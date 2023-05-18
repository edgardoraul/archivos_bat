Attribute VB_Name = "Cursos"
Option Explicit

' Color que será usado para seleccionar y contar prendas
Const verde = 40
Const gris = 16

' Columna a partir de la cual comienzan los productos
Const DesdeAqui = 5

' Cantidad de columnas que corresponden a cada producto
Public productos As Byte

' Lleva la cuenta de la cantidad de cada talle
Public contador As Integer

' Ultimas Filas y Columnas
Public ultimaFila As Integer
Public ultimaColumna As Byte

' Condicional para los separados
Public separados As Boolean

' Talles
Public talleArriba As Variant
Public talleCalzado As Variant
Public talleAbajo As Variant
Public talleCabeza As Variant
Public talleCosas As Variant
Public talleCinto As Variant
Public talleSinTalle As Variant
Public cantProd As Byte
Public cantCols As Byte
Public productoActual As String

' Textos
Public TITULARES As Variant
Public hojasTotalizadoras As Variant
Public hojita As Variant

Function negritaCentrado(rango)
' Da formato de centrado, en negrita y rodeado con líneas
    With rango
        .HorizontalAlignment = xlCenter
        .Font.Bold = True
        .Borders.LineStyle = xlContinuous
    End With
End Function


Sub constructorHoja()
' CONSTRUYE LAS HOJAS NECESARIAS PARA TRABAJAR
    'Elimina cualquier hipervícunlo que haya
    Worksheets(1).Hyperlinks.Delete
     
    ' Formatea el titular de la planilla
    With Worksheets(1).Range("A1:D1")
        .Merge
        .Font.Size = 16
    End With
    Call negritaCentrado(Worksheets(1).Range("A1:D1"))
    
    
    Dim z As Integer 'Columna inicial: Es la del primer producto
    Dim x As Byte 'Columnas totales en TOTALES Y SEPARADOS. El doble que "z"
    Dim titular As String 'El nombre de c/producto en PLANILLA
    Dim hojasTotales As Byte 'Cantidad de hojas del documento. Irá incrementando a medida que se vayan creando
    
    ' Textos a utilizar más adelante
    TITULARES = Array("TALLES", "TOTALES", "SEPARADOS", "FALTANTES")
    
    ultimaColumna = Worksheets(1).Cells(2, Columns.Count).End(xlToLeft).Column
    ultimaFila = Worksheets(1).Cells(Rows.Count, 2).End(xlUp).Row
    
    ' Se construye sólo una vez al comenzar el archivo
    If ThisWorkbook.Worksheets.Count < 2 Then
        Call constructorTotalesSeparados
    End If
    
    ' Cuenta cantidad de hojas en el archivo
    hojasTotales = ThisWorkbook.Worksheets.Count
    
    z = DesdeAqui
    cantProd = ultimaColumna - DesdeAqui + 1
    cantCols = cantProd * 2
    x = 2

    
    For z = DesdeAqui To ultimaColumna
        Worksheets(1).Activate
        Worksheets(1).Cells(2, z).Activate
        titular = Worksheets(1).Cells(2, z).Value
        productoActual = Worksheets(1).Cells(2, z).Address
        
        ' Agrega siempre en desde la última hoja
        If z > hojasTotales Then
            Worksheets.Add(after:=Worksheets(hojasTotales)).Name = titular
            
            Call constructorProducto(hojasTotales + 1, titular, x)
            x = x + 2
            hojasTotales = ThisWorkbook.Worksheets.Count
            
        End If
    Next z
End Sub

Function constructorTotalesSeparados()

' CONSTUYE LAS HOJAS DE TOTALES, SEPARADOS Y FALTANTES
hojasTotalizadoras = Array("TOTALES", "SEPARADOS", "FALTANTES")

For Each hojita In hojasTotalizadoras
    ' Coloca el nombre correcto
    Worksheets.Add(after:=Worksheets(ThisWorkbook.Worksheets.Count)).Name = hojita
    
    ' Copia el título de la hoja principal: PLANILLA
    Sheets(hojita).Range("A1").Value = "=" & Worksheets(1).Name & "!$A$1"
    
Next hojita

End Function

Sub constructorProducto(indice, titulo, prodActual)

' CONSTRUYE LA PLANILLA INTERNA DE CADA PRODUCTO
Dim leyendaTallesArriba As String
Dim leyendaTallesCalzado As String
Dim leyendaTallesAbajo As String
Dim leyendaTallesCabeza As String
Dim leyendaTallesCosas As String
Dim leyendaTallesCinto As String
Dim leyendaSinTalles As String
Dim eleccion As Integer
Dim grupoTalles As Variant
Dim arreglo As Variant
Dim fila As Byte
Dim leyenda As String
Dim direccion As String

leyendaTallesArriba = "Remeras, tricotas, garibaldinas, mamelucos, camperas, etc..."
leyendaTallesCalzado = "Zapatos, borcegos, botines, etc..."
leyendaTallesAbajo = "Camisas, pantalones, bombachas, etc..."
leyendaTallesCabeza = "Quepis, casquetes, gorras plato, etc..."
leyendaTallesCosas = "Diestro, surdo, tipo de arma, grupo sanguíneo, etc..."
leyendaTallesCinto = "Cinturones"
leyendaSinTalles = "Sin Talles"

leyenda = "Escribí un número del 1 al 7." & vbNewLine & vbNewLine & "1: " & leyendaTallesArriba & vbNewLine & "2: " & leyendaTallesCalzado & vbNewLine & "3: " & leyendaTallesAbajo & vbNewLine & "4: " & leyendaTallesCabeza & vbNewLine & "5: " & leyendaTallesCosas & vbNewLine & "6: " & leyendaTallesCinto & vbNewLine & "7: " & leyendaSinTalles & "."

talleArriba = Array("3XS", "XXS", "XS", "S", "M", "L", "XL", "XXL", "3XL", "4XL", "5XL", "6XL")
talleCalzado = Array(30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50, 51, 52, 53, 54, 55)
talleAbajo = Array(32, 34, 36, 38, 40, 42, 44, 46, 48, 50, 52, 54, 56, 58, 60, 62, 64, 66)
talleCabeza = Array(50, 51, 52, 53, 54, 55, 56, 57, 58, 59, 60, 61, 62, 63, 64, 65, 66, 67, 68, 69, 70)
talleCosas = Array("CIN", "MOL", "BER", "BRO", "TAU", "IZQ", "DER", "ABN", "ABP", "AN", "AP", "BN", "BP", "ON", "OP")
talleCinto = Array(80, 81, 82, 83, 84, 85, 86, 87, 88, 89, 90, 91, 92, 93, 94, 95, 96, 97, 98, 99, 100, 101, 102, 103, 104, 105, 106, 107, 108, 109, 110, 111, 112, 113, 114, 115, 116, 117, 118, 119, 120, 121, 122, 123, 124, 125, 126, 127, 128, 129, 130, 131, 132, 133, 134, 135, 136, 137, 138, 139, 140, 141, 142, 143, 144, 145, 146, 147, 148, 149, 150)
talleSinTalle = Array(1)

    With Worksheets(indice)
        ' Copia el título de la primer hoja "PLANILLA"
        .Range("A1").Value = Worksheets(1).Range("A1").Value
        
        ' Los respectivos títulos y subtítulos del producto
        .Range("A2").Value = titulo
        .Range("A3").Value = "TALLES"
        .Range("B3").Value = "TOTALES"
        .Range("C3").Value = "SEPARADOS"
        .Range("D3").Value = "FALTANTES"
    End With
    
    ' Hipervínculos a las pestañas respectivas
    With Worksheets(indice).Hyperlinks
        '-> Planilla principal
        .Add Anchor:=Worksheets(indice).Range("A1"), Address:="", SubAddress:="'" & Worksheets(1).Name & "'!A1", ScreenTip:="Ir a " & Sheets(1).Name & "."
        
        '-> A la columna correspondiente de la planilla principal
        .Add Anchor:=Worksheets(indice).Range("A2"), Address:="", SubAddress:="'" & Worksheets(1).Name & "'!" & productoActual & "", ScreenTip:="Ir a " & titulo & " en " & Sheets(1).Name & "."
        
        '-> A los TOTALES
        .Add Anchor:=Worksheets(indice).Range("B3"), Address:="", SubAddress:="'" & Worksheets(2).Name & "'!A1", ScreenTip:="Ir a " & Worksheets(2).Name & ".", TextToDisplay:=Range("B3").Value
        
        '-> A los SEPARADOS
        .Add Anchor:=Worksheets(indice).Range("C3"), Address:="", SubAddress:="'" & Worksheets(3).Name & "'!A1", ScreenTip:="Ir a " & Worksheets(3).Name & ".", TextToDisplay:=Range("C3").Value
        
        '-> A los FALTANTES
        .Add Anchor:=Worksheets(indice).Range("D3"), Address:="", SubAddress:="'" & Worksheets(4).Name & "'!A1", ScreenTip:="Ir a " & Worksheets(4).Name & ".", TextToDisplay:=Range("D3").Value
    End With
    
    ' Hypervínculos desde la planilla principal al respectivo producto
    With Worksheets(1).Range(productoActual).Hyperlinks
        .Add Anchor:=Worksheets(1).Range(productoActual), Address:="", _
            SubAddress:="'" & Worksheets(indice).Name & "'!A2", _
            ScreenTip:="Ir a " & Worksheets(indice).Name & "."
    End With
    
    eleccion = VBA.InputBox(Prompt:=leyenda, Title:="Grupo de Talles: " & titulo)
    
    MsgBox "Elegiste la opción: " & eleccion
volver:
    ' Elección del grupo de talles
    Select Case eleccion
        Case 1
            grupoTalles = talleArriba
        Case 2
            grupoTalles = talleCalzado
        Case 3
            grupoTalles = talleAbajo
        Case 4
            grupoTalles = talleCabeza
        Case 5
            grupoTalles = talleCosas
        Case 6
            grupoTalles = talleCinto
        Case 7
            grupoTalles = talleSinTalle
        Case Else
            MsgBox ("Tenés que elegir alguna opción")
            GoTo volver
    End Select

    fila = 4
    If WorksheetFunction.CountA(grupoTalles) > 0 Then
        For Each arreglo In grupoTalles
             
            ' Datos en la planilla de TOTALES y SEPARADOS
            Call datosTotales(fila, arreglo, titulo, WorksheetFunction.CountA(grupoTalles), prodActual)
            
            ' Datos de cada opción de talle
            Worksheets(indice).Cells(fila, 1).Value = arreglo
            
            ' Datos de TOTALES
            direccion = Worksheets(2).Cells(fila, prodActual).Address
            Worksheets(indice).Cells(fila, 2).Formula = "=" & Worksheets(2).Name & "!" & direccion
            
            ' Datos de SEPARADOS
            direccion = Worksheets(3).Cells(fila, prodActual).Address
            Worksheets(indice).Cells(fila, 3).Formula = "=" & Worksheets(3).Name & "!" & direccion
            
            ' La resta de los talles que faltan
            Worksheets(indice).Cells(fila, 4).Value = "=B" & fila & "-C" & fila & ""
            
            ' Incrementa contador
            fila = fila + 1
            
            ' Limpia la variable
            direccion = ""
        Next arreglo
    End If
    
    ' Las sumatorias
    With Worksheets(indice)
        .Cells(fila, 1).Value = "TOTALES"
        .Cells(fila, 2).Value = "=SUM(B4:B" & fila - 1 & ")"
        .Cells(fila, 3).Value = "=SUM(C4:C" & fila - 1 & ")"
        .Cells(fila, 4).Value = "=SUM(D4:D" & fila - 1 & ")"
    End With
    
    ' Combinar celdas
    Worksheets(indice).Range("A1:D1").Merge
    Worksheets(indice).Range("A2:D2").Merge
    
    ' Negrita y fuente más grande para las autosumas
    With Worksheets(indice).Range(Cells(fila, 1), Cells(fila, 4))
        .Font.Bold = True
        .Font.Size = 16
    End With
    
    ' Negrita y centrado de títulos y subtítulos
    With Worksheets(indice).Range("A1:D3")
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    
    ' Redimensionar la tabla
    With Worksheets(indice).Range("A2").CurrentRegion
        .EntireColumn.AutoFit
        .Borders.LineStyle = xlContinuous
    End With
    
    
End Sub

Sub datosTotales(filita, datos, hojaOrigen, cantidadTalles, prodActual)
' Completa datos para llenar en la planilla TOTALES

Dim hojita As Variant

For Each hojita In hojasTotalizadoras
    ' Títulos del producto
    Sheets(hojita).Cells(2, prodActual - 1).Value = hojaOrigen
    
    ' Combina el título del producto con la celda adyacente
    'Worksheets(hojita).Range(Cells(2, prodActual - 1), Cells(2, prodActual)).Merge
    'Call negritaCentrado(Sheets(hojita).Range(Cells(2, prodActual - 1), Cells(2, prodActual)))
    
    ' Subtítulo del talle
    Sheets(hojita).Cells(3, prodActual - 1).Value = "TALLES"
    Call negritaCentrado(Sheets(hojita).Cells(3, prodActual - 1))
    
    ' Subtítulos de totales, separados y faltantes
    Sheets(hojita).Cells(3, prodActual).Value = hojita
    Call negritaCentrado(Sheets(hojita).Cells(3, prodActual))
    
    ' Datos de los talles y formato
    Sheets(hojita).Cells(filita, prodActual - 1).Value = datos
    With Sheets(hojita).Cells(filita, prodActual - 1)
        .Value = datos
        .Borders.LineStyle = xlContinuous
    End With
    
    ' Formato a los datos
    With Sheets(hojita).Cells(filita, prodActual)
        .Borders.LineStyle = xlContinuous
    End With
    
    ' Fórmulas en la hoja FALTANTES
    With Worksheets(4).Cells(filita, prodActual)
        .Value = "=" & Worksheets(2).Name & "!R" & filita & "C" & prodActual & "-" & Worksheets(3).Name & "!R" & filita & "C" & prodActual & ""
    End With
    
    ' Autosuma, leyenda "total" y formato
    If filita - 3 = cantidadTalles Then
        Sheets(hojita).Cells(filita + 1, prodActual).FormulaR1C1 = "=SUM(R4C" & prodActual & ":R" & filita & "C" & prodActual & ")"
        With Sheets(hojita).Cells(filita + 1, prodActual)
            .Borders.LineStyle = xlContinuous
            .Font.Bold = True
        End With
        
        Sheets(hojita).Cells(filita + 1, prodActual - 1).Value = "Total"
        With Sheets(hojita).Cells(filita + 1, prodActual - 1)
            .Borders.LineStyle = xlContinuous
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
        End With
    End If
    
Next hojita

End Sub

Sub formato()
' DA FORMATO Y CREA FILAS HACIA ABAJO A MEDIDA QUE SE
' CARGAN LOS PRODUCTOS.
Dim numero As Integer
Dim miRango As Range

ultimaColumna = Worksheets(1).Cells(2, Columns.Count).End(xlToLeft).Column
ultimaFila = Worksheets(1).Cells(Rows.Count, 2).End(xlUp).Row

numero = ultimaFila - 2
Set miRango = Range(Cells(3, 1), Cells(ultimaFila, ultimaColumna))

' SALE DE LA MACRO PARA EVITAR SOBREESCRIBIR TITULOS
If numero = 1 Then
    Exit Sub
End If

' FORMATEA PRIMERA COLUMNA A MEDIDA QUE SE COMPLETA
With Cells(ultimaFila, 1)
    .Value = numero
    .Font.ColorIndex = gris
End With

' FORMATEA FILAS CON TITULOS
With Range(Cells(2, 1), Cells(2, ultimaColumna))
    .HorizontalAlignment = xlCenter
    .Borders.LineStyle = xlContinuous
    .Font.Bold = True
End With

' Convierte en mayúsuculas los títulos
Dim celda As Variant
For Each celda In Range(Cells(2, 1), Cells(2, ultimaColumna))
    celda.Value = UCase(celda.Value)
Next celda

' FORMATEA FILAS CON DATOS RELEVANTES PRESELECCIONADOS
miRango.FormatConditions.Add Type:=xlCellValue, Operator:=xlNotEqual, Formula1:=xlEmptyCellReferences
With miRango.Borders
    .LineStyle = xlContinuous
End With

' Se autoposiciona para facilitar el completado
Cells(ultimaFila, 3).Activate

End Sub

Sub marcar()
Attribute marcar.VB_Description = "Sirve para marcar lo que está separado, inventariado, contado y listo para embalar y ser entregado al cadete."
Attribute marcar.VB_ProcData.VB_Invoke_Func = "M\n14"
' Ctrol +  May + M
' MARCA LAS CELDAS CON COLOR
' SOLO LAS CELDAS MARCADAS PUEDEN SUMARSE
Dim ultimaColumna As Byte
ultimaColumna = Worksheets(1).Cells(2, Columns.Count).End(xlToLeft).Column

If ActiveCell.Column >= DesdeAqui And ActiveCell.Column <= ultimaColumna And ActiveCell.Row > 2 Then
    
    If ActiveCell.Interior.ColorIndex = verde Or Selection.Interior.ColorIndex = verde Then
        ActiveCell.Interior.Color = xlNone
        Selection.Interior.ColorIndex = xlNone
    Else
        ActiveCell.Interior.ColorIndex = verde
        Selection.Interior.ColorIndex = verde
    End If
End If

End Sub

Sub proteger()
' PROTEGE EL LIBRO Y SUS HOJAS


On Error GoTo Fin
Dim hojita As Worksheet

Application.ScreenUpdating = False
ActiveWorkbook.Unprotect Password:="Rerda"
For Each hojita In ActiveWorkbook.Worksheets
    If hojita.Visible = True Then
        With hojita
            'If hojita.Index = 1 Then
             '   .Protect Password:="Rerda", AllowFormattingCells:=True, AllowFormattingRows:=True
            'ElseIf hojita.Index > 3 Then
                ' Se desprotegen las hojas 2 y 3 ya que sirven para guardar datos dinámicos
                .Unprotect Password:="Rerda"
            'End If
        End With
    End If
Next

' Se posiciona en el principio
Worksheets(1).Activate
ThisWorkbook.Save
Fin:
End Sub

Sub buscar_producto()
Attribute buscar_producto.VB_Description = "Cuenta los talles de productos en cuestión."
Attribute buscar_producto.VB_ProcData.VB_Invoke_Func = "J\n14"
' Hoja: TOTALES
' Recorre la fila de los títulos para coincidir con el texto "TALLES".

' Primer hoja
Worksheets(1).Activate
Worksheets(1).Range("A1").Activate

' Controla si están creadas las planillas
If ThisWorkbook.Worksheets.Count = 1 Then
    MsgBox ("Primero tenés que generar todas planillas")
    Exit Sub
End If

' Variables para contar
Dim e As Byte
Dim f As Byte
Dim cuentaProducto As Byte

' Asigando este valor falso
separados = False

' Total de columnas incluídas las vacías
Dim columnas As Byte

' Cantidad de productos en la hoja PLANILLA
productos = Worksheets(1).Cells(2, Columns.Count).End(xlToLeft).Column - DesdeAqui + 1
columnas = productos * 2
    
' Recorre las columnas de hoja TOTALES evadiendo las vacías
For e = 1 To columnas
    
    If Worksheets(2).Cells(3, e).Value = "TALLES" Then
        cuentaProducto = cuentaProducto + 1
        Call RecorreProductos(e, productos, cuentaProducto, separados)
    Else
        GoTo proxima
    End If

proxima:
Next e

' Posiciona en la hoja TOTALES
Worksheets(2).Activate

' Formato lindo
Call formatear(2, columnas)

' Muestra el aviso de proceso terminado
MsgBox "TOTALES COMPLETADOS! :-)", vbOKOnly

End Sub
Sub buscar_separados()
Attribute buscar_separados.VB_Description = "Ejecuta la cuenta de los productos que están marcados y separados."
Attribute buscar_separados.VB_ProcData.VB_Invoke_Func = "I\n14"
' Hoja: SEPARADOS
' Recorre la fila de los títulos para coincidir con el texto "TALLES".

' Primer hoja
Worksheets(1).Activate
Worksheets(1).Range("A1").Activate

' Controla si están creadas las planillas
If ThisWorkbook.Worksheets.Count = 1 Then
    MsgBox ("Primero tenés que generar todas planillas")
    Exit Sub
End If


' Variables para contar
Dim e As Byte
Dim f As Byte
Dim cuentaProducto As Byte

' Asigando este valor verdadero
separados = True

' Total de columnas incluídas las vacías
Dim columnas As Byte

' Cantidad de productos en la hoja PLANILLA
productos = Worksheets(1).Cells(2, Columns.Count).End(xlToLeft).Column - DesdeAqui + 1
columnas = productos * 2
    
' Recorre las columnas de hoja TOTALES evadiendo las vacías
For e = 1 To columnas
    
    If Worksheets(3).Cells(3, e).Value = "TALLES" Then
        cuentaProducto = cuentaProducto + 1
        Call RecorreProductos(e, productos, cuentaProducto, separados)
    Else
        GoTo proxima
    End If

proxima:
Next e

' Posiciona en la hoja TOTALES
Worksheets(3).Activate

' Formato lindo
Call formatear(3, columnas)
Call formatear(4, columnas)

' Muestra el aviso de proceso terminado
MsgBox "Productos marcados y separados -> contados todo OK :-)", vbOKOnly
End Sub


Sub RecorreProductos(columna, productos, cuentaProducto, separados)
' Hoja TOTALES
' Recorre cada fila correspondiente al talle para luego
' llamar a la función que realiza la búsqueda.

Dim g As Byte
Dim talleBuscar As Variant
Dim talleEncontrar As Variant
Dim tipoCuenta As Byte

If separados = True Then
    tipoCuenta = 3
Else
    tipoCuenta = 2
End If

Worksheets(tipoCuenta).Activate

    g = 4 ' -> desde la 4° fila empieza.
    
    Do Until Worksheets(tipoCuenta).Cells(g, columna).Value = ""
                
        ' Talle a buscar
        talleBuscar = Worksheets(tipoCuenta).Cells(g, columna).Value
        
        ' Llama la macro que cuenta los talles
        ' Necesita el valor buscado y la columna a partir de la cual buscar
        Call CuentaTalles(talleBuscar, cuentaProducto + DesdeAqui - 1, separados)
        
        
        ' Se inserta el resultado de la suma de talles
        If contador > 0 Then
            Worksheets(tipoCuenta).Cells(g, columna + 1).Value = contador
        End If
        
        ' Baja una fila y repite
        g = g + 1
        
        ' Resetea el contador
        contador = 0
    Loop
End Sub

Function CuentaTalles(talleBuscar, cuentaProducto, condicion)
' CUENTA LOS TALLES MARCADOS DEL PRODUCTO.
' TAMBIÉN COMPLETA POR DIFERENCIA UNA HOJA CON LOS FALTANTES.

' Todas las filas con datos de la Planilla
Dim ultimaFila As Integer

' Acumulador simple
Dim h As Integer
Dim talleEncontrar As Variant

ultimaFila = Worksheets(1).Cells(Rows.Count, 2).End(xlUp).Row
    
' Cuenta la cantidad de talles. Necesita ubicar:
' la fila y la columna para empezar a contar.

For h = 3 To ultimaFila
    ' Controla si hay talle y si está marcado con color

    talleEncontrar = Worksheets(1).Cells(h, cuentaProducto).Value
    
    If talleEncontrar = talleBuscar And condicion = False Then
        ' Acumula lo contado
        contador = contador + 1
    
    ElseIf talleEncontrar = talleBuscar And condicion = True And Worksheets(1).Cells(h, cuentaProducto).Interior.ColorIndex = verde Then
        
        ' Acumula lo contado
        contador = contador + 1
        
    End If

Next h

End Function

Function formatear(hoja, columnas)
' Formatea la hoja con algo lindo
Worksheets(hoja).Activate

Dim i As Byte

' Fusiona las 4 celdas del título de la hoja y le agrega hiperlink a la principal
With Worksheets(hoja).Range("A1:D1")
    .Merge
    .Hyperlinks.Add Anchor:=Worksheets(hoja).Range("A1"), Address:="", _
        SubAddress:="'" & Worksheets(1).Name & "'!A1", _
        ScreenTip:="Ir a " & Sheets(1).Name & "."
    .Font.Bold = True
    .Font.Size = 14
End With

With Worksheets(hoja).Range("A3").CurrentRegion
 ' Negrita y fuente más grande para las autosumas
    .EntireColumn.AutoFit
End With

Call negritaCentrado(Worksheets(hoja).Range(Cells(3, 1), Cells(3, columnas)))

i = 2
Do While i <= columnas
    Worksheets(hoja).Range(Cells(2, i), Cells(2, i - 1)).Merge
    With Worksheets(hoja).Cells(2, i - 1)
        .HorizontalAlignment = xlCenter
        .Font.Bold = True
    End With
    
    With Worksheets(hoja).Range(Cells(2, i - 1), Cells(2, i))
        .Borders.LineStyle = xlContinuous
    End With
    i = i + 2
Loop



End Function
