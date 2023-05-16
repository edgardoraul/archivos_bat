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


Sub constructorHoja()
' CONSTRUYE LAS HOJAS NECESARIAS PARA TRABAJAR
    Dim z As Integer
    Dim x As Byte
    Dim titular As String
    Dim hojasTotales As Byte
    
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

    Debug.Print "Cant Prod: " & cantProd, "Cant Cols: " & cantCols, "Col actual: " & x
    
    For z = DesdeAqui To ultimaColumna
        Worksheets(1).Activate
        Worksheets(1).Cells(2, z).Activate
        titular = Worksheets(1).Cells(2, z).Value
        
        ' Agrega siempre en desde la última hoja
        If z > hojasTotales + 1 Then
            Worksheets.Add(After:=Worksheets(hojasTotales)).Name = titular
            
            Call constructorProducto(hojasTotales + 1, titular, x)
            x = x + 2
            hojasTotales = ThisWorkbook.Worksheets.Count
            Debug.Print "Cant Prod: " & cantProd, "Cant Cols: " & cantCols, "Col actual: " & x
    
        End If
    Next z
End Sub

Function constructorTotalesSeparados()
' CONSTUYE LAS HOJAS DE TOTALES Y SEPARADOS
    
    ' Crear las hoja 1 Totales y su título
    Worksheets.Add(After:=Worksheets(ThisWorkbook.Worksheets.Count)).Name = "TOTALES"
    Worksheets(2).Range("A1").Value = "=" & Worksheets(1).Name & "!$A$1"
    
    ' Crea la hoja 3 y su título
    Worksheets.Add(After:=Worksheets(ThisWorkbook.Worksheets.Count)).Name = "SEPARADOS"
    Worksheets(3).Range("A1").Value = "=" & Worksheets(1).Name & "!$A$1"

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
    
    
    eleccion = VBA.InputBox(Prompt:=leyenda, Title:="Grupo de Talles")
    
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
            
            ' Datos de TOTALES CORREGIR AQUÍ LA FORMULA
            direccion = Worksheets(2).Cells(fila, prodActual).Address
            Debug.Print direccion
            Worksheets(indice).Cells(fila, 2).Formula = "=" & Worksheets(2).Name & "!" & direccion
            
            ' Datos de SEPARADOS CORREGIR AQUÍ LA FORMULA
            direccion = Worksheets(3).Cells(fila, prodActual).Address
            Debug.Print direccion
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
    Worksheets(indice).Range("A2:D2").Merge
    
    ' Negrita y fuente más grande para las autosumas
    With Worksheets(indice).Range(Cells(fila, 1), Cells(fila, 4))
        .Font.Bold = True
        .Font.Size = 16
    End With
    
    ' Negrita y centrado de subtítulos
    With Worksheets(indice).Range("A2:D3")
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

    ' Títulos del producto
    Worksheets(2).Cells(2, prodActual - 1).Value = hojaOrigen
    Worksheets(3).Cells(2, prodActual - 1).Value = hojaOrigen

    ' Subtítulo del talle
    Worksheets(2).Cells(3, prodActual - 1).Value = "TALLES"
    Worksheets(3).Cells(3, prodActual - 1).Value = "TALLES"

    ' Subtítulos de totales
    Worksheets(2).Cells(3, prodActual).Value = "TOTALES"
    Worksheets(3).Cells(3, prodActual).Value = "SEPARADOS"

    
    ' Datos de los talles
    Worksheets(2).Cells(filita, prodActual - 1).Value = datos
    Worksheets(3).Cells(filita, prodActual - 1).Value = datos
    
    ' Autosuma
    If filita - 3 = cantidadTalles Then
        Worksheets(2).Cells(filita + 1, prodActual).FormulaR1C1 = "=SUM(R4C" & prodActual & ":R" & filita & "C" & prodActual & ")"
        Worksheets(3).Cells(filita + 1, prodActual).FormulaR1C1 = "=SUM(R4C" & prodActual & ":R" & filita & "C" & prodActual & ")"
    End If

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
                Debug.Print hojita.Index
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

Dim e As Byte
Dim f As Byte
Dim cuentaProducto As Byte

' Asigando este valor falso
separados = False

' Total de columnas incluídas las vacías
Dim columnas As Byte

' Cantida de productos en la hoja PLANILLA
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

Dim e As Byte
Dim f As Byte
Dim cuentaProducto As Byte

' Asigando este valor verdadero
separados = True

' Total de columnas incluídas las vacías
Dim columnas As Byte

' Cantida de productos en la hoja PLANILLA
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
        Else
            Worksheets(tipoCuenta).Cells(g, columna + 1).Value = ""
        End If
        
        ' Baja una fila y repite
        g = g + 1
        
        ' Resetea el contador
        contador = 0
    Loop
End Sub

Function CuentaTalles(talleBuscar, cuentaProducto, condicion)
' CUENTA LOS TALLES MARCADOS DEL PRODUCTO

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

Dim i As Byte


With Worksheets(hoja).Range("A3").CurrentRegion
 ' Negrita y fuente más grande para las autosumas
    .EntireColumn.AutoFit
End With

With Worksheets(hoja).Range(Cells(3, 1), Cells(3, columnas))
    .HorizontalAlignment = xlCenter
    .Borders.LineStyle = xlContinuous
    .Font.Bold = True
End With

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
