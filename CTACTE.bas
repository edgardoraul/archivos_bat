Attribute VB_Name = "CTACTE"
Option Explicit
Sub Rotulador()
    Sheets("Planilla").Select

    ' Declaración de Variables y su tipo de dato
    Dim apellidoNombre As String
    Dim dniCuit As String
    Dim direccion As String
    Dim codigoPostal As Variant
    Dim ciudad As String
    Dim provincia As String
    Dim telefono As String
    Dim codigoNis As String
    Dim hoja As Worksheet
    Dim fecha As String

    fecha = Format(Date, "yyyy-mm-dd")
    
     ' Controlar quién es el vendedor
    
    'If Range("T2").Value = "" Then
    'Cells(
    'If Cells(2, Columns.Count).End(xlToLeft).Column).Value = "" Then
    If Cells(2, Cells(1, Columns.Count).End(xlToLeft).Column).Value = "" Then
        MsgBox ("¿Y qué viajante, vendedor o sucursal sos vos?")
        Range("T2").Select
        Exit Sub
    End If
    
    ' Dando formato a la página para imprimir
    With ActiveSheet.PageSetup
        .Orientation = xlLandscape
        .TopMargin = Application.CentimetersToPoints(1.9)
        .RightMargin = Application.CentimetersToPoints(0.6)
        .BottomMargin = Application.CentimetersToPoints(1.9)
        .LeftMargin = Application.CentimetersToPoints(0.6)
        .HeaderMargin = Application.CentimetersToPoints(0.8)
        .FooterMargin = Application.CentimetersToPoints(0.8)
        .CenterHorizontally = True
        .PaperSize = xlPaperA4
    End With

    ' Controla si está parado en una celda equivocada
    If ActiveCell.Column <> Cells(1, Columns.Count).End(xlToLeft).Column - 1 Then
        MsgBox ("Debes elegir alguna compra que tenga algún tipo de flete.")
        Cells.Range("S2").Select
        Exit Sub
    
    ElseIf ActiveCell.Value = "" Then
        Cells(Cells(1, Columns.Count).End(xlToLeft).Column - 1).Select
        Exit Sub
    End If
    
    
    ' Variables para operar
    apellidoNombre = ActiveCell.Offset(0, -20).Value
    dniCuit = ActiveCell.Offset(0, -10).Value
    direccion = ActiveCell.Offset(0, -5).Value
    codigoPostal = ActiveCell.Offset(0, -3).Value
    ciudad = ActiveCell.Offset(0, -2).Value
    provincia = ActiveCell.Offset(0, -1).Value
    telefono = ActiveCell.Offset(0, -4).Value
    
    ' Controla si está completo el DNI/CUIT y el CP
    If dniCuit = "" Then
        MsgBox "Te faltó completar el DNI/CUIT."
        
        ' Celda del DNI/CUIT
        ActiveCell.Offset(0, -10).Activate
        Exit Sub
    
    ElseIf apellidoNombre = "" Then
        MsgBox "Te faltó completar el Apellido y Nombre."
        
        ' Celda del Apellido y Nombre
        ActiveCell.Offset(0, -20).Activate
        Exit Sub
        
    ElseIf codigoPostal = "" Then
        MsgBox "Te faltó completar el Código Postal."
        
        ' Celda del CP
        ActiveCell.Offset(0, -3).Activate
        Exit Sub
    
    ElseIf ciudad = "" Then
        MsgBox "Te faltó completar la Ciudad."
        
        ' Celda del Ciudad
        ActiveCell.Offset(0, -2).Activate
        Exit Sub
    
    ElseIf provincia = "" Then
        MsgBox "Te faltó completar la Provincia."
        
        ' Celda del Provincia
        ActiveCell.Offset(0, -1).Activate
        Exit Sub
    
    ElseIf telefono = "" Then
        MsgBox "Te faltó completar el Teléfono."
        
        ' Celda del Teléfono
        ActiveCell.Offset(0, -4).Activate
        Exit Sub
    End If


    ' SI ES A DOMICILIO -----------
    If ActiveCell.Value = Sheets("Opciones").Range("A5").Value Then
    
        If direccion = "" Then
            MsgBox "Te faltó completar la Dirección."
        
            ' Celda del Dirección
            ActiveCell.Offset(0, -5).Activate
            Exit Sub
        End If
        
        ' Colocar los datos
        With Sheets("A Domicilio")
            .Range("C16").Value = UCase(apellidoNombre)
            .Range("P15").Value = dniCuit
            .Range("C18").Value = direccion
            .Range("C23").Value = UCase(provincia)
            .Range("E21").Value = codigoPostal
            .Range("G21").Value = UCase(ciudad)
            .Range("P23").Value = telefono
        End With
        
        ' Generar Proforma
        Call proforma(apellidoNombre, direccion, provincia, codigoPostal, ciudad, telefono, fecha)
        
        ' Generar rotulo
        Call Rotulo_Correo_Argentino(apellidoNombre, Sheets("A Domicilio"), fecha)
        

    ' SI ES A RETIRAR EN SUCURSAL ----------------
    ElseIf ActiveCell.Value = Sheets("Opciones").Range("A3").Value Then
        
        ' Colocar los datos
        With Sheets("A Sucursal")
            .Range("C16").Value = UCase(apellidoNombre)
            .Range("R16").Value = dniCuit
        End With
        
        ' Validando si existe o no el dato.
        On Error Resume Next
        codigoNis = Sheets("Sucursales").Range("$F$4:$F$5000").Find(what:=codigoPostal, LookIn:=xlValues, LookAt:=xlPart).Offset(0, -5)
    
        If codigoNis = "" Then
            ' Todo salió mal
            Sheets("Planilla").Activate
            MsgBox ("El código postal " & codigoPostal & " no corresponde con ninguna sucursal del Correo. Intentá con otro. ")
            Sheets("Sucursales").Select
            MsgBox ("Buscá aquí un código postal de sucursal disponible")
            Exit Sub
        End If
        
        ' Completando el resto de datos.
        With Sheets("A Sucursal")
            .Range("S18").Value = codigoNis
            .Range("R22").Value = telefono
        End With
        
        ' Generar Proforma
        Call proforma(apellidoNombre, "Retiro en Sucursal del Correo Argentino Cód. NIS " & codigoNis, provincia, codigoPostal, ciudad, telefono, fecha)
        
        ' Generar rotulo
        Call Rotulo_Correo_Argentino(apellidoNombre, Sheets("A Sucursal"), fecha)
        
    
    ' SI ES PAGO DE FLETE EN DESTINO
    ElseIf ActiveCell.Value = Sheets("Opciones").Range("A4").Value Then

        ' Colocar los datos
        With Sheets("Pago en Destino")
            .Range("C16").Value = UCase(apellidoNombre)
            .Range("R16").Value = dniCuit
        End With
        
        ' Validando si existe o no el dato.
        On Error Resume Next
        codigoNis = Sheets("Sucursales").Range("$F$4:$F$5000").Find(what:=codigoPostal, LookIn:=xlValues, LookAt:=xlPart).Offset(0, -5)
    
        If codigoNis = "" Then
            ' Todo salió mal
            Sheets("Planilla").Activate
            MsgBox ("El código postal " & codigoPostal & " no corresponde con ninguna sucursal del Correo. Intentá con otro. ")
            Sheets("Sucursales").Activate
            MsgBox ("Buscá aquí un código postal de sucursal disponible")
            Exit Sub
        End If
        
        ' Completando el resto de datos.
        With Sheets("Pago en Destino")
            .Range("S18").Value = codigoNis
            .Range("R22").Value = telefono
        End With
        
        Sheets("Planilla").Activate
        
        ' Generar Proforma
        Call proforma(apellidoNombre, "Retiro en Sucursal del Correo Argentino Cód. NIS " & codigoNis, provincia, codigoPostal, ciudad, telefono, fecha)
        
        ' Generar rotulo
        Call Rotulo_Correo_Argentino(apellidoNombre, Sheets("Pago en Destino"), fecha)
    
    
    '   SI ES RETIRO EN LOCAL -----------
    ElseIf ActiveCell.Value = Sheets("Opciones").Range("A2").Value Then
        
        ' Colocar los datos
        With Sheets("Retiro en Local")
            .Range("C16").Value = UCase(apellidoNombre)
            .Range("R16").Value = dniCuit
            .Range("R22").Value = telefono
        End With
        
        ' Generar rotulo
        Call Rotulo_Correo_Argentino(apellidoNombre, Sheets("Retiro en Local"), fecha)
        Sheets("Planilla").Activate
    
    End If


End Sub


Function Rotulo_Correo_Argentino(apellidoNombre, rotulo, fecha)
'
' Rotulos Macro
' Guarda los rótulos en pdf
'
' Acceso directo: CTRL+MAY+Ñ
'   Guarda antes de crear el archivo
    ThisWorkbook.Save
    
    ' Declaración de Variables y su tipo de datos
    Dim nombre As String
    Dim ruta As String
    Dim nombreCarpeta As String

    ' Variables necesarias
    nombre = fecha & ". " & apellidoNombre
    ruta = ThisWorkbook.Path
    nombreCarpeta = "\Rotulos\"
    
    ' Comprobando si existe o no la carpeta
    If Dir(ruta, vbDirectory) <> "" Then
        If Dir(ruta & nombreCarpeta, vbDirectory) = "" Then
            MkDir ruta & nombreCarpeta
        End If
    End If
    
    ' Generando el archivo pdf
    rotulo.ExportAsFixedFormat Type:=xlTypePDF, fileName:= _
        ruta & nombreCarpeta & UCase(nombre) & ".pdf", _
        OpenAfterPublish:=True
    Sheets("Planilla").Activate
End Function

Function Validar_CP(cp)
' Valida si el código postal es correcto, existe o no.

    codigoNis = Sheets("Sucursales").Range("F1576:F5000").Find(what:=cp, LookIn:=xlValues, SearchOrder:=xlByRows, LookAt:=xlWhole).Offset(0, -5)
    
    If codigoNis = "" Then
        ' Todo salió mal
        Sheets("Planilla").Activate
        MsgBox ("El código postal " & cp & " no existe. Intentá con otro. " & codigoNis)
        Exit Function
    End If
End Function
Sub marcar()
    ' MARCA CON UN COLOR TODA UNA SELECCION. O LO VUELVE A LA NORMALIDAD
    Dim rojo, verde, azul, inferior, superior, filas As Byte
    
    inferior = 150
    superior = 255
    rojo = Int(inferior + Rnd * (superior - inferior + 1))
    verde = Int(inferior + Rnd * (superior - inferior + 1))
    azul = Int(inferior + Rnd * (superior - inferior + 1))
    filas = Selection.Rows.Count - 1
    
    If Selection.Row > 1 And (Selection.Row + filas) < 35 Then
        ActiveSheet.Unprotect "Rerda"
        Selection.Interior.color = RGB(rojo, verde, azul)
        ActiveSheet.Protect "Rerda"
    ElseIf Selection.Row = 1 Or (Selection.Row + filas) > 34 Then
        ' NADA
        MsgBox "La 1° fila de títulos no se selecciona.", vbCritical, "¡Guarda!"
    Else
        MsgBox "Tenés que seleccionar algo antes " & vbNewLine & "que esté entre el título y el pié."
    End If
    Debug.Print Selection.Row + filas
End Sub
Sub desmarcar()
    ' VUELVE A LA NORMALIDAD
    If Selection.Row = 1 Then
        MsgBox "La 1° fila de títulos no se selecciona.", vbCritical, "¡Guarda!"
    Else
        ActiveSheet.Unprotect "Rerda"
        Selection.Interior.ColorIndex = xlNone
        ActiveSheet.Protect "Rerda"
    End If
End Sub
' Sirve para controlar si corresponde o no una factura proforma
Function proforma(apellidoNombre, direccion, provincia, codigoPostal, ciudad, telefono, fecha)
    Dim acumulador As Byte
    Dim cantidad As Byte
    Dim precio As Double
    Dim sku As String
    Dim color As String
    Dim talle As String
    Dim cotizacion As Double
    cotizacion = 0
        
    
    
    'Limpiando información previa
    With Sheets("Proforma")
        .Range("A21:D49").ClearContents
        .Range("H21:H49").ClearContents
        .Range("I7:I14").ClearContents
        .Range("I17:I18").ClearContents
    End With
    
    ' Dando formato a la página para imprimir
    With Sheets("Proforma").PageSetup
        .Orientation = xlPortrait
        .TopMargin = Application.CentimetersToPoints(1.9)
        .RightMargin = Application.CentimetersToPoints(0.6)
        .BottomMargin = Application.CentimetersToPoints(1.9)
        .LeftMargin = Application.CentimetersToPoints(0.6)
        .HeaderMargin = Application.CentimetersToPoints(0.8)
        .FooterMargin = Application.CentimetersToPoints(0.8)
        .CenterHorizontally = True
        .PaperSize = xlPaperA4
    End With
    
    If provincia = "TIERRA DEL FUEGO" Or provincia = "Tierra del Fuego" Then
        ' Pide cotización del dólar
        Do While cotizacion = 0
            cotizacion = Application.InputBox(Prompt:="Cotización del dólar", Title:="Factura Proforma", Default:=1)
        Loop
        
        'hacer proforma
        With Sheets("Proforma")
            .Cells(7, 9).Value = UCase(apellidoNombre)
            .Cells(9, 9).Value = direccion
            .Cells(11, 9).Value = UCase(ciudad)
            .Cells(12, 9).Value = codigoPostal
            .Cells(13, 9).Value = UCase(provincia)
            .Cells(17, 9).Value = "'" & telefono
        End With
        acumulador = 0
        
        ' Bucle que recorre la venta
        Worksheets("Planilla").Activate
        ActiveCell.Offset(0, -19).Activate
                
        Do While ActiveCell.Offset(0, -1).Value = apellidoNombre Or ActiveCell.Offset(0, -1).Value = ""
        
            ' Toma datos de la "Planilla"
            With Sheets("Planilla")
                sku = ActiveCell.Value
                cantidad = ActiveCell.Offset(0, 4).Value
                color = ActiveCell.Offset(0, 3).Value
                talle = ActiveCell.Offset(0, 2).Value
                precio = ActiveCell.Offset(0, 5).Value
            End With
            Debug.Print sku, cantidad, color, talle, precio
            
            ' Los vuelca en la Proforma
            With Sheets("Proforma")
                ' Copiando el código
                .Cells(21 + acumulador, 2).Value = sku
            
                ' Copiando el talle
                .Cells(21 + acumulador, 4).Value = talle
            
                ' Copiando el color
                .Cells(21 + acumulador, 3).Value = color
            
                ' Copiando la cantidad
                .Cells(21 + acumulador, 1).Value = cantidad
            
                ' Copiando el monto
                .Cells(21 + acumulador, 8).Value = precio / cotizacion
            End With
                        
            ' Aumentando el contador
            acumulador = acumulador + 1
            
            ' Saltando de celda
            ActiveCell.Offset(1, 0).Activate
            
            ' Salir cuando se acabe todo
            If ActiveCell.Value = "" Then
                Exit Do
            End If
        Loop
        
        ' Imprimiendo el rótulo
        Call Rotulo_Correo_Argentino("Factura Proforma - " & apellidoNombre, Sheets("Proforma"), fecha)
    End If
End Function

Sub CorreoArgentino()
' GENERADOR DE HOJA DEL CORREO ARGENTINO

' Variables necesarias
Dim planillaGeneradora As Workbook
Dim nombreArchivoCorreo As String
Dim ruta As String
Dim nombreCarpeta As String
Dim i As Byte
Dim e As Byte
Dim comienzo As Byte

e = 0
comienzo = 13

Set planillaGeneradora = ThisWorkbook


nombreArchivoCorreo = "CorreoArgentino.xlsm"


ruta = planillaGeneradora.Path & "\" & nombreArchivoCorreo

' Controla si existe o no el archivo en la misma carpeta
If Dir(ruta) = "" Then
    MsgBox "El archivo " & nombreArchivoCorreo & " debe estar en la misma carpeta que esta planilla."
    Exit Sub
End If

' Abrir el archivo
Workbooks.Open ruta

' Borra el contenido
For i = 13 To 32
    Cells(i, 3) = ""
    Cells(i, 10) = ""
Next i

' Copia contenido: Denominación - DNI/CUIT: - CP: - Provincia - Vendedor/Viajante
    
    ' Si hay más de 20, se genera un nuevo archivo.
    For i = 2 To 34
        If planillaGeneradora.Worksheets("Planilla").Cells(i, 2) <> "" Then
            ' Datos del nombre/apellido y demás
            Cells(comienzo + e, 3).Value = UCase(planillaGeneradora.Worksheets("Planilla").Cells(i, 2).Value) & " - DNI/CUIT: " & planillaGeneradora.Worksheets("Planilla").Cells(i, 12).Value & " - CP " & planillaGeneradora.Worksheets("Planilla").Cells(i, 19).Value & " - " & planillaGeneradora.Worksheets("Planilla").Cells(i, 21).Value
            
            ' Dato del Vendedor
            Cells(comienzo + e, 10).Value = planillaGeneradora.Worksheets("Planilla").Cells(2, 23).Value
        
            ' Incrementamos en 1 el contador "e"
            e = e + 1
        End If
        
        
    
    
        If e > 20 Then
            ThisWorkbook.Save
            MsgBox "Te sobrepasaste de 20 renglones. Guardá este documento y hacé otro más."
            Exit Sub
        End If
    
    Next i

End Sub


Sub deposito()
' GENERA UNA PLANILLA SÓLO PARA USO EXCLUSIVO DEL DEPOSITO
' Declaración de Variables y su tipo de datos
Dim ruta As String
Dim ultima As Byte
Dim i As Byte
Dim NombreDeArchivo As String
Dim ExisteArchivo As String
Dim nombre As String
Dim nombreCarpeta As String

' Variables necesarias. Un nivel más arriba
nombreCarpeta = ThisWorkbook.Path & "\..\"
ultima = 34
ruta = "'" & nombreCarpeta & "[Stock.XLS]Sheet1'!$A$2:$G$10000"

NombreDeArchivo = nombreCarpeta & "Stock.XLS"
ExisteArchivo = Dir(NombreDeArchivo)

' Comprueba si existe el archivo Stock.XLS
If ExisteArchivo = "" Then
    MsgBox "El archivo Stock.XLS debe estar en la misma carpeta que esta planilla"
    Exit Sub
End If

Sheets("Depósito").Activate

' Reemplaza la que hubiere



' Creando las columnas
Cells(1, 1).Value = "Cliente"
Cells(1, 2).Value = "Descripción"
Cells(1, 3).Value = "Código"
Cells(1, 4).Value = "Color"
Cells(1, 5).Value = "Talle"
Cells(1, 6).Value = "Cantidad"
Cells(1, 7).Value = "Ubicación"

' Completando los datos
For i = 2 To ultima
    ' Cliente
    Cells(i, 1).Value = Sheets("Planilla").Cells(i, 2).Value
    
    ' Descripción. Sólo si hay código
    If Sheets("Planilla").Cells(i, 3).Value = "" Then
        Cells(i, 2).ClearContents
    Else
        Cells(i, 2).Value = "=VLOOKUP(C" & i & "," & ruta & ", 2, FALSE)"
    End If
    
    ' Código
    Cells(i, 3).Value = "'" & Sheets("Planilla").Cells(i, 3).Value
    
    ' Color
    Cells(i, 4).Value = Sheets("Planilla").Cells(i, 6).Value
    
    ' Talle
    Cells(i, 5).Value = Sheets("Planilla").Cells(i, 5).Value
    
    ' Cantidad
    Cells(i, 6).Value = Sheets("Planilla").Cells(i, 7).Value
        
    ' La ubicación
    If Cells(i, 3) = "" Then
        Cells(i, 7).Value = ""
    Else
        Cells(i, 7).Formula = "=VLOOKUP(C" & i & "," & ruta & ", 3, FALSE)"
    End If
Next i

' Ordenando alfabéticamente esta columna de ubicación
With Range("A1:G1")
    .AutoFilter
    .Rows("1").RowHeight = 27
    .Font.Bold = True
    .Font.Size = 12
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
End With

Range("A1").CurrentRegion.Sort Key1:=Range("G1"), Order1:=xlAscending, Header:=xlGuess
With Selection
    .AutoFilter
End With

With Range("A1").CurrentRegion
    .Columns.AutoFit
End With

' Colocando totales de productos y dando formato
Cells(ultima + 1, 5).Value = "TOTALES:"
Cells(ultima + 1, 6).Select
Cells(ultima + 1, 6).Value = "=SUM(F2:F" & ultima & ")"
Range(Cells(ultima + 1, 5), Cells(ultima + 1, 6)).Select
With Selection
    .Font.Bold = True
    .Font.Size = 15
    .HorizontalAlignment = xlRight
    .VerticalAlignment = xlBottom
End With


' Colocando el total de rótulos a imprimir
Cells(ultima + 1, 2).Value = "ROTULOS:"
Cells(ultima + 1, 3).Value = "=COUNTA(A2:A" & ultima & ")"
Range(Cells(ultima + 1, 2), Cells(ultima + 1, 3)).Select
With Selection
    .Font.Bold = True
    .Font.Size = 15
    .VerticalAlignment = xlBottom
    .HorizontalAlignment = xlRight
End With
Cells(ultima + 1, 3).HorizontalAlignment = xlLeft

With Range("A1").CurrentRegion
    .Borders.LineStyle = xlContinuous
End With

' Formato de impresión
With ActiveSheet.PageSetup
    .Orientation = xlPortrait
    .PaperSize = xlPaperA4
    .LeftMargin = Application.CentimetersToPoints(0.64)
    .RightMargin = Application.CentimetersToPoints(0.64)
    .TopMargin = Application.CentimetersToPoints(4)
    .BottomMargin = Application.CentimetersToPoints(1.91)
    .HeaderMargin = Application.CentimetersToPoints(0.76)
    .FooterMargin = Application.CentimetersToPoints(0.76)
    .CenterHorizontally = True
    .CenterVertically = False
    .PrintArea = ActiveSheet.Range("A1:G" & ultima + 1).Address
    .Zoom = False
    .FitToPagesTall = 1
    .FitToPagesWide = 1
    .CenterHeader = "&B&20&F" & vbNewLine & "SOLO PARA USO EN DEPOSITO"
End With

' Guarda cambios
ThisWorkbook.Save

' Exporta un txt
Call exportarTxt
End Sub
Sub construirCtaCte()
' CONSTRUYE LA PLANILLA DE CUENTA CORRIENTE

' Nombre Planilla de Ventas
Dim planillaVentas As String
planillaVentas = ThisWorkbook.Name

' Carpeta actual de Planilla de Ventas
Dim carpetaActual As String
carpetaActual = Workbooks(planillaVentas).Path

' Creación archivo Cuentas Corrientes
Dim archivoCtaCte As String
archivoCtaCte = Year(Date) & ". CTACTE.xlsx"

' Control si existe de antes
If Dir(carpetaActual & "\..\" & archivoCtaCte, vbNormal) = "" Then
    Workbooks.Add.SaveAs fileName:=(carpetaActual & "\..\" & archivoCtaCte)
Else
    ' Workbooks.Open (carpetaActual & "\..\" & archivoCtaCte)
    MsgBox "Ya está creada la planilla de las cuentas corrientes de antes." & vbNewLine & "Esperá al año que viene."
    Exit Sub
End If


' Arreglo con la definición pernsolizada de los meses
Dim arrayMeses As Variant
arrayMeses = Array("ENE", "FEB", "MAR", "ABR", "MAY", "JUN", "JUL", "AGO", "SEP", "OCT", "NOV", "DIC")

' Construcción de las hojas
Dim i As Byte
i = 12
Do While i <= 12
    If i = 0 Then
        GoTo Salida
        ' Sale del bucle
    ElseIf i = 12 Then
        ActiveSheet.Name = arrayMeses(i - 1) & "-" & arrayMeses(0)
    Else
        Worksheets.Add.Name = arrayMeses(i - 1) & "-" & arrayMeses(i)
    End If
    i = i - 1
    Call construirHoja
Loop
Salida:

Workbooks(archivoCtaCte).Close SaveChanges:=True
End Sub

Function construirHoja()
Dim titulares As Variant
Dim i As Byte
titulares = Array("FAC NUM", "FECHA", "DNI", "CUENTA", "CBU", "CLIENTE", "IMPO.TOTAL", "CUOTAS", "IMP.CUOTA", "TELEFONO", "DOMICILIO", "LOCALIDAD", "PROVINCIA", "VENDEDOR")
With ActiveSheet
    For i = 0 To UBound(titulares)
        With Cells(1, i + 1)
            .Value = titulares(i)
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .Interior.ColorIndex = 15
            .Borders.LineStyle = xlContinuous
        End With
    Next i
End With

End Function

Sub CTACTE()
' CARGA DATOS DE LA PLANILLA EN EL ARCHIVO DE LAS CUENTAS CORRIENTES
Dim a As String
' Nombre Planilla de Ventas
Static planillaVentas As String
planillaVentas = ThisWorkbook.Name

' Carpeta actual de Planilla de Ventas
Dim carpetaActual As String
carpetaActual = Workbooks(planillaVentas).Path

' Nombre archivo Cuentas Corrientes
Dim archivoCtaCte As String
archivoCtaCte = Year(Date) & ". CTACTE.xlsx"

' Pestaña del archivo Cuentas Corrientes
Dim arrayMeses As Variant
arrayMeses = Array("ENE-FEB", "FEB-MAR", "MAR-ABR", "ABR-MAY", "MAY-JUN", "JUN-JUL", "JUL-AGO", "AGO-SEP", "SEP-OCT", "OCT-NOV", "NOV-DIC", "DIC-ENE")
Dim mesCtaCte As String

' Sirve para posicionar la pestaña al comenzar
If Day(Date) <= 19 Then
    mesCtaCte = Month(Date + 1)
Else
    mesCtaCte = Month(Date)
End If
Debug.Print mesCtaCte

' Abrir archivo Cuentas Corrientes
On Error GoTo ManejoError
Workbooks.Open(carpetaActual & "\..\" & archivoCtaCte).Sheets(arrayMeses(mesCtaCte - 1)).Activate


' Ultima fila de la pestaña actual
Dim ultimaFila As Byte

' Copiando información de la planilla al archivo
Dim i As Byte
For i = 2 To 34
    ' Siempre actualizar la última fila
    ultimaFila = Workbooks(archivoCtaCte).Sheets(arrayMeses(mesCtaCte - 1)).Cells(Rows.Count, 1).End(xlUp).Row + 1
    Debug.Print ultimaFila
    Debug.Print Workbooks(planillaVentas).Sheets("Planilla").Cells(i, 13).Value
    
    If Workbooks(planillaVentas).Worksheets("Planilla").Cells(i, 13).Value = Workbooks(planillaVentas).Sheets("Opciones").Range("C2").Value Then
        
        With Workbooks(archivoCtaCte).Sheets(arrayMeses(mesCtaCte - 1))
            ' N° Factura
            If Workbooks(planillaVentas).Sheets("Planilla").Cells(i, 11).Value <> "" Then
                .Cells(ultimaFila, 1).Value = Workbooks(planillaVentas).Sheets("Planilla").Cells(i, 11).Value
            Else
                MsgBox "Te falta facturar!!"
                Workbooks.Open(carpetaActual & "\..\" & archivoCtaCte).Close
                Workbooks(planillaVentas).Worksheets("Planilla").Cells(i, 11).Activate
                Exit Sub
            End If
            
            ' Fecha
            .Cells(ultimaFila, 2).Value = Workbooks(planillaVentas).Sheets("Planilla").Cells(i, 1).Value
            
            ' DNI
            .Cells(ultimaFila, 3).Value = Workbooks(planillaVentas).Sheets("Planilla").Cells(i, 12).Value
            
            ' N° Cuenta de Caja de Ahorro
            .Cells(ultimaFila, 4).Value = "'" & Workbooks(planillaVentas).Sheets("Planilla").Cells(i, 15).Value
            
            ' CBU
            .Cells(ultimaFila, 5).Value = "'" & Workbooks(planillaVentas).Sheets("Planilla").Cells(i, 16).Value
            
            ' Cliente
            .Cells(ultimaFila, 6).Value = Workbooks(planillaVentas).Sheets("Planilla").Cells(i, 2).Value
            
            ' Importe total de la factura
            .Cells(ultimaFila, 7).Value = CCur(Workbooks(planillaVentas).Sheets("Planilla").Cells(i, 9).Value)
            
            ' Cantidad de Cuotas
            .Cells(ultimaFila, 8).Value = Workbooks(planillaVentas).Sheets("Planilla").Cells(i, 14).Value
            
            ' Importe de la cuota
            .Cells(ultimaFila, 9).Value = CCur(Cells(ultimaFila, 7).Value / Cells(ultimaFila, 8).Value)
            
            ' Teléfono
            .Cells(ultimaFila, 10).Value = Workbooks(planillaVentas).Sheets("Planilla").Cells(i, 18).Value
            
            ' Domicilio
            .Cells(ultimaFila, 11).Value = Workbooks(planillaVentas).Sheets("Planilla").Cells(i, 17).Value & " - CP: " & Workbooks(planillaVentas).Sheets("Planilla").Cells(i, 19).Value
            
            ' Localidad
            .Cells(ultimaFila, 12).Value = Workbooks(planillaVentas).Sheets("Planilla").Cells(i, 20).Value
            
            ' Provincia
            .Cells(ultimaFila, 13).Value = Workbooks(planillaVentas).Sheets("Planilla").Cells(i, 21).Value
            
            ' Vendedor
            .Cells(ultimaFila, 14).Value = Workbooks(planillaVentas).Sheets("Planilla").Cells(2, 23).Value
        End With
    End If
Next i

ActiveSheet.UsedRange.EntireColumn.AutoFit
ActiveSheet.UsedRange.EntireRow.AutoFit
Range("A1").Activate

ManejoError:
    ' En caso de error, mostrar mensaje y finalizar la macro
    MsgBox "Ufa! Ele! Ubicá bien tu planilla:" & vbNewLine & "Así ->" & vbNewLine & "\VENTAS AÑO\VENTAS MES\" & planillaVentas, vbCritical
    Exit Sub

End Sub


Sub exportarTxt()

' GENERA UN ARCHIVO DE TEXTO PARA IMPORTAR AL D.F.
Dim fila As Long
Dim columna As Long

Dim txt As String
Dim textoArchivo As String
Dim server As String
Dim carpetaDestino As String
Dim carpetaActual As String
Dim nombreArchivo As String

Dim largo As Byte
Dim limite As Byte
Dim i As Byte
Dim ultimaFila As Byte
Dim resto As Byte
Dim cantArchivos As Byte

Dim planillaVentas As Object

Dim txtTemporal As Workbook


' Planilla de Ventas

Set planillaVentas = ActiveWorkbook
carpetaActual = planillaVentas.Path

nombreArchivo = Len(planillaVentas.Name)
ultimaFila = planillaVentas.Sheets("Depósito").Cells(Rows.Count, 2).End(xlUp).Row - 1
server = "\\SER-DF\d\A Remitar TXT\"
carpetaDestino = planillaVentas.Sheets("Planilla").Cells(2, Columns.Count).End(xlToLeft).Value & "\"
Debug.Print nombreArchivo; ultimaFila; carpetaDestino


' Crea un archivo temporal
txt = "TXT Temporal"
Application.DisplayAlerts = False
Set txtTemporal = Workbooks.Add
txtTemporal.SaveAs (planillaVentas.Path & "\" & txt)
Application.DisplayAlerts = True
limite = 30

' Separación de talles y colores ==========
' Completa planilla para exportar
For fila = 2 To ultimaFila
    
    ' 1º) Stock
    'planillaVentas.Sheets("Depósito").Cells(fila, 6).Value = "'" & txtTemporal.Sheets(1).Cells(fila - 1, 1).Value
    txtTemporal.Sheets(1).Cells(fila - 1, 1).Value = planillaVentas.Sheets("Depósito").Cells(fila, 6).Value
    
    ' 2º Codigo
    'planillaVentas.Sheets("Depósito").Cells(fila, 3).Value = "'" & txtTemporal.Sheets(1).Cells(fila - 1, 2).Value
    txtTemporal.Sheets(1).Cells(fila - 1, 2).Value = "'" & planillaVentas.Sheets("Depósito").Cells(fila, 3).Value
    
    ' 3° Color
    'planillaVentas.Sheets("Depósito").Cells(fila, 4).Value = "'" & Left(txtTemporal.Sheets(1).Cells(fila - 1, 3).Value, InStr(txtTemporal.Sheets(1).Cells(fila - 1, 3).Value, "."))
    largo = InStr(planillaVentas.Sheets("Depósito").Cells(fila, 4).Value, ".")
    If largo > 1 Then
        largo = largo - 1
    End If
    txtTemporal.Sheets(1).Cells(fila - 1, 3).Value = "'" & Left(planillaVentas.Sheets("Depósito").Cells(fila, 4).Value, largo)
    
    ' 4° Talle
    'planillaVentas.Sheets("Depósito").Cells(fila, 3).Value = "'" & txtTemporal.Sheets(1).Cells(fila - 1, 4).Value
    txtTemporal.Sheets(1).Cells(fila - 1, 4).Value = "'" & planillaVentas.Sheets("Depósito").Cells(fila, 5).Value
Next fila

' Ajuste ultima Fila - HARDCODEO ESTO PARA PROBAR
ultimaFila = txtTemporal.Sheets(1).Cells(Rows.Count, 1).End(xlUp).Row
Debug.Print ultimaFila

' Si se pasa del tope (30 líneas), serán "n" archivos con 30 líneas y otro con el resto
' de items que quedaron fuera. Sería el resto de una división, el módulo.
resto = ultimaFila Mod limite
cantArchivos = Int(ultimaFila / limite) + 1
Debug.Print "Archivos a importar: " & cantArchivos


' Generación del txt
Call generarTxt(fila - 1, ultimaFila, "", cantArchivos, planillaVentas.Name, carpetaDestino, limite, resto, server)

' Eliminación del archivo temporal
Debug.Print txtTemporal.Name

txtTemporal.Close (True)

Kill (planillaVentas.Path & "\" & txt & ".xlsx")


End Sub

Function generarTxt(fila, ultimaFila, textoArchivo, cantArchivos, nombreArchivo, carpetaDestino, limite, resto, server)
Dim rutaArchivo As String
Dim i As Byte
Dim tope As Byte
fila = 0


' Generación del txt
For i = 1 To cantArchivos
tope = i * limite
    If i = cantArchivos Then
        tope = ultimaFila
    End If
    
    For fila = (limite * (i - 1)) + 1 To tope
        Cells(fila, 1).Activate
        textoArchivo = textoArchivo _
            & Cells(fila, 1).Value _
            & "+" & Cells(fila, 2).Value _
            & "!" & Cells(fila, 3).Value _
            & "!" & Cells(fila, 4).Value _
            & vbNewLine
            Debug.Print "Archivo N°: " & i, "Fila N° :" & fila
    Next fila
    
    ' Si es mayor a uno, se van nombrando incrementalmente
    If cantArchivos > 1 Then
        nombreArchivo = Left(nombreArchivo, Len(nombreArchivo) - 5) & " - " & i & ".txt"
    Else
        nombreArchivo = Left(nombreArchivo, Len(nombreArchivo) - 5) & ".txt"
    End If
    
    rutaArchivo = server & carpetaDestino & nombreArchivo
    Debug.Print textoArchivo
    Open rutaArchivo For Output As #1
    Print #1, textoArchivo
    Close #1
    
    MsgBox "Datos exportados con éxito a " & rutaArchivo, vbInformation, "Cargar detalle desde txt"
Next i



End Function


