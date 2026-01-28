Attribute VB_Name = "WebReformada"
Option Explicit
Public ruta As String

Function PintarFila(Hojilla As String, Fila As Integer, DesdeColumna As Integer, HastaColumna As Integer)
    ' Pinta filas impares
    If Fila Mod 2 <> 0 Then
        Worksheets(Hojilla).Range(Cells(Fila, DesdeColumna), Cells(Fila, HastaColumna)).Interior.color = RGB(240, 240, 240)
    End If
End Function

'Sub CorreoArgentino(packar As Workbook, VentaWeb As Workbook)
Sub CorreoArgentino()

' GENERADOR DE HOJA DEL CORREO ARGENTINO
Dim VentaWeb As Workbook
Dim packar As Workbook
Dim nombreArchivoCorreo As String
Dim ruta As String
Dim nombreCarpeta As String
Dim i As Byte
Dim e As Byte
Dim comienzo As Byte
Dim ultima As Integer

Set VentaWeb = ActiveWorkbook
ultima = VentaWeb.Worksheets("ventas").Cells(Rows.Count, 2).End(xlUp).Row - 1
Debug.Print ultima

e = 0
comienzo = 13


ruta = VentaWeb.Path

Workbooks.Open ruta & "\..\CorreoArgentino.xlsm"
Set packar = ActiveWorkbook
nombreArchivoCorreo = packar.Name

' Borra el contenido
For i = 13 To 32
    With packar.Worksheets("Correo Argentino")
        .Cells(i, 3).Activate
        .Cells(i, 3) = ""
        .Cells(i, 10) = ""
    End With
Next i

' Copia contenido: Denominación - DNI/CUIT: - CP: - Provincia - Vendedor/Viajante
    
    ' Si hay más de 20, se genera un nuevo archivo.
    For i = 2 To ultima
        If VentaWeb.Worksheets("ventas").Cells(i, 1).Value <> "" And VentaWeb.Worksheets("ventas").Cells(i, 9).Value <> "Retira en Local" Then
            ' Datos del nombre/apellido y demás
            packar.Worksheets("Correo Argentino").Cells(comienzo + e, 3).Value = UCase(VentaWeb.Worksheets("ventas").Cells(i, 2).Value) & " - DNI/CUIT: " & VentaWeb.Worksheets("ventas").Cells(i, 7).Value & " - CP " & VentaWeb.Worksheets("ventas").Cells(i, 11).Value & " - " & VentaWeb.Worksheets("ventas").Cells(i, 12).Value
            
            ' Dato del Vendedor
            packar.Worksheets("Correo Argentino").Cells(comienzo + e, 10).Value = "Venta Web Nº #" & VentaWeb.Worksheets("ventas").Cells(i, 1).Value
        
            ' Incrementamos en 1 el contador "e"
            e = e + 1
            Debug.Print e
        End If
        
    
        If e > 20 Then
            packar.Save
            MsgBox "Te sobrepasaste de 20 renglones. Guardá este documento y hacé otro más."
            Exit Sub
        End If
    
    Next i

End Sub


Function generarRoutuloRetiro(ruta As String, nombre As String, telefono As String, dni As String, fecha As String, Venta As String)
    Dim numVenta As Worksheet
    Set numVenta = Worksheets(Venta)
    
    ' GENERA PESTAÑAS CON ROTULOS PARA RETIRO EN LOCAL
    ' Enmarcando
    numVenta.Range("A1:H21").Select
    With Selection
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With
    
    ' Formato de impresión
    With numVenta.PageSetup
        .Orientation = xlPortrait
        .PaperSize = xlPaperA4
        .LeftMargin = Application.CentimetersToPoints(0.64)
        .RightMargin = Application.CentimetersToPoints(0.64)
        .TopMargin = Application.CentimetersToPoints(2.5)
        .BottomMargin = Application.CentimetersToPoints(1.91)
        .HeaderMargin = Application.CentimetersToPoints(0.76)
        .FooterMargin = Application.CentimetersToPoints(0.76)
        .CenterHorizontally = True
        .CenterVertically = False
        .PrintArea = "$A$1:$H$21"
        .Zoom = False
        .FitToPagesTall = 1
        .FitToPagesWide = 1
    End With
    
   
    ' Dando un alto a la fila
    numVenta.Range("A2:A2").RowHeight = 30
    
     ' Insertando la imagen
    'On Error Resume Next
    numVenta.Pictures.Insert(ruta & "..\logo.png").Select
    
    ' Centrando el logo
    With Selection
        .Top = 4
        .Left = 155
    End With
    
   
    ' Leyenda de retiro
    numVenta.Range("A4:H6").Select
    With Selection
        .Merge
        .Font.Size = 30
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    numVenta.Range("A4").Value = "RETIRA EN ENTREPISO"
    
    ' Nombre del cliente, en mayúsculas
    numVenta.Range("A9:H10").Select
    With Selection
        .Merge
        .Font.Size = 25
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Interior.color = RGB(220, 220, 220)
    End With
    numVenta.Range("A9").Value = UCase(nombre)
    
    ' dni/cuit del cliente
    With numVenta
        .Range("A12").Value = "DNI/CUIT:"
        .Range("A12").HorizontalAlignment = xlRight
        .Range("A12:B12").Font.Bold = True
        .Range("A12:B12").Font.Size = 13
        .Range("B12").Value = "'" & dni
    End With
    
    
    
    ' Teléfono del cliente
    numVenta.Range("a14:d14").Select
    With Selection
        .Font.Bold = True
    End With
    
    With numVenta
        .Range("a14").Value = "TELEFONO:"
        .Range("a14").HorizontalAlignment = xlRight
        .Range("b14").HorizontalAlignment = xlLeft
        .Range("b14").Value = telefono
    End With
    
    
    ' FIRMA
    numVenta.Range("A20").Select
    With Selection
        .Value = "FIRMA:"
        .HorizontalAlignment = xlRight
        .Font.Bold = True
    End With
    numVenta.Range("b20:d20").Borders(xlEdgeBottom).LineStyle = xlContinuous
    
    ' FECHA
    numVenta.Range("f20").Select
    With Selection
        .Value = "FECHA RETIRO:"
        .HorizontalAlignment = xlRight
        .Font.Bold = True
    End With
    numVenta.Range("g20:h20").Borders(xlEdgeBottom).LineStyle = xlContinuous
    
    ' FECHA ELABORACIÓN
    With numVenta.Range("F12")
        .Value = "FECHA DE ELABORACIÓN:"
        .HorizontalAlignment = xlRight
    End With
    
    With numVenta.Range("G12")
        .Value = "'" & fecha
        .HorizontalAlignment = xlLeft
    End With
    
    numVenta.Range("F12:H12").Font.Bold = True
    
    ' NUMERO DE VENTA
    numVenta.Range("f14").Select
    With Selection
        .Value = "N° DE VENTA WEB:"
        .Font.Bold = True
        .HorizontalAlignment = xlRight
    End With
    
    numVenta.Range("g14").Select
    With Selection
        .Value = Venta
        .HorizontalAlignment = xlLeft
        .Font.Bold = True
        .Font.Size = 15
    End With
    
End Function

Function formatPrint(ultima, i)
' Dando formato apaisado, expandido a A4 y con titulares. Una sóla página.

' Delimitando el tamaño de hojas y márgenes
Dim filasTotales As Integer
filasTotales = ultima + 1

' Formatea la última columna que NO saldrá impresa, sólo para acomodar, nada más
Range("A:I").Columns.AutoFit

' Centrando el contenido
Range("E:E").HorizontalAlignment = xlCenter
Cells(ultima + 1, 5).HorizontalAlignment = xlRight

' Acomoda el texto de las celdas con datos
Columns("B").ColumnWidth = 40
Columns("D").ColumnWidth = 50
Columns("E").ColumnWidth = 14
'Range(Cells(2, 1), Cells(ultima, 11)).WrapText = True

' Ocultando columnas
Columns("G:H").Hidden = True

' Ajusta automáticamente la altura de las filas
Range(Cells(2, 1), Cells(ultima, 11)).Rows.AutoFit

' Formato de impresión
With ActiveSheet.PageSetup
    .Orientation = xlLandscape
    .PaperSize = xlPaperA4
    .LeftMargin = Application.CentimetersToPoints(0.64)
    .RightMargin = Application.CentimetersToPoints(0.64)
    .TopMargin = Application.CentimetersToPoints(2.5)
    ' Para evitar el blanco
    '.TopMargin = Application.CentimetersToPoints(5)
    .BottomMargin = Application.CentimetersToPoints(1.91)
    .HeaderMargin = Application.CentimetersToPoints(0.76)
    .FooterMargin = Application.CentimetersToPoints(0.76)
    .CenterHorizontally = True
    .CenterVertically = False
    .PrintArea = ActiveSheet.Range("A1:I" & ultima + 1).Address
    .Zoom = False
    .FitToPagesTall = 1
    .FitToPagesWide = 1
    .CenterHeader = "&B&20&F"
    ' Para espacio en blanco0
    '.CenterHeader = vbNewLine & vbNewLine & vbNewLine & "&B&20&F"
End With
    
End Function

Function formato(ultima, i)
' DA FORMATO A LA TABLA DE EXCEL PARA QUE SE VEA BONITA

Range("A1").CurrentRegion.Select
With Selection
    .Borders(xlEdgeLeft).LineStyle = xlContinuous
    .Borders(xlEdgeRight).LineStyle = xlContinuous
    .VerticalAlignment = xlTop
End With

' Formateando los encabezados
Rows("1").RowHeight = 27
Range("A1").Select
Range(Selection, Selection.End(xlToRight)).Select
    With Selection
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.color = RGB(250, 250, 250)
        .WrapText = True
    End With
    
' Agregando bordes VERTICALES a toda la tabla
Range("A2").CurrentRegion.Select
Selection.Borders(xlDiagonalDown).LineStyle = xlNone
Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With

' Agregando bordes HORIZONTALES en titular
Range("A1:M1").Select
Selection.Borders.LineStyle = xlContinuous

' Se cuentan cuantas celdas ocupadas hasta el final
Range(Cells(ultima, 1), Cells(ultima, 13)).Select
    With Selection
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With

' Colocando totales de productos y dando formato
Cells(ultima + 1, 3).Value = "TOTALES:"
Cells(ultima + 1, 6).Select
Cells(ultima + 1, 6).Value = "=SUM(F2:F" & ultima & ")"
Range(Cells(ultima + 1, 3), Cells(ultima + 1, 6)).Select
    With Selection
        .Font.Bold = True
        .Font.Size = 15
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlBottom
    End With


' Colocando el total de rótulos a imprimir
Cells(ultima + 1, 2).Value = "ROTULOS"
Cells(ultima + 1, 1).Value = "=COUNTA(A2:A" & ultima & ")"
Range(Cells(ultima + 1, 1), Cells(ultima + 1, 2)).Select
With Selection
    .Font.Bold = True
    .Font.Size = 15
    .VerticalAlignment = xlBottom
    .HorizontalAlignment = xlRight
End With
Cells(ultima + 1, 2).HorizontalAlignment = xlLeft

' Colocando un borde superior
For i = 3 To ultima
    Range(Cells(i, 1), Cells(i, 13)).Select
    If Cells(i, 1).Value <> "" Then
        With Selection
            .Borders(xlEdgeTop).LineStyle = xlContinuous
        End With
    End If
Next i

' Autofit para la última columna
Range("J:M").EntireColumn.AutoFit
Range("A:E").EntireColumn.AutoFit

' Ajustar el contenido
With Range("C:E").EntireColumn
    .WrapText = True
    .ShrinkToFit = True
End With

' Moviendo la columna del código
Columns("D").Cut
Columns("C").Insert Shift:=xlToRight

End Function

Function GuardarArchivo(fecha, ruta, nombreArchivo)
' VALIDANDO NOMBRE DE ARCHIVO A GENERAR Y GUARDAR

' Variables a utilizar
Dim nombre As String
Dim cuenta As String

' Asignando algunos valores
ruta = ruta & "WEB\"


'Controlando si la compu EDGAR está prendida y conectada a red.
If Dir(ruta, vbDirectory) = "" Then
    MsgBox ("No hay acceso la compu EDGAR. Debes prender esa compu y que se conecte a la red.")
    Exit Function
End If

' Definiendo unas variables
Dim archivos As String
Dim u As Integer
Dim denominacion As String
    
' Preparación de variables
u = 1
archivos = Dir(ruta)
    
' Recorrido de la carpeta
ActiveWorkbook.Sheets.Add(after:=ActiveWorkbook.ActiveSheet).Name = "Listado"
Sheets("Listado").Visible = False
Sheets(1).Name = "ventas"
Sheets("ventas").Select

Do While Len(archivos) > 0
    Sheets("Listado").Cells(u, 1).Value = archivos
    archivos = Dir()
    u = u + 1
Loop
nombre = ruta & Sheets("Listado").Cells(u - 1, 1).Value

' Controlando que no se esté duplicando el mismo archivo con otro nombre
If ActiveWorkbook.Name = Sheets("Listado").Cells(u - 1, 1).Value Then
    MsgBox ("Ya creaste este archivo antes. Generá uno nuevo.")
    ActiveWorkbook.Close SaveChanges:=False
    Exit Function
End If


' Guardando el archivo
Dim parteNumero As String
Dim nombreNumero As Integer
Dim e As Integer
e = 1
parteNumero = Mid(Sheets("Listado").Cells(u - 1, 1).Value, 11, 7)
nombreNumero = CInt(parteNumero) + 1
parteNumero = CStr(nombreNumero)
    
' Agregando ceros para tener un nombre coherente
Do While Len(parteNumero) < 6
    parteNumero = "0" & parteNumero
    e = e + 1
Loop
nombreArchivo = "Ventas Web " & parteNumero & ". " & fecha & ".xlsx"
nombre = ruta & nombreArchivo

Worksheets("ventas").Range("A1").Select
ActiveWorkbook.SaveAs Filename:=nombre, FileFormat:=xlOpenXMLWorkbook
Worksheets(1).Name = "ventas"
ActiveWorkbook.Save
Application.ThisWorkbook.Save

End Function


' Sirve para controlar si corresponde o no una factura proforma
Function proforma(apellidoNombre, direccion, provincia, codigoPostal, ciudad, telefono, fecha, RotulosCorreo, VentaWeb, ultimaFila, ruta)
    Dim acumulador As Byte
    Dim cantidad As Byte
    Dim precio As Double
    Dim sku As String
    Dim color As String
    Dim talle As String
    Dim cotizacion As Double
    Dim i As Integer
    cotizacion = 0
    
    
    'Limpiando información previa
    With RotulosCorreo.Worksheets("Proforma")
        .Range("A21:D49").ClearContents
        .Range("H21:H49").ClearContents
        .Range("I7:I14").ClearContents
        .Range("I17:I18").ClearContents
    End With
    
    ' Dando formato a la página para imprimir
    With RotulosCorreo.Worksheets("Proforma").PageSetup
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
        With RotulosCorreo.Worksheets("Proforma")
            .Cells(7, 9).Value = UCase(apellidoNombre)
            .Cells(9, 9).Value = direccion
            .Cells(11, 9).Value = UCase(ciudad)
            .Cells(12, 9).Value = codigoPostal
            .Cells(13, 9).Value = UCase(provincia)
            .Cells(17, 9).Value = "'" & telefono
        End With
        acumulador = 0
        
        ' Bucle que recorre la venta
        Do
            With VentaWeb.Worksheets("ventas")
                sku = .Cells(acumulador + 2, 3).Value
                cantidad = .Cells(acumulador + 2, 6).Value
                color = .Cells(acumulador + 2, 5).Value
                talle = .Cells(acumulador + 2, 5).Value
                precio = .Cells(acumulador + 2, 16).Value
            End With
            Debug.Print sku, cantidad, color, talle, precio

            
            ' Los vuelca en la Proforma
            With RotulosCorreo.Worksheets("Proforma")
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
            
        Loop While VentaWeb.Worksheets("ventas").Cells(acumulador + 2, 1).Value = "" And VentaWeb.Worksheets("ventas").Cells(acumulador + 2, 2).Value <> "ROTULOS"
                
        ' Imprimiendo el rótulo
        Call Rotulo_Correo_Argentino("Factura Proforma - " & apellidoNombre, RotulosCorreo.Worksheets("Proforma"), fecha, ruta)
    End If
End Function



Sub bucleRotular()
    ' Va rotulando por bucle
    Dim Fila As Integer
    Dim ultimaFila As Integer
    Dim RotulosCorreo As Workbook
    Dim VentaWeb As Workbook
    
    Set VentaWeb = ActiveWorkbook
    Debug.Print VentaWeb.Name
    
    ruta = VentaWeb.Path
    
    Workbooks.Open ruta & "\..\RotulosCorreo.xlsm"
    Set RotulosCorreo = ActiveWorkbook
    
    Debug.Print RotulosCorreo.Name
    
    ultimaFila = VentaWeb.Worksheets("ventas").Cells(Rows.Count, 2).End(xlUp).Row - 1
    Fila = 2
    VentaWeb.Worksheets("ventas").Activate
    VentaWeb.Worksheets("ventas").Range("M2").Activate
    
    For Fila = 2 To ultimaFila
        If VentaWeb.Worksheets("ventas").Cells(Fila, 13) <> "" Then
            Call Rotulador(VentaWeb, RotulosCorreo, Fila, ultimaFila, ruta)
        End If
    Next Fila
    
    
End Sub
Function Rotulador(VentaWeb As Workbook, RotulosCorreo As Workbook, Fila As Integer, ultimaFila As Integer, ruta As String)
    
    ' Declaración de Variables y su tipo de dato
    Dim Rotulos As Workbook
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
    
    Set Rotulos = RotulosCorreo

    fecha = Format(Date, "yyyy-mm-dd")
    
    
    With VentaWeb.Worksheets("ventas")
        apellidoNombre = .Cells(Fila, 2).Value
        dniCuit = .Cells(Fila, 7).Value
        direccion = .Cells(Fila, 14).Value & " " & .Cells(Fila, 15).Value
        codigoPostal = .Cells(Fila, 11).Value
        ciudad = .Cells(Fila, 10).Value
        provincia = .Cells(Fila, 12).Value
        telefono = .Cells(Fila, 8).Value
    End With
    
   
    ' SI ES A DOMICILIO -----------
    If VentaWeb.Worksheets("ventas").Cells(Fila, 13).Value = "Correo Argentino Clasico - Envio a domicilio" Then
        
        ' Colocar los datos
        With RotulosCorreo.Worksheets("A Domicilio")
            ' Colocando el número de Venta
            .Range("D7").Value = "Web Nº #" & VentaWeb.Worksheets("ventas").Cells(Fila, 1).Value
            .Range("C16").Value = UCase(apellidoNombre)
            .Range("P15").Value = dniCuit
            .Range("C18").Value = direccion
            .Range("C23").Value = UCase(provincia)
            .Range("E21").Value = codigoPostal
            .Range("G21").Value = UCase(ciudad)
            .Range("P23").Value = telefono
        End With
        
        ' Generar Proforma
        Call proforma(apellidoNombre, direccion, provincia, codigoPostal, ciudad, telefono, fecha, RotulosCorreo, VentaWeb, ultimaFila, ruta)
        
        ' Generar rotulo
        Call Rotulo_Correo_Argentino(apellidoNombre, RotulosCorreo.Worksheets("A Domicilio"), fecha, ruta)
        

    ' SI ES A RETIRAR EN SUCURSAL ----------------
    ElseIf VentaWeb.Worksheets("ventas").Cells(Fila, 13).Value = "Punto de retiro" Then
        
        ' Colocar los datos
        With RotulosCorreo.Worksheets("A Sucursal")
            .Range("C16").Value = UCase(apellidoNombre)
            .Range("R16").Value = dniCuit
            .Range("E7").Value = "Web Nº #" & VentaWeb.Worksheets("ventas").Cells(Fila, 1).Value
        End With
        
        ' Validando si existe o no el dato.
        On Error Resume Next
        codigoNis = RotulosCorreo.Worksheets("Sucursales").Range("$F$4:$F$5000").Find(what:=codigoPostal, LookIn:=xlValues, LookAt:=xlPart).Offset(0, -5)
    
        If codigoNis = "" Then
            ' Todo salió mal
            VentaWeb.Worksheets("ventas").Activate
            MsgBox ("El código postal " & codigoPostal & " no corresponde con ninguna sucursal del Correo. Intentá con otro. ")
            RotulosCorreo.Worksheets("Sucursales").Select
            MsgBox ("Buscá aquí un código postal de sucursal disponible")
            
            ' Marcar advertencia
            VentaWeb.Worksheets("ventas").Cells(Fila, 9).Value = "Buscar CP"
            Exit Function
        End If
        
        ' Completando el resto de datos.
        With RotulosCorreo.Worksheets("A Sucursal")
            .Range("S18").Value = codigoNis
            .Range("R22").Value = telefono
        End With
        
        ' Generar Proforma
        Call proforma(apellidoNombre, "Retiro en Sucursal del Correo Argentino Cód. NIS " & codigoNis, provincia, codigoPostal, ciudad, telefono, fecha, RotulosCorreo, VentaWeb, ultimaFila, ruta)
        
        ' Generar rotulo
        Call Rotulo_Correo_Argentino(apellidoNombre, RotulosCorreo.Worksheets("A Sucursal"), fecha, ruta)
    
    Else
        ' Rascarse las bolas
    
    End If


End Function


Function Rotulo_Correo_Argentino(apellidoNombre, rotulo, fecha, ruta)
'
' Rotulos Macro
' Guarda los rótulos en pdf
'
' Acceso directo: CTRL+MAY+Ñ
    
    ' Declaración de Variables y su tipo de datos
    Dim nombre As String
    Dim nombreCarpeta As String

    ' Variables necesarias
    nombre = fecha & ". " & apellidoNombre
    nombreCarpeta = "\Rotulos\"
    
    ' Comprobando si existe o no la carpeta
    If Dir(ruta, vbDirectory) <> "" Then
        If Dir(ruta & nombreCarpeta, vbDirectory) = "" Then
            MkDir ruta & nombreCarpeta
        End If
    End If
    
    ' Generando el archivo pdf
    rotulo.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
        ruta & nombreCarpeta & UCase(nombre) & ".pdf", _
        OpenAfterPublish:=True

End Function

Function Validar_CP(cp)
' Valida si el código postal es correcto, existe o no.

    codigoNis = RotulosCorreo.Worksheets("Sucursales").Range("F1576:F5000").Find(what:=cp, LookIn:=xlValues, searchorder:=xlByRows, LookAt:=xlWhole).Offset(0, -5)
    
    If codigoNis = "" Then
        ' Todo salió mal
        VentaWeb.Worksheets("ventas").Activate
        MsgBox ("El código postal " & cp & " no existe. Intentá con otro. " & codigoNis)
        Exit Function
    End If
End Function

Sub ventasWebReformada()

' Controlar que no se haya hecho formato antes
If ActiveSheet.Range("I1").Value = "Detalle" Then
    MsgBox ("Ya le diste formato a esta planilla. " & VBA.vbNewLine & "Probá con otra.")
    Range("A1").Select
    Exit Sub
End If

' Declarando variables a utilizar
Dim nombre As String
Dim telefono As String
Dim dni As String
Dim numVenta As String
Dim planilla As Object
Dim packar As Workbook
Dim ultima As Integer
Dim fecha As String
Dim i As Integer
Dim Rotulos As Integer
Dim ruta As String
Dim img As String
Dim nombreArchivo As String
Dim VentaWeb As Workbook
Dim RotulosCorreo As Workbook



' Sirve para averiguar el nombre de la computadora actual
Dim ws As Object
Set ws = CreateObject("WScript.network")

' Asignando algunos valores de acuerdo en qué equipo de la red esté
If ws.ComputerName = "EDGAR" Then
    ruta = "D:\Web\Listados de Ventas Online\"
    Debug.Print "Estoy en la computadora: " & ws.ComputerName
Else
    ruta = "\\EDGAR\Web\Listados de Ventas Online\"
    Debug.Print "Estoy en una computadora de la red, llamada: " & ws.ComputerName
End If
Debug.Print "Se guardan los archivos en: " & ruta

Rotulos = 0
fecha = Day(Date) & "-" & Month(Date) & "-" & Year(Date)


' Guardando el archivo con nombre específico
ActiveSheet.Name = "ventas"
Cells.Font.Size = 11
Call GuardarArchivo(fecha, ruta, nombreArchivo)
Range("A1").Activate
ultima = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row

' Borrando información innecesaria
Range("Y:Y").EntireColumn.Copy
Range("AO:AO").EntireColumn.PasteSpecial

' Copiando y pegando columna de los precios
Range("AG:AG").EntireColumn.Copy
Range("AT:AT").EntireColumn.PasteSpecial

Range("Q:R").EntireColumn.Copy
Range("AP:AQ").EntireColumn.PasteSpecial

Range("B:K, P:T, X:X, Z:AE, AG:AG, AI:AN").EntireColumn.Delete
Range("C:E").EntireColumn.Insert
Range("H:H").EntireColumn.Copy
Range("C:C").PasteSpecial
Range("A1").Value = "Venta"
Range("F:F").Select
Selection.NumberFormat = "General"
Range("G:G").Select
Selection.NumberFormat = "0"

' CORRIGIENDO NOMBRE DE CLIENTES
For i = 2 To ultima
    ' Mayúsculas en los nombres
    Cells(i, 2).Value = UCase(Cells(i, 2).Value)
    Cells(i, 8).Value = UCase(Cells(i, 8).Value)
    Cells(i, 3).Value = UCase(Cells(i, 3).Value)

    'Recorremos los nombres de los clientes
    If Cells(i, 2).Value <> Cells(i, 3).Value Then
        Cells(i, 2).Value = Cells(i, 2).Value & " - " & Cells(i, 3).Value
    End If
Next i
Range("B1").Value = "Cliente"

' Elimino la columna innecesaria
Range("C:C").EntireColumn.Delete
Range("G:G").EntireColumn.Delete
Range("C:D").EntireColumn.Insert

' Moviendo el detalle
Range("M:M").EntireColumn.Copy
Range("C:C").PasteSpecial
Range("M:M").EntireColumn.Delete
Range("C1").Value = "Descripción"

' Moviendo la cantidad
Range("M:M").EntireColumn.Copy
Range("F:F").PasteSpecial
Range("L:L").EntireColumn.Delete
Range("F1").Value = "Cantidad"

' Eliminando columna innecesaria
Range("P:Q").EntireColumn.Delete
'Range("D:E").EntireColumn.Delete
Range("L:L").EntireColumn.Delete



' Purgando los teléfonos
For i = 2 To ultima
    Cells(i, 8).Value = Right(Cells(i, 8).Value, 10)
Next i

' Insertando columna para la ubicación
Range("I:I").EntireColumn.Insert
Range("I1:I1").Value = "Detalles"

' Dando formato
'For i = 2 To ultima
    ' Columna de las ubicaciones de productos
    'With Cells(i, 9)
        '.WrapText = False
        '.ShrinkToFit = True
    'End With
'Next i

' Limpiando el contenido de "(sin color)" y "(sin talle)"
' Con comas y doble paréntesis
Cells.Replace what:="((Sin Color)), ", Replacement:="", LookAt:=xlPart, searchorder:= _
        xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Cells.Replace what:="((Sin Talle)), ", Replacement:="", LookAt:=xlPart, searchorder:= _
        xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False

' Doble paréntesis
Cells.Replace what:="((Sin Color))", Replacement:="", LookAt:=xlPart, searchorder:= _
        xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Cells.Replace what:="((Sin Talle))", Replacement:="", LookAt:=xlPart, searchorder:= _
        xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False

' Con comas
Cells.Replace what:="(Sin Color), ", Replacement:="", LookAt:=xlPart, searchorder:= _
        xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Cells.Replace what:="(Sin Talle), ", Replacement:="", LookAt:=xlPart, searchorder:= _
        xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False

' Simples
Cells.Replace what:="(Sin Color)", Replacement:="", LookAt:=xlPart, searchorder:= _
        xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Cells.Replace what:="(Sin Talle)", Replacement:="", LookAt:=xlPart, searchorder:= _
        xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False


' Generando las columnas de código/talle/color/cantidad
Range("C:C").Select
Selection.TextToColumns Destination:=Range("C1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="(", FieldInfo:=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
Range("A2").Activate
Cells.Replace what:=")", Replacement:="", LookAt:=xlPart, searchorder:= _
        xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Range("D1").Value = "Código"
Range("E1").Value = "Variante"


' Borra cosas innecesarias
Do While ActiveCell.Value <> ""
    If ActiveCell.Offset(0, 1) = "" Then
        ActiveCell.Value = ""
        ActiveCell.Offset(0, 13) = ""
    End If
    ActiveCell.Offset(1, 0).Activate
Loop

' DANDO FORMATO A TODA LA PLANILLA
Call formato(ultima, i)

' Agrandando columna de los códigos
Columns("C").ColumnWidth = 8
Range(Cells(2, 3), Cells(ultima - 1, 3)).WrapText = False

' Coloreando filas impares
For i = 3 To ultima - 1
    Call PintarFila("ventas", i, 1, 9)
Next i


' GENERANDO LOS ROTULOS DE RETIRO
Dim HojaRetiroLocal As String
For i = 2 To ultima
    If Worksheets("ventas").Cells(i, 13).Value = "Retiras en Rerda Mendoza" Or Worksheets(1).Cells(i, 13).Value = "Rerda S.A. - Sastrería Militar" Or Worksheets(1).Cells(i, 13).Value = "Local Rerda" Then
        
        ' Se coloca la leyenda en la celda
        Debug.Print Cells(i, 9).Value
        Cells(i, 9).Value = "Retira en Local"
        
        ' Variables a completar
        nombre = Cells(i, 2).Value
        telefono = Cells(i, 8).Value
        dni = Cells(i, 7).Value
        numVenta = Cells(i, 1).Value
        
        ' Contador de rótulos a imprimir
        Rotulos = Rotulos + 1
        
        ' Se agrega una pestaña para el respectivo rótulo
        HojaRetiroLocal = "Venta Nº " & numVenta
        ActiveWorkbook.Worksheets.Add(after:=ActiveWorkbook.Worksheets("ventas")).Name = HojaRetiroLocal
        
        ' Se genera el rótulo respectivo
        Call generarRoutuloRetiro(ruta, nombre, telefono, dni, fecha, HojaRetiroLocal)
        
    End If
    Worksheets("ventas").Activate
Next i

'Posicionando al principio
Worksheets.Add(after:=Worksheets("ventas")).Name = "Depósito"
Worksheets.Add(after:=Sheets("Depósito")).Name = "Exportar TXT"
Worksheets("ventas").Activate
Range("A1").Activate

' Definiendo este archivo
Set planilla = ActiveWorkbook

' DANDO FORMATO DE IMPRESION
Call formatPrint(ultima, i)



' Dando un aviso condicional sólo si hay rótulos, se lo contrario, no.
If Rotulos > 0 Then
    MsgBox ("Tenés " & Rotulos & " rótulos de retiro en local, para imprimir." & VBA.vbNewLine & "Aquí abajo en las pestañas.")
End If


' Generando una planilla sólo para dpto. DEPOSITO
Call deposito

' Guardando en un objeto este archivo
Set VentaWeb = ActiveWorkbook


' Generar la rotulación y la proforma de corresponder
ruta = ruta & "..\"

'Call bucleRotular

' Generando una planilla de informe para el Correo Argentino


' Posicionando al principio
planilla.Worksheets("ventas").Activate
ActiveWorkbook.Save
MsgBox "Proceso terminado"
End Sub

Sub deposito()
' GENERA UNA PLANILLA SÓLO PARA USO EXCLUSIVO DEL DEPOSITO
Dim ruta As String
Dim nombreArchivo As String
Dim ultima As Byte
Dim i As Integer
Dim enrutacion As String
Dim web As Boolean
enrutacion = ActiveWorkbook.Path & "\..\"
Debug.Print enrutacion

ruta = "'" & enrutacion & "[Stock.XLS]Sheet1'!$A$2:$G$10000"
ultima = Worksheets("ventas").Cells(Rows.Count, 2).End(xlUp).Row - 1
web = False


' Control si existe la pestaña "Depósito"
Call CrearHoja("Depósito")


' Creando las columnas
With Worksheets("Depósito")
    .Activate
    .Cells.Clear
    .Cells(1, 1).Value = "Nº Venta"
    .Cells(1, 2).Value = "Cliente"
    .Cells(1, 3).Value = "Descripción"
    .Cells(1, 4).Value = "Código"
    .Cells(1, 5).Value = "Variante"
    .Cells(1, 6).Value = "Cantidad"
    .Cells(1, 7).Value = "Ubicación"
End With

' Completando los datos
With Worksheets("Depósito")
    For i = 2 To ultima
        ' Venta
        .Cells(i, 1).Value = Worksheets("ventas").Cells(i, 1).Value
        
        ' Cliente
        .Cells(i, 2).Value = Worksheets("ventas").Cells(i, 2).Value
        
        ' Descripción
        .Cells(i, 3).Value = "=VLOOKUP(D" & i & "," & ruta & ",2,FALSE)"
        
        ' Reemplazo de la descripción en pestaña "ventas"
        Worksheets("ventas").Cells(i, 3).Value = "'" & Worksheets("ventas").Cells(i, 3).Value
        Worksheets("ventas").Cells(i, 4).Value = "=VLOOKUP(C" & i & "," & ruta & ",2,FALSE)"
        
        ' Código
        .Cells(i, 4).Value = "'" & Worksheets("ventas").Cells(i, 3).Value
        
        ' Variante
        .Cells(i, 5).Value = Worksheets("ventas").Cells(i, 5).Value
        
        ' Cantidad
        .Cells(i, 6).Value = Worksheets("ventas").Cells(i, 6).Value
            
        ' La ubiCación
        .Cells(i, 7).Formula = "=VLOOKUP(D" & i & "," & ruta & ",3,FALSE)"
    Next i
End With

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

' Moviendo la columna del código
Columns("D").Cut
Columns("C").Insert Shift:=xlToRight

' Ocultando columnas
Columns("A:B").Hidden = True

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

' Colocando bordes
With Range("A1").CurrentRegion
    .Borders.LineStyle = xlContinuous
End With

' Coloreando filas impares
For i = 3 To ultima
    Call PintarFila("Depósito", i, 3, 7)
Next i

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
    ' Le agrego saltos de línea para evitar el Espacio en Blanco
    '.CenterHeader = vbNewLine & vbNewLine & "&B&20&F" & vbNewLine & "SOLO PARA USO EN DEPOSITO"
End With
Call exportarTxt
Worksheets("ventas").Activate
Range("A1").Activate

End Sub

Function CrearHoja(nombreHoja As String) As Boolean
    ' controla si una hoja existe o no
    Dim existe As Boolean
     
    On Error Resume Next
    existe = (Worksheets(nombreHoja).Name <> "")
     
    If Not existe Then
        Worksheets.Add(after:=Worksheets(Worksheets.Count)).Name = nombreHoja
    End If
     
    CrearHoja = existe
     
End Function


Sub exportarTxt()

' GENERA UN ARCHIVO DE TEXTO PARA IMPORTAR AL D.F.

Dim Fila As Long, Columna As Long
Dim textoArchivo As String
Dim server As String

Dim carpetaDestino As String

Dim nombreArchivo As String
Dim limite As Byte
Dim item As Variant
Dim i As Byte
Dim ultimaFila As Byte
Dim resto As Byte
Dim cantArchivos As Byte
Dim RangoVariante As Range


Dim matrixCodColor As Object
ultimaFila = Worksheets("Depósito").Cells(Rows.Count, 6).End(xlUp).Row - 1
nombreArchivo = Len(ActiveWorkbook.Name)
server = "\\SER-DF\D\A Remitar TXT"
carpetaDestino = "\WEB\"
Set matrixCodColor = CreateObject("Scripting.Dictionary")
limite = 30

Worksheets("Exportar TXT").Activate

matrixCodColor.Add "01", "PITON GRIS"
matrixCodColor.Add "02", "PITON BEIGE"
matrixCodColor.Add "03", "BEIGE"
matrixCodColor.Add "04", "WOODLAND"
matrixCodColor.Add "05", "DIGITAL DESERT O DIGITAL DESERTICO"
matrixCodColor.Add "06", "VERDE"
matrixCodColor.Add "07", "ACU"
matrixCodColor.Add "08", "MULTICAM"
matrixCodColor.Add "09", "NEGRO"
matrixCodColor.Add "10", "TRI DESERT"
matrixCodColor.Add "11", "DIGITAL WOODLAND"
matrixCodColor.Add "12", "PITON VERDE"
matrixCodColor.Add "13", "GRIS"
matrixCodColor.Add "14", "REAL TREE"
matrixCodColor.Add "15", "DIGITAL TIGER WOODLAND"
matrixCodColor.Add "16", "DIGITAL FUERZA AEREA"
matrixCodColor.Add "17", "DIGITAL NAVAL"
matrixCodColor.Add "18", "AZUL"
matrixCodColor.Add "19", "CAMUFLADO"
matrixCodColor.Add "20", "BLANCO"
matrixCodColor.Add "21", "CPL MULTICAM BLACK"
matrixCodColor.Add "22", "ROJO"
matrixCodColor.Add "23", "DIGITAL RUSO"
matrixCodColor.Add "24", "BORDO"
matrixCodColor.Add "25", "CAMEL"
matrixCodColor.Add "26", "VIAL TUCUMAN"
matrixCodColor.Add "27", "DIGITAL GRIS"
matrixCodColor.Add "30", "STAR SEG"
matrixCodColor.Add "ABR", "ABROJO"
matrixCodColor.Add "ESTAMP", "ESTAMP"
matrixCodColor.Add "REF", "REFLECTIVO"

' Limpiar la hoja
Worksheets("Exportar TXT").Cells.Clear

' Separación de talles y colores ==========
' Datos fuentes
Worksheets("Depósito").Activate
Worksheets("Depósito").Range(Cells(2, 5), Cells(ultimaFila, 5)).Select
Set RangoVariante = Worksheets("Depósito").Range(Cells(2, 5), Cells(ultimaFila, 5))
Selection.Copy
Worksheets("Exportar TXT").Range("C1").PasteSpecial xlPasteValues
Application.CutCopyMode = False
Worksheets("Exportar TXT").Activate

' Separar en columnas
' Comprobar si hay datos en el rango "Variante" antes de procesar
If Application.WorksheetFunction.CountA(RangoVariante) > 0 Then
    Sheets("Exportar TXT").Range(Cells(1, 3), Cells(ultimaFila + 1, 3)).TextToColumns _
        Destination:=Range(Cells(1, 3), Cells(ultimaFila + 1, 3)), _
        DataType:=xlDelimited, _
        ConsecutiveDelimiter:=True, _
        Tab:=False, _
        Semicolon:=False, _
        Comma:=True
End If


Range("A1").Activate

' Acomodar los datos del 1° el color y 2° el talle
For Fila = 1 To ultimaFila
    Cells(Fila, 3).Select
    ' Recorre el diccionario buscando coincidencia
    For Each item In matrixCodColor
        If matrixCodColor(item) = Cells(Fila, 3).Value Then
            Cells(Fila, 3).Value = "'" & item
            GoTo proximaFila
        ElseIf Cells(Fila, 3).Value = "" Then
            GoTo proximaFila
        End If
    Next item

' Traslandando el talle a la siguiente columna
Cells(Fila, 4).Value = Cells(Fila, 3).Value
Cells(Fila, 3).Value = ""
proximaFila:
Next Fila

' Borrar espacios en blanco
Range(Cells(1, 4), Cells(ultimaFila + 1, 4)).Replace what:=" ", Replacement:="", LookAt:=xlPart, searchorder:= _
        xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False


' Completa planilla para exportar
For Fila = 1 To ultimaFila - 1
    
    ' 1º) Stock
    Cells(Fila, 1).Value = "'" & Worksheets("Depósito").Cells(Fila + 1, 6).Value
    
    ' 2º Codigo
    Cells(Fila, 2).Value = "'" & Worksheets("Depósito").Cells(Fila + 1, 3).Value
    
Next Fila

' Ajuste ultima Fila - HARDCODEO ESTO PARA PROBAR
ultimaFila = ultimaFila - 1

' Si se pasa del tope (30 líneas), serán "n" archivos con 30 líneas y otro con el resto
' de items que quedaron fuera. Sería el resto de una división, el módulo.
resto = ultimaFila Mod limite
cantArchivos = Int(ultimaFila / limite) + 1
Debug.Print "Archivos a importar: " & cantArchivos


' Generación del txt
Call generarTxt(Fila, ultimaFila, "", cantArchivos, nombreArchivo, carpetaDestino, limite, resto, server)

End Sub

Function generarTxt(Fila, ultimaFila, textoArchivo, cantArchivos, nombreArchivo, carpetaDestino, limite, resto, server)
Dim rutaArchivo As String
Dim i As Byte
Dim tope As Byte
Fila = 0


' Generación del txt
For i = 1 To cantArchivos
tope = i * limite
    If i = cantArchivos Then
        tope = ultimaFila
    End If
    
    For Fila = (limite * (i - 1)) + 1 To tope
        Cells(Fila, 1).Activate
        textoArchivo = textoArchivo _
            & Cells(Fila, 1).Value _
            & "+" & Cells(Fila, 2).Value _
            & "!" & Cells(Fila, 3).Value _
            & "!" & Cells(Fila, 4).Value _
            & vbNewLine
            Debug.Print "Archivo N°: " & i, "Fila N° :" & Fila
    Next Fila
    
    ' Si es mayor a uno, se van nombrando incrementalmente
    If cantArchivos > 1 Then
        nombreArchivo = Left(ActiveWorkbook.Name, Len(ActiveWorkbook.Name) - 5) & " - " & i & ".txt"
    Else
        nombreArchivo = Left(ActiveWorkbook.Name, Len(ActiveWorkbook.Name) - 5) & ".txt"
    End If
    
    rutaArchivo = server & carpetaDestino & nombreArchivo
    Debug.Print textoArchivo
    
   
    ' Lo comenté porque generaba un error. No debería.
    Open rutaArchivo For Output As #1
    Print #1, textoArchivo
    Close #1
    
    MsgBox "Datos exportados con éxito a " & rutaArchivo, vbInformation, "Cargar detalle desde txt"
Next i



End Function


