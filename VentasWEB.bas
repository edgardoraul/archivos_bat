Attribute VB_Name = "VentasWEB"
Option Explicit
Sub GuardarArchivo()
' VALIDANDO NOMBRE DE ARCHIVO A GENERAR
' Variables a utilizar
Dim ruta As String
Dim nombre As String
Dim cuenta As String
Dim fecha As String

' Asignando algunos valores
'ruta = "\\EDGARD\Web\Listados de Ventas Online\WEB\"
ruta = "D:\Web\Listados de Ventas Online\WEB\"
fecha = Day(Date) & "-" & Month(Date) & "-" & Year(Date)

'Controlando si la carpeta existe, de lo contrario, crearla en local
If Dir(ruta, vbDirectory) = "" Then
    MkDir (ruta)
    MsgBox ("No hay acceso la compu EDGARD. Debes prender esa compu y que se conecte a la red.")
Else
    
End If

' Definiendo unas variables
Dim archivos As String
Dim u As Integer
Dim denominacion As String
    
' Preparación de variables
u = 1
archivos = Dir(ruta)
    
' Recorrido de la carpeta
ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook _
    .ActiveSheet).Name = "Listado"
Sheets("Listado").Visible = False
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
    Exit Sub
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
nombre = ruta & "Ventas Web " & parteNumero & ". " & fecha & ".xlsx"

Sheets("ventas").Range("A1").Select
ActiveWorkbook.SaveAs Filename:=nombre, FileFormat:=xlOpenXMLStrictWorkbook, ConflictResolution:=xlUserResolution, AddToMru:=True, Local:=True
ActiveWorkbook.Save
End Sub
Function formatPrint(ultimaFila, i)
' Dando formato apaisado, expandido a A4 y con titulares. Una sóla página.

' Delimitando el tamaño de hojas y márgenes
Dim filasTotales As Integer
filasTotales = ultimaFila + 1

' Formatea la última columna que NO saldrá impresa, sólo para acomodar, nada más
Range("D:K").Columns.AutoFit

' Centrando el contenido
Range("E:E").HorizontalAlignment = xlCenter

' Acomoda el texto de las celdas con datos
Range("B:B").ColumnWidth = 40
Range("C:C").ColumnWidth = 50
Range(Cells(2, 1), Cells(ultimaFila, 11)).WrapText = True

' Ajusta automáticamente la altura de las filas
Range(Cells(2, 1), Cells(ultimaFila, 10)).Rows.AutoFit

' Formato de impresión
With ActiveSheet.PageSetup
    .Orientation = xlLandscape
    .PaperSize = xlPaperA4
    .LeftMargin = Application.CentimetersToPoints(0.64)
    .RightMargin = Application.CentimetersToPoints(0.64)
    .TopMargin = Application.CentimetersToPoints(2.5)
    .BottomMargin = Application.CentimetersToPoints(1.91)
    .HeaderMargin = Application.CentimetersToPoints(0.76)
    .FooterMargin = Application.CentimetersToPoints(0.76)
    .CenterHorizontally = True
    .CenterVertically = False
    .PrintArea = ActiveSheet.Range("A1:H" & filasTotales).Address
    .Zoom = False
    .FitToPagesTall = 1
    .FitToPagesWide = 1
    .CenterHeader = "&B&20&F"
End With
    
End Function
Function formato(ultimaFila, i)
' Enmarcando toda la tabla
Range("A1").CurrentRegion.Select
With Selection
    .Borders(xlEdgeLeft).LineStyle = xlContinuous
    .Borders(xlEdgeRight).LineStyle = xlContinuous
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
Range("A1:K1").Select
Selection.Borders.LineStyle = xlContinuous

' Se cuentan cuantas celdas ocupadas hasta el final
Range(Cells(ultimaFila, 1), Cells(ultimaFila, 11)).Select
    With Selection
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With

' Colocando totales de productos y dando formato
Cells(ultimaFila + 1, 5).Value = "TOTALES:"
Cells(ultimaFila + 1, 6).Select
Cells(ultimaFila + 1, 6).Value = "=SUM(F2:F" & ultimaFila & ")"
Range(Cells(ultimaFila + 1, 5), Cells(ultimaFila + 1, 6)).Select
    With Selection
        .Font.Bold = True
        .Font.Size = 15
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlBottom
    End With
Cells(ultimaFila + 1, 6).Borders.LineStyle = xlContinuous

' Colocando el total de rótulos a imprimir
Cells(ultimaFila + 1, 2).Value = "ROTULOS:"
Cells(ultimaFila + 1, 3).Value = "=COUNTA(A2:A" & ultimaFila & ")"
Range(Cells(ultimaFila + 1, 2), Cells(ultimaFila + 1, 3)).Select
With Selection
    .Font.Bold = True
    .Font.Size = 15
    .VerticalAlignment = xlBottom
    .HorizontalAlignment = xlRight
End With
Cells(ultimaFila + 1, 3).HorizontalAlignment = xlLeft

' Colocando un borde superior
For i = 3 To ultimaFila
    Range(Cells(i, 1), Cells(i, 11)).Select
    If Cells(i, 1).Value <> "" Then
        With Selection
            .Borders(xlEdgeTop).LineStyle = xlContinuous
        End With
    End If
Next i

End Function
Function guardando(savename)
    ' Guarda y renombra el archivo con número y fecha
    savename = Application.GetSaveAsFilename(fileFilter:="Exel Files (*.xlsx), *.xlsx")
    ActiveWorkbook.SaveAs Filename:=savename, FileFormat:=51
End Function
Sub ventasWeb()
Dim savename As String
Dim ultimaFila As Integer
Dim i As Integer
'Call guardando(savename)
Call GuardarArchivo
Range("A1").Activate
ultimaFila = ActiveSheet.Cells(Rows.count, 1).End(xlUp).Row

' Borrando información innecesaria
Range("B:K, P:T, X:X, Z:AE, AG:AG, AI:AN").EntireColumn.Delete
Range("C:E").EntireColumn.Insert
Range("H:H").EntireColumn.Copy
Range("C:C").PasteSpecial
Range("A1").Value = "Núm. Venta"
Range("F:F").Select
Selection.NumberFormat = "General"
Range("G:G").Select
Selection.NumberFormat = "0"

'Recorremos los nombres de los clientes
For i = 2 To ultimaFila
    
    If Cells(i, 2).Value = Cells(i, 3).Value Then
    Else
        Cells(i, 2).Value = Cells(i, 2).Value & " - " & Cells(i, 3).Value
    End If
    
Next i
Range("B1").Value = "Cliente"

' Elimino la columna innecesaria
Range("C:C").EntireColumn.Delete
Range("G:G").EntireColumn.Delete
Range("C:D").EntireColumn.Insert

'Moviendo el detalle
Range("M:M").EntireColumn.Copy
Range("C:C").PasteSpecial
Range("M:M").EntireColumn.Delete
Range("C1").Value = "Descripción"

'Moviendo la cantidad
Range("M:M").EntireColumn.Copy
Range("F:F").PasteSpecial
Range("L:L").EntireColumn.Delete
Range("F1").Value = "Cantidad"

'Eliminando columna innecesaria
Range("L:L").EntireColumn.Delete

'Purgando los teléfonos
For i = 2 To ultimaFila
    Cells(i, 8).Value = Right(Cells(i, 8).Value, 10)
Next i

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


'Borra cosas innecesarias
Do While ActiveCell.Value <> ""
    If ActiveCell.Offset(0, 1) = "" Then
        ActiveCell.Value = ""
    End If
    ActiveCell.Offset(1, 0).Activate
Loop

' DANDO FORMATO A TODA LA PLANILLA
Call formato(ultimaFila, i)

' DANDO FORMATO DE IMPRESION
Call formatPrint(ultimaFila, i)

' GENERANDO LOS ROTULOS DE RETIRO
For i = 2 To ultimaFila
    If Cells(i, 10).Value = "Correo Argentino - Envio a domicilio" Then
    ' Imprimir rótulo
    End If
Next i

'Posicionando al principio
Sheets("ventas").Range("A1").Activate
End Sub
