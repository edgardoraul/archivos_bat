Attribute VB_Name = "VentasWEB"
Option Explicit


Function correo(numVenta, nombre, ultima, i, packar, planilla)
    ' GENERA UN LISTADO DE VENTAS Y N° GUIAS PARA EL CORREO
    ' Acumulador negativo para evitar filas en blanco
    
    Dim vacia As Byte
    vacia = 0
    ' Acumulador para el caso de las segunda o X tanda
    Dim tanda As Byte
    
    ' Borrando el contenido viejo
    packar.Sheets(1).Range("A9:C39").ClearContents
    
    ' Completando la información
    For i = 2 To ultima
        ' Asignando el valor a cada N° vta.
        numVenta = planilla.Sheets("ventas").Cells(i, 1).Value
        nombre = planilla.Sheets("ventas").Cells(i, 2).Value
        Debug.Print numVenta & "-" & nombre
        
        ' Controlando espacios vacíos
        If numVenta = "" Or planilla.Sheets("ventas").Cells(i, 9).Value = "Retira en Local" Then
            vacia = vacia + 1
        End If
        
        ' Recorremos la planilla del Correo
        ' Controlamos que el número de venta esté completo
        ' y además que NO SEA un retiro en Local
        If numVenta <> "" And planilla.Sheets("ventas").Cells(i, 9).Value <> "Retira en Local" Then
            packar.Sheets(1).Cells(i + 7 - vacia, 1).Value = numVenta
            packar.Sheets(1).Cells(i + 7 - vacia, 2).Value = nombre
        End If
    Next i
    
    
End Function

Function generarRoutuloRetiro(nombre, telefono, dni, fecha, numVenta, ruta)
    ' GENERA PESTAÑAS CON ROTULOS PARA RETIRO EN LOCAL
    ' Enmarcando
    ActiveSheet.Range("A1:H21").Select
    With Selection
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With
    
    ' Formato de impresión
    With ActiveSheet.PageSetup
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
        .PrintArea = ActiveSheet.Range("A1:H21")
        .Zoom = False
        .FitToPagesTall = 1
        .FitToPagesWide = 1
    End With
    
   
    ' Dando un alto a la fila
    ActiveSheet.Range("A2:A2").RowHeight = 30
    
     ' Insertando la imagen
    ActiveSheet.Pictures.Insert(ruta & "..\logo.png").Select
    
    ' Centrando el logo
    With Selection
        .Top = 4
        .Left = 155
    End With
    
   
    ' Leyenda de retiro
    Range("A4:H6").Select
    With Selection
        .Merge
        .Font.Size = 30
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    Range("A4").Value = "RETIRA EN ENTREPISO"
    
    ' Nombre del cliente, en mayúsculas
    Range("A9:H10").Select
    With Selection
        .Merge
        .Font.Size = 25
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Interior.color = RGB(220, 220, 220)
    End With
    Range("A9").Value = UCase(nombre)
    
    ' dni/cuit del cliente
    Range("a12").Value = "DNI/CUIT:"
    Range("a12").HorizontalAlignment = xlRight
    Range("a12:b12").Font.Bold = True
    Range("a12:b12").Font.Size = 13
    Range("b12").Value = "'" & dni
    
    ' Teléfono del cliente
    Range("a14:d14").Select
    With Selection
        .Font.Bold = True
    End With
    Range("a14").Value = "TELEFONO:"
    Range("a14").HorizontalAlignment = xlRight
    Range("b14").HorizontalAlignment = xlLeft
    Range("b14").Value = telefono
    
    ' FIRMA
    Range("A20").Select
    With Selection
        .Value = "FIRMA:"
        .HorizontalAlignment = xlRight
        .Font.Bold = True
    End With
    Range("b20:d20").Borders(xlEdgeBottom).LineStyle = xlContinuous
    
    ' FECHA
    Range("f20").Select
    With Selection
        .Value = "FECHA RETIRO:"
        .HorizontalAlignment = xlRight
        .Font.Bold = True
    End With
    Range("g20:h20").Borders(xlEdgeBottom).LineStyle = xlContinuous
    
    ' FECHA ELABORACIÓN
    Range("f12").Value = "FECHA DE ELABORACIÓN:"
    Range("f12").HorizontalAlignment = xlRight
    Range("f12:h12").Font.Bold = True
    Range("g12").Value = fecha
    Range("g12").HorizontalAlignment = xlLeft
    
    ' NUMERO DE VENTA
    Range("f14").Select
    With Selection
        .Value = "N° DE VENTA WEB:"
        .Font.Bold = True
        .HorizontalAlignment = xlRight
    End With
    
    Range("g14").Select
    With Selection
        .Value = numVenta
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
Range("D:K").Columns.AutoFit

' Centrando el contenido
Range("E:E").HorizontalAlignment = xlCenter
Cells(ultima + 1, 5).HorizontalAlignment = xlRight

' Acomoda el texto de las celdas con datos
Range("B:B").ColumnWidth = 40
Range("C:C").ColumnWidth = 50
Range("A:A").ColumnWidth = 7
Range("E:E").ColumnWidth = 12
'Range(Cells(2, 1), Cells(ultima, 11)).WrapText = True

' Ajusta automáticamente la altura de las filas
Range(Cells(2, 1), Cells(ultima, 11)).Rows.AutoFit

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
    .PrintArea = ActiveSheet.Range("A1:I" & ultima + 1).Address
    .Zoom = False
    .FitToPagesTall = 1
    .FitToPagesWide = 1
    .CenterHeader = "&B&20&F"
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
Cells(ultima + 1, 6).Borders.LineStyle = xlContinuous

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


End Function

Sub GuardarArchivo(fecha, ruta, nombreArchivo)
' VALIDANDO NOMBRE DE ARCHIVO A GENERAR Y GUARDAR

' Variables a utilizar
Dim nombre As String
Dim cuenta As String

' Asignando algunos valores
ruta = ruta & "WEB\"


'Controlando si la compu EDGARD está prendida y conectada a red.
If Dir(ruta, vbDirectory) = "" Then
    MsgBox ("No hay acceso la compu EDGARD. Debes prender esa compu y que se conecte a la red.")
    Exit Sub
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
nombreArchivo = "Ventas Web " & parteNumero & ". " & fecha & ".xlsx"
nombre = ruta & "Ventas Web " & parteNumero & ". " & fecha & ".xlsx"

Sheets("ventas").Range("A1").Select
ActiveWorkbook.SaveAs Filename:=nombre, FileFormat:=xlOpenXMLStrictWorkbook, ConflictResolution:=xlUserResolution, AddToMru:=True, Local:=True
ActiveWorkbook.Save
Application.ThisWorkbook.Save


End Sub


Sub ventasWeb()

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
Dim packar As Object
Dim ultima As Integer
Dim fecha As String
Dim i As Integer
Dim rotulos As Integer
Dim ruta As String
Dim img As String
Dim nombreArchivo As String

' Sirve para averiguar el nombre de la computadora actual
Dim ws As Object
Set ws = CreateObject("WScript.network")

' Asignando algunos valores de acuerdo en qué equipo de la red esté
If ws.ComputerName = "EDGARD" Then
    ruta = "D:\Web\Listados de Ventas Online\"
    Debug.Print "Estoy en la computadora: " & ws.ComputerName
Else
    ruta = "\\EDGARD\Web\Listados de Ventas Online\"
    Debug.Print "Estoy en una computadora de la red, llamada: " & ws.ComputerName
End If
Debug.Print "Se guardan los archivos en: " & ruta

rotulos = 0
fecha = Day(Date) & "-" & Month(Date) & "-" & Year(Date)


' Guardando el archivo con nombre específico
Call GuardarArchivo(fecha, ruta, nombreArchivo)
Range("A1").Activate
ultima = ActiveSheet.Cells(Rows.count, 1).End(xlUp).Row

' Borrando información innecesaria
Range("Y:Y").EntireColumn.Copy
Range("AO:AO").EntireColumn.PasteSpecial
Range("B:K, P:T, X:X, Z:AE, AG:AG, AI:AN").EntireColumn.Delete
Range("C:E").EntireColumn.Insert
Range("H:H").EntireColumn.Copy
Range("C:C").PasteSpecial
Range("A1").Value = "Núm. Venta"
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
Range("L:L").EntireColumn.Delete


' Purgando los teléfonos
For i = 2 To ultima
    ' Columna de los teléfonos
    Cells(i, 8).Value = Right(Cells(i, 8).Value, 10)
Next i

' Insertando columna para la ubicación
Range("I:I").EntireColumn.Insert
Range("I1:I1").Value = "Detalles"

' Dando formato
For i = 2 To ultima
    ' Columna de las ubicaciones de productos
    With Cells(i, 9)
        .WrapText = False
        .ShrinkToFit = True
    End With
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
        ActiveCell.Offset(0, 13) = ""
    End If
    ActiveCell.Offset(1, 0).Activate
Loop

' DANDO FORMATO A TODA LA PLANILLA
Call formato(ultima, i)


' GENERANDO LOS ROTULOS DE RETIRO
For i = 2 To ultima
    If Sheets("ventas").Cells(i, 13).Value = "Retiras en Rerda Mendoza" Then
        
        ' Se coloca la leyenda en la celda
        Debug.Print Cells(i, 9).Value
        Cells(i, 9).Value = "Retira en Local"
        
        ' Variables a completar
        nombre = Cells(i, 2).Value
        telefono = Cells(i, 8).Value
        dni = Cells(i, 7).Value
        numVenta = Cells(i, 1).Value
        
        ' Contador de rótulos a imprimir
        rotulos = rotulos + 1
        
        ' Se agrega una pestaña para el respectivo rótulo
        ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Sheets("ventas")).Name = "Venta N° " & Cells(i, 1).Value
        
        ' Se genera el rótulo respectivo
        Call generarRoutuloRetiro(nombre, telefono, dni, fecha, numVenta, ruta)
        
    End If
    Sheets("ventas").Activate
Next i

'Posicionando al principio
Sheets.Add(After:=Sheets("ventas")).Name = "Depósito"
Sheets.Add(After:=Sheets("Depósito")).Name = "Exportar TXT"
Sheets("ventas").Activate
Range("A1").Activate

' Definiendo este archivo
Set planilla = ActiveWorkbook

' DANDO FORMATO DE IMPRESION
Call formatPrint(ultima, i)



' Dando un aviso condicional sólo si hay rótulos, se lo contrario, no.
If rotulos > 0 Then
    MsgBox ("Tenés " & rotulos & " rótulos de retiro en local, para imprimir." & VBA.vbNewLine & "Aquí abajo en las pestañas.")
End If


' Generando una planilla sólo para dpto. DEPOSITO
Call deposito


' Abrir el archivo
ruta = ruta & "..\"
Workbooks.Open ruta & "ENCOMIENDAS_WEB.xlsx"
Set packar = ActiveWorkbook

' Generando una planilla de informe para el Correo Argentino
Call correo(numVenta, nombre, ultima, i, packar, planilla)


' Posicionando al principio
planilla.Sheets("ventas").Activate
ActiveWorkbook.Save
End Sub

Sub deposito()
Attribute deposito.VB_Description = "Genera planilla de depósito para la Web."
Attribute deposito.VB_ProcData.VB_Invoke_Func = "S\n14"
' GENERA UNA PLANILLA SÓLO PARA USO EXCLUSIVO DEL DEPOSITO
Dim ruta As String
Dim nombreArchivo As String
Dim ultima As Byte
Dim i As Byte
Dim enrutacion As String
Dim web As Boolean
enrutacion = ActiveWorkbook.Path & "\..\"
Debug.Print enrutacion

ruta = "'" & enrutacion & "[Stock.XLS]Sheet1'!$A$2:$G$10000"
ultima = Sheets("ventas").Cells(Rows.count, 2).End(xlUp).Row - 1
web = False


' Control si existe la pestaña "Depósito"
Call CrearHoja("Depósito")


' Creando las columnas
Sheets("Depósito").Activate
Sheets("Depósito").Cells.Clear
Cells(1, 1).Value = "Nº Venta"
Cells(1, 2).Value = "Cliente"
Cells(1, 3).Value = "Descripción"
Cells(1, 4).Value = "Código"
Cells(1, 5).Value = "Variante"
Cells(1, 6).Value = "Cantidad"
Cells(1, 7).Value = "Ubicación"

' Completando los datos
For i = 2 To ultima
    ' Venta
    Cells(i, 1).Value = Sheets(1).Cells(i, 1).Value
    
    ' Cliente
    Cells(i, 2).Value = Sheets(1).Cells(i, 2).Value
    
    ' Descripción
    Cells(i, 3).Value = "=VLOOKUP(D" & i & "," & ruta & ",2,FALSE)"
    
    ' Código
    Cells(i, 4).Value = Sheets(1).Cells(i, 4).Value
    
    ' Variante
    Cells(i, 5).Value = Sheets(1).Cells(i, 5).Value
    
    ' Cantidad
    Cells(i, 6).Value = Sheets(1).Cells(i, 6).Value
        
    ' La ubiación
    Cells(i, 7).Formula = "=VLOOKUP(D" & i & "," & ruta & ",7,FALSE)"
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
    .Orientation = xlLandscape
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
Call exportarTxt
Sheets("ventas").Activate
Range("A1").Activate

End Sub

Function CrearHoja(nombreHoja As String) As Boolean
    ' controla si una hoja existe o no
    Dim existe As Boolean
     
    On Error Resume Next
    existe = (Worksheets(nombreHoja).Name <> "")
     
    If Not existe Then
        Worksheets.Add(After:=Worksheets(Worksheets.count)).Name = nombreHoja
    End If
     
    CrearHoja = existe
     
End Function


Sub exportarTxt()

' GENERA UN ARCHIVO DE TEXTO PARA IMPORTAR AL D.F.

Dim fila As Long, columna As Long
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

Dim matrixCodColor As Object
ultimaFila = Sheets("Depósito").Cells(Rows.count, 2).End(xlUp).Row - 1
nombreArchivo = Len(ActiveWorkbook.Name)
server = "\\SER-DF\A Remitar TXT"
carpetaDestino = "\WEB\"
Set matrixCodColor = CreateObject("Scripting.Dictionary")
limite = 30

Sheets("Exportar TXT").Activate

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

' Recorre la colección
'For Each item In matrixCodColor
    'Debug.Print item & ": " & matrixCodColor(item)
'Next item

' Limpiar la hoja
Sheets("Exportar TXT").Cells.Clear

' Separación de talles y colores ==========
' Datos fuentes
Sheets("Depósito").Activate
Sheets("Depósito").Range(Cells(2, 5), Cells(ultimaFila, 5)).Select
Selection.Copy
Sheets("Exportar TXT").Range("C1").PasteSpecial xlPasteValues
Application.CutCopyMode = False
Sheets("Exportar TXT").Activate

' Separar en columnas
Sheets("Exportar TXT").Range(Cells(1, 3), Cells(ultimaFila + 1, 3)).TextToColumns _
    Destination:=Range(Cells(1, 3), Cells(ultimaFila + 1, 3)), _
    DataType:=xlDelimited, _
    ConsecutiveDelimiter:=True, _
    Tab:=False, _
    Semicolon:=False, _
    Comma:=True

Range("a1").Activate

' Acomodar los datos del 1° el color y 2° el talle
For fila = 1 To ultimaFila
    Cells(fila, 3).Select
    ' Recorre el diccionario buscando coincidencia
    For Each item In matrixCodColor
        If matrixCodColor(item) = Cells(fila, 3).Value Then
            Cells(fila, 3).Value = "'" & item
            GoTo proximaFila
        ElseIf Cells(fila, 3).Value = "" Then
            GoTo proximaFila
        End If
    Next item

' Traslandando el talle a la siguiente columna
Cells(fila, 4).Value = Cells(fila, 3).Value
Cells(fila, 3).Value = ""
proximaFila:
Next fila

' Borrar espacios en blanco
Range(Cells(1, 4), Cells(ultimaFila + 1, 4)).Replace what:=" ", Replacement:="", LookAt:=xlPart, searchorder:= _
        xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False


' Completa planilla para exportar
For fila = 1 To ultimaFila - 1
    
    ' 1º) Stock
    Cells(fila, 1).Value = "'" & Sheets("Depósito").Cells(fila + 1, 6).Value
    
    ' 2º Codigo
    Cells(fila, 2).Value = "'" & Sheets("Depósito").Cells(fila + 1, 4).Value
    
Next fila

' Ajuste ultima Fila - HARDCODEO ESTO PARA PROBAR
ultimaFila = ultimaFila - 1

' Si se pasa del tope (30 líneas), serán "n" archivos con 30 líneas y otro con el resto
' de items que quedaron fuera. Sería el resto de una división, el módulo.
resto = ultimaFila Mod limite
cantArchivos = Int(ultimaFila / limite) + 1
Debug.Print "Archivos a importar: " & cantArchivos


' Generación del txt
Call generarTxt(fila, ultimaFila, "", cantArchivos, nombreArchivo, carpetaDestino, limite, resto, server)

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
        nombreArchivo = Left(ActiveWorkbook.Name, Len(ActiveWorkbook.Name) - 5) & " - " & i & ".txt"
    Else
        nombreArchivo = Left(ActiveWorkbook.Name, Len(ActiveWorkbook.Name) - 5) & ".txt"
    End If
    
    rutaArchivo = server & carpetaDestino & nombreArchivo
    Debug.Print textoArchivo
    Open rutaArchivo For Output As #1
    Print #1, textoArchivo
    Close #1
    
    MsgBox "Datos exportados con éxito a " & rutaArchivo, vbInformation, "Cargar detalle desde txt"
Next i



End Function
