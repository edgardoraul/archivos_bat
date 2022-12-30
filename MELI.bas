Attribute VB_Name = "MELI"
Option Explicit

Sub AA_MELI()
Attribute AA_MELI.VB_Description = "Crea las planillas para MercadoLibre."
Attribute AA_MELI.VB_ProcData.VB_Invoke_Func = "K\n14"
' ============================================================
' GENERA EN FORMA AUTOMATIZADA LAS PLANILLAS DE VENTAS DE MELI
' ============================================================

' ============================================================
' Controlar que no se haya hecho formato antes
If Range("J1").Value = "Firma Control" Then
    MsgBox ("Ya le diste formato a esta planilla. Probá con otra.")
    Range("A1").Select
    Exit Sub
End If

' ============================================================
' Creando las hojas nuevas
Cells.Select
Cells.ClearFormats
ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook _
    .Worksheets(ActiveWorkbook.Worksheets.count)).Name = "Planilla"

Application.Worksheets(1).Select
'Debug.Print Application.Worksheets.count



' Copiando información importante a la Planilla
Range("A:A, C:C, H:H, J:J, L:L, N:N, O:O, P:P, AW:AW, AU:AU").Select
Selection.Copy
Sheets("Planilla").Paste
Application.CutCopyMode = False

' Borrando la hoja innecesaria
Worksheets(1).Delete

' Acomodando las columnas
Columns(3).EntireColumn.Insert
Range("D:D").Copy
Range("L:L").PasteSpecial xlPasteAll

Range("K:K").Copy
Range("C:C").PasteSpecial xlPasteAll
Application.CutCopyMode = False

'Range("J:J").Copy
'Range("D:D").PasteSpecial xlPasteAll
'Application.CutCopyMode = False

' Eliminando las columnas sobrantes
Columns(4).EntireColumn.Delete
Columns(10).EntireColumn.Delete

' Insertando unas columnas necesarias
'Columns(9).EntireColumn.Insert
Columns(10).EntireColumn.Insert

' Colocando títulos
Range("B1").Value = "Nº de Venta"
Range("C1").Value = "Cliente"
Range("D1").Value = "Descripción"
Range("E1").Value = "Código"
Range("I1").Value = "Detalles"
Range("J1").Value = "Firma Control"


' ============================================================
' BORRANDO TEXTOS INNECESARIOS

' Definiendo la variable como arreglo
Dim cadenaOriginal As Variant

' Listado de expresiones a borrar. Deber tener al último el "" para quelo tome el bucle.
cadenaOriginal = Array("-CL-EG", "-PR-EG", "-PR", "-CL", " T:52-56", " T:2XS-XL", "T:34-44", "T:34-48", "T:36-48", "T:46-48", " T:46-50", "T:50-54", "T:56-60", "T:62-66", "T:34/44", "T:34/48", "T:36/48", "T:46/48", "T:38-48", "T:50/54", " 50-54", "T:56/60", "T:62/66", "T:XXS-XXL", "T:2XS-2XL", "2XS/2XL", "XXS/XXL", "T:3XL-5XL", "T:2XS/2XL", "3XL/5XL", " envío gratis", " envio gratis", " rerda", " envío gratis", " envio gratis", " en cuotas", "en cuotas", " premium", "premium", "Unico", "Único", "Regulable", " cuotas", " talles especiales", "talle especial", " - ", " . ", "   ", "...", "..", "")

' Definiendo el largo del array con un bucle While
Dim largo As Integer
largo = 0
Do While largo >= 0
    If cadenaOriginal(largo) = "" Then Exit Do
    largo = largo + 1
Loop

' Definiendo el contador
Dim i As Integer

' Bucle. Debe coincidir el largo del arreglo con el fin del bucle
For i = 0 To largo
    Cells.Replace what:=cadenaOriginal(i), Replacement:="", LookAt:=xlPart, _
        searchorder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
Next

' ============================================================
' FORMATEANDO LA TABLA PARA DEJARLA BONITA

' Formateando las columnas
Columns("A").ColumnWidth = 9.5
Columns("B").ColumnWidth = 16.5
Columns("C").ColumnWidth = 23.71
Columns("D").ColumnWidth = 40.14
Columns("E").ColumnWidth = 10.29
Columns("F").ColumnWidth = 10.86
Columns("G").ColumnWidth = 6
Columns("H").ColumnWidth = 9.3
Columns("I").ColumnWidth = 10.7
Columns("J").ColumnWidth = 10.7

' Formateando los encabezados
Rows("1").RowHeight = 25.5
Range("A1").Select
Range(Selection, Selection.End(xlToRight)).Select
    With Selection
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.color = RGB(250, 250, 250)
        .WrapText = True
    End With
    
' Formateando la columna de fechas
Range("A1").Select
Range(Selection, Selection.End(xlDown)).Select
With Selection
    .NumberFormat = "m/d/yyyy"
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

' Agregando bordes HORIZONTALES al final de tabla
Range("K1").Select
' Se cuentan cuantas celdas ocupadas hasta el final
Dim ultima As Integer
ultima = Cells(Rows.count, 1).End(xlUp).Row
Range(Cells(ultima, 1), Cells(ultima, 11)).Select
    With Selection
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With

' Colocando totales de productos y dando formato
Cells(ultima + 1, 7).Value = "TOTALES:"
Cells(ultima + 1, 8).Select
Cells(ultima + 1, 8).Value = "=SUM(H2:H" & ultima & ")"
Range(Cells(ultima + 1, 7), Cells(ultima + 1, 8)).Select
    With Selection
        .Font.Bold = True
        .Font.Size = 15
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlBottom
    End With


' Dando formato a la columna de los números de ventas
Range(Cells(2, 2), Cells(ultima, 2)).Select
With Selection
    .NumberFormat = "###"
    .HorizontalAlignment = xlRight
End With

' Filtrando la cantidad de rótulos mediante bucle
' Parándose en la última celda de la columna
Dim contador As Integer
contador = 0

' Con el bucle iremos subiendo y controlando si la celda superior es igual a la que estamos parada,
' de ser así, subimos a la que está arriba y borramos el valor de la que está abajo y otras correspondientes a las fechas número de venta y nombre de cliente.
Do While ultima - contador > 1
    Cells(ultima - contador, 11).Select
    
    ' Si no tiene nada es porque hay un retiro en Local
    If Cells(ultima - contador, 11).Value = "" Then
        Cells(ultima - contador, 11).Value = "Retira en Local"
    
    ' Controlando valores iguales
    ElseIf Cells(ultima - contador, 11).Value = Cells(ultima - contador - 1, 11).Value Then
        Cells(ultima - contador - 1, 11).Select
        Cells(ultima - contador, 11).ClearContents
        Range(Cells(ultima - contador, 1), Cells(ultima - contador, 3)).ClearContents
        
    End If
    
    ' Aprovechando para colocar un borde superior
    If Cells(ultima - contador, 11).Value <> "" Then
        Range(Cells(ultima - contador, 1), Cells(ultima - contador, 11)).Select
        With Selection
            .Borders(xlEdgeTop).LineStyle = xlContinuous
        End With
    End If
    contador = contador + 1
Loop


' Colocando totales de Rótulos y dando formato. Pero antes deben estar filtrados cuántos son realmente
Cells(ultima + 1, 3).Value = " ROTULOS"
Cells(ultima + 1, 2).Select
Cells(ultima + 1, 2).Value = "=COUNTA(K2:K" & ultima & ")-COUNTIF(K2:K" & ultima & ", ""Retira en Local"")"
Range(Cells(ultima + 1, 2), Cells(ultima + 1, 3)).Select
    With Selection
        .Font.Bold = True
        .Font.Size = 15
        .VerticalAlignment = xlBottom
    End With

' Formato para acomodar el texto en toda la tabla imprimible
Range(Cells(1, 1), Cells(ultima, "K")).Select
    With Selection
    End With


' ==========================================================
' FORMATO PARA IMPRIMIR UNA SOLA PÁGINA

' Delimitando el tamaño de hojas y márgenes
Dim filasTotales As Integer
filasTotales = ultima + 1

' Formatea la última columna que NO saldrá impresa, sólo para acomodar, nada más
Range("K:K").Columns.AutoFit

' Acomoda el texto de las celdas con datos
Range(Cells(2, 1), Cells(ultima, 11)).WrapText = True

' Ajusta automáticamente la altura de las filas
Range(Cells(2, 1), Cells(ultima, 10)).Rows.AutoFit

' Borde externo faltante
Range(Cells(2, 1), Cells(ultima, 11)).BorderAround LineStyle:=xlContinuous

' Formato de impresión
With Worksheets("Planilla").PageSetup
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
    .PrintArea = Sheets("Planilla").Range("A1:J" & filasTotales).Address
    .Zoom = False
    .FitToPagesTall = 1
    .FitToPagesWide = 1
    .CenterHeader = "&B&20&F"
End With


' ==========================================================
' VALIDANDO NOMBRE DE ARCHIVO A GENERAR

' Variables a utilizar
Dim ruta As String
Dim nombre As String
Dim cuenta As String
Dim fecha As String

' Asignando algunos valores
'ruta = "\\EDGARD\Web\Listados de Ventas Online\MELI"
ruta = "D:\Web\Listados de Ventas Online\MELI"
fecha = Day(Date) & "-" & Month(Date) & "-" & Year(Date)

'Controlando si la carpeta existe, de lo contrario, crearla en local
If Dir(ruta, vbDirectory) <> "" Then
    ruta = "D:\Web\Listados de Ventas Online\MELI"
    MkDir (ruta & "1")
    MkDir (ruta & "2")
    MsgBox ("No hay acceso la compu EDGARD. Se guardan en " & ruta & "1 y en " & ruta & "2")
End If

' Preguntando al usuario qué cuenta es
cuenta = Application.InputBox(Prompt:="¿Qué cuenta de MercadoLibre es? ¿1 ó 2?", Title:="Cuenta de MercadoLibre", Default:=1)
If cuenta <> 1 And cuenta <> 2 Then
    MsgBox ("¿Cuenta 1 ó 2?. Elegí bien.")
    cuenta = Application.InputBox(Prompt:="¿Qué cuenta de MercadoLibre es? ¿1 ó 2?", Title:="Cuenta de MercadoLibre", Default:=1)
End If
' En base a la respuesta, determinar la carpeta definitiva
ruta = ruta & cuenta & "\"


' Definiendo unas variables
Dim archivos As String
Dim u As Integer
Dim denominacion As String
    
' Preparación de variables
u = 1
archivos = Dir(ruta)
    
' Recorrido de la carpeta
ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook _
    .Worksheets("Planilla")).Name = "Listado"
Sheets("Listado").Visible = False
Sheets("Planilla").Select

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


' Controlando si es cuenta 1
If cuenta = 1 Then
    parteNumero = Mid(Sheets("Listado").Cells(u - 1, 1).Value, 8, 7)
    nombreNumero = CInt(parteNumero) + 1
    parteNumero = CStr(nombreNumero)
    
    ' Agregando ceros para tener un nombre coherente
    Do While Len(parteNumero) < 6
        parteNumero = "0" & parteNumero
        e = e + 1
    Loop
    nombre = ruta & "Pedidos " & parteNumero & ". " & fecha & ".xlsx"

ElseIf cuenta = 2 Then
    parteNumero = Mid(Sheets("Listado").Cells(u - 1, 1).Value, 19, 7)
    nombreNumero = CInt(parteNumero) + 1
    parteNumero = CStr(nombreNumero)
    
    ' Agregando ceros para tener un nombre coherente
    Do While Len(parteNumero) < 6
        parteNumero = "0" & parteNumero
        e = e + 1
    Loop
    nombre = ruta & "CUENTA 2 - Pedidos " & parteNumero & ". " & fecha & ".xlsx"
Else
    MsgBox ("Elegí: 1 ó 2")
End If

Sheets("Planilla").Range("A1").Select

ActiveWorkbook.SaveAs Filename:=nombre, FileFormat:=xlOpenXMLStrictWorkbook, ConflictResolution:=xlUserResolution, AddToMru:=True, Local:=True
ActiveWorkbook.Save

' Controlando si es la cuenta 2
If cuenta = 2 Then
    Call correo(ruta, i, ultima)
Else
    Exit Sub
End If

' ============== FIN DE TODA LA MACRO.
End Sub

Function correo(ruta, i, ultima)
' ======= LLAMADA A LA GENERACION DE PLANILLA PARA EL CORREO
' GENERA UN LISTADO DE VENTAS Y N° GUIAS PARA EL CORREO
    Dim numVenta As String
    Dim Cliente As String
    Dim tn As String
    Dim packar As Object
    Dim planilla As Object
    Dim hoy As Date
    Dim continuacion As Integer
    hoy = Date
    Set planilla = ActiveWorkbook
    
    ruta = ruta & "\..\ENCOMIENDAS_MELI2.xlsx"
    ' Abrir el archivo
    Workbooks.Open ruta
    Set packar = ActiveWorkbook
    
    ' CONTROL DE FECHAS
    If packar.Sheets(1).Range("A1").Value <> hoy Then
        packar.Sheets(1).Range("A1").Value = Date
        continuacion = 7
        
        ' Borrando el contenido viejo
        packar.Sheets(1).Range("A9:C39").ClearContents
    Else
        continuacion = packar.Sheets(1).Cells(39, 1).End(xlUp).Row - 1
        Debug.Print continuacion
    End If
     
    
    ' Completando la información
    For i = 2 To ultima
        ' Asignando el valor a cada N° vta.
        numVenta = planilla.Sheets(1).Cells(i, 2).Value
        Cliente = planilla.Sheets(1).Cells(i, 3).Value
        tn = planilla.Sheets(1).Cells(i, 11).Value
        
        ' Recorremos la planilla del Correo
        ' Controlamos que el número de venta esté completo
        ' y además que NO SEA un retiro en Local
        If numVenta <> "" And planilla.Sheets(1).Cells(i, 9).Value <> "Retira en Local" Then
            packar.Sheets(1).Cells(i + continuacion, 1).Value = numVenta
            packar.Sheets(1).Cells(i + continuacion, 2).Value = Cliente
            packar.Sheets(1).Cells(i + continuacion, 3).Value = tn
        End If
    Next i
End Function

