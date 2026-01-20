Attribute VB_Name = "MELI"
Option Explicit

Function PintarFila(Hojilla As String, Fila As Integer, DesdeColumna As Integer, HastaColumna As Integer)
    Dim TextoOriginal As String
    
    ' Pinta filas impares
    If Fila Mod 2 <> 0 Then
        Worksheets(Hojilla).Range(Cells(Fila, DesdeColumna), Cells(Fila, HastaColumna)).Interior.color = RGB(240, 240, 240)
        
        ' Formateando el dato del código
        Worksheets(Hojilla).Cells(Fila, 4).Font.color = RGB(240, 240, 240)
        Worksheets(Hojilla).Cells(Fila, 4).Characters(Start:=1, Length:=7).Font.ColorIndex = xlAutomatic

    End If
End Function

Sub AA_MELI()
Attribute AA_MELI.VB_Description = "Crea las planillas para MercadoLibre."
Attribute AA_MELI.VB_ProcData.VB_Invoke_Func = "K\n14"
' ============================================================
' GENERA EN FORMA AUTOMATIZADA LAS PLANILLAS DE VENTAS DE MELI
' ============================================================
Dim ultima As Integer
Dim i As Integer

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
Cells.Font.Size = 11
ActiveWorkbook.Sheets.Add(after:=ActiveWorkbook _
    .Worksheets(ActiveWorkbook.Worksheets.Count)).Name = "Planilla"
Application.Worksheets(1).Select


' Dando valor a la última fila
ultima = Cells(Rows.Count, 1).End(xlUp).Row

' Borrando una columna del barrio
Range("AH:AH").Delete

' Corrigiendo información
Cells(1, 49).Activate

' Bucle que recorre las filas y obtiene el nombre+apellido y/o razón social.
For i = 2 To ultima
    Debug.Print Cells(i, 49).Value
    Cells(i, 49).Activate
    
    ' Obtención del usuario en mayúsculas
    Cells(i, 49).Value = Cells(i, 4).Value
Next i


' Copiando información importante a la Planilla
Range("A:A, C:C, H:H, J:J, L:L, N:N, O:O, P:P, AW:AW, AU:AU").Select
Selection.Copy
Sheets("Planilla").Paste
Application.CutCopyMode = False

' Copiando unas ciertas columnas
Range("AS:AT").Select
Selection.Copy
Sheets("Planilla").Activate
Range("M1").Activate
Range("M1").PasteSpecial xlPasteAll
Application.CutCopyMode = False


' Reacomodando los datos
For i = 2 To ultima
    Cells(i, 13).Value = Cells(i, 13).Value & " " & Cells(i, 14).Value
Next i

' Borramos la última columna innecesaria
Columns(14).EntireColumn.Delete

' Borrando la hoja innecesaria
Application.DisplayAlerts = False
Worksheets(1).Delete
Application.DisplayAlerts = True

' Acomodando las columnas
Columns(3).EntireColumn.Insert
Range("D:D").Copy
Range("L:L").PasteSpecial xlPasteAll

Range("K:K").Copy
Range("C:C").PasteSpecial xlPasteAll
Application.CutCopyMode = False

' Eliminando las columnas sobrantes
Columns(4).EntireColumn.Delete
Columns(10).EntireColumn.Delete

' Insertando unas columnas necesarias
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
cadenaOriginal = Array("-CL-EG", "-PR-EG", "-PR", "-CL", " T:52-56", " T:2XS-XL", "T:34-44", "T:34-48", "T:36-48", "T:46-48", " T:46-50", "T:50-54", "T:56-60", "T:62-66", "T:34/44", "T:34/48", "T:36/48", "T:46/48", "T:38-48", "T:50/54", " 50-54", "T:56/60", "T:62/66", "T:XXS-XXL", "T:2XS-2XL", "2XS/2XL", "XXS/XXL", "T:3XL-5XL", "T:2XS/2XL", "3XL/5XL", " envío gratis", " envio gratis", " rerda", " envío gratis", " envio gratis", " en cuotas", "en cuotas", " premium", "premium", "Unico", "Único", "Regulable", " cuotas", " talles especiales", "talle especial", " - ", " . ", "   ", "...", "..", "Unico", "Único", "")

' Definiendo el largo del array con un bucle While
Dim largo As Integer
largo = 0
Do While largo >= 0
    If cadenaOriginal(largo) = "" Then Exit Do
    largo = largo + 1
Loop

' Bucle. Debe coincidir el largo del arreglo con el fin del bucle
For i = 0 To largo
    Cells.Replace what:=cadenaOriginal(i), Replacement:="", LookAt:=xlPart, _
        searchorder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
Next

' ============================================================
' FORMATEANDO LA TABLA PARA DEJARLA BONITA

' Mini título
Range("H1").Value = "Cant"
Dim Codigucho As Range
Dim TextoOriginal As String


' Formateando las columnas
Columns("A").ColumnWidth = 11
Columns("A").Hidden = True
Columns("B").ColumnWidth = 19
Columns("C").ColumnWidth = 27
Columns("D").ColumnWidth = 45
Columns("E").ColumnWidth = 11
Columns("F").ColumnWidth = 15
Columns("G").ColumnWidth = 6
Columns("H").ColumnWidth = 5
Columns("I").ColumnWidth = 10.7
Columns("J").ColumnWidth = 10.7
Columns("J").Hidden = True

' Cambiando de lugar las columnas
' 1. Cortar Columna E y moverla después de la C (posición temporal)
Columns("E").Cut
Columns("D").Insert Shift:=xlToRight


' Formateando el dato del código
For i = 2 To ultima
    Set Codigucho = Worksheets("Planilla").Cells(i, 4)
    TextoOriginal = CStr(Codigucho.Value)
    With Codigucho
        .HorizontalAlignment = xlLeft
        .NumberFormat = "@"
        .Value = TextoOriginal
        .Font.color = vbWhite
        .Characters(Start:=1, Length:=7).Font.ColorIndex = xlAutomatic
    End With
Next i


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

' Filtra cantidad de rótulos
' última celda de la columna
Dim contador As Integer
contador = 0

' Con el bucle iremos subiendo y controlando si la celda superior es igual a la que estamos parada,
' de ser así, subimos a la que está arriba y borramos el valor de la que está abajo y otras correspondientes a las fechas número de venta y nombre de cliente.
Do While ultima - contador > 1
    Cells(ultima - contador, 11).Select
    
    ' Si no tiene nada es porque hay un retiro en Local
    If Cells(ultima - contador, 11).Value = "" Then
        Cells(ultima - contador, 11).Value = "Retira en Local"
        
        ' Completa el nombre del cliente que está vacío
        Cells(ultima - contador, 3).Value = Cells(ultima - contador, 13).Value
    
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
    
    ' Pintando la fila
    Call PintarFila("Planilla", ultima - contador, 1, 9)
    
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

' Borrando la última colummna ahora innecesaria
Columns(13).EntireColumn.Delete

' Redimensionado un par de columnas
With Range(Cells(2, 3), Cells(ultima, 4))
    .ShrinkToFit = True
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

' Formato a la descripción y los clientes.
With Range(Cells(2, 3), Cells(ultima, 4))
    .ShrinkToFit = True
    .WrapText = False
End With

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
    ' Le agrego saltos de líneas para evadir la parte que no fucniona de la impresora
    '.CenterHeader = vbNewLine & vbNewLine & vbNewLine & vbNewLine & "&B&20&F"
End With

' Ajustanto texto a la columna de los clientes
Range(Cells(2, 3), Cells(ultima, 3)).WrapText = True



' ==========================================================
' VALIDANDO NOMBRE DE ARCHIVO A GENERAR

' Variables a utilizar
Dim ruta As String
Dim nombre As String
Dim cuenta As String
Dim fecha As String

' Sirve para averiguar el nombre de la computadora actual
Dim ws As Object
Set ws = CreateObject("WScript.network")

' Asignando algunos valores de acuerdo en qué equipo de la red esté
If ws.ComputerName = "EDGAR" Then
    ruta = "D:\Web\Listados de Ventas Online\MELI"
    Debug.Print "Estoy en la computadora: " & ws.ComputerName
Else
    ruta = "\\EDGAR\Web\Listados de Ventas Online\MELI"
    Debug.Print "Estoy en una computadora de la red, llamada: " & ws.ComputerName
End If
Debug.Print "Se guardan los archivos en: " & ruta

fecha = Day(Date) & "-" & Month(Date) & "-" & Year(Date)

'Controlando si la carpeta existe, de lo contrario, crearla en local
If Dir(ruta, vbDirectory) <> "" Then
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
ActiveWorkbook.Sheets.Add(after:=ActiveWorkbook _
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

ActiveWorkbook.SaveAs nombre
ActiveWorkbook.Save

' Generando planilla para Depósito
Call depositoMeli

Sheets("Planilla").Activate
Range("A1").Activate


' Controlando si es la cuenta 2
If cuenta = 2 Then
    Call correo(ruta, i, ultima)
Else
    Exit Sub
End If

' ============== FIN DE TODA LA MACRO.
End Sub



Sub depositoMeli()
Attribute depositoMeli.VB_Description = "Regenera la planilla del Depósito"
Attribute depositoMeli.VB_ProcData.VB_Invoke_Func = "D\n14"
' Planilla para el Depósito
' GENERA UNA PLANILLA SÓLO PARA USO EXCLUSIVO DEL DEPOSITO

Dim rutaFormula As String
Dim carpeta As String
Dim i As Integer
Dim ultima As Byte
Dim meli As Boolean

carpeta = ActiveWorkbook.Path
rutaFormula = "'" & carpeta & "\..\" & "[Stock.XLS]Sheet1'!$A$2:$G$10000"
'rutaFormula = carpeta & "\..\" & "[Stock.XLS]Sheet1'!$A$2:$G$10000"


' Nueva hoja con nombre Depósito
ultima = Worksheets("Planilla").Cells(Rows.Count, 2).End(xlUp).Row - 1
Call CrearHoja(ActiveWorkbook, "Depósito")
ActiveWorkbook.Worksheets("Depósito").Activate

' Creando las columnas
With Worksheets("Depósito")
    ' Limpia la hoja, por las dudas, de todo contenido y formato
    .Cells.Clear
    .Cells.Select
    .Cells.ClearFormats
    .Cells.Font.Size = 11
    .Cells(1, 1).Value = "Nº Venta"
    .Cells(1, 2).Value = "Cliente"
    .Cells(1, 3).Value = "Código"
    .Cells(1, 4).Value = "Descripción"
    .Cells(1, 5).Value = "Color"
    .Cells(1, 6).Value = "Talle"
    .Cells(1, 7).Value = "Cant"
    .Cells(1, 8).Value = "Ubicación"
End With

' Ocultando columnas al pede
With Worksheets("Depósito")
    .Columns("A").Hidden = True
    .Columns("B").Hidden = True
End With

' Completando los datos
For i = 2 To ultima
    With Worksheets("Depósito")

        ' Venta
        If .Cells(i, 2).Value = "" Then
            .Cells(i, 1).Value = ""
        Else
            .Cells(i, 1).Value = "'" & Worksheets("Planilla").Cells(i, 2).Value
        End If
        
        ' Cliente
        .Cells(i, 2).Value = Worksheets("Planilla").Cells(i, 3).Value
        
        ' Descripción
        .Cells(i, 4).Value = "=VLOOKUP(LEFT(C" & i & ", 7)," & rutaFormula & ",2,FALSE)"
        
        ' Reemplazando la descripción por las fórmulas en Planilla
        Worksheets("Planilla").Cells(i, 5).Value = "=VLOOKUP(LEFT(D" & i & ", 7)," & rutaFormula & ",2,FALSE)"
        
        ' Código
        .Cells(i, 3).Value = "'" & Worksheets("Planilla").Cells(i, 4).Value
        .Cells(i, 3).Font.color = vbWhite
        .Cells(i, 3).Characters(Start:=1, Length:=7).Font.ColorIndex = xlAutomatic
        .Cells(i, 3).HorizontalAlignment = xlLeft
        
        ' Color
        .Cells(i, 5).Value = Worksheets("Planilla").Cells(i, 6).Value
        
        ' Talle
        .Cells(i, 6).Value = Worksheets("Planilla").Cells(i, 7).Value
        
        ' Cantidad
        .Cells(i, 7).Value = Worksheets("Planilla").Cells(i, 8).Value
           
        ' La ubicación. Si está en planta baja agregar esta aclaración
        If Worksheets("Planilla").Cells(i, 9).Value = "planta baja" Then
            .Cells(i, 8).Value = "10. PLANTA BAJA"
        Else
            .Cells(i, 8).Formula = "=VLOOKUP(LEFT(C" & i & ", 7)," & rutaFormula & ",3,FALSE)"
        End If
    
    End With
Next i

' Ordenando alfabéticamente esta columna de ubicación
With Worksheets("Depósito").Range("A1:H1")
    .AutoFilter
    .Rows("1").RowHeight = 27
    .Font.Bold = True
    .Font.Size = 13.5
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
End With

Worksheets("Depósito").Range("A1").CurrentRegion.Sort Key1:=Range("H1"), Order1:=xlAscending, Header:=xlGuess
With Selection
    .AutoFilter
End With

With Worksheets("Depósito").Range("A1").CurrentRegion
    .Columns.AutoFit
    .Columns("E").ColumnWidth = 30
    .Columns("A").Hidden = True
    .Columns("B").Hidden = True
End With

' Colocando totales de productos y dando formato
With Worksheets("Depósito")
    .Cells(ultima + 1, 6).Value = "TOTALES:"
    .Cells(ultima + 1, 7).Select
    .Cells(ultima + 1, 7).Value = "=SUM(G2:G" & ultima & ")"
    .Range(Cells(ultima + 1, 6), Cells(ultima + 1, 7)).Select
End With

With Selection
    .Font.Bold = True
    .Font.Size = 16
    .HorizontalAlignment = xlRight
    .VerticalAlignment = xlBottom
End With


With Worksheets("Depósito").Range(Cells(1, 3), Cells(ultima, 8))
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
    .PrintArea = ActiveSheet.Range("A1:H" & ultima + 1).Address
    .Zoom = False
    .FitToPagesTall = 1
    .FitToPagesWide = 1
    .CenterHeader = "&B&20&F" & vbNewLine & "SOLO PARA USO EN DEPOSITO"
    ' Agrego saltos de líne para evitar zona blanca
    '.CenterHeader = vbNewLine & vbNewLine & vbNewLine & "&B&20&F" & vbNewLine & "SOLO PARA USO EN DEPOSITO"
End With

' Colorear filas
For i = 2 To ultima
    Call PintarFila("Depósito", i, 1, 8)
Next i

Worksheets("Planilla").Activate

' Genera TXT. Comprueba de qué cuenta es.
Dim cuenta As Byte
If Left(ActiveWorkbook.Name, 8) = "CUENTA 2" Then
    cuenta = 2
Else
    cuenta = 1
End If

Call exportarTxt("MELI" & cuenta, ActiveWorkbook)
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
    If packar.Worksheets(1).Range("A1").Value <> hoy Then
        packar.Worksheets(1).Range("A1").Value = Date
        continuacion = 7
        
        ' Borrando el contenido viejo
        packar.Worksheets(1).Range("A9:C39").ClearContents
    Else
        continuacion = packar.Worksheets(1).Cells(39, 1).End(xlUp).Row - 1
        Debug.Print continuacion
    End If
     
    
    ' Completando la información
    For i = 2 To ultima
        ' Asignando el valor a cada N° vta.
        numVenta = planilla.Worksheets(1).Cells(i, 2).Value
        Cliente = planilla.Worksheets(1).Cells(i, 3).Value
        tn = planilla.Worksheets(1).Cells(i, 11).Value
        
        ' Recorremos la planilla del Correo
        ' Controlamos que el número de venta esté completo
        ' y además que NO SEA un retiro en Local
        If numVenta <> "" And planilla.Worksheets(1).Cells(i, 9).Value <> "Retira en Local" Then
            With packar.Worksheets(1)
                .Cells(i + continuacion, 1).Value = "'" & numVenta
                .Cells(i + continuacion, 2).Value = Cliente
                .Cells(i + continuacion, 3).Value = tn
            End With
        End If
    Next i
End Function


Function CrearHoja(archivo As Workbook, nombreHoja As String) As Boolean
    ' controla si una hoja existe o no
    Dim existe As Boolean
     
    On Error Resume Next
    existe = archivo.Worksheets(nombreHoja).Name <> ""
     
    If Not existe Then
        archivo.Worksheets.Add(after:=Worksheets(Worksheets.Count)).Name = nombreHoja
    End If
     
    CrearHoja = existe
     
End Function



Sub exportarTxt(carpeta As String, planillaActual As Workbook)

' GENERA UN ARCHIVO DE TEXTO PARA IMPORTAR AL D.F.
' Variable temporal

Dim Fila As Byte
Dim codigo As String
Dim Equivalencia As Workbook
Dim camino As String
camino = planillaActual.Path & "\..\Equivalencia.XLS"
Set Equivalencia = Workbooks.Open(camino, False, True)



Dim textoArchivo As String
Dim server As String
Dim txt As String
txt = "Exportar TXT"
Dim carpetaDestino As String

Dim nombreArchivo As String
Dim limite As Byte
Dim item As Variant
Dim i As Byte
Dim ultimaFila As Byte
Dim resto As Byte
Dim cantArchivos As Byte
Dim ruta As Range

Set ruta = Equivalencia.Worksheets(1).Range("A1:G10000")
ultimaFila = planillaActual.Worksheets("Planilla").Cells(Rows.Count, 2).End(xlUp).Row - 1
nombreArchivo = planillaActual.Name
server = "\\SER-DF\D\A Remitar TXT"
carpetaDestino = "\" & carpeta & "\"
limite = ultimaFila
Debug.Print nombreArchivo

' Crea la hoja
Call CrearHoja(planillaActual, txt)

' Limpiar la hoja
planillaActual.Worksheets(txt).Cells.Clear

' Completa planilla para exportar
For Fila = 1 To ultimaFila - 1
    With planillaActual.Worksheets(txt)
        ' 1º) Stock
        .Cells(Fila, 1).Value = "'" & planillaActual.Worksheets("Depósito").Cells(Fila + 1, 7).Value
    
        ' 2º Codigo
        codigo = "" & planillaActual.Worksheets("Depósito").Cells(Fila + 1, 3).Value & ""
        .Cells(Fila, 2).Value = Left(codigo, 7)

        ' 3° Color
        On Error Resume Next
        .Cells(Fila, 3) = "'" & Application.WorksheetFunction.VLookup(codigo, ruta, 4, False)
        
        ' 4° Talle
        .Cells(Fila, 4) = "'" & Application.WorksheetFunction.VLookup(codigo, ruta, 6, False)
    End With
Next Fila

' Corrección de "Depósito" ==========
For Fila = 2 To ultimaFila
    With planillaActual.Worksheets("Depósito")
        ' Corrección del Color
        .Cells(Fila, 5).Activate
        .Cells(Fila, 3).Activate
        .Cells(Fila, 5).Value = Application.WorksheetFunction.VLookup(.Cells(Fila, 3).Value, ruta, 4, False) & ". " & Application.WorksheetFunction.VLookup(.Cells(Fila, 3).Value, ruta, 5, False)
        
        ' Corrección del Talle
        .Cells(Fila, 6).Value = "'" & Application.WorksheetFunction.VLookup(.Cells(Fila, 3).Value, ruta, 6, False)
    End With
Next Fila

' Corrección de "Planilla" =========
For Fila = 2 To ultimaFila
    With planillaActual.Worksheets("Planilla")
        ' Corrección del Color
        .Cells(Fila, 6).Value = Application.WorksheetFunction.VLookup(.Cells(Fila, 4).Value, ruta, 4, False) & ". " & Application.WorksheetFunction.VLookup(.Cells(Fila, 4).Value, ruta, 5, False)
        
        ' Corrección del Talle
        .Cells(Fila, 7).Value = "'" & Application.WorksheetFunction.VLookup(.Cells(Fila, 4).Value, ruta, 6, False)
    End With
Next Fila


' Ajuste ultima Fila - HARDCODEO ESTO PARA PROBAR
ultimaFila = ultimaFila - 1

' Si se pasa del tope (30 líneas), serán "n" archivos con 30 líneas y otro con el resto
' de items que quedaron fuera. Sería el resto de una división, el módulo.
resto = ultimaFila Mod limite
cantArchivos = Int(ultimaFila / limite) + 1
Debug.Print "Archivos a importar: " & cantArchivos

' Cierra el archivo con el listado de las equivalencias
Equivalencia.Close False

' Generación del txt
Call generarTxt(Fila, ultimaFila, "", cantArchivos, planillaActual, carpetaDestino, limite, resto, server, txt)
planillaActual.Sheets("Depósito").Activate
End Sub

Function generarTxt(Fila, ultimaFila, textoArchivo, cantArchivos, archivoFuente, carpetaDestino, limite, resto, server, hoja)
Dim rutaArchivo As String
Dim nombreArchivo As String
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
        archivoFuente.Sheets(hoja).Activate
        Cells(Fila, 1).Activate
        textoArchivo = textoArchivo _
            & Cells(Fila, 1).Value _
            & "+" & archivoFuente.Sheets(hoja).Cells(Fila, 2).Value _
            & "!" & archivoFuente.Sheets(hoja).Cells(Fila, 3).Value _
            & "!" & archivoFuente.Sheets(hoja).Cells(Fila, 4).Value _
            & vbNewLine
            Debug.Print "Archivo N°: " & i, "Fila N° :" & Fila
    Next Fila
    
    ' Si es mayor a uno, se van nombrando incrementalmente
    If cantArchivos > 1 Then
        nombreArchivo = Left(archivoFuente.Name, Len(archivoFuente.Name) - 5) & " - " & i & ".txt"
    Else
        nombreArchivo = Left(archivoFuente.Name, Len(archivoFuente.Name) - 5) & ".txt"
    End If
    
    rutaArchivo = server & carpetaDestino & nombreArchivo
    Debug.Print textoArchivo
    Open rutaArchivo For Output As #1
    Print #1, textoArchivo
    Close #1
    
    MsgBox "Datos exportados con éxito a " & rutaArchivo, vbInformation, "Cargar detalle desde txt"
Next i
End Function



