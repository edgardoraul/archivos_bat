Attribute VB_Name = "MELI"
Option Explicit

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
    MsgBox ("Ya le diste formato a esta planilla. Prob� con otra.")
    Range("A1").Select
    Exit Sub
End If

' ============================================================
' Creando las hojas nuevas
Cells.Select
Cells.ClearFormats
ActiveWorkbook.Sheets.Add(after:=ActiveWorkbook _
    .Worksheets(ActiveWorkbook.Worksheets.Count)).Name = "Planilla"
Application.Worksheets(1).Select


' Dando valor a la �ltima fila
ultima = Cells(Rows.Count, 1).End(xlUp).Row

' Borrando una columna del barrio
Range("AH:AH").Delete

' Corrigiendo informaci�n
Cells(1, 49).Activate

' Bucle que recorre las filas y obtiene el nombre+apellido y/o raz�n social.
For i = 2 To ultima
    Debug.Print Cells(i, 49).Value
    Cells(i, 49).Activate
    
    ' Obtenci�n del usuario en may�sculas
    Cells(i, 49).Value = Cells(i, 4).Value
Next i


' Copiando informaci�n importante a la Planilla
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

' Borramos la �ltima columna innecesaria
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

' Colocando t�tulos
Range("B1").Value = "N� de Venta"
Range("C1").Value = "Cliente"
Range("D1").Value = "Descripci�n"
Range("E1").Value = "C�digo"
Range("I1").Value = "Detalles"
Range("J1").Value = "Firma Control"


' ============================================================
' BORRANDO TEXTOS INNECESARIOS

' Definiendo la variable como arreglo
Dim cadenaOriginal As Variant

' Listado de expresiones a borrar. Deber tener al �ltimo el "" para quelo tome el bucle.
cadenaOriginal = Array("-CL-EG", "-PR-EG", "-PR", "-CL", " T:52-56", " T:2XS-XL", "T:34-44", "T:34-48", "T:36-48", "T:46-48", " T:46-50", "T:50-54", "T:56-60", "T:62-66", "T:34/44", "T:34/48", "T:36/48", "T:46/48", "T:38-48", "T:50/54", " 50-54", "T:56/60", "T:62/66", "T:XXS-XXL", "T:2XS-2XL", "2XS/2XL", "XXS/XXL", "T:3XL-5XL", "T:2XS/2XL", "3XL/5XL", " env�o gratis", " envio gratis", " rerda", " env�o gratis", " envio gratis", " en cuotas", "en cuotas", " premium", "premium", "Unico", "�nico", "Regulable", " cuotas", " talles especiales", "talle especial", " - ", " . ", "   ", "...", "..", "Unico", "�nico", "")

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
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
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


' Dando formato a la columna de los n�meros de ventas
Range(Cells(2, 2), Cells(ultima, 2)).Select
With Selection
    .NumberFormat = "###"
    .HorizontalAlignment = xlRight
End With

' Filtra cantidad de r�tulos
' �ltima celda de la columna
Dim contador As Integer
contador = 0

' Con el bucle iremos subiendo y controlando si la celda superior es igual a la que estamos parada,
' de ser as�, subimos a la que est� arriba y borramos el valor de la que est� abajo y otras correspondientes a las fechas n�mero de venta y nombre de cliente.
Do While ultima - contador > 1
    Cells(ultima - contador, 11).Select
    
    ' Si no tiene nada es porque hay un retiro en Local
    If Cells(ultima - contador, 11).Value = "" Then
        Cells(ultima - contador, 11).Value = "Retira en Local"
        
        ' Completa el nombre del cliente que est� vac�o
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
    contador = contador + 1
Loop

' Colocando totales de R�tulos y dando formato. Pero antes deben estar filtrados cu�ntos son realmente
Cells(ultima + 1, 3).Value = " ROTULOS"
Cells(ultima + 1, 2).Select
Cells(ultima + 1, 2).Value = "=COUNTA(K2:K" & ultima & ")-COUNTIF(K2:K" & ultima & ", ""Retira en Local"")"
Range(Cells(ultima + 1, 2), Cells(ultima + 1, 3)).Select
    With Selection
        .Font.Bold = True
        .Font.Size = 15
        .VerticalAlignment = xlBottom
    End With

' Borrando la �ltima colummna ahora innecesaria
Columns(13).EntireColumn.Delete

' Redimensionado un par de columnas
With Range(Cells(2, 3), Cells(ultima, 4))
    .ShrinkToFit = True
End With

' ==========================================================
' FORMATO PARA IMPRIMIR UNA SOLA P�GINA

' Delimitando el tama�o de hojas y m�rgenes
Dim filasTotales As Integer
filasTotales = ultima + 1

' Formatea la �ltima columna que NO saldr� impresa, s�lo para acomodar, nada m�s
Range("K:K").Columns.AutoFit

' Acomoda el texto de las celdas con datos
Range(Cells(2, 1), Cells(ultima, 11)).WrapText = True

' Ajusta autom�ticamente la altura de las filas
Range(Cells(2, 1), Cells(ultima, 10)).Rows.AutoFit

' Formato a la descripci�n y los clientes.
With Range(Cells(2, 3), Cells(ultima, 4))
    .ShrinkToFit = True
    .WrapText = False
End With

' Borde externo faltante
Range(Cells(2, 1), Cells(ultima, 11)).BorderAround LineStyle:=xlContinuous

' Formato de impresi�n
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
    ' Le agrego saltos de l�neas para evadir la parte que no fucniona de la impresora
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

' Asignando algunos valores de acuerdo en qu� equipo de la red est�
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

' Preguntando al usuario qu� cuenta es
cuenta = Application.InputBox(Prompt:="�Qu� cuenta de MercadoLibre es? �1 � 2?", Title:="Cuenta de MercadoLibre", Default:=1)
If cuenta <> 1 And cuenta <> 2 Then
    MsgBox ("�Cuenta 1 � 2?. Eleg� bien.")
    cuenta = Application.InputBox(Prompt:="�Qu� cuenta de MercadoLibre es? �1 � 2?", Title:="Cuenta de MercadoLibre", Default:=1)
End If
' En base a la respuesta, determinar la carpeta definitiva
ruta = ruta & cuenta & "\"


' Definiendo unas variables
Dim archivos As String
Dim u As Integer
Dim denominacion As String
    
' Preparaci�n de variables
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

' Controlando que no se est� duplicando el mismo archivo con otro nombre
If ActiveWorkbook.Name = Sheets("Listado").Cells(u - 1, 1).Value Then
    MsgBox ("Ya creaste este archivo antes. Gener� uno nuevo.")
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
    MsgBox ("Eleg�: 1 � 2")
End If

Sheets("Planilla").Range("A1").Select

ActiveWorkbook.SaveAs nombre
ActiveWorkbook.Save

' Generando planilla para Dep�sito
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
Attribute depositoMeli.VB_Description = "Regenera la planilla del Dep�sito"
Attribute depositoMeli.VB_ProcData.VB_Invoke_Func = "D\n14"
' Planilla para el Dep�sito
' GENERA UNA PLANILLA S�LO PARA USO EXCLUSIVO DEL DEPOSITO

Dim rutaFormula As String
Dim carpeta As String
Dim i As Byte
Dim ultima As Byte
Dim meli As Boolean

carpeta = ActiveWorkbook.Path
rutaFormula = "'" & carpeta & "\..\" & "[Stock.XLS]Sheet1'!$A$2:$G$10000"

' Nueva hoja con nombre Dep�sito
ultima = Sheets("Planilla").Cells(Rows.Count, 2).End(xlUp).Row - 1
Call CrearHoja(ActiveWorkbook, "Dep�sito")
ActiveWorkbook.Sheets("Dep�sito").Activate

' Creando las columnas
With Sheets("Dep�sito")
    ' Limpia la hoja, por las dudas, de todo contenido y formato
    .Cells.Clear
    .Cells(1, 1).Value = "N� Venta"
    .Cells(1, 2).Value = "Cliente"
    .Cells(1, 3).Value = "Descripci�n"
    .Cells(1, 4).Value = "C�digo"
    .Cells(1, 5).Value = "Color"
    .Cells(1, 6).Value = "Talle"
    .Cells(1, 7).Value = "Cantidad"
    .Cells(1, 8).Value = "Ubicaci�n"
End With

' Completando los datos
For i = 2 To ultima
    ' Venta
    If Sheets(1).Cells(i, 2).Value = "" Then
        Cells(i, 1).Value = ""
    Else
        Cells(i, 1).Value = "'" & Sheets(1).Cells(i, 2).Value
    End If
    
    ' Cliente
    Cells(i, 2).Value = Sheets(1).Cells(i, 3).Value
    
    ' Descripci�n
    Cells(i, 3).Value = "=VLOOKUP(LEFT(D" & i & ", 7)," & rutaFormula & ",2,FALSE)"
    
    ' C�digo
    Cells(i, 4).Value = "'" & Sheets(1).Cells(i, 5).Value
    
    ' Color
    Cells(i, 5).Value = Sheets(1).Cells(i, 6).Value
    
    ' Talle
    Cells(i, 6).Value = Sheets(1).Cells(i, 7).Value
    
    ' Cantidad
    Cells(i, 7).Value = Sheets(1).Cells(i, 8).Value
       
    ' La ubicaci�n. Si est� en planta baja agregar esta aclaraci�n
    If Sheets(1).Cells(i, 9).Value = "planta baja" Then
        Cells(i, 8).Value = "10. PLANTA BAJA"
    Else
        Cells(i, 8).Formula = "=VLOOKUP(LEFT(D" & i & ", 7)," & rutaFormula & ",3,FALSE)"
    End If
    
Next i

' Ordenando alfab�ticamente esta columna de ubicaci�n
With Range("A1:H1")
    .AutoFilter
    .Rows("1").RowHeight = 27
    .Font.Bold = True
    .Font.Size = 12
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
End With

Range("A1").CurrentRegion.Sort Key1:=Range("H1"), Order1:=xlAscending, Header:=xlGuess
With Selection
    .AutoFilter
End With

With Range("A1").CurrentRegion
    .Columns.AutoFit
End With

' Colocando totales de productos y dando formato
Cells(ultima + 1, 6).Value = "TOTALES:"
Cells(ultima + 1, 7).Select
Cells(ultima + 1, 7).Value = "=SUM(G2:G" & ultima & ")"
Range(Cells(ultima + 1, 6), Cells(ultima + 1, 7)).Select
With Selection
    .Font.Bold = True
    .Font.Size = 15
    .HorizontalAlignment = xlRight
    .VerticalAlignment = xlBottom
End With


' Colocando el total de r�tulos a imprimir
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

' Formato de impresi�n
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
    .PrintArea = ActiveSheet.Range("A1:H" & ultima + 1).Address
    .Zoom = False
    .FitToPagesTall = 1
    .FitToPagesWide = 1
    .CenterHeader = "&B&20&F" & vbNewLine & "SOLO PARA USO EN DEPOSITO"
    ' Agrego saltos de l�ne para evitar zona blanca
    '.CenterHeader = vbNewLine & vbNewLine & vbNewLine & "&B&20&F" & vbNewLine & "SOLO PARA USO EN DEPOSITO"
End With

Sheets("Planilla").Activate

' Genera TXT. Comprueba de qu� cuenta es.
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
' GENERA UN LISTADO DE VENTAS Y N� GUIAS PARA EL CORREO
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
     
    
    ' Completando la informaci�n
    For i = 2 To ultima
        ' Asignando el valor a cada N� vta.
        numVenta = planilla.Sheets(1).Cells(i, 2).Value
        Cliente = planilla.Sheets(1).Cells(i, 3).Value
        tn = planilla.Sheets(1).Cells(i, 11).Value
        
        ' Recorremos la planilla del Correo
        ' Controlamos que el n�mero de venta est� completo
        ' y adem�s que NO SEA un retiro en Local
        If numVenta <> "" And planilla.Sheets(1).Cells(i, 9).Value <> "Retira en Local" Then
            packar.Sheets(1).Cells(i + continuacion, 1).Value = "'" & numVenta
            packar.Sheets(1).Cells(i + continuacion, 2).Value = Cliente
            packar.Sheets(1).Cells(i + continuacion, 3).Value = tn
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

Dim fila As Byte
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

Set ruta = Equivalencia.Sheets(1).Range("A1:G10000")
ultimaFila = planillaActual.Sheets("Planilla").Cells(Rows.Count, 2).End(xlUp).Row - 1
nombreArchivo = planillaActual.Name
server = "\\SER-DF\D\A Remitar TXT"
carpetaDestino = "\" & carpeta & "\"
limite = ultimaFila
Debug.Print nombreArchivo

' Crea la hoja
Call CrearHoja(planillaActual, txt)

' Limpiar la hoja
planillaActual.Sheets(txt).Cells.Clear

' Completa planilla para exportar
For fila = 1 To ultimaFila - 1
    With planillaActual.Sheets(txt)
        ' 1�) Stock
        .Cells(fila, 1).Value = "'" & planillaActual.Sheets("Dep�sito").Cells(fila + 1, 7).Value
    
        ' 2� Codigo
        codigo = "" & planillaActual.Sheets("Dep�sito").Cells(fila + 1, 4).Value & ""
        .Cells(fila, 2).Value = Left(codigo, 7)

        ' 3� Color
        On Error Resume Next
        .Cells(fila, 3) = "'" & Application.WorksheetFunction.VLookup(codigo, ruta, 4, False)
        
        ' 4� Talle
        .Cells(fila, 4) = "'" & Application.WorksheetFunction.VLookup(codigo, ruta, 6, False)
    End With
Next fila

' Ajuste ultima Fila - HARDCODEO ESTO PARA PROBAR
ultimaFila = ultimaFila - 1

' Si se pasa del tope (30 l�neas), ser�n "n" archivos con 30 l�neas y otro con el resto
' de items que quedaron fuera. Ser�a el resto de una divisi�n, el m�dulo.
resto = ultimaFila Mod limite
cantArchivos = Int(ultimaFila / limite) + 1
Debug.Print "Archivos a importar: " & cantArchivos

' Cierra el archivo con el listado de las equivalencias
Equivalencia.Close False

' Generaci�n del txt
Call generarTxt(fila, ultimaFila, "", cantArchivos, planillaActual, carpetaDestino, limite, resto, server, txt)
planillaActual.Sheets("Dep�sito").Activate
End Sub

Function generarTxt(fila, ultimaFila, textoArchivo, cantArchivos, archivoFuente, carpetaDestino, limite, resto, server, hoja)
Dim rutaArchivo As String
Dim nombreArchivo As String
Dim i As Byte
Dim tope As Byte
fila = 0


' Generaci�n del txt
For i = 1 To cantArchivos
tope = i * limite
    If i = cantArchivos Then
        tope = ultimaFila
    End If
    
    For fila = (limite * (i - 1)) + 1 To tope
        archivoFuente.Sheets(hoja).Activate
        Cells(fila, 1).Activate
        textoArchivo = textoArchivo _
            & Cells(fila, 1).Value _
            & "+" & archivoFuente.Sheets(hoja).Cells(fila, 2).Value _
            & "!" & archivoFuente.Sheets(hoja).Cells(fila, 3).Value _
            & "!" & archivoFuente.Sheets(hoja).Cells(fila, 4).Value _
            & vbNewLine
            Debug.Print "Archivo N�: " & i, "Fila N� :" & fila
    Next fila
    
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
    
    MsgBox "Datos exportados con �xito a " & rutaArchivo, vbInformation, "Cargar detalle desde txt"
Next i
End Function



