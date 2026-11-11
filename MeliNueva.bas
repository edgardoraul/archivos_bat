Attribute VB_Name = "MeliNueva"
Option Explicit

Public RUTA As String
Public Const MIPC As String = "D:\Web\Listados de Ventas Online\MELY"
Public Const MIPCENRED As String = "\\EDGAR\Web\Listados de Ventas Online\MELY"
Public UltimaTotal As Long
Public IdCliente As String
Public Provincias As Variant
Public Respuesta As String

Public Sub Workbook_Open()
    Application.OnKey "^+K", "AA_MELI_Nueva"
End Sub
Sub BB_SEPARAR_SUCURSALES()
    Dim i As Integer
    Dim c As Integer
    Dim ultimaFila As Long
    Dim Codigo As String
    Dim Color As String
    Dim Talle As String
    Dim Cantidad As Integer
    
    
    ' Analiza si esta venta corresponde a otra sucursal
    ' Obtener la base de la ruta donde están los archivos
    Call EstablecerRuta

    ' Determinar los dos archivos a utilizar: la planilla a analizar
    ' y el listado con el stock de las sucursales.
    Dim PlanillaActual As Workbook
    Dim StockSucursales As Workbook
    
    Set PlanillaActual = ActiveWorkbook
    RUTA = RUTA & "\..\"
    With PlanillaActual.Worksheets(1)
        ultimaFila = .Cells(.Rows.Count, 3).End(xlUp).Row - 1
    End With
    
    
    ' Abre el CSV
    Set StockSucursales = Workbooks.Open(RUTA & "StockDeArticulos.csv", Local:=True)
    
    ' Guarda el archivo abierto como un libro de Excel moderno (.xlsx)
    Application.DisplayAlerts = False
    StockSucursales.SaveAs fileName:=RUTA & "StockDeArticulos.xlsx", FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True
    
    ' Lógiga de búsqueda y análisis desde la planilla original
    With PlanillaActual.Worksheets(1)
        For i = 2 To ultimaFila
            Codigo = Left(.Cells(i, 3).Value, 7)
            Color = .Cells(i, 5).Value
            If Color <> "" Then
                Color = "'" & Trim(Split(Color, ".")(0))
            End If
            Talle = .Cells(i, 6).Value
            Cantidad = .Cells(i, 7).Value
            
            Debug.Print Codigo, Color, Talle
            Call BuscaStock(StockSucursales, Codigo, Color, Talle, Cantidad)
            .Cells(i, 8).Value = Respuesta
        Next i
    End With
    
    ' Lo cerramos para guardar memoria
    StockSucursales.Close SaveChanges:=False

End Sub
Function BuscaStock(Base As Workbook, Codigo As String, Color As String, Talle As String, Cantidad As Integer)
    Dim wsStock As Worksheet
    Dim colSKU As Integer, colColor As Integer, colTalle As Integer, colPieza As Integer
    Dim rngEncontrado As Range
    Dim primerCelda As String
    Dim stockPieza As Integer
    Dim iCol As Long, ultimaColumnaStock As Long
    Dim mayorStock As Integer, stockActual As Integer
    Dim depositoElegido As String
    
    Set wsStock = Base.Worksheets(1)
    Base.Activate
    wsStock.Activate
    Respuesta = ""
    
    ' Identificar columnas por encabezado
    With wsStock
        colSKU = .Rows(1).Find(What:="SKU", LookAt:=xlWhole).Column
        colColor = .Rows(1).Find(What:="COLOR", LookAt:=xlWhole).Column
        colTalle = .Rows(1).Find(What:="TALLE", LookAt:=xlWhole).Column
        colPieza = 5 ' Columna E (PIEZA)
        
        ultimaColumnaStock = .Cells(1, .Columns.Count).End(xlToLeft).Column
        
        ' Búsqueda rápida del SKU en su columna específica
        Set rngEncontrado = .Columns(colSKU).Find(What:=Codigo, LookAt:=xlWhole, MatchCase:=False)
        
        If Not rngEncontrado Is Nothing Then
            primerCelda = rngEncontrado.Address
            Do
                ' Verificar si en la misma fila coinciden Color y Talle
                If .Cells(rngEncontrado.Row, colColor).Value = Color And _
                   .Cells(rngEncontrado.Row, colTalle).Value = Talle Then
                   
                    ' Coincidencia triple encontrada
                    stockPieza = Val(.Cells(rngEncontrado.Row, colPieza).Value)
                    .Cells(rngEncontrado.Row, colPieza).Activate
                    
                    If stockPieza >= 2 Then
                        Debug.Print "ok"
                        Exit Function
                    Else
                        mayorStock = stockPieza
                        depositoElegido = "PIEZA"
                        For iCol = (colPieza + 1) To ultimaColumnaStock
                            stockActual = Val(.Cells(rngEncontrado.Row, iCol).Value)
                            If stockActual > mayorStock Then
                                mayorStock = stockActual
                                depositoElegido = .Cells(1, iCol).Value
                            End If
                        Next iCol
                        MsgBox "Depósito: " & depositoElegido & " -> : " & mayorStock
                        Debug.Print "Depósito: " & depositoElegido & " -> : " & mayorStock
                        
                        Respuesta = depositoElegido
                        
                        Exit Function
                    End If
                    
                End If
                
                Set rngEncontrado = .Columns(colSKU).FindNext(rngEncontrado)
            Loop While Not rngEncontrado Is Nothing And rngEncontrado.Address <> primerCelda
        End If
        
        Debug.Print "Artículo no encontrado en stock."
        
        
    End With
    
    
    
End Function

Sub AA_MELI_Nueva()
' GENERACION DE PLANILLAS DE MELI AÑO 2027
    Call GuardarCopiaConSecuencial
End Sub

Public Sub EstablecerRuta()
    Dim pcName As String
    Dim workgroupName As String
    Dim objWMIService As Object
    Dim colItems As Object
    Dim objItem As Object
    
    pcName = UCase(Environ("COMPUTERNAME"))
    
    On Error Resume Next
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    Set colItems = objWMIService.ExecQuery("Select Workgroup from Win32_ComputerSystem")
    For Each objItem In colItems
        workgroupName = UCase(objItem.Workgroup)
    Next objItem
    On Error GoTo 0
    
    If pcName = "EDGAR" Then
        RUTA = MIPC
    ElseIf workgroupName = "RERDA" Then
        RUTA = MIPCENRED
    Else
        RUTA = ""
        MsgBox "Debes estar en la red del local de Rerda para trabajar y en su grupo de trabajo", vbExclamation, "Acceso restringido"
    End If
End Sub

Sub GuardarCopiaConSecuencial()
    Dim fso As Object
    Dim folderObj As Object
    Dim fileObj As Object
    Dim wbOriginal As Workbook
    Dim wbTemp As Workbook
    Dim tempPath As String
    Dim finalPath As String
    Dim todayStr As String
    Dim filePrefix As String
    Dim fileName As String
    Dim currentNum As Long
    Dim maxNum As Long
    Dim numStr As String
    Dim posDot As Long
    Dim UltFilaAcum As Long
    Dim UltFilaParcial As Long
    Dim PlanillaNumero As Byte
    

    ' 1. Verificar/Obtener la RUTA global
    If RUTA = "" Then Call EstablecerRuta
    If RUTA = "" Then Exit Sub

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(RUTA) Then
        MsgBox "La carpeta de destino no está accesible:" & vbCrLf & RUTA, vbCritical, "Error de Ruta"
        Exit Sub
    End If

    Set wbOriginal = ActiveWorkbook
    tempPath = wbOriginal.Path & "\MELI-Temp.xlsx"
    
    On Error Resume Next
    If fso.FileExists(tempPath) Then fso.DeleteFile tempPath, True
    On Error GoTo 0

    ' 2. Guardar como excel el libro actual
    wbOriginal.SaveAs fileName:=tempPath, _
        FileFormat:=xlOpenXMLWorkbook, _
        CreateBackup:=False
    
    ' 3. Abrir la copia temporal para trabajar
    Set wbTemp = Workbooks.Open(tempPath)
    
    ' 4. Limpieza de datos pasando el libro como parámetro
    Call LimpiezaDatos(wbTemp)

    ' 5. Determinar fecha
    todayStr = Format(Date, "yyyy-mm-dd")
    filePrefix = todayStr & ". "

    ' 6. Analizar archivos correlativos (Independiente de la fecha)
    maxNum = 0
    Set folderObj = fso.GetFolder(RUTA)

    For Each fileObj In folderObj.Files
        fileName = fileObj.Name
        
        ' Verificar que contenga la estructura " - MELI" al final antes de la extensión
        posDot = InStrRev(fileName, ".")
        If posDot > 0 Then fileName = Left(fileName, posDot - 1)
        
        ' Si el archivo termina en " - MELI", extrae los 5 dígitos numéricos justo anteriores
        If Right(fileName, 7) = " - MELI" Then
            numStr = Mid(fileName, Len(fileName) - 11, 5) ' Extrae el bloque de 5 dígitos
            If IsNumeric(numStr) Then
                currentNum = CLng(numStr)
                If currentNum > maxNum Then maxNum = currentNum
            End If
        End If
    Next fileObj

    ' 7. Generar nuevo nombre
    numStr = Format(maxNum + 1, "00000")
    finalPath = RUTA & "\" & filePrefix & numStr & " - MELI.xlsx"
    
    ' 8. Controles de archivos de datos
    If Not fso.FileExists(RUTA & "\..\Stock.XLS") Then
        MsgBox "El archivo Stock.XLS debe estar en la carpeta " & vbNewLine & RUTA, vbCritical, "Error"
        Exit Sub
    End If
        
    If Not fso.FileExists(RUTA & "\..\Equivalencia.XLS") Then
        MsgBox "El archivo Equivalencia.XLS debe estar en la carpeta " & vbNewLine & RUTA, vbCritical, "Error"
        wbTemp.Close SaveChanges:=False
        Exit Sub
    End If
    
    ' 9. Completado de información pasando el libro
    Call CompletaInfo(wbTemp)
    
    ' 10. Formato a Ventas
    Call FormatoTabla(wbTemp, wbTemp.Worksheets("Ventas"), True)
    
    ' 11. Guardar y cerrar la copia procesada
    wbTemp.SaveAs fileName:=finalPath, FileFormat:=51
    
    ' 12. Crear pestaña para Depósito
    Call AB_DepositoMeli
End Sub

Sub CompletaInfo(ByRef Planilla As Workbook)
    Dim i As Long
    Dim rutaArchivo As String
    Dim rutaEquivalencia As String

    rutaArchivo = "'" & RUTA & "\..\[Stock.XLS]Sheet1'!"
    rutaEquivalencia = "'" & RUTA & "\..\[Equivalencia.XLS]Sheet1'!"

    
    'Application.ScreenUpdating = False
    With Planilla.Worksheets(1)
        For i = 2 To UltimaTotal
            With .Cells(i, 3)
                .Value = "'" & .Value
                .HorizontalAlignment = xlLeft
                .Font.Color = vbWhite
                .Characters(Start:=1, Length:=7).Font.ColorIndex = xlAutomatic
            End With
            
            If .Cells(i, 3).Value = "" Then
                GoTo Siguiente
            ElseIf Len(.Cells(i, 3).Value) > 7 Then
                
                ' Color
               .Cells(i, 5).Formula = "=IF(ISNA(VLOOKUP(C" & i & ", " & rutaEquivalencia & "$A$2:$F$10000, 4, FALSE)), """", " & _
                       "IF(VLOOKUP(C" & i & ", " & rutaEquivalencia & "$A$2:$F$10000, 4, FALSE)="""", """", " & _
                       "VLOOKUP(C" & i & ", " & rutaEquivalencia & "$A$2:$F$10000, 4, FALSE) & "". "" & VLOOKUP(C" & i & ", " & rutaEquivalencia & "$A$2:$F$10000, 5, FALSE)))"
                
                ' Talle
                .Cells(i, 6).Formula = "=IF(ISNA(VLOOKUP(C" & i & ", " & rutaEquivalencia & "$A$2:$F$10000, 6, FALSE)), """", " & _
                       "IF(VLOOKUP(C" & i & ", " & rutaEquivalencia & "$A$2:$F$10000, 6, FALSE)="""", """", " & _
                       "VLOOKUP(C" & i & ", " & rutaEquivalencia & "$A$2:$F$10000, 6, FALSE)))"
                .Cells(i, 6).HorizontalAlignment = xlCenter
            End If
            
            ' Descripción
            .Cells(i, 4).Formula = "=IFERROR(VLOOKUP(LEFT(C" & i & ", 7), " & rutaArchivo & "$A$2:$C$10000, 2, FALSE), """")"
            
Siguiente:
        Next i
    End With
    
    'Application.ScreenUpdating = True
End Sub



Sub LimpiezaDatos(ByRef Libro As Workbook)
    With Libro.Worksheets(1)
        .Activate
        .Range("A1").Activate
        
        .Name = "Ventas"
 
        .Cells.Replace What:="-CL", Replacement:="", LookAt:=xlPart, _
        searchorder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
        .Cells.Replace What:="-PR", Replacement:="", LookAt:=xlPart, _
        searchorder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
        UltimaTotal = .Cells(.Rows.Count, 3).End(xlUp).Row
        Debug.Print UltimaTotal
    End With
End Sub

Sub FormatoTabla(Archivo As Workbook, Hoja As Worksheet, Orientacion As Boolean)
'Sub FormatoTabla()
' Orientacion => False: Portrait (Vertical)
' Orientacion => True: Landscape (Horizontal)

Dim ultimaFila As Long
Dim i As Integer
    
If Orientacion = True Then ' => VENTAS
    With Hoja
        .Activate
        .Range("A1").Activate
        
        ' Ultima Fila
        ultimaFila = .Cells(.Rows.Count, 3).End(xlUp).Row
        
        ' Bordes tabla
        Range(.Cells(1, 1), .Cells(ultimaFila, 9)).Select
        With Selection
            .Cells.Font.Size = 14
            .Borders(xlInsideVertical).LineStyle = xlContinuous
            .Borders(xlInsideVertical).ColorIndex = 0
            .Borders(xlInsideVertical).TintAndShade = 0
            .Borders(xlInsideVertical).Weight = xlThin
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeRight).LineStyle = xlContinuous
        End With
        Range(.Cells(1, 1), .Cells(1, 9)).Select
        With Selection
            .Borders.LineStyle = xlContinuous
            .EntireRow.HorizontalAlignment = xlCenter
            .Font.Bold = True
            .Interior.Color = RGB(250, 250, 250)
        End With
        
        ' Columna Nº Venta
        .Columns(1).ColumnWidth = 25
        
        ' Columna Cliente
        .Columns(2).ColumnWidth = 38
        .Columns(2).EntireColumn.WrapText = True
        
        ' Columna Producto
        .Columns(4).ColumnWidth = 55
        
        ' Columna Color
        .Columns(5).ColumnWidth = 15
        .Columns(5).EntireColumn.HorizontalAlignment = xlCenter
        .Columns(5).EntireColumn.WrapText = True
        
        ' Columna Talle
        .Columns(6).ColumnWidth = 7
        .Columns(6).EntireColumn.HorizontalAlignment = xlCenter
        
        ' Columna Cantidad
        .Columns(7).ColumnWidth = 7
        
        ' Columna Detalle
        .Columns(8).ColumnWidth = 15
        .Columns(8).EntireColumn.WrapText = True
        
        ' Seguimiento
        .Columns(9).AutoFit
        
        ' Separador de ventas
        For i = 2 To ultimaFila
            If .Cells(i, 1).Value <> "" Then
                Range(.Cells(i, 1), .Cells(i, 9)).Borders(xlEdgeTop).LineStyle = xlContinuous
            End If
        Next i
        
        ' Totales
        .Cells(ultimaFila + 1, 6).Value = "TOTALES:"
        .Cells(ultimaFila + 1, 7).Formula = "=SUM(G2:G" & ultimaFila & ")"
        .Cells(ultimaFila + 1, 6).HorizontalAlignment = xlRight
        .Cells(ultimaFila + 1, 7).WrapText = False
        
        With Range(.Cells(ultimaFila + 1, 6), .Cells(ultimaFila + 1, 7))
            .Font.Bold = True
            .Font.Size = 20
        End With
        
        ' Cantidad de rótulos
        .Cells(ultimaFila + 1, 2).Value = "ROTULOS: "
        .Cells(ultimaFila + 1, 3).Formula = "=COUNTA(A2:A" & ultimaFila & ")"
        With Range(.Cells(ultimaFila + 1, 2), .Cells(ultimaFila + 1, 3))
            .Font.Bold = True
            .Font.Size = 20
        End With
        .Cells(ultimaFila + 1, 2).HorizontalAlignment = xlRight
        .Cells(ultimaFila + 1, 3).HorizontalAlignment = xlLeft
        
        ' Los encabezados
        With .Rows(1)
            .RowHeight = 25
            .HorizontalAlignment = xlCenter
        End With
    End With
    
    ' Formato impresión
    With Hoja.PageSetup
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
        .PrintArea = Hoja.Range("A1:H" & (ultimaFila + 1)).Address
        .Zoom = False
        .FitToPagesTall = 1
        .FitToPagesWide = 1
        .CenterHeader = "&B&20&F"
    End With
    
ElseIf Orientacion = False Then ' => DEPOSITO
    With Hoja
        .Activate
        .Range("A1").Activate
        
        ' Ultima Fila
        ultimaFila = .Cells(.Rows.Count, 1).End(xlUp).Row
        
        ' Bordes tabla
        Range(.Cells(1, 1), .Cells(ultimaFila, 6)).Select
        With Selection
            .Cells.Font.Size = 14
            .Borders.LineStyle = xlContinuous
            .Borders.ColorIndex = 0
            .Borders.TintAndShade = 0
            .Borders.Weight = xlThin
            .Borders.LineStyle = xlContinuous
            .Borders.LineStyle = xlContinuous
            .Borders.LineStyle = xlContinuous
        End With
        Range(.Cells(1, 1), .Cells(1, 6)).Select
        With Selection
            .Borders.LineStyle = xlContinuous
            .EntireRow.HorizontalAlignment = xlCenter
            .Font.Bold = True
            .Interior.Color = RGB(250, 250, 250)
        End With
        
        ' Columna Color
        .Columns(3).ColumnWidth = 15
        .Columns(3).EntireColumn.HorizontalAlignment = xlCenter
        .Columns(3).EntireColumn.WrapText = True
        
        ' Ancho desubicación
        .Columns(6).ColumnWidth = 22
        .Columns(6).EntireColumn.WrapText = True
        
        ' Agranda la letra y redimensiona las columnas
        With .Range("A1").CurrentRegion
            .Font.Size = 14
            .Columns.AutoFit
        End With
        
        ' Aumenta el alto de las filas
        With .Range("A1").CurrentRegion
            .Rows.RowHeight = 20
        End With
        
        ' Totales
        .Cells(ultimaFila + 1, 4).Value = "TOTALES:"
        .Cells(ultimaFila + 1, 5).Formula = "=SUM(E2:E" & ultimaFila & ")"
        .Cells(ultimaFila + 1, 4).HorizontalAlignment = xlRight
        .Cells(ultimaFila + 1, 5).WrapText = False
        
        With Range(.Cells(ultimaFila + 1, 4), .Cells(ultimaFila + 1, 5))
            .Font.Bold = True
            .Font.Size = 20
        End With
        
        ' Los encabezados
        With .Rows(1)
            .RowHeight = 25
            .HorizontalAlignment = xlCenter
        End With
        
    End With
    
    ' Formato impresión
    With Hoja.PageSetup
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
        .PrintArea = Hoja.Range("A1:F" & (ultimaFila + 1)).Address
        .Zoom = False
        .FitToPagesTall = 1
        .FitToPagesWide = 1
        .CenterHeader = "&B&20&F" & vbNewLine & "SOLO PARA USO EN DEPOSITO"
    End With
Else
    ' Muchísimo
End If

End Sub

Sub AB_DepositoMeli()
    Dim Archivo As Workbook
    Dim Hoja As Worksheet
    Dim ultimaFila As Long
    Dim i As Integer
    Dim rutaEquivalencia As String
    Dim server As String
    Dim carpetaDestino As String
    Dim nombreArchivo As String
        
    Set Archivo = ActiveWorkbook
    rutaEquivalencia = Archivo.Path
    rutaEquivalencia = "'" & rutaEquivalencia & "\..\[Stock.XLS]Sheet1'!"
    nombreArchivo = Len(ActiveWorkbook.Name)
    server = "\\SER-DF\D\A Remitar TXT"
    carpetaDestino = "\MELI1\"
    Debug.Print carpetaDestino & vbNewLine & server & vbNewLine & nombreArchivo
    
    With Archivo.Worksheets("Ventas")
        ' Ultima Fila: la toma de Ventas,
        ' sin tener en cuenta la celda de los totales
        ultimaFila = .Cells(.Rows.Count, 3).End(xlUp).Row - 1
    End With
    
    ' Asigna el objeto Deposito =======
    
    ' Desactivar alertas para eliminar la hoja sin confirmación previa
    Application.DisplayAlerts = False
    
    ' Comprobar si la hoja ya existe y eliminarla
    On Error Resume Next
    Set Hoja = Archivo.Worksheets("Depósito")
    On Error GoTo 0
    
    If Not Hoja Is Nothing Then
        Hoja.Delete
    End If
    
    ' Restaurar alertas
    Application.DisplayAlerts = True
    
    ' Agregar la nueva hoja al final del libro
    Set Hoja = Archivo.Worksheets.Add(After:=Archivo.Worksheets(Archivo.Worksheets.Count))
    Hoja.Name = "Depósito"
    
    ' TITULARES
    With Hoja
        .Cells(1, 1).Value = "Código"
        .Cells(1, 2).Value = "Producto"
        .Cells(1, 3).Value = "Color"
        .Cells(1, 4).Value = "Talle"
        .Cells(1, 5).Value = "Cant."
        .Cells(1, 6).Value = "Ubicación"
    End With
    
    ' Copia y pega información
    With Archivo.Worksheets("Ventas")
        For i = 2 To ultimaFila
            ' Código
            Hoja.Cells(i, 1).Value = "'" & Left(.Cells(i, 3).Value, 7)
            
            ' Producto
            Hoja.Cells(i, 2).Value = .Cells(i, 4).Value
            
            ' Color
            Hoja.Cells(i, 3).Value = .Cells(i, 5).Value
            
            ' Talle
            Hoja.Cells(i, 4).Value = .Cells(i, 6).Value
            
            ' Cantidad
            Hoja.Cells(i, 5).Value = .Cells(i, 7).Value
            
            ' Desubicación
            Hoja.Cells(i, 6).Formula = "=IFERROR(VLOOKUP(A" & i & ", " & rutaEquivalencia & "$A$2:$C$10000, 3, FALSE), """")"
            
            ' Otros depósitos
            If LCase(.Cells(i, 8).Value) = "planta baja" Or LCase(.Cells(i, 8).Value) = "mendoza" Then
                Hoja.Cells(i, 6).Value = "12. Planta Baja"
                .Cells(i, 8).Value = "Planta Baja"
            ElseIf LCase(.Cells(i, 8).Value) = "tucuman" Or LCase(.Cells(i, 8).Value) = "neuquen" Or LCase(.Cells(i, 8).Value) = "rgallegos" Then
                Hoja.Range(Cells(i, 1), Cells(i, 5)).ClearContents
                Hoja.Cells(i, 6).Value = .Cells(i, 8).Value
            End If
            
        Next i
    End With
    
    ' Ordenando alfabéticamente la última columna de ubicación
    With Hoja.Range("A1:F1")
        .AutoFilter
        .Rows("1").RowHeight = 27
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    Worksheets("Depósito").Range("A1").CurrentRegion.Sort Key1:=Range("F1"), Order1:=xlAscending, Header:=xlGuess
    With Selection
        .AutoFilter
    End With
    
    ' Centra los datos
    With Hoja.Range(Cells(1, 3), Cells(ultimaFila, 5))
        .HorizontalAlignment = xlCenter
    End With
    
    ' Coloreando celdas
    For i = 2 To ultimaFila
        Call PintaFila(Hoja.Name, i, 1, 6)
    Next i
    
    ' Dar formato
    Call FormatoTabla(Archivo, Hoja, False)
    
    ' Crear el txt
    Call generarTxt(i, ultimaFila, "", 1, nombreArchivo, carpetaDestino, ultimaFila, 0, server)
    
End Sub
Function PintaFila(Hojilla As String, fila As Integer, DesdeColumna As Integer, HastaColumna As Integer)
    ' Pinta filas impares
    If fila Mod 2 <> 0 Then
        Worksheets(Hojilla).Range(Cells(fila, DesdeColumna), Cells(fila, HastaColumna)).Interior.Color = RGB(240, 240, 240)
    End If
End Function

Function generarTxt(fila, ultimaFila, textoArchivo, cantArchivos, nombreArchivo, carpetaDestino, limite, resto, server)
    Dim rutaArchivo As String
    Dim rutaCarpeta As String
    Dim i As Integer
    Dim tope As Long ' Cambiado a Long para evitar desbordamientos
    Dim Archivo As Workbook
    Dim CantLetras As Integer ' Cambiado a Integer para evitar desbordamiento si el texto es largo
    Dim fso As Object
    
    fila = 0
    Set Archivo = ActiveWorkbook

    ' Instanciar FileSystemObject para verificar la red/carpeta
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Definir y validar la ruta del directorio
    rutaCarpeta = server & carpetaDestino
    
    ' Asegurar que la ruta termine en barra diagonal '\'
    If Right(rutaCarpeta, 1) <> "\" Then rutaCarpeta = rutaCarpeta & "\"

    ' Verificar si la carpeta compartida o servidor está disponible
    If Not fso.FolderExists(rutaCarpeta) Then
        MsgBox "Atención: La carpeta de red no está disponible o accesible:" & vbCrLf & _
               rutaCarpeta & vbCrLf & vbCrLf & _
               "Se omitirá la generación del archivo .txt y la macro continuará.", _
               vbExclamation, "Servidor / Red No Disponible"
        Exit Function
    End If

    ' Generación del txt
    For i = 1 To cantArchivos
        tope = i * limite
        If i = cantArchivos Then
            tope = ultimaFila
        End If
        
        With Archivo.Worksheets("Depósito")
            For fila = (limite * (i - 1)) + 1 To tope - 1
            
                CantLetras = InStr(1, .Cells(fila + 1, 3).Value, ".")
                
                If CantLetras <= 0 Then
                    ' Si no encuentra punto, toma la longitud total del texto
                    CantLetras = Len(.Cells(fila + 1, 3).Value) + 1
                End If
                
                ' 1º. Cantidad (+) 2º. Código (!) 3º. Color (!) 4º. Talle
                textoArchivo = textoArchivo _
                    & .Cells(fila + 1, 5).Value _
                    & "+" & .Cells(fila + 1, 1).Value _
                    & "!" & Left(.Cells(fila + 1, 3).Value, CantLetras - 1) _
                    & "!" & .Cells(fila + 1, 4).Value _
                    & vbNewLine
            Next fila
        End With
        
        ' Asignación incremental de nombre
        If cantArchivos > 1 Then
            nombreArchivo = Left(ActiveWorkbook.Name, Len(ActiveWorkbook.Name) - 5) & " - " & i & ".txt"
        Else
            nombreArchivo = Left(ActiveWorkbook.Name, Len(ActiveWorkbook.Name) - 5) & ".txt"
        End If
        
        rutaArchivo = rutaCarpeta & nombreArchivo
        
        ' Captura de errores durante la apertura/escritura del archivo
        On Error Resume Next
        
        Open rutaArchivo For Output As #1
        Print #1, textoArchivo
        Close #1
        
        ' Verificar si hubo algún error de escritura/permisos al guardar
        If Err.Number <> 0 Then
            MsgBox "Ocurrió un error al intentar guardar el archivo en la red:" & vbCrLf & _
                   rutaArchivo & vbCrLf & _
                   "Detalle del error: " & Err.Description, vbExclamation, "Error de Exportación"
            Err.Clear
        Else
            MsgBox "Datos exportados con éxito a " & rutaArchivo, vbInformation, "Cargar detalle desde txt"
        End If
        
        On Error GoTo 0
        
        textoArchivo = ""
    Next i
End Function

Sub ConstructorPlanilla()
    Dim FilaFinal As Long
    Dim FilaInicial As Long
    Dim FilasRestantes As Long
    Dim NumPlanilla As Byte
    Dim PlanillaOriginal As Workbook
    Set PlanillaOriginal = ActiveWorkbook
    With PlanillaOriginal.Worksheets(1)
        NumPlanilla = 1
        FilaInicial = 1
        FilaFinal = 20
        UltimaTotal = .Cells(.Rows.Count, 3).End(xlUp).Row
        Debug.Print "Ultima fila total: " & UltimaTotal
        FilasRestantes = UltimaTotal - FilaInicial
        
        Do While UltimaTotal > FilaInicial + 20
            .Cells(FilaFinal, 1).Activate
            If UltimaTotal >= (FilaFinal - FilaInicial + 1) Then
                Do While .Cells(FilaFinal, 1).Value = ""
                    FilaFinal = FilaFinal - 1
                    .Cells(FilaFinal, 1).Activate
                Loop
                FilaFinal = FilaFinal - 1
                .Cells(FilaFinal, 1).Activate
            End If
            
            FilasRestantes = UltimaTotal - FilaFinal
            Debug.Print "Fila inicial de la planilla Nº " & NumPlanilla & ": " & .Cells(FilaInicial, 1).Row
            Debug.Print "Fila final de la planilla Nº " & NumPlanilla & ": " & .Cells(FilaFinal, 1).Row
            Debug.Print "Cayó en un carrito. Tiene que ser otra anterior..."
            Debug.Print "Filas restantes: " & FilasRestantes
            
            ' Se determina la ultima fila en la planilla que se procesa
            FilaFinal = ActiveCell.Row
            Debug.Print "Fila final de la planilla Nº " & NumPlanilla & ": " & FilaFinal
            
            ' Incrementamos la planilla
            FilaInicial = FilaFinal + 1
            NumPlanilla = NumPlanilla + 1
            FilasRestantes = FilasRestantes - FilaFinal
            FilaFinal = FilaFinal + 20
        Loop
    End With
End Sub
