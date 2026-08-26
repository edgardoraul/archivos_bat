Attribute VB_Name = "MeliNueva"
Option Explicit

Public RUTA As String
Public Const MIPC As String = "D:\Web\Listados de Ventas Online\MELY"
Public Const MIPCENRED As String = "\\EDGAR\Web\Listados de Ventas Online\MELY"
Public UltimaTotal As Long
Public IdCliente As String
Public Provincias As Variant


Sub MeliNueva()
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

Public Sub GuardarCopiaConSecuencial()
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
    
    ' 4.1 Particionar la planilla
    ' Como máximo se hará una planilla con 20 renglones, no más.
    ' Pero si justo cae en medio de un carrito de compras, pues pasa a la
    ' planilla siguiente.
    Call ConstructorPlantillas
    

    ' 5. Determinar fecha
    todayStr = Format(Date, "yyyy-mm-dd")
    filePrefix = todayStr & ". "

    ' 6. Analizar archivos correlativos
    maxNum = 0
    Set folderObj = fso.GetFolder(RUTA)

    For Each fileObj In folderObj.Files
        fileName = fileObj.Name
        If Left(fileName, Len(filePrefix)) = filePrefix Then
            posDot = InStrRev(fileName, ".")
            If posDot > 0 Then fileName = Left(fileName, posDot - 1)
            
            If Mid(fileName, Len(filePrefix) + 6, 8) = " - MELI" Then
                numStr = Mid(fileName, Len(filePrefix) + 1, 5)
                If IsNumeric(numStr) Then
                    currentNum = CLng(numStr)
                    If currentNum > maxNum Then maxNum = currentNum
                End If
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
        
    If Not fso.FileExists(RUTA & "\..\Equivalencia2.XLS") Then
        MsgBox "El archivo Equivalencia2.XLS debe estar en la carpeta " & vbNewLine & RUTA, vbCritical, "Error"
        wbTemp.Close SaveChanges:=False
        Exit Sub
    End If
    
    ' 9. Completado de información pasando el libro
    Call CompletaInfo(wbTemp)
    
    ' 10. Formato a Ventas
    'Call FormatoTabla(wbTemp, wbTemp.Worksheets("Ventas"), True)
    
    ' 11. Crear pestaña para Depósito
    'Call Deposito(wbTemp, wbTemp.Worksheets("Deposito"))
    
    ' 12. Formato a Depósito
    'Call FormatoTabla(wbTemp, wbTemp.Worksheets("Deposito"), False)
    
    ' 13. Generar TXT de importación
    Call TxtImportacion(wbTemp)
    
    ' 14. Guardar y cerrar la copia procesada
    wbTemp.SaveAs fileName:=finalPath, FileFormat:=51
    
    MsgBox "Proceso completado con éxito:" & vbCrLf & finalPath, vbInformation, "Éxito"
End Sub

Sub CompletaInfo(ByRef Planilla As Workbook)
    Dim i As Long
    Dim rutaArchivo As String
    Dim rutaEquivalencia As String

    rutaArchivo = "'" & RUTA & "\..\[Stock.XLS]Sheet1'!"
    rutaEquivalencia = "'" & RUTA & "\..\[Equivalencia2.XLS]Sheet1'!"
    
    'Application.ScreenUpdating = False
    With Planilla.Worksheets(1)
        For i = 2 To UltimaTotal
            With .Cells(i, 3)
                .Value = "'" & .Value
                .HorizontalAlignment = xlLeft
                .Font.color = vbWhite
                .Characters(Start:=1, Length:=7).Font.ColorIndex = xlAutomatic
            End With
            
            If .Cells(i, 3).Value = "" Then
                GoTo Siguiente
            ElseIf Len(.Cells(i, 3).Value) > 7 Then
                
                ' Color
                .Cells(i, 5).Formula = "=IF(ISNA(VLOOKUP(C" & i & ", " & rutaEquivalencia & "$A$2:$F$10000, 4, FALSE)), ""Sin equivalencia"", IF(VLOOKUP(C" & i & ", " & rutaEquivalencia & "$A$2:$F$10000, 4, FALSE)="""", """", VLOOKUP(C" & i & ", " & rutaEquivalencia & "$A$2:$F$10000, 4, FALSE)))"
                .Cells(i, 5).HorizontalAlignment = xlCenter
                
                ' Leyenda
                '.Cells(i, 6).Formula = "=IF(ISNA(VLOOKUP(C" & i & ", " & rutaEquivalencia & "$A$2:$F$10000, 5, FALSE)), ""Sin equivalencia"", IF(VLOOKUP(C" & i & ", " & rutaEquivalencia & "$A$2:$F$10000, 5, FALSE)="""", """", VLOOKUP(C" & i & ", " & rutaEquivalencia & "$A$2:$F$10000, 5, FALSE)))"
                '.Cells(i, 6).HorizontalAlignment = xlCenter
                
                ' Talle
                .Cells(i, 6).Formula = "=IF(ISNA(VLOOKUP(C" & i & ", " & rutaEquivalencia & "$A$2:$F$10000, 6, FALSE)), ""Sin equivalencia"", IF(VLOOKUP(C" & i & ", " & rutaEquivalencia & "$A$2:$F$10000, 6, FALSE)="""", """", VLOOKUP(C" & i & ", " & rutaEquivalencia & "$A$2:$F$10000, 6, FALSE)))"
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
        .Name = "Ventas"

        
        .Cells.Replace What:="-CL", Replacement:="", LookAt:=xlPart, _
        searchorder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
        .Cells.Replace What:="-PR", Replacement:="", LookAt:=xlPart, _
        searchorder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
        UltimaTotal = .Cells(.Rows.Count, 1).End(xlUp).Row
    End With
End Sub

Sub Deposito(Archivo As Workbook, Hoja As Worksheet)

End Sub
'Sub FormatoTabla(Archivo As Workbook, Hoja As Worksheet, Orientacion As Boolean)
Sub FormatoTabla()
' Orientacion => True: Portrait (Vertical)
' Orientacion => False: Landscape (Horizontal)

' Sólo para testing
Dim Orientacion As Boolean
Dim Archivo As Workbook
Dim Hoja As Worksheet
Dim UltimaFila As Long
Dim i As Byte

Orientacion = True
Set Archivo = ActiveWorkbook
Set Hoja = Archivo.Worksheets("Ventas")
    
    
If Orientacion = True Then ' => VENTAS
    With Hoja
        ' Ultima Fila
        UltimaFila = .Cells(.Rows.Count, 1).End(xlUp).Row
        
        ' Bordes tabla
        Range(.Cells(1, 1), .Cells(UltimaFila, 9)).Select
        With Selection
            .Cells.Font.Name = "Consolas"
            .Cells.Font.Size = 14
            .Borders(xlInsideVertical).LineStyle = xlContinuous
            .Borders(xlInsideVertical).ColorIndex = 0
            .Borders(xlInsideVertical).TintAndShade = 0
            .Borders(xlInsideVertical).Weight = xlThin
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeRight).LineStyle = xlContinuous
        End With
        Range(.Cells(1, 1), .Cells(1, 10)).Select
        With Selection
            .Borders.LineStyle = xlContinuous
            .EntireRow.HorizontalAlignment = xlCenter
            .Font.Bold = True
        End With
        
        ' Columna Nº Venta
        .Columns(1).ColumnWidth = 25
        
        ' Columna Cliente
        .Columns(2).ColumnWidth = 38
        .Columns(2).EntireColumn.WrapText = True
        
        ' Columna Producto
        .Columns(4).ColumnWidth = 55
        
        ' Columna Color
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
        For i = 2 To UltimaFila
            If .Cells(i, 1).Value <> "" Then
                Range(.Cells(i, 1), .Cells(i, 10)).Borders(xlEdgeTop).LineStyle = xlContinuous
            End If
        Next i
        
        ' Totales
        .Cells(UltimaFila + 1, 6).Value = "TOTALES:"
        .Cells(UltimaFila + 1, 7).Formula = "=SUM(G2:G" & UltimaFila & ")"
        .Cells(UltimaFila + 1, 6).HorizontalAlignment = xlRight
        .Cells(UltimaFila + 1, 7).WrapText = False
        
        With Range(.Cells(UltimaFila + 1, 6), .Cells(UltimaFila + 1, 7))
            .Font.Name = "Arial"
            .Font.Bold = True
            .Font.Size = 20
        End With
        
        ' Cantidad de rótulos
        .Cells(UltimaFila + 1, 2).Value = "ROTULOS: "
        .Cells(UltimaFila + 1, 3).Formula = "=COUNTA(A2:A" & UltimaFila & ")"
        With Range(.Cells(UltimaFila + 1, 2), .Cells(UltimaFila + 1, 3))
            .Font.Name = "Arial"
            .Font.Bold = True
            .Font.Size = 20
        End With
        .Cells(UltimaFila + 1, 2).HorizontalAlignment = xlRight
        .Cells(UltimaFila + 1, 3).HorizontalAlignment = xlLeft
        
        ' Los encabezados
        With .Rows(1)
            .RowHeight = 25
            .HorizontalAlignment = xlCenter
            .Font.Bold = True
            .Interior.color = RGB(250, 250, 250)
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
        .PrintArea = Hoja.Range("A1:H" & (UltimaFila + 1)).Address
        .Zoom = False
        .FitToPagesTall = 1
        .FitToPagesWide = 1
        .CenterHeader = "&B&20&F"
    End With
    

Else '=> DEPOSITO
    
    With Archivo.Worksheets("Ventas")
        ' Ultima Fila la toma de Ventas
        UltimaFila = .Cells(.Rows.Count, 1).End(xlUp).Row
        
        ' Obtiene información de ventas
        Range(.Cells(1, 3), .Cells(UltimaFila, 8)).Copy

    End With
    
    ' Asigna el objeto Deposito
    Set Hoja = Archivo.Worksheets("Deposito")
    
    With Hoja
        
        .Activate
        .Cells.Clear
        .Cells.Select
        .Cells.ClearFormats
        .Cells.Font.Size = 12
        .Cells.Font.Name = "Arial"
        .Cells.RowHeight = 17
        .Cells(1, 1).Value = "Código"
        .Cells(1, 2).Value = "Descripción"
        .Cells(1, 3).Value = "Color"
        .Cells(1, 4).Value = "Talle"
        .Cells(1, 5).Value = "Cant"
        .Cells(1, 6).Value = "Ubicación"
        Range("A1").Paste
        
    End With
End If

End Sub
Sub TxtImportacion(Archivo)

End Sub

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
        UltimaTotal = .Cells(.Rows.Count, 1).End(xlUp).Row
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
