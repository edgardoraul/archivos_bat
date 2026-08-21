Attribute VB_Name = "DJ"
Option Explicit

' Constante para definir el ancho de la impresión en centímetros
Const ANCHO_IMPRESION_CM As Double = 12

Sub GenerarRotuloCD()
    Dim fd As FileDialog
    Dim folderPath As String, folderName As String, genrePart As String
    Dim fileName As String, cleanName As String, savePath As String
    Dim colorVal As Long
    Dim ws As Worksheet
    Dim r As Long, i As Long, lastSongRow As Long, endRow As Long, dotPos As Long
    Dim ptsTarget As Double, ptsA As Double, ptsC As Double, ptsBTarget As Double
    Dim rngLabel As Range
    Dim shpFrame As Shape, shp As Shape

    ' 1. Seleccionar carpeta mediante cuadro de diálogo
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    fd.Title = "Seleccione la carpeta del CD"
    If fd.Show <> -1 Then Exit Sub
    folderPath = fd.SelectedItems(1)
    
    folderName = Mid(folderPath, InStrRev(folderPath, "\") + 1)
    
    ' 2. Evaluar tipo de música a partir del 6º caracter
    If Len(folderName) >= 6 Then
        genrePart = UCase(Mid(folderName, 6))
    Else
        genrePart = UCase(folderName)
    End If
    
    ' Asignación de COLOR
    If InStr(genrePart, "DEEP HOUSE") > 0 Then
        colorVal = RGB(0, 176, 80)      ' Verde
        
    ElseIf InStr(genrePart, "PROGRESSIVE HOUSE") > 0 Then
        colorVal = RGB(255, 127, 0)    ' Naranja
        
    ElseIf InStr(genrePart, "PROGRESSIVE") > 0 Then
        colorVal = RGB(20, 20, 255)      ' Azul
        
    ElseIf InStr(genrePart, "TRANCE") > 0 Then
        colorVal = RGB(255, 20, 255)    ' Magenta
        
    ElseIf InStr(genrePart, "ACID JAZZ") > 0 Then
        colorVal = RGB(0, 0, 0)        ' Negro
        
    ElseIf InStr(genrePart, "HOUSE") > 0 Then
        colorVal = RGB(139, 69, 19)    ' Marrón
        
    Else
        colorVal = RGB(128, 128, 128)  ' Gris
    End If

    Set ws = ActiveSheet
    ws.Cells.Clear
    
    ' Limpiar formas preexistentes
    For Each shp In ws.Shapes
        shp.Delete
    Next shp

    ' 3. Fuente por defecto
    With ws.Cells.Font
        .Name = "Consolas"
        .Size = 14
    End With

    ' 4. Ancho de columnas A y C = 3, y ajuste dinámico de columna B
    ws.Columns("A").ColumnWidth = 3
    ws.Columns("C").ColumnWidth = 3
    
    ptsTarget = Application.CentimetersToPoints(ANCHO_IMPRESION_CM)
    ptsA = ws.Columns("A").Width
    ptsC = ws.Columns("C").Width
    ptsBTarget = ptsTarget - ptsA - ptsC
    
    If ptsBTarget > 0 Then
        ws.Columns("B").ColumnWidth = 10
        ws.Columns("B").ColumnWidth = ws.Columns("B").ColumnWidth * (ptsBTarget / ws.Columns("B").Width)
    End If

    ' 5. Título en B2
    With ws.Range("B2")
        .Value = folderName
        .Font.Name = "Arial"
        .Font.Size = 20
        .Font.Bold = True
        .Font.color = colorVal
        .HorizontalAlignment = xlCenter
        .WrapText = True
    End With

    ' 6. Listado de archivos a partir de B4 (sin extensión)
    r = 4
    fileName = Dir(folderPath & "\*.*")
    
    Do While fileName <> ""
        If Left(fileName, 2) <> "~$" And UCase(fileName) <> UCase(folderName & ".xlsm") And UCase(fileName) <> UCase(folderName & ".xlsx") Then
            
            dotPos = InStrRev(fileName, ".")
            If dotPos > 1 Then
                cleanName = Left(fileName, dotPos - 1)
            Else
                cleanName = fileName
            End If
            
            With ws.Cells(r, 2)
                .Value = cleanName
                .WrapText = True
                
                If Len(cleanName) >= 3 Then
                    .Characters(1, 3).Font.color = colorVal
                Else
                    .Characters(1, Len(cleanName)).Font.color = colorVal
                End If
                
                With .Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .color = RGB(210, 210, 210)
                    .Weight = xlThin
                End With
            End With
            
            r = r + 1
        End If
        fileName = Dir
    Loop

    lastSongRow = r - 1

    ' Quitar borde inferior al último archivo
    If lastSongRow >= 4 Then
        ws.Cells(lastSongRow, 2).Borders(xlEdgeBottom).LineStyle = xlNone
    End If

    ' Ajuste de altura de filas de archivos (mínimo 25)
    For i = 4 To lastSongRow
        ws.Rows(i).AutoFit
        If ws.Rows(i).RowHeight < 25 Then ws.Rows(i).RowHeight = 25
    Next i

    ' Fila límite del marco (última con archivo + 2)
    If lastSongRow >= 4 Then
        endRow = lastSongRow + 2
    Else
        endRow = 5
    End If

    ws.Rows(endRow - 1).RowHeight = 15
    ws.Rows(endRow).RowHeight = 15
    ws.Rows(1).RowHeight = 12
    ws.Rows(3).RowHeight = 12

    ' 7. Insertar marco rectangular con esquinas inferiores redondeadas
    Set rngLabel = ws.Range(ws.Cells(1, 1), ws.Cells(endRow, 3))
    
    Set shpFrame = ws.Shapes.AddShape(msoShapeRound2SameRectangle, _
                                      rngLabel.Left, rngLabel.Top, _
                                      rngLabel.Width, rngLabel.Height)
    With shpFrame
        .Rotation = 180 ' Invierte la forma para que las curvas queden abajo
        .Fill.Visible = msoFalse
        With .Line
            .DashStyle = msoLineDash
            .ForeColor.RGB = colorVal
            .Weight = 1
        End With
    End With

    ' 8. Configurar impresión con márgenes en 0
    With ws.PageSetup
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
        .LeftMargin = 0
        .RightMargin = 0
        .TopMargin = 0
        .BottomMargin = 0
        .PrintArea = rngLabel.Address
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .Zoom = False
    End With

    ' 9. Guardar libro como .xlsx
    savePath = folderPath & "\" & folderName & ".xlsx"
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs fileName:=savePath, FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True

    MsgBox "Rótulo generado correctamente en:" & vbCrLf & savePath, vbInformation, "Listo"
End Sub

