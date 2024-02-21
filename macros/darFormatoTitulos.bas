Attribute VB_Name = "Módulo5"
Sub D_DarFormato()
Attribute D_DarFormato.VB_Description = "DA FORMATO CON BORDES Y TITULOS"
Attribute D_DarFormato.VB_ProcData.VB_Invoke_Func = " \n14"
'
' DARFORMATO Macro
' DA FORMATO CON BORDES Y TITULOS
'

'
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Range(Selection, Selection.End(xlUp)).Select
    Selection.Font.Bold = False
    ActiveWorkbook.Save
    Selection.End(xlToRight).Select
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Columns.AutoFit
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Columns("D:D").ColumnWidth = 39
    Columns("C:C").ColumnWidth = 26.71
    Columns("J:J").ColumnWidth = 12.86
    Columns("I:I").ColumnWidth = 11.86
    Selection.Rows.AutoFit
    Range("A1").Select
    ActiveWorkbook.Save
End Sub
