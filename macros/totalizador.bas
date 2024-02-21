Attribute VB_Name = "Módulo3"
Sub H_Totalizador()
Attribute H_Totalizador.VB_ProcData.VB_Invoke_Func = "t\n14"
'
' totalizador Macro
'
' Acceso directo: CTRL+t
'
    ActiveCell.Offset(-1, -1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "TOTAL"
    ActiveCell.Range("A1:B1").Select
    ActiveCell.Offset(0, 1).Range("A1").Activate
    Selection.Font.Bold = True
    With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    ActiveWorkbook.Save
End Sub
