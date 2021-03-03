Attribute VB_Name = "Módulo7"
Sub J_TotalizarRotulos()
Attribute J_TotalizarRotulos.VB_ProcData.VB_Invoke_Func = "r\n14"
'
' TotalizarRotulos Macro
'
' Acceso directo: CTRL+r
'
    ActiveCell.FormulaR1C1 = "ROTULOS"
    ActiveCell.Offset(0, -1).Range("A1:B1").Select
    ActiveCell.Activate
    Selection.Font.Bold = True
    With Selection.Font
        .Name = "Arial"
        .Size = 15
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    ActiveCell.Offset(0, 0).Range("A1").Select
    ActiveWorkbook.Save
End Sub
