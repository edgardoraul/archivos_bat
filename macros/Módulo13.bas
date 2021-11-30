Attribute VB_Name = "Módulo5"
Sub Colorear()
Attribute Colorear.VB_Description = "PInta las celdas o filas...bla bla de un color para dejar controladas."
Attribute Colorear.VB_ProcData.VB_Invoke_Func = "o\n14"
'
' Colorear Macro
' PInta las celdas o filas...bla bla de un color para dejar controladas.
'
' Acceso directo: CTRL+o
'
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    ActiveWorkbook.Save
End Sub
