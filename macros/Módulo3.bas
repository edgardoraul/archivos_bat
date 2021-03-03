Attribute VB_Name = "Módulo6"
Sub C_QuitarRelleno()
Attribute C_QuitarRelleno.VB_Description = "QUITA EL RELLENO"
Attribute C_QuitarRelleno.VB_ProcData.VB_Invoke_Func = " \n14"
'
' QUITAR_RELLENO Macro
' QUITA EL RELLENO
'

'
    Cells.Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("A1").Select
    ActiveWorkbook.Save
End Sub
