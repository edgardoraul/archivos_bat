Attribute VB_Name = "Módulo4"
Sub A_EliminarColumas()
Attribute A_EliminarColumas.VB_Description = "Elimina las columnas de las planillas para procesar las ventas."
Attribute A_EliminarColumas.VB_ProcData.VB_Invoke_Func = " \n14"
'
' EliminarColumas Macro
' Elimina las columnas de las planillas para procesar las ventas.
'

'
    Columns("B:B").Select
    Selection.Delete Shift:=xlToLeft
    Columns("D:H").Select
    Selection.Delete Shift:=xlToLeft
    Columns("E:E").Select
    Selection.Delete Shift:=xlToLeft
    Columns("F:F").Select
    Selection.Delete Shift:=xlToLeft
    Columns("I:AL").Select
    Selection.Delete Shift:=xlToLeft
    Columns("J:J").Select
    Selection.Delete Shift:=xlToLeft
    Columns("K:K").Select
    Selection.Delete Shift:=xlToLeft
    Columns("K:K").Select
    Selection.Delete Shift:=xlToLeft
    Range("J1").Select
    ActiveCell.FormulaR1C1 = "Firma Control"
    Range("I1").Select
    ActiveCell.FormulaR1C1 = "Firma Recepción"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Id Venta"
    Range("A1").Select
    ActiveWorkbook.Save
End Sub
