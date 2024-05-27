Attribute VB_Name = "InsertarFilasSuperior"
Sub InsertarFilasSuperior()
Attribute InsertarFilasSuperior.VB_ProcData.VB_Invoke_Func = "f\n14"
'
' InsertarFilasSuperior Macro
'
' Acceso directo: CTRL+f
'
    Rows("3:3").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A2:E2").Select
    Range("E2").Activate
    Selection.Copy
    Range("A3").Select
    ActiveSheet.Paste
    Range("A2:E2").Select
    Range("E2").Activate
    Application.CutCopyMode = False
    Selection.ClearContents
    Range("A2").Select
    
    Range("C3:E3").Select
    Selection.AutoFill Destination:=Range("C2:E3"), Type:=xlFillDefault
    Range("C2:E3").Select
    Range("B2").Select
End Sub
