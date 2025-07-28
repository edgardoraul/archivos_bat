Attribute VB_Name = "RotuladorDetalle"
Option Explicit
Public Const pass As String = "Rerda2025"
Public ROTULO As Worksheet
Public ultima As Integer

Sub Proteger()
' ACTIVA Y PROTEGE LAS HOJAS =========
    Set ROTULO = ThisWorkbook.Worksheets("ROTULO")
    Dim hojita As Worksheet
    For Each hojita In Application.Worksheets
        hojita.Protect Password:=pass
    Next hojita
    ThisWorkbook.Protect Password:=pass
End Sub

Sub Desproteger()
' DESACTIVA Y DESPROTEGE LAS HOJAS ====
    Dim hojita As Worksheet
        For Each hojita In Application.Worksheets
            hojita.Unprotect Password:=pass
        Next hojita
    ThisWorkbook.Unprotect Password:=pass
End Sub

Function Alertar()
' ACTIVA LAS ALERTAS ==================
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Function

Function Desalertar()
' DESACTIVA LAS ALERTAS ===============
    Application.EnableEvents = False
    Application.ScreenUpdating = False
End Function

Function UltimaFila()
    ultima = ROTULO.Cells(Rows.Count, 6).End(xlUp).Row
    If ultima < 7 Then
        ultima = 7
    End If
End Function

Sub AddVariante()
' AGREGA UNA FILA DE VARIANTES ===============
    ROTULO.Unprotect pass
    With ROTULO
        .Rows(7).Insert Shift:=xlShiftDown, CopyOrigin:=xlFormatFromRightOrBelow
        .Range(Cells(8, 1), Cells(8, 6)).Select
    End With
    Selection.Copy
    ROTULO.Range("A7").Activate
    ActiveSheet.Paste
    Application.CutCopyMode = False
    With ROTULO
        .Range(Cells(7, 1), Cells(7, 5)).Value = ""
        .Cells(7, 6).Value = " "
    End With
    Call UltimaFila
    ROTULO.Protect pass
    Call UltimaFila
    Debug.Print "Se agregó la fila " & ultima
End Sub

Sub DelVariante()
' REMUEVE UNA FILA DE VARIANTES ==============
    Call UltimaFila
    If ultima = 7 Then
        MsgBox "No se puede borrar la última fila"
        Exit Sub
    End If
    With ROTULO
        .Unprotect pass
        .Rows(7).EntireRow.Delete
        .Protect pass
    End With
End Sub
