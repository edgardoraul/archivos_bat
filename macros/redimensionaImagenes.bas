Attribute VB_Name = "Módulo1"

Public Sub FitPic()
Attribute FitPic.VB_ProcData.VB_Invoke_Func = "h\n14"
'https://www.extendoffice.com/documents/excel/1060-excel-resize-picture-to-fit-cell.html
    On Error GoTo NOT_SHAPE
    Dim PicWtoHRatio As Single
    Dim CellWtoHRatio As Single
    With Selection
        PicWtoHRatio = .Width / .Height
    End With
    With Selection.TopLeftCell
        CellWtoHRatio = .Width / .RowHeight
    End With
    Select Case PicWtoHRatio / CellWtoHRatio
    Case Is > 1
        With Selection
            .Width = .TopLeftCell.Width - 4
            .Height = .Width / PicWtoHRatio - 4
        End With
    Case Else
        With Selection
            .Height = .TopLeftCell.RowHeight - 4
            .Width = .Height * PicWtoHRatio - 4
        End With
    End Select
    With Selection
        .Top = .TopLeftCell.Top + 4
        .Left = .TopLeftCell.Left + 4
    End With
    Exit Sub
NOT_SHAPE:
    MsgBox "Select a picture before running this macro."
End Sub


