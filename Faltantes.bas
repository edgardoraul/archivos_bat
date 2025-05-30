Attribute VB_Name = "Faltantes"
Public Const pass As String = "Rerda2025"
Public ultimaConDatos As Integer
Public ultimaDerecha As Integer
Public Const naranja As String = 40
Option Explicit


' PROTEGER
Function proteger()
    Dim ws As Worksheet
    Sheets("CODIGOS").Visible = False
    Sheets("VARIANTES").Visible = False
    For Each ws In ThisWorkbook.Worksheets
        ws.Protect pass
    Next ws
    ThisWorkbook.Protect pass
End Function

' DESPROTEGER
Function desproteger()
    ThisWorkbook.Unprotect pass
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ws.Unprotect pass
    Next
End Function

' Advertencia de borrado
Function Advertencia() As Boolean
    Dim Respuesta As VbMsgBoxResult ' Usar VbMsgBoxResult es más claro para los resultados de MsgBox

    ' Mostrar el cuadro de diálogo con Aceptar y Cancelar
    Respuesta = MsgBox("¿En serio?" & vbNewLine & "Vas a borrar la información cargada en esta fila o columna.", vbOKCancel + vbQuestion, "Confirmación")

    ' Analizar la respuesta
    If Respuesta = vbOK Then ' El usuario ha seleccionado "Aceptar"
    
        ' Opcional: MsgBox "Has seleccionado Aceptar. El proceso de borrado continuará.", vbInformation
        Advertencia = True ' Indica que el proceso de borrado puede continuar
        
    ElseIf Respuesta = vbCancel Then ' El usuario ha seleccionado "Cancelar"
    
        ' Opcional: MsgBox "Has seleccionado Cancelar. El proceso de borrado se detendrá.", vbInformation
        Advertencia = False ' Indica que el proceso de borrado DEBE detenerse
        
    End If
End Function

' Generando las últimas celdas con datos.
Function ultima(ultimaConDatos)
    ultimaConDatos = Cells(Rows.Count, 1).End(xlUp).Row
    ultimaDerecha = Cells(4, Columns.Count).End(xlToLeft).Column
    Debug.Print "Ultima fila: " & ultimaConDatos
    Debug.Print "Ultima columna: " & ultimaDerecha
End Function

Sub Marcar()
Attribute Marcar.VB_ProcData.VB_Invoke_Func = "M\n14"
    ' Ctrol +  May + M
    ' MARCA LAS CELDAS CON COLOR
    ' SOLO LAS CELDAS MARCADAS PUEDEN SUMARSE
    Dim ultimaColumna As Byte
    Call ultima(ultimaConDatos)
    ultimaColumna = Worksheets(1).Cells(2, Columns.Count).End(xlToLeft).Column
    
    If ActiveCell.Column >= 5 And ActiveCell.Column <= ultimaDerecha And ActiveCell.Row > 4 Then
        
        If ActiveCell.Interior.ColorIndex = naranja Or Selection.Interior.ColorIndex = naranja Then
            ActiveCell.Interior.Color = xlNone
            Selection.Interior.ColorIndex = xlNone
        Else
            ActiveCell.Interior.ColorIndex = naranja
            Selection.Interior.ColorIndex = naranja
        End If
    End If
End Sub

' Inserta una persona en la fila Nº 5.
Sub InsPersona()
    Rows(5).Insert Shift:=xlShiftDown, CopyOrigin:=xlFormatFromRightOrBelow
    Cells(5, 1).Activate
    Call ultima(ultimaConDatos)
    Cells(ultimaConDatos, 2).Formula = "=COUNTA(" & "B5:B" & ultimaConDatos - 1 & ")"
    ThisWorkbook.Save
End Sub

' Borra una fila, siempre la Nº 5. Deja una solita, nada más.
Sub BorrarPersona()
    Call Advertencia
    Call ultima(ultimaConDatos)
    If ultimaConDatos > 6 Then
        Rows(5).Delete
        Cells(5, 1).Activate
    Else
        MsgBox "No se puede borrar esta fila."
    End If
    ThisWorkbook.Save
End Sub

' Inserta un producto: son 3 columnas, con su formato, fórmula y restricciones.
Sub InsProducto()
    Application.DisplayAlerts = False
    Call ultima(ultimaConDatos)
    Columns("E:G").Select
    Selection.Insert CopyOrigin:=xlFormatFromRightOrBelow
    Range(Cells(2, 8), Cells(ultimaConDatos - 1, 10)).Select
    Selection.Copy
    Range("E2").Activate
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    Columns("H:J").Select
    Columns("H:J").Copy
    Columns("E:G").Select
    Selection.PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
    Range("E2").Select
    Range("E2").Value = ""
    Range(Cells(5, 5), Cells(ultimaConDatos - 1, 7)).Value = ""
    Application.DisplayAlerts = True
    ActiveWorkbook.Save
End Sub

' Borra un producto (grupo de 3 columnas), siempre las columnas E a G.
Sub RemProducto()
    Call Advertencia
    Call ultima(ultimaConDatos)
    Application.DisplayAlerts = False
    If ultimaDerecha < 8 Then
        MsgBox "No se puede eliminar el último producto."
        Range("E2").Select
        ActiveWorkbook.Save
        Exit Sub
    End If
    Columns("E:G").Select
    Selection.Delete
    Range("E2").Select
    Application.DisplayAlerts = True
    ActiveWorkbook.Save
End Sub

' Calcula un resumen de totales, separados y faltantes.
Sub Faltantes()
    Debug.Print "Calculiando faltantes..."
End Sub
