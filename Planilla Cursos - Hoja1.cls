VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Calculate()
    ' Salir si no hay un autofiltro activo en la hoja.
    If Me.AutoFilterMode = False Then Exit Sub
    Call ultima
    Me.AutoFilterMode = False
    Range(Cells(4, 1), Cells(4, ultimaDerecha)).Interior.ColorIndex = 15
    
    Range(Cells(4, 1), Cells(4, ultimaDerecha)).Locked = True
    ActiveSheet.Protect pass
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim ColMonitorear As Range
    
    Call ultima
    
    
    
    ' Define el rango de celdas que activará el evento (por ejemplo, columna A)
    Set ColMonitorear = Intersect(Target, Range(Cells(5, 1), Cells(ultimaConDatos, 1)))
      
    If ColMonitorear Is Nothing Then
        Exit Sub
    End If
    
    
    If ColMonitorear = Worksheets("VARIANTES").Range("C3").Value Then
        Call Desactivar(Target.row)
    ElseIf ColMonitorear = Worksheets("VARIANTES").Range("C2").Value Then
        Call Activar(Target.row)
    Else
        Debug.Print "Nada de nada"
    End If

End Sub
