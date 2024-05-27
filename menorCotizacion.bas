Attribute VB_Name = "menorCotizacion"
Option Explicit
Sub resaltarCeldas()
    Dim ws As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim i As Long, j As Long
    Dim minValue As Double
    Dim minCol As Long
    
    ' Definir la hoja de trabajo
    Set ws = Worksheets(1)
    
    ' Encontrar la última fila y columna con datos
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ' Iterar sobre cada fila
    For i = 3 To lastRow
        ' Inicializar el valor mínimo y la columna correspondiente
        minValue = ws.Cells(i, 5).Value ' Empezamos desde la quinta columna
        minCol = 5
        
        ' Encontrar el valor mínimo en la fila (ignorando ceros)
        For j = 5 To lastCol ' Empezamos desde la tercera columna
            If ws.Cells(i, j).Value > 0 And ws.Cells(i, j).Value < minValue Then
                minValue = ws.Cells(i, j).Value
                minCol = j
            End If
        Next j
        
        ' Comparar el valor mínimo con el valor en la columna principal (columna A)
        If minValue > ws.Cells(i, 3).Value And ws.Cells(i, 3).Value > 0 Then
            ' Cambiar el color de fondo de la celda en la columna principal a rosa
            ws.Cells(i, 3).Interior.color = RGB(255, 192, 203) ' Color rosa claro
        Else
            ws.Cells(i, 3).Interior.color = xlNone
        End If
    Next i
End Sub

