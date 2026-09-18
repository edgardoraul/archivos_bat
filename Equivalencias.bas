Attribute VB_Name = "Equivalencias"
Option Explicit
Sub GenerarEquivalencias()
    Dim wsEquiv As Worksheet
    Dim wsTabla As Worksheet
    Dim ultimaFila As Long
    Dim i As Long
    Dim valArticulo As String
    Dim valColor As String
    Dim valTalle As String
    Dim talleMapeado As String
    Dim celdaEncontrada As Range

    On Error GoTo ManejoErrores
    
    Set wsEquiv = ActiveWorkbook.Sheets("Equivalencia")
    Set wsTabla = ActiveWorkbook.Sheets("Tabla")
    
    ' Obtener última fila según columna B (Artículo)
    ultimaFila = wsEquiv.Cells(wsEquiv.Rows.Count, "B").End(xlUp).Row
    
    'Application.ScreenUpdating = False
    
    For i = 2 To ultimaFila
        ' Para ver dónde estamos
        wsEquiv.Cells(i, 1).Activate
        
        ' Lectura de valores por columna (B: Artículo, D: Color, F: Talle)
        valArticulo = Trim(wsEquiv.Cells(i, "B").Value)
        valColor = Trim(wsEquiv.Cells(i, "D").Value)
        valTalle = Trim(wsEquiv.Cells(i, "F").Value)
        
        talleMapeado = ""
        
        ' Buscar el talle en la hoja Tabla
        If valTalle <> "" Then
            Set celdaEncontrada = wsTabla.Cells.Find(What:=valTalle, LookIn:=xlValues, LookAt:=xlWhole)
            
            If Not celdaEncontrada Is Nothing Then
                ' Toma el valor de la primera columna (Columna A) de la fila hallada
                talleMapeado = wsTabla.Cells(celdaEncontrada.Row, 1).Value
            Else
                ' Si no lo encuentra en Tabla, mantiene el talle original
                talleMapeado = valTalle
            End If
        End If
        
        ' Genera el código concatenado en la Columna A (Código)
        wsEquiv.Cells(i, "A").Value = valArticulo & valColor & talleMapeado
    Next i

    'Application.ScreenUpdating = True
    MsgBox "Equivalencias generadas correctamente.", vbInformation, "Proceso Finalizado"
    Exit Sub

ManejoErrores:
    'Application.ScreenUpdating = True
    MsgBox "Ocurrió un error: " & Err.Description, vbCritical, "Error"
End Sub
