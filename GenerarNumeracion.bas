Attribute VB_Name = "GenerarNumeracion"
Option Explicit
Sub GenerarPaginacionExacta()
    Dim inicio As Long, fin As Long, i As Long, fila As Long
    Dim numFormateado As String
    
    ' Entradas del usuario
    inicio = InputBox("Ingrese el número inicial:", "Paginación")
    fin = InputBox("Ingrese el número final:", "Paginación")
    
    
    If fin < inicio Then Exit Sub
    
    With ActiveSheet
        .Range("A1").CurrentRegion
        .Cells.Clear
        .ResetAllPageBreaks
        .Cells.PageBreak = xlPageBreakNone
        .Range("A1").Activate
    End With

    Application.ScreenUpdating = False
    
    ' Configuración de fuente
    With Cells.Font
        .Name = "Arial"
        .Size = 10
    End With
    
    fila = 1
    For i = inicio To fin
        ' Formato de 8 dígitos
        numFormateado = "'" & Format(i, "00000000")
        Cells(fila, 8).Value = numFormateado
        Cells(fila, 8).HorizontalAlignment = xlRight
        
        ' Insertar salto de página después de cada número
        If i < fin Then
            ActiveSheet.HPageBreaks.Add Before:=Cells(fila + 1, 1)
        End If
        fila = fila + 1
    Next i
    
    ' Configuración de página (Márgenes en Centímetros)
    With ActiveSheet.PageSetup
        .PaperSize = xlPaperA4
        ' 1.5 cm de margen superior y derecho
        .TopMargin = Application.CentimetersToPoints(1.05)
        .RightMargin = Application.CentimetersToPoints(0.5)
        ' Márgenes restantes mínimos para no desplazar el número
        .LeftMargin = Application.CentimetersToPoints(2)
        .BottomMargin = Application.CentimetersToPoints(2)
        .HeaderMargin = 0
        .FooterMargin = 0
        .PrintGridlines = False
    End With
    
    Application.ScreenUpdating = True
    MsgBox "Hojas generadas con éxito.", vbInformation
End Sub

