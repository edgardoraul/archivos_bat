Attribute VB_Name = "ExportaPdf"
Option Explicit
Sub ExportarComoPDF()

    Dim nombreArchivo As String
    Dim rutaArchivo As String
    Dim posPunto As Integer
    
    ' Obtener el nombre del archivo activo sin la extensión
    posPunto = InStrRev(ActiveWorkbook.Name, ".")
    nombreArchivo = Left(ActiveWorkbook.Name, posPunto - 1)
    
    ' Obtener la ruta del archivo activo
    rutaArchivo = ActiveWorkbook.Path & "\" & nombreArchivo
    
    Debug.Print nombreArchivo
    
    ' Formato de impresión
    With ActiveSheet.PageSetup
        '.Orientation = xlPortrait
        '.PaperSize = xlPaperA4
        '.LeftMargin = Application.CentimetersToPoints(0.64)
        '.RightMargin = Application.CentimetersToPoints(0.64)
        '.TopMargin = Application.CentimetersToPoints(2.5)
        '.BottomMargin = Application.CentimetersToPoints(1.91)
        '.HeaderMargin = Application.CentimetersToPoints(0.76)
        '.FooterMargin = Application.CentimetersToPoints(0.76)
        '.CenterHorizontally = True '
        '.CenterVertically = False '
        '.PrintArea = ActiveSheet.Range("A1:H21")
        '.Zoom = False
        '.FitToPagesTall = 1
        '.FitToPagesWide = 1
    End With
    
    ' Exportar como PDF con el mismo nombre y en la misma carpeta
    
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, fileName:=rutaArchivo & ".pdf", _
        Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, _
        OpenAfterPublish:=True
        
    ' Cerrar el archivo activo
    Application.DisplayAlerts = False ' Desactivar alertas para cerrar guardando cambios
    ActiveWorkbook.Close SaveChanges:=True
    Application.DisplayAlerts = True ' Reactivar alertas
    
    ' Muestra el archivo en carpeta para enviar por mail o imprimir.
    'Shell "explorer " & rutaArchivo, vbNormalFocus

End Sub


