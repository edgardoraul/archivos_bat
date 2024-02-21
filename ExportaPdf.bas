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
    
    ' Exportar como PDF con el mismo nombre y en la misma carpeta
    
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=rutaArchivo & ".pdf", _
        Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, _
        OpenAfterPublish:=True
        
    ' Cerrar el archivo activo
    Application.DisplayAlerts = False ' Desactivar alertas para cerrar guardando cambios
    ActiveWorkbook.Close SaveChanges:=True
    Application.DisplayAlerts = True ' Reactivar alertas
    
    ' Muestra el archivo en carpeta para enviar por mail o imprimir.
    'Shell "explorer " & rutaArchivo, vbNormalFocus

End Sub


