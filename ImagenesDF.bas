Attribute VB_Name = "ImagenesDF"
Option Explicit
Dim ultima As String
Dim ruta As String
Dim ruta_base As String
Dim codigo As String
Dim archivo As String
Dim acumulador As Integer

Sub ImagenesDF()
' CREA LAS RUTAS PARA IMPORTAR IMAGENES AL PESCADO DRAGÓN
' LA PRIMERA PARTE DEL ENLACE QUE NO CAMBIA
ruta = ThisWorkbook.Path & "\Exportado.txt"
ruta_base = Worksheets("Constantes").Cells(13, 2).Value
archivo = "\1.jpg"

' 1º - SE PARE EN LA PRIMERA HOJA y la limpia
Worksheets("Listado").Activate
ActiveSheet.Cells.ClearContents

' 2º - Crea los títulos
With Worksheets("Listado")
    .Cells(1, 1).Value = "Código"
    .Cells(1, 2).Value = "Enlace"
End With

acumulador = 2

' 3º - Ejecuta una función de copia
Call copiaCodigos("Variables")
Call copiaCodigos("Con Color")
Call copiaCodigos("Simples")
Call copiaCodigos("Con Talles")


MsgBox "Archivo exportado en: " & ruta, vbInformation
End Sub

Function copiaCodigos(hoja)
    Dim i As Integer
    Dim Servidor As String
    Dim total As Long
    
    Servidor = Worksheets("Constantes").Range("B15").Value
    ' Va copiando códigos de arriba hacia abajo y generando enlaces
    ultima = Worksheets(hoja).Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To ultima
        ' Parándose en la celda correspondiente
        'Cells(acumulador, 1).Activate
        Cells(1, 3).Activate
        Cells(1, 3).Value = acumulador - 1
        
        If Worksheets(hoja).Cells(i, 8).Value = "" Then
            GoTo Siguiente
        End If
        
        
        ' Obteniendo datos
        codigo = Left(Worksheets(hoja).Cells(i, 3).Value, 7)
        If codigo = Left(Worksheets(hoja).Cells(i - 1, 3).Value, 7) Then
            GoTo Siguiente
        End If
        ruta = "" & ruta_base & codigo & archivo & ""
        
        ' Escribiendo registros
        Cells(acumulador, 1).Value = codigo
        Cells(acumulador, 2).Value = ruta
        Call CopiarImagen_PorCodigo(codigo, i, archivo, Servidor, ruta_base)
        
        ' Aumento el acumulador para no pisar información cargada
        acumulador = acumulador + 1
        
Siguiente:
    Next i
    Call ExportarComoTXT_Tab
    
End Function


Function ExportarComoTXT_Tab()
    Dim archivo As Integer
    
    Dim fila As Long, col As Long
    Dim ultimaFila As Long, ultimaCol As Long
    Dim linea As String
    
    ruta = ThisWorkbook.Path & "\Exportado.txt"
    
    '?? Abrir archivo para escritura
    archivo = FreeFile
    Open ruta For Output As #archivo
    
    '?? Detectar rango usado
    With ActiveSheet
        ultimaFila = .Cells(.Rows.Count, 1).End(xlUp).Row
        ultimaCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
        
        '?? Recorrer filas
        For fila = 1 To ultimaFila
            linea = ""
            For col = 1 To ultimaCol
                linea = linea & .Cells(fila, col).Text
                If col < ultimaCol Then linea = linea & vbTab
            Next col
            Print #archivo, linea
        Next fila
    End With
    
    Close #archivo
    
    
End Function


Function CopiarImagen_PorCodigo(codigo, i, NombreImagen, CarpetaRaizDestino, RutaDelOrigen)
    
    On Error GoTo ErrorHandler
    
    ' Declaración de constantes y variables
    'Const NombreImagen As String = "1.jpg"
    'Dim CarpetaRaizDestino As String
    Dim RutaOrigen As String
    Dim RutaDestino As String
    Dim FSO As Object ' Se usa para manipular archivos y carpetas
    
    ' Obtener la ruta raíz de destino de la celda específica (puedes cambiar "Hoja1" y "B1" a tus necesidades)
    'CarpetaRaizDestino = ThisWorkbook.Sheets("Hoja1").Range("B1").Value
    
    ' Construir las rutas completas
    ' La Carpeta de Origen es igual al Código
    ' La Carpeta de Destino es igual al Código
    
    RutaOrigen = RutaDelOrigen & codigo & NombreImagen
    RutaDestino = CarpetaRaizDestino & codigo & NombreImagen
    
    ' Crear el objeto FileSystemObject
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    ' Verificar si la carpeta de destino existe. Si no, la creará.
    If Not FSO.FolderExists(CarpetaRaizDestino & codigo) Then
        FSO.CreateFolder CarpetaRaizDestino & codigo
    End If
    
    ' Copiar el archivo
    ' Se usa 'True' para sobrescribir el archivo si ya existe
    FSO.CopyFile RutaOrigen, RutaDestino, True
    
    ' Limpiar el objeto
    Set FSO = Nothing
    

    Exit Function
    
ErrorHandler:
    ' Manejo de errores
    MsgBox "Error: " & Err.Number & " - " & Err.Description, vbCritical
    CopiarImagen_PorCodigo = False
    Worksheets("Listado").Cells(i, 3).Value = "Error, no se copio"
    
End Function
