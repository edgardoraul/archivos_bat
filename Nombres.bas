Attribute VB_Name = "Nombres"
Option Explicit
Global BackgroundColor As Byte
Global TextColor As Byte
Public Const clave As String = "Rerda2025"
Public ultima As Integer

Sub CompletarNombres()
    Dim columnaImagen As Boolean
    Dim esSegundaImagen As Boolean
    Dim col As Integer
    Dim CantNombre As Integer
    Dim CantNombreReal As Integer
    Dim img As Picture
    Dim archivoImagen As String
    Dim archivoImagen2 As String
    Dim usarSegundaImagen As Boolean
    Dim celdaDestino As Range
    Dim contenido As String
    Dim filaFuente As Integer
    Dim celdaFuente As String
    Dim r As Long
    
    ' Ultima fila con datos
    ultima = Worksheets("Listado").Cells(Rows.Count, 1).End(xlUp).Row

    ' Controla si ya está editado
    If Sheets("Nombres").Range("B1").Value <> "" Then
        MsgBox "Ya está editado este archivo." & vbNewLine & "Guardá una copia limpia para trabajar " & vbNewLine & "o borrá este contenido."
        Exit Sub
    End If

    CantNombreReal = ultima - 1
    CantNombre = CantNombreReal

    If CantNombre <= 0 Then
        MsgBox "Operación cancelada o cantidad inválida."
        Exit Sub
    ElseIf CantNombre > 40000 Then
         MsgBox "¿Cuántos nombres vas a imprimir? " & vbNewLine & "Hasta 40000 (límite de filas)"
         Exit Sub
    ElseIf CantNombre Mod 2 > 0 Then
        CantNombre = CantNombre + 1
    End If

    ' Validar contenido
    On Error Resume Next
    contenido = Application.InputBox("Escribí el texto que se va a repetir", Type:=2)
    On Error GoTo 0

    If contenido = "" Then
        MsgBox "Operación cancelada o no ingresaste nada."
        Exit Sub
    End If

IMAGEN1:
    On Error Resume Next
    archivoImagen = Application.GetOpenFilename("Archivos de imagen (*.jpg; *.jpeg; *.png; *.gif),*.jpg;*.jpeg;*.png;*.gif", , "Selecciona la PRIMERA imagen")
    On Error GoTo 0

    If archivoImagen = "False" Or archivoImagen = "" Then
        MsgBox "Tenés que elegir alguna imagen"
        GoTo IMAGEN1
    End If

    ' Preguntar si se agrega la segunda columna de imagen
    If MsgBox("¿Querés agregar una segunda columna de imagen?", vbQuestion + vbYesNo, "Imagen Adicional") = vbYes Then
        usarSegundaImagen = True
IMAGEN2:
        On Error Resume Next
        archivoImagen2 = Application.GetOpenFilename("Archivos de imagen (*.jpg; *.jpeg; *.png; *.gif),*.jpg;*.jpeg;*.png;*.gif", , "Selecciona la SEGUNDA imagen")
        On Error GoTo 0

        If archivoImagen2 = "False" Or archivoImagen2 = "" Then
            MsgBox "Tenés que elegir la segunda imagen"
            GoTo IMAGEN2
        End If
    Else
        usarSegundaImagen = False
    End If

COLOR:
    Call colorFondo
    If BackgroundColor < 1 Or BackgroundColor > 56 Then
        MsgBox "Tenés que elegir el número de un color válido (1-56)"
        GoTo COLOR
    End If

    r = 1
    filaFuente = 1

    While r <= CantNombre / 2
        ' Obtener primer nombre
        celdaFuente = ObtenerNombreFuente(filaFuente, CantNombreReal)

        If Not usarSegundaImagen Then
            ' --- ESTRUCTURA NORMAL (4 COLUMNAS -> Rótulo = 7,5 cm) ---
            
            ' Col 1: Img 1
            col = 1
            Set celdaDestino = Cells(r, col)
            Call formato(True, False, False, celdaDestino, BackgroundColor, TextColor)
            Set img = ActiveSheet.Pictures.Insert(archivoImagen)
            Call redimensionar(img, celdaDestino)

            ' Col 2: Texto
            col = 2
            Set celdaDestino = Cells(r, col)
            Call formato(False, False, False, celdaDestino, BackgroundColor, TextColor)
            Call InsertarTexto(celdaDestino, filaFuente, celdaFuente, contenido, CantNombreReal)

            ' Col 3: Img 1
            col = 3
            Set celdaDestino = Cells(r, col)
            Call formato(True, False, False, celdaDestino, BackgroundColor, TextColor)
            Set img = ActiveSheet.Pictures.Insert(archivoImagen)
            Call redimensionar(img, celdaDestino)

            ' Col 4: Texto
            col = 4
            filaFuente = filaFuente + 1
            celdaFuente = ObtenerNombreFuente(filaFuente, CantNombreReal)
            Set celdaDestino = Cells(r, col)
            Call formato(False, False, False, celdaDestino, BackgroundColor, TextColor)
            Call InsertarTexto(celdaDestino, filaFuente, celdaFuente, contenido, CantNombreReal)

        Else
            ' --- ESTRUCTURA CON SEGUNDA IMAGEN (6 COLUMNAS -> Rótulo = 7,5 cm) ---
            
            ' Col 1 (A): Img 1
            col = 1
            Set celdaDestino = Cells(r, col)
            Call formato(True, False, True, celdaDestino, BackgroundColor, TextColor)
            Set img = ActiveSheet.Pictures.Insert(archivoImagen)
            Call redimensionar(img, celdaDestino)

            ' Col 2 (B): Texto
            col = 2
            Set celdaDestino = Cells(r, col)
            Call formato(False, False, True, celdaDestino, BackgroundColor, TextColor)
            Call InsertarTexto(celdaDestino, filaFuente, celdaFuente, contenido, CantNombreReal)

            ' Col 3 (C): Img 2 (Sin borde izquierdo)
            col = 3
            Set celdaDestino = Cells(r, col)
            Call formato(True, True, True, celdaDestino, BackgroundColor, TextColor)
            Set img = ActiveSheet.Pictures.Insert(archivoImagen2)
            Call redimensionar(img, celdaDestino)

            ' Col 4 (D): Img 1
            col = 4
            Set celdaDestino = Cells(r, col)
            Call formato(True, False, True, celdaDestino, BackgroundColor, TextColor)
            Set img = ActiveSheet.Pictures.Insert(archivoImagen)
            Call redimensionar(img, celdaDestino)

            ' Col 5 (E): Texto
            col = 5
            filaFuente = filaFuente + 1
            celdaFuente = ObtenerNombreFuente(filaFuente, CantNombreReal)
            Set celdaDestino = Cells(r, col)
            Call formato(False, False, True, celdaDestino, BackgroundColor, TextColor)
            Call InsertarTexto(celdaDestino, filaFuente, celdaFuente, contenido, CantNombreReal)

            ' Col 6 (F): Img 2 (Sin borde izquierdo)
            col = 6
            Set celdaDestino = Cells(r, col)
            Call formato(True, True, True, celdaDestino, BackgroundColor, TextColor)
            Set img = ActiveSheet.Pictures.Insert(archivoImagen2)
            Call redimensionar(img, celdaDestino)
        End If

        r = r + 1
        filaFuente = filaFuente + 1
    Wend
    
    ' --- CONFIGURACIÓN DE ENCABEZADO (Número de Página) ---
    With ActiveSheet.PageSetup
        .CenterHeader = "Página &P"
    End With

    On Error Resume Next
    ActiveWorkbook.Save
    On Error GoTo 0

    Set img = Nothing
    Set celdaDestino = Nothing
    Call Proteger
End Sub

' Función auxiliar para obtener el nombre sin desbordar en impares
Function ObtenerNombreFuente(filaFuente As Integer, cantReal As Integer) As String
    If filaFuente <= cantReal Then
        ObtenerNombreFuente = UCase(Worksheets("Listado").Cells(filaFuente + 1, 1).Value & " " & Sheets("Listado").Cells(filaFuente + 1, 2).Value)
    Else
        ObtenerNombreFuente = ""
    End If
End Function

' Función auxiliar para escribir el texto
Sub InsertarTexto(celdaDestino As Range, filaFuente As Integer, celdaFuente As String, contenido As String, cantReal As Integer)
    Dim campo3 As String
    
    If filaFuente <= cantReal Then
        campo3 = UCase(Sheets("Listado").Cells(filaFuente + 1, 3).Value)
    Else
        campo3 = ""
    End If

    If celdaFuente = "" Then
        celdaDestino.Value = contenido
        If campo3 <> "" Then celdaDestino.Value = celdaDestino.Value & vbNewLine & campo3
    Else
        If campo3 = "" Then
            celdaDestino.Value = celdaFuente & vbNewLine & contenido
        Else
            celdaDestino.Value = celdaFuente & vbNewLine & contenido & vbNewLine & campo3
        End If
    End If
    
    celdaDestino.Value = RTrim(celdaDestino.Value)
End Sub

Function formato(columnaImagen As Boolean, esSegundaImagen As Boolean, modoDosImagenes As Boolean, celdaDestino As Range, BackgroundColor As Byte, TextColor As Byte)
    Const ALTO = 54
    Const ANCHOIMAGEN = 9.5 ' ~1.9 cm (celda cuadrada)
    
    ' Ajuste del ancho de texto para mantener total exacto de 7.5 cm
    Dim anchoTexto As Double
    If modoDosImagenes Then
        anchoTexto = 19.5 ' ~3.7 cm (1.9 + 3.7 + 1.9 = 7.5 cm)
        celdaDestino.Font.Size = 8
    Else
        anchoTexto = 29.5 ' ~5.6 cm (1.9 + 5.6 = 7.5 cm)
        celdaDestino.Font.Size = 10
    End If

    With celdaDestino
        .Interior.ColorIndex = BackgroundColor
        .Font.ColorIndex = TextColor
        .Font.Bold = True
        
        .Font.Name = "Arial"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter

        .Borders.LineStyle = xlDouble
        .Borders.ColorIndex = 16
        .Borders.Weight = xlMedium

        .RowHeight = ALTO
        If columnaImagen Then
            .EntireColumn.ColumnWidth = ANCHOIMAGEN
            If esSegundaImagen Then
                .Borders(xlEdgeLeft).LineStyle = xlNone
            End If
        Else
            .EntireColumn.ColumnWidth = anchoTexto
            If .Column > 1 Then
                 .Borders(xlEdgeLeft).LineStyle = xlNone
            End If
            .WrapText = True
        End If
    End With
End Function

Function redimensionar(img As Picture, celdaDestino As Range)
On Error GoTo ErrorHandler

    Dim PicWtoHRatio As Single
    Dim CellWtoHRatio As Single
    Dim targetWidth As Single
    Dim targetHeight As Single

    PicWtoHRatio = img.Width / img.Height
    CellWtoHRatio = celdaDestino.Width / celdaDestino.Height

    If PicWtoHRatio / CellWtoHRatio > 1 Then
        targetWidth = celdaDestino.Width - 4
        targetHeight = targetWidth / PicWtoHRatio
    Else
        targetHeight = celdaDestino.Height - 4
        targetWidth = targetHeight * PicWtoHRatio
    End If

    img.Width = targetWidth
    img.Height = targetHeight

    img.Top = celdaDestino.Top + (celdaDestino.Height - img.Height) / 2
    img.Left = celdaDestino.Left + (celdaDestino.Width - img.Width) / 2

    Exit Function

ErrorHandler:
    Resume Next
End Function

Sub colorFondo()
    Dim tempSheet As Worksheet
    Dim x As Byte
    Dim inputColor As Variant
    Dim inputColorText As Variant
    
    Call Desproteger

    Set tempSheet = ActiveWorkbook.Worksheets.Add(After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Sheets.Count))
    tempSheet.Name = "PaletaColoresTemp"

     For x = 1 To 56
        If x <= 14 Then
            Cells(x, 1).Interior.ColorIndex = x
            Cells(x, 2).Value = x
            Cells(x, 2).Borders(xlEdgeRight).LineStyle = xlSolid
        ElseIf x <= 28 Then
            Cells(x - 14, 3).Interior.ColorIndex = x
            Cells(x - 14, 4).Value = x
            Cells(x - 14, 4).Borders(xlEdgeRight).LineStyle = xlSolid
        ElseIf x <= 42 Then
            Cells(x - 28, 5).Interior.ColorIndex = x
            Cells(x - 28, 6).Value = x
            Cells(x - 28, 6).Borders(xlEdgeRight).LineStyle = xlSolid
        Else
            Cells(x - 42, 7).Interior.ColorIndex = x
            Cells(x - 42, 8).Value = x
            Cells(x - 42, 8).Borders(xlEdgeRight).LineStyle = xlSolid
        End If
    Next x

    tempSheet.Columns("A:H").AutoFit

    inputColor = Application.InputBox("Escribí el número de color de fondo (1-56) y presiona Enter.", Type:=1)
    inputColorText = Application.InputBox("Escribí el número de color del texto (1-56) y presiona Enter.", Type:=1)
    
    Call Desalertar
    Sheets("Nombres").Activate
    
    tempSheet.Delete
    
    If IsNumeric(inputColor) Then
        BackgroundColor = CByte(inputColor)
    Else
        BackgroundColor = 0
    End If

    If IsNumeric(inputColorText) Then
        TextColor = CByte(inputColorText)
    Else
        TextColor = 0
    End If
    
    Set tempSheet = Nothing
    Call Proteger
End Sub

Function Proteger()
    Call Desalertar
    Dim Archivo As Workbook
    Dim i As Byte
    Set Archivo = ThisWorkbook
    
    For i = 1 To Archivo.Worksheets.Count
        If Archivo.Worksheets(i).Name <> "Nombres" Then
            Archivo.Worksheets(i).Protect clave
        End If
    Next i
    Archivo.Protect clave
    Archivo.Save
    Call Alertar
End Function

Function Desproteger()
    Call Desalertar
    Dim Archivo As Workbook
    Dim i As Byte
    Set Archivo = ThisWorkbook
    Archivo.Unprotect clave
    For i = 1 To Archivo.Worksheets.Count
        Archivo.Worksheets(i).Unprotect clave
    Next i
    Call Alertar
    Archivo.Save
End Function

Function Desalertar()
    Application.DisplayAlerts = False
End Function

Function Alertar()
    Application.DisplayAlerts = True
End Function
