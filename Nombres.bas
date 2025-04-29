Attribute VB_Name = "Nombres"
Option Explicit
Global BackgroundColor As Byte
Global TextColor As Byte
Public Const clave As String = "Rerda2025"
Public ultima As Integer

Sub CompletarNombres()
    Dim columnaImagen As Boolean
    Dim columna As Integer
    Dim fila As Integer
    Dim f As Integer
    Dim c As Integer
    Dim CantNombre As Integer
    Dim img As Picture
    Dim archivoImagen As String
    Dim celdaDestino As Range
    Dim contenido As String
    Dim filaFuente As Integer
    Dim celdaFuente As String
    
    ' Ultima fila con datos
    ultima = Sheets("Listado").Cells(Rows.Count, 1).End(xlUp).Row

    ' Usamos Long para filas y columnas en bucles grandes, aunque Byte/Integer es suficiente aquí
    Dim r As Long
    Dim col As Long

    ' Controla si ya está editado
    If Sheets("Nombres").Range("B1").Value <> "" Then
        MsgBox "Ya está editado este archivo." & vbNewLine & "Guardá una copia limpia para trabajar " & vbNewLine & "o borrá este contenido."
        Exit Sub
    End If


    ' Validar la entrada de CantNombre
    ' Para manejar si el usuario cancela el InputBox
    On Error Resume Next
    
    'CantNombre = Application.InputBox("¿Cuántos Nombres necesitás?", Type:=1) ' Type:=1 asegura que la entrada sea un número
    ' Restablecer el manejo de errores
    On Error GoTo 0
    
    CantNombre = ultima - 1

    ' El usuario canceló o ingresó 0
    If CantNombre <= 0 Then
        MsgBox "Operación cancelada o cantidad inválida."
        Exit Sub
        
    ' Limite razonable para filas en versiones antiguas
    ElseIf CantNombre < 0 Or CantNombre > 65536 Then
         MsgBox ("¿Cuántos nombres vas a imprimir? " & vbNewLine & "Hasta 65536 (límite de filas en versiones antiguas)") ' Ajustado el mensaje
         Exit Sub
    
    ' Para saber asegurarse que siempre sea par
    ElseIf CantNombre Mod 2 > 0 Then
        CantNombre = CantNombre + 1
    End If

    ' Validar la entrada de contenido
    ' Para manejar si el usuario cancela el InputBox
    On Error Resume Next
    
    ' Type:=2 asegura que la entrada sea texto
    contenido = Application.InputBox("Escribí el texto que se va a repetir", Type:=2)
    contenido = UCase(contenido)
    
    ' Restablecer el manejo de errores
    On Error GoTo 0

    ' El usuario canceló o no ingresó texto
    If contenido = "" Then
        MsgBox "Operación cancelada o no ingresastre nada."
        Exit Sub
    End If


IMAGEN:
    ' Validar la selección de archivo de imagen
    ' Para manejar si el usuario cancela el GetOpenFilename
    On Error Resume Next
    archivoImagen = Application.GetOpenFilename("Archivos de imagen (*.jpg; *.jpeg; *.png; *.gif),*.jpg;*.jpeg;*.png;*.gif", , "Selecciona una imagen")
    
    ' Restablecer el manejo de errores
    On Error GoTo 0

    ' GetOpenFilename devuelve "False" si se cancela
    If archivoImagen = "False" Or archivoImagen = "" Then
        MsgBox "Tenés que elegir alguna imagen"
        
        ' Volver a pedir la imagen
        GoTo IMAGEN
    End If

COLOR:
    ' Validar la selección del color de fondo
    ' colorFondo ahora devuelve el color directamente
    Call colorFondo
    If BackgroundColor < 1 Or BackgroundColor > 56 Then
        MsgBox "Tenés que elegir el número de un color válido (1-56)"
        
        ' Volver a pedir el color
        GoTo COLOR
    End If


    ' Comienza en primera fila
    r = 1
    filaFuente = 1
    
    ' El bucle ahora recorre la cantidad total de nombres, distribuyendo en 2 columnas
    ' Recorre las filas necesarias para la mitad de nombres
    While r <= CantNombre / 2
        
        celdaFuente = UCase(Sheets("Listado").Cells(filaFuente + 1, 1).Value & " " & Sheets("Listado").Cells(filaFuente + 1, 2).Value)
        col = 1
        Set celdaDestino = Cells(r, col)
        columnaImagen = True
        Call formato(columnaImagen, celdaDestino, BackgroundColor, TextColor) ' Pasar la celda de destino
        
        ' Insertar la imagen y obtener una referencia al objeto Picture insertado
        Set img = ActiveSheet.Pictures.Insert(archivoImagen)
        ' Redimensionar y posicionar usando el objeto img y la celda de destino
        Call redimensionar(img, celdaDestino)

        ' Columna 2 (Texto)
        col = 2
        Set celdaDestino = Cells(r, col)
        columnaImagen = False
        
        ' Pasar la celda de destino
        Call formato(columnaImagen, celdaDestino, BackgroundColor, TextColor)
        
        ' Contenido
        celdaDestino.Value = celdaFuente & vbNewLine & contenido & vbNewLine & UCase(Sheets("Listado").Cells(filaFuente + 1, 3).Value)
        celdaDestino.Value = RTrim(celdaDestino.Value)

        ' Columna 3 (Imagen)
        col = 3
        Set celdaDestino = Cells(r, col)
        columnaImagen = True
        Call formato(columnaImagen, celdaDestino, BackgroundColor, TextColor)
        
        ' Insertar la imagen y obtener una referencia al objeto Picture insertado
        Set img = ActiveSheet.Pictures.Insert(archivoImagen)
        
        ' Redimensionar y posicionar usando el objeto img y la celda de destino
        Call redimensionar(img, celdaDestino)

        ' Columna 4 (Texto)
        col = 4
        filaFuente = filaFuente + 1
        celdaFuente = UCase(Sheets("Listado").Cells(filaFuente + 1, 1).Value & " " & Sheets("Listado").Cells(filaFuente + 1, 2).Value)
        Set celdaDestino = Cells(r, col)
        columnaImagen = False
        Call formato(columnaImagen, celdaDestino, BackgroundColor, TextColor)
        
        ' Contenido
        celdaDestino.Value = celdaFuente & vbNewLine & contenido & vbNewLine & UCase(Sheets("Listado").Cells(filaFuente + 1, 3).Value)
        celdaDestino.Value = RTrim(celdaDestino.Value)
        

        ' Contador para las filas
        r = r + 1
        filaFuente = filaFuente + 1
    Wend

    ' Guardar el libro
    On Error Resume Next ' Para manejar si el usuario cancela al guardar
    ActiveWorkbook.Save
    On Error GoTo 0 ' Restablecer el manejo de errores

    ' Limpiar objetos
    Set img = Nothing
    Set celdaDestino = Nothing
    Call Proteger
End Sub

' Modificada para aceptar el rango de la celda de destino
Function formato(columnaImagen As Boolean, celdaDestino As Range, BackgroundColor As Byte, TextColor As Byte)
' Da formato a la tarjeta
Const ALTO = 54
Const ANCHOTEXTO = 30
Const ANCHOIMAGEN = 7

With celdaDestino
    ' Color de Fondo
    .Interior.ColorIndex = BackgroundColor
    .Font.ColorIndex = TextColor

    ' Fuente
    .Font.Bold = True
    .Font.Size = 11
    .Font.Name = "Arial"
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter

    ' Bordes
    '.Borders.LineStyle = xlContinuous
    .Borders.LineStyle = xlDouble
    ' Usar ColorIndex para bordes también para máxima compatibilidad si RGB da problemas
    .Borders.ColorIndex = 16 ' Un gris oscuro, compatible con ColorIndex
    .Borders.Weight = xlMedium

    ' Dimensiones
    .RowHeight = ALTO
    If columnaImagen = True Then
        .EntireColumn.ColumnWidth = ANCHOIMAGEN
    Else
        .EntireColumn.ColumnWidth = ANCHOTEXTO
        ' Asegurarse de que el borde izquierdo se elimine solo si no es la primera columna
        If .Column > 1 Then
             .Borders(xlEdgeLeft).LineStyle = xlNone
        End If
        .WrapText = True
    End If
End With

End Function

' Modificada para aceptar el objeto Picture y el rango de la celda de destino
Function redimensionar(img As Picture, celdaDestino As Range)
' Redimensiona y posiciona la imagen para que ocupe toda la celda
On Error GoTo ErrorHandler ' Manejo de errores más específico

    Dim PicWtoHRatio As Single
    Dim CellWtoHRatio As Single
    Dim targetWidth As Single
    Dim targetHeight As Single

    ' Calcular la relación de aspecto de la imagen
    PicWtoHRatio = img.Width / img.Height

    ' Calcular la relación de aspecto de la celda de destino
    ' Usamos .Width y .Height del rango, que son más fiables que .TopLeftCell en algunos casos
    CellWtoHRatio = celdaDestino.Width / celdaDestino.Height

    ' Determinar el tamaño objetivo manteniendo la relación de aspecto
    If PicWtoHRatio / CellWtoHRatio > 1 Then ' La imagen es más ancha que la celda (relativamente)
        ' Ajustar al ancho de la celda (menos un pequeño margen)
        targetWidth = celdaDestino.Width - 4
        targetHeight = targetWidth / PicWtoHRatio
    Else ' La imagen es más alta que la celda (relativamente)
        ' Ajustar a la altura de la celda (menos un pequeño margen)
        targetHeight = celdaDestino.Height - 4
        targetWidth = targetHeight * PicWtoHRatio
    End If

    ' Aplicar el nuevo tamaño a la imagen
    img.Width = targetWidth
    img.Height = targetHeight

    ' Posicionar la imagen en el centro de la celda de destino
    ' Usamos las propiedades Top y Left de la celda de destino
    img.Top = celdaDestino.Top + (celdaDestino.Height - img.Height) / 2
    img.Left = celdaDestino.Left + (celdaDestino.Width - img.Width) / 2

    Exit Function ' Salir de la función si todo va bien

ErrorHandler:
    ' Puedes agregar un mensaje de error si lo deseas, pero el error "NOT_SHAPE" original no era crítico
    ' MsgBox "Error en redimensionar: " & Err.Description
    Resume Next ' Continuar la ejecución después del error (si el error no es grave)
End Function

' Modificada para devolver el color seleccionado
Sub colorFondo()
    Dim tempSheet As Worksheet
    Dim x As Byte
    Dim inputColor As Variant ' Usamos Variant para manejar posible cancelación del InputBox
    Dim inputColorText As Variant
    
    Call Desproteger

    
    ' Agregar una hoja temporal para mostrar los colores
    Set tempSheet = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Sheets.Count))
    tempSheet.Name = "PaletaColoresTemp" ' Darle un nombre para identificarla

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

    ' Ajustar el ancho de las columnas para que se vea bien
    tempSheet.Columns("A:H").AutoFit

    ' Pedir al usuario que elija un número de color
    ' Usamos Type:=1 para asegurar que la entrada sea un número
    inputColor = Application.InputBox("Escribí el número de color de fondo (1-56) y presiona Enter.", Type:=1)
    
    ' Elegimos el color del texto
    inputColorText = Application.InputBox("Escribí el número de color del texto (1-56) y presiona Enter.", Type:=1)
    
    Call Desalertar
    Sheets("Nombres").Activate
    
    ' Eliminar la hoja temporal
    tempSheet.Delete
    
    ' Verificar si la entrada es un número válido
    If IsNumeric(inputColor) Then
        BackgroundColor = CByte(inputColor) ' Asignar el color a la variable global
    Else
        BackgroundColor = 0 ' Indicar que no se seleccionó un color de fondo válido
    End If

    If IsNumeric(inputColorText) Then
        TextColor = CByte(inputColorText) ' Asignar el color a la variable global
    Else
        TextColor = 0 ' Indicar que no se seleccionó un color de texto válido
    End If
    
    ' Limpiar objetos
    Set tempSheet = Nothing
    
    Call Proteger
End Sub


Function Proteger()
    ' Protege el archivo y sus hojas excepto "Listados"
    Call Desalertar
    Dim Archivo As Workbook
    Dim i As Byte
    Set Archivo = ThisWorkbook
    
    For i = 1 To Archivo.Sheets.Count
        If Archivo.Sheets(i).Name <> "Nombres" Then
            Archivo.Sheets(i).Protect clave
        End If
    Next i
    Archivo.Protect clave
    Archivo.Save
    Call Alertar
End Function

Function Desproteger()
    ' Desprotege el archivo y sus hojas excepto "Listados"
    Call Desalertar
    Dim Archivo As Workbook
    Dim i As Byte
    Set Archivo = ThisWorkbook
    Archivo.Unprotect clave
    For i = 1 To Archivo.Sheets.Count
        Archivo.Sheets(i).Unprotect clave
    Next i
    Call Alertar
    Archivo.Save
End Function

Function Desalertar()
    ' Desactivar la alerta de eliminación
    Application.DisplayAlerts = False
End Function

Function Alertar()
    ' Reactivar las alertas
    Application.DisplayAlerts = True
End Function

