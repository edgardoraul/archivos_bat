Attribute VB_Name = "CopiarSubcarpetasConPatron"
Option Explicit

Dim lastFolderPath As String ' Variable global para almacenar la última carpeta seleccionada

Sub CopiarSubcarpetasConPatron_Actualizada()
    ' Declaración de variables para mayor compatibilidad y limpieza
    Dim origenPath As String
    Dim destinoPath As String
    Dim origenFolder As Object
    Dim tempFolder As Object
    Dim subFolderOrigen As Object
    Dim tempSubFolder As Object
    Dim newFolderName As String
    
    ' Se requiere la librería "Microsoft Scripting Runtime" para FSO.
    ' Se utiliza CreateObject para compatibilidad sin necesidad de referencia.
    Dim FSO As Object
    
    ' Variables para Regex (expresiones regulares)
    Dim regex As Object
    Dim match As Object
    
    ' Variables para limpieza (archivos y subcarpetas dentro de las carpetas copiadas)
    Dim subFolder As Object
    Dim file As Object
    
    ' Variable para manejar la ruta seleccionada en el diálogo (se elimina lastFolderPath)
    Dim selectedPath As String

    ' --- 1. Seleccionar carpeta de origen ---
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Seleccionar carpeta de origen"
        ' Inicia en la ruta del libro activo si no hay una ruta previa (mejora de tu código)
        .InitialFileName = ActiveWorkbook.Path & "\"
        
        If .Show = -1 Then ' Si el usuario selecciona una carpeta
            origenPath = .SelectedItems(1)
            Debug.Print "Origen: " & origenPath
        Else ' Si el usuario cancela
            MsgBox "Operación cancelada por el usuario.", vbInformation
            Exit Sub
        End If
    End With
    
    ' --- 2. Definir rutas de destino y preparar FSO ---
    destinoPath = "D:\Web\imagenes_rerda\" ' Carpeta final de destino
    Dim tempPath As String
    tempPath = destinoPath & "Temp\" ' Ruta de la carpeta temporal
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    ' Asegurar que la carpeta de destino final exista
    If Not FSO.FolderExists(destinoPath) Then
        FSO.CreateFolder destinoPath
        Debug.Print "Carpeta de destino creada: " & destinoPath
    End If
    
    ' Crear o limpiar la carpeta temporal
    If FSO.FolderExists(tempPath) Then
        FSO.DeleteFolder tempPath, True ' Borrar el contenido, True = Forzar borrado
        Debug.Print "Carpeta temporal previa borrada."
    End If
    FSO.CreateFolder tempPath ' Crear la carpeta temporal
    Set tempFolder = FSO.GetFolder(tempPath)
    Debug.Print "Carpeta temporal creada: " & tempPath
    
    ' --- 3. Inicializar expresión regular ---
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = False
    regex.IgnoreCase = True
    regex.pattern = "\d{7}" ' Patrón: 7 dígitos numéricos
    
    ' --- 4. Recorrer, Filtrar y Limpiar Subcarpetas ---
    Set origenFolder = FSO.GetFolder(origenPath)
    For Each subFolderOrigen In origenFolder.SubFolders
        Debug.Print "Procesando: " & subFolderOrigen.Name
        
        ' Verificar si el nombre de la subcarpeta cumple con el patrón
        If regex.Test(subFolderOrigen.Name) Then
            Set match = regex.Execute(subFolderOrigen.Name)
            If match.Count > 0 Then
                newFolderName = match(0).Value ' Extraer el patrón (7 dígitos)
                
                ' **Copiar la subcarpeta a la carpeta temporal. El 'True' al final es para sobrescribir.**
                ' Aunque es una carpeta temporal que acabamos de crear, esto es una buena práctica.
                On Error Resume Next ' Manejar posible error de ruta o permisos durante la copia
                FSO.CopyFolder subFolderOrigen.Path, tempPath & newFolderName, True
                On Error GoTo 0 ' Reactivar el manejo de errores
                
                Debug.Print "Copiado a temporal como: " & newFolderName
                
                ' Obtener la carpeta recién copiada en la carpeta temporal
                Set tempSubFolder = Nothing ' Limpiar la variable
                Set tempSubFolder = FSO.GetFolder(tempPath & newFolderName)
                
                ' Borrar subcarpetas dentro de la carpeta recién copiada (limpieza)
                For Each subFolder In tempSubFolder.SubFolders
                    subFolder.Delete True ' Borrado forzado
                Next subFolder
                
                ' Eliminar el archivo Thumbs.db dentro de la carpeta recién copiada (limpieza)
                For Each file In tempSubFolder.Files
                    If UCase(file.Name) = "THUMBS.DB" Then ' Usar UCase para mejor comparación
                        FSO.DeleteFile file.Path, True ' Borrado forzado
                    End If
                Next file
            End If
        End If
    Next subFolderOrigen
    
    ' --- 5. Copiar de la carpeta temporal a la carpeta final de destino (Reemplazo) ---
    For Each tempSubFolder In tempFolder.SubFolders
        ' **EL REEMPLAZO SE GARANTIZA AQUÍ:**
        ' FSO.CopyFolder con 'True' sobrescribe el destino si ya existe.
        ' Si la carpeta destinoPath & "\" & tempSubFolder.Name existe, se sobrescribe.
        FSO.CopyFolder tempSubFolder.Path, destinoPath, True
        Debug.Print "Copiado a destino final (sobrescrito): " & tempSubFolder.Name
    Next tempSubFolder
    
    ' --- 6. Borrar la carpeta temporal ---
    FSO.DeleteFolder tempFolder.Path, True ' Borrado forzado
    
    ' --- 7. Finalizar
    Set FSO = Nothing
    Set regex = Nothing
    MsgBox "Proceso de copiado y reemplazo completado.", vbInformation
End Sub


Sub RenombrarCarpetasPorCodigo_Limpio()
Attribute RenombrarCarpetasPorCodigo_Limpio.VB_ProcData.VB_Invoke_Func = "O\n14"
    Dim CarpetaBase As String
    Dim FD As FileDialog
    Dim FSO As Object, Carpeta As Object
    Dim ListaNombres As Collection
    Dim nombre As Variant
    Dim UltFila As Long
    Dim codigo As String
    Dim celda As Range
    Dim baseNombre As String, nuevoNombre As String
    Dim rutaVieja As String, rutaNueva As String
    
    ' Elegir carpeta base
    Set FD = Application.FileDialog(msoFileDialogFolderPicker)
    With FD
        .Title = "Selecciona la carpeta que contiene las subcarpetas"
        If .Show <> -1 Then Exit Sub
        CarpetaBase = .SelectedItems(1)
    End With
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set Carpeta = FSO.GetFolder(CarpetaBase)
    
    ' Tomar la lista de subcarpetas primero (para no alterar la colección mientras renombramos)
    Set ListaNombres = New Collection
    Dim sf As Object
    For Each sf In Carpeta.SubFolders
        ListaNombres.Add sf.Name
    Next sf
    
    UltFila = Cells(Rows.Count, 1).End(xlUp).Row
    
    For Each nombre In ListaNombres
        rutaVieja = CarpetaBase & "\" & CStr(nombre)
        ' Si la subcarpeta fue movida/eliminada entre medio, saltar
        If Len(Dir(rutaVieja, vbDirectory)) = 0 Then GoTo Siguiente
        
        ' Tomar los primeros 7 caracteres del nombre de la carpeta
        codigo = Left$(CStr(nombre), 7)
        
        ' Buscar el código en la columna A (coincidencia exacta, sin distinguir may/min)
        Set celda = Range("F1:F" & UltFila).Find(What:=codigo, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
        If Not celda Is Nothing Then
            ' Construir el nombre base A & " - " & B (si B está vacío, sólo A)
            If Trim$(CStr(celda.Offset(0, -1).Value)) <> "" Then
                baseNombre = CStr(celda.Value) & " - " & CStr(celda.Offset(0, -1).Value)
            Else
                baseNombre = CStr(celda.Value)
            End If
            
            ' Limpiar y normalizar (permitir letras, números, punto, guion medio, guion bajo y un solo espacio)
            nuevoNombre = LimpiarYNormalizarNombre(baseNombre, codigo)
            
            ' Si el nombre resultante es igual al actual (tras limpiar), no hace falta renombrar
            If StrComp(nuevoNombre, CStr(nombre), vbTextCompare) = 0 Then GoTo Siguiente
            
            ' Asegurar unicidad si ya existe
            rutaNueva = RutaUnica(CarpetaBase, nuevoNombre)
            
            On Error Resume Next
            Name rutaVieja As rutaNueva
            If Err.Number <> 0 Then
                ' Si algo falla, continuar sin frenar
                Debug.Print "No se pudo renombrar: " & rutaVieja & " -> " & rutaNueva & " | Error: " & Err.Number & " - " & Err.Description
                Err.Clear
            End If
            On Error GoTo 0
        End If
Siguiente:
        ' continúa con la siguiente subcarpeta
    Next nombre
    
    MsgBox "Proceso terminado.", vbInformation
End Sub

' --- Limpia caracteres y normaliza espacios ---
' Permite: letras, números, punto (.), guion (-), guion bajo (_), y espacio.
' Convierte múltiples espacios a uno y recorta extremos.
' Elimina cualquier carácter extraño (\, /, :, *, ?, ", <, >, |, comas, etc.).
' Si el resultado queda vacío, usa el código de 7 caracteres como respaldo.
Function LimpiarYNormalizarNombre(ByVal s As String, ByVal respaldoCodigo As String) As String
    Dim rx As Object
    
    ' 1) Reemplazar caracteres NO permitidos por nada
    '    Permitidos: A-Z a-z 0-9 . - _ espacio
    Set rx = CreateObject("VBScript.RegExp")
    rx.Global = True
    rx.IgnoreCase = True
    rx.pattern = "[^A-Za-z0-9.\-_ ]"
    s = rx.Replace(s, "")
    
    ' 2) Colapsar espacios múltiples a uno solo
    rx.pattern = "\s+"
    s = rx.Replace(s, " ")
    
    ' 3) Recortar espacios extremos
    s = Trim$(s)
    
    ' 4) Quitar puntos/espacios al final y al inicio (Windows no permite terminar con espacio o punto)
    Do While Len(s) > 0 And (Right$(s, 1) = "." Or Right$(s, 1) = " ")
        s = Left$(s, Len(s) - 1)
    Loop
    Do While Len(s) > 0 And (Left$(s, 1) = "." Or Left$(s, 1) = " ")
        s = Mid$(s, 2)
    Loop
    
    ' 5) Evitar nombres reservados de Windows
    If EsNombreReservado(s) Then s = s & "_"
    
    ' 6) Truncar por seguridad (evitar rutas larguísimas)
    If Len(s) > 200 Then s = Left$(s, 200)
    
    ' 7) Respaldo si quedó vacío
    If Len(s) = 0 Then
        s = respaldoCodigo
        If Len(s) = 0 Then s = "Carpeta"
    End If
    
    LimpiarYNormalizarNombre = s
End Function

' --- Verifica nombres reservados Windows ---
Function EsNombreReservado(ByVal n As String) As Boolean
    Dim r As Variant
    For Each r In Array("CON", "PRN", "AUX", "NUL", _
                        "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9", _
                        "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9")
        If StrComp(n, r, vbTextCompare) = 0 Then
            EsNombreReservado = True
            Exit Function
        End If
    Next r
End Function

' --- Genera una ruta única si ya existe una carpeta con ese nombre ---
Function RutaUnica(ByVal base As String, ByVal nombre As String) As String
    Dim ruta As String, i As Long
    ruta = base & "\" & nombre
    If Len(Dir(ruta, vbDirectory)) = 0 Then
        RutaUnica = ruta
        Exit Function
    End If
    
    i = 1
    Do
        ruta = base & "\" & nombre & " (" & i & ")"
        i = i + 1
    Loop While Len(Dir(ruta, vbDirectory)) <> 0
    RutaUnica = ruta
End Function


