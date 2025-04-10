Attribute VB_Name = "copiaRenombraImgenes_Rerda"
Option Explicit

Dim lastFolderPath As String ' Variable global para almacenar la última carpeta seleccionada
Dim newFolderName As String ' Variable global para almacenar el nuevo nombre de carpeta

Sub CopiarPegarRenombrarBorrarSubcarpetas()
Attribute CopiarPegarRenombrarBorrarSubcarpetas.VB_ProcData.VB_Invoke_Func = "P\n14"
    Dim origenPath As String
    Dim destinoPath As String
    Dim origenFolder As Object
    Dim destinoFolder As Object
    Dim subFolder As Object
    Dim response As Integer
    Dim control As String
    
    ' Obtener la ruta de la carpeta de origen
    With Application.FileDialog(msoFileDialogFolderPicker)
        If lastFolderPath <> "" Then
            If Right(lastFolderPath, 1) = "\\" Then
                lastFolderPath = Left(lastFolderPath, Len(lastFolderPath) - 1)
            End If
            origenPath = Left(lastFolderPath, InStrRev(lastFolderPath, "\\") - 1)
            .InitialFileName = origenPath
        Else
            .InitialFileName = ActiveWorkbook.Path & "\"
        End If
        .Title = "Seleccionar carpeta de origen"
        .Show
    
        If .SelectedItems.Count = 0 Then Exit Sub
        origenPath = .SelectedItems(1)
        lastFolderPath = origenPath
    End With
    
    ' Verificar si existe 1.jpg, si no, crear una copia del primer archivo de imagen encontrado
    Call AsegurarImagen1(origenPath)
    
    ' Obtener el nombre de la carpeta de origen
    newFolderName = Left(GetFolderName(origenPath), 7)
    
    ' Carpeta de destino
    destinoPath = "D:\Web\imagenes_rerda\"
    
    ' Verificar si la carpeta de destino ya existe con el nuevo nombre
    If FolderExists(destinoPath & "\" & newFolderName) Then
        response = MsgBox("La carpeta con el nombre '" & newFolderName & "' ya existe en la ubicación de destino. ¿Desea reemplazarla?", vbYesNoCancel + vbExclamation, "Carpeta existente")
        If response = vbYes Then
            DeleteFolder destinoPath & "\" & newFolderName
        ElseIf response = vbNo Then
            CopiarPegarRenombrarBorrarSubcarpetas
            Exit Sub
        ElseIf response = vbCancel Then
            Exit Sub
        End If
    End If
    
    ' Copiar la carpeta con el nuevo nombre
    Set origenFolder = CreateObject("Scripting.FileSystemObject").GetFolder(origenPath)
    origenFolder.Copy Destination:=destinoPath & "\" & newFolderName
    
    ' Obtener la carpeta copiada
    Set destinoFolder = CreateObject("Scripting.FileSystemObject").GetFolder(destinoPath & "\" & newFolderName)
    
    ' Borrar subcarpetas
    For Each subFolder In destinoFolder.SubFolders
        subFolder.Delete
    Next subFolder
    
    control = destinoFolder
    Debug.Print destinoFolder
    Call Agregar1jpgSiNoExiste(control)
    
    
    
End Sub

Sub AsegurarImagen1(origenPath As String)
    Dim fso As Object, archivo As Object
    Dim archivoImagen As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Si ya existe 1.jpg, no hacer nada
    If fso.FileExists(origenPath & "\1.jpg") Then Exit Sub
    
    ' Buscar el primer archivo de imagen
    For Each archivo In fso.GetFolder(origenPath).Files
        If LCase(fso.GetExtensionName(archivo.Name)) = "jpg" Or _
           LCase(fso.GetExtensionName(archivo.Name)) = "jpeg" Or _
           LCase(fso.GetExtensionName(archivo.Name)) = "png" Then
            archivoImagen = archivo.Path
            Exit For
        End If
    Next archivo
    
    ' Si encontró un archivo de imagen, copiarlo como 1.jpg
    If archivoImagen <> "" Then
        fso.CopyFile archivoImagen, origenPath & "\1.jpg"
    End If
End Sub

Function FolderExists(folderPath As String) As Boolean
    If Dir(folderPath, vbDirectory) <> "" Then
        FolderExists = True
    Else
        FolderExists = False
    End If
End Function

Sub DeleteFolder(folderPath As String)
    On Error Resume Next
    Kill folderPath & "\*.*"
    RmDir folderPath
    On Error GoTo 0
End Sub

Function GetFolderName(folderPath As String) As String
    Dim folderArray() As String
    folderArray = Split(folderPath, "\")
    GetFolderName = folderArray(UBound(folderArray))
End Function


Function Agregar1jpgSiNoExiste(rutaCarpeta As String) As Boolean
    Dim fso As Object
    Dim archivo As Object
    Dim archivo1jpg As String
    Dim primerJPG As String
    
    On Error GoTo errHandler
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FolderExists(rutaCarpeta) Then
        Debug.Print "La carpeta no existe: " & rutaCarpeta
        Agregar1jpgSiNoExiste = False
        Exit Function
    End If
    
    archivo1jpg = fso.BuildPath(rutaCarpeta, "1.jpg")
    
    ' Si ya existe "1.jpg", no hacer nada
    If fso.FileExists(archivo1jpg) Then
        Agregar1jpgSiNoExiste = True
        Exit Function
    End If
    
    ' Buscar el primer .jpg (o .jpeg)
    For Each archivo In fso.GetFolder(rutaCarpeta).Files
        If Not (archivo.Attributes And 2) = 0 Then GoTo Siguiente ' Saltar ocultos
        If EsExtensionJPG(archivo.Name) Then
            primerJPG = archivo.Path
            fso.CopyFile primerJPG, archivo1jpg
            Debug.Print "Se copió: " & primerJPG & " como 1.jpg en " & rutaCarpeta
            Agregar1jpgSiNoExiste = True
            Exit Function
        End If
Siguiente:
    Next archivo

    ' Si llegó hasta acá, no encontró ninguna imagen
    Debug.Print "No se encontró ninguna imagen JPG en: " & rutaCarpeta
    Agregar1jpgSiNoExiste = False
    Exit Function

errHandler:
    Debug.Print "Error en carpeta: " & rutaCarpeta & " -> " & Err.Description
    Agregar1jpgSiNoExiste = False
End Function

Private Function EsExtensionJPG(nombreArchivo As String) As Boolean
    Dim ext As String
    ext = LCase(Right(nombreArchivo, Len(nombreArchivo) - InStrRev(nombreArchivo, ".")))
    EsExtensionJPG = (ext = "jpg" Or ext = "jpeg")
End Function

