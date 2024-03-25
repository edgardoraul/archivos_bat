Attribute VB_Name = "CopiarSubcarpetasConPatron"
Option Explicit

Dim lastFolderPath As String ' Variable global para almacenar la última carpeta seleccionada

Sub CopiarSubcarpetasConPatron()
    Dim origenPath As String
    Dim destinoPath As String
    Dim origenFolder As Object
    Dim destinoFolder As Object
    Dim subFolderOrigen As Object
    Dim subFolderDestino As Object
    Dim newFolderName As String ' Variable para almacenar el nuevo nombre de la carpeta
    Dim pattern As String
    Dim regex As Object
    Dim matches As Object
    Dim match As Object
    Dim subFolder As Object ' Definición de la variable subFolder
    Dim fso As Object
    Dim file As Object
    
    ' Obtener la ruta de la carpeta de origen
    With Application.FileDialog(msoFileDialogFolderPicker)
        If lastFolderPath <> "" Then
            If Right(lastFolderPath, 1) = "\" Then
                lastFolderPath = Left(lastFolderPath, Len(lastFolderPath) - 1) ' Eliminar la barra diagonal al final, si existe
            End If
            origenPath = lastFolderPath
        Else
            .InitialFileName = ActiveWorkbook.Path & "\"
        End If
        .Title = "Seleccionar carpeta de origen"
        .Show
    
        If .SelectedItems.Count = 0 Then
            Exit Sub
        Else
            origenPath = .SelectedItems(1)
            Debug.Print origenPath
            lastFolderPath = origenPath ' Guardar la última carpeta seleccionada
        End If
    End With
    
    ' Obtener la carpeta de destino
    destinoPath = "D:\Web\imagenes_rerda\" ' Cambia esta ruta por la ruta deseada
    
    ' Inicializar expresión regular para buscar el patrón en los nombres de las carpetas
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = False
    regex.IgnoreCase = True
    regex.pattern = "\d{7}" ' Patrón: 7 dígitos numéricos, espacio, guión medio, espacio
    
    ' Recorrer las subcarpetas de la carpeta de origen
    Set origenFolder = CreateObject("Scripting.FileSystemObject").GetFolder(origenPath)
    For Each subFolderOrigen In origenFolder.SubFolders
        Debug.Print subFolderOrigen.Name
        ' Verificar si el nombre de la subcarpeta cumple con el patrón utilizando expresiones regulares
        If regex.Test(subFolderOrigen.Name) Then
            ' Copiar y pegar la subcarpeta en el destino con el nuevo nombre
            newFolderName = regex.Execute(subFolderOrigen.Name)(0)
            Set destinoFolder = CreateObject("Scripting.FileSystemObject").GetFolder(destinoPath)
            CreateObject("Scripting.FileSystemObject").CopyFolder subFolderOrigen.Path, destinoPath & newFolderName, True
            
            ' Obtener la carpeta recién copiada en el destino
            Set subFolderDestino = FindSubFolder(destinoFolder, newFolderName)
            
            ' Borrar subcarpetas dentro de la carpeta recién copiada
            If Not subFolderDestino Is Nothing Then
                For Each subFolder In subFolderDestino.SubFolders
                    subFolder.Delete
                Next subFolder
            End If
            
            ' Eliminar el archivo Thumbs.db dentro de la carpeta recién copiada
            Set fso = CreateObject("Scripting.FileSystemObject")
            For Each file In subFolderDestino.Files
                If file.Name = "Thumbs.db" Then
                    fso.DeleteFile file.Path
                End If
            Next file
        End If
    Next subFolderOrigen
End Sub

Function FindSubFolder(parentFolder As Object, folderName As String) As Object
    Dim subFolder As Object
    For Each subFolder In parentFolder.SubFolders
        If subFolder.Name = folderName Then
            Set FindSubFolder = subFolder
            Exit Function
        End If
    Next subFolder
    Set FindSubFolder = Nothing
End Function

