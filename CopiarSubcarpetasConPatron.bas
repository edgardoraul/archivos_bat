Attribute VB_Name = "CopiarSubcarpetasConPatron"
Option Explicit

Dim lastFolderPath As String ' Variable global para almacenar la �ltima carpeta seleccionada

Sub CopiarSubcarpetasConPatron()
    Dim origenPath As String
    Dim destinoPath As String
    Dim origenFolder As Object
    Dim destinoFolder As Object
    Dim subFolder As Object
    Dim newFolderName As String ' Variable para almacenar el nuevo nombre de la carpeta
    Dim pattern As String
    
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
            lastFolderPath = origenPath ' Guardar la �ltima carpeta seleccionada
        End If
    End With
    
    ' Obtener la carpeta de destino
    destinoPath = "D:\Web\imagenes_rerda\" ' Cambia esta ruta por la ruta deseada
    
    ' Recorrer las subcarpetas de la carpeta de origen
    Set origenFolder = CreateObject("Scripting.FileSystemObject").GetFolder(origenPath)
    For Each subFolder In origenFolder.SubFolders
        ' Verificar si el nombre de la subcarpeta cumple con el patr�n
        pattern = "###### -" ' Patr�n: 7 caracteres num�ricos, un espacio y un gui�n medio
        If Len(subFolder.Name) = Len(pattern) Then
            If IsNumeric(Left(subFolder.Name, 7)) And Mid(subFolder.Name, 8, 1) = " " And Mid(subFolder.Name, 9, 1) = "-" Then
                ' Copiar y pegar la subcarpeta en el destino con el nuevo nombre
                newFolderName = Left(subFolder.Name, 7)
                Set destinoFolder = CreateObject("Scripting.FileSystemObject").GetFolder(destinoPath)
                CreateObject("Scripting.FileSystemObject").CopyFolder subFolder.Path, destinoFolder.Path & "\" & newFolderName, True
            End If
        End If
    Next subFolder
End Sub

