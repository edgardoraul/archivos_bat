Attribute VB_Name = "CopiarSubcarpetasConPatron"
Option Explicit

Dim lastFolderPath As String ' Variable global para almacenar la �ltima carpeta seleccionada

Sub CopiarSubcarpetasConPatron()
Attribute CopiarSubcarpetasConPatron.VB_ProcData.VB_Invoke_Func = " \n14"
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
    Dim subFolder As Object ' Definici�n de la variable subFolder
    Dim fso As Object
    Dim file As Object
    Dim tempFolder As Object
    Dim tempSubFolder As Object
    
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
    
    ' Obtener la carpeta de destino temporal
    destinoPath = "D:\Web\imagenes_rerda\" ' Carpeta final de destino
    
    ' Crear la carpeta temporal si no existe
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(destinoPath & "..\Temp") Then
        fso.CreateFolder destinoPath & "..\Temp"
    End If
    
    ' Obtener la carpeta de destino temporal
    Set tempFolder = fso.GetFolder(destinoPath & "..\Temp")
    
    ' Inicializar expresi�n regular para buscar el patr�n en los nombres de las carpetas
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = False
    regex.IgnoreCase = True
    regex.pattern = "\d{7}" ' Patr�n: 7 d�gitos num�ricos, espacio, gui�n medio, espacio
    
    ' Recorrer las subcarpetas de la carpeta de origen
    Set origenFolder = fso.GetFolder(origenPath)
    For Each subFolderOrigen In origenFolder.SubFolders
        Debug.Print subFolderOrigen.Name
        ' Verificar si el nombre de la subcarpeta cumple con el patr�n utilizando expresiones regulares
        If regex.Test(subFolderOrigen.Name) Then
            ' Copiar y pegar la subcarpeta en la carpeta temporal con el nuevo nombre
            newFolderName = regex.Execute(subFolderOrigen.Name)(0)
            fso.CopyFolder subFolderOrigen.Path, tempFolder.Path & "\" & newFolderName, True
            
            ' Obtener la carpeta reci�n copiada en la carpeta temporal
            Set tempSubFolder = fso.GetFolder(tempFolder.Path & "\" & newFolderName)
            
            ' Borrar subcarpetas dentro de la carpeta reci�n copiada
            For Each subFolder In tempSubFolder.SubFolders
                subFolder.Delete
            Next subFolder
            
            ' Eliminar el archivo Thumbs.db dentro de la carpeta reci�n copiada
            For Each file In tempSubFolder.Files
                If file.Name = "Thumbs.db" Then
                    fso.DeleteFile file.Path
                End If
            Next file
        End If
    Next subFolderOrigen
    
    ' Copiar las carpetas de la carpeta temporal a la carpeta final de destino
    For Each tempSubFolder In tempFolder.SubFolders
        fso.CopyFolder tempSubFolder.Path, destinoPath, True
        Debug.Print tempSubFolder.Name
    Next tempSubFolder
    
    ' Borrar la carpeta temporal
    fso.DeleteFolder tempFolder.Path
End Sub

