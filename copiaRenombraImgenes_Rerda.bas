Attribute VB_Name = "copiaRenombraImgenes_Rerda"
Option Explicit

Dim lastFolderPath As String ' Variable global para almacenar la última carpeta seleccionada
Dim newFolderName As String ' Variable global para almacenar el nuevo nombre de carpeta

Sub CopiarPegarRenombrarBorrarSubcarpetas()
    Dim origenPath As String
    Dim destinoPath As String
    Dim origenFolder As Object
    Dim destinoFolder As Object
    Dim subFolder As Object
    Dim response As Integer
    
    ' Obtener la ruta de la carpeta de origen
    With Application.FileDialog(msoFileDialogFolderPicker)
        If lastFolderPath <> "" Then
            If Right(lastFolderPath, 1) = "\" Then
                lastFolderPath = Left(lastFolderPath, Len(lastFolderPath) - 1) ' Eliminar la barra diagonal al final, si existe
            End If
            origenPath = Left(lastFolderPath, InStrRev(lastFolderPath, "\") - 1) ' Obtener la ruta un nivel arriba
            .InitialFileName = origenPath
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
    
    ' Obtener el nombre de la carpeta de origen
    newFolderName = Left(GetFolderName(origenPath), 7)
    
    ' Obtener la carpeta de destino
    destinoPath = "D:\Web\imagenes_rerda\" ' Cambia esta ruta por la ruta deseada
    
    ' Verificar si la carpeta de destino ya existe con el nuevo nombre
    If FolderExists(destinoPath & "\" & newFolderName) Then
        response = MsgBox("La carpeta con el nombre '" & newFolderName & "' ya existe en la ubicación de destino. ¿Desea reemplazarla?", vbYesNoCancel + vbExclamation, "Carpeta existente")
        If response = vbYes Then
            ' Borrar la carpeta de destino existente
            DeleteFolder destinoPath & "\" & newFolderName
        ElseIf response = vbNo Then
            ' Seleccionar otra carpeta
            CopiarPegarRenombrarBorrarSubcarpetas
            Exit Sub
        ElseIf response = vbCancel Then
            Exit Sub ' Cancelar la operación
        End If
    End If
    
    ' Copiar y pegar la carpeta de origen con el nuevo nombre
    Set origenFolder = CreateObject("Scripting.FileSystemObject").GetFolder(origenPath)
    origenFolder.Copy Destination:=destinoPath & "\" & newFolderName
    
    ' Obtener la carpeta copiada
    Set destinoFolder = CreateObject("Scripting.FileSystemObject").GetFolder(destinoPath & "\" & newFolderName)
    
    ' Borrar subcarpetas
    For Each subFolder In destinoFolder.SubFolders
        subFolder.Delete
    Next subFolder
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

