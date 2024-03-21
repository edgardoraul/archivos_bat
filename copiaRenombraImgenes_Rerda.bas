Attribute VB_Name = "copiaRenombraImgenes_Rerda"
Option Explicit
Sub CopiarPegarRenombrarBorrarSubcarpetas()
Attribute CopiarPegarRenombrarBorrarSubcarpetas.VB_ProcData.VB_Invoke_Func = "P\n14"
    Dim origenPath As String
    Dim destinoPath As String
    Dim origenFolder As Object
    Dim destinoFolder As Object
    Dim subFolder As Object
    
    ' Obtener la ruta de la carpeta de origen
    With Application.FileDialog(msoFileDialogFolderPicker)
        .InitialFileName = ActiveWorkbook.Path & "\"
        .Title = "Seleccionar carpeta"
        .Show
    
        If .SelectedItems.Count = 0 Then
            Exit Sub
        Else
            origenPath = .SelectedItems(1)
            Debug.Print origenPath
        End If

    End With
    
    ' Obtener la carpeta de destino
    destinoPath = "D:\Web\imagenes_rerda\" ' Cambia esta ruta por la ruta deseada
    
    ' Copiar y pegar la carpeta
    Set origenFolder = CreateObject("Scripting.FileSystemObject").GetFolder(origenPath)
    origenFolder.Copy Destination:=destinoPath
    
    ' Obtener la carpeta copiada
    Set destinoFolder = CreateObject("Scripting.FileSystemObject").GetFolder(destinoPath & "\" & origenFolder.Name)
    
    ' Borrar subcarpetas
    For Each subFolder In destinoFolder.subfolders
        subFolder.Delete
    Next subFolder
    
    ' Renombrar la carpeta copiada
    destinoFolder.Name = Left(destinoFolder.Name, 7)
End Sub

