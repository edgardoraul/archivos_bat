Attribute VB_Name = "MeliNueva"
Option Explicit
Public RUTA As String
Public Sub EstablecerRuta()
    Dim pcName As String
    Dim workgroupName As String
    Dim objWMIService As Object
    Dim colItems As Object
    Dim objItem As Object
    
    ' Obtener nombre del equipo actual
    pcName = UCase(Environ("COMPUTERNAME"))
    
    ' Obtener Grupo de Trabajo vía WMI
    On Error Resume Next
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    Set colItems = objWMIService.ExecQuery("Select Workgroup from Win32_ComputerSystem")
    
    For Each objItem In colItems
        workgroupName = UCase(objItem.Workgroup)
    Next objItem
    On Error GoTo 0
    
    ' Evaluar condiciones
    If pcName = "EDGAR" Then
        RUTA = "D:\Web\Listados de Ventas Online\MELI"
    ElseIf workgroupName = "RERDA" Then
        RUTA = "\\EDGAR\Web\Listados de Ventas Online\MELI"
    Else
        RUTA = ""
        MsgBox "Debes estar en la red del local de Rerda para trabajar y en su grupo de trabajo", vbExclamation, "Acceso restringido"
    End If
    Debug.Print RUTA
End Sub

Public Sub GuardarCopiaConSecuencial()
    Dim fso As Object
    Dim folderObj As Object
    Dim fileObj As Object
    Dim wbOriginal As Workbook
    Dim wbTemp As Workbook
    Dim tempPath As String
    Dim finalPath As String
    Dim todayStr As String
    Dim filePrefix As String
    Dim fileName As String
    Dim currentNum As Long
    Dim maxNum As Long
    Dim numStr As String
    Dim extension As String
    Dim posDot As Long

    ' 1. Verificar/Obtener la RUTA global
    If RUTA = "" Then Call EstablecerRuta
    If RUTA = "" Then Exit Sub ' Si no hay ruta válida, detiene el proceso

    ' Verificar que la carpeta exista realmente en la red/disco
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(RUTA) Then
        MsgBox "La carpeta de destino no está accesible:" & vbCrLf & RUTA, vbCritical, "Error de Ruta"
        Exit Sub
    End If

    Set wbOriginal = ActiveWorkbook
    
    ' Definir la ruta del archivo temporal local
    tempPath = ActiveWorkbook.Path & "\MELI-Temp.xlsx"
    
    ' Salvar error si el archivo temporal anterior quedó abierto
    On Error Resume Next
    If fso.FileExists(tempPath) Then fso.DeleteFile tempPath, True
    On Error GoTo 0

    ' 2. Guardar copia del libro actual como "MELI-Temp.xlsx" sin alterar el original
    On Error Resume Next
    wbOriginal.SaveCopyAs tempPath
    
    ' 3. Abrir la copia temporal para trabajar sobre ella y cerrar el original
    Set wbTemp = Workbooks.Open(tempPath)
    wbOriginal.Close SaveChanges:=False

    ' 4. Determinar la fecha de hoy en formato AAAA-MM-DD
    todayStr = Format(Date, "yyyy-mm-dd")
    filePrefix = todayStr & ". "

    ' 5. Analizar los archivos de la carpeta RUTA para encontrar el mayor número del patrón
    maxNum = 0
    Set folderObj = fso.GetFolder(RUTA)

    For Each fileObj In folderObj.Files
        fileName = fileObj.Name
        ' Verificar si inicia con la fecha de hoy "YYYY-MM-DD. "
        If Left(fileName, Len(filePrefix)) = filePrefix Then
            ' Quitar la extensión del nombre
            posDot = InStrRev(fileName, ".")
            If posDot > 0 Then fileName = Left(fileName, posDot - 1)
            
            ' Intentar extraer el número correlativo de 5 dígitos
            ' Formato esperado: "YYYY-MM-DD. XXXXX - MELI"
            If Mid(fileName, Len(filePrefix) + 6, 8) = " - MELI" Then
                numStr = Mid(fileName, Len(filePrefix) + 1, 5)
                If IsNumeric(numStr) Then
                    currentNum = CLng(numStr)
                    If currentNum > maxNum Then maxNum = currentNum
                End If
            End If
        End If
    Next fileObj

    ' 6. Generar el nuevo nombre correlativo (aumentando en +1)
    numStr = Format(maxNum + 1, "00000")
    finalPath = RUTA & "\" & filePrefix & numStr & " - MELI.xlsx"

    ' 7. Guardar el libro temporal en la carpeta RUTA con el nombre definitivo y cerrarlo
    ' 51 = xlOpenXMLWorkbook (.xlsx)
    wbTemp.SaveAs fileName:=finalPath, FileFormat:=51
    
    MsgBox "Archivo guardado exitosamente como:" & vbCrLf & finalPath, vbInformation, "Proceso Completado"
End Sub

Sub MeliNueva()
' GENERACION DE PLANILLAS DE MELI AÑO 2027
    Call GuardarCopiaConSecuencial
End Sub
