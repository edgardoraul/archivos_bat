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
Sub MeliNueva()
' GENERACION DE PLANILLAS DE MELI AÑO 2027
    Call EstablecerRuta
End Sub
