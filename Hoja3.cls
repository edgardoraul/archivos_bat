VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Private Sub Worksheet_Activate()
    Dim respuesta As Variant
    
    Call Desproteger
    ' Desactiva las alertas para evitar mensajes al agregar/eliminar hojas
    Call Desalertar
    
    ' Mostrar el cuadro de diálogo con el mensaje, título y botones
    Sheets("Listado").Activate
    Sheets("Listado").Cells(1, 1).Activate
    respuesta = MsgBox("¿Vas a generar los rótulos?", vbYesNo, "Confirmación")
    Debug.Print respuesta

    ' Verificar la respuesta del usuario
    If respuesta <> vbYes Then
        ' El usuario hizo clic en Cancelar (No)
        Call Alertar
        Call Proteger
        Exit Sub
    End If

    ' Permite que el código continúe si la hoja "GenerarNombres" no se encuentra
    On Error Resume Next

    ' Busca la hoja llamada "GenerarNombre"
    Dim targetSheet As Worksheet
    Set targetSheet = ThisWorkbook.Sheets("GenerarNombres")
    
    ' Restablece el manejo de errores
    On Error GoTo 0

    ' Verifica si la hoja "GenerarNombre" fue encontrada
    If Not targetSheet Is Nothing Then
        Call CompletarNombres
    End If

    ' Reactiva las alertas
    Call Alertar

    ' Limpiar objeto
    Set targetSheet = Nothing
    Call Proteger
End Sub

