VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Private Sub Worksheet_Activate()
    Dim respuesta As Integer
    ' quita seguridad
    Call Desproteger
    ' Desactiva las alertas para evitar mensajes al agregar/eliminar hojas
    Call Desalertar
    
    'Application.InputBox ("¿Vas a borrar los rótulos creados?, Type:=2)
    
    ' Mostrar el cuadro de diálogo con el mensaje, título y botones
    Sheets("Listado").Activate
    Sheets("Listado").Cells(1, 1).Activate
    respuesta = MsgBox("¿Vas a borrar los rótulos creados?", vbYesNo, "Confirmación")

    ' Verificar la respuesta del usuario
    If respuesta = vbNo Then
        ' El usuario hizo clic en Cancelar (No)
        Exit Sub
    End If

    On Error Resume Next ' Permite que el código continúe si la hoja "GenerarNombre" no se encuentra

    ' Busca la hoja llamada "GenerarNombre"
    Dim targetSheet As Worksheet
    Set targetSheet = ThisWorkbook.Sheets("GenerarNombres")
    
    ' Borrar la planilla vieja
    Sheets("Nombres").Delete
    

    On Error GoTo 0 ' Restablece el manejo de errores

    ' Verifica si la hoja "GenerarNombre" fue encontrada
    If Not targetSheet Is Nothing Then
        ' Si se encontró, agrega una nueva hoja ANTES de ella y la nombra "Nombre"
        ThisWorkbook.Sheets.Add Before:=targetSheet, Count:=1, Type:=xlWorksheet
        ActiveSheet.Name = "Nombres"
        
        ' Establece el color de la pestaña a azul
        ActiveSheet.Tab.ColorIndex = 46
    Else
        ' Si no se encontró la hoja "GenerarNombre", puedes decidir qué hacer:
        ' Opción 1: No hacer nada (el código actual no agregaría la hoja si no encuentra "GenerarNombre")
        ' Opción 2: Agregar la hoja "Nombre" al final del libro (descomenta las siguientes 3 líneas)
        ' ThisWorkbook.Sheets.Add After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count), Count:=1, Type:=xlWorksheet
        ' ActiveSheet.Name = "Nombre"
        ' MsgBox "La hoja 'GenerarNombre' no fue encontrada, la hoja 'Nombre' se agregó al final." ' Esto sí mostraría un mensaje
        ' Opción 3: Agregar la hoja "Nombre" al principio del libro (descomenta las siguientes 3 líneas)
        ' ThisWorkbook.Sheets.Add Before:=ThisWorkbook.Sheets(1), Count:=1, Type:=xlWorksheet
        ' ActiveSheet.Name = "Nombre"
        ' MsgBox "La hoja 'GenerarNombre' no fue encontrada, la hoja 'Nombre' se agregó al inicio." ' Esto sí mostraría un mensaje

        ' Como la solicitud original era SIN mensajes, la Opción 1 (no hacer nada si no encuentra)
        ' es la que cumple con el requisito de no mostrar alertas.
        ' Si necesitas que siempre agregue la hoja "Nombres", elige la Opción 2 o 3 y elimina el MsgBox.
    End If

    ' Reactiva las alertas
    Call Alertar

    ' Limpiar objeto
    Set targetSheet = Nothing
    Call Proteger
    Sheets("Listado").Activate
    Sheets("Listado").Cells(1, 1).Activate
End Sub



