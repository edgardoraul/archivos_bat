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
    
    'Application.InputBox ("�Vas a borrar los r�tulos creados?, Type:=2)
    
    ' Mostrar el cuadro de di�logo con el mensaje, t�tulo y botones
    Sheets("Listado").Activate
    Sheets("Listado").Cells(1, 1).Activate
    respuesta = MsgBox("�Vas a borrar los r�tulos creados?", vbYesNo, "Confirmaci�n")

    ' Verificar la respuesta del usuario
    If respuesta = vbNo Then
        ' El usuario hizo clic en Cancelar (No)
        Exit Sub
    End If

    On Error Resume Next ' Permite que el c�digo contin�e si la hoja "GenerarNombre" no se encuentra

    ' Busca la hoja llamada "GenerarNombre"
    Dim targetSheet As Worksheet
    Set targetSheet = ThisWorkbook.Sheets("GenerarNombres")
    
    ' Borrar la planilla vieja
    Sheets("Nombres").Delete
    

    On Error GoTo 0 ' Restablece el manejo de errores

    ' Verifica si la hoja "GenerarNombre" fue encontrada
    If Not targetSheet Is Nothing Then
        ' Si se encontr�, agrega una nueva hoja ANTES de ella y la nombra "Nombre"
        ThisWorkbook.Sheets.Add Before:=targetSheet, Count:=1, Type:=xlWorksheet
        ActiveSheet.Name = "Nombres"
        
        ' Establece el color de la pesta�a a azul
        ActiveSheet.Tab.ColorIndex = 46
    Else
        ' Si no se encontr� la hoja "GenerarNombre", puedes decidir qu� hacer:
        ' Opci�n 1: No hacer nada (el c�digo actual no agregar�a la hoja si no encuentra "GenerarNombre")
        ' Opci�n 2: Agregar la hoja "Nombre" al final del libro (descomenta las siguientes 3 l�neas)
        ' ThisWorkbook.Sheets.Add After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count), Count:=1, Type:=xlWorksheet
        ' ActiveSheet.Name = "Nombre"
        ' MsgBox "La hoja 'GenerarNombre' no fue encontrada, la hoja 'Nombre' se agreg� al final." ' Esto s� mostrar�a un mensaje
        ' Opci�n 3: Agregar la hoja "Nombre" al principio del libro (descomenta las siguientes 3 l�neas)
        ' ThisWorkbook.Sheets.Add Before:=ThisWorkbook.Sheets(1), Count:=1, Type:=xlWorksheet
        ' ActiveSheet.Name = "Nombre"
        ' MsgBox "La hoja 'GenerarNombre' no fue encontrada, la hoja 'Nombre' se agreg� al inicio." ' Esto s� mostrar�a un mensaje

        ' Como la solicitud original era SIN mensajes, la Opci�n 1 (no hacer nada si no encuentra)
        ' es la que cumple con el requisito de no mostrar alertas.
        ' Si necesitas que siempre agregue la hoja "Nombres", elige la Opci�n 2 o 3 y elimina el MsgBox.
    End If

    ' Reactiva las alertas
    Call Alertar

    ' Limpiar objeto
    Set targetSheet = Nothing
    Call Proteger
    Sheets("Listado").Activate
    Sheets("Listado").Cells(1, 1).Activate
End Sub



