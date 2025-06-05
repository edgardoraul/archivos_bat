Attribute VB_Name = "Faltantes"
Public Const pass As String = "Rerda2025"
Public ultimaConDatos As Integer
Public ultimaDerecha As Integer
Public Const naranja As String = 40
Public Respuesta As Variant

Option Explicit
' Desalertar
Function NoAlertas()
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
End Function

' Alertar
Function SiAlertas()
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Function

' PROTEGER
Function proteger()
    Dim ws As Worksheet
    If Sheets("CODIGOS").Visible = True Then
        Sheets("CODIGOS").Visible = False
    End If
    
    If Sheets("VARIANTES").Visible = True Then
        Sheets("VARIANTES").Visible = False
    End If
    For Each ws In ThisWorkbook.Worksheets
        ws.Protect pass
    Next ws
    ThisWorkbook.Protect pass
End Function

' DESPROTEGER
Function desproteger()
    ThisWorkbook.Unprotect pass
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ws.Unprotect pass
    Next
End Function

' Advertencia de borrado
Function Advertencia()

    ' Mostrar el cuadro de di�logo con Aceptar y Cancelar
    Respuesta = MsgBox("�Cuidado!" & vbNewLine & "Vas a borrar la informaci�n cargada.", vbYesNo, "Confirmaci�n")
    Debug.Print Respuesta
    ' Analizar la respuesta
    If Respuesta = vbYes Then ' El usuario ha seleccionado "Aceptar"
    
        ' Opcional: MsgBox "Has seleccionado Aceptar. El proceso de borrado continuar�.", vbInformation
        Debug.Print Respuesta & ". Aceptasteeeee!!!"
        
    Else ' El usuario ha seleccionado "Cancelar"
    
        ' Opcional: MsgBox "Has seleccionado Cancelar. El proceso de borrado se detendr�.", vbInformation
        Debug.Print Respuesta & ". �Chau! Cancelaste."
    End If
    
End Function

' Generando las �ltimas celdas con datos.
Function ultima()
    ultimaConDatos = Cells(Rows.Count, 1).End(xlUp).Row
    ultimaDerecha = Cells(4, Columns.Count).End(xlToLeft).Column
    Debug.Print "Ultima fila: " & ultimaConDatos
    Debug.Print "Ultima columna: " & ultimaDerecha
    ThisWorkbook.Save
End Function

Sub Marcar()
Attribute Marcar.VB_ProcData.VB_Invoke_Func = "M\n14"
    ' Ctrol +  May + M
    ' MARCA LAS CELDAS CON COLOR
    ' SOLO LAS CELDAS MARCADAS PUEDEN SUMARSE
    Call NoAlertas
    Sheets("LISTADO").Unprotect pass
    Call ultima
    If ActiveCell.Column >= 5 And ActiveCell.Column <= ultimaDerecha And ActiveCell.Row > 4 Then
        
        If ActiveCell.Interior.ColorIndex = naranja Or Selection.Interior.ColorIndex = naranja Then
            ActiveCell.Interior.Color = xlNone
            Selection.Interior.ColorIndex = xlNone
        Else
            ActiveCell.Interior.ColorIndex = naranja
            Selection.Interior.ColorIndex = naranja
        End If
    End If
    Sheets("LISTADO").Protect pass
    Call SiAlertas
End Sub

' Inserta una persona en la fila N� 5.
Sub InsPersona()
    Call NoAlertas
    Sheets("LISTADO").Unprotect pass
    Rows(5).Insert Shift:=xlShiftDown, CopyOrigin:=xlFormatFromRightOrBelow
    Cells(5, 1).Activate
    Call ultima
    Cells(ultimaConDatos, 2).Formula = "=COUNTA(" & "B5:B" & ultimaConDatos - 1 & ")"
    Call ultima
    Sheets("LISTADO").Protect pass
    Call SiAlertas
    ThisWorkbook.Save
End Sub

' Borra una fila, siempre la N� 5. Deja una solita, nada m�s.
Sub BorrarPersona()
    Sheets("LISTADO").Unprotect pass
    
    ' Se advierte sobre el borrado. Se sale si se cancela.
    Call Advertencia
    Call NoAlertas
    If Respuesta <> vbYes Then Exit Sub
    
    Call ultima
    If ultimaConDatos > 6 Then
        Rows(5).Delete
        Cells(5, 1).Activate
    Else
        MsgBox "No se puede borrar esta fila."
    End If
    Call ultima
    Sheets("LISTADO").Protect pass
    Call SiAlertas
    ThisWorkbook.Save
End Sub

' Inserta un producto: son 3 columnas, con su formato, f�rmula y restricciones.
Sub InsProducto()
    Sheets("LISTADO").Unprotect pass
    Call NoAlertas
    Call ultima
    Columns("E:G").Select
    Selection.Insert CopyOrigin:=xlFormatFromRightOrBelow
    Range(Cells(2, 8), Cells(ultimaConDatos - 1, 10)).Select
    Selection.Copy
    Range("E2").Activate
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Columns("H:J").Select
    Columns("H:J").Copy
    Columns("E:G").Select
    Selection.PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
    Range("E2").Activate
    Range("E2").Value = ""
    Range(Cells(5, 5), Cells(ultimaConDatos - 1, 7)).Value = ""
    Range("E2").Select
    Call ultima
    Sheets("LISTADO").Protect pass
    Call SiAlertas
    ThisWorkbook.Save
End Sub

' Borra un producto (grupo de 3 columnas), siempre las columnas E a G.
Sub RemProducto()
    Sheets("LISTADO").Unprotect pass
    Call ultima
    
    ' Advirtiendo sobre el borrado. Se sale si se cancela
    Call Advertencia
    Call NoAlertas
    If Respuesta <> vbYes Then Exit Sub
    If ultimaDerecha < 8 Then
        MsgBox "No se puede eliminar el �ltimo producto."
        Range("E2").Select
        ActiveWorkbook.Save
        Exit Sub
    End If
    Columns("E:G").Select
    Selection.Delete
    Range("E2").Select
    Call ultima
    Sheets("LISTADO").Protect pass
    Call SiAlertas
    ThisWorkbook.Save
End Sub

' Funci�n auxiliar para limpiar nombres de hojas (privada a este m�dulo)
Private Function CleanSheetName(sIn As String) As String
    Dim sOut As String
    sOut = sIn
    sOut = Replace(sOut, "/", "_")
    sOut = Replace(sOut, "\", "_")
    sOut = Replace(sOut, "*", "_")
    sOut = Replace(sOut, "?", "_")
    sOut = Replace(sOut, "[", "_")
    sOut = Replace(sOut, "]", "_")
    sOut = Replace(sOut, ":", "_")
    ' Podr�as a�adir m�s reemplazos si es necesario
    If Len(sOut) > 31 Then sOut = Left(sOut, 31) ' Asegurar longitud m�xima
    CleanSheetName = sOut
End Function


Sub Faltantes()
    Dim ws As Worksheet
    Dim shFaltantes As Worksheet
    Dim shPlanilla As Worksheet
    Dim productSheet As Worksheet
    Dim productCode As String
    Dim j_col As Long ' �ndice de columna para productos
    Dim lastProductColPlanilla As Long
    Dim i_sheetCounter As Long ' Para bucle de borrado de hojas

    Dim preservedSheetNames As Object
    Set preservedSheetNames = CreateObject("Scripting.Dictionary")

    ' Configuraci�n inicial y desprotecci�n
    Call NoAlertas ' Tu funci�n (DisplayAlerts = False, Calculation = xlCalculationManual)
    Application.ScreenUpdating = False ' Manejar expl�citamente para evitar parpadeo
    
    On Error GoTo Faltantes_ErrorHandler ' Establecer manejador de errores para esta subrutina

    Call desproteger ' Tu funci�n para desproteger todo

    ' --- INICIO DE LA NUEVA L�GICA ---

    ' 1. Eliminar hojas excepto las preservadas
    preservedSheetNames(UCase("LISTADO")) = True
    preservedSheetNames(UCase("VARIANTES")) = True
    preservedSheetNames(UCase("CODIGOS")) = True

    For i_sheetCounter = ThisWorkbook.Worksheets.Count To 1 Step -1 ' Iterar hacia atr�s al eliminar
        Set ws = ThisWorkbook.Worksheets(i_sheetCounter)
        If Not preservedSheetNames.Exists(UCase(ws.Name)) Then
            ws.Delete ' DisplayAlerts ya est� en False
        End If
    Next i_sheetCounter
    Set ws = Nothing ' Limpiar referencia

    ' 2. Crear hoja "FALTANTES"
    Set shFaltantes = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    On Error Resume Next ' Por si el nombre "FALTANTES" ya existe de alguna manera (improbable aqu�)
    shFaltantes.Name = "FALTANTES"
    If Err.Number <> 0 Then
        MsgBox "Error cr�tico: No se pudo nombrar la hoja como 'FALTANTES'." & vbNewLine & _
               "Puede que ya exista una hoja protegida o con un nombre problem�tico." & vbNewLine & _
               "Error: " & Err.Description, vbCritical, "Error al Crear Hoja FALTANTES"
        Err.Clear
        GoTo Faltantes_Cleanup ' Salir si no se puede crear FALTANTES
    End If
    On Error GoTo Faltantes_ErrorHandler ' Restaurar manejador de errores principal de esta Sub

    ' 3. Identificar c�digos de producto desde "PLANILLA" y crear hojas de producto
    On Error Resume Next ' Intentar obtener la hoja "PLANILLA"
    Set shPlanilla = ThisWorkbook.Sheets("LISTADO")
    On Error GoTo Faltantes_ErrorHandler ' Restaurar manejador de errores principal

    If shPlanilla Is Nothing Then
        MsgBox "La hoja 'LISTADO' no fue encontrada. No se pueden crear hojas de producto.", vbExclamation
        ' El resto de la l�gica de esta secci�n se omitir� si shPlanilla no existe.
        ' El MsgBox al final indicar� que la l�gica de llenado est� pendiente.
    Else
        ' Determinar la �ltima columna con c�digos de producto en la Fila 2 de "PLANILLA"
        If IsEmpty(shPlanilla.Cells(2, shPlanilla.Columns.Count).Value) And shPlanilla.Cells(2, 1).Value = "" Then
             ' Si la �ltima celda de la fila 2 est� vac�a Y la primera tambi�n, la fila podr�a estar vac�a
            If shPlanilla.Cells(2, shPlanilla.Columns.Count).End(xlToLeft).Column = 1 And IsEmpty(shPlanilla.Cells(2, 1).Value) Then
                lastProductColPlanilla = 0 ' Indica que la fila est� esencialmente vac�a o solo datos en col A
            Else
                lastProductColPlanilla = shPlanilla.Cells(2, shPlanilla.Columns.Count).End(xlToLeft).Column
            End If
        Else
            lastProductColPlanilla = shPlanilla.Cells(2, shPlanilla.Columns.Count).End(xlToLeft).Column
        End If

        If lastProductColPlanilla < 5 Then
            Debug.Print "No se encontraron c�digos de producto en la Fila 2 de 'PLANILLA' a partir de la Columna E."
        Else
            For j_col = 5 To lastProductColPlanilla Step 3 ' Desde Col E, cada 3 columnas
                productCode = Trim(CStr(shPlanilla.Cells(2, j_col).Value))

                If productCode <> "" Then
                    Dim tempSheetNameAttempt As String
                    tempSheetNameAttempt = Left(productCode, 31) ' Nombre de hoja tentativo

                    Set productSheet = Nothing ' Reiniciar para cada intento
                    On Error Resume Next     ' Activar manejo de errores para la creaci�n y nombrado de hoja
                    Set productSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
                    productSheet.Name = tempSheetNameAttempt
                    
                    If Err.Number <> 0 Then ' Si fall� el primer intento de nombrar
                        Err.Clear
                        Dim altSheetName As String
                        altSheetName = "Prod_" & CleanSheetName(productCode) ' Usar CleanSheetName y prefijo
                        If Len(altSheetName) > 31 Then altSheetName = Left(altSheetName, 31)
                        
                        ' Si productSheet no se pudo crear en el Add, este Name fallar�.
                        ' Asumimos que Add funcion� pero el Name inicial fall�.
                        If Not productSheet Is Nothing Then
                           productSheet.Name = altSheetName
                        Else ' Esto ser�a muy raro si el Add fall� silenciosamente
                           Debug.Print "Error al A�ADIR hoja para producto: " & productCode & ". Se omitir�."
                           GoTo SiguienteCodigoProducto ' Saltar a la siguiente iteraci�n de j_col
                        End If

                        If Err.Number <> 0 Then ' Si a�n falla con nombre alternativo
                            Err.Clear
                            If Not productSheet Is Nothing Then
                                productSheet.Name = "HojaAuto" & Format(Now, "HHmmss") & "_" & j_col
                            End If
                        End If
                    End If
                    On Error GoTo Faltantes_ErrorHandler ' Restaurar manejador principal

                    If Not productSheet Is Nothing Then
                        Debug.Print "Creada hoja de producto: '" & productSheet.Name & "' para c�digo: '" & productCode & "'"
                        ' Aqu� es donde, m�s adelante, llenar�as esta productSheet.
                        ' Ejemplo: productSheet.Cells(1, 1).Value = "Datos para Producto: " & productCode
                    Else
                        Debug.Print "No se pudo crear la hoja para el c�digo de producto: '" & productCode & "'"
                    End If
                End If
SiguienteCodigoProducto:
            Next j_col
        End If ' End If lastProductColPlanilla < 5
    End If ' End If shPlanilla Is Nothing
    Set shPlanilla = Nothing ' Limpiar referencia
    Set productSheet = Nothing
    
    ' --- FIN DE LA NUEVA L�GICA ---

    ' Las siguientes l�neas son de tu estructura original para Faltantes:
    Call ultima ' Tu funci�n: establece ultimaConDatos y ultimaDerecha (basado en ActiveSheet)
    
    If ultimaConDatos > 0 Then
        Dim wsActivar As Worksheet
        Set wsActivar = Nothing ' Asegurar que est� limpio
        On Error Resume Next
        Set wsActivar = ActiveSheet ' La hoja activa puede ser la �ltima creada.
        On Error GoTo Faltantes_ErrorHandler

        If Not wsActivar Is Nothing Then
            Debug.Print "Intentando activar celda (" & ultimaConDatos & ", 1) en la hoja: '" & wsActivar.Name & "'"
            On Error Resume Next ' Error local para la activaci�n
            wsActivar.Cells(ultimaConDatos, 1).Activate
            If Err.Number <> 0 Then
                Debug.Print "Advertencia: No se pudo activar la celda. Error: " & Err.Description
                Err.Clear
            End If
            On Error GoTo Faltantes_ErrorHandler
        Else
            Debug.Print "No hay hoja activa para la operaci�n .Activate post 'ultima'."
        End If
        Set wsActivar = Nothing
    Else
        Debug.Print "'ultimaConDatos' es 0 despu�s de llamar a 'ultima', no se activ� celda."
    End If
    
    ' ...aqu� ir� el resto del c�digo....
    MsgBox "Estructura de hojas base creada." & vbNewLine & _
           "La l�gica para llenar estas hojas con datos se implementar� a continuaci�n.", vbInformation, "Siguientes Pasos"
    
Faltantes_Cleanup:
    Call proteger ' Tu funci�n
    
    ' En lugar de llamar a SiAlertas, restauramos expl�citamente lo que NoAlertas cambi�
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True ' Asegurar que la actualizaci�n de pantalla se restaura

    ' Limpieza final de objetos
    Set preservedSheetNames = Nothing
    Set shFaltantes = Nothing
    ' Las variables globales como ultimaConDatos no se "Set Nothing"
    Exit Sub

Faltantes_ErrorHandler:
    MsgBox "Error en Sub Faltantes: " & Err.Number & " - " & Err.Description, vbCritical, "Error en Macro Faltantes"
    ' Considera si quieres intentar la limpieza incluso despu�s de un error
    Resume Faltantes_Cleanup
End Sub

