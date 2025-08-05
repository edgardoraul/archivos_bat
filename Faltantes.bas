Attribute VB_Name = "Faltantes"
Public Const pass As String = "Rerda2025"
Public ultimaConDatos As Integer
Public ultimaDerecha As Integer
Public Const naranja As String = 40
Public Respuesta As Variant
Public Talle As Variant
Public Color As Variant
Public Cantidad As Integer
Public Producto As Object
Public ordenTalles As Variant
Public ultimaResumen As Integer
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
    If Worksheets("CODIGOS").Visible = True Then
        Worksheets("CODIGOS").Visible = False
    End If
    
    If Worksheets("VARIANTES").Visible = True Then
        Worksheets("VARIANTES").Visible = False
    End If
    For Each ws In ThisWorkbook.Worksheets

        ws.Protect pass, AllowFiltering:=True, AllowSorting:=True

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

    ' Mostrar el cuadro de diálogo con Aceptar y Cancelar
    Respuesta = MsgBox("¡Cuidado!" & vbNewLine & "Vas a borrar la información cargada.", vbYesNo, "Confirmación")
    
End Function

' Generando las últimas celdas con datos.
Function ultima()
    ultimaConDatos = Worksheets("LISTADO").Cells(Rows.Count, 1).End(xlUp).row - 4
    ultimaDerecha = Worksheets("LISTADO").Cells(4, Columns.Count).End(xlToLeft).Column
End Function

Function Activar(Fila)
    ' ACTIVA LA FILA DE UN CANDIDATO
    
    Call NoAlertas
    ActiveSheet.Unprotect pass
    With Range(Cells(Fila, 1), Cells(Fila, ultimaDerecha))
        .Locked = False
        .Font.ColorIndex = xlAutomatic
    End With
    ActiveSheet.Protect pass
    Call SiAlertas
End Function

Function Desactivar(Fila)
    ' BLOQUEA UNA FILA DE UN CANDIDATO PARA QUE NO SEA TENIDA EN CUENTA
    
    Call NoAlertas
    ActiveSheet.Unprotect pass
    With Range(Cells(Fila, 1), Cells(Fila, ultimaDerecha))
        .Font.ColorIndex = 48
    End With
    
    Range(Cells(Fila, 2), Cells(Fila, ultimaDerecha)).Locked = True
    
    ActiveSheet.Protect pass
    Call SiAlertas
End Function

Sub Filtrar()
    ' HABILITA EL FILTRADO Y ORDENAMIENTO DE DATOS DE LA PLANILLA
    Call ultima
    ActiveSheet.Unprotect pass
    With Range(Cells(4, 1), Cells(4, ultimaDerecha))
        .Locked = False
        .Interior.Color = xlAutomatic
        .AutoFilter
    End With
End Sub

Sub Marcar()
Attribute Marcar.VB_ProcData.VB_Invoke_Func = "M\n14"
    ' Ctrol +  May + M
    ' MARCA LAS CELDAS CON COLOR
    ' SOLO LAS CELDAS MARCADAS PUEDEN SUMARSE
    Call NoAlertas
    Worksheets("LISTADO").Unprotect pass
    Call ultima
    If ActiveCell.Column >= 5 And ActiveCell.Column <= ultimaDerecha And ActiveCell.row > 4 Then
        
        If ActiveCell.Interior.ColorIndex = naranja Or Selection.Interior.ColorIndex = naranja Then
            ActiveCell.Interior.Color = xlNone
            Selection.Interior.ColorIndex = xlNone
        Else
            ActiveCell.Interior.ColorIndex = naranja
            Selection.Interior.ColorIndex = naranja
        End If
    End If
    Worksheets("LISTADO").Protect pass
    Call SiAlertas
End Sub

' Inserta una persona en la fila Nº 5.
Sub InsPersona()
    Dim criterio As String
    
    Call NoAlertas
    
    Worksheets("LISTADO").Unprotect pass
    Rows(5).Insert Shift:=xlShiftDown, CopyOrigin:=xlFormatFromRightOrBelow
    criterio = ",=""" & Worksheets("VARIANTES").Range("C3").Value & """"
    Debug.Print criterio
    
    '--- NUEVAS LÍNEAS ---
    ' Restablece solo el color de fondo de la fila a "Sin relleno".
    Rows(5).Interior.ColorIndex = xlNone
    ' Restablece solo el color de la letra a "Automático" (normalmente negro).
    Rows(5).Font.ColorIndex = xlAutomatic
    
    Cells(5, 1).Activate
    Call ultima
    
    ' Celda con los inactivos
    Cells(ultimaConDatos + 2, 2).Formula = "=COUNTIF(" & "A5:A" & ultimaConDatos & ",""=Activo"")"
    
    ' Celda con los inactivos
    Cells(ultimaConDatos + 3, 2).Formula = "=COUNTIF(" & "A5:A" & ultimaConDatos & ",""=Inactivo"")"
    
    ' Celda con los totales
    Cells(ultimaConDatos + 4, 2).Formula = "=COUNTA(" & "B5:B" & ultimaConDatos & ")"
    
    
    
    Call ultima
    
    Worksheets("LISTADO").Protect pass
    
    Call SiAlertas
    ThisWorkbook.Save
End Sub

' Borra una fila, siempre la Nº 5. Deja una solita, nada más.
Sub BorrarPersona()
    Worksheets("LISTADO").Unprotect pass
    
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
    Worksheets("LISTADO").Protect pass
    Call SiAlertas
    ThisWorkbook.Save
End Sub

' Inserta un producto: son 3 columnas, con su formato, fórmula y restricciones.
Sub InsProducto()
    Worksheets("LISTADO").Unprotect pass
    Call NoAlertas
    Call ultima
    Columns("E:G").Select
    Selection.Insert CopyOrigin:=xlFormatFromRightOrBelow
    Range(Cells(2, 8), Cells(ultimaConDatos, 10)).Select
    Selection.Copy
    Range("E2").Activate
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Columns("H:J").Select
    Columns("H:J").Copy
    Columns("E:G").Select
    Selection.PasteSpecial Paste:=xlPasteFormats
        
     '--- NUEVAS COLUMNAS ---
    ' Restablece solo el color de fondo de la fila a "Sin relleno".
    Selection.Interior.ColorIndex = xlNone
    
    ' Restablece solo el color de la letra a "Automático" (normalmente negro).
    Selection.Font.ColorIndex = xlAutomatic
    
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
    Worksheets("LISTADO").Unprotect pass
    Call ultima
    
    ' Advirtiendo sobre el borrado. Se sale si se cancela
    Call Advertencia
    Call NoAlertas
    If Respuesta <> vbYes Then Exit Sub
    If ultimaDerecha < 12 Then
        MsgBox "No se puede eliminar el último producto."
        Range("E2").Select
        ActiveWorkbook.Save
        Exit Sub
    End If
    Columns("E:G").Select
    Selection.Delete
    Range("E2").Select
    
    Call ultima
    
    Worksheets("LISTADO").Protect pass
    
    Call SiAlertas
    
    ThisWorkbook.Save
End Sub

' Función auxiliar para limpiar nombres de hojas (privada a este módulo)
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
    ' Podrías añadir más reemplazos si es necesario
    If Len(sOut) > 31 Then sOut = Left(sOut, 31) ' Asegurar longitud máxima
    CleanSheetName = sOut
End Function


'==================================================================================================
' MACRO PRINCIPAL: Faltantes
'==================================================================================================
Sub Faltantes()
    Dim ws As Worksheet
    Dim shFaltantes As Worksheet
    Dim shPlanilla As Worksheet
    Dim productSheet As Worksheet
    Dim productCode As String
    Dim j_col As Long ' Índice de columna para productos
    Dim lastProductColPlanilla As Long
    Dim i_sheetCounter As Long ' Para bucle de borrado de hojas
    Dim lastDataRowOnProductSheet As Integer
    Dim nextPasteRow As Integer
    ordenTalles = Array("3XS", "XXS", "XS", "S", "M", "L", "XL", "XXL", "3XL", "4XL", "5XL", "6XL")
    ultimaResumen = 0
    Dim preservedSheetNames As Object
    Set preservedSheetNames = CreateObject("Scripting.Dictionary")

    ' Configuración inicial y desprotección
    Call NoAlertas ' Tu función (DisplayAlerts = False, Calculation = xlCalculationManual)
    Application.ScreenUpdating = False ' Manejar explícitamente para evitar parpadeo
    
    On Error GoTo Faltantes_ErrorHandler ' Establecer manejador de errores para esta subrutina

    Call desproteger ' Tu función para desproteger todo

    ' --- INICIO DE LA NUEVA LÓGICA ---

    ' 1. Eliminar hojas excepto las preservadas
    preservedSheetNames(UCase("LISTADO")) = True
    preservedSheetNames(UCase("VARIANTES")) = True
    preservedSheetNames(UCase("CODIGOS")) = True


    For Each ws In ThisWorkbook.Worksheets
        If ws.Index > 3 Then
            ws.Delete
        End If
    Next ws
    Set ws = Nothing ' Limpiar referencia

    ' 2. Crear hoja "FALTANTES"
    Set shFaltantes = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    On Error Resume Next ' Por si el nombre "FALTANTES" ya existe de alguna manera (improbable aquí)
    shFaltantes.Name = "FALTANTES"
    If Err.Number <> 0 Then
        MsgBox "Error crítico: No se pudo nombrar la hoja como 'FALTANTES'." & vbNewLine & _
               "Puede que ya exista una hoja protegida o con un nombre problemático." & vbNewLine & _
               "Error: " & Err.Description, vbCritical, "Error al Crear Hoja FALTANTES"
        Err.Clear
        GoTo Faltantes_Cleanup ' Salir si no se puede crear FALTANTES
    End If
    On Error GoTo Faltantes_ErrorHandler ' Restaurar manejador de errores principal de esta Sub

    ' 3. Identificar códigos de producto desde "PLANILLA" y crear hojas de producto
    On Error Resume Next ' Intentar obtener la hoja "PLANILLA"
    Set shPlanilla = ThisWorkbook.Worksheets("LISTADO")
    On Error GoTo Faltantes_ErrorHandler ' Restaurar manejador de errores principal

    If shPlanilla Is Nothing Then
        MsgBox "La hoja 'LISTADO' no fue encontrada. No se pueden crear hojas de producto.", vbExclamation
        ' El resto de la lógica de esta sección se omitirá si shPlanilla no existe.
        ' El MsgBox al final indicará que la lógica de llenado está pendiente.
    Else
        ' Determinar la última columna con códigos de producto en la Fila 2 de "PLANILLA"
        If IsEmpty(shPlanilla.Cells(2, shPlanilla.Columns.Count).Value) And shPlanilla.Cells(2, 1).Value = "" Then
             ' Si la última celda de la fila 2 está vacía Y la primera también, la fila podría estar vacía
            If shPlanilla.Cells(2, shPlanilla.Columns.Count).End(xlToLeft).Column = 1 And IsEmpty(shPlanilla.Cells(2, 1).Value) Then
                lastProductColPlanilla = 0 ' Indica que la fila está esencialmente vacía o solo datos en col A
            Else
                lastProductColPlanilla = shPlanilla.Cells(2, shPlanilla.Columns.Count).End(xlToLeft).Column
            End If
        Else
            lastProductColPlanilla = shPlanilla.Cells(2, shPlanilla.Columns.Count).End(xlToLeft).Column
        End If

        If lastProductColPlanilla < 5 Then
            Debug.Print "No se encontraron códigos de producto en la Fila 2 de 'PLANILLA' a partir de la Columna E."
        Else
            For j_col = 5 To lastProductColPlanilla Step 3 ' Desde Col E, cada 3 columnas
                productCode = Trim(CStr(shPlanilla.Cells(2, j_col).Value))

                If Not productCode Then
                    Dim tempSheetNameAttempt As String
                    tempSheetNameAttempt = productCode
                    
                    Set productSheet = Nothing ' Reiniciar para cada intento
                    On Error Resume Next     ' Activar manejo de errores para la creación y nombrado de hoja

                    Set productSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Sheets.Count))
                    productSheet.Name = tempSheetNameAttempt
                    
                    If Err.Number <> 0 Then ' Si falló el primer intento de nombrar
                        Err.Clear
                        Dim altSheetName As String
                        altSheetName = "Prod_" & CleanSheetName(productCode) ' Usar CleanSheetName y prefijo
                        If Len(altSheetName) > 31 Then altSheetName = Left(altSheetName, 31)
                        
                        ' Si productSheet no se pudo crear en el Add, este Name fallará.
                        ' Asumimos que Add funcionó pero el Name inicial falló.
                        If Not productSheet Is Nothing Then
                           productSheet.Name = altSheetName
                        Else ' Esto sería muy raro si el Add falló silenciosamente
                           Debug.Print "Error al AÑADIR hoja para producto: " & productCode & ". Se omitirá."
                           GoTo SiguienteCodigoProducto ' Saltar a la siguiente iteración de j_col
                        End If

                        If Err.Number <> 0 Then ' Si aún falla con nombre alternativo
                            Err.Clear
                            If Not productSheet Is Nothing Then
                                productSheet.Name = "HojaAuto" & Format(Now, "HHmmss") & "_" & j_col
                            End If
                        End If
                    End If
                    On Error GoTo Faltantes_ErrorHandler ' Restaurar manejador principal

                    If Not productSheet Is Nothing Then
                        
                        ' CREANDO LOS TITULARES DE CADA HOJA
                        productSheet.Cells(1, 1).Value = productSheet.Name & ": "
                        Range(productSheet.Cells(2, 1), productSheet.Cells(2, 4)).Merge
                        With productSheet
                            .Cells(2, 1).HorizontalAlignment = xlRight
                            .Cells(2, 1).Value = "Código: "
                            .Cells(1, 2).HorizontalAlignment = xlRight
                            .Cells(2, 5).Value = productSheet.Name
                            .Cells(1, 2).HorizontalAlignment = xlLeft
                            .Cells(3, 1).Value = Trim(CStr(shPlanilla.Cells(3, j_col).Value))
                            .Cells(4, 1).Value = "TALLE"
                            .Cells(4, 2).Value = "COLOR"
                            .Cells(4, 3).Value = "TOTAL"
                            .Cells(4, 4).Value = "SEPARADOS"
                            .Cells(4, 5).Value = "FALTANTES"
                        End With
                        Range(productSheet.Cells(3, 1), productSheet.Cells(3, 5)).Merge

                        ' CREANDO LOS HIPERVINCULOS DE LAS HOJAS HACIA EL RESPECTIVO
                        productSheet.Cells(2, 5).Select
                        With Selection.Hyperlinks
                            .Add Anchor:=Selection, Address:="", SubAddress:="LISTADO!" & Chr(64 + j_col) & 2
                        End With
                        
                        With productSheet
                            .Range("E2").Font.Size = 24
                            .Range("E2").Font.Bold = True
                        End With
                        
                        
                        ' Calculando los totales y faltantes
                        Call CalcFal(productSheet.Name, j_col)
                        
                        

                        
                    Else
                        Debug.Print "No se pudo crear la hoja para el código de producto: '" & productCode & "'"
                    End If
                End If
SiguienteCodigoProducto:
            Next j_col
        End If ' End If lastProductColPlanilla < 5
    End If ' End If shPlanilla Is Nothing
    Set shPlanilla = Nothing ' Limpiar referencia
    Set productSheet = Nothing
    
    ' --- FIN DE LA NUEVA LÓGICA ---

    If ultimaConDatos > 0 Then
        Dim wsActivar As Worksheet
        Set wsActivar = Nothing ' Asegurar que está limpio
        On Error Resume Next
        Set wsActivar = ActiveSheet ' La hoja activa puede ser la última creada.
        On Error GoTo Faltantes_ErrorHandler

        If Not wsActivar Is Nothing Then
            On Error Resume Next ' Error local para la activación
            wsActivar.Cells(ultimaConDatos, 1).Activate
            If Err.Number <> 0 Then
                Err.Clear
            End If
            On Error GoTo Faltantes_ErrorHandler
        End If
        Set wsActivar = Nothing
    End If
    
    ' ...COLOCANDO TITULARES Y ENLACES....
    ' MOSTRANDO EL RESUMEN =======================
    With Worksheets("FALTANTES")
       .Activate
       .Range("A:E").Columns.AutoFit
    End With
   

      
Faltantes_Cleanup:
    Call proteger ' Tu función
    
    ' En lugar de llamar a SiAlertas, restauramos explícitamente lo que NoAlertas cambió
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True ' Asegurar que la actualización de pantalla se restaura

    ' Limpieza final de objetos
    Set preservedSheetNames = Nothing
    Set shFaltantes = Nothing
    ' Las variables globales como ultimaConDatos no se "Set Nothing"
    Exit Sub

Faltantes_ErrorHandler:
    MsgBox "Error en Sub Faltantes: " & Err.Number & " - " & Err.Description, vbCritical, "Error en Macro Faltantes"
    Resume Faltantes_Cleanup

   
End Sub

'========================================================================
' FUNCIÓN AUXILIAR PARA OBTENER EL ÍNDICE DE ORDENACIÓN DEL TALLE
'========================================================================
Private Function GetTalleSortIndex(ByVal Talle As String) As Integer
    ' Esta función devuelve la posición de un talle en el array global 'ordenTalles'.
    ' Los talles no encontrados se colocan al final.
    Dim i As Integer
    
    ' El array 'ordenTalles' debe estar disponible en el módulo
    On Error Resume Next ' Si 'ordenTalles' no está inicializado, evitará un error
    For i = LBound(ordenTalles) To UBound(ordenTalles)
        If UCase(ordenTalles(i)) = UCase(Talle) Then
            GetTalleSortIndex = i
            Exit Function
        End If
    Next i
    On Error GoTo 0
    
    ' Si el talle no se encuentra en el array, devolver un número alto para que vaya al final
    GetTalleSortIndex = 9999
End Function

' ======= CUENTA Y CALCULA LOS FALTANTES (VERSIÓN CON ORDENACIÓN AVANZADA) =======
Sub CalcFal(Producto As String, Col_Talle As Long)
    ' Declaración de variables
    Dim dict As Object
    Dim wsListado As Worksheet
    Dim wsProducto As Worksheet
    Dim i As Long, j As Long, k As Long ' Contadores para bucles
    Dim ultimaFilaListado As Long
    
    'Dim Talle As Variant
    'Dim Color As String
    Dim Cantidad As Long
    Dim key As Variant
    
    Dim outputRow As Long
    Dim dataArray As Variant
    
    ' --- Variables para la ordenación ---
    Dim unsortedData() As Variant ' Array para volcar el diccionario
    Dim tempRow As Variant      ' Array para el intercambio en la ordenación (bubble sort)
    Dim talle1_index As Integer, talle2_index As Integer
    Dim talle1_val As String, talle2_val As String
    Dim color1_val As String, color2_val As String
    Dim mustSwap As Boolean
    
    ' Crear el objeto Dictionary
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Establecer referencias a las hojas
    Set wsListado = ThisWorkbook.Worksheets("LISTADO")
    Set wsProducto = ThisWorkbook.Worksheets(Producto)
    
    ' Obtener la última fila con datos
    ultimaFilaListado = wsListado.Cells(wsListado.Rows.Count, 1).End(xlUp).row
    
    ' Inicializando la primera celda donde se copiará el resumen en FALTANTES

    '================================================================
    ' PASO 1: AGREGAR DATOS DESDE LA HOJA "LISTADO" (Sin cambios)
    '================================================================
    For i = 5 To ultimaFilaListado - 1
        Talle = wsListado.Cells(i, Col_Talle).Value
        Color = wsListado.Cells(i, Col_Talle).Offset(0, 1).Value
        
        'If IsNumeric(wsListado.Cells(i, Col_Talle).Offset(0, 2).value) And Not IsEmpty(Talle) And Talle <> "" Then
            Cantidad = CLng(wsListado.Cells(i, Col_Talle).Offset(0, 2).Value)
            
            If Cantidad > 0 And Worksheets("LISTADO").Cells(i, 1).Value = Worksheets("VARIANTES").Range("C2").Value Then
                key = Trim(CStr(Talle)) & "|" & Trim(CStr(Color))
                
                If Not dict.exists(key) Then
                    dict.Add key, Array(0, 0)
                End If
                
                dataArray = dict(key)
                dataArray(0) = dataArray(0) + Cantidad
                
                If wsListado.Cells(i, Col_Talle).Offset(0, 2).Interior.ColorIndex = naranja And Worksheets("LISTADO").Cells(i, 1).Value = Worksheets("VARIANTES").Range("C2").Value Then
                    dataArray(1) = dataArray(1) + Cantidad
                End If
                
                dict(key) = dataArray
            
            End If
        'End If
    Next i

    If dict.Count > 0 Then
        '================================================================
        ' PASO 2: VOLCAR DATOS A UN ARRAY Y ORDENARLO (Lógica Mejorada)
        '================================================================
        ReDim unsortedData(1 To dict.Count, 1 To 5)
        i = 1
        For Each key In dict.keys
            unsortedData(i, 1) = Split(key, "|")(0) ' Talle
            unsortedData(i, 2) = Split(key, "|")(1) ' Color
            dataArray = dict(key)
            unsortedData(i, 3) = dataArray(0) ' Total
            unsortedData(i, 4) = dataArray(1) ' Separado
            unsortedData(i, 5) = unsortedData(i, 3) - unsortedData(i, 4) ' Faltante
            i = i + 1
        Next key
        
        ' Algoritmo de Ordenación Avanzado (Bubble Sort)
        For i = 1 To dict.Count - 1
            For j = i + 1 To dict.Count
                ' Obtener valores para una comparación más limpia
                talle1_index = GetTalleSortIndex(unsortedData(i, 1))
                talle2_index = GetTalleSortIndex(unsortedData(j, 1))
                talle1_val = unsortedData(i, 1)
                talle2_val = unsortedData(j, 1)
                color1_val = unsortedData(i, 2)
                color2_val = unsortedData(j, 2)
                
                mustSwap = False ' Por defecto, no intercambiar

                ' --- Inicio de la lógica de decisión ---
                If talle1_index > talle2_index Then
                    ' Criterio 1: Ordenar por el array 'ordenTalles'. El índice mayor va después.
                    mustSwap = True
                ElseIf talle1_index = talle2_index Then
                    ' Los índices son iguales. Esto significa que o son el mismo talle estándar,
                    ' o ambos son talles "no estándar" (ambos con índice 9999).
                    
                    If talle1_val <> talle2_val Then
                        ' Criterio 2: Son talles "no estándar" diferentes. Hay que ordenarlos.
                        If IsNumeric(talle1_val) And IsNumeric(talle2_val) Then
                            ' Si ambos son numéricos, comparar como números.
                            If CDbl(talle1_val) > CDbl(talle2_val) Then mustSwap = True
                        Else
                            ' Si no, comparar como texto.
                            If talle1_val > talle2_val Then mustSwap = True
                        End If
                    Else
                        ' Criterio 3: Los talles son idénticos. El desempate es el color.
                        If color1_val > color2_val Then
                            mustSwap = True
                        End If
                    End If
                End If
                ' --- Fin de la lógica de decisión ---

                If mustSwap Then
                    ' Intercambiar las filas completas si es necesario
                    ReDim tempRow(1 To 5)
                    For k = 1 To 5
                        tempRow(k) = unsortedData(i, k)
                        unsortedData(i, k) = unsortedData(j, k)
                        unsortedData(j, k) = tempRow(k)
                    Next k
                End If
            Next j
        Next i

        '================================================================
        ' PASO 3: VOLCAR LOS DATOS YA ORDENADOS A LA HOJA
        '================================================================
        ' (Esta sección es idéntica a la versión anterior)
        outputRow = 5
        For i = 1 To dict.Count
            wsProducto.Cells(outputRow, 1).Value = unsortedData(i, 1) ' Talle
            wsProducto.Cells(outputRow, 2).Value = unsortedData(i, 2) ' Color
            wsProducto.Cells(outputRow, 3).Value = unsortedData(i, 3) ' Total
            
            If unsortedData(i, 4) > 0 Then
                wsProducto.Cells(outputRow, 4).Value = unsortedData(i, 4)
            Else
                wsProducto.Cells(outputRow, 4).Value = Empty
            End If
            
            If unsortedData(i, 5) > 0 Then
                wsProducto.Cells(outputRow, 5).Value = unsortedData(i, 5)
            Else
                wsProducto.Cells(outputRow, 5).Value = Empty
            End If
            
            outputRow = outputRow + 1
        Next i
        
        '================================================================
        ' PASO 4 y 5: TOTALES Y FORMATO (Sin cambios)
        '================================================================
        With wsProducto
            .Range("A3").Font.Size = 14
            .Range("A3:E4").Font.Bold = True
            .Cells(outputRow, 2).Value = "TOTALES"
            .Cells(outputRow, 2).Font.Bold = True
            
            .Cells(outputRow, 3).Formula = "=SUM(C5:C" & outputRow - 1 & ")"
            .Cells(outputRow, 4).Formula = "=SUM(D5:D" & outputRow - 1 & ")"
            .Cells(outputRow, 5).Formula = "=SUM(E5:E" & outputRow - 1 & ")"
            
            .Range(.Cells(outputRow, 3), .Cells(outputRow, 5)).Font.Bold = True
            
            .Range("A2:E" & outputRow).Borders.LineStyle = xlContinuous
        
            .Range("C1:C2").Font.Bold = True
            
            .Range("A1:E" & outputRow).Columns.AutoFit
        End With
        
        '================================================================
        ' PASO 6: COPIAR RESUMEN A LA HOJA "FALTANTES"
        '================================================================
        If ultimaResumen < 1 Then
            ultimaResumen = 1
        End If
        
        ' CREANDO LOS HIPERVINCULOS DE LAS HOJAS HACIA EL RESUMEN
        Cells(1, 1).Select
        With Selection.Hyperlinks
            .Add Anchor:=Selection, Address:="", SubAddress:="FALTANTES!E" & ultimaResumen & "", TextToDisplay:="<<-  Ir al Resumen"
        End With
        
        
        wsProducto.Range("A2:E" & outputRow & "").Select
        With Selection
            .Copy
        End With
        Sheets("FALTANTES").Activate
        Sheets("FALTANTES").Cells(ultimaResumen, 1).Activate
        ActiveSheet.Paste
        
        ' Agregar un hipervínculo a la respectiva pestaña
       
        
        ' Incrementar las filas para el pròximo pegue
        ultimaResumen = ultimaResumen + outputRow + 1
        Range("A1:E" & outputRow).Columns.AutoFit
        
    Else
        wsProducto.Cells(4, 1).Value = "No se encontraron pedidos para este producto."
    End If
    
    Call Creditos
    
    ' Liberar memoria
    Set dict = Nothing
    Set wsListado = Nothing
    Set wsProducto = Nothing
End Sub

Sub Creditos()
    ' Calcula totales, faltantes y diferencias en montos
    Dim F, C As Integer
    Dim Acumulado As Long
    Dim formulaString As String
    Dim formulaStringEntregado As String
    Dim Listado As Worksheet
    Dim CantSeparada As Variant
    Set Listado = ThisWorkbook.Worksheets("LISTADO")
    
    Call ultima
    
    For F = 5 To ultimaConDatos
        formulaString = "="
        formulaStringEntregado = "="
        
        For C = 7 To ultimaDerecha - 4 Step 3
            If Listado.Cells(F, 1).Value = Worksheets("VARIANTES").Range("C3").Value Then
                GoTo Seguir
            End If
            
            ' Fórmula en Créditos Aprobados
            formulaString = formulaString & Listado.Cells(F, C).Address & "*" & Listado.Cells(1, C - 1).Address & "+"
        
            ' Fórmula en Créditos Entregados
            Debug.Print Listado.Cells(F, C).Interior.ColorIndex
            
            If Listado.Cells(F, C).Interior.ColorIndex = naranja Then
                CantSeparada = Listado.Cells(F, C).Address
            Else
                CantSeparada = 0
            End If
            
            formulaStringEntregado = formulaStringEntregado & CantSeparada & "*" & Listado.Cells(1, C - 1).Address & "+"

Seguir:
        Next C
        
        ' Quitamos el último "+" que sobra al final de la cadena
        If Len(formulaString) > 1 Then
            formulaString = Left(formulaString, Len(formulaString) - 1)
        End If
        
        If Len(formulaStringEntregado) > 1 Then
            formulaStringEntregado = Left(formulaStringEntregado, Len(formulaStringEntregado) - 1)
        End If
        
        ' Asignamos la fórmula completa a la celda de destino una sola vez
        If Listado.Cells(F, 1).Value = Worksheets("VARIANTES").Range("C2").Value Then
            Listado.Cells(F, ultimaDerecha - 3).Formula = formulaString
            Listado.Cells(F, ultimaDerecha - 2).Formula = formulaStringEntregado
            Listado.Cells(F, ultimaDerecha - 1).Formula = "=" & Listado.Cells(F, ultimaDerecha - 3).Address & "-" & Listado.Cells(F, ultimaDerecha - 2).Address
        End If
    Next F
End Sub
