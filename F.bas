Attribute VB_Name = "F"
Option Explicit

Sub ProcesarFaltantesDeProductos()

    Dim shListado As Worksheet
    Dim shFaltantes As Worksheet
    Dim shProducto As Worksheet
    
    Dim nombreHojaListado As String
    Dim nombreHojaFaltantes As String
    Dim nombreProducto As String
    Dim nombreEncabezadoColumna As String
    
    Dim ultimaFilaListado As Long
    Dim ultimaColumnaListado As Long
    Dim i As Long 'Para filas
    Dim j As Long 'Para columnas (productos)
    
    Dim filaDestinoProducto As Long
    Dim filaInicioDatosProducto As Long
    Dim filaDestinoFaltantes As Long
    
    Dim dictTalles As Object
    Dim talleKey As Variant
    Dim valorCelda As String
    
    Dim totalSeparadosProducto As Long
    Dim totalFaltantesProducto As Long
    
    Dim ws As Worksheet

    Dim faltantesConsolidado As Object
    Dim sortedProductNames As Object
    Dim productSummary As Object
    Dim pKey As Variant
    Dim tKeySorted As Variant

    ' --- CONFIGURACIÓN INICIAL ---
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False

    ' --- 1. IDENTIFICAR Y VALIDAR HOJA "LISTADO" ---
    nombreHojaListado = "LISTADO"
    On Error Resume Next
    Set shListado = ThisWorkbook.Sheets(nombreHojaListado)
    On Error GoTo ErrorHandler
    
    If shListado Is Nothing Then
        MsgBox "La hoja requerida '" & nombreHojaListado & "' no fue encontrada. El proceso no puede continuar.", vbCritical
        GoTo CleanupAndExit
    End If

    ' --- 2. ELIMINAR TODAS LAS OTRAS HOJAS EXCEPTO "LISTADO" ---
    For Each ws In ThisWorkbook.Worksheets
        If UCase(ws.Name) <> UCase(nombreHojaListado) Then
            ws.Delete
        End If
    Next ws

    ' --- 3. CREAR PESTAÑA "FALTANTES" ---
    nombreHojaFaltantes = "FALTANTES"
    Set shFaltantes = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    shFaltantes.Name = nombreHojaFaltantes
    filaDestinoFaltantes = 1
    
    Set faltantesConsolidado = CreateObject("Scripting.Dictionary")
    faltantesConsolidado.CompareMode = vbTextCompare

    ' --- DETERMINAR RANGO DE DATOS EN HOJA "LISTADO" ---
    If shListado.Cells(1, 1).Value = "" Then
        MsgBox "La hoja '" & nombreHojaListado & "' parece estar vacía o no tiene el formato esperado en A1.", vbExclamation
    End If
    ultimaFilaListado = shListado.Cells(shListado.Rows.Count, 1).End(xlUp).Row
    ultimaColumnaListado = shListado.Cells(1, shListado.Columns.Count).End(xlToLeft).Column

    ' --- ITERAR POR CADA PRODUCTO EN "LISTADO" ---
    For j = 3 To ultimaColumnaListado
        nombreEncabezadoColumna = Trim(CStr(shListado.Cells(1, j).Value))

        If LCase(nombreEncabezadoColumna) = "observaciones" Or LCase(nombreEncabezadoColumna) = "entrega" Then
            GoTo SiguienteProducto
        End If

        nombreProducto = nombreEncabezadoColumna
        If nombreProducto = "" Then GoTo SiguienteProducto
        
        Set dictTalles = CreateObject("Scripting.Dictionary")
        dictTalles.CompareMode = vbTextCompare
        
        totalSeparadosProducto = 0
        totalFaltantesProducto = 0

        Set shProducto = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        
        Dim tempProdSheetName As String
        tempProdSheetName = Left(nombreProducto, 31)

        On Error Resume Next
        shProducto.Name = tempProdSheetName
        If Err.Number <> 0 Then
            Err.Clear
            tempProdSheetName = Replace(tempProdSheetName, "/", "_")
            tempProdSheetName = Replace(tempProdSheetName, "\", "_")
            tempProdSheetName = Replace(tempProdSheetName, "*", "_")
            tempProdSheetName = Replace(tempProdSheetName, "?", "_")
            tempProdSheetName = Replace(tempProdSheetName, "[", "_")
            tempProdSheetName = Replace(tempProdSheetName, "]", "_")
            tempProdSheetName = Replace(tempProdSheetName, ":", "_")
            
            If Len(tempProdSheetName) > 25 Then
                tempProdSheetName = Left(tempProdSheetName, 25)
            End If
            shProducto.Name = "Prod_" & j & "_" & tempProdSheetName
            If Len(shProducto.Name) > 31 Then shProducto.Name = Left(shProducto.Name, 31)
            
            If Err.Number <> 0 Then
                Err.Clear
                shProducto.Name = "ProductoAuto" & Format(Now, "HHmmss") & j
            End If
        End If
        On Error GoTo ErrorHandler
        
        With shProducto
            .Cells(1, 1).Value = "Producto: " & nombreProducto
            .Cells(1, 1).Font.Bold = True
            .Cells(1, 1).Font.Size = 14
            .Hyperlinks.Add Anchor:=.Cells(2, 1), Address:="", SubAddress:="'" & nombreHojaFaltantes & "'!A1", TextToDisplay:="<< Volver a " & nombreHojaFaltantes
            
            .Cells(4, 1).Value = "Talle"
            .Cells(4, 2).Value = "Separados"
            .Cells(4, 3).Value = "Faltantes (Sin Talle Espec.)"
            .Range("A4:C4").Font.Bold = True
            filaInicioDatosProducto = 5
            filaDestinoProducto = filaInicioDatosProducto
        End With
        
        For i = 2 To ultimaFilaListado
            valorCelda = Trim(CStr(shListado.Cells(i, j).Value))
            
            If UCase(valorCelda) = "ENT" Then
            ElseIf valorCelda = "" Then
                If Not dictTalles.Exists("ZZ. Sin Talle Especificado") Then
                    dictTalles("ZZ. Sin Talle Especificado") = Array(0, 0)
                End If
                Dim countsFaltante As Variant
                countsFaltante = dictTalles("ZZ. Sin Talle Especificado")
                countsFaltante(1) = countsFaltante(1) + 1
                dictTalles("ZZ. Sin Talle Especificado") = countsFaltante
            Else
                If Not dictTalles.Exists(valorCelda) Then
                    dictTalles(valorCelda) = Array(0, 0)
                End If
                Dim countsSeparado As Variant
                countsSeparado = dictTalles(valorCelda)
                countsSeparado(0) = countsSeparado(0) + 1
                dictTalles(valorCelda) = countsSeparado
            End If
        Next i
        
        If dictTalles.Count > 0 Then
            For Each talleKey In dictTalles.Keys
                shProducto.Cells(filaDestinoProducto, 1).Value = talleKey
                ' --- MODIFICACIÓN: Dejar celda vacía si el conteo es 0 ---
                shProducto.Cells(filaDestinoProducto, 2).Value = IIf(CLng(dictTalles(talleKey)(0)) = 0, "", dictTalles(talleKey)(0))
                shProducto.Cells(filaDestinoProducto, 3).Value = IIf(CLng(dictTalles(talleKey)(1)) = 0, "", dictTalles(talleKey)(1))
                ' --- FIN MODIFICACIÓN ---
                
                totalSeparadosProducto = totalSeparadosProducto + CLng(dictTalles(talleKey)(0))
                totalFaltantesProducto = totalFaltantesProducto + CLng(dictTalles(talleKey)(1))
                
                filaDestinoProducto = filaDestinoProducto + 1
            Next talleKey
        
            Dim ultimaDataRowTalles As Long
            ultimaDataRowTalles = filaDestinoProducto - 1
            
            If ultimaDataRowTalles >= filaInicioDatosProducto Then
                shProducto.Sort.SortFields.Clear
                shProducto.Range("A" & filaInicioDatosProducto & ":C" & ultimaDataRowTalles).Sort _
                    Key1:=shProducto.Range("A" & filaInicioDatosProducto), Order1:=xlAscending, Header:=xlNo
            End If

            shProducto.Cells(filaDestinoProducto, 1).Value = "TOTAL " & nombreProducto
            shProducto.Cells(filaDestinoProducto, 1).Font.Bold = True
            
            If ultimaDataRowTalles >= filaInicioDatosProducto Then
                shProducto.Cells(filaDestinoProducto, 2).Formula = "=SUM(B" & filaInicioDatosProducto & ":B" & ultimaDataRowTalles & ")"
                shProducto.Cells(filaDestinoProducto, 3).Formula = "=SUM(C" & filaInicioDatosProducto & ":C" & ultimaDataRowTalles & ")"
            Else
                shProducto.Cells(filaDestinoProducto, 2).Value = totalSeparadosProducto ' Suma fórmulas darán 0, esto también
                shProducto.Cells(filaDestinoProducto, 3).Value = totalFaltantesProducto
            End If
            shProducto.Cells(filaDestinoProducto, 2).Font.Bold = True
            shProducto.Cells(filaDestinoProducto, 3).Font.Bold = True
        End If
        shProducto.Columns("A:C").AutoFit
        
        Set productSummary = CreateObject("Scripting.Dictionary")
        Set productSummary("tallesData") = dictTalles
        productSummary("actualSheetName") = shProducto.Name
        productSummary("totalSeparados") = totalSeparadosProducto
        productSummary("totalFaltantes") = totalFaltantesProducto
        
        If Not faltantesConsolidado.Exists(nombreProducto) Then
            faltantesConsolidado.Add nombreProducto, productSummary
        Else
            faltantesConsolidado(nombreProducto & "_dup" & j) = productSummary
        End If

SiguienteProducto:
        Set dictTalles = Nothing
        Set productSummary = Nothing
        Set shProducto = Nothing
    Next j
    
    Set sortedProductNames = CreateObject("System.Collections.ArrayList")
    For Each pKey In faltantesConsolidado.Keys
        sortedProductNames.Add pKey
    Next pKey
    sortedProductNames.Sort

    For Each pKey In sortedProductNames
        Dim currentNombreProductoSorted As String
        currentNombreProductoSorted = CStr(pKey)
        
        Dim currentProductData As Object
        Set currentProductData = faltantesConsolidado(currentNombreProductoSorted)
        
        Dim currentTallesData As Object
        Set currentTallesData = currentProductData("tallesData")
        
        Dim currentActualSheetName As String
        currentActualSheetName = currentProductData("actualSheetName")

        Dim currentTotalSeparados As Long
        currentTotalSeparados = currentProductData("totalSeparados")

        Dim currentTotalFaltantes As Long
        currentTotalFaltantes = currentProductData("totalFaltantes")

        With shFaltantes
            .Cells(filaDestinoFaltantes, 1).Value = "Producto: " & currentNombreProductoSorted
            .Cells(filaDestinoFaltantes, 1).Font.Bold = True
            .Cells(filaDestinoFaltantes, 1).Font.Size = 12
            .Hyperlinks.Add Anchor:=.Cells(filaDestinoFaltantes, 1), Address:="", SubAddress:="'" & currentActualSheetName & "'!A1", TextToDisplay:="Producto: " & currentNombreProductoSorted
            filaDestinoFaltantes = filaDestinoFaltantes + 1
            
            .Cells(filaDestinoFaltantes, 2).Value = "Talle"
            .Cells(filaDestinoFaltantes, 3).Value = "Separados"
            .Cells(filaDestinoFaltantes, 4).Value = "Faltantes (Sin Talle Espec.)"
            .Range(.Cells(filaDestinoFaltantes, 2), .Cells(filaDestinoFaltantes, 4)).Font.Bold = True
            filaDestinoFaltantes = filaDestinoFaltantes + 1
            
            If currentTallesData.Count > 0 Then
                Dim sortedTalleKeysFaltantes As Object
                Set sortedTalleKeysFaltantes = CreateObject("System.Collections.ArrayList")
                For Each tKeySorted In currentTallesData.Keys
                    sortedTalleKeysFaltantes.Add tKeySorted
                Next
                sortedTalleKeysFaltantes.Sort

                For Each tKeySorted In sortedTalleKeysFaltantes
                    If CLng(currentTallesData(tKeySorted)(0)) > 0 Or CLng(currentTallesData(tKeySorted)(1)) > 0 Then
                        .Cells(filaDestinoFaltantes, 2).Value = tKeySorted
                        ' --- MODIFICACIÓN: Dejar celda vacía si el conteo es 0 ---
                        .Cells(filaDestinoFaltantes, 3).Value = IIf(CLng(currentTallesData(tKeySorted)(0)) = 0, "", currentTallesData(tKeySorted)(0))
                        .Cells(filaDestinoFaltantes, 4).Value = IIf(CLng(currentTallesData(tKeySorted)(1)) = 0, "", currentTallesData(tKeySorted)(1))
                        ' --- FIN MODIFICACIÓN ---
                        filaDestinoFaltantes = filaDestinoFaltantes + 1
                    End If
                Next tKeySorted
                Set sortedTalleKeysFaltantes = Nothing
                
                .Cells(filaDestinoFaltantes, 2).Value = "TOTAL " & currentNombreProductoSorted
                .Cells(filaDestinoFaltantes, 2).Font.Bold = True
                .Cells(filaDestinoFaltantes, 3).Value = currentTotalSeparados ' Los totales generales sí muestran 0 si es el caso
                .Cells(filaDestinoFaltantes, 3).Font.Bold = True
                .Cells(filaDestinoFaltantes, 4).Value = currentTotalFaltantes ' Los totales generales sí muestran 0 si es el caso
                .Cells(filaDestinoFaltantes, 4).Font.Bold = True
            Else
                .Cells(filaDestinoFaltantes, 2).Value = "(Sin movimientos para este producto)"
                .Cells(filaDestinoFaltantes, 2).Font.Italic = True
            End If
            filaDestinoFaltantes = filaDestinoFaltantes + 2
        End With
    Next pKey
    
    If shFaltantes.Cells(1, 1).Value <> "" Then
      shFaltantes.Columns("A:D").AutoFit

      With shFaltantes.PageSetup
          .LeftHeader = "&P"
          .CenterHeader = ""
          .RightHeader = "&A"
          .LeftFooter = ""
          .CenterFooter = "&F"
          .RightFooter = ""
      End With

      shFaltantes.Activate
      shFaltantes.Cells(1, 1).Select
    ElseIf Not shListado Is Nothing Then
        shListado.Activate
    End If

CleanupAndExit:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    If Err.Number <> 0 And Err.Source <> "" Then
        MsgBox "Ocurrió un error: " & Err.Description, vbCritical
    ElseIf Err.Number = 0 Then
        'MsgBox "Proceso de faltantes completado.", vbInformation
    End If
    If Err.Number <> 0 Then Err.Clear
    
    Set shListado = Nothing
    Set shFaltantes = Nothing
    Set shProducto = Nothing
    Set dictTalles = Nothing
    Set ws = Nothing
    Set faltantesConsolidado = Nothing
    Set sortedProductNames = Nothing
    Set productSummary = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Error en tiempo de ejecución " & Err.Number & ": " & Err.Description, vbCritical
    Resume CleanupAndExit

End Sub

