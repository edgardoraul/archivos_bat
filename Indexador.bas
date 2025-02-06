Attribute VB_Name = "Indexador"
Option Explicit
Global maximaColumna As Integer

Sub indexador()
    Dim ultima As Integer
    Dim i As Integer
    Dim palabra As String
    maximaColumna = 29
    
    ' Empieza por la última hoja, el índice.
    ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count).Activate
    Range("A1").Activate
    ultima = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Bucle que recorre la columna A desde la 3° fila, hasta la última fila que tenga datos
    
    For i = Cells(1, 1).End(xlDown).Row To ultima
        palabra = Cells(i, 2).Value
        
        Debug.Print "Fila " & i & " -> " & palabra
        
        ' Limpiamos la fila de datos para ser insertados.
        Range(Cells(i, 5), Cells(i, Columns.Count)).ClearContents
        
        ' Se llama la función para que nos arroje el enlace a la celda
        ' y el nombre de la hoja donde la encontró
        Call buscarPalabraExacta(palabra, i)
    Next i
    
    Debug.Print "Columna máxima: " & maximaColumna
    With Range(Cells(1, 1), Cells(ultima, maximaColumna))
        .Borders.LineStyle = xlContinuous
        .EntireColumn.AutoFit
    End With
End Sub

Function buscarPalabraExacta(palabra As String, fila As Integer)
    Dim sh As Worksheet
    Dim c As Range
    Dim firstAddress As String
    Dim col As Integer
    
    ' Recorre todas las hojas excepto la última
    For Each sh In ThisWorkbook.Sheets
        If sh.Index < ThisWorkbook.Sheets.Count Then
            
            With sh.UsedRange
                Set c = .Find(What:=palabra, LookIn:=xlValues, LookAt:=xlPart)
                If Not c Is Nothing Then
                    firstAddress = c.Address
                    
                    
                    Do
                        ' Validar que la celda contenga la palabra según las condiciones especificadas
                        If EsPalabraValida(c.Value, palabra) Then
                            ' Coloca los resultados en la fila correspondiente de la última hoja
                            With ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
                                
                                ' Para que reemplace el contenido
                                col = col + 1
                                If col < 5 Then col = 5
                                .Cells(fila, col).Activate
                                .Cells(fila, col).Value = sh.Name
                                .Hyperlinks.Add Anchor:=.Cells(fila, col), _
                                    Address:="", SubAddress:="'" & sh.Name & "'!" & c.Address, _
                                    TextToDisplay:=sh.Name, _
                                    ScreenTip:=sh.Name & "!" & c.Address
                                    
                                    ' Depuracion
                                    Debug.Print sh.Name & "!" & c.Address
                                
                                ' Tamaño de letra 8
                                .Cells(fila, col).Font.Size = 8
                                Debug.Print "Encontrado en: " & sh.Name & " -> " & c.Address
                                
                            End With
                            
                            ' Acumulando un valor extra de columna
                            If col > maximaColumna Then
                                maximaColumna = col
                                Debug.Print "Columna máxima: " & maximaColumna
                            End If
                        End If
                        Set c = .FindNext(c)
                        If c Is Nothing Then GoTo meta
                    Loop While Not c Is Nothing And c.Address <> firstAddress
                End If
            End With
        End If
meta:
    Next sh
End Function



' Función auxiliar que valida si la palabra encontrada cumple las condiciones requeridas.
Function EsPalabraValida(texto As String, palabra As String) As Boolean
    Dim pos As Integer
    Dim endPos As Integer
    Dim okPre As Boolean, okPost As Boolean
    
    ' 1. Coincidencia exacta (después de quitar espacios al inicio y fin)
    If StrComp(Trim(texto), palabra, vbTextCompare) = 0 Then
        EsPalabraValida = True
        Exit Function
    End If
    
    pos = 1
    Do
        pos = InStr(pos, texto, palabra, vbTextCompare)
        If pos = 0 Then Exit Do
        
        endPos = pos + Len(palabra) - 1
        
        ' Determinar el estado del carácter precedente:
        If pos = 1 Then
            ' Si está al comienzo, no se evalúa el carácter previo.
            okPre = True
        Else
            okPre = (Mid(texto, pos - 1, 1) = " ")
        End If
        
        ' Determinar el estado del carácter siguiente:
        If endPos = Len(texto) Then
            ' Si está al final, no se evalúa el carácter siguiente.
            okPost = True
        Else
            okPost = (Mid(texto, endPos + 1, 1) = " ")
        End If
        
        ' 2. Si la palabra está al comienzo (pos = 1), debe tener un espacio justo después.
        If pos = 1 And endPos < Len(texto) Then
            If okPost Then
                EsPalabraValida = True
                Exit Function
            End If
        End If
        
        ' 3. Si la palabra está al final (endPos = Len(texto)), debe tener un espacio justo antes.
        If endPos = Len(texto) And pos > 1 Then
            If okPre Then
                EsPalabraValida = True
                Exit Function
            End If
        End If
        
        ' 4. Si la palabra está en medio (no al inicio ni al final), debe estar rodeada por espacios.
        If pos > 1 And endPos < Len(texto) Then
            If okPre And okPost Then
                EsPalabraValida = True
                Exit Function
            End If
        End If
        
        pos = pos + 1
    Loop
    EsPalabraValida = False
End Function



Sub reemplazarCaracteres()
    Dim ws As Worksheet
    Dim c As Range
    Dim cambiosTotales As Long
    Dim cambiosHoja As Long
    Dim resultados As Object
    Set resultados = CreateObject("Scripting.Dictionary")
    
    ' Inicializar contadores
    cambiosTotales = 0
    
    ' Recorremos todas las hojas
    For Each ws In ThisWorkbook.Worksheets
        cambiosHoja = 0
        
        ' Buscar y reemplazar en texto: "=" al final sin espacios
        For Each c In ws.UsedRange
            If Not IsEmpty(c.Value) And VarType(c.Value) = vbString Then
                ' Caso: "=" al final sin espacios
                If Right(c.Value, 1) = "=" And Len(c.Value) > 1 Then
                    If Not Left(c.Value, Len(c.Value) - 1) Like "*=*" Then
                        c.Value = Left(c.Value, Len(c.Value) - 1) & " ="
                        cambiosHoja = cambiosHoja + 1
                    End If
                End If
                
                ' Caso: "=" sin espacios ni al inicio ni al final
                If InStr(1, c.Value, "=") > 1 And Not (Mid(c.Value, InStr(1, c.Value, "=") - 1, 1) = " " And Mid(c.Value, InStr(1, c.Value, "=") + 1, 1) = " ") Then
                    c.Value = Replace(c.Value, "=", " = ")
                    cambiosHoja = cambiosHoja + 1
                End If
            End If
        Next c
        
        ' Guardar resultados por hoja
        If cambiosHoja > 0 Then
            resultados.Add ws.Name, cambiosHoja
            cambiosTotales = cambiosTotales + cambiosHoja
        End If
    Next ws
    
    ' Mostrar resultados
    Debug.Print "Total de cambios realizados: " & cambiosTotales
    Dim key As Variant
    For Each key In resultados.Keys
        Debug.Print "Hoja: " & key & ", Cambios: " & resultados(key)
    Next key
End Sub


