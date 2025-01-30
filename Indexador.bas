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
                    Debug.Print c.Value
                    
                    
                    Do
                        ' Verificar que la celda contenga exactamente la palabra buscada
                        If StrComp(Trim(c.Value), palabra, vbTextCompare) = 0 Then
                            ' Coloca los resultados en la fila correspondiente de la última hoja
                            With ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
                                
                                ' Para que reemplace el contenido
                                col = col + 1
                                If col < 5 Then col = 5
                                .Cells(fila, col).Activate
                                .Cells(fila, col).Value = sh.Name
                                .Hyperlinks.Add Anchor:=.Cells(fila, col), Address:="", SubAddress:= _
                                sh.Name & "!" & c.Address, TextToDisplay:=sh.Name
                                
                                ' Tamaño de 8 a la letra
                                .Cells(fila, col).Font.Size = 8
                                Debug.Print "Encontrado en: " & sh.Name
                                
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


