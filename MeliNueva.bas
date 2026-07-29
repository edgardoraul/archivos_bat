Attribute VB_Name = "MeliNueva"
Option Explicit
Public RUTA As String
Public Const MIPC As String = "D:\Web\Listados de Ventas Online\MELY"
Public Const MIPCENRED As String = "\\EDGAR\Web\Listados de Ventas Online\MELY"
Public UltimaTotal As Long
Public IdCliente As String

Sub MeliNueva()
' GENERACION DE PLANILLAS DE MELI AÑO 2027
    Call GuardarCopiaConSecuencial
End Sub

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
        RUTA = MIPC
    ElseIf workgroupName = "RERDA" Then
        RUTA = MIPCENRED
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
    
    ' 3. Abrir la copia temporal para trabajar sobre ella.
    Set wbTemp = Workbooks.Open(tempPath)
    
    ' Ver más adelante dónde lo pongo, tendría que ir al final
    wbOriginal.Close SaveChanges:=False
    
    ' 4. Limpieza de datos
    'Call LimpiezaDatos(wbTemp)
    Call LimpiezaDatos

    ' 5. Determinar la fecha de hoy en formato AAAA-MM-DD
    todayStr = Format(Date, "yyyy-mm-dd")
    filePrefix = todayStr & ". "

    ' 6. Analizar los archivos de la carpeta RUTA para encontrar el mayor número del patrón
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

    ' 7. Generar el nuevo nombre correlativo (aumentando en +1)
    numStr = Format(maxNum + 1, "00000")
    finalPath = RUTA & "\" & filePrefix & numStr & " - MELI.xlsx"
    
    ' 8. Controlando si existe el archivo Stock.XLS para las descripciones
    tempPath = RUTA & "\..\Stock.XLS"
    If fso.FileExists(tempPath) Then
        Debug.Print "Stock.XLS OK"
    Else
        MsgBox ("El archivo Stock.XLS debe estar en la carpeta " & vbNewLine & RUTA)
        Exit Sub
    End If
        
    ' 9. Controlando si existe el archivo Equivalencia2.XLS para las variantes color/talle
    tempPath = RUTA & "\..\Equivalencia2.XLS"
    If fso.FileExists(tempPath) Then
        Debug.Print "Equivalencia2.XLS OK"
    Else
        MsgBox ("El archivo Equivalencia2.XLS debe estar en la carpeta " & vbNewLine & RUTA)
        Exit Sub
    End If
    
    ' 10. Completado de información del cliente, código, descripción, color y talle
    'Call CompletaInfo(wbTemp)
    Call CompletaInfo

    ' 11. Guardar el libro temporal en la carpeta RUTA con el nombre definitivo y cerrarlo
    wbTemp.SaveAs fileName:=finalPath, FileFormat:=51
End Sub


Sub CompletaInfo()
'Sub CompletaInfo(Planilla As Workbook)
    Dim Planilla As Workbook
    Set Planilla = ActiveWorkbook
    ' COMPLETA DATOS: descripción, color/leyenda
    ' en base a equivalencias.
    
    ' Abre un archivo de rótulos para realizar las búsquedas
    Dim Rotulos As Workbook
    Set Rotulos = AbrirArchivoRotulos()

    Dim i As Integer
    Dim rutaArchivo As String
    Dim rutaEquivalencia As String
    rutaArchivo = "'" & RUTA & "\..\[Stock.XLS]Sheet1'!"
    rutaEquivalencia = "'" & RUTA & "\..\[Equivalencia2.XLS]Sheet1'!"
    Debug.Print UltimaTotal
    
    With Planilla.Worksheets(1)
    
        For i = 2 To UltimaTotal
        
            ' Limpia de enlaces innecesarios
            .Cells(i, 1).ClearHyperlinks
        
            ' Convierte a texto el SKU
            .Cells(i, 3).Value = "'" & .Cells(i, 3).Value
            
            ' Esconde los últimos caracteres del SKU
            With .Cells(i, 3)
                .HorizontalAlignment = xlLeft
                .Font.color = vbWhite
                .Characters(Start:=1, Length:=7).Font.ColorIndex = xlAutomatic
            End With
            
            
            ' Salta fila con código vacío
            If .Cells(i, 3) = "" Then
                GoTo Siguiente
            
            ' Sólo se valida si es equivalencia
            ElseIf Len(.Cells(i, 3).Value) > 7 Then
                ' Color
                'Cells(i, 5).Formula = "=IFERROR(VLOOKUP(C" & i & ", " & rutaEquivalencia & "$A$2:$F$10000, 4, FALSE), """")"
                .Cells(i, 5).Activate
                .Cells(i, 5).Formula = "=IF(ISNA(VLOOKUP(C" & i & ", " & rutaEquivalencia & "$A$2:$F$10000, 4, FALSE)), ""Sin equivalencia"", IF(VLOOKUP(C" & i & ", " & rutaEquivalencia & "$A$2:$F$10000, 4, FALSE)="""", """", VLOOKUP(C" & i & ", " & rutaEquivalencia & "$A$2:$F$10000, 4, FALSE)))"
                .Cells(i, 5).HorizontalAlignment = xlCenter
                
                ' Leyenda
                .Cells(i, 6).Activate
                '.Cells(i, 6).Formula = "=IFERROR(VLOOKUP(C" & i & ", " & rutaEquivalencia & "$A$2:$F$10000, 5, FALSE), """")"
                .Cells(i, 6).Formula = "=IF(ISNA(VLOOKUP(C" & i & ", " & rutaEquivalencia & "$A$2:$F$10000, 5, FALSE)), ""Sin equivalencia"", IF(VLOOKUP(C" & i & ", " & rutaEquivalencia & "$A$2:$F$10000, 5, FALSE)="""", """", VLOOKUP(C" & i & ", " & rutaEquivalencia & "$A$2:$F$10000, 5, FALSE)))"
                .Cells(i, 6).HorizontalAlignment = xlCenter
                
                ' Talle
                .Cells(i, 7).Activate
                '.Cells(i, 7).Formula = "=IFERROR(VLOOKUP(C" & i & ", " & rutaEquivalencia & "$A$2:$F$10000, 6, FALSE), """")"
                .Cells(i, 7).Formula = "=IF(ISNA(VLOOKUP(C" & i & ", " & rutaEquivalencia & "$A$2:$F$10000, 6, FALSE)), ""Sin equivalencia"", IF(VLOOKUP(C" & i & ", " & rutaEquivalencia & "$A$2:$F$10000, 6, FALSE)="""", """", VLOOKUP(C" & i & ", " & rutaEquivalencia & "$A$2:$F$10000, 6, FALSE)))"
                .Cells(i, 7).HorizontalAlignment = xlCenter
            End If
            
            ' Descripción: sólo toma los 7 primeros caracteres
            .Cells(i, 4).Activate
            .Cells(i, 4).Formula = "=IFERROR(VLOOKUP(LEFT(C" & i & ", 7), " & rutaArchivo & "$A$2:$C$10000, 2, FALSE), """")"
            
            ' Busca el usuario del cliente
            Call BuscarCliente(.Cells(i, 1).Value, Rotulos)
            .Cells(i, 2).Value = IdCliente
            
Siguiente:
        Next i
    End With
    
End Sub

Public Function AbrirArchivoRotulos() As Workbook
    Dim rutaArchivo As Variant
    Dim wb As Workbook
    
    ' Abrir cuadro de diálogo para seleccionar el archivo Excel
    rutaArchivo = Application.GetOpenFilename( _
        FileFilter:="Archivos de Excel (*.xls; *.xlsx; *.xlsm), *.xls; *.xlsx; *.xlsm", _
        Title:="Selecciona el archivo de Rótulos")
        
    ' Si el usuario cancela, la función devuelve Nothing
    If rutaArchivo = False Then
        MsgBox "No se seleccionó ningún archivo.", vbExclamation, "Proceso cancelado"
        Set AbrirArchivoRotulos = Nothing
        Exit Function
    End If
    
    ' Abrir el libro y asignarlo como resultado de la función
    On Error Resume Next
    Set wb = Workbooks.Open(fileName:=rutaArchivo, ReadOnly:=True)
    On Error GoTo 0
    
    If wb Is Nothing Then
        MsgBox "No se pudo abrir el archivo seleccionado.", vbCritical, "Error"
        Set AbrirArchivoRotulos = Nothing
    Else
        ' Retorna el objeto Workbook abierto
        Set AbrirArchivoRotulos = wb
    End If
End Function

Public Function BuscarCliente(ByVal NumVenta As String, ByRef LibroRotulos As Workbook) As String
    Dim ws As Worksheet
    Dim coincidenciaEnHoja As Range
    Dim ultimaCoincidencia As Range
    Dim celdaEvaluar As Range
    Dim textoCelda As String
    Dim maxFilas As Long
    
    ' Validar entrada
    If LibroRotulos Is Nothing Or Trim(NumVenta) = "" Then
        IdCliente = ""
        BuscarCliente = ""
        Exit Function
    End If
    
    Set ultimaCoincidencia = Nothing
    
    ' 1. Buscar en TODAS las celdas de cada hoja
    For Each ws In LibroRotulos.Worksheets
        ' ws.Cells abarca la totalidad de la hoja
        Set coincidenciaEnHoja = ws.Cells.Find( _
            What:=NumVenta, _
            After:=ws.Cells(1, 1), _
            LookIn:=xlValues, _
            LookAt:=xlPart, _
            SearchDirection:=xlPrevious)
        
        If Not coincidenciaEnHoja Is Nothing Then
            Set ultimaCoincidencia = coincidenciaEnHoja
        End If
    Next ws
    
    ' 2. Bajar filas en la misma columna donde se encontró, buscando el patrón (...)
    If Not ultimaCoincidencia Is Nothing Then
        Set celdaEvaluar = ultimaCoincidencia
        maxFilas = 0
        
        Do While maxFilas < 100
            Set celdaEvaluar = celdaEvaluar.Offset(1, 0) ' Baja 1 fila en la columna de la coincidencia
            textoCelda = Trim(CStr(celdaEvaluar.Value))
            
            If Len(textoCelda) >= 2 Then
                If Left(textoCelda, 1) = "(" And Right(textoCelda, 1) = ")" Then
                    ' Purgar paréntesis
                    IdCliente = Mid(textoCelda, 2, Len(textoCelda) - 2)
                    BuscarCliente = IdCliente
                    Exit Function
                End If
            End If
            
            maxFilas = maxFilas + 1
        Loop
        
        IdCliente = "No encontrado ===="
    Else
        IdCliente = "Nada ===="
    End If
    
    BuscarCliente = IdCliente
    
    ' Cerrar el archivo
    
End Function

Sub LimpiezaDatos()
    ' LIMPIA LOS DATOS, ELIMINA COLUMNAS Y TITULA CORRECTAMENTE
    
    Dim Libro As Workbook
    Set Libro = ActiveWorkbook
    With Libro.Worksheets(1)
        ' Limpiando de formatos y dejando una base
        .Cells.ClearFormats
        .Cells.Font.Name = "Consolas"
        .Cells.Font.Size = 14
        .Name = "Ventas"
        
        ' Elimina filas innecesarias
        .Range("A1:A5").EntireRow.Delete
        
        ' Elimina Columnas innecesarias
        .Range("AW1:BN1").EntireColumn.Delete
        .Range("AF1:AU1").EntireColumn.Delete
        .Range("X1:AE1").EntireColumn.Delete
        .Range("H1:V1").EntireColumn.Delete
        .Range("B1:F1").EntireColumn.Delete
        
        ' Agrega columnas necesarias
        .Range("B:G").Insert Shift:=xlToRight
        
        ' Títulos de Columnas necesarias
        .Range("A1").Value = "Nº Venta"
        .Range("B1").Value = "Cliente"
        .Range("D1").Value = "Descripción"
        .Range("E1").Value = "Color"
        .Range("F1").Value = "Leyenda Color"
        .Range("G1").Value = "Talle"
        .Range("H1").Value = "Cant."
        
        ' Mueve los sku
        .Columns("I:I").Cut Destination:=Columns("C:C")
        .Range("I1:I1").EntireColumn.Delete
        .Range("C1").Value = "Código"
        
        ' Agrega una columna para detalles
        .Columns("I:I").Insert Shift:=xlToLeft
        .Range("I1").Value = "Detalle"
        .Range("J1").Value = "Nº Seguimiento"
        
        ' Limpia los sku de sufijos -CL y -PR
        .Cells.Replace What:="-CL", Replacement:="", LookAt:=xlPart, _
        searchorder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
        .Cells.Replace What:="-PR", Replacement:="", LookAt:=xlPart, _
        searchorder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
        ' Crea el valor de la última fila de la planilla total en su conjunto
        UltimaTotal = .Cells(Rows.Count, 1).End(xlUp).Row
        Debug.Print UltimaTotal
    End With
End Sub

