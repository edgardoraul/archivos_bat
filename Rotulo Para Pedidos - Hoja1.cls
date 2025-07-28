VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    '--- Declaración de Variables ---
    Dim CeldaMonitoreada As Range
    Dim RangoBusqueda As Range
    Dim ColBusqueda As Long
    Dim ValorBusqueda As String

    '--- Configuración ---
    'Celda que, al cambiar, dispara la macro.
    Set CeldaMonitoreada = Me.Range("A5")
    
    'Rango donde se buscará el valor (en la hoja "ARTICULOS").
    'Es mejor definirlo como objeto Range para más flexibilidad.
    Set RangoBusqueda = ThisWorkbook.Worksheets("ARTICULOS").Range("A2:C10000")
    
    'Columna del resultado dentro de RangoBusqueda.
    ColBusqueda = 2

    '--- Lógica Principal ---
    'Verifica si la celda modificada (Target) es la que nos interesa.
    If Not Intersect(Target, CeldaMonitoreada) Is Nothing Then
    
        'Guardamos el valor a buscar para usarlo en la fórmula.
        ValorBusqueda = CeldaMonitoreada.Value
        
        'Desactivamos los eventos para evitar que la macro se llame a sí misma.
        Call Desalertar
        
        'Manejo de errores si la celda está vacía.
        If Trim(ValorBusqueda) = "" Then
            Me.Range("B5").ClearContents 'Limpia la celda de resultado si A6 está vacía.
        Else
            'Insertamos la fórmula en B6.
            'Usamos IFERROR para mostrar "" si no se encuentra el valor, en vez de #N/A.
            'VBA siempre usa los nombres de función en inglés (VLOOKUP, IFERROR, FALSE).
            Me.Range("B5").Formula = "=IFERROR(VLOOKUP(" & CeldaMonitoreada.Address & ", " & _
                                     "ARTICULOS!" & RangoBusqueda.Address & ", " & _
                                     ColBusqueda & ", FALSE), """")"
        End If
        
        'Volvemos a activar los eventos. ¡Muy importante!
        Call Alertar
        
    End If
    
End Sub

