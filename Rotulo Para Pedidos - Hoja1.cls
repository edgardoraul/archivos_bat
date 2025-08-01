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
    '--- Declaraci�n de Variables ---
    Dim CeldaMonitoreada As Range
    Dim RangoBusqueda As Range
    Dim ColBusqueda As Long
    Dim ValorBusqueda As String

    '--- Configuraci�n ---
    'Celda que, al cambiar, dispara la macro.
    Set CeldaMonitoreada = Me.Range("A5")
    
    'Rango donde se buscar� el valor (en la hoja "ARTICULOS").
    'Es mejor definirlo como objeto Range para m�s flexibilidad.
    Set RangoBusqueda = ThisWorkbook.Worksheets("ARTICULOS").Range("A2:C10000")
    
    'Columna del resultado dentro de RangoBusqueda.
    ColBusqueda = 2

    '--- L�gica Principal ---
    'Verifica si la celda modificada (Target) es la que nos interesa.
    If Not Intersect(Target, CeldaMonitoreada) Is Nothing Then
    
        'Guardamos el valor a buscar para usarlo en la f�rmula.
        ValorBusqueda = CeldaMonitoreada.Value
        
        'Desactivamos los eventos para evitar que la macro se llame a s� misma.
        Call Desalertar
        
        'Manejo de errores si la celda est� vac�a.
        If Trim(ValorBusqueda) = "" Then
            Me.Range("B5").ClearContents 'Limpia la celda de resultado si A6 est� vac�a.
        Else
            'Insertamos la f�rmula en B6.
            'Usamos IFERROR para mostrar "" si no se encuentra el valor, en vez de #N/A.
            'VBA siempre usa los nombres de funci�n en ingl�s (VLOOKUP, IFERROR, FALSE).
            Me.Range("B5").Formula = "=IFERROR(VLOOKUP(" & CeldaMonitoreada.Address & ", " & _
                                     "ARTICULOS!" & RangoBusqueda.Address & ", " & _
                                     ColBusqueda & ", FALSE), """")"
        End If
        
        'Volvemos a activar los eventos. �Muy importante!
        Call Alertar
        
    End If
    
End Sub

