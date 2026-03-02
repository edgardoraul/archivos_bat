Attribute VB_Name = "AnalisisVentas"
Option Explicit
Sub ImagenesProducteca()
Dim producteca As Workbook
Dim imagenes As Workbook
Dim i As Integer
Dim e As Integer
Dim sku As String
Dim ultima_p As Long
Dim ultima_i As Long
Dim Hojilla As String
Dim c As Range
Dim total As Long

'hojilla = "Con Color"
'hojilla = "Variables"
Hojilla = "Simples"


Workbooks.Open ("D:\Web\archivos_bat\Producteca_img.xlsx")
Workbooks.Open ("D:\Web\archivos_bat\ListadoImagenesWeb.xlsm")

Set producteca = Workbooks("Producteca_img.xlsx")
Set imagenes = Workbooks("ListadoImagenesWeb.xlsm")

ultima_p = producteca.Worksheets(1).Cells(Rows.Count, 1).End(xlUp).Row
ultima_i = imagenes.Worksheets(Hojilla).Cells(Rows.Count, 1).End(xlUp).Row

producteca.Worksheets(1).Activate
producteca.Worksheets(1).Cells(1, 8).Activate
producteca.Worksheets(1).Cells(4, 9).Value = ultima_p - 1



For i = 2 To ultima_p
    sku = producteca.Worksheets(1).Cells(i, 3).Value
    If producteca.Worksheets(1).Cells(i, 7).Value = "Cambiado" Then
        GoTo Siguiente
    End If
    
    ' Buscar en imágenes ==========
    For e = 2 To ultima_i
        If Hojilla = "Con Color" And imagenes.Worksheets(Hojilla).Cells(e, 2).Value = sku Then
            producteca.Worksheets(1).Cells(i, 6).Value = imagenes.Worksheets(Hojilla).Cells(e, 7).Value
            producteca.Worksheets(1).Cells(i, 7).Value = "Cambiado"
            GoTo Siguiente
        ElseIf imagenes.Worksheets(Hojilla).Cells(e, 3).Value = Left(sku, 7) & "##" Then
            producteca.Worksheets(1).Cells(i, 6).Value = imagenes.Worksheets(Hojilla).Cells(e, 7).Value
            producteca.Worksheets(1).Cells(i, 7).Value = "Cambiado"
            GoTo Siguiente
        End If
        
        ' Avance
        producteca.Worksheets(1).Cells(1, 9).Value = e
        
    Next e
Siguiente:
    producteca.Worksheets(1).Cells(3, 9).Value = (i - 1) / (ultima_p - 1)
    producteca.Worksheets(1).Cells(2, 9).Value = i
Next i

End Sub
