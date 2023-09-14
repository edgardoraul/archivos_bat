Attribute VB_Name = "CTACTE"
Option Explicit
Sub Rotulador()
    Sheets("Planilla").Select

    ' Declaraci�n de Variables y su tipo de dato
    Dim apellidoNombre As String
    Dim dniCuit As String
    Dim direccion As String
    Dim codigoPostal As Variant
    Dim ciudad As String
    Dim provincia As String
    Dim telefono As String
    Dim codigoNis As String
    Dim hoja As Worksheet
    Dim fecha As String

    fecha = Format(Date, "yyyy-mm-dd")
    
     ' Controlar qui�n es el vendedor
    
    'If Range("T2").Value = "" Then
    'Cells(
    'If Cells(2, Columns.Count).End(xlToLeft).Column).Value = "" Then
    If Cells(2, Cells(1, Columns.Count).End(xlToLeft).Column).Value = "" Then
        MsgBox ("�Y qu� viajante, vendedor o sucursal sos vos?")
        Range("T2").Select
        Exit Sub
    End If
    
    ' Dando formato a la p�gina para imprimir
    With ActiveSheet.PageSetup
        .Orientation = xlLandscape
        .TopMargin = Application.CentimetersToPoints(1.9)
        .RightMargin = Application.CentimetersToPoints(0.6)
        .BottomMargin = Application.CentimetersToPoints(1.9)
        .LeftMargin = Application.CentimetersToPoints(0.6)
        .HeaderMargin = Application.CentimetersToPoints(0.8)
        .FooterMargin = Application.CentimetersToPoints(0.8)
        .CenterHorizontally = True
        .PaperSize = xlPaperA4
    End With

    ' Controla si est� parado en una celda equivocada
    If ActiveCell.Column <> Cells(1, Columns.Count).End(xlToLeft).Column - 1 Then
        MsgBox ("Debes elegir alguna compra que tenga alg�n tipo de flete.")
        Cells.Range("S2").Select
        Exit Sub
    
    ElseIf ActiveCell.Value = "" Then
        Cells(Cells(1, Columns.Count).End(xlToLeft).Column - 1).Select
        Exit Sub
    End If
    
    
    ' Variables para operar
    apellidoNombre = ActiveCell.Offset(0, -20).Value
    dniCuit = ActiveCell.Offset(0, -10).Value
    direccion = ActiveCell.Offset(0, -5).Value
    codigoPostal = ActiveCell.Offset(0, -3).Value
    ciudad = ActiveCell.Offset(0, -2).Value
    provincia = ActiveCell.Offset(0, -1).Value
    telefono = ActiveCell.Offset(0, -4).Value
    
    ' Controla si est� completo el DNI/CUIT y el CP
    If dniCuit = "" Then
        MsgBox "Te falt� completar el DNI/CUIT."
        
        ' Celda del DNI/CUIT
        ActiveCell.Offset(0, -10).Activate
        Exit Sub
    
    ElseIf apellidoNombre = "" Then
        MsgBox "Te falt� completar el Apellido y Nombre."
        
        ' Celda del Apellido y Nombre
        ActiveCell.Offset(0, -20).Activate
        Exit Sub
        
    ElseIf codigoPostal = "" Then
        MsgBox "Te falt� completar el C�digo Postal."
        
        ' Celda del CP
        ActiveCell.Offset(0, -3).Activate
        Exit Sub
    
    ElseIf ciudad = "" Then
        MsgBox "Te falt� completar la Ciudad."
        
        ' Celda del Ciudad
        ActiveCell.Offset(0, -2).Activate
        Exit Sub
    
    ElseIf provincia = "" Then
        MsgBox "Te falt� completar la Provincia."
        
        ' Celda del Provincia
        ActiveCell.Offset(0, -1).Activate
        Exit Sub
    
    ElseIf telefono = "" Then
        MsgBox "Te falt� completar el Tel�fono."
        
        ' Celda del Tel�fono
        ActiveCell.Offset(0, -4).Activate
        Exit Sub
    End If


    ' SI ES A DOMICILIO -----------
    If ActiveCell.Value = Sheets("Opciones").Range("A5").Value Then
    
        If direccion = "" Then
            MsgBox "Te falt� completar la Direcci�n."
        
            ' Celda del Direcci�n
            ActiveCell.Offset(0, -5).Activate
            Exit Sub
        End If
        
        ' Colocar los datos
        With Sheets("A Domicilio")
            .Range("C16").Value = UCase(apellidoNombre)
            .Range("P15").Value = dniCuit
            .Range("C18").Value = direccion
            .Range("C23").Value = UCase(provincia)
            .Range("E21").Value = codigoPostal
            .Range("G21").Value = UCase(ciudad)
            .Range("P23").Value = telefono
        End With
        
        ' Generar Proforma
        Call proforma(apellidoNombre, direccion, provincia, codigoPostal, ciudad, telefono, fecha)
        
        ' Generar rotulo
        Call Rotulo_Correo_Argentino(apellidoNombre, Sheets("A Domicilio"), fecha)
        

    ' SI ES A RETIRAR EN SUCURSAL ----------------
    ElseIf ActiveCell.Value = Sheets("Opciones").Range("A3").Value Then
        
        ' Colocar los datos
        With Sheets("A Sucursal")
            .Range("C16").Value = UCase(apellidoNombre)
            .Range("R16").Value = dniCuit
        End With
        
        ' Validando si existe o no el dato.
        On Error Resume Next
        codigoNis = Sheets("Sucursales").Range("$F$4:$F$5000").Find(What:=codigoPostal, LookIn:=xlValues, LookAt:=xlPart).Offset(0, -5)
    
        If codigoNis = "" Then
            ' Todo sali� mal
            Sheets("Planilla").Activate
            MsgBox ("El c�digo postal " & codigoPostal & " no corresponde con ninguna sucursal del Correo. Intent� con otro. ")
            Sheets("Sucursales").Select
            MsgBox ("Busc� aqu� un c�digo postal de sucursal disponible")
            Exit Sub
        End If
        
        ' Completando el resto de datos.
        With Sheets("A Sucursal")
            .Range("S18").Value = codigoNis
            .Range("R22").Value = telefono
        End With
        
        ' Generar Proforma
        Call proforma(apellidoNombre, "Retiro en Sucursal del Correo Argentino C�d. NIS " & codigoNis, provincia, codigoPostal, ciudad, telefono, fecha)
        
        ' Generar rotulo
        Call Rotulo_Correo_Argentino(apellidoNombre, Sheets("A Sucursal"), fecha)
        
    
    ' SI ES PAGO DE FLETE EN DESTINO
    ElseIf ActiveCell.Value = Sheets("Opciones").Range("A4").Value Then

        ' Colocar los datos
        With Sheets("Pago en Destino")
            .Range("C16").Value = UCase(apellidoNombre)
            .Range("R16").Value = dniCuit
        End With
        
        ' Validando si existe o no el dato.
        On Error Resume Next
        codigoNis = Sheets("Sucursales").Range("$F$4:$F$5000").Find(What:=codigoPostal, LookIn:=xlValues, LookAt:=xlPart).Offset(0, -5)
    
        If codigoNis = "" Then
            ' Todo sali� mal
            Sheets("Planilla").Activate
            MsgBox ("El c�digo postal " & codigoPostal & " no corresponde con ninguna sucursal del Correo. Intent� con otro. ")
            Sheets("Sucursales").Activate
            MsgBox ("Busc� aqu� un c�digo postal de sucursal disponible")
            Exit Sub
        End If
        
        ' Completando el resto de datos.
        With Sheets("Pago en Destino")
            .Range("S18").Value = codigoNis
            .Range("R22").Value = telefono
        End With
        
        Sheets("Planilla").Activate
        
        ' Generar Proforma
        Call proforma(apellidoNombre, "Retiro en Sucursal del Correo Argentino C�d. NIS " & codigoNis, provincia, codigoPostal, ciudad, telefono, fecha)
        
        ' Generar rotulo
        Call Rotulo_Correo_Argentino(apellidoNombre, Sheets("Pago en Destino"), fecha)
    
    
    '   SI ES RETIRO EN LOCAL -----------
    ElseIf ActiveCell.Value = Sheets("Opciones").Range("A2").Value Then
        
        ' Colocar los datos
        With Sheets("Retiro en Local")
            .Range("C16").Value = UCase(apellidoNombre)
            .Range("R16").Value = dniCuit
            .Range("R22").Value = telefono
        End With
        
        ' Generar rotulo
        Call Rotulo_Correo_Argentino(apellidoNombre, Sheets("Retiro en Local"), fecha)
        Sheets("Planilla").Activate
    
    End If


End Sub


Function Rotulo_Correo_Argentino(apellidoNombre, rotulo, fecha)
'
' Rotulos Macro
' Guarda los r�tulos en pdf
'
' Acceso directo: CTRL+MAY+�
'   Guarda antes de crear el archivo
    ThisWorkbook.Save
    
    ' Declaraci�n de Variables y su tipo de datos
    Dim nombre As String
    Dim ruta As String
    Dim nombreCarpeta As String

    ' Variables necesarias
    nombre = fecha & ". " & apellidoNombre
    ruta = ThisWorkbook.Path
    nombreCarpeta = "\Rotulos\"
    
    ' Comprobando si existe o no la carpeta
    If Dir(ruta, vbDirectory) <> "" Then
        If Dir(ruta & nombreCarpeta, vbDirectory) = "" Then
            MkDir ruta & nombreCarpeta
        End If
    End If
    
    ' Generando el archivo pdf
    rotulo.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
        ruta & nombreCarpeta & UCase(nombre) & ".pdf", _
        OpenAfterPublish:=True
    Sheets("Planilla").Activate
End Function

Function Validar_CP(cp)
' Valida si el c�digo postal es correcto, existe o no.

    codigoNis = Sheets("Sucursales").Range("F1576:F5000").Find(What:=cp, LookIn:=xlValues, SearchOrder:=xlByRows, LookAt:=xlWhole).Offset(0, -5)
    
    If codigoNis = "" Then
        ' Todo sali� mal
        Sheets("Planilla").Activate
        MsgBox ("El c�digo postal " & cp & " no existe. Intent� con otro. " & codigoNis)
        Exit Function
    End If
End Function
Sub marcar()
    ' MARCA CON UN COLOR TODA UNA SELECCION. O LO VUELVE A LA NORMALIDAD
    Dim rojo, verde, azul, inferior, superior, filas As Byte
    
    inferior = 150
    superior = 255
    rojo = Int(inferior + Rnd * (superior - inferior + 1))
    verde = Int(inferior + Rnd * (superior - inferior + 1))
    azul = Int(inferior + Rnd * (superior - inferior + 1))
    filas = Selection.Rows.Count - 1
    
    If Selection.Row > 1 And (Selection.Row + filas) < 35 Then
        ActiveSheet.Unprotect "Rerda"
        Selection.Interior.color = RGB(rojo, verde, azul)
        ActiveSheet.Protect "Rerda"
    ElseIf Selection.Row = 1 Or (Selection.Row + filas) > 34 Then
        ' NADA
        MsgBox "La 1� fila de t�tulos no se selecciona.", vbCritical, "�Guarda!"
    Else
        MsgBox "Ten�s que seleccionar algo antes " & vbNewLine & "que est� entre el t�tulo y el pi�."
    End If
    Debug.Print Selection.Row + filas
End Sub
Sub desmarcar()
    ' VUELVE A LA NORMALIDAD
    If Selection.Row = 1 Then
        MsgBox "La 1� fila de t�tulos no se selecciona.", vbCritical, "�Guarda!"
    Else
        ActiveSheet.Unprotect "Rerda"
        Selection.Interior.ColorIndex = xlNone
        ActiveSheet.Protect "Rerda"
    End If
End Sub
' Sirve para controlar si corresponde o no una factura proforma
Function proforma(apellidoNombre, direccion, provincia, codigoPostal, ciudad, telefono, fecha)
    Dim acumulador As Byte
    Dim cantidad As Byte
    Dim precio As Double
    Dim sku As String
    Dim color As String
    Dim talle As String
    Dim cotizacion As Double
    cotizacion = 0
        
    Do While cotizacion = 0
        cotizacion = Application.InputBox(Prompt:="Cotizaci�n del d�lar", Title:="Factura Proforma", Default:=1)
    Loop
    
    'Limpiando informaci�n previa
    With Sheets("Proforma")
        .Range("A21:D49").ClearContents
        .Range("H21:H49").ClearContents
        .Range("I7:I14").ClearContents
        .Range("I17:I18").ClearContents
    End With
    
    ' Dando formato a la p�gina para imprimir
    With Sheets("Proforma").PageSetup
        .Orientation = xlPortrait
        .TopMargin = Application.CentimetersToPoints(1.9)
        .RightMargin = Application.CentimetersToPoints(0.6)
        .BottomMargin = Application.CentimetersToPoints(1.9)
        .LeftMargin = Application.CentimetersToPoints(0.6)
        .HeaderMargin = Application.CentimetersToPoints(0.8)
        .FooterMargin = Application.CentimetersToPoints(0.8)
        .CenterHorizontally = True
        .PaperSize = xlPaperA4
    End With
    
    If provincia = "TIERRA DEL FUEGO" Or provincia = "Tierra del Fuego" Then
        'hacer proforma
        With Sheets("Proforma")
            .Cells(7, 9).Value = UCase(apellidoNombre)
            .Cells(9, 9).Value = direccion
            .Cells(11, 9).Value = UCase(ciudad)
            .Cells(12, 9).Value = codigoPostal
            .Cells(13, 9).Value = UCase(provincia)
            .Cells(17, 9).Value = "'" & telefono
        End With
        acumulador = 0
        
        ' Bucle que recorre la venta
        Worksheets("Planilla").Activate
        ActiveCell.Offset(0, -19).Activate
                
        Do While ActiveCell.Offset(0, -1).Value = apellidoNombre Or ActiveCell.Offset(0, -1).Value = ""
        
            ' Toma datos de la "Planilla"
            With Sheets("Planilla")
                sku = ActiveCell.Value
                cantidad = ActiveCell.Offset(0, 4).Value
                color = ActiveCell.Offset(0, 3).Value
                talle = ActiveCell.Offset(0, 2).Value
                precio = ActiveCell.Offset(0, 5).Value
            End With
            Debug.Print sku, cantidad, color, talle, precio
            
            ' Los vuelca en la Proforma
            With Sheets("Proforma")
                ' Copiando el c�digo
                .Cells(21 + acumulador, 2).Value = sku
            
                ' Copiando el talle
                .Cells(21 + acumulador, 4).Value = talle
            
                ' Copiando el color
                .Cells(21 + acumulador, 3).Value = color
            
                ' Copiando la cantidad
                .Cells(21 + acumulador, 1).Value = cantidad
            
                ' Copiando el monto
                .Cells(21 + acumulador, 8).Value = precio / cotizacion
            End With
                        
            ' Aumentando el contador
            acumulador = acumulador + 1
            
            ' Saltando de celda
            ActiveCell.Offset(1, 0).Activate
            
            ' Salir cuando se acabe todo
            If ActiveCell.Value = "" Then
                Exit Do
            End If
        Loop
        
        ' Imprimiendo el r�tulo
        Call Rotulo_Correo_Argentino("Factura Proforma - " & apellidoNombre, Sheets("Proforma"), fecha)
    End If
End Function

Sub CorreoArgentino()
' GENERADOR DE HOJA DEL CORREO ARGENTINO

' Variables necesarias
Dim planillaGeneradora As Workbook
Dim nombreArchivoCorreo As String
Dim ruta As String
Dim nombreCarpeta As String
Dim i As Byte
Dim e As Byte
Dim comienzo As Byte

e = 0
comienzo = 13

Set planillaGeneradora = ThisWorkbook


nombreArchivoCorreo = "CorreoArgentino.xlsm"


ruta = planillaGeneradora.Path & "\" & nombreArchivoCorreo

' Controla si existe o no el archivo en la misma carpeta
If Dir(ruta) = "" Then
    MsgBox "El archivo " & nombreArchivoCorreo & " debe estar en la misma carpeta que esta planilla."
    Exit Sub
End If

' Abrir el archivo
Workbooks.Open ruta

' Borra el contenido
For i = 13 To 32
    Cells(i, 3) = ""
    Cells(i, 10) = ""
Next i

' Copia contenido: Denominaci�n - DNI/CUIT: - CP: - Provincia - Vendedor/Viajante
    
    ' Si hay m�s de 20, se genera un nuevo archivo.
    For i = 2 To 34
        If planillaGeneradora.Worksheets("Planilla").Cells(i, 2) <> "" Then
            ' Datos del nombre/apellido y dem�s
            Cells(comienzo + e, 3).Value = UCase(planillaGeneradora.Worksheets("Planilla").Cells(i, 2).Value) & " - DNI/CUIT: " & planillaGeneradora.Worksheets("Planilla").Cells(i, 12).Value & " - CP " & planillaGeneradora.Worksheets("Planilla").Cells(i, 19).Value & " - " & planillaGeneradora.Worksheets("Planilla").Cells(i, 21).Value
            
            ' Dato del Vendedor
            Cells(comienzo + e, 10).Value = planillaGeneradora.Worksheets("Planilla").Cells(2, 23).Value
        
            ' Incrementamos en 1 el contador "e"
            e = e + 1
        End If
        
        
    
    
        If e > 20 Then
            ThisWorkbook.Save
            MsgBox "Te sobrepasaste de 20 renglones. Guard� este documento y hac� otro m�s."
            Exit Sub
        End If
    
    Next i

End Sub


Sub deposito()
' GENERA UNA PLANILLA S�LO PARA USO EXCLUSIVO DEL DEPOSITO
' Declaraci�n de Variables y su tipo de datos
Dim ruta As String
Dim ultima As Byte
Dim i As Byte
Dim NombreDeArchivo As String
Dim ExisteArchivo As String
Dim nombre As String
Dim nombreCarpeta As String

' Variables necesarias
nombreCarpeta = ThisWorkbook.Path & "\"
ultima = 34
ruta = "'" & nombreCarpeta & "[Stock.XLS]Sheet1'!$A$2:$G$10000"

NombreDeArchivo = nombreCarpeta & "Stock.XLS"
ExisteArchivo = Dir(NombreDeArchivo)

' Comprueba si existe el archivo Stock.XLS
If ExisteArchivo = "" Then
    MsgBox "El archivo Stock.XLS debe estar en la misma carpeta que esta planilla"
    Exit Sub
End If

Sheets("Dep�sito").Activate

' Reemplaza la que hubiere



' Creando las columnas
Cells(1, 1).Value = "Cliente"
Cells(1, 2).Value = "Descripci�n"
Cells(1, 3).Value = "C�digo"
Cells(1, 4).Value = "Color"
Cells(1, 5).Value = "Talle"
Cells(1, 6).Value = "Cantidad"
Cells(1, 7).Value = "Ubicaci�n"

' Completando los datos
For i = 2 To ultima
    ' Cliente
    Cells(i, 1).Value = Sheets("Planilla").Cells(i, 2).Value
    
    ' Descripci�n
    Cells(i, 2).Value = Sheets("Planilla").Cells(i, 4).Formula
    
    ' C�digo
    Cells(i, 3).Value = Sheets("Planilla").Cells(i, 3).Value
    
    ' Color
    Cells(i, 4).Value = Sheets("Planilla").Cells(i, 6).Value
    
    ' Talle
    Cells(i, 5).Value = Sheets("Planilla").Cells(i, 5).Value
    
    ' Cantidad
    Cells(i, 6).Value = Sheets("Planilla").Cells(i, 7).Value
        
    ' La ubicaci�n
    If Cells(i, 3) = "" Then
        Cells(i, 7).Value = ""
    Else
        Cells(i, 7).Formula = "=VLOOKUP(C" & i & "," & ruta & ",7,FALSE)"
    End If
Next i

' Ordenando alfab�ticamente esta columna de ubicaci�n
With Range("A1:G1")
    .AutoFilter
    .Rows("1").RowHeight = 27
    .Font.Bold = True
    .Font.Size = 12
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
End With

Range("A1").CurrentRegion.Sort Key1:=Range("G1"), Order1:=xlAscending, Header:=xlGuess
With Selection
    .AutoFilter
End With

With Range("A1").CurrentRegion
    .Columns.AutoFit
End With

' Colocando totales de productos y dando formato
Cells(ultima + 1, 5).Value = "TOTALES:"
Cells(ultima + 1, 6).Select
Cells(ultima + 1, 6).Value = "=SUM(F2:F" & ultima & ")"
Range(Cells(ultima + 1, 5), Cells(ultima + 1, 6)).Select
With Selection
    .Font.Bold = True
    .Font.Size = 15
    .HorizontalAlignment = xlRight
    .VerticalAlignment = xlBottom
End With


' Colocando el total de r�tulos a imprimir
Cells(ultima + 1, 2).Value = "ROTULOS:"
Cells(ultima + 1, 3).Value = "=COUNTA(A2:A" & ultima & ")"
Range(Cells(ultima + 1, 2), Cells(ultima + 1, 3)).Select
With Selection
    .Font.Bold = True
    .Font.Size = 15
    .VerticalAlignment = xlBottom
    .HorizontalAlignment = xlRight
End With
Cells(ultima + 1, 3).HorizontalAlignment = xlLeft

With Range("A1").CurrentRegion
    .Borders.LineStyle = xlContinuous
End With

' Formato de impresi�n
With ActiveSheet.PageSetup
    .Orientation = xlPortrait
    .PaperSize = xlPaperA4
    .LeftMargin = Application.CentimetersToPoints(0.64)
    .RightMargin = Application.CentimetersToPoints(0.64)
    .TopMargin = Application.CentimetersToPoints(4)
    .BottomMargin = Application.CentimetersToPoints(1.91)
    .HeaderMargin = Application.CentimetersToPoints(0.76)
    .FooterMargin = Application.CentimetersToPoints(0.76)
    .CenterHorizontally = True
    .CenterVertically = False
    .PrintArea = ActiveSheet.Range("A1:G" & ultima + 1).Address
    .Zoom = False
    .FitToPagesTall = 1
    .FitToPagesWide = 1
    .CenterHeader = "&B&20&F" & vbNewLine & "SOLO PARA USO EN DEPOSITO"
End With

ThisWorkbook.Save

End Sub
Sub construirCtaCte()
' CONSTRUYE LA PLANILLA DE CUENTA CORRIENTE

' Nombre Planilla de Ventas
Dim planillaVentas As String
planillaVentas = ThisWorkbook.Name

' Carpeta actual de Planilla de Ventas
Dim carpetaActual As String
carpetaActual = Workbooks(planillaVentas).Path

' Creaci�n archivo Cuentas Corrientes
Dim archivoCtaCte As String
archivoCtaCte = Year(Date) & ". CTACTE.xlsx"

' Control si existe de antes
If Dir(carpetaActual & "\..\" & archivoCtaCte, vbNormal) = "" Then
    Workbooks.Add.SaveAs Filename:=(carpetaActual & "\..\" & archivoCtaCte)
Else
    ' Workbooks.Open (carpetaActual & "\..\" & archivoCtaCte)
    MsgBox "Ya est� creada la planilla de las cuentas corrientes de antes." & vbNewLine & "Esper� al a�o que viene."
    Exit Sub
End If


' Arreglo con la definici�n pernsolizada de los meses
Dim arrayMeses As Variant
arrayMeses = Array("ENE", "FEB", "MAR", "ABR", "MAY", "JUN", "JUL", "AGO", "SEP", "OCT", "NOV", "DIC")

' Construcci�n de las hojas
Dim i As Byte
i = 12
Do While i <= 12
    If i = 0 Then
        GoTo Salida
        ' Sale del bucle
    ElseIf i = 12 Then
        ActiveSheet.Name = arrayMeses(i - 1) & "-" & arrayMeses(0)
    Else
        Worksheets.Add.Name = arrayMeses(i - 1) & "-" & arrayMeses(i)
    End If
    i = i - 1
    Call construirHoja
Loop
Salida:

Workbooks(archivoCtaCte).Close SaveChanges:=True
End Sub

Function construirHoja()
Dim titulares As Variant
Dim i As Byte
titulares = Array("FAC NUM", "FECHA", "DNI", "CUENTA", "CBU", "CLIENTE", "IMPO.TOTAL", "CUOTAS", "IMP.CUOTA", "TELEFONO", "DOMICILIO", "LOCALIDAD", "PROVINCIA", "VENDEDOR")
With ActiveSheet
    For i = 0 To UBound(titulares)
        With Cells(1, i + 1)
            .Value = titulares(i)
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .Interior.ColorIndex = 15
            .Borders.LineStyle = xlContinuous
        End With
    Next i
End With

End Function

Sub CTACTE()
' CARGA DATOS DE LA PLANILLA EN EL ARCHIVO DE LAS CUENTAS CORRIENTES

' Nombre Planilla de Ventas
Dim planillaVentas As String
planillaVentas = ThisWorkbook.Name

' Carpeta actual de Planilla de Ventas
Dim carpetaActual As String
carpetaActual = Workbooks(planillaVentas).Path

' Nombre archivo Cuentas Corrientes
Dim archivoCtaCte As String
archivoCtaCte = Year(Date) & ". CTACTE.xlsx"

' Pesta�a del archivo Cuentas Corrientes
Dim arrayMeses As Variant
arrayMeses = Array("ENE-FEB", "FEB-MAR", "MAR-ABR", "ABR-MAY", "MAY-JUN", "JUN-JUL", "JUL-AGO", "AGO-SEP", "SEP-OCT", "OCT-NOV", "NOV-DIC", "DIC-ENE")
Dim mesCtaCte As String

' Sirve para posicionar la pesta�a al comenzar
If Day(Date) < 21 Then
    mesCtaCte = Month(Date + 1)
Else
    mesCtaCte = Month(Date)
End If
Debug.Print mesCtaCte

' Abrir archivo Cuentas Corrientes
Workbooks.Open(carpetaActual & "\..\" & archivoCtaCte).Sheets(arrayMeses(mesCtaCte - 1)).Activate
'Workbooks.Open(carpetaActual & "\" & archivoCtaCte).Sheets(arrayMeses(mesCtaCte - 1)).Activate

' Ultima fila de la pesta�a actual
Dim ultimaFila As Byte

' Copiando informaci�n de la planilla al archivo
Dim i As Byte
For i = 2 To 34
    ' Siempre actualizar la �ltima fila
    ultimaFila = Workbooks(archivoCtaCte).Sheets(arrayMeses(mesCtaCte - 1)).Cells(Rows.Count, 1).End(xlUp).Row + 1
    Debug.Print ultimaFila
    Debug.Print Workbooks(planillaVentas).Sheets("Planilla").Cells(i, 13).Value
    
    If Workbooks(planillaVentas).Worksheets("Planilla").Cells(i, 13).Value = Workbooks(planillaVentas).Sheets("Opciones").Range("C2").Value Then
        
        With Workbooks(archivoCtaCte).Sheets(arrayMeses(mesCtaCte - 1))
            ' N� Factura
            .Cells(ultimaFila, 1).Value = Workbooks(planillaVentas).Sheets("Planilla").Cells(i, 11).Value
            
            ' Fecha
            .Cells(ultimaFila, 2).Value = Workbooks(planillaVentas).Sheets("Planilla").Cells(i, 1).Value
            
            ' DNI
            .Cells(ultimaFila, 3).Value = Workbooks(planillaVentas).Sheets("Planilla").Cells(i, 12).Value
            
            ' N� Cuenta de Caja de Ahorro
            .Cells(ultimaFila, 4).Value = "'" & Workbooks(planillaVentas).Sheets("Planilla").Cells(i, 15).Value
            
            ' CBU
            .Cells(ultimaFila, 5).Value = "'" & Workbooks(planillaVentas).Sheets("Planilla").Cells(i, 16).Value
            
            ' Cliente
            .Cells(ultimaFila, 6).Value = Workbooks(planillaVentas).Sheets("Planilla").Cells(i, 2).Value
            
            ' Importe total de la factura
            .Cells(ultimaFila, 7).Value = CCur(Workbooks(planillaVentas).Sheets("Planilla").Cells(i, 9).Value)
            
            ' Cantidad de Cuotas
            .Cells(ultimaFila, 8).Value = Workbooks(planillaVentas).Sheets("Planilla").Cells(i, 14).Value
            
            ' Importe de la cuota
            .Cells(ultimaFila, 9).Value = CCur(Cells(ultimaFila, 7).Value / Cells(ultimaFila, 8).Value)
            
            ' Tel�fono
            .Cells(ultimaFila, 10).Value = Workbooks(planillaVentas).Sheets("Planilla").Cells(i, 18).Value
            
            ' Domicilio
            .Cells(ultimaFila, 11).Value = Workbooks(planillaVentas).Sheets("Planilla").Cells(i, 17).Value & " - CP: " & Workbooks(planillaVentas).Sheets("Planilla").Cells(i, 19).Value
            
            ' Localidad
            .Cells(ultimaFila, 12).Value = Workbooks(planillaVentas).Sheets("Planilla").Cells(i, 20).Value
            
            ' Provincia
            .Cells(ultimaFila, 13).Value = Workbooks(planillaVentas).Sheets("Planilla").Cells(i, 21).Value
            
            ' Vendedor
            .Cells(ultimaFila, 14).Value = Workbooks(planillaVentas).Sheets("Planilla").Cells(2, 23).Value
        End With
    End If
Next i

ActiveSheet.UsedRange.EntireColumn.AutoFit
ActiveSheet.UsedRange.EntireRow.AutoFit
Range("A1").Activate

End Sub
