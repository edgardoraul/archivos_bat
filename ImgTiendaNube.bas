Attribute VB_Name = "ImgTiendaNube"
Option Explicit

' Definiendo las variables
Dim URL As String
Dim CantS As Integer
Dim codigo As Long
Dim ultimaFila As Integer
Dim i As Integer
Dim e As Integer
Dim extension As String
Dim imagenes As String
Dim contador As Integer
Dim tabla As String
Dim conglomerado As String
Dim color As String
Dim acumulado As String
Dim cuenta As Integer
Dim tipo As String
Dim ruta As String
Dim rutaImgRenombradas As String
Dim subcarpeta As String
Dim origen As String
Dim archivoNuevo As String
Dim cantidadImg As Integer
Dim destino As String
Dim archivoAntiguo As String
Dim xPath As String
Dim xFile As String
Dim xCount As Integer
Dim cantidad As String
Const rutaOrigenImg = "D:\Web\imagenes_rerda\"
Const rutaFinal = "D:\OneDrive\Dragonfish Color y Talle\Articulos\"



' Una función para obtener las rutas de carpetas
Function OBTENER_RUTA_CARPETA_ARCHIVO(ruta As String) As String
    Set objeto = New FileSystemObject
    Set archivo = objeto.GetFile(ruta)
    OBTENER_RUTA_CARPETA_ARCHIVO = archivo.ParentFolder.Path
End Function


Sub GeneradorImagenesVariables()
' PRODUCTO CON VARIENTE DE TALLE ===================

Sheets("Variables").Activate


' Rellenando la url
URL = Sheets("Constantes").Range("B1").Value
extension = ".jpg"
tabla = "/tabla" & extension


' Obteniendo la última fila
Range("A1").Select
ultimaFila = Range(Selection, Selection.End(xlDown)).Count

' Bucle que recorre toda la columna
For i = 2 To ultimaFila
    If Cells(i, 8) >= 1 Then
        ' Asigando la cantidad de imágenes que tiene el producto
        CantS = Cells(i, 8).Value
        imagenes = ""
        
        For contador = 1 To CantS
            imagenes = imagenes & "," & URL & Cells(i, 6).Value & "/" & contador & extension
        Next
                
        ' Controlando si tiene tabla
        If Cells(i, 9).Value = 1 Then
            imagenes = imagenes & "," & URL & Cells(i, 6).Value & tabla
        End If
        
        ' Insertando el resultado completo de las imágenes
        Cells(i, 7).Value = Right(imagenes, Len(imagenes) - 1)
    End If
Next

ThisWorkbook.Save
End Sub


Sub generadorImagenesConColor()
' PRODUCTO CON VARIANTE DE COLOR ===================

Sheets("Con Color").Activate

' Rellenando la url
URL = Sheets("Constantes").Range("B1").Value
extension = ".jpg"

' Obteniendo la última fila
Range("A1").Select
ultimaFila = Range(Selection, Selection.End(xlDown)).Count

' Ordenando de mayor a menor el ID para facilitar la construcción de los enlaces del Padre.
Range("A1").Sort Key1:=Range("A1"), Order1:=xlDescending, Header:=xlNo

' Bucle que recorre toda la columna
For i = 2 To ultimaFila
    
    ' Contando cuantas veces se repite el código del producto
    If Cells(i, 4).Value = "Padre" Then
        Cells(i, 8).Value = Application.CountIf(ActiveSheet.Range("F2:F" & ultimaFila), Cells(i, 6).Value) - 1
    End If

    ' Generando los enlaces
    If Cells(i, 8).Value >= 1 Then
            
        ' Asigando la cantidad de imágenes que tiene el producto
        CantS = Cells(i, 8).Value
        imagenes = ""
        
        For contador = 1 To CantS
            ' Acumula y nombra las imágenes en base primero a su código de color, seguido de su orden
            imagenes = imagenes & "," & URL & Cells(i, 6).Value & "/" & Cells(i, 4).Value & contador & extension
        Next
        
        ' Insertando el resultado completo de las imágenes
        Cells(i, 7).Value = Right(imagenes, Len(imagenes) - 1)
        
        ' Insertando un nuevo loop para los artículos que son Padres
        If Cells(i, 4).Value = "Padre" Then
            ' Pruebo colocando el último valor de la fila anterior
            Cells(i, 7).Value = Cells(i - 1, 7).Value
        End If
    End If
    
    ' Otro loop para colocar la sumatoria concatenada de todas las imágenes de las
    ' variantes en el padre. Corroborando primero si es un Padre.
    If Cells(i, 4).Value = "Padre" Then
        acumulado = ""
        cuenta = Cells(i, 8).Value
        For e = 1 To cuenta
            acumulado = acumulado & "," & Cells(i - e, 7).Value
        Next
        
        ' Insertando el valor acumulado en la celda y eliminando la última coma.
        Cells(i, 7).Value = Right(acumulado, Len(acumulado) - 1)
    End If

Next
Range("A1").Sort Key1:=Range("A1"), Order1:=xlAscending, Header:=xlYes
ThisWorkbook.Save

End Sub

Sub generadorImagenesConTalle()
' PRODUCTO CON VARIANTE DE TALLE ===================

Worksheets("Con Talles").Activate

' Rellenando la url
URL = Worksheets("Constantes").Range("B1").Value
extension = ".jpg"

' Obteniendo la última fila
Cells(2, 1).Activate
ultimaFila = Cells(Rows.Count, 1).End(xlUp).Row

' Ordenando de mayor a menor el ID para facilitar la construcción de los enlaces del Padre.
Range("A1").Sort Key1:=Range("A1"), Order1:=xlDescending, Header:=xlNo

' Bucle que recorre toda la columna
For i = 2 To ultimaFila
    
    ' Generando los enlaces
    If Cells(i, 8).Value >= 1 Then
            
        ' Asigando la cantidad de imágenes que tiene el producto
        CantS = Cells(i, 8).Value
        imagenes = ""
        
        For contador = 1 To CantS
            ' Acumula y nombra las imágenes en base primero a su código de color, seguido de su orden
            imagenes = imagenes & "," & URL & Cells(i, 6).Value & "/" & Cells(i, 4).Value & contador & extension
        Next
        
        ' Insertando el resultado completo de las imágenes
        Cells(i, 7).Value = Right(imagenes, Len(imagenes) - 1)
        
       
    End If
    
    ' Otro loop para colocar la sumatoria concatenada de todas las imágenes de las
    ' variantes en el padre. Corroborando primero si es un Padre.
    If Cells(i, 4).Value = "Padre" Then
        acumulado = ""
        cuenta = Cells(i, 8).Value
        For e = 1 To cuenta
            acumulado = acumulado & "," & Cells(i - e, 7).Value
        Next
        
        ' Insertando el valor acumulado en la celda y eliminando la última coma.
        Cells(i, 7).Value = Right(acumulado, Len(acumulado) - 1)
    End If

Next
Range("A1").Sort Key1:=Range("A1"), Order1:=xlAscending, Header:=xlYes
ThisWorkbook.Save

End Sub

Sub GeneradorImagenesSimples()
' PRODUCTO SIMPLE ===================

Sheets("Simples").Activate

' Rellenando la url
URL = Sheets("Constantes").Range("B1").Value
extension = ".jpg"

' Obteniendo la última fila
Range("A1").Select
ultimaFila = Range(Selection, Selection.End(xlDown)).Count

' Bucle que recorre toda la columna
For i = 2 To ultimaFila
    
    If Cells(i, 8) >= 1 Then
        ' Asigando la cantidad de imágenes que tiene el producto
        CantS = Cells(i, 8).Value
        imagenes = ""
        
        For contador = 1 To CantS
            imagenes = imagenes & "," & URL & Cells(i, 6).Value & "/" & contador & extension
        Next
        
        ' Insertando el resultado completo de las imágenes
        Cells(i, 7).Value = Right(imagenes, Len(imagenes) - 1)
    End If
Next

ThisWorkbook.Save
End Sub


Sub copiarImgVariables()
    ' DESCRIPCION: Copia y renombra imágenes con variantes de talles
    ' Creamos coordenadas para trabajar
    Dim subcarpeta As String
    Dim Carpeta As String
    Dim archivoViejo As String
    Dim archivoNuevo As String
    Dim origen As String
    Dim destino As String
    Dim fs As Object
    Dim cantidadImg As Integer
    Dim skuActual As String
    Dim skuAnterior As String
    Dim codigoActual As String
    Dim codigoAnterior As String
    Dim Seguir As String
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    Range("A1").Select
    ultimaFila = Range(Selection, Selection.End(xlDown)).Count
    
    extension = ".jpg"

    ruta = rutaOrigenImg

    rutaImgRenombradas = rutaFinal
    
    Cells(ultimaFila + 2, 5).Activate

    ' Bucle para recorrer toda la columna de los códigos y todas las carpetas con las imágenes
    For i = 2 To ultimaFila
    
        ' Posicionándose en lo que importa
        Cells(ultimaFila + 2, 5).Value = "Procedados: " & i & " de " & ultimaFila
        
        ' Código -> Corresponde a la carpeta en la que están las imágenes numeradas
        subcarpeta = Cells(i, 6).Value
        
        ' SKU limpio. Corresponde al código en si mismo que tiene el producto
        skuActual = Left(Cells(i, 3).Value, 7)
        codigoActual = Cells(i, 6).Value
        
        
        ' Controlando si sku actual tiene el mismo código que el sku anterior
        If skuActual = codigoActual Then
            Debug.Print "El sku actual " & skuActual & " coincide con el código " & codigoActual
        ElseIf skuActual <> codigoActual Then
            Debug.Print "El sku actual " & skuActual & " es talle grande del código " & codigoActual
            subcarpeta = skuActual
            FSO.CopyFolder (ruta & codigoActual), (ruta & subcarpeta), True
        End If
        
        ' Contar la cantidad de imágenes que hay una carpeta determinada
        xPath = ruta & subcarpeta & "\*" & extension
        xFile = Dir(xPath)
        
        xCount = 0
        Do While xFile <> ""
            xCount = xCount + 1
            If xFile = "tabla.jpg" Then
                xCount = xCount - 1
                Cells(i, 9).Value = 1
            End If
            xFile = Dir()
        Loop
        
        ' Insertando el resultado encontrado en la planilla
        Cells(i, 8).Value = xCount
        
        ' Extrayendo de la planilla la cantidad de imágenes
        cantidadImg = xCount
        
        Debug.Print "El Código " & subcarpeta & " tiene " & cantidadImg & " imágenes."
        
        If cantidadImg < 1 Then
            Cells(i, 10).Value = "El código " & subcarpeta & " no tiene imágenes."
            Cells(i, 8).Value = ""
            GoTo Seguir
        Else
            Cells(i, 10).Value = ""
        End If
        
        Debug.Print "Estamos en la fila N° " & i
        
        
        
        ' Controlando si tiene tabla de talles
        If Cells(i, 9).Value = 1 Then
            cantidadImg = cantidadImg + 1
        End If
        
        
        ' Creando nuevos nombres de archivos de fotos mediante bucle
        For e = 1 To cantidadImg
            
            ' Carpeta y nombre de archivo de Origen
            If e = cantidadImg And Cells(i, 9).Value = 1 Then
                origen = ruta & subcarpeta & "\" & "tabla" & extension
            Else
                origen = ruta & subcarpeta & "\" & e & extension
            End If
            
            ' Nuevo nombre de archivo
            archivoNuevo = skuActual & "'''" & e & extension
            
            ' Carpeta y nombre nuevo de destino
            destino = rutaImgRenombradas & archivoNuevo
            FileCopy origen, destino
            Debug.Print origen & " está copiado como " & destino
            
        Next
Seguir:
    Next
    

    
End Sub
Sub copiarImgColor()
' DESCRIPCION: Copia y renombra imágenes con variantes de COLOR

' Creamos coordenadas para trabajar
Range("A1").Select
ultimaFila = Range(Selection, Selection.End(xlDown)).Count
extension = ".jpg"
ruta = rutaOrigenImg
rutaImgRenombradas = rutaFinal

Debug.Print ultimaFila

' Copiando Imágenes. Recorremos toda la tabla desde arriba hasta abajo
For i = 2 To ultimaFila
    'Definiendo la cantidad de imágenes que tiene esta variante
    codigo = Cells(i, 6).Value
    xPath = ruta & codigo & "\*" & extension
    xFile = Dir(xPath)
    If xFile = "" Then
        Cells(i, 8).Value = "Sin imágenes"
        GoTo Seguir
    End If
    
    'Averiguando cuántas imágenes hay en la variante o padre seleccionada
    xCount = 0
    Do While xFile <> ""
        'Si coincide el nombre del archivo con el color
        If Left(xFile, 2) = Cells(i, 4).Value Then
            xCount = xCount + 1
            Cells(i, (8 + xCount)).Value = xFile
        End If
        ' Aquí ya cambia de valor
        xFile = Dir()
    Loop
    
    
    'Extrayendo la cantidad de imágenes que tiene cada publicación
    cantidadImg = xCount
    
    'Anotando el resultado
    Cells(i, 8).Value = cantidadImg
    
    'Renombrando cada imagen y copiándola al destino
    For e = 1 To cantidadImg

        ' Foto de portada. Una sola.
        If Cells(i, (8 + e)).Value = "1.jpg" Then
            archivoAntiguo = Cells(i, (8 + e)).Value
            origen = ruta & codigo & "\" & archivoAntiguo
            archivoNuevo = codigo & "'''" & extension
        
        ElseIf Cells(i, (8 + e)).Value = "2.jpg" Then
            archivoAntiguo = Cells(i, (8 + e)).Value
            origen = ruta & codigo & "\" & archivoAntiguo
            archivoNuevo = codigo & "'''1" & extension
        
        Else
            ' Fotos de las variantes de color
            color = Left(Cells(i, (8 + e)).Value, 2)
            cantidad = Mid(Cells(i, (8 + e)).Value, 3, (Len(Cells(i, (8 + e))) - Len(extension) - 2))
            archivoAntiguo = color & cantidad & extension
            origen = ruta & codigo & "\" & archivoAntiguo
            archivoNuevo = codigo & "'" & color & "''" & cantidad & extension
        End If
        
        'Definiendo el destino final del archivo de la imagen
        destino = rutaImgRenombradas & archivoNuevo
        
        'Copiando el achivo con el nuevo nombre
        Debug.Print "Fila " & i & " tiene " & cantidadImg & " # " & archivoAntiguo & " -> " & archivoNuevo
               
        FileCopy origen, destino
        
        Debug.Print origen & " -> " & destino
Seguir:
    Next
    
Next
  
Call generadorImagenesConColor
End Sub

Sub copiarImgTalle()

' DESCRIPCION: Copia y renombra imágenes con variantes de TALLE
' Creamos coordenadas para trabajar
Worksheets("Con Talles").Activate
Cells(2, 1).Activate
ultimaFila = Cells(Rows.Count, 1).End(xlUp).Row

extension = ".jpg"
ruta = rutaOrigenImg
rutaImgRenombradas = rutaFinal

Debug.Print ultimaFila

' Copiando Imágenes. Recorremos toda la tabla desde arriba hasta abajo
For i = 2 To ultimaFila
    'Definiendo la cantidad de imágenes que tiene esta variante
    codigo = Cells(i, 6).Value
    xPath = ruta & codigo & "\*" & extension
    xFile = Dir(xPath)
    If xFile = "" Then
        Cells(i, 8).Value = "Sin imágenes"
        GoTo Seguir
    End If
    
    'Averiguando cuántas imágenes hay en la variante
    xCount = 0
    Do While xFile <> ""
        'Si coincide el nombre del archivo con el talle
        Debug.Print Left(xFile, InStr(xFile, ".") - 2)
        If Left(xFile, InStr(xFile, ".") - 2) = Cells(i, 4).Value Then
            xCount = xCount + 1
            Cells(i, (8 + xCount)).Value = xFile
        End If
        ' Aquí ya cambia de valor
        xFile = Dir()
    Loop
    
    
    'Extrayendo la cantidad de imágenes que tiene cada publicación
    cantidadImg = xCount
    
    'Anotando el resultado
    Cells(i, 8).Value = cantidadImg
    
    'Renombrando cada imagen y copiándola al destino
    For e = 1 To cantidadImg

        ' Foto de portada. Una sola.
        If Cells(i, (8 + e)).Value = "1.jpg" Then
            archivoAntiguo = Cells(i, (8 + e)).Value
            origen = ruta & codigo & "\" & archivoAntiguo
            archivoNuevo = codigo & "'''" & extension
        
        ElseIf Cells(i, (8 + e)).Value = "2.jpg" Then
            archivoAntiguo = Cells(i, (8 + e)).Value
            origen = ruta & codigo & "\" & archivoAntiguo
            archivoNuevo = codigo & "'''1" & extension
        
        Else
            ' Fotos de las variantes de color
            color = Left(Cells(i, (8 + e)).Value, 2)
            cantidad = Mid(Cells(i, (8 + e)).Value, 3, (Len(Cells(i, (8 + e))) - Len(extension) - 2))
            archivoAntiguo = color & cantidad & extension
            origen = ruta & codigo & "\" & archivoAntiguo
            archivoNuevo = codigo & "''" & color & "'" & cantidad & extension
        End If
        
        'Definiendo el destino final del archivo de la imagen
        destino = rutaImgRenombradas & archivoNuevo
        
        'Copiando el achivo con el nuevo nombre
        Debug.Print "Fila " & i & " tiene " & cantidadImg & " # " & archivoAntiguo & " -> " & archivoNuevo
               
        FileCopy origen, destino
        
        Debug.Print origen & " -> " & destino
Seguir:
    Next
    
Next
Call generadorImagenesConTalle
End Sub



Sub CopiarPrimeraImagenComo1jpg()
    Dim FSO As Object
    Dim CarpetaBase As String
    Dim Carpeta As Object, subcarpeta As Object
    Dim archivo As Object, primerImagen As String
    Dim dialogo As FileDialog
    Dim extensionesImagenes As Variant
    Dim encontrado As Boolean
    
    extensionesImagenes = Array("jpg")

    ' Elegir la carpeta base
    Set dialogo = Application.FileDialog(msoFileDialogFolderPicker)
    dialogo.Title = "Seleccionar carpeta base"
    
    If dialogo.Show <> -1 Then
        MsgBox "No se seleccionó ninguna carpeta.", vbExclamation
        Exit Sub
    End If
    
    CarpetaBase = dialogo.SelectedItems(1)
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    If Not FSO.FolderExists(CarpetaBase) Then
        MsgBox "La carpeta no existe.", vbCritical
        Exit Sub
    End If
    
    Set Carpeta = FSO.GetFolder(CarpetaBase)
    
    ' Iterar subcarpetas de primer nivel
    For Each subcarpeta In Carpeta.SubFolders
        ' Saltar si es carpeta oculta o se llama .git
        If LCase(subcarpeta.Name) = ".git" Then GoTo SiguienteSubcarpeta
        If (subcarpeta.Attributes And 2) <> 0 Then GoTo SiguienteSubcarpeta ' 2 = Hidden
        
        ' Verificar si ya existe 1.jpg
        If FSO.FileExists(subcarpeta.Path & "\1.jpg") Then GoTo SiguienteSubcarpeta
        
        ' Buscar primera imagen en la subcarpeta
        encontrado = False
        For Each archivo In subcarpeta.Files
            If Not (archivo.Attributes And 2) = 0 Then GoTo SiguienteArchivo ' Saltar archivos ocultos
            If EsImagen(archivo.Name, extensionesImagenes) Then
                primerImagen = archivo.Path
                FSO.CopyFile archivo.Path, subcarpeta.Path & "\1.jpg"
                encontrado = True
                Exit For
            End If
SiguienteArchivo:
        Next archivo
        
SiguienteSubcarpeta:
    Debug.Print "Analizando la carpeta: " & subcarpeta
    Next subcarpeta
    
    MsgBox "Proceso completado.", vbInformation
End Sub

Function EsImagen(nombreArchivo As String, extensiones As Variant) As Boolean
    Dim ext As Variant
    For Each ext In extensiones
        If LCase(Right(nombreArchivo, Len(ext) + 1)) = "." & LCase(ext) Then
            EsImagen = True
            Exit Function
        End If
    Next ext
    EsImagen = False
End Function


Function Control()
    ActiveSheet.Cells(3, 6).Activate
    
    Do While ActiveCell <> ""
        If ActiveCell.Offset(-1, 0).Value = ActiveCell.Value Then
            ' hacer algo
        Else
            ActiveCell.Offset(0, 5).Value = "Sin repetir"
        End If
        ActiveCell.Offset(1, 0).Activate
    Loop
End Function
