# importar-varias-columnas-de-un-excel-a-otro-excel
Sub ImportarArchivo()

    Dim LibroDestino As Workbook
    Dim LibroOrigen As Workbook
    Dim Ruta As String
    Dim Uf As Long
    
    Set LibroDestino = ThisWorkbook
    
    Ruta = Application.GetOpenFilename(Title:="Por favor seleccione un libro") 'seleccion de ruta del archivo de excel
    
    If Ruta = "False" Then
   '   MsgBox ("Este es el mensaje")
     
        Exit Sub
    End If
    
    Set LibroOrigen = Workbooks.Open(Ruta)
    
 ' el significado de Sheets(2) es la segunda hoja del excel que importamos que empieza desde la letra C
 
    Uf = LibroOrigen.Sheets(2).Range("C" & Rows.Count).End(xlUp).Row
    
    ' Aqui el comando ("E2:E" & Uf) define el rango de incio como columna inicial y fonal donde copia y pega en la hoja destino siendo el>>
    'Sheets(1).Range("A7") la primera hoja y pega desde el A7 para este caso
    
    LibroOrigen.Sheets(2).Range("E2:E" & Uf).Copy Destination:=LibroDestino.Sheets(1).Range("A7")
    LibroOrigen.Sheets(2).Range("M2:M" & Uf).Copy Destination:=LibroDestino.Sheets(1).Range("B7")
    LibroOrigen.Sheets(2).Range("I2:I" & Uf).Copy Destination:=LibroDestino.Sheets(1).Range("D7")
     
    
    LibroOrigen.Close
    


End Sub
