Attribute VB_Name = "Module2"
Function GetYearPosition(ByVal searchYear As Integer) As Long
    Dim ws As Worksheet
    Dim yearRange As Range
    Dim yearCell As Range
    Dim uniqueYears As Collection
    Dim i As Long
    
    ' Establecer la hoja de trabajo "Archivos"
    Set ws = ThisWorkbook.Sheets("Archivos")
    
    ' Rango de los años (columna A desde la fila 2 hasta la 50)
    Set yearRange = ws.Range("A2:A1000")
    
    ' Crear una colección para almacenar los años únicos
    Set uniqueYears = New Collection
    
    On Error Resume Next
    ' Llenar la colección con años únicos
    For Each yearCell In yearRange
        If Not IsEmpty(yearCell.value) Then
            uniqueYears.Add yearCell.value, CStr(yearCell.value)
        End If
    Next yearCell
    On Error GoTo 0
  
    ' Buscar la posición del año
    For i = 1 To uniqueYears.Count
      
        If uniqueYears(i) = searchYear Then
           
            
                GetYearPosition = i - 1 ' Restar 1 para que la posición comience desde 0
                Exit Function
            
          
        End If
    Next i
    
    ' Si no se encuentra el año, devolver -1
    GetYearPosition = -1
End Function
' === Helpers: detectar último mes con CSV en "Archivos" (cols D:E) ===
Private Function UltimoMesCSV(ByVal year As Long) As Integer
    Dim ws As Worksheet, lastRow As Long, r As Long, c As Long
    Dim s As String, mm As Integer, maxm As Integer
    Set ws = ThisWorkbook.Sheets("Archivos")
    maxm = 0
    For c = 4 To 5 ' D y E
        lastRow = ws.Cells(ws.Rows.Count, c).End(xlUp).Row
        For r = 1 To lastRow
            s = Trim$(CStr(ws.Cells(r, c).value))
            ' formato esperado: "MM.yyyy.csv"
            If Len(s) >= 11 And Right$(s, 4) = ".csv" Then
                If Mid$(s, 4, 4) = CStr(year) Then
                    mm = Val(Left$(s, 2))
                    If mm >= 1 And mm <= 12 Then
                        If mm > maxm Then maxm = mm
                    End If
                End If
            End If
        Next r
    Next c
    UltimoMesCSV = maxm ' 0 si no hay ninguno
End Function

' === Helper: ubicar inicio/fin del bloque de un año en F29 (columna A) ===
Private Sub LimitesBloqueF29(ByVal year As Long, ByRef startRow As Long, ByRef endRow As Long)
    Dim ws As Worksheet, lastRow As Long, r As Long, found As Boolean
    Set ws = ThisWorkbook.Sheets("F29")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    startRow = 0: endRow = 0

    ' buscar el rótulo del año en col A
    For r = 1 To lastRow
        If ws.Cells(r, "A").value = year Then
            startRow = r               ' fila del rótulo del año (encabezado)
            Exit For
        End If
    Next r

    If startRow = 0 Then Exit Sub

    ' fin del bloque: siguiente rótulo de año o fin de hoja
    For r = startRow + 1 To lastRow
        If IsNumeric(ws.Cells(r, "A").value) And ws.Cells(r, "A").value >= 1900 And ws.Cells(r, "A").value <= 9999 Then
            endRow = r - 1
            Exit For
        End If
    Next r
    If endRow = 0 Then endRow = lastRow
End Sub

' === Helper: ¿la fila es de “código” (020, 142, etc.)? -> B numérica ===
Private Function FilaConCodigo(ws As Worksheet, ByVal r As Long) As Boolean
    FilaConCodigo = IsNumeric(ws.Cells(r, "B").value)
End Function

' === Acción: limpia meses futuros (sin CSV) dentro del bloque del año en F29 ===
Public Sub CorrigeMesesFuturosF29(ByVal year As Long)
    Dim ws As Worksheet
    Dim startRow As Long, endRow As Long
    Dim lastMonth As Integer, clearFromCol As Integer, r As Long, c As Long

    Set ws = ThisWorkbook.Sheets("F29")

    ' 1) ¿hasta qué mes hay CSV?
    lastMonth = UltimoMesCSV(year)        ' 0..12
    If lastMonth = 12 Then Exit Sub       ' todo completo -> nada que limpiar

    ' 2) Límites del bloque del año en F29
    Call LimitesBloqueF29(year, startRow, endRow)
    If startRow = 0 Or endRow = 0 Then Exit Sub

    ' 3) Limpiar desde el mes siguiente al último disponible
    '    C=3 => mes 1, ..., N=14 => mes 12
    clearFromCol = 3 + lastMonth          ' si lastMonth=5 (mayo), se limpia desde col 8 (junio)
    If clearFromCol <= 14 Then
        Application.ScreenUpdating = False
        For r = startRow + 1 To endRow     ' saltar la fila del rótulo del año
            If FilaConCodigo(ws, r) Then
                For c = clearFromCol To 14
                    ws.Cells(r, c).ClearContents
                Next c
            End If
        Next r
        Application.ScreenUpdating = True
    End If

    ' 4) Recalcular totales en col O (si ya tienes este sub en tu módulo)
    On Error Resume Next
    RecalcTotalesF29
    On Error GoTo 0
End Sub

Sub Boton_F29_fom()
    
    Dim ws As Worksheet
    Dim startYear As Long
    Dim endYear As Long
    Dim yearDiff As Long
    
    Dim i As Long
    Dim yearInt As Long
    Dim CountYear As Integer
    Dim wsDest As Worksheet
    
    
    Call LimpiarValores(3)
    Call LimpiaFormatea(3)
    
    ' Definir la hoja de trabajo
    Set ws = ThisWorkbook.Sheets("F29")
    
    ' Leer los valores de los años
    startYear = ws.Range("G2").value
    endYear = ws.Range("I2").value
    
    ' Calcular la diferencia de años (número de replicaciones)
     yearDiff = endYear - startYear
    
     
     ' Verificar que la diferencia de años sea mayor que 0
     If yearDiff < 0 Then
         MsgBox "El valor de año final debe ser mayor o igual que el de inicio.", vbExclamation
         Exit Sub
     End If
    
    If Not Validaryear(startYear, "C") Or Not Validaryear(endYear, "C") Then
    
        MsgBox "Periodo Anual fuera de rango"
    Else
      
    
       Set wsDest = ThisWorkbook.Sheets("Ventas")
       wsDest.Cells(1, 7).value = startYear
       wsDest.Cells(1, 9).value = endYear
       
       
       Set wsDest = ThisWorkbook.Sheets("Compras")
       wsDest.Cells(1, 7).value = startYear
       wsDest.Cells(1, 9).value = endYear
       
       Call ReplicarPlanillaCalculo(startYear, 3, "F29", yearDiff)
       
       
       positionRow = 6 ' Inicializamos la posición de la fila en 3 para el primer año
       
           
       CountYear = 1
       For i = 0 To yearDiff
           yearInt = startYear + i
       
           Call ProcesarF29(yearInt, positionRow, CountYear)
           ' Después de cada año, incrementamos positionRow en 19 (para el siguiente año)
           positionRow = positionRow + 60
           CountYear = CountYear + 13
           
       Next i
    End If
    
    ' <<< limpia meses sin CSV en el ÚLTIMO año (p.ej. 2025)
    Call CorrigeMesesFuturosF29(endYear)
 
End Sub
Function Validaryear(ByVal yearBuscado As Long, ByVal Col As String) As Boolean
    Dim ws As Worksheet
    Dim rngYears As Range
    Dim celda As Range
    Dim tieneDatos As Boolean
    
    ' Establecer la hoja "Archivos"
    Set ws = ThisWorkbook.Sheets("Archivos")
    
    ' Definir el rango de años (columna A hasta fila 100, ajusta según tus datos)
    Set rngYears = ws.Range("A2:A100") ' Comienza en fila 2 para evitar encabezados
    
    ' Inicializar la variable de retorno como Falso
    Validaryear = False
    
    ' Buscar el año en la columna A
    For Each celda In rngYears
        If celda.value = yearBuscado Then
            ' Verificar si hay datos en la columna especificada para la fila correspondiente
            If ws.Cells(celda.Row, columns(Col).Column).value <> "" Then
                Validaryear = True ' Si se encuentra al menos un dato, es válido
                Exit Function       ' Salir de la función
            End If
        End If
    Next celda
End Function



Private Sub ProcesarF29(ByVal yearInput As Long, ByVal positionRow As Integer, ByVal CountY As Integer)

   ' Definir variables
    Dim F29Path As String
  
    Dim wsFiles As Worksheet
    Dim wsDataset As Worksheet
    Dim mes As Integer
    Dim posArchivo As Integer
    Dim posicion As Long
    Dim Dataset As String
    posicion = GetYearPosition(yearInput)
    If posicion > 0 Then
        posicion = posicion * 13
           
    End If
    posicion = posicion + 1
    
    ' Obtener valores de los parámetros de la hoja Param
    F29Path = Sheets("Param").Range("B2").value
    ' Concatenar el año a la ruta
    F29Path = F29Path & "\" & CStr(yearInput)
     
   
    Set wsFiles = Sheets("Archivos")
    Set wsDataset = Sheets("dataset")
    Dim archivoPDF As String
    Dim archivoTXT As String
    posArchivo = CountY
    ' Leer 12 archivos y procesarlos
    posArchivo = posicion
     
    
    For mes = 1 To 12
        ' Construir la ruta y el nombre del archivo PDF
        
        archivoPDF = F29Path & "\" & wsFiles.Cells(posArchivo + mes, 3).value
        
        archivoTXT = wsFiles.Cells(posArchivo + mes, 3).value
       
        If wsFiles.Cells(posArchivo + mes, 3).value <> "" Then
            Dataset = wsDataset.Cells(posArchivo + mes, 2).value
            'MsgBox Dataset
            If Dataset <> "" Then
                Call BuscarCodigos(mes, positionRow, Dataset, yearInput)
            Else
                archivoTXT = archivoTXT & ": Sin datos"
                MsgBox archivoTXT
            End If
            
        End If
        
    Next mes
   

End Sub

Sub BuscarCodigos(mes As Integer, ByVal positionRow As Integer, Dataset As String, ByVal yearInput As Long)
'Sub BuscarCodigos(mes As Integer, archivoPDF As String, ByVal positionRow As Integer, Dataset As String)
    ' Definir variables
    Dim i As Integer
    Dim wsSummary As Worksheet
    Dim codigo As String
    Dim desc As String
    Dim resultado As String
    Dim LineData As String
    Dim Records As Integer
    Dim pos As Integer
    Dim Lista As Collection
    Set Lista = New Collection ' Inicializa la colección
    Dim Result As Variant
    
     
    
    Set Lista = GetDataPdf(Dataset, yearInput)
   
    Records = 90
    
    ' Obtener referencia a la hoja Summary
    Set wsSummary = ThisWorkbook.Sheets("F29")
    
    pos = positionRow
    
    
    For i = pos To Records + pos
        codigo = wsSummary.Cells(pos, 2).value
        desc = UCase(wsSummary.Cells(pos, 1).value)
         
        If codigo <> "" And desc <> "" And IsNumeric(codigo) Then
                            
                'Result = ObtenerValorPorCodigo(Lista, codigo)
                Result = ObtenerValorPorCodigoBinario(Lista, codigo)
             
                
                If Result = -1 Then
                    wsSummary.Cells(pos, mes + 2).value = ""
                Else
                    wsSummary.Cells(pos, mes + 2).value = Result
                End If
           
        
        End If
        pos = pos + 1
    Next i

End Sub
Function ObtenerValorPorCodigoBinario(ByRef Lista As Collection, ByVal CodigoBuscado As String) As Variant
    Dim inicio As Long
    Dim fin As Long
    Dim medio As Long
    Dim CodigoActual As Variant

    inicio = 1
    fin = Lista.Count
    
    ' Convertir CódigoBuscado a número si es posible
    On Error Resume Next
    
    On Error GoTo 0

    While inicio <= fin
        medio = (inicio + fin) \ 2
        CodigoActual = Lista(medio)(1) ' Suponemos que el código está en el índice 1 del elemento

        ' Comparar el CódigoBuscado con el CódigoActual
        If CodigoActual = CodigoBuscado Or CInt(CodigoActual) = CInt(CodigoBuscado) Then
            ObtenerValorPorCodigoBinario = Lista(medio)(3) ' Retorna el ValorActual
            Exit Function
        ElseIf CInt(CodigoBuscado) < CInt(CodigoActual) Then
            fin = medio - 1
        Else
            inicio = medio + 1
        End If
    Wend

    ' Si no se encuentra el código, devuelve -1
    ObtenerValorPorCodigoBinario = -1
End Function


Function GetDataPdf(Dataset As String, ByVal yearInput As Long) As Collection
 
    Dim texto As String
    Dim Bloques() As String
    Dim Lineas() As String
    Dim codigo As Variant
    Dim Descripcion As String
    Dim Valor As String
    Dim resultado As String
    Dim CodigosValidos As Object
    Dim i As Integer, j As Integer
    Dim Bloque As String
    Dim PosTotal As Long
    Dim miLista As Collection
    Dim posicion As Long
    
    Dim ws As Worksheet
    
    Set miLista = New Collection ' Inicializas la colección
  
    texto = Dataset
  
    If posicion > 0 Then
        ' Cortar el texto desde la primera aparición de "Código Glosa Valor"
        texto = Mid(texto, posicion)
    End If
    ' Encontrar la posición de la primera aparición de "Código Glosa Valor"
    posicion = InStr(texto, "Código Glosa Valor")
    
    texto = Replace(texto, "Código Glosa Valor", "|")
    
    texto = Replace(texto, "TOTAL A PAGAR DENTRO DEL PLAZO LEGAL", "|")
    texto = Replace(texto, "+", "|")
    
     
    ' Cortar el texto antes de "TOTAL A PAGAR"
    'PosTotal = InStr(1, texto, "TOTAL A PAGAR")
    'If PosTotal > 0 Then
    '    texto = Left(texto, PosTotal - 1) ' Corta antes de "TOTAL"
    'End If
    
    ' Separar el texto en bloques usando los códigos como delimitadores
    Dim Flag As Boolean
    
     
    Bloques = Split(texto, "|")
    'MsgBox texto
    ' Procesar cada bloque
    resultado = ""
  
    For i = LBound(Bloques) To UBound(Bloques)
        Bloque = Trim(Bloques(i))
        If Bloque <> "" Then
            ' Dividir el bloque en palabras
            Lineas = Split(Bloque, " ")
            
            ' Extraer código, descripción y valor
         
            If BuscarEnArray(Lineas(0), yearInput) Then
                
                codigo = Lineas(0)
                Valor = Lineas(UBound(Lineas))
                Valor = Trim(Lineas(UBound(Lineas)))
                Valor = FormatearNumero(Valor)
                
               
                ' Construir la descripción manualmente
                Descripcion = ""
                If UBound(Lineas) > 1 Then
                    For j = 1 To UBound(Lineas) - 1
                        Descripcion = Descripcion & Lineas(j) & " "
                    Next j
                    Descripcion = Trim(Descripcion)
                    
                End If
                Descripcion = "    " & Descripcion & "    "
                If CInt(Lineas(0)) = 91 Then
                    miLista.Add Array(miLista, codigo, Descripcion, Lineas(1))
                    'MsgBox Lineas(1)
                Else
                    miLista.Add Array(miLista, codigo, Descripcion, Valor)
                End If

                ' Agregar el resultado formateado
               
                
               
                
                'Call ImprimirListaEnHoja(miLista)
               
                 
            End If
        End If
    Next i
   
   
   Set miLista = OrdenarListaConWorksheetFunction(miLista)
   'MsgBox ObtenerValorPorCodigo(miLista, 520)
   Set GetDataPdf = miLista
End Function
Function OrdenarListaConWorksheetFunction(miLista As Collection) As Collection
    Dim ws As Worksheet
    Dim miLista2 As Collection
    Dim i As Long
    
    ' Crear la lista como una colección
    Set miLista2 = New Collection
     
   
     ' Verificar si la hoja "Sheet1" existe
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Temp")
    On Error GoTo 0
    
    ' Si la hoja no existe, la creamos
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = "Temp"
    End If
    
    ' Ocultar la hoja
    ws.Visible = xlSheetHidden ' Oculta la hoja (modo normal)
    ' Si deseas ocultarla completamente, usa: ws.Visible = xlSheetVeryHidden
    
    ' Pasar los datos de la lista a la hoja de cálculo
    For i = 1 To miLista.Count
        ws.Cells(i, 1).value = miLista(i)(1) ' Código
        ws.Cells(i, 2).value = miLista(i)(2) ' Descripción
        ws.Cells(i, 3).value = miLista(i)(3) ' Valor
    Next i
    
    ' Ordenar los datos en la hoja de cálculo por la columna "Código"
    ws.Sort.SortFields.Clear
    ws.Sort.SortFields.Add Key:=ws.Range("A1:A" & miLista.Count), _
                           SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ws.Sort
        .SetRange ws.Range("A1:C" & miLista.Count)
        .header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ' Recuperar los datos ordenados de la hoja de cálculo
    ' miLista.Clear
    For i = 1 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        miLista2.Add Array(miLista2, ws.Cells(i, 1).value, ws.Cells(i, 2).value, ws.Cells(i, 3).value)
       
    Next i
    
     
    ' Limpiar solo los contenidos de la hoja
    ws.Cells.ClearContents
    
    Set OrdenarListaConWorksheetFunction = miLista2
    
End Function

 
Function FormatearNumero(ByVal Valor As String) As String
    Dim NumeroFormateado As String
    
    On Error GoTo ManejoErrores

    ' Validar si contiene decimales (coma)
    If InStr(Valor, ",") > 0 Or InStr(Valor, ".") Then
        ' Reemplazar separadores de miles (comas) por puntos
        Valor = Replace(Valor, ".", "") ' Eliminar puntos, suponiendo que son separadores de miles
        Valor = Replace(Valor, ",", "") ' Convertir la coma decimal en punto decimal
    End If
     
    
    FormatearNumero = Valor
    Exit Function

ManejoErrores:
    FormatearNumero = "Error: Ocurrió un problema al procesar el valor"
End Function


Sub ImprimirListaEnHoja(ByRef Lista As Collection)
    ' Imprime los datos de la colección en una hoja de Excel
    Dim ws As Worksheet
    Dim i As Long
    Dim j As Long
    Dim Col As Long
    Dim elemento As Variant
    
    ' Crear o seleccionar la hoja
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Listado")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = "Listado"
    End If
    On Error GoTo 0
    
    ' Limpiar la hoja
    'ws.Cells.Clear
    ws.Cells(1, 1).value = "Código"
    ws.Cells(1, 2).value = "Descripción"
    ws.Cells(1, 3).value = "Valor"
    
    ' Escribir los datos en la hoja
    i = 2
    
    
    For Each elemento In Lista
        ws.Cells(i, Col + 1).value = elemento(1) ' Código
        ws.Cells(i, Col + 2).value = elemento(2) ' Descripción
        ws.Cells(i, Col + 3).value = elemento(3) ' Valor
        i = i + 1
    Next elemento
      
     
    
    'MsgBox "Lista impresa en la hoja 'Listado'.", vbInformation
End Sub
Function ProcesaDataSecuencial(rutaPDF As String, ByVal yearInput As Long) As String 'Obtiene el texto completo desde el PDF F29
   
    Dim AcroApp As Object
    Dim AcroAVDoc As Object
    Dim AcroPDDoc As Object
    Dim jsObj As Object
    Dim NumPaginas As Integer
    Dim i As Integer, j As Integer
    Dim NumPalabras As Integer
    Dim TextoExtraido As String
    Dim ruta As String
    Dim palabra As String
    Dim LastPalabra As String
    Dim Flag As Boolean
    Dim Cadena As String
    ruta = ""
    ruta = rutaPDF
    ' Crear objeto para Adobe Acrobat
    
    'On Error Resume Next
    Set AcroApp = CreateObject("AcroExch.App")
    Set AcroAVDoc = CreateObject("AcroExch.AVDoc")
    'On Error GoTo 0
    
    'If AcroApp Is Nothing Or AcroAVDoc Is Nothing Then
    '    MsgBox "Adobe Acrobat Reader no está configurado correctamente o no está instalado.", vbCritical
        
    'End If
    LastPalabra = ""
    
    ' Abrir el PDF
    If AcroAVDoc.Open(ruta, "") Then
        Set AcroPDDoc = AcroAVDoc.GetPDDoc
        Set jsObj = AcroPDDoc.GetJSObject

        ' Obtener el número total de páginas
        NumPaginas = jsObj.numPages

        ' Extraer texto de cada página
        
        For i = 0 To NumPaginas - 1
            NumPalabras = jsObj.getPageNumWords(i)
            Cadena = ""
            For j = 0 To NumPalabras - 1
                Flag = True
                           
                If Not IsNumeric(LastPalabra) Then
                    Flag = False
                    
                End If
            
                ' Obtener la palabra de la página actual
                palabra = jsObj.getPageNthWord(i, j)
                palabra = FormatearNumero(palabra)
                LastPalabra = palabra
                
               
                'If ObtenerPenultimaPalabra(Cadena) = "Ley" And Flag And IsNumeric(Palabra) Then
                '    MsgBox Cadena
                'End If
                
                    ' Validar si la palabra está en CodigosValidos.Keys y tiene 3 o menos caracteres
                'If ValidarCodigo(palabra) And Len(palabra) <= 3 And Flag And ObtenerPenultimaPalabra(Cadena) <> "Ley" Then
                If BuscarEnArray(palabra, yearInput) And Len(palabra) <= 3 And Flag Then
                    ' Si la palabra es un código válido y tiene 3 o menos caracteres, agregar "|" al inicio de la palabra
                     
                    palabra = "|" & palabra
                    Cadena = ""
                End If
                Cadena = Cadena & LastPalabra & " "
                TextoExtraido = TextoExtraido & palabra & " "
            Next j
            TextoExtraido = TextoExtraido & vbCrLf ' Salto de línea entre páginas
        Next i
               
        ProcesaDataSecuencial = TextoExtraido ' Retorna el dataset
         
        ' Cerrar el documento
        AcroAVDoc.Close True
    Else
        MsgBox "No se pudo abrir el archivo PDF.", vbExclamation
    End If

    ' Cerrar objetos
    AcroApp.Exit
    Set AcroApp = Nothing
    Set AcroAVDoc = Nothing
    Set AcroPDDoc = Nothing
    Set jsObj = Nothing
End Function

Function BuscarEnArray(codigo As String, Optional yearInput As Long = 0) As Boolean
    Dim arr As Variant
    Dim i As Long
    BuscarEnArray = False

    If Len(codigo) <= 3 Then
        If codigo = "039" Or codigo = "39" Then
            If yearInput >= 2023 Then
                BuscarEnArray = True
            End If
            Exit Function
        End If
    End If
    
    ' Definimos el array con datos de prueba
    arr = Array("010", "020", "028", "030", "048", "049", "050", "054", "056", "062", "066", "068", "077", "089", "091", "91", "110", "111", _
"113", "115", "120", "122", "123", "127", "142", "151", "152", "153", "154", "155", "156", "157", "164", "409", "500", "501", "502", "503", _
"504", "509", "510", "511", "512", "513", "514", "515", "516", "517", "518", "519", "520", "521", "522", "523", "524", "525", "526", "527", _
"528", "529", "530", "531", "532", "534", "535", "536", "537", "538", "539", "540", "541", "542", "543", "544", "547", "548", "550", "553", _
"557", "560", "562", "563", "564", "565", "566", "573", "584", "585", "586", "587", "588", "589", "592", "593", "594", "595", "700", "701", _
"702", "703", "708", "709", "711", "712", "713", "714", "715", "716", "717", "718", "720", "721", "722", "723", "724", "729", "730", "731", _
"732", "734", "735", "738", "739", "740", "741", "742", "743", "744", "745", "749", "750", "751", "755", "756", "757", "758", "759", "761", _
"762", "763", "764", "772", "773", "774", "775", "776", "777", "778", "779", "780", "782", "783", "784", "785", "786", "787", "788", "789", _
"791", "792", "793", "794", "796", "797", "798", "799", "800", "801", "802", "803", "804", "805", "806", "807", "808", "809", "810")
    
    ' Recorremos el array
    For i = LBound(arr) To UBound(arr)
        If arr(i) = codigo And Len(codigo) <= 3 Then
            BuscarEnArray = True
            Exit Function ' Salimos de la función si encontramos el valor
        End If
    Next i
    
    
End Function
 


Sub ImprimirListaEnHoja2(Data As String)
    ' Imprime los datos de la colección en una hoja de Excel
    Dim ws As Worksheet
    Dim i As Long
    Dim elemento As Variant
    
    
    
    ' Crear o seleccionar la hoja
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Listado2")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = "Listado2"
    End If
    On Error GoTo 0
    
    ' Limpiar la hoja
    'ws.Cells.Clear
   
    ' Escribir los datos en la hoja
    i = 1
    Do While i < 12
        ' Tarea a realizar en cada fila
        If ws.Cells(i, 1).value = "" Then
            ws.Cells(i, 1).value = Data
            Exit Do
        End If
        
        ' Incrementar el contador
        i = i + 1
    Loop
     
End Sub

Sub boton01()
    
    Dim ws As Worksheet
    Dim startYear As Long
    Dim endYear As Long
    Dim yearDiff As Long
    Dim yearInt As Long
    
    Dim i As Long
    Dim CountYear As Integer
    
    Call LimpiarValores(1)
    Call LimpiaFormatea(1)
    
    ' Definir la hoja de trabajo
    Set ws = ThisWorkbook.Sheets("Compras")
    
    ' Leer los valores de los años desde las celdas G1 (inicio) y G2 (fin)
    startYear = ws.Range("G1").value
    endYear = ws.Range("I1").value
    
    ' Calcular la diferencia de años (número de replicaciones)
    yearDiff = endYear - startYear
    
    ' Verificar que la diferencia de años sea mayor que 0
    If yearDiff < 0 Then
        MsgBox "El valor de año final debe ser mayor o igual que el de inicio.", vbExclamation
        Exit Sub
    End If
    
    If Not Validaryear(startYear, "D") Or Not Validaryear(endYear, "D") Then
    
        MsgBox "Periodo Anual fuera de rango"
    Else
       Call ReplicarPlanillaCalculo(startYear, 1, "Compras", yearDiff)
       
       
       ' Llamar a la función ProcesarRC para cada año
       
       positionRow = 5 ' Inicializamos la posición de la fila en 5 para el primer año
        
       CountYear = 1
       For i = 0 To yearDiff
           yearInt = startYear + i
        
           Call ProcesarRC(yearInt, positionRow, CountYear)
           ' Después de cada año, incrementamos positionRow en 19 (para el siguiente año)
           positionRow = positionRow + 19
           CountYear = CountYear + 13
           If i > 0 Then
              ' MsgBox i
               Call ReplicarDynamicFormula("Compras", i, 1)
           End If
       Next i
    End If
End Sub
Private Sub ProcesarRC(ByVal yearInput As Long, ByVal positionRow As Long, ByVal CountY As Integer)

    ' Definir variables
    Dim RC_Path As String
    'Dim RC_Fname As String
    Dim filepath As String
    Dim wsSummary As Worksheet
    Dim wsFiles As Worksheet
    Dim mes As Integer
    Dim arr() As Variant
    Dim pos As Integer
    Dim posArchivo As Integer
    
    Dim posicion As Integer
    posicion = 0
    
    posicion = GetYearPosition(yearInput)
    
    If posicion > 0 Then
        posicion = posicion * 13
         
    End If
    posicion = posicion + 1
  

     ' Obtener valores de los parámetros de la hoja Param
    RC_Path = Sheets("Param").Range("B4").value
    ' Concatenar el año a la ruta
    RC_Path = RC_Path & "\" & yearInput ' Concatenar año al final de la ruta
    'RC_Fname = Sheets("Param").Range("B5").value
    
    ' Obtener referencia a la hoja Summary
    Set wsSummary = ThisWorkbook.Sheets("Compras")
    Set wsFiles = Sheets("Archivos")
    
    ' Iterar 12 archivos y procesarlos , Tipo Doc : 33,34,46,56,61,110
    pos = positionRow
    posArchivo = posicion
    For mes = 1 To 12
        
        ' Procesar el archivo TXT y buscar los códigos
        filepath = RC_Path & "\" & wsFiles.Cells(posArchivo + mes, 4).value
        'MsgBox wsFiles.Cells(posArchivo + mes, 4).value
        If wsFiles.Cells(posArchivo + mes, 4).value <> "" Then
            'Notas crédito
            arr = Array("61")
            wsSummary.Cells(pos + mes, 4) = Abs(SumaPorColumnaParametrica(filepath, mes, "Tipo Doc", arr, "Monto IVA Recuperable", 12))
            
            
            'Total Crédito fiscal
            arr = Array("33", "39", "46", "56", "914")
            wsSummary.Cells(pos + mes, 9) = SumaPorColumnaParametrica(filepath, mes, "Tipo Doc", arr, "Monto IVA Recuperable", 12) - wsSummary.Cells(pos + mes, 4)
            
            'Exento
            arr = Array("33", "34", "39", "46", "56", "914")
            wsSummary.Cells(pos + mes, 12) = SumaPorColumnaParametrica(filepath, mes, "Tipo Doc", arr, "Monto Exento", 10)
            arr = Array("61")
            wsSummary.Cells(pos + mes, 12) = wsSummary.Cells(pos + mes, 12) + SumaPorColumnaParametrica(filepath, mes, "Tipo Doc", arr, "Monto Exento", 10)
        End If
        
    Next mes

End Sub
 
Sub Buton02()
    
    Dim ws As Worksheet
    Dim yearInt As Long
    Dim yearDiff As Long
    
    Dim startYear As Long
    Dim endYear As Long
    
    Dim i As Long
    
    Dim CountYear As Integer
  
    Call LimpiarValores(2)
    Call LimpiaFormatea(2)
    
    ' Definir la hoja de trabajo
    Set ws = ThisWorkbook.Sheets("Ventas")
    
    ' Leer los valores de los años desde las celdas G1 (inicio) y G2 (fin)
    startYear = ws.Range("G1").value
    endYear = ws.Range("I1").value
    
    ' Calcular la diferencia de años (número de replicaciones)
    yearDiff = endYear - startYear
    
    ' Verificar que la diferencia de años sea mayor que 0
    If yearDiff < 0 Then
        MsgBox "El valor de año final debe ser mayor o igual que el de inicio.", vbExclamation
        Exit Sub
    End If
    
    If Not Validaryear(startYear, "E") Or Not Validaryear(endYear, "E") Then
        
            MsgBox "Periodo Anual fuera de rango"
    Else
        
        Call ReplicarPlanillaCalculo(startYear, 2, "Ventas", yearDiff)
          
        
        positionRow = 4 ' Inicializamos la posición de la fila en 4 para el primer año
         
        CountYear = 1
        For i = 0 To yearDiff
            yearInt = startYear + i
            ' Llamamos a la función ProcesarRV, pasando tanto el año como la posición de la fila
             
            Call ProcesarRV(yearInt, positionRow, CountYear)
            ' Después de cada año, incrementamos positionRow en 19 (para el siguiente año)
            positionRow = positionRow + 19
            CountYear = CountYear + 13
            
            If i > 0 Then
               ' MsgBox i
                Call ReplicarDynamicFormula("Ventas", i, 2)
            End If
        Next i
        
        'Call ReplicarDynamicFormula("Ventas", 24, 1, 24)
    End If
End Sub

' Pone en negrita los totales de O cuando la fila tiene código numérico en B
Sub FormatTotalesF29()
    Dim ws As Worksheet, lastRow As Long, r As Long, hasCode As Boolean
    Set ws = ThisWorkbook.Sheets("F29")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    For r = 6 To lastRow
        hasCode = (Len(ws.Cells(r, "B").value) > 0 And IsNumeric(ws.Cells(r, "B").value))
        If hasCode Then
            With ws.Cells(r, "O")
                .Font.Bold = True
                .NumberFormat = "#,##0"
            End With
        Else
            ws.Cells(r, "O").Font.Bold = False
        End If
    Next r
End Sub

' --- Rellena O con la suma de C:N cuando hay código numérico en B ---
Sub RecalcTotalesF29()
    Dim ws As Worksheet
    Dim lastRow As Long, r As Long
    Dim hasCode As Boolean

    Set ws = ThisWorkbook.Sheets("F29")
    Application.ScreenUpdating = False

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    For r = 6 To lastRow
        hasCode = False
        If Len(ws.Cells(r, "B").value) > 0 Then
            hasCode = IsNumeric(ws.Cells(r, "B").value)
        End If

        If hasCode Then
            ws.Cells(r, "O").FormulaR1C1 = "=SUM(RC[-12]:RC[-1])"
        Else
            ws.Cells(r, "O").ClearContents
        End If
    Next r

    Application.ScreenUpdating = True
End Sub

' ---------- NUEVO: helper para limpiar (dejar en blanco) filas fijas ----------
Private Sub LimpiarFilasPuente(ByVal hoja As String, ParamArray filas())
    Dim ws As Worksheet, k As Long, f As Long
    Set ws = ThisWorkbook.Sheets(hoja)
    For k = LBound(filas) To UBound(filas)
        f = CLng(filas(k))
        ws.Rows(f).ClearContents
    Next k
End Sub
' ------------------------------------------------------------------------------

Sub ReplicarPlanillaCalculo(startYear As Long, NumeroHoja As Integer, NombreHoja As String, yearDiff As Long)
    Dim ws As Worksheet
    Dim sourceRange As Range
    Dim destinationRange As Range
    Dim offsetRows As Long
    Dim i As Long
    Dim yearInt As Long
    Dim inicio As String
    Dim destino As String
    Dim Rango As String
    Dim NumberoffsetRows As Integer
    Dim NumberoffsetYear As Integer

    NumberoffsetYear = 2
    Set ws = ThisWorkbook.Sheets(NombreHoja)

    Select Case NumeroHoja
        Case 1
            inicio = "B5": destino = "B3": Rango = "B3:M18": NumberoffsetRows = 3
        Case 2
            inicio = "B4": destino = "B2": Rango = "B2:M17": NumberoffsetRows = 3
        Case 3
            inicio = "A4": destino = "A4": Rango = "A4:N61": NumberoffsetRows = 2: NumberoffsetYear = 0
        Case Else
            MsgBox "Número de hoja no válido.", vbExclamation
            Exit Sub
    End Select

    Set sourceRange = ws.Range(Rango)
    offsetRows = NumberoffsetRows
    ws.Range(inicio).value = startYear

    For i = 1 To yearDiff
        Set destinationRange = ws.Range(destino).Offset((sourceRange.Rows.Count + offsetRows) * i, 0)
        sourceRange.Copy
        destinationRange.PasteSpecial Paste:=xlPasteAll
        destinationRange.Cells(1, 1).Offset(NumberoffsetYear, 0).value = startYear + i
    Next i

    Application.CutCopyMode = False

    If NumeroHoja = 3 Then
        RecalcTotalesF29
        FormatTotalesF29
    End If

    ' ======= CORRECCIÓN SOLICITADA =======
    ' Ventas: limpiar fila 20.  Compras: limpiar filas 20 y 21.
    If NumeroHoja = 2 Then
        LimpiarFilasPuente NombreHoja, 20
    ElseIf NumeroHoja = 1 Then
        LimpiarFilasPuente NombreHoja, 20, 21
    End If
    ' =====================================
End Sub

' Salto correcto por bloque de 60 filas
Function GetAdjustedCell(columnOffset As Integer, baseRow As Long) As String
    GetAdjustedCell = "$C$" & (baseRow + (60 * columnOffset))
End Function

Sub ReplicarDynamicFormula(ByVal NombreHoja As String, ByVal columnOffset As Integer, ByVal tipo As Integer)
    Dim wsDest As Worksheet
    Dim i As Integer
    Dim formulaString01 As String
    Dim formulaString02 As String
    Dim formulaString03 As String
    Dim formulaString04 As String
    Dim formulaString05 As String
    Dim startRow As Integer
    Dim rowOffset As Integer

    Set wsDest = ThisWorkbook.Sheets(NombreHoja)

    Select Case tipo
        '========================
        ' COMPRAS
        '========================
        Case 1
            startRow = (columnOffset * 19) + 6
            rowOffset = startRow
            formulaString01 = "=DESREF('F29'!" & GetAdjustedCell(columnOffset, 38) & ";0;FILA()-" & rowOffset & ")"
            formulaString02 = "=DESREF('F29'!" & GetAdjustedCell(columnOffset, 58) & ";0;FILA()-" & rowOffset & ")"
            formulaString04 = "=DESREF('F29'!" & GetAdjustedCell(columnOffset, 44) & ";0;FILA()-" & rowOffset & ")"
            formulaString05 = "=DESREF('F29'!" & GetAdjustedCell(columnOffset, 34) & ";0;FILA()-" & rowOffset & ")"

        '========================
        ' VENTAS
        '========================
        Case 2
            startRow = (columnOffset * 19) + 5
            rowOffset = startRow
            formulaString01 = "=DESREF('F29'!" & GetAdjustedCell(columnOffset, 30) & ";0;FILA()-" & rowOffset & ")"
            formulaString02 = "=DESREF('F29'!" & GetAdjustedCell(columnOffset, 7) & ";0;FILA()-" & rowOffset & ")"
            formulaString03 = "=DESREF('F29'!" & GetAdjustedCell(columnOffset, 6) & ";0;FILA()-" & rowOffset & ")"

        Case Else
            MsgBox "Número de hoja no válido.", vbExclamation
            Exit Sub
    End Select

    For i = startRow To startRow + 11
        wsDest.Cells(i, 3).FormulaLocal = formulaString01   ' Compras: Cod 528 / Ventas: Cod 538
        wsDest.Cells(i, 6).FormulaLocal = formulaString02   ' Compras: Cod 537 / Ventas: Cod 142
        If tipo = 1 Then
            wsDest.Cells(i, 7).FormulaLocal = formulaString04 ' Compras: Cod 504
            wsDest.Cells(i, 11).FormulaLocal = formulaString05 ' Compras: Exento
        Else
            wsDest.Cells(i, 10).FormulaLocal = formulaString03 ' Ventas: Cod 715
        End If
    Next i
End Sub

Sub LimpiaFormatea(NumeroHoja As Integer)
    Dim ws As Worksheet
    Dim Rango As Range
    Dim NombreHoja As String

    Select Case NumeroHoja
        Case 1: NombreHoja = "Compras"
        Case 2: NombreHoja = "Ventas"
        Case 3: NombreHoja = "F29"
        Case Else
            MsgBox "Número de hoja no válido.", vbExclamation
            Exit Sub
    End Select

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(NombreHoja)
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "La hoja '" & NombreHoja & "' no existe.", vbExclamation
        Exit Sub
    End If

    Select Case NumeroHoja
        Case 1, 2
            Set Rango = ws.Range("B20:L60")
        Case 3
            Set Rango = ws.Range("A61:N1000")
    End Select

    If Rango.MergeCells Then Rango.UnMerge
    Rango.EntireRow.Delete
End Sub


Sub LimpiarValores(NumeroHoja As Integer)
    Dim ws As Worksheet
    Dim Rango As Range
    Dim celda As Range
    Dim NombreHoja As String
    
  

    ' Determina el nombre de la hoja basado en el número de hoja
    Select Case NumeroHoja
        Case 1
            NombreHoja = "Compras"
        Case 2
            NombreHoja = "Ventas"
        Case 3
            NombreHoja = "F29"
        Case Else
            MsgBox "Número de hoja no válido.", vbExclamation
            Exit Sub
    End Select

    ' Establece la referencia a la hoja de trabajo
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(NombreHoja)
    On Error GoTo 0

    If ws Is Nothing Then
        MsgBox "No se pudo encontrar la hoja: " & NombreHoja, vbCritical
        Exit Sub
    End If

    ' Establece el rango basado en la hoja seleccionada
    Select Case NumeroHoja
        Case 1
            Set Rango = ws.Range("C6:M17")
         
        Case 2
            Set Rango = ws.Range("C5:L17")
           
        Case 3
            Set Rango = ws.Range("C5:N60")
           
    End Select
    
    ' Limpia las celdas del rango que no contienen fórmulas
    For Each celda In Rango
        If Not celda.HasFormula Then
            celda.ClearContents ' Limpia solo celdas sin fórmula
        End If
    Next celda
End Sub

Private Sub ProcesarRV(ByVal yearInput As Long, ByVal positionRow As Long, ByVal CountY As Integer)
    
    ' Definir variables
    Dim RV_Path As String
    'Dim RV_Fname As String
    Dim filepath As String
    Dim wsSummary As Worksheet
    Dim wsFiles As Worksheet
    Dim mes As Integer
    Dim arr() As Variant
    Dim pos As Integer
    Dim posArchivo As Integer
    Dim posicion As Integer
    posicion = 0
    
    posicion = GetYearPosition(yearInput)
 
    
    
    If posicion > 0 Then
        posicion = posicion * 13
        posicon = posicion + 1
         
    End If
    posicion = posicion + 1
    'MsgBox posicion
    ' Obtener valores de los parámetros de la hoja Param
    RV_Path = Sheets("Param").Range("B6").value
    ' Concatenar el año a la ruta
    RV_Path = RV_Path & "\" & yearInput ' Concatenar año al final de la ruta
    'RV_Fname = Sheets("Param").Range("B7").value
    'MsgBox RV_Path
    
    ' Obtener referencia a la hoja Summary
    Set wsSummary = ThisWorkbook.Sheets("Ventas")
    Set wsFiles = Sheets("Archivos")
    
    ' Iterar 12 archivos y procesarlos
    pos = positionRow
    'posArchivo = CountY
    posArchivo = posicion
    'MsgBox pos
    
    For mes = 1 To 12
      
        ' Procesar el archivo TXT y buscar los códigos
        
        filepath = RV_Path & "\" & wsFiles.Cells(posArchivo + mes, 5).value
        'MsgBox wsFiles.Cells(posArchivo + mes, 5).value
        If wsFiles.Cells(posArchivo + mes, 5).value <> "" Then
        
            arr = Array("33", "39", "43", "46", "56", "61")
            wsSummary.Cells(pos + mes, 4) = SumaPorColumnaParametrica(filepath, mes, "Tipo Doc", arr, "Monto IVA", 12)
            
            ' Ventas y/o Servicios exentos o no gravados
            arr = Array("33", "34", "41", "46", "56", "61")
            wsSummary.Cells(pos + mes, 8) = SumaPorColumnaParametrica(filepath, mes, "Tipo Doc", arr, "Monto Exento", 10)
            
            arr = Array("110", "111", "112")
            wsSummary.Cells(pos + mes, 11) = SumaPorColumnaParametrica(filepath, mes, "Tipo Doc", arr, "Monto Exento", 10)
        End If
        
    Next mes
    
    

End Sub


 
Private Function SumaPorColumnaParametrica(filepath, mes, ByVal nombreColumnaB As String, ByRef valoresColumnaB() As Variant, ByVal nombreColumnaRegresar As String, ByVal numeroColumnaRegresar As Integer) As Double
    Dim fileContent As String
    Dim fileLine As String
    Dim headerRow As String
    Dim sumaColumnaRegresar As Double
    Dim numeroColumnaRevisada&
    Dim i&
    
    If Dir(filepath) = "" Then
        MsgBox "La ruta del archivo es incorrecta o el archivo no existe: validar 'ruta\YYYY\nombrearchivo.csv'", vbExclamation, "Error"
        Exit Function
    End If
    
    Open filepath For Input As #1
    fileContent = Input$(LOF(1), #1)
    Close #1
      
    
    Dim lines() As String
    lines = Split(fileContent, vbCrLf)
   
    If UBound(lines) < 1 Then
        fileContent = Trim(fileContent) ' Elimina espacios o saltos al inicio y al final
        fileContent = Replace(fileContent, vbCrLf, vbLf) ' Convertir todos los `vbCrLf` en `vbLf`
        fileContent = Replace(fileContent, vbCr, vbLf)  ' Convertir `vbCr` solitarios en `vbLf`
        lines = Split(fileContent, vbLf)
  
    End If
    ' Obtener la cabecera del archivo
    Dim Hcols() As String
    headerRow = lines(0)
    'MsgBox lines(1)
    Hcols = Split(headerRow, ";")
    
    numeroColumnaRevisada = numeroColumnaRegresar
    ' Verificar que la cabecera de Tipo Doc sea correcta
    If InStr(1, headerRow, nombreColumnaB, vbTextCompare) = 0 Or Hcols(1) <> nombreColumnaB Then
        
        Logger "E", "Archivo " & mes, "La cabecera '" & nombreColumnaB & "' no se encuentra en la fila 1 pos 2."
        Exit Function
    End If
    
    ' Verificar que la cabecera de Monto IVA Recuperable sea correcta - 202412 se ha comentado ya que no esta tolerante a falla por datos no númericos y las instrucciones
    ' se ha reemplazado con el siguiente codigo.
    If InStr(1, headerRow, nombreColumnaRegresar, vbTextCompare) = 0 Then
        Logger "E", "Archivo " & mes, "La cabecera '" & nombreColumnaRegresar & "' no se encuentra en la fila 1 pos " & numeroColumnaRegresar
        Exit Function
    
    ElseIf Hcols(numeroColumnaRegresar - 1) <> nombreColumnaRegresar Then
 
    '    numeroColumnaRevisada = Application.Match(nombreColumnaRegresar, Hcols, 0)
        Logger "W", "Archivo " & mes, "La cabecera '" & nombreColumnaRegresar & "' reposicionada a pos " & numeroColumnaRevisada
    End If
    Dim resultado As Variant

    'If Hcols(numeroColumnaRegresar - 1) <> nombreColumnaRegresar Then
    If Hcols(numeroColumnaRegresar - 1) <> nombreColumnaRegresar Then
        
        ' Intenta encontrar la coincidencia
        On Error Resume Next
        resultado = Application.Match(nombreColumnaRegresar, Hcols, 0)
        On Error GoTo 0
    
        ' Verifica si la búsqueda fue exitosa
        If IsError(resultado) Then
            'MsgBox "El valor '" & nombreColumnaRegresar & "' no se encontró en el rango.", vbExclamation
            Logger "E", "Archivo " & mes, "La cabecera '" & nombreColumnaRegresar & "' no se encuentra en la fila 1 pos " & numeroColumnaRegresar
        Else
            numeroColumnaRevisada = resultado
            Logger "W", "Archivo " & mes, "La cabecera '" & nombreColumnaRegresar & "' reposicionada a pos " & numeroColumnaRevisada
            
        End If
    End If

    Dim NameCol As String
    'MsgBox headerRow
    'MsgBox Hcols(numeroColumnaRegresar)
    'MsgBox numeroColumnaRevisada
    ' Recorrer las líneas del archivo y realizar la suma de la columna correspondiente
    For i = 1 To UBound(lines)
     ' MsgBox Trim(lines(i))
      If Trim(lines(i)) <> "" Then
        'MsgBox Hcols(numeroColumnaRegresar)
        fileLine = lines(i)
        
        
        Dim columns() As String
        columns = Split(fileLine, ";")
        NameCol = columns(numeroColumnaRevisada - 1)
        'MsgBox NameCol
        If UBound(columns) >= numeroColumnaRevisada - 1 Then
          
          ' Verificar si el valor de la columna B se encuentra en el arreglo de valores
          If IsInArray(columns(2 - 1), valoresColumnaB) Then 'Columna B
              'MsgBox columns(numeroColumnaRevisada - 1)
              ' Realizar la suma de la columna correspondiente
              'If IsNumeric(Trim(columns(numeroColumnaRevisada - 1).value)) And Len(Trim(columns(numeroColumnaRevisada - 1).value)) > 0 Then
                
 
              If columns(numeroColumnaRevisada - 1) <> "-" Then
                If IsNumeric(columns(numeroColumnaRevisada - 1)) And columns(numeroColumnaRevisada - 1) <> "-" Then
                  
                  sumaColumnaRegresar = sumaColumnaRegresar + CDbl(columns(numeroColumnaRevisada - 1))
                  'sumaColumnaRegresar = sumaColumnaRegresar + CDbl(columns(numeroColumnaRevisada - 1))
                  
                Else
                  If columns(numeroColumnaRevisada - 1) <> "" Then
                    Logger "W", "Compras", "Archivo " & mes & " tiene valor inválido en fila " & i
                  End If
                End If
              End If
          End If
        End If
      End If
    Next i
    
    ' Mostrar la suma obtenida
    SumaPorColumnaParametrica = sumaColumnaRegresar
End Function

Function IsInArray(ByVal value As Variant, ByVal arr As Variant) As Boolean
    Dim element As Variant
    For Each element In arr
        If element = value Then
            IsInArray = True
            Exit Function
        End If
    Next element
    IsInArray = False
End Function
Sub ImprimirListayear(ByRef Lista As Collection)
    ' Imprime los datos de la colección en una hoja de Excel
    Dim ws As Worksheet
    Dim i As Long
    Dim j As Long
    Dim Col As Long
    Dim elemento As Variant
    
    ' Crear o seleccionar la hoja
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Listadoyear")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = "Listadoyear"
    End If
    On Error GoTo 0
    
    ' Limpiar la hoja
    ws.Cells.Clear
    ws.Cells(1, 1).value = "Código"
    ws.Cells(1, 2).value = "Descripción"
    ws.Cells(1, 3).value = "Valor"
    
    ' Escribir los datos en la hoja
    i = 2
    
    
    For Each elemento In Lista
        ws.Cells(i, Col + 1).value = elemento
         
        i = i + 1
    Next elemento
      
     
    
    'MsgBox "Lista impresa en la hoja 'Listado'.", vbInformation
End Sub

 
Sub LeerRuta()
    Dim HojaResultados As Worksheet
    ' Definir la hoja de resultados
    Set HojaResultados = Sheets("Archivos")
    
    
    Dim miListayear As Collection
    Set miListayear = New Collection ' Inicializas la colección
    Dim respuesta As VbMsgBoxResult
    respuesta = MsgBox("¿Requiere incluir carga de archivo PDF para F29?", vbYesNo + vbQuestion, "Confirmación")
    
    
    ' Obtener las rutas de las carpetas desde la hoja "Params"
    F29Folder = Sheets("Param").Range("B2").value
    RCFolder = Sheets("Param").Range("B4").value
    RVFolder = Sheets("Param").Range("B6").value
    
   
    Set miListayear = Agregaryear(LeerCarpetasPoranio(F29Folder, 1, 0, "F29Folder"), miListayear)
   
    Set miListayear = Agregaryear(LeerCarpetasPoranio(RCFolder, 1, 0, "RCFolder"), miListayear)
        
    Set miListayear = Agregaryear(LeerCarpetasPoranio(RVFolder, 1, 0, "RVFolder"), miListayear)
    
    Call ImprimiryearArchivo(miListayear)
    ' Evaluar la respuesta del usuario
    If respuesta = vbYes Then
        HojaResultados.Range("C2:E100").ClearContents
        Call LimpiarHoja("Dataset")
        Call LeerCarpetasPoranio(F29Folder, 1, 1, "F29Folder")
        Call LeerCarpetasPoranio(RCFolder, 2, 1, "RCFolder")
        Call LeerCarpetasPoranio(RVFolder, 3, 1, "RVFolder")
        
    Else
        HojaResultados.Range("D2:E100").ClearContents
        Call LeerCarpetasPoranio(RCFolder, 2, 1, "RCFolder")
        Call LeerCarpetasPoranio(RVFolder, 3, 1, "RVFolder")
    End If
    
    
End Sub
Sub ImprimiryearArchivo(ByRef ListaYears As Collection)
    Dim ws As Worksheet
    Dim filaInicio As Long
    Dim year As Variant
    Dim mes As Integer
    
    ' Seleccionar o inicializar la hoja "Archivo"
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Archivos")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = "Archivos"
    End If
    On Error GoTo 0

    ' Limpiar la hoja
    'ws.Cells.Clear
 
    
    
    ws.Range("A2:B1000").ClearContents
    ws.Cells(1, 1).value = "Año"
    ws.Cells(1, 2).value = "Mes"

    ' Inicializar la fila de inicio
    filaInicio = 2

    ' Recorrer cada año en la lista
    For Each year In ListaYears
        ' Imprimir el año y los meses del 1 al 12
        For mes = 1 To 12
            ws.Cells(filaInicio, 1).value = year
            ws.Cells(filaInicio, 2).value = ObtenerNombreMes(mes)
            filaInicio = filaInicio + 1
        Next mes
        
        ' Agregar 3 filas en blanco antes del siguiente año
        filaInicio = filaInicio + 1
    Next year

    ' Ajustar columnas
    ws.columns("A:B").AutoFit
End Sub

Function Agregaryear(Lista As Collection, ByRef destino As Collection) As Collection
    Dim elemento As Variant
    Dim yaExiste As Boolean
    Dim item As Variant
 

    ' Agregar elementos de Lista a Destino si no existen
    For Each elemento In Lista
        yaExiste = False
        For Each item In destino
            If item = elemento Then
                yaExiste = True
                Exit For
            End If
        Next item

        If Not yaExiste Then
            destino.Add elemento
        End If
    Next elemento
    Set Agregaryear = destino
End Function


Function LeerCarpetasPoranio(rutaBase, ByVal tipo As Integer, modo As Integer, NombreForm As String) As Collection
    
    Dim fileSystem As Object
    Dim carpeta As Object
    Dim subCarpeta As Object
    Dim anio As Long
    Dim posicion As Integer
    Dim Listyear As Collection
    Dim i As Integer
    Dim Flag As Boolean
    
    Set Listyear = New Collection ' Inicializas la colección
    Flag = True
    
     
    ' Crear una instancia del sistema de archivos
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    
    ' Validar si la ruta base es válida y existe
    If fileSystem.FolderExists(rutaBase) Then
        Set carpeta = fileSystem.GetFolder(rutaBase)
        
    Else
        MsgBox "Error: La carpeta " & NombreForm & " no existe o la ruta es inválida."
        Flag = False
        Set carpeta = Nothing ' Opcional para manejar nulos
        
        
    End If

    If Flag Then
        ' Recorrer todas las subcarpetas
        posicion = 2
        i = 0
        For Each subCarpeta In carpeta.SubFolders
            ' Validar si el nombre de la subcarpeta es un año (formato YYYY)
            If IsNumeric(subCarpeta.Name) And Len(subCarpeta.Name) = 4 Then
                anio = CLng(subCarpeta.Name)
                Listyear.Add (anio)
                
                
                If modo > 0 Then
                    If i = 0 Then
                        posicion = ObtenerFilaInicioDesdeHoja(anio)
                    End If
                    ' Llamar a la función LeerArchivos con el año como parámetro
                    Call LeerArchivos(anio, tipo, posicion)
                    posicion = posicion + 13
                End If
            End If
            i = i + 1
        Next subCarpeta
        
        
    End If
    Set LeerCarpetasPoranio = Listyear
End Function

Sub LeerArchivos(ByVal yearInput As Long, ByVal tipo As Integer, ByVal posicion As Integer)

    Dim F29Folder As String
    Dim RCFolder As String
    Dim RVFolder As String
    Dim ruta As String
    Dim Archivos As String
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim pos As Integer
    Dim NumeroMes As Integer
    Dim HojaResultados As Worksheet
    Dim HojaResultadosLabelyyyymm As Worksheet
    Dim F29Path As String
    Dim Path As String
    Dim Data As String
    
        
    F29Path = Sheets("Param").Range("B2").value
    ' Concatenar el año a la ruta
    
    ' Obtener las rutas de las carpetas desde la hoja "Params"
    F29Folder = Sheets("Param").Range("B2").value
    RCFolder = Sheets("Param").Range("B4").value
    RVFolder = Sheets("Param").Range("B6").value

    ' Concatenar el año a la ruta
    F29Folder = F29Folder & "\" & yearInput ' Concatenar año al final de la ruta
    RCFolder = RCFolder & "\" & yearInput ' Concatenar año al final de la ruta
    RVFolder = RVFolder & "\" & yearInput ' Concatenar año al final de la ruta
    
    ' Definir la hoja de resultados
    Set HojaResultados = Sheets("Archivos")
    

    Select Case tipo
        Case 1
            ruta = F29Folder & "\"
            Archivos = Dir(ruta & "*.pdf")
           
            i = 1
            j = 3 ' Columna C en la hoja de resultados
            pos = posicion
        
         
            Do While Archivos <> "" And i <= 12
               
                NumeroMes = ExtraerMesArchivo(Archivos)
               
                HojaResultados.Cells(pos - 1 + NumeroMes, j).value = Archivos
                Path = ""
                Path = F29Path & "\" & HojaResultados.Cells(pos - 1 + NumeroMes, j - 2).value & "\" & Archivos
                
                If Path <> "" Then
                    Data = ProcesaDataSecuencial(Path, yearInput)
                    Call Imprimir("dataset", Archivos, Data, pos - 1 + NumeroMes)
                End If
                Archivos = Dir
                'pos = pos + 1
                i = i + 1
            Loop
            Logger "I", "Archivos F29", "Se han encontrado " & i - 1 & " archivos para procesar."
    
        Case 2
            ' Leer los archivos de la segunda carpeta (RCFolder)
            ruta = RCFolder & "\"
            Archivos = Dir(ruta & "*.csv")
            i = 1
            j = 4 ' Columna D en la hoja de resultados
            pos = posicion
            Do While Archivos <> "" And i <= 12
                NumeroMes = ExtraerMesArchivo(Archivos)
                HojaResultados.Cells(pos - 1 + NumeroMes, j).value = Archivos
                Archivos = Dir
             
                i = i + 1
            Loop
            Logger "I", "Archivos RC", "Se han encontrado " & i - 1 & " archivos para procesar."
    
        Case Else
            ' Leer los archivos de la tercera carpeta (RVFolder)
            ruta = RVFolder & "\"
            Archivos = Dir(ruta & "*.csv")
            i = 1
            j = 5 ' Columna E en la hoja de resultados
            pos = posicion
            Do While Archivos <> "" And i <= 12
                NumeroMes = ExtraerMesArchivo(Archivos)
                HojaResultados.Cells(pos - 1 + NumeroMes, j).value = Archivos
                Archivos = Dir
           
                i = i + 1
            Loop
            Logger "I", "Archivos RV", "Se han encontrado " & i - 1 & " archivos para procesar."

    End Select
    Dim wsDest As Worksheet

   
  Call DeterminarRangoAños("Ventas", 3)
  Call DeterminarRangoAños("Compras", 4)
  Call DeterminarRangoAños("Ventas", 5)

End Sub

Function ObtenerFilaInicioDesdeHoja(ByVal AñoBuscado As Long) As Long
    Dim ws As Worksheet
    Dim PrimerAño As Long
    Dim FilaBase As Long
    Dim PosicionRelativa As Long

    ' Referenciar la hoja "Archivo"
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Archivos")
    If ws Is Nothing Then
        MsgBox "La hoja 'Archivo' no existe.", vbCritical
        ObtenerFilaInicioDesdeHoja = -1
        Exit Function
    End If
    On Error GoTo 0

    ' Obtener el PrimerAño desde la celda A2
    PrimerAño = ws.Cells(2, 1).value
    If Not IsNumeric(PrimerAño) Or PrimerAño = 0 Then
        MsgBox "No se encontró un año válido en la celda A2.", vbCritical
        ObtenerFilaInicioDesdeHoja = -1
        Exit Function
    End If

    ' Fila inicial del PrimerAño
    FilaBase = 2

    ' Calcular la posición relativa (diferencia en años)
    PosicionRelativa = AñoBuscado - PrimerAño

    ' Calcular la fila de inicio
    ObtenerFilaInicioDesdeHoja = FilaBase + PosicionRelativa * 13
End Function
Sub Imprimir(nombre As String, Archivo As String, Data As String, pos As Integer)
    ' Imprime los datos de la colección en una hoja de Excel
    Dim ws As Worksheet
    Dim i As Long
    Dim j As Long
    Dim Col As Long
    Dim elemento As Variant
    
    ' Crear o seleccionar la hoja
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(nombre)
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = nombre
    End If
    On Error GoTo 0
    
    ' Limpiar la hoja
    'ws.Cells.Clear
    ws.Cells(pos, 1).value = Archivo
    ws.Cells(pos, 2).value = Data
   
    
End Sub
Sub LimpiarHoja(nombre As String)
    ' Imprime los datos de la colección en una hoja de Excel
    Dim ws As Worksheet
    
    
    ' Crear o seleccionar la hoja
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(nombre)
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = nombre
    End If
    On Error GoTo 0
    
    ' Limpiar la hoja
    ws.Cells.Clear
    
     ' Ocultar normalmente
    ws.Visible = xlSheetHidden
   
   
    
End Sub
Sub DeterminarRangoAños(NombreHoja As String, Columna As Integer)
    Dim wsSrc As Worksheet
    Dim wsDest As Worksheet
    Dim UltimaFila As Long
    Dim AñoInicio As Long
    Dim AñoFinal As Long
    Dim i As Long
    Dim AñoActual As Long

    ' Referenciar las hojas
    Set wsSrc = ThisWorkbook.Sheets("Archivos")
    Set wsDest = ThisWorkbook.Sheets(NombreHoja)
    
    ' Obtener la última fila de datos en la columna A
    UltimaFila = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row

    ' Inicializar valores de inicio y final
    AñoInicio = 0
    AñoFinal = 0

    ' Recorrer las filas en la hoja "Archivos"
    For i = 2 To UltimaFila
        AñoActual = wsSrc.Cells(i, 1).value ' Leer el año en la columna A
        If wsSrc.Cells(i, Columna).value <> "" Then ' Verificar si hay un archivo en la columna C
            If AñoInicio = 0 Then
                AñoInicio = AñoActual ' Asignar el primer año encontrado
            End If
            AñoFinal = AñoActual ' Actualizar el último año encontrado
        End If
    Next i
     

    ' Asignar los valores directamente en las celdas de la hoja "F29"
    wsDest.Cells(1, 7).value = AñoInicio ' Columna G: Año de inicio
    wsDest.Cells(1, 9).value = AñoFinal ' Columna I: Año final
 
End Sub

Function ObtenerNombreMes(ByVal NumeroMes As Integer) As String
    Dim meses As Variant
    ' Arreglo con los nombres de los meses
    meses = Array("Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", _
                  "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre")
    
    ' Validar que el número del mes esté en el rango válido (1 a 12)
    If NumeroMes >= 1 And NumeroMes <= 12 Then
        ObtenerNombreMes = meses(NumeroMes - 1)
    Else
        ObtenerNombreMes = "Mes inválido"
    End If
End Function


Sub Logger(tipo As String, seccion As String, Descripcion As String)
    Dim ws As Worksheet
    Dim consecutivo&
    Dim UltimaFila As Long
    
    ' Establecer la hoja de Excel donde se almacenará el log
    Set ws = ThisWorkbook.Sheets("Log")
    
    ' Obtener la última fila utilizada en la hoja de Excel
    UltimaFila = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Obtener el consecutivo sumando 1 a la última fila
    consecutivo = UltimaFila + 1
    
    ' Escribir los valores en las columnas correspondientes
    ws.Cells(consecutivo, 1).value = consecutivo - 1
    ws.Cells(consecutivo, 2).value = Format(Now(), "dd/mm/yyyy hh:mm:ss") ' Agregar fecha/hora formateada
    ws.Cells(consecutivo, 3).value = seccion
    ws.Cells(consecutivo, 4).value = tipo
    ws.Cells(consecutivo, 5).value = Descripcion
    
    ' Establecer el color de fondo de la celda de acuerdo al tipo
    Select Case tipo
        Case "E"
            ws.Cells(consecutivo, 4).Font.Color = RGB(255, 0, 0) ' Rojo
        Case "W"
            ws.Cells(consecutivo, 4).Font.Color = RGB(255, 165, 0) ' Naranja
        Case "I"
            ws.Cells(consecutivo, 4).Font.Color = RGB(0, 0, 255) ' Azul
        Case "S"
            ws.Cells(consecutivo, 4).Font.Color = RGB(0, 128, 0) ' Verde
    End Select
End Sub

Sub LimpiarLog()
    Dim ws As Worksheet
    
    ' Establecer la hoja de Excel donde se encuentra el log
    Set ws = ThisWorkbook.Sheets("Log")
    
    ' Limpiar el contenido de las columnas del log
    ws.Range("A2:E" & Application.Max(2, ws.Cells(ws.Rows.Count, 1).End(xlUp).Row)).ClearContents
    
    ' Limpiar el color de fondo de las celdas del tipo
    ws.Range("D2:D" & Application.Max(2, ws.Cells(ws.Rows.Count, 1).End(xlUp).Row)).ClearFormats
End Sub


Option Explicit

' --- PÉGALA AQUÍ (fuera de cualquier Sub/Function) ---
Public Function ExtraerMesArchivo(ByVal nombre As String) As Integer
    Dim partes() As String
    Dim mm As String
    
    ' Caso A: nombres tipo "..._MM_YYYY.ext"
    partes = Split(nombre, "_")
    If UBound(partes) >= 2 Then
        mm = Left$(partes(2), 2)
        If IsNumeric(mm) Then
            ExtraerMesArchivo = CInt(mm)
            Exit Function
        End If
    End If
    
    ' Caso B: nombres que empiezan "MM." (ej: "01.20.pdf")
    If Len(nombre) >= 2 And IsNumeric(Left$(nombre, 2)) Then
        ExtraerMesArchivo = CInt(Left$(nombre, 2))
        Exit Function
    End If
    
    Err.Raise vbObjectError + 513, , "Nombre de archivo inesperado: " & nombre
End Function
' --- FIN FUNCIÓN ---


