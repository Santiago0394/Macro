Option Explicit

'---------------------------------------------
' UTILIDADES DE UBICACIÓN / AÑOS / BLOQUES
'---------------------------------------------
Function GetYearPosition(ByVal searchYear As Integer) As Long
    Dim ws As Worksheet
    Dim yearRange As Range
    Dim yearCell As Range
    Dim uniqueYears As Collection
    Dim i As Long

    Set ws = ThisWorkbook.Sheets("Archivos")
    Set yearRange = ws.Range("A2:A1000")
    Set uniqueYears = New Collection

    On Error Resume Next
    For Each yearCell In yearRange
        If Not IsEmpty(yearCell.value) Then
            uniqueYears.Add yearCell.value, CStr(yearCell.value)
        End If
    Next yearCell
    On Error GoTo 0

    For i = 1 To uniqueYears.Count
        If uniqueYears(i) = searchYear Then
            GetYearPosition = i - 1
            Exit Function
        End If
    Next i

    GetYearPosition = -1
End Function

Private Function UltimoMesCSV(ByVal year As Long) As Integer
    Dim ws As Worksheet, lastRow As Long, r As Long, c As Long
    Dim s As String, mm As Integer, maxm As Integer
    Set ws = ThisWorkbook.Sheets("Archivos")
    maxm = 0
    For c = 4 To 5 ' D y E
        lastRow = ws.Cells(ws.Rows.Count, c).End(xlUp).Row
        For r = 1 To lastRow
            s = Trim$(CStr(ws.Cells(r, c).value))
            If Len(s) >= 11 And Right$(s, 4) = ".csv" Then
                If Mid$(s, 4, 4) = CStr(year) Then
                    mm = val(Left$(s, 2))
                    If mm >= 1 And mm <= 12 Then
                        If mm > maxm Then maxm = mm
                    End If
                End If
            End If
        Next r
    Next c
    UltimoMesCSV = maxm
End Function

Private Sub LimitesBloqueF29(ByVal year As Long, ByRef startRow As Long, ByRef endRow As Long)
    Dim ws As Worksheet, lastRow As Long, r As Long
    Set ws = ThisWorkbook.Sheets("F29")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    startRow = 0: endRow = 0

    For r = 1 To lastRow
        If ws.Cells(r, "A").value = year Then
            startRow = r
            Exit For
        End If
    Next r
    If startRow = 0 Then Exit Sub

    For r = startRow + 1 To lastRow
        If IsNumeric(ws.Cells(r, "A").value) _
           And ws.Cells(r, "A").value >= 1900 _
           And ws.Cells(r, "A").value <= 9999 Then
            endRow = r - 1
            Exit For
        End If
    Next r
    If endRow = 0 Then endRow = lastRow
End Sub

'---------------------------------------------
' ANCLA: buscar "Impuesto determinado o remanente"
'---------------------------------------------
Private Function FindRowImpuestoRem(ByVal year As Long) As Long
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("F29")
    Dim startRow As Long, endRow As Long, r As Long, t As String
    FindRowImpuestoRem = 0
    LimitesBloqueF29 year, startRow, endRow
    If startRow = 0 Or endRow = 0 Then Exit Function
    For r = startRow + 1 To endRow
        t = UCase$(Trim$(CStr(ws.Cells(r, "A").value)))
        If (InStr(t, "IMPUESTO DETERMINADO") > 0) And (InStr(t, "REMANENTE") > 0) Then
            FindRowImpuestoRem = r
            Exit Function
        End If
    Next r
End Function

'---------------------------------------------
' Asegura un código EXACTAMENTE en una fila fija,
' moviendo si existe o insertando si no existe,
' y borra duplicados del mismo código.
' Devuelve la fila final del código.
'---------------------------------------------
Private Function EnsureCodeAtRow(ByVal year As Long, ByVal targetRow As Long, _
                                 ByVal code As String, ByVal glosa As String) As Long
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("F29")
    Dim startRow As Long, endRow As Long, r As Long, curRow As Long, codeNum As Long
    Dim rowsToDelete As Collection: Set rowsToDelete = New Collection
    codeNum = CLng(code)

    LimitesBloqueF29 year, startRow, endRow
    If startRow = 0 Or endRow = 0 Then Exit Function

    ws.Range(ws.Cells(startRow + 1, "B"), ws.Cells(endRow, "B")).NumberFormat = "000"

    curRow = 0
    For r = startRow + 1 To endRow
        If val(ws.Cells(r, "B").value) = codeNum Then
            If curRow = 0 Then
                curRow = r
            Else
                rowsToDelete.Add r
            End If
        End If
    Next r

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    If curRow = 0 Then
        ws.Rows(targetRow).Insert Shift:=xlDown
        If targetRow > 1 Then
            ws.Rows(targetRow).EntireRow.RowHeight = ws.Rows(targetRow - 1).EntireRow.RowHeight
            ws.Rows(targetRow).Font.Bold = ws.Rows(targetRow - 1).Font.Bold
            ws.Rows(targetRow).Interior.Color = ws.Rows(targetRow - 1).Interior.Color
        End If
        If Len(Trim$(glosa)) > 0 Then ws.Cells(targetRow, "A").value = glosa
        ws.Cells(targetRow, "B").NumberFormat = "000"
        ws.Cells(targetRow, "B").value = codeNum
        ws.Cells(targetRow, "C").Resize(1, 12).ClearContents
        ws.Cells(targetRow, "C").Resize(1, 12).NumberFormat = "#,##0"
        curRow = targetRow
    Else
        If curRow <> targetRow Then
            If curRow < targetRow Then targetRow = targetRow - 1
            ws.Rows(curRow).Cut
            ws.Rows(targetRow).Insert Shift:=xlDown
            curRow = targetRow
        End If
        ws.Cells(curRow, "B").NumberFormat = "000"
        If Len(Trim$(glosa)) > 0 And Trim$(CStr(ws.Cells(curRow, "A").value)) = "" Then
            ws.Cells(curRow, "A").value = glosa
        End If
    End If

    Dim i As Long
    For i = rowsToDelete.Count To 1 Step -1
        ws.Rows(rowsToDelete(i)).Delete
    Next i

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    EnsureCodeAtRow = curRow
End Function

' === 062 -> 123 -> 703 -> 048 -> 151 -> 596 -> 810 -> 049 -> 091 ===
Private Sub EnsureOrden062_123_703(ByVal year As Long)
    Dim startRow As Long, endRow As Long
    Dim labelRow As Long, fallbackRow As Long
    Dim r062 As Long, r123 As Long, r703 As Long, r048 As Long, r151 As Long
    Dim r596 As Long, r810 As Long, r049 As Long, r091 As Long

    LimitesBloqueF29 year, startRow, endRow
    If startRow = 0 Or endRow = 0 Then Exit Sub

    ' 062 debajo del rótulo "Impuesto determinado o remanente"
    labelRow = FindRowImpuestoRem(year)
    If labelRow = 0 Then
        fallbackRow = Application.WorksheetFunction.Max(startRow + 2, endRow - 1)
        r062 = EnsureCodeAtRow(year, fallbackRow, "062", "PPM NETO DETERMINADO")
    Else
        r062 = EnsureCodeAtRow(year, labelRow + 1, "062", "PPM NETO DETERMINADO")
    End If

    ' Orden fijo
    r123 = EnsureCodeAtRow(year, r062 + 1, "123", "")
    r703 = EnsureCodeAtRow(year, r123 + 1, "703", "")
    r048 = EnsureCodeAtRow(year, r703 + 1, "048", "RET. IMP. ÚNICO TRAB. ART. 74 N° 1 LIR")
    r151 = EnsureCodeAtRow(year, r048 + 1, "151", "")
    r596 = EnsureCodeAtRow(year, r151 + 1, "596", "RETENCIÓN CAMBIO DE SUJETO")
    r810 = EnsureCodeAtRow(year, r596 + 1, "810", "")
    r049 = EnsureCodeAtRow(year, r810 + 1, "049", "")   ' <-- NUEVO: 049 antes del 091
    r091 = EnsureCodeAtRow(year, r049 + 1, "091", "TOTAL A PAGAR DENTRO DEL PLAZO LEGAL")
End Sub







'---------------------------------------------
' Otras utilidades F29
'---------------------------------------------
Private Function FilaConCodigo(ws As Worksheet, ByVal r As Long) As Boolean
    FilaConCodigo = IsNumeric(ws.Cells(r, "B").value)
End Function

Public Sub CorrigeMesesFuturosF29(ByVal year As Long)
    Dim ws As Worksheet
    Dim startRow As Long, endRow As Long
    Dim lastMonth As Integer, clearFromCol As Integer, r As Long, c As Long
    Set ws = ThisWorkbook.Sheets("F29")

    lastMonth = UltimoMesCSV(year)
    If lastMonth = 12 Then Exit Sub

    LimitesBloqueF29 year, startRow, endRow
    If startRow = 0 Or endRow = 0 Then Exit Sub

    clearFromCol = 3 + lastMonth
    If clearFromCol <= 14 Then
        Application.ScreenUpdating = False
        For r = startRow + 1 To endRow
            If FilaConCodigo(ws, r) Then
                For c = clearFromCol To 14
                    ws.Cells(r, c).ClearContents
                Next c
            End If
        Next r
        Application.ScreenUpdating = True
    End If

    On Error Resume Next
    RecalcTotalesF29
    On Error GoTo 0
End Sub

'---------------------------------------------
' BOTÓN F29
'---------------------------------------------
Sub Boton_F29_fom()
    Dim ws As Worksheet
    Dim startYear As Long, endYear As Long, yearDiff As Long
    Dim i As Long, yearInt As Long, CountYear As Integer
    Dim wsDest As Worksheet
    Dim positionRow As Long  ' necesario

    Call LimpiarValores(3)
    Call LimpiaFormatea(3)

    Set ws = ThisWorkbook.Sheets("F29")
    startYear = ws.Range("G2").value
    endYear = ws.Range("I2").value
    yearDiff = endYear - startYear

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

        positionRow = 6
        CountYear = 1
        For i = 0 To yearDiff
            yearInt = startYear + i
            Call ProcesarF29(yearInt, positionRow, CountYear)
            positionRow = positionRow + 60
            CountYear = CountYear + 13
        Next i
    End If

    Call CorrigeMesesFuturosF29(endYear)
End Sub

Function Validaryear(ByVal yearBuscado As Long, ByVal Col As String) As Boolean
    Dim ws As Worksheet
    Dim rngYears As Range
    Dim celda As Range

    Set ws = ThisWorkbook.Sheets("Archivos")
    Set rngYears = ws.Range("A2:A100")

    Validaryear = False
    For Each celda In rngYears
        If celda.value = yearBuscado Then
            If ws.Cells(celda.Row, columns(Col).Column).value <> "" Then
                Validaryear = True
                Exit Function
            End If
        End If
    Next celda
End Function

'---------------------------------------------
' PROCESAR F29
'---------------------------------------------
Private Sub ProcesarF29(ByVal yearInput As Long, ByVal positionRow As Integer, ByVal CountY As Integer)
    Dim F29Path As String
    Dim wsFiles As Worksheet
    Dim wsDataset As Worksheet
    Dim mes As Integer
    Dim posArchivo As Integer
    Dim posicion As Long
    Dim Dataset As String
    Dim archivoPDF As String, archivoTXT As String

    posicion = GetYearPosition(yearInput)
    If posicion > 0 Then posicion = posicion * 13
    posicion = posicion + 1

    F29Path = Sheets("Param").Range("B2").value
    F29Path = F29Path & "\" & CStr(yearInput)

    Set wsFiles = Sheets("Archivos")
    Set wsDataset = Sheets("dataset")

    posArchivo = posicion

    ' **Orden determinista 062->123->703 a partir de 2023**
    If yearInput >= 2023 Then
        EnsureOrden062_123_703 yearInput
    End If

    For mes = 1 To 12
        archivoPDF = F29Path & "\" & wsFiles.Cells(posArchivo + mes, 3).value
        archivoTXT = wsFiles.Cells(posArchivo + mes, 3).value

        If wsFiles.Cells(posArchivo + mes, 3).value <> "" Then
            Dataset = wsDataset.Cells(posArchivo + mes, 2).value
            If Dataset <> "" Then
                Call BuscarCodigos(mes, positionRow, Dataset, yearInput)
            Else
                archivoTXT = archivoTXT & ": Sin datos"
                MsgBox archivoTXT
            End If
        End If
    Next mes
End Sub

'---------------------------------------------
' LECTURA DE LISTA/DATASET PARA F29
'---------------------------------------------
Sub BuscarCodigos(mes As Integer, ByVal positionRow As Integer, Dataset As String, ByVal yearObjetivo As Long)
    Dim i As Integer, wsSummary As Worksheet
    Dim codigo As String, desc As String, Result As Variant
    Dim Records As Integer, pos As Integer
    Dim Lista As Collection: Set Lista = New Collection

    Set Lista = GetDataPdf(Dataset, yearObjetivo)
    Records = 90
    Set wsSummary = ThisWorkbook.Sheets("F29")

    pos = positionRow
    For i = pos To Records + pos
        codigo = wsSummary.Cells(pos, 2).value
        desc = UCase(wsSummary.Cells(pos, 1).value)

        If Len(codigo) > 0 And IsNumeric(codigo) Then
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
    Dim inicio As Long, fin As Long, medio As Long
    Dim CodigoActual As Variant

    inicio = 1
    fin = Lista.Count

    While inicio <= fin
        medio = (inicio + fin) \ 2
        CodigoActual = Lista(medio)(1)
        If CodigoActual = CodigoBuscado Or CInt(CodigoActual) = CInt(CodigoBuscado) Then
            ObtenerValorPorCodigoBinario = Lista(medio)(3)
            Exit Function
        ElseIf CInt(CodigoBuscado) < CInt(CodigoActual) Then
            fin = medio - 1
        Else
            inicio = medio + 1
        End If
    Wend
    ObtenerValorPorCodigoBinario = -1
End Function

Private Function PosicionAnioEnTexto(ByVal texto As String, ByVal yearObjetivo As Long) As Long
    Dim yearStr As String
    Dim i As Long
    Dim prevChar As String, nextChar As String

    yearStr = CStr(yearObjetivo)
    For i = 1 To Len(texto) - Len(yearStr) + 1
        If Mid$(texto, i, Len(yearStr)) = yearStr Then
            If i > 1 Then prevChar = Mid$(texto, i - 1, 1) Else prevChar = ""
            If i + Len(yearStr) <= Len(texto) Then
                nextChar = Mid$(texto, i + Len(yearStr), 1)
            Else
                nextChar = ""
            End If
            If Not (Len(prevChar) > 0 And prevChar Like "[0-9]") _
               And Not (Len(nextChar) > 0 And nextChar Like "[0-9]") Then
                PosicionAnioEnTexto = i
                Exit Function
            End If
        End If
    Next i
End Function

Private Function DetectYearInText(ByVal texto As String) As Long
    Dim i As Long, posible As Long
    Dim fragmento As String
    Dim prevChar As String, nextChar As String

    For i = 1 To Len(texto) - 3
        fragmento = Mid$(texto, i, 4)
        If IsNumeric(fragmento) Then
            If i > 1 Then prevChar = Mid$(texto, i - 1, 1) Else prevChar = ""
            If i + 4 <= Len(texto) Then
                nextChar = Mid$(texto, i + 4, 1)
            Else
                nextChar = ""
            End If
            If Not (Len(prevChar) > 0 And prevChar Like "[0-9]") _
               And Not (Len(nextChar) > 0 And nextChar Like "[0-9]") Then
                posible = CLng(fragmento)
                If posible >= 1900 And posible <= 2100 Then
                    DetectYearInText = posible
                    Exit Function
                End If
            End If
        End If
    Next i
End Function

Function GetDataPdf(Dataset As String, Optional ByVal yearObjetivo As Long = 0) As Collection
    Dim texto As String, Bloques() As String
    Dim codigo As String, Descripcion As String, Valor As String
    Dim contenido As String
    Dim i As Long, j As Long, k As Long
    Dim Bloque As String
    Dim tokens() As String
    Dim valueIdx As Long
    Dim miLista As Collection: Set miLista = New Collection
    Dim posicion As Long
    Dim posYear As Long
    Dim currentYear As Long
    Dim detectedYear As Long

    texto = Dataset

    If yearObjetivo <> 0 Then
        posYear = PosicionAnioEnTexto(texto, yearObjetivo)
        If posYear > 0 Then texto = Mid$(texto, posYear)  ' descarta declaraciones de años posteriores
    End If

    ' 1) Localiza el encabezado (con/sin acento) y corta desde ahí
    posicion = InStr(1, texto, "Código Glosa Valor", vbTextCompare)
    If posicion = 0 Then posicion = InStr(1, texto, "Cdigo Glosa Valor", vbTextCompare)
    If posicion > 0 Then texto = Mid$(texto, posicion)

    ' 2) Normaliza separadores a "|"
    texto = Replace(texto, "Código Glosa Valor", "|", , , vbTextCompare)
    texto = Replace(texto, "Cdigo Glosa Valor", "|", , , vbTextCompare)
    texto = Replace(texto, "TOTAL A PAGAR DENTRO DEL PLAZO LEGAL", "|", , , vbTextCompare)
    texto = Replace(texto, "TOTAL A PAGAR", "|", , , vbTextCompare)
    texto = Replace(texto, "+", "|")

    Bloques = Split(texto, "|")

    For i = LBound(Bloques) To UBound(Bloques)
        Bloque = Application.Trim(Bloques(i))
        If Len(Bloque) = 0 Then GoTo SiguienteBloque

        If yearObjetivo <> 0 Then
            detectedYear = DetectYearInText(Bloque)
            If detectedYear >= 1900 And detectedYear <= 2100 Then currentYear = detectedYear
        End If

        If Not ParseBloqueCodigo(Bloque, codigo, contenido) Then GoTo SiguienteBloque

        If yearObjetivo <> 0 Then
            If currentYear <> 0 And currentYear <> yearObjetivo Then GoTo SiguienteBloque
        End If

        tokens = Split(Application.Trim(contenido), " ")
        valueIdx = -1

        ' === Valor: desde el final busca el primer número REAL que no parezca código ===
        ' (ignora tokens numéricos de 3 dígitos para no capturar 151/596/048, etc.)
        Valor = ""
        For k = UBound(tokens) To 0 Step -1
            Valor = FormatearNumero(tokens(k))
            If Valor <> "" And Valor <> "-" And IsNumeric(Valor) Then
                If Len(Valor) > 3 Then Exit For
            End If
        Next k
        If Valor <> "" And IsNumeric(Valor) Then valueIdx = k

        ' Si aún no encontró (casos raros con montos muy pequeños), repite sin la regla de >3
        If Valor = "" Or Not IsNumeric(Valor) Then
            For k = UBound(tokens) To 0 Step -1
                Valor = FormatearNumero(tokens(k))
                If Valor <> "" And Valor <> "-" And IsNumeric(Valor) Then
                    valueIdx = k
                    Exit For
                End If
            Next k
        End If

        ' Descripción = tokens intermedios
        Descripcion = ""
        If UBound(tokens) >= 0 Then
            If valueIdx = -1 Then valueIdx = UBound(tokens) + 1
            For j = 0 To valueIdx - 1
                If Len(tokens(j)) > 0 Then
                    Descripcion = Descripcion & tokens(j) & " "
                End If
            Next j
            Descripcion = Trim$(Descripcion)
        End If
        Descripcion = "    " & Descripcion & "    "

        ' Caso 091 tolerante (si faltara un token de valor)
        If CLng(codigo) = 91 And (Valor = "" Or Not IsNumeric(Valor)) Then
            If UBound(tokens) >= 0 Then
                Valor = FormatearNumero(tokens(UBound(tokens)))
            End If
        End If

        miLista.Add Array(miLista, codigo, Descripcion, Valor)

        ' Final de la primera declaración: evitar arrastrar meses/años posteriores
        If codigo = "091" Then
            If yearObjetivo = 0 Then
                Exit For
            ElseIf currentYear = 0 Or currentYear = yearObjetivo Then
                Exit For
            Else
                ' Reinicia el año actual cuando el primer total pertenece a otra declaración,
                ' para que los siguientes bloques sin etiqueta de año no hereden el periodo equivocado.
                currentYear = 0
            End If
        End If

SiguienteBloque:
    Next i

    Set miLista = OrdenarListaConWorksheetFunction(miLista)
    Set GetDataPdf = miLista
End Function





Function OrdenarListaConWorksheetFunction(miLista As Collection) As Collection
    Dim ws As Worksheet, miLista2 As Collection, i As Long
    Set miLista2 = New Collection

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Temp")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = "Temp"
    End If
    ws.Visible = xlSheetHidden

    For i = 1 To miLista.Count
        ws.Cells(i, 1).value = miLista(i)(1)
        ws.Cells(i, 2).value = miLista(i)(2)
        ws.Cells(i, 3).value = miLista(i)(3)
    Next i

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

    For i = 1 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        miLista2.Add Array(miLista2, ws.Cells(i, 1).value, ws.Cells(i, 2).value, ws.Cells(i, 3).value)
    Next i

    ws.Cells.ClearContents
    Set OrdenarListaConWorksheetFunction = miLista2
End Function

Private Function ParseBloqueCodigo(ByVal Bloque As String, ByRef codigo As String, ByRef contenido As String) As Boolean
    Dim regex As Object, matches As Object
    Dim codigoRaw As String

    Bloque = Trim$(Bloque)
    If Len(Bloque) = 0 Then Exit Function

    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Global = False
        .IgnoreCase = True
        .Pattern = "^\s*(\d{1,3})(?:\s*[-:–—]?\s*)(.*)$"
    End With

    If Not regex.Test(Bloque) Then Exit Function

    Set matches = regex.Execute(Bloque)
    If matches.Count = 0 Then Exit Function

    codigoRaw = matches(0).SubMatches(0)
    If Not BuscarEnArray(codigoRaw) Then Exit Function

    codigo = Format$(CLng(codigoRaw), "000")
    contenido = Application.Trim(matches(0).SubMatches(1))
    ParseBloqueCodigo = True
End Function

Function FormatearNumero(ByVal Valor As String) As String
    Dim s As String, ch As String, i As Long, out As String, neg As Boolean

    s = Valor
    s = Replace(s, Chr(160), "")         ' NBSP
    s = Replace(s, ChrW(8239), "")       ' thin NBSP
    s = Replace(s, "-", "-")             ' minus unicode
    s = Replace(s, "–", "-")
    s = Replace(s, "—", "-")
    s = Trim$(s)

    ' Paréntesis = negativo
    If Left$(s, 1) = "(" And Right$(s, 1) = ")" Then
        neg = True
        s = Mid$(s, 2, Len(s) - 2)
    End If

    ' Quitar separadores de miles/decimales chilenos
    s = Replace(s, ".", "")
    s = Replace(s, ",", "")

    ' Dejar solo dígitos
    out = ""
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch Like "[0-9]" Then out = out & ch
    Next i

    If out = "" Then
        FormatearNumero = ""
    ElseIf neg Then
        FormatearNumero = "-" & out
    Else
        FormatearNumero = out
    End If
End Function


'---------------------------------------------
' COMPRAS
'---------------------------------------------
Sub boton01()
    Dim ws As Worksheet
    Dim startYear As Long, endYear As Long, yearDiff As Long
    Dim yearInt As Long, i As Long, CountYear As Integer
    Dim positionRow As Long

    Call LimpiarValores(1)
    Call LimpiaFormatea(1)

    Set ws = ThisWorkbook.Sheets("Compras")
    startYear = ws.Range("G1").value
    endYear = ws.Range("I1").value
    yearDiff = endYear - startYear

    If yearDiff < 0 Then
        MsgBox "El valor de año final debe ser mayor o igual que el de inicio.", vbExclamation
        Exit Sub
    End If

    If Not Validaryear(startYear, "D") Or Not Validaryear(endYear, "D") Then
        MsgBox "Periodo Anual fuera de rango"
    Else
        Call ReplicarPlanillaCalculo(startYear, 1, "Compras", yearDiff)

        positionRow = 5
        CountYear = 1
        For i = 0 To yearDiff
            yearInt = startYear + i
            Call ProcesarRC(yearInt, positionRow, CountYear)
            positionRow = positionRow + 19
            CountYear = CountYear + 13
            If i > 0 Then
                Call ReplicarDynamicFormula("Compras", i, 1)
            End If
        Next i
    End If
End Sub

Private Sub ProcesarRC(ByVal yearInput As Long, ByVal positionRow As Long, ByVal CountY As Integer)
    Dim RC_Path As String, filepath As String
    Dim wsSummary As Worksheet, wsFiles As Worksheet
    Dim mes As Integer, arr() As Variant
    Dim pos As Integer, posArchivo As Integer
    Dim posicion As Integer

    posicion = GetYearPosition(yearInput)
    If posicion > 0 Then posicion = posicion * 13
    posicion = posicion + 1

    RC_Path = Sheets("Param").Range("B4").value
    RC_Path = RC_Path & "\" & yearInput

    Set wsSummary = ThisWorkbook.Sheets("Compras")
    Set wsFiles = Sheets("Archivos")

    pos = positionRow
    posArchivo = posicion

    For mes = 1 To 12
        filepath = RC_Path & "\" & wsFiles.Cells(posArchivo + mes, 4).value
        If wsFiles.Cells(posArchivo + mes, 4).value <> "" Then
            arr = Array("61")
            wsSummary.Cells(pos + mes, 4) = Abs(SumaPorColumnaParametrica(filepath, mes, "Tipo Doc", arr, "Monto IVA Recuperable", 12))

            arr = Array("33", "39", "46", "56", "914")
            wsSummary.Cells(pos + mes, 9) = SumaPorColumnaParametrica(filepath, mes, "Tipo Doc", arr, "Monto IVA Recuperable", 12) - wsSummary.Cells(pos + mes, 4)

            arr = Array("33", "34", "39", "46", "56", "914")
            wsSummary.Cells(pos + mes, 12) = SumaPorColumnaParametrica(filepath, mes, "Tipo Doc", arr, "Monto Exento", 10)
            arr = Array("61")
            wsSummary.Cells(pos + mes, 12) = wsSummary.Cells(pos + mes, 12) + SumaPorColumnaParametrica(filepath, mes, "Tipo Doc", arr, "Monto Exento", 10)
        End If
    Next mes
End Sub

'---------------------------------------------
' VENTAS
'---------------------------------------------
Sub Buton02()
    Dim ws As Worksheet
    Dim yearInt As Long, yearDiff As Long
    Dim startYear As Long, endYear As Long
    Dim i As Long, CountYear As Integer
    Dim positionRow As Long

    Call LimpiarValores(2)
    Call LimpiaFormatea(2)

    Set ws = ThisWorkbook.Sheets("Ventas")
    startYear = ws.Range("G1").value
    endYear = ws.Range("I1").value
    yearDiff = endYear - startYear

    If yearDiff < 0 Then
        MsgBox "El valor de año final debe ser mayor o igual que el de inicio.", vbExclamation
        Exit Sub
    End If

    If Not Validaryear(startYear, "E") Or Not Validaryear(endYear, "E") Then
        MsgBox "Periodo Anual fuera de rango"
    Else
        Call ReplicarPlanillaCalculo(startYear, 2, "Ventas", yearDiff)

        positionRow = 4
        CountYear = 1
        For i = 0 To yearDiff
            yearInt = startYear + i
            Call ProcesarRV(yearInt, positionRow, CountYear)
            positionRow = positionRow + 19
            CountYear = CountYear + 13
            If i > 0 Then
                Call ReplicarDynamicFormula("Ventas", i, 2)
            End If
        Next i
    End If
End Sub

Private Sub ProcesarRV(ByVal yearInput As Long, ByVal positionRow As Long, ByVal CountY As Integer)
    Dim RV_Path As String, filepath As String
    Dim wsSummary As Worksheet, wsFiles As Worksheet
    Dim mes As Integer, arr() As Variant, filtros() As Variant
    Dim pos As Integer, posArchivo As Integer
    Dim posicion As Integer
    Dim total33 As Double, total56 As Double, total61 As Double

    posicion = GetYearPosition(yearInput)
    If posicion > 0 Then posicion = posicion * 13
    posicion = posicion + 1

    RV_Path = Sheets("Param").Range("B6").value & "\" & CStr(yearInput)

    Set wsSummary = ThisWorkbook.Sheets("Ventas")
    Set wsFiles = Sheets("Archivos")

    pos = positionRow
    posArchivo = posicion

    For mes = 1 To 12
        filepath = RV_Path & "\" & wsFiles.Cells(posArchivo + mes, 5).value

        If wsFiles.Cells(posArchivo + mes, 5).value <> "" Then
            If yearInput >= 2023 Then
                filtros = Array("33")
                total33 = SumaPorColumnaParametrica(filepath, mes, "Tipo Doc", filtros, "Monto IVA", 12)

                filtros = Array("56")
                total56 = SumaPorColumnaParametrica(filepath, mes, "Tipo Doc", filtros, "Monto IVA", 12)

                filtros = Array("61")
                total61 = SumaPorColumnaParametrica(filepath, mes, "Tipo Doc", filtros, "Monto IVA", 12)

                wsSummary.Cells(pos + mes, 4).value = total33 + total56 - total61
            Else
                arr = Array("33", "39", "43", "46", "56", "61")
                wsSummary.Cells(pos + mes, 4).value = SumaPorColumnaParametrica(filepath, mes, "Tipo Doc", arr, "Monto IVA", 12)
            End If

            arr = Array("33", "34", "41", "46", "56", "61")
            wsSummary.Cells(pos + mes, 8).value = SumaPorColumnaParametrica(filepath, mes, "Tipo Doc", arr, "Monto Exento", 10)

            arr = Array("110", "111", "112")
            wsSummary.Cells(pos + mes, 11).value = SumaPorColumnaParametrica(filepath, mes, "Tipo Doc", arr, "Monto Exento", 10)
        End If
    Next mes
End Sub

'---------------------------------------------
' FORMATOS Y TOTALES F29
'---------------------------------------------
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

Private Sub LimpiarFilasPuente(ByVal hoja As String, ParamArray filas())
    Dim ws As Worksheet, k As Long, f As Long
    Dim rngRow As Range, c As Range
    Dim seen As Object
    Set ws = ThisWorkbook.Sheets(hoja)
    Set seen = CreateObject("Scripting.Dictionary")

    For k = LBound(filas) To UBound(filas)
        f = CLng(filas(k))
        Set rngRow = ws.Range("B" & f & ":M" & f)

        On Error Resume Next
        rngRow.ClearContents
        If Err.Number <> 0 Then
            Err.Clear: On Error GoTo 0
            For Each c In rngRow.Cells
                If c.MergeCells Then
                    If Not seen.Exists(c.MergeArea.Address) Then
                        c.MergeArea.ClearContents
                        seen.Add c.MergeArea.Address, True
                    End If
                Else
                    c.ClearContents
                End If
            Next c
        Else
            On Error GoTo 0
        End If
    Next k
End Sub

Sub ReplicarPlanillaCalculo(startYear As Long, NumeroHoja As Integer, NombreHoja As String, yearDiff As Long)
    Dim ws As Worksheet
    Dim sourceRange As Range, destinationRange As Range
    Dim offsetRows As Long, i As Long
    Dim inicio As String, destino As String, Rango As String
    Dim NumberoffsetRows As Integer, NumberoffsetYear As Integer

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

    If NumeroHoja = 2 Then
        Dim stepRows As Long, baseRow As Long, k As Long, r As Long
        stepRows = ws.Range(Rango).Rows.Count + NumberoffsetRows
        baseRow = 20
        LimpiarFilasPuente NombreHoja, baseRow
        For k = 1 To yearDiff
            r = baseRow + (stepRows * k) - 2
            LimpiarFilasPuente NombreHoja, r, r + 1, r + 2
        Next k
    End If

    If NumeroHoja = 1 Then
        Dim stepRowsC As Long, baseRowC As Long, kc As Long, rc As Long
        stepRowsC = ws.Range(Rango).Rows.Count + NumberoffsetRows
        baseRowC = 21
        LimpiarFilasPuente NombreHoja, 20, 21
        For kc = 1 To yearDiff
            rc = baseRowC + (stepRowsC * kc) - 2
            LimpiarFilasPuente NombreHoja, rc, rc + 1, rc + 2
        Next kc
    End If
End Sub

Function GetAdjustedCell(columnOffset As Integer, baseRow As Long) As String
    GetAdjustedCell = "$C$" & (baseRow + (60 * columnOffset))
End Function

Sub ReplicarDynamicFormula(ByVal NombreHoja As String, ByVal columnOffset As Integer, ByVal tipo As Integer)
    Dim wsDest As Worksheet
    Dim i As Integer
    Dim formulaString01 As String, formulaString02 As String, formulaString03 As String, formulaString04 As String, formulaString05 As String
    Dim startRow As Integer, rowOffset As Integer

    Set wsDest = ThisWorkbook.Sheets(NombreHoja)

    Select Case tipo
        Case 1 ' COMPRAS
            startRow = (columnOffset * 19) + 6
            rowOffset = startRow
            Dim baseYearC As Long: baseYearC = ThisWorkbook.Sheets("Compras").Range("G1").value
            Dim yC As Long: yC = baseYearC + columnOffset
            Dim r528 As Long: r528 = FindRowByYearAndCode(yC, "528")
            Dim r537 As Long: r537 = FindRowByYearAndCode(yC, "537")
            Dim r504 As Long: r504 = FindRowByYearAndCode(yC, "504")
            Dim r562 As Long: r562 = FindRowByYearAndCode(yC, "562")

            If r528 > 0 Then formulaString01 = "=DESREF('F29'!$C$" & r528 & ";0;FILA()-" & rowOffset & ")" Else formulaString01 = "0"
            If r537 > 0 Then formulaString02 = "=DESREF('F29'!$C$" & r537 & ";0;FILA()-" & rowOffset & ")" Else formulaString02 = "0"
            If r504 > 0 Then formulaString04 = "=DESREF('F29'!$C$" & r504 & ";0;FILA()-" & rowOffset & ")" Else formulaString04 = "0"
            If r562 > 0 Then formulaString05 = "=DESREF('F29'!$C$" & r562 & ";0;FILA()-" & rowOffset & ")" Else formulaString05 = "0"

        Case 2 ' VENTAS
            startRow = (columnOffset * 19) + 5
            rowOffset = startRow
            Dim baseYearV As Long: baseYearV = ThisWorkbook.Sheets("Ventas").Range("G1").value
            Dim y As Long: y = baseYearV + columnOffset
            Dim r538 As Long: r538 = FindRowByYearAndCode(y, "538")
            Dim r142 As Long: r142 = FindRowByYearAndCode(y, "142")
            Dim r020 As Long: r020 = FindRowByYearAndCode(y, "020")

            If r538 > 0 Then formulaString01 = "=DESREF('F29'!$C$" & r538 & ";0;FILA()-" & rowOffset & ")" Else formulaString01 = "0"
            If r142 > 0 Then formulaString02 = "=DESREF('F29'!$C$" & r142 & ";0;FILA()-" & rowOffset & ")" Else formulaString02 = "0"
            If r020 > 0 Then formulaString03 = "=DESREF('F29'!$C$" & r020 & ";0;FILA()-" & rowOffset & ")" Else formulaString03 = "0"
        Case Else
            MsgBox "Número de hoja no válido.", vbExclamation
            Exit Sub
    End Select

    For i = startRow To startRow + 11
        If tipo = 1 Then
            wsDest.Cells(i, 3).FormulaLocal = formulaString01
            wsDest.Cells(i, 6).FormulaLocal = formulaString02
            wsDest.Cells(i, 7).FormulaLocal = formulaString04
            wsDest.Cells(i, 11).FormulaLocal = formulaString05
        Else
            wsDest.Cells(i, 3).FormulaLocal = formulaString01
            wsDest.Cells(i, 6).FormulaLocal = formulaString02
            wsDest.Cells(i, 10).FormulaLocal = formulaString03
        End If
    Next i
End Sub

Sub LimpiaFormatea(NumeroHoja As Integer)
    Dim ws As Worksheet, Rango As Range, NombreHoja As String
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
        Case 1, 2: Set Rango = ws.Range("B20:L60")
        Case 3:    Set Rango = ws.Range("A61:N1000")
    End Select

    If Rango.MergeCells Then Rango.UnMerge
    Rango.EntireRow.Delete
End Sub

Sub LimpiarValores(NumeroHoja As Integer)
    Dim ws As Worksheet, Rango As Range, celda As Range, NombreHoja As String
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
        MsgBox "No se pudo encontrar la hoja: " & NombreHoja, vbCritical
        Exit Sub
    End If

    Select Case NumeroHoja
        Case 1: Set Rango = ws.Range("C6:M17")
        Case 2: Set Rango = ws.Range("C5:L17")
        Case 3: Set Rango = ws.Range("C5:N60")
    End Select

    For Each celda In Rango
        If Not celda.HasFormula Then celda.ClearContents
    Next celda
End Sub

'---------------------------------------------
' LECTURA DE CARPETAS / ARCHIVOS
'---------------------------------------------
Sub LeerRuta()
    Dim HojaResultados As Worksheet
    Set HojaResultados = Sheets("Archivos")

    Dim miListayear As Collection: Set miListayear = New Collection
    Dim respuesta As VbMsgBoxResult
    Dim F29Folder As String, RCFolder As String, RVFolder As String

    respuesta = MsgBox("Requiere incluir carga de archivo PDF para F29?", vbYesNo + vbQuestion, "Confirmación")

    F29Folder = Sheets("Param").Range("B2").value
    RCFolder = Sheets("Param").Range("B4").value
    RVFolder = Sheets("Param").Range("B6").value

    Set miListayear = Agregaryear(LeerCarpetasPoranio(F29Folder, 1, 0, "F29Folder"), miListayear)
    Set miListayear = Agregaryear(LeerCarpetasPoranio(RCFolder, 1, 0, "RCFolder"), miListayear)
    Set miListayear = Agregaryear(LeerCarpetasPoranio(RVFolder, 1, 0, "RVFolder"), miListayear)

    Call ImprimiryearArchivo(miListayear)

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

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Archivos")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = "Archivos"
    End If
    On Error GoTo 0

    ws.Range("A2:B1000").ClearContents
    ws.Cells(1, 1).value = "Año"
    ws.Cells(1, 2).value = "Mes"

    filaInicio = 2

    For Each year In ListaYears
        For mes = 1 To 12
            ws.Cells(filaInicio, 1).value = year
            ws.Cells(filaInicio, 2).value = ObtenerNombreMes(mes)
            filaInicio = filaInicio + 1
        Next mes
        filaInicio = filaInicio + 1
    Next year
    ws.columns("A:B").AutoFit
End Sub

Function Agregaryear(Lista As Collection, ByRef destino As Collection) As Collection
    Dim elemento As Variant, yaExiste As Boolean, item As Variant
    For Each elemento In Lista
        yaExiste = False
        For Each item In destino
            If item = elemento Then yaExiste = True: Exit For
        Next item
        If Not yaExiste Then destino.Add elemento
    Next elemento
    Set Agregaryear = destino
End Function

Function LeerCarpetasPoranio(rutaBase, ByVal tipo As Integer, modo As Integer, NombreForm As String) As Collection
    Dim fileSystem As Object, carpeta As Object, subCarpeta As Object
    Dim anio As Long, posicion As Integer
    Dim Listyear As Collection: Set Listyear = New Collection
    Dim i As Integer, Flag As Boolean

    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    If fileSystem.FolderExists(rutaBase) Then
        Set carpeta = fileSystem.GetFolder(rutaBase)
        Flag = True
    Else
        MsgBox "Error: La carpeta " & NombreForm & " no existe o la ruta es inválida."
        Flag = False
        Set carpeta = Nothing
    End If

    If Flag Then
        posicion = 2
        i = 0
        For Each subCarpeta In carpeta.SubFolders
            If IsNumeric(subCarpeta.Name) And Len(subCarpeta.Name) = 4 Then
                anio = CLng(subCarpeta.Name)
                Listyear.Add (anio)
                If modo > 0 Then
                    If i = 0 Then posicion = ObtenerFilaInicioDesdeHoja(anio)
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
    Dim F29Folder As String, RCFolder As String, RVFolder As String
    Dim ruta As String, Archivos As String, i As Integer, j As Integer, pos As Integer
    Dim NumeroMes As Integer, HojaResultados As Worksheet
    Dim F29Path As String, Path As String, Data As String

    F29Path = Sheets("Param").Range("B2").value
    F29Folder = Sheets("Param").Range("B2").value & "\" & yearInput
    RCFolder = Sheets("Param").Range("B4").value & "\" & yearInput
    RVFolder = Sheets("Param").Range("B6").value & "\" & yearInput

    Set HojaResultados = Sheets("Archivos")

    Select Case tipo
        Case 1
            ruta = F29Folder & "\"
            Archivos = Dir(ruta & "*.pdf")
            i = 1: j = 3: pos = posicion
            Do While Archivos <> "" And i <= 12
                NumeroMes = ExtraerMesArchivo(Archivos)
                HojaResultados.Cells(pos - 1 + NumeroMes, j).value = Archivos
                Path = F29Path & "\" & HojaResultados.Cells(pos - 1 + NumeroMes, j - 2).value & "\" & Archivos
                If Path <> "" Then
                    Data = ProcesaDataSecuencial(Path)
                    Call Imprimir("dataset", Archivos, Data, pos - 1 + NumeroMes)
                End If
                Archivos = Dir
                i = i + 1
            Loop
            Logger "I", "Archivos F29", "Se han encontrado " & i - 1 & " archivos para procesar."
        Case 2
            ruta = RCFolder & "\"
            Archivos = Dir(ruta & "*.csv")
            i = 1: j = 4: pos = posicion
            Do While Archivos <> "" And i <= 12
                NumeroMes = ExtraerMesArchivo(Archivos)
                HojaResultados.Cells(pos - 1 + NumeroMes, j).value = Archivos
                Archivos = Dir
                i = i + 1
            Loop
            Logger "I", "Archivos RC", "Se han encontrado " & i - 1 & " archivos para procesar."
        Case Else
            ruta = RVFolder & "\"
            Archivos = Dir(ruta & "*.csv")
            i = 1: j = 5: pos = posicion
            Do While Archivos <> "" And i <= 12
                NumeroMes = ExtraerMesArchivo(Archivos)
                HojaResultados.Cells(pos - 1 + NumeroMes, j).value = Archivos
                Archivos = Dir
                i = i + 1
            Loop
            Logger "I", "Archivos RV", "Se han encontrado " & i - 1 & " archivos para procesar."
    End Select

    Call DeterminarRangoAos("Ventas", 3)
    Call DeterminarRangoAos("Compras", 4)
    Call DeterminarRangoAos("Ventas", 5)
End Sub

Function ObtenerFilaInicioDesdeHoja(ByVal AoBuscado As Long) As Long
    Dim ws As Worksheet, PrimerAo As Long, FilaBase As Long, PosicionRelativa As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Archivos")
    If ws Is Nothing Then
        MsgBox "La hoja 'Archivo' no existe.", vbCritical
        ObtenerFilaInicioDesdeHoja = -1
        Exit Function
    End If
    On Error GoTo 0

    PrimerAo = ws.Cells(2, 1).value
    If Not IsNumeric(PrimerAo) Or PrimerAo = 0 Then
        MsgBox "No se encontró un año válido en la celda A2.", vbCritical
        ObtenerFilaInicioDesdeHoja = -1
        Exit Function
    End If

    FilaBase = 2
    PosicionRelativa = AoBuscado - PrimerAo
    ObtenerFilaInicioDesdeHoja = FilaBase + PosicionRelativa * 13
End Function

Sub Imprimir(nombre As String, Archivo As String, Data As String, pos As Integer)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(nombre)
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = nombre
    End If
    On Error GoTo 0
    ws.Cells(pos, 1).value = Archivo
    ws.Cells(pos, 2).value = Data
End Sub

Sub LimpiarHoja(nombre As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(nombre)
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = nombre
    End If
    On Error GoTo 0
    ws.Cells.Clear
    ws.Visible = xlSheetHidden
End Sub

Sub DeterminarRangoAos(NombreHoja As String, Columna As Integer)
    Dim wsSrc As Worksheet, wsDest As Worksheet
    Dim UltimaFila As Long, AoInicio As Long, AoFinal As Long
    Dim i As Long, AoActual As Long

    Set wsSrc = ThisWorkbook.Sheets("Archivos")
    Set wsDest = ThisWorkbook.Sheets(NombreHoja)
    UltimaFila = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row

    AoInicio = 0: AoFinal = 0
    For i = 2 To UltimaFila
        AoActual = wsSrc.Cells(i, 1).value
        If wsSrc.Cells(i, Columna).value <> "" Then
            If AoInicio = 0 Then AoInicio = AoActual
            AoFinal = AoActual
        End If
    Next i

    wsDest.Cells(1, 7).value = AoInicio
    wsDest.Cells(1, 9).value = AoFinal
End Sub

Function ObtenerNombreMes(ByVal NumeroMes As Integer) As String
    Dim meses As Variant
    meses = Array("Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", _
                  "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre")
    If NumeroMes >= 1 And NumeroMes <= 12 Then
        ObtenerNombreMes = meses(NumeroMes - 1)
    Else
        ObtenerNombreMes = "Mes inválido"
    End If
End Function

Sub Logger(tipo As String, seccion As String, Descripcion As String)
    Dim ws As Worksheet, consecutivo&, UltimaFila As Long
    Set ws = ThisWorkbook.Sheets("Log")
    UltimaFila = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    consecutivo = UltimaFila + 1
    ws.Cells(consecutivo, 1).value = consecutivo - 1
    ws.Cells(consecutivo, 2).value = Format(Now(), "dd/mm/yyyy hh:mm:ss")
    ws.Cells(consecutivo, 3).value = seccion
    ws.Cells(consecutivo, 4).value = tipo
    ws.Cells(consecutivo, 5).value = Descripcion
    Select Case tipo
        Case "E": ws.Cells(consecutivo, 4).Font.Color = RGB(255, 0, 0)
        Case "W": ws.Cells(consecutivo, 4).Font.Color = RGB(255, 165, 0)
        Case "I": ws.Cells(consecutivo, 4).Font.Color = RGB(0, 0, 255)
        Case "S": ws.Cells(consecutivo, 4).Font.Color = RGB(0, 128, 0)
    End Select
End Sub

Sub LimpiarLog()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Log")
    ws.Range("A2:E" & Application.Max(2, ws.Cells(ws.Rows.Count, 1).End(xlUp).Row)).ClearContents
    ws.Range("D2:D" & Application.Max(2, ws.Cells(ws.Rows.Count, 1).End(xlUp).Row)).ClearFormats
End Sub

'---------------------------------------------
' EXTRAER MES DEL NOMBRE DE ARCHIVO
'---------------------------------------------
Public Function ExtraerMesArchivo(ByVal nombre As String) As Integer
    Dim partes() As String, mm As String
    partes = Split(nombre, "_")
    If UBound(partes) >= 2 Then
        mm = Left$(partes(2), 2)
        If IsNumeric(mm) Then
            ExtraerMesArchivo = CInt(mm)
            Exit Function
        End If
    End If
    If Len(nombre) >= 2 And IsNumeric(Left$(nombre, 2)) Then
        ExtraerMesArchivo = CInt(Left$(nombre, 2))
        Exit Function
    End If
    Err.Raise vbObjectError + 513, , "Nombre de archivo inesperado: " & nombre
End Function

Private Function FindRowByYearAndCode(ByVal year As Long, ByVal code As String) As Long
    Dim ws As Worksheet
    Dim startRow As Long, endRow As Long, r As Long
    Set ws = ThisWorkbook.Sheets("F29")
    FindRowByYearAndCode = 0
    LimitesBloqueF29 year, startRow, endRow
    If startRow = 0 Or endRow = 0 Then Exit Function
    For r = startRow + 1 To endRow
        If val(ws.Cells(r, "B").value) = val(code) Then
            FindRowByYearAndCode = r
            Exit Function
        End If
    Next r
End Function
'---------------------------------------------
' PROCESO PDF SECUENCIAL (si se usa)
' Inserta "|" antes de cualquier código válido y
' normaliza el código a "000". Para el resto,
' limpia números con FormatearNumero y deja texto tal cual.
'---------------------------------------------
Function ProcesaDataSecuencial(rutaPDF As String) As String
    Dim AcroApp As Object, AcroAVDoc As Object, AcroPDDoc As Object, jsObj As Object
    Dim i As Long, j As Long, nPages As Long, nWords As Long
    Dim tok As String, val As String, out As String

    Set AcroApp = CreateObject("AcroExch.App")
    Set AcroAVDoc = CreateObject("AcroExch.AVDoc")

    If Not AcroAVDoc.Open(rutaPDF, "") Then
        MsgBox "No se pudo abrir el archivo PDF.", vbExclamation
        GoTo Cleanup
    End If

    Set AcroPDDoc = AcroAVDoc.GetPDDoc
    Set jsObj = AcroPDDoc.GetJSObject

    nPages = jsObj.numPages
    For i = 0 To nPages - 1
        nWords = jsObj.getPageNumWords(i)
        For j = 0 To nWords - 1
            tok = Trim$(jsObj.getPageNthWord(i, j))

            If BuscarEnArray(tok) Then
                out = out & "|" & Format$(CLng(tok), "000") & " "
            Else
                val = FormatearNumero(tok)
                If val <> "" Then out = out & val & " " Else out = out & tok & " "
            End If
        Next j
        out = out & vbCrLf
    Next i

    ProcesaDataSecuencial = out

Cleanup:
    On Error Resume Next
    If Not AcroAVDoc Is Nothing Then AcroAVDoc.Close True
    If Not AcroApp Is Nothing Then AcroApp.Exit
    Set jsObj = Nothing
    Set AcroPDDoc = Nothing
    Set AcroAVDoc = Nothing
    Set AcroApp = Nothing
End Function

Private Function SoloDigitos(ByVal s As String) As Boolean
    Dim k As Long, ch As String
    If Len(s) = 0 Then Exit Function
    For k = 1 To Len(s)
        ch = Mid$(s, k, 1)
        If ch < "0" Or ch > "9" Then Exit Function
    Next k
    SoloDigitos = True
End Function

Function BuscarEnArray(codigo As String) As Boolean
    Dim s As String, i As Long
    Dim valid As Variant

    ' Lista en formato "000"
    valid = Array( _
        "010", "020", "028", "030", "039", "048", "049", "050", "054", "056", "062", "066", "068", "077", "089", "091", _
        "110", "111", "113", "115", "120", "122", "123", "127", "142", "151", "152", "153", "154", "155", "156", "157", "164", "409", _
        "500", "501", "502", "503", "504", "509", "510", "511", "512", "513", "514", "515", "516", "517", "518", "519", "520", "521", _
        "522", "523", "524", "525", "526", "527", "528", "529", "530", "531", "532", "534", "535", "536", "537", "538", "539", "540", _
        "541", "542", "543", "544", "547", "548", "550", "553", "557", "560", "562", "563", "564", "565", "566", "573", "584", "585", _
        "586", "587", "588", "589", "592", "593", "594", "595", "596", _
        "700", "701", "702", "703", "708", "709", "711", "712", "713", "714", "715", "716", "717", "718", "720", "721", "722", "723", _
        "724", "729", "730", "731", "732", "734", "735", "738", "739", "740", "741", "742", "743", "744", "745", "749", "750", "751", _
        "755", "756", "757", "758", "759", "761", "762", "763", "764", "772", "773", "774", "775", "776", "777", "778", "779", "780", _
        "782", "783", "784", "785", "786", "787", "788", "789", "791", "792", "793", "794", "796", "797", "798", "799", "800", "801", "802", "803", "804", "805", "806", "807", "808", "809", "810" _
    )

    BuscarEnArray = False
    s = Trim$(codigo)

    ' *** Filtro estricto: solo 1–3 dígitos ***
    If Len(s) = 0 Or Len(s) > 3 Then Exit Function
    If Not SoloDigitos(s) Then Exit Function

    ' Ahora es seguro normalizar
    s = Format$(CLng(s), "000")

    For i = LBound(valid) To UBound(valid)
        If valid(i) = s Then
            BuscarEnArray = True
            Exit Function
        End If
    Next i
End Function
