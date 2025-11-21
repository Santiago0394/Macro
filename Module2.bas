' Devuelve s si no está vacío; en caso contrario, fallback
Private Function NzText(ByVal s As String, ByVal fallback As String) As String
    If Len(Trim$(CStr(s))) > 0 Then NzText = Trim$(CStr(s)) Else NzText = fallback
End Function

' Limpia espacios extras y los padding que ponemos al armar la descripción
Private Function SanitizeGlosa(ByVal s As String) As String
    Dim t As String
    t = Trim$(CStr(s))
    t = Replace(t, vbTab, " ")
    t = Replace(t, "  ", " ")
    Do While InStr(t, "  ") > 0
        t = Replace(t, "  ", " ")
    Loop
    SanitizeGlosa = t
End Function

' Busca en las cadenas del sheet "dataset" la glosa del código indicado para ese año
Private Function GlosaDesdeDatasetPorCodigo(ByVal year As Long, ByVal codeNum As Long) As String
    Dim wsFiles As Worksheet, wsData As Worksheet
    Dim pos As Long, m As Long, dataStr As String
    Dim lst As Collection, i As Long
    On Error GoTo salir

    Set wsFiles = ThisWorkbook.Sheets("Archivos")
    Set wsData = ThisWorkbook.Sheets("dataset")

    pos = ObtenerFilaInicioDesdeHoja(year) ' inicio del bloque (fila del mes enero)
    If pos <= 0 Then GoTo salir

    For m = 0 To 11
        dataStr = CStr(wsData.Cells(pos + m, 2).value) ' col B = dataset
        If Len(dataStr) > 0 Then
            Set lst = GetDataPdf(dataStr)
            For i = 1 To lst.Count
                If Val(lst(i)(1)) = codeNum Then
                    GlosaDesdeDatasetPorCodigo = SanitizeGlosa(lst(i)(2))
                    Exit Function
                End If
            Next i
        End If
    Next m
salir:
End Function


' === Inserta/ubica "Tasa PPM" y calcula por CÓDIGOS (numerador = 062) ===
Private Sub EnsureFilaTasaPPM(ByVal year As Long)
    If year < 2023 Then Exit Sub

    Const LABEL_TASA As String = "Tasa PPM"

    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("F29")
    Dim startRow As Long, endRow As Long, r As Long, c As Long
    Dim r091 As Long, r062 As Long, r020 As Long, r142 As Long, r715 As Long, r538 As Long
    Dim rowTasa As Long, targetRow As Long

    ' Bloque del año
    Call LimitesBloqueF29(year, startRow, endRow)
    If startRow = 0 Or endRow = 0 Then Exit Sub

    ' Filas por CÓDIGO
    r091 = FindRowByYearAndCode(year, "091")         ' para UBICAR debajo de 091
    r062 = FindRowByYearAndCode(year, "062")         ' <-- NUMERADOR correcto
    r020 = FindRowByYearAndCode(year, "020")
    r142 = FindRowByYearAndCode(year, "142")
    r715 = FindRowByYearAndCode(year, "715")
    r538 = FindRowByYearAndCode(year, "538")
    If r538 = 0 Then r538 = FindRowByYearAndLabel(year, "Total débitos")

    ' Si falta alguno, no hacemos nada
    If r091 = 0 Or r062 = 0 Or r020 = 0 Or r142 = 0 Or r715 = 0 Or r538 = 0 Then Exit Sub

    ' Debajo del 091
    targetRow = r091 + 1

    ' ¿Ya existe "Tasa PPM"?
    rowTasa = 0
    For r = startRow + 1 To endRow
        If UCase$(Trim$(CStr(ws.Cells(r, "A").value))) = UCase$(LABEL_TASA) Then
            rowTasa = r: Exit For
        End If
    Next r

    Application.ScreenUpdating = False

    If rowTasa = 0 Then
        ws.Rows(targetRow).Insert Shift:=xlDown
        ws.Rows(targetRow).EntireRow.RowHeight = ws.Rows(targetRow - 1).EntireRow.RowHeight
        ws.Rows(targetRow).Font.Bold = ws.Rows(targetRow - 1).Font.Bold
        ws.Rows(targetRow).Interior.Color = ws.Rows(targetRow - 1).Interior.Color
        ws.Cells(targetRow, "A").value = LABEL_TASA
        ws.Cells(targetRow, "B").ClearContents
        ws.Cells(targetRow, "C").Resize(1, 12).NumberFormat = "0.0000" ' o "0.00%" si lo quieres en %
        rowTasa = targetRow
    ElseIf rowTasa <> targetRow Then
        ws.Rows(rowTasa).Cut
        ws.Rows(targetRow).Insert Shift:=xlDown
        rowTasa = targetRow
    End If

    ' Fórmula por CÓDIGOS: 062 / (020 + 142 + 715 + 538/19)
    For c = 3 To 14
        ws.Cells(rowTasa, c).FormulaR1C1 = _
            "=IFERROR(R" & r062 & "C/(R" & r020 & "C+R" & r142 & "C+R" & r715 & "C+R" & r538 & "C/19),0)"
    Next c

    Application.ScreenUpdating = True
End Sub



Private Sub EnsureFilaCodigo091_DebajoImpuestoAPagar(ByVal year As Long)
    Const LABEL_TARGET As String = "Impuesto a Pagar"
    Const GLOSA As String = "TOTAL A PAGAR DENTRO DEL PLAZO LEGAL"

    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("F29")
    Dim startRow As Long, endRow As Long
    Dim r As Long, rowImp As Long, row091 As Long
    Dim lastCodeRow As Long, targetRow As Long

    ' Bloque del año
    Call LimitesBloqueF29(year, startRow, endRow)
    If startRow = 0 Or endRow = 0 Then Exit Sub

    ' Normaliza formato de códigos en el bloque
    ws.Range(ws.Cells(startRow + 1, "B"), ws.Cells(endRow, "B")).NumberFormat = "000"

    ' Localiza "Impuesto a Pagar", una fila 091 (si existe) y el último renglón con código
    For r = startRow + 1 To endRow
        If UCase$(Trim$(CStr(ws.Cells(r, "A").value))) = UCase$(LABEL_TARGET) Then rowImp = r
        If Val(ws.Cells(r, "B").value) = 91 Then row091 = r
        If IsNumeric(ws.Cells(r, "B").value) Then lastCodeRow = r
    Next r

    ' Debajo de “Impuesto a Pagar”; si no existiera, al final de los códigos
    If rowImp > 0 Then
        targetRow = rowImp + 1
    ElseIf lastCodeRow > 0 Then
        targetRow = lastCodeRow + 1
    Else
        targetRow = endRow + 1
    End If

    Application.ScreenUpdating = False

    If row091 > 0 Then
        ' Mover si no está en la posición correcta
        If row091 <> targetRow Then
            ws.Rows(row091).Cut
            ws.Rows(targetRow).Insert Shift:=xlDown
        End If
    Else
        ' Insertar nueva fila
        ws.Rows(targetRow).Insert Shift:=xlDown

        ' Copiar estética de la fila superior
        ws.Rows(targetRow).EntireRow.RowHeight = ws.Rows(targetRow - 1).EntireRow.RowHeight
        ws.Rows(targetRow).Font.Bold = ws.Rows(targetRow - 1).Font.Bold
        ws.Rows(targetRow).Interior.Color = ws.Rows(targetRow - 1).Interior.Color

        ' Glosa, código y formato de meses
        ws.Cells(targetRow, "A").value = GLOSA
        ws.Cells(targetRow, "B").NumberFormat = "000"
        ws.Cells(targetRow, "B").value = 91
        ws.Cells(targetRow, "C").Resize(1, 12).ClearContents
        ws.Cells(targetRow, "C").Resize(1, 12).NumberFormat = "#,##0"
    End If

    Application.ScreenUpdating = True
End Sub

' Devuelve la fila del código dentro del bloque [startRow..endRow]
' Busca en B (normal) o en A (caso 2025: "703"/"123" en col A).
Private Function RowOfCodigo(ws As Worksheet, startRow As Long, endRow As Long, ByVal codeNum As Long) As Long
    Dim r As Long, a As String
    RowOfCodigo = 0
    For r = startRow + 1 To endRow
        If Val(ws.Cells(r, "B").value) = codeNum Then
            RowOfCodigo = r
        Else
            a = Trim$(CStr(ws.Cells(r, "A").value))
            If Len(a) > 0 And IsNumeric(a) Then
                If Val(a) = codeNum Then RowOfCodigo = r
            End If
        End If
    Next r
End Function

' ---- NUEVO helper: detecta si la fila es de "código"
Private Function EsFilaDeCodigo(ws As Worksheet, ByVal r As Long) As Boolean
    Dim a As String, b As Variant

    ' B numérica (caso normal)
    b = ws.Cells(r, "B").value
    If Len(Trim$(b & "")) > 0 And IsNumeric(b) Then
        EsFilaDeCodigo = True
        Exit Function
    End If

    ' A con número de 1–3 dígitos (p.ej. 703 / 123) y no es la etiqueta
    a = Trim$(CStr(ws.Cells(r, "A").value))
    If Len(a) > 0 And IsNumeric(a) Then
        If Len(a) <= 3 Then EsFilaDeCodigo = True
    End If
End Function
Private Sub EnsureFilaImpuestoAPagar(ByVal year As Long)
    Const LABEL_IMP As String = "Impuesto determinado o remanente"
    Const LABEL_TARGET As String = "Impuesto a Pagar"

    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("F29")
    Dim startRow As Long, endRow As Long
    Dim rowLabel As Long, rowTarget As Long
    Dim lastCodeRow As Long, r As Long, c As Long
    Dim targetRow As Long

    ' Límites del bloque del año
    Call LimitesBloqueF29(year, startRow, endRow)
    If startRow = 0 Or endRow = 0 Then Exit Sub

    ' ===== PARCHE SOLO 2025 =====
    ' Si el 2025 es el último bloque, debajo de endRow puede haber filas con B numérica y A vacía.
    ' Las consideramos parte del bloque para que "Impuesto a Pagar" quede DESPUÉS de ellas.
    If year = 2025 Then
        Dim lastUsedB As Long, r2 As Long
        lastUsedB = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
        For r2 = endRow + 1 To lastUsedB
            ' si aparece algo en A, paramos (sería otra sección)
            If Len(Trim$(CStr(ws.Cells(r2, "A").value))) > 0 Then Exit For
            If IsNumeric(ws.Cells(r2, "B").value) Then endRow = r2
        Next r2
    End If
    ' ===== FIN PARCHE =====

    ' Fila de "Impuesto determinado o remanente"
    rowLabel = FindRowByYearAndLabel(year, LABEL_IMP)
    If rowLabel = 0 Then Exit Sub

    ' Buscar el ÚLTIMO renglón con código (col B numérica) dentro del bloque
    lastCodeRow = 0
    For r = rowLabel To endRow
        If IsNumeric(ws.Cells(r, "B").value) Then
            lastCodeRow = r
        End If
    Next r
    If lastCodeRow = 0 Then Exit Sub

    ' ¿Ya existe "Impuesto a Pagar"?
    rowTarget = 0
    For r = startRow + 1 To endRow
        If UCase$(Trim$(CStr(ws.Cells(r, "A").value))) = UCase$(LABEL_TARGET) Then
            rowTarget = r
            Exit For
        End If
    Next r

    targetRow = lastCodeRow + 1

    Application.ScreenUpdating = False

    ' Insertar o mover la fila justo DESPUÉS del último código
    If rowTarget = 0 Then
        ws.Rows(targetRow).Insert Shift:=xlDown
        ws.Rows(targetRow).EntireRow.RowHeight = ws.Rows(targetRow - 1).EntireRow.RowHeight
        ws.Rows(targetRow).Font.Bold = True
        ws.Rows(targetRow).Interior.Color = ws.Rows(targetRow - 1).Interior.Color
        ws.Cells(targetRow, "A").value = LABEL_TARGET
        ws.Cells(targetRow, "B").ClearContents
        ws.Cells(targetRow, "C").Resize(1, 12).NumberFormat = "#,##0"
    ElseIf rowTarget <> targetRow Then
        ws.Rows(rowTarget).Cut
        ws.Rows(targetRow).Insert Shift:=xlDown
        rowTarget = targetRow
    End If

    ' Fórmulas por mes (C:N) desde la etiqueta hasta el ÚLTIMO código
    For c = 3 To 14
        ws.Cells(targetRow, c).FormulaR1C1 = "=SUM(R" & rowLabel & "C:R" & lastCodeRow & "C)"
    Next c

    ' Columna O vacía
    ws.Cells(targetRow, "O").ClearContents

    Application.ScreenUpdating = True
End Sub


Private Sub EnsureFilaCodigo049_Debajo048(ByVal year As Long)
    Const LABEL_IMP As String = "Impuesto determinado o remanente"
    ' Respaldo si no encontramos la glosa en el dataset
    Const GLOSA As String = "Ret. 3% Rta 42 N°1 reint. prest. Tasa 0%"

    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("F29")
    Dim startRow As Long, endRow As Long
    Dim rowLabel As Long, row062 As Long, row123 As Long, row703 As Long
    Dim row048 As Long, row049 As Long
    Dim targetRow As Long, r As Long
    Dim glosaReal As String

    ' Límites del bloque del año
    Call LimitesBloqueF29(year, startRow, endRow)
    If startRow = 0 Or endRow = 0 Then Exit Sub

    ' Normaliza formato de códigos del bloque
    ws.Range(ws.Cells(startRow + 1, "B"), ws.Cells(endRow, "B")).NumberFormat = "000"

    ' Ubicar filas clave
    rowLabel = FindRowByYearAndLabel(year, LABEL_IMP)
    For r = startRow + 1 To endRow
        Select Case Val(ws.Cells(r, "B").value)
            Case 62:  row062 = r
            Case 123: row123 = r
            Case 703: row703 = r
            Case 48:  row048 = r
            Case 49:  row049 = r
        End Select
    Next r

    ' Posición objetivo para 049: debajo de 048; luego 703 > 123 > 062 > rótulo
    If row048 > 0 Then
        targetRow = row048 + 1
    ElseIf row703 > 0 Then
        targetRow = row703 + 1
    ElseIf row123 > 0 Then
        targetRow = row123 + 1
    ElseIf row062 > 0 Then
        targetRow = row062 + 1
    ElseIf rowLabel > 0 Then
        targetRow = rowLabel + 1
    Else
        targetRow = endRow + 1
    End If

    ' Glosa real desde el dataset; si no hay, usa la constante GLOSA
    glosaReal = NzText(GlosaDesdeDatasetPorCodigo(year, 49), GLOSA)

    Application.ScreenUpdating = False

    If row049 > 0 Then
        ' Mover si hace falta…
        If row049 <> targetRow Then
            ws.Rows(row049).Cut
            ws.Rows(targetRow).Insert Shift:=xlDown
            row049 = targetRow
        End If
        ' …y asegurar la glosa correcta aunque ya existiera
        ws.Cells(row049, "A").value = glosaReal
        ws.Cells(row049, "B").NumberFormat = "000"
        ws.Cells(row049, "B").value = 49
    Else
        ' Insertar nueva fila en la posición objetivo
        ws.Rows(targetRow).Insert Shift:=xlDown

        ' Copiar estética de la fila superior
        ws.Rows(targetRow).EntireRow.RowHeight = ws.Rows(targetRow - 1).EntireRow.RowHeight
        ws.Rows(targetRow).Font.Bold = ws.Rows(targetRow - 1).Font.Bold
        ws.Rows(targetRow).Interior.Color = ws.Rows(targetRow - 1).Interior.Color

        ' Glosa, código y meses
        ws.Cells(targetRow, "A").value = glosaReal
        ws.Cells(targetRow, "B").NumberFormat = "000"
        ws.Cells(targetRow, "B").value = 49
        ws.Cells(targetRow, "C").Resize(1, 12).ClearContents
        ws.Cells(targetRow, "C").Resize(1, 12).NumberFormat = "#,##0"
    End If

    Application.ScreenUpdating = True
End Sub





Private Sub EnsureFilaCodigo810_Debajo596(ByVal year As Long)
    Const LABEL_IMP As String = "Impuesto determinado o remanente"
    Const GLOSA As String = "TOTAL A PAGAR DENTRO DEL PLAZO LEGAL"

    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("F29")
    Dim startRow As Long, endRow As Long
    Dim rowLabel As Long, row062 As Long, row123 As Long, row703 As Long
    Dim row048 As Long, row151 As Long, row596 As Long, row810 As Long
    Dim targetRow As Long, r As Long

    ' Límites del bloque del año
    Call LimitesBloqueF29(year, startRow, endRow)
    If startRow = 0 Or endRow = 0 Then Exit Sub

    ' Normaliza formato de códigos del bloque
    ws.Range(ws.Cells(startRow + 1, "B"), ws.Cells(endRow, "B")).NumberFormat = "000"

    ' Ubicar filas clave
    rowLabel = FindRowByYearAndLabel(year, LABEL_IMP)
    For r = startRow + 1 To endRow
        Select Case Val(ws.Cells(r, "B").value)
            Case 62:  row062 = r
            Case 123: row123 = r
            Case 703: row703 = r
            Case 48:  row048 = r
            Case 151: row151 = r
            Case 596: row596 = r
            Case 810: row810 = r
        End Select
    Next r

    ' Posición objetivo para 810: 596 -> 151 -> 048 -> 703 -> 123 -> 062 -> rótulo
    If row596 > 0 Then
        targetRow = row596 + 1
    ElseIf row151 > 0 Then
        targetRow = row151 + 1
    ElseIf row048 > 0 Then
        targetRow = row048 + 1
    ElseIf row703 > 0 Then
        targetRow = row703 + 1
    ElseIf row123 > 0 Then
        targetRow = row123 + 1
    ElseIf row062 > 0 Then
        targetRow = row062 + 1
    ElseIf rowLabel > 0 Then
        targetRow = rowLabel + 1
    Else
        targetRow = endRow + 1 ' Plan B
    End If

    Application.ScreenUpdating = False

    If row810 > 0 Then
        If row810 <> targetRow Then
            ws.Rows(row810).Cut
            ws.Rows(targetRow).Insert Shift:=xlDown
        End If
    Else
        ws.Rows(targetRow).Insert Shift:=xlDown

        ' Copiar estética de la fila superior
        ws.Rows(targetRow).EntireRow.RowHeight = ws.Rows(targetRow - 1).EntireRow.RowHeight
        ws.Rows(targetRow).Font.Bold = ws.Rows(targetRow - 1).Font.Bold
        ws.Rows(targetRow).Interior.Color = ws.Rows(targetRow - 1).Interior.Color

        ' Glosa, código y meses
        ws.Cells(targetRow, "A").value = GLOSA
        ws.Cells(targetRow, "B").NumberFormat = "000"
        ws.Cells(targetRow, "B").value = 810
        ws.Cells(targetRow, "C").Resize(1, 12).ClearContents
        ws.Cells(targetRow, "C").Resize(1, 12).NumberFormat = "#,##0"
    End If

    Application.ScreenUpdating = True
End Sub

' Devuelve el índice del primer token que sea un código válido (o -1 si no hay)
Private Function FirstCodeIndex(ByRef toks() As String) As Long
    Dim k As Long
    For k = LBound(toks) To UBound(toks)
        If BuscarEnArray(Trim$(toks(k))) Then
            FirstCodeIndex = k
            Exit Function
        End If
    Next k
    FirstCodeIndex = -1
End Function

' Devuelve el último token no vacío del arreglo
Private Function LastNonEmptyToken(ByRef toks() As String) As String
    Dim i As Long
    For i = UBound(toks) To LBound(toks) Step -1
        If Len(Trim$(toks(i))) > 0 Then
            LastNonEmptyToken = Trim$(toks(i))
            Exit Function
        End If
    Next i
    LastNonEmptyToken = ""
End Function


Private Sub EnsureFilaCodigo596_Debajo151(ByVal year As Long)
    Const LABEL_IMP As String = "Impuesto determinado o remanente"
    Const GLOSA As String = "RETENCIÓN CAMBIO DE SUJETO"

    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("F29")
    Dim startRow As Long, endRow As Long
    Dim rowLabel As Long, row062 As Long, row123 As Long, row703 As Long, row048 As Long, row151 As Long, row596 As Long
    Dim targetRow As Long, r As Long

    ' Límites del bloque del año
    Call LimitesBloqueF29(year, startRow, endRow)
    If startRow = 0 Or endRow = 0 Then Exit Sub

    ' Normaliza formato de códigos del bloque
    ws.Range(ws.Cells(startRow + 1, "B"), ws.Cells(endRow, "B")).NumberFormat = "000"

    ' Ubicar filas clave
    rowLabel = FindRowByYearAndLabel(year, LABEL_IMP)
    For r = startRow + 1 To endRow
        Select Case Val(ws.Cells(r, "B").value)
            Case 62:  row062 = r
            Case 123: row123 = r
            Case 703: row703 = r
            Case 48:  row048 = r
            Case 151: row151 = r
            Case 596: row596 = r
        End Select
    Next r

    ' Posición objetivo: 151 -> 048 -> 703 -> 123 -> 062 -> rótulo
    If row151 > 0 Then
        targetRow = row151 + 1
    ElseIf row048 > 0 Then
        targetRow = row048 + 1
    ElseIf row703 > 0 Then
        targetRow = row703 + 1
    ElseIf row123 > 0 Then
        targetRow = row123 + 1
    ElseIf row062 > 0 Then
        targetRow = row062 + 1
    ElseIf rowLabel > 0 Then
        targetRow = rowLabel + 1
    Else
        targetRow = endRow + 1 ' plan B
    End If

    Application.ScreenUpdating = False

    If row596 > 0 Then
        If row596 <> targetRow Then
            ws.Rows(row596).Cut
            ws.Rows(targetRow).Insert Shift:=xlDown
            row596 = targetRow
        End If
    Else
        ws.Rows(targetRow).Insert Shift:=xlDown

        ' Copiar estética de la fila superior
        ws.Rows(targetRow).EntireRow.RowHeight = ws.Rows(targetRow - 1).EntireRow.RowHeight
        ws.Rows(targetRow).Font.Bold = ws.Rows(targetRow - 1).Font.Bold
        ws.Rows(targetRow).Interior.Color = ws.Rows(targetRow - 1).Interior.Color

        ' Glosa, código y meses
        ws.Cells(targetRow, "A").value = GLOSA
        ws.Cells(targetRow, "B").NumberFormat = "000"
        ws.Cells(targetRow, "B").value = 596
        ws.Cells(targetRow, "C").Resize(1, 12).ClearContents
        ws.Cells(targetRow, "C").Resize(1, 12).NumberFormat = "#,##0"
    End If

    Application.ScreenUpdating = True
End Sub


Private Sub EnsureFilaCodigo151_Debajo048(ByVal year As Long)
    Const LABEL_IMP As String = "Impuesto determinado o remanente"
    Const GLOSA As String = "RETENCIÓN TASA LEY 21.133 SOBRE RENTAS"

    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("F29")
    Dim startRow As Long, endRow As Long
    Dim rowLabel As Long, row062 As Long, row123 As Long, row703 As Long, row048 As Long, row151 As Long
    Dim targetRow As Long, r As Long

    ' Límites del bloque del año
    Call LimitesBloqueF29(year, startRow, endRow)
    If startRow = 0 Or endRow = 0 Then Exit Sub

    ' Normaliza formato de códigos del bloque
    ws.Range(ws.Cells(startRow + 1, "B"), ws.Cells(endRow, "B")).NumberFormat = "000"

    ' Ubicar filas clave
    rowLabel = FindRowByYearAndLabel(year, LABEL_IMP)
    For r = startRow + 1 To endRow
        Select Case Val(ws.Cells(r, "B").value)
            Case 62:  row062 = r
            Case 123: row123 = r
            Case 703: row703 = r
            Case 48:  row048 = r
            Case 151: row151 = r
        End Select
    Next r

    ' Posición objetivo: 151 debajo de 048 -> 703 -> 123 -> 062 -> rótulo
    If row048 > 0 Then
        targetRow = row048 + 1
    ElseIf row703 > 0 Then
        targetRow = row703 + 1
    ElseIf row123 > 0 Then
        targetRow = row123 + 1
    ElseIf row062 > 0 Then
        targetRow = row062 + 1
    ElseIf rowLabel > 0 Then
        targetRow = rowLabel + 1
    Else
        targetRow = endRow + 1 ' plan B
    End If

    Application.ScreenUpdating = False

    If row151 > 0 Then
        If row151 <> targetRow Then
            ws.Rows(row151).Cut
            ws.Rows(targetRow).Insert Shift:=xlDown
            row151 = targetRow
        End If
        ' Asegura la glosa correcta aunque ya existiera
        ws.Cells(row151, "A").value = GLOSA
        ws.Cells(row151, "B").NumberFormat = "000"
        ws.Cells(row151, "B").value = 151
    Else
        ws.Rows(targetRow).Insert Shift:=xlDown

        ' Copiar estética de la fila superior
        ws.Rows(targetRow).EntireRow.RowHeight = ws.Rows(targetRow - 1).EntireRow.RowHeight
        ws.Rows(targetRow).Font.Bold = ws.Rows(targetRow - 1).Font.Bold
        ws.Rows(targetRow).Interior.Color = ws.Rows(targetRow - 1).Interior.Color

        ' Glosa, código y meses
        ws.Cells(targetRow, "A").value = GLOSA
        ws.Cells(targetRow, "B").NumberFormat = "000"
        ws.Cells(targetRow, "B").value = 151
        ws.Cells(targetRow, "C").Resize(1, 12).ClearContents
        ws.Cells(targetRow, "C").Resize(1, 12).NumberFormat = "#,##0"
    End If

    Application.ScreenUpdating = True
End Sub



Private Sub EnsureFilaCodigo048_Debajo703(ByVal year As Long)
    Const LABEL_IMP As String = "Impuesto determinado o remanente"
    Const GLOSA As String = "RET. IMP. UNICO TRAB. ART. 74 N°1 LIR"

    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("F29")
    Dim startRow As Long, endRow As Long
    Dim rowLabel As Long, row062 As Long, row123 As Long, row703 As Long, row048 As Long
    Dim targetRow As Long, r As Long

    Call LimitesBloqueF29(year, startRow, endRow)
    If startRow = 0 Or endRow = 0 Then Exit Sub

    ws.Range(ws.Cells(startRow + 1, "B"), ws.Cells(endRow, "B")).NumberFormat = "000"

    rowLabel = FindRowByYearAndLabel(year, LABEL_IMP)
    For r = startRow + 1 To endRow
        Select Case Val(ws.Cells(r, "B").value)
            Case 62:  row062 = r
            Case 123: row123 = r
            Case 703: row703 = r
            Case 48:  row048 = r
        End Select
    Next r

    If row703 > 0 Then
        targetRow = row703 + 1
    ElseIf row123 > 0 Then
        targetRow = row123 + 1
    ElseIf row062 > 0 Then
        targetRow = row062 + 1
    ElseIf rowLabel > 0 Then
        targetRow = rowLabel + 1
    Else
        targetRow = endRow + 1
    End If

    Application.ScreenUpdating = False

    If row048 > 0 Then
        If row048 <> targetRow Then
            ws.Rows(row048).Cut
            ws.Rows(targetRow).Insert Shift:=xlDown
            row048 = targetRow
        End If
    Else
        ws.Rows(targetRow).Insert Shift:=xlDown
        ws.Rows(targetRow).EntireRow.RowHeight = ws.Rows(targetRow - 1).EntireRow.RowHeight
        ws.Rows(targetRow).Font.Bold = ws.Rows(targetRow - 1).Font.Bold
        ws.Rows(targetRow).Interior.Color = ws.Rows(targetRow - 1).Interior.Color

        ' <<< Glosa obligatoria para que BuscarCodigos lo procese >>>
        ws.Cells(targetRow, "A").value = GLOSA
        ws.Cells(targetRow, "B").NumberFormat = "000"
        ws.Cells(targetRow, "B").value = 48
        ws.Cells(targetRow, "C").Resize(1, 12).ClearContents
        ws.Cells(targetRow, "C").Resize(1, 12).NumberFormat = "#,##0"
    End If

    Application.ScreenUpdating = True
End Sub




Private Sub EnsureFilaCodigo703_Debajo123(ByVal year As Long)
    Const LABEL_IMP As String = "Impuesto determinado o remanente"
    Const cod As String = "703"
    Const GLOSA As String = "" ' opcional: coloca aquí el texto exacto si quieres fijar la glosa

    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("F29")
    Dim startRow As Long, endRow As Long
    Dim rowLabel As Long, row062 As Long, row123 As Long, row703 As Long
    Dim targetRow As Long, r As Long

    ' Límites del bloque del año
    Call LimitesBloqueF29(year, startRow, endRow)
    If startRow = 0 Or endRow = 0 Then Exit Sub

    ' Normaliza formato "000" en columna B dentro del bloque
    ws.Range(ws.Cells(startRow + 1, "B"), ws.Cells(endRow, "B")).NumberFormat = "000"

    ' Ubicar filas clave
    rowLabel = FindRowByYearAndLabel(year, LABEL_IMP)
    For r = startRow + 1 To endRow
        If Val(ws.Cells(r, "B").value) = 62 Then row062 = r
        If Val(ws.Cells(r, "B").value) = 123 Then row123 = r
        If Val(ws.Cells(r, "B").value) = 703 Then row703 = r
    Next r

    ' Posición objetivo: debajo de 123; si no hay 123, debajo de 062; luego del rótulo
    If row123 > 0 Then
        targetRow = row123 + 1
    ElseIf row062 > 0 Then
        targetRow = row062 + 1
    ElseIf rowLabel > 0 Then
        targetRow = rowLabel + 1
    Else
        targetRow = endRow + 1  ' plan B: al final del bloque
    End If

    Application.ScreenUpdating = False

    If row703 > 0 Then
        If row703 <> targetRow Then
            ws.Rows(row703).Cut
            ws.Rows(targetRow).Insert Shift:=xlDown
            row703 = targetRow
        End If
    Else
        ws.Rows(targetRow).Insert Shift:=xlDown

        ' Copiar estética de la fila superior
        ws.Rows(targetRow).EntireRow.RowHeight = ws.Rows(targetRow - 1).EntireRow.RowHeight
        ws.Rows(targetRow).Font.Bold = ws.Rows(targetRow - 1).Font.Bold
        ws.Rows(targetRow).Interior.Color = ws.Rows(targetRow - 1).Interior.Color

        ' Setear glosa (si quieres), código y limpiar meses
        If Len(GLOSA) > 0 Then ws.Cells(targetRow, "A").value = GLOSA
        ws.Cells(targetRow, "B").NumberFormat = "000"
        ws.Cells(targetRow, "B").value = 703
        ws.Cells(targetRow, "C").Resize(1, 12).ClearContents
        ws.Cells(targetRow, "C").Resize(1, 12).NumberFormat = "#,##0"
    End If

    Application.ScreenUpdating = True
End Sub


Private Sub EnsureFilaCodigo123_Debajo062(ByVal year As Long)
    Const LABEL_IMP As String = "Impuesto determinado o remanente"
    Const cod As String = "123"
    Const GLOSA As String = "" ' déjalo vacío para no forzar texto; si quieres un nombre fijo, colócalo aquí

    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("F29")
    Dim startRow As Long, endRow As Long
    Dim rowLabel As Long, row062 As Long, row123 As Long, targetRow As Long, r As Long

    ' Límites del bloque del año
    Call LimitesBloqueF29(year, startRow, endRow)
    If startRow = 0 Or endRow = 0 Then Exit Sub

    ' Normaliza formato de códigos del bloque
    ws.Range(ws.Cells(startRow + 1, "B"), ws.Cells(endRow, "B")).NumberFormat = "000"

    ' Fila del rótulo y de 062 (si existe)
    rowLabel = FindRowByYearAndLabel(year, LABEL_IMP)

    For r = startRow + 1 To endRow
        If Val(ws.Cells(r, "B").value) = 62 Then row062 = r
        If Val(ws.Cells(r, "B").value) = 123 Then row123 = r
    Next r

    ' Posición objetivo: debajo de 062; si no existe 062, debajo del rótulo
    If row062 > 0 Then
        targetRow = row062 + 1
    ElseIf rowLabel > 0 Then
        targetRow = rowLabel + 1
    Else
        targetRow = endRow + 1 ' plan B: al final del bloque
    End If

    Application.ScreenUpdating = False

    If row123 > 0 Then
        ' Mover si no está en la posición correcta
        If row123 <> targetRow Then
            ws.Rows(row123).Cut
            ws.Rows(targetRow).Insert Shift:=xlDown
            row123 = targetRow
        End If
    Else
        ' Insertar nueva fila en targetRow
        ws.Rows(targetRow).Insert Shift:=xlDown

        ' Copiar estética de la fila superior para no romper formatos
        ws.Rows(targetRow).EntireRow.RowHeight = ws.Rows(targetRow - 1).EntireRow.RowHeight
        ws.Rows(targetRow).Font.Bold = ws.Rows(targetRow - 1).Font.Bold
        ws.Rows(targetRow).Interior.Color = ws.Rows(targetRow - 1).Interior.Color

        ' Setear glosa (opcional), código y limpiar meses
        If Len(GLOSA) > 0 Then ws.Cells(targetRow, "A").value = GLOSA
        ws.Cells(targetRow, "B").NumberFormat = "000"
        ws.Cells(targetRow, "B").value = 123
        ws.Cells(targetRow, "C").Resize(1, 12).ClearContents
        ws.Cells(targetRow, "C").Resize(1, 12).NumberFormat = "#,##0"
    End If

    Application.ScreenUpdating = True
End Sub


Private Sub EnsureFilaCodigo062_DebajoImpuestoDet(ByVal year As Long)
    Const LABEL_IMP As String = "Impuesto determinado o remanente"
    Const cod As String = "062"
    Const GLOSA As String = "PPM NETO DETERMINADO"

    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("F29")
    Dim startRow As Long, endRow As Long
    Dim rowLabel As Long, row062 As Long, targetRow As Long, r As Long

    ' Límites del bloque del año
    Call LimitesBloqueF29(year, startRow, endRow)
    If startRow = 0 Or endRow = 0 Then Exit Sub

    ' Normaliza formato de códigos del bloque
    ws.Range(ws.Cells(startRow + 1, "B"), ws.Cells(endRow, "B")).NumberFormat = "000"

    ' Fila de la etiqueta y del 062 si ya existe
    rowLabel = FindRowByYearAndLabel(year, LABEL_IMP)
    If rowLabel = 0 Then
        ' Si por alguna razón no está la etiqueta, lo dejamos al final del bloque
        targetRow = endRow + 1
    Else
        targetRow = rowLabel + 1
    End If

    For r = startRow + 1 To endRow
        If Val(ws.Cells(r, "B").value) = 62 Then
            row062 = r
            Exit For
        End If
    Next r

    Application.ScreenUpdating = False

    If row062 > 0 Then
        ' Si existe y no está en el lugar correcto, lo movemos
        If row062 <> targetRow Then
            ws.Rows(row062).Cut
            ws.Rows(targetRow).Insert Shift:=xlDown
            row062 = targetRow
        End If
    Else
        ' No existe: insertar nueva fila en targetRow
        ws.Rows(targetRow).Insert Shift:=xlDown

        ' Copiar estética básica de la fila superior para no romper formatos
        ws.Rows(targetRow).EntireRow.RowHeight = ws.Rows(targetRow - 1).EntireRow.RowHeight
        ws.Rows(targetRow).Font.Bold = ws.Rows(targetRow - 1).Font.Bold
        ws.Rows(targetRow).Interior.Color = ws.Rows(targetRow - 1).Interior.Color

        ' Setear glosa, código y limpiar meses
        ws.Cells(targetRow, "A").value = GLOSA
        ws.Cells(targetRow, "B").NumberFormat = "000"
        ws.Cells(targetRow, "B").value = 62
        ws.Cells(targetRow, "C").Resize(1, 12).ClearContents
        ws.Cells(targetRow, "C").Resize(1, 12).NumberFormat = "#,##0"
    End If

    Application.ScreenUpdating = True
End Sub


Private Function FindRowByYearAndLabel(ByVal year As Long, ByVal labelText As String) As Long
    Dim ws As Worksheet, startRow As Long, endRow As Long, r As Long
    Set ws = ThisWorkbook.Sheets("F29")
    FindRowByYearAndLabel = 0
    Call LimitesBloqueF29(year, startRow, endRow)
    If startRow = 0 Or endRow = 0 Then Exit Function

    For r = startRow + 1 To endRow
        If UCase$(Trim$(CStr(ws.Cells(r, "A").value))) = UCase$(Trim$(labelText)) Then
            FindRowByYearAndLabel = r
            Exit Function
        End If
    Next r
End Function


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

Private Sub EnsureFilaCodigoEnF29(ByVal year As Long, ByVal code As String, ByVal GLOSA As String, _
                                  Optional ByVal preferAfterCode As String = "020", _
                                  Optional ByVal fallbackBeforeCode As String = "142")

    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("F29")
    Dim startRow As Long, endRow As Long, r As Long
    Dim row020 As Long, row142 As Long, row039 As Long, insertRow As Long
    Dim codeNum As Long: codeNum = CLng(code)

    Call LimitesBloqueF29(year, startRow, endRow)
    If startRow = 0 Or endRow = 0 Then Exit Sub

    ' Formatear la columna de códigos del bloque del año a 3 dígitos
    ws.Range(ws.Cells(startRow + 1, "B"), ws.Cells(endRow, "B")).NumberFormat = "000"

    ' Ubicar filas de 020, 142 y 039 (comparación numérica)
    For r = startRow + 1 To endRow
        Select Case True
            Case Val(ws.Cells(r, "B").value) = Val(preferAfterCode):   row020 = r
            Case Val(ws.Cells(r, "B").value) = Val(fallbackBeforeCode): row142 = r
            Case Val(ws.Cells(r, "B").value) = codeNum:                 row039 = r
        End Select
    Next r

    ' Decidir la posición ideal
    If row020 > 0 Then
        insertRow = row020 + 1              ' después de 020
    ElseIf row142 > 0 Then
        insertRow = row142                  ' antes de 142
    Else
        insertRow = startRow + 1            ' muy arriba si no hay 020/142
    End If

    ' Si ya existe 039 pero mal ubicado, mover la fila
    If row039 > 0 And row039 <> insertRow Then
        ws.Rows(row039).Cut
        ws.Rows(insertRow).Insert Shift:=xlDown
        row039 = insertRow
    End If

    ' Si no existe, insertarla
    If row039 = 0 Then
        ws.Rows(insertRow).Insert Shift:=xlDown
        ' Copiar estética de la fila superior
        ws.Rows(insertRow).EntireRow.RowHeight = ws.Rows(insertRow - 1).EntireRow.RowHeight
        ws.Rows(insertRow).Font.Bold = ws.Rows(insertRow - 1).Font.Bold
        ws.Rows(insertRow).Interior.Color = ws.Rows(insertRow - 1).Interior.Color

        ws.Cells(insertRow, "A").value = GLOSA   ' Glosa
        ws.Cells(insertRow, "B").NumberFormat = "000"
        ws.Cells(insertRow, "B").value = codeNum ' 039 mostrado como 000
        ws.Cells(insertRow, "C").Resize(1, 12).ClearContents
        ws.Cells(insertRow, "C").Resize(1, 12).NumberFormat = "#,##0"
    Else
        ' Asegura formato 000 si ya existía
        ws.Cells(row039, "B").NumberFormat = "000"
    End If
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
    
    ' Asegura la fila del 039 desde 2023 en adelante
    If yearInput >= 2023 Then
        EnsureFilaCodigoEnF29 yearInput, "039", "IVA TOT RET. TERC.(TASA ART. 14)", "020", "142"
    End If
    
    ' NUEVO: Asegura 062 justo debajo de "Impuesto determinado o remanente" SOLO 2023–2025
    If yearInput >= 2023 And yearInput <= 2025 Then
        EnsureFilaCodigo062_DebajoImpuestoDet yearInput
        ' NUEVO: Asegura 123 inmediatamente debajo de 062
        EnsureFilaCodigo123_Debajo062 yearInput
        EnsureFilaCodigo703_Debajo123 yearInput
        EnsureFilaCodigo048_Debajo703 yearInput
        EnsureFilaCodigo151_Debajo048 yearInput
        EnsureFilaCodigo596_Debajo151 yearInput
        EnsureFilaCodigo810_Debajo596 yearInput
        EnsureFilaCodigo049_Debajo048 yearInput
        EnsureFilaImpuestoAPagar yearInput
        EnsureFilaCodigo091_DebajoImpuestoAPagar yearInput
        EnsureFilaTasaPPM yearInput
    End If
     
    
    For mes = 1 To 12
        ' Construir la ruta y el nombre del archivo PDF
        
        archivoPDF = F29Path & "\" & wsFiles.Cells(posArchivo + mes, 3).value
        
        archivoTXT = wsFiles.Cells(posArchivo + mes, 3).value
       
        If wsFiles.Cells(posArchivo + mes, 3).value <> "" Then
            Dataset = wsDataset.Cells(posArchivo + mes, 2).value
            'MsgBox Dataset
            If Dataset <> "" Then
                Call BuscarCodigos(mes, positionRow, Dataset)
            Else
                archivoTXT = archivoTXT & ": Sin datos"
                MsgBox archivoTXT
            End If
            
        End If
        
    Next mes
   

End Sub

Sub BuscarCodigos(mes As Integer, ByVal positionRow As Integer, Dataset As String)
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
    
     
    
    Set Lista = GetDataPdf(Dataset)
   
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


Function GetDataPdf(Dataset As String) As Collection
    Dim texto As String
    Dim Bloques() As String
    Dim Lineas() As String
    Dim codigo As Variant
    Dim Descripcion As String
    Dim Valor As String
    Dim resultado As String
    Dim i As Integer
    Dim Bloque As String
    Dim PosTotal As Long
    Dim miLista As Collection
    Dim posicion As Long

    Set miLista = New Collection
    texto = Dataset

    If posicion > 0 Then
        texto = Mid(texto, posicion)
    End If

    posicion = InStr(texto, "Código Glosa Valor")
    texto = Replace(texto, "Código Glosa Valor", "|")
    texto = Replace(texto, "TOTAL A PAGAR DENTRO DEL PLAZO LEGAL", "|")
    texto = Replace(texto, "+", "|")

    Bloques = Split(texto, "|")
    resultado = ""

    For i = LBound(Bloques) To UBound(Bloques)
        Bloque = Trim$(Bloques(i))
        If Bloque <> "" Then
            Lineas = Split(Bloque, " ")

            Dim t As Long, nextIdx As Long
            Dim idxValor As Long, k As Long
            Dim scanEnd As Long, valorToken As String

            t = LBound(Lineas)
            Do While t <= UBound(Lineas)
                If BuscarEnArray(Trim$(Lineas(t))) Then
                    ' Buscar el inicio del próximo código (bucles SAFE: sin AND)
                    nextIdx = t + 1
                    Do While nextIdx <= UBound(Lineas)
                        If BuscarEnArray(Trim$(Lineas(nextIdx))) Then Exit Do
                        nextIdx = nextIdx + 1
                    Loop

                    ' Último token no vacío antes del siguiente código o del fin del bloque
                    scanEnd = nextIdx - 1
                    If scanEnd > UBound(Lineas) Then scanEnd = UBound(Lineas)

                    idxValor = -1
                    For k = scanEnd To t + 1 Step -1
                        If Len(Trim$(Lineas(k))) > 0 Then
                            idxValor = k
                            Exit For
                        End If
                    Next k

                    codigo = Trim$(Lineas(t))

                    If idxValor <> -1 And idxValor >= LBound(Lineas) And idxValor <= UBound(Lineas) Then
                        valorToken = Trim$(Lineas(idxValor))
                    Else
                        valorToken = ""   ' no hay valor explícito
                    End If
                    Valor = FormatearNumero(valorToken)

                    ' Descripción = tokens entre el código y el valor
                    Descripcion = ""
                    If idxValor <> -1 Then
                        For k = t + 1 To idxValor - 1
                            If Len(Trim$(Lineas(k))) > 0 Then
                                Descripcion = Descripcion & Trim$(Lineas(k)) & " "
                            End If
                        Next k
                    End If
                    Descripcion = "    " & Trim$(Descripcion) & "    "

                    ' Ahora (mantiene tu excepción pero segura):
                    If IsNumeric(codigo) And CInt(codigo) = 91 Then
                        If t + 1 <= UBound(Lineas) Then
                            miLista.Add Array(miLista, codigo, Descripcion, FormatearNumero(Trim$(Lineas(t + 1))))
                        Else
                            ' si no hay t+1, usa el valor ya calculado con idxValor
                            miLista.Add Array(miLista, codigo, Descripcion, Valor)
                        End If
                    Else
                        miLista.Add Array(miLista, codigo, Descripcion, Valor)
                    End If

                    ' Continuar desde el siguiente posible código
                    If nextIdx > t Then
                        t = nextIdx
                    Else
                        t = t + 1
                    End If
                Else
                    t = t + 1
                End If
            Loop
        End If
    Next i

    Set miLista = OrdenarListaConWorksheetFunction(miLista)
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
Function ProcesaDataSecuencial(rutaPDF As String) As String 'Obtiene el texto completo desde el PDF F29
   
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
                If BuscarEnArray(palabra) And Len(palabra) <= 3 And Flag Then
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
Function BuscarEnArray(codigo As String) As Boolean
    Dim arr As Variant
    Dim i As Long
    
    ' Definimos el array con datos de prueba
    arr = Array("010", "020", "028", "030", "039", "39", "048", "48", "049", "050", "054", "056", "062", "066", "068", "077", "089", "091", "91", "110", "111", _
"113", "115", "120", "122", "123", "127", "142", "151", "152", "153", "154", "155", "156", "157", "164", "409", "500", "501", "502", "503", _
"504", "509", "510", "511", "512", "513", "514", "515", "516", "517", "518", "519", "520", "521", "522", "523", "524", "525", "526", "527", _
"528", "529", "530", "531", "532", "534", "535", "536", "537", "538", "539", "540", "541", "542", "543", "544", "547", "548", "550", "553", _
"557", "560", "562", "563", "564", "565", "566", "573", "584", "585", "586", "587", "588", "589", "592", "593", "594", "595", "596", "700", "701", _
"702", "703", "708", "709", "711", "712", "713", "714", "715", "716", "717", "718", "720", "721", "722", "723", "724", "729", "730", "731", _
"732", "734", "735", "738", "739", "740", "741", "742", "743", "744", "745", "749", "750", "751", "755", "756", "757", "758", "759", "761", _
"762", "763", "764", "772", "773", "774", "775", "776", "777", "778", "779", "780", "782", "783", "784", "785", "786", "787", "788", "789", _
"791", "792", "793", "794", "796", "797", "798", "799", "800", "801", "802", "803", "804", "805", "806", "807", "808", "809", "810")


    
    ' Inicializamos el resultado como Falso
    BuscarEnArray = False
    
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

' --- Helper robusto: limpia filas aunque tengan celdas combinadas
Private Sub LimpiarFilasPuente(ByVal hoja As String, ParamArray filas())
    Dim ws As Worksheet, k As Long, f As Long
    Dim rngRow As Range, c As Range
    Dim seen As Object
    Set ws = ThisWorkbook.Sheets(hoja)
    Set seen = CreateObject("Scripting.Dictionary")

    For k = LBound(filas) To UBound(filas)
        f = CLng(filas(k))

        ' Limita al bloque de la planilla (B:M). Ajusta si tu rango cambia.
        Set rngRow = ws.Range("B" & f & ":M" & f)

        On Error Resume Next
        rngRow.ClearContents                   ' intento rápido
        If Err.Number <> 0 Then                ' si hay combinadas que molestan...
            Err.Clear: On Error GoTo 0
            ' limpia celda a celda y, si está combinada, limpia toda el área combinada 1 sola vez
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
    
    ' ======= LIMPIEZA DE FILAS PUENTE =======
    ' Ventas: limpiar 3 filas antes de cada bloque anual (quedan con títulos)
    If NumeroHoja = 2 Then
        Dim stepRows As Long, baseRow As Long, k As Long, r As Long
        stepRows = ws.Range(Rango).Rows.Count + NumberoffsetRows   ' = 16 + 3 = 19
        baseRow = 20                                              ' primera “puente” conocida

        ' Limpia la primera “puente” del primer bloque
        LimpiarFilasPuente NombreHoja, baseRow

        ' Limpia 37–39, 56–58, 75–77, ... (base + step*k - 2 .. base + step*k)
        For k = 1 To yearDiff
            r = baseRow + (stepRows * k) - 2
            LimpiarFilasPuente NombreHoja, r, r + 1, r + 2
        Next k
    End If

    ' Compras: limpiar 3 filas “puente” en cada salto de año
    If NumeroHoja = 1 Then
        Dim stepRowsC As Long, baseRowC As Long, kc As Long, rc As Long
        stepRowsC = ws.Range(Rango).Rows.Count + NumberoffsetRows   ' 16 + 3 = 19
        baseRowC = 21                                               ' primer “puente” en Compras

        ' Limpia las primeras filas puente del primer bloque (si arrastran algo)
        LimpiarFilasPuente NombreHoja, 20, 21

        ' Limpia 38–40, 57–59, 76–78, ... (base + step*k - 2 .. base + step*k)
        For kc = 1 To yearDiff
            rc = baseRowC + (stepRowsC * kc) - 2
            LimpiarFilasPuente NombreHoja, rc, rc + 1, rc + 2
        Next kc
    End If
    
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
            Dim baseYearC As Long: baseYearC = ThisWorkbook.Sheets("Compras").Range("G1").value
            Dim yC As Long: yC = baseYearC + columnOffset
            Dim r528 As Long: r528 = FindRowByYearAndCode(yC, "528")
            Dim r537 As Long: r537 = FindRowByYearAndCode(yC, "537")
            Dim r504 As Long: r504 = FindRowByYearAndCode(yC, "504")
            Dim r562 As Long: r562 = FindRowByYearAndCode(yC, "562") ' S/L° Compras (sin derecho a CF)

            If r528 > 0 Then
                formulaString01 = "=DESREF('F29'!$C$" & r528 & ";0;FILA()-" & rowOffset & ")"
            Else
                formulaString01 = "0"
            End If
            If r537 > 0 Then
                formulaString02 = "=DESREF('F29'!$C$" & r537 & ";0;FILA()-" & rowOffset & ")"
            Else
                formulaString02 = "0"
            End If
            If r504 > 0 Then
                formulaString04 = "=DESREF('F29'!$C$" & r504 & ";0;FILA()-" & rowOffset & ")"
            Else
                formulaString04 = "0"
            End If
            If r562 > 0 Then
                formulaString05 = "=DESREF('F29'!$C$" & r562 & ";0;FILA()-" & rowOffset & ")"
            Else
                formulaString05 = "0"
            End If

        '========================
        ' VENTAS
        '========================
        Case 2
            ' Filas de inicio del bloque de meses en "Ventas"
            startRow = (columnOffset * 19) + 5
            rowOffset = startRow

            ' Año que corresponde a este bloque replicado
            Dim baseYearV As Long: baseYearV = ThisWorkbook.Sheets("Ventas").Range("G1").value
            Dim y As Long: y = baseYearV + columnOffset

            ' Códigos a leer EN F29 (SOLO LECTURA)
            Dim r538 As Long: r538 = FindRowByYearAndCode(y, "538")
            Dim r142 As Long: r142 = FindRowByYearAndCode(y, "142")
            Dim r020 As Long: r020 = FindRowByYearAndCode(y, "020")   ' <<--- Exportaciones

            ' Construcción de fórmulas
            If r538 > 0 Then
                formulaString01 = "=DESREF('F29'!$C$" & r538 & ";0;FILA()-" & rowOffset & ")"
            Else
                formulaString01 = "0"
            End If

            If r142 > 0 Then
                formulaString02 = "=DESREF('F29'!$C$" & r142 & ";0;FILA()-" & rowOffset & ")"
            Else
                formulaString02 = "0"
            End If

            ' *** CORREGIDO: Columna J ahora trae Cod 020 (no 715) ***
            If r020 > 0 Then
                formulaString03 = "=DESREF('F29'!$C$" & r020 & ";0;FILA()-" & rowOffset & ")"
            Else
                formulaString03 = "0"
            End If

        Case Else
            MsgBox "Número de hoja no válido.", vbExclamation
            Exit Sub
    End Select

    ' Volcado de 12 meses (Ene–Dic) al bloque correspondiente
    For i = startRow To startRow + 11
        If tipo = 1 Then
            wsDest.Cells(i, 3).FormulaLocal = formulaString01   ' Compras: Cod 528
            wsDest.Cells(i, 6).FormulaLocal = formulaString02   ' Compras: Cod 537
            wsDest.Cells(i, 7).FormulaLocal = formulaString04   ' Compras: Cod 504
            wsDest.Cells(i, 11).FormulaLocal = formulaString05  ' Compras: Exento (562)
        Else
            wsDest.Cells(i, 3).FormulaLocal = formulaString01   ' Ventas: Cod 538
            wsDest.Cells(i, 6).FormulaLocal = formulaString02   ' Ventas: Cod 142
            wsDest.Cells(i, 10).FormulaLocal = formulaString03  ' *** Ventas: Cod 020 (Exportaciones) ***
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
    Dim filepath As String
    Dim wsSummary As Worksheet
    Dim wsFiles As Worksheet
    Dim mes As Integer
    Dim arr() As Variant           ' <-- arreglo (con paréntesis)
    Dim filtros() As Variant       ' <-- arreglo (con paréntesis) para pasar a la función
    Dim pos As Integer
    Dim posArchivo As Integer
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
            ' S/L°Ventas (CLP):
            '   2023+ = (TipoDoc 33) - (TipoDoc 61)
            '   anteriores = 33,39,43,46,56,61
            If yearInput >= 2023 Then
            ' 33 + 56 - 61
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

            ' Ventas y/o Servicios exentos o no gravados
            arr = Array("33", "34", "41", "46", "56", "61")
            wsSummary.Cells(pos + mes, 8).value = SumaPorColumnaParametrica(filepath, mes, "Tipo Doc", arr, "Monto Exento", 10)

            ' Exportaciones / otros exentos
            arr = Array("110", "111", "112")
            wsSummary.Cells(pos + mes, 11).value = SumaPorColumnaParametrica(filepath, mes, "Tipo Doc", arr, "Monto Exento", 10)
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
                    Data = ProcesaDataSecuencial(Path)
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

Private Function FindRowByYearAndCode(ByVal year As Long, ByVal code As String) As Long
    Dim ws As Worksheet
    Dim startRow As Long, endRow As Long, r As Long
    Set ws = ThisWorkbook.Sheets("F29")
    FindRowByYearAndCode = 0
    Call LimitesBloqueF29(year, startRow, endRow)
    If startRow = 0 Or endRow = 0 Then Exit Function
    For r = startRow + 1 To endRow
        If Val(ws.Cells(r, "B").value) = Val(code) Then
            FindRowByYearAndCode = r
            Exit Function
        End If
    Next r
End Function





