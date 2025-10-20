Attribute VB_Name = "MD_Exportador"

Sub ExportarTodasLasHojasAPDF_Robusto()
    Dim ws As Worksheet, wsTemp As Worksheet
    Dim celdaTaxa As Range, celdaIdentif As Range, celda As Range
    Dim filaInicio As Long, filaFin As Long, colInicio As Long, colFin As Long
    Dim carpetaDestino As String, rngExportar As Range

    If ThisWorkbook.Path = "" Then
        MsgBox "Guarda primero el archivo Excel.", vbCritical
        Exit Sub
    End If

    carpetaDestino = ThisWorkbook.Path & "\PDFs_Iniciales\"
    If Dir(carpetaDestino, vbDirectory) = "" Then MkDir carpetaDestino

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    On Error Resume Next: Sheets("TEMPORAL_EXPORT").Delete: On Error GoTo 0
    Set wsTemp = Sheets.Add: wsTemp.Name = "TEMPORAL_EXPORT"

    For Each ws In ThisWorkbook.Worksheets
        If ws.Visible = xlSheetVisible And ws.Name <> wsTemp.Name Then

            Set celdaTaxa = Nothing
            For Each celda In ws.Columns("A").Cells
                If Trim(celda.Text) = "TAXA" Then
                    Set celdaTaxa = celda
                    Exit For
                End If
            Next celda
            If celdaTaxa Is Nothing Then GoTo SiguienteHoja

            filaInicio = celdaTaxa.Row
            colInicio = 1 ' Columna A

            ' Buscar IDENTIFICACION DE MUESTRAS en la fila de TAXA (puede estar combinada)
            Set celdaIdentif = Nothing
            Dim c As Range
            For Each c In ws.Rows(filaInicio).Cells
                Dim textoCelda As String
                textoCelda = Trim(c.Text)
                If Len(textoCelda) >= 12 Then
                    If Right(textoCelda, 12) = " DE MUESTRAS" Then
                        Set celdaIdentif = c
                        Exit For
                    End If
                End If
            Next c
            If celdaIdentif Is Nothing Then GoTo SiguienteHoja


            Dim colFinTemp As Long
            If celdaIdentif.MergeCells Then
                colFinTemp = celdaIdentif.MergeArea.Columns(celdaIdentif.MergeArea.Columns.Count).Column
            Else
                colFinTemp = celdaIdentif.Column
            End If
            colFin = colFinTemp

            ' Buscar Fin_Tabla en la columna A
            Dim celdaFinTabla As Range
            Set celdaFinTabla = ws.Columns("A").Find("Fin_Tabla", LookIn:=xlValues, LookAt:=xlPart)
            If celdaFinTabla Is Nothing Then GoTo SiguienteHoja

            filaFin = celdaFinTabla.Row

            ' Definir el rango a exportar: desde una fila antes de TAXA hasta Fin_Tabla, y columnas de A hasta el final de IDENTIFICACION DE MUESTRAS
            Set rngExportar = ws.Range(ws.Cells(filaInicio - 1, colInicio), ws.Cells(filaFin, colFin))

            Call ReducirTablaPorColorYExportar(ws, wsTemp, rngExportar, carpetaDestino, ws.Name, colInicio, colFin)

SiguienteHoja:
        End If
    Next ws

    wsTemp.Delete
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub


Sub ReducirTablaPorColorYExportar(ws As Worksheet, wsTemp As Worksheet, rngExportar As Range, carpetaDestino As String, nombreHoja As String, colInicio As Long, colFin As Long)
    Dim COLOR_FIJO As Long: COLOR_FIJO = RGB(255, 255, 0)
    Const MAX_FILAS As Long = 70

    Dim filaTaxa As Long
    filaTaxa = rngExportar.Cells(2, 1).Row
    Dim encabezado As Range
    Dim encabezadoInicio As Long, encabezadoFin As Long
    encabezadoInicio = filaTaxa - 1 ' siempre una fila antes de TAXA
    encabezadoFin = encabezadoInicio

    ' Buscar hacia abajo hasta la última fila consecutiva amarilla
    Do While ws.Cells(encabezadoFin + 1, colInicio).Interior.Color = COLOR_FIJO
        encabezadoFin = encabezadoFin + 1
    Loop

    ' Ahora sí definimos correctamente el encabezado completo
    Set encabezado = ws.Range(ws.Cells(encabezadoInicio, colInicio), ws.Cells(encabezadoFin, colFin))

    Dim filaStart As Long: filaStart = rngExportar.Row
    Dim filaEnd As Long: filaEnd = rngExportar.Row + rngExportar.Rows.Count - 1

    ' Listar todas las filas visibles y pintadas
    Dim filaBloques() As Long, colorBloques() As Long, esFijo() As Boolean
    Dim totalFilasExportar As Long: totalFilasExportar = 0
    Dim fila As Long
    For fila = filaStart To filaEnd
        If Not ws.Rows(fila).Hidden Then
            Dim colorFila As Long
            colorFila = ws.Cells(fila, 1).Interior.Color
            If colorFila <> 16777215 Then ' No blanco
                totalFilasExportar = totalFilasExportar + 1
                ReDim Preserve filaBloques(1 To totalFilasExportar)
                ReDim Preserve colorBloques(1 To totalFilasExportar)
                ReDim Preserve esFijo(1 To totalFilasExportar)
                filaBloques(totalFilasExportar) = fila
                colorBloques(totalFilasExportar) = colorFila
                esFijo(totalFilasExportar) = (colorFila = COLOR_FIJO)
            End If
        End If
    Next fila

    If totalFilasExportar = 0 Then
        MsgBox "No se encontraron filas pintadas y visibles para exportar.", vbInformation
        Exit Sub
    End If

    ' Agrupar bloques consecutivos de igual color y tipo
    Dim bloquesInicio() As Long, bloquesFin() As Long, bloquesColor() As Long, bloquesEsFijo() As Boolean
    Dim contador As Long: contador = 0
    Dim i As Long

    For i = 1 To totalFilasExportar
        If contador = 0 Then
            contador = 1
            ReDim Preserve bloquesInicio(1 To contador)
            ReDim Preserve bloquesFin(1 To contador)
            ReDim Preserve bloquesColor(1 To contador)
            ReDim Preserve bloquesEsFijo(1 To contador)
            bloquesInicio(contador) = i
            bloquesColor(contador) = colorBloques(i)
            bloquesEsFijo(contador) = esFijo(i)
        ElseIf colorBloques(i) <> bloquesColor(contador) Or esFijo(i) <> bloquesEsFijo(contador) Then
            contador = contador + 1
            ReDim Preserve bloquesInicio(1 To contador)
            ReDim Preserve bloquesFin(1 To contador)
            ReDim Preserve bloquesColor(1 To contador)
            ReDim Preserve bloquesEsFijo(1 To contador)
            bloquesInicio(contador) = i
            bloquesColor(contador) = colorBloques(i)
            bloquesEsFijo(contador) = esFijo(i)
        End If
        bloquesFin(contador) = i
    Next i

    ' Exportar bloques variables (no fijos, no amarillo)
    Dim idxSubgrupo As Long: idxSubgrupo = 1

    For i = 1 To contador
        If Not bloquesEsFijo(i) Then
            Dim parte As Long: parte = 1
            Dim totalFilasBloque As Long: totalFilasBloque = bloquesFin(i) - bloquesInicio(i) + 1

            Dim j As Long
            For j = 1 To totalFilasBloque Step MAX_FILAS
                Dim filasEnParte As Long
                filasEnParte = Application.Min(MAX_FILAS, totalFilasBloque - (j - 1))
                wsTemp.Cells.Clear

                ' === NUEVO: nombre base y archivo de decimales para esta parte ===
                Dim nombreBase As String
                nombreBase = carpetaDestino & nombreHoja & "_Subgrupo" & idxSubgrupo & "_Parte" & parte

                Dim decFile As Integer
                decFile = FreeFile
                Open nombreBase & "_decimales.txt" For Output As #decFile
                ' === FIN NUEVO ===


                ' Pega encabezado solo como valores/formato
                encabezado.Copy
                wsTemp.Cells(1, 1).PasteSpecial xlPasteValuesAndNumberFormats
                wsTemp.Cells(1, 1).PasteSpecial xlPasteFormats
                ' === NUEVO: decimales del encabezado (filas 1..encabezado.Rows.Count en wsTemp) ===
                Dim hdrRow As Long, origHdrRow As Long, colIdx As Long
                Dim relCol As Long, dec As Long
                Dim totalCols As Long
                totalCols = colFin - colInicio + 1

                For hdrRow = 1 To encabezado.Rows.Count
                    origHdrRow = encabezado.Row + (hdrRow - 1)  ' fila ORIGINAL en ws
                    For colIdx = colInicio To colFin
                        relCol = colIdx - colInicio + 1         ' columna relativa en wsTemp
                        If IsNumeric(ws.Cells(origHdrRow, colIdx).Value) Then
                            dec = CalcularDecimales(ws.Cells(origHdrRow, colIdx))
                            Print #decFile, "FILA=" & hdrRow & ";COL=" & relCol & ";DECIMALES=" & dec
                        Else
                            Print #decFile, "FILA=" & hdrRow & ";COL=" & relCol & ";DECIMALES="
                        End If
                    Next colIdx
                Next hdrRow
                ' === FIN NUEVO ===

                Application.CutCopyMode = False

                wsTemp.Range(wsTemp.Cells(1, 1), wsTemp.Cells(encabezado.Rows.Count, colFin - colInicio + 1)).Interior.ColorIndex = xlNone

                ' Copiar filas visibles y pintadas para esta parte
                Dim filaTemp As Long: filaTemp = encabezado.Rows.Count + 1
                Dim k As Long
                For k = bloquesInicio(i) + (j - 1) To Application.Min(bloquesInicio(i) + (j - 1) + filasEnParte - 1, bloquesFin(i))
                    Dim filaOriginal As Long: filaOriginal = filaBloques(k)

                    Call ForzarFormatoNumerico(ws.Range(ws.Cells(filaOriginal, colInicio), ws.Cells(filaOriginal, colFin)))

                    ws.Range(ws.Cells(filaOriginal, colInicio), ws.Cells(filaOriginal, colFin)).Copy
                    wsTemp.Cells(filaTemp, 1).PasteSpecial xlPasteValuesAndNumberFormats
                    wsTemp.Cells(filaTemp, 1).PasteSpecial xlPasteFormats
                    wsTemp.Range(wsTemp.Cells(filaTemp, 1), wsTemp.Cells(filaTemp, colFin - colInicio + 1)).Interior.ColorIndex = xlNone
                    ' === NUEVO: decimales de la fila de datos recién pegada ===
                    For colIdx = colInicio To colFin
                        relCol = colIdx - colInicio + 1
                        If IsNumeric(ws.Cells(filaOriginal, colIdx).Value) Then
                            dec = CalcularDecimales(ws.Cells(filaOriginal, colIdx))
                            Print #decFile, "FILA=" & filaTemp & ";COL=" & relCol & ";DECIMALES=" & dec
                        Else
                            Print #decFile, "FILA=" & filaTemp & ";COL=" & relCol & ";DECIMALES="
                        End If
                    Next colIdx
                    ' === FIN NUEVO ===

                    filaTemp = filaTemp + 1
                Next k

                Dim m As Long
                For m = 1 To contador
                    If bloquesEsFijo(m) Then
                        If filaBloques(bloquesInicio(m)) > filaBloques(bloquesFin(i)) Then
                            Dim l As Long
                            For l = bloquesInicio(m) To bloquesFin(m)
                                Dim filaFooter As Long: filaFooter = filaBloques(l)
                                ws.Range(ws.Cells(filaFooter, colInicio), ws.Cells(filaFooter, colFin)).Copy
                                wsTemp.Cells(filaTemp, 1).PasteSpecial xlPasteValuesAndNumberFormats
                                wsTemp.Cells(filaTemp, 1).PasteSpecial xlPasteFormats
                                wsTemp.Range(wsTemp.Cells(filaTemp, 1), wsTemp.Cells(filaTemp, colFin - colInicio + 1)).Interior.ColorIndex = xlNone
                                ' === NUEVO: decimales de la fila de footer recién pegada ===
                                For colIdx = colInicio To colFin
                                    relCol = colIdx - colInicio + 1
                                    If IsNumeric(ws.Cells(filaFooter, colIdx).Value) Then
                                        dec = CalcularDecimales(ws.Cells(filaFooter, colIdx))
                                        Print #decFile, "FILA=" & filaTemp & ";COL=" & relCol & ";DECIMALES=" & dec
                                    Else
                                        Print #decFile, "FILA=" & filaTemp & ";COL=" & relCol & ";DECIMALES="
                                    End If
                                Next colIdx
                                ' === FIN NUEVO ===

                                filaTemp = filaTemp + 1
                            Next l
                        End If
                    End If
                Next m
                Application.CutCopyMode = False

                ' Ajustes y exportar PDF
                wsTemp.Cells.WrapText = True
                Dim col As Range
                For Each col In wsTemp.UsedRange.Columns
                    If col.ColumnWidth < 20 Then col.ColumnWidth = col.ColumnWidth + 5
                Next col

                wsTemp.Cells.Font.Size = 9 ' Opcional

                With wsTemp.PageSetup
                    .Orientation = xlLandscape
                    .Zoom = False
                    .LeftMargin = Application.InchesToPoints(0.5)
                    .RightMargin = Application.InchesToPoints(0.5)
                    .TopMargin = Application.InchesToPoints(0.5)
                    .BottomMargin = Application.InchesToPoints(0.5)
                End With

                Dim nombreArchivo As String
                nombreArchivo = carpetaDestino & nombreHoja & "_Parte" & idxSubgrupo & "_" & (j \ MAX_FILAS + 1) & ".pdf"
                
                ' === NUEVO: cerrar decimales para asegurar que exista y esté completo ===
                Close #decFile
                ' === FIN NUEVO ===


                ' ==== Exportar el mapeo de celdas combinadas y formato de la tabla en wsTemp ====
                Call ExportarMapaCombinadasYFormato(wsTemp.UsedRange, nombreBase & "_mapa.txt", nombreBase & "_formato.txt")

                ' === ELIMINAR SOLO SEPARADORES DE MILES (manteniendo decimales y estilo) ===
                Dim celdaFmt As Range
                Dim fmtOriginal As String

                Application.UseSystemSeparators = False
                Application.ThousandsSeparator = ""
                Application.DecimalSeparator = "."

                For Each celdaFmt In wsTemp.UsedRange
                    If IsNumeric(celdaFmt.Value) And Not IsEmpty(celdaFmt.Value) Then
                        fmtOriginal = celdaFmt.NumberFormat
                        ' Si el formato contiene patrón de miles (coma, espacio o agrupación)
                        If InStr(fmtOriginal, "#,##") > 0 Or InStr(fmtOriginal, "# ##") > 0 Or InStr(fmtOriginal, " ") > 0 Then
                            ' Reemplaza solo la parte de agrupación, sin tocar los decimales
                            fmtOriginal = Replace(fmtOriginal, "#,##", "##")
                            fmtOriginal = Replace(fmtOriginal, "# ##", "##")
                            fmtOriginal = Replace(fmtOriginal, " ", "")  ' elimina espacios finos invisibles (U+202F)
                            celdaFmt.NumberFormat = fmtOriginal
                        End If
                    End If
                Next celdaFmt
                ' === FIN ELIMINAR SEPARADORES DE MILES ===





                ' ==== Exportar CSV usando wsTemp.Copy ====
                
                wsTemp.Copy
                ActiveWorkbook.SaveAs Filename:=nombreBase & ".csv", FileFormat:=xlCSV
                ActiveWorkbook.Close SaveChanges:=False


                parte = parte + 1
            Next j
            idxSubgrupo = idxSubgrupo + 1
        End If
    Next i

End Sub

Sub ExportarMapaCombinadasYFormato(rango As Range, archivo As String, archivo_formato As String)
    Dim mapFile As Integer, fmtFile As Integer
    mapFile = FreeFile: Open archivo For Output As #mapFile
    fmtFile = FreeFile: Open archivo_formato For Output As #fmtFile

    ' === NUEVO: cargar decimales desde el TXT asociado a esta parte ===
    Dim decPath As String
    decPath = Replace(archivo_formato, "_formato.txt", "_decimales.txt")

    Dim decDict As Object: Set decDict = CreateObject("Scripting.Dictionary")
    Dim ff As Integer, ln As String
    If Len(Dir(decPath)) > 0 Then
        ff = FreeFile
        Open decPath For Input As #ff
        Do While Not EOF(ff)
            Line Input #ff, ln
            If Len(Trim$(ln)) > 0 Then
                ' Espera: FILA=R;COL=C;DECIMALES=K (K puede estar vacío)
                Dim parts() As String, r As String, c As String, d As String, key As String
                parts = Split(ln, ";")
                If UBound(parts) >= 2 Then
                    r = Split(parts(0), "=")(1)
                    c = Split(parts(1), "=")(1)
                    If InStr(parts(2), "=") > 0 Then d = Split(parts(2), "=")(1) Else d = ""
                    key = CStr(Val(r)) & ":" & CStr(Val(c))
                    If Not decDict.Exists(key) Then decDict.Add key, d
                End If
            End If
        Loop
        Close #ff
    End If
    ' === FIN NUEVO ===

    Dim celda As Range
    Dim decimales As Variant

    For Each celda In rango
        ' Mapeo de celdas combinadas (solo ancla)
        If celda.MergeCells Then
            If celda.Address = celda.MergeArea.Cells(1, 1).Address Then
                Print #mapFile, "FILA=" & celda.Row & ";COL=" & celda.Column & _
                                ";ROWS=" & celda.MergeArea.Rows.Count & ";COLS=" & celda.MergeArea.Columns.Count
            End If
        End If

        ' === NUEVO: usar decimales del TXT si existen (coordenadas relativas en wsTemp: empiezan en 1,1) ===
        Dim k As String
        k = CStr(celda.Row) & ":" & CStr(celda.Column)

        If decDict.Exists(k) Then
            decimales = decDict(k)  ' puede ser "" (no numérico) o un número
        Else
            ' Fallback: método anterior por si no hay TXT (compatibilidad)
            If IsNumeric(celda.Value) Then
                If InStr(1, celda.Text, Application.DecimalSeparator) > 0 Then
                    decimales = Len(Split(celda.Text, Application.DecimalSeparator)(1))
                Else
                    decimales = 0
                End If
            Else
                decimales = ""
            End If
        End If
        ' === FIN NUEVO ===

        ' Guardar formato + decimales
        Print #fmtFile, "FILA=" & celda.Row & ";COL=" & celda.Column & _
                        ";ITALIC=" & celda.Font.Italic & _
                        ";BOLD=" & celda.Font.Bold & _
                        ";FMT=" & celda.NumberFormat & _
                        ";DECIMALES=" & decimales
    Next celda

    Close #mapFile
    Close #fmtFile
End Sub


Function CalcularDecimales(c As Range) As Long
    On Error GoTo Fin
    Dim fmt As String, secPos As String
    Dim i As Long, inQuote As Boolean, ch As String
    Dim cleaned As String, pDot As Long, pComma As Long, decPos As Long
    Dim count0 As Long

    fmt = CStr(c.NumberFormatLocal)
    secPos = Split(fmt, ";")(0)

    ' Quitar literales entre comillas
    cleaned = ""
    For i = 1 To Len(secPos)
        ch = Mid$(secPos, i, 1)
        If ch = """" Then
            inQuote = Not inQuote
        ElseIf Not inQuote Then
            cleaned = cleaned & ch
        End If
    Next i

    ' Buscar separador decimal
    pDot = InStr(cleaned, ".")
    pComma = InStr(cleaned, ",")
    If pDot = 0 And pComma = 0 Then GoTo ContarPorValor

    ' Contar ceros a la derecha del separador
    decPos = IIf(pDot > 0, pDot, pComma)
    For i = decPos + 1 To Len(cleaned)
        ch = Mid$(cleaned, i, 1)
        If ch = "0" Then
            count0 = count0 + 1
        ElseIf ch Like "[#,?]" Then
            ' Ignorar opcionales, se manejarán por valor
        ElseIf ch = "%" Or ch = "E" Or ch = "e" Then
            Exit For
        End If
    Next i

    ' Si el formato solo tenía '#' (ningún 0), contar por valor
    If count0 = 0 Then GoTo ContarPorValor
    CalcularDecimales = count0
    Exit Function

ContarPorValor:
    ' --- Nuevo bloque: si formato usa #, usar el valor real ---
    If IsNumeric(c.Value) Then
        Dim s As String
        s = CStr(c.Value)
        If InStr(s, ".") > 0 Then
            CalcularDecimales = Len(Split(s, ".")(1))
        Else
            CalcularDecimales = 0
        End If
    Else
        CalcularDecimales = 0
    End If
    Exit Function

Fin:
    CalcularDecimales = 0
End Function


Sub ForzarFormatoNumerico(rng As Range)
    Dim c As Range
    For Each c In rng
        If IsNumeric(c.Value) And c.NumberFormatLocal = "General" Then
            c.NumberFormatLocal = "0.##########"
        End If
    Next c
End Sub












