'=============================================
'Motor de busqueda seccion esquemas
'=============================================


Sub EliminarGraficosEnRango(ws As Worksheet, rng As Range, Optional excluir As String = "")
    Dim shp As Shape
    Dim i As Long
    
    ' Recorremos de atr�s hacia adelante porque al eliminar shapes
    ' la colecci�n se reindexa y si vas hacia adelante te saltas elementos
    For i = ws.Shapes.Count To 1 Step -1
        Set shp = ws.Shapes(i)
        If shp.Name <> excluir Then
            If ShapeIntersectaConRango(shp, rng) Then
                shp.Delete
            End If
        End If
    Next i
End Sub
Function ShapeIntersectaConRango(shp As Shape, rng As Range) As Boolean
    Dim shpLeft As Double, shpTop As Double
    Dim shpRight As Double, shpBottom As Double
    Dim rngLeft As Double, rngTop As Double
    Dim rngRight As Double, rngBottom As Double
    
    shpLeft = shp.Left
    shpTop = shp.Top
    shpRight = shp.Left + shp.Width
    shpBottom = shp.Top + shp.Height
    
    rngLeft = rng.Left
    rngTop = rng.Top
    rngRight = rng.Left + rng.Width
    rngBottom = rng.Top + rng.Height
    
    ' Verifica si hay solapamiento entre los dos rect�ngulos
    If shpLeft < rngRight And shpRight > rngLeft And _
       shpTop < rngBottom And shpBottom > rngTop Then
        ShapeIntersectaConRango = True
    Else
        ShapeIntersectaConRango = False
    End If
End Function
Sub Celda_IM() ' Modo de ejecuci�n
    Dim ws As Worksheet
    Dim rng As Range
    Dim celdaDestino As String
    Dim rangoLimpiar As String
    Dim pregunta As String
    
    Set ws = ThisWorkbook.Sheets("ESQUEMA")
    
    pregunta = InputBox("Indicar la posicion de la celda (H1 a H9): ")
    If Trim(pregunta) = "" Then Exit Sub
    pregunta = UCase(Trim(pregunta))
    
    Select Case pregunta 
        Case "H1": celdaDestino = "BV56": rangoLimpiar = "BV56:CO115"
        Case "H2": celdaDestino = "CQ56": rangoLimpiar = "CQ56:DJ115"
        Case "H3": celdaDestino = "DL56": rangoLimpiar = "DL56:EE115"
        Case "H4": celdaDestino = "EG56": rangoLimpiar = "EG56:EZ115"
        Case "H5": celdaDestino = "FB56": rangoLimpiar = "FB56:FU115"
        Case "H6": celdaDestino = "FW56": rangoLimpiar = "FW56:GP115"
        Case "H7": celdaDestino = "GR56": rangoLimpiar = "GR56:HK115"
        Case "H8": celdaDestino = "HM56": rangoLimpiar = "HM56:IF115"
        Case "H9": celdaDestino = "IH56": rangoLimpiar = "IH56:JA115"
        Case Else
            MsgBox "Posicion no valida. Debe indicar un valor entre H1 y H9.", vbExclamation
            Exit Sub
    End Select
    
    Set rng = ws.Range(rangoLimpiar)
    
    ' Limpiar celdas
    rng.Clear
    rng.Interior.Color = &HF2F2F2
    
    ' Borrar gr�ficos y formas que se solapen con el rango
    EliminarGraficosEnRango ws, rng
    
    Dim codigoIM As String
    codigoIM = "ESQ_CL_IM_24"
    Resultado_IM = "1"
    
    MoverBloque codigoIM, "BD_Esquemas", ws, celdaDestino
End Sub
Sub Celda_QM() ' Modo de ejecuci�n
    Dim ws As Worksheet
    Dim rng As Range
    Dim celdaDestino As String
    Dim rangoLimpiar As String
    Dim pregunta As String
    
    Set ws = ThisWorkbook.Sheets("ESQUEMA")
    
    pregunta = InputBox("Indicar la posicion de la celda (H1 a H9): ")
    If Trim(pregunta) = "" Then Exit Sub
    pregunta = UCase(Trim(pregunta))
    
    Select Case pregunta
        Case "H1": celdaDestino = "BV56": rangoLimpiar = "BV56:CO115"
        Case "H2": celdaDestino = "CQ56": rangoLimpiar = "CQ56:DJ115"
        Case "H3": celdaDestino = "DL56": rangoLimpiar = "DL56:EE115"
        Case "H4": celdaDestino = "EG56": rangoLimpiar = "EG56:EZ115"
        Case "H5": celdaDestino = "FB56": rangoLimpiar = "FB56:FU115"
        Case "H6": celdaDestino = "FW56": rangoLimpiar = "FW56:GP115"
        Case "H7": celdaDestino = "GR56": rangoLimpiar = "GR56:HK115"
        Case "H8": celdaDestino = "HM56": rangoLimpiar = "HM56:IF115"
        Case "H9": celdaDestino = "IH56": rangoLimpiar = "IH56:JA115"
        Case Else
            MsgBox "posicion no valida. Debe indicar un valor entre H1 y H9.", vbExclamation
            Exit Sub
    End Select
    
    Set rng = ws.Range(rangoLimpiar)
    
    ' Limpiar celdas
    rng.Clear
    rng.Interior.Color = &HF2F2F2
    
    ' Borrar gr�ficos y formas que se solapen con el rango
    EliminarGraficosEnRango ws, rng
    
    Dim codigoQM As String
    codigoQM = "ESQ_CL_QM_24"
    Resultado_QM = "1"
    
    MoverBloque codigoQM, "BD_Esquemas", ws, celdaDestino
End Sub
Sub Celda_GBC() ' Modo de ejecuci�n
    Dim ws As Worksheet
    Dim rng As Range
    Dim celdaDestino As String
    Dim rangoLimpiar As String
    Dim pregunta As String
    
    Set ws = ThisWorkbook.Sheets("ESQUEMA")
    
    pregunta = InputBox("Indicar la posicion de la celda (H1 a H9): ")
    If Trim(pregunta) = "" Then Exit Sub
    pregunta = UCase(Trim(pregunta))
    
    Select Case pregunta
        Case "H1": celdaDestino = "BV56": rangoLimpiar = "BV56:CO115"
        Case "H2": celdaDestino = "CQ56": rangoLimpiar = "CQ56:DJ115"
        Case "H3": celdaDestino = "DL56": rangoLimpiar = "DL56:EE115"
        Case "H4": celdaDestino = "EG56": rangoLimpiar = "EG56:EZ115"
        Case "H5": celdaDestino = "FB56": rangoLimpiar = "FB56:FU115"
        Case "H6": celdaDestino = "FW56": rangoLimpiar = "FW56:GP115"
        Case "H7": celdaDestino = "GR56": rangoLimpiar = "GR56:HK115"
        Case "H8": celdaDestino = "HM56": rangoLimpiar = "HM56:IF115"
        Case "H9": celdaDestino = "IH56": rangoLimpiar = "IH56:JA115"
        Case Else
            MsgBox "posicion no valida. Debe indicar un valor entre H1 y H9.", vbExclamation
            Exit Sub
    End Select
    
    Set rng = ws.Range(rangoLimpiar)
    
    ' Limpiar celdas
    rng.Clear
    rng.Interior.Color = &HF2F2F2
    
    ' Borrar gr�ficos y formas que se solapen con el rango
    EliminarGraficosEnRango ws, rng
    
    Dim codigoGBC As String
    codigoGBC = "ESQ_CL_GBC_24"
    Resultado_GBC = "1"
    
    MoverBloque codigoGBC, "BD_Esquemas", ws, celdaDestino
End Sub
Sub Celda_DM1A() ' Modo de ejecucion
    Dim ws As Worksheet
    Dim rng As Range
    Dim celdaDestino As String
    Dim rangoLimpiar As String
    Dim pregunta As String
    
    Set ws = ThisWorkbook.Sheets("ESQUEMA")
    
    pregunta = InputBox("Indicar la posicion de la celda (H1 a H9): ")
    If Trim(pregunta) = "" Then Exit Sub
    pregunta = UCase(Trim(pregunta))
    
    Select Case pregunta
        Case "H1": celdaDestino = "BV56": rangoLimpiar = "BV56:CO115"
        Case "H2": celdaDestino = "CQ56": rangoLimpiar = "CQ56:DJ115"
        Case "H3": celdaDestino = "DL56": rangoLimpiar = "DL56:EE115"
        Case "H4": celdaDestino = "EG56": rangoLimpiar = "EG56:EZ115"
        Case "H5": celdaDestino = "FB56": rangoLimpiar = "FB56:FU115"
        Case "H6": celdaDestino = "FW56": rangoLimpiar = "FW56:GP115"
        Case "H7": celdaDestino = "GR56": rangoLimpiar = "GR56:HK115"
        Case "H8": celdaDestino = "HM56": rangoLimpiar = "HM56:IF115"
        Case "H9": celdaDestino = "IH56": rangoLimpiar = "IH56:JA115"
        Case Else
            MsgBox "posicion no valida. Debe indicar un valor entre H1 y H9.", vbExclamation
            Exit Sub
    End Select
    
    ' === NUEVO: guardar la posicion en la variable publica ===
    PosicionUnifilarActual = CLng(Mid(pregunta, 2, 1))
    
    Set rng = ws.Range(rangoLimpiar)
    rng.Clear
    rng.Interior.Color = &HF2F2F2
    EliminarGraficosEnRango ws, rng
    
    Dim codigoDM1A As String
    codigoDM1A = "ESQ_CL_DM1A_24"
    
    MoverBloque codigoDM1A, "BD_Esquemas", ws, celdaDestino
End Sub
Sub Limpiar_TodasLasPosiciones()
    Dim ws As Worksheet
    Dim rng As Range
    
    Set ws = ThisWorkbook.Sheets("ESQUEMA")
    Set rng = ws.Range("BV56:HV116")
    
    ' Limpiar solo el contenido (no formato)
    rng.ClearContents
    
    ' Borrar gr�ficos y formas que se solapen con el rango,
    ' EXCLUYENDO el bot�n "Boton_Limpiar" (ajusta el nombre real aqu�)
    EliminarGraficosEnRango ws, rng, "Boton_Limpiar"
    
    MsgBox "Se limpia el contenido y los graficos de todas las posiciones (H1 a H9).", vbInformation
End Sub

