'=============================================
'Motor de busqueda seccion esquemas
'=============================================

Sub MoverBloque(ref As String, BD As String, wsDestino As Worksheet, celdaDestino As String)
    Dim origen As Range
    Dim destino As Range
    Dim wsDatos As Worksheet

    Set wsDatos = ThisWorkbook.Sheets(BD)

    On Error GoTo errHandler
    Set origen = wsDatos.Range(ref)
    Set destino = wsDestino.Range(celdaDestino)

    origen.Copy Destination:=destino

    wsDestino.Range(destino, destino.Offset(origen.Rows.Count - 1, origen.Columns.Count - 1)).Select
    Exit Sub

errHandler:
    MsgBox "Referencia no encontrada ", vbCritical
End Sub
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
        Case "H1": celdaDestino = "CM45": rangoLimpiar = "CM45:DF104"
        Case "H2": celdaDestino = "DH45": rangoLimpiar = "DH45:EA104"
        Case "H3": celdaDestino = "EC45": rangoLimpiar = "EC45:EV104"
        Case "H4": celdaDestino = "EX45": rangoLimpiar = "EX45:FQ104"
        Case "H5": celdaDestino = "FS45": rangoLimpiar = "FS45:GL104"
        Case "H6": celdaDestino = "GN45": rangoLimpiar = "GN45:HG104"
        Case "H7": celdaDestino = "HI45": rangoLimpiar = "HI45:IB104"
        Case "H8": celdaDestino = "ID45": rangoLimpiar = "ID45:HT115"
        Case "H9": celdaDestino = "IY45": rangoLimpiar = "IY45:JR104"
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
        Case "H1": celdaDestino = "CM45": rangoLimpiar = "CM45:DF104"
        Case "H2": celdaDestino = "DH45": rangoLimpiar = "DH45:EA104"
        Case "H3": celdaDestino = "EC45": rangoLimpiar = "EC45:EV104"
        Case "H4": celdaDestino = "EX45": rangoLimpiar = "EX45:FQ104"
        Case "H5": celdaDestino = "FS45": rangoLimpiar = "FS45:GL104"
        Case "H6": celdaDestino = "GN45": rangoLimpiar = "GN45:HG104"
        Case "H7": celdaDestino = "HI45": rangoLimpiar = "HI45:IB104"
        Case "H8": celdaDestino = "ID45": rangoLimpiar = "ID45:HT115"
        Case "H9": celdaDestino = "IY45": rangoLimpiar = "IY45:JR104"
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
        Case "H1": celdaDestino = "CM45": rangoLimpiar = "CM45:DF104"
        Case "H2": celdaDestino = "DH45": rangoLimpiar = "DH45:EA104"
        Case "H3": celdaDestino = "EC45": rangoLimpiar = "EC45:EV104"
        Case "H4": celdaDestino = "EX45": rangoLimpiar = "EX45:FQ104"
        Case "H5": celdaDestino = "FS45": rangoLimpiar = "FS45:GL104"
        Case "H6": celdaDestino = "GN45": rangoLimpiar = "GN45:HG104"
        Case "H7": celdaDestino = "HI45": rangoLimpiar = "HI45:IB104"
        Case "H8": celdaDestino = "ID45": rangoLimpiar = "ID45:HT115"
        Case "H9": celdaDestino = "IY45": rangoLimpiar = "IY45:JR104"
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
Sub Celda_DM1A() ' Modo de ejecuci�n
    Dim ws As Worksheet
    Dim rng As Range
    Dim celdaDestino As String
    Dim rangoLimpiar As String
    Dim pregunta As String
    
    Set ws = ThisWorkbook.Sheets("ESQUEMA")
    
    ' ===============================
    ' Preguntar posicion de la celda
    ' ===============================
    pregunta = InputBox("Indicar la posicion de la celda (H1 a H9): ")
    
    If Trim(pregunta) = "" Then Exit Sub ' El usuario cancel� o dej� vac�o
    
    pregunta = UCase(Trim(pregunta))
    
    ' ===============================
    ' Determinar celda destino y rango a limpiar segun posicion
    ' ===============================
    Select Case pregunta
        Case "H1"
            celdaDestino = "CM45"
            rangoLimpiar = "CM45:DF104"
        Case "H2"
            celdaDestino = "DH45"
            rangoLimpiar = "DH45:EA104"
        Case "H3"
            celdaDestino = "EC45"
            rangoLimpiar = "EC45:EV104"
        Case "H4" 
            celdaDestino = "EX45"
            rangoLimpiar = "EX45:FQ104"
        Case "H5"
            celdaDestino = "FS45"
            rangoLimpiar = "FS45:GL104"
        Case "H6"
            celdaDestino = "GN45"
            rangoLimpiar = "GN45:HG104"
        Case "H7"
            celdaDestino = "HI45"
            rangoLimpiar = "HI45:IB104"
        Case "H8"
            celdaDestino = "ID45"
            rangoLimpiar = "ID45:HT115"
        Case "H9"
            celdaDestino = "IY45"
            rangoLimpiar = "IY45:JR104"
        Case Else
            MsgBox "posicion no valida. Debe indicar un valor entre H1 y H9.", vbExclamation
            Exit Sub
    End Select
    
    Set rng = ws.Range(rangoLimpiar)
    
    ' Limpiar celdas
    rng.Clear
    rng.Interior.Color = &HF2F2F2
    
    ' Borrar graficos y formas que se solapen con el rango
    EliminarGraficosEnRango ws, rng
    
    Dim codigoDM1A As String
    codigoDM1A = "ESQ_CL_DM1A_24"
    
    MoverBloque codigoDM1A, "BD_Esquemas", ws, celdaDestino
End Sub
Sub Limpiar_TodasLasPosiciones()
    Dim ws As Worksheet
    Dim rng As Range
    
    Set ws = ThisWorkbook.Sheets("ESQUEMA")
    Set rng = ws.Range("CM45:HV116")
    
    ' Limpiar solo el contenido (no formato)
    rng.ClearContents
    
    ' Borrar gr�ficos y formas que se solapen con el rango,
    ' EXCLUYENDO el bot�n "Boton_Limpiar" (ajusta el nombre real aqu�)
    EliminarGraficosEnRango ws, rng, "Boton_Limpiar"
    
    MsgBox "Se limpia el contenido y los graficos de todas las posiciones (H1 a H9).", vbInformation
End Sub

