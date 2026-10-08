'=================== HOJA ESQUEMA========
Sub Celda_IM() ' Modo de ejecucion
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
        Case "H1": celdaDestino = "CE57": rangoLimpiar = "CE57:CX116"
        Case "H2": celdaDestino = "CZ57": rangoLimpiar = "CZ57:DS116"
        Case "H3": celdaDestino = "DU57": rangoLimpiar = "DU57:EN116"
        Case "H4": celdaDestino = "EP57": rangoLimpiar = "EP57:FI116"
        Case "H5": celdaDestino = "FK57": rangoLimpiar = "FK57:GD116"
        Case "H6": celdaDestino = "GF57": rangoLimpiar = "FW56:GY116"
        Case "H7": celdaDestino = "HA57": rangoLimpiar = "GR56:HT116"
        Case "H8": celdaDestino = "HV57": rangoLimpiar = "HA57:IO116"
        Case "H9": celdaDestino = "IQ57": rangoLimpiar = "IQ57:JJ116"
        Case Else
            MsgBox "Posicion no valida. Debe indicar un valor entre H1 y H9.", vbExclamation
            Exit Sub
    End Select
    
    ' === NUEVO: guardar la posicion en la variable publica ===
    PosicionUnifilarActual = CLng(Mid(pregunta, 2, 1))
    Set rng = ws.Range(rangoLimpiar)
    
    ' Limpiar celdas
    rng.Clear
    rng.Interior.Color = &HF2F2F2
    
    ' Borrar gr?ficos y formas que se solapen con el rango
    EliminarGraficosEnRango ws, rng
    
    Dim codigoIM As String
    codigoIM = "ESQ_CL_IM_24"
    Resultado_IM = "1"
    
    MoverBloque codigoIM, "BD_Esquemas", ws, celdaDestino
End Sub
Sub Celda_QM() ' Modo de ejecuci?n
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
        Case "H1": celdaDestino = "CE57": rangoLimpiar = "CE57:CX116"
        Case "H2": celdaDestino = "CZ57": rangoLimpiar = "CZ57:DS116"
        Case "H3": celdaDestino = "DU57": rangoLimpiar = "DU57:EN116"
        Case "H4": celdaDestino = "EP57": rangoLimpiar = "EP57:FI116"
        Case "H5": celdaDestino = "FK57": rangoLimpiar = "FK57:GD116"
        Case "H6": celdaDestino = "FW56": rangoLimpiar = "FW56:GY116"
        Case "H7": celdaDestino = "GR56": rangoLimpiar = "GR56:HT116"
        Case "H8": celdaDestino = "HA57": rangoLimpiar = "HA57:IO116"
        Case "H9": celdaDestino = "IQ57": rangoLimpiar = "IQ57:JJ116"
        Case Else
            MsgBox "posicion no valida. Debe indicar un valor entre H1 y H9.", vbExclamation
            Exit Sub
    End Select
    ' === NUEVO: guardar la posicion en la variable publica ===
    PosicionUnifilarActual = CLng(Mid(pregunta, 2, 1))
    Set rng = ws.Range(rangoLimpiar)
    
    ' Limpiar celdas
    rng.Clear
    rng.Interior.Color = &HF2F2F2
    
    ' Borrar gr?ficos y formas que se solapen con el rango
    EliminarGraficosEnRango ws, rng
    
    Dim codigoQM As String
    codigoQM = "ESQ_CL_QM_24"
    Resultado_QM = "1"
    
    MoverBloque codigoQM, "BD_Esquemas", ws, celdaDestino
End Sub
Sub Celda_GBC() ' Modo de ejecuci?n
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
        Case "H1": celdaDestino = "CE57": rangoLimpiar = "CE57:CX116"
        Case "H2": celdaDestino = "CZ57": rangoLimpiar = "CZ57:DS116"
        Case "H3": celdaDestino = "DU57": rangoLimpiar = "DU57:EN116"
        Case "H4": celdaDestino = "EP57": rangoLimpiar = "EP57:FI116"
        Case "H5": celdaDestino = "FK57": rangoLimpiar = "FK57:GD116"
        Case "H6": celdaDestino = "FW56": rangoLimpiar = "FW56:GY116"
        Case "H7": celdaDestino = "GR56": rangoLimpiar = "GR56:HT116"
        Case "H8": celdaDestino = "HA57": rangoLimpiar = "HA57:IO116"
        Case "H9": celdaDestino = "IQ57": rangoLimpiar = "IQ57:JJ116"
        Case Else
            MsgBox "posicion no valida. Debe indicar un valor entre H1 y H9.", vbExclamation
            Exit Sub
    End Select
    ' === NUEVO: guardar la posicion en la variable publica ===
    PosicionUnifilarActual = CLng(Mid(pregunta, 2, 1))
    Set rng = ws.Range(rangoLimpiar)
    
    ' Limpiar celdas
    rng.Clear
    rng.Interior.Color = &HF2F2F2
    
    ' Borrar gr?ficos y formas que se solapen con el rango
    EliminarGraficosEnRango ws, rng
    
    Dim codigoGBC As String
    codigoGBC = "ESQ_CL_GBC_24"
    Resultado_GBC = "1"
    
    MoverBloque codigoGBC, "BD_Esquemas", ws, celdaDestino
End Sub
Sub Celda_GAM() ' Modo de ejecuci?n
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
        Case "H1": celdaDestino = "CE57": rangoLimpiar = "CE57:CX116"
        Case "H2": celdaDestino = "CZ57": rangoLimpiar = "CZ57:DS116"
        Case "H3": celdaDestino = "DU57": rangoLimpiar = "DU57:EN116"
        Case "H4": celdaDestino = "EP57": rangoLimpiar = "EP57:FI116"
        Case "H5": celdaDestino = "FK57": rangoLimpiar = "FK57:GD116"
        Case "H6": celdaDestino = "FW56": rangoLimpiar = "FW56:GY116"
        Case "H7": celdaDestino = "GR56": rangoLimpiar = "GR56:HT116"
        Case "H8": celdaDestino = "HA57": rangoLimpiar = "HA57:IO116"
        Case "H9": celdaDestino = "IQ57": rangoLimpiar = "IQ57:JJ116"
        Case Else
            MsgBox "posicion no valida. Debe indicar un valor entre H1 y H9.", vbExclamation
            Exit Sub
    End Select
    ' === NUEVO: guardar la posicion en la variable publica ===
    PosicionUnifilarActual = CLng(Mid(pregunta, 2, 1))
    Set rng = ws.Range(rangoLimpiar)
    
    ' Limpiar celdas
    rng.Clear
    rng.Interior.Color = &HF2F2F2
    
    ' Borrar gr?ficos y formas que se solapen con el rango
    EliminarGraficosEnRango ws, rng
    
    Dim codigoGBC As String
    codigoGBC = "ESQ_CL_GAM_24"
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
        Case "H1": celdaDestino = "CE57": rangoLimpiar = "CE57:CX116"
        Case "H2": celdaDestino = "CZ57": rangoLimpiar = "CZ57:DS116"
        Case "H3": celdaDestino = "DU57": rangoLimpiar = "DU57:EN116"
        Case "H4": celdaDestino = "EP57": rangoLimpiar = "EP57:FI116"
        Case "H5": celdaDestino = "FK57": rangoLimpiar = "FK57:GD116"
        Case "H6": celdaDestino = "FW56": rangoLimpiar = "FW56:GY116"
        Case "H7": celdaDestino = "GR56": rangoLimpiar = "GR56:HT116"
        Case "H8": celdaDestino = "HA57": rangoLimpiar = "HA57:IO116"
        Case "H9": celdaDestino = "IQ57": rangoLimpiar = "IQ57:JJ116"
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
    Set rng = ws.Range("CE57:HV116")
    
    ' Limpiar solo el contenido (no formato)
    rng.ClearContents
    
    ' Borrar gr?ficos y formas que se solapen con el rango,
    ' EXCLUYENDO el bot?n "Boton_Limpiar" (ajusta el nombre real aqu?)
    EliminarGraficosEnRango ws, rng, "Boton_Limpiar"
    
    MsgBox "Se limpia el contenido y los graficos de todas las posiciones (H1 a H9).", vbInformation
End Sub
