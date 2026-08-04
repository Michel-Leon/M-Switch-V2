'=========================================================
' CONFIGURACION GLOBAL (version final)
'=========================================================
Public PosicionUnifilarActual As Long
Public Const FILA_INICIO_UNIFILAR As Long = 44
Public Const ANCHO_SLOT As Long = 17
Public Const NUM_SLOTS_TOTAL As Long = 10
Public Const NUM_POSICIONES As Long = 9      ' H1 a H9 
Public Const NOMBRE_HOJA_CONTROL As String = "_CTRL_UNIFILAR"
Public Const FILA_CONTROL_INICIO As Long = 2
Sub MoverBloque(ref As String, BD As String, wsDestino As Worksheet, celdaDestino As String)
    Dim origen As Range
    Dim destino As Range
    Dim wsDatos As Worksheet

    Set wsDatos = ThisWorkbook.Sheets(BD)

    On Error GoTo errBuscarRango
    Set origen = wsDatos.Range(ref)
    Set destino = wsDestino.Range(celdaDestino)
    destino.Resize(origen.Rows.Count, origen.Columns.Count).UnMerge 
    On Error GoTo errCopiar

    origen.Copy Destination:=destino
    Exit Sub

errBuscarRango:
    MsgBox "Referencia no encontrada: " & ref & vbCrLf & _
           "(" & Err.Number & ") " & Err.Description, vbCritical
    Exit Sub

errCopiar:
    MsgBox "Error al copiar el bloque '" & ref & "' a " & wsDestino.Name & "!" & celdaDestino & vbCrLf & _
           "(" & Err.Number & ") " & Err.Description, vbCritical
End Sub
'=========================================================
' Ejecutar UNA sola vez para crear la hoja de control oculta
'=========================================================
Sub CrearHojaControl()
    Dim ws As Worksheet, hoja As Worksheet
    Dim existe As Boolean, i As Long
    
    existe = False
    For Each hoja In ThisWorkbook.Sheets
        If hoja.Name = NOMBRE_HOJA_CONTROL Then existe = True: Exit For
    Next hoja
    If existe Then
        MsgBox "La hoja de control ya existe.", vbInformation
        Exit Sub
    End If
    
    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = NOMBRE_HOJA_CONTROL
    ws.Range("A1").Value = "Posicion": ws.Range("B1").Value = "CodigoRef": ws.Range("C1").Value = "Ancho"
    For i = 1 To NUM_POSICIONES
        ws.Cells(FILA_CONTROL_INICIO + i - 1, "A").Value = i
    Next i
    ws.Visible = xlSheetVeryHidden
    MsgBox "Hoja de control creada.", vbInformation
End Sub

'=========================================================
' Asigna/sobreescribe el bloque de una posicion (H1-H9) y repinta TODO
'=========================================================
Sub Asignar_Posicion_Unifilar(posicion As Long, codigoRef As String, anchoBloque As Long)
    Dim wsCtrl As Worksheet
    Dim fila As Long
    
    If posicion < 1 Or posicion > NUM_POSICIONES Then
        MsgBox "Posicion no valida (1 a " & NUM_POSICIONES & ").", vbExclamation: Exit Sub
    End If
    If anchoBloque <> 17 And anchoBloque <> 34 Then
        MsgBox "Ancho de bloque no valido. Debe ser 17 o 34.", vbExclamation: Exit Sub
    End If
    
    Set wsCtrl = ThisWorkbook.Sheets(NOMBRE_HOJA_CONTROL)
    fila = FILA_CONTROL_INICIO + posicion - 1
    
    wsCtrl.Cells(fila, "B").Value = codigoRef   ' sobreescribe si ya habia algo en esa posicion
    wsCtrl.Cells(fila, "C").Value = anchoBloque
    
    RecalcularYColocarTodo
End Sub

Sub Limpiar_Posicion_Unifilar(posicion As Long)
    Dim wsCtrl As Worksheet
    Set wsCtrl = ThisWorkbook.Sheets(NOMBRE_HOJA_CONTROL)
    wsCtrl.Cells(FILA_CONTROL_INICIO + posicion - 1, "B").ClearContents
    wsCtrl.Cells(FILA_CONTROL_INICIO + posicion - 1, "C").ClearContents
    RecalcularYColocarTodo
End Sub

'=========================================================
' Limpia TODO el area y vuelve a pegar cada posicion en orden,
' acumulando el ancho real (asi H2 despues de un bloque de 34 en H1
' termina en el slot 3 automaticamente)
'=========================================================
Sub RecalcularYColocarTodo()
    Dim ws As Worksheet, wsCtrl As Worksheet, rng As Range
    Dim i As Long, fila As Long
    Dim codigoRef As String, anchoBloque As Long
    Dim slotsUsados As Long, slotActual As Long, colInicio As Long
    Dim celdaDestino As String
    
    Set ws = ThisWorkbook.Sheets("UNIFILAR")
    Set wsCtrl = ThisWorkbook.Sheets(NOMBRE_HOJA_CONTROL)
    Set rng = ws.Range("CD44:IU147")
    
    Application.ScreenUpdating = False
    rng.ClearContents
    rng.UnMerge 
    EliminarGraficosEnRango ws, rng, "Boton_Limpiar"
    
    slotActual = 1
    For i = 1 To NUM_POSICIONES
        fila = FILA_CONTROL_INICIO + i - 1
        codigoRef = wsCtrl.Cells(fila, "B").Value
        
        If codigoRef <> "" Then
            anchoBloque = CLng(wsCtrl.Cells(fila, "C").Value)
            slotsUsados = anchoBloque \ ANCHO_SLOT
            
            If slotActual + slotsUsados - 1 > NUM_SLOTS_TOTAL Then
                MsgBox "La posicion H" & i & " no cabe (limite " & NUM_SLOTS_TOTAL & " slots).", vbExclamation
                Application.ScreenUpdating = True
                Exit Sub
            End If
            
            colInicio = ws.Range("CD44").Column + (slotActual - 1) * ANCHO_SLOT
            celdaDestino = ws.Cells(FILA_INICIO_UNIFILAR, colInicio).Address(False, False)
            
            MoverBloque codigoRef, "Diagrama_unifilar", ws, celdaDestino
            slotActual = slotActual + slotsUsados
        End If
    Next i
    
    Application.ScreenUpdating = True
End Sub