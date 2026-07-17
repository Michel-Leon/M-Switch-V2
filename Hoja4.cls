'=============================================
'Motor de busqueda seccion esquemas
'=============================================

Sub MoverBloque(ref As String, BD As String)
    Dim origen As Range
    Dim destino As Range
    Dim wsDatos As Worksheet
    Dim wsActual As Worksheet

    Set wsDatos = ThisWorkbook.Sheets(BD)
    Set wsActual = ActiveSheet

    On Error GoTo errHandler
    Set origen = wsDatos.Range(ref)
    Set destino = wsActual.Range("JR30")

    origen.Copy Destination:=destino

    wsActual.Range(destino, destino.Offset(origen.Rows.Count - 1, origen.Columns.Count - 1)).Select
    Exit Sub

errHandler:
        MsgBox "Referencia no encontrada ", vbCritical
End Sub
Sub EliminarGraficosEnRango(ws As Worksheet, rng As Range)
    Dim shp As Shape
    Dim i As Long
    
    ' Recorremos de atrás hacia adelante porque al eliminar shapes
    ' la colección se reindexa y si vas hacia adelante te saltas elementos
    For i = ws.Shapes.Count To 1 Step -1
        Set shp = ws.Shapes(i)
        If ShapeIntersectaConRango(shp, rng) Then
            shp.Delete
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
    
    ' Verifica si hay solapamiento entre los dos rectángulos
    If shpLeft < rngRight And shpRight > rngLeft And _
       shpTop < rngBottom And shpBottom > rngTop Then
        ShapeIntersectaConRango = True
    Else
        ShapeIntersectaConRango = False
    End If
End Function

Sub Celda_IM() ' Modo de ejecución
    Dim ws As Worksheet
    Dim rng As Range
    
    Set ws = ThisWorkbook.Sheets("ESQUEMA")
    Set rng = ws.Range("JJ22:LG166")
    
    ' Limpiar celdas
    rng.Clear
    rng.Interior.Color = &HF2F2F2
    
    ' Borrar gráficos y formas que se solapen con el rango
    EliminarGraficosEnRango ws, rng
    
    Dim codigoIM As String
    codigoIM = "ESQ_CL_IM_24"
    Resultado_IM = "1"
    MoverBloque codigoIM, "BD_Esquemas"
End Sub
Sub Celda_QM() ' Modo de ejecución
    Dim ws As Worksheet
    Dim rng As Range
    
    Set ws = ThisWorkbook.Sheets("ESQUEMA")
    Set rng = ws.Range("JJ22:LG166")
    
    ' Limpiar celdas
    rng.Clear
    rng.Interior.Color = &HF2F2F2
    
    ' Borrar gráficos y formas que se solapen con el rango
    EliminarGraficosEnRango ws, rng
    
    Dim codigoQM As String
    codigoQM = "ESQ_CL_QM_24"
    Resultado_QM = "1"
    MoverBloque codigoQM, "BD_Esquemas"
End Sub
Sub Celda_GBC() ' Modo de ejecución
    Dim ws As Worksheet
    Dim rng As Range
    
    Set ws = ThisWorkbook.Sheets("ESQUEMA")
    Set rng = ws.Range("JJ22:LG166")
    
    ' Limpiar celdas
    rng.Clear
    rng.Interior.Color = &HF2F2F2
    
    ' Borrar gráficos y formas que se solapen con el rango
    EliminarGraficosEnRango ws, rng
    
    Dim codigoGBC As String
    codigoGBC = "ESQ_CL_GBC_24"
    Resultado_GBC = "1"
    MoverBloque codigoGBC, "BD_Esquemas"
End Sub
Sub Celda_DM1A() ' Modo de ejecución
    Dim ws As Worksheet
    Dim rng As Range
    
    Set ws = ThisWorkbook.Sheets("ESQUEMA")
    Set rng = ws.Range("JJ22:LG166")
    
    ' Limpiar celdas
    rng.Clear
    rng.Interior.Color = &HF2F2F2
    
    ' Borrar gráficos y formas que se solapen con el rango
    EliminarGraficosEnRango ws, rng
    
    Dim codigoDM1A As String
    codigoDM1A = "ESQ_CL_DM1A_24"
    Resultado_DM1A = "1"
    MoverBloque codigoDM1A, "BD_Esquemas"
End Sub


