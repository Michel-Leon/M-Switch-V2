Attribute VB_Name = "M�dulo1"
Public Sub Limpiar_Layout() 
    ThisWorkbook.Sheets("Layout").Range("JI96:LI187").Clear
    ThisWorkbook.Sheets("Layout").Range("JI96:LI187").Interior.Color = &HF2F2F2
End Sub 
Public Sub Limpiar_Unifilar()
    ThisWorkbook.Sheets("Unifilar").Range("JI96:LI187").Clear
    ThisWorkbook.Sheets("Unifilar").Range("JI96:LI187").Interior.Color = &HF2F2F2
End Sub
Public Sub Restablecer_Layout()
    ThisWorkbook.Sheets("Layout").Range("BR44:IE147").Clear
    ThisWorkbook.Sheets("Layout").Range("BR148:IE165").Clear
    call DibujarLineaPunteada_Fija
End Sub
Public Sub Restablecer_unifilar()
    ThisWorkbook.Sheets("Unifilar").Range("BR44:IE147").Clear
    ThisWorkbook.Sheets("Unifilar").Range("BR148:IE165").Clear
    call DibujarLineaPunteada_Fijau
    DibujarLineaPunteada_FijaVU
End Sub
Sub DibujarLineaPunteada_Fija() 
    Dim ws As Worksheet
    Dim rng As Range

    ' Define la hoja donde aplicar la línea (puedes cambiar "Hoja1" por el nombre de tu hoja)
    Set ws = ThisWorkbook.Sheets("Layout")
    
    ' Define el rango fijo
    Set rng = ws.Range("BR148:IE148")
    
    ' Limpia el borde inferior anterior (opcional)
    rng.Borders(xlEdgeTop).LineStyle = xlNone

    ' Aplica línea punteada fina color #BFBFBF
    With rng.Borders(xlEdgeTop)
        .LineStyle = xlDot
        .Weight = xlHairline
        .Color = RGB(191, 191, 191)
    End With
End Sub
Sub DibujarLineaPunteada_Fijau()
    Dim ws As Worksheet
    Dim rng As Range

    ' Define la hoja donde aplicar la línea (puedes cambiar "Hoja1" por el nombre de tu hoja)
    Set ws = ThisWorkbook.Sheets("unifilar")
    
    ' Define el rango fijo
    Set rng = ws.Range("BR148:IE148")
    
    ' Limpia el borde inferior anterior (opcional)
    rng.Borders(xlEdgeTop).LineStyle = xlNone

    ' Aplica línea punteada fina color #BFBFBF
    With rng.Borders(xlEdgeTop)
        .LineStyle = xlDot
        .Weight = xlHairline
        .Color = RGB(191, 191, 191)
    End With  
End Sub
Sub DibujarLineaPunteada_FijaVU() 
    Dim ws As Worksheet
    Dim rng As Range

    ' Define la hoja donde aplicar la línea (puedes cambiar "Hoja1" por el nombre de tu hoja)
    Set ws = ThisWorkbook.Sheets("unifilar")
    
    ' Define el rango fijo
    Set rng = ws.Range("BR44:BR141")
    
    ' Limpia el borde inferior anterior (opcional)
    rng.Borders(xlEdgeLeft).LineStyle = xlNone

    ' Aplica línea punteada fina color #BFBFBF
    With rng.Borders(xlEdgeLeft)
        .LineStyle = xlDot
        .Weight = xlHairline
        .Color = RGB(191, 191, 191)
    End With
End Sub



Sub seleccion_IM()
    Dim shp As Shape
    Dim groupname As String
    groupname = "Celda IM"
    
    On Error Resume Next
    Set shp = ActiveSheet.Shapes(groupname)
    On Error GoTo 0
    
    If Not shp Is Nothing Then
        ' Alterna la visibilidad del grupo seleccionado
        shp.Visible = Not shp.Visible
        ' Oculta los dem s grupos
        ActiveSheet.Shapes("Celda QM").Visible = False
        ActiveSheet.Shapes("Celda GAM").Visible = False
        ActiveSheet.Shapes("Celda DM1").Visible = False
        ActiveSheet.Shapes("Celda GBC").Visible = False
        ActiveSheet.Shapes("Seleccion COTA").Visible = False
    end If
end Sub
Sub seleccion_QM()
    Dim shp As Shape
    Dim groupname As String
    groupname = "Celda QM"
    
    On Error Resume Next
    Set shp = ActiveSheet.Shapes(groupname)
    On Error GoTo 0
    
    If Not shp Is Nothing Then
        ' Alterna la visibilidad del grupo seleccionado
        shp.Visible = Not shp.Visible
        ' Oculta los dem s grupos
        ActiveSheet.Shapes("Celda IM").Visible = False
        ActiveSheet.Shapes("Celda GAM").Visible = False
        ActiveSheet.Shapes("Celda DM1").Visible = False
        ActiveSheet.Shapes("Celda GBC").Visible = False
        ActiveSheet.Shapes("Seleccion COTA").Visible = False
    end If
end Sub
Sub seleccion_GAM()
    Dim shp As Shape
    Dim groupname As String
    groupname = "Celda GAM"
    
    On Error Resume Next
    Set shp = ActiveSheet.Shapes(groupname)
    On Error GoTo 0
    
    If Not shp Is Nothing Then
        ' Alterna la visibilidad del grupo seleccionado
        shp.Visible = Not shp.Visible
        ' Oculta los dem s grupos
        ActiveSheet.Shapes("Celda IM").Visible = False
        ActiveSheet.Shapes("Celda QM").Visible = False
        ActiveSheet.Shapes("Celda DM1").Visible = False
        ActiveSheet.Shapes("Celda GBC").Visible = False
        ACTIVEsheet.Shapes("Seleccion COTA").Visible = False

    end If
end Sub
Sub seleccion_DM1()
    Dim shp As Shape
    Dim groupname As String
    groupname = "Celda DM1"
    
    On Error Resume Next
    Set shp = ActiveSheet.Shapes(groupname)
    On Error GoTo 0
    
    If Not shp Is Nothing Then
        ' Alterna la visibilidad del grupo seleccionado
        shp.Visible = Not shp.Visible
        ' Oculta los dem s grupos
        ActiveSheet.Shapes("Celda IM").Visible = False
        ActiveSheet.Shapes("Celda QM").Visible = False
        ActiveSheet.Shapes("Celda GAM").Visible = False
        ActiveSheet.Shapes("Celda GBC").Visible = False
        ActiveSheet.Shapes("Seleccion COTA").Visible = False

    end If
end Sub
Sub seleccion_GBC()
    Dim shp As Shape
    Dim groupname As String
    groupname = "Celda GBC"
    
    On Error Resume Next
    Set shp = ActiveSheet.Shapes(groupname)
    On Error GoTo 0
    
    If Not shp Is Nothing Then
        ' Alterna la visibilidad del grupo seleccionado
        shp.Visible = Not shp.Visible
        ' Oculta los dem s grupos
        ActiveSheet.Shapes("Celda IM").Visible = False
        ActiveSheet.Shapes("Celda QM").Visible = False
        ActiveSheet.Shapes("Celda GAM").Visible = False
        ActiveSheet.Shapes("Celda DM1").Visible = False
        ActiveSheet.Shapes("Seleccion COTA").Visible = False

    end If
end Sub
Sub seleccion_ACOTAR()
    Dim shp As Shape
    Dim groupname As String
    groupname = "Seleccion COTA"
    
    On Error Resume Next
    Set shp = ActiveSheet.Shapes(groupname)
    On Error GoTo 0
    
    If Not shp Is Nothing Then
        ' Alterna la visibilidad del grupo seleccionado
        shp.Visible = Not shp.Visible
        ' Oculta los dem s grupos
        ActiveSheet.Shapes("Celda IM").Visible = False
        ActiveSheet.Shapes("Celda QM").Visible = False
        ActiveSheet.Shapes("Celda GAM").Visible = False
        ActiveSheet.Shapes("Celda DM1").Visible = False
        ActiveSheet.Shapes("Celda GBC").Visible = False

    end If
end Sub
