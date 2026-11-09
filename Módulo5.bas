Attribute VB_Name = "M�dulo5"
Public Sub Seleccion_Menu_IM()
    Dim shp As Shape
    Dim groupname As String
    groupname = "Menu IM"
    
    On Error Resume Next
    Set shp = ActiveSheet.Shapes(groupname)
    On Error GoTo 0
    
    If Not shp Is Nothing Then
        ' Alterna la visibilidad del grupo seleccionado
        shp.Visible = Not shp.Visible
        ' Oculta los dem s grupos
        ActiveSheet.Shapes("Menu QM").Visible = False
        ActiveSheet.Shapes("Menu GAM").Visible = False
        ActiveSheet.Shapes("Menu DM1").Visible = False
        ActiveSheet.Shapes("Menu GBC").Visible = False
        ActiveSheet.Shapes("Menu Cotas").Visible = False
      End If
End Sub
Public Sub Seleccion_Menu_QM()
    Dim shp As Shape
    Dim groupname As String
    groupname = "Menu QM"
    
    On Error Resume Next
    Set shp = ActiveSheet.Shapes(groupname)
    On Error GoTo 0
    
    If Not shp Is Nothing Then
        ' Alterna la visibilidad del grupo seleccionado
        shp.Visible = Not shp.Visible
        ' Oculta los demás grupos
        ActiveSheet.Shapes("Menu IM").Visible = False
        ActiveSheet.Shapes("Menu GAM").Visible = False
        ActiveSheet.Shapes("Menu DM1").Visible = False
        ActiveSheet.Shapes("Menu GBC").Visible = False
        ActiveSheet.Shapes("Menu Cotas").Visible = False
      End If
End Sub
public Sub Seleccion_Menu_GAM()
    Dim shp As Shape
    Dim groupname As String
    groupname = "Menu GAM"
    
    On Error Resume Next
    Set shp = ActiveSheet.Shapes(groupname)
    On Error GoTo 0
    
    If Not shp Is Nothing Then
        ' Alterna la visibilidad del grupo seleccionado
        shp.Visible = Not shp.Visible
        ' Oculta los demás grupos
        ActiveSheet.Shapes("Menu IM").Visible = False
        ActiveSheet.Shapes("Menu QM").Visible = False
        ActiveSheet.Shapes("Menu DM1").Visible = False
        ActiveSheet.Shapes("Menu GBC").Visible = False
        ActiveSheet.Shapes("Menu Cotas").Visible = False
      End If
End Sub
public Sub Seleccion_Menu_DM1()
    Dim shp As Shape
    Dim groupname As String
    groupname = "Menu DM1"
    
    On Error Resume Next
    Set shp = ActiveSheet.Shapes(groupname)
    On Error GoTo 0
    
    If Not shp Is Nothing Then
        ' Alterna la visibilidad del grupo seleccionado
        shp.Visible = Not shp.Visible
        ' Oculta los demás grupos
        ActiveSheet.Shapes("Menu IM").Visible = False
        ActiveSheet.Shapes("Menu QM").Visible = False
        ActiveSheet.Shapes("Menu GAM").Visible = False
        ActiveSheet.Shapes("Menu GBC").Visible = False
        ActiveSheet.Shapes("Menu Cotas").Visible = False
      End If
End Sub
public Sub Seleccion_Menu_GBC()
    Dim shp As Shape
    Dim groupname As String
    groupname = "Menu GBC"
    
    On Error Resume Next
    Set shp = ActiveSheet.Shapes(groupname)
    On Error GoTo 0
    
    If Not shp Is Nothing Then
        ' Alterna la visibilidad del grupo seleccionado
        shp.Visible = Not shp.Visible
        ' Oculta los demás grupos
        ActiveSheet.Shapes("Menu IM").Visible = False
        ActiveSheet.Shapes("Menu QM").Visible = False
        ActiveSheet.Shapes("Menu GAM").Visible = False
        ActiveSheet.Shapes("Menu DM1").Visible = False
        ActiveSheet.Shapes("Menu Cotas").Visible = False
      End If
End Sub
public Sub Seleccion_Menu_Cotas()
    Dim shp As Shape
    Dim groupname As String
    groupname = "Menu Cotas"
    
    On Error Resume Next
    Set shp = ActiveSheet.Shapes(groupname)
    On Error GoTo 0
    
    If Not shp Is Nothing Then
        ' Alterna la visibilidad del grupo seleccionado
        shp.Visible = Not shp.Visible
        ' Oculta los demás grupos
        ActiveSheet.Shapes("Menu IM").Visible = False
        ActiveSheet.Shapes("Menu QM").Visible = False
        ActiveSheet.Shapes("Menu GAM").Visible = False
        ActiveSheet.Shapes("Menu DM1").Visible = False
        ActiveSheet.Shapes("Menu GBC").Visible = False
      End If
End Sub