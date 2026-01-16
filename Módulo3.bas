Attribute VB_Name = "Módulo3"
'------------------ funciones para el configurador ------------------
Sub seleccion_Contactores()
    Dim shp As Shape
    Dim groupname As String
    groupname = "AC_CONTACTORES"
    
    On Error Resume Next
    Set shp = ActiveSheet.Shapes(groupname)
    On Error GoTo 0
    
    If Not shp Is Nothing Then
        ' Alterna la visibilidad del grupo seleccionado
        shp.Visible = Not shp.Visible
        ' Oculta los demás grupos
        ActiveSheet.Shapes("AC_SENSORES").Visible = False
        ActiveSheet.Shapes("AC_MODULO SINCRONISMO").Visible = False
        ActiveSheet.Shapes("AC_TRANSFORMADOR DE CORRIENTE").Visible = False
        ActiveSheet.Shapes("AC_VPIS").Visible = False
        ActiveSheet.Shapes("AC_RELE DE PROTECCION").Visible = False
        ActiveSheet.Shapes("AC_TRANSFORMADOR DE POTENCIA").Visible = False
        ActiveSheet.Shapes("AC_FUSIBLES").Visible = False
        ActiveSheet.Shapes("AC_ANALIZADOR DE REDES").Visible = False
        ActiveSheet.Shapes("AC_MEDIDPOR TARIFARIO").Visible = False
        ActiveSheet.Shapes("AC_GABINETE BAJA TENSION").Visible = False
        ActiveSheet.Shapes("AC_MANDO MOTORIZADO").Visible = False
        ActiveSheet.Shapes("AC_GATEWAY").Visible = False
        ActiveSheet.Shapes("AC_ADAPTADOR DE TENSION").Visible = False
        ActiveSheet.Shapes("AC_ZOCALO").Visible = False
    End If
End Sub
Sub seleccion_Sensores()
    Dim shp As Shape
    Dim groupname As String
    groupname = "AC_SENSORES"
    
    On Error Resume Next
    Set shp = ActiveSheet.Shapes(groupname)
    On Error GoTo 0
    
    If Not shp Is Nothing Then
        ' Alterna la visibilidad del grupo seleccionado
        shp.Visible = Not shp.Visible
        ' Oculta los demás grupos
        ActiveSheet.Shapes("AC_CONTACTORES").Visible = False
        ActiveSheet.Shapes("AC_MODULO SINCRONISMO").Visible = False
        ActiveSheet.Shapes("AC_TRANSFORMADOR DE CORRIENTE").Visible = False
        ActiveSheet.Shapes("AC_VPIS").Visible = False
        ActiveSheet.Shapes("AC_RELE DE PROTECCION").Visible = False
        ActiveSheet.Shapes("AC_TRANSFORMADOR DE POTENCIA").Visible = False
        ActiveSheet.Shapes("AC_FUSIBLES").Visible = False
        ActiveSheet.Shapes("AC_ANALIZADOR DE REDES").Visible = False
        ActiveSheet.Shapes("AC_MEDIDPOR TARIFARIO").Visible = False
        ActiveSheet.Shapes("AC_GABINETE BAJA TENSION").Visible = False
        ActiveSheet.Shapes("AC_MANDO MOTORIZADO").Visible = False
        ActiveSheet.Shapes("AC_GATEWAY").Visible = False
        ActiveSheet.Shapes("AC_ADAPTADOR DE TENSION").Visible = False
        ActiveSheet.Shapes("AC_ZOCALO").Visible = False
    End If
End Sub
Sub seleccion_Modulo_Sincronismo()
    Dim shp As Shape
    Dim groupname As String
    groupname = "AC_MODULO SINCRONISMO"
    
    On Error Resume Next
    Set shp = ActiveSheet.Shapes(groupname)
    On Error GoTo 0
    
    If Not shp Is Nothing Then
        ' Alterna la visibilidad del grupo seleccionado
        shp.Visible = Not shp.Visible
        ' Oculta los demás grupos
        ActiveSheet.Shapes("AC_CONTACTORES").Visible = False
        ActiveSheet.Shapes("AC_SENSORES").Visible = False
        ActiveSheet.Shapes("AC_TRANSFORMADOR DE CORRIENTE").Visible = False
        ActiveSheet.Shapes("AC_VPIS").Visible = False
        ActiveSheet.Shapes("AC_RELE DE PROTECCION").Visible = False 
        ActiveSheet.Shapes("AC_TRANSFORMADOR DE POTENCIA").Visible = False
        ActiveSheet.Shapes("AC_FUSIBLES").Visible = False
        ActiveSheet.Shapes("AC_ANALIZADOR DE REDES").Visible = False
        ActiveSheet.Shapes("AC_MEDIDPOR TARIFARIO").Visible = False
        ActiveSheet.Shapes("AC_GABINETE BAJA TENSION").Visible = False
        ActiveSheet.Shapes("AC_MANDO MOTORIZADO").Visible = False
        ActiveSheet.Shapes("AC_GATEWAY").Visible = False
        ActiveSheet.Shapes("AC_ADAPTADOR DE TENSION").Visible = False
        ActiveSheet.Shapes("AC_ZOCALO").Visible = False
    End If
End Sub
Sub seleccion_Transformador_Corriente()
    Dim shp As Shape
    Dim groupname As String
    groupname = "AC_TRANSFORMADOR DE CORRIENTE"
    
    On Error Resume Next    
    Set shp = ActiveSheet.Shapes(groupname)
    On Error GoTo 0
    If Not shp Is Nothing Then
        ' Alterna la visibilidad del grupo seleccionado
        shp.Visible = Not shp.Visible
        ' Oculta los demás grupos
        ActiveSheet.Shapes("AC_CONTACTORES").Visible = False
        ActiveSheet.Shapes("AC_SENSORES").Visible = False
        ActiveSheet.Shapes("AC_MODULO SINCRONISMO").Visible = False
        ActiveSheet.Shapes("AC_VPIS").Visible = False
        ActiveSheet.Shapes("AC_RELE DE PROTECCION").Visible = False
        ActiveSheet.Shapes("AC_TRANSFORMADOR DE POTENCIA").Visible = False
        ActiveSheet.Shapes("AC_FUSIBLES").Visible = False
        ActiveSheet.Shapes("AC_ANALIZADOR DE REDES").Visible = False
        ActiveSheet.Shapes("AC_MEDIDPOR TARIFARIO").Visible = False
        ActiveSheet.Shapes("AC_GABINETE BAJA TENSION").Visible = False
        ActiveSheet.Shapes("AC_MANDO MOTORIZADO").Visible = False
        ActiveSheet.Shapes("AC_GATEWAY").Visible = False
        ActiveSheet.Shapes("AC_ADAPTADOR DE TENSION").Visible = False
        ActiveSheet.Shapes("AC_ZOCALO").Visible = False
    End If
End Sub
Sub seleccion_VPIS()
    Dim shp As Shape
    Dim groupname As String
    groupname = "AC_VPIS"
    
    On Error Resume Next    
    Set shp = ActiveSheet.Shapes(groupname)
    On Error GoTo 0
    If Not shp Is Nothing Then
        ' Alterna la visibilidad del grupo seleccionado
        shp.Visible = Not shp.Visible
        ' Oculta los demás grupos
        ActiveSheet.Shapes("AC_CONTACTORES").Visible = False
        ActiveSheet.Shapes("AC_SENSORES").Visible = False
        ActiveSheet.Shapes("AC_MODULO SINCRONISMO").Visible = False
        ActiveSheet.Shapes("AC_TRANSFORMADOR DE CORRIENTE").Visible = False
        ActiveSheet.Shapes("AC_RELE DE PROTECCION").Visible = False
        ActiveSheet.Shapes("AC_TRANSFORMADOR DE POTENCIA").Visible = False
        ActiveSheet.Shapes("AC_FUSIBLES").Visible = False
        ActiveSheet.Shapes("AC_ANALIZADOR DE REDES").Visible = False
        ActiveSheet.Shapes("AC_MEDIDPOR TARIFARIO").Visible = False
        ActiveSheet.Shapes("AC_GABINETE BAJA TENSION").Visible = False
        ActiveSheet.Shapes("AC_MANDO MOTORIZADO").Visible = False
        ActiveSheet.Shapes("AC_GATEWAY").Visible = False
        ActiveSheet.Shapes("AC_ADAPTADOR DE TENSION").Visible = False
        ActiveSheet.Shapes("AC_ZOCALO").Visible = False
    End If
End Sub
Sub seleccion_Rele_Proteccion()
    Dim shp As Shape
    Dim groupname As String
    groupname = "AC_RELE DE PROTECCION"
    
    On Error Resume Next    
    Set shp = ActiveSheet.Shapes(groupname)
    On Error GoTo 0
    If Not shp Is Nothing Then
        ' Alterna la visibilidad del grupo seleccionado
        shp.Visible = Not shp.Visible
        ' Oculta los demás grupos
        ActiveSheet.Shapes("AC_CONTACTORES").Visible = False
        ActiveSheet.Shapes("AC_SENSORES").Visible = False
        ActiveSheet.Shapes("AC_MODULO SINCRONISMO").Visible = False
        ActiveSheet.Shapes("AC_TRANSFORMADOR DE CORRIENTE").Visible = False
        ActiveSheet.Shapes("AC_VPIS").Visible = False        
        ActiveSheet.Shapes("AC_TRANSFORMADOR DE POTENCIA").Visible = False
        ActiveSheet.Shapes("AC_FUSIBLES").Visible = False
        ActiveSheet.Shapes("AC_ANALIZADOR DE REDES").Visible = False
        ActiveSheet.Shapes("AC_MEDIDPOR TARIFARIO").Visible = False
        ActiveSheet.Shapes("AC_GABINETE BAJA TENSION").Visible = False
        ActiveSheet.Shapes("AC_MANDO MOTORIZADO").Visible = False
        ActiveSheet.Shapes("AC_GATEWAY").Visible = False
        ActiveSheet.Shapes("AC_ADAPTADOR DE TENSION").Visible = False
        ActiveSheet.Shapes("AC_ZOCALO").Visible = False        
    End If
End Sub 
Sub seleccion_Transformador_Potencia()
    Dim shp As Shape
    Dim groupname As String
    groupname = "AC_TRANSFORMADOR DE POTENCIA"
    
    On Error Resume Next    
    Set shp = ActiveSheet.Shapes(groupname)
    On Error GoTo 0
    If Not shp Is Nothing Then
        ' Alterna la visibilidad del grupo seleccionado
        shp.Visible = Not shp.Visible
        ' Oculta los demás grupos
        ActiveSheet.Shapes("AC_CONTACTORES").Visible = False
        ActiveSheet.Shapes("AC_SENSORES").Visible = False
        ActiveSheet.Shapes("AC_MODULO SINCRONISMO").Visible = False
        ActiveSheet.Shapes("AC_TRANSFORMADOR DE CORRIENTE").Visible = False
        ActiveSheet.Shapes("AC_VPIS").Visible = False   
        ActiveSheet.Shapes("AC_RELE DE PROTECCION").Visible = False
        ActiveSheet.Shapes("AC_FUSIBLES").Visible = False
        ActiveSheet.Shapes("AC_ANALIZADOR DE REDES").Visible = False
        ActiveSheet.Shapes("AC_MEDIDPOR TARIFARIO").Visible = False
        ActiveSheet.Shapes("AC_GABINETE BAJA TENSION").Visible = False
        ActiveSheet.Shapes("AC_MANDO MOTORIZADO").Visible = False
        ActiveSheet.Shapes("AC_GATEWAY").Visible = False
        ActiveSheet.Shapes("AC_ADAPTADOR DE TENSION").Visible = False
        ActiveSheet.Shapes("AC_ZOCALO").Visible = False        
    End If
End Sub
Sub seleccion_Fusibles()
    Dim shp As Shape
    Dim groupname As String
    groupname = "AC_FUSIBLES"
    
    On Error Resume Next    
    Set shp = ActiveSheet.Shapes(groupname)
    On Error GoTo 0
    If Not shp Is Nothing Then
        ' Alterna la visibilidad del grupo seleccionado
        shp.Visible = Not shp.Visible
        ' Oculta los demás grupos
        ActiveSheet.Shapes("AC_CONTACTORES").Visible = False
        ActiveSheet.Shapes("AC_SENSORES").Visible = False
        ActiveSheet.Shapes("AC_MODULO SINCRONISMO").Visible = False
        ActiveSheet.Shapes("AC_TRANSFORMADOR DE CORRIENTE").Visible = False
        ActiveSheet.Shapes("AC_VPIS").Visible = False
        ActiveSheet.Shapes("AC_RELE DE PROTECCION").Visible = False        
        ActiveSheet.Shapes("AC_TRANSFORMADOR DE POTENCIA").Visible = False
        ActiveSheet.Shapes("AC_ANALIZADOR DE REDES").Visible = False        
        ActiveSheet.Shapes("AC_MEDIDPOR TARIFARIO").Visible = False
        ActiveSheet.Shapes("AC_GABINETE BAJA TENSION").Visible = False
        ActiveSheet.Shapes("AC_MANDO MOTORIZADO").Visible = False        
        ActiveSheet.Shapes("AC_GATEWAY").Visible = False
        ActiveSheet.Shapes("AC_ADAPTADOR DE TENSION").Visible = False
        ActiveSheet.Shapes("AC_ZOCALO").Visible = False        
    End If
End Sub     
Sub seleccion_Analizador_Redes()
    Dim shp As Shape
    Dim groupname As String
    groupname = "AC_ANALIZADOR DE REDES"
    
    On Error Resume Next    
    Set shp = ActiveSheet.Shapes(groupname)
    On Error GoTo 0
    If Not shp Is Nothing Then
        ' Alterna la visibilidad del grupo seleccionado
        shp.Visible = Not shp.Visible
        ' Oculta los demás grupos
        ActiveSheet.Shapes("AC_CONTACTORES").Visible = False
        ActiveSheet.Shapes("AC_SENSORES").Visible = False
        ActiveSheet.Shapes("AC_MODULO SINCRONISMO").Visible = False
        ActiveSheet.Shapes("AC_TRANSFORMADOR DE CORRIENTE").Visible = False
        ActiveSheet.Shapes("AC_VPIS").Visible = False
        ActiveSheet.Shapes("AC_RELE DE PROTECCION").Visible = False        
        ActiveSheet.Shapes("AC_TRANSFORMADOR DE POTENCIA").Visible = False
        ActiveSheet.Shapes("AC_FUSIBLES").Visible = False        
        ActiveSheet.Shapes("AC_MEDIDPOR TARIFARIO").Visible = False
        ActiveSheet.Shapes("AC_GABINETE BAJA TENSION").Visible = False
        ActiveSheet.Shapes("AC_MANDO MOTORIZADO").Visible = False        
        ActiveSheet.Shapes("AC_GATEWAY").Visible = False
        ActiveSheet.Shapes("AC_ADAPTADOR DE TENSION").Visible = False
        ActiveSheet.Shapes("AC_ZOCALO").Visible = False        
    End If
End Sub
Sub seleccion_Medidor_Tarifario()
    Dim shp As Shape
    Dim groupname As String
    groupname = "AC_MEDIDPOR TARIFARIO"
    
    On Error Resume Next    
    Set shp = ActiveSheet.Shapes(groupname)
    On Error GoTo 0
    If Not shp Is Nothing Then
        ' Alterna la visibilidad del grupo seleccionado
        shp.Visible = Not shp.Visible
        ' Oculta los demás grupos
        ActiveSheet.Shapes("AC_CONTACTORES").Visible = False
        ActiveSheet.Shapes("AC_SENSORES").Visible = False
        ActiveSheet.Shapes("AC_MODULO SINCRONISMO").Visible = False
        ActiveSheet.Shapes("AC_TRANSFORMADOR DE CORRIENTE").Visible = False
        ActiveSheet.Shapes("AC_VPIS").Visible = False       
        ActiveSheet.Shapes("AC_RELE DE PROTECCION").Visible = False
        ActiveSheet.Shapes("AC_TRANSFORMADOR DE POTENCIA").Visible = False
        ActiveSheet.Shapes("AC_FUSIBLES").Visible = False        
        ActiveSheet.Shapes("AC_ANALIZADOR DE REDES").Visible = False
        ActiveSheet.Shapes("AC_GABINETE BAJA TENSION").Visible = False
        ActiveSheet.Shapes("AC_MANDO MOTORIZADO").Visible = False        
        ActiveSheet.Shapes("AC_GATEWAY").Visible = False
        ActiveSheet.Shapes("AC_ADAPTADOR DE TENSION").Visible = False
        ActiveSheet.Shapes("AC_ZOCALO").Visible = False        
    End If
End Sub 
Sub seleccion_Gabinete_Baja_Tension()
    Dim shp As Shape
    Dim groupname As String
    groupname = "AC_GABINETE BAJA TENSION"
    
    On Error Resume Next    
    Set shp = ActiveSheet.Shapes(groupname)
    On Error GoTo 0
    If Not shp Is Nothing Then
        ' Alterna la visibilidad del grupo seleccionado
        shp.Visible = Not shp.Visible
        ' Oculta los demás grupos
        ActiveSheet.Shapes("AC_CONTACTORES").Visible = False
        ActiveSheet.Shapes("AC_SENSORES").Visible = False
        ActiveSheet.Shapes("AC_MODULO SINCRONISMO").Visible = False
        ActiveSheet.Shapes("AC_TRANSFORMADOR DE CORRIENTE").Visible = False
        ActiveSheet.Shapes("AC_VPIS").Visible = False
        ActiveSheet.Shapes("AC_RELE DE PROTECCION").Visible = False        
        ActiveSheet.Shapes("AC_TRANSFORMADOR DE POTENCIA").Visible = False
        ActiveSheet.Shapes("AC_FUSIBLES").Visible = False        
        ActiveSheet.Shapes("AC_ANALIZADOR DE REDES").Visible = False
        ActiveSheet.Shapes("AC_MEDIDPOR TARIFARIO").Visible = False
        ActiveSheet.Shapes("AC_MANDO MOTORIZADO").Visible = False        
        ActiveSheet.Shapes("AC_GATEWAY").Visible = False
        ActiveSheet.Shapes("AC_ADAPTADOR DE TENSION").Visible = False
        ActiveSheet.Shapes("AC_ZOCALO").Visible = False        
    End If
End Sub 
Sub seleccion_Mando_Motorizado()
    Dim shp As Shape
    Dim groupname As String
    groupname = "AC_MANDO MOTORIZADO"
    
    On Error Resume Next    
    Set shp = ActiveSheet.Shapes(groupname)
    On Error GoTo 0
    If Not shp Is Nothing Then
        ' Alterna la visibilidad del grupo seleccionado
        shp.Visible = Not shp.Visible
        ' Oculta los demás grupos
        ActiveSheet.Shapes("AC_CONTACTORES").Visible = False
        ActiveSheet.Shapes("AC_SENSORES").Visible = False
        ActiveSheet.Shapes("AC_MODULO SINCRONISMO").Visible = False
        ActiveSheet.Shapes("AC_TRANSFORMADOR DE CORRIENTE").Visible = False
        ActiveSheet.Shapes("AC_VPIS").Visible = False
        ActiveSheet.Shapes("AC_RELE DE PROTECCION").Visible = False        
        ActiveSheet.Shapes("AC_TRANSFORMADOR DE POTENCIA").Visible = False
        ActiveSheet.Shapes("AC_FUSIBLES").Visible = False        
        ActiveSheet.Shapes("AC_ANALIZADOR DE REDES").Visible = False
        ActiveSheet.Shapes("AC_MEDIDPOR TARIFARIO").Visible = False
        ActiveSheet.Shapes("AC_GABINETE BAJA TENSION").Visible = False        
        ActiveSheet.Shapes("AC_GATEWAY").Visible = False
        ActiveSheet.Shapes("AC_ADAPTADOR DE TENSION").Visible = False
        ActiveSheet.Shapes("AC_ZOCALO").Visible = False        
    End If
End Sub 
Sub seleccion_Gateway()
    Dim shp As Shape
    Dim groupname As String
    groupname = "AC_GATEWAY"
    
    On Error Resume Next    
    Set shp = ActiveSheet.Shapes(groupname)
    On Error GoTo 0
    If Not shp Is Nothing Then
        ' Alterna la visibilidad del grupo seleccionado
        shp.Visible = Not shp.Visible
        ' Oculta los demás grupos
        ActiveSheet.Shapes("AC_CONTACTORES").Visible = False
        ActiveSheet.Shapes("AC_SENSORES").Visible = False
        ActiveSheet.Shapes("AC_MODULO SINCRONISMO").Visible = False
        ActiveSheet.Shapes("AC_TRANSFORMADOR DE CORRIENTE").Visible = False
        ActiveSheet.Shapes("AC_VPIS").Visible = False
        ActiveSheet.Shapes("AC_RELE DE PROTECCION").Visible = False        
        ActiveSheet.Shapes("AC_TRANSFORMADOR DE POTENCIA").Visible = False
        ActiveSheet.Shapes("AC_FUSIBLES").Visible = False        
        ActiveSheet.Shapes("AC_ANALIZADOR DE REDES").Visible = False
        ActiveSheet.Shapes("AC_MEDIDPOR TARIFARIO").Visible = False
        ActiveSheet.Shapes("AC_GABINETE BAJA TENSION").Visible = False        
        ActiveSheet.Shapes("AC_MANDO MOTORIZADO").Visible = False
        ActiveSheet.Shapes("AC_ADAPTADOR DE TENSION").Visible = False
        ActiveSheet.Shapes("AC_ZOCALO").Visible = False        
    End If
End Sub
Sub seleccion_Adaptador_Tension()
    Dim shp As Shape
    Dim groupname As String
    groupname = "AC_ADAPTADOR DE TENSION"
    
    On Error Resume Next    
    Set shp = ActiveSheet.Shapes(groupname)
    On Error GoTo 0
    If Not shp Is Nothing Then
        ' Alterna la visibilidad del grupo seleccionado
        shp.Visible = Not shp.Visible
        ' Oculta los demás grupos
        ActiveSheet.Shapes("AC_CONTACTORES").Visible = False
        ActiveSheet.Shapes("AC_SENSORES").Visible = False
        ActiveSheet.Shapes("AC_MODULO SINCRONISMO").Visible = False
        ActiveSheet.Shapes("AC_TRANSFORMADOR DE CORRIENTE").Visible = False
        ActiveSheet.Shapes("AC_VPIS").Visible = False
        ActiveSheet.Shapes("AC_RELE DE PROTECCION").Visible = False        
        ActiveSheet.Shapes("AC_TRANSFORMADOR DE POTENCIA").Visible = False
        ActiveSheet.Shapes("AC_FUSIBLES").Visible = False        
        ActiveSheet.Shapes("AC_ANALIZADOR DE REDES").Visible = False
        ActiveSheet.Shapes("AC_MEDIDPOR TARIFARIO").Visible = False
        ActiveSheet.Shapes("AC_GABINETE BAJA TENSION").Visible = False        
        ActiveSheet.Shapes("AC_MANDO MOTORIZADO").Visible = False
        ActiveSheet.Shapes("AC_GATEWAY").Visible = False
        ActiveSheet.Shapes("AC_ZOCALO").Visible = False        
    End If
End Sub
Sub seleccion_Zocalo()
    Dim shp As Shape
    Dim groupname As String
    groupname = "AC_ZOCALO"
    
    On Error Resume Next    
    Set shp = ActiveSheet.Shapes(groupname)
    On Error GoTo 0
    If Not shp Is Nothing Then
        ' Alterna la visibilidad del grupo seleccionado
        shp.Visible = Not shp.Visible
        ' Oculta los demás grupos
        ActiveSheet.Shapes("AC_CONTACTORES").Visible = False
        ActiveSheet.Shapes("AC_SENSORES").Visible = False
        ActiveSheet.Shapes("AC_MODULO SINCRONISMO").Visible = False
        ActiveSheet.Shapes("AC_TRANSFORMADOR DE CORRIENTE").Visible = False
        ActiveSheet.Shapes("AC_VPIS").Visible = False
        ActiveSheet.Shapes("AC_RELE DE PROTECCION").Visible = False        
        ActiveSheet.Shapes("AC_TRANSFORMADOR DE POTENCIA").Visible = False
        ActiveSheet.Shapes("AC_FUSIBLES").Visible = False        
        ActiveSheet.Shapes("AC_ANALIZADOR DE REDES").Visible = False
        ActiveSheet.Shapes("AC_MEDIDPOR TARIFARIO").Visible = False
        ActiveSheet.Shapes("AC_GABINETE BAJA TENSION").Visible = False        
        ActiveSheet.Shapes("AC_MANDO MOTORIZADO").Visible = False
        ActiveSheet.Shapes("AC_GATEWAY").Visible = False
        ActiveSheet.Shapes("AC_ADAPTADOR DE TENSION").Visible = False        
    End If
End Sub
