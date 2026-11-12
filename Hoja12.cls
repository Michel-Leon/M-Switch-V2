'======== Mecanicos ===========
Private mbCargando As Boolean          'evita eventos en cascada
' Prepara la hoja destino y limpia la zona (celdas + shapes)
    Private Sub PrepararHoja(ByRef ws As Worksheet, ByVal sZona As String)
        Set ws = ThisWorkbook.Sheets("FrontalFM")
        Application.ScreenUpdating = False
        Application.CutCopyMode = False
        LimpiarZona ws, sZona
    End Sub

    ' Devuelve True si alguno de los valores esta vacio
    Private Function Vacio(ParamArray v() As Variant) As Boolean
        Dim i As Long
        For i = LBound(v) To UBound(v)
            If Trim$(CStr(v(i) & "")) = "" Then
                Vacio = True
                Exit Function
            End If
        Next i
    End Function

    Private Sub AvisoFaltan(ByVal sQue As String)
        Application.ScreenUpdating = True
        MsgBox "Faltan selecciones en: " & sQue, vbExclamation
    End Sub
'CELDA DM1A
    Private Sub ReleDM1_Change()
        Dim dato As String
        dato = Me.OLEObjects("ReleDM1").Object.Value
        '--- Contactor: P5U20 solo existe sin contactor
        With Me.OLEObjects("ContacDM1").Object
            .Clear
            Select Case dato
                Case "P3U30", "P3T32"
                    .AddItem "SI"
                    .AddItem "NO"
                Case "P5U20"
                    .AddItem "NO"
            End Select
        End With
        '--- Corriente
        With Me.OLEObjects("CorrienteDM1").Object
            .Clear
            Select Case dato
                Case "P3U30"
                    .AddItem "LPCTS"
                    .AddItem "CT"
                Case "P3T32"
                    .AddItem "CT"
                Case "P5U20"
                    .AddItem "LPCTS"
            End Select
        End With
        Me.OLEObjects("MtensionDM1").Object.Clear
        Me.OLEObjects("PM_DM1").Object.Clear
    End Sub

    Private Sub CorrienteDM1_Change()
        Dim dato As String, rele As String
        dato = Me.OLEObjects("CorrienteDM1").Object.Value
        rele = Me.OLEObjects("ReleDM1").Object.Value
        With Me.OLEObjects("MtensionDM1").Object
            .Clear
            Select Case dato
                Case "LPCTS"
                    .AddItem "NO"
                Case "CT"
                    If rele = "P3U30" Then .AddItem "NO"   'P3T32 solo existe con PT
                    .AddItem "PT"
            End Select
        End With
        Me.OLEObjects("PM_DM1").Object.Clear
    End Sub

    Private Sub MtensionDM1_Change()
        Dim contac As String, mt As String
        contac = Me.OLEObjects("ContacDM1").Object.Value
        mt = Me.OLEObjects("MtensionDM1").Object.Value
        With Me.OLEObjects("PM_DM1").Object
            .Clear
            If mt = "" Then Exit Sub
            '--- PM = SI solo existe con contactor SI y PT
            If contac = "SI" And mt = "PT" Then .AddItem "SI"
            .AddItem "NO"
        End With
    End Sub

    Private Sub ContacDM1_Change()
        MtensionDM1_Change      'si cambian el contactor, se recalcula PM
    End Sub
Public Sub Worksheet_Activate()
    '======== CELDA IM ==========================
        With  Me.OLEObjects("TensionIM").Object
            .clear
            .AddItem "24Kv"
        End With
        with  Me.OLEObjects("ApliIM").Object
            .clear
            .AddItem "Seccionamiento"
        End With
        with  Me.OLEObjects("ContacIM").Object
            .clear
            .AddItem "SI"
            .AddItem "NO"
        End With
        with  Me.OLEObjects("ZocaloIM").Object
            .clear
            .AddItem "SIN"
            .AddItem "400mm"
            .AddItem "600mm"
        End With
    '======= CELDA QM ==========================
        With  Me.OLEObjects("TensionQM").Object
            .clear
            .AddItem "24Kv"
        End With
        with  Me.OLEObjects("ApliQM").Object
            .clear
            .AddItem "Proteccion"
        End With
        with  Me.OLEObjects("ContacQM").Object
            .clear
            .AddItem "SI"
            .AddItem "NO"
        End With
        with  Me.OLEObjects("ZocaloQM").Object
            .clear
            .AddItem "SIN"
            .AddItem "400mm"
            .AddItem "600mm"
        End With
    '======= CELDA GAM ========================
        With  Me.OLEObjects("TensionGAM").Object
            .clear
            .AddItem "24Kv"
        End With
        with  Me.OLEObjects("ApliGAM").Object
            .clear
            .AddItem "Remonte"
        End With
        with  Me.OLEObjects("TipoGAM").Object
            .clear
            .AddItem "GAM_0"
        End With
        with  Me.OLEObjects("ZocaloGAM").Object
            .clear
            .AddItem "SIN"
            .AddItem "400mm"
            .AddItem "600mm"
        End With
    '======= CELDA GBC ========================
        With  Me.OLEObjects("TensionGBC").Object
            .clear
            .AddItem "24Kv"
        End With
        with  Me.OLEObjects("ApliGBC").Object
            .clear
            .AddItem "Medida"
        End With
        with  Me.OLEObjects("TipoGBC").Object
            .clear
            .AddItem "GBC_A"
            .AddItem "GBC_B"
            .AddItem "GBC_C"
            .AddItem "GBC_D"
        End With
        with  Me.OLEObjects("OrientacionGBC").Object
            .clear
            .AddItem "Derecha-Izquierda"
            .AddItem "Izquierda-Derecha"
        End With
        with  Me.OLEObjects("ZocaloGBC").Object
            .clear
            .AddItem "SIN"
            .AddItem "400mm"
            .AddItem "600mm"
        End With
    '======= CELDA DM1 ========================
        With Me.OLEObjects("TensionDM1").Object
            .Clear
            .AddItem "24Kv"
        End With
        With Me.OLEObjects("ReleDM1").Object
            .Clear
            .AddItem "P3U30"
            .AddItem "P3T32"
            .AddItem "P5U20"
        End With
        With Me.OLEObjects("ZocaloDM1").Object
            .Clear
            .AddItem "SIN"
            .AddItem "400mm"
            .AddItem "600mm"
        End With
        Me.OLEObjects("ContacDM1").Object.Clear
        Me.OLEObjects("CorrienteDM1").Object.Clear
        Me.OLEObjects("MtensionDM1").Object.Clear
        Me.OLEObjects("PM_DM1").Object.Clear
    With Me.OLEObjects("CotasMC").Object 
        .clear   ' Solo llena la lista si está vacía
        .AddItem "DM1A"
        .AddItem "QM"
        .AddItem "IM"
        .AddItem "GAM"
        .AddItem "GBC"
      
    End With   
end sub
Private Sub LimpiarZona(ws As Worksheet, ByVal sRango As String)
    Dim rngZona As Range, shp As Shape, i As Long

    Set rngZona = ws.Range(sRango)

    Application.CutCopyMode = False          'clave: sale del modo copiar

    'borrar de atras hacia adelante porque la coleccion se reindexa
    For i = ws.Shapes.Count To 1 Step -1
        Set shp = ws.Shapes(i)
        If Not Intersect(shp.TopLeftCell, rngZona) Is Nothing Then
            'no borrar los combos ActiveX si estuvieran en la zona
            If shp.Type <> msoOLEControlObject Then shp.Delete
        End If
    Next i

    rngZona.Clear
End Sub
Public sub celda_IM()
    '========================
        Set wsActual = ThisWorkbook.Sheets("Mecanicos")
        Celda = "KS97"
        Application.ScreenUpdating = False
        LimpiarZona wsActual, "KP95:NI189"
    '========================
    Dim dato1 As String, dato2 As String, dato3 As String, dato4 As String
    dato1 = Me.OLEObjects("TensionIM").Object.Value
    dato2 = Me.OLEObjects("ApliIM").Object.Value
    select case dato2
        case "Seccionamiento"
            dato2 = "SEC"
    end select
    dato3 = Me.OLEObjects("ContacIM").Object.Value
    dato4 = Me.OLEObjects("ZocaloIM").Object.Value
    CodeIM = "MC_IM" & dato1  & dato2 & "_" & dato3 & dato4
    MoverBloquMecanico CodeIM, "IM", wsActual, Celda
end sub
Public sub celda_QM()
    '========================
        Set wsActual = ThisWorkbook.Sheets("Mecanicos")
        Celda = "KS97"
        Application.ScreenUpdating = False
        LimpiarZona wsActual, "KP95:NI189"
    '========================
    Dim dato1 As String, dato2 As String, dato3 As String, dato4 As String
    dato1 = Me.OLEObjects("TensionQM").Object.Value
    dato2 = Me.OLEObjects("ApliQM").Object.Value
    select case dato2
        case "Proteccion"
            dato2 = "PRO"
    end select
    dato3 = Me.OLEObjects("ContacQM").Object.Value
    dato4 = Me.OLEObjects("ZocaloQM").Object.Value
    CodeQM = "MC_QM" & dato1  & dato2 & "_" & dato3 & dato4
    MoverBloquMecanico CodeQM, "QM", wsActual, Celda
end sub
Public sub celda_GAM()
    '========================
        Set wsActual = ThisWorkbook.Sheets("Mecanicos")
        Celda = "KS97"
        Application.ScreenUpdating = False
        LimpiarZona wsActual, "KP95:NI189"
    '========================
    Dim dato1 As String, dato2 As String, dato3 As String, dato4 As String
    dato1 = Me.OLEObjects("TensionGAM").Object.Value
    dato2 = Me.OLEObjects("ApliGAM").Object.Value
    select case dato2
        case "Remonte"
            dato2 = "REM"
    end select
    dato3 = Me.OLEObjects("TipoGAM").Object.Value
    dato4 = Me.OLEObjects("ZocaloGAM").Object.Value
    CodeGAM = "MC_GAM" & dato1  & dato2 & "_" & dato3 & dato4
    MoverBloquMecanico CodeGAM, "GAM", wsActual, Celda
end sub
Public sub celda_GBC()
    '========================
        Set wsActual = ThisWorkbook.Sheets("Mecanicos")
        Celda = "KS85"
        Application.ScreenUpdating = False
        LimpiarZona wsActual, "KP78:NI189"
    '========================
    Dim dato1 As String, dato2 As String, dato3 As String, dato4 As String, dato5 As String
    dato1 = Me.OLEObjects("TensionGBC").Object.Value
    dato2 = Me.OLEObjects("ApliGBC").Object.Value
    select case dato2
        case "Medida"
            dato2 = "MED"
    end select
    dato3 = Me.OLEObjects("TipoGBC").Object.Value
    dato4 = Me.OLEObjects("OrientacionGBC").Object.Value
    select case dato4
        case "Derecha-Izquierda"
            dato4 = "DI"
        case "Izquierda-Derecha"
            dato4 = "ID"
    end select
    dato5 = Me.OLEObjects("ZocaloGBC").Object.Value
    CodeGBC = "MC_" & dato1  & dato2  & dato3 & "_" & dato4 & dato5
    MoverBloquMecanico CodeGBC, "GBC 24kV 2 MED", wsActual, Celda
end sub 
Public Sub celda_DM1A()
    Dim dato1 As String, dato2 As String, dato3 As String, dato4 As String
    Dim dato5 As String, dato6 As String, dato7 As String
    dato1 = Me.OLEObjects("TensionDM1").Object.Value
    dato2 = Left(Me.OLEObjects("ReleDM1").Object.Value, 3)   'P3U30→P3U, P3T32→P3T, P5U20→P5U
    dato3 = Me.OLEObjects("ContacDM1").Object.Value
    dato4 = Me.OLEObjects("CorrienteDM1").Object.Value
    dato5 = Me.OLEObjects("MtensionDM1").Object.Value
    dato6 = Me.OLEObjects("ZocaloDM1").Object.Value
    dato7 = Me.OLEObjects("PM_DM1").Object.Value
    If dato6 = "SIN" Then dato6 = "NO"     'en la lista el zócalo sin base es "NO"

    If dato1 = "" Or dato2 = "" Or dato3 = "" Or dato4 = "" Or _
       dato5 = "" Or dato6 = "" Or dato7 = "" Then
        MsgBox "Complete todas las opciones de la celda DM1A.", vbExclamation
        Exit Sub
    End If
    '========================
    Set wsActual = ThisWorkbook.Sheets("Mecanicos")
    Celda = "KS97"
    Application.ScreenUpdating = False
    LimpiarZona wsActual, "KO78:NI189"
    '========================
    CodeDM1A = "MC_" & dato1 & "DM1A" & dato2 & "_" & dato3 & dato4 & "_" & dato5 & dato6 & dato7
    MoverBloquMecanico CodeDM1A, "DM", wsActual, Celda
    Application.ScreenUpdating = True
End Sub
Public Sub Generar_CotasMecanico()
   '========================
        Set wsActual = ThisWorkbook.Sheets("Mecanicos")
        Celda = "KS97"
        Application.ScreenUpdating = False
        LimpiarZona wsActual, "KP95:NI189"
    '========================
    Dim dato1 As String
    dato1 = Me.OLEObjects("CotasMC").Object.Value
    ResultadoCotas = "Cotas_" & dato1
    MoverBloquMecanico ResultadoCotas, "IM", wsActual, Celda
End Sub
'-----------------------------------------
public sub  Limpiar_zona2()
    thisWorkbook.Sheets("Mecanicos").Range("KO78:NI189").Clear
End Sub 
Public Sub Limpiar_zona1()
    Dim borde As Variant

    With ThisWorkbook.Sheets("Mecanicos").Range("CM45:IZ148")
        .Clear
        For Each borde In Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight)
            With .Borders(borde)
                .LineStyle = xlContinuous
                .Weight = xlHairline          ' Línea punteada extra delgada
                .Color = RGB(191, 191, 191)   ' #BFBFBF
            End With
        Next borde
    End With
    With ThisWorkbook.Sheets("Mecanicos").Range("CM149:IZ166")
        .Clear
    end with
End Sub