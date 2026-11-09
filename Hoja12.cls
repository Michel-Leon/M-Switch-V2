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
        With  Me.OLEObjects("TensionDM1").Object
            .clear
            .AddItem "24Kv"
        End With
        with  Me.OLEObjects("ReleDM1").Object
            .clear
            .AddItem "P3U30"
        End With
        with  Me.OLEObjects("ContacDM1").Object
            .clear
            .AddItem "SI"
            .AddItem "NO"
        End With
        with  Me.OLEObjects("CorrienteDM1").Object
            .clear
            .AddItem "LPCT"
            .AddItem "CT"
        End With
        with  Me.OLEObjects("MtensionDM1").Object
            .clear
            .AddItem "PT"
            .AddItem "NO"
        End With
        with  Me.OLEObjects("ZocaloDM1").Object
            .clear
            .AddItem "SIN"
            .AddItem "400mm"
            .AddItem "600mm"
        End With
        with  Me.OLEObjects("PM_DM1").Object
            .clear
            .AddItem "SI"
            .AddItem "NO"
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
'-----------------------------------------
public sub  Limpiar_zona2()
    thisWorkbook.Sheets("Mecanicos").Range("KO94:NI189").Clear
End Sub 