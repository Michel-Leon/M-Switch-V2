'======== UNIFILAR ===========
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
Private Sub Worksheet_Activate()
    With Me.CotasUN
        If .ListCount = 0 Then          ' Solo llena la lista si está vacía
            .AddItem "DM1A"
            .AddItem "QM"
            .AddItem "IM"
            .AddItem "GAM"
            .AddItem "GBC"
        End If
    End With
End Sub

' Se ejecuta automáticamente al elegir una opción en el ComboBox
Private Sub CotasUnifilar_Change()
    Generar_CotasUnifilar
End Sub

Public Sub Generar_CotasUnifilar()
    Dim wsActual As Worksheet
    Dim Celda As String
    Dim dato1 As String
    Dim ResultadoCotas As String

    dato1 = Trim(Me.CotasUN.Value)
    If dato1 = "" Then
        MsgBox "Seleccione una opción en el ComboBox.", vbExclamation
        Exit Sub
    End If

    Set wsActual = Me
    Celda = "KS81"

    Application.ScreenUpdating = False
    LimpiarZona wsActual, "KO78:NI189"

    ResultadoCotas = "C_" & dato1
    MoverBloquMecanico ResultadoCotas, "Diagrama_unifilar", wsActual, Celda
    Application.ScreenUpdating = True
End Sub
public sub  Limpiar_zona2()
    thisWorkbook.Sheets("Unifilar").Range("KO78:NI189").Clear 
End Sub 