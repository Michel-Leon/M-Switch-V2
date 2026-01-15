'================ HIJA UNIFILAR ================
'================= VARIABLES GLOBALES ===========================
Dim resultadoUNIF As String
PRIVATE Sub TensionDM_Change()
    Dim Tension as String
    Tension = Me.OLEObjects("TensionDM").Object.Value
    With Me.OLEObjects("EspecialDM").Object
        .Clear
        Select case Tension
            Case "24"
            .AddItem ""
            .AddItem "2CT"
            '.AddItem "CT-PT"
            .AddItem "NA"
            Case "36"
            .AddItem ""
            .AddItem "2CT" 
            .AddItem "NA"
        End Select
    End With
End Sub
PRIVATE Sub TipoGBC_Change()
    Dim Tipo as String
    Tipo = Me.OLEObjects("TipoGBC").Object.Value
    With Me.OLEObjects("MedidoresGBC").Object
        .Clear 
        Select case Tipo
            Case "A", "B", "D" 
                .AddItem ""
                .AddItem "1M"
                .AddItem "2M"
            Case "C"
                .AddItem ""
                .AddItem "1M" 
                .AddItem "2M"
        End Select        
    End With
End Sub

''--------------------------------CARGA DE COMBOS AL ACTIVAR LA HOJA-------------------------------
Public Sub Worksheet_Activate()
    With Me.OLEObjects("TensionIM").Object
        .Clear
        .AddItem ""
        .AddItem "24"
        .AddItem "36"
    End With
    With Me.OLEObjects("DirreccionIM").Object
        .Clear
        .AddItem ""
        .AddItem "Central"
        .AddItem "Izquierda"
        .AddItem "Derecha"
    End With
    With Me.OLEObjects("AppIM").Object
        .Clear
        .AddItem ""
        .AddItem "Seccionamiento"
    End With
    '---------------------------------------------------------------
     With Me.OLEObjects("TensionQM").Object
        .Clear
        .AddItem ""
        .AddItem "24"
        .AddItem "36"
    End With
    With Me.OLEObjects("DireccionQM").Object
        .Clear
        .AddItem ""
        .AddItem "Central"
        .AddItem "Izquierda"
        .AddItem "Derecha"
    End With
    With Me.OLEObjects("AppQM").Object
        .Clear
        .AddItem ""
        .AddItem "Proteccion"
    End With
    '---------------------------------------------------------------
    With Me.OLEObjects("TensionGAM").Object
        .Clear
        .AddItem ""
        .AddItem "24"
        .AddItem "36"
    End With
    With Me.OLEObjects("DireccionGAM").Object
        .Clear
        .AddItem ""
        .AddItem "Izquierda"
        .AddItem "Derecha"
    End With
    With Me.OLEObjects("AppGAM").Object
        .Clear
        .AddItem ""
        .AddItem "Remonte"
    End With
    With Me.OLEObjects("TipoGAM").Object
        .Clear
        .AddItem ""
        .AddItem "0"
        .AddItem "2"
    End With
    '---------------------------------------------------------------
    With Me.OLEObjects("TensionDM").Object
        .Clear
        .AddItem ""
        .AddItem "24"
        .AddItem "36"
    End With
    With Me.OLEObjects("DireccionDM").Object
        .Clear
        .AddItem ""
        .AddItem "Central"
        .AddItem "Izquierda"
        .AddItem "Derecha"
    End With
    With Me.OLEObjects("AppDM").Object
        .Clear
        .AddItem ""
        .AddItem "Proteccion"
    End With
    '---------------------------------------------------------------
    With Me.OLEObjects("TensionGBC").Object
        .Clear
        .AddItem ""
        .AddItem "24"
        .AddItem "36"
    End With
    With Me.OLEObjects("DireccionGBC").Object
        .Clear
        .AddItem ""
        
        .AddItem "Izquierda"
        .AddItem "Derecha"
    End With
    With Me.OLEObjects("AppGBC").Object
        .Clear
        .AddItem ""
        .AddItem "Medida"
    End With
    With Me.OLEObjects("TipoGBC").Object
        .Clear
        .AddItem ""
        .AddItem "A"
        .AddItem "B" 
        .AddItem "C"
        .AddItem "D"
    End With
    
End Sub
Sub MoverBloque(ref As String, ParamArray BD() As Variant)
    Dim origen As Range
    Dim destino As Range
    Dim wsDatos As Worksheet
    Dim wsActual As Worksheet
    Dim i As Long

    Set wsActual = ActiveSheet
    Set destino = wsActual.Range("JK97")

    On Error Resume Next
    For i = LBound(BD) To UBound(BD)
        Set wsDatos = ThisWorkbook.Sheets(BD(i))
        Set origen = wsDatos.Range(ref)

        If Not origen Is Nothing Then
            origen.Copy Destination:=destino
            'wsActual.Range(destino, destino.Offset(origen.Rows.Count - 1, origen.Columns.Count - 1)).Select
            Exit Sub
        End If
        Set origen = Nothing
    Next i
    On Error GoTo 0

    MsgBox "Referencia '" & ref & "' no encontrada en ninguna base de datos.", vbCritical
End Sub
Public Sub Boton1IM()
    ThisWorkbook.Sheets("Layout").Range("JI96:LI187").Clear
    ThisWorkbook.Sheets("Layout").Range("JI94:LI187").Interior.Color = &HF2F2F2
    Dim Tension as String
    Dim Direcccion As String
    Dim Aplicacion As String
    Select case Me.OLEObjects("DirreccionIM").Object.Value
        Case "Central"
            Direcccion = "CTR"
        Case "Izquierda"
            Direcccion = "IZ"
        Case "Derecha"
            Direcccion = "DR"
        Case Else
    End Select
    Select case Me.OLEObjects("AppIM").Object.Value
        Case "Seccionamiento"
            Aplicacion = "SCC"
        Case Else
    End Select
    Tension = Me.OLEObjects("TensionIM").Object.Value
    resultadoUNIF = "UN_IM" & Tension & "X" & Aplicacion & "X" & Direcccion
    'MsgBox resultadoUNIF
    MoverBloque resultadoUNIF, "UNIM" , "UIM 36kv"
    
End Sub
Public Sub Boton1QM()
    ThisWorkbook.Sheets("Layout").Range("JI96:LI187").Clear
    ThisWorkbook.Sheets("Layout").Range("JI94:LI187").Interior.Color = &HF2F2F2
    Dim Tension as String
    Dim Direcccion As String
    Dim Aplicacion As String
    Select case Me.OLEObjects("DireccionQM").Object.Value
        Case "Central"
            Direcccion = "CTR"
        Case "Izquierda"
            Direcccion = "IZ"
        Case "Derecha"
            Direcccion = "DR"
        Case Else
    End Select 
    Select case Me.OLEObjects("AppQM").Object.Value
        Case "Proteccion"
            Aplicacion = "PRT"
        Case Else
    End Select
    Tension = Me.OLEObjects("TensionQM").Object.Value
    resultadoUNIF = "UN_QM" & Tension & "X" & Aplicacion & "X" & Direcccion
    'MsgBox resultadoUNIF
    MoverBloque resultadoUNIF, "UNQM" , "UQM 36kv"
End Sub
Public Sub Boton1GAM()
    ThisWorkbook.Sheets("Layout").Range("JI96:LI187").Clear
    ThisWorkbook.Sheets("Layout").Range("JI94:LI187").Interior.Color = &HF2F2F2
    Dim Tension as String
    Dim Direcccion As String
    Dim Aplicacion As String
    Dim Tipo As String

    Select case Me.OLEObjects("DireccionGAM").Object.Value
        Case "Central"
            Direcccion = "CTR"
        Case "Izquierda"
            Direcccion = "IZ"
        Case "Derecha"
            Direcccion = "DR"
        Case Else
    End Select
    Select case Me.OLEObjects("AppGAM").Object.Value
        Case "Remonte"
            Aplicacion = "RMT"
        Case Else
    End Select
    Tipo = Me.OLEObjects("TipoGAM").Object.Value
    Tension = Me.OLEObjects("TensionGAM").Object.Value
    resultadoUNIF = "UN_GAM" & Tipo & Tension & "X" & Aplicacion & "X" & Direcccion
    'MsgBox resultadoUNIF
    MoverBloque resultadoUNIF, "UNGAM" , "UGAM 24kv", "UNGAM 36kv"
End Sub
Public Sub Boton1DM()
    ThisWorkbook.Sheets("Layout").Range("JI96:LI187").Clear
    ThisWorkbook.Sheets("Layout").Range("JI94:LI187").Interior.Color = &HF2F2F2
    Dim Tension as String
    Dim Direcccion As String
    Dim Aplicacion As String
    Dim Especial As String
    Select case Me.OLEObjects("DireccionDM").Object.Value
        Case "Central"
            Direcccion = "CTR"
        Case "Izquierda"
            Direcccion = "IZ"
        Case "Derecha"
            Direcccion = "DR"
        Case Else
    End Select
    Select case Me.OLEObjects("AppDM").Object.Value
        Case "Proteccion"
            Aplicacion = "PRT"
        Case Else
    End Select
    Select case Me.OLEObjects("EspecialDM").Object.Value
        Case "2CT"
            Especial = "2CT"
        Case "CT-PT"
            Especial = "CTPT"
        Case "NA"
            Especial = "NA"
        Case Else
    End Select
    Tension = Me.OLEObjects("TensionDM").Object.Value

    resultadoUNIF = "UN_DM1A" & Tension & "X" & Aplicacion & "X" & Direcccion & "X" & Especial
    'MsgBox resultadoUNIF
    MoverBloque resultadoUNIF, "UNDM1" , "UNDM1 36kv"
End Sub
Public Sub Boton1GBC()
    ThisWorkbook.Sheets("Layout").Range("JI96:LI187").Clear
    ThisWorkbook.Sheets("Layout").Range("JI94:LI187").Interior.Color = &HF2F2F2
    Dim Tension as String
    Dim Direcccion As String
    Dim Aplicacion As String
    Dim Medidores As String
    Dim Tipo As String
    Select case Me.OLEObjects("DireccionGBC").Object.Value
        Case "Central"
            Direcccion = "CTR"
        Case "Izquierda"
            Direcccion = "IZ"
        Case "Derecha"
            Direcccion = "DR" 
        Case Else
    End Select
    Select case Me.OLEObjects("AppGBC").Object.Value
        Case "Medida"
            Aplicacion = "MD"
        Case Else
    End Select
    Select case Me.OLEObjects("MedidoresGBC").Object.Value
        Case "1M"
            Medidores = "1M"
        Case "2M"
            Medidores = "2M"
        Case Else
    End Select
    Tension = Me.OLEObjects("TensionGBC").Object.Value
    Tipo = Me.OLEObjects("TipoGBC").Object.Value    
    resultadoUNIF = "UN_GBC" & Tipo & Tension & "X" & Aplicacion & "X" & Medidores & "X" & Direcccion 
    'MsgBox resultadoUNIF
    MoverBloque resultadoUNIF, "UNGBC" , "UGBC 36kv"
End Sub
