
'=============Hoja configurador=============
Public Seleccion As String

Private Sub ClaseSensores_Change()
    Dim clase As String
    clase = Me.OLEObjects("ClaseSensores").Object.Value
    With Me.OLEObjects("Sensor").Object
        .Clear
        Select Case clase
            Case "TH110"
                .AddItem "Kit de sensores TH110 x3"
            Case "CL110"
                .AddItem "Sensor CL110"
        End Select
    End With
End Sub

Public Sub Worksheet_activate() 
    Nivel_Tension = thisWorkbook.Sheets("Portada 2").range("DA112").Value
    'msgBox Nivel_Tension
    With Me.OLEObjects("Contactor").Object
        .Clear
        .AddItem "Contactos seccionador Abierto - Cerrado - A tierra"
        
    End With
    With Me.OLEObjects("ClaseSensores").Object
        .Clear
        .AddItem "TH110"
        .AddItem "CL110"
    End With
    With Me.OLEObjects("VPIS").Object
        .Clear
        select case Nivel_Tension
            case "24KV"
                .AddItem "VPIS (2 - 4) KV"
                .AddItem "VPIS (3,4 - 6,3)KV"
            case "36KV"
                .AddItem "VPIS (2 - 4) KV"
                .AddItem "VPIS (13 - 24)KV"
            End Select
    End With
    With Me.OLEObjects("RELE").Object
        .Clear
        .AddItem "Rele P3U30-6AAA2BCAA"
        .AddItem "Rele P3T32-CGG1A-AA1NA-BBAAA"
        .AddItem "Rele P5U20-AABD-IABAA-BAEI"
    End With
    With Me.OLEObjects("Modulo").Object
        .Clear
        .AddItem "AGC-150"
    End With
    With Me.OLEObjects("TRACORRIENTE").Object
        .Clear
        select case Nivel_Tension
            case "24KV"
                .AddItem "50/5A Clase 0.5S 10VA 17,5kV"
                .AddItem "100/5A Clase 0.5S 10VA 17,5kV"
                .AddItem "150/5A Clase 0.5S 10VA 17,5kV"
                .AddItem "200/5A Clase 0.5S 10VA 17,5kV"
                .AddItem "50/5A Clase 0.2S 10VA 17,5kV"
                .AddItem "100/5A Clase 0.2S 10VA 17,5kV"
                .AddItem "150/5A Clase 0.2S 10VA 17,5kV"
                .AddItem "200/5A Clase 0.2S 10VA 17,5kV"
                .AddItem "50/5A Clase 5P20 10VA 17,5kV"
                .AddItem "100/5A Clase 5P20 10VA 17,5kV"
                .AddItem "150/5A Clase 5P20 10VA 17,5kV"
                .AddItem "200/5A Clase 5P20 10VA 17,5kV"
                .AddItem "50/5A Clase 10P20 10VA 17,5kV"
                .AddItem "150/5A Clase 10P20 10VA 17,5kV"
                .AddItem "200/5A Clase 10P20 10VA 17,5kV"
                .AddItem "100/5A N1: 0.5S N2:5P20 10VA 17,5kV"
                .AddItem "150/5A N1: 0.5S N2:5P20 10VA 17,5kV"
                .AddItem "200/5A N1: 0.5S N2:5P20 10VA 17,5kV"
                .AddItem "50/5A N1: 0.5S N2:10P20 10VA 17,5kV"
                .AddItem "100/5A N1: 0.5S N2:10P20 10VA 17,5kV"
                .AddItem "150/5A N1: 0.5S N2:10P20 10VA 17,5kV"
                .AddItem "200/5A N1: 0.5S N2:10P20 10VA 17,5kV"
            Case "36KV"
                .AddItem "50/5A Clase 0.5S 10VA 36kV"
                .AddItem "100/5A Clase 0.5S 10VA 36kV"
                .AddItem "150/5A Clase 0.5S 10VA 36kV"
                .AddItem "200/5A Clase 0.5S 10VA 36kV"
                .AddItem "50/5A Clase 0.2S 10VA 36kV"
                .AddItem "100/5A Clase 0.2S 10VA 36kV"
                .AddItem "150/5A Clase 0.2S 10VA 36kV"
                .AddItem "200/5A Clase 0.2S 10VA 36kV"
                .AddItem "50/5A Clase 5P20 10VA 36kV"
                .AddItem "100/5A Clase 5P20 10VA 36kV"
                .AddItem "150/5A Clase 5P20 10VA 36kV"
                .AddItem "200/5A Clase 5P20 10VA 36kV"
                .AddItem "50/5A Clase 10P20 10VA 36kV"
                .AddItem "150/5A Clase 10P20 10VA 36kV"
                .AddItem "200/5A Clase 10P20 10VA 36kV"
                .AddItem "100/5A N1: 0.5S N2:5P20 10VA 36kV"
                .AddItem "150/5A N1: 0.5S N2:5P20 10VA 36kV"
                .AddItem "200/5A N1: 0.5S N2:5P20 10VA 36kV"
                .AddItem "50/5A N1: 0.5S N2:10P20 10VA 36kV"
                .AddItem "100/5A N1: 0.5S N2:10P20 10VA 36kV"
                .AddItem "150/5A N1: 0.5S N2:10P20 10VA 36kV"
                .AddItem "200/5A N1: 0.5S N2:10P20 10VA 36kV"
            End Select
    End With
    With Me.OLEObjects("TRAPOTENCIA").Object
        .Clear
        select case Nivel_Tension
            case "24KV"
                .AddItem "PT Clase 0.2 10VA S.Normal 17,5kV"
                .AddItem "PT Clase 0.5 10VA S.Normal 17,5kV"
                .AddItem "PT Clase 3P 10VA S.Normal 17,5kV"
                .AddItem "PT N1: 0.5 N2:3P 10VA S.Normal 17,5kV"
                .AddItem "PT N1:Clase 0.2 N2: 3P 10VA S.Normal 17,5kV"
            case "36KV"
                .AddItem "PT Clase 0.2 10VA S.Normal 36kV"
                .AddItem "PT Clase 0.5 10VA S.Normal 36kV"
                .AddItem "PT Clase 3P 10VA S.Normal 36kV"
                .AddItem "PT N1: 0.5 N2:3P 10VA S.Normal 36kV"
                .AddItem "PT N1:Clase 0.2 N2: 3P 10VA S.Normal 36kV"
            End Select
    End With
    With Me.OLEObjects("FUSIBLE").Object
        .Clear
        select case Nivel_Tension
            case "24KV"
            .AddItem "6.3 Amp 10-24kV"
            .AddItem "10 Amp 10-24kV"
            .AddItem "16 Amp 10-24kV"
            .AddItem "20 Amp 10-24kV"
            .AddItem "25 Amp 10-24kV"
            .AddItem "31,5 Amp 10-24kV"
            .AddItem "40 Amp 10-24kV"
            .AddItem "50 Amp 10-24kV"
            .AddItem "63 Amp 10-24kV"
            .AddItem "80 Amp 10-24kV"
            .AddItem "100 Amp 10-24kV"
            .AddItem "125 Amp 10-24kV"
            .AddItem "160 Amp 10-24kV"
            .AddItem "200 Amp 10-24kV"
        Case "36KV"
            .AddItem "6.3 Amp 20-36kV"
            .AddItem "110 Amp 20-36kV"
            .AddItem "16 Amp 20-36kV"
            .AddItem "20 Amp 20-36kV"
            .AddItem "25 Amp 20-36kV"
            .AddItem "31,5 Amp 20-36kV"
            .AddItem "40 Amp 20-36kV"
            .AddItem "50 Amp 20-36kV"
            .AddItem "63 Amp 20-36kV"
            .AddItem "80 Amp 20-36kV"
            .AddItem "100 Amp 20-36kV"
        end select      
    End With
    With Me.OLEObjects("AnalizadorREDES").Object
        .Clear
        .AddItem "PM5340"
        .AddItem "PM5560"
        .AddItem "PM8240"
    End With
    With Me.OLEObjects("MedidorTarifario").Object
        .Clear
        .AddItem "Landis Clase 0.5s RS485"
        .AddItem "Landis Clase 0.5s RS485 + Ethernet"
        .AddItem "Landis Clase 0.2s RS485 + Ethernet"
        .AddItem "Microstar Clase 0.5s RS485"
        .AddItem "Inhemeter Clase 0.5s RS485"
        .AddItem "Itron Clase 0.5s RS485"
        .AddItem "Itron Clase 0.2s RS485"
    End With
    With Me.OLEObjects("GabineteBT").Object
        .Clear
        .AddItem "Gabinete de baja tension"
    End With
    With Me.OLEObjects("MandoMotorizado").Object
        .Clear
        .AddItem "Motorizado 120VAC IM"
        .AddItem "Motorizado 120VAC QM"
    End With
    With Me.OLEObjects("Gateway").Object
        .Clear
        .AddItem "PAS600"
        .AddItem "PAS600L"
    End With
    With Me.OLEObjects("AdaptadorTension").Object
        .Clear
        .AddItem "Adaptador de tension"
    End With
    With Me.OLEObjects("Zocalo").Object
        .Clear
        .AddItem "Zocalos 400mm"
    End With
End Sub
 
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim wsDB As Worksheet
    Dim listaOpciones As String
    Dim ultimaFila As Long
    Dim i As Long
    Dim columnas As Variant
    Dim col As Variant
    Dim celdaDestino As Range
    
    ' Columnas a monitorear
    columnas = Array("CI", "CZ", "DQ", "EH", "EY", "FP", "GG", "GX", "HO", "IF")
    
    ' Verificar si el cambio fue en fila 69 y en alguna de las columnas
    If Target.Row <> 69 Then Exit Sub
    
    Dim esColumnaValida As Boolean
    esColumnaValida = False
    For Each col In columnas
        If Target.Column = Range(col & "1").Column Then
            esColumnaValida = True
            Exit For
        End If
    Next col
    
    If Not esColumnaValida Then Exit Sub
    
    ' Configurar
    Set wsDB = ThisWorkbook.Sheets("Hoja2")
    ultimaFila = wsDB.Cells(wsDB.Rows.Count, "A").End(xlUp).Row
    
    ' Celda destino en la misma columna, fila 72
    Set celdaDestino = Me.Cells(72, Target.Column)
    
    ' Limpiar destino
    celdaDestino.Value = ""
    On Error Resume Next
    'celdaDestino.Validation.Delete
    On Error GoTo 0
    
End Sub

Public Sub Codigo_Contactores(descripcion As String, Fila As Long)
    ' Asegurar que Nivel_Tension tenga valor
    If Nivel_Tension = "" Then
        Nivel_Tension = UCase(Trim(ThisWorkbook.Sheets("Portada 2").Range("DA112").Value))
    End If 
    Dim HojaDestino As Worksheet
    Dim HojaOrigen As Worksheet
    Dim codigo As String
    Dim ultimaFila As Long
    Dim i As Long 
    Dim numeroCelda As Variant
    Dim columnas As Variant
    Dim colBusqDesc As String
    Dim colBusqCod As String
    Dim encontrado As Boolean
    'MsgBox "Nivel_Tension = [" & Nivel_Tension & "]"
    ' Columnas destino en fila 114
    columnas = Array("CI", "CZ", "DQ", "EH", "EY", "FP", "GG", "GX", "HO", "IF")
    
    Set HojaDestino = ThisWorkbook.Sheets("CONFIGURADOR")
    Set HojaOrigen = ThisWorkbook.Sheets("Hoja2")
        
    If descripcion = "" Then
        MsgBox "Seleccione un contactor del ComboBox.", vbExclamation
        Exit Sub
    End If
    
    ' Definir columnas de búsqueda según nivel de tensión
    Select Case Nivel_Tension
        Case "24KV"
            colBusqDesc = "AE"
            colBusqCod = "AF"
        ' Agrega más casos si hay otros niveles:
        Case "36KV"
            colBusqDesc = "AI"
            colBusqCod = "AJ"
        Case Else
            MsgBox "Nivel de tensión no válido: " & Nivel_Tension, vbExclamation
            Exit Sub
    End Select
    
    ' Buscar la descripción en Hoja2
    ultimaFila = HojaOrigen.Cells(HojaOrigen.Rows.Count, colBusqDesc).End(xlUp).Row
    encontrado = False
    
    For i = 2 To ultimaFila
        If Trim(HojaOrigen.Range(colBusqDesc & i).Value) = descripcion Then
            codigo = HojaOrigen.Range(colBusqCod & i).Value
            encontrado = True
            Exit For
        End If
    Next i
    
    If Not encontrado Then
        MsgBox "No se encontró el contactor: " & descripcion, vbExclamation
        Exit Sub
    End If
    
    ' Preguntar número de celda (1-10)
    numeroCelda = Application.InputBox( _
        Prompt:="Ingrese el numero +H de la celda:" & vbCrLf & vbCrLf , _
        Title:="Seleccionar celda destino", _
        Type:=1)  ' Type 1 = solo números
    
    ' Validar entrada
    If numeroCelda = False Then Exit Sub  ' Usuario canceló
    
    If numeroCelda < 1 Or numeroCelda > 10 Then
        MsgBox "Número no válido. Debe ser entre 1 y 10.", vbExclamation
        Exit Sub
    End If
    
    ' Pegar el código en la celda correspondiente
    HojaDestino.Range(columnas(numeroCelda - 1) & Fila).NumberFormat = "@"
    HojaDestino.Range(columnas(numeroCelda - 1) & Fila).Value = codigo
    
    'MsgBox "Código '" & codigo & "' copiado en celda " & numeroCelda & _
     '      " (columna " & columnas(numeroCelda - 1) & fila).", vbInformation

End Sub

public sub Generara_Contactor
    Seleccion = Me.OLEObjects("Contactor").Object.Value
    codigo_Contactores Seleccion, 114
End Sub
public sub Generar_Sensores()
    Seleccion = Me.OLEObjects("Sensor").Object.Value
    IF Me.OLEObjects("ClaseSensores").Object.Value = "TH110" Then
        codigo_Contactores Seleccion, 118
    Elseif Me.OLEObjects("ClaseSensores").Object.Value = "CL110" Then
        codigo_Contactores Seleccion, 122
    End If
End Sub
Public sub Generar_VPIS()
    Seleccion = Me.OLEObjects("VPIS").Object.Value
    codigo_Contactores Seleccion, 126
End sub
Public sub Generar_RELE()
    Seleccion = Me.OLEObjects("RELE").Object.Value
    codigo_Contactores Seleccion, 130
End sub

Public Sub Generar_Modulo()
    Seleccion = Me.OLEObjects("Modulo").Object.Value
    codigo_Contactores Seleccion, 134
End Sub
Public Sub Generar_TRACORRIENTE()
    Seleccion = Me.OLEObjects("TRACORRIENTE").Object.Value
    codigo_Contactores Seleccion, 138
End Sub
Public Sub Generar_TRAPOTENCIA()
    Seleccion = Me.OLEObjects("TRAPOTENCIA").Object.Value
    codigo_Contactores Seleccion, 142
End Sub
Public Sub Generar_FUSIBLE()
    Seleccion = Me.OLEObjects("FUSIBLE").Object.Value
    codigo_Contactores Seleccion, 146
End Sub
PUBLIC Sub Generar_AnalizadorREDES()
    Seleccion = Me.OLEObjects("AnalizadorREDES").Object.Value
    codigo_Contactores Seleccion, 150
End Sub
public sub Generar_MedidorTarifario()
    Seleccion = Me.OLEObjects("MedidorTarifario").Object.Value
    codigo_Contactores Seleccion, 154
End sub
public sub Generar_GabineteBT()
    Seleccion = Me.OLEObjects("GabineteBT").Object.Value
    codigo_Contactores Seleccion, 158
End sub
public sub Generar_MandoMotorizado()
    Seleccion = Me.OLEObjects("MandoMotorizado").Object.Value
    codigo_Contactores Seleccion, 162
End sub
public sub Generar_AdaptadorTension()
    Seleccion = Me.OLEObjects("AdaptadorTension").Object.Value
    codigo_Contactores Seleccion, 166
End sub
public sub Generar_Zocalo()
    Seleccion = Me.OLEObjects("Zocalo").Object.Value
    codigo_Contactores Seleccion, 170
End sub

