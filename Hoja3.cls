'============= Hoja layout============

'================= VARIABLES GLOBALES ===========================
Dim resultadoMeta As String
Dim resultadoacotacion As String
''--------------------------------CARGA DE COMBOS AL ACTIVAR LA HOJA-------------------------------
Private Sub NivelTensionGAM_Change()
   DiM SeleccionNiveltensionGam as string
   SeleccionNiveltensionGam = Me.OLEObjects("NivelTensionGAM").Object.Value
   With Me.OLEObjects("TipoGAM").Object
        .Clear
        Select case SeleccionNiveltensionGam
            Case "24"
                .AddItem "GAM_0"
                .AddItem "GAM_2"
            Case "36"
                .AddItem "GAM_0 sin baraje"
                .AddItem "GAM_0 con baraje"
                .AddItem "GAM_2 con baraje"
        End Select
    End With
End Sub
Private sUB NivelTensionDM_Change()
   DiM SeleccionNiveltensionDM as string
   SeleccionNiveltensionDM = Me.OLEObjects("NivelTensionDM").Object.Value
   With Me.OLEObjects("TipoDM").Object
        .Clear
        Select case SeleccionNiveltensionDM
            Case "24"
                .AddItem "DM1-A"
                .AddItem "DM1-A Especial"
            Case "36"
                .AddItem "DM1-A"
                
        End Select
    End With
End Sub
Public Sub Worksheet_Activate()
    '---------------------------------------------------------------
    With Me.OLEObjects("NivelTensionIM").Object
        .Clear
        .AddItem ""
        .AddItem "24"
        .AddItem "36"
    End With
    With Me.OLEObjects("AppIM").Object
        .Clear
        .AddItem ""
        .AddItem "Seccionamiento"
    End With
    '---------------------------------------------------------------
    With Me.OLEObjects("NivelTensionQM").Object
        .Clear
        .AddItem ""
        .AddItem "24"
        .AddItem "36"
    End With
    With Me.OLEObjects("AppQM").Object
        .Clear
        .AddItem ""
        .AddItem "Proteccion"
    End With
    '---------------------------------------------------------------
    With Me.OLEObjects("NivelTensionGAM").Object
        .Clear
        .AddItem ""
        .AddItem "24"
        .AddItem "36"
    End With
    With Me.OLEObjects("AppGAM").Object
        .Clear
        .AddItem ""
        .AddItem "Remonte"
    End With
    '---------------------------------------------------------------
    With Me.OLEObjects("NivelTensionDM").Object
        .Clear
        .AddItem ""
        .AddItem "24"
        .AddItem "36"
    End With
    With Me.OLEObjects("AppDM").Object
        .Clear
        .AddItem ""
        .AddItem "Proteccion"
    End With
    '---------------------------------------------------------------
    With Me.OLEObjects("NivelTensionGBC").Object
        .Clear
        .AddItem ""
        .AddItem "24"
        .AddItem "36"
    End With
    With Me.OLEObjects("AppGBC").Object
        .Clear
        .AddItem ""
        .AddItem "Medida"
    End With
    With Me.OLEObjects("TipoGBC").Object
        .Clear
        .AddItem ""
        .AddItem "GBC-A"
        .AddItem "GBC-B"
        .AddItem "GBC-C"
        .AddItem "GBC-D"
    End With
    With Me.OLEObjects("OrientacionGBC").Object
        .Clear
        .AddItem ""
        .AddItem "Derecha a izquierda"
        .AddItem "Izquierda a derecha"
    End With
    '---------------------------------------------------------------
    With Me.OLEObjects("TIPODECELDA").Object
        .Clear
        .AddItem "IM"
        .AddItem "QM"
        .AddItem "GAM"
        .AddItem "DM1"
        .AddItem "GBC"
    End With
    With Me.OLEObjects("TENSIONCELDA").Object
        .Clear
        .AddItem "24"
        .AddItem "36"
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
            wsActual.Range(destino, destino.Offset(origen.Rows.Count - 1, origen.Columns.Count - 1)).Select
            Exit Sub
        End If
        Set origen = Nothing
    Next i
    On Error GoTo 0

    MsgBox "Referencia '" & ref & "' no encontrada en ninguna base de datos.", vbCritical
End Sub
Sub MoverBloque2(ref As String, ParamArray BD() As Variant)
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
    Dim Tension As String
    Dim Aplicacion As String
    Select Case Me.OLEObjects("AppIM").Object.Value
        Case "Seccionamiento"
            Aplicacion="SECC"
        Case Else
    End Select
    Tension = Me.OLEObjects("NivelTensionIM").Object.Value
    resultadoMeta = "IM_" & Tension & "X" & Aplicacion
    'MsgBox "Referencia a buscar: " & resultadoMeta
    MoverBloque resultadoMeta, "IM"
End Sub
Public Sub Boton2QM()
    ThisWorkbook.Sheets("Layout").Range("JI96:LI187").Clear
    ThisWorkbook.Sheets("Layout").Range("JI94:LI187").Interior.Color = &HF2F2F2
    Dim Tension As String
    Dim Aplicacion As String
    Select Case Me.OLEObjects("AppQM").Object.Value
        Case "Proteccion"
            Aplicacion="PROT"
        Case Else
    End Select
    Tension = Me.OLEObjects("NivelTensionQM").Object.Value
    resultadoMeta = "QM_" & Tension & "X" & Aplicacion
    'MsgBox "Referencia a buscar: " & resultadoMeta
    MoverBloque resultadoMeta, "QM"
End Sub
Public Sub Boton3GAM()
    ThisWorkbook.Sheets("Layout").Range("JI96:LI187").Clear
    ThisWorkbook.Sheets("Layout").Range("JI94:LI187").Interior.Color = &HF2F2F2
    Dim Tension As String
    Dim Aplicacion As String
    Dim Tipo As String 
    Select Case Me.OLEObjects("AppGAM").Object.Value
        Case "Remonte"
            Aplicacion="REM"
        Case Else
    End Select 
    Select Case me.OLEObjects("TipoGAM").Object.Value
        Case "GAM_0"
            Tipo="0"
        Case "GAM_2"
            Tipo="2"
        Case "GAM_0 con baraje"
            Tipo="0BRR"
        Case "GAM_2 con baraje"
            Tipo="2BRR"
        Case "GAM_0 sin baraje"
            Tipo="0"
    End Select
    Tension = Me.OLEObjects("NivelTensionGAM").Object.Value
   
    resultadoMeta = "GAM_" & Tension & "X" & Aplicacion & "X" & Tipo
    'MsgBox "Referencia a buscar: " & resultadoMeta
    MoverBloque resultadoMeta, "GAM"
End Sub
Public Sub Boton4DM()
    ThisWorkbook.Sheets("Layout").Range("JI96:LI187").Clear
    ThisWorkbook.Sheets("Layout").Range("JI94:LI187").Interior.Color = &HF2F2F2
    Dim Tension As String
    Dim Aplicacion As String
    Dim Tipo As String
    Select Case Me.OLEObjects("AppDM").Object.Value
        Case "Proteccion"
            Aplicacion="PROT"
        Case Else
    End Select
    Select Case me.OLEObjects("TipoDM").Object.Value
        case "DM1-A"
            Tipo="A"
        case "DM1-A Especial"
            Tipo="AESP"
    End Select
    Tension = Me.OLEObjects("NivelTensionDM").Object.Value
   
    resultadoMeta = "DM_" & Tension & "X" & Aplicacion & "X" & Tipo
    'MsgBox "Referencia a buscar: " & resultadoMeta
    MoverBloque resultadoMeta, "DM1"
End Sub
Public Sub Boton5GBC()
    ThisWorkbook.Sheets("Layout").Range("JI96:LI187").Clear
    ThisWorkbook.Sheets("Layout").Range("JI94:LI187").Interior.Color = &HF2F2F2
    Dim Tension As String
    Dim Aplicacion As String
    Dim Tipo As String
    Dim Orientacion As String
    Dim Medidores As String
    Select Case Me.OLEObjects("AppGBC").Object.Value
        Case "Medida"
            Aplicacion="MD"
        Case Else
    End Select
    Select Case me.OLEObjects("TipoGBC").Object.Value
        Case "GBC-A"
            Tipo="A"
        Case "GBC-B"
            Tipo="B"
        Case "GBC-C"
            Tipo="C"
        Case "GBC-D"
            Tipo="D"
        Case Else
    End Select
    Select Case me.OLEObjects("OrientacionGBC").Object.Value
        Case "Derecha a izquierda"
            Orientacion="DI"
        Case "Izquierda a derecha"
            Orientacion="ID"
        Case Else
    End Select
    Tension = Me.OLEObjects("NivelTensionGBC").Object.Value

    resultadoMeta = "GBC_" & Tension & "X" & Aplicacion & "X" & Tipo & "X" & Orientacion
    'MsgBox "Referencia a buscar: " & resultadoMeta
    MoverBloque resultadoMeta, "GBC 24kV 1 MED","GBC 24kV 2 MED","GBC 36kV 1 MED","GBC 36kV 2 MED"
End Sub

Public Sub Boton6ACOTAR()
    ThisWorkbook.Sheets("Layout").Range("JI96:LI187").Clear
    ThisWorkbook.Sheets("Layout").Range("JI94:LI187").Interior.Color = &HF2F2F2
    DiM TIPODECELDA As String
    DIM TENSIONCELDA As String    
    TIPODECELDA = Me.OLEObjects("TIPODECELDA").Object.Value
    TENSIONCELDA = Me.OLEObjects("TENSIONCELDA").Object.Value
    
    if TIPODECELDA = "IM" and TENSIONCELDA = "24" then
        resultadoacotacion= "AC_" & "1600X375X940"
    ElseIf TIPODECELDA = "IM" and TENSIONCELDA = "36" Then
        resultadoacotacion= "AC_" & "2250X375X1500"
    ElseIf TIPODECELDA = "QM" and TENSIONCELDA = "24" Then 
        resultadoacotacion= "AC_" & "1600X375X940"  
    ElseIf TIPODECELDA = "QM" and TENSIONCELDA = "36" Then 
        resultadoacotacion= "AC_" & "2250X375X1500" 
    ElseIf TIPODECELDA = "GAM" and TENSIONCELDA = "24" Then  
        resultadoacotacion= "AC_" & "1600X375X940" 
    ElseIf TIPODECELDA = "GAM" and TENSIONCELDA = "36" Then  
        resultadoacotacion= "AC_" & "2250X375X1500"
    ElseIf TIPODECELDA = "DM1" and TENSIONCELDA = "24" Then  
        resultadoacotacion= "AC_" & "1700X750X1140" 
    ElseIf TIPODECELDA = "DM1" and TENSIONCELDA = "36" Then 
        resultadoacotacion= "AC_" & "2250X750X1800"
    ElseIf TIPODECELDA = "GBC" and TENSIONCELDA = "24" Then 
        resultadoacotacion= "AC_" & "2050X750X1020"
    ElseIf TIPODECELDA = "GBC" and TENSIONCELDA = "36" Then 
        resultadoacotacion= "AC_" & "2250X750X1750"      
    end if
    MoverBloque2 resultadoacotacion, "IM"
End Sub
Sub AgregarELEMENTOSALBOM(codigoBuscado As String)

    Dim hojaDestino As Worksheet
    Dim hojaDatos As Worksheet
    Dim hojasDatos As Variant
    Dim nombreHoja As Variant
    Dim celdaCodigo As Range
    Dim filaCodigo As Long
    Dim referencia As String, tipo As String
    Dim Dato1 As String, Dato2 As String, Dato3 As String, Dato4 As String, Dato5 As String, Dato6 As String
    Dim detalles As String
    Dim ultimaFila As Long
    Dim encontrado As Boolean
    Dim valorN As String, valorC As String, valorTag As String, Ubicacion As String
    Dim filaDestino As Long

    hojasDatos = Array( _
    "IM", _
    "QM", _
    "GAM", _
    "DM1", _
    "GBC 24kV 1 MED", _
    "GBC 24kV 2 MED", _
    "GBC 36kV 1 MED", _
    "GBC 36kV 2 MED", _
    "UNIM", _
    "UNQM", _
    "UNGAM", _
    "UNDM1", _
    "UNGBC")


    Set hojaDestino = ThisWorkbook.Sheets("BOM")
    encontrado = False

    If codigoBuscado = "" Then
        MsgBox "No se ha generado ningún código.", vbExclamation
        Exit Sub
    End If

    ' Buscar en cada hoja
    For Each nombreHoja In hojasDatos
        Set hojaDatos = ThisWorkbook.Sheets(nombreHoja)
        
        Set celdaCodigo = hojaDatos.Range("B:B").Find(What:=codigoBuscado, LookIn:=xlValues, LookAt:=xlWhole)
        
        If Not celdaCodigo Is Nothing Then
            filaCodigo = celdaCodigo.Row

            referencia = hojaDatos.Cells(filaCodigo, "C").Value
            tipo = hojaDatos.Cells(filaCodigo, "D").Value

            ' Extraer directamente los datos
            Dato1 = hojaDatos.Cells(filaCodigo, "E").Value
            Dato2 = hojaDatos.Cells(filaCodigo, "F").Value
            Dato3 = hojaDatos.Cells(filaCodigo, "G").Value
            Dato4 = hojaDatos.Cells(filaCodigo, "H").Value
            

            detalles = Dato1 & "/ " & Dato2 & "/ " & Dato3 & "/ " & Dato4
            encontrado = True
            Exit For
        End If
    Next nombreHoja
 
    If Not encontrado Then
        MsgBox "Código no encontrado en las hojas de datos.", vbExclamation
        'MsgBox hojaDatos.Name
        Exit Sub
    End If

    ' Pedir datos al usuario
    valorTag = InputBox("Ingrese el TAG la celda", "TAG")
   

    'Ubicacion = "+N" & valorN & " +C" & valorC

    ' Buscar la siguiente fila disponible cada 3 filas a partir de fila 56
    filaDestino = 56
    Do While hojaDestino.Range("BJ" & filaDestino).Text <> "" Or _
              hojaDestino.Range("BR" & filaDestino).Text <> "" Or _
              hojaDestino.Range("CI" & filaDestino).Text <> "" Or _
              hojaDestino.Range("CQ" & filaDestino).Text <> "" Or _
              hojaDestino.Range("DH" & filaDestino).Text <> ""
        filaDestino = filaDestino + 3
    Loop

    ' Escribir los datos en la hoja con el orden de la imagen
    With hojaDestino
        .Range("BJ" & filaDestino).Value = (filaDestino - 56) \ 3 + 1     ' Item
        .Range("BR" & filaDestino).Value = referencia                     ' Referencia
        .Range("CN" & filaDestino).Value = tipo                           ' Aplicacion
        .Range("HF" & filaDestino).Value = valorTag                       ' TAG
        .Range("EP" & filaDestino).Value = detalles                       ' Detalles 
                      
    End With

    MsgBox "Referencia agregada al BOM correctamente.", vbInformation

End Sub

Sub botonBOMM()
    'MsgBox resultado
    AgregarELEMENTOSALBOM resultadoMeta
    
End Sub

