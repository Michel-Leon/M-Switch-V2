VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Configurador_IM 
   ClientHeight    =   8960.001
   ClientLeft      =   -3100
   ClientTop       =   -420
   ClientWidth     =   7900
   OleObjectBlob   =   "Configurador_IM.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Configurador_IM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Initialize()
    ' Tama�o de la ventana
    Me.Height = 650
    Me.Width = 932
     ' Centrar en la pantalla
    Me.StartUpPosition = 0  ' Manual, para poder controlar posici?n
    Me.Top = (Application.Height - Me.Height) / 2
    Me.Left = (Application.Width - Me.Width) / 2
    MsgBox Dato1
    If Resultado_IM = "1" Then
        Me.OptionButton1.Value = True
        Me.OptionButton7.Value = True
        Me.OptionButton9.Value = True
        Me.OptionButton10.Value = True
        Me.OptionButton12.Value = True
        Me.OptionButton15.Value = True
        Me.OptionButton17.Value = True
        Me.OptionButton19.Value = True
        Me.OptionButton23.Value = True
        Me.OptionButton25.Value = True
        Me.OptionButton27.Value = True
        Me.OptionButton30.Value = True
        Me.OptionButton66.Value = True ' conexion a al izquierda
        Me.Rf1.Caption = "SM61N3JHC6Z7CIMES"
        Me.Ct1.Caption = "1"
    End If
    
End Sub
Private Sub SetRef(ByVal isChecked As Boolean, _
                    ByRef refCtl As MSForms.Label, _
                    ByRef ctCtl As MSForms.Label, _
                    ByVal refValue As String, _
                    ByVal ctValue As String)
    If isChecked Then
        refCtl.Caption = refValue
        ctCtl.Caption = ctValue
    Else
        refCtl.Caption = ""
        ctCtl.Caption = ""
    End If
End Sub
'============= A)
Private Sub OptionButton14_Click()
    SetRef OptionButton14.Value, Me.Rf2, Me.Ct2, "KIT-CONT-B-01", "1"
End Sub
Private Sub OptionButton15_Click()
    SetRef OptionButton15.Value, Me.Rf2, Me.Ct2, "", ""
End Sub
'============= B)
Private Sub OptionButton16_Click()
    SetRef OptionButton16.Value, Me.Rf3, Me.Ct3, "KIT-TH110-01", "1"
End Sub
Private Sub OptionButton17_Click()
    SetRef OptionButton17.Value, Me.Rf3, Me.Ct3, "", ""
End Sub
'============= C)
Private Sub OptionButton18_Click()
    SetRef OptionButton18.Value, Me.Rf4, Me.Ct4, "ZBRTT1", "1"
End Sub
Private Sub OptionButton19_Click()
    SetRef OptionButton19.Value, Me.Rf4, Me.Ct4, "", ""
End Sub
'============= D)
Private Sub OptionButton20_Click()
    SetRef OptionButton20.Value, Me.Rf5, Me.Ct5, "VPI62403", "1"
End Sub
Private Sub OptionButton21_Click()
    SetRef OptionButton21.Value, Me.Rf5, Me.Ct5, "VPI62404", "1"
End Sub
Private Sub OptionButton22_Click()
    SetRef OptionButton22.Value, Me.Rf5, Me.Ct5, "VPI62405", "1"
End Sub
Private Sub OptionButton23_Click()
    SetRef OptionButton23.Value, Me.Rf5, Me.Ct5, "", ""
End Sub
'============= E)
Private Sub OptionButton24_Click()
    SetRef OptionButton24.Value, Me.Rf6, Me.Ct6, "GAB-CONT-MT", "1"
End Sub
Private Sub OptionButton25_Click()
    SetRef OptionButton25.Value, Me.Rf6, Me.Ct6, "", ""
End Sub
'============= F)
Private Sub OptionButton26_Click()
    SetRef OptionButton26.Value, Me.Rf7, Me.Ct7, "KIT-MCH-IM-120V", "1"
End Sub
Private Sub OptionButton27_Click()
    SetRef OptionButton27.Value, Me.Rf7, Me.Ct7, "", ""
End Sub
'============= G)
Private Sub OptionButton28_Click()
    SetRef OptionButton28.Value, Me.Rf8, Me.Ct8, "KIT-ZOCALO-400-SM6-24", "1"
End Sub
Private Sub OptionButton29_Click()
    SetRef OptionButton29.Value, Me.Rf8, Me.Ct8, "KIT-ZOCALO-600-SM6-24", "1"
End Sub
Private Sub OptionButton30_Click()
    SetRef OptionButton30.Value, Me.Rf8, Me.Ct8, "", ""
End Sub
Private Sub Registar_Click()
    With ThisWorkbook.Sheets("FT IM")
        '============= 1)
            If Me.OptionButton1.Value = True Then 'Voltaje nominal
                .Range("CZ32").Value = "17.5 kV"
            ElseIf Me.OptionButton2.Value = True Then
                .Range("CZ32").Value = "24 kV"
            End If
        '============= 2)
            If Me.OptionButton3.Value = True Then 'Voltaje de operacion
                .Range("CZ36").Value = "11.4 kV"
            ElseIf Me.OptionButton4.Value = True Then
                .Range("CZ36").Value = "13.2  kV"
            ElseIf Me.OptionButton5.Value = True Then
                .Range("CZ36").Value = "13.8 kV"
            End If
        '============= 3)
            If Me.OptionButton6.Value = True Then 'Corriente
                .Range("CZ40").Value = "630 A"
            ElseIf Me.OptionButton7.Value = True Then
                .Range("CZ40").Value = "1250 A"
            End If
        '============= 4)
            If Me.OptionButton8.Value = True Then 'corriente corto
                .Range("CZ52").Value = "20 kA/1s"
            ElseIf Me.OptionButton9.Value = True Then
                .Range("CZ52").Value = "25 kA/1S"
            End If
        '============= 6)
            If Me.OptionButton12.Value = True Then
                .Range("DA114").Value = "AC-AFL 12,5kA/1s"
            ElseIf Me.OptionButton13.Value = True Then
                .Range("DA114").Value = "IAC-AFLR 12,5kA/1s"
            End If
        '============= A)
            If Me.OptionButton14.Value = True Then
                .Range("FH23").Value = "Incluido"
            ElseIf Me.OptionButton15.Value = True Then
                .Range("FH23").Value = "No Incluido"
            End If
        '============= B)
            If Me.OptionButton16.Value = True Then
                .Range("FH35").Value = "sensores TH110 x3"
            ElseIf Me.OptionButton17.Value = True Then
                .Range("FH35").Value = "No Incluido"
            End If
        '============= C)
            If Me.OptionButton18.Value = True Then
                .Range("FH40").Value = "Sensor CL110"
            ElseIf Me.OptionButton19.Value = True Then
                .Range("FH40").Value = "No Incluido"
            End If
        '============= D)
            If Me.OptionButton20.Value = True Then
                .Range("FH45").Value = "VPIS (2 - 4) KV"
            ElseIf Me.OptionButton21.Value = True Then
                .Range("FH45").Value = "VPIS (3,4 - 6,3)KV - VPIS (13 - 24kV)"
            ElseIf Me.OptionButton22.Value = True Then
                .Range("FH45").Value = "VPIS (4-8)KV"
            ElseIf Me.OptionButton23.Value = True Then
                .Range("FH45").Value = "(9-17)kV"
            End If
        '============= E)
            If Me.OptionButton24.Value = True Then
                .Range("FH50").Value = "Incluido"
            ElseIf Me.OptionButton25.Value = True Then
                .Range("FH50").Value = "No Incluido"
            End If
        '============= F)
            If Me.OptionButton26.Value = True Then
                .Range("FH55").Value = "Motorizado 120VAC IM"
            ElseIf Me.OptionButton27.Value = True Then
                .Range("FH55").Value = "No Incluido"
            End If
        '============= G)
            If Me.OptionButton28.Value = True Then
                .Range("FH60").Value = "400mm"
            ElseIf Me.OptionButton29.Value = True Then
                .Range("FH60").Value = "600mm"
            ElseIf Me.OptionButton30.Value = True Then
                .Range("FH60").Value = "No Incluido"
            End If
    End With
    MsgBox "Datos Cargados a la ficha tecnica", vbInformation, "Aviso"
    Codigo_unifilar
    'cerrar el formulario
    Unload Me
End Sub
Public Sub Codigo_unifilar()
    Dim codigo_IM As String
    
    Debug.Print "Dato1 al entrar: [" & Dato1 & "]"
    
    If Dato1 = "" Then
        MsgBox "Dato1 está vacío. Seleccione primero el tipo (SM624...).", vbExclamation
        Exit Sub
    End If

    Dato2 = "02"
    If Me.OptionButton17.Value = True Then
        Dato3 = "01"
    ElseIf Me.OptionButton16.Value = True Then
        Dato3 = "02"
    End If
    IF Me.OptionButton66.Value = True Then
        Dato4 = "01"
    ElseIf Me.OptionButton67.Value = True Then
        Dato4 = "02"
    ElseIf Me.OptionButton68.Value = True Then
        Dato4 = "03"
    Else
        Dato4 = ""
    End If
    If PosicionUnifilarActual = 0 Then
        MsgBox "No se ha definido la posicion (H1-H9) para este bloque.", vbExclamation
        Exit Sub
    End If

    codigo_IM = "UN_" & Dato1 & Dato2 & Dato3 & Dato4
    Asignar_Posicion_Unifilar PosicionUnifilarActual, codigo_IM, 17
End Sub

