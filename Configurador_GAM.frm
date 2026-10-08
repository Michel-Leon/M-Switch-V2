VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Configurador_GAM 
   Caption         =   "GAM"
   ClientHeight    =   8960.001
   ClientLeft      =   -70
   ClientTop       =   -280
   ClientWidth     =   7900
   OleObjectBlob   =   "Configurador_GAM.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Configurador_GAM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub UserForm_Initialize()
    'Tama?o de la ventana
    Me.Height = 580
    Me.Width = 932
     ' Centrar en la pantalla
    Me.StartUpPosition = 0  ' Manual, para poder controlar posici?n
    Me.Top = (Application.Height - Me.Height) / 2
    Me.Left = (Application.Width - Me.Width) / 2
    Me.OptionButton1.Value = True
    Me.OptionButton7.Value = True
    Me.OptionButton9.Value = True
    Me.OptionButton10.Value = True
    Me.OptionButton12.Value = True
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
    SetRef OptionButton14.Value, Me.Rf2, Me.Ct2, "KIT-TH110-01", "1"
End Sub
Private Sub OptionButton15_Click()
    SetRef OptionButton15.Value, Me.Rf2, Me.Ct2, "", ""
End Sub
'============= B)
Private Sub OptionButton16_Click()
    SetRef OptionButton16.Value, Me.Rf3, Me.Ct3, "ZBRTT1", "1"
End Sub
Private Sub OptionButton17_Click()
    SetRef OptionButton17.Value, Me.Rf3, Me.Ct3, "", ""
End Sub
'============= C)
Private Sub OptionButton18_Click()
    SetRef OptionButton18.Value, Me.Rf4, Me.Ct4, "KIT-ZOCALO-400-SM6-24", "1"
End Sub
Private Sub OptionButton19_Click()
    SetRef OptionButton19.Value, Me.Rf4, Me.Ct4, "KIT-ZOCALO-600-SM6-24", ""
End Sub
Private Sub OptionButton20_Click()
    SetRef OptionButton20.Value, Me.Rf4, Me.Ct4, "", ""
End Sub
'============= 2)
Private Sub OptionButton21_Click()
    SetRef OptionButton21.Value, Me.Rf1, Me.Ct1, "SE-GAM-0-24KV", "1"
End Sub
Private Sub OptionButton22_Click()
    SetRef OptionButton22.Value, Me.Rf1, Me.Ct1, "SE-GAM-2-24KV", "1"
End Sub
Private Sub OptionButton23_Click()
    SetRef OptionButton23.Value, Me.Rf1, Me.Ct1, "SM61DAJHJ6Z7CGAM0", "1"
End Sub
Private Sub OptionButton24_Click()
    SetRef OptionButton24.Value, Me.Rf1, Me.Ct1, "SM61G2JHJ6Z7CGAM2", "1"
End Sub
Private Sub Registar_Click()
    With ThisWorkbook.Sheets("FT GAM 24kV")
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
                .Range("FH28").Value = "Incluido"
            ElseIf Me.OptionButton17.Value = True Then
                .Range("FH28").Value = "No Incluido"
            End If
        '============= C)
            If Me.OptionButton18.Value = True Then
                .Range("FH31").Value = "Incluido"
            ElseIf Me.OptionButton19.Value = True Then
                .Range("FH31").Value = "No Incluido"
            End If
    End With
    MsgBox "Datos Cargados a la ficha técnica", vbInformation, "Aviso"
    Codigo_unifilar
    'cerrar el formulario
    Unload Me
End Sub
Public Sub Codigo_unifilar()
    Dim codigo_GAM As String
    Dato2 = "01"
    If Me.OptionButton23.Value = True Or Me.OptionButton24.Value = True Then
        Dato3 = "02"
    ElseIf Me.OptionButton21.Value = True Or Me.OptionButton22.Value = True Then
        Dato3 = "01"
    End If
    If Me.OptionButton25.Value = True Then
        Dato4 = "02"
    ElseIf Me.OptionButton26.Value = True Then
        Dato4 = "01"
    End If
    codigo_GAM = "UN_" & Dato1 & Dato2 & Dato3 & Dato4
    If PosicionUnifilarActual = 0 Then
        MsgBox "No se ha definido la posicion (H1-H9) para este bloque.", vbExclamation
        Exit Sub
    End If

    Asignar_Posicion_Unifilar PosicionUnifilarActual, codigo_GAM, 17
End Sub


    

