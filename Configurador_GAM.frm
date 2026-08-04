VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Configurador_GAM 
   Caption         =   "DM1A"
   ClientHeight    =   8960.001
   ClientLeft      =   90
   ClientTop       =   320
   ClientWidth     =   12280
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
Private sub OptionButton16_Click()
    SetRef OptionButton16.Value, Me.Rf3, Me.Ct3, "ZBRTT1", "1"
End sub
private sub OptionButton17_Click()
    SetRef OptionButton17.Value, Me.Rf3, Me.Ct3, "", ""
End sub
'============= C)
private sub OptionButton18_Click()
    SetRef OptionButton18.Value, Me.Rf4, Me.Ct4, "KIT-ZOCALO-400-SM6-24", "1"
End sub
private sub OptionButton19_Click()
    SetRef OptionButton19.Value, Me.Rf4, Me.Ct4, "KIT-ZOCALO-600-SM6-24", ""
End sub
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
    with thisWorkbook.sheets("FT GAM 24kV")
        '============= 1)
            if Me.OptionButton1.Value = True then 'Voltaje nominal
                .Range("CZ32").Value = "17.5 kV" 
            elseif Me.OptionButton2.Value = True then 
                .Range("CZ32").Value = "24 kV"
            end if
        '============= 2)
            if Me.OptionButton3.Value = True then 'Voltaje de operacion
                .Range("CZ36").Value = "11.4 kV"
            elseif Me.OptionButton4.Value = True then 
                .Range("CZ36").Value = "13.2  kV"
            elseif Me.OptionButton5.Value = True then
                .Range("CZ36").Value = "13.8 kV"    
            end if
        '============= 3)
            if Me.OptionButton6.Value = True then 'Corriente
                .Range("CZ40").Value = "630 A"
            elseif Me.OptionButton7.Value = True then 
                .Range("CZ40").Value = "1250 A"
            end if
        '============= 4)
            if Me.OptionButton8.Value = True then 'corriente corto
                .Range("CZ52").Value = "20 kA/1s"
            elseif Me.OptionButton9.Value = True then 
                .Range("CZ52").Value = "25 kA/1S"
            end if
        '============= 6) 
            if Me.OptionButton12.Value = True then 
                .Range("DA114").Value = "AC-AFL 12,5kA/1s"
            elseif Me.OptionButton13.Value = True then 
                .Range("DA114").Value = "IAC-AFLR 12,5kA/1s"  
            end if 
        '============= A)
            if Me.OptionButton14.Value = True then 
                .Range("FH23").Value = "Incluido"
            elseif Me.OptionButton15.Value = True then
                .Range("FH23").Value = "No Incluido"
            end if
        '============= B)
            if Me.OptionButton16.Value = True then 
                .Range("FH28").Value = "Incluido"
            elseif Me.OptionButton17.Value = True then
                .Range("FH28").Value = "No Incluido"
            end if
        '============= C)
            if Me.OptionButton18.Value = True then 
                .Range("FH31").Value = "Incluido"
            elseif Me.OptionButton19.Value = True then
                .Range("FH31").Value = "No Incluido"
            end if
    end with
    MsgBox "Datos Cargados a la ficha técnica", vbInformation, "Aviso"
    codigo_unifilar
    'cerrar el formulario
    Unload Me
End Sub  
public sub codigo_unifilar()
    dim codigo_GAM As String
    Dato2="01"
    if Me.OptionButton23.Value = True or Me.OptionButton24.Value = True then 
        Dato3="02"
    elseif Me.OptionButton21.Value = True or Me.OptionButton22.Value = True then 
        Dato3="01"
    end if  
    if Me.OptionButton25.Value = True then 
        Dato4="02"
    elseif Me.OptionButton26.Value = True then 
        Dato4="01"
    end if
    Codigo_GAM = "UN_" & Dato1 & Dato2 & Dato3 & Dato4
    If PosicionUnifilarActual = 0 Then
        MsgBox "No se ha definido la posicion (H1-H9) para este bloque.", vbExclamation
        Exit Sub
    End If

    Asignar_Posicion_Unifilar PosicionUnifilarActual, codigo_GAM, 17
End sub


    