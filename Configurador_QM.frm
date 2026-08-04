VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Configurador_QM 
   Caption         =   "DM1A"
   ClientHeight    =   7360
   ClientLeft      =   90
   ClientTop       =   330
   ClientWidth     =   12220
   OleObjectBlob   =   "Configurador_QM.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Configurador_QM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub UserForm_Initialize()
    'Tama�o de la ventana
    Me.Height = 580
    Me.Width = 932
     ' Centrar en la pantalla
    Me.StartUpPosition = 0  ' Manual, para poder controlar posici�n
    Me.Top = (Application.Height - Me.Height) / 2
    Me.Left = (Application.Width - Me.Width) / 2
     Me.OptionButton1.Value = True
        Me.OptionButton7.Value = True
        Me.OptionButton9.Value = True
        me.OptionButton10.Value = True
        Me.OptionButton12.Value = True
        Me.OptionButton15.Value = True
        Me.OptionButton17.Value = True
        Me.OptionButton19.Value = True
        Me.OptionButton23.Value = True
        Me.OptionButton25.Value = True
        Me.OptionButton27.Value = True
        Me.OptionButton30.Value = True
        Me.Rf1.Caption = "SM61Q3JHH6Z7CQMES"
        Me.Ct1.Caption = "1"
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
private sub OptionButton16_Click()
    SetRef OptionButton18.Value, Me.Rf4, Me.Ct4, "ZBRTT1", "1"
End sub
private sub OptionButton17_Click()
    SetRef OptionButton19.Value, Me.Rf4, Me.Ct4, "", ""
End sub
'============= C)
Private sub OptionButton18_Click()
    SetRef OptionButton16.Value, Me.Rf3, Me.Ct3, "KIT-TH110-01", "1"
End sub
private sub OptionButton19_Click()
    SetRef OptionButton17.Value, Me.Rf3, Me.Ct3, "", ""
End sub
'============= D)
private sub OptionButton20_Click()
    SetRef OptionButton20.Value, Me.Rf5, Me.Ct5, "VPI62403", "1"
End sub
private sub OptionButton21_Click()
    SetRef OptionButton21.Value, Me.Rf5, Me.Ct5, "VPI62404", "1"
End sub
Private sub OptionButton22_Click()
    SetRef OptionButton22.Value, Me.Rf5, Me.Ct5, "VPI62405", "1"
End sub
Private sub OptionButton23_Click()
    SetRef OptionButton23.Value, Me.Rf5, Me.Ct5, "", ""
End sub
'============= E)
Private sub OptionButton24_Click()
    SetRef OptionButton24.Value, Me.Rf6, Me.Ct6, "2050151000003", "3"
End sub
Private sub OptionButton25_Click()
    SetRef OptionButton25.Value, Me.Rf6, Me.Ct6, "2050151000004", "3"
End sub
Private sub OptionButton26_Click()
    SetRef OptionButton26.Value, Me.Rf6, Me.Ct6, "2050151000005", "3"
End sub
Private sub OptionButton27_Click()
    SetRef OptionButton27.Value, Me.Rf6, Me.Ct6, "2050151000006", "3"
End sub
Private sub OptionButton28_Click()
    SetRef OptionButton28.Value, Me.Rf7, Me.Ct7, "2050151000007", "3"
End sub
Private sub OptionButton29_Click()
    SetRef OptionButton29.Value, Me.Rf7, Me.Ct7, "2050151000008", "3"
End sub
Private sub OptionButton30_Click()
    SetRef OptionButton30.Value, Me.Rf7, Me.Ct7, "2050151000009", "3"
End sub
Private sub OptionButton31_Click()
    SetRef OptionButton31.Value, Me.Rf8, Me.Ct8, "2050151000010", "3"
End sub
Private sub OptionButton32_Click()
    SetRef OptionButton32.Value, Me.Rf8, Me.Ct8, "2050151000011", "3"
End sub
private sub OptionButton33_Click()
    SetRef OptionButton33.Value, Me.Rf8, Me.Ct8, "2050151000012", "3"
End sub
Private sub OptionButton34_Click()
    SetRef OptionButton34.Value, Me.Rf8, Me.Ct8, "2050151000013", "3"
End sub
private sub OptionButton35_Click()
    SetRef OptionButton35.Value, Me.Rf8, Me.Ct8, "2050151000014", "3"
End sub
private sub OptionButton36_Click()
    SetRef OptionButton36.Value, Me.Rf8, Me.Ct8, "2050151000015", "3"
End sub
private sub OptionButton37_Click()
    SetRef OptionButton37.Value, Me.Rf8, Me.Ct8, "2050151000016", "3"
End sub
'============= F)
private sub OptionButton38_Click()
    SetRef OptionButton38.Value, Me.Rf9, Me.Ct9, "GAB-CONT-MT", "1" 
end sub
private sub OptionButton39_Click()
    SetRef OptionButton39.Value, Me.Rf9, Me.Ct9, "", ""
End sub
'============= G)
private sub OptionButton40_Click()
    SetRef OptionButton40.Value, Me.Rf10, Me.Ct10, "KIT-MCH-QM-120V", "1"

End sub
private sub OptionButton41_Click()
    SetRef OptionButton41.Value, Me.Rf10, Me.Ct10, "", ""
End sub
'============= H)
private sub OptionButton42_Click()
    SetRef OptionButton42.Value, Me.Rf11, Me.Ct11, "KIT-ZOCALO-400-SM6-24", "1"
End sub
private sub OptionButton43_Click()
    SetRef OptionButton43.Value, Me.Rf11, Me.Ct11, "KIT-ZOCALO-600-SM6-24", "1"
End sub
private sub OptionButton44_Click()
    SetRef OptionButton44.Value, Me.Rf11, Me.Ct11, "", ""
End sub
Private Sub Registar_Click()
    with thisWorkbook.sheets("FT QM")
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
        '============= C)
            if Me.OptionButton18.Value = True then 
                .Range("FH35").Value = "sensores TH110 x3" 
            elseif Me.OptionButton19.Value = True then 
                .Range("FH35").Value = "No Incluido"
            end if
        '============= B)
            if Me.OptionButton16.Value = True then 
                .Range("FH40").Value = "Sensor CL110"
            elseif Me.OptionButton17.Value = True then 
                .Range("FH40").Value = "No Incluido"
            end if
        '============= D)
            if Me.OptionButton20.Value = True then 
                .Range("FH45").Value = "VPIS (2 - 4) KV"
            elseif Me.OptionButton21.Value = True then 
                .Range("FH45").Value = "VPIS (3,4 - 6,3)KV - VPIS (13 - 24kV)"
            elseif Me.OptionButton22.Value = True then 
                .Range("FH45").Value = "VPIS (4-8)KV"
            elseif Me.OptionButton23.Value = True then 
                .Range("FH45").Value = "(9-17)kV"
            end if
        '============= E)
            if Me.OptionButton24.Value = True then 
                .Range("FH50").Value = "6.3 Amp 10-24kV"
            elseif Me.OptionButton25.Value = True then 
                .Range("FH50").Value = "10 Amp 10-24kV"
            elseif Me.OptionButton26.Value = True then 
                .Range("FH50").Value = "16 Amp 10-24kV"
            elseif Me.OptionButton27.Value = True then 
                .Range("FH50").Value = "20 Amp 10-24kV"
            Elseif Me.OptionButton28.Value = True then 
                .Range("FH50").Value = "25 Amp 10-24kV"  
            Elseif Me.OptionButton29.Value = True then
                .Range("FH50").Value = "31,5 Amp 10-24kV"   
            Elseif Me.OptionButton30.Value = True then     
                .Range("FH50").Value = "40 Amp 10-24kV" 
            Elseif Me.OptionButton31.Value = True then 
                .Range("FH50").Value = "50 Amp 10-24kV"
            Elseif Me.OptionButton32.Value = True then 
                .Range("FH50").Value = "63 Amp 10-24kV"
            Elseif Me.OptionButton33.Value = True then 
                .Range("FH50").Value = "80 Amp 10-24kV"
            Elseif Me.OptionButton34.Value = True then 
                .Range("FH50").Value = "100 Amp 10-24kV"
            Elseif Me.OptionButton35.Value = True then 
                .Range("FH50").Value = "125 Amp 10-24kV"
            Elseif Me.OptionButton36.Value = True then 
                .Range("FH50").Value = "160 Amp 10-24kV"
            Elseif Me.OptionButton37.Value = True then 
                .Range("FH50").Value = "200 Amp 10-24kV"
            end if
        '============= F)    
            if Me.OptionButton38.Value = True then
                .Range("FH55").Value = "200/5A Clase 10P20 10VA 17.5kV"
            Elseif Me.OptionButton39.Value = True then
                .Range("FH55").Value = "50/5A N1: 0.5S N2:5P20 10VA 17.5kV"
            end if 
        '============= G)
            if Me.OptionButton40.Value = True then   
                .Range("FH60").Value = "Motorizado 120VAC QM"
            ELseif Me.OptionButton41.Value = True then
                .Range("FH60").Value = "No Incluido"
            End if  
        '============= H)
            if Me.OptionButton42.Value = True then   
                .Range("FH65").Value = "400mm"
            ELseif Me.OptionButton43.Value = True then  
                .Range("FH65").Value = "600mm"
            Elseif Me.OptionButton44.Value = True then  
                .Range("FH65").Value = "No Incluido"    
            end if       
    end with
    MsgBox "Datos Cargados a la ficha técnica", vbInformation, "Aviso"
    codigo_unifilar
    'cerrar el formulario
    Unload Me
End Sub
Public sub Codigo_unifilar()
    Dim codigo_QM As String
    Dato2 = "03"
    If Me.OptionButton17.Value = True Then
        Dato3 = "02"
    elseif Me.OptionButton16.Value = True Then
        Dato3 = "01"
    End If

    codigo_QM = "UN_" & Dato1 & Dato2 & Dato3
    If PosicionUnifilarActual = 0 Then
        MsgBox "No se ha definido la posicion (H1-H9) para este bloque.", vbExclamation
        Exit Sub
    End If

    Asignar_Posicion_Unifilar PosicionUnifilarActual, codigo_QM, 17    
end sub
