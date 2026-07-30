VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Configurador_IM 
   Caption         =   "DM1A"
   ClientHeight    =   7360
   ClientLeft      =   90
   ClientTop       =   330
   ClientWidth     =   12110
   OleObjectBlob   =   "Configurador_IM.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Configurador_IM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub UserForm_Initialize()
    ' Tamaño de la ventana
    Me.Height = 580
    Me.Width = 932
     ' Centrar en la pantalla
    Me.StartUpPosition = 0  ' Manual, para poder controlar posici�n
    Me.Top = (Application.Height - Me.Height) / 2
    Me.Left = (Application.Width - Me.Width) / 2
    msgbox Resultado_IM
    if Resultado_IM  = "1" Then
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
Private sub OptionButton16_Click()
    SetRef OptionButton16.Value, Me.Rf3, Me.Ct3, "KIT-TH110-01", "1"
End sub
private sub OptionButton17_Click()
    SetRef OptionButton17.Value, Me.Rf3, Me.Ct3, "", ""
End sub
'============= C)
private sub OptionButton18_Click()
    SetRef OptionButton18.Value, Me.Rf4, Me.Ct4, "ZBRTT1", "1"
End sub
private sub OptionButton19_Click()
    SetRef OptionButton19.Value, Me.Rf4, Me.Ct4, "", ""
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
    SetRef OptionButton24.Value, Me.Rf6, Me.Ct6, "GAB-CONT-MT", "1"
End sub
Private sub OptionButton25_Click()
    SetRef OptionButton25.Value, Me.Rf6, Me.Ct6, "", ""
End sub
'============= F)
Private sub OptionButton26_Click()
    SetRef OptionButton26.Value, Me.Rf7, Me.Ct7, "KIT-MCH-IM-120V", "1"
End sub
Private sub OptionButton27_Click()
    SetRef OptionButton27.Value, Me.Rf7, Me.Ct7, "", ""
End sub
'============= G)
Private sub OptionButton28_Click()
    SetRef OptionButton28.Value, Me.Rf8, Me.Ct8, "KIT-ZOCALO-400-SM6-24", "1"
    
End sub
Private sub OptionButton29_Click()
    SetRef OptionButton29.Value, Me.Rf8, Me.Ct8, "KIT-ZOCALO-600-SM6-24", "1"
  
End sub
private sub OptionButton30_Click()
    SetRef OptionButton30.Value, Me.Rf8, Me.Ct8, "", ""

End sub
Private Sub Registar_Click()
    with thisWorkbook.sheets("FT IM")
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
                .Range("FH35").Value = "sensores TH110 x3" 
            elseif Me.OptionButton17.Value = True then 
                .Range("FH35").Value = "No Incluido"
            end if
        '============= C)
            if Me.OptionButton18.Value = True then 
                .Range("FH40").Value = "Sensor CL110"
            elseif Me.OptionButton19.Value = True then 
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
                .Range("FH50").Value = "Incluido"
            elseif Me.OptionButton25.Value = True then 
                .Range("FH50").Value = "No Incluido"
            end if
        '============= F)
            if Me.OptionButton26.Value = True then 
                .Range("FH55").Value = "Motorizado 120VAC IM"
            elseif Me.OptionButton27.Value = True then
                .Range("FH55").Value = "No Incluido"    
            end if
        '============= G)
            if Me.OptionButton28.Value = True then   
                .Range("FH60").Value = "400mm"
            ELseif Me.OptionButton29.Value = True then  
                .Range("FH60").Value = "600mm"
            Elseif Me.OptionButton30.Value = True then  
                .Range("FH60").Value = "No Incluido"    
            end if       
    end with
    MsgBox "Datos Cargados a la ficha técnica", vbInformation, "Aviso"
    'cerrar el formulario
    Unload Me
End Sub