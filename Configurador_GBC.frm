VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Configurador_GBC 
   Caption         =   "DM1A"
   ClientHeight    =   7360
   ClientLeft      =   100
   ClientTop       =   330
   ClientWidth     =   12260
   OleObjectBlob   =   "Configurador_GBC.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Configurador_GBC"
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
'============= 1)
Private Sub OptionButton68_Click()
    SetRef OptionButton68.Value, Me.Rf1, Me.Ct1, "SE-GBC-A-24KV", "1"
End Sub
Private Sub OptionButton69_Click()
    SetRef OptionButton69.Value, Me.Rf1, Me.Ct1, "SE-GBC-B-24KV", "1"
End Sub
Private Sub OptionButton70_Click()
    SetRef OptionButton70.Value, Me.Rf1, Me.Ct1, "SE-GBC-C-24KV", "1"
End Sub
Private Sub OptionButton71_Click()
    SetRef OptionButton71.Value, Me.Rf1, Me.Ct1, "SE-GBC-D-24KV", "1"
End Sub
'============= A)
Private Sub OptionButton27_Click()
    SetRef OptionButton27.Value, Me.Rf2, Me.Ct2, "CTMT_17,5_5/5_0.2s_BR", "3"
End Sub
Private Sub OptionButton28_Click()
    SetRef OptionButton28.Value, Me.Rf2, Me.Ct2, "CTMT_17,5_7.5/5_0.2s_BR", "3"
End Sub
Private Sub OptionButton29_Click()
    SetRef OptionButton29.Value, Me.Rf2, Me.Ct2, "CTMT_17,5_10/5_0.2s_BR", "3"
End Sub
Private Sub OptionButton30_Click()
    SetRef OptionButton30.Value, Me.Rf2, Me.Ct2, "CTMT_17,5_15/5_0.2s_BR", "3"
End Sub
Private Sub OptionButton31_Click()
    SetRef OptionButton31.Value, Me.Rf2, Me.Ct2, "CTMT_17,5_20/5_0.2s_BR", "3"
End Sub
Private Sub OptionButton32_Click()
    SetRef OptionButton32.Value, Me.Rf2, Me.Ct2, "CTMT_17,5_25/5_0.2s_BR", "3"
End Sub
Private Sub OptionButton33_Click()
    SetRef OptionButton33.Value, Me.Rf2, Me.Ct2, "CTMT_17,5_30/5_0.2s_BR", "3"
End Sub
Private Sub OptionButton34_Click()
    SetRef OptionButton34.Value, Me.Rf2, Me.Ct2, "CTMT_17,5_40/5_0.2s_BR", "3"
End Sub
Private Sub OptionButton35_Click()
    SetRef OptionButton35.Value, Me.Rf2, Me.Ct2, "CTMT_17,5_50/5_0.2s_BR", "3"
End Sub
Private Sub OptionButton36_Click()
    SetRef OptionButton36.Value, Me.Rf2, Me.Ct2, "CTMT_17,5_60/5_0.2s_BR", "3"
End Sub
Private Sub OptionButton37_Click()
    SetRef OptionButton37.Value, Me.Rf2, Me.Ct2, "CTMT_17,5_70/5_0.2s_BR", "3"
End Sub
Private Sub OptionButton38_Click()
    SetRef OptionButton38.Value, Me.Rf2, Me.Ct2, "CTMT_17,5_80/5_0.2s_BR", "3"
End Sub
Private Sub OptionButton39_Click()
    SetRef OptionButton39.Value, Me.Rf2, Me.Ct2, "CTMT_17,5_100/5_0.2s_BR", "3"
End Sub
Private Sub OptionButton40_Click()
    SetRef OptionButton40.Value, Me.Rf2, Me.Ct2, "", ""
End Sub
'============= B)
Private Sub OptionButton41_Click()
    SetRef OptionButton41.Value, Me.Rf3, Me.Ct3, "PT-13.2-120-02-10-BR-N", "3"
End Sub
Private Sub OptionButton42_Click()
    SetRef OptionButton42.Value, Me.Rf3, Me.Ct3, "PT-13.2-120-05-10-BR-N", "3"
End Sub
Private Sub OptionButton43_Click()
    SetRef OptionButton43.Value, Me.Rf3, Me.Ct3, "PT-13.2-120-3P-10-BR-N", "3"
End Sub
Private Sub OptionButton44_Click()
    SetRef OptionButton44.Value, Me.Rf3, Me.Ct3, "PT-13.2-120-05-3P-10-BR-N", "3"
End Sub
Private Sub OptionButton45_Click()
    SetRef OptionButton45.Value, Me.Rf3, Me.Ct3, "PT-13.2-120-02-3P-10-BR-N", "3"
End Sub
Private Sub OptionButton46_Click()
    SetRef OptionButton46.Value, Me.Rf3, Me.Ct3, "", ""
End Sub
'============= C)
Private Sub OptionButton47_Click()
    SetRef OptionButton47.Value, Me.Rf4, Me.Ct4, "2010472000012", "1"
End Sub
Private Sub OptionButton48_Click()
    SetRef OptionButton48.Value, Me.Rf4, Me.Ct4, "2010472000023", "1"
End Sub
Private Sub OptionButton49_Click()
    SetRef OptionButton49.Value, Me.Rf5, Me.Ct5, "2010472000030", "1"
End Sub
Private Sub OptionButton50_Click()
    SetRef OptionButton50.Value, Me.Rf5, Me.Ct5, "2010060000003", "1"
End Sub 
Private Sub OptionButton51_Click()
    SetRef OptionButton51.Value, Me.Rf5, Me.Ct5, "2010665000005", "1"
End Sub
Private Sub OptionButton52_Click()
    SetRef OptionButton52.Value, Me.Rf6, Me.Ct6, "2010401000006", "1"
End Sub
Private Sub OptionButton53_Click()
    SetRef OptionButton53.Value, Me.Rf6, Me.Ct6, "2010451000001", "1"
End Sub
Private Sub OptionButton54_Click()
    SetRef OptionButton54.Value, Me.Rf6, Me.Ct6, "METSEION7400", "1"
End Sub
Private Sub OptionButton55_Click()
    SetRef OptionButton55.Value, Me.Rf6, Me.Ct6, "", ""
End Sub
'============= D)
Private Sub OptionButton56_Click()
    SetRef OptionButton56.Value, Me.Rf7, Me.Ct7, "2010472000012", "1"
End Sub
Private Sub OptionButton57_Click()
    SetRef OptionButton57.Value, Me.Rf7, Me.Ct7, "2010472000023", "1"
End Sub
Private Sub OptionButton58_Click()
    SetRef OptionButton58.Value, Me.Rf7, Me.Ct7, "2010472000030", "1"
End Sub
Private Sub OptionButton59_Click()
    SetRef OptionButton59.Value, Me.Rf7, Me.Ct7, "2010060000003", "1"
End Sub
Private Sub OptionButton60_Click()
    SetRef OptionButton60.Value, Me.Rf7, Me.Ct7, "2010665000005", "1"
End Sub
Private Sub OptionButton61_Click()
    SetRef OptionButton61.Value, Me.Rf7, Me.Ct7, "2010401000006", "1"
End Sub
Private Sub OptionButton62_Click()
    SetRef OptionButton62.Value, Me.Rf7, Me.Ct7, "2010451000001", "1"
End Sub
Private Sub OptionButton63_Click()  
    SetRef OptionButton63.Value, Me.Rf7, Me.Ct7, "METSEION7400", "1"
End Sub
Private Sub OptionButton64_Click()
    SetRef OptionButton64.Value, Me.Rf7, Me.Ct7, "", ""
End Sub
'============= E)
Private Sub OptionButton65_Click()
    SetRef OptionButton65.Value, Me.Rf8, Me.Ct8, "KIT-ZOCALO-400-SM6-24", "1"
End Sub
Private Sub OptionButton66_Click()
    SetRef OptionButton66.Value, Me.Rf8, Me.Ct8, "KIT-ZOCALO-600-SM6-24", "1"
End Sub
Private Sub OptionButton67_Click()
    SetRef OptionButton67.Value, Me.Rf8, Me.Ct8, "", ""
End Sub
Private Sub Registar_Click()
    with thisWorkbook.sheets("FT GBC")
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
                .Range("FH23").Value = "5A Clase 0.5s 10VA 17,5kV"
            elseif Me.OptionButton15.Value = True then  
                .Range("FH23").Value = "7.5/5A Clase 0.5s10VA 17,5kV"
            elseif Me.OptionButton16.Value = True then 
                .Range("FH23").Value = "10/5A Clase 0.5s10VA 17,5kV"    
            elseif Me.OptionButton17.Value = True then 
                .Range("FH23").Value = "15/5A Clase 0.5s10VA 17,5kV"    
            elseif Me.OptionButton18.Value = True then 
                .Range("FH23").Value = "20/5A Clase 0.5s10VA 17,5kV"   
            elseif Me.OptionButton19.Value = True then 
                .Range("FH23").Value = "25/5A Clase 0.5s10VA 17,5kV" 
            elseif Me.OptionButton20.Value = True then 
                .Range("FH23").Value = "30/5A Clase 0.5s10VA 17,5kV"
            elseif Me.OptionButton21.Value = True then 
                .Range("FH23").Value = "40/5A Clase 0.5s10VA 17,5kV"  
            elseif Me.OptionButton22.Value = True then 
                .Range("FH23").Value = "50/5A Clase 0.5s10VA 17,5kV"  
            elseif Me.OptionButton23.Value = True then 
                .Range("FH23").Value = "60/5A Clase 0.5s10VA 17,5kV"  
            elseif Me.OptionButton24.Value = True then 
                .Range("FH23").Value = "70/5A Clase 0.5s10VA 17,5kV"  
            elseif Me.OptionButton25.Value = True then 
                .Range("FH23").Value = "80/5A Clase 0.5s10VA 17,5kV"  
            elseif Me.OptionButton26.Value = True then
                .Range("FH23").Value = "100/5A Clase 0.5s10VA 17,5kV"  
            elseif Me.OptionButton27.Value = True then 
                .Range("FH23").Value = "5/5A Clase 0.2s 10VA 17,5kV"
            elseif Me.OptionButton28.Value = True then 
                .Range("FH23").Value = "7.5/5A Clase 0.2s10VA 17,5kV"
            elseif Me.OptionButton29.Value = True then  
                .Range("FH23").Value = "10/5A Clase 0.2s10VA 17,5kV"
            elseif Me.OptionButton30.Value = True then 
                .Range("FH23").Value = "15/5A Clase 0.2s10VA 17,5kV"    
            elseif Me.OptionButton31.Value = True then 
                .Range("FH23").Value = "20/5A Clase 0.2s10VA 17,5kV"   
            elseif Me.OptionButton32.Value = True then 
                .Range("FH23").Value = "25/5A Clase 0.2s10VA 17,5kV" 
            elseif Me.OptionButton33.Value = True then 
                .Range("FH23").Value = "30/5A Clase 0.2s10VA 17,5kV"
            elseif Me.OptionButton34.Value = True then 
                .Range("FH23").Value = "40/5A Clase 0.2s10VA 17,5kV"  
            elseif Me.OptionButton35.Value = True then 
                .Range("FH23").Value = "50/5A Clase 0.2s10VA 17,5kV"  
            elseif Me.OptionButton36.Value = True then 
                .Range("FH23").Value = "60/5A Clase 0.2s10VA 17,5kV"  
            elseif Me.OptionButton37.Value = True then 
                .Range("FH23").Value = "70/5A Clase 0.2s10VA 17,5kV"  
            elseif Me.OptionButton38.Value = True then 
                .Range("FH23").Value = "80/5A Clase 0.2s10VA 17,5kV"  
            Elseif Me.OptionButton39.Value = True then
                .Range("FH23").Value = "100/5A Clase 0.2s10VA 17,5kV"  
            elseif Me.OptionButton40.Value = True then
                .Range("FH23").Value = "No Incluido"    
            end if
        '============= B)
            if Me.OptionButton41.Value = True then 
                .Range("FH33").Value = "PT Clase 0.2 10VA S.Normal 17,5kV"
            elseif Me.OptionButton42.Value = True then 
                .Range("FH33").Value = "PT Clase 0.5 10VA S.Normal 17,5kV"
            elseif Me.OptionButton43.Value = True then 
                .Range("FH33").Value = "PT Clase 3P 10VA S.Normal 17,5kV"
            elseif Me.OptionButton44.Value = True then 
                .Range("FH33").Value = "PT N1: 0.5 N2:3P 10VA S.Normal 17,5kV"
            elseif Me.OptionButton45.Value = True then 
                .Range("FH33").Value = "PT N1:Clase 0.2 N2: 3P 10VA S.Normal 17,5kV"
            elseif Me.OptionButton46.Value = True then 
                .Range("FH33").Value = "No Incluido"
            end if
        '============= C)
            if Me.OptionButton47.Value = True then 
                .Range("FH43").Value = "Landis Clase 0.5s RS485"
            elseif Me.OptionButton48.Value = True then 
                .Range("FH43").Value = "Landis Clase 0.5s RS485 + Ethernet"
            elseif Me.OptionButton49.Value = True then 
                .Range("FH43").Value = "Landis Clase 0.2s RS485 + Ethernet"
            elseif Me.OptionButton50.Value = True then 
                .Range("FH43").Value = "Microstar Clase 0.5s RS485"
            elseif Me.OptionButton51.Value = True then 
                .Range("FH43").Value = "Inhemeter Clase 0.5s RS485"
            elseif Me.OptionButton52.Value = True then 
                .Range("FH43").Value = "Itron Clase 0.5s RS485"
            Elseif Me.OptionButton53.Value = True then 
                .Range("FH43").Value = "Itron Clase 0.2s RS485"
            elseif Me.OptionButton54.Value = True then 
                .Range("FH43").Value = "ION7400"
            elseif Me.OptionButton55.Value = True then 
                .Range("FH43").Value = "Suministro Cliente"        
            end if    
        '============= D)
            if Me.OptionButton56.Value = True then 
                .Range("FH53").Value = "Landis Clase 0.5s RS485"
            elseif Me.OptionButton57.Value = True then 
                .Range("FH53").Value = "Landis Clase 0.5s RS485 + Ethernet"
            elseif Me.OptionButton58.Value = True then 
                .Range("FH53").Value = "Landis Clase 0.2s RS485 + Ethernet"
            elseif Me.OptionButton59.Value = True then 
                .Range("FH53").Value = "Microstar Clase 0.5s RS485"
            elseif Me.OptionButton60.Value = True then 
                .Range("FH53").Value = "Inhemeter Clase 0.5s RS485"
            elseif Me.OptionButton61.Value = True then 
                .Range("FH53").Value = "Itron Clase 0.5s RS485"
            Elseif Me.OptionButton62.Value = True then 
                .Range("FH53").Value = "Itron Clase 0.2s RS485"
            elseif Me.OptionButton63.Value = True then 
                .Range("FH53").Value = "ION7400"
            elseif Me.OptionButton64.Value = True then 
                .Range("FH53").Value = "Suministro Cliente"        
            end if 
        '============= E)
            if Me.OptionButton65.Value = True then 
                .Range("FH63").Value = "400mm"
            elseif Me.OptionButton66.Value = True then 
                .Range("FH63").Value = "600mm"
            elseif Me.OptionButton67.Value = True then 
                .Range("FH63").Value = "No Incluido"        
            end if     
    end with
    MsgBox "Datos Cargados a la ficha técnica", vbInformation, "Aviso"
    'cerrar el formulario
    Unload Me
End Sub