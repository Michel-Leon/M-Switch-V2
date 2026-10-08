VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Configurador_DM1A 
   Caption         =   "DM1A"
   ClientHeight    =   8960.001
   ClientLeft      =   -170
   ClientTop       =   -860
   ClientWidth     =   7900
   OleObjectBlob   =   "Configurador_DM1A.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Configurador_DM1A"
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
    Me.OptionButton15.Value = True
    Me.OptionButton17.Value = True
    Me.OptionButton19.Value = True
    Me.OptionButton23.Value = True
    Me.OptionButton27.Value = True
    Me.OptionButton30.Value = True
    Me.OptionButton47.Value = True
    Me.OptionButton54.Value = True
    Me.OptionButton58.Value = True
    Me.OptionButton60.Value = True
    Me.OptionButton62.Value = True
    Me.OptionButton65.Value = True
    Me.Rf1.Caption = "SM61D1JHD6X7CDM1A"
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
    SetRef OptionButton24.Value, Me.Rf6, Me.Ct6, "KIT-P3U30", "1"
End Sub
Private Sub OptionButton25_Click()
    SetRef OptionButton25.Value, Me.Rf6, Me.Ct6, "KIT-P3T32", "1"
End Sub
Private Sub OptionButton26_Click()
    SetRef OptionButton26.Value, Me.Rf6, Me.Ct6, "KIT-P5U20", "1"
End Sub
Private Sub OptionButton27_Click()
    SetRef OptionButton27.Value, Me.Rf6, Me.Ct6, "", ""
End Sub
'============= F)
Private Sub OptionButton28_Click()
    SetRef OptionButton28.Value, Me.Rf7, Me.Ct7, "KIT-SYNC-MT-01", "1"
End Sub
Private Sub OptionButton29_Click()
    SetRef OptionButton29.Value, Me.Rf7, Me.Ct7, "KIT-ATS-MT-01", "1"
End Sub
Private Sub OptionButton30_Click()
    SetRef OptionButton30.Value, Me.Rf7, Me.Ct7, "", ""
End Sub
'============= G)
Private Sub OptionButton31_Click()
    SetRef OptionButton31.Value, Me.Rf8, Me.Ct8, "CTMT_17,5_50/5_5P20_BR", "3"
    Me.OptionButton47.Value = False
End Sub
Private Sub OptionButton32_Click()
    SetRef OptionButton32.Value, Me.Rf8, Me.Ct8, "CTMT_17,5_100/5_5P20_BR", "3"
    Me.OptionButton47.Value = False
End Sub
Private Sub OptionButton33_Click()
    SetRef OptionButton33.Value, Me.Rf8, Me.Ct8, "CTMT_17,5_150/5_5P20_BR", "3"
    Me.OptionButton47.Value = False
End Sub
Private Sub OptionButton34_Click()
    SetRef OptionButton34.Value, Me.Rf8, Me.Ct8, "CTMT_17,5_200/5_5P20_BR", "3"
    Me.OptionButton47.Value = False
End Sub
Private Sub OptionButton35_Click()
    SetRef OptionButton35.Value, Me.Rf8, Me.Ct8, "CTMT_17,5_50/5_10P20_BR", "3"
    Me.OptionButton47.Value = False
End Sub
Private Sub OptionButton36_Click()
    SetRef OptionButton36.Value, Me.Rf8, Me.Ct8, "CTMT_17,5_100/5_10P20_BR", "3"
    Me.OptionButton47.Value = False
End Sub
Private Sub OptionButton37_Click()
    SetRef OptionButton37.Value, Me.Rf8, Me.Ct8, "CTMT_17,5_150/5_10P20_BR", "3"
    Me.OptionButton47.Value = False
End Sub
Private Sub OptionButton38_Click()
    SetRef OptionButton38.Value, Me.Rf8, Me.Ct8, "CTMT_17,5_200/5_10P20_BR", "3"
    Me.OptionButton47.Value = False
End Sub
Private Sub OptionButton39_Click()
    SetRef OptionButton39.Value, Me.Rf8, Me.Ct8, "CTMT_17,5_50/5_0.5s_5P20_BR", "3"
    Me.OptionButton47.Value = False
End Sub
Private Sub OptionButton40_Click()
    SetRef OptionButton40.Value, Me.Rf8, Me.Ct8, "CTMT_17,5_100/5_0.5s_5P20_BR", "3"
    Me.OptionButton47.Value = False
End Sub
Private Sub OptionButton41_Click()
    SetRef OptionButton41.Value, Me.Rf8, Me.Ct8, "CTMT_17,5_150/5_0.5s_5P20_BR", "3"
    Me.OptionButton47.Value = False
End Sub
Private Sub OptionButton42_Click()
    SetRef OptionButton42.Value, Me.Rf8, Me.Ct8, "CTMT_17,5_200/5_0.5s_5P20_BR", "3"
    Me.OptionButton47.Value = False
End Sub
Private Sub OptionButton43_Click()
    SetRef OptionButton43.Value, Me.Rf8, Me.Ct8, "CTMT_17,5_50/5_0.5s_10P20_BR", "3"
    Me.OptionButton47.Value = False
End Sub
Private Sub OptionButton44_Click()
    SetRef OptionButton44.Value, Me.Rf8, Me.Ct8, "CTMT_17,5_100/5_0.5s_10P20_BR", "3"
    Me.OptionButton47.Value = False
End Sub
Private Sub OptionButton45_Click()
    SetRef OptionButton45.Value, Me.Rf8, Me.Ct8, "CTMT_17,5_150/5_0.5s_10P20_BR", "3"
    Me.OptionButton47.Value = False
End Sub
Private Sub OptionButton46_Click()
    SetRef OptionButton46.Value, Me.Rf8, Me.Ct8, "CTMT_17,5_200/5_0.5s_10P20_BR", "3"
    Me.OptionButton47.Value = False
End Sub
Private Sub OptionButton47_Click()
    SetRef OptionButton47.Value, Me.Rf8, Me.Ct8, "", ""
End Sub
Private Sub OptionButton48_Click()
    SetRef OptionButton48.Value, Me.Rf8, Me.Ct8, "", ""
    Me.OptionButton47.Value = False
End Sub
'============= H)
Private Sub OptionButton49_Click()
    SetRef OptionButton49.Value, Me.Rf9, Me.Ct9, "PT-13.2-120-02-10-BR-N", "3"
End Sub
Private Sub OptionButton50_Click()
    SetRef OptionButton50.Value, Me.Rf9, Me.Ct9, "PT-13.2-120-05-10-BR-N", "3"
End Sub
Private Sub OptionButton51_Click()
    SetRef OptionButton51.Value, Me.Rf9, Me.Ct9, "PT-13.2-120-3P-10-BR-N", "3"
End Sub
Private Sub OptionButton52_Click()
    SetRef OptionButton52.Value, Me.Rf9, Me.Ct9, "PT-13.2-120-05-3P-10-BR-N", "3"
End Sub
Private Sub OptionButton53_Click()
    SetRef OptionButton53.Value, Me.Rf9, Me.Ct9, "PT-13.2-120-02-3P-10-BR-N", "3"
End Sub
Private Sub OptionButton54_Click()
    SetRef OptionButton54.Value, Me.Rf9, Me.Ct9, "", ""
End Sub
'============= I)
Private Sub OptionButton55_Click()
    SetRef OptionButton55.Value, Me.Rf10, Me.Ct10, "KIT-AR-01", "1"
End Sub
Private Sub OptionButton56_Click()
    SetRef OptionButton56.Value, Me.Rf10, Me.Ct10, "KIT-AR-02", "1"
End Sub
Private Sub OptionButton57_Click()
    SetRef OptionButton57.Value, Me.Rf10, Me.Ct10, "KIT-AR-03", "1"
End Sub
Private Sub OptionButton58_Click()
    SetRef OptionButton58.Value, Me.Rf10, Me.Ct10, "", ""
End Sub
'============= J)
Private Sub OptionButton59_Click()
    SetRef OptionButton59.Value, Me.Rf11, Me.Ct11, "GAB-CONT-MT", "1"
End Sub
Private Sub OptionButton60_Click()
    SetRef OptionButton60.Value, Me.Rf11, Me.Ct11, "", ""
End Sub
'============= K)
Private Sub OptionButton61_Click()
    SetRef OptionButton61.Value, Me.Rf12, Me.Ct12, "EMS59572", "1"
End Sub
Private Sub OptionButton62_Click()
    SetRef OptionButton62.Value, Me.Rf12, Me.Ct12, "", ""
End Sub
'============= L)
Private Sub OptionButton63_Click()
    SetRef OptionButton63.Value, Me.Rf13, Me.Ct13, "KIT-ZOCALO-400-SM6-24", "1"
End Sub
Private Sub OptionButton64_Click()
    SetRef OptionButton64.Value, Me.Rf13, Me.Ct13, "KIT-ZOCALO-600-SM6-24", "1"
End Sub
Private Sub OptionButton65_Click()
    SetRef OptionButton65.Value, Me.Rf13, Me.Ct13, "", ""
End Sub


Private Sub Registar_Click()
    With ThisWorkbook.Sheets("FT DM1A")
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
                .Range("FH50").Value = "Relé P3U30-6AAA2BCAA"
            ElseIf Me.OptionButton25.Value = True Then
                .Range("FH50").Value = "Relé P3T32-CGG1A-AA1NA-BBAAA"
            ElseIf Me.OptionButton26.Value = True Then
                .Range("FH50").Value = "Relé P5U20-AABD-IABAA-BAEI"
            ElseIf Me.OptionButton27.Value = True Then
                .Range("FH50").Value = "No Incluido"
            End If
        '============= F)
            If Me.OptionButton28.Value = True Then
                .Range("FH55").Value = "AGC-150"
            ElseIf Me.OptionButton29.Value = True Then
                .Range("FH55").Value = "ATS-270 Comap"
            ElseIf Me.OptionButton30.Value = True Then
                .Range("FH55").Value = "No Incluido"
            End If
        '============= G)
            If Me.OptionButton31.Value = True Then
                .Range("FH60").Value = "50/5A Clase 5P20 10VA 17.5kV"
            ElseIf Me.OptionButton32.Value = True Then
                .Range("FH60").Value = "100/5A Clase 5P20 10VA 17.5kV"
            ElseIf Me.OptionButton33.Value = True Then
                .Range("FH60").Value = "150/5A Clase 5P20 10VA 17.5kV"
            ElseIf Me.OptionButton34.Value = True Then
                .Range("FH60").Value = "200/5A Clase 5P20 10VA 17.5kV"
            ElseIf Me.OptionButton35.Value = True Then
                .Range("FH60").Value = "50/5A Clase 10P20 10VA 17.5kV"
            ElseIf Me.OptionButton36.Value = True Then
                .Range("FH60").Value = "100/5A Clase 10P20 10VA 17.5kV"
            ElseIf Me.OptionButton37.Value = True Then
                .Range("FH60").Value = "150/5A Clase 10P20 10VA 17.5kV"
            ElseIf Me.OptionButton38.Value = True Then
                .Range("FH60").Value = "200/5A Clase 10P20 10VA 17.5kV"
            ElseIf Me.OptionButton39.Value = True Then
                .Range("FH60").Value = "50/5A N1: 0.5S N2:5P20 10VA 17.5kV"
            ElseIf Me.OptionButton40.Value = True Then
                .Range("FH60").Value = "100/5A N1: 0.5S N2:5P20 10VA 17.5kV"
            ElseIf Me.OptionButton41.Value = True Then
                .Range("FH60").Value = "150/5A N1: 0.5S N2:5P20 10VA 17.5kV"
            ElseIf Me.OptionButton42.Value = True Then
                .Range("FH60").Value = "200/5A N1: 0.5S N2:5P20 10VA 17.5kV"
            ElseIf Me.OptionButton43.Value = True Then
                .Range("FH60").Value = "50/5A N1: 0.5S N2:10P20 10VA 17.5kV"
            ElseIf Me.OptionButton44.Value = True Then
                .Range("FH60").Value = "100/5A N1: 0.5S N2:10P20 10VA 17.5kV"
            ElseIf Me.OptionButton45.Value = True Then
                .Range("FH60").Value = "150/5A N1: 0.5S N2:10P20 10VA 17.5kV"
            ElseIf Me.OptionButton46.Value = True Then
                .Range("FH60").Value = "200/5A N1: 0.5S N2:10P20 10VA 17.5kV"
            ElseIf Me.OptionButton47.Value = True Then
                .Range("FH60").Value = "LPCT's 5-1250A / 22.5mV Cl 5P250"
            ElseIf Me.OptionButton48.Value = True Then
                .Range("FH60").Value = "No Incluido"
            End If
        '============= H)
            If Me.OptionButton49.Value = True Then
                .Range("FH70").Value = "PT Clase 0.2 10VA S.Normal 17,5kV"
            ElseIf Me.OptionButton50.Value = True Then
                .Range("FH70").Value = "PT Clase 0.5 10VA S.Normal 17,5kV"
            ElseIf Me.OptionButton51.Value = True Then
                .Range("FH70").Value = "PT Clase 3P 10VA S.Normal 17,5kV"
            ElseIf Me.OptionButton52.Value = True Then
                .Range("FH70").Value = "PT N1: 0.5 N2:3P 10VA S.Normal 17,5kV"
            ElseIf Me.OptionButton53.Value = True Then
                .Range("FH70").Value = "PT N1:Clase 0.2 N2: 3P 10VA S.Normal 17,5kV"
            ElseIf Me.OptionButton54.Value = True Then
                .Range("FH70").Value = "No Incluido"
            End If
        '============= I)
            If Me.OptionButton55.Value = True Then
                .Range("FH80").Value = "PM5340"
            ElseIf Me.OptionButton56.Value = True Then
                .Range("FH80").Value = "PM5560"
            ElseIf Me.OptionButton57.Value = True Then
                .Range("FH80").Value = "PM8240"
            ElseIf Me.OptionButton58.Value = True Then
                .Range("FH80").Value = "No Incluido"
            End If
        '============= J)
            If Me.OptionButton59.Value = True Then
                .Range("FH85").Value = "Incluido"
            ElseIf Me.OptionButton60.Value = True Then
                .Range("FH85").Value = "No Incluido"
            End If
        '============= K)
            If Me.OptionButton61.Value = True Then
                .Range("FH90").Value = "Incluido"
            ElseIf Me.OptionButton62.Value = True Then
                .Range("FH90").Value = "No Incluido"
            End If
        '============= L)
            If Me.OptionButton63.Value = True Then
                .Range("FH95").Value = "400mm"
            ElseIf Me.OptionButton64.Value = True Then
                .Range("FH95").Value = "600mm"
            ElseIf Me.OptionButton65.Value = True Then
                .Range("FH95").Value = "No Incluido"
            End If
    End With
    MsgBox "Datos Cargados a la ficha técnica", vbInformation, "Aviso"
    Codigo_unifilar
    'cerrar el formulario
    Unload Me
End Sub
' tomar valores para codigo bloque unifilar
Public Sub Codigo_unifilar()
    Dim codigo_DM1A As String

    Dato2 = "04"

    If Me.OptionButton17.Value = True Then
        Dato3 = "01"
    ElseIf Me.OptionButton16.Value = True Then
        Dato3 = "02"
    End If

    If Me.OptionButton47.Value = True Then
        Dato4 = "01"
    ElseIf HayActivoEnRango(31, 46) And HayActivoEnRango(49, 53) And HayActivoEnRango(55, 57) Then
        Dato4 = "03"
    ElseIf HayActivoEnRango(31, 46) And HayActivoEnRango(49, 53) Then
        Dato4 = "02"
    Else
        Dato4 = ""
    End If

    codigo_DM1A = "UN_" & Dato1 & Dato2 & Dato3 & Dato4

    If PosicionUnifilarActual = 0 Then
        MsgBox "No se ha definido la posicion (H1-H9) para este bloque.", vbExclamation
        Exit Sub
    End If

    Asignar_Posicion_Unifilar PosicionUnifilarActual, codigo_DM1A, 34
End Sub
Private Function HayActivoEnRango(desde As Integer, hasta As Integer) As Boolean
    Dim i As Integer
    For i = desde To hasta
        If Me.Controls("OptionButton" & i).Value = True Then
            HayActivoEnRango = True
            Exit Function
        End If
    Next i
    HayActivoEnRango = False
End Function

