'=============Hoja configurador=============
private sub ClaseSensores_Change()
    dim clase as string
    clase = Me.OLEObjects("ClaseSensores").Object.Value
    With Me.OLEObjects("Sensor").Object
        .clear
        select case clase
            case "TH110"
                .additem "Kit de sensores TH110"
            case "CL110"
                .additem "CL110-x1"
                .additem "CL110-x2"
                .additem "CL110-x4"
        end select
    end with
end sub
Public sub Worksheet_activate()
    
    With Me.OLEObjects("Contactor").Object
        .clear
        .additem "Contactor abierto"
        .additem "Contactor cerrado"
    end with
    With Me.OLEObjects("ClaseSensores").Object
        .clear
        .additem "TH110"
        .additem "CL110"
    end with
    with Me.OLEObjects("VPIS").Object
        .clear
        .additem "VPIS (2 - 4) KV"
        .additem "VPIS (3,4 - 6,3)KV"
    end with
    With Me.OLEObjects("RELE").Object
        .clear
        .additem "Relé P3U30-6AAA2BCAA"
        .additem "Relé P3T32-CGG1A-AA1NA-BBAAA"
        .additem "Relé P5U20-AABD-IABAA-BAEI"
    end with
    with Me.OLEObjects("Modulo").Object
        .clear
        .additem "AGC-150"
    end with
    with Me.OLEObjects("TRACORRIENTE").Object
        .clear
        .additem "50/5A Clase 0.5S 10VA"
        .additem "100/5A Clase 0.5S 10VA"
        .additem "150/5A Clase 0.5S 10VA"
        .additem "200/5A Clase 0.5S 10VA"
        .additem "50/5A Clase 0.2S 10VA"
        .additem "100/5A Clase 0.2S 10VA"
        .additem "150/5A Clase 0.2S 10VA"
        .additem "200/5A Clase 0.2S 10VA"
        .additem "50/5A Clase 5P20 10VA"
        .additem "100/5A Clase 5P20 10VA"    
        .additem "150/5A Clase 5P20 10VA"
        .additem "200/5A Clase 5P20 10VA"
        .additem "50/5A Clase 10P20 10VA"
        .additem "100/5A Clase 10P20 10VA"
        .additem "150/5A Clase 10P20 10VA"
        .additem "200/5A Clase 10P20 10VA"
        .additem "50/5A N1: 0.5S N2:5P20 10VA"
        .additem "100/5A N1: 0.5S N2:5P20 10VA"
        .additem "150/5A N1: 0.5S N2:5P20 10VA"
        .additem "200/5A N1: 0.5S N2:5P20 10VA"
        .additem "50/5A N1: 0.5S N2:10P20 10VA"
        .additem "100/5A N1: 0.5S N2:10P20 10VA"
        .additem "150/5A N1: 0.5S N2:10P20 10VA"
        .additem "200/5A N1: 0.5S N2:10P20 10VA"
    end with
    with Me.OLEObjects("TRAPOTENCIA").Object
        .clear 
        .additem "PT Clase 0.2 10VA S.Normal"
        .additem "PT Clase 0.5 10VA S.Normal"
        .additem "PT Clase 3P 10VA S.Normal"
        .additem "PT N1: 0.5 N2:3P 10VA S.Normal"
        .additem "PT N1:Clase 0.2 N2: 3P 10VA S.Normal"
    end with
    with Me.OLEObjects("FUSIBLE").Object
        .clear
        .additem "6.3 Amp 10-24kV"
        .additem "10 Amp 10-24kV"
        .additem "16 Amp 10-24kV"
        .additem "20 Amp 10-24kV"
        .additem "25 Amp 10-24kV"
        .additem "31,5 Amp 10-24kV"
        .additem "40 Amp 10-24kV"
        .additem "50 Amp 10-24kV"
        .additem "63 Amp 10-24kV"
        .additem "80 Amp 10-24kV"
        .additem "100 Amp 10-24kV"
        .additem "125 Amp 10-24kV"
        .additem "160 Amp 10-24kV"
        .additem "200 Amp 10-24kV"
    end with
    With Me.OLEObjects("AnalizadorREDES").Object
        .clear
        .additem "PM5340"
        .additem "PM5560"
        .additem "PM8240"
    end with
    With Me.OLEObjects("MedidorTarifario").Object
        .clear
        .additem "Landis Clase 0.5s RS485"
        .additem "Landis Clase 0.5s RS485 + Ethernet"
        .additem "Landis Clase 0.2s RS485 + Ethernet"
        .additem "Microstar Clase 0.5s RS485"
        .additem "Inhemeter Clase 0.5s RS485"
        .additem "Itron Clase 0.5s RS485"
        .additem "Itron Clase 0.2s RS485"
    end with
    With Me.OLEObjects("GabineteBT").Object
        .clear
        .additem "Gabinete de baja tensión"
    end with
    With Me.OLEObjects("MandoMotorizado").Object
        .clear
        .additem "Motorizado 120VAC IM"
        .additem "Motorizado 120VAC QM"
    end with
    With Me.OLEObjects("Gateway").Object
        .clear
        .additem "PAS600"
        .additem "PAS600L"
    end with
    With Me.OLEObjects("AdaptadorTension").Object
        .clear
        .additem "Adaptador de tensión"
    end with
    with Me.OLEObjects("Zocalo").Object
        .clear
        .additem "Zócalos 400mm"
    end with
end sub
 