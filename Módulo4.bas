Attribute VB_Name = "M�dulo4"
Public Sub SM624_Click()
   Dato1 = "01"
    With thisworkbook.sheets("Panel")
        .Range("DK89").Value = "17,5kV"
        .Range("DK97").Value = "630 A"
        .Range("DK105").Value = "20 kA/1s"
        .Range("DK113").Value = "AIS"
        .Range("DK121").Value = "SF6"
    End With
    msgbox "Modulo de Diseño habilitado.", vbInformation, "SM624"
End Sub
Public Sub SM636_Click()
    Dato1 = "" 
    With thisworkbook.sheets("Panel")
        .Range("DK89").Value = "34,5kV"
        .Range("DK97").Value = "630A"
        .Range("DK105").Value = "20kA/1S"
        .Range("DK113").Value = "AIS"
        .Range("DK121").Value = "SF6"
    End With
     msgbox "Modulo de Diseño Inhabilitado.", vbCritical, "SM636"
End Sub
Public Sub Premset_Click()
    Dato1 = ""
    With thisworkbook.sheets("Panel")
        .Range("DK89").Value = "17,5kV"
        .Range("DK97").Value = "630/1250A"
        .Range("DK105").Value = "25kA/3S"
        .Range("DK113").Value = "2SIS"
        .Range("DK121").Value = "-"
    End With
    msgbox "Modulo de Diseño Inhabilitado.", vbCritical, "Premset"
End Sub
Public Sub AiSeT_Click()
    Dato1 = ""
    With thisworkbook.sheets("Panel")
        .Range("DK89").Value = "17,5kV"
        .Range("DK97").Value = "630A"
        .Range("DK105").Value = "20kA/1S"
        .Range("DK113").Value = "AIS"
        .Range("DK121").Value = "-"
    End With
    msgbox "Modulo de Diseño Inhabilitado.", vbCritical, "AiSeT libre de SF6"
End Sub
Public Sub PIX_Click()
    Dato1 = ""
    With thisworkbook.sheets("Panel")
        .Range("DK89").Value = "17,5kV"
        .Range("DK97").Value = "2500/4000A"
        .Range("DK105").Value = "31,5/40kA"
        .Range("DK113").Value = "AIS"
        .Range("DK121").Value = "-"
    End With
    msgbox "Modulo de Diseño Inhabilitado.", vbCritical, "PIX"
End Sub
Public sub Limpiar_ctr_unifilar()
    thisworkbook.sheets("_CTRL_UNIFILAR").Range("B2:C10").Value = ""
End sub
