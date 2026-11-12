' Hoja PANEL
Private Sub Worksheet_Activate()
   With thisWorkbook.Sheets("Panel")
      .range ("DK89:DK121").value = ""
   End With
End Sub 