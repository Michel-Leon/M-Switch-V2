Attribute VB_Name = "M�dulo4"
Public Sub SM624_Click()
   Dato1 = "01"
    msgbox "Dato1 = " & Dato1
End Sub
Public sub Limpiar_ctr_unifilar()
    thisworkbook.sheets("_CTRL_UNIFILAR").Range("B2:C10").Value = ""
End sub