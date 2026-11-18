'=login=
Private Sub Worksheet_Activate()
    with thisworkbook.sheets("login")
        .range("HX83").value ""
        .range("HX98").value ""
    end with
End Sub

