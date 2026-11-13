'======== DOCUMENTACION =========
Public Sub Worksheet_Activate()
    With Me.oleObjects("HojaENT").Object
        If .ListCount = 0 Then
        .AddItem "Añadir" 'Hoja_Entrada
        .AddItem "No anadir"
        End If
    End With
    with Me.oleObjects("Esquema").Object
        If .ListCount = 0 Then
        .AddItem "Añadir" 'Esquema
        .AddItem "No anadir"
        End If
    End With
    with Me.oleObjects("Unifilar").Object
        If .ListCount = 0 Then
        .AddItem "Añadir" 'Unifilar
        .AddItem "No anadir"
        End If
    End With
    with Me.oleObjects("Mecanico").Object
        If .ListCount = 0 Then
        .AddItem "Añadir" 'Mecanico
        .AddItem "No anadir"
        End If
    End With
    with Me.oleObjects("Resumen").Object
        If .ListCount = 0 Then
        .AddItem "Añadir" 'Resumen
        .AddItem "No anadir"
        End If
    End With
    with Me.oleObjects("Oferta1").Object
        If .ListCount = 0 Then
        .AddItem "Añadir" 'Oferta
        .AddItem "No anadir"
        End If
    End With
    with Me.oleObjects("Ficha_TEC").Object
        .clear
        .AddItem ""
        .AddItem "Ficha tecnica DM1A" 'FT_DM1A
        .AddItem "Ficha tecnica QM"   'FT_QM
        .AddItem "Ficha tecnica IM"   'FT_IM
        .AddItem "Ficha tecnica GAM"  'FT_GAM
        .AddItem "Ficha tecnica GBC"  'FT_GBC
    End With
    with Me.oleObjects("Complemento").Object
        .clear
        .AddItem ""
        .AddItem "Complemento DM1A"  'Comp_DM1A
        .AddItem "Complemento QM"    'Comp_QM
        .AddItem "Complemento IM"    'Comp_IM
        .AddItem "Complemento GAM 0"    'Comp_GAM_0
        .AddItem "Complemento GAM 2"    'Comp_GAM_2
        .AddItem "Complemento GBC A"    'Comp_GBC_A
        .AddItem "Complemento GBC B"    'Comp_GBC_B
        .AddItem "Complemento GBC C"    'Comp_GBC_C
        .AddItem "Complemento GBC D"    'Comp_GBC_D
    End With
    with Me.oleObjects("Oferta2").Object
        .clear
        .AddItem "Añadir adicional" 'Oferta
        .AddItem "No anadir adicional"
    End With
End Sub