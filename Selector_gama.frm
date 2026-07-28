VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Selector_gama 
   Caption         =   "NAVEGADOR"
   ClientHeight    =   6130
   ClientLeft      =   70
   ClientTop       =   310
   ClientWidth     =   6420
   OleObjectBlob   =   "Selector_gama.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Selector_gama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Initialize()
    'Tamaño de la ventana
    Me.Height = 488
    Me.Width = 493
End Sub    
Private Sub SM624_Click()
   ThisWorkbook.Sheets("ESQUEMA").Activate
   Unload Me
End Sub

Private Sub salir_Click()
   Unload Me
End Sub
