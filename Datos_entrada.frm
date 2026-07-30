VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Datos_entrada 
   ClientHeight    =   10570
   ClientLeft      =   130
   ClientTop       =   440
   ClientWidth     =   12760
   OleObjectBlob   =   "Datos_entrada.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Datos_entrada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub UserForm_Initialize()
    'Tamaño de la ventana
    Me.Height = 630
    Me.Width = 649
End Sub
Private Sub Saltar_Click()
   Selector_gama.Show
   'cerra formulario
   Unload Me
End Sub
