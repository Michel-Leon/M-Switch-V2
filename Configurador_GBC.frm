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
