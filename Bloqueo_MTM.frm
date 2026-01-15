VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Bloqueo_MTM 
   Caption         =   "LOGIN"
   ClientHeight    =   5100
   ClientLeft      =   30
   ClientTop       =   160
   ClientWidth     =   3220
   OleObjectBlob   =   "Bloqueo_MTM.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Bloqueo_MTM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub cmdLogin_Click()
    
    usuario = txtUsuario.Text
    clave = txtPassword.Text

    
    If usuario = "ADMIN" And clave = "Admin@GT78" Then
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"", True)"
    Application.DisplayFormulaBar = True
    ActiveWindow.DisplayHeadings = True
    ThisWorkbook.Names.Add Name:="usuario_actual", RefersTo:="=""" & usuario & """"
    Sheets("IM").Visible = True
    Sheets("IM").Activate
    Sheets("QM").Visible = True
    Sheets("QM").Activate
    Sheets("GAM").Visible = True
    Sheets("GAM").Activate
    Sheets("DM1").Visible = True
    Sheets("DM1").Activate
    Sheets("GBC 24kV 1 MED").Visible = True
    Sheets("GBC 24kV 1 MED").Activate
    Sheets("GBC 36kV 1 MED").Visible = True
    Sheets("GBC 36kV 1 MED").Activate
   
    Unload Me
    
Else
    MsgBox "Usuario o contraseña incorrectos", vbCritical
    txtPassword.Text = ""
    txtUsuario.SetFocus
End If

End Sub

Private Sub UserForm_Initialize()
   
    Me.StartUpPosition = 0
    Me.Top = 100
    Me.Left = 400
End Sub




