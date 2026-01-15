VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Bloqueo_UNF 
   Caption         =   "LOGIN"
   ClientHeight    =   5120
   ClientLeft      =   40
   ClientTop       =   160
   ClientWidth     =   3220
   OleObjectBlob   =   "Bloqueo_UNF.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Bloqueo_UNF"
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
    Sheets("UNIM").Visible = True
    Sheets("UNIM").Activate
    Sheets("UIM 36kv").Visible = True
    Sheets("UIM 36kv").Activate
    Sheets("UNQM").Visible = True
    Sheets("UNQM").Activate
    Sheets("UQM 36kv").Visible = True
    Sheets("UQM 36kv").Activate
    Sheets("UNGAM").Visible = True
    Sheets("UNGAM").Activate
    Sheets("UGAM 36kv").Visible = True
    Sheets("UGAM 36kv").Activate
    Sheets("UNDM1").Visible = True
    Sheets("UNDM1").Activate
    Sheets("UDM1 36kv").Visible = True
    Sheets("UDM1 36kv").Activate
    Sheets("UNGBC").Visible = True
    Sheets("UNGBC").Activate
    Sheets("UGBC 36kv").Visible = True
    Sheets("UGBC 36kv").Activate
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

