Attribute VB_Name = "M�dulo1"


Public Resultado_IM As String
Public Resultado_DM1A As String
Public Resultado_QM As String
Public Resultado_GBC As String

Sub Formulario_DM1()
    Configurador_DM1A.Show
    Limpiar_ficha_DM1A
End Sub
Sub Formulario_IM()
    Configurador_IM.Show
End Sub
Sub Formulario_QM()
    Configurador_QM.Show
End Sub
Sub fORMULARIO_GBC()
    Configurador_GBC.Show
End Sub
sub navegador()
    Navegador_Menu.Show
End sub
public sub Limpiar_ficha_DM1A()
    WITH ThisWorkbook.Sheets("FT DM1A")
        .Range("EV58").Value = ""
        .Range("EV62").Value = ""
        .Range("EV66").Value = ""
        .Range("EW140").Value = ""
        .Range("HD49").Value = ""
        .Range("HD61").Value = ""
        .Range("HD66").Value = ""
        .Range("HD71").Value = ""
        .Range("HD76").Value = ""
        .Range("HD81").Value = ""
        .Range("HD86").Value = ""
        .Range("HD96").Value = ""
        .Range("HD106").Value = ""
        .Range("HD111").Value = ""
        .Range("HD116").Value = ""
        .Range("HD121").Value = ""
    end with
End Sub

Sub abrir_formulario()

    Dim wsLogin As Worksheet
    Dim wsUsuarios As Worksheet
    Dim wsConfig As Worksheet
    Dim usuarioIngresado As String
    Dim claveIngresada As String
    Dim ultimaFila As Long
    Dim i As Long
    Dim encontrado As Boolean
    Dim nombreUsuario As String

    ' Referencias a las hojas
    Set wsLogin = ThisWorkbook.Sheets("LOGIN")
    Set wsUsuarios = ThisWorkbook.Sheets("USUARIOS")
    Set wsConfig = ThisWorkbook.Sheets("ESQUEMA")

    ' Leer usuario y clave ingresados
    usuarioIngresado = Trim(wsLogin.Range("HX83").Value)
    claveIngresada = Trim(wsLogin.Range("HX98").Value)

    ' Validar que no estén vacíos
    If usuarioIngresado = "" Or claveIngresada = "" Then
        MsgBox "Debe ingresar usuario y contraseña.", vbExclamation, "Login"
        Exit Sub
    End If

    ' Buscar el usuario en la columna E de USUARIOS
    encontrado = False
    ultimaFila = wsUsuarios.Cells(wsUsuarios.Rows.Count, "E").End(xlUp).Row

    For i = 1 To ultimaFila
        If Trim(wsUsuarios.Cells(i, "E").Value) = usuarioIngresado Then
            ' Usuario encontrado, ahora validar clave
            If Trim(wsUsuarios.Cells(i, "F").Value) = claveIngresada Then
                encontrado = True
                nombreUsuario = Trim(wsUsuarios.Cells(i, "D").Value)
                Exit For
            End If
        End If
    Next i

    If encontrado Then
        ' Escribir el nombre en M-Configurator, celda JO4
        wsConfig.Range("JO4").Value = nombreUsuario

        ' Abrir el formulario
        Selector_gama.Show
    Else
        MsgBox "Usuario o contraseña incorrectos.", vbCritical, "Login"
    End If

End Sub