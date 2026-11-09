Attribute VB_Name = "M�dulo1"
'============== Variables Globales ===================
Public Linea As String
Public Celda As String
Public wsActual As Worksheet
'-----------------------------------------------------
Public CodeIM As String
public CodeQM As String
Public CodeGAM As String
Public CodeDM1 As String
Public CodeGBC As String
'=============== Mover Bloque ========================
Public Sub MoverBloquMecanico(ByVal ref As String, ByVal BD As String, _
                       ByVal wsDestino As Worksheet, ByVal celdaDestino As String)

    Dim origen As Range, destino As Range, wsDatos As Worksheet

    On Error GoTo errHandler

    Set wsDatos = ThisWorkbook.Sheets(BD)
    Set origen = wsDatos.Range(ref)
    Set destino = wsDestino.Range(celdaDestino)

    origen.Copy Destination:=destino
    Application.CutCopyMode = False

    Exit Sub

    errHandler:
        MsgBox "Fallo al mover el bloque." & vbCrLf & _
            "Ref: " & ref & vbCrLf & _
            "BD: " & BD & vbCrLf & _
            "Destino: " & celdaDestino & vbCrLf & vbCrLf & _
            "Error " & Err.Number & ": " & Err.Description, vbCritical
End Sub
Private Function ResolverBloque(ByVal nombre As String) As Range
    Dim n As Name
    On Error Resume Next

    ' 1) Nombre global
    Set ResolverBloque = ThisWorkbook.Names(nombre).RefersToRange
    If Not ResolverBloque Is Nothing Then Exit Function
    Err.Clear

    ' 2) Nombre local en cualquier hoja
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        Set ResolverBloque = ws.Names(nombre).RefersToRange
        If Not ResolverBloque Is Nothing Then Exit Function
        Err.Clear
    Next ws

    On Error GoTo 0
End Function
Sub OcultarBARRAS()
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"", false)"
    Application.DisplayFormulaBar = False
    Application.DisplayStatusBar = False
    ActiveWindow.DisplayWorkbookTabs = False
    ActiveWindow.DisplayHeadings = False
    ActiveWindow.DisplayHorizontalScrollBar = False
    ActiveWindow.DisplayVerticalScrollBar = False
    ActiveSheet.Shapes("Mostrar Barras").Visible = True
    ActiveSheet.Shapes("ocultar Barras").Visible = False
End Sub
' Mostrar las barras y elementos de la interfaz
Sub MostarBARRAS()
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"", true)"
    Application.DisplayFormulaBar = True
    Application.DisplayStatusBar = True
    ActiveWindow.DisplayWorkbookTabs = True
    ActiveWindow.DisplayHeadings = True
    ActiveWindow.DisplayHorizontalScrollBar = True
    ActiveWindow.DisplayVerticalScrollBar = True
    ActiveSheet.Shapes("ocultar Barras").Visible = True
    ActiveSheet.Shapes("Mostrar Barras").Visible = False
End Sub
Sub MostrarTodasLasHojas()
    Dim ws As Worksheet
    Dim usuario As String
    Dim clave As String
    Dim intentos As Integer

    Const USUARIO_VALIDO As String = "ADMIN"
    Const CLAVE_VALIDA As String = "ADMIN2026"
    Const MAX_INTENTOS As Integer = 3

    Do
        intentos = intentos + 1

        usuario = InputBox("Ingrese el usuario:", "Acceso restringido")
        If StrPtr(usuario) = 0 Then Exit Sub   ' Cancel�

        clave = InputBox("Ingrese la contraseña:", "Acceso restringido")
        If StrPtr(clave) = 0 Then Exit Sub     ' Cancel�

        If UCase(Trim(usuario)) = USUARIO_VALIDO And Trim(clave) = CLAVE_VALIDA Then
            Exit Do
        Else
            If intentos >= MAX_INTENTOS Then
                MsgBox "Acceso denegado. Se agotaron los intentos.", vbCritical, "Error"
                Exit Sub
            End If
            MsgBox "Usuario o contraseña incorrectos." & vbNewLine & _
                   "Intento " & intentos & " de " & MAX_INTENTOS, vbExclamation, "Error"
        End If
    Loop

    ' Acceso concedido
    Application.ScreenUpdating = False
    For Each ws In ThisWorkbook.Worksheets
        ws.Visible = xlSheetVisible
    Next ws
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"", TRUE)"
    Application.DisplayFormulaBar = True
    Application.ScreenUpdating = True

    MsgBox "Acceso concedido. Todas las hojas est�n visibles.", vbInformation, "Listo"
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
    Set wsConfig = ThisWorkbook.Sheets("INICIO")

    ' Verificar que la estructura del libro no est� protegida
    If ThisWorkbook.ProtectStructure Then
        MsgBox "La estructura del libro esta protegida." & vbNewLine & _
               "Desprotejala en Revisar > Proteger libro.", vbCritical, "Error"
        Exit Sub
    End If

    ' Leer usuario y clave ingresados
    usuarioIngresado = Trim(wsLogin.Range("IJ83").Value)
    claveIngresada = Trim(wsLogin.Range("IJ98").Value)

    If usuarioIngresado = "" Or claveIngresada = "" Then
        MsgBox "Debe ingresar usuario y contraseña.", vbExclamation, "Login"
        Exit Sub
    End If

    ' Buscar el usuario en la columna E de USUARIOS
    encontrado = False
    ultimaFila = wsUsuarios.Cells(wsUsuarios.Rows.Count, "E").End(xlUp).Row

    For i = 1 To ultimaFila
        If Trim(wsUsuarios.Cells(i, "E").Value) = usuarioIngresado Then
            If Trim(wsUsuarios.Cells(i, "F").Value) = claveIngresada Then
                encontrado = True
                nombreUsuario = Trim(wsUsuarios.Cells(i, "D").Value)
                Exit For
            End If
        End If
    Next i

    If encontrado Then

        Application.ScreenUpdating = False

        ' 1) PRIMERO mostrar y activar la hoja destino
        wsConfig.Visible = xlSheetVisible
        wsConfig.Activate

        ' 2) Escribir el nombre del usuario
        wsConfig.Range("L185").Value = nombreUsuario

        ' 3) DESPU�S ocultar el login (ya hay otra hoja visible)
        'wsLogin.Visible = xlSheetVeryHidden

        ' Limpiar los campos del login por seguridad
        wsLogin.Range("IJ83").Value = ""
        wsLogin.Range("IJ98").Value = "" 
 
        Application.ScreenUpdating = True

    Else
        MsgBox "Usuario o contraseña incorrectos.", vbCritical, "Login"
        wsLogin.Range("HX98").Value = ""
    End If

End Sub