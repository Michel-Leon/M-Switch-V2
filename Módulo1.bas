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
public ResultadoCotas As String
'=============== Mover Bloque ========================
Public Sub MoverBloquMecanico(ByVal ref As String, ByVal BD As String, _
                       ByVal wsDestino As Worksheet, ByVal celdaDestino As String)

    Dim origen As Range, destino As Range

    ' 1) Buscar el rango por nombre del libro, sin importar la hoja
    On Error Resume Next
    Set origen = ThisWorkbook.Names(ref).RefersToRange
    ' 2) Si no existe como nombre de libro, buscarlo en la hoja BD
    If origen Is Nothing Then Set origen = ThisWorkbook.Sheets(BD).Range(ref)
    On Error GoTo errHandler

    If origen Is Nothing Then
        Err.Raise 1004, , "No se encontró el rango '" & ref & "' como nombre del libro ni en la hoja '" & BD & "'."
    End If

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
