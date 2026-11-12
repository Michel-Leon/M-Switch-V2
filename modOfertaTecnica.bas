Attribute VB_Name = "modOfertaTecnica"
Option Explicit

'==========================================================
' Modulo: modOfertaTecnica
' Registra en la hoja "Oferta Tecnica" una celda completa
' (encabezado + celda principal + accesorios seleccionados),
' tomando descripcion y marca de la hoja del listado.
'==========================================================

' ---------------- Configuracion ----------------
Private Const HOJA_OFERTA As String = "Oferta Tecnica"
Private Const HOJA_LISTADO As String = "ACCESORIOS"
Private Const FILA_INICIO As Long = 56
Private Const FILA_FIN As Long = 173

' Columnas en Oferta Tecnica
Private Const COL_ITEM As String = "BU"     ' Item
Private Const COL_DESC As String = "CB"     ' Descripcion
Private Const COL_REF As String = "GU"      ' Referencia
Private Const COL_MARCA As String = "IG"    ' Marca
Private Const COL_CANT As String = "JC"     ' Cantidad
Private Const COLS_FORMATO As String = "BU,CB,GU,IG,JC"

' Columnas en el listado (hoja Accesorios 24kV)
Private Const LST_DESC As String = "D"
Private Const LST_MARCA As String = "F"
Private Const LST_REF As String = "G"

' Texto de la fila de encabezado. {n} se reemplaza por el numero del item.
Private Const FORMATO_ENCABEZADO As String = "Celda S{n}H{n}"
' Marca de la celda principal (fila n.1)
Private Const MARCA_CELDA As String = "CTS"

'==========================================================
' API PUBLICA
'==========================================================

' Escribe un paquete completo:
'   fila n        -> BU = n, CB = S{n}H{n}                     (gris)
'   fila n.1      -> la primera referencia, con descripcionCelda y marca CTS
'   filas n.2 ... -> lineas de "extras" (solo descripcion; config. por defecto)
'   filas sig.    -> cada referencia restante, con descripcion y marca del listado
' refs y cants son arreglos paralelos (referencia y cantidad).
' extras (opcional): Collection o Array de textos que van solo en Descripcion.
' Las referencias vacias (opciones "No incluido") se omiten.
Public Function RegistrarCeldaEnOferta(ByVal descripcionCelda As String, _
                                       refs As Variant, cants As Variant, _
                                       Optional extras As Variant) As Boolean
    Dim ws As Worksheet, fila As Long, filaUltima As Long, n As Long, subN As Long
    Dim i As Long, ref As String, cant As String, x As Variant
    Dim desc As String, marca As String, faltantes As String
    Dim lineas As New Collection, ln As Variant

    If Not HojaExiste(HOJA_LISTADO) Then
        MsgBox "No existe la hoja '" & HOJA_LISTADO & "' en este libro." & vbLf & _
               "Copia alli el listado de accesorios.", vbCritical
        Exit Function
    End If
    Set ws = ThisWorkbook.Worksheets(HOJA_OFERTA)

    ' 1. Armar la lista de lineas: Array(descripcion, referencia, marca, cantidad)
    '    a) Celda principal
    ref = Trim$(refs(LBound(refs)))
    If Len(ref) > 0 Then
        lineas.Add Array(descripcionCelda, ref, MARCA_CELDA, CantidadTexto(cants(LBound(cants))))
    End If

    '    b) Configuracion por defecto (solo descripcion)
    If Not IsMissing(extras) Then
        For Each x In extras
            If Len(Trim$(CStr(x))) > 0 Then lineas.Add Array(Trim$(CStr(x)), "", "", "")
        Next x
    End If

    '    c) Accesorios seleccionados (descripcion y marca desde el listado)
    For i = LBound(refs) + 1 To UBound(refs)
        ref = Trim$(refs(i))
        If Len(ref) > 0 Then
            If BuscarEnListado(ref, desc, marca) Then
                lineas.Add Array(desc, ref, marca, CantidadTexto(cants(i)))
            Else
                lineas.Add Array("", ref, "", CantidadTexto(cants(i)))
                faltantes = faltantes & vbLf & "   - " & ref
            End If
        End If
    Next i

    If lineas.Count = 0 Then
        MsgBox "No hay nada seleccionado para registrar.", vbExclamation
        Exit Function
    End If

    ' 2. Verificar que cabe TODO el paquete (encabezado + lineas).
    '    Cada linea visual puede ocupar varias filas combinadas.
    fila = SiguienteFila(ws)
    filaUltima = fila
    If fila > 0 Then
        For i = 1 To lineas.Count
            filaUltima = FilaSiguiente(ws, filaUltima)
        Next i
    End If
    If fila = 0 Or filaUltima > FILA_FIN Then
        MsgBox "No hay espacio en " & HOJA_OFERTA & ": se necesitan " & lineas.Count + 1 & _
               " lineas y el limite es la fila " & FILA_FIN & ".", vbExclamation
        Exit Function
    End If

    On Error GoTo Fin
    Application.ScreenUpdating = False

    ' 3. Fila de encabezado
    n = UltimoGrupo(ws) + 1
    EscribirTexto ws.Range(COL_ITEM & fila), CStr(n)
    ws.Range(COL_DESC & fila).Value = Replace(FORMATO_ENCABEZADO, "{n}", CStr(n))
    AplicarFormatoFila ws, fila, True

    ' 4. Una fila por cada linea
    For Each ln In lineas
        fila = FilaSiguiente(ws, fila)     ' salta al siguiente bloque combinado
        subN = subN + 1

        EscribirTexto ws.Range(COL_ITEM & fila), n & "." & subN
        ws.Range(COL_DESC & fila).Value = ln(0)
        If Len(ln(1)) > 0 Then EscribirTexto ws.Range(COL_REF & fila), CStr(ln(1))
        If Len(ln(2)) > 0 Then ws.Range(COL_MARCA & fila).Value = ln(2)
        If Len(ln(3)) > 0 Then ws.Range(COL_CANT & fila).Value = Val(ln(3))
        AplicarFormatoFila ws, fila, False
    Next ln

    RegistrarCeldaEnOferta = True

Fin:
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then
        MsgBox "Error al registrar en " & HOJA_OFERTA & ": " & Err.Description, vbCritical
    ElseIf Len(faltantes) > 0 Then
        MsgBox "Estas referencias no estan en '" & HOJA_LISTADO & _
               "' y quedaron sin descripcion:" & faltantes, vbExclamation
    End If
End Function

' Cantidad como texto; si viene vacia se asume 1.
Private Function CantidadTexto(ByVal valor As Variant) As String
    CantidadTexto = Trim$(CStr(valor))
    If Len(CantidadTexto) = 0 Then CantidadTexto = "1"
End Function

' Borra todos los items del rango y deja el formato en blanco.
Public Sub LimpiarOfertaTecnica()
    Dim ws As Worksheet, f As Long, col As Variant

    If MsgBox("Se borraran los items de las filas " & FILA_INICIO & " a " & _
              FILA_FIN & ". Continuar?", vbYesNo + vbQuestion) = vbNo Then Exit Sub

    Set ws = ThisWorkbook.Worksheets(HOJA_OFERTA)
    Application.ScreenUpdating = False
    For f = FILA_INICIO To FILA_FIN
        For Each col In Split(COLS_FORMATO, ",")
            ws.Range(col & f).MergeArea.ClearContents
        Next col
        AplicarFormatoFila ws, f, False
    Next f
    Application.ScreenUpdating = True
End Sub

'==========================================================
' FUNCIONES DE APOYO
'==========================================================

' Busca la referencia en el listado (sin distinguir mayusculas).
' Devuelve True y llena desc y marca si la encuentra.
Private Function BuscarEnListado(ByVal ref As String, ByRef desc As String, _
                                 ByRef marca As String) As Boolean
    Dim wl As Worksheet, f As Long, ultima As Long
    Set wl = ThisWorkbook.Worksheets(HOJA_LISTADO)
    ultima = wl.Cells(wl.Rows.Count, LST_REF).End(xlUp).Row

    For f = 2 To ultima
        If StrComp(Trim$(CStr(wl.Range(LST_REF & f).Value)), ref, vbTextCompare) = 0 Then
            desc = Trim$(CStr(wl.Range(LST_DESC & f).Value))
            marca = Trim$(CStr(wl.Range(LST_MARCA & f).Value))
            BuscarEnListado = True
            Exit Function
        End If
    Next f
End Function

Private Function HojaExiste(ByVal nombre As String) As Boolean
    Dim sh As Object
    On Error Resume Next
    Set sh = ThisWorkbook.Worksheets(nombre)
    HojaExiste = Not sh Is Nothing
End Function

' Primera fila del bloque siguiente al ultimo item escrito en BU.
' Devuelve 0 si el rango esta lleno.
Private Function SiguienteFila(ws As Worksheet) As Long
    Dim f As Long, sig As Long
    For f = FILA_FIN To FILA_INICIO Step -1
        If Len(LeerItem(ws, f)) > 0 Then
            sig = FilaSiguiente(ws, f)
            If sig > FILA_FIN Then
                SiguienteFila = 0
            Else
                SiguienteFila = sig
            End If
            Exit Function
        End If
    Next f
    SiguienteFila = FILA_INICIO
End Function

' Primera fila de la linea visual siguiente. Si BU esta combinada
' verticalmente (por ejemplo 3 filas por linea), salta el bloque completo.
Private Function FilaSiguiente(ws As Worksheet, ByVal fila As Long) As Long
    Dim blq As Range
    Set blq = ws.Range(COL_ITEM & fila).MergeArea
    FilaSiguiente = blq.Row + blq.Rows.Count
End Function

' Mayor item principal (entero, sin punto) registrado en BU.
Private Function UltimoGrupo(ws As Worksheet) As Long
    Dim f As Long, txt As String, maxN As Long
    For f = FILA_INICIO To FILA_FIN
        txt = LeerItem(ws, f)
        If Len(txt) > 0 And InStr(txt, ".") = 0 Then
            If Val(txt) > maxN Then maxN = Val(txt)
        End If
    Next f
    UltimoGrupo = maxN
End Function

' Lee el item de BU como texto (coma decimal -> punto).
Private Function LeerItem(ws As Worksheet, ByVal fila As Long) As String
    LeerItem = Replace(Trim$(CStr(ws.Range(COL_ITEM & fila).Value)), ",", ".")
End Function

' Escribe como TEXTO: 1.10 no se convierte en 1.1 ni en fecha,
' y referencias numericas largas no pasan a notacion cientifica.
Private Sub EscribirTexto(celda As Range, ByVal valor As String)
    celda.MergeArea.NumberFormat = "@"
    celda.Value = valor
End Sub

' Encabezado: fondo #D9D9D9, letra #808080. Normal: fondo blanco, letra negra.
Private Sub AplicarFormatoFila(ws As Worksheet, ByVal fila As Long, ByVal esEncabezado As Boolean)
    Dim col As Variant, rng As Range
    For Each col In Split(COLS_FORMATO, ",")
        Set rng = ws.Range(col & fila).MergeArea
        If esEncabezado Then
            rng.Interior.Color = RGB(217, 217, 217)
            rng.Font.Color = RGB(128, 128, 128)
            rng.Font.Bold = True
        Else
            rng.Interior.Color = RGB(255, 255, 255)
            rng.Font.Color = RGB(0, 0, 0)
        End If
    Next col
End Sub
