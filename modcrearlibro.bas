Attribute VB_Name = "modcrearlibro"
Option Explicit

'==========================================================================
'  CONFIGURACI?N
'==========================================================================
' Nombre de la hoja donde est?n los ComboBox (HojaENT, Esquema, etc.)
Private Const HOJA_CONTROL As String = "Documentacion"          '<-- CAMBIAR por el nombre real

'Nombre oculto donde se guarda la ruta del �ltimo libro creado
Private Const NOMBRE_RUTA As String = "RutaLibroGenerado"

' Formato de las hojas del libro nuevo 
Private Const ANCHO_COLUMNA As Double = 0.38
Private Const ALTO_FILA As Double = 3.6

' ---- NOMBRES DE LAS �REAS DE IMPRESI�N (�nico lugar a modificar) ----
' Puede ser un nombre definido o el nombre de una hoja con �rea de impresi�n.
' Se busca sin importar may�sculas, tildes, espacios o guiones bajos.
' Orden = orden de los ComboBox: HojaENT|Esquema|Unifilar|Mecanico|Resumen|Oferta1
Private Const AREAS_BASE As String = "Hoja de entrada|Esquema|Unifilar|Mecanicos|Resumen_tecnico|Oferta"
' Orden = elementos del ComboBox Ficha_TEC
Private Const AREAS_FT As String = "FT_DM1A|FT_QM|FT_IM|FT_GAM|FT_GBC"
' Orden = elementos del ComboBox Complemento
Private Const AREAS_COMP As String = "Comp_DM1A|Comp_QM|Comp_IM|Comp_GAM_0|Comp_GAM_2|Comp_GBC_A|Comp_GBC_B|Comp_GBC_C|Comp_GBC_D"
' �rea que a�ade el ComboBox Oferta2
Private Const AREA_OFERTA_ADIC As String = "Oferta"


'==========================================================================
'  MACRO 1: CREAR EL LIBRO (Hoja de entrada ... Oferta 1)
'==========================================================================
Public Sub CrearLibro()
    Dim wsCtrl As Worksheet
    Dim areas As Collection
    Dim ruta As Variant
    Dim wb As Workbook
    Dim nombre As Variant
    Dim primera As Boolean

    Set wsCtrl = ThisWorkbook.Worksheets(HOJA_CONTROL)
    Set areas = New Collection

    ' ListIndex = 0  ->  "A�adir"
    Dim combos As Variant, nombresBase As Variant, k As Long
    combos = Array("HojaENT", "Esquema", "Unifilar", "Mecanico", "Resumen", "Oferta1")
    nombresBase = Split(AREAS_BASE, "|")
    For k = 0 To UBound(combos)
        If Seleccion(wsCtrl, CStr(combos(k))) = 0 Then areas.Add nombresBase(k)
    Next k

    If areas.Count = 0 Then
        MsgBox "No hay ninguna hoja marcada como 'Añadir'.", vbExclamation
        Exit Sub
    End If

    ' Preguntar ubicaci�n y nombre del libro
    ruta = Application.GetSaveAsFilename( _
                InitialFileName:="Oferta.xlsx", _
                FileFilter:="Libro de Excel (*.xlsx), *.xlsx", _
                Title:="Seleccione d�nde guardar el libro y escriba su nombre")
    If VarType(ruta) = vbBoolean Then Exit Sub          ' Cancel�

    If Dir(CStr(ruta)) <> "" Then
        If MsgBox("El archivo ya existe:" & vbCrLf & ruta & vbCrLf & vbCrLf & _
                  "�Desea reemplazarlo?", vbYesNo + vbQuestion) = vbNo Then Exit Sub
    End If

    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.Calculate                       ' valores de f�rmulas actualizados

    Set wb = Workbooks.Add(xlWBATWorksheet)              ' Libro con una sola hoja
    primera = True
    For Each nombre In areas
        CopiarArea CStr(nombre), wb, primera
        primera = False
    Next nombre

    RomperVinculos wb
    wb.Worksheets(1).Activate

    Application.DisplayAlerts = False
    wb.SaveAs Filename:=CStr(ruta), FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True

    GuardarRuta CStr(ruta)
    MsgBox "Libro creado correctamente:" & vbCrLf & ruta, vbInformation

Salir:
    Application.CutCopyMode = False
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub
ErrHandler:
    MsgBox "Error al crear el libro: " & Err.Description, vbCritical
    Resume Salir
End Sub


'==========================================================================
'  MACRO 2: A�ADIR HOJAS ADICIONALES (Ficha t�cnica, Complemento, Oferta 2)
'==========================================================================
Public Sub AnadirHojas()
    Dim wsCtrl As Worksheet
    Dim areas As Collection
    Dim wb As Workbook
    Dim nombre As Variant
    Dim idx As Long

    Set wsCtrl = ThisWorkbook.Worksheets(HOJA_CONTROL)
    Set areas = New Collection

    idx = Seleccion(wsCtrl, "Ficha_TEC")
    If idx >= 0 Then areas.Add Split(AREAS_FT, "|")(idx)

    idx = Seleccion(wsCtrl, "Complemento")
    If idx >= 0 Then areas.Add Split(AREAS_COMP, "|")(idx)

    If Seleccion(wsCtrl, "Oferta2") = 0 Then areas.Add AREA_OFERTA_ADIC

    If areas.Count = 0 Then
        MsgBox "No hay ninguna hoja adicional seleccionada.", vbExclamation
        Exit Sub
    End If

    Set wb = ObtenerLibroDestino()
    If wb Is Nothing Then Exit Sub

    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.Calculate                       ' valores de f�rmulas actualizados

    For Each nombre In areas
        CopiarArea CStr(nombre), wb, False
    Next nombre

    RomperVinculos wb
    wb.Save
    MsgBox "Hojas a�adidas correctamente a:" & vbCrLf & wb.FullName, vbInformation

Salir:
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    Exit Sub
ErrHandler:
    MsgBox "Error al a�adir hojas: " & Err.Description, vbCritical
    Resume Salir
End Sub


'==========================================================================
'  PROCEDIMIENTOS AUXILIARES
'==========================================================================

' Devuelve el ListIndex del ComboBox (-1 si no hay selecci�n)
Private Function Seleccion(ByVal ws As Worksheet, ByVal nombreCombo As String) As Long
    Seleccion = ws.OLEObjects(nombreCombo).Object.ListIndex
End Function

' Copia el rango con nombre a una hoja del libro destino y la configura
Private Sub CopiarArea(ByVal nombreArea As String, ByVal wb As Workbook, ByVal usarPrimeraHoja As Boolean)
    Dim rngOrigen As Range
    Dim rngDestino As Range
    Dim ws As Worksheet

    Set rngOrigen = ObtenerRango(nombreArea)
    If rngOrigen Is Nothing Then
        MsgBox "No se encontro el Area de impresión con nombre '" & nombreArea & "'.", vbExclamation
        Exit Sub
    End If

    If usarPrimeraHoja Then
        Set ws = wb.Worksheets(1)
    Else
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    End If
    ws.Name = NombreUnico(wb, nombreArea)

    ' Formato de celdas (antes de pegar im�genes para que no se deformen)
    ws.Cells.ColumnWidth = ANCHO_COLUMNA
    ws.Cells.RowHeight = ALTO_FILA

    ' Si el �rea de impresi�n corta una celda combinada, se ampl�a para incluirla
    Set rngOrigen = ExpandirCombinadas(rngOrigen)
    Set rngDestino = ws.Range("A1").Resize(rngOrigen.Rows.Count, rngOrigen.Columns.Count)

    ' PasteSpecial e im�genes requieren que la hoja destino est� activa
    wb.Activate
    ws.Activate

    ' 1) Copia completa: formatos, celdas combinadas, textos y valores fijos.
    '    (Copiar con Destination crea las combinadas igual que el original,
    '     as� no aparece el error "celdas combinadas del mismo tama�o")
    rngOrigen.Copy Destination:=rngDestino
    Application.CutCopyMode = False

    ' Las formas que Excel copia solo se eliminan: se copian despu�s en su posici�n exacta
    BorrarFormas ws

    ' 2) Mapa de f�rmulas: cada f�rmula se reemplaza por su valor,
    '    bloque por bloque y respetando celdas combinadas e im�genes en celda
    ConvertirFormulasEnValores rngOrigen, rngDestino

    ' 3) Im�genes flotantes (logos, fotos, gr�ficos) en la misma posici�n y tama�o
    CopiarImagenes rngOrigen, rngDestino

    ConfigurarImpresion ws, rngDestino
    ActiveWindow.DisplayGridlines = False
    ws.Range("A1").Select
End Sub

' =====================================================================
'  MAPA DE F�RMULAS -> VALORES
'  Paso 1 (origen):  SpecialCells detecta las f�rmulas del �rea y cada
'                    bloque se escribe como valor de una sola vez.
'                    Si un bloque falla (combinadas, matriciales...) se
'                    hace celda por celda.
'  Paso 2 (origen):  f�rmulas con desbordamiento (Excel 365).
'  Paso 3 (destino): verificaci�n: cualquier f�rmula que haya quedado en
'                    el libro nuevo se reemplaza por el valor del origen.
' =====================================================================
Private Sub ConvertirFormulasEnValores(ByVal rngOrigen As Range, ByVal rngDestino As Range)
    Dim formulas As Range, area As Range, c As Range, r As Range
    Dim oCelda As Object

    On Error Resume Next
    Set formulas = rngOrigen.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0

    If Not formulas Is Nothing Then
        ' ---- Paso 1: por bloques, con respaldo celda por celda ----
        For Each area In formulas.Areas
            If AreaTieneImagen(area) Then
                For Each c In area.Cells
                    CongelarCelda c, rngOrigen, rngDestino
                Next c
            Else
                On Error Resume Next
                Err.Clear
                DestinoDe(area, rngOrigen, rngDestino).Value = area.Value
                If Err.Number <> 0 Then
                    Err.Clear
                    On Error GoTo 0
                    For Each c In area.Cells
                        CongelarCelda c, rngOrigen, rngDestino
                    Next c
                End If
                On Error GoTo 0
            End If
        Next area

        ' ---- Paso 2: rangos desbordados (Excel 365) ----
        On Error Resume Next
        For Each c In formulas.Cells
            Set oCelda = c
            If oCelda.HasSpill Then
                Set r = Nothing
                Set r = Intersect(oCelda.SpillingToRange, rngOrigen)
                If Not r Is Nothing Then DestinoDe(r, rngOrigen, rngDestino).Value = r.Value
            End If
        Next c
        On Error GoTo 0
    End If

    ' ---- Paso 3: verificaci�n en el libro nuevo ----
    VerificarSinFormulas rngOrigen, rngDestino
End Sub

' Escribe el valor de UNA celda del origen en su equivalente del destino
Private Sub CongelarCelda(ByVal c As Range, ByVal rngOrigen As Range, ByVal rngDestino As Range)
    Dim r As Range, d As Range

    On Error Resume Next
    If c.HasArray Then
        ' F�rmula matricial: se escribe el bloque completo de la matriz
        Set r = Intersect(c.CurrentArray, rngOrigen)
        If Not r Is Nothing Then DestinoDe(r, rngOrigen, rngDestino).Value = r.Value

    ElseIf c.MergeCells Then
        ' Combinada: solo la celda principal guarda valor
        If c.Address = c.MergeArea.Cells(1, 1).Address Then
            Set d = DestinoDe(c.MergeArea, rngOrigen, rngDestino)
            If TieneImagenEnCelda(c) Then
                PegarValorConImagen c.MergeArea, d
            Else
                d.Cells(1, 1).Value = c.Value
            End If
        End If

    Else
        Set d = DestinoDe(c, rngOrigen, rngDestino)
        If TieneImagenEnCelda(c) Then
            PegarValorConImagen c, d
        Else
            d.Value = c.Value
        End If
    End If
    On Error GoTo 0
End Sub

' Imagen dentro de celda / IMAGEN(): pegar como valor conserva la imagen
Private Sub PegarValorConImagen(ByVal o As Range, ByVal d As Range)
    On Error Resume Next
    Err.Clear
    o.Copy
    d.PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    If Err.Number <> 0 Then d.Cells(1, 1).Value = o.Cells(1, 1).Value
    On Error GoTo 0
End Sub

' Busca f�rmulas que hayan quedado en el DESTINO y las reemplaza por el
' valor del origen (no depende de que el origen permita SpecialCells)
Private Sub VerificarSinFormulas(ByVal rngOrigen As Range, ByVal rngDestino As Range)
    Dim quedan As Range, c As Range, o As Range

    On Error Resume Next
    Set quedan = rngDestino.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0
    If quedan Is Nothing Then Exit Sub

    For Each c In quedan.Cells
        Set o = rngOrigen.Cells(c.Row - rngDestino.Row + 1, c.Column - rngDestino.Column + 1)
        If c.HasFormula Then CongelarCelda o, rngOrigen, rngDestino
        ' �ltimo recurso: congelar con el valor mostrado en el destino
        If c.HasFormula Then
            On Error Resume Next
            If c.MergeCells Then
                If c.Address = c.MergeArea.Cells(1, 1).Address Then c.Value = o.Value
            Else
                c.Value = o.Value
            End If
            On Error GoTo 0
        End If
    Next c
End Sub

Private Function AreaTieneImagen(ByVal a As Range) As Boolean
    Dim c As Range
    For Each c In a.Cells
        If TieneImagenEnCelda(c) Then AreaTieneImagen = True: Exit Function
    Next c
End Function

' Imagen dentro de celda, funci�n IMAGEN() o tipo de dato enriquecido
Private Function TieneImagenEnCelda(ByVal c As Range) As Boolean
    Dim rico As Boolean
    On Error Resume Next
    rico = c.HasRichDataType            ' Excel 365; en versiones antiguas se ignora
    On Error GoTo 0
    TieneImagenEnCelda = rico Or (InStr(1, c.Formula, "IMAGE(", vbTextCompare) > 0)
End Function

' Rango equivalente en el destino
Private Function DestinoDe(ByVal r As Range, ByVal rngOrigen As Range, ByVal rngDestino As Range) As Range
    Set DestinoDe = rngDestino.Cells(1, 1).Offset(r.Row - rngOrigen.Row, r.Column - rngOrigen.Column) _
                              .Resize(r.Rows.Count, r.Columns.Count)
End Function

' Ampl�a el rango hasta que ninguna celda combinada quede cortada en el borde
Private Function ExpandirCombinadas(ByVal rng As Range) As Range
    Dim r As Range, u As Range, c As Range, borde As Range
    Dim cambio As Boolean

    Set r = rng.Areas(1)
    Do
        cambio = False
        Set borde = Union(r.Rows(1), r.Rows(r.Rows.Count), r.Columns(1), r.Columns(r.Columns.Count))
        Set u = r
        For Each c In borde.Cells
            If c.MergeCells Then Set u = Union(u, c.MergeArea)
        Next c
        Set u = Envolvente(u)
        If u.Address <> r.Address Then
            Set r = u
            cambio = True
        End If
    Loop While cambio

    Set ExpandirCombinadas = r
End Function

' Rect�ngulo que contiene todas las �reas de un rango
Private Function Envolvente(ByVal u As Range) As Range
    Dim a As Range
    Dim f1 As Long, c1 As Long, f2 As Long, c2 As Long

    f1 = u.Areas(1).Row: c1 = u.Areas(1).Column
    f2 = f1: c2 = c1
    For Each a In u.Areas
        If a.Row < f1 Then f1 = a.Row
        If a.Column < c1 Then c1 = a.Column
        If a.Row + a.Rows.Count - 1 > f2 Then f2 = a.Row + a.Rows.Count - 1
        If a.Column + a.Columns.Count - 1 > c2 Then c2 = a.Column + a.Columns.Count - 1
    Next a
    Set Envolvente = u.Worksheet.Range(u.Worksheet.Cells(f1, c1), u.Worksheet.Cells(f2, c2))
End Function

Private Sub BorrarFormas(ByVal ws As Worksheet)
    Dim i As Long
    For i = ws.Shapes.Count To 1 Step -1
        ws.Shapes(i).Delete
    Next i
End Sub

' Copia las im�genes/formas cuya esquina superior izquierda est� dentro del �rea,
' manteniendo la misma posici�n relativa a la celda y el mismo tama�o
Private Sub CopiarImagenes(ByVal rngOrigen As Range, ByVal rngDestino As Range)
    Dim wsO As Worksheet, wsD As Worksheet
    Dim shp As Shape, nueva As Shape
    Dim celdaO As Range, celdaD As Range

    Set wsO = rngOrigen.Worksheet
    Set wsD = rngDestino.Worksheet

    For Each shp In wsO.Shapes
        If EsImagenCopiable(shp) Then
            Set celdaO = shp.TopLeftCell
            If Not Intersect(celdaO, rngOrigen) Is Nothing Then

                ' Celda equivalente en el destino
                Set celdaD = rngDestino.Cells(1, 1).Offset(celdaO.Row - rngOrigen.Row, _
                                                          celdaO.Column - rngOrigen.Column)

                Set nueva = PegarForma(shp, wsD)
                If Not nueva Is Nothing Then
                    On Error Resume Next
                    nueva.LockAspectRatio = msoFalse
                    nueva.Left = celdaD.Left + (shp.Left - celdaO.Left)
                    nueva.Top = celdaD.Top + (shp.Top - celdaO.Top)
                    nueva.Width = shp.Width
                    nueva.Height = shp.Height
                    nueva.LockAspectRatio = shp.LockAspectRatio
                    nueva.Placement = shp.Placement
                    nueva.Name = shp.Name
                    On Error GoTo 0
                End If
            End If
        End If
    Next shp
End Sub

' Excluye comentarios, controles (ComboBox, botones) y formas ocultas
Private Function EsImagenCopiable(ByVal shp As Shape) As Boolean
    Select Case shp.Type
        Case msoComment, msoFormControl, msoOLEControlObject
            EsImagenCopiable = False
        Case Else
            EsImagenCopiable = shp.Visible
    End Select
End Function

' Im�genes vinculadas, gr�ficos o "im�genes de c�mara" se pegan como imagen fija
' para que no dependan del libro original
Private Function DebeCopiarseComoImagen(ByVal shp As Shape) As Boolean
    Dim f As String
    If shp.Type = msoLinkedPicture Or shp.Type = msoChart Then
        DebeCopiarseComoImagen = True
        Exit Function
    End If
    On Error Resume Next
    f = shp.DrawingObject.Formula
    On Error GoTo 0
    DebeCopiarseComoImagen = (Len(f) > 0)
End Function

' Copia y pega una forma, con reintentos (el portapapeles a veces falla)
Private Function PegarForma(ByVal shp As Shape, ByVal wsD As Worksheet) As Shape
    Dim intento As Integer
    Dim n As Long
    Dim comoImagen As Boolean

    comoImagen = DebeCopiarseComoImagen(shp)
    n = wsD.Shapes.Count

    For intento = 1 To 5
        On Error Resume Next
        Err.Clear
        If comoImagen Then
            Application.ScreenUpdating = True        ' evita im�genes en blanco
            shp.CopyPicture Appearance:=xlScreen, Format:=xlPicture
            Application.ScreenUpdating = False
        Else
            shp.Copy
        End If
        DoEvents
        wsD.Paste
        If Err.Number = 0 And wsD.Shapes.Count > n Then
            On Error GoTo 0
            Set PegarForma = wsD.Shapes(wsD.Shapes.Count)
            Application.CutCopyMode = False
            Exit Function
        End If
        On Error GoTo 0
        DoEvents
    Next intento

    Application.CutCopyMode = False
End Function

' Configuraci�n de p�gina: �rea de impresi�n, m�rgenes 0, centrado, horizontal
Private Sub ConfigurarImpresion(ByVal ws As Worksheet, ByVal rng As Range)
    On Error Resume Next
    Application.PrintCommunication = False
    On Error GoTo 0

    With ws.PageSetup
        .PrintArea = rng.Address
        .Orientation = xlLandscape
        .LeftMargin = 0
        .RightMargin = 0
        .TopMargin = 0
        .BottomMargin = 0
        .HeaderMargin = 0
        .FooterMargin = 0
        .CenterHorizontally = True
        .CenterVertically = True
        ' Ajustar a 1 p�gina (borrar estas 3 l�neas si no lo desea)
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
    End With

    On Error Resume Next
    Application.PrintCommunication = True
    On Error GoTo 0
End Sub

' Busca el �rea por nombre definido (libro u hoja) o por nombre de hoja (su �rea
' de impresi�n). Ignora may�sculas, tildes, espacios y guiones bajos.
Private Function ObtenerRango(ByVal nombre As String) As Range
    Dim r As Range
    Dim nm As Name
    Dim sh As Worksheet
    Dim objetivo As String, nombreCorto As String

    objetivo = Normalizar(nombre)

    ' 1) Nombres definidos (se omiten los autom�ticos Print_Area / Print_Titles)
    For Each nm In ThisWorkbook.Names
        nombreCorto = Mid$(nm.Name, InStrRev(nm.Name, "!") + 1)
        If Normalizar(nombreCorto) = objetivo Then
            On Error Resume Next
            Set r = nm.RefersToRange
            On Error GoTo 0
            If Not r Is Nothing Then Exit For
        End If
    Next nm

    ' 2) Hoja con ese nombre -> su �rea de impresi�n
    If r Is Nothing Then
        For Each sh In ThisWorkbook.Worksheets
            If Normalizar(sh.Name) = objetivo Then
                If sh.PageSetup.PrintArea <> "" Then
                    On Error Resume Next
                    Set r = sh.Range(sh.PageSetup.PrintArea).Areas(1)
                    On Error GoTo 0
                End If
                Exit For
            End If
        Next sh
    End If

    If Not r Is Nothing Then Set r = r.Areas(1)
    Set ObtenerRango = r
End Function

' Texto comparable: min�sculas, sin tildes, sin espacios ni guiones bajos
Private Function Normalizar(ByVal t As String) As String
    Dim conTilde As Variant, sinTilde As Variant, i As Long
    t = LCase$(Trim$(t))
    conTilde = Array("�", "�", "�", "�", "�", "�", "�")
    sinTilde = Array("a", "e", "i", "o", "u", "u", "n")
    For i = 0 To UBound(conTilde)
        t = Replace(t, conTilde(i), sinTilde(i))
    Next i
    t = Replace(t, " ", "")
    t = Replace(t, "_", "")
    t = Replace(t, "-", "")
    Normalizar = t
End Function

' Evita nombres de hoja repetidos: Oferta, Oferta (2), Oferta (3)...
Private Function NombreUnico(ByVal wb As Workbook, ByVal base As String) As String
    Dim candidato As String
    Dim n As Long

    candidato = Left$(base, 31)
    n = 1
    Do While ExisteHoja(wb, candidato)
        n = n + 1
        candidato = Left$(base, 31 - Len(" (" & n & ")")) & " (" & n & ")"
    Loop
    NombreUnico = candidato
End Function

Private Function ExisteHoja(ByVal wb As Workbook, ByVal nombre As String) As Boolean
    Dim sh As Object
    On Error Resume Next
    Set sh = wb.Sheets(nombre)
    On Error GoTo 0
    ExisteHoja = Not sh Is Nothing
End Function

' Quita v�nculos al libro original (nombres y referencias externas)
Private Sub RomperVinculos(ByVal wb As Workbook)
    Dim links As Variant
    Dim i As Long

    On Error Resume Next
    links = wb.LinkSources(xlExcelLinks)
    If Not IsEmpty(links) Then
        For i = LBound(links) To UBound(links)
            wb.BreakLink Name:=links(i), Type:=xlLinkTypeExcelLinks
        Next i
    End If

    For i = wb.Names.Count To 1 Step -1
        If InStr(wb.Names(i).RefersTo, "[") > 0 Or InStr(wb.Names(i).RefersTo, "#REF!") > 0 Then
            wb.Names(i).Delete
        End If
    Next i
    On Error GoTo 0
End Sub

' Guarda la ruta del libro creado para usarla en AnadirHojas
Private Sub GuardarRuta(ByVal ruta As String)
    ThisWorkbook.Names.Add Name:=NOMBRE_RUTA, _
        RefersTo:="=""" & Replace(ruta, """", """""") & """", Visible:=False
End Sub

Private Function LeerRuta() As String
    Dim s As String
    On Error Resume Next
    s = ThisWorkbook.Names(NOMBRE_RUTA).RefersTo
    On Error GoTo 0
    If Len(s) > 3 Then LeerRuta = Replace(Mid$(s, 3, Len(s) - 3), """""", """")
End Function

' Localiza el libro creado: abierto, en la ruta guardada, o lo pide al usuario
Private Function ObtenerLibroDestino() As Workbook
    Dim ruta As String
    Dim archivo As String
    Dim wb As Workbook
    Dim sel As Variant

    ruta = LeerRuta()
    If ruta <> "" Then
        archivo = Mid$(ruta, InStrRev(ruta, Application.PathSeparator) + 1)
        On Error Resume Next
        Set wb = Workbooks(archivo)
        On Error GoTo 0
        If wb Is Nothing Then
            If Dir(ruta) <> "" Then Set wb = Workbooks.Open(ruta)
        End If
    End If

    If wb Is Nothing Then
        sel = Application.GetOpenFilename( _
                FileFilter:="Libro de Excel (*.xlsx), *.xlsx", _
                Title:="Seleccione el libro al que desea a�adir las hojas")
        If VarType(sel) = vbBoolean Then Exit Function
        Set wb = Workbooks.Open(CStr(sel))
        GuardarRuta CStr(sel)
    End If

    Set ObtenerLibroDestino = wb
End Function


' =====================================================================
'  MACRO DE APOYO: mapa de f�rmulas, combinadas e im�genes por �rea
'  Usa las mismas �reas de impresi�n que CrearLibro y AnadirHojas
' =====================================================================
Public Sub Mapear_Formulas()
    Dim listas As Variant, nombres As Variant, n As Variant, lista As Variant
    Dim wbRep As Workbook, wsRep As Worksheet, ws As Worksheet
    Dim rngArea As Range, rngAmpliada As Range, celdas As Range, a As Range, c As Range
    Dim shp As Shape
    Dim fila As Long, nImg As Long, nFlot As Long
    Dim combinadas As String, cortada As String, vistos As String

    listas = Array(AREAS_BASE, AREAS_FT, AREAS_COMP, AREA_OFERTA_ADIC)

    Set wbRep = Workbooks.Add(xlWBATWorksheet)
    Set wsRep = wbRep.Worksheets(1)
    wsRep.Name = "Mapa_Formulas"
    wsRep.Range("A1:J1").Value = Array("�rea", "Hoja", "Rango del �rea", "�rea ajustada", _
        "Corta combinadas", "Rango con f�rmula", "N� celdas", "Tiene combinadas", _
        "Im�genes en celda", "Im�genes flotantes")
    wsRep.Range("A1:J1").Font.Bold = True
    fila = 2
    vistos = "|"

    For Each lista In listas
        nombres = Split(lista, "|")
        For Each n In nombres
            ' Evitar repetir un �rea que aparece en dos listas (ej. Oferta)
            If InStr(vistos, "|" & Normalizar(CStr(n)) & "|") = 0 Then
                vistos = vistos & Normalizar(CStr(n)) & "|"

                Set rngArea = ObtenerRango(CStr(n))
                If rngArea Is Nothing Then
                    wsRep.Cells(fila, 1).Resize(1, 3).Value = Array(CStr(n), "-", "(NO ENCONTRADA)")
                    wsRep.Cells(fila, 1).Resize(1, 10).Interior.Color = RGB(255, 199, 206)
                    fila = fila + 1
                Else
                    Set ws = rngArea.Worksheet
                    Set rngAmpliada = ExpandirCombinadas(rngArea)
                    cortada = IIf(rngAmpliada.Address = rngArea.Address, "No", "S�")

                    nFlot = 0
                    For Each shp In ws.Shapes
                        If EsImagenCopiable(shp) Then
                            If Not Intersect(shp.TopLeftCell, rngAmpliada) Is Nothing Then nFlot = nFlot + 1
                        End If
                    Next shp

                    Set celdas = Nothing
                    On Error Resume Next
                    Set celdas = rngAmpliada.SpecialCells(xlCellTypeFormulas)
                    On Error GoTo 0

                    If celdas Is Nothing Then
                        wsRep.Cells(fila, 1).Resize(1, 10).Value = Array(CStr(n), ws.Name, _
                            rngArea.Address(0, 0), rngAmpliada.Address(0, 0), cortada, _
                            "(sin f�rmulas)", 0, "-", 0, nFlot)
                        fila = fila + 1
                    Else
                        For Each a In celdas.Areas
                            If IsNull(a.MergeCells) Then
                                combinadas = "Parcial"
                            ElseIf a.MergeCells Then
                                combinadas = "S�"
                            Else
                                combinadas = "No"
                            End If
                            nImg = 0
                            For Each c In a.Cells
                                If TieneImagenEnCelda(c) Then nImg = nImg + 1
                            Next c
                            wsRep.Cells(fila, 1).Resize(1, 10).Value = Array(CStr(n), ws.Name, _
                                rngArea.Address(0, 0), rngAmpliada.Address(0, 0), cortada, _
                                a.Address(0, 0), a.Cells.Count, combinadas, nImg, nFlot)
                            fila = fila + 1
                        Next a
                    End If
                End If
            End If
        Next n
    Next lista

    wsRep.Columns("A:J").AutoFit
    MsgBox "Mapa de f�rmulas generado.", vbInformation
End Sub
