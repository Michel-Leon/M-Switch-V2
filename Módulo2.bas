Attribute VB_Name = "M�dulo2"
Public LibroProyecto As Workbook 
Public NombreLibroProyecto As String

Public RutaLibroProyecto As String

Public LibroImportado As Workbook
Public HojaImportada As Worksheet
Public AreaImportacion As String
 
Sub MostrarTodasLasHojas()
    Dim ws As Worksheet 
    For Each ws In ThisWorkbook.Worksheets
        ws.Visible = xlSheetVisible
    Next ws
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"", TRUE)"
    Application.DisplayFormulaBar = True
End Sub 
Sub OcultarBARRAS()
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"", false)"
    Application.DisplayFormulaBar = False
    Application.DisplayStatusBar = False
    ActiveWindow.DisplayHeadings = False
    ActiveWindow.DisplayWorkbookTabs = False
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
    ActiveWindow.DisplayHeadings = True
    ActiveWindow.DisplayWorkbookTabs = True
    ActiveWindow.DisplayHorizontalScrollBar = True
    ActiveWindow.DisplayVerticalScrollBar = True
    ActiveSheet.Shapes("ocultar Barras").Visible = True
    ActiveSheet.Shapes("Mostrar Barras").Visible = False
End Sub
Sub CambiarZoomEnTodasLasHojas()
    Dim zoomStr As String
    Dim zoomVal As Integer
    Dim ws As Worksheet

    ' Solicitar el valor de zoom al usuario
    zoomStr = InputBox("Ingrese el nivel de zoom deseado (ej. 100 para 100%):", "Cambiar Zoom")

    ' Cancelar si el usuario presiona Cancelar o deja vacío
    If zoomStr = "" Then Exit Sub

    ' Validar que el valor ingresado sea numérico
    If Not IsNumeric(zoomStr) Then
        MsgBox "Por favor ingrese un número válido.", vbExclamation
        Exit Sub
    End If

    zoomVal = CInt(zoomStr)

    ' Validar que el zoom esté dentro de un rango razonable
    If zoomVal < 10 Or zoomVal > 400 Then
        MsgBox "El nivel de zoom debe estar entre 10 y 400.", vbExclamation
        Exit Sub
    End If

    ' Aplicar el zoom a todas las hojas
    For Each ws In ThisWorkbook.Worksheets
        ws.Activate
        ActiveWindow.Zoom = zoomVal
    Next ws

    MsgBox "Zoom ajustado a " & zoomVal & "% en todas las hojas.", vbInformation
End Sub
'=============== SECCION PARA CREAR LIBRO DE PROYECTO ========================
' Crea un nuevo libro con el área de impresión de la hoja activa
Sub crearnuevolibroconproyecto()
    Dim nombreProyecto As String
    Dim libroOrigen As Workbook
    Dim hojaOrigen As Worksheet
    Dim nuevoLibro As Workbook
    Dim hojaDestino As Worksheet
    Dim rangoImpresion As Range
    Dim rutaGuardar As String
    Dim areaImpresion As String
    Dim ws As Worksheet
    
    ' Variables para pegar solo valores en rango específico
    Dim filaOffset As Long, colOffset As Long
    Dim celdaInicioDestino As Range
    Dim rangoValores As Range
    
    Set libroOrigen = ThisWorkbook
    Set hojaOrigen = libroOrigen.ActiveSheet

    ' Validar si existe un área de impresión
    If hojaOrigen.PageSetup.PrintArea = "" Then
        MsgBox "No hay un área de impresión definida en la hoja activa.", vbExclamation
        Exit Sub
    End If

    ' Solicitar nombre del proyecto
    nombreProyecto = InputBox("Ingrese el nombre del proyecto:", "Nombre del proyecto")
    If nombreProyecto = "" Then
        MsgBox "Operación cancelada.", vbExclamation
        Exit Sub
    End If

    ' Obtener el rango del área de impresión
    areaImpresion = hojaOrigen.PageSetup.PrintArea
    Set rangoImpresion = hojaOrigen.Range(areaImpresion)

    ' Crear nuevo libro y pegar el contenido
    Set nuevoLibro = Workbooks.Add
    Set hojaDestino = nuevoLibro.Sheets(1)
    hojaDestino.Name = nombreProyecto

    rangoImpresion.Copy
    hojaDestino.Range("A1").PasteSpecial xlPasteAll
    Application.CutCopyMode = False

    ' ---- Pegar solo valores en el rango BA154:IN170 ----
    Set rangoValores = hojaOrigen.Range("BA154:IN170")
    
    ' Calcular posición relativa dentro del área de impresión
    filaOffset = rangoValores.Row - rangoImpresion.Row
    colOffset = rangoValores.Column - rangoImpresion.Column
    
    ' Definir la celda inicial en el nuevo libro
    Set celdaInicioDestino = hojaDestino.Cells(1 + filaOffset, 1 + colOffset)
    
    ' Pegar solo valores (sin usar portapapeles)
    hojaDestino.Range(celdaInicioDestino, _
        celdaInicioDestino.Offset(rangoValores.Rows.Count - 1, rangoValores.Columns.Count - 1)).Value = rangoValores.Value
    ' ----------------------------------------------------

   

    ' Aplicar configuración de tamaño de celdas y página a todas las hojas
    For Each ws In nuevoLibro.Sheets
        With ws
            .Cells.RowHeight = 3.6
            .Cells.ColumnWidth = 0.38
            With .PageSetup
                .Orientation = xlLandscape
                .LeftMargin = Application.CentimetersToPoints(0)
                .RightMargin = Application.CentimetersToPoints(0)
                .TopMargin = Application.CentimetersToPoints(0)
                .BottomMargin = Application.CentimetersToPoints(0)
                .CenterHorizontally = True
                .CenterVertically = True
                .Zoom = 100
            End With
        End With
    Next ws

    ' Guardar archivo
    With Application.FileDialog(msoFileDialogSaveAs)
        .Title = "Guardar nuevo archivo de proyecto"
        .InitialFileName = nombreProyecto & ".xlsx"
        .FilterIndex = 1
        If .Show = -1 Then
            rutaGuardar = .SelectedItems(1)
            nuevoLibro.SaveAs rutaGuardar, FileFormat:=xlOpenXMLWorkbook
            Set LibroProyecto = nuevoLibro
            RutaLibroProyecto = rutaGuardar
            MsgBox "El archivo se ha guardado correctamente en:" & vbCrLf & rutaGuardar
        Else
            MsgBox "No se guardó el archivo.", vbInformation
            nuevoLibro.Close False
            Exit Sub
        End If
    End With
     ' Establecer el área de impresión en el nuevo libro
    hojaDestino.PageSetup.PrintArea = hojaDestino.Range("A1").Resize(rangoImpresion.Rows.Count, rangoImpresion.Columns.Count).Address
End Sub
'=============== SECCION PARA AGREGAR HOJA A NUEVO LIBRO CREADO ========================
' Agrega una hoja consecutiva al libro de proyecto creado
Sub AgregarHojaConsecutivaAlLibroCreado()
    Dim hojaOrigen As Worksheet
    Dim nuevaHoja As Worksheet
    Dim rangoImpresion As Range
    Dim areaImpresion As String
    Dim i As Integer
    Dim nombreHoja As String
    Dim existeHoja As Boolean 
    Dim ws As Worksheet
    Dim rangoValores As Range
    Dim filaOffset As Long, colOffset As Long
    Dim celdaInicioDestino As Range
    
    ' Verificar si el libro existe
    If LibroProyecto Is Nothing Then
        If RutaLibroProyecto <> "" Then
            On Error Resume Next
            Set LibroProyecto = Workbooks.Open(RutaLibroProyecto)
            On Error GoTo 0
            If LibroProyecto Is Nothing Then
                MsgBox "No se pudo reabrir el archivo del proyecto.", vbExclamation
                Exit Sub
            End If
        Else
            MsgBox "No se encontró el libro del proyecto. Por favor, cree uno primero.", vbExclamation
            Exit Sub
        End If
    End If
    
    Set hojaOrigen = ThisWorkbook.ActiveSheet
    
    ' Validar área de impresión
    If hojaOrigen.PageSetup.PrintArea = "" Then
        MsgBox "No hay un área de impresión definida en la hoja activa.", vbExclamation
        Exit Sub
    End If
    
    ' Obtener rango de impresión
    areaImpresion = hojaOrigen.PageSetup.PrintArea
    Set rangoImpresion = hojaOrigen.Range(areaImpresion)
    
    ' Buscar próximo nombre disponible (001, 002, ...)
    i = 1
    Do
        nombreHoja = Format(i, "000")
        existeHoja = False
        For Each ws In LibroProyecto.Sheets
            If ws.Name = nombreHoja Then
                existeHoja = True
                Exit For
            End If
        Next
        If Not existeHoja Then Exit Do
        i = i + 1
    Loop
    
    ' Crear hoja y asignar nombre
    Set nuevaHoja = LibroProyecto.Sheets.Add(After:=LibroProyecto.Sheets(LibroProyecto.Sheets.Count))
    nuevaHoja.Name = nombreHoja
    
    ' Copiar todo el rango de impresión
    rangoImpresion.Copy
    nuevaHoja.Range("A1").PasteSpecial xlPasteAll
    Application.CutCopyMode = False

    ' ---- Pegar solo valores en BA154:IN170 ----
    Set rangoValores = hojaOrigen.Range("BA154:IN170")
    
    filaOffset = rangoValores.Row - rangoImpresion.Row
    colOffset = rangoValores.Column - rangoImpresion.Column
    
    Set celdaInicioDestino = nuevaHoja.Cells(1 + filaOffset, 1 + colOffset)
    
    nuevaHoja.Range(celdaInicioDestino, _
        celdaInicioDestino.Offset(rangoValores.Rows.Count - 1, rangoValores.Columns.Count - 1)).Value = rangoValores.Value
    ' ------------------------------------------
    
    ' Ajustar área de impresión en nueva hoja
    nuevaHoja.PageSetup.PrintArea = nuevaHoja.Range("A1").Resize(rangoImpresion.Rows.Count, rangoImpresion.Columns.Count).Address
    
    ' Configuración de página
    With nuevaHoja
        .Cells.RowHeight = 3.6
        .Cells.ColumnWidth = 0.38
        With .PageSetup
            .Orientation = xlLandscape
            .LeftMargin = Application.CentimetersToPoints(0)
            .RightMargin = Application.CentimetersToPoints(0)
            .TopMargin = Application.CentimetersToPoints(0)
            .BottomMargin = Application.CentimetersToPoints(0)
            .CenterHorizontally = True
            .CenterVertically = True
            .Zoom = 100
        End With
    End With
    
    MsgBox "Hoja '" & nombreHoja & "' agregada correctamente al libro del proyecto.", vbInformation
End Sub
'=============== SECCION PARA IMPORTAR HOJA EN EL PROGRAMA ========================
' Importa una hoja desde otro archivo de Excel y la pega en la hoja activa
SUB IMPORTARHOJAS()
    Dim rutaArchivo As Variant
    Dim hojaDestino As Worksheet
    Dim nombreHoja As String
    Dim listaHojas As String
    Dim hoja As Worksheet

    ' Seleccionar archivo
    rutaArchivo = Application.GetOpenFilename("Archivos de Excel (*.xls*), *.xls*", , "Selecciona el archivo que deseas importar")
    If rutaArchivo = False Then Exit Sub

    ' Abrir archivo y mostrar hojas
    Set LibroImportado = Workbooks.Open(rutaArchivo)
    listaHojas = ""
    For Each hoja In LibroImportado.Sheets
        listaHojas = listaHojas & hoja.Name & vbCrLf
    Next hoja

    nombreHoja = InputBox("Selecciona una hoja:" & vbCrLf & listaHojas, "Seleccionar hoja")
    If nombreHoja = "" Then
        LibroImportado.Close False
        Exit Sub
    End If

    On Error Resume Next
    Set HojaImportada = LibroImportado.Sheets(nombreHoja)
    On Error GoTo 0

    If HojaImportada Is Nothing Then
        MsgBox "La hoja no existe.", vbCritical
        LibroImportado.Close False
        Exit Sub
    End If

    ' Validar área de impresión
    AreaImportacion = HojaImportada.PageSetup.PrintArea
    If AreaImportacion = "" Then
        MsgBox "No hay área de impresión definida.", vbExclamation
        LibroImportado.Close False
        Exit Sub
    End If

    ' Copiar área
    HojaImportada.Range(AreaImportacion).Copy

    ' Pegar en hoja activa del archivo base
    Set hojaDestino = ThisWorkbook.ActiveSheet
    hojaDestino.Range("AX28").PasteSpecial xlPasteAll
    Application.CutCopyMode = False

    MsgBox "Área importada. Puedes modificarla y luego usar la opción de devolver cambios.", vbInformation
END SUB
'=============== SECCION PARA EXPORTAR HOJA PROGRAMA ========================
' Devuelve los cambios realizados en el área importada al archivo original
Sub DevolverCambiosAlLibroImportado()
    On Error GoTo errorHandler

    ' Validaciones
    If LibroImportado Is Nothing Then
        MsgBox "No hay un libro importado activo.", vbExclamation
        Exit Sub
    End If

    If HojaImportada Is Nothing Then
        MsgBox "No hay una hoja importada registrada.", vbExclamation
        Exit Sub
    End If

    If AreaImportacion = "" Then
        MsgBox "No se ha registrado el área de impresión.", vbExclamation
        Exit Sub
    End If

    ' Copiar el área desde archivo base (A1 en adelante)
    Dim hojaBase As Worksheet
    Dim rangoOrigen As Range
    Dim rangoDestino As Range
    Dim filas As Long, columnas As Long

    Set hojaBase = ThisWorkbook.ActiveSheet
    filas = HojaImportada.Range(AreaImportacion).Rows.Count
    columnas = HojaImportada.Range(AreaImportacion).Columns.Count

    Set rangoOrigen = hojaBase.Range("AX28").Resize(filas, columnas)
    rangoOrigen.Copy

    ' Pegar sobre el área original en el archivo importado
    Set rangoDestino = HojaImportada.Range(AreaImportacion)
    rangoDestino.PasteSpecial xlPasteAll
    Application.CutCopyMode = False

    MsgBox "Cambios devueltos con formato, fórmulas y objetos gráficos.", vbInformation
    Exit Sub

    errorHandler:
        MsgBox "Error al devolver los cambios: " & Err.Description, vbCritical
End Sub
Sub CerrarSesion() 
    On Error Resume Next
    ThisWorkbook.Names("usuario_actual").Delete
    Sheets("IM").Visible = xlSheetVeryHidden
    Sheets("QM").Visible = xlSheetVeryHidden
    Sheets("GAM").Visible = xlSheetVeryHidden 
    Sheets("DM1").Visible = xlSheetVeryHidden
    Sheets("UNIM").Visible = xlSheetVeryHidden
    Sheets("UIM 36kv").Visible = xlSheetVeryHidden
    Sheets("UNQM").Visible = xlSheetVeryHidden
    Sheets("UQM 36kv").Visible = xlSheetVeryHidden
    Sheets("UNGAM").Visible = xlSheetVeryHidden
    Sheets("UGAM 36kv").Visible = xlSheetVeryHidden
    Sheets("UNDM1").Visible = xlSheetVeryHidden
    Sheets("UDM1 36kv").Visible = xlSheetVeryHidden
    Sheets("UNGBC").Visible = xlSheetVeryHidden
    Sheets("UGBC 36kv").Visible = xlSheetVeryHidden
    Sheets("GBC 24kV 1 MED").Visible = xlSheetVeryHidden
    Sheets("GBC 24kV 2 MED").Visible = xlSheetVeryHidden
    Sheets("GBC 36kV 1 MED").Visible = xlSheetVeryHidden
    Sheets("GBC 36kV 2 MED").Visible = xlSheetVeryHidden
    Sheets("USUARIOS").Visible = xlSheetVeryHidden
    Sheets("ACCESOS").Visible = xlSheetVeryHidden
    Sheets("INICIO").Activate
    frmLogin.Show
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"", false)"
    Application.DisplayFormulaBar = False
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = False
    Application.DisplayStatusBar = False
    ActiveWindow.DisplayWorkbookTabs = False
    ActiveWindow.DisplayHorizontalScrollBar = False
    ActiveWindow.DisplayVerticalScrollBar = False 
End Sub

Sub abrirlogin1()
    Bloqueo_MTM.Show
End Sub
Sub abrirlogin2()
    Bloqueo_UNF.Show
End Sub
Sub abrirlogin3()
    Bloqueo_USUARIOS.Show
End Sub