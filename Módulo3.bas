Attribute VB_Name = "M�dulo3"

'=============================================
' Utilidades generales - Graficos/Shapes
' (Movido desde el modulo de la hoja ESQUEMA
'  para que sea visible desde cualquier modulo
'  del proyecto, ej. RecalcularYColocarTodo)
'=============================================

Sub EliminarGraficosEnRango(ws As Worksheet, rng As Range, Optional excluir As String = "")
    Dim shp As Shape
    Dim i As Long

    ' Recorremos de atras hacia adelante porque al eliminar shapes
    ' la coleccion se reindexa y si vas hacia adelante te saltas elementos
    For i = ws.Shapes.Count To 1 Step -1
        Set shp = ws.Shapes(i)
        If shp.Name <> excluir Then
            If ShapeIntersectaConRango(shp, rng) Then
                shp.Delete
            End If
        End If
    Next i
End Sub

Function ShapeIntersectaConRango(shp As Shape, rng As Range) As Boolean
    Dim shpLeft As Double, shpTop As Double
    Dim shpRight As Double, shpBottom As Double
    Dim rngLeft As Double, rngTop As Double
    Dim rngRight As Double, rngBottom As Double

    shpLeft = shp.Left
    shpTop = shp.Top
    shpRight = shp.Left + shp.Width
    shpBottom = shp.Top + shp.Height

    rngLeft = rng.Left
    rngTop = rng.Top
    rngRight = rng.Left + rng.Width
    rngBottom = rng.Top + rng.Height

    ' Verifica si hay solapamiento entre los dos rectangulos
    If shpLeft < rngRight And shpRight > rngLeft And _
       shpTop < rngBottom And shpBottom > rngTop Then
        ShapeIntersectaConRango = True
    Else
        ShapeIntersectaConRango = False
    End If
End Function

Public Sub limpiar_zona_unifilares()
    Dim rng As Range
    Set rng = ThisWorkbook.Worksheets("Unifilar").Range("CD44:IQ147")

    ' Limpiar el contenido y formato previo del rango (incluye bordes anteriores)
    rng.Clear

    ' Margen con línea punteada extra delgada en gris, alrededor de todo el rango
    ' (equivale a la 3ra opción de tu imagen: "línea punteada", NO "Ninguna" ni la sólida)
    With rng
        .BorderAround LineStyle:=xlDot, Weight:=xlHairline, Color:=RGB(191, 191, 191)
    End With
End Sub

