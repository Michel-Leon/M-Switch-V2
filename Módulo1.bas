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