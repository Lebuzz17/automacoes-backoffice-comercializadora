Attribute VB_Name = "Módulo4"
Function PROCV_EhVermelho(valor As Variant, tabela As Range, colRet As Long) As Boolean
    Dim c As Range
    Dim resultado As Range

    ' Procura o valor na primeira coluna da tabela
    Set c = tabela.Columns(1).Find(What:=valor, LookAt:=xlWhole, LookIn:=xlValues)
    If c Is Nothing Then
        PROCV_EhVermelho = False
    Else
        ' Vai para a célula correspondente na coluna desejada
        Set resultado = c.Offset(0, colRet - 1)
        ' Testa se a fonte é vermelha
        If resultado.Font.Color = vbRed Then
            PROCV_EhVermelho = True
        Else
            PROCV_EhVermelho = False
        End If
    End If
End Function

