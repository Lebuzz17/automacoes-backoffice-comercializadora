Attribute VB_Name = "Módulo1"
Public Function SomaCorFonte(rngValores As Range, rngCorFonte As Range) As Double
    Dim cel As Range
    Dim soma As Double
    Dim corFonte As Long

    corFonte = rngCorFonte.Font.Color

    For Each cel In rngValores
        If cel.Font.Color = corFonte Then
            If IsNumeric(cel.Value) Then
                soma = soma + cel.Value
            End If
        End If
    Next cel

    SomaCorFonte = soma
End Function

