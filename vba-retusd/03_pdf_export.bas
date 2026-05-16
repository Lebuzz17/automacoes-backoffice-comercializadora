Attribute VB_Name = "Módulo3"
Option Explicit

' PROCV que só considera as linhas cuja FONTE tem a mesma cor de rngCorFonte.
'   lookup_value  : valor a procurar (chave)
'   table_array   : tabela (1ª coluna = chaves)
'   col_index     : coluna de retorno (>=1)
'   rngValores    : intervalo "espelho" por LINHA, onde será checada a cor da FONTE
'   rngCorFonte   : célula modelo cuja cor de FONTE será usada no filtro
'   [aprox]       : FALSE = exato (padrão) | TRUE = aproximado (tabela ordenada)
Public Function ProcvPorCorFonte(lookup_value As Variant, _
                                 table_array As Range, _
                                 col_index As Long, _
                                 rngValores As Range, _
                                 rngCorFonte As Range, _
                                 Optional aprox As Boolean = False) As Variant
    Application.Volatile True
    
    On Error GoTo falha

    If table_array Is Nothing Or rngValores Is Nothing Or rngCorFonte Is Nothing Then
        ProcvPorCorFonte = CVErr(xlErrValue): Exit Function
    End If
    If col_index < 1 Or col_index > table_array.Columns.Count Then
        ProcvPorCorFonte = CVErr(xlErrValue): Exit Function
    End If
    If rngValores.Rows.Count <> table_array.Rows.Count Then
        ' precisa alinhar 1:1 as linhas
        ProcvPorCorFonte = CVErr(xlErrRef): Exit Function
    End If

    Dim corFonte As Long
    Dim nRows As Long, r As Long
    Dim key As Variant, bestRow As Long, bestKey As Variant

    corFonte = rngCorFonte.Font.Color
    nRows = table_array.Rows.Count

    If Not aprox Then
        ' -------- PROCV EXATO --------
        For r = 1 To nRows
            If rngValores.Cells(r, 1).Font.Color = corFonte Then
                If ValuesEqual(lookup_value, table_array.Cells(r, 1).Value) Then
                    ProcvPorCorFonte = table_array.Cells(r, col_index).Value
                    Exit Function
                End If
            End If
        Next r
        ProcvPorCorFonte = CVErr(xlErrNA)
    Else
        ' -------- PROCV APROXIMADO (1ª coluna ORDENADA) --------
        bestRow = 0
        For r = 1 To nRows
            If rngValores.Cells(r, 1).Font.Color = corFonte Then
                key = table_array.Cells(r, 1).Value
                If IsNumeric(lookup_value) And IsNumeric(key) Then
                    If CDbl(key) <= CDbl(lookup_value) Then
                        If bestRow = 0 Or CDbl(key) > CDbl(bestKey) Then bestRow = r: bestKey = key
                    End If
                ElseIf IsDate(lookup_value) And IsDate(key) Then
                    If CDbl(CDate(key)) <= CDbl(CDate(lookup_value)) Then
                        If bestRow = 0 Or CDbl(CDate(key)) > CDbl(CDate(bestKey)) Then bestRow = r: bestKey = key
                    End If
                Else
                    If StrComp(CStr(key), CStr(lookup_value), vbTextCompare) <= 0 Then
                        If bestRow = 0 Or StrComp(CStr(key), CStr(bestKey), vbTextCompare) > 0 Then bestRow = r: bestKey = key
                    End If
                End If
            End If
        Next r
        If bestRow > 0 Then
            ProcvPorCorFonte = table_array.Cells(bestRow, col_index).Value
        Else
            ProcvPorCorFonte = CVErr(xlErrNA)
        End If
    End If

    Exit Function
falha:
    ProcvPorCorFonte = CVErr(xlErrNA)
End Function

' --- helper: igualdade segura para número/data/texto ---
Private Function ValuesEqual(a As Variant, b As Variant) As Boolean
    If IsNumeric(a) And IsNumeric(b) Then
        ValuesEqual = (CDbl(a) = CDbl(b))
    ElseIf IsDate(a) And IsDate(b) Then
        ValuesEqual = (CLng(CDate(a)) = CLng(CDate(b)))
    Else
        ValuesEqual = (StrComp(CStr(a), CStr(b), vbTextCompare) = 0)
    End If
End Function

