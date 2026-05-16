Attribute VB_Name = "Módulo2"
Option Explicit

' === CONFIG ===
Private Const NOME_MOLDE As String = "ND 2025 - 014 - TEMPLATE"   ' aba-molde oculta no RETUSD
Private Const RESETAR_CONTADOR_NO_MOLDE As Boolean = False  ' mude para True quando quiser resetar' === MENSAGENS CUSTOMIZÁVEIS ===
Private Const MSG_PROMPT_DATA As String = "Insira a data de vencimento da ND (ex.: 18/08/2025):"
Private Const TIT_PROMPT_DATA  As String = "Definir data"
Private Const MSG_DATA_INVALIDA As String = "Digite uma data válida (ex.: 18/08/2025) ou clique Cancelar."



' === Dispare esta pelo checkbox (um por linha) ===
Sub Check_GerarND_RETUSD()
    Dim sh As Worksheet, cb As CheckBox
    Dim lin As Long
    Dim sufixo As String, valNome As Variant, valCNPJ As Variant, dtF61 As Date, valColB As Variant

    Set sh = ActiveSheet
    Set cb = sh.CheckBoxes(Application.Caller)
    If cb.Value <> 1 Then Exit Sub
    lin = cb.TopLeftCell.Row

    ' 1ª palavra da B continua como sufixo
    sufixo = FirstWord(sh.Cells(lin, "B").Text)

    ' AGORA BUSCA EM AD (nome) E AE (CNPJ)
    valNome = sh.Cells(lin, "AD").Value      ' -> vai para D22 (mescla D:I)
    valCNPJ = sh.Cells(lin, "AE").Value      ' -> vai para E26 (mescla E:H)
    valColB = sh.Cells(lin, "B").Value

    ' Data com opção de Cancelar/fechar (X)
    Dim vData As Variant
    vData = InputBoxDataOpcional(MSG_PROMPT_DATA, TIT_PROMPT_DATA)
    If IsEmpty(vData) Then
        cb.Value = xlOff          ' devolve o checkbox
        Exit Sub                  ' aborta a macro
    End If
    dtF61 = CDate(vData)


    ' Cria a nova ND a partir do molde
    CriarND_DeMolde_RETUSD sufixo, valNome, valCNPJ, dtF61, valColB

    cb.Value = xlOff
End Sub

Private Sub CriarND_DeMolde_RETUSD(ByVal sufixoB As String, _
                                   ByVal valNome As Variant, _
                                   ByVal valCNPJ As Variant, _
                                   ByVal dtF61 As Date, _
                                   ByVal repeteColB As Variant)

    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim wsMolde As Worksheet, wsNovo As Worksheet
    Dim visPrev As XlSheetVisibility
    Dim ano As String, numMolde As Long, proxNum As String
    Dim baseNome As String, sufixo As String, nomeDesejado As String, nomePdfDesejado As String
    Dim evOn As Boolean, suOn As Boolean, daOn As Boolean

    On Error Resume Next
    Set wsMolde = wb.Worksheets(NOME_MOLDE)
    On Error GoTo 0
    If wsMolde Is Nothing Then
        MsgBox "Molde não encontrado no RETUSD: '" & NOME_MOLDE & "'.", vbCritical
        Exit Sub
    End If

    If Not ParseNomeND(wsMolde.Name, ano, numMolde) Then
        MsgBox "Nome do molde deve ser 'ND <ano> - <número> - <sufixo>'.", vbCritical
        Exit Sub
    End If
    
    If RESETAR_CONTADOR_NO_MOLDE Then
        Dim wsCfg As Worksheet
        On Error Resume Next
        Set wsCfg = ThisWorkbook.Worksheets("Config")
        On Error GoTo 0
        If Not wsCfg Is Nothing Then
            wsCfg.Range("B1").Value = CLng(numMolde)  ' reinicia no número do MOLDE
        End If
    End If

    proxNum = ProximoNumeroPersistente(wb, numMolde)

    sufixo = SanitizeForSheetName(FirstWord(sufixoB))
    If Len(sufixo) = 0 Then sufixo = "ND"
    baseNome = "ND " & ano & " - " & proxNum & " - "
    If Len(baseNome & sufixo) > 31 Then sufixo = Left$(sufixo, 31 - Len(baseNome))
    nomeDesejado = baseNome & sufixo
    
    ' Nome do PDF no molde "ND <ano> - <sufixo>" (sem o número do meio)
    nomePdfDesejado = "ND " & ano & " - " & sufixo


    ' copiar molde e pegar a cópia ativa
    evOn = Application.EnableEvents: Application.EnableEvents = False
    suOn = Application.ScreenUpdating: Application.ScreenUpdating = False
    daOn = Application.DisplayAlerts:  Application.DisplayAlerts = False

    visPrev = wsMolde.Visible
    wsMolde.Visible = xlSheetVisible
    wsMolde.Copy After:=wb.Worksheets(wb.Worksheets.Count)
    wsMolde.Visible = visPrev

    Set wsNovo = ActiveSheet

    ' preencher (usando MergeArea por causa das mesclas)
    With wsNovo
        .Visible = xlSheetVisible
        .Range("E14:H14").Value = "NOTA DE DÉBITO Nº " & ano & " - " & proxNum
        .Range("E18").Value = Format(Date, "dddd, dd ""de"" mmmm ""de"" yyyy")

        ' D22 é a mescla D:I – NOME
        .Range("D22").MergeArea.Cells(1, 1).Value2 = valNome

        ' E26 é a mescla E:H – CNPJ (como texto, para preservar zeros à esquerda)
        .Range("E26").MergeArea.NumberFormat = "@"
        .Range("E26").MergeArea.Cells(1, 1).Value = CStr(valCNPJ)

        ' F61 é a mescla F:G – DATA
        .Range("F61").MergeArea.Cells(1, 1).Value2 = dtF61
        .Range("F61").MergeArea.NumberFormat = "dd/mm/yyyy"
        
        ' repete o valor da coluna B em toda a faixa
        .Range("C30:C53").Value2 = repeteColB

    End With
    
    ' ===== FILTRO: mostra apenas linhas onde a coluna I não é vazia =====
    With wsNovo
        .Range("$D$29:$I$54").AutoFilter Field:=6, Criteria1:="<>"
    End With


    ' renomear por último
    wsNovo.Name = NomeUnicoNoWorkbook(wb, nomeDesejado)
    
   ' ===== EXPORTAR A PRÓPRIA ABA CRIADA (wsNovo) COMO PDF =====
    Dim oldPA As String
    Dim pdfPath As String
    Dim nomePdf As String

' use o nome da nova aba (ou use nomePdfDesejado se você já o definiu)
    nomePdf = wsNovo.Name

' limpar PrintArea para pegar a folha inteira
    oldPA = wsNovo.PageSetup.PrintArea
    wsNovo.PageSetup.PrintArea = ""

' caminho do arquivo
    If Len(ThisWorkbook.Path) > 0 Then
        pdfPath = ThisWorkbook.Path & "\" & nomePdf & ".pdf"
    Else
        pdfPath = Environ$("USERPROFILE") & "\Desktop\" & nomePdf & ".pdf"
    End If

' exporta a PLANILHA INTEIRA criada (wsNovo), não a ActiveSheet genérica
    wsNovo.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=pdfPath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=True, _
        OpenAfterPublish:=True

' restaura PrintArea anterior
    wsNovo.PageSetup.PrintArea = oldPA

finalizar:
    Application.DisplayAlerts = daOn
    Application.ScreenUpdating = suOn
    Application.EnableEvents = evOn
End Sub



' === Helpers ===

' Lê "ND <ano> - <número> - <sufixo>" (hífen normal; tolera espaços extras e NBSP)
Private Function ParseNomeND(ByVal nm As String, ByRef ano As String, ByRef numero As Long) As Boolean
    Dim s As String, parts As Variant
    s = Replace(Trim$(nm), Chr(160), " ")
    parts = Split(s, " - ")
    If UBound(parts) <> 2 Then Exit Function
    If Left$(parts(0), 3) <> "ND " Then Exit Function
    ano = Trim$(Mid$(parts(0), 4))
    If Not IsNumeric(parts(1)) Then Exit Function
    numero = CLng(parts(1))
    ParseNomeND = True
End Function

' === Helper para escrever com segurança em células mescladas ===
Private Sub PutMerged(ByVal target As Range, ByVal v As Variant)
    If target.MergeCells Then
        target.MergeArea.Value2 = v      ' escreve em toda a área mesclada
    Else
        target.Value2 = v
    End If
End Sub


Private Function NomeUnicoNoWorkbook(ByVal wb As Workbook, ByVal baseName As String) As String
    Dim tryName As String, i As Long
    tryName = Trim$(baseName)
    i = 1
    Do While AbaExiste(wb, tryName)
        i = i + 1
        tryName = Trim$(baseName) & " (" & i & ")"
    Loop
    NomeUnicoNoWorkbook = tryName
End Function

Private Function AbaExiste(ByVal wb As Workbook, ByVal nm As String) As Boolean
    Dim sh As Worksheet
    On Error Resume Next
    Set sh = wb.Worksheets(nm)
    AbaExiste = Not (sh Is Nothing)
    On Error GoTo 0
End Function

Private Function SanitizeForSheetName(ByVal s As String) As String
    Dim bad As Variant, i As Long
    s = Trim$(s)
    bad = Array(":", "\", "/", "?", "*", "[", "]")
    For i = LBound(bad) To UBound(bad)
        s = Replace(s, bad(i), "")
    Next
    SanitizeForSheetName = s
End Function

Private Function FirstWord(ByVal s As String) As String
    Dim t As String, p As Long
    t = Replace(CStr(s), Chr(160), " ")
    t = Trim$(t)
    p = InStr(1, t, " ")
    If p > 0 Then FirstWord = Left$(t, p - 1) Else FirstWord = t
End Function

' === Caixa de data que permite Cancelar/fechar ===
Private Function InputBoxDataOpcional(ByVal prompt As String, ByVal titulo As String) As Variant
    Dim resp As Variant
    Do
        resp = Application.InputBox(prompt, titulo, Type:=2)  ' Type 2 = texto; trataremos como data
        ' Cancelar/fechar (X) retorna Boolean False
        If VarType(resp) = vbBoolean And resp = False Then
            InputBoxDataOpcional = Empty
            Exit Function
        End If
        ' Enter vazio = tratar como cancelado (se preferir obrigatória, remova este bloco)
        If Trim$(CStr(resp)) = "" Then
            InputBoxDataOpcional = Empty
            Exit Function
        End If
        If IsDate(resp) Then
            InputBoxDataOpcional = DateValue(resp)
            Exit Function
        Else
            MsgBox MSG_DATA_INVALIDA, vbExclamation
        End If
    Loop
End Function

' Retorna o próximo número da ND (formato "000") e atualiza Config!B1.
' Se a planilha Config não existir, ela é criada (muito oculta) e inicializada com defaultStart.
Private Function ProximoNumeroPersistente(ByVal wb As Workbook, ByVal defaultStart As Long) As String
    Dim wsCfg As Worksheet
    Dim v As Variant
    Dim n As Long

    ' tenta obter a planilha Config
    On Error Resume Next
    Set wsCfg = wb.Worksheets("Config")
    On Error GoTo 0

    ' se não existir, cria e inicializa
    If wsCfg Is Nothing Then
        Set wsCfg = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        On Error Resume Next
        wsCfg.Name = "Config"
        On Error GoTo 0

        ' marca como muito oculta (não aparece nem no menu Ocultar/Mostrar)
        wsCfg.Visible = xlSheetVeryHidden

        wsCfg.Range("A1").Value = "UltimoNumeroND"
        wsCfg.Range("B1").Value = CLng(defaultStart) ' ex.: 14 vindo do molde
    End If

    ' lê último número; se inválido, usa defaultStart
    v = wsCfg.Range("B1").Value
    If Not IsNumeric(v) Then
        n = CLng(defaultStart)
    Else
        n = CLng(v)
    End If

    ' incrementa, grava e retorna formatado
    n = n + 1
    wsCfg.Range("B1").Value = n

    ProximoNumeroPersistente = Format$(n, "000")
End Function


