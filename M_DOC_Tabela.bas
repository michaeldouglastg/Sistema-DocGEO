Attribute VB_Name = "M_DOC_Tabela"
Option Explicit


'' =========================================================================================
'' MACRO PARA GERAR O DOCUMENTO "TABELA ANALÍTICA" EM WORD (VERSÃO FINAL)
'' =========================================================================================
'Public Sub GerarTabelaAnaliticaWord(dadosPropriedade As Object, dadosTecnico As Object)
'
'    On Error GoTo ErroWord
'    ' --- ETAPA 1: Garante que os dados UTM estejam atualizados ---
'    Call PreencherTabelaConversaoUTM
'
'    ' --- ETAPA 2: Configuração ---
'    Dim wsPrincipal As Worksheet: Set wsPrincipal = ThisWorkbook.Sheets(ObterNomeAbaAtiva())
'    Dim loPrincipal As ListObject: Set loPrincipal = wsPrincipal.ListObjects(ObterNomeTabelaAtiva())
'
'    Dim wsConversao As Worksheet: Set wsConversao = ThisWorkbook.Sheets("TEMP_CONVERSAO")
'    Dim loConversao As ListObject: Set loConversao = wsConversao.ListObjects("tbl_Conversao")
'
'    Dim wordApp As Word.Application, wordDoc As Word.Document
'    Dim i As Long
'
'    ' --- ETAPA 3: Gerar e Formatar o Documento Word ---
'    On Error Resume Next
'    Set wordApp = GetObject(, "Word.Application")
'    On Error GoTo ErroWord
'    If wordApp Is Nothing Then Set wordApp = New Word.Application
'
'    wordApp.Visible = False: wordApp.ScreenUpdating = False
'    Set wordDoc = wordApp.Documents.Add
'
'    ' Aplica as formatações gerais
'    With wordDoc
'        .PageSetup.TopMargin = wordApp.CentimetersToPoints(2.5)
'        .PageSetup.BottomMargin = wordApp.CentimetersToPoints(2.5)
'        .PageSetup.LeftMargin = wordApp.CentimetersToPoints(2.25)
'        .PageSetup.RightMargin = wordApp.CentimetersToPoints(3)
'        With .Content
'            .Font.Name = "Arial": .Font.Size = 12
'            .ParagraphFormat.LineSpacingRule = wdLineSpaceSingle
'        End With
'    End With
'
'    ' Usa o objeto Selection para construir o documento
'    With wordApp.Selection
'        ' Título
'        .ParagraphFormat.Alignment = wdAlignParagraphCenter
'        .Font.Bold = True: .Font.Underline = wdUnderlineSingle: .Font.Size = 14
'        .TypeText "TABELA ANALÍTICA"
'        .TypeParagraph: .TypeParagraph
'
'        .Font.Name = "Arial": .Font.Bold = False: .Font.Size = 12: .Font.Underline = wdUnderlineNone
'        '.ParagraphFormat.Alignment = wdAlignParagraphLeft
'
'         ' --- CABEÇALHO EM DUAS COLUNAS COM TABELA INVISÍVEL ---
'        Dim tblHeader As Word.Table
'        Dim perimetroTotal As Double
'        perimetroTotal = Application.WorksheetFunction.Sum(loPrincipal.ListColumns("Distância").DataBodyRange)
'
'        Set tblHeader = wordDoc.Tables.Add(Range:=.Range, NumRows:=7, NumColumns:=2)
'        tblHeader.Borders.Enable = False
'
'        ' Preenche a tabela usando os dados recebidos do formulário
'        With tblHeader
'            SetCellTextBoldLabel .cell(1, 1), "Imóvel: "
'            SetCellTextBoldLabel .cell(2, 1), "Proprietário: "
'            SetCellTextBoldLabel .cell(3, 1), "Município: "
'            SetCellTextBoldLabel .cell(4, 1), "Estado: "
'            SetCellTextBoldLabel .cell(5, 1), "Sistema UTM: "
'            SetCellTextBoldLabel .cell(6, 1), "Área Medida e Demarcada: "
'            SetCellTextBoldLabel .cell(7, 1), "Perímetro Demarcado: "
'            SetCellTextBoldLabel .cell(1, 2), dadosPropriedade("Imóvel")
'            SetCellTextBoldLabel .cell(2, 2), dadosPropriedade("Proprietário")
'            SetCellTextBoldLabel .cell(3, 2), dadosPropriedade("Município")
'            SetCellTextBoldLabel .cell(4, 2), dadosPropriedade("Estado")
'            SetCellTextBoldLabel .cell(5, 2), dadosPropriedade("Sistema UTM")
'            SetCellTextBoldLabel .cell(6, 2), dadosPropriedade("Área")
'            SetCellTextBoldLabel .cell(7, 2), dadosPropriedade("Perímetro")
'            .Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
'        End With
'
'        ' Move o cursor para FORA da tabela
'        Dim rng As Word.Range
'        Set rng = wordDoc.Content
'        rng.Collapse wdCollapseEnd
'        rng.Select
'
'        .TypeParagraph
'
'
'        ' Bloco de Informações
'        Dim perimetroTotal As Double
'        perimetroTotal = Application.WorksheetFunction.Sum(loPrincipal.ListColumns("Distância").DataBodyRange)
'
'        .Font.Bold = True: .TypeText "Proprietário: ": .Font.Bold = False: .TypeText dadosPropriedade("Proprietário") & vbCrLf
'        .Font.Bold = True: .TypeText "Imóvel Rural: ": .Font.Bold = False: .TypeText dadosPropriedade("Denominação") & vbCrLf
'        .Font.Bold = True: .TypeText "Município: ": .Font.Bold = False: .TypeText dadosPropriedade("Município/UF") & vbCrLf
'        .Font.Bold = True: .TypeText "Comarca: ": .Font.Bold = False: .TypeText dadosPropriedade("Comarca") & vbCrLf
'        .Font.Bold = True: .TypeText "Matrícula: ": .Font.Bold = False: .TypeText dadosPropriedade("Matrícula") & vbCrLf
'        .Font.Bold = True: .TypeText "Código INCRA: ": .Font.Bold = False: .TypeText dadosPropriedade("Cód. Incra/SNCR") & vbCrLf
'        .Font.Bold = True: .TypeText "Área (ha): ": .Font.Bold = False: .TypeText dadosPropriedade("Natureza/Área") & vbCrLf
'        .Font.Bold = True: .TypeText "Perímetro (m): ": .Font.Bold = False: .TypeText Format(perimetroTotal, "0.00") & vbCrLf
'        .TypeParagraph
'
'        ' Título da Tabela
'        .Font.Bold = True: .Font.Size = 12
'        .TypeText "Descrição"
'        .TypeParagraph
'
'        ' Tabela Principal
'        Dim tblWord As Word.Table, numLinhasTabela As Long
'        numLinhasTabela = loPrincipal.ListRows.Count + 1
'        Set tblWord = wordDoc.Tables.Add(Range:=.Range, NumRows:=numLinhasTabela, NumColumns:=6)
'
'        With tblWord
'            .Borders.Enable = True: .Range.Font.Name = "Arial": .Range.Font.Size = 9
'            .Range.ParagraphFormat.LineSpacingRule = wdLineSpaceSingle
'            .Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
'            .Range.Cells.VerticalAlignment = wdCellAlignVerticalCenter
'
'            With .Rows(1).Range
'                .Font.Bold = True
'                .Shading.BackgroundPatternColor = wdColorGray15
'            End With
'
'            .cell(1, 1).Range.Text = "De": .cell(1, 2).Range.Text = "Para"
'            .cell(1, 3).Range.Text = "Coord. N(Y)": .cell(1, 4).Range.Text = "Coord. E(X)"
'            .cell(1, 5).Range.Text = "Azimute": .cell(1, 6).Range.Text = "Distância"
'
'            ' Preenche o corpo da tabela
'            If loPrincipal.ListRows.Count > 0 Then
'                For i = 1 To loPrincipal.ListRows.Count
'                    .cell(i + 1, 1).Range.Text = loPrincipal.ListRows(i).Range(1).value ' De
'                    .cell(i + 1, 2).Range.Text = loPrincipal.ListRows(i).Range(5).value ' Para
'
'                    .cell(i + 1, 3).Range.Text = Format(loConversao.ListRows(i).Range(3).value, "0.000") ' UTM N
'                    .cell(i + 1, 4).Range.Text = Format(loConversao.ListRows(i).Range(2).value, "0.000") ' UTM E
'
'                    .cell(i + 1, 5).Range.Text = loPrincipal.ListRows(i).Range(6).value ' Azimute
'                    .cell(i + 1, 6).Range.Text = Format(loPrincipal.ListRows(i).Range(7).value, "0.00") ' Distância
'                Next i
'            End If
'        End With
'    End With
'
'    ' --- ETAPA 4: Finalização ---
'    wordApp.ScreenUpdating = True
'    Dim caminhoArquivo As String, nomeArquivo As String
'    nomeArquivo = "Tabela Analitica - " & dadosPropriedade("Denominação") & ".docx"
'
'    ' --- BLOCO DE SALVAMENTO PDF ATUALIZADO ---
'    Dim arquivoDestino As Variant
'
'    ' Abre a janela de diálogo para o usuário escolher onde salvar
'    arquivoDestino = Application.GetSaveAsFilename(InitialFileName:=nomeArquivo, _
'                     FileFilter:="Arquivo PDF (*.pdf), *.pdf", _
'                     Title:="Salvar PDF Como")
'
'    ' Verifica se o usuário cancelou a janela
'    If arquivoDestino = False Then
'        MsgBox "Operação cancelada pelo usuário.", vbExclamation
'        wordDoc.Close SaveChanges:=wdDoNotSaveChanges
'        wordApp.Quit
'        Set wordDoc = Nothing: Set wordApp = Nothing
'        Unload frmAguarde
'        Exit Sub
'    End If
'
'    ' Exporta para o caminho escolhido pelo usuário
'    wordDoc.ExportAsFixedFormat OutputFileName:=arquivoDestino, ExportFormat:=wdExportFormatPDF
'    ' ------------------------------------------
'
'    Unload frmAguarde
'
'    Dim resposta As VbMsgBoxResult
'    resposta = MsgBox("Tabela Analítica gerada com sucesso!" & vbCrLf & "Arquivo: " & nomeArquivo & vbCrLf & "Deseja abrir o arquivo agora?", vbQuestion + vbYesNo, "Geração Concluída")
'
'    If resposta = vbYes Then wordApp.Visible = True Else wordDoc.Close: wordApp.Quit
'
'    Set wordDoc = Nothing: Set wordApp = Nothing
'    Exit Sub
'
'ErroWord:
'    Unload frmAguarde
'
'    MsgBox "Ocorreu um erro ao gerar a Tabela Analítica: " & Err.Description, vbCritical
'    If Not wordApp Is Nothing Then wordApp.Quit SaveChanges:=False
'    Set wordApp = Nothing
'End Sub
'' =========================================================================================
'' FUNÇÃO PARA GERAR O TEXTO DA TABELA ANALÍTICA PARA O PREVIEW
'' =========================================================================================
'Public Function GerarTextoTabelaAnalitica(dadosPropriedade As Object, dadosTecnico As Object) As String
'
'    On Error GoTo ErroFuncao
'    ' Garante que os dados UTM estejam atualizados
'    Call PreencherTabelaConversaoUTM
'
'    Dim wsPrincipal As Worksheet: Set wsPrincipal = ThisWorkbook.Sheets(ObterNomeAbaAtiva())
'    Dim loPrincipal As ListObject: Set loPrincipal = wsPrincipal.ListObjects(ObterNomeTabelaAtiva())
'    Dim wsConversao As Worksheet: Set wsConversao = ThisWorkbook.Sheets("TEMP_CONVERSAO")
'    Dim loConversao As ListObject: Set loConversao = wsConversao.ListObjects("tbl_Conversao")
'    Dim textoFinal As String, i As Long, perimetroTotal As Double
'
'    perimetroTotal = Application.WorksheetFunction.Sum(loPrincipal.ListColumns("Distância").DataBodyRange)
'
'    textoFinal = "TABELA ANALÍTICA" & vbCrLf & vbCrLf
'    textoFinal = textoFinal & "Proprietário: " & vbTab & dadosPropriedade("Proprietário") & vbCrLf
'    textoFinal = textoFinal & "Imóvel Rural: " & vbTab & dadosPropriedade("Denominação") & vbCrLf
'    textoFinal = textoFinal & "Município: " & vbTab & vbTab & dadosPropriedade("Município/UF") & vbCrLf
'    textoFinal = textoFinal & "Comarca: " & vbTab & vbTab & dadosPropriedade("Comarca") & vbCrLf
'    textoFinal = textoFinal & "Matrícula: " & vbTab & vbTab & dadosPropriedade("Matrícula") & vbCrLf
'    textoFinal = textoFinal & "Código INCRA: " & vbTab & dadosPropriedade("Cód. Incra/SNCR") & vbCrLf
'    textoFinal = textoFinal & "Área (ha): " & vbTab & vbTab & dadosPropriedade("Natureza/Área") & vbCrLf
'    textoFinal = textoFinal & "Perímetro (m): " & vbTab & Format(perimetroTotal, "0.00") & vbCrLf & vbCrLf
'
'    textoFinal = textoFinal & "Descrição" & vbCrLf
'    textoFinal = textoFinal & String(150, "-") & vbCrLf
'    textoFinal = textoFinal & "De" & vbTab & "Para" & vbTab & "Coord. N(Y)" & vbTab & "Coord. E(X)" & vbTab & "Azimute" & vbTab & "Distância" & vbCrLf
'    textoFinal = textoFinal & String(150, "-") & vbCrLf
'
'    If loPrincipal.ListRows.Count > 0 Then
'        For i = 1 To loPrincipal.ListRows.Count
'            textoFinal = textoFinal & loPrincipal.ListRows(i).Range(1).value & vbTab ' De
'            textoFinal = textoFinal & loPrincipal.ListRows(i).Range(5).value & vbTab ' Para
'            textoFinal = textoFinal & Format(loConversao.ListRows(i).Range(3).value, "0.000") & vbTab ' UTM N
'            textoFinal = textoFinal & Format(loConversao.ListRows(i).Range(2).value, "0.000") & vbTab ' UTM E
'            textoFinal = textoFinal & loPrincipal.ListRows(i).Range(6).value & vbTab ' Azimute
'            textoFinal = textoFinal & Format(loPrincipal.ListRows(i).Range(7).value, "0.00") & vbCrLf ' Distância
'        Next i
'    End If
'
'    GerarTextoTabelaAnalitica = textoFinal
'    Exit Function
'
'ErroFuncao:
'    GerarTextoTabelaAnalitica = "Ocorreu um erro ao gerar o texto da Tabela Analítica: " & Err.Description
'End Function





' =========================================================================================
' FUNÇÃO PARA GERAR O TEXTO DA TABELA ANALÍTICA (VERSÃO ROBUSTA E CORRIGIDA)
' =========================================================================================
Public Function GerarTextoTabelaAnalitica(dadosPropriedade As Object, dadosTecnico As Object) As String
    On Error GoTo ErroFuncao
    
    ' Garante que os dados UTM estejam atualizados
    'Call PreencherTabelaConversaoUTM
    
    Dim wsPrincipal As Worksheet: Set wsPrincipal = ThisWorkbook.Sheets(ObterNomeAbaAtiva())
    Dim loPrincipal As ListObject: Set loPrincipal = wsPrincipal.ListObjects(ObterNomeTabelaAtiva())
    Dim wsConversao As Worksheet: Set wsConversao = ThisWorkbook.Sheets("TEMP_CONVERSAO")
    Dim loConversao As ListObject: Set loConversao = wsConversao.ListObjects("tbl_Conversao")
    Dim textoFinal As String, i As Long, perimetroTotal As Double
    
    ' --- CÁLCULO DE PERÍMETRO SEGURO ---
    ' Soma apenas as células que contêm números, ignorando erros ou texto.
    Dim cell As Range
    perimetroTotal = 0
    For Each cell In loPrincipal.ListColumns("Distância").DataBodyRange.Cells
        If IsNumeric(cell.Value) Then
            perimetroTotal = perimetroTotal + CDbl(cell.Value)
        End If
    Next cell
    
    ' --- Construção do Cabeçalho ---
    textoFinal = "TABELA ANALÍTICA" & vbCrLf & vbCrLf
    textoFinal = textoFinal & "Proprietário: " & vbTab & dadosPropriedade("Proprietário") & vbCrLf
    textoFinal = textoFinal & "Imóvel Rural: " & vbTab & dadosPropriedade("Denominação") & vbCrLf
    textoFinal = textoFinal & "Município: " & vbTab & vbTab & dadosPropriedade("Município/UF") & vbCrLf
    textoFinal = textoFinal & "Perímetro (m): " & vbTab & Format(perimetroTotal, "0.00") & vbCrLf & vbCrLf
    
    textoFinal = textoFinal & "Descrição da Tabela de Coordenadas UTM:" & vbCrLf
    textoFinal = textoFinal & String(150, "-") & vbCrLf
    textoFinal = textoFinal & "De" & vbTab & "Para" & vbTab & "Coord. N(Y)" & vbTab & "Coord. E(X)" & vbTab & "Azimute" & vbTab & "Distância (m)" & vbCrLf
    textoFinal = textoFinal & String(150, "-") & vbCrLf
    
    ' --- Construção Segura do Corpo da Tabela ---
    If loPrincipal.ListRows.Count > 0 Then
        For i = 1 To loPrincipal.ListRows.Count
            Dim utmN As Variant, utmE As Variant, dist As Variant
            
            ' Verifica se a linha correspondente existe na tabela de conversão
            If i <= loConversao.ListRows.Count Then
                utmN = loConversao.ListRows(i).Range(3).Value
                utmE = loConversao.ListRows(i).Range(2).Value
            Else
                utmN = "N/A"
                utmE = "N/A"
            End If
            
            dist = loPrincipal.ListRows(i).Range(7).Value
            
            textoFinal = textoFinal & loPrincipal.ListRows(i).Range(1).Value & vbTab ' De
            textoFinal = textoFinal & loPrincipal.ListRows(i).Range(5).Value & vbTab ' Para
            
            ' Formata apenas se for numérico
            If IsNumeric(utmN) Then textoFinal = textoFinal & Format(utmN, "0.000") & vbTab Else textoFinal = textoFinal & utmN & vbTab
            If IsNumeric(utmE) Then textoFinal = textoFinal & Format(utmE, "0.000") & vbTab Else textoFinal = textoFinal & utmE & vbTab
            
            textoFinal = textoFinal & loPrincipal.ListRows(i).Range(6).Value & vbTab ' Azimute
            
            ' Formata apenas se for numérico
            If IsNumeric(dist) Then textoFinal = textoFinal & Format(dist, "0.00") & vbCrLf Else textoFinal = textoFinal & dist & vbCrLf
        Next i
    End If
    
    GerarTextoTabelaAnalitica = textoFinal
    Exit Function
    
ErroFuncao:
    GerarTextoTabelaAnalitica = "Ocorreu um erro ao gerar o texto da Tabela Analítica: " & Err.Description
End Function

'' =========================================================================================
'' MACRO PARA GERAR O DOCUMENTO "TABELA ANALÍTICA" EM WORD
'' =========================================================================================
''Public Sub GerarTabelaAnaliticaWord(dadosPropriedade As Object, dadosTecnico As Object)
'Public Sub GerarTabelaAnaliticaWord0(dadosPropriedade As Object, dadosTecnico As Object, Optional gerarComoPDF As Boolean = False)
'
'    If Not M_Word_Engine.Word_Setup(False, 1.27, 1.27, 1.27, 1.27) Then Exit Sub
'    Dim wordApp As Object: Set wordApp = M_Word_Engine.GetWordApp()
'    Dim wordDoc As Object: Set wordDoc = M_Word_Engine.GetWordDoc()
'
'    With wordApp.Selection
'        .ParagraphFormat.Alignment = wdAlignParagraphCenter
'        .Font.Bold = True
'        .Font.Size = 12
'        .TypeText "TABELA ANALÍTICA"
'        .TypeParagraph
'        .TypeParagraph
'        .ParagraphFormat.Alignment = wdAlignParagraphLeft
'        .Font.Bold = False
'        .Font.Size = 10
'
'        .TypeText "Proprietário: " & dadosPropriedade("Proprietário") & vbCrLf
'        .TypeText "Imóvel Rural: " & dadosPropriedade("Denominação") & vbCrLf
'        ' Adicione outros campos do cabeçalho aqui...
'        .TypeParagraph
'
'        Dim tblWord As Word.Table
'        Set tblWord = wordDoc.Tables.Add(Range:=.Range, NumRows:=loPrincipal.ListRows.Count + 1, NumColumns:=6)
'        With tblWord
'            .Borders.Enable = True
'            .Rows(1).Range.Font.Bold = True
'            .cell(1, 1).Range.Text = "De"
'            .cell(1, 2).Range.Text = "Para"
'            .cell(1, 3).Range.Text = "Coord. N(Y)"
'            .cell(1, 4).Range.Text = "Coord. E(X)"
'            .cell(1, 5).Range.Text = "Azimute"
'            .cell(1, 6).Range.Text = "Distância (m)"
'
'            For i = 1 To loPrincipal.ListRows.Count
'                .cell(i + 1, 1).Range.Text = loPrincipal.ListRows(i).Range(1).Value
'                .cell(i + 1, 2).Range.Text = loPrincipal.ListRows(i).Range(5).Value
'                .cell(i + 1, 3).Range.Text = Format(loConversao.ListRows(i).Range(3).Value, "0.000")
'                .cell(i + 1, 4).Range.Text = Format(loConversao.ListRows(i).Range(2).Value, "0.000")
'                .cell(i + 1, 5).Range.Text = loPrincipal.ListRows(i).Range(6).Value
'                .cell(i + 1, 6).Range.Text = Format(loPrincipal.ListRows(i).Range(7).Value, "0.00")
'            Next i
'        End With
'    End With
'
'    ' --- ETAPA 4: FINALIZAÇÃO INTELIGENTE (WORD OU PDF) ---
'    Dim nomeArquivo As String
'    nomeArquivo = "Tabela - " & M_Utils.File_SanitizeName(dadosPropriedade("Denominação")) & _
'                  " - " & M_Utils.File_SanitizeName(confrontanteSelecionado)
'
'    Dim caminho As String
'    If pastaDestino <> "" Then
'        caminho = M_Word_Engine.Word_Teardown(nomeArquivo, gerarComoPDF, False, pastaDestino, False)
'    Else
'        caminho = M_Word_Engine.Word_Teardown(nomeArquivo, gerarComoPDF, True)
'        If caminho <> "" Then MsgBox "Documento gerado!", vbInformation
'    End If
'    Unload frmAguarde
'
'ErroWord:
'    Unload frmAguarde
'    MsgBox "Ocorreu um erro ao gerar a Tabela Analítica: " & Err.Description, vbCritical
'    If Not wordApp Is Nothing Then wordApp.Quit SaveChanges:=False
'    Set wordApp = Nothing
'End Sub

Public Sub GerarTabelaAnaliticaWord(dadosPropriedade As Object, dadosTecnico As Object, Optional gerarComoPDF As Boolean = False)
    
    On Error GoTo ErroWord
    
    ' --- ETAPA 1: CRÍTICO - ATUALIZAR DADOS UTM ---
    ' Esta linha NÃO pode estar comentada, senão a tabela loConversao fica vazia e gera o erro.
    'Call PreencherTabelaConversaoUTM descomentar depois
    
    ' --- ETAPA 2: Configuração ---
    
    ' --- ETAPA 1: Coleta e Filtro de Dados ---
    Dim wsPrincipal As Worksheet: Set wsPrincipal = ThisWorkbook.Sheets(M_Config.App_GetNomeAbaAtiva())
    Dim loPrincipal As ListObject: Set loPrincipal = wsPrincipal.ListObjects(M_Config.App_GetNomeTabelaAtiva())
    Dim wsConversao As Worksheet: Set wsConversao = ThisWorkbook.Sheets("TEMP_CONVERSAO")
    Dim loConversao As ListObject: Set loConversao = wsConversao.ListObjects("tbl_Conversao")
    Dim i As Long
    
    frmAguarde.Show vbModeless
    frmAguarde.AtualizarStatus "Gerando Tabela Analítica..."
    
    ' --- ETAPA 2: Gerar e Formatar o Documento Word ---
    If Not M_Word_Engine.Word_Setup(False, 1.27, 1.27, 1.27, 1.27) Then Exit Sub
    Dim wordApp As Object: Set wordApp = M_Word_Engine.GetWordApp()
    Dim wordDoc As Object: Set wordDoc = M_Word_Engine.GetWordDoc()
    
    With wordApp.Selection
        .ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Font.Bold = True
        .Font.Size = 12
        .TypeText "TABELA ANALÍTICA"
        .TypeParagraph
        .TypeParagraph
        .ParagraphFormat.Alignment = wdAlignParagraphLeft
        .Font.Bold = False
        .Font.Size = 10
        
        .TypeText "Proprietário: " & dadosPropriedade("Proprietário") & vbCrLf
        .TypeText "Imóvel Rural: " & dadosPropriedade("Denominação") & vbCrLf
        ' (Adicione outros campos aqui se necessário)
        .TypeParagraph
        
        Dim tblWord As Word.Table
        Set tblWord = wordDoc.Tables.Add(Range:=.Range, NumRows:=loPrincipal.ListRows.Count + 1, NumColumns:=6)
        
        With tblWord
            .Borders.Enable = True
            .Rows(1).Range.Font.Bold = True
            .cell(1, 1).Range.Text = "De"
            .cell(1, 2).Range.Text = "Para"
            .cell(1, 3).Range.Text = "Coord. N(Y)"
            .cell(1, 4).Range.Text = "Coord. E(X)"
            .cell(1, 5).Range.Text = "Azimute"
            .cell(1, 6).Range.Text = "Distância (m)"
            
            For i = 1 To loPrincipal.ListRows.Count
                ' Preenche De/Para da tabela principal
                .cell(i + 1, 1).Range.Text = loPrincipal.ListRows(i).Range(1).Value
                .cell(i + 1, 2).Range.Text = loPrincipal.ListRows(i).Range(5).Value
                
                ' --- CORREÇÃO DE SEGURANÇA AQUI ---
                ' Verifica se existe a linha correspondente na tabela de conversão
                If i <= loConversao.ListRows.Count Then
                    .cell(i + 1, 3).Range.Text = Format(loConversao.ListRows(i).Range(3).Value, "0.000")
                    .cell(i + 1, 4).Range.Text = Format(loConversao.ListRows(i).Range(2).Value, "0.000")
                Else
                    .cell(i + 1, 3).Range.Text = "N/A"
                    .cell(i + 1, 4).Range.Text = "N/A"
                End If
                ' ----------------------------------
                
                .cell(i + 1, 5).Range.Text = loPrincipal.ListRows(i).Range(6).Value
                .cell(i + 1, 6).Range.Text = Format(loPrincipal.ListRows(i).Range(7).Value, "0.00")
            Next i
        End With
    End With
    
    ' --- ETAPA 4: FINALIZAÇÃO INTELIGENTE (WORD OU PDF) ---
    Dim nomeArquivo As String
    nomeArquivo = "Tabela Analítica - " & M_Utils.File_SanitizeName(dadosPropriedade("Denominação"))
    
    Dim caminho As String
    caminho = M_Word_Engine.Word_Teardown(nomeArquivo, gerarComoPDF)
    
    If caminho <> "" Then MsgBox "Tabela Analítica gerado com SUCESSO!", vbInformation
    Unload frmAguarde
    Exit Sub

ErroWord:
    On Error Resume Next
    Unload frmAguarde
    On Error GoTo 0
    MsgBox "ERRO ao gerar o Tabela Analítica: " & Err.Description, vbCritical
    If Not wordApp Is Nothing Then wordApp.Quit SaveChanges:=False
    Set wordApp = Nothing
End Sub

' =========================================================================================
' MACRO ATALHO PARA GERAR PDF (Chama a macro principal)
' =========================================================================================
Public Sub GerarTabelaAnaliticaPDF(dadosPropriedade As Object, dadosTecnico As Object)
    ' Chama a rotina do Word passando True para forçar a geração do PDF no final
    Call GerarTabelaAnaliticaWord(dadosPropriedade, dadosTecnico, True)
End Sub
