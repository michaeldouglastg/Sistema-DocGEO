Attribute VB_Name = "M_DOC_Anuencia"
Option Explicit

' =========================================================================================
' FUNÇÃO PARA GERAR O TEXTO DA CARTA DE ANUÊNCIA PARA O PREVIEW
' =========================================================================================
Public Function GerarTextoAnuencia(confrontanteSelecionado As String, dadosPropriedade As Object, dadosTecnico As Object) As String
    On Error GoTo ErroFuncao
    
    ' --- ETAPA 1: Filtrar os dados ---
    Dim dadosFiltrados As Collection: Set dadosFiltrados = New Collection
    Dim wsPrincipal As Worksheet: Set wsPrincipal = ThisWorkbook.Sheets(ObterNomeAbaAtiva())
    Dim loPrincipal As ListObject: Set loPrincipal = wsPrincipal.ListObjects(ObterNomeTabelaAtiva())
    Dim i As Long
    
    For i = 1 To loPrincipal.ListRows.Count
        If loPrincipal.ListRows(i).Range(8).Value = confrontanteSelecionado Then
            dadosFiltrados.Add loPrincipal.ListRows(i).Range
        End If
    Next i
    
    If dadosFiltrados.Count = 0 Then
        GerarTextoAnuencia = "ERRO: Nenhum segmento encontrado para o confrontante selecionado."
        Exit Function
    End If
    
    ' --- ETAPA 2: Construir o Texto ---
    Dim TextoAnuencia As String
    
    ' Título
    TextoAnuencia = "DECLARAÇÃO DE RECONHECIMENTO DE LIMITES" & vbCrLf & vbCrLf
    
    ' Parágrafo de Declaração
    TextoAnuencia = TextoAnuencia & vbTab & confrontanteSelecionado & ", proprietários do imóvel rural, no município de, " & dadosPropriedade("Município/UF") & ";" & vbCrLf
    TextoAnuencia = TextoAnuencia & "Confrontante de, " & dadosPropriedade("Proprietário") & ", CPF: " & dadosPropriedade("CPF") & ", proprietária do imóvel rural denominado, " & dadosPropriedade("Denominação")
    TextoAnuencia = TextoAnuencia & ", Matrícula: " & dadosPropriedade("Matrícula") & ", na comarca e município de, " & dadosPropriedade("Comarca") & ", declaramos não existir nenhuma disputa ou discordância sobre os limites comuns existentes entre os citados imóveis." & vbCrLf & vbCrLf
    
    ' Tabela de Coordenadas (em formato de texto)
    TextoAnuencia = TextoAnuencia & "Descrição do trecho de confrontação:" & vbCrLf & vbCrLf
    TextoAnuencia = TextoAnuencia & "De" & vbTab & "Para" & vbTab & "Azimute" & vbTab & "Distância (m)" & vbTab & "E(X) Longitude" & vbTab & "N(Y) Latitude" & vbTab & "Altitude" & vbCrLf
    TextoAnuencia = TextoAnuencia & String(200, "-") & vbCrLf
    
    ' Primeira linha (ponto de partida)
    TextoAnuencia = TextoAnuencia & "" & vbTab & dadosFiltrados(1).Cells(1).Value & vbTab & "" & vbTab & "" & vbTab & dadosFiltrados(1).Cells(2).Value & vbTab & dadosFiltrados(1).Cells(3).Value & vbTab & Format(dadosFiltrados(1).Cells(4).Value, "0.00") & vbCrLf
    
    ' Loop para as linhas de dados
    Dim totalDistancia As Double: totalDistancia = 0
    For i = 1 To dadosFiltrados.Count
        Dim linhaPrincipal As Long: linhaPrincipal = dadosFiltrados(i).Row - loPrincipal.HeaderRowRange.Row
        TextoAnuencia = TextoAnuencia & dadosFiltrados(i).Cells(1).Value & vbTab ' De
        TextoAnuencia = TextoAnuencia & dadosFiltrados(i).Cells(5).Value & vbTab ' Para
        TextoAnuencia = TextoAnuencia & dadosFiltrados(i).Cells(6).Value & vbTab ' Azimute
        TextoAnuencia = TextoAnuencia & Format(dadosFiltrados(i).Cells(7).Value, "0.00") & vbTab ' Distância
        TextoAnuencia = TextoAnuencia & loPrincipal.ListRows(linhaPrincipal).Range(2).Value & vbTab ' Longitude
        TextoAnuencia = TextoAnuencia & loPrincipal.ListRows(linhaPrincipal).Range(3).Value & vbTab ' Latitude
        TextoAnuencia = TextoAnuencia & Format(loPrincipal.ListRows(linhaPrincipal).Range(4).Value, "0.00") & vbCrLf ' Altitude
        If IsNumeric(dadosFiltrados(i).Cells(7).Value) Then totalDistancia = totalDistancia + CDbl(dadosFiltrados(i).Cells(7).Value)
    Next i
    
    TextoAnuencia = TextoAnuencia & String(200, "-") & vbCrLf
    TextoAnuencia = TextoAnuencia & "Total: " & dadosFiltrados.Count & vbTab & vbTab & "Somatória: " & Format(totalDistancia, "0.00") & vbCrLf & vbCrLf
    
    ' Texto Final
    TextoAnuencia = TextoAnuencia & vbTab & "Declaramos ainda que o profissional, " & dadosTecnico("Nome do Técnico") & ", " & dadosTecnico("Formação") & ", " & dadosTecnico("TRT/ART") _
               & ", credenciado pelo INCRA sob o cód. " & dadosTecnico("Cód. Incra") & ", nos indicou as demarcações do limite entre as nossas propriedades, tanto no campo como nas suas representações gráficas." & vbCrLf _
               & vbTab & "Concordamos com essa demarcação, expressa na planta e no memorial descritivo, ambos em anexo, e reconhecemos esta descrição como o limite legal entre nossas propriedades." & vbCrLf & vbCrLf

    ' Data
    Dim dataTexto As String, dataCapitalizada As String
    dataTexto = Format(Date, "dd") & " de " & Format(Date, "mmmm") & " de " & Format(Date, "yyyy")
    dataCapitalizada = StrConv(dataTexto, vbProperCase)
    
    dataTexto = Replace(dataCapitalizada, " De ", " de ")
    TextoAnuencia = TextoAnuencia & String(8, vbTab) & dadosPropriedade("Município/UF") & ", " & dataTexto & "." & String(4, vbCrLf)
    
    ' Assinaturas
    TextoAnuencia = TextoAnuencia & "____________________________________" & vbTab & "____________________________________" & vbCrLf
    TextoAnuencia = TextoAnuencia & "Proprietário(a) do Imóvel" & vbTab & "Confrontante" & vbCrLf
    TextoAnuencia = TextoAnuencia & dadosPropriedade("Proprietário") & vbTab & confrontanteSelecionado & vbCrLf
    'TextoAnuencia = TextoAnuencia & "CPF: " & dadosPropriedade("CPF") & vbTab & "CPF: " & M_Utils.GetCadastroValue(M_Config.LBL_CONFRONTANTE_CPF) & vbCrLf & String(4, vbCrLf)
    TextoAnuencia = TextoAnuencia & "CPF: " & dadosPropriedade("CPF") & vbTab & "CPF: _______________" & vbCrLf & String(4, vbCrLf)
    
    TextoAnuencia = TextoAnuencia & String(10, vbTab) & "____________________________________" & vbCrLf
    TextoAnuencia = TextoAnuencia & String(10, vbTab) & "Responsável Técnico" & vbCrLf
    TextoAnuencia = TextoAnuencia & String(10, vbTab) & dadosTecnico("Nome do Técnico")
    
    ' Retorna o texto completo
    GerarTextoAnuencia = TextoAnuencia
    Exit Function
    
ErroFuncao:
    GerarTextoAnuencia = "Ocorreu um erro ao gerar o texto da Anuência: " & Err.Description
End Function
' =========================================================================================
' MACRO PARA GERAR A CARTA DE ANUÊNCIA (INDIVIDUAL OU EM MASSA)
' =========================================================================================
Public Sub GerarCartaAnuencia(confrontanteSelecionado As String, _
                              dadosPropriedade As Object, _
                              dadosTecnico As Object, _
                              Optional gerarComoPDF As Boolean = False, _
                              Optional pastaDestino As String = "")

    On Error GoTo ErroWord
    
    ' --- ETAPA 1: Coleta e Filtro de Dados ---
    Dim dadosFiltrados As Collection: Set dadosFiltrados = New Collection
    Dim wsPrincipal As Worksheet: Set wsPrincipal = ThisWorkbook.Sheets(M_Config.App_GetNomeAbaAtiva())
    Dim loPrincipal As ListObject: Set loPrincipal = wsPrincipal.ListObjects(M_Config.App_GetNomeTabelaAtiva())
    Dim i As Long
    
    frmAguarde.Show vbModeless
    frmAguarde.AtualizarStatus "Gerando Carta de Anuência..."
    
    For i = 1 To loPrincipal.ListRows.Count
        If loPrincipal.ListRows(i).Range(8).Value = confrontanteSelecionado Then
            dadosFiltrados.Add loPrincipal.ListRows(i).Range
        End If
    Next i
    If dadosFiltrados.Count = 0 Then
        MsgBox "Nenhum segmento encontrado para o confrontante selecionado.", vbInformation
        Exit Sub
    End If
    
    ' --- ETAPA 2: Gerar e Formatar o Documento Word ---
    If Not M_Word_Engine.Word_Setup(False, 1.27, 1.27, 1.27, 1.27) Then Exit Sub
    Dim wordApp As Object: Set wordApp = M_Word_Engine.GetWordApp()
    Dim wordDoc As Object: Set wordDoc = M_Word_Engine.GetWordDoc()

    ' Usa o objeto Selection para construir o documento
    With wordApp.Selection
        ' Título
        .ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Font.Bold = True: .Font.Underline = wdUnderlineSingle: .Font.Size = 14
        .TypeText "DECLARAÇÃO DE RECONHECIMENTO DE LIMITES"
        .TypeParagraph: .TypeParagraph
        .Font.Name = "Arial": .Font.Size = 12: .Font.Underline = wdUnderlineNone

        ' Parágrafo de Declaração
        .ParagraphFormat.Alignment = wdAlignParagraphJustify
        .Font.Bold = True: .TypeText vbTab & confrontanteSelecionado
        .Font.Bold = False: .TypeText ", proprietários do imóvel rural, no município de, "
        .Font.Bold = True: .TypeText dadosPropriedade("Município/UF")
        .Font.Bold = False: .TypeText "; Confrontante de, "
        .Font.Bold = True: .TypeText dadosPropriedade("Proprietário") & ", CPF: " & dadosPropriedade("CPF")
        .Font.Bold = False: .TypeText ", proprietária do imóvel rural denominado, "
        .Font.Bold = True: .TypeText dadosPropriedade("Denominação")
        .Font.Bold = False: .TypeText ", "
        .Font.Bold = True: .TypeText "Matrícula: " & dadosPropriedade("Matrícula")
        .Font.Bold = False: .TypeText ", na comarca e município de, "
        .Font.Bold = True: .TypeText dadosPropriedade("Comarca")
        .Font.Bold = False: .TypeText ", declaramos não existir nenhuma disputa ou discordância sobre os limites comuns existentes entre os citados imóveis."
        .TypeParagraph: .TypeParagraph

        ' Tabela de Coordenadas
        .Font.Bold = False: .TypeText "Descrição do trecho de confrontação:": .TypeParagraph: .TypeParagraph
        Dim tblWord As Word.Table, numLinhasTabela As Long
        numLinhasTabela = dadosFiltrados.Count + 3
        Set tblWord = wordDoc.Tables.Add(Range:=.Range, NumRows:=numLinhasTabela, NumColumns:=7)
        With tblWord
            .Range.Font.Name = "Arial": .Range.Font.Size = 8
            .Borders.Enable = True
            .Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
            .Range.Cells.VerticalAlignment = wdCellAlignVerticalCenter
            
            ' Cabeçalho
            With .Rows(1).Range
                .Font.Bold = True
                .Shading.BackgroundPatternColor = wdColorGray15
            End With
            .cell(1, 1).Range.Text = "De": .cell(1, 2).Range.Text = "Para": .cell(1, 3).Range.Text = "Azimute"
            .cell(1, 4).Range.Text = "Distância (m)": .cell(1, 5).Range.Text = "E(X) Longitude"
            .cell(1, 6).Range.Text = "N(Y) Latitude": .cell(1, 7).Range.Text = "Altitude"
            
            ' Corpo da Tabela
            .cell(2, 2).Range.Text = dadosFiltrados(1).Cells(1).Value: .cell(2, 5).Range.Text = dadosFiltrados(1).Cells(2).Value
            .cell(2, 6).Range.Text = dadosFiltrados(1).Cells(3).Value: .cell(2, 7).Range.Text = Format(dadosFiltrados(1).Cells(4).Value, "0.00")
            Dim totalDistancia As Double: totalDistancia = 0
            For i = 1 To dadosFiltrados.Count
                Dim linhaPrincipal As Long: linhaPrincipal = dadosFiltrados(i).Row - loPrincipal.HeaderRowRange.Row
                .cell(i + 2, 1).Range.Text = dadosFiltrados(i).Cells(1).Value: .cell(i + 2, 2).Range.Text = dadosFiltrados(i).Cells(5).Value
                .cell(i + 2, 3).Range.Text = dadosFiltrados(i).Cells(6).Value: .cell(i + 2, 4).Range.Text = Format(dadosFiltrados(i).Cells(7).Value, "0.00")
                .cell(i + 2, 5).Range.Text = loPrincipal.ListRows(linhaPrincipal).Range(2).Value
                .cell(i + 2, 6).Range.Text = loPrincipal.ListRows(linhaPrincipal).Range(3).Value
                .cell(i + 2, 7).Range.Text = Format(loPrincipal.ListRows(linhaPrincipal).Range(4).Value, "0.00")
                totalDistancia = totalDistancia + CDbl(dadosFiltrados(i).Cells(7).Value)
            Next i
            
            ' Rodapé da Tabela
            With .Rows.Last.Range
                .Font.Bold = True
                .Shading.BackgroundPatternColor = wdColorGray15
            End With
            .cell(numLinhasTabela, 1).Range.Text = "Total: " & dadosFiltrados.Count
            .cell(numLinhasTabela, 3).Range.Text = "Somatória: " & Format(totalDistancia, "0.00")
            .cell(numLinhasTabela, 1).Merge MergeTo:=.cell(numLinhasTabela, 2)
            .cell(numLinhasTabela, 2).Merge MergeTo:=.cell(numLinhasTabela, 3)
        End With
    End With

    ' Move o cursor para FORA da tabela
    Dim rng As Word.Range
    Set rng = wordDoc.Content
    rng.Collapse wdCollapseEnd
    rng.Select
    
    ' Continua a construção com o objeto Selection
    With wordApp.Selection
        .TypeParagraph ' Adiciona espaçamento
        
        ' Texto Final
        .ParagraphFormat.Alignment = wdAlignParagraphJustify
        .Font.Size = 12
        .Font.Bold = False: .TypeText vbTab & "Declaramos ainda que o profissional, "
        .Font.Bold = True: .TypeText dadosTecnico("Nome do Técnico")
        .Font.Bold = False: .TypeText ", "
        .Font.Bold = True: .TypeText dadosTecnico("Formação")
        .Font.Bold = False: .TypeText ", "
        .Font.Bold = True: .TypeText dadosTecnico("TRT/ART")
        .Font.Bold = False: .TypeText ", credenciado pelo INCRA sob o cód. "
        .Font.Bold = True: .TypeText dadosTecnico("Cód. Incra")
        .Font.Bold = False: .TypeText ", nos indicou as demarcações do limite entre as nossas propriedades, tanto no campo como nas suas representações gráficas." & vbCrLf _
                                  & vbTab & "Concordamos com essa demarcação, expressa na planta e no memorial descritivo, ambos em anexo, e reconhecemos esta descrição como o limite legal entre nossas propriedades."
        .TypeParagraph
        .TypeParagraph

        ' Data
        Dim dataTexto As String, dataCapitalizada As String
        dataTexto = Format(Date, "dd") & " de " & Format(Date, "mmmm") & " de " & Format(Date, "yyyy")
        dataCapitalizada = StrConv(dataTexto, vbProperCase)
        dataTexto = Replace(dataCapitalizada, " De ", " de ")
        
        .ParagraphFormat.Alignment = wdAlignParagraphRight
        .Font.Bold = True: .TypeText dadosPropriedade("Município/UF") & ", " & dataTexto & "."
        .TypeParagraph: .TypeParagraph: .TypeParagraph: .TypeParagraph
    End With

    ' Bloco de Assinaturas
    Dim tblAssinaturas As Word.Table
    Set tblAssinaturas = wordDoc.Tables.Add(Range:=wordApp.Selection.Range, NumRows:=2, NumColumns:=2)
    tblAssinaturas.Borders.Enable = False
    With tblAssinaturas
        .Range.Font.Size = 11
        .cell(1, 1).Range.Text = "____________________________________" & vbCrLf & "Proprietário(a) do Imóvel" & vbCrLf & dadosPropriedade("Proprietário") & vbCrLf & "CPF: " & dadosPropriedade("CPF")
        .cell(1, 1).Range.Paragraphs(1).Range.Font.Bold = False
        .cell(1, 1).Range.Paragraphs(2).Range.Font.Bold = True
        .cell(1, 1).Range.Paragraphs(3).Range.Font.Bold = False
        .cell(1, 1).Range.Paragraphs(4).Range.Font.Bold = False
        .cell(1, 1).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
        
        '.cell(1, 2).Range.Text = "____________________________________" & vbCrLf & "Confrontante" & vbCrLf & confrontanteSelecionado & vbCrLf & "CPF: _______________"
        .cell(1, 2).Range.Text = "____________________________________" & vbCrLf & "Confrontante" & vbCrLf & confrontanteSelecionado & vbCrLf & "CPF: " & M_Utils.GetCadastroValue(M_Config.LBL_CONFRONTANTE_CPF)
        .cell(1, 2).Range.Paragraphs(1).Range.Font.Bold = False
        .cell(1, 2).Range.Paragraphs(2).Range.Font.Bold = True
        .cell(1, 2).Range.Paragraphs(3).Range.Font.Bold = False
        .cell(1, 2).Range.Paragraphs(4).Range.Font.Bold = False
        .cell(1, 2).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
        
        .cell(2, 2).Range.Text = "____________________________________" & vbCrLf & "Responsável Técnico" & vbCrLf & dadosTecnico("Nome do Técnico") & vbCrLf & dadosTecnico("Formação") & vbCrLf & dadosTecnico("Registro (CFT/CREA)") & " / INCRA: " & dadosTecnico("Cód. Incra") & vbCrLf & dadosTecnico("TRT/ART")
        .cell(2, 2).Range.Paragraphs(1).Range.Font.Bold = False
        .cell(2, 2).Range.Paragraphs(2).Range.Font.Bold = True
        .cell(2, 2).Range.Paragraphs(3).Range.Font.Bold = False
        .cell(2, 2).Range.Paragraphs(4).Range.Font.Bold = False
        .cell(2, 2).Range.Paragraphs(5).Range.Font.Bold = False
        .cell(2, 2).Range.Paragraphs(6).Range.Font.Bold = False
        .cell(2, 2).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    End With
        
    ' --- ETAPA 3: FINALIZAÇÃO INTELIGENTE ---
    Dim nomeArquivo As String
    nomeArquivo = "Carta de Anuência - " & M_Utils.File_SanitizeName(dadosPropriedade("Denominação")) & _
                  " - " & M_Utils.File_SanitizeName(confrontanteSelecionado)
    
    Dim caminho As String
    If pastaDestino <> "" Then
        caminho = M_Word_Engine.Word_Teardown(nomeArquivo, gerarComoPDF, False, pastaDestino, False)
    Else
        caminho = M_Word_Engine.Word_Teardown(nomeArquivo, gerarComoPDF, True)
        If caminho <> "" Then MsgBox "Carta de Anuência gerada com SUCESSO!", vbInformation
    End If
    Unload frmAguarde
    Exit Sub
    
ErroWord:
    On Error Resume Next
    Unload frmAguarde
    On Error GoTo 0
    MsgBox "Erro ao gerar para " & confrontanteSelecionado & ": " & Err.Description, vbCritical
    If Not wordApp Is Nothing Then wordApp.Quit SaveChanges:=False
    Set wordApp = Nothing
End Sub

' =========================================================================================
' MACRO PARA GERAR A CARTA DE ANUÊNCIA EM PDF
' =========================================================================================
Public Sub GerarAnuenciaPDF(confrontanteSelecionado As String, dadosProp As Object, dadosTec As Object, Optional pastaDestino As String = "")
    Call GerarCartaAnuencia(confrontanteSelecionado, dadosProp, dadosTec, True, pastaDestino)
End Sub
