Attribute VB_Name = "M_DOC_Tabela"
Option Explicit

' =========================================================================================
' FUNÇÃO PARA GERAR O TEXTO DA TABELA ANALÍTICA PARA O PREVIEW
' =========================================================================================
Public Function GerarTextoTabelaAnalitica(dadosPropriedade As Object, dadosTecnico As Object) As String
    On Error GoTo ErroFuncao

    Dim wsPrincipal As Worksheet: Set wsPrincipal = ThisWorkbook.Sheets(M_Config.App_GetNomeAbaAtiva())
    Dim loPrincipal As ListObject: Set loPrincipal = wsPrincipal.ListObjects(M_Config.App_GetNomeTabelaAtiva())
    Dim wsConversao As Worksheet: Set wsConversao = ThisWorkbook.Sheets("TEMP_CONVERSAO")
    Dim loConversao As ListObject: Set loConversao = wsConversao.ListObjects("tbl_Conversao")
    Dim textoFinal As String, i As Long, perimetroTotal As Double

    ' Cálculo de perímetro seguro
    Dim cell As Range
    perimetroTotal = 0
    For Each cell In loPrincipal.ListColumns("Distância").DataBodyRange.Cells
        If IsNumeric(cell.Value) Then
            perimetroTotal = perimetroTotal + CDbl(cell.Value)
        End If
    Next cell
    
    ' Cálculo de área usando fórmula de Shoelace (coordenadas UTM)
    Dim areaM2 As Double, areaHa As Double, j As Long
    Dim N1 As Double, E1 As Double, N2 As Double, E2 As Double
    areaM2 = 0
    
    On Error Resume Next ' Ignora erros de conversão
    For i = 1 To loPrincipal.ListRows.Count
        j = i + 1
        If j > loPrincipal.ListRows.Count Then j = 1
        
        N1 = CDbl(loPrincipal.ListRows(i).Range(2).Value)
        E1 = CDbl(loPrincipal.ListRows(i).Range(3).Value)
        N2 = CDbl(loPrincipal.ListRows(j).Range(2).Value)
        E2 = CDbl(loPrincipal.ListRows(j).Range(3).Value)
        
        areaM2 = areaM2 + (E1 * N2 - E2 * N1)
    Next i
    On Error GoTo ErroFuncao ' Volta ao tratamento normal
    
    areaM2 = Abs(areaM2) / 2
    areaHa = areaM2 / 10000

    ' Título
    textoFinal = "TABELA ANALÍTICA" & vbCrLf & vbCrLf

    ' Cabeçalho
    textoFinal = textoFinal & "Imóvel: " & vbTab & vbTab & dadosPropriedade("Denominação") & vbCrLf
    textoFinal = textoFinal & "Proprietário: " & vbTab & dadosPropriedade("Proprietário") & vbCrLf
    textoFinal = textoFinal & "Município: " & vbTab & vbTab & dadosPropriedade("Município/UF") & vbCrLf
    textoFinal = textoFinal & "Estado: " & vbTab & vbTab & dadosPropriedade("Estado") & vbCrLf
    textoFinal = textoFinal & "Sistema UTM: " & vbTab & dadosPropriedade("Sistema UTM") & vbCrLf
    textoFinal = textoFinal & "Área medida e demarcada: " & vbTab & Format(areaHa, "#,##0.0000") & " hectares" & vbCrLf
    textoFinal = textoFinal & "Perímetro demarcado: " & vbTab & Format(perimetroTotal, "#,##0.00") & " metros" & vbCrLf & vbCrLf

    ' Descrição
    textoFinal = textoFinal & "DESCRIÇÃO" & vbCrLf
    textoFinal = textoFinal & String(150, "-") & vbCrLf
    textoFinal = textoFinal & "De" & vbTab & "Para" & vbTab & "Coord. N(Y)" & vbTab & "Coord. E(X)" & vbTab & "Azimute" & vbTab & "Distância" & vbCrLf
    textoFinal = textoFinal & String(150, "-") & vbCrLf

    ' Corpo da tabela
    If loPrincipal.ListRows.Count > 0 Then
        For i = 1 To loPrincipal.ListRows.Count
            Dim utmN As Variant, utmE As Variant, dist As Variant

            ' Lê coordenadas UTM diretamente da tabela principal
            utmN = loPrincipal.ListRows(i).Range(2).Value ' Coord. N(Y)
            utmE = loPrincipal.ListRows(i).Range(3).Value ' Coord. E(X)
            dist = loPrincipal.ListRows(i).Range(7).Value

            textoFinal = textoFinal & loPrincipal.ListRows(i).Range(1).Value & vbTab ' De
            textoFinal = textoFinal & loPrincipal.ListRows(i).Range(5).Value & vbTab ' Para

            ' Formata apenas se for numérico
            If IsNumeric(utmN) Then textoFinal = textoFinal & Format(utmN, "#,##0.00") & vbTab Else textoFinal = textoFinal & utmN & vbTab
            If IsNumeric(utmE) Then textoFinal = textoFinal & Format(utmE, "#,##0.00") & vbTab Else textoFinal = textoFinal & utmE & vbTab

            textoFinal = textoFinal & loPrincipal.ListRows(i).Range(6).Value & vbTab ' Azimute

            ' Formata apenas se for numérico
            If IsNumeric(dist) Then textoFinal = textoFinal & Format(dist, "#,##0.00 m") & vbCrLf Else textoFinal = textoFinal & dist & vbCrLf
        Next i
    End If

    textoFinal = textoFinal & String(150, "-") & vbCrLf
    
    textoFinal = textoFinal & "Perímetro: " & Format(perimetroTotal, "#,##0.00 m") & vbCrLf
    textoFinal = textoFinal & "Área m²: " & Format(areaM2, "#,##0.00 m²") & vbCrLf
    textoFinal = textoFinal & "Área ha: " & Format(areaHa, "#,##0.0000 ha") & vbCrLf & vbCrLf

    ' Data
    Dim dataTexto As String, dataCapitalizada As String
    dataTexto = Format(Date, "dd") & " de " & Format(Date, "mmmm") & " de " & Format(Date, "yyyy")
    dataCapitalizada = StrConv(dataTexto, vbProperCase)
    dataTexto = Replace(dataCapitalizada, " De ", " de ")

    textoFinal = textoFinal & vbTab & vbTab & vbTab & dadosPropriedade("Município/UF") & ", " & dataTexto & "." & vbCrLf & vbCrLf & vbCrLf

    ' Assinatura
    textoFinal = textoFinal & "____________________________________" & vbCrLf
    textoFinal = textoFinal & "Responsável Técnico" & vbCrLf
    textoFinal = textoFinal & dadosTecnico("Nome do Técnico") & vbCrLf
    textoFinal = textoFinal & dadosTecnico("Formação") & vbCrLf
    textoFinal = textoFinal & dadosTecnico("Registro (CFT/CREA)") & " / INCRA: " & dadosTecnico("Cód. Incra") & vbCrLf
    textoFinal = textoFinal & dadosTecnico("TRT/ART")

    GerarTextoTabelaAnalitica = textoFinal
    Exit Function

ErroFuncao:
    GerarTextoTabelaAnalitica = "Ocorreu um erro ao gerar o texto da Tabela Analítica: " & Err.Description
End Function

' =========================================================================================
' MACRO PARA GERAR A TABELA ANALÍTICA EM WORD
' =========================================================================================
Public Sub GerarTabelaAnaliticaWord(dadosPropriedade As Object, dadosTecnico As Object, Optional gerarComoPDF As Boolean = False)

    On Error GoTo ErroWord

    ' --- ETAPA 1: Coleta de Dados ---
    Dim wsPrincipal As Worksheet: Set wsPrincipal = ThisWorkbook.Sheets(M_Config.App_GetNomeAbaAtiva())
    Dim loPrincipal As ListObject: Set loPrincipal = wsPrincipal.ListObjects(M_Config.App_GetNomeTabelaAtiva())
    Dim wsConversao As Worksheet: Set wsConversao = ThisWorkbook.Sheets("TEMP_CONVERSAO")
    Dim loConversao As ListObject: Set loConversao = wsConversao.ListObjects("tbl_Conversao")
    Dim i As Long

    frmAguarde.Show vbModeless
    frmAguarde.AtualizarStatus "Gerando Tabela Analítica..."

    ' Cálculo de perímetro seguro
    Dim perimetroTotal As Double, cell As Range
    perimetroTotal = 0
    For Each cell In loPrincipal.ListColumns("Distância").DataBodyRange.Cells
        If IsNumeric(cell.Value) Then
            perimetroTotal = perimetroTotal + CDbl(cell.Value)
        End If
    Next cell
    
    ' Cálculo de área usando fórmula de Shoelace (coordenadas UTM)
    Dim areaM2 As Double, areaHa As Double, j As Long
    Dim N1 As Double, E1 As Double, N2 As Double, E2 As Double
    areaM2 = 0
    
    On Error Resume Next ' Ignora erros de conversão
    For i = 1 To loPrincipal.ListRows.Count
        j = i + 1
        If j > loPrincipal.ListRows.Count Then j = 1
        
        N1 = CDbl(loPrincipal.ListRows(i).Range(2).Value)
        E1 = CDbl(loPrincipal.ListRows(i).Range(3).Value)
        N2 = CDbl(loPrincipal.ListRows(j).Range(2).Value)
        E2 = CDbl(loPrincipal.ListRows(j).Range(3).Value)
        
        areaM2 = areaM2 + (E1 * N2 - E2 * N1)
    Next i
    On Error GoTo ErroWord ' Volta ao tratamento normal
    
    areaM2 = Abs(areaM2) / 2
    areaHa = areaM2 / 10000

    ' --- ETAPA 2: Gerar e Formatar o Documento Word ---
    If Not M_Word_Engine.Word_Setup(False, 2.5, 2.5, 2.25, 3#) Then Exit Sub
    Dim wordApp As Object: Set wordApp = M_Word_Engine.GetWordApp()
    Dim wordDoc As Object: Set wordDoc = M_Word_Engine.GetWordDoc()

    With wordApp.Selection
        ' Título
        .ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Font.Bold = True: .Font.Underline = wdUnderlineSingle: .Font.Size = 14
        .TypeText "TABELA ANALÍTICA"
        .TypeParagraph: .TypeParagraph

        .Font.Name = "Arial": .Font.Bold = False: .Font.Size = 12: .Font.Underline = wdUnderlineNone

        ' --- CABEÇALHO EM DUAS COLUNAS COM TABELA INVISÍVEL ---
        Dim tblHeader As Word.Table
        Set tblHeader = wordDoc.Tables.Add(Range:=.Range, NumRows:=7, NumColumns:=2)
        tblHeader.Borders.Enable = False

        With tblHeader
            ' Coluna 1: Labels (fonte normal) | Coluna 2: Valores (fonte negrito)
            .cell(1, 1).Range.Text = "Imóvel:"
            .cell(1, 2).Range.Text = dadosPropriedade("Denominação")

            .cell(2, 1).Range.Text = "Proprietário:"
            .cell(2, 2).Range.Text = dadosPropriedade("Proprietário")

            .cell(3, 1).Range.Text = "Município:"
            .cell(3, 2).Range.Text = dadosPropriedade("Município/UF")

            .cell(4, 1).Range.Text = "Estado:"
            .cell(4, 2).Range.Text = dadosPropriedade("Estado")

            .cell(5, 1).Range.Text = "Sistema UTM:"
            .cell(5, 2).Range.Text = dadosPropriedade("Sistema UTM")

            .cell(6, 1).Range.Text = "Área medida e demarcada:"
            .cell(6, 2).Range.Text = Format(areaHa, "#,##0.0000") & " hectares"

            .cell(7, 1).Range.Text = "Perímetro demarcado:"
            .cell(7, 2).Range.Text = Format(perimetroTotal, "#,##0.00") & " metros"

            ' Formatação: Coluna 1 normal, Coluna 2 negrito
            Dim r As Long
            For r = 1 To 7
                .cell(r, 1).Range.Font.Bold = False
                .cell(r, 2).Range.Font.Bold = True
            Next r

            .Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
        End With

        ' Move o cursor para FORA da tabela
        Dim rng As Word.Range
        Set rng = wordDoc.Content
        rng.Collapse wdCollapseEnd
        rng.Select

        .TypeParagraph

        ' Subtítulo "DESCRIÇÃO"
        .ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Font.Bold = True: .Font.Size = 12
        .TypeText "DESCRIÇÃO"
        .TypeParagraph

        ' Tabela de Coordenadas
        Dim tblWord As Word.Table, numLinhasTabela As Long
        numLinhasTabela = loPrincipal.ListRows.Count + 1  ' +1 cabeçalho
        Set tblWord = wordDoc.Tables.Add(Range:=.Range, NumRows:=numLinhasTabela, NumColumns:=6)

        With tblWord
            .Borders.Enable = True
            .Range.Font.Name = "Arial": .Range.Font.Size = 9
            .Range.Font.Bold = False
            .Range.ParagraphFormat.LineSpacingRule = wdLineSpaceSingle
            .Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
            .Range.Cells.VerticalAlignment = wdCellAlignVerticalCenter

            ' Cabeçalho
            With .Rows(1).Range
                .Font.Bold = True
                .Shading.BackgroundPatternColor = wdColorGray15
            End With

            .cell(1, 1).Range.Text = "De"
            .cell(1, 2).Range.Text = "Para"
            .cell(1, 3).Range.Text = "Coord. N(Y)"
            .cell(1, 4).Range.Text = "Coord. E(X)"
            .cell(1, 5).Range.Text = "Azimute"
            .cell(1, 6).Range.Text = "Distância"

            ' Corpo da tabela - Lê coordenadas diretamente da tabela principal
            If loPrincipal.ListRows.Count > 0 Then
                For i = 1 To loPrincipal.ListRows.Count
                    .cell(i + 1, 1).Range.Text = loPrincipal.ListRows(i).Range(1).Value ' De
                    .cell(i + 1, 2).Range.Text = loPrincipal.ListRows(i).Range(5).Value ' Para
                    .cell(i + 1, 3).Range.Text = Format(loPrincipal.ListRows(i).Range(2).Value, "#,##0.00") ' UTM N
                    .cell(i + 1, 4).Range.Text = Format(loPrincipal.ListRows(i).Range(3).Value, "#,##0.00") ' UTM E
                    .cell(i + 1, 5).Range.Text = loPrincipal.ListRows(i).Range(6).Value ' Azimute
                    .cell(i + 1, 6).Range.Text = Format(loPrincipal.ListRows(i).Range(7).Value, "#,##0.00 m") ' Distância
                Next i
            End If
        End With
    End With

    ' Move o cursor para FORA da tabela de coordenadas
    Set rng = wordDoc.Content
    rng.Collapse wdCollapseEnd
    rng.Select

    ' --- TABELA DE RODAPÉ (2 linhas x 1 coluna) ---
    With wordApp.Selection
        .TypeParagraph
        .TypeParagraph
        .TypeParagraph
        .TypeParagraph
        
        Dim tblRodape As Word.Table
        Set tblRodape = wordDoc.Tables.Add(Range:=.Range, NumRows:=2, NumColumns:=1)
        
        With tblRodape
            .Borders.Enable = True
            .Range.Font.Name = "Arial": .Range.Font.Size = 10
            .Range.Font.Bold = True
            .Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
            .Range.Cells.VerticalAlignment = wdCellAlignVerticalCenter
            
            .cell(1, 1).Range.Text = "Perímetro: " & Format(perimetroTotal, "#,##0.00 m")
            .cell(2, 1).Range.Text = "Área: " & Format(areaM2, "#,##0.00 m²") & "    Área: " & Format(areaHa, "#,##0.0000 ha")
        End With
    End With
    
    ' Move o cursor para FORA da tabela de rodapé
    Set rng = wordDoc.Content
    rng.Collapse wdCollapseEnd
    rng.Select
    
    With wordApp.Selection
        .TypeParagraph
        .TypeParagraph
        .TypeParagraph
        .TypeParagraph
        .TypeParagraph

        ' Data
        Dim dataTexto As String, dataCapitalizada As String
        dataTexto = Format(Date, "dd") & " de " & Format(Date, "mmmm") & " de " & Format(Date, "yyyy")
        dataCapitalizada = StrConv(dataTexto, vbProperCase)
        dataTexto = Replace(dataCapitalizada, " De ", " de ")

        .ParagraphFormat.Alignment = wdAlignParagraphRight
        .Font.Bold = True: .Font.Size = 12
        .TypeText dadosPropriedade("Município/UF") & ", " & dataTexto & "."
        .TypeParagraph: .TypeParagraph: .TypeParagraph: .TypeParagraph
    End With

    ' Move o cursor para FORA
    Set rng = wordDoc.Content
    rng.Collapse wdCollapseEnd
    rng.Select
    
    With wordApp.Selection
        .TypeParagraph
        .TypeParagraph
    End With

    ' Bloco de Assinaturas
    Dim tblAssinaturas As Word.Table
    Set tblAssinaturas = wordDoc.Tables.Add(Range:=wordApp.Selection.Range, NumRows:=1, NumColumns:=1)
    tblAssinaturas.Borders.Enable = False

    With tblAssinaturas
        .Range.Font.Size = 12
        .cell(1, 1).Range.Text = "____________________________________" & vbCrLf & _
                                 "Responsável Técnico" & vbCrLf & _
                                 dadosTecnico("Nome do Técnico") & vbCrLf & _
                                 dadosTecnico("Formação") & vbCrLf & _
                                 dadosTecnico("Registro (CFT/CREA)") & " / INCRA: " & dadosTecnico("Cód. Incra") & vbCrLf & _
                                 dadosTecnico("TRT/ART")
        .cell(1, 1).Range.Paragraphs(1).Range.Font.Bold = False
        .cell(1, 1).Range.Paragraphs(2).Range.Font.Bold = True
        .cell(1, 1).Range.Paragraphs(3).Range.Font.Bold = False
        .cell(1, 1).Range.Paragraphs(4).Range.Font.Bold = False
        .cell(1, 1).Range.Paragraphs(5).Range.Font.Bold = False
        .cell(1, 1).Range.Paragraphs(6).Range.Font.Bold = False
        .cell(1, 1).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    End With

    ' --- ETAPA 3: FINALIZAÇÃO ---
    Dim nomeArquivo As String
    nomeArquivo = "Tabela Analítica - " & M_Utils.File_SanitizeName(dadosPropriedade("Denominação"))

    Dim caminho As String
    caminho = M_Word_Engine.Word_Teardown(nomeArquivo, gerarComoPDF)

    If caminho <> "" Then MsgBox "Tabela Analítica gerada com SUCESSO!", vbInformation
    Unload frmAguarde
    Exit Sub

ErroWord:
    On Error Resume Next
    Unload frmAguarde
    On Error GoTo 0
    MsgBox "ERRO ao gerar a Tabela Analítica: " & Err.Description, vbCritical
    If Not wordApp Is Nothing Then wordApp.Quit SaveChanges:=False
    Set wordApp = Nothing
End Sub

' =========================================================================================
' MACRO PARA GERAR A TABELA ANALÍTICA EM PDF
' =========================================================================================
Public Sub GerarTabelaAnaliticaPDF(dadosProp As Object, dadosTec As Object)
    Call GerarTabelaAnaliticaWord(dadosProp, dadosTec, True)
End Sub
