Attribute VB_Name = "M_DOC_Memorial"
' =========================================================================================
'   Geração de Documentos
'   Reúne todas as macros que criam arquivos externos (Word, PDF) e os textos de pré-visualização.
' =========================================================================================
' =========================================================================================
' FUNÇÃO PARA GERAR O TEXTO DO MEMORIAL PARA O PREVIEW (VERSÃO COMPLETA)
' =========================================================================================
Public Function GerarTextoMemorial(dadosPropriedade As Object, dadosTecnico As Object) As String
    
    On Error GoTo ErroFuncao
    Dim wsPrincipal As Worksheet: Set wsPrincipal = ThisWorkbook.Sheets(ObterNomeAbaAtiva())
    Dim loPrincipal As ListObject: Set loPrincipal = wsPrincipal.ListObjects(ObterNomeTabelaAtiva())
    Dim sistemaAtivo As String: sistemaAtivo = M_Config.App_GetSistemaAtivo()
    Dim TextoMemorial As String, i As Long, perimetroTotal As Double
    
    perimetroTotal = Application.WorksheetFunction.Sum(loPrincipal.ListColumns("Distância").DataBodyRange)
    
    ' Cabeçalho
    TextoMemorial = "MEMORIAL DESCRITIVO" & vbCrLf & vbCrLf
    TextoMemorial = TextoMemorial & "Propriedade:" & vbTab & dadosPropriedade("Denominação") & vbTab & vbTab & "Matrícula: " & dadosPropriedade("Matrícula") & vbCrLf
    TextoMemorial = TextoMemorial & "Proprietário:" & vbTab & dadosPropriedade("Proprietário") & vbTab & "Código Incra: " & dadosPropriedade("Cód. Incra/SNCR") & vbCrLf
    TextoMemorial = TextoMemorial & "Município:" & vbTab & dadosPropriedade("Município/UF") & vbTab & "Comarca: " & dadosPropriedade("Comarca") & vbCrLf
    TextoMemorial = TextoMemorial & "Área SGL (ha):" & vbTab & Format(dadosPropriedade("Area (SGL)"), "#,##0.0000") & vbTab & "Perímetro (m): " & Format(perimetroTotal, "0.00") & vbCrLf & vbCrLf

    ' Corpo do Memorial
'    TextoMemorial = TextoMemorial & vbTab & "Inicia-se a descrição deste perímetro no vértice " & loPrincipal.ListRows(1).Range(1).Value _
'                    & ", de coordenadas (Longitude: " & loPrincipal.ListRows(1).Range(2).Value _
'                    & ", Latitude: " & loPrincipal.ListRows(1).Range(3).Value _
'                    & " e Altitude: " & loPrincipal.ListRows(1).Range(4).Value & " m); "
                    
                    
    TextoMemorial = TextoMemorial & vbTab & "Inicia-se a descrição deste perímetro no vértice " & loPrincipal.ListRows(1).Range(1).Value & ", de coordenadas ("
    If sistemaAtivo = "SGL" Then
        TextoMemorial = TextoMemorial & "Longitude: " & loPrincipal.ListRows(1).Range(2).Value _
                        & ", Latitude: " & loPrincipal.ListRows(1).Range(3).Value
    Else
        TextoMemorial = TextoMemorial & "Coord. N(Y): " & loPrincipal.ListRows(1).Range(2).Value _
                        & ", Coord. E(X): " & loPrincipal.ListRows(1).Range(3).Value
    End If
    TextoMemorial = TextoMemorial & " e Altitude: " & loPrincipal.ListRows(1).Range(4).Value & " m); "
                    
    Dim confrontanteAnterior As String: confrontanteAnterior = ""
    For i = 1 To loPrincipal.ListRows.Count
        Dim confrontanteAtual As String, tipoDivisa As String
        confrontanteAtual = loPrincipal.ListRows(i).Range(8).Value
        tipoDivisa = loPrincipal.ListRows(i).Range(10).Value
        If confrontanteAtual <> confrontanteAnterior Then
            If i > 1 Then TextoMemorial = Left(TextoMemorial, Len(TextoMemorial) - 2) & ". "
            TextoMemorial = TextoMemorial & tipoDivisa & "; deste, segue confrontando com " & confrontanteAtual & ", com os seguintes azimutes e distâncias: "
        End If
        Dim azimute As String, distancia As String, verticePara As String
        
        ' --- LINHAS MODIFICADAS ---
        ' Configuração de casas decimais (Mude para "0.000" se quiser 3 casas)
        Dim formatoDecimais As String: formatoDecimais = "0.00"
        
        azimute = loPrincipal.ListRows(i).Range(6).Value
        distancia = Format(loPrincipal.ListRows(i).Range(7).Value, formatoDecimais)
        
'        Azimute = loPrincipal.ListRows(i).Range(6).value
'        distancia = loPrincipal.ListRows(i).Range(7).value
        verticePara = loPrincipal.ListRows(i).Range(5).Value
        If i = loPrincipal.ListRows.Count Then
             TextoMemorial = TextoMemorial & azimute & " e " & distancia & " m até o vértice " & verticePara & ", ponto inicial da descrição deste perímetro."
        Else
'             TextoMemorial = TextoMemorial & Azimute & " e " & distancia & " m até o vértice " & verticePara & ", (Longitude: " _
'                        & loPrincipal.ListRows(i + 1).Range(2).Value & ", Latitude: " _
'                        & loPrincipal.ListRows(i + 1).Range(3).Value & " e Altitude: " _
'                        & loPrincipal.ListRows(i + 1).Range(4).Value & " m); "
                        
             TextoMemorial = TextoMemorial & azimute & " e " & distancia & " m até o vértice " & verticePara & ", ("
             If sistemaAtivo = "SGL" Then
                 TextoMemorial = TextoMemorial & "Longitude: " & loPrincipal.ListRows(i + 1).Range(2).Value _
                            & ", Latitude: " & loPrincipal.ListRows(i + 1).Range(3).Value
             Else
                 TextoMemorial = TextoMemorial & "Coord. N(Y): " & loPrincipal.ListRows(i + 1).Range(2).Value _
                            & ", Coord. E(X): " & loPrincipal.ListRows(i + 1).Range(3).Value
             End If
             TextoMemorial = TextoMemorial & " e Altitude: " & loPrincipal.ListRows(i + 1).Range(4).Value & " m); "
                            
        End If
        confrontanteAnterior = confrontanteAtual
    Next i
    
    ' Rodapé
    TextoMemorial = TextoMemorial & vbCrLf & vbCrLf & vbTab & "Todas as coordenadas aqui descritas estão georreferenciadas ao Sistema Geodésico Brasileiro tendo como datum o SIRGAS2000. A área foi obtida pelas coordenadas cartesianas locais, referenciada ao Sistema Geodésico Local (SGL-SIGEF). Todos os azimutes foram calculados pela fórmula do Problema Geodésico Inverso (Puissant). Perímetro e Distâncias foram calculados pelas coordenadas cartesianas geocêntricas." & vbCrLf
    TextoMemorial = TextoMemorial & String(4, vbCrLf) & vbTab & "Observações:" & vbCrLf
    TextoMemorial = TextoMemorial & vbTab & "A planta anexa é parte integrante deste memorial descritivo." & vbCrLf & vbCrLf
    
    ' Data
'    Dim dataTexto As String
'    dataTexto = StrConv(Format(Date, "dd 'de' mmmm 'de' yyyy"), vbProperCase)
'    dataTexto = Replace(dataTexto, " De ", " de ")
    
    Dim dataTexto As String, dataCapitalizada As String
    dataTexto = Format(Date, "dd") & " de " & Format(Date, "mmmm") & " de " & Format(Date, "yyyy")
    dataCapitalizada = StrConv(dataTexto, vbProperCase)
    dataTexto = Replace(dataCapitalizada, " De ", " de ")
    
    TextoMemorial = TextoMemorial & vbTab & vbTab & vbTab & dadosPropriedade("Município/UF") & ", " & dataTexto & "." & vbCrLf & vbCrLf & vbCrLf & vbCrLf
    
    ' Assinaturas
    TextoMemorial = TextoMemorial & "____________________________________" & vbCrLf & "Proprietário(a) do Imóvel" & vbCrLf & dadosPropriedade("Proprietário") & vbCrLf & "CPF: " & dadosPropriedade("CPF") & vbCrLf & vbCrLf & vbCrLf
    TextoMemorial = TextoMemorial & "____________________________________" & vbCrLf & "Responsável Técnico" & vbCrLf & dadosTecnico("Nome do Técnico") & vbCrLf & dadosTecnico("Formação") & vbCrLf & dadosTecnico("Registro (CFT/CREA)") & " / INCRA: " & dadosTecnico("Cód. Incra") & vbCrLf & dadosTecnico("TRT/ART")
    
    GerarTextoMemorial = TextoMemorial
    Exit Function
    
ErroFuncao:
    GerarTextoMemorial = "Ocorreu um erro ao gerar o texto do memorial: " & Err.Description
End Function

' =========================================================================================
' MACRO PARA GERAR O MEMORIAL DESCRITIVO EM WORD
' =========================================================================================
Public Sub GerarMemorialWord(dadosPropriedade As Object, dadosTecnico As Object, Optional gerarComoPDF As Boolean = False)

    On Error GoTo ErroWord
    
    ' --- ETAPA 1: Coleta e Filtro de Dados ---
    Dim wsPrincipal As Worksheet: Set wsPrincipal = ThisWorkbook.Sheets(M_Config.App_GetNomeAbaAtiva())
    Dim loPrincipal As ListObject: Set loPrincipal = wsPrincipal.ListObjects(M_Config.App_GetNomeTabelaAtiva())
    Dim sistemaAtivo As String: sistemaAtivo = M_Config.App_GetSistemaAtivo()
    Dim i As Long
    
    frmAguarde.Show vbModeless
    frmAguarde.AtualizarStatus "Gerando Memorial Descritivo..."
    
    ' --- ETAPA 2: Gerar e Formatar o Documento Word ---
    If Not M_Word_Engine.Word_Setup(False, 1.27, 1.27, 1.27, 1.27) Then Exit Sub
    Dim wordApp As Object: Set wordApp = M_Word_Engine.GetWordApp()
    Dim wordDoc As Object: Set wordDoc = M_Word_Engine.GetWordDoc()

    ' Usa o objeto Selection para construir o documento
    With wordApp.Selection
        ' Título
        .ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Font.Bold = True: .Font.Underline = wdUnderlineSingle: .Font.Size = 16
        .TypeText "MEMORIAL DESCRITIVO"
        .TypeParagraph: .TypeParagraph
        
        .Font.Name = "Arial": .Font.Bold = False: .Font.Size = 12: .Font.Underline = wdUnderlineNone
        
        ' --- CABEÇALHO EM DUAS COLUNAS COM TABELA INVISÍVEL ---
        Dim tblHeader As Word.Table
        Dim perimetroTotal As Double
        perimetroTotal = Application.WorksheetFunction.Sum(loPrincipal.ListColumns("Distância").DataBodyRange)

        Set tblHeader = wordDoc.Tables.Add(Range:=.Range, NumRows:=4, NumColumns:=2)
        tblHeader.Borders.Enable = False

        ' Preenche a tabela usando os dados recebidos do formulário
        With tblHeader
            SetCellTextBoldLabel .cell(1, 1), "Propriedade: ", dadosPropriedade("Denominação")
            SetCellTextBoldLabel .cell(1, 2), "Matrícula: ", dadosPropriedade("Matrícula")
            SetCellTextBoldLabel .cell(2, 1), "Proprietário: ", dadosPropriedade("Proprietário")
            SetCellTextBoldLabel .cell(2, 2), "Código Incra: ", dadosPropriedade("Cód. Incra/SNCR")
            SetCellTextBoldLabel .cell(3, 1), "Município: ", dadosPropriedade("Município/UF")
            SetCellTextBoldLabel .cell(3, 2), "Comarca: ", dadosPropriedade("Comarca")
            SetCellTextBoldLabel .cell(4, 1), "Área SGL (ha): ", Format(dadosPropriedade("Area (SGL)"), "#,##0.0000")
            SetCellTextBoldLabel .cell(4, 2), "Perímetro (m): ", Format(perimetroTotal, "0.00")
            .Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
        End With

        ' Move o cursor para FORA da tabela
        Dim rng As Word.Range
        Set rng = wordDoc.Content
        rng.Collapse wdCollapseEnd
        rng.Select

        .TypeParagraph

        ' Corpo do Memorial (construído pedaço por pedaço)
        .ParagraphFormat.Alignment = wdAlignParagraphJustify
        .Font.Bold = False: .TypeText vbTab & "Inicia-se a descrição deste perímetro no vértice "
        .Font.Bold = True: .TypeText loPrincipal.ListRows(1).Range(1).Value _
'        .Font.Bold = False: .TypeText ", de coordenadas (Longitude: " & loPrincipal.ListRows(1).Range(2).Value _
'            & ", Latitude: " & loPrincipal.ListRows(1).Range(3).Value _
'            & " e Altitude: " & loPrincipal.ListRows(1).Range(4).Value & " m); "

        
            
        .Font.Bold = False: .TypeText ", de coordenadas ("
        If sistemaAtivo = "SGL" Then
            .TypeText "Longitude: " & loPrincipal.ListRows(1).Range(2).Value _
                & ", Latitude: " & loPrincipal.ListRows(1).Range(3).Value
        Else
            .TypeText "Coord. N(Y): " & loPrincipal.ListRows(1).Range(2).Value _
                & ", Coord. E(X): " & loPrincipal.ListRows(1).Range(3).Value
        End If
        .TypeText " e Altitude: " & loPrincipal.ListRows(1).Range(4).Value & " m); "
        
        

        Dim confrontanteAnterior As String: confrontanteAnterior = ""
        For i = 1 To loPrincipal.ListRows.Count
            Dim confrontanteAtual As String, tipoDivisa As String
            confrontanteAtual = loPrincipal.ListRows(i).Range(8).Value
            tipoDivisa = loPrincipal.ListRows(i).Range(10).Value

            If confrontanteAtual <> confrontanteAnterior Then
                If i > 1 Then .TypeText ". "
                .TypeText tipoDivisa & "; deste, segue confrontando com "
                .Font.Bold = True: .TypeText confrontanteAtual
                .Font.Bold = False: .TypeText ", com os seguintes azimutes e distâncias: "
            End If

            Dim azimute As String, distancia As String, verticePara As String
            ' --- LINHAS MODIFICADAS ---
            ' Configuração de casas decimais (Mude para "0.000" se quiser 3 casas)
            Dim formatoDecimais As String: formatoDecimais = "0.00"
            
            azimute = loPrincipal.ListRows(i).Range(6).Value
            distancia = Format(loPrincipal.ListRows(i).Range(7).Value, formatoDecimais)
            
'            Azimute = loPrincipal.ListRows(i).Range(6).value
'            distancia = loPrincipal.ListRows(i).Range(7).value
            verticePara = loPrincipal.ListRows(i).Range(5).Value

            If i = loPrincipal.ListRows.Count Then
                .TypeText azimute & " e " & distancia & " m até o vértice "
                .Font.Bold = True: .TypeText verticePara
                .Font.Bold = False: .TypeText ", ponto inicial da descrição deste perímetro."
            Else
                .TypeText azimute & " e " & distancia & " m até o vértice "
                .Font.Bold = True: .TypeText verticePara
'                .Font.Bold = False: .TypeText ", (Longitude: " & loPrincipal.ListRows(i + 1).Range(2).Value _
'                    & ", Latitude: " & loPrincipal.ListRows(i + 1).Range(3).Value _
'                    & " e Altitude: " & loPrincipal.ListRows(i + 1).Range(4).Value & " m); "



            .Font.Bold = False: .TypeText ", ("
            If sistemaAtivo = "SGL" Then
                .TypeText "Longitude: " & loPrincipal.ListRows(i + 1).Range(2).Value _
                    & ", Latitude: " & loPrincipal.ListRows(i + 1).Range(3).Value
            Else
                .TypeText "Coord. N(Y): " & loPrincipal.ListRows(i + 1).Range(2).Value _
                    & ", Coord. E(X): " & loPrincipal.ListRows(i + 1).Range(3).Value
            End If
            .TypeText " e Altitude: " & loPrincipal.ListRows(i + 1).Range(4).Value & " m); "



            End If
            confrontanteAnterior = confrontanteAtual
        Next i

        .TypeParagraph: .TypeParagraph

        ' Texto Final
        .ParagraphFormat.Alignment = wdAlignParagraphJustify
        .Font.Size = 12
        .Font.Bold = False: .TypeText vbTab & "Todas as coordenadas aqui descritas estão georreferenciadas ao Sistema Geodésico Brasileiro tendo como datum o SIRGAS2000. A área foi obtida pelas coordenadas cartesianas locais, referenciada ao Sistema Geodésico Local (SGL-SIGEF). Todos os azimutes foram calculados pela fórmula do Problema Geodésico Inverso (Puissant). Perímetro e Distâncias foram calculados pelas coordenadas cartesianas geocêntricas."

        .TypeParagraph: .TypeParagraph: .TypeParagraph: .TypeParagraph:

        .Font.Bold = True: .Font.Size = 14: .TypeText vbTab & "Observações:"
        .TypeParagraph
        .Font.Bold = False: .Font.Size = 12: .TypeText vbTab & "A planta anexa é parte integrante deste memorial descritivo."
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
    Set tblAssinaturas = wordDoc.Tables.Add(Range:=wordApp.Selection.Range, NumRows:=2, NumColumns:=1)
    tblAssinaturas.Borders.Enable = False

    With tblAssinaturas
        .Range.Font.Size = 12
        .cell(1, 1).Range.Text = "____________________________________" & vbCrLf & "Proprietário(a) do Imóvel" & vbCrLf & dadosPropriedade("Proprietário") & vbCrLf & "CPF: " & dadosPropriedade("CPF") & vbCrLf & vbCrLf & vbCrLf
        .cell(1, 1).Range.Paragraphs(1).Range.Font.Bold = False
        .cell(1, 1).Range.Paragraphs(2).Range.Font.Bold = True
        .cell(1, 1).Range.Paragraphs(3).Range.Font.Bold = False
        .cell(1, 1).Range.Paragraphs(4).Range.Font.Bold = False
        .cell(1, 1).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter

        .cell(2, 1).Range.Text = "____________________________________" & vbCrLf & "Responsável Técnico" & vbCrLf & dadosTecnico("Nome do Técnico") & vbCrLf & dadosTecnico("Formação") & vbCrLf & dadosTecnico("Registro (CFT/CREA)") & " / INCRA: " & dadosTecnico("Cód. Incra") & vbCrLf & dadosTecnico("TRT/ART")
        .cell(2, 1).Range.Paragraphs(1).Range.Font.Bold = False
        .cell(2, 1).Range.Paragraphs(2).Range.Font.Bold = True
        .cell(2, 1).Range.Paragraphs(3).Range.Font.Bold = False
        .cell(2, 1).Range.Paragraphs(4).Range.Font.Bold = False
        .cell(2, 1).Range.Paragraphs(5).Range.Font.Bold = False
        .cell(2, 1).Range.Paragraphs(6).Range.Font.Bold = False
        .cell(2, 1).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    End With
    
    Dim nomeArquivo As String
    nomeArquivo = "Memorial - " & M_Utils.File_SanitizeName(dadosPropriedade("Denominação"))
    
    Dim caminho As String
    caminho = M_Word_Engine.Word_Teardown(nomeArquivo, gerarComoPDF)
    
    If caminho <> "" Then MsgBox "Memorial Descritivo gerado com SUCESSO!", vbInformation
    Unload frmAguarde
    Exit Sub

ErroWord:
    On Error Resume Next
    Unload frmAguarde
    On Error GoTo 0
    MsgBox "ERRO ao gerar o Memorial Descritivo: " & Err.Description, vbCritical
    If Not wordApp Is Nothing Then wordApp.Quit SaveChanges:=False
    Set wordApp = Nothing
End Sub

' =========================================================================================
' MACRO PARA GERAR O MEMORIAL DESCRITIVO EM PDF
' =========================================================================================
Public Sub GerarMemorialPDF(dadosProp As Object, dadosTec As Object)
    Call GerarMemorialWord(dadosProp, dadosTec, True)
End Sub
