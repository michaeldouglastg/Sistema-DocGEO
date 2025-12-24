Attribute VB_Name = "M_DOC_Laudo"
Option Explicit

' =========================================================================================
' FUNÇÃO PARA GERAR O TEXTO DO LAUDO TÉCNICO PARA O PREVIEW
' =========================================================================================
Public Function GerarTextoLaudo(dadosPropriedade As Object, dadosTecnico As Object) As String
    Dim TextoLaudo As String
    
    TextoLaudo = "LAUDO TÉCNICO" & String(2, vbCrLf)
    TextoLaudo = TextoLaudo & vbTab & "Eu, " & dadosTecnico("Nome do Técnico") & ", " & dadosTecnico("Formação") & ", CFT: " & dadosTecnico("Registro (CFT/CREA)") _
                 & ", sob efeitos do termo de responsabilidade técnica TRT: " & dadosTecnico("TRT/ART") & ", Atesto, " _
                 & "sob as penas da lei, que efetuei pessoalmente o levantamento da área e que os valores corretos dos azimutes e distâncias e a identificação das confrontações " _
                 & "são as apresentadas nesta oportunidade, na planta e no memorial descritivo que acompanham o presente laudo." & vbCrLf
                 
    TextoLaudo = TextoLaudo & vbTab & "Ao efetuar os trabalhos, constatei o seguinte: o imóvel " & dadosPropriedade("Denominação") & ", Situado Na " & dadosPropriedade("Endereço Propriedade") & ", Matricula: " & dadosPropriedade("Matrícula") _
                 & " de propriedade de " & dadosPropriedade("Proprietário") & ", CPF: " & dadosPropriedade("CPF") & ", " & dadosPropriedade("Nacionalidade") & ", " & dadosPropriedade("Profissão") & ", " _
                 & dadosPropriedade("Estado Civil") & ", maior, portador da CI nº " & dadosPropriedade("RG") & ", expedida por " & dadosPropriedade("Expedição") & " em " & dadosPropriedade("Data Expedição") & ", residente e domiciliado na " & dadosPropriedade("Endereço Proprietário") _
                 & ", possui descrição tabular precária, devido o uso de equipamentos de pouca precisão da época." & vbCrLf
                 
    TextoLaudo = TextoLaudo & vbTab & "O presente levantamento foi efetuado com aparelhos geodésicos GPS de altíssima precisão, intra-muros, uma vez que as divisas são claras (cerca de arames bem antigas) e respeitadas há muitos anos. Além disso, todos os confrontantes confirmaram que a referida cerca respeita os limites de seus imóveis." & String(4, vbCrLf)
    
    Dim dataTexto As String, dataCapitalizada As String
    dataTexto = Format(Date, "dd") & " de " & Format(Date, "mmmm") & " de " & Format(Date, "yyyy")
    dataCapitalizada = StrConv(dataTexto, vbProperCase)
    dataTexto = Replace(dataCapitalizada, " De ", " de ")
    
''    Dim dataTexto As String
''    dataTexto = StrConv(Format(Date, "dd 'de' mmmm 'de' yyyy"), vbProperCase)
''    dataTexto = Replace(dataTexto, " De ", " de ")
    TextoLaudo = TextoLaudo & String(8, vbTab) & dadosPropriedade("Município/UF") & ", " & dataTexto & "." & String(4, vbCrLf)
    
    TextoLaudo = TextoLaudo & "____________________________________" & vbCrLf
    TextoLaudo = TextoLaudo & "Proprietário do Imóvel" & vbCrLf
    TextoLaudo = TextoLaudo & dadosPropriedade("Proprietário") & vbCrLf
    TextoLaudo = TextoLaudo & "CPF: " & dadosPropriedade("CPF") & String(4, vbCrLf)
    
    TextoLaudo = TextoLaudo & "____________________________________" & vbCrLf
    TextoLaudo = TextoLaudo & "Responsável Técnico" & vbCrLf
    TextoLaudo = TextoLaudo & dadosTecnico("Nome do Técnico") & vbCrLf
    TextoLaudo = TextoLaudo & dadosTecnico("Formação") & vbCrLf
    TextoLaudo = TextoLaudo & "CFT: " & dadosPropriedade("Registro (CFT/CREA)") & vbTab & dadosTecnico("TRT/ART") & vbCrLf
    
    GerarTextoLaudo = TextoLaudo
End Function

' =========================================================================================
' MACRO PARA GERAR O DOCUMENTO "LAUDO TÉCNICO" EM WORD
' =========================================================================================
'Public Sub GerarLaudoTecnicoWord(dadosPropriedade As Object, dadosTecnico As Object)
Public Sub GerarLaudoTecnicoWord(dadosPropriedade As Object, dadosTecnico As Object, Optional gerarComoPDF As Boolean = False)
    
    On Error GoTo ErroWord
    
    ' --- ETAPA 1: Coleta e Filtro de Dados ---
    Dim wsPrincipal As Worksheet: Set wsPrincipal = ThisWorkbook.Sheets(M_Config.App_GetNomeAbaAtiva())
    Dim loPrincipal As ListObject: Set loPrincipal = wsPrincipal.ListObjects(M_Config.App_GetNomeTabelaAtiva())
    Dim i As Long
    
    frmAguarde.Show vbModeless
    frmAguarde.AtualizarStatus "Gerando Laudo Técnico..."
    
    ' --- ETAPA 2: Gerar e Formatar o Documento Word ---
    If Not M_Word_Engine.Word_Setup(False, 1.27, 1.27, 1.27, 1.27) Then Exit Sub
    Dim wordApp As Object: Set wordApp = M_Word_Engine.GetWordApp()
    Dim wordDoc As Object: Set wordDoc = M_Word_Engine.GetWordDoc()

    ' Usa o objeto Selection para construir o documento
    With wordApp.Selection
        ' Título
        .ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Font.Bold = True: .Font.Underline = wdUnderlineSingle: .Font.Size = 14
        .TypeText "LAUDO TÉCNICO"
        .TypeParagraph: .TypeParagraph
        
        .Font.Name = "Arial": .Font.Size = 12: .Font.Underline = wdUnderlineNone: .Font.Bold = False
        .ParagraphFormat.Alignment = wdAlignParagraphJustify
        .ParagraphFormat.FirstLineIndent = wordApp.CentimetersToPoints(2)

        ' Parágrafo 1
        .TypeText "Eu, "
        .Font.Bold = True: .TypeText dadosTecnico("Nome do Técnico"): .Font.Bold = False
        .Font.Bold = True: .TypeText ", " & dadosTecnico("Formação") & ", CFT: " & dadosTecnico("Registro (CFT/CREA)")
        .Font.Bold = False: .TypeText ", sob efeitos do termo de responsabilidade técnica "
        .Font.Bold = True: .TypeText "TRT: " & dadosTecnico("TRT/ART"): .Font.Bold = False
        .TypeText ", Atesto, sob as penas da lei, que efetuei pessoalmente o levantamento da área e que os valores corretos dos azimutes e distâncias e a identificação das confrontações são as apresentadas nesta oportunidade, na planta e no memorial descritivo que acompanham o presente laudo."
        .TypeParagraph
        
        ' Parágrafo 2
        .TypeText "Ao efetuar os trabalhos, constatei o seguinte: o imóvel "
        .Font.Bold = True: .TypeText dadosPropriedade("Denominação"): .Font.Bold = False
        .Font.Bold = True: .TypeText ", Situado Na " & dadosPropriedade("Endereço Propriedade") & ", "
        .Font.Bold = True: .TypeText "Matricula: " & dadosPropriedade("Matrícula"): .Font.Bold = False
        .Font.Bold = True: .TypeText " de propriedade de "
        .Font.Bold = True: .TypeText dadosPropriedade("Proprietário"): .Font.Bold = False
        .TypeText ", "
        .Font.Bold = True: .TypeText "CPF: " & dadosPropriedade("CPF"): .Font.Bold = False
        .TypeText ", " & dadosPropriedade("Nacionalidade") & ", " & dadosPropriedade("Profissão") & ", " & dadosPropriedade("Estado Civil") & ", maior, portador da "
        .Font.Bold = False: .TypeText "CI nº " & dadosPropriedade("RG"): .Font.Bold = False
        .Font.Bold = False: .TypeText ", expedida por " & dadosPropriedade("Expedição"): .Font.Bold = False
        .Font.Bold = False: .TypeText " em " & dadosPropriedade("Data Expedição"): .Font.Bold = False
        .TypeText ", residente e domiciliado na " & dadosPropriedade("Endereço Proprietário") & ", possui descrição tabular precária, devido o uso de equipamentos de pouca precisão da época."
        .TypeParagraph
        
        ' Parágrafo 3
        .TypeText "O presente levantamento foi efetuado com aparelhos geodésicos GPS de altíssima precisão, intra-muros, uma vez que as divisas são claras (cerca de arames bem antigas) e respeitadas há muitos anos. Além disso, todos os confrontantes confirmaram que a referida cerca respeita os limites de seus imóveis."
        .TypeParagraph: .TypeParagraph: .TypeParagraph
        
        ' Data
        .ParagraphFormat.FirstLineIndent = 0
        .ParagraphFormat.Alignment = wdAlignParagraphCenter
        
        Dim dataTexto As String, dataCapitalizada As String
        dataTexto = Format(Date, "dd") & " de " & Format(Date, "mmmm") & " de " & Format(Date, "yyyy")
        dataCapitalizada = StrConv(dataTexto, vbProperCase)
        dataTexto = Replace(dataCapitalizada, " De ", " de ")
        
        .ParagraphFormat.Alignment = wdAlignParagraphRight
        .Font.Bold = True: .TypeText dadosPropriedade("Município/UF") & ", " & dataTexto & "."
        
        
'        Dim dataTexto As String
'        dataTexto = StrConv(Format(Date, "dd 'de' mmmm 'de' yyyy"), vbProperCase)
'        dataTexto = Replace(dataTexto, " De ", " de ")
        '.TypeText dadosPropriedade("Município/UF") & ", " & dataTexto & "."
        .TypeParagraph: .TypeParagraph: .TypeParagraph: .TypeParagraph
        
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
    
            .cell(2, 1).Range.Text = "____________________________________" & vbCrLf & "Responsável Técnico" & vbCrLf & dadosTecnico("Nome do Técnico") & vbCrLf & dadosTecnico("Formação") & vbCrLf & "CFT: " & dadosTecnico("Registro (CFT/CREA)") & " / INCRA: " & dadosTecnico("Cód. Incra") & vbCrLf & dadosTecnico("TRT/ART")
            .cell(2, 1).Range.Paragraphs(1).Range.Font.Bold = False
            .cell(2, 1).Range.Paragraphs(2).Range.Font.Bold = True
            .cell(2, 1).Range.Paragraphs(3).Range.Font.Bold = False
            .cell(2, 1).Range.Paragraphs(4).Range.Font.Bold = False
            .cell(2, 1).Range.Paragraphs(5).Range.Font.Bold = False
            .cell(2, 1).Range.Paragraphs(6).Range.Font.Bold = False
            .cell(2, 1).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
        End With
        
    End With
    
    Dim nomeArquivo As String
    nomeArquivo = "Laudo Técnico - " & M_Utils.File_SanitizeName(dadosPropriedade("Denominação"))
    
    Dim caminho As String
    caminho = M_Word_Engine.Word_Teardown(nomeArquivo, False)
    
    If caminho <> "" Then MsgBox "Laudo Técnico gerado com SUCESSO!", vbInformation
    Unload frmAguarde
    Exit Sub

ErroWord:
    On Error Resume Next
    Unload frmAguarde
    On Error GoTo 0
    MsgBox "ERRO ao gerar o Laudo Técnico: " & Err.Description, vbCritical
    If Not wordApp Is Nothing Then wordApp.Quit SaveChanges:=False
    Set wordApp = Nothing
End Sub

' =========================================================================================
' MACRO PARA GERAR O DOCUMENTO "LAUDO TÉCNICO" EM PDF
' =========================================================================================
Public Sub GerarLaudoTecnicoPDF(dadosPropriedade As Object, dadosTecnico As Object)
    Call GerarLaudoTecnicoWord(dadosPropriedade, dadosTecnico, True)
End Sub
