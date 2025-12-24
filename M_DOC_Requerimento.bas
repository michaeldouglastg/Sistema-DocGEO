Attribute VB_Name = "M_DOC_Requerimento"
Option Explicit

' =========================================================================================
' FUNÇÃO PARA GERAR O TEXTO DO REQUERIMENTO PARA O PREVIEW (VERSÃO FINAL)
' =========================================================================================
Public Function GerarTextoRequerimento(dadosPropriedade As Object, dadosTecnico As Object) As String
    
    Dim textoReq As String
    
    ' Destinatário
    textoReq = "ILMO SR. Do " & UCase(dadosPropriedade("Cartório (CNS)")) & " DE " & UCase(dadosPropriedade("Município/UF")) & String(4, vbCrLf)
    
    ' Parágrafo 1: Dados do Proprietário e Pedido
    textoReq = textoReq & vbTab & "Eu, " & dadosPropriedade("Proprietário") & ", " & "CPF: " & dadosPropriedade("CPF")
    textoReq = textoReq & "; abaixo assinado, vem, nos termos da legislação vigente, especialmente, da Lei n° 10.931, de 02.08.2004, que alterou os arts. 212 e 213 da Lei n° 6.015/73 – LRP, requerer a "
    textoReq = textoReq & "AVERBAÇÃO E RETIFICAÇÃO"
    textoReq = textoReq & " do Georreferenciamento, do imóvel de sua propriedade; para tanto expõe o seguinte e ao final requer:" & String(2, vbCrLf)
    
    ' Item 1: Dados do Imóvel
    textoReq = textoReq & "1 – A requerente é proprietária do imóvel, " & dadosPropriedade("Denominação") & ", "
    textoReq = textoReq & "Matricula:" & dadosPropriedade("Matrícula")
    textoReq = textoReq & ", com o levantamento feito com GPS de altíssima precisão foi encontrada uma "
    textoReq = textoReq & "Área Total de " & dadosPropriedade("Natureza/Área") & " ha"
    textoReq = textoReq & ", sendo assim pedimos a atualização da matricula com as descrições em coordenadas Geográficas como segue memorial em anexo." & String(2, vbCrLf)
    
    ' Sub-item 1
    textoReq = textoReq & "- Localizada no município e comarca de, " & dadosPropriedade("Município/UF")
    textoReq = textoReq & ", e cadastrada no INCRA sob o código nº " & dadosPropriedade("Cód. Incra/SNCR") & "." & String(2, vbCrLf)
    
    ' Item 2
    textoReq = textoReq & "2 – Por todo o exposto vem requerer o seguinte; " & "AVERBAÇÃO E RETIFICAÇÃO"
    textoReq = textoReq & " do Georreferenciamento, junto ao Sigef (Sistema de Gestão Fundiária) sendo assim o proprietário requer, nos termos do art. 212 e 213 da Lei nº 6.015/73. Para tal, fazem a juntada de novos trabalhos topográficos e demais documentos para a devida avaliação e decisão." & String(2, vbCrLf)
    
    ' Item 3
    textoReq = textoReq & "3 – A requerente declara, sob penas da lei, juntamente com o " & "Responsável Técnico"
    textoReq = textoReq & ", que efetuou levantamento topográfico, " & dadosTecnico("Nome do Técnico") & ", " & dadosTecnico("Formação")
    textoReq = textoReq & ", " & "CFT: " & dadosTecnico("Registro (CFT/CREA)")
    textoReq = textoReq & ", que também assina este requerimento, que todas as informações e dados juntados a este requerimento são verdadeiros, bem como declaram terem conhecimento das disposições legais do art. 213, § 14, da Lei n. 6.015/73 - LRP:" & String(2, vbCrLf)
    
    ' Citação da Lei
    textoReq = textoReq & vbTab & """" & "Art. 213 – Verificado a qualquer tempo não serem verdadeiras os fatos constantes do memorial descritivo, responderá o requerente e o profissional que o elaborou pelos prejuízos causados, independentemente das sanções disciplinares e penais" & """" & "." & String(4, vbCrLf)
    
    ' Data
    Dim dataTexto As String
    dataTexto = StrConv(Format(Date, "dd 'de' mmmm 'de' yyyy"), vbProperCase)
    dataTexto = Replace(dataTexto, " De ", " de ")
    textoReq = textoReq & vbTab & vbTab & dadosPropriedade("Município/UF") & ", " & dataTexto & "." & String(4, vbCrLf)
    
    ' Assinaturas
    textoReq = textoReq & "____________________________________" & vbCrLf
    textoReq = textoReq & "Proprietário do Imóvel" & vbCrLf
    textoReq = textoReq & dadosPropriedade("Proprietário") & vbCrLf
    textoReq = textoReq & "CPF: " & dadosPropriedade("CPF") & String(4, vbCrLf)
    
    textoReq = textoReq & "____________________________________" & vbCrLf
    textoReq = textoReq & "Responsável Técnico" & vbCrLf
    textoReq = textoReq & dadosTecnico("Nome do Técnico") & vbCrLf
    textoReq = textoReq & dadosTecnico("Formação") & vbCrLf
    textoReq = textoReq & dadosTecnico("Registro (CFT/CREA)") & " / INCRA: " & dadosTecnico("Cód. Incra")
    
    ' Retorna o texto completo
    GerarTextoRequerimento = textoReq
    
End Function

' =========================================================================================
' MACRO PARA GERAR O REQUERIMENTO EM WORD (ADAPTADA PARA RECEBER DADOS)
' =========================================================================================
Public Sub GerarRequerimentoWord(dadosPropriedade As Object, dadosTecnico As Object, Optional gerarComoPDF As Boolean = False)
    
    On Error GoTo ErroWord
    
    ' --- ETAPA 1: Coleta e Filtro de Dados ---
    Dim wsPrincipal As Worksheet: Set wsPrincipal = ThisWorkbook.Sheets(M_Config.App_GetNomeAbaAtiva())
    Dim loPrincipal As ListObject: Set loPrincipal = wsPrincipal.ListObjects(M_Config.App_GetNomeTabelaAtiva())
    Dim i As Long
    
    frmAguarde.Show vbModeless
    frmAguarde.AtualizarStatus "Gerando Requerimento..."
    
    If Not M_Word_Engine.Word_Setup(False, 1.27, 1.27, 1.27, 1.27) Then Exit Sub
    Dim wordApp As Object: Set wordApp = M_Word_Engine.GetWordApp()
    Dim wordDoc As Object: Set wordDoc = M_Word_Engine.GetWordDoc()

    ' Usa o objeto Selection para construir o documento
    With wordApp.Selection
        ' Destinatário
        .ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Font.Underline = wdUnderlineSingle: .Font.Size = 14: .Font.Bold = True
        .TypeText "ILMO SR. Do "
        .TypeText UCase(dadosPropriedade("Cartório (CNS)")) & " DE " & UCase(dadosPropriedade("Município/UF"))
        .TypeParagraph: .TypeParagraph
        
        ' Parágrafo 1: Dados do Proprietário e Pedido
        .ParagraphFormat.Alignment = wdAlignParagraphJustify
        .Font.Underline = wdUnderlineNone: .Font.Size = 12: .Font.Bold = False
        .TypeText "Eu, "
        .Font.Bold = True: .TypeText dadosPropriedade("Proprietário"): .Font.Bold = False
        .TypeText ", "
        .Font.Bold = True: .TypeText "CPF: " & dadosPropriedade("CPF"): .Font.Bold = False
        .TypeText "; abaixo assinado, vem, nos termos da legislação vigente, especialmente, da Lei n° 10.931, de 02.08.2004, que alterou os arts. 212 e 213 da Lei n° 6.015/73 – LRP, requerer a "
        .Font.Bold = True: .TypeText "AVERBAÇÃO E RETIFICAÇÃO": .Font.Bold = False
        .TypeText " do Georreferenciamento, do imóvel de sua propriedade; para tanto expõe o seguinte e ao final requer:"
        .TypeParagraph: .TypeParagraph
        
        ' Item 1: Dados do Imóvel
        .ParagraphFormat.FirstLineIndent = 0
        .TypeText "1 – A requerente é proprietária do imóvel, "
        .Font.Bold = True: .TypeText dadosPropriedade("Denominação"): .Font.Bold = False
        .TypeText ", "
        .Font.Bold = True: .TypeText "Matricula:" & dadosPropriedade("Matrícula"): .Font.Bold = False
        .TypeText ", com o levantamento feito com GPS de altíssima precisão foi encontrada uma "
        .Font.Bold = True: .TypeText "Área Total de " & dadosPropriedade("Natureza/Área") & " ha": .Font.Bold = False
        .TypeText ", sendo assim pedimos a atualização da matricula com as descrições em coordenadas Geográficas como segue memorial em anexo."
        .TypeParagraph: .TypeParagraph
        
        ' Sub-item 1
        .TypeText "- Localizada no município e comarca de, "
        .Font.Bold = True: .TypeText dadosPropriedade("Município/UF"): .Font.Bold = False
        .TypeText ", e cadastrada no INCRA sob o código nº "
        .Font.Bold = True: .TypeText dadosPropriedade("Cód. Incra/SNCR"): .Font.Bold = False
        .TypeText "."
        .TypeParagraph: .TypeParagraph
        
        ' Item 2
        .TypeText "2 – Por todo o exposto vem requerer o seguinte; "
        .Font.Bold = True: .TypeText "AVERBAÇÃO E RETIFICAÇÃO": .Font.Bold = False
        .TypeText " do Georreferenciamento, junto ao Sigef (Sistema de Gestão Fundiária) sendo assim o proprietário requer, nos termos do art. 212 e 213 da Lei nº 6.015/73. Para tal, fazem a juntada de novos trabalhos topográficos e demais documentos para a devida avaliação e decisão."
        .TypeParagraph: .TypeParagraph
        
        ' Item 3
        .TypeText "3 – A requerente declara, sob penas da lei, juntamente com o "
        .Font.Bold = True: .TypeText "Responsável Técnico": .Font.Bold = False
        .TypeText ", que efetuou levantamento topográfico, "
        .Font.Bold = True: .TypeText dadosTecnico("Nome do Técnico"): .Font.Bold = False
        .TypeText ", "
        .Font.Bold = True: .TypeText dadosTecnico("Formação"): .Font.Bold = False
        .TypeText ", "
        .Font.Bold = True: .TypeText "CFT: " & dadosTecnico("Registro (CFT/CREA)"): .Font.Bold = False
        .TypeText ", que também assina este requerimento, que todas as informações e dados juntados a este requerimento são verdadeiros, bem como declaram terem conhecimento das disposições legais do art. 213, § 14, da Lei n. 6.015/73 - LRP:"
        .TypeParagraph: .TypeParagraph
        
        ' Citação da Lei
        .ParagraphFormat.LeftIndent = wordApp.CentimetersToPoints(2)
        .Font.Size = 9
        .TypeText Chr(34) & "Art. 213 – Verificado a qualquer tempo não serem verdadeiras os fatos constantes do memorial descritivo, responderá o requerente e o profissional que o elaborou pelos prejuízos causados, independentemente das sanções disciplinares e penais" & Chr(34) & "."
        .TypeParagraph
        .ParagraphFormat.LeftIndent = 0 ' Reseta a indentação
        .Font.Size = 12
        .TypeParagraph: .TypeParagraph: .TypeParagraph:
        
        ' Data
        Dim dataTexto As String
        dataTexto = StrConv(Format(Date, "dd 'de' mmmm 'de' yyyy"), vbProperCase)
        dataTexto = Replace(dataTexto, " De ", " de ")
        .ParagraphFormat.Alignment = wdAlignParagraphRight
        .Font.Bold = True: .TypeText dadosPropriedade("Município/UF") & ", " & dataTexto & "."
        .Font.Bold = False
        .TypeParagraph: .TypeParagraph: .TypeParagraph: .TypeParagraph
        
        ' Assinaturas
        .ParagraphFormat.Alignment = wdAlignParagraphCenter
        .TypeText "____________________________________"
        .TypeParagraph
        .Font.Bold = True: .TypeText "Proprietário do Imóvel": .Font.Bold = False
        .TypeParagraph
        .TypeText dadosPropriedade("Proprietário")
        .TypeParagraph
        .TypeText "CPF: " & dadosPropriedade("CPF")
        .TypeParagraph: .TypeParagraph: .TypeParagraph: .TypeParagraph:
        
        .TypeText "____________________________________"
        .TypeParagraph
        .Font.Bold = True: .TypeText "Responsável Técnico": .Font.Bold = False
        .TypeParagraph
        .TypeText dadosTecnico("Nome do Técnico")
        .TypeParagraph
        .TypeText dadosTecnico("Formação")
        .TypeParagraph
        .TypeText dadosTecnico("Registro (CFT/CREA)") & " / INCRA: " & dadosTecnico("Cód. Incra")
    End With
    
    Dim nomeArquivo As String
    nomeArquivo = "Requerimento - " & M_Utils.File_SanitizeName(dadosPropriedade("Denominação"))
    
    Dim caminho As String
    caminho = M_Word_Engine.Word_Teardown(nomeArquivo, False)
    
    If caminho <> "" Then MsgBox "Requerimento gerado com SUCESSO!", vbInformation
    Unload frmAguarde
    Exit Sub

ErroWord:
    On Error Resume Next
    Unload frmAguarde
    On Error GoTo 0
    MsgBox "ERRO ao gerar o Requerimento: " & Err.Description, vbCritical
    If Not wordApp Is Nothing Then wordApp.Quit SaveChanges:=False
    Set wordApp = Nothing
End Sub

' =========================================================================================
' MACRO PARA GERAR O REQUERIMENTO EM PDF
' =========================================================================================
Public Sub GerarRequerimentoPDF(dadosPropriedade As Object, dadosTecnico As Object)
    Call GerarRequerimentoWord(dadosPropriedade, dadosTecnico, True)
End Sub
