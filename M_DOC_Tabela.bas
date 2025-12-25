Attribute VB_Name = "M_DOC_Tabela"
Option Explicit

' =========================================================================================
' FUN��O PARA GERAR O TEXTO DA TABELA ANAL�TICA PARA O PREVIEW
' =========================================================================================
Public Function GerarTextoTabelaAnalitica(dadosPropriedade As Object, dadosTecnico As Object) As String
    On Error GoTo ErroFuncao

    Dim wsPrincipal As Worksheet: Set wsPrincipal = ThisWorkbook.Sheets(M_Config.App_GetNomeAbaAtiva())
    Dim loPrincipal As ListObject: Set loPrincipal = wsPrincipal.ListObjects(M_Config.App_GetNomeTabelaAtiva())
    Dim wsConversao As Worksheet: Set wsConversao = ThisWorkbook.Sheets("TEMP_CONVERSAO")
    Dim loConversao As ListObject: Set loConversao = wsConversao.ListObjects("tbl_Conversao")
    Dim textoFinal As String, i As Long, perimetroTotal As Double

    ' C�lculo de per�metro seguro
    Dim cell As Range
    perimetroTotal = 0
    For Each cell In loPrincipal.ListColumns("Dist�ncia").DataBodyRange.Cells
        If IsNumeric(cell.Value) Then
            perimetroTotal = perimetroTotal + CDbl(cell.Value)
        End If
    Next cell

    ' T�tulo
    textoFinal = "TABELA ANAL�TICA" & vbCrLf & vbCrLf

    ' Cabe�alho
    textoFinal = textoFinal & "Im�vel: " & vbTab & vbTab & dadosPropriedade("Denomina��o") & vbCrLf
    textoFinal = textoFinal & "Propriet�rio: " & vbTab & dadosPropriedade("Propriet�rio") & vbCrLf
    textoFinal = textoFinal & "Munic�pio: " & vbTab & vbTab & dadosPropriedade("Munic�pio/UF") & vbCrLf
    textoFinal = textoFinal & "Estado: " & vbTab & vbTab & dadosPropriedade("Estado") & vbCrLf
    textoFinal = textoFinal & "Sistema UTM: " & vbTab & dadosPropriedade("Sistema UTM") & vbCrLf
    textoFinal = textoFinal & "�rea medida e demarcada: " & vbTab & Format(dadosPropriedade("Area (SGL)"), "#,##0.0000") & " hectares" & vbCrLf
    textoFinal = textoFinal & "Per�metro demarcado: " & vbTab & Format(perimetroTotal, "#,##0.00") & " metros" & vbCrLf & vbCrLf

    ' Descri��o
    textoFinal = textoFinal & "DESCRI��O" & vbCrLf
    textoFinal = textoFinal & String(150, "-") & vbCrLf
    textoFinal = textoFinal & "De" & vbTab & "Para" & vbTab & "Coord. N(Y)" & vbTab & "Coord. E(X)" & vbTab & "Azimute" & vbTab & "Dist�ncia" & vbCrLf
    textoFinal = textoFinal & String(150, "-") & vbCrLf

    ' Corpo da tabela
    If loPrincipal.ListRows.Count > 0 Then
        For i = 1 To loPrincipal.ListRows.Count
            Dim utmN As Variant, utmE As Variant, dist As Variant

            ' Verifica se a linha correspondente existe na tabela de convers�o
            If i <= loConversao.ListRows.Count Then
                utmN = loConversao.ListRows(i).Range(2).Value ' Coord. N(Y)
                utmE = loConversao.ListRows(i).Range(3).Value ' Coord. E(X)
            Else
                utmN = "N/A"
                utmE = "N/A"
            End If

            dist = loPrincipal.ListRows(i).Range(7).Value

            textoFinal = textoFinal & loPrincipal.ListRows(i).Range(1).Value & vbTab ' De
            textoFinal = textoFinal & loPrincipal.ListRows(i).Range(5).Value & vbTab ' Para

            ' Formata apenas se for num�rico
            If IsNumeric(utmN) Then textoFinal = textoFinal & Format(utmN, "#,##0.00") & vbTab Else textoFinal = textoFinal & utmN & vbTab
            If IsNumeric(utmE) Then textoFinal = textoFinal & Format(utmE, "#,##0.00") & vbTab Else textoFinal = textoFinal & utmE & vbTab

            textoFinal = textoFinal & loPrincipal.ListRows(i).Range(6).Value & vbTab ' Azimute

            ' Formata apenas se for num�rico
            If IsNumeric(dist) Then textoFinal = textoFinal & Format(dist, "#,##0.00 m") & vbCrLf Else textoFinal = textoFinal & dist & vbCrLf
        Next i
    End If

    textoFinal = textoFinal & String(150, "-") & vbCrLf
    textoFinal = textoFinal & "Per�metro: " & Format(perimetroTotal, "#,##0.00 m") & vbTab & vbTab & "�rea: " & Format(dadosPropriedade("Area (SGL)"), "#,##0.0000 m�") & vbCrLf & vbCrLf

    ' Data
    Dim dataTexto As String, dataCapitalizada As String
    dataTexto = Format(Date, "dd") & " de " & Format(Date, "mmmm") & " de " & Format(Date, "yyyy")
    dataCapitalizada = StrConv(dataTexto, vbProperCase)
    dataTexto = Replace(dataCapitalizada, " De ", " de ")

    textoFinal = textoFinal & vbTab & vbTab & vbTab & dadosPropriedade("Munic�pio/UF") & ", " & dataTexto & "." & vbCrLf & vbCrLf & vbCrLf

    ' Assinatura
    textoFinal = textoFinal & "____________________________________" & vbCrLf
    textoFinal = textoFinal & "Respons�vel T�cnico" & vbCrLf
    textoFinal = textoFinal & dadosTecnico("Nome do T�cnico") & vbCrLf
    textoFinal = textoFinal & dadosTecnico("Forma��o") & vbCrLf
    textoFinal = textoFinal & dadosTecnico("Registro (CFT/CREA)") & " / INCRA: " & dadosTecnico("C�d. Incra") & vbCrLf
    textoFinal = textoFinal & dadosTecnico("TRT/ART")

    GerarTextoTabelaAnalitica = textoFinal
    Exit Function

ErroFuncao:
    GerarTextoTabelaAnalitica = "Ocorreu um erro ao gerar o texto da Tabela Anal�tica: " & Err.Description
End Function

' =========================================================================================
' MACRO PARA GERAR A TABELA ANAL�TICA EM WORD
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
    frmAguarde.AtualizarStatus "Gerando Tabela Anal�tica..."

    ' C�lculo de per�metro seguro
    Dim perimetroTotal As Double, cell As Range
    perimetroTotal = 0
    For Each cell In loPrincipal.ListColumns("Dist�ncia").DataBodyRange.Cells
        If IsNumeric(cell.Value) Then
            perimetroTotal = perimetroTotal + CDbl(cell.Value)
        End If
    Next cell

    ' --- ETAPA 2: Gerar e Formatar o Documento Word ---
    If Not M_Word_Engine.Word_Setup(False, 2.5, 2.5, 2.25, 3#) Then Exit Sub
    Dim wordApp As Object: Set wordApp = M_Word_Engine.GetWordApp()
    Dim wordDoc As Object: Set wordDoc = M_Word_Engine.GetWordDoc()

    With wordApp.Selection
        ' T�tulo
        .ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Font.Bold = True: .Font.Underline = wdUnderlineSingle: .Font.Size = 14
        .TypeText "TABELA ANAL�TICA"
        .TypeParagraph: .TypeParagraph

        .Font.Name = "Arial": .Font.Bold = False: .Font.Size = 12: .Font.Underline = wdUnderlineNone

        ' --- CABE�ALHO EM DUAS COLUNAS COM TABELA INVIS�VEL ---
        Dim tblHeader As Word.Table
        Set tblHeader = wordDoc.Tables.Add(Range:=.Range, NumRows:=7, NumColumns:=2)
        tblHeader.Borders.Enable = False

        With tblHeader
            SetCellTextBoldLabel .cell(1, 1), "Im�vel: ", dadosPropriedade("Denomina��o")
            SetCellTextBoldLabel .cell(2, 1), "Propriet�rio: ", dadosPropriedade("Propriet�rio")
            SetCellTextBoldLabel .cell(3, 1), "Munic�pio: ", dadosPropriedade("Munic�pio/UF")
            SetCellTextBoldLabel .cell(4, 1), "Estado: ", dadosPropriedade("Estado")
            SetCellTextBoldLabel .cell(5, 1), "Sistema UTM: ", dadosPropriedade("Sistema UTM")
            SetCellTextBoldLabel .cell(6, 1), "�rea medida e demarcada: ", Format(dadosPropriedade("Area (SGL)"), "#,##0.0000") & " hectares"
            SetCellTextBoldLabel .cell(7, 1), "Per�metro demarcado: ", Format(perimetroTotal, "#,##0.00") & " metros"
            .Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
        End With

        ' Move o cursor para FORA da tabela
        Dim rng As Word.Range
        Set rng = wordDoc.Content
        rng.Collapse wdCollapseEnd
        rng.Select

        .TypeParagraph

        ' Subt�tulo "DESCRI��O"
        .ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Font.Bold = True: .Font.Size = 12
        .TypeText "DESCRI��O"
        .TypeParagraph

        ' Tabela de Coordenadas
        Dim tblWord As Word.Table, numLinhasTabela As Long
        numLinhasTabela = loPrincipal.ListRows.Count + 2  ' +1 cabe�alho, +1 rodap�
        Set tblWord = wordDoc.Tables.Add(Range:=.Range, NumRows:=numLinhasTabela, NumColumns:=6)

        With tblWord
            .Borders.Enable = True
            .Range.Font.Name = "Arial": .Range.Font.Size = 9
            .Range.ParagraphFormat.LineSpacingRule = wdLineSpaceSingle
            .Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
            .Range.Cells.VerticalAlignment = wdCellAlignVerticalCenter

            ' Cabe�alho
            With .Rows(1).Range
                .Font.Bold = True
                .Shading.BackgroundPatternColor = wdColorGray15
            End With

            .cell(1, 1).Range.Text = "De"
            .cell(1, 2).Range.Text = "Para"
            .cell(1, 3).Range.Text = "Coord. N(Y)"
            .cell(1, 4).Range.Text = "Coord. E(X)"
            .cell(1, 5).Range.Text = "Azimute"
            .cell(1, 6).Range.Text = "Dist�ncia"

            ' Corpo da tabela
            If loPrincipal.ListRows.Count > 0 Then
                For i = 1 To loPrincipal.ListRows.Count
                    .cell(i + 1, 1).Range.Text = loPrincipal.ListRows(i).Range(1).Value ' De
                    .cell(i + 1, 2).Range.Text = loPrincipal.ListRows(i).Range(5).Value ' Para

                    ' Verifica se existe a linha correspondente na tabela de convers�o
                    If i <= loConversao.ListRows.Count Then
                        .cell(i + 1, 3).Range.Text = Format(loConversao.ListRows(i).Range(2).Value, "#,##0.00") ' UTM N
                        .cell(i + 1, 4).Range.Text = Format(loConversao.ListRows(i).Range(3).Value, "#,##0.00") ' UTM E
                    Else
                        .cell(i + 1, 3).Range.Text = "N/A"
                        .cell(i + 1, 4).Range.Text = "N/A"
                    End If

                    .cell(i + 1, 5).Range.Text = loPrincipal.ListRows(i).Range(6).Value ' Azimute
                    .cell(i + 1, 6).Range.Text = Format(loPrincipal.ListRows(i).Range(7).Value, "#,##0.00 m") ' Dist�ncia
                Next i
            End If

            ' Rodap� da tabela
            With .Rows.Last.Range
                .Font.Bold = True
                .Shading.BackgroundPatternColor = wdColorGray15
            End With
            .cell(numLinhasTabela, 1).Range.Text = "Per�metro: " & Format(perimetroTotal, "#,##0.00 m")
            .cell(numLinhasTabela, 4).Range.Text = "�rea: " & Format(dadosPropriedade("Area (SGL)"), "#,##0.0000 m�")
            .cell(numLinhasTabela, 1).Merge MergeTo:=.cell(numLinhasTabela, 3)
            .cell(numLinhasTabela, 2).Merge MergeTo:=.cell(numLinhasTabela, 3)
        End With
    End With

    ' Move o cursor para FORA da tabela
    Set rng = wordDoc.Content
    rng.Collapse wdCollapseEnd
    rng.Select

    With wordApp.Selection
        .TypeParagraph
        .TypeParagraph

        ' Data
        Dim dataTexto As String, dataCapitalizada As String
        dataTexto = Format(Date, "dd") & " de " & Format(Date, "mmmm") & " de " & Format(Date, "yyyy")
        dataCapitalizada = StrConv(dataTexto, vbProperCase)
        dataTexto = Replace(dataCapitalizada, " De ", " de ")

        .ParagraphFormat.Alignment = wdAlignParagraphRight
        .Font.Bold = True: .Font.Size = 12
        .TypeText dadosPropriedade("Munic�pio/UF") & ", " & dataTexto & "."
        .TypeParagraph: .TypeParagraph: .TypeParagraph: .TypeParagraph
    End With

    ' Bloco de Assinaturas
    Dim tblAssinaturas As Word.Table
    Set tblAssinaturas = wordDoc.Tables.Add(Range:=wordApp.Selection.Range, NumRows:=1, NumColumns:=1)
    tblAssinaturas.Borders.Enable = False

    With tblAssinaturas
        .Range.Font.Size = 12
        .cell(1, 1).Range.Text = "____________________________________" & vbCrLf & _
                                 "Respons�vel T�cnico" & vbCrLf & _
                                 dadosTecnico("Nome do T�cnico") & vbCrLf & _
                                 dadosTecnico("Forma��o") & vbCrLf & _
                                 dadosTecnico("Registro (CFT/CREA)") & " / INCRA: " & dadosTecnico("C�d. Incra") & vbCrLf & _
                                 dadosTecnico("TRT/ART")
        .cell(1, 1).Range.Paragraphs(1).Range.Font.Bold = False
        .cell(1, 1).Range.Paragraphs(2).Range.Font.Bold = True
        .cell(1, 1).Range.Paragraphs(3).Range.Font.Bold = False
        .cell(1, 1).Range.Paragraphs(4).Range.Font.Bold = False
        .cell(1, 1).Range.Paragraphs(5).Range.Font.Bold = False
        .cell(1, 1).Range.Paragraphs(6).Range.Font.Bold = False
        .cell(1, 1).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    End With

    ' --- ETAPA 3: FINALIZA��O ---
    Dim nomeArquivo As String
    nomeArquivo = "Tabela Anal�tica - " & M_Utils.File_SanitizeName(dadosPropriedade("Denomina��o"))

    Dim caminho As String
    caminho = M_Word_Engine.Word_Teardown(nomeArquivo, gerarComoPDF)

    If caminho <> "" Then MsgBox "Tabela Anal�tica gerada com SUCESSO!", vbInformation
    Unload frmAguarde
    Exit Sub

ErroWord:
    On Error Resume Next
    Unload frmAguarde
    On Error GoTo 0
    MsgBox "ERRO ao gerar a Tabela Anal�tica: " & Err.Description, vbCritical
    If Not wordApp Is Nothing Then wordApp.Quit SaveChanges:=False
    Set wordApp = Nothing
End Sub

' =========================================================================================
' MACRO PARA GERAR A TABELA ANAL�TICA EM PDF
' =========================================================================================
Public Sub GerarTabelaAnaliticaPDF(dadosProp As Object, dadosTec As Object)
    Call GerarTabelaAnaliticaWord(dadosProp, dadosTec, True)
End Sub
