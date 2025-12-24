Attribute VB_Name = "M_Importacao"
Option Explicit
' ==============================================================================
' MODULO: M_IMPORTACAO
' DESCRICAO: IMPORTACAO DE DADOS (CSV, PDF SIGEF, KML)
' ==============================================================================

' ==============================================================================
' 1. IMPORTACAO DE ARQUIVOS CSV
' ==============================================================================
Public Sub Imp_LerArquivoCSV()
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim wsPrincipal As Worksheet
    Dim loPrincipal As ListObject
    Dim dictVertices As Object
    Dim caminhoArquivo1 As Variant, caminhoArquivo2 As Variant
    Dim nomeAba As String, nomeTabela As String
    
    nomeAba = M_Config.App_GetNomeAbaAtiva()
    nomeTabela = M_Config.App_GetNomeTabelaAtiva()
    Set wsPrincipal = wb.Sheets(nomeAba)
    
    On Error Resume Next
    Set loPrincipal = wsPrincipal.ListObjects(nomeTabela)
    On Error GoTo ErroImportacao
    
    If loPrincipal Is Nothing Then
        MsgBox "A tabela '" & nomeTabela & "' nao foi encontrada na aba '" & nomeAba & "'.", vbCritical
        Exit Sub
    End If
    
    Set dictVertices = CreateObject("Scripting.Dictionary")
    
    caminhoArquivo2 = Application.GetOpenFilename("Arquivos de Texto (*.csv),*.csv", , "Selecione o CSV de COORDENADAS (X, Y, Z)")
    If caminhoArquivo2 = False Then Exit Sub
    
    caminhoArquivo1 = Application.GetOpenFilename("Arquivos de Texto (*.csv),*.csv", , "Selecione o CSV de CONFRONTANTES")
    If caminhoArquivo1 = False Then Exit Sub
    
    Call M_Utils.Utils_OtimizarPerformance(True)
    
    ' Desbloqueia planilha para escrita
    M_SheetProtection.DesbloquearPlanilha wsPrincipal
    
    Call M_Dados.Dados_LimparTudo
    
    loPrincipal.ListColumns(2).Range.NumberFormat = "@"
    loPrincipal.ListColumns(3).Range.NumberFormat = "@"
    
    ' --- ETAPA 1: LER COORDENADAS (CSV 2) ---
    Dim conteudoCSV As String
    Dim linhas() As String, linha As Variant, dadosLinha() As String
    Dim coordWKT As String, coordSplit() As String
    Dim lonDMS As String, latDMS As String
    
    conteudoCSV = LerArquivoUTF8(CStr(caminhoArquivo2))
    linhas = Split(conteudoCSV, vbLf)
    
    For Each linha In linhas
        linha = Trim(Replace(linha, vbCr, ""))
        If linha <> "" And Not UCase(Left(linha, 6)) = "QRCODE" Then
            dadosLinha = Split(linha, ";")
            If UBound(dadosLinha) >= 12 Then
                coordWKT = Replace(Replace(dadosLinha(12), "POINT (", ""), ")", "")
                coordSplit = Split(coordWKT, " ")
                lonDMS = M_Utils.Str_DD_Para_DMS(Val(coordSplit(0)))
                latDMS = M_Utils.Str_DD_Para_DMS(Val(coordSplit(1)))
                If Not dictVertices.Exists(dadosLinha(1)) Then
                    dictVertices.Add Key:=dadosLinha(1), Item:=Array(lonDMS, latDMS, dadosLinha(11))
                End If
            End If
        End If
    Next linha
    
    ' --- ETAPA 2: LER LIMITES (CSV 1) E PREENCHER TABELA ---
    conteudoCSV = LerArquivoUTF8(CStr(caminhoArquivo1))
    linhas = Split(conteudoCSV, vbLf)
    
    Dim newRow As ListRow
    Dim verticeOrigem As String
    Dim azimuteDecimal As Double
    Dim altitudeVal As Variant
    
    For Each linha In linhas
        linha = Trim(Replace(linha, vbCr, ""))
        If linha <> "" And Not UCase(Left(linha, 6)) = "QRCODE" Then
            dadosLinha = Split(linha, ";")
            If UBound(dadosLinha) >= 6 Then
                Set newRow = loPrincipal.ListRows.Add
                verticeOrigem = dadosLinha(1)
                
                With newRow
                    .Range(1).Value = verticeOrigem
                    .Range(5).Value = dadosLinha(2)
                    .Range(8).Value = dadosLinha(6)
                    .Range(9).Value = dadosLinha(3)
                    
                    azimuteDecimal = CDbl(Replace(dadosLinha(4), ".", ","))
                    .Range(6).Value = M_Utils.Str_FormatAzimute(azimuteDecimal)
                    .Range(7).Value = CDbl(Replace(dadosLinha(5), ".", ","))
                    
                    If dictVertices.Exists(verticeOrigem) Then
                        .Range(2).Value = dictVertices(verticeOrigem)(0)
                        .Range(3).Value = dictVertices(verticeOrigem)(1)
                        altitudeVal = dictVertices(verticeOrigem)(2)
                        If IsNumeric(Replace(CStr(altitudeVal), ".", ",")) Then
                            .Range(4).Value = CDbl(Replace(CStr(altitudeVal), ".", ","))
                        Else
                            .Range(4).Value = 0
                        End If
                    End If
                End With
            End If
        End If
    Next linha
    
    M_SheetProtection.BloquearPlanilha wsPrincipal
    
    Call M_App_Logica.Processo_PosImportacao
    Call M_Utils.Utils_OtimizarPerformance(False)
    
    MsgBox "Importacao CSV concluida com sucesso!", vbInformation
    Exit Sub
    
ErroImportacao:
    M_SheetProtection.BloquearPlanilha wsPrincipal
    Call M_Utils.Utils_OtimizarPerformance(False)
    MsgBox "Erro na importacao CSV: " & Err.Description, vbCritical
End Sub

Public Sub Imp_LerArquivoCSV1()
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim wsPrincipal As Worksheet
    Dim loPrincipal As ListObject
    Dim dictVertices As Object
    Dim caminhoArquivo1 As Variant, caminhoArquivo2 As Variant
    Dim nomeAba As String, nomeTabela As String
    
    nomeAba = M_Config.App_GetNomeAbaAtiva()
    nomeTabela = M_Config.App_GetNomeTabelaAtiva()
    Set wsPrincipal = wb.Sheets(nomeAba)
    
    On Error Resume Next
    Set loPrincipal = wsPrincipal.ListObjects(nomeTabela)
    On Error GoTo ErroImportacao
    
    If loPrincipal Is Nothing Then
        MsgBox "A tabela '" & nomeTabela & "' nao foi encontrada na aba '" & nomeAba & "'.", vbCritical
        Exit Sub
    End If
    
    Set dictVertices = CreateObject("Scripting.Dictionary")
    
    caminhoArquivo2 = Application.GetOpenFilename("Arquivos de Texto (*.csv),*.csv", , "Selecione o CSV de COORDENADAS (X, Y, Z)")
    If caminhoArquivo2 = False Then Exit Sub
    
    caminhoArquivo1 = Application.GetOpenFilename("Arquivos de Texto (*.csv),*.csv", , "Selecione o CSV de CONFRONTANTES")
    If caminhoArquivo1 = False Then Exit Sub
    
    Call M_Utils.Utils_OtimizarPerformance(True)
    M_SheetProtection.DesbloquearPlanilha wsPrincipal
    Call M_Dados.Dados_LimparTudo
    
    ' Formata colunas como Texto para receber o formato DMS
    loPrincipal.ListColumns(2).Range.NumberFormat = "@"
    loPrincipal.ListColumns(3).Range.NumberFormat = "@"
    
    ' --- ETAPA 1: LER COORDENADAS (CSV 2) ---
    Dim conteudoCSV As String
    Dim linhas() As String, linha As Variant, dadosLinha() As String
    Dim coordWKT As String, coordSplit() As String
    Dim lonDMS As String, latDMS As String
    Dim valLon As Double, valLat As Double
    
    conteudoCSV = LerArquivoUTF8(CStr(caminhoArquivo2))
    linhas = Split(conteudoCSV, vbLf)
    
    For Each linha In linhas
        linha = Trim(Replace(linha, vbCr, ""))
        If linha <> "" And Not UCase(Left(linha, 6)) = "QRCODE" Then
            dadosLinha = Split(linha, ";")
            If UBound(dadosLinha) >= 12 Then
                ' Limpa o WKT "POINT (X Y)"
                coordWKT = Replace(Replace(dadosLinha(12), "POINT (", ""), ")", "")
                coordSplit = Split(coordWKT, " ")
                
                ' CORREÇÃO AQUI: Tratamento robusto de Double (Ponto ou Vírgula)
                ' Substitui ponto por vírgula para garantir que o CDbl funcione no Excel BR
                If UBound(coordSplit) >= 1 Then
                    valLon = CDbl(Replace(coordSplit(0), ".", ","))
                    valLat = CDbl(Replace(coordSplit(1), ".", ","))
                    
                    ' Converte para String DMS mantendo sinal e precisão
                    lonDMS = M_Utils.Str_DD_Para_DMS(valLon)
                    latDMS = M_Utils.Str_DD_Para_DMS(valLat)
                    
                    If Not dictVertices.Exists(dadosLinha(1)) Then
                        dictVertices.Add Key:=dadosLinha(1), Item:=Array(lonDMS, latDMS, dadosLinha(11))
                    End If
                End If
            End If
        End If
    Next linha
    
    ' --- ETAPA 2: LER LIMITES (CSV 1) E PREENCHER TABELA ---
    conteudoCSV = LerArquivoUTF8(CStr(caminhoArquivo1))
    linhas = Split(conteudoCSV, vbLf)
    
    Dim newRow As ListRow
    Dim verticeOrigem As String
    Dim azimuteDecimal As Double
    Dim altitudeVal As Variant
    
    For Each linha In linhas
        linha = Trim(Replace(linha, vbCr, ""))
        If linha <> "" And Not UCase(Left(linha, 6)) = "QRCODE" Then
            dadosLinha = Split(linha, ";")
            If UBound(dadosLinha) >= 6 Then
                Set newRow = loPrincipal.ListRows.Add
                verticeOrigem = dadosLinha(1)
                
                With newRow
                    .Range(1).Value = verticeOrigem
                    .Range(5).Value = dadosLinha(2)
                    .Range(8).Value = dadosLinha(6)
                    .Range(9).Value = dadosLinha(3)
                    
                    ' Tratamento Azimute
'                    If IsNumeric(Replace(dadosLinha(4), ".", ",")) Then
'                        azimuteDecimal = CDbl(Replace(dadosLinha(4), ".", ","))
'                        ' Garante saída GMS no Azimute também
'                        .Range(6).Value = M_Utils.Str_DD_Para_DMS(azimuteDecimal)
'                    Else
'                        .Range(6).Value = dadosLinha(4)
'                    End If
'
'                    ' Distância
'                    .Range(7).Value = CDbl(Replace(dadosLinha(5), ".", ","))

                    ' Azimute SGL (Graus e Minutos)
                    If IsNumeric(Replace(dadosLinha(4), ".", ",")) Then
                        azimuteDecimal = CDbl(Replace(dadosLinha(4), ".", ","))
                        ' Usa a função específica DM
                        .Range(6).Value = M_Utils.Str_DD_Para_DM(azimuteDecimal)
                    Else
                        .Range(6).Value = dadosLinha(4)
                    End If
                    
                    ' Distância
                    .Range(7).Value = CDbl(Replace(dadosLinha(5), ".", ","))
                    
                    ' Preenche Lat/Lon/Alt se encontrou no Dicionário
                    If dictVertices.Exists(verticeOrigem) Then
                        .Range(2).Value = dictVertices(verticeOrigem)(0)
                        .Range(3).Value = dictVertices(verticeOrigem)(1)
                        altitudeVal = dictVertices(verticeOrigem)(2)
                        
                        If IsNumeric(Replace(CStr(altitudeVal), ".", ",")) Then
                            .Range(4).Value = CDbl(Replace(CStr(altitudeVal), ".", ","))
                        Else
                            .Range(4).Value = 0
                        End If
                    End If
                End With
            End If
        End If
    Next linha
    
    M_SheetProtection.BloquearPlanilha wsPrincipal
    Call M_App_Logica.Processo_PosImportacao
    Call M_Utils.Utils_OtimizarPerformance(False)
    
    MsgBox "Importacao CSV concluida com sucesso!", vbInformation
    Exit Sub
    
ErroImportacao:
    M_SheetProtection.BloquearPlanilha wsPrincipal
    Call M_Utils.Utils_OtimizarPerformance(False)
    MsgBox "Erro na importacao CSV: " & Err.Description, vbCritical
End Sub

' ==============================================================================
' 2. IMPORTACAO DE PDF (MEMORIAL SIGEF)
' ==============================================================================
Public Sub Imp_LerPDF_SIGEF()
    Dim wordApp As Object, wordDoc As Object, wordTabela As Object
    Dim loPrincipal As ListObject
    Dim wsPrincipal As Worksheet
    Dim caminhoPDF As Variant
    Dim i As Long, pontosImportados As Long
    Dim nomeAba As String, nomeTabela As String
    
    nomeAba = M_Config.App_GetNomeAbaAtiva()
    nomeTabela = M_Config.App_GetNomeTabelaAtiva()
    
    caminhoPDF = Application.GetOpenFilename("Arquivos PDF (*.pdf), *.pdf", , "Selecione o Memorial SIGEF")
    If caminhoPDF = False Then Exit Sub
    
    Call M_Utils.Utils_OtimizarPerformance(True)
    
    Set wsPrincipal = ThisWorkbook.Sheets(nomeAba)
    M_SheetProtection.DesbloquearPlanilha wsPrincipal
    
    Call M_Dados.Dados_LimparTudo
    
    Set loPrincipal = wsPrincipal.ListObjects(nomeTabela)
    loPrincipal.ListColumns(2).Range.NumberFormat = "@"
    loPrincipal.ListColumns(3).Range.NumberFormat = "@"
    
    On Error GoTo ErroPDF
    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = False
    Set wordDoc = wordApp.Documents.Open(caminhoPDF, ReadOnly:=True, Visible:=False)
    
    Call Imp_ParseCabecalhoPDF(wordDoc)
    
    If wordDoc.Tables.Count = 0 Then
        MsgBox "Nenhuma tabela encontrada no PDF.", vbExclamation
        GoTo FimPDF
    End If
    
    For Each wordTabela In wordDoc.Tables
        For i = 1 To wordTabela.Rows.Count
            If wordTabela.Rows(i).Cells.Count >= 8 Then
                Dim deStr As String
                deStr = M_Utils.Str_LimparCaractereWord(wordTabela.cell(i, 1).Range.Text)
                
                If UCase(deStr) <> "CODIGO" And deStr <> "" And deStr <> "--" Then
                    Dim novaLinha As ListRow
                    Set novaLinha = loPrincipal.ListRows.Add
                    pontosImportados = pontosImportados + 1
                    
                    With novaLinha
                        .Range(1).Value = deStr
                        .Range(2).Value = M_Utils.Str_LimparCaractereWord(wordTabela.cell(i, 2).Range.Text)
                        .Range(3).Value = M_Utils.Str_LimparCaractereWord(wordTabela.cell(i, 3).Range.Text)
                        .Range(4).Value = CDbl(Replace(M_Utils.Str_LimparCaractereWord(wordTabela.cell(i, 4).Range.Text), ".", ","))
                        .Range(5).Value = M_Utils.Str_LimparCaractereWord(wordTabela.cell(i, 5).Range.Text)
                        .Range(6).Value = M_Utils.Str_LimparCaractereWord(wordTabela.cell(i, 6).Range.Text)
                        .Range(7).Value = CDbl(Replace(M_Utils.Str_LimparCaractereWord(wordTabela.cell(i, 7).Range.Text), ".", ","))
                        .Range(8).Value = M_Utils.Str_LimparCaractereWord(wordTabela.cell(i, 8).Range.Text)
                    End With
                End If
            End If
        Next i
    Next wordTabela
    
FimPDF:
    wordDoc.Close False
    wordApp.Quit
    Set wordDoc = Nothing
    Set wordApp = Nothing
    
    M_SheetProtection.BloquearPlanilha wsPrincipal
    
    Call M_App_Logica.Processo_PosImportacao
    Call M_Utils.Utils_OtimizarPerformance(False)
    
    If pontosImportados > 0 Then MsgBox pontosImportados & " vertices importados!", vbInformation
    Exit Sub
    
ErroPDF:
    If Not wordApp Is Nothing Then wordApp.Quit
    M_SheetProtection.BloquearPlanilha wsPrincipal
    Call M_Utils.Utils_OtimizarPerformance(False)
    MsgBox "Erro ao ler PDF: " & Err.Description, vbCritical
End Sub

Private Sub Imp_ParseCabecalhoPDF(wordDoc As Object)
    Dim textoCompleto As String
    textoCompleto = Replace(wordDoc.Content.Text, vbLf, vbCr) & vbCr
    
    Dim matr As String, denom As String, area As String, prop As String
    Dim cpf As String, incra As String, mun As String, cns As String
    
    denom = M_Utils.Str_ExtractBetween(textoCompleto, "Denominacao:", "Natureza da Area:")
    area = M_Utils.Str_ExtractBetween(textoCompleto, "Natureza da Area:", "Proprietario(a):")
    prop = M_Utils.Str_ExtractBetween(textoCompleto, "Proprietario(a):", "CPF:")
    cpf = M_Utils.Str_ExtractBetween(textoCompleto, "CPF:", "Matricula do imovel:")
    matr = M_Utils.Str_ExtractBetween(textoCompleto, "Matricula do imovel:", "Codigo INCRA/SNCR:")
    incra = M_Utils.Str_ExtractBetween(textoCompleto, "Codigo INCRA/SNCR:", "Municipio/UF:")
    mun = M_Utils.Str_ExtractBetween(textoCompleto, "Municipio/UF:", "Cartorio (CNS):")
    cns = M_Utils.Str_ExtractBetween(textoCompleto, "Cartorio (CNS):", "Responsavel Tecnico(a):")
    
    Dim colunasProp As Variant, valoresProp As Variant
    colunasProp = Array(M_Config.LBL_MATRICULA, M_Config.LBL_PROPRIEDADE, M_Config.LBL_NATUREZA, _
                        M_Config.LBL_PROP_NOME, M_Config.LBL_PROP_CPF, M_Config.LBL_INCRA, _
                        M_Config.LBL_MUNICIPIO, "Cartorio (CNS)")
    valoresProp = Array(matr, denom, area, prop, cpf, incra, mun, cns)
    
    Call M_Dados.Dados_UpsertRegistro(M_Config.SH_BD_PROP, M_Config.TBL_DB_PROP, _
                                       M_Config.LBL_MATRICULA, matr, colunasProp, valoresProp)
    
    Dim respTec As String, formacao As String, cred As String, conselho As String, trt As String
    respTec = M_Utils.Str_ExtractBetween(textoCompleto, "Responsavel Tecnico(a):", "Formacao:")
    formacao = M_Utils.Str_ExtractBetween(textoCompleto, "Formacao:", "Codigo de credenciamento:")
    cred = M_Utils.Str_ExtractBetween(textoCompleto, "Codigo de credenciamento:", "Conselho Profissional:", "Sistema Geodesico")
    conselho = M_Utils.Str_ExtractBetween(textoCompleto, "Conselho Profissional:", "Documento de RT:")
    trt = M_Utils.Str_ExtractBetween(textoCompleto, "Documento de RT:", "Coordenadas:", "Area (Sistema")
    
    Dim colunasTec As Variant, valoresTec As Variant
    colunasTec = Array("Registro (CFT/CREA)", M_Config.LBL_RT_NOME, "Formacao", "Cod. Incra", M_Config.LBL_RT_ART)
    valoresTec = Array(conselho, respTec, formacao, cred, trt)
    
    Call M_Dados.Dados_UpsertRegistro(M_Config.SH_BD_TEC, M_Config.TBL_DB_TEC, _
                                       "Registro (CFT/CREA)", conselho, colunasTec, valoresTec)
End Sub

Private Function LerArquivoUTF8(caminho As String) As String
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    With stream
        .Type = 2
        .Charset = "UTF-8"
        .Open
        .LoadFromFile caminho
        LerArquivoUTF8 = .ReadText
        .Close
    End With
End Function

' ==============================================================================
' 3. IMPORTACAO DE ARQUIVO KML (GOOGLE EARTH)
' ==============================================================================
Public Sub Imp_LerArquivoKML()
    Dim xmlDoc As Object, nodeList As Object, node As Object
    Dim wsSGL As Worksheet, loSGL As ListObject
    Dim caminhoArquivo As Variant, kmlString As String
    Dim pontosImportados As Long
    Dim stream As Object
    
    caminhoArquivo = Application.GetOpenFilename("Arquivos KML (*.kml), *.kml", , "Selecione o arquivo KML")
    If caminhoArquivo = False Then Exit Sub
    
    Call M_Utils.Utils_OtimizarPerformance(True)
    
    Set wsSGL = ThisWorkbook.Sheets(M_Config.SH_SGL)
    Set loSGL = wsSGL.ListObjects(M_Config.TBL_SGL)
    
    M_SheetProtection.DesbloquearPlanilha wsSGL
    
    Call M_Dados.Dados_LimparTabela(M_Config.SH_SGL, M_Config.TBL_SGL)
    
    loSGL.ListColumns(2).Range.NumberFormat = "@"
    loSGL.ListColumns(3).Range.NumberFormat = "@"
    
    On Error GoTo ErroKML
    
    Set stream = CreateObject("ADODB.Stream")
    stream.Open
    stream.Charset = "UTF-8"
    stream.LoadFromFile caminhoArquivo
    kmlString = stream.ReadText
    stream.Close
    
    Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
    xmlDoc.async = False
    xmlDoc.SetProperty "SelectionNamespaces", "xmlns:kml='http://www.opengis.net/kml/2.2'"
    
    If Not xmlDoc.LoadXML(kmlString) Then
        If Not xmlDoc.Load(caminhoArquivo) Then
            MsgBox "Erro ao ler estrutura do KML.", vbCritical
            GoTo FimKML
        End If
    End If
    
    pontosImportados = 0
    
    ' A. Tenta ler Placemarks (Pontos isolados)
    Set nodeList = xmlDoc.SelectNodes("//kml:Placemark/kml:Point/kml:coordinates")
    If nodeList.Length = 0 Then Set nodeList = xmlDoc.SelectNodes("//Placemark/Point/coordinates")
    
    If nodeList.Length > 0 Then
        For Each node In nodeList
            pontosImportados = pontosImportados + 1
            Call Imp_ProcessarNoKML(loSGL, node.ParentNode.ParentNode, pontosImportados)
        Next node
    End If
    
    ' B. Tenta ler Poligonos/Linhas
    If pontosImportados = 0 Then
        Set node = xmlDoc.SelectSingleNode("//kml:Polygon//kml:coordinates")
        If node Is Nothing Then Set node = xmlDoc.SelectSingleNode("//kml:LineString//kml:coordinates")
        If node Is Nothing Then Set node = xmlDoc.SelectSingleNode("//Polygon//coordinates")
        
        If Not node Is Nothing Then
            Dim coordsArr() As String, i As Long
            Dim coordLimpa As String
            
            coordLimpa = node.Text
            coordLimpa = Replace(coordLimpa, vbLf, " ")
            coordLimpa = Replace(coordLimpa, vbCr, " ")
            coordLimpa = Application.WorksheetFunction.Trim(coordLimpa)
            coordsArr = Split(coordLimpa, " ")
            
            For i = LBound(coordsArr) To UBound(coordsArr)
                If Trim(coordsArr(i)) <> "" Then
                    pontosImportados = pontosImportados + 1
                    Call Imp_AdicionarPontoSGL(loSGL, "P-" & pontosImportados, CStr(coordsArr(i)))
                End If
            Next i
        End If
    End If
    
FimKML:
    Set xmlDoc = Nothing
    
    M_SheetProtection.BloquearPlanilha wsSGL
    
    If pontosImportados > 0 Then
        Call M_App_Logica.Processo_PosImportacao
        Call M_Utils.Utils_OtimizarPerformance(False)
        MsgBox pontosImportados & " pontos importados do KML com sucesso!", vbInformation
    Else
        Call M_Utils.Utils_OtimizarPerformance(False)
        MsgBox "Nenhuma coordenada encontrada no KML.", vbExclamation
    End If
    Exit Sub
    
ErroKML:
    M_SheetProtection.BloquearPlanilha wsSGL
    Call M_Utils.Utils_OtimizarPerformance(False)
    MsgBox "Erro na importacao KML: " & Err.Description, vbCritical
End Sub

Private Sub Imp_ProcessarNoKML(lo As ListObject, nodePlacemark As Object, idx As Long)
    Dim nome As String, coords As String
    Dim nodeName As Object, nodeCoord As Object
    
    Set nodeName = nodePlacemark.SelectSingleNode("kml:name")
    If nodeName Is Nothing Then Set nodeName = nodePlacemark.SelectSingleNode("name")
    If Not nodeName Is Nothing Then nome = nodeName.Text Else nome = "P-" & idx
    
    Set nodeCoord = nodePlacemark.SelectSingleNode(".//kml:coordinates")
    If nodeCoord Is Nothing Then Set nodeCoord = nodePlacemark.SelectSingleNode(".//coordinates")
    
    If Not nodeCoord Is Nothing Then
        Call Imp_AdicionarPontoSGL(lo, nome, nodeCoord.Text)
    End If
End Sub

Private Sub Imp_AdicionarPontoSGL(lo As ListObject, nome As String, coordString As String)
    Dim partes() As String
    Dim lonDec As Double, latDec As Double, alt As Double
    Dim novaLinha As ListRow
    
    partes = Split(Trim(coordString), ",")
    
    If UBound(partes) >= 1 Then
        lonDec = Val(partes(0))
        latDec = Val(partes(1))
        If UBound(partes) >= 2 Then alt = Val(partes(2)) Else alt = 0
        
        Set novaLinha = lo.ListRows.Add
        With novaLinha
            .Range(1).Value = nome
            .Range(2).Value = M_Utils.Str_DD_Para_DMS(lonDec)
            .Range(3).Value = M_Utils.Str_DD_Para_DMS(latDec)
            .Range(4).Value = alt
            .Range(9).Value = "Cerca"
        End With
    End If
End Sub
