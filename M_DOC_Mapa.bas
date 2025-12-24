Attribute VB_Name = "M_DOC_Mapa"
Option Explicit

' ==============================================================================
' MÓDULO: M_DOC_MAPA
' DESCRIÇÃO: GERAÇÃO DE PLANTA TOPOGRÁFICA (MAPA 10X) VETORIAL E PDF A1
' ==============================================================================

' --- 1. CONSTANTES DE MAPEAMENTO (CÉLULAS NOMEADAS) ---

' Áreas Gráficas
Private Const RNG_COORDENADAS As String = "areaCoordenadas"      ' Colagem da Tabela
Private Const RNG_LOGO As String = "areaLogo"                    ' Logo da Empresa
Private Const RNG_CONVENCOES As String = "areaConvecoes"         ' Legenda/Convenções
Private Const RNG_ROSA As String = "areaRosa"                    ' Rosa dos Ventos
Private Const RNG_MAPA_LOCAL As String = "areaMapa"              ' Imagem Satélite/Croqui
Private Const RNG_POLIGONO As String = "imgPoligono"             ' Área do Gráfico Principal
Private Const RNG_REGUA As String = "areaRegua"                  ' Área da Régua de Escala

' Campos de Texto (Cabeçalho/Rodapé)
Private Const RNG_TITULO As String = "areaTitulo"
Private Const RNG_ESCALA As String = "areaEscala"
Private Const RNG_FOLHA As String = "areaFolha"
Private Const RNG_EMAIL_TEL As String = "areaEmailTelefone"

' Campos Técnicos (Datum/Observações)
Private Const RNG_OBS_COORD As String = "areaOBSCoordenadas"
Private Const RNG_OBS_DATUM As String = "areaOBSDatum"
Private Const RNG_OBS_MERID As String = "areaOBSMeridiano"
Private Const RNG_OBS_LEVANT As String = "areaOBSTipoLevantamento"

' Informações do Imóvel (Carimbo)
Private Const RNG_INFO_PROP As String = "areaINFOPropriedade"
Private Const RNG_INFO_DONO As String = "areaINFOProprietario"
Private Const RNG_INFO_MUN As String = "areaINFOMunicipio"
Private Const RNG_INFO_COMARCA As String = "areaINFOComarcas"
Private Const RNG_INFO_CARTORIO As String = "areaINFOCartorio"
Private Const RNG_INFO_MAT As String = "areaINFOMAT"
Private Const RNG_INFO_INCRA As String = "areaINFOCodIncra"
Private Const RNG_INFO_AREA As String = "areaINFOAreaTotal"
Private Const RNG_INFO_PERIM As String = "areaINFOPerimetro"
Private Const RNG_INFO_DATA As String = "areaINFOData"

' Assinaturas
Private Const RNG_ASSIN_DONO As String = "areaPROPNome"
Private Const RNG_ASSIN_RT As String = "areaRTNome"
Private Const RNG_ASSIN_FORMACAO As String = "areaRTFormacao"
Private Const RNG_ASSIN_CREA As String = "areaRTCrea"

' ==============================================================================
' 2. ROTINA PRINCIPAL (ORQUESTRA A GERAÇÃO DO MAPA)
' ==============================================================================
Public Sub GerarMapaExcel(dadosProp As Object, dadosTec As Object, _
                          titulo As String, escala1 As String, escala2 As String, _
                          email As String, telefone As String, _
                          txtCoord As String, txtDatum As String, _
                          txtMeridiano As String, txtLevantamento As String, _
                          pathLogo As String, pathMapaLocal As String, _
                          pathRosa As String, pathConvencoes As String)
    
    Dim wsMapa As Worksheet
    Dim loUTM As ListObject
    
    Call Utils_OtimizarPerformance(True)
    
    ' Referências
    Set wsMapa = ThisWorkbook.Sheets("MAPA10X")
    Set loUTM = ThisWorkbook.Sheets(M_Config.SH_UTM).ListObjects(M_Config.TBL_UTM)
    
    wsMapa.Activate
    
    ' --- A. PREENCHIMENTO DE TEXTOS (CARIMBO E OBSERVAÇÕES) ---
    On Error Resume Next
    
    ' 1. Cabeçalho e Escala
    wsMapa.Range(RNG_TITULO).Value = UCase(titulo)
    wsMapa.Range(RNG_ESCALA).Value = "1 / " & escala2
    wsMapa.Range(RNG_FOLHA).Value = "01"
    
    ' 2. Dados do Frame Mapa (E-mail, Telefone, Datum, Obs) - AGORA DINÂMICOS
    wsMapa.Range(RNG_EMAIL_TEL).Value = email & " | " & telefone
    wsMapa.Range(RNG_OBS_COORD).Value = txtCoord
    wsMapa.Range(RNG_OBS_DATUM).Value = txtDatum
    wsMapa.Range(RNG_OBS_MERID).Value = txtMeridiano
    wsMapa.Range(RNG_OBS_LEVANT).Value = txtLevantamento
    
    ' 3. Dados do Proprietário e Propriedade (Vindo do Dicionário dadosProp)
    ' Certifique-se que as Células Nomeadas existem na planilha MAPA10X
    wsMapa.Range(RNG_INFO_PROP).Value = dadosProp(M_Config.LBL_PROPRIEDADE)
    wsMapa.Range(RNG_INFO_DONO).Value = dadosProp(M_Config.LBL_PROP_NOME)
    wsMapa.Range(RNG_INFO_MUN).Value = dadosProp(M_Config.LBL_MUNICIPIO)
    wsMapa.Range(RNG_INFO_COMARCA).Value = dadosProp(M_Config.LBL_COMARCA)
    wsMapa.Range(RNG_INFO_CARTORIO).Value = dadosProp(M_Config.LBL_CARTORIO)
    wsMapa.Range(RNG_INFO_MAT).Value = dadosProp(M_Config.LBL_MATRICULA)
    wsMapa.Range(RNG_INFO_INCRA).Value = dadosProp(M_Config.LBL_INCRA)
    
    ' 4. Métricas Calculadas
    wsMapa.Range(RNG_INFO_AREA).Value = Format(dadosProp("Area (SGL)"), "0.0000") & " ha"
    wsMapa.Range(RNG_INFO_PERIM).Value = Format(ThisWorkbook.Sheets(M_Config.SH_UTM).Range(M_Config.CELL_UTM_PERIMETRO).Value, "0.00") & " m"
    wsMapa.Range(RNG_INFO_DATA).Value = Format(Date, "dd/mm/yyyy")
    
    ' 5. Assinaturas (Técnico e Proprietário)
    wsMapa.Range(RNG_ASSIN_DONO).Value = dadosProp(M_Config.LBL_PROP_NOME)
    wsMapa.Range(RNG_ASSIN_RT).Value = dadosTec(M_Config.LBL_RT_NOME)
    wsMapa.Range(RNG_ASSIN_FORMACAO).Value = dadosTec("Formação")
    wsMapa.Range(RNG_ASSIN_CREA).Value = dadosTec("Registro (CFT/CREA)")
    
    On Error GoTo 0
    
    ' --- B. INSERÇÃO DE IMAGENS ---
    Call InserirImagemNaCelula(wsMapa, pathLogo, RNG_LOGO)
    Call InserirImagemNaCelula(wsMapa, pathMapaLocal, RNG_MAPA_LOCAL)
    Call InserirImagemNaCelula(wsMapa, pathRosa, RNG_ROSA)
    Call InserirImagemNaCelula(wsMapa, pathConvencoes, RNG_CONVENCOES)
    
    ' --- C. GERAÇÃO DA TABELA UTM ---
    Call GerarTabelaComoImagem(wsMapa, loUTM)
    
    ' --- D. GERAÇÃO DO GRÁFICO VETORIAL ---
    Call InserirGraficoVivo(wsMapa, loUTM)
    
    ' Finalização
    wsMapa.Activate
    wsMapa.Range("A1").Select
    Call Utils_OtimizarPerformance(False)
    
    MsgBox "Mapa gerado com sucesso!", vbInformation
End Sub

' ==============================================================================
' 3. GRÁFICO VETORIAL "VIVO" (ESTILO COM EIXOS E GRADE)
' ==============================================================================
Private Sub InserirGraficoVivo(wsMapa As Worksheet, loUTM As ListObject)
    Dim rngDestino As Range
    Dim chtObj As ChartObject
    Dim cht As Chart
    Dim arrX() As Double, arrY() As Double
    Dim arrLabels() As String ' <--- Faltava declarar isso na sua versão
    Dim i As Long, qtd As Long
    Dim nomeGrafico As String
    Dim arrRaw As Variant
    Dim pt As Point
    
    nomeGrafico = "MapaPoligono" ' Nome fixo para facilitar a remoção depois
    
    ' 1. Define área de destino (Usa a constante do módulo)
    On Error Resume Next
    Set rngDestino = wsMapa.Range(RNG_POLIGONO).MergeArea
    If rngDestino Is Nothing Then Exit Sub
    On Error GoTo 0

    ' 2. Remove gráfico anterior se existir
    On Error Resume Next
    wsMapa.ChartObjects(nomeGrafico).Delete
    On Error GoTo 0

    ' 3. Prepara os Dados (UTM) com Fechamento
    If loUTM.ListRows.Count < 2 Then Exit Sub
    arrRaw = loUTM.DataBodyRange.Value
    qtd = UBound(arrRaw, 1)

    ReDim arrX(1 To qtd + 1)
    ReDim arrY(1 To qtd + 1)
    ReDim arrLabels(1 To qtd + 1)

    For i = 1 To qtd
        arrLabels(i) = CStr(arrRaw(i, 1)) ' Nome do Ponto
        arrY(i) = CDbl(arrRaw(i, 2))      ' Norte (Y)
        arrX(i) = CDbl(arrRaw(i, 3))      ' Este (X)
    Next i
    
    ' Fechamento do Polígono (Repete o primeiro ponto no final)
    arrY(qtd + 1) = arrY(1)
    arrX(qtd + 1) = arrX(1)
    arrLabels(qtd + 1) = arrLabels(1)

    ' 4. Cria o Objeto Gráfico
    Set chtObj = wsMapa.ChartObjects.Add(Left:=rngDestino.Left, _
                                         Top:=rngDestino.Top, _
                                         Width:=rngDestino.Width, _
                                         Height:=rngDestino.Height)
    chtObj.Name = nomeGrafico
    Set cht = chtObj.Chart

    ' 5. Configura a Série e o Visual
    With cht
        .ChartType = xlXYScatterLines
        
        ' Limpa séries automáticas
        Do While .SeriesCollection.Count > 0
            .SeriesCollection(1).Delete
        Loop

        ' Adiciona a Série do Perímetro
        With .SeriesCollection.NewSeries
            .Name = "Perímetro"
            .XValues = arrX
            .Values = arrY
            
            ' Estilo da Linha (Azul Técnico)
            .Format.Line.ForeColor.RGB = RGB(0, 176, 240)
            .Format.Line.Weight = 2
            
            ' Estilo dos Marcadores (Vértices)
            .MarkerStyle = xlMarkerStyleCircle
            .MarkerSize = 5
            .MarkerBackgroundColor = RGB(0, 32, 96)
            .MarkerForegroundColor = RGB(0, 32, 96)
            
            ' Rótulos dos Pontos (P-1, P-2...)
            .HasDataLabels = True
            .DataLabels.Position = xlLabelPositionAbove

            For i = 1 To UBound(arrLabels)
                Set pt = .Points(i)
                pt.DataLabel.Text = arrLabels(i)
                pt.DataLabel.Font.Name = "Arial"
                pt.DataLabel.Font.Size = 7
                pt.DataLabel.Font.Color = RGB(0, 0, 0)
                pt.DataLabel.Font.Bold = True
            Next i
        End With

        ' 6. Limpeza Geral
        .HasTitle = False
        .HasLegend = False
        .ChartArea.Border.LineStyle = xlNone
        .ChartArea.Format.Fill.Visible = msoFalse
        .PlotArea.Format.Fill.Visible = msoFalse

        ' --- 7. CONFIGURAÇÃO DOS EIXOS (COM GRADE) ---
        
        ' Eixo Y (Norte)
        .HasAxis(xlValue) = True
        With .Axes(xlValue)
            .Border.Color = RGB(150, 150, 150)
            .HasMajorGridlines = True
            .MajorGridlines.Border.Color = RGB(220, 220, 220) ' Grade Cinza Claro
            .TickLabels.Font.Size = 8
            .TickLabels.NumberFormat = "#,##0"
        End With

        ' Eixo X (Este)
        .HasAxis(xlCategory) = True
        With .Axes(xlCategory)
            .Border.Color = RGB(150, 150, 150)
            .HasMajorGridlines = True
            .MajorGridlines.Border.Color = RGB(220, 220, 220)
            .TickLabels.Font.Size = 8
            .TickLabels.NumberFormat = "#,##0"
        End With
    End With

    ' 8. Ajuste de Escala (1:1 - Sem distorção)
    Call AjustarEscalaGraficoMapa(cht, arrX, arrY)
    
    ' 9. Gera a Régua (Baseado na escala calculada)
    ' A régua precisa saber quantos metros cabem na largura do gráfico
    Dim larguraPlotMetros As Double
    larguraPlotMetros = cht.Axes(xlCategory).MaximumScale - cht.Axes(xlCategory).MinimumScale
    Call DesenharReguaVetorial(wsMapa, larguraPlotMetros)

End Sub

' ==============================================================================
' 4. AJUSTE DE ESCALA (MANTÉM PROPORÇÃO 1:1 EM GRÁFICO COM EIXOS)
' ==============================================================================
Private Sub AjustarEscalaGraficoMapa(cht As Chart, arrX() As Double, arrY() As Double)
    Dim minX As Double, maxX As Double, minY As Double, maxY As Double
    Dim deltaX As Double, deltaY As Double, centroX As Double, centroY As Double
    Dim plotW As Double, plotH As Double, novoDelta As Double
    Dim i As Long

    ' 1. Encontra limites dos dados
    minX = arrX(1): maxX = arrX(1): minY = arrY(1): maxY = arrY(1)
    For i = LBound(arrX) To UBound(arrX)
        If arrX(i) < minX Then minX = arrX(i)
        If arrX(i) > maxX Then maxX = arrX(i)
        If arrY(i) < minY Then minY = arrY(i)
        If arrY(i) > maxY Then maxY = arrY(i)
    Next i

    ' 2. Adiciona margem de 10%
    deltaX = (maxX - minX) * 1.1
    deltaY = (maxY - minY) * 1.1
    
    ' Evita erro se houver apenas 1 ponto ou pontos iguais
    If deltaX < 50 Then deltaX = 50
    If deltaY < 50 Then deltaY = 50
    
    centroX = (minX + maxX) / 2
    centroY = (minY + maxY) / 2

    ' 3. Pega dimensões físicas do gráfico (em pontos)
    plotW = cht.Parent.Width
    plotH = cht.Parent.Height

    ' Reseta escalas automáticas antes de calcular
    cht.Axes(xlCategory).MinimumScaleIsAuto = True
    cht.Axes(xlCategory).MaximumScaleIsAuto = True
    cht.Axes(xlValue).MinimumScaleIsAuto = True
    cht.Axes(xlValue).MaximumScaleIsAuto = True

    ' 4. Matemática de Proporção (Aspect Ratio 1:1)
    ' Se o gráfico é mais "largo" que os dados, aumentamos a escala X.
    ' Se o gráfico é mais "alto" que os dados, aumentamos a escala Y.
    If (deltaX / plotW) > (deltaY / plotH) Then
        ' Os dados são muito largos. O fator limitante é o X.
        ' Calculamos quanto Y precisamos mostrar para manter a proporção.
        novoDelta = deltaX * (plotH / plotW)
        
        cht.Axes(xlCategory).MinimumScale = centroX - (deltaX / 2)
        cht.Axes(xlCategory).MaximumScale = centroX + (deltaX / 2)
        
        cht.Axes(xlValue).MinimumScale = centroY - (novoDelta / 2)
        cht.Axes(xlValue).MaximumScale = centroY + (novoDelta / 2)
    Else
        ' Os dados são muito altos. O fator limitante é o Y.
        novoDelta = deltaY * (plotW / plotH)
        
        cht.Axes(xlValue).MinimumScale = centroY - (deltaY / 2)
        cht.Axes(xlValue).MaximumScale = centroY + (deltaY / 2)
        
        cht.Axes(xlCategory).MinimumScale = centroX - (novoDelta / 2)
        cht.Axes(xlCategory).MaximumScale = centroX + (novoDelta / 2)
    End If
End Sub

' ==============================================================================
' 4. AJUSTE DE ESCALA (IMPEDE DISTORÇÃO)
' ==============================================================================
'Private Function AjustarGraficoProporcional0(cht As Chart, arrX As Variant, arrY As Variant) As Double
'    Dim minX As Double, maxX As Double, minY As Double, maxY As Double
'    Dim deltaX As Double, deltaY As Double
'    Dim centroX As Double, centroY As Double
'    Dim maxDelta As Double
'    Dim margem As Double
'
'    ' Encontra Min/Max
'    minX = Application.Min(arrX): maxX = Application.Max(arrX)
'    minY = Application.Min(arrY): maxY = Application.Max(arrY)
'
'    deltaX = maxX - minX
'    deltaY = maxY - minY
'    centroX = (minX + maxX) / 2
'    centroY = (minY + maxY) / 2
'
'    ' Define o "bounding box" quadrado para manter proporção
'    If deltaX > deltaY Then maxDelta = deltaX Else maxDelta = deltaY
'
'    margem = maxDelta * 0.1 ' 10% de margem
'    maxDelta = maxDelta + (margem * 2)
'
'    ' Aplica aos eixos (mesmo invisíveis, eles controlam a escala)
'    With cht.Axes(xlCategory) ' X
'        .MinimumScale = centroX - (maxDelta / 2)
'        .MaximumScale = centroX + (maxDelta / 2)
'    End With
'
'    With cht.Axes(xlValue) ' Y
'        .MinimumScale = centroY - (maxDelta / 2)
'        .MaximumScale = centroY + (maxDelta / 2)
'    End With
'
'    ' Retorna quantos METROS cabem na largura total do gráfico
'    AjustarGraficoProporcional = maxDelta
'End Function

' ==============================================================================
' 5. RÉGUA DE ESCALA VETORIAL
' ==============================================================================
Private Sub DesenharReguaVetorial(ws As Worksheet, maxMetrosNoGrafico As Double)
    Dim rngRegua As Range
    Dim shp As Shape
    Dim larguraReguaPx As Double
    Dim metrosPorPx As Double
    Dim tamanhoReguaMetros As Double
    Dim larguraVisualPx As Double
    
    On Error Resume Next
    Set rngRegua = ws.Range(RNG_REGUA)
    ws.Shapes("ReguaEscala").Delete
    On Error GoTo 0
    
    ' Calcula proporção
    larguraReguaPx = ws.Range(RNG_POLIGONO).Width ' Largura do gráfico em pontos
    metrosPorPx = maxMetrosNoGrafico / larguraReguaPx
    
    ' Define um tamanho "bonito" para a régua (50m, 100m, 200m, 500m...)
    Dim baseRegua As Double
    baseRegua = (rngRegua.Width * metrosPorPx) * 0.8 ' Tenta usar 80% da área disponível
    
    ' Define um tamanho "bonito" para a régua
    If baseRegua > 1000 Then
        tamanhoReguaMetros = 1000
    ElseIf baseRegua > 500 Then
        tamanhoReguaMetros = 500
    ElseIf baseRegua > 200 Then
        tamanhoReguaMetros = 200
    ElseIf baseRegua > 100 Then
        tamanhoReguaMetros = 100
    Else
        tamanhoReguaMetros = 50
    End If
    
    larguraVisualPx = tamanhoReguaMetros / metrosPorPx
    
    ' Desenha a Linha
    Set shp = ws.Shapes.AddLine(rngRegua.Left, rngRegua.Top + (rngRegua.Height / 2), _
                                rngRegua.Left + larguraVisualPx, rngRegua.Top + (rngRegua.Height / 2))
    shp.Name = "ReguaEscala"
    shp.Line.Weight = 2
    shp.Line.ForeColor.RGB = RGB(0, 0, 0)
    shp.Line.EndArrowheadStyle = msoArrowheadTriangle
    shp.Line.BeginArrowheadStyle = msoArrowheadTriangle
    
    ' Adiciona o Texto (Ex: "100 m") no meio
    Dim txt As Shape
    On Error Resume Next
    ws.Shapes("TextoRegua").Delete
    On Error GoTo 0
    
    Set txt = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, _
                                   rngRegua.Left, rngRegua.Top, _
                                   larguraVisualPx, rngRegua.Height / 2)
    txt.Name = "TextoRegua"
    txt.TextFrame2.TextRange.Text = tamanhoReguaMetros & " m"
    txt.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    txt.Fill.Visible = msoFalse
    txt.Line.Visible = msoFalse
End Sub

' ==============================================================================
' 6. UTILITÁRIOS (IMAGENS E TABELA)
' ==============================================================================
Private Sub InserirImagemNaCelula(ws As Worksheet, caminhoImg As String, nomeRange As String)
    If caminhoImg = "" Or Dir(caminhoImg) = "" Then Exit Sub
    
    Dim rng As Range, shp As Shape
    On Error Resume Next
    Set rng = ws.Range(nomeRange).MergeArea
    
    ' Limpa anteriores na área
    For Each shp In ws.Shapes
        If Not Intersect(shp.TopLeftCell, rng) Is Nothing And shp.Type = msoPicture Then
            shp.Delete
        End If
    Next shp
    On Error GoTo 0
    
    Set shp = ws.Shapes.AddPicture(caminhoImg, msoFalse, msoTrue, rng.Left, rng.Top, -1, -1)
    
    ' Ajuste Best Fit
    With shp
        .LockAspectRatio = msoTrue
        If (.Width / .Height) > (rng.Width / rng.Height) Then
            .Width = rng.Width - 4
            .Top = rng.Top + (rng.Height - .Height) / 2
            .Left = rng.Left + 2
        Else
            .Height = rng.Height - 4
            .Top = rng.Top + 2
            .Left = rng.Left + (rng.Width - .Width) / 2
        End If
    End With
End Sub

' ==============================================================================
' GERA A IMAGEM DA TABELA (AJUSTE FINO: LARGURA TOTAL x LIMITE ALTURA)
' ==============================================================================
Private Sub GerarTabelaComoImagem(wsMapa As Worksheet, loUTM As ListObject)
    Dim wsTemp As Worksheet
    Dim rngDados As Range, rngDestino As Range
    Dim shp As Shape
    Dim i As Long, qtd As Long
    Dim arrDados As Variant
    Dim alturaMax As Double, larguraDestino As Double
    
    ' 1. Prepara a aba temporária
    On Error Resume Next
    Set wsTemp = ThisWorkbook.Sheets("TEMP_MAPA")
    If wsTemp Is Nothing Then
        Set wsTemp = ThisWorkbook.Sheets.Add
        wsTemp.Name = "TEMP_MAPA"
        wsTemp.Visible = xlSheetVeryHidden
    End If
    On Error GoTo 0
    wsTemp.Cells.Clear

    ' 2. Cabeçalho
    wsTemp.Range("A1:F1").Value = Array("PONTO", "PARA", "COORD. N (Y)", "COORD. E (X)", "AZIMUTE", "DISTÂNCIA")
    With wsTemp.Range("A1:F1")
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Interior.Color = RGB(220, 220, 220)
    End With

    ' 3. Preenche Dados
    If loUTM.ListRows.Count > 0 Then
        qtd = loUTM.ListRows.Count
        arrDados = loUTM.DataBodyRange.Value

        For i = 1 To qtd
            wsTemp.Cells(i + 1, 1).Value = arrDados(i, 1)
            wsTemp.Cells(i + 1, 2).Value = arrDados(i, 5)
            wsTemp.Cells(i + 1, 3).Value = Format(arrDados(i, 2), "0.000")
            wsTemp.Cells(i + 1, 4).Value = Format(arrDados(i, 3), "0.000")
            
            ' Azimute
            If IsNumeric(arrDados(i, 6)) Then
                wsTemp.Cells(i + 1, 5).Value = M_Utils.Str_FormatAzimute(CDbl(arrDados(i, 6)))
            Else
                wsTemp.Cells(i + 1, 5).Value = arrDados(i, 6)
            End If
            
            wsTemp.Cells(i + 1, 6).Value = Format(arrDados(i, 7), "0.00")
        Next i

        ' Formatação Visual da Tabela (Fonte menor ajuda a caber)
        Set rngDados = wsTemp.Range("A1:F" & qtd + 1)
        With rngDados
            .Borders.LineStyle = xlContinuous
            .HorizontalAlignment = xlCenter
            .Font.Name = "Arial Narrow" ' Fonte mais estreita ajuda
            .Font.Size = 8
            .Columns.AutoFit
        End With

        ' 4. Copia
        rngDados.CopyPicture Appearance:=xlScreen, Format:=xlPicture

        ' 5. Define Destino (CORREÇÃO DA LARGURA)
        wsMapa.Activate
        
        On Error Resume Next
        ' Pega a área mesclada completa para saber a largura real
        Set rngDestino = wsMapa.Range(RNG_COORDENADAS).MergeArea
        If rngDestino Is Nothing Then Set rngDestino = wsMapa.Range("B2:K90") ' Fallback de segurança
        On Error GoTo 0
        
        ' Remove anterior
        For Each shp In wsMapa.Shapes
            If Not Intersect(shp.TopLeftCell, rngDestino) Is Nothing And shp.Type = msoPicture Then
                shp.Delete
            End If
        Next shp

        ' Seleciona apenas a primeira célula para colar (evita erro de mesclagem)
        rngDestino.Cells(1, 1).Select
        wsMapa.Paste
        
        Set shp = wsMapa.Shapes(wsMapa.Shapes.Count)
        
        ' --- 6. LÓGICA DE DIMENSIONAMENTO (PRIORIDADE LARGURA) ---
        alturaMax = rngDestino.Height
        larguraDestino = rngDestino.Width
        
        With shp
            .LockAspectRatio = msoTrue
            .Top = rngDestino.Top
            .Left = rngDestino.Left
            
            ' APLICA A LARGURA TOTAL DA ÁREA AZUL
            .Width = larguraDestino
            
            ' VERIFICA SE ESTOUROU A ALTURA
            ' (Se a tabela for muito comprida, ela vai passar do limite inferior ao ser esticada)
            If .Height > alturaMax Then
                .Height = alturaMax
                ' Ao reduzir a altura, a largura reduzirá automaticamente (proporcional)
                ' criando espaço branco nas laterais, o que é inevitável para não cortar dados.
            End If
        End With
    End If

    Application.CutCopyMode = False
End Sub


' ==============================================================================
' 7. PDF A1 (VERSÃO BLINDADA)
' ==============================================================================
Public Sub GerarPDFMapa(dadosProp As Object, Optional Silencioso As Boolean = False)
    Dim wsMapa As Worksheet
    Dim nomeArquivo As String, caminhoArquivo As String, pastaSalvar As String
    
    Set wsMapa = ThisWorkbook.Sheets("MAPA10X")
    
    Call Utils_OtimizarPerformance(True)
    
    nomeArquivo = "MAPA_A1_" & M_Utils.File_SanitizeName(dadosProp(M_Config.LBL_PROPRIEDADE)) & "_" & Format(Date, "dd-mm-yy") & ".pdf"
    
    If Not Silencioso Then
        pastaSalvar = M_Utils.UI_SelecionarPasta()
        If pastaSalvar = "" Then
            Call Utils_OtimizarPerformance(False)
            Exit Sub
        End If
    Else
        pastaSalvar = ThisWorkbook.path & "\"
    End If
    caminhoArquivo = pastaSalvar & nomeArquivo
    
    With wsMapa.PageSetup
        .PrintArea = "$A$1:$X$97"
        .Orientation = xlLandscape
        
        ' Tenta A1, ignora erro se não tiver driver
        On Error Resume Next
        .PaperSize = 144
        If Err.Number <> 0 Then
            Err.Clear: .PaperSize = xlPaperA3
            If Err.Number <> 0 Then .PaperSize = xlPaperA4
        End If
        On Error GoTo 0
        
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .LeftMargin = 0: .RightMargin = 0: .TopMargin = 0: .BottomMargin = 0
        .CenterHorizontally = True: .CenterVertically = True
    End With
    
    On Error GoTo ErroPDF
    wsMapa.ExportAsFixedFormat Type:=xlTypePDF, filename:=caminhoArquivo, _
                               Quality:=xlQualityStandard, IncludeDocProperties:=True, _
                               IgnorePrintAreas:=False, OpenAfterPublish:=False
    
    Call Utils_OtimizarPerformance(False)
    
    If Not Silencioso Then
        If MsgBox("Mapa gerado!" & vbCrLf & "Deseja abrir?", vbYesNo + vbInformation) = vbYes Then
            ThisWorkbook.FollowHyperlink caminhoArquivo
        End If
    End If
    Exit Sub

ErroPDF:
    Call Utils_OtimizarPerformance(False)
    If Not Silencioso Then MsgBox "Erro PDF: " & Err.Description, vbCritical
End Sub

' ==============================================================================
' BOTÃO DA PLANILHA: GERAR MAPA RÁPIDO (SEM USERFORM)
' ==============================================================================
Public Sub Btn_GerarMapa_Direto()
    Dim dadosProp As Object, dadosTec As Object
    
    ' 1. Cria dicionários vazios para evitar erro de objeto
    ' (Como não estamos no UserForm, não temos os dados carregados)
    Set dadosProp = CreateObject("Scripting.Dictionary")
    Set dadosTec = CreateObject("Scripting.Dictionary")
    
    ' Dica: Se quiser preencher algo padrão, faça assim:
    ' dadosProp(M_Config.LBL_PROPRIEDADE) = "MAPA RÁPIDO"
    
    ' 2. Chama a rotina de geração passando tudo VAZIO ("")
    ' O desenho (polígono) e a tabela de coordenadas serão gerados
    ' pois eles vêm da tabela UTM, e não do formulário.
    Call M_DOC_Mapa.GerarMapaExcel(dadosProp, _
                                   dadosTec, _
                                   "", _
                                   "", _
                                   "", _
                                   "", _
                                   "", _
                                   "", _
                                   "", _
                                   "", _
                                   "", _
                                   "", _
                                   "", _
                                   "", _
                                   "")
                                   
    MsgBox "Mapa gerado! (Modo Rápido - Sem Carimbo)", vbInformation
End Sub

