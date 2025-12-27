Attribute VB_Name = "M_App_Logica"
Option Explicit
' ==============================================================================
' MODULO: M_APP_LOGICA
' DESCRICAO: REGRAS DE NEGOCIO COM PROTECAO DE PLANILHAS
' ==============================================================================

Public Sub Processo_AtualizarMetricas()
    Dim wsPainel As Worksheet, wsSGL As Worksheet, wsUTM As Worksheet
    Dim loSGL As ListObject, loUTM As ListObject
    Dim sistemaAtivo As String
    Dim areaHaUTM As Double, areaM2UTM As Double, perimetroUTM As Double
    Dim areaHaSGL As Double, areaM2SGL As Double, perimetroSGL As Double
    
    On Error GoTo ErroCalculo
    
    sistemaAtivo = M_Config.App_GetSistemaAtivo()
    Set wsPainel = ThisWorkbook.Sheets(M_Config.SH_PAINEL)
    Set wsSGL = ThisWorkbook.Sheets(M_Config.SH_SGL)
    Set wsUTM = ThisWorkbook.Sheets(M_Config.SH_UTM)
    
    areaHaUTM = 0: areaM2UTM = 0: perimetroUTM = 0
    areaHaSGL = 0: areaM2SGL = 0: perimetroSGL = 0
    
    ' Calcular Area UTM
    On Error Resume Next
    Set loUTM = wsUTM.ListObjects(M_Config.TBL_UTM)
    If loUTM.ListRows.Count >= 3 Then
        Dim arrE2 As Variant, arrN2 As Variant
        arrN2 = loUTM.ListColumns(2).DataBodyRange.Value
        arrE2 = loUTM.ListColumns(3).DataBodyRange.Value
        arrN2 = Application.Transpose(arrN2)
        arrE2 = Application.Transpose(arrE2)
        areaM2UTM = M_Math_Geo.Geo_Area_Gauss(arrE2, arrN2)
        areaHaUTM = areaM2UTM / 10000
        perimetroUTM = Application.WorksheetFunction.Sum(loUTM.ListColumns(7).DataBodyRange)
    End If
    On Error GoTo ErroCalculo
    
    ' Calcular Area SGL
    On Error Resume Next
    Set loSGL = wsSGL.ListObjects(M_Config.TBL_SGL)
    If loSGL.ListRows.Count >= 3 Then
        Call Processo_Calc_Area_SGL_Avancado(loSGL, areaM2SGL, areaHaSGL)
        perimetroSGL = Application.WorksheetFunction.Sum(loSGL.ListColumns(7).DataBodyRange)
    End If
    On Error GoTo ErroCalculo
    
    ' Desbloquear planilhas
    M_SheetProtection.DesbloquearPlanilha wsSGL
    M_SheetProtection.DesbloquearPlanilha wsUTM
    M_SheetProtection.DesbloquearPlanilha wsPainel
    
    ' Gravar nas celulas das planilhas de dados
    EscreverCelulaSegura wsSGL, M_Config.CELL_SGL_AREA_HA, areaHaSGL
    EscreverCelulaSegura wsSGL, M_Config.CELL_SGL_AREA_M2, areaM2SGL
    EscreverCelulaSegura wsSGL, M_Config.CELL_SGL_PERIMETRO, perimetroSGL
    
    EscreverCelulaSegura wsUTM, M_Config.CELL_UTM_AREA_HA, areaHaUTM
    EscreverCelulaSegura wsUTM, M_Config.CELL_UTM_AREA_M2, areaM2UTM
    EscreverCelulaSegura wsUTM, M_Config.CELL_UTM_PERIMETRO, perimetroUTM
    
    ' Atualizar shapes - PAINEL_PRINCIPAL
    AtualizarTextoShape wsPainel, "shp_Valor_Ha_SGL", Format(areaHaSGL, "0.0000") & " ha"
    AtualizarTextoShape wsPainel, "shp_Valor_Ha_UTM", Format(areaHaUTM, "0.0000") & " ha"
    If sistemaAtivo = "SGL" Then
        AtualizarTextoShape wsPainel, "shp_Valor_M2", Format(areaM2SGL, "#,##0.00") & " m2"
        AtualizarTextoShape wsPainel, "shp_Valor_Perimetro", Format(perimetroSGL, "#,##0.00") & " m"
    Else
        AtualizarTextoShape wsPainel, "shp_Valor_M2", Format(areaM2UTM, "#,##0.00") & " m2"
        AtualizarTextoShape wsPainel, "shp_Valor_Perimetro", Format(perimetroUTM, "#,##0.00") & " m"
    End If
    
    ' Atualizar shapes - DADOS_PRINCIPAL_SGL
    AtualizarTextoShape wsSGL, "shp_Valor_Ha_SGL", Format(areaHaSGL, "0.0000") & " ha"
    AtualizarTextoShape wsSGL, "shp_Valor_Ha_UTM", Format(areaHaUTM, "0.0000") & " ha"
    AtualizarTextoShape wsSGL, "shp_Valor_M2", Format(areaM2SGL, "#,##0.00") & " m2"
    AtualizarTextoShape wsSGL, "shp_Valor_Perimetro", Format(perimetroSGL, "#,##0.00") & " m"
    
    ' Atualizar shapes - DADOS_PRINCIPAL_UTM
    AtualizarTextoShape wsUTM, "shp_Valor_Ha_SGL", Format(areaHaSGL, "0.0000") & " ha"
    AtualizarTextoShape wsUTM, "shp_Valor_Ha_UTM", Format(areaHaUTM, "0.0000") & " ha"
    AtualizarTextoShape wsUTM, "shp_Valor_M2", Format(areaM2UTM, "#,##0.00") & " m2"
    AtualizarTextoShape wsUTM, "shp_Valor_Perimetro", Format(perimetroUTM, "#,##0.00") & " m"
    
    ' Bloquear planilhas
    M_SheetProtection.BloquearPlanilha wsSGL
    M_SheetProtection.BloquearPlanilha wsUTM
    M_SheetProtection.BloquearPlanilha wsPainel
    Exit Sub
    
ErroCalculo:
    MsgBox "Erro ao atualizar metricas: " & Err.Description, vbExclamation
End Sub

Private Sub AtualizarTextoShape(ws As Worksheet, nomeShape As String, texto As String)
    On Error Resume Next
    ws.Shapes(nomeShape).TextFrame2.TextRange.Text = texto
    On Error GoTo 0
End Sub

Private Sub Processo_Calc_Area_SGL_Avancado(lo As ListObject, ByRef outM2 As Double, ByRef outHa As Double)
    Dim i As Long, qtd As Long
    Dim latSoma As Double, lonSoma As Double, altSoma As Double
    Dim lat0 As Double, lon0 As Double, alt0 As Double
    Dim arrDados As Variant
    Dim E_sgl() As Double, N_sgl() As Double
    
    arrDados = lo.DataBodyRange.Value
    qtd = UBound(arrDados, 1)
    
    For i = 1 To qtd
        latSoma = latSoma + M_Utils.Str_DMS_Para_DD(CStr(arrDados(i, 3)))
        lonSoma = lonSoma + M_Utils.Str_DMS_Para_DD(CStr(arrDados(i, 2)))
        If IsNumeric(arrDados(i, 4)) Then altSoma = altSoma + CDbl(arrDados(i, 4))
    Next i
    
    lat0 = latSoma / qtd
    lon0 = lonSoma / qtd
    alt0 = altSoma / qtd
    
    Dim ptOrigem As Type_Geocentrica
    ptOrigem = M_Math_Geo.Geo_Geod_Para_Geoc(lon0, lat0, alt0)
    
    ReDim E_sgl(1 To qtd)
    ReDim N_sgl(1 To qtd)
    
    For i = 1 To qtd
        Dim latPt As Double, lonPt As Double, altPt As Double
        lonPt = M_Utils.Str_DMS_Para_DD(CStr(arrDados(i, 2)))
        latPt = M_Utils.Str_DMS_Para_DD(CStr(arrDados(i, 3)))
        If IsNumeric(arrDados(i, 4)) Then
            altPt = CDbl(arrDados(i, 4))
        Else
            altPt = 0
        End If
        
        Dim ptGeoc As Type_Geocentrica
        ptGeoc = M_Math_Geo.Geo_Geod_Para_Geoc(lonPt, latPt, altPt)
        
        Dim ptTopo As Type_Topocentrica
        ptTopo = M_Math_Geo.Geo_Geoc_Para_Topoc(ptGeoc.x, ptGeoc.y, ptGeoc.Z, lon0, lat0, ptOrigem.x, ptOrigem.y, ptOrigem.Z)
        
        E_sgl(i) = ptTopo.E
        N_sgl(i) = ptTopo.N
    Next i
    
    outM2 = M_Math_Geo.Geo_Area_Gauss(E_sgl, N_sgl)
    outHa = outM2 / 10000
End Sub

Public Sub Processo_Conv_SGL_UTM()
    Dim wsSGL As Worksheet, wsUTM As Worksheet
    Dim loSGL As ListObject, loUTM As ListObject
    Dim i As Long, qtd As Long
    Dim arrSGL As Variant

    On Error Resume Next
    Set wsSGL = ThisWorkbook.Sheets(M_Config.SH_SGL)
    Set wsUTM = ThisWorkbook.Sheets(M_Config.SH_UTM)
    Set loSGL = wsSGL.ListObjects(M_Config.TBL_SGL)
    Set loUTM = wsUTM.ListObjects(M_Config.TBL_UTM)
    On Error GoTo 0

    If loSGL Is Nothing Or loUTM Is Nothing Then Exit Sub
    If loSGL.ListRows.Count = 0 Then Exit Sub

    Call M_Utils.Utils_OtimizarPerformance(True)
    M_SheetProtection.DesbloquearPlanilha wsUTM

    Call M_Dados.Dados_LimparTabela(M_Config.SH_UTM, M_Config.TBL_UTM)

    arrSGL = loSGL.DataBodyRange.Value
    qtd = UBound(arrSGL, 1)

    Dim linhasAtuais As Long: linhasAtuais = loUTM.ListRows.Count
    If linhasAtuais < qtd Then
        For i = 1 To qtd - linhasAtuais: loUTM.ListRows.Add: Next i
    End If

    Dim arrOut() As Variant
    ReDim arrOut(1 To qtd, 1 To loUTM.ListColumns.Count)

    Dim latDD As Double, lonDD As Double
    Dim utmAtual As Type_UTM
    Dim zonaPadrao As Integer

    ' Detecta fuso automaticamente da primeira coordenada
    lonDD = M_Utils.Str_DMS_Para_DD(CStr(arrSGL(1, 2)))
    zonaPadrao = M_Math_Geo.Geo_GetZonaUTM(lonDD)

    Dim cacheN() As Double, cacheE() As Double
    ReDim cacheN(1 To qtd), cacheE(1 To qtd)

    ' Primeira passagem: Converte todas as coordenadas Geo → UTM
    For i = 1 To qtd
        lonDD = M_Utils.Str_DMS_Para_DD(CStr(arrSGL(i, 2)))
        latDD = M_Utils.Str_DMS_Para_DD(CStr(arrSGL(i, 3)))

        ' USA FUNÇÃO REFATORADA: Converter_GeoParaUTM (com 3 parâmetros)
        utmAtual = M_Math_Geo.Converter_GeoParaUTM(latDD, lonDD, zonaPadrao)

        cacheN(i) = utmAtual.Norte
        cacheE(i) = utmAtual.Leste

        arrOut(i, 1) = arrSGL(i, 1)
        arrOut(i, 2) = Round(utmAtual.Norte, 2)  ' 2 casas decimais
        arrOut(i, 3) = Round(utmAtual.Leste, 2)
        arrOut(i, 4) = arrSGL(i, 4)
        arrOut(i, 8) = arrSGL(i, 8)
        arrOut(i, 9) = arrSGL(i, 9)
        arrOut(i, 10) = arrSGL(i, 10)
    Next i

    ' Segunda passagem: Calcula azimute e distância entre pontos consecutivos
    For i = 1 To qtd
        Dim idxProx As Long
        If i < qtd Then
            idxProx = i + 1
            arrOut(i, 5) = arrSGL(i + 1, 1)
        Else
            idxProx = 1
            arrOut(i, 5) = arrSGL(1, 1)
        End If

        ' USA FUNÇÃO REFATORADA: Calcular_DistanciaAzimute_UTM (calcula tudo de uma vez)
        Dim calc As Type_CalculoPonto
        calc = M_Math_Geo.Calcular_DistanciaAzimute_UTM(cacheN(i), cacheE(i), cacheN(idxProx), cacheE(idxProx))

        ' NOVO: Aplica correção de Convergência Meridiana para obter Azimute Geodésico
        ' Azimute Geodésico = Azimute de Grid + Convergência Meridiana
        Dim azimuteGeod As Double
        lonDD = M_Utils.Str_DMS_Para_DD(CStr(arrSGL(i, 2)))
        latDD = M_Utils.Str_DMS_Para_DD(CStr(arrSGL(i, 3)))
        azimuteGeod = M_Math_Geo.Converter_AzimuteGridParaGeod(calc.AzimuteDecimal, latDD, lonDD, zonaPadrao)

        ' USA NOVA FUNÇÃO: Str_FormatAzimuteGMS (com segundos: GGG°MM'SS")
        arrOut(i, 6) = M_Utils.Str_FormatAzimuteGMS(azimuteGeod)
        arrOut(i, 7) = Round(calc.Distancia, 3)
    Next i

    loUTM.DataBodyRange.Value = arrOut

    M_SheetProtection.BloquearPlanilha wsUTM
    Call Processo_AtualizarMetricas
    Call M_Utils.Utils_OtimizarPerformance(False)
End Sub

Public Sub Processo_PosImportacao()
    Dim ws As Worksheet, lo As ListObject
    Set ws = ThisWorkbook.Sheets(M_Config.App_GetNomeAbaAtiva())
    Set lo = ws.ListObjects(M_Config.App_GetNomeTabelaAtiva())

    If lo.ListRows.Count = 0 Then Exit Sub

    Call M_Utils.Utils_OtimizarPerformance(True)
    M_SheetProtection.DesbloquearPlanilha ws

    Call M_UI_Main.UI_DetectarFusoHemisferio

    lo.ListColumns(4).DataBodyRange.NumberFormat = "0.00"
    lo.ListColumns(7).DataBodyRange.NumberFormat = "0.000"

    Dim formulaDesc As String
    formulaDesc = "=IFERROR(VLOOKUP(TRIM([@Tipo])," & M_Config.TBL_PARAMETROS & ",2,FALSE), ""--"")"
    On Error Resume Next
    lo.ListColumns(10).DataBodyRange.Formula = formulaDesc
    On Error GoTo 0

    ' NOVO: Preenche valores padrao para campos de validacao INCRA (se existirem)
    Call PreencherValoresPadraoINCRA(lo)

    M_SheetProtection.BloquearPlanilha ws

    Call Processo_AtualizarMetricas
    Call Processo_Conv_SGL_UTM
    Call M_UI_Main.UI_Resize_ListBox
    Call M_UI_Main.UI_Refresh_ListBox
    Call M_Graficos.Grafico_PlotarPoligono(M_Config.SH_PAINEL)
    Call M_Graficos.Grafico_PlotarPoligono(M_Config.SH_CROQUI)

    Call M_Utils.Utils_OtimizarPerformance(False)
End Sub

Private Sub EscreverCelulaSegura(ws As Worksheet, EnderecoOuNome As String, valor As Variant)
    On Error Resume Next
    ws.Range(EnderecoOuNome).Value = valor
    On Error GoTo 0
End Sub

Private Sub PreencherValoresPadraoINCRA(lo As ListObject)
    '----------------------------------------------------------------------------------
    ' Preenche valores padrao para campos de validacao INCRA apos importacao
    ' Se os campos nao existirem, a funcao simplesmente retorna sem erro
    '----------------------------------------------------------------------------------
    Dim colPrecisaoH As ListColumn, colPrecisaoV As ListColumn
    Dim colMetodo As ListColumn, colCodLimite As ListColumn
    Dim i As Long

    On Error Resume Next

    ' Tenta localizar as colunas de validacao INCRA
    Set colPrecisaoH = lo.ListColumns("Precisao H (m)")
    Set colPrecisaoV = lo.ListColumns("Precisao V (m)")
    Set colMetodo = lo.ListColumns("Metodo Posic.")
    Set colCodLimite = lo.ListColumns("Cod. Limite")

    ' Se pelo menos uma coluna existe, preenche valores padrao
    If Not colPrecisaoH Is Nothing Or Not colPrecisaoV Is Nothing Or _
       Not colMetodo Is Nothing Or Not colCodLimite Is Nothing Then

        For i = 1 To lo.ListRows.Count
            ' Preenche Precisao H (padrao: 0.30m para limite artificial)
            If Not colPrecisaoH Is Nothing Then
                If IsEmpty(colPrecisaoH.DataBodyRange(i).Value) Or _
                   colPrecisaoH.DataBodyRange(i).Value = "" Or _
                   colPrecisaoH.DataBodyRange(i).Value = 0 Then
                    colPrecisaoH.DataBodyRange(i).Value = 0.3
                End If
            End If

            ' Preenche Precisao V (padrao: 0.50m)
            If Not colPrecisaoV Is Nothing Then
                If IsEmpty(colPrecisaoV.DataBodyRange(i).Value) Or _
                   colPrecisaoV.DataBodyRange(i).Value = "" Or _
                   colPrecisaoV.DataBodyRange(i).Value = 0 Then
                    colPrecisaoV.DataBodyRange(i).Value = 0.5
                End If
            End If

            ' Preenche Metodo Posicionamento (padrao: GNSS-RTK)
            If Not colMetodo Is Nothing Then
                If IsEmpty(colMetodo.DataBodyRange(i).Value) Or _
                   colMetodo.DataBodyRange(i).Value = "" Then
                    colMetodo.DataBodyRange(i).Value = "GNSS-RTK"
                End If
            End If

            ' Preenche Codigo Limite (padrao: LA1 - Cerca)
            If Not colCodLimite Is Nothing Then
                If IsEmpty(colCodLimite.DataBodyRange(i).Value) Or _
                   colCodLimite.DataBodyRange(i).Value = "" Then
                    colCodLimite.DataBodyRange(i).Value = "LA1"
                End If
            End If
        Next i

        ' Formata colunas numericas
        If Not colPrecisaoH Is Nothing Then colPrecisaoH.DataBodyRange.NumberFormat = "0.00"
        If Not colPrecisaoV Is Nothing Then colPrecisaoV.DataBodyRange.NumberFormat = "0.00"
    End If

    On Error GoTo 0
End Sub

' ==============================================================================
' CALCULAR AZIMUTE E DISTANCIA SEPARADOS - SGL
' ==============================================================================
Public Sub Calcular_Azimute_SGL()
    Dim wsSGL As Worksheet
    Dim loSGL As ListObject
    Dim i As Long, qtd As Long
    Dim lat1 As Double, lon1 As Double, lat2 As Double, lon2 As Double
    Dim azimute As Double
    
    On Error GoTo Erro
    Set wsSGL = ThisWorkbook.Sheets(M_Config.SH_SGL)
    Set loSGL = wsSGL.ListObjects(M_Config.TBL_SGL)
    
    If loSGL.ListRows.Count < 2 Then
        MsgBox "Minimo 2 vertices necessarios.", vbExclamation
        Exit Sub
    End If
    
    M_SheetProtection.DesbloquearPlanilha wsSGL
    Call M_Utils.Utils_OtimizarPerformance(True)
    
    qtd = loSGL.ListRows.Count
    
    For i = 1 To qtd
        lon1 = M_Utils.Str_DMS_Para_DD(CStr(loSGL.DataBodyRange(i, 2).Value))
        lat1 = M_Utils.Str_DMS_Para_DD(CStr(loSGL.DataBodyRange(i, 3).Value))
        
        If i < qtd Then
            lon2 = M_Utils.Str_DMS_Para_DD(CStr(loSGL.DataBodyRange(i + 1, 2).Value))
            lat2 = M_Utils.Str_DMS_Para_DD(CStr(loSGL.DataBodyRange(i + 1, 3).Value))
        Else
            lon2 = M_Utils.Str_DMS_Para_DD(CStr(loSGL.DataBodyRange(1, 2).Value))
            lat2 = M_Utils.Str_DMS_Para_DD(CStr(loSGL.DataBodyRange(1, 3).Value))
        End If
        
        azimute = M_Math_Geo.Geo_Azimute_Puissant(lat1, lon1, lat2, lon2)
        'azimute = M_Math_Geo.Geo_Azimute_Plano(lat1, lon1, lat2, lon2)
        loSGL.DataBodyRange(i, 6).Value = M_Utils.Str_FormatAzimute(azimute)
    Next i
    
    Call M_Utils.Utils_OtimizarPerformance(False)
    M_SheetProtection.BloquearPlanilha wsSGL
    MsgBox "Azimutes calculados!", vbInformation
    Exit Sub
    
Erro:
    Call M_Utils.Utils_OtimizarPerformance(False)
    M_SheetProtection.BloquearPlanilha wsSGL
    MsgBox "Erro: " & Err.Description, vbCritical
End Sub

Public Sub Calcular_Distancia_SGL()
    Dim wsSGL As Worksheet
    Dim loSGL As ListObject
    Dim i As Long, qtd As Long
    Dim lat1 As Double, lon1 As Double, lat2 As Double, lon2 As Double
    Dim distancia As Double
    
    On Error GoTo Erro
    Set wsSGL = ThisWorkbook.Sheets(M_Config.SH_SGL)
    Set loSGL = wsSGL.ListObjects(M_Config.TBL_SGL)
    
    If loSGL.ListRows.Count < 2 Then
        MsgBox "Minimo 2 vertices necessarios.", vbExclamation
        Exit Sub
    End If
    
    M_SheetProtection.DesbloquearPlanilha wsSGL
    Call M_Utils.Utils_OtimizarPerformance(True)
    
    qtd = loSGL.ListRows.Count
    
    For i = 1 To qtd
        lon1 = M_Utils.Str_DMS_Para_DD(CStr(loSGL.DataBodyRange(i, 2).Value))
        lat1 = M_Utils.Str_DMS_Para_DD(CStr(loSGL.DataBodyRange(i, 3).Value))
        
        If i < qtd Then
            lon2 = M_Utils.Str_DMS_Para_DD(CStr(loSGL.DataBodyRange(i + 1, 2).Value))
            lat2 = M_Utils.Str_DMS_Para_DD(CStr(loSGL.DataBodyRange(i + 1, 3).Value))
        Else
            lon2 = M_Utils.Str_DMS_Para_DD(CStr(loSGL.DataBodyRange(1, 2).Value))
            lat2 = M_Utils.Str_DMS_Para_DD(CStr(loSGL.DataBodyRange(1, 3).Value))
        End If
        
        distancia = M_Math_Geo.Math_Distancia_Geodesica(lat1, lon1, lat2, lon2)
        loSGL.DataBodyRange(i, 7).Value = Round(distancia, 2)
    Next i
    
    Call M_Utils.Utils_OtimizarPerformance(False)
    M_SheetProtection.BloquearPlanilha wsSGL
    MsgBox "Distancias calculadas!", vbInformation
    Exit Sub
    
Erro:
    Call M_Utils.Utils_OtimizarPerformance(False)
    M_SheetProtection.BloquearPlanilha wsSGL
    MsgBox "Erro: " & Err.Description, vbCritical
End Sub

' ==============================================================================
' CALCULAR AZIMUTE E DISTANCIA SEPARADOS - UTM
' ==============================================================================
Public Sub Calcular_Azimute_UTM()
    Dim wsUTM As Worksheet
    Dim loUTM As ListObject
    Dim i As Long, qtd As Long
    Dim N1 As Double, E1 As Double, N2 As Double, e2 As Double

    On Error GoTo Erro
    Set wsUTM = ThisWorkbook.Sheets(M_Config.SH_UTM)
    Set loUTM = wsUTM.ListObjects(M_Config.TBL_UTM)

    If loUTM.ListRows.Count < 2 Then
        MsgBox "Minimo 2 vertices necessarios.", vbExclamation
        Exit Sub
    End If

    M_SheetProtection.DesbloquearPlanilha wsUTM
    Call M_Utils.Utils_OtimizarPerformance(True)

    qtd = loUTM.ListRows.Count

    For i = 1 To qtd
        N1 = CDbl(loUTM.DataBodyRange(i, 2).Value)
        E1 = CDbl(loUTM.DataBodyRange(i, 3).Value)

        If i < qtd Then
            N2 = CDbl(loUTM.DataBodyRange(i + 1, 2).Value)
            e2 = CDbl(loUTM.DataBodyRange(i + 1, 3).Value)
        Else
            N2 = CDbl(loUTM.DataBodyRange(1, 2).Value)
            e2 = CDbl(loUTM.DataBodyRange(1, 3).Value)
        End If

        ' USA FUNÇÃO REFATORADA: Calcular_DistanciaAzimute_UTM
        Dim calc As Type_CalculoPonto
        calc = M_Math_Geo.Calcular_DistanciaAzimute_UTM(N1, E1, N2, e2)

        ' NOVO: Aplica correção de Convergência Meridiana
        ' Primeiro, obtém fuso e hemisfério da tabela
        Dim fusoUTM As Integer, hemisferio As String
        On Error Resume Next
        fusoUTM = M_UI_Main.UI_GetFusoAtual()
        hemisferio = M_UI_Main.UI_GetHemisferioAtual()
        If fusoUTM = 0 Then fusoUTM = 23  ' Padrão para Brasil
        If hemisferio = "" Then hemisferio = "S"
        On Error GoTo Erro

        ' Converte UTM → Geo para obter lat/lon e calcular convergência
        Dim geoAtual As Type_Geo
        geoAtual = M_Math_Geo.Converter_UTMParaGeo(N1, E1, fusoUTM, hemisferio)

        ' Aplica correção de convergência meridiana
        Dim azimuteGeod As Double
        azimuteGeod = M_Math_Geo.Converter_AzimuteGridParaGeod(calc.AzimuteDecimal, geoAtual.Latitude, geoAtual.Longitude, fusoUTM)

        ' USA NOVA FUNÇÃO: Str_FormatAzimuteGMS (com segundos)
        loUTM.DataBodyRange(i, 6).Value = M_Utils.Str_FormatAzimuteGMS(azimuteGeod)
    Next i

    Call M_Utils.Utils_OtimizarPerformance(False)
    M_SheetProtection.BloquearPlanilha wsUTM
    MsgBox "Azimutes calculados!", vbInformation
    Exit Sub

Erro:
    Call M_Utils.Utils_OtimizarPerformance(False)
    M_SheetProtection.BloquearPlanilha wsUTM
    MsgBox "Erro: " & Err.Description, vbCritical
End Sub

Public Sub Calcular_Distancia_UTM()
    Dim wsUTM As Worksheet
    Dim loUTM As ListObject
    Dim i As Long, qtd As Long
    Dim N1 As Double, E1 As Double, N2 As Double, e2 As Double
    Dim distancia As Double
    
    On Error GoTo Erro
    Set wsUTM = ThisWorkbook.Sheets(M_Config.SH_UTM)
    Set loUTM = wsUTM.ListObjects(M_Config.TBL_UTM)
    
    If loUTM.ListRows.Count < 2 Then
        MsgBox "Minimo 2 vertices necessarios.", vbExclamation
        Exit Sub
    End If
    
    M_SheetProtection.DesbloquearPlanilha wsUTM
    Call M_Utils.Utils_OtimizarPerformance(True)
    
    qtd = loUTM.ListRows.Count
    
    For i = 1 To qtd
        N1 = CDbl(loUTM.DataBodyRange(i, 2).Value)
        E1 = CDbl(loUTM.DataBodyRange(i, 3).Value)
        
        If i < qtd Then
            N2 = CDbl(loUTM.DataBodyRange(i + 1, 2).Value)
            e2 = CDbl(loUTM.DataBodyRange(i + 1, 3).Value)
        Else
            N2 = CDbl(loUTM.DataBodyRange(1, 2).Value)
            e2 = CDbl(loUTM.DataBodyRange(1, 3).Value)
        End If
        
        distancia = M_Math_Geo.Math_Distancia_Euclidiana(N1, E1, N2, e2)
        loUTM.DataBodyRange(i, 7).Value = Round(distancia, 2)
    Next i
    
    Call M_Utils.Utils_OtimizarPerformance(False)
    M_SheetProtection.BloquearPlanilha wsUTM
    MsgBox "Distancias calculadas!", vbInformation
    Exit Sub
    
Erro:
    Call M_Utils.Utils_OtimizarPerformance(False)
    M_SheetProtection.BloquearPlanilha wsUTM
    MsgBox "Erro: " & Err.Description, vbCritical
End Sub
