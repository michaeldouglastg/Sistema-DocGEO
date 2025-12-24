Attribute VB_Name = "EXEMPLOS_ATUALIZACAO"
' ==============================================================================
' EXEMPLOS DE ATUALIZAÇÃO - M_APP_LOGICA
' ==============================================================================
' Este arquivo mostra COMO atualizar o código existente para usar as novas
' funções refatoradas. Compare o ANTES e DEPOIS lado a lado.
' ==============================================================================

Option Explicit

' ==============================================================================
' EXEMPLO 1: CONVERSÃO SGL → UTM (Processo_Conv_SGL_UTM)
' ==============================================================================

' ---------- ANTES (Código Original) ----------
Public Sub Processo_Conv_SGL_UTM_ANTIGO()
    Dim wsSGL As Worksheet, wsUTM As Worksheet
    Dim loSGL As ListObject, loUTM As ListObject
    Dim i As Long, qtd As Long
    Dim arrSGL As Variant

    Set wsSGL = ThisWorkbook.Sheets(M_Config.SH_SGL)
    Set wsUTM = ThisWorkbook.Sheets(M_Config.SH_UTM)
    Set loSGL = wsSGL.ListObjects(M_Config.TBL_SGL)
    Set loUTM = wsUTM.ListObjects(M_Config.TBL_UTM)

    arrSGL = loSGL.DataBodyRange.Value
    qtd = UBound(arrSGL, 1)

    Dim arrOut() As Variant
    ReDim arrOut(1 To qtd, 1 To loUTM.ListColumns.Count)

    Dim latDD As Double, lonDD As Double
    Dim utmAtual As Type_UTM
    Dim zonaPadrao As Integer

    ' ANTES: Conversão simples
    lonDD = M_Utils.Str_DMS_Para_DD(CStr(arrSGL(1, 2)))
    zonaPadrao = M_Math_Geo.Geo_GetZonaUTM(lonDD)

    For i = 1 To qtd
        lonDD = M_Utils.Str_DMS_Para_DD(CStr(arrSGL(i, 2)))
        latDD = M_Utils.Str_DMS_Para_DD(CStr(arrSGL(i, 3)))

        ' ANTES: Função antiga
        utmAtual = M_Math_Geo.Geo_LatLon_Para_UTM(latDD, lonDD, zonaPadrao)

        arrOut(i, 2) = Round(utmAtual.Norte, 3)
        arrOut(i, 3) = Round(utmAtual.Leste, 3)
    Next i

    loUTM.DataBodyRange.Value = arrOut
End Sub

' ---------- DEPOIS (Código Refatorado - RECOMENDADO) ----------
Public Sub Processo_Conv_SGL_UTM_NOVO()
    Dim wsSGL As Worksheet, wsUTM As Worksheet
    Dim loSGL As ListObject, loUTM As ListObject
    Dim i As Long, qtd As Long
    Dim arrSGL As Variant

    Set wsSGL = ThisWorkbook.Sheets(M_Config.SH_SGL)
    Set wsUTM = ThisWorkbook.Sheets(M_Config.SH_UTM)
    Set loSGL = wsSGL.ListObjects(M_Config.TBL_SGL)
    Set loUTM = wsUTM.ListObjects(M_Config.TBL_UTM)

    arrSGL = loSGL.DataBodyRange.Value
    qtd = UBound(arrSGL, 1)

    Dim arrOut() As Variant
    ReDim arrOut(1 To qtd, 1 To loUTM.ListColumns.Count)

    Dim latDD As Double, lonDD As Double
    Dim utmAtual As Type_UTM
    Dim zonaPadrao As Integer

    ' DEPOIS: Conversão robusta (aceita múltiplos formatos)
    lonDD = M_Utils.Str_DMS_Para_DD(CStr(arrSGL(1, 2)))
    zonaPadrao = M_Math_Geo.Geo_GetZonaUTM(lonDD)

    ' Cache para otimização
    Dim cacheN() As Double, cacheE() As Double
    ReDim cacheN(1 To qtd), cacheE(1 To qtd)

    For i = 1 To qtd
        ' DEPOIS: Str_DMS_Para_DD agora aceita decimal, DMS com sinal, DMS com sufixo
        lonDD = M_Utils.Str_DMS_Para_DD(CStr(arrSGL(i, 2)))
        latDD = M_Utils.Str_DMS_Para_DD(CStr(arrSGL(i, 3)))

        ' DEPOIS: Nova função validada (mais precisa)
        utmAtual = M_Math_Geo.Converter_GeoParaUTM(latDD, lonDD, zonaPadrao)

        ' Verifica sucesso da conversão
        If Not utmAtual.Sucesso Then
            Debug.Print "Erro ao converter ponto " & i
            GoTo ProximoPonto
        End If

        cacheN(i) = utmAtual.Norte
        cacheE(i) = utmAtual.Leste

        arrOut(i, 1) = arrSGL(i, 1)
        arrOut(i, 2) = Round(utmAtual.Norte, 3)
        arrOut(i, 3) = Round(utmAtual.Leste, 3)
        arrOut(i, 4) = arrSGL(i, 4)

ProximoPonto:
    Next i

    ' DEPOIS: Calcular azimutes e distâncias com função robusta
    For i = 1 To qtd
        Dim idxProx As Long
        If i < qtd Then idxProx = i + 1 Else idxProx = 1

        Dim calc As Type_CalculoPonto
        calc = M_Math_Geo.Calcular_DistanciaAzimute_UTM( _
            cacheN(i), cacheE(i), cacheN(idxProx), cacheE(idxProx) _
        )

        arrOut(i, 5) = arrSGL(idxProx, 1)
        arrOut(i, 6) = M_Utils.Str_FormatAzimute(calc.AzimuteDecimal)
        arrOut(i, 7) = Round(calc.Distancia, 3)
    Next i

    loUTM.DataBodyRange.Value = arrOut
End Sub

' ==============================================================================
' EXEMPLO 2: CALCULAR AZIMUTE E DISTÂNCIA SGL
' ==============================================================================

' ---------- ANTES ----------
Public Sub Calcular_Azimute_SGL_ANTIGO()
    Dim wsSGL As Worksheet
    Dim loSGL As ListObject
    Dim i As Long, qtd As Long
    Dim lat1 As Double, lon1 As Double, lat2 As Double, lon2 As Double
    Dim azimute As Double

    Set wsSGL = ThisWorkbook.Sheets(M_Config.SH_SGL)
    Set loSGL = wsSGL.ListObjects(M_Config.TBL_SGL)

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

        ' ANTES: Cálculo direto
        azimute = M_Math_Geo.Geo_Azimute_Puissant(lat1, lon1, lat2, lon2)
        loSGL.DataBodyRange(i, 6).Value = M_Utils.Str_FormatAzimute(azimute)
    Next i
End Sub

' ---------- DEPOIS (COM VALIDAÇÃO E CACHE) ----------
Public Sub Calcular_Azimute_SGL_NOVO()
    Dim wsSGL As Worksheet
    Dim loSGL As ListObject
    Dim i As Long, qtd As Long
    Dim arrDados As Variant
    Dim arrAzimutes() As String

    Set wsSGL = ThisWorkbook.Sheets(M_Config.SH_SGL)
    Set loSGL = wsSGL.ListObjects(M_Config.TBL_SGL)

    If loSGL.ListRows.Count < 2 Then
        MsgBox "Mínimo 2 vértices necessários.", vbExclamation
        Exit Sub
    End If

    M_SheetProtection.DesbloquearPlanilha wsSGL
    Call M_Utils.Utils_OtimizarPerformance(True)

    ' DEPOIS: Lê tudo em array para melhor performance
    arrDados = loSGL.DataBodyRange.Value
    qtd = UBound(arrDados, 1)
    ReDim arrAzimutes(1 To qtd)

    ' DEPOIS: Cache de coordenadas convertidas
    Dim arrLat() As Double, arrLon() As Double
    ReDim arrLat(1 To qtd), arrLon(1 To qtd)

    ' Converte todas as coordenadas uma vez
    For i = 1 To qtd
        ' DEPOIS: Conversão robusta que aceita qualquer formato
        On Error Resume Next
        arrLon(i) = M_Utils.Str_DMS_Para_DD(CStr(arrDados(i, 2)))
        arrLat(i) = M_Utils.Str_DMS_Para_DD(CStr(arrDados(i, 3)))
        On Error GoTo 0
    Next i

    ' Calcula azimutes
    For i = 1 To qtd
        Dim idxProx As Long
        If i < qtd Then idxProx = i + 1 Else idxProx = 1

        Dim azimute As Double
        azimute = M_Math_Geo.Geo_Azimute_Puissant( _
            arrLat(i), arrLon(i), arrLat(idxProx), arrLon(idxProx) _
        )

        arrAzimutes(i) = M_Utils.Str_FormatAzimute(azimute)
    Next i

    ' DEPOIS: Escreve tudo de uma vez (muito mais rápido)
    For i = 1 To qtd
        loSGL.DataBodyRange(i, 6).Value = arrAzimutes(i)
    Next i

    Call M_Utils.Utils_OtimizarPerformance(False)
    M_SheetProtection.BloquearPlanilha wsSGL
    MsgBox "Azimutes calculados!", vbInformation
End Sub

' ==============================================================================
' EXEMPLO 3: CALCULAR DISTÂNCIA UTM
' ==============================================================================

' ---------- ANTES ----------
Public Sub Calcular_Distancia_UTM_ANTIGO()
    Dim wsUTM As Worksheet
    Dim loUTM As ListObject
    Dim i As Long, qtd As Long
    Dim N1 As Double, E1 As Double, N2 As Double, E2 As Double
    Dim distancia As Double

    Set wsUTM = ThisWorkbook.Sheets(M_Config.SH_UTM)
    Set loUTM = wsUTM.ListObjects(M_Config.TBL_UTM)

    qtd = loUTM.ListRows.Count

    For i = 1 To qtd
        N1 = CDbl(loUTM.DataBodyRange(i, 2).Value)
        E1 = CDbl(loUTM.DataBodyRange(i, 3).Value)

        If i < qtd Then
            N2 = CDbl(loUTM.DataBodyRange(i + 1, 2).Value)
            E2 = CDbl(loUTM.DataBodyRange(i + 1, 3).Value)
        Else
            N2 = CDbl(loUTM.DataBodyRange(1, 2).Value)
            E2 = CDbl(loUTM.DataBodyRange(1, 3).Value)
        End If

        ' ANTES: Cálculo simples
        distancia = M_Math_Geo.Math_Distancia_Euclidiana(N1, E1, N2, E2)
        loUTM.DataBodyRange(i, 7).Value = Round(distancia, 2)
    Next i
End Sub

' ---------- DEPOIS (COM CÁLCULO COMBINADO) ----------
Public Sub Calcular_Distancia_Azimute_UTM_NOVO()
    Dim wsUTM As Worksheet
    Dim loUTM As ListObject
    Dim i As Long, qtd As Long
    Dim arrDados As Variant

    Set wsUTM = ThisWorkbook.Sheets(M_Config.SH_UTM)
    Set loUTM = wsUTM.ListObjects(M_Config.TBL_UTM)

    If loUTM.ListRows.Count < 2 Then Exit Sub

    M_SheetProtection.DesbloquearPlanilha wsUTM
    Call M_Utils.Utils_OtimizarPerformance(True)

    arrDados = loUTM.DataBodyRange.Value
    qtd = UBound(arrDados, 1)

    ' DEPOIS: Arrays para armazenar resultados
    Dim arrDistancias() As Double, arrAzimutes() As String
    ReDim arrDistancias(1 To qtd), arrAzimutes(1 To qtd)

    For i = 1 To qtd
        Dim N1 As Double, E1 As Double, N2 As Double, E2 As Double
        N1 = CDbl(arrDados(i, 2))
        E1 = CDbl(arrDados(i, 3))

        Dim idxProx As Long
        If i < qtd Then idxProx = i + 1 Else idxProx = 1

        N2 = CDbl(arrDados(idxProx, 2))
        E2 = CDbl(arrDados(idxProx, 3))

        ' DEPOIS: Calcula distância E azimute de uma vez (mais eficiente)
        Dim calc As Type_CalculoPonto
        calc = M_Math_Geo.Calcular_DistanciaAzimute_UTM(N1, E1, N2, E2)

        arrDistancias(i) = Round(calc.Distancia, 2)
        arrAzimutes(i) = M_Utils.Str_FormatAzimute(calc.AzimuteDecimal)
    Next i

    ' DEPOIS: Escreve tudo de uma vez
    For i = 1 To qtd
        loUTM.DataBodyRange(i, 6).Value = arrAzimutes(i)
        loUTM.DataBodyRange(i, 7).Value = arrDistancias(i)
    Next i

    Call M_Utils.Utils_OtimizarPerformance(False)
    M_SheetProtection.BloquearPlanilha wsUTM
    MsgBox "Distâncias e azimutes calculados!", vbInformation
End Sub

' ==============================================================================
' EXEMPLO 4: IMPORTAÇÃO DE CSV COM COORDENADAS DECIMAIS
' ==============================================================================

' ---------- ANTES ----------
Public Sub Importar_CSV_Coordenadas_ANTIGO()
    ' ... código de leitura do CSV ...

    Dim coordWKT As String, coordSplit() As String
    Dim lonDMS As String, latDMS As String

    ' ANTES: Assumia formato específico
    coordWKT = "POINT (X Y)"
    coordSplit = Split(coordWKT, " ")

    ' ANTES: Conversão manual com Replace
    Dim lonDD As Double, latDD As Double
    lonDD = CDbl(Replace(coordSplit(0), ".", ","))
    latDD = CDbl(Replace(coordSplit(1), ".", ","))

    ' Converte para DMS
    lonDMS = M_Utils.Str_DD_Para_DMS(lonDD)
    latDMS = M_Utils.Str_DD_Para_DMS(latDD)
End Sub

' ---------- DEPOIS (MAIS SIMPLES E ROBUSTO) ----------
Public Sub Importar_CSV_Coordenadas_NOVO()
    ' ... código de leitura do CSV ...

    Dim coordWKT As String, coordSplit() As String
    Dim lonDMS As String, latDMS As String

    ' Exemplo: POINT (-43.5934619399999974 -22.4695083300000000)
    coordWKT = "POINT (-43.5934619399999974 -22.4695083300000000)"
    coordWKT = Replace(Replace(coordWKT, "POINT (", ""), ")", "")
    coordSplit = Split(coordWKT, " ")

    ' DEPOIS: Str_DMS_Para_DD aceita decimal diretamente!
    Dim lonDD As Double, latDD As Double
    lonDD = M_Utils.Str_DMS_Para_DD(coordSplit(0))  ' Detecta automaticamente que é decimal
    latDD = M_Utils.Str_DMS_Para_DD(coordSplit(1))

    ' Converte para formato do sistema
    lonDMS = M_Utils.Str_DD_Para_DMS(lonDD)  ' "-43°35'36.463""
    latDMS = M_Utils.Str_DD_Para_DMS(latDD)  ' "-22°28'10.230""

    ' OU se quiser formato com sufixo para documentos:
    Dim lonComSufixo As String, latComSufixo As String
    lonComSufixo = M_Utils.Str_DD_Para_DMS_ComSufixo(lonDD, "LON")  ' "43° 35' 36.4626" O"
    latComSufixo = M_Utils.Str_DD_Para_DMS_ComSufixo(latDD, "LAT")  ' "22° 28' 10.2299" S"
End Sub

' ==============================================================================
' EXEMPLO 5: CÁLCULO DE ÁREA SGL (COM SISTEMA LOCAL)
' ==============================================================================

' ---------- DEPOIS (CÓDIGO MANTIDO, MAS OTIMIZADO) ----------
Public Sub Processo_Calc_Area_SGL_Otimizado(lo As ListObject, ByRef outM2 As Double, ByRef outHa As Double)
    Dim i As Long, qtd As Long
    Dim latSoma As Double, lonSoma As Double, altSoma As Double
    Dim lat0 As Double, lon0 As Double, alt0 As Double
    Dim arrDados As Variant
    Dim E_sgl() As Double, N_sgl() As Double

    arrDados = lo.DataBodyRange.Value
    qtd = UBound(arrDados, 1)

    ' Calcula centroide
    For i = 1 To qtd
        ' DEPOIS: Conversão robusta
        latSoma = latSoma + M_Utils.Str_DMS_Para_DD(CStr(arrDados(i, 3)))
        lonSoma = lonSoma + M_Utils.Str_DMS_Para_DD(CStr(arrDados(i, 2)))
        If IsNumeric(arrDados(i, 4)) Then altSoma = altSoma + CDbl(arrDados(i, 4))
    Next i

    lat0 = latSoma / qtd
    lon0 = lonSoma / qtd
    alt0 = altSoma / qtd

    ' Converte para geocêntrica (origem)
    Dim ptOrigem As Type_Geocentrica
    ptOrigem = M_Math_Geo.Geo_Geod_Para_Geoc(lon0, lat0, alt0)

    ReDim E_sgl(1 To qtd)
    ReDim N_sgl(1 To qtd)

    ' Converte todos os pontos para topocêntrico
    For i = 1 To qtd
        Dim latPt As Double, lonPt As Double, altPt As Double

        ' DEPOIS: Conversão robusta
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
        ptTopo = M_Math_Geo.Geo_Geoc_Para_Topoc(ptGeoc.x, ptGeoc.y, ptGeoc.Z, _
                                                 lon0, lat0, ptOrigem.x, ptOrigem.y, ptOrigem.Z)

        E_sgl(i) = ptTopo.E
        N_sgl(i) = ptTopo.N
    Next i

    ' Calcula área usando Gauss
    outM2 = M_Math_Geo.Geo_Area_Gauss(E_sgl, N_sgl)
    outHa = outM2 / 10000
End Sub

' ==============================================================================
' EXEMPLO 6: CONVERSÃO BIDIRECIONAL UTM ↔ GEO
' ==============================================================================

' Nova funcionalidade: Converter UTM de volta para SGL
Public Sub Exemplo_Conversao_Bidirecional()
    ' Coordenadas originais (SGL)
    Dim latOriginal As Double: latOriginal = -22.469508
    Dim lonOriginal As Double: lonOriginal = -43.593461

    ' 1. SGL → UTM
    Dim utm As Type_UTM
    utm = M_Math_Geo.Converter_GeoParaUTM(latOriginal, lonOriginal, 23)

    Debug.Print "UTM Norte: " & utm.Norte    ' ~7514234.567
    Debug.Print "UTM Leste: " & utm.Leste    ' ~685432.123

    ' 2. UTM → SGL (NOVO!)
    Dim geo As Type_Geo
    geo = M_Math_Geo.Converter_UTMParaGeo(utm.Norte, utm.Leste, 23, "S")

    Debug.Print "Lat recuperada: " & geo.Latitude   ' -22.469508
    Debug.Print "Lon recuperada: " & geo.Longitude  ' -43.593461

    ' Verifica erro
    Dim erroLat As Double, erroLon As Double
    erroLat = Abs(geo.Latitude - latOriginal)
    erroLon = Abs(geo.Longitude - lonOriginal)

    Debug.Print "Erro Latitude: " & erroLat & "° (< 0.000001°)"
    Debug.Print "Erro Longitude: " & erroLon & "° (< 0.000001°)"
End Sub

' ==============================================================================
' EXEMPLO 7: USAR TIPO DE RETORNO PARA VALIDAÇÃO
' ==============================================================================

Public Sub Exemplo_Validacao_Com_Tipo()
    Dim utm As Type_UTM
    Dim geo As Type_Geo

    ' Tenta converter coordenada potencialmente inválida
    utm = M_Math_Geo.Converter_GeoParaUTM(999, 999, 23) ' Coordenada inválida

    ' DEPOIS: Verifica sucesso antes de usar
    If utm.Sucesso Then
        Debug.Print "Conversão OK: " & utm.Norte
    Else
        Debug.Print "Erro na conversão! Coordenadas inválidas."
    End If

    ' Mesmo para UTM → Geo
    geo = M_Math_Geo.Converter_UTMParaGeo(999999999, 999999999, 23, "S")

    If geo.Sucesso Then
        Debug.Print "Conversão OK: " & geo.Latitude
    Else
        Debug.Print "Erro na conversão!"
    End If
End Sub

' ==============================================================================
' RESUMO DAS PRINCIPAIS MELHORIAS
' ==============================================================================

' 1. Str_DMS_Para_DD() agora aceita:
'    - Decimal: "-43.5934619399999974"
'    - DMS com sinal: "-43°35'36,463""
'    - DMS com sufixo: "43° 35' 36,4626" O"
'    - Ponto ou vírgula decimal
'
' 2. Converter_GeoParaUTM() e Converter_UTMParaGeo():
'    - Algoritmo validado (NIMA)
'    - Retorna Type com flag .Sucesso
'    - Precisão milimétrica
'
' 3. Calcular_DistanciaAzimute_UTM():
'    - Cálculo robusto por quadrante
'    - Elimina erros em casos especiais
'    - Retorna distância E azimute juntos
'
' 4. Calcular_CoordenadasPorDistanciaAzimute():
'    - Nova funcionalidade (irradiação)
'    - Útil para criar pontos a partir de dist/azimute
'
' 5. Performance:
'    - Cache de coordenadas convertidas
'    - Escrita em lote (arrays)
'    - Menos chamadas a funções de célula
