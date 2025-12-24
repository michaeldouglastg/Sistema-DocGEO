Attribute VB_Name = "M_Math_Geo"
Option Explicit
' ==============================================================================
' MODULO: M_MATH_GEO (REFATORADO)
' DESCRICAO: BIBLIOTECA DE CALCULOS GEODESICOS E CONVERSOES UTM
' VERSAO: 2.0 - Algoritmos validados e otimizados
' ==============================================================================

' ==============================================================================
' TIPOS DE DADOS PERSONALIZADOS
' ==============================================================================

Public Type Type_UTM
    Norte As Double
    Leste As Double
    fuso As Integer
    Hemisferio As String ' "N" ou "S"
    Sucesso As Boolean
End Type

Public Type Type_Geo
    Latitude As Double
    Longitude As Double
    Sucesso As Boolean
End Type

Public Type Type_Geocentrica
    x As Double
    y As Double
    Z As Double
End Type

Public Type Type_Topocentrica
    E As Double
    N As Double
    U As Double
End Type

Public Type Type_CalculoPonto
    Distancia As Double
    AzimuteDecimal As Double
    Sucesso As Boolean
End Type

Public Type Type_PontoUTM
    Norte As Double
    Leste As Double
    Sucesso As Boolean
End Type

' ==============================================================================
' CONSTANTES GEODESICAS (SIRGAS 2000 / WGS84)
' ==============================================================================
Private Const PI As Double = 3.14159265358979
Private Const SEMI_EIXO As Double = 6378137#           ' Semi-eixo maior (a)
Private Const ACHAT As Double = 0.00335281068118       ' Achatamento (f)
Private Const K0 As Double = 0.9996                    ' Fator de escala UTM
Private Const FALSO_NORTE_SUL As Double = 10000000#    ' False Northing (hemisfério sul)
Private Const FALSO_LESTE As Double = 500000#          ' False Easting

Public Const CONST_PI As Double = 3.14159265358979
Public Const CONST_RAIO_TERRA As Double = 6371000
Public Const CONST_SEMI_EIXO_MAIOR As Double = 6378137#
Public Const CONST_ACHATAMENTO As Double = 0.00335281068118
Public Const CONST_FATOR_K0 As Double = 0.9996

' ==============================================================================
' CONVERSAO GEOGRAFICA (LAT/LON) PARA UTM - VERSAO VALIDADA
' ==============================================================================

Public Function Converter_GeoParaUTM(ByVal Latitude As Double, ByVal Longitude As Double, ByVal fuso As Integer) As Type_UTM
    '----------------------------------------------------------------------------------
    ' Algoritmo: Transversa de Mercator (Elipsoide WGS84 / SIRGAS2000)
    ' Fonte: NIMA (National Imagery and Mapping Agency) Technical Manual
    ' Precisão: Milimétrica para coordenadas no território brasileiro
    '----------------------------------------------------------------------------------
    Dim resultado As Type_UTM

    On Error GoTo ErroConversao

    ' Constantes do Elipsoide
    Dim k0 As Double: k0 = CONST_FATOR_K0
    Dim a As Double: a = CONST_SEMI_EIXO_MAIOR
    Dim f As Double: f = 1 / 298.257223563  ' Achatamento inverso
    Dim e2 As Double: e2 = 2 * f - f ^ 2
    Dim e_linha2 As Double: e_linha2 = e2 / (1 - e2)

    ' Conversão para Radianos
    Dim lat_rad As Double: lat_rad = Latitude * PI / 180
    Dim lon_rad As Double: lon_rad = Longitude * PI / 180

    ' Meridiano Central do Fuso (MC)
    ' Fórmula: MC = (Fuso * 6) - 183
    Dim lon_cm_rad As Double: lon_cm_rad = ((fuso * 6) - 183) * PI / 180

    ' Variáveis Auxiliares de Cálculo
    Dim N As Double, T As Double, C As Double, A_term As Double, M As Double

    N = a / Sqr(1 - e2 * Sin(lat_rad) ^ 2)
    T = Tan(lat_rad) ^ 2
    C = e_linha2 * Cos(lat_rad) ^ 2
    A_term = (lon_rad - lon_cm_rad) * Cos(lat_rad)

    ' Cálculo do Arco do Meridiano (M)
    M = a * ((1 - e2 / 4 - 3 * e2 ^ 2 / 64 - 5 * e2 ^ 3 / 256) * lat_rad _
        - (3 * e2 / 8 + 3 * e2 ^ 2 / 32 + 45 * e2 ^ 3 / 1024) * Sin(2 * lat_rad) _
        + (15 * e2 ^ 2 / 256 + 45 * e2 ^ 3 / 1024) * Sin(4 * lat_rad) _
        - (35 * e2 ^ 3 / 3072) * Sin(6 * lat_rad))

    ' --- Cálculo Final Leste (E) ---
    resultado.Leste = k0 * N * (A_term + (1 - T + C) * A_term ^ 3 / 6 + _
                      (5 - 18 * T + T ^ 2 + 72 * C - 58 * e_linha2) * A_term ^ 5 / 120) + FALSO_LESTE

    ' --- Cálculo Final Norte (N) ---
    resultado.Norte = k0 * (M + N * Tan(lat_rad) * (A_term ^ 2 / 2 + _
                      (5 - T + 9 * C + 4 * C ^ 2) * A_term ^ 4 / 24 + _
                      (61 - 58 * T + T ^ 2 + 600 * C - 330 * e_linha2) * A_term ^ 6 / 720))

    ' Correção para Hemisfério Sul
    If Latitude < 0 Then
        resultado.Norte = resultado.Norte + FALSO_NORTE_SUL
        resultado.Hemisferio = "S"
    Else
        resultado.Hemisferio = "N"
    End If

    ' Preenche o restante da estrutura
    resultado.fuso = fuso
    resultado.Sucesso = True

    Converter_GeoParaUTM = resultado
    Exit Function

ErroConversao:
    resultado.Sucesso = False
    resultado.Norte = 0
    resultado.Leste = 0
    Converter_GeoParaUTM = resultado
End Function

' ==============================================================================
' CONVERSAO UTM PARA GEOGRAFICA (LAT/LON) - VERSAO VALIDADA
' ==============================================================================

Public Function Converter_UTMParaGeo(ByVal Norte As Double, ByVal Leste As Double, _
                                      ByVal fuso As Integer, ByVal Hemisferio As String) As Type_Geo
    '----------------------------------------------------------------------------------
    ' Algoritmo: Inversa da Transversa de Mercator
    ' Converte coordenadas UTM para Latitude/Longitude (SIRGAS 2000)
    '----------------------------------------------------------------------------------
    Dim resultado As Type_Geo
    On Error GoTo ErroConversao

    ' Constantes do Elipsoide
    Dim k0 As Double: k0 = CONST_FATOR_K0
    Dim a As Double: a = CONST_SEMI_EIXO_MAIOR
    Dim f As Double: f = 1 / 298.257223563
    Dim e2 As Double: e2 = 2 * f - f ^ 2
    Dim e_linha2 As Double: e_linha2 = e2 / (1 - e2)

    ' Variáveis de Cálculo
    Dim x As Double, y As Double, lon_cm_rad As Double, M As Double, mu As Double
    Dim e1 As Double, phi1_rad As Double, n1 As Double, T1 As Double, C1 As Double
    Dim R1 As Double, D As Double, lat_rad As Double, lon_rad As Double

    ' Remove False Easting e False Northing
    x = Leste - FALSO_LESTE

    If UCase(Left(Hemisferio, 1)) = "S" Then
        y = Norte - FALSO_NORTE_SUL
    Else
        y = Norte
    End If

    ' Meridiano Central
    lon_cm_rad = (fuso * 6 - 183) * PI / 180

    ' Cálculo do Footpoint Latitude (phi1)
    M = y / k0
    mu = M / (a * (1 - e2 / 4 - 3 * e2 ^ 2 / 64 - 5 * e2 ^ 3 / 256))
    e1 = (1 - Sqr(1 - e2)) / (1 + Sqr(1 - e2))

    phi1_rad = mu + (3 * e1 / 2 - 27 * e1 ^ 3 / 32) * Sin(2 * mu) _
                  + (21 * e1 ^ 2 / 16 - 55 * e1 ^ 4 / 32) * Sin(4 * mu) _
                  + (151 * e1 ^ 3 / 96) * Sin(6 * mu) _
                  + (1097 * e1 ^ 4 / 512) * Sin(8 * mu)

    ' Raios de Curvatura e Termos Auxiliares
    n1 = a / Sqr(1 - e2 * Sin(phi1_rad) ^ 2)
    T1 = Tan(phi1_rad) ^ 2
    C1 = e_linha2 * Cos(phi1_rad) ^ 2
    R1 = a * (1 - e2) / (1 - e2 * Sin(phi1_rad) ^ 2) ^ 1.5
    D = x / (n1 * k0)

    ' Latitude Final
    lat_rad = phi1_rad - (n1 * Tan(phi1_rad) / R1) * _
              (D ^ 2 / 2 - (5 + 3 * T1 + 10 * C1 - 4 * C1 ^ 2 - 9 * e_linha2) * D ^ 4 / 24 + _
              (61 + 90 * T1 + 298 * C1 + 45 * T1 ^ 2 - 252 * e_linha2 - 3 * C1 ^ 2) * D ^ 6 / 720)

    ' Longitude Final
    lon_rad = lon_cm_rad + (D - (1 + 2 * T1 + C1) * D ^ 3 / 6 + _
              (5 - 2 * C1 + 28 * T1 - 3 * C1 ^ 2 + 8 * e_linha2 + 24 * T1 ^ 2) * D ^ 5 / 120) / Cos(phi1_rad)

    ' Converte para Graus
    resultado.Latitude = lat_rad * 180 / PI
    resultado.Longitude = lon_rad * 180 / PI
    resultado.Sucesso = True

    Converter_UTMParaGeo = resultado
    Exit Function

ErroConversao:
    resultado.Sucesso = False
    Converter_UTMParaGeo = resultado
End Function

' ==============================================================================
' FUNCOES DE COMPATIBILIDADE COM O SISTEMA ATUAL
' ==============================================================================

Public Function Geo_LatLon_Para_UTM(ByVal Lat As Double, ByVal Lon As Double, _
                                     Optional ByVal ZonaForcada As Long = 0) As Type_UTM
    '----------------------------------------------------------------------------------
    ' Função de compatibilidade com o sistema atual
    ' Calcula fuso automaticamente ou usa o forçado
    '----------------------------------------------------------------------------------
    Dim fusoCalculado As Integer

    If ZonaForcada > 0 Then
        fusoCalculado = ZonaForcada
    Else
        fusoCalculado = Geo_GetZonaUTM(Lon)
    End If

    Geo_LatLon_Para_UTM = Converter_GeoParaUTM(Lat, Lon, fusoCalculado)
End Function

Public Function Geo_UTM_Para_LatLon(ByVal Norte As Double, ByVal Este As Double, _
                                     ByVal Zona As Long, ByVal hemisferioSul As Boolean) As Object
    '----------------------------------------------------------------------------------
    ' Retorna Dictionary para compatibilidade com código antigo
    '----------------------------------------------------------------------------------
    Dim resultado As Object
    Dim geo As Type_Geo
    Dim hemisferio As String

    If hemisferioSul Then hemisferio = "S" Else hemisferio = "N"

    geo = Converter_UTMParaGeo(Norte, Este, Zona, hemisferio)

    Set resultado = CreateObject("Scripting.Dictionary")
    resultado.Add "Latitude", geo.Latitude
    resultado.Add "Longitude", geo.Longitude

    Set Geo_UTM_Para_LatLon = resultado
End Function

Public Function Geo_GetZonaUTM(Longitude As Double) As Integer
    '----------------------------------------------------------------------------------
    ' Calcula o fuso UTM baseado na Longitude
    ' Fórmula: Fuso = Int((Longitude + 180) / 6) + 1
    '----------------------------------------------------------------------------------
    Dim lonNormalizada As Double
    lonNormalizada = Longitude

    ' Garante intervalo [-180, 180]
    If lonNormalizada > 180 Then lonNormalizada = lonNormalizada - 360
    If lonNormalizada < -180 Then lonNormalizada = lonNormalizada + 360

    Geo_GetZonaUTM = Int((lonNormalizada + 180) / 6) + 1
End Function

' ==============================================================================
' CALCULOS DE DISTANCIA E AZIMUTE - SISTEMA UTM (PLANO)
' ==============================================================================

Public Function Calcular_DistanciaAzimute_UTM(ByVal Norte1 As Double, ByVal Leste1 As Double, _
                                                ByVal Norte2 As Double, ByVal Leste2 As Double) As Type_CalculoPonto
    '----------------------------------------------------------------------------------
    ' Calcula distância e azimute entre dois pontos UTM
    ' Usa geometria plana (válido para pontos próximos no mesmo fuso)
    '----------------------------------------------------------------------------------
    Dim DeltaNorte As Double, DeltaLeste As Double
    Dim AzimuteDecimal As Double
    Dim resultado As Type_CalculoPonto

    ' Calcula as diferenças
    DeltaNorte = Norte2 - Norte1
    DeltaLeste = Leste2 - Leste1

    ' --- Cálculo da Distância (Pitágoras) ---
    resultado.Distancia = Sqr(DeltaNorte ^ 2 + DeltaLeste ^ 2)

    ' --- Cálculo do Azimute (Robusto por Quadrante) ---
    If DeltaNorte = 0 And DeltaLeste = 0 Then
        AzimuteDecimal = 0
    Else
        Dim anguloBaseRad As Double
        Dim anguloBaseGraus As Double

        ' Calcula o ângulo base no primeiro quadrante (0-90°)
        If Abs(DeltaNorte) > 0.000001 Then
            anguloBaseRad = Atn(Abs(DeltaLeste) / Abs(DeltaNorte))
            anguloBaseGraus = anguloBaseRad * 180 / PI

            ' Ajusta para o quadrante correto
            If DeltaLeste > 0 And DeltaNorte > 0 Then       ' 1º Quadrante (NE)
                AzimuteDecimal = anguloBaseGraus
            ElseIf DeltaLeste > 0 And DeltaNorte < 0 Then   ' 2º Quadrante (SE)
                AzimuteDecimal = 180 - anguloBaseGraus
            ElseIf DeltaLeste < 0 And DeltaNorte < 0 Then   ' 3º Quadrante (SW)
                AzimuteDecimal = 180 + anguloBaseGraus
            ElseIf DeltaLeste < 0 And DeltaNorte > 0 Then   ' 4º Quadrante (NW)
                AzimuteDecimal = 360 - anguloBaseGraus
            End If
        Else
            ' Casos especiais (eixos E-W)
            If DeltaLeste > 0 Then AzimuteDecimal = 90 Else AzimuteDecimal = 270
        End If

        ' Casos especiais (eixos N-S quando DeltaLeste = 0)
        If Abs(DeltaLeste) < 0.000001 Then
            If DeltaNorte > 0 Then AzimuteDecimal = 0 Else AzimuteDecimal = 180
        End If
    End If

    resultado.AzimuteDecimal = AzimuteDecimal
    resultado.Sucesso = True
    Calcular_DistanciaAzimute_UTM = resultado
End Function

Public Function Calcular_CoordenadasPorDistanciaAzimute(ByVal NorteInicial As Double, ByVal LesteInicial As Double, _
                                                          ByVal Distancia As Double, ByVal AzimuteDecimal As Double) As Type_PontoUTM
    '----------------------------------------------------------------------------------
    ' Calcula coordenadas de um novo ponto a partir de:
    ' - Ponto inicial (Norte, Leste)
    ' - Distância
    ' - Azimute (em graus decimais)
    '----------------------------------------------------------------------------------
    Dim AzimuteRad As Double
    Dim DeltaNorte As Double
    Dim DeltaLeste As Double
    Dim resultado As Type_PontoUTM

    ' Converte Azimute para Radianos
    AzimuteRad = AzimuteDecimal * (PI / 180)

    ' Calcula as projeções (trigonometria de azimute)
    DeltaNorte = Distancia * Cos(AzimuteRad)
    DeltaLeste = Distancia * Sin(AzimuteRad)

    ' Soma às coordenadas iniciais
    resultado.Norte = NorteInicial + DeltaNorte
    resultado.Leste = LesteInicial + DeltaLeste
    resultado.Sucesso = True

    Calcular_CoordenadasPorDistanciaAzimute = resultado
End Function

' ==============================================================================
' FUNCOES DE COMPATIBILIDADE - GEOMETRIA PLANA
' ==============================================================================

Public Function Math_Distancia_Euclidiana(x1 As Double, y1 As Double, x2 As Double, y2 As Double) As Double
    Math_Distancia_Euclidiana = Sqr((x2 - x1) ^ 2 + (y2 - y1) ^ 2)
End Function

Public Function Geo_Azimute_Plano(x1 As Double, y1 As Double, x2 As Double, y2 As Double) As Double
    Dim calc As Type_CalculoPonto
    calc = Calcular_DistanciaAzimute_UTM(y1, x1, y2, x2) ' Note: invertido para compatibilidade
    Geo_Azimute_Plano = calc.AzimuteDecimal
End Function

' ==============================================================================
' AREA DE POLIGONO - METODO DE GAUSS
' ==============================================================================

Public Function Geo_Area_Gauss(arrX As Variant, arrY As Variant) As Double
    '----------------------------------------------------------------------------------
    ' Calcula área de polígono usando método de Gauss (Shoelace Formula)
    ' Válido para polígonos fechados em coordenadas planas (UTM ou Topocêntricas)
    '----------------------------------------------------------------------------------
    Dim i As Long, N As Long
    Dim area As Double

    On Error GoTo ErroCalculo

    N = UBound(arrX)
    If N <> UBound(arrY) Then
        Geo_Area_Gauss = 0
        Exit Function
    End If

    area = 0
    For i = 1 To N - 1
        area = area + (arrX(i) * arrY(i + 1) - arrX(i + 1) * arrY(i))
    Next i

    ' Fecha o polígono (último ponto para o primeiro)
    area = area + (arrX(N) * arrY(1) - arrX(1) * arrY(N))

    Geo_Area_Gauss = Abs(area) / 2
    Exit Function

ErroCalculo:
    Geo_Area_Gauss = 0
End Function

' ==============================================================================
' CALCULOS GEODESICOS - COORDENADAS GEOGRAFICAS (SGL)
' ==============================================================================

Public Function Geo_Azimute_Puissant(lat1 As Double, lon1 As Double, lat2 As Double, lon2 As Double) As Double
    '----------------------------------------------------------------------------------
    ' Azimute Geodésico pela Fórmula de Puissant
    ' Mais preciso que azimute plano para coordenadas geográficas
    '----------------------------------------------------------------------------------
    Dim dLon As Double, dLat As Double
    Dim latMed As Double
    Dim azimute As Double

    dLon = (lon2 - lon1) * CONST_PI / 180
    dLat = (lat2 - lat1) * CONST_PI / 180
    latMed = (lat1 + lat2) / 2 * CONST_PI / 180

    Dim x As Double, y As Double
    x = dLon * Cos(latMed)
    y = dLat

    azimute = Application.WorksheetFunction.Atan2(y, x) * 180 / CONST_PI
    azimute = 90 - azimute

    If azimute < 0 Then azimute = azimute + 360
    If azimute >= 360 Then azimute = azimute - 360

    Geo_Azimute_Puissant = azimute
End Function

Public Function Math_Distancia_Geodesica(lat1 As Double, lon1 As Double, lat2 As Double, lon2 As Double) As Double
    '----------------------------------------------------------------------------------
    ' Distância Geodésica pela Fórmula de Haversine
    ' Considera a curvatura da Terra (esférica)
    '----------------------------------------------------------------------------------
    Dim R As Double: R = CONST_RAIO_TERRA
    Dim dLat As Double, dLon As Double
    Dim a As Double, C As Double

    dLat = (lat2 - lat1) * CONST_PI / 180
    dLon = (lon2 - lon1) * CONST_PI / 180

    Dim lat1Rad As Double: lat1Rad = lat1 * CONST_PI / 180
    Dim lat2Rad As Double: lat2Rad = lat2 * CONST_PI / 180

    a = Sin(dLat / 2) * Sin(dLat / 2) + _
        Cos(lat1Rad) * Cos(lat2Rad) * Sin(dLon / 2) * Sin(dLon / 2)
    C = 2 * Application.WorksheetFunction.Atan2(Sqr(1 - a), Sqr(a))

    Math_Distancia_Geodesica = R * C
End Function

' ==============================================================================
' TRANSFORMACOES DE COORDENADAS - GEOCENTRICAS E TOPOCENTRICAS
' ==============================================================================

Public Function Geo_Geod_Para_Geoc(ByVal Lon As Double, ByVal Lat As Double, ByVal H As Double) As Type_Geocentrica
    '----------------------------------------------------------------------------------
    ' Converte Geodésicas (Lat, Lon, Alt) para Geocêntricas (X, Y, Z)
    ' Sistema: Cartesiano 3D com origem no centro da Terra
    '----------------------------------------------------------------------------------
    Dim e2 As Double: e2 = 2 * ACHAT - ACHAT * ACHAT
    Dim latRad As Double: latRad = Lat * PI / 180
    Dim lonRad As Double: lonRad = Lon * PI / 180
    Dim N_val As Double: N_val = SEMI_EIXO / Sqr(1 - (e2 * Sin(latRad) ^ 2))

    Dim resultado As Type_Geocentrica
    resultado.x = (N_val + H) * Cos(latRad) * Cos(lonRad)
    resultado.y = (N_val + H) * Cos(latRad) * Sin(lonRad)
    resultado.Z = (N_val * (1 - e2) + H) * Sin(latRad)

    Geo_Geod_Para_Geoc = resultado
End Function

Public Function Geo_Geoc_Para_Topoc(ByVal x As Double, ByVal y As Double, ByVal Z As Double, _
                                     ByVal lon0 As Double, ByVal lat0 As Double, _
                                     ByVal X0 As Double, ByVal Y0 As Double, ByVal Z0 As Double) As Type_Topocentrica
    '----------------------------------------------------------------------------------
    ' Converte Geocêntricas para Topocêntricas (ENU - East, North, Up)
    ' Sistema local: origem em (lat0, lon0), eixos E-N-U
    '----------------------------------------------------------------------------------
    Dim latRad As Double: latRad = lat0 * PI / 180
    Dim lonRad As Double: lonRad = lon0 * PI / 180

    Dim dX As Double: dX = x - X0
    Dim dY As Double: dY = y - Y0
    Dim dZ As Double: dZ = Z - Z0

    Dim resultado As Type_Topocentrica
    resultado.E = -Sin(lonRad) * dX + Cos(lonRad) * dY
    resultado.N = -Sin(latRad) * Cos(lonRad) * dX - Sin(latRad) * Sin(lonRad) * dY + Cos(latRad) * dZ
    resultado.U = Cos(latRad) * Cos(lonRad) * dX + Cos(latRad) * Sin(lonRad) * dY + Sin(latRad) * dZ

    Geo_Geoc_Para_Topoc = resultado
End Function
