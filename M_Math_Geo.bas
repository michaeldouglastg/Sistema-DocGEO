Attribute VB_Name = "M_Math_Geo"
Option Explicit
' ==============================================================================
' MODULO: M_MATH_GEO
' DESCRICAO: BIBLIOTECA DE CALCULOS GEODESICOS
' ==============================================================================

Public Type Type_UTM
    Norte As Double
    Leste As Double
    fuso As Integer
    Hemisferio As String ' "N" ou "S"
    'Zona As Integer
    Sucesso As Boolean
    'Hemisferio As String
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

Private Const pi As Double = 3.14159265358979
Private Const SEMI_EIXO As Double = 6378137#
Private Const ACHAT As Double = 0.00335281068118

Public Function Converter_GeoParaUTM(ByVal Latitude As Double, ByVal Longitude As Double, ByVal fuso As Integer) As Type_UTM
    '----------------------------------------------------------------------------------
    ' Algoritmo: Transversa de Mercator (Elipsoide WGS84 / SIRGAS2000)
    ' Adaptado para o sistema M_Math_Geo
    '----------------------------------------------------------------------------------
    Dim resultado As Type_UTM
    
    On Error GoTo ErroConversao
    
    ' Constantes do Elipsoide (WGS84/SIRGAS2000)
    Dim k0 As Double: k0 = 0.9996
    Dim a As Double: a = 6378137#
    Dim f As Double: f = 1 / 298.257223563
    Dim e2 As Double: e2 = 2 * f - f ^ 2
    Dim e_linha2 As Double: e_linha2 = e2 / (1 - e2)
    
    ' Otimiza��o: Pi calculado nativamente no VBA � mais r�pido que chamar WorksheetFunction
    Dim pi As Double: pi = 4 * Atn(1)
    
    ' Convers�o para Radianos
    Dim lat_rad As Double: lat_rad = Latitude * pi / 180
    Dim lon_rad As Double: lon_rad = Longitude * pi / 180
    
    ' Meridiano Central do Fuso (MC)
    ' F�rmula: MC = (Fuso * 6) - 183
    Dim lon_cm_rad As Double: lon_cm_rad = ((fuso * 6) - 183) * pi / 180
    
    ' Vari�veis Auxiliares de C�lculo
    Dim N As Double, T As Double, C As Double, A_term As Double, M As Double
    
    N = a / Sqr(1 - e2 * Sin(lat_rad) ^ 2)
    T = Tan(lat_rad) ^ 2
    C = e_linha2 * Cos(lat_rad) ^ 2
    A_term = (lon_rad - lon_cm_rad) * Cos(lat_rad)
    
    ' C�lculo do Arco do Meridiano (M)
    M = a * ((1 - e2 / 4 - 3 * e2 ^ 2 / 64 - 5 * e2 ^ 3 / 256) * lat_rad _
        - (3 * e2 / 8 + 3 * e2 ^ 2 / 32 + 45 * e2 ^ 3 / 1024) * Sin(2 * lat_rad) _
        + (15 * e2 ^ 2 / 256 + 45 * e2 ^ 3 / 1024) * Sin(4 * lat_rad) _
        - (35 * e2 ^ 3 / 3072) * Sin(6 * lat_rad))
    
    ' --- C�lculo Final Leste (E) ---
    resultado.Leste = k0 * N * (A_term + (1 - T + C) * A_term ^ 3 / 6 + _
                      (5 - 18 * T + T ^ 2 + 72 * C - 58 * e_linha2) * A_term ^ 5 / 120) + 500000
    
    ' --- C�lculo Final Norte (N) ---
    resultado.Norte = k0 * (M + N * Tan(lat_rad) * (A_term ^ 2 / 2 + _
                      (5 - T + 9 * C + 4 * C ^ 2) * A_term ^ 4 / 24 + _
                      (61 - 58 * T + T ^ 2 + 600 * C - 330 * e_linha2) * A_term ^ 6 / 720))
    
    ' Corre��o para Hemisf�rio Sul
    If Latitude < 0 Then
        resultado.Norte = resultado.Norte + 10000000
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

Public Function Calc_Fuso_From_Lon(ByVal Longitude As Double) As Integer
    '----------------------------------------------------------------------------------
    ' Calcula o fuso UTM baseado na Longitude WGS84
    ' F�rmula: Fuso = ParteInteira((Longitude + 180) / 6) + 1
    '----------------------------------------------------------------------------------
    Dim fuso As Integer
    
    ' Garante que a longitude esteja no intervalo [-180, 180]
    If Longitude > 180 Then Longitude = Longitude - 360
    If Longitude < -180 Then Longitude = Longitude + 360
    
    ' C�lculo do Fuso
    fuso = Int((Longitude + 180) / 6) + 1
    
    Calc_Fuso_From_Lon = fuso
End Function

' ==================================================================================
' FUN��O DE COMPATIBILIDADE (PONTE)
' Mant�m o nome antigo para n�o quebrar o resto do sistema, mas usa a matem�tica nova.
' ==================================================================================
Public Function Geo_LatLon_Para_UTM(ByVal Lat As Double, ByVal Lon As Double) As Type_UTM
    Dim fusoAuto As Integer
    
    ' 1. Calcula o fuso automaticamente (usando a fun��o auxiliar que criamos)
    fusoAuto = Calc_Fuso_From_Lon(Lon)
    
    ' 2. Chama a NOVA fun��o robusta passando o fuso calculado
    Geo_LatLon_Para_UTM = Converter_GeoParaUTM(Lat, Lon, fusoAuto)
    
End Function
                                     
'Public Function Geo_LatLon_Para_UTM(ByVal latitude As Double, ByVal longitude As Double) As Type_UTM
Public Function Geo_LatLon_Para_UTM0(ByVal Latitude As Double, ByVal Longitude As Double, Optional ByVal ZonaForcada As Long = 0) As Type_UTM
    
    On Error GoTo ErroCalculo
    
    Dim resultado As Type_UTM
    'Dim zona As Integer
'
'    If ZonaForcada > 0 Then
'        zona = ZonaForcada
'    Else
'        zona = Geo_GetZonaUTM(longitude)
'    End If
'    resultado.zona = zona
    
    Dim Zona As Long
    If ZonaForcada > 0 Then
        Zona = ZonaForcada
    Else
        Zona = Geo_GetZonaUTM(Longitude)
    End If

    
    Dim a As Double: a = 6378137#
    Dim f As Double: f = 1 / 298.257222101
    Dim k0 As Double: k0 = 0.9996
    Dim e2 As Double: e2 = 2 * f - f * f
    Dim e_linha2 As Double: e_linha2 = e2 / (1 - e2)
    
    Dim latRad As Double: latRad = Latitude * pi / 180
    Dim lonRad As Double: lonRad = Longitude * pi / 180
    Dim lon0 As Double: lon0 = (Zona * 6 - 183) * pi / 180
    
    Dim N_radius As Double, T As Double, C As Double, A_term As Double, M As Double
    N_radius = a / Sqr(1 - e2 * Sin(latRad) ^ 2)
    T = Tan(latRad) ^ 2
    C = e_linha2 * Cos(latRad) ^ 2
    A_term = (lonRad - lon0) * Cos(latRad)
    
    M = a * ((1 - e2 / 4 - 3 * e2 ^ 2 / 64 - 5 * e2 ^ 3 / 256) * latRad _
           - (3 * e2 / 8 + 3 * e2 ^ 2 / 32 + 45 * e2 ^ 3 / 1024) * Sin(2 * latRad) _
           + (15 * e2 ^ 2 / 256 + 45 * e2 ^ 3 / 1024) * Sin(4 * latRad) _
           - (35 * e2 ^ 3 / 3072) * Sin(6 * latRad))
    
    resultado.Leste = k0 * N_radius * (A_term + (1 - T + C) * A_term ^ 3 / 6 _
                    + (5 - 18 * T + T ^ 2 + 72 * C - 58 * e_linha2) * A_term ^ 5 / 120) + 500000
    
    resultado.Norte = k0 * (M + N_radius * Tan(latRad) * (A_term ^ 2 / 2 _
                    + (5 - T + 9 * C + 4 * C ^ 2) * A_term ^ 4 / 24 _
                    + (61 - 58 * T + T ^ 2 + 600 * C - 330 * e_linha2) * A_term ^ 6 / 720))
    
    If Latitude < 0 Then resultado.Norte = resultado.Norte + 10000000
    
    resultado.Sucesso = True
    Geo_LatLon_Para_UTM = resultado
    Exit Function
    
ErroCalculo:
    resultado.Sucesso = False
    Geo_LatLon_Para_UTM = resultado
End Function

' ==============================================================================
' CONVERTER UTM PARA LAT/LON (SIRGAS 2000)
' ==============================================================================
Public Function Geo_UTM_Para_LatLon(ByVal Norte As Double, ByVal Este As Double, _
                                     ByVal Zona As Long, ByVal hemisferioSul As Boolean) As Object
    Dim resultado As Object
    Set resultado = CreateObject("Scripting.Dictionary")
    
    ' Constantes SIRGAS 2000 / WGS84
    Dim a As Double: a = 6378137#
    Dim f As Double: f = 1 / 298.257222101
    Dim k0 As Double: k0 = 0.9996
    Dim E0 As Double: E0 = 500000
    Dim N0 As Double
    
    If hemisferioSul Then
        N0 = 10000000
    Else
        N0 = 0
    End If
    
    ' Calculos auxiliares
    Dim e2 As Double: e2 = 2 * f - f * f
    Dim E1 As Double: E1 = (1 - Sqr(1 - e2)) / (1 + Sqr(1 - e2))
    Dim lonOrigem As Double: lonOrigem = (Zona - 1) * 6 - 180 + 3
    
    Dim x As Double: x = Este - E0
    Dim y As Double: y = Norte - N0
    
    Dim M As Double: M = y / k0
    Dim mu As Double: mu = M / (a * (1 - e2 / 4 - 3 * e2 * e2 / 64 - 5 * e2 * e2 * e2 / 256))
    
    ' Latitude do ponto de pe
    Dim phi1 As Double
    phi1 = mu + (3 * E1 / 2 - 27 * E1 * E1 * E1 / 32) * Sin(2 * mu) _
             + (21 * E1 * E1 / 16 - 55 * E1 * E1 * E1 * E1 / 32) * Sin(4 * mu) _
             + (151 * E1 * E1 * E1 / 96) * Sin(6 * mu) _
             + (1097 * E1 * E1 * E1 * E1 / 512) * Sin(8 * mu)
    
    ' Raios de curvatura
    Dim N_val As Double: N_val = a / Sqr(1 - e2 * Sin(phi1) * Sin(phi1))
    Dim T As Double: T = Tan(phi1) * Tan(phi1)
    Dim C As Double: C = (e2 / (1 - e2)) * Cos(phi1) * Cos(phi1)
    Dim R As Double: R = a * (1 - e2) / ((1 - e2 * Sin(phi1) * Sin(phi1)) ^ 1.5)
    Dim D As Double: D = x / (N_val * k0)
    
    ' Latitude final
    Dim Lat As Double
    Lat = phi1 - (N_val * Tan(phi1) / R) * (D * D / 2 _
          - (5 + 3 * T + 10 * C - 4 * C * C - 9 * (e2 / (1 - e2))) * D * D * D * D / 24 _
          + (61 + 90 * T + 298 * C + 45 * T * T - 252 * (e2 / (1 - e2)) - 3 * C * C) * D * D * D * D * D * D / 720)
    
    ' Longitude final
    Dim Lon As Double
    Lon = lonOrigem + (D - (1 + 2 * T + C) * D * D * D / 6 _
          + (5 - 2 * C + 28 * T - 3 * C * C + 8 * (e2 / (1 - e2)) + 24 * T * T) * D * D * D * D * D / 120) / Cos(phi1)
    
    ' Converter para graus
    Lat = Lat * 180 / CONST_PI
    Lon = Lon * 180 / CONST_PI
    
    resultado.Add "Latitude", Lat
    resultado.Add "Longitude", Lon
    
    Set Geo_UTM_Para_LatLon = resultado
End Function

Public Function Geo_GetZonaUTM(Longitude As Double) As Integer
    Geo_GetZonaUTM = Int((Longitude + 180) / 6) + 1
End Function

Public Function Geo_Area_Gauss(arrX As Variant, arrY As Variant) As Double
    Dim i As Long, N As Long
    Dim area As Double
    
    N = UBound(arrX)
    If N <> UBound(arrY) Then Exit Function
    
    area = 0
    For i = 1 To N - 1
        area = area + (arrX(i) * arrY(i + 1) - arrX(i + 1) * arrY(i))
    Next i
    
    area = area + (arrX(N) * arrY(1) - arrX(1) * arrY(N))
    Geo_Area_Gauss = Abs(area) / 2
End Function

Public Function Math_Distancia_Euclidiana(x1 As Double, y1 As Double, x2 As Double, y2 As Double) As Double
    Math_Distancia_Euclidiana = Sqr((x2 - x1) ^ 2 + (y2 - y1) ^ 2)
End Function

Public Function Geo_Azimute_Plano(x1 As Double, y1 As Double, x2 As Double, y2 As Double) As Double
    Dim deltaX As Double, deltaY As Double
    Dim azimute As Double
    
    deltaX = x2 - x1
    deltaY = y2 - y1
    
    If Abs(deltaX) < 0.000001 And Abs(deltaY) < 0.000001 Then
        Geo_Azimute_Plano = 0
        Exit Function
    End If
    
    On Error Resume Next
    azimute = Application.WorksheetFunction.Atan2(deltaY, deltaX) * 180 / pi
    If Err.Number <> 0 Then azimute = 0
    On Error GoTo 0
    
    If azimute < 0 Then azimute = azimute + 360
    Geo_Azimute_Plano = azimute
End Function

Public Function Geo_Geod_Para_Geoc(ByVal Lon As Double, ByVal Lat As Double, ByVal H As Double) As Type_Geocentrica
    Dim e2 As Double: e2 = 2 * ACHAT - ACHAT * ACHAT
    Dim latRad As Double: latRad = Lat * pi / 180
    Dim lonRad As Double: lonRad = Lon * pi / 180
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
    Dim latRad As Double: latRad = lat0 * pi / 180
    Dim lonRad As Double: lonRad = lon0 * pi / 180
    
    Dim dX As Double: dX = x - X0
    Dim dY As Double: dY = y - Y0
    Dim dZ As Double: dZ = Z - Z0
    
    Dim resultado As Type_Topocentrica
    resultado.E = -Sin(lonRad) * dX + Cos(lonRad) * dY
    resultado.N = -Sin(latRad) * Cos(lonRad) * dX - Sin(latRad) * Sin(lonRad) * dY + Cos(latRad) * dZ
    resultado.U = Cos(latRad) * Cos(lonRad) * dX + Cos(latRad) * Sin(lonRad) * dY + Sin(latRad) * dZ
    
    Geo_Geoc_Para_Topoc = resultado
End Function

' ==============================================================================
' AZIMUTE GEODESICO (PUISSANT) - PARA COORDENADAS SGL
' ==============================================================================
Public Function Geo_Azimute_Puissant(lat1 As Double, lon1 As Double, lat2 As Double, lon2 As Double) As Double
    Dim dLon As Double, dLat As Double
    Dim latMed As Double
    Dim azimute As Double
    
    dLon = (lon2 - lon1) * CONST_PI / 180
    dLat = (lat2 - lat1) * CONST_PI / 180
    latMed = (lat1 + lat2) / 2 * CONST_PI / 180
    
    Dim x As Double, y As Double
    x = dLon * Cos(latMed)
    y = dLat

    ' Excel VBA: ATAN2(x, y) = atan2(y, x) matemático
    ' Para azimute: Az = 90° - atan2(y, x) = 90° - ATAN2(x, y)
    ' Mas testando, ATAN2(y, x) funciona melhor - verificar convenção
    azimute = Application.WorksheetFunction.Atan2(y, x) * 180 / CONST_PI
    azimute = 90 - azimute
    
    If azimute < 0 Then azimute = azimute + 360
    If azimute >= 360 Then azimute = azimute - 360
    
    Geo_Azimute_Puissant = azimute
End Function

' ==============================================================================
' DISTANCIA GEODESICA (HAVERSINE) - PARA COORDENADAS SGL
' ==============================================================================
Public Function Math_Distancia_Geodesica(lat1 As Double, lon1 As Double, lat2 As Double, lon2 As Double) As Double
    Dim R As Double: R = 6371000 ' Raio da Terra em metros
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
