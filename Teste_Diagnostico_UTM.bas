Sub Teste_Diagnostico_UTM()
    ' ==============================================================================
    ' DIAGNÓSTICO: Verificar conversão UTM com dados reais
    ' ==============================================================================

    Dim lat As Double, lon As Double
    Dim utmAntigo As Type_UTM, utmNovo As Type_UTM
    Dim fusoAuto As Integer
    Dim resultado As String

    ' Coordenadas do ponto HVZV-P-21400
    ' SGL: -43°35'36,463" (Lon) -22°28'10,230" (Lat)
    lat = -22.469508333
    lon = -43.59346194

    resultado = "=== DIAGNÓSTICO CONVERSÃO UTM ===" & vbCrLf & vbCrLf

    ' Testa detecção de fuso
    fusoAuto = M_Math_Geo.Geo_GetZonaUTM(lon)
    resultado = resultado & "FUSO DETECTADO: " & fusoAuto & vbCrLf & vbCrLf

    ' Testa conversão com Converter_GeoParaUTM
    utmNovo = M_Math_Geo.Converter_GeoParaUTM(lat, lon, 23)
    resultado = resultado & "Converter_GeoParaUTM (Fuso 23):" & vbCrLf
    resultado = resultado & "  Norte: " & Format(utmNovo.Norte, "0.0000") & vbCrLf
    resultado = resultado & "  Leste: " & Format(utmNovo.Leste, "0.0000") & vbCrLf
    resultado = resultado & "  Hemisferio: " & utmNovo.Hemisferio & vbCrLf
    resultado = resultado & vbCrLf

    ' Testa conversão com fuso detectado
    utmNovo = M_Math_Geo.Converter_GeoParaUTM(lat, lon, fusoAuto)
    resultado = resultado & "Converter_GeoParaUTM (Fuso Auto=" & fusoAuto & "):" & vbCrLf
    resultado = resultado & "  Norte: " & Format(utmNovo.Norte, "0.0000") & vbCrLf
    resultado = resultado & "  Leste: " & Format(utmNovo.Leste, "0.0000") & vbCrLf
    resultado = resultado & "  Hemisferio: " & utmNovo.Hemisferio & vbCrLf
    resultado = resultado & vbCrLf

    ' Valores esperados
    resultado = resultado & "VALORES ESPERADOS:" & vbCrLf
    resultado = resultado & "  Norte: 7514524,6000" & vbCrLf
    resultado = resultado & "  Leste: 644711,6600" & vbCrLf
    resultado = resultado & vbCrLf

    ' Diferenças
    Dim diffN As Double, diffE As Double
    diffN = utmNovo.Norte - 7514524.6
    diffE = utmNovo.Leste - 644711.66
    resultado = resultado & "DIFERENÇAS:" & vbCrLf
    resultado = resultado & "  Delta Norte: " & Format(diffN, "0.00") & " m" & vbCrLf
    resultado = resultado & "  Delta Leste: " & Format(diffE, "0.00") & " m" & vbCrLf

    Debug.Print resultado
    MsgBox resultado, vbInformation, "Diagnóstico UTM"
End Sub
