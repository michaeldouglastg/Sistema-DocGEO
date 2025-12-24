Attribute VB_Name = "TesteFinalRefatoracao"
Option Explicit

Sub Teste_Final_Refatoracao()
    ' ==============================================================================
    ' TESTE FINAL - REFATORAÇÃO VALIDADA
    ' Versão corrigida com valores reais validados
    ' ==============================================================================

    Dim resultado As String
    Dim totalTestes As Integer: totalTestes = 0
    Dim testesPassados As Integer: testesPassados = 0

    resultado = "=== TESTE FINAL DA REFATORACAO ===" & vbCrLf & vbCrLf

    ' --------------------------------------------------------------------------
    ' TESTE 1: Conversão DMS → DD (formato atual)
    ' --------------------------------------------------------------------------
    totalTestes = totalTestes + 1
    resultado = resultado & "TESTE 1: DMS com sinal → DD" & vbCrLf

    Dim test1 As Double
    test1 = M_Utils.Str_DMS_Para_DD("-43°35'36,463""")

    If Abs(test1 - (-43.59346194)) < 0.00001 Then
        resultado = resultado & "  ✅ PASSOU" & vbCrLf
        testesPassados = testesPassados + 1
    Else
        resultado = resultado & "  ❌ FALHOU - Obtido: " & test1 & vbCrLf
    End If
    resultado = resultado & vbCrLf

    ' --------------------------------------------------------------------------
    ' TESTE 2: Conversão DMS com sufixo → DD
    ' --------------------------------------------------------------------------
    totalTestes = totalTestes + 1
    resultado = resultado & "TESTE 2: DMS com sufixo O → DD" & vbCrLf

    Dim test2 As Double
    test2 = M_Utils.Str_DMS_Para_DD("43° 35' 36,4626"" O")

    If Abs(test2 - (-43.59346183)) < 0.00001 Then
        resultado = resultado & "  ✅ PASSOU" & vbCrLf
        testesPassados = testesPassados + 1
    Else
        resultado = resultado & "  ❌ FALHOU - Obtido: " & test2 & vbCrLf
    End If
    resultado = resultado & vbCrLf

    ' --------------------------------------------------------------------------
    ' TESTE 3: Conversão Decimal Puro (CSV SIGEF)
    ' --------------------------------------------------------------------------
    totalTestes = totalTestes + 1
    resultado = resultado & "TESTE 3: Decimal puro (CSV SIGEF)" & vbCrLf

    Dim test3 As Double
    test3 = M_Utils.Str_DMS_Para_DD("-43.5934619399999974")

    If Abs(test3 - (-43.59346194)) < 0.0000001 Then
        resultado = resultado & "  ✅ PASSOU" & vbCrLf
        testesPassados = testesPassados + 1
    Else
        resultado = resultado & "  ❌ FALHOU - Obtido: " & test3 & vbCrLf
    End If
    resultado = resultado & vbCrLf

    ' --------------------------------------------------------------------------
    ' TESTE 4: Conversão Geo → UTM (validado)
    ' --------------------------------------------------------------------------
    totalTestes = totalTestes + 1
    resultado = resultado & "TESTE 4: Geo → UTM (função antiga vs nova)" & vbCrLf

    Dim utmAntigo As Type_UTM, utmNovo As Type_UTM
    utmAntigo = M_Math_Geo.Geo_LatLon_Para_UTM(-22.469508, -43.593461, 23)
    utmNovo = M_Math_Geo.Converter_GeoParaUTM(-22.469508, -43.593461, 23)

    Dim deltaN As Double, deltaE As Double
    deltaN = Abs(utmNovo.Norte - utmAntigo.Norte)
    deltaE = Abs(utmNovo.Leste - utmAntigo.Leste)

    If deltaN < 0.001 And deltaE < 0.001 Then
        resultado = resultado & "  ✅ PASSOU (diferença < 1mm)" & vbCrLf
        testesPassados = testesPassados + 1
    Else
        resultado = resultado & "  ❌ FALHOU" & vbCrLf
        resultado = resultado & "    Delta Norte: " & deltaN & " m" & vbCrLf
        resultado = resultado & "    Delta Leste: " & deltaE & " m" & vbCrLf
    End If
    resultado = resultado & vbCrLf

    ' --------------------------------------------------------------------------
    ' TESTE 5: Azimute Robusto (NE = 45°)
    ' --------------------------------------------------------------------------
    totalTestes = totalTestes + 1
    resultado = resultado & "TESTE 5: Azimute robusto (NE = 45°)" & vbCrLf

    Dim calc As Type_CalculoPonto
    calc = M_Math_Geo.Calcular_DistanciaAzimute_UTM(0, 0, 100, 100)

    If Abs(calc.AzimuteDecimal - 45) < 0.1 Then
        resultado = resultado & "  ✅ PASSOU (Az=" & calc.AzimuteDecimal & "°)" & vbCrLf
        testesPassados = testesPassados + 1
    Else
        resultado = resultado & "  ❌ FALHOU - Obtido: " & calc.AzimuteDecimal & "°" & vbCrLf
    End If
    resultado = resultado & vbCrLf

    ' --------------------------------------------------------------------------
    ' TESTE 6: Azimute em todos os quadrantes
    ' --------------------------------------------------------------------------
    totalTestes = totalTestes + 1
    resultado = resultado & "TESTE 6: Azimute nos 4 quadrantes" & vbCrLf

    Dim azNE As Double, azSE As Double, azSW As Double, azNW As Double
    azNE = M_Math_Geo.Calcular_DistanciaAzimute_UTM(0, 0, 100, 100).AzimuteDecimal   ' 45°
    azSE = M_Math_Geo.Calcular_DistanciaAzimute_UTM(0, 0, -100, 100).AzimuteDecimal  ' 135°
    azSW = M_Math_Geo.Calcular_DistanciaAzimute_UTM(0, 0, -100, -100).AzimuteDecimal ' 225°
    azNW = M_Math_Geo.Calcular_DistanciaAzimute_UTM(0, 0, 100, -100).AzimuteDecimal  ' 315°

    Dim quadrantesOK As Boolean
    quadrantesOK = (Abs(azNE - 45) < 0.1) And _
                   (Abs(azSE - 135) < 0.1) And _
                   (Abs(azSW - 225) < 0.1) And _
                   (Abs(azNW - 315) < 0.1)

    If quadrantesOK Then
        resultado = resultado & "  ✅ PASSOU" & vbCrLf
        resultado = resultado & "    NE=" & azNE & "° SE=" & azSE & "° SW=" & azSW & "° NW=" & azNW & "°" & vbCrLf
        testesPassados = testesPassados + 1
    Else
        resultado = resultado & "  ❌ FALHOU" & vbCrLf
    End If
    resultado = resultado & vbCrLf

    ' --------------------------------------------------------------------------
    ' TESTE 7: Conversão UTM → Geo (ida e volta)
    ' --------------------------------------------------------------------------
    totalTestes = totalTestes + 1
    resultado = resultado & "TESTE 7: UTM → Geo (conversão inversa)" & vbCrLf

    ' Converte Geo → UTM → Geo
    Dim latOriginal As Double: latOriginal = -22.469508
    Dim lonOriginal As Double: lonOriginal = -43.593461

    Dim utmIda As Type_UTM
    utmIda = M_Math_Geo.Converter_GeoParaUTM(latOriginal, lonOriginal, 23)

    Dim geoVolta As Type_Geo
    geoVolta = M_Math_Geo.Converter_UTMParaGeo(utmIda.Norte, utmIda.Leste, 23, "S")

    Dim erroLat As Double, erroLon As Double
    erroLat = Abs(geoVolta.Latitude - latOriginal)
    erroLon = Abs(geoVolta.Longitude - lonOriginal)

    If erroLat < 0.000001 And erroLon < 0.000001 Then
        resultado = resultado & "  ✅ PASSOU (erro < 10cm)" & vbCrLf
        testesPassados = testesPassados + 1
    Else
        resultado = resultado & "  ❌ FALHOU" & vbCrLf
        resultado = resultado & "    Erro Lat: " & erroLat & "°" & vbCrLf
        resultado = resultado & "    Erro Lon: " & erroLon & "°" & vbCrLf
    End If
    resultado = resultado & vbCrLf

    ' --------------------------------------------------------------------------
    ' RESUMO FINAL
    ' --------------------------------------------------------------------------
    resultado = resultado & "================================" & vbCrLf
    resultado = resultado & "RESULTADO FINAL:" & vbCrLf
    resultado = resultado & "  Testes executados: " & totalTestes & vbCrLf
    resultado = resultado & "  Testes passados: " & testesPassados & vbCrLf
    resultado = resultado & "  Taxa de sucesso: " & Format((testesPassados / totalTestes) * 100, "0.0") & "%" & vbCrLf
    resultado = resultado & vbCrLf

    If testesPassados = totalTestes Then
        resultado = resultado & "🎉 TODOS OS TESTES PASSARAM!" & vbCrLf
        resultado = resultado & "✅ REFATORAÇÃO VALIDADA COM SUCESSO!" & vbCrLf
    Else
        resultado = resultado & "⚠️ ALGUNS TESTES FALHARAM" & vbCrLf
        resultado = resultado & "Verifique os módulos importados" & vbCrLf
    End If

    Debug.Print resultado
    MsgBox resultado, vbInformation, "Teste Final da Refatoração"

End Sub
