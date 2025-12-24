Attribute VB_Name = "VerificarModulosImportados"
Option Explicit

Sub Verificar_Modulos_Importados()
    ' ==============================================================================
    ' VERIFICA QUAIS MÓDULOS ESTÃO IMPORTADOS NO EXCEL
    ' ==============================================================================

    Dim resultado As String
    resultado = "=== VERIFICAÇÃO DE MÓDULOS ===" & vbCrLf & vbCrLf

    ' Verifica se as funções refatoradas existem
    On Error Resume Next

    ' Teste 1: Str_FormatAzimuteGMS (só existe no refatorado)
    Dim testAz As String
    testAz = M_Utils.Str_FormatAzimuteGMS(123.9117)
    If Err.Number = 0 Then
        resultado = resultado & "✅ M_Utils.Str_FormatAzimuteGMS() EXISTE" & vbCrLf
        resultado = resultado & "   Resultado: " & testAz & vbCrLf
    Else
        resultado = resultado & "❌ M_Utils.Str_FormatAzimuteGMS() NÃO EXISTE" & vbCrLf
        resultado = resultado & "   Erro: " & Err.Description & vbCrLf
    End If
    Err.Clear
    resultado = resultado & vbCrLf

    ' Teste 2: Calcular_DistanciaAzimute_UTM (só existe no refatorado)
    Dim testCalc As Type_CalculoPonto
    testCalc = M_Math_Geo.Calcular_DistanciaAzimute_UTM(0, 0, 100, 100)
    If Err.Number = 0 Then
        resultado = resultado & "✅ M_Math_Geo.Calcular_DistanciaAzimute_UTM() EXISTE" & vbCrLf
        resultado = resultado & "   Azimute: " & Format(testCalc.AzimuteDecimal, "0.00") & "°" & vbCrLf
        resultado = resultado & "   Distância: " & Format(testCalc.Distancia, "0.00") & " m" & vbCrLf
    Else
        resultado = resultado & "❌ M_Math_Geo.Calcular_DistanciaAzimute_UTM() NÃO EXISTE" & vbCrLf
        resultado = resultado & "   Erro: " & Err.Description & vbCrLf
    End If
    Err.Clear
    resultado = resultado & vbCrLf

    ' Teste 3: Converter_GeoParaUTM (existe em ambos)
    Dim testUTM As Type_UTM
    testUTM = M_Math_Geo.Converter_GeoParaUTM(-22.469508, -43.593462, 23)
    If Err.Number = 0 Then
        resultado = resultado & "✅ M_Math_Geo.Converter_GeoParaUTM() EXISTE" & vbCrLf
        resultado = resultado & "   Norte: " & Format(testUTM.Norte, "0.0000") & vbCrLf
        resultado = resultado & "   Leste: " & Format(testUTM.Leste, "0.0000") & vbCrLf
        resultado = resultado & vbCrLf
        resultado = resultado & "   ESPERADO:" & vbCrLf
        resultado = resultado & "   Norte: 7514524,6000" & vbCrLf
        resultado = resultado & "   Leste: 644711,6600" & vbCrLf
        resultado = resultado & vbCrLf
        Dim diffN As Double, diffE As Double
        diffN = testUTM.Norte - 7514524.6
        diffE = testUTM.Leste - 644711.66
        resultado = resultado & "   DIFERENÇAS:" & vbCrLf
        resultado = resultado & "   Delta Norte: " & Format(diffN, "0.00") & " m"
        If Abs(diffN) > 100 Then resultado = resultado & " ❌ GRANDE!"
        resultado = resultado & vbCrLf
        resultado = resultado & "   Delta Leste: " & Format(diffE, "0.00") & " m"
        If Abs(diffE) > 100 Then resultado = resultado & " ❌ GRANDE!"
        resultado = resultado & vbCrLf
    Else
        resultado = resultado & "❌ M_Math_Geo.Converter_GeoParaUTM() NÃO EXISTE" & vbCrLf
        resultado = resultado & "   Erro: " & Err.Description & vbCrLf
    End If
    Err.Clear
    resultado = resultado & vbCrLf

    ' Conclusão
    resultado = resultado & "================================" & vbCrLf
    resultado = resultado & "CONCLUSÃO:" & vbCrLf
    resultado = resultado & vbCrLf
    resultado = resultado & "Se alguma função NÃO EXISTE:" & vbCrLf
    resultado = resultado & "→ Você precisa importar os módulos refatorados" & vbCrLf
    resultado = resultado & vbCrLf
    resultado = resultado & "Se a diferença é GRANDE (>100m):" & vbCrLf
    resultado = resultado & "→ O módulo M_Math_Geo ANTIGO está sendo usado" & vbCrLf
    resultado = resultado & "→ Remova M_Math_Geo antigo e importe o refatorado" & vbCrLf

    Debug.Print resultado
    MsgBox resultado, vbInformation, "Verificação de Módulos"
End Sub
