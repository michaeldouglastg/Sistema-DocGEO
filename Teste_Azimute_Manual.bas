Sub Teste_Azimute_Manual()
    ' ==============================================================================
    ' TESTE: Calcula azimute manualmente para comparar
    ' ==============================================================================

    Dim resultado As String
    resultado = "=== TESTE DE AZIMUTE MANUAL ===" & vbCrLf & vbCrLf

    ' Coordenadas obtidas
    Dim N1 As Double, E1 As Double, N2 As Double, E2 As Double
    N1 = 7514524.6
    E1 = 644711.65  ' Valor obtido
    N2 = 7514523.79
    E2 = 644712.84

    Dim DeltaN As Double, DeltaE As Double
    DeltaN = N2 - N1  ' = -0.81
    DeltaE = E2 - E1  ' = 1.19

    resultado = resultado & "COORDENADAS OBTIDAS:" & vbCrLf
    resultado = resultado & "  P1: N=" & N1 & " E=" & E1 & vbCrLf
    resultado = resultado & "  P2: N=" & N2 & " E=" & E2 & vbCrLf
    resultado = resultado & "  DeltaN: " & DeltaN & vbCrLf
    resultado = resultado & "  DeltaE: " & DeltaE & vbCrLf
    resultado = resultado & vbCrLf

    ' Calcula azimute
    Dim anguloRad As Double, anguloGraus As Double, azimute As Double
    anguloRad = Application.WorksheetFunction.Atan2(Abs(DeltaN), Abs(DeltaE))
    anguloGraus = anguloRad * 180 / (4 * Atn(1))

    ' Quadrante: DeltaE > 0 e DeltaN < 0 = SE (2º quadrante)
    azimute = 180 - anguloGraus

    resultado = resultado & "CÁLCULO DO AZIMUTE (valores obtidos):" & vbCrLf
    resultado = resultado & "  Ângulo base: " & Format(anguloGraus, "0.000") & "°" & vbCrLf
    resultado = resultado & "  Quadrante: SE (2º)" & vbCrLf
    resultado = resultado & "  Azimute: " & Format(azimute, "0.000") & "°" & vbCrLf
    resultado = resultado & "  Azimute GMS: " & M_Utils.Str_FormatAzimuteGMS(azimute) & vbCrLf
    resultado = resultado & vbCrLf

    ' Agora com coordenadas esperadas
    E1 = 644711.66  ' Valor esperado
    E2 = 644712.85  ' Ajustado proporcionalmente
    N2 = 7514523.8  ' Valor esperado

    DeltaN = N2 - N1
    DeltaE = E2 - E1

    resultado = resultado & "COORDENADAS ESPERADAS:" & vbCrLf
    resultado = resultado & "  P1: N=" & N1 & " E=" & E1 & vbCrLf
    resultado = resultado & "  P2: N=" & N2 & " E=" & E2 & vbCrLf
    resultado = resultado & "  DeltaN: " & DeltaN & vbCrLf
    resultado = resultado & "  DeltaE: " & DeltaE & vbCrLf
    resultado = resultado & vbCrLf

    anguloRad = Application.WorksheetFunction.Atan2(Abs(DeltaN), Abs(DeltaE))
    anguloGraus = anguloRad * 180 / (4 * Atn(1))
    azimute = 180 - anguloGraus

    resultado = resultado & "CÁLCULO DO AZIMUTE (valores esperados):" & vbCrLf
    resultado = resultado & "  Ângulo base: " & Format(anguloGraus, "0.000") & "°" & vbCrLf
    resultado = resultado & "  Azimute: " & Format(azimute, "0.000") & "°" & vbCrLf
    resultado = resultado & "  Azimute GMS: " & M_Utils.Str_FormatAzimuteGMS(azimute) & vbCrLf
    resultado = resultado & "  Esperado: 123°54'42""" & vbCrLf
    resultado = resultado & vbCrLf

    resultado = resultado & "================================" & vbCrLf
    resultado = resultado & "CONCLUSÃO:" & vbCrLf
    resultado = resultado & vbCrLf
    resultado = resultado & "Diferenças de 1cm nas coordenadas UTM causam" & vbCrLf
    resultado = resultado & "diferenças de ~18' no azimute." & vbCrLf
    resultado = resultado & vbCrLf
    resultado = resultado & "Para obter azimutes exatos, precisamos de" & vbCrLf
    resultado = resultado & "coordenadas UTM exatas (mesmo milímetro)." & vbCrLf

    MsgBox resultado, vbInformation, "Teste de Azimute"
End Sub
