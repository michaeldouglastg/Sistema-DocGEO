Attribute VB_Name = "Teste_Debug_Convergencia"
Option Explicit

' ==============================================================================
' TESTE DE DEBUG - CONVERGÊNCIA MERIDIANA
' ==============================================================================

Public Sub DEBUG_TestarConvergenciaPrimeiroSegmento()
    ' Dados do primeiro segmento HVZV-P-21400 → HVZV-P-21401
    Dim N1 As Double, E1 As Double
    Dim N2 As Double, E2 As Double
    Dim fusoUTM As Integer
    Dim hemisferio As String
    Dim resultado As String

    ' Coordenadas fornecidas
    E1 = 644711.65
    N1 = 7514524.6
    E2 = 644712.84
    N2 = 7514523.79

    ' IMPORTANTE: Ajuste conforme seu caso
    fusoUTM = 23  ' Ajustar conforme necessário (18-25)
    hemisferio = "S"  ' Sul ou Norte

    resultado = "=== DEBUG: CONVERGÊNCIA MERIDIANA ===" & vbCrLf & vbCrLf

    ' 1. Calcular azimute de grid (sem correção)
    Dim calc As Type_CalculoPonto
    calc = M_Math_Geo.Calcular_DistanciaAzimute_UTM(N1, E1, N2, E2)

    resultado = resultado & "1. AZIMUTE DE GRID (sem correção):" & vbCrLf
    resultado = resultado & "   Decimal: " & Format(calc.AzimuteDecimal, "0.000000") & "°" & vbCrLf
    resultado = resultado & "   DMS: " & M_Utils.Str_FormatAzimuteGMS(calc.AzimuteDecimal) & vbCrLf & vbCrLf

    ' 2. Converter UTM → Geo
    Dim geoAtual As Type_Geo
    geoAtual = M_Math_Geo.Converter_UTMParaGeo(N1, E1, fusoUTM, hemisferio)

    resultado = resultado & "2. COORDENADAS GEOGRÁFICAS (ponto inicial):" & vbCrLf
    resultado = resultado & "   Latitude: " & Format(geoAtual.Latitude, "0.000000") & "° "
    If geoAtual.Latitude < 0 Then resultado = resultado & "(Sul)" Else resultado = resultado & "(Norte)"
    resultado = resultado & vbCrLf
    resultado = resultado & "   Longitude: " & Format(geoAtual.Longitude, "0.000000") & "° "
    If geoAtual.Longitude < 0 Then resultado = resultado & "(Oeste)" Else resultado = resultado & "(Leste)"
    resultado = resultado & vbCrLf & vbCrLf

    ' 3. Calcular meridiano central
    Dim meridianoCentral As Double
    meridianoCentral = (fusoUTM * 6) - 183

    resultado = resultado & "3. FUSO UTM E MERIDIANO CENTRAL:" & vbCrLf
    resultado = resultado & "   Fuso: " & fusoUTM & vbCrLf
    resultado = resultado & "   Meridiano Central: " & meridianoCentral & "°" & vbCrLf
    resultado = resultado & "   ΔLon = Lon - λ0 = " & Format(geoAtual.Longitude - meridianoCentral, "0.000000") & "°" & vbCrLf & vbCrLf

    ' 4. Calcular convergência meridiana
    Dim convergencia As Double
    convergencia = M_Math_Geo.Calcular_ConvergenciaMeridiana(geoAtual.Latitude, geoAtual.Longitude, fusoUTM)

    resultado = resultado & "4. CONVERGÊNCIA MERIDIANA (γ):" & vbCrLf
    resultado = resultado & "   Decimal: " & Format(convergencia, "0.000000") & "°" & vbCrLf
    resultado = resultado & "   DMS: " & M_Utils.Str_DD_Para_DMS(convergencia) & vbCrLf
    If convergencia > 0 Then
        resultado = resultado & "   (Positiva - adiciona ao azimute de grid)" & vbCrLf
    Else
        resultado = resultado & "   (Negativa - subtrai do azimute de grid)" & vbCrLf
    End If
    resultado = resultado & vbCrLf

    ' 5. Aplicar correção
    Dim azimuteGeod As Double
    azimuteGeod = M_Math_Geo.Converter_AzimuteGridParaGeod(calc.AzimuteDecimal, geoAtual.Latitude, geoAtual.Longitude, fusoUTM)

    resultado = resultado & "5. AZIMUTE GEODÉSICO (com correção):" & vbCrLf
    resultado = resultado & "   Fórmula: Az_Geod = Az_Grid + γ" & vbCrLf
    resultado = resultado & "   " & Format(calc.AzimuteDecimal, "0.000000") & "° + " & Format(convergencia, "0.000000") & "° = " & Format(azimuteGeod, "0.000000") & "°" & vbCrLf
    resultado = resultado & "   DMS: " & M_Utils.Str_FormatAzimuteGMS(azimuteGeod) & vbCrLf & vbCrLf

    ' 6. Comparação
    resultado = resultado & "=== COMPARAÇÃO ===" & vbCrLf
    resultado = resultado & "Azimute de Grid: " & M_Utils.Str_FormatAzimuteGMS(calc.AzimuteDecimal) & vbCrLf
    resultado = resultado & "Azimute Geodésico: " & M_Utils.Str_FormatAzimuteGMS(azimuteGeod) & vbCrLf
    resultado = resultado & "Azimute Esperado: 123°54'42""" & vbCrLf & vbCrLf

    ' 7. Informações adicionais
    resultado = resultado & "=== INFORMAÇÕES IMPORTANTES ===" & vbCrLf
    resultado = resultado & "• Se o ponto está a OESTE do meridiano central:" & vbCrLf
    resultado = resultado & "  - No Hemisfério SUL: γ > 0 (positiva)" & vbCrLf
    resultado = resultado & "  - No Hemisfério NORTE: γ < 0 (negativa)" & vbCrLf
    resultado = resultado & "• Se o ponto está a LESTE do meridiano central:" & vbCrLf
    resultado = resultado & "  - No Hemisfério SUL: γ < 0 (negativa)" & vbCrLf
    resultado = resultado & "  - No Hemisfério NORTE: γ > 0 (positiva)" & vbCrLf & vbCrLf
    resultado = resultado & "• Seu ponto (E=" & E1 & "):"& vbCrLf
    If E1 < 500000 Then
        resultado = resultado & "  - Está a OESTE do meridiano central (E < 500000)" & vbCrLf
    Else
        resultado = resultado & "  - Está a LESTE do meridiano central (E > 500000)" & vbCrLf
    End If

    ' Exibir resultado
    Debug.Print resultado

    ' Salvar em arquivo
    Dim fso As Object, arquivo As Object
    Dim caminhoArquivo As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    If ThisWorkbook.Path <> "" Then
        caminhoArquivo = ThisWorkbook.Path & "\DEBUG_Convergencia_Resultado.txt"
    Else
        caminhoArquivo = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\DEBUG_Convergencia_Resultado.txt"
    End If

    Set arquivo = fso.CreateTextFile(caminhoArquivo, True)
    If Err.Number = 0 Then
        arquivo.WriteLine resultado
        arquivo.Close
        MsgBox "Resultado salvo em:" & vbCrLf & caminhoArquivo, vbInformation, "Debug Concluído"
    Else
        MsgBox resultado, vbInformation, "Debug - Veja também a janela Immediate (Ctrl+G)"
    End If
    On Error GoTo 0

    Set fso = Nothing
End Sub
