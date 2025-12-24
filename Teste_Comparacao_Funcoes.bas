Attribute VB_Name = "TesteComparacaoFuncoes"
Option Explicit

Sub Teste_Comparacao_Funcoes()
    ' ==============================================================================
    ' TESTE: Compara funções antigas vs novas
    ' ==============================================================================

    Dim resultado As String
    resultado = "=== COMPARAÇÃO FUNÇÕES ANTIGAS VS NOVAS ===" & vbCrLf & vbCrLf

    ' Coordenadas de teste (Rio de Janeiro, aproximado)
    Dim lat As Double: lat = -22.469508
    Dim lon As Double: lon = -43.593461
    Dim fuso As Integer: fuso = 23

    resultado = resultado & "Coordenadas de teste:" & vbCrLf
    resultado = resultado & "  Latitude: " & lat & vbCrLf
    resultado = resultado & "  Longitude: " & lon & vbCrLf
    resultado = resultado & "  Fuso: " & fuso & vbCrLf & vbCrLf

    ' ------------------------------------------------------------------------------
    ' TESTE 1: Função ANTIGA (Geo_LatLon_Para_UTM)
    ' ------------------------------------------------------------------------------
    resultado = resultado & "FUNÇÃO ANTIGA (Geo_LatLon_Para_UTM):" & vbCrLf

    On Error Resume Next
    Dim utmAntigo As Type_UTM
    utmAntigo = M_Math_Geo.Geo_LatLon_Para_UTM(lat, lon, fuso)

    If Err.Number = 0 Then
        resultado = resultado & "  Norte: " & utmAntigo.Norte & vbCrLf
        resultado = resultado & "  Leste: " & utmAntigo.Leste & vbCrLf
        resultado = resultado & "  Hemisferio: " & utmAntigo.Hemisferio & vbCrLf
        resultado = resultado & "  Sucesso: " & utmAntigo.Sucesso & vbCrLf
    Else
        resultado = resultado & "  ERRO: " & Err.Description & vbCrLf
    End If
    Err.Clear
    On Error GoTo 0

    resultado = resultado & vbCrLf

    ' ------------------------------------------------------------------------------
    ' TESTE 2: Função NOVA (Converter_GeoParaUTM)
    ' ------------------------------------------------------------------------------
    resultado = resultado & "FUNÇÃO NOVA (Converter_GeoParaUTM):" & vbCrLf

    On Error Resume Next
    Dim utmNovo As Type_UTM
    utmNovo = M_Math_Geo.Converter_GeoParaUTM(lat, lon, fuso)

    If Err.Number = 0 Then
        resultado = resultado & "  Norte: " & utmNovo.Norte & vbCrLf
        resultado = resultado & "  Leste: " & utmNovo.Leste & vbCrLf
        resultado = resultado & "  Hemisferio: " & utmNovo.Hemisferio & vbCrLf
        resultado = resultado & "  Sucesso: " & utmNovo.Sucesso & vbCrLf
    Else
        resultado = resultado & "  ERRO: " & Err.Description & vbCrLf
    End If
    Err.Clear
    On Error GoTo 0

    resultado = resultado & vbCrLf

    ' ------------------------------------------------------------------------------
    ' COMPARAÇÃO
    ' ------------------------------------------------------------------------------
    If utmAntigo.Sucesso And utmNovo.Sucesso Then
        resultado = resultado & "DIFERENÇAS:" & vbCrLf
        resultado = resultado & "  Delta Norte: " & Abs(utmNovo.Norte - utmAntigo.Norte) & " m" & vbCrLf
        resultado = resultado & "  Delta Leste: " & Abs(utmNovo.Leste - utmAntigo.Leste) & " m" & vbCrLf

        If Abs(utmNovo.Norte - utmAntigo.Norte) < 0.001 And Abs(utmNovo.Leste - utmAntigo.Leste) < 0.001 Then
            resultado = resultado & vbCrLf & "✅ FUNÇÕES PRODUZEM MESMO RESULTADO!" & vbCrLf
        Else
            resultado = resultado & vbCrLf & "⚠️ DIFERENÇA DETECTADA!" & vbCrLf
        End If
    End If

    resultado = resultado & vbCrLf & "================================" & vbCrLf

    ' ------------------------------------------------------------------------------
    ' TESTE 3: Conversão de String Decimal (CSV)
    ' ------------------------------------------------------------------------------
    resultado = resultado & vbCrLf & "TESTE CONVERSÃO DECIMAL (CSV):" & vbCrLf

    Dim csvString As String: csvString = "-43.5934619399999974"
    Dim lonConvertida As Double

    resultado = resultado & "  String CSV: """ & csvString & """" & vbCrLf

    lonConvertida = M_Utils.Str_DMS_Para_DD(csvString)

    resultado = resultado & "  Valor convertido: " & lonConvertida & vbCrLf
    resultado = resultado & "  Esperado: -43.5934619399999974" & vbCrLf
    resultado = resultado & "  Diferença: " & Abs(lonConvertida - (-43.5934619399999974)) & vbCrLf

    If Abs(lonConvertida - (-43.5934619399999974)) < 0.0000001 Then
        resultado = resultado & "  ✅ PASSOU" & vbCrLf
    Else
        resultado = resultado & "  ❌ FALHOU" & vbCrLf
    End If

    Debug.Print resultado
    MsgBox resultado, vbInformation, "Comparação de Funções"

End Sub
