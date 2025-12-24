Attribute VB_Name = "TesteImportacaoModulos"
Option Explicit

Sub Teste_Importacao_Modulos()
    ' ==============================================================================
    ' TESTE RÁPIDO: Verifica se os módulos refatorados foram importados
    ' ==============================================================================

    Dim resultado As String
    resultado = "=== VERIFICAÇÃO DE IMPORTAÇÃO ===" & vbCrLf & vbCrLf

    ' ------------------------------------------------------------------------------
    ' TESTE 1: Função básica existe?
    ' ------------------------------------------------------------------------------
    On Error Resume Next
    Dim test1 As Double
    test1 = M_Utils.Str_DMS_Para_DD("-43.5")

    If Err.Number = 0 Then
        resultado = resultado & "✅ M_Utils.Str_DMS_Para_DD existe" & vbCrLf
        resultado = resultado & "   Resultado: " & test1 & vbCrLf
    Else
        resultado = resultado & "❌ M_Utils.Str_DMS_Para_DD NÃO encontrada!" & vbCrLf
        resultado = resultado & "   Erro: " & Err.Description & vbCrLf
    End If
    Err.Clear
    On Error GoTo 0

    resultado = resultado & vbCrLf

    ' ------------------------------------------------------------------------------
    ' TESTE 2: Nova função existe?
    ' ------------------------------------------------------------------------------
    On Error Resume Next
    Dim test2 As Type_UTM
    test2 = M_Math_Geo.Converter_GeoParaUTM(-22, -43, 23)

    If Err.Number = 0 Then
        resultado = resultado & "✅ M_Math_Geo.Converter_GeoParaUTM existe" & vbCrLf
        resultado = resultado & "   Norte: " & test2.Norte & vbCrLf
        resultado = resultado & "   Leste: " & test2.Leste & vbCrLf
    Else
        resultado = resultado & "❌ M_Math_Geo.Converter_GeoParaUTM NÃO encontrada!" & vbCrLf
        resultado = resultado & "   Erro: " & Err.Description & vbCrLf
        resultado = resultado & vbCrLf
        resultado = resultado & "⚠️ ATENÇÃO: Você importou os módulos refatorados?" & vbCrLf
        resultado = resultado & "   1. Remover M_Utils e M_Math_Geo antigos" & vbCrLf
        resultado = resultado & "   2. Importar M_Utils_REFATORADO.bas (renomear para M_Utils)" & vbCrLf
        resultado = resultado & "   3. Importar M_Math_Geo_REFATORADO.bas (renomear para M_Math_Geo)" & vbCrLf
    End If
    Err.Clear
    On Error GoTo 0

    resultado = resultado & vbCrLf

    ' ------------------------------------------------------------------------------
    ' TESTE 3: Type personalizado existe?
    ' ------------------------------------------------------------------------------
    On Error Resume Next
    Dim test3 As Type_CalculoPonto
    test3.Distancia = 100
    test3.AzimuteDecimal = 45

    If Err.Number = 0 Then
        resultado = resultado & "✅ Type_CalculoPonto existe" & vbCrLf
    Else
        resultado = resultado & "❌ Type_CalculoPonto NÃO encontrado!" & vbCrLf
        resultado = resultado & "   Erro: " & Err.Description & vbCrLf
    End If
    Err.Clear
    On Error GoTo 0

    resultado = resultado & vbCrLf
    resultado = resultado & "================================" & vbCrLf

    Debug.Print resultado
    MsgBox resultado, vbInformation, "Verificação de Importação"

End Sub
