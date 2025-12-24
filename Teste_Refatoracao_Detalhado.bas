Attribute VB_Name = "TesteRefatoracaoDetalhado"
Option Explicit

Sub Teste_Refatoracao_Detalhado()
    ' ==============================================================================
    ' TESTE DETALHADO DA REFATORAÇÃO
    ' Mostra exatamente qual teste falhou e os valores obtidos
    ' ==============================================================================

    Dim resultado As String
    resultado = "=== TESTES DA REFATORAÇÃO ===" & vbCrLf & vbCrLf

    ' ------------------------------------------------------------------------------
    ' TESTE 1: Conversão DMS → DD (formato atual do sistema)
    ' ------------------------------------------------------------------------------
    resultado = resultado & "TESTE 1: DMS com sinal → DD" & vbCrLf
    Dim test1 As Double
    test1 = M_Utils.Str_DMS_Para_DD("-43°35'36,463""")

    resultado = resultado & "  Entrada: ""-43°35'36,463""""" & vbCrLf
    resultado = resultado & "  Esperado: -43.59346194" & vbCrLf
    resultado = resultado & "  Obtido: " & test1 & vbCrLf
    resultado = resultado & "  Diferença: " & Abs(test1 - (-43.59346194)) & vbCrLf

    If Abs(test1 - (-43.59346194)) < 0.00001 Then
        resultado = resultado & "  ✅ PASSOU" & vbCrLf & vbCrLf
    Else
        resultado = resultado & "  ❌ FALHOU" & vbCrLf & vbCrLf
    End If

    ' ------------------------------------------------------------------------------
    ' TESTE 2: Conversão DMS → DD (formato com sufixo)
    ' ------------------------------------------------------------------------------
    resultado = resultado & "TESTE 2: DMS com sufixo O → DD" & vbCrLf
    Dim test2 As Double
    test2 = M_Utils.Str_DMS_Para_DD("43° 35' 36,4626"" O")

    resultado = resultado & "  Entrada: ""43° 35' 36,4626"""" O""" & vbCrLf
    resultado = resultado & "  Esperado: -43.59346183" & vbCrLf
    resultado = resultado & "  Obtido: " & test2 & vbCrLf
    resultado = resultado & "  Diferença: " & Abs(test2 - (-43.59346183)) & vbCrLf

    If Abs(test2 - (-43.59346183)) < 0.00001 Then
        resultado = resultado & "  ✅ PASSOU" & vbCrLf & vbCrLf
    Else
        resultado = resultado & "  ❌ FALHOU" & vbCrLf & vbCrLf
    End If

    ' ------------------------------------------------------------------------------
    ' TESTE 3: Conversão DD → DMS
    ' ------------------------------------------------------------------------------
    resultado = resultado & "TESTE 3: DD → DMS" & vbCrLf
    Dim test3 As String
    test3 = M_Utils.Str_DD_Para_DMS(-43.593461)

    resultado = resultado & "  Entrada: -43.593461" & vbCrLf
    resultado = resultado & "  Esperado: -43°35'36.458""" & vbCrLf
    resultado = resultado & "  Obtido: " & test3 & vbCrLf

    ' Verifica se contém -43°35'36
    If InStr(test3, "-43°35'36") > 0 Then
        resultado = resultado & "  ✅ PASSOU" & vbCrLf & vbCrLf
    Else
        resultado = resultado & "  ❌ FALHOU" & vbCrLf & vbCrLf
    End If

    ' ------------------------------------------------------------------------------
    ' TESTE 4: Conversão Geo → UTM
    ' ------------------------------------------------------------------------------
    resultado = resultado & "TESTE 4: Geo → UTM" & vbCrLf
    Dim test4 As Type_UTM
    test4 = M_Math_Geo.Converter_GeoParaUTM(-22.469508, -43.593461, 23)

    resultado = resultado & "  Entrada: Lat=-22.469508, Lon=-43.593461, Fuso=23" & vbCrLf
    resultado = resultado & "  Norte esperado: ~7514234 (±100)" & vbCrLf
    resultado = resultado & "  Norte obtido: " & test4.Norte & vbCrLf
    resultado = resultado & "  Leste esperado: ~685432 (±100)" & vbCrLf
    resultado = resultado & "  Leste obtido: " & test4.Leste & vbCrLf
    resultado = resultado & "  Hemisfério: " & test4.Hemisferio & vbCrLf
    resultado = resultado & "  Sucesso: " & test4.Sucesso & vbCrLf

    If test4.Sucesso And Abs(test4.Norte - 7514234) < 100 And Abs(test4.Leste - 685432) < 100 Then
        resultado = resultado & "  ✅ PASSOU" & vbCrLf & vbCrLf
    Else
        resultado = resultado & "  ❌ FALHOU" & vbCrLf & vbCrLf
    End If

    ' ------------------------------------------------------------------------------
    ' TESTE 5: Cálculo de azimute (NE = 45°)
    ' ------------------------------------------------------------------------------
    resultado = resultado & "TESTE 5: Azimute NE (45°)" & vbCrLf
    Dim test5 As Type_CalculoPonto
    test5 = M_Math_Geo.Calcular_DistanciaAzimute_UTM(0, 0, 100, 100)

    resultado = resultado & "  Entrada: (0,0) → (100,100)" & vbCrLf
    resultado = resultado & "  Azimute esperado: 45°" & vbCrLf
    resultado = resultado & "  Azimute obtido: " & test5.AzimuteDecimal & "°" & vbCrLf
    resultado = resultado & "  Distância obtida: " & test5.Distancia & " m" & vbCrLf
    resultado = resultado & "  Diferença: " & Abs(test5.AzimuteDecimal - 45) & vbCrLf

    If Abs(test5.AzimuteDecimal - 45) < 0.1 Then
        resultado = resultado & "  ✅ PASSOU" & vbCrLf & vbCrLf
    Else
        resultado = resultado & "  ❌ FALHOU" & vbCrLf & vbCrLf
    End If

    ' ------------------------------------------------------------------------------
    ' TESTE 6: Conversão decimal puro (CSV SIGEF)
    ' ------------------------------------------------------------------------------
    resultado = resultado & "TESTE 6: Decimal puro (CSV)" & vbCrLf
    Dim test6 As Double
    test6 = M_Utils.Str_DMS_Para_DD("-43.5934619399999974")

    resultado = resultado & "  Entrada: ""-43.5934619399999974""" & vbCrLf
    resultado = resultado & "  Esperado: -43.5934619399999974" & vbCrLf
    resultado = resultado & "  Obtido: " & test6 & vbCrLf
    resultado = resultado & "  Diferença: " & Abs(test6 - (-43.5934619399999974)) & vbCrLf

    If Abs(test6 - (-43.5934619399999974)) < 0.0000001 Then
        resultado = resultado & "  ✅ PASSOU" & vbCrLf & vbCrLf
    Else
        resultado = resultado & "  ❌ FALHOU" & vbCrLf & vbCrLf
    End If

    ' Mostra resultado
    Debug.Print resultado

    ' Cria formulário de mensagem
    Dim frm As Object
    On Error Resume Next
    Set frm = CreateObject("Forms.Form.1")
    On Error GoTo 0

    If frm Is Nothing Then
        ' Se não conseguir criar form, usa MsgBox
        MsgBox resultado, vbInformation, "Resultado dos Testes"
    Else
        ' Usa form para mostrar texto completo
        MsgBox resultado, vbInformation, "Resultado dos Testes"
    End If

End Sub
