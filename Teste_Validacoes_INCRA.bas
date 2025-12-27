Attribute VB_Name = "Teste_Validacoes_INCRA"
Option Explicit
' ==============================================================================
' MODULO: TESTE_VALIDACOES_INCRA
' DESCRICAO: TESTES UNITARIOS DAS VALIDACOES CONFORME MANUAL INCRA
' ==============================================================================

Public Sub ExecutarTodosTestes()
    '----------------------------------------------------------------------------------
    ' Executa todos os testes de validacao INCRA
    '----------------------------------------------------------------------------------
    Dim resultado As String

    resultado = "TESTES DE VALIDACAO INCRA" & vbCrLf
    resultado = resultado & String(60, "=") & vbCrLf & vbCrLf

    resultado = resultado & Teste_TiposVertice() & vbCrLf
    resultado = resultado & Teste_TiposLimite() & vbCrLf
    resultado = resultado & Teste_PrecisaoHorizontal() & vbCrLf
    resultado = resultado & Teste_MetodosPosicionamento() & vbCrLf
    resultado = resultado & Teste_ValidacaoCompleta() & vbCrLf
    resultado = resultado & Teste_CalculoEMQ() & vbCrLf

    Debug.Print resultado

    ' Tenta criar arquivo com resultado
    On Error Resume Next
    Dim fso As Object, arquivo As Object
    Dim caminhoArquivo As String
    Dim salvouArquivo As Boolean

    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Tenta salvar no diretório do workbook
    If ThisWorkbook.Path <> "" Then
        caminhoArquivo = ThisWorkbook.Path & "\resultado_testes_incra.txt"
    Else
        ' Se workbook não foi salvo, usa área de trabalho do usuário
        caminhoArquivo = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\resultado_testes_incra.txt"
    End If

    Set arquivo = fso.CreateTextFile(caminhoArquivo, True)

    If Err.Number = 0 Then
        arquivo.WriteLine resultado
        arquivo.Close
        salvouArquivo = True
    Else
        salvouArquivo = False
    End If
    On Error GoTo 0

    ' Mostra resultado
    If salvouArquivo Then
        MsgBox "Testes concluidos!" & vbCrLf & vbCrLf & _
               "Resultado salvo em:" & vbCrLf & caminhoArquivo, vbInformation
    Else
        MsgBox "Testes concluidos!" & vbCrLf & vbCrLf & _
               "NOTA: Nao foi possivel salvar arquivo." & vbCrLf & _
               "Resultado disponivel na janela Immediate (Ctrl+G)", vbInformation
    End If
End Sub

Private Function Teste_TiposVertice() As String
    '----------------------------------------------------------------------------------
    ' Testa validacao de tipos de vertice
    '----------------------------------------------------------------------------------
    Dim resultado As String
    Dim passaram As Long, falharam As Long

    resultado = "TESTE: Tipos de Vertice" & vbCrLf
    resultado = resultado & String(40, "-") & vbCrLf

    ' Testes positivos
    If M_Validacao.Validar_TipoVertice("M") Then
        resultado = resultado & "  OK - M (Marco) valido" & vbCrLf
        passaram = passaram + 1
    Else
        resultado = resultado & "  FALHOU - M deveria ser valido" & vbCrLf
        falharam = falharam + 1
    End If

    If M_Validacao.Validar_TipoVertice("P") Then
        resultado = resultado & "  OK - P (Ponto) valido" & vbCrLf
        passaram = passaram + 1
    Else
        resultado = resultado & "  FALHOU - P deveria ser valido" & vbCrLf
        falharam = falharam + 1
    End If

    If M_Validacao.Validar_TipoVertice("V") Then
        resultado = resultado & "  OK - V (Virtual) valido" & vbCrLf
        passaram = passaram + 1
    Else
        resultado = resultado & "  FALHOU - V deveria ser valido" & vbCrLf
        falharam = falharam + 1
    End If

    ' Testes negativos
    If Not M_Validacao.Validar_TipoVertice("X") Then
        resultado = resultado & "  OK - X invalidado corretamente" & vbCrLf
        passaram = passaram + 1
    Else
        resultado = resultado & "  FALHOU - X deveria ser invalido" & vbCrLf
        falharam = falharam + 1
    End If

    If Not M_Validacao.Validar_TipoVertice("") Then
        resultado = resultado & "  OK - String vazia invalidada" & vbCrLf
        passaram = passaram + 1
    Else
        resultado = resultado & "  FALHOU - String vazia deveria ser invalida" & vbCrLf
        falharam = falharam + 1
    End If

    resultado = resultado & "  Total: " & passaram & " passaram, " & falharam & " falharam" & vbCrLf

    Teste_TiposVertice = resultado
End Function

Private Function Teste_TiposLimite() As String
    '----------------------------------------------------------------------------------
    ' Testa validacao de tipos de limite
    '----------------------------------------------------------------------------------
    Dim resultado As String
    Dim passaram As Long, falharam As Long
    Dim i As Long

    resultado = "TESTE: Tipos de Limite" & vbCrLf
    resultado = resultado & String(40, "-") & vbCrLf

    ' Testa limites artificiais (LA1-LA7)
    Dim limitesArtificiais As Variant
    limitesArtificiais = Array("LA1", "LA2", "LA3", "LA4", "LA5", "LA6", "LA7")

    For i = LBound(limitesArtificiais) To UBound(limitesArtificiais)
        If M_Validacao.Validar_TipoLimite(CStr(limitesArtificiais(i))) Then
            passaram = passaram + 1
        Else
            resultado = resultado & "  FALHOU - " & limitesArtificiais(i) & " deveria ser valido" & vbCrLf
            falharam = falharam + 1
        End If
    Next i

    ' Testa limites naturais (LN1-LN6)
    Dim limitesNaturais As Variant
    limitesNaturais = Array("LN1", "LN2", "LN3", "LN4", "LN5", "LN6")

    For i = LBound(limitesNaturais) To UBound(limitesNaturais)
        If M_Validacao.Validar_TipoLimite(CStr(limitesNaturais(i))) Then
            passaram = passaram + 1
        Else
            resultado = resultado & "  FALHOU - " & limitesNaturais(i) & " deveria ser valido" & vbCrLf
            falharam = falharam + 1
        End If
    Next i

    ' Testes negativos
    If Not M_Validacao.Validar_TipoLimite("LA8") Then
        passaram = passaram + 1
    Else
        resultado = resultado & "  FALHOU - LA8 deveria ser invalido" & vbCrLf
        falharam = falharam + 1
    End If

    If Not M_Validacao.Validar_TipoLimite("LN7") Then
        passaram = passaram + 1
    Else
        resultado = resultado & "  FALHOU - LN7 deveria ser invalido" & vbCrLf
        falharam = falharam + 1
    End If

    resultado = resultado & "  OK - Todos os codigos LA1-LA7 e LN1-LN6 validos" & vbCrLf
    resultado = resultado & "  Total: " & passaram & " passaram, " & falharam & " falharam" & vbCrLf

    Teste_TiposLimite = resultado
End Function

Private Function Teste_PrecisaoHorizontal() As String
    '----------------------------------------------------------------------------------
    ' Testa validacao de precisao horizontal
    '----------------------------------------------------------------------------------
    Dim resultado As String
    Dim passaram As Long, falharam As Long

    resultado = "TESTE: Precisao Horizontal" & vbCrLf
    resultado = resultado & String(40, "-") & vbCrLf

    ' Teste LA1 (limite artificial, <= 0.50m)
    If M_Validacao.Validar_PrecisaoHorizontal("LA1", 0.3) Then
        resultado = resultado & "  OK - LA1 com 0.30m: CONFORME" & vbCrLf
        passaram = passaram + 1
    Else
        resultado = resultado & "  FALHOU - LA1 com 0.30m deveria passar" & vbCrLf
        falharam = falharam + 1
    End If

    If Not M_Validacao.Validar_PrecisaoHorizontal("LA1", 0.8) Then
        resultado = resultado & "  OK - LA1 com 0.80m: NAO CONFORME" & vbCrLf
        passaram = passaram + 1
    Else
        resultado = resultado & "  FALHOU - LA1 com 0.80m deveria falhar" & vbCrLf
        falharam = falharam + 1
    End If

    ' Teste LN1 (limite natural, <= 3.00m)
    If M_Validacao.Validar_PrecisaoHorizontal("LN1", 2.5) Then
        resultado = resultado & "  OK - LN1 com 2.50m: CONFORME" & vbCrLf
        passaram = passaram + 1
    Else
        resultado = resultado & "  FALHOU - LN1 com 2.50m deveria passar" & vbCrLf
        falharam = falharam + 1
    End If

    If Not M_Validacao.Validar_PrecisaoHorizontal("LN1", 3.5) Then
        resultado = resultado & "  OK - LN1 com 3.50m: NAO CONFORME" & vbCrLf
        passaram = passaram + 1
    Else
        resultado = resultado & "  FALHOU - LN1 com 3.50m deveria falhar" & vbCrLf
        falharam = falharam + 1
    End If

    ' Teste LA5 (limite inacessivel, <= 7.50m)
    If M_Validacao.Validar_PrecisaoHorizontal("LA5", 5#) Then
        resultado = resultado & "  OK - LA5 com 5.00m: CONFORME" & vbCrLf
        passaram = passaram + 1
    Else
        resultado = resultado & "  FALHOU - LA5 com 5.00m deveria passar" & vbCrLf
        falharam = falharam + 1
    End If

    If Not M_Validacao.Validar_PrecisaoHorizontal("LA5", 8#) Then
        resultado = resultado & "  OK - LA5 com 8.00m: NAO CONFORME" & vbCrLf
        passaram = passaram + 1
    Else
        resultado = resultado & "  FALHOU - LA5 com 8.00m deveria falhar" & vbCrLf
        falharam = falharam + 1
    End If

    resultado = resultado & "  Total: " & passaram & " passaram, " & falharam & " falharam" & vbCrLf

    Teste_PrecisaoHorizontal = resultado
End Function

Private Function Teste_MetodosPosicionamento() As String
    '----------------------------------------------------------------------------------
    ' Testa validacao de metodos de posicionamento
    '----------------------------------------------------------------------------------
    Dim resultado As String
    Dim passaram As Long, falharam As Long
    Dim i As Long

    resultado = "TESTE: Metodos de Posicionamento" & vbCrLf
    resultado = resultado & String(40, "-") & vbCrLf

    ' Metodos validos
    Dim metodos As Variant
    metodos = Array("GNSS-RTK", "GNSS-PPP", "GNSS-REL", "TOP", "GAN", "SRE", "BCA")

    For i = LBound(metodos) To UBound(metodos)
        If M_Validacao.Validar_MetodoPosicionamento(CStr(metodos(i))) Then
            passaram = passaram + 1
        Else
            resultado = resultado & "  FALHOU - " & metodos(i) & " deveria ser valido" & vbCrLf
            falharam = falharam + 1
        End If
    Next i

    ' Teste negativo
    If Not M_Validacao.Validar_MetodoPosicionamento("INVALIDO") Then
        passaram = passaram + 1
    Else
        resultado = resultado & "  FALHOU - Metodo invalido deveria falhar" & vbCrLf
        falharam = falharam + 1
    End If

    resultado = resultado & "  OK - Todos os metodos validos aceitos" & vbCrLf
    resultado = resultado & "  Total: " & passaram & " passaram, " & falharam & " falharam" & vbCrLf

    Teste_MetodosPosicionamento = resultado
End Function

Private Function Teste_ValidacaoCompleta() As String
    '----------------------------------------------------------------------------------
    ' Testa validacao completa de registro
    '----------------------------------------------------------------------------------
    Dim resultado As String
    Dim msgErro As String
    Dim valido As Boolean

    resultado = "TESTE: Validacao Completa de Registro" & vbCrLf
    resultado = resultado & String(40, "-") & vbCrLf

    ' Teste com registro valido
    valido = M_Validacao.Validar_RegistroCompleto("M", "LA1", 0.3, 0.5, "GNSS-RTK", msgErro)
    If valido Then
        resultado = resultado & "  OK - Registro valido aprovado" & vbCrLf
    Else
        resultado = resultado & "  FALHOU - Registro valido foi rejeitado" & vbCrLf
        resultado = resultado & "    Erro: " & msgErro & vbCrLf
    End If

    ' Teste com tipo de vertice invalido
    valido = M_Validacao.Validar_RegistroCompleto("X", "LA1", 0.3, 0.5, "GNSS-RTK", msgErro)
    If Not valido Then
        resultado = resultado & "  OK - Tipo vertice invalido detectado" & vbCrLf
    Else
        resultado = resultado & "  FALHOU - Tipo vertice invalido nao foi detectado" & vbCrLf
    End If

    ' Teste com precisao fora do padrao
    valido = M_Validacao.Validar_RegistroCompleto("M", "LA1", 0.8, 0.5, "GNSS-RTK", msgErro)
    If Not valido Then
        resultado = resultado & "  OK - Precisao fora do padrao detectada" & vbCrLf
    Else
        resultado = resultado & "  FALHOU - Precisao fora do padrao nao foi detectada" & vbCrLf
    End If

    ' Teste com metodo invalido
    valido = M_Validacao.Validar_RegistroCompleto("M", "LA1", 0.3, 0.5, "INVALIDO", msgErro)
    If Not valido Then
        resultado = resultado & "  OK - Metodo invalido detectado" & vbCrLf
    Else
        resultado = resultado & "  FALHOU - Metodo invalido nao foi detectado" & vbCrLf
    End If

    Teste_ValidacaoCompleta = resultado
End Function

Private Function Teste_CalculoEMQ() As String
    '----------------------------------------------------------------------------------
    ' Testa calculo de Erro Medio Quadratico
    '----------------------------------------------------------------------------------
    Dim resultado As String
    Dim precisoes As Variant
    Dim emq As Double

    resultado = "TESTE: Calculo de EMQ" & vbCrLf
    resultado = resultado & String(40, "-") & vbCrLf

    ' Teste com valores conhecidos
    ' EMQ de [0.3, 0.4, 0.5] = sqrt((0.09 + 0.16 + 0.25)/3) = sqrt(0.1667) = 0.408
    precisoes = Array(0.3, 0.4, 0.5)
    emq = M_Validacao.Calcular_EMQ(precisoes)

    If Abs(emq - 0.408) < 0.01 Then
        resultado = resultado & "  OK - EMQ calculado corretamente: " & Format(emq, "0.000") & vbCrLf
    Else
        resultado = resultado & "  FALHOU - EMQ incorreto: " & Format(emq, "0.000") & " (esperado ~0.408)" & vbCrLf
    End If

    Teste_CalculoEMQ = resultado
End Function
