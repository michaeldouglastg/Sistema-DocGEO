Attribute VB_Name = "M_Validacao"
Option Explicit
' ==============================================================================
' MODULO: M_VALIDACAO
' DESCRICAO: VALIDACOES DE CONFORMIDADE COM MANUAL TECNICO INCRA
' REFERENCIA: Portaria N 2.502/2022 - Manual Tecnico 2 Edicao
' ==============================================================================

' ==============================================================================
' CONSTANTES DE PRECISAO CONFORME MANUAL INCRA (Cap. 1.4.4)
' ==============================================================================
Private Const PREC_LIMITE_ARTIFICIAL As Double = 0.5        ' LA1-LA4: <= 0.50m
Private Const PREC_LIMITE_NATURAL As Double = 3#             ' LN1-LN6: <= 3.00m
Private Const PREC_LIMITE_INACESSIVEL As Double = 7.5       ' LA5-LA7: <= 7.50m
Private Const PREC_VERTICAL_PADRAO As Double = 1#           ' Padrao vertical: <= 1.00m

' ==============================================================================
' VALIDACAO DE TIPO DE VERTICE (Cap. 1.5 do Manual)
' ==============================================================================
Public Function Validar_TipoVertice(tipo As String) As Boolean
    '----------------------------------------------------------------------------------
    ' Valida se o tipo de vertice esta conforme Manual INCRA
    ' Tipos validos:
    '   M - Marco (materializado no terreno)
    '   P - Ponto (feicao natural ou artificial identificavel)
    '   V - Virtual (calculado, sem materializacao fisica)
    '----------------------------------------------------------------------------------
    Dim tipoUpper As String
    tipoUpper = UCase(Trim(tipo))

    Select Case tipoUpper
        Case "M", "P", "V"
            Validar_TipoVertice = True
        Case Else
            Validar_TipoVertice = False
    End Select
End Function

Public Function Obter_DescricaoTipoVertice(tipo As String) As String
    '----------------------------------------------------------------------------------
    ' Retorna descricao do tipo de vertice conforme manual
    '----------------------------------------------------------------------------------
    Select Case UCase(Trim(tipo))
        Case "M"
            Obter_DescricaoTipoVertice = "Marco (materializado)"
        Case "P"
            Obter_DescricaoTipoVertice = "Ponto (feicao identificavel)"
        Case "V"
            Obter_DescricaoTipoVertice = "Virtual (calculado)"
        Case Else
            Obter_DescricaoTipoVertice = "TIPO INVALIDO"
    End Select
End Function

' ==============================================================================
' VALIDACAO DE TIPO DE LIMITE/DIVISA (Cap. 2 do Manual)
' ==============================================================================
Public Function Validar_TipoLimite(codigoLimite As String) As Boolean
    '----------------------------------------------------------------------------------
    ' Valida se o codigo de limite esta conforme Manual INCRA
    ' Limites Artificiais: LA1-LA7
    ' Limites Naturais: LN1-LN6
    '----------------------------------------------------------------------------------
    Dim cod As String
    cod = UCase(Trim(codigoLimite))

    Select Case cod
        ' Limites Artificiais
        Case "LA1", "LA2", "LA3", "LA4", "LA5", "LA6", "LA7"
            Validar_TipoLimite = True
        ' Limites Naturais
        Case "LN1", "LN2", "LN3", "LN4", "LN5", "LN6"
            Validar_TipoLimite = True
        Case Else
            Validar_TipoLimite = False
    End Select
End Function

Public Function Obter_DescricaoLimite(codigoLimite As String) As String
    '----------------------------------------------------------------------------------
    ' Retorna descricao oficial do tipo de limite conforme Manual INCRA
    '----------------------------------------------------------------------------------
    Select Case UCase(Trim(codigoLimite))
        ' Limites Artificiais
        Case "LA1"
            Obter_DescricaoLimite = "Cerca"
        Case "LA2"
            Obter_DescricaoLimite = "Estrada"
        Case "LA3"
            Obter_DescricaoLimite = "Rio/Corrego Canalizado"
        Case "LA4"
            Obter_DescricaoLimite = "Vala, Rego, Canal"
        Case "LA5"
            Obter_DescricaoLimite = "Limite Inacessivel (Artificial)"
        Case "LA6"
            Obter_DescricaoLimite = "Limite Inacessivel (Serra, Escarpa)"
        Case "LA7"
            Obter_DescricaoLimite = "Limite Inacessivel (Rio, Corrego, Lago)"
        ' Limites Naturais
        Case "LN1"
            Obter_DescricaoLimite = "Talvegue de Rio/Corrego"
        Case "LN2"
            Obter_DescricaoLimite = "Crista de Serra/Espigao"
        Case "LN3"
            Obter_DescricaoLimite = "Margem de Rio/Corrego"
        Case "LN4"
            Obter_DescricaoLimite = "Margem de Lago/Lagoa"
        Case "LN5"
            Obter_DescricaoLimite = "Margem de Oceano"
        Case "LN6"
            Obter_DescricaoLimite = "Limite Seco de Praia/Mangue"
        Case Else
            Obter_DescricaoLimite = "CODIGO INVALIDO"
    End Select
End Function

Public Function Obter_PrecisaoRequerida(codigoLimite As String) As Double
    '----------------------------------------------------------------------------------
    ' Retorna precisao horizontal requerida em metros conforme Manual INCRA (Cap. 1.4.4)
    '----------------------------------------------------------------------------------
    Select Case UCase(Trim(codigoLimite))
        ' Limites Artificiais Acessiveis: <= 0.50m
        Case "LA1", "LA2", "LA3", "LA4"
            Obter_PrecisaoRequerida = PREC_LIMITE_ARTIFICIAL
        ' Limites Inacessiveis: <= 7.50m
        Case "LA5", "LA6", "LA7"
            Obter_PrecisaoRequerida = PREC_LIMITE_INACESSIVEL
        ' Limites Naturais: <= 3.00m
        Case "LN1", "LN2", "LN3", "LN4", "LN5", "LN6"
            Obter_PrecisaoRequerida = PREC_LIMITE_NATURAL
        Case Else
            Obter_PrecisaoRequerida = -1  ' Codigo invalido
    End Select
End Function

' ==============================================================================
' VALIDACAO DE PRECISAO HORIZONTAL (Cap. 1.4.4 do Manual)
' ==============================================================================
Public Function Validar_PrecisaoHorizontal(codigoLimite As String, precisaoMedida As Double) As Boolean
    '----------------------------------------------------------------------------------
    ' Valida se a precisao horizontal medida atende o requisito do Manual INCRA
    ' Retorna True se conforme, False se fora do padrao
    '----------------------------------------------------------------------------------
    Dim precisaoRequerida As Double

    If precisaoMedida < 0 Then
        Validar_PrecisaoHorizontal = False
        Exit Function
    End If

    precisaoRequerida = Obter_PrecisaoRequerida(codigoLimite)

    If precisaoRequerida < 0 Then
        ' Codigo de limite invalido
        Validar_PrecisaoHorizontal = False
    Else
        ' Verifica se precisao medida <= requerida
        Validar_PrecisaoHorizontal = (precisaoMedida <= precisaoRequerida)
    End If
End Function

Public Function Validar_PrecisaoVertical(precisaoMedida As Double) As Double
    '----------------------------------------------------------------------------------
    ' Valida se a precisao vertical atende padrao (geralmente <= 1.00m)
    '----------------------------------------------------------------------------------
    If precisaoMedida < 0 Then
        Validar_PrecisaoVertical = False
    Else
        Validar_PrecisaoVertical = (precisaoMedida <= PREC_VERTICAL_PADRAO)
    End If
End Function

' ==============================================================================
' VALIDACAO DE METODO DE POSICIONAMENTO (Cap. 1.4.3 e 3 do Manual)
' ==============================================================================
Public Function Validar_MetodoPosicionamento(metodo As String) As Boolean
    '----------------------------------------------------------------------------------
    ' Valida se o metodo de posicionamento e um dos aceitos pelo Manual INCRA
    ' Metodos validos (Cap. 3):
    '   - GNSS-RTK (Real Time Kinematic)
    '   - GNSS-PPP (Precise Point Positioning)
    '   - GNSS-REL (GNSS Relativo)
    '   - TOP (Topografia Classica)
    '   - GAN (Geometria Analitica)
    '   - SRE (Sensoriamento Remoto)
    '   - BCA (Base Cartografica)
    '----------------------------------------------------------------------------------
    Dim metodoUpper As String
    metodoUpper = UCase(Trim(metodo))

    Select Case metodoUpper
        Case "GNSS-RTK", "GNSS-PPP", "GNSS-REL", "TOP", "GAN", "SRE", "BCA"
            Validar_MetodoPosicionamento = True
        Case Else
            Validar_MetodoPosicionamento = False
    End Select
End Function

Public Function Obter_DescricaoMetodo(metodo As String) As String
    '----------------------------------------------------------------------------------
    ' Retorna descricao completa do metodo de posicionamento
    '----------------------------------------------------------------------------------
    Select Case UCase(Trim(metodo))
        Case "GNSS-RTK"
            Obter_DescricaoMetodo = "GNSS - Real Time Kinematic"
        Case "GNSS-PPP"
            Obter_DescricaoMetodo = "GNSS - Precise Point Positioning"
        Case "GNSS-REL"
            Obter_DescricaoMetodo = "GNSS - Relativo"
        Case "TOP"
            Obter_DescricaoMetodo = "Topografia Classica"
        Case "GAN"
            Obter_DescricaoMetodo = "Geometria Analitica"
        Case "SRE"
            Obter_DescricaoMetodo = "Sensoriamento Remoto"
        Case "BCA"
            Obter_DescricaoMetodo = "Base Cartografica"
        Case Else
            Obter_DescricaoMetodo = "METODO NAO ESPECIFICADO"
    End Select
End Function

' ==============================================================================
' FUNCOES DE VALIDACAO COMPLETA DE REGISTRO
' ==============================================================================
Public Function Validar_RegistroCompleto(tipo As String, codigoLimite As String, _
                                          precisaoH As Double, precisaoV As Double, _
                                          metodo As String, ByRef mensagemErro As String) As Boolean
    '----------------------------------------------------------------------------------
    ' Valida todos os campos de um registro conforme Manual INCRA
    ' Retorna True se conforme, False se houver erros
    ' mensagemErro contera detalhes do problema
    '----------------------------------------------------------------------------------
    Dim erros As String
    erros = ""

    ' Valida tipo de vertice
    If Not Validar_TipoVertice(tipo) Then
        erros = erros & "- Tipo de vertice invalido (deve ser M, P ou V)" & vbCrLf
    End If

    ' Valida tipo de limite
    If Not Validar_TipoLimite(codigoLimite) Then
        erros = erros & "- Codigo de limite invalido (deve ser LA1-LA7 ou LN1-LN6)" & vbCrLf
    End If

    ' Valida precisao horizontal
    If Not Validar_PrecisaoHorizontal(codigoLimite, precisaoH) Then
        Dim precReq As Double
        precReq = Obter_PrecisaoRequerida(codigoLimite)
        erros = erros & "- Precisao horizontal fora do padrao (medida: " & Format(precisaoH, "0.00") & "m, " & _
                        "requerida: <= " & Format(precReq, "0.00") & "m)" & vbCrLf
    End If

    ' Valida precisao vertical
    If Not Validar_PrecisaoVertical(precisaoV) Then
        erros = erros & "- Precisao vertical fora do padrao (medida: " & Format(precisaoV, "0.00") & "m, " & _
                        "requerida: <= " & Format(PREC_VERTICAL_PADRAO, "0.00") & "m)" & vbCrLf
    End If

    ' Valida metodo de posicionamento
    If Not Validar_MetodoPosicionamento(metodo) Then
        erros = erros & "- Metodo de posicionamento invalido" & vbCrLf
    End If

    If erros = "" Then
        Validar_RegistroCompleto = True
        mensagemErro = ""
    Else
        Validar_RegistroCompleto = False
        mensagemErro = "ERROS DE VALIDACAO INCRA:" & vbCrLf & erros
    End If
End Function

' ==============================================================================
' FUNCOES DE RELATORIO DE QUALIDADE
' ==============================================================================
Public Function Calcular_EMQ(arrPrecisoes As Variant) As Double
    '----------------------------------------------------------------------------------
    ' Calcula Erro Medio Quadratico (RMS) de um conjunto de precisoes
    ' Usado para relatorio de qualidade posicional
    '----------------------------------------------------------------------------------
    Dim soma As Double, i As Long, n As Long

    On Error GoTo ErroCalculo

    If Not IsArray(arrPrecisoes) Then
        Calcular_EMQ = -1
        Exit Function
    End If

    n = UBound(arrPrecisoes) - LBound(arrPrecisoes) + 1
    If n = 0 Then
        Calcular_EMQ = -1
        Exit Function
    End If

    soma = 0
    For i = LBound(arrPrecisoes) To UBound(arrPrecisoes)
        soma = soma + (CDbl(arrPrecisoes(i)) ^ 2)
    Next i

    Calcular_EMQ = Sqr(soma / n)
    Exit Function

ErroCalculo:
    Calcular_EMQ = -1
End Function

Public Function Gerar_RelatorioQualidade(nomePlanilha As String, nomeTabela As String) As String
    '----------------------------------------------------------------------------------
    ' Gera relatorio de qualidade posicional para a tabela especificada
    ' Retorna texto formatado com estatisticas
    '----------------------------------------------------------------------------------
    Dim ws As Worksheet, tbl As ListObject
    Dim i As Long, qtd As Long
    Dim arrPrecH() As Double, arrPrecV() As Double
    Dim emqH As Double, emqV As Double
    Dim maxH As Double, maxV As Double
    Dim minH As Double, minV As Double
    Dim conforme As Long, naoConforme As Long
    Dim relatorio As String

    On Error GoTo ErroRelatorio

    Set ws = ThisWorkbook.Sheets(nomePlanilha)
    Set tbl = ws.ListObjects(nomeTabela)
    qtd = tbl.ListRows.Count

    If qtd = 0 Then
        Gerar_RelatorioQualidade = "Nenhum dado disponivel para analise."
        Exit Function
    End If

    ReDim arrPrecH(1 To qtd)
    ReDim arrPrecV(1 To qtd)
    maxH = 0: maxV = 0
    minH = 999: minV = 999
    conforme = 0: naoConforme = 0

    ' Coleta dados (assumindo colunas de precisao existem)
    ' NOTA: Ajustar indices de coluna conforme estrutura real da tabela
    For i = 1 To qtd
        ' Aqui seria necessario ler as colunas de precisao
        ' Por enquanto, exemplo generico
        ' arrPrecH(i) = tbl.DataBodyRange(i, colPrecH).Value
        ' arrPrecV(i) = tbl.DataBodyRange(i, colPrecV).Value
    Next i

    ' Calcula estatisticas
    emqH = Calcular_EMQ(arrPrecH)
    emqV = Calcular_EMQ(arrPrecV)

    ' Formata relatorio
    relatorio = "RELATORIO DE QUALIDADE POSICIONAL" & vbCrLf
    relatorio = relatorio & String(50, "=") & vbCrLf & vbCrLf
    relatorio = relatorio & "Total de vertices: " & qtd & vbCrLf
    relatorio = relatorio & "Vertices conformes: " & conforme & vbCrLf
    relatorio = relatorio & "Vertices nao conformes: " & naoConforme & vbCrLf & vbCrLf
    relatorio = relatorio & "PRECISAO HORIZONTAL:" & vbCrLf
    relatorio = relatorio & "  EMQ: " & Format(emqH, "0.000") & " m" & vbCrLf
    relatorio = relatorio & "  Minima: " & Format(minH, "0.000") & " m" & vbCrLf
    relatorio = relatorio & "  Maxima: " & Format(maxH, "0.000") & " m" & vbCrLf & vbCrLf
    relatorio = relatorio & "PRECISAO VERTICAL:" & vbCrLf
    relatorio = relatorio & "  EMQ: " & Format(emqV, "0.000") & " m" & vbCrLf
    relatorio = relatorio & "  Minima: " & Format(minV, "0.000") & " m" & vbCrLf
    relatorio = relatorio & "  Maxima: " & Format(maxV, "0.000") & " m" & vbCrLf

    Gerar_RelatorioQualidade = relatorio
    Exit Function

ErroRelatorio:
    Gerar_RelatorioQualidade = "Erro ao gerar relatorio: " & Err.Description
End Function

' ==============================================================================
' FUNCOES AUXILIARES
' ==============================================================================
Public Function Obter_ListaMetodos() As Variant
    '----------------------------------------------------------------------------------
    ' Retorna array com todos os metodos de posicionamento validos
    ' Util para preencher ComboBox na interface
    '----------------------------------------------------------------------------------
    Obter_ListaMetodos = Array("GNSS-RTK", "GNSS-PPP", "GNSS-REL", "TOP", "GAN", "SRE", "BCA")
End Function

Public Function Obter_ListaTiposVertice() As Variant
    '----------------------------------------------------------------------------------
    ' Retorna array com tipos de vertice validos
    '----------------------------------------------------------------------------------
    Obter_ListaTiposVertice = Array("M", "P", "V")
End Function

Public Function Obter_ListaLimitesArtificiais() As Variant
    '----------------------------------------------------------------------------------
    ' Retorna array com codigos de limites artificiais
    '----------------------------------------------------------------------------------
    Obter_ListaLimitesArtificiais = Array("LA1", "LA2", "LA3", "LA4", "LA5", "LA6", "LA7")
End Function

Public Function Obter_ListaLimitesNaturais() As Variant
    '----------------------------------------------------------------------------------
    ' Retorna array com codigos de limites naturais
    '----------------------------------------------------------------------------------
    Obter_ListaLimitesNaturais = Array("LN1", "LN2", "LN3", "LN4", "LN5", "LN6")
End Function
