# GUIA DE VALIDAÇÕES INCRA
## Sistema DocGEO - Implementação de Validações Conforme Manual Técnico

**Versão:** 1.0
**Data:** 27 de dezembro de 2024
**Referência:** Portaria Nº 2.502/2022 - Manual Técnico INCRA 2ª Edição

---

## 📋 ÍNDICE

1. [Visão Geral](#visão-geral)
2. [Novos Módulos](#novos-módulos)
3. [Estrutura de Dados](#estrutura-de-dados)
4. [Como Usar](#como-usar)
5. [Procedimentos de Setup](#procedimentos-de-setup)
6. [Validações Disponíveis](#validações-disponíveis)
7. [Integração com Sistema Existente](#integração-com-sistema-existente)
8. [Testes](#testes)
9. [Referências do Manual INCRA](#referências-do-manual-incra)

---

## 1. VISÃO GERAL

Este pacote adiciona **validações completas de conformidade** com o Manual Técnico do INCRA ao Sistema DocGEO, incluindo:

### ✅ Validações Implementadas

- **Tipos de Vértices** (M, P, V) - conforme Cap. 1.5 do Manual
- **Tipos de Limites/Divisas** (LA1-LA7, LN1-LN6) - conforme Cap. 2
- **Precisão Horizontal** por tipo de limite - conforme Cap. 1.4.4
- **Precisão Vertical** - conforme padrões técnicos
- **Métodos de Posicionamento** - conforme Cap. 1.4.3 e 3
- **Cálculo de EMQ** (Erro Médio Quadrático) - para relatórios de qualidade

### 📦 Arquivos Adicionados

| Arquivo | Descrição |
|---------|-----------|
| `M_Validacao.bas` | Módulo principal com funções de validação |
| `M_Setup_Parametros.bas` | Setup e manutenção de parâmetros INCRA |
| `Teste_Validacoes_INCRA.bas` | Suite de testes unitários |
| `dados_parametros_incra.csv` | Dados oficiais dos parâmetros |
| `GUIA_VALIDACOES_INCRA.md` | Esta documentação |

---

## 2. NOVOS MÓDULOS

### 2.1 M_Validacao.bas

Módulo principal de validações. Funções principais:

#### Validação de Tipos de Vértices
```vba
' Valida se tipo é M, P ou V
Function Validar_TipoVertice(tipo As String) As Boolean

' Retorna descrição do tipo
Function Obter_DescricaoTipoVertice(tipo As String) As String
```

#### Validação de Tipos de Limites
```vba
' Valida se código é LA1-LA7 ou LN1-LN6
Function Validar_TipoLimite(codigoLimite As String) As Boolean

' Retorna descrição oficial do limite
Function Obter_DescricaoLimite(codigoLimite As String) As String

' Retorna precisão requerida em metros
Function Obter_PrecisaoRequerida(codigoLimite As String) As Double
```

#### Validação de Precisão
```vba
' Valida se precisão horizontal atende requisito
Function Validar_PrecisaoHorizontal(codigoLimite As String, precisaoMedida As Double) As Boolean

' Valida precisão vertical (padrão <= 1.00m)
Function Validar_PrecisaoVertical(precisaoMedida As Double) As Boolean
```

#### Validação de Métodos de Posicionamento
```vba
' Valida se método é aceito pelo INCRA
Function Validar_MetodoPosicionamento(metodo As String) As Boolean

' Retorna descrição completa do método
Function Obter_DescricaoMetodo(metodo As String) As String
```

#### Validação Completa
```vba
' Valida todos os campos de um registro
Function Validar_RegistroCompleto(tipo As String, codigoLimite As String, _
                                   precisaoH As Double, precisaoV As Double, _
                                   metodo As String, ByRef mensagemErro As String) As Boolean
```

#### Relatórios de Qualidade
```vba
' Calcula Erro Médio Quadrático (RMS)
Function Calcular_EMQ(arrPrecisoes As Variant) As Double

' Gera relatório de qualidade posicional
Function Gerar_RelatorioQualidade(nomePlanilha As String, nomeTabela As String) As String
```

### 2.2 M_Setup_Parametros.bas

Módulo para setup inicial e manutenção de parâmetros.

#### Funções Principais
```vba
' Popula tabela com códigos oficiais INCRA
Sub Setup_PopularParametrosINCRA()

' Verifica estrutura das tabelas
Sub Setup_VerificarEstruturaDados()

' Adiciona colunas de validação nas tabelas
Sub Setup_AdicionarColunasValidacao()
```

### 2.3 Teste_Validacoes_INCRA.bas

Suite completa de testes unitários.

```vba
' Executa todos os testes
Sub ExecutarTodosTestes()
```

---

## 3. ESTRUTURA DE DADOS

### 3.1 Novos Campos nas Tabelas SGL e UTM

As tabelas principais precisam incluir 4 novos campos:

| Campo | Tipo | Descrição | Exemplo |
|-------|------|-----------|---------|
| **Precisao H (m)** | Número (0.00) | Precisão horizontal em metros | 0.30 |
| **Precisao V (m)** | Número (0.00) | Precisão vertical em metros | 0.50 |
| **Metodo Posic.** | Texto | Código do método de posicionamento | GNSS-RTK |
| **Cod. Limite** | Texto | Código do tipo de limite | LA1 |

### 3.2 Tabela de Parâmetros INCRA

A tabela `tbl_Parametros` deve conter os códigos oficiais:

| Codigo | Descricao | Tipo | Precisao_Requerida |
|--------|-----------|------|--------------------|
| LA1 | Cerca | Artificial | 0.50m |
| LA2 | Estrada | Artificial | 0.50m |
| LA3 | Rio/Córrego Canalizado | Artificial | 0.50m |
| LA4 | Vala, Rego, Canal | Artificial | 0.50m |
| LA5 | Limite Inacessível (Artificial) | Artificial | 7.50m |
| LA6 | Limite Inacessível (Serra, Escarpa) | Artificial | 7.50m |
| LA7 | Limite Inacessível (Rio, Córrego, Lago) | Artificial | 7.50m |
| LN1 | Talvegue de Rio/Córrego | Natural | 3.00m |
| LN2 | Crista de Serra/Espigão | Natural | 3.00m |
| LN3 | Margem de Rio/Córrego | Natural | 3.00m |
| LN4 | Margem de Lago/Lagoa | Natural | 3.00m |
| LN5 | Margem de Oceano | Natural | 3.00m |
| LN6 | Limite Seco de Praia/Mangue | Natural | 3.00m |
| M | Marco (materializado) | Vertice | - |
| P | Ponto (feição identificável) | Vertice | - |
| V | Virtual (calculado) | Vertice | - |
| GNSS-RTK | GNSS - Real Time Kinematic | Metodo | - |
| GNSS-PPP | GNSS - Precise Point Positioning | Metodo | - |
| GNSS-REL | GNSS - Relativo | Metodo | - |
| TOP | Topografia Clássica | Metodo | - |
| GAN | Geometria Analítica | Metodo | - |
| SRE | Sensoriamento Remoto | Metodo | - |
| BCA | Base Cartográfica | Metodo | - |

---

## 4. COMO USAR

### 4.1 Primeira Execução (Setup Inicial)

Execute os procedimentos de setup **NA ORDEM**:

#### Passo 1: Popular Parâmetros INCRA
```vba
Sub ExecutarSetupInicial()
    ' Popula tabela de parâmetros
    Call M_Setup_Parametros.Setup_PopularParametrosINCRA()
End Sub
```

#### Passo 2: Verificar Estrutura de Dados
```vba
Sub VerificarEstrutura()
    ' Verifica se as colunas necessárias existem
    Call M_Setup_Parametros.Setup_VerificarEstruturaDados()
End Sub
```

#### Passo 3: Adicionar Colunas (se necessário)
```vba
Sub AdicionarColunas()
    ' Adiciona as 4 novas colunas nas tabelas SGL e UTM
    Call M_Setup_Parametros.Setup_AdicionarColunasValidacao()
End Sub
```

### 4.2 Uso nas Validações de Entrada

#### Exemplo 1: Validar Tipo de Vértice
```vba
Dim tipo As String
tipo = txtTipoVertice.Value

If Not M_Validacao.Validar_TipoVertice(tipo) Then
    MsgBox "Tipo de vértice inválido. Use M, P ou V.", vbExclamation
    Exit Sub
End If
```

#### Exemplo 2: Validar Código de Limite
```vba
Dim codLimite As String
codLimite = cboCodigoLimite.Value

If Not M_Validacao.Validar_TipoLimite(codLimite) Then
    MsgBox "Código de limite inválido. Use LA1-LA7 ou LN1-LN6.", vbExclamation
    Exit Sub
End If

' Mostra precisão requerida
Dim precReq As Double
precReq = M_Validacao.Obter_PrecisaoRequerida(codLimite)
lblPrecisaoRequerida.Caption = "Precisão requerida: <= " & Format(precReq, "0.00") & "m"
```

#### Exemplo 3: Validar Precisão Horizontal
```vba
Dim codLimite As String, precisaoH As Double
codLimite = cboCodigoLimite.Value
precisaoH = CDbl(txtPrecisaoH.Value)

If Not M_Validacao.Validar_PrecisaoHorizontal(codLimite, precisaoH) Then
    Dim precReq As Double
    precReq = M_Validacao.Obter_PrecisaoRequerida(codLimite)
    MsgBox "Precisão horizontal fora do padrão!" & vbCrLf & _
           "Medida: " & Format(precisaoH, "0.00") & "m" & vbCrLf & _
           "Requerida: <= " & Format(precReq, "0.00") & "m", vbExclamation
    Exit Sub
End If
```

#### Exemplo 4: Validação Completa de Registro
```vba
Dim msgErro As String

If Not M_Validacao.Validar_RegistroCompleto( _
        tipo:=txtTipo.Value, _
        codigoLimite:=cboCodigoLimite.Value, _
        precisaoH:=CDbl(txtPrecisaoH.Value), _
        precisaoV:=CDbl(txtPrecisaoV.Value), _
        metodo:=cboMetodo.Value, _
        mensagemErro:=msgErro) Then

    MsgBox msgErro, vbExclamation, "Validação INCRA"
    Exit Sub
End If

' Se chegou aqui, dados estão conformes
MsgBox "Dados validados com sucesso!", vbInformation
```

### 4.3 Populando ComboBoxes com Valores Válidos

#### ComboBox de Tipos de Vértices
```vba
Private Sub UserForm_Initialize()
    Dim tiposVertice As Variant
    tiposVertice = M_Validacao.Obter_ListaTiposVertice()  ' Retorna Array("M", "P", "V")

    cboTipoVertice.Clear
    Dim i As Long
    For i = LBound(tiposVertice) To UBound(tiposVertice)
        cboTipoVertice.AddItem tiposVertice(i)
    Next i
End Sub
```

#### ComboBox de Códigos de Limites
```vba
Private Sub UserForm_Initialize()
    Dim limitesArt As Variant, limitesNat As Variant
    limitesArt = M_Validacao.Obter_ListaLimitesArtificiais()  ' LA1-LA7
    limitesNat = M_Validacao.Obter_ListaLimitesNaturais()     ' LN1-LN6

    cboCodigoLimite.Clear

    ' Adiciona limites artificiais
    Dim i As Long
    For i = LBound(limitesArt) To UBound(limitesArt)
        cboCodigoLimite.AddItem limitesArt(i) & " - " & _
                                M_Validacao.Obter_DescricaoLimite(CStr(limitesArt(i)))
    Next i

    ' Adiciona limites naturais
    For i = LBound(limitesNat) To UBound(limitesNat)
        cboCodigoLimite.AddItem limitesNat(i) & " - " & _
                                M_Validacao.Obter_DescricaoLimite(CStr(limitesNat(i)))
    Next i
End Sub
```

#### ComboBox de Métodos de Posicionamento
```vba
Private Sub UserForm_Initialize()
    Dim metodos As Variant
    metodos = M_Validacao.Obter_ListaMetodos()

    cboMetodo.Clear
    Dim i As Long
    For i = LBound(metodos) To UBound(metodos)
        cboMetodo.AddItem metodos(i) & " - " & _
                          M_Validacao.Obter_DescricaoMetodo(CStr(metodos(i)))
    Next i
End Sub
```

---

## 5. PROCEDIMENTOS DE SETUP

### 5.1 Checklist de Implementação

- [ ] Importar módulos VBA (`M_Validacao.bas`, `M_Setup_Parametros.bas`, `Teste_Validacoes_INCRA.bas`)
- [ ] Atualizar `M_Config.bas` com novas constantes
- [ ] Executar `Setup_PopularParametrosINCRA()`
- [ ] Executar `Setup_VerificarEstruturaDados()`
- [ ] Executar `Setup_AdicionarColunasValidacao()` (se necessário)
- [ ] Executar `ExecutarTodosTestes()` para validar implementação
- [ ] Atualizar formulários de entrada de dados
- [ ] Atualizar processos de importação
- [ ] Atualizar geração de documentos

### 5.2 Verificação de Instalação

Execute este código para verificar se tudo foi instalado corretamente:

```vba
Sub VerificarInstalacao()
    Dim resultado As String

    resultado = "VERIFICACAO DE INSTALACAO" & vbCrLf
    resultado = resultado & String(50, "=") & vbCrLf & vbCrLf

    ' Testa se módulo está disponível
    On Error Resume Next
    Dim teste As Boolean
    teste = M_Validacao.Validar_TipoVertice("M")

    If Err.Number = 0 Then
        resultado = resultado & "OK - Modulo M_Validacao carregado" & vbCrLf
    Else
        resultado = resultado & "ERRO - Modulo M_Validacao nao encontrado" & vbCrLf
    End If
    On Error GoTo 0

    ' Verifica constantes em M_Config
    On Error Resume Next
    Dim prec As Double
    prec = M_Config.PREC_LIMITE_ARTIFICIAL

    If Err.Number = 0 Then
        resultado = resultado & "OK - Constantes INCRA em M_Config" & vbCrLf
    Else
        resultado = resultado & "ERRO - Constantes INCRA nao encontradas em M_Config" & vbCrLf
    End If
    On Error GoTo 0

    ' Verifica estrutura de dados
    resultado = resultado & vbCrLf
    Call M_Setup_Parametros.Setup_VerificarEstruturaDados()

    MsgBox resultado, vbInformation
End Sub
```

---

## 6. VALIDAÇÕES DISPONÍVEIS

### 6.1 Resumo das Validações

| Validação | Função | Critério | Referência Manual |
|-----------|--------|----------|-------------------|
| Tipo de Vértice | `Validar_TipoVertice()` | M, P ou V | Cap. 1.5 |
| Tipo de Limite | `Validar_TipoLimite()` | LA1-LA7, LN1-LN6 | Cap. 2 |
| Precisão LA1-LA4 | `Validar_PrecisaoHorizontal()` | ≤ 0.50m | Cap. 1.4.4 |
| Precisão LN1-LN6 | `Validar_PrecisaoHorizontal()` | ≤ 3.00m | Cap. 1.4.4 |
| Precisão LA5-LA7 | `Validar_PrecisaoHorizontal()` | ≤ 7.50m | Cap. 1.4.4 |
| Precisão Vertical | `Validar_PrecisaoVertical()` | ≤ 1.00m | Padrão Técnico |
| Método Posicionamento | `Validar_MetodoPosicionamento()` | GNSS-RTK, PPP, REL, TOP, GAN, SRE, BCA | Cap. 1.4.3 e 3 |

### 6.2 Tabela de Códigos INCRA

#### Limites Artificiais (LA)

| Código | Descrição | Precisão |
|--------|-----------|----------|
| LA1 | Cerca | ≤ 0.50m |
| LA2 | Estrada | ≤ 0.50m |
| LA3 | Rio/Córrego Canalizado | ≤ 0.50m |
| LA4 | Vala, Rego, Canal | ≤ 0.50m |
| LA5 | Limite Inacessível (Artificial) | ≤ 7.50m |
| LA6 | Limite Inacessível (Serra, Escarpa) | ≤ 7.50m |
| LA7 | Limite Inacessível (Rio, Córrego, Lago) | ≤ 7.50m |

#### Limites Naturais (LN)

| Código | Descrição | Precisão |
|--------|-----------|----------|
| LN1 | Talvegue de Rio/Córrego | ≤ 3.00m |
| LN2 | Crista de Serra/Espigão | ≤ 3.00m |
| LN3 | Margem de Rio/Córrego | ≤ 3.00m |
| LN4 | Margem de Lago/Lagoa | ≤ 3.00m |
| LN5 | Margem de Oceano | ≤ 3.00m |
| LN6 | Limite Seco de Praia/Mangue | ≤ 3.00m |

---

## 7. INTEGRAÇÃO COM SISTEMA EXISTENTE

### 7.1 Atualizar M_App_Logica.bas

Adicionar validações no processo de pós-importação:

```vba
Public Sub Processo_PosImportacao()
    ' ... código existente ...

    ' ADICIONAR: Validação dos dados importados
    Call Validar_DadosImportados()

    ' ... restante do código ...
End Sub

Private Sub Validar_DadosImportados()
    Dim ws As Worksheet, tbl As ListObject
    Dim i As Long, qtdErros As Long
    Dim msgErro As String, relatorioErros As String

    Set ws = ThisWorkbook.Sheets(M_Config.App_GetNomeAbaAtiva())
    Set tbl = ws.ListObjects(M_Config.App_GetNomeTabelaAtiva())

    If tbl.ListRows.Count = 0 Then Exit Sub

    For i = 1 To tbl.ListRows.Count
        ' Lê campos (ajustar índices conforme estrutura real)
        Dim tipo As String, codLimite As String
        Dim precisaoH As Double, precisaoV As Double
        Dim metodo As String

        tipo = tbl.DataBodyRange(i, 8).Value  ' Coluna "Tipo"
        codLimite = tbl.DataBodyRange(i, 11).Value  ' Coluna "Cod. Limite"
        precisaoH = tbl.DataBodyRange(i, 12).Value  ' Coluna "Precisao H"
        precisaoV = tbl.DataBodyRange(i, 13).Value  ' Coluna "Precisao V"
        metodo = tbl.DataBodyRange(i, 14).Value  ' Coluna "Metodo Posic."

        If Not M_Validacao.Validar_RegistroCompleto(tipo, codLimite, precisaoH, precisaoV, metodo, msgErro) Then
            qtdErros = qtdErros + 1
            relatorioErros = relatorioErros & "Linha " & i & ": " & msgErro & vbCrLf
        End If
    Next i

    If qtdErros > 0 Then
        MsgBox "Foram encontrados " & qtdErros & " erros de validacao INCRA:" & vbCrLf & vbCrLf & _
               relatorioErros, vbExclamation, "Validacao INCRA"
    Else
        MsgBox "Todos os dados estao conformes com o Manual INCRA!", vbInformation
    End If
End Sub
```

### 7.2 Atualizar M_DOC_Memorial.bas

Adicionar informações de método de posicionamento no memorial:

```vba
' No rodapé do memorial (linha ~87)
TextoMemorial = TextoMemorial & vbCrLf & vbTab & _
    "Todas as coordenadas aqui descritas estão georreferenciadas ao Sistema Geodésico " & _
    "Brasileiro tendo como datum o SIRGAS2000. A área foi obtida pelas coordenadas " & _
    "cartesianas locais, referenciada ao Sistema Geodésico Local (SGL-SIGEF). " & _
    "Todos os azimutes foram calculados pela fórmula do Problema Geodésico Inverso (Puissant). " & _
    "Perímetro e Distâncias foram calculados pelas coordenadas cartesianas geocêntricas." & vbCrLf

' ADICIONAR: Informação sobre método de posicionamento
Dim metodoUtilizado As String
metodoUtilizado = ObterMetodoPrevalente()  ' Função a criar

TextoMemorial = TextoMemorial & vbTab & _
    "Método de Posicionamento: " & M_Validacao.Obter_DescricaoMetodo(metodoUtilizado) & vbCrLf
```

---

## 8. TESTES

### 8.1 Executar Suite de Testes

```vba
Sub TestarValidacoes()
    Call Teste_Validacoes_INCRA.ExecutarTodosTestes()
End Sub
```

Resultado esperado:
- Arquivo `resultado_testes_incra.txt` criado
- Todos os testes devem passar (0 falhas)

### 8.2 Testes Manuais Recomendados

1. **Teste de Tipo de Vértice:**
   - Inserir M → deve aceitar
   - Inserir X → deve rejeitar

2. **Teste de Código de Limite:**
   - Inserir LA1 → deve aceitar
   - Inserir LA8 → deve rejeitar

3. **Teste de Precisão:**
   - LA1 com 0.30m → deve aceitar
   - LA1 com 0.80m → deve rejeitar
   - LN1 com 2.50m → deve aceitar
   - LN1 com 3.50m → deve rejeitar

4. **Teste de Método:**
   - GNSS-RTK → deve aceitar
   - INVALIDO → deve rejeitar

---

## 9. REFERÊNCIAS DO MANUAL INCRA

### Capítulo 1.4.3 - Métodos de Posicionamento
Métodos aceitos para determinação de coordenadas.

### Capítulo 1.4.4 - Precisão Posicional
Requisitos de precisão por tipo de limite:
- Limites artificiais (LA1-LA4): ≤ 0.50m
- Limites naturais (LN1-LN6): ≤ 3.00m
- Limites inacessíveis (LA5-LA7): ≤ 7.50m

### Capítulo 1.5 - Tipos de Vértices
- **M (Marco):** Vértice materializado
- **P (Ponto):** Feição identificável
- **V (Virtual):** Calculado

### Capítulo 2 - Limites e Confrontações
Classificação oficial de limites artificiais e naturais.

### Capítulo 3 - Métodos de Posicionamento
Descrição detalhada de cada método aceito.

---

## 📞 SUPORTE

Para questões sobre implementação:
1. Consulte o arquivo `RELATORIO_CONFORMIDADE_INCRA.md`
2. Execute `Setup_VerificarEstruturaDados()` para diagnosticar problemas
3. Execute `ExecutarTodosTestes()` para validar instalação

---

**Documento gerado automaticamente pelo Sistema DocGEO**
**Versão: 1.0 | Data: 27/12/2024**
