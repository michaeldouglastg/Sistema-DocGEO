VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmGerenciadorDB 
   Caption         =   "GERENCIAMENTO DO BANCO DE DADOS"
   ClientHeight    =   7785
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10695
   OleObjectBlob   =   "frmGerenciadorDB.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmGerenciadorDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' --- VARIÁVEIS DE CONTROLE ---
Private wsProp As Worksheet
Private wsTec As Worksheet
Private loProp As ListObject
Private loTec As ListObject
Private LinhaAtiva As Long ' Armazena a linha sendo editada

' ==============================================================================
' 1. INICIALIZAÇÃO E NAVEGAÇÃO
' ==============================================================================
Private Sub UserForm_Initialize()
    ' Configurar Referências às Tabelas
    Set wsProp = ThisWorkbook.Sheets("BD_PROPRIEDADES")
    Set wsTec = ThisWorkbook.Sheets("BD_TECNICOS")
    
    ' Assume que suas tabelas se chamam "TabelaPropriedades" e "TabelaTecnicos"
    ' Se não forem tabelas oficiais (Ctrl+T), ajustaremos para Range normal depois
    On Error Resume Next
    Set loProp = wsProp.ListObjects(1)
    Set loTec = wsTec.ListObjects(1)
    On Error GoTo 0
    
    ' Configurações Visuais Iniciais
    Me.mpgGerenciador.Value = 0 ' Começa no Menu
    Me.mpgGerenciador.Style = fmTabStyleNone ' Esconde as abas superiores (navegação por botão)
    
    ' Carrega as Listas
    Call AtualizarListaProp
    Call AtualizarListaTec
    
    ' Define Estado Inicial
    Call ModoCriacao("Prop")
    Call ModoCriacao("Tec")
End Sub

' Botões do Menu (Page 0)
Private Sub cmdIrParaPropriedades_Click()
    Me.mpgGerenciador.Value = 1
End Sub

Private Sub cmdIrParaTecnicos_Click()
    Me.mpgGerenciador.Value = 2
End Sub

' Botões de Voltar
Private Sub cmdVoltarProp_Click()
    Me.mpgGerenciador.Value = 0
End Sub
Private Sub cmdVoltarTec_Click()
    Me.mpgGerenciador.Value = 0
End Sub

' ==============================================================================
' 2. MÉTODOS AUXILIARES (Design Patterns: State Machine)
' ==============================================================================
Private Sub ModoCriacao(tipo As String)
    ' Configura visual para Novo Registro (AZUL)
    LinhaAtiva = 0
    
    If tipo = "Prop" Then
        ' Limpa Campos
'        Me.cboBuscaProp.Value = ""
'        Me.txtDenominacao.Value = ""
'        Me.txtProprietario.Value = ""
'        Me.txtMatricula.Value = ""
'        Me.txtDenominacao.Value = ""
        
        Me.txtMatricula.Value = ""
        Me.txtCodIncra.Value = ""
        Me.txtNaturezaArea.Value = ""
        Me.txtEndereco1.Value = ""
        Me.txtMunicipio.Value = ""
        Me.txtComarca.Value = ""
        Me.txtCartorio.Value = ""
        Me.txtCartorioCNS.Value = ""
        Me.txtProprietario.Value = ""
        Me.txtCPF.Value = ""
        Me.txtRG.Value = ""
        Me.txtExpedicao.Value = ""
        Me.txtDataExpedicao.Value = ""
        Me.txtNacionalidade.Value = ""
        Me.txtEstadoCivil.Value = ""
        'Me.txtProfissao.Value = ""
        Me.txtEndereco2.Value = ""
        
        ' Garante scroll no topo
        Me.fraDetalhesProp.ScrollTop = 0
        
        ' Visual
        Me.lblStatusProp.Caption = "MODO: NOVO REGISTRO"
        Me.lblStatusProp.BackColor = RGB(0, 120, 215) ' Azul
        Me.lblStatusProp.ForeColor = vbWhite
        Me.cmdExcluirProp.Enabled = False
        Me.lstPropriedades.ListIndex = -1 ' Tira seleção
    Else
        ' Lógica igual para Técnico...
        Me.cboBuscaTec.Value = ""
        Me.txtNomeTec.Value = ""
        ' ...
        Me.lblStatusTec.Caption = "MODO: NOVO REGISTRO"
        Me.lblStatusTec.BackColor = RGB(0, 120, 215)
        Me.lblStatusTec.ForeColor = vbWhite
        Me.cmdExcluirTec.Enabled = False
        Me.lstTecnicos.ListIndex = -1
    End If
End Sub

Private Sub ModoEdicao(tipo As String)
    ' Configura visual para Edição (AMARELO)
    If tipo = "Prop" Then
        Me.lblStatusProp.Caption = "MODO: EDIÇÃO"
        Me.lblStatusProp.BackColor = RGB(255, 192, 0) ' Amarelo
        Me.lblStatusProp.ForeColor = vbBlack
        Me.cmdExcluirProp.Enabled = True
    Else
        Me.lblStatusTec.Caption = "MODO: EDIÇÃO"
        Me.lblStatusTec.BackColor = RGB(255, 192, 0)
        Me.lblStatusTec.ForeColor = vbBlack
        Me.cmdExcluirTec.Enabled = True
    End If
End Sub

' ==============================================================================
' 3. LÓGICA DE PROPRIEDADES (PAGE 1)
' ==============================================================================

' Carrega ListBox e ComboBox
Private Sub AtualizarListaProp()
    Dim arrDados As Variant
    
    ' Pega dados da tabela (sem cabeçalho)
    If loProp.ListRows.Count > 0 Then
        arrDados = loProp.DataBodyRange.Value
        
        ' Preenche ListBox
        Me.lstPropriedades.List = arrDados
        
        ' Preenche ComboBox (Coluna 1 = Nome da Propriedade)
        ' Dica: Para combobox, podemos pegar só a coluna específica
        Me.cboBuscaProp.List = Application.Index(arrDados, 0, 1)
    Else
        Me.lstPropriedades.Clear
        Me.cboBuscaProp.Clear
    End If
End Sub

' Ao Clicar na Lista -> Preenche Campos e entra em Edição
Private Sub lstPropriedades_Click0()
    If Me.lstPropriedades.ListIndex = -1 Then Exit Sub
    
    Dim idx As Long
    idx = Me.lstPropriedades.ListIndex
    
    ' Guarda a linha real do Excel para salvar depois
    ' (Header + Index + 1) -> Ajuste conforme sua planilha
    LinhaAtiva = idx + 1
    
    ' Preenche TextBoxes com base na Matriz do Listbox
    ' Column(0) = Coluna A, Column(1) = Coluna B...
    Me.txtDenominacao.Value = Me.lstPropriedades.List(idx, 0)
    Me.cboBuscaProp.Value = Me.lstPropriedades.List(idx, 0) ' Sincroniza Combo
    Me.txtProprietario.Value = Me.lstPropriedades.List(idx, 1)
    Me.txtMatricula.Value = Me.lstPropriedades.List(idx, 2)
    ' ... preencher o resto ...
    
    Call ModoEdicao("Prop")
End Sub

Private Sub lstPropriedades_Click()
    ' Se nada estiver selecionado, sai
    If Me.lstPropriedades.ListIndex = -1 Then Exit Sub
    
    Dim idx As Long
    idx = Me.lstPropriedades.ListIndex
    
    ' Guarda a linha real da tabela para salvar/editar depois
    LinhaAtiva = idx + 1
    
    ' Referência à linha da tabela para leitura direta
    ' Isso garante que datas e números venham corretos, sem formatação de texto do ListBox
    Dim rngLinha As Range
    Set rngLinha = loProp.ListRows(LinhaAtiva).Range
    
    ' Desliga tratamento de erro temporariamente para campos vazios ou nomes errados
    On Error Resume Next
    
    Me.txtDenominacao.Value = rngLinha.Cells(1, 1).Value
    Me.txtMatricula.Value = rngLinha.Cells(1, 2).Value
    Me.txtCodIncra.Value = rngLinha.Cells(1, 3).Value
    Me.txtNaturezaArea.Value = rngLinha.Cells(1, 4).Value
    Me.txtEndereco1.Value = rngLinha.Cells(1, 5).Value
    Me.txtMunicipio.Value = rngLinha.Cells(1, 6).Value
    Me.txtComarca.Value = rngLinha.Cells(1, 7).Value
    Me.txtCartorio.Value = rngLinha.Cells(1, 8).Value
    Me.txtCartorioCNS.Value = rngLinha.Cells(1, 9).Value
    Me.txtProprietario.Value = rngLinha.Cells(1, 10).Value
    Me.txtCPF.Value = rngLinha.Cells(1, 11).Value
    Me.txtRG.Value = rngLinha.Cells(1, 12).Value
    Me.txtExpedicao.Value = rngLinha.Cells(1, 13).Value
    Me.txtDataExpedicao.Value = rngLinha.Cells(1, 14).Value
    Me.txtNacionalidade.Value = rngLinha.Cells(1, 15).Value
    Me.txtEstadoCivil.Value = rngLinha.Cells(1, 16).Value
    'Me.txtProfissao.Value = rngLinha.Cells(1, 17).Value ' Se o textbox existir
    Me.txtEndereco2.Value = rngLinha.Cells(1, 18).Value
    
    ' Sincroniza o Combo
    Me.cboBuscaProp.Value = Me.txtDenominacao.Value
    
    On Error GoTo 0
    
    ' Muda visual para MODO EDIÇÃO
    Call ModoEdicao("Prop")
End Sub

' Ao Selecionar no ComboBox -> Busca na Lista
Private Sub cboBuscaProp_Change()
    Dim i As Long
    If Me.cboBuscaProp.ListIndex > -1 Then
        ' Encontra item correspondente na ListBox
        Me.lstPropriedades.Selected(Me.cboBuscaProp.ListIndex) = True
        ' O evento lstPropriedades_Click será disparado automaticamente aqui
    End If
End Sub

' Botão NOVO
Private Sub cmdNovoProp_Click()
    Call ModoCriacao("Prop")
    
    Me.fraDetalhesProp.ScrollTop = 0
    Me.txtDenominacao.SetFocus
End Sub

' Botão SALVAR (Inteligente: Cria ou Edita)
Private Sub cmdSalvarProp_Click0()
    ' Validação básica
    If Me.txtDenominacao.Text = "" Then
        MsgBox "Nome da Propriedade obrigatório!", vbExclamation
        Exit Sub
    End If
    
    Dim rngDestino As Range
    
    If LinhaAtiva = 0 Then
        ' MODO CRIAR: Adiciona nova linha na tabela
        Set rngDestino = loProp.ListRows.Add.Range
    Else
        ' MODO EDITAR: Aponta para a linha existente
        Set rngDestino = loProp.ListRows(LinhaAtiva).Range
    End If
    
    ' Salva dados
    With rngDestino
        .Cells(1, 1).Value = Me.txtDenominacao.Value
        .Cells(1, 2).Value = Me.txtProprietario.Value
        .Cells(1, 3).Value = Me.txtMatricula.Value
        ' ... salvar resto ...
    End With
    
    MsgBox "Salvo com sucesso!", vbInformation
    Call AtualizarListaProp
    Call ModoCriacao("Prop") ' Reseta para novo
End Sub

Private Sub cmdSalvarProp_Click()
    ' 1. Validação Obrigatória (Exemplo: Denominação e Proprietário)
    If Me.txtDenominacao.Text = "" Or Me.txtProprietario.Text = "" Then
        MsgBox "Os campos 'Denominação' e 'Proprietário' são obrigatórios!", vbExclamation
        Exit Sub
    End If
    
    Dim rngDestino As Range
    
    ' 2. Define onde salvar (Nova linha ou Linha existente)
    If LinhaAtiva = 0 Then
        ' MODO CRIAR: Adiciona nova linha na tabela
        Set rngDestino = loProp.ListRows.Add.Range
    Else
        ' MODO EDITAR: Aponta para a linha que estamos editando
        Set rngDestino = loProp.ListRows(LinhaAtiva).Range
    End If
    
    ' 3. Salva dados nas células (Mapeamento 1 a 18)
    On Error Resume Next
    With rngDestino
        .Cells(1, 1).Value = Me.txtDenominacao.Value
        .Cells(1, 2).Value = Me.txtMatricula.Value
        .Cells(1, 3).Value = Me.txtCodIncra.Value
        .Cells(1, 4).Value = Me.txtNaturezaArea.Value
        .Cells(1, 5).Value = Me.txtEndereco1.Value
        .Cells(1, 6).Value = Me.txtMunicipio.Value
        .Cells(1, 7).Value = Me.txtComarca.Value
        .Cells(1, 8).Value = Me.txtCartorio.Value
        .Cells(1, 9).Value = Me.txtCartorioCNS.Value
        .Cells(1, 10).Value = Me.txtProprietario.Value
        .Cells(1, 11).Value = Me.txtCPF.Value
        .Cells(1, 12).Value = Me.txtRG.Value
        .Cells(1, 13).Value = Me.txtExpedicao.Value
        
        ' Tratamento para Data (evita salvar texto se estiver vazio)
        If IsDate(Me.txtDataExpedicao.Value) Then
            .Cells(1, 14).Value = CDate(Me.txtDataExpedicao.Value)
        Else
            .Cells(1, 14).Value = Me.txtDataExpedicao.Value
        End If
        
        .Cells(1, 15).Value = Me.txtNacionalidade.Value
        .Cells(1, 16).Value = Me.txtEstadoCivil.Value
        '.Cells(1, 17).Value = Me.txtProfissao.Value
        .Cells(1, 18).Value = Me.txtEndereco2.Value
    End With
    On Error GoTo 0
    
    MsgBox "Propriedade salva com sucesso!", vbInformation
    
    ' 4. Atualiza a interface
    Call AtualizarListaProp
    Call ModoCriacao("Prop") ' Reseta e limpa os campos
End Sub

' Botão EXCLUIR
Private Sub cmdExcluirProp_Click()
    If LinhaAtiva = 0 Then Exit Sub
    
    If MsgBox("Tem certeza que deseja excluir este registro?", vbQuestion + vbYesNo) = vbYes Then
        loProp.ListRows(LinhaAtiva).Delete
        Call AtualizarListaProp
        Call ModoCriacao("Prop")
    End If
End Sub





' ==============================================================================
' 4. LÓGICA DE TÉCNICOS (PAGE 2) - GESTÃO BD_TECNICOS
' ==============================================================================

' --- Carrega ListBox e ComboBox de Técnicos ---
Private Sub AtualizarListaTec()
    Dim arrDados As Variant
    
    ' Limpa antes de carregar
    Me.lstTecnicos.Clear
    Me.cboBuscaTec.Clear
    
    ' Verifica se a tabela existe e tem dados
    If Not loTec Is Nothing Then
        If loTec.ListRows.Count > 0 Then
            ' Pega dados da tabela (sem cabeçalho)
            arrDados = loTec.DataBodyRange.Value
            
            ' Preenche ListBox com todas as colunas
            Me.lstTecnicos.List = arrDados
            
            ' Preenche ComboBox apenas com a Coluna 1 (Nome do Técnico)
            ' Application.Index extrai apenas a coluna 1 da matriz
            Me.cboBuscaTec.List = Application.Index(arrDados, 0, 1)
        End If
    End If
End Sub

' --- Ao Clicar na Lista -> Preenche Campos e entra em Edição ---
Private Sub lstTecnicos_Click()
    ' Se nada estiver selecionado, sai
    If Me.lstTecnicos.ListIndex = -1 Then Exit Sub
    
    Dim idx As Long
    idx = Me.lstTecnicos.ListIndex
    
    ' Guarda a linha real da tabela para salvar/editar depois
    LinhaAtiva = idx + 1
    
    ' Preenche TextBoxes com base na Matriz do Listbox
    ' Ajuste os índices (0, 1, 2...) conforme a ordem das colunas no seu Excel
    On Error Resume Next ' Evita erro se a coluna estiver vazia
    Me.txtNomeTec.Value = Me.lstTecnicos.List(idx, 0)      ' Coluna A
    Me.cboBuscaTec.Value = Me.lstTecnicos.List(idx, 0)     ' Sincroniza Combo
    Me.txtFormacaoTec.Value = Me.lstTecnicos.List(idx, 1)  ' Coluna B
    Me.txtRegistroTec.Value = Me.lstTecnicos.List(idx, 2)  ' Coluna C
    Me.txtEmailTec.Value = Me.lstTecnicos.List(idx, 3)     ' Coluna D
    Me.txtTelefoneTec.Value = Me.lstTecnicos.List(idx, 4)  ' Coluna E
    On Error GoTo 0
    
    ' Muda visual para AMARELO (Edição)
    Call ModoEdicao("Tec")
End Sub

' --- Ao Selecionar no ComboBox -> Busca na Lista ---
Private Sub cboBuscaTec_Change()
    ' Se selecionar algo no combo, força a seleção na lista
    If Me.cboBuscaTec.ListIndex > -1 Then
        Me.lstTecnicos.Selected(Me.cboBuscaTec.ListIndex) = True
        ' O evento lstTecnicos_Click será disparado automaticamente aqui
    End If
End Sub

' --- Botão NOVO (Limpa tudo para inserir) ---
Private Sub cmdNovoTec_Click()
    Call ModoCriacao("Tec") ' Reseta visual para AZUL
    Me.txtNomeTec.SetFocus
End Sub

' --- Botão SALVAR (Inteligente: Cria ou Edita) ---
Private Sub cmdSalvarTec_Click()
    ' 1. Validação Obrigatória
    If Me.txtNomeTec.Text = "" Then
        MsgBox "O Nome do Técnico é obrigatório!", vbExclamation
        Me.txtNomeTec.SetFocus
        Exit Sub
    End If
    
    Dim rngDestino As Range
    
    ' 2. Define onde salvar (Nova linha ou Linha existente)
    If LinhaAtiva = 0 Then
        ' MODO CRIAR: Adiciona nova linha na tabela
        Set rngDestino = loTec.ListRows.Add.Range
    Else
        ' MODO EDITAR: Aponta para a linha que estamos editando
        Set rngDestino = loTec.ListRows(LinhaAtiva).Range
    End If
    
    ' 3. Salva dados nas células
    ' Ajuste a ordem conforme suas colunas no Excel
    With rngDestino
        .Cells(1, 1).Value = Me.txtNomeTec.Value
        .Cells(1, 2).Value = Me.txtFormacaoTec.Value
        .Cells(1, 3).Value = Me.txtRegistroTec.Value
        .Cells(1, 4).Value = Me.txtEmailTec.Value
        .Cells(1, 5).Value = Me.txtTelefoneTec.Value
    End With
    
    MsgBox "Técnico salvo com sucesso!", vbInformation
    
    ' 4. Atualiza a interface
    Call AtualizarListaTec
    Call ModoCriacao("Tec") ' Volta para modo de criação
End Sub

' --- Botão EXCLUIR ---
Private Sub cmdExcluirTec_Click()
    ' Só funciona se estiver editando alguém (LinhaAtiva > 0)
    If LinhaAtiva = 0 Then Exit Sub
    
    If MsgBox("Tem certeza que deseja excluir o técnico '" & Me.txtNomeTec.Value & "'?", _
              vbQuestion + vbYesNo, "Excluir Técnico") = vbYes Then
              
        ' Deleta a linha da tabela
        loTec.ListRows(LinhaAtiva).Delete
        
        ' Atualiza tudo
        Call AtualizarListaTec
        Call ModoCriacao("Tec")
        
        MsgBox "Registro excluído.", vbInformation
    End If
End Sub
