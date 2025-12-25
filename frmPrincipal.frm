VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPrincipal 
   Caption         =   "Pr�-Visualiza��o do Documento"
   ClientHeight    =   11700
   ClientLeft      =   30
   ClientTop       =   75
   ClientWidth     =   32385
   OleObjectBlob   =   "frmPrincipal.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ==============================================================================
' USERFORM: frmPrincipal
' DESCRICAO: FORMULARIO PRINCIPAL DE GERACAO DE DOCUMENTOS
' Cole este codigo no UserForm frmPrincipal
' ==============================================================================
Option Explicit

Public Status As String
Public Opcao As String
Public Boleano As Boolean

'Private pathLogo As String
'Private pathMapaLocal As String
'Private pathRegua As String

Private pathLogo As String
Private pathMapaLocal As String
Private pathRosa As String
Private pathConvencoes As String

' ==============================================================================
' INICIALIZACAO
' ==============================================================================
Private Sub UserForm_Initialize()
    Me.Height = 500
    Me.Width = 950
    Me.fraMenu.Height = 475
    Me.fraMenu.Width = 280
    Me.lblStatus.Caption = Status
    
    Call PopularComboBoxPropriedades
    Call PopularComboBoxTecnicos
    
    Me.fraDocumento.Visible = True
    Me.fraPropriedade.Visible = False
    Me.fraProprietario.Visible = False
    Me.fraTecnico.Visible = False
    Me.fraConfrontante.Visible = False
    Me.fraMapa.Visible = False
    
    Call LimparPreview
    Me.optMemorial.Value = False
    Me.cmbVoltar.Visible = False
    Me.cmbAvancar.Visible = False
    Me.cmdCadastrar.Visible = False
    Me.cmbNovo.Visible = False
    Me.cmdGerarWord.Visible = False
    Me.cmdGerarPDF.Visible = False
    Me.fraGerador.Visible = False
    
    Me.imgLogo.PictureSizeMode = fmPictureSizeModeZoom
    Me.imgMapa.PictureSizeMode = fmPictureSizeModeZoom
End Sub

Private Sub cmbNovo_Click()
    Me.fraDocumento.Visible = True
    Me.fraPropriedade.Visible = False
    Me.fraProprietario.Visible = False
    Me.fraTecnico.Visible = False
    Me.fraConfrontante.Visible = False
    Me.fraMapa.Visible = False
    Me.cmbVoltar.Visible = False
    Me.cmbAvancar.Visible = True
    Me.cmdCadastrar.Visible = False
    Me.cmbNovo.Visible = False
    Me.cmdGerarWord.Visible = False
    Me.cmdGerarPDF.Visible = False
    Me.fraGerador.Visible = False
    
    Call LimparPreview
    Me.optMemorial.Value = False
    Me.optAnuencia.Value = False
    Me.optLaudo.Value = False
    Me.optRequerimento.Value = False
    Me.optTabela.Value = False
    Me.optMapa.Value = False
    
    Status = "fraDocumento"
    Me.lblStatus.Caption = Status
End Sub

' ==============================================================================
' CONFIGURACAO DE INTERFACE
' ==============================================================================
Private Sub ConfigurarInterface(frameAtivo As Object, Optional mostrarBotoesGerar As Boolean = False, Optional mostrarBotoesMapa As Boolean = False)
    Me.fraDocumento.Visible = False
    Me.fraPropriedade.Visible = False
    Me.fraProprietario.Visible = False
    Me.fraTecnico.Visible = False
    Me.fraConfrontante.Visible = False
    Me.fraMapa.Visible = False
    
    If Not frameAtivo Is Nothing Then
        frameAtivo.Visible = True
        frameAtivo.Top = 18
        frameAtivo.Left = 12
    End If
    
    Me.cmbVoltar.Visible = True
    Me.cmdCadastrar.Visible = False
    Me.cmbAvancar.Visible = True
    Me.cmbNovo.Visible = False
    
    Me.fraGerador.Visible = mostrarBotoesGerar
    Me.cmdGerarWord.Visible = mostrarBotoesGerar
    Me.cmdGerarPDF.Visible = mostrarBotoesGerar
    
    Me.cmdGerarPDFMapa.Visible = mostrarBotoesMapa
    Me.cmdGerarDXF.Visible = mostrarBotoesMapa
    
    If mostrarBotoesMapa Then
        Me.fraGerador.Visible = mostrarBotoesMapa
        Me.cmdGerarPDFMapa.Top = 12
        Me.cmdGerarDXF.Top = 12
        Me.cmbAvancar.Visible = False
    End If
    
    If mostrarBotoesGerar Then
        Me.cmdGerarWord.Top = 12
        Me.cmdGerarPDF.Top = 12
        Me.cmbAvancar.Visible = False
    End If
End Sub

' ==============================================================================
' NAVEGACAO - AVANCAR
' ==============================================================================
Private Sub cmbAvancar_Click()
    Me.lblStatus.Caption = Status
    
    Select Case Status
        Case "fraDocumento"
            Call ConfigurarInterface(Me.fraPropriedade)
            Status = "fraPropriedade"
            
        Case "fraPropriedade"
            Call ConfigurarInterface(Me.fraProprietario)
            Me.cmdCadastrar.Visible = True
            Status = "fraProprietario"
            
        Case "fraProprietario"
            Call ConfigurarInterface(Me.fraTecnico, mostrarBotoesGerar:=True)
            Me.cmdCadastrar.Visible = True
            Me.cmbAvancar.Visible = False
            Me.cmbNovo.Visible = True
            Status = "fraTecnico"
            
        Case "fraDocumento_02"
            Call ConfigurarInterface(Me.fraConfrontante)
            Status = "fraConfrontante_02"
            
        Case "fraConfrontante_02"
            Call ConfigurarInterface(Me.fraPropriedade)
            Me.cmdCadastrar.Visible = True
            Call GerarPreview
            Status = "fraPropriedade_02"
            
        Case "fraPropriedade_02"
            Call ConfigurarInterface(Me.fraProprietario)
            Me.cmdCadastrar.Visible = True
            Status = "fraProprietario_02"
            
        Case "fraProprietario_02"
            Call ConfigurarInterface(Me.fraTecnico, mostrarBotoesGerar:=True)
            Me.cmdCadastrar.Visible = True
            Me.cmbAvancar.Visible = False
            Me.cmbNovo.Visible = True
            Status = "fraTecnico_02"
            
        Case "fraDocumento_03"
            Call ConfigurarInterface(Me.fraMapa)
            Status = "fraMapa_03"
            
        Case "fraMapa_03"
            Call ConfigurarInterface(Me.fraPropriedade)
            Me.cmdCadastrar.Visible = True
            Status = "fraPropriedade_03"
            
        Case "fraPropriedade_03"
            Call ConfigurarInterface(Me.fraProprietario)
            Me.cmdCadastrar.Visible = True
            Status = "fraProprietario_03"
            
        Case "fraProprietario_03"
            Call ConfigurarInterface(Me.fraTecnico, mostrarBotoesMapa:=True)
            Me.cmdCadastrar.Visible = True
            Me.cmbAvancar.Visible = False
            Me.cmbNovo.Visible = True
            Status = "fraTecnico_03"
    End Select
    
    Me.lblStatus.Caption = Status
End Sub

' ==============================================================================
' NAVEGACAO - VOLTAR
' ==============================================================================
Private Sub cmbVoltar_Click()
    Me.lblStatus.Caption = Status
    
    Select Case Status
        Case "fraPropriedade"
            Call ConfigurarInterface(Me.fraDocumento)
            Me.cmbVoltar.Visible = False
            Me.cmdCadastrar.Visible = False
            Status = "fraDocumento"
            
        Case "fraConfrontante_02"
            Call ConfigurarInterface(Me.fraDocumento)
            Me.cmbVoltar.Visible = False
            Me.cmdCadastrar.Visible = False
            Status = "fraDocumento_02"
            
        Case "fraMapa_03"
            Call ConfigurarInterface(Me.fraDocumento)
            Me.cmbVoltar.Visible = False
            Me.cmdCadastrar.Visible = False
            Status = "fraDocumento_03"
            
        Case "fraProprietario"
            Call ConfigurarInterface(Me.fraPropriedade)
            Me.cmdCadastrar.Visible = True
            Status = "fraPropriedade"
            
        Case "fraPropriedade_02"
            Call ConfigurarInterface(Me.fraConfrontante)
            Status = "fraConfrontante_02"
            
        Case "fraPropriedade_03"
            Call ConfigurarInterface(Me.fraMapa)
            Status = "fraMapa_03"
            
        Case "fraTecnico"
            Call ConfigurarInterface(Me.fraProprietario)
            Me.cmdCadastrar.Visible = True
            Me.fraGerador.Visible = False
            Status = "fraProprietario"
            
        Case "fraProprietario_02"
            Call ConfigurarInterface(Me.fraPropriedade)
            Me.cmdCadastrar.Visible = True
            Status = "fraPropriedade_02"
            
        Case "fraProprietario_03"
            Call ConfigurarInterface(Me.fraPropriedade)
            Me.cmdCadastrar.Visible = True
            Status = "fraPropriedade_03"
            
        Case "fraTecnico_02"
            Call ConfigurarInterface(Me.fraProprietario)
            Me.cmdCadastrar.Visible = True
            Me.fraGerador.Visible = False
            Status = "fraProprietario_02"
            
        Case "fraTecnico_03"
            Call ConfigurarInterface(Me.fraProprietario)
            Me.cmdCadastrar.Visible = True
            Status = "fraProprietario_03"
            
        Case "fraGerar_03"
            Call ConfigurarInterface(Me.fraTecnico)
            Me.cmdCadastrar.Visible = True
            Me.cmbAvancar.Visible = True
            Status = "fraTecnico_03"
    End Select
    
    Me.lblStatus.Caption = Status
End Sub

' ==============================================================================
' SELECAO DE TIPO DE DOCUMENTO
' ==============================================================================
Private Sub optMemorial_Click()
    Me.lblTitulo.Caption = Me.optMemorial.Caption
    Me.fraConfrontante.Visible = False
    Me.fraDocumento.Visible = False
    Me.fraPropriedade.Top = 18
    Me.fraPropriedade.Left = 12
    Me.fraPropriedade.Visible = True
    Me.cmbVoltar.Visible = True
    Me.cmbAvancar.Visible = True
    Me.cmdCadastrar.Visible = True
    Status = "fraPropriedade"
    Call LimparPreview
    Call GerarPreview
    Me.lblStatus.Caption = Status
End Sub

Private Sub optAnuencia_Click()
    Me.lblTitulo.Caption = Me.optAnuencia.Caption
    Me.fraDocumento.Visible = False
    Me.fraConfrontante.Visible = True
    Me.fraConfrontante.Top = 18
    Me.fraConfrontante.Left = 12
    Me.optGerarIndividual.Value = True
    Me.fraGerarIndividual.Visible = True
    Me.fraGerarTodos.Visible = False
    Me.cmbVoltar.Visible = True
    Me.cmbAvancar.Visible = True
    Status = "fraConfrontante_02"
    Call PopularComboBoxConfrontantes
    Call LimparPreview
    Me.lblStatus.Caption = Status
End Sub

Private Sub optLaudo_Click()
    Me.lblTitulo.Caption = Me.optLaudo.Caption
    Me.fraConfrontante.Visible = False
    Me.fraDocumento.Visible = False
    Me.fraPropriedade.Top = 18
    Me.fraPropriedade.Left = 12
    Me.fraPropriedade.Visible = True
    Me.cmbVoltar.Visible = True
    Me.cmbAvancar.Visible = True
    Me.cmdCadastrar.Visible = True
    Status = "fraPropriedade"
    Call LimparPreview
    Call GerarPreview
    Me.lblStatus.Caption = Status
End Sub

Private Sub optRequerimento_Click()
    Me.lblTitulo.Caption = Me.optRequerimento.Caption
    Me.fraConfrontante.Visible = False
    Me.fraDocumento.Visible = False
    Me.fraPropriedade.Top = 18
    Me.fraPropriedade.Left = 12
    Me.fraPropriedade.Visible = True
    Me.cmbVoltar.Visible = True
    Me.cmbAvancar.Visible = True
    Me.cmdCadastrar.Visible = True
    Status = "fraPropriedade"
    Call LimparPreview
    Call GerarPreview
    Me.lblStatus.Caption = Status
End Sub

Private Sub optTabela_Click()
    Me.lblTitulo.Caption = Me.optTabela.Caption
    Me.fraConfrontante.Visible = False
    Me.fraDocumento.Visible = False
    Me.fraPropriedade.Top = 18
    Me.fraPropriedade.Left = 12
    Me.fraPropriedade.Visible = True
    Me.cmbVoltar.Visible = True
    Me.cmbAvancar.Visible = True
    Me.cmdCadastrar.Visible = True
    Status = "fraPropriedade"
    Call LimparPreview
    Call GerarPreview
    Me.lblStatus.Caption = Status
End Sub

Private Sub optMapa_Click()
    Me.fraConfrontante.Visible = False
    Me.fraDocumento.Visible = False
    Me.fraPropriedade.Visible = False
    Me.fraProprietario.Visible = False
    Me.fraTecnico.Visible = False
    Me.fraMapa.Visible = True
    Me.fraMapa.Top = 18
    Me.fraMapa.Left = 12
    Me.cmbVoltar.Visible = False
    Me.cmdCadastrar.Visible = True
    Me.cmbAvancar.Visible = True
    Me.cmdGerarWord.Visible = False
    Me.cmdGerarPDF.Visible = False
    Me.cmdGerarPDFMapa.Visible = True
    Status = "fraMapa_03"
    
    Dim wsParam As Worksheet
    Set wsParam = ThisWorkbook.Sheets(M_Config.SH_PARAMETROS)
    Me.txtTitulo.Text = wsParam.Range("K1").Value
    Me.txtEscala1.Text = wsParam.Range("K2").Value
    Me.txtEscala2.Text = wsParam.Range("K3").Value
    pathLogo = wsParam.Range("K4").Value
    pathMapaLocal = wsParam.Range("K5").Value
    'pathRegua = wsParam.Range("K6").Value
    
    CarregarImagemPreview Me.imgLogo, pathLogo
    CarregarImagemPreview Me.imgMapa, pathMapaLocal
    Me.lblStatus.Caption = Status
End Sub

Private Sub optGerarIndividual_Click()
    Me.fraGerarTodos.Visible = False
    Me.fraGerarIndividual.Visible = True
    Me.fraGerarIndividual.Top = 84
End Sub

Private Sub optGerarTodos_Click()
    Me.fraGerarTodos.Visible = True
    Me.fraGerarIndividual.Visible = False
    If Me.lstConfrontantesMassa.ListCount = 0 Then
        Call PopularComboBoxConfrontantes
    End If
End Sub

' ==============================================================================
' GERACAO DE DOCUMENTOS
' ==============================================================================
Private Sub cmdGerarWord_Click()
    Dim dadosProp As Object, dadosTec As Object
    Set dadosProp = ColetarDadosPropriedade()
    Set dadosTec = ColetarDadosTecnico()
    
    If Me.optMemorial.Value Then
        Call GerarMemorialWord(dadosProp, dadosTec)
        
    ElseIf Me.optAnuencia.Value Then
        If Me.optGerarIndividual.Value = True Then
            If Me.cboConfrontantesAnuencia.ListIndex = -1 Then
                MsgBox "Selecione um confrontante.", vbExclamation
                Exit Sub
            End If
            frmAguarde.Show
            Call GerarCartaAnuencia(Me.cboConfrontantesAnuencia.Value, dadosProp, dadosTec)
            
        ElseIf Me.optGerarTodos.Value = True Then
            Dim pastaDestino As String
            With Application.FileDialog(msoFileDialogFolderPicker)
                .Title = "Selecione a pasta para salvar as Anuencias"
                If .Show = -1 Then pastaDestino = .SelectedItems(1) Else Exit Sub
            End With
            
            frmAguarde.Show
            Dim i As Long, cont As Long: cont = 0
            For i = 0 To Me.lstConfrontantesMassa.ListCount - 1
                If Me.lstConfrontantesMassa.Selected(i) Then
                    Call GerarCartaAnuencia(Me.lstConfrontantesMassa.List(i), dadosProp, dadosTec, False, pastaDestino)
                    cont = cont + 1
                End If
            Next i
            Unload frmAguarde
            MsgBox cont & " documentos gerados com sucesso na pasta:" & vbCrLf & pastaDestino, vbInformation
        End If
        
    ElseIf Me.optRequerimento.Value Then
        Call GerarRequerimentoWord(dadosProp, dadosTec)
        
    ElseIf Me.optLaudo.Value Then
        Call GerarLaudoTecnicoWord(dadosProp, dadosTec)
        
    ElseIf Me.optTabela.Value Then
        Call GerarTabelaAnaliticaWord(dadosProp, dadosTec)
    End If
End Sub

Private Sub cmdGerarPDF_Click()
    Dim dadosProp As Object, dadosTec As Object
    Set dadosProp = ColetarDadosPropriedade()
    Set dadosTec = ColetarDadosTecnico()
    
    If Me.optMemorial.Value Then
        Call GerarMemorialPDF(dadosProp, dadosTec)
        
    ElseIf Me.optAnuencia.Value Then
        If Me.optGerarIndividual.Value = True Then
            If Me.cboConfrontantesAnuencia.ListIndex = -1 Then
                MsgBox "Selecione um confrontante.", vbExclamation
                Exit Sub
            End If
            frmAguarde.Show
            Call GerarAnuenciaPDF(Me.cboConfrontantesAnuencia.Value, dadosProp, dadosTec)
            
        ElseIf Me.optGerarTodos.Value = True Then
            Dim pastaDestino As String
            With Application.FileDialog(msoFileDialogFolderPicker)
                .Title = "Selecione a pasta para salvar os PDFs"
                If .Show = -1 Then pastaDestino = .SelectedItems(1) Else Exit Sub
            End With
            
            frmAguarde.Show
            Dim i As Long, cont As Long: cont = 0
            For i = 0 To Me.lstConfrontantesMassa.ListCount - 1
                If Me.lstConfrontantesMassa.Selected(i) Then
                    Call GerarAnuenciaPDF(Me.lstConfrontantesMassa.List(i), dadosProp, dadosTec, pastaDestino)
                    cont = cont + 1
                End If
            Next i
            Unload frmAguarde
            MsgBox cont & " PDFs gerados com sucesso na pasta:" & vbCrLf & pastaDestino, vbInformation
        End If
        
    ElseIf Me.optRequerimento.Value Then
        Call GerarRequerimentoPDF(dadosProp, dadosTec)

    ElseIf Me.optLaudo.Value Then
        Call GerarLaudoTecnicoPDF(dadosProp, dadosTec)

    ElseIf Me.optTabela.Value Then
        Call GerarTabelaAnaliticaPDF(dadosProp, dadosTec)
    End If
End Sub

Private Sub cmdGerarDXF_Click()
    Dim dadosProp As Object
    Set dadosProp = ColetarDadosPropriedade()
    Call M_DOC_DXF.GerarArquivoDXF(dadosProp)
End Sub

' ==============================================================================
' PRE-VISUALIZACAO
' ==============================================================================
Private Sub GerarPreview()
    Dim dadosProp As Object, dadosTec As Object
    
    Set dadosProp = ColetarDadosPropriedade()
    Set dadosTec = ColetarDadosTecnico()
    
    Call LimparPreview
    
    If Me.optMemorial.Value = True Then
        Me.txtPreview.Text = M_DOC_Memorial.GerarTextoMemorial(dadosProp, dadosTec)
        
    ElseIf Me.optAnuencia.Value = True Then
        If Me.optGerarTodos.Value = True Then
            Me.txtPreview.Text = "PRE-VISUALIZACAO INDISPONIVEL EM LOTE." & vbCrLf & vbCrLf & _
                "Voce selecionou a opcao 'GERAR TODOS'." & vbCrLf & _
                "Para ver o texto de uma carta especifica, mude para 'GERAR INDIVIDUAL' e selecione o confrontante."
        Else
            If Me.cboConfrontantesAnuencia.ListIndex = -1 Then
                Me.txtPreview.Text = "Selecione um confrontante na lista acima para visualizar o documento."
                Exit Sub
            End If
            Me.txtPreview.Text = M_DOC_Anuencia.GerarTextoAnuencia(Me.cboConfrontantesAnuencia.Value, dadosProp, dadosTec)
        End If
        
    ElseIf Me.optRequerimento.Value = True Then
        Me.txtPreview.Text = M_DOC_Requerimento.GerarTextoRequerimento(dadosProp, dadosTec)
        
    ElseIf Me.optLaudo.Value = True Then
        Me.txtPreview.Text = M_DOC_Laudo.GerarTextoLaudo(dadosProp, dadosTec)
        
    ElseIf Me.optTabela.Value = True Then
        Me.txtPreview.Text = M_DOC_Tabela.GerarTextoTabelaAnalitica(dadosProp, dadosTec)
        
    ElseIf Me.optMapa.Value = True Then
        Me.txtPreview.Text = "PRE-VISUALIZACAO DO MAPA (A3/A1):" & vbCrLf & vbCrLf & _
            "Este documento sera gerado diretamente como um arquivo grafico (Word ou PDF) contendo:" & vbCrLf & _
            "- Desenho do Poligono (Croqui)" & vbCrLf & _
            "- Tabela de Coordenadas UTM" & vbCrLf & _
            "- Carimbo Tecnico e Imagens" & vbCrLf & vbCrLf & _
            "O texto nao e editavel nesta janela."
    End If
    
    If Len(Me.txtPreview.Text) > 10 And InStr(1, UCase(Me.txtPreview.Text), "ERRO") = 0 Then
        Me.cmdGerarWord.Enabled = True
        Me.cmdGerarPDF.Enabled = True
    ElseIf Me.optMapa.Value = True Then
        Me.cmdGerarWord.Enabled = True
        Me.cmdGerarPDF.Enabled = True
    Else
        Me.cmdGerarWord.Enabled = False
        Me.cmdGerarPDF.Enabled = False
    End If
End Sub

Private Sub LimparPreview()
    Me.txtPreview.Text = ""
    Me.cmdGerarWord.Enabled = False
    Me.cmdGerarPDF.Enabled = False
End Sub

Private Sub cmbAtualizar_Click()
    Call GerarPreview
End Sub

Private Sub cmdCopiarPreview_Click()
    If Me.txtPreview.Text = "" Then
        MsgBox "Nao ha texto na caixa de pre-visualizacao para copiar.", vbInformation, "Aviso"
        Exit Sub
    End If
    
    Dim clipboard As MSForms.DataObject
    Set clipboard = New MSForms.DataObject
    clipboard.SetText Me.txtPreview.Text
    clipboard.PutInClipboard
    Set clipboard = Nothing
    
    MsgBox "O texto da pre-visualizacao foi copiado com sucesso para a area de transferencia!", vbInformation, "Copiado"
End Sub

' ==============================================================================
' CADASTROS - PROPRIEDADE
' ==============================================================================
Private Sub cboBuscaPropriedade_Change()
    Dim ws As Worksheet, tbl As ListObject, foundRow As Range, busca As String
    busca = Me.cboBuscaPropriedade.Value
    If busca = "" Then Exit Sub
    
    Set ws = ThisWorkbook.Sheets("BD_PROPRIEDADES")
    Set tbl = ws.ListObjects("tbl_PropriedadesDB")
    Set foundRow = tbl.ListColumns(1).DataBodyRange.Find(busca, LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not foundRow Is Nothing Then
        Dim iRow As Long: iRow = foundRow.Row - tbl.HeaderRowRange.Row
        Me.txtDenominacao.Value = tbl.DataBodyRange.Cells(iRow, 1).Value
        Me.txtMatricula.Value = tbl.DataBodyRange.Cells(iRow, 2).Value
        Me.txtCodIncra.Value = tbl.DataBodyRange.Cells(iRow, 3).Value
        Me.txtNaturezaArea.Value = tbl.DataBodyRange.Cells(iRow, 4).Value
        Me.txtEndereco1.Value = tbl.DataBodyRange.Cells(iRow, 5).Value
        Me.txtMunicipio.Value = tbl.DataBodyRange.Cells(iRow, 6).Value
        Me.txtComarca.Value = tbl.DataBodyRange.Cells(iRow, 7).Value
        Me.txtCartorio.Value = tbl.DataBodyRange.Cells(iRow, 8).Value
        Me.txtCartorioCNS.Value = tbl.DataBodyRange.Cells(iRow, 9).Value
        Me.txtProprietario.Value = tbl.DataBodyRange.Cells(iRow, 10).Value
        Me.txtCPF.Value = tbl.DataBodyRange.Cells(iRow, 11).Value
        Me.txtRG.Value = tbl.DataBodyRange.Cells(iRow, 12).Value
        Me.txtExpedicao.Value = tbl.DataBodyRange.Cells(iRow, 13).Value
        Me.txtDataExpedicao.Value = tbl.DataBodyRange.Cells(iRow, 14).Value
        Me.txtNacionalidade.Value = tbl.DataBodyRange.Cells(iRow, 15).Value
        Me.txtEstadoCivil.Value = tbl.DataBodyRange.Cells(iRow, 16).Value
        'Me.txtProfissao.Value = tbl.DataBodyRange.Cells(iRow, 17).Value
        Me.txtEndereco2.Value = tbl.DataBodyRange.Cells(iRow, 18).Value
        cboBuscaProprietario.Text = cboBuscaPropriedade.Text
    End If
    
    Call LimparPreview
    Call GerarPreview
End Sub

Private Sub cmdCadastrarPropriedade_Click()
    Dim ws As Worksheet, tbl As ListObject, newRow As ListRow
    Set ws = ThisWorkbook.Sheets("BD_PROPRIEDADES")
    Set tbl = ws.ListObjects("tbl_PropriedadesDB")
    Set newRow = tbl.ListRows.Add
    
    With newRow
        .Range(1).Value = Me.txtDenominacao.Value
        .Range(2).Value = Me.txtMatricula.Value
        .Range(3).Value = Me.txtCodIncra.Value
        .Range(4).Value = Me.txtNaturezaArea.Value
        .Range(5).Value = Me.txtEndereco1.Value
        .Range(6).Value = Me.txtMunicipio.Value
        .Range(7).Value = Me.txtComarca.Value
        .Range(8).Value = Me.txtCartorio.Value
        .Range(9).Value = Me.txtCartorioCNS.Value
        .Range(10).Value = Me.txtProprietario.Value
        .Range(11).Value = Me.txtCPF.Value
        .Range(12).Value = Me.txtRG.Value
        .Range(13).Value = Me.txtExpedicao.Value
        .Range(14).Value = Me.txtDataExpedicao.Value
        .Range(15).Value = Me.txtNacionalidade.Value
        .Range(16).Value = Me.txtEstadoCivil.Value
        '.Range(17).Value = Me.txtProfissao.Value
        .Range(18).Value = Me.txtEndereco2.Value
    End With
    
    Call PopularComboBoxPropriedades
    Me.cboBuscaPropriedade.Value = Me.txtDenominacao.Value
    MsgBox "Propriedade '" & Me.txtDenominacao.Value & "' cadastrada com sucesso!", vbInformation
End Sub

' ==============================================================================
' CADASTROS - TECNICO
' ==============================================================================
Private Sub cboBuscaTecnico_Change()
    Dim ws As Worksheet, tbl As ListObject, foundRow As Range, busca As String
    busca = Me.cboBuscaTecnico.Value
    If busca = "" Then Exit Sub
    
    Set ws = ThisWorkbook.Sheets("BD_TECNICOS")
    Set tbl = ws.ListObjects("tbl_TecnicosDB")
    Set foundRow = tbl.ListColumns(1).DataBodyRange.Find(busca, LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not foundRow Is Nothing Then
        Dim iRow As Long: iRow = foundRow.Row - tbl.HeaderRowRange.Row
        Me.txtNomeTecnico.Value = tbl.DataBodyRange.Cells(iRow, 1).Value
        Me.txtFormacao.Value = tbl.DataBodyRange.Cells(iRow, 2).Value
        Me.txtRegistro.Value = tbl.DataBodyRange.Cells(iRow, 3).Value
        Me.txtCodIncraTecnico.Value = tbl.DataBodyRange.Cells(iRow, 4).Value
        Me.txtTRTART.Value = tbl.DataBodyRange.Cells(iRow, 5).Value
    End If
    
    Call LimparPreview
    Call GerarPreview
End Sub

Private Sub cmdCadastrarTecnico_Click()
    Dim ws As Worksheet, tbl As ListObject, newRow As ListRow
    Set ws = ThisWorkbook.Sheets("BD_TECNICOS")
    Set tbl = ws.ListObjects("tbl_TecnicosDB")
    Set newRow = tbl.ListRows.Add
    
    With newRow
        .Range(1).Value = Me.txtNomeTecnico.Value
        .Range(2).Value = Me.txtFormacao.Value
        .Range(3).Value = Me.txtRegistro.Value
        .Range(4).Value = Me.txtCodIncraTecnico.Value
        .Range(5).Value = Me.txtTRTART.Value
    End With
    
    Call PopularComboBoxTecnicos
    Me.cboBuscaTecnico.Value = Me.txtNomeTecnico.Value
    MsgBox "Responsavel Tecnico '" & Me.txtNomeTecnico.Value & "' cadastrado com sucesso!", vbInformation
End Sub

' ==============================================================================
' FUNCOES AUXILIARES
' ==============================================================================
Private Function ColetarDadosPropriedade() As Object
    Dim D As Object: Set D = CreateObject("Scripting.Dictionary")
    Dim wsPainel As Worksheet
    Set wsPainel = ThisWorkbook.Sheets("PAINEL_PRINCIPAL")
    
    Dim areaSGL As Variant
    On Error Resume Next
    areaSGL = wsPainel.Range("AreaSGL").Value
    If Err.Number <> 0 Then areaSGL = 0
    On Error GoTo 0
    
    D.Add "Denomina��o", Me.txtDenominacao.Value
    D.Add "Matr�cula", Me.txtMatricula.Value
    D.Add "C�d. Incra/SNCR", Me.txtCodIncra.Value
    D.Add "Natureza/�rea", Me.txtNaturezaArea.Value
    D.Add "Endere�o Propriedade", Me.txtEndereco1.Value
    D.Add "Munic�pio/UF", Me.txtMunicipio.Value
    D.Add "Comarca", Me.txtComarca.Value
    D.Add "Cart�rio", Me.txtCartorio.Value
    D.Add "Cart�rio (CNS)", Me.txtCartorioCNS.Value
    D.Add "Propriet�rio", Me.txtProprietario.Value
    D.Add "CPF", Me.txtCPF.Value
    D.Add "RG", Me.txtRG.Value
    D.Add "Expedi��o", Me.txtExpedicao.Value
    D.Add "Data Expedi��o", Me.txtDataExpedicao.Value
    D.Add "Nacionalidade", Me.txtNacionalidade.Value
    D.Add "Estado Civil", Me.txtEstadoCivil.Value
    D.Add "Endere�o Propriet�rio", Me.txtEndereco2.Value
    D.Add "Area (SGL)", areaSGL
    
    Set ColetarDadosPropriedade = D
End Function

Private Function ColetarDadosTecnico() As Object
    Dim D As Object: Set D = CreateObject("Scripting.Dictionary")
    D.Add "Nome do T�cnico", Me.txtNomeTecnico.Value
    D.Add "Forma��o", Me.txtFormacao.Value
    D.Add "Registro (CFT/CREA)", Me.txtRegistro.Value
    D.Add "C�d. Incra", Me.txtCodIncraTecnico.Value
    D.Add "TRT/ART", Me.txtTRTART.Value
    Set ColetarDadosTecnico = D
End Function

Private Sub PopularComboBoxPropriedades()
    Dim ws As Worksheet, tbl As ListObject, i As Long
    Set ws = ThisWorkbook.Sheets(M_Config.SH_BD_PROP)
    Set tbl = ws.ListObjects(M_Config.TBL_DB_PROP)
    
    Me.cboBuscaPropriedade.Clear
    Me.cboBuscaProprietario.Clear
    
    If Not tbl.DataBodyRange Is Nothing Then
        For i = 1 To tbl.ListRows.Count
            Me.cboBuscaPropriedade.AddItem tbl.DataBodyRange.Cells(i, 1).Value
            Me.cboBuscaProprietario.AddItem tbl.DataBodyRange.Cells(i, 1).Value
        Next i
    End If
End Sub

Private Sub PopularComboBoxTecnicos()
    Dim ws As Worksheet, tbl As ListObject, i As Long
    Set ws = ThisWorkbook.Sheets(M_Config.SH_BD_TEC)
    Set tbl = ws.ListObjects(M_Config.TBL_DB_TEC)
    
    Me.cboBuscaTecnico.Clear
    
    If Not tbl.DataBodyRange Is Nothing Then
        For i = 1 To tbl.ListRows.Count
            Me.cboBuscaTecnico.AddItem tbl.DataBodyRange.Cells(i, 1).Value
        Next i
    End If
End Sub

Private Sub PopularComboBoxConfrontantes()
    Dim ws As Worksheet, tbl As ListObject, i As Long
    Dim confrontante As String
    Dim dictConf As Object
    
    Set dictConf = CreateObject("Scripting.Dictionary")
    Set ws = ThisWorkbook.Sheets(M_Config.App_GetNomeAbaAtiva())
    Set tbl = ws.ListObjects(M_Config.App_GetNomeTabelaAtiva())
    
    Me.cboConfrontantesAnuencia.Clear
    Me.lstConfrontantesMassa.Clear
    
    If tbl.DataBodyRange Is Nothing Then Exit Sub
    
    For i = 1 To tbl.ListRows.Count
        confrontante = Trim(CStr(tbl.DataBodyRange.Cells(i, 8).Value))
        If confrontante <> "" And Not dictConf.Exists(confrontante) Then
            dictConf.Add confrontante, 1
            Me.cboConfrontantesAnuencia.AddItem confrontante
            Me.lstConfrontantesMassa.AddItem confrontante
        End If
    Next i
End Sub

Private Sub cboConfrontantesAnuencia_Change()
    Me.fraDocumento.Visible = False
    Me.fraConfrontante.Visible = False
    Me.fraPropriedade.Visible = True
    Me.fraPropriedade.Top = 18
    Me.fraPropriedade.Left = 12
    Me.cmbVoltar.Visible = True
    Me.cmdCadastrar.Visible = True
    Me.cmbAvancar.Visible = True
    Me.cmbNovo.Visible = True
    Status = "fraPropriedade_02"
    Call GerarPreview
    Me.lblStatus.Caption = Status
End Sub

Private Sub cmdCadastrar_Click()
    If Status = "fraTecnico" Then
        'MsgBox "Tecnico"
        frmGerenciadorDB.Show
    ElseIf Status = "fraPropriedade" Then
        'MsgBox "Propriedade"
        frmGerenciadorDB.Show
    ElseIf Status = "fraProprietario" Then
        'MsgBox "Propriet�rio"
        frmGerenciadorDB.Show
    End If
End Sub

Private Sub cmdResponsavel_Click()
    Me.fraProprietario.Visible = False
    Me.fraTecnico.Top = 18
    Me.fraTecnico.Left = 12
    Me.fraTecnico.Visible = True
End Sub

Private Sub cmdFechar_Click()
    Unload Me
End Sub

' ==============================================================================
' EVENTOS DE CLIQUE NAS IMAGENS (PARA CARREGAR ARQUIVO)
' Certifique-se de que os controles Image no formul�rio tenham estes nomes:
' imgLogo, imgMapa, imgRosa, imgConvencoes
' ==============================================================================

Private Sub imgLogo_Click()
    pathLogo = SelecionarImagemDialogo()
    If pathLogo <> "" Then CarregarImagemPreview Me.imgLogo, pathLogo
End Sub

Private Sub imgMapa_Click()
    pathMapaLocal = SelecionarImagemDialogo()
    If pathMapaLocal <> "" Then CarregarImagemPreview Me.imgMapa, pathMapaLocal
End Sub

Private Sub imgRosa_Click()
    pathRosa = SelecionarImagemDialogo()
    If pathRosa <> "" Then CarregarImagemPreview Me.imgRosa, pathRosa
End Sub

Private Sub imgConvencoes_Click()
    pathConvencoes = SelecionarImagemDialogo()
    If pathConvencoes <> "" Then CarregarImagemPreview Me.imgConvencoes, pathConvencoes
End Sub

' ==============================================================================
' BOT�O 1: GERAR MAPA NO EXCEL (DESENHO + IMAGENS)(COM TODOS OS CAMPOS)
' Nome sugerido para o bot�o: cmdGerarExcelMapa
' ==============================================================================
Private Sub cmdGerarExcelMapa_Click()
    ' 1. Valida��o Simples
    If Me.txtTitulo.Text = "" Then
        MsgBox "Por favor, digite um t�tulo para o mapa.", vbExclamation
        Me.txtTitulo.SetFocus
        Exit Sub
    End If

    ' 2. Coleta de Dados dos outros frames (Propriedade/T�cnico)
    Dim dadosProp As Object, dadosTec As Object
    Set dadosProp = ColetarDadosPropriedade()
    Set dadosTec = ColetarDadosTecnico()
    
    ' 3. Chama o M�dulo M_DOC_Mapa passando TUDO
    ' Ajuste os nomes "Me.txt..." abaixo conforme o nome real das suas caixas de texto
    Call M_DOC_Mapa.GerarMapaExcel(dadosProp, _
                                   dadosTec, _
                                   Me.txtTitulo.Text, _
                                   Me.txtEscala1.Text, _
                                   Me.txtEscala2.Text, _
                                   Me.txtEmail.Text, _
                                   Me.txtTelefone.Text, _
                                   Me.txtObsCoord.Text, _
                                   Me.txtObsDatum.Text, _
                                   Me.txtObsMeridiano.Text, _
                                   Me.txtObsLevantamento.Text, _
                                   pathLogo, _
                                   pathMapaLocal, _
                                   pathRosa, _
                                   pathConvencoes)
End Sub

' ==============================================================================
' BOT�O 2: GERAR PDF DO MAPA (A1)
' Nome sugerido para o bot�o: cmdGerarPDFMapa
' ==============================================================================
Private Sub cmdGerarPDFMapa_Click()
    Dim dadosProp As Object
    
    ' Coleta apenas dados da propriedade (para o nome do arquivo)
    Set dadosProp = ColetarDadosPropriedade()
    
    frmAguarde.Show vbModeless
    frmAguarde.AtualizarStatus "Exportando Mapa A1 para PDF..."
    
    ' Chama a rotina de PDF (False = Mostra mensagem e pergunta se quer abrir)
    Call M_DOC_Mapa.GerarPDFMapa(dadosProp, False)
    
    Unload frmAguarde
End Sub

' ==============================================================================
' FUN��ES AUXILIARES (SELE��O E PREVIEW DE IMAGEM)
' ==============================================================================
Private Function SelecionarImagemDialogo() As String
    Dim fDialog As FileDialog
    Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
    
    With fDialog
        .Title = "Selecione a Imagem"
        .Filters.Clear
        .Filters.Add "Imagens", "*.jpg; *.jpeg; *.png; *.bmp"
        .AllowMultiSelect = False
        
        If .Show = -1 Then
            SelecionarImagemDialogo = .SelectedItems(1)
        Else
            SelecionarImagemDialogo = ""
        End If
    End With
End Function

Private Sub CarregarImagemPreview(imgControl As MSForms.Image, path As String)
    On Error Resume Next
    If path <> "" And Dir(path) <> "" Then
        imgControl.Picture = LoadPicture(path)
        imgControl.PictureSizeMode = fmPictureSizeModeZoom ' Ajusta sem distorcer
    End If
    On Error GoTo 0
End Sub
