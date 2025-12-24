Attribute VB_Name = "M_UI_Main"
Option Explicit
' ==============================================================================
' MODULO: M_UI_MAIN
' DESCRICAO: CONTROLE DE INTERFACE DO USUARIO (FRONT-END).
' GERENCIA BOTOES DA PLANILHA, LISTBOX, TELA CHEIA E VISUALIZACAO.
' ==============================================================================

' ==============================================================================
' 1. BOTOES DE SISTEMA E IMPORTACAO
' ==============================================================================
Public Sub UI_Show_Importador()
    frmImportar.Show
End Sub

Public Sub UI_Show_Cadastro()
    frmPrincipal.Show
End Sub

Public Sub UI_Show_BD()
    frmGerenciadorDB.Show
End Sub


Public Sub UI_Ver_MAPA()
    ThisWorkbook.Sheets(M_Config.SH_MAPA).Activate
    frmControle.Show
End Sub

Public Sub UI_Ver_BD()
    ThisWorkbook.Sheets(M_Config.SH_BD_PROP).Activate
End Sub

' ALTERNA ENTRE MODO TELA CHEIA E MODO NORMAL
Public Sub UI_ToggleFullscreen()
    Dim EstadoAtual As Boolean
    
    ' Se a Faixa de Opcoes estiver visivel, queremos esconder (Modo Fullscreen)
    EstadoAtual = Application.CommandBars("Ribbon").Visible
    
    Call M_Utils.Utils_OtimizarPerformance(True)
    
    If EstadoAtual = True Then
        ' Entrar em Tela Cheia
        Application.DisplayFormulaBar = False
        Application.DisplayStatusBar = False
        ActiveWindow.DisplayWorkbookTabs = False
        ActiveWindow.DisplayHeadings = False
        ActiveWindow.DisplayGridlines = False
        Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",False)"
    Else
        ' Sair da Tela Cheia
        Application.DisplayFormulaBar = True
        Application.DisplayStatusBar = False
        ActiveWindow.DisplayWorkbookTabs = True
        ActiveWindow.DisplayHeadings = False
        ActiveWindow.DisplayGridlines = False
        Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",True)"
    End If
    
    Call M_Utils.Utils_OtimizarPerformance(False)
End Sub

' ==============================================================================
' 2. CONTROLE DA LISTBOX (PAINEL PRINCIPAL)
' ==============================================================================

' ATUALIZA A LISTBOX COM DADOS DA TABELA ATIVA (COM CABECALHO E FORMATACAO)
Public Sub UI_Refresh_ListBox()
    Dim ws As Worksheet
    Dim objOLE As OLEObject
    Dim lst As MSForms.ListBox
    Dim tbl As ListObject
    Dim arrDados As Variant
    Dim arrFinal() As Variant
    Dim nomeAbaDados As String, nomeTabela As String
    Dim i As Long, C As Long
    Dim numLinhas As Long, numCols As Long
    
    On Error GoTo ErroUI
    
    ' 1. Identifica a fonte de dados e a tabela
    nomeAbaDados = M_Config.App_GetNomeAbaAtiva()
    nomeTabela = M_Config.App_GetNomeTabelaAtiva()
    Set ws = ThisWorkbook.Sheets(nomeAbaDados)
    Set tbl = ws.ListObjects(nomeTabela)
    
    ' 2. Referencia a ListBox na planilha Painel
    Set objOLE = ThisWorkbook.Sheets(M_Config.SH_PAINEL).OLEObjects(M_Config.LB_PRINCIPAL)
    Set lst = objOLE.Object
    
    ' 3. Prepara os Dados em Memoria (Muito mais rapido)
    If tbl.ListRows.Count > 0 Then
        arrDados = tbl.DataBodyRange.Value
        numLinhas = UBound(arrDados, 1)
        numCols = tbl.ListColumns.Count
        
        ' Redimensiona array final: Linhas de Dados + 1 (Cabecalho), Colunas (0 a N-1)
        ReDim arrFinal(0 To numLinhas, 0 To numCols - 1)
        
        ' A. Preenche o Cabecalho (Linha 0)
        For C = 1 To numCols
            arrFinal(0, C - 1) = tbl.HeaderRowRange(1, C).Value
        Next C
        
        ' B. Preenche os Dados (Linha 1 em diante) com Formatacao
        For i = 1 To numLinhas
            For C = 1 To numCols
                ' Verifica se e a coluna 7 (Distancia) para formatar
                If C = 7 Then
                    ' Formata com 2 casas decimais
                    If IsNumeric(arrDados(i, C)) Then
                        arrFinal(i, C - 1) = Format(arrDados(i, C), "0.00")
                    Else
                        arrFinal(i, C - 1) = arrDados(i, C)
                    End If
                Else
                    ' Outras colunas: copia simples
                    arrFinal(i, C - 1) = arrDados(i, C)
                End If
            Next C
        Next i
        
        ' 4. Configura e Preenche a ListBox
        With lst
            .Clear
            .ColumnCount = numCols
            .ColumnHeads = False
            .ColumnWidths = "80;80;80;50;80;50;50;180;30;80"
            .List = arrFinal
        End With
    Else
        ' Se tabela vazia, limpa tudo
        lst.Clear
    End If
    
    Call UI_Resize_ListBox
    Exit Sub
    
ErroUI:
    ' Tratamento de erro silencioso
End Sub

' AJUSTA POSICAO E TAMANHO DA LISTBOX BASEADO EM CELULAS DE REFERENCIA
Public Sub UI_Resize_ListBox()
    Dim ws As Worksheet
    Dim lst As OLEObject
    Dim rngInicio As Range, rngFim As Range
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(M_Config.SH_PAINEL)
    Set lst = ws.OLEObjects(M_Config.LB_PRINCIPAL)
    
    ' Define area (ex: N12 ate AH32)
    Set rngInicio = ws.Range("C13")
    Set rngFim = ws.Range("AH33")
    
    If Not lst Is Nothing Then
        With lst
            .Top = rngInicio.Top
            .Left = rngInicio.Left
            .Width = (rngFim.Left + rngFim.Width) - rngInicio.Left
            .Height = (rngFim.Top + rngFim.Height) - rngInicio.Top
        End With
    End If
End Sub

' ==============================================================================
' 3. BOTOES DE NAVEGACAO DE DADOS (SGL / UTM)
' ==============================================================================
Public Sub UI_AtivarVisao_SGL()
    ' A logica de "qual esta ativo" e lida dos OptionButtons pelo M_Config.
    ' Esta macro garante que os calculos e a lista sejam atualizados.
    Call M_App_Logica.Processo_AtualizarMetricas
    Call UI_Refresh_ListBox
    Call M_Graficos.Grafico_PlotarPoligono(M_Config.SH_PAINEL)
End Sub

Public Sub UI_AtivarVisao_UTM()
    ' Ao mudar para UTM, garantimos que a conversao esteja atualizada
    'Call M_Utils.Utils_OtimizarPerformance(True)
    'Call M_App_Logica.Processo_Conv_SGL_UTM
    Call M_App_Logica.Processo_AtualizarMetricas
    Call UI_Refresh_ListBox
    Call M_Graficos.Grafico_PlotarPoligono(M_Config.SH_PAINEL)
    'Call M_Utils.Utils_OtimizarPerformance(False)
End Sub

Public Sub UI_Ver_Grafico_Croqui()
    ThisWorkbook.Sheets(M_Config.SH_CROQUI).Activate
    Call M_Graficos.Grafico_PlotarPoligono(M_Config.SH_CROQUI)
End Sub

Public Sub UI_Ver_Planilha_UTM()
    Call M_App_Logica.Processo_Conv_SGL_UTM
    ThisWorkbook.Sheets(M_Config.SH_UTM).Activate
End Sub

Public Sub UI_Ver_Planilha_SGL()
    ThisWorkbook.Sheets(M_Config.SH_SGL).Activate
End Sub

Public Sub UI_Ver_Panel_Principal()
    ThisWorkbook.Sheets(M_Config.SH_PAINEL).Activate
End Sub

' ==============================================================================
' 4. BOTOES DE CONTROLE DE GRAFICO (WRAPPERS)
' ==============================================================================
' PAINEL
Public Sub UI_Btn_ZoomIn()
    Call M_Graficos.Grafico_Zoom(M_Config.SH_CROQUI, 1.2)
End Sub

Public Sub UI_Btn_ZoomOut()
    Call M_Graficos.Grafico_Zoom(M_Config.SH_CROQUI, 0.8)
End Sub

Public Sub UI_Btn_PanUp()
    Call M_Graficos.Grafico_Pan(M_Config.SH_CROQUI, "CIMA")
End Sub

Public Sub UI_Btn_PanDown()
    Call M_Graficos.Grafico_Pan(M_Config.SH_CROQUI, "BAIXO")
End Sub

Public Sub UI_Btn_PanLeft()
    Call M_Graficos.Grafico_Pan(M_Config.SH_CROQUI, "ESQUERDA")
End Sub

Public Sub UI_Btn_PanRight()
    Call M_Graficos.Grafico_Pan(M_Config.SH_CROQUI, "DIREITA")
End Sub

Public Sub UI_Btn_Reset()
    Call M_Graficos.Grafico_PlotarPoligono(M_Config.SH_PAINEL)
End Sub

' ABA CROQUI (Botoes duplicados na outra aba chamam estas)
Public Sub UI_Btn_Croqui_ZoomIn()
    Call M_Graficos.Grafico_Zoom(M_Config.SH_CROQUI, 1.2)
End Sub

Public Sub UI_Btn_Croqui_ZoomOut()
    Call M_Graficos.Grafico_Zoom(M_Config.SH_CROQUI, 0.8)
End Sub

Public Sub UI_Btn_Croqui_Reset()
    Call M_Graficos.Grafico_PlotarPoligono(M_Config.SH_CROQUI)
End Sub

' ==============================================================================
' BOTOES DE ROTULOS (MOSTRAR/OCULTAR NOMES DOS PONTOS)
' ==============================================================================
' Botao para PAINEL_PRINCIPAL
Public Sub UI_Btn_Rotulos()
    Call M_Graficos.Grafico_AlternarRotulos(M_Config.SH_PAINEL)
End Sub

' Botao para CROQUI
Public Sub UI_Btn_Croqui_Rotulos()
    Call M_Graficos.Grafico_AlternarRotulos(M_Config.SH_CROQUI)
End Sub

' ==============================================================================
' 5. BOTOES DE EXPORTACAO (ACIONADOS PELA PLANILHA)
' ==============================================================================
Public Sub UI_Btn_GerarDXF()
    Dim dadosProp As Object
    
    ' 1. Coleta nome para o arquivo
    Set dadosProp = UI_ColetarDadosMinimos("Gerar DXF")
    If dadosProp Is Nothing Then Exit Sub
    
    ' 2. Chama o exportador
    Call M_DOC_Exportacao.ExportarDXF(dadosProp)
End Sub

Public Sub UI_Btn_GerarKML()
    Dim dadosProp As Object
    
    ' 1. Coleta nome para o arquivo
    Set dadosProp = UI_ColetarDadosMinimos("Gerar KML (Google Earth)")
    If dadosProp Is Nothing Then Exit Sub
    
    ' 2. Chama o exportador
    Call M_DOC_Exportacao.ExportarKML(dadosProp)
End Sub

' --- FUNCAO AUXILIAR PRIVADA DO MODULO UI ---
' Cria um dicionario temporario apenas com o Nome da Propriedade
Private Function UI_ColetarDadosMinimos(titulo As String) As Object
    Dim nomeProp As String
    Dim dict As Object
    
    nomeProp = InputBox("Informe o nome da Propriedade/Projeto para o arquivo:", titulo, "Projeto_Sem_Titulo")
    
    If Trim(nomeProp) = "" Then
        Set UI_ColetarDadosMinimos = Nothing
        Exit Function
    End If
    
    Set dict = CreateObject("Scripting.Dictionary")
    dict.Add M_Config.LBL_PROPRIEDADE, nomeProp
    
    Set UI_ColetarDadosMinimos = dict
End Function

' ==============================================================================
' POPULAR COMBOBOX DE FUSO E HEMISFERIO
' ==============================================================================
Public Sub UI_PopularComboBoxUTM()
    Dim wsPainel As Worksheet
    Dim cboFuso As MSForms.ComboBox
    Dim cboHemisferio As MSForms.ComboBox
    Dim i As Long
    
    On Error Resume Next
    Set wsPainel = ThisWorkbook.Sheets(M_Config.SH_PAINEL)
    Set cboFuso = wsPainel.OLEObjects(M_Config.CBO_FUSO).Object
    Set cboHemisferio = wsPainel.OLEObjects(M_Config.CBO_HEMISFERIO).Object
    On Error GoTo 0
    
    If cboFuso Is Nothing Or cboHemisferio Is Nothing Then Exit Sub
    
    ' Popular Fusos (18 a 25 - Brasil)
    cboFuso.Clear
    For i = 18 To 25
        cboFuso.AddItem i
    Next i
    
    ' Popular Hemisferios
    cboHemisferio.Clear
    cboHemisferio.AddItem "Sul"
    cboHemisferio.AddItem "Norte"
    
    ' Valores padrao (podem ser alterados pela deteccao automatica)
    If cboFuso.ListIndex = -1 Then cboFuso.ListIndex = 4  ' Fuso 22
    If cboHemisferio.ListIndex = -1 Then cboHemisferio.ListIndex = 0  ' Sul
End Sub

Public Sub UI_DetectarFusoHemisferio()
    Dim wsPainel As Worksheet
    Dim wsSGL As Worksheet
    Dim loSGL As ListObject
    Dim cboFuso As MSForms.ComboBox
    Dim cboHemisferio As MSForms.ComboBox
    Dim latDD As Double, lonDD As Double
    Dim fusoDetectado As Long
    Dim i As Long
    
    On Error Resume Next
    Set wsPainel = ThisWorkbook.Sheets(M_Config.SH_PAINEL)
    Set wsSGL = ThisWorkbook.Sheets(M_Config.SH_SGL)
    Set loSGL = wsSGL.ListObjects(M_Config.TBL_SGL)
    Set cboFuso = wsPainel.OLEObjects(M_Config.CBO_FUSO).Object
    Set cboHemisferio = wsPainel.OLEObjects(M_Config.CBO_HEMISFERIO).Object
    On Error GoTo 0
    
    If loSGL Is Nothing Or cboFuso Is Nothing Or cboHemisferio Is Nothing Then Exit Sub
    If loSGL.ListRows.Count = 0 Then Exit Sub
    
    ' Pega primeira coordenada (Longitude coluna 2, Latitude coluna 3)
    lonDD = M_Utils.Str_DMS_Para_DD(CStr(loSGL.ListRows(1).Range(2).Value))
    latDD = M_Utils.Str_DMS_Para_DD(CStr(loSGL.ListRows(1).Range(3).Value))
    
    ' Detecta Fuso
    fusoDetectado = M_Math_Geo.Geo_GetZonaUTM(lonDD)
    For i = 0 To cboFuso.ListCount - 1
        If CInt(cboFuso.List(i)) = fusoDetectado Then
            cboFuso.ListIndex = i
            Exit For
        End If
    Next i
    
    ' Detecta Hemisferio
    If latDD < 0 Then
        cboHemisferio.ListIndex = 0  ' Sul
    Else
        cboHemisferio.ListIndex = 1  ' Norte
    End If
End Sub

Public Function UI_GetFusoSelecionado() As Long
    Dim wsPainel As Worksheet
    Dim cboFuso As MSForms.ComboBox
    
    On Error Resume Next
    Set wsPainel = ThisWorkbook.Sheets(M_Config.SH_PAINEL)
    Set cboFuso = wsPainel.OLEObjects(M_Config.CBO_FUSO).Object
    
    If cboFuso Is Nothing Or cboFuso.ListIndex = -1 Then
        UI_GetFusoSelecionado = 0  ' Automatico
    Else
        UI_GetFusoSelecionado = CLng(cboFuso.Value)
    End If
    On Error GoTo 0
End Function

Public Function UI_GetHemisferioSul() As Boolean
    Dim wsPainel As Worksheet
    Dim cboHemisferio As MSForms.ComboBox
    
    On Error Resume Next
    Set wsPainel = ThisWorkbook.Sheets(M_Config.SH_PAINEL)
    Set cboHemisferio = wsPainel.OLEObjects(M_Config.CBO_HEMISFERIO).Object
    
    If cboHemisferio Is Nothing Or cboHemisferio.ListIndex = -1 Then
        UI_GetHemisferioSul = True  ' Padrao Sul
    Else
        UI_GetHemisferioSul = (cboHemisferio.Value = "Sul")
    End If
    On Error GoTo 0
End Function

Public Sub UI_Btn_VisualizarMapa()
    'Call M_DOC_Mapa.UI_VisualizarMapa
    'ThisWorkbook.Sheets(M_Config.SH_MAPA).Activate
    'frmControle.Show
End Sub
