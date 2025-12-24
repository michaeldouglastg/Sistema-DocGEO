Attribute VB_Name = "M_Dados"
Option Explicit
' ==============================================================================
' MODULO: M_DADOS
' DESCRICAO: CAMADA DE ACESSO A DADOS COM PROTECAO DE PLANILHAS
' ==============================================================================

Public Sub Dados_LimparTabela(nomePlanilha As String, nomeTabela As String)
    Dim ws As Worksheet, lo As ListObject
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(nomePlanilha)
    If ws Is Nothing Then Exit Sub
    Set lo = ws.ListObjects(nomeTabela)
    If lo Is Nothing Then Exit Sub
    
    M_SheetProtection.DesbloquearPlanilha ws
    If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.Delete
    M_SheetProtection.BloquearPlanilha ws
    On Error GoTo 0
End Sub

Public Sub Dados_LimparTudo()
    Dim wsPainel As Worksheet
    
    If MsgBox("Deseja limpar todos os dados das tabelas SGL e UTM?", _
              vbYesNo + vbQuestion, "Limpar Dados") <> vbYes Then Exit Sub
    
    Call M_Utils.Utils_OtimizarPerformance(True)
    
    Call Dados_LimparTabela(M_Config.SH_SGL, M_Config.TBL_SGL)
    Call Dados_LimparTabela(M_Config.SH_UTM, M_Config.TBL_UTM)
    Call Dados_LimparTabela(M_Config.SH_TEMP_CONV, M_Config.TBL_CONVERSAO)
    
    Call M_Graficos.Grafico_Limpar(M_Config.SH_PAINEL)
    Call M_Graficos.Grafico_Limpar(M_Config.SH_CROQUI)
    
    Set wsPainel = ThisWorkbook.Sheets(M_Config.SH_PAINEL)
    M_SheetProtection.DesbloquearPlanilha wsPainel
    
    On Error Resume Next
    wsPainel.OLEObjects("optSGL").Object.Value = True
    wsPainel.OLEObjects("optUTM").Object.Value = False
    wsPainel.Range(M_Config.CELL_SGL_AREA_HA).Value = ""
    wsPainel.Range(M_Config.CELL_SGL_AREA_M2).Value = ""
    wsPainel.Range(M_Config.CELL_SGL_PERIMETRO).Value = ""
    
    AtualizarShape wsPainel, "shp_Label_Sistema", "AREA TOTAL:"
    AtualizarShape wsPainel, "shp_Valor_Ha", "0,0000 ha"
    AtualizarShape wsPainel, "shp_Valor_M2", "0,00 m2"
    AtualizarShape wsPainel, "shp_Valor_Perimetro", "0,00 m"
    On Error GoTo 0
    
    M_SheetProtection.BloquearPlanilha wsPainel
    Call M_UI_Main.UI_Refresh_ListBox
    Call M_Utils.Utils_OtimizarPerformance(False)
    
    MsgBox "Dados limpos e sistema resetado.", vbInformation
End Sub

Private Sub AtualizarShape(ws As Worksheet, nome As String, texto As String)
    On Error Resume Next
    ws.Shapes(nome).TextFrame2.TextRange.Text = texto
    On Error GoTo 0
End Sub

Public Function Dados_GetArrayTabela(nomePlanilha As String, nomeTabela As String) As Variant
    Dim ws As Worksheet, tbl As ListObject
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(nomePlanilha)
    Set tbl = ws.ListObjects(nomeTabela)
    On Error GoTo 0
    
    If tbl Is Nothing Then Exit Function
    If tbl.ListRows.Count > 0 Then
        Dados_GetArrayTabela = tbl.DataBodyRange.Value
    Else
        Dados_GetArrayTabela = Empty
    End If
End Function

Public Sub Dados_UpsertRegistro(nomePlanilha As String, nomeTabela As String, _
                                 colNomeChave As String, valorChave As String, _
                                 arrColunas As Variant, arrValores As Variant)
    Dim ws As Worksheet, tbl As ListObject
    Dim colIndex As Long, matchRow As Variant
    Dim targetRow As ListRow
    Dim i As Long
    
    If valorChave = "" Then Exit Sub
    
    Set ws = ThisWorkbook.Sheets(nomePlanilha)
    Set tbl = ws.ListObjects(nomeTabela)
    
    M_SheetProtection.DesbloquearPlanilha ws
    
    On Error Resume Next
    colIndex = tbl.ListColumns(colNomeChave).Index
    matchRow = Application.Match(valorChave, tbl.ListColumns(colIndex).DataBodyRange, 0)
    On Error GoTo 0
    
    If Not IsError(matchRow) Then
        If MsgBox("O registro '" & valorChave & "' ja existe. Deseja atualizar?", _
                  vbYesNo + vbQuestion, "Atualizar") = vbYes Then
            Set targetRow = tbl.ListRows(matchRow)
        Else
            M_SheetProtection.BloquearPlanilha ws
            Exit Sub
        End If
    Else
        Set targetRow = tbl.ListRows.Add(AlwaysInsert:=True)
    End If
    
    On Error Resume Next
    For i = LBound(arrColunas) To UBound(arrColunas)
        If arrValores(i) <> "" Then
            targetRow.Range(tbl.ListColumns(arrColunas(i)).Index).Value = arrValores(i)
        End If
    Next i
    On Error GoTo 0
    
    M_SheetProtection.BloquearPlanilha ws
End Sub

Public Function Dados_BuscarValor(nomePlanilha As String, nomeTabela As String, _
                                   valorProcurado As String, _
                                   Optional indiceColunaBusca As Long = 1, _
                                   Optional indiceColunaRetorno As Long = 2) As String
    Dim ws As Worksheet, tbl As ListObject
    Dim rngBusca As Range, foundCell As Range
    Dim linhaTabela As Long
    
    On Error GoTo ErroBusca
    Set ws = ThisWorkbook.Sheets(nomePlanilha)
    Set tbl = ws.ListObjects(nomeTabela)
    Set rngBusca = tbl.ListColumns(indiceColunaBusca).DataBodyRange
    Set foundCell = rngBusca.Find(What:=valorProcurado, LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not foundCell Is Nothing Then
        linhaTabela = foundCell.Row - tbl.HeaderRowRange.Row
        Dados_BuscarValor = CStr(tbl.DataBodyRange(linhaTabela, indiceColunaRetorno).Value)
    Else
        Dados_BuscarValor = ""
    End If
    Exit Function
    
ErroBusca:
    Dados_BuscarValor = ""
End Function

Public Function Dados_LerLinhaParaDict(nomePlanilha As String, nomeTabela As String, _
                                        valorChave As String, colChaveIndex As Long) As Object
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim ws As Worksheet, tbl As ListObject
    Dim rngBusca As Range, foundCell As Range
    Dim i As Long, rowIdx As Long
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(nomePlanilha)
    Set tbl = ws.ListObjects(nomeTabela)
    Set rngBusca = tbl.ListColumns(colChaveIndex).DataBodyRange
    Set foundCell = rngBusca.Find(What:=valorChave, LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not foundCell Is Nothing Then
        rowIdx = foundCell.Row - tbl.HeaderRowRange.Row
        For i = 1 To tbl.ListColumns.Count
            dict.Add tbl.HeaderRowRange(1, i).Value, tbl.DataBodyRange(rowIdx, i).Value
        Next i
    End If
    On Error GoTo 0
    
    Set Dados_LerLinhaParaDict = dict
End Function
