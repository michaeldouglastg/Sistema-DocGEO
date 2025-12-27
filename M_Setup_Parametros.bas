Attribute VB_Name = "M_Setup_Parametros"
Option Explicit
' ==============================================================================
' MODULO: M_SETUP_PARAMETROS
' DESCRICAO: INICIALIZACAO E ATUALIZACAO DE PARAMETROS INCRA
' ==============================================================================

Public Sub Setup_PopularParametrosINCRA()
    '----------------------------------------------------------------------------------
    ' Popula a tabela de parametros com os codigos oficiais INCRA
    ' Executa apenas se a tabela estiver vazia ou mediante confirmacao do usuario
    '----------------------------------------------------------------------------------
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim i As Long
    Dim dados() As Variant
    Dim resposta As VbMsgBoxResult

    On Error GoTo ErroSetup

    ' Verifica se tabela de parametros existe
    Set ws = ThisWorkbook.Sheets(M_Config.SH_PARAMETROS)
    Set tbl = ws.ListObjects(M_Config.TBL_PARAMETROS)

    ' Se ja houver dados, pergunta ao usuario
    If tbl.ListRows.Count > 0 Then
        resposta = MsgBox("A tabela de parametros ja contem dados." & vbCrLf & _
                          "Deseja adicionar/atualizar os codigos oficiais INCRA?", _
                          vbYesNo + vbQuestion, "Atualizar Parametros INCRA")
        If resposta <> vbYes Then Exit Sub
    End If

    Call M_Utils.Utils_OtimizarPerformance(True)
    M_SheetProtection.DesbloquearPlanilha ws

    ' Dados dos parametros INCRA
    ' Formato: Array(Codigo, Descricao, Tipo, Precisao)
    dados = Array( _
        Array("LA1", "Cerca", "Artificial", "0.50m"), _
        Array("LA2", "Estrada", "Artificial", "0.50m"), _
        Array("LA3", "Rio/Corrego Canalizado", "Artificial", "0.50m"), _
        Array("LA4", "Vala, Rego, Canal", "Artificial", "0.50m"), _
        Array("LA5", "Limite Inacessivel (Artificial)", "Artificial", "7.50m"), _
        Array("LA6", "Limite Inacessivel (Serra, Escarpa)", "Artificial", "7.50m"), _
        Array("LA7", "Limite Inacessivel (Rio, Corrego, Lago)", "Artificial", "7.50m"), _
        Array("LN1", "Talvegue de Rio/Corrego", "Natural", "3.00m"), _
        Array("LN2", "Crista de Serra/Espigao", "Natural", "3.00m"), _
        Array("LN3", "Margem de Rio/Corrego", "Natural", "3.00m"), _
        Array("LN4", "Margem de Lago/Lagoa", "Natural", "3.00m"), _
        Array("LN5", "Margem de Oceano", "Natural", "3.00m"), _
        Array("LN6", "Limite Seco de Praia/Mangue", "Natural", "3.00m"), _
        Array("M", "Marco (materializado)", "Vertice", "-"), _
        Array("P", "Ponto (feicao identificavel)", "Vertice", "-"), _
        Array("V", "Virtual (calculado)", "Vertice", "-"), _
        Array("GNSS-RTK", "GNSS - Real Time Kinematic", "Metodo", "-"), _
        Array("GNSS-PPP", "GNSS - Precise Point Positioning", "Metodo", "-"), _
        Array("GNSS-REL", "GNSS - Relativo", "Metodo", "-"), _
        Array("TOP", "Topografia Classica", "Metodo", "-"), _
        Array("GAN", "Geometria Analitica", "Metodo", "-"), _
        Array("SRE", "Sensoriamento Remoto", "Metodo", "-"), _
        Array("BCA", "Base Cartografica", "Metodo", "-") _
    )

    ' Adiciona os dados linha por linha
    For i = LBound(dados) To UBound(dados)
        Call AdicionarOuAtualizar_Parametro(tbl, dados(i)(0), dados(i)(1), dados(i)(2), dados(i)(3))
    Next i

    M_SheetProtection.BloquearPlanilha ws
    Call M_Utils.Utils_OtimizarPerformance(False)

    MsgBox "Parametros INCRA atualizados com sucesso!" & vbCrLf & _
           "Total de registros: " & (UBound(dados) - LBound(dados) + 1), _
           vbInformation, "Setup Concluido"
    Exit Sub

ErroSetup:
    Call M_Utils.Utils_OtimizarPerformance(False)
    On Error Resume Next
    M_SheetProtection.BloquearPlanilha ws
    On Error GoTo 0
    MsgBox "Erro ao popular parametros INCRA: " & Err.Description, vbCritical
End Sub

Private Sub AdicionarOuAtualizar_Parametro(tbl As ListObject, codigo As String, _
                                            descricao As String, tipo As String, precisao As String)
    '----------------------------------------------------------------------------------
    ' Adiciona ou atualiza um registro na tabela de parametros
    ' Se o codigo ja existir, atualiza; senao, adiciona novo
    '----------------------------------------------------------------------------------
    Dim rng As Range, found As Range
    Dim novaLinha As ListRow

    On Error Resume Next
    Set rng = tbl.ListColumns(1).DataBodyRange
    Set found = rng.Find(What:=codigo, LookIn:=xlValues, LookAt:=xlWhole)
    On Error GoTo 0

    If Not found Is Nothing Then
        ' Atualiza registro existente
        Dim linhaIdx As Long
        linhaIdx = found.Row - tbl.HeaderRowRange.Row
        tbl.DataBodyRange(linhaIdx, 1).Value = codigo
        tbl.DataBodyRange(linhaIdx, 2).Value = descricao
        If tbl.ListColumns.Count >= 3 Then tbl.DataBodyRange(linhaIdx, 3).Value = tipo
        If tbl.ListColumns.Count >= 4 Then tbl.DataBodyRange(linhaIdx, 4).Value = precisao
    Else
        ' Adiciona novo registro
        Set novaLinha = tbl.ListRows.Add
        novaLinha.Range(1, 1).Value = codigo
        novaLinha.Range(1, 2).Value = descricao
        If tbl.ListColumns.Count >= 3 Then novaLinha.Range(1, 3).Value = tipo
        If tbl.ListColumns.Count >= 4 Then novaLinha.Range(1, 4).Value = precisao
    End If
End Sub

Public Sub Setup_VerificarEstruturaDados()
    '----------------------------------------------------------------------------------
    ' Verifica se as tabelas principais possuem as colunas necessarias
    ' para as validacoes INCRA (Precisao H, Precisao V, Metodo, Cod. Limite)
    '----------------------------------------------------------------------------------
    Dim wsSGL As Worksheet, wsUTM As Worksheet
    Dim tblSGL As ListObject, tblUTM As ListObject
    Dim relatorio As String

    On Error Resume Next
    Set wsSGL = ThisWorkbook.Sheets(M_Config.SH_SGL)
    Set wsUTM = ThisWorkbook.Sheets(M_Config.SH_UTM)
    Set tblSGL = wsSGL.ListObjects(M_Config.TBL_SGL)
    Set tblUTM = wsUTM.ListObjects(M_Config.TBL_UTM)
    On Error GoTo 0

    relatorio = "VERIFICACAO DE ESTRUTURA DE DADOS" & vbCrLf
    relatorio = relatorio & String(50, "=") & vbCrLf & vbCrLf

    ' Verifica tabela SGL
    relatorio = relatorio & "TABELA SGL:" & vbCrLf
    relatorio = relatorio & "  Total de colunas: " & tblSGL.ListColumns.Count & vbCrLf
    relatorio = relatorio & "  Colunas necessarias para INCRA:" & vbCrLf
    relatorio = relatorio & "    - Tipo (vertice): " & IIf(ExisteColuna(tblSGL, "Tipo"), "OK", "FALTA") & vbCrLf
    relatorio = relatorio & "    - Descricao (limite): " & IIf(ExisteColuna(tblSGL, "Descricao"), "OK", "FALTA") & vbCrLf
    relatorio = relatorio & "    - Precisao H: " & IIf(ExisteColuna(tblSGL, "Precisao H"), "OK", "FALTA (ADICIONAR)") & vbCrLf
    relatorio = relatorio & "    - Precisao V: " & IIf(ExisteColuna(tblSGL, "Precisao V"), "OK", "FALTA (ADICIONAR)") & vbCrLf
    relatorio = relatorio & "    - Metodo Posic.: " & IIf(ExisteColuna(tblSGL, "Metodo"), "OK", "FALTA (ADICIONAR)") & vbCrLf
    relatorio = relatorio & "    - Cod. Limite: " & IIf(ExisteColuna(tblSGL, "Cod"), "OK", "FALTA (ADICIONAR)") & vbCrLf
    relatorio = relatorio & vbCrLf

    ' Verifica tabela UTM
    relatorio = relatorio & "TABELA UTM:" & vbCrLf
    relatorio = relatorio & "  Total de colunas: " & tblUTM.ListColumns.Count & vbCrLf
    relatorio = relatorio & "  Colunas necessarias para INCRA:" & vbCrLf
    relatorio = relatorio & "    - Tipo (vertice): " & IIf(ExisteColuna(tblUTM, "Tipo"), "OK", "FALTA") & vbCrLf
    relatorio = relatorio & "    - Descricao (limite): " & IIf(ExisteColuna(tblUTM, "Descricao"), "OK", "FALTA") & vbCrLf
    relatorio = relatorio & "    - Precisao H: " & IIf(ExisteColuna(tblUTM, "Precisao H"), "OK", "FALTA (ADICIONAR)") & vbCrLf
    relatorio = relatorio & "    - Precisao V: " & IIf(ExisteColuna(tblUTM, "Precisao V"), "OK", "FALTA (ADICIONAR)") & vbCrLf
    relatorio = relatorio & "    - Metodo Posic.: " & IIf(ExisteColuna(tblUTM, "Metodo"), "OK", "FALTA (ADICIONAR)") & vbCrLf
    relatorio = relatorio & "    - Cod. Limite: " & IIf(ExisteColuna(tblUTM, "Cod"), "OK", "FALTA (ADICIONAR)") & vbCrLf

    MsgBox relatorio, vbInformation, "Verificacao de Estrutura"
End Sub

Private Function ExisteColuna(tbl As ListObject, nomeColuna As String) As Boolean
    '----------------------------------------------------------------------------------
    ' Verifica se uma coluna existe na tabela
    '----------------------------------------------------------------------------------
    Dim col As ListColumn

    On Error Resume Next
    Set col = tbl.ListColumns(nomeColuna)
    ExisteColuna = Not col Is Nothing
    On Error GoTo 0
End Function

Public Sub Setup_AdicionarColunasValidacao()
    '----------------------------------------------------------------------------------
    ' Adiciona as colunas de validacao INCRA nas tabelas SGL e UTM
    ' ATENCAO: Esta funcao modifica a estrutura das tabelas
    '----------------------------------------------------------------------------------
    Dim wsSGL As Worksheet, wsUTM As Worksheet
    Dim tblSGL As ListObject, tblUTM As ListObject
    Dim resposta As VbMsgBoxResult

    resposta = MsgBox("Esta operacao adicionara colunas de validacao INCRA nas tabelas SGL e UTM:" & vbCrLf & _
                      "  - Precisao H (m)" & vbCrLf & _
                      "  - Precisao V (m)" & vbCrLf & _
                      "  - Metodo Posic." & vbCrLf & _
                      "  - Cod. Limite" & vbCrLf & vbCrLf & _
                      "Deseja continuar?", _
                      vbYesNo + vbQuestion, "Adicionar Colunas")

    If resposta <> vbYes Then Exit Sub

    On Error GoTo ErroAdicionar

    Set wsSGL = ThisWorkbook.Sheets(M_Config.SH_SGL)
    Set wsUTM = ThisWorkbook.Sheets(M_Config.SH_UTM)
    Set tblSGL = wsSGL.ListObjects(M_Config.TBL_SGL)
    Set tblUTM = wsUTM.ListObjects(M_Config.TBL_UTM)

    Call M_Utils.Utils_OtimizarPerformance(True)

    ' Adiciona colunas na tabela SGL
    M_SheetProtection.DesbloquearPlanilha wsSGL
    Call AdicionarColunaSeFaltando(tblSGL, "Precisao H (m)", "0.00")
    Call AdicionarColunaSeFaltando(tblSGL, "Precisao V (m)", "0.00")
    Call AdicionarColunaSeFaltando(tblSGL, "Metodo Posic.", "GNSS-RTK")
    Call AdicionarColunaSeFaltando(tblSGL, "Cod. Limite", "LA1")
    M_SheetProtection.BloquearPlanilha wsSGL

    ' Adiciona colunas na tabela UTM
    M_SheetProtection.DesbloquearPlanilha wsUTM
    Call AdicionarColunaSeFaltando(tblUTM, "Precisao H (m)", "0.00")
    Call AdicionarColunaSeFaltando(tblUTM, "Precisao V (m)", "0.00")
    Call AdicionarColunaSeFaltando(tblUTM, "Metodo Posic.", "GNSS-RTK")
    Call AdicionarColunaSeFaltando(tblUTM, "Cod. Limite", "LA1")
    M_SheetProtection.BloquearPlanilha wsUTM

    Call M_Utils.Utils_OtimizarPerformance(False)

    MsgBox "Colunas de validacao INCRA adicionadas com sucesso!", vbInformation
    Exit Sub

ErroAdicionar:
    Call M_Utils.Utils_OtimizarPerformance(False)
    MsgBox "Erro ao adicionar colunas: " & Err.Description, vbCritical
End Sub

Private Sub AdicionarColunaSeFaltando(tbl As ListObject, nomeColuna As String, valorPadrao As String)
    '----------------------------------------------------------------------------------
    ' Adiciona uma coluna na tabela se ela nao existir
    ' Preenche com valor padrao para linhas existentes
    '----------------------------------------------------------------------------------
    Dim col As ListColumn
    Dim novaCol As ListColumn

    On Error Resume Next
    Set col = tbl.ListColumns(nomeColuna)
    On Error GoTo 0

    If col Is Nothing Then
        ' Adiciona nova coluna
        Set novaCol = tbl.ListColumns.Add
        novaCol.Name = nomeColuna

        ' Preenche com valor padrao
        If tbl.ListRows.Count > 0 And valorPadrao <> "" Then
            novaCol.DataBodyRange.Value = valorPadrao
        End If
    End If
End Sub
