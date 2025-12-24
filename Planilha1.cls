VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Planilha1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' ==============================================================================
' PLANILHA: PAINEL_PRINCIPAL (Planilha1)
' DESCRICAO: CONTROLA EVENTOS DA INTERFACE PRINCIPAL
' Cole este codigo no objeto da planilha PAINEL_PRINCIPAL
' ==============================================================================
Option Explicit
' Variável para evitar "piscar" a tela desnecessariamente
'Dim UltimoBotaoFocado As String
'Dim UltimoIconeFocado As String

Private Sub Worksheet_Activate()
    ' 1. Ajusta o tamanho da ListBox conforme as celulas N12:AH32
    Call M_UI_Main.UI_Resize_ListBox
    
    ' 2. Garante que um sistema esteja selecionado (Padrao SGL)
    If Not optSGL.Value And Not optUTM.Value Then
        optSGL.Value = True
    End If
    
    ' 3. Atualiza a visualizacao com base na selecao atual
    If optSGL.Value = True Then
        Call M_UI_Main.UI_AtivarVisao_SGL
    Else
        Call M_UI_Main.UI_AtivarVisao_UTM
    End If
End Sub

Private Sub optSGL_Click()
    If optSGL.Value = True Then
        lblInfo.Caption = "As Coordenadas estao em SGL"
        Call M_UI_Main.UI_AtivarVisao_SGL
    End If
End Sub

Private Sub optUTM_Click()
    If optUTM.Value = True Then
        lblInfo.Caption = "As Coordenadas estao em UTM"
        Call M_UI_Main.UI_AtivarVisao_UTM
    End If
End Sub

Private Sub cboFuso_GotFocus()
    cboFuso.DropDown
End Sub

Private Sub cboHemisferio_GotFocus()
    cboHemisferio.DropDown
End Sub

Private Sub lstPrincipal_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim nomeAba As String, nomeTabela As String
    
    If lstPrincipal.ListIndex = -1 Then Exit Sub
    If lstPrincipal.ListIndex = 0 Then Exit Sub
    
    nomeAba = M_Config.App_GetNomeAbaAtiva()
    nomeTabela = M_Config.App_GetNomeTabelaAtiva()
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(nomeAba)
    Set tbl = ws.ListObjects(nomeTabela)
    
    If Not tbl Is Nothing Then
        ws.Activate
        tbl.DataBodyRange.Rows(lstPrincipal.ListIndex).Select
    End If
    On Error GoTo 0
End Sub

Private Sub lstPrincipal_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim i As Long, j As Long
    Dim textoCopiado As String
    Dim dataObj As MSForms.DataObject
    
    ' Ctrl+C (KeyCode 67 com Shift = 2)
    If KeyCode = 67 And Shift = 2 Then
        If lstPrincipal.ListIndex = -1 Then Exit Sub
        
        If lstPrincipal.MultiSelect <> 0 Then
            ' Multiplas Linhas
            For i = 0 To lstPrincipal.ListCount - 1
                If lstPrincipal.Selected(i) Then
                    For j = 0 To lstPrincipal.ColumnCount - 1
                        textoCopiado = textoCopiado & lstPrincipal.List(i, j) & vbTab
                    Next j
                    textoCopiado = Left(textoCopiado, Len(textoCopiado) - 1) & vbCrLf
                End If
            Next i
        Else
            ' Linha Unica
            For j = 0 To lstPrincipal.ColumnCount - 1
                textoCopiado = textoCopiado & lstPrincipal.List(lstPrincipal.ListIndex, j) & vbTab
            Next j
            textoCopiado = Left(textoCopiado, Len(textoCopiado) - 1)
        End If
        
        If Len(textoCopiado) > 0 Then
            Set dataObj = New MSForms.DataObject
            dataObj.SetText textoCopiado
            dataObj.PutInClipboard
            MsgBox "Linha(s) copiada(s) para a area de transferencia!", vbInformation
        End If
    End If
End Sub
