Attribute VB_Name = "M_SheetProtection"
Option Explicit
' ==============================================================================
' MODULO: M_SHEETPROTECTION
' DESCRICAO: GERENCIA BLOQUEIO E DESBLOQUEIO DE PLANILHAS
' ==============================================================================

Private Const SENHA_PADRAO As String = "12345" ' ALTERE PARA SUA SENHA

' ==============================================================================
' FUNCOES PUBLICAS
' ==============================================================================

Public Sub DesbloquearPlanilha(ws As Worksheet)
'    On Error Resume Next
'    If ws.ProtectContents Then
'        ws.Unprotect Password:=SENHA_PADRAO
'    End If
'    On Error GoTo 0
End Sub

Public Sub BloquearPlanilha(ws As Worksheet, Optional PermitirFiltros As Boolean = True)
'    On Error Resume Next
'    ws.Protect Password:=SENHA_PADRAO, _
'                UserInterfaceOnly:=True, _
'                AllowFiltering:=PermitirFiltros, _
'                AllowSorting:=PermitirFiltros, _
'                AllowFormattingColumns:=True
'    On Error GoTo 0
End Sub

Public Sub DesbloquearTodas()
'    Dim ws As Worksheet
'    For Each ws In ThisWorkbook.Worksheets
'        DesbloquearPlanilha ws
'    Next ws
End Sub

Public Sub BloquearTodas()
'    Dim ws As Worksheet
'    For Each ws In ThisWorkbook.Worksheets
'        BloquearPlanilha ws
'    Next ws
End Sub

Public Sub InicializarProtecao()
    BloquearTodas
End Sub

Public Function EstaProtegida(ws As Worksheet) As Boolean
    EstaProtegida = ws.ProtectContents
End Function
