VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EstaPastaDeTrabalho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' ==============================================================================
' ESTAPASTADETRABALHO (ThisWorkbook)
' Cole este codigo no objeto ThisWorkbook
' ==============================================================================
Option Explicit

Private Sub Workbook_Open()
    Dim wsPainel As Worksheet
    Dim wsTemp As Worksheet
    
    On Error Resume Next
    
    ' Inicializa protecao de todas as planilhas
    Call M_SheetProtection.InicializarProtecao
    
    Call M_UI_Main.UI_PopularComboBoxUTM
    
    Set wsPainel = ThisWorkbook.Sheets(M_Config.SH_PAINEL)
    If wsPainel Is Nothing Then Exit Sub
    
    ' Force Refresh
    If ActiveSheet.Name = M_Config.SH_PAINEL Then
        For Each wsTemp In ThisWorkbook.Worksheets
            If wsTemp.Name <> M_Config.SH_PAINEL And wsTemp.Visible = xlSheetVisible Then
                wsTemp.Activate
                Exit For
            End If
        Next wsTemp
    End If
    
    wsPainel.Activate
    On Error GoTo 0
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' Garante que planilhas estejam protegidas ao fechar
    Call M_SheetProtection.BloquearTodas
End Sub

