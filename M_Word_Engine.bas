Attribute VB_Name = "M_Word_Engine"
Option Explicit
' ==============================================================================
' MODULO: M_WORD_ENGINE
' DESCRICAO: MOTOR DE AUTOMACAO DO WORD COM SETUP/TEARDOWN CENTRALIZADOS
' ==============================================================================

' --- CONSTANTES DO WORD ---
Public Const wdAlignParagraphLeft As Integer = 0
Public Const wdAlignParagraphCenter As Integer = 1
Public Const wdAlignParagraphRight As Integer = 2
Public Const wdAlignParagraphJustify As Integer = 3
Public Const wdLineSpaceSingle As Integer = 0
Public Const wdUnderlineNone As Integer = 0
Public Const wdUnderlineSingle As Integer = 1
Public Const wdCollapseEnd As Integer = 0
Public Const wdColorGray15 As Long = 14277081
Public Const wdExportFormatPDF As Integer = 17
Public Const wdDoNotSaveChanges As Integer = 0
Public Const wdCellAlignVerticalCenter As Integer = 1

' --- VARIAVEIS GLOBAIS DO MOTOR ---
Private g_WordApp As Object
Private g_WordDoc As Object

' ==============================================================================
' 1. SETUP - INICIALIZACAO DO WORD
' ==============================================================================
Public Function Word_Setup(Optional Visivel As Boolean = False, _
                           Optional TopCm As Double = 1.27, _
                           Optional BottomCm As Double = 1.27, _
                           Optional LeftCm As Double = 1.27, _
                           Optional RightCm As Double = 1.27) As Boolean
    On Error GoTo ErroSetup
    
    On Error Resume Next
    Set g_WordApp = GetObject(, "Word.Application")
    On Error GoTo ErroSetup
    
    If g_WordApp Is Nothing Then
        Set g_WordApp = CreateObject("Word.Application")
    End If
    
    If g_WordApp Is Nothing Then
        MsgBox "Nao foi possivel iniciar o Microsoft Word.", vbCritical
        Word_Setup = False
        Exit Function
    End If
    
    g_WordApp.Visible = Visivel
    If Not Visivel Then g_WordApp.ScreenUpdating = False
    
    Set g_WordDoc = g_WordApp.Documents.Add
    
    With g_WordDoc.PageSetup
        .TopMargin = Application.CentimetersToPoints(TopCm)
        .BottomMargin = Application.CentimetersToPoints(BottomCm)
        .LeftMargin = Application.CentimetersToPoints(LeftCm)
        .RightMargin = Application.CentimetersToPoints(RightCm)
    End With
    
    With g_WordDoc.Content
        .Font.Name = "Arial"
        .Font.Size = 12
        .ParagraphFormat.LineSpacingRule = wdLineSpaceSingle
        .ParagraphFormat.SpaceBefore = 0
        .ParagraphFormat.SpaceAfter = 0
    End With
    
    Word_Setup = True
    Exit Function
    
ErroSetup:
    MsgBox "Erro ao inicializar Word: " & Err.Description, vbCritical
    Word_Setup = False
End Function

' ==============================================================================
' 2. TEARDOWN - SALVAMENTO E FECHAMENTO
' ==============================================================================
Public Function Word_Teardown(nomeArquivo As String, _
                              Optional ComoPDF As Boolean = False, _
                              Optional PerguntarLocal As Boolean = True, _
                              Optional pastaDestino As String = "", _
                              Optional MostrarAoFinal As Boolean = True) As String
    On Error GoTo ErroTeardown
    
    Dim caminhoFinal As String
    Dim ext As String, filtro As String, titulo As String
    Dim arquivoDestino As Variant
    
    g_WordApp.ScreenUpdating = True
    
    If ComoPDF Then
        ext = ".pdf"
        filtro = "Arquivo PDF (*.pdf), *.pdf"
        titulo = "Salvar como PDF"
    Else
        ext = ".docx"
        filtro = "Documento Word (*.docx), *.docx"
        titulo = "Salvar como Word"
    End If
    
    If pastaDestino <> "" Then
        caminhoFinal = pastaDestino & "\" & nomeArquivo & ext
    ElseIf PerguntarLocal Then
        arquivoDestino = Application.GetSaveAsFilename( _
            InitialFileName:=nomeArquivo & ext, _
            FileFilter:=filtro, _
            Title:=titulo)
        
        If arquivoDestino = False Then
            g_WordDoc.Close SaveChanges:=wdDoNotSaveChanges
            Word_Limpar
            Word_Teardown = ""
            Exit Function
        End If
        caminhoFinal = CStr(arquivoDestino)
    Else
        caminhoFinal = Environ("USERPROFILE") & "\Documents\" & nomeArquivo & ext
    End If
    
    If ComoPDF Then
        g_WordDoc.ExportAsFixedFormat OutputFileName:=caminhoFinal, ExportFormat:=wdExportFormatPDF
        g_WordDoc.Close SaveChanges:=wdDoNotSaveChanges
    Else
        g_WordDoc.SaveAs2 filename:=caminhoFinal
        If Not MostrarAoFinal Then
            g_WordDoc.Close
        End If
    End If
    
    If MostrarAoFinal And Not ComoPDF Then
        g_WordApp.Visible = True
    Else
        If g_WordApp.Documents.Count = 0 Then
            g_WordApp.Quit
        End If
    End If
    
    Word_Teardown = caminhoFinal
    Word_Limpar
    Exit Function
    
ErroTeardown:
    MsgBox "Erro ao salvar documento: " & Err.Description, vbCritical
    On Error Resume Next
    If Not g_WordApp Is Nothing Then g_WordApp.Visible = True
    Word_Limpar
    Word_Teardown = ""
End Function

Private Sub Word_Limpar()
    Set g_WordDoc = Nothing
    Set g_WordApp = Nothing
End Sub

' ==============================================================================
' 3. GETTERS PARA USO EXTERNO
' ==============================================================================
Public Function GetWordApp() As Object
    Set GetWordApp = g_WordApp
End Function

Public Function GetWordDoc() As Object
    Set GetWordDoc = g_WordDoc
End Function

Public Function GetSelection() As Object
    If Not g_WordApp Is Nothing Then
        Set GetSelection = g_WordApp.Selection
    End If
End Function

' ==============================================================================
' 4. FUNCOES DE ESCRITA
' ==============================================================================
Public Sub Word_EscreverTitulo(texto As String, Optional Tamanho As Integer = 14)
    With g_WordApp.Selection
        .ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Font.Bold = True
        .Font.Underline = wdUnderlineSingle
        .Font.Size = Tamanho
        .TypeText Text:=texto
        .TypeParagraph
        .TypeParagraph
        .Font.Bold = False
        .Font.Underline = wdUnderlineNone
        .Font.Size = 12
        .ParagraphFormat.Alignment = wdAlignParagraphJustify
    End With
End Sub

Public Sub Word_EscreverParagrafo(texto As String, _
                                   Optional Negrito As Boolean = False, _
                                   Optional Alinhamento As Integer = 3)
    With g_WordApp.Selection
        .ParagraphFormat.Alignment = Alinhamento
        .Font.Bold = Negrito
        .TypeText Text:=texto
        .TypeParagraph
    End With
End Sub

Public Sub Word_PularLinha(Optional qtd As Integer = 1)
    Dim i As Integer
    For i = 1 To qtd
        g_WordApp.Selection.TypeParagraph
    Next i
End Sub

Public Sub Word_EscreverTexto(texto As String, Optional Negrito As Boolean = False)
    With g_WordApp.Selection
        .Font.Bold = Negrito
        .TypeText Text:=texto
    End With
End Sub

' ==============================================================================
' 5. FUNCOES DE TABELA
' ==============================================================================
Public Function Word_CriarTabela(numLinhas As Long, numColunas As Long, _
                                  Optional CabecalhoCinza As Boolean = True) As Object
    Dim rangeTabela As Object
    Dim tabela As Object
    
    Set rangeTabela = g_WordDoc.Content
    rangeTabela.Collapse Direction:=wdCollapseEnd
    
    Set tabela = g_WordDoc.Tables.Add(Range:=g_WordApp.Selection.Range, _
                                       NumRows:=numLinhas, NumColumns:=numColunas)
    
    With tabela
        .Borders.Enable = True
        .Range.Font.Name = "Arial"
        .Range.Font.Size = 9
        .Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Range.Cells.VerticalAlignment = wdCellAlignVerticalCenter
        
        If CabecalhoCinza Then
            With .Rows(1).Range
                .Font.Bold = True
                .Shading.BackgroundPatternColor = wdColorGray15
            End With
        End If
    End With
    
    Set Word_CriarTabela = tabela
End Function

Public Sub Word_MoverAposTabela()
    Dim rng As Object
    Set rng = g_WordDoc.Content
    rng.Collapse Direction:=wdCollapseEnd
    rng.Select
    g_WordApp.Selection.TypeParagraph
End Sub

' ==============================================================================
' 6. FUNCOES DE FORMATACAO
' ==============================================================================
Public Sub Word_FindAndBold(textoProcurado As String)
    If Trim(textoProcurado) = "" Or InStr(1, textoProcurado, "nao encontrado") > 0 Then Exit Sub
    
    Dim rng As Object
    Set rng = g_WordDoc.Content
    
    With rng.Find
        .ClearFormatting
        .Text = textoProcurado
        .Replacement.Text = textoProcurado
        .Replacement.Font.Bold = True
        .Execute Replace:=2
    End With
End Sub

Public Sub Word_SetCellBoldLabel(cell As Object, label As String, Value As String)
    Dim cellRng As Object
    Set cellRng = cell.Range
    
    cellRng.Text = label & Value
    cellRng.End = cellRng.End - 1
    
    With cellRng.Find
        .Text = label
        .MatchCase = False
        If .Execute Then
            cellRng.Font.Bold = True
        End If
    End With
End Sub

Public Sub Word_SetCellTextBoldLabel(cell As Object, label As String, Value As String)
    Dim cellRng As Object
    Set cellRng = cell.Range
    
    cellRng.Text = label & Value
    cellRng.End = cellRng.End - 1
    
    With cellRng.Find
        .Text = label
        .MatchCase = False
        If .Execute Then
            cellRng.Font.Bold = True
        End If
    End With
End Sub

Public Sub SetCellTextBoldLabel(cell As Object, label As String, Value As String)
    Call Word_SetCellTextBoldLabel(cell, label, Value)
End Sub

' ==============================================================================
' 7. FORMATACAO DE DATA
' ==============================================================================
Public Function Word_FormatarData() As String
    Dim dataTexto As String
    dataTexto = Format(Date, "dd") & " de " & Format(Date, "mmmm") & " de " & Format(Date, "yyyy")
    dataTexto = StrConv(dataTexto, vbProperCase)
    Word_FormatarData = Replace(dataTexto, " De ", " de ")
End Function
