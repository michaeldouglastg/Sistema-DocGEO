Attribute VB_Name = "M_Graficos"
Option Explicit
' ==============================================================================
' MODULO: M_GRAFICOS
' DESCRICAO: CONTROLE AVANCADO DE GRAFICOS (ZOOM, PAN, PLOTAGEM, ROTULOS).
' ==============================================================================

Private Const NOME_GRAFICO As String = "Mapa"

' ==============================================================================
' 1. BOTAO DE RESET (ALINHAR E CENTRALIZAR) - VERSAO FORCADA
' ==============================================================================
Public Sub Grafico_Reset(nomePlanilha As String)
    Dim ws As Worksheet
    Dim cht As Chart
    
    On Error GoTo ErroReset
    Set ws = ThisWorkbook.Sheets(nomePlanilha)
    
    ' Workaround: Ativa a planilha para garantir que o Excel desenhe o grafico
    If ActiveSheet.Name <> nomePlanilha Then ws.Activate
    
    ' Verifica existencia
    If ws.ChartObjects.Count = 0 Then
        MsgBox "Nenhum grafico encontrado na aba " & nomePlanilha, vbExclamation
        Exit Sub
    End If
    
    ' Tenta pegar o grafico pelo nome
    On Error Resume Next
    Set cht = ws.ChartObjects(NOME_GRAFICO).Chart
    On Error GoTo ErroReset
    
    If cht Is Nothing Then
        MsgBox "O grafico com nome '" & NOME_GRAFICO & "' nao foi encontrado na aba " & nomePlanilha & "." & vbCrLf & _
               "Verifique se o nome do grafico esta correto.", vbCritical
        Exit Sub
    End If
    
    Call M_Utils.Utils_OtimizarPerformance(True)
    
    ' 1. RESET TOTAL: Forca eixos para Automatico (Remove Zoom Travado)
    With cht.Axes(xlCategory)
        .MinimumScaleIsAuto = True
        .MaximumScaleIsAuto = True
    End With
    With cht.Axes(xlValue)
        .MinimumScaleIsAuto = True
        .MaximumScaleIsAuto = True
    End With
    
    ' 2. Redesenha o Poligono e Aplica Escala Proporcional
    Call Grafico_PlotarPoligono(nomePlanilha)
    
    Call M_Utils.Utils_OtimizarPerformance(False)
    Exit Sub
    
ErroReset:
    Call M_Utils.Utils_OtimizarPerformance(False)
    MsgBox "Erro ao resetar grafico: " & Err.Description, vbCritical
End Sub

' ==============================================================================
' 2. PLOTAGEM DE POLIGONO (COM FECHAMENTO E ROTULOS)
' ==============================================================================
Public Sub Grafico_PlotarPoligono(nomePlanilha As String)
    Dim wsGrafico As Worksheet, wsDados As Worksheet
    Dim loDados As ListObject
    Dim cht As Chart
    Dim arrX() As Double, arrY() As Double, arrLabels() As String
    Dim i As Long, qtd As Long
    
    On Error GoTo ErroPlotagem
    Set wsGrafico = ThisWorkbook.Sheets(nomePlanilha)
    
    ' Evita erro se o grafico nao existir
    On Error Resume Next
    Set cht = wsGrafico.ChartObjects(NOME_GRAFICO).Chart
    On Error GoTo ErroPlotagem
    
    If cht Is Nothing Then Exit Sub
    
    Set wsDados = ThisWorkbook.Sheets(M_Config.SH_UTM)
    Set loDados = wsDados.ListObjects(M_Config.TBL_UTM)
    
    If loDados.ListRows.Count < 2 Then Exit Sub
    
    ' --- Carregar Dados ---
    Dim arrRaw As Variant
    arrRaw = loDados.DataBodyRange.Value
    qtd = UBound(arrRaw, 1)
    
    ReDim arrX(1 To qtd + 1)
    ReDim arrY(1 To qtd + 1)
    ReDim arrLabels(1 To qtd + 1)
    
    For i = 1 To qtd
        arrLabels(i) = CStr(arrRaw(i, 1))
        arrY(i) = CDbl(arrRaw(i, 2)) ' N
        arrX(i) = CDbl(arrRaw(i, 3)) ' E
    Next i
    
    ' Ponto de fechamento
    arrLabels(qtd + 1) = arrLabels(1)
    arrY(qtd + 1) = arrY(1)
    arrX(qtd + 1) = arrX(1)
    
    ' --- Atualizar Serie ---
    ' Limpa series antigas
    Do While cht.SeriesCollection.Count > 0
        cht.SeriesCollection(1).Delete
    Loop
    
    With cht.SeriesCollection.NewSeries
        .Name = "Perimetro"
        .XValues = arrX
        .Values = arrY
        .ChartType = xlXYScatterLines
        
        ' Rotulos
        .HasDataLabels = True
        .DataLabels.Position = xlLabelPositionAbove
        
        Dim pt As Point
        For i = 1 To qtd + 1
            Set pt = .Points(i)
            pt.DataLabel.Text = arrLabels(i)
            pt.DataLabel.Font.Size = 8
        Next i
    End With
    
    With cht.SeriesCollection(1).Format.Line
        .ForeColor.RGB = RGB(0, 0, 0)
        .Weight = 1
    End With
    
    ' --- Ajustar Escala Proporcional ---
    Call Grafico_AjustarEscalaProporcional(cht, arrX, arrY)
    
    Exit Sub
    
ErroPlotagem:
    ' Falha silenciosa na plotagem automatica
End Sub

' ==============================================================================
' AJUSTE DE ESCALA PROPORCIONAL (PRIVADA)
' ==============================================================================
Private Sub Grafico_AjustarEscalaProporcional(cht As Chart, arrX() As Double, arrY() As Double)
    Dim minX As Double, maxX As Double, minY As Double, maxY As Double
    Dim deltaX As Double, deltaY As Double
    Dim centroX As Double, centroY As Double
    Dim plotW As Double, plotH As Double
    Dim novoDelta As Double
    
    ' Calcula limites dos DADOS (nao do grafico atual)
    minX = arrX(1): maxX = arrX(1): minY = arrY(1): maxY = arrY(1)
    
    Dim i As Long
    For i = LBound(arrX) To UBound(arrX)
        If arrX(i) < minX Then minX = arrX(i)
        If arrX(i) > maxX Then maxX = arrX(i)
        If arrY(i) < minY Then minY = arrY(i)
        If arrY(i) > maxY Then maxY = arrY(i)
    Next i
    
    deltaX = maxX - minX
    deltaY = maxY - minY
    centroX = (minX + maxX) / 2
    centroY = (minY + maxY) / 2
    
    ' Margem de 10%
    deltaX = deltaX * 1.1
    deltaY = deltaY * 1.1
    
    If deltaX = 0 Then deltaX = 50 ' Evita erro em ponto unico
    If deltaY = 0 Then deltaY = 50
    
    ' Tenta obter tamanho da plotagem
    On Error Resume Next
    plotW = cht.PlotArea.Width
    plotH = cht.PlotArea.Height
    If plotW <= 0 Then plotW = 100
    If plotH <= 0 Then plotH = 100
    On Error GoTo 0
    
    ' Aplica Proporcao 1:1
    If (deltaX / plotW) > (deltaY / plotH) Then
        ' Ajusta Y para crescer
        novoDelta = deltaX * (plotH / plotW)
        With cht.Axes(xlCategory)
            .MinimumScale = centroX - (deltaX / 2)
            .MaximumScale = centroX + (deltaX / 2)
        End With
        With cht.Axes(xlValue)
            .MinimumScale = centroY - (novoDelta / 2)
            .MaximumScale = centroY + (novoDelta / 2)
        End With
    Else
        ' Ajusta X para crescer
        novoDelta = deltaY * (plotW / plotH)
        With cht.Axes(xlValue)
            .MinimumScale = centroY - (deltaY / 2)
            .MaximumScale = centroY + (deltaY / 2)
        End With
        With cht.Axes(xlCategory)
            .MinimumScale = centroX - (novoDelta / 2)
            .MaximumScale = centroX + (novoDelta / 2)
        End With
    End If
End Sub

' ==============================================================================
' 3. ZOOM E PAN (SIMPLIFICADOS)
' ==============================================================================
Public Sub Grafico_Zoom(nomePlanilha As String, ByVal FatorZoom As Double)
    Dim ws As Worksheet
    Dim cht As Chart
    Dim minX As Double, maxX As Double, minY As Double, maxY As Double
    Dim spanX As Double, spanY As Double, centroX As Double, centroY As Double
    
    Set ws = ThisWorkbook.Sheets(nomePlanilha)
    
    On Error Resume Next
    Set cht = ws.ChartObjects(NOME_GRAFICO).Chart
    If cht Is Nothing Then Exit Sub
    On Error GoTo 0
    
    With cht.Axes(xlCategory)
        minX = .MinimumScale
        maxX = .MaximumScale
        spanX = maxX - minX
        centroX = (minX + maxX) / 2
        .MinimumScale = centroX - (spanX / FatorZoom / 2)
        .MaximumScale = centroX + (spanX / FatorZoom / 2)
    End With
    
    With cht.Axes(xlValue)
        minY = .MinimumScale
        maxY = .MaximumScale
        spanY = maxY - minY
        centroY = (minY + maxY) / 2
        .MinimumScale = centroY - (spanY / FatorZoom / 2)
        .MaximumScale = centroY + (spanY / FatorZoom / 2)
    End With
End Sub

Public Sub Grafico_Pan(nomePlanilha As String, Direcao As String)
    Dim ws As Worksheet
    Dim cht As Chart
    Dim deltaX As Double, deltaY As Double
    
    Set ws = ThisWorkbook.Sheets(nomePlanilha)
    
    On Error Resume Next
    Set cht = ws.ChartObjects(NOME_GRAFICO).Chart
    If cht Is Nothing Then Exit Sub
    On Error GoTo 0
    
    With cht.Axes(xlCategory)
        deltaX = (.MaximumScale - .MinimumScale) * 0.1
    End With
    With cht.Axes(xlValue)
        deltaY = (.MaximumScale - .MinimumScale) * 0.1
    End With
    
    Select Case UCase(Direcao)
        Case "CIMA": Call MoverEixo(cht.Axes(xlValue), deltaY)
        Case "BAIXO": Call MoverEixo(cht.Axes(xlValue), -deltaY)
        Case "ESQUERDA": Call MoverEixo(cht.Axes(xlCategory), -deltaX)
        Case "DIREITA": Call MoverEixo(cht.Axes(xlCategory), deltaX)
    End Select
End Sub

Private Sub MoverEixo(eixo As Axis, delta As Double)
    eixo.MinimumScale = eixo.MinimumScale + delta
    eixo.MaximumScale = eixo.MaximumScale + delta
End Sub

' ==============================================================================
' 4. CONTROLE DE ROTULOS (MOSTRAR/OCULTAR)
' ==============================================================================
Public Sub Grafico_AlternarRotulos(nomePlanilha As String)
    Dim ws As Worksheet
    Dim cht As Chart
    
    On Error GoTo ErroRotulo
    Set ws = ThisWorkbook.Sheets(nomePlanilha)
    
    ' Garante que o grafico exista
    If ws.ChartObjects.Count = 0 Then Exit Sub
    
    On Error Resume Next
    Set cht = ws.ChartObjects(NOME_GRAFICO).Chart
    On Error GoTo ErroRotulo
    
    If cht Is Nothing Then Exit Sub
    
    Call M_Utils.Utils_OtimizarPerformance(True)
    
    ' Verifica se existe alguma serie plotada
    If cht.SeriesCollection.Count > 0 Then
        With cht.SeriesCollection(1)
            ' Inverte o estado atual (Se True vira False, se False vira True)
            .HasDataLabels = Not .HasDataLabels
            
            ' Se estiver ativando, garante a posicao correta
            If .HasDataLabels Then
                .DataLabels.Position = xlLabelPositionAbove
            End If
        End With
    End If
    
    Call M_Utils.Utils_OtimizarPerformance(False)
    Exit Sub
    
ErroRotulo:
    Call M_Utils.Utils_OtimizarPerformance(False)
    MsgBox "Erro ao alterar rotulos: " & Err.Description, vbExclamation
End Sub

' ==============================================================================
' 5. LIMPEZA VISUAL DO GRAFICO (REMOVE SERIES, MANTEM FORMATACAO)
' ==============================================================================
Public Sub Grafico_Limpar(nomePlanilha As String)
    Dim ws As Worksheet
    Dim cht As Chart
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(nomePlanilha)
    
    ' Tenta pegar o grafico
    If ws.ChartObjects.Count > 0 Then
        Set cht = ws.ChartObjects(NOME_GRAFICO).Chart
        If Not cht Is Nothing Then
            ' Remove todas as series de dados (linhas/pontos)
            Do While cht.SeriesCollection.Count > 0
                cht.SeriesCollection(1).Delete
            Loop
        End If
    End If
    On Error GoTo 0
End Sub


