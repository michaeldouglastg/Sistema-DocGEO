Attribute VB_Name = "M_DOC_DXF"
Option Explicit

' ==============================================================================
' MÓDULO: M_DOC_DXF
' DESCRIÇÃO: EXPORTAÇÃO DA GEOMETRIA PARA ARQUIVO DXF (AUTOCAD)
' ==============================================================================

Public Sub GerarArquivoDXF(dadosProp As Object, Optional PastaExportacao As String = "")
    Dim wsUTM As Worksheet
    Dim loUTM As ListObject
    Dim i As Long, qtd As Long
    Dim caminhoArquivo As String, nomeArquivo As String
    Dim fNum As Integer
    Dim x As Double, y As Double, Z As Double
    Dim ptNome As String
    
    ' 1. Configuração e Referências
    Set wsUTM = ThisWorkbook.Sheets(M_Config.SH_UTM)
    Set loUTM = wsUTM.ListObjects(M_Config.TBL_UTM)
    
    If loUTM.ListRows.Count < 2 Then
        MsgBox "Não há dados suficientes na tabela UTM para gerar o DXF.", vbExclamation
        Exit Sub
    End If
    
    Call Utils_OtimizarPerformance(True)
    
    ' 2. Definir Nome do Arquivo
'    nomeArquivo = "Planta_DXF_" & M_Utils.File_SanitizeName(dadosProp(M_Config.LBL_PROPRIEDADE)) & ".dxf"
'    caminhoArquivo = ThisWorkbook.path & "\" & nomeArquivo
    If PastaExportacao <> "" Then
            caminhoArquivo = PastaExportacao
        Else
            caminhoArquivo = M_Utils.UI_SelecionarPasta()
            If caminhoArquivo = "" Then Exit Sub
        End If
        
        nomeArquivo = "Planta_DXF_..."
        caminhoArquivo = caminhoArquivo & nomeArquivo
    
    ' 3. Iniciar Escrita do Arquivo Texto
    fNum = FreeFile
    Open caminhoArquivo For Output As #fNum
    
    ' --- CABEÇALHO DXF (Minimalista - Compatível R12) ---
    Print #fNum, "0"
    Print #fNum, "SECTION"
    Print #fNum, "2"
    Print #fNum, "HEADER"
    Print #fNum, "9"
    Print #fNum, "$ACADVER"
    Print #fNum, "1"
    Print #fNum, "AC1009" ' Versão R11/R12 (Abre em qualquer AutoCAD)
    Print #fNum, "0"
    Print #fNum, "ENDSEC"
    
    ' --- TABELAS (LAYERS) ---
    Print #fNum, "0"
    Print #fNum, "SECTION"
    Print #fNum, "2"
    Print #fNum, "TABLES"
    Print #fNum, "0"
    Print #fNum, "TABLE"
    Print #fNum, "2"
    Print #fNum, "LAYER"
    Print #fNum, "70"
    Print #fNum, "2" ' Quantidade de layers definidos abaixo
    
    ' Layer 1: PERIMETRO (Cor 1 = Vermelho)
    Print #fNum, "0"; vbCrLf; "LAYER"
    Print #fNum, "2"; vbCrLf; "PERIMETRO"
    Print #fNum, "70"; vbCrLf; "0"
    Print #fNum, "62"; vbCrLf; "1"
    Print #fNum, "6"; vbCrLf; "CONTINUOUS"
    
    ' Layer 2: TEXTO (Cor 7 = Branco/Preto)
    Print #fNum, "0"; vbCrLf; "LAYER"
    Print #fNum, "2"; vbCrLf; "TEXTO"
    Print #fNum, "70"; vbCrLf; "0"
    Print #fNum, "62"; vbCrLf; "7"
    Print #fNum, "6"; vbCrLf; "CONTINUOUS"
    
    Print #fNum, "0"
    Print #fNum, "ENDTAB"
    Print #fNum, "0"
    Print #fNum, "ENDSEC"
    
    ' --- ENTIDADES (DESENHO) ---
    Print #fNum, "0"
    Print #fNum, "SECTION"
    Print #fNum, "2"
    Print #fNum, "ENTITIES"
    
    qtd = loUTM.ListRows.Count
    
    ' 4. Desenhar Polígono (POLYLINE)
    Print #fNum, "0"
    Print #fNum, "POLYLINE"
    Print #fNum, "8" ' Layer
    Print #fNum, "PERIMETRO"
    Print #fNum, "66" ' Flag: Seguem vértices
    Print #fNum, "1"
    Print #fNum, "10"; vbCrLf; "0.0"
    Print #fNum, "20"; vbCrLf; "0.0"
    Print #fNum, "30"; vbCrLf; "0.0"
    
    ' Loop dos Vértices
    For i = 1 To qtd
        x = CDbl(loUTM.DataBodyRange(i, 3).Value) ' E (X)
        y = CDbl(loUTM.DataBodyRange(i, 2).Value) ' N (Y)
        Z = 0 ' Z (Opcional: pode pegar da coluna Altitude se quiser 3D)
        
        Print #fNum, "0"
        Print #fNum, "VERTEX"
        Print #fNum, "8"
        Print #fNum, "PERIMETRO"
        Print #fNum, "10"
        Print #fNum, Replace(CStr(x), ",", ".") ' Força ponto decimal (Padrão CAD)
        Print #fNum, "20"
        Print #fNum, Replace(CStr(y), ",", ".")
        Print #fNum, "30"
        Print #fNum, Replace(CStr(Z), ",", ".")
    Next i
    
    ' Fechamento do Polígono (Repete o primeiro ponto)
    x = CDbl(loUTM.DataBodyRange(1, 3).Value)
    y = CDbl(loUTM.DataBodyRange(1, 2).Value)
    Print #fNum, "0"; vbCrLf; "VERTEX"
    Print #fNum, "8"; vbCrLf; "PERIMETRO"
    Print #fNum, "10"; vbCrLf; Replace(CStr(x), ",", ".")
    Print #fNum, "20"; vbCrLf; Replace(CStr(y), ",", ".")
    Print #fNum, "30"; vbCrLf; "0.0"
    
    Print #fNum, "0"
    Print #fNum, "SEQEND" ' Fim da Polyline
    
    ' 5. Desenhar Nomes dos Pontos (TEXT)
    For i = 1 To qtd
        ptNome = loUTM.DataBodyRange(i, 1).Value
        x = CDbl(loUTM.DataBodyRange(i, 3).Value)
        y = CDbl(loUTM.DataBodyRange(i, 2).Value)
        
        Print #fNum, "0"
        Print #fNum, "TEXT"
        Print #fNum, "8"
        Print #fNum, "TEXTO"
        Print #fNum, "10"
        Print #fNum, Replace(CStr(x + 1), ",", ".") ' Desloca 1m para não ficar em cima da linha
        Print #fNum, "20"
        Print #fNum, Replace(CStr(y + 1), ",", ".")
        Print #fNum, "40"
        Print #fNum, "2.0" ' Altura do Texto
        Print #fNum, "1"
        Print #fNum, ptNome
    Next i
    
    ' Fim do Arquivo
    Print #fNum, "0"
    Print #fNum, "ENDSEC"
    Print #fNum, "0"
    Print #fNum, "EOF"
    
    Close #fNum
    
    Call Utils_OtimizarPerformance(False)
    
    If MsgBox("Arquivo DXF gerado com sucesso!" & vbCrLf & vbCrLf & _
              "Salvo em: " & nomeArquivo & vbCrLf & vbCrLf & _
              "Deseja abrir a pasta do arquivo?", vbYesNo + vbInformation) = vbYes Then
        Shell "explorer.exe /select," & caminhoArquivo, vbNormalFocus
    End If
End Sub

