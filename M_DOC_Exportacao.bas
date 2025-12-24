Attribute VB_Name = "M_DOC_Exportacao"
Option Explicit

' ==============================================================================
' MÓDULO: M_DOC_EXPORTACAO
' DESCRIÇÃO: EXPORTAÇÃO DE DADOS PARA KML (GOOGLE EARTH) E DXF (AUTOCAD)
' ==============================================================================

' ==============================================================================
' 1. EXPORTAR PARA DXF (AUTOCAD)
' ==============================================================================
Public Sub ExportarDXF(dadosProp As Object)
    Dim wsUTM As Worksheet
    Dim loUTM As ListObject
    Dim i As Long, qtd As Long
    Dim pastaSalvar As String, nomeArquivo As String, caminhoCompleto As String
    Dim fNum As Integer
    Dim x As Double, y As Double, Z As Double
    Dim xProx As Double, yProx As Double
    Dim ptNome As String
    
    ' 1. Referências
    Set wsUTM = ThisWorkbook.Sheets(M_Config.SH_UTM)
    Set loUTM = wsUTM.ListObjects(M_Config.TBL_UTM)
    
    If loUTM.ListRows.Count < 2 Then
        MsgBox "Não há dados suficientes na tabela UTM para gerar o DXF.", vbExclamation
        Exit Sub
    End If
    
    ' 2. Selecionar Pasta
    pastaSalvar = M_Utils.UI_SelecionarPasta()
    If pastaSalvar = "" Then Exit Sub
    
    nomeArquivo = "Planta_" & M_Utils.File_SanitizeName(dadosProp(M_Config.LBL_PROPRIEDADE)) & ".dxf"
    caminhoCompleto = pastaSalvar & nomeArquivo
    
    Call Utils_OtimizarPerformance(True)
    
    ' 3. Gerar Arquivo
    fNum = FreeFile
    Open caminhoCompleto For Output As #fNum
    
    ' --- CABEÇALHO DXF MÍNIMO ---
    Print #fNum, "0"
    Print #fNum, "SECTION"
    Print #fNum, "2"
    Print #fNum, "ENTITIES"
    
    qtd = loUTM.ListRows.Count
    
    ' 4. Loop pelos pontos (Linhas, Textos e Pontos)
    For i = 1 To qtd
        ' Mapeamento: Col 1=Nome, Col 2=N(Y), Col 3=E(X), Col 4=Alt(Z)
        ptNome = CStr(loUTM.DataBodyRange(i, 1).Value)
        y = CDbl(loUTM.DataBodyRange(i, 2).Value) ' Norte
        x = CDbl(loUTM.DataBodyRange(i, 3).Value) ' Este
        
        ' Z (Opcional - se vazio assume 0)
        If IsNumeric(loUTM.DataBodyRange(i, 4).Value) Then
            Z = CDbl(loUTM.DataBodyRange(i, 4).Value)
        Else
            Z = 0
        End If
        
        ' Define o próximo ponto para fechar a linha
        Dim idxProx As Long
        If i < qtd Then idxProx = i + 1 Else idxProx = 1
        
        yProx = CDbl(loUTM.DataBodyRange(idxProx, 2).Value)
        xProx = CDbl(loUTM.DataBodyRange(idxProx, 3).Value)
        
        ' ENTIDADE: LINE (Linha do perímetro)
        Print #fNum, "0"
        Print #fNum, "LINE"
        Print #fNum, "8" ' Layer
        Print #fNum, "PERIMETRO"
        Print #fNum, "10"
        Print #fNum, Replace(CStr(x), ",", ".") ' X Inicial
        Print #fNum, "20"
        Print #fNum, Replace(CStr(y), ",", ".") ' Y Inicial
        Print #fNum, "30"
        Print #fNum, "0.0"
        Print #fNum, "11"
        Print #fNum, Replace(CStr(xProx), ",", ".") ' X Final
        Print #fNum, "21"
        Print #fNum, Replace(CStr(yProx), ",", ".") ' Y Final
        Print #fNum, "31"
        Print #fNum, "0.0"
        
        ' ENTIDADE: TEXT (Nome do Ponto)
        Print #fNum, "0"
        Print #fNum, "TEXT"
        Print #fNum, "8"
        Print #fNum, "TEXTO"
        Print #fNum, "10"
        Print #fNum, Replace(CStr(x + 1), ",", ".") ' Desloca um pouco
        Print #fNum, "20"
        Print #fNum, Replace(CStr(y + 1), ",", ".")
        Print #fNum, "30"
        Print #fNum, "0.0"
        Print #fNum, "40"
        Print #fNum, "2.0" ' Altura do texto
        Print #fNum, "1"
        Print #fNum, ptNome
        
        ' ENTIDADE: POINT (Marcador)
        Print #fNum, "0"
        Print #fNum, "POINT"
        Print #fNum, "8"
        Print #fNum, "PONTOS"
        Print #fNum, "10"
        Print #fNum, Replace(CStr(x), ",", ".")
        Print #fNum, "20"
        Print #fNum, Replace(CStr(y), ",", ".")
        Print #fNum, "30"
        Print #fNum, Replace(CStr(Z), ",", ".")
    Next i
    
    ' Rodapé DXF
    Print #fNum, "0"
    Print #fNum, "ENDSEC"
    Print #fNum, "0"
    Print #fNum, "EOF"
    
    Close #fNum
    
    Call Utils_OtimizarPerformance(False)
    
    If MsgBox("Arquivo DXF gerado!" & vbCrLf & nomeArquivo & vbCrLf & "Deseja abrir o local?", vbYesNo + vbInformation) = vbYes Then
        Shell "explorer.exe /select," & caminhoCompleto, vbNormalFocus
    End If
End Sub

' ==============================================================================
' 2. EXPORTAR PARA KML (GOOGLE EARTH)
' ==============================================================================
Public Sub ExportarKML(dadosProp As Object)
    Dim wsSGL As Worksheet
    Dim loSGL As ListObject
    Dim i As Long, qtd As Long
    Dim pastaSalvar As String, nomeArquivo As String, caminhoCompleto As String
    Dim kmlHeader As String, kmlFooter As String, kmlBody As String
    Dim coordsPoligono As String
    Dim latDD As Double, lonDD As Double
    Dim ptNome As String
    
    ' 1. Referências (Usa a tabela SGL pois ela já tem Lat/Lon originais)
    Set wsSGL = ThisWorkbook.Sheets(M_Config.SH_SGL)
    Set loSGL = wsSGL.ListObjects(M_Config.TBL_SGL)
    
    If loSGL.ListRows.Count < 2 Then
        MsgBox "Dados insuficientes para KML.", vbExclamation
        Exit Sub
    End If
    
    ' 2. Selecionar Pasta
    pastaSalvar = M_Utils.UI_SelecionarPasta()
    If pastaSalvar = "" Then Exit Sub
    
    nomeArquivo = "GoogleEarth_" & M_Utils.File_SanitizeName(dadosProp(M_Config.LBL_PROPRIEDADE)) & ".kml"
    caminhoCompleto = pastaSalvar & nomeArquivo
    
    Call Utils_OtimizarPerformance(True)
    
    ' 3. Montagem do KML
    qtd = loSGL.ListRows.Count
    coordsPoligono = ""
    
    kmlHeader = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf & _
                "<kml xmlns=""http://www.opengis.net/kml/2.2"">" & vbCrLf & _
                "<Document>" & vbCrLf & _
                "<name>" & dadosProp(M_Config.LBL_PROPRIEDADE) & "</name>" & vbCrLf & _
                "<Style id=""polyStyle""><LineStyle><color>ff0000ff</color><width>2</width></LineStyle><PolyStyle><color>400000ff</color></PolyStyle></Style>"
    
    kmlBody = ""
    
    ' Loop para Pontos e Polígono
    For i = 1 To qtd
        ptNome = loSGL.DataBodyRange(i, 1).Value
        
        ' Converte SGL (DMS) para Decimal usando nosso M_Utils robusto
        lonDD = M_Utils.Str_DMS_Para_DD(CStr(loSGL.DataBodyRange(i, 2).Value))
        latDD = M_Utils.Str_DMS_Para_DD(CStr(loSGL.DataBodyRange(i, 3).Value))
        
        ' KML usa formato: Longitude,Latitude,Altitude
        Dim coordStr As String
        coordStr = Replace(CStr(lonDD), ",", ".") & "," & Replace(CStr(latDD), ",", ".") & ",0"
        
        ' Adiciona ao Polígono
        coordsPoligono = coordsPoligono & coordStr & " "
        
        ' Adiciona Placemark (Ponto individual)
        kmlBody = kmlBody & "<Placemark>" & vbCrLf & _
                  "<name>" & ptNome & "</name>" & vbCrLf & _
                  "<Point><coordinates>" & coordStr & "</coordinates></Point>" & vbCrLf & _
                  "</Placemark>" & vbCrLf
    Next i
    
    ' Fecha o Polígono (Repete o primeiro ponto)
    Dim lon1 As Double, lat1 As Double
    lon1 = M_Utils.Str_DMS_Para_DD(CStr(loSGL.DataBodyRange(1, 2).Value))
    lat1 = M_Utils.Str_DMS_Para_DD(CStr(loSGL.DataBodyRange(1, 3).Value))
    coordsPoligono = coordsPoligono & Replace(CStr(lon1), ",", ".") & "," & Replace(CStr(lat1), ",", ".") & ",0"
    
    ' Monta o Objeto Polígono
    Dim kmlPoly As String
    kmlPoly = "<Placemark>" & vbCrLf & _
              "<name>Perímetro</name>" & vbCrLf & _
              "<styleUrl>#polyStyle</styleUrl>" & vbCrLf & _
              "<Polygon><outerBoundaryIs><LinearRing><coordinates>" & vbCrLf & _
              coordsPoligono & vbCrLf & _
              "</coordinates></LinearRing></outerBoundaryIs></Polygon>" & vbCrLf & _
              "</Placemark>" & vbCrLf
              
    kmlFooter = "</Document></kml>"
    
    ' 4. Salvar Arquivo
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    With stream
        .Type = 2 ' Text
        .Charset = "UTF-8"
        .Open
        .WriteText kmlHeader & kmlPoly & kmlBody & kmlFooter
        .SaveToFile caminhoCompleto, 2 ' Overwrite
        .Close
    End With
    
    Call Utils_OtimizarPerformance(False)
    MsgBox "Arquivo KML gerado com sucesso!", vbInformation
End Sub

