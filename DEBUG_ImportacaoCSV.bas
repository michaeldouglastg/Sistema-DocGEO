Attribute VB_Name = "DEBUG_ImportacaoCSV"
Option Explicit

Sub Debug_Primeira_Linha_CSV()
    ' ==============================================================================
    ' DEBUG: Verificar como a primeira linha do CSV está sendo lida
    ' ==============================================================================

    Dim wsSGL As Worksheet
    Dim loSGL As ListObject
    Dim resultado As String

    On Error Resume Next
    Set wsSGL = ThisWorkbook.Sheets(M_Config.SH_SGL)
    Set loSGL = wsSGL.ListObjects(M_Config.TBL_SGL)
    On Error GoTo 0

    If loSGL Is Nothing Then
        MsgBox "Tabela SGL não encontrada!", vbCritical
        Exit Sub
    End If

    If loSGL.ListRows.Count = 0 Then
        MsgBox "Tabela SGL está vazia! Importe o CSV primeiro.", vbExclamation
        Exit Sub
    End If

    resultado = "=== DEBUG: PRIMEIRA LINHA DO CSV ===" & vbCrLf & vbCrLf

    ' Pega valores brutos da primeira linha
    Dim nomeRaw As String, lonRaw As String, latRaw As String
    nomeRaw = CStr(loSGL.ListRows(1).Range(1).Value)
    lonRaw = CStr(loSGL.ListRows(1).Range(2).Value)
    latRaw = CStr(loSGL.ListRows(1).Range(3).Value)

    resultado = resultado & "VALORES BRUTOS NA PLANILHA SGL:" & vbCrLf
    resultado = resultado & "  Nome: " & nomeRaw & vbCrLf
    resultado = resultado & "  Longitude (col 2): " & lonRaw & vbCrLf
    resultado = resultado & "  Latitude (col 3): " & latRaw & vbCrLf
    resultado = resultado & vbCrLf

    ' Converte para decimal
    Dim lonDD As Double, latDD As Double
    lonDD = M_Utils.Str_DMS_Para_DD(lonRaw)
    latDD = M_Utils.Str_DMS_Para_DD(latRaw)

    resultado = resultado & "CONVERSÃO DMS → DD:" & vbCrLf
    resultado = resultado & "  Longitude DD: " & Format(lonDD, "0.00000000") & vbCrLf
    resultado = resultado & "  Latitude DD: " & Format(latDD, "0.00000000") & vbCrLf
    resultado = resultado & vbCrLf

    ' Verifica se valores estão corretos
    resultado = resultado & "VALORES ESPERADOS:" & vbCrLf
    resultado = resultado & "  Longitude: -43.59346194" & vbCrLf
    resultado = resultado & "  Latitude: -22.46950833" & vbCrLf
    resultado = resultado & vbCrLf

    Dim diffLon As Double, diffLat As Double
    diffLon = lonDD - (-43.59346194)
    diffLat = latDD - (-22.46950833)

    resultado = resultado & "DIFERENÇAS:" & vbCrLf
    resultado = resultado & "  Delta Lon: " & Format(diffLon, "0.00000000")
    If Abs(diffLon) > 0.001 Then resultado = resultado & " ❌ GRANDE!"
    resultado = resultado & vbCrLf
    resultado = resultado & "  Delta Lat: " & Format(diffLat, "0.00000000")
    If Abs(diffLat) > 0.001 Then resultado = resultado & " ❌ GRANDE!"
    resultado = resultado & vbCrLf
    resultado = resultado & vbCrLf

    ' Calcula UTM com esses valores
    Dim fusoDetectado As Integer
    fusoDetectado = M_Math_Geo.Geo_GetZonaUTM(lonDD)

    resultado = resultado & "FUSO DETECTADO:" & vbCrLf
    resultado = resultado & "  Fuso: " & fusoDetectado
    If fusoDetectado <> 23 Then resultado = resultado & " ❌ ESPERADO: 23"
    resultado = resultado & vbCrLf
    resultado = resultado & vbCrLf

    ' Converte para UTM
    Dim utm As Type_UTM
    utm = M_Math_Geo.Converter_GeoParaUTM(latDD, lonDD, fusoDetectado)

    resultado = resultado & "CONVERSÃO GEO → UTM:" & vbCrLf
    resultado = resultado & "  Norte: " & Format(utm.Norte, "0.0000") & vbCrLf
    resultado = resultado & "  Leste: " & Format(utm.Leste, "0.0000") & vbCrLf
    resultado = resultado & vbCrLf

    resultado = resultado & "ESPERADO:" & vbCrLf
    resultado = resultado & "  Norte: 7514524.6000" & vbCrLf
    resultado = resultado & "  Leste: 644711.6600" & vbCrLf
    resultado = resultado & vbCrLf

    Dim diffN As Double, diffE As Double
    diffN = utm.Norte - 7514524.6
    diffE = utm.Leste - 644711.66

    resultado = resultado & "DIFERENÇAS UTM:" & vbCrLf
    resultado = resultado & "  Delta Norte: " & Format(diffN, "0.00") & " m"
    If Abs(diffN) > 100 Then resultado = resultado & " ❌ GRANDE!"
    resultado = resultado & vbCrLf
    resultado = resultado & "  Delta Leste: " & Format(diffE, "0.00") & " m"
    If Abs(diffE) > 100 Then resultado = resultado & " ❌ GRANDE!"
    resultado = resultado & vbCrLf
    resultado = resultado & vbCrLf

    ' Verifica valores na tabela UTM
    Dim wsUTM As Worksheet
    Dim loUTM As ListObject
    On Error Resume Next
    Set wsUTM = ThisWorkbook.Sheets(M_Config.SH_UTM)
    Set loUTM = wsUTM.ListObjects(M_Config.TBL_UTM)

    If Not loUTM Is Nothing And loUTM.ListRows.Count > 0 Then
        Dim utmNorteNaPlanilha As Double, utmLesteNaPlanilha As Double
        utmNorteNaPlanilha = CDbl(loUTM.ListRows(1).Range(2).Value)
        utmLesteNaPlanilha = CDbl(loUTM.ListRows(1).Range(3).Value)

        resultado = resultado & "================================" & vbCrLf
        resultado = resultado & "VALORES ATUAIS NA PLANILHA UTM:" & vbCrLf
        resultado = resultado & "  Norte: " & Format(utmNorteNaPlanilha, "0.0000") & vbCrLf
        resultado = resultado & "  Leste: " & Format(utmLesteNaPlanilha, "0.0000") & vbCrLf
        resultado = resultado & vbCrLf

        Dim diffNPlanilha As Double, diffEPlanilha As Double
        diffNPlanilha = utmNorteNaPlanilha - utm.Norte
        diffEPlanilha = utmLesteNaPlanilha - utm.Leste

        resultado = resultado & "DIFERENÇA (Planilha vs Calculado agora):" & vbCrLf
        resultado = resultado & "  Delta Norte: " & Format(diffNPlanilha, "0.00") & " m"
        If Abs(diffNPlanilha) > 10 Then resultado = resultado & " ❌ DIFERENTE!"
        resultado = resultado & vbCrLf
        resultado = resultado & "  Delta Leste: " & Format(diffEPlanilha, "0.00") & " m"
        If Abs(diffEPlanilha) > 10 Then resultado = resultado & " ❌ DIFERENTE!"
        resultado = resultado & vbCrLf
    End If
    On Error GoTo 0

    Debug.Print resultado
    MsgBox resultado, vbInformation, "Debug Importação CSV"
End Sub
