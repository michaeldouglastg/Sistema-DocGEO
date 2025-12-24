Attribute VB_Name = "VERIFICAR_VERSAO_M_App_Logica"
Option Explicit

Sub Verificar_Versao_M_App_Logica()
    ' ==============================================================================
    ' VERIFICA QUAL VERSÃO DE M_App_Logica ESTÁ NO EXCEL
    ' ==============================================================================

    Dim resultado As String
    resultado = "=== VERIFICAÇÃO: VERSÃO DE M_App_Logica ===" & vbCrLf & vbCrLf

    ' Testa se a macro foi atualizada
    ' A versão atualizada tem comentários específicos no código
    ' Vamos verificar indiretamente pelo comportamento

    ' Cria dados de teste na tabela SGL
    Dim wsSGL As Worksheet
    Dim loSGL As ListObject

    On Error Resume Next
    Set wsSGL = ThisWorkbook.Sheets(M_Config.SH_SGL)
    Set loSGL = wsSGL.ListObjects(M_Config.TBL_SGL)
    On Error GoTo 0

    If loSGL Is Nothing Then
        MsgBox "Tabela SGL não encontrada!", vbCritical
        Exit Sub
    End If

    ' Verifica se existe dados na SGL
    If loSGL.ListRows.Count = 0 Then
        resultado = resultado & "⚠️ ATENÇÃO: Tabela SGL está vazia!" & vbCrLf
        resultado = resultado & "   Não é possível testar a conversão." & vbCrLf
        resultado = resultado & vbCrLf
        resultado = resultado & "SOLUÇÃO:" & vbCrLf
        resultado = resultado & "1. Importe o CSV primeiro" & vbCrLf
        resultado = resultado & "2. Execute este teste novamente" & vbCrLf

        MsgBox resultado, vbExclamation, "Tabela Vazia"
        Exit Sub
    End If

    ' Pega primeira linha
    Dim lonStr As String, latStr As String
    lonStr = CStr(loSGL.ListRows(1).Range(2).Value)
    latStr = CStr(loSGL.ListRows(1).Range(3).Value)

    Dim lonDD As Double, latDD As Double
    lonDD = M_Utils.Str_DMS_Para_DD(lonStr)
    latDD = M_Utils.Str_DMS_Para_DD(latStr)

    resultado = resultado & "COORDENADAS LIDAS DA PLANILHA SGL:" & vbCrLf
    resultado = resultado & "  Longitude: " & lonStr & " → " & Format(lonDD, "0.00000000") & "°" & vbCrLf
    resultado = resultado & "  Latitude: " & latStr & " → " & Format(latDD, "0.00000000") & "°" & vbCrLf
    resultado = resultado & vbCrLf

    ' Testa conversão manual (como deveria ser)
    Dim fusoCorreto As Integer
    fusoCorreto = M_Math_Geo.Geo_GetZonaUTM(lonDD)

    Dim utmCorreto As Type_UTM
    utmCorreto = M_Math_Geo.Converter_GeoParaUTM(latDD, lonDD, fusoCorreto)

    resultado = resultado & "CONVERSÃO MANUAL (como deveria ser):" & vbCrLf
    resultado = resultado & "  Fuso: " & fusoCorreto & vbCrLf
    resultado = resultado & "  Norte: " & Format(utmCorreto.Norte, "0.0000") & vbCrLf
    resultado = resultado & "  Leste: " & Format(utmCorreto.Leste, "0.0000") & vbCrLf
    resultado = resultado & vbCrLf

    ' Verifica o que está na planilha UTM
    Dim wsUTM As Worksheet
    Dim loUTM As ListObject

    On Error Resume Next
    Set wsUTM = ThisWorkbook.Sheets(M_Config.SH_UTM)
    Set loUTM = wsUTM.ListObjects(M_Config.TBL_UTM)

    If Not loUTM Is Nothing And loUTM.ListRows.Count > 0 Then
        Dim norteNaPlanilha As Double, lesteNaPlanilha As Double
        norteNaPlanilha = CDbl(loUTM.ListRows(1).Range(2).Value)
        lesteNaPlanilha = CDbl(loUTM.ListRows(1).Range(3).Value)

        resultado = resultado & "VALORES NA PLANILHA UTM (gerados pelo sistema):" & vbCrLf
        resultado = resultado & "  Norte: " & Format(norteNaPlanilha, "0.0000") & vbCrLf
        resultado = resultado & "  Leste: " & Format(lesteNaPlanilha, "0.0000") & vbCrLf
        resultado = resultado & vbCrLf

        Dim diffN As Double, diffE As Double
        diffN = norteNaPlanilha - utmCorreto.Norte
        diffE = lesteNaPlanilha - utmCorreto.Leste

        resultado = resultado & "DIFERENÇA (Planilha vs Manual):" & vbCrLf
        resultado = resultado & "  Delta Norte: " & Format(diffN, "0.00") & " m"
        If Abs(diffN) < 1 Then
            resultado = resultado & " ✅ CORRETO!"
        ElseIf Abs(diffN) < 100 Then
            resultado = resultado & " ⚠️ PEQUENA DIFERENÇA"
        Else
            resultado = resultado & " ❌ DIFERENÇA GRANDE!"
        End If
        resultado = resultado & vbCrLf

        resultado = resultado & "  Delta Leste: " & Format(diffE, "0.00") & " m"
        If Abs(diffE) < 1 Then
            resultado = resultado & " ✅ CORRETO!"
        ElseIf Abs(diffE) < 100 Then
            resultado = resultado & " ⚠️ PEQUENA DIFERENÇA"
        Else
            resultado = resultado & " ❌ DIFERENÇA GRANDE!"
        End If
        resultado = resultado & vbCrLf
        resultado = resultado & vbCrLf

        resultado = resultado & "================================" & vbCrLf
        resultado = resultado & "DIAGNÓSTICO:" & vbCrLf
        resultado = resultado & vbCrLf

        If Abs(diffN) > 100 Or Abs(diffE) > 100 Then
            resultado = resultado & "❌ M_App_Logica NÃO FOI ATUALIZADO!" & vbCrLf
            resultado = resultado & vbCrLf
            resultado = resultado & "SOLUÇÃO:" & vbCrLf
            resultado = resultado & "1. Abra VBA (Alt+F11)" & vbCrLf
            resultado = resultado & "2. Remova o módulo M_App_Logica antigo" & vbCrLf
            resultado = resultado & "3. File → Import File... → M_App_Logica.bas" & vbCrLf
            resultado = resultado & "4. Re-importe o CSV" & vbCrLf
        Else
            resultado = resultado & "✅ M_App_Logica está ATUALIZADO!" & vbCrLf
            resultado = resultado & vbCrLf
            resultado = resultado & "As coordenadas UTM estão corretas!" & vbCrLf
        End If
    Else
        resultado = resultado & "⚠️ Tabela UTM está vazia!" & vbCrLf
        resultado = resultado & "   Execute a conversão SGL→UTM primeiro." & vbCrLf
    End If
    On Error GoTo 0

    Debug.Print resultado
    MsgBox resultado, vbInformation, "Verificação de Versão"
End Sub
