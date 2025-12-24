Attribute VB_Name = "M_Utils"
Option Explicit
' ==============================================================================
' MODULO: M_UTILS (REFATORADO)
' DESCRICAO: FERRAMENTAS UTILITARIAS COM CONVERSOES ROBUSTAS
' VERSAO: 2.0 - Integrado com lógica validada de outro sistema
' ==============================================================================

' ==============================================================================
' PERFORMANCE
' ==============================================================================
Public Sub Utils_OtimizarPerformance(Ligar As Boolean)
    With Application
        .ScreenUpdating = Not Ligar
        .EnableEvents = Not Ligar
        .DisplayStatusBar = Not Ligar
        If Ligar Then
            .Calculation = xlCalculationManual
        Else
            .Calculation = xlCalculationAutomatic
        End If
    End With
End Sub

' ==============================================================================
' CONVERSAO DE COORDENADAS - VERSAO ROBUSTA E UNIVERSAL
' ==============================================================================

' ------------------------------------------------------------------------------
' FUNÇÃO PRINCIPAL: DMS PARA DECIMAL (UNIVERSAL)
' ------------------------------------------------------------------------------
' Suporta múltiplos formatos:
' 1. "-43°35'36,463"          (formato atual do sistema)
' 2. "22° 28' 10,2299" S"     (formato com sufixo S/N/O/L)
' 3. "-43.5934619399999974"   (já em decimal - passa direto)
' 4. "43°35'36.463" W         (ponto decimal em vez de vírgula)
' ------------------------------------------------------------------------------
Public Function Str_DMS_Para_DD(ByVal dmsString As String) As Double
    Dim textoOriginal As String, textoLimpo As String
    Dim charAtual As String
    Dim i As Long
    Dim partes() As String
    Dim sinal As Integer
    Dim numGrau As Double, numMin As Double, numSeg As Double

    On Error GoTo ErroConversao

    textoOriginal = Trim(dmsString)
    If textoOriginal = "" Then
        Str_DMS_Para_DD = 0
        Exit Function
    End If

    ' --- DETECÇÃO RÁPIDA: Se já é decimal (não tem ° nem ') ---
    If InStr(textoOriginal, Chr(176)) = 0 And InStr(textoOriginal, "'") = 0 Then
        ' CORREÇÃO: Val() sempre usa ponto como decimal, independente da configuração regional
        ' Normaliza vírgula para ponto primeiro
        Dim decimalNormalizado As String
        decimalNormalizado = Replace(textoOriginal, ",", ".")

        ' Val() ignora configuração regional e sempre usa ponto como decimal
        Str_DMS_Para_DD = Val(decimalNormalizado)
        Exit Function
    End If

    ' --- DETERMINAÇÃO DO SINAL ---
    sinal = 1

    ' Verifica se tem sinal negativo explícito
    If InStr(1, textoOriginal, "-") > 0 Then
        sinal = -1
    ' Ou se tem sufixo indicando hemisfério Sul/Oeste
    ElseIf InStr(1, UCase(textoOriginal), "S") > 0 Or _
           InStr(1, UCase(textoOriginal), "W") > 0 Or _
           InStr(1, UCase(textoOriginal), "O") > 0 Then
        sinal = -1
    End If

    ' --- LIMPEZA DA STRING ---
    ' Remove letras de quadrante
    textoLimpo = textoOriginal
    textoLimpo = Replace(textoLimpo, "N", "", , , vbTextCompare)
    textoLimpo = Replace(textoLimpo, "S", "", , , vbTextCompare)
    textoLimpo = Replace(textoLimpo, "E", "", , , vbTextCompare)
    textoLimpo = Replace(textoLimpo, "L", "", , , vbTextCompare)
    textoLimpo = Replace(textoLimpo, "W", "", , , vbTextCompare)
    textoLimpo = Replace(textoLimpo, "O", "", , , vbTextCompare)

    ' Remove símbolos de grau/minuto/segundo, substituindo por espaços
    textoLimpo = Replace(textoLimpo, Chr(176), " ")  ' Chr(176) = símbolo °
    textoLimpo = Replace(textoLimpo, "'", " ")
    textoLimpo = Replace(textoLimpo, """", " ")
    textoLimpo = Replace(textoLimpo, "-", "")

    ' Normaliza vírgula para ponto decimal (padrão internacional)
    textoLimpo = Replace(textoLimpo, ",", ".")

    ' Remove espaços duplicados
    Do While InStr(textoLimpo, "  ") > 0
        textoLimpo = Replace(textoLimpo, "  ", " ")
    Loop
    textoLimpo = Trim(textoLimpo)

    ' --- EXTRAÇÃO DAS PARTES NUMÉRICAS ---
    partes = Split(textoLimpo, " ")

    numGrau = 0: numMin = 0: numSeg = 0

    If UBound(partes) >= 0 Then numGrau = Val(partes(0))
    If UBound(partes) >= 1 Then numMin = Val(partes(1))
    If UBound(partes) >= 2 Then numSeg = Val(partes(2))

    ' Validação básica
    If numMin >= 60 Or numSeg >= 60 Then GoTo ErroConversao

    ' --- CÁLCULO FINAL ---
    If UBound(partes) = 0 Then
        ' Apenas graus (já é decimal)
        Str_DMS_Para_DD = numGrau * sinal
    Else
        ' Formato DMS completo
        Str_DMS_Para_DD = sinal * (Abs(numGrau) + (numMin / 60) + (numSeg / 3600))
    End If

    Exit Function

ErroConversao:
    Str_DMS_Para_DD = 0
    Debug.Print "Erro ao converter DMS para DD: '" & dmsString & "'"
End Function

' ------------------------------------------------------------------------------
' FUNÇÃO: DECIMAL PARA DMS (FORMATO PADRÃO DO SISTEMA)
' ------------------------------------------------------------------------------
' Retorna formato: "-43°35'36.463" (com 3 casas decimais nos segundos)
' Compatível com o formato atual do sistema DocGEO
' ------------------------------------------------------------------------------
Public Function Str_DD_Para_DMS(ByVal CoordenadaDecimal As Double) As String
    Dim graus As Long, minutos As Long, segundos As Double, tempCoord As Double
    Dim sinal As String

    If CoordenadaDecimal < 0 Then sinal = "-" Else sinal = ""
    tempCoord = Abs(CoordenadaDecimal)
    graus = Int(tempCoord)
    minutos = Int((tempCoord - graus) * 60)
    segundos = ((tempCoord - graus) * 60 - minutos) * 60

    ' Formato padrão do sistema: -GG°MM'SS.SSS"
    Str_DD_Para_DMS = sinal & graus & Chr(176) & Format(minutos, "00") & "'" & Format(segundos, "00.000") & Chr(34)
End Function

' ------------------------------------------------------------------------------
' FUNÇÃO: DECIMAL PARA DMS COM SUFIXO (FORMATO ALTERNATIVO)
' ------------------------------------------------------------------------------
' Retorna formato: "22° 28' 10.2299" S" ou "43° 35' 36.4626" O"
' Para uso em exportações ou documentos que exigem esse formato
' ------------------------------------------------------------------------------
Public Function Str_DD_Para_DMS_ComSufixo(ByVal ValorDecimal As Double, ByVal Tipo As String) As String
    Dim graus As Double, minutos As Long
    Dim segundos As Double
    Dim sufixo As String
    Dim ValorAbs As Double

    ValorAbs = Abs(ValorDecimal)

    graus = Int(ValorAbs)
    minutos = Int((ValorAbs - graus) * 60)
    segundos = (((ValorAbs - graus) * 60) - minutos) * 60

    ' Define o sufixo (N/S para Latitude, L/O para Longitude)
    If UCase(Tipo) = "LAT" Then
        If ValorDecimal < 0 Then sufixo = " S" Else sufixo = " N"
    ElseIf UCase(Tipo) = "LON" Then
        If ValorDecimal < 0 Then sufixo = " O" Else sufixo = " L"
    Else
        sufixo = ""
    End If

    ' Formato com 4 casas decimais e sufixo
    Str_DD_Para_DMS_ComSufixo = CStr(graus) & "° " & Format(minutos, "00") & "' " & _
                                Format(segundos, "00.0000") & """" & sufixo
End Function

' ------------------------------------------------------------------------------
' FUNÇÃO PARA AZIMUTE (GG°MM') - SEM SEGUNDOS
' ------------------------------------------------------------------------------
' Formato usado para azimutes no memorial descritivo
' Exemplo: "123°45'"
' ------------------------------------------------------------------------------
Public Function Str_DD_Para_DM(ByVal grausDecimal As Double) As String
    Dim graus As Double
    Dim minutos As Double
    Dim resto As Double

    grausDecimal = Abs(grausDecimal)

    graus = Int(grausDecimal)

    ' Calcula minutos e arredonda (segundos não são exibidos)
    resto = (grausDecimal - graus) * 60
    minutos = Round(resto, 0)

    ' Ajuste de overflow
    If minutos >= 60 Then
        minutos = 0
        graus = graus + 1
    End If

    ' Retorna apenas GG°MM'
    Str_DD_Para_DM = graus & "°" & Format(minutos, "00") & "'"
End Function

' ------------------------------------------------------------------------------
' FUNÇÃO: FORMATAR AZIMUTE (GRAUS DECIMAIS PARA GGG°MM')
' ------------------------------------------------------------------------------
' Normaliza azimute para 0-360° e formata como 3 dígitos de graus
' Exemplo: 45.5 → "045°30'"
' ------------------------------------------------------------------------------
Public Function Str_FormatAzimute(ByVal azimuteDecimal As Double) As String
    ' Normaliza para 0-360
    If azimuteDecimal < 0 Then azimuteDecimal = azimuteDecimal + 360
    If azimuteDecimal >= 360 Then azimuteDecimal = azimuteDecimal - 360

    Dim minutosTotais As Long, graus As Long, minutos As Long
    minutosTotais = Round(azimuteDecimal * 60, 0)
    graus = minutosTotais \ 60
    minutos = minutosTotais Mod 60
    If graus = 360 Then graus = 0

    Str_FormatAzimute = Format(graus, "000") & Chr(176) & Format(minutos, "00") & "'"
End Function

' ------------------------------------------------------------------------------
' FUNÇÃO: FORMATAR AZIMUTE GMS (GRAUS DECIMAIS PARA GGG°MM'SS")
' ------------------------------------------------------------------------------
' Normaliza azimute para 0-360° e formata com SEGUNDOS
' Usado para coordenadas UTM onde precisão maior é necessária
' Exemplo: 123.9117 → "123°54'42""
' ------------------------------------------------------------------------------
Public Function Str_FormatAzimuteGMS(ByVal azimuteDecimal As Double) As String
    ' Normaliza para 0-360
    If azimuteDecimal < 0 Then azimuteDecimal = azimuteDecimal + 360
    If azimuteDecimal >= 360 Then azimuteDecimal = azimuteDecimal - 360

    Dim graus As Long, minutos As Long, segundos As Long
    Dim tempDecimal As Double

    graus = Int(azimuteDecimal)
    tempDecimal = (azimuteDecimal - graus) * 60
    minutos = Int(tempDecimal)
    segundos = Round((tempDecimal - minutos) * 60, 0)

    ' Ajustes de overflow
    If segundos >= 60 Then
        segundos = segundos - 60
        minutos = minutos + 1
    End If
    If minutos >= 60 Then
        minutos = minutos - 60
        graus = graus + 1
    End If
    If graus >= 360 Then graus = 0

    Str_FormatAzimuteGMS = Format(graus, "000") & Chr(176) & Format(minutos, "00") & "'" & Format(segundos, "00") & Chr(34)
End Function

' ------------------------------------------------------------------------------
' FUNÇÃO: AZIMUTE GMS PARA DECIMAL
' ------------------------------------------------------------------------------
' Converte string de azimute (sem quadrante) para graus decimais
' Exemplo: "123°45'30"" → 123.7583
' ------------------------------------------------------------------------------
Public Function Str_Azimute_Para_DD(ByVal GMS_String As String) As Double
    Dim tempStr As String
    Dim partes() As String
    Dim graus As Double, minutos As Double, segundos As Double

    On Error GoTo ErroConversao
    tempStr = Trim(GMS_String)

    ' Limpa a string, deixando apenas números e espaços
    tempStr = Replace(tempStr, Chr(176), " ")  ' Chr(176) = símbolo °
    tempStr = Replace(tempStr, "'", " ")
    tempStr = Replace(tempStr, """", " ")
    tempStr = Replace(tempStr, ",", ".")

    ' Remove espaços duplicados
    Do While InStr(tempStr, "  ")
        tempStr = Replace(tempStr, "  ", " ")
    Loop
    tempStr = Trim(tempStr)

    partes = Split(tempStr, " ")

    graus = 0: minutos = 0: segundos = 0
    If UBound(partes) >= 0 Then graus = CDbl(partes(0))
    If UBound(partes) >= 1 Then minutos = CDbl(partes(1))
    If UBound(partes) >= 2 Then segundos = CDbl(partes(2))

    Str_Azimute_Para_DD = graus + (minutos / 60) + (segundos / 3600)
    Exit Function

ErroConversao:
    Str_Azimute_Para_DD = -1
End Function

' ==============================================================================
' CONVERSÃO DE RUMO PARA AZIMUTE E VICE-VERSA
' ==============================================================================

' ------------------------------------------------------------------------------
' FUNÇÃO: RUMO PARA AZIMUTE
' ------------------------------------------------------------------------------
' Converte rumo (ex: "N 45°30' E") para azimute (45.5°)
' Quadrantes: NE (0-90), SE (90-180), SW (180-270), NW (270-360)
' ------------------------------------------------------------------------------
Public Function Str_Rumo_Para_Azimute(ByVal RumoString As String) As Double
    Dim quadrante As String
    Dim anguloDecimal As Double
    Dim azimuteFinal As Double
    Dim startQ As String, endQ As String
    Dim anguloStr As String

    On Error GoTo ErroConversao

    RumoString = Trim(UCase(RumoString))

    ' Detecta quadrante Norte/Sul
    If InStr(RumoString, "N") > 0 Then
        startQ = "N"
    ElseIf InStr(RumoString, "S") > 0 Then
        startQ = "S"
    Else
        GoTo ErroConversao
    End If

    ' Detecta quadrante Leste/Oeste
    If InStr(RumoString, "E") > 0 Or InStr(RumoString, "L") > 0 Then
        endQ = "E"
    ElseIf InStr(RumoString, "W") > 0 Or InStr(RumoString, "O") > 0 Then
        endQ = "W"
    Else
        GoTo ErroConversao
    End If

    quadrante = startQ & endQ

    ' Extrai o ângulo removendo letras
    anguloStr = RumoString
    anguloStr = Replace(anguloStr, "N", "")
    anguloStr = Replace(anguloStr, "S", "")
    anguloStr = Replace(anguloStr, "E", "")
    anguloStr = Replace(anguloStr, "L", "")
    anguloStr = Replace(anguloStr, "W", "")
    anguloStr = Replace(anguloStr, "O", "")
    anguloStr = Trim(anguloStr)

    ' Converte o ângulo para decimal
    anguloDecimal = Str_DMS_Para_DD(anguloStr)
    If anguloDecimal < 0 Or anguloDecimal > 90 Then GoTo ErroConversao

    ' Aplica regra de conversão baseada no quadrante
    Select Case quadrante
        Case "NE": azimuteFinal = anguloDecimal
        Case "SE": azimuteFinal = 180 - anguloDecimal
        Case "SW": azimuteFinal = 180 + anguloDecimal
        Case "NW": azimuteFinal = 360 - anguloDecimal
        Case Else: GoTo ErroConversao
    End Select

    Str_Rumo_Para_Azimute = azimuteFinal
    Exit Function

ErroConversao:
    Str_Rumo_Para_Azimute = -1
End Function

' ------------------------------------------------------------------------------
' FUNÇÃO: AZIMUTE PARA RUMO
' ------------------------------------------------------------------------------
' Converte azimute (45.5°) para rumo ("N 45°30' E")
' ------------------------------------------------------------------------------
Public Function Str_Azimute_Para_Rumo(ByVal AzimuteDecimal As Double) As String
    Dim anguloRumo As Double
    Dim quadrante As String
    Dim rumoGMS As String

    ' Determina quadrante e ângulo
    If AzimuteDecimal >= 0 And AzimuteDecimal < 90 Then
        quadrante = "NE"
        anguloRumo = AzimuteDecimal
    ElseIf AzimuteDecimal >= 90 And AzimuteDecimal < 180 Then
        quadrante = "SE"
        anguloRumo = 180 - AzimuteDecimal
    ElseIf AzimuteDecimal >= 180 And AzimuteDecimal < 270 Then
        quadrante = "SW"
        anguloRumo = AzimuteDecimal - 180
    ElseIf AzimuteDecimal >= 270 And AzimuteDecimal <= 360 Then
        quadrante = "NW"
        anguloRumo = 360 - AzimuteDecimal
    End If

    ' Casos especiais (eixos cardeais)
    If AzimuteDecimal = 0 Then
        Str_Azimute_Para_Rumo = "Norte"
        Exit Function
    ElseIf AzimuteDecimal = 90 Then
        Str_Azimute_Para_Rumo = "Leste"
        Exit Function
    ElseIf AzimuteDecimal = 180 Then
        Str_Azimute_Para_Rumo = "Sul"
        Exit Function
    ElseIf AzimuteDecimal = 270 Then
        Str_Azimute_Para_Rumo = "Oeste"
        Exit Function
    End If

    ' Formata ângulo em DMS
    rumoGMS = Str_DD_Para_DMS(anguloRumo)

    ' Monta string final
    Str_Azimute_Para_Rumo = rumoGMS & " " & quadrante
End Function

' ==============================================================================
' STRINGS
' ==============================================================================
Public Function Str_ExtractBetween(texto As String, rotuloInicio As String, _
                                    ParamArray rotulosFimPossiveis() As Variant) As String
    On Error GoTo ErroExtracao

    Dim posInicio As Long, posValor As Long, posFim As Long
    Dim posFinalDefinitiva As Long, rotulo As Variant

    posInicio = InStr(1, texto, rotuloInicio, vbTextCompare)
    If posInicio = 0 Then GoTo ErroExtracao

    posValor = posInicio + Len(rotuloInicio)
    posFinalDefinitiva = Len(texto) + 1

    If InStr(posValor, texto, vbCr) > 0 Then posFinalDefinitiva = InStr(posValor, texto, vbCr)

    For Each rotulo In rotulosFimPossiveis
        posFim = InStr(posValor, texto, CStr(rotulo), vbTextCompare)
        If posFim > 0 And posFim < posFinalDefinitiva Then posFinalDefinitiva = posFim
    Next rotulo

    Str_ExtractBetween = Trim(Mid(texto, posValor, posFinalDefinitiva - posValor))
    Exit Function

ErroExtracao:
    Str_ExtractBetween = ""
End Function

Public Function Str_LimparCaractereWord(texto As String) As String
    Dim s As String: s = texto
    If Len(s) > 2 Then
        If Asc(Right(s, 1)) < 32 Then s = Left(s, Len(s) - 1)
        If Asc(Right(s, 1)) < 32 Then s = Left(s, Len(s) - 1)
    End If
    Str_LimparCaractereWord = Trim(s)
End Function

Public Function File_SanitizeName(ByVal filename As String) As String
    Dim invalidChars As Variant, i As Long
    invalidChars = Array("/", "\", ":", "*", "?", Chr(34), "<", ">", "|")
    File_SanitizeName = filename
    For i = LBound(invalidChars) To UBound(invalidChars)
        File_SanitizeName = Replace(File_SanitizeName, invalidChars(i), "")
    Next i
    File_SanitizeName = Trim(File_SanitizeName)
End Function

Public Function SanitizeFilename(ByVal filename As String) As String
    SanitizeFilename = File_SanitizeName(filename)
End Function

' ==============================================================================
' SELETOR DE PASTA
' ==============================================================================
Public Function UI_SelecionarPasta() As String
    Dim diag As FileDialog
    Set diag = Application.FileDialog(msoFileDialogFolderPicker)
    With diag
        .Title = "Selecione a pasta"
        .AllowMultiSelect = False
        If .Show = -1 Then
            UI_SelecionarPasta = .SelectedItems(1)
            If Right(UI_SelecionarPasta, 1) <> "\" Then UI_SelecionarPasta = UI_SelecionarPasta & "\"
        Else
            UI_SelecionarPasta = ""
        End If
    End With
End Function

' ==============================================================================
' BUSCA EM CADASTROS
' ==============================================================================
Public Function GetCadastroValue(label As String, Optional occurrence As Long = 1) As String
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim searchRange As Range
    Dim firstFound As Range, currentFound As Range
    Dim i As Long

    On Error GoTo NotFound
    Set ws = ThisWorkbook.Sheets(M_Config.SHEET_CADASTROS)
    Set tbl = ws.ListObjects(M_Config.TBL_CADASTROS)
    Set searchRange = tbl.ListColumns(1).DataBodyRange
    Set currentFound = searchRange.Find(What:=label, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)

    If Not currentFound Is Nothing Then
        If occurrence > 1 Then
            Set firstFound = currentFound
            For i = 2 To occurrence
                Set currentFound = searchRange.FindNext(After:=currentFound)
                If currentFound Is Nothing Or currentFound.Address = firstFound.Address Then
                    GetCadastroValue = ""
                    Exit Function
                End If
            Next i
        End If
        GetCadastroValue = CStr(currentFound.Offset(0, 1).Value)
    Else
NotFound:
        GetCadastroValue = ""
    End If
End Function
