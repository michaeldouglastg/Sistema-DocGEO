Attribute VB_Name = "M_Utils"
Option Explicit
' ==============================================================================
' MODULO: M_UTILS
' DESCRICAO: FERRAMENTAS UTILITARIAS
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
' CONVERSAO DE COORDENADAS
' ==============================================================================
Public Function Str_DMS_Para_DD(ByVal dmsString As String) As Double
    Dim textoOriginal As String, textoLimpo As String
    Dim charAtual As String
    Dim i As Long
    Dim partes() As String
    Dim sinal As Integer
    Dim numGrau As Double, numMin As Double, numSeg As Double
    
    textoOriginal = Trim(dmsString)
    If textoOriginal = "" Then Exit Function
    
    sinal = 1
    If InStr(1, textoOriginal, "-") > 0 Then
        sinal = -1
    ElseIf InStr(1, UCase(textoOriginal), "S") > 0 Or _
           InStr(1, UCase(textoOriginal), "W") > 0 Or _
           InStr(1, UCase(textoOriginal), "O") > 0 Then
        sinal = -1
    End If
    
    textoLimpo = ""
    For i = 1 To Len(textoOriginal)
        charAtual = Mid(textoOriginal, i, 1)
        If IsNumeric(charAtual) Or charAtual = "," Or charAtual = "." Then
            textoLimpo = textoLimpo & charAtual
        Else
            textoLimpo = textoLimpo & " "
        End If
    Next i
    
    textoLimpo = Application.WorksheetFunction.Trim(textoLimpo)
    partes = Split(textoLimpo, " ")
    
    numGrau = 0: numMin = 0: numSeg = 0
    If UBound(partes) >= 0 Then numGrau = Val(Replace(partes(0), ",", "."))
    If UBound(partes) >= 1 Then numMin = Val(Replace(partes(1), ",", "."))
    If UBound(partes) >= 2 Then numSeg = Val(Replace(partes(2), ",", "."))
    
    If UBound(partes) = 0 Then
        Str_DMS_Para_DD = numGrau * sinal
    Else
        Str_DMS_Para_DD = sinal * (Abs(numGrau) + (numMin / 60) + (numSeg / 3600))
    End If
End Function

Public Function Str_DD_Para_DMS(ByVal CoordenadaDecimal As Double) As String
    Dim graus As Long, minutos As Long, segundos As Double, tempCoord As Double
    Dim sinal As String
    
    If CoordenadaDecimal < 0 Then sinal = "-" Else sinal = ""
    tempCoord = Abs(CoordenadaDecimal)
    graus = Int(tempCoord)
    minutos = Int((tempCoord - graus) * 60)
    segundos = ((tempCoord - graus) * 60 - minutos) * 60
    
    Str_DD_Para_DMS = sinal & graus & Chr(176) & Format(minutos, "00") & "'" & Format(segundos, "00.000") & Chr(34)
End Function

' ------------------------------------------------------------------------------
' FUNÇÃO PARA AZIMUTE SGL (GG°MM')
' ------------------------------------------------------------------------------
Public Function Str_DD_Para_DM(ByVal grausDecimal As Double) As String
    Dim graus As Double
    Dim minutos As Double
    Dim resto As Double
    
    grausDecimal = Abs(grausDecimal)
    
    graus = Int(grausDecimal)
    
    ' Calcula minutos e arredonda o resto (segundos não são exibidos, então arredondamos o minuto)
    resto = (grausDecimal - graus) * 60
    minutos = Round(resto, 0) ' Arredonda para o minuto mais próximo
    
    If minutos >= 60 Then
        minutos = 0
        graus = graus + 1
    End If
    
    ' Retorna apenas GG°MM'
    Str_DD_Para_DM = graus & "°" & Format(minutos, "00") & "'"
End Function

Public Function Str_FormatAzimute(ByVal azimuteDecimal As Double) As String
    If azimuteDecimal < 0 Then azimuteDecimal = azimuteDecimal + 360
    If azimuteDecimal >= 360 Then azimuteDecimal = azimuteDecimal - 360

    Dim minutosTotais As Long, graus As Long, minutos As Long
    minutosTotais = Round(azimuteDecimal * 60, 0)
    graus = minutosTotais \ 60
    minutos = minutosTotais Mod 60
    If graus = 360 Then graus = 0

    Str_FormatAzimute = Format(graus, "000") & Chr(176) & Format(minutos, "00") & "'"
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
