Attribute VB_Name = "M_Config"
Option Explicit
' ==============================================================================
' MODULO: M_CONFIG (UNIFICADO)
' DESCRICAO: CENTRALIZA TODAS AS CONSTANTES E FUNCOES DE ESTADO
' ==============================================================================

' --- NOMES DAS PLANILHAS ---
Public Const SH_PAINEL As String = "PAINEL_PRINCIPAL"
Public Const SH_SGL As String = "DADOS_PRINCIPAL_SGL"
Public Const SH_UTM As String = "DADOS_PRINCIPAL_UTM"
Public Const SH_PARAMETROS As String = "PARAMETROS"
Public Const SH_CADASTROS As String = "CADASTROS"
Public Const SH_VISUALIZACAO As String = "PRE-VISUALIZAR"
Public Const SH_BD_PROP As String = "BD_PROPRIEDADES"
Public Const SH_BD_TEC As String = "BD_TECNICOS"
Public Const SH_TEMP_CONV As String = "TEMP_CONVERSAO"
Public Const SH_CROQUI As String = "CROQUI"
Public Const SH_MAPA As String = "MAPA10X"
Public Const SH_UX_DESIGN As String = "UX"

' --- NOMES DAS TABELAS ---
Public Const TBL_SGL As String = "tbl_Principal_SGL"
Public Const TBL_UTM As String = "tbl_Principal_UTM"
Public Const TBL_PARAMETROS As String = "tbl_Parametros"
Public Const TBL_CADASTROS As String = "tbl_Cadastros"
Public Const TBL_DB_PROP As String = "tbl_PropriedadesDB"
Public Const TBL_DB_TEC As String = "tbl_TecnicosDB"
Public Const TBL_CONVERSAO As String = "tbl_Conversao"

' --- CONTROLES E CELULAS ---
Public Const LB_PRINCIPAL As String = "lstPrincipal"

' Celulas SGL
Public Const CELL_SGL_AREA_HA As String = "AreaSGL"
Public Const CELL_SGL_AREA_M2 As String = "AreaM2"
Public Const CELL_SGL_PERIMETRO As String = "Perimetro"

' Celulas UTM
Public Const CELL_UTM_AREA_HA As String = "AreaSGL2"
Public Const CELL_UTM_AREA_M2 As String = "AreaM22"
Public Const CELL_UTM_PERIMETRO As String = "Perimetro2"

' --- CONSTANTES GEODESICAS ---
Public Const CONST_RAIO_TERRA As Double = 6371000
Public Const CONST_PI As Double = 3.14159265358979
Public Const CONST_SEMI_EIXO_MAIOR As Double = 6378137#
Public Const CONST_ACHATAMENTO As Double = 0.00335281068118
Public Const CONST_FATOR_K0 As Double = 0.9996

' --- ROTULOS DE CADASTRO ---
Public Const LBL_PROPRIEDADE As String = "Propriedade"
Public Const LBL_MATRICULA As String = "Matricula"
Public Const LBL_INCRA As String = "Codigo Incra"
Public Const LBL_MUNICIPIO As String = "Municipio"
Public Const LBL_COMARCA As String = "Comarca"
Public Const LBL_CARTORIO As String = "Cartorio"
Public Const LBL_NATUREZA As String = "Natureza/Area"
Public Const LBL_CONFRONTANTE_CPF As String = "CPF Confrontante"

Public Const LBL_PROP_NOME As String = "Proprietario"
Public Const LBL_PROP_CPF As String = "CPF"
Public Const LBL_PROP_NACIONALIDADE As String = "Nacionalidade"
Public Const LBL_PROP_ESTADO_CIVIL As String = "Estado Civil"
Public Const LBL_PROP_PROFISSAO As String = "Profissao"
Public Const LBL_PROP_RG As String = "RG"
Public Const LBL_PROP_RG_EXPEDICAO As String = "Expedicao"
Public Const LBL_PROP_RG_DATA As String = "Data Expedicao"
Public Const LBL_PROP_ENDERECO As String = "Endereco Completo"

Public Const LBL_RT_NOME As String = "Responsavel Tecnico"
Public Const LBL_RT_ART As String = "TRT/ART"
Public Const LBL_RT_TITULO As String = "Titulo"
Public Const LBL_RT_REGISTRO As String = "Registro (CFT/CREA)"
Public Const LBL_RT_INCRA As String = "Codigo INCRA"

Public Const LBL_MAPA_TITULO As String = "PLANTA TOPOGRAFICA - RETIFICACAO"
Public Const LBL_DATUM As String = "SIRGAS 2000"
Public Const LBL_MERIDIANO As String = "45 WGr"

' --- ALIASES PARA COMPATIBILIDADE (M_Config1) ---
Public Const SHEET_PRINCIPAL As String = "PAINEL_PRINCIPAL"
Public Const SHEET_DADOS_PRINCIPAL_SGL As String = "DADOS_PRINCIPAL_SGL"
Public Const SHEET_DADOS_PRINCIPAL_UTM As String = "DADOS_PRINCIPAL_UTM"
Public Const SHEET_PARAMETROS As String = "PARAMETROS"
Public Const SHEET_CADASTROS As String = "CADASTROS"
Public Const SHEET_VISUALIZACAO As String = "PRE-VISUALIZAR"
Public Const SHEET_BD_PROPRIEDADES As String = "BD_PROPRIEDADES"
Public Const SHEET_BD_TECNICOS As String = "BD_TECNICOS"
Public Const TBL_PROPRIEDADES_DB As String = "tbl_PropriedadesDB"
Public Const TBL_TECNICOS_DB As String = "tbl_TecnicosDB"
Public Const LISTBOX_PRINCIPAL As String = "lstPrincipal"
Public Const CELL_VISUALIZACAO_OUTPUT As String = "C7"
Public Const RAIO_TERRA_M As Double = 6371000

' Aliases de Rotulos (CAD_)
Public Const CAD_PROPRIEDADE As String = "Propriedade"
Public Const CAD_MATRICULA As String = "Matricula"
Public Const CAD_INCRA As String = "Codigo Incra"
Public Const CAD_MUNICIPIO As String = "Municipio"
Public Const CAD_COMARCA As String = "Comarca"
Public Const CAD_PROPRIETARIO_NOME As String = "Proprietario"
Public Const CAD_PROPRIETARIO_CPF As String = "CPF"
Public Const CAD_PROPRIETARIO_NACIONALIDADE As String = "Nacionalidade"
Public Const CAD_PROPRIETARIO_ESTADO_CIVIL As String = "Estado Civil"
Public Const CAD_PROPRIETARIO_PROFISSAO As String = "Profissao"
Public Const CAD_PROPRIETARIO_RG As String = "RG"
Public Const CAD_PROPRIETARIO_RG_EXPEDICAO As String = "Expedicao"
Public Const CAD_PROPRIETARIO_RG_DATA As String = "Data Expedicao"
Public Const CAD_PROPRIETARIO_ENDERECO As String = "Endereco Completo"
Public Const CAD_CONFRONTANTE_CPF As String = "CPF Confrontante"
Public Const CAD_RT_NOME As String = "Responsavel Tecnico"
Public Const CAD_RT_TITULO As String = "Titulo"
Public Const CAD_RT_REGISTRO As String = "Registro (CFT/CREA)"
Public Const CAD_RT_INCRA As String = "Codigo INCRA"
Public Const CAD_RT_ART As String = "TRT/ART"
Public Const CAD_CARTORIO As String = "Cartorio"
Public Const CAD_NATUREZA As String = "Natureza/Area"

' Nomes dos ComboBox UTM
Public Const CBO_FUSO As String = "cboFuso"
Public Const CBO_HEMISFERIO As String = "cboHemisferio"

' ==============================================================================
' FUNCOES DE ESTADO
' ==============================================================================
Public Function App_GetSistemaAtivo() As String
    Dim wsPainel As Worksheet
    Dim optSGL As Object
    
    On Error Resume Next
    Set wsPainel = ThisWorkbook.Sheets(SH_PAINEL)
    If wsPainel Is Nothing Then
        App_GetSistemaAtivo = "SGL"
        Exit Function
    End If
    
    Set optSGL = wsPainel.OLEObjects("optSGL").Object
    If Not optSGL Is Nothing Then
        If optSGL.Value = True Then
            App_GetSistemaAtivo = "SGL"
        Else
            App_GetSistemaAtivo = "UTM"
        End If
    Else
        App_GetSistemaAtivo = "SGL"
    End If
    On Error GoTo 0
End Function

Public Function App_GetNomeAbaAtiva() As String
    If App_GetSistemaAtivo() = "SGL" Then
        App_GetNomeAbaAtiva = SH_SGL
    Else
        App_GetNomeAbaAtiva = SH_UTM
    End If
End Function

Public Function App_GetNomeTabelaAtiva() As String
    If App_GetSistemaAtivo() = "SGL" Then
        App_GetNomeTabelaAtiva = TBL_SGL
    Else
        App_GetNomeTabelaAtiva = TBL_UTM
    End If
End Function

' --- ALIASES PARA COMPATIBILIDADE ---
Public Function ObterSistemaAtivo() As String
    ObterSistemaAtivo = App_GetSistemaAtivo()
End Function

Public Function ObterNomeAbaAtiva() As String
    ObterNomeAbaAtiva = App_GetNomeAbaAtiva()
End Function

Public Function ObterNomeTabelaAtiva() As String
    ObterNomeTabelaAtiva = App_GetNomeTabelaAtiva()
End Function
