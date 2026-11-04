Attribute VB_Name = "modConfig"
Option Explicit
Option Private Module   ' rotinas internas: fora da lista de macros
'==============================================================================
' modConfig
' Constantes do projeto, leitura do tema ativo e criacao da estrutura basica
' (planilhas Config, Dados e Dashboard, tabela de temas e validacao).
'
' Observacao: este arquivo e ASCII puro (sem acentos) para nao corromper na
' importacao pelo Editor do VBA no Mac. Textos acentuados exibidos ao usuario
' sao montados com modUtil.U("...\u00e1...").
'==============================================================================

' Nomes fixos (sem colchetes: o Excel os proibe em nomes de planilha)
Public Const PLAN_CONFIG As String = "Config"
Public Const PLAN_DADOS As String = "Dados"
Public Const PLAN_DASH As String = "Dashboard"
Public Const TAB_DADOS As String = "tblDados"
Public Const TAB_TEMAS As String = "tblTemas"
Public Const NOME_TEMA_ATIVO As String = "cfgTemaAtivo"
Public Const NOME_LISTA_TEMAS As String = "lstTemas"

' Posicao das colunas na tabela de temas (1 = Nome do Tema)
Public Const COL_NOME As Long = 1
Public Const COL_PRIMARIA As Long = 2      ' R, G, B em 2, 3, 4
Public Const COL_SECUNDARIA As Long = 5    ' R, G, B em 5, 6, 7
Public Const COL_FUNDO As Long = 8         ' R, G, B em 8, 9, 10
Public Const COL_TEXTO As Long = 11        ' R, G, B em 11, 12, 13
Public Const COL_FONTE As Long = 14
Public Const NUM_COLS_TEMA As Long = 14

' Tema resolvido em cores RGB prontas para uso
Public Type TemaDS
    Nome As String
    Primaria As Long
    Secundaria As Long
    Fundo As Long
    Texto As Long
    Fonte As String
End Type

'------------------------------------------------------------------------------
' GarantirEstrutura
' Proposito: garantir que o livro tenha Config (com temas e Tema Ativo), Dados
'            e Dashboard. Cria o que faltar; nunca apaga dados existentes.
' Parametros: wb - livro alvo.
' Retorno:    True se a estrutura esta pronta; False se o usuario cancelou.
' Efeitos:    pode adicionar ou renomear planilhas, criar nomes definidos,
'             a tabela tblTemas e a validacao de dados do Tema Ativo.
'------------------------------------------------------------------------------
Public Function GarantirEstrutura(ByVal wb As Workbook) As Boolean
    Dim wsCfg As Worksheet, wsDados As Worksheet, wsDash As Worksheet, wsAtiva As Worksheet

    ' Guarda a planilha ativa ANTES de criar Config: Worksheets.Add ativa a
    ' planilha nova (achado no teste do suplemento .xlam).
    If ActiveWorkbook Is wb And TypeName(ActiveSheet) = "Worksheet" Then Set wsAtiva = ActiveSheet

    Set wsCfg = modUtil.ObterPlanilha(wb, PLAN_CONFIG)
    If wsCfg Is Nothing Then
        Set wsCfg = wb.Worksheets.Add(Before:=wb.Sheets(1))
        wsCfg.Name = PLAN_CONFIG
    End If
    GarantirTabelaTemas wsCfg

    Set wsDados = modUtil.ObterPlanilha(wb, PLAN_DADOS)
    If wsDados Is Nothing Then
        Set wsDados = OferecerPlanilhaAtivaComoDados(wsAtiva)
        If wsDados Is Nothing Then
            Set wsDados = wb.Worksheets.Add(After:=wsCfg)
            wsDados.Name = PLAN_DADOS
        End If
    End If

    Set wsDash = modUtil.ObterPlanilha(wb, PLAN_DASH)
    If wsDash Is Nothing Then
        Set wsDash = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        wsDash.Name = PLAN_DASH
    End If

    GarantirEstrutura = True
End Function

'------------------------------------------------------------------------------
' OferecerPlanilhaAtivaComoDados
' Proposito: no uso como suplemento (.xlam), aproveitar a planilha ativa que ja
'            tem dados, renomeando-a para "Dados" com confirmacao do usuario.
' Parametros: ws - planilha que estava ativa quando o usuario clicou
'             (Nothing se nao era uma planilha do livro alvo).
' Retorno:    a planilha renomeada, ou Nothing se nao houver candidata/recusa.
' Efeitos:    renomeia a planilha (reversivel pelo usuario).
'------------------------------------------------------------------------------
Private Function OferecerPlanilhaAtivaComoDados(ByVal ws As Worksheet) As Worksheet
    Dim resp As VbMsgBoxResult

    If ws Is Nothing Then Exit Function
    If ws.Name = PLAN_CONFIG Or ws.Name = PLAN_DASH Then Exit Function
    If Application.WorksheetFunction.CountA(ws.Range("A1").CurrentRegion) = 0 Then Exit Function

    resp = MsgBox(modUtil.U("A planilha """) & ws.Name & modUtil.U( _
        """ tem dados a partir de A1." & vbLf & vbLf & _
        "Usar esta planilha como fonte do dashboard? Ela ser\u00e1 renomeada para ""Dados""."), _
        vbYesNo + vbQuestion, "Dashboard")
    If resp = vbYes Then
        ws.Name = PLAN_DADOS
        Set OferecerPlanilhaAtivaComoDados = ws
    End If
End Function

'------------------------------------------------------------------------------
' GarantirTabelaTemas
' Proposito: criar, se faltar, o bloco de configuracao em Config: titulo,
'            celula Tema Ativo (nome cfgTemaAtivo), tabela tblTemas com quatro
'            temas, nome lstTemas e validacao em lista.
' Parametros: ws - planilha Config.
' Efeitos:    escreve em A1:N10 apenas quando tblTemas nao existe e A6 esta vazio.
'------------------------------------------------------------------------------
Private Sub GarantirTabelaTemas(ByVal ws As Worksheet)
    Dim wb As Workbook, lo As ListObject, dados As Variant, nm As Name

    Set wb = ws.Parent
    Set lo = modUtil.ObterTabela(ws, TAB_TEMAS)

    If lo Is Nothing Then
        If Len(modUtil.TextoSeguro(ws.Range("A6").Value)) > 0 Then
            Err.Raise vbObjectError + 601, "GarantirTabelaTemas", modUtil.U( _
                "Config!A6 j\u00e1 tem conte\u00fado, mas a tabela tblTemas n\u00e3o existe. " & _
                "Libere A6:N10 ou recrie a planilha Config.")
        End If
        ws.Range("A1").Value = modUtil.U("Configura\u00e7\u00e3o do Dashboard")
        ws.Range("A1").Font.Bold = True
        ws.Range("A1").Font.Size = 16
        ws.Range("A3").Value = "Tema Ativo"
        ws.Range("A3").Font.Bold = True
        ws.Range("A4").Value = modUtil.U("Escolha o tema na lista ao lado; o dashboard muda na hora.")
        ws.Range("A4").Font.Italic = True

        dados = TemasPadrao()
        ws.Range("A6").Resize(UBound(dados, 1), UBound(dados, 2)).Value = dados
        Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A6").Resize(UBound(dados, 1), UBound(dados, 2)), , xlYes)
        lo.Name = TAB_TEMAS
        ws.Range("B3").Value = dados(2, COL_NOME)
        ws.Columns("A").ColumnWidth = 22
        ws.Range("B:M").ColumnWidth = 9
        ws.Columns("N").ColumnWidth = 16
        ws.Range("B3").ColumnWidth = 22
    End If

    ' Nomes definidos (idempotente)
    On Error Resume Next
    Set nm = wb.Names(NOME_TEMA_ATIVO)
    On Error GoTo 0
    If nm Is Nothing Then wb.Names.Add Name:=NOME_TEMA_ATIVO, RefersTo:="='" & PLAN_CONFIG & "'!$B$3"
    Set nm = Nothing
    On Error Resume Next
    Set nm = wb.Names(NOME_LISTA_TEMAS)
    On Error GoTo 0
    If nm Is Nothing Then wb.Names.Add Name:=NOME_LISTA_TEMAS, RefersTo:="=" & TAB_TEMAS & "[" & lo.ListColumns(COL_NOME).Name & "]"

    ' Validacao em lista do Tema Ativo: so recria se faltar ou estiver errada.
    ' (Recriar sempre falhou com a janela do livro minimizada: erro 1004.)
    Dim formula1 As String
    On Error Resume Next
    formula1 = wb.Names(NOME_TEMA_ATIVO).RefersToRange.Validation.Formula1
    On Error GoTo 0
    If formula1 <> "=" & NOME_LISTA_TEMAS Then
        On Error Resume Next
        With wb.Names(NOME_TEMA_ATIVO).RefersToRange.Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=" & NOME_LISTA_TEMAS
            .IgnoreBlank = False
            .InCellDropdown = True
            .ErrorTitle = "Tema"
            .ErrorMessage = modUtil.U("Escolha um tema da tabela tblTemas.")
        End With
        If Err.Number <> 0 Then Application.StatusBar = modUtil.U("Aviso: n\u00e3o foi poss\u00edvel recriar a lista do Tema Ativo.")
        On Error GoTo 0
    End If
End Sub

'------------------------------------------------------------------------------
' TemasPadrao
' Proposito: devolver a tabela de temas inicial (cabecalho + 4 temas, um escuro).
' Retorno:   matriz Variant(1 To 5, 1 To 14).
' Efeitos:   nenhum.
'------------------------------------------------------------------------------
Public Function TemasPadrao() As Variant
    Dim linhas As Variant, m() As Variant, i As Long, j As Long

    linhas = Array( _
        Array("Nome do Tema", _
              U("Cor Prim\u00e1ria R"), U("Cor Prim\u00e1ria G"), U("Cor Prim\u00e1ria B"), _
              U("Cor Secund\u00e1ria R"), U("Cor Secund\u00e1ria G"), U("Cor Secund\u00e1ria B"), _
              "Fundo R", "Fundo G", "Fundo B", _
              "Cor do Texto R", "Cor do Texto G", "Cor do Texto B", "Fonte"), _
        Array("Corporativo Azul", 30, 64, 175, 217, 119, 6, 248, 250, 252, 15, 23, 42, "Arial"), _
        Array("Noite Escura", 56, 189, 248, 167, 139, 250, 17, 24, 39, 226, 232, 240, "Verdana"), _
        Array("Verde Floresta", 21, 128, 61, 202, 138, 4, 240, 253, 244, 20, 83, 45, "Georgia"), _
        Array("Grafite e Coral", 55, 65, 81, 239, 68, 68, 255, 255, 255, 17, 24, 39, "Calibri"))

    ReDim m(1 To UBound(linhas) + 1, 1 To NUM_COLS_TEMA)
    For i = 0 To UBound(linhas)
        For j = 0 To NUM_COLS_TEMA - 1
            m(i + 1, j + 1) = linhas(i)(j)
        Next j
    Next i
    TemasPadrao = m
End Function

Private Function U(ByVal s As String) As String
    U = modUtil.U(s)
End Function

'------------------------------------------------------------------------------
' LerTemaAtivo
' Proposito: localizar em tblTemas a linha do Tema Ativo e converter em TemaDS.
' Parametros: wb    - livro alvo.
'             t     - (saida) tema resolvido.
'             aviso - (saida) texto de aviso quando houve recurso a padrao.
' Retorno:    True sempre que conseguiu produzir um tema utilizavel.
' Efeitos:    nenhum (somente leitura).
' Regras:     tema nao encontrado -> primeiro tema da tabela; tabela vazia ou
'             ausente -> tema embutido "Corporativo Azul"; componente RGB
'             invalido -> limitado a 0..255 (texto vira 0).
'------------------------------------------------------------------------------
Public Function LerTemaAtivo(ByVal wb As Workbook, ByRef t As TemaDS, ByRef aviso As String) As Boolean
    Dim ws As Worksheet, lo As ListObject, v As Variant, nomeAtivo As String
    Dim i As Long, linha As Long

    Set ws = modUtil.ObterPlanilha(wb, PLAN_CONFIG)
    If Not ws Is Nothing Then Set lo = modUtil.ObterTabela(ws, TAB_TEMAS)

    If lo Is Nothing Then
        v = TemasPadrao()
        linha = 2
        aviso = modUtil.U("Tabela de temas n\u00e3o encontrada; usando o tema padr\u00e3o.")
    ElseIf lo.DataBodyRange Is Nothing Then
        v = TemasPadrao()
        linha = 2
        aviso = modUtil.U("Tabela de temas vazia; usando o tema padr\u00e3o.")
    Else
        v = lo.Range.Value       ' inclui o cabecalho na linha 1
        On Error Resume Next
        nomeAtivo = modUtil.TextoSeguro(wb.Names(NOME_TEMA_ATIVO).RefersToRange.Value)
        On Error GoTo 0
        For i = 2 To UBound(v, 1)
            If StrComp(modUtil.TextoSeguro(v(i, COL_NOME)), nomeAtivo, vbTextCompare) = 0 Then
                linha = i
                Exit For
            End If
        Next i
        If linha = 0 Then
            linha = 2
            aviso = modUtil.U("Tema """) & nomeAtivo & modUtil.U(""" n\u00e3o encontrado; usando """) & _
                    modUtil.TextoSeguro(v(2, COL_NOME)) & """."
        End If
    End If

    If UBound(v, 2) < NUM_COLS_TEMA Then
        v = TemasPadrao()
        linha = 2
        aviso = modUtil.U("A tabela de temas tem menos de 14 colunas; usando o tema padr\u00e3o.")
    End If

    t.Nome = modUtil.TextoSeguro(v(linha, COL_NOME))
    t.Primaria = modUtil.CorRGB(v(linha, COL_PRIMARIA), v(linha, COL_PRIMARIA + 1), v(linha, COL_PRIMARIA + 2))
    t.Secundaria = modUtil.CorRGB(v(linha, COL_SECUNDARIA), v(linha, COL_SECUNDARIA + 1), v(linha, COL_SECUNDARIA + 2))
    t.Fundo = modUtil.CorRGB(v(linha, COL_FUNDO), v(linha, COL_FUNDO + 1), v(linha, COL_FUNDO + 2))
    t.Texto = modUtil.CorRGB(v(linha, COL_TEXTO), v(linha, COL_TEXTO + 1), v(linha, COL_TEXTO + 2))
    t.Fonte = modUtil.TextoSeguro(v(linha, COL_FONTE))
    If Len(t.Fonte) = 0 Then t.Fonte = "Arial"
    LerTemaAtivo = True
End Function
