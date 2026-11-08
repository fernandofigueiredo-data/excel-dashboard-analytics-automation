Attribute VB_Name = "modTema"
Option Explicit
'==============================================================================
' modTema
' Design System: le o Tema Ativo em Config, cria ou atualiza os Estilos de
' Celula DS_* e recolore Shapes e graficos DS_* com cores RGB explicitas.
'
' Estilos (todos com IncludeNumber = False, para preservar formatos numericos):
'   DS_Texto      fundo e texto padrao (Dashboard inteira e corpo de tblDados)
'   DS_Header     cabecalho de tabela (cor primaria)
'   DS_Card       valor do card de KPI
'   DS_CardRotulo rotulo do card (fonte menor; necessario para hierarquia)
'   DS_Titulo     titulo do Dashboard (necessario para destaque)
'   DS_Aviso      mensagens ao usuario (necessario: quebra de linha e destaque)
' Como as celulas guardam o NOME do estilo, redefinir o estilo atualiza todas
' as celulas na hora, sem percorrer celula a celula.
'==============================================================================

'------------------------------------------------------------------------------
' AplicarTema
' Proposito: aplicar o Tema Ativo a todo o livro.
' Parametros: wb - livro alvo (opcional; padrao = livro ativo).
' Efeitos:    redefine os estilos DS_*, aplica-os a tblDados e a tblTemas e
'             recolore Shapes e graficos cujo nome comeca com "DS_" no
'             Dashboard. Restaura ScreenUpdating/Calculation mesmo apos erro.
'------------------------------------------------------------------------------
Public Sub AplicarTema(Optional ByVal wb As Workbook)
    Dim t As TemaDS, aviso As String, lo As ListObject, ws As Worksheet

    On Error GoTo Falha
    If wb Is Nothing Then Set wb = modUtil.LivroAlvo()
    If wb Is Nothing Then Exit Sub
    If modUtil.ObterPlanilha(wb, modConfig.PLAN_CONFIG) Is Nothing Then
        MsgBox modUtil.U("Este livro ainda n\u00e3o tem Dashboard. Clique em Atualizar primeiro."), _
               vbInformation, "Dashboard"
        Exit Sub
    End If

    modUtil.DesligarApp
    modConfig.LerTemaAtivo wb, t, aviso
    DefinirEstilos wb, t

    Set lo = modUtil.ObterTabelaNoLivro(wb, modConfig.TAB_DADOS)
    If Not lo Is Nothing Then AplicarEstilosTabela lo
    Set lo = modUtil.ObterTabelaNoLivro(wb, modConfig.TAB_TEMAS)
    If Not lo Is Nothing Then AplicarEstilosTabela lo

    Set ws = modUtil.ObterPlanilha(wb, modConfig.PLAN_DASH)
    If Not ws Is Nothing Then
        ColorirObjetos ws, t
        ws.Tab.Color = t.Primaria
    End If

    modUtil.RestaurarApp
    If Len(aviso) > 0 Then Application.StatusBar = aviso   ' depois de restaurar
    Exit Sub

Falha:
    Dim d As String
    d = Err.Number & " - " & Err.Description
    modUtil.RestaurarApp
    MsgBox modUtil.U("N\u00e3o foi poss\u00edvel aplicar o tema." & vbLf & vbLf & "Detalhe: ") & d, _
           vbExclamation, "Dashboard"
End Sub

'------------------------------------------------------------------------------
' TratarMudancaConfig
' Proposito: ponto unico chamado pelo evento Worksheet_Change de Config (modelo
'            .xlsm) ou pelo evento de aplicacao clsEventosApp (suplemento .xlam).
' Parametros: Sh - planilha alterada; Target - celulas alteradas.
' Efeitos:    chama AplicarTema se mudou o Tema Ativo ou a tabela de temas.
'------------------------------------------------------------------------------
Public Sub TratarMudancaConfig(ByVal Sh As Object, ByVal Target As Range)
    Dim wb As Workbook, alvo As Range, lo As ListObject

    On Error GoTo Saida
    If TypeName(Sh) <> "Worksheet" Then Exit Sub
    If StrComp(Sh.Name, modConfig.PLAN_CONFIG, vbTextCompare) <> 0 Then Exit Sub
    Set wb = Sh.Parent

    Set alvo = wb.Names(modConfig.NOME_TEMA_ATIVO).RefersToRange
    Set lo = modUtil.ObterTabela(Sh, modConfig.TAB_TEMAS)
    If Not lo Is Nothing Then Set alvo = Application.Union(alvo, lo.Range)

    If Not Application.Intersect(Target, alvo) Is Nothing Then AplicarTema wb
Saida:
End Sub

'------------------------------------------------------------------------------
' DefinirEstilos
' Proposito: criar (se faltar) e redefinir os estilos DS_* com o tema.
' Parametros: wb - livro alvo; t - tema resolvido.
' Efeitos:    altera Workbook.Styles; todas as celulas com esses estilos mudam.
'------------------------------------------------------------------------------
Private Sub DefinirEstilos(ByVal wb As Workbook, ByRef t As TemaDS)
    Dim fundoCard As Long, textoClaro As Long

    fundoCard = modUtil.Misturar(t.Fundo, t.Primaria, 0.1)

    ConfigurarEstilo ObterEstilo(wb, "DS_Texto"), t.Fonte, 11, False, t.Texto, t.Fundo, xlLeft
    ConfigurarEstilo ObterEstilo(wb, "DS_Header"), t.Fonte, 11, True, _
                     modUtil.Contraste(t.Primaria), t.Primaria, xlLeft
    ConfigurarEstilo ObterEstilo(wb, "DS_Titulo"), t.Fonte, 20, True, _
                     modUtil.CorLegivel(t.Primaria, t.Fundo, t.Texto), t.Fundo, xlLeft
    ConfigurarEstilo ObterEstilo(wb, "DS_CardRotulo"), t.Fonte, 10, False, t.Texto, fundoCard, xlCenter
    ConfigurarEstilo ObterEstilo(wb, "DS_Card"), t.Fonte, 20, True, _
                     modUtil.CorLegivel(t.Primaria, fundoCard, t.Texto), fundoCard, xlCenter
    ConfigurarEstilo ObterEstilo(wb, "DS_Aviso"), t.Fonte, 10, False, t.Texto, _
                     modUtil.Misturar(t.Fundo, t.Secundaria, 0.18), xlLeft
    wb.Styles("DS_Aviso").WrapText = True
    wb.Styles("DS_Aviso").Font.Italic = True
    wb.Styles("DS_Card").ShrinkToFit = True          ' numeros grandes nao viram ####
    wb.Styles("DS_CardRotulo").ShrinkToFit = True    ' rotulos longos cabem no card

    ' Bordas: linha de destaque no cabecalho e barra lateral nos cards
    textoClaro = modUtil.Misturar(t.Primaria, t.Texto, 0.3)
    BordaEstilo wb.Styles("DS_Header"), xlEdgeBottom, xlBottom, textoClaro, xlMedium
    BordaEstilo wb.Styles("DS_Card"), xlEdgeLeft, xlLeft, t.Secundaria, xlThick
    BordaEstilo wb.Styles("DS_CardRotulo"), xlEdgeLeft, xlLeft, t.Secundaria, xlThick
End Sub

' Devolve o estilo existente ou cria um novo com o nome dado.
Private Function ObterEstilo(ByVal wb As Workbook, ByVal nome As String) As Style
    On Error Resume Next
    Set ObterEstilo = wb.Styles(nome)
    On Error GoTo 0
    If ObterEstilo Is Nothing Then Set ObterEstilo = wb.Styles.Add(nome)
End Function

'------------------------------------------------------------------------------
' ConfigurarEstilo
' Proposito: definir fonte, cores e alinhamento de um estilo.
' Parametros: s - estilo; fonte, tamanho, negrito; corTexto, corFundo (RGB);
'             horiz - alinhamento horizontal (xlLeft, xlCenter...).
' Efeitos:    IncludeNumber = False para nao apagar formatos numericos.
'------------------------------------------------------------------------------
Private Sub ConfigurarEstilo(ByVal s As Style, ByVal fonte As String, ByVal tamanho As Double, _
        ByVal negrito As Boolean, ByVal corTexto As Long, ByVal corFundo As Long, ByVal horiz As Long)
    s.IncludeNumber = False
    s.IncludeFont = True
    s.IncludeAlignment = True
    s.IncludeBorder = True
    s.IncludePatterns = True
    s.IncludeProtection = False
    s.Font.Name = fonte
    s.Font.Size = tamanho
    s.Font.Bold = negrito
    s.Font.Italic = False
    s.Font.Color = corTexto
    s.Interior.Pattern = xlSolid
    s.Interior.Color = corFundo
    s.HorizontalAlignment = horiz
    s.VerticalAlignment = xlCenter
    s.WrapText = False
End Sub

'------------------------------------------------------------------------------
' BordaEstilo
' Proposito: definir uma borda de estilo. Style.Borders aceita xlLeft/xlBottom
'            em algumas versoes e xlEdgeLeft/xlEdgeBottom em outras; tenta as duas.
' Parametros: s - estilo; idx1, idx2 - indices alternativos; cor (RGB); peso.
' Efeitos:    erro de borda nao interrompe (borda e cosmetica).
'------------------------------------------------------------------------------
Private Sub BordaEstilo(ByVal s As Style, ByVal idx1 As Long, ByVal idx2 As Long, _
        ByVal cor As Long, ByVal peso As Long)
    On Error Resume Next
    With s.Borders(idx1)
        .LineStyle = xlContinuous
        .Weight = peso
        .Color = cor
    End With
    If Err.Number <> 0 Then
        Err.Clear
        With s.Borders(idx2)
            .LineStyle = xlContinuous
            .Weight = peso
            .Color = cor
        End With
    End If
    On Error GoTo 0
End Sub

' Aplica DS_Header ao cabecalho e DS_Texto ao corpo de uma tabela (2 operacoes).
Private Sub AplicarEstilosTabela(ByVal lo As ListObject)
    If Not lo.HeaderRowRange Is Nothing Then lo.HeaderRowRange.Style = "DS_Header"
    If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.Style = "DS_Texto"
End Sub

'------------------------------------------------------------------------------
' ColorirObjetos
' Proposito: aplicar o tema a Shapes (botoes) e graficos com nome "DS_*".
' Parametros: ws - planilha Dashboard; t - tema.
' Efeitos:    altera preenchimento, linhas e fontes dos objetos.
'------------------------------------------------------------------------------
Public Sub ColorirObjetos(ByVal ws As Worksheet, ByRef t As TemaDS)
    Dim shp As Shape, co As ChartObject

    For Each shp In ws.Shapes
        If Left$(shp.Name, 3) = "DS_" And shp.Type <> msoChart Then ColorirShape shp, t
    Next shp
    For Each co In ws.ChartObjects
        If Left$(co.Name, 3) = "DS_" Then modGraficos.EstilizarGrafico co, t
    Next co
End Sub

' Botao: fundo primario, texto em cor de contraste, fonte do tema.
Private Sub ColorirShape(ByVal shp As Shape, ByRef t As TemaDS)
    On Error Resume Next        ' propriedades cosmeticas; falha nao interrompe
    shp.Fill.Visible = msoTrue
    shp.Fill.Solid
    shp.Fill.ForeColor.RGB = t.Primaria
    shp.Line.Visible = msoFalse
    With shp.TextFrame2.TextRange.Font
        .Name = t.Fonte
        .Size = 12
        .Bold = msoTrue
        .Fill.ForeColor.RGB = modUtil.Contraste(t.Primaria)
    End With
    On Error GoTo 0
End Sub

