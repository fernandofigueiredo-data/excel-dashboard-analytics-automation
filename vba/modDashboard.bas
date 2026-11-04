Attribute VB_Name = "modDashboard"
Option Explicit
'==============================================================================
' modDashboard
' Rotina principal AtualizarTudo e montagem da planilha Dashboard.
'
' REGRA DE CLASSIFICACAO (implementada em modDados.ClassificarColunas)
'   Amostra de ate 500 celulas nao vazias por coluna, nas 2.000 primeiras
'   linhas; >= 80% datas -> data; >= 80% numeros -> numerica (salvo
'   identificadores e codigos de classe); nenhuma celula valida -> vazia
'   (ignorada); caso contrario -> categorica. Detalhes no topo de modDados.
'
' GERACAO PADRAO
'   - card "Registros"        =ROWS(tblDados)
'   - cards Soma e Media      ate as 3 primeiras colunas numericas
'                             (AGGREGATE ignora celulas com erro)
'   - graficos                8 tipos automaticos + linha no tempo, com as
'                             regras do guia de graficos: ver modGraficos.
'
' NENHUM numero e digitado: cards e area auxiliar sao formulas. Textos de
' rotulo e avisos sao escritos pelo VBA.
'==============================================================================

Private Const CARDS_POR_LINHA As Long = 4
Private Const MAX_NUMERICAS As Long = 3
Private Const TOP_N As Long = 10
Private Const LIN_AVISO As Long = 3
Private Const LIN_CARDS As Long = 5
Private Const COL_INI As Long = 2          ' coluna B
Private Const COL_FIM As Long = 16         ' coluna P (ultima do painel)
Private Const COL_AUX As Long = 19         ' coluna S (area auxiliar)
Private Const LIN_AUX As Long = 5          ' primeira linha de resultados auxiliares

'------------------------------------------------------------------------------
' AtualizarTudo  (botao "Atualizar")
' Proposito: criar/ajustar a estrutura e tblDados, aplicar o tema e reconstruir
'            o Dashboard.
' Parametros: nenhum (atua no livro ativo; ver modUtil.LivroAlvo).
' Efeitos:    pode criar planilhas, reescrever cabecalhos de Dados, recriar todo
'             o conteudo da planilha Dashboard (celulas, graficos e Shapes DS_*)
'             e os nomes dsG_*. ScreenUpdating, Calculation e eventos sao
'             desligados e restaurados, tambem em caso de erro.
'------------------------------------------------------------------------------
Public Sub AtualizarTudo()
    Dim wb As Workbook, lo As ListObject, msg As String, temDados As Boolean, d As String

    On Error GoTo Falha
    Set wb = modUtil.LivroAlvo()
    If wb Is Nothing Then
        MsgBox modUtil.U("Abra o livro com os dados antes de clicar em Atualizar."), vbInformation, "Dashboard"
        Exit Sub
    End If

    modUtil.DesligarApp
    Application.StatusBar = "Dashboard: preparando a estrutura..."
    modConfig.GarantirEstrutura wb

    Application.StatusBar = "Dashboard: preparando tblDados..."
    temDados = modDados.PrepararTabela(wb, lo, msg)

    modTema.AplicarTema wb                  ' cria os estilos antes de usa-los
    Application.StatusBar = "Dashboard: montando cards e graficos..."
    ConstruirDashboard wb, lo, temDados, msg
    modTema.AplicarTema wb                  ' colore os graficos recem-criados
    modGraficos.FinalizarFormatos wb        ' por ultimo: o Excel para Mac pode refazer o formato

    modUtil.RestaurarApp
    Application.StatusBar = modUtil.U("Dashboard atualizado em ") & Format$(Now, "hh:nn:ss")
    Exit Sub

Falha:
    d = Err.Number & " - " & Err.Description
    modUtil.RestaurarApp
    EscreverFalha wb, d
    MsgBox modUtil.U("N\u00e3o foi poss\u00edvel atualizar o dashboard." & vbLf & vbLf & "Detalhe: ") & d, _
           vbExclamation, "Dashboard"
End Sub

' Callbacks da Faixa de Opcoes (usadas apenas no suplemento .xlam).
Public Sub Faixa_AtualizarTudo(control As IRibbonControl)
    AtualizarTudo
End Sub

Public Sub Faixa_AplicarTema(control As IRibbonControl)
    modTema.AplicarTema
End Sub

'------------------------------------------------------------------------------
' ConstruirDashboard
' Proposito: limpar e montar o Dashboard: titulo, avisos, cards e graficos.
' Parametros: wb - livro; lo - tblDados (Nothing se nao ha dados);
'             temDados - resultado de PrepararTabela; msgDados - motivo se False.
' Efeitos:    apaga todo o conteudo anterior da planilha Dashboard.
'------------------------------------------------------------------------------
Private Sub ConstruirDashboard(ByVal wb As Workbook, ByVal lo As ListObject, _
        ByVal temDados As Boolean, ByVal msgDados As String)
    Dim ws As Worksheet, tipos As Variant, avisos As Collection, nums As Collection
    Dim iData As Long, iNum As Long, nVazias As Long, nIds As Long, j As Long, k As Long
    Dim nomeCol As String, ref As String, fmt As String, linGraf As Long
    Dim esq As Double, larg As Double, topo As Double

    Set ws = modUtil.ObterPlanilha(wb, modConfig.PLAN_DASH)
    LimparDashboard wb, ws
    PrepararLayout ws
    Set avisos = New Collection

    If Not temDados Then
        avisos.Add msgDados
        EscreverAvisos ws, avisos
        CriarBotao ws
        FinalizarVisual ws
        Exit Sub
    End If

    ' Classificacao das colunas
    tipos = modDados.ClassificarColunas(lo)
    Set nums = New Collection
    For j = 1 To UBound(tipos)
        Select Case tipos(j)
            Case "N"
                If nums.Count < MAX_NUMERICAS Then nums.Add j
            Case "D"
                If iData = 0 Then iData = j
            Case "V"
                nVazias = nVazias + 1
            Case "I"
                nIds = nIds + 1
        End Select
    Next j
    If nums.Count > 0 Then iNum = nums(1)

    ws.Range("B2").Value = modUtil.U("Base: tblDados (planilha Dados) \u00b7 ") & lo.ListColumns.Count & _
        modUtil.U(" colunas \u00b7 atualizado em ") & Format$(Now, "dd/mm/yyyy hh:nn")

    ' Cards
    k = 0
    AdicionarCard ws, k, "Registros", "=ROWS(" & modConfig.TAB_DADOS & ")", "#,##0"
    k = k + 1
    For j = 1 To nums.Count
        nomeCol = lo.ListColumns(nums(j)).Name
        ref = modUtil.RefColuna(modConfig.TAB_DADOS, nomeCol)
        fmt = modGraficos.FormatoColuna(nums(j), False, lo)
        AdicionarCard ws, k, "Soma de " & nomeCol, "=IFERROR(AGGREGATE(9,6," & ref & "),0)", fmt
        k = k + 1
        fmt = modGraficos.FormatoColuna(nums(j), True, lo)
        AdicionarCard ws, k, modUtil.U("M\u00e9dia de ") & nomeCol, "=IFERROR(AGGREGATE(1,6," & ref & "),""-"")", fmt
        k = k + 1
    Next j

    ' Graficos abaixo dos cards
    linGraf = LIN_CARDS + (((k - 1) \ CARDS_POR_LINHA) + 1) * 4
    esq = ws.Columns(COL_INI).Left
    larg = (ws.Columns(COL_FIM + 1).Left - esq - 12) / 2
    topo = ws.Rows(linGraf).Top

    modGraficos.GerarGraficos wb, ws, lo, tipos, nums, iData, avisos, esq, larg, topo, COL_AUX

    If iNum = 0 Then
        avisos.Add modUtil.U("Nenhuma coluna num\u00e9rica: n\u00e3o h\u00e1 cards de soma e m\u00e9dia, " & _
                   "e os gr\u00e1ficos mostram a contagem de registros.")
    End If
    If nVazias > 0 Then
        avisos.Add nVazias & modUtil.U(" coluna(s) sem valores foram ignoradas.")
    End If
    If nIds > 0 Then
        avisos.Add nIds & modUtil.U(" coluna(s) de c\u00f3digo/identificador n\u00e3o entram em somas e m\u00e9dias.")
    End If

    EscreverAvisos ws, avisos
    CriarBotao ws
    FinalizarVisual ws
End Sub

'------------------------------------------------------------------------------
' LimparDashboard
' Proposito: remover tudo o que a montagem anterior criou.
' Parametros: wb - livro; ws - planilha Dashboard.
' Efeitos:    apaga os graficos gerados ("DS_G*"), Shapes "DS_*", nomes "dsG_*"
'             e todo o conteudo/formatacao das celulas. Graficos com outro nome
'             (feitos pelo usuario) ficam onde estao.
'------------------------------------------------------------------------------
Private Sub LimparDashboard(ByVal wb As Workbook, ByVal ws As Worksheet)
    Dim i As Long
    For i = ws.ChartObjects.Count To 1 Step -1
        If Left$(ws.ChartObjects(i).Name, 4) = "DS_G" Then ws.ChartObjects(i).Delete
    Next i
    For i = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(i).Name, 3) = "DS_" Then ws.Shapes(i).Delete
    Next i
    For i = wb.Names.Count To 1 Step -1
        If Left$(wb.Names(i).Name, 4) = "dsG_" Then wb.Names(i).Delete
    Next i
    ws.Cells.Clear
    ws.Rows.RowHeight = ws.StandardHeight
End Sub

'------------------------------------------------------------------------------
' PrepararLayout
' Proposito: grade de colunas, fundo, titulo e cabecalho da area auxiliar.
' Parametros: ws - planilha Dashboard.
' Efeitos:    aplica DS_Texto a planilha inteira (uma operacao) e DS_Titulo.
'------------------------------------------------------------------------------
Private Sub PrepararLayout(ByVal ws As Worksheet)
    ws.Cells.Style = "DS_Texto"
    ws.Columns(1).ColumnWidth = 2
    ws.Range(ws.Columns(COL_INI), ws.Columns(COL_FIM)).ColumnWidth = 11
    ws.Range(ws.Columns(COL_FIM + 1), ws.Columns(COL_AUX - 1)).ColumnWidth = 3
    ws.Range(ws.Columns(COL_AUX), ws.Columns(COL_AUX + 4)).ColumnWidth = 18

    ws.Range(ws.Cells(1, COL_INI), ws.Cells(1, COL_FIM)).Style = "DS_Titulo"
    ws.Cells(1, COL_INI).Value = "Dashboard"
    ws.Rows(1).RowHeight = 36

    ws.Cells(2, COL_AUX).Value = modUtil.U("\u00c1rea auxiliar dos gr\u00e1ficos (gerada automaticamente; n\u00e3o editar)")
End Sub

'------------------------------------------------------------------------------
' AdicionarCard
' Proposito: desenhar um card de KPI como bloco de celulas mescladas.
' Parametros: ws; k - posicao do card (0, 1, 2...); rotulo; formula (ingles,
'             atribuida por Formula2); fmt - formato numerico do valor.
' Efeitos:    mescla 1x3 (rotulo, DS_CardRotulo) e 2x3 (valor, DS_Card).
'------------------------------------------------------------------------------
Private Sub AdicionarCard(ByVal ws As Worksheet, ByVal k As Long, ByVal rotulo As String, _
        ByVal formula As String, ByVal fmt As String)
    Dim lin As Long, col As Long

    lin = LIN_CARDS + (k \ CARDS_POR_LINHA) * 4
    col = COL_INI + (k Mod CARDS_POR_LINHA) * 4

    With ws.Cells(lin, col).Resize(1, 3)
        .Merge
        .Style = "DS_CardRotulo"
        .Cells(1, 1).Value = rotulo
    End With
    With ws.Cells(lin + 1, col).Resize(2, 3)
        .Merge
        .Style = "DS_Card"
        .NumberFormat = fmt
        .Cells(1, 1).Formula2 = formula
    End With
    ws.Rows(lin).RowHeight = 20
    ws.Rows(lin + 1).RowHeight = 22
    ws.Rows(lin + 2).RowHeight = 22
End Sub

'------------------------------------------------------------------------------
' EscreverAvisos
' Proposito: mostrar mensagens ao usuario na linha 3 do Dashboard.
' Parametros: ws; avisos - Collection de textos (vazia = nada a mostrar).
' Efeitos:    mescla B3:P3 com estilo DS_Aviso e ajusta a altura da linha.
'------------------------------------------------------------------------------
Private Sub EscreverAvisos(ByVal ws As Worksheet, ByVal avisos As Collection)
    Dim i As Long, txt As String
    If avisos.Count = 0 Then Exit Sub
    For i = 1 To avisos.Count
        If i > 1 Then txt = txt & vbLf
        txt = txt & "- " & avisos(i)
    Next i
    With ws.Range(ws.Cells(LIN_AVISO, COL_INI), ws.Cells(LIN_AVISO, COL_FIM))
        .Merge
        .Style = "DS_Aviso"
        .Cells(1, 1).Value = txt
    End With
    ws.Rows(LIN_AVISO).RowHeight = 16 * avisos.Count + 8
End Sub

' Botoes "Atualizar" e "Fixar grafico" (Shapes nativas com OnAction; sem Form
' Controls nem ActiveX).
Private Sub CriarBotao(ByVal ws As Worksheet)
    NovoBotao ws, "DS_BtnAtualizar", "Atualizar", "AtualizarTudo", COL_FIM - 2, COL_FIM + 1
    NovoBotao ws, "DS_BtnFixar", modUtil.U("Fixar gr\u00e1fico"), "FixarGrafico", COL_FIM - 6, COL_FIM - 2
End Sub

' Um botao na linha 1, da coluna c1 ate a borda esquerda da coluna c2.
Private Sub NovoBotao(ByVal ws As Worksheet, ByVal nome As String, ByVal rotulo As String, _
        ByVal macro As String, ByVal c1 As Long, ByVal c2 As Long)
    Dim shp As Shape, esq As Double
    esq = ws.Columns(c1).Left
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, esq, ws.Rows(1).Top + 5, _
                                 ws.Columns(c2).Left - esq - 6, 26)
    shp.Name = nome
    shp.OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & macro
    shp.TextFrame2.TextRange.Text = rotulo
    On Error Resume Next
    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
    shp.Placement = xlMove
    On Error GoTo 0
End Sub

' Ativa o Dashboard, esconde as linhas de grade e posiciona em A1.
Private Sub FinalizarVisual(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ws.Range("A1").Select
    On Error GoTo 0
End Sub

' Em caso de erro, deixa a mensagem visivel no proprio Dashboard.
Private Sub EscreverFalha(ByVal wb As Workbook, ByVal detalhe As String)
    Dim ws As Worksheet
    On Error Resume Next
    If wb Is Nothing Then Exit Sub
    Set ws = modUtil.ObterPlanilha(wb, modConfig.PLAN_DASH)
    If ws Is Nothing Then Exit Sub
    ws.Cells(LIN_AVISO, COL_INI).Value = modUtil.U("Erro ao montar o dashboard: ") & detalhe
End Sub
