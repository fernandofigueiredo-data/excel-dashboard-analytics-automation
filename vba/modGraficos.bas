Attribute VB_Name = "modGraficos"
Option Explicit
Option Private Module   ' rotinas internas: fora da lista de macros
'==============================================================================
' modGraficos
' Gera automaticamente 8 tipos de grafico (mais a linha no tempo) seguindo o
' "Guia de Graficos no Excel para Analise de Dados".
'
' REGRAS DO GUIA APLICADAS A TODOS
'   - O titulo diz a conclusao: formula na area auxiliar, ligada ao titulo.
'   - Barras e colunas comecam em zero; no maximo 7 categorias (as 7 maiores,
'     e o titulo avisa quantas existem).
'   - Nada de 3D; menos ruido: serie unica sem legenda e sem linhas de grade,
'     com rotulos de dados diretos.
'   - Uma cor de destaque (a maior categoria) e cinza para o contexto.
'   - Fonte de 12 pt.
'
' QUAL COLUNA ALIMENTA CADA GRAFICO
'   catA   = 1a coluna categorica
'   catB   = 2a coluna categorica (sem ela, catA)
'   partes = categorica diferente de catA com 2 a 4 grupos (texto antes de
'            codigo numerico); sem ela, a de menos grupos (3 maiores + Outros)
'   medida = 1a coluna numerica (sem numerica: contagem de registros)
'
' REGRAS DE CADA GRAFICO (gerado so quando os dados cumprem a regra)
'   Colunas agrupadas  catA x 1 a 3 medidas DA MESMA ESCALA (total entre 1/10
'                      e 10 vezes o da 1a). Escalas diferentes nao se misturam.
'   Barras agrupadas   catB x as mesmas medidas, horizontais, maior no topo.
'   Colunas empilhadas catA x partes: ate 4 partes (3 maiores + "Outros"),
'                      do tom escuro ao claro.
'   Barras empilhadas  mesmos dados em 100% (compara so a composicao).
'   Pizza              categorica com 2 a 5 grupos em que uma fatia tem
'                      50% ou mais (e diferente da usada na rosca).
'   Rosca              categorica com 2 a 5 grupos (texto antes de codigo
'                      numerico); furo de 60%; maior
'                      fatia as 12 h; rotulos com nome e porcentagem.
'   Dispersao (XY)     1a x 2a coluna numerica, 20 pontos ou mais, linha de
'                      tendencia linear e titulos dos eixos.
'   Arvore hierarquica catA > partes (sem partes, so catA). Nao esta no guia:
'                      segue as regras gerais.
'   Linha              soma mensal da medida pela 1a coluna de data.
'
' Os graficos leem a area auxiliar (formulas de matriz dinamica) por nomes
' definidos dsG_* com referencia de derramamento; a arvore le o intervalo
' derramado atual (o Excel nao aceita nome nesse tipo de grafico).
'==============================================================================

Private Const MAX_CAT As Long = 7
Private Const MAX_PARTES As Long = 4
Private Const MAX_FATIAS As Long = 5
Private Const MIN_PONTOS As Long = 20
Private Const DOMINANCIA As Double = 0.5
Private Const ESCALA As Double = 10
Private Const ALTURA As Double = 280
Private Const XL_TREEMAP As Long = 117

Private mWb As Workbook
Private mWs As Worksheet
Private mLo As ListObject
Private mAvisos As Collection
Private mMeds As Collection          ' medidas agrupaveis (indices de coluna)
Private mFmts As Collection          ' formato de eixo por grafico (FinalizarFormatos)
Private mINum As Long                ' 1a medida (0 = contagem)
Private mMedida As String            ' nome da medida para os titulos
Private mCol As Long                 ' proxima coluna livre da area auxiliar
Private mN As Long                   ' graficos ja posicionados
Private mEsq As Double, mLarg As Double, mTopo As Double

'------------------------------------------------------------------------------
' GerarGraficos
' Proposito: escolher as colunas de cada grafico, escrever a area auxiliar e
'            criar os graficos em grade de 2 por linha.
' Parametros: wb, ws (Dashboard), lo (tblDados); tipos - classificacao das
'             colunas; nums - indices numericos; iData - coluna de data (0 =
'             nenhuma); avisos - recebe os motivos de grafico nao gerado;
'             esq, larg, topo - grade; colAux - 1a coluna da area auxiliar.
' Retorno:    quantidade de graficos criados.
'------------------------------------------------------------------------------
Public Function GerarGraficos(ByVal wb As Workbook, ByVal ws As Worksheet, ByVal lo As ListObject, _
        ByVal tipos As Variant, ByVal nums As Collection, ByVal iData As Long, ByVal avisos As Collection, _
        ByVal esq As Double, ByVal larg As Double, ByVal topo As Double, ByVal colAux As Long) As Long
    Dim cats As Collection, dist As Collection, j As Long, i As Long
    Dim catA As Long, catB As Long, partes As Long, nPartes As Long
    Dim dimRosca As Long, dimPizza As Long, nDist As Long
    Dim ehCod As Collection, grupos As Collection, nota As Long, melhor As Long

    Set mWb = wb: Set mWs = ws: Set mLo = lo: Set mAvisos = avisos
    Set mFmts = New Collection
    mCol = colAux: mN = 0: mEsq = esq: mLarg = larg: mTopo = topo

    mINum = 0
    If nums.Count > 0 Then mINum = nums(1)
    If mINum > 0 Then mMedida = NomeCol(mINum) Else mMedida = "registros"
    Set mMeds = New Collection
    If mINum > 0 Then
        mMeds.Add mINum
        For i = 2 To nums.Count
            If MesmaEscala(nums(1), nums(i)) Then mMeds.Add nums(i)
        Next i
    End If

    ' Categoricas e quantidade de grupos (contagem exata ate 50)
    Set cats = New Collection: Set dist = New Collection: Set ehCod = New Collection
    For j = 1 To UBound(tipos)
        If tipos(j) = "C" Then
            Set grupos = modDados.ValoresDistintos(lo, j, 50)
            cats.Add j
            dist.Add grupos.Count
            ehCod.Add GruposNumericos(grupos)
        End If
    Next j
    If cats.Count > 0 Then catA = cats(1)
    If cats.Count > 1 Then catB = cats(2) Else catB = catA

    ' Partes: ate 4 grupos (guia); prefere texto a codigo numerico; sem isso,
    ' a categoria com menos grupos (3 maiores + "Outros").
    For i = 1 To cats.Count
        If cats(i) <> catA And dist(i) >= 2 Then
            nota = IIf(dist(i) <= MAX_PARTES, 0, 1000 + dist(i)) + IIf(ehCod(i), 100, 0) + i
            If partes = 0 Or nota < melhor Then partes = cats(i): nPartes = dist(i): melhor = nota
        End If
    Next i
    ' Rosca: 2 a 5 grupos, texto antes de codigo
    melhor = 0
    For i = 1 To cats.Count
        If dist(i) >= 2 And dist(i) <= MAX_FATIAS Then
            nota = IIf(ehCod(i), 100, 0) + i
            If dimRosca = 0 Or nota < melhor Then dimRosca = cats(i): melhor = nota
        End If
    Next i
    ' Pizza: 2 a 5 grupos com fatia dominante, diferente da rosca
    melhor = 0
    For i = 1 To cats.Count
        If cats(i) <> dimRosca And dist(i) >= 2 And dist(i) <= MAX_FATIAS Then
            nota = IIf(ehCod(i), 100, 0) + i
            If dimPizza = 0 Or nota < melhor Then
                If Dominante(cats(i)) Then dimPizza = cats(i): melhor = nota
            End If
        End If
    Next i

    ' 1-2: colunas e barras agrupadas
    If catA > 0 Then
        GraficoRanking "ColunasAgrupadas", catA, False
        GraficoRanking "BarrasAgrupadas", catB, True
    Else
        Aviso U("Nenhuma coluna de categoria: colunas, barras, empilhadas e \u00e1rvore n\u00e3o foram geradas.")
    End If

    ' 3-4 e 8: empilhadas e arvore
    If catA > 0 And partes > 0 Then
        GraficosComposicao catA, partes, nPartes
    ElseIf catA > 0 Then
        Aviso U("Colunas e barras empilhadas precisam de uma 2\u00aa coluna de categoria com 2 ou mais grupos.")
        GraficoArvoreSimples
    End If

    ' 5-6: rosca e pizza
    If dimRosca > 0 Then
        GraficoFatias "Rosca", dimRosca
    Else
        Aviso U("Rosca n\u00e3o gerada: nenhuma categoria tem de 2 a 5 grupos (o guia limita a 5 fatias).")
    End If
    If dimPizza > 0 Then
        GraficoFatias "Pizza", dimPizza
    Else
        Aviso U("Pizza n\u00e3o gerada: o guia pede pizza s\u00f3 quando uma fatia domina (50% ou mais) " & _
                "em uma categoria de at\u00e9 5 grupos, e nenhuma outra categoria cumpre isso.")
    End If

    ' 7: dispersao
    If nums.Count >= 2 Then
        GraficoDispersao nums(1), nums(2)
    Else
        Aviso U("Dispers\u00e3o n\u00e3o gerada: precisa de duas colunas num\u00e9ricas.")
    End If

    ' Linha no tempo
    If iData > 0 Then
        GraficoLinha iData
    Else
        Aviso U("Nenhuma coluna de data: o gr\u00e1fico de linha n\u00e3o foi gerado.")
    End If

    GerarGraficos = mN
End Function

'==============================================================================
' Blocos de cada grafico
'==============================================================================

' Colunas ou barras agrupadas: categoria x 1 a 3 medidas, 7 maiores.
Private Sub GraficoRanking(ByVal kind As String, ByVal jCat As Long, ByVal horizontal As Boolean)
    Dim col As Long, cCat As Range, i As Long, nm As Long, co As ChartObject, s As Series

    nm = mMeds.Count
    If nm = 0 Then nm = 1
    col = Bloco(nm + 1)
    Set cCat = Cel(5, col)
    cCat.Formula2 = "=IFERROR(" & GruposOrdenados(jCat, MAX_CAT) & ",""(vazio)"")"
    Cel(4, col).Value = NomeCol(jCat)
    For i = 1 To nm
        Cel(5, col + i).Formula2 = "=IFERROR(" & Agg(MedidaN(i), Ref(jCat) & "," & Sp(cCat)) & ",0)"
        Cel(4, col + i).Value = NomeMedidaN(i)
        FormatarValores col + i, MedidaN(i)
        NomeSpill "dsG_" & kind & "_V" & i, Cel(5, col + i)
    Next i
    EstiloCabecalho col, col + nm
    NomeSpill "dsG_" & kind & "_C", cCat

    Cel(3, col).Formula2 = "=IFERROR(INDEX(" & Sp(cCat) & ",1)&" & Q(" lidera em ") & "&" & Q(mMedida) & _
        "&"" (""&TEXT(INDEX(" & Sp(Cel(5, col + 1)) & ",1)/" & TotalMedida() & ",""0%"")&" & Q(" do total)") & _
        "&IF(" & NGrupos(jCat) & ">" & MAX_CAT & "," & Q(U(" \u00b7 ") & MAX_CAT & " maiores de ") & "&" & _
        NGrupos(jCat) & ",""""),"""")"
    mWs.Calculate

    Set co = NovoGrafico(kind)
    For i = 1 To nm
        Set s = co.Chart.SeriesCollection.NewSeries
        s.Name = NomeMedidaN(i)
        LigarSerie s, "dsG_" & kind & "_C", "dsG_" & kind & "_V" & i, cCat, Cel(5, col + i)
    Next i
    If horizontal Then co.Chart.ChartType = xlBarClustered Else co.Chart.ChartType = xlColumnClustered
    co.Chart.HasLegend = (nm > 1)
    TituloConclusao co.Chart, Cel(3, col), NomeMedidaN(1) & " por " & NomeCol(jCat)
    EixoDesdeZero co.Chart, Cel(5, col + 1)
    RegistrarFormato co.Name, CodigoFormato(mINum)
End Sub

' Colunas empilhadas, barras empilhadas 100% e arvore: categoria x partes.
Private Sub GraficosComposicao(ByVal jCat As Long, ByVal jPar As Long, ByVal nDistP As Long)
    Dim col As Long, cRows As Range, nNomeadas As Long, outros As Boolean, np As Long, i As Long
    Dim f As String, somas As String, linha1 As String, colArv As Long, cArv As Range, co As ChartObject

    outros = (nDistP > MAX_PARTES)
    If outros Then nNomeadas = MAX_PARTES - 1 Else nNomeadas = nDistP
    np = nNomeadas
    If outros Then np = np + 1

    col = Bloco(np + 1)
    Set cRows = Cel(5, col)
    cRows.Formula2 = "=IFERROR(" & GruposOrdenados(jCat, MAX_CAT) & ",""(vazio)"")"
    Cel(4, col).Value = NomeCol(jCat)
    NomeSpill "dsG_Emp_C", cRows
    For i = 1 To nNomeadas
        Cel(4, col + i).Formula2 = "=IFERROR(INDEX(" & GruposOrdenados(jPar, 0) & "," & i & "),"""")"
        Cel(5, col + i).Formula2 = "=IFERROR(" & Agg(mINum, Ref(jCat) & "," & Sp(cRows) & "," & Ref(jPar) & "," & _
            Cel(4, col + i).Address) & ",0)"
    Next i
    If outros Then
        Cel(4, col + np).Value = "Outros"
        f = "=IFERROR(" & Agg(mINum, Ref(jCat) & "," & Sp(cRows))
        For i = 1 To nNomeadas
            f = f & "-" & Sp(Cel(5, col + i))
        Next i
        Cel(5, col + np).Formula2 = f & ",0)"
    End If
    For i = 1 To np
        FormatarValores col + i, mINum
        NomeSpill "dsG_Emp_V" & i, Cel(5, col + i)
        If i > 1 Then somas = somas & ","
        If i > 1 Then linha1 = linha1 & "+"
        somas = somas & Sp(Cel(5, col + i))
        linha1 = linha1 & "INDEX(" & Sp(Cel(5, col + i)) & ",1)"
    Next i
    somas = "SUM(" & somas & ")"
    EstiloCabecalho col, col + np

    ' Titulos (linha 3): colunas empilhadas, barras 100% e arvore
    Cel(3, col).Formula2 = "=IFERROR(INDEX(" & Sp(cRows) & ",1)&" & Q(" tem o maior total de ") & "&" & Q(mMedida) & _
        "&""; ""&" & Cel(4, col + 1).Address & "&" & Q(U(" \u00e9 a maior parte (")) & "&TEXT(SUM(" & Sp(Cel(5, col + 1)) & _
        ")/" & somas & ",""0%"")&"")"","""")"
    Cel(3, col + 1).Formula2 = "=IFERROR(" & Cel(4, col + 1).Address & "&" & Q(" responde por ") & "&TEXT(SUM(" & _
        Sp(Cel(5, col + 1)) & ")/" & somas & ",""0%"")&" & Q(" de ") & "&" & Q(mMedida) & "&" & Q(" nos ") & _
        "&ROWS(" & Sp(cRows) & ")&" & Q(" maiores grupos de " & NomeCol(jCat)) & ","""")"
    Cel(3, col + 2).Formula2 = "=IFERROR(INDEX(" & Sp(cRows) & ",1)&" & Q(" concentra ") & "&TEXT((" & linha1 & ")/" & _
        somas & ",""0%"")&" & Q(" de " & mMedida & " entre os grupos mostrados") & ","""")"

    ' Tabela longa da arvore: categoria | parte | valor (so valores > 0)
    colArv = Bloco(3)
    Set cArv = Cel(5, colArv)
    cArv.Formula2 = "=IFERROR(LET(lin," & Sp(cRows) & ",p," & mWs.Range(Cel(4, col + 1), Cel(4, col + np)).Address & _
        ",m,HSTACK(" & somas_Lista(col, np) & "),nr,ROWS(lin),np,COLUMNS(p),rr,TOCOL(IF(SEQUENCE(1,np),lin)),pp,TOCOL(IF(SEQUENCE(nr),p)),vv,TOCOL(m)," & _
        "FILTER(HSTACK(rr,pp,vv),vv>0)),"""")"
    Cel(4, colArv).Value = NomeCol(jCat)
    Cel(4, colArv + 1).Value = NomeCol(jPar)
    Cel(4, colArv + 2).Value = mMedida
    EstiloCabecalho colArv, colArv + 2
    FormatarValores colArv + 2, mINum
    NomeSpill "dsG_Arvore_T", cArv
    mWs.Calculate

    ' Colunas empilhadas
    Set co = NovoGrafico("ColunasEmpilhadas")
    SeriesPartes co.Chart, col, np
    co.Chart.ChartType = xlColumnStacked
    co.Chart.HasLegend = True
    TituloConclusao co.Chart, Cel(3, col), mMedida & " por " & NomeCol(jCat) & " e " & NomeCol(jPar)
    EixoDesdeZero co.Chart, Cel(5, col + 1)
    RegistrarFormato co.Name, CodigoFormato(mINum)

    ' Barras empilhadas 100%
    Set co = NovoGrafico("BarrasEmpilhadas")
    SeriesPartes co.Chart, col, np
    co.Chart.ChartType = xlBarStacked100
    co.Chart.HasLegend = True
    TituloConclusao co.Chart, Cel(3, col + 1), U("Composi\u00e7\u00e3o de ") & mMedida & " por " & NomeCol(jCat)
    RegistrarFormato co.Name, "P"

    ' Arvore hierarquica
    GraficoArvore cArv, Cel(3, col + 2), mMedida & " por " & NomeCol(jCat) & " e " & NomeCol(jPar)
End Sub

' Lista "F5#,G5#,..." das colunas de partes (para HSTACK).
Private Function somas_Lista(ByVal col As Long, ByVal np As Long) As String
    Dim i As Long, s As String
    For i = 1 To np
        If i > 1 Then s = s & ","
        s = s & Sp(Cel(5, col + i))
    Next i
    somas_Lista = s
End Function

' Uma serie por parte; o nome da serie fica ligado ao cabecalho da parte.
Private Sub SeriesPartes(ByVal ch As Chart, ByVal col As Long, ByVal np As Long)
    Dim i As Long, s As Series
    For i = 1 To np
        Set s = ch.SeriesCollection.NewSeries
        On Error Resume Next
        s.Name = "='" & mWs.Name & "'!" & Cel(4, col + i).Address
        If Err.Number <> 0 Then Err.Clear: s.Name = modUtil.TextoSeguro(Cel(4, col + i).Value)
        On Error GoTo 0
        LigarSerie s, "dsG_Emp_C", "dsG_Emp_V" & i, Cel(5, col), Cel(5, col + i)
    Next i
End Sub

' Arvore de um nivel (quando nao ha 2a categoria): usa o bloco de colunas agrupadas.
Private Sub GraficoArvoreSimples()
    Dim cCat As Range, cVal As Range, col As Long, cArv As Range
    On Error Resume Next
    Set cCat = mWb.Names("dsG_ColunasAgrupadas_C").RefersToRange
    Set cVal = mWb.Names("dsG_ColunasAgrupadas_V1").RefersToRange
    On Error GoTo 0
    If cCat Is Nothing Or cVal Is Nothing Then Exit Sub
    col = Bloco(2)
    Set cArv = Cel(5, col)
    cArv.Formula2 = "=IFERROR(FILTER(HSTACK(" & cCat.Cells(1, 1).Address & "#," & cVal.Cells(1, 1).Address & "#)," & _
        cVal.Cells(1, 1).Address & "#>0),"""")"
    Cel(4, col).Value = "Grupo"
    Cel(4, col + 1).Value = mMedida
    EstiloCabecalho col, col + 1
    FormatarValores col + 1, mINum
    NomeSpill "dsG_Arvore_T", cArv
    Cel(3, col).Formula2 = "=" & cCat.Cells(1, 1).Offset(-2, 0).Address    ' mesmo titulo das colunas agrupadas
    mWs.Calculate
    GraficoArvore cArv, Cel(3, col), mMedida
End Sub

' Cria a arvore (treemap) a partir do intervalo derramado atual.
Private Sub GraficoArvore(ByVal cArv As Range, ByVal cTitulo As Range, ByVal reserva As String)
    Dim fonte As Range, shp As Object, lft As Double, tp As Double

    Set fonte = IntervaloDerramado(cArv)
    If fonte.Rows.Count < 2 Then
        Aviso U("\u00c1rvore hier\u00e1rquica n\u00e3o gerada: menos de 2 grupos com valor positivo.")
        Exit Sub
    End If
    lft = mEsq + (mN Mod 2) * (mLarg + 12)
    tp = mTopo + (mN \ 2) * (ALTURA + 12)
    ' No Excel para Mac, SetSourceData no treemap da erro 445 ("o objeto nao da
    ' suporte a esta acao"). O caminho que funciona e o da interface: selecionar
    ' a tabela e criar o grafico, que entao le a selecao.
    On Error Resume Next
    mWs.Activate
    fonte.Select
    Set shp = mWs.Shapes.AddChart2(-1, XL_TREEMAP, lft, tp, mLarg, ALTURA)
    If Err.Number = 0 Then
        If shp.Chart.SeriesCollection.Count = 0 Then Err.Raise 445
    End If
    mWs.Range("A1").Select
    If Err.Number <> 0 Then
        Err.Clear
        If Not shp Is Nothing Then shp.Delete
        On Error GoTo 0
        Aviso U("\u00c1rvore hier\u00e1rquica n\u00e3o gerada: esta vers\u00e3o do Excel n\u00e3o permitiu criar o gr\u00e1fico por macro.")
        Exit Sub
    End If
    shp.Name = "DS_GArvore"
    shp.Chart.HasLegend = False         ' os rotulos ja trazem os nomes (menos ruido)
    On Error GoTo 0
    mN = mN + 1
    TituloConclusao shp.Chart, cTitulo, reserva
    RegistrarFormato shp.Name, "-"
End Sub

' Rosca ou pizza: categoria de 2 a 5 grupos, da maior para a menor fatia.
Private Sub GraficoFatias(ByVal kind As String, ByVal jDim As Long)
    Dim col As Long, cCat As Range, cVal As Range, co As ChartObject, s As Series, verbo As String

    col = Bloco(2)
    Set cCat = Cel(5, col)
    Set cVal = Cel(5, col + 1)
    cCat.Formula2 = "=IFERROR(" & GruposOrdenados(jDim, 0) & ",""(vazio)"")"
    cVal.Formula2 = "=IFERROR(" & Agg(mINum, Ref(jDim) & "," & Sp(cCat)) & ",0)"
    Cel(4, col).Value = NomeCol(jDim)
    Cel(4, col + 1).Value = NomeMedidaN(1)
    EstiloCabecalho col, col + 1
    FormatarValores col + 1, mINum
    NomeSpill "dsG_" & kind & "_C", cCat
    NomeSpill "dsG_" & kind & "_V", cVal
    If kind = "Pizza" Then verbo = " concentra " Else verbo = U(" \u00e9 a maior fatia: ")
    Cel(3, col).Formula2 = "=IFERROR(INDEX(" & Sp(cCat) & ",1)&" & Q(verbo) & "&TEXT(INDEX(" & Sp(cVal) & _
        ",1)/SUM(" & Sp(cVal) & "),""0%"")&" & Q(" de " & mMedida) & ","""")"
    mWs.Calculate

    Set co = NovoGrafico(kind)
    Set s = co.Chart.SeriesCollection.NewSeries
    s.Name = NomeMedidaN(1)
    LigarSerie s, "dsG_" & kind & "_C", "dsG_" & kind & "_V", cCat, cVal
    If kind = "Pizza" Then co.Chart.ChartType = xlPie Else co.Chart.ChartType = xlDoughnut
    co.Chart.HasLegend = False
    On Error Resume Next
    co.Chart.ChartGroups(1).FirstSliceAngle = 0
    If kind = "Rosca" Then co.Chart.ChartGroups(1).DoughnutHoleSize = 60
    On Error GoTo 0
    TituloConclusao co.Chart, Cel(3, col), mMedida & " por " & NomeCol(jDim)
    RegistrarFormato co.Name, "-"
End Sub

' Dispersao: pares numericos validos das duas primeiras medidas.
Private Sub GraficoDispersao(ByVal jX As Long, ByVal jY As Long)
    Dim col As Long, cX As Range, cY As Range, filtro As String, n As Long, co As ChartObject, s As Series
    Dim nx As String, ny As String

    col = Bloco(2)
    Set cX = Cel(5, col)
    Set cY = Cel(5, col + 1)
    filtro = "ISNUMBER(" & Ref(jX) & ")*ISNUMBER(" & Ref(jY) & ")"
    cX.Formula2 = "=IFERROR(FILTER(" & Ref(jX) & "," & filtro & "),"""")"
    cY.Formula2 = "=IFERROR(FILTER(" & Ref(jY) & "," & filtro & "),"""")"
    mWs.Calculate
    On Error Resume Next
    n = Application.WorksheetFunction.Count(IntervaloDerramado(cX))
    On Error GoTo 0
    If n < MIN_PONTOS Then
        cX.ClearContents
        cY.ClearContents
        Aviso U("Dispers\u00e3o n\u00e3o gerada: h\u00e1 ") & n & U(" pares v\u00e1lidos e o guia pede 20 pontos ou mais.")
        Exit Sub
    End If

    nx = NomeCol(jX): ny = NomeCol(jY)
    Cel(4, col).Value = nx
    Cel(4, col + 1).Value = ny
    EstiloCabecalho col, col + 1
    FormatarValores col, jX
    FormatarValores col + 1, jY
    NomeSpill "dsG_Dispersao_X", cX
    NomeSpill "dsG_Dispersao_Y", cY
    Cel(3, col).Formula2 = "=IFERROR(LET(k,CORREL(" & Sp(cX) & "," & Sp(cY) & "),IF(k>=0.4," & _
        Q("Quanto maior " & nx & ", maior " & ny) & ",IF(k<=-0.4," & Q("Quanto maior " & nx & ", menor " & ny) & "," & _
        Q(nx & " e " & ny & U(" n\u00e3o mostram rela\u00e7\u00e3o clara")) & "))&IF(ABS(k)>=0.7," & Q(U(" (rela\u00e7\u00e3o forte)")) & _
        ",IF(ABS(k)>=0.4," & Q(U(" (rela\u00e7\u00e3o moderada)")) & ",""""))),"""")"
    mWs.Calculate

    Set co = NovoGrafico("Dispersao")
    Set s = co.Chart.SeriesCollection.NewSeries
    s.Name = ny
    LigarSerie s, "dsG_Dispersao_X", "dsG_Dispersao_Y", cX, cY
    co.Chart.ChartType = xlXYScatter
    co.Chart.HasLegend = False
    On Error Resume Next
    s.Trendlines.Add Type:=xlLinear
    With co.Chart.Axes(xlCategory)
        .HasTitle = True
        .AxisTitle.Text = nx
    End With
    With co.Chart.Axes(xlValue)
        .HasTitle = True
        .AxisTitle.Text = ny
    End With
    On Error GoTo 0
    TituloConclusao co.Chart, Cel(3, col), ny & " x " & nx
    RegistrarFormato co.Name, "XY" & CodigoFormato(jX) & CodigoFormato(jY)
End Sub

' Linha: soma mensal da medida pela coluna de data (contagem sem medida).
Private Sub GraficoLinha(ByVal iData As Long)
    Dim col As Long, c1 As Range, c2 As Range, d As String, crit As String, co As ChartObject, s As Series

    col = Bloco(2)
    Set c1 = Cel(5, col)
    Set c2 = Cel(5, col + 1)
    d = Ref(iData)
    c1.Formula2 = "=IFERROR(LET(d," & d & ",n,FILTER(d,ISNUMBER(d)),SORT(UNIQUE(DATE(YEAR(n),MONTH(n),1)))),"""")"
    crit = d & ","">=""&" & Sp(c1) & "," & d & ",""<""&DATE(YEAR(" & Sp(c1) & "),MONTH(" & Sp(c1) & ")+1,1)"
    c2.Formula2 = "=IFERROR(" & Agg(mINum, crit) & ",0)"
    Cel(4, col).Value = U("M\u00eas (") & NomeCol(iData) & ")"
    Cel(4, col + 1).Value = NomeMedidaN(1)
    EstiloCabecalho col, col + 1
    mWs.Range(c1, Cel(5 + 1200, col)).NumberFormat = "mmm/yy"
    FormatarValores col + 1, mINum
    NomeSpill "dsG_Linha_C", c1
    NomeSpill "dsG_Linha_V", c2
    Cel(3, col).Formula2 = "=IFERROR(LET(v," & Sp(c2) & ",n,ROWS(v),IF(n<2," & Q(NomeMedidaN(1) & U(" por m\u00eas")) & "," & _
        Q(NomeMedidaN(1)) & "&IF(INDEX(v,n)>=INDEX(v,n-1)," & Q(" subiu ") & "," & Q(" caiu ") & ")&TEXT(ABS(INDEX(v,n)/INDEX(v,n-1)-1),""0%"")&" & _
        Q(U(" no \u00faltimo m\u00eas")) & "))," & Q(NomeMedidaN(1) & U(" por m\u00eas")) & ")"
    mWs.Calculate

    Set co = NovoGrafico("Linha")
    Set s = co.Chart.SeriesCollection.NewSeries
    s.Name = NomeMedidaN(1)
    LigarSerie s, "dsG_Linha_C", "dsG_Linha_V", c1, c2
    co.Chart.ChartType = xlLineMarkers
    co.Chart.HasLegend = False
    On Error Resume Next
    co.Chart.Axes(xlCategory).CategoryType = xlCategoryScale
    co.Chart.Axes(xlCategory).TickLabels.NumberFormatLinked = False
    co.Chart.Axes(xlCategory).TickLabels.NumberFormat = "mmm/yy"
    On Error GoTo 0
    TituloConclusao co.Chart, Cel(3, col), NomeMedidaN(1) & U(" por m\u00eas")
    EixoDesdeZero co.Chart, c2
    RegistrarFormato co.Name, CodigoFormato(mINum)
End Sub

'==============================================================================
' Estilo dos graficos (chamado por modTema na montagem e na troca de tema)
'==============================================================================

'------------------------------------------------------------------------------
' EstilizarGrafico
' Proposito: aplicar o tema e as regras visuais do guia a um grafico DS_G*.
' Parametros: co - o grafico; t - tema ativo.
' Efeitos:    cores RGB, fontes de 12 pt, rotulos, legenda e linhas de grade.
'             On Error Resume Next e deliberado: tudo aqui e cosmetico e algumas
'             propriedades variam entre Mac, Windows e tipos de grafico.
'------------------------------------------------------------------------------
Public Sub EstilizarGrafico(ByVal co As ChartObject, ByRef t As TemaDS)
    Dim ch As Chart, kind As String, cinza As Long, grade As Long, fundoG As Long
    Dim i As Long, n As Long, s As Series, cor As Long

    On Error Resume Next
    Set ch = co.Chart
    kind = Mid$(co.Name, 5)
    cinza = modUtil.Misturar(t.Fundo, t.Texto, 0.22)
    grade = modUtil.Misturar(t.Fundo, t.Texto, 0.12)
    fundoG = modUtil.Misturar(t.Fundo, t.Primaria, 0.05)

    ' Base comum
    ch.ChartArea.Format.Fill.Visible = msoTrue
    ch.ChartArea.Format.Fill.Solid
    ch.ChartArea.Format.Fill.ForeColor.RGB = fundoG
    ch.ChartArea.Format.Line.Visible = msoFalse
    ch.PlotArea.Format.Fill.Visible = msoFalse
    ch.ChartArea.Font.Name = t.Fonte
    ch.ChartArea.Font.Size = 12
    ch.ChartArea.Font.Color = t.Texto
    If ch.HasTitle Then
        ch.ChartTitle.Font.Name = t.Fonte
        ch.ChartTitle.Font.Size = 13
        ch.ChartTitle.Font.Bold = True
        ch.ChartTitle.Font.Color = t.Texto
    End If
    If ch.HasLegend Then ch.Legend.Font.Color = t.Texto
    ch.Axes(xlCategory).TickLabels.Font.Color = t.Texto
    ch.Axes(xlValue).TickLabels.Font.Color = t.Texto
    ch.Axes(xlCategory).Format.Line.ForeColor.RGB = grade
    ch.Axes(xlValue).Format.Line.Visible = msoFalse
    ch.Axes(xlValue).MajorGridlines.Format.Line.ForeColor.RGB = grade

    Select Case kind
        Case "ColunasAgrupadas", "BarrasAgrupadas"
            n = ch.SeriesCollection.Count
            ch.ChartGroups(1).GapWidth = 60
            If n = 1 Then
                ' Uma cor para o que importa (a maior), cinza para o contexto
                Set s = ch.SeriesCollection(1)
                s.Format.Fill.Visible = msoTrue
                s.Format.Fill.Solid
                s.Format.Fill.ForeColor.RGB = cinza
                s.Format.Line.Visible = msoFalse
                s.Points(1).Format.Fill.ForeColor.RGB = t.Primaria
                s.HasDataLabels = True
                s.DataLabels.Font.Color = t.Texto
                s.DataLabels.Font.Size = 11
                s.DataLabels.Position = xlLabelPositionOutsideEnd
                ch.Axes(xlValue).HasMajorGridlines = False
                ch.HasLegend = False
            Else
                For i = 1 To n
                    Set s = ch.SeriesCollection(i)
                    s.Format.Fill.Visible = msoTrue
                    s.Format.Fill.Solid
                    s.Format.Fill.ForeColor.RGB = Tom(t, i)
                    s.Format.Line.Visible = msoFalse
                Next i
                ch.Axes(xlValue).HasMajorGridlines = True
                ch.HasLegend = True
                ch.Legend.Position = xlLegendPositionBottom
                ch.Legend.Font.Color = t.Texto
            End If
            If kind = "BarrasAgrupadas" Then
                ch.Axes(xlCategory).ReversePlotOrder = True      ' maior no topo
                ch.Axes(xlCategory).Crosses = xlMaximum
            End If

        Case "ColunasEmpilhadas", "BarrasEmpilhadas"
            n = ch.SeriesCollection.Count
            ch.ChartGroups(1).GapWidth = 55
            For i = 1 To n
                Set s = ch.SeriesCollection(i)
                If s.Name = "Outros" Then cor = cinza Else cor = Tom(t, i)
                s.Format.Fill.Visible = msoTrue
                s.Format.Fill.Solid
                s.Format.Fill.ForeColor.RGB = cor
                s.Format.Line.Visible = msoTrue
                s.Format.Line.ForeColor.RGB = fundoG
                s.Format.Line.Weight = 1
            Next i
            ch.HasLegend = True
            ch.Legend.Position = xlLegendPositionTop
            ch.Legend.Font.Color = t.Texto
            ch.Axes(xlValue).HasMajorGridlines = True
            If kind = "BarrasEmpilhadas" Then
                ch.Axes(xlCategory).ReversePlotOrder = True
                ch.Axes(xlCategory).Crosses = xlMaximum
            End If

        Case "Rosca", "Pizza"
            Set s = ch.SeriesCollection(1)
            n = s.Points.Count
            For i = 1 To n
                If i = 1 Then cor = t.Primaria Else cor = modUtil.Misturar(cinza, t.Fundo, (i - 2) * 0.18)
                s.Points(i).Format.Fill.Visible = msoTrue
                s.Points(i).Format.Fill.Solid
                s.Points(i).Format.Fill.ForeColor.RGB = cor
                s.Points(i).Format.Line.Visible = msoTrue
                s.Points(i).Format.Line.ForeColor.RGB = fundoG
                s.Points(i).Format.Line.Weight = 1.5
            Next i
            s.HasDataLabels = True
            With s.DataLabels
                .ShowCategoryName = True
                .ShowPercentage = True
                .ShowValue = False
                .ShowSeriesName = False
                .Separator = "  "
                .NumberFormat = "0%"
                .Font.Size = 11
                .Font.Color = t.Texto
            End With
            If kind = "Pizza" Then
                s.DataLabels.Position = xlLabelPositionOutsideEnd
            Else
                For i = 1 To n            ' rotulo dentro da fatia: cor de contraste
                    s.Points(i).DataLabel.Font.Color = modUtil.Contraste(s.Points(i).Format.Fill.ForeColor.RGB)
                Next i
            End If
            ch.HasLegend = False

        Case "Dispersao"
            Set s = ch.SeriesCollection(1)
            s.MarkerStyle = xlMarkerStyleCircle
            s.MarkerSize = 6
            s.MarkerBackgroundColor = t.Primaria
            s.MarkerForegroundColor = t.Primaria
            s.Format.Fill.Transparency = 0.35
            s.Format.Line.Visible = msoFalse
            s.Trendlines(1).Format.Line.ForeColor.RGB = t.Secundaria
            s.Trendlines(1).Format.Line.Weight = 2.25
            ch.Axes(xlCategory).AxisTitle.Font.Color = t.Texto
            ch.Axes(xlValue).AxisTitle.Font.Color = t.Texto
            ch.Axes(xlValue).HasMajorGridlines = True

        Case "Linha"
            Set s = ch.SeriesCollection(1)
            n = s.Points.Count
            s.Format.Line.Visible = msoTrue
            s.Format.Line.ForeColor.RGB = t.Primaria
            s.Format.Line.Weight = 2.25
            s.MarkerStyle = xlMarkerStyleCircle
            s.MarkerSize = 5
            s.MarkerBackgroundColor = t.Primaria
            s.MarkerForegroundColor = t.Primaria
            ' Destaque do ultimo valor: marcador e rotulo
            s.Points(n).MarkerSize = 9
            s.Points(n).MarkerBackgroundColor = t.Secundaria
            s.Points(n).MarkerForegroundColor = t.Secundaria
            s.Points(n).HasDataLabel = True
            s.Points(n).DataLabel.Position = xlLabelPositionAbove
            s.Points(n).DataLabel.Font.Bold = True
            s.Points(n).DataLabel.Font.Color = t.Texto
            ch.Axes(xlValue).HasMajorGridlines = True

        Case "Arvore"
            ' Regra 5 do guia: tons da cor primaria no grupo principal (o 1o da
            ' tabela, que esta ordenada) e cinza para os demais. A ordem dos
            ' pontos segue a da tabela dsG_Arvore_T.
            ColorirArvore co, t, cinza
    End Select
    On Error GoTo 0
End Sub

Private Sub ColorirArvore(ByVal co As ChartObject, ByRef t As TemaDS, ByVal cinza As Long)
    ' No treemap de 2 niveis, cada grupo tem um ponto "pai" antes dos seus blocos
    ' (testado no Mac: 15 linhas + 5 grupos = 20 pontos).
    Dim v As Variant, s As Series, r As Long, p As Long, k As Long, nPts As Long, pais As Boolean
    Dim grupo As String, anterior As String, topo As String, cor As Long
    On Error Resume Next
    v = co.Parent.Parent.Names("dsG_Arvore_T").RefersToRange.Value
    If Not IsArray(v) Then Exit Sub
    Set s = co.Chart.SeriesCollection(1)
    nPts = s.Points.Count
    pais = (UBound(v, 2) >= 3 And nPts > UBound(v, 1))
    topo = modUtil.TextoSeguro(v(1, 1))
    anterior = Chr$(0)
    For r = 1 To UBound(v, 1)
        grupo = modUtil.TextoSeguro(v(r, 1))
        If grupo <> anterior Then
            k = 0
            anterior = grupo
            If pais Then
                p = p + 1
                PintarPonto s, p, CorArvore(t, cinza, grupo = topo, 1)
            End If
        End If
        k = k + 1
        p = p + 1
        If p > nPts Then Exit For
        PintarPonto s, p, CorArvore(t, cinza, grupo = topo, k)
    Next r
    On Error GoTo 0
End Sub

Private Function CorArvore(ByRef t As TemaDS, ByVal cinza As Long, ByVal destaque As Boolean, ByVal k As Long) As Long
    If destaque Then
        CorArvore = modUtil.Misturar(t.Primaria, t.Fundo, (k - 1) * 0.22)
    Else
        CorArvore = modUtil.Misturar(cinza, t.Fundo, (k - 1) * 0.15)
    End If
End Function

' Pinta um ponto (bloco, fatia) e deixa o rotulo legivel sobre a cor.
Private Sub PintarPonto(ByVal s As Series, ByVal i As Long, ByVal cor As Long)
    On Error Resume Next
    s.Points(i).Format.Fill.Visible = msoTrue
    s.Points(i).Format.Fill.Solid
    s.Points(i).Format.Fill.ForeColor.RGB = cor
    s.Points(i).DataLabel.Font.Color = modUtil.Contraste(cor)
    On Error GoTo 0
End Sub

' Tons da cor primaria, do escuro ao claro (series de apoio).
Private Function Tom(ByRef t As TemaDS, ByVal i As Long) As Long
    Tom = modUtil.Misturar(t.Primaria, t.Fundo, (i - 1) * 0.28)
End Function

'==============================================================================
' Formatos numericos de eixos e rotulos (reaplicados no fim de AtualizarTudo)
'==============================================================================

'------------------------------------------------------------------------------
' FinalizarFormatos
' Proposito: aplicar formato neutro aos eixos de valores e aos rotulos. Testado
'            no Excel para Mac: o eixo le o formato na sintaxe local e o Excel
'            refaz o formato do 1o grafico ao criar outro; por isso e o ultimo
'            passo de AtualizarTudo.
' Parametros: wb - livro alvo.
'------------------------------------------------------------------------------
Public Sub FinalizarFormatos(ByVal wb As Workbook)
    Dim ws As Worksheet, co As ChartObject, cod As String, s As Series

    If mFmts Is Nothing Then Exit Sub
    Set ws = modUtil.ObterPlanilha(wb, modConfig.PLAN_DASH)
    If ws Is Nothing Then Exit Sub
    For Each co In ws.ChartObjects
        cod = ""
        On Error Resume Next
        cod = mFmts(co.Name)
        On Error GoTo 0
        If Len(cod) > 0 And cod <> "-" Then
            On Error Resume Next
            If Left$(cod, 2) = "XY" Then
                FormatarEixo co.Chart.Axes(xlCategory), Mid$(cod, 3, 1)
                FormatarEixo co.Chart.Axes(xlValue), Mid$(cod, 4, 1)
            Else
                FormatarEixo co.Chart.Axes(xlValue), cod
                For Each s In co.Chart.SeriesCollection
                    If s.HasDataLabels Then s.DataLabels.NumberFormat = FormatoNeutro(cod)
                    If s.Points(s.Points.Count).HasDataLabel Then _
                        s.Points(s.Points.Count).DataLabel.NumberFormat = FormatoNeutro(cod)
                Next s
            End If
            On Error GoTo 0
        End If
    Next co
End Sub

Private Sub FormatarEixo(ByVal ax As Axis, ByVal cod As String)
    On Error Resume Next
    ax.TickLabels.NumberFormatLinked = False
    ax.TickLabels.NumberFormat = FormatoNeutro(cod)
    On Error GoTo 0
End Sub

' "P" -> 0%; "N" -> milhar sem moeda e sem decimais (sintaxe local no Mac).
Private Function FormatoNeutro(ByVal cod As String) As String
    If cod = "P" Then
        FormatoNeutro = "0%"
    Else
#If Mac Then
        FormatoNeutro = "#" & Application.International(xlThousandsSeparator) & "##0"
#Else
        FormatoNeutro = "#,##0"
#End If
    End If
End Function

Private Sub RegistrarFormato(ByVal nome As String, ByVal cod As String)
    On Error Resume Next
    mFmts.Remove nome
    On Error GoTo 0
    mFmts.Add cod, nome
End Sub

' "P" se a coluna esta em porcentagem, "N" nos demais casos (0 = contagem).
Private Function CodigoFormato(ByVal j As Long) As String
    If j > 0 Then
        If InStr(1, FormatoColuna(j), "%") > 0 Then CodigoFormato = "P" Else CodigoFormato = "N"
    Else
        CodigoFormato = "N"
    End If
End Function

'------------------------------------------------------------------------------
' FormatoColuna
' Proposito: formato numerico de uma coluna de tblDados para cards e area
'            auxiliar. "Geral" vira #,##0 (somas) ou #,##0.00 (medias).
'------------------------------------------------------------------------------
Public Function FormatoColuna(ByVal j As Long, Optional ByVal ehMedia As Boolean = False, _
        Optional ByVal lo As ListObject) As String
    Dim f As String
    If lo Is Nothing Then Set lo = mLo
    On Error Resume Next
    f = CStr(lo.ListColumns(j).DataBodyRange.Cells(1, 1).NumberFormat)
    On Error GoTo 0
    If Len(f) = 0 Or f = "General" Or f = "@" Then
        If ehMedia Then f = "#,##0.00" Else f = "#,##0"
    End If
    FormatoColuna = f
End Function

'==============================================================================
' Utilitarios internos
'==============================================================================

Private Function Cel(ByVal lin As Long, ByVal col As Long) As Range
    Set Cel = mWs.Cells(lin, col)
End Function

Private Function Sp(ByVal r As Range) As String
    Sp = r.Address & "#"
End Function

Private Function Q(ByVal s As String) As String
    Q = modUtil.TextoFormula(s)
End Function

Private Function U(ByVal s As String) As String
    U = modUtil.U(s)
End Function

Private Function Ref(ByVal j As Long) As String
    Ref = modUtil.RefColuna(modConfig.TAB_DADOS, mLo.ListColumns(j).Name)
End Function

Private Function NomeCol(ByVal j As Long) As String
    NomeCol = mLo.ListColumns(j).Name
End Function

Private Function MedidaN(ByVal i As Long) As Long
    If mMeds.Count = 0 Then MedidaN = 0 Else MedidaN = mMeds(i)
End Function

Private Function NomeMedidaN(ByVal i As Long) As String
    If MedidaN(i) = 0 Then NomeMedidaN = "Registros" Else NomeMedidaN = NomeCol(MedidaN(i))
End Function

' SUMIFS da medida j (ou COUNTIFS se j = 0) com a lista "faixa,criterio,...".
Private Function Agg(ByVal j As Long, ByVal crit As String) As String
    If j > 0 Then
        Agg = "SUMIFS(" & Ref(j) & "," & crit & ")"
    Else
        Agg = "COUNTIFS(" & crit & ")"
    End If
End Function

Private Function TotalMedida() As String
    If mINum > 0 Then
        TotalMedida = "AGGREGATE(9,6," & Ref(mINum) & ")"
    Else
        TotalMedida = "ROWS(" & modConfig.TAB_DADOS & ")"
    End If
End Function

' Grupos unicos (sem vazios/erros) ordenados pela medida; topN = 0 -> todos.
Private Function GruposOrdenados(ByVal jCat As Long, ByVal topN As Long) As String
    Dim s As String
    s = "LET(c," & Ref(jCat) & ",u,UNIQUE(FILTER(c,IFERROR(c<>"""",FALSE))),s,SORTBY(u," & Agg(mINum, "c,u") & ",-1),"
    If topN > 0 Then s = s & "TAKE(s," & topN & "))" Else s = s & "s)"
    GruposOrdenados = s
End Function

Private Function NGrupos(ByVal jCat As Long) As String
    NGrupos = "LET(c," & Ref(jCat) & ",ROWS(UNIQUE(FILTER(c,IFERROR(c<>"""",FALSE)))))"
End Function

' Reserva colunas na area auxiliar e devolve a primeira.
Private Function Bloco(ByVal largura As Long) As Long
    Bloco = mCol
    mCol = mCol + largura + 1
End Function

Private Sub EstiloCabecalho(ByVal c1 As Long, ByVal c2 As Long)
    mWs.Range(Cel(4, c1), Cel(4, c2)).Style = "DS_Header"
    mWs.Range(mWs.Columns(c1), mWs.Columns(c2)).ColumnWidth = 16
End Sub

Private Sub FormatarValores(ByVal col As Long, ByVal j As Long)
    mWs.Range(Cel(5, col), Cel(5 + 5000, col)).NumberFormat = FormatoColuna(j)
End Sub

Private Sub NomeSpill(ByVal nome As String, ByVal celula As Range)
    mWb.Names.Add Name:=nome, RefersTo:="='" & mWs.Name & "'!" & celula.Address & "#"
End Sub

Private Sub Aviso(ByVal txt As String)
    mAvisos.Add txt
End Sub

' Grafico vazio na proxima posicao da grade (2 por linha).
Private Function NovoGrafico(ByVal kind As String) As ChartObject
    Dim co As ChartObject, lft As Double, tp As Double
    lft = mEsq + (mN Mod 2) * (mLarg + 12)
    tp = mTopo + (mN \ 2) * (ALTURA + 12)
    Set co = mWs.ChartObjects.Add(lft, tp, mLarg, ALTURA)
    co.Name = "DS_G" & kind
    Do While co.Chart.SeriesCollection.Count > 0
        co.Chart.SeriesCollection(1).Delete
    Loop
    mN = mN + 1
    Set NovoGrafico = co
End Function

'------------------------------------------------------------------------------
' TituloConclusao
' Proposito: ligar o titulo do grafico a celula com a frase-conclusao (regra 1
'            do guia). Se o Excel recusar a ligacao, usa o texto calculado.
'------------------------------------------------------------------------------
Private Sub TituloConclusao(ByVal ch As Chart, ByVal cel As Range, ByVal reserva As String)
    Dim tt As Object, txt As String
    txt = modUtil.TextoSeguro(cel.Value)
    ch.HasTitle = True
    Set tt = ch.ChartTitle
    On Error Resume Next
    If Len(txt) > 0 Then
        tt.Formula = "='" & mWs.Name & "'!" & cel.Address
        If Err.Number <> 0 Then
            Err.Clear
            tt.Text = txt
        End If
    Else
        tt.Text = reserva
    End If
    On Error GoTo 0
End Sub

' Liga a serie aos nomes dsG_*; se o Excel recusar, ao intervalo derramado atual.
Private Sub LigarSerie(ByVal s As Series, ByVal nomeX As String, ByVal nomeY As String, _
        ByVal cX As Range, ByVal cY As Range)
    On Error Resume Next
    s.XValues = "=" & modUtil.RefNomeLivro(mWb, nomeX)
    s.Values = "=" & modUtil.RefNomeLivro(mWb, nomeY)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        s.XValues = IntervaloDerramado(cX)
        s.Values = IntervaloDerramado(cY)
    End If
    On Error GoTo 0
End Sub

Private Function IntervaloDerramado(ByVal c As Range) As Range
    On Error Resume Next
    Set IntervaloDerramado = c.SpillingToRange
    On Error GoTo 0
    If IntervaloDerramado Is Nothing Then Set IntervaloDerramado = c
End Function

' Barras e colunas comecam em zero (regra 1 do guia) quando nao ha negativos.
Private Sub EixoDesdeZero(ByVal ch As Chart, ByVal celValores As Range)
    On Error Resume Next
    If Application.WorksheetFunction.Min(IntervaloDerramado(celValores)) >= 0 Then
        ch.Axes(xlValue).MinimumScale = 0
    End If
    On Error GoTo 0
End Sub

' Duas medidas tem a mesma escala se o total de uma estiver entre 1/10 e 10
' vezes o da outra (o guia: "duas escalas diferentes: crie dois graficos").
Private Function MesmaEscala(ByVal j1 As Long, ByVal j2 As Long) As Boolean
    Dim t1 As Double, t2 As Double, r As Double
    On Error GoTo Nao
    t1 = Application.WorksheetFunction.Aggregate(9, 6, mLo.ListColumns(j1).DataBodyRange)
    t2 = Application.WorksheetFunction.Aggregate(9, 6, mLo.ListColumns(j2).DataBodyRange)
    If t1 = 0 Then Exit Function
    r = Abs(t2 / t1)
    MesmaEscala = (r >= 1 / ESCALA And r <= ESCALA)
Nao:
End Function

' A maior fatia da categoria tem 50% ou mais da medida (e nada e negativo)?
Private Function GruposNumericos(ByVal grupos As Collection) As Boolean
    ' True se os grupos sao codigos numericos (ex.: TC = 1, 2, 3)
    If grupos.Count = 0 Then Exit Function
    GruposNumericos = IsNumeric(grupos(1)) And VarType(grupos(1)) <> vbString
End Function

Private Function Dominante(ByVal jCat As Long) As Boolean
    Dim grupos As Collection, g As Variant, v As Variant, total As Double, maior As Double
    Dim faixaCat As Range, faixaMed As Range

    Set grupos = modDados.ValoresDistintos(mLo, jCat, MAX_FATIAS)
    Set faixaCat = mLo.ListColumns(jCat).DataBodyRange
    If mINum > 0 Then Set faixaMed = mLo.ListColumns(mINum).DataBodyRange
    For Each g In grupos
        If mINum > 0 Then
            v = Application.SumIfs(faixaMed, faixaCat, g)
        Else
            v = Application.CountIfs(faixaCat, g)
        End If
        If IsError(v) Then Exit Function
        If v < 0 Then Exit Function
        total = total + v
        If v > maior Then maior = v
    Next g
    If total > 0 Then Dominante = (maior / total >= DOMINANCIA)
End Function
