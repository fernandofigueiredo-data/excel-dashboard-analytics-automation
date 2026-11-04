Attribute VB_Name = "modFixar"
Option Explicit
'==============================================================================
' modFixar
' "Fixar grafico": copia um grafico gerado (DS_G*) do Dashboard para a planilha
' "Meus graficos", junto com as formulas da area auxiliar que o alimentam.
' A copia passa a ser do usuario: pode trocar o tipo, as series e o formato.
' Atualizar e a troca de tema nao a apagam nem a repintam, e ela continua
' lendo tblDados (acompanha a base quando os dados mudam).
'
' COMO FUNCIONA
'   1. Descobre as colunas da area auxiliar que o grafico usa: series, titulo,
'      nomes dsG_* e as referencias das formulas, ate fechar o conjunto.
'   2. Reescreve essas formulas (linhas 3 a 5) na proxima faixa livre da area
'      auxiliar de "Meus graficos", deslocando as referencias ($S$5 -> $AD$5).
'   3. Cria nomes dsM<n>_* (os dsG_* sao apagados a cada Atualizar).
'   4. Copia o grafico e religa series e titulo a faixa nova. A arvore
'      (treemap) e recriada, porque o Excel para Mac nao religa a serie dela.
'
' LIMITES
'   - Se a nova base nao tiver as colunas usadas, o grafico fixado fica vazio.
'   - Apagar um grafico fixado nao apaga a faixa auxiliar dele (pode ser
'     apagada a mao, nas colunas da area auxiliar de "Meus graficos").
'==============================================================================

Private Const LIN_TOPO As Long = 4        ' 1a linha dos graficos fixados
Private Const COL_INI As Long = 2         ' coluna B
Private Const COL_AUX As Long = 19        ' coluna S: inicio da area auxiliar
Private Const ESPACO As Double = 12
Private Const XL_TREEMAP As Long = 117

Private mDash As String                   ' nome da planilha Dashboard
Private mMeus As String                   ' nome da planilha Meus graficos

'------------------------------------------------------------------------------
' FixarGrafico
' Proposito: macro do botao "Fixar grafico" do Dashboard (e da Faixa de Opcoes
'            no suplemento). Usa o grafico selecionado ou pergunta qual fixar.
'------------------------------------------------------------------------------
Public Sub FixarGrafico()
    Dim wb As Workbook, ws As Worksheet, co As ChartObject, d As String, nome As String

    Set wb = modUtil.LivroAlvo()
    If wb Is Nothing Then Exit Sub
    Set ws = modUtil.ObterPlanilha(wb, modConfig.PLAN_DASH)
    If ws Is Nothing Then
        MsgBox modUtil.U("Este livro ainda n\u00e3o tem Dashboard. Clique em Atualizar primeiro."), vbInformation, "Dashboard"
        Exit Sub
    End If

    Set co = GraficoSelecionado(ws)
    If co Is Nothing Then Set co = EscolherGrafico(ws)
    If co Is Nothing Then Exit Sub

    On Error GoTo Falha
    modUtil.DesligarApp
    nome = Fixar(wb, ws, co)
    modUtil.RestaurarApp
    MsgBox modUtil.U("O gr\u00e1fico foi copiado para a aba """) & mMeus & """." & vbLf & vbLf & _
           modUtil.U("Essa c\u00f3pia \u00e9 sua: clique com o bot\u00e3o direito nela para trocar o tipo " & _
           "(Alterar Tipo de Gr\u00e1fico), os dados (Selecionar Dados) ou o formato. " & _
           "Atualizar e a troca de tema n\u00e3o mexem nela, e ela continua lendo a tabela tblDados."), _
           vbInformation, "Dashboard"
    Exit Sub

Falha:
    d = Err.Number & " - " & Err.Description
    modUtil.RestaurarApp
    MsgBox modUtil.U("N\u00e3o foi poss\u00edvel fixar o gr\u00e1fico." & vbLf & vbLf & "Detalhe: ") & d, _
           vbExclamation, "Dashboard"
End Sub

Public Sub Faixa_FixarGrafico(control As IRibbonControl)
    FixarGrafico
End Sub

'==============================================================================
' Escolha do grafico
'==============================================================================

' Grafico gerado que esta selecionado no Dashboard (Nothing se nao houver).
Private Function GraficoSelecionado(ByVal ws As Worksheet) As ChartObject
    Dim co As Object
    On Error Resume Next
    If ActiveChart Is Nothing Then Exit Function
    Set co = ActiveChart.Parent
    If TypeName(co) <> "ChartObject" Then Exit Function
    If Not co.Parent Is ws Then Exit Function
    If Left$(co.Name, 4) <> "DS_G" Then Exit Function
    Set GraficoSelecionado = co
End Function

' Lista os graficos gerados (na ordem da tela) e pergunta o numero.
Private Function EscolherGrafico(ByVal ws As Worksheet) As ChartObject
    Dim lista As Collection, co As ChartObject, i As Long, j As Long, txt As String, resp As String, k As Long

    Set lista = New Collection
    For Each co In ws.ChartObjects
        If Left$(co.Name, 4) = "DS_G" Then
            k = 0
            For i = 1 To lista.Count
                If Antes(co, lista(i)) Then k = i: Exit For
            Next i
            If k = 0 Then lista.Add co Else lista.Add co, Before:=k
        End If
    Next co
    If lista.Count = 0 Then
        MsgBox modUtil.U("N\u00e3o h\u00e1 gr\u00e1ficos gerados no Dashboard. Clique em Atualizar primeiro."), _
               vbInformation, "Dashboard"
        Exit Function
    End If

    For i = 1 To lista.Count
        txt = txt & i & " - " & Resumo(lista(i)) & vbLf
    Next i
    resp = InputBox(modUtil.U("Qual gr\u00e1fico voc\u00ea quer fixar para editar? Digite o n\u00famero:") & _
                    vbLf & vbLf & txt, modUtil.U("Fixar gr\u00e1fico"), "1")
    resp = Trim$(resp)
    If Len(resp) = 0 Then Exit Function
    If Not IsNumeric(resp) Then GoTo Invalido
    j = CLng(resp)
    If j < 1 Or j > lista.Count Then GoTo Invalido
    Set EscolherGrafico = lista(j)
    Exit Function
Invalido:
    MsgBox modUtil.U("Digite um n\u00famero de 1 a ") & lista.Count & ".", vbExclamation, "Dashboard"
End Function

' Ordem de leitura: de cima para baixo, da esquerda para a direita.
Private Function Antes(ByVal a As ChartObject, ByVal b As ChartObject) As Boolean
    If Abs(a.Top - b.Top) > 5 Then Antes = (a.Top < b.Top) Else Antes = (a.Left < b.Left)
End Function

' Titulo curto do grafico para a lista.
Private Function Resumo(ByVal co As ChartObject) As String
    Dim s As String
    On Error Resume Next
    If co.Chart.HasTitle Then s = co.Chart.ChartTitle.Text
    On Error GoTo 0
    If Len(s) = 0 Then s = Mid$(co.Name, 5)
    If Len(s) > 48 Then s = Left$(s, 45) & "..."
    Resumo = s
End Function

'==============================================================================
' Copia
'==============================================================================

'------------------------------------------------------------------------------
' Fixar
' Proposito: copiar a faixa auxiliar e o grafico para "Meus graficos".
' Retorno:    nome do grafico criado.
'------------------------------------------------------------------------------
Private Function Fixar(ByVal wb As Workbook, ByVal ws As Worksheet, ByVal co As ChartObject) As String
    Dim wsM As Worksheet, kind As String, n As Long, minC As Long, maxC As Long, destC As Long, delta As Long
    Dim s As Series, f As String, tit As String, lft As Double, tp As Double, novo As Object, nome As String

    mDash = ws.Name
    mMeus = modUtil.U("Meus gr\u00e1ficos")
    Set wsM = GarantirPlanilhaMeus(wb, ws)
    kind = Mid$(co.Name, 5)
    n = ProximoNumero(wb, wsM)
    nome = "Meu" & n & "_" & kind

    ' 1. Colunas usadas pelo grafico
    minC = 0: maxC = 0
    On Error Resume Next
    For Each s In co.Chart.SeriesCollection
        f = ""
        f = s.Formula
        If Len(f) > 0 Then Coletar wb, f, minC, maxC
    Next s
    tit = ""
    tit = co.Chart.ChartTitle.Formula
    On Error GoTo 0
    If Left$(tit, 1) = "=" Then Ler tit, False, minC, maxC Else tit = ""
    If kind = "Arvore" Then Coletar wb, "=" & modUtil.RefNomeLivro(wb, "dsG_Arvore_T"), minC, maxC
    If maxC = 0 Then Err.Raise vbObjectError + 1, , modUtil.U("os dados deste gr\u00e1fico n\u00e3o foram encontrados na \u00e1rea auxiliar. Clique em Atualizar e tente de novo.")
    FecharConjunto ws, minC, maxC

    ' 2. Formulas na proxima faixa livre
    destC = ProximaColuna(wsM)
    delta = destC - minC
    CopiarFaixa ws, wsM, minC, maxC, delta, n
    wsM.Calculate

    ' 3 e 4. Grafico
    PosicaoLivre wsM, co.Width, lft, tp
    If kind = "Arvore" Then
        Set novo = RecriarArvore(wb, wsM, co, delta, lft, tp)
    Else
        Set novo = CopiarGrafico(wb, ws, wsM, co, delta, n, lft, tp)
    End If
    novo.Name = nome
    If Len(tit) > 0 Then
        On Error Resume Next
        novo.Chart.ChartTitle.Formula = Reescrever(tit, delta, False, 0, 0)
        On Error GoTo 0
    End If

    On Error Resume Next
    wsM.Activate
    wsM.Range("A1").Select
    ActiveWindow.ScrollRow = 1
    On Error GoTo 0
    Fixar = nome
End Function

' Cria "Meus graficos" logo depois do Dashboard, se ainda nao existir.
Private Function GarantirPlanilhaMeus(ByVal wb As Workbook, ByVal ws As Worksheet) As Worksheet
    Dim wsM As Worksheet
    Set wsM = modUtil.ObterPlanilha(wb, mMeus)
    If Not wsM Is Nothing Then
        Set GarantirPlanilhaMeus = wsM
        Exit Function
    End If
    Set wsM = wb.Worksheets.Add(After:=ws)
    wsM.Name = mMeus
    On Error Resume Next
    wsM.Cells.Style = "DS_Texto"
    wsM.Columns(1).ColumnWidth = 2
    wsM.Range(wsM.Columns(COL_INI), wsM.Columns(16)).ColumnWidth = 11
    wsM.Range(wsM.Cells(1, COL_INI), wsM.Cells(1, 16)).Style = "DS_Titulo"
    wsM.Rows(1).RowHeight = 36
    wsM.Cells(1, COL_INI).Value = mMeus
    wsM.Cells(2, COL_INI).Value = modUtil.U("Gr\u00e1ficos fixados: edite \u00e0 vontade (bot\u00e3o direito \u203a " & _
        "Alterar Tipo de Gr\u00e1fico ou Selecionar Dados). Atualizar e o tema n\u00e3o mexem nesta aba.")
    wsM.Cells(2, COL_AUX).Value = modUtil.U("\u00c1rea auxiliar dos gr\u00e1ficos fixados (n\u00e3o editar)")
    wsM.Activate
    ActiveWindow.DisplayGridlines = False
    On Error GoTo 0
    Set GarantirPlanilhaMeus = wsM
End Function

' Proximo n livre: maior n ja usado em nomes dsM<n>_* ou graficos Meu<n>_* + 1.
Private Function ProximoNumero(ByVal wb As Workbook, ByVal wsM As Worksheet) As Long
    Dim nm As Name, co As ChartObject, n As Long, p As Long, maior As Long
    For Each nm In wb.Names
        If Left$(nm.Name, 3) = "dsM" Then
            p = InStr(4, nm.Name, "_")
            If p > 4 Then
                If IsNumeric(Mid$(nm.Name, 4, p - 4)) Then
                    n = CLng(Mid$(nm.Name, 4, p - 4))
                    If n > maior Then maior = n
                End If
            End If
        End If
    Next nm
    For Each co In wsM.ChartObjects
        If Left$(co.Name, 3) = "Meu" Then
            p = InStr(4, co.Name, "_")
            If p > 4 Then
                If IsNumeric(Mid$(co.Name, 4, p - 4)) Then
                    n = CLng(Mid$(co.Name, 4, p - 4))
                    If n > maior Then maior = n
                End If
            End If
        End If
    Next co
    ProximoNumero = maior + 1
End Function

' Junta as colunas citadas por um texto de formula (serie, nome ou titulo);
' os nomes dsG_* citados entram pelo endereco a que se referem.
Private Sub Coletar(ByVal wb As Workbook, ByVal f As String, ByRef minC As Long, ByRef maxC As Long)
    Dim p As Long, nomeG As String
    Ler f, False, minC, maxC
    p = InStr(1, f, "dsG_")
    Do While p > 0
        nomeG = Identificador(f, p)
        On Error Resume Next
        Ler wb.Names(nomeG).RefersTo, False, minC, maxC
        On Error GoTo 0
        p = InStr(p + Len(nomeG), f, "dsG_")
    Loop
End Sub

' Amplia [minC, maxC] ate incluir tudo o que as formulas da faixa citam.
Private Sub FecharConjunto(ByVal ws As Worksheet, ByRef minC As Long, ByRef maxC As Long)
    Dim c As Long, r As Long, a As Long, b As Long, mudou As Boolean
    Do
        mudou = False
        a = minC: b = maxC
        For c = minC To maxC
            For r = 3 To 5
                If ws.Cells(r, c).HasFormula Then Ler ws.Cells(r, c).Formula2, True, a, b
            Next r
        Next c
        If a < minC Or b > maxC Then
            minC = a: maxC = b
            mudou = True
        End If
    Loop While mudou
End Sub

' Primeira coluna livre da area auxiliar de "Meus graficos" (uma de folga).
Private Function ProximaColuna(ByVal wsM As Worksheet) As Long
    Dim r As Long, c As Long, maior As Long
    For r = 3 To 5
        c = wsM.Cells(r, wsM.Columns.Count).End(xlToLeft).Column
        If c > maior Then maior = c
    Next r
    ' A linha 5 pode derramar para a direita
    c = wsM.Cells(5, maior).Column
    On Error Resume Next
    If wsM.Cells(5, maior).HasSpill Then c = wsM.Cells(5, maior).SpillParent.SpillingToRange.Column + _
        wsM.Cells(5, maior).SpillParent.SpillingToRange.Columns.Count - 1
    On Error GoTo 0
    If c > maior Then maior = c
    If maior < COL_AUX Then ProximaColuna = COL_AUX Else ProximaColuna = maior + 2
End Function

' Copia as linhas 3 a 5 das colunas minC..maxC, deslocadas em delta colunas.
Private Sub CopiarFaixa(ByVal ws As Worksheet, ByVal wsM As Worksheet, ByVal minC As Long, ByVal maxC As Long, _
        ByVal delta As Long, ByVal n As Long)
    Dim c As Long, r As Long, src As Range, dst As Range

    wsM.Cells(2, minC + delta).Value = modUtil.U("Gr\u00e1fico fixado ") & n & " (" & Format$(Now, "dd/mm/yyyy hh:nn") & ")"
    For c = minC To maxC
        wsM.Columns(c + delta).ColumnWidth = ws.Columns(c).ColumnWidth
        On Error Resume Next
        wsM.Range(wsM.Cells(5, c + delta), wsM.Cells(5005, c + delta)).NumberFormat = ws.Cells(5, c).NumberFormat
        wsM.Cells(4, c + delta).Style = ws.Cells(4, c).Style.Name
        On Error GoTo 0
        For r = 3 To 5
            Set src = ws.Cells(r, c)
            Set dst = wsM.Cells(r, c + delta)
            If Not FilhoDeDerramamento(src) Then
                If src.HasFormula Then
                    dst.Formula2 = Reescrever(src.Formula2, delta, True, 0, 0)
                ElseIf Len(modUtil.TextoSeguro(src.Value)) > 0 Then
                    dst.Value = src.Value
                End If
            End If
        Next r
    Next c
End Sub

' True para uma celula preenchida pelo derramamento de outra (nao a ancora).
Private Function FilhoDeDerramamento(ByVal c As Range) As Boolean
    On Error Resume Next
    If c.HasSpill Then FilhoDeDerramamento = (c.SpillParent.Address <> c.Address)
End Function

' Proxima posicao da grade de 2 graficos por linha (abaixo do mais baixo).
Private Sub PosicaoLivre(ByVal wsM As Worksheet, ByVal larg As Double, ByRef lft As Double, ByRef tp As Double)
    Dim co As ChartObject, baixo As Double, ultimo As ChartObject
    lft = wsM.Columns(COL_INI).Left
    tp = wsM.Rows(LIN_TOPO).Top
    For Each co In wsM.ChartObjects
        If ultimo Is Nothing Then
            Set ultimo = co
        ElseIf co.Top > ultimo.Top + 5 Or (Abs(co.Top - ultimo.Top) <= 5 And co.Left > ultimo.Left) Then
            Set ultimo = co
        End If
    Next co
    If ultimo Is Nothing Then Exit Sub
    If ultimo.Left < lft + 5 Then           ' ultimo esta na coluna da esquerda
        lft = ultimo.Left + ultimo.Width + ESPACO
        tp = ultimo.Top
    Else
        For Each co In wsM.ChartObjects
            If co.Top + co.Height > baixo Then baixo = co.Top + co.Height
        Next co
        tp = baixo + ESPACO
    End If
End Sub

' Copia o grafico, cria os nomes dsM<n>_* e religa as series.
Private Function CopiarGrafico(ByVal wb As Workbook, ByVal ws As Worksheet, ByVal wsM As Worksheet, _
        ByVal co As ChartObject, ByVal delta As Long, ByVal n As Long, ByVal lft As Double, ByVal tp As Double) As ChartObject
    Dim novo As ChartObject, s As Series, f As String, nf As String, antes As Long

    antes = wsM.ChartObjects.Count
    co.Copy
    wsM.Activate
    wsM.Cells(LIN_TOPO, COL_INI).Select
    wsM.Paste
    If wsM.ChartObjects.Count <= antes Then Err.Raise vbObjectError + 2, , modUtil.U("o Excel n\u00e3o colou a c\u00f3pia do gr\u00e1fico.")
    Set novo = wsM.ChartObjects(wsM.ChartObjects.Count)
    novo.Left = lft
    novo.Top = tp
    novo.Width = co.Width
    novo.Height = co.Height

    For Each s In novo.Chart.SeriesCollection
        f = ""
        On Error Resume Next
        f = s.Formula
        On Error GoTo 0
        If Len(f) > 0 Then
            nf = TrocarNomes(wb, Reescrever(f, delta, False, 0, 0), delta, n)
            LigarFormula s, nf
        End If
    Next s
    Set CopiarGrafico = novo
End Function

' Aplica a formula =SERIES(...) inteira; se o Excel recusar, parte por parte.
Private Sub LigarFormula(ByVal s As Series, ByVal nf As String)
    Dim partes As Collection
    On Error Resume Next
    s.Formula = nf
    If Err.Number = 0 Then Exit Sub
    Err.Clear
    Set partes = ArgumentosSerie(nf)
    If partes.Count < 3 Then Exit Sub
    If Len(partes(1)) > 0 Then
        If Left$(partes(1), 1) = """" Then s.Name = Mid$(partes(1), 2, Len(partes(1)) - 2) Else s.Name = "=" & partes(1)
    End If
    If Len(partes(2)) > 0 Then s.XValues = "=" & partes(2)
    s.Values = "=" & partes(3)
    On Error GoTo 0
End Sub

' Argumentos de =SERIES(a,b,c,d) separados nas virgulas de nivel 1.
Private Function ArgumentosSerie(ByVal f As String) As Collection
    Dim c As New Collection, i As Long, ch As String, nivel As Long, aspas As Boolean, atual As String, ini As Long
    ini = InStr(1, f, "(")
    For i = ini + 1 To Len(f)
        ch = Mid$(f, i, 1)
        If ch = """" Then aspas = Not aspas
        If Not aspas And ch = "(" Then nivel = nivel + 1
        If Not aspas And ch = ")" Then
            If nivel = 0 Then Exit For
            nivel = nivel - 1
        End If
        If Not aspas And nivel = 0 And ch = "," Then
            c.Add atual
            atual = ""
        Else
            atual = atual & ch
        End If
    Next i
    c.Add atual
    Set ArgumentosSerie = c
End Function

' Troca dsG_X por dsM<n>_X no texto, criando cada nome novo na faixa copiada.
Private Function TrocarNomes(ByVal wb As Workbook, ByVal f As String, ByVal delta As Long, ByVal n As Long) As String
    Dim p As Long, nomeG As String, nomeM As String, alvo As String
    p = InStr(1, f, "dsG_")
    Do While p > 0
        nomeG = Identificador(f, p)
        nomeM = "dsM" & n & "_" & Mid$(nomeG, 5)
        If Not NomeExiste(wb, nomeM) Then
            alvo = Reescrever(wb.Names(nomeG).RefersTo, delta, False, 0, 0)
            wb.Names.Add Name:=nomeM, RefersTo:=alvo
        End If
        f = Left$(f, p - 1) & nomeM & Mid$(f, p + Len(nomeG))
        p = InStr(p + Len(nomeM), f, "dsG_")
    Loop
    TrocarNomes = f
End Function

' Recria a arvore a partir da faixa copiada (selecao + AddChart2, como em
' modGraficos.GraficoArvore: no Mac o treemap nao aceita SetSourceData).
Private Function RecriarArvore(ByVal wb As Workbook, ByVal wsM As Worksheet, ByVal co As ChartObject, _
        ByVal delta As Long, ByVal lft As Double, ByVal tp As Double) As Object
    Dim ref As String, ancora As Range, fonte As Range, shp As Shape
    ref = Reescrever(wb.Names("dsG_Arvore_T").RefersTo, delta, False, 0, 0)
    ref = Replace(Mid$(ref, 2), "#", "")
    Set ancora = Application.Range(ref)
    Set fonte = ancora
    On Error Resume Next
    Set fonte = ancora.SpillingToRange
    On Error GoTo 0
    If fonte Is Nothing Then Set fonte = ancora
    wsM.Activate
    fonte.Select
    Set shp = wsM.Shapes.AddChart2(-1, XL_TREEMAP, lft, tp, co.Width, co.Height)
    wsM.Range("A1").Select
    On Error Resume Next
    shp.Chart.HasLegend = False
    shp.Chart.HasTitle = True
    On Error GoTo 0
    Set RecriarArvore = wsM.ChartObjects(wsM.ChartObjects.Count)
End Function

'==============================================================================
' Referencias
'==============================================================================

'------------------------------------------------------------------------------
' Reescrever
' Proposito: deslocar em delta colunas as referencias absolutas ($COL$LIN) que
'            apontam para o Dashboard e trocar "Dashboard!" por "'Meus graficos'!".
' Parametros: f - texto da formula; delta - colunas a deslocar;
'             semQualif - True para formulas de celula do Dashboard (referencia
'             sem planilha = Dashboard); minC/maxC - recebem as colunas vistas;
'             trocar - False para so ler (nao troca a planilha).
' Retorno:    o texto reescrito. Trechos entre aspas nao sao tocados.
'------------------------------------------------------------------------------
Private Function Reescrever(ByVal f As String, ByVal delta As Long, ByVal semQualif As Boolean, _
        ByRef minC As Long, ByRef maxC As Long, Optional ByVal trocar As Boolean = True) As String
    Dim i As Long, j As Long, k As Long, n As Long, ch As String, out As String, aspas As Boolean
    Dim letras As String, col As Long, qual As String, doDash As Boolean, anteriorDash As Boolean

    n = Len(f)
    i = 1
    Do While i <= n
        ch = Mid$(f, i, 1)
        If ch = """" Then
            aspas = Not aspas
            out = out & ch
            i = i + 1
        ElseIf ch = "$" And Not aspas Then
            j = i + 1
            Do While j <= n
                If Not EhLetra(Mid$(f, j, 1)) Then Exit Do
                j = j + 1
            Loop
            k = j + 1
            Do While k <= n
                If Not EhDigito(Mid$(f, k, 1)) Then Exit Do
                k = k + 1
            Loop
            If j - i - 1 >= 1 And j - i - 1 <= 3 And j <= n And k > j + 1 Then
                If Mid$(f, j, 1) = "$" Then
                    ' De qual planilha e esta referencia?
                    qual = Qualificador(f, i)
                    If Len(qual) > 0 Then
                        doDash = (StrComp(NomeDoQualificador(qual), mDash, vbTextCompare) = 0)
                    ElseIf i > 1 And Mid$(f, i - 1, 1) = ":" Then
                        doDash = anteriorDash
                    Else
                        doDash = semQualif
                    End If
                    anteriorDash = doDash
                    If doDash Then
                        letras = Mid$(f, i + 1, j - i - 1)
                        col = ColunaDe(letras)
                        If minC = 0 Or col < minC Then minC = col
                        If col > maxC Then maxC = col
                        If Len(qual) > 0 And trocar Then
                            out = Left$(out, Len(out) - Len(qual)) & "'" & Replace(mMeus, "'", "''") & "'!"
                        End If
                        out = out & "$" & LetrasDe(col + delta) & Mid$(f, j, k - j)
                    Else
                        out = out & Mid$(f, i, k - i)
                    End If
                    i = k
                Else
                    out = out & ch
                    i = i + 1
                End If
            Else
                out = out & ch
                i = i + 1
            End If
        Else
            out = out & ch
            i = i + 1
        End If
    Loop
    Reescrever = out
End Function

' So le: junta em minC/maxC as colunas do Dashboard citadas em f.
Private Sub Ler(ByVal f As String, ByVal semQualif As Boolean, ByRef minC As Long, ByRef maxC As Long)
    Reescrever f, 0, semQualif, minC, maxC, False
End Sub

' Texto "Planilha!" ou "'Planilha'!" logo antes da posicao p (ou "").
Private Function Qualificador(ByVal f As String, ByVal p As Long) As String
    Dim i As Long
    If p < 3 Then Exit Function
    If Mid$(f, p - 1, 1) <> "!" Then Exit Function
    If Mid$(f, p - 2, 1) = "'" Then
        i = InStrRev(f, "'", p - 3)
        If i > 0 Then Qualificador = Mid$(f, i, p - i)
    Else
        i = p - 2
        Do While i >= 1
            If Not (EhLetra(Mid$(f, i, 1)) Or EhDigito(Mid$(f, i, 1)) Or Mid$(f, i, 1) = "_" Or Mid$(f, i, 1) = ".") Then Exit Do
            i = i - 1
        Loop
        Qualificador = Mid$(f, i + 1, p - i - 1)
    End If
End Function

' "'Planilha'!" ou "Planilha!" -> "Planilha" (sem o livro entre colchetes).
Private Function NomeDoQualificador(ByVal q As String) As String
    Dim s As String, p As Long
    s = Left$(q, Len(q) - 1)
    If Left$(s, 1) = "'" Then s = Mid$(s, 2, Len(s) - 2)
    s = Replace(s, "''", "'")
    p = InStr(1, s, "]")
    If p > 0 Then s = Mid$(s, p + 1)
    NomeDoQualificador = s
End Function

' Nome definido que comeca na posicao p.
Private Function Identificador(ByVal f As String, ByVal p As Long) As String
    Dim i As Long, ch As String
    i = p
    Do While i <= Len(f)
        ch = Mid$(f, i, 1)
        If Not (EhLetra(ch) Or EhDigito(ch) Or ch = "_" Or ch = ".") Then Exit Do
        i = i + 1
    Loop
    Identificador = Mid$(f, p, i - p)
End Function

Private Function NomeExiste(ByVal wb As Workbook, ByVal nome As String) As Boolean
    Dim nm As Name
    On Error Resume Next
    Set nm = wb.Names(nome)
    NomeExiste = Not nm Is Nothing
End Function

Private Function EhLetra(ByVal ch As String) As Boolean
    EhLetra = (ch >= "A" And ch <= "Z") Or (ch >= "a" And ch <= "z")
End Function

Private Function EhDigito(ByVal ch As String) As Boolean
    EhDigito = (ch >= "0" And ch <= "9")
End Function

Private Function ColunaDe(ByVal letras As String) As Long
    Dim i As Long
    For i = 1 To Len(letras)
        ColunaDe = ColunaDe * 26 + (Asc(UCase$(Mid$(letras, i, 1))) - 64)
    Next i
End Function

Private Function LetrasDe(ByVal col As Long) As String
    Dim s As String, r As Long
    Do While col > 0
        r = (col - 1) Mod 26
        s = Chr$(65 + r) & s
        col = (col - 1) \ 26
    Loop
    LetrasDe = s
End Function
