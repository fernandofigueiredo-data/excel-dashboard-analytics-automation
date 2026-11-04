Attribute VB_Name = "modDados"
Option Explicit
Option Private Module   ' rotinas internas: fora da lista de macros
'==============================================================================
' modDados
' Prepara a tabela tblDados a partir do que foi colado em Dados!A1 e classifica
' cada coluna para o dashboard.
'
' REGRA DE CLASSIFICACAO DE COLUNAS
'   Amostra: ate AMOSTRA_MAX celulas NAO vazias de cada coluna, lidas das
'   primeiras AMOSTRA_LINHAS linhas (leitura unica em array, rapida em bases
'   grandes). Celulas com erro (#N/D etc.) sao ignoradas na contagem.
'   Seja n = celulas validas na amostra:
'     n = 0                                   -> "V" vazia (ignorada)
'     >= 80% com valor do tipo Data           -> "D" data
'     >= 80% com valor numerico (nao texto)   -> "N" numerica, exceto:
'        - cabecalho de identificador (palavra codigo, cod, id, cnpj, cpf,
'          cep, matricula)                    -> "I" identificador (ignorado)
'        - so inteiros com ate 25 valores distintos e distintos <= 5% de n,
'          E cabecalho com cara de codigo (sigla de ate 4 letras maiusculas,
'          como TC/TI, ou palavra tipo, classe, categoria, grupo, segmento,
'          nivel, faixa) -> "C" categorica (codigo de classe)
'          Sem a condicao do cabecalho, quantidades como Itens = 1..12 viravam
'          categoria (achado no teste do suplemento).
'     caso contrario                          -> "C" categorica
'   Observacoes:
'     - O tipo vem do VALOR da celula, nao do formato: "12" digitado como texto
'       e "01/02/2026" guardado como texto contam como categoria.
'     - Verdadeiro/Falso conta como categoria.
'     - Datas sao reconhecidas porque o Excel as entrega como Date em .Value.
'     - Somar ou tirar media de codigo nao tem significado; por isso as
'       excecoes acima (encontradas testando a base IFData do Banco Central).
'
' RENOMEACAO DE CABECALHOS (previsivel)
'   - vazio ou com erro          -> "Coluna_<n>"  (n = posicao da coluna)
'   - repetido (sem diferenciar maiusculas) -> "<nome>_2", "<nome>_3", ...
'   - quebras de linha viram espaco; espacos nas pontas sao removidos.
'==============================================================================

Private Const AMOSTRA_LINHAS As Long = 2000
Private Const AMOSTRA_MAX As Long = 500
Private Const LIMIAR As Double = 0.8
Private Const MAX_DISTINTOS_CODIGO As Long = 25
Private Const RAZAO_DISTINTOS_CODIGO As Double = 0.05

'------------------------------------------------------------------------------
' PrepararTabela
' Proposito: localizar a regiao continua a partir de Dados!A1, corrigir os
'            cabecalhos e criar ou redimensionar o ListObject tblDados.
' Parametros: wb  - livro alvo.
'             lo  - (saida) a tabela tblDados, quando houver dados.
'             msg - (saida) explicacao quando nao ha dados utilizaveis.
' Retorno:    True se ha cabecalho e ao menos uma linha de dados.
' Efeitos:    reescreve a linha de cabecalho, desfaz celulas mescladas da
'             regiao, cria/renomeia/redimensiona tblDados e remove o estilo de
'             tabela padrao (o visual vem dos estilos DS_*).
'------------------------------------------------------------------------------
Public Function PrepararTabela(ByVal wb As Workbook, ByRef lo As ListObject, ByRef msg As String) As Boolean
    Dim ws As Worksheet, regiao As Range, outra As ListObject

    Set ws = modUtil.ObterPlanilha(wb, modConfig.PLAN_DADOS)
    Set lo = modUtil.ObterTabelaNoLivro(wb, modConfig.TAB_DADOS)

    If Not lo Is Nothing Then
        If Not lo.Parent Is ws Then
            msg = modUtil.U("J\u00e1 existe uma tabela tblDados na planilha """) & lo.Parent.Name & _
                  modUtil.U(""". Renomeie-a ou mova os dados para a planilha Dados.")
            Set lo = Nothing
            Exit Function
        End If
    End If

    ' Uma tabela com outro nome em A1 e adotada (renomeada para tblDados)
    If lo Is Nothing Then
        If Not ws.Range("A1").ListObject Is Nothing Then
            Set lo = ws.Range("A1").ListObject
            lo.Name = modConfig.TAB_DADOS
        End If
    End If

    Set regiao = ws.Range("A1").CurrentRegion
    If Application.WorksheetFunction.CountA(regiao) = 0 Then
        msg = modUtil.U("Nenhum dado encontrado. Cole a base na planilha Dados a partir da c\u00e9lula A1, " & _
              "com os nomes das colunas na linha 1, e clique em Atualizar.")
        Exit Function
    End If
    If regiao.Row <> 1 Or regiao.Column <> 1 Then Set regiao = ws.Range("A1", regiao.Cells(regiao.Rows.Count, regiao.Columns.Count))

    ' Outra tabela dentro da regiao impede criar tblDados
    For Each outra In ws.ListObjects
        If Not outra Is lo Then
            If Not Application.Intersect(outra.Range, regiao) Is Nothing Then
                msg = modUtil.U("A tabela """) & outra.Name & modUtil.U(""" est\u00e1 encostada nos dados de A1. " & _
                      "Deixe uma linha e uma coluna vazias entre ela e a base.")
                Set lo = Nothing
                Exit Function
            End If
        End If
    Next outra

    If IsNull(regiao.MergeCells) Then
        regiao.UnMerge
    ElseIf regiao.MergeCells Then
        regiao.UnMerge
    End If

    CorrigirCabecalhos regiao.Rows(1)

    If regiao.Rows.Count < 2 Then
        msg = modUtil.U("A base tem s\u00f3 a linha de cabe\u00e7alho. Cole os registros a partir da linha 2.")
        If Not lo Is Nothing Then Set lo = Nothing
        Exit Function
    End If

    If lo Is Nothing Then
        Set lo = ws.ListObjects.Add(xlSrcRange, regiao, , xlYes)
        lo.Name = modConfig.TAB_DADOS
    ElseIf lo.Range.Address <> regiao.Address Then
        lo.Resize regiao
    End If

    On Error Resume Next
    lo.TableStyle = ""        ' sem estilo de tabela: o tema usa DS_Header/DS_Texto
    lo.ShowTableStyleRowStripes = False
    On Error GoTo 0

    PrepararTabela = True
End Function

'------------------------------------------------------------------------------
' CorrigirCabecalhos
' Proposito: aplicar a regra de renomeacao (topo do modulo) a linha 1.
' Parametros: linha - primeira linha da regiao de dados.
' Efeitos:    reescreve a linha inteira de uma vez, formatada como texto.
'------------------------------------------------------------------------------
Private Sub CorrigirCabecalhos(ByVal linha As Range)
    Dim v As Variant, n As Long, j As Long, k As Long
    Dim nome As String, base As String, vistos As Collection

    n = linha.Columns.Count
    If n = 1 Then
        ReDim v(1 To 1, 1 To 1)
        v(1, 1) = linha.Cells(1, 1).Value
    Else
        v = linha.Value
    End If

    Set vistos = New Collection
    For j = 1 To n
        nome = modUtil.TextoSeguro(v(1, j))
        nome = Replace(Replace(Replace(nome, vbCrLf, " "), vbLf, " "), vbCr, " ")
        nome = Trim$(nome)
        If Len(nome) > 200 Then nome = Left$(nome, 200)
        If Len(nome) = 0 Then nome = "Coluna_" & j

        base = nome
        k = 1
        Do While JaExiste(vistos, nome)
            k = k + 1
            nome = base & "_" & k
        Loop
        vistos.Add nome, LCase$(nome)
        v(1, j) = nome
    Next j

    linha.NumberFormat = "@"
    linha.Value = v
End Sub

' Retorna True se a chave ja esta na Collection (chaves nao diferenciam caixa).
Private Function JaExiste(ByVal c As Collection, ByVal chave As String) As Boolean
    Dim x As Variant
    On Error Resume Next
    x = c(LCase$(chave))
    JaExiste = (Err.Number = 0)
    Err.Clear
End Function

'------------------------------------------------------------------------------
' ClassificarColunas
' Proposito: classificar cada coluna de tblDados pela regra do topo do modulo.
' Parametros: lo - tabela tblDados com ao menos uma linha de dados.
' Retorno:    Variant(1 To nColunas) com "N", "D", "C", "I" ou "V".
' Efeitos:    nenhum (somente leitura, uma leitura em array).
'------------------------------------------------------------------------------
Public Function ClassificarColunas(ByVal lo As ListObject) As Variant
    Dim nLin As Long, nCol As Long, i As Long, j As Long
    Dim v As Variant, tipos() As Variant, x As Variant
    Dim n As Long, nD As Long, nN As Long, soInteiros As Boolean
    Dim distintos As Collection

    nCol = lo.ListColumns.Count
    ReDim tipos(1 To nCol)

    nLin = lo.DataBodyRange.Rows.Count
    If nLin > AMOSTRA_LINHAS Then nLin = AMOSTRA_LINHAS
    If nLin = 1 And nCol = 1 Then
        ReDim v(1 To 1, 1 To 1)
        v(1, 1) = lo.DataBodyRange.Cells(1, 1).Value
    Else
        v = lo.DataBodyRange.Resize(nLin).Value
    End If

    For j = 1 To nCol
        n = 0: nD = 0: nN = 0: soInteiros = True
        Set distintos = New Collection
        For i = 1 To nLin
            x = v(i, j)
            If Not IsError(x) And Not IsEmpty(x) Then
                If VarType(x) = vbString Then
                    If Len(x) > 0 Then n = n + 1
                Else
                    n = n + 1
                    Select Case VarType(x)
                        Case vbDate
                            nD = nD + 1
                        Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
                            nN = nN + 1
                            If x <> Fix(x) Then soInteiros = False
                            If distintos.Count <= MAX_DISTINTOS_CODIGO Then
                                If Not JaExiste(distintos, CStr(x)) Then distintos.Add CStr(x), LCase$(CStr(x))
                            End If
                    End Select
                End If
                If n >= AMOSTRA_MAX Then Exit For
            End If
        Next i

        If n = 0 Then
            tipos(j) = "V"
        ElseIf nD / n >= LIMIAR Then
            tipos(j) = "D"
        ElseIf nN / n >= LIMIAR Then
            If EhIdentificador(lo.ListColumns(j).Name) Then
                tipos(j) = "I"
            ElseIf soInteiros And distintos.Count <= MAX_DISTINTOS_CODIGO _
                   And distintos.Count <= RAZAO_DISTINTOS_CODIGO * nN _
                   And PareceCodigoDeClasse(lo.ListColumns(j).Name) Then
                tipos(j) = "C"
            Else
                tipos(j) = "N"
            End If
        Else
            tipos(j) = "C"
        End If
    Next j

    ClassificarColunas = tipos
End Function

'------------------------------------------------------------------------------
' EhIdentificador
' Proposito: reconhecer pelo cabecalho colunas numericas que sao codigos de
'            identificacao (somar ou tirar media delas nao tem significado).
' Parametros: cabecalho - nome da coluna.
' Retorno:    True se alguma palavra do cabecalho, sem acento e em minusculas,
'             for codigo, cod, id, cnpj, cpf, cep ou matricula.
'------------------------------------------------------------------------------
Private Function EhIdentificador(ByVal cabecalho As String) As Boolean
    Dim s As String, i As Long, c As String, palavras As Variant, p As Variant

    s = LCase$(SemAcento(cabecalho))
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        If Not (c Like "[a-z0-9]") Then Mid$(s, i, 1) = " "
    Next i
    palavras = Split(Application.WorksheetFunction.Trim(s), " ")
    For Each p In palavras
        Select Case p
            Case "codigo", "cod", "id", "cnpj", "cpf", "cep", "matricula"
                EhIdentificador = True
                Exit Function
        End Select
    Next p
End Function

'------------------------------------------------------------------------------
' PareceCodigoDeClasse
' Proposito: dizer se o cabecalho de uma coluna de inteiros sugere um codigo de
'            classificacao (e nao uma quantidade).
' Parametros: cabecalho - nome da coluna.
' Retorno:    True para sigla de 1 a 4 caracteres maiusculos/digitos com ao menos
'             uma letra (TC, TI, UF2) ou se alguma palavra, sem acento e em
'             minusculas, for tipo, classe, categoria, grupo, segmento, nivel ou
'             faixa.
'------------------------------------------------------------------------------
Private Function PareceCodigoDeClasse(ByVal cabecalho As String) As Boolean
    Dim t As String, i As Long, c As String, temLetra As Boolean
    Dim s As String, p As Variant

    t = Trim$(cabecalho)
    If Len(t) >= 1 And Len(t) <= 4 Then
        PareceCodigoDeClasse = True
        For i = 1 To Len(t)
            c = Mid$(t, i, 1)
            If c Like "[A-Z]" Then
                temLetra = True
            ElseIf Not (c Like "[0-9]") Then
                PareceCodigoDeClasse = False
                Exit For
            End If
        Next i
        If PareceCodigoDeClasse And temLetra Then Exit Function
        PareceCodigoDeClasse = False
    End If

    s = LCase$(SemAcento(cabecalho))
    For i = 1 To Len(s)
        If Not (Mid$(s, i, 1) Like "[a-z0-9]") Then Mid$(s, i, 1) = " "
    Next i
    For Each p In Split(Application.WorksheetFunction.Trim(s), " ")
        Select Case p
            Case "tipo", "classe", "categoria", "grupo", "segmento", "nivel", "faixa"
                PareceCodigoDeClasse = True
                Exit Function
        End Select
    Next p
End Function

' Remove os acentos mais comuns do portugues (comparacoes de cabecalho).
Private Function SemAcento(ByVal s As String) As String
    Dim de As String, para As String, i As Long
    de = modUtil.U("\u00e1\u00e0\u00e2\u00e3\u00e9\u00ea\u00ed\u00f3\u00f4\u00f5\u00fa\u00fc\u00e7" & _
                   "\u00c1\u00c0\u00c2\u00c3\u00c9\u00ca\u00cd\u00d3\u00d4\u00d5\u00da\u00dc\u00c7")
    para = "aaaaeeiooouucAAAAEEIOOOUUC"
    For i = 1 To Len(de)
        s = Replace(s, Mid$(de, i, 1), Mid$(para, i, 1))
    Next i
    SemAcento = s
End Function

'------------------------------------------------------------------------------
' ValoresDistintos
' Proposito: grupos distintos de uma coluna (sem vazios e erros), comparando
'            sem diferenciar maiusculas, como fazem SUMIFS/COUNTIFS.
' Parametros: lo - tabela; j - coluna; limite - para de contar apos limite+1.
' Retorno:    Collection com o 1o valor encontrado de cada grupo.
' Efeitos:    nenhum (uma leitura em array da coluna).
'------------------------------------------------------------------------------
Public Function ValoresDistintos(ByVal lo As ListObject, ByVal j As Long, ByVal limite As Long) As Collection
    Dim v As Variant, i As Long, x As Variant, k As String, c As Collection

    Set c = New Collection
    Set ValoresDistintos = c
    If lo.DataBodyRange Is Nothing Then Exit Function
    If lo.DataBodyRange.Rows.Count = 1 Then
        ReDim v(1 To 1, 1 To 1)
        v(1, 1) = lo.ListColumns(j).DataBodyRange.Cells(1, 1).Value
    Else
        v = lo.ListColumns(j).DataBodyRange.Value
    End If
    For i = 1 To UBound(v, 1)
        x = v(i, 1)
        If Not IsError(x) And Not IsEmpty(x) Then
            k = LCase$(CStr(x))
            If Len(k) > 0 Then
                If Not JaExiste(c, k) Then
                    c.Add x, k
                    If c.Count > limite Then Exit For
                End If
            End If
        End If
    Next i
End Function
