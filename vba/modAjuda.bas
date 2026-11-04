Attribute VB_Name = "modAjuda"
Option Explicit
Option Private Module   ' rotinas internas: fora da lista de macros
'==============================================================================
' modAjuda
' Cria a aba "Como usar" do modelo .xlsm: passo a passo para quem nao conhece
' VBA. O texto fica aqui (versionado com o codigo); a aba usa os estilos DS_*
' e por isso acompanha a troca de tema.
' Para recriar a aba: no Editor do VBA, janela Imediata, digite
'     modAjuda.CriarAbaComoUsar
' Linhas curtas e sem quebra automatica: a troca de tema redefine os estilos e
' poderia desfazer uma quebra de linha aplicada direto na celula.
'==============================================================================

Private Const PLAN_AJUDA As String = "Como usar"

'------------------------------------------------------------------------------
' CriarAbaComoUsar
' Proposito: criar (ou refazer) a aba "Como usar" como primeira aba do livro.
' Parametros: wb - livro alvo (opcional; padrao = este arquivo).
' Efeitos:    apaga e reescreve o conteudo da aba; ativa a aba ao final.
'------------------------------------------------------------------------------
Public Sub CriarAbaComoUsar(Optional ByVal wb As Workbook)
    Dim ws As Worksheet, c As Collection, i As Long, lin As Long, tipo As String, txt As String

    If wb Is Nothing Then Set wb = ThisWorkbook
    Set ws = modUtil.ObterPlanilha(wb, PLAN_AJUDA)
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(Before:=wb.Sheets(1))
        ws.Name = PLAN_AJUDA
    End If

    ws.Cells.Clear
    ws.Cells.Style = "DS_Texto"
    ws.Columns(1).ColumnWidth = 3
    ws.Columns(2).ColumnWidth = 5
    ws.Columns(3).ColumnWidth = 110

    Set c = Conteudo()
    lin = 2
    For i = 1 To c.Count
        tipo = Left$(c(i), 2)
        txt = modUtil.U(Mid$(c(i), 3))
        Select Case tipo
            Case "T:"                                   ' titulo
                ws.Cells(lin, 2).Value = txt
                ws.Range(ws.Cells(lin, 2), ws.Cells(lin, 3)).Style = "DS_Titulo"
                ws.Rows(lin).RowHeight = 36
            Case "S:"                                   ' frase de abertura
                ws.Cells(lin, 2).Value = txt
            Case "H:"                                   ' titulo de secao
                lin = lin + 1
                ws.Cells(lin, 2).Value = txt
                ws.Range(ws.Cells(lin, 2), ws.Cells(lin, 3)).Style = "DS_Header"
                ws.Rows(lin).RowHeight = 22
            Case "N:"                                   ' passo numerado "n|texto"
                ws.Cells(lin, 2).Value = Split(txt, "|")(0) & "."
                ws.Cells(lin, 2).HorizontalAlignment = xlRight
                ws.Cells(lin, 3).Value = Split(txt, "|")(1)
            Case "B:"                                   ' item com marcador
                ws.Cells(lin, 2).Value = ChrW(&H2022)
                ws.Cells(lin, 2).HorizontalAlignment = xlRight
                ws.Cells(lin, 3).Value = txt
            Case "C:"                                   ' continuacao da linha anterior
                ws.Cells(lin, 3).Value = txt
        End Select
        lin = lin + 1
    Next i

    On Error Resume Next
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ws.Range("A1").Select
    On Error GoTo 0
End Sub

' Texto da aba. Prefixos: T titulo, S abertura, H secao, N passo, B item,
' C continuacao. Acentos em \uXXXX (arquivo ASCII; ver modUtil.U).
Private Function Conteudo() As Collection
    Dim c As New Collection

    c.Add "T:Como usar o Dashboard"
    c.Add "S:Transforma qualquer tabela em um painel com cards e 8 tipos de gr\u00e1fico, com um clique no bot\u00e3o Atualizar."

    c.Add "H:Passo a passo"
    c.Add "N:1|Ao abrir o arquivo, clique em ""Habilitar Macros"". Sem isso, o bot\u00e3o Atualizar n\u00e3o funciona."
    c.Add "N:2|V\u00e1 \u00e0 aba Dados e apague o exemplo: clique em A1, pressione Cmd+A (Ctrl+A no Windows) e depois Delete."
    c.Add "N:3|Cole a sua base a partir da c\u00e9lula A1, com os nomes das colunas na linha 1."
    c.Add "C:Se a base vier de outra planilha, use Editar \u203a Colar Especial \u203a Valores."
    c.Add "N:4|V\u00e1 \u00e0 aba Dashboard e clique no bot\u00e3o Atualizar, no canto superior direito."
    c.Add "C:Sem o bot\u00e3o, use Ferramentas \u203a Macro \u203a Macros, escolha AtualizarTudo e clique em Executar."
    c.Add "N:5|Leia a faixa de avisos abaixo do t\u00edtulo: ela explica por que algum gr\u00e1fico n\u00e3o foi gerado."
    c.Add "N:6|Para mudar as cores, escolha outro tema na c\u00e9lula B3 da aba Config. O painel muda na hora."
    c.Add "N:7|Salve com Cmd+S (Ctrl+S no Windows), mantendo o formato .xlsm para n\u00e3o perder as macros."

    c.Add "H:Como preparar a base"
    c.Add "B:Uma linha por registro e uma coluna por informa\u00e7\u00e3o."
    c.Add "B:Cabe\u00e7alho na linha 1, sem linhas de t\u00edtulo acima da tabela e sem c\u00e9lulas mescladas."
    c.Add "B:Datas guardadas como datas e n\u00fameros como n\u00fameros (n\u00e3o como texto)."
    c.Add "B:Colunas de c\u00f3digo (C\u00f3digo, ID, CNPJ, CPF, CEP) n\u00e3o entram em somas e m\u00e9dias."

    c.Add "H:O que o dashboard mostra"
    c.Add "B:Cards: n\u00famero de registros e soma e m\u00e9dia das 3 primeiras colunas num\u00e9ricas."
    c.Add "B:Colunas e barras agrupadas: comparam categorias (no m\u00e1ximo 7, as maiores)."
    c.Add "B:Colunas e barras empilhadas: mostram de que partes cada categoria \u00e9 feita (at\u00e9 4 partes)."
    c.Add "B:\u00c1rvore hier\u00e1rquica: o peso de cada categoria e das suas partes."
    c.Add "B:Rosca: a participa\u00e7\u00e3o de cada grupo, em categorias de at\u00e9 5 grupos."
    c.Add "B:Pizza: s\u00f3 aparece quando um grupo concentra 50% ou mais do total."
    c.Add "B:Dispers\u00e3o: a rela\u00e7\u00e3o entre as duas primeiras colunas num\u00e9ricas (com 20 pontos ou mais)."
    c.Add "B:Linha: a evolu\u00e7\u00e3o m\u00eas a m\u00eas, quando a base tem uma coluna de data."
    c.Add "B:O t\u00edtulo de cada gr\u00e1fico j\u00e1 diz a conclus\u00e3o. As regras seguem o Guia de Gr\u00e1ficos no Excel."

    c.Add "H:Temas"
    c.Add "B:H\u00e1 4 temas prontos. O tema escolhido em Config!B3 vale para as abas, os cards e os gr\u00e1ficos."
    c.Add "B:Para criar um tema, acrescente uma linha na tabela de temas da aba Config:"
    c.Add "C:nome, tr\u00eas cores em R, G, B (de 0 a 255) e a fonte. O novo tema entra na lista sozinho."

    c.Add "H:Editar um gr\u00e1fico"
    c.Add "N:1|No Dashboard, clique no gr\u00e1fico que voc\u00ea quer editar e depois no bot\u00e3o Fixar gr\u00e1fico."
    c.Add "C:Sem gr\u00e1fico selecionado, o bot\u00e3o mostra a lista dos gr\u00e1ficos para voc\u00ea escolher pelo n\u00famero."
    c.Add "N:2|Uma c\u00f3pia aparece na aba Meus gr\u00e1ficos. Clique nela com o bot\u00e3o direito para trocar o tipo"
    c.Add "C:(Alterar Tipo de Gr\u00e1fico), os dados (Selecionar Dados) ou o formato."
    c.Add "N:3|A c\u00f3pia continua lendo a tabela tblDados e acompanha a base. Atualizar e o tema n\u00e3o mexem nela."
    c.Add "B:Se a nova base n\u00e3o tiver as colunas usadas pela c\u00f3pia, ela fica vazia: apague-a e fixe de novo."
    c.Add "H:Cuidados"
    c.Add "B:A aba Dashboard \u00e9 refeita a cada Atualizar. Para editar um gr\u00e1fico, use o bot\u00e3o Fixar gr\u00e1fico."
    c.Add "B:Gr\u00e1ficos feitos por voc\u00ea n\u00e3o s\u00e3o apagados pelo Atualizar, mas o ideal \u00e9 deix\u00e1-los em outra aba."
    c.Add "B:Depois de mudar os dados, clique em Atualizar para refazer os gr\u00e1ficos e os t\u00edtulos."
    c.Add "B:Quem receber o arquivo sem habilitar as macros v\u00ea o painel, mas n\u00e3o consegue atualizar nem trocar o tema."

    c.Add "H:Problemas comuns"
    c.Add "B:O bot\u00e3o n\u00e3o faz nada: as macros est\u00e3o desabilitadas. Feche, reabra e clique em Habilitar Macros."
    c.Add "B:Aviso ""Nenhum dado encontrado"": a base n\u00e3o come\u00e7a na c\u00e9lula A1 da aba Dados."
    c.Add "B:Uma coluna de n\u00fameros n\u00e3o virou card: os n\u00fameros podem estar guardados como texto."
    c.Add "C:Selecione a coluna e use Dados \u203a Texto para Colunas \u203a Concluir."
    c.Add "B:N\u00e3o apareceu o gr\u00e1fico de linha: a coluna de data est\u00e1 como texto (por exemplo ""06/2026"")."
    c.Add "B:Um gr\u00e1fico n\u00e3o apareceu: veja o motivo na faixa de avisos do Dashboard."

    Set Conteudo = c
End Function
