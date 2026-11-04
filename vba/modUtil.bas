Attribute VB_Name = "modUtil"
Option Explicit
Option Private Module   ' rotinas internas: fora da lista de macros
'==============================================================================
' modUtil
' Funcoes auxiliares: texto Unicode, conversao segura, cores RGB, busca de
' planilhas/tabelas/nomes, referencias estruturadas e controle do estado do
' Application (ScreenUpdating, Calculation, EnableEvents).
' Sem Scripting.Dictionary, FileSystemObject ou Declare: compativel com Mac.
'==============================================================================

Private mNivel As Long                 ' profundidade de chamadas DesligarApp
Private mCalc As XlCalculation         ' modo de calculo salvo na 1a chamada

'------------------------------------------------------------------------------
' U
' Proposito: converter sequencias \uXXXX em caracteres Unicode (ex.: "\u00e3").
' Parametros: s - texto ASCII com escapes.
' Retorno:    texto com acentos.
' Efeitos:    nenhum. Existe porque os .bas sao ASCII puro (importacao no Mac).
'------------------------------------------------------------------------------
Public Function U(ByVal s As String) As String
    Dim i As Long, r As String, h As String
    i = 1
    Do While i <= Len(s)
        h = Mid$(s, i + 2, 4)
        If Mid$(s, i, 2) = "\u" And Len(h) = 4 And EhHex(h) Then
            r = r & ChrW(CLng("&H" & h))
            i = i + 6
        Else
            r = r & Mid$(s, i, 1)
            i = i + 1
        End If
    Loop
    U = r
End Function

Private Function EhHex(ByVal h As String) As Boolean
    Dim i As Long
    For i = 1 To Len(h)
        If InStr(1, "0123456789abcdefABCDEF", Mid$(h, i, 1)) = 0 Then Exit Function
    Next i
    EhHex = True
End Function

'------------------------------------------------------------------------------
' TextoSeguro
' Proposito: converter qualquer valor de celula em texto sem gerar erro.
' Parametros: v - valor (pode ser erro de formula, Null ou Empty).
' Retorno:    texto aparado; "" para erro, Null ou Empty.
'------------------------------------------------------------------------------
Public Function TextoSeguro(ByVal v As Variant) As String
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then Exit Function
    If IsArray(v) Or IsObject(v) Then Exit Function
    TextoSeguro = Trim$(CStr(v))
End Function

'------------------------------------------------------------------------------
' Limitar / CorRGB
' Proposito: montar uma cor com RGB(r, g, b) a partir de valores de celula,
'            limitando cada componente ao intervalo 0..255.
' Parametros: r, g, b - valores (texto invalido vira 0).
' Retorno:    Long de cor RGB.
'------------------------------------------------------------------------------
Public Function Limitar(ByVal v As Variant) As Long
    Dim d As Double
    If IsError(v) Then Exit Function
    If Not IsNumeric(v) Then Exit Function
    d = CDbl(v)
    If d < 0 Then d = 0
    If d > 255 Then d = 255
    Limitar = CLng(d)
End Function

Public Function CorRGB(ByVal r As Variant, ByVal g As Variant, ByVal b As Variant) As Long
    CorRGB = RGB(Limitar(r), Limitar(g), Limitar(b))
End Function

' Componentes de uma cor RGB (Long = R + G*256 + B*65536)
Public Function CompR(ByVal c As Long) As Long
    CompR = c And &HFF&
End Function

Public Function CompG(ByVal c As Long) As Long
    CompG = (c \ &H100&) And &HFF&
End Function

Public Function CompB(ByVal c As Long) As Long
    CompB = (c \ &H10000) And &HFF&
End Function

'------------------------------------------------------------------------------
' Misturar
' Proposito: interpolar duas cores (usado para derivar o fundo dos cards).
' Parametros: c1, c2 - cores RGB; peso - fracao de c2 (0 = c1, 1 = c2).
' Retorno:    cor RGB resultante.
'------------------------------------------------------------------------------
Public Function Misturar(ByVal c1 As Long, ByVal c2 As Long, ByVal peso As Double) As Long
    Misturar = RGB( _
        CLng(CompR(c1) + (CompR(c2) - CompR(c1)) * peso), _
        CLng(CompG(c1) + (CompG(c2) - CompG(c1)) * peso), _
        CLng(CompB(c1) + (CompB(c2) - CompB(c1)) * peso))
End Function

'------------------------------------------------------------------------------
' Luminancia / Contraste
' Proposito: escolher texto claro ou escuro legivel sobre um fundo.
' Parametros: fundo - cor RGB.
' Retorno:    RGB(255,255,255) para fundos escuros; RGB(17,24,39) para claros.
'------------------------------------------------------------------------------
Public Function Luminancia(ByVal c As Long) As Double
    Luminancia = 0.2126 * CompR(c) + 0.7152 * CompG(c) + 0.0722 * CompB(c)
End Function

Public Function Contraste(ByVal fundo As Long) As Long
    If Luminancia(fundo) < 150 Then
        Contraste = RGB(255, 255, 255)
    Else
        Contraste = RGB(17, 24, 39)
    End If
End Function

'------------------------------------------------------------------------------
' CorLegivel
' Proposito: usar a cor desejada como texto apenas se contrastar com o fundo.
' Parametros: desejada, fundo, alternativa - cores RGB.
' Retorno:    desejada, ou alternativa se a diferenca de luminancia for pequena.
'------------------------------------------------------------------------------
Public Function CorLegivel(ByVal desejada As Long, ByVal fundo As Long, ByVal alternativa As Long) As Long
    If Abs(Luminancia(desejada) - Luminancia(fundo)) < 60 Then
        CorLegivel = alternativa
    Else
        CorLegivel = desejada
    End If
End Function

'------------------------------------------------------------------------------
' ObterPlanilha / ObterTabela
' Proposito: localizar sem erro uma planilha pelo nome, ou um ListObject.
' Retorno:    o objeto, ou Nothing se nao existir.
'------------------------------------------------------------------------------
Public Function ObterPlanilha(ByVal wb As Workbook, ByVal nome As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If StrComp(ws.Name, nome, vbTextCompare) = 0 Then
            Set ObterPlanilha = ws
            Exit Function
        End If
    Next ws
End Function

Public Function ObterTabela(ByVal ws As Worksheet, ByVal nome As String) As ListObject
    Dim lo As ListObject
    For Each lo In ws.ListObjects
        If StrComp(lo.Name, nome, vbTextCompare) = 0 Then
            Set ObterTabela = lo
            Exit Function
        End If
    Next lo
End Function

' Procura um ListObject pelo nome em todas as planilhas do livro.
Public Function ObterTabelaNoLivro(ByVal wb As Workbook, ByVal nome As String) As ListObject
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        Set ObterTabelaNoLivro = ObterTabela(ws, nome)
        If Not ObterTabelaNoLivro Is Nothing Then Exit Function
    Next ws
End Function

'------------------------------------------------------------------------------
' RefColuna
' Proposito: montar referencia estruturada segura para uma coluna de tabela.
' Parametros: tabela - nome do ListObject; coluna - cabecalho da coluna.
' Retorno:    texto no formato tabela[[coluna]], com ' [ ] # escapados por '.
' Efeitos:    nenhum. O formato com colchetes duplos aceita espacos e simbolos.
'------------------------------------------------------------------------------
Public Function RefColuna(ByVal tabela As String, ByVal coluna As String) As String
    Dim c As String
    c = Replace(coluna, "'", "''")
    c = Replace(c, "[", "'[")
    c = Replace(c, "]", "']")
    c = Replace(c, "#", "'#")
    RefColuna = tabela & "[[" & c & "]]"
End Function

'------------------------------------------------------------------------------
' RefNomeLivro
' Proposito: referenciar um nome definido com o prefixo do livro, formato
'            exigido pelas series de graficos ('Livro.xlsx'!nome).
'------------------------------------------------------------------------------
Public Function RefNomeLivro(ByVal wb As Workbook, ByVal nome As String) As String
    RefNomeLivro = "'" & Replace(wb.Name, "'", "''") & "'!" & nome
End Function

'------------------------------------------------------------------------------
' TextoFormula
' Proposito: transformar texto em literal de formula ("" escapado).
'------------------------------------------------------------------------------
Public Function TextoFormula(ByVal s As String) As String
    TextoFormula = """" & Replace(s, """", """""") & """"
End Function

'------------------------------------------------------------------------------
' LivroAlvo
' Proposito: definir em qual livro as rotinas atuam. Funciona tanto no modelo
'            .xlsm (o proprio livro esta ativo) quanto no suplemento .xlam
'            (atua no livro que o usuario esta usando).
' Retorno:    ActiveWorkbook, ou Nothing se nao houver livro utilizavel.
'------------------------------------------------------------------------------
Public Function LivroAlvo() As Workbook
    Dim wb As Workbook
    On Error Resume Next
    Set wb = ActiveWorkbook
    On Error GoTo 0
    If wb Is Nothing Then Exit Function
    If wb.IsAddin Then Exit Function
    Set LivroAlvo = wb
End Function

'------------------------------------------------------------------------------
' DesligarApp / RestaurarApp
' Proposito: desligar ScreenUpdating, Calculation automatico e eventos durante a
'            execucao e restaurar o estado original depois, inclusive apos erro.
' Efeitos:    chamadas aninhadas sao contadas; so a ultima RestaurarApp
'             devolve o estado salvo pela primeira DesligarApp.
'------------------------------------------------------------------------------
Public Sub DesligarApp()
    If mNivel = 0 Then mCalc = Application.Calculation
    mNivel = mNivel + 1
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
End Sub

Public Sub RestaurarApp()
    On Error Resume Next
    If mNivel > 0 Then mNivel = mNivel - 1
    If mNivel = 0 Then
        ' mCalc = 0 significa que o estado se perdeu (projeto VBA reiniciado)
        If mCalc = 0 Then mCalc = xlCalculationAutomatic
        Application.Calculation = mCalc
        Application.EnableEvents = True     ' nunca deixar eventos desligados
        Application.ScreenUpdating = True
        Application.StatusBar = False
    End If
End Sub
