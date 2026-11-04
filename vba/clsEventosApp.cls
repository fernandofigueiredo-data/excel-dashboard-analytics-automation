VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEventosApp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'==============================================================================
' clsEventosApp  (somente no suplemento .xlam)
' Proposito: o suplemento nao pode colocar codigo na planilha Config de cada
'            livro; por isso escuta as alteracoes de TODAS as planilhas pelo
'            evento de aplicacao e repassa para modTema.TratarMudancaConfig,
'            que so age na planilha Config, no Tema Ativo ou em tblTemas.
' Ligacao:   feita em EstaPastaDeTrabalho (Workbook_Open) do suplemento.
'==============================================================================

Public WithEvents App As Application

Private Sub App_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    modTema.TratarMudancaConfig Sh, Target
End Sub
