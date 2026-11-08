# Dashboard no Excel com Design System

O usuário cola uma base tabular em **Dados!A1** e clica em **Atualizar**. A macro faz quatro coisas:

1. converte a base na tabela `tblDados`;
2. classifica cada coluna;
3. monta cards e gráficos na planilha **Dashboard**;
4. aplica o tema escolhido em **Config!B3**.

Requisito: Microsoft 365 (Mac ou Windows). As fórmulas usam `LET`, `UNIQUE`, `FILTER`, `SORTBY` e `TAKE`.

## Uso rápido: `DashboardGenerator.xlsm` (já montado)

O arquivo abre na aba **Como usar**, que traz este passo a passo dentro do próprio Excel.

1. Abra `DashboardGenerator.xlsm` e clique em **Habilitar Macros** quando o Excel perguntar.
2. Na planilha **Dados**, apague o exemplo e cole a sua base a partir de **A1**, com os nomes das colunas na linha 1.
3. Na planilha **Dashboard**, clique em **Atualizar**.
4. Para trocar o tema, escolha outro na lista em **Config!B3**.
5. Para editar um gráfico do seu jeito, clique nele e depois em **Fixar gráfico** (veja [Editar os gráficos](#editar-os-gráficos-fixar-gráfico)).

O arquivo vem com as 240 linhas fictícias de exemplo e o dashboard já gerado. O roteiro A, mais abaixo, só é necessário para montar o arquivo de novo a partir dos `.bas`, por exemplo depois de alterar o código.

## Conteúdo

| Arquivo | Para quê |
|---|---|
| `vba/modConfig.bas` | Constantes, tabela de temas, criação da estrutura e leitura do Tema Ativo |
| `vba/modDados.bas` | `tblDados`, renomeação de cabeçalhos e **regra de classificação** (no topo do módulo) |
| `vba/modTema.bas` | `AplicarTema`: estilos `DS_*` e cores de Shapes e gráficos |
| `vba/modDashboard.bas` | `AtualizarTudo`: cards, avisos e botões |
| `vba/modFixar.bas` | `FixarGrafico`: copia um gráfico gerado para a aba **Meus gráficos**, onde o usuário pode editá-lo |
| `vba/modAjuda.bas` | Cria a aba **"Como usar"** do `.xlsm` (o texto do passo a passo fica neste módulo) |
| `vba/modGraficos.bas` | **Os 8 tipos de gráfico automáticos** e as regras do guia de gráficos |
| `vba/modUtil.bas` | Funções auxiliares: cores RGB, nomes, estado do Application |
| `vba/Config_Worksheet_Change.txt` | Evento da planilha Config (só no modelo `.xlsm`) |
| `vba/clsEventosApp.cls` + `vba/EstaPastaDeTrabalho_xlam.txt` | Mesmo evento, para o suplemento `.xlam` |
| `DashboardGenerator.xlam` | **Suplemento pronto e testado**, com a guia Dashboard na Faixa de Opções |
| `DashboardGenerator.xlsm` | **Pronto para uso:** módulos importados, evento da Config, botões Atualizar e Fixar gráfico e dashboard de exemplo |
| `Dashboard_base.xlsx` | Config com 4 temas e a validação do Tema Ativo, Dados com **240 linhas fictícias de exemplo** (com a coluna Canal, em que "Digital" domina, para a pizza aparecer) e Dashboard vazio |
| `scripts/adicionar_faixa_xlam.py` | Opcional: cria a guia "Dashboard" na Faixa de Opções do `.xlam` |

Os `.bas` são **ASCII puro**. Os acentos dos textos exibidos ao usuário estão escritos como `\u00e1` e são convertidos por `modUtil.U`. Os comentários foram escritos sem acento de propósito: o Editor do VBA no Mac costuma corromper acentos na importação.

> **Importante no Mac:** o Editor do VBA **não importa** arquivos de uma pasta cujo caminho tenha acento. A pasta deste projeto (`análise-de-dados`) tem acento, e a importação falha com "Erro de Sistema &H80004005". Antes de importar, copie a pasta `vba/` para um caminho sem acento, por exemplo `~/Documents/vba_dashboard/`.

## A. Modelo `.xlsm` (uma pasta de trabalho)

1. Abra `Dashboard_base.xlsx` no Excel.
2. **Ferramentas › Macro › Editor do Visual Basic** (ou ⌥ F11).
3. No Editor, use **Arquivo › Importar Arquivo…** e importe, um de cada vez (o diálogo não aceita seleção múltipla), os 8 arquivos `.bas` da pasta `vba/`. Depois de rodar `AtualizarTudo`, digite `modAjuda.CriarAbaComoUsar` na janela Imediata do Editor para criar a aba "Como usar".
4. No painel do projeto, clique duas vezes na planilha **Config** e cole o conteúdo de `Config_Worksheet_Change.txt`.
5. **Depurar › Compilar VBAProject**. Não deve aparecer nenhum erro.
6. Volte ao Excel e use **Arquivo › Salvar Como…** com o formato **Pasta de Trabalho Habilitada para Macro do Excel (.xlsm)**.
7. **Ferramentas › Macro › Macros…**, escolha `AtualizarTudo` e clique em **Executar**. O Dashboard é montado e ganha os botões **Atualizar** e **Fixar gráfico**, já ligados às macros.
8. Para trocar o tema: escolha outro em **Config!B3**.
9. Para trocar a base: apague o conteúdo de **Dados**, cole a sua base em A1 (cabeçalho na linha 1) e clique em **Atualizar**.

Os botões são criados pela própria macro. Se precisar religar um deles: clique com o botão direito nele, depois **Atribuir Macro…** › `AtualizarTudo` (ou `FixarGrafico`).

## B. Suplemento `.xlam` (qualquer pasta de trabalho)

**Já montado:** `DashboardGenerator.xlam`, com a guia "Dashboard" na Faixa de Opções.

- **Para usar numa sessão:** abra o arquivo com dois cliques e clique em **Habilitar Macros**. A guia aparece em todas as pastas de trabalho abertas.
- **Para carregar sempre que o Excel abrir:** use **Ferramentas › Suplementos do Excel… › Procurar…** e marque o arquivo.

Os passos abaixo servem só para remontar o suplemento a partir dos `.bas`.

1. Crie uma pasta de trabalho em branco e abra o Editor do VBA (⌥ F11).
2. Importe os 7 arquivos `.bas` (todos menos o `modAjuda.bas`) e o `clsEventosApp.cls`.
3. Clique duas vezes em **EstaPastaDeTrabalho** e cole o conteúdo de `EstaPastaDeTrabalho_xlam.txt`. **Não** use `Config_Worksheet_Change.txt` aqui.
4. Rode **Depurar › Compilar VBAProject**.
5. **Arquivo › Salvar Como…** › formato **Suplemento do Excel (.xlam)**, com o nome `DashboardGenerator.xlam`.
6. Opcional, para ter a guia na Faixa de Opções: feche o Excel e rode o comando abaixo. Ele gera o arquivo `DashboardGenerator_faixa.xlam`.

```bash
python3 scripts/adicionar_faixa_xlam.py DashboardGenerator.xlam
```

7. **Ferramentas › Suplementos do Excel…** › **Procurar…**, escolha o `.xlam` e marque-o.
8. Abra a pasta de trabalho com os dados, com a planilha de dados ativa, e clique em **Dashboard › Atualizar dashboard**. Sem a Faixa de Opções: abra **Ferramentas › Macro › Macros…**, digite `AtualizarTudo` e clique em **Executar**.

No primeiro uso, se a pasta de trabalho não tiver a planilha `Dados`, a macro pergunta se pode usar a planilha ativa e renomeá-la para **Dados**. Depois cria **Config** e **Dashboard**.

**Compartilhamento (e-mail, OneDrive, rede):** cards e gráficos são fórmulas e gráficos comuns do Excel. Quem recebe vê o dashboard e os números se recalculam sem macro, inclusive no Excel para a Web. Só **Atualizar** (reconstruir) e a troca de tema dependem do suplemento ou do `.xlsm`.

## O que a macro gera

- **Classificação de colunas.** A macro examina até 500 células não vazias por coluna, nas primeiras 2.000 linhas:
  - 80% ou mais com datas: a coluna é de **data**;
  - 80% ou mais com números: a coluna é **numérica**, com duas exceções:
    - cabeçalho com a palavra código, cod, id, CNPJ, CPF, CEP ou matrícula: a coluna é **identificador** e não entra em somas nem médias;
    - só números inteiros, com até 25 valores distintos que somam no máximo 5% das células, **e** cabeçalho com cara de código (sigla de até 4 letras maiúsculas, como TC e TI, ou as palavras tipo, classe, categoria, grupo, segmento, nível ou faixa): a coluna é um **código de classe** e conta como **categórica**. Quantidades como "Itens" continuam numéricas;
  - nenhuma célula preenchida: a coluna é **vazia** e é ignorada;
  - qualquer outro caso: a coluna é **categórica**.

  O tipo vem do valor da célula, não do formato.
- **Cards.**
  - **Registros**: `=ROWS(tblDados)`.
  - **Soma** e **Média** para até 3 colunas numéricas. Usam `AGGREGATE`, que ignora células com erro.
- **Gráficos: 8 tipos automáticos mais a linha no tempo**, seguindo o *Guia de Gráficos no Excel*. Cada gráfico só é gerado quando os dados cumprem a regra dele. Quando não cumprem, o Dashboard diz o motivo na faixa de avisos.

| Gráfico | Dados usados | Regra do guia |
|---|---|---|
| Colunas agrupadas | 1ª categoria × 1 a 3 colunas numéricas **da mesma escala** | Até 7 categorias (as 7 maiores); eixo em zero; com uma série, a maior em destaque e as demais em cinza, com rótulos diretos |
| Barras agrupadas | 2ª categoria × as mesmas medidas | Horizontais, da maior (no topo) para a menor; até 7 categorias |
| Colunas empilhadas | 1ª categoria × partes de outra categoria | Até 4 partes (as 3 maiores + "Outros"), do tom escuro ao claro |
| Barras empilhadas | Mesmos dados, em 100% | "Empilhada 100%": compara só a composição |
| Árvore hierárquica | 1ª categoria > partes | Fora do guia: segue as regras gerais, com o grupo principal em tons da cor primária e os demais em cinza |
| Rosca | Categoria com 2 a 5 grupos | Até 5 fatias, maior fatia às 12 h, furo de 60%, nome e porcentagem nos rótulos |
| Pizza | Categoria com 2 a 5 grupos em que **uma fatia tem 50% ou mais** | "Só quando uma fatia domina" |
| Dispersão (XY) | 1ª × 2ª coluna numérica | 20 pontos ou mais; linha de tendência; títulos dos eixos |
| Linha | Soma mensal pela coluna de data | Evolução no tempo, com o último ponto destacado |

  Regras que valem para todos:
  - o **título diz a conclusão** e é calculado por fórmula, por exemplo "Sudeste lidera em Receita (24% do total)", "Quanto maior Receita, maior Quantidade (relação forte)" ou "Receita subiu 25% no último mês";
  - não há 3D;
  - a fonte tem 12 pt;
  - uma cor de destaque para o que importa e cinza para o contexto.

  Como as colunas são escolhidas:
  - **partes** e **fatias** dão preferência a categorias de texto ("Tipo de Controle") sobre códigos numéricos (TC = 1, 2, 3);
  - medidas de **escalas diferentes** (por exemplo Receita e Quantidade) não se misturam no mesmo gráfico agrupado.

  Como os grupos são somados (`SUMIFS` montado por `modGraficos.Agg` e `Igual`):
  - a categoria é comparada pelo texto exato, sem diferenciar maiúsculas: `<18`, `>=60` e `Kit*` são lidos como texto. Sem isso, o `SUMIFS` lê `<18` como "menor que 18" e `*` como curinga;
  - células com erro (`#N/D` etc.) na coluna numérica ficam de fora, pelo critério `">-1E+307"` na própria medida. Os cards fazem o mesmo com `AGGREGATE`, então o total dos gráficos bate com o dos cards.
- **De onde os gráficos leem.** Da área auxiliar na coluna S em diante, que tem fórmulas de matriz dinâmica. Essa área é exposta pelos nomes `dsG_*` (por exemplo `=Dashboard!$S$5#`), então os gráficos crescem com `tblDados`.
- **Estilos.** `DS_Texto`, `DS_Header`, `DS_Card`, mais três estilos extras, cada um com motivo:
  - `DS_CardRotulo`: hierarquia entre o rótulo e o valor do card;
  - `DS_Titulo`: destaque do título;
  - `DS_Aviso`: mensagens que quebram linha.

  Trocar o tema **redefine** esses estilos, e todas as células que os usam mudam juntas.

## Editar os gráficos (Fixar gráfico)

Os gráficos do Dashboard são refeitos a cada **Atualizar**, então uma edição feita neles se perde. Para editar um gráfico do seu jeito:

1. No Dashboard, clique no gráfico e depois no botão **Fixar gráfico** (no suplemento: **Dashboard › Fixar gráfico**, na Faixa de Opções). Sem gráfico selecionado, o botão mostra a lista dos gráficos e pede o número.
2. Uma cópia aparece na aba **Meus gráficos**, criada na primeira vez. Clique nela com o botão direito para trocar o tipo (**Alterar Tipo de Gráfico**), as séries (**Selecionar Dados**) ou o formato.

A cópia é do usuário:

- **Atualizar** e a troca de tema não a apagam nem a repintam;
- ela continua lendo `tblDados`: quando a base muda, os números e o título acompanham.

Como funciona: a macro copia, para a área auxiliar de **Meus gráficos** (coluna S em diante), as fórmulas que alimentam o gráfico, com as referências deslocadas. Depois cria nomes `dsM<n>_*` para elas e religa as séries da cópia a esses nomes. Os nomes `dsG_*` do Dashboard não servem, porque são apagados a cada Atualizar.

Limitações:

- A **árvore** fixada é recriada a partir da faixa copiada, porque o Excel para Mac não religa a série desse tipo de gráfico. Por isso ela sai com as cores padrão do Excel, e não com as do tema.
- Se a nova base não tiver as colunas que a cópia usa, ela fica vazia. Apague a cópia e fixe o gráfico de novo.
- Apagar um gráfico fixado não apaga a faixa auxiliar dele em **Meus gráficos**.

## Temas

São 4 temas: **Corporativo Azul**, **Noite Escura** (o tema escuro), **Verde Floresta** e **Grafite e Coral**.

Para criar um tema, acrescente uma linha em `tblTemas`: nome, três cores em R, G, B (de 0 a 255) e a fonte. A lista do Tema Ativo inclui o novo tema automaticamente. Prefira fontes que existem no Mac e no Windows (Arial, Verdana, Georgia, Calibri).

## Limitações conhecidas

- A planilha **Dashboard** é sempre recriada inteira: não coloque conteúdo manual nela. Os gráficos que o Atualizar apaga são só os que ele mesmo criou (nome começando com `DS_G`); um gráfico feito à mão fica, mas as células embaixo dele são limpas. Para editar um gráfico gerado, use **Fixar gráfico**.
- A base precisa começar em **Dados!A1**, com o cabeçalho na linha 1. Se a planilha de origem tem linhas de título acima da tabela, como a `Base_IFData`, copie só a tabela e cole em Dados!A1 usando **Colar Especial › Valores**.
- As colunas e barras agrupadas usam a **primeira** e a **segunda** coluna categórica. Numa base em que essa coluna é um nome único por linha (por exemplo Instituição), o gráfico vira um ranking "10 maiores" e não um agrupamento como TCB ou tipo de controle.
- Datas guardadas como texto contam como coluna categórica.
- Uma coluna de inteiros com poucos valores diferentes e cabeçalho em sigla curta maiúscula (por exemplo "ANO" ou "QTD") é tratada como código de classe e fica sem cards. Para ela virar medida, renomeie a coluna com um nome por extenso, como "Quantidade".

## O que foi testado (Excel 16.114.1 para Mac, macOS 27.2, pt-BR)

Seguindo o roteiro A: importação, compilação sem erros e `AtualizarTudo` sobre a base de exemplo. Resultado:

- 5 cards;
- gráfico de colunas por Região;
- gráfico de linha mensal.

Outros cenários testados:

- **Troca para "Noite Escura" em Config!B3:** o evento reaplicou o tema em estilos, botão e gráficos.
- **Base vazia:** aparece a mensagem no Dashboard, sem erro.
- **Base só com texto e cabeçalhos vazio e duplicado:** os cabeçalhos viraram `Coluna_2` e `cidade_2`, apareceram o card de contagem, o gráfico de contagem e os avisos.

Duas correções saíram desses testes:

- o eixo dos gráficos agora usa um formato neutro de milhar;
- somas de colunas inteiras não mostram mais `,00`.

A correção do eixo foi validada na janela Imediata do Editor e depois de ponta a ponta, na montagem do `.xlsm` de 26/10/2026 (seção abaixo), já com o `modDashboard.bas` final.

## Teste com a base IFData (Banco Central, junho de 2026)

Base usada: `Analise_IFData_Junho_2026.xlsx`, aba `Base_IFData`, com 1.622 linhas e 25 colunas. A Tabela1 foi colada como valores em Dados!A1 numa cópia; o arquivo original não foi alterado. As capturas de tela estão em `teste_ifdata/`.

- **Antes da correção:** as "3 primeiras numéricas" eram Código, TC e TI, isto é, um identificador e dois códigos de classe. O dashboard mostrou "Soma de Código" e "Média de TC", e o eixo das colunas começava em 93,5 milhões.
- **Depois da correção:** os cards passaram a ser Ativo Total, Carteira de Crédito e Títulos e Valores Mobiliários. O gráfico mostra as 10 maiores instituições por Ativo Total, com o eixo começando em zero.

Conferência com a planilha:

| Indicador | Dashboard | Referência na planilha |
|---|---|---|
| Soma do Ativo Total (R$ mil) | 19.497.752.152 | ANÁLISE_FINAL: 19.497.752.152 |
| Média do Ativo Total | 12.411.045 | ANÁLISE_FINAL: 12.411.045 |
| Top 10 por Ativo Total | BB 2.461.086.914 … Safra 256.683.465 | Cálculos_Apoio!Y15:Z24: mesma ordem e mesmos valores |

A aba `Gráficos_Análise` tem 13 gráficos, quase todos agregados por TCB (8 grupos), tipo de controle (3) ou UF. O gerador automático só reproduz o "Top 10 instituições por Ativo Total". Agrupamentos escolhidos pelo analista, percentuais de participação e o ROE agregado (lucro ÷ patrimônio) não saem da regra padrão; seriam o papel de uma aba de configuração em uma próxima versão.

## Montagem do `DashboardGenerator.xlsm` (26/10/2026)

A montagem seguiu o roteiro A no Excel 16.114.1 para Mac. Ela revelou mais um caso: ao criar o segundo gráfico ou aplicar o tema, o eixo do primeiro gráfico voltava a exibir "250000,0". Agora `AtualizarTudo` reaplica o formato dos eixos no final.

Conferências feitas no arquivo final:

- compilação sem erros;
- `AtualizarTudo` executado pela caixa de Macros, com os dois eixos corretos ("250.000" e "100.000");
- botão **Atualizar** clicado no próprio Dashboard;
- arquivo salvo como `.xlsm`.

## Teste do suplemento `.xlam` (30/10/2026, Excel 16.114.1 para Mac)

O suplemento foi montado com os 5 módulos, a `clsEventosApp` e o código de EstaPastaDeTrabalho, e a Faixa de Opções foi acrescentada pelo script. Ele foi aberto direto, sem instalação em Suplementos do Excel, e testado em `Livro_qualquer_teste.xlsx`: uma pasta de trabalho comum, com uma única aba "Vendas" de 300 linhas e sem Config nem Dados. As capturas de tela estão em `teste_xlam/`.

| Teste | Resultado |
|---|---|
| Guia "Dashboard" na Faixa de Opções (customUI) | Aparece, com os botões Atualizar dashboard e Aplicar tema |
| Atualizar dashboard pela faixa | 1ª tentativa: **falhou**. A aba "Vendas" não foi oferecida e foi criada uma aba Dados vazia. **Corrigido** em `modConfig` (a aba ativa agora é guardada antes de criar a Config) |
| Depois da correção | Pergunta se pode renomear "Vendas" para Dados. Com Sim, cria Config e Dashboard: 300 registros, soma e média do Faturamento, gráfico por Loja e gráfico mensal |
| Botão Atualizar no Dashboard | Executa a macro do suplemento (16:06 → 16:07) |
| Troca para "Noite Escura" em Config!B3 | O evento de aplicação do suplemento reaplicou o tema |
| Fechar sem salvar | O livro de teste e o suplemento foram fechados sem alterar outros arquivos |

O que não foi testado nessa data: o botão **Aplicar tema** da faixa, a instalação permanente em Suplementos do Excel e o Windows. Os dois primeiros foram testados em 08/11/2026 (ver [Testes finais](#testes-finais-de-ponta-a-ponta-08112026-excel-161144-para-mac)); o Windows continua sem teste.

Observação (revista em 08/11/2026): a nota de 04/11 dizia que o Excel não abre o `.xlsm` e o `.xlam` juntos. O teste de 08/11 mostrou outra coisa: com o suplemento instalado e carregado, o `.xlsm` abre normalmente. O que o Excel recusa ("não pode abrir duas pastas de trabalho com o mesmo nome") é abrir um segundo `DashboardGenerator.xlam` enquanto outro já está carregado, que foi o que aconteceu em 04/11, porque a cópia instalada já estava aberta.

## Ajuste da regra de código de classe (30/10/2026)

O ajuste foi feito no `modDados` e aplicado ao `.xlam` e ao `.xlsm`. Testes com o suplemento:

| Base | Resultado |
|---|---|
| Livro com "Itens" (1 a 12) | Os cards **Soma de Itens 1.894** e **Média de Itens 6,31** voltaram, iguais ao valor calculado fora do Excel |
| IFData | TC e TI seguem como categoria e Código como identificador. Cards e Top 10 iguais aos do teste anterior |

O `.xlsm` recebeu o `modDados` novo, compilou e foi salvo. O `AtualizarTudo` não foi rodado de novo nele, porque o dashboard de exemplo salvo não muda: a coluna Quantidade tem 40 valores distintos. Nessa data o `.xlsm` ainda não tinha a correção do `modConfig` (guardar a aba ativa). Em 08/11/2026 o código dos dois arquivos foi extraído e comparado com `vba/`: é o mesmo, já com essa correção.

Durante este teste, o Excel mostrou "O arquivo não pode ser acessado" e desativou a AutoRecuperação até o Excel ser reiniciado. Causa provável: apaguei à mão o arquivo de trava `~$DashboardGenerator.xlam` que o Excel cria na pasta. A causa não foi confirmada, e o aviso não apareceu nos testes anteriores.

## 8 tipos de gráfico automáticos (31/10/2026, Excel 16.114.1 para Mac)

O código novo está em `vba/modGraficos.bas`; `modDashboard`, `modTema`, `modDados` e `modConfig` foram ajustados. Ele foi montado no `.xlsm` e no `.xlam`, e as capturas de tela estão em `teste_graficos/`.

| Teste | Resultado |
|---|---|
| `.xlsm` com a base de exemplo | Os 8 tipos mais a linha, sem avisos, depois de acrescentar a coluna Canal. Sem ela, a pizza não sai: nenhuma fatia domina, e o aviso explica isso |
| Troca para "Noite Escura" | Os 9 gráficos são recoloridos pelo evento da Config |
| `.xlam` num livro comum (aba "Vendas" com 400 linhas) | Os 9 gráficos pela guia Dashboard; a pizza saiu como "App concentra 61% de Faturamento" |
| Troca para "Verde Floresta" pelo suplemento | Os 9 gráficos são recoloridos pelo evento do suplemento |

Problemas encontrados no Excel para Mac e já contornados:

- **Árvore hierárquica (treemap):** o comando que liga os dados ao gráfico (`SetSourceData`) dá erro 445 nesse tipo de gráfico. A macro seleciona a tabela auxiliar e cria o gráfico a partir da seleção, como faz a interface do Excel.
- **Pontos da árvore:** o treemap tem um ponto "pai" por grupo antes dos blocos do grupo (15 linhas + 5 grupos = 20 pontos). A coloração leva isso em conta.
- **Validação do Tema Ativo:** recriar a validação a cada Atualizar falhava (erro 1004) quando a janela do livro estava minimizada. Agora ela só é recriada se estiver faltando.

Limitações:

- A **árvore** lê o intervalo da área auxiliar como estava no último Atualizar. Se o número de grupos mudar, clique em Atualizar.
- O título da árvore não recebe o negrito e a cor do tema, porque o Excel não aplica essa formatação nesse tipo de gráfico.
- Os rótulos de porcentagem da pizza e da rosca podem diferir do título em 1 ponto percentual, por arredondamento.
- A base IFData não foi retestada com os 8 gráficos: o arquivo não está mais em `Documents`.

## Aba "Como usar" (31/10/2026)

O `.xlsm` ganhou uma aba **Como usar**, que é a primeira do arquivo e a que aparece ao abrir. Ela tem sete seções:

- passo a passo em 7 etapas;
- como preparar a base;
- o que cada gráfico mostra;
- temas;
- editar um gráfico (desde 04/11/2026);
- cuidados;
- problemas comuns.

A aba usa os estilos do tema (`DS_Titulo`, `DS_Header`, `DS_Texto`) e por isso muda de cor junto com o Dashboard.

**Para alterar o texto:** edite `vba/modAjuda.bas`, reimporte o módulo e rode `modAjuda.CriarAbaComoUsar` na janela Imediata. A aba é apagada e refeita.

O suplemento `.xlam` não cria essa aba, porque ela pertence ao modelo.

## Editar gráficos: Fixar gráfico (04/11/2026, Excel 16.114.1 para Mac)

Duas mudanças:

- **O Atualizar preserva os gráficos do usuário.** Antes ele apagava todos os gráficos do Dashboard; agora apaga só os que ele criou (`DS_G*`).
- **Botão Fixar gráfico.** O código está em `vba/modFixar.bas`. O botão no Dashboard é criado por `modDashboard`, e a Faixa de Opções do suplemento ganhou o botão com ícone de alfinete.

Os testes foram feitos numa cópia do `.xlsm`. Depois os módulos foram montados no `.xlsm` e no `.xlam`, e os dois compilaram sem erros.

| Teste | Resultado |
|---|---|
| Fixar pela lista (colunas empilhadas: 3 séries com nome ligado ao cabeçalho) | Cópia em **Meus gráficos**, com as séries religadas a `dsM1_Emp_*` e as fórmulas deslocadas da coluna Y para a S |
| Fixar a árvore | Recriada a partir da faixa copiada, com o título ligado |
| Fixar com o gráfico selecionado | Fixa direto, sem mostrar a lista |
| Troca para "Noite Escura" | Dashboard recolorido; as cópias mantêm as cores |
| Atualizar de novo | Cópias intactas; um gráfico criado à mão no Dashboard continua lá |
| Receita de uma linha de Dados alterada | Valores e título da cópia mudaram ("Norte tem o maior total…") sem clicar em Atualizar |
| Botão Fixar gráfico da faixa, numa pasta em branco | Mostra "Este livro ainda não tem Dashboard" |

Encontrado e corrigido durante o teste: a árvore não cria nomes `dsM<n>_*`, e por isso o gráfico fixado depois dela repetia o número. Agora o número também conta os gráficos `Meu<n>_*` que já existem.

## Testes finais de ponta a ponta (08/11/2026, Excel 16.114.4 para Mac)

Objetivo: fechar o que faltava testar (botão **Aplicar tema** da faixa e instalação permanente) e conferir os números dos gráficos, não só se eles aparecem.

**Como os números foram conferidos.** Depois de cada Atualizar, a área auxiliar do Dashboard (linhas 3 a 60, colunas S em diante) e a aba Dados foram lidas do Excel por AppleScript. As somas por grupo, por mês e por categoria × parte foram recalculadas em Python a partir dos mesmos dados e comparadas valor a valor. Bases usadas:

- a base de exemplo do `.xlsm` (240 linhas);
- uma base de 120 linhas com categorias `<18`, `18-29`, `30-59`, `>=60` e produtos `Kit` e `Kit*`;
- uma base de 80 linhas com `=NA()` em duas células da coluna numérica.

### Auditoria antes dos testes

- O código dentro do `.xlsm` e do `.xlam` foi extraído e comparado com `vba/`: era o mesmo, só com diferenças de maiúsculas que o Editor do VBA faz sozinho.
- O `.xlsm` publicado levava uma referência ao painel do suplemento web "Claude for Excel" (`xl/webextensions`, com um identificador de arquivo), gravada ao salvar com o painel aberto. Ela foi removida do pacote; o projeto VBA não foi tocado.

### Erros encontrados e corrigidos

| Erro | Efeito no dashboard | Correção |
|---|---|---|
| Categorias como `<18` e `>=60` viravam comparação no `SUMIFS` | Grupos com soma 0; o título apontava o líder errado ("30-59 lidera em Valor") | `modGraficos.Igual`: critério `"="&texto`, para a categoria ser lida como texto |
| `*` e `?` na categoria viravam curinga | `Kit*` somava também `Kit` (28.058,79 em vez de 13.514,92): "Kit* lidera em Valor (50% do total)" | `Igual` escapa `~`, `*` e `?` quando o texto tem curinga. Testado no Excel para Mac: o `~` só é lido como escape quando o critério tem `*` ou `?` |
| Um `#N/D` na coluna numérica zerava o grupo inteiro | Grupos com erro iam a 0 e para o topo: "Sul lidera em Valor (0% do total)" | `modGraficos.Agg` acrescenta o critério `">-1E+307"` na própria medida, que deixa as células com erro de fora |
| Achado durante a correção: `Application.SumIfs` chamado pelo VBA com o critério `">-1E+307"` deu erro 6 ("Estouro") | Na primeira execução numa pasta com aba a renomear e `#N/D` na medida, o Atualizar parava com "Erro ao montar o dashboard: 6 - Estouro"; a segunda execução funcionava | `modGraficos.Dominante` (regra da pizza) passou a somar os grupos em VBA, numa passada pelas linhas, sem `SumIfs`. As fórmulas das células, que usam o mesmo critério, deram os números certos em todos os testes |

Também mudou a mensagem do **Aplicar tema** numa pasta sem Dashboard: antes pedia para "executar AtualizarTudo"; agora diz "Clique em Atualizar primeiro", como o Fixar gráfico.

Só `modGraficos.bas` e `modTema.bas` mudaram. Os dois foram reimportados no `.xlsm` e no `.xlam`, os projetos compilaram sem erro, e o código extraído dos arquivos finais é igual ao de `vba/`.

### Resultados (versão final)

| Teste | Resultado |
|---|---|
| `.xlsm` com a base de exemplo | 5 cards e 9 gráficos, sem avisos. 66 valores da área auxiliar iguais ao cálculo independente. Cards: 240 registros, Receita 823.011,33 (média 3.429,21) e Quantidade 4.843 (média 20,18), iguais ao cálculo independente |
| Base com `<18`, `>=60` e `Kit*`, pelo suplemento | 9 gráficos, sem avisos; 58 valores iguais. "<18 lidera em Valor (29% do total)" e "Kit lidera em Valor (26% do total)" |
| Base com `#N/D`, pelo suplemento, na primeira execução (com a pergunta de renomear a aba) | 8 gráficos; a pizza não sai porque nenhuma fatia domina, e o aviso diz isso. 45 valores iguais; a soma dos grupos é igual ao card (42.818,53) |
| Botão **Aplicar tema** da faixa | Com os eventos desligados, `Config!B3` foi trocado para "Noite Escura" e o Dashboard continuou azul. O botão aplicou o tema: fundo RGB(17, 24, 39) |
| Botão **Fixar gráfico** da faixa, com o gráfico selecionado | Fixou direto, sem a lista. Cópia `Meu1_ColunasAgrupadas` com séries em `dsM1_*`, título ligado e fórmulas copiadas com as referências deslocadas. Ao somar 1.000 a uma linha de Dados, a cópia passou de 12.431,24 para 13.431,24 sem clicar em Atualizar |
| Instalação permanente (Ferramentas › Suplementos do Excel) | O suplemento instalado carregou ao abrir o Excel, antes de qualquer arquivo. A guia **Dashboard** apareceu numa pasta comum, e o Atualizar gerou o painel com os 58 valores certos. O Excel fechou sem pedir para salvar o suplemento |
| `.xlsm` com o suplemento instalado e carregado | Abriu normalmente, na aba **Como usar**. A troca de tema por `Config!B3` rodou nos dois eventos (o da planilha e o do suplemento) sem erro. A macro do botão Fixar gráfico do modelo fixou a rosca |
| Pasta em branco | **Aplicar tema** e **Fixar gráfico** mostram "Este livro ainda não tem Dashboard. Clique em Atualizar primeiro." |
| `.xlsm` final, sem o painel web | O pacote abre, pede para habilitar as macros e abre na aba **Como usar** |

As capturas de antes e depois estão em `teste_bordas/`.

### Observações do ambiente (Excel para Mac)

- **Acesso ao projeto VBA desligado.** Nesta máquina, "confiar no acesso ao modelo de objeto do projeto VBA" está desligado, então os módulos foram trocados pelo Editor do VBA: **Arquivo › Remover** e **Arquivo › Importar Arquivo…**, de uma pasta sem acento no caminho.
- **Dois menus "Importar".** Com o Editor aberto, a barra de menus pode voltar a ser a da planilha. Aí **Arquivo › Importar** abre o assistente de importação de texto do Excel, não a importação de módulo. Confira se a barra mostra **Depurar** e **Executar** antes de importar.
- **Painel web no `.xlsm`.** Antes de publicar, confira que o `.xlsm` não tem `xl/webextensions`. O Excel grava ali o painel de um suplemento web, como o Claude for Excel, se ele estiver aberto ao salvar.

O que continua sem teste: **Windows**.
