# ImportaExcel

Ferramenta em Delphi XE7 para importar planilhas (Excel ou CSV) para a base Firebird do **ADMERP**. Cada linha da planilha é validada antes de ser gravada. A gravação pode ser feita direto no banco `.FDB` ou num script `.SQL`/`.TXT` para executar depois.

## Tipos de importação

| Tipo | Tabela principal | Observação |
|---|---|---|
| Clie/Forn | `CLIEFORN` | Cadastra também o grupo de cliente (`GRUPO_CLIENTE`) quando informado pela descrição |
| Transportadoras | `TRAN` | Cadastro principal e placas; a aba de adicionais CT-e/MDF-e não é importada |
| Grupos | `GRUP_PROD` | |
| SubGrupos | `SUB_GRUP_PROD` | |
| Marcas | `MARCA` | |
| Produtos | `PROD` | Inclui tributos (`PROD_TRIBUTOS`), custos e estoque |
| Grades | `GRADE`, `GRADE_LINHA`, `GRADE_COLUNA` | Liga o produto à grade e lança o estoque por tamanho/cor (`GRADE_LANCAMENTO`, `GRADE_ESTOQUE`, ajuste em `GRAD_AJUS` e `PROD_AJUS`). Os produtos precisam ter sido importados antes |
| Títulos a Pagar | `TITUP` | Inclui baixas (`BTITUP`); pode cadastrar o fornecedor quando não existir |
| Títulos a Receber | `TITUR` | Inclui baixas (`BTITUR`); pode cadastrar o cliente quando não existir |
| CUSTOM | informada na tela | Importa colunas com o mesmo nome dos campos da tabela informada em **Nome tabela CUSTOM** |

As colunas aceitas por cada tipo, na ordem recomendada, aparecem em **Carregar > Colunas**. Na lista, `Ctrl+C` copia o nome da coluna selecionada.

## Como usar

1. Em **Dados Origem**, selecione a planilha (`.xlsx`, `.xls` ou `.csv`). A primeira linha precisa ser o cabeçalho com os nomes das colunas.
2. Em **Dados Destino**, informe a base `.FDB` ou um arquivo `.SQL`/`.TXT`. O botão **Usar caminho padrão** aponta para `C:\ADMERP\DADOS.FDB`.
3. Escolha o tipo de importação.
4. Confira e ajuste os dados na grade, se precisar (menu **Editar** e clique direito na grade).
5. Clique em **Importar**.

### Validação e erros

- A importação é linha a linha, a partir de **Iniciar da linha**.
- Se uma linha tiver valor inválido, a importação para nela e mostra todos os problemas da linha de uma vez. As linhas anteriores ficam gravadas.
- O campo **Iniciar da linha** fica na linha que parou: corrija o valor e clique em **Importar** de novo para continuar dali, sem duplicar as anteriores.
- Colunas com nome não reconhecido geram um aviso antes de começar, com a opção de continuar ignorando essas colunas.
- Ao final, o generator da tabela é ajustado para o maior código existente (no script, o comando vai no final do arquivo).

### Modo Update

Para alterar cadastros existentes em vez de incluir novos, clique com o botão direito no cabeçalho da coluna que será a chave e escolha **Update**. A coluna fica em azul e aparece em **Colunas Update**. Cada linha altera os registros que correspondem à chave; as demais colunas são os campos atualizados.

### Grades

Uma planilha só, com uma linha por combinação do produto (tamanho, ou tamanho e cor):

| Coluna | Uso |
|---|---|
| PROD | Código do produto (obrigatória) |
| EMPR | Empresa onde o estoque aparece (padrão 1) |
| GRADE | Código ou descrição da grade. Pela descrição, a grade é cadastrada se não existir |
| TIT_LINHA / TIT_COLUNA | Títulos usados ao cadastrar a grade (ex.: TAMANHO e COR) |
| LINHA | Tamanho (obrigatória) |
| COLUNA | Cor (opcional). Sem ela, a grade é de uma dimensão e usa a coluna UNICO |
| QTD ou EST / MAX | Saldo final da combinação: QTD vale para o estoque fiscal e o físico; EST só fiscal e MAX só físico |

- O importador lança só a diferença até o saldo da planilha: reimportar a mesma planilha não duplica o estoque. Quantidade vazia não mexe no estoque e negativo é aceito.
- O estoque entra como um ajuste (um `GRAD_AJUS` e um `PROD_AJUS` por importação), e o estoque total do produto é a soma da grade.
- Produto sem grade que já tem estoque é barrado (zere o estoque antes de vincular a grade); produto ligado a outra grade também.
- Tamanho ou cor novos são cadastrados na grade, e todos os produtos da grade ganham a combinação com estoque zero.
- Não usa colunas Update. No script, as linhas que o banco `.FDB` barraria não alteram nada.

### Destino: FDB ou script

| | Banco `.FDB` | Script `.SQL` / `.TXT` |
|---|---|---|
| Gravação | Direto na base, linha a linha | Gera os comandos para executar depois |
| Cidades | Localizadas na tabela `CIDADE` durante a importação | Resolvidas por subconsulta na execução do script |
| Códigos repetidos | Verificados durante a importação | Recusados pelo banco na execução do script |

**Faça backup da base antes de importar no FDB.** A gravação não pode ser desfeita automaticamente.

## Menus

| Menu | Função |
|---|---|
| Carregar > Cabeçalho | Carrega na grade o cabeçalho de outra planilha |
| Carregar > Colunas | Lista as colunas aceitas pelo tipo selecionado |
| Editar | Adicionar/deletar coluna e linha, procurar, substituir e dividir |
| Limpar | Apaga os dados de um tipo na base `.FDB` (Clie/Forn, Transportadoras, Grupos, SubGrupos, Marcas, Produtos, Grades, Títulos a Pagar e a Receber). Pede confirmação e reinicia o generator. Limpar Produtos apaga também o estoque e os lançamentos das grades. Limpar Grades mantém os produtos: tira a grade deles e apaga o cadastro das grades, o estoque por tamanho e cor e os ajustes feitos pela importação de grades (o estoque total volta ao de antes da importação) |

Clique direito na grade: em uma célula, preenche a coluna copiando o valor ou em série; no cabeçalho, mescla, copia ou marca a coluna como Update.

## Requisitos

- Windows com o client do Firebird (a base do ADMERP usa Firebird 3).
- Microsoft Excel instalado para abrir `.xls`/`.xlsx`. Para `.csv` (separado por `;`) não é necessário.
- Acesso ao servidor Firebird com o usuário configurado no importador, quando o destino for `.FDB`.

## Compilação

Abra `ImportaExcel.dproj` no **Delphi XE7** (RAD Studio 15.0) e compile, ou pela linha de comando:

```bash
DCC32.EXE -B ImportaExcel.dpr
```

O acesso ao banco usa dbExpress (`Data.DBXFirebird`).

## Estrutura do código

| Arquivo | Conteúdo |
|---|---|
| `importa_excel.pas` | Tela principal: carga da planilha, validações gerais, menus e o laço de importação |
| `Colunas.pas` | Lista de colunas aceitas por tipo (tela Colunas e aviso de colunas não reconhecidas) |
| `Units/uImportaClieForn.pas` | Importação de clientes e fornecedores |
| `Units/uImportaTran.pas` | Importação de transportadoras |
| `Units/uImportaGrupo.pas`, `uImportaSubGrupo.pas`, `uImportaMarca.pas` | Grupos, subgrupos e marcas |
| `Units/uImportaProduto.pas` | Importação de produtos e tributos |
| `Units/uImportaGrade.pas` | Importação de grades (tamanho e cor) e do estoque por combinação |
| `Units/uImportaTituP.pas`, `uImportaTituR.pas` | Títulos a pagar e a receber |
| `Units/uImportaCUSTOM.pas` | Importação para tabela informada pelo usuário |
| `Units/uUtil.pas` | Leitura de Excel/CSV e funções auxiliares |
| `Units/uProcurar.pas`, `uSubstituir.pas`, `uDividir.pas` | Ferramentas de edição da grade |
| `empresa.pas`, `importando.pas` | Tela de dados da empresa e tela de andamento |
