# Média Diária de Exportação - MDIC

Rotina automatizada que verifica, baixa e processa o boletim semanal da **Balança Comercial Brasileira** (MDIC/SECEX), filtra os produtos de interesse e grava os indicadores de exportação no banco Oracle.

## Fluxo do pipeline

1. **Verificação de atualização** (`web_scraper.py`)
   Abre via Selenium (Chrome) a página de resultados da balança comercial (`principais_resultados.html`) e extrai a data da última atualização publicada no `h4.date`.

2. **Decisão de download** (`balanca_semanal.py :: verifica_necesidade_download`)
   Compara a data extraída com a data atual. Só segue com o pipeline se a atualização já é válida (`hoje >= data_referencia`); caso contrário, encerra sem baixar nada.

3. **Download do arquivo** (`baixar_arquivo`)
   Baixa a planilha `Setores_Produtos.xlsx` diretamente da URL pública do MDIC (`https://balanca.economia.gov.br/balanca/semanal/Setores_Produtos.xlsx`) e lê os dados com `pandas`/`openpyxl` (pulando as 6 primeiras linhas de cabeçalho).

4. **Padronização das colunas** (`renomear_colunas`)
   Renomeia as 14 primeiras colunas da planilha para nomes legíveis (valor US$, toneladas, preço médio, variações percentuais etc.).

5. **Filtragem dos produtos** (`filtrar_dataframe`)
   Normaliza os textos (remove acentuação/caixa) e filtra apenas os produtos definidos em `main.py`:
   - Algodão em bruto
   - Madeira em bruto
   - Milho não moído, exceto milho doce

6. **Transformação largo → longo** (`transforma_de_largo_para_longo`)
   Usa `melt` para transformar as colunas de indicadores em linhas (`PRODUTO`, `INDICADOR`, `VALOR`), adiciona a coluna `DATA` (ano atual + mês informado) e extrai o número da `SEMANA` a partir do nome do arquivo original da planilha.

7. **Persistência no Oracle** (`banco_oracle.py`)
   Conecta ao Oracle (via `oracledb`, usando Oracle Wallet) e insere os registros em lote na tabela `MEDIA_DIARIA_EXPORTACAO_MDIC` (colunas `PRODUTO`, `INDICADOR`, `VALOR`, `DATA`, `SEMANA`).

> O método `salvar_arquivo` (exportação para `.xlsx` local com rótulo de confidencialidade via MSIP/Excel COM) está presente na classe mas atualmente **desativado** em `main.py` — a persistência é feita apenas no banco.

## Rotina semanal para Excel / Power BI (`extrair_para_excel.py`)

Baixa a divulgação atual da [página de principais resultados](https://balanca.economia.gov.br/balanca/pg_principal_bc/principais_resultados.html), arquiva **todos** os arquivos do site e mantém o histórico de todas as divulgações em `saida/balanca_semanal_mdic.xlsx`, numa **tabela única** pronta para o Power BI.

```bat
extrair-para-excel.bat                  :: uso normal (é o que o agendador chama)
extrair-para-excel.bat --reprocessar    :: reconstrói o Excel inteiro a partir de saida\brutos, sem baixar nada
```

### Como funciona

1. Baixa a página (HTML estático, sem Selenium) e todos os arquivos linkados nela (planilhas e `Nota.pdf`). Salva tudo em `saida/brutos/<ano>-<mes>_semana<n>/`. **Essa pasta é a fonte da verdade**: as planilhas do site têm URL fixa e são substituídas toda semana.
2. Processa cada divulgação arquivada que ainda não está no Excel, além da atual. Uma semana que falhou ou não foi gravada entra sozinha na execução seguinte.
3. Valida antes de gravar:
   - a página e as planilhas são da mesma semana;
   - o layout das planilhas é o esperado;
   - os produtos somam o setor, os setores somam o total da `Tabela_Resumo` e as semanas somam o mês;
   - a variação publicada bate com média diária ÷ base do ano anterior.

   Se algo falhar, aquela divulgação não entra e o erro vai para `saida/extracao.log`.
4. Grava o Excel num arquivo temporário e troca no fim. Reexecutar a mesma semana substitui as linhas dela, sem duplicar. Se a estrutura das tabelas mudar no código, o Excel é reconstruído automaticamente a partir dos brutos.

O código de saída é 1 quando a divulgação atual falha, e o Agendador de Tarefas mostra isso em "Resultado da última execução". **Agende para rodar diariamente, por exemplo às 18h.** O MDIC publica às segundas, entre 15h e 15h30. Rodar a mais não custa nada e cobre atrasos da publicação.

### Abas do Excel

| Aba | Conteúdo |
|---|---|
| `DADOS` | **A tabela.** Uma linha por produto (exportação ou importação) em cada divulgação semanal. Cerca de 575 linhas por semana. |
| `TEXTOS` | Um parágrafo por linha das seções Destaques, Totais e Setores e Produtos. |
| `LEIA-ME` | Regras de uso e o dicionário de todas as colunas. |

`DADOS` e `TEXTOS` são Tabelas do Excel com esses mesmos nomes, que é o que o Power BI lista ao importar.

**Setor e total não são gravados: são a soma dos produtos** (conferido a cada divulgação contra a `Tabela_Resumo` do MDIC). Assim, somar qualquer coluna de valor dentro de uma divulgação nunca conta nada em dobro.

Colunas de `DADOS`, em grupos:

- **Quando**: `DATA_REFERENCIA` (último dia coberto), `DATA_PUBLICACAO`, `ANO`, `MES`, `SEMANA`, `DIAS_UTEIS_ACUMULADO_MES`, `DIAS_UTEIS_SEMANA`, `ULTIMA_DO_MES`, `MAIS_RECENTE`.
- **O quê**: `FLUXO`, `SETOR`, `PRODUTO`.
- **Valor**: `VALOR_ACUMULADO_MES_USD`, `VALOR_SEMANA_USD`, `MEDIA_DIARIA_USD`, `MEDIA_DIARIA_ANO_ANTERIOR_USD`, `VARIACAO_VALOR`.
- **Peso**: as mesmas cinco, em toneladas.
- **Preço**: `PRECO_MEDIO_USD_TON`, `VARIACAO_PRECO`.

Convenções:
- Valores em **US$** (não em US$ mil), para o Power BI exibir "14,25 bi" corretamente. Peso em toneladas.
- Variações em **fração**: 0,285 = 28,5%, sempre contra a média diária do mesmo mês do ano anterior.
- `MEDIA_DIARIA_ANO_ANTERIOR_*` é a base publicada pelo MDIC na mesma planilha, e se refere ao **mês completo** do ano anterior. Por isso a comparação é sempre por média diária.
- `VALOR_SEMANA_USD` é o acumulado da divulgação menos o da anterior do mesmo mês, e fica **vazio** quando a semana anterior não foi coletada.
- Onde o MDIC publica 0% porque não houve embarque no mês, a coluna traz **−100%**; onde não existe base no ano anterior, fica **vazia**.

### Montando no Power BI

1. Obter dados → Excel → marque as Tabelas `DADOS` e `TEXTOS`.
2. Não precisa de relacionamento nem de tabela de dimensão: é uma tabela só. Use `DATA_REFERENCIA` como eixo de tempo.
3. Formate as colunas `VARIACAO_*` como porcentagem e marque como **Não resumir**: variação não se soma nem se tira média.
4. Para a variação de qualquer agrupamento (setor, total, ou uma lista de produtos sua), crie uma medida:

   ```dax
   Variação = DIVIDE( SUM(DADOS[MEDIA_DIARIA_USD]), SUM(DADOS[MEDIA_DIARIA_ANO_ANTERIOR_USD]) ) - 1
   ```

   Saldo e corrente de comércio saem de `CALCULATE` filtrando `FLUXO`.
5. Cada divulgação é uma foto do mês até aquela semana. Por isso:
   - painel da situação atual: filtre `MAIS_RECENTE = Verdadeiro`;
   - comparar meses: filtre `ULTIMA_DO_MES = Verdadeiro`;
   - evolução dentro do mês: use `MEDIA_DIARIA_USD` no eixo `DATA_REFERENCIA`;
   - somar semanas: use `VALOR_SEMANA_USD`, nunca `VALOR_ACUMULADO_MES_USD` de divulgações diferentes do mesmo mês.

## Estrutura de arquivos

```
media-diaria-exportacao-mdic/
├── main.py                          # Orquestra o pipeline (entry point)
├── extrair_para_excel.py            # Rotina semanal: arquiva o site e atualiza o Excel (Power BI)
├── extrair-para-excel.bat           # Execução da rotina semanal (agendador de tarefas)
├── media-diaria-exportacao-mdic.bat # Script de execução do fluxo antigo (Oracle)
├── .env                             # Credenciais/config do Oracle (não versionado)
└── src/
    ├── web_scraper.py               # Selenium: extrai data da última atualização do MDIC
    ├── balanca_semanal.py           # Download, parsing, filtro e transformação da planilha
    ├── principais_resultados.py     # Parsing e validação da página e das planilhas de uma divulgação
    └── banco_oracle.py              # Conexão e inserção dos dados no Oracle
```

## Execução

O `.bat` navega até o diretório do projeto e roda o script principal:

```bat
cd /d "C:\...\3.Rotinas-atuomacao\tasks\media-diaria-exportacao-mdic"
python main.py
```

Indicado para ser chamado por um agendador de tarefas (Task Scheduler do Windows).

## Configuração (.env)

O arquivo `.env` (não incluído neste README) deve conter as variáveis usadas em `banco_oracle.py`:

| Variável            | Descrição                                          |
|---------------------|-----------------------------------------------------|
| `ORACLE_LIB_DIR`     | Diretório do Oracle Instant Client                  |
| `DB_USER`            | Usuário do banco Oracle                             |
| `DB_PASSWORD`        | Senha do banco Oracle                               |
| `DB_DSN`             | DSN/alias de conexão (ex.: TNS do Oracle Wallet)     |
| `ORACLE_WALLET_DIR`  | Diretório do Oracle Wallet (`config_dir`)            |

## Dependências principais

- `selenium` + `webdriver-manager` (extração da data de atualização)
- `pandas` + `openpyxl` (leitura/transformação da planilha)
- `unidecode` (normalização de texto para o filtro)
- `requests` (download do arquivo)
- `oracledb` (conexão com Oracle via Wallet)
- `python-dotenv` (carregamento do `.env`)
- `pywin32` — usado apenas em `ExcelFileHandler.set_confidentiality` (rotulagem MSIP via Excel COM), método atualmente não utilizado no fluxo principal.

## Observações

- A tabela de destino no Oracle é `MEDIA_DIARIA_EXPORTACAO_MDIC`.
- O pipeline é idempotente em relação à checagem de data: só reprocessa quando há uma atualização nova publicada pelo MDIC.
- `src/__init__.py` está vazio, apenas marca `src` como pacote Python.
