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
3. Valida antes de gravar, sempre sobre a divulgação **inteira** (todos os produtos, exportação e importação):
   - a página e as planilhas são da mesma semana;
   - o layout das planilhas é o esperado;
   - os produtos somam o setor, os setores somam o total da `Tabela_Resumo` e as semanas somam o mês;
   - a variação publicada bate com média diária ÷ base do ano anterior.

   Se algo falhar, aquela divulgação não entra e o erro vai para `saida/extracao.log`.
4. **Só no fim** reduz ao que está configurado para captura (ver abaixo) e grava.
5. Grava o Excel num arquivo temporário e troca no fim. Reexecutar a mesma semana substitui as linhas dela, sem duplicar. Se a estrutura das tabelas mudar no código, o Excel é reconstruído automaticamente a partir dos brutos.

### O que é capturado (configurável)

No topo de [`src/principais_resultados.py`](src/principais_resultados.py):

```python
PRODUTOS_CAPTURADOS = [   # None = todos os produtos da divulgação
    'Algodão em bruto',
    'Madeira em bruto',
    'Milho não moído, exceto milho doce',
]
FLUXOS_CAPTURADOS = [EXPORTACAO]   # [EXPORTACAO, IMPORTACAO] para capturar os dois
```

O recorte vale só para a tabela final. **Os arquivos brutos continuam guardando a divulgação completa**, então para ampliar a captura basta editar as listas e rodar `extrair-para-excel.bat --reprocessar`: todo o histórico já arquivado é refeito com os produtos novos, sem perder nada.

A comparação do nome ignora acento, maiúscula/minúscula e espaços extras. Se um produto configurado não existir na divulgação (sinal de que o MDIC mudou o nome), a execução **para com erro** em vez de gravar a semana sem ele.

O código de saída é 1 quando a divulgação atual falha, e o Agendador de Tarefas mostra isso em "Resultado da última execução". **Agende para rodar diariamente, por exemplo às 18h.** O MDIC publica às segundas, entre 15h e 15h30. Rodar a mais não custa nada e cobre atrasos da publicação.

### Abas do Excel

| Aba | Conteúdo |
|---|---|
| `DADOS` | **A tabela de análise.** Uma linha por produto capturado em cada divulgação semanal (hoje: 3 produtos de exportação). |
| `TEXTOS` | Um parágrafo por linha das seções Destaques, Totais e Setores e Produtos, como publicados — inclui exportação e importação. |
| `LEIA-ME` | Regras de uso e o dicionário de todas as colunas, montado a partir do que está configurado para captura. |

`DADOS` e `TEXTOS` são Tabelas do Excel com esses mesmos nomes, que é o que o Power BI lista ao importar.

Setores e totais não são gravados. Com a captura em todos os produtos (`PRODUTOS_CAPTURADOS = None`), eles são exatamente a soma das linhas. **Com o recorte atual, a soma das linhas não é o total do setor nem do Brasil** — é só o subtotal daqueles produtos.

Colunas de `DADOS`, em grupos:

- **Quando**: `DATA_REFERENCIA` (último dia coberto), `DATA_PUBLICACAO`, `ANO`, `MES`, `SEMANA`, `DIAS_UTEIS_ACUMULADO_MES`, `DIAS_UTEIS_SEMANA`, `ULTIMA_DO_MES` e `MAIS_RECENTE` (1 = sim, 0 = não).
- **O quê**: `PRODUTO`.
- **Valor**: `VALOR_ACUMULADO_MES_USD`, `VALOR_SEMANA_USD`, `MEDIA_DIARIA_USD`, `MEDIA_DIARIA_ANO_ANTERIOR_USD`, `VARIACAO_VALOR`.
- **Peso**: as mesmas cinco, em toneladas.
- **Preço**: `PRECO_MEDIO_USD_TON`, `VARIACAO_PRECO`.

Convenções:
- Valores em **US$** (não em US$ mil), para o Power BI exibir "14,25 bi" corretamente. Peso em toneladas.
- Variações em **fração**: 0,285 = 28,5%, sempre contra a média diária do mesmo mês do ano anterior.
- `MEDIA_DIARIA_ANO_ANTERIOR_*` é a base publicada pelo MDIC na mesma planilha, e se refere ao **mês completo** do ano anterior. Por isso a comparação é sempre por média diária.
- `VALOR_SEMANA_USD` é o acumulado da divulgação menos o da anterior do mesmo mês, e fica **vazio** quando a semana anterior não foi coletada.
- Onde o MDIC publica 0% porque não houve embarque no mês, a coluna traz **−100%**; onde não existe base no ano anterior, fica **vazia**.

### Levando para o banco

A tabela nova tem o mesmo formato da aba `DADOS` — uma linha por produto, uma coluna por métrica. Não é o formato da antiga `MEDIA_DIARIA_EXPORTACAO_MDIC` (longo, com os indicadores `A-1`), e nada aqui grava naquela tabela.

O DDL sai do próprio código, então nunca fica fora de sincronia com o que a rotina gera:

```bat
extrair-para-excel.bat --ddl                          :: mostra o CREATE TABLE na tela
extrair-para-excel.bat --ddl > sql\criar_tabelas.sql  :: regera o arquivo depois de mudar colunas
```

O resultado está em [`sql/criar_tabelas.sql`](sql/criar_tabelas.sql) e traz, para `DADOS` e `TEXTOS`:

- o `CREATE TABLE` com os tipos já definidos;
- um índice único na chave natural (`ANO, MES, SEMANA, FLUXO, PRODUTO`), que faz o **banco** barrar divulgação duplicada, não só a rotina;
- um `COMMENT ON COLUMN` por coluna, com a mesma descrição da aba LEIA-ME — assim a documentação vai junto para o banco.

Ajuste o nome e o esquema das tabelas ao padrão do banco de destino antes de rodar. O DDL não foi executado num banco real.

Como a tabela já vem no formato largo, o Power BI lê dela direto, sem pivot e sem Power Query.

### Montando no Power BI (a partir do Excel)

1. Obter dados → Excel → marque as Tabelas `DADOS` e `TEXTOS`.
2. Não precisa de relacionamento nem de tabela de dimensão: é uma tabela só. Use `DATA_REFERENCIA` como eixo de tempo.
3. Formate as colunas `VARIACAO_*` como porcentagem e marque como **Não resumir**: variação não se soma nem se tira média.
4. Para a variação de um conjunto de produtos (não some as colunas `VARIACAO_*`), crie uma medida:

   ```dax
   Variação = DIVIDE( SUM(DADOS[MEDIA_DIARIA_USD]), SUM(DADOS[MEDIA_DIARIA_ANO_ANTERIOR_USD]) ) - 1
   ```

   Isso dá o número certo em qualquer agrupamento dos produtos capturados. Saldo e corrente de comércio só seriam possíveis capturando também a importação.
5. Cada divulgação é uma foto do mês até aquela semana. Por isso:
   - painel da situação atual: filtre `MAIS_RECENTE = 1`;
   - comparar meses: filtre `ULTIMA_DO_MES = 1`;
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

### Rotina semanal (`extrair_para_excel.py` / `src/principais_resultados.py`)

Só bibliotecas externas comuns — nada de Selenium/Chrome, Oracle ou `.env`:

| Biblioteca | Versão usada | Para quê |
|---|---|---|
| `requests` | 2.34.2 | Baixar a página e as planilhas do MDIC |
| `pandas` | 3.0.5 | Ler e tratar os dados das planilhas |
| `openpyxl` | 3.1.5 | Ler `.xlsx` e gravar o Excel de saída (formatado, com Tabelas) |
| `urllib3` | 2.7.0 | Suprimir o aviso de certificado SSL do site do governo (usa `requests`, não é instalada à parte) |

O resto (`re`, `dataclasses`, `pathlib`, `html.parser`, `argparse`, `logging`, `datetime` etc.) já vem pronto no Python, sem instalar nada.

```bat
pip install requests pandas openpyxl
```

### Fluxo antigo (`main.py` / Oracle)

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
