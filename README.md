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

## Estrutura de arquivos

```
media-diaria-exportacao-mdic/
├── main.py                          # Orquestra o pipeline (entry point)
├── media-diaria-exportacao-mdic.bat # Script de execução (usado em agendador de tarefas)
├── .env                             # Credenciais/config do Oracle (não versionado)
└── src/
    ├── web_scraper.py               # Selenium: extrai data da última atualização do MDIC
    ├── balanca_semanal.py           # Download, parsing, filtro e transformação da planilha
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
