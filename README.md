# Média Diária de Exportação - MDIC

Rotina automatizada que baixa o boletim semanal da **Balança Comercial Brasileira** (MDIC/SECEX), valida a divulgação inteira, filtra os produtos de interesse e grava na tabela do data lake **`BRADODW_IM_PRICING.BZ_MEDIA_SEMANAL_DE_EXPORTACAO`**. A Nota (PDF) da divulgação vai para o bucket da OCI.

## Rotina semanal (`main.py`)

```bat
Carregar_MDIC.bat              :: uso normal (é o que o agendador deve chamar)
Carregar_MDIC.bat --conferir   :: mostra o que seria gravado no banco, sem gravar nem enviar nada
```

### Como funciona

1. **Baixa três arquivos de URL fixa** (`https://balanca.economia.gov.br/balanca/semanal/`), que o MDIC substitui toda semana:

   | Arquivo | Para quê |
   |---|---|
   | `Setores_Produtos.xlsx` | Os produtos, e o título diz de qual semana é a divulgação |
   | `Tabela_Resumo.xlsx` | Data de referência, dias úteis e os totais usados na validação |
   | `Nota.pdf` | O texto da divulgação, enviado ao bucket (ver "Nota no bucket da OCI") |

   As planilhas não ficam guardadas: depois de processadas, o que importa está no banco.

2. **Valida antes de gravar**, sempre sobre a divulgação **inteira** (todos os 575 produtos, exportação e importação):
   - as duas planilhas são da mesma semana, e as abas EXP e IMP também;
   - o layout das planilhas é o esperado (coluna inesperada ou faltando derruba a execução);
   - os produtos somam o setor, os setores somam o total da `Tabela_Resumo` e as semanas somam o mês;
   - a variação publicada bate com média diária ÷ base do ano anterior.

   Se algo falhar, nada é gravado e o erro vai para `saida/extracao.log`.

3. **Só no fim** reduz ao que está configurado para captura (ver abaixo).

4. **Lê o histórico da própria tabela**, junta a divulgação atual e recalcula as colunas que dependem das outras semanas (`VALOR_SEMANA_USD`, `PESO_SEMANA_TON`, `ULTIMA_DO_MES`, `MAIS_RECENTE`).

5. **Grava só o que mudou** (ver "Gravação no data lake"). Cada linha gravada sai no log: se é nova ou quais colunas mudaram.

6. **Envia a Nota ao bucket** (ver "Nota no bucket da OCI").

O código de saída é 1 quando algo falha, e o Agendador de Tarefas mostra isso em "Resultado da última execução". Banco e Nota são independentes: se um falhar, o outro segue, e a execução termina com código 1. **Agende para rodar diariamente, por exemplo às 18h.** O MDIC publica às segundas, entre 15h e 15h30. Rodar a mais não custa nada (sem divulgação nova, nada é gravado) e cobre atrasos da publicação. A execução leva de 40 s a 1 min: quase tudo é o `funcoes_uteis` carregando o Oracle Client da rede e conectando; o resto da rotina leva ~2 s.

> **Não deixe passar uma semana.** O site só tem a divulgação atual. Uma semana que não for capturada não volta mais (nem a Nota), e a semana seguinte do mesmo mês fica com `VALOR_SEMANA_USD` / `PESO_SEMANA_TON` vazios.

### `--conferir`

Faz tudo (baixa, valida, lê o banco, calcula) e para antes de gravar: mostra no log quais linhas seriam gravadas e o que mudaria em cada uma, sem gravar nada no banco e sem baixar nem enviar a Nota. Serve para testar depois de mexer no código ou na configuração de captura.

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

Produto incluído aqui passa a entrar **a partir da próxima divulgação**: as semanas anteriores não são refeitas, porque as planilhas não ficam guardadas.

A comparação do nome ignora acento, maiúscula/minúscula e espaços extras. Se um produto configurado não existir na divulgação (sinal de que o MDIC mudou o nome), a execução **para com erro** em vez de gravar a semana sem ele.

Para capturar importação também, a coluna `FLUXO` precisa entrar no `ESQUEMA` — senão o mesmo produto nos dois fluxos viraria duas linhas indistinguíveis. O código avisa se isso for esquecido.

### A tabela

Uma linha por produto capturado em cada divulgação semanal.

| Coluna | Descrição |
|---|---|
| `DATA_REFERENCIA` | Último dia coberto: os dados vão do dia 1º do mês até esta data. Eixo de tempo. |
| `SEMANA` | Número da divulgação no mês ("até a Nª semana"). |
| `ULTIMA_DO_MES` | 1 na última divulgação coletada de cada mês, 0 nas demais. Para comparar meses. |
| `MAIS_RECENTE` | 1 na divulgação mais recente, 0 nas demais. Para a situação atual. |
| `PRODUTO` | Produto (grupo CUCI), como publicado pelo MDIC. |
| `VALOR_ACUMULADO_MES_USD` | Valor FOB em US$ do dia 1º até `DATA_REFERENCIA`. |
| `VALOR_SEMANA_USD` | Valor FOB só da semana (acumulado menos o da divulgação anterior do mês). |
| `MEDIA_DIARIA_USD` | Valor por dia útil no acumulado do mês. |
| `VARIACAO_VALOR` | Variação da média diária em US$ contra o mesmo mês completo do ano anterior. |
| `PESO_ACUMULADO_MES_TON` | Toneladas do dia 1º até `DATA_REFERENCIA`. |
| `PESO_SEMANA_TON` | Toneladas só da semana. |
| `MEDIA_DIARIA_TON` | Toneladas por dia útil no acumulado do mês. |
| `VARIACAO_PESO` | Variação da média diária em toneladas contra o mesmo mês do ano anterior. |
| `PRECO_MEDIO_USD_TON` | US$ por tonelada (valor ÷ peso). Vazio quando não houve peso. |
| `VARIACAO_PRECO_MEDIO_TON` | Variação do preço médio contra o mesmo mês do ano anterior. Vazio sem um dos preços. |
| `DATA_CARGA_DW` | Quando a linha entrou no DW (preenchida pelo banco). |

Setores e totais não são gravados. **Com o recorte atual, a soma das linhas não é o total do setor nem do Brasil** — é só o subtotal daqueles produtos, e não dá para calcular saldo nem corrente de comércio.

Convenções:
- Valores em **US$** (não em US$ mil). Peso em toneladas.
- Variações em **fração**: 0,285 = 28,5%, sempre da **média diária** contra o mesmo mês **completo** do ano anterior — é assim que o MDIC publica, e é o que torna comparável um mês parcial com um mês fechado.
- `VALOR_SEMANA_USD` / `PESO_SEMANA_TON` são o acumulado da divulgação menos o da anterior do mesmo mês, e ficam **vazios** quando a semana anterior não foi coletada.
- Onde o MDIC publica 0% porque não houve embarque no mês, a coluna traz **−100%**; onde não existe base no ano anterior, fica **vazia**.

> **Limite do recorte atual:** as colunas `VARIACAO_*` valem por produto. Para a variação de um **conjunto** de produtos seria preciso somar as médias diárias do ano anterior, que não estão na tabela. Se isso passar a ser necessário, inclua `MEDIA_DIARIA_ANO_ANTERIOR_USD` / `_TON` no `ESQUEMA` (elas continuam sendo lidas e validadas internamente) e crie as colunas na tabela.

### Gravação no data lake

Feita por [`src/banco.py`](src/banco.py), seguindo o padrão dos outros scripts do time: conexão, leitura e gravação vêm do módulo compartilhado `funcoes_uteis`, em `\\boci01filesrv\Comercial_Dados\utils\Funcoes_Python` (o mesmo `Y:\utils\Funcoes_Python`; o código usa o caminho de rede porque a unidade `Y:` não existe no Agendador de Tarefas nem no Airflow).

- **Só esta tabela**: a rotina lê e grava apenas a `BZ_MEDIA_SEMANAL_DE_EXPORTACAO`; não toca em nenhuma outra tabela do banco.
- **Credencial**: **nenhuma fica neste projeto** — não há `.env` aqui. `conecta_banco()` pega usuário e senha do Windows Credential Manager (`oracle_db_cred`) e o resto da conexão (`LIB_DIR`, `DSN`, `CONFIG_DIR`) do `.env` central do `funcoes_uteis`, em `utils\Credenciais`.
- **Histórico**: a tabela é lida inteira a cada execução e é a base para recalcular as colunas que dependem das outras semanas.
- **Gravação**: `upsert_df_banco()` faz `MERGE` pela chave natural (`DATA_REFERENCIA, PRODUTO`) — rodar de novo a mesma semana atualiza aquelas linhas em vez de duplicar, e nada é apagado. O índice único na chave faz o próprio banco barrar duplicatas.
- **Tudo ou nada**: o `upsert_df_banco` só faz commit no fim. Se a gravação falhar no meio, a conexão é fechada sem commit e o Oracle desfaz o que ficou pendente — a divulgação nunca fica gravada pela metade.
- **Só o que mudou**: o `upsert_df_banco` roda o MERGE linha a linha, uma ida ao servidor por linha. Por isso o `apenas_mudancas()` compara com o que já está gravado e manda só as linhas diferentes — normalmente as 3 da divulgação nova mais as 3 da anterior, que perdem `ULTIMA_DO_MES` / `MAIS_RECENTE`.
- **Mesma semana com outra data**: se a tabela já tiver a mesma semana do mês gravada com outra `DATA_REFERENCIA`, a execução para com erro em vez de gravar — o MERGE casa pela data, então gravaria a semana duas vezes.

A tabela é **bronze**: todas as colunas são `VARCHAR2`, como as outras `BZ_` do schema. O dado entra cru, do jeito que o MDIC publicou, e a conversão para número e data fica para a camada de cima. Datas vão em ISO (`2026-09-20`), que ordena como texto; números com até 15 dígitos significativos.

### Nota no bucket da OCI

Feito por [`src/bucket.py`](src/bucket.py).

- **Destino**: bucket `brado-inteligencia-mercado`, em `mdic/<ano>/Nota-<mes>-<dia>.pdf`. A data é a `DATA_REFERENCIA` (último dia da semana). Ex.: 3ª semana de set/2026 → `mdic/2026/Nota-09-20.pdf`.
- **Reenvio**: mesma semana = mesmo nome = mesmo arquivo. Rodar de novo só sobrescreve com o mesmo PDF.
- **Conferência**: a Nota é baixada à parte das planilhas. Se a capa for de outra semana, ela não é enviada.
- **Falha**: o SDK da OCI tenta de novo sozinho em falhas temporárias. Se ainda falhar, o banco é gravado mesmo assim, o erro vai para o log e a execução termina com código 1. A execução do dia seguinte reenvia.
- **Credencial**: `~/.oci/config` (perfil `DEFAULT`), fora do projeto. Precisa existir no usuário que roda o agendador.
- Não fica cópia local da Nota.

### Histórico migrado da tabela antiga

O histórico de mar/2025 a mar/2026 que estava em `OTAVIO_MIRANDA.MEDIA_DIARIA_EXPORTACAO_MDIC` já foi copiado para a tabela (migração única). Pontos a saber sobre essas linhas:

- **Data de referência reconstruída**: a tabela antiga só tinha o dia 1º do mês. A data foi reconstruída pela convenção do MDIC (semanas de segunda a domingo; sábado/domingo no início do mês entram na semana 1) e conferida contra os dias úteis (`VALOR_ACUMULADO_MES_USD ÷ MEDIA_DIARIA_USD`). **27 das 29 divulgações fecham exato**; `2026-01-31` semana 5 e `2026-03-29` semana 4 não fecham.
- **Variação duvidosa nessas mesmas duas**: a variação publicada não bate com média diária ÷ base do ano anterior (erro de 6 a 13 p.p.). Valor, peso e preço do período atual são coerentes nas 29 — só a comparação com 2025 é duvidosa.
- Os 5 indicadores `A-1` da tabela antiga não têm coluna na nova e ficaram de fora.

### Como usar os dados (Power BI)

1. Converta os tipos (a tabela é toda texto): `DATA_REFERENCIA` para data, as colunas de valor, peso, preço e variação para número decimal.
2. Não precisa de relacionamento nem de tabela de dimensão: é uma tabela só. Use `DATA_REFERENCIA` como eixo de tempo.
3. Formate as colunas `VARIACAO_*` como porcentagem e marque como **Não resumir**: variação não se soma nem se tira média.
4. Cada divulgação é uma foto do mês até aquela semana. Por isso:
   - painel da situação atual: filtre `MAIS_RECENTE = 1`;
   - comparar meses: filtre `ULTIMA_DO_MES = 1`;
   - evolução dentro do mês: use `MEDIA_DIARIA_USD` no eixo `DATA_REFERENCIA`;
   - somar semanas: use `VALOR_SEMANA_USD`, nunca `VALOR_ACUMULADO_MES_USD` de divulgações diferentes do mesmo mês.

### Dependências

```bat
pip install -r requirements.txt
```

As versões testadas estão fixadas em [`requirements.txt`](requirements.txt) (Python 3.14):

| Biblioteca | Para quê |
|---|---|
| `requests` | Baixar as planilhas e a Nota |
| `pandas` | Ler e tratar os dados das planilhas |
| `openpyxl` | Motor do pandas para ler `.xlsx` |
| `urllib3` | Suprimir o aviso de certificado SSL do site do governo (vem com `requests`) |
| `oci` | Enviar a Nota ao bucket |
| `pymupdf` | Ler a capa da Nota para conferir a semana |
| `oracledb`, `keyring`, `python-dotenv`, `numpy` | Usados pelo `funcoes_uteis` |

O `funcoes_uteis` também precisa do Oracle Instant Client e de acesso à rede.

---

## Estrutura de arquivos

```
media_semanal_de_exportacao/
├── Carregar_MDIC.bat                # Execução da rotina semanal (agendador de tarefas)
├── main.py                          # Rotina semanal: baixa, valida, grava no banco e envia a Nota
├── requirements.txt                 # Dependências com as versões testadas
├── src/
│   ├── principais_resultados.py     # Leitura, validação e tratamento de uma divulgação
│   ├── banco.py                     # Histórico e gravação na tabela bronze (usa o funcoes_uteis)
│   └── bucket.py                    # Envio da Nota ao bucket da OCI
└── saida/                           # Não versionado
    └── extracao.log                 # Log de todas as execuções
```

`src/__init__.py` está vazio, apenas marca `src` como pacote Python.
