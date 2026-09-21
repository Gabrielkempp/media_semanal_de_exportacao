-- ---------------------------------------------------------------------------------------------------
-- Tabelas da rotina semanal do MDIC, para o banco novo.
--
-- GERADO A PARTIR DO CODIGO: nao edite na mao. Para regerar depois de mudar colunas:
--     extrair-para-excel.bat --ddl > sql\criar_tabelas.sql
--
-- Ajuste o nome/esquema das tabelas conforme o padrao do banco de destino antes de rodar.
-- ---------------------------------------------------------------------------------------------------

CREATE TABLE BALANCA_SEMANAL_MDIC_DADOS (
    DATA_REFERENCIA                DATE,
    DATA_PUBLICACAO                DATE,
    ANO                            NUMBER(4),
    MES                            NUMBER(2),
    SEMANA                         NUMBER(2),
    DIAS_UTEIS_ACUMULADO_MES       NUMBER(2),
    DIAS_UTEIS_SEMANA              NUMBER(2),
    ULTIMA_DO_MES                  NUMBER(1),
    MAIS_RECENTE                   NUMBER(1),
    PRODUTO                        VARCHAR2(400),
    VALOR_ACUMULADO_MES_USD        NUMBER(18,2),
    VALOR_SEMANA_USD               NUMBER(18,2),
    MEDIA_DIARIA_USD               NUMBER(18,6),
    MEDIA_DIARIA_ANO_ANTERIOR_USD  NUMBER(18,6),
    VARIACAO_VALOR                 NUMBER(18,10),
    PESO_ACUMULADO_MES_TON         NUMBER(18,3),
    PESO_SEMANA_TON                NUMBER(18,3),
    MEDIA_DIARIA_TON               NUMBER(18,6),
    MEDIA_DIARIA_ANO_ANTERIOR_TON  NUMBER(18,6),
    VARIACAO_PESO                  NUMBER(18,10),
    PRECO_MEDIO_USD_TON            NUMBER(18,3),
    VARIACAO_PRECO                 NUMBER(18,10)
);

-- Chave natural: o banco passa a impedir a mesma divulgação gravada duas vezes
CREATE UNIQUE INDEX UX_DADOS ON BALANCA_SEMANAL_MDIC_DADOS (ANO, MES, SEMANA, PRODUTO);

COMMENT ON COLUMN BALANCA_SEMANAL_MDIC_DADOS.DATA_REFERENCIA IS 'Último dia coberto pela divulgação: os dados vão do dia 1º do mês até esta data. Use como eixo de tempo.';
COMMENT ON COLUMN BALANCA_SEMANAL_MDIC_DADOS.DATA_PUBLICACAO IS 'Data em que o MDIC publicou a divulgação.';
COMMENT ON COLUMN BALANCA_SEMANAL_MDIC_DADOS.ANO IS 'Ano de referência dos dados.';
COMMENT ON COLUMN BALANCA_SEMANAL_MDIC_DADOS.MES IS 'Mês de referência dos dados.';
COMMENT ON COLUMN BALANCA_SEMANAL_MDIC_DADOS.SEMANA IS 'Número da divulgação no mês ("até a Nª semana").';
COMMENT ON COLUMN BALANCA_SEMANAL_MDIC_DADOS.DIAS_UTEIS_ACUMULADO_MES IS 'Dias úteis do dia 1º do mês até DATA_REFERENCIA.';
COMMENT ON COLUMN BALANCA_SEMANAL_MDIC_DADOS.DIAS_UTEIS_SEMANA IS 'Dias úteis só da semana desta divulgação.';
COMMENT ON COLUMN BALANCA_SEMANAL_MDIC_DADOS.ULTIMA_DO_MES IS '1 na última divulgação coletada de cada mês, 0 nas demais. Use para comparar meses.';
COMMENT ON COLUMN BALANCA_SEMANAL_MDIC_DADOS.MAIS_RECENTE IS '1 na divulgação mais recente, 0 nas demais. Use para ver a situação atual.';
COMMENT ON COLUMN BALANCA_SEMANAL_MDIC_DADOS.PRODUTO IS 'Produto (grupo CUCI), como publicado pelo MDIC.';
COMMENT ON COLUMN BALANCA_SEMANAL_MDIC_DADOS.VALOR_ACUMULADO_MES_USD IS 'Valor FOB em US$ do dia 1º do mês até DATA_REFERENCIA.';
COMMENT ON COLUMN BALANCA_SEMANAL_MDIC_DADOS.VALOR_SEMANA_USD IS 'Valor FOB em US$ só da semana (acumulado desta divulgação menos o da anterior do mesmo mês). Vazio quando a divulgação anterior do mês não foi coletada.';
COMMENT ON COLUMN BALANCA_SEMANAL_MDIC_DADOS.MEDIA_DIARIA_USD IS 'Valor por dia útil (VALOR_ACUMULADO_MES_USD ÷ DIAS_UTEIS_ACUMULADO_MES). Use para comparar períodos.';
COMMENT ON COLUMN BALANCA_SEMANAL_MDIC_DADOS.MEDIA_DIARIA_ANO_ANTERIOR_USD IS 'Média diária do mesmo mês do ano anterior (mês completo), como publicada pelo MDIC na mesma planilha. É a base da variação.';
COMMENT ON COLUMN BALANCA_SEMANAL_MDIC_DADOS.VARIACAO_VALOR IS 'MEDIA_DIARIA_USD ÷ MEDIA_DIARIA_ANO_ANTERIOR_USD − 1, em fração (0,285 = 28,5%). −100% = sem embarque no mês; vazio = sem base no ano anterior.';
COMMENT ON COLUMN BALANCA_SEMANAL_MDIC_DADOS.PESO_ACUMULADO_MES_TON IS 'Peso em toneladas do dia 1º do mês até DATA_REFERENCIA.';
COMMENT ON COLUMN BALANCA_SEMANAL_MDIC_DADOS.PESO_SEMANA_TON IS 'Peso em toneladas só da semana (mesma lógica de VALOR_SEMANA_USD).';
COMMENT ON COLUMN BALANCA_SEMANAL_MDIC_DADOS.MEDIA_DIARIA_TON IS 'Toneladas por dia útil.';
COMMENT ON COLUMN BALANCA_SEMANAL_MDIC_DADOS.MEDIA_DIARIA_ANO_ANTERIOR_TON IS 'Toneladas por dia útil no mesmo mês do ano anterior (mês completo), como publicado pelo MDIC.';
COMMENT ON COLUMN BALANCA_SEMANAL_MDIC_DADOS.VARIACAO_PESO IS 'MEDIA_DIARIA_TON ÷ MEDIA_DIARIA_ANO_ANTERIOR_TON − 1, em fração (mesmas regras de VARIACAO_VALOR).';
COMMENT ON COLUMN BALANCA_SEMANAL_MDIC_DADOS.PRECO_MEDIO_USD_TON IS 'Preço médio em US$ por tonelada (valor ÷ peso). Vazio quando não houve peso.';
COMMENT ON COLUMN BALANCA_SEMANAL_MDIC_DADOS.VARIACAO_PRECO IS 'Variação do preço médio contra o mesmo mês do ano anterior, em fração. Vazio quando um dos preços não existe.';

CREATE TABLE BALANCA_SEMANAL_MDIC_TEXTOS (
    DATA_REFERENCIA  DATE,
    ANO              NUMBER(4),
    MES              NUMBER(2),
    SEMANA           NUMBER(2),
    MAIS_RECENTE     NUMBER(1),
    ORDEM            NUMBER(4),
    SECAO            VARCHAR2(40),
    FLUXO            VARCHAR2(20),
    SUBSECAO         VARCHAR2(60),
    TEXTO            CLOB
);

-- Chave natural: o banco passa a impedir a mesma divulgação gravada duas vezes
CREATE UNIQUE INDEX UX_TEXTOS ON BALANCA_SEMANAL_MDIC_TEXTOS (ANO, MES, SEMANA, ORDEM);

COMMENT ON COLUMN BALANCA_SEMANAL_MDIC_TEXTOS.DATA_REFERENCIA IS 'Como em DADOS: liga o texto à divulgação.';
COMMENT ON COLUMN BALANCA_SEMANAL_MDIC_TEXTOS.ANO IS 'Ano de referência.';
COMMENT ON COLUMN BALANCA_SEMANAL_MDIC_TEXTOS.MES IS 'Mês de referência.';
COMMENT ON COLUMN BALANCA_SEMANAL_MDIC_TEXTOS.SEMANA IS 'Número da divulgação no mês.';
COMMENT ON COLUMN BALANCA_SEMANAL_MDIC_TEXTOS.MAIS_RECENTE IS '1 nos textos da divulgação mais recente, 0 nos demais.';
COMMENT ON COLUMN BALANCA_SEMANAL_MDIC_TEXTOS.ORDEM IS 'Posição do texto na página.';
COMMENT ON COLUMN BALANCA_SEMANAL_MDIC_TEXTOS.SECAO IS 'Seção da página: Destaques, Totais ou Setores e Produtos.';
COMMENT ON COLUMN BALANCA_SEMANAL_MDIC_TEXTOS.FLUXO IS 'Exportação, Importação, Saldo ou Corrente (vazio = texto geral).';
COMMENT ON COLUMN BALANCA_SEMANAL_MDIC_TEXTOS.SUBSECAO IS 'Subtítulo do texto na página.';
COMMENT ON COLUMN BALANCA_SEMANAL_MDIC_TEXTOS.TEXTO IS 'Texto como publicado pelo MDIC.';

