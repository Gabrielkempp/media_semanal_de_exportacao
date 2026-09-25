# Histórico e gravação na tabela bronze. Acessa só a TABELA_BANCO.

import logging
import sys

import pandas as pd

from src import principais_resultados as pr

# mesmo que Y:\utils\Funcoes_Python (Y: não existe no Agendador/Airflow)
funcoes_path = r"\\boci01filesrv\Comercial_Dados\utils\Funcoes_Python"

log = logging.getLogger('mdic')


# Junta `atual` ao histórico, recalcula e grava só o que mudou. Retorna as linhas gravadas.
def gravar(atual, conferir=False):
    # import aqui: falha de rede cai no log da execução
    if funcoes_path not in sys.path:
        sys.path.append(funcoes_path)
    import funcoes_uteis as banco

    conn = banco.conecta_banco()
    try:
        gravadas = banco.leitura_dados_banco(connection=conn,
                                             query=f"SELECT {', '.join(pr.ESQUEMA)} FROM {pr.TABELA_BANCO}")
        # NULL -> None (NaN != NaN)
        gravadas = gravadas.astype(object).where(pd.notna(gravadas), None)
        indice = _por_chave(gravadas)

        dados = pr.calcular_colunas_historicas(combinar(para_tipos(gravadas), atual))
        mudancas = apenas_mudancas(para_texto(dados), indice)

        log.info(f"Banco: {len(mudancas)} linhas a gravar em {pr.TABELA_BANCO} "
                 f"(tabela tem {len(gravadas)}, ficará com {len(dados)})")
        for linha in resumo_mudancas(mudancas, indice):
            log.info(f"  {linha}")
        if conferir:
            log.info("Conferência: nada foi gravado")
            return 0
        if mudancas.empty:
            return 0
        linhas = banco.upsert_df_banco(connection=conn, table_name=pr.TABELA_BANCO, df=mudancas,
                                       chave_colunas=pr.CHAVE_NATURAL)
        log.info(f"Banco: {linhas} linhas gravadas")
        return linhas
    finally:
        # sem commit = rollback: nunca grava a divulgação pela metade
        conn.close()


# Texto do banco -> tipos de cálculo (inverso do para_texto).
def para_tipos(gravadas):
    df = gravadas.copy()
    for coluna in pr.ESQUEMA:
        if coluna in pr.COLUNAS_DATA:
            df[coluna] = pd.to_datetime(df[coluna])
        elif coluna not in pr.COLUNAS_TEXTO:
            # exato; pd.to_numeric erra o 15º dígito
            df[coluna] = df[coluna].astype(float)
    return pr.ajustar_tipos(df)


# Histórico + divulgação atual (substitui a mesma data). Mesma semana com outra data = erro.
def combinar(historico, atual):
    data = atual['DATA_REFERENCIA'].iloc[0]
    semana = int(atual['SEMANA'].iloc[0])
    mesma_semana = ((historico['DATA_REFERENCIA'].dt.to_period('M') == data.to_period('M'))
                    & (historico['SEMANA'] == semana))
    outras_datas = sorted(historico.loc[mesma_semana & (historico['DATA_REFERENCIA'] != data), 'DATA_REFERENCIA']
                          .dt.strftime('%Y-%m-%d').unique())
    if outras_datas:
        raise ValueError(f"A tabela já tem a {semana}ª semana de {data:%m/%Y} com a data {outras_datas}, "
                         f"mas a divulgação atual é de {data:%Y-%m-%d}. Verifique antes de gravar.")

    partes = [historico[historico['DATA_REFERENCIA'] != data], atual]
    dados = pd.concat([p for p in partes if len(p)], ignore_index=True)
    return dados.sort_values('DATA_REFERENCIA', kind='stable').reset_index(drop=True)


def _formato(coluna):
    if coluna in pr.COLUNAS_DATA:
        # ISO: ordena como texto
        return lambda valor: valor.strftime('%Y-%m-%d')
    if coluna in pr.COLUNAS_INTEIRAS or coluna in pr.COLUNAS_SIM_NAO:
        return lambda valor: str(int(valor))
    if coluna in pr.COLUNAS_TEXTO:
        return str
    # 15 dígitos: mesmo texto calculado ou relido do banco
    return lambda valor: f'{float(valor):.15g}'


# Bronze = tudo VARCHAR2; vazio = NULL.
def para_texto(dados):
    textos = {}
    for coluna in pr.ESQUEMA:
        formato = _formato(coluna)
        # dtype=object: mantém None (NaN chegaria ao banco como "nan")
        textos[coluna] = pd.Series([None if pd.isna(valor) else formato(valor) for valor in dados[coluna]],
                                   index=dados.index, dtype=object)
    return pd.DataFrame(textos, index=dados.index)


# {chave natural: linha} de um df nas colunas do ESQUEMA.
def _por_chave(df):
    posicoes = [pr.ESQUEMA.index(coluna) for coluna in pr.CHAVE_NATURAL]
    return {tuple(linha[p] for p in posicoes): linha for linha in df[pr.ESQUEMA].itertuples(index=False, name=None)}


# Só as linhas diferentes do gravado (o MERGE vai ao servidor 1 vez por linha).
def apenas_mudancas(df, indice):
    return df[[indice.get(chave) != linha for chave, linha in _por_chave(df).items()]]


# Por linha a gravar: 'nova' ou as colunas que mudam (antes -> depois).
def resumo_mudancas(mudancas, indice):
    for chave, linha in _por_chave(mudancas).items():
        descricao = ' | '.join(chave)
        anterior = indice.get(chave)
        if anterior is None:
            yield f"{descricao}: nova"
        else:
            diferencas = [f"{coluna} {antes} -> {depois}"
                          for coluna, antes, depois in zip(pr.ESQUEMA, anterior, linha) if antes != depois]
            yield f"{descricao}: {', '.join(diferencas)}"
