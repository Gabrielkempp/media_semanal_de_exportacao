# Leitura, validação e tratamento da divulgação semanal da Balança Comercial (MDIC/SECEX).
import io
import logging
import re
import unicodedata
from dataclasses import dataclass

import pandas as pd
import pymupdf
import requests
import urllib3

BASE = "https://balanca.economia.gov.br/balanca/semanal/"
ARQ_SETORES_PRODUTOS = 'Setores_Produtos.xlsx'
ARQ_TABELA_RESUMO = 'Tabela_Resumo.xlsx'
ARQ_NOTA = 'Nota.pdf'
URLS = {nome: BASE + nome for nome in (ARQ_SETORES_PRODUTOS, ARQ_TABELA_RESUMO, ARQ_NOTA)}
PLANILHAS = (ARQ_SETORES_PRODUTOS, ARQ_TABELA_RESUMO)

MESES = ['janeiro', 'fevereiro', 'março', 'abril', 'maio', 'junho',
         'julho', 'agosto', 'setembro', 'outubro', 'novembro', 'dezembro']
MESES_ABREV = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez']

# título da planilha: 'Até 2ª Semana de Setembro/2026'
RE_SEMANA = re.compile(r'(\d+)\s*[ªº°]?\s*semana\s+de\s+([a-zç]+)\s*/\s*(\d{4})', re.IGNORECASE)
# capa da Nota: '3ª Semana SETEMBRO de 2026'
RE_SEMANA_NOTA = re.compile(r'(\d+)\s*[ªº°]?\s*semana\s+(?:de\s+)?([a-zç]+)\s+de\s+(\d{4})', re.IGNORECASE)
# Tabela_Resumo: '3ª semana (14 a 20)'
RE_SEMANA_RESUMO = re.compile(r'^(\d+)\s*[ªº°] semana\s*\((\d{1,2})\s*a\s*(\d{1,2})\)')
RE_SETOR = re.compile(r'^[A-Z] - (.+)$')

EXPORTACAO, IMPORTACAO = 'Exportação', 'Importação'

# Captura: a divulgação inteira é validada; só isto é gravado (None = todos os produtos)
PRODUTOS_CAPTURADOS = [
    'Algodão em bruto',
    'Madeira em bruto',
    'Milho não moído, exceto milho doce',
]
# os dois fluxos exigem FLUXO no ESQUEMA
FLUXOS_CAPTURADOS = [EXPORTACAO]

ESQUEMA = [
    'DATA_REFERENCIA', 'SEMANA', 'ULTIMA_DO_MES', 'MAIS_RECENTE', 'PRODUTO',
    'VALOR_ACUMULADO_MES_USD', 'VALOR_SEMANA_USD', 'MEDIA_DIARIA_USD', 'VARIACAO_VALOR',
    'PESO_ACUMULADO_MES_TON', 'PESO_SEMANA_TON', 'MEDIA_DIARIA_TON', 'VARIACAO_PESO',
    'PRECO_MEDIO_USD_TON', 'VARIACAO_PRECO_MEDIO_TON',
]
CHAVE_NATURAL = ['DATA_REFERENCIA', 'PRODUTO']
# a tabela tem também DATA_CARGA_DW, preenchida pelo banco
TABELA_BANCO = 'BRADODW_IM_PRICING.BZ_MEDIA_SEMANAL_DE_EXPORTACAO'
# tipos por coluna (as que não aparecem aqui são número)
COLUNAS_INTEIRAS = {'SEMANA'}
COLUNAS_DATA = {'DATA_REFERENCIA'}
# gravadas como 1 / 0
COLUNAS_SIM_NAO = {'ULTIMA_DO_MES', 'MAIS_RECENTE'}
COLUNAS_TEXTO = {'PRODUTO'}

# (grupo, período) -> (coluna, fator): US$ mil -> US$, % -> fração
# "_" = uso interno; None = descartada
COLUNAS_SETORES_PRODUTOS = {
    ('US$ Mil', 'ATUAL'): ('VALOR_ACUMULADO_MES_USD', 1000),
    ('US$ Mil', 'ANTERIOR'): None,
    ('US$ Mil Por Média Diária', 'ATUAL'): ('MEDIA_DIARIA_USD', 1000),
    ('US$ Mil Por Média Diária', 'ANTERIOR'): ('_MEDIA_DIARIA_ANO_ANTERIOR_USD', 1000),
    ('Toneladas', 'ATUAL'): ('PESO_ACUMULADO_MES_TON', 1),
    ('Toneladas', 'ANTERIOR'): None,
    ('Toneladas por Média Diária', 'ATUAL'): ('MEDIA_DIARIA_TON', 1),
    ('Toneladas por Média Diária', 'ANTERIOR'): ('_MEDIA_DIARIA_ANO_ANTERIOR_TON', 1),
    ('Preço (US$/Tonelada)', 'ATUAL'): ('PRECO_MEDIO_USD_TON', 1),
    ('Preço (US$/Tonelada)', 'ANTERIOR'): ('_PRECO_ANO_ANTERIOR', 1),
    ('Variação (%) Por Média Diária', 'Valor US$'): ('VARIACAO_VALOR', 0.01),
    ('Variação (%) Por Média Diária', 'Toneladas'): ('VARIACAO_PESO', 0.01),
    ('Variação (%) Por Média Diária', 'Preço'): ('VARIACAO_PRECO_MEDIO_TON', 0.01),
}

# site sem cadeia completa de certificados
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

log = logging.getLogger('mdic')


@dataclass
class Publicacao:
    ano: int
    mes: int
    semana: int

    # Ex.: 'Set/2026'.
    def rotulo_mes(self, ano=None):
        return f"{MESES_ABREV[self.mes - 1]}/{ano or self.ano}"

    @property
    def descricao(self):
        return f"{self.rotulo_mes()} - {self.semana}ª semana"


def baixar(url):
    response = requests.get(url, verify=False, timeout=120)
    response.raise_for_status()
    return response.content


# 'Até 2ª Semana de Setembro/2026' -> Publicacao(2026, 9, 2).
def ler_periodo(texto, padrao=RE_SEMANA):
    m = padrao.search(texto)
    if not m:
        raise ValueError(f"Período de referência não encontrado em: {texto!r}")
    nome_mes = m.group(2).lower()
    if nome_mes not in MESES:
        raise ValueError(f"Mês desconhecido {nome_mes!r} em: {texto!r}")
    return Publicacao(ano=int(m.group(3)), mes=MESES.index(nome_mes) + 1, semana=int(m.group(1)))


# A Nota é baixada à parte das planilhas: erro se a capa for de outra semana.
def conferir_nota(conteudo, pub):
    with pymupdf.open(stream=conteudo, filetype='pdf') as pdf:
        capa = ' '.join(pdf[0].get_text().split())
    periodo = ler_periodo(capa, padrao=RE_SEMANA_NOTA)
    if periodo != pub:
        raise ValueError(f"{ARQ_NOTA} é de {periodo.descricao}, mas as planilhas são de {pub.descricao}")


# Colunas do ESQUEMA, na ordem, com os tipos padronizados.
def ajustar_tipos(df):
    df = df[ESQUEMA].copy()
    for coluna in df.columns:
        if coluna in COLUNAS_INTEIRAS:
            df[coluna] = df[coluna].astype('Int64')
        elif coluna in COLUNAS_DATA:
            df[coluna] = pd.to_datetime(df[coluna])
        elif coluna in COLUNAS_SIM_NAO:
            df[coluna] = df[coluna].astype(bool).astype('Int64')
    return df


# Retorna (publicação, produtos, setores, dias úteis do mês). Setores: só para validação.
def extrair_setores_produtos(conteudo):
    produtos, setores, pub, dias_uteis = [], [], None, None
    planilha = pd.ExcelFile(io.BytesIO(conteudo))
    for aba, fluxo in (('EXP', EXPORTACAO), ('IMP', IMPORTACAO)):
        df = planilha.parse(aba, header=None)

        # "... Set/2026: 8 dias úteis; ... Até 2ª Semana de Setembro/2026"
        titulo = str(df.iat[3, 0])
        periodo = ler_periodo(titulo)
        if pub is None:
            pub = periodo
        elif periodo != pub:
            raise ValueError(f"{ARQ_SETORES_PRODUTOS}: aba {aba} é de {periodo.descricao}, mas EXP é de {pub.descricao}")
        m = re.search(re.escape(pub.rotulo_mes()) + r':\s*(\d+)\s*dias', titulo)
        dias_uteis = int(m.group(1)) if m else None

        # cabeçalho com células mescladas
        grupos = df.iloc[4].ffill()
        subtitulos = df.iloc[6]
        if grupos[0] != 'Descrição':
            raise ValueError(f"Layout inesperado em {ARQ_SETORES_PRODUTOS} [{aba}]: {grupos.tolist()}")

        periodos = {pub.rotulo_mes(): 'ATUAL', pub.rotulo_mes(pub.ano - 1): 'ANTERIOR'}
        colunas = {}
        for c in range(1, df.shape[1]):
            grupo, sub = grupos[c], subtitulos[c]
            if pd.isna(grupo) or pd.isna(sub):
                continue
            chave = (grupo, periodos.get(sub, sub))
            if chave not in COLUNAS_SETORES_PRODUTOS:
                raise ValueError(f"Coluna inesperada em {ARQ_SETORES_PRODUTOS} [{aba}]: {grupo!r} / {sub!r}")
            if COLUNAS_SETORES_PRODUTOS[chave]:
                colunas[c] = COLUNAS_SETORES_PRODUTOS[chave]
        faltando = {d[0] for d in COLUNAS_SETORES_PRODUTOS.values() if d} - {d for d, _ in colunas.values()}
        if faltando:
            raise ValueError(f"Colunas não encontradas em {ARQ_SETORES_PRODUTOS} [{aba}]: {faltando}")

        corpo = df.iloc[8:, [0, *colunas]]
        corpo = corpo[corpo[0].notna()]
        descricao = corpo[0].astype(str).str.split().str.join(' ')
        manter = ~descricao.str.startswith('Fonte:')
        corpo, descricao = corpo[manter], descricao[manter]

        linhas = pd.DataFrame({destino: corpo[c].astype(float) * fator for c, (destino, fator) in colunas.items()})
        # acumulado: 2 casas tiram o ruído do x1000; médias mantêm a precisão da fonte
        linhas['VALOR_ACUMULADO_MES_USD'] = linhas['VALOR_ACUMULADO_MES_USD'].map(lambda v: round(v, 2))
        setor = descricao.str.extract(RE_SETOR, expand=False)
        linhas = linhas.assign(FLUXO=fluxo, SETOR=setor.ffill())
        eh_setor = setor.notna()
        setores.append(linhas[eh_setor])
        produtos.append(linhas[~eh_setor].assign(PRODUTO=descricao[~eh_setor]))

    return pub, pd.concat(produtos, ignore_index=True), pd.concat(setores, ignore_index=True), dias_uteis


# Corrige os 0% do MDIC (-100% ou sem base) e o preço 0 sem peso; confere o resto.
def corrigir_variacoes(produtos):
    df = produtos
    for variacao, atual, base in (('VARIACAO_VALOR', 'MEDIA_DIARIA_USD', '_MEDIA_DIARIA_ANO_ANTERIOR_USD'),
                                  ('VARIACAO_PESO', 'MEDIA_DIARIA_TON', '_MEDIA_DIARIA_ANO_ANTERIOR_TON')):
        normais = (df[atual] > 0) & (df[base] > 0)
        calculada = df.loc[normais, atual] / df.loc[normais, base] - 1
        divergentes = (calculada - df.loc[normais, variacao]).abs() > 1e-6 + 1e-6 * calculada.abs()
        if divergentes.any():
            exemplos = df.loc[divergentes[divergentes].index, ['FLUXO', 'PRODUTO', variacao]].head(3).to_dict('records')
            raise ValueError(f"{variacao} publicada não bate com a média diária ÷ base: {exemplos}")
        df.loc[(df[atual] == 0) & (df[base] > 0), variacao] = -1.0
        df.loc[df[base].fillna(0) == 0, variacao] = None

    sem_preco = df['PRECO_MEDIO_USD_TON'].fillna(0) == 0
    df.loc[sem_preco, 'PRECO_MEDIO_USD_TON'] = None
    df.loc[sem_preco | (df['_PRECO_ANO_ANTERIOR'].fillna(0) == 0), 'VARIACAO_PRECO_MEDIO_TON'] = None
    return df


# Data de referência, dias úteis e totais do mês (US$) para validação.
def extrair_tabela_resumo(conteudo, pub):
    df = pd.read_excel(io.BytesIO(conteudo), header=None)
    rotulos = df[0].fillna('').astype(str).str.strip()
    nome_mes = MESES[pub.mes - 1].capitalize()

    if not rotulos.str.endswith(f'{nome_mes} de {pub.ano}').any():
        raise ValueError(f"{ARQ_TABELA_RESUMO} não é de {nome_mes}/{pub.ano}")
    cab = rotulos.index[rotulos == 'Período']
    esperado = ['Período', 'Dias Úteis', 'EXPORTAÇÃO', 'IMPORTAÇÃO', 'CORRENTE', 'SALDO']
    if (len(cab) != 1
            or [v for v in df.iloc[cab[0], :10].fillna('') if v] != esperado
            or df.iloc[cab[0] + 1, 2:10].tolist() != ['Valor', 'Média p/ dia útil'] * 4):
        raise ValueError(f"Layout inesperado no cabeçalho de {ARQ_TABELA_RESUMO}")

    # US$ milhões -> US$
    def valores(i):
        return {EXPORTACAO: float(df.iat[i, 2]) * 1e6, IMPORTACAO: float(df.iat[i, 4]) * 1e6}

    re_mes = re.compile(rf'^{nome_mes} \(até a (\d+)\s*[ªº°] semana\)$')
    mes, semanas = None, {}
    for i, rotulo in rotulos.items():
        if m := re_mes.match(rotulo):
            if int(m.group(1)) != pub.semana:
                raise ValueError(f"{ARQ_TABELA_RESUMO} não é da mesma divulgação de {ARQ_SETORES_PRODUTOS}: {rotulo!r}")
            mes = {'dias_uteis': int(df.iat[i, 1]), 'valores': valores(i)}
        elif m := RE_SEMANA_RESUMO.match(rotulo):
            semanas[int(m.group(1))] = {'dias_uteis': int(df.iat[i, 1]), 'dia_fim': int(m.group(3)), 'valores': valores(i)}

    if mes is None or set(semanas) != set(range(1, pub.semana + 1)):
        raise ValueError(f"Mês ou semanas não encontrados em {ARQ_TABELA_RESUMO}: semanas={sorted(semanas)}")
    for fluxo in (EXPORTACAO, IMPORTACAO):
        soma = sum(s['valores'][fluxo] for s in semanas.values())
        if abs(soma - mes['valores'][fluxo]) > 10:
            raise ValueError(f"{ARQ_TABELA_RESUMO}: soma das semanas de {fluxo} ({soma:,.2f}) != mês ({mes['valores'][fluxo]:,.2f})")

    return {
        'data_referencia': pd.Timestamp(pub.ano, pub.mes, semanas[pub.semana]['dia_fim']),
        'dias_uteis_mes': mes['dias_uteis'],
        'total_mes_usd': mes['valores'],
    }


# Confere somas (produtos = setor, setores = total), repetidos e dias úteis.
def validar_somas(produtos, setores, resumo, dias_uteis_produtos, tolerancia=10.0):
    erros = []
    repetidos = produtos.duplicated(['FLUXO', 'PRODUTO'])
    if repetidos.any():
        erros.append(f"Produtos repetidos: {produtos.loc[repetidos, 'PRODUTO'].tolist()}")
    for fluxo in (EXPORTACAO, IMPORTACAO):
        s = setores[setores['FLUXO'] == fluxo].set_index('SETOR')
        p = produtos[produtos['FLUXO'] == fluxo].groupby('SETOR')
        for coluna in ('VALOR_ACUMULADO_MES_USD', '_MEDIA_DIARIA_ANO_ANTERIOR_USD'):
            for setor, soma in p[coluna].sum().items():
                if abs(soma - s.at[setor, coluna]) > tolerancia:
                    erros.append(f"{fluxo}/{setor}: soma dos produtos em {coluna} ({soma:,.2f}) != setor ({s.at[setor, coluna]:,.2f})")
        total = resumo['total_mes_usd'][fluxo]
        if abs(s['VALOR_ACUMULADO_MES_USD'].sum() - total) > tolerancia:
            erros.append(f"{fluxo}: soma dos setores ({s['VALOR_ACUMULADO_MES_USD'].sum():,.2f}) != total ({total:,.2f})")
    if dias_uteis_produtos is not None and dias_uteis_produtos != resumo['dias_uteis_mes']:
        erros.append(f"Dias úteis divergentes: {ARQ_SETORES_PRODUTOS}={dias_uteis_produtos}, "
                     f"{ARQ_TABELA_RESUMO}={resumo['dias_uteis_mes']}")
    if erros:
        raise ValueError("Inconsistências encontradas:\n  " + "\n  ".join(erros))


# Ignora acento, caixa e espaços extras.
def _chave_produto(nome):
    sem_acento = ''.join(c for c in unicodedata.normalize('NFKD', str(nome).lower()) if not unicodedata.combining(c))
    return ' '.join(sem_acento.split())


# Reduz aos fluxos e produtos configurados. Produto não encontrado = erro.
def filtrar_captura(produtos):
    if len(FLUXOS_CAPTURADOS) > 1 and 'FLUXO' not in ESQUEMA:
        raise ValueError("Para capturar mais de um fluxo, a coluna FLUXO precisa entrar no ESQUEMA "
                         f"(hoje FLUXOS_CAPTURADOS = {FLUXOS_CAPTURADOS})")
    filtrado = produtos[produtos['FLUXO'].isin(FLUXOS_CAPTURADOS)]
    if PRODUTOS_CAPTURADOS is not None:
        desejados = {_chave_produto(p): p for p in PRODUTOS_CAPTURADOS}
        chaves = filtrado['PRODUTO'].map(_chave_produto)
        filtrado = filtrado[chaves.isin(desejados)]
        encontrados = set(chaves)
        faltando = [nome for chave, nome in desejados.items() if chave not in encontrados]
        if faltando:
            raise ValueError(f"Produtos configurados em PRODUTOS_CAPTURADOS que não existem na divulgação "
                             f"(o MDIC pode ter mudado o nome): {faltando}")
    if filtrado.empty:
        raise ValueError(f"Nenhuma linha restou após o filtro (fluxos={FLUXOS_CAPTURADOS})")
    log.info(f"Capturados {len(filtrado)} produtos de {len(produtos)} publicados "
             f"(fluxos: {', '.join(FLUXOS_CAPTURADOS)})")
    return filtrado.reset_index(drop=True)


# {nome do arquivo: bytes} -> (publicação, linhas da tabela).
def processar_divulgacao(arquivos):
    faltando = [nome for nome in PLANILHAS if nome not in arquivos]
    if faltando:
        raise ValueError(f"Arquivos ausentes na divulgação: {faltando}")

    pub, produtos, setores, dias_produtos = extrair_setores_produtos(arquivos[ARQ_SETORES_PRODUTOS])
    resumo = extrair_tabela_resumo(arquivos[ARQ_TABELA_RESUMO], pub=pub)
    validar_somas(produtos, setores, resumo, dias_uteis_produtos=dias_produtos)
    produtos = filtrar_captura(corrigir_variacoes(produtos))

    dados = produtos.assign(
        DATA_REFERENCIA=resumo['data_referencia'],
        SEMANA=pub.semana,
        # calculadas depois, com o histórico (calcular_colunas_historicas)
        VALOR_SEMANA_USD=None, PESO_SEMANA_TON=None, ULTIMA_DO_MES=True, MAIS_RECENTE=True,
    )
    return pub, ajustar_tipos(dados)


# Valores da semana e marcações de recência, que dependem das outras divulgações.
def calcular_colunas_historicas(dados):
    dados = dados.assign(_MES=pd.to_datetime(dados['DATA_REFERENCIA']).dt.to_period('M'))
    chave = ['_MES', 'SEMANA', 'PRODUTO']
    anterior = dados[chave + ['VALOR_ACUMULADO_MES_USD', 'PESO_ACUMULADO_MES_TON']].copy()
    anterior['SEMANA'] += 1
    anterior.columns = chave + ['_VALOR_ANTERIOR', '_PESO_ANTERIOR']
    dados = dados.merge(anterior, on=chave, how='left')

    divulgacoes = set(zip(dados['_MES'], dados['SEMANA']))
    tem_anterior = [(mes, semana - 1) in divulgacoes for mes, semana in zip(dados['_MES'], dados['SEMANA'])]
    primeira = dados['SEMANA'] == 1
    for destino, acumulado, ja_acumulado in (('VALOR_SEMANA_USD', 'VALOR_ACUMULADO_MES_USD', '_VALOR_ANTERIOR'),
                                             ('PESO_SEMANA_TON', 'PESO_ACUMULADO_MES_TON', '_PESO_ANTERIOR')):
        # produto ausente na semana anterior = 0; semana anterior não coletada = vazio
        base = dados[ja_acumulado].fillna(0).where(tem_anterior).mask(primeira, 0)
        dados[destino] = (dados[acumulado] - base).round(3)

    dados['ULTIMA_DO_MES'] = dados['SEMANA'] == dados.groupby('_MES')['SEMANA'].transform('max')
    dados['MAIS_RECENTE'] = dados['DATA_REFERENCIA'] == dados['DATA_REFERENCIA'].max()
    return ajustar_tipos(dados)
