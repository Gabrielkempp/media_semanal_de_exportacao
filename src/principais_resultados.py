"""
Extração da divulgação semanal "Balança Comercial Preliminar Parcial do Mês" (MDIC/SECEX).

Cada divulgação vira linhas da tabela única DADOS (um produto por linha) e da tabela TEXTOS.
Tudo é lido dos arquivos brutos da divulgação (página HTML + planilhas), que precisam ser da mesma semana.
Setores e totais não são gravados: são a soma dos produtos (isso é validado a cada divulgação).
"""
import io
import re
from dataclasses import dataclass
from html.parser import HTMLParser
from pathlib import PurePosixPath
from urllib.parse import urljoin, urlparse

import pandas as pd
import requests
import urllib3

URL_PAGINA = "https://balanca.economia.gov.br/balanca/pg_principal_bc/principais_resultados.html"
ARQ_PAGINA = 'principais_resultados.html'
ARQ_SETORES_PRODUTOS = 'Setores_Produtos.xlsx'
ARQ_TABELA_RESUMO = 'Tabela_Resumo.xlsx'
# Arquivos usados no processamento; os demais links da página são apenas arquivados
URLS_ESSENCIAIS = {
    ARQ_SETORES_PRODUTOS: urljoin(URL_PAGINA, '../semanal/Setores_Produtos.xlsx'),
    ARQ_TABELA_RESUMO: urljoin(URL_PAGINA, '../semanal/Tabela_Resumo.xlsx'),
}
EXTENSOES_ARQUIVADAS = ('.xlsx', '.xls', '.csv', '.pdf', '.zip')

MESES = ['janeiro', 'fevereiro', 'março', 'abril', 'maio', 'junho',
         'julho', 'agosto', 'setembro', 'outubro', 'novembro', 'dezembro']
MESES_ABREV = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez']

# O site usa "2ª" e "2º" indistintamente
RE_SEMANA = re.compile(r'(\d+)\s*[ªº°]?\s*semana\s+de\s+([a-zç]+)\s*/\s*(\d{4})', re.IGNORECASE)
RE_PASTA = re.compile(r'^(\d{4})-(\d{2})_semana(\d+)$')

EXPORTACAO, IMPORTACAO = 'Exportação', 'Importação'
NOMES_FLUXO = {  # como aparecem na página -> nome padronizado
    'Exportações': EXPORTACAO, 'Importações': IMPORTACAO,
    'Balança Comercial': 'Saldo', 'Corrente de Comércio': 'Corrente',
}

# Dicionário de dados: define as colunas, a ordem e o texto da aba LEIA-ME
DESCRICOES = {
    'DADOS': {
        'DATA_REFERENCIA': 'Último dia coberto pela divulgação: os dados vão do dia 1º do mês até esta data. Use como eixo de tempo.',
        'DATA_PUBLICACAO': 'Data em que o MDIC publicou a divulgação.',
        'ANO': 'Ano de referência dos dados.',
        'MES': 'Mês de referência dos dados.',
        'SEMANA': 'Número da divulgação no mês ("até a Nª semana").',
        'DIAS_UTEIS_ACUMULADO_MES': 'Dias úteis do dia 1º do mês até DATA_REFERENCIA.',
        'DIAS_UTEIS_SEMANA': 'Dias úteis só da semana desta divulgação.',
        'ULTIMA_DO_MES': 'VERDADEIRO na última divulgação coletada de cada mês. Use para comparar meses.',
        'MAIS_RECENTE': 'VERDADEIRO na divulgação mais recente. Use para ver a situação atual.',
        'FLUXO': 'Exportação ou Importação.',
        'SETOR': 'Setor de atividade econômica (classificação do MDIC).',
        'PRODUTO': 'Produto (grupo CUCI), como publicado pelo MDIC.',
        'VALOR_ACUMULADO_MES_USD': 'Valor FOB em US$ do dia 1º do mês até DATA_REFERENCIA.',
        'VALOR_SEMANA_USD': 'Valor FOB em US$ só da semana (acumulado desta divulgação menos o da anterior do mesmo mês). '
                            'Vazio quando a divulgação anterior do mês não foi coletada.',
        'MEDIA_DIARIA_USD': 'Valor por dia útil (VALOR_ACUMULADO_MES_USD ÷ DIAS_UTEIS_ACUMULADO_MES). Use para comparar períodos.',
        'MEDIA_DIARIA_ANO_ANTERIOR_USD': 'Média diária do mesmo mês do ano anterior (mês completo), como publicada pelo MDIC '
                                         'na mesma planilha. É a base da variação.',
        'VARIACAO_VALOR': 'MEDIA_DIARIA_USD ÷ MEDIA_DIARIA_ANO_ANTERIOR_USD − 1, em fração (0,285 = 28,5%). '
                          '−100% = sem embarque no mês; vazio = sem base no ano anterior.',
        'PESO_ACUMULADO_MES_TON': 'Peso em toneladas do dia 1º do mês até DATA_REFERENCIA.',
        'PESO_SEMANA_TON': 'Peso em toneladas só da semana (mesma lógica de VALOR_SEMANA_USD).',
        'MEDIA_DIARIA_TON': 'Toneladas por dia útil.',
        'MEDIA_DIARIA_ANO_ANTERIOR_TON': 'Toneladas por dia útil no mesmo mês do ano anterior (mês completo), como publicado pelo MDIC.',
        'VARIACAO_PESO': 'MEDIA_DIARIA_TON ÷ MEDIA_DIARIA_ANO_ANTERIOR_TON − 1, em fração (mesmas regras de VARIACAO_VALOR).',
        'PRECO_MEDIO_USD_TON': 'Preço médio em US$ por tonelada (valor ÷ peso). Vazio quando não houve peso.',
        'VARIACAO_PRECO': 'Variação do preço médio contra o mesmo mês do ano anterior, em fração. Vazio quando um dos preços não existe.',
    },
    'TEXTOS': {
        'DATA_REFERENCIA': 'Como em DADOS: liga o texto à divulgação.',
        'ANO': 'Ano de referência.',
        'MES': 'Mês de referência.',
        'SEMANA': 'Número da divulgação no mês.',
        'MAIS_RECENTE': 'VERDADEIRO nos textos da divulgação mais recente.',
        'ORDEM': 'Posição do texto na página.',
        'SECAO': 'Seção da página: Destaques, Totais ou Setores e Produtos.',
        'FLUXO': 'Exportação, Importação, Saldo ou Corrente (vazio = texto geral).',
        'SUBSECAO': 'Subtítulo do texto na página.',
        'TEXTO': 'Texto como publicado pelo MDIC.',
    },
}
ESQUEMA = {tabela: list(colunas) for tabela, colunas in DESCRICOES.items()}
COLUNAS_INTEIRAS = {'ANO', 'MES', 'SEMANA', 'DIAS_UTEIS_ACUMULADO_MES', 'DIAS_UTEIS_SEMANA', 'ORDEM'}
COLUNAS_DATA = {'DATA_REFERENCIA', 'DATA_PUBLICACAO'}

# (grupo do cabeçalho, período) -> (coluna de destino, fator). ATUAL = mês da divulgação; ANTERIOR = mesmo mês do ano anterior.
# Valores vão de US$ mil para US$ e variações de % para fração (28,5% -> 0,285), como o Power BI espera.
# Colunas iniciadas por "_" são usadas só no tratamento e não vão para a tabela.
COLUNAS_SETORES_PRODUTOS = {
    ('US$ Mil', 'ATUAL'): ('VALOR_ACUMULADO_MES_USD', 1000),
    ('US$ Mil', 'ANTERIOR'): None,  # total do mês do ano anterior: a média diária já basta como base
    ('US$ Mil Por Média Diária', 'ATUAL'): ('MEDIA_DIARIA_USD', 1000),
    ('US$ Mil Por Média Diária', 'ANTERIOR'): ('MEDIA_DIARIA_ANO_ANTERIOR_USD', 1000),
    ('Toneladas', 'ATUAL'): ('PESO_ACUMULADO_MES_TON', 1),
    ('Toneladas', 'ANTERIOR'): None,
    ('Toneladas por Média Diária', 'ATUAL'): ('MEDIA_DIARIA_TON', 1),
    ('Toneladas por Média Diária', 'ANTERIOR'): ('MEDIA_DIARIA_ANO_ANTERIOR_TON', 1),
    ('Preço (US$/Tonelada)', 'ATUAL'): ('PRECO_MEDIO_USD_TON', 1),
    ('Preço (US$/Tonelada)', 'ANTERIOR'): ('_PRECO_ANO_ANTERIOR', 1),
    ('Variação (%) Por Média Diária', 'Valor US$'): ('VARIACAO_VALOR', 0.01),
    ('Variação (%) Por Média Diária', 'Toneladas'): ('VARIACAO_PESO', 0.01),
    ('Variação (%) Por Média Diária', 'Preço'): ('VARIACAO_PRECO', 0.01),
}

# O servidor do governo não envia a cadeia completa de certificados
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


def id_publicacao(ano, mes, semana):
    """Identificador da divulgação no formato AAAAMMS, ex.: 2ª semana de set/2026 -> 2026092."""
    return ano * 1000 + mes * 10 + semana


def id_da_pasta(nome):
    m = RE_PASTA.match(nome)
    return id_publicacao(int(m.group(1)), int(m.group(2)), int(m.group(3))) if m else None


@dataclass
class Publicacao:
    ano: int
    mes: int
    semana: int
    data_publicacao: pd.Timestamp

    @property
    def id(self):
        return id_publicacao(self.ano, self.mes, self.semana)

    def rotulo_mes(self, ano=None):
        """Rótulo usado nos cabeçalhos das planilhas, ex.: 'Set/2026'."""
        return f"{MESES_ABREV[self.mes - 1]}/{ano or self.ano}"

    @property
    def descricao(self):
        return f"{self.rotulo_mes()} - {self.semana}ª semana"

    @property
    def pasta(self):
        return f"{self.ano}-{self.mes:02d}_semana{self.semana}"


def baixar(url):
    response = requests.get(url, verify=False, timeout=120)
    response.raise_for_status()
    return response.content


def ler_periodo(texto):
    """'Até 2ª Semana de Setembro/2026' -> (2, 9, 2026)."""
    m = RE_SEMANA.search(texto)
    if not m:
        raise ValueError(f"Período de referência não encontrado em: {texto!r}")
    nome_mes = m.group(2).lower()
    if nome_mes not in MESES:
        raise ValueError(f"Mês desconhecido {nome_mes!r} em: {texto!r}")
    return int(m.group(1)), MESES.index(nome_mes) + 1, int(m.group(3))


def validar_periodo(pub, texto, fonte):
    if ler_periodo(texto) != (pub.semana, pub.mes, pub.ano):
        raise ValueError(f"{fonte} não é da mesma divulgação da página ({pub.descricao}): {texto!r}")


def ajustar_tipos(df, tabela):
    """Ordena as colunas conforme o ESQUEMA e padroniza os tipos (inteiros, datas, verdadeiro/falso)."""
    df = df[ESQUEMA[tabela]].copy()
    for coluna in df.columns:
        if coluna in COLUNAS_INTEIRAS:
            df[coluna] = df[coluna].astype('Int64')
        elif coluna in COLUNAS_DATA:
            df[coluna] = pd.to_datetime(df[coluna])
        elif coluna in ('ULTIMA_DO_MES', 'MAIS_RECENTE'):
            df[coluna] = df[coluna].astype(bool)
    return df


# ---------------------------------------------------------------- página HTML

class _ParserPagina(HTMLParser):
    """Lê cabeçalho da publicação, links de download e textos das seções Destaques, Totais e Setores e Produtos."""
    SECOES = {'destaques': 'Destaques', 'totais': 'Totais', 'setores-e-produtos': 'Setores e Produtos'}
    CAPTURADAS = ('p', 'h2', 'h3', 'h4', 'li', 'strong')

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.divs = []            # (id, classes) das divs abertas
        self.titulos = {}         # 'h3'/'h4' -> (texto, profundidade da div que o contém)
        self.captura = None       # (tag, classes) do elemento cujo texto está sendo lido
        self.buffer = []
        self.periodo = None
        self.datas = []
        self.links = []
        self.paragrafos = []      # na ordem em que aparecem na página
        self.destaque = {}        # caixa de destaque em leitura

    def _secao(self):
        return next((self.SECOES[i] for i, _ in self.divs if i in self.SECOES), None)

    def handle_starttag(self, tag, attrs):
        attrs = dict(attrs)
        classes = (attrs.get('class') or '').split()
        if tag == 'div':
            self.divs.append((attrs.get('id'), classes))
        elif tag == 'a' and attrs.get('href'):
            self.links.append(attrs['href'])
        elif tag in self.CAPTURADAS and self.captura is None:
            self.captura = (tag, classes)
            self.buffer = []

    def handle_endtag(self, tag):
        if tag == 'div' and self.divs:
            self.divs.pop()
            self.titulos = {h: t for h, t in self.titulos.items() if t[1] <= len(self.divs)}
            return
        if not self.captura or tag != self.captura[0]:
            return
        _, classes = self.captura
        self.captura = None
        texto = ' '.join(''.join(self.buffer).split())
        secao = self._secao()
        classes_div = self.divs[-1][1] if self.divs else []

        if tag == 'h2' and 'mes' in classes:
            self.periodo = texto
        elif tag == 'h4' and 'date' in classes:
            self.datas.append(texto)
        elif secao == 'Destaques':
            self._ler_destaque(tag, classes, classes_div, texto)
        elif secao and tag in ('h3', 'h4'):
            self.titulos[tag] = (texto, len(self.divs))
        elif secao and tag == 'p' and texto:
            self.paragrafos.append({'SECAO': secao, 'FLUXO': self.titulos.get('h3', (None,))[0],
                                    'SUBSECAO': self.titulos.get('h4', (None,))[0], 'itens': [texto]})

    def _ler_destaque(self, tag, classes, classes_div, texto):
        # Cada caixa tem um título (fluxo), subtítulos (mês / acumulado) e itens <li class="texto">
        if tag == 'p' and 'titulo' in classes_div:
            self.destaque = {'FLUXO': texto}
        elif tag == 'strong' and 'subtitulo' in classes_div:
            self.destaque = {'SECAO': 'Destaques', 'FLUXO': self.destaque.get('FLUXO'), 'SUBSECAO': texto, 'itens': []}
            self.paragrafos.append(self.destaque)
        elif tag == 'li' and 'texto' in classes and 'itens' in self.destaque:
            self.destaque['itens'].append(texto)

    def handle_data(self, data):
        if self.captura:
            self.buffer.append(data)


def analisar_pagina(html):
    parser = _ParserPagina()
    parser.feed(html)
    parser.close()
    return parser


def publicacao_da_pagina(parser):
    if not parser.periodo:
        raise ValueError("Cabeçalho do período (h2.mes) não encontrado na página")
    semana, mes, ano = ler_periodo(parser.periodo)
    atualizado = next((d for d in parser.datas if d.startswith('Atualizado em')), None)
    if not atualizado:
        raise ValueError(f"Data de atualização não encontrada nos h4.date: {parser.datas}")
    data_publicacao = pd.to_datetime(atualizado.split('em ')[1].strip(), format='%d/%m/%Y')
    return Publicacao(ano, mes, semana, data_publicacao)


def links_da_pagina(parser):
    """{nome do arquivo: url} de todos os arquivos para download da página (+ os essenciais)."""
    arquivos = dict(URLS_ESSENCIAIS)
    for href in parser.links:
        url = urljoin(URL_PAGINA, href.strip().strip('"\''))  # o site tem um href com aspas sobrando
        nome = PurePosixPath(urlparse(url).path).name
        if nome.lower().endswith(EXTENSOES_ARQUIVADAS):
            arquivos.setdefault(nome, url)
    return arquivos


def _juntar_itens(itens):
    """['Total:', 'US$ 3,36 bilhões', 'crescimento de 193,8%'] -> 'Total: US$ 3,36 bilhões; crescimento de 193,8%'."""
    if len(itens) > 1 and itens[0].endswith(':'):
        return f"{itens[0]} {'; '.join(itens[1:])}"
    return '; '.join(itens)


def textos_da_pagina(parser):
    secoes = {p['SECAO'] for p in parser.paragrafos}
    if secoes != set(_ParserPagina.SECOES.values()):
        raise ValueError(f"Seções de texto esperadas não encontradas. Encontradas: {secoes}")
    return pd.DataFrame([{
        'ORDEM': ordem,
        'SECAO': p['SECAO'],
        'FLUXO': NOMES_FLUXO.get(p['FLUXO'], p['FLUXO']),
        'SUBSECAO': p['SUBSECAO'],
        'TEXTO': _juntar_itens(p['itens']),
    } for ordem, p in enumerate(parser.paragrafos, start=1)])


# ---------------------------------------------------------------- Setores_Produtos.xlsx

def extrair_setores_produtos(conteudo, pub):
    """Retorna (produtos, setores, dias úteis do mês). Setores são usados só para validar a soma dos produtos."""
    produtos, setores = [], []
    dias_uteis = None
    periodos = {pub.rotulo_mes(): 'ATUAL', pub.rotulo_mes(pub.ano - 1): 'ANTERIOR'}
    for aba, fluxo in (('EXP', EXPORTACAO), ('IMP', IMPORTACAO)):
        df = pd.read_excel(io.BytesIO(conteudo), sheet_name=aba, header=None)

        # Ex.: "EXPORTAÇÃO BRASILEIRA\nCUCI ...\nSet/2026: 8 dias úteis; Set/2025: 22 dias úteis.\nAté 2ª Semana ..."
        titulo = str(df.iat[3, 0])
        validar_periodo(pub, titulo, f"{ARQ_SETORES_PRODUTOS} [{aba}]")
        m = re.search(re.escape(pub.rotulo_mes()) + r':\s*(\d+)\s*dias', titulo)
        dias_uteis = int(m.group(1)) if m else None

        # Cabeçalho: grupo na linha 5 (células mescladas) e período/subtítulo na linha 7
        grupos = df.iloc[4].ffill()
        subtitulos = df.iloc[6]
        if grupos[0] != 'Descrição':
            raise ValueError(f"Layout inesperado em {ARQ_SETORES_PRODUTOS} [{aba}]: {grupos.tolist()}")

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

        setor = None
        for _, linha in df.iloc[8:].iterrows():
            descricao = linha[0]
            if pd.isna(descricao) or str(descricao).startswith('Fonte:'):
                continue
            descricao = ' '.join(str(descricao).split())
            registro = {'FLUXO': fluxo}
            for c, (destino, fator) in colunas.items():
                valor = float(linha[c]) * fator if pd.notna(linha[c]) else None
                # O acumulado vem exato em US$ (3 casas em US$ mil): o arredondamento só tira ruído da multiplicação.
                # Médias diárias ficam com a precisão da fonte, senão a variação de produtos pequenos se distorce.
                registro[destino] = round(valor, 2) if valor is not None and destino == 'VALOR_ACUMULADO_MES_USD' else valor
            m = re.match(r'^[A-Z] - (.+)$', descricao)
            if m:
                setor = m.group(1)
                setores.append({**registro, 'SETOR': setor})
            else:
                produtos.append({**registro, 'SETOR': setor, 'PRODUTO': descricao})

    return pd.DataFrame(produtos), pd.DataFrame(setores), dias_uteis


def corrigir_variacoes(produtos):
    """
    A planilha do MDIC mostra 0% quando a variação é -100% (sem embarque no mês) ou quando não existe base no ano
    anterior (variação indefinida), e preço 0 quando não houve peso. Aqui cada caso recebe o valor correto.
    Nos demais casos confere se a variação publicada bate com média atual ÷ base − 1.
    """
    df = produtos
    for variacao, atual, base in (('VARIACAO_VALOR', 'MEDIA_DIARIA_USD', 'MEDIA_DIARIA_ANO_ANTERIOR_USD'),
                                  ('VARIACAO_PESO', 'MEDIA_DIARIA_TON', 'MEDIA_DIARIA_ANO_ANTERIOR_TON')):
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
    df.loc[sem_preco | (df['_PRECO_ANO_ANTERIOR'].fillna(0) == 0), 'VARIACAO_PRECO'] = None
    return df.drop(columns=['_PRECO_ANO_ANTERIOR'])


# ---------------------------------------------------------------- Tabela_Resumo.xlsx

def extrair_tabela_resumo(conteudo, pub):
    """
    Lê da Tabela_Resumo o que a tabela de produtos não traz: datas e dias úteis de cada semana, e os totais
    do mês (US$) usados para validar a soma dos produtos.
    """
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

    def valores(i):  # exportação e importação em US$ (a planilha está em US$ milhões)
        return {EXPORTACAO: float(df.iat[i, 2]) * 1e6, IMPORTACAO: float(df.iat[i, 4]) * 1e6}

    mes, semanas = None, {}
    for i, rotulo in rotulos.items():
        if m := re.match(rf'^{nome_mes} \(até a (\d+)\s*[ªº°] semana\)$', rotulo):
            if int(m.group(1)) != pub.semana:
                raise ValueError(f"{ARQ_TABELA_RESUMO} não é da mesma divulgação da página: {rotulo!r}")
            mes = {'dias_uteis': int(df.iat[i, 1]), 'valores': valores(i)}
        elif m := re.match(r'^(\d+)\s*[ªº°] semana\s*\((\d{1,2})\s*a\s*(\d{1,2})\)', rotulo):
            semanas[int(m.group(1))] = {'dias_uteis': int(df.iat[i, 1]), 'dia_fim': int(m.group(3)), 'valores': valores(i)}

    if mes is None or set(semanas) != set(range(1, pub.semana + 1)):
        raise ValueError(f"Mês ou semanas não encontrados em {ARQ_TABELA_RESUMO}: semanas={sorted(semanas)}")
    for fluxo in (EXPORTACAO, IMPORTACAO):
        soma = sum(s['valores'][fluxo] for s in semanas.values())
        if abs(soma - mes['valores'][fluxo]) > 10:
            raise ValueError(f"{ARQ_TABELA_RESUMO}: soma das semanas de {fluxo} ({soma:,.2f}) != mês ({mes['valores'][fluxo]:,.2f})")

    atual = semanas[pub.semana]
    return {
        'data_referencia': pd.Timestamp(pub.ano, pub.mes, atual['dia_fim']),
        'dias_uteis_mes': mes['dias_uteis'],
        'dias_uteis_semana': atual['dias_uteis'],
        'total_mes_usd': mes['valores'],
    }


# ---------------------------------------------------------------- processamento

def validar_somas(produtos, setores, resumo, dias_uteis_produtos, tolerancia=10.0):
    """Produtos somam o setor (valor e base do ano anterior) e os setores somam o total da Tabela_Resumo."""
    erros = []
    if produtos.duplicated(['FLUXO', 'PRODUTO']).any():
        erros.append(f"Produtos repetidos: {produtos[produtos.duplicated(['FLUXO', 'PRODUTO'])]['PRODUTO'].tolist()}")
    for fluxo in (EXPORTACAO, IMPORTACAO):
        s = setores[setores['FLUXO'] == fluxo].set_index('SETOR')
        p = produtos[produtos['FLUXO'] == fluxo].groupby('SETOR')
        for coluna in ('VALOR_ACUMULADO_MES_USD', 'MEDIA_DIARIA_ANO_ANTERIOR_USD'):
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


def processar_divulgacao(arquivos):
    """arquivos: {nome do arquivo: bytes} de uma divulgação. Retorna {'DADOS': ..., 'TEXTOS': ...}."""
    faltando = [n for n in (ARQ_PAGINA, *URLS_ESSENCIAIS) if n not in arquivos]
    if faltando:
        raise ValueError(f"Arquivos ausentes na divulgação: {faltando}")

    parser = analisar_pagina(arquivos[ARQ_PAGINA].decode('utf-8'))
    pub = publicacao_da_pagina(parser)
    textos = textos_da_pagina(parser)
    produtos, setores, dias_produtos = extrair_setores_produtos(arquivos[ARQ_SETORES_PRODUTOS], pub)
    resumo = extrair_tabela_resumo(arquivos[ARQ_TABELA_RESUMO], pub)
    validar_somas(produtos, setores, resumo, dias_produtos)
    produtos = corrigir_variacoes(produtos)

    chave = {'DATA_REFERENCIA': resumo['data_referencia'], 'ANO': pub.ano, 'MES': pub.mes, 'SEMANA': pub.semana}
    dados = produtos.assign(
        **chave,
        DATA_PUBLICACAO=pub.data_publicacao,
        DIAS_UTEIS_ACUMULADO_MES=resumo['dias_uteis_mes'],
        DIAS_UTEIS_SEMANA=resumo['dias_uteis_semana'],
        # Dependem das outras divulgações; calculadas em calcular_colunas_historicas
        VALOR_SEMANA_USD=None, PESO_SEMANA_TON=None, ULTIMA_DO_MES=True, MAIS_RECENTE=True,
    )
    textos = textos.assign(**chave, MAIS_RECENTE=True)
    return {'DADOS': ajustar_tipos(dados, 'DADOS'), 'TEXTOS': ajustar_tipos(textos, 'TEXTOS')}


def calcular_colunas_historicas(dados, textos):
    """Colunas que dependem do conjunto de divulgações: valores da semana e as marcações de recência."""
    chave = ['ANO', 'MES', 'SEMANA']
    anterior = dados[chave + ['FLUXO', 'PRODUTO', 'VALOR_ACUMULADO_MES_USD', 'PESO_ACUMULADO_MES_TON']].copy()
    anterior['SEMANA'] += 1
    anterior.columns = chave + ['FLUXO', 'PRODUTO', '_VALOR_ANTERIOR', '_PESO_ANTERIOR']
    dados = dados.merge(anterior, on=chave + ['FLUXO', 'PRODUTO'], how='left')

    divulgacoes = set(zip(dados['ANO'], dados['MES'], dados['SEMANA']))
    tem_anterior = pd.Series([(a, m, s - 1) in divulgacoes for a, m, s in zip(dados['ANO'], dados['MES'], dados['SEMANA'])],
                             index=dados.index)
    primeira = dados['SEMANA'] == 1
    for destino, acumulado, ja_acumulado in (('VALOR_SEMANA_USD', 'VALOR_ACUMULADO_MES_USD', '_VALOR_ANTERIOR'),
                                             ('PESO_SEMANA_TON', 'PESO_ACUMULADO_MES_TON', '_PESO_ANTERIOR')):
        # Produto ausente na divulgação anterior do mês tinha acumulado 0; sem a divulgação anterior, fica vazio
        base = dados[ja_acumulado].fillna(0).where(tem_anterior).mask(primeira, 0)
        dados[destino] = (dados[acumulado] - base).round(3)

    dados['ULTIMA_DO_MES'] = dados['SEMANA'] == dados.groupby(['ANO', 'MES'])['SEMANA'].transform('max')
    mais_recente = dados['DATA_REFERENCIA'].max()
    dados['MAIS_RECENTE'] = dados['DATA_REFERENCIA'] == mais_recente
    textos = textos.assign(MAIS_RECENTE=textos['DATA_REFERENCIA'] == mais_recente)
    return ajustar_tipos(dados, 'DADOS'), ajustar_tipos(textos, 'TEXTOS')
