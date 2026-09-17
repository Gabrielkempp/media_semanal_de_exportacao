import requests
import pandas as pd
from unidecode import unidecode
from datetime import datetime
from openpyxl import load_workbook, Workbook
from openpyxl.writer.excel import save_workbook
from src.web_scraper import WebScraper

# Interface para o padrão DIP (Princípio da Inversão de Dependência)
class FileHandler:
    def download_file(self, url):
        raise NotImplementedError

    def save_file(self, df, file_path):
        raise NotImplementedError


# Implementação específica para o Excel
class ExcelFileHandler(FileHandler):
    def download_file(self, url):
        response = requests.get(url,verify=False)
        response.raise_for_status()
        return response.content

    def save_file(self, df, file_path):
        df.to_excel(file_path, index=False)

    def set_confidentiality(self, workbook_path):
        try:
            excel = win32.gencache.EnsureDispatch('Excel.Application')
            workbook = excel.Workbooks.Open(workbook_path)
            
            # Propriedades Personalizadas (Formato correto para MSIP/MIP)
            company_guid = 'b69092d7-9252-4aea-8128-47a7c5579722'
            workbook.CustomDocumentProperties.Add(
                Name=f"MSIP_Label_{company_guid}_Enabled",
                LinkToContent=False,
                Type=4,  # Tipo 4 = String/Text
                Value="true"
            )

            workbook.CustomDocumentProperties.Add(
                Name=f"MSIP_Label_{company_guid}_SetDate",
                LinkToContent=False,
                Type=4,
                Value='2024-05-17T14:40:25Z'
            )

            workbook.CustomDocumentProperties.Add(
                Name=f"MSIP_Label_{company_guid}_Method",
                LinkToContent=False,
                Type=4,
                Value="Privileged"
            )

            workbook.CustomDocumentProperties.Add(
                Name=f"MSIP_Label_{company_guid}_Name",
                LinkToContent=False,
                Type=4,
                Value="Interno"
            )

            workbook.CustomDocumentProperties.Add(
                Name=f"MSIP_Label_{company_guid}_SiteId",
                LinkToContent=False,
                Type=4,
                Value='f683419f-a47e-4a1c-a361-004903b6d070'
            )

            workbook.CustomDocumentProperties.Add(
                Name=f"MSIP_Label_{company_guid}_ActionId",
                LinkToContent=False,
                Type=4,
                Value='f22dcdfa-0fa4-48d0-95ad-df53715510cc'
            )

            workbook.CustomDocumentProperties.Add(
                Name=f"MSIP_Label_{company_guid}_ContentBits",
                LinkToContent=False,
                Type=4,
                Value='3'
            )

            # Salvar o arquivo e fechar o Excel
            workbook.Save()
            workbook.Close()
            excel.Quit()

            print(f"✅ Rótulo de confidencialidade aplicado com sucesso!")
        except Exception as e:
            print(f"❌ Erro ao aplicar rótulo de confidencialidade: {e}")

# Classe principal para manipulação e processamento de dados
class BalancaSemanalMDIC:
    def __init__(self, file_handler, url, filtro):
        self.url = url
        self.filtro = filtro
        self.df = None
        self.df_infos = None
        self.arquivo_nome = None
        self.file_handler = file_handler

    def baixar_arquivo(self):
        try:
            print('url:', self.url)
            excel_content = self.file_handler.download_file(self.url)
            print('excel_content: ', excel_content)
            self.df = pd.read_excel(excel_content, engine='openpyxl', skiprows=6)
            self.df_infos = pd.read_excel(excel_content, engine='openpyxl')
            self.arquivo_nome = self.df_infos.iloc[2, 0].split('.')[1]
        except Exception as e:
            print(f"Erro ao baixar ou processar o arquivo: {e}")

    def get_ultima_atualizacao(self):
        url = "https://balanca.economia.gov.br/balanca/pg_principal_bc/principais_resultados.html"
        scraper = WebScraper(url)
        scraper.setup_driver()
        scraper.open_website()
        ultima_atualizacao = scraper.get_first_h4_date_text().split('em ')[1].replace('/', '-')
        scraper.close_driver()
        return ultima_atualizacao

    def verifica_necesidade_download(self, date):
        data_referencia = datetime.strptime(date, "%d-%m-%Y")
        return datetime.now() >= data_referencia

    def renomear_colunas(self):
        novos_nomes = {
            self.df.columns[0]: 'Descrição',
            self.df.columns[1]: 'US$ Mil',
            self.df.columns[2]: 'US$ Mil A-1',
            self.df.columns[3]: 'US$ Mil Por Média Diária',
            self.df.columns[4]: 'US$ Mil Por Média Diária A-1',
            self.df.columns[5]: 'Toneladas',
            self.df.columns[6]: 'Toneladas A-1',
            self.df.columns[7]: 'Toneladas por Média Diária',
            self.df.columns[8]: 'Toneladas por Média Diária A-1',
            self.df.columns[9]: 'Preço (US$/Tonelada)',
            self.df.columns[10]: 'Preço (US$/Tonelada) A-1',
            self.df.columns[11]: 'Variação (%) Por Média Diária Valor US$',
            self.df.columns[12]: 'Variação (%) Por Média Diária Toneladas',
            self.df.columns[13]: 'Variação (%) Por Média Diária Preço'
        }
        self.df.rename(columns=novos_nomes, inplace=True)

    def filtrar_dataframe(self):
        self.df['Descrição_normalizada'] = self.df['Descrição'].apply(self.normalizar_texto)
        filtro_normalizado = [self.normalizar_texto(item) for item in self.filtro]
        df_filtrado = self.df[self.df['Descrição_normalizada'].isin(filtro_normalizado)].drop(columns=['Descrição_normalizada'])
        return df_filtrado

    def normalizar_texto(self, texto):
        return unidecode(texto).lower() if isinstance(texto, str) else texto

    def transforma_de_largo_para_longo(self, df_largo, mes, semana):
        df_longo = df_largo.melt(
            id_vars='Descrição',
            var_name='INDICADOR',
            value_name='VALOR'
        )
        
        df_longo.rename(columns={'Descrição': 'PRODUTO'}, inplace=True)
        # df_longo['DATA'] = str(datetime.now().year) + '-' + mes + '-01'
        df_longo['DATA'] =  datetime(int(datetime.now().year), int(mes),1)
        
        # Extração segura do número da semana
        try:
            semana_str = ''.join(filter(str.isdigit, self.arquivo_nome.split(' ')[1]))
            df_longo['SEMANA'] = int(semana_str)
        except (IndexError, ValueError):
            print(f"⚠️ Erro ao extrair número da semana. Definindo valor padrão.")
            df_longo['SEMANA'] = semana  # Valor padrão caso haja erro

        return df_longo

    def salvar_arquivo(self, df_filtrado, ultima_atualizacao):
        # nome_arquivo_final = f'balanco_semanal_{ultima_atualizacao}_{self.format_text(self.arquivo_nome)}.xlsx'
        # caminho_relativo = f"..\\..\\..\\..\\..\\01.Bases\\15.Exportaca_MDIC\\{nome_arquivo_final}"
        # self.file_handler.save_file(df_filtrado, caminho_relativo)
        # self.file_handler.set_confidentiality(caminho_relativo)
        # print(f"✅ Arquivo salvo com sucesso: {nome_arquivo_final}")
        print('Aqui é para salvar no banco')
        print(df_filtrado)

    def format_text(self, text):
        return text.replace('Até', 'ate').replace('ª', '').replace(' ', '_').replace('/', '-').replace('ç', 'c').replace('\n', '')

