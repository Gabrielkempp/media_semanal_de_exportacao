from src.balanca_semanal import ExcelFileHandler, BalancaSemanalMDIC
from src.banco_oracle import BancoOracle

def main():
    print('Rodando media MDIC')
    url = "https://balanca.economia.gov.br/balanca/semanal/Setores_Produtos.xlsx"
    filtro = ['Algodão em bruto', 'Madeira em bruto', 'Milho não moído, exceto milho doce']
    file_handler = ExcelFileHandler()
    
    processor = BalancaSemanalMDIC(file_handler, url, filtro)
    ultima_atualizacao = processor.get_ultima_atualizacao()
    
    if processor.verifica_necesidade_download(ultima_atualizacao):
        processor.baixar_arquivo()
        processor.renomear_colunas()
        df_filtrado = processor.filtrar_dataframe()
        df_longo = processor.transforma_de_largo_para_longo(df_filtrado, ultima_atualizacao.split('-')[1], 10)
        # processor.salvar_arquivo(df_longo, ultima_atualizacao)
        banco = BancoOracle()
        banco.inserir_exportacoes_em_lote(df_longo)
        banco.fechar_conexao()


    else:
        print('✅ Não há nada novo para baixar.')

if __name__ == "__main__":
    main()
