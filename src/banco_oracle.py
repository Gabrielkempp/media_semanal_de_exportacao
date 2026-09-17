import os
from dotenv import load_dotenv
import pandas as pd

import oracledb

class BancoOracle:
    def __init__(self):
        load_dotenv()  # Carrega variáveis do .env

        oracledb.init_oracle_client(lib_dir=os.getenv("ORACLE_LIB_DIR"))

        self.connection = oracledb.connect(
            user=os.getenv("DB_USER"),
            password=os.getenv("DB_PASSWORD"),
            dsn=os.getenv("DB_DSN"),
            config_dir=os.getenv("ORACLE_WALLET_DIR")
        )
        self.cursor = self.connection.cursor()
        print("✅ Conexão com Oracle estabelecida!")

    def consultar_exportacoes(self, produto=None):
        query_base = """
            SELECT id, PRODUTO, INDICADOR, VALOR, DATA, SEMANA
            FROM MEDIA_DIARIA_EXPORTACAO_MDIC
        """
        if produto:
            query_base += " WHERE PRODUTO = :produto ORDER BY DATA DESC"
            self.cursor.execute(query_base, {"produto": produto})
        else:
            query_base += " ORDER BY DATA DESC"
            self.cursor.execute(query_base)

        return self.cursor.fetchall()

    def inserir_exportacao(self, produto, indicador, valor, data, semana):
        query = """
            INSERT INTO MEDIA_DIARIA_EXPORTACAO_MDIC 
            (PRODUTO, INDICADOR, VALOR, DATA, SEMANA)
            VALUES (:1, :2, :3, :4, :5)
        """
        self.cursor.execute(query, (produto, indicador, valor, data, semana))
        self.connection.commit()
        print("✅ Registro inserido com sucesso!")

    def inserir_exportacoes_em_lote(self, df: pd.DataFrame):
        """
        Insere múltiplos registros na tabela MEDIA_DIARIA_EXPORTACAO_MDIC
        a partir de um DataFrame pandas.
        Espera colunas: PRODUTO, INDICADOR, VALOR, DATA, SEMANA
        """
        print(df)
        required_columns = {"PRODUTO", "INDICADOR", "VALOR", "DATA", "SEMANA"}
        if not required_columns.issubset(df.columns):
            raise ValueError(f"O DataFrame precisa conter as colunas: {required_columns}")

        query = """
            INSERT INTO MEDIA_DIARIA_EXPORTACAO_MDIC 
            (PRODUTO, INDICADOR, VALOR, DATA, SEMANA)
            VALUES (:1, :2, :3, :4, :5)
        """

        # Converte DataFrame para lista de tuplas
        registros = [
            (
                row["PRODUTO"],
                row["INDICADOR"],
                row["VALOR"],
                row["DATA"],
                row["SEMANA"]
            )
            for _, row in df.iterrows()
        ]

        self.cursor.executemany(query, registros)
        self.connection.commit()
        print(f"✅ {len(registros)} registros inseridos com sucesso!")

    def fechar_conexao(self):
        self.cursor.close()
        self.connection.close()
        print("🔒 Conexão encerrada.")

if __name__ == "__main__": 
    banco = BancoOracle()
    resultados = banco.consultar_exportacoes()

    banco.fechar_conexao()
