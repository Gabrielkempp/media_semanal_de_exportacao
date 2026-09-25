# Rotina semanal MDIC -> BRADODW_IM_PRICING.BZ_MEDIA_SEMANAL_DE_EXPORTACAO (Nota no bucket da OCI).
#   python main.py              -> baixa, valida, grava e envia a Nota
#   python main.py --conferir   -> mostra o que seria gravado, sem gravar nem enviar nada
import argparse
import logging
import sys
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

from src import banco, bucket, principais_resultados as pr

PASTA_SAIDA = Path(__file__).parent / 'saida'
ARQUIVO_LOG = PASTA_SAIDA / 'extracao.log'

log = logging.getLogger('mdic')


# Downloads em paralelo -> {nome: bytes}.
def baixar(nomes):
    with ThreadPoolExecutor(max_workers=len(nomes)) as pool:
        return dict(zip(nomes, pool.map(pr.baixar, [pr.URLS[nome] for nome in nomes])))


# Baixa a Nota, confere a semana e envia ao bucket.
def enviar_nota(pub, data_referencia):
    conteudo = pr.baixar(pr.URLS[pr.ARQ_NOTA])
    pr.conferir_nota(conteudo, pub)
    bucket.enviar_nota(conteudo, data_referencia)


# Falha vai para o log e não interrompe as outras etapas.
def executar(etapa, funcao, *args, **kwargs):
    try:
        funcao(*args, **kwargs)
        return True
    except Exception:
        log.exception(f"{etapa} falhou")
        return False


def configurar_log():
    PASTA_SAIDA.mkdir(exist_ok=True)
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s %(levelname)s %(message)s',
        handlers=[logging.FileHandler(ARQUIVO_LOG, encoding='utf-8'), logging.StreamHandler()],
    )


def main():
    argumentos = argparse.ArgumentParser(description='Rotina semanal MDIC: baixa, valida, grava no banco e envia a Nota.')
    argumentos.add_argument('--conferir', action='store_true',
                            help='mostra o que seria gravado, sem gravar no banco nem enviar a Nota')
    args = argumentos.parse_args()

    configurar_log()
    try:
        pub, dados = pr.processar_divulgacao(baixar(pr.PLANILHAS))
        log.info(f"Divulgação no site: {pub.descricao}")
    except Exception:
        log.exception("Execução interrompida")
        return 1

    # banco e Nota são independentes; qualquer falha = código 1
    banco_ok = executar("Gravação no banco", banco.gravar, dados, conferir=args.conferir)
    nota_ok = args.conferir or executar("Envio da Nota ao bucket", enviar_nota,
                                        pub, data_referencia=dados['DATA_REFERENCIA'].iloc[0])
    return 0 if banco_ok and nota_ok else 1


if __name__ == "__main__":
    sys.exit(main())
