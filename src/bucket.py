# Envio da Nota ao bucket da OCI. Credencial: ~/.oci/config (fora deste projeto).
import base64
import hashlib
import logging

import oci

PERFIL = 'DEFAULT'
BUCKET = 'brado-inteligencia-mercado'

log = logging.getLogger('mdic')


# 2026-09-20 -> 'mdic/2026/Nota-09-20.pdf'.
def caminho_nota(data_referencia):
    return f"mdic/{data_referencia:%Y}/Nota-{data_referencia:%m-%d}.pdf"


# Mesma semana = mesmo nome: reenviar sobrescreve com o mesmo arquivo.
def enviar_nota(conteudo, data_referencia):
    # tenta de novo sozinho em falha temporária (rede, 429, 5xx)
    client = oci.object_storage.ObjectStorageClient(oci.config.from_file(profile_name=PERFIL),
                                                    retry_strategy=oci.retry.DEFAULT_RETRY_STRATEGY)
    destino = caminho_nota(data_referencia)
    # MD5: o servidor recusa o arquivo se chegar corrompido
    md5 = base64.b64encode(hashlib.md5(conteudo).digest()).decode()
    client.put_object(client.get_namespace().data, BUCKET, destino, conteudo,
                      content_type='application/pdf', content_md5=md5)
    log.info(f"Nota enviada: {BUCKET}/{destino}")
    return destino
