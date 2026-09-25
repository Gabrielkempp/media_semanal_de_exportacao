# Teste de conexao com o bucket da OCI. Somente leitura: nao grava, altera nem apaga nada.

import oci

# Configuracao
PERFIL         = "DEFAULT"                      # perfil do ~/.oci/config
BUCKET         = "brado-inteligencia-mercado"   # None = testa apenas a autenticacao
COMPARTMENT_ID = None                           # opcional: OCID do compartimento, para listar os buckets


def linha(ok, etapa, detalhe=""):
    print(f"[{'OK    ' if ok else 'FALHOU'}] {etapa}" + (f" -> {detalhe}" if detalhe else ""))


# 1. Autenticacao (obrigatoria)
try:
    config = oci.config.from_file(profile_name=PERFIL)
    oci.config.validate_config(config)        # levanta excecao se algum campo estiver invalido
    client = oci.object_storage.ObjectStorageClient(config)
    namespace = client.get_namespace().data   # primeira chamada assinada = prova da autenticacao
    linha(True, "1. Autenticacao", f"namespace={namespace} | regiao={config['region']}")
except oci.exceptions.ConfigFileNotFound as e:
    linha(False, "1. Autenticacao", f"config nao encontrado ({e})")
    raise SystemExit(1)
except oci.exceptions.InvalidConfig as e:
    linha(False, "1. Autenticacao", f"campos invalidos no config ({e})")
    raise SystemExit(1)
except oci.exceptions.ServiceError as e:
    linha(False, "1. Autenticacao", f"{e.status} {e.code}: {e.message}")
    raise SystemExit(1)

# 2. Listar buckets (opcional)
if COMPARTMENT_ID:
    try:
        nomes = [b.name for b in client.list_buckets(namespace, COMPARTMENT_ID).data]
        linha(True, "2. Listar buckets", f"{len(nomes)} encontrado(s): {', '.join(nomes) or '-'}")
    except oci.exceptions.ServiceError as e:
        linha(False, "2. Listar buckets", f"{e.status} {e.code}: {e.message}")
else:
    print("[PULADO] 2. Listar buckets (COMPARTMENT_ID nao informado)")

# 3-4. Acesso e leitura do bucket (opcional)
if BUCKET:
    try:
        b = client.get_bucket(namespace, BUCKET).data
        linha(True, "3. Acesso ao bucket", f"{b.name} | compartimento={b.compartment_id}")
    except oci.exceptions.ServiceError as e:
        linha(False, "3. Acesso ao bucket", f"{e.status} {e.code}: {e.message}")

    try:
        objetos = client.list_objects(
            namespace, BUCKET, fields="name,size,timeCreated", limit=5
        ).data.objects
        linha(True, "4. Leitura do conteudo", f"{len(objetos)} objeto(s) na amostra")
        for obj in objetos:
            print(f"         {obj.time_created}  {obj.size:>9} B  {obj.name}")
    except oci.exceptions.ServiceError as e:
        linha(False, "4. Leitura do conteudo", f"{e.status} {e.code}: {e.message}")
else:
    print("[PULADO] 3-4. Checagens de bucket (BUCKET nao informado)")

print("\nConexao e autenticacao verificadas. Nenhum dado foi enviado ao bucket.")