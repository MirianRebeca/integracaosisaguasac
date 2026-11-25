
# -*- coding: utf-8 -*-
"""
Gerador de XML - Controle Semestral SAC / SISAGUA

- Lê os relatórios da aba LABB.
- Obtém o código da forma de abastecimento na coluna E da aba RTs.
- Consulta automaticamente CEVS e pontos de captação no SISAGUA.
- Gera XML para conferência e envio manual pelo SoapUI.
- NÃO envia o controle semestral ao SISAGUA.
"""

from pathlib import Path
from datetime import datetime, date
from decimal import Decimal, InvalidOperation
from getpass import getpass
from urllib.request import Request, urlopen
from urllib.error import HTTPError, URLError
import re
import unicodedata
import xml.etree.ElementTree as ET
from xml.dom import minidom

from openpyxl import load_workbook

import urllib.request
import urllib.error
import time
import http.client


# ============================================================
# CONFIGURAÇÕES
# ============================================================

PASTA = Path(__file__).resolve().parent
ARQUIVO_EXCEL = PASTA / "ld lq preenchido.xlsx"

ABA_LABB = "LABB"
ABA_RTS = "RTs"
ABA_AUTH = "Auth"

COLUNA_CODIGO_SAC = 5  # Coluna E

URL_CONSULTA = (
    "http://sisagua.saude.gov.br/"
    "sisagua/integracao/controlemensal"
)

URL_SEMESTRAL = (
    "http://sisagua.saude.gov.br/"
    "sisagua/integracao/controlesemestral"
)

PASTA_SAIDA = PASTA / "XML_GERADOS"

NS_SOAP = "http://schemas.xmlsoap.org/soap/envelope/"
NS_ENDPOINT = "http://endpoint.ws.sisagua.saude.gov.br/"

ET.register_namespace("soapenv", NS_SOAP)
ET.register_namespace("end", NS_ENDPOINT)

PARAMETROS_SEM_LD_LQ = {
    "CONDUTIVIDADE_ELETRICA",
    "GOSTO_ODOR",
    "PH",
}

#PARA SALVAR

ABA_SISAGUA = "SISAGUA"

PASTA_RELATORIOS = Path(
    r"Y:\setores_\setores\assessoria\vigilancia sanitária"
    r"\planilhas\Relatórios SISAGUA"
)

URL_LOGIN = "https://sisagua.saude.gov.br/sisagua/login.jsf"
URL_CONTROLE_SEMESTRAL = (
    "http://sisagua.saude.gov.br/"
    "sisagua/paginas/seguro/controleSacSemestral/"
    "controleSacSemestral.jsf?faces-redirect=true"
)

# ============================================================
# MAPEAMENTO DAS LINHAS DA LABB
# ============================================================

MAPA = {}


def registrar(linhas, grupo, parametros):
    for linha, parametro in zip(linhas, parametros):
        MAPA[linha] = (grupo, parametro)


registrar(range(10, 23), "SUBSTANCIAS_INORGANICAS", [
    "ANTIMONIO", "ARSENIO", "BARIO", "CADMIO", "CHUMBO",
    "COBRE", "CROMO", "MERCURIO", "NIQUEL", "NITRATO",
    "NITRITO", "SELENIO", "URANIO"
])

registrar(range(25, 41), "SUBSTANCIAS_ORGANICAS", [
    "_1_2_DICLOROETANO", "ACRILAMIDA", "BENZENO",
    "BENZO_A_PIRENO", "CLORETO_VINILA", "DI_2_ETILHEXIL",
    "DICLOROMETANO", "DIOXANO", "EPICLORIDRINA",
    "ETILBENZENO", "PENTACLOROFENOL",
    "TETRACLORETO_CARBONO", "TETRACLOROETENO",
    "TOLUENO", "TRICLOROETENO", "XILENOS"
])

registrar(range(43, 83), "AGROTOXICOS", [
    "_24D", "ALACLOR", "ALDICARBE", "ALDRIN", "AMETRINA",
    "ATRAZINA", "CARBENDAZIM", "CARBOFURANO",
    "CIPROCONAZOL", "CLORDANO", "CLOROTALONIL",
    "CLORPIRIFOS", "DDT_DDD_DDE", "DIFENOCONAZOL",
    "DIMETOATO_OMETOATO", "DIURON", "EPOXICONAZOL",
    "FIPRONIL", "FLUTRIAFOL", "GLIFOSATO_AMPA",
    "HIDROXI_ATRAZINA", "LINDANO", "MALATIONA",
    "MANCOZEBE", "METAMIDOFOS_ACETATO", "METOLACLORO",
    "METRIBUZIM", "MOLINATO", "PARAQUATE", "PICLORAM",
    "PROFENOFOS", "PROPARGITO",
    "PROTIOCONAZOL_PROTICONAZOLDESTIO", "SIMAZINA",
    "TEBUCONAZOL", "TERBUFOS", "TIAMETOXAM",
    "TIODICARBE", "TIRAM", "TRIFLURALINA"
])

registrar(range(85, 94), "OUTROS", [
    "CONDUTIVIDADE_ELETRICA", "COR_VERDADEIRA",
    "DEMANDA_BIOQUIMICA_OXIGENIO",
    "DEMANDA_QUIMICA_OXIGENIO", "FOSFORO_TOTAL",
    "NITROGENIO_AMONIACAL_TOTAL", "OXIGENIO_DISSOLVIDO",
    "PH", "TURBIDEZ"
])

registrar(range(96, 101), "PRODUTOS_SECUNDARIOS_DESINFECCAO", [
    "_2_4_6_TRICLOROFENOL", "_2_4_DICLOROFENOL",
    "ACIDOS_HALOACETICOS", "CLORAMINAS", "TRIHALOMETANOS"
])

registrar(range(103, 105), "RADIOATIVIDADE", [
    "ALFA_TOTAL", "BETA_TOTAL"
])

registrar(range(107, 122), "PARAMETROS_ORGANOLEPTICOS", [
    "_1_2_DICLOROBENZENO", "_1_4_DICLOROBENZENO",
    "ALUMINIO", "AMONIA", "CLORETO", "DUREZA_TOTAL",
    "FERRO", "GOSTO_ODOR", "MANGANES",
    "MONOCLOROBENZENO", "SODIO",
    "SOLIDOS_DISSOLVIDOS_TOTAIS", "SULFATO",
    "SULFETO_HIDROGENIO", "ZINCO"
])

registrar(range(124, 137), "SUBSTANCIAS_INORGANICAS", [
    "ANTIMONIO", "ARSENIO", "BARIO", "CADMIO", "CHUMBO",
    "COBRE", "CROMO", "MERCURIO", "NIQUEL", "NITRATO",
    "NITRITO", "SELENIO", "URANIO"
])

registrar(range(139, 155), "SUBSTANCIAS_ORGANICAS", [
    "_1_2_DICLOROETANO", "ACRILAMIDA", "BENZENO",
    "BENZO_A_PIRENO", "CLORETO_VINILA", "DI_2_ETILHEXIL",
    "DICLOROMETANO", "DIOXANO", "EPICLORIDRINA",
    "ETILBENZENO", "PENTACLOROFENOL",
    "TETRACLORETO_CARBONO", "TETRACLOROETENO",
    "TOLUENO", "TRICLOROETENO", "XILENOS"
])

registrar(range(157, 197), "AGROTOXICOS", [
    "_24D", "ALACLOR", "ALDICARBE", "ALDRIN", "AMETRINA",
    "ATRAZINA", "CARBENDAZIM", "CARBOFURANO",
    "CIPROCONAZOL", "CLORDANO", "CLOROTALONIL",
    "CLORPIRIFOS", "DDT_DDD_DDE", "DIFENOCONAZOL",
    "DIMETOATO_OMETOATO", "DIURON", "EPOXICONAZOL",
    "FIPRONIL", "FLUTRIAFOL", "GLIFOSATO_AMPA",
    "HIDROXI_ATRAZINA", "LINDANO", "MALATIONA",
    "MANCOZEBE", "METAMIDOFOS_ACETATO", "METOLACLORO",
    "METRIBUZIM", "MOLINATO", "PARAQUATE", "PICLORAM",
    "PROFENOFOS", "PROPARGITO",
    "PROTIOCONAZOL_PROTICONAZOLDESTIO", "SIMAZINA",
    "TEBUCONAZOL", "TERBUFOS", "TIAMETOXAM",
    "TIODICARBE", "TIRAM", "TRIFLURALINA"
])


# ============================================================
# FUNÇÕES GERAIS
# ============================================================

def texto(v):
    return "" if v is None else str(v).strip()


def normalizar_nome(v):
    s = texto(v).lower()
    s = unicodedata.normalize("NFD", s)
    s = "".join(
        ch for ch in s
        if unicodedata.category(ch) != "Mn"
    )
    return re.sub(r"\s+", " ", s).strip()


def chave_cliente(v):
    """
    Compara nomes ignorando espaços, pontuação e diferenças
    como 'Copa Energia S A' e 'Copa Energia SA'.
    """
    s = normalizar_nome(v)
    return re.sub(r"[^a-z0-9]", "", s)


def extrair_data(v):
    if isinstance(v, datetime):
        return v.strftime("%d/%m/%Y")
    if isinstance(v, date):
        return v.strftime("%d/%m/%Y")

    m = re.search(
        r"(?<!\d)(\d{2}/\d{2}/\d{4})(?!\d)",
        texto(v)
    )
    return m.group(1) if m else None


def data_obj(s):
    return datetime.strptime(s, "%d/%m/%Y").date()


def decimal_valor(v):
    if v is None:
        return None

    s = texto(v)

    if not s or s in {"-", "—"}:
        return None

    s = s.replace(" ", "")

    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    else:
        s = s.replace(",", ".")

    try:
        return Decimal(s)
    except InvalidOperation:
        return None


def formatar_numero(v, campo):
    d = decimal_valor(v)

    if d is None:
        return None

    casas = max(0, -d.as_tuple().exponent)
    antes = len(
        str(abs(d).quantize(Decimal(1), rounding="ROUND_DOWN"))
    )

    if antes > 5 or casas > 6:
        raise ValueError(
            f"{campo}: valor {v!r} excede o limite "
            "de 5 dígitos inteiros/6 decimais."
        )

    s = format(d, "f")

    if "." in s:
        s = s.rstrip("0").rstrip(".")

    return s.replace(".", ",")


def resultado_xml(v, parametro):
    if v is None or texto(v) == "":
        return "MENOR_LQ"

    s = texto(v)

    if s.startswith("<"):
        return "MENOR_LQ"

    if parametro == "GOSTO_ODOR":
        m = re.match(r"^\s*(-?\d+(?:[,.]\d+)?)", s)
        if m:
            return formatar_numero(m.group(1), "Resultado")

    n = formatar_numero(s, "Resultado")

    if n is None:
        raise ValueError(f"Resultado não reconhecido: {s!r}")

    return n

def criar_elemento(pai, nome, valor):
    e = ET.SubElement(pai, nome)
    e.text = str(valor)
    return e


def pedir_opcao(titulo, opcoes):
    print("\n" + titulo)

    for i, item in enumerate(opcoes, 1):
        print(f"  {i}. {item}")

    while True:
        escolha = input("Escolha o número: ").strip()

        if escolha.isdigit() and 1 <= int(escolha) <= len(opcoes):
            return opcoes[int(escolha) - 1]

        print("Opção inválida.")


def pedir_texto(rotulo, padrao=None):
    sufixo = f" [{padrao}]" if padrao else ""
    v = input(f"{rotulo}{sufixo}: ").strip()
    return v or (padrao or "")


# ============================================================
# LOCALIZAÇÃO DO CÓDIGO SAC NA ABA RTS
# ============================================================

def localizar_codigo_sac(ws_rts, cliente):
    """
    Procura o nome do cliente em qualquer célula da linha
    da aba RTs. O código SAC é sempre lido da coluna E.
    """

    chave = chave_cliente(cliente)

    if not chave:
        raise ValueError("Nome do cliente vazio.")

    # Confere o cabeçalho da coluna E.
    cabecalho = normalizar_nome(
        ws_rts.cell(1, COLUNA_CODIGO_SAC).value
    )

    if cabecalho != "codig forma abastecimento".replace(" ", ""):
        if cabecalho.replace(" ", "") != "codigoformaabastecimento":
            raise ValueError(
                "Não encontrei o cabeçalho "
                "'codigoFormaAbastecimento' na célula E1 da aba RTs."
            )

    encontrados = []

    for linha in range(2, ws_rts.max_row + 1):
        nome_encontrado = False

        for coluna in range(1, ws_rts.max_column + 1):
            valor = texto(ws_rts.cell(linha, coluna).value)

            if chave_cliente(valor) == chave:
                nome_encontrado = True
                break

        if nome_encontrado:
            codigo = texto(
                ws_rts.cell(linha, COLUNA_CODIGO_SAC).value
            )

            if codigo:
                encontrados.append(codigo)

    codigos = sorted(set(encontrados))

    if not codigos:
        raise ValueError(
            f"O cliente '{cliente}' não possui "
            "codigoFormaAbastecimento preenchido na coluna E da aba RTs. "
            "Consulta cancelada."
        )

    if len(codigos) > 1:
        raise ValueError(
            f"Foram encontrados códigos SAC diferentes para '{cliente}': "
            f"{', '.join(codigos)}. Verifique a aba RTs."
        )

    return codigos[0]


# ============================================================
# AUTENTICAÇÃO
# ============================================================

def obter_credenciais(wb, rt):
    """
    Se existir a aba Auth:
      A = RT
      B = Username
      C = Password

    Se não houver cadastro correspondente, solicita os dados
    sem exibir a senha na tela.
    """

    if ABA_AUTH in wb.sheetnames:
        ws = wb[ABA_AUTH]

        for linha in range(2, ws.max_row + 1):
            rt_planilha = texto(ws.cell(linha, 1).value)

            if normalizar_nome(rt_planilha) == normalizar_nome(rt):
                usuario = texto(ws.cell(linha, 2).value)
                senha = texto(ws.cell(linha, 3).value)

                if usuario and senha:
                    return usuario, senha

    print("\nCredenciais SISAGUA não encontradas na aba Auth.")

    usuario = input("Username SISAGUA: ").strip()
    senha = getpass("Password SISAGUA: ")

    if not usuario or not senha:
        raise ValueError("Username e Password são obrigatórios.")

    return usuario, senha


# ============================================================
# CONSULTA AUTOMÁTICA AO SISAGUA
# ============================================================

def nome_local(elemento):
    return elemento.tag.split("}")[-1].lower()


def consultar_sisagua(codigo_sac, ano, usuario, senha):
    """
    Consulta a forma de abastecimento.
    Esta função realiza apenas uma operação de consulta.
    Não envia dados de controle semestral.
    """

    envelope = ET.Element(f"{{{NS_SOAP}}}Envelope")
    ET.SubElement(envelope, f"{{{NS_SOAP}}}Header")

    body = ET.SubElement(envelope, f"{{{NS_SOAP}}}Body")

    operacao = ET.SubElement(
        body,
        f"{{{NS_ENDPOINT}}}consultarFormaAbast"
    )

    parametros = ET.SubElement(operacao, "parametros")

    criar_elemento(
        parametros,
        "codigoFormaAbastecimento",
        codigo_sac
    )

    criar_elemento(
        parametros,
        "anoReferencia",
        ano
    )

    xml_envio = ET.tostring(
        envelope,
        encoding="utf-8",
        xml_declaration=True
    )

    requisicao = Request(
        URL_CONSULTA,
        data=xml_envio,
        headers={
            "Content-Type": "text/xml; charset=utf-8",
            "Username": usuario,
            "Password": senha,
        },
        method="POST"
    )

    print("\nConsultando forma de abastecimento no SISAGUA...")

    try:
        with urlopen(requisicao, timeout=60) as resposta:
            xml_resposta = resposta.read()

    except HTTPError as erro:
        detalhe = erro.read().decode("utf-8", errors="replace")
        raise RuntimeError(
            f"Erro HTTP {erro.code} na consulta SISAGUA:\n{detalhe}"
        )

    except URLError as erro:
        raise RuntimeError(
            f"Não foi possível conectar ao SISAGUA: {erro.reason}"
        )

    try:
        raiz = ET.fromstring(xml_resposta)
    except ET.ParseError:
        raise RuntimeError(
            "O SISAGUA retornou uma resposta que não é XML válido."
        )

    # Verifica falha SOAP.
    for elemento in raiz.iter():
        if nome_local(elemento) == "faultstring":
            raise RuntimeError(
                "Falha SOAP na consulta: " + texto(elemento.text)
            )

    valores = {}

    for elemento in raiz.iter():
        nome = nome_local(elemento)
        valor = texto(elemento.text)

        if valor:
            valores.setdefault(nome, []).append(valor)

    # Verifica o retorno de sucesso.
    sucesso = valores.get("sucesso", [])

    if sucesso and sucesso[-1].lower() in {"false", "0"}:
        mensagem = valores.get("mensagem", ["Sem mensagem"])[-1]
        raise RuntimeError(
            f"O SISAGUA não confirmou a consulta: {mensagem}"
        )

    cevs_lista = valores.get("cevs", [])

    if not cevs_lista:
        mensagem = valores.get("mensagem", [""])[-1]
        raise RuntimeError(
            "Não foi possível localizar o CEVS na resposta do SISAGUA. "
            f"Mensagem: {mensagem}\n"
            "Resposta recebida:\n"
            + xml_resposta.decode("utf-8", errors="replace")[:4000]
        )

    cevs = cevs_lista[0]

    
    # Localiza os pontos de captação na estrutura retornada pelo SISAGUA.
    pontos = []

    for elemento in raiz.iter():
        if nome_local(elemento) != "pontocaptacao":
            continue

        dados_ponto = {}

        for filho in elemento:
            nome_filho = nome_local(filho)
            valor = texto(filho.text)

            if nome_filho == "codigo":
                dados_ponto["codigo"] = valor
            elif nome_filho == "nome":
                dados_ponto["nome"] = valor
            elif nome_filho == "categoria":
                dados_ponto["categoria"] = valor
            elif nome_filho == "tipocaptacao":
                dados_ponto["tipo"] = valor

        if dados_ponto.get("codigo"):
            pontos.append(dados_ponto)


    print("\nConsulta realizada.")
    print(f"CEVS retornado: {cevs}")

    if pontos:
        print("\nPontos de captação encontrados:")

        for p in pontos:
            descricao = p["nome"] or "Sem nome informado"
            categoria = p["categoria"]

            print(
                f"  Código: {p['codigo']} | "
                f"{descricao} | {categoria}"
            )
    else:
        print(
            "\nA resposta não apresentou códigos de ponto de captação."
        )

    return cevs, pontos


def testar_conexao_sisagua():
    servidor = "sisagua.saude.gov.br"
    caminho = "/sisagua/integracao/controlesemestral"

    conexao = http.client.HTTPConnection(
        servidor,
        port=80,
        timeout=30
    )

    try:
        print("\n========== TESTE DE CONEXÃO SISAGUA ==========")

        inicio = time.perf_counter()

        conexao.connect()

        tempo_conexao = time.perf_counter() - inicio

        print(
            f"Tempo para estabelecer conexão: "
            f"{tempo_conexao:.2f} segundos"
        )

        inicio_resposta = time.perf_counter()

        conexao.request("HEAD", caminho)

        resposta = conexao.getresponse()

        tempo_resposta = time.perf_counter() - inicio_resposta

        print(f"Status HTTP: {resposta.status}")
        print(
            f"Tempo até resposta do servidor: "
            f"{tempo_resposta:.2f} segundos"
        )

        print("==============================================")

    except Exception as erro:
        print(f"Erro no teste: {repr(erro)}")

    finally:
        conexao.close()
testar_conexao_sisagua()

def enviar_xml(arquivo_xml, usuario, senha):
    """Envia o XML semestral ao SISAGUA."""

    inicio_total = time.perf_counter()

    arquivo_xml = Path(arquivo_xml)

    if not arquivo_xml.exists():
        raise FileNotFoundError(
            f"Arquivo XML não encontrado: {arquivo_xml}"
        )

    # Lê exatamente o XML que foi gerado.
    conteudo_xml = arquivo_xml.read_bytes()

    print(f"\nTamanho do XML: {len(conteudo_xml) / 1024:.2f} KB")
    print(
        f"Leitura do arquivo: "
        f"{time.perf_counter() - inicio_total:.2f} segundos"
    )

    headers = {
        "Content-Type": "text/xml; charset=utf-8",
        "Username": usuario,
        "Password": senha
    }

    requisicao = urllib.request.Request(
        URL_SEMESTRAL,
        data=conteudo_xml,
        headers=headers,
        method="POST"
    )

    print("\nEnviando XML ao SISAGUA...")

    
    inicio_envio = time.perf_counter()

    try:
        with urllib.request.urlopen(
            requisicao,
            timeout=180
        ) as resposta:

            tempo_ate_resposta = time.perf_counter()

            retorno = resposta.read().decode("utf-8")

            tempo_leitura_resposta = time.perf_counter()

            print(
                f"Tempo até receber os cabeçalhos: "
                f"{tempo_ate_resposta - inicio_envio:.2f} segundos"
            )

            print(
                f"Tempo para receber o conteúdo: "
                f"{tempo_leitura_resposta - tempo_ate_resposta:.2f} segundos"
            )

            print(
                f"Tempo total da comunicação: "
                f"{tempo_leitura_resposta - inicio_envio:.2f} segundos"
            )

    except urllib.error.HTTPError as erro:
        detalhe = erro.read().decode(
            "utf-8", errors="replace"
        )

        raise RuntimeError(
            f"Erro HTTP {erro.code} no envio:\n{detalhe}"
        )

    except TimeoutError as erro:
        tempo_decorrido = time.perf_counter() - inicio_envio
        raise RuntimeError(
            "Tempo limite de 180 segundos excedido durante a comunicação "
            "com o SISAGUA. O servidor pode ter recebido o XML e não ter "
            "conseguido retornar a confirmação. Antes de tentar novamente, "
            "verifique no SISAGUA se o lançamento foi registrado. "
            f"Tempo decorrido: {tempo_decorrido:.2f} segundos."
        ) from erro

    except urllib.error.URLError as erro:
        raise RuntimeError(
            f"Falha de conexão com o SISAGUA: {erro.reason}"
        )


    # Interpreta a resposta SOAP.
    inicio_leitura = time.perf_counter()

    raiz = ET.fromstring(retorno)

    dados = {}

    for elemento in raiz.iter():
        nome = nome_local(elemento)

        if nome in ("codigo", "mensagem", "sucesso"):
            dados[nome] = texto(elemento.text)

    print(
        f"Interpretação da resposta: "
        f"{time.perf_counter() - inicio_leitura:.2f} segundos"
    )

    codigo = dados.get("codigo", "")
    mensagem = dados.get("mensagem", "")
    sucesso = dados.get("sucesso", "").lower()

    print("\n========== RETORNO SISAGUA ==========")
    print(f"Código: {codigo}")
    print(f"Mensagem: {mensagem}")
    print(f"Sucesso: {sucesso}")
    print("=====================================")

    print(
        f"Tempo total da função: "
        f"{time.perf_counter() - inicio_total:.2f} segundos"
    )

    if codigo == "200" and sucesso == "true":
        print("\nXML enviado e aceito pelo SISAGUA!")
        return True

    print("\nO SISAGUA não confirmou o recebimento.")
    return False



# ============================================================
# SELEÇÃO DOS RELATÓRIOS
# ============================================================

def listar_relatorios(ws):
    relatorios = []
    clientes = {}

    for col_data in range(4, ws.max_column + 1, 2):
        cliente = texto(ws.cell(2, col_data).value)
        if not cliente:
            continue

        coleta = extrair_data(ws.cell(4, col_data).value)
        if not coleta:
            continue

        registro = {
            "col_data": col_data,
            "col_resultado": col_data + 1,
            "cliente": cliente,
            "cliente_norm": chave_cliente(cliente),
            "tipo": texto(ws.cell(3, col_data).value),
            "coleta": coleta,
            "rt": texto(ws.cell(5, col_data).value),
            "cargo": texto(ws.cell(6, col_data).value),
            "nome_relatorio": texto(ws.cell(1, col_data).value),
        }
        relatorios.append(registro)
        clientes.setdefault(registro["cliente_norm"], cliente)

    if not relatorios:
        raise ValueError("Não encontrei relatórios válidos na aba LABB.")

    nomes = sorted(clientes.values(), key=str.casefold)
    print("\nCLIENTES ENCONTRADOS NA ABA LABB")
    for i, nome in enumerate(nomes, 1):
        qtd = sum(1 for r in relatorios if r["cliente_norm"] == chave_cliente(nome))
        print(f"  {i}. {nome} ({qtd} relatórios)")

    while True:
        escolha = input(
            "\nDigite os números dos clientes separados por vírgula "
            "ou TODOS para processar a lista inteira: "
        ).strip()

        if escolha.upper() == "TODOS":
            escolhidos = nomes
            break

        try:
            indices = [int(x.strip()) for x in escolha.split(",") if x.strip()]
            if not indices or len(set(indices)) != len(indices):
                raise ValueError
            if any(i < 1 or i > len(nomes) for i in indices):
                raise ValueError
            escolhidos = [nomes[i - 1] for i in indices]
            break
        except ValueError:
            print("Seleção inválida. Exemplo: 1, 3, 5 ou TODOS.")

    return [
        (
            nome,
            [r for r in relatorios if r["cliente_norm"] == chave_cliente(nome)]
        )
        for nome in escolhidos
    ]


# ============================================================
# GERAÇÃO DO XML
# ============================================================

def gerar_xml(
    ws,
    cliente_escolhido,
    selecionados,
    codigo_sac,
    cevs,
    codigo_ponto_captacao,
    ano,
    semestre,
    data_preenchimento,
    rt,
    cargo
):

    data_preenchimento_obj = data_obj(data_preenchimento)

    inicio_semestre = date(
        int(ano),
        1 if semestre == "1" else 7,
        1
    )

    fim_semestre = date(
        int(ano),
        6 if semestre == "1" else 12,
        30 if semestre == "1" else 31
    )

    if data_preenchimento_obj > date.today():
        raise ValueError(
            "Data de preenchimento não pode estar no futuro."
        )

    envelope = ET.Element(f"{{{NS_SOAP}}}Envelope")
    ET.SubElement(envelope, f"{{{NS_SOAP}}}Header")

    body = ET.SubElement(envelope, f"{{{NS_SOAP}}}Body")

    operacao = ET.SubElement(
        body,
        f"{{{NS_ENDPOINT}}}empresaEnviarControleSacSemestral"
    )

    solucao = ET.SubElement(
        operacao,
        "solucaoAbastecimentoSAC"
    )

    criar_elemento(
        solucao,
        "codigoFormaAbastecimento",
        codigo_sac
    )

    criar_elemento(
        solucao,
        "anoReferencia",
        ano
    )

    periodo_container = ET.SubElement(solucao, "periodos")
    periodo = ET.SubElement(periodo_container, "periodo")

    criar_elemento(periodo, "semestre", semestre)

    tratamento = ET.SubElement(periodo, "tratamentoAgua")

    criar_elemento(tratamento, "cevs", cevs)
    criar_elemento(
        tratamento,
        "dataPreenchimento",
        data_preenchimento
    )
    criar_elemento(
        tratamento,        "responsavelPelasInformacoes",
        rt
    )
    criar_elemento(
        tratamento,
        "cargoResponsavel",
        cargo
    )
    criar_elemento(tratamento, "naoOperouSemestre", "false")

    analises = ET.SubElement(
        tratamento,
        "analisesNovaNorma"
    )

    total_parametros = 0
    total_campos = 0
    resumo_pontos = {}
    pendencias = []
    usados = set()

    for rel in selecionados:

        tipo_norm = normalizar_nome(rel["tipo"])
        agua_bruta = tipo_norm.startswith("agua bruta")

        ponto_padrao = (
            "PONTO_CAPTACAO"
            if agua_bruta
            else "SAIDA_DO_TRATAMENTO"
        )

        try:
            data_coleta = data_obj(rel["coleta"])
        except ValueError:
            pendencias.append(
                f"{rel['nome_relatorio']}: data de coleta inválida."
            )
            continue

        if not (inicio_semestre <= data_coleta <= fim_semestre):
            pendencias.append(
                f"{rel['nome_relatorio']}: coleta {rel['coleta']} "
                f"fora do semestre {semestre}/{ano}."
            )
            continue

        if data_coleta > data_preenchimento_obj:
            pendencias.append(
                f"{rel['nome_relatorio']}: coleta posterior "
                "à data de preenchimento."
            )
            continue

        faixa = (
            range(10, 94)
            if agua_bruta
            else range(96, 197)
        )

        for linha in faixa:

            if linha not in MAPA:
                continue

            grupo, parametro = MAPA[linha]

            data_analise_txt = extrair_data(
                ws.cell(linha, rel["col_data"]).value
            )

            if not data_analise_txt:
                continue

            try:
                data_analise = data_obj(data_analise_txt)
            except ValueError:
                pendencias.append(
                    f"{rel['nome_relatorio']} / linha {linha}: "
                    "data de análise inválida."
                )
                continue

            if data_analise < data_coleta:
                pendencias.append(
                    f"{rel['nome_relatorio']} / "
                    f"{texto(ws.cell(linha, 1).value)}: "
                    "data de análise anterior à coleta."
                )
                continue

            if not (
                inicio_semestre <= data_analise <= fim_semestre
            ):
                pendencias.append(
                    f"{rel['nome_relatorio']} / "
                    f"{texto(ws.cell(linha, 1).value)}: "
                    f"análise {data_analise_txt} fora do semestre."
                )
                continue

            if data_analise > data_preenchimento_obj:
                pendencias.append(
                    f"{rel['nome_relatorio']} / "
                    f"{texto(ws.cell(linha, 1).value)}: "
                    "análise posterior ao preenchimento."
                )
                continue

            if grupo == "OUTROS" and ponto_padrao != "PONTO_CAPTACAO":
                pendencias.append(
                    f"{rel['nome_relatorio']} / "
                    f"{texto(ws.cell(linha, 1).value)}: "
                    "grupo OUTROS não enviado fora do ponto de captação."
                )
                continue

            ponto = (
                "PONTO_DE_CONSUMO"
                if grupo == "RADIOATIVIDADE"
                else ponto_padrao
            )

            resultado_bruto = ws.cell(
                linha,
                rel["col_resultado"]
            ).value

            resultado = resultado_xml(
                resultado_bruto,
                parametro
            )

            ld = None
            lq = None

            if parametro not in PARAMETROS_SEM_LD_LQ:
                ld = formatar_numero(
                    ws.cell(linha, 2).value,
                    "LD"
                )

                lq = formatar_numero(
                    ws.cell(linha, 3).value,
                    "LQ"
                )

                ld_dec = decimal_valor(
                    ws.cell(linha, 2).value
                )

                lq_dec = decimal_valor(
                    ws.cell(linha, 3).value
                )

                if (
                    ld_dec is not None
                    and lq_dec is not None
                    and lq_dec < ld_dec
                ):
                    pendencias.append(
                        f"Linha {linha} / "
                        f"{texto(ws.cell(linha, 1).value)}: "
                        "LQ menor que LD."
                    )
                    continue

            campos = [
                ("Data da coleta", rel["coleta"]),
                ("Data da análise", data_analise_txt),
                ("Resultado", resultado),
            ]

            if ld is not None:
                campos.append(("LD", ld))

            if lq is not None:
                campos.append(("LQ", lq))

            for campo, valor in campos:
                item = ET.SubElement(
                    analises,
                    "analiseNovaNorma"
                )

                criar_elemento(item, "grupoParametro", grupo)
                criar_elemento(item, "parametro", parametro)
                criar_elemento(item, "pontoMonitoramento", ponto)
                criar_elemento(item, "campo", campo)
                criar_elemento(item, "valor", valor)

                if ponto == "PONTO_CAPTACAO":
                    criar_elemento(
                        item,
                        "codigoPontoCaptacao",
                        codigo_ponto_captacao
                    )

                total_campos += 1

            chave = (rel["col_data"], linha)

            if chave not in usados:
                usados.add(chave)
                total_parametros += 1

            resumo_pontos[ponto] = (
                resumo_pontos.get(ponto, 0) + 1
            )

    if total_parametros == 0:
        raise ValueError(
            "Nenhum parâmetro elegível foi encontrado para gerar o XML."
        )

    if pendencias:
        print("\nATENÇÃO: foram encontradas pendências:")

        for p in pendencias:
            print(" -", p)

        resposta = input(
            "\nDeseja gerar o XML mesmo assim, "
            "incluindo apenas os registros válidos? (S/N): "
        ).strip().upper()

        if resposta != "S":
            print("Geração cancelada.")
            return

    PASTA_SAIDA.mkdir(exist_ok=True)

    nome_base = re.sub(
        r"[^A-Za-z0-9_-]+",
        "_",
        cliente_escolhido
    ).strip("_")

    arquivo_xml = (
        PASTA_SAIDA
        / f"{nome_base}_SAC_{ano}_{semestre}.xml"
    )

    bruto = ET.tostring(
        envelope,
        encoding="utf-8"
    )

    bonito = minidom.parseString(
        bruto
    ).toprettyxml(
        indent="   ",
        encoding="utf-8"
    )

    arquivo_xml.write_bytes(bonito)

    arquivo_log = arquivo_xml.with_suffix(".txt")

    linhas_log = [
        f"Cliente: {cliente_escolhido}",
        f"Código SAC: {codigo_sac}",
        f"CEVS: {cevs}",
        f"Ano/semestre: {ano}/{semestre}",
        f"Responsável: {rt}",
        f"Cargo: {cargo}",
        f"Relatórios considerados: {len(selecionados)}",
        f"Parâmetros incluídos: {total_parametros}",
        f"Registros de campo no XML: {total_campos}",
        "Pontos de monitoramento:"
    ]

    for ponto, qtd in sorted(resumo_pontos.items()):
        linhas_log.append(
            f"  {ponto}: {qtd} parâmetros"
        )

    if pendencias:
        linhas_log.append("\nPENDÊNCIAS:")
        linhas_log.extend(
            f"- {p}" for p in pendencias
        )

    arquivo_log.write_text(
        "\n".join(linhas_log),
        encoding="utf-8"
    )

    print("\nXML GERADO PARA CONFERÊNCIA")
    print(f"Arquivo: {arquivo_xml}")
    print(f"Parâmetros incluídos: {total_parametros}")
    print(f"Campos no XML: {total_campos}")
    print(f"Relatórios considerados: {len(selecionados)}")
    print(f"Resumo: {arquivo_log}")

    return arquivo_xml


#======================
#PARA SALVAR#======================

def localizar_estabelecimento_sisagua(wb, cliente):
    if ABA_SISAGUA not in wb.sheetnames:
        raise ValueError(
            f"A aba {ABA_SISAGUA!r} não existe na planilha."
        )

    ws = wb[ABA_SISAGUA]
    chave = chave_cliente(cliente)

    encontrados = []

    for linha in range(2, ws.max_row + 1):
        cliente_planilha = texto(ws.cell(linha, 1).value)

        if chave_cliente(cliente_planilha) == chave:
            nome_sisagua = texto(ws.cell(linha, 2).value)
            perfil = texto(ws.cell(linha, 3).value)

            encontrados.append({
                "linha": linha,
                "nome": nome_sisagua,
                "perfil": perfil
            })

    if not encontrados:
        raise ValueError(
            f"Cliente '{cliente}' não encontrado na aba SISAGUA."
        )

    if len(encontrados) > 1:
        raise ValueError(
            f"Há mais de um cadastro para '{cliente}' "
            f"na aba SISAGUA. Linhas: "
            f"{[x['linha'] for x in encontrados]}"
        )

    cadastro = encontrados[0]

    if not cadastro["nome"] or not cadastro["perfil"]:
        raise ValueError(
            f"Cadastro incompleto na aba SISAGUA, "
            f"linha {cadastro['linha']}."
        )

    perfis_validos = {
        "2.3 empresa regional municipal",
        "3 empresa municipal"
    }

    if normalizar_nome(cadastro["perfil"]) not in perfis_validos:
        raise ValueError(
            f"Perfil inválido na linha {cadastro['linha']}: "
            f"{cadastro['perfil']}"
        )

    return cadastro["nome"], cadastro["perfil"]


def aguardar_download(arquivo, timeout=90):
    import time

    arquivo = Path(arquivo)
    temporario = Path(str(arquivo) + ".crdownload")

    inicio = time.time()

    while time.time() - inicio < timeout:
        if arquivo.exists() and not temporario.exists():
            if arquivo.stat().st_size > 0:
                return True

        time.sleep(1)

    raise TimeoutError(
        f"O PDF não foi baixado dentro de {timeout} segundos."
    )


def clicar_por_texto(driver, texto_botao, timeout=20):
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    xpath = (
        f"//*[self::button or self::a or self::span]"
        f"[normalize-space(.)='{texto_botao}']"
    )

    elementos = WebDriverWait(driver, timeout).until(
        EC.presence_of_all_elements_located(
            (By.XPATH, xpath)
        )
    )

    visiveis = [e for e in elementos if e.is_displayed()]

    if not visiveis:
        raise RuntimeError(
            f"Não encontrei o botão '{texto_botao}'."
        )

    if len(visiveis) > 1:
        # Prefere o botão ou link, evitando spans internos duplicados.
        principais = [
            e for e in visiveis
            if e.tag_name in ("button", "a")
        ]

        if len(principais) == 1:
            visiveis = principais
        else:
            raise RuntimeError(
                f"Encontrei mais de um elemento visível "
                f"com o texto '{texto_botao}'."
            )

    visiveis[0].click()


def iniciar_chrome():
    from selenium import webdriver

    PASTA_RELATORIOS.mkdir(parents=True, exist_ok=True)

    opcoes = webdriver.ChromeOptions()

    preferencias = {
        "download.default_directory": str(PASTA_RELATORIOS.resolve()),
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True,
        "safebrowsing.enabled": True
    }

    opcoes.add_experimental_option(
        "prefs",
        preferencias
    )

    opcoes.add_argument("--start-maximized")

    return webdriver.Chrome(options=opcoes)


def baixar_relatorio_sisagua(
    cliente,
    ano,
    semestre,
    nome_estabelecimento,
    perfil,
    usuario,
    senha
):
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.common.exceptions import TimeoutException

    driver = None

    nome_arquivo_original = (
        PASTA_RELATORIOS
        / "relatorioControleSemestralAnoVigente.pdf"
    )

    nome_cliente = re.sub(
        r'[\\/:*?"<>|]',
        "_",
        cliente
    ).strip()

    nome_semestre = (
        "1º Semestre"
        if semestre == "1"
        else "2º Semestre"
    )

    arquivo_final = (
        PASTA_RELATORIOS
        / f"{nome_cliente} - {ano} - {nome_semestre}.pdf"
    )

    try:
        # Evita confundir um arquivo antigo com o novo download.
        if nome_arquivo_original.exists():
            nome_arquivo_original.unlink()

        driver = iniciar_chrome()
        espera = WebDriverWait(driver, 30)

        print("\nAbrindo SISAGUA no Chrome...")

        driver.get(URL_LOGIN)

        import time

        time.sleep(5)

        print("\nURL aberta:", driver.current_url)
        print("Título da página:", driver.title)
        print("Início do HTML:")
        print(driver.page_source[:1500])

        # Login.
        espera.until(
            EC.visibility_of_element_located((By.ID, "email"))
        ).send_keys(usuario)

        driver.find_element(By.ID, "senha").send_keys(senha)
        driver.find_element(By.ID, "btnEntrar").click()

        # Seleção do perfil.
        if normalizar_nome(perfil) == (
            "2.3 empresa regional municipal"
        ):
            id_perfil = "perfis:0:j_idt46"
        else:
            id_perfil = "perfis:1:j_idt46"

        espera.until(
            EC.element_to_be_clickable((By.ID, id_perfil))
        ).click()

        print(f"Perfil selecionado: {perfil}")

                # Seleção do estabelecimento pelo nome completo.
        print("Localizando estabelecimento...")

        elementos = espera.until(
            EC.presence_of_all_elements_located((
                By.XPATH,
                f"//*[normalize-space(text())="
                f"{repr(nome_estabelecimento)}]"
            ))
        )

        candidatos = [
            e for e in elementos
            if e.is_displayed()
        ]

        if len(candidatos) != 1:
            raise RuntimeError(
                "Não foi possível identificar unicamente o "
                "estabelecimento na lista do SISAGUA. "
                f"Correspondências encontradas: {len(candidatos)}. "
                "A automação foi interrompida para evitar "
                "selecionar o estabelecimento errado."
            )

        elemento = candidatos[0]

        # Realiza o clique normal pelo Selenium.
        driver.execute_script(
            "arguments[0].scrollIntoView({block: 'center'});",
            elemento
        )

        elemento.click()

        print(
            f"Estabelecimento selecionado: "
            f"{nome_estabelecimento}"
        )

        # Aguarda a conclusão do carregamento da página.
        espera.until(
            lambda d: d.execute_script(
                "return document.readyState"
            ) == "complete"
        )

        time.sleep(3)

        print(
            "URL após selecionar estabelecimento:",
            driver.current_url
        )

        print("Título:", driver.title)

        
                      # Navegação pelos menus do SISAGUA.
        from selenium.webdriver.common.action_chains import ActionChains

        acoes = ActionChains(driver)

        print("\nAcessando CONTROLE...")

        # Posiciona o mouse sobre CONTROLE.
        menu_controle = espera.until(
            EC.visibility_of_element_located((
                By.XPATH,
                "//span[contains(@class,'ui-menuitem-text') "
                "and normalize-space()='CONTROLE']"
            ))
        )
        acoes.move_to_element(menu_controle).perform()

        print("Acessando SOLUÇÃO ALTERNATIVA COLETIVA...")

        # Localiza o submenu dentro do menu CONTROLE.
        menu_solucao = espera.until(
            EC.visibility_of_element_located((
                By.XPATH,
                "//li[contains(@class,'ui-menu-parent')]"
                "[.//span[normalize-space()='CONTROLE']]"
                "//li[contains(@class,'ui-menu-parent')]"
                "[.//span[normalize-space()='SOLUÇÃO ALTERNATIVA COLETIVA']]"
                "/a"
            ))
        )

        acoes.move_to_element(menu_solucao).perform()

        print("Acessando SEMESTRAL...")

        # Clica no link SEMESTRAL dentro do submenu.
        link_semestral = espera.until(
            EC.element_to_be_clickable((
                By.CSS_SELECTOR,
                "a[href*='controleSacSemestral/controleSacSemestral.jsf']"
            ))
        )

        link_semestral.click()

        # Aguarda a abertura da tela Semestral.
        espera.until(
            EC.presence_of_element_located((
                By.ID,
                "listaSacs_data"
            ))
        )

        print("Tela Semestral aberta.")

        # Clica no lápis do SAC.
        lapis = espera.until(
            EC.element_to_be_clickable((
                By.CSS_SELECTOR,
                "#listaSacs_data span.ui-icon-pencil"
            ))
        )

        lapis.click()

        print("SAC aberto pelo ícone de lápis.")
        
        # Aguarda a tela de cadastro/visualização.
        espera.until(
            EC.presence_of_element_located(
                (By.ID, "modalVisualizarControleSemestral")
            )
        )

        # A tela abre no segundo semestre.
        # Só troca a aba quando o semestre solicitado for 1.
        if semestre == "1":
            clicar_por_texto(driver, "1° Semestre")

        # Abre a visualização do relatório.
        clicar_por_texto(driver, "VISUALIZAR")

        # Aguarda a atualização do conteúdo da janela.
        modal = espera.until(
            EC.visibility_of_element_located(
                (By.ID, "modalVisualizarControleSemestral")
            )
        )

        # Imprimir é o comando dentro da janela de visualização.
        botao_imprimir = espera.until(
            EC.element_to_be_clickable((
                By.XPATH,
                "//a[@title='Imprimir']"
            ))
        )

        botao_imprimir.click()

        print("\nComando de impressão acionado.")

        print("\nAguardando download do PDF...")

        aguardar_download(nome_arquivo_original)

        if arquivo_final.exists():
            arquivo_final.unlink()

        nome_arquivo_original.rename(arquivo_final)

        print("\nPDF salvo com sucesso:")
        print(arquivo_final)

        return True

    except Exception as e:
        import traceback

        print("\nERRO NA AUTOMAÇÃO DO PDF:")
        print(f"Tipo do erro: {type(e).__name__}")
        print(f"Descrição: {repr(e)}")
        traceback.print_exc()
        raise

    finally:
        if driver:
            driver.quit()

# ============================================================
# EXECUÇÃO PRINCIPAL - PROCESSAMENTO EM LOTE
# ============================================================

def processar_cliente(
    wb, ws, ws_rts, cliente_escolhido, selecionados,
    ano, semestre, data_preenchimento
):
    print("\n" + "=" * 70)
    print(f"INICIANDO CLIENTE: {cliente_escolhido}")
    print("=" * 70)

    nome_estabelecimento, perfil_sisagua = (
        localizar_estabelecimento_sisagua(wb, cliente_escolhido)
    )
    print(f"Estabelecimento SISAGUA: {nome_estabelecimento}")
    print(f"Perfil SISAGUA: {perfil_sisagua}")

    rts = sorted(
        {r["rt"] for r in selecionados if r["rt"]},
        key=str.casefold
    )
    if not rts:
        raise ValueError("Nenhum RT foi encontrado na linha 5 dos relatórios.")

    if len(rts) == 1:
        rt = rts[0]
        print(f"RT identificado: {rt}")
    else:
        rt = pedir_opcao(
            f"Selecione o RT responsável por {cliente_escolhido}:",
            rts
        )

    cargos = sorted({
        r["cargo"] for r in selecionados
        if normalizar_nome(r["rt"]) == normalizar_nome(rt) and r["cargo"]
    })
    if not cargos:
        raise ValueError(f"Não encontrei cargo para o RT {rt}.")

    cargo = cargos[0] if len(cargos) == 1 else pedir_opcao(
        f"Selecione o cargo do RT {rt}:",
        cargos
    )

    codigo_sac = localizar_codigo_sac(ws_rts, cliente_escolhido)
    print(f"Código da forma de abastecimento: {codigo_sac}")

    usuario, senha = obter_credenciais(wb, rt)

    cevs, pontos = consultar_sisagua(codigo_sac, ano, usuario, senha)
    codigos_pontos = sorted({p["codigo"] for p in pontos if p.get("codigo")})

    if not codigos_pontos:
        raise ValueError(
            "A consulta não retornou código de ponto de captação. "
            "Não é possível preencher automaticamente o XML."
        )

    if len(codigos_pontos) == 1:
        codigo_ponto_captacao = codigos_pontos[0]
        print(f"Ponto de captação identificado: {codigo_ponto_captacao}")
    else:
        opcoes = []
        for p in pontos:
            if not p.get("codigo"):
                continue
            descricao = f"{p['codigo']} - {p.get('nome') or 'Sem nome'}"
            if p.get("categoria"):
                descricao += f" ({p['categoria']})"
            if descricao not in opcoes:
                opcoes.append(descricao)

        escolhido = pedir_opcao(
            f"A consulta retornou mais de um ponto para {cliente_escolhido}. "
            "Selecione o ponto correspondente aos relatórios:",
            opcoes
        )
        codigo_ponto_captacao = escolhido.split(" - ")[0]

    arquivo_xml = gerar_xml(
        ws, cliente_escolhido, selecionados, codigo_sac, cevs,
        codigo_ponto_captacao, ano, semestre, data_preenchimento, rt, cargo
    )

    if arquivo_xml is None:
        raise RuntimeError("Geração do XML cancelada; envio não realizado.")

    envio_confirmado = enviar_xml(arquivo_xml, usuario, senha)
    if not envio_confirmado:
        raise RuntimeError(
            "O SISAGUA não confirmou o recebimento do XML. PDF não baixado."
        )

    baixar_relatorio_sisagua(
        cliente_escolhido, ano, semestre,
        nome_estabelecimento, perfil_sisagua, usuario, senha
    )
    return "ENVIADO E PDF BAIXADO"


def main():
    if not ARQUIVO_EXCEL.exists():
        raise FileNotFoundError(
            f"Não encontrei {ARQUIVO_EXCEL.name}. "
            "Deixe o script na mesma pasta da planilha."
        )

    wb = load_workbook(ARQUIVO_EXCEL, data_only=True)

    for aba in (ABA_LABB, ABA_RTS):
        if aba not in wb.sheetnames:
            raise ValueError(f"A aba {aba!r} não existe na planilha.")

    ws = wb[ABA_LABB]
    ws_rts = wb[ABA_RTS]

    lote = listar_relatorios(ws)
    total_clientes = len(lote)
    resultados = []

    # Mantém o período solicitado para este processamento em lote.
    ano = "2026"
    semestre = "2"
    data_preenchimento = date.today().strftime("%d/%m/%Y")

    print("\nCONFIGURAÇÃO DO LOTE")
    print(f"Ano de referência: {ano}")
    print(f"Semestre: {semestre}")
    print(f"Data de preenchimento: {data_preenchimento}")
    print(f"Clientes selecionados: {total_clientes}")

    confirmar = input(
        "\nConfirma o processamento completo (XML, envio e PDF) "
        "para os clientes selecionados? (S/N): "
    ).strip().upper()
    if confirmar != "S":
        print("Processamento cancelado pelo usuário.")
        return

    for indice, (cliente, selecionados) in enumerate(lote, 1):
        print(f"\n\n### CLIENTE {indice} DE {total_clientes} ###")
        try:
            status = processar_cliente(
                wb, ws, ws_rts, cliente, selecionados,
                ano, semestre, data_preenchimento
            )
            resultados.append((cliente, status))
            print(f"\nCONCLUÍDO: {cliente} — {status}")

        except Exception as erro:
            resultados.append((cliente, f"ERRO: {type(erro).__name__}: {erro}"))
            print(f"\nFALHA NO CLIENTE {cliente}: {erro}")
            print("O lote continuará com o próximo cliente.")

    print("\n" + "=" * 70)
    print("RESUMO FINAL DO LOTE")
    print("=" * 70)
    for cliente, status in resultados:
        print(f"- {cliente}: {status}")

    concluidos = sum(1 for _, s in resultados if s == "ENVIADO E PDF BAIXADO")
    falhas = len(resultados) - concluidos
    print(f"\nTotal: {len(resultados)} | Concluídos: {concluidos} | "
          f"Com falha: {falhas}")


if __name__ == "__main__":
    try:
        main()
    except Exception as erro:
        print(f"\nERRO FATAL: {erro}")

    input("\nPressione ENTER para fechar...")
