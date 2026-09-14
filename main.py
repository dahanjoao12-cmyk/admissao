import asyncio
import json
import os
import re
import shutil
import unicodedata
import uuid
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import List
from zoneinfo import ZoneInfo

from dotenv import load_dotenv
from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from docx import Document

from models import CAMPOS_ADMISSAO, DadosAdmissao, GerarArquivosResponse

load_dotenv()

BASE_DIR = Path(__file__).resolve().parent
GENERATED_DIR = BASE_DIR / "generated"
GENERATED_DIR.mkdir(exist_ok=True)
TEMPLATE_PATH = BASE_DIR / "template_contrato.docx"

BRASILIA_TZ = ZoneInfo("America/Sao_Paulo")

# Estado em memória dos jobs de geração (sem banco de dados, conforme escopo do projeto).
# Guarda apenas o nome-base usado nos arquivos baixáveis. Os arquivos em si ficam no disco
# até a limpeza diária (ver _limpeza_diaria_loop) apagar tudo de uma vez, de madrugada.
JOBS_META: dict[str, dict] = {}


def _slugify_nome(nome: str) -> str:
    sem_acentos = unicodedata.normalize("NFKD", nome).encode("ascii", "ignore").decode("ascii")
    slug = re.sub(r"[^a-zA-Z0-9]+", "", sem_acentos).lower()
    return slug or "colaborador"


def _limpar_todos_os_jobs() -> None:
    for job_dir in GENERATED_DIR.iterdir():
        if job_dir.is_dir():
            shutil.rmtree(job_dir, ignore_errors=True)
    JOBS_META.clear()


async def _limpeza_diaria_loop() -> None:
    """Apaga todos os arquivos gerados uma vez por dia, às 3h no horário de Brasília.

    Não apagamos na hora do download: os arquivos ficam disponíveis no servidor
    durante o dia (para o RH baixar/reenviar quando precisar) e são varridos de
    uma vez só de madrugada, quando ninguém deve estar usando o sistema.
    """
    while True:
        agora = datetime.now(BRASILIA_TZ)
        proxima_execucao = agora.replace(hour=3, minute=0, second=0, microsecond=0)
        if proxima_execucao <= agora:
            proxima_execucao += timedelta(days=1)
        await asyncio.sleep((proxima_execucao - agora).total_seconds())
        _limpar_todos_os_jobs()


ANTHROPIC_MODEL = os.environ.get("ANTHROPIC_MODEL", "claude-sonnet-5")

app = FastAPI(title="Admissão RH — API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.on_event("startup")
async def _iniciar_limpeza_diaria() -> None:
    asyncio.create_task(_limpeza_diaria_loop())

# --------------------------------------------------------------------------
# Extração de dados via IA (Anthropic)
# --------------------------------------------------------------------------

SYSTEM_PROMPT = """Você é um assistente especialista em extrair dados cadastrais de documentos brasileiros a partir de imagens ou PDFs, para fins de admissão de funcionários.

Você pode receber mais de um documento na mesma solicitação — por exemplo, um documento de identificação (RG, CPF ou CNH) e, separadamente, um comprovante de residência (conta de luz, água, telefone ou internet). Combine as informações de TODOS os documentos enviados num único resultado: dados pessoais normalmente vêm do documento de identificação, e o endereço normalmente vem do comprovante de residência (mas também pode vir do próprio documento de identificação, quando presente).

Analise cuidadosamente os documentos fornecidos e devolva a resposta APENAS como um objeto JSON válido — sem markdown, sem texto antes ou depois, sem comentários — contendo exatamente estas chaves:

{
  "nome_completo": "",
  "cpf": "",
  "rg": "",
  "rg_orgao_emissor": "",
  "rg_uf_emissor": "",
  "rg_data_emissao": "",
  "data_nascimento": "",
  "sexo": "",
  "naturalidade_cidade": "",
  "naturalidade_uf": "",
  "nome_mae": "",
  "nome_pai": "",
  "endereco_logradouro": "",
  "endereco_numero": "",
  "endereco_complemento": "",
  "endereco_bairro": "",
  "endereco_cidade": "",
  "endereco_uf": "",
  "endereco_cep": ""
}

Regras:
- Nunca invente ou deduza informação que não esteja visível nos documentos. Se um campo não for encontrado em nenhum documento, use string vazia "".
- Formate o CPF como 000.000.000-00 quando encontrado.
- Formate datas como DD/MM/AAAA.
- "sexo" deve ser "Masculino" ou "Feminino" quando identificável.
- "rg_orgao_emissor" é a sigla do órgão emissor (ex: "SSP"), separada da UF, que vai em "rg_uf_emissor" (ex: "RJ").
- "naturalidade_cidade" e "naturalidade_uf" vêm separados (ex: cidade "Rio de Janeiro", uf "RJ").
- Os campos "endereco_*" vêm cada um separado (não junte tudo num só campo). "endereco_cep" apenas números.
- Responda SOMENTE com o objeto JSON."""

IMAGE_MEDIA_TYPES = {"image/png", "image/jpeg", "image/jpg", "image/webp", "image/gif"}


def _build_content_block(content_type: str, data_b64: str) -> dict:
    if content_type == "application/pdf":
        return {
            "type": "document",
            "source": {"type": "base64", "media_type": "application/pdf", "data": data_b64},
        }
    media_type = content_type if content_type in IMAGE_MEDIA_TYPES else "image/jpeg"
    return {
        "type": "image",
        "source": {"type": "base64", "media_type": media_type, "data": data_b64},
    }


def _extract_json(text: str) -> dict:
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass
    match = re.search(r"\{.*\}", text, re.DOTALL)
    if match:
        return json.loads(match.group(0))
    raise ValueError("A IA não retornou um JSON válido.")


@app.post("/api/extrair")
async def extrair_dados(documentos: List[UploadFile] = File(...)):
    if not documentos:
        raise HTTPException(status_code=400, detail="Nenhum documento enviado.")

    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        raise HTTPException(
            status_code=503,
            detail=(
                "ANTHROPIC_API_KEY não configurada no servidor. Defina a variável "
                "de ambiente (ou arquivo .env) para habilitar a extração via IA."
            ),
        )

    import base64

    import anthropic

    content_blocks = []
    for upload in documentos:
        raw = await upload.read()
        if not raw:
            continue
        b64 = base64.b64encode(raw).decode("utf-8")
        content_blocks.append(_build_content_block(upload.content_type or "", b64))

    if not content_blocks:
        raise HTTPException(status_code=400, detail="Os arquivos enviados estão vazios.")

    content_blocks.append(
        {
            "type": "text",
            "text": "Extraia os dados cadastrais dos documentos de identificação acima.",
        }
    )

    client = anthropic.Anthropic(api_key=api_key)

    try:
        response = client.messages.create(
            model=ANTHROPIC_MODEL,
            max_tokens=1024,
            system=SYSTEM_PROMPT,
            messages=[{"role": "user", "content": content_blocks}],
        )
    except anthropic.APIError as exc:
        raise HTTPException(status_code=502, detail=f"Erro ao chamar a API da Anthropic: {exc}")

    texto_resposta = "".join(
        block.text for block in response.content if getattr(block, "type", None) == "text"
    )

    try:
        dados_brutos = _extract_json(texto_resposta)
    except (ValueError, json.JSONDecodeError):
        raise HTTPException(
            status_code=502,
            detail="A IA retornou uma resposta que não pôde ser interpretada como JSON.",
        )

    dados = DadosAdmissao(**{k: dados_brutos.get(k, "") or "" for k in CAMPOS_ADMISSAO})
    return dados.model_dump()


# --------------------------------------------------------------------------
# Geração de arquivos (contrato .docx + layout de importação do Domínio)
# --------------------------------------------------------------------------


def _replace_placeholders_in_paragraph(paragraph, contexto: dict) -> None:
    if not paragraph.runs:
        return
    texto_completo = "".join(run.text for run in paragraph.runs)
    if "{{" not in texto_completo:
        return
    novo_texto = texto_completo
    for chave, valor in contexto.items():
        novo_texto = novo_texto.replace("{{" + chave + "}}", str(valor))
    if novo_texto == texto_completo:
        return
    paragraph.runs[0].text = novo_texto
    for run in paragraph.runs[1:]:
        run.text = ""


def _juntar(*partes: str, sep: str = " ") -> str:
    return sep.join(p for p in partes if p)


def _montar_campos_combinados(dados: DadosAdmissao) -> dict:
    """Strings de exibição (para o contrato), combinando os campos estruturados."""
    endereco_linha1 = _juntar(
        dados.endereco_logradouro,
        f", {dados.endereco_numero}" if dados.endereco_numero else "",
        sep="",
    )
    endereco_completo = " - ".join(
        p
        for p in [
            endereco_linha1,
            dados.endereco_complemento,
            dados.endereco_bairro,
            _juntar(dados.endereco_cidade, dados.endereco_uf, sep="/"),
            f"CEP {dados.endereco_cep}" if dados.endereco_cep else "",
        ]
        if p
    )
    return {
        "endereco_completo": endereco_completo,
        "naturalidade": _juntar(dados.naturalidade_cidade, dados.naturalidade_uf, sep="/"),
        "rg_orgao_uf": _juntar(dados.rg_orgao_emissor, dados.rg_uf_emissor, sep="/"),
    }


def gerar_contrato_docx(dados: DadosAdmissao, destino: Path) -> None:
    if not TEMPLATE_PATH.exists():
        raise HTTPException(
            status_code=500,
            detail=(
                f"Template não encontrado em {TEMPLATE_PATH.name}. Execute "
                "scripts/create_template.py ou adicione o modelo real da empresa."
            ),
        )
    doc = Document(str(TEMPLATE_PATH))
    contexto = {**dados.model_dump(), **_montar_campos_combinados(dados)}

    for paragraph in doc.paragraphs:
        _replace_placeholders_in_paragraph(paragraph, contexto)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    _replace_placeholders_in_paragraph(paragraph, contexto)

    doc.save(str(destino))


with open(BASE_DIR / "data" / "municipios_dominio.json", encoding="utf-8") as _f:
    _MUNICIPIOS_DOMINIO: dict[str, int] = json.load(_f)

PAIS_BRASIL_CODIGO = "30"  # Código do Brasil na aba "Código dos países" do Domínio.


def _codigo_municipio(nome_cidade: str) -> str:
    """Busca exata (sem acento/caixa) na tabela oficial de municípios do Domínio.

    Só preenche quando encontra correspondência exata — na dúvida, deixa em
    branco para o RH completar manualmente, em vez de arriscar um código errado.
    """
    if not nome_cidade:
        return ""
    chave = unicodedata.normalize("NFKD", nome_cidade).encode("ascii", "ignore").decode("ascii").strip().upper()
    codigo = _MUNICIPIOS_DOMINIO.get(chave)
    return str(codigo) if codigo is not None else ""


def _somente_digitos(valor: str) -> str:
    return re.sub(r"\D", "", valor or "")


def _mapear_sexo_dominio(sexo: str) -> str:
    s = (sexo or "").strip().lower()
    if s.startswith("m"):
        return "M"
    if s.startswith("f"):
        return "F"
    return ""


# Posição (1-indexado, conforme o "Dicionário dos campos" do Domínio) -> valor.
# Campos de folha de pagamento (cargo, salário, departamento, banco, horas,
# categoria, sindicato etc.) ficam em branco: não vêm de nenhum documento de
# identificação, e por decisão do escritório são preenchidos manualmente pelo
# RH direto no Domínio Web, não por este sistema.
def _montar_linha_dominio(dados: DadosAdmissao) -> list[str]:
    campos: dict[int, str] = {
        1: "0",  # Código empresa — instrução do próprio Domínio: sempre 0.
        3: dados.nome_completo,
        4: _somente_digitos(dados.cpf),
        17: dados.endereco_logradouro,
        18: dados.endereco_numero,
        19: dados.endereco_complemento,
        20: dados.endereco_bairro,
        21: dados.endereco_cidade,
        22: dados.endereco_uf,
        23: _somente_digitos(dados.endereco_cep),
        24: dados.data_nascimento,
        43: _mapear_sexo_dominio(dados.sexo),
        44: PAIS_BRASIL_CODIGO,  # País nacionalidade — assumimos Brasil por padrão.
        49: dados.nome_pai,
        50: dados.nome_mae,
        52: _codigo_municipio(dados.endereco_cidade),
        53: _codigo_municipio(dados.naturalidade_cidade),
        54: PAIS_BRASIL_CODIGO,  # País do endereço.
        55: PAIS_BRASIL_CODIGO,  # País de nascimento.
        56: dados.rg,
        57: dados.rg_orgao_emissor,
        58: dados.rg_uf_emissor,
        59: dados.rg_data_emissao,
        76: dados.naturalidade_uf,
    }
    return [campos.get(pos, "") for pos in range(1, 79)]


def gerar_layout_dominio(dados: DadosAdmissao, destino: Path) -> None:
    """Gera o "empregados.txt" no layout oficial de importação de admissão do
    Domínio Web: 78 colunas separadas por TAB, sem cabeçalho, uma linha por
    empregado (mesmo formato que a macro da planilha modelo do escritório gera).

    Só preenchemos os campos que vêm de documento (identidade e endereço) — os
    campos de folha de pagamento ficam em branco, para o RH completar dentro do
    próprio Domínio Web.
    """
    linha = "\t".join(_montar_linha_dominio(dados))
    with destino.open("w", encoding="cp1252", errors="replace", newline="\r\n") as f:
        f.write(linha + "\n")


@app.post("/api/gerar", response_model=GerarArquivosResponse)
async def gerar_arquivos(dados: DadosAdmissao):
    if not dados.nome_completo.strip():
        raise HTTPException(status_code=400, detail="nome_completo é obrigatório.")

    job_id = uuid.uuid4().hex
    job_dir = GENERATED_DIR / job_id
    job_dir.mkdir(parents=True, exist_ok=True)

    docx_path = job_dir / "contrato_trabalho.docx"
    dominio_path = job_dir / "layout_dominio.txt"

    gerar_contrato_docx(dados, docx_path)
    gerar_layout_dominio(dados, dominio_path)

    # Nome-base usado nos arquivos baixados: {nomedocolaborador}_{ddmmyyyy} (data de
    # geração dos documentos — a mesma convenção que será usada na pasta do cloud).
    base_nome = f"{_slugify_nome(dados.nome_completo)}_{date.today().strftime('%d%m%Y')}"

    JOBS_META[job_id] = {"base_nome": base_nome}

    return GerarArquivosResponse(
        job_id=job_id,
        docx_url=f"/api/download/{job_id}/docx",
        dominio_url=f"/api/download/{job_id}/dominio",
    )


@app.get("/api/download/{job_id}/{tipo}")
async def download_arquivo(job_id: str, tipo: str):
    if not re.fullmatch(r"[a-f0-9]{32}", job_id):
        raise HTTPException(status_code=400, detail="job_id inválido.")

    meta = JOBS_META.get(job_id)
    if not meta:
        raise HTTPException(status_code=404, detail="Arquivo não encontrado (job expirado ou já apagado na limpeza diária).")

    job_dir = GENERATED_DIR / job_id
    if tipo == "docx":
        caminho = job_dir / "contrato_trabalho.docx"
        media_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        nome_arquivo = f"{meta['base_nome']}_contrato.docx"
    elif tipo == "dominio":
        caminho = job_dir / "layout_dominio.txt"
        media_type = "text/plain"
        nome_arquivo = f"{meta['base_nome']}_dominio.txt"
    else:
        raise HTTPException(status_code=404, detail="Tipo de arquivo desconhecido.")

    if not caminho.exists():
        raise HTTPException(status_code=404, detail="Arquivo não encontrado (job expirado ou inválido).")

    return FileResponse(caminho, media_type=media_type, filename=nome_arquivo)


# --------------------------------------------------------------------------
# Frontend estático
# --------------------------------------------------------------------------

app.mount("/", StaticFiles(directory=str(BASE_DIR / "static"), html=True), name="static")
