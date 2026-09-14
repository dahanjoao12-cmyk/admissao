from pydantic import BaseModel

# Ordem também usada pelo prompt de extração (IA) e pela tela de validação.
CAMPOS_ADMISSAO = [
    "nome_completo",
    "cpf",
    "rg",
    "rg_orgao_emissor",
    "rg_uf_emissor",
    "rg_data_emissao",
    "data_nascimento",
    "sexo",
    "naturalidade_cidade",
    "naturalidade_uf",
    "nome_mae",
    "nome_pai",
    "endereco_logradouro",
    "endereco_numero",
    "endereco_complemento",
    "endereco_bairro",
    "endereco_cidade",
    "endereco_uf",
    "endereco_cep",
]


class DadosAdmissao(BaseModel):
    nome_completo: str = ""
    cpf: str = ""
    rg: str = ""
    rg_orgao_emissor: str = ""
    rg_uf_emissor: str = ""
    rg_data_emissao: str = ""
    data_nascimento: str = ""
    sexo: str = ""
    naturalidade_cidade: str = ""
    naturalidade_uf: str = ""
    nome_mae: str = ""
    nome_pai: str = ""
    endereco_logradouro: str = ""
    endereco_numero: str = ""
    endereco_complemento: str = ""
    endereco_bairro: str = ""
    endereco_cidade: str = ""
    endereco_uf: str = ""
    endereco_cep: str = ""


class GerarArquivosResponse(BaseModel):
    job_id: str
    docx_url: str
    dominio_url: str
