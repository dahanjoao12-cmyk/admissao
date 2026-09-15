from pydantic import BaseModel

# Campos preenchidos pela IA a partir dos documentos. Usado para filtrar a
# resposta do prompt de extração — os campos de folha de pagamento (abaixo,
# em DadosAdmissao) são preenchidos manualmente pelo RH, não pela IA.
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
    # --- Documentais (extraídos pela IA) ---
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

    # --- Folha de pagamento (preenchidos manualmente pelo RH) ---
    codigo_empregado: str = ""
    codigo_filial: str = ""
    codigo_servico: str = ""
    codigo_cargo: str = ""
    codigo_departamento: str = ""
    codigo_centro_custos: str = ""
    codigo_convenio_coletivo: str = ""
    codigo_sindicato_cadastro: str = ""

    salario: str = ""
    data_admissao: str = ""
    codigo_esocial_empregado: str = ""
    vinculo_empregaticio: str = ""
    categoria: str = ""
    tipo_indicativo_admissao: str = ""
    tipo_admissao: str = ""
    codigo_categoria_esocial: str = ""
    atividade_simples_nacional: str = ""

    horas_dia: str = ""
    horas_semana: str = ""
    horas_mes: str = ""
    codigo_jornada: str = ""
    forma_calculo_pagamento: str = ""
    forma_pagamento: str = ""
    vencimento_ferias: str = ""

    optante_fgts: str = ""
    data_opcao_fgts: str = ""
    ocorrencia_sefip: str = ""
    multiplos_vinculos: str = ""
    pagou_contribuicao_sindical: str = ""

    codigo_banco: str = ""
    conta_corrente: str = ""
    digito_conta: str = ""
    tipo_conta: str = ""

    contrato_experiencia: str = ""
    clausula_assecuratoria: str = ""
    dias_contrato_experiencia: str = ""
    inicio_prazo_determinado: str = ""
    fim_prazo_determinado: str = ""
    dias_prorrogacao: str = ""
    prorrogacao_prazo_determinado: str = ""
    motivo_prorrogacao: str = ""

    grau_instrucao: str = ""
    deficiente_fisico: str = ""
    situacao_qualificacao_cadastral: str = ""
    raca_cor: str = ""
    estado_civil: str = ""
    recebia_seguro_desemprego: str = ""
    tipo_endereco: str = ""

    email: str = ""
    ddd_telefone: str = ""
    telefone: str = ""
    ddd_contato_2: str = ""
    contato_2: str = ""


class GerarArquivosResponse(BaseModel):
    job_id: str
    docx_url: str
    dominio_url: str
    conferencia_url: str
