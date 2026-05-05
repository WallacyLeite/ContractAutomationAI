"""
MVP Prototipo — Gerador de Contratos Automatizado
Contrato: Reestabelecimento de BPC
Backend: FastAPI + docxtpl + OpenAI
"""

import os
import re
import json
from pathlib import Path
from datetime import datetime

from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from pydantic import BaseModel
from dotenv import load_dotenv
from docxtpl import DocxTemplate
import openai

# Carrega variáveis de ambiente do arquivo .env
load_dotenv()

app = FastAPI(title="Gerador de Contratos BPC", version="1.0.0")

# Libera CORS para o frontend local funcionar sem bloqueio do navegador
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["GET", "POST"],
    allow_headers=["*"],
)

# Caminhos importantes do projeto
BASE_DIR      = Path(__file__).parent
TEMPLATE_PATH = BASE_DIR / "templates" / "contrato_bpc_pronto.docx"
GENERATED_DIR = BASE_DIR / "generated"
GENERATED_DIR.mkdir(exist_ok=True)


# ---------------------------------------------------------------------------
# Modelo de dados — campos do Contrato de Reestabelecimento de BPC
# ---------------------------------------------------------------------------
class ContratoData(BaseModel):
    nome_cliente: str
    nacionalidade: str
    estado_civil: str
    profissao: str
    rg: str
    cpf: str
    telefone: str
    endereco: str
    cep: str
    cidade: str
    data_dia: str   # ex: "05"
    data_mes: str   # ex: "maio"
    data_ano: str   # ex: "2026"


# ---------------------------------------------------------------------------
# ETAPA IA — Normalização dos dados com OpenAI
#
# Corrige formatação, capitalização e padrões antes de preencher o template.
# Se a API falhar, o sistema continua com os dados originais (fallback seguro).
# ---------------------------------------------------------------------------
def processar_com_openai(dados: dict) -> dict:
    api_key = os.getenv("OPENAI_API_KEY", "")

    if not api_key or api_key == "your_openai_api_key_here":
        print("[OpenAI] Chave não configurada. Usando dados originais.")
        return dados

    try:
        client = openai.OpenAI(api_key=api_key)

        prompt = (
            "Você é um assistente jurídico especializado em contratos previdenciários brasileiros.\n"
            "Normalize os dados abaixo para preencher um Contrato de Reestabelecimento de BPC (LOAS).\n"
            "Regras:\n"
            "- Capitalize nomes, profissões e endereços corretamente\n"
            "- Formate CPF no padrão: 000.000.000-00\n"
            "- Formate RG com pontos e traço quando aplicável\n"
            "- Formate CEP como: 00000-000\n"
            "- Formate telefone como: (00) 00000-0000\n"
            "- Estado civil em minúsculas: solteiro, casado, divorciado, viúvo, separado\n"
            "- Mês por extenso em português (ex: 'maio', 'outubro')\n"
            "- Retorne APENAS um JSON com as mesmas chaves, sem texto adicional\n\n"
            f"Dados:\n{json.dumps(dados, ensure_ascii=False, indent=2)}"
        )

        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[{"role": "user", "content": prompt}],
            response_format={"type": "json_object"},
            temperature=0.1,
        )

        dados_normalizados = json.loads(response.choices[0].message.content)
        print("[OpenAI] Dados normalizados com sucesso.")
        return dados_normalizados

    except Exception as e:
        # Falha silenciosa — nunca interrompe a geração do contrato
        print(f"[OpenAI] Erro: {e}. Continuando com dados originais.")
        return dados


# ---------------------------------------------------------------------------
# ROTA: GET /health — Verifica se a API está no ar
# ---------------------------------------------------------------------------
@app.get("/health")
def health_check():
    return {"status": "ok", "message": "API funcionando corretamente"}


# ---------------------------------------------------------------------------
# ROTA: POST /generate-contract — Gera e retorna o contrato em .docx
#
# Fluxo:
# 1. Recebe JSON com os dados do formulário
# 2. Valida se o template Word preparado existe
# 3. Processa dados com OpenAI (com fallback automático)
# 4. Preenche o template com docxtpl
# 5. Salva o arquivo em /generated com nome organizado
# 6. Retorna o .docx para download
# ---------------------------------------------------------------------------
@app.post("/generate-contract")
def generate_contract(dados: ContratoData):

    # Verifica se o template preparado existe
    if not TEMPLATE_PATH.exists():
        raise HTTPException(
            status_code=500,
            detail=(
                "Template não encontrado em backend/templates/contrato_bpc_pronto.docx. "
                "Execute primeiro: python prepare_template.py"
            ),
        )

    # Converte o modelo Pydantic para dicionário Python
    dados_dict = dados.model_dump()

    # --- Etapa IA: normaliza dados (com fallback automático se falhar) ---
    dados_finais = processar_com_openai(dados_dict)

    # Carrega o template Word com docxtpl
    doc = DocxTemplate(str(TEMPLATE_PATH))

    # Preenche todas as variáveis {{ campo }} do template
    doc.render(dados_finais)

    # Gera nome de arquivo seguro: contrato_NomeCliente_YYYYMMDD_HHMMSS.docx
    nome_raw  = dados_finais.get("nome_cliente", "cliente")
    nome_safe = re.sub(r"[^\w\s-]", "", nome_raw).strip().replace(" ", "_")
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    nome_arquivo = f"contrato_BPC_{nome_safe}_{timestamp}.docx"

    # Salva na pasta /generated
    caminho_saida = GENERATED_DIR / nome_arquivo
    doc.save(str(caminho_saida))

    print(f"[Contrato] Gerado: {nome_arquivo}")

    # Retorna o .docx como resposta para download automático
    return FileResponse(
        path=str(caminho_saida),
        filename=nome_arquivo,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
