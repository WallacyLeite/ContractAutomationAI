# Contract Generator MVP

🇺🇸 [English](#english) | 🇧🇷 [Português](#português)

---

<a name="english"></a>
# 🇺🇸 English

> MVP for automatic generation of legal contracts in Word (.docx) from a web form, with optional data normalization via OpenAI.

## Overview

**Contract Generator MVP** automates the creation of standardized legal contracts. The user fills out a simple form in the browser, the data is sent to a Python API, optionally processed by OpenAI (formatting, capitalization, typo correction), and the final contract is generated as a `.docx` file available for immediate download.

Built for social security law firms that need to issue repetitive contracts quickly and consistently.

## Features

- Responsive web form with field validation and input masks (CPF, phone, ZIP code)
- REST API in FastAPI returning the contract as a `.docx` file
- Automatic Word template filling via `docxtpl`
- Data normalization with OpenAI GPT-4o-mini (with silent fallback if API is not configured)
- Organized file naming: `contrato_BPC_ClientName_timestamp.docx`
- Utility script to adapt existing Word templates to use `{{ }}` variables

## Stack

| Layer | Technology |
|---|---|
| Frontend | HTML5, CSS3, Vanilla JavaScript |
| Backend | Python 3.11+ · FastAPI · Uvicorn |
| Word Template | docxtpl · python-docx |
| AI (optional) | OpenAI API — GPT-4o-mini |
| Environment Variables | python-dotenv |

## Project Structure

```
Prototipo/
├── backend/
│   ├── main.py                  # Main API (routes, docxtpl, OpenAI)
│   ├── prepare_template.py      # Converts existing Word template to docxtpl format
│   ├── create_template.py       # Creates a Word template from scratch (general use)
│   ├── requirements.txt
│   ├── .env.example
│   ├── templates/
│   │   ├── CONTRATO DE Reestabelecimento de BPC.docx   # Original template
│   │   └── contrato_bpc_pronto.docx                    # Prepared template (generated)
│   └── generated/               # Generated contracts are stored here
│
└── frontend/
    ├── index.html
    ├── style.css
    └── script.js
```

## Prerequisites

- Python 3.11 or higher
- pip
- Modern browser (Chrome, Edge, Firefox)

## Installation

**1. Clone the repository**

```bash
git clone https://github.com/your-username/prototipo.git
cd prototipo
```

**2. Install Python dependencies**

```bash
cd backend
pip install -r requirements.txt
```

**3. Set up environment variables**

```bash
cp .env.example .env
```

Edit the `.env` file:

```env
OPENAI_API_KEY=sk-...  # optional — the system works without this key
```

**4. Prepare the Word template**

```bash
python prepare_template.py
```

This script reads the original template and generates `templates/contrato_bpc_pronto.docx` with docxtpl `{{ }}` variables. Run only once (or whenever the original template is changed).

## Running

**Backend**

```bash
cd backend
uvicorn main:app --reload
```

API available at `http://localhost:8000`
Interactive docs (Swagger): `http://localhost:8000/docs`

**Frontend**

Open `frontend/index.html` directly in your browser. No additional server needed.

## API Endpoints

| Method | Route | Description |
|---|---|---|
| `GET` | `/health` | Health check |
| `POST` | `/generate-contract` | Receives JSON data and returns the `.docx` |

**Sample payload for `/generate-contract`:**

```json
{
  "nome_cliente": "Maria Aparecida da Silva",
  "nacionalidade": "brasileira",
  "estado_civil": "solteira",
  "profissao": "do lar",
  "rg": "2.345.678",
  "cpf": "123.456.789-00",
  "telefone": "(62) 99999-8888",
  "endereco": "Rua das Flores, 45, Setor Norte, Goiânia – GO",
  "cep": "74000-000",
  "cidade": "Goiânia",
  "data_dia": "05",
  "data_mes": "maio",
  "data_ano": "2026"
}
```

## Template Variables

The template uses [docxtpl](https://docxtpl.readthedocs.io/) syntax. To add or edit fields, use the variables below directly in the `.docx` file:

| Variable | Description |
|---|---|
| `{{ nome_cliente }}` | Client full name |
| `{{ nacionalidade }}` | Nationality |
| `{{ estado_civil }}` | Marital status |
| `{{ profissao }}` | Profession |
| `{{ rg }}` | ID number |
| `{{ cpf }}` | CPF (Brazilian tax ID) |
| `{{ telefone }}` | Phone number with area code |
| `{{ endereco }}` | Full address |
| `{{ cep }}` | ZIP code |
| `{{ cidade }}` | City |
| `{{ data_dia }}` | Signing day |
| `{{ data_mes }}` | Signing month (written out) |
| `{{ data_ano }}` | Signing year |

## System Flow

```
HTML Form
    │
    │  POST /generate-contract (JSON)
    ▼
  FastAPI
    │
    ├── [OpenAI]  Normalizes and formats data
    │   └── Silent fallback if API is not configured
    │
    ├── [docxtpl] Loads contrato_bpc_pronto.docx
    │             Fills {{ }} variables
    │             Saves to /generated/contrato_BPC_Name_timestamp.docx
    │
    └── FileResponse (.docx)
             │
             ▼
    Automatic download in browser
```

## Notes

- Generated contracts are saved in `backend/generated/`. This folder is **not tracked by git** (`.gitignore`).
- The OpenAI key is **optional**. Without it, the system generates the contract using the form data exactly as submitted.
- This MVP has no authentication, database, or cloud storage by design — simplicity and local use are the focus.

## License

This project is licensed under the [MIT License](LICENSE).

---

<a name="português"></a>
# 🇧🇷 Português

> MVP para geração automática de contratos jurídicos em Word (.docx) a partir de um formulário web, com normalização opcional via OpenAI.

## Visão Geral

O **Contract Generator MVP** automatiza a geração de contratos jurídicos padronizados. O usuário preenche um formulário simples no navegador, os dados são enviados para uma API em Python, opcionalmente processados pela OpenAI (formatação, capitalização, correção de digitação), e o contrato final é gerado em `.docx` e disponibilizado para download imediato.

Desenvolvido para escritórios de advocacia previdenciária que precisam emitir contratos repetitivos com velocidade e padronização.

## Funcionalidades

- Formulário web responsivo com validação de campos e máscaras automáticas (CPF, telefone, CEP)
- API REST em FastAPI com retorno do contrato em `.docx`
- Preenchimento automático de template Word via `docxtpl`
- Normalização de dados com OpenAI GPT-4o-mini (com fallback automático se a API não estiver configurada)
- Nomeação organizada dos arquivos gerados: `contrato_BPC_NomeCliente_timestamp.docx`
- Script utilitário para adaptar templates Word existentes para uso com variáveis `{{ }}`

## Stack

| Camada | Tecnologia |
|---|---|
| Frontend | HTML5, CSS3, JavaScript puro |
| Backend | Python 3.11+ · FastAPI · Uvicorn |
| Template Word | docxtpl · python-docx |
| IA (opcional) | OpenAI API — GPT-4o-mini |
| Variáveis de ambiente | python-dotenv |

## Estrutura do Projeto

```
Prototipo/
├── backend/
│   ├── main.py                  # API principal (rotas, docxtpl, OpenAI)
│   ├── prepare_template.py      # Converte template Word existente para docxtpl
│   ├── create_template.py       # Cria template Word do zero (uso geral)
│   ├── requirements.txt
│   ├── .env.example
│   ├── templates/
│   │   ├── CONTRATO DE Reestabelecimento de BPC.docx   # Template original
│   │   └── contrato_bpc_pronto.docx                    # Template preparado (gerado)
│   └── generated/               # Contratos gerados ficam aqui
│
└── frontend/
    ├── index.html
    ├── style.css
    └── script.js
```

## Pré-requisitos

- Python 3.11 ou superior
- pip
- Navegador moderno (Chrome, Edge, Firefox)

## Instalação

**1. Clone o repositório**

```bash
git clone https://github.com/seu-usuario/prototipo.git
cd prototipo
```

**2. Instale as dependências Python**

```bash
cd backend
pip install -r requirements.txt
```

**3. Configure as variáveis de ambiente**

```bash
cp .env.example .env
```

Edite o arquivo `.env`:

```env
OPENAI_API_KEY=sk-...  # opcional — o sistema funciona sem esta chave
```

**4. Prepare o template Word**

```bash
python prepare_template.py
```

Esse script lê o template original e gera `templates/contrato_bpc_pronto.docx` com as variáveis `{{ }}` do docxtpl. Execute apenas uma vez (ou sempre que o template original for alterado).

## Como Rodar

**Backend**

```bash
cd backend
uvicorn main:app --reload
```

API disponível em `http://localhost:8000`
Documentação interativa (Swagger): `http://localhost:8000/docs`

**Frontend**

Abra o arquivo `frontend/index.html` diretamente no navegador. Nenhum servidor adicional é necessário.

## Endpoints da API

| Método | Rota | Descrição |
|---|---|---|
| `GET` | `/health` | Verifica se a API está no ar |
| `POST` | `/generate-contract` | Recebe dados JSON e retorna o `.docx` |

**Exemplo de payload para `/generate-contract`:**

```json
{
  "nome_cliente": "Maria Aparecida da Silva",
  "nacionalidade": "brasileira",
  "estado_civil": "solteira",
  "profissao": "do lar",
  "rg": "2.345.678",
  "cpf": "123.456.789-00",
  "telefone": "(62) 99999-8888",
  "endereco": "Rua das Flores, 45, Setor Norte, Goiânia – GO",
  "cep": "74000-000",
  "cidade": "Goiânia",
  "data_dia": "05",
  "data_mes": "maio",
  "data_ano": "2026"
}
```

## Variáveis do Template Word

O template utiliza a sintaxe do [docxtpl](https://docxtpl.readthedocs.io/). Para adicionar ou editar campos, use as variáveis abaixo diretamente no arquivo `.docx`:

| Variável | Descrição |
|---|---|
| `{{ nome_cliente }}` | Nome completo do contratante |
| `{{ nacionalidade }}` | Nacionalidade |
| `{{ estado_civil }}` | Estado civil |
| `{{ profissao }}` | Profissão |
| `{{ rg }}` | Número do RG |
| `{{ cpf }}` | CPF formatado |
| `{{ telefone }}` | Telefone com DDD |
| `{{ endereco }}` | Endereço completo |
| `{{ cep }}` | CEP |
| `{{ cidade }}` | Cidade |
| `{{ data_dia }}` | Dia da assinatura |
| `{{ data_mes }}` | Mês da assinatura (por extenso) |
| `{{ data_ano }}` | Ano da assinatura |

## Fluxo do Sistema

```
Formulário HTML
     │
     │  POST /generate-contract (JSON)
     ▼
  FastAPI
     │
     ├── [OpenAI]  Normaliza e formata os dados
     │   └── Fallback silencioso se a API não estiver configurada
     │
     ├── [docxtpl] Carrega contrato_bpc_pronto.docx
     │             Preenche as variáveis {{ }}
     │             Salva em /generated/contrato_BPC_Nome_timestamp.docx
     │
     └── FileResponse (.docx)
              │
              ▼
     Download automático no navegador
```

## Observações

- Os contratos gerados ficam salvos em `backend/generated/`. Essa pasta **não é versionada** no git (`.gitignore`).
- A chave da OpenAI é **opcional**. Sem ela, o sistema gera o contrato normalmente usando os dados exatamente como preenchidos no formulário.
- Este MVP não possui autenticação, banco de dados ou integração com armazenamento em nuvem por design — o foco é simplicidade e funcionamento local.

## Licença

Este projeto está sob a licença [MIT](LICENSE).
