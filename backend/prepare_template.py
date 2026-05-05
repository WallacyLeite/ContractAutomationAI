"""
Script utilitário — Adapta o template BPC para uso com docxtpl.

Substitui os campos em branco (____) e tabs pelos marcadores {{ variavel }},
preservando a formatação original do documento Word.

Execute uma vez antes de subir o backend:
    python prepare_template.py

Entrada:  templates/CONTRATO DE Reestabelecimento de BPC.docx
Saída:    templates/contrato_bpc_pronto.docx
"""

import re
from docx import Document

TEMPLATE_ENTRADA = "templates/CONTRATO DE Reestabelecimento de BPC.docx"
TEMPLATE_SAIDA   = "templates/contrato_bpc_pronto.docx"


# ---------------------------------------------------------------------------
# Helpers de leitura e escrita em parágrafos do python-docx
# ---------------------------------------------------------------------------

def ler_texto(paragrafo):
    """Retorna o texto completo do parágrafo concatenando todos os runs."""
    return "".join(r.text for r in paragrafo.runs)


def gravar_texto(paragrafo, novo_texto):
    """
    Coloca o novo texto no primeiro run e limpa os demais.
    Isso preserva a formatação do parágrafo (fonte, tamanho, negrito do run[0]).
    """
    if not paragrafo.runs:
        return
    paragrafo.runs[0].text = novo_texto
    for run in paragrafo.runs[1:]:
        run.text = ""


# ---------------------------------------------------------------------------
# Mapeamento de substituições — um bloco por campo do template
# ---------------------------------------------------------------------------

def adaptar(texto):
    """
    Aplica todas as substituições de variáveis ao texto de um parágrafo.
    Retorna o texto modificado (ou o original se nada mudou).
    """
    t = texto

    # --- DATAS com tabs: "Goiânia, [TAB]de[TAB]de 20[TAB]." ---
    # Cobre variações como "Goiânia, ___de___ de 20_____."
    if re.search(r'Goi.nia', t, re.IGNORECASE):
        t = re.sub(
            r'(Goi.nia,?\s*)(?:\t+|_{2,})\s*de\s*(?:\t+|_{2,})\s*de\s*20\s*(?:\t+|_{2,})\s*\.?',
            r'\1{{ data_dia }} de {{ data_mes }} de {{ data_ano }}.',
            t, flags=re.IGNORECASE
        )

    # --- DATAS com barras: "[TAB],[TAB]/[TAB]/20[TAB]." ---
    # Linha da renúncia e autorização: ",  /  /20  ."
    # O "20" já está no padrão e é consumido — data_ano recebe o ano completo (ex: 2026)
    if re.search(r'/\s*(?:\t+|_{2,})?\s*/\s*20', t):
        t = re.sub(
            r'[\s,]*(?:\t+|_{0,})\s*/\s*(?:\t+|_{0,})\s*/\s*20\s*(?:\t+|_{0,})\.?',
            '{{ cidade }}, {{ data_dia }}/{{ data_mes }}/{{ data_ano }}.',
            t
        )

    # --- OUTORGANTE: ____ ---
    t = re.sub(r'(OUTORGANTE:\s*)_{2,}', r'\1{{ nome_cliente }}', t)

    # --- Nome: ____ (aparece em múltiplos blocos) ---
    t = re.sub(r'(Nome:\s*)_{2,}', r'\1{{ nome_cliente }}', t)

    # --- CPF: ____ (standalone) ---
    t = re.sub(r'(CPF:\s*)_{2,}', r'\1{{ cpf }}', t)

    # --- CONTRATANTE: ____ ---
    t = re.sub(r'(CONTRATANTE:\s*)_{2,}', r'\1{{ nome_cliente }}', t)

    # --- Nacionalidade: ____ ---
    t = re.sub(r'(Nacionalidade:\s*)_{2,}', r'\1{{ nacionalidade }}', t)

    # --- Estado civil: ____ e profissão: ____ na mesma linha ---
    t = re.sub(r'(Estado civil:\s*)_{2,}', r'\1{{ estado_civil }}', t)
    t = re.sub(r'(profiss[aã]o:\s*)_{2,}', r'\1{{ profissao }}', t)

    # --- RG ____ CPF: ____ na mesma linha ---
    t = re.sub(r'(RG\s*)_{2,}', r'\1{{ rg }}', t)

    # --- telefone: ____ ---
    t = re.sub(r'(telefone:\s*)_{2,}', r'\1{{ telefone }}', t, flags=re.IGNORECASE)

    # --- (a): ____ (linha de endereço) ---
    t = re.sub(r'(\(a\):\s*)_{2,}', r'\1{{ endereco }}', t)

    # --- "fazem: ____ e Eduardo" (cabeçalho do contrato) ---
    t = re.sub(r'(fazem:\s*)_{2,}', r'\1{{ nome_cliente }}', t)

    # --- Assinatura: ____ ---
    t = re.sub(r'(Assinatura:\s*)_{2,}', r'\1{{ nome_cliente }}', t)

    # --- Declaração de Renúncia: "Eu, ____, portador do CPF:____, residente..." ---
    if t.strip().startswith('Eu,') and 'portador do CPF' in t:
        t = re.sub(r'(Eu,\s*(?:\t+)?)_{2,}', r'\1{{ nome_cliente }}', t)
        t = re.sub(r'(portador do CPF:\s*)_{2,}', r'\1{{ cpf }}', t)
        t = re.sub(r'(residente e domiciliado na\s*(?:\t+)?)_{0,}', r'\1{{ endereco }}', t)

    # --- Linha de CEP e Cidade da renúncia ---
    if re.match(r'^\s*CEP:', t):
        t = re.sub(r'(CEP:\s*)(?:_{2,}|\t+)', r'\1{{ cep }}', t)
        t = re.sub(r'(Cidade\s*-\s*)(?:_{2,}|\t+)', r'\1{{ cidade }}', t)

    # --- Autorização: "Eu, [TAB], brasileiro(a), estado civil: [TAB], CPF n° [TAB], RG n° [TAB]" ---
    # Nesta seção os campos em branco são tabs, não underscores
    if t.strip().startswith('Eu,') and re.search(r'brasileiro', t, re.IGNORECASE):
        t = re.sub(r'(Eu,\s*)(?:\t+|_{2,})', r'\1{{ nome_cliente }}', t)
        t = re.sub(r'(estado\s+civil:\s*)(?:\t+|_{2,})', r'\1{{ estado_civil }}', t, flags=re.IGNORECASE)
        t = re.sub(r'(CPF\s*n[°oº\.]\s*)(?:\t+|_{2,})', r'\1{{ cpf }}', t)
        t = re.sub(r'(RG\s*n[°oº\.]\s*)(?:\t+|_{2,})', r'\1{{ rg }}', t)

    return t


# ---------------------------------------------------------------------------
# Processamento principal
# ---------------------------------------------------------------------------

def preparar():
    doc = Document(TEMPLATE_ENTRADA)
    alteracoes = 0

    for i, para in enumerate(doc.paragraphs):
        original = ler_texto(para)
        if not original.strip():
            continue

        modificado = adaptar(original)

        if modificado != original:
            gravar_texto(para, modificado)
            alteracoes += 1
            # Mostra as primeiras 100 chars para conferência
            print(f"  [§{i:03d}] {modificado[:100]}")

    doc.save(TEMPLATE_SAIDA)

    print(f"\n{alteracoes} parágrafos adaptados.")
    print(f"Template salvo em: {TEMPLATE_SAIDA}\n")
    print("Variáveis disponíveis no template:")
    for v in [
        "{{ nome_cliente }}", "{{ nacionalidade }}", "{{ estado_civil }}",
        "{{ profissao }}", "{{ rg }}", "{{ cpf }}", "{{ telefone }}",
        "{{ endereco }}", "{{ cep }}", "{{ cidade }}",
        "{{ data_dia }}", "{{ data_mes }}", "{{ data_ano }}",
    ]:
        print(f"  {v}")


if __name__ == "__main__":
    preparar()
