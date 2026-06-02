/**
 * MVP Prototipo — Contrato BPC
 * Frontend: coleta dados do formulário, envia para o backend e dispara download do .docx
 */

// URL base do backend
// Em arquivo local, cai no backend de desenvolvimento.
// Em produção, usa a mesma origem do site publicado.
const API_BASE = window.location.origin === "null"
  ? "http://localhost:8000"
  : window.location.origin;

// Nomes dos meses em português, usados para converter a data do input
const MESES_PT = [
  "janeiro", "fevereiro", "março", "abril", "maio", "junho",
  "julho", "agosto", "setembro", "outubro", "novembro", "dezembro",
];

// Referências aos elementos do DOM
const form       = document.getElementById("contratoForm");
const submitBtn  = document.getElementById("submitBtn");
const btnText    = document.getElementById("btnText");
const btnLoading = document.getElementById("btnLoading");
const statusMsg  = document.getElementById("statusMsg");


// ---------------------------------------------------------------------------
// Máscaras de campo — formatação automática enquanto o usuário digita
// ---------------------------------------------------------------------------

function mascaraCPF(valor) {
  return valor
    .replace(/\D/g, "")
    .slice(0, 11)
    .replace(/(\d{3})(\d)/, "$1.$2")
    .replace(/(\d{3})(\d)/, "$1.$2")
    .replace(/(\d{3})(\d{1,2})$/, "$1-$2");
}

function mascaraTelefone(valor) {
  const digits = valor.replace(/\D/g, "").slice(0, 11);
  if (digits.length <= 10) {
    return digits.replace(/(\d{2})(\d{4})(\d{0,4})/, "($1) $2-$3");
  }
  return digits.replace(/(\d{2})(\d{5})(\d{0,4})/, "($1) $2-$3");
}

function mascaraCEP(valor) {
  return valor
    .replace(/\D/g, "")
    .slice(0, 8)
    .replace(/(\d{5})(\d)/, "$1-$2");
}

// Aplica as máscaras nos campos correspondentes
document.getElementById("cpf").addEventListener("input", (e) => {
  e.target.value = mascaraCPF(e.target.value);
});
document.getElementById("telefone").addEventListener("input", (e) => {
  e.target.value = mascaraTelefone(e.target.value);
});
document.getElementById("cep").addEventListener("input", (e) => {
  e.target.value = mascaraCEP(e.target.value);
});


// ---------------------------------------------------------------------------
// Converte data ISO (YYYY-MM-DD) para os três campos do template
// Retorna: { data_dia, data_mes, data_ano }
// ---------------------------------------------------------------------------
function decomporData(dataISO) {
  if (!dataISO) return { data_dia: "", data_mes: "", data_ano: "" };
  const [ano, mes, dia] = dataISO.split("-");
  return {
    data_dia: dia,
    data_mes: MESES_PT[parseInt(mes, 10) - 1],
    data_ano: ano, // ano completo (ex: 2026) — o prepare_template consumiu o "20" fixo do template
  };
}


// ---------------------------------------------------------------------------
// Exibe mensagem de feedback para o usuário
// ---------------------------------------------------------------------------
function exibirStatus(mensagem, tipo = "success") {
  statusMsg.textContent = mensagem;
  statusMsg.className = `status-msg ${tipo}`;
  statusMsg.hidden = false;
  statusMsg.scrollIntoView({ behavior: "smooth", block: "nearest" });
}

function ocultarStatus() {
  statusMsg.hidden = true;
}


// ---------------------------------------------------------------------------
// Controla o estado do botão durante o carregamento
// ---------------------------------------------------------------------------
function setLoading(ativo) {
  submitBtn.disabled = ativo;
  btnText.hidden = ativo;
  btnLoading.hidden = !ativo;
}


// ---------------------------------------------------------------------------
// Valida campos obrigatórios antes de enviar
// ---------------------------------------------------------------------------
const CAMPOS_OBRIGATORIOS = [
  "nome_cliente", "nacionalidade", "estado_civil", "profissao",
  "cpf", "rg", "telefone", "endereco", "cep", "cidade", "data_assinatura",
];

function validar(dados) {
  let valido = true;

  CAMPOS_OBRIGATORIOS.forEach((campo) => {
    const el = document.getElementById(campo);
    if (!el) return;
    const valor = campo === "data_assinatura"
      ? el.value
      : dados[campo];

    if (!valor || valor.toString().trim() === "") {
      el.classList.add("invalid");
      valido = false;
    } else {
      el.classList.remove("invalid");
    }
  });

  return valido;
}


// ---------------------------------------------------------------------------
// Dispara download automático do .docx recebido como blob
// ---------------------------------------------------------------------------
function dispararDownload(blob, nomeArquivo) {
  const url  = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href     = url;
  link.download = nomeArquivo;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}


// ---------------------------------------------------------------------------
// Handler principal — submit do formulário
// ---------------------------------------------------------------------------
form.addEventListener("submit", async (event) => {
  event.preventDefault();
  ocultarStatus();

  // Decompõe a data do input em dia, mês (por extenso) e ano
  const dataRaw = document.getElementById("data_assinatura").value;
  const { data_dia, data_mes, data_ano } = decomporData(dataRaw);

  // Monta o payload que o backend espera (campos do ContratoData)
  const dados = {
    nome_cliente:  document.getElementById("nome_cliente").value.trim(),
    nacionalidade: document.getElementById("nacionalidade").value.trim(),
    estado_civil:  document.getElementById("estado_civil").value.trim(),
    profissao:     document.getElementById("profissao").value.trim(),
    cpf:           document.getElementById("cpf").value.trim(),
    rg:            document.getElementById("rg").value.trim(),
    telefone:      document.getElementById("telefone").value.trim(),
    endereco:      document.getElementById("endereco").value.trim(),
    cep:           document.getElementById("cep").value.trim(),
    cidade:        document.getElementById("cidade").value.trim(),
    data_dia,
    data_mes,
    data_ano,
  };

  // Valida antes de enviar
  if (!validar(dados) || !dataRaw) {
    if (!dataRaw) document.getElementById("data_assinatura").classList.add("invalid");
    exibirStatus("Preencha todos os campos obrigatórios (*) antes de continuar.", "error");
    return;
  }

  setLoading(true);

  try {
    // Envia os dados para o backend
    const response = await fetch(`${API_BASE}/generate-contract`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(dados),
    });

    if (!response.ok) {
      const erro = await response.json().catch(() => ({}));
      throw new Error(erro.detail || `Erro ${response.status} ao gerar contrato`);
    }

    // Recebe o .docx como blob e dispara o download
    const blob = await response.blob();
    const contentDisposition = response.headers.get("Content-Disposition") || "";
    const match = contentDisposition.match(/filename="?([^"]+)"?/);
    const nomeArquivo = match ? match[1] : `contrato_BPC_${dados.nome_cliente}.docx`;

    dispararDownload(blob, nomeArquivo);
    exibirStatus(
      `Contrato gerado com sucesso! Download de "${nomeArquivo}" iniciado.`,
      "success"
    );

  } catch (error) {
    console.error("[Erro]", error);
    exibirStatus(`Falha ao gerar contrato: ${error.message}`, "error");

  } finally {
    setLoading(false);
  }
});


// Remove destaque de inválido quando o usuário começa a corrigir o campo
form.querySelectorAll("input, select, textarea").forEach((el) => {
  el.addEventListener("input", () => el.classList.remove("invalid"));
  el.addEventListener("change", () => el.classList.remove("invalid"));
});
