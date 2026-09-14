// ─────────────────────────────────────────────────────────────
// NAVEGAÇÃO ENTRE TELAS
// ─────────────────────────────────────────────────────────────
const telas = { nova: "tela-nova", historico: "tela-historico", config: "tela-config" };
function mostrarTela(nome) {
  Object.values(telas).forEach(id => document.getElementById(id).classList.remove("active"));
  document.getElementById(telas[nome]).classList.add("active");
  document.querySelectorAll("nav button").forEach(b => b.classList.remove("active"));
  document.getElementById(`btn-${nome}`).classList.add("active");
  if (nome === "historico") carregarHistorico();
}
document.getElementById("btn-nova").onclick = () => mostrarTela("nova");
document.getElementById("btn-historico").onclick = () => mostrarTela("historico");
document.getElementById("btn-config").onclick = () => mostrarTela("config");

// ─────────────────────────────────────────────────────────────
// CONFIGURAÇÃO (URL do Apps Script, salva no localStorage)
// ─────────────────────────────────────────────────────────────
const CONFIG_KEY = "philo_apps_script_url";
document.getElementById("appsScriptUrl").value = localStorage.getItem(CONFIG_KEY) || "";
document.getElementById("btnSalvarConfig").onclick = () => {
  const url = document.getElementById("appsScriptUrl").value.trim();
  localStorage.setItem(CONFIG_KEY, url);
  document.getElementById("status-config").textContent = "Salvo neste navegador.";
};
function getBackendUrl() {
  return localStorage.getItem(CONFIG_KEY) || "";
}

// ─────────────────────────────────────────────────────────────
// GERAR PROPOSTA
// ─────────────────────────────────────────────────────────────
const MESES_LABEL = ["Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov","Dez"];

document.getElementById("btnGerar").onclick = async () => {
  const status = document.getElementById("status");
  const btn = document.getElementById("btnGerar");
  const backendUrl = getBackendUrl();
  if (!backendUrl) {
    status.innerHTML = '<span class="erro">Configure a URL do Apps Script na aba "Configuração" antes de gerar a proposta.</span>';
    return;
  }

  try {
    btn.disabled = true;

    const cidade = document.getElementById("cidade").value.trim();
    const uf = document.getElementById("uf").value.trim().toUpperCase();
    if (!cidade || !uf) throw new Error("Informe cidade e UF.");

    status.textContent = "Buscando irradiação solar (HSP) real da cidade...";
    const { hspMensal } = await PhiloCalc.buscarHSPMensal(cidade, uf);

    const consumoMeses = Array.from(document.querySelectorAll(".mes")).map(el => parseFloat(el.value) || 0);
    const consumoMedioMensal = consumoMeses.reduce((a, b) => a + b, 0) / 12;
    const tarifa = parseFloat(document.getElementById("tarifa").value) || 0;
    const potenciaPainelW = parseFloat(document.getElementById("potenciaPainel").value) || 0;
    const potenciaSistemaInformada = parseFloat(document.getElementById("potenciaSistema").value);
    const valorMaterialFornecedor = parseFloat(document.getElementById("valorMaterial").value) || 0;
    const tipoInstalacao = document.getElementById("tipoInstalacao").value;

    const inputs = {
      cidade, uf,
      unidadesConsumidoras: [{ consumoMedioMensal, tarifa }],
      potenciaSistemaKwp: isNaN(potenciaSistemaInformada) ? undefined : potenciaSistemaInformada,
      potenciaPainelW,
      hspMensal,
      valorMaterialFornecedor,
      tipoInstalacao,
    };

    status.textContent = "Calculando dimensionamento, custos, DRE e payback...";
    const dadosIniciais = PhiloCalc.calcularDadosIniciais(inputs);
    const custos = PhiloCalc.calcularCustos(inputs, dadosIniciais);
    const dre = PhiloCalc.calcularDRE(custos);
    const payback = PhiloCalc.calcularPayback(dre, dadosIniciais);
    const resumo = PhiloCalc.montarResumoProposta({ inputs, dadosIniciais, custos, dre, payback });

    status.textContent = "Gerando documento (Word/PDF) e salvando no histórico...";
    const payload = {
      action: "gerarProposta",
      nomeCliente: document.getElementById("nomeCliente").value.trim(),
      fornecedor: document.getElementById("fornecedor").value.trim(),
      resumo,
    };

    const resp = await fetch(backendUrl, {
      method: "POST",
      headers: { "Content-Type": "text/plain;charset=utf-8" }, // evita preflight CORS no Apps Script
      body: JSON.stringify(payload),
    });
    const data = await resp.json();
    if (!data.ok) throw new Error(data.erro || "Erro desconhecido no backend.");

    document.getElementById("r-potencia").textContent = resumo.potenciaSistemaKwp.toFixed(2);
    document.getElementById("r-avista").textContent = "R$ " + resumo.condicoesComerciais.aVista.toLocaleString("pt-BR", { minimumFractionDigits: 2 });
    document.getElementById("r-payback").textContent = resumo.payback;
    document.getElementById("link-docx").href = data.linkDocx;
    document.getElementById("link-pdf").href = data.linkPdf;
    document.getElementById("resultado").style.display = "block";
    status.textContent = "Proposta gerada com sucesso.";
  } catch (err) {
    status.innerHTML = `<span class="erro">Erro: ${err.message}</span>`;
  } finally {
    btn.disabled = false;
  }
};

// ─────────────────────────────────────────────────────────────
// HISTÓRICO
// ─────────────────────────────────────────────────────────────
async function carregarHistorico() {
  const tbody = document.getElementById("tabela-historico");
  const backendUrl = getBackendUrl();
  if (!backendUrl) {
    tbody.innerHTML = '<tr><td colspan="7" class="erro">Configure a URL do Apps Script primeiro.</td></tr>';
    return;
  }
  tbody.innerHTML = '<tr><td colspan="7" style="color:var(--muted)">Carregando...</td></tr>';
  try {
    const resp = await fetch(`${backendUrl}?action=historico`);
    const data = await resp.json();
    if (!data.ok) throw new Error(data.erro);
    if (!data.propostas.length) {
      tbody.innerHTML = '<tr><td colspan="7" style="color:var(--muted)">Nenhuma proposta gerada ainda.</td></tr>';
      return;
    }
    tbody.innerHTML = data.propostas.map(p => `
      <tr>
        <td>${p.data}</td>
        <td>${p.cliente}</td>
        <td>${p.local}</td>
        <td>${p.potenciaKwp}</td>
        <td>R$ ${Number(p.valorAVista).toLocaleString("pt-BR", { minimumFractionDigits: 2 })}</td>
        <td>${p.payback}</td>
        <td><a href="${p.linkPdf}" target="_blank" style="color:var(--accent)">PDF</a></td>
      </tr>
    `).join("");
  } catch (err) {
    tbody.innerHTML = `<tr><td colspan="7" class="erro">Erro ao carregar: ${err.message}</td></tr>`;
  }
}
