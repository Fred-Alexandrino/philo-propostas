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
// UNIDADES CONSUMIDORAS (geradora + beneficiárias)
// ─────────────────────────────────────────────────────────────
const MESES_LABEL = ["Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov","Dez"];
let contadorUnidades = 0;

function criarBlocoUnidade(tipo) {
  const idx = contadorUnidades++;
  const isGeradora = tipo === "geradora";
  const div = document.createElement("div");
  div.className = "unidade-bloco";
  div.dataset.unidadeId = idx;
  div.dataset.tipo = tipo;
  div.innerHTML = `
    <div class="unidade-header">
      <span class="tag ${tipo}">${isGeradora ? "Unidade Geradora" : "Unidade Beneficiária"}</span>
      ${isGeradora ? "" : '<button type="button" class="remover">Remover</button>'}
    </div>
    <div class="grid">
      <div class="field"><label>Identificação (opcional)</label><input class="uc-nome" placeholder="Ex: Apto 101"></div>
      <div class="field"><label>Endereço da unidade (opcional)</label><input class="uc-endereco" placeholder="Rua, número, bairro"></div>
    </div>
    <div class="meses-grid">
      ${MESES_LABEL.map(m => `<div class="field"><label>${m}</label><input type="number" class="uc-mes" value="0"></div>`).join("")}
    </div>
    <div class="field" style="margin-top:14px; max-width:260px;">
      <label>Tarifa média (R$/kWh)</label>
      <input type="number" step="0.0001" class="uc-tarifa" placeholder="Ex: 0.9246">
    </div>
  `;
  if (!isGeradora) {
    div.querySelector(".remover").onclick = () => div.remove();
  }
  return div;
}

function coletarUnidades() {
  return Array.from(document.querySelectorAll(".unidade-bloco")).map(bloco => {
    const consumoMeses = Array.from(bloco.querySelectorAll(".uc-mes")).map(el => parseFloat(el.value) || 0);
    const consumoMedioMensal = consumoMeses.reduce((a, b) => a + b, 0) / 12;
    return {
      tipo: bloco.dataset.tipo,
      nome: bloco.querySelector(".uc-nome").value.trim(),
      endereco: bloco.querySelector(".uc-endereco").value.trim(),
      consumoMedioMensal,
      tarifa: parseFloat(bloco.querySelector(".uc-tarifa").value) || 0,
    };
  });
}

document.getElementById("unidades-container").appendChild(criarBlocoUnidade("geradora"));
document.getElementById("btnAddUnidade").onclick = () => {
  document.getElementById("unidades-container").appendChild(criarBlocoUnidade("beneficiaria"));
};

// ─────────────────────────────────────────────────────────────
// CÁLCULO DA POTÊNCIA MÍNIMA (sob demanda — depende do HSP da cidade)
// ─────────────────────────────────────────────────────────────
document.getElementById("btnCalcularPotencia").onclick = async () => {
  const box = document.getElementById("potenciaSugerida");
  const cidade = document.getElementById("cidade").value.trim();
  const uf = document.getElementById("uf").value.trim().toUpperCase();
  if (!cidade || !uf) {
    box.style.display = "block";
    box.textContent = "Informe cidade e UF antes de calcular.";
    return;
  }
  box.style.display = "block";
  box.textContent = "Buscando irradiação solar (HSP) da cidade...";
  try {
    const { hspMensal } = await PhiloCalc.buscarHSPMensal(cidade, uf);
    const unidades = coletarUnidades().map(u => ({ consumoMedioMensal: u.consumoMedioMensal, tarifa: u.tarifa }));
    const dados = PhiloCalc.calcularDadosIniciais({
      unidadesConsumidoras: unidades,
      hspMensal,
      potenciaPainelW: parseFloat(document.getElementById("potenciaPainel").value) || 550,
    });
    box.textContent = `Potência mínima necessária: ${dados.potenciaMinimaKwp.toFixed(2)} kWp `
      + `(${dados.quantidadePaineis} painéis nessa potência de placa) — consumo total considerado: ${Math.round(dados.consumoTotal)} kWh/mês.`;
    if (!document.getElementById("potenciaSistema").value) {
      document.getElementById("potenciaSistema").placeholder = `Sugestão: ${dados.potenciaMinimaKwp.toFixed(2)} kWp`;
    }
  } catch (err) {
    box.textContent = `Erro ao calcular: ${err.message}`;
  }
};

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

    const unidades = coletarUnidades();
    if (!unidades.length) throw new Error("Adicione ao menos a unidade geradora.");
    const potenciaPainelW = parseFloat(document.getElementById("potenciaPainel").value) || 0;
    const potenciaSistemaInformada = parseFloat(document.getElementById("potenciaSistema").value);
    const valorMaterialFornecedor = parseFloat(document.getElementById("valorMaterial").value) || 0;
    const tipoInstalacao = document.getElementById("tipoInstalacao").value;

    const inputs = {
      cidade, uf,
      unidadesConsumidoras: unidades.map(u => ({ consumoMedioMensal: u.consumoMedioMensal, tarifa: u.tarifa })),
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
      enderecoCliente: document.getElementById("enderecoCliente").value.trim(),
      fornecedor: document.getElementById("fornecedor").value.trim(),
      unidades: unidades.map(u => ({ tipo: u.tipo, nome: u.nome, endereco: u.endereco })),
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
