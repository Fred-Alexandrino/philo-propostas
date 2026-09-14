// ─────────────────────────────────────────────────────────────
// PWA — registro do service worker e botão de instalação
// ─────────────────────────────────────────────────────────────
if ("serviceWorker" in navigator) {
  window.addEventListener("load", () => {
    navigator.serviceWorker.register("sw.js").then((reg) => reg.update());
  });
}

let promptInstalacaoAdiado = null;
const btnInstalar = document.getElementById("btn-instalar");

window.addEventListener("beforeinstallprompt", (event) => {
  event.preventDefault();
  promptInstalacaoAdiado = event;
  btnInstalar.style.display = "inline-block"; // Chrome/Edge (desktop e Android)
});

btnInstalar.onclick = async () => {
  if (!promptInstalacaoAdiado) return;
  promptInstalacaoAdiado.prompt();
  await promptInstalacaoAdiado.userChoice;
  promptInstalacaoAdiado = null;
  btnInstalar.style.display = "none";
};

window.addEventListener("appinstalled", () => {
  btnInstalar.style.display = "none";
});

// iOS/Safari não dispara "beforeinstallprompt" — o usuário instala via
// Compartilhar → Adicionar à Tela de Início. Mostramos essa dica uma vez.
const ehIOS = /iphone|ipad|ipod/i.test(navigator.userAgent);
const jaInstalado = window.matchMedia("(display-mode: standalone)").matches || navigator.standalone;
if (ehIOS && !jaInstalado) {
  btnInstalar.style.display = "inline-block";
  btnInstalar.textContent = "📲 Como instalar";
  btnInstalar.onclick = () => {
    alert('Para instalar no iPhone/iPad: toque no ícone de Compartilhar (□↑) na barra do Safari e depois em "Adicionar à Tela de Início".');
  };
}

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
    const potenciaPainelW = parseFloat(document.getElementById("potenciaPainel").value) || 0;
    const dados = PhiloCalc.calcularDadosIniciais({
      unidadesConsumidoras: unidades,
      hspMensal,
      potenciaPainelW: potenciaPainelW || undefined,
    });
    let texto = `Potência mínima necessária para atender o consumo: <strong>${dados.potenciaMinimaKwp.toFixed(2)} kWp</strong> `
      + `(consumo total considerado: ${Math.round(dados.consumoTotal)} kWh/mês).`;
    if (dados.quantidadePaineisSugerida) {
      const kwpResultante = (dados.quantidadePaineisSugerida * potenciaPainelW) / 1000;
      texto += ` Com placas de ${potenciaPainelW}Wp, seriam necessários pelo menos `
        + `<strong>${dados.quantidadePaineisSugerida} módulos</strong> (${kwpResultante.toFixed(2)} kWp instalados, `
        + `já que não dá pra fracionar um painel).`;
      if (!document.getElementById("quantidadeModulos").value) {
        document.getElementById("quantidadeModulos").value = dados.quantidadePaineisSugerida;
        atualizarPotenciaResultante();
      }
    } else {
      texto += " Informe a potência da placa para ver a quantidade sugerida de módulos.";
    }
    box.innerHTML = texto;
  } catch (err) {
    box.textContent = `Erro ao calcular: ${err.message}`;
  }
};

// Potência do sistema é sempre uma consequência de quantidade × potência do módulo
function atualizarPotenciaResultante() {
  const box = document.getElementById("potenciaResultante");
  const qtd = parseFloat(document.getElementById("quantidadeModulos").value) || 0;
  const painelW = parseFloat(document.getElementById("potenciaPainel").value) || 0;
  if (qtd && painelW) {
    const kwp = (qtd * painelW) / 1000;
    box.style.display = "block";
    box.textContent = `Potência do sistema: ${kwp.toFixed(2)} kWp (${qtd} módulos × ${painelW}Wp)`;
  } else {
    box.style.display = "none";
  }
}
document.getElementById("quantidadeModulos").addEventListener("input", atualizarPotenciaResultante);
document.getElementById("potenciaPainel").addEventListener("input", atualizarPotenciaResultante);

// ─────────────────────────────────────────────────────────────
// DOWNLOAD DIRETO (base64 → Blob) — nada fica salvo no Drive,
// o navegador baixa o arquivo na hora e você escolhe onde salvar.
// ─────────────────────────────────────────────────────────────
function prepararDownload(elementoLink, base64, nomeArquivo, mimeType) {
  const bytes = Uint8Array.from(atob(base64), c => c.charCodeAt(0));
  const blob = new Blob([bytes], { type: mimeType });
  if (elementoLink.dataset.blobUrl) URL.revokeObjectURL(elementoLink.dataset.blobUrl);
  const url = URL.createObjectURL(blob);
  elementoLink.href = url;
  elementoLink.download = nomeArquivo;
  elementoLink.dataset.blobUrl = url;
  elementoLink.removeAttribute("target"); // download não deve abrir em nova aba
}

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
    const quantidadeModulos = parseFloat(document.getElementById("quantidadeModulos").value) || 0;
    if (!potenciaPainelW || !quantidadeModulos) throw new Error("Informe a potência da placa e a quantidade de módulos.");
    const valorMaterialFornecedor = parseFloat(document.getElementById("valorMaterial").value) || 0;
    const maoDeObraInstalacao = parseFloat(document.getElementById("maoDeObraInstalacao").value) || 0;
    const materialInstalacao = parseFloat(document.getElementById("materialInstalacao").value) || 0;
    const impostoInstalacao = parseFloat(document.getElementById("impostoInstalacao").value) || 0;

    const inputs = {
      cidade, uf,
      unidadesConsumidoras: unidades.map(u => ({ consumoMedioMensal: u.consumoMedioMensal, tarifa: u.tarifa })),
      quantidadeModulos,
      potenciaPainelW,
      hspMensal,
      valorMaterialFornecedor,
      maoDeObraInstalacao,
      materialInstalacao,
      impostoInstalacao,
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
    prepararDownload(document.getElementById("link-docx"), data.docxBase64, data.nomeArquivo + ".docx",
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
    prepararDownload(document.getElementById("link-pdf"), data.pdfBase64, data.nomeArquivo + ".pdf", "application/pdf");
    document.getElementById("resultado").style.display = "block";
    status.textContent = "Proposta gerada com sucesso. Clique para baixar os arquivos.";
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
    tbody.innerHTML = '<tr><td colspan="6" class="erro">Configure a URL do Apps Script primeiro.</td></tr>';
    return;
  }
  tbody.innerHTML = '<tr><td colspan="6" style="color:var(--muted)">Carregando...</td></tr>';
  try {
    const resp = await fetch(`${backendUrl}?action=historico`);
    const data = await resp.json();
    if (!data.ok) throw new Error(data.erro);
    if (!data.propostas.length) {
      tbody.innerHTML = '<tr><td colspan="6" style="color:var(--muted)">Nenhuma proposta gerada ainda.</td></tr>';
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
      </tr>
    `).join("");
  } catch (err) {
    tbody.innerHTML = `<tr><td colspan="6" class="erro">Erro ao carregar: ${err.message}</td></tr>`;
  }
}
