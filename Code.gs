/**
 * PHILO PROPOSTAS — BACKEND (Google Apps Script)
 * =================================================
 * Cole este código em: Extensões → Apps Script (dentro da planilha "Propostas Philo").
 *
 * CONFIGURAÇÃO NECESSÁRIA ANTES DE PUBLICAR:
 * 1. Crie uma aba chamada "Config" na planilha com:
 *      A1: TEMPLATE_DOC_ID   B1: <ID do Google Doc modelo da proposta>
 *      A2: PASTA_DRIVE_ID    B2: <ID de uma pasta no Drive para guardar os documentos gerados>
 * 2. Prepare o Google Doc modelo (veja instruções abaixo desta função) com marcadores
 *    {{TAG}} nos lugares onde hoje há valores fixos no Word original.
 * 3. Em Publicar → Implantar como app da web:
 *      - Executar como: Eu (sua conta)
 *      - Quem pode acessar: Qualquer pessoa
 * 4. Copie a URL gerada (.../exec) e cole no painel, na aba "Configuração".
 *
 * MARCADORES ESPERADOS NO GOOGLE DOC MODELO (troque os valores do Word original por estes):
 *   {{LOCAL}} {{DATA}} {{CONSUMO_MEDIO}} {{TARIFA_MEDIA}} {{POTENCIA_SISTEMA}}
 *   {{ESTIMATIVA_GERACAO}} {{QTD_PAINEIS}} {{POTENCIA_PAINEL}} {{AREA_NECESSARIA}}
 *   {{VALOR_A_VISTA}} {{VALOR_PARCELADO_TOTAL}} {{VALOR_PARCELA}}
 *   {{VALOR_ENTRADA}} {{VALOR_ENTRADA_PARCELA}}
 *   {{ECONOMIA_MES}} {{ECONOMIA_ANO}} {{RETORNO_MES}} {{RETORNO_ANO}}
 *   {{PAYBACK}} {{PRODUCAO_ANUAL}} {{REAJUSTE_TARIFA}} {{DEGRADACAO_SISTEMA}}
 *   {{NOME_CLIENTE}} {{ENDERECO_CLIENTE}} {{FORNECEDOR}}
 *   {{UNIDADES_BENEFICIARIAS}} (opcional — só preencha esse marcador no modelo se o projeto
 *   for de geração compartilhada / autoconsumo remoto; fica vazio quando há só uma unidade)
 */

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    if (body.action === "gerarProposta") {
      return jsonOutput(gerarProposta(body));
    }
    return jsonOutput({ ok: false, erro: "Ação desconhecida." });
  } catch (err) {
    return jsonOutput({ ok: false, erro: err.message });
  }
}

function doGet(e) {
  try {
    if (e.parameter.action === "historico") {
      return jsonOutput({ ok: true, propostas: listarHistorico() });
    }
    if (e.parameter.action === "setup") {
      return jsonOutput(autoconfigurar());
    }
    return jsonOutput({ ok: false, erro: "Ação desconhecida." });
  } catch (err) {
    return jsonOutput({ ok: false, erro: err.message });
  }
}

// ─────────────────────────────────────────────────────────────
// AUTOCONFIGURAÇÃO — cria a aba Config, a pasta no Drive e o
// Google Doc modelo (já com a identidade visual da Philo e os
// marcadores {{TAG}}) automaticamente. Rode uma vez abrindo, no
// navegador, a URL do Web App + "?action=setup".
// Rodar de novo não duplica nada que já esteja configurado.
// ─────────────────────────────────────────────────────────────
const LOGO_URL = "https://raw.githubusercontent.com/Fred-Alexandrino/philo-propostas/main/logo-doc.png";
const VERDE_PHILO = "#1e8a4f";
const GRAFITE_PHILO = "#3d3d3d";

function autoconfigurar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let configSheet = ss.getSheetByName("Config");
  if (!configSheet) {
    configSheet = ss.insertSheet("Config");
  }
  const config = getConfigSafe(configSheet);

  let pastaId = config["PASTA_DRIVE_ID"];
  if (!pastaId) {
    const pasta = DriveApp.createFolder("Propostas Philo - Documentos");
    pastaId = pasta.getId();
  }

  let templateId = config["TEMPLATE_DOC_ID"];
  if (!templateId) {
    templateId = criarModeloProposta(DriveApp.getFolderById(pastaId));
  }

  configSheet.clear();
  configSheet.appendRow(["TEMPLATE_DOC_ID", templateId]);
  configSheet.appendRow(["PASTA_DRIVE_ID", pastaId]);

  return {
    ok: true,
    mensagem: "Configuração concluída automaticamente.",
    templateDocUrl: `https://docs.google.com/document/d/${templateId}/edit`,
    pastaUrl: `https://drive.google.com/drive/folders/${pastaId}`,
  };
}

function getConfigSafe(sheet) {
  const config = {};
  if (sheet.getLastRow() === 0) return config;
  const values = sheet.getDataRange().getValues();
  values.forEach(row => { if (row[0]) config[row[0]] = row[1]; });
  return config;
}

// Monta o Google Doc modelo do zero, replicando a estrutura da
// "PROPOSTA COMERCIAL - MODELO.docx" com os marcadores {{TAG}} no
// lugar dos valores, e a logo da Philo no topo.
function criarModeloProposta(pasta) {
  const doc = DocumentApp.create("Modelo Proposta Philo");
  const arquivo = DriveApp.getFileById(doc.getId());
  pasta.addFile(arquivo);
  DriveApp.getRootFolder().removeFile(arquivo); // tira da raiz, deixa só na pasta

  const corpo = doc.getBody();
  corpo.setMarginTop(36).setMarginBottom(36);

  // Logo no topo
  try {
    const logoBlob = UrlFetchApp.fetch(LOGO_URL).getBlob();
    const img = corpo.appendImage(logoBlob);
    img.setWidth(180);
    img.setHeight(180 * (img.getHeight() / img.getWidth()));
  } catch (err) {
    corpo.appendParagraph("PHILO SOLUÇÕES ENERGÉTICAS").setBold(true);
  }

  corpo.appendParagraph("");
  addPar(corpo, "DATA: {{DATA}}");
  addPar(corpo, "LOCAL: {{LOCAL}}");
  addPar(corpo, "CLIENTE: {{NOME_CLIENTE}}");
  addPar(corpo, "ENDEREÇO: {{ENDERECO_CLIENTE}}");
  corpo.appendParagraph("");

  addPar(corpo, "Os dados para elaboração da presente proposta têm como base o consumo de energia elétrica em kWh, dos últimos 12 meses.");
  corpo.appendParagraph("");

  addTitulo(corpo, "CONSUMO");
  addTabela(corpo, [
    ["CONSUMO MÉDIO", "{{CONSUMO_MEDIO}} kWh"],
    ["VALOR DA TARIFA", "R$ {{TARIFA_MEDIA}}"],
  ]);

  addPar(corpo, "A potência do sistema de geração é calculada com base no consumo de energia e na irradiação solar do local.");
  addTabela(corpo, [
    ["POTÊNCIA DO SISTEMA", "{{POTENCIA_SISTEMA}} kWp"],
    ["ESTIMATIVA MÉDIA DE GERAÇÃO (MÊS)", "{{ESTIMATIVA_GERACAO}} kWh"],
  ]);

  addTitulo(corpo, "PAINÉIS SOLARES");
  addTabela(corpo, [
    ["QUANTIDADE", "{{QTD_PAINEIS}}"],
    ["POTÊNCIA PLACA", "{{POTENCIA_PAINEL}} Wp"],
    ["NÍVEL DE QUALIDADE", "Tier 1 Bloomberg"],
    ["TECNOLOGIA", "Monocristalino"],
    ["ÁREA NECESSÁRIA", "{{AREA_NECESSARIA}} m²"],
  ]);

  addTitulo(corpo, "INVERSOR");
  addTabela(corpo, [
    ["QUANTIDADE", "1"],
    ["FABRICANTE", "{{FORNECEDOR}}"],
    ["MONITORAMENTO", "WIRELESS"],
  ]);

  addPar(corpo, "*As garantias informadas são de responsabilidade dos fabricantes de cada equipamento. Os equipamentos poderão sofrer alterações em comum acordo entre empresa e cliente.");

  addTitulo(corpo, "CONDIÇÕES COMERCIAIS");
  addTabela(corpo, [
    ["À VISTA", "R$ {{VALOR_A_VISTA}}"],
    ["TOTAL PARCELADO (12x)", "R$ {{VALOR_PARCELA}} — total R$ {{VALOR_PARCELADO_TOTAL}}"],
    ["ENTRADA + PARCELADO", "Entrada R$ {{VALOR_ENTRADA}} + 12x R$ {{VALOR_ENTRADA_PARCELA}}"],
  ]);
  addPar(corpo, "*Esta proposta é válida por 5 dias úteis, ou enquanto durar o estoque.");

  addTitulo(corpo, "ECONOMIA E RETORNO DO INVESTIMENTO");
  addTabela(corpo, [
    ["ECONOMIA MÉDIA MENSAL", "R$ {{ECONOMIA_MES}}"],
    ["RETORNO SOBRE INVESTIMENTO (MÊS)", "{{RETORNO_MES}}"],
    ["ECONOMIA NO ANO", "R$ {{ECONOMIA_ANO}}"],
    ["RETORNO SOBRE INVESTIMENTO (ANO)", "{{RETORNO_ANO}}"],
  ]);

  addTitulo(corpo, "PREMISSAS");
  addTabela(corpo, [
    ["Produção de energia solar em kWh/ano", "{{PRODUCAO_ANUAL}}"],
    ["Tarifa média de energia em R$/kWh", "{{TARIFA_MEDIA}}"],
    ["Aumento anual (máximo) na tarifa de energia elétrica", "{{REAJUSTE_TARIFA}}"],
    ["Decréscimo anual de eficiência no sistema", "{{DEGRADACAO_SISTEMA}}"],
    ["TEMPO DE RETORNO DO INVESTIMENTO (PAY-BACK)", "{{PAYBACK}}"],
  ]);

  const beneficiarias = corpo.appendParagraph("Unidades beneficiárias: {{UNIDADES_BENEFICIARIAS}}");
  beneficiarias.editAsText().setForegroundColor(GRAFITE_PHILO);

  corpo.appendParagraph("");
  corpo.appendParagraph("");
  addPar(corpo, "____________________________________");
  addPar(corpo, "CONTRATANTE");
  corpo.appendParagraph("");
  addPar(corpo, "____________________________________");
  addPar(corpo, "CONTRATADA – PHILO SOLUÇÕES ENERGÉTICAS").setBold(true);
  addPar(corpo, "FRED ALEXANDRINO – ENGENHEIRO ELETRICISTA");
  addPar(corpo, "CREA-CE Nº 061835737-8");

  doc.saveAndClose();
  return doc.getId();
}

function addTitulo(corpo, texto) {
  const p = corpo.appendParagraph(texto);
  p.setBold(true);
  p.editAsText().setForegroundColor(VERDE_PHILO);
  p.setSpacingBefore(14);
  return p;
}

function addPar(corpo, texto) {
  return corpo.appendParagraph(texto);
}

function addTabela(corpo, linhas) {
  const tabela = corpo.appendTable(linhas);
  tabela.getRow(0); // no-op, só garante que existe ao menos 1 linha
  for (let i = 0; i < linhas.length; i++) {
    tabela.getCell(i, 0).editAsText().setBold(true);
  }
  return tabela;
}

function jsonOutput(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function getConfig() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Config");
  if (!sheet) throw new Error('Aba "Config" não encontrada na planilha. Crie uma aba chamada exatamente "Config" com TEMPLATE_DOC_ID e PASTA_DRIVE_ID.');
  const values = sheet.getDataRange().getValues();
  const config = {};
  values.forEach(row => { if (row[0]) config[row[0]] = row[1]; });
  return config;
}

function gerarProposta(body) {
  const config = getConfig();
  const templateId = config["TEMPLATE_DOC_ID"];
  const pastaId = config["PASTA_DRIVE_ID"];
  if (!templateId) throw new Error('Preencha TEMPLATE_DOC_ID na aba "Config" (ID do Google Doc modelo).');

  const r = body.resumo;
  const nomeArquivo = `Proposta - ${body.nomeCliente || "Cliente"} - ${r.local} - ${r.data}`;

  // 1) Copia o template
  let pasta;
  try {
    pasta = pastaId ? DriveApp.getFolderById(pastaId) : DriveApp.getRootFolder();
  } catch (err) {
    throw new Error(`PASTA_DRIVE_ID inválido na aba "Config" (${pastaId}). Verifique o ID da pasta no Drive.`);
  }
  let copia;
  try {
    copia = DriveApp.getFileById(templateId).makeCopy(nomeArquivo, pasta);
  } catch (err) {
    throw new Error(`TEMPLATE_DOC_ID inválido na aba "Config" (${templateId}). Verifique o ID do Google Doc modelo.`);
  }
  const doc = DocumentApp.openById(copia.getId());
  const corpo = doc.getBody();

  const fmt = (n) => Number(n).toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const fmtPerc = (n) => (Number(n) * 100).toFixed(1).replace(".", ",") + "%";

  const substituicoes = {
    "{{LOCAL}}": r.local,
    "{{DATA}}": r.data,
    "{{CONSUMO_MEDIO}}": r.consumoMedioKwh,
    "{{TARIFA_MEDIA}}": fmt(r.tarifaMedia),
    "{{POTENCIA_SISTEMA}}": r.potenciaSistemaKwp.toFixed(2),
    "{{ESTIMATIVA_GERACAO}}": r.estimativaGeracaoMesKwh,
    "{{QTD_PAINEIS}}": r.quantidadePaineis,
    "{{POTENCIA_PAINEL}}": r.potenciaPainelW,
    "{{AREA_NECESSARIA}}": r.areaNecessariaM2,
    "{{VALOR_A_VISTA}}": fmt(r.condicoesComerciais.aVista),
    "{{VALOR_PARCELADO_TOTAL}}": fmt(r.condicoesComerciais.parcelado.total),
    "{{VALOR_PARCELA}}": fmt(r.condicoesComerciais.parcelado.valorParcela),
    "{{VALOR_ENTRADA}}": fmt(r.condicoesComerciais.entradaParcelada.entrada),
    "{{VALOR_ENTRADA_PARCELA}}": fmt(r.condicoesComerciais.entradaParcelada.valorParcela),
    "{{ECONOMIA_MES}}": fmt(r.economiaMes),
    "{{ECONOMIA_ANO}}": fmt(r.economiaAno),
    "{{RETORNO_MES}}": fmtPerc(r.retornoMensalPerc),
    "{{RETORNO_ANO}}": fmtPerc(r.retornoAnualPerc),
    "{{PAYBACK}}": r.payback,
    "{{PRODUCAO_ANUAL}}": r.premissas.producaoAnualKwh,
    "{{REAJUSTE_TARIFA}}": fmtPerc(r.premissas.reajusteAnualTarifa),
    "{{DEGRADACAO_SISTEMA}}": fmtPerc(r.premissas.degradacaoAnualSistema),
    "{{NOME_CLIENTE}}": body.nomeCliente || "",
    "{{ENDERECO_CLIENTE}}": body.enderecoCliente || "",
    "{{FORNECEDOR}}": body.fornecedor || "",
    "{{UNIDADES_BENEFICIARIAS}}": listarUnidadesBeneficiarias(body.unidades),
  };
  Object.keys(substituicoes).forEach(tag => corpo.replaceText(tag, String(substituicoes[tag])));
  doc.saveAndClose();

  // 2) Exporta como .docx
  const token = ScriptApp.getOAuthToken();
  const docxUrl = `https://docs.google.com/document/d/${copia.getId()}/export?format=docx`;
  const docxBlob = UrlFetchApp.fetch(docxUrl, { headers: { Authorization: "Bearer " + token } })
    .getBlob().setName(nomeArquivo + ".docx");
  const arquivoDocx = pasta.createFile(docxBlob);
  arquivoDocx.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  // 3) Exporta como PDF
  const pdfBlob = DriveApp.getFileById(copia.getId()).getAs("application/pdf").setName(nomeArquivo + ".pdf");
  const arquivoPdf = pasta.createFile(pdfBlob);
  arquivoPdf.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  // 4) Registra no histórico
  registrarHistorico({
    data: r.data,
    cliente: body.nomeCliente || "",
    local: r.local,
    potenciaKwp: r.potenciaSistemaKwp.toFixed(2),
    valorAVista: r.condicoesComerciais.aVista.toFixed(2),
    payback: r.payback,
    linkDocx: arquivoDocx.getUrl(),
    linkPdf: arquivoPdf.getUrl(),
    jsonCompleto: JSON.stringify(body),
  });

  return {
    ok: true,
    linkDocx: arquivoDocx.getUrl(),
    linkPdf: arquivoPdf.getUrl(),
  };
}

// Se houver unidades beneficiárias (geração compartilhada/autoconsumo remoto),
// monta uma lista de texto com nome + endereço de cada uma, para uso opcional
// no marcador {{UNIDADES_BENEFICIARIAS}} do modelo.
function listarUnidadesBeneficiarias(unidades) {
  if (!unidades || !unidades.length) return "";
  const beneficiarias = unidades.filter(u => u.tipo === "beneficiaria");
  if (!beneficiarias.length) return "";
  return beneficiarias
    .map((u, i) => `${i + 1}. ${u.nome || "Unidade beneficiária"}${u.endereco ? " — " + u.endereco : ""}`)
    .join("\n");
}

function registrarHistorico(row) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Propostas");
  if (!sheet) {
    sheet = ss.insertSheet("Propostas");
    sheet.appendRow(["Data", "Cliente", "Local", "Potência (kWp)", "Valor à Vista", "Payback", "Link Word", "Link PDF", "JSON"]);
  }
  sheet.appendRow([row.data, row.cliente, row.local, row.potenciaKwp, row.valorAVista, row.payback, row.linkDocx, row.linkPdf, row.jsonCompleto]);
}

function listarHistorico() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Propostas");
  if (!sheet || sheet.getLastRow() < 2) return [];
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues();
  return values.map(row => ({
    data: row[0], cliente: row[1], local: row[2], potenciaKwp: row[3],
    valorAVista: row[4], payback: row[5], linkDocx: row[6], linkPdf: row[7],
  })).reverse();
}
