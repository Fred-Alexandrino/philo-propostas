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
 *   {{NOME_CLIENTE}} {{FORNECEDOR}}
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
    return jsonOutput({ ok: false, erro: "Ação desconhecida." });
  } catch (err) {
    return jsonOutput({ ok: false, erro: err.message });
  }
}

function jsonOutput(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function getConfig() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Config");
  const values = sheet.getDataRange().getValues();
  const config = {};
  values.forEach(row => { if (row[0]) config[row[0]] = row[1]; });
  return config;
}

function gerarProposta(body) {
  const config = getConfig();
  const templateId = config["TEMPLATE_DOC_ID"];
  const pastaId = config["PASTA_DRIVE_ID"];
  if (!templateId) throw new Error('Configure TEMPLATE_DOC_ID na aba "Config".');

  const r = body.resumo;
  const nomeArquivo = `Proposta - ${body.nomeCliente || "Cliente"} - ${r.local} - ${r.data}`;

  // 1) Copia o template
  const pasta = pastaId ? DriveApp.getFolderById(pastaId) : DriveApp.getRootFolder();
  const copia = DriveApp.getFileById(templateId).makeCopy(nomeArquivo, pasta);
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
    "{{FORNECEDOR}}": body.fornecedor || "",
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
