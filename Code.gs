/**
 * PHILO PROPOSTAS — BACKEND (Google Apps Script)
 * =================================================
 * Cole este código em: Extensões → Apps Script (dentro da planilha "Propostas Philo").
 *
 * COMO FUNCIONA: o modelo é o SEU arquivo .docx de verdade (com marcadores
 * {{TAG}} no lugar dos valores), não uma recriação. O Apps Script abre o
 * .docx como um ZIP (Utilities.unzip), edita o texto interno do documento
 * (word/document.xml) e monta um novo .docx (Utilities.zip) — sem depender
 * do Google Docs para isso, então a formatação original fica 100% intacta.
 *
 * CONFIGURAÇÃO NECESSÁRIA ANTES DE PUBLICAR:
 * 1. Ative o serviço avançado do Drive: no editor do Apps Script, no menu
 *    lateral "Serviços" → "+" → adicione "Drive API" (é só usado para
 *    gerar o PDF; sem isso o Word ainda funciona, mas o PDF não).
 * 2. Faça upload do arquivo "Modelo_Proposta_SEM_GRAFICOS.docx" (ou a versão
 *    mais atual do modelo com marcadores) para o seu Google Drive.
 * 3. Crie uma aba chamada "Config" na planilha com:
 *      A1: TEMPLATE_DOCX_ID   B1: <ID do arquivo .docx que você subiu>
 *    (o ID é o trecho da URL do Drive entre /d/ e /view ou /edit)
 * 4. Implantar → Nova implantação → App da Web → Executar como "Eu",
 *    acesso "Qualquer pessoa".
 *
 * MARCADORES USADOS NO MODELO:
 *   {{DATA}} {{LOCAL}} {{NOME_CLIENTE}} {{ENDERECO_CLIENTE}}
 *   {{CONSUMO_MEDIO}} {{TARIFA_MEDIA}} {{POTENCIA_SISTEMA}} {{ESTIMATIVA_GERACAO}}
 *   {{QTD_PAINEIS}} {{POTENCIA_PAINEL}} {{AREA_NECESSARIA}} {{FORNECEDOR}}
 *   {{VALOR_A_VISTA}} {{VALOR_PARCELADO_TOTAL}} {{VALOR_PARCELA}}
 *   {{VALOR_ENTRADA}} {{VALOR_ENTRADA_PARCELA}}
 *   {{ECONOMIA_MES}} {{ECONOMIA_ANO}} {{RETORNO_MES}} {{RETORNO_ANO}}
 *   {{PAYBACK}} {{PRODUCAO_ANUAL}} {{REAJUSTE_TARIFA}} {{DEGRADACAO_SISTEMA}}
 *   {{UNIDADES_BENEFICIARIAS}}
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
  if (!sheet) throw new Error('Aba "Config" não encontrada. Crie uma aba "Config" com TEMPLATE_DOCX_ID.');
  const values = sheet.getDataRange().getValues();
  const config = {};
  values.forEach(row => { if (row[0]) config[row[0]] = row[1]; });
  return config;
}

// Aceita tanto um ID puro quanto um link completo do Drive colado por engano.
function extrairIdDrive(valor) {
  if (!valor) return null;
  const texto = String(valor).trim();
  const match = texto.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (match) return match[1];
  if (/^[a-zA-Z0-9_-]{15,}$/.test(texto)) return texto;
  return null;
}

function gerarProposta(body) {
  const config = getConfig();
  const templateId = extrairIdDrive(config["TEMPLATE_DOCX_ID"]);
  if (!templateId) throw new Error('Preencha TEMPLATE_DOCX_ID na aba "Config" com o ID do arquivo .docx no Drive.');

  const r = body.resumo;
  const nomeArquivo = `Proposta - ${body.nomeCliente || "Cliente"} - ${r.local} - ${r.data}`;

  // 1) Abre o .docx original como ZIP
  let templateBlob;
  try {
    templateBlob = DriveApp.getFileById(templateId).getBlob();
  } catch (err) {
    throw new Error(`TEMPLATE_DOCX_ID inválido (${templateId}). Verifique o ID do arquivo .docx no Drive.`);
  }
  const arquivosZip = Utilities.unzip(templateBlob);

  // 2) Encontra e edita o word/document.xml (onde fica todo o texto)
  const idxDocXml = arquivosZip.findIndex(b => b.getName() === "word/document.xml");
  if (idxDocXml === -1) throw new Error("Arquivo .docx inválido: word/document.xml não encontrado.");
  let xml = arquivosZip[idxDocXml].getDataAsString("UTF-8");

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
  Object.keys(substituicoes).forEach(tag => {
    // split/join evita problemas de caracteres especiais de regex no valor
    xml = xml.split(tag).join(escaparXml(String(substituicoes[tag])));
  });

  // 3) Substitui o word/document.xml modificado de volta no pacote e remonta o .docx
  arquivosZip[idxDocXml] = Utilities.newBlob(xml, "application/xml", "word/document.xml");
  const novoDocxBlob = Utilities.zip(arquivosZip, nomeArquivo + ".docx")
    .setContentType("application/vnd.openxmlformats-officedocument.wordprocessingml.document");

  // 4) Gera o PDF convertendo uma cópia temporária via Google Drive (removida no final)
  let pdfBase64 = null;
  let arquivoTempId = null;
  try {
    const arquivoTemp = Drive.Files.insert(
      { title: nomeArquivo, mimeType: MimeType.GOOGLE_DOCS },
      novoDocxBlob,
      { convert: true }
    );
    arquivoTempId = arquivoTemp.id;
    const pdfBlob = DriveApp.getFileById(arquivoTempId).getAs("application/pdf");
    pdfBase64 = Utilities.base64Encode(pdfBlob.getBytes());
  } catch (err) {
    // Se a "Drive API" (serviço avançado) não estiver ativada, ainda entregamos o Word.
    pdfBase64 = null;
  } finally {
    if (arquivoTempId) DriveApp.getFileById(arquivoTempId).setTrashed(true);
  }

  // 5) Registra no histórico (só os números — nenhum arquivo fica salvo)
  registrarHistorico({
    data: r.data,
    cliente: body.nomeCliente || "",
    local: r.local,
    potenciaKwp: r.potenciaSistemaKwp.toFixed(2),
    valorAVista: r.condicoesComerciais.aVista.toFixed(2),
    payback: r.payback,
    jsonCompleto: JSON.stringify(body),
  });

  return {
    ok: true,
    nomeArquivo,
    docxBase64: Utilities.base64Encode(novoDocxBlob.getBytes()),
    pdfBase64,
    avisoPdf: pdfBase64 ? null : 'PDF não gerado — ative o serviço avançado "Drive API" no Apps Script (Serviços → + → Drive API) e tente de novo.',
  };
}

// Escapa caracteres especiais de XML nos valores inseridos (evita quebrar o documento
// se algum nome de cliente tiver &, <, > etc.)
function escaparXml(texto) {
  return texto
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

// Se houver unidades beneficiárias (geração compartilhada/autoconsumo remoto),
// monta uma lista de texto com nome + endereço de cada uma.
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
    sheet.appendRow(["Data", "Cliente", "Local", "Potência (kWp)", "Valor à Vista", "Payback", "JSON"]);
  }
  sheet.appendRow([row.data, row.cliente, row.local, row.potenciaKwp, row.valorAVista, row.payback, row.jsonCompleto]);
}

function listarHistorico() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Propostas");
  if (!sheet || sheet.getLastRow() < 2) return [];
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
  return values.map(row => ({
    data: row[0], cliente: row[1], local: row[2], potenciaKwp: row[3],
    valorAVista: row[4], payback: row[5],
  })).reverse();
}
