/**
 * MOTOR DE CÁLCULO — PROPOSTA COMERCIAL PHILO SOLUÇÕES ENERGÉTICAS
 * ==================================================================
 * Replica fielmente as fórmulas de "ORÇAMENTO MODELO - PHILO.xlsx".
 * Cada constante abaixo tem a célula de origem indicada no comentário,
 * para você conferir/ajustar caso a planilha original mude.
 *
 * Uso típico:
 *   const hsp = await buscarHSPMensal(cidade, uf);
 *   const dados = calcularDadosIniciais({ ...inputs, hspMensal: hsp });
 *   const custos = calcularCustos(inputs, dados);
 *   const dre = calcularDRE(custos);
 *   const payback = calcularPayback(dre, dados);
 *   const proposta = montarResumoProposta({ inputs, dados, custos, dre, payback });
 */

// ─────────────────────────────────────────────────────────────
// CONSTANTES DO MODELO (extraídas da planilha — ajuste se mudar)
// ─────────────────────────────────────────────────────────────
const CONST = {
  EFICIENCIA_SISTEMA: 0.8,        // DADOS INICIAIS!F23
  TAMANHO_PLACA_M2: 2.5,          // DADOS INICIAIS!C9 (aprox. p/ placas ~550-710W)
  PESO_PLACA_KG: 25,              // DADOS INICIAIS!C11 (kg por placa)
  BDI: 0.214,                     // DRE!D6 — Benefícios e Despesas Indiretas
  REPASSE_FORNECEDOR: 0.12,       // DRE!C11 — comissão/repasse ao fornecedor do kit
  NF_SERVICOS: 0.07,              // DRE!C12 — imposto sobre nota fiscal de serviço
  PROVISAO_RISCO: 0.015,          // DRE!C19
  DESCONTO_A_VISTA: 0.10,         // DRE!L2 = G4*0.9 → 10% de desconto à vista
  PARCELAS_PADRAO: 12,

  // CUSTOS DIRETOS — valores sugeridos automaticamente (ajustáveis na tela):
  MAO_DE_OBRA_POR_KWP: 350,        // R$/kWp — solicitado por Fred
  IMPOSTO_INSTALACAO_PERC: 0.10,   // 10% em cima do valor final (sale price), calculado de forma fechada (ver sugerirCustosInstalacao)
  MARGEM_LUCRO_PERC: 0.20,         // 20% — margem de lucro padrão, editável por proposta
  // Material CA — tabela CUSTOS PROJETO!I4:L6, varia por tipo de telhado/local
  CUSTOS_INSTALACAO: {
    "telhado_fortaleza": { materialCAPorKwp: 220.00 },
    "solo_fortaleza":     { materialCAPorKwp: 244.20 },
    "telhado_outras":     { materialCAPorKwp: 123.67 },
    "solo_outras":        { materialCAPorKwp: 441.31 },
  },

  // CUSTOS INDIRETOS (por kWp, salvo indicação contrária)
  INSTALACAO_TERCEIRIZADA_POR_KWP: 300,  // CUSTOS PROJETO!D31
  ART_PROJETO: 88.78,                     // CUSTOS PROJETO!D32 (valor fixo)
  VISITA_TECNICA_POR_KWP: 50,             // CUSTOS PROJETO!D33

  // PAYBACK (25 anos) — CUSTOS PROJETO / PAYBACK PROJETO I E II
  REAJUSTE_TARIFARIO_ANUAL: 0.10,   // PAYBACK!D8
  DEGRADACAO_SISTEMA_ANUAL: 0.005,  // PAYBACK!D9
  HORIZONTE_ANOS: 25,

  // Componentes tarifários (Lei 14.300 — fio B / compensação decrescente da GD)
  // PAYBACK!P26:P30 — ajuste se a distribuidora/tarifa do cliente for diferente
  PERC_TE: 0.4611,
  PERC_TUSD: 0.5389,
  PERC_ICMS_SOBRE_TUSD: 0.27,
  BANDEIRA_VERMELHA_MEDIA: 0.0485,
};

// ─────────────────────────────────────────────────────────────
// 1) GEOCODIFICAÇÃO + HSP REAL POR CIDADE (substitui a tabela fixa)
// ─────────────────────────────────────────────────────────────

/**
 * Geocodifica "Cidade, UF" usando Nominatim (OpenStreetMap) — gratuito, sem chave.
 * Requer um User-Agent identificável (exigência da política de uso do Nominatim).
 */
async function geocodificarCidade(cidade, uf) {
  const query = encodeURIComponent(`${cidade}, ${uf}, Brasil`);
  const url = `https://nominatim.openstreetmap.org/search?q=${query}&format=json&limit=1`;
  const resp = await fetch(url, { headers: { "User-Agent": "Philo-Propostas/1.0 (contato@philoenergetica.com.br)" } });
  const data = await resp.json();
  if (!data.length) throw new Error(`Cidade não encontrada: ${cidade}, ${uf}`);
  return { lat: parseFloat(data[0].lat), lon: parseFloat(data[0].lon) };
}

/**
 * Busca a climatologia mensal de irradiância (HSP, kWh/m²/dia) na API pública
 * da NASA POWER (ALLSKY_SFC_SW_DWN = irradiância global no plano horizontal,
 * o mesmo dado que o Cresesb fornece manualmente). Gratuito, sem chave.
 * Retorna um array de 12 valores [Jan..Dez].
 */
async function buscarHSPMensal(cidade, uf) {
  const { lat, lon } = await geocodificarCidade(cidade, uf);
  const url = `https://power.larc.nasa.gov/api/temporal/climatology/point`
    + `?parameters=ALLSKY_SFC_SW_DWN&community=RE&longitude=${lon}&latitude=${lat}&format=JSON`;
  const resp = await fetch(url);
  const data = await resp.json();
  const meses = ["JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG","SEP","OCT","NOV","DEC"];
  const serie = data.properties.parameter.ALLSKY_SFC_SW_DWN;
  return { lat, lon, hspMensal: meses.map(m => serie[m]) };
}

// ─────────────────────────────────────────────────────────────
// 2) DADOS INICIAIS — dimensionamento do sistema
// ─────────────────────────────────────────────────────────────

/**
 * @param {Object} inputs
 * @param {Array<{consumoMedioMensal: number, tarifa: number}>} inputs.unidadesConsumidoras
 * @param {number} inputs.potenciaSistemaKwp   - kWp escolhido (você pode ajustar acima do mínimo)
 * @param {number} inputs.potenciaPainelW
 * @param {Array<number>} inputs.hspMensal     - 12 valores vindos de buscarHSPMensal()
 * @param {Array<number>} inputs.diasPorMes    - opcional, default calendário padrão
 */
function calcularDadosIniciais(inputs) {
  const diasPorMesPadrao = [31,28,31,30,31,30,31,31,30,31,30,31];
  const diasPorMes = inputs.diasPorMes || diasPorMesPadrao;

  // Tarifa média ponderada entre UCs — DADOS INICIAIS!B30:D44
  const consumoTotal = inputs.unidadesConsumidoras.reduce((s, uc) => s + uc.consumoMedioMensal, 0);
  const tarifaMedia = inputs.unidadesConsumidoras.reduce(
    (s, uc) => s + (uc.consumoMedioMensal / consumoTotal) * uc.tarifa, 0
  );

  const hspMedio = inputs.hspMensal.reduce((a, b) => a + b, 0) / 12;               // F20
  const consumoDiario = consumoTotal / 30;                                         // F21
  const hspEfetivo = hspMedio * CONST.EFICIENCIA_SISTEMA;                          // F25
  const potenciaMinimaKwp = consumoDiario / hspEfetivo;                            // F26
  const quantidadePaineisSugerida = inputs.potenciaPainelW
    ? Math.ceil((potenciaMinimaKwp * 1000) / inputs.potenciaPainelW)
    : null; // referência — não define a potência do sistema, só orienta a escolha

  // A potência do sistema é sempre uma consequência da quantidade de módulos
  // escolhida × a potência do módulo — nunca o inverso. Se a quantidade não for
  // informada ainda (ex: só calculando a potência mínima), cai no mínimo teórico
  // exato, sem arredondar para um número de painéis.
  let potenciaSistemaKwp, quantidadePaineis;
  if (inputs.quantidadeModulos) {
    quantidadePaineis = inputs.quantidadeModulos;
    potenciaSistemaKwp = (inputs.quantidadeModulos * inputs.potenciaPainelW) / 1000;
  } else if (inputs.potenciaSistemaKwp) {
    potenciaSistemaKwp = inputs.potenciaSistemaKwp;
    quantidadePaineis = Math.ceil((potenciaSistemaKwp * 1000) / inputs.potenciaPainelW);
  } else {
    potenciaSistemaKwp = potenciaMinimaKwp;
    quantidadePaineis = quantidadePaineisSugerida;
  }
  const areaSistemaM2 = quantidadePaineis * CONST.TAMANHO_PLACA_M2;                 // C10
  const pesoSistemaKg = quantidadePaineis * CONST.PESO_PLACA_KG;                    // C11

  // Geração mensal prevista — PROJEÇÃO DO SISTEMA-ESTIMATIVA!G8:G19
  const geracaoMensalKwh = inputs.hspMensal.map((hsp, i) =>
    potenciaSistemaKwp * hsp * diasPorMes[i] * CONST.EFICIENCIA_SISTEMA
  );
  const geracaoMediaMensalKwh = geracaoMensalKwh.reduce((a, b) => a + b, 0) / 12;
  const geracaoAnualKwh = geracaoMensalKwh.reduce((a, b) => a + b, 0);

  return {
    tarifaMedia, consumoTotal, consumoDiario, hspMedio,
    potenciaMinimaKwp, quantidadePaineisSugerida,
    potenciaSistemaKwp, quantidadePaineis,
    areaSistemaM2, pesoSistemaKg,
    geracaoMensalKwh, geracaoMediaMensalKwh, geracaoAnualKwh,
  };
}

// ─────────────────────────────────────────────────────────────
// 3) CUSTOS DO PROJETO
// ─────────────────────────────────────────────────────────────

/**
 * @param {Object} inputs
 * @param {number} inputs.valorMaterialFornecedor  - valor total do kit (fornecedor)
 * @param {number} inputs.maoDeObraInstalacao      - R$ — preenchido manualmente por proposta
 * @param {number} inputs.materialInstalacao       - R$ — preenchido manualmente por proposta
 * @param {number} inputs.impostoInstalacao        - R$ — preenchido manualmente por proposta
 * @param {Object} dadosIniciais - retorno de calcularDadosIniciais()
 */
function calcularCustos(inputs, dadosIniciais) {
  const materiaisKit = inputs.valorMaterialFornecedor;                              // MATERIAL!G7
  const maoDeObraInstalacao = inputs.maoDeObraInstalacao || 0;
  const materialInstalacao = inputs.materialInstalacao || 0;
  const impostoInstalacao = inputs.impostoInstalacao || 0;

  return { materiaisKit, maoDeObraInstalacao, materialInstalacao, impostoInstalacao };
}

function calcularCustosIndiretos(kwp, extras = {}) {
  const instalacaoTerceirizada = kwp * CONST.INSTALACAO_TERCEIRIZADA_POR_KWP;
  const artProjeto = CONST.ART_PROJETO;
  const visitaTecnica = kwp * CONST.VISITA_TECNICA_POR_KWP;
  const engLaudo = extras.engLaudo || 0;
  const artCat = extras.artCat || 0;
  return instalacaoTerceirizada + artProjeto + visitaTecnica + engLaudo + artCat;
}

/**
 * Sugere valores de mão de obra, material CA e imposto para pré-preencher a
 * tela (o usuário pode ajustar cada um depois):
 *  - Mão de obra: R$ 350/kWp
 *  - Material CA: tabela por tipo de telhado/local (mesma da planilha original)
 *  - Imposto: 10% sobre o valor final da proposta — como o valor final também
 *    depende do imposto (ele entra na soma antes da margem), a conta é resolvida
 *    de forma fechada: valorFinal = base*(1+margem) / (1 - imposto%*(1+margem))
 */
function sugerirCustosInstalacao({ potenciaSistemaKwp, valorMaterialFornecedor, tipoInstalacao, margemLucroPerc }) {
  const kwp = potenciaSistemaKwp;
  const cfg = CONST.CUSTOS_INSTALACAO[tipoInstalacao] || CONST.CUSTOS_INSTALACAO.telhado_outras;
  const margem = margemLucroPerc ?? CONST.MARGEM_LUCRO_PERC;

  const maoDeObraInstalacao = kwp * CONST.MAO_DE_OBRA_POR_KWP;
  const materialInstalacao = kwp * cfg.materialCAPorKwp;

  const base = (valorMaterialFornecedor || 0) + maoDeObraInstalacao + materialInstalacao;
  const divisor = 1 - CONST.IMPOSTO_INSTALACAO_PERC * (1 + margem);
  const valorFinalEstimado = (base * (1 + margem)) / divisor;
  const impostoInstalacao = CONST.IMPOSTO_INSTALACAO_PERC * valorFinalEstimado;

  return { maoDeObraInstalacao, materialInstalacao, impostoInstalacao };
}

// ─────────────────────────────────────────────────────────────
// 4) PREÇO FINAL — kit + material CA + mão de obra + imposto + margem
// ─────────────────────────────────────────────────────────────

/**
 * Valor final = valor do kit + material CA + mão de obra + imposto + margem de lucro.
 * A margem de lucro é um percentual (padrão 20%) sobre a soma dos quatro
 * primeiros itens, e é totalmente editável por proposta.
 */
function calcularDRE(custos, opts = {}) {
  const margemLucroPerc = opts.margemLucroPerc ?? CONST.MARGEM_LUCRO_PERC;

  const subtotal = custos.materiaisKit + custos.maoDeObraInstalacao + custos.materialInstalacao + custos.impostoInstalacao;
  const margemLucro = margemLucroPerc * subtotal;
  const valorFinal = subtotal + margemLucro;

  const totalAVista = valorFinal * (1 - CONST.DESCONTO_A_VISTA);
  const parcelaTotalParcelado = valorFinal / CONST.PARCELAS_PADRAO;
  const parcelaEntradaParcelada = custos.materiaisKit / CONST.PARCELAS_PADRAO;
  const entradaParcelada = valorFinal - custos.materiaisKit;

  return {
    subtotal, margemLucroPerc, margemLucro,
    receita: valorFinal, valorVendaFinal: valorFinal, lucro: margemLucro,
    condicoesComerciais: {
      aVista: totalAVista,
      parcelado: { parcelas: CONST.PARCELAS_PADRAO, valorParcela: parcelaTotalParcelado, total: valorFinal },
      entradaParcelada: { parcelas: CONST.PARCELAS_PADRAO, valorParcela: parcelaEntradaParcelada, entrada: entradaParcelada, total: valorFinal },
    },
  };
}

// ─────────────────────────────────────────────────────────────
// 5) PAYBACK — projeção de 25 anos
// ─────────────────────────────────────────────────────────────

function calcularPayback(dre, dadosIniciais) {
  // Nota: a planilha calcula uma "TARIFA EQUIVALENTE" (TE+TUSD-ICMS+bandeira,
  // PAYBACK!R26:R30) mas o loop de 25 anos usa a tarifa média direto (D5=D44).
  // Mantemos o cálculo abaixo só como referência informativa, sem usá-lo no loop.
  const tarifaEquivalenteRef =
    (dadosIniciais.tarifaMedia * CONST.PERC_TE) +
    (dadosIniciais.tarifaMedia * CONST.PERC_TUSD) * (1 - CONST.PERC_ICMS_SOBRE_TUSD) +
    CONST.BANDEIRA_VERMELHA_MEDIA;

  const anos = [];
  let tarifaAno = dadosIniciais.tarifaMedia;
  let geracaoAno = dadosIniciais.geracaoAnualKwh;
  let acumulado = -dre.valorVendaFinal; // investimento inicial (ano 0)

  for (let ano = 1; ano <= CONST.HORIZONTE_ANOS; ano++) {
    if (ano > 1) {
      tarifaAno *= (1 + CONST.REAJUSTE_TARIFARIO_ANUAL);
      geracaoAno = Math.round(geracaoAno * (1 - CONST.DEGRADACAO_SISTEMA_ANUAL) * 100) / 100;
    }
    // Fator de compensação decrescente da GD (Lei 14.300) — PAYBACK!I5 (fórmula original)
    const fatorCompensacao = 0.4 + 0.6 * (1 - 0.28 * CONST.PERC_TUSD * Math.min(1, 0.15 + 0.15 * ano));
    const economiaAno = geracaoAno * tarifaAno * fatorCompensacao;
    acumulado += economiaAno;
    anos.push({ ano, tarifaAno, geracaoAno, economiaAno, acumulado });
  }

  const primeiroAnoPositivo = anos.find(a => a.acumulado >= 0);
  const anoAnterior = anos[anos.indexOf(primeiroAnoPositivo) - 1] || { acumulado: -dre.valorVendaFinal };
  const fracaoAno = -anoAnterior.acumulado / primeiroAnoPositivo.economiaAno;
  const anosPayback = (primeiroAnoPositivo.ano - 1) + fracaoAno;

  const economiaMes = anos[0].economiaAno / 12;
  const economiaAno1 = anos[0].economiaAno;

  return {
    anos,
    economiaMediaMensal: economiaMes,
    economiaAno1,
    retornoInvestimentoMensal: economiaMes / dre.valorVendaFinal,
    retornoInvestimentoAnual: economiaAno1 / dre.valorVendaFinal,
    paybackAnos: Math.floor(anosPayback),
    paybackMeses: Math.round((anosPayback % 1) * 12),
    tarifaEquivalenteRef,
  };
}

// ─────────────────────────────────────────────────────────────
// 6) RESUMO FINAL — campos prontos para preencher a Proposta em Word
// ─────────────────────────────────────────────────────────────

function montarResumoProposta({ inputs, dadosIniciais, custos, dre, payback }) {
  return {
    local: `${inputs.cidade} - ${inputs.uf}`,
    data: new Date().toLocaleDateString("pt-BR"),
    consumoMedioKwh: Math.round(dadosIniciais.consumoTotal),
    tarifaMedia: dadosIniciais.tarifaMedia,
    potenciaSistemaKwp: dadosIniciais.potenciaSistemaKwp,
    estimativaGeracaoMesKwh: Math.round(dadosIniciais.geracaoMediaMensalKwh),
    quantidadePaineis: dadosIniciais.quantidadePaineis,
    potenciaPainelW: inputs.potenciaPainelW,
    areaNecessariaM2: Math.round(dadosIniciais.areaSistemaM2),
    condicoesComerciais: dre.condicoesComerciais,
    economiaMes: payback.economiaMediaMensal,
    economiaAno: payback.economiaAno1,
    retornoMensalPerc: payback.retornoInvestimentoMensal,
    retornoAnualPerc: payback.retornoInvestimentoAnual,
    payback: `${payback.paybackAnos} ano(s) e ${payback.paybackMeses} mês(es)`,
    premissas: {
      producaoAnualKwh: Math.round(dadosIniciais.geracaoAnualKwh),
      tarifaMediaRsPorKwh: dadosIniciais.tarifaMedia,
      reajusteAnualTarifa: CONST.REAJUSTE_TARIFARIO_ANUAL,
      degradacaoAnualSistema: CONST.DEGRADACAO_SISTEMA_ANUAL,
    },
  };
}

const PhiloCalc = {
  CONST,
  geocodificarCidade,
  buscarHSPMensal,
  calcularDadosIniciais,
  calcularCustos,
  calcularCustosIndiretos,
  sugerirCustosInstalacao,
  calcularDRE,
  calcularPayback,
  montarResumoProposta,
};

// Funciona tanto no Node.js (backend/testes) quanto direto no navegador (frontend)
if (typeof module !== "undefined" && module.exports) {
  module.exports = PhiloCalc;
} else if (typeof window !== "undefined") {
  window.PhiloCalc = PhiloCalc;
}
