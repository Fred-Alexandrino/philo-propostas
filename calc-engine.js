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

  // CUSTOS INDIRETOS (por kWp, salvo indicação contrária) — os diretos
  // (mão de obra, material de instalação, imposto) agora são preenchidos
  // manualmente por proposta na tela do painel.
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
  const kwp = dadosIniciais.potenciaSistemaKwp;
  const materiaisKit = inputs.valorMaterialFornecedor;                              // MATERIAL!G7

  // Custos diretos — agora preenchidos manualmente por proposta (variam por
  // fornecedor/região), em vez de estimados por kWp.
  const maoDeObraInstalacao = inputs.maoDeObraInstalacao || 0;
  const materialInstalacao = inputs.materialInstalacao || 0;
  const impostoInstalacao = inputs.impostoInstalacao || 0;
  const custosDiretos = maoDeObraInstalacao + materialInstalacao + impostoInstalacao;

  // Custos indiretos — continuam estimados por kWp (ART, visita técnica, terceirização)
  const instalacaoTerceirizada = kwp * CONST.INSTALACAO_TERCEIRIZADA_POR_KWP;
  const artProjeto = CONST.ART_PROJETO;
  const visitaTecnica = kwp * CONST.VISITA_TECNICA_POR_KWP;
  const engLaudo = inputs.engLaudo || 0;
  const artCat = inputs.artCat || 0;
  const custosIndiretos = instalacaoTerceirizada + artProjeto + visitaTecnica + engLaudo + artCat;

  const custosDiversos = custosDiretos + custosIndiretos;                           // CUSTOS PROJETO!C4

  return { materiaisKit, custosDiretos, custosIndiretos, custosDiversos };
}

// ─────────────────────────────────────────────────────────────
// 4) DRE — precificação e condições comerciais
// ─────────────────────────────────────────────────────────────

function calcularDRE(custos, opts = {}) {
  const comissaoVendaPerc = opts.comissaoVendaPerc || 0;
  const artRespTecPerc = opts.artRespTecPerc || 0;

  // Aplica BDI sobre cada linha de custo — DRE!E6:E7
  const vlrVendaKit = custos.materiaisKit / (1 - CONST.BDI);
  const vlrVendaDiversos = custos.custosDiversos / (1 - CONST.BDI);
  const receita = vlrVendaKit + vlrVendaDiversos;                                   // DRE!E4

  // Impostos — DRE!E9:E12
  const repasseFornecedor = CONST.REPASSE_FORNECEDOR * (receita - custos.materiaisKit);
  const nfServicos = (receita - custos.materiaisKit - repasseFornecedor) * CONST.NF_SERVICOS;
  const totalImpostos = repasseFornecedor + nfServicos;

  // Custos (visão DRE) — DRE!E14:E20
  const comissaoVenda = comissaoVendaPerc * receita;
  const provisaoRisco = CONST.PROVISAO_RISCO * receita;
  const artRespTec = artRespTecPerc * custos.materiaisKit;
  const totalCustos = custos.materiaisKit + custos.custosDiversos + comissaoVenda + provisaoRisco + artRespTec;

  const lucro = receita - totalImpostos - totalCustos;                              // DRE!E22
  const margemLucro = lucro / receita;                                             // DRE!G22

  // Condições comerciais — DRE!G2:M9
  const valorVendaFinal = receita;                                                  // (ajuste manual de centavos não replicado)
  const totalAVista = valorVendaFinal * (1 - CONST.DESCONTO_A_VISTA);
  const parcelaTotalParcelado = valorVendaFinal / CONST.PARCELAS_PADRAO;
  const parcelaEntradaParcelada = custos.materiaisKit / CONST.PARCELAS_PADRAO;
  const entradaParcelada = valorVendaFinal - custos.materiaisKit;

  return {
    receita, valorVendaFinal, totalImpostos, totalCustos, lucro, margemLucro,
    condicoesComerciais: {
      aVista: totalAVista,
      parcelado: { parcelas: CONST.PARCELAS_PADRAO, valorParcela: parcelaTotalParcelado, total: valorVendaFinal },
      entradaParcelada: { parcelas: CONST.PARCELAS_PADRAO, valorParcela: parcelaEntradaParcelada, entrada: entradaParcelada, total: valorVendaFinal },
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
