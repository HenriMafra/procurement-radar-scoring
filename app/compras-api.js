// =====================================================================
//  compras-api.js - consulta REAL de status de pregao
//
//  Fonte: API de dados abertos do Compras.gov (espelho do PNCP / Lei 14.133)
//    https://dadosabertos.compras.gov.br/modulo-contratacoes/1_consultarContratacoes_PNCP_14133
//
//  Armadilhas ja pagas (nao mexer sem reler):
//   1. unidadeOrgaoCodigoUnidade EXIGE 6 digitos com zero a esquerda.
//      "70001"  -> 0 registros
//      "070001" -> 36 registros   (mesma UASG, mesmo periodo)
//   2. codigoModalidade e OBRIGATORIO. Sem ele a API responde 404.
//      5 = Pregao Eletronico (legado SIASG) | 6 = Pregao Eletronico (14.133)
//      O mesmo pregao pode estar em uma ou outra; por isso tentamos as duas.
//   3. NAO filtrar por CNPJ para identificar orgao: a Justica Eleitoral inteira
//      usa o CNPJ do TSE. Buscando 90009/2026 por CNPJ vem 17 pregoes distintos,
//      um de cada TRE. Quem identifica a unidade e a UASG.
// =====================================================================
'use strict';

const https = require('https');

const HOST = 'dadosabertos.compras.gov.br';
const ROTA = '/modulo-contratacoes/1_consultarContratacoes_PNCP_14133';
const MODALIDADES = [5, 6];
const MAX_PAGINAS = 6;

function getJson(url, timeoutMs = 45000) {
  return new Promise((resolve) => {
    const req = https.get(url, {
      rejectUnauthorized: false,
      headers: { Accept: 'application/json', 'User-Agent': 'ENTERPRISECORE-Radar/1.0' },
    }, (r) => {
      let d = '';
      r.setEncoding('utf8');
      r.on('data', (c) => { d += c; });
      r.on('end', () => {
        if (r.statusCode !== 200) return resolve({ ok: false, status: r.statusCode, erro: `HTTP ${r.statusCode}` });
        try { resolve({ ok: true, status: 200, dados: JSON.parse(d) }); }
        catch (e) { resolve({ ok: false, status: 200, erro: 'resposta nao e JSON valido' }); }
      });
    });
    req.on('error', (e) => resolve({ ok: false, status: 0, erro: e.message }));
    req.setTimeout(timeoutMs, () => { req.destroy(); resolve({ ok: false, status: 0, erro: 'timeout' }); });
  });
}

// "70001" -> "070001" | "9777/2024" -> null (nao e UASG, e numero de processo)
function normalizaUasg(v) {
  const s = String(v == null ? '' : v).trim();
  if (!/^\d{1,6}$/.test(s)) return null;
  return s.padStart(6, '0');
}

function normalizaNumero(v) {
  const s = String(v == null ? '' : v).trim();
  const m = /(\d{1,6})/.exec(s.replace(/\D+$/, ''));
  return m ? String(Number(m[1])) : null;
}

function hoje() { return new Date().toISOString().slice(0, 10); }

/**
 * Busca um pregao especifico. Devolve { ok, achado, registro, tentativas, erro }.
 */
async function consultarPregao({ uasg, numero, ano }) {
  const u = normalizaUasg(uasg);
  const n = normalizaNumero(numero);
  const a = String(ano || new Date().getFullYear()).slice(0, 4);

  if (!u) return { ok: false, achado: false, erro: `UASG invalida: "${uasg}" (precisa ser numerica, ate 6 digitos)` };
  if (!n) return { ok: false, achado: false, erro: `Numero do pregao invalido: "${numero}"` };

  const fim = a === String(new Date().getFullYear()) ? hoje() : `${a}-12-31`;
  const tentativas = [];

  for (const mod of MODALIDADES) {
    for (let pag = 1; pag <= MAX_PAGINAS; pag++) {
      const url = `https://${HOST}${ROTA}?pagina=${pag}&tamanhoPagina=100`
        + `&unidadeOrgaoCodigoUnidade=${u}&codigoModalidade=${mod}`
        + `&dataPublicacaoPncpInicial=${a}-01-01&dataPublicacaoPncpFinal=${fim}`;

      const r = await getJson(url);
      if (!r.ok) { tentativas.push(`mod ${mod} pag ${pag}: ${r.erro}`); break; }

      const itens = Array.isArray(r.dados && r.dados.resultado) ? r.dados.resultado : [];
      const achado = itens.find((x) => String(x.numeroCompra) === n);
      if (achado) return { ok: true, achado: true, registro: mapear(achado), tentativas };

      if (itens.length < 100) break; // ultima pagina
    }
  }
  return { ok: true, achado: false, tentativas, erro: `Pregao ${n}/${a} nao encontrado na UASG ${u}` };
}

function mapear(x) {
  return {
    numeroControlePNCP: x.numeroControlePNCP || '',
    idCompra: x.idCompra || '',
    orgao: x.orgaoEntidadeRazaoSocial || '',
    unidadeCodigo: x.unidadeOrgaoCodigoUnidade || '',
    unidadeNome: x.unidadeOrgaoNomeUnidade || '',
    uf: x.unidadeOrgaoUfSigla || '',
    numeroCompra: String(x.numeroCompra || ''),
    ano: String(x.anoCompraPncp || ''),
    modalidade: x.modalidadeNome || '',
    situacao: x.situacaoCompraNomePncp || '',
    existeResultado: !!x.existeResultado,
    srp: !!x.srp,
    processo: x.processo || '',
    objeto: x.objetoCompra || '',
    valorEstimado: Number(x.valorTotalEstimado || 0),
    valorHomologado: Number(x.valorTotalHomologado || 0),
    dataPublicacao: x.dataPublicacaoPncp || '',
    dataAbertura: x.dataAberturaPropostaPncp || '',
    dataEncerramento: x.dataEncerramentoPropostaPncp || '',
    dataAtualizacao: x.dataAtualizacaoPncp || '',
    excluida: !!x.contratacaoExcluida,
    // e o idCompra que vai no parametro compra=, NAO a UASG
    linkPortal: x.idCompra
      ? `https://cnetmobile.estaleiro.serpro.gov.br/comprasnet-web/public/compras/acompanhamento-compra?compra=${x.idCompra}`
      : '',
  };
}

/**
 * Resultados por item (quem venceu, quanto foi homologado).
 * A resposta NAO traz numeroCompra: a chave de juncao e o idCompra.
 * O endpoint de ITENS (2_consultar...) foi testado e da timeout de forma
 * consistente mesmo com pagina de 100 - por isso nao entra no polling.
 */
async function consultarResultados({ uasg, idCompra, ano }) {
  const u = normalizaUasg(uasg);
  if (!u || !idCompra) return [];
  const a = String(ano || new Date().getFullYear()).slice(0, 4);
  const fim = a === String(new Date().getFullYear()) ? hoje() : `${a}-12-31`;

  const achados = [];
  for (let pag = 1; pag <= 4; pag++) {
    const url = `https://${HOST}/modulo-contratacoes/3_consultarResultadoItensContratacoes_PNCP_14133`
      + `?pagina=${pag}&tamanhoPagina=500&unidadeOrgaoCodigoUnidade=${u}`
      + `&dataResultadoPncpInicial=${a}-01-01&dataResultadoPncpFinal=${fim}`;
    const r = await getJson(url);
    if (!r.ok) break;
    const itens = Array.isArray(r.dados && r.dados.resultado) ? r.dados.resultado : [];
    for (const x of itens) {
      if (String(x.idCompra) !== String(idCompra)) continue;
      achados.push({
        item: x.numeroItemPncp,
        fornecedor: x.nomeRazaoSocialFornecedor || '',
        cnpjFornecedor: x.niFornecedor || '',
        porte: x.porteFornecedorNome || '',
        quantidade: Number(x.quantidadeHomologada || 0),
        valorUnitario: Number(x.valorUnitarioHomologado || 0),
        valorTotal: Number(x.valorTotalHomologado || 0),
        situacao: x.situacaoCompraItemResultadoNome || '',
        cancelamento: x.motivoCancelamento || '',
      });
    }
    if (itens.length < 500) break;
  }
  achados.sort((p, q) => Number(p.item) - Number(q.item));
  return achados;
}

// ---------------------------------------------------------------------
//  PNCP: historico de manutencao e documentos publicados.
//
//  ESTE E O EQUIVALENTE PUBLICO DA ABA "AVISOS" DO QUADRO INFORMATIVO.
//  Confirmado no CNJ 90012/2026: o portal mostrava "Evento de Suspensao...
//  Motivo: Para fins de retificacao do edital" e o /historico devolve, no
//  mesmo minuto, tipoLogManutencaoNome=Retificacao com essa justificativa.
//  Nao passa por captcha: e a API aberta do PNCP.
//
//  O que ele NAO traz: o texto das mensagens do pregoeiro na sala de
//  disputa, nem esclarecimentos/impugnacoes. Isso continua fora de alcance.
// ---------------------------------------------------------------------
const PNCP = 'pncp.gov.br';

// "07421906000129-1-000033/2026" -> { cnpj, seq, ano }
function partesControlePncp(numeroControlePNCP) {
  const m = /^(\d{14})-\d+-0*(\d+)\/(\d{4})$/.exec(String(numeroControlePNCP || '').trim());
  return m ? { cnpj: m[1], seq: m[2], ano: m[3] } : null;
}

async function consultarPncpDetalhe(numeroControlePNCP) {
  const p = partesControlePncp(numeroControlePNCP);
  if (!p) return { ok: false, eventosOk: false, arquivosOk: false, eventos: [], arquivos: [] };
  const base = `https://${PNCP}/api/pncp/v1/orgaos/${p.cnpj}/compras/${p.ano}/${p.seq}`;

  const [rh, ra] = await Promise.all([getJson(`${base}/historico`), getJson(`${base}/arquivos`)]);

  const eventos = (rh.ok && Array.isArray(rh.dados) ? rh.dados : []).map((e) => ({
    quando: e.logManutencaoDataInclusao || '',
    tipo: e.tipoLogManutencaoNome || '',
    categoria: e.categoriaLogManutencaoNome || '',
    justificativa: e.justificativa || '',
    documentoTipo: e.documentoTipo || '',
    documentoTitulo: e.documentoTitulo || '',
    por: e.usuarioNome || '',
  })).sort((a, b) => String(a.quando).localeCompare(String(b.quando)));

  const arquivos = (ra.ok && Array.isArray(ra.dados) ? ra.dados : [])
    .filter((a) => a.statusAtivo !== false)
    .map((a) => ({
      titulo: a.titulo || '',
      tipo: a.tipoDocumentoNome || a.tipoDocumentoDescricao || '',
      publicadoEm: a.dataPublicacaoPncp || '',
      url: a.url || a.uri || '',
    }));

  return {
    ok: rh.ok || ra.ok,
    eventosOk: rh.ok,
    arquivosOk: ra.ok,
    eventos,
    arquivos
  };
}

async function consultarPncpResultados(numeroControlePNCP) {
  const p = partesControlePncp(numeroControlePNCP);
  if (!p) return [];
  const base = `https://${PNCP}/api/pncp/v1/orgaos/${p.cnpj}/compras/${p.ano}/${p.seq}`;
  const rItens = await getJson(`${base}/itens`);
  const itens = (rItens.ok && Array.isArray(rItens.dados)) ? rItens.dados : [];
  const achados = [];
  for (const it of itens) {
    const num = it.numeroItem;
    const rRes = await getJson(`${base}/itens/${num}/resultados`);
    if (rRes.ok && Array.isArray(rRes.dados)) {
      for (const res of rRes.dados) {
        achados.push({
          item: num,
          fornecedor: res.nomeRazaoSocialFornecedor || '',
          cnpjFornecedor: res.niFornecedor || '',
          porte: res.porteFornecedorNome || '',
          quantidade: Number(res.quantidadeHomologada || 0),
          valorUnitario: Number(res.valorUnitarioHomologado || 0),
          valorTotal: Number(res.valorTotalHomologado || 0),
          situacao: res.situacaoCompraItemResultadoNome || '',
          dataResultado: res.dataResultado || '',
          dataInclusao: res.dataInclusao || '',
          cancelamento: res.motivoCancelamento || '',
        });
      }
    }
  }
  achados.sort((a, b) => Number(a.item) - Number(b.item));
  return achados;
}

/** Consulta completa: contratacao + resultados + historico/documentos do PNCP. */
async function consultarTudo({ uasg, numero, ano, numeroControlePNCP, registroAtual }) {
  let base = await consultarPregao({ uasg, numero, ano });

  // Se não localizou na API de dadosabertos mas temos o controle PNCP, usamos o PNCP diretamente
  if (!base.achado && numeroControlePNCP) {
    base = {
      ok: true,
      achado: true,
      registro: Object.assign({}, registroAtual || {}, {
        numeroControlePNCP,
        numeroCompra: String(numero || ''),
        ano: String(ano || ''),
        uasg: String(uasg || ''),
        unidadeCodigo: String(uasg || ''),
        modalidade: 'Pregão - Eletrônico'
      })
    };
  }

  if (!base.achado) return base;
  const g = base.registro;
  const ctrl = g.numeroControlePNCP || numeroControlePNCP;

  // 1. Resultados prioritariamente pela API oficial direta do PNCP
  let resultados = [];
  if (ctrl) {
    try { resultados = await consultarPncpResultados(ctrl); } catch (_) {}
  }
  // Fallback para espelho dadosabertos caso PNCP não retorne itens
  if (!resultados.length && g.idCompra) {
    try { resultados = await consultarResultados({ uasg, idCompra: g.idCompra, ano }); } catch (_) {}
  }

  g.resultados = resultados;
  g.qtdResultados = resultados.length;
  g.existeResultado = resultados.length > 0;
  g.totalHomologadoItens = resultados.reduce((s, x) => s + (Number(x.valorTotal) || 0), 0);
  g.valorHomologado = g.totalHomologadoItens;
  g.vencedores = [...new Set(resultados.map((x) => x.fornecedor).filter(Boolean))];

  // 2. Histórico e Arquivos oficiais do PNCP (com blindagem contra quedas de rede)
  let det = null;
  if (ctrl) {
    try { det = await consultarPncpDetalhe(ctrl); } catch (_) {}
  }

  // Preservação do histórico: só atualiza se det respondeu com sucesso E quantidade válida
  if (det && det.eventosOk && Array.isArray(det.eventos) && det.eventos.length >= (registroAtual?.eventosPncp?.length || 0)) {
    g.eventosPncp = det.eventos;
    g.qtdEventosPncp = det.eventos.length;
  } else if (registroAtual && Array.isArray(registroAtual.eventosPncp) && registroAtual.eventosPncp.length) {
    g.eventosPncp = registroAtual.eventosPncp;
    g.qtdEventosPncp = registroAtual.eventosPncp.length;
  } else {
    g.eventosPncp = det?.eventos || [];
    g.qtdEventosPncp = g.eventosPncp.length;
  }

  // Preservação de arquivos: só atualiza se respondeu com sucesso E quantidade válida
  if (det && det.arquivosOk && Array.isArray(det.arquivos) && det.arquivos.length >= (registroAtual?.arquivos?.length || 0)) {
    g.arquivos = det.arquivos;
    g.qtdArquivos = det.arquivos.length;
  } else if (registroAtual && Array.isArray(registroAtual.arquivos) && registroAtual.arquivos.length) {
    g.arquivos = registroAtual.arquivos;
    g.qtdArquivos = registroAtual.arquivos.length;
  } else {
    g.arquivos = det?.arquivos || [];
    g.qtdArquivos = g.arquivos.length;
  }

  const ult = g.eventosPncp[g.eventosPncp.length - 1];
  g.ultimoEventoPncp = ult ? `${ult.tipo}${ult.justificativa ? ' — ' + ult.justificativa : ''}` : '';

  return base;
}

// Campos cuja mudanca significa movimentacao real no pregao
const VIGIADOS = [
  ['situacao', 'Situação'],
  ['existeResultado', 'Resultado publicado'],
  ['valorHomologado', 'Valor homologado'],
  ['dataEncerramento', 'Encerramento das propostas'],
  ['dataAbertura', 'Abertura das propostas'],
  ['excluida', 'Contratação excluída'],
  ['qtdResultados', 'Itens com resultado'],
  ['totalHomologadoItens', 'Total homologado nos itens'],
  ['qtdEventosPncp', 'Eventos no PNCP'],
  ['qtdArquivos', 'Documentos publicados'],
];

// Campo ausente e campo zerado sao a MESMA coisa. Sem isso, adicionar um campo
// novo ao snapshot faz todo item antigo parecer "mudou" (undefined -> 0).
const NUMERICOS = new Set(['valorHomologado', 'qtdResultados', 'totalHomologadoItens', 'qtdEventosPncp', 'qtdArquivos']);
function valorNorm(campo, v) {
  if (NUMERICOS.has(campo)) return String(Number(v || 0));
  if (v === null || v === undefined) return '';
  if (typeof v === 'boolean') return v ? 'true' : 'false';
  return String(v);
}

function formatarDataPtBr(isoStr) {
  if (!isoStr) return '';
  try {
    const d = new Date(isoStr);
    if (isNaN(d.getTime())) return isoStr;
    return d.toLocaleString('pt-BR', { timeZone: 'America/Sao_Paulo' });
  } catch (_) {
    return isoStr;
  }
}

function assinatura(reg) {
  if (!reg) return '';
  return VIGIADOS.map(([k]) => `${k}=${valorNorm(k, reg[k])}`).join('|');
}

/** Traduz e agrupa eventos técnicos do portal em informações claras e apresentáveis */
function sintetizarEventosPncp(novosEventos) {
  const avisos = [];
  let itensEmAndamento = 0;
  let itensComResultado = 0;
  let arquivosSubstituidos = 0;
  let inclusaoContratacao = false;
  let inclusaoDocumento = false;

  for (const e of novosEventos) {
    const cat = (e.categoria || '').toLowerCase();
    const just = (e.justificativa || '').toLowerCase();
    const tipo = (e.tipo || '').toLowerCase();

    if (just.includes('mudança de situação para em andamento')) {
      itensEmAndamento++;
      continue;
    }
    if (just.includes('inclusão do resultado do item') || cat.includes('resultado de item')) {
      itensComResultado++;
      continue;
    }
    if (just.includes('arquivo excluído pelo sistema para inclusão de um novo')) {
      arquivosSubstituidos++;
      continue;
    }
    if (tipo === 'inclusão' && cat.includes('contratação') && !e.justificativa) {
      // SÓ considera Registro Inicial se o evento ocorreu recentemente (últimos 3 dias)
      const dt = new Date(e.quando);
      const diffDias = !isNaN(dt.getTime()) ? (Date.now() - dt.getTime()) / (1000 * 60 * 60 * 24) : 999;
      if (diffDias <= 3) {
        inclusaoContratacao = true;
      }
      continue;
    }
    if (tipo === 'inclusão' && cat.includes('documento') && !e.justificativa) {
      inclusaoDocumento = true;
      continue;
    }
    if (cat.includes('item de contratação') && !e.justificativa) {
      continue;
    }

    // Eventos com justificativa ou decisões explícitas do órgão
    if (tipo.includes('suspens')) {
      avisos.push(`🛑 Suspensão do Certame: ${e.justificativa || 'Processo suspenso no portal oficial.'}`);
    } else if (tipo.includes('cancel')) {
      avisos.push(`❌ Cancelamento do Certame: ${e.justificativa || 'Processo cancelado pelo órgão.'}`);
    } else if (tipo.includes('anula')) {
      avisos.push(`⛔ Anulação do Certame: ${e.justificativa || 'Processo anulado por vício formal.'}`);
    } else if (tipo.includes('revog')) {
      avisos.push(`⚠️ Revogação do Certame: ${e.justificativa || 'Processo revogado por interesse público.'}`);
    } else if (tipo.includes('retifica')) {
      const doc = e.documentoTipo ? ` [${e.documentoTipo}]` : '';
      avisos.push(`📝 Retificação no Edital${doc}: ${e.justificativa || 'Edital alterado pelo órgão licitante.'}`);
    } else {
      avisos.push(`📌 Aviso Oficial (${e.tipo}${e.categoria ? ' - ' + e.categoria : ''}): ${e.justificativa || 'Atualização publicada pelo órgão.'}`);
    }
  }

  if (itensEmAndamento > 0) {
    avisos.push(`🔄 Andamento da Disputa: ${itensEmAndamento > 1 ? itensEmAndamento + ' itens entraram' : 'Item entrou'} na fase 'Em Andamento' (abertura/análise de propostas)`);
  }
  if (itensComResultado > 0) {
    avisos.push(`⚖️ Julgamento de Propostas: Resultados e classificação formalizados no sistema pelo pregoeiro`);
  }
  if (arquivosSubstituidos > 0) {
    avisos.push(`📄 Atualização de Documentos: Versão anterior de arquivo substituída por nova versão oficial`);
  }
  if (inclusaoContratacao) {
    avisos.push(`📢 Registro Inicial: Edital registrado e publicado oficialmente no portal PNCP`);
  }
  if (inclusaoDocumento) {
    avisos.push(`📄 Inserção de Documentos: Edital e anexos disponibilizados no sistema`);
  }

  return avisos;
}

/** Compara duas leituras e gera informações 100% completas, claras e apresentáveis */
function compararRegistros(antes, depois) {
  if (!antes || !depois) return [];
  const mudancas = [];

  // 1. MUDANÇA DE SITUAÇÃO OFICIAL (ex: Divulgada -> Homologada / Suspensa)
  if (antes.situacao !== depois.situacao && depois.situacao) {
    mudancas.push(`📢 Situação no Portal: Alterada de "${antes.situacao || 'Inicial'}" para "${depois.situacao}"`);
  }

  // 2. EXCLUSÃO
  if (!antes.excluida && depois.excluida) {
    mudancas.push('⚠️ Alerta Crítico: Contratação foi EXCLUÍDA no portal oficial');
  }

  // 3. DATAS CRÍTICAS (Prorrogações de Abertura / Encerramento)
  if (antes.dataAbertura && depois.dataAbertura && antes.dataAbertura !== depois.dataAbertura) {
    mudancas.push(`📅 Nova Data de Abertura de Propostas: ${formatarDataPtBr(depois.dataAbertura)}`);
  }
  if (antes.dataEncerramento && depois.dataEncerramento && antes.dataEncerramento !== depois.dataEncerramento) {
    mudancas.push(`📅 Novo Prazo de Encerramento: ${formatarDataPtBr(depois.dataEncerramento)}`);
  }

  // 4. RESULTADOS E HOMOLOGAÇÃO
  const qtdResAntes = Number(antes.qtdResultados || 0);
  const qtdResDepois = Number(depois.qtdResultados || 0);
  const valAntes = Number(antes.valorHomologado || antes.totalHomologadoItens || 0);
  const valDepois = Number(depois.valorHomologado || depois.totalHomologadoItens || 0);
  const temNovoResultado = (!antes.existeResultado && depois.existeResultado) || (qtdResDepois > qtdResAntes);

  if (temNovoResultado || (valDepois > valAntes && valDepois > 0)) {
    const novosResultados = (depois.resultados || []).slice(qtdResAntes);
    const valFmt = valDepois > 0 ? valDepois.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' }) : '';
    
    if (novosResultados.length) {
      const listaQuem = novosResultados.map((x) => {
        const vFmt = Number(x.valorTotal || 0).toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
        return `Item ${x.item}: ${x.fornecedor} (${vFmt})`;
      }).join('; ');
      
      const totalFmt = valFmt ? ` (Valor Total Homologado: ${valFmt})` : '';
      mudancas.push(`🏆 Resultado Oficial Publicado${totalFmt} — ${listaQuem}`);
    } else if (valFmt) {
      mudancas.push(`🏆 Resultado Homologado no Portal: Valor Total de ${valFmt}`);
    } else {
      mudancas.push('🏆 Resultado Oficial de julgamento publicado no portal');
    }
  }

  // 5. DOCUMENTOS ANEXADOS (Chave única estrita para nunca repetir arquivo conhecido)
  const arquivosAntesChaves = new Set((antes.arquivos || []).map(a => a.url || `${a.titulo}|${a.tipo}`));
  const novosArquivos = (depois.arquivos || []).filter(a => !arquivosAntesChaves.has(a.url || `${a.titulo}|${a.tipo}`));
  for (const arq of novosArquivos) {
    const nomeDoc = arq.titulo || arq.tipo || 'Documento';
    const tipoDoc = arq.tipo && arq.titulo && arq.tipo !== arq.titulo ? ` (Tipo: ${arq.tipo})` : '';
    mudancas.push(`📄 Arquivo Disponível: "${nomeDoc}"${tipoDoc} [Disponível para download]`);
  }

  // 6. COMUNICADOS E EVENTOS DO ÓRGÃO TRADUZIDOS (Chave única estrita para nunca repetir evento conhecido)
  const eventosAntesChaves = new Set((antes.eventosPncp || []).map(e => `${e.quando}|${e.tipo}|${e.categoria}|${e.justificativa}`));
  const novosEventos = (depois.eventosPncp || []).filter(e => !eventosAntesChaves.has(`${e.quando}|${e.tipo}|${e.categoria}|${e.justificativa}`));
  if (novosEventos.length) {
    const traduzidos = sintetizarEventosPncp(novosEventos);
    for (const t of traduzidos) {
      mudancas.push(t);
    }
  }

  return mudancas;
}

module.exports = {
  consultarPregao, consultarResultados, consultarTudo,
  assinatura, compararRegistros, normalizaUasg, normalizaNumero,
};
