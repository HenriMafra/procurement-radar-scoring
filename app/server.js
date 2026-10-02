
function isMensagemRecente(textoOuData, maxDias = 2) {
  if (!textoOuData) return true;
  const m = /(\d{1,2})\/(\d{1,2})\/(\d{4})/.exec(textoOuData);
  if (m) {
    const dia = parseInt(m[1], 10);
    const mes = parseInt(m[2], 10) - 1;
    const ano = parseInt(m[3], 10);
    const dataMsg = new Date(ano, mes, dia);
    const hoje = new Date();
    hoje.setHours(0, 0, 0, 0);
    const diffMs = hoje.getTime() - dataMsg.getTime();
    const diffDias = diffMs / (1000 * 60 * 60 * 24);
    if (diffDias > maxDias) return false;
  }
  return true;
}


const EXTENSION_AUTH_TOKEN = (() => {
  try {
    const authPath = path.join(__dirname, '..', 'data', 'extension-auth.json');
    if (fs.existsSync(authPath)) {
      return JSON.parse(fs.readFileSync(authPath, 'utf8')).token;
    }
  } catch (_) {}
  return 'enterprisecore_sentinela_9a7d3f5b8e1c2a4d6f0b8e9a7c5d3b1f';
})();

const { execFile } = require('child_process');
// =====================================================================
//  ENTERPRISECORE Radar - servidor local (Node puro, sem dependencias)
//  - Serve a interface (ui.html)
//  - Entrega os dados (dados.json real ou exemplo.json)
//  - Entrega o catalogo de contatos
//  - Abre o Outlook via PowerShell/COM com o e-mail pronto
//  - Registra envios (log) para marcar "JA ENVIADO"
//  - Permite processar e baixar/abrir a planilha Excel de 18 colunas
// =====================================================================
'use strict';

const http = require('http');
const fs = require('fs');
const os = require('os');
const path = require('path');
const { spawn, exec } = require('child_process');
const comprasApi = require('./compras-api');

const ROOT = path.resolve(__dirname, '..');
// PORT e definido mais abaixo, a partir de config/ambiente.json

const P = {
  ui:          path.join(__dirname, 'ui.html'),
  dados:       path.join(ROOT, 'data', 'dados.json'),      // gerado na Fase 2
  exemplo:     path.join(ROOT, 'data', 'exemplo.json'),    // Fase 1
  contatos:    path.join(ROOT, 'config', 'contatos.json'),
  log:         path.join(ROOT, 'data', 'enviados-log.json'),
  boletins:    path.join(ROOT, 'data', 'boletins'),        // onde o usuario solta os .xlsx do dia
  xlsx:        path.join(ROOT, 'data', 'Planilha-ENTERPRISECORE-Licitacoes.xlsx'),
  psEnviar:    path.join(ROOT, 'scripts', 'enviar-outlook.ps1'),
  psAlertaAuto:path.join(ROOT, 'scripts', 'enviar-alerta-auto.ps1'),
  pyEnviar:    path.join(ROOT, 'scripts', 'enviar_email.py'),
  psProcessar: path.join(ROOT, 'scripts', 'processar-boletins.ps1'),
  psOft:       path.join(ROOT, 'scripts', 'gerar-oft-lote.ps1'),
  acompanhados:path.join(ROOT, 'data', 'acompanhados.json'),
  historico:   path.join(ROOT, 'data', 'historico-eventos.json'),
  escopo:      path.join(ROOT, 'config', 'escopo-enterprisecore.json'),
  logoB64:     path.join(__dirname, 'logo_b64.txt'),
  logoJpg:     path.join(__dirname, 'enterprisecore_logo.jpg'),
  logoCompacto:path.join(__dirname, 'enterprisecore_logo_grupo_compact.png'),
  pdf:         path.join(ROOT, 'ENTERPRISECORE_Manual_Completo.pdf'),
  manualUsoBoletins:        path.join(ROOT, 'docs', 'manual-uso-boletins_timbrado.pdf'),
  manualConstrucaoBoletins: path.join(ROOT, 'docs', 'manual-construcao-boletins_timbrado.pdf'),
  manualUsoRos:             path.join(ROOT, 'docs', 'manual-uso-ros_timbrado.pdf'),
  manualConstrucaoRos:      path.join(ROOT, 'docs', 'manual-construcao-ros_timbrado.pdf'),
  manualUsoPregoes:         path.join(ROOT, 'docs', 'manual-uso-pregoes_timbrado.pdf'),
  manualConstrucaoPregoes:  path.join(ROOT, 'docs', 'manual-construcao-pregoes_timbrado.pdf'),
};

// ---------- helpers ----------
function sendJson(res, code, obj) {
  const body = JSON.stringify(obj);
  res.writeHead(code, {
    'Content-Type': 'application/json; charset=utf-8',
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization',
  });
  res.end(body);
}

function readBody(req) {
  return new Promise((resolve, reject) => {
    let data = '';
    req.on('data', (c) => { data += c; if (data.length > 20 * 1024 * 1024) reject(new Error('payload grande demais')); });
    req.on('end', () => resolve(data));
    req.on('error', reject);
  });
}

function lerJson(file, fallback) {
  try {
    if (!fs.existsSync(file)) return fallback;
    let raw = fs.readFileSync(file, 'utf8');
    // O Bloco de Notas e o PowerShell gravam UTF-8 COM BOM. O JSON.parse quebra
    // com o BOM na frente, e o erro passava batido: a pessoa editava o config,
    // salvava, e nada mudava. Remove o BOM antes de interpretar.
    if (raw.charCodeAt(0) === 0xFEFF) raw = raw.slice(1);
    return JSON.parse(raw);
  } catch (e) {
    console.error(`\n[ERRO] Nao consegui ler ${file}`);
    console.error(`       ${e.message}`);
    console.error('       O arquivo provavelmente tem erro de digitacao (virgula a mais, aspas faltando).');
    console.error('       Cole o conteudo em https://jsonlint.com para achar a linha.\n');
    return fallback;
  }
}

// =====================================================================
//  VARREDURA AUTOMATICA DA VIGILANCIA (a cada 5 min, em segundo plano)
//  Assim a pagina ja abre atualizada: o navegador so le o arquivo,
//  quem consulta as APIs e o servidor, sozinho.
// =====================================================================
const INTERVALO_MIN = 1;
let VARREDURA = { rodando: false, ultima: null, ultimoErro: '', comMudanca: 0, verificados: 0 };

function consultarChatPCP(codigoLicitacao) {
  return new Promise((resolve) => {
    if (!codigoLicitacao) return resolve([]);
    const url = `https://compras.api.portaldecompraspublicas.com.br/v1/licitacao/${codigoLicitacao}/chat`;
    const https = require('https');
    const req = https.get(url, {
      rejectUnauthorized: false,
      headers: {
        'User-Agent': 'ENTERPRISECORE-Radar/1.0',
        'Accept': 'application/json'
      }
    }, (res) => {
      let data = '';
      res.on('data', c => data += c);
      res.on('end', () => {
        if (res.statusCode !== 200) return resolve([]);
        try {
          const json = JSON.parse(data);
          const frases = Array.isArray(json.frasesChat) ? json.frasesChat : [];
          resolve(frases);
        } catch (e) {
          resolve([]);
        }
      });
    });
    req.on('error', () => resolve([]));
    req.setTimeout(10000, () => { req.destroy(); resolve([]); });
  });
}

function consultarItensPCP(codigoLicitacao) {
  return new Promise((resolve) => {
    if (!codigoLicitacao) return resolve(null);
    const url = `https://compras.api.portaldecompraspublicas.com.br/v2/licitacao/${codigoLicitacao}/itens?pagina=1`;
    const https = require('https');
    const req = https.get(url, {
      rejectUnauthorized: false,
      headers: {
        'User-Agent': 'ENTERPRISECORE-Radar/1.0',
        'Accept': 'application/json'
      }
    }, (res) => {
      let data = '';
      res.on('data', c => data += c);
      res.on('end', () => {
        if (res.statusCode !== 200) return resolve(null);
        try {
          resolve(JSON.parse(data));
        } catch (e) {
          resolve(null);
        }
      });
    });
    req.on('error', () => resolve(null));
    req.setTimeout(10000, () => { req.destroy(); resolve(null); });
  });
}

function consultarDocumentosPCP(codigoLicitacao) {
  return new Promise((resolve) => {
    if (!codigoLicitacao) return resolve([]);
    const url = `https://compras.api.portaldecompraspublicas.com.br/v2/licitacao/${codigoLicitacao}/documentos/processo`;
    const https = require('https');
    const req = https.get(url, {
      rejectUnauthorized: false,
      headers: {
        'User-Agent': 'ENTERPRISECORE-Radar/1.0',
        'Accept': 'application/json'
      }
    }, (res) => {
      let data = '';
      res.on('data', c => data += c);
      res.on('end', () => {
        if (res.statusCode !== 200) return resolve([]);
        try {
          const json = JSON.parse(data);
          resolve(Array.isArray(json) ? json : []);
        } catch (e) {
          resolve([]);
        }
      });
    });
    req.on('error', () => resolve([]));
    req.setTimeout(10000, () => { req.destroy(); resolve([]); });
  });
}

function consultarPCP(item) {
  return new Promise((resolve) => {
    let url = '';
    const link = (typeof item === 'object' ? item.registro?.linkPortal || item.linkPortal : String(item || ''));
    if (link && link.includes('portaldecompraspublicas.com.br/processos/')) {
      const pathSuffix = link.split('portaldecompraspublicas.com.br/processos/')[1];
      if (pathSuffix) {
        url = `https://compras.api.portaldecompraspublicas.com.br/v2/licitacao/${pathSuffix}`;
      }
    }

    if (!url) {
      const termo = typeof item === 'object' ? (item.numero || item.apelido || '') : String(item || '');
      url = `https://compras.api.portaldecompraspublicas.com.br/v2/licitacao/processos?objeto=${encodeURIComponent(termo)}`;
    }

    const https = require('https');
    const req = https.get(url, {
      rejectUnauthorized: false,
      headers: {
        'User-Agent': 'ENTERPRISECORE-Radar/1.0',
        'Accept': 'application/json'
      }
    }, (res) => {
      let data = '';
      res.on('data', c => data += c);
      res.on('end', () => {
        if (res.statusCode !== 200) return resolve({ achado: false, erro: `HTTP ${res.statusCode}` });
        try {
          const json = JSON.parse(data);
          if (json.codigoLicitacao) {
            const sitDesc = json.statusProcesso?.descricao || json.statusProcessoPublico?.descricao || json.status?.descricao || 'Publicada';
            const sitCod = json.statusProcesso?.codigo || json.statusProcessoPublico?.codigo || json.status?.codigo || 1;
            return resolve({
              achado: true,
              registro: {
                codigoLicitacao: json.codigoLicitacao,
                numeroCompra: json.numeroProcesso || json.numero,
                orgao: json.razaoSocialComprador || json.razaoSocial,
                situacao: sitDesc,
                statusCodigo: sitCod,
                autoridadeCompetente: json.nomeAutoridadeCompetente || '',
                pregoeiro: json.nomeOperador || json.nomePregoeiro || '',
                dataAbertura: json.dataHoraAbertura || json.dataHoraInicioPropostas,
                dataEncerramento: json.dataHoraFechamento || json.dataHoraInicioLances,
                dataPublicacao: json.dataHoraPublicacao,
                objeto: json.objeto || json.resumo,
                linkPortal: link
              }
            });
          }
          const items = json.result || json.items || (Array.isArray(json) ? json : []);
          if (!items.length) return resolve({ achado: false, erro: 'Nenhum processo encontrado no PCP' });
          const first = items[0];
          const sitDesc = first.statusProcesso?.descricao || first.status?.descricao || 'Publicada';
          const sitCod = first.statusProcesso?.codigo || first.status?.codigo || 1;
          resolve({
            achado: true,
            registro: {
              codigoLicitacao: first.codigoLicitacao,
              numeroCompra: first.numero,
              orgao: first.razaoSocial,
              situacao: sitDesc,
              statusCodigo: sitCod,
              autoridadeCompetente: '',
              pregoeiro: '',
              dataAbertura: first.dataHoraInicioPropostas,
              dataEncerramento: first.dataHoraInicioLances,
              dataPublicacao: first.dataHoraPublicacao,
              objeto: first.resumo,
              linkPortal: `https://www.portaldecompraspublicas.com.br/processos${first.urlReferencia || ''}`
            }
          });
        } catch (e) {
          resolve({ achado: false, erro: e.message });
        }
      });
    });
    req.on('error', e => resolve({ achado: false, erro: e.message }));
    req.setTimeout(12000, () => { req.destroy(); resolve({ achado: false, erro: 'timeout' }); });
  });
}

// Consulta um item e aplica o resultado nele. Devolve as mudancas detectadas.
async function atualizarItem(item, historico) {
  const agora = new Date().toISOString();
  item.ultimaChecagem = agora;

  // 1. Portal de Compras Públicas
  if (item.portal === 'compraspublicas') {
    const codLicitacao = item.codigoLicitacao || item.registro?.codigoLicitacao;
    const r = await consultarPCP(item);
    const mudancas = [];

    if (r.achado) {
      item.erro = '';

      // Consulta itens/lotes e documentos em tempo real
      let sitItens = '';
      let valorFinalVencedor = 0;
      let itensData = null;
      let docs = [];

      if (codLicitacao) {
        itensData = await consultarItensPCP(codLicitacao);
        if (itensData?.lotes?.result?.length) {
          const primeiroLote = itensData.lotes.result[0];
          const primeiroItem = primeiroLote.itens?.[0];
          if (primeiroItem?.situacao?.descricao) {
            sitItens = primeiroItem.situacao.descricao;
          }
          if (primeiroItem?.melhorLance) {
            valorFinalVencedor = Number(primeiroItem.melhorLance);
          }
        }
        docs = await consultarDocumentosPCP(codLicitacao);
      }

      const temTermoHomologacao = docs.some(d => (d.nome || d.tituloDocumento || '').includes('Homologação'));
      const temTermoAdjudicacao = docs.some(d => (d.nome || d.tituloDocumento || '').includes('Adjudicação'));

      // Situação consolidada do certame
      let situacaoFinal = r.registro.situacao;
      if (sitItens === 'Homologado' || temTermoHomologacao || (r.registro.statusCodigo === 7 && temTermoHomologacao)) {
        situacaoFinal = 'Homologado';
      } else if (sitItens === 'Adjudicado' || temTermoAdjudicacao) {
        situacaoFinal = 'Adjudicado';
      }

      const novaAssinatura = `situacao=${situacaoFinal}|status=${r.registro.statusCodigo}|itens=${sitItens}|termoHomolog=${temTermoHomologacao}|abertura=${r.registro.dataAbertura}|lances=${r.registro.dataEncerramento}`;
      
      if (item.assinatura && item.assinatura !== novaAssinatura) {
        const sitAnterior = item.registro?.situacaoProcessada || item.registro?.situacao;
        if (sitAnterior !== situacaoFinal) {
          if (situacaoFinal === 'Homologado') {
            const autoridade = r.registro.autoridadeCompetente ? ` por ${r.registro.autoridadeCompetente} (Autoridade Competente)` : '';
            const vlrStr = valorFinalVencedor > 0 ? ` — R$ ${valorFinalVencedor.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}` : '';
            mudancas.push(`🏆 HOMOLOGAÇÃO: Processo homologado oficialmente no Portal de Compras Públicas${autoridade}! Vencedor: ENTERPRISECORE SOLUÇÕES EM TELEINFORMÁTICA LTDA${vlrStr}`);
          } else if (situacaoFinal === 'Adjudicado') {
            mudancas.push(`⚖️ ADJUDICAÇÃO: Lote adjudicado no Portal de Compras Públicas!`);
          } else {
            mudancas.push(`Status no PCP alterado para: "${situacaoFinal}"`);
          }
        }
        if (item.registro?.dataEncerramento !== r.registro.dataEncerramento && r.registro.dataEncerramento) {
          mudancas.push(`Data da Sessão/Lances ajustada para: ${r.registro.dataEncerramento}`);
        }
      }

      item.assinatura = novaAssinatura;
      const codOriginal = item.codigoLicitacao || item.registro?.codigoLicitacao;
      item.registro = Object.assign({}, item.registro, r.registro, {
        situacao: situacaoFinal,
        situacaoOriginalPCP: r.registro.situacao,
        situacaoItens: sitItens,
        situacaoProcessada: situacaoFinal,
        valorHomologado: valorFinalVencedor || item.registro?.valorEstimado
      });
      if (codOriginal) {
        item.registro.codigoLicitacao = codOriginal;
        item.codigoLicitacao = codOriginal;
      }

      // Leitor de Documentos Oficiais Publicados
      if (docs.length) {
        if (!Array.isArray(item.documentosConhecidos)) {
          item.documentosConhecidos = docs.map(d => d.nome || d.tituloDocumento).filter(Boolean);
          // Se houver Termo de Homologação recém-descoberto que nunca foi notificado em eventos:
          const jaTemEventoHomolog = arr(item.eventos).some(e => e.texto && e.texto.includes('Termo de Homologação'));
          if (temTermoHomologacao && !jaTemEventoHomolog) {
            mudancas.push('📄 DOCUMENTO OFICIAL: Termo de Homologação publicado no Portal de Compras Públicas');
          }
        } else {
          const conhecidos = new Set(item.documentosConhecidos);
          for (const d of docs) {
            const nomeDoc = d.nome || d.tituloDocumento;
            if (nomeDoc && !conhecidos.has(nomeDoc)) {
              conhecidos.add(nomeDoc);
              mudancas.push(`📄 NOVO DOCUMENTO: "${nomeDoc}" publicado no portal`);
            }
          }
          item.documentosConhecidos = Array.from(conhecidos);
        }
      }
    } else if (!codLicitacao) {
      item.erro = r.erro || 'não localizado no PCP';
    }

    // Leitor de Chat / Mensagens em Tempo Real do Pregoeiro no PCP
    if (codLicitacao) {
      const frases = await consultarChatPCP(codLicitacao);
      if (frases.length) {
        if (!Array.isArray(item.ultimasFrasesConhecidas)) {
          item.ultimasFrasesConhecidas = frases.map(f => `${f.dataHoraFrase || ''}|${f.apelido || ''}|${f.frase || ''}`);
          const ultimas = frases.slice(-3);
          for (const uf of ultimas) {
            const ev = { quando: uf.dataHoraFrase || agora, texto: `[CHAT ${uf.apelido || 'PREGOEIRO'}]: ${uf.frase}`, autor: uf.apelido || 'PREGOEIRO', novo: false };
            item.eventos = arr(item.eventos).concat([ev]);
          }
        } else {
          const conhecidas = new Set(item.ultimasFrasesConhecidas);
          const novasFrases = [];
          for (const f of frases) {
            const chave = `${f.dataHoraFrase || ''}|${f.apelido || ''}|${f.frase || ''}`;
            if (!conhecidas.has(chave)) {
              conhecidas.add(chave);
              novasFrases.push(f);
            }
          }
          item.ultimasFrasesConhecidas = Array.from(conhecidas);
          for (const nf of novasFrases) {
            const txt = `[CHAT ${nf.apelido || 'PREGOEIRO'}]: ${nf.frase}`;
            mudancas.push(txt);
            const ev = { quando: nf.dataHoraFrase || agora, texto: txt, autor: nf.apelido || 'PREGOEIRO', novo: true };
            item.eventos = arr(item.eventos).concat([ev]);
            historico[item.id] = arr(historico[item.id]).concat([ev]);
          }
        }
      }
    }

    if (mudancas.length) {
      item.mudouAgora = true;
      const mudancasStatus = mudancas.filter(m => !m.startsWith('[CHAT'));
      if (mudancasStatus.length) {
        const evs = mudancasStatus.map(m => ({ quando: agora, texto: m, autor: 'PORTAL COMPRAS PÚBLICAS', novo: true }));
        item.eventos = arr(item.eventos).concat(evs);
        historico[item.id] = arr(historico[item.id]).concat(evs);
      }
    } else {
      item.mudouAgora = false;
    }

    return mudancas;
  }

  // 2. Licitações-e (manual / link direto seguro)
  if (item.portal === 'licitacoes-e' || item.semConsultaAutomatica) {
    item.mudouAgora = false;
    return [];
  }

  // 3. Compras.gov.br (Serpro / PNCP Oficial)
  const r = await comprasApi.consultarTudo({
    uasg: item.uasg,
    numero: item.numero,
    ano: item.ano,
    numeroControlePNCP: item.registro?.numeroControlePNCP || item.numeroControlePNCP,
    registroAtual: item.registro,
  });
  if (!r.achado) {
    item.erro = r.erro || 'não foi possível consultar';
    item.mudouAgora = false;
    return [];
  }

  item.erro = '';
  const novaAssinatura = comprasApi.assinatura(r.registro);
  const mudancas = comprasApi.compararRegistros(item.registro, r.registro);

  // --- OPÇÃO 1: Análise Automática de Documentos e Atas no PNCP ---
  if (Array.isArray(r.registro?.arquivos) && r.registro.arquivos.length) {
    if (!item.arquivosInicializados) {
      item.arquivosAnalisados = (r.registro.arquivos || []).map((a) => a.url).filter(Boolean);
      item.arquivosInicializados = true;
    } else {
      item.arquivosAnalisados = item.arquivosAnalisados || [];
      for (const arq of r.registro.arquivos) {
        if (arq.url && !item.arquivosAnalisados.includes(arq.url)) {
          item.arquivosAnalisados.push(arq.url);
          try {
            const resDoc = await executarPython('analisador-documentos.py', [arq.url, arq.titulo || '']);
            if (resDoc.ok && resDoc.dados) {
              if (Array.isArray(resDoc.dados.decisoes) && resDoc.dados.decisoes.length) {
                for (const dec of resDoc.dados.decisoes) {
                  mudancas.push(`[PNCP / ATO DOCUMENTAL]: ${dec.termo.toUpperCase()} em "${dec.arquivo}": ${dec.trecho}`);
                }
              }
            }
          } catch (_) {}
        }
      }
    }
  }

  // --- OPÇÃO 4: Monitoramento do Diário Oficial da União (DOU) ---
  const procRadical = (item.processo || item.registro?.processo || (item.apelido?.includes('CNMP') ? '19.00.6300.0000129' : '')).split('/')[0];
  if (procRadical && procRadical.length > 5) {
    try {
      const resDou = await executarPython('monitor-dou.py', [procRadical]);
      if (resDou.ok && Array.isArray(resDou.dados)) {
        if (!item.douInicializado) {
          item.douConhecidos = resDou.dados.map((m) => m.url).filter(Boolean);
          item.douInicializado = true;
        } else {
          item.douConhecidos = item.douConhecidos || [];
          for (const mat of resDou.dados) {
            if (mat.url && !item.douConhecidos.includes(mat.url)) {
              item.douConhecidos.push(mat.url);
              mudancas.push(`[DOU - DIÁRIO OFICIAL DA UNIÃO]: ${mat.titulo} (${mat.data} - ${mat.secao}): ${mat.resumo} (${mat.url})`);
            }
          }
        }
      }
    } catch (_) {}
  }

  if (novaAssinatura !== item.assinatura && mudancas.length) {
    item.mudouAgora = true;
    item.vistoEm = '';
    const eventos = mudancas.map((m) => ({ quando: agora, texto: m, autor: 'COMPRAS.GOV / PNCP', novo: true }));
    item.eventos = arr(item.eventos).concat(eventos);
    historico[item.id] = arr(historico[item.id]).concat(eventos);
  } else {
    item.mudouAgora = false;
  }
  item.assinatura = novaAssinatura;
  item.registro = r.registro;
  return mudancas;
}

async function dispararAlertaEmail(item, mudancas, opcoes = {}) {
  try {
    const isTeste = !!opcoes.isTeste;
    // REGRA ESTRITA DO USUÁRIO: Testes vão EXCLUSIVAMENTE para henri.mafra com prefixo 'TESTE - '
    const dests = isTeste 
      ? 'henri.mafra@grupoenterprisecore.com' 
      : ((Array.isArray(item.responsaveis) && item.responsaveis.length)
          ? item.responsaveis.join('; ')
          : 'henri.mafra@grupoenterprisecore.com');

    const apelidoOuOrgao = item.apelido || item.registro?.orgao || item.numero;
    const assunto = isTeste
      ? `TESTE - [B2G-RADAR-ENTERPRISECORE] ${opcoes.motivoTeste || 'Validação de Alerta e Layout'}: ${apelidoOuOrgao}`
      : `[B2G-RADAR-ENTERPRISECORE] 🚨 Nova movimentação: ${apelidoOuOrgao}`;

    const mudancasHtml = mudancas.map((m) => `<li style="margin-bottom:8px;line-height:1.5;color:#111827"><b>${esc(m)}</b></li>`).join('');

    const logo = logoDataUri();
    const logoImg = logo ? `<img src="${logo}" alt="EnterpriseCore CyberSecurity" style="height:38px;display:block;border:0;" />` : `<strong style="font-size:18px;color:#0f172a">ENTERPRISECORE</strong>`;

    const html = `
      <!DOCTYPE html>
      <html>
      <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
      </head>
      <body style="margin:0;padding:0;background:#f1f5f9;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Helvetica,Arial,sans-serif;">
        <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="background:#f1f5f9;padding:24px 12px;">
          <tr>
            <td align="center">
              <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="680" style="max-width:680px;background:#ffffff;border:1px solid #cbd5e1;border-radius:8px;overflow:hidden;box-shadow:0 4px 14px rgba(0,0,0,0.06);">
                
                <!-- Cabeçalho Limpo / Corporativo Fundo Branco -->
                <tr>
                  <td style="padding:18px 24px;background:#ffffff;border-bottom:2px solid #e2e8f0;">
                    <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%">
                      <tr>
                        <td style="vertical-align:middle;">
                          ${logoImg}
                        </td>
                        <td align="right" style="vertical-align:middle;">
                          <div style="font-size:13px;font-weight:800;color:#0f172a;letter-spacing:-0.01em;text-transform:uppercase;">ENTERPRISECORE RADAR &bull; VIGILÂNCIA 24/7</div>
                          <div style="font-size:11px;color:#64748b;margin-top:2px;">Chave de Regra: <span style="font-family:monospace;font-weight:700;">[B2G-RADAR-ENTERPRISECORE]</span></div>
                        </td>
                      </tr>
                    </table>
                  </td>
                </tr>

                <!-- Conteúdo Principal -->
                <tr>
                  <td style="padding:24px 28px;background:#ffffff;">
                    
                    <!-- Destaque Vermelho do Alerta -->
                    <div style="background:#fff1f2;border-left:5px solid #e11d48;padding:16px 18px;border-radius:4px;margin-bottom:22px;">
                      <div style="color:#be123c;font-size:13px;font-weight:900;text-transform:uppercase;letter-spacing:0.5px;">
                        🚨 NOVA MOVIMENTAÇÃO / AVISO OFICIAL:
                      </div>
                      <ul style="margin:10px 0 0 0;padding-left:20px;font-size:14px;color:#1e293b;">
                        ${mudancasHtml}
                      </ul>
                    </div>

                    <!-- Dados Cadastrais da Oportunidade -->
                    <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="border-collapse:collapse;margin-bottom:20px;font-size:13.5px;">
                      <tr style="border-bottom:1px solid #f1f5f9;">
                        <td style="padding:9px 0;color:#64748b;width:140px;"><b>Órgão:</b></td>
                        <td style="padding:9px 0;color:#0f172a;font-weight:700;">${item.registro?.orgao || item.apelido || '—'}</td>
                      </tr>
                      <tr style="border-bottom:1px solid #f1f5f9;">
                        <td style="padding:9px 0;color:#64748b;"><b>Pregão / Processo:</b></td>
                        <td style="padding:9px 0;color:#0f172a;"><b>${(item.numero && item.ano && !String(item.numero).includes(item.ano)) ? item.numero + "/" + item.ano : (item.numero || "")}</b> (${item.uasg ? 'UASG ' + item.uasg : (item.portal === 'compraspublicas' ? 'Compras Públicas' : 'Licitações-e')})</td>
                      </tr>
                      <tr style="border-bottom:1px solid #f1f5f9;">
                        <td style="padding:9px 0;color:#64748b;"><b>Portal:</b></td>
                        <td style="padding:9px 0;color:#0f172a;">${item.portal === 'licitacoes-e' ? 'Licitações-e (Banco do Brasil)' : (item.portal === 'compraspublicas' ? 'Portal de Compras Públicas' : 'Compras.gov.br (Serpro / PNCP)')}</td>
                      </tr>
                      <tr style="border-bottom:1px solid #f1f5f9;">
                        <td style="padding:9px 0;color:#64748b;"><b>Valor Estimado:</b></td>
                        <td style="padding:9px 0;color:#0f172a;font-weight:800;">${Number(item.registro?.valorEstimado) > 0 ? 'R$ ' + Number(item.registro?.valorEstimado).toLocaleString('pt-BR') : 'Sigiloso'}</td>
                      </tr>
                      <tr style="border-bottom:1px solid #f1f5f9;">
                        <td style="padding:9px 0;color:#64748b;"><b>Responsáveis:</b></td>
                        <td style="padding:9px 0;color:#0f172a;"><span style="background:#f8fafc;border:1px solid #e2e8f0;padding:3px 8px;border-radius:4px;font-family:monospace;font-size:12px;">${dests}</span></td>
                      </tr>
                    </table>

                    ${item.registro?.objeto ? `
                      <div style="background:#f8fafc;border:1px solid #e2e8f0;padding:12px 14px;border-radius:6px;font-size:12.5px;color:#334155;margin-bottom:22px;line-height:1.55;">
                        <b style="color:#0f172a;">Objeto:</b> ${item.registro.objeto}
                      </div>` : ''}

                    <!-- Botão de Acesso Oficial -->
                    <div style="margin-top:24px;padding-top:16px;border-top:1px solid #f1f5f9;text-align:center;">
                      ${item.registro?.linkPortal ? `<a href="${item.registro.linkPortal}" target="_blank" style="display:inline-block;background:#0f172a;color:#ffffff;padding:12px 24px;text-decoration:none;border-radius:6px;font-weight:700;font-size:13.5px;box-shadow:0 2px 6px rgba(0,0,0,0.12);">Acessar Sessão do Pregão no Portal Oficial &rarr;</a>` : ''}
                    </div>

                  </td>
                </tr>

                <!-- Rodapé Corporativo Claro -->
                <tr>
                  <td style="padding:14px 24px;background:#f8fafc;border-top:1px solid #e2e8f0;font-size:11.5px;color:#64748b;text-align:center;">
                    ENTERPRISECORE Radar &bull; Central de Vigilância Estratégica 24/7 &bull; Enterprise IT Group
                  </td>
                </tr>

              </table>
            </td>
          </tr>
        </table>
      </body>
      </html>
    `;

    const tmp = path.join(os.tmpdir(), `enterprisecore-alert-${Date.now()}-${Math.floor(Math.random()*1000)}.json`);
    fs.writeFileSync(tmp, JSON.stringify({ to: dests, subject: assunto, html }), 'utf8');

    // Marca no item que o alerta foi despachado
    const agora = new Date().toISOString();
    item.ultimoAlertaEnviado = agora;
    if (Array.isArray(item.eventos)) {
      item.eventos.forEach(e => {
        if (mudancas.some(m => e.texto && e.texto.includes(m))) {
          e.alertaEnviado = true;
        }
      });
    }

    return new Promise((resolve) => {
      if (process.platform === 'win32' && fs.existsSync(P.psAlertaAuto)) {
        const ps = spawn('powershell.exe', [
          '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass',
          '-File', P.psAlertaAuto, '-JsonPath', tmp
        ], { windowsHide: true });
        ps.stdout.on('data', d => console.log(`[alerta-ps] ${d.toString().trim()}`));
        ps.stderr.on('data', d => console.error(`[alerta-ps-err] ${d.toString().trim()}`));
        ps.on('close', (code) => {
          if (code === 0) {
            console.log(`[alerta] ✅ E-mail despachado via Outlook COM para: ${dests} (${assunto})`);
            try { fs.unlinkSync(tmp); } catch (_) {}
            resolve(true);
          } else {
            console.error(`[alerta] ⚠️ PowerShell código ${code}. Tentando fallback via Python...`);
            const pyExec = process.platform === 'win32' ? 'python' : 'python3';
            const py = spawn(pyExec, [P.pyEnviar, tmp]);
            py.on('close', () => {
              try { fs.unlinkSync(tmp); } catch (_) {}
              resolve(false);
            });
          }
        });
      } else {
        const pyExec = process.platform === 'win32' ? 'python' : 'python3';
        const py = spawn(pyExec, [P.pyEnviar, tmp]);
        py.on('close', () => {
          try { fs.unlinkSync(tmp); } catch (_) {}
          resolve(true);
        });
      }
    });
  } catch (e) {
    console.error('[alerta] Falha ao despachar e-mail:', e.message);
  }
}

async function varrerVigilancia(motivo) {
  if (VARREDURA.rodando) return;
  const lista = arr(lerJson(P.acompanhados, []));
  if (!lista.length) { VARREDURA.ultima = new Date().toISOString(); return; }

  VARREDURA.rodando = true;
  const historico = lerJson(P.historico, {});
  let comMudanca = 0;
  try {
    for (const item of lista) {
      const m = await atualizarItem(item, historico);
      if (m.length) {
        comMudanca++;
        await dispararAlertaEmail(item, m);
      } else {
        // Checa se há eventos novos pendentes de alerta que nunca foram enviados por e-mail
        const pendentes = arr(item.eventos).filter(e => e.novo && !e.alertaEnviado).map(e => e.texto);
        if (pendentes.length) {
          await dispararAlertaEmail(item, pendentes);
        }
      }
      await new Promise((s) => setTimeout(s, 400));
    }
    fs.writeFileSync(P.acompanhados, JSON.stringify(lista, null, 2), 'utf8');
    fs.writeFileSync(P.historico, JSON.stringify(historico, null, 2), 'utf8');
    VARREDURA.ultimoErro = '';
  } catch (e) {
    VARREDURA.ultimoErro = e.message;
  }
  VARREDURA.rodando = false;
  VARREDURA.ultima = new Date().toISOString();
  VARREDURA.comMudanca = comMudanca;
  VARREDURA.verificados = lista.length;
  console.log(`[vigilancia] ${motivo}: ${lista.length} verificado(s), ${comMudanca} com movimentacao`);
}

// =====================================================================
//  PORTABILIDADE - nada de caminho fixo de maquina aqui.
//  Tudo que muda de PC vem de config/ambiente.json; quando o campo esta
//  vazio, o Radar procura sozinho nos locais mais comuns.
// =====================================================================
const HOME = process.env.USERPROFILE || require('os').homedir();

function existe(p) { try { return !!p && fs.existsSync(p); } catch (_) { return false; } }
function primeiroQueExiste(lista) { return lista.find(existe) || ''; }

// O python.exe da WindowsApps e um atalho de 0 byte da Microsoft Store:
// executa nada e sai com codigo 9009. Tem que ser descartado.
function pythonValido(p) {
  if (!existe(p)) return false;
  if (/WindowsApps/i.test(p)) return false;
  try { return fs.statSync(p).size > 0; } catch (_) { return false; }
}

function acharPython() {
  const candidatos = [
    path.join(HOME, 'Python', 'python.exe'),
    path.join(HOME, 'AppData', 'Local', 'Programs', 'Python', 'Python313', 'python.exe'),
    path.join(HOME, 'AppData', 'Local', 'Programs', 'Python', 'Python312', 'python.exe'),
    path.join(HOME, 'AppData', 'Local', 'Programs', 'Python', 'Python311', 'python.exe'),
    'C:\\Python313\\python.exe', 'C:\\Python312\\python.exe', 'C:\\Python311\\python.exe',
    'C:\\Program Files\\Python313\\python.exe', 'C:\\Program Files\\Python312\\python.exe',
  ];
  return candidatos.find(pythonValido) || '';
}

const AMB = lerJson(path.join(ROOT, 'config', 'ambiente.json'), {});
const painelCfg = AMB.painelROs || {};

const PORT = Number(AMB.porta) > 0 ? Number(AMB.porta) : 8790;

const PAINEL = {
  pasta: painelCfg.pasta || primeiroQueExiste([
    path.join(HOME, 'Gestao_ROs_Bitrix'),
    path.join(HOME, 'Desktop', 'Gestao_ROs_Bitrix'),
    path.join(HOME, 'Documents', 'Gestao_ROs_Bitrix'),
  ]),
  python: pythonValido(AMB.python) ? AMB.python : acharPython(),
  porta: Number(painelCfg.porta) > 0 ? Number(painelCfg.porta) : 8501,
  proc: null,
  log: '',
};
PAINEL.url = `http://localhost:${PAINEL.porta}`;

// scripts.json do Painel de Scripts (opcional)
const PATH_SCRIPTS = AMB.painelScripts || primeiroQueExiste([
  path.join(HOME, 'Desktop', 'PAINEL SCRIPTS', 'scripts.json'),
  path.join(HOME, 'PAINEL SCRIPTS', 'scripts.json'),
]);

// Detecta o painel no ar mesmo se ele tiver sido aberto pelo .bat, fora do Radar
function painelNoAr() {
  return new Promise((resolve) => {
    const req = require('http').get({ host: 'localhost', port: PAINEL.porta, path: '/', timeout: 1500 }, (r) => {
      r.resume();
      resolve(r.statusCode > 0);
    });
    req.on('error', () => resolve(false));
    req.on('timeout', () => { req.destroy(); resolve(false); });
  });
}

function esc(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}
function arr(x) { return Array.isArray(x) ? x : (x ? [x] : []); }
function dataBr(iso) {
  const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(String(iso || ''));
  return m ? `${m[3]}/${m[2]}/${m[1]}` : '';
}

// ---------- composicao do e-mail (usada pelo .OFT e pelo envio direto) ----------
function logoDataUri() {
  try { if (fs.existsSync(P.logoB64)) return fs.readFileSync(P.logoB64, 'utf8').trim(); } catch (_) {}
  return '';
}

let _logoCompactoCache = null;
function logoCompactoDataUri() {
  if (_logoCompactoCache !== null) return _logoCompactoCache;
  try {
    const buf = fs.readFileSync(P.logoCompacto);
    _logoCompactoCache = `data:image/png;base64,${buf.toString('base64')}`;
  } catch (_) { _logoCompactoCache = ''; }
  return _logoCompactoCache;
}

// Mesma logica do siglaUf() do ui.html: "Santa Catarina (SC)" -> "SC".
function siglaUf(estado) {
  const m = /\(([A-Z]{2})\)\s*$/.exec(String(estado || '').trim());
  return m ? m[1] : String(estado || '').trim().toUpperCase();
}

// Corta o objeto num limite legivel, sem quebrar palavra no meio.
// Os boletins trazem lotes inteiros num paragrafo so - despejar tudo no e-mail
// vira parede de texto e ninguem le.
function resumirObjeto(txt, limite = 620) {
  const t = String(txt || '').replace(/\s+/g, ' ').trim();
  if (t.length <= limite) return { texto: t, cortado: false };
  const corte = t.lastIndexOf(' ', limite);
  return { texto: t.slice(0, corte > 0 ? corte : limite) + '…', cortado: true };
}

// A coluna UASG do boletim vem poluida com numero de processo na maioria das
// linhas (ex.: "23305.013201.2026-95"). Rotular isso como UASG num e-mail para
// um colega e passar informacao errada - so chama de UASG o que parece UASG.
function rotuloUasg(v) {
  const s = String(v == null ? '' : v).trim();
  if (!s) return '';
  if (/^\d{1,6}$/.test(s)) return 'UASG ' + esc(s);
  return 'processo ' + esc(s);
}

function prazoTexto(x) {
  const d = x.diasAtePregao;
  if (d === null || d === undefined || d === '') return '';
  const n = Number(d);
  if (isNaN(n)) return '';
  if (n < 0) return `venceu há ${Math.abs(n)} dia${Math.abs(n) === 1 ? '' : 's'}`;
  if (n === 0) return 'é HOJE';
  if (n === 1) return 'é AMANHÃ';
  return `faltam ${n} dias`;
}

// Traduz os sinais tecnicos em frase de gente. O campo motivoMatch e a
// decomposicao de pontos do motor: serve para auditar, nao para mandar por e-mail.
function porQueInteressa(x) {
  const p = [];
  const parc = arr(x.parceiros);
  const conc = arr(x.concorrentesDet);
  const sol = arr(x.solucoesDet);
  if (parc.length) p.push(`cita <b>${esc(parc.join(', '))}</b>, de quem somos parceiros`);
  if (conc.length) p.push(`cita <b>${esc(conc.join(', '))}</b> — oportunidade de substituição`);
  if (sol.length) p.push(`escopo bate com <b>${esc(sol.join(', '))}</b>`);
  if (!p.length) return '';
  let frase = p.join('; ') + '.';
  if (Number(x.densidadeFator) < 1) {
    frase += ' <i>Atenção: o sinal aparece pouco num objeto longo — confirmar se o item é central ou acessório no lote.</i>';
  }
  return frase.charAt(0).toUpperCase() + frase.slice(1);
}

// Data + hora como no modelo da equipe: "12/08/2026 | 14:00"
function dataHoraPregao(x) {
  const d = dataBr(x.dataPregao);
  if (!d) return '—';
  let h = String(x.horaPregao || '').trim();
  if (!h) h = '00:00';
  const m = /^(\d{1,2})[:h]?(\d{2})?/.exec(h);
  if (m) h = String(m[1]).padStart(2, '0') + ':' + (m[2] || '00');
  return `${d} | ${h}`;
}

// A "Solucao" do modelo e UM valor so. Nao dá para juntar as detectadas com "/"
// porque o proprio nome das categorias ja tem barra ("NGFW / Firewall / IPS"),
// e o resultado vira uma linha ilegivel. Mostra a principal e sinaliza as demais.
function solucaoTexto(x) {
  const s = arr(x.solucoesDet).filter(Boolean);
  if (s.length === 1) return s[0];
  if (s.length > 1) return `${s[0]}  (+${s.length - 1} ${s.length === 2 ? 'outra' : 'outras'})`;
  if (x.solucao && !/^(a verificar|perifericos)/i.test(x.solucao)) return x.solucao;
  return 'Outros';
}

// Um bloco por publicacao, no formato do comunicado usado pela equipe:
// numero, ficha de dados (Orgao, Solucao, No Pregao, UASG, Data) e o Objeto inteiro.
// Sem score e sem faixa de proposito: este e-mail vai para diretoria, nao e a tela
// de triagem. O que ordena a lista continua sendo o score, mas ele nao aparece aqui.
function blocoLicitacaoHtml(x, pos) {
  const campo = (rot, val) => `
    <tr>
      <td width="130" style="padding:5px 14px 5px 0;color:#6B665C;font-size:12px;font-weight:bold;vertical-align:top;white-space:nowrap">${esc(rot)}</td>
      <td style="padding:5px 0;color:#1C1A17;font-size:13px;line-height:1.45;vertical-align:top">${val}</td>
    </tr>`;

  return `
  <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="margin:0 0 26px 0">
    <tr><td style="padding-bottom:8px">
      <span style="display:inline-block;background:#2A2620;color:#FCFAF2;font-size:13px;font-weight:bold;padding:3px 11px">${pos}</span>
    </td></tr>
    <tr><td style="border:1px solid #D6CDB8;background:#FCFAF2;padding:14px 18px">

      <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%">
        ${campo('Órgão', esc(x.orgao || '—'))}
        ${campo('Solução', esc(solucaoTexto(x)))}
        ${campo('Nº Pregão', esc(x.numeroPregao || '—'))}
        ${campo('UASG', uasgTexto(x.uasg))}
        ${campo('Data do Pregão', `<b>${esc(dataHoraPregao(x))}</b>`)}
        ${x.estado ? campo('Estado', esc(x.estado)) : ''}
        ${x.sitioCompras ? campo('Portal', esc(x.sitioCompras)) : ''}
      </table>

      <div style="height:14px;line-height:14px">&nbsp;</div>

      <div style="font-size:12px;font-weight:bold;color:#6B665C;padding-bottom:5px">Objeto:</div>
      <div style="font-size:13px;color:#1C1A17;line-height:1.6;text-align:justify">${esc(String(x.objeto || '').replace(/\s+/g, ' ').trim())}</div>

      ${x.link ? `
      <div style="height:14px;line-height:14px">&nbsp;</div>
      <a href="${esc(x.link)}" style="color:#22506E;font-size:12.5px;font-weight:bold;text-decoration:none">→ Abrir edital</a>` : ''}

    </td></tr>
  </table>`;
}

// No modelo da equipe, UASG ausente aparece como travessao.
// A coluna do boletim vem poluida com numero de processo em quase toda linha,
// entao so mostra o que realmente parece UASG.
function uasgTexto(v) {
  const s = String(v == null ? '' : v).trim();
  if (/^\d{1,6}$/.test(s)) return esc(s);
  return '—';
}

function emailHtml(itens, titulo, opcoes) {
  const cfgEmail = (lerJson(P.escopo, {}).email) || {};
  const o = opcoes || {};
  const logo = logoCompactoDataUri();
  const cab = logo
    ? `<img src="${esc(logo)}" alt="Enterprise IT Group" style="height:34px;display:block;border:0" />`
    : `<div style="font-size:22px;font-weight:bold;color:#1C1A17">ENTERPRISECORE</div>`;

  const tit = titulo || cfgEmail.titulo || '📋 Radar de Publicações';
  const regiao = o.regiao || cfgEmail.regiao || '';
  const nota = cfgEmail.rodapeNota || 'Como as publicações podem ocorrer durante todo o dia, o corte será realizado em D-1.';
  const assin = cfgEmail.assinatura || 'Comunicado interno · Enterprise IT Group';
  const hoje = new Date().toLocaleDateString('pt-BR');
  const n = itens.length;

  return `<!DOCTYPE html>
<html><head><meta charset="utf-8"><title>${esc(tit)}</title></head>
<body style="margin:0;padding:0;background:#F6F1E4">
  <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="background:#F6F1E4">
    <tr><td align="center" style="padding:26px 12px">
      <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="720" style="max-width:720px;background:#FCFAF2;border:1px solid #D6CDB8">

        <tr><td style="padding:20px 26px;border-bottom:2px solid #2A2620">${cab}</td></tr>

        <tr><td style="padding:26px 26px 0;font-family:Segoe UI,Arial,sans-serif">
          <div style="font-size:21px;font-weight:bold;color:#1C1A17;line-height:1.3">${esc(tit)}</div>
          <div style="font-size:13px;color:#6B665C;padding-top:6px">
            ${esc(hoje)}${regiao ? '  ·  ' + esc(regiao) : ''}
          </div>
        </td></tr>

        <tr><td align="center" style="padding:24px 26px 26px;font-family:Segoe UI,Arial,sans-serif">
          <table role="presentation" cellpadding="0" cellspacing="0" border="0">
            <tr><td style="border:1px solid #D6CDB8;background:#EFE7D3;padding:9px 22px;font-size:14px;color:#1C1A17">
              Total: <b>${n}</b> ${n === 1 ? 'Publicação' : 'Publicações'}
            </td></tr>
          </table>
        </td></tr>

        <tr><td style="padding:0 26px;font-family:Segoe UI,Arial,sans-serif">
          ${itens.map((x, i) => blocoLicitacaoHtml(x, i + 1)).join('\n')}
        </td></tr>

        <tr><td style="padding:6px 26px 26px;font-family:Segoe UI,Arial,sans-serif">
          <div style="font-size:12.5px;color:#4B4740;line-height:1.6">${esc(nota)}</div>
          <div style="height:16px;line-height:16px">&nbsp;</div>
          <div style="font-size:13px;color:#1C1A17">Atenciosamente,</div>
          <div style="height:20px;line-height:20px">&nbsp;</div>
          <div style="border-top:1px solid #D6CDB8;padding-top:12px;font-size:11.5px;color:#948C7C">
            ${esc(assin)}
          </div>
        </td></tr>

      </table>
    </td></tr>
  </table>
</body></html>`;
}

function assuntoDe(x) {
  const p = [];
  if (x.faixa) p.push(`[${x.faixa}]`);
  p.push(x.orgao || 'Licitacao');
  if (x.numeroPregao) p.push(`- ${x.numeroPregao}`);
  return p.join(' ').slice(0, 200);
}

// Assunto do alerta individual (pasta "Individuais" do .OFT em lote).
// Enxuto de proposito: so o que precisa pra achar/identificar o processo na
// caixa de entrada - o resto (alerta, UF, data) ja esta no corpo do e-mail.
function assuntoIndividual(x) {
  let orgao = String(x.orgao || 'Órgão não informado').replace(/\s+/g, ' ').trim();
  if (orgao.length > 60) orgao = orgao.slice(0, 59) + '…';
  return x.numeroPregao ? `${orgao} — Pregão ${x.numeroPregao}` : orgao;
}

// Alerta individual, uma licitacao por e-mail - modelo da equipe comercial
// (barra vermelha de alerta + cabecalho azul-marinho + ficha de dados em
// cartao com borda + objeto). Diferente do emailHtml() (digest de varias
// publicacoes para a diretoria): este vai para a equipe comercial pedir
// posicionamento imediato, entao nao leva o aviso de "uso interno...
// diretoria e liderancas executivas".
function emailHtmlIndividual(x) {
  const logo = logoCompactoDataUri();
  const logoImg = logo
    ? `<img src="${esc(logo)}" alt="Enterprise IT Group" height="32" style="display:block;border:0" />`
    : `<div style="font-size:18px;font-weight:800;color:#0f2942;letter-spacing:-0.5px">Enterprise IT Group</div>`;

  const FF = "font-family:'Segoe UI',Arial,sans-serif";
  const NAVY = '#0f2942';
  const AZUL_VIVO = '#a9d3fb';
  const VERMELHO = '#dc2626';
  const GOLD = '#FFD54A';
  const GOLD_ESCURO = '#E6B800';
  // Nota: outlook classico (motor do Word) so renderiza texto claro em fundo
  // escuro de forma confiavel as vezes - ja vimos "Equipe Comercial" branco
  // sair cinza-apagado num fundo azul-marinho solido. Por isso o cabecalho
  // usa fundo claro com texto escuro (seguro) e reserva cor forte pra
  // borda/destaque, nao pro texto corrido. Tambem evitamos border-radius e
  // overflow:hidden nas tabelas - nao renderizam no Outlook classico mesmo.

  const linha = (rot, val, destaque, ultima) => `
          <tr>
            <td width="34%" bgcolor="${destaque ? GOLD : '#f3f6fb'}" style="${FF};background-color:${destaque ? GOLD : '#f3f6fb'};padding:10px 14px;${ultima ? '' : 'border-bottom:1px solid ' + (destaque ? GOLD_ESCURO : '#e2e8f0') + ';'}font-weight:bold;color:${NAVY};vertical-align:top;font-size:12.5px">${esc(rot)}</td>
            <td width="66%" bgcolor="${destaque ? GOLD : '#ffffff'}" style="${FF};background-color:${destaque ? GOLD : '#ffffff'};padding:10px 14px;${ultima ? '' : 'border-bottom:1px solid ' + (destaque ? GOLD_ESCURO : '#e2e8f0') + ';'}color:#1e293b;line-height:1.4;font-size:13px${destaque ? ';font-weight:bold' : ''}">${val}</td>
          </tr>`;

  const tituloSecao = (texto) => `
        <table border="0" cellpadding="0" cellspacing="0" width="100%" style="margin:0 0 12px">
          <tr>
            <td width="10" style="padding:0"><table cellpadding="0" cellspacing="0"><tr><td width="9" height="9" bgcolor="${NAVY}" style="background-color:${NAVY};font-size:1px;line-height:1px">&nbsp;</td></tr></table></td>
            <td style="${FF};padding-left:8px;font-size:13.5px;font-weight:bold;color:${NAVY};letter-spacing:0.3px;border-bottom:2px solid ${GOLD}">${texto}</td>
          </tr>
        </table>`;

  return `<!DOCTYPE html>
<html lang="pt-BR"><head><meta charset="UTF-8"><title>Publicação Identificada - Enterprise IT Group</title></head>
<body style="margin:0;padding:22px 0;background-color:#e9edf2;${FF};-webkit-font-smoothing:antialiased">

  <table align="center" border="0" cellpadding="0" cellspacing="0" width="100%" style="max-width:560px;margin:0 auto;background-color:#ffffff;border:2px solid ${NAVY};border-collapse:collapse;mso-table-lspace:0pt;mso-table-rspace:0pt">

    <tr>
      <td bgcolor="${VERMELHO}" style="${FF};background-color:${VERMELHO};padding:12px 22px;color:#ffffff;font-size:14.5px;font-weight:bold;letter-spacing:0.2px;border-bottom:3px solid #991b1b">
        ⚠️ Publicação Identificada
      </td>
    </tr>

    <tr>
      <td bgcolor="${AZUL_VIVO}" style="background-color:${AZUL_VIVO};padding:18px 22px;border-bottom:3px solid ${NAVY}">
        <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse:collapse">
          <tr>
            <td valign="middle" style="${FF};color:${NAVY}">
              <div style="font-size:16px;font-weight:bold;margin-bottom:4px;color:${NAVY}">Equipe Comercial</div>
              <div style="font-size:13px;color:#1e293b;line-height:1.4">
                Esta oportunidade exige posicionamento <strong style="color:${VERMELHO}">imediato</strong>.
              </div>
            </td>
            <td align="right" valign="middle" style="${FF};padding-left:16px">${logoImg}</td>
          </tr>
        </table>
      </td>
    </tr>

    <tr>
      <td style="padding:26px 24px;background-color:#ffffff">

        ${tituloSecao('Dados da Licitação')}

        <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse:collapse;border:1px solid ${NAVY};margin-bottom:28px">
          ${linha('Órgão', esc(x.orgao || '—'))}
          ${linha('Solução', esc(solucaoTexto(x)))}
          ${linha('Nº do Pregão', esc(x.numeroPregao || '—'))}
          ${linha('UASG', uasgTexto(x.uasg))}
          ${linha('⏰ Data do Pregão', `<b>${esc(dataHoraPregao(x))}</b>`, true, true)}
        </table>

        ${tituloSecao('Objeto da Contratação')}

        <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse:collapse;margin-bottom:24px">
          <tr>
            <td bgcolor="#f3f6fb" style="${FF};background-color:#f3f6fb;border:1px solid #d7dee8;border-left:5px solid ${NAVY};padding:15px 17px;font-size:12.5px;color:#1e293b;line-height:1.6;text-align:justify">
              ${esc(String(x.objeto || '').replace(/\s+/g, ' ').trim())}
            </td>
          </tr>
        </table>

        <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-top:2px solid ${GOLD}">
          <tr><td style="${FF};font-size:12.5px;font-weight:bold;color:${NAVY};padding-top:12px">Equipe Comercial · Enterprise IT Group</td></tr>
        </table>

      </td>
    </tr>

  </table>

</body></html>`;
}

// Resolve os itens a partir dos ids enviados pela interface
function itensPorIds(ids) {
  const set = new Set(arr(ids).map(String));
  const todos = arr(lerJson(P.dados, { licitacoes: [] }).licitacoes);
  return todos.filter((x) => set.has(String(x.id)));
}

// ---------- servidor ----------
const server = http.createServer(async (req, res) => {
  try {
    const url = req.url.split('?')[0];

    if (req.method === 'OPTIONS') {
      res.writeHead(204, {
        'Access-Control-Allow-Origin': '*',
        'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
        'Access-Control-Allow-Headers': 'Content-Type, Authorization, X-EnterpriseCore-Radar-Token',
        'Access-Control-Allow-Private-Network': 'true',
      });
      return res.end();
    }

    // Interface
    if (req.method === 'GET' && (url === '/' || url === '/index.html')) {
      fs.readFile(P.ui, (err, buf) => {
        if (err) { res.writeHead(500); return res.end('ui.html nao encontrado'); }
        res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
        res.end(buf);
      });
      return;
    }

    // Logo ENTERPRISECORE (PNG / JPG)
    if (req.method === 'GET' && (url === '/logo_enterprisecore.png' || url === '/logo.png' || url === '/public/logo_enterprisecore.png')) {
      const pngPath = path.join(__dirname, 'logo_enterprisecore.png');
      if (fs.existsSync(pngPath)) {
        res.writeHead(200, { 'Content-Type': 'image/png' });
        return fs.createReadStream(pngPath).pipe(res);
      }
    }
    if (req.method === 'GET' && (url === '/enterprisecore_logo.jpg' || url === '/public/enterprisecore_logo.jpg')) {
      const logoPath = path.join(__dirname, 'enterprisecore_logo.jpg');
      if (fs.existsSync(logoPath)) {
        res.writeHead(200, { 'Content-Type': 'image/jpeg' });
        return fs.createReadStream(logoPath).pipe(res);
      }
    }

    // Dados das licitacoes (prioriza o real, cai no exemplo)
    if (req.method === 'GET' && url === '/api/dados') {
      const usaReal = fs.existsSync(P.dados);
      const obj = lerJson(usaReal ? P.dados : P.exemplo, { licitacoes: [] });
      return sendJson(res, 200, {
        fonte: usaReal ? 'real' : 'exemplo',
        geradoEm: obj.geradoEm || null,
        boletinsDir: P.boletins,
        licitacoes: Array.isArray(obj.licitacoes) ? obj.licitacoes : [],
      });
    }

    // Contatos
    if (req.method === 'GET' && url === '/api/contatos') {
      const obj = lerJson(P.contatos, { pessoas: [] });
      return sendJson(res, 200, { pessoas: Array.isArray(obj.pessoas) ? obj.pessoas : [] });
    }

    // IDs ja enviados
    if (req.method === 'GET' && url === '/api/enviados') {
      const log = lerJson(P.log, []);
      const ids = Array.isArray(log) ? [...new Set(log.map((x) => x.id))] : [];
      return sendJson(res, 200, { ids });
    }

    // Marca ids como enviados (append no log)
    if (req.method === 'POST' && url === '/api/marcar-enviado') {
      const body = JSON.parse((await readBody(req)) || '{}');
      const log = lerJson(P.log, []);
      const agora = new Date().toISOString();
      (body.ids || []).forEach((id) => log.push({ id, to: body.to || '', cc: body.cc || '', quando: agora }));
      fs.writeFileSync(P.log, JSON.stringify(log, null, 2), 'utf8');
      return sendJson(res, 200, { ok: true });
    }

    // Escopo ENTERPRISECORE completo - alimenta o dossie para analise por IA
    if (req.method === 'GET' && url === '/api/escopo') {
      const cfg = lerJson(P.escopo, {});
      return sendJson(res, 200, {
        parceiros: arr(cfg.fabricantes).map((f) => ({ nome: f.nome, solucao: f.solucao })),
        concorrentes: arr(cfg.concorrentes).map((c) => c.nome),
        solucoes: arr(cfg.categorias).map((c) => ({
          solucao: c.solucao, fabricantePadrao: c.fabricantePadrao, palavrasChave: arr(c.palavrasChave),
        })),
        faixas: arr(cfg.faixas),
        pesos: cfg.pesos || {},
        estadosPreferidos: cfg.estadosPreferidos || null,
      });
    }

    // Previa do e-mail: exatamente o HTML que vai para o .OFT
    if (req.method === 'POST' && url === '/api/previa-email') {
      const body = JSON.parse((await readBody(req)) || '{}');
      const itens = itensPorIds(body.ids);
      if (!itens.length) return sendJson(res, 400, { ok: false, erro: 'Nenhuma licitacao selecionada.' });
      res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
      return res.end(emailHtml(itens, body.assunto));
    }

    // Lista de scripts do Painel de Scripts no Desktop
    if (req.method === 'GET' && url === '/api/scripts-painel') {
      const list = lerJson(PATH_SCRIPTS, []);
      return sendJson(res, 200, { ok: true, scripts: list });
    }

    // Executa um script do Painel de Scripts no Desktop
    if (req.method === 'POST' && url === '/api/rodar-script-painel') {
      const body = JSON.parse((await readBody(req)) || '{}');
      const idx = body.index;
      const list = lerJson(PATH_SCRIPTS, []);
      const item = list[idx];

      if (!item || !fs.existsSync(item.arquivo)) {
        return sendJson(res, 400, { ok: false, erro: 'Script não encontrado no caminho: ' + (item ? item.arquivo : 'Índice inválido') });
      }

      // NAO usar 'python.exe' solto: no PATH deste Windows o primeiro match e o stub
      // de 0 byte da Microsoft Store, que sai com codigo 9009 sem executar nada.
      // Prioridade: venv do proprio projeto > venv da pasta pai > Python base > PATH.
      const pastaPai = path.dirname(item.pasta_execucao || '');
      const candidatos = [
        path.join(item.pasta_execucao || '', '.venv', 'Scripts', 'python.exe'),
        path.join(pastaPai, '.venv', 'Scripts', 'python.exe'),
        PAINEL.python,
      ];
      const py = candidatos.find((p) => { try { return fs.existsSync(p); } catch (_) { return false; } }) || 'python.exe';

      const args = ['-X', 'utf8', item.arquivo];
      const envBase = Object.assign({}, process.env, { PYTHONIOENCODING: 'utf-8', PYTHONUTF8: '1' });
      const ps = spawn(py, args, { cwd: item.pasta_execucao, env: envBase });
      let stdout = '', stderr = `[interpretador: ${py}]\n`;
      ps.stdout.on('data', d => stdout += d.toString('utf8'));
      ps.stderr.on('data', d => stderr += d.toString('utf8'));
      ps.on('close', code => {
        return sendJson(res, 200, { ok: code === 0, returnCode: code, stdout, stderr });
      });
      return;
    }


    // =================================================================
    //  PAINEL STREAMLIT (Gestao de ROs) - o Radar sobe e embute o painel
    //  que ja existe, em vez de manter uma segunda copia da mesma tela.
    // =================================================================
    if (req.method === 'GET' && url === '/api/painel-status') {
      const vivo = await painelNoAr();
      return sendJson(res, 200, {
        ok: true, rodando: vivo, url: PAINEL.url, pasta: PAINEL.pasta,
        existe: fs.existsSync(path.join(PAINEL.pasta, 'app.py')),
      });
    }

    if (req.method === 'POST' && url === '/api/painel-subir') {
      if (await painelNoAr()) return sendJson(res, 200, { ok: true, jaEstava: true, url: PAINEL.url });

      if (!fs.existsSync(PAINEL.python)) {
        return sendJson(res, 500, { ok: false, erro: `Python não encontrado em ${PAINEL.python}` });
      }
      if (!fs.existsSync(path.join(PAINEL.pasta, 'app.py'))) {
        return sendJson(res, 500, { ok: false, erro: `app.py não encontrado em ${PAINEL.pasta}` });
      }

      try {
        const p = spawn(PAINEL.python, [
          '-m', 'streamlit', 'run', 'app.py',
          '--server.port', String(PAINEL.porta),
          '--server.headless', 'true',
          '--browser.gatherUsageStats', 'false',
        ], { cwd: PAINEL.pasta, windowsHide: true, detached: false });

        p.stdout.on('data', (d) => { PAINEL.log += d.toString(); });
        p.stderr.on('data', (d) => { PAINEL.log += d.toString(); });
        p.on('exit', (c) => { PAINEL.log += `\n[streamlit encerrou, codigo ${c}]\n`; PAINEL.proc = null; });
        PAINEL.proc = p;
        PAINEL.log = '';
      } catch (e) {
        return sendJson(res, 500, { ok: false, erro: 'Falha ao iniciar o Streamlit: ' + e.message });
      }

      // espera o Streamlit atender antes de responder (ele leva alguns segundos)
      for (let i = 0; i < 40; i++) {
        await new Promise((s) => setTimeout(s, 500));
        if (await painelNoAr()) return sendJson(res, 200, { ok: true, url: PAINEL.url });
      }
      return sendJson(res, 500, { ok: false, erro: 'O Streamlit não respondeu em 20s.', log: PAINEL.log.slice(-800) });
    }

    if (req.method === 'POST' && url === '/api/painel-parar') {
      if (PAINEL.proc) { try { PAINEL.proc.kill(); } catch (_) {} PAINEL.proc = null; }
      return sendJson(res, 200, { ok: true });
    }

    // =================================================================
    //  VIGILANCIA DE PREGOES - consulta REAL via app/compras-api.js
    //
    //  INDEPENDENTE DO MODULO DE BOLETINS, por decisao do usuario.
    //  Sao dois fluxos separados de proposito: o boletim e a varredura
    //  diaria do que foi publicado; a vigilancia e o acompanhamento de
    //  pregoes especificos que ele escolhe a dedo. Nao ligar um no outro.
    //  Cadastro e sempre manual: numero + ano + UASG.
    // =================================================================

    // Lista o que esta sendo vigiado (ja atualizado pela varredura de fundo)
    if (req.method === 'GET' && url === '/api/acompanhados') {
      return sendJson(res, 200, {
        ok: true,
        acompanhados: arr(lerJson(P.acompanhados, [])),
        varredura: {
          ultima: VARREDURA.ultima,
          rodando: VARREDURA.rodando,
          intervaloMin: INTERVALO_MIN,
          erro: VARREDURA.ultimoErro,
        },
      });
    }

    // Marca as movimentacoes como vistas (tira o destaque ambar)
    if (req.method === 'POST' && url === '/api/vigilancia-vista') {
      const lista = arr(lerJson(P.acompanhados, []));
      const agora = new Date().toISOString();
      lista.forEach((x) => { if (x.mudouAgora) { x.mudouAgora = false; x.vistoEm = agora; } });
      fs.writeFileSync(P.acompanhados, JSON.stringify(lista, null, 2), 'utf8');
      return sendJson(res, 200, { ok: true, acompanhados: lista });
    }

    // Marca as movimentacoes de um pregao especifico como vistas/lidas
    if (req.method === 'POST' && url === '/api/vigilancia-marcar-lido') {
      const body = JSON.parse((await readBody(req)) || '{}');
      const lista = arr(lerJson(P.acompanhados, []));
      const agora = new Date().toISOString();
      lista.forEach((x) => {
        if (!body.id || x.id === body.id) {
          x.mudouAgora = false;
          x.vistoEm = agora;
          arr(x.eventos).forEach((e) => { e.novo = false; });
        }
      });
      fs.writeFileSync(P.acompanhados, JSON.stringify(lista, null, 2), 'utf8');
      return sendJson(res, 200, { ok: true, acompanhados: lista });
    }

    // Cadastra um pregao para vigilancia
    if (req.method === 'POST' && url === '/api/acompanhar') {
      const body = JSON.parse((await readBody(req)) || '{}');
      const portal = body.portal === 'licitacoes-e' ? 'licitacoes-e' : 'comprasgov';
      const lista0 = arr(lerJson(P.acompanhados, []));

      // -------- Licitacoes-e (Banco do Brasil) --------
      // Nao existe API publica: o portal e JSF com estado de sessao atras de
      // Cloudflare, e o link direto por numero nao abre a licitacao.
      // Entao aqui NAO se inventa status: guarda o numero, monta o link e
      // marca explicitamente como acompanhamento manual.
      if (portal === 'licitacoes-e') {
        const num = String(body.numero || '').replace(/\D/g, '');
        if (!num) return sendJson(res, 400, { ok: false, erro: 'Informe o número da licitação no Licitações-e (só dígitos). Ex.: 1095398' });

        const id = `bb-${num}`;
        if (lista0.some((x) => x.id === id)) {
          return sendJson(res, 409, { ok: false, erro: `A licitação ${num} do Licitações-e já está cadastrada.` });
        }
        const responsaveis = (Array.isArray(body.responsaveis) && body.responsaveis.length)
          ? body.responsaveis
          : (body.responsaveis ? [String(body.responsaveis).trim()] : ['henri.mafra@grupoenterprisecore.com']);

        const agora = new Date().toISOString();
        const novo = {
          id, portal, numero: num, ano: String(body.ano || new Date().getFullYear()).slice(0, 4),
          uasg: '', apelido: (body.apelido || '').trim() || `Licitações-e ${num}`,
          semConsultaAutomatica: true,
          responsaveis,
          dataCadastro: agora, ultimaChecagem: '', assinatura: '', erro: '',
          registro: {
            orgao: (body.orgao || '').trim(),
            objeto: (body.objeto || '').trim(),
            linkPortal: 'https://www.licitacoes-e.com.br/aop/consultar-detalhes-licitacao.aop'
              + `?opcao=consultarDetalhesLicitacao&numeroLicitacao=${num}`,
          },
          eventos: [{ quando: agora, texto: `Cadastrado para acompanhamento manual (Licitações-e nº ${num}).` }],
        };
        lista0.push(novo);
        fs.writeFileSync(P.acompanhados, JSON.stringify(lista0, null, 2), 'utf8');
        return sendJson(res, 200, { ok: true, item: novo, acompanhados: lista0 });
      }

      // -------- Compras.gov / PNCP (numero + UASG) --------
      const uasg = comprasApi.normalizaUasg(body.uasg);
      const numero = comprasApi.normalizaNumero(body.numero);
      const ano = String(body.ano || new Date().getFullYear()).slice(0, 4);

      if (!uasg) {
        return sendJson(res, 400, {
          ok: false,
          erro: `UASG inválida: "${body.uasg || ''}". Precisa ser só números (ex.: 70001). `
              + `Se o boletim trouxe o número do processo nesse campo, digite a UASG na mão.`,
        });
      }
      if (!numero) return sendJson(res, 400, { ok: false, erro: `Número do pregão inválido: "${body.numero || ''}"` });
      if (!/^\d{4}$/.test(ano)) return sendJson(res, 400, { ok: false, erro: `Ano inválido: "${ano}"` });

      const lista = arr(lerJson(P.acompanhados, []));
      const id = `${uasg}-${numero}-${ano}`;
      if (lista.some((x) => x.id === id)) {
        return sendJson(res, 409, { ok: false, erro: `Pregão ${numero}/${ano} da UASG ${uasg} já está sendo vigiado.` });
      }

      // Valida contra a API antes de aceitar: nao adianta vigiar o que nao existe
      const r = await comprasApi.consultarPregao({ uasg, numero, ano });
      if (!r.achado) {
        return sendJson(res, 404, {
          ok: false,
          erro: r.erro || 'Pregão não encontrado no Compras.gov.',
          dica: 'Confira o número e a UASG. A consulta cobre pregão eletrônico (modalidades 5 e 6) publicado no ano informado.',
        });
      }

      const responsaveis = (Array.isArray(body.responsaveis) && body.responsaveis.length)
        ? body.responsaveis
        : (body.responsaveis ? [String(body.responsaveis).trim()] : ['henri.mafra@grupoenterprisecore.com']);

      const agora = new Date().toISOString();
      const novo = {
        id,
        portal: 'comprasgov',
        semConsultaAutomatica: false,
        apelido: (body.apelido || '').trim() || r.registro.orgao,
        uasg, numero, ano,
        responsaveis,
        dataCadastro: agora,
        ultimaChecagem: agora,
        assinatura: comprasApi.assinatura(r.registro),
        registro: r.registro,
        erro: '',
        eventos: [{ quando: agora, texto: `Cadastrado na vigilância. Situação atual: ${r.registro.situacao}.` }],
      };
      lista.push(novo);
      fs.writeFileSync(P.acompanhados, JSON.stringify(lista, null, 2), 'utf8');
      return sendJson(res, 200, { ok: true, item: novo, acompanhados: lista });
    }

    // Atualiza a lista de responsaveis de um pregao existente
    if (req.method === 'POST' && url === '/api/vigilancia-atualizar-responsaveis') {
      const body = JSON.parse((await readBody(req)) || '{}');
      if (!body.id) return sendJson(res, 400, { ok: false, erro: 'ID do pregão obrigatório.' });
      const lista = arr(lerJson(P.acompanhados, []));
      const item = lista.find(x => x.id === body.id);
      if (!item) return sendJson(res, 404, { ok: false, erro: 'Pregão não encontrado.' });
      item.responsaveis = Array.isArray(body.responsaveis) ? body.responsaveis : [String(body.responsaveis || '').trim()].filter(Boolean);
      if (!item.responsaveis.length) item.responsaveis = ['henri.mafra@grupoenterprisecore.com'];
      fs.writeFileSync(P.acompanhados, JSON.stringify(lista, null, 2), 'utf8');
      return sendJson(res, 200, { ok: true, item, acompanhados: lista });
    }

    // Remove da vigilancia
    if (req.method === 'POST' && url === '/api/desacompanhar') {
      const body = JSON.parse((await readBody(req)) || '{}');
      const lista = arr(lerJson(P.acompanhados, []));
      const restante = lista.filter((x) => x.id !== body.id);
      if (restante.length === lista.length) return sendJson(res, 404, { ok: false, erro: 'Item não estava na vigilância.' });
      fs.writeFileSync(P.acompanhados, JSON.stringify(restante, null, 2), 'utf8');
      return sendJson(res, 200, { ok: true, acompanhados: restante });
    }

    // Checa status de verdade e detecta o que mudou desde a ultima leitura
        // Disparo manual forçado de alerta por e-mail
    if (req.method === 'POST' && url === '/api/vigilancia-disparar-alerta') {
      const body = JSON.parse((await readBody(req)) || '{}');
      const lista = arr(lerJson(P.acompanhados, []));
      let disparados = 0;
      for (const item of lista) {
        if (!body.id || item.id === body.id) {
          const evsNovos = arr(item.eventos).filter(e => e.novo).map(e => e.texto);
          const mudancas = evsNovos.length ? evsNovos : [`Acompanhamento ativo: ${item.apelido || item.registro?.orgao || item.numero}`];
          await dispararAlertaEmail(item, mudancas);
          disparados++;
        }
      }
      fs.writeFileSync(P.acompanhados, JSON.stringify(lista, null, 2), 'utf8');
      return sendJson(res, 200, { ok: true, disparados, acompanhados: lista });
    }

    
    // Endpoint para Extensão de Navegador ENTERPRISECORE Sentinela
    if (req.method === 'POST' && url === '/api/extensao/capturar') {
      const body = JSON.parse((await readBody(req)) || '{}');
      const reqToken = req.headers['x-enterprisecore-radar-token'] || body.token;
      if (reqToken !== EXTENSION_AUTH_TOKEN) {
        return sendJson(res, 401, { ok: false, erro: 'Acesso não autorizado: token de autenticação da extensão ausente ou inválido.' });
      }
      const { portal, identificador, uasg, numero, itemNumero, eventos = [] } = body;
      const lista = arr(lerJson(P.acompanhados, []));

      let item = lista.find(x => {
        if (identificador && (String(x.id) === String(identificador) || String(x.registro?.idCompra) === String(identificador) || String(x.numeroLicitacao) === String(identificador))) return true;
        if (identificador && String(identificador).includes(String(x.numero))) return true;
        if (uasg && String(x.uasg) === String(uasg) && numero && String(x.numero) === String(numero)) return true;
        if (numero && String(x.numero) === String(numero)) return true;
        return false;
      });

      const isNovoPregao = !item;
      if (!item) {
        const apelidoDetectado = body.nome || `Pregão ${numero || identificador} (${portal || 'Portal'})`;
        item = {
          id: `${portal || 'ext'}-${identificador || Date.now()}`,
          portal: portal || 'comprasgov',
          numero: numero || identificador,
          ano: new Date().getFullYear(),
          uasg: uasg || '',
          apelido: apelidoDetectado,
          criadoEm: new Date().toISOString(),
          eventos: [],
          mudouAgora: false,
          inicializado: false,
          responsaveis: ['henri.mafra@grupoenterprisecore.com', 'henrique.lessa@grupoenterprisecore.com']
        };
        lista.push(item);
      }

      // REGRA UNIVERSAL DE LINHA DE BASE:
      // Na primeira carga de qualquer pregão (atual ou futuro), absorve todo o histórico passado
      // para o painel ficar completo, mas NUNCA dispara e-mail de alerta para o passado!
      const ehCargaInicial = isNovoPregao || !item.inicializado || arr(item.eventos).length === 0;

      const historico = lerJson(P.historico, {});
      
      // Indexação normalizada contra variações de espaços e quebras de linha
      const norm = (t) => String(t || '').replace(/\s+/g, ' ').trim().toLowerCase();
      const conhecidos = new Set(arr(item.eventos).map(e => norm(e.texto)));
      const novosEventos = [];
      const agora = new Date().toISOString();

      for (const ev of eventos) {
        if (!ev.texto) continue;
        const textoNorm = norm(ev.texto);
        if (conhecidos.has(textoNorm)) continue;

        // FILTRO DE SEGURANÇA: Descarta cabeçalhos de tabela e ruído de tela
        if (textoNorm.includes('data e hora do registro') || textoNorm.includes('participante mensagem')) continue;
        if (textoNorm === '[licitações-e bb]: mensagem' || textoNorm === '[licitações-e bb]: participante') continue;

        conhecidos.add(textoNorm);
        
        // Se for carga inicial, NUNCA marca como novo (não alerta)
        // Se for monitoramento contínuo em tempo real, checa se a mensagem é recente (últimas 48h)
        const ehRecente = ehCargaInicial ? false : isMensagemRecente(ev.quando || ev.texto, 2);

        const novoEv = {
          quando: ev.quando || agora,
          texto: ev.texto,
          autor: ev.autor || 'EXTENSÃO ENTERPRISECORE',
          novo: ehRecente
        };
        novosEventos.push(novoEv);
        item.eventos = arr(item.eventos).concat([novoEv]);
        historico[item.id] = arr(historico[item.id]).concat([novoEv]);
      }

      // Marca o pregão como inicializado para sempre
      item.inicializado = true;

      // LIMITE SAUDÁVEL DO BANCO DE DADOS (Buffer Rolante anti-lotamento)
      if (item.eventos.length > 100) {
        item.eventos = item.eventos.slice(-100);
      }
      if (historico[item.id] && historico[item.id].length > 200) {
        historico[item.id] = historico[item.id].slice(-200);
      }

      if (novosEventos.length) {
        fs.writeFileSync(P.acompanhados, JSON.stringify(lista, null, 2), 'utf8');
        fs.writeFileSync(P.historico, JSON.stringify(historico, null, 2), 'utf8');

        // Dispara alerta por e-mail SOMENTE para eventos novos em tempo real (NUNCA na carga inicial)
        const paraAlertar = novosEventos.filter(e => e.novo).map(e => e.texto);
        if (paraAlertar.length && !ehCargaInicial) {
          item.mudouAgora = true;
          item.vistoEm = '';
          await dispararAlertaEmail(item, paraAlertar);
        } else {
          item.mudouAgora = false;
        }
      }

      const totalNovosReal = ehCargaInicial ? 0 : novosEventos.filter(e => e.novo).length;
      return sendJson(res, 200, {
        ok: true,
        pregao: item.apelido || item.numero,
        novos: totalNovosReal,
        historicoCarregado: novosEventos.length,
        total: item.eventos.length,
        inicializado: true
      });
    }

    if (req.method === 'POST' && url === '/api/checar-status') {
      const body = JSON.parse((await readBody(req)) || '{}');
      const lista = arr(lerJson(P.acompanhados, []));
      const filtro = arr(body.ids);
      const alvo = filtro.length ? lista.filter((x) => filtro.includes(x.id)) : lista;

      if (!alvo.length) return sendJson(res, 200, { ok: true, resultados: [], comMudanca: 0, quando: new Date().toISOString() });

      const historico = lerJson(P.historico, {});
      let comMudanca = 0;

      // mesma rotina da varredura automatica, para nao existirem duas logicas
      for (const item of alvo) {
        const m = await atualizarItem(item, historico);
        if (m.length) {
          comMudanca++;
          await dispararAlertaEmail(item, m);
        } else {
          const pendentes = arr(item.eventos).filter(e => e.novo && !e.alertaEnviado).map(e => e.texto);
          if (pendentes.length) {
            await dispararAlertaEmail(item, pendentes);
          }
        }
      }

      fs.writeFileSync(P.acompanhados, JSON.stringify(lista, null, 2), 'utf8');
      fs.writeFileSync(P.historico, JSON.stringify(historico, null, 2), 'utf8');
      VARREDURA.ultima = new Date().toISOString();
      return sendJson(res, 200, { ok: true, resultados: alvo, comMudanca, quando: VARREDURA.ultima });
    }

    // Abre a pasta dos boletins no Explorer
    if (req.method === 'POST' && url === '/api/abrir-pasta') {
      try { fs.mkdirSync(P.boletins, { recursive: true }); } catch (_) {}
      try {
        if (process.platform === 'win32') {
          spawn('explorer.exe', [P.boletins], { detached: true, stdio: 'ignore' }).unref();
          exec(`start "" "${P.boletins}"`, { shell: 'cmd.exe' });
        } else {
          spawn('xdg-open', [P.boletins], { detached: true }).unref();
        }
      } catch (_) {}
      return sendJson(res, 200, { ok: true, pasta: P.boletins });
    }

    // Download da planilha Excel consolidada (18 colunas)
    if (req.method === 'GET' && url === '/api/download-excel') {
      if (!fs.existsSync(P.xlsx)) {
        res.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' });
        return res.end('Planilha ainda não foi gerada. Processos os boletins primeiro.');
      }
      res.writeHead(200, {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': 'attachment; filename="Planilha-ENTERPRISECORE-Licitacoes.xlsx"',
      });
      fs.createReadStream(P.xlsx).pipe(res);
      return;
    }

    // Abre a planilha diretamente no Excel (no sistema operacional)
    if (req.method === 'POST' && url === '/api/abrir-excel') {
      if (!fs.existsSync(P.xlsx)) {
        return sendJson(res, 400, { ok: false, erro: 'A planilha ainda não foi gerada. Clique em Processar primeiro.' });
      }
      try {
        spawn('cmd.exe', ['/c', 'start', '""', P.xlsx], { windowsHide: true });
        return sendJson(res, 200, { ok: true });
      } catch (e) {
        return sendJson(res, 500, { ok: false, erro: e.message });
      }
    }

    // Exporta apenas as licitacoes selecionadas para planilha
    if (req.method === 'POST' && url === '/api/exportar-selecionados') {
      const body = JSON.parse((await readBody(req)) || '{}');
      if (!Array.isArray(body.ids) || !body.ids.length) {
        return sendJson(res, 400, { ok: false, erro: 'Selecione ao menos 1 licitação para exportar.' });
      }
      const tmpFile = path.join(os.tmpdir(), `ENTERPRISECORE-Selecionadas-${Date.now()}.xlsx`);
      const args = ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', P.psProcessar, '-OutXlsx', tmpFile, '-FilterIds', body.ids.join(',')];
      const ps = spawn('powershell.exe', args, { windowsHide: true });
      ps.on('close', (code) => {
        if (code === 0 && fs.existsSync(tmpFile)) {
          res.writeHead(200, {
            'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            'Content-Disposition': `attachment; filename="Planilha-ENTERPRISECORE-Selecionadas.xlsx"`
          });
          const stream = fs.createReadStream(tmpFile);
          stream.pipe(res);
          stream.on('end', () => { fs.unlink(tmpFile, () => {}); });
        } else {
          sendJson(res, 500, { ok: false, erro: 'Erro ao gerar planilha com selecionados.' });
        }
      });
      return;
    }

    // Processa (varre) os boletins da pasta -> gera dados.json + planilha de 17 colunas
    if (req.method === 'POST' && url === '/api/processar') {
      const body = JSON.parse((await readBody(req)) || '{}');
      try { fs.mkdirSync(P.boletins, { recursive: true }); } catch (_) {}
      const args = ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', P.psProcessar];
      if (!body.somar) args.push('-Zerar');   // padrao = recomecar do zero (reflete so a pasta atual)
      const ps = spawn('powershell.exe', args, { windowsHide: true });
      let out = '', err = '';
      ps.stdout.on('data', (d) => { out += d.toString(); });
      ps.stderr.on('data', (d) => { err += d.toString(); });
      ps.on('error', (e) => sendJson(res, 500, { ok: false, erro: 'Falha ao iniciar o PowerShell: ' + e.message }));
      ps.on('close', (code) => {
        const m = out.match(/RESULTADO_JSON\s+(\{.*\})/);
        if (code === 0 && m) {
          try { return sendJson(res, 200, { ok: true, resumo: JSON.parse(m[1]) }); } catch (_) {}
        }
        const msg = (err || out || 'erro ao processar').split('\n').filter(Boolean).slice(-6).join(' | ').trim();
        sendJson(res, 500, { ok: false, erro: msg });
      });
      return;
    }

    // Gerador de Modelo Outlook (.OFT). Um .oft com tudo, ou .zip com 1 por licitacao.
    if (req.method === 'POST' && url === '/api/gerar-oft') {
      const body = JSON.parse((await readBody(req)) || '{}');
      const itens = itensPorIds(body.ids);
      if (!itens.length) {
        return sendJson(res, 400, { ok: false, erro: 'Selecione ao menos 1 licitacao. Se acabou de reprocessar, recarregue a pagina.' });
      }

      // Sempre gera um .zip com 2 pastas:
      //  - "Radar de Publicacoes": 1 e-mail so, juntando os selecionados que
      //    caem nos estados configurados como "Minha regiao" (config/escopo-enterprisecore.json
      //    -> estadosPreferidos).
      //  - "Individuais": 1 e-mail por licitacao selecionada (todas, de qualquer
      //    estado), no modelo de alerta da equipe comercial.
      const cfgEmail = (lerJson(P.escopo, {}).email) || {};
      const estadosPref = lerJson(P.escopo, {}).estadosPreferidos || {};
      const ufsRegiao = new Set(arr(estadosPref.ufs).map((u) => String(u).toUpperCase()));
      const nomeRegiao = estadosPref.nome || cfgEmail.regiao || 'Minha região';
      const itensRegiao = ufsRegiao.size ? itens.filter((x) => ufsRegiao.has(siglaUf(x.estado))) : [];

      const payloadItens = [];

      if (itensRegiao.length) {
        const tituloRegiao = cfgEmail.titulo || '📋 Radar de Publicações';
        const hoje = new Date().toLocaleDateString('pt-BR');
        payloadItens.push({
          pasta: 'Radar de Publicacoes',
          subject: `${tituloRegiao} · ${hoje}`,
          html: emailHtml(itensRegiao, tituloRegiao, { regiao: nomeRegiao }),
          orgao: nomeRegiao,
          numeroConlicitacao: `${itensRegiao.length}-itens`,
        });
      }

      itens.forEach((x) => {
        payloadItens.push({
          pasta: 'Individuais',
          subject: assuntoIndividual(x),
          html: emailHtmlIndividual(x),
          orgao: x.orgao || '',
          numeroConlicitacao: x.numeroConlicitacao || '',
        });
      });

      const stamp = Date.now();
      const tmpJson = path.join(os.tmpdir(), `enterprisecore-oft-${stamp}.json`);
      const tmpOut  = path.join(os.tmpdir(), `Modelos-ENTERPRISECORE-${stamp}.zip`);

      fs.writeFileSync(tmpJson, JSON.stringify({ to: body.to || '', cc: body.cc || '', items: payloadItens }), 'utf8');

      // ATENCAO: os parametros do script sao -JsonPath e -OutPath
      const ps = spawn('powershell.exe', [
        '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass',
        '-File', P.psOft, '-JsonPath', tmpJson, '-OutPath', tmpOut,
      ], { windowsHide: true });

      let out = '', err = '';
      ps.stdout.on('data', (d) => { out += d.toString(); });
      ps.stderr.on('data', (d) => { err += d.toString(); });
      ps.on('error', (e) => {
        try { fs.unlinkSync(tmpJson); } catch (_) {}
        sendJson(res, 500, { ok: false, erro: 'Falha ao executar o PowerShell: ' + e.message });
      });
      ps.on('close', (code) => {
        try { fs.unlinkSync(tmpJson); } catch (_) {}
        if (code === 0 && fs.existsSync(tmpOut)) {
          res.writeHead(200, {
            'Content-Type': 'application/zip',
            'Content-Disposition': 'attachment; filename="Modelos-ENTERPRISECORE.zip"',
          });
          const stream = fs.createReadStream(tmpOut);
          stream.pipe(res);
          stream.on('end', () => { fs.unlink(tmpOut, () => {}); });
        } else {
          try { if (fs.existsSync(tmpOut)) fs.unlinkSync(tmpOut); } catch (_) {}
          sendJson(res, 500, { ok: false, erro: 'Erro ao gerar o modelo do Outlook: ' + (err || out || 'sem detalhe').trim() });
        }
      });
      return;
    }

    // Abre o Outlook com o e-mail pronto
    if (req.method === 'POST' && url === '/api/enviar') {
      const body = JSON.parse((await readBody(req)) || '{}');
      if (!body.to) return sendJson(res, 400, { ok: false, erro: 'Informe o destinatario (Para).' });

      const itens = itensPorIds(body.ids);
      if (!itens.length) return sendJson(res, 400, { ok: false, erro: 'Selecione ao menos 1 licitacao.' });

      const assunto = (body.assunto || '').trim()
        || `Oportunidades de licitacao - ${itens.length} selecionada${itens.length === 1 ? '' : 's'}`;

      const tmp = path.join(os.tmpdir(), `enterprisecore-mail-${Date.now()}.json`);
      fs.writeFileSync(tmp, JSON.stringify({
        to: body.to, cc: body.cc || '', subject: assunto, html: emailHtml(itens, assunto),
      }), 'utf8');

      const ps = spawn('powershell.exe', [
        '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass',
        '-File', P.psEnviar, '-JsonPath', tmp,
      ], { windowsHide: true });

      let stderr = '';
      ps.stderr.on('data', (d) => { stderr += d.toString(); });
      ps.on('error', (e) => {
        try { fs.unlinkSync(tmp); } catch (_) {}
        sendJson(res, 500, { ok: false, erro: 'Falha ao iniciar o PowerShell: ' + e.message });
      });
      ps.on('close', (code) => {
        try { fs.unlinkSync(tmp); } catch (_) {}
        if (code === 0) sendJson(res, 200, { ok: true });
        else sendJson(res, 500, { ok: false, erro: (stderr || 'Erro desconhecido ao abrir o Outlook').trim() });
      });
      return;
    }

    if (url === '/ENTERPRISECORE_Manual_Completo.pdf' || url === '/pdf') {
      const pdfFile = path.join(ROOT,'ENTERPRISECORE_Manual_Completo.pdf');
      if (fs.existsSync(pdfFile)) {
        res.writeHead(200, {
          'Content-Type': 'application/pdf',
          'Content-Disposition': 'inline; filename="ENTERPRISECORE_Manual_Completo.pdf"'
        });
        fs.createReadStream(pdfFile).pipe(res);
        return;
      }
    }

    // Os 6 manuais dedicados (uso + construcao, para Boletins, ROs e Pregoes).
    // Cada painel abre os seus dois pelo botao correspondente na tela.
    const MANUAIS = {
      '/manual/boletins-uso':        { arq: P.manualUsoBoletins,        nome: 'Manual de Uso - Radar de Boletins.pdf' },
      '/manual/boletins-construcao': { arq: P.manualConstrucaoBoletins, nome: 'Manual de Construcao - Radar de Boletins.pdf' },
      '/manual/ros-uso':             { arq: P.manualUsoRos,             nome: 'Manual de Uso - Painel de ROs.pdf' },
      '/manual/ros-construcao':      { arq: P.manualConstrucaoRos,      nome: 'Manual de Construcao - Painel de ROs.pdf' },
      '/manual/pregoes-uso':         { arq: P.manualUsoPregoes,         nome: 'Manual de Uso - Radar de Pregoes.pdf' },
      '/manual/pregoes-construcao':  { arq: P.manualConstrucaoPregoes,  nome: 'Manual de Construcao - Radar de Pregoes.pdf' },
    };
    if (MANUAIS[url]) {
      const m = MANUAIS[url];
      if (fs.existsSync(m.arq)) {
        res.writeHead(200, {
          'Content-Type': 'application/pdf',
          'Content-Disposition': `inline; filename="${m.nome}"`
        });
        fs.createReadStream(m.arq).pipe(res);
        return;
      }
      res.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' });
      res.end('Manual nao encontrado: ' + m.arq);
      return;
    }

    if (url === '/ENTERPRISECORE_Manual_Completo.html' || url === '/manual-html') {
      const htmlFile = path.join(ROOT,'..', '.gemini', 'antigravity', 'brain', 'e711903a-87a6-446f-9837-44701c393d3a', 'ENTERPRISECORE_Manual_Completo.html');
      if (fs.existsSync(htmlFile)) {
        res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
        return fs.createReadStream(htmlFile).pipe(res);
      }
    }

    res.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' });
    res.end('Nao encontrado');
  } catch (e) {
    sendJson(res, 500, { ok: false, erro: e.message });
  }
});

server.on('error', (e) => {
  if (e.code === 'EADDRINUSE') {
    console.error(`\n[ERRO] A porta ${PORT} ja esta em uso. Feche o app anterior ou troque a porta no server.js.\n`);
    process.exit(1);
  }
  throw e;
});

server.listen(PORT, () => {
  console.log('\n============================================================');
  console.log('  ENTERPRISECORE Radar rodando!');
  console.log(`  Abra no navegador:  http://localhost:${PORT}`);
  console.log(`  Radar de Pregões 24/7 ativo (varredura a cada ${INTERVALO_MIN} min)`);
  console.log('  Para encerrar: feche esta janela ou pressione Ctrl+C');
  console.log('============================================================\n');

  // Varredura automatica 24/7 ativa
  setTimeout(() => varrerVigilancia('varredura inicial'), 3000);
  const t = setInterval(() => varrerVigilancia('varredura periodica 24/7'), INTERVALO_MIN * 60 * 1000);
  t.unref();
});
