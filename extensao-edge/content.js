// =====================================================================
//  ENTERPRISECORE Radar — Sentinela de Pregões (Content Script Hardened V1.3)
//  Compatível com Licitações-e (BB), Compras.gov.br e Compras Públicas
//  Proteção contra Context Invalidation e Tratamento de Telas de Detalhes
// =====================================================================

(function () {
  'use strict';

  if (window.__ENTERPRISECORE_SENTINELA_INJETADO__) return;
  window.__ENTERPRISECORE_SENTINELA_INJETADO__ = true;

  let modoSentinelaAtivo = true;
  let timerSentinela = null;
  let ultimoUrlConhecido = window.location.href;

  // ---------- 0. Proteção Anti-Invalidation do Contexto ----------
  function isContextoValido() {
    try {
      return typeof chrome !== 'undefined' && Boolean(chrome?.runtime?.id);
    } catch (_) {
      return false;
    }
  }

  function pararSentinelaSeInvalido() {
    if (!isContextoValido()) {
      if (timerSentinela) {
        clearInterval(timerSentinela);
        timerSentinela = null;
      }
      return true;
    }
    return false;
  }

  function escHtml(str) {
    return String(str || '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  // ---------- 1. Detecção Inteligente do Contexto ----------
  function detectarContexto() {
    const url = window.location.href;
    const hostname = window.location.hostname;
    const bodyTxt = document.body ? document.body.innerText || '' : '';
    const cacheSessao = sessionStorage.getItem('enterprisecore_pregao_id') || '';

    // A. Licitações-e (Banco do Brasil)
    if (hostname.includes('licitacoes-e.com.br') || hostname.includes('bb.com.br')) {
      let numLic = '';
      try {
        const urlObj = new URL(url);
        numLic = urlObj.searchParams.get('numeroLicitacao') ||
                 urlObj.searchParams.get('numLicitacao') ||
                 urlObj.searchParams.get('licitacao') ||
                 urlObj.searchParams.get('id') || '';
      } catch (_) {}

      if (!numLic) {
        const inputs = document.querySelectorAll('input[name*="licitacao" i], input[id*="licitacao" i], input[name*="Licitacao"]');
        for (const inp of inputs) {
          const val = (inp.value || '').trim();
          if (/^\d{5,8}$/.test(val)) {
            numLic = val;
            break;
          }
        }
      }

      if (!numLic) {
        const m = /(?:n[ºo°]?\s*(?:da)?\s*licita[çc][ãa]o|licita[çc][ãa]o\s*n[ºo°]?|licita[çc][ãa]o\s*[:\-])[\s:]*(\d{5,8})/i.exec(bodyTxt);
        if (m) numLic = m[1];
      }

      if (!numLic) {
        const m = /n[ºo°][\s:]*(\d{5,8})/i.exec(bodyTxt);
        if (m) numLic = m[1];
      }

      if (!numLic && cacheSessao) {
        numLic = cacheSessao;
      }

      if (numLic) {
        sessionStorage.setItem('enterprisecore_pregao_id', numLic);
      }

      return {
        portal: 'licitacoes-e',
        identificador: (numLic || '').replace(/[^0-9]/g, ''),
        itemNumero: '',
        nome: numLic ? `Licitações-e (Nº ${numLic})` : 'Licitações-e (Nº pendente)'
      };
    }

    // B. Compras.gov.br / Serpro
    if (hostname.includes('serpro.gov.br') || hostname.includes('comprasnet.gov.br') || hostname.includes('compras.gov.br')) {
      let compra = '';
      let itemNumero = '';
      try {
        const urlObj = new URL(url);
        compra = urlObj.searchParams.get('compra') ||
                 urlObj.searchParams.get('idCompra') ||
                 urlObj.searchParams.get('numeroCompra') || '';
        const mItem = /\/item\/(\d+)/i.exec(url);
        if (mItem) itemNumero = mItem[1];
      } catch (_) {}

      if (!compra) {
        const mC = /(?:compra|uasg)[^\d]*(\d{15,18})/i.exec(bodyTxt);
        if (mC) compra = mC[1];
      }
      if (!compra) {
        const mUasgPreg = /(?:uasg\s*[:\-]?\s*(\d{5,6}))[^\n\r]*?(?:preg[ãa]o\s*[:\-]?\s*(\d{1,6})\/(\d{4}))/i.exec(bodyTxt);
        if (mUasgPreg) {
          compra = `${mUasgPreg[1].padStart(6, '0')}05${mUasgPreg[2].padStart(5, '0')}${mUasgPreg[3]}`;
        }
      }
      // Detecções diretas por órgãos acompanhados
      if (!compra && bodyTxt.includes('806030') && bodyTxt.includes('90227')) {
        compra = '80603005902272026';
      }
      if (!compra && bodyTxt.includes('590001') && bodyTxt.includes('29/2026')) {
        compra = '59000105000292026';
      }
      if (!compra && bodyTxt.includes('070001') && bodyTxt.includes('90009')) {
        compra = '07000105900092026';
      }

      if (!compra && cacheSessao) compra = cacheSessao;
      if (compra) sessionStorage.setItem('enterprisecore_pregao_id', compra);

      return {
        portal: 'comprasgov',
        identificador: (compra || '').replace(/[^0-9]/g, ''),
        itemNumero: (itemNumero || '').replace(/[^0-9]/g, ''),
        nome: compra ? `Compras.gov (Compra ${compra})` : 'Compras.gov'
      };
    }

    // D. Compras Eletrônicas PROCERGS / Compras-RS
    if (hostname.includes('procergs.rs.gov.br') || hostname.includes('compras.rs.gov.br')) {
      let numEdital = '';
      let idOffer = '';
      try {
        const urlObj = new URL(url);
        idOffer = urlObj.searchParams.get('idOffer') || urlObj.searchParams.get('idOfferFiltered') || '';
        const mEd = /\/editais\/(\d+_\d+|\d+\/\d+)/i.exec(url);
        if (mEd) numEdital = mEd[1].replace('_', '/');
      } catch (_) {}

      if (!numEdital) {
        const m = /(?:edital|preg[ãa]o)[^\d]*(\d{2,4}\/\d{4})/i.exec(bodyTxt);
        if (m) numEdital = m[1];
      }
      if (!numEdital && cacheSessao) numEdital = cacheSessao;
      if (numEdital) sessionStorage.setItem('enterprisecore_pregao_id', numEdital);

      const ident = numEdital || idOffer;
      return {
        portal: 'procergs',
        identificador: ident,
        itemNumero: '',
        nome: `PROCERGS ${numEdital ? '(Pregão ' + numEdital + ')' : ''}`
      };
    }

    // C. Portal de Compras Públicas
    if (hostname.includes('portaldecompraspublicas.com.br')) {
      let cod = '';
      const mCod = /(\d{5,7})/i.exec(url);
      if (mCod) cod = mCod[1];
      if (!cod) {
        const m = /(?:processo|licita[çc][ãa]o|preg[ãa]o)[\s:]*(\d{5,7})/i.exec(bodyTxt);
        if (m) cod = m[1];
      }
      if (!cod && cacheSessao) cod = cacheSessao;
      if (cod) sessionStorage.setItem('enterprisecore_pregao_id', cod);

      return {
        portal: 'compraspublicas',
        identificador: cod,
        itemNumero: '',
        nome: cod ? `Compras Públicas (Nº ${cod})` : 'Portal de Compras Públicas'
      };
    }

    return null;
  }

  // ---------- 2. Extrator de Dados do Licitações-e (BB) ----------
  function extrairDadosLicitacoesE() {
    const eventos = [];
    const rows = document.querySelectorAll('table tr');
    let count = 0;

    for (const r of rows) {
      if (count++ > 50) break;
      if (['INPUT', 'TEXTAREA'].includes(r.tagName)) continue;

      const tds = r.querySelectorAll('td');
      if (tds.length >= 2) {
        const col0 = (tds[0].innerText || '').trim();
        const col1 = (tds[1].innerText || '').trim();
        const col2 = tds.length >= 3 ? (tds[2].innerText || '').trim() : '';

        // Valida se a primeira coluna é realmente uma data ou hora (ex: 03/09/2026 ou 11:20:00)
        if (/\d{1,2}[\/:]\d{2}/.test(col0) && !/data/i.test(col0)) {
          const autor = col1.toUpperCase().includes('PREGOEIRO') ? 'PREGOEIRO' : 'SISTEMA';
          const msg = col2 || col1;
          if (msg.length > 3) {
            eventos.push({
              autor: autor,
              quando: col0,
              texto: `[LICITAÇÕES-E BB] ${col0} | ${col1}: ${msg.replace(/\s+/g, ' ')}`
            });
          }
          continue;
        }
      }

      // Fallback para avisos fora de tabela com data/hora
      const txt = (r.innerText || '').trim();
      if (/data\s*e\s*hora/i.test(txt) || (/participante/i.test(txt) && /mensagem/i.test(txt))) continue;
      if (!/\d{1,2}\/\d{1,2}\/\d{2,4}|\d{1,2}:\d{2}/.test(txt)) continue;

      if (txt.includes('Pregoeiro') || txt.includes('Sistema') || txt.includes('Lote') || txt.includes('Decisão') || txt.includes('Despacho')) {
        if (txt.length > 20 && txt.length < 2000 && !txt.includes('Consultar') && !txt.includes('Voltar')) {
          eventos.push({
            autor: txt.includes('Pregoeiro') ? 'PREGOEIRO' : 'SISTEMA',
            texto: `[LICITAÇÕES-E BB]: ${txt.replace(/\s+/g, ' ')}`
          });
        }
      }
    }

    // Leitura da Situação da Licitação (na tela de detalhes consultar-detalhes-licitacao.aop)
    const bodyText = document.body ? document.body.innerText || '' : '';
    const mSit = /(?:situa[çc][ãa]o\s*(?:da\s*licita[çc][ãa]o)?\s*[:\-])\s*([^\n\r<]{3,60})/i.exec(bodyText);
    if (mSit) {
      const sit = mSit[1].trim();
      if (sit && !sit.toLowerCase().includes('consultar') && sit.length > 3) {
        eventos.push({
          autor: 'SISTEMA',
          texto: `[LICITAÇÕES-E BB]: Situação do certame: ${sit}`
        });
      }
    }

    return eventos;
  }

  
  // ---------- 3.B Extrator de Dados da PROCERGS / Compras-RS ----------
  function extrairDadosProcergs() {
    const eventos = [];
    const bodyTxt = document.body ? document.body.innerText || '' : '';
    
    // Situação
    const mSit = /Situa[çc][ãa]o\s*[:\-]?\s*([^\n\r<]{3,60})/i.exec(bodyTxt);
    if (mSit) {
      const sit = mSit[1].trim();
      if (sit && !sit.toLowerCase().includes('consultar') && sit.length > 3) {
        eventos.push({
          autor: 'SISTEMA',
          texto: `[PROCERGS / COMPRAS-RS]: Situação: ${sit}`
        });
      }
    }

    // Histórico de Eventos do Lote (Tabela Estruturada de 4 colunas)
    const rows = document.querySelectorAll('table tr');
    for (const r of rows) {
      const tds = r.querySelectorAll('td');
      if (tds.length >= 4) {
        const dataHora = (tds[0].innerText || '').trim();
        const nomeEvento = (tds[1].innerText || '').trim();
        const responsavel = (tds[2].innerText || '').trim();
        const descr = (tds[3].innerText || '').trim();

        if (/\d{1,2}\/\d{1,2}\/\d{4}/.test(dataHora) && !/data/i.test(dataHora)) {
          eventos.push({
            autor: responsavel.toUpperCase().includes('PREGOEIRO') || responsavel.toLowerCase().includes('mateus') ? 'PREGOEIRO' : 'SISTEMA',
            quando: dataHora,
            texto: `[PROCERGS] ${dataHora} | ${nomeEvento} (${responsavel}): ${descr.replace(/\s+/g, ' ')}`
          });
          continue;
        }
      }

      const txt = (r.innerText || '').trim();
      if (txt.includes('Pregoeiro') || txt.includes('Esclarecimento') || txt.includes('Impugnação') || txt.includes('Ata') || txt.includes('Lote')) {
        if (txt.length > 20 && txt.length < 2000 && !txt.includes('Consultar') && !txt.includes('Voltar') && !txt.includes('Posição do edital')) {
          eventos.push({
            autor: 'PREGOEIRO',
            texto: `[PROCERGS / COMPRAS-RS]: ${txt.replace(/\s+/g, ' ')}`
          });
        }
      }
    }
    return eventos;
  }

  // ---------- 3. Extrator de Dados do Compras.gov.br (Omni-Channel Serpro) ----------
  function extrairDadosComprasnet() {
    const eventos = [];
    const textContent = document.body ? document.body.innerText || '' : '';
    if (textContent.length > 500000) return eventos;

    // 1. Decisão do Pregoeiro em Recurso
    const regexPregoeiro = /Decisão\s+do\s+pregoeiro[\s\S]*?Decisão\s+tomada[\s:]*([^\n]+)[\s\S]*?Data\s+decisão[\s:]*([^\n]+)[\s\S]*?Fundamentação[\s:]*([\s\S]*?)(?=Revisão\s+da\s+autoridade|Recursos\s+e\s+contrarrazões|$)/i;
    const mPreg = regexPregoeiro.exec(textContent);
    if (mPreg) {
      eventos.push({
        autor: 'PREGOEIRO',
        quando: mPreg[2].trim().slice(0, 30),
        texto: `[RECURSO ADM - DECISÃO DO PREGOEIRO]: Decisão tomada: ${mPreg[1].trim().toUpperCase()}. Fundamentação: ${mPreg[3].trim().slice(0, 800)}...`
      });
    }

    // 2. Revisão da Autoridade Competente
    const regexAutoridade = /Revisão\s+da\s+autoridade\s+competente[\s\S]*?Decisão\s+tomada[\s:]*([^\n]+)[\s\S]*?Data\s+decisão[\s:]*([^\n]+)[\s\S]*?Fundamentação[\s:]*([\s\S]*?)(?=Decisão\s+do\s+pregoeiro|Recursos\s+e\s+contrarrazões|$)/i;
    const mAut = regexAutoridade.exec(textContent);
    if (mAut) {
      eventos.push({
        autor: 'AUTORIDADE COMPETENTE',
        quando: mAut[2].trim().slice(0, 30),
        texto: `[RECURSO ADM - REVISÃO AUTORIDADE COMPETENTE]: Decisão tomada: ${mAut[1].trim().toUpperCase()}. Fundamentação: ${mAut[3].trim().slice(0, 800)}...`
      });
    }

    // 3. Varredura Estruturada de Tabelas e Linhas (table tr, mat-row, br-item)
    const rows = document.querySelectorAll('table tr, mat-row, .mat-row, .br-item, [role="row"], .linha-mensagem');
    for (const r of rows) {
      if (['INPUT', 'TEXTAREA', 'SELECT', 'FORM'].includes(r.tagName)) continue;
      const txt = (r.innerText || '').trim();
      if (!txt || txt.length < 15 || txt.length > 2500) continue;

      if (/data\s*e\s*hora/i.test(txt) && /participante|mensagem/i.test(txt)) continue;

      const temData = /\d{1,2}\/\d{1,2}\/\d{2,4}|\d{1,2}:\d{2}/.test(txt);
      const temChave = /pregoeiro|sistema|fornecedor|arrematante|licitante|proposta|desclassific|habilitad|inabilitad|recurso|adjudic|homolog/i.test(txt);

      if (temData && temChave) {
        const autor = /pregoeiro/i.test(txt) ? 'PREGOEIRO' : (/sistema/i.test(txt) ? 'SISTEMA' : 'FORNECEDOR');
        eventos.push({
          autor: autor,
          texto: `[COMPRAS.GOV]: ${txt.replace(/\s+/g, ' ')}`
        });
      }
    }

    // 4. Balões de Chat e Componentes de Mensagem (Angular, GovBR, Web Components)
    const msgElements = document.querySelectorAll(
      '.chat-msg, .mensagem-item, .item-mensagem, tr.mensagem, ' +
      '.mensagem, .msg, .texto-mensagem, .card-mensagem, .balao-mensagem, ' +
      '[class*="mensagem" i], [class*="chat" i], [class*="evento" i]'
    );
    let count = 0;
    for (const el of msgElements) {
      if (count++ > 60) break;
      if (['INPUT', 'TEXTAREA', 'SELECT', 'FORM', 'BODY', 'HTML'].includes(el.tagName)) continue;
      const txt = (el.innerText || '').trim();
      if (txt.length > 15 && txt.length < 2000 && !txt.includes('Consultar') && !txt.includes('Voltar')) {
        if (/pregoeiro|sistema|arrematante|aberto prazo|declarado vencedor|proposta readequada|desclassific|habilitad/i.test(txt)) {
          const autor = /pregoeiro/i.test(txt) ? 'PREGOEIRO' : 'SISTEMA';
          eventos.push({
            autor: autor,
            texto: `[CHAT / COMPRASNET]: ${txt.replace(/\s+/g, ' ')}`
          });
        }
      }
    }

    // 5. Fallback por Linhas de Texto do Chat
    const chatLines = textContent.split('\n');
    for (const line of chatLines) {
      const l = line.trim();
      if (l.length > 20 && l.length < 1000) {
        if (/^(?:\[?(?:Pregoeiro|Sistema|Coordenador)\]?[\s:]+)/i.test(l) || /^(?:(?:\d{2}\/\d{2}\/\d{4}|\d{2}:\d{2})[^\n]+(?:Pregoeiro|Sistema))/i.test(l)) {
          const autor = /pregoeiro/i.test(l) ? 'PREGOEIRO' : 'SISTEMA';
          eventos.push({
            autor: autor,
            texto: `[MENSAGEM SERPRO]: ${l.replace(/\s+/g, ' ')}`
          });
        }
      }
    }

    // Deduplica eventos por texto normalizado
    const vistos = new Set();
    const eventosUnicos = [];
    for (const ev of eventos) {
      const chave = ev.texto.replace(/\[[^\]]+\]:\s*/, '').replace(/\s+/g, ' ').trim().toLowerCase();
      if (!vistos.has(chave)) {
        vistos.add(chave);
        eventosUnicos.push(ev);
      }
    }

    return eventosUnicos;
  }

  // ---------- 4. Sincronização Segura com o Radar ----------
  function sincronizarComRadar(manual = false) {
    if (pararSentinelaSeInvalido()) return;

    const ctx = detectarContexto();
    if (!ctx || !ctx.identificador) {
      if (manual) {
        mostrarToast('⚠️ Preencha o Nº da Licitação no cartão abaixo para sincronizar.', 'warn');
        const inp = document.getElementById('enterprisecoreInputManual');
        if (inp) inp.focus();
      }
      return;
    }

    let eventos = [];
    if (ctx.portal === 'comprasgov') eventos = extrairDadosComprasnet();
    else if (ctx.portal === 'licitacoes-e') eventos = extrairDadosLicitacoesE();
    else if (ctx.portal === 'procergs') eventos = extrairDadosProcergs();

    if (!eventos.length) {
      if (manual) mostrarToast('ℹ️ Nenhuma nova mensagem ou decisão visível nesta tela.', 'info');
      return;
    }

    atualizarBadgeStatus('Sincronizando...', 'sync');

    try {
      if (pararSentinelaSeInvalido()) return;

      chrome.runtime.sendMessage(
        {
          action: 'SINCRONIZAR_RADAR',
          payload: {
            portal: ctx.portal,
            identificador: ctx.identificador,
            itemNumero: ctx.itemNumero,
            nome: ctx.nome,
            eventos: eventos
          }
        },
        (res) => {
          if (pararSentinelaSeInvalido()) return;
          if (chrome.runtime.lastError) {
            atualizarBadgeStatus('Radar Offline', 'error');
            if (manual) mostrarToast('❌ Radar local não respondeu (verifique se está rodando).', 'error');
            return;
          }

          if (res && res.ok) {
            const dados = res.dados || {};
            const horaAtual = new Date().toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' });
            if (dados.novos > 0) {
              mostrarToast(`🚨 ENTERPRISECORE Radar: ${dados.novos} nova(s) movimentação(ões) capturada(s) e alertada(s)!`, 'success');
              atualizarBadgeStatus(`Novo Alerta (${horaAtual})`, 'warn');
            } else {
              if (manual) {
                mostrarToast(`✅ Radar conectado! Todas as informações desta tela já estão sincronizadas.`, 'success');
              }
              atualizarBadgeStatus(`Ativo (às ${horaAtual})`, 'ok');
            }
          } else {
            const msgErro = res?.dados?.erro || 'Erro ao sincronizar';
            atualizarBadgeStatus('Erro', 'warn');
            if (manual) mostrarToast(`⚠️ ${msgErro}`, 'warn');
          }
        }
      );
    } catch (e) {
      if (String(e).includes('context invalidated')) {
        pararSentinelaSeInvalido();
      } else {
        console.warn('[ENTERPRISECORE Radar]', e);
      }
    }
  }

  // ---------- 5. Injeção do Widget Flutuante ----------
  function injetarWidget() {
    if (pararSentinelaSeInvalido()) return;

    let widget = document.getElementById('enterprisecore-radar-widget');
    const ctx = detectarContexto();
    if (!ctx) return;

    if (!widget) {
      widget = document.createElement('div');
      widget.id = 'enterprisecore-radar-widget';
      widget.className = 'enterprisecore-sentinela-box';
      document.body.appendChild(widget);
    }

    const temId = Boolean(ctx.identificador);

    widget.innerHTML = `
      <div class="enterprisecore-widget-header">
        <div class="enterprisecore-widget-title">
          <span class="enterprisecore-dot" style="background:${temId ? '#10b981' : '#f59e0b'};box-shadow:0 0 8px ${temId ? '#10b981' : '#f59e0b'}"></span>
          <b>ENTERPRISECORE RADAR</b> <small>Sentinela</small>
        </div>
        <button class="enterprisecore-btn-close" id="enterprisecoreBtnMin" title="Minimizar">_</button>
      </div>
      <div class="enterprisecore-widget-body" id="enterprisecoreWidgetBody">
        <div class="enterprisecore-target-name" id="enterprisecoreTargetName" style="color:${temId ? '#38bdf8' : '#f59e0b'}">${escHtml(ctx.nome)}</div>
        
        ${!temId ? `
          <div style="background:#1e293b;padding:8px;border-radius:6px;margin-bottom:10px;border:1px solid #334155">
            <span style="font-size:11px;color:#cbd5e1;display:block;margin-bottom:4px">Informe o Nº da Licitação:</span>
            <div style="display:flex;gap:4px">
              <input type="text" id="enterprisecoreInputManual" placeholder="Ex: 1095715" style="flex:1;padding:5px 8px;background:#0f172a;color:#fff;border:1px solid #475569;border-radius:4px;font-size:12px">
              <button id="enterprisecoreBtnSalvarManual" style="background:#0284c7;color:#fff;border:none;padding:5px 10px;border-radius:4px;cursor:pointer;font-weight:700;font-size:11px">Salvar</button>
            </div>
          </div>
        ` : `
          <div class="enterprisecore-status-row">
            <span>Status:</span>
            <b id="enterprisecoreStatusTxt" class="enterprisecore-txt-ok">Ativo</b>
          </div>
        `}

        <div class="enterprisecore-actions">
          <button class="enterprisecore-btn-sync" id="enterprisecoreBtnSync">⚡ Sincronizar Agora</button>
        </div>
        
        <div class="enterprisecore-toggle-row">
          <label class="enterprisecore-switch">
            <input type="checkbox" id="enterprisecoreToggleAuto" ${modoSentinelaAtivo ? 'checked' : ''}>
            <span class="enterprisecore-slider"></span>
          </label>
          <span style="font-size:11px;color:#94a3b8">Auto-sentinela (60s sem F5)</span>
        </div>
      </div>
    `;

    document.getElementById('enterprisecoreBtnSync').addEventListener('click', () => sincronizarComRadar(true));
    document.getElementById('enterprisecoreBtnMin').addEventListener('click', () => {
      const b = document.getElementById('enterprisecoreWidgetBody');
      b.style.display = b.style.display === 'none' ? 'block' : 'none';
    });

    if (!temId) {
      const btnSalvar = document.getElementById('enterprisecoreBtnSalvarManual');
      const inp = document.getElementById('enterprisecoreInputManual');
      if (btnSalvar && inp) {
        btnSalvar.addEventListener('click', () => {
          const val = inp.value.trim().replace(/[^0-9]/g, '');
          if (val) {
            sessionStorage.setItem('enterprisecore_pregao_id', val);
            mostrarToast(`✅ Licitação Nº ${val} identificada com sucesso!`, 'success');
            injetarWidget();
            sincronizarComRadar(true);
          }
        });
        inp.addEventListener('keypress', (e) => {
          if (e.key === 'Enter') btnSalvar.click();
        });
      }
    }

    const chkAuto = document.getElementById('enterprisecoreToggleAuto');
    if (chkAuto) {
      chkAuto.addEventListener('change', (e) => {
        modoSentinelaAtivo = e.target.checked;
        if (modoSentinelaAtivo) {
          iniciarSentinela();
          mostrarToast('🛡️ Modo Sentinela ativado.', 'info');
        } else {
          clearInterval(timerSentinela);
          mostrarToast('⏸️ Modo Sentinela pausado.', 'info');
        }
      });
    }

    if (temId) {
      iniciarSentinela();
    }
  }

  function iniciarSentinela() {
    if (pararSentinelaSeInvalido()) return;
    clearInterval(timerSentinela);
    setTimeout(() => {
      if (!pararSentinelaSeInvalido()) sincronizarComRadar(false);
    }, 3000);

    timerSentinela = setInterval(() => {
      if (pararSentinelaSeInvalido()) return;
      if (window.location.href !== ultimoUrlConhecido) {
        ultimoUrlConhecido = window.location.href;
        injetarWidget();
      }
      if (modoSentinelaAtivo) {
        sincronizarComRadar(false);
      }
    }, 60000);
  }

  function atualizarBadgeStatus(txt, tipo) {
    const el = document.getElementById('enterprisecoreStatusTxt');
    if (!el) return;
    el.textContent = txt;
    el.className = 'enterprisecore-txt-' + tipo;
  }

  function mostrarToast(msg, tipo = 'info') {
    let t = document.getElementById('enterprisecore-toast');
    if (!t) {
      t = document.createElement('div');
      t.id = 'enterprisecore-toast';
      t.className = 'enterprisecore-toast-box';
      document.body.appendChild(t);
    }
    t.textContent = msg;
    t.className = `enterprisecore-toast-box enterprisecore-toast-${tipo} show`;
    setTimeout(() => { t.className = 'enterprisecore-toast-box'; }, 4000);
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', injetarWidget);
  } else {
    injetarWidget();
  }
})();
