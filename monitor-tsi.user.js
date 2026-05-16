// ==UserScript==
// @name         Monitor Operacional TSI
// @namespace    http://tampermonkey.net/
// @version      21.0
// @description  Monitor de apontamentos em tempo real com escalados vs apontados
// @author       TSI
// @match        https://tsi-app.com/planejamento-operacional*
// @match        https://tsi-app.com/pedidoEapt*
// @grant        none
// @require      https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js
// @updateURL    https://raw.githubusercontent.com/brendosep164/monitor-tsi/main/monitor-tsi.user.js
// @downloadURL  https://raw.githubusercontent.com/brendosep164/monitor-tsi/main/monitor-tsi.user.js
// ==/UserScript==

(function() {
  'use strict';

  if (!window.location.href.includes('planejamento-operacional')) return;
  if (window.self !== window.top) return;

  const AVATAR_URL = 'https://i.imgur.com/9HPkbTi.png';

  const BG_IFRAME_IDS = ['_mon_ifr_A', '_mon_ifr_B', '_mon_ifr_C', '_mon_ifr_D'];
  const IFR_PAG2   = '_mon_pag2';
  const IFR_ESCALA = '_mon_escala';

  // ─── ESTADO ────────────────────────────────────────────────────────────────
  let iframesInUse  = {};
  let fetchQueue    = [];
  let inQueue       = new Set();
  let operations    = [];
  let apontCache    = {};
  let expanded      = new Set();
  let monitoradas   = new Set();
  let minimized     = false;
  let refreshTimer  = null;
  let watchdogTimer = null;

  // ── FILTRO / ORDENAÇÃO ──────────────────────────────────────────────────────
  let filterText  = "";
  let sortCol     = null;   // 'esc' | 'apt' | 'hora' | 'status'
  let sortDir     = 1;      // 1 = asc, -1 = desc
  // ops já notificadas — persiste no localStorage com chave por data (limpa automaticamente a cada dia)
  const _hoje = new Date().toISOString().slice(0, 10); // 'YYYY-MM-DD'
  const NOTIF_KEY = '_monNotificadas_' + _hoje;
  // Limpa chaves de dias anteriores
  (function() {
    try {
      Object.keys(localStorage).forEach(k => {
        if (k.startsWith('_monNotificadas_') && k !== NOTIF_KEY) localStorage.removeItem(k);
      });
    } catch(e) {}
  })();
  function notificadasLoad() { try { return new Set((JSON.parse(localStorage.getItem(NOTIF_KEY) || '[]')).map(String)); } catch(e) { return new Set(); } }
  function notificadasSave() { try { localStorage.setItem(NOTIF_KEY, JSON.stringify([...notificadas])); } catch(e) {} }
  let notificadas = notificadasLoad();

  // ── CACHE PERSISTENTE (sessionStorage) ──────────────────────────────────────
  const CACHE_KEY = '_monCache_v2';
  const CACHE_TTL = 5 * 60 * 1000;

  function cacheSave() {
    try {
      const payload = {};
      Object.entries(apontCache).forEach(([k, v]) => {
        if (v && v !== 'loading') payload[k] = { d: v, ts: Date.now() };
      });
      sessionStorage.setItem(CACHE_KEY, JSON.stringify(payload));
    } catch(e) {}
  }

  function cacheLoad() {
    try {
      const raw = sessionStorage.getItem(CACHE_KEY);
      if (!raw) return {};
      const payload = JSON.parse(raw);
      const now = Date.now();
      const out = {};
      Object.entries(payload).forEach(([k, v]) => {
        if (v && v.d && (now - v.ts) < CACHE_TTL) out[k] = v.d;
      });
      return out;
    } catch(e) { return {}; }
  }

  // ── IFRAMES ─────────────────────────────────────────────────────────────────
  function criarIfr(id) {
    if (document.getElementById(id)) return;
    const ifr = document.createElement('iframe');
    ifr.id = id;
    ifr.style.cssText = 'position:fixed;width:1px;height:1px;opacity:0;pointer-events:none;border:none;top:-9999px;left:-9999px;';
    document.body.appendChild(ifr);
  }

  function ensureIframes() {
    BG_IFRAME_IDS.forEach(criarIfr);
    criarIfr(IFR_PAG2);
    criarIfr(IFR_ESCALA);
  }

  function releaseIfr(ifrId) {
    delete iframesInUse[ifrId];
    const ifr = document.getElementById(ifrId);
    if (ifr) {
      ifr.onload = null;
      try { ifr.src = 'about:blank'; } catch(e) {}
    }
    setTimeout(processQueue, 50);
  }

  function startWatchdog() {
    if (watchdogTimer) clearInterval(watchdogTimer);
    watchdogTimer = setInterval(() => {
      const now = Date.now();
      Object.entries(iframesInUse).forEach(([key, info]) => {
        if (now - info.since > 35000) {
          console.warn('[Monitor] watchdog destravou', key, 'op:', info.opId);
          if (info.opId && apontCache[info.opId] === 'loading') {
            apontCache[info.opId] = {
              solicitado: 0, escalado: 0, apontado: 0,
              colaboradores: [], escalados: [], faltando: [],
              pdfLinks: [], xlsLinks: [], _erro: true
            };
          }
          delete iframesInUse[key];
          activeFetches = Math.max(0, activeFetches - 1);
        }
      });
      if (fetchQueue.length > 0 && activeFetches < MAX_CONCURRENT) {
        processQueue();
      }
    }, 8000);
  }

  // ── JANELA ──────────────────────────────────────────────────────────────────
  function dentroJanela(op) {
    if (!op.hora) return false;
    const [h, m] = op.hora.split(':').map(Number);
    if (isNaN(h)) return false;
    const agora    = new Date();
    const agoraMin = agora.getHours() * 60 + agora.getMinutes();
    const opMin    = h * 60 + (m || 0);
    let diffMin    = opMin - agoraMin;
    // Corrige virada de meia-noite apenas para ops futuras (evita que op de ontem
    // apareça como "daqui a pouco" quando diffMin ficaria positivo grande)
    if (diffMin > 180) diffMin -= 1440;
    // Ops passadas: sempre inclui (sem limite inferior)
    // Ops futuras: apenas até +3h
    return diffMin <= 180;
  }

  function monKey(op) { return op.id || op.chave; }
  function naJanela(op) { return monitoradas.has(monKey(op)) || dentroJanela(op); }

  // ── PARSER DE OPS ────────────────────────────────────────────────────────────
  function parseBubble(img) {
    if (!img) return null;
    const src = img.getAttribute('src') || '';
    const title = img.getAttribute('data-original-title') || img.getAttribute('title') || '';
    let status = 0;
    if (src.includes('statusbubble_1')) status = 1;
    else if (src.includes('statusbubble_2')) status = 2;
    else if (src.includes('statusbubble_3')) status = 3;
    else return null;
    return { status, title };
  }

  function parseOpsFromDoc(doc) {
    const ops = [];
    const mainTable = doc.querySelector('table tbody');
    if (!mainTable) return ops;
    mainTable.querySelectorAll('tr').forEach(row => {
      const cells = row.querySelectorAll('td');
      if (cells.length < 9) return;
      const allLinks = [...row.querySelectorAll('a[onclick*="planejamento-operacional-edit"]')];
      const linkEl = allLinks[0] || null;
      const linkElView = allLinks[1] || null; // olhinho = segundo link
      const eyeLink = row.querySelector('a[href*="pedidoEgeral"]') || row.querySelector('a[href*="pedidoE"]');
      let id = '';
      if (linkEl) {
        const match = linkEl.getAttribute('onclick').match(/planejamento-operacional-edit([A-Za-z0-9+\/=_-]+?)_\d+[',\s]/);
        if (match) id = match[1];
      }
      const g = i => cells[i]?.textContent?.trim() || '';
      const bubbles = [];
      row.querySelectorAll('img[src*="statusbubble_"]').forEach(img => {
        const b = parseBubble(img);
        if (b) bubbles.push(b);
      });
      const liderCell = cells[11];
      const lider = liderCell?.getAttribute('title') || liderCell?.getAttribute('data-original-title') || liderCell?.textContent?.trim() || '';
      window._monLinkEls = window._monLinkEls || {}; window._monLinkElsView = window._monLinkElsView || {}; window._monEyeHrefs = window._monEyeHrefs || {};
      if (linkEl && id) window._monLinkEls[id] = linkEl;
      if (linkElView && id) window._monLinkElsView[id] = linkElView;
      if (eyeLink && id) window._monEyeHrefs[id] = eyeLink.getAttribute('href');
      ops.push({ chave: g(0), sigla: g(1), site: g(2), qtd: parseInt(g(3)) || 0, hora: g(9), lider, liderCompleto: lider, status: g(24).toLowerCase(), time: g(8), id, bubbles });
    });
    return ops;
  }

  // ── NOTIFICAÇÕES ─────────────────────────────────────────────────────────────
  function pedirPermissaoNotificacao() {
    if (!('Notification' in window)) return;
    if (Notification.permission === 'granted') {
      new Notification('🔔 Monitor TSI', { body: 'Notificações já estão ativas!', icon: AVATAR_URL });
      return;
    }
    if (Notification.permission === 'denied') {
      alert('Notificações bloqueadas!\n\n1. Clique no cadeado\n2. Notificações → Permitir\n3. Recarregue');
      return;
    }
    Notification.requestPermission().then(perm => {
      const btn = document.getElementById('mon-notif-btn');
      if (perm === 'granted') {
        new Notification('✅ Monitor TSI ativado!', { body: 'Você receberá alertas.', icon: AVATAR_URL });
        if (btn) { btn.textContent = 'Notif. ativa'; btn.dataset.state = 'on'; }
      } else {
        if (btn) { btn.textContent = 'Bloqueado'; btn.dataset.state = 'off'; }
      }
    });
  }

  // Op considerada concluída quando todos os P1–P11 estão marcados como "Sim" no modal
  // Usa o campo todosConfirmados gravado no cache durante o fetch da operação
  function isConcluido(op) {
    const d = apontCache[op.id];
    if (!d || d === 'loading') return false;
    return d.todosConfirmados === true;
  }

  function notify(op, d) {
    if (!('Notification' in window) || Notification.permission !== 'granted') return;
    if ((op.time || '') !== 'VD') return;
    if (isConcluido(op)) return; // ← não notifica ops já concluídas
    try { new Notification('✅ Operação Completa — TSI', { body: `${op.sigla} | ${op.site}\n${d.apontado}/${d.solicitado} apontados`, icon: AVATAR_URL }); } catch(e) {}
  }

  window._monPedirNotif = pedirPermissaoNotificacao;

  // ── PROGRESSO ────────────────────────────────────────────────────────────────
  function updateProgress(loaded, total) {
    const circle  = document.getElementById('mon-progress-circle');
    const overlay = document.getElementById('mon-progress-overlay');
    const text    = document.getElementById('mon-progress-text');
    const img     = document.getElementById('mon-avatar-img');
    if (!circle || !overlay || !text) return;
    const pct  = total === 0 ? 100 : Math.round((loaded / total) * 100);
    const r    = 22, circ = 2 * Math.PI * r;
    circle.style.strokeDashoffset = circ - (pct / 100) * circ;
    if (pct >= 100) {
      circle.style.stroke   = 'var(--mon-green)';
      overlay.style.opacity = '0';
      text.style.display    = 'none';
      if (img) { img.style.animation = 'none'; img.style.transform = 'none'; }
    } else {
      circle.style.stroke   = 'var(--mon-amber)';
      overlay.style.opacity = '0.55';
      text.style.display    = 'flex';
      text.textContent      = pct + '%';
      if (img) img.style.animation = 'mon-shake 0.5s ease-in-out infinite alternate';
    }
  }

  // ── FETCH COM DOMPARSER (substitui iframes) ──────────────────────────────────
  // Como o script roda no contexto da página (mesma origem), fetch funciona direto.
  function fetchDoc(url) {
    return fetch(url, { credentials: 'include' })
      .then(r => {
        if (!r.ok) throw new Error('HTTP ' + r.status);
        return r.text();
      })
      .then(html => {
        const parser = new DOMParser();
        return parser.parseFromString(html, 'text/html');
      });
  }

  // ── FILA COM DEDUPLICAÇÃO ────────────────────────────────────────────────────
  // Concorrência controlada por semáforo (máx 4 simultâneos)
  let activeFetches = 0;
  const MAX_CONCURRENT = 4;

  function getAvailableIframe() {
    return BG_IFRAME_IDS.find(id => !iframesInUse[id]);
  }

  function loadUrl(ifr, url, timeout, onDone) {
    // Mantido apenas para compatibilidade com enviarEscala
    let fired = false;
    const fire = (result) => {
      if (fired) return;
      fired = true;
      clearTimeout(timer);
      ifr.onload = null;
      onDone(result);
    };
    const timer = setTimeout(() => fire(null), timeout);
    ifr.onload = function() {
      try {
        const href = ifr.contentDocument && ifr.contentDocument.location && ifr.contentDocument.location.href;
        if (!href || href === 'about:blank') return;
      } catch(e) {}
      setTimeout(() => {
        try { fire(ifr.contentDocument); }
        catch(e) { fire(null); }
      }, 1800);
    };
    try { ifr.src = url; } catch(e) { fire(null); }
  }

  function processQueue() {
    if (fetchQueue.length === 0) return;
    if (activeFetches >= MAX_CONCURRENT) return;

    const { op, callback } = fetchQueue.shift();
    inQueue.delete(op.id);

    if (!op.id) { setTimeout(processQueue, 10); return; }

    activeFetches++;
    // Marca um slot de iframe como "em uso" só pro watchdog continuar funcionando
    const fakeId = '_fetch_' + op.id;
    iframesInUse[fakeId] = { opId: op.id, since: Date.now() };

    const oldCache = (apontCache[op.id] && apontCache[op.id] !== 'loading') ? apontCache[op.id] : null;

    const release = (dados) => {
      activeFetches--;
      delete iframesInUse[fakeId];
      if (dados !== null) apontCache[op.id] = dados;
      callback(apontCache[op.id], oldCache);
      updateMetrics();
      setTimeout(processQueue, 30);
    };

    const fallback = (escalados) => {
      release({
        solicitado: op.qtd, escalado: escalados.length, apontado: 0,
        colaboradores: [], escalados: escalados || [],
        faltando: escalados || [], pdfLinks: [], xlsLinks: []
      });
    };

    if (!apontCache[op.id]) apontCache[op.id] = 'loading';

    // 1) Busca a página da operação para pegar os links de escala e apontamento
    fetchDoc('https://tsi-app.com/planejamento-operacional-edit' + op.id + '_1')
      .then(doc => {
        let listaEnviada = false, todosConfirmados = false;
        try {
          let etapa = 0, confirmadas = 0;
          doc.querySelectorAll('table tbody tr').forEach(row => {
            const radios = row.querySelectorAll('input[type="radio"]');
            if (radios.length >= 2) { etapa++; if (radios[0].checked) confirmadas++; }
          });
          listaEnviada     = etapa >= 8  && confirmadas >= 8;
          todosConfirmados = etapa >= 11 && confirmadas >= 11;
        } catch(e) {}

        let escalaHref, eaptHref;
        const escalaLink = doc.querySelector('a[href*="pedidoEescala"]');
        const eaptLink   = doc.querySelector('a[href*="pedidoEapt"]');
        if (escalaLink && eaptLink) {
          escalaHref = escalaLink.getAttribute('href');
          eaptHref   = eaptLink.getAttribute('href');
        } else {
          const eg = doc.querySelector('a[href*="pedidoEgeral"]');
          if (!eg) { fallback([]); return; }
          escalaHref = eg.getAttribute('href').replace('pedidoEgeral', 'pedidoEescala');
          eaptHref   = eg.getAttribute('href').replace('pedidoEgeral', 'pedidoEapt');
        }

        // 2) Busca escala
        return fetchDoc('https://tsi-app.com/' + escalaHref)
          .then(doc2 => {
            const escalados = [], pdfLinks = [], xlsLinks = [];
            const tbl = doc2.querySelector('table.tables.table-fixed.card-table.table-bordered');
            if (tbl) {
              tbl.querySelectorAll('tbody tr').forEach(row => {
                if (row.classList.contains('strikethrough')) return;
                if (row.querySelector('td.strikethrough')) return;
                const cells = row.querySelectorAll('td');
                if (cells.length < 5) return;
                const nome    = cells[2]?.textContent?.trim();
                const cpf     = cells[3]?.textContent?.trim();
                if (!nome || nome.length < 3) return;
                const datacad = cells[12]?.textContent?.trim() || '';
                const advisor = cells[13]?.textContent?.trim() || '';
                escalados.push({ nome, cpf, tipo: cells[4]?.textContent?.trim(), datacad, advisor });
              });
            }
            const xlsLabels = ['Layout 1 (SHEIN)','Layout 2 (Cordovil)','Layout 3 (SBC)','Layout 4 (SBF)','Layout 5 (Endereço)','Layout 6 (KISOC)'];
            // Lider: extrai o nome completo da coluna advisor (col 13) da tabela de escala
            const lideresSet = new Set();
            escalados.forEach(e => { if (e.advisor && e.advisor.length > 2) lideresSet.add(e.advisor); });
            const lideres = lideresSet.size > 0
              ? [...lideresSet]
              : (op.liderCompleto ? [op.liderCompleto] : (op.lider ? [op.lider] : []));
            doc2.querySelectorAll('a[href*="escalaprelistaLiderPDF_"]').forEach(a => {
              pdfLinks.push({ label: a.textContent.trim() || 'Lista p/ Assinaturas', href: a.getAttribute('href') });
            });
            doc2.querySelectorAll('a[href*="escalaprelistaLiderXLS"]').forEach((a, i) => xlsLinks.push({ label: xlsLabels[i] || a.textContent.trim(), href: a.getAttribute('href') }));

            if (!naJanela(op)) {
              release({ solicitado: op.qtd, escalado: escalados.length, apontado: 0, colaboradores: [], escalados, faltando: escalados, pdfLinks, xlsLinks, _soEscala: true, listaEnviada, todosConfirmados, lideres });
              return;
            }

            // 3) Busca apontamentos
            return fetchDoc('https://tsi-app.com/' + eaptHref)
              .then(doc3 => {
                const colaboradores = [], faltasConfirmadas = [];
                const tbl3 = doc3.querySelector('table.tables.table-fixed.card-table:not(.table-bordered)');
                if (tbl3) {
                  tbl3.querySelectorAll('tbody tr').forEach(row => {
                    const cells = row.querySelectorAll('td');
                    if (cells.length < 10) return;
                    const nome   = cells[0]?.textContent?.trim();
                    const cpf    = cells[1]?.textContent?.trim();
                    const origem = cells[9]?.textContent?.trim();
                    const inicio = cells[8]?.textContent?.trim();
                    if (!nome || nome.length < 3) return;
                    const isFalta = origem === 'FALTA';
                    if (isFalta) { faltasConfirmadas.push({ nome, cpf, tipo: cells[2]?.textContent?.trim(), inicio }); return; }
                    if (!inicio) return;
                    // DIST do INÍCIO OPORTUNIDADE = coluna 10, formato "0,08 km" ou "399,56 km"
                    const distRaw = cells[10]?.textContent?.trim() || '';
                    const distMatch = distRaw.match(/([\d.,]+)\s*km/i);
                    const distNum = distMatch ? parseFloat(distMatch[1].replace(',', '.')) : 0;
                    colaboradores.push({ nome, cpf, tipo: cells[2]?.textContent?.trim(), inicio, dist: distNum });
                  });
                }
                const apontadosCPF = new Set(colaboradores.map(c => c.cpf));
                const faltando = escalados.filter(e => !apontadosCPF.has(e.cpf));
                release({ solicitado: op.qtd, escalado: escalados.length, apontado: colaboradores.length, colaboradores, escalados, faltando, faltasConfirmadas, pdfLinks, xlsLinks, listaEnviada, todosConfirmados, eaptHref, lideres });
              });
          });
      })
      .catch(() => fallback([]));
  }

  function enfileirar(op, callback, force) {
    if (!op.id) return;
    if (inQueue.has(op.id)) return;
    if (!force && apontCache[op.id] === 'loading') return;
    if (!force && apontCache[op.id] && apontCache[op.id] !== 'loading') return;
    inQueue.add(op.id);
    fetchQueue.push({ op, callback });
    processQueue();
  }

  function fetchApontamentos(op, callback) {
    if (!op.id) return;
    const cached = apontCache[op.id];
    if (cached !== undefined && cached !== 'loading' && cached !== null) {
      callback(cached, null);
      return;
    }
    enfileirar(op, callback);
  }

  // ── ENVIAR ESCALA ─────────────────────────────────────────────────────────────
  function enviarEscala(opId, btnEl) {
    const ifr = document.getElementById(IFR_ESCALA);
    if (!ifr || !opId) return;

    const origTxt = btnEl.innerHTML;
    btnEl.disabled = true;
    btnEl.innerHTML = 'Enviando…';
    btnEl.style.opacity = '0.6';

    let done = false;
    const fail = (msg) => {
      if (done) return; done = true;
      btnEl.disabled = false;
      btnEl.innerHTML = '✗ ' + msg;
      btnEl.className = btnEl.className.replace('mon-send-btn', 'mon-send-btn mon-send-btn--err');
      btnEl.style.opacity = '1';
      setTimeout(() => { btnEl.innerHTML = origTxt; btnEl.className = btnEl.className.replace(' mon-send-btn--err', ''); }, 3000);
    };
    const safetyTimer = setTimeout(() => fail('timeout'), 25000);

    ifr.onload = null;
    ifr.src = 'https://tsi-app.com/planejamento-operacional-edit' + opId + '_1';
    ifr.onload = function() {
        if (done) return;
        try {
          const doc = ifr.contentDocument;
          if (!doc || !doc.body) { fail('modal vazio'); clearTimeout(safetyTimer); return; }
          let marcados = 0;
          for (let i = 1; i <= 8; i++) {
            const radio = doc.querySelector('input[name="p' + i + '_confirm"][value="S"]');
            if (radio) {
              radio.checked = true;
              radio.dispatchEvent(new Event('change', { bubbles: true }));
              radio.dispatchEvent(new Event('click',  { bubbles: true }));
              marcados++;
            }
          }
          if (marcados === 0) { fail('radios não encontrados'); clearTimeout(safetyTimer); return; }
          setTimeout(() => {
            if (done) return;
            const saveBtn = doc.querySelector('button[name="submitF"]');
            if (!saveBtn) { fail('btn salvar não encontrado'); clearTimeout(safetyTimer); return; }
            saveBtn.click();
            setTimeout(() => {
              if (done) return;
              done = true;
              clearTimeout(safetyTimer);
              btnEl.innerHTML = '✓ Enviado!';
              btnEl.style.opacity = '1';
              setTimeout(() => window.location.reload(), 1500);
            }, 500);
          }, 50);
        } catch(e) { fail('erro'); clearTimeout(safetyTimer); }
    };
  }

  window._monEnviarEscala = enviarEscala;
  window._monEnviarReport = enviarReport;

  // ── ENVIAR REPORT ─────────────────────────────────────────────────────────────
  function enviarReport(opId, btnEl) {
    const ifr = document.getElementById(IFR_ESCALA);
    if (!ifr || !opId) return;

    // Quantidade de apontados já em cache para esta operação
    const d = apontCache[opId];
    const qtdApontados = (d && d !== 'loading' && d.apontado != null) ? d.apontado : 0;

    const origTxt = btnEl.innerHTML;
    btnEl.disabled = true;
    btnEl.innerHTML = 'Enviando…';
    btnEl.style.opacity = '0.6';

    let done = false;
    const fail = (msg) => {
      if (done) return; done = true;
      btnEl.disabled = false;
      btnEl.innerHTML = '✗ ' + msg;
      btnEl.className = btnEl.className.replace('mon-send-btn', 'mon-send-btn mon-send-btn--err');
      btnEl.style.opacity = '1';
      setTimeout(() => { btnEl.innerHTML = origTxt; btnEl.className = btnEl.className.replace(' mon-send-btn--err', ''); }, 3000);
    };
    const safetyTimer = setTimeout(() => fail('timeout'), 25000);

    ifr.onload = null;
    ifr.src = 'https://tsi-app.com/planejamento-operacional-edit' + opId + '_1';
    ifr.onload = function() {
        if (done) return;
        try {
          const doc = ifr.contentDocument;
          if (!doc || !doc.body) { fail('modal vazio'); clearTimeout(safetyTimer); return; }

          let marcados = 0;

          // Marca P1 a P11 com "S"
          for (let i = 1; i <= 11; i++) {
            const radio = doc.querySelector('input[name="p' + i + '_confirm"][value="S"]');
            if (radio) {
              radio.checked = true;
              radio.dispatchEvent(new Event('change', { bubbles: true }));
              radio.dispatchEvent(new Event('click',  { bubbles: true }));
              marcados++;
            }
          }

          if (marcados === 0) { fail('radios não encontrados'); clearTimeout(safetyTimer); return; }

          // P10: preenche também o campo de quantidade de apontamentos
          // Tenta selectors comuns para o campo de quantidade do P10
          const p10QuantSelectors = [
            'input[name="p10_qtd"]',
            'input[name="p10_quantidade"]',
            'input[name="p10_quant"]',
            'input[name="p10_valor"]',
          ];
          let p10Input = null;
          for (const sel of p10QuantSelectors) {
            p10Input = doc.querySelector(sel);
            if (p10Input) break;
          }
          // Fallback: procura input de texto/number dentro da mesma linha do radio do p10
          if (!p10Input) {
            const p10Radio = doc.querySelector('input[name="p10_confirm"]');
            if (p10Radio) {
              const row = p10Radio.closest('tr') || p10Radio.closest('div');
              if (row) {
                p10Input = row.querySelector('input[type="text"], input[type="number"]');
              }
            }
          }
          if (p10Input) {
            p10Input.value = qtdApontados;
            p10Input.dispatchEvent(new Event('input',  { bubbles: true }));
            p10Input.dispatchEvent(new Event('change', { bubbles: true }));
          }

          setTimeout(() => {
            if (done) return;
            const saveBtn = doc.querySelector('button[name="submitF"]');
            if (!saveBtn) { fail('btn salvar não encontrado'); clearTimeout(safetyTimer); return; }
            saveBtn.click();
            setTimeout(() => {
              if (done) return;
              done = true;
              clearTimeout(safetyTimer);
              btnEl.innerHTML = '✓ Enviado!';
              btnEl.style.opacity = '1';
              setTimeout(() => window.location.reload(), 1500);
            }, 500);
          }, 50);
        } catch(e) { fail('erro'); clearTimeout(safetyTimer); }
    };
  }

  // ── GERAR RELATÓRIO WHATSAPP ──────────────────────────────────────────────────
  window._monGerarRelatorio = function(opId, btnEl) {
    const d = apontCache[opId];
    if (!d || d === 'loading') { alert('Aguarde os dados carregarem.'); return; }
    const op = operations.find(o => o.id === opId);
    if (!op) return;

    // Colaboradores entregues = apontados com dist <= 2km
    // dist=0 significa que não encontrou dado de distância — conta normalmente
    const entregues = (d.colaboradores || []).filter(c => c.dist === 0 || c.dist <= 2);
    const entregueCount = entregues.length;

    // Líderes: prioridade = lideres da escala; fallback = op.liderCompleto ou op.lider
    const lideres = (d.lideres && d.lideres.length > 0)
      ? d.lideres
      : (op.liderCompleto ? [op.liderCompleto] : (op.lider ? [op.lider] : ['—']));

    // Data de hoje
    const hoje = new Date();
    const dia  = String(hoje.getDate()).padStart(2,'0');
    const mes  = String(hoje.getMonth()+1).padStart(2,'0');
    const ano  = hoje.getFullYear();

    // Nome completo do líder: prioriza liderCompleto, fallback para lider
    const lideresCompletos = lideres.map(l => {
      if (l && l !== '—') return l;
      return op.liderCompleto || op.lider || '—';
    });

    const reportAtualizado = d.todosConfirmados === true;
    const tituloReport = reportAtualizado
      ? `*💚 REPORT ATUALIZADO - TIME VERDE 💚*`
      : `*💚 REPORT - TIME VERDE 💚*`;

    const texto = [
      tituloReport,
      ``,
      `📅 ${dia}/${mes}/${ano}`,
      ``,
      `*CHAVE DA OPERAÇÃO:* ${op.chave}`,
      `*HORÁRIO:* ${op.hora || '—'}`,
      ``,
      `*SOLICITADO:* ${String(op.qtd).padStart(2,'0')}`,
      `*ENTREGUE:* ${String(entregueCount).padStart(2,'0')}`,
      ``,
      `*LÍDER:* ${lideresCompletos.join(' / ')}`,
    ].join('\n');

    navigator.clipboard.writeText(texto)
      .then(() => {
        const orig = btnEl.innerHTML;
        btnEl.innerHTML = '✅ Copiado!';
        btnEl.style.color = 'var(--mon-green)';
        setTimeout(() => { btnEl.innerHTML = orig; btnEl.style.color = ''; }, 2500);
      })
      .catch(() => {
        // Fallback: abre prompt com o texto para copiar manualmente
        prompt('Copie o relatório:', texto);
      });
  };



  // ── PRINT APONTAMENTOS ────────────────────────────────────────────────────────
  // Carrega a página de apontamentos em iframe oculto (mesmo domínio = sem CORS),
  // aplica html2canvas direto no documento do iframe e baixa/copia o PNG.
  window._monPrintApontamentos = function(opId, btnEl) {
    const d = apontCache[opId];
    if (!d || d === 'loading') { alert('Aguarde os dados carregarem.'); return; }
    const href = d.eaptHref;
    if (!href) { alert('Clique em ↻ Atualizar para habilitar o print desta operação.'); return; }

    const origHTML = btnEl.innerHTML;
    btnEl.disabled = true;
    btnEl.innerHTML = '⏳ Carregando…';

    const fail = (msg) => {
      btnEl.innerHTML = '✗ ' + msg;
      setTimeout(() => { btnEl.innerHTML = origHTML; btnEl.disabled = false; }, 3500);
    };

    // Cria iframe oculto com tamanho real para renderização correta
    const printIfr = document.createElement('iframe');
    printIfr.id = '_mon_print_ifr';
    printIfr.style.cssText = [
      'position:fixed',
      'top:-9999px',
      'left:-9999px',
      'width:1400px',
      'height:900px',
      'opacity:0',
      'pointer-events:none',
      'border:none',
      'z-index:-1',
    ].join(';');
    document.body.appendChild(printIfr);

    const cleanup = () => { try { document.body.removeChild(printIfr); } catch(e) {} };

    const safetyTimeout = setTimeout(() => { cleanup(); fail('Timeout ao carregar'); }, 30000);

    printIfr.onload = () => {
      // Aguarda o DOM do iframe renderizar completamente
      setTimeout(() => {
        try {
          const ifrDoc = printIfr.contentDocument || printIfr.contentWindow.document;
          const ifrWin = printIfr.contentWindow;

          // Injeta html2canvas no iframe (mesmo domínio, funciona sem CORS)
          const script = ifrDoc.createElement('script');
          script.src = 'https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js';
          script.onload = () => {
            setTimeout(() => {
              try {
                // ── Acha o contador "X apontamento(s) registrado(s)"
                // Usa TreeWalker nos nós de TEXTO para achar o nó mais específico,
                // evitando pegar o textContent herdado de ancestrais (que engloba a página toda).
                let counterEl = null;
                const walker = ifrDoc.createTreeWalker(ifrDoc.body, NodeFilter.SHOW_TEXT);
                let node;
                while ((node = walker.nextNode())) {
                  if (/\d+\s*apontamento.*registrado/i.test(node.nodeValue)) {
                    // Sobe até incluir badge FALTAS mas sem englobar a tabela
                    let t = node.parentElement;
                    for (let i = 0; i < 4; i++) {
                      if (!t || !t.parentElement || t.parentElement === ifrDoc.body) break;
                      const p = t.parentElement;
                      if (p.querySelector('table')) break;
                      t = p;
                    }
                    counterEl = t;
                    break;
                  }
                }

                // ── Acha a tabela principal
                const tbl = ifrDoc.querySelector('table.tables.table-fixed.card-table:not(.table-bordered)')
                         || ifrDoc.querySelector('table.card-table')
                         || ifrDoc.querySelector('table');

                if (!tbl) { clearTimeout(safetyTimeout); cleanup(); fail('Tabela não encontrada'); return; }

                // ── Encontra coluna DATA para saber onde cortar à direita
                const theadRows = Array.from(tbl.querySelectorAll('thead tr'));
                let dataThEl = null;
                if (theadRows[1]) {
                  for (const cell of theadRows[1].querySelectorAll('th,td')) {
                    if (/^DATA$/i.test(cell.textContent.trim())) { dataThEl = cell; break; }
                  }
                }
                if (!dataThEl && theadRows[0]) {
                  for (const cell of theadRows[0].querySelectorAll('th,td')) {
                    if (/^DATA$/i.test(cell.textContent.trim())) { dataThEl = cell; break; }
                  }
                }

                // ── Calcula área de recorte
                const scrollX = ifrWin.scrollX || 0;
                const scrollY = ifrWin.scrollY || 0;
                // Se não achou o contador, usa o topo da tabela menos 40px para garantir espaço
                const counterRect = counterEl
                  ? counterEl.getBoundingClientRect()
                  : { top: tbl.getBoundingClientRect().top - 40, left: tbl.getBoundingClientRect().left };
                const tblRect    = tbl.getBoundingClientRect();
                const lastRow    = tbl.querySelector('tbody tr:last-child');
                const bottomRect = lastRow ? lastRow.getBoundingClientRect() : tblRect;
                const dataRect   = dataThEl ? dataThEl.getBoundingClientRect() : null;

                const cropTop    = counterRect.top  - 10 + scrollY;
                const cropLeft   = counterRect.left - 6  + scrollX;
                const cropBottom = bottomRect.bottom + 6 + scrollY;
                const cropRight  = dataRect ? dataRect.right + 6 + scrollX : tblRect.right + 6 + scrollX;

                btnEl.innerHTML = '📸 Gerando…';

                // ── Captura com html2canvas diretamente no documento do iframe
                ifrWin.html2canvas(ifrDoc.body, {
                  backgroundColor: '#ffffff',
                  scale: 2,
                  useCORS: true,
                  allowTaint: true,
                  logging: false,
                  scrollX: 0,
                  scrollY: 0,
                  windowWidth: ifrDoc.body.scrollWidth,
                  windowHeight: ifrDoc.body.scrollHeight,
                }).then(canvas => {
                  clearTimeout(safetyTimeout);

                  const scale = 2;
                  const sx = Math.round(cropLeft  * scale);
                  const sy = Math.round(Math.max(0, cropTop) * scale);
                  const sw = Math.max(1, Math.round((cropRight  - cropLeft) * scale));
                  const sh = Math.max(1, Math.round((cropBottom - cropTop)  * scale));

                  const out = document.createElement('canvas');
                  out.width  = sw;
                  out.height = sh;
                  out.getContext('2d').drawImage(canvas, sx, sy, sw, sh, 0, 0, sw, sh);

                  const dataUrl = out.toDataURL('image/png');

                  // ── Tenta clipboard; se falhar faz download automático
                  const doDownload = () => {
                    const a = document.createElement('a');
                    a.href = dataUrl;
                    const opChave = (operations.find(o => o.id === opId) || {}).chave || opId;
                    a.download = 'apontamentos_' + opChave + '_' + new Date().toISOString().slice(0,16).replace(/[T:]/g,'-') + '.png';
                    document.body.appendChild(a);
                    a.click();
                    document.body.removeChild(a);
                  };

                  const onSuccess = () => {
                    btnEl.innerHTML = '✅ Print copiado!';
                    btnEl.style.color = 'var(--mon-green)';
                    setTimeout(() => { btnEl.innerHTML = origHTML; btnEl.style.color = ''; btnEl.disabled = false; }, 2500);
                  };
                  const onDownloaded = () => {
                    btnEl.innerHTML = '✅ Baixado!';
                    btnEl.style.color = 'var(--mon-green)';
                    setTimeout(() => { btnEl.innerHTML = origHTML; btnEl.style.color = ''; btnEl.disabled = false; }, 2500);
                  };

                  if (navigator.clipboard && navigator.clipboard.write) {
                    fetch(dataUrl).then(r => r.blob()).then(blob => {
                      navigator.clipboard.write([new ClipboardItem({ 'image/png': blob })])
                        .then(onSuccess)
                        .catch(() => { doDownload(); onDownloaded(); });
                    }).catch(() => { doDownload(); onDownloaded(); });
                  } else {
                    doDownload();
                    onDownloaded();
                  }

                  cleanup();
                }).catch(e => {
                  clearTimeout(safetyTimeout);
                  cleanup();
                  fail('Erro canvas: ' + e.message);
                });
              } catch(e) {
                clearTimeout(safetyTimeout);
                cleanup();
                fail('Erro: ' + e.message);
              }
            }, 100);
          };
          script.onerror = () => { clearTimeout(safetyTimeout); cleanup(); fail('Erro ao carregar html2canvas'); };
          ifrDoc.head.appendChild(script);
        } catch(e) {
          clearTimeout(safetyTimeout);
          cleanup();
          fail('Erro de acesso ao iframe: ' + e.message);
        }
      }, 300);
    };

    printIfr.src = 'https://tsi-app.com/' + href;
  };

  // ── FILTRO / SORT ─────────────────────────────────────────────────────────────
  let activeStatusFilter = 'all';

  function getStatusRank(op) {
    const d = apontCache[op.id];
    if (!d || d === 'loading') return 99;
    if (d.apontado >= d.solicitado && d.escalado >= d.solicitado) return 0; // completo
    if (d.apontado > 0) return 1;  // parcial
    if (d.escalado >= d.solicitado) return 2; // esc ok
    if (d.escalado > 0) return 3;  // esc parcial
    return 4; // nenhum
  }

  function matchesStatusFilter(op) {
    if (activeStatusFilter === 'all') return true;
    const d = apontCache[op.id];
    if (!d || d === 'loading') return activeStatusFilter === 'nenhum';
    const escOk = d.escalado >= d.solicitado;
    const aptOk = d.apontado >= d.solicitado;
    if (activeStatusFilter === 'completo') return aptOk && escOk;
    if (activeStatusFilter === 'parcial')  return d.apontado > 0 && !aptOk;
    if (activeStatusFilter === 'esc')      return d.apontado === 0 && escOk;
    if (activeStatusFilter === 'esc_inc')  return d.escalado > 0 && !escOk;
    if (activeStatusFilter === 'nenhum')   return d.apontado === 0 && d.escalado === 0;
    return true;
  }

  function getVisibleOps() {
    const q = filterText.toLowerCase().trim();
    let ops = operations.filter(op => {
      if (q && !op.chave.toLowerCase().includes(q) &&
               !op.sigla.toLowerCase().includes(q) &&
               !(op.site||'').toLowerCase().includes(q) &&
               !(op.lider||'').toLowerCase().includes(q)) return false;
      return matchesStatusFilter(op);
    });

    if (sortCol) {
      ops = [...ops].sort((a, b) => {
        let va, vb;
        const da = apontCache[a.id], db = apontCache[b.id];
        if (sortCol === 'esc') {
          va = (da && da !== 'loading' && a.qtd > 0) ? da.escalado / a.qtd : -1;
          vb = (db && db !== 'loading' && b.qtd > 0) ? db.escalado / b.qtd : -1;
        } else if (sortCol === 'apt') {
          va = (da && da !== 'loading' && a.qtd > 0) ? da.apontado / a.qtd : -1;
          vb = (db && db !== 'loading' && b.qtd > 0) ? db.apontado / b.qtd : -1;
        } else if (sortCol === 'hora') {
          va = a.hora || '99:99'; vb = b.hora || '99:99';
        } else if (sortCol === 'status') {
          va = getStatusRank(a); vb = getStatusRank(b);
        }
        if (va < vb) return -1 * sortDir;
        if (va > vb) return  1 * sortDir;
        return 0;
      });
    }
    return ops;
  }

  window._monSetFilter = function(val) {
    filterText = val;
    const clearBtn = document.getElementById('mon-filter-clear');
    if (clearBtn) clearBtn.style.display = val ? 'block' : 'none';
    renderTable();
  };

  window._monClearFilter = function() {
    filterText = '';
    const inp = document.getElementById('mon-filter-input');
    if (inp) inp.value = '';
    const clearBtn = document.getElementById('mon-filter-clear');
    if (clearBtn) clearBtn.style.display = 'none';
    renderTable();
  };

  window._monSetStatusFilter = function(val, btnEl) {
    activeStatusFilter = val;
    document.querySelectorAll('.mon-chip').forEach(b => b.classList.remove('active'));
    if (btnEl) btnEl.classList.add('active');
    renderTable();
  };

  window._monToggleSort = function(col, thEl) {
    if (sortCol === col) {
      sortDir = sortDir * -1;
    } else {
      sortCol = col;
      sortDir = 1;
    }
    // Atualiza visual das colunas
    document.querySelectorAll('.mon-th-sort').forEach(th => {
      th.classList.remove('sort-asc', 'sort-desc');
    });
    if (thEl) thEl.classList.add(sortDir === 1 ? 'sort-asc' : 'sort-desc');
    renderTable();
  };

  // ── ATUALIZAÇÃO MANUAL ────────────────────────────────────────────────────────
  function manualRefresh() {
    if (refreshTimer) clearInterval(refreshTimer);
    if (_alignTimeout) clearTimeout(_alignTimeout);
    monCarregarContatos();
    iframesInUse = {};
    fetchQueue   = [];
    inQueue      = new Set();
    apontCache   = {};
    try { sessionStorage.removeItem(CACHE_KEY); } catch(e) {}
    BG_IFRAME_IDS.forEach(id => {
      const ifr = document.getElementById(id);
      if (ifr) { ifr.onload = null; try { ifr.src = 'about:blank'; } catch(e) {} }
    });
    fetchOperations();
    scheduleAlignedRefresh();
  }
  window._monRefresh = manualRefresh;

  // ── ATUALIZAÇÃO SILENCIOSA ────────────────────────────────────────────────────
  function silentRefresh() {
    // Usa apenas as ops da página atual que está sendo visualizada
    const ops = parseOpsFromDoc(document);
    if (ops.length === 0) return;

    const oldIds  = new Set(operations.map(o => o.id));
    const newIds  = new Set(ops.map(o => o.id));

    operations.filter(o => !newIds.has(o.id)).forEach(o => {
      delete apontCache[o.id];
      expanded.delete(o.chave);
      monitoradas.delete(monKey(o));
    });

    ops.forEach(o => { if (dentroJanela(o)) monitoradas.add(monKey(o)); });
    operations = ops;
    renderTable();

    ops.filter(o => o.id).forEach((op, i) => {
      setTimeout(() => {
        const cached = apontCache[op.id];
        const emJanela = naJanela(op);

        if (cached === 'loading') return;

        if (!oldIds.has(op.id)) {
          monitoradas.add(monKey(op));
          enfileirar(op, (novo, old) => {
            updateCells(op, novo, old);
            updateMetrics();
          });
          return;
        }

        // Se tinha cache só de escala mas agora entrou na janela, invalida e rebusca com apontamentos
        const eraSoEscala = cached && cached !== 'loading' && cached._soEscala;
        if (emJanela || eraSoEscala) {
          delete apontCache[op.id];
          if (emJanela) monitoradas.add(monKey(op));
          enfileirar(op, (novo, old) => {
            updateCells(op, novo, old);
            updateMetrics();
            cacheSave();
            if (expanded.has(op.chave)) {
              const idx = operations.findIndex(o => o.chave === op.chave);
              const det = document.getElementById('det-' + idx);
              if (det) det.querySelector('.mon-detail-inner').innerHTML = renderDetail(op);
            }
          });
        } else {
          // Rebusca se não tem cache, tem erro, ou tem só escala
          if (!cached || cached._erro || cached._soEscala) {
            enfileirar(op, (novo) => {
              updateCells(op, novo, null);
              cacheSave();
            });
          } else {
            // Já tem cache completo: apenas reavalia notificação sem rebuscar
            updateCells(op, cached, null);
          }
        }
      }, i * 250);
    });

    const sub = document.getElementById('mon-sub');
    if (sub) sub.textContent = 'Atualizado ' + new Date().toLocaleTimeString('pt-BR');
  }

  function updateCells(op, d, old) {
    if (!d || d === 'loading') return;
    const row = document.querySelector(`tr[data-chave="${op.chave}"]`);
    if (!row) return;
    const cells = row.querySelectorAll('td');
    if (cells[2]) cells[2].innerHTML = escaladoBadge(d, op.qtd);
    if (naJanela(op)) {
      if (cells[3]) cells[3].innerHTML = apontBadge(d, op.qtd);
    }
    if (cells[6]) cells[6].innerHTML = situacaoBadge(d, op) + escalaEnviadaBadge(op);

    const _nid = String(op.id);

    // Se a op voltou ao estado ativo (não concluída) mas ainda não completou,
    // remove do set para poder notificar novamente quando completar
    if (!isConcluido(op) && notificadas.has(_nid)) {
      const jaCompleta = d && typeof d.apontado === 'number' && typeof d.solicitado === 'number'
                         && d.solicitado > 0 && d.apontado >= d.solicitado;
      if (!jaCompleta) {
        notificadas.delete(_nid);
        notificadasSave();
      }
    }

    // Notifica apenas uma vez quando completar — somente se planejamento ou em andamento (não concluída)
    const _completa = d && d !== 'loading' && !d._erro &&
                      typeof d.apontado === 'number' && typeof d.solicitado === 'number' &&
                      d.solicitado > 0 && d.apontado >= d.solicitado;
    if (_completa && !isConcluido(op) && !notificadas.has(_nid)) {
      notify(op, d);
      notificadas.add(_nid);
      notificadasSave();
    }
  }

  // ── DRAG / RESIZE / MINIMIZE ──────────────────────────────────────────────────
  // Shield único reutilizável — evita múltiplos overlays empilhados ou esquecidos
  let _monShield = null;
  function _showShield(cursor) {
    if (!_monShield) {
      _monShield = document.createElement('div');
      _monShield.style.cssText = 'position:fixed;inset:0;z-index:199999;';
      document.body.appendChild(_monShield);
    }
    _monShield.style.cursor = cursor;
  }
  function _hideShield() {
    if (_monShield) { _monShield.remove(); _monShield = null; }
    document.body.style.userSelect = '';
  }

  // Altura salva antes de minimizar para restaurar corretamente
  let _savedPanelHeight = '';

  function initControls(panel) {
    panel.style.position = 'fixed';
    let isDocked = true; // painel começa ancorado na direita via CSS

    // ── RESIZE (alça esquerda) ──────────────────────────────────────────────
    const rh = document.createElement('div');
    rh.className = 'mon-resize-handle';
    rh.addEventListener('mousedown', e => {
      e.preventDefault();
      e.stopPropagation();
      const startX = e.clientX, startW = panel.offsetWidth;
      document.body.style.userSelect = 'none';
      _showShield('ew-resize');
      const mv = e => {
        const newW = Math.min(Math.max(startW + (startX - e.clientX), 380), window.innerWidth - 40);
        panel.style.width = newW + 'px';
      };
      const up = () => {
        _hideShield();
        window.removeEventListener('mousemove', mv);
        window.removeEventListener('mouseup', up);
      };
      window.addEventListener('mousemove', mv);
      window.addEventListener('mouseup', up);
    });
    panel.appendChild(rh);

    // ── DRAG (header) ───────────────────────────────────────────────────────
    const header = panel.querySelector('#mon-header');
    if (!header) return;

    header.addEventListener('mousedown', e => {
      if (e.target.tagName === 'BUTTON' || e.target.closest('button')) return;
      e.preventDefault();

      // Converte de right-anchored para posição absoluta na primeira arrastada
      if (isDocked) {
        const r = panel.getBoundingClientRect();
        panel.style.left  = r.left + 'px';
        panel.style.top   = r.top  + 'px';
        panel.style.right = 'auto';
        isDocked = false;
      }

      const r  = panel.getBoundingClientRect();
      const ox = e.clientX - r.left;
      const oy = e.clientY - r.top;
      document.body.style.userSelect = 'none';
      header.style.cursor = 'grabbing';
      _showShield('grabbing');

      const mv = e => {
        panel.style.left = Math.max(0, Math.min(e.clientX - ox, window.innerWidth  - panel.offsetWidth))  + 'px';
        panel.style.top  = Math.max(0, Math.min(e.clientY - oy, window.innerHeight - 60)) + 'px';
      };
      const up = () => {
        header.style.cursor = '';
        _hideShield();
        window.removeEventListener('mousemove', mv);
        window.removeEventListener('mouseup', up);
      };
      window.addEventListener('mousemove', mv);
      window.addEventListener('mouseup', up);
    });
  }

  window._monMinimize = function() {
    const body  = document.getElementById('mon-body');
    const btn   = document.getElementById('mon-min-btn');
    const panel = document.getElementById('mon-panel');
    if (!body || !btn || !panel) return;
    minimized = !minimized;
    if (minimized) {
      _savedPanelHeight = panel.style.height || '';
      body.style.display  = 'none';
      panel.style.height  = 'auto';
      panel.style.overflow = 'visible';
      btn.innerHTML = '&#9633;';
    } else {
      body.style.display  = '';
      panel.style.height  = _savedPanelHeight || '100vh';
      panel.style.overflow = 'hidden';
      btn.innerHTML = '&#8212;';
    }
  };

  // ── BOTÃO FLUTUANTE ───────────────────────────────────────────────────────────
  function injectButton() {
    if (document.getElementById('btn-mon')) return;
    if (window.self !== window.top) return;
    const btn = document.createElement('button');
    btn.id = 'btn-mon';
    btn.innerHTML = `<span class="mon-fab-dot"></span> Monitor`;
    btn.onclick = toggleMonitor;
    document.body.appendChild(btn);
  }

  // ── ESTILOS ───────────────────────────────────────────────────────────────────
  function injectStyles() {
    if (document.getElementById('mon-style')) return;
    const s = document.createElement('style');
    s.id = 'mon-style';
    s.textContent = `
      @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@400;600&display=swap');

      /* ── VARIÁVEIS TEMA EQUILIBRADO ── */
      :root {
        --mon-bg:          #e8ecf2;
        --mon-surface:     #f0f3f8;
        --mon-surface2:    #dde2eb;
        --mon-surface3:    #d2d8e4;
        --mon-border:      #c6cedd;
        --mon-border2:     #b3bdce;
        --mon-text:        #1b2333;
        --mon-text-dim:    #4e5d72;
        --mon-text-faint:  #8898ad;
        --mon-green:       #0a7c57;
        --mon-green-light: #0e9e6e;
        --mon-green-bg:    rgba(10,124,87,0.11);
        --mon-green-border:rgba(10,124,87,0.28);
        --mon-amber:       #a85800;
        --mon-amber-bg:    rgba(168,88,0,0.1);
        --mon-amber-border:rgba(168,88,0,0.28);
        --mon-red:         #b91c1c;
        --mon-red-bg:      rgba(185,28,28,0.09);
        --mon-red-border:  rgba(185,28,28,0.28);
        --mon-blue:        #1d4ed8;
        --mon-blue-bg:     rgba(29,78,216,0.09);
        --mon-blue-border: rgba(29,78,216,0.28);
        --mon-accent:      #4338ca;
        --mon-accent-bg:   rgba(67,56,202,0.1);
        --mon-accent-border:rgba(67,56,202,0.28);
        --mon-shadow-sm:   0 1px 3px rgba(0,0,0,0.1), 0 1px 2px rgba(0,0,0,0.07);
        --mon-shadow:      0 4px 16px rgba(0,0,0,0.13), 0 1px 4px rgba(0,0,0,0.07);
        --mon-shadow-lg:   0 20px 60px rgba(0,0,0,0.18), 0 4px 16px rgba(0,0,0,0.1);
        --mon-radius:      10px;
        --mon-radius-sm:   6px;
        --mon-radius-xs:   4px;
        --mon-font: 'Inter', system-ui, sans-serif;
        --mon-mono: 'JetBrains Mono', 'Consolas', monospace;
      }

      /* ── KEYFRAMES ── */
      @keyframes mon-shake {
        0%   { transform: rotate(-6deg) scale(1.03); }
        100% { transform: rotate(6deg)  scale(1.03); }
      }
      @keyframes mon-bar-fill {
        from { width: 0; }
        to   { width: var(--bar-w); }
      }
      @keyframes mon-fadein {
        from { opacity: 0; transform: translateY(4px); }
        to   { opacity: 1; transform: translateY(0); }
      }
      @keyframes mon-pulse-dot {
        0%, 100% { opacity: 1; transform: scale(1); }
        50%       { opacity: 0.6; transform: scale(0.85); }
      }
      @keyframes mon-hg-thread-blink {
        0%, 100% { opacity: 0.7; }
        50%       { opacity: 0.2; }
      }
      @keyframes mon-row-in {
        from { opacity: 0; transform: translateX(4px); }
        to   { opacity: 1; transform: translateX(0); }
      }
      @keyframes mon-slide-in {
        from { opacity: 0; transform: translateX(20px); }
        to   { opacity: 1; transform: translateX(0); }
      }

      /* ── BOTÃO FLUTUANTE ── */
      #btn-mon {
        position: fixed; bottom: 24px; right: 24px; z-index: 99999;
        display: flex; align-items: center; gap: 8px;
        background: var(--mon-accent);
        color: #fff;
        border: none;
        padding: 10px 20px; border-radius: 99px;
        font-size: 13px; font-family: var(--mon-font);
        font-weight: 600; cursor: pointer; letter-spacing: 0.2px;
        box-shadow: 0 4px 20px rgba(79,70,229,0.35), 0 1px 4px rgba(79,70,229,0.2);
        transition: all 0.2s cubic-bezier(0.4,0,0.2,1);
      }
      #btn-mon:hover {
        background: #4338ca;
        box-shadow: 0 6px 28px rgba(79,70,229,0.45);
        transform: translateY(-1px);
      }
      #btn-mon:active { transform: translateY(0); }
      .mon-fab-dot {
        width: 7px; height: 7px; border-radius: 50%;
        background: #a5f3b8;
        display: inline-block;
        animation: mon-pulse-dot 2s ease-in-out infinite;
      }

      /* ── PAINEL PRINCIPAL ── */
      #mon-panel {
        font-family: var(--mon-font);
        font-size: 13px;
        background: var(--mon-bg);
        color: var(--mon-text);
        border-left: 1px solid var(--mon-border);
        box-shadow: var(--mon-shadow-lg);
        animation: mon-slide-in 0.25s cubic-bezier(0.4,0,0.2,1);
      }

      /* ── RESIZE HANDLE ── */
      .mon-resize-handle {
        position: absolute; left: 0; top: 0; width: 4px; height: 100%;
        cursor: ew-resize; z-index: 10; background: transparent;
        transition: background 0.15s;
      }
      .mon-resize-handle:hover { background: var(--mon-accent); opacity: 0.3; }

      /* ── SCROLLBAR — sempre visível, larga ── */
      #mon-panel ::-webkit-scrollbar { width: 8px; }
      #mon-panel ::-webkit-scrollbar-track { background: var(--mon-surface3); border-radius: 4px; }
      #mon-panel ::-webkit-scrollbar-thumb {
        background: var(--mon-border2);
        border-radius: 4px;
        border: 2px solid var(--mon-surface3);
        min-height: 40px;
      }
      #mon-panel ::-webkit-scrollbar-thumb:hover { background: var(--mon-text-faint); }
      #mon-table-wrap { scrollbar-gutter: stable; }

      /* ── HEADER ── */
      #mon-header {
        background: var(--mon-surface);
        border-bottom: 1px solid var(--mon-border);
        padding: 0 16px;
        height: 56px;
        display: flex; align-items: center; justify-content: space-between;
        flex-shrink: 0; user-select: none; cursor: grab;
        gap: 12px;
      }
      #mon-header:active { cursor: grabbing; }
      .mon-logo {
        display: flex; align-items: center; gap: 10px; flex-shrink: 0;
      }
      .mon-logo-icon {
        width: 32px; height: 32px; border-radius: 50%;
        display: flex; align-items: center; justify-content: center;
        flex-shrink: 0; overflow: hidden;
        box-shadow: 0 2px 8px rgba(0,100,0,0.35);
      }
      .mon-logo-icon img {
        width: 32px; height: 32px; object-fit: cover; display: block;
      }
      .mon-logo-text {
        display: flex; flex-direction: column; gap: 0px;
      }
      .mon-logo-title {
        font-size: 13px; font-weight: 700; color: var(--mon-text);
        letter-spacing: -0.2px;
      }
      .mon-logo-sub {
        font-size: 11px; color: var(--mon-text-faint); font-weight: 400;
      }

      /* ── STATUS PILL ── */
      .mon-status-pill {
        display: inline-flex; align-items: center; gap: 5px;
        padding: 4px 10px; border-radius: 99px;
        border: 1px solid;
        font-size: 11px; font-weight: 600; letter-spacing: 0.2px;
        white-space: nowrap;
      }
      .mon-status-pill[data-state="live"] {
        color: var(--mon-green); border-color: var(--mon-green-border);
        background: var(--mon-green-bg);
      }
      .mon-status-pill[data-state="sync"] {
        color: var(--mon-amber); border-color: var(--mon-amber-border);
        background: var(--mon-amber-bg);
      }
      .mon-status-pill[data-state="offline"] {
        color: var(--mon-text-faint); border-color: var(--mon-border);
        background: transparent;
      }
      .mon-status-dot {
        width: 6px; height: 6px; border-radius: 50%; background: currentColor;
        flex-shrink: 0;
      }
      .mon-status-pill[data-state="live"] .mon-status-dot {
        animation: mon-pulse-dot 1.8s ease-in-out infinite;
      }
      .mon-status-pill[data-state="sync"] .mon-status-dot {
        animation: mon-spin 1s linear infinite;
        border-radius: 0;
        clip-path: polygon(50% 0%, 100% 100%, 0% 100%);
      }

      /* ── BOTÕES HEADER ── */
      .mon-hdr-btn {
        height: 32px; padding: 0 12px; border-radius: var(--mon-radius-sm);
        border: 1px solid var(--mon-border2); background: var(--mon-surface);
        color: var(--mon-text-dim); font-size: 12px; font-family: var(--mon-font);
        font-weight: 500; cursor: pointer; transition: all 0.15s;
        display: inline-flex; align-items: center; gap: 5px;
        white-space: nowrap;
      }
      .mon-hdr-btn:hover {
        background: var(--mon-surface2); color: var(--mon-text);
        border-color: var(--mon-border2);
      }
      .mon-hdr-btn--green {
        color: var(--mon-green); border-color: var(--mon-green-border);
        background: var(--mon-green-bg);
      }
      .mon-hdr-btn--green:hover { background: rgba(5,150,105,0.14); }
      .mon-hdr-btn[data-state="on"] { color: var(--mon-green); }
      .mon-hdr-btn[data-state="off"] { color: var(--mon-red); }

      .mon-icon-btn {
        width: 32px; height: 32px; border-radius: var(--mon-radius-sm);
        border: 1px solid var(--mon-border); background: transparent;
        color: var(--mon-text-faint); font-size: 14px; cursor: pointer;
        display: inline-flex; align-items: center; justify-content: center;
        transition: all 0.15s; flex-shrink: 0;
      }
      .mon-icon-btn:hover { background: var(--mon-surface2); color: var(--mon-text); border-color: var(--mon-border2); }

      /* ── MÉTRICAS ── */
      #mon-metrics {
        display: grid; grid-template-columns: repeat(4, 1fr);
        border-bottom: 1px solid var(--mon-border);
        flex-shrink: 0;
        background: var(--mon-surface);
      }
      .mon-metric {
        padding: 14px 16px; position: relative;
        border-right: 1px solid var(--mon-border);
      }
      .mon-metric:last-child { border-right: none; }
      .mon-metric-label {
        font-size: 10px; color: var(--mon-text-faint); letter-spacing: 0.5px;
        text-transform: uppercase; margin-bottom: 5px; font-weight: 600;
      }
      .mon-metric-val {
        font-size: 28px; font-weight: 700; line-height: 1;
        color: var(--mon-text); letter-spacing: -1px;
      }
      .mon-metric-val.green  { color: var(--mon-green); }
      .mon-metric-val.amber  { color: var(--mon-amber); }
      .mon-metric-val.red    { color: var(--mon-red); }
      .mon-metric-accent {
        position: absolute; bottom: 0; left: 16px; right: 16px;
        height: 2px; border-radius: 1px; opacity: 0;
        transition: opacity 0.3s;
      }
      .mon-metric:hover .mon-metric-accent { opacity: 1; }
      .mon-metric:nth-child(1) .mon-metric-accent { background: var(--mon-accent); }
      .mon-metric:nth-child(2) .mon-metric-accent { background: var(--mon-green); }
      .mon-metric:nth-child(3) .mon-metric-accent { background: var(--mon-amber); }
      .mon-metric:nth-child(4) .mon-metric-accent { background: var(--mon-red); }

      /* ── FILTER BAR ── */
      #mon-filter-bar {
        display: flex; flex-wrap: wrap; align-items: center; gap: 8px;
        padding: 10px 14px; border-bottom: 1px solid var(--mon-border);
        background: var(--mon-surface); flex-shrink: 0;
      }
      #mon-filter-input-wrap {
        display: flex; align-items: center; gap: 7px;
        background: var(--mon-bg); border: 1.5px solid var(--mon-border);
        border-radius: var(--mon-radius-sm); padding: 6px 11px;
        flex: 1; min-width: 180px; transition: border-color 0.15s;
      }
      #mon-filter-input-wrap:focus-within {
        border-color: var(--mon-accent);
        box-shadow: 0 0 0 3px rgba(79,70,229,0.1);
      }
      .mon-filter-icon { color: var(--mon-text-faint); font-size: 14px; flex-shrink: 0; }
      #mon-filter-input {
        flex: 1; background: transparent; border: none; outline: none;
        color: var(--mon-text); font-size: 12.5px; font-family: var(--mon-font);
      }
      #mon-filter-input::placeholder { color: var(--mon-text-faint); }
      #mon-filter-clear {
        background: none; border: none; color: var(--mon-text-faint);
        cursor: pointer; font-size: 12px; padding: 0; line-height: 1;
        display: none; transition: color 0.15s;
      }
      #mon-filter-clear:hover { color: var(--mon-red); }

      /* ── STATUS CHIPS ── */
      #mon-status-chips { display: flex; flex-wrap: wrap; gap: 5px; }
      .mon-chip {
        padding: 4px 11px; border-radius: 99px; font-size: 11px; font-weight: 500;
        font-family: var(--mon-font); cursor: pointer; border: 1.5px solid var(--mon-border);
        background: transparent; color: var(--mon-text-dim);
        transition: all 0.15s; letter-spacing: 0.1px; white-space: nowrap;
      }
      .mon-chip:hover { background: var(--mon-surface2); color: var(--mon-text); border-color: var(--mon-border2); }
      .mon-chip.active { background: var(--mon-accent-bg); border-color: var(--mon-accent-border); color: var(--mon-accent); font-weight: 600; }
      .mon-chip--completo.active { background: var(--mon-green-bg); border-color: var(--mon-green-border); color: var(--mon-green); }
      .mon-chip--parcial.active  { background: var(--mon-amber-bg); border-color: var(--mon-amber-border); color: var(--mon-amber); }
      .mon-chip--esc.active      { background: var(--mon-blue-bg); border-color: var(--mon-blue-border); color: var(--mon-blue); }
      .mon-chip--esc-inc.active  { background: rgba(234,88,12,0.07); border-color: rgba(234,88,12,0.25); color: #ea580c; }
      .mon-chip--nenhum.active   { background: var(--mon-red-bg); border-color: var(--mon-red-border); color: var(--mon-red); }

      /* ── TABELA ── */
      #mon-table-wrap { flex: 1; overflow-y: auto; background: var(--mon-bg); }
      #mon-table {
        width: 100%; border-collapse: collapse; font-size: 12.5px;
        table-layout: fixed;
      }
      #mon-table thead th {
        padding: 9px 13px;
        text-align: left; font-size: 10.5px; font-weight: 600;
        color: var(--mon-text-faint); letter-spacing: 0.4px;
        text-transform: uppercase;
        background: var(--mon-surface);
        border-bottom: 1.5px solid var(--mon-border);
        position: sticky; top: 0; z-index: 2;
        white-space: nowrap;
      }
      #mon-table thead th.center { text-align: center; }

      /* ── ROWS ── */
      tr.op-row {
        border-bottom: 1px solid var(--mon-border);
        cursor: pointer;
        transition: background 0.1s;
        animation: mon-row-in 0.2s ease both;
      }
      tr.op-row:hover td { background: rgba(79,70,229,0.03); }
      tr.op-row.is-expanded td { background: rgba(79,70,229,0.04); }
      tr.op-row td {
        padding: 11px 13px; vertical-align: middle;
        background: transparent; transition: background 0.1s;
      }

      .mon-chave {
        font-family: var(--mon-mono);
        font-size: 11px; color: var(--mon-accent);
        white-space: nowrap; overflow: hidden; text-overflow: ellipsis;
        letter-spacing: -0.2px; display: block;
        font-weight: 600;
      }
      .mon-sigla {
        font-weight: 700; font-size: 13px; color: var(--mon-text);
        letter-spacing: -0.3px;
      }
      .mon-hora {
        font-family: var(--mon-mono);
        font-size: 12px; color: var(--mon-text);
        letter-spacing: -0.2px; font-weight: 600;
      }
      .mon-lider {
        font-size: 12px; color: var(--mon-text-dim);
        white-space: nowrap; overflow: hidden; text-overflow: ellipsis;
        display: block;
      }
      .mon-chevron { color: var(--mon-text-faint); font-size: 10px; transition: transform 0.2s; display: inline-block; }
      tr.op-row.is-expanded .mon-chevron { transform: rotate(180deg); }

      /* ── PROGRESS BADGE ── */
      .mon-prog-cell { text-align: center; }
      .mon-prog-num {
        font-size: 13px; font-weight: 700; letter-spacing: -0.3px;
      }
      .mon-prog-bar {
        height: 3px; background: var(--mon-surface3); border-radius: 2px;
        overflow: hidden; margin-top: 4px;
      }
      .mon-prog-fill {
        height: 100%; border-radius: 2px;
        animation: mon-bar-fill 0.6s cubic-bezier(0.4,0,0.2,1) both;
        width: var(--bar-w);
      }
      .mon-prog-pending { font-size: 12px; color: var(--mon-text-faint); }
      .mon-prog-na { font-size: 11px; color: var(--mon-text-faint); }

      /* ── STATUS BADGES ── */
      .mon-status-badge {
        display: inline-flex; align-items: center; gap: 4px;
        padding: 3px 9px; border-radius: 99px;
        font-size: 11px; font-weight: 600; letter-spacing: 0.1px;
        white-space: nowrap; border: 1px solid transparent;
      }
      .mon-status-badge.completo  { color: var(--mon-green); background: var(--mon-green-bg); border-color: var(--mon-green-border); }
      .mon-status-badge.parcial   { color: var(--mon-amber); background: var(--mon-amber-bg); border-color: var(--mon-amber-border); }
      .mon-status-badge.esc-ok    { color: var(--mon-blue); background: var(--mon-blue-bg); border-color: var(--mon-blue-border); }
      .mon-status-badge.nenhum    { color: var(--mon-red); background: var(--mon-red-bg); border-color: var(--mon-red-border); }
      .mon-status-badge.neutro    { color: var(--mon-text-faint); background: transparent; }
      .mon-envelope { font-size: 12px; margin-left: 4px; opacity: 0.7; }

      /* ── SORT HEADERS ── */
      .mon-th-sort { cursor: pointer; user-select: none; }
      .mon-th-sort:hover { color: var(--mon-text-dim); }
      .mon-sort-arrow { font-size: 9px; margin-left: 3px; opacity: 0.35; }
      .mon-th-sort.sort-asc  .mon-sort-arrow::after { content: '▲'; opacity: 1; color: var(--mon-accent); }
      .mon-th-sort.sort-desc .mon-sort-arrow::after { content: '▼'; opacity: 1; color: var(--mon-accent); }
      .mon-th-sort:not(.sort-asc):not(.sort-desc) .mon-sort-arrow::after { content: '⇅'; }

      /* ── DETALHE EXPANDIDO ── */
      tr.op-detail td { padding: 0 !important; border-bottom: 1px solid var(--mon-border); }
      .mon-detail-inner {
        background: var(--mon-bg);
        border-top: 1px solid var(--mon-border);
        padding: 16px;
        animation: mon-fadein 0.2s ease;
      }
      .mon-detail-header {
        display: flex; align-items: center; gap: 8px; margin-bottom: 14px;
        flex-wrap: wrap;
      }
      .mon-key-chip {
        font-family: var(--mon-mono); font-size: 11.5px;
        color: var(--mon-accent); background: var(--mon-accent-bg);
        border: 1px solid var(--mon-accent-border); border-radius: var(--mon-radius-xs);
        padding: 3px 9px; letter-spacing: -0.2px; font-weight: 600;
      }
      .mon-copy-btn {
        background: transparent; border: 1px solid var(--mon-border);
        color: var(--mon-text-dim); padding: 4px 10px;
        border-radius: var(--mon-radius-xs); font-size: 11px; font-family: var(--mon-font);
        cursor: pointer; transition: all 0.15s;
      }
      .mon-copy-btn:hover { background: var(--mon-surface2); color: var(--mon-text); border-color: var(--mon-border2); }
      .mon-open-btn {
        background: var(--mon-accent-bg); border: 1px solid var(--mon-accent-border);
        color: var(--mon-accent) !important; text-decoration: none;
        padding: 4px 10px; border-radius: var(--mon-radius-xs);
        font-size: 11px; font-family: var(--mon-font); font-weight: 600;
        cursor: pointer; display: inline-flex; align-items: center; gap: 4px;
        transition: all 0.15s;
      }
      .mon-open-btn:hover { background: rgba(79,70,229,0.14); }

      .mon-stat-grid {
        display: grid; grid-template-columns: 1fr 1fr; gap: 10px;
        margin-bottom: 14px;
      }
      .mon-stat-card {
        background: var(--mon-surface);
        border: 1px solid var(--mon-border);
        border-radius: var(--mon-radius-sm);
        padding: 13px 15px;
        box-shadow: var(--mon-shadow-sm);
      }
      .mon-stat-card-header {
        display: flex; justify-content: space-between; align-items: baseline;
        margin-bottom: 9px;
      }
      .mon-stat-card-label {
        font-size: 10px; color: var(--mon-text-faint); letter-spacing: 0.5px;
        text-transform: uppercase; font-weight: 600;
      }
      .mon-stat-card-num {
        font-size: 20px; font-weight: 700; letter-spacing: -0.5px;
      }
      .mon-stat-card-sub {
        font-size: 12px; color: var(--mon-text-faint);
        font-weight: 400; margin-left: 1px;
      }
      .mon-stat-card-pct { font-size: 10px; color: var(--mon-text-faint); margin-top: 6px; }

      /* ── ACTIONS ── */
      .mon-actions { display: flex; flex-wrap: wrap; gap: 7px; margin-bottom: 14px; align-items: center; }

      .mon-send-btn {
        display: inline-flex; align-items: center; gap: 5px;
        background: var(--mon-surface);
        border: 1px solid var(--mon-border2);
        color: var(--mon-text-dim);
        padding: 6px 13px; border-radius: var(--mon-radius-sm);
        font-size: 11.5px; font-family: var(--mon-font); font-weight: 500;
        cursor: pointer; letter-spacing: 0.1px;
        transition: all 0.15s; box-shadow: var(--mon-shadow-sm);
      }
      .mon-send-btn:hover { background: var(--mon-surface2); color: var(--mon-text); border-color: var(--mon-border2); }
      .mon-send-btn:disabled { cursor: not-allowed; opacity: 0.45; }
      .mon-send-btn--err { color: var(--mon-red) !important; border-color: var(--mon-red-border) !important; background: var(--mon-red-bg) !important; }
      .mon-send-btn--report { background: var(--mon-amber-bg) !important; border-color: var(--mon-amber-border) !important; color: var(--mon-amber) !important; }
      .mon-send-btn--report:hover { background: rgba(217,119,6,0.14) !important; }
      .mon-send-btn--print { background: var(--mon-green-bg) !important; border-color: var(--mon-green-border) !important; color: var(--mon-green) !important; }
      .mon-send-btn--print:hover { background: rgba(5,150,105,0.14) !important; }
      .mon-send-btn--relatorio { background: rgba(22,163,74,0.07) !important; border-color: rgba(22,163,74,0.25) !important; color: #16a34a !important; }
      .mon-send-btn--relatorio:hover { background: rgba(22,163,74,0.14) !important; }
      .mon-send-btn--open { background: var(--mon-accent-bg) !important; border-color: var(--mon-accent-border) !important; color: var(--mon-accent) !important; text-decoration: none; }
      .mon-send-btn--open:hover { background: rgba(79,70,229,0.14) !important; }

      .mon-dl-btn {
        display: inline-flex; align-items: center; gap: 5px;
        background: var(--mon-surface);
        border: 1px solid var(--mon-border2);
        color: var(--mon-text-dim);
        padding: 6px 12px; border-radius: var(--mon-radius-sm);
        font-size: 11.5px; font-family: var(--mon-font);
        cursor: pointer; transition: all 0.15s; text-decoration: none;
        box-shadow: var(--mon-shadow-sm);
      }
      .mon-dl-btn:hover { background: var(--mon-surface2); color: var(--mon-text); }

      .mon-xls-menu { position: relative; display: inline-block; }
      .mon-xls-dropdown {
        display: none; position: absolute; bottom: calc(100% + 6px); left: 0;
        background: var(--mon-surface); border: 1px solid var(--mon-border2);
        border-radius: var(--mon-radius-sm); padding: 4px;
        z-index: 99999; min-width: 210px;
        box-shadow: var(--mon-shadow);
        animation: mon-fadein 0.15s ease;
      }
      .mon-xls-menu.open .mon-xls-dropdown { display: block; }
      .mon-xls-dropdown a {
        display: block; padding: 9px 12px; color: var(--mon-text-dim);
        font-size: 12px; font-family: var(--mon-font); text-decoration: none;
        transition: all 0.1s; border-radius: var(--mon-radius-xs);
      }
      .mon-xls-dropdown a:hover { background: var(--mon-surface2); color: var(--mon-text); }

      /* ── LISTAS LADO A LADO ── */
      #mon-panel .mon-lists-grid {
        display: grid !important; grid-template-columns: 1fr 1fr !important; gap: 10px !important;
        margin: 0 !important; padding: 0 !important;
      }
      #mon-panel .mon-list-panel {
        border-radius: var(--mon-radius-sm) !important;
        border: 1px solid var(--mon-border) !important;
        overflow: hidden !important;
        background: var(--mon-surface) !important;
        box-shadow: var(--mon-shadow-sm) !important;
      }
      #mon-panel .mon-list-panel-header {
        display: flex !important; align-items: center !important; gap: 8px !important;
        padding: 9px 13px !important;
        background: var(--mon-bg) !important;
        border-bottom: 1px solid var(--mon-border) !important;
        margin: 0 !important;
      }
      #mon-panel .mon-list-panel-dot {
        width: 7px !important; height: 7px !important; border-radius: 50% !important;
        flex-shrink: 0 !important; display: inline-block !important;
      }
      #mon-panel .mon-list-panel-title {
        font-size: 10.5px !important; font-weight: 700 !important; letter-spacing: 0.4px !important;
        text-transform: uppercase !important; color: var(--mon-text-dim) !important; flex: 1 !important;
        margin: 0 !important;
      }
      #mon-panel .mon-list-panel-count {
        font-size: 10.5px !important; font-weight: 700 !important;
        padding: 2px 8px !important; border-radius: 99px !important;
        line-height: 1.4 !important; border: 1px solid !important;
      }
      #mon-panel .mon-list-panel-body {
        max-height: 230px !important; overflow-y: auto !important;
      }
      #mon-panel .mon-list-row {
        padding: 9px 13px !important;
        border-bottom: 1px solid var(--mon-border) !important;
        transition: background 0.1s !important;
        display: block !important;
      }
      #mon-panel .mon-list-row:last-child { border-bottom: none !important; }
      #mon-panel .mon-list-row:hover { background: var(--mon-surface2) !important; }
      #mon-panel .mon-list-row-name {
        font-size: 12px !important; font-weight: 600 !important; color: var(--mon-text) !important;
        margin-bottom: 3px !important; display: block !important;
        white-space: normal !important;
      }
      #mon-panel .mon-list-row-meta {
        display: flex !important; align-items: center !important; gap: 8px !important;
      }
      #mon-panel .mon-list-row-tipo {
        font-size: 11px !important; color: var(--mon-text-dim) !important;
      }
      #mon-panel .mon-list-row-time {
        font-size: 11px !important; color: var(--mon-amber) !important;
        font-family: var(--mon-mono) !important; margin-left: auto !important;
        background: var(--mon-amber-bg) !important;
        border: 1px solid var(--mon-amber-border) !important;
        border-radius: var(--mon-radius-xs) !important; padding: 2px 8px !important;
        letter-spacing: -0.2px !important; font-weight: 600 !important;
      }
      #mon-panel .mon-list-empty {
        padding: 20px 13px !important; font-size: 12px !important;
        color: var(--mon-text-faint) !important; text-align: center !important;
        display: block !important;
      }

      .mon-loading-detail {
        display: flex; align-items: center; gap: 9px;
        color: var(--mon-text-dim); font-size: 12px; padding: 4px 0;
      }
      .mon-loading-spinner {
        width: 15px; height: 15px; border-radius: 50%;
        border: 2px solid var(--mon-border);
        border-top-color: var(--mon-accent);
        animation: mon-spin 0.7s linear infinite; flex-shrink: 0;
      }

      .mon-empty-detail { color: var(--mon-text-faint); font-size: 12px; padding: 8px 0; }

      /* ── BOTÕES ESCALA WA / GMAIL ── */
      #mon-panel .mon-send-btn--wa-escala {
        background: var(--mon-green-bg) !important;
        color: var(--mon-green) !important;
        border-color: var(--mon-green-border) !important;
      }
      #mon-panel .mon-send-btn--wa-escala:hover {
        background: var(--mon-green) !important;
        color: #fff !important;
      }
      #mon-panel .mon-send-btn--gmail {
        background: var(--mon-blue-bg) !important;
        color: var(--mon-blue) !important;
        border-color: rgba(29,78,216,0.25) !important;
      }
      #mon-panel .mon-send-btn--gmail:hover {
        background: var(--mon-blue) !important;
        color: #fff !important;
      }

      /* ── PROGRESS AVATAR ── */
      .mon-avatar-wrap {
        position: relative; width: 44px; height: 44px; flex-shrink: 0;
      }
      #mon-avatar-img {
        width: 40px; height: 40px; border-radius: 50%; object-fit: cover;
        position: absolute; top: 2px; left: 2px; z-index: 1;
        transform-origin: center;
      }
      #mon-progress-overlay {
        position: absolute; top: 2px; left: 2px; width: 40px; height: 40px;
        border-radius: 50%; background: rgba(255,255,255,0.65); z-index: 2;
        transition: opacity 0.3s;
      }
      #mon-progress-text {
        position: absolute; top: 2px; left: 2px; width: 40px; height: 40px;
        border-radius: 50%; z-index: 3;
        display: flex; align-items: center; justify-content: center;
        font-size: 9px; font-weight: 700; color: var(--mon-text); font-family: var(--mon-mono);
      }

      /* ── HIGHLIGHT DE BUSCA ── */
      mark {
        background: rgba(79,70,229,0.12) !important;
        color: var(--mon-accent) !important;
        border-radius: 2px !important;
        padding: 0 1px !important;
      }
    `;
    document.head.appendChild(s);
  }

  // ── PAINEL ────────────────────────────────────────────────────────────────────
  function createPanel() {
    if (document.getElementById('mon-panel')) return;
    injectStyles();

    const p = document.createElement('div');
    p.id = 'mon-panel';
    p.style.cssText = `
      position:fixed;top:0;right:0;width:1000px;height:100vh;
      z-index:99998;display:none;flex-direction:column;overflow:hidden;
    `;

    const r = 20, circ = 2 * Math.PI * r;
    const notifState = !('Notification' in window) ? 'off' :
      Notification.permission === 'granted' ? 'on' :
      Notification.permission === 'denied'  ? 'off' : 'default';
    const notifLabel = notifState === 'on' ? '🔔 Notif. ativa' : notifState === 'off' ? '🔕 Bloqueado' : '🔔 Ativar notif.';

    p.innerHTML = `
      <!-- HEADER -->
      <div id="mon-header">
        <div class="mon-logo">
          <div class="mon-logo-icon"><img src="https://upload.wikimedia.org/wikipedia/commons/1/10/Palmeiras_logo.svg" alt="TSI" onerror="this.style.display='none';this.parentElement.textContent='M'" /></div>
          <div class="mon-logo-text">
            <div class="mon-logo-title">Monitor TSI</div>
            <div style="display:flex;align-items:center;gap:4px;margin-top:1px;">
              <div class="mon-logo-sub" id="mon-sub">Inicializando…</div>
              <div id="mon-hourglass-wrap" title="Próxima atualização automática" style="display:flex;align-items:center;gap:2px;line-height:1;">
                <span id="mon-hg-icon" style="font-size:11px;line-height:1;display:inline-block;transition:transform 0.3s;">⏳</span>
                <span id="mon-hg-count" style="font-size:10px;font-family:var(--mon-mono);color:var(--mon-text-faint);font-weight:600;">—</span>
              </div>
            </div>
          </div>
        </div>
        <div style="display:flex;align-items:center;gap:7px;flex-wrap:nowrap">
          <span id="mon-live" class="mon-status-pill" data-state="offline">
            <span class="mon-status-dot"></span>
            <span>Offline</span>
          </span>
          <button class="mon-hdr-btn mon-hdr-btn--green" onclick="window._monRefresh()" title="Atualizar tudo agora">
            ↻ Atualizar
          </button>
          <button id="mon-notif-btn" class="mon-hdr-btn" data-state="${notifState}"
            onclick="window._monPedirNotif()">${notifLabel}</button>
          <!-- Avatar progress -->
          <div class="mon-avatar-wrap" title="Progresso de carregamento">
            <img id="mon-avatar-img" src="${AVATAR_URL}" />
            <div id="mon-progress-overlay"></div>
            <div id="mon-progress-text" style="display:none">0%</div>
            <svg width="44" height="44" style="position:absolute;top:0;left:0;transform:rotate(-90deg);z-index:4">
              <circle cx="22" cy="22" r="${r}" fill="none" stroke="var(--mon-border)" stroke-width="2.5"/>
              <circle id="mon-progress-circle" cx="22" cy="22" r="${r}" fill="none" stroke="var(--mon-amber)" stroke-width="2.5" stroke-dasharray="${circ}" stroke-dashoffset="${circ}" style="transition:stroke-dashoffset 0.4s ease,stroke 0.4s ease"/>
            </svg>
          </div>
          <button class="mon-icon-btn" id="mon-min-btn" onclick="window._monMinimize()" title="Minimizar">&#8212;</button>
          <button class="mon-icon-btn" onclick="document.getElementById('mon-panel').style.display='none';document.getElementById('btn-mon').innerHTML='<span class=mon-fab-dot></span> Monitor'" title="Fechar">&#10005;</button>
        </div>
      </div>

      <!-- BODY -->
      <div id="mon-body" style="display:flex;flex-direction:column;flex:1;overflow:hidden">
        <!-- MÉTRICAS -->
        <div id="mon-metrics">
          <div class="mon-metric">
            <div class="mon-metric-label">Operações</div>
            <div class="mon-metric-val" id="m-total">—</div>
            <div class="mon-metric-accent"></div>
          </div>
          <div class="mon-metric">
            <div class="mon-metric-label">Completas</div>
            <div class="mon-metric-val green" id="m-ok">—</div>
            <div class="mon-metric-accent"></div>
          </div>
          <div class="mon-metric">
            <div class="mon-metric-label">Parciais</div>
            <div class="mon-metric-val amber" id="m-inc">—</div>
            <div class="mon-metric-accent"></div>
          </div>
          <div class="mon-metric">
            <div class="mon-metric-label">Sem apontamento</div>
            <div class="mon-metric-val red" id="m-zero">—</div>
            <div class="mon-metric-accent"></div>
          </div>
        </div>

        <!-- FILTRO + STATUS CHIPS -->
        <div id="mon-filter-bar">
          <div id="mon-filter-input-wrap">
            <span class="mon-filter-icon">⌕</span>
            <input id="mon-filter-input" type="text" placeholder="Filtrar por chave, sigla ou líder…"
              oninput="window._monSetFilter(this.value)" autocomplete="off" spellcheck="false" />
            <button id="mon-filter-clear" onclick="window._monClearFilter()" title="Limpar">✕</button>
          </div>
          <div id="mon-status-chips">
            <button class="mon-chip mon-chip--all active" onclick="window._monSetStatusFilter('all',this)">Todos</button>
            <button class="mon-chip mon-chip--completo" onclick="window._monSetStatusFilter('completo',this)">✓ Completo</button>
            <button class="mon-chip mon-chip--parcial"  onclick="window._monSetStatusFilter('parcial',this)">△ Parcial</button>
            <button class="mon-chip mon-chip--esc"      onclick="window._monSetStatusFilter('esc',this)">Esc. ok</button>
            <button class="mon-chip mon-chip--esc-inc"  onclick="window._monSetStatusFilter('esc_inc',this)">⚠ Esc. inc.</button>
            <button class="mon-chip mon-chip--nenhum"   onclick="window._monSetStatusFilter('nenhum',this)">✗ Nenhum</button>
          </div>
        </div>

        <!-- TABELA -->
        <div id="mon-table-wrap">
          <table id="mon-table">
            <thead>
              <tr>
                <th style="width:15%">Chave</th>
                <th style="width:9%">Sigla</th>
                <th class="center mon-th-sort" data-col="esc" style="width:13%" onclick="window._monToggleSort('esc',this)">Esc / Sol <span class="mon-sort-arrow"></span></th>
                <th class="center mon-th-sort" data-col="apt" style="width:13%" onclick="window._monToggleSort('apt',this)">Apt / Sol <span class="mon-sort-arrow"></span></th>
                <th class="mon-th-sort" data-col="hora" style="width:8%" onclick="window._monToggleSort('hora',this)">Hora <span class="mon-sort-arrow"></span></th>
                <th style="width:20%">Líder</th>
                <th class="center mon-th-sort" data-col="status" style="width:18%" onclick="window._monToggleSort('status',this)">Status <span class="mon-sort-arrow"></span></th>
                <th style="width:4%"></th>
              </tr>
            </thead>
            <tbody id="mon-tbody"></tbody>
          </table>
        </div>
      </div>
    `;
    document.body.appendChild(p);
    initControls(p);
  }


  // ── BADGES ────────────────────────────────────────────────────────────────────
  function colorForPct(pct) {
    if (pct >= 100) return 'var(--mon-green)';
    if (pct > 0)    return 'var(--mon-indigo)';
    return 'var(--mon-text-faint)';
  }
  function colorForAptPct(pct) {
    if (pct >= 100) return 'var(--mon-green)';
    if (pct > 0)    return 'var(--mon-amber)';
    return 'var(--mon-text-faint)';
  }

  function escaladoBadge(d, qtd) {
    if (!d || d === 'loading') return `<span class="mon-prog-pending">…/${qtd}</span>`;
    const pct = qtd > 0 ? Math.min(100, Math.round((d.escalado / qtd) * 100)) : 0;
    const cor = colorForPct(pct);
    return `
      <div class="mon-prog-cell">
        <div class="mon-prog-num" style="color:${cor}">${d.escalado}<span style="color:var(--mon-text-faint);font-size:10px;font-weight:400">/${qtd}</span></div>
        <div class="mon-prog-bar"><div class="mon-prog-fill" style="--bar-w:${pct}%;background:${cor}"></div></div>
      </div>`;
  }

  function apontBadge(d, qtd) {
    if (!d || d === 'loading') return `<span class="mon-prog-pending">…/${qtd}</span>`;
    const pct = qtd > 0 ? Math.min(100, Math.round((d.apontado / qtd) * 100)) : 0;
    const cor = colorForAptPct(pct);
    return `
      <div class="mon-prog-cell">
        <div class="mon-prog-num" style="color:${cor}">${d.apontado}<span style="color:var(--mon-text-faint);font-size:10px;font-weight:400">/${qtd}</span></div>
        <div class="mon-prog-bar"><div class="mon-prog-fill" style="--bar-w:${pct}%;background:${cor}"></div></div>
      </div>`;
  }

  function escalaEnviadaBadge(op) {
    const d = apontCache[op.id];
    if (!d || d === 'loading' || !d.listaEnviada) return '';
    return '<span class="mon-envelope" title="Lista enviada ao cliente">📋</span>';
  }

  function situacaoBadge(d, op) {
    if (!d || d === 'loading') return '<span class="mon-status-badge neutro">—</span>';
    const escOk = d.escalado >= d.solicitado;
    const aptOk = d.apontado >= d.solicitado;
    if (d._soEscala) {
      if (d.escalado === 0) return '<span class="mon-status-badge nenhum">✗ Nenhum</span>';
      if (escOk)            return '<span class="mon-status-badge esc-ok">✓ Esc. ok</span>';
      return `<span class="mon-status-badge nenhum">Esc. ${d.escalado}/${d.solicitado}</span>`;
    }
    if (aptOk && escOk) return '<span class="mon-status-badge completo">✓ Completo</span>';
    if (d.apontado === 0 && d.escalado === 0) return '<span class="mon-status-badge nenhum">✗ Nenhum</span>';
    const listaEnviada = d.listaEnviada || (op && apontCache[op.id] && apontCache[op.id].listaEnviada);
    if (d.apontado === 0 && listaEnviada) return escOk ? '<span class="mon-status-badge esc-ok">✓ Esc. ok</span>' : `<span class="mon-status-badge nenhum">Esc. ${d.escalado}/${d.solicitado}</span>`;
    if (d.apontado === 0 && escOk)        return '<span class="mon-status-badge esc-ok">✓ Esc. ok</span>';
    if (d.apontado === 0)                 return `<span class="mon-status-badge nenhum">Esc. ${d.escalado}/${d.solicitado}</span>`;
    return `<span class="mon-status-badge parcial">△ Apt ${d.apontado}/${d.solicitado}</span>`;
  }

  // ── MÉTRICAS ──────────────────────────────────────────────────────────────────
  function updateMetrics() {
    let ok = 0, inc = 0, zero = 0;
    operations.forEach(op => {
      if (!naJanela(op)) return;
      const d = apontCache[op.id];
      if (!d || d === 'loading') return;
      if (d.apontado === 0) zero++;
      else if (d.apontado >= d.solicitado) ok++;
      else inc++;
    });
    const s = (id, v) => { const el = document.getElementById(id); if (el) el.textContent = v; };
    s('m-total', operations.length);
    s('m-ok',    ok);
    s('m-inc',   inc);
    s('m-zero',  zero);
  }

  function setLive(state, label) {
    const el = document.getElementById('mon-live');
    if (!el) return;
    el.dataset.state = state;
    el.querySelector('span:last-child').textContent = label;
  }

  // ── RENDER ────────────────────────────────────────────────────────────────────
  function renderTable() {
    const tbody = document.getElementById('mon-tbody');
    if (!tbody) return;
    tbody.innerHTML = '';
    if (operations.length === 0) {
      tbody.innerHTML = `<tr><td colspan="8" style="text-align:center;padding:3rem;color:var(--mon-text-faint);font-size:13px">Nenhuma operação encontrada</td></tr>`;
      updateMetrics();
      return;
    }

    const visibleOps = getVisibleOps();

    // Atualiza contador no input placeholder
    const inp = document.getElementById('mon-filter-input');
    if (inp && !filterText) inp.placeholder = `Filtrar por chave, sigla ou site… (${operations.length} ops)`;

    if (visibleOps.length === 0) {
      tbody.innerHTML = `<tr><td colspan="8" style="text-align:center;padding:3rem;color:var(--mon-text-faint);font-size:13px">Nenhuma operação corresponde ao filtro</td></tr>`;
      updateMetrics();
      return;
    }

    visibleOps.forEach((op, idx) => {
      const isExp    = expanded.has(op.chave);
      const d        = apontCache[op.id];
      const emJanela = naJanela(op);
      const temDados = d && d !== 'loading';

      const escPct = temDados && op.qtd > 0 ? Math.min(100, Math.round((d.escalado / op.qtd) * 100)) : 0;
      const aptPct = temDados && emJanela && op.qtd > 0 ? Math.min(100, Math.round((d.apontado / op.qtd) * 100)) : 0;
      const escCor = colorForPct(escPct);
      const aptCor = colorForAptPct(aptPct);

      // Highlight do texto filtrado
      const hl = (txt) => {
        if (!filterText) return txt;
        const q = filterText.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
        return txt.replace(new RegExp('(' + q + ')', 'gi'), '<mark style="background:rgba(129,140,248,0.3);color:var(--mon-indigo);border-radius:2px">$1</mark>');
      };

      const tr = document.createElement('tr');
      tr.className = 'op-row' + (isExp ? ' is-expanded' : '');
      tr.dataset.chave = op.chave;

      tr.innerHTML = `
        <td><span class="mon-chave" title="${op.chave}">${hl(op.chave)}</span></td>
        <td><span class="mon-sigla">${hl(op.sigla)}</span></td>
        <td>
          ${op.id
            ? (temDados
                ? `<div class="mon-prog-cell">
                     <div class="mon-prog-num" style="color:${escCor}">${d.escalado}<span style="color:var(--mon-text-faint);font-size:10px;font-weight:400">/${op.qtd}</span></div>
                     <div class="mon-prog-bar"><div class="mon-prog-fill" style="--bar-w:${escPct}%;background:${escCor}"></div></div>
                   </div>`
                : `<span class="mon-prog-pending">…/${op.qtd}</span>`)
            : '<span class="mon-prog-na">—</span>'}
        </td>
        <td>
          ${emJanela
            ? (temDados
                ? `<div class="mon-prog-cell">
                     <div class="mon-prog-num" style="color:${aptCor}">${d.apontado}<span style="color:var(--mon-text-faint);font-size:10px;font-weight:400">/${op.qtd}</span></div>
                     <div class="mon-prog-bar"><div class="mon-prog-fill" style="--bar-w:${aptPct}%;background:${aptCor}"></div></div>
                   </div>`
                : `<span class="mon-prog-pending">…/${op.qtd}</span>`)
            : '<span class="mon-prog-na">—</span>'}
        </td>
        <td><span class="mon-hora">${op.hora}</span></td>
        <td><span class="mon-lider">${hl(op.lider)}</span></td>
        <td style="text-align:center">
          ${situacaoBadge(temDados ? d : null, op)}${escalaEnviadaBadge(op)}
        </td>
        <td style="text-align:center"><span class="mon-chevron">▼</span></td>
      `;
      tr.onclick = () => toggleRow(op, idx);
      tbody.appendChild(tr);

      if (isExp) {
        const det = document.createElement('tr');
        det.id = 'det-' + idx;
        det.className = 'op-detail';
        det.innerHTML = `<td colspan="8"><div class="mon-detail-inner">${renderDetail(op)}</div></td>`;
        tbody.appendChild(det);
      }
    });
    updateMetrics();
  }

  function renderDetail(op) {
    const d = apontCache[op.id];
    if (!d || d === 'loading') return `
      <div class="mon-loading-detail">
        <div class="mon-loading-spinner"></div>
        Carregando dados…
      </div>`;

    const escPct = op.qtd > 0 ? Math.min(100, Math.round((d.escalado / op.qtd) * 100)) : 0;
    const aptPct = op.qtd > 0 ? Math.min(100, Math.round((d.apontado / op.qtd) * 100)) : 0;
    const escCor = colorForPct(escPct);
    const aptCor = colorForAptPct(aptPct);

    const chaveEsc = op.chave.replace(/'/g, "\\'").replace(/"/g, '&quot;');

    let html = `
      <div class="mon-detail-header">
        <span class="mon-key-chip">${op.chave}</span>
        <button class="mon-copy-btn"
          onclick="event.stopPropagation();(function(btn){navigator.clipboard.writeText('${chaveEsc}').then(()=>{btn.textContent='✓ Copiado';btn.style.color='var(--mon-green)';setTimeout(()=>{btn.textContent='⎘ Copiar chave';btn.style.color='';},1800)}).catch(()=>{btn.textContent='✗ Erro';setTimeout(()=>{btn.textContent='⎘ Copiar chave';btn.style.color='';},1800)})})(this)">
          ⎘ Copiar chave
        </button>
        <button class="mon-copy-btn mon-open-btn" onclick="event.stopPropagation();var el=window._monLinkEls&&window._monLinkEls['${op.id}'];if(el){loadiframe('planejamento-operacional-edit${op.id}_3','Editar Planejamento',570,'modal1500');if(window.$)$('#modal1500').modal('show');}">🔎 Abrir OP</button>
      </div>

      <div class="mon-stat-grid">
        <div class="mon-stat-card">
          <div class="mon-stat-card-header">
            <span class="mon-stat-card-label">Escalados</span>
            <span class="mon-stat-card-num" style="color:${escCor}">${d.escalado}<span class="mon-stat-card-sub">/${op.qtd}</span></span>
          </div>
          <div class="mon-prog-bar"><div class="mon-prog-fill" style="--bar-w:${escPct}%;background:${escCor}"></div></div>
          <div class="mon-stat-card-pct">${escPct}% escalado</div>
        </div>
        <div class="mon-stat-card">
          <div class="mon-stat-card-header">
            <span class="mon-stat-card-label">Apontados</span>
            <span class="mon-stat-card-num" style="color:${aptCor}">${d.apontado}<span class="mon-stat-card-sub">/${op.qtd}</span></span>
          </div>
          <div class="mon-prog-bar"><div class="mon-prog-fill" style="--bar-w:${aptPct}%;background:${aptCor}"></div></div>
          <div class="mon-stat-card-pct">${aptPct}% apontado</div>
        </div>
      </div>
    `;

    const pdfLinks = d.pdfLinks || [], xlsLinks = d.xlsLinks || [];
    const qtdApt = d.apontado || 0;
    const temEscala = d.escalado > 0;
    const escalaAtualizada = d.listaEnviada === true;
    html += `<div class="mon-actions">
      <button class="mon-send-btn" onclick="event.stopPropagation();if(confirm('Confirmar envio de escala?'))window._monEnviarEscala('${op.id}',this)">✓ Escala enviada</button>
      <button class="mon-send-btn mon-send-btn--report" onclick="event.stopPropagation();if(confirm('Confirmar envio de report?'))window._monEnviarReport('${op.id}',this)" title="Preenche P1–P11 e coloca ${qtdApt} apontamentos no P10">📋 Report enviado</button>
      <button class="mon-send-btn mon-send-btn--print" onclick="event.stopPropagation();window._monPrintApontamentos('${op.id}',this)">🖨️ Print apontamentos</button>
      <button class="mon-send-btn mon-send-btn--relatorio" onclick="event.stopPropagation();window._monGerarRelatorio('${op.id}',this)">📲 ${escalaAtualizada && d.todosConfirmados ? 'Report WhatsApp (atualizado)' : 'Report WhatsApp'}</button>
      ${temEscala ? `<button class="mon-send-btn mon-send-btn--wa-escala" onclick="event.stopPropagation();window._monGerarMsgEscala('${op.id}',this)">💬 ${escalaAtualizada ? 'Msg escala (atualizada)' : 'Msg escala WA'}</button><button class="mon-send-btn mon-send-btn--gmail" onclick="event.stopPropagation();window._monAbrirGmailEscala('${op.id}',this)" onmouseenter="(function(btn){var emails=window._monEmailsDaOpById&&window._monEmailsDaOpById('${op.id}');btn.title=emails&&emails.length?'Para: '+emails.join(', '):'⚠ Nenhum destinatário cadastrado';})(this)">${'<svg width="14" height="11" viewBox="0 0 24 18" fill="none" xmlns="http://www.w3.org/2000/svg" style="vertical-align:middle;margin-right:4px"><path d="M0 3.6L12 11.4L24 3.6V16.8C24 17.46 23.46 18 22.8 18H1.2C0.54 18 0 17.46 0 16.8V3.6Z" fill="#EA4335"/><path d="M0 3.6L12 11.4L24 3.6" fill="#FBBC05"/><path d="M0 1.2C0 0.54 0.54 0 1.2 0H22.8C23.46 0 24 0.54 24 1.2V3.6L12 11.4L0 3.6V1.2Z" fill="#4285F4"/><path d="M0 3.6V1.2C0 0.54 0.54 0 1.2 0L12 7.8L0 3.6Z" fill="#34A853"/><path d="M24 3.6V1.2C24 0.54 23.46 0 22.8 0L12 7.8L24 3.6Z" fill="#EA4335"/></svg>'} ${escalaAtualizada ? 'Gmail (atualizado)' : 'Gmail escala'}</button>` : ''}`;
    pdfLinks.forEach(l => {
      html += `<a href="https://tsi-app.com/${l.href}" target="_blank" class="mon-dl-btn">📄 ${l.label || 'Assinatura'}</a>`;
    });
    if (xlsLinks.length > 0) {
      html += `<div class="mon-xls-menu" id="xls-menu-${op.id}">
        <button class="mon-dl-btn" onclick="event.stopPropagation();var m=document.getElementById('xls-menu-${op.id}');m.classList.toggle('open');var close=function(ev){if(!m.contains(ev.target)){m.classList.remove('open');document.removeEventListener('click',close);}};setTimeout(function(){document.addEventListener('click',close);},0);">📊 XLS ▾</button>
        <div class="mon-xls-dropdown">`;
      xlsLinks.forEach(l => { html += `<a href="https://tsi-app.com/${l.href}" target="_blank" onclick="event.stopPropagation();document.getElementById('xls-menu-${op.id}').classList.remove('open');">${l.label}</a>`; });
      html += `</div></div>`;
    }
    html += `</div>`;

    const faltando = d.faltando || [];
    const colab = d.colaboradores || [];

    if (colab.length === 0 && faltando.length === 0) {
      html += '<div class="mon-empty-detail">Sem dados de escala/apontamento.</div>';
    } else {
      html += `<div class="mon-lists-grid">`;

      // ── APONTADOS (esquerda) ──
      html += `
        <div class="mon-list-panel mon-list-panel--ok">
          <div class="mon-list-panel-header">
            <span class="mon-list-panel-dot" style="background:var(--mon-green)"></span>
            <span class="mon-list-panel-title">Apontados</span>
            <span class="mon-list-panel-count" style="background:var(--mon-green-bg);color:var(--mon-green)">${colab.length}</span>
          </div>
          <div class="mon-list-panel-body">`;
      if (colab.length === 0) {
        html += `<div class="mon-list-empty">Nenhum apontamento ainda</div>`;
      } else {
        colab.forEach(c => {
          const kmColor = c.dist > 0 ? (c.dist < 1 ? 'var(--mon-green)' : 'var(--mon-red)') : '';
          const kmBg    = c.dist > 0 ? (c.dist < 1 ? 'var(--mon-green-bg)' : 'var(--mon-red-bg)') : '';
          const kmBorder= c.dist > 0 ? (c.dist < 1 ? 'var(--mon-green-border,rgba(22,163,74,0.25))' : 'var(--mon-red-border,rgba(220,38,38,0.25))') : '';
          const kmTag = c.dist > 0 ? `<span class="mon-list-row-tipo" style="background:${kmBg};color:${kmColor};border:1px solid ${kmBorder};border-radius:4px;padding:1px 6px;font-weight:600">📍 ${c.dist.toFixed(2)} km</span>` : '';
          html += `
            <div class="mon-list-row">
              <div class="mon-list-row-name">${c.nome}</div>
              <div class="mon-list-row-meta">
                <span class="mon-list-row-tipo">${c.tipo||'—'}</span>
                <span class="mon-list-row-time">${c.inicio}</span>
                ${kmTag}
              </div>
            </div>`;
        });
      }
      html += `</div></div>`;

      // ── FALTANDO (direita) ──
      html += `
        <div class="mon-list-panel mon-list-panel--warn">
          <div class="mon-list-panel-header">
            <span class="mon-list-panel-dot" style="background:var(--mon-red)"></span>
            <span class="mon-list-panel-title">Faltando</span>
            <span class="mon-list-panel-count" style="background:var(--mon-red-bg);color:var(--mon-red)">${faltando.length}</span>
            ${(d.faltasConfirmadas||[]).length > 0 ? '<span onclick="event.stopPropagation();var b=document.getElementById(\'fc-'+op.id+'\');var a=this.querySelector(\'.fc-arr\');if(b.style.display===\'none\'){b.style.display=\'block\';a.style.transform=\'rotate(180deg)\';}else{b.style.display=\'none\';a.style.transform=\'\';}" style="display:inline-flex;align-items:center;gap:4px;cursor:pointer;user-select:none;margin-left:auto;background:var(--mon-red-bg);border:1px solid var(--mon-red-border,rgba(220,38,38,0.25));border-radius:5px;padding:2px 8px;font-size:10px;font-weight:700;color:var(--mon-red)">⚠ Faltas ('+((d.faltasConfirmadas||[]).length)+')<span class="fc-arr" style="transition:transform 0.2s;font-size:9px">▼</span></span>' : ''}
          </div>
          <div id="fc-${op.id}" style="display:none;padding:6px 10px;border-bottom:1px solid var(--mon-border)">
            ${(d.faltasConfirmadas||[]).map(c => '<div class="mon-list-row" style="background:var(--mon-red-bg);border-radius:6px;margin-bottom:3px;border:none"><div class="mon-list-row-name" style="color:var(--mon-red)">'+c.nome+'</div><div class="mon-list-row-meta"><span class="mon-list-row-tipo">'+(c.tipo||'—')+'</span>'+(c.inicio ? '<span class="mon-list-row-tipo">🕐 '+c.inicio+'</span>' : '')+'</div></div>').join('')}
          </div>
          <div class="mon-list-panel-body">`;
      if (faltando.length === 0) {
        html += `<div class="mon-list-empty" style="color:var(--mon-green)">✓ Todos apontados</div>`;
      } else {
        faltando.forEach(c => {
          html += `
            <div class="mon-list-row">
              <div class="mon-list-row-name" style="color:var(--mon-red)">${c.nome}</div>
              <div class="mon-list-row-meta">
                <span class="mon-list-row-tipo">${c.tipo||'—'}</span>
                ${c.advisor ? `<span class="mon-list-row-tipo" style="color:var(--mon-text-secondary)">👤 ${c.advisor}</span>` : ''}
                ${c.datacad ? `<span class="mon-list-row-tipo" style="color:var(--mon-text-secondary)">🕐 ${c.datacad}</span>` : ''}
              </div>
            </div>`;
        });

      }
      html += `</div></div>`;

      html += `</div>`;
    }
    return html;
  }

  function toggleRow(op, idx) {
    const isOpen = expanded.has(op.chave);
    expanded.clear();
    if (!isOpen) {
      expanded.add(op.chave);
    }
    renderTable();
    if (!isOpen) {
      setTimeout(() => {
        const det = document.getElementById('det-' + idx);
        if (det) det.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
      }, 50);
      const cached = apontCache[op.id];
      if (!cached || cached === 'loading') {
        const poll = setInterval(() => {
          const c = apontCache[op.id];
          if (c && c !== 'loading') {
            clearInterval(poll);
            const det = document.getElementById('det-' + idx);
            if (det) det.querySelector('.mon-detail-inner').innerHTML = renderDetail(op);
            updateMetrics();
          }
        }, 500);
        setTimeout(() => clearInterval(poll), 40000);
      }
    }
  }

  // ── TOGGLE ────────────────────────────────────────────────────────────────────
  function toggleMonitor() {
    if (!document.getElementById('mon-panel')) createPanel();
    const p   = document.getElementById('mon-panel');
    const btn = document.getElementById('btn-mon');
    if (p.style.display === 'none' || !p.style.display) {
      p.style.display = 'flex'; p.style.flexDirection = 'column';
      btn.innerHTML = '<span class="mon-fab-dot"></span> Monitor';
    } else {
      p.style.display = 'none';
      btn.innerHTML = '<span class="mon-fab-dot"></span> Monitor';
    }
  }

  // ── AMPULHETA ANIMADA (countdown visual) ──────────────────────────────────────
  let _hgInterval = null;

  function startCountdown(totalSec) {
    if (totalSec === undefined) {
      const agora = new Date();
      totalSec = 60 - agora.getSeconds();
    }
    if (totalSec <= 0) totalSec = 60;

    if (_hgInterval) clearInterval(_hgInterval);

    let remaining = totalSec;
    // Alterna entre ⏳ e ⌛ para dar sensação de movimento
    let flip = false;

    const tick = () => {
      const count = document.getElementById('mon-hg-count');
      const icon  = document.getElementById('mon-hg-icon');
      if (!count) return;

      count.textContent = remaining + 's';
      if (icon) {
        flip = !flip;
        icon.textContent = flip ? '⌛' : '⏳';
      }
      remaining--;

      if (remaining < 0) {
        clearInterval(_hgInterval);
        _hgInterval = null;
        if (count) count.textContent = '—';
        if (icon)  icon.textContent  = '⏳';
      }
    };

    tick();
    _hgInterval = setInterval(tick, 1000);
  }

  // ── TIMER ALINHADO AO MINUTO FECHADO ─────────────────────────────────────────
  let _alignTimeout = null;

  function scheduleAlignedRefresh() {
    if (refreshTimer)    clearInterval(refreshTimer);
    if (_alignTimeout)   clearTimeout(_alignTimeout);
    refreshTimer   = null;
    _alignTimeout  = null;

    const agora            = new Date();
    const msAteProxMinuto  = (60 - agora.getSeconds()) * 1000 - agora.getMilliseconds();

    _alignTimeout = setTimeout(() => {
      _alignTimeout = null;
      silentRefresh();
      startCountdown(); // reinicia ampulheta
      refreshTimer = setInterval(() => {
        silentRefresh();
        startCountdown();
      }, 60 * 1000);
    }, msAteProxMinuto);

    // Inicia a ampulheta com o tempo real até o próximo minuto
    startCountdown(Math.round(msAteProxMinuto / 1000));
  }

  function startMonitor() {
    window._monRunning = true;
    monCarregarContatos();
    fetchOperations();
    scheduleAlignedRefresh();
    startWatchdog();
  }

  // ── DETECTA PÁGINA ATUAL DA URL ─────────────────────────────────────────────
  function getCurrentPageSuffix() {
    const match = window.location.href.match(/planejamento-operacional(_\d+)?/);
    if (!match) return '';
    return match[1] || ''; // ex: '' = pag1, '_2' = pag2, '_3' = pag3, etc.
  }

  function getCurrentPageUrl() {
    const suffix = getCurrentPageSuffix();
    return 'https://tsi-app.com/planejamento-operacional' + suffix;
  }

  function fetchOperations() {
    setLive('sync', 'Sincronizando…');

    const finalizar = (opsAll) => {
      const seen = new Set();
      const ops  = opsAll.filter(o => { if (seen.has(o.chave)) return false; seen.add(o.chave); return true; });

      const savedCache = cacheLoad();
      const prevCache  = apontCache;

      operations  = ops;
      window._monOps = ops;
      apontCache  = {};
      ops.forEach(o => {
        if (o.id) {
          const mem = prevCache[o.id];
          const ses = savedCache[o.id];
          if (mem && mem !== 'loading') apontCache[o.id] = mem;
          else if (ses) apontCache[o.id] = ses;
        }
      });

      expanded    = new Set([...expanded].filter(c => ops.some(o => o.chave === c)));
      monitoradas = new Set();
      fetchQueue  = [];
      inQueue     = new Set();

      const opsComId = ops.filter(o => o.id);
      opsComId.forEach(o => { if (dentroJanela(o)) monitoradas.add(monKey(o)); });
      renderTable();
      setLive('live', 'Ao vivo');
      const sub = document.getElementById('mon-sub');
      if (sub) sub.textContent = 'Atualizado ' + new Date().toLocaleTimeString('pt-BR');

      const total = opsComId.length;
      let loaded = 0;

      opsComId.forEach(op => {
        if (apontCache[op.id]) {
          apontCache[op.id]._stale = true;
          loaded++;
        }
      });
      updateProgress(loaded, total);
      if (loaded > 0) updateMetrics();

      opsComId.forEach((op) => {
        const hadCache = !!apontCache[op.id];
        enfileirar(op, (novo) => {
          if (dentroJanela(op)) monitoradas.add(monKey(op));
          if (!hadCache) loaded++;
          updateProgress(loaded, total);
          updateCells(op, novo, null);
          updateMetrics();
          cacheSave();
        }, true);
      });
    };

    // Aguarda a tabela ter linhas (DOM pode ainda estar renderizando após navegação SPA)
    let tentativas = 0;
    const tentarParsear = () => {
      const opsDoc = parseOpsFromDoc(document);
      if (opsDoc.length > 0) {
        console.log('[Monitor] Ops encontradas na página:', opsDoc.length);
        finalizar(opsDoc);
        return;
      }
      tentativas++;
      if (tentativas < 10) {
        // Tenta de novo em 800ms (até ~8s no total)
        setTimeout(tentarParsear, 800);
      } else {
        console.warn('[Monitor] Nenhuma op encontrada após 10 tentativas');
        finalizar([]);
      }
    };

    tentarParsear();
  }

  // ── DETECTA NAVEGAÇÃO SPA (mudança de URL sem reload) ────────────────────────
  function watchPageNavigation() {
    let lastUrl = window.location.href;

    // Observa mudanças no DOM que indiquem troca de página (paginação SPA)
    const observer = new MutationObserver(() => {
      const currentUrl = window.location.href;
      if (currentUrl !== lastUrl) {
        lastUrl = currentUrl;
        // Só age se ainda estiver em planejamento-operacional
        if (!currentUrl.includes('planejamento-operacional')) return;
        console.log('[Monitor] Navegação detectada:', currentUrl);
        // Aguarda a nova página renderizar e então recarrega as ops
        setTimeout(() => {
          iframesInUse = {};
          fetchQueue   = [];
          inQueue      = new Set();
          apontCache   = {};
          fetchOperations();
        }, 2000);
      }
    });

    observer.observe(document.body, { childList: true, subtree: true });

    // Também intercepta pushState/replaceState (navegação SPA via history API)
    const origPush    = history.pushState.bind(history);
    const origReplace = history.replaceState.bind(history);

    history.pushState = function(...args) {
      origPush(...args);
      setTimeout(() => {
        const url = window.location.href;
        if (url !== lastUrl && url.includes('planejamento-operacional')) {
          lastUrl = url;
          console.log('[Monitor] pushState detectado:', url);
          setTimeout(() => {
            iframesInUse = {};
            fetchQueue   = [];
            inQueue      = new Set();
            apontCache   = {};
            fetchOperations();
          }, 2000);
        }
      }, 100);
    };

    history.replaceState = function(...args) {
      origReplace(...args);
      setTimeout(() => {
        const url = window.location.href;
        if (url !== lastUrl && url.includes('planejamento-operacional')) {
          lastUrl = url;
          console.log('[Monitor] replaceState detectado:', url);
          setTimeout(() => {
            iframesInUse = {};
            fetchQueue   = [];
            inQueue      = new Set();
            apontCache   = {};
            fetchOperations();
          }, 2000);
        }
      }, 100);
    };
  }

  // ── INIT ──────────────────────────────────────────────────────────────────────
  setTimeout(() => {
    ensureIframes();
    injectButton();
    createPanel();
    if (!window._monRunning) startMonitor();
    watchPageNavigation();
    if (Notification.permission === 'default') pedirPermissaoNotificacao();
  }, 2000);

  // ══════════════════════════════════════════════════════════════════════════════
  // ADDON v20: MENSAGEM DE ESCALA (WA) + GMAIL DE ESCALA
  // ══════════════════════════════════════════════════════════════════════════════

  // ── CONTATOS (JSONBin — mesmo backend do Gerador) ────────────────────────────
  const _MON_BIN_ID  = '69dd9cfa36566621a8ae40e1';
  const _MON_BIN_KEY = '$2a$10$re7SEj86dL3mQnxKBMLFvu7f566NmQucI1RwyW5t9tfYCrCQUExt.';
  const _MON_BIN_URL = 'https://api.jsonbin.io/v3/b/' + _MON_BIN_ID;
  let _monContatos = null;

  function monCarregarContatos() {
    try {
      const local = localStorage.getItem('tsi_contatos');
      if (local) _monContatos = JSON.parse(local);
    } catch(e) {}
    fetch(_MON_BIN_URL + '/latest', { headers: { 'X-Master-Key': _MON_BIN_KEY } })
      .then(r => r.json())
      .then(j => {
        if (j.record && j.record.contatos && Object.keys(j.record.contatos).length > 0) {
          _monContatos = j.record.contatos;
          try { localStorage.setItem('tsi_contatos', JSON.stringify(_monContatos)); } catch(e) {}
        }
      })
      .catch(() => {});
  }

  function monEmailsDaOp(op) {
    if (!_monContatos || !op.chave) return [];
    const chaveUp = op.chave.toUpperCase();
    const unidade = Object.keys(_monContatos).find(k => chaveUp.startsWith(k.toUpperCase()));
    if (!unidade) return [];
    const u = _monContatos[unidade];
    return (u && Array.isArray(u.emails)) ? u.emails : [];
  }

  // Expõe lookup por opId para uso no tooltip inline do botão Gmail
  window._monEmailsDaOpById = function(opId) {
    const op = operations.find(o => o.id === opId);
    if (!op) return [];
    return monEmailsDaOp(op);
  };

  // ── NOME DO USUÁRIO LOGADO ────────────────────────────────────────────────────
  function monNomeUsuario() {
    try {
      const el = document.querySelector('.headertop-nomeappsub');
      if (!el) return '';
      const nomeCompleto = el.textContent.trim();
      const primeiro = nomeCompleto.split(' ')[0] || '';
      return primeiro.charAt(0).toUpperCase() + primeiro.slice(1).toLowerCase();
    } catch(e) { return ''; }
  }

  // ── SAUDAÇÃO ──────────────────────────────────────────────────────────────────
  function monSaudacao() {
    const h = new Date().getHours();
    return h < 12 ? 'Bom dia' : h < 18 ? 'Boa tarde' : 'Boa noite';
  }

  function monDataHoje() {
    const d = new Date();
    const dd = String(d.getDate()).padStart(2, '0');
    const mm = String(d.getMonth() + 1).padStart(2, '0');
    return dd + '/' + mm + '/' + d.getFullYear();
  }

  function monHoraFormatada(hora) {
    if (!hora) return '';
    return hora.replace(':', 'h').replace(/h00$/, 'h');
  }

  function monHoraAssunto(hora) {
    if (!hora) return '';
    return hora.replace(':', 'H');
  }

  // ── MENSAGEM WHATSAPP DE ESCALA ───────────────────────────────────────────────
  window._monGerarMsgEscala = function(opId, btnEl) {
    const d  = apontCache[opId];
    const op = operations.find(o => o.id === opId);
    if (!op) return;

    const atualizada = d && d !== 'loading' && d.listaEnviada === true;
    const data       = monDataHoje();
    const hora       = monHoraFormatada(op.hora);
    const saudacao   = monSaudacao();

    const texto = atualizada
      ? saudacao + ', time. Segue a escala TSI *atualizada* de *' + data + '* às *' + hora + '*.'
      : saudacao + ', time. Segue a escala TSI de *' + data + '* às *' + hora + '*.';

    navigator.clipboard.writeText(texto)
      .then(() => {
        const orig = btnEl.innerHTML;
        btnEl.innerHTML = '✅ Copiado!';
        btnEl.style.color = 'var(--mon-green)';
        setTimeout(() => { btnEl.innerHTML = orig; btnEl.style.color = ''; }, 2500);
      })
      .catch(() => { prompt('Copie a mensagem:', texto); });
  };

  // ── GMAIL DE ESCALA ───────────────────────────────────────────────────────────
  window._monAbrirGmailEscala = function(opId, btnEl) {
    const d  = apontCache[opId];
    const op = operations.find(o => o.id === opId);
    if (!op) return;

    const atualizada  = d && d !== 'loading' && d.listaEnviada === true;
    const data        = monDataHoje();
    const horaExib    = monHoraFormatada(op.hora);
    const horaAssunto = monHoraAssunto(op.hora);
    const nome        = monNomeUsuario();
    const assinatura  = '\n\nAtenciosamente,\n' + (nome ? nome + '\n' : '') + 'Assistente de Planejamento | TSI';

    const emails = monEmailsDaOp(op);
    const to     = emails.join(',');

    if (emails.length === 0) {
      const orig = btnEl.innerHTML;
      btnEl.innerHTML = '⚠ Sem destinatários';
      btnEl.style.color = 'var(--mon-amber)';
      btnEl.style.borderColor = 'var(--mon-amber-border)';
      setTimeout(() => { btnEl.innerHTML = orig; btnEl.style.color = ''; btnEl.style.borderColor = ''; }, 3500);
      return;
    }

    const assunto = atualizada
      ? 'TSI - ESCALA ATUALIZADA | ' + data + ' | ' + horaAssunto
      : 'TSI - ESCALA | ' + data + ' | ' + horaAssunto;

    const corpo = atualizada
      ? 'Bom dia,\n\nEncaminho a versão atualizada da escala TSI referente ao dia ' + data + ', turno das ' + horaExib + '. Pedimos que desconsiderem a versão anterior.\n\nQualquer dúvida, estou à disposição.' + assinatura
      : 'Bom dia,\n\nEncaminho a escala TSI referente ao dia ' + data + ', turno das ' + horaExib + ', para conhecimento e organização das atividades.\n\nSolicito a conferência das informações. Qualquer dúvida, estou à disposição.' + assinatura;

    const url = 'https://mail.google.com/mail/?view=cm&fs=1'
      + (to ? '&to=' + encodeURIComponent(to) : '')
      + '&su=' + encodeURIComponent(assunto)
      + '&body=' + encodeURIComponent(corpo);

    window.open(url, '_blank');

    const orig = btnEl.innerHTML;
    btnEl.innerHTML = '✅ Gmail aberto!';
    btnEl.style.color = 'var(--mon-green)';
    setTimeout(() => { btnEl.innerHTML = orig; btnEl.style.color = ''; }, 3000);
  };


})();
