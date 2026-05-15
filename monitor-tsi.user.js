// ==UserScript==
// @name         Monitor Operacional TSI
// @namespace    http://tampermonkey.net/
// @version      11.15
// @description  Monitor de apontamentos em tempo real com escalados vs apontados
// @author       TSI
// @match        https://tsi-app.com/planejamento-operacional*
// @grant        none
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
      const linkEl = row.querySelector('a[onclick*="planejamento-operacional-edit"]');
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
      ops.push({ chave: g(0), sigla: g(1), site: g(2), qtd: parseInt(g(3)) || 0, hora: g(9), lider: g(11), status: g(24).toLowerCase(), time: g(8), id, bubbles });
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
                const nome = cells[2]?.textContent?.trim();
                const cpf  = cells[3]?.textContent?.trim();
                if (!nome || nome.length < 3) return;
                escalados.push({ nome, cpf, tipo: cells[4]?.textContent?.trim() });
              });
            }
            const xlsLabels = ['Layout 1 (SHEIN)','Layout 2 (Cordovil)','Layout 3 (SBC)','Layout 4 (SBF)','Layout 5 (Endereço)','Layout 6 (KISOC)'];
            doc2.querySelectorAll('a[href*="escalaprelistaLiderPDF_"]').forEach(a => pdfLinks.push({ label: a.textContent.trim() || 'Lista p/ Assinaturas', href: a.getAttribute('href') }));
            doc2.querySelectorAll('a[href*="escalaprelistaLiderXLS"]').forEach((a, i) => xlsLinks.push({ label: xlsLabels[i] || a.textContent.trim(), href: a.getAttribute('href') }));

            if (!naJanela(op)) {
              release({ solicitado: op.qtd, escalado: escalados.length, apontado: 0, colaboradores: [], escalados, faltando: escalados, pdfLinks, xlsLinks, _soEscala: true, listaEnviada, todosConfirmados });
              return;
            }

            // 3) Busca apontamentos
            return fetchDoc('https://tsi-app.com/' + eaptHref)
              .then(doc3 => {
                const colaboradores = [];
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
                    if (origem === 'FALTA') return;
                    if (!inicio) return;
                    colaboradores.push({ nome, cpf, tipo: cells[2]?.textContent?.trim(), inicio });
                  });
                }
                const apontadosCPF = new Set(colaboradores.map(c => c.cpf));
                const faltando = escalados.filter(e => !apontadosCPF.has(e.cpf));
                release({ solicitado: op.qtd, escalado: escalados.length, apontado: colaboradores.length, colaboradores, escalados, faltando, pdfLinks, xlsLinks, listaEnviada, todosConfirmados });
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
      setTimeout(() => {
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
            }, 2500);
          }, 700);
        } catch(e) { fail('erro'); clearTimeout(safetyTimer); }
      }, 2000);
    };
  }

  window._monEnviarEscala = enviarEscala;

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
      setTimeout(() => {
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
            }, 2500);
          }, 700);
        } catch(e) { fail('erro'); clearTimeout(safetyTimer); }
      }, 2000);
    };
  }

  window._monEnviarReport = enviarReport;

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
    refreshTimer = setInterval(silentRefresh, 60 * 1000);
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
    if (cells[3]) cells[3].innerHTML = escaladoBadge(d, op.qtd);
    if (naJanela(op)) {
      if (cells[4]) cells[4].innerHTML = apontBadge(d, op.qtd);
    }
    if (cells[7]) cells[7].innerHTML = situacaoBadge(d, op) + escalaEnviadaBadge(op);

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
  function initControls(panel) {
    panel.style.position = 'fixed';
    const rh = document.createElement('div');
    rh.className = 'mon-resize-handle';
    rh.addEventListener('mousedown', e => {
      e.preventDefault();
      const startX = e.clientX, startW = panel.offsetWidth;
      document.body.style.userSelect = 'none';
      const mv = e => { panel.style.width = Math.min(Math.max(startW + (startX - e.clientX), 420), window.innerWidth - 80) + 'px'; };
      const up = () => { document.body.style.userSelect = ''; document.removeEventListener('mousemove', mv); document.removeEventListener('mouseup', up); };
      document.addEventListener('mousemove', mv);
      document.addEventListener('mouseup', up);
    });
    panel.appendChild(rh);

    const header = panel.querySelector('#mon-header');
    if (header) {
      header.addEventListener('mousedown', e => {
        if (e.target.tagName === 'BUTTON' || e.target.closest('button')) return;
        e.preventDefault();
        header.style.cursor = 'grabbing';
        const rect = panel.getBoundingClientRect();
        const ox = e.clientX - rect.left, oy = e.clientY - rect.top;
        document.body.style.userSelect = 'none';
        const mv = e => {
          panel.style.left  = Math.max(0, Math.min(e.clientX - ox, window.innerWidth - panel.offsetWidth)) + 'px';
          panel.style.top   = Math.max(0, Math.min(e.clientY - oy, window.innerHeight - 60)) + 'px';
          panel.style.right = 'auto';
        };
        const up = () => {
          header.style.cursor = '';
          document.body.style.userSelect = '';
          document.removeEventListener('mousemove', mv);
          document.removeEventListener('mouseup', up);
        };
        document.addEventListener('mousemove', mv);
        document.addEventListener('mouseup', up);
      });
    }
  }

  window._monMinimize = function() {
    const body  = document.getElementById('mon-body');
    const btn   = document.getElementById('mon-min-btn');
    const panel = document.getElementById('mon-panel');
    if (!body || !btn || !panel) return;
    minimized = !minimized;
    body.style.display = minimized ? 'none' : '';
    panel.style.height = minimized ? 'auto' : '100vh';
    btn.innerHTML = minimized ? '&#9633;' : '&#8212;';
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
      /* ── VARIÁVEIS ── */
      :root {
        --mon-bg:        #0c0c0f;
        --mon-surface:   #111116;
        --mon-surface2:  #16161d;
        --mon-border:    #1f1f2e;
        --mon-border2:   #2a2a3a;
        --mon-text:      #c8c8d8;
        --mon-text-dim:  #6b6b88;
        --mon-text-faint:#32324a;
        --mon-green:     #34d474;
        --mon-green-bg:  rgba(52,212,116,0.08);
        --mon-amber:     #f59e0b;
        --mon-amber-bg:  rgba(245,158,11,0.08);
        --mon-red:       #f87171;
        --mon-red-bg:    rgba(248,113,113,0.08);
        --mon-indigo:    #818cf8;
        --mon-indigo-bg: rgba(129,140,248,0.1);
        --mon-radius:    10px;
        --mon-radius-sm: 6px;
      }

      /* ── KEYFRAMES ── */
      @keyframes mon-shake {
        0%   { transform: rotate(-8deg) scale(1.04); }
        100% { transform: rotate(8deg)  scale(1.04); }
      }
      @keyframes mon-bar-fill {
        from { width: 0; }
        to   { width: var(--bar-w); }
      }
      @keyframes mon-fadein {
        from { opacity: 0; transform: translateY(6px); }
        to   { opacity: 1; transform: translateY(0); }
      }
      @keyframes mon-pulse-dot {
        0%, 100% { opacity: 1; box-shadow: 0 0 0 0 currentColor; }
        50%       { opacity: 0.7; box-shadow: 0 0 0 4px transparent; }
      }
      @keyframes mon-spin {
        to { transform: rotate(360deg); }
      }
      @keyframes mon-row-in {
        from { opacity: 0; transform: translateX(6px); }
        to   { opacity: 1; transform: translateX(0); }
      }

      /* ── BOTÃO FLUTUANTE ── */
      #btn-mon {
        position: fixed; bottom: 20px; right: 20px; z-index: 99999;
        display: flex; align-items: center; gap: 8px;
        background: var(--mon-surface);
        color: var(--mon-text);
        border: 1px solid var(--mon-border2);
        padding: 9px 18px; border-radius: 24px;
        font-size: 12px; font-family: 'Segoe UI', system-ui, sans-serif;
        font-weight: 600; cursor: pointer; letter-spacing: 0.5px;
        box-shadow: 0 4px 24px rgba(0,0,0,0.5), 0 0 0 1px rgba(255,255,255,0.04);
        transition: all 0.2s;
      }
      #btn-mon:hover {
        background: var(--mon-surface2);
        border-color: var(--mon-border2);
        box-shadow: 0 6px 32px rgba(0,0,0,0.6);
        transform: translateY(-1px);
      }
      .mon-fab-dot {
        width: 7px; height: 7px; border-radius: 50%;
        background: var(--mon-green);
        display: inline-block;
        animation: mon-pulse-dot 2s ease-in-out infinite;
      }

      /* ── PAINEL ── */
      #mon-panel {
        font-family: 'Segoe UI', system-ui, sans-serif;
        font-size: 12px;
        background: var(--mon-bg);
        color: var(--mon-text);
        border-left: 1px solid var(--mon-border);
        box-shadow: -12px 0 60px rgba(0,0,0,0.7);
      }

      /* ── RESIZE HANDLE ── */
      .mon-resize-handle {
        position: absolute; left: 0; top: 0; width: 4px; height: 100%;
        cursor: ew-resize; z-index: 10; background: transparent;
        transition: background 0.15s;
      }
      .mon-resize-handle:hover { background: var(--mon-indigo); opacity: 0.4; }

      /* ── SCROLLBAR ── */
      #mon-panel ::-webkit-scrollbar { width: 3px; }
      #mon-panel ::-webkit-scrollbar-track { background: transparent; }
      #mon-panel ::-webkit-scrollbar-thumb { background: var(--mon-border2); border-radius: 2px; }

      /* ── HEADER ── */
      #mon-header {
        background: var(--mon-surface);
        border-bottom: 1px solid var(--mon-border);
        padding: 0 16px;
        height: 58px;
        display: flex; align-items: center; justify-content: space-between;
        flex-shrink: 0; user-select: none; cursor: grab;
      }
      #mon-header:active { cursor: grabbing; }
      .mon-logo {
        display: flex; align-items: center; gap: 10px;
      }
      .mon-logo-icon {
        width: 30px; height: 30px; border-radius: 8px;
        background: linear-gradient(135deg, var(--mon-indigo), #6366f1);
        display: flex; align-items: center; justify-content: center;
        font-size: 14px; color: #fff; font-weight: 700;
        flex-shrink: 0;
      }
      .mon-logo-text {
        display: flex; flex-direction: column; gap: 1px;
      }
      .mon-logo-title {
        font-size: 12px; font-weight: 700; color: #e8e8f0;
        letter-spacing: 0.8px; text-transform: uppercase;
      }
      .mon-logo-sub {
        font-size: 10px; color: var(--mon-text-dim);
      }

      /* ── STATUS PILL ── */
      .mon-status-pill {
        display: inline-flex; align-items: center; gap: 5px;
        padding: 3px 10px; border-radius: 20px;
        border: 1px solid;
        font-size: 10px; font-weight: 600; letter-spacing: 0.5px;
      }
      .mon-status-pill[data-state="live"] {
        color: var(--mon-green); border-color: rgba(52,212,116,0.25);
        background: var(--mon-green-bg);
      }
      .mon-status-pill[data-state="sync"] {
        color: var(--mon-amber); border-color: rgba(245,158,11,0.25);
        background: var(--mon-amber-bg);
      }
      .mon-status-pill[data-state="offline"] {
        color: var(--mon-text-faint); border-color: var(--mon-border);
        background: transparent;
      }
      .mon-status-dot {
        width: 5px; height: 5px; border-radius: 50%; background: currentColor;
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
        height: 30px; padding: 0 12px; border-radius: var(--mon-radius-sm);
        border: 1px solid var(--mon-border2); background: transparent;
        color: var(--mon-text-dim); font-size: 11px; font-family: inherit;
        font-weight: 500; cursor: pointer; transition: all 0.15s;
        display: inline-flex; align-items: center; gap: 6px;
      }
      .mon-hdr-btn:hover {
        background: var(--mon-surface2); color: var(--mon-text);
        border-color: var(--mon-border2);
      }
      .mon-hdr-btn--green { color: var(--mon-green); border-color: rgba(52,212,116,0.25); }
      .mon-hdr-btn--green:hover { background: var(--mon-green-bg); color: var(--mon-green); }
      .mon-hdr-btn[data-state="on"] { color: var(--mon-green); }
      .mon-hdr-btn[data-state="off"] { color: var(--mon-red); }

      .mon-icon-btn {
        width: 30px; height: 30px; border-radius: var(--mon-radius-sm);
        border: 1px solid var(--mon-border); background: transparent;
        color: var(--mon-text-dim); font-size: 14px; cursor: pointer;
        display: inline-flex; align-items: center; justify-content: center;
        transition: all 0.15s;
      }
      .mon-icon-btn:hover { background: var(--mon-surface2); color: var(--mon-text); }

      /* ── MÉTRICAS ── */
      #mon-metrics {
        display: grid; grid-template-columns: repeat(4, 1fr);
        border-bottom: 1px solid var(--mon-border);
        flex-shrink: 0;
      }
      .mon-metric {
        padding: 14px 16px; position: relative;
        border-right: 1px solid var(--mon-border);
      }
      .mon-metric:last-child { border-right: none; }
      .mon-metric-label {
        font-size: 9px; color: var(--mon-text-dim); letter-spacing: 1px;
        text-transform: uppercase; margin-bottom: 6px; font-weight: 600;
      }
      .mon-metric-val {
        font-size: 26px; font-weight: 700; line-height: 1;
        color: var(--mon-text-dim);
      }
      .mon-metric-val.green  { color: var(--mon-green); }
      .mon-metric-val.amber  { color: var(--mon-amber); }
      .mon-metric-val.red    { color: var(--mon-red); }
      .mon-metric-bar {
        position: absolute; bottom: 0; left: 0;
        height: 2px; background: currentColor; opacity: 0.3;
        transition: width 0.8s ease;
      }

      /* ── TABELA ── */
      #mon-table-wrap { flex: 1; overflow-y: auto; }
      #mon-table {
        width: 100%; border-collapse: collapse; font-size: 12px;
        table-layout: fixed;
      }
      #mon-table thead th {
        padding: 8px 12px;
        text-align: left; font-size: 9px; font-weight: 700;
        color: var(--mon-text-faint); letter-spacing: 1.2px;
        text-transform: uppercase;
        background: var(--mon-surface);
        border-bottom: 1px solid var(--mon-border);
        position: sticky; top: 0; z-index: 2;
      }
      #mon-table thead th.center { text-align: center; }

      /* ── ROWS ── */
      tr.op-row {
        border-bottom: 1px solid var(--mon-border);
        cursor: pointer;
        transition: background 0.12s;
        animation: mon-row-in 0.25s ease both;
      }
      tr.op-row:hover td { background: var(--mon-surface); }
      tr.op-row.is-expanded td { background: var(--mon-surface2); }
      tr.op-row td {
        padding: 10px 12px; vertical-align: middle;
        background: transparent; transition: background 0.12s;
      }
      .mon-chave {
        font-family: 'Consolas', monospace;
        font-size: 11px; color: #a0a0c0;
        white-space: nowrap; overflow: hidden; text-overflow: ellipsis;
        letter-spacing: 0.3px;
      }
      .mon-sigla {
        font-weight: 700; font-size: 13px; color: #e0e0f0;
      }
      .mon-site {
        font-size: 11px; color: var(--mon-text-dim);
        white-space: nowrap; overflow: hidden; text-overflow: ellipsis;
      }
      .mon-hora {
        font-family: 'Consolas', monospace;
        font-size: 12px; color: var(--mon-text);
        letter-spacing: 0.5px;
      }
      .mon-lider {
        font-size: 11px; color: var(--mon-text-dim);
        white-space: nowrap; overflow: hidden; text-overflow: ellipsis;
      }
      .mon-chevron { color: var(--mon-text-faint); font-size: 10px; transition: transform 0.2s; }
      tr.op-row.is-expanded .mon-chevron { transform: rotate(180deg); }

      /* ── PROGRESS BADGE ── */
      .mon-prog-cell { text-align: center; }
      .mon-prog-num {
        font-size: 12px; font-weight: 700; letter-spacing: 0.3px;
      }
      .mon-prog-bar {
        height: 3px; background: rgba(255,255,255,0.06); border-radius: 2px;
        overflow: hidden; margin-top: 4px;
      }
      .mon-prog-fill {
        height: 100%; border-radius: 2px;
        animation: mon-bar-fill 0.7s cubic-bezier(0.4,0,0.2,1) both;
        width: var(--bar-w);
      }
      .mon-prog-pending { font-size: 11px; color: var(--mon-text-faint); }
      .mon-prog-na { font-size: 10px; color: var(--mon-text-faint); }

      /* ── STATUS BADGES ── */
      .mon-status-badge {
        display: inline-flex; align-items: center; gap: 4px;
        padding: 3px 8px; border-radius: 20px;
        font-size: 10px; font-weight: 700; letter-spacing: 0.3px;
        white-space: nowrap;
      }
      .mon-status-badge.completo  { color: var(--mon-green);  background: var(--mon-green-bg); }
      .mon-status-badge.parcial   { color: var(--mon-amber);  background: var(--mon-amber-bg); }
      .mon-status-badge.esc-ok    { color: var(--mon-indigo); background: var(--mon-indigo-bg); }
      .mon-status-badge.nenhum    { color: var(--mon-red);    background: var(--mon-red-bg); }
      .mon-status-badge.neutro    { color: var(--mon-text-faint); background: transparent; }
      .mon-envelope { font-size: 12px; margin-left: 4px; opacity: 0.8; }

      /* ── DETALHE EXPANDIDO ── */
      tr.op-detail td { padding: 0 !important; border-bottom: 1px solid var(--mon-border); }
      .mon-detail-inner {
        background: var(--mon-bg);
        border-top: 1px solid var(--mon-border);
        padding: 16px;
        animation: mon-fadein 0.2s ease;
      }
      .mon-detail-header {
        display: flex; align-items: center; gap: 10px; margin-bottom: 14px;
      }
      .mon-key-chip {
        font-family: 'Consolas', monospace; font-size: 11px;
        color: var(--mon-text); background: var(--mon-surface2);
        border: 1px solid var(--mon-border2); border-radius: 4px;
        padding: 3px 8px; letter-spacing: 0.3px;
      }
      .mon-copy-btn {
        background: transparent; border: 1px solid var(--mon-border);
        color: var(--mon-text-faint); padding: 3px 9px;
        border-radius: 4px; font-size: 10px; font-family: inherit;
        cursor: pointer; transition: all 0.15s;
      }
      .mon-copy-btn:hover { background: var(--mon-surface2); color: var(--mon-text); }

      .mon-stat-grid {
        display: grid; grid-template-columns: 1fr 1fr; gap: 10px;
        margin-bottom: 14px;
      }
      .mon-stat-card {
        background: var(--mon-surface);
        border: 1px solid var(--mon-border);
        border-radius: var(--mon-radius-sm);
        padding: 12px 14px;
      }
      .mon-stat-card-header {
        display: flex; justify-content: space-between; align-items: baseline;
        margin-bottom: 8px;
      }
      .mon-stat-card-label {
        font-size: 9px; color: var(--mon-text-dim); letter-spacing: 1px;
        text-transform: uppercase; font-weight: 700;
      }
      .mon-stat-card-num {
        font-size: 18px; font-weight: 700;
      }
      .mon-stat-card-sub {
        font-size: 10px; color: var(--mon-text-faint);
        margin-left: 2px;
      }
      .mon-stat-card-pct { font-size: 9px; color: var(--mon-text-faint); margin-top: 4px; }

      /* ── ACTIONS ── */
      .mon-actions { display: flex; flex-wrap: wrap; gap: 8px; margin-bottom: 14px; align-items: center; }

      .mon-send-btn {
        display: inline-flex; align-items: center; gap: 6px;
        background: var(--mon-indigo-bg);
        border: 1px solid rgba(129,140,248,0.3);
        color: var(--mon-indigo);
        padding: 6px 14px; border-radius: var(--mon-radius-sm);
        font-size: 11px; font-family: inherit; font-weight: 600;
        cursor: pointer; letter-spacing: 0.3px;
        transition: all 0.15s;
      }
      .mon-send-btn:hover { background: rgba(129,140,248,0.18); }
      .mon-send-btn:disabled { cursor: not-allowed; opacity: 0.5; }
      .mon-send-btn--err { color: var(--mon-red) !important; border-color: rgba(248,113,113,0.3) !important; background: var(--mon-red-bg) !important; }
      .mon-send-btn--report { background: rgba(245,158,11,0.1); border-color: rgba(245,158,11,0.3) !important; color: var(--mon-amber) !important; }
      .mon-send-btn--report:hover { background: rgba(245,158,11,0.2); }

      /* ── FILTER BAR ── */
      #mon-filter-bar {
        display: flex; flex-wrap: wrap; align-items: center; gap: 8px;
        padding: 8px 12px; border-bottom: 1px solid var(--mon-border);
        background: var(--mon-surface); flex-shrink: 0;
      }
      #mon-filter-input-wrap {
        display: flex; align-items: center; gap: 6px;
        background: var(--mon-bg); border: 1px solid var(--mon-border2);
        border-radius: var(--mon-radius-sm); padding: 5px 10px;
        flex: 1; min-width: 200px;
      }
      #mon-filter-input-wrap:focus-within { border-color: var(--mon-indigo); }
      .mon-filter-icon { color: var(--mon-text-dim); font-size: 14px; }
      #mon-filter-input {
        flex: 1; background: transparent; border: none; outline: none;
        color: var(--mon-text); font-size: 12px; font-family: inherit;
      }
      #mon-filter-input::placeholder { color: var(--mon-text-faint); }
      #mon-filter-clear {
        background: none; border: none; color: var(--mon-text-faint);
        cursor: pointer; font-size: 11px; padding: 0; line-height: 1;
        display: none;
      }
      #mon-filter-clear:hover { color: var(--mon-red); }
      #mon-status-chips { display: flex; flex-wrap: wrap; gap: 5px; }
      .mon-chip {
        padding: 4px 10px; border-radius: 99px; font-size: 10px; font-weight: 600;
        font-family: inherit; cursor: pointer; border: 1px solid var(--mon-border2);
        background: transparent; color: var(--mon-text-dim);
        transition: all 0.15s; letter-spacing: 0.3px;
      }
      .mon-chip:hover { background: var(--mon-surface2); color: var(--mon-text); }
      .mon-chip.active { background: var(--mon-indigo-bg); border-color: rgba(129,140,248,0.4); color: var(--mon-indigo); }
      .mon-chip--completo.active { background: var(--mon-green-bg); border-color: rgba(52,212,116,0.4); color: var(--mon-green); }
      .mon-chip--parcial.active  { background: var(--mon-amber-bg); border-color: rgba(245,158,11,0.4); color: var(--mon-amber); }
      .mon-chip--esc.active      { background: var(--mon-indigo-bg); border-color: rgba(129,140,248,0.4); color: var(--mon-indigo); }
      .mon-chip--nenhum.active   { background: var(--mon-red-bg); border-color: rgba(248,113,113,0.4); color: var(--mon-red); }

      /* ── SORT HEADERS ── */
      .mon-th-sort { cursor: pointer; user-select: none; white-space: nowrap; }
      .mon-th-sort:hover { color: var(--mon-text); }
      .mon-sort-arrow { font-size: 9px; margin-left: 3px; opacity: 0.4; }
      .mon-th-sort.sort-asc  .mon-sort-arrow { content: '▲'; opacity: 1; color: var(--mon-indigo); }
      .mon-th-sort.sort-desc .mon-sort-arrow { content: '▼'; opacity: 1; color: var(--mon-indigo); }
      .mon-th-sort.sort-asc  .mon-sort-arrow::after { content: '▲'; }
      .mon-th-sort.sort-desc .mon-sort-arrow::after { content: '▼'; }
      .mon-th-sort:not(.sort-asc):not(.sort-desc) .mon-sort-arrow::after { content: '⇅'; }
      #mon-table thead th { white-space: nowrap; }

      .mon-dl-btn {
        display: inline-flex; align-items: center; gap: 5px;
        background: transparent;
        border: 1px solid var(--mon-border2);
        color: var(--mon-text-dim);
        padding: 5px 12px; border-radius: var(--mon-radius-sm);
        font-size: 11px; font-family: inherit;
        cursor: pointer; transition: all 0.15s; text-decoration: none;
      }
      .mon-dl-btn:hover { background: var(--mon-surface2); color: var(--mon-text); }

      .mon-xls-menu { position: relative; display: inline-block; }
      .mon-xls-dropdown {
        display: none; position: absolute; bottom: 100%; left: 0;
        margin-bottom: 6px;
        background: var(--mon-surface); border: 1px solid var(--mon-border2);
        border-radius: var(--mon-radius-sm); padding: 4px;
        z-index: 999; min-width: 190px;
        box-shadow: 0 8px 32px rgba(0,0,0,0.7);
      }
      .mon-xls-menu:hover .mon-xls-dropdown { display: block; }
      .mon-xls-dropdown a {
        display: block; padding: 7px 10px; color: var(--mon-text-dim);
        font-size: 11px; font-family: inherit; text-decoration: none;
        transition: all 0.1s; border-radius: 4px;
      }
      .mon-xls-dropdown a:hover { background: var(--mon-surface2); color: var(--mon-text); }

      /* ── LISTAS ── */
      .mon-list-label {
        font-size: 9px; letter-spacing: 1px; text-transform: uppercase;
        font-weight: 700; margin-bottom: 8px;
      }
      .mon-list-label.warn { color: var(--mon-red); }
      .mon-list-label.muted { color: var(--mon-text-faint); }

      .mon-list-table {
        width: 100%; border-collapse: collapse; font-size: 11px; margin-bottom: 14px;
      }
      .mon-list-table thead th {
        padding: 5px 10px; text-align: left;
        font-size: 9px; color: var(--mon-text-faint);
        letter-spacing: 1px; text-transform: uppercase; font-weight: 700;
        border-bottom: 1px solid var(--mon-border);
        background: transparent; position: static;
      }
      .mon-list-table tbody tr {
        border-bottom: 1px solid var(--mon-border);
        transition: background 0.1s;
      }
      .mon-list-table tbody tr:last-child { border-bottom: none; }
      .mon-list-table tbody tr:hover td { background: var(--mon-surface); }
      .mon-list-table td { padding: 6px 10px; }
      .mon-list-table td.name { font-weight: 600; }
      .mon-list-table td.name.warn { color: var(--mon-red); }
      .mon-list-table td.name.ok   { color: var(--mon-text); }
      .mon-list-table td.tipo { color: var(--mon-text-dim); }
      .mon-list-table td.time { color: var(--mon-text-dim); font-family: 'Consolas', monospace; font-size: 10px; }

      .mon-empty-detail { color: var(--mon-text-faint); font-size: 11px; }

      /* ── LISTAS LADO A LADO ── */
      #mon-panel .mon-lists-grid {
        display: grid !important; grid-template-columns: 1fr 1fr !important; gap: 10px !important;
        margin: 0 !important; padding: 0 !important;
      }
      #mon-panel .mon-list-panel {
        border-radius: var(--mon-radius-sm) !important;
        border: 1px solid var(--mon-border) !important;
        overflow: hidden !important;
        background: var(--mon-bg) !important;
      }
      #mon-panel .mon-list-panel-header {
        display: flex !important; align-items: center !important; gap: 8px !important;
        padding: 9px 12px !important;
        background: var(--mon-surface) !important;
        border-bottom: 1px solid var(--mon-border) !important;
        margin: 0 !important;
      }
      #mon-panel .mon-list-panel-dot {
        width: 6px !important; height: 6px !important; border-radius: 50% !important;
        flex-shrink: 0 !important; display: inline-block !important;
      }
      #mon-panel .mon-list-panel-title {
        font-size: 10px !important; font-weight: 700 !important; letter-spacing: 0.8px !important;
        text-transform: uppercase !important; color: var(--mon-text-dim) !important; flex: 1 !important;
        margin: 0 !important;
      }
      #mon-panel .mon-list-panel-count {
        font-size: 10px !important; font-weight: 700 !important;
        padding: 2px 7px !important; border-radius: 10px !important;
        line-height: 1.4 !important;
      }
      #mon-panel .mon-list-panel-body {
        max-height: 240px !important; overflow-y: auto !important;
      }
      #mon-panel .mon-list-row {
        padding: 8px 12px !important;
        border-bottom: 1px solid var(--mon-border) !important;
        transition: background 0.1s !important;
        display: block !important;
      }
      #mon-panel .mon-list-row:last-child { border-bottom: none !important; }
      #mon-panel .mon-list-row:hover { background: var(--mon-surface) !important; }
      #mon-panel .mon-list-row-name {
        font-size: 11px !important; font-weight: 600 !important; color: var(--mon-text) !important;
        margin-bottom: 3px !important; display: block !important;
        white-space: normal !important;
      }
      #mon-panel .mon-list-row-meta {
        display: flex !important; align-items: center !important; gap: 8px !important;
      }
      #mon-panel .mon-list-row-tipo {
        font-size: 10px !important; color: var(--mon-text-dim) !important;
      }
      #mon-panel .mon-list-row-time {
        font-size: 11px !important; color: #f59e0b !important;
        font-family: 'Consolas', monospace !important; margin-left: auto !important;
        background: rgba(245,158,11,0.12) !important;
        border: 1px solid rgba(245,158,11,0.3) !important;
        border-radius: 4px !important; padding: 2px 8px !important;
        letter-spacing: 0.5px !important; font-weight: 700 !important;
      }
      #mon-panel .mon-list-empty {
        padding: 20px 12px !important; font-size: 11px !important;
        color: var(--mon-text-faint) !important; text-align: center !important;
        display: block !important;
      }
      .mon-loading-detail {
        display: flex; align-items: center; gap: 8px;
        color: var(--mon-text-dim); font-size: 11px;
      }
      .mon-loading-spinner {
        width: 14px; height: 14px; border-radius: 50%;
        border: 2px solid var(--mon-border2);
        border-top-color: var(--mon-indigo);
        animation: mon-spin 0.7s linear infinite; flex-shrink: 0;
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
        border-radius: 50%; background: rgba(0,0,0,0.55); z-index: 2;
        transition: opacity 0.3s;
      }
      #mon-progress-text {
        position: absolute; top: 2px; left: 2px; width: 40px; height: 40px;
        border-radius: 50%; z-index: 3;
        display: flex; align-items: center; justify-content: center;
        font-size: 9px; font-weight: 700; color: #fff; font-family: 'Consolas', monospace;
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
      position:fixed;top:0;right:0;width:1020px;height:100vh;
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
          <div class="mon-logo-icon">M</div>
          <div class="mon-logo-text">
            <div class="mon-logo-title">Monitor TSI</div>
            <div class="mon-logo-sub" id="mon-sub">Inicializando…</div>
          </div>
        </div>
        <div style="display:flex;align-items:center;gap:8px">
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
          </div>
          <div class="mon-metric">
            <div class="mon-metric-label">Completas</div>
            <div class="mon-metric-val green" id="m-ok">—</div>
          </div>
          <div class="mon-metric">
            <div class="mon-metric-label">Parciais</div>
            <div class="mon-metric-val amber" id="m-inc">—</div>
          </div>
          <div class="mon-metric">
            <div class="mon-metric-label">Sem apontamento</div>
            <div class="mon-metric-val red" id="m-zero">—</div>
          </div>
        </div>

        <!-- FILTRO + STATUS CHIPS -->
        <div id="mon-filter-bar">
          <div id="mon-filter-input-wrap">
            <span class="mon-filter-icon">⌕</span>
            <input id="mon-filter-input" type="text" placeholder="Filtrar por chave, sigla ou site…"
              oninput="window._monSetFilter(this.value)" autocomplete="off" spellcheck="false" />
            <button id="mon-filter-clear" onclick="window._monClearFilter()" title="Limpar">✕</button>
          </div>
          <div id="mon-status-chips">
            <button class="mon-chip mon-chip--all active" onclick="window._monSetStatusFilter('all',this)">Todos</button>
            <button class="mon-chip mon-chip--completo" onclick="window._monSetStatusFilter('completo',this)">✓ Completo</button>
            <button class="mon-chip mon-chip--parcial"  onclick="window._monSetStatusFilter('parcial',this)">△ Parcial</button>
            <button class="mon-chip mon-chip--esc"      onclick="window._monSetStatusFilter('esc',this)">Esc. ok</button>
            <button class="mon-chip mon-chip--nenhum"   onclick="window._monSetStatusFilter('nenhum',this)">✗ Nenhum</button>
          </div>
        </div>

        <!-- TABELA -->
        <div id="mon-table-wrap">
          <table id="mon-table">
            <thead>
              <tr>
                <th style="width:13%">Chave</th>
                <th style="width:7%">Sigla</th>
                <th style="width:17%">Site</th>
                <th class="center mon-th-sort" data-col="esc" style="width:11%" onclick="window._monToggleSort('esc',this)">Esc / Sol <span class="mon-sort-arrow"></span></th>
                <th class="center mon-th-sort" data-col="apt" style="width:11%" onclick="window._monToggleSort('apt',this)">Apt / Sol <span class="mon-sort-arrow"></span></th>
                <th class="mon-th-sort" data-col="hora" style="width:7%"  onclick="window._monToggleSort('hora',this)">Hora <span class="mon-sort-arrow"></span></th>
                <th style="width:14%">Líder</th>
                <th class="center mon-th-sort" data-col="status" style="width:16%" onclick="window._monToggleSort('status',this)">Status <span class="mon-sort-arrow"></span></th>
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
      tbody.innerHTML = `<tr><td colspan="9" style="text-align:center;padding:3rem;color:var(--mon-text-faint);font-size:13px">Nenhuma operação encontrada</td></tr>`;
      updateMetrics();
      return;
    }

    const visibleOps = getVisibleOps();

    // Atualiza contador no input placeholder
    const inp = document.getElementById('mon-filter-input');
    if (inp && !filterText) inp.placeholder = `Filtrar por chave, sigla ou site… (${operations.length} ops)`;

    if (visibleOps.length === 0) {
      tbody.innerHTML = `<tr><td colspan="9" style="text-align:center;padding:3rem;color:var(--mon-text-faint);font-size:13px">Nenhuma operação corresponde ao filtro</td></tr>`;
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
        <td><span class="mon-chave" style="color:#a0a0c0" title="${op.chave}">${hl(op.chave)}</span></td>
        <td><span class="mon-sigla">${hl(op.sigla)}</span></td>
        <td><span class="mon-site" title="${op.site}">${hl(op.site)}</span></td>
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
        det.innerHTML = `<td colspan="9"><div class="mon-detail-inner">${renderDetail(op)}</div></td>`;
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
    html += `<div class="mon-actions">
      <button class="mon-send-btn" onclick="event.stopPropagation();window._monEnviarEscala('${op.id}',this)">✓ Escala enviada</button>
      <button class="mon-send-btn mon-send-btn--report" onclick="event.stopPropagation();window._monEnviarReport('${op.id}',this)" title="Preenche P1–P11 e coloca ${qtdApt} apontamentos no P10">📋 Report enviado</button>`;
    pdfLinks.forEach(l => {
      html += `<a href="https://tsi-app.com/${l.href}" target="_blank" class="mon-dl-btn">📄 ${l.label || 'Assinatura'}</a>`;
    });
    if (xlsLinks.length > 0) {
      html += `<div class="mon-xls-menu">
        <button class="mon-dl-btn">📊 XLS ▾</button>
        <div class="mon-xls-dropdown">`;
      xlsLinks.forEach(l => { html += `<a href="https://tsi-app.com/${l.href}" target="_blank">${l.label}</a>`; });
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
          html += `
            <div class="mon-list-row">
              <div class="mon-list-row-name">${c.nome}</div>
              <div class="mon-list-row-meta">
                <span class="mon-list-row-tipo">${c.tipo||'—'}</span>
                <span class="mon-list-row-time">${c.inicio}</span>
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
    if (expanded.has(op.chave)) { expanded.delete(op.chave); renderTable(); return; }
    expanded.add(op.chave);
    renderTable();
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

  function startMonitor() {
    window._monRunning = true;
    fetchOperations();
    refreshTimer  = setInterval(silentRefresh, 60 * 1000);
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

})();
