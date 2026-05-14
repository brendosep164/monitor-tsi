// ==UserScript==
// @name         Monitor Operacional TSI
// @namespace    http://tampermonkey.net/
// @version      10.1
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
    }
    setTimeout(processQueue, 30);
  }

  function startWatchdog() {
    if (watchdogTimer) clearInterval(watchdogTimer);
    watchdogTimer = setInterval(() => {
      const now = Date.now();
      Object.entries(iframesInUse).forEach(([ifrId, info]) => {
        if (now - info.since > 35000) {
          console.warn('[Monitor] watchdog destravou iframe', ifrId, 'op:', info.opId);
          if (info.opId && apontCache[info.opId] === 'loading') {
            apontCache[info.opId] = {
              solicitado: 0, escalado: 0, apontado: 0,
              colaboradores: [], escalados: [], faltando: [],
              pdfLinks: [], xlsLinks: [], _erro: true
            };
          }
          releaseIfr(ifrId);
        }
      });
    }, 15000);
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
    if (diffMin < -720) diffMin += 1440;
    return diffMin >= -720 && diffMin <= 180;
  }

  function monKey(op) { return op.id || op.chave; }
  function naJanela(op) { return monitoradas.has(monKey(op)) || dentroJanela(op); }

  // ── PARSER DE OPS ────────────────────────────────────────────────────────────
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
        const match = linkEl.getAttribute('onclick').match(/planejamento-operacional-edit([A-Za-z0-9+\/=_-]+)_1/);
        if (match) id = match[1];
      }
      const g = i => cells[i]?.innerText?.trim() || '';
      ops.push({ chave: g(0), sigla: g(1), site: g(2), qtd: parseInt(g(3)) || 0, hora: g(9), lider: g(11), status: g(24).toLowerCase(), time: g(8), id });
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
        if (btn) { btn.textContent = '🔔 ativo'; btn.style.color = '#4ade80'; }
      } else {
        if (btn) { btn.textContent = '🔕 bloqueado'; btn.style.color = '#f87171'; }
      }
    });
  }

  function notify(op, d) {
    if (!('Notification' in window) || Notification.permission !== 'granted') return;
    if ((op.time || '') !== 'VD') return;
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
      circle.style.stroke   = '#4ade80';
      overlay.style.opacity = '0';
      text.style.display    = 'none';
      if (img) { img.style.animation = 'none'; img.style.transform = 'none'; }
    } else {
      circle.style.stroke   = '#fb923c';
      overlay.style.opacity = '0.55';
      text.style.display    = 'flex';
      text.textContent      = pct + '%';
      if (img) img.style.animation = 'mon-shake 0.5s ease-in-out infinite alternate';
    }
  }

  // ── FILA COM DEDUPLICAÇÃO ────────────────────────────────────────────────────
  function getAvailableIframe() {
    return BG_IFRAME_IDS.find(id => !iframesInUse[id]);
  }

  function loadUrl(ifr, url, timeout, onDone) {
    let fired = false;

    const fire = (result) => {
      if (fired) return;
      fired = true;
      clearTimeout(timer);
      ifr.onload = null;
      onDone(result);
    };

    const timer = setTimeout(() => {
      console.warn('[Monitor] timeout', url.split('?')[0].split('/').pop());
      fire(null);
    }, timeout);

    ifr.onload = function() {
      setTimeout(() => {
        try { fire(ifr.contentDocument); }
        catch(e) { fire(null); }
      }, 1800);
    };

    try {
      ifr.src = url;
    } catch(e) {
      fire(null);
    }
  }

  function processQueue() {
    if (fetchQueue.length === 0) return;
    const ifrId = getAvailableIframe();
    if (!ifrId) return;

    const { op, callback } = fetchQueue.shift();
    inQueue.delete(op.id);

    iframesInUse[ifrId] = { opId: op.id, since: Date.now() };
    const ifr = document.getElementById(ifrId);

    const release = (dados) => {
      if (dados !== null) apontCache[op.id] = dados;
      releaseIfr(ifrId);
      callback(apontCache[op.id], null);
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

    if (!op.id) {
      releaseIfr(ifrId);
      return;
    }

    apontCache[op.id] = 'loading';

    // ── Passo 1: modal da operação ──────────────────────────────────────────
    loadUrl(ifr, 'https://tsi-app.com/planejamento-operacional-edit' + op.id + '_1', 14000, (doc) => {
      if (!doc) { fallback([]); return; }

      let escalaHref, eaptHref;
      try {
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
      } catch(e) { fallback([]); return; }

      // ── Passo 2: escala (TODAS as ops, dentro ou fora da janela) ─────────
      loadUrl(ifr, 'https://tsi-app.com/' + escalaHref, 14000, (doc2) => {
        const escalados = [], pdfLinks = [], xlsLinks = [];
        if (doc2) {
          try {
            const tbl = doc2.querySelector('table.tables.table-fixed.card-table.table-bordered');
            if (tbl) {
              tbl.querySelectorAll('tbody tr').forEach(row => {
                if (row.classList.contains('strikethrough')) return;
                if (row.querySelector('td.strikethrough')) return;
                const cells = row.querySelectorAll('td');
                if (cells.length < 5) return;
                const nome = cells[2]?.innerText?.trim();
                const cpf  = cells[3]?.innerText?.trim();
                if (!nome || nome.length < 3) return;
                escalados.push({ nome, cpf, tipo: cells[4]?.innerText?.trim() });
              });
            }
            const xlsLabels = ['Layout 1 (SHEIN)','Layout 2 (Cordovil)','Layout 3 (SBC)','Layout 4 (SBF)','Layout 5 (Endereço)','Layout 6 (KISOC)'];
            doc2.querySelectorAll('a[href*="escalaLiderPDF_"]').forEach(a => pdfLinks.push({ label: a.innerText.trim(), href: a.getAttribute('href') }));
            doc2.querySelectorAll('a[href*="escalaprelistaLiderXLS"]').forEach((a, i) => xlsLinks.push({ label: xlsLabels[i] || a.innerText.trim(), href: a.getAttribute('href') }));
          } catch(e) {}
        }

        // Fora da janela: registra só escala, sem buscar apontamentos
        if (!naJanela(op)) {
          release({
            solicitado: op.qtd, escalado: escalados.length, apontado: 0,
            colaboradores: [], escalados, faltando: escalados,
            pdfLinks, xlsLinks, _soEscala: true
          });
          return;
        }

        // ── Passo 3: apontamentos (só ops dentro da janela de 3h) ─────────
        loadUrl(ifr, 'https://tsi-app.com/' + eaptHref, 14000, (doc3) => {
          const colaboradores = [];
          if (doc3) {
            try {
              const tbl = doc3.querySelector('table.tables.table-fixed.card-table:not(.table-bordered)');
              if (tbl) {
                tbl.querySelectorAll('tbody tr').forEach(row => {
                  const cells = row.querySelectorAll('td');
                  if (cells.length < 10) return;
                  const nome   = cells[0]?.innerText?.trim();
                  const cpf    = cells[1]?.innerText?.trim();
                  const origem = cells[9]?.innerText?.trim();
                  const inicio = cells[8]?.innerText?.trim();
                  if (!nome || nome.length < 3) return;
                  if (origem === 'FALTA') return;
                  if (!inicio) return;
                  colaboradores.push({ nome, cpf, tipo: cells[2]?.innerText?.trim(), inicio });
                });
              }
            } catch(e) {}
          }
          const apontadosCPF = new Set(colaboradores.map(c => c.cpf));
          const faltando = escalados.filter(e => !apontadosCPF.has(e.cpf));
          release({
            solicitado: op.qtd, escalado: escalados.length, apontado: colaboradores.length,
            colaboradores, escalados, faltando, pdfLinks, xlsLinks
          });
        });
      });
    });
  }

  function enfileirar(op, callback) {
    if (!op.id) return;
    if (inQueue.has(op.id)) return;
    if (apontCache[op.id] === 'loading') return;
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
    btnEl.innerHTML = '⏳ enviando...';
    btnEl.style.opacity = '0.6';

    let done = false;
    const fail = (msg) => {
      if (done) return; done = true;
      btnEl.disabled = false;
      btnEl.innerHTML = '✗ ' + msg;
      btnEl.style.color = '#f87171';
      btnEl.style.opacity = '1';
      setTimeout(() => { btnEl.innerHTML = origTxt; btnEl.style.color = ''; }, 3000);
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
              btnEl.innerHTML = '✓ enviado!';
              btnEl.style.color = '#4ade80';
              btnEl.style.opacity = '1';
              setTimeout(() => window.location.reload(), 1500);
            }, 2500);
          }, 700);
        } catch(e) { fail('erro'); clearTimeout(safetyTimer); }
      }, 2000);
    };
  }

  window._monEnviarEscala = enviarEscala;

  // ── ATUALIZAÇÃO MANUAL ────────────────────────────────────────────────────────
  function manualRefresh() {
    if (refreshTimer) clearInterval(refreshTimer);
    iframesInUse = {};
    fetchQueue   = [];
    inQueue      = new Set();
    apontCache   = {};
    BG_IFRAME_IDS.forEach(id => {
      const ifr = document.getElementById(id);
      if (ifr) ifr.onload = null;
    });
    fetchOperations();
    refreshTimer = setInterval(silentRefresh, 60 * 1000);
  }
  window._monRefresh = manualRefresh;

  // ── ATUALIZAÇÃO SILENCIOSA ────────────────────────────────────────────────────
  function silentRefresh() {
    const ops = parseOpsFromDoc(document);
    if (ops.length === 0) return;

    const oldIds  = new Set(operations.map(o => o.id));
    const newIds  = new Set(ops.map(o => o.id));

    operations.filter(o => !newIds.has(o.id)).forEach(o => {
      delete apontCache[o.id];
      expanded.delete(o.chave);
      monitoradas.delete(monKey(o));
    });

    ops.forEach(o => {
      if (dentroJanela(o)) monitoradas.add(monKey(o));
    });
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

        if (emJanela) {
          delete apontCache[op.id];
          enfileirar(op, (novo, old) => {
            updateCells(op, novo, old);
            updateMetrics();
            if (expanded.has(op.chave)) {
              const idx = operations.findIndex(o => o.chave === op.chave);
              const det = document.getElementById('det-' + idx);
              if (det) det.querySelector('div').innerHTML = renderDetail(op);
            }
          });
        } else {
          if (!cached || cached._erro) {
            enfileirar(op, (novo) => {
              updateCells(op, novo, null);
            });
          }
        }
      }, i * 250);
    });

    const sub = document.getElementById('mon-sub');
    if (sub) sub.textContent = 'sync ' + new Date().toLocaleTimeString('pt-BR');
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
    // STATUS: sempre atualiza com badge real
    if (cells[7]) cells[7].innerHTML = situacaoBadge(d);
    if (old && old !== 'loading' && old.apontado < old.solicitado && d.apontado >= d.solicitado && d.apontado > 0) notify(op, d);
  }

  // ── DRAG / RESIZE / MINIMIZE ──────────────────────────────────────────────────
  function initControls(panel) {
    panel.style.position = 'fixed';
    const rh = document.createElement('div');
    rh.style.cssText = 'position:absolute;left:0;top:0;width:5px;height:100%;cursor:ew-resize;z-index:10;background:transparent;';
    rh.addEventListener('mouseover', () => rh.style.background = 'rgba(255,255,255,0.08)');
    rh.addEventListener('mouseout',  () => rh.style.background = 'transparent');
    rh.addEventListener('mousedown', e => {
      e.preventDefault();
      const startX = e.clientX, startW = panel.offsetWidth;
      document.body.style.userSelect = 'none';
      const mv = e => { panel.style.width = Math.min(Math.max(startW + (startX - e.clientX), 400), window.innerWidth - 80) + 'px'; };
      const up = () => { document.body.style.userSelect = ''; document.removeEventListener('mousemove', mv); document.removeEventListener('mouseup', up); };
      document.addEventListener('mousemove', mv);
      document.addEventListener('mouseup', up);
    });
    panel.appendChild(rh);

    const header = panel.querySelector('#mon-header');
    if (header) {
      header.style.cursor = 'grab';
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
        const up = () => { header.style.cursor = 'grab'; document.body.style.userSelect = ''; document.removeEventListener('mousemove', mv); document.removeEventListener('mouseup', up); };
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
    btn.textContent    = minimized ? '□' : '—';
  };

  // ── BOTÃO FLUTUANTE ───────────────────────────────────────────────────────────
  function injectButton() {
    if (document.getElementById('btn-mon')) return;
    if (window.self !== window.top) return;
    const btn = document.createElement('button');
    btn.id = 'btn-mon';
    btn.innerHTML = '▶ MONITOR';
    btn.style.cssText = `
      position:fixed;bottom:16px;right:16px;z-index:99999;
      background:#0f0f0f;color:#e0e0e0;border:1px solid #333;
      padding:8px 16px;border-radius:6px;font-size:12px;
      font-family:monospace;font-weight:700;cursor:pointer;
      letter-spacing:1px;box-shadow:0 2px 12px rgba(0,0,0,0.6);
      transition:all 0.2s;
    `;
    btn.onmouseover = () => btn.style.background = '#1a1a1a';
    btn.onmouseout  = () => btn.style.background = '#0f0f0f';
    btn.onclick = toggleMonitor;
    document.body.appendChild(btn);
  }

  // ── PAINEL ────────────────────────────────────────────────────────────────────
  function createPanel() {
    if (document.getElementById('mon-panel')) return;

    if (!document.getElementById('mon-style')) {
      const s = document.createElement('style');
      s.id = 'mon-style';
      s.textContent = `
        @keyframes mon-shake{0%{transform:rotate(-10deg) scale(1.05)}100%{transform:rotate(10deg) scale(1.05)}}
        @keyframes mon-pulse{0%,100%{opacity:1}50%{opacity:0.3}}
        @keyframes mon-bar-fill{from{width:0}to{width:var(--bar-w)}}
        #mon-panel{font-family:'Segoe UI',monospace,sans-serif}
        #mon-panel ::-webkit-scrollbar{width:4px}
        #mon-panel ::-webkit-scrollbar-track{background:#111}
        #mon-panel ::-webkit-scrollbar-thumb{background:#2a2a2a;border-radius:2px}
        #mon-panel tr.op-row:hover td{background:#141414}
        .mon-bar-wrap{height:3px;background:rgba(255,255,255,0.06);border-radius:2px;overflow:hidden;margin-top:3px}
        .mon-bar-inner{height:100%;border-radius:2px;animation:mon-bar-fill 0.8s cubic-bezier(0.4,0,0.2,1) both;width:var(--bar-w)}
        .mon-send-btn{display:inline-flex;align-items:center;gap:6px;background:rgba(129,140,248,0.08);border:1px solid rgba(129,140,248,0.22);color:#818cf8;padding:5px 12px;border-radius:5px;font-size:10px;font-family:monospace;cursor:pointer;letter-spacing:1px;font-weight:600;transition:all 0.2s}
        .mon-send-btn:hover{background:rgba(129,140,248,0.16)}
        .mon-send-btn:disabled{cursor:not-allowed;opacity:0.6}
        .mon-dl-btn{display:inline-flex;align-items:center;gap:4px;background:rgba(255,255,255,0.04);border:1px solid rgba(255,255,255,0.1);color:#666;padding:4px 10px;border-radius:4px;font-size:10px;font-family:monospace;cursor:pointer;transition:all 0.15s;text-decoration:none}
        .mon-dl-btn:hover{background:rgba(255,255,255,0.08);color:#aaa}
        .mon-xls-menu{position:relative;display:inline-block}
        .mon-xls-dropdown{display:none;position:absolute;bottom:100%;left:0;margin-bottom:4px;background:#111;border:1px solid rgba(255,255,255,0.1);border-radius:6px;padding:4px 0;z-index:999;min-width:180px;box-shadow:0 8px 24px rgba(0,0,0,0.8)}
        .mon-xls-menu:hover .mon-xls-dropdown{display:block}
        .mon-xls-dropdown a{display:block;padding:6px 12px;color:#666;font-size:10px;font-family:monospace;text-decoration:none;transition:all 0.1s;border-bottom:1px solid rgba(255,255,255,0.04)}
        .mon-xls-dropdown a:last-child{border-bottom:none}
        .mon-xls-dropdown a:hover{background:rgba(255,255,255,0.06);color:#aaa}
        .mon-refresh-btn{background:rgba(74,222,128,0.07);border:1px solid rgba(74,222,128,0.2);color:#4ade80;padding:3px 10px;border-radius:4px;font-size:10px;font-family:monospace;cursor:pointer;transition:all 0.15s;letter-spacing:0.5px}
        .mon-refresh-btn:hover{background:rgba(74,222,128,0.14)}
        .mon-refresh-btn:active{transform:scale(0.96)}
      `;
      document.head.appendChild(s);
    }

    const p = document.createElement('div');
    p.id = 'mon-panel';
    p.style.cssText = `
      position:fixed;top:0;right:0;width:980px;height:100vh;
      background:#0d0d0d;color:#ccc;z-index:99998;
      box-shadow:-6px 0 32px rgba(0,0,0,0.7);
      display:none;flex-direction:column;overflow:hidden;
      font-family:'Segoe UI',monospace,sans-serif;font-size:12px;
      border-left:1px solid #1e1e1e;
    `;

    const r = 22, circ = 2 * Math.PI * r;
    const notifLabel = !('Notification' in window) ? 'sem suporte' :
      Notification.permission === 'granted' ? '🔔 ativo' :
      Notification.permission === 'denied'  ? '🔕 bloqueado' : '🔔 ativar';
    const notifColor = Notification.permission === 'granted' ? '#4ade80' :
      Notification.permission === 'denied'  ? '#f87171' : '#666';

    p.innerHTML = `
      <div id="mon-header" style="background:#111;border-bottom:1px solid #1e1e1e;padding:10px 14px;display:flex;align-items:center;justify-content:space-between;flex-shrink:0;user-select:none">
        <div style="display:flex;align-items:center;gap:10px">
          <span style="color:#fff;font-weight:700;font-size:13px;letter-spacing:0.8px">MONITOR OPERACIONAL</span>
          <span id="mon-live" style="font-size:10px;color:#444;padding:2px 7px;border:1px solid #222;border-radius:2px;letter-spacing:0.5px">OFFLINE</span>
        </div>
        <div style="display:flex;align-items:center;gap:7px">
          <button class="mon-refresh-btn" onclick="window._monRefresh()" title="Atualizar tudo agora">↻ atualizar</button>
          <button id="mon-notif-btn" onclick="window._monPedirNotif()" style="background:#111;border:1px solid #222;color:${notifColor};padding:3px 8px;border-radius:4px;font-size:10px;cursor:pointer;font-family:monospace">${notifLabel}</button>
          <div style="position:relative;width:50px;height:50px;flex-shrink:0" title="Progresso">
            <img id="mon-avatar-img" src="${AVATAR_URL}" style="width:46px;height:46px;border-radius:50%;object-fit:cover;position:absolute;top:2px;left:2px;z-index:1;transform-origin:center" />
            <div id="mon-progress-overlay" style="position:absolute;top:2px;left:2px;width:46px;height:46px;border-radius:50%;background:rgba(0,0,0,0.6);z-index:2;transition:opacity 0.3s"></div>
            <div id="mon-progress-text" style="position:absolute;top:2px;left:2px;width:46px;height:46px;border-radius:50%;z-index:3;display:flex;align-items:center;justify-content:center;font-size:10px;font-weight:700;color:#fff;font-family:monospace">0%</div>
            <svg width="50" height="50" style="position:absolute;top:0;left:0;transform:rotate(-90deg);z-index:4">
              <circle cx="25" cy="25" r="${r}" fill="none" stroke="#1e1e1e" stroke-width="3"/>
              <circle id="mon-progress-circle" cx="25" cy="25" r="${r}" fill="none" stroke="#fb923c" stroke-width="3" stroke-dasharray="${circ}" stroke-dashoffset="${circ}" style="transition:stroke-dashoffset 0.4s ease,stroke 0.4s ease"/>
            </svg>
          </div>
          <span id="mon-sub" style="font-size:10px;color:#444">—</span>
          <button id="mon-min-btn" onclick="window._monMinimize()" style="background:#111;border:1px solid #222;color:#555;padding:2px 8px;border-radius:2px;font-size:13px;cursor:pointer;font-family:monospace;line-height:1">—</button>
          <button onclick="document.getElementById('mon-panel').style.display='none';document.getElementById('btn-mon').innerHTML='▶ MONITOR'" style="background:#111;border:1px solid #222;color:#555;padding:2px 8px;border-radius:2px;font-size:12px;cursor:pointer;font-family:monospace">✕</button>
        </div>
      </div>

      <div id="mon-body" style="display:flex;flex-direction:column;flex:1;overflow:hidden">
        <div style="display:grid;grid-template-columns:repeat(4,1fr);gap:1px;background:#1a1a1a;flex-shrink:0;border-bottom:1px solid #1e1e1e">
          <div style="background:#0d0d0d;padding:10px 14px">
            <div style="font-size:9px;color:#444;letter-spacing:1px;margin-bottom:4px;text-transform:uppercase">Operações</div>
            <div style="font-size:22px;font-weight:700;color:#aaa" id="m-total">—</div>
          </div>
          <div style="background:#0d0d0d;padding:10px 14px">
            <div style="font-size:9px;color:#444;letter-spacing:1px;margin-bottom:4px;text-transform:uppercase">Completas</div>
            <div style="font-size:22px;font-weight:700;color:#4ade80" id="m-ok">—</div>
          </div>
          <div style="background:#0d0d0d;padding:10px 14px">
            <div style="font-size:9px;color:#444;letter-spacing:1px;margin-bottom:4px;text-transform:uppercase">Parciais</div>
            <div style="font-size:22px;font-weight:700;color:#fb923c" id="m-inc">—</div>
          </div>
          <div style="background:#0d0d0d;padding:10px 14px">
            <div style="font-size:9px;color:#444;letter-spacing:1px;margin-bottom:4px;text-transform:uppercase">Sem apont.</div>
            <div style="font-size:22px;font-weight:700;color:#f87171" id="m-zero">—</div>
          </div>
        </div>

        <div style="flex:1;overflow-y:auto">
          <table style="width:100%;border-collapse:collapse;font-size:12px;table-layout:fixed">
            <thead>
              <tr style="background:#111;border-bottom:1px solid #1e1e1e;position:sticky;top:0;z-index:1">
                <th style="padding:8px 10px;text-align:left;font-size:9px;color:#444;letter-spacing:1px;font-weight:600;width:14%">CHAVE</th>
                <th style="padding:8px 10px;text-align:left;font-size:9px;color:#444;letter-spacing:1px;font-weight:600;width:7%">SIGLA</th>
                <th style="padding:8px 10px;text-align:left;font-size:9px;color:#444;letter-spacing:1px;font-weight:600;width:16%">SITE</th>
                <th style="padding:8px 10px;text-align:center;font-size:9px;color:#444;letter-spacing:1px;font-weight:600;width:12%">ESC/SOL</th>
                <th style="padding:8px 10px;text-align:center;font-size:9px;color:#444;letter-spacing:1px;font-weight:600;width:12%">APT/SOL</th>
                <th style="padding:8px 10px;text-align:left;font-size:9px;color:#444;letter-spacing:1px;font-weight:600;width:7%">HORA</th>
                <th style="padding:8px 10px;text-align:left;font-size:9px;color:#444;letter-spacing:1px;font-weight:600;width:12%">LÍDER</th>
                <th style="padding:8px 10px;text-align:center;font-size:9px;color:#444;letter-spacing:1px;font-weight:600;width:16%">STATUS</th>
                <th style="padding:8px 10px;width:4%"></th>
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

  function toggleMonitor() {
    if (!document.getElementById('mon-panel')) createPanel();
    const p   = document.getElementById('mon-panel');
    const btn = document.getElementById('btn-mon');
    if (p.style.display === 'none' || !p.style.display) {
      p.style.display = 'flex'; p.style.flexDirection = 'column';
      btn.innerHTML = '■ MONITOR';
    } else {
      p.style.display = 'none';
      btn.innerHTML = '▶ MONITOR';
    }
  }

  function startMonitor() {
    window._monRunning = true;
    fetchOperations();
    refreshTimer  = setInterval(silentRefresh, 60 * 1000);
    startWatchdog();
  }

  function fetchOperations() {
    setLive('SYNC', '#fb923c');
    const ops1 = parseOpsFromDoc(document);
    const ifr2 = document.getElementById(IFR_PAG2);

    const finalizar = (opsAll) => {
      const seen = new Set();
      const ops  = opsAll.filter(o => { if (seen.has(o.chave)) return false; seen.add(o.chave); return true; });

      operations  = ops;
      apontCache  = {};
      expanded    = new Set();
      monitoradas = new Set();
      fetchQueue  = [];
      inQueue     = new Set();

      const opsComId = ops.filter(o => o.id);
      opsComId.forEach(o => { if (dentroJanela(o)) monitoradas.add(monKey(o)); });
      renderTable();
      setLive('LIVE', '#4ade80');
      const sub = document.getElementById('mon-sub');
      if (sub) sub.textContent = 'sync ' + new Date().toLocaleTimeString('pt-BR');

      let loaded = 0;
      const total = opsComId.length;
      updateProgress(0, total);

      opsComId.forEach((op) => {
        enfileirar(op, (novo) => {
          if (dentroJanela(op)) monitoradas.add(monKey(op));
          loaded++;
          updateProgress(loaded, total);
          updateCells(op, novo, null);
          updateMetrics();
        });
      });
    };

    if (!ifr2) { finalizar(ops1); return; }

    let pag2done = false;
    const pag2timer = setTimeout(() => {
      if (!pag2done) { pag2done = true; finalizar(ops1); }
    }, 8000);

    ifr2.onload = null;
    ifr2.src = 'https://tsi-app.com/planejamento-operacional_2';
    ifr2.onload = function() {
      if (pag2done) return;
      pag2done = true;
      clearTimeout(pag2timer);
      setTimeout(() => {
        try { finalizar([...ops1, ...parseOpsFromDoc(ifr2.contentDocument)]); }
        catch(e) { finalizar(ops1); }
      }, 1500);
    };
  }

  // ── RENDER ────────────────────────────────────────────────────────────────────
  function renderTable() {
    const tbody = document.getElementById('mon-tbody');
    if (!tbody) return;
    tbody.innerHTML = '';
    if (operations.length === 0) {
      tbody.innerHTML = '<tr><td colspan="9" style="text-align:center;padding:3rem;color:#333">NENHUMA OPERAÇÃO</td></tr>';
      return;
    }
    operations.forEach((op, idx) => {
      const isExp    = expanded.has(op.chave);
      const d        = apontCache[op.id];
      const emJanela = naJanela(op);
      const temDados = d && d !== 'loading';

      const escPct = temDados && op.qtd > 0 ? Math.min(100, Math.round((d.escalado / op.qtd) * 100)) : 0;
      const aptPct = temDados && emJanela && op.qtd > 0 ? Math.min(100, Math.round((d.apontado / op.qtd) * 100)) : 0;
      const escCor = escPct >= 100 ? '#4ade80' : escPct > 0 ? '#818cf8' : '#2a2a2a';
      const aptCor = aptPct >= 100 ? '#4ade80' : aptPct > 0 ? '#fb923c' : '#2a2a2a';

      const tr = document.createElement('tr');
      tr.className = 'op-row';
      tr.dataset.chave = op.chave;
      tr.style.cssText = `border-bottom:1px solid #161616;cursor:pointer;transition:background 0.1s;${isExp ? 'background:#141414;' : ''}`;

      tr.innerHTML = `
        <td style="padding:8px 10px;color:#777;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;font-size:10px" title="${op.chave}">${op.chave}</td>
        <td style="padding:8px 10px;color:#ddd;font-weight:700;font-size:12px">${op.sigla}</td>
        <td style="padding:8px 10px;color:#666;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;font-size:11px" title="${op.site}">${op.site}</td>
        <td style="padding:6px 10px;text-align:center">
          ${op.id
            ? (temDados
                ? `<div style="font-size:12px;font-weight:700;color:${escCor}">${d.escalado}/${op.qtd}</div><div class="mon-bar-wrap"><div class="mon-bar-inner" style="--bar-w:${escPct}%;background:${escCor}"></div></div>`
                : `<span style="color:#333;font-size:11px">.../${op.qtd}</span>`)
            : '<span style="color:#2a2a2a">—</span>'}
        </td>
        <td style="padding:6px 10px;text-align:center">
          ${emJanela
            ? (temDados
                ? `<div style="font-size:12px;font-weight:700;color:${aptCor}">${d.apontado}/${op.qtd}</div><div class="mon-bar-wrap"><div class="mon-bar-inner" style="--bar-w:${aptPct}%;background:${aptCor}"></div></div>`
                : `<span style="color:#333;font-size:11px">.../${op.qtd}</span>`)
            : '<span style="color:#2a2a2a;font-size:10px">—</span>'}
        </td>
        <td style="padding:8px 10px;color:#666;font-size:11px">${op.hora}</td>
        <td style="padding:8px 10px;color:#888;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;font-size:11px">${op.lider}</td>
        <td style="padding:8px 10px;text-align:center">
          ${situacaoBadge(temDados ? d : null)}
        </td>
        <td style="padding:8px 10px;text-align:center;color:#333;font-size:13px">${isExp ? '▴' : '▾'}</td>
      `;
      tr.onclick = () => toggleRow(op, idx);
      tbody.appendChild(tr);

      if (isExp) {
        const det = document.createElement('tr');
        det.id = 'det-' + idx;
        det.style.cssText = 'border-bottom:1px solid #1e1e1e;';
        det.innerHTML = `<td colspan="9" style="padding:0"><div style="background:#0a0a0a;border-top:1px solid #1a1a1a;padding:10px 14px">${renderDetail(op)}</div></td>`;
        tbody.appendChild(det);
      }
    });
    updateMetrics();
  }

  function renderDetail(op) {
    const d = apontCache[op.id];
    if (!d || d === 'loading') return '<span style="color:#444;font-size:11px">⏳ carregando...</span>';

    const escPct = op.qtd > 0 ? Math.min(100, Math.round((d.escalado / op.qtd) * 100)) : 0;
    const aptPct = op.qtd > 0 ? Math.min(100, Math.round((d.apontado / op.qtd) * 100)) : 0;
    const escCor = escPct >= 100 ? '#4ade80' : escPct > 0 ? '#818cf8' : '#333';
    const aptCor = aptPct >= 100 ? '#4ade80' : aptPct > 0 ? '#fb923c' : '#333';

    let html = `
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:12px;min-width:0">
        <div style="background:rgba(255,255,255,0.03);border:1px solid #1a1a1a;border-radius:6px;padding:10px 12px;min-width:0;overflow:hidden">
          <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:6px">
            <span style="font-size:8px;color:#333;letter-spacing:1.5px;text-transform:uppercase">Escalados</span>
            <span style="font-size:14px;font-weight:700;color:${escCor}">${d.escalado}<span style="color:#2a2a2a;font-size:10px">/${op.qtd}</span></span>
          </div>
          <div style="height:4px;background:rgba(255,255,255,0.06);border-radius:2px;overflow:hidden">
            <div style="height:100%;width:${escPct}%;background:${escCor};border-radius:2px;transition:width 0.8s ease"></div>
          </div>
          <div style="font-size:9px;color:#2a2a2a;margin-top:4px">${escPct}% escalado</div>
        </div>
        <div style="background:rgba(255,255,255,0.03);border:1px solid #1a1a1a;border-radius:6px;padding:10px 12px;min-width:0;overflow:hidden">
          <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:6px">
            <span style="font-size:8px;color:#333;letter-spacing:1.5px;text-transform:uppercase">Apontados</span>
            <span style="font-size:14px;font-weight:700;color:${aptCor}">${d.apontado}<span style="color:#2a2a2a;font-size:10px">/${op.qtd}</span></span>
          </div>
          <div style="height:4px;background:rgba(255,255,255,0.06);border-radius:2px;overflow:hidden">
            <div style="height:100%;width:${aptPct}%;background:${aptCor};border-radius:2px;transition:width 0.8s ease"></div>
          </div>
          <div style="font-size:9px;color:#2a2a2a;margin-top:4px">${aptPct}% apontado</div>
        </div>
      </div>`;

    const pdfLinks = d.pdfLinks || [], xlsLinks = d.xlsLinks || [];
    html += `<div style="display:flex;flex-wrap:wrap;gap:6px;margin-bottom:12px;align-items:center">
      <button class="mon-send-btn" onclick="event.stopPropagation();window._monEnviarEscala('${op.id}',this)">✓ ESCALA ENVIADA</button>`;
    pdfLinks.forEach(l => {
      html += `<a href="https://tsi-app.com/${l.href}" target="_blank" class="mon-dl-btn">📄 PDF${l.label ? ' · '+l.label.split(' ')[0] : ''}</a>`;
    });
    if (xlsLinks.length > 0) {
      html += `<div class="mon-xls-menu"><button class="mon-dl-btn">📊 XLS ▾</button><div class="mon-xls-dropdown">`;
      xlsLinks.forEach(l => { html += `<a href="https://tsi-app.com/${l.href}" target="_blank">${l.label}</a>`; });
      html += `</div></div>`;
    }
    html += `</div>`;

    const faltando = d.faltando || [];
    if (faltando.length > 0) {
      html += `<div style="font-size:9px;color:#f87171;letter-spacing:1px;text-transform:uppercase;margin-bottom:6px">⚠ Faltando apontamento (${faltando.length})</div>
        <table style="width:100%;border-collapse:collapse;font-size:11px;margin-bottom:12px">
          <thead><tr>
            <th style="padding:4px 8px;text-align:left;color:#333;font-size:9px;border-bottom:1px solid #1a1a1a">NOME</th>
            <th style="padding:4px 8px;text-align:left;color:#333;font-size:9px;border-bottom:1px solid #1a1a1a">TIPO</th>
          </tr></thead><tbody>`;
      faltando.forEach(c => {
        html += `<tr style="border-bottom:1px solid #141414"><td style="padding:5px 8px;color:#f87171;font-weight:600">${c.nome}</td><td style="padding:5px 8px;color:#555">${c.tipo||'—'}</td></tr>`;
      });
      html += '</tbody></table>';
    }

    const colab = d.colaboradores || [];
    if (colab.length > 0) {
      html += `<div style="font-size:9px;color:#444;letter-spacing:1px;text-transform:uppercase;margin-bottom:6px">Apontados (${colab.length})</div>
        <table style="width:100%;border-collapse:collapse;font-size:11px">
          <thead><tr>
            <th style="padding:4px 8px;text-align:left;color:#333;font-size:9px;border-bottom:1px solid #1a1a1a">NOME</th>
            <th style="padding:4px 8px;text-align:left;color:#333;font-size:9px;border-bottom:1px solid #1a1a1a">TIPO</th>
            <th style="padding:4px 8px;text-align:left;color:#333;font-size:9px;border-bottom:1px solid #1a1a1a">INÍCIO</th>
          </tr></thead><tbody>`;
      colab.forEach(c => {
        html += `<tr style="border-bottom:1px solid #141414"><td style="padding:5px 8px;color:#bbb;font-weight:600">${c.nome}</td><td style="padding:5px 8px;color:#555">${c.tipo||'—'}</td><td style="padding:5px 8px;color:#666;font-family:monospace">${c.inicio}</td></tr>`;
      });
      html += '</tbody></table>';
    }

    if (colab.length === 0 && faltando.length === 0) {
      html += '<div style="color:#333;font-size:11px">sem dados de escala/apontamento</div>';
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
          if (det) det.querySelector('div').innerHTML = renderDetail(op);
          updateMetrics();
        }
      }, 500);
      setTimeout(() => clearInterval(poll), 40000);
    }
  }

  function escaladoBadge(d, qtd) {
    if (!d || d === 'loading') return `<span style="color:#333;font-size:11px">.../${qtd}</span>`;
    const pct = qtd > 0 ? Math.min(100, Math.round((d.escalado / qtd) * 100)) : 0;
    const cor = pct >= 100 ? '#4ade80' : pct > 0 ? '#818cf8' : '#2a2a2a';
    return `<div style="font-size:12px;font-weight:700;color:${cor}">${d.escalado}/${qtd}</div><div class="mon-bar-wrap"><div class="mon-bar-inner" style="--bar-w:${pct}%;background:${cor}"></div></div>`;
  }

  function apontBadge(d, qtd) {
    if (!d || d === 'loading') return `<span style="color:#333;font-size:11px">.../${qtd}</span>`;
    const pct = qtd > 0 ? Math.min(100, Math.round((d.apontado / qtd) * 100)) : 0;
    const cor = pct >= 100 ? '#4ade80' : pct > 0 ? '#fb923c' : '#2a2a2a';
    return `<div style="font-size:12px;font-weight:700;color:${cor}">${d.apontado}/${qtd}</div><div class="mon-bar-wrap"><div class="mon-bar-inner" style="--bar-w:${pct}%;background:${cor}"></div></div>`;
  }

  function situacaoBadge(d) {
    if (!d || d === 'loading') return '<span style="color:#2a2a2a;font-size:10px">—</span>';
    const escOk = d.escalado >= d.solicitado;
    const aptOk = d.apontado >= d.solicitado;
    // Ops fora da janela: só tem dados de escala (_soEscala), sem apontamento
    if (d._soEscala) {
      if (d.escalado === 0) return '<span style="color:#f87171;font-size:10px;font-weight:700;letter-spacing:0.5px">✗ NENHUM</span>';
      if (escOk)            return '<span style="color:#4ade80;font-size:10px;font-weight:700;letter-spacing:0.5px">✓ ESC OK</span>';
      return `<span style="color:#818cf8;font-size:10px;font-weight:700;letter-spacing:0.5px">ESC ${d.escalado}/${d.solicitado}</span>`;
    }
    // Ops dentro da janela: mostra status completo com apontamentos
    if (aptOk && escOk)   return '<span style="color:#4ade80;font-size:10px;font-weight:700;letter-spacing:0.5px">✓ COMPLETO</span>';
    if (d.apontado === 0 && d.escalado === 0) return '<span style="color:#f87171;font-size:10px;font-weight:700;letter-spacing:0.5px">✗ NENHUM</span>';
    if (d.apontado === 0) return `<span style="color:#f87171;font-size:10px;font-weight:700;letter-spacing:0.5px">ESC ${d.escalado}/${d.solicitado}</span>`;
    return `<span style="color:#fb923c;font-size:10px;font-weight:700;letter-spacing:0.5px">△ APT ${d.apontado}/${d.solicitado}</span>`;
  }

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

  function setLive(label, cor) {
    const el = document.getElementById('mon-live');
    if (el) { el.textContent = label; el.style.color = cor; el.style.borderColor = cor + '33'; }
  }

  // ── INIT ──────────────────────────────────────────────────────────────────────
  setTimeout(() => {
    ensureIframes();
    injectButton();
    createPanel();
    if (!window._monRunning) startMonitor();
    if (Notification.permission === 'default') pedirPermissaoNotificacao();
  }, 2000);

})();