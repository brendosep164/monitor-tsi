// ==UserScript==
// @name         Monitor Operacional TSI
// @namespace    http://tampermonkey.net/
// @version      11.0
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

  // ── IFRAMES ──────────────────────────────────────────────────────────────────
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
    if (ifr) ifr.onload = null;
    setTimeout(processQueue, 30);
  }

  function startWatchdog() {
    if (watchdogTimer) clearInterval(watchdogTimer);
    watchdogTimer = setInterval(() => {
      const now = Date.now();
      Object.entries(iframesInUse).forEach(([ifrId, info]) => {
        if (now - info.since > 35000) {
          if (info.opId && apontCache[info.opId] === 'loading') {
            apontCache[info.opId] = { solicitado: 0, escalado: 0, apontado: 0, colaboradores: [], escalados: [], faltando: [], pdfLinks: [], xlsLinks: [], _erro: true };
          }
          releaseIfr(ifrId);
        }
      });
    }, 15000);
  }

  // ── JANELA ───────────────────────────────────────────────────────────────────
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

  // ── PARSER ───────────────────────────────────────────────────────────────────
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
        if (btn) { btn.textContent = 'Notif. ativas'; btn.classList.add('is-active'); }
      } else {
        if (btn) { btn.textContent = 'Bloqueado'; btn.classList.add('is-blocked'); }
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
    const bar     = document.getElementById('mon-progress-bar');
    const label   = document.getElementById('mon-progress-label');
    if (!bar || !label) return;
    const pct = total === 0 ? 100 : Math.round((loaded / total) * 100);
    bar.style.width = pct + '%';
    if (pct >= 100) {
      bar.style.background = 'var(--clr-ok)';
      label.textContent = 'Carregado';
      label.style.color = 'var(--clr-ok)';
    } else {
      bar.style.background = 'var(--clr-warn)';
      label.textContent = pct + '%';
      label.style.color = 'var(--clr-warn)';
    }
  }

  // ── FILA ─────────────────────────────────────────────────────────────────────
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
    const timer = setTimeout(() => fire(null), timeout);
    ifr.onload = function() {
      setTimeout(() => {
        try { fire(ifr.contentDocument); } catch(e) { fire(null); }
      }, 1800);
    };
    try { ifr.src = url; } catch(e) { fire(null); }
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
      release({ solicitado: op.qtd, escalado: escalados.length, apontado: 0, colaboradores: [], escalados: escalados || [], faltando: escalados || [], pdfLinks: [], xlsLinks: [] });
    };

    if (!op.id) { releaseIfr(ifrId); return; }
    apontCache[op.id] = 'loading';

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

        if (!naJanela(op)) {
          release({ solicitado: op.qtd, escalado: escalados.length, apontado: 0, colaboradores: [], escalados, faltando: escalados, pdfLinks, xlsLinks, _soEscala: true });
          return;
        }

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
          release({ solicitado: op.qtd, escalado: escalados.length, apontado: colaboradores.length, colaboradores, escalados, faltando, pdfLinks, xlsLinks });
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
    if (cached !== undefined && cached !== 'loading' && cached !== null) { callback(cached, null); return; }
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
      btnEl.innerHTML = 'Erro: ' + msg;
      btnEl.style.opacity = '1';
      setTimeout(() => { btnEl.innerHTML = origTxt; }, 3000);
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
            if (radio) { radio.checked = true; radio.dispatchEvent(new Event('change', { bubbles: true })); radio.dispatchEvent(new Event('click', { bubbles: true })); marcados++; }
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
              btnEl.innerHTML = 'Enviado!';
              btnEl.classList.add('sent');
              setTimeout(() => window.location.reload(), 1500);
            }, 2500);
          }, 700);
        } catch(e) { fail('erro'); clearTimeout(safetyTimer); }
      }, 2000);
    };
  }

  window._monEnviarEscala = enviarEscala;

  // ── REFRESH ───────────────────────────────────────────────────────────────────
  function manualRefresh() {
    if (refreshTimer) clearInterval(refreshTimer);
    iframesInUse = {};
    fetchQueue   = [];
    inQueue      = new Set();
    apontCache   = {};
    BG_IFRAME_IDS.forEach(id => { const ifr = document.getElementById(id); if (ifr) ifr.onload = null; });
    fetchOperations();
    refreshTimer = setInterval(silentRefresh, 60 * 1000);
  }
  window._monRefresh = manualRefresh;

  function silentRefresh() {
    const ops = parseOpsFromDoc(document);
    if (ops.length === 0) return;
    const oldIds = new Set(operations.map(o => o.id));
    const newIds = new Set(ops.map(o => o.id));
    operations.filter(o => !newIds.has(o.id)).forEach(o => { delete apontCache[o.id]; expanded.delete(o.chave); monitoradas.delete(monKey(o)); });
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
          enfileirar(op, (novo, old) => { updateCells(op, novo, old); updateMetrics(); });
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
              if (det) det.querySelector('.det-inner').innerHTML = renderDetail(op);
            }
          });
        } else {
          if (!cached || cached._erro) enfileirar(op, (novo) => { updateCells(op, novo, null); });
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
    if (naJanela(op)) { if (cells[4]) cells[4].innerHTML = apontBadge(d, op.qtd); }
    if (cells[7]) cells[7].innerHTML = situacaoBadge(d);
    if (old && old !== 'loading' && old.apontado < old.solicitado && d.apontado >= d.solicitado && d.apontado > 0) notify(op, d);
  }

  // ── DRAG / RESIZE / MINIMIZE ──────────────────────────────────────────────────
  function initControls(panel) {
    panel.style.position = 'fixed';
    const rh = document.createElement('div');
    rh.style.cssText = 'position:absolute;left:0;top:0;width:5px;height:100%;cursor:ew-resize;z-index:10;';
    rh.addEventListener('mousedown', e => {
      e.preventDefault();
      const startX = e.clientX, startW = panel.offsetWidth;
      document.body.style.userSelect = 'none';
      const mv = e => { panel.style.width = Math.min(Math.max(startW + (startX - e.clientX), 480), window.innerWidth - 80) + 'px'; };
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
    btn.innerHTML = 'Monitor';
    btn.style.cssText = `
      position:fixed;bottom:20px;right:20px;z-index:99999;
      background:#1C2733;color:#8FA8C0;
      border:1px solid #2D3F50;
      padding:9px 18px;border-radius:6px;font-size:11px;
      font-family:'DM Mono',monospace,'Courier New';font-weight:500;cursor:pointer;
      letter-spacing:0.08em;box-shadow:0 4px 16px rgba(0,0,0,0.5);
      transition:all 0.15s;
    `;
    btn.onmouseover = () => { btn.style.background = '#243140'; btn.style.color = '#B8CCDC'; };
    btn.onmouseout  = () => { btn.style.background = '#1C2733'; btn.style.color = '#8FA8C0'; };
    btn.onclick = toggleMonitor;
    document.body.appendChild(btn);
  }

  // ── ESTILOS ───────────────────────────────────────────────────────────────────
  function injectStyles() {
    if (document.getElementById('mon-style')) return;
    const s = document.createElement('style');
    s.id = 'mon-style';
    s.textContent = `
      @import url('https://fonts.googleapis.com/css2?family=DM+Mono:wght@400;500&family=IBM+Plex+Sans:wght@400;500;600&display=swap');

      :root {
        --mon-bg:        #111820;
        --mon-bg2:       #141D26;
        --mon-bg3:       #192330;
        --mon-border:    #243040;
        --mon-border2:   #1D2A38;
        --mon-text:      #C4D4E0;
        --mon-text-dim:  #607585;
        --mon-text-faint:#3A5060;
        --clr-ok:        #3ECF8E;
        --clr-partial:   #F0A050;
        --clr-danger:    #E05C5C;
        --clr-esc:       #5B9CF6;
        --clr-accent:    #4A90D9;
        --font-ui:       'IBM Plex Sans', system-ui, sans-serif;
        --font-mono:     'DM Mono', 'Courier New', monospace;
      }

      #mon-panel {
        font-family: var(--font-ui);
        font-size: 12px;
        background: var(--mon-bg);
        color: var(--mon-text);
      }

      #mon-panel *::-webkit-scrollbar { width: 5px; }
      #mon-panel *::-webkit-scrollbar-track { background: var(--mon-bg); }
      #mon-panel *::-webkit-scrollbar-thumb { background: var(--mon-border); border-radius: 3px; }
      #mon-panel *::-webkit-scrollbar-thumb:hover { background: #2D4050; }

      /* HEADER */
      #mon-header {
        background: var(--mon-bg2);
        border-bottom: 1px solid var(--mon-border);
        padding: 0 16px;
        height: 52px;
        display: flex;
        align-items: center;
        justify-content: space-between;
        flex-shrink: 0;
        user-select: none;
      }

      .mon-title {
        font-family: var(--font-ui);
        font-size: 11px;
        font-weight: 600;
        color: var(--mon-text);
        letter-spacing: 0.12em;
        text-transform: uppercase;
      }

      .mon-live-badge {
        font-family: var(--font-mono);
        font-size: 9px;
        font-weight: 500;
        letter-spacing: 0.1em;
        padding: 2px 8px;
        border-radius: 3px;
        border: 1px solid;
        text-transform: uppercase;
      }

      .mon-live-badge.offline { color: var(--mon-text-faint); border-color: var(--mon-border); }
      .mon-live-badge.live    { color: var(--clr-ok); border-color: rgba(62,207,142,0.25); background: rgba(62,207,142,0.06); }
      .mon-live-badge.sync    { color: var(--clr-partial); border-color: rgba(240,160,80,0.25); background: rgba(240,160,80,0.06); }

      /* HEADER BUTTONS */
      .mon-hbtn {
        font-family: var(--font-ui);
        font-size: 11px;
        padding: 5px 12px;
        border-radius: 5px;
        cursor: pointer;
        transition: all 0.12s;
        border: 1px solid var(--mon-border);
        background: transparent;
        color: var(--mon-text-dim);
        letter-spacing: 0.02em;
      }
      .mon-hbtn:hover { background: var(--mon-bg3); color: var(--mon-text); border-color: #2D4050; }
      .mon-hbtn.primary { color: var(--clr-ok); border-color: rgba(62,207,142,0.3); }
      .mon-hbtn.primary:hover { background: rgba(62,207,142,0.08); }
      .mon-hbtn.is-active { color: var(--clr-ok); border-color: rgba(62,207,142,0.3); }
      .mon-hbtn.is-blocked { color: var(--clr-danger); border-color: rgba(224,92,92,0.3); }

      /* PROGRESS BAR */
      .mon-progress-wrap {
        flex: 1;
        max-width: 180px;
        margin: 0 8px;
      }
      .mon-progress-track {
        height: 3px;
        background: var(--mon-border);
        border-radius: 2px;
        overflow: hidden;
      }
      #mon-progress-bar {
        height: 100%;
        width: 0%;
        background: var(--clr-partial);
        border-radius: 2px;
        transition: width 0.35s ease, background 0.35s;
      }
      #mon-progress-label {
        font-family: var(--font-mono);
        font-size: 9px;
        color: var(--mon-text-faint);
        margin-top: 3px;
        letter-spacing: 0.05em;
      }

      /* MÉTRICAS */
      #mon-metrics {
        display: grid;
        grid-template-columns: repeat(4, 1fr);
        border-bottom: 1px solid var(--mon-border);
        flex-shrink: 0;
      }

      .mon-metric-card {
        padding: 12px 16px;
        border-right: 1px solid var(--mon-border2);
      }
      .mon-metric-card:last-child { border-right: none; }

      .mon-metric-label {
        font-size: 9px;
        font-weight: 600;
        letter-spacing: 0.12em;
        text-transform: uppercase;
        color: var(--mon-text-faint);
        margin-bottom: 4px;
      }

      .mon-metric-value {
        font-family: var(--font-mono);
        font-size: 24px;
        font-weight: 500;
        line-height: 1;
      }

      /* TABELA */
      .mon-table-wrap {
        flex: 1;
        overflow-y: auto;
      }

      #mon-table {
        width: 100%;
        border-collapse: collapse;
        table-layout: fixed;
      }

      #mon-table thead tr {
        background: var(--mon-bg2);
        border-bottom: 1px solid var(--mon-border);
        position: sticky;
        top: 0;
        z-index: 2;
      }

      #mon-table thead th {
        padding: 8px 12px;
        text-align: left;
        font-size: 9px;
        font-weight: 600;
        letter-spacing: 0.12em;
        text-transform: uppercase;
        color: var(--mon-text-faint);
      }
      #mon-table thead th.center { text-align: center; }

      #mon-table tbody tr.op-row {
        border-bottom: 1px solid var(--mon-border2);
        cursor: pointer;
        transition: background 0.1s;
      }
      #mon-table tbody tr.op-row:hover td { background: var(--mon-bg3); }
      #mon-table tbody tr.op-row.expanded td { background: var(--mon-bg3); }

      #mon-table tbody td {
        padding: 9px 12px;
        vertical-align: middle;
      }

      .cell-chave {
        font-family: var(--font-mono);
        font-size: 10px;
        color: var(--mon-text-dim);
        overflow: hidden;
        text-overflow: ellipsis;
        white-space: nowrap;
      }

      .cell-sigla {
        font-family: var(--font-mono);
        font-size: 12px;
        font-weight: 500;
        color: var(--mon-text);
      }

      .cell-site {
        font-size: 11px;
        color: var(--mon-text-dim);
        overflow: hidden;
        text-overflow: ellipsis;
        white-space: nowrap;
      }

      .cell-hora {
        font-family: var(--font-mono);
        font-size: 11px;
        color: var(--mon-text-dim);
      }

      .cell-lider {
        font-size: 11px;
        color: var(--mon-text-dim);
        overflow: hidden;
        text-overflow: ellipsis;
        white-space: nowrap;
      }

      .cell-toggle {
        text-align: center;
        color: var(--mon-text-faint);
        font-size: 10px;
      }

      /* BADGE NÚMERO */
      .mon-num-badge {
        display: flex;
        flex-direction: column;
        align-items: center;
        gap: 3px;
      }
      .mon-num-badge .num {
        font-family: var(--font-mono);
        font-size: 12px;
        font-weight: 500;
        line-height: 1;
      }
      .mon-num-badge .bar-track {
        width: 100%;
        height: 2px;
        background: rgba(255,255,255,0.07);
        border-radius: 1px;
        overflow: hidden;
      }
      .mon-num-badge .bar-fill {
        height: 100%;
        border-radius: 1px;
        transition: width 0.6s ease;
      }

      /* STATUS BADGE */
      .mon-status {
        display: inline-flex;
        align-items: center;
        gap: 4px;
        font-size: 10px;
        font-weight: 600;
        letter-spacing: 0.06em;
        padding: 2px 8px;
        border-radius: 3px;
        border: 1px solid transparent;
        white-space: nowrap;
      }
      .mon-status.ok      { color: var(--clr-ok); border-color: rgba(62,207,142,0.2); background: rgba(62,207,142,0.07); }
      .mon-status.partial { color: var(--clr-partial); border-color: rgba(240,160,80,0.2); background: rgba(240,160,80,0.07); }
      .mon-status.danger  { color: var(--clr-danger); border-color: rgba(224,92,92,0.2); background: rgba(224,92,92,0.07); }
      .mon-status.esc     { color: var(--clr-esc); border-color: rgba(91,156,246,0.2); background: rgba(91,156,246,0.07); }
      .mon-status.empty   { color: var(--mon-text-faint); }

      /* DETALHE */
      tr.det-row { border-bottom: 1px solid var(--mon-border); }
      .det-inner {
        background: var(--mon-bg);
        border-top: 1px solid var(--mon-border2);
        padding: 14px 16px;
      }

      .det-stats-grid {
        display: grid;
        grid-template-columns: 1fr 1fr;
        gap: 10px;
        margin-bottom: 14px;
      }

      .det-stat-card {
        background: var(--mon-bg2);
        border: 1px solid var(--mon-border2);
        border-radius: 6px;
        padding: 12px 14px;
      }

      .det-stat-title {
        font-size: 9px;
        font-weight: 600;
        letter-spacing: 0.12em;
        text-transform: uppercase;
        color: var(--mon-text-faint);
        margin-bottom: 8px;
      }

      .det-stat-row {
        display: flex;
        align-items: baseline;
        justify-content: space-between;
        margin-bottom: 6px;
      }

      .det-stat-num {
        font-family: var(--font-mono);
        font-size: 20px;
        font-weight: 500;
        line-height: 1;
      }

      .det-stat-den {
        font-family: var(--font-mono);
        font-size: 12px;
        color: var(--mon-text-faint);
      }

      .det-stat-pct {
        font-family: var(--font-mono);
        font-size: 11px;
        color: var(--mon-text-dim);
      }

      .det-bar-track {
        height: 4px;
        background: rgba(255,255,255,0.06);
        border-radius: 2px;
        overflow: hidden;
      }
      .det-bar-fill {
        height: 100%;
        border-radius: 2px;
        transition: width 0.7s cubic-bezier(0.4,0,0.2,1);
      }

      /* AÇÕES */
      .det-actions {
        display: flex;
        flex-wrap: wrap;
        gap: 6px;
        margin-bottom: 14px;
      }

      .mon-btn-escala {
        font-family: var(--font-ui);
        font-size: 11px;
        font-weight: 500;
        padding: 6px 14px;
        border-radius: 5px;
        border: 1px solid rgba(91,156,246,0.3);
        background: rgba(91,156,246,0.08);
        color: var(--clr-esc);
        cursor: pointer;
        transition: all 0.15s;
        letter-spacing: 0.02em;
      }
      .mon-btn-escala:hover { background: rgba(91,156,246,0.14); }
      .mon-btn-escala:disabled { opacity: 0.5; cursor: not-allowed; }
      .mon-btn-escala.sent { color: var(--clr-ok); border-color: rgba(62,207,142,0.3); background: rgba(62,207,142,0.08); }

      .mon-btn-link {
        font-family: var(--font-ui);
        font-size: 11px;
        padding: 5px 12px;
        border-radius: 5px;
        border: 1px solid var(--mon-border);
        background: transparent;
        color: var(--mon-text-dim);
        cursor: pointer;
        transition: all 0.12s;
        text-decoration: none;
        display: inline-flex;
        align-items: center;
        gap: 5px;
      }
      .mon-btn-link:hover { background: var(--mon-bg3); color: var(--mon-text); border-color: #2D4050; }

      .mon-xls-menu { position: relative; display: inline-block; }
      .mon-xls-dropdown {
        display: none;
        position: absolute;
        bottom: calc(100% + 4px);
        left: 0;
        background: var(--mon-bg2);
        border: 1px solid var(--mon-border);
        border-radius: 6px;
        padding: 4px 0;
        z-index: 999;
        min-width: 190px;
        box-shadow: 0 8px 24px rgba(0,0,0,0.6);
      }
      .mon-xls-menu:hover .mon-xls-dropdown { display: block; }
      .mon-xls-dropdown a {
        display: block;
        padding: 7px 12px;
        color: var(--mon-text-dim);
        font-size: 11px;
        text-decoration: none;
        transition: all 0.1s;
        border-bottom: 1px solid var(--mon-border2);
        font-family: var(--font-ui);
      }
      .mon-xls-dropdown a:last-child { border-bottom: none; }
      .mon-xls-dropdown a:hover { background: var(--mon-bg3); color: var(--mon-text); }

      /* LISTAS DE COLABORADORES */
      .det-section-title {
        font-size: 9px;
        font-weight: 600;
        letter-spacing: 0.12em;
        text-transform: uppercase;
        margin-bottom: 8px;
        padding-bottom: 6px;
        border-bottom: 1px solid var(--mon-border2);
      }
      .det-section-title.danger { color: var(--clr-danger); }
      .det-section-title.dim    { color: var(--mon-text-faint); }

      .det-colab-table {
        width: 100%;
        border-collapse: collapse;
        font-size: 11px;
        margin-bottom: 14px;
      }
      .det-colab-table th {
        padding: 5px 8px;
        text-align: left;
        font-size: 9px;
        font-weight: 600;
        letter-spacing: 0.1em;
        text-transform: uppercase;
        color: var(--mon-text-faint);
        border-bottom: 1px solid var(--mon-border2);
      }
      .det-colab-table td {
        padding: 6px 8px;
        border-bottom: 1px solid var(--mon-border2);
        vertical-align: middle;
      }
      .det-colab-table tr:last-child td { border-bottom: none; }
      .det-colab-table .name { color: var(--mon-text); font-weight: 500; }
      .det-colab-table .name.danger { color: var(--clr-danger); }
      .det-colab-table .tipo { color: var(--mon-text-dim); }
      .det-colab-table .inicio { font-family: var(--font-mono); color: var(--mon-text-faint); font-size: 10px; }
    `;
    document.head.appendChild(s);
  }

  // ── PAINEL ────────────────────────────────────────────────────────────────────
  function createPanel() {
    if (document.getElementById('mon-panel')) return;
    injectStyles();

    const notifLabel = !('Notification' in window) ? 'Sem suporte' :
      Notification.permission === 'granted' ? 'Notif. ativas' :
      Notification.permission === 'denied'  ? 'Bloqueado' : 'Ativar notif.';
    const notifCls = Notification.permission === 'granted' ? 'mon-hbtn is-active' :
      Notification.permission === 'denied'  ? 'mon-hbtn is-blocked' : 'mon-hbtn';

    const p = document.createElement('div');
    p.id = 'mon-panel';
    p.style.cssText = `
      position:fixed;top:0;right:0;width:1020px;height:100vh;
      background:var(--mon-bg);color:var(--mon-text);z-index:99998;
      box-shadow:-8px 0 40px rgba(0,0,0,0.7);
      display:none;flex-direction:column;overflow:hidden;
      font-family:var(--font-ui);font-size:12px;
      border-left:1px solid var(--mon-border);
    `;

    p.innerHTML = `
      <div id="mon-header">
        <div style="display:flex;align-items:center;gap:10px">
          <span class="mon-title">Monitor Operacional</span>
          <span id="mon-live" class="mon-live-badge offline">Offline</span>
        </div>
        <div style="display:flex;align-items:center;gap:8px">
          <div class="mon-progress-wrap">
            <div class="mon-progress-track"><div id="mon-progress-bar"></div></div>
            <div id="mon-progress-label">—</div>
          </div>
          <span id="mon-sub" style="font-size:10px;color:var(--mon-text-faint);font-family:var(--font-mono);letter-spacing:0.04em"></span>
          <button class="mon-hbtn primary" onclick="window._monRefresh()">↻ Atualizar</button>
          <button id="mon-notif-btn" class="${notifCls}" onclick="window._monPedirNotif()">${notifLabel}</button>
          <button id="mon-min-btn" class="mon-hbtn" onclick="window._monMinimize()" style="padding:5px 10px;font-size:13px">—</button>
          <button class="mon-hbtn" onclick="document.getElementById('mon-panel').style.display='none';document.getElementById('btn-mon').innerHTML='Monitor'" style="padding:5px 10px;font-size:13px">✕</button>
        </div>
      </div>

      <div id="mon-body" style="display:flex;flex-direction:column;flex:1;overflow:hidden">
        <div id="mon-metrics">
          <div class="mon-metric-card">
            <div class="mon-metric-label">Operações</div>
            <div class="mon-metric-value" id="m-total" style="color:var(--mon-text)">—</div>
          </div>
          <div class="mon-metric-card">
            <div class="mon-metric-label">Completas</div>
            <div class="mon-metric-value" id="m-ok" style="color:var(--clr-ok)">—</div>
          </div>
          <div class="mon-metric-card">
            <div class="mon-metric-label">Parciais</div>
            <div class="mon-metric-value" id="m-inc" style="color:var(--clr-partial)">—</div>
          </div>
          <div class="mon-metric-card">
            <div class="mon-metric-label">Sem apontamento</div>
            <div class="mon-metric-value" id="m-zero" style="color:var(--clr-danger)">—</div>
          </div>
        </div>

        <div class="mon-table-wrap">
          <table id="mon-table">
            <thead>
              <tr>
                <th style="width:13%">Chave</th>
                <th style="width:7%">Sigla</th>
                <th style="width:17%">Site</th>
                <th style="width:11%" class="center">Esc / Sol</th>
                <th style="width:11%" class="center">Apt / Sol</th>
                <th style="width:7%">Hora</th>
                <th style="width:14%">Líder</th>
                <th style="width:14%" class="center">Status</th>
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

  function toggleMonitor() {
    if (!document.getElementById('mon-panel')) createPanel();
    const p   = document.getElementById('mon-panel');
    const btn = document.getElementById('btn-mon');
    if (p.style.display === 'none' || !p.style.display) {
      p.style.display = 'flex'; p.style.flexDirection = 'column';
      btn.innerHTML = '■ Monitor';
    } else {
      p.style.display = 'none';
      btn.innerHTML = 'Monitor';
    }
  }

  // ── RENDER ────────────────────────────────────────────────────────────────────
  function renderTable() {
    const tbody = document.getElementById('mon-tbody');
    if (!tbody) return;
    tbody.innerHTML = '';
    if (operations.length === 0) {
      tbody.innerHTML = '<tr><td colspan="9" style="text-align:center;padding:3rem;color:var(--mon-text-faint);font-size:12px">Nenhuma operação encontrada</td></tr>';
      return;
    }
    operations.forEach((op, idx) => {
      const isExp    = expanded.has(op.chave);
      const d        = apontCache[op.id];
      const emJanela = naJanela(op);
      const temDados = d && d !== 'loading';

      const tr = document.createElement('tr');
      tr.className = 'op-row' + (isExp ? ' expanded' : '');
      tr.dataset.chave = op.chave;

      tr.innerHTML = `
        <td class="cell-chave" title="${op.chave}">${op.chave}</td>
        <td class="cell-sigla">${op.sigla}</td>
        <td class="cell-site" title="${op.site}">${op.site}</td>
        <td style="padding:7px 12px;text-align:center">${op.id ? (temDados ? escaladoBadge(d, op.qtd) : `<span style="color:var(--mon-text-faint);font-family:var(--font-mono);font-size:11px">…/${op.qtd}</span>`) : '<span style="color:var(--mon-text-faint)">—</span>'}</td>
        <td style="padding:7px 12px;text-align:center">${emJanela ? (temDados ? apontBadge(d, op.qtd) : `<span style="color:var(--mon-text-faint);font-family:var(--font-mono);font-size:11px">…/${op.qtd}</span>`) : '<span style="color:var(--mon-text-faint)">—</span>'}</td>
        <td class="cell-hora">${op.hora}</td>
        <td class="cell-lider" title="${op.lider}">${op.lider}</td>
        <td style="text-align:center">${situacaoBadge(temDados ? d : null)}</td>
        <td class="cell-toggle">${isExp ? '▴' : '▾'}</td>
      `;
      tr.onclick = () => toggleRow(op, idx);
      tbody.appendChild(tr);

      if (isExp) {
        const det = document.createElement('tr');
        det.id = 'det-' + idx;
        det.className = 'det-row';
        det.innerHTML = `<td colspan="9" style="padding:0"><div class="det-inner">${renderDetail(op)}</div></td>`;
        tbody.appendChild(det);
      }
    });
    updateMetrics();
  }

  function renderDetail(op) {
    const d = apontCache[op.id];
    if (!d || d === 'loading') return '<span style="color:var(--mon-text-faint);font-size:11px">Carregando…</span>';

    const escPct = op.qtd > 0 ? Math.min(100, Math.round((d.escalado / op.qtd) * 100)) : 0;
    const aptPct = op.qtd > 0 ? Math.min(100, Math.round((d.apontado / op.qtd) * 100)) : 0;
    const escCor = escPct >= 100 ? 'var(--clr-ok)' : escPct > 0 ? 'var(--clr-esc)' : 'var(--mon-text-faint)';
    const aptCor = aptPct >= 100 ? 'var(--clr-ok)' : aptPct > 0 ? 'var(--clr-partial)' : 'var(--mon-text-faint)';

    let html = `<div class="det-stats-grid">
      <div class="det-stat-card">
        <div class="det-stat-title">Escalados</div>
        <div class="det-stat-row">
          <span class="det-stat-num" style="color:${escCor}">${d.escalado}<span class="det-stat-den">/${op.qtd}</span></span>
          <span class="det-stat-pct">${escPct}%</span>
        </div>
        <div class="det-bar-track"><div class="det-bar-fill" style="width:${escPct}%;background:${escCor}"></div></div>
      </div>
      <div class="det-stat-card">
        <div class="det-stat-title">Apontados</div>
        <div class="det-stat-row">
          <span class="det-stat-num" style="color:${aptCor}">${d.apontado}<span class="det-stat-den">/${op.qtd}</span></span>
          <span class="det-stat-pct">${aptPct}%</span>
        </div>
        <div class="det-bar-track"><div class="det-bar-fill" style="width:${aptPct}%;background:${aptCor}"></div></div>
      </div>
    </div>`;

    const pdfLinks = d.pdfLinks || [], xlsLinks = d.xlsLinks || [];
    html += `<div class="det-actions">
      <button class="mon-btn-escala" onclick="event.stopPropagation();window._monEnviarEscala('${op.id}',this)">✓ Confirmar escala enviada</button>`;
    pdfLinks.forEach(l => {
      html += `<a href="https://tsi-app.com/${l.href}" target="_blank" class="mon-btn-link">📄 PDF${l.label ? ' · ' + l.label.split(' ')[0] : ''}</a>`;
    });
    if (xlsLinks.length > 0) {
      html += `<div class="mon-xls-menu"><button class="mon-btn-link">📊 XLS ▾</button><div class="mon-xls-dropdown">`;
      xlsLinks.forEach(l => { html += `<a href="https://tsi-app.com/${l.href}" target="_blank">${l.label}</a>`; });
      html += `</div></div>`;
    }
    html += `</div>`;

    const faltando = d.faltando || [];
    if (faltando.length > 0) {
      html += `<div class="det-section-title danger">⚠ Aguardando apontamento — ${faltando.length} colaborador${faltando.length > 1 ? 'es' : ''}</div>
        <table class="det-colab-table"><thead><tr>
          <th>Nome</th><th>Tipo</th>
        </tr></thead><tbody>`;
      faltando.forEach(c => {
        html += `<tr><td class="name danger">${c.nome}</td><td class="tipo">${c.tipo || '—'}</td></tr>`;
      });
      html += '</tbody></table>';
    }

    const colab = d.colaboradores || [];
    if (colab.length > 0) {
      html += `<div class="det-section-title dim">Apontados — ${colab.length}</div>
        <table class="det-colab-table"><thead><tr>
          <th>Nome</th><th>Tipo</th><th>Início</th>
        </tr></thead><tbody>`;
      colab.forEach(c => {
        html += `<tr><td class="name">${c.nome}</td><td class="tipo">${c.tipo || '—'}</td><td class="inicio">${c.inicio}</td></tr>`;
      });
      html += '</tbody></table>';
    }

    if (colab.length === 0 && faltando.length === 0) {
      html += '<div style="color:var(--mon-text-faint);font-size:11px;padding:8px 0">Sem dados de escala ou apontamento.</div>';
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
          if (det) det.querySelector('.det-inner').innerHTML = renderDetail(op);
          updateMetrics();
        }
      }, 500);
      setTimeout(() => clearInterval(poll), 40000);
    }
  }

  // ── BADGES ────────────────────────────────────────────────────────────────────
  function escaladoBadge(d, qtd) {
    if (!d || d === 'loading') return `<span style="color:var(--mon-text-faint);font-family:var(--font-mono);font-size:11px">…/${qtd}</span>`;
    const pct = qtd > 0 ? Math.min(100, Math.round((d.escalado / qtd) * 100)) : 0;
    const cor = pct >= 100 ? 'var(--clr-ok)' : pct > 0 ? 'var(--clr-esc)' : 'var(--mon-text-faint)';
    return `<div class="mon-num-badge"><span class="num" style="color:${cor}">${d.escalado}<span style="color:var(--mon-text-faint);font-size:10px">/${qtd}</span></span><div class="bar-track" style="width:56px"><div class="bar-fill" style="width:${pct}%;background:${cor}"></div></div></div>`;
  }

  function apontBadge(d, qtd) {
    if (!d || d === 'loading') return `<span style="color:var(--mon-text-faint);font-family:var(--font-mono);font-size:11px">…/${qtd}</span>`;
    const pct = qtd > 0 ? Math.min(100, Math.round((d.apontado / qtd) * 100)) : 0;
    const cor = pct >= 100 ? 'var(--clr-ok)' : pct > 0 ? 'var(--clr-partial)' : 'var(--mon-text-faint)';
    return `<div class="mon-num-badge"><span class="num" style="color:${cor}">${d.apontado}<span style="color:var(--mon-text-faint);font-size:10px">/${qtd}</span></span><div class="bar-track" style="width:56px"><div class="bar-fill" style="width:${pct}%;background:${cor}"></div></div></div>`;
  }

  function situacaoBadge(d) {
    if (!d || d === 'loading') return '<span class="mon-status empty">—</span>';
    const escOk = d.escalado >= d.solicitado;
    const aptOk = d.apontado >= d.solicitado;
    if (d._soEscala) {
      if (d.escalado === 0)  return '<span class="mon-status danger">Nenhum</span>';
      if (escOk)             return '<span class="mon-status ok">Esc. OK</span>';
      return `<span class="mon-status esc">Esc. ${d.escalado}/${d.solicitado}</span>`;
    }
    if (aptOk && escOk)      return '<span class="mon-status ok">Completo</span>';
    if (d.apontado === 0 && d.escalado === 0) return '<span class="mon-status danger">Nenhum</span>';
    if (d.apontado === 0)    return `<span class="mon-status esc">Esc. ${d.escalado}/${d.solicitado}</span>`;
    return `<span class="mon-status partial">Parcial ${d.apontado}/${d.solicitado}</span>`;
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

  function setLive(label, cls) {
    const el = document.getElementById('mon-live');
    if (el) { el.textContent = label; el.className = 'mon-live-badge ' + cls; }
  }

  // ── MOTOR ─────────────────────────────────────────────────────────────────────
  function startMonitor() {
    window._monRunning = true;
    fetchOperations();
    refreshTimer  = setInterval(silentRefresh, 60 * 1000);
    startWatchdog();
  }

  function fetchOperations() {
    setLive('Sincronizando', 'sync');
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
      setLive('Online', 'live');
      const sub = document.getElementById('mon-sub');
      if (sub) sub.textContent = 'Atualizado ' + new Date().toLocaleTimeString('pt-BR');

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

  // ── INIT ──────────────────────────────────────────────────────────────────────
  setTimeout(() => {
    ensureIframes();
    injectButton();
    createPanel();
    if (!window._monRunning) startMonitor();
    if (Notification.permission === 'default') pedirPermissaoNotificacao();
  }, 2000);

})();
