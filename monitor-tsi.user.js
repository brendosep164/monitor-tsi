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
 
  const QUEUE_IFRS = ['_mon_A','_mon_B','_mon_C','_mon_D'];
  const IFR_PAG2   = '_mon_pag2';
  const IFR_ESCALA = '_mon_escala';
  const AVATAR_URL = 'https://i.imgur.com/9HPkbTi.png';
 
  let iframesInUse = new Set();
  let fetchQueue   = [];
  let inQueue      = new Set();
  let operations   = [];
  let apontCache   = {};
  let expanded     = new Set();
  let monitoradas  = new Set();
  let minimized    = false;
  let refreshTimer = null;
 
  // ── CACHE PERSISTENTE ─────────────────────────────────
  window.addEventListener('beforeunload', () => {
    try {
      const s = {};
      Object.entries(apontCache).forEach(([k,v]) => { if (v && v !== 'loading') s[k] = v; });
      sessionStorage.setItem('_mc12', JSON.stringify(s));
    } catch(e) {}
  });
  function restoreCache() {
    try {
      const r = sessionStorage.getItem('_mc12');
      if (r) { apontCache = JSON.parse(r); sessionStorage.removeItem('_mc12'); }
    } catch(e) {}
  }
 
  // ── IFRAMES (sem sandbox) ─────────────────────────────
  function criarIfr(id) {
    if (document.getElementById(id)) return;
    const f = document.createElement('iframe');
    f.id = id;
    f.style.cssText = 'position:fixed;width:1px;height:1px;opacity:0;pointer-events:none;border:none;top:-9999px;left:-9999px;';
    document.body.appendChild(f);
  }
  function ensureIframes() { QUEUE_IFRS.forEach(criarIfr); criarIfr(IFR_PAG2); criarIfr(IFR_ESCALA); }
 
  // ── JANELA ────────────────────────────────────────────
  function dentroJanela(op) {
    if (!op.hora) return false;
    const [h,m] = op.hora.split(':').map(Number);
    if (isNaN(h)) return false;
    const now = new Date();
    const nowMin = now.getHours()*60+now.getMinutes();
    const opMin  = h*60+(m||0);
    let d = opMin - nowMin;
    if (d < -720) d += 1440;
    return d >= -720 && d <= 180;
  }
  function monKey(op) { return op.id || op.chave; }
  function naJanela(op) { return monitoradas.has(monKey(op)) || dentroJanela(op); }
 
  // ── PARSER ────────────────────────────────────────────
  function parseOps(doc) {
    const ops = [];
    const tbody = doc.querySelector('table tbody');
    if (!tbody) return ops;
    tbody.querySelectorAll('tr').forEach(row => {
      const cells = row.querySelectorAll('td');
      if (cells.length < 9) return;
      const lnk = row.querySelector('a[onclick*="planejamento-operacional-edit"]');
      let id = '';
      if (lnk) { const m = lnk.getAttribute('onclick').match(/edit([A-Za-z0-9+\/=_-]+)_1/); if (m) id = m[1]; }
      const g = i => cells[i]?.innerText?.trim() || '';
      ops.push({ chave:g(0), sigla:g(1), site:g(2), qtd:parseInt(g(3))||0, hora:g(9), lider:g(11), status:g(24).toLowerCase(), time:g(8), id });
    });
    return ops;
  }
 
  // ── NOTIFICAÇÕES ──────────────────────────────────────
  function pedirPermissao() {
    if (!('Notification' in window)) return;
    if (Notification.permission === 'granted') { new Notification('Monitor TSI',{body:'Notificações ativas!',icon:AVATAR_URL}); return; }
    if (Notification.permission === 'denied')  { alert('Notificações bloqueadas. Clique no cadeado e permita.'); return; }
    Notification.requestPermission().then(p => {
      const b = document.getElementById('mon-notif-btn');
      if (p === 'granted') { if (b) { b.textContent='🔔 Ativo'; b.className='mon-hbtn active'; } }
      else                  { if (b) { b.textContent='🔕 Bloqueado'; b.className='mon-hbtn blocked'; } }
    });
  }
  function notify(op, d) {
    if (!('Notification' in window) || Notification.permission !== 'granted') return;
    if ((op.time||'') !== 'VD') return;
    try { new Notification('Completo — TSI',{body:`${op.sigla} · ${op.site}\n${d.apontado}/${d.solicitado} apontados`,icon:AVATAR_URL}); } catch(e) {}
  }
  window._monPedirNotif = pedirPermissao;
 
  // ── LOAD URL (closure seguro) ─────────────────────────
  function loadUrl(ifr, url, timeout, cb) {
    let fired = false;
    const fire = res => { if (fired) return; fired=true; clearTimeout(t); ifr.onload=null; cb(res); };
    const t = setTimeout(() => fire(null), timeout);
    ifr.onload = function() { setTimeout(() => { try { fire(ifr.contentDocument); } catch(e) { fire(null); } }, 1800); };
    try { ifr.src = url; } catch(e) { fire(null); }
  }
 
  // ── FILA ──────────────────────────────────────────────
  function getIfr() { return QUEUE_IFRS.find(id => !iframesInUse.has(id)); }
 
  function processQueue() {
    if (!fetchQueue.length) return;
    const ifrId = getIfr();
    if (!ifrId) return;
    const {op, cb} = fetchQueue.shift();
    inQueue.delete(op.id);
    iframesInUse.add(ifrId);
    const ifr = document.getElementById(ifrId);
 
    const release = dados => {
      if (dados) apontCache[op.id] = dados;
      iframesInUse.delete(ifrId);
      ifr.onload = null;
      if (cb) cb(apontCache[op.id]);
      setTimeout(processQueue, 30);
    };
 
    const fallback = (esc=[]) => release({
      solicitado:op.qtd, escalado:esc.length, apontado:0,
      colaboradores:[], escalados:esc, faltando:esc,
      listaLinks:[], xlsLinks:[], _soEscala:true
    });
 
    if (!op.id) { iframesInUse.delete(ifrId); setTimeout(processQueue,30); return; }
    apontCache[op.id] = 'loading';
 
    // Passo 1: modal
    loadUrl(ifr, 'https://tsi-app.com/planejamento-operacional-edit'+op.id+'_1', 14000, doc => {
      if (!doc) { fallback(); return; }
      let escHref, aptHref;
      try {
        const eL = doc.querySelector('a[href*="pedidoEescala"]');
        const aL = doc.querySelector('a[href*="pedidoEapt"]');
        if (eL && aL) { escHref=eL.getAttribute('href'); aptHref=aL.getAttribute('href'); }
        else {
          const eg = doc.querySelector('a[href*="pedidoEgeral"]');
          if (!eg) { fallback(); return; }
          escHref = eg.getAttribute('href').replace('pedidoEgeral','pedidoEescala');
          aptHref = eg.getAttribute('href').replace('pedidoEgeral','pedidoEapt');
        }
      } catch(e) { fallback(); return; }
 
      // Passo 2: escala
      loadUrl(ifr, 'https://tsi-app.com/'+escHref, 14000, doc2 => {
        const escalados=[], listaLinks=[], xlsLinks=[];
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
                escalados.push({nome, cpf, tipo:cells[4]?.innerText?.trim()});
              });
            }
            // Lista para assinatura por lider
            doc2.querySelectorAll('a[href*="escalaprelistaLiderPDF_"]').forEach(a => {
              listaLinks.push({label:a.innerText.trim(), href:a.getAttribute('href')});
            });
            // XLS
            const xlsLabels = ['Layout 1 (SHEIN)','Layout 2 (Cordovil)','Layout 3 (SBC)','Layout 4 (SBF)','Layout 5 (Endereco)','Layout 6 (KISOC)'];
            doc2.querySelectorAll('a[href*="escalaprelistaLiderXLS"]').forEach((a,i) => {
              xlsLinks.push({label:xlsLabels[i]||a.innerText.trim(), href:a.getAttribute('href')});
            });
          } catch(e) {}
        }
 
        // Fora da janela: so escala
        if (!naJanela(op)) {
          release({solicitado:op.qtd, escalado:escalados.length, apontado:0, colaboradores:[], escalados, faltando:escalados, listaLinks, xlsLinks, _soEscala:true});
          return;
        }
 
        // Passo 3: apontamentos
        loadUrl(ifr, 'https://tsi-app.com/'+aptHref, 14000, doc3 => {
          const colabs=[];
          if (doc3) {
            try {
              const tbl = doc3.querySelector('table.tables.table-fixed.card-table:not(.table-bordered)');
              if (tbl) {
                tbl.querySelectorAll('tbody tr').forEach(row => {
                  const c = row.querySelectorAll('td');
                  if (c.length < 10) return;
                  const nome   = c[0]?.innerText?.trim();
                  const cpf    = c[1]?.innerText?.trim();
                  const origem = c[9]?.innerText?.trim();
                  const inicio = c[8]?.innerText?.trim();
                  if (!nome || nome.length < 3) return;
                  if (origem === 'FALTA') return;
                  if (!inicio) return;
                  colabs.push({nome, cpf, tipo:c[2]?.innerText?.trim(), inicio});
                });
              }
            } catch(e) {}
          }
          const cpfs = new Set(colabs.map(c=>c.cpf));
          const faltando = escalados.filter(e=>!cpfs.has(e.cpf));
          release({solicitado:op.qtd, escalado:escalados.length, apontado:colabs.length, colaboradores:colabs, escalados, faltando, listaLinks, xlsLinks});
        });
      });
    });
  }
 
  function enfileirar(op, cb) {
    if (!op.id) return;
    if (inQueue.has(op.id)) return;
    if (apontCache[op.id] === 'loading') return;
    inQueue.add(op.id);
    fetchQueue.push({op, cb});
    processQueue();
  }
 
  // ── ENVIAR ESCALA ─────────────────────────────────────
  function enviarEscala(opId, btn) {
    const ifr = document.getElementById(IFR_ESCALA);
    if (!ifr || !opId) return;
    const orig = btn.innerHTML;
    btn.disabled = true; btn.innerHTML = 'Enviando…';
    let done = false;
    const fail = msg => {
      if (done) return; done=true;
      btn.disabled=false; btn.innerHTML='Erro: '+msg;
      setTimeout(()=>{btn.innerHTML=orig;},3000);
    };
    const safe = setTimeout(()=>fail('timeout'),25000);
    ifr.onload = null;
    ifr.src = 'https://tsi-app.com/planejamento-operacional-edit'+opId+'_1';
    ifr.onload = function() {
      setTimeout(()=>{
        if (done) return;
        try {
          const doc = ifr.contentDocument;
          if (!doc?.body) { fail('modal vazio'); clearTimeout(safe); return; }
          let n=0;
          for (let i=1;i<=8;i++) {
            const r = doc.querySelector(`input[name="p${i}_confirm"][value="S"]`);
            if (r) { r.checked=true; r.dispatchEvent(new Event('change',{bubbles:true})); r.dispatchEvent(new Event('click',{bubbles:true})); n++; }
          }
          if (!n) { fail('formulario nao encontrado'); clearTimeout(safe); return; }
          setTimeout(()=>{
            if (done) return;
            const s = doc.querySelector('button[name="submitF"]');
            if (!s) { fail('botao nao encontrado'); clearTimeout(safe); return; }
            s.click();
            setTimeout(()=>{
              if (done) return;
              done=true; clearTimeout(safe);
              btn.innerHTML='Enviado!'; btn.style.color='#3ECF8E';
              setTimeout(()=>window.location.reload(),1500);
            },2500);
          },700);
        } catch(e) { fail('erro'); clearTimeout(safe); }
      },2000);
    };
  }
  window._monEnviarEscala = enviarEscala;
 
  // ── COPIAR CHAVE ──────────────────────────────────────
  function copiarChave(chave, btn) {
    const orig = btn.innerHTML;
    const ok = () => { btn.innerHTML='Copiado!'; btn.style.color='#3ECF8E'; setTimeout(()=>{ btn.innerHTML=orig; btn.style.color=''; },2000); };
    if (navigator.clipboard) { navigator.clipboard.writeText(chave).then(ok).catch(()=>{}); return; }
    const ta = document.createElement('textarea');
    ta.value=chave; document.body.appendChild(ta); ta.select(); document.execCommand('copy'); document.body.removeChild(ta); ok();
  }
  window._monCopiarChave = copiarChave;
 
  // ── REFRESH ───────────────────────────────────────────
  function manualRefresh() {
    if (refreshTimer) clearInterval(refreshTimer);
    iframesInUse.clear(); fetchQueue=[]; inQueue=new Set();
    QUEUE_IFRS.forEach(id=>{ const f=document.getElementById(id); if(f) f.onload=null; });
    Object.keys(apontCache).forEach(k=>{ if(apontCache[k]==='loading') delete apontCache[k]; });
    fetchOperations();
    refreshTimer = setInterval(silentRefresh,60000);
  }
  window._monRefresh = manualRefresh;
 
  function silentRefresh() {
    const ops = parseOps(document);
    if (!ops.length) return;
    const oldIds = new Set(operations.map(o=>o.id));
    const newIds = new Set(ops.map(o=>o.id));
    operations.filter(o=>!newIds.has(o.id)).forEach(o=>{
      delete apontCache[o.id]; expanded.delete(o.chave); monitoradas.delete(monKey(o));
    });
    ops.forEach(o=>{ if(dentroJanela(o)) monitoradas.add(monKey(o)); });
    operations = ops;
    renderTable();
    ops.filter(o=>o.id).forEach((op,i)=>{
      setTimeout(()=>{
        const cached = apontCache[op.id];
        if (cached==='loading') return;
        if (!oldIds.has(op.id)) {
          monitoradas.add(monKey(op));
          enfileirar(op,d=>{ updateCells(op,d); updateMetrics(); });
        } else if (naJanela(op)) {
          delete apontCache[op.id];
          enfileirar(op,d=>{ updateCells(op,d); updateMetrics(); });
        } else if (!cached||cached._erro) {
          enfileirar(op,d=>updateCells(op,d));
        }
      },i*150);
    });
    const sub=document.getElementById('mon-sub');
    if(sub) sub.textContent=new Date().toLocaleTimeString('pt-BR');
  }
 
  function updateCells(op, d) {
    if (!d||d==='loading') return;
    const row=document.querySelector(`tr[data-chave="${op.chave}"]`);
    if (!row) return;
    const cells=row.querySelectorAll('td');
    if (cells[3]) cells[3].innerHTML=renderEscBadge(d,op.qtd);
    if (naJanela(op)&&cells[4]) cells[4].innerHTML=renderAptBadge(d,op.qtd);
    if (cells[7]) cells[7].innerHTML=renderStatus(d);
    if (expanded.has(op.chave)) {
      const idx=operations.findIndex(o=>o.chave===op.chave);
      const det=document.getElementById('det-'+idx);
      if (det) det.querySelector('.det-inner').innerHTML=renderDetail(op);
    }
  }
 
  // ── CONTROLES ─────────────────────────────────────────
  function initControls(panel) {
    const rh=document.createElement('div');
    rh.style.cssText='position:absolute;left:0;top:0;width:5px;height:100%;cursor:ew-resize;z-index:10;';
    rh.addEventListener('mousedown',e=>{
      e.preventDefault();
      const sx=e.clientX,sw=panel.offsetWidth;
      document.body.style.userSelect='none';
      const mv=e=>panel.style.width=Math.min(Math.max(sw+sx-e.clientX,480),window.innerWidth-60)+'px';
      const up=()=>{ document.body.style.userSelect=''; document.removeEventListener('mousemove',mv); document.removeEventListener('mouseup',up); };
      document.addEventListener('mousemove',mv); document.addEventListener('mouseup',up);
    });
    panel.appendChild(rh);
    const hdr=panel.querySelector('#mon-header');
    if (hdr) {
      hdr.addEventListener('mousedown',e=>{
        if(e.target.closest('button,a,input')) return;
        e.preventDefault(); hdr.style.cursor='grabbing';
        const rect=panel.getBoundingClientRect(),ox=e.clientX-rect.left,oy=e.clientY-rect.top;
        document.body.style.userSelect='none';
        const mv=e=>{
          panel.style.left=Math.max(0,Math.min(e.clientX-ox,window.innerWidth-panel.offsetWidth))+'px';
          panel.style.top=Math.max(0,Math.min(e.clientY-oy,window.innerHeight-60))+'px';
          panel.style.right='auto';
        };
        const up=()=>{ hdr.style.cursor='grab'; document.body.style.userSelect=''; document.removeEventListener('mousemove',mv); document.removeEventListener('mouseup',up); };
        document.addEventListener('mousemove',mv); document.addEventListener('mouseup',up);
      });
    }
  }
 
  window._monMinimize=function(){
    const body=document.getElementById('mon-body'),btn=document.getElementById('mon-min-btn'),panel=document.getElementById('mon-panel');
    if(!body||!btn||!panel) return;
    minimized=!minimized;
    body.style.display=minimized?'none':'';
    panel.style.height=minimized?'auto':'100vh';
    btn.textContent=minimized?'□':'—';
  };
 
  // ── BOTAO ─────────────────────────────────────────────
  function injectButton() {
    if(document.getElementById('btn-mon')||window.self!==window.top) return;
    const btn=document.createElement('button');
    btn.id='btn-mon'; btn.textContent='Monitor';
    btn.style.cssText='position:fixed;bottom:20px;right:20px;z-index:99999;background:#0B1520;color:#5B9CF6;border:1px solid #1E3A5F;padding:9px 18px;border-radius:7px;font-size:11px;font-family:"DM Mono",monospace;font-weight:500;cursor:pointer;letter-spacing:0.08em;box-shadow:0 4px 20px rgba(0,0,0,0.6);transition:all 0.15s;';
    btn.onmouseover=()=>{ btn.style.background='#112030'; };
    btn.onmouseout=()=>{ btn.style.background='#0B1520'; };
    btn.onclick=toggleMonitor;
    document.body.appendChild(btn);
  }
 
  // ── ESTILOS ───────────────────────────────────────────
  function injectStyles() {
    if(document.getElementById('mon-style')) return;
    const s=document.createElement('style');
    s.id='mon-style';
    s.textContent=`
      @import url('https://fonts.googleapis.com/css2?family=DM+Mono:wght@400;500&family=Inter:wght@400;500;600&display=swap');
      #mon-panel{--bg:#0A0F18;--bg2:#0D1520;--bg3:#111D2C;--bg4:#162234;--border:#1A2D42;--border2:#142030;--text:#C5D5E5;--dim:#506878;--faint:#2A4060;--ok:#3ECF8E;--warn:#F0A050;--danger:#E05C5C;--esc:#5B9CF6;--font:'Inter',system-ui,sans-serif;--mono:'DM Mono','Courier New',monospace;}
      @keyframes mon-slidein{from{opacity:0;transform:translateX(16px)}to{opacity:1;transform:none}}
      @keyframes mon-fadein{from{opacity:0;transform:translateY(4px)}to{opacity:1;transform:none}}
      @keyframes mon-blink{0%,100%{opacity:1}50%{opacity:0.3}}
      #mon-panel{font-family:var(--font);font-size:12px;animation:mon-slidein 0.2s ease;}
      #mon-panel *{box-sizing:border-box;}
      #mon-panel ::-webkit-scrollbar{width:4px;}
      #mon-panel ::-webkit-scrollbar-track{background:var(--bg);}
      #mon-panel ::-webkit-scrollbar-thumb{background:var(--border);border-radius:2px;}
      #mon-header{background:var(--bg2);border-bottom:1px solid var(--border);padding:0 16px;height:54px;display:flex;align-items:center;justify-content:space-between;flex-shrink:0;user-select:none;cursor:grab;}
      .mon-logo{font-size:11px;font-weight:600;letter-spacing:0.15em;text-transform:uppercase;color:var(--text);}
      .mon-live{font-family:var(--mono);font-size:9px;letter-spacing:0.1em;padding:2px 8px;border-radius:3px;border:1px solid;margin-left:10px;text-transform:uppercase;}
      .mon-live.offline{color:var(--faint);border-color:var(--faint);}
      .mon-live.sync{color:var(--warn);border-color:rgba(240,160,80,0.3);background:rgba(240,160,80,0.06);animation:mon-blink 1s infinite;}
      .mon-live.live{color:var(--ok);border-color:rgba(62,207,142,0.3);background:rgba(62,207,142,0.06);}
      .mon-hbtn{font-family:var(--font);font-size:11px;padding:5px 12px;border-radius:5px;cursor:pointer;transition:all 0.12s;border:1px solid var(--border);background:transparent;color:var(--dim);}
      .mon-hbtn:hover{background:var(--bg3);color:var(--text);}
      .mon-hbtn.ok{color:var(--ok);border-color:rgba(62,207,142,0.25);}
      .mon-hbtn.ok:hover{background:rgba(62,207,142,0.06);}
      .mon-hbtn.active{color:var(--ok);border-color:rgba(62,207,142,0.3);}
      .mon-hbtn.blocked{color:var(--danger);border-color:rgba(224,92,92,0.3);}
      .mon-prog-wrap{flex:1;max-width:150px;margin:0 10px;}
      .mon-prog-track{height:2px;background:var(--border);border-radius:1px;overflow:hidden;}
      #mon-prog-bar{height:100%;width:0%;background:var(--warn);border-radius:1px;transition:width 0.4s ease,background 0.4s;}
      #mon-prog-label{font-family:var(--mono);font-size:9px;color:var(--faint);margin-top:3px;}
      #mon-metrics{display:grid;grid-template-columns:repeat(4,1fr);border-bottom:1px solid var(--border);flex-shrink:0;}
      .mon-metric{padding:14px 16px;border-right:1px solid var(--border2);}
      .mon-metric:last-child{border-right:none;}
      .mon-metric-label{font-size:9px;font-weight:600;letter-spacing:0.12em;text-transform:uppercase;color:var(--faint);margin-bottom:6px;}
      .mon-metric-val{font-family:var(--mono);font-size:26px;font-weight:500;line-height:1;}
      .mon-metric-bar{height:2px;background:var(--border);border-radius:1px;overflow:hidden;margin-top:8px;}
      .mon-metric-bar-fill{height:100%;border-radius:1px;transition:width 0.6s ease;}
      .mon-table-wrap{flex:1;overflow-y:auto;}
      #mon-table{width:100%;border-collapse:collapse;table-layout:fixed;}
      #mon-table thead tr{background:var(--bg2);border-bottom:1px solid var(--border);position:sticky;top:0;z-index:2;}
      #mon-table thead th{padding:8px 12px;text-align:left;font-size:9px;font-weight:600;letter-spacing:0.12em;text-transform:uppercase;color:var(--faint);}
      #mon-table thead th.c{text-align:center;}
      #mon-table tbody tr.op-row{border-bottom:1px solid var(--border2);cursor:pointer;transition:background 0.1s;}
      #mon-table tbody tr.op-row:hover td{background:var(--bg3);}
      #mon-table tbody tr.op-row.exp td{background:var(--bg3);}
      #mon-table tbody td{padding:9px 12px;vertical-align:middle;}
      .c-chave{font-family:var(--mono);font-size:10px;color:var(--dim);overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
      .c-sigla{font-family:var(--mono);font-size:12px;font-weight:500;color:var(--text);}
      .c-site{font-size:11px;color:var(--dim);overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
      .c-hora{font-family:var(--mono);font-size:11px;color:var(--dim);}
      .c-lider{font-size:11px;color:var(--dim);overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
      .c-tog{text-align:center;color:var(--faint);font-size:10px;}
      .num-badge{display:flex;flex-direction:column;align-items:center;gap:3px;}
      .num-badge .num{font-family:var(--mono);font-size:12px;font-weight:500;line-height:1;}
      .num-badge .den{font-family:var(--mono);font-size:9px;color:var(--faint);}
      .num-badge .btrack{width:52px;height:2px;background:rgba(255,255,255,0.07);border-radius:1px;overflow:hidden;}
      .num-badge .bfill{height:100%;border-radius:1px;transition:width 0.5s ease;}
      .mon-status{display:inline-flex;align-items:center;font-size:10px;font-weight:600;letter-spacing:0.06em;padding:3px 10px;border-radius:4px;border:1px solid transparent;white-space:nowrap;}
      .mon-status.ok{color:var(--ok);border-color:rgba(62,207,142,0.2);background:rgba(62,207,142,0.07);}
      .mon-status.warn{color:var(--warn);border-color:rgba(240,160,80,0.2);background:rgba(240,160,80,0.07);}
      .mon-status.danger{color:var(--danger);border-color:rgba(224,92,92,0.2);background:rgba(224,92,92,0.07);}
      .mon-status.esc{color:var(--esc);border-color:rgba(91,156,246,0.2);background:rgba(91,156,246,0.07);}
      .mon-status.empty{color:var(--faint);}
      tr.det-row{border-bottom:1px solid var(--border);}
      .det-inner{background:var(--bg);border-top:1px solid var(--border2);padding:16px;animation:mon-fadein 0.2s ease;}
      .det-grid{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:14px;}
      .det-card{background:var(--bg2);border:1px solid var(--border2);border-radius:6px;padding:12px 14px;min-width:0;}
      .det-card-label{font-size:9px;font-weight:600;letter-spacing:0.12em;text-transform:uppercase;color:var(--faint);margin-bottom:8px;}
      .det-card-row{display:flex;align-items:baseline;justify-content:space-between;margin-bottom:6px;}
      .det-card-num{font-family:var(--mono);font-size:22px;font-weight:500;line-height:1;}
      .det-card-den{font-family:var(--mono);font-size:11px;color:var(--faint);}
      .det-card-pct{font-family:var(--mono);font-size:11px;color:var(--dim);}
      .det-bar-track{height:4px;background:rgba(255,255,255,0.06);border-radius:2px;overflow:hidden;}
      .det-bar-fill{height:100%;border-radius:2px;transition:width 0.7s cubic-bezier(0.4,0,0.2,1);}
      .det-actions{display:flex;flex-wrap:wrap;gap:6px;margin-bottom:14px;align-items:center;}
      .det-btn-escala{font-family:var(--font);font-size:11px;font-weight:500;padding:6px 14px;border-radius:5px;border:1px solid rgba(91,156,246,0.3);background:rgba(91,156,246,0.07);color:var(--esc);cursor:pointer;transition:all 0.15s;}
      .det-btn-escala:hover{background:rgba(91,156,246,0.13);}
      .det-btn-escala:disabled{opacity:0.5;cursor:not-allowed;}
      .det-btn-copy{font-family:var(--mono);font-size:10px;padding:5px 12px;border-radius:5px;border:1px solid var(--border);background:transparent;color:var(--dim);cursor:pointer;transition:all 0.12s;}
      .det-btn-copy:hover{background:var(--bg3);color:var(--text);}
      .det-btn-link{font-family:var(--font);font-size:11px;padding:5px 12px;border-radius:5px;border:1px solid var(--border);background:transparent;color:var(--dim);cursor:pointer;transition:all 0.12s;text-decoration:none;display:inline-flex;align-items:center;gap:5px;}
      .det-btn-link:hover{background:var(--bg3);color:var(--text);}
      .det-xls-menu{position:relative;display:inline-block;}
      .det-xls-drop{display:none;position:absolute;bottom:calc(100% + 4px);left:0;background:var(--bg2);border:1px solid var(--border);border-radius:6px;padding:4px 0;z-index:999;min-width:185px;box-shadow:0 8px 24px rgba(0,0,0,0.6);}
      .det-xls-menu:hover .det-xls-drop{display:block;}
      .det-xls-drop a{display:block;padding:7px 12px;color:var(--dim);font-size:11px;text-decoration:none;transition:all 0.1s;border-bottom:1px solid var(--border2);font-family:var(--font);}
      .det-xls-drop a:last-child{border-bottom:none;}
      .det-xls-drop a:hover{background:var(--bg3);color:var(--text);}
      .det-section-title{font-size:9px;font-weight:600;letter-spacing:0.12em;text-transform:uppercase;margin-bottom:8px;padding-bottom:6px;border-bottom:1px solid var(--border2);}
      .det-section-title.danger{color:var(--danger);}
      .det-section-title.dim{color:var(--faint);}
      .det-colab-table{width:100%;border-collapse:collapse;font-size:11px;margin-bottom:14px;}
      .det-colab-table th{padding:5px 8px;text-align:left;font-size:9px;font-weight:600;letter-spacing:0.1em;text-transform:uppercase;color:var(--faint);border-bottom:1px solid var(--border2);}
      .det-colab-table td{padding:6px 8px;border-bottom:1px solid var(--border2);vertical-align:middle;}
      .det-colab-table tr:last-child td{border-bottom:none;}
      .det-colab-table .nm{color:var(--text);font-weight:500;}
      .det-colab-table .nm.danger{color:var(--danger);}
      .det-colab-table .tp{color:var(--dim);}
      .det-colab-table .ini{font-family:var(--mono);color:var(--faint);font-size:10px;}
    `;
    document.head.appendChild(s);
  }
 
  // ── PAINEL ────────────────────────────────────────────
  function createPanel() {
    if(document.getElementById('mon-panel')) return;
    injectStyles();
    const notifTxt=!('Notification' in window)?'Sem suporte':Notification.permission==='granted'?'Notif. ativas':Notification.permission==='denied'?'Bloqueado':'Notificacoes';
    const notifCls=Notification.permission==='granted'?'mon-hbtn active':Notification.permission==='denied'?'mon-hbtn blocked':'mon-hbtn';
    const p=document.createElement('div');
    p.id='mon-panel';
    p.style.cssText='position:fixed;top:0;right:0;width:1020px;height:100vh;background:var(--bg);color:var(--text);z-index:99998;box-shadow:-8px 0 40px rgba(0,0,0,0.7);display:none;flex-direction:column;overflow:hidden;border-left:1px solid var(--border);';
    p.innerHTML=`
      <div id="mon-header">
        <div style="display:flex;align-items:center">
          <span class="mon-logo">Monitor Operacional</span>
          <span id="mon-live" class="mon-live offline">Offline</span>
        </div>
        <div style="display:flex;align-items:center;gap:8px">
          <div class="mon-prog-wrap">
            <div class="mon-prog-track"><div id="mon-prog-bar"></div></div>
            <div id="mon-prog-label">—</div>
          </div>
          <span id="mon-sub" style="font-family:var(--mono);font-size:9px;color:var(--faint)"></span>
          <button class="mon-hbtn ok" onclick="window._monRefresh()">Atualizar</button>
          <button id="mon-notif-btn" class="${notifCls}" onclick="window._monPedirNotif()">${notifTxt}</button>
          <button id="mon-min-btn" class="mon-hbtn" onclick="window._monMinimize()" style="padding:5px 10px">—</button>
          <button class="mon-hbtn" onclick="document.getElementById('mon-panel').style.display='none';document.getElementById('btn-mon').textContent='Monitor'" style="padding:5px 10px">x</button>
        </div>
      </div>
      <div id="mon-body" style="display:flex;flex-direction:column;flex:1;overflow:hidden">
        <div id="mon-metrics">
          <div class="mon-metric">
            <div class="mon-metric-label">Operacoes</div>
            <div class="mon-metric-val" id="m-total" style="color:var(--text)">—</div>
            <div class="mon-metric-bar"><div class="mon-metric-bar-fill" id="mb-total" style="background:var(--esc)"></div></div>
          </div>
          <div class="mon-metric">
            <div class="mon-metric-label">Completas</div>
            <div class="mon-metric-val" id="m-ok" style="color:var(--ok)">—</div>
            <div class="mon-metric-bar"><div class="mon-metric-bar-fill" id="mb-ok" style="background:var(--ok)"></div></div>
          </div>
          <div class="mon-metric">
            <div class="mon-metric-label">Parciais</div>
            <div class="mon-metric-val" id="m-inc" style="color:var(--warn)">—</div>
            <div class="mon-metric-bar"><div class="mon-metric-bar-fill" id="mb-inc" style="background:var(--warn)"></div></div>
          </div>
          <div class="mon-metric">
            <div class="mon-metric-label">Sem apontamento</div>
            <div class="mon-metric-val" id="m-zero" style="color:var(--danger)">—</div>
            <div class="mon-metric-bar"><div class="mon-metric-bar-fill" id="mb-zero" style="background:var(--danger)"></div></div>
          </div>
        </div>
        <div class="mon-table-wrap">
          <table id="mon-table">
            <thead><tr>
              <th style="width:13%">Chave</th>
              <th style="width:7%">Sigla</th>
              <th style="width:17%">Site</th>
              <th style="width:11%" class="c">Esc/Sol</th>
              <th style="width:11%" class="c">Apt/Sol</th>
              <th style="width:7%">Hora</th>
              <th style="width:14%">Lider</th>
              <th style="width:14%" class="c">Status</th>
              <th style="width:4%"></th>
            </tr></thead>
            <tbody id="mon-tbody"></tbody>
          </table>
        </div>
      </div>
    `;
    document.body.appendChild(p);
    initControls(p);
  }
 
  function toggleMonitor() {
    if(!document.getElementById('mon-panel')) createPanel();
    const p=document.getElementById('mon-panel'),btn=document.getElementById('btn-mon');
    if(p.style.display==='none'||!p.style.display){p.style.display='flex';p.style.flexDirection='column';if(btn)btn.textContent='Fechar';}
    else{p.style.display='none';if(btn)btn.textContent='Monitor';}
  }
 
  // ── RENDER TABLE ──────────────────────────────────────
  function renderTable() {
    const tbody=document.getElementById('mon-tbody');
    if(!tbody) return;
    tbody.innerHTML='';
    if(!operations.length){
      tbody.innerHTML='<tr><td colspan="9" style="text-align:center;padding:3rem;color:var(--faint)">Nenhuma operacao encontrada</td></tr>';
      return;
    }
    operations.forEach((op,idx)=>{
      const isExp=expanded.has(op.chave);
      const d=apontCache[op.id];
      const emJanela=naJanela(op);
      const temD=d&&d!=='loading';
      const tr=document.createElement('tr');
      tr.className='op-row'+(isExp?' exp':'');
      tr.dataset.chave=op.chave;
      if(!emJanela) tr.style.opacity='0.4';
      tr.innerHTML=`
        <td class="c-chave" title="${op.chave}">${op.chave}</td>
        <td class="c-sigla">${op.sigla}</td>
        <td class="c-site" title="${op.site}">${op.site}</td>
        <td style="padding:7px 12px;text-align:center">${op.id?(temD?renderEscBadge(d,op.qtd):`<span style="color:var(--faint);font-family:var(--mono);font-size:11px">.../${op.qtd}</span>`):'<span style="color:var(--faint)">-</span>'}</td>
        <td style="padding:7px 12px;text-align:center">${emJanela?(temD?renderAptBadge(d,op.qtd):`<span style="color:var(--faint);font-family:var(--mono);font-size:11px">.../${op.qtd}</span>`):'<span style="color:var(--faint)">-</span>'}</td>
        <td class="c-hora">${op.hora}</td>
        <td class="c-lider" title="${op.lider}">${op.lider}</td>
        <td style="text-align:center">${renderStatus(temD?d:null)}</td>
        <td class="c-tog">${isExp?'▴':'▾'}</td>
      `;
      tr.onclick=()=>toggleRow(op,idx);
      tbody.appendChild(tr);
      if(isExp){
        const det=document.createElement('tr');
        det.id='det-'+idx; det.className='det-row';
        det.innerHTML=`<td colspan="9" style="padding:0"><div class="det-inner">${renderDetail(op)}</div></td>`;
        tbody.appendChild(det);
      }
    });
    updateMetrics();
  }
 
  function renderDetail(op) {
    const d=apontCache[op.id];
    if(!d||d==='loading') return '<span style="color:var(--faint);font-size:11px">Carregando...</span>';
    const escPct=op.qtd>0?Math.min(100,Math.round((d.escalado/op.qtd)*100)):0;
    const aptPct=op.qtd>0?Math.min(100,Math.round((d.apontado/op.qtd)*100)):0;
    const escCor=escPct>=100?'var(--ok)':escPct>0?'var(--esc)':'var(--faint)';
    const aptCor=aptPct>=100?'var(--ok)':aptPct>0?'var(--warn)':'var(--faint)';
 
    let html=`<div class="det-grid">
      <div class="det-card">
        <div class="det-card-label">Escalados</div>
        <div class="det-card-row"><span class="det-card-num" style="color:${escCor}">${d.escalado}<span class="det-card-den">/${op.qtd}</span></span><span class="det-card-pct">${escPct}%</span></div>
        <div class="det-bar-track"><div class="det-bar-fill" style="width:${escPct}%;background:${escCor}"></div></div>
      </div>
      <div class="det-card">
        <div class="det-card-label">Apontados</div>
        <div class="det-card-row"><span class="det-card-num" style="color:${aptCor}">${d.apontado}<span class="det-card-den">/${op.qtd}</span></span><span class="det-card-pct">${aptPct}%</span></div>
        <div class="det-bar-track"><div class="det-bar-fill" style="width:${aptPct}%;background:${aptCor}"></div></div>
      </div>
    </div>`;
 
    const listaLinks=d.listaLinks||[],xlsLinks=d.xlsLinks||[];
    html+=`<div class="det-actions">
      <button class="det-btn-escala" onclick="event.stopPropagation();window._monEnviarEscala('${op.id}',this)">Confirmar escala enviada</button>
      <button class="det-btn-copy" onclick="event.stopPropagation();window._monCopiarChave('${op.chave}',this)">Copiar chave</button>`;
    listaLinks.forEach(l=>{
      html+=`<a href="https://tsi-app.com/${l.href}" target="_blank" class="det-btn-link">Lista${l.label?' - '+l.label.split(' ')[0]:''}</a>`;
    });
    if(xlsLinks.length){
      html+=`<div class="det-xls-menu"><button class="det-btn-link">XLS</button><div class="det-xls-drop">`;
      xlsLinks.forEach(l=>{html+=`<a href="https://tsi-app.com/${l.href}" target="_blank">${l.label}</a>`;});
      html+=`</div></div>`;
    }
    html+=`</div>`;
 
    const faltando=d.faltando||[];
    if(faltando.length){
      html+=`<div class="det-section-title danger">Aguardando apontamento - ${faltando.length}</div>
        <table class="det-colab-table"><thead><tr><th>Nome</th><th>Tipo</th></tr></thead><tbody>`;
      faltando.forEach(c=>{html+=`<tr><td class="nm danger">${c.nome}</td><td class="tp">${c.tipo||'-'}</td></tr>`;});
      html+=`</tbody></table>`;
    }
    const colab=d.colaboradores||[];
    if(colab.length){
      html+=`<div class="det-section-title dim">Apontados - ${colab.length}</div>
        <table class="det-colab-table"><thead><tr><th>Nome</th><th>Tipo</th><th>Inicio</th></tr></thead><tbody>`;
      colab.forEach(c=>{html+=`<tr><td class="nm">${c.nome}</td><td class="tp">${c.tipo||'-'}</td><td class="ini">${c.inicio}</td></tr>`;});
      html+=`</tbody></table>`;
    }
    if(!colab.length&&!faltando.length) html+=`<div style="color:var(--faint);font-size:11px;padding:8px 0">Sem dados de escala ou apontamento.</div>`;
    return html;
  }
 
  function toggleRow(op,idx){
    if(expanded.has(op.chave)){expanded.delete(op.chave);renderTable();return;}
    expanded.add(op.chave);
    renderTable();
    const c=apontCache[op.id];
    if(!c||c==='loading'){
      const poll=setInterval(()=>{
        const cc=apontCache[op.id];
        if(cc&&cc!=='loading'){
          clearInterval(poll);
          const det=document.getElementById('det-'+idx);
          if(det) det.querySelector('.det-inner').innerHTML=renderDetail(op);
          updateMetrics();
        }
      },500);
      setTimeout(()=>clearInterval(poll),40000);
    }
  }
 
  // ── BADGES ────────────────────────────────────────────
  function renderEscBadge(d,qtd){
    if(!d||d==='loading') return `<span style="color:var(--faint);font-family:var(--mono);font-size:11px">.../${qtd}</span>`;
    const pct=qtd>0?Math.min(100,Math.round((d.escalado/qtd)*100)):0;
    const cor=pct>=100?'var(--ok)':pct>0?'var(--esc)':'var(--faint)';
    return `<div class="num-badge"><span class="num" style="color:${cor}">${d.escalado}<span class="den">/${qtd}</span></span><div class="btrack"><div class="bfill" style="width:${pct}%;background:${cor}"></div></div></div>`;
  }
  function renderAptBadge(d,qtd){
    if(!d||d==='loading') return `<span style="color:var(--faint);font-family:var(--mono);font-size:11px">.../${qtd}</span>`;
    const pct=qtd>0?Math.min(100,Math.round((d.apontado/qtd)*100)):0;
    const cor=pct>=100?'var(--ok)':pct>0?'var(--warn)':'var(--faint)';
    return `<div class="num-badge"><span class="num" style="color:${cor}">${d.apontado}<span class="den">/${qtd}</span></span><div class="btrack"><div class="bfill" style="width:${pct}%;background:${cor}"></div></div></div>`;
  }
  function renderStatus(d){
    if(!d||d==='loading') return '<span class="mon-status empty">—</span>';
    if(d._soEscala){
      if(!d.escalado) return '<span class="mon-status danger">Nenhum</span>';
      if(d.escalado>=d.solicitado) return '<span class="mon-status esc">Esc. OK</span>';
      return `<span class="mon-status esc">Esc. ${d.escalado}/${d.solicitado}</span>`;
    }
    if(d.apontado>=d.solicitado&&d.escalado>=d.solicitado) return '<span class="mon-status ok">Completo</span>';
    if(!d.apontado&&!d.escalado) return '<span class="mon-status danger">Nenhum</span>';
    if(!d.apontado) return `<span class="mon-status esc">Esc. ${d.escalado}/${d.solicitado}</span>`;
    return `<span class="mon-status warn">Parcial ${d.apontado}/${d.solicitado}</span>`;
  }
 
  // ── METRICAS ──────────────────────────────────────────
  function updateProgress(loaded,total){
    const bar=document.getElementById('mon-prog-bar'),lbl=document.getElementById('mon-prog-label');
    if(!bar||!lbl) return;
    const pct=total===0?100:Math.round((loaded/total)*100);
    bar.style.width=pct+'%';
    if(pct>=100){bar.style.background='var(--ok)';lbl.textContent='Carregado';lbl.style.color='var(--ok)';}
    else{bar.style.background='var(--warn)';lbl.textContent=pct+'%';lbl.style.color='var(--warn)';}
  }
  function updateMetrics(){
    let ok=0,inc=0,zero=0;
    const emJanela=operations.filter(o=>naJanela(o));
    emJanela.forEach(op=>{
      const d=apontCache[op.id];
      if(!d||d==='loading') return;
      if(d.apontado===0) zero++;
      else if(d.apontado>=d.solicitado) ok++;
      else inc++;
    });
    const total=emJanela.length;
    const sv=(id,v)=>{const el=document.getElementById(id);if(el)el.textContent=v;};
    const sb=(id,pct)=>{const el=document.getElementById(id);if(el)el.style.width=pct+'%';};
    sv('m-total',total); sb('mb-total',total?100:0);
    sv('m-ok',ok);       sb('mb-ok',total?Math.round(ok/total*100):0);
    sv('m-inc',inc);     sb('mb-inc',total?Math.round(inc/total*100):0);
    sv('m-zero',zero);   sb('mb-zero',total?Math.round(zero/total*100):0);
  }
  function setLive(txt,cls){const el=document.getElementById('mon-live');if(el){el.textContent=txt;el.className='mon-live '+cls;}}
 
  // ── MOTOR ─────────────────────────────────────────────
  function startMonitor(){
    window._monRunning=true;
    fetchOperations();
    refreshTimer=setInterval(silentRefresh,60000);
  }
 
  function fetchOperations(){
    setLive('Sincronizando','sync');
    const ops1=parseOps(document);
    const ifr2=document.getElementById(IFR_PAG2);
 
    const finalizar=opsAll=>{
      const seen=new Set();
      const ops=opsAll.filter(o=>{if(seen.has(o.chave))return false;seen.add(o.chave);return true;});
 
      restoreCache(); // restaura antes de redefinir operations
 
      operations=ops;
      expanded=new Set();
      monitoradas=new Set();
      fetchQueue=[]; inQueue=new Set();
 
      const opsComId=ops.filter(o=>o.id);
      opsComId.forEach(o=>{if(dentroJanela(o))monitoradas.add(monKey(o));});
      renderTable();
      setLive('Online','live');
      const sub=document.getElementById('mon-sub');
      if(sub) sub.textContent=new Date().toLocaleTimeString('pt-BR');
 
      // Ops com cache valido: atualiza direto
      const comCache=opsComId.filter(o=>apontCache[o.id]&&apontCache[o.id]!=='loading');
      const semCache=opsComId.filter(o=>!apontCache[o.id]||apontCache[o.id]==='loading');
 
      comCache.forEach(op=>{
        if(dentroJanela(op))monitoradas.add(monKey(op));
        updateCells(op,apontCache[op.id]);
      });
      updateMetrics();
 
      let loaded=comCache.length;
      const total=opsComId.length;
      updateProgress(loaded,total);
 
      semCache.forEach(op=>{
        enfileirar(op,d=>{
          if(dentroJanela(op))monitoradas.add(monKey(op));
          loaded++;
          updateProgress(loaded,total);
          updateCells(op,d);
          updateMetrics();
        });
      });
    };
 
    if(!ifr2){finalizar(ops1);return;}
    let p2done=false;
    const p2t=setTimeout(()=>{if(!p2done){p2done=true;finalizar(ops1);}},8000);
    ifr2.onload=null;
    ifr2.src='https://tsi-app.com/planejamento-operacional_2';
    ifr2.onload=function(){
      if(p2done)return; p2done=true; clearTimeout(p2t);
      setTimeout(()=>{try{finalizar([...ops1,...parseOps(ifr2.contentDocument)]);}catch(e){finalizar(ops1);}},1500);
    };
  }
 
  // ── INIT ──────────────────────────────────────────────
  setTimeout(()=>{
    ensureIframes();
    injectButton();
    createPanel();
    if(!window._monRunning) startMonitor();
    if(Notification.permission==='default') pedirPermissao();
  },2000);
 
})();
 
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
