<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Gestão de Doca — Sistema Operacional</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.2/dist/chart.umd.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
<style>
*{box-sizing:border-box;margin:0;padding:0}
:root{
  --c-bg:#0d1117;
  --c-surface:#161b22;
  --c-surface2:#1c2331;
  --c-border:#30363d;
  --c-border2:#444c56;
  --c-text:#c9d1d9;
  --c-muted:#8b949e;
  --c-accent:#e6a817;
  --c-accent-dim:#b8820e;
  --c-green:#3fb950;
  --c-red:#f85149;
  --c-blue:#58a6ff;
  --c-purple:#bc8cff;
  --c-orange:#ffa657;
  --c-teal:#39d353;
  --font:'Consolas','SF Mono','Cascadia Code',monospace;
}
body{background:var(--c-bg);color:var(--c-text);font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;font-size:13px;line-height:1.5;min-height:100vh}

/* ─── SCREENS ─────────────────── */
.screen{display:none;min-height:100vh}
.screen.active{display:flex;flex-direction:column}

/* ─── LOGIN ─────────────────────── */
#screen-login{align-items:center;justify-content:center;background:var(--c-bg)}
.login-card{
  background:var(--c-surface);
  border:1px solid var(--c-border);
  border-radius:8px;
  padding:40px 44px;
  width:360px;
}
.login-card .brand{display:flex;align-items:center;gap:10px;margin-bottom:32px}
.brand-icon{width:36px;height:36px;background:var(--c-accent);border-radius:6px;display:flex;align-items:center;justify-content:center}
.brand-icon svg{width:20px;height:20px;fill:#000}
.brand-name{font-size:15px;font-weight:600;color:#fff}
.brand-sub{font-size:11px;color:var(--c-muted)}
.login-card label{display:block;font-size:11px;font-weight:600;color:var(--c-muted);text-transform:uppercase;letter-spacing:.06em;margin-bottom:6px}
.login-card input{
  width:100%;background:var(--c-bg);border:1px solid var(--c-border);
  border-radius:6px;padding:9px 12px;color:var(--c-text);font-size:13px;outline:none;
  transition:border-color .15s;
}
.login-card input:focus{border-color:var(--c-accent)}
.login-card .field{margin-bottom:16px}
.login-err{font-size:12px;color:var(--c-red);margin-top:8px;display:none}
.btn-primary{
  width:100%;background:var(--c-accent);border:none;border-radius:6px;
  padding:10px;color:#000;font-weight:700;font-size:13px;cursor:pointer;
  transition:background .15s;margin-top:8px;
}
.btn-primary:hover{background:var(--c-accent-dim)}
.btn-sm{padding:5px 12px;border-radius:5px;border:1px solid var(--c-border);background:var(--c-surface2);color:var(--c-text);font-size:12px;cursor:pointer;transition:all .15s}
.btn-sm:hover{border-color:var(--c-border2);background:var(--c-border)}
.btn-green{background:#238636;border-color:#2ea043;color:#fff}
.btn-green:hover{background:#2ea043}
.btn-red{background:#b62324;border-color:#da3633;color:#fff}
.btn-red:hover{background:#da3633}
.btn-amber{background:#9e6a03;border-color:var(--c-accent);color:#fff}
.btn-amber:hover{background:#b8820e}

/* ─── TOPBAR ─────────────────────── */
.topbar{
  background:var(--c-surface);border-bottom:1px solid var(--c-border);
  padding:0 20px;height:44px;display:flex;align-items:center;gap:0;
  flex-shrink:0;position:sticky;top:0;z-index:100;
}
.topbar-brand{display:flex;align-items:center;gap:8px;padding-right:20px;border-right:1px solid var(--c-border);margin-right:4px}
.topbar-brand .icon{width:24px;height:24px;background:var(--c-accent);border-radius:4px;display:flex;align-items:center;justify-content:center}
.topbar-brand .icon svg{width:13px;height:13px;fill:#000}
.topbar-brand span{font-size:13px;font-weight:600;color:#fff}
.nav-tabs{display:flex;gap:1px;flex:1;padding-left:4px}
.nav-tab{
  padding:0 16px;height:44px;display:flex;align-items:center;gap:6px;
  font-size:12px;font-weight:500;color:var(--c-muted);cursor:pointer;
  border-bottom:2px solid transparent;transition:all .15s;white-space:nowrap;
  background:none;border-top:none;border-left:none;border-right:none;
}
.nav-tab:hover{color:var(--c-text)}
.nav-tab.active{color:var(--c-accent);border-bottom-color:var(--c-accent)}
.nav-tab svg{width:13px;height:13px;fill:currentColor;opacity:.7}
.topbar-right{display:flex;align-items:center;gap:12px;margin-left:auto}
.clock-display{font-family:var(--font);font-size:12px;color:var(--c-muted);background:var(--c-bg);border:1px solid var(--c-border);padding:3px 10px;border-radius:4px}
.status-pill{font-size:10px;font-weight:700;padding:3px 8px;border-radius:3px;background:rgba(63,185,80,.12);color:var(--c-green);border:1px solid rgba(63,185,80,.25)}
.status-pill.off{background:rgba(248,81,73,.1);color:var(--c-red);border-color:rgba(248,81,73,.25)}
.btn-logout{background:none;border:1px solid var(--c-border);border-radius:4px;padding:4px 10px;color:var(--c-muted);font-size:11px;cursor:pointer;transition:all .15s}
.btn-logout:hover{border-color:var(--c-border2);color:var(--c-text)}

/* ─── MAIN LAYOUT ─────────────────── */
.main-content{flex:1;overflow:auto;padding:16px 20px}

/* ─── SECTION HEADER ─────────────── */
.sec-header{display:flex;justify-content:space-between;align-items:center;margin-bottom:14px}
.sec-title{font-size:13px;font-weight:600;color:#fff}
.sec-sub{font-size:11px;color:var(--c-muted);margin-top:1px}

/* ─── KPI STRIP ─────────────────────── */
.kpi-strip{display:grid;grid-template-columns:repeat(auto-fit,minmax(130px,1fr));gap:1px;background:var(--c-border);border:1px solid var(--c-border);border-radius:6px;overflow:hidden;margin-bottom:16px}
.kpi{background:var(--c-surface);padding:12px 14px}
.kpi-label{font-size:10px;font-weight:600;text-transform:uppercase;letter-spacing:.07em;color:var(--c-muted);margin-bottom:4px}
.kpi-val{font-size:22px;font-weight:700;color:#fff;font-family:var(--font);line-height:1}
.kpi-sub{font-size:10px;color:var(--c-muted);margin-top:3px}
.kpi-val.green{color:var(--c-green)}
.kpi-val.amber{color:var(--c-accent)}
.kpi-val.blue{color:var(--c-blue)}
.kpi-val.red{color:var(--c-red)}

/* ─── PROGRESS ─────────────────────── */
.prog-wrap{background:var(--c-surface);border:1px solid var(--c-border);border-radius:6px;padding:12px 14px;margin-bottom:16px}
.prog-header{display:flex;justify-content:space-between;font-size:11px;margin-bottom:7px}
.prog-bar{height:6px;background:var(--c-border);border-radius:3px;overflow:hidden}
.prog-fill{height:100%;background:var(--c-accent);border-radius:3px;transition:width .4s ease}

/* ─── GRID ─────────────────────── */
.grid-2{display:grid;grid-template-columns:1fr 1fr;gap:14px}
.grid-3{display:grid;grid-template-columns:2fr 1fr;gap:14px}
.col{display:flex;flex-direction:column;gap:14px}

/* ─── PANEL ─────────────────────── */
.panel{background:var(--c-surface);border:1px solid var(--c-border);border-radius:6px;overflow:hidden}
.panel-head{padding:10px 14px;border-bottom:1px solid var(--c-border);display:flex;align-items:center;justify-content:space-between;gap:8px}
.panel-title{font-size:11px;font-weight:600;text-transform:uppercase;letter-spacing:.07em;color:var(--c-muted);display:flex;align-items:center;gap:6px}
.panel-title .dot{width:6px;height:6px;border-radius:50%}
.dot-green{background:var(--c-green)}
.dot-amber{background:var(--c-accent)}
.dot-blue{background:var(--c-blue)}
.dot-red{background:var(--c-red)}
.dot-muted{background:var(--c-muted)}
.panel-body{padding:12px 14px}

/* ─── FORM ROW ─────────────────────── */
.form-row{display:flex;flex-wrap:wrap;gap:8px;align-items:flex-end}
.fld{display:flex;flex-direction:column;gap:4px}
.fld label{font-size:10px;font-weight:600;text-transform:uppercase;letter-spacing:.06em;color:var(--c-muted)}
.fld input,.fld select{
  background:var(--c-bg);border:1px solid var(--c-border);border-radius:5px;
  padding:6px 10px;color:var(--c-text);font-size:12px;outline:none;
  transition:border-color .15s;
}
.fld input:focus,.fld select:focus{border-color:var(--c-accent)}
.fld select option{background:var(--c-surface2)}

/* ─── DOCAS GRID ─────────────────────── */
.docas-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(120px,1fr));gap:6px;padding:12px 14px}
.doca{
  border:1px solid var(--c-border);border-radius:5px;padding:10px 8px;
  text-align:center;cursor:default;transition:border-color .15s,background .15s;
  min-height:110px;display:flex;flex-direction:column;align-items:center;justify-content:space-between;
}
.doca-num{font-family:var(--font);font-size:10px;font-weight:700;color:var(--c-muted);letter-spacing:.04em;margin-bottom:4px}
.doca-status{font-size:9px;font-weight:700;text-transform:uppercase;letter-spacing:.06em;padding:2px 6px;border-radius:3px;margin-bottom:6px}
.doca-info{font-size:11px;color:var(--c-text);font-weight:500;line-height:1.3;min-height:28px;display:flex;align-items:center;justify-content:center;text-align:center}
.doca-timer{font-family:var(--font);font-size:11px;font-weight:700;color:var(--c-accent);margin-bottom:4px}
.doca-actions{display:flex;flex-direction:column;gap:3px;width:100%}
.doca-actions .btn-sm{font-size:10px;padding:3px 8px;width:100%}

/* Doca states */
.doca.livre{background:var(--c-bg)}
.doca.chamado{background:rgba(88,166,255,.05);border-color:var(--c-blue)}
.doca.ocupada{background:rgba(63,185,80,.05);border-color:var(--c-green)}
.doca.carreta{background:rgba(248,81,73,.05);border-color:var(--c-red)}
.st-livre{background:rgba(139,148,158,.1);color:var(--c-muted)}
.st-chamado{background:rgba(88,166,255,.15);color:var(--c-blue)}
.st-ocupada{background:rgba(63,185,80,.15);color:var(--c-green)}
.st-carreta{background:rgba(248,81,73,.15);color:var(--c-red)}
@keyframes pulse-border{0%,100%{border-color:var(--c-blue)}50%{border-color:rgba(88,166,255,.3)}}
.doca.chamado{animation:pulse-border 1.4s infinite}

/* ─── FILA ─────────────────────── */
.fila-empresa{border-bottom:1px solid var(--c-border)}
.fila-empresa:last-child{border-bottom:none}
.fila-emp-head{padding:7px 14px;font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.06em;color:var(--c-muted);background:var(--c-bg);display:flex;justify-content:space-between;align-items:center}
.fila-emp-count{font-family:var(--font);background:var(--c-border);border-radius:3px;padding:1px 6px;color:var(--c-text);font-size:10px}
.fila-items{padding:4px 0}
.fila-item{padding:5px 14px;display:flex;align-items:center;gap:8px;font-size:11px;border-bottom:1px solid rgba(48,54,61,.5)}
.fila-item:last-child{border-bottom:none}
.fila-pos{width:18px;height:18px;border-radius:3px;background:var(--c-surface2);border:1px solid var(--c-border);font-family:var(--font);font-size:9px;font-weight:700;display:flex;align-items:center;justify-content:center;flex-shrink:0;color:var(--c-accent)}
.fila-nome{flex:1;color:var(--c-text)}
.fila-pkgs{font-family:var(--font);font-size:10px;color:var(--c-muted)}

/* ─── LOG ─────────────────────── */
.log-list{max-height:220px;overflow-y:auto;padding:8px 14px}
.log-entry{display:flex;gap:8px;align-items:flex-start;padding:4px 0;border-bottom:1px solid rgba(48,54,61,.5);font-size:11px}
.log-entry:last-child{border-bottom:none}
.log-t{font-family:var(--font);font-size:10px;color:var(--c-muted);white-space:nowrap;padding-top:1px;min-width:54px}
.log-dot2{width:5px;height:5px;border-radius:50%;flex-shrink:0;margin-top:4px}
.log-msg{color:var(--c-text)}

/* ─── CHART PANEL ─────────────────── */
.chart-wrap{padding:12px 14px}
.chart-wrap canvas{max-height:200px}

/* ─── TABLE ─────────────────────── */
.tbl-wrap{overflow-x:auto}
table{width:100%;border-collapse:collapse}
thead th{background:var(--c-bg);color:var(--c-muted);font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.06em;padding:8px 12px;text-align:left;border-bottom:1px solid var(--c-border);white-space:nowrap}
tbody td{padding:7px 12px;border-bottom:1px solid rgba(48,54,61,.5);font-size:12px;color:var(--c-text)}
tbody tr:last-child td{border-bottom:none}
tbody tr:hover td{background:rgba(255,255,255,.02)}
.td-mono{font-family:var(--font);font-size:11px}
.badge{display:inline-block;padding:2px 6px;border-radius:3px;font-size:10px;font-weight:700}
.badge-green{background:rgba(63,185,80,.15);color:var(--c-green);border:1px solid rgba(63,185,80,.25)}
.badge-blue{background:rgba(88,166,255,.15);color:var(--c-blue);border:1px solid rgba(88,166,255,.25)}

/* ─── MOTORISTA PORTAL ─────────────── */
#screen-motorista{background:var(--c-bg);align-items:center;justify-content:center;flex:1}
.motorista-wrap{width:100%;max-width:420px;padding:20px}
.m-card{background:var(--c-surface);border:1px solid var(--c-border);border-radius:8px;overflow:hidden}
.m-card-head{background:var(--c-bg);border-bottom:1px solid var(--c-border);padding:16px 20px;display:flex;align-items:center;gap:10px}
.m-card-head .icon{width:32px;height:32px;background:var(--c-accent);border-radius:5px;display:flex;align-items:center;justify-content:center;flex-shrink:0}
.m-card-head .icon svg{width:16px;height:16px;fill:#000}
.m-card-head h2{font-size:14px;font-weight:600;color:#fff}
.m-card-head p{font-size:11px;color:var(--c-muted)}
.m-card-body{padding:20px}
.status-block{text-align:center;padding:20px 0}
.big-num{font-family:var(--font);font-size:72px;font-weight:700;line-height:1;color:var(--c-accent)}
.big-doca{font-family:var(--font);font-size:48px;font-weight:700;color:var(--c-green);line-height:1}
.status-tag{display:inline-flex;align-items:center;gap:5px;padding:4px 12px;border-radius:3px;font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.05em;margin-top:8px}
.st-waiting{background:rgba(230,168,23,.1);color:var(--c-accent);border:1px solid rgba(230,168,23,.25)}
.st-called{background:rgba(88,166,255,.1);color:var(--c-blue);border:1px solid rgba(88,166,255,.25)}
.st-unload{background:rgba(63,185,80,.1);color:var(--c-green);border:1px solid rgba(63,185,80,.25)}
.st-done{background:rgba(63,185,80,.1);color:var(--c-green);border:1px solid rgba(63,185,80,.25)}
.m-timer{font-family:var(--font);font-size:28px;font-weight:700;color:var(--c-accent);text-align:center;margin:10px 0}
.m-btn-action{
  width:100%;padding:13px;border:none;border-radius:6px;font-size:14px;font-weight:700;
  cursor:pointer;transition:all .15s;display:flex;align-items:center;justify-content:center;gap:8px;margin-top:12px;
}
.m-btn-iniciar{background:var(--c-green);color:#000}
.m-btn-iniciar:hover{background:#3dd456}
.m-btn-finalizar{background:var(--c-red);color:#fff}
.m-btn-finalizar:hover{background:#ff6459}
.m-btn-novo{background:var(--c-surface2);color:var(--c-text);border:1px solid var(--c-border)}
.m-btn-novo:hover{border-color:var(--c-border2)}

/* ─── MODAL ─────────────────────── */
.modal-bg{position:fixed;inset:0;background:rgba(0,0,0,.7);z-index:500;display:none;align-items:center;justify-content:center}
.modal-bg.open{display:flex}
.modal{background:var(--c-surface);border:1px solid var(--c-border);border-radius:8px;padding:24px;width:340px}
.modal h3{font-size:14px;font-weight:600;color:#fff;margin-bottom:8px}
.modal p{font-size:12px;color:var(--c-muted);line-height:1.5;margin-bottom:16px}
.modal-btns{display:flex;gap:8px;justify-content:flex-end}

/* ─── EMPRESA TAG ─────────────────── */
.emp-tag{display:inline-block;padding:1px 6px;border-radius:3px;font-size:10px;font-weight:700;background:rgba(139,148,158,.1);color:var(--c-muted);border:1px solid var(--c-border)}

/* ─── EMPRESA ADD ─────────────────── */
.empresa-add-row{display:flex;gap:6px;margin-top:6px;align-items:center}
.empresa-add-row input{flex:1;background:var(--c-bg);border:1px solid var(--c-border);border-radius:5px;padding:5px 8px;color:var(--c-text);font-size:12px;outline:none}
.empresa-add-row input:focus{border-color:var(--c-accent)}

/* ─── LEGEND ─────────────────────── */
.legend{display:flex;flex-wrap:wrap;gap:10px;padding:10px 14px;border-top:1px solid var(--c-border)}
.legend-item{display:flex;align-items:center;gap:5px;font-size:10px;color:var(--c-muted)}
.legend-swatch{width:10px;height:10px;border-radius:2px}

/* ─── FILTER ROW ─────────────────── */
.filter-row{display:flex;flex-wrap:wrap;gap:8px;align-items:flex-end;padding:12px 14px;border-bottom:1px solid var(--c-border)}

/* responsive */
@media(max-width:900px){
  .grid-3{grid-template-columns:1fr}
  .grid-2{grid-template-columns:1fr}
}
@media(max-width:600px){
  .nav-tab span{display:none}
  .kpi-val{font-size:18px}
}
</style>
</head>
<body>

<!-- ════════════════════ LOGIN ════════════════════ -->
<div id="screen-login" class="screen active">
  <div class="login-card">
    <div class="brand">
      <div class="brand-icon">
        <svg viewBox="0 0 24 24"><path d="M3 8l7-5 4 3 4-3 3 2v12l-3 2-4-3-4 3-7-4V8z"/></svg>
      </div>
      <div>
        <div class="brand-name">DocaOps</div>
        <div class="brand-sub">Sistema de Gestão Operacional</div>
      </div>
    </div>
    <div class="field">
      <label>Usuário</label>
      <input type="text" id="loginUser" placeholder="gestor" autocomplete="off">
    </div>
    <div class="field">
      <label>Senha</label>
      <input type="password" id="loginPass" placeholder="••••••" onkeydown="if(event.key==='Enter')doLogin()">
    </div>
    <div id="loginErr" class="login-err">Usuário ou senha incorretos.</div>
    <button class="btn-primary" onclick="doLogin()">Entrar no Sistema</button>
    <div style="margin-top:20px;padding-top:16px;border-top:1px solid var(--c-border);font-size:11px;color:var(--c-muted)">
      <strong style="color:var(--c-text)">Acesso padrão:</strong><br>
      Usuário: <code style="color:var(--c-accent)">gestor</code> &nbsp;|&nbsp; Senha: <code style="color:var(--c-accent)">1234</code>
    </div>
  </div>
</div>

<!-- ════════════════════ APP PRINCIPAL ════════════════════ -->
<div id="screen-app" class="screen">

  <!-- TOPBAR -->
  <div class="topbar">
    <div class="topbar-brand">
      <div class="icon"><svg viewBox="0 0 24 24" fill="none" stroke="#000" stroke-width="2"><path d="M3 8l7-5 4 3 4-3 3 2v12l-3 2-4-3-4 3-7-4V8z"/></svg></div>
      <span>DocaOps</span>
    </div>
    <div class="nav-tabs">
      <button class="nav-tab active" onclick="switchTab('gestao',this)">
        <svg viewBox="0 0 16 16"><path d="M2 4h12M2 8h8M2 12h10"/></svg>
        <span>Doca em Tempo Real</span>
      </button>
      <button class="nav-tab" onclick="switchTab('dashboard',this)">
        <svg viewBox="0 0 16 16"><rect x="1" y="9" width="4" height="6"/><rect x="6" y="5" width="4" height="10"/><rect x="11" y="2" width="4" height="13"/></svg>
        <span>Dashboard & Analytics</span>
      </button>
    </div>
    <div class="topbar-right">
      <div id="opStatusPill" class="status-pill off">● OPERAÇÃO INATIVA</div>
      <div class="clock-display" id="clock">--:--:--</div>
      <button class="btn-logout" onclick="doLogout()">Sair</button>
    </div>
  </div>

  <!-- ── TAB: DOCA ────────────────── -->
  <div id="tab-gestao" class="main-content" style="display:flex;flex-direction:column;gap:14px">

    <!-- KPIs + controles topo -->
    <div style="display:flex;gap:10px;flex-wrap:wrap;align-items:center">
      <div style="flex:1">
        <div class="kpi-strip">
          <div class="kpi">
            <div class="kpi-label">Meta</div>
            <div style="display:flex;align-items:center;gap:6px;margin-top:2px">
              <input type="number" id="kpiMetaInput" value="5000" style="width:80px;background:var(--c-bg);border:1px solid var(--c-border);border-radius:4px;padding:2px 6px;color:#fff;font-size:18px;font-weight:700;font-family:var(--font);outline:none">
              <button class="btn-sm" onclick="G.definirMeta()">✓</button>
            </div>
          </div>
          <div class="kpi">
            <div class="kpi-label">Descarregados</div>
            <div class="kpi-val green" id="kpiDesc">0</div>
            <div class="kpi-sub">pacotes</div>
          </div>
          <div class="kpi">
            <div class="kpi-label">Na Fila</div>
            <div class="kpi-val amber" id="kpiFila">0</div>
            <div class="kpi-sub">motoristas</div>
          </div>
          <div class="kpi">
            <div class="kpi-label">Docas Ativas</div>
            <div class="kpi-val blue" id="kpiDocas">0</div>
            <div class="kpi-sub">de 13</div>
          </div>
          <div class="kpi">
            <div class="kpi-label">Progresso</div>
            <div class="kpi-val" id="kpiPct" style="font-size:18px">0%</div>
            <div class="prog-bar" style="margin-top:5px"><div class="prog-fill" id="kpiBar" style="width:0%"></div></div>
          </div>
        </div>
      </div>
      <button id="btnIniciarOp" onclick="G.iniciarOperacao()" class="btn-sm btn-green" style="padding:9px 16px;font-size:12px;white-space:nowrap">
        ▶ Iniciar Operação
      </button>
    </div>

    <!-- Add Motorista -->
    <div class="panel">
      <div class="panel-head">
        <div class="panel-title"><span class="dot dot-blue"></span>Adicionar Motorista à Fila</div>
      </div>
      <div class="panel-body">
        <div class="form-row">
          <div class="fld" style="min-width:140px;flex:1">
            <label>Nome</label>
            <input type="text" id="gNome" placeholder="Nome do motorista">
          </div>
          <div class="fld" style="min-width:100px">
            <label>Empresa</label>
            <select id="gEmpresa"></select>
          </div>
          <div class="fld" style="width:90px">
            <label>Pacotes</label>
            <input type="number" id="gPacotes" placeholder="0">
          </div>
          <div class="fld" style="align-self:flex-end">
            <button class="btn-sm btn-amber" onclick="G.adicionarMotorista()">+ Adicionar</button>
          </div>
        </div>
      </div>
    </div>

    <!-- Grid principal -->
    <div class="grid-3">
      <!-- Coluna esquerda: Docas + Fila -->
      <div class="col">
        <!-- Docas -->
        <div class="panel">
          <div class="panel-head">
            <div class="panel-title"><span class="dot dot-amber"></span>Docas de Desembarque</div>
            <div style="display:flex;gap:10px;font-size:10px;color:var(--c-muted)">
              <span><span class="dot dot-blue" style="display:inline-block;width:6px;height:6px;border-radius:50%;margin-right:3px"></span>Chamado</span>
              <span><span class="dot dot-green" style="display:inline-block;width:6px;height:6px;border-radius:50%;margin-right:3px"></span>Descarregando</span>
              <span><span class="dot dot-red" style="display:inline-block;width:6px;height:6px;border-radius:50%;margin-right:3px"></span>Carreta</span>
            </div>
          </div>
          <div class="docas-grid" id="docasContainer"></div>
        </div>

        <!-- Fila -->
        <div class="panel">
          <div class="panel-head">
            <div class="panel-title"><span class="dot dot-muted"></span>Fila de Espera por Empresa</div>
          </div>
          <div id="filaContainer"></div>
        </div>
      </div>

      <!-- Coluna direita: Gráfico + Log -->
      <div class="col">
        <div class="panel">
          <div class="panel-head"><div class="panel-title"><span class="dot dot-green"></span>Entradas por Hora</div></div>
          <div class="chart-wrap"><canvas id="horariosChart"></canvas></div>
        </div>
        <div class="panel">
          <div class="panel-head"><div class="panel-title"><span class="dot dot-muted"></span>Log de Eventos</div></div>
          <div class="log-list" id="logContainer"></div>
        </div>
      </div>
    </div>
  </div>

  <!-- ── TAB: DASHBOARD ─────────────── -->
  <div id="tab-dashboard" class="main-content" style="display:none;flex-direction:column;gap:14px">

    <!-- Filtros + Add Operação -->
    <div class="panel">
      <div class="panel-head"><div class="panel-title"><span class="dot dot-amber"></span>Registrar Operação de Dupla</div></div>
      <div class="filter-row">
        <div class="fld" style="flex:1;min-width:140px">
          <label>Dupla</label>
          <input type="text" id="dTeam" list="teamsList" placeholder="Nome da dupla">
          <datalist id="teamsList"></datalist>
        </div>
        <div class="fld">
          <label>Início</label>
          <input type="datetime-local" id="dStart">
        </div>
        <div class="fld">
          <label>Fim</label>
          <input type="datetime-local" id="dEnd">
        </div>
        <div class="fld" style="width:90px">
          <label>Pacotes</label>
          <input type="number" id="dPkgs" placeholder="0">
        </div>
        <div class="fld" style="align-self:flex-end">
          <button class="btn-sm btn-amber" onclick="D.adicionarOp()">+ Registrar</button>
        </div>
        <div class="fld" style="margin-left:auto;align-self:flex-end;display:flex;gap:6px">
          <select id="dFilter" onchange="D.applyFilter(this.value)" class="fld select">
            <option value="all">Todos os períodos</option>
            <option value="today" selected>Hoje</option>
            <option value="yesterday">Ontem</option>
            <option value="week">Esta semana</option>
            <option value="month">Este mês</option>
          </select>
          <button class="btn-sm" onclick="D.exportar()">↓ Excel</button>
          <button class="btn-sm btn-red" onclick="D.limpar()">⊘ Limpar</button>
        </div>
      </div>
    </div>

    <!-- KPIs Dashboard -->
    <div class="kpi-strip">
      <div class="kpi"><div class="kpi-label">Total Pacotes</div><div class="kpi-val amber" id="dTotal">0</div></div>
      <div class="kpi"><div class="kpi-label">Caminhões</div><div class="kpi-val blue" id="dTrucks">0</div></div>
      <div class="kpi"><div class="kpi-label">Dupla Mais Rápida</div><div class="kpi-val green" id="dFastest" style="font-size:14px;margin-top:2px">—</div></div>
      <div class="kpi"><div class="kpi-label">Dupla Mais Lenta</div><div class="kpi-val red" id="dSlowest" style="font-size:14px;margin-top:2px">—</div></div>
      <div class="kpi"><div class="kpi-label">Média Eficiência</div><div class="kpi-val" id="dAvgEff" style="font-size:16px;margin-top:2px">0 pac/h</div></div>
    </div>

    <!-- Charts grid -->
    <div class="grid-2">
      <div class="panel"><div class="panel-head"><div class="panel-title"><span class="dot dot-amber"></span>Volume por Dupla</div></div><div class="chart-wrap"><canvas id="volChart"></canvas></div></div>
      <div class="panel"><div class="panel-head"><div class="panel-title"><span class="dot dot-blue"></span>Eficiência (pac/h)</div></div><div class="chart-wrap"><canvas id="effChart"></canvas></div></div>
      <div class="panel"><div class="panel-head"><div class="panel-title"><span class="dot dot-green"></span>Caminhões por Dupla</div></div><div class="chart-wrap"><canvas id="truckChart"></canvas></div></div>
      <div class="panel"><div class="panel-head"><div class="panel-title"><span class="dot dot-muted"></span>Distribuição de Volume</div></div><div class="chart-wrap"><canvas id="distChart"></canvas></div></div>
    </div>
    <div class="panel">
      <div class="panel-head"><div class="panel-title"><span class="dot dot-amber"></span>Desempenho Acumulado por Hora</div></div>
      <div class="chart-wrap" style="padding:12px 14px"><canvas id="lineChart" style="max-height:220px"></canvas></div>
      <div class="legend" id="chartLegend"></div>
    </div>

    <!-- Tabela -->
    <div class="panel">
      <div class="panel-head"><div class="panel-title"><span class="dot dot-muted"></span>Detalhamento das Operações</div></div>
      <div class="tbl-wrap">
        <table id="opsTable">
          <thead><tr><th>Dupla</th><th>Início</th><th>Término</th><th>Pacotes</th><th>Caminhões</th><th>Duração (h)</th><th>Eficiência</th></tr></thead>
          <tbody></tbody>
        </table>
      </div>
    </div>
  </div>

</div><!-- /screen-app -->

<!-- PORTAL MOTORISTA (tela separada) -->
<div id="screen-motorista" class="screen">
  <div class="motorista-wrap">
    <div class="m-card">
      <div class="m-card-head">
        <div class="icon"><svg viewBox="0 0 24 24" fill="none" stroke="#000" stroke-width="2"><path d="M20 7H4a1 1 0 00-1 1v10a1 1 0 001 1h16a1 1 0 001-1V8a1 1 0 00-1-1z"/><path d="M16 7V5a2 2 0 00-2-2h-4a2 2 0 00-2 2v2"/></svg></div>
        <div>
          <h2>Portal do Motorista</h2>
          <p>Registro de chegada e acompanhamento</p>
        </div>
      </div>
      <div class="m-card-body">
        <!-- Registro -->
        <div id="mFormView">
          <div style="display:flex;flex-direction:column;gap:10px">
            <div class="fld">
              <label>Seu Nome Completo</label>
              <input type="text" id="mNome" placeholder="Nome do motorista" style="font-size:13px;padding:8px 10px">
            </div>
            <div class="fld">
              <label>WhatsApp (com DDD)</label>
              <input type="tel" id="mWpp" placeholder="11999998888" style="font-size:13px;padding:8px 10px">
            </div>
            <div class="fld">
              <label>Empresa</label>
              <select id="mEmpresa" style="font-size:13px;padding:8px 10px"></select>
              <div class="empresa-add-row">
                <input type="text" id="mNovaEmpresa" placeholder="Outra empresa? Digite aqui...">
                <button class="btn-sm btn-amber" onclick="M.addEmpresa()" style="white-space:nowrap">+ Adicionar</button>
              </div>
            </div>
            <div class="fld">
              <label>Volume de Pacotes</label>
              <input type="number" id="mPkgs" placeholder="Nº de pacotes" style="font-size:13px;padding:8px 10px">
            </div>
            <button class="m-btn-action m-btn-iniciar" onclick="M.registrar()">
              <svg width="16" height="16" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><path d="M5 12l5 5L20 7"/></svg>
              ENTRAR NA FILA
            </button>
            <button class="m-btn-action m-btn-novo" onclick="showLogin()" style="margin-top:4px;font-size:12px;padding:9px">← Voltar para o login</button>
          </div>
        </div>
        <!-- Status -->
        <div id="mStatusView" style="display:none">
          <div id="mStatusFila" class="status-block">
            <div style="font-size:12px;color:var(--c-muted);margin-bottom:4px">Sua posição na fila</div>
            <div class="big-num" id="mPosNum">—</div>
            <div class="status-tag st-waiting">● Aguardando chamada</div>
            <div style="font-size:11px;color:var(--c-muted);margin-top:12px">Fique atento. Você será chamado em breve.</div>
          </div>
          <div id="mStatusChamado" class="status-block" style="display:none">
            <div style="font-size:12px;color:var(--c-blue);font-weight:600;margin-bottom:8px">▲ SUA VEZ! Dirija-se à:</div>
            <div class="big-doca" id="mDocaNum">—</div>
            <div style="font-size:11px;color:var(--c-muted);margin-top:4px">Doca de desembarque</div>
            <button class="m-btn-action m-btn-iniciar" onclick="M.iniciarDescarga()" style="margin-top:16px">▶ Iniciar Descarga</button>
          </div>
          <div id="mStatusDesc" style="display:none">
            <div style="text-align:center;padding:16px 0 8px">
              <div style="font-size:11px;color:var(--c-muted);margin-bottom:4px">Descarga em andamento</div>
              <div class="m-timer" id="mTimer">00:00:00</div>
              <div class="status-tag st-unload">● Em descarga</div>
            </div>
            <button class="m-btn-action m-btn-finalizar" onclick="M.finalizar()">■ Finalizar Descarga</button>
          </div>
          <div id="mStatusFim" style="display:none">
            <div class="status-block">
              <div style="font-size:32px;margin-bottom:8px">✓</div>
              <div style="font-size:16px;font-weight:600;color:var(--c-green)">Descarga Concluída</div>
              <div style="font-size:11px;color:var(--c-muted);margin-top:6px">Obrigado pelo serviço.</div>
            </div>
            <button class="m-btn-action m-btn-novo" onclick="M.novo()">← Novo Registro</button>
          </div>
          <div style="margin-top:12px;padding-top:12px;border-top:1px solid var(--c-border);font-size:11px;color:var(--c-muted);text-align:center" id="mNomeDisplay"></div>
        </div>
      </div>
    </div>
    <div style="text-align:center;margin-top:12px;font-size:11px;color:var(--c-muted)">
      Gestor? <a href="#" onclick="showLogin()" style="color:var(--c-accent)">Acessar painel de gestão</a>
    </div>
  </div>
</div>

<!-- MODAL CONFIRMAÇÃO -->
<div class="modal-bg" id="modalConfirm">
  <div class="modal">
    <h3>Confirmar exclusão</h3>
    <p>Todos os dados das operações serão removidos permanentemente. Esta ação não pode ser desfeita.</p>
    <div class="modal-btns">
      <button class="btn-sm" onclick="closeModal()">Cancelar</button>
      <button class="btn-sm btn-red" onclick="D.confirmarLimpar()">Limpar tudo</button>
    </div>
  </div>
</div>

<script>
// ═══════════════════════════════════════════════
//  ESTADO
// ═══════════════════════════════════════════════
const S = {
  fila:[],
  docas: (() => {
    // Docas 169 a 181 = 13 docas
    const ids = Array.from({length:13},(_,i)=>169+i);
    return ids.map(id=>({id,status:'livre',motorista:null}));
  })(),
  meta:5000, iniciada:false,
  historico:{}, descarregados:0,
  log:[],
  ops:[], filteredOps:[],
  teamColors:{}, colorIdx:0,
  motorista:null,
  empresas:['Prálog','Imediato','Hawk','Ontime'],
};

const PALETA = ['#e6a817','#58a6ff','#3fb950','#f85149','#bc8cff','#ffa657','#39d353','#ff7b72','#d2a8ff','#79c0ff'];

// ═══════════════════════════════════════════════
//  PERSISTÊNCIA
// ═══════════════════════════════════════════════
function save(){
  try{
    localStorage.setItem('doc_fila',JSON.stringify(S.fila));
    localStorage.setItem('doc_docas',JSON.stringify(S.docas));
    localStorage.setItem('doc_meta',S.meta);
    localStorage.setItem('doc_iniciada',S.iniciada);
    localStorage.setItem('doc_hist',JSON.stringify(S.historico));
    localStorage.setItem('doc_desc',S.descarregados);
    localStorage.setItem('doc_log',JSON.stringify(S.log.slice(0,60)));
    localStorage.setItem('doc_ops',JSON.stringify(S.ops));
    localStorage.setItem('doc_colors',JSON.stringify(S.teamColors));
    localStorage.setItem('doc_motorista',JSON.stringify(S.motorista));
    localStorage.setItem('doc_empresas',JSON.stringify(S.empresas));
  }catch(e){}
}

function loadState(){
  try{
    S.fila=JSON.parse(localStorage.getItem('doc_fila')||'[]');
    const sd=JSON.parse(localStorage.getItem('doc_docas')||'null');
    if(sd)S.docas=sd;
    S.meta=parseInt(localStorage.getItem('doc_meta')||'5000');
    S.iniciada=JSON.parse(localStorage.getItem('doc_iniciada')||'false');
    S.historico=JSON.parse(localStorage.getItem('doc_hist')||'{}');
    S.descarregados=parseInt(localStorage.getItem('doc_desc')||'0');
    S.log=JSON.parse(localStorage.getItem('doc_log')||'[]');
    S.ops=JSON.parse(localStorage.getItem('doc_ops')||'null')||genSampleOps();
    S.teamColors=JSON.parse(localStorage.getItem('doc_colors')||'{}');
    S.colorIdx=Object.keys(S.teamColors).length;
    S.motorista=JSON.parse(localStorage.getItem('doc_motorista')||'null');
    const se=JSON.parse(localStorage.getItem('doc_empresas')||'null');
    if(se&&se.length)S.empresas=se;
  }catch(e){}
}

function genSampleOps(){
  const names=["Carlos & Bruna","Lucas & Fernanda","Matheus & Gabriela","Rafael & Juliana","Vinicius & Patricia","Gustavo & Sandra","Thiago & Amanda","Felipe & Renata","Ricardo & Camila","Diego & Beatriz"];
  const today=new Date().toISOString().slice(0,10);
  return names.map((name,i)=>{
    const h=13+(i%5),dur=60+Math.floor(Math.random()*90);
    const st=new Date(`${today}T${String(h).padStart(2,'0')}:${String((i*7)%60).padStart(2,'0')}`);
    const en=new Date(st.getTime()+dur*60000);
    return{id:i+1,team:name,startTime:st.toISOString(),endTime:en.toISOString(),packages:1500+Math.floor(Math.random()*800)};
  });
}

// ═══════════════════════════════════════════════
//  HELPERS
// ═══════════════════════════════════════════════
function addLog(tipo,msg){
  const h=new Date().toLocaleTimeString('pt-BR',{hour:'2-digit',minute:'2-digit',second:'2-digit'});
  S.log.unshift({tipo,msg,h});
  if(S.log.length>60)S.log.length=60;
}
function getColor(t){
  if(!S.teamColors[t]){S.teamColors[t]=PALETA[S.colorIdx%PALETA.length];S.colorIdx++;}
  return S.teamColors[t];
}
function getInitials(n){
  if(!n)return'??';
  const p=n.replace(/&/g,'e').split(/\s+e\s+/i).map(x=>x.trim());
  if(p.length>=2&&p[0]&&p[1])return(p[0][0]+p[1][0]).toUpperCase();
  const w=n.split(' ');
  if(w.length>=2)return(w[0][0]+w[w.length-1][0]).toUpperCase();
  return n.slice(0,2).toUpperCase();
}
function fmtDt(d){return new Date(d).toLocaleString('pt-BR');}
function closeModal(){document.getElementById('modalConfirm').classList.remove('open');}

function populateEmpresaSelects(){
  ['gEmpresa','mEmpresa'].forEach(id=>{
    const el=document.getElementById(id);
    if(!el)return;
    const cur=el.value;
    el.innerHTML=S.empresas.map(e=>`<option value="${e}">${e}</option>`).join('');
    if(cur&&S.empresas.includes(cur))el.value=cur;
  });
}

// ═══════════════════════════════════════════════
//  LOGIN / NAVEGAÇÃO
// ═══════════════════════════════════════════════
function showScreen(id){
  document.querySelectorAll('.screen').forEach(s=>s.classList.remove('active'));
  document.getElementById(id).classList.add('active');
}
function doLogin(){
  const u=document.getElementById('loginUser').value.trim();
  const p=document.getElementById('loginPass').value;
  const err=document.getElementById('loginErr');
  if(u==='gestor'&&p==='1234'){
    err.style.display='none';
    showScreen('screen-app');
    G.render();
    document.getElementById('dFilter').value='today';
    D.applyFilter('today');
  } else {
    err.style.display='block';
  }
}
function doLogout(){
  document.getElementById('loginUser').value='';
  document.getElementById('loginPass').value='';
  showScreen('screen-login');
}
function showLogin(){showScreen('screen-login');}

// Link para portal motorista — não precisa de login
window.abrirPortalMotorista=function(){
  showScreen('screen-motorista');
  M.render();
};
// Adicionar link no rodapé login
document.addEventListener('DOMContentLoaded',()=>{
  const lc=document.querySelector('.login-card');
  const link=document.createElement('div');
  link.style.cssText='margin-top:12px;text-align:center;font-size:11px;color:var(--c-muted)';
  link.innerHTML='Motorista? <a href="#" onclick="abrirPortalMotorista()" style="color:var(--c-accent)">Acessar portal do motorista</a>';
  lc.appendChild(link);
});

// ═══════════════════════════════════════════════
//  TABS
// ═══════════════════════════════════════════════
function switchTab(id,btn){
  document.querySelectorAll('.nav-tab').forEach(b=>b.classList.remove('active'));
  btn.classList.add('active');
  document.getElementById('tab-gestao').style.display='none';
  document.getElementById('tab-dashboard').style.display='none';
  document.getElementById('tab-'+id).style.display='flex';
  document.getElementById('tab-'+id).style.flexDirection='column';
  if(id==='dashboard')D.applyFilter(document.getElementById('dFilter').value);
}

// ═══════════════════════════════════════════════
//  CHARTS
// ═══════════════════════════════════════════════
let CH={};
const CHART_DEFAULTS={
  responsive:true,maintainAspectRatio:true,
  plugins:{legend:{display:false}},
  scales:{
    x:{ticks:{color:'#8b949e',font:{size:10}},grid:{color:'rgba(255,255,255,0.04)'}},
    y:{ticks:{color:'#8b949e',font:{size:10}},grid:{color:'rgba(255,255,255,0.04)'},beginAtZero:true}
  }
};
function initCharts(){
  CH.horarios=new Chart(document.getElementById('horariosChart'),{
    type:'bar',
    data:{labels:Array.from({length:24},(_,i)=>`${i}h`),datasets:[{label:'Caminhões',data:Array(24).fill(0),backgroundColor:'#e6a81799',borderColor:'#e6a817',borderWidth:1,borderRadius:2}]},
    options:{...CHART_DEFAULTS}
  });
  function mkBar(id){
    return new Chart(document.getElementById(id),{type:'bar',data:{labels:[],datasets:[{data:[],backgroundColor:[],borderRadius:2}]},options:{...CHART_DEFAULTS}});
  }
  CH.vol=mkBar('volChart');CH.eff=mkBar('effChart');CH.truck=mkBar('truckChart');
  CH.dist=new Chart(document.getElementById('distChart'),{
    type:'doughnut',data:{labels:[],datasets:[{data:[],backgroundColor:[]}]},
    options:{responsive:true,plugins:{legend:{position:'right',labels:{color:'#8b949e',font:{size:10},boxWidth:10,padding:8}}}}
  });
  CH.line=new Chart(document.getElementById('lineChart'),{
    type:'line',data:{labels:[],datasets:[]},
    options:{responsive:true,maintainAspectRatio:true,plugins:{legend:{labels:{color:'#8b949e',font:{size:10},boxWidth:10,padding:8}}},
    scales:{x:{ticks:{color:'#8b949e',font:{size:10}},grid:{color:'rgba(255,255,255,0.04)'}},y:{ticks:{color:'#8b949e'},grid:{color:'rgba(255,255,255,0.04)'},beginAtZero:true}},
    elements:{line:{tension:.3,borderWidth:1.5},point:{radius:2}}}
  });
}

// ═══════════════════════════════════════════════
//  G — GESTÃO
// ═══════════════════════════════════════════════
const G={
  iniciarOperacao(){
    if(S.iniciada)return;
    S.iniciada=true;
    addLog('sys','Operação iniciada.');
    const b=document.getElementById('btnIniciarOp');
    b.textContent='● Operação Ativa';b.disabled=true;b.className='btn-sm';b.style.opacity='.5';
    document.getElementById('opStatusPill').textContent='● OPERAÇÃO ATIVA';
    document.getElementById('opStatusPill').classList.remove('off');
    this.chamarProximo();this.render();save();
  },
  definirMeta(){
    const v=parseInt(document.getElementById('kpiMetaInput').value);
    if(v>0){S.meta=v;addLog('sys',`Meta: ${v.toLocaleString()} pacotes.`);save();this.render();}
  },
  adicionarMotorista(){
    const nome=document.getElementById('gNome').value.trim();
    const empresa=document.getElementById('gEmpresa').value;
    const pkgs=parseInt(document.getElementById('gPacotes').value);
    if(!nome||!pkgs||pkgs<=0){alert('Preencha nome e pacotes.');return;}
    const m={id:Date.now(),nome,empresa,pacotes:pkgs};
    S.fila.push(m);
    addLog('reg',`${nome} (${empresa}) — ${pkgs.toLocaleString()} pacotes.`);
    document.getElementById('gNome').value='';document.getElementById('gPacotes').value='';
    save();this.chamarProximo();this.render();
  },
  chamarProximo(){
    if(!S.iniciada||!S.fila.length)return;
    const doca=S.docas.find(d=>d.status==='livre');
    if(!doca)return;
    const m=S.fila.shift();
    doca.status='chamado';doca.motorista=m;
    addLog('call',`${m.nome} → Doca ${doca.id}`);
    // sync motorista portal
    if(S.motorista&&S.motorista.id===m.id){S.motorista.status='chamado';S.motorista.docaId=doca.id;save();M.render();}
    save();this.render();
  },
  iniciarDoca(id){
    const d=S.docas.find(x=>x.id===id);
    if(!d||d.status!=='chamado')return;
    d.status='ocupada';d.motorista.horaInicio=new Date().toISOString();
    addLog('sys',`Descarga iniciada — Doca ${id}.`);
    if(S.motorista&&S.motorista.docaId===id){S.motorista.status='desc';S.motorista.horaInicio=d.motorista.horaInicio;}
    save();this.render();
  },
  finalizarDoca(id){
    const d=S.docas.find(x=>x.id===id);
    if(!d||d.status!=='ocupada')return;
    const m=d.motorista;
    const hr=new Date(m.horaInicio).getHours();
    S.historico[hr]=(S.historico[hr]||0)+1;
    S.descarregados+=m.pacotes;
    addLog('done',`${m.nome} — Doca ${id} — +${m.pacotes.toLocaleString()} pac.`);
    if(S.motorista&&S.motorista.id===m.id)S.motorista.status='fim';
    d.status='livre';d.motorista=null;
    save();this.chamarProximo();this.render();
  },
  ocuparCarreta(id){
    const d=S.docas.find(x=>x.id===id);
    if(!d||d.status!=='livre')return;
    d.status='carreta_manual';
    addLog('sys',`Doca ${id} — carreta.`);
    save();this.render();
  },
  finalizarCarreta(id){
    const d=S.docas.find(x=>x.id===id);
    if(!d||d.status!=='carreta_manual')return;
    d.status='livre';d.motorista=null;
    addLog('done',`Carreta finalizada — Doca ${id}.`);
    save();this.chamarProximo();this.render();
  },
  render(){
    // KPIs
    const pct=Math.min(100,Math.round(S.descarregados/S.meta*100));
    document.getElementById('kpiDesc').textContent=S.descarregados.toLocaleString('pt-BR');
    document.getElementById('kpiFila').textContent=S.fila.length;
    document.getElementById('kpiDocas').textContent=S.docas.filter(d=>d.status!=='livre').length;
    document.getElementById('kpiPct').textContent=pct+'%';
    document.getElementById('kpiBar').style.width=pct+'%';
    if(S.iniciada){
      const b=document.getElementById('btnIniciarOp');
      b.textContent='● Operação Ativa';b.disabled=true;b.style.opacity='.5';
      document.getElementById('opStatusPill').textContent='● OPERAÇÃO ATIVA';
      document.getElementById('opStatusPill').classList.remove('off');
    }
    // Docas
    const dc=document.getElementById('docasContainer');
    dc.innerHTML='';
    S.docas.forEach(d=>{
      const el=document.createElement('div');
      el.className='doca '+d.status.replace('_manual','').replace('carreta_manual','carreta');
      const num=`<div class="doca-num">DOCA ${d.id}</div>`;
      if(d.status==='livre'){
        el.innerHTML=`${num}<div class="doca-status st-livre">Livre</div><div class="doca-info">—</div><div class="doca-timer"></div><div class="doca-actions"><button class="btn-sm" onclick="G.ocuparCarreta(${d.id})" style="font-size:9px">+ Carreta</button></div>`;
      } else if(d.status==='carreta_manual'){
        el.innerHTML=`${num}<div class="doca-status st-carreta">Carreta</div><div class="doca-info">Manual</div><div class="doca-timer"></div><div class="doca-actions"><button class="btn-sm btn-red" onclick="G.finalizarCarreta(${d.id})">Finalizar</button></div>`;
      } else if(d.status==='chamado'){
        el.innerHTML=`${num}<div class="doca-status st-chamado">Chamado</div><div class="doca-info">${d.motorista.nome}</div><div class="doca-timer"></div><div class="doca-actions"><button class="btn-sm btn-green" onclick="G.iniciarDoca(${d.id})">Iniciar</button></div>`;
      } else {
        const el2=Date.now()-new Date(d.motorista.horaInicio).getTime();
        const h=String(Math.floor(el2/3600000)).padStart(2,'0');
        const m2=String(Math.floor((el2%3600000)/60000)).padStart(2,'0');
        const s=String(Math.floor((el2%60000)/1000)).padStart(2,'0');
        el.innerHTML=`${num}<div class="doca-status st-ocupada">Ativa</div><div class="doca-info">${d.motorista.nome}</div><div class="doca-timer">${h}:${m2}:${s}</div><div class="doca-actions"><button class="btn-sm btn-red" onclick="G.finalizarDoca(${d.id})">Finalizar</button></div>`;
      }
      dc.appendChild(el);
    });
    // Fila
    const fc=document.getElementById('filaContainer');
    fc.innerHTML='';
    if(!S.fila.length){fc.innerHTML='<div style="padding:14px;font-size:11px;color:var(--c-muted);text-align:center">Nenhum motorista na fila.</div>';return;}
    S.empresas.forEach(emp=>{
      const items=S.fila.filter(x=>x.empresa===emp);
      const sec=document.createElement('div');sec.className='fila-empresa';
      const rows=items.length?items.map(x=>{
        const pos=S.fila.findIndex(f=>f.id===x.id)+1;
        return `<div class="fila-item"><div class="fila-pos">${pos}</div><div class="fila-nome">${x.nome}</div><div class="fila-pkgs">${x.pacotes.toLocaleString()}</div></div>`;
      }).join(''):`<div class="fila-item" style="color:var(--c-muted);font-size:10px;padding:6px 14px">Vazia</div>`;
      sec.innerHTML=`<div class="fila-emp-head">${emp}<span class="fila-emp-count">${items.length}</span></div><div class="fila-items">${rows}</div>`;
      fc.appendChild(sec);
    });
    // Log
    const lc=document.getElementById('logContainer');
    const colors={sys:'#58a6ff',reg:'#3fb950',call:'#e6a817',done:'#bc8cff',err:'#f85149'};
    lc.innerHTML=S.log.length?S.log.map(e=>`<div class="log-entry"><div class="log-t">${e.h}</div><div class="log-dot2" style="background:${colors[e.tipo]||'#8b949e'}"></div><div class="log-msg">${e.msg}</div></div>`).join(''):'<div style="font-size:11px;color:var(--c-muted)">Aguardando eventos...</div>';
    // Chart
    CH.horarios.data.datasets[0].data=Array.from({length:24},(_,i)=>S.historico[i]||0);
    CH.horarios.update();
  }
};

// ═══════════════════════════════════════════════
//  D — DASHBOARD
// ═══════════════════════════════════════════════
const D={
  applyFilter(val){
    const now=new Date(),today=new Date(now.getFullYear(),now.getMonth(),now.getDate());
    let st,en;
    switch(val){
      case 'today':st=today;en=new Date(today.getTime()+864e5);break;
      case 'yesterday':st=new Date(today.getTime()-864e5);en=today;break;
      case 'week':st=new Date(today);st.setDate(st.getDate()-now.getDay());en=new Date(now.getTime()+864e5);break;
      case 'month':st=new Date(now.getFullYear(),now.getMonth(),1);en=new Date(now.getTime()+864e5);break;
      default:S.filteredOps=[...S.ops];this.render();return;
    }
    S.filteredOps=S.ops.filter(op=>{const d=new Date(op.startTime);return d>=st&&d<en;});
    this.render();
  },
  adicionarOp(){
    const team=document.getElementById('dTeam').value.trim();
    const s=document.getElementById('dStart').value;
    const e=document.getElementById('dEnd').value;
    const p=parseInt(document.getElementById('dPkgs').value);
    if(!team||!s||!e||!p){alert('Preencha todos os campos.');return;}
    if(new Date(e)<=new Date(s)){alert('Término deve ser após o início.');return;}
    S.ops.push({id:Date.now(),team,startTime:new Date(s).toISOString(),endTime:new Date(e).toISOString(),packages:p});
    getColor(team);
    document.getElementById('dTeam').value='';document.getElementById('dPkgs').value='';
    save();this.applyFilter(document.getElementById('dFilter').value);this.populateDL();
  },
  populateDL(){
    const dl=document.getElementById('teamsList');
    dl.innerHTML=[...new Set(S.ops.map(o=>o.team))].map(t=>`<option value="${t}">`).join('');
  },
  getMetrics(){
    const teams=[...new Set(S.filteredOps.map(o=>o.team))];
    const m={};
    teams.forEach(t=>{
      const ops=S.filteredOps.filter(o=>o.team===t);
      const pkgs=ops.reduce((s,o)=>s+o.packages,0);
      const hrs=ops.reduce((s,o)=>s+(new Date(o.endTime)-new Date(o.startTime))/36e5,0);
      m[t]={pkgs,trucks:ops.length,hrs,eff:hrs>0?pkgs/hrs:0};
    });
    return m;
  },
  render(){
    const m=this.getMetrics();
    const teams=Object.keys(m);
    const inits=teams.map(t=>getInitials(t));
    const colors=teams.map(t=>getColor(t));
    // KPIs
    const total=Object.values(m).reduce((s,x)=>s+x.pkgs,0);
    const trucks=Object.values(m).reduce((s,x)=>s+x.trucks,0);
    const avgEff=teams.length?Object.values(m).reduce((s,x)=>s+x.eff,0)/teams.length:0;
    document.getElementById('dTotal').textContent=total.toLocaleString('pt-BR');
    document.getElementById('dTrucks').textContent=trucks;
    document.getElementById('dAvgEff').textContent=avgEff.toFixed(0)+' pac/h';
    const sorted=Object.entries(m).sort((a,b)=>b[1].eff-a[1].eff);
    if(sorted.length>=2){
      document.getElementById('dFastest').textContent=`▲ ${getInitials(sorted[0][0])} ${sorted[0][1].eff.toFixed(0)} p/h`;
      document.getElementById('dSlowest').textContent=`▼ ${getInitials(sorted[sorted.length-1][0])} ${sorted[sorted.length-1][1].eff.toFixed(0)} p/h`;
    } else {document.getElementById('dFastest').textContent='—';document.getElementById('dSlowest').textContent='—';}
    // Bar charts
    function updBar(ch,labels,data,bgs){ch.data.labels=labels;ch.data.datasets[0].data=data;ch.data.datasets[0].backgroundColor=bgs;ch.update();}
    updBar(CH.vol,inits,teams.map(t=>m[t].pkgs),colors);
    updBar(CH.eff,inits,teams.map(t=>parseFloat(m[t].eff.toFixed(1))),colors);
    updBar(CH.truck,inits,teams.map(t=>m[t].trucks),colors);
    // Doughnut
    CH.dist.data.labels=inits;CH.dist.data.datasets[0].data=teams.map(t=>m[t].pkgs);CH.dist.data.datasets[0].backgroundColor=colors;CH.dist.update();
    // Line
    const hrs=Array.from({length:11},(_,i)=>`${String(i+13).padStart(2,'0')}:00`);
    const lineSets=teams.map(t=>({label:getInitials(t),data:hrs.map((_,idx)=>{const h=idx+13;return S.filteredOps.filter(o=>o.team===t&&new Date(o.endTime).getHours()<h+1).reduce((s,o)=>s+o.packages,0);}),borderColor:getColor(t),fill:false,pointRadius:2}));
    CH.line.data.labels=hrs;CH.line.data.datasets=lineSets;CH.line.update();
    // Legend
    document.getElementById('chartLegend').innerHTML=teams.map((t,i)=>`<div class="legend-item"><div class="legend-swatch" style="background:${colors[i]}"></div><span>${inits[i]}: ${t}</span></div>`).join('');
    // Table
    const tb=document.querySelector('#opsTable tbody');
    tb.innerHTML=!S.filteredOps.length?`<tr><td colspan="7" style="text-align:center;color:var(--c-muted);padding:16px">Nenhuma operação encontrada.</td></tr>`:
    [...S.filteredOps].sort((a,b)=>new Date(b.startTime)-new Date(a.startTime)).map(op=>{
      const dur=(new Date(op.endTime)-new Date(op.startTime))/36e5;
      const eff=dur>0?(op.packages/dur).toFixed(1):'0';
      const c=getColor(op.team);
      return`<tr><td><span style="display:inline-block;width:20px;height:20px;border-radius:3px;background:${c};text-align:center;font-size:9px;font-weight:700;line-height:20px;color:#000;margin-right:6px">${getInitials(op.team)}</span>${op.team}</td><td class="td-mono">${fmtDt(op.startTime)}</td><td class="td-mono">${fmtDt(op.endTime)}</td><td class="td-mono">${op.packages.toLocaleString('pt-BR')}</td><td>1</td><td class="td-mono">${dur.toFixed(2)}</td><td><span class="badge badge-green">${eff}</span></td></tr>`;
    }).join('');
    this.populateDL();
  },
  exportar(){
    const data=S.filteredOps.map(op=>{const dur=(new Date(op.endTime)-new Date(op.startTime))/36e5;return{'Dupla':op.team,'Início':fmtDt(op.startTime),'Término':fmtDt(op.endTime),'Pacotes':op.packages,'Duração (h)':dur.toFixed(2),'Eficiência (pac/h)':(op.packages/dur).toFixed(1)};});
    const ws=XLSX.utils.json_to_sheet(data);const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,'Operações');
    XLSX.writeFile(wb,`operacoes-${new Date().toISOString().slice(0,10)}.xlsx`);
  },
  limpar(){document.getElementById('modalConfirm').classList.add('open');},
  confirmarLimpar(){S.ops=[];S.filteredOps=[];S.teamColors={};S.colorIdx=0;closeModal();save();this.render();}
};

// ═══════════════════════════════════════════════
//  M — MOTORISTA
// ═══════════════════════════════════════════════
const M={
  addEmpresa(){
    const n=document.getElementById('mNovaEmpresa').value.trim();
    if(!n)return;
    if(!S.empresas.includes(n))S.empresas.push(n);
    document.getElementById('mNovaEmpresa').value='';
    populateEmpresaSelects();save();
    document.getElementById('mEmpresa').value=n;
  },
  registrar(){
    const nome=document.getElementById('mNome').value.trim();
    const wpp=document.getElementById('mWpp').value.trim();
    const empresa=document.getElementById('mEmpresa').value;
    const pkgs=parseInt(document.getElementById('mPkgs').value);
    if(!nome||!pkgs||pkgs<=0){alert('Preencha nome e pacotes.');return;}
    const mot={id:Date.now(),nome,whatsapp:wpp,empresa,pacotes:pkgs,status:'fila',docaId:null,horaInicio:null};
    S.motorista=mot;
    S.fila.push({id:mot.id,nome:mot.nome,empresa:mot.empresa,pacotes:mot.pacotes});
    addLog('reg',`${nome} entrou pela fila (portal).`);
    save();G.chamarProximo();G.render();this.render();
  },
  iniciarDescarga(){
    const m=S.motorista;if(!m||m.status!=='chamado')return;
    m.status='desc';m.horaInicio=new Date().toISOString();
    const doca=S.docas.find(d=>d.motorista&&d.motorista.id===m.id);
    if(doca){doca.status='ocupada';doca.motorista.horaInicio=m.horaInicio;}
    addLog('sys',`${m.nome} iniciou — Doca ${m.docaId}.`);
    save();this.render();G.render();
  },
  finalizar(){
    const m=S.motorista;if(!m||m.status!=='desc')return;
    const doca=S.docas.find(d=>d.motorista&&d.motorista.id===m.id);
    if(doca)G.finalizarDoca(doca.id);
    m.status='fim';save();this.render();
  },
  novo(){
    S.motorista=null;
    document.getElementById('mNome').value='';document.getElementById('mWpp').value='';document.getElementById('mPkgs').value='';
    save();this.render();
  },
  updateTimer(){
    if(!S.motorista||S.motorista.status!=='desc')return;
    const el=document.getElementById('mTimer');if(!el)return;
    const ms=Date.now()-new Date(S.motorista.horaInicio).getTime();
    const h=String(Math.floor(ms/3600000)).padStart(2,'0');
    const mi=String(Math.floor((ms%3600000)/60000)).padStart(2,'0');
    const s=String(Math.floor((ms%60000)/1000)).padStart(2,'0');
    el.textContent=`${h}:${mi}:${s}`;
  },
  render(){
    const m=S.motorista;
    const fv=document.getElementById('mFormView');
    const sv=document.getElementById('mStatusView');
    if(!m){fv.style.display='block';sv.style.display='none';return;}
    fv.style.display='none';sv.style.display='block';
    document.getElementById('mNomeDisplay').textContent=`${m.nome} · ${m.empresa} · ${m.pacotes.toLocaleString()} pacotes`;
    ['mStatusFila','mStatusChamado','mStatusDesc','mStatusFim'].forEach(id=>document.getElementById(id).style.display='none');
    if(m.status==='fila'){
      document.getElementById('mStatusFila').style.display='block';
      const pos=S.fila.findIndex(x=>x.id===m.id)+1;
      document.getElementById('mPosNum').textContent=pos>0?pos:'—';
    } else if(m.status==='chamado'){
      document.getElementById('mStatusChamado').style.display='block';
      document.getElementById('mDocaNum').textContent=`DOCA ${m.docaId}`;
    } else if(m.status==='desc'){
      document.getElementById('mStatusDesc').style.display='block';
      this.updateTimer();
    } else if(m.status==='fim'){
      document.getElementById('mStatusFim').style.display='block';
    }
  }
};

// ═══════════════════════════════════════════════
//  INIT
// ═══════════════════════════════════════════════
document.addEventListener('DOMContentLoaded',()=>{
  loadState();
  initCharts();
  populateEmpresaSelects();
  D.populateDL();

  // Clock
  function tick(){
    document.getElementById('clock').textContent=new Date().toLocaleTimeString('pt-BR');
  }
  tick();setInterval(tick,1000);

  // Auto-refresh 1s
  setInterval(()=>{
    if(document.getElementById('screen-app').classList.contains('active')){
      G.render();
    }
    if(document.getElementById('screen-motorista').classList.contains('active')){
      M.updateTimer();
      if(S.motorista&&S.motorista.status==='fila'){
        const pos=S.fila.findIndex(x=>x.id===S.motorista.id)+1;
        const el=document.getElementById('mPosNum');
        if(el)el.textContent=pos>0?pos:'—';
      }
    }
  },1000);
});
</script>
</body>
</html>
