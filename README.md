
<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>DocaOps — Sistema Operacional</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.2/dist/chart.umd.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
<style>
*{box-sizing:border-box;margin:0;padding:0}
:root{
  --bg:#0d1117;--s1:#161b22;--s2:#1c2331;--s3:#21262d;
  --b1:#30363d;--b2:#444c56;--b3:#6e7681;
  --tx:#c9d1d9;--muted:#8b949e;
  --acc:#e6a817;--acc2:#b8820e;
  --green:#3fb950;--red:#f85149;--blue:#58a6ff;--purple:#bc8cff;--orange:#ffa657;
}
body{background:var(--bg);color:var(--tx);font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;font-size:13px;line-height:1.5;min-height:100vh}
.screen{display:none;min-height:100vh}.screen.active{display:flex;flex-direction:column}

/* LOGIN */
#sc-login{align-items:center;justify-content:center}
.lcard{background:var(--s1);border:1px solid var(--b1);border-radius:10px;padding:40px 44px;width:360px}
.lbrand{display:flex;align-items:center;gap:10px;margin-bottom:32px}
.lbrand .ico{width:36px;height:36px;background:var(--acc);border-radius:7px;display:flex;align-items:center;justify-content:center}
.lbrand .ico svg{width:18px;height:18px}
.lbrand h1{font-size:15px;font-weight:700;color:#fff}.lbrand p{font-size:11px;color:var(--muted)}
.lfield{margin-bottom:14px}
.lfield label{display:block;font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.07em;color:var(--muted);margin-bottom:5px}
.lfield input{width:100%;background:var(--bg);border:1px solid var(--b1);border-radius:6px;padding:9px 12px;color:var(--tx);font-size:13px;outline:none;transition:border-color .15s}
.lfield input:focus{border-color:var(--acc)}
.lerr{font-size:11px;color:var(--red);margin-top:6px;display:none}
.lbtn{width:100%;background:var(--acc);border:none;border-radius:6px;padding:11px;color:#000;font-weight:700;font-size:13px;cursor:pointer;margin-top:4px;transition:background .15s}
.lbtn:hover{background:var(--acc2)}
.lmoto{margin-top:20px;padding-top:16px;border-top:1px solid var(--b1);text-align:center;font-size:11px;color:var(--muted)}
.lmoto a{color:var(--acc);text-decoration:none}

/* TOPBAR */
.topbar{background:var(--s1);border-bottom:1px solid var(--b1);height:44px;padding:0 20px;display:flex;align-items:center;flex-shrink:0;position:sticky;top:0;z-index:100}
.tbrand{display:flex;align-items:center;gap:8px;padding-right:20px;border-right:1px solid var(--b1);margin-right:4px}
.tbrand .ico{width:24px;height:24px;background:var(--acc);border-radius:4px;display:flex;align-items:center;justify-content:center}
.tbrand .ico svg{width:12px;height:12px}
.tbrand span{font-size:13px;font-weight:700;color:#fff}
.navtabs{display:flex;gap:1px;flex:1;padding-left:4px}
.ntab{padding:0 16px;height:44px;display:flex;align-items:center;gap:6px;font-size:12px;font-weight:500;color:var(--muted);cursor:pointer;border:none;border-bottom:2px solid transparent;background:none;transition:all .15s;white-space:nowrap}
.ntab:hover{color:var(--tx)}.ntab.active{color:var(--acc);border-bottom-color:var(--acc)}
.ntab svg{width:12px;height:12px;fill:currentColor;opacity:.7}
.tright{display:flex;align-items:center;gap:10px;margin-left:auto}
.tclock{font-family:'SF Mono','Consolas',monospace;font-size:12px;color:var(--muted);background:var(--bg);border:1px solid var(--b1);padding:3px 10px;border-radius:4px}
.tpill{font-size:10px;font-weight:700;padding:3px 8px;border-radius:3px}
.tpill.on{background:rgba(63,185,80,.12);color:var(--green);border:1px solid rgba(63,185,80,.25)}
.tpill.off{background:rgba(248,81,73,.1);color:var(--red);border:1px solid rgba(248,81,73,.25)}
.tbtn{background:none;border:1px solid var(--b1);border-radius:4px;padding:4px 10px;color:var(--muted);font-size:11px;cursor:pointer;transition:all .15s}
.tbtn:hover{border-color:var(--b2);color:var(--tx)}

/* LAYOUT */
.mcontent{flex:1;overflow:auto;padding:14px 18px;display:flex;flex-direction:column;gap:12px}

/* PANEL */
.panel{background:var(--s1);border:1px solid var(--b1);border-radius:7px;overflow:hidden}
.phead{padding:9px 14px;border-bottom:1px solid var(--b1);display:flex;align-items:center;justify-content:space-between;gap:8px}
.ptitle{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.08em;color:var(--muted);display:flex;align-items:center;gap:6px}
.dot{width:6px;height:6px;border-radius:50%;display:inline-block}
.dg{background:var(--green)}.da{background:var(--acc)}.db{background:var(--blue)}.dr{background:var(--red)}.dm{background:var(--b3)}.dp{background:var(--purple)}

/* KPI STRIP */
.kstrip{display:grid;grid-template-columns:repeat(auto-fit,minmax(110px,1fr));gap:1px;background:var(--b1);border:1px solid var(--b1);border-radius:7px;overflow:hidden}
.kpi{background:var(--s1);padding:11px 14px}
.kl{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.07em;color:var(--muted);margin-bottom:3px}
.kv{font-size:22px;font-weight:800;color:#fff;font-family:'SF Mono','Consolas',monospace;line-height:1}
.kv.g{color:var(--green)!important;background:none!important}
.kv.a{color:var(--acc)!important;background:none!important}
.kv.b{color:var(--blue)!important;background:none!important}
.kv.r{color:var(--red)!important;background:none!important}
.kv.p{color:var(--purple)!important;background:none!important}
.ks{font-size:10px;color:var(--muted);margin-top:2px}
.pbar{height:5px;background:var(--b1);border-radius:3px;overflow:hidden;margin-top:5px}
.pfill{height:100%;background:linear-gradient(90deg,var(--acc2),var(--acc));border-radius:3px;transition:width .4s}

/* BOTÕES */
.btn{padding:5px 12px;border-radius:5px;border:1px solid var(--b1);background:var(--s2);color:var(--tx);font-size:11px;font-weight:600;cursor:pointer;transition:all .15s;white-space:nowrap}
.btn:hover{border-color:var(--b2);background:var(--b1)}
.btn.g{background:#1a3a1e;border-color:#2ea043;color:var(--green)}.btn.g:hover{background:#2ea043;color:#000}
.btn.r{background:#3a1a1a;border-color:var(--red);color:var(--red)}.btn.r:hover{background:var(--red);color:#fff}
.btn.a{background:#3a2f00;border-color:var(--acc);color:var(--acc)}.btn.a:hover{background:var(--acc);color:#000}
.btn.b{background:#0a1f3a;border-color:var(--blue);color:var(--blue)}.btn.b:hover{background:var(--blue);color:#000}
.btn.p{background:#2a1a40;border-color:var(--purple);color:var(--purple)}.btn.p:hover{background:var(--purple);color:#000}

/* FORM */
.frow{display:flex;flex-wrap:wrap;gap:8px;align-items:flex-end;padding:10px 14px}
.fld{display:flex;flex-direction:column;gap:3px}
.fld label{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.06em;color:var(--muted)}
.fld input,.fld select{background:var(--bg);border:1px solid var(--b1);border-radius:5px;padding:6px 9px;color:var(--tx);font-size:12px;outline:none;transition:border-color .15s}
.fld input:focus,.fld select:focus{border-color:var(--acc)}
.fld select option{background:var(--s2)}

/* ─── PAINEL CENTRAL ─── */
.ops-layout{display:grid;grid-template-columns:200px 200px 1fr;gap:10px;padding:12px 14px;align-items:start}

/* LISTAS DRAG (duplas / motoristas) */
.drag-list{border:1px solid var(--b1);border-radius:6px;background:var(--bg);overflow:hidden;display:flex;flex-direction:column}
.drag-list-head{padding:7px 10px;background:var(--s2);border-bottom:1px solid var(--b1);font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.07em;color:var(--muted);display:flex;align-items:center;justify-content:space-between;flex-shrink:0}
.drag-list-head span.cnt{font-family:'SF Mono','Consolas',monospace;font-size:11px;color:var(--acc)}
.drag-items{overflow-y:auto;max-height:320px}
.drag-item{padding:7px 10px;border-bottom:1px solid rgba(48,54,61,.5);display:flex;align-items:center;gap:8px;cursor:grab;user-select:none;transition:background .12s}
.drag-item:last-child{border-bottom:none}
.drag-item:hover{background:var(--s2)}
.drag-item.dragging{opacity:.35;cursor:grabbing}
.drag-avatar{width:26px;height:26px;border-radius:4px;display:flex;align-items:center;justify-content:center;font-size:9px;font-weight:700;color:#000;flex-shrink:0;font-family:'SF Mono','Consolas',monospace}
.drag-info{flex:1;min-width:0}
.drag-name{font-size:11px;font-weight:600;color:var(--tx);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.drag-sub{font-size:9px;color:var(--muted)}
.drag-handle{font-size:11px;color:var(--b3);flex-shrink:0}
.drag-empty{padding:14px 10px;font-size:10px;color:var(--muted);text-align:center;font-style:italic;line-height:1.5}
.drag-type-badge{font-size:8px;font-weight:700;padding:1px 5px;border-radius:3px;flex-shrink:0}
.tb-dupla{background:rgba(230,168,23,.15);color:var(--acc)}
.tb-moto{background:rgba(88,166,255,.15);color:var(--blue)}

/* ─── DOCAS GRID ─── */
.docas-wrap{display:flex;flex-direction:column;gap:6px}
.docas-label{font-size:10px;color:var(--muted);display:flex;align-items:center;gap:12px;margin-bottom:2px}
.docas-label span{display:flex;align-items:center;gap:4px}
.docas-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(120px,1fr));gap:5px}

/* DOCA CARD */
.doca{border:1px solid var(--b1);border-radius:6px;padding:8px;min-height:130px;display:flex;flex-direction:column;align-items:center;text-align:center;gap:3px;transition:all .15s;position:relative;background:var(--bg)}
.doca.over{background:rgba(230,168,23,.06);border-color:var(--acc);border-style:dashed}
.doca.dup{background:rgba(63,185,80,.04);border-color:var(--green)}
.doca.mot{background:rgba(88,166,255,.04);border-color:var(--blue)}
.doca.both{background:rgba(188,140,255,.04);border-color:var(--purple)}
.doca.cart{background:rgba(248,81,73,.04);border-color:var(--red)}
.doca-num{font-family:'SF Mono','Consolas',monospace;font-size:9px;font-weight:700;color:var(--muted);letter-spacing:.04em}
.doca-badge{font-size:8px;font-weight:700;text-transform:uppercase;letter-spacing:.05em;padding:2px 5px;border-radius:3px;margin-bottom:1px}
.db-livre{background:rgba(110,118,129,.1);color:var(--muted)}
.db-dupla{background:rgba(63,185,80,.15);color:var(--green)}
.db-moto{background:rgba(88,166,255,.15);color:var(--blue)}
.db-full{background:rgba(188,140,255,.15);color:var(--purple)}
.db-cart{background:rgba(248,81,73,.15);color:var(--red)}
.doca-slots{display:flex;flex-direction:column;gap:2px;width:100%;margin:2px 0}
.slot{font-size:9px;padding:2px 5px;border-radius:3px;display:flex;align-items:center;gap:3px;width:100%;text-align:left}
.slot-dup{background:rgba(63,185,80,.1);color:var(--green);border:1px solid rgba(63,185,80,.2)}
.slot-mot{background:rgba(88,166,255,.1);color:var(--blue);border:1px solid rgba(88,166,255,.2)}
.slot-dot{width:5px;height:5px;border-radius:50%;flex-shrink:0}
.slot-name{flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.doca-timer{font-family:'SF Mono','Consolas',monospace;font-size:10px;font-weight:700;color:var(--acc)}
.doca-kpis{font-size:8px;color:var(--muted);line-height:1.4}
.doca-actions{display:flex;flex-direction:column;gap:2px;width:100%;margin-top:auto}
.doca-actions .btn{font-size:9px;padding:3px 5px;width:100%}
.drop-hint{font-size:9px;color:var(--b3);margin-top:4px}
@keyframes pulse-g{0%,100%{border-color:var(--green)}50%{border-color:rgba(63,185,80,.2)}}
@keyframes pulse-b{0%,100%{border-color:var(--blue)}50%{border-color:rgba(88,166,255,.2)}}
@keyframes pulse-p{0%,100%{border-color:var(--purple)}50%{border-color:rgba(188,140,255,.2)}}
.doca.dup{animation:pulse-g 2.5s infinite}
.doca.mot{animation:pulse-b 2.5s infinite}
.doca.both{animation:pulse-p 2.5s infinite}

/* ─── KPIs PRODUTIVIDADE ─── */
.prod-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(200px,1fr));gap:6px;padding:10px 14px}
.prod-card{background:var(--bg);border:1px solid var(--b1);border-radius:6px;padding:9px 11px;display:flex;flex-direction:column;gap:4px}
.prod-header{display:flex;align-items:center;gap:6px}
.prod-avatar{width:22px;height:22px;border-radius:3px;display:flex;align-items:center;justify-content:center;font-size:8px;font-weight:700;color:#000;flex-shrink:0;font-family:'SF Mono','Consolas',monospace}
.prod-name{font-size:10px;font-weight:700;color:var(--tx);flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.prod-doca{font-family:'SF Mono','Consolas',monospace;font-size:9px;color:var(--muted)}
.prod-metrics{display:grid;grid-template-columns:1fr 1fr 1fr;gap:3px;margin-top:2px}
.pm{background:var(--s2);border-radius:3px;padding:3px 5px;text-align:center}
.pm-v{font-size:11px;font-weight:800;font-family:'SF Mono','Consolas',monospace;line-height:1}
.pm-l{font-size:8px;color:var(--muted);margin-top:1px}
.prod-bar-wrap{margin-top:3px}
.prod-bar{height:3px;background:var(--b1);border-radius:2px;overflow:hidden}
.prod-bar-fill{height:100%;border-radius:2px;transition:width .5s}
.prod-status{font-size:8px;display:flex;align-items:center;gap:4px;color:var(--muted)}
.pst-dot{width:5px;height:5px;border-radius:50%;flex-shrink:0}

/* FILA */
.fila-wrap{}
.femp{border-bottom:1px solid var(--b1)}.femp:last-child{border-bottom:none}
.femp-head{padding:6px 14px;background:var(--bg);border-bottom:1px solid rgba(48,54,61,.5);font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.06em;color:var(--muted);display:flex;align-items:center;justify-content:space-between}
.femp-cnt{font-family:'SF Mono','Consolas',monospace;background:var(--b1);border-radius:3px;padding:0 5px;font-size:10px;color:var(--tx)}
.fitem{padding:5px 14px;display:flex;align-items:center;gap:8px;font-size:11px;border-bottom:1px solid rgba(48,54,61,.4)}
.fitem:last-child{border-bottom:none}
.fpos{width:18px;height:18px;border-radius:3px;background:var(--s2);border:1px solid var(--b1);font-size:9px;font-weight:700;display:flex;align-items:center;justify-content:center;flex-shrink:0;color:var(--acc);font-family:'SF Mono','Consolas',monospace}
.fnome{flex:1}.fpkgs{font-family:'SF Mono','Consolas',monospace;font-size:10px;color:var(--muted)}
.fbtn{font-size:9px;padding:2px 6px}

/* LOG */
.log-wrap{max-height:180px;overflow-y:auto;padding:6px 14px}
.lentry{display:flex;gap:8px;padding:3px 0;border-bottom:1px solid rgba(48,54,61,.4);font-size:11px}
.lentry:last-child{border-bottom:none}
.lt{font-family:'SF Mono','Consolas',monospace;font-size:10px;color:var(--muted);white-space:nowrap;min-width:54px}
.ld{width:5px;height:5px;border-radius:50%;flex-shrink:0;margin-top:4px}
.lm{color:var(--tx)}

/* LAYOUT GESTAO */
.gestao-main{display:grid;grid-template-columns:1fr 300px;gap:12px}
.gestao-left{display:flex;flex-direction:column;gap:12px}
.gestao-right{display:flex;flex-direction:column;gap:12px}
@media(max-width:1100px){.gestao-main{grid-template-columns:1fr}}
@media(max-width:900px){.ops-layout{grid-template-columns:1fr 1fr;}.ops-layout .docas-wrap{grid-column:1/-1}}

/* DASH */
.dash-grid2{display:grid;grid-template-columns:1fr 1fr;gap:12px}
.dash-grid3{display:grid;grid-template-columns:1fr 1fr 1fr;gap:12px}
.cw{padding:12px 14px}.cw canvas{max-height:200px}
.twrap{overflow-x:auto}
table{width:100%;border-collapse:collapse}
thead th{background:var(--bg);color:var(--muted);font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.06em;padding:7px 12px;text-align:left;border-bottom:1px solid var(--b1);white-space:nowrap}
tbody td{padding:6px 12px;border-bottom:1px solid rgba(48,54,61,.5);font-size:11px}
tbody tr:last-child td{border-bottom:none}
tbody tr:hover td{background:rgba(255,255,255,.02)}
.mono{font-family:'SF Mono','Consolas',monospace;font-size:10px}
.badge{display:inline-block;padding:1px 6px;border-radius:3px;font-size:10px;font-weight:700}
.badge.g{background:rgba(63,185,80,.15);color:var(--green);border:1px solid rgba(63,185,80,.2)}
.badge.a{background:rgba(230,168,23,.15);color:var(--acc);border:1px solid rgba(230,168,23,.2)}
.badge.r{background:rgba(248,81,73,.1);color:var(--red);border:1px solid rgba(248,81,73,.2)}
.ch-legend{display:flex;flex-wrap:wrap;gap:12px;padding:8px 14px;border-top:1px solid var(--b1)}
.ch-li{display:flex;align-items:center;gap:5px;font-size:10px;color:var(--muted)}
.ch-sw{width:12px;height:3px;border-radius:2px}
@media(max-width:900px){.dash-grid2{grid-template-columns:1fr}.dash-grid3{grid-template-columns:1fr}}

/* MOTORISTA SCREEN */
#sc-moto{align-items:center;justify-content:center;background:var(--bg);flex:1}
.mwrap{width:100%;max-width:400px;padding:20px}
.mcard{background:var(--s1);border:1px solid var(--b1);border-radius:9px;overflow:hidden}
.mhead{background:var(--bg);border-bottom:1px solid var(--b1);padding:14px 18px;display:flex;align-items:center;gap:10px}
.mhead .ico{width:30px;height:30px;background:var(--acc);border-radius:5px;display:flex;align-items:center;justify-content:center;flex-shrink:0}
.mhead .ico svg{width:15px;height:15px}
.mhead h2{font-size:13px;font-weight:700;color:#fff}.mhead p{font-size:10px;color:var(--muted)}
.mbody{padding:18px}
.mstatus{text-align:center;padding:16px 0}
.bignum{font-family:'SF Mono','Consolas',monospace;font-size:64px;font-weight:800;line-height:1;color:var(--acc)}
.bigdoca{font-family:'SF Mono','Consolas',monospace;font-size:44px;font-weight:800;color:var(--green);line-height:1}
.stag{display:inline-flex;align-items:center;gap:5px;padding:3px 10px;border-radius:3px;font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.05em;margin-top:8px}
.sw{background:rgba(230,168,23,.1);color:var(--acc);border:1px solid rgba(230,168,23,.25)}
.sc{background:rgba(88,166,255,.1);color:var(--blue);border:1px solid rgba(88,166,255,.25)}
.sd{background:rgba(63,185,80,.1);color:var(--green);border:1px solid rgba(63,185,80,.25)}
.mtimer{font-family:'SF Mono','Consolas',monospace;font-size:28px;font-weight:700;color:var(--acc);text-align:center;margin:8px 0}
.mabtn{width:100%;padding:12px;border:none;border-radius:6px;font-size:13px;font-weight:700;cursor:pointer;display:flex;align-items:center;justify-content:center;gap:8px;margin-top:10px;transition:all .15s}
.mabtn.g{background:var(--green);color:#000}.mabtn.g:hover{background:#4dd860}
.mabtn.r{background:var(--red);color:#fff}.mabtn.r:hover{background:#ff6459}
.mabtn.x{background:var(--s2);color:var(--tx);border:1px solid var(--b1)}.mabtn.x:hover{border-color:var(--b2)}
.emp-add{display:flex;gap:6px;margin-top:6px}
.emp-add input{flex:1;background:var(--bg);border:1px solid var(--b1);border-radius:5px;padding:6px 9px;color:var(--tx);font-size:12px;outline:none}
.emp-add input:focus{border-color:var(--acc)}

/* MODAL */
.modal-bg{position:fixed;inset:0;background:rgba(0,0,0,.75);z-index:500;display:none;align-items:center;justify-content:center}
.modal-bg.open{display:flex}
.modal{background:var(--s1);border:1px solid var(--b1);border-radius:8px;padding:22px;width:320px}
.modal h3{font-size:13px;font-weight:700;color:#fff;margin-bottom:7px}
.modal p{font-size:11px;color:var(--muted);line-height:1.6;margin-bottom:14px}
.mbtns{display:flex;gap:8px;justify-content:flex-end}
</style>
</head>
<body>

<!-- ═══ LOGIN ═══ -->
<div id="sc-login" class="screen active">
  <div class="lcard">
    <div class="lbrand">
      <div class="ico"><svg viewBox="0 0 24 24" fill="none" stroke="#000" stroke-width="2.5"><path d="M3 9l9-7 9 7v11a2 2 0 01-2 2H5a2 2 0 01-2-2z"/><polyline points="9,22 9,12 15,12 15,22"/></svg></div>
      <div><h1>DocaOps</h1><p>Sistema de Gestão Operacional</p></div>
    </div>
    <div class="lfield"><label>Usuário</label><input type="text" id="lu" placeholder="Seu usuário" autocomplete="off"></div>
    <div class="lfield"><label>Senha</label><input type="password" id="lp" placeholder="••••••" onkeydown="if(event.key==='Enter')doLogin()"></div>
    <div id="lerr" class="lerr">Usuário ou senha incorretos.</div>
    <button class="lbtn" onclick="doLogin()">Acessar o Sistema</button>
    <div class="lmoto">Motorista? <a href="#" onclick="showMoto()">Acessar portal de chegada</a></div>
  </div>
</div>

<!-- ═══ APP ═══ -->
<div id="sc-app" class="screen">
  <div class="topbar">
    <div class="tbrand">
      <div class="ico"><svg viewBox="0 0 24 24" fill="none" stroke="#000" stroke-width="2.5"><path d="M3 9l9-7 9 7v11a2 2 0 01-2 2H5a2 2 0 01-2-2z"/><polyline points="9,22 9,12 15,12 15,22"/></svg></div>
      <span>DocaOps</span>
    </div>
    <div class="navtabs">
      <button class="ntab active" onclick="setTab('gestao',this)">
        <svg viewBox="0 0 16 16" fill="currentColor"><rect x="1" y="1" width="6" height="6" rx="1"/><rect x="9" y="1" width="6" height="6" rx="1"/><rect x="1" y="9" width="6" height="6" rx="1"/><rect x="9" y="9" width="6" height="6" rx="1"/></svg>
        <span>Doca em Tempo Real</span>
      </button>
      <button class="ntab" onclick="setTab('dash',this)">
        <svg viewBox="0 0 16 16" fill="currentColor"><path d="M1 14V8h3v6H1zm4-8v8h3V6H5zm4 3v5h3V9H9zm4-5v10h2V4h-2z"/></svg>
        <span>Dashboard & Analytics</span>
      </button>
    </div>
    <div class="tright">
      <div id="tpill" class="tpill off">● INATIVO</div>
      <div class="tclock" id="clock">--:--:--</div>
      <button class="tbtn" onclick="doLogout()">Sair</button>
    </div>
  </div>

  <!-- ═══ TAB GESTÃO ═══ -->
  <div id="tab-gestao" class="mcontent">

    <!-- KPIs PRINCIPAIS -->
    <div class="kstrip">
      <div class="kpi">
        <div class="kl">Meta de Pacotes</div>
        <div style="display:flex;align-items:center;gap:5px;margin-top:2px">
          <input type="number" id="metaInput" value="5000" style="width:80px;background:var(--bg);border:1px solid var(--b1);border-radius:4px;padding:2px 6px;color:#fff;font-size:20px;font-weight:800;font-family:'SF Mono','Consolas',monospace;outline:none">
          <button class="btn a" onclick="G.setMeta()" style="padding:3px 8px;font-size:10px">✓</button>
        </div>
      </div>
      <div class="kpi"><div class="kl">Descarregados</div><div class="kv g" id="kDesc">0</div><div class="ks">pacotes hoje</div></div>
      <div class="kpi"><div class="kl">Docas Ativas</div><div class="kv b" id="kAtivas">0</div><div class="ks">de 13</div></div>
      <div class="kpi"><div class="kl">Duplas na Doca</div><div class="kv a" id="kDuplas">0</div><div class="ks">trabalhando</div></div>
      <div class="kpi"><div class="kl">Motoristas na Doca</div><div class="kv p" id="kMoots">0</div><div class="ks">descarregando</div></div>
      <div class="kpi"><div class="kl">Fila</div><div class="kv" id="kFila">0</div><div class="ks">aguardando</div></div>
      <div class="kpi" style="grid-column:span 2">
        <div class="kl">Progresso da Meta</div>
        <div style="display:flex;align-items:baseline;gap:8px;margin:2px 0 4px"><div class="kv" id="kPct">0%</div><div class="ks" id="kPctSub">0 / 5.000</div></div>
        <div class="pbar"><div class="pfill" id="kBar" style="width:0%"></div></div>
      </div>
    </div>

    <div class="gestao-main">
      <div class="gestao-left">

        <!-- CADASTRO RÁPIDO -->
        <div class="panel">
          <div class="phead">
            <div class="ptitle"><span class="dot da"></span>Cadastrar</div>
            <div style="display:flex;gap:6px">
              <button class="btn a" id="btnNewDupla" onclick="G.toggleForm('dupla')">+ Dupla</button>
              <button class="btn b" id="btnNewMoto" onclick="G.toggleForm('moto')">+ Motorista</button>
              <button class="btn r" onclick="G.addCarreta()">+ Carreta</button>
            </div>
          </div>
          <!-- Form dupla -->
          <div id="formDupla" style="display:none;border-bottom:1px solid var(--b1)">
            <div class="frow">
              <div class="fld" style="flex:1;min-width:150px"><label>Nomes da Dupla</label><input type="text" id="dNomes" placeholder="ex: João e Maria"></div>
              <div class="fld" style="width:90px"><label>Pacotes</label><input type="number" id="dPkgs" placeholder="0"></div>
              <div class="fld" style="align-self:flex-end"><button class="btn g" onclick="G.addDupla()">Adicionar</button></div>
            </div>
          </div>
          <!-- Form motorista -->
          <div id="formMoto" style="display:none;border-bottom:1px solid var(--b1)">
            <div class="frow">
              <div class="fld" style="flex:1;min-width:130px"><label>Nome</label><input type="text" id="mNome" placeholder="Nome do motorista"></div>
              <div class="fld" style="min-width:100px"><label>Empresa</label><select id="mEmp"></select></div>
              <div class="fld" style="width:90px"><label>Pacotes</label><input type="number" id="mPkgs" placeholder="0"></div>
              <div class="fld" style="align-self:flex-end"><button class="btn b" onclick="G.addMoto()">Adicionar</button></div>
            </div>
          </div>
        </div>

        <!-- DRAG & DROP CENTRAL -->
        <div class="panel">
          <div class="phead">
            <div class="ptitle"><span class="dot da"></span>Alocar nas Docas — Arraste para a doca desejada</div>
            <div style="display:flex;gap:10px;font-size:10px;color:var(--muted)">
              <span><span style="color:var(--green)">●</span> Dupla</span>
              <span><span style="color:var(--blue)">●</span> Motorista</span>
              <span><span style="color:var(--purple)">●</span> Ambos</span>
            </div>
          </div>
          <div class="ops-layout">
            <!-- Duplas disponíveis -->
            <div class="drag-list" id="poolDuplas">
              <div class="drag-list-head">Duplas <span class="cnt" id="cntDuplas">0</span></div>
              <div class="drag-items" id="itemsDuplas"></div>
            </div>
            <!-- Motoristas disponíveis -->
            <div class="drag-list" id="poolMotos">
              <div class="drag-list-head">Motoristas <span class="cnt" id="cntMotos">0</span></div>
              <div class="drag-items" id="itemsMotos"></div>
            </div>
            <!-- Grid de docas -->
            <div class="docas-wrap">
              <div class="docas-label">
                <span><span class="dot da"></span>Docas 169–181</span>
                <span><span class="dot dg"></span>Dupla ativa</span>
                <span><span class="dot db"></span>Motorista</span>
                <span><span class="dot dp"></span>Dupla+Moto</span>
                <span><span class="dot dr"></span>Carreta</span>
              </div>
              <div class="docas-grid" id="docasGrid"></div>
            </div>
          </div>
        </div>

        <!-- KPIs DE PRODUTIVIDADE POR DOCA -->
        <div class="panel" id="prodPanel">
          <div class="phead"><div class="ptitle"><span class="dot dg"></span>Produtividade em Tempo Real — Docas Ativas</div></div>
          <div class="prod-grid" id="prodGrid">
            <div style="padding:14px;font-size:11px;color:var(--muted);text-align:center;grid-column:1/-1">Nenhuma doca ativa no momento.</div>
          </div>
        </div>

      </div><!-- /gestao-left -->

      <!-- RIGHT SIDEBAR -->
      <div class="gestao-right">
        <!-- Fila de espera -->
        <div class="panel">
          <div class="phead"><div class="ptitle"><span class="dot db"></span>Fila de Espera</div></div>
          <div id="filaWrap"></div>
        </div>
        <!-- Chart horários -->
        <div class="panel">
          <div class="phead"><div class="ptitle"><span class="dot dg"></span>Entradas por Hora</div></div>
          <div class="cw"><canvas id="hrChart"></canvas></div>
        </div>
        <!-- Log -->
        <div class="panel">
          <div class="phead"><div class="ptitle"><span class="dot dm"></span>Log de Eventos</div></div>
          <div class="log-wrap" id="logWrap"></div>
        </div>
      </div>
    </div>
  </div>

  <!-- ═══ TAB DASHBOARD ═══ -->
  <div id="tab-dash" class="mcontent" style="display:none">
    <div class="panel">
      <div class="phead"><div class="ptitle"><span class="dot da"></span>Registrar Operação Manual</div></div>
      <div class="frow">
        <div class="fld" style="flex:1;min-width:140px"><label>Dupla/Motorista</label><input type="text" id="opTeam" list="teamsDL" placeholder="Nome"><datalist id="teamsDL"></datalist></div>
        <div class="fld"><label>Início</label><input type="datetime-local" id="opStart"></div>
        <div class="fld"><label>Fim</label><input type="datetime-local" id="opEnd"></div>
        <div class="fld" style="width:90px"><label>Pacotes</label><input type="number" id="opPkgs" placeholder="0"></div>
        <button class="btn a" onclick="D.addOp()">+ Registrar</button>
        <div style="margin-left:auto;display:flex;gap:6px;align-items:center">
          <select id="fPeriod" onchange="D.filter(this.value)" style="background:var(--bg);border:1px solid var(--b1);border-radius:5px;padding:6px 9px;color:var(--tx);font-size:11px;outline:none">
            <option value="all">Todos</option><option value="today" selected>Hoje</option>
            <option value="yesterday">Ontem</option><option value="week">Esta semana</option><option value="month">Este mês</option>
          </select>
          <button class="btn" onclick="D.exportar()">↓ Excel</button>
          <button class="btn r" onclick="D.limpar()">⊘ Limpar</button>
        </div>
      </div>
    </div>

    <div class="kstrip">
      <div class="kpi"><div class="kl">Total Pacotes</div><div class="kv a" id="dTot">0</div></div>
      <div class="kpi"><div class="kl">Caminhões</div><div class="kv b" id="dTrk">0</div></div>
      <div class="kpi"><div class="kl">Duplas/Motos</div><div class="kv" id="dTeams">0</div></div>
      <div class="kpi"><div class="kl">Mais Rápida</div><div class="kv g" id="dFast" style="font-size:13px;margin-top:3px">—</div></div>
      <div class="kpi"><div class="kl">Mais Lenta</div><div class="kv r" id="dSlow" style="font-size:13px;margin-top:3px">—</div></div>
      <div class="kpi"><div class="kl">Média pac/h</div><div class="kv" id="dAvg">0</div></div>
      <div class="kpi"><div class="kl">Tempo Médio</div><div class="kv" id="dAvgT" style="font-size:16px;margin-top:3px">—</div></div>
    </div>

    <div class="dash-grid3">
      <div class="panel"><div class="phead"><div class="ptitle"><span class="dot da"></span>Volume por Dupla</div></div><div class="cw"><canvas id="cVol"></canvas></div></div>
      <div class="panel"><div class="phead"><div class="ptitle"><span class="dot db"></span>Eficiência (pac/h)</div></div><div class="cw"><canvas id="cEff"></canvas></div></div>
      <div class="panel"><div class="phead"><div class="ptitle"><span class="dot dg"></span>Caminhões</div></div><div class="cw"><canvas id="cTrk"></canvas></div></div>
    </div>
    <div class="dash-grid2">
      <div class="panel"><div class="phead"><div class="ptitle"><span class="dot dm"></span>Distribuição</div></div><div class="cw"><canvas id="cDist"></canvas></div></div>
      <div class="panel">
        <div class="phead"><div class="ptitle"><span class="dot da"></span>Acumulado por Hora</div></div>
        <div class="cw"><canvas id="cLine"></canvas></div>
        <div class="ch-legend" id="lineLegend"></div>
      </div>
    </div>
    <div class="panel">
      <div class="phead"><div class="ptitle"><span class="dot dm"></span>Detalhamento das Operações</div></div>
      <div class="twrap">
        <table id="opsTable">
          <thead><tr><th>Nome</th><th>Tipo</th><th>Início</th><th>Término</th><th>Pacotes</th><th>Duração (h)</th><th>Eficiência</th><th>Rating</th></tr></thead>
          <tbody></tbody>
        </table>
      </div>
    </div>
  </div>
</div>

<!-- ═══ PORTAL MOTORISTA ═══ -->
<div id="sc-moto" class="screen">
  <div class="mwrap">
    <div class="mcard">
      <div class="mhead">
        <div class="ico"><svg viewBox="0 0 24 24" fill="none" stroke="#000" stroke-width="2.5"><rect x="1" y="3" width="15" height="13" rx="2"/><path d="M16 8l5 3v5h-5V8z"/><circle cx="5.5" cy="18.5" r="2.5"/><circle cx="18.5" cy="18.5" r="2.5"/></svg></div>
        <div><h2>Portal do Motorista</h2><p>Registro e acompanhamento</p></div>
      </div>
      <div class="mbody">
        <div id="mForm">
          <div style="display:flex;flex-direction:column;gap:10px">
            <div class="fld"><label>Nome Completo</label><input type="text" id="mfNome" placeholder="Seu nome" style="font-size:13px;padding:8px 10px"></div>
            <div class="fld"><label>WhatsApp</label><input type="tel" id="mfWpp" placeholder="11999998888" style="font-size:13px;padding:8px 10px"></div>
            <div class="fld">
              <label>Empresa</label>
              <select id="mfEmp" style="font-size:13px;padding:8px 10px"></select>
              <div class="emp-add">
                <input type="text" id="mfNovaEmp" placeholder="Outra empresa? Digite aqui...">
                <button class="btn a" onclick="M.addEmpresa()">+ Add</button>
              </div>
            </div>
            <div class="fld"><label>Volume de Pacotes</label><input type="number" id="mfPkgs" placeholder="Nº de pacotes" style="font-size:13px;padding:8px 10px"></div>
            <button class="mabtn g" onclick="M.registrar()">✓ ENTRAR NA FILA</button>
            <button class="mabtn x" onclick="showLogin()" style="font-size:11px;padding:8px;margin-top:2px">← Voltar ao login</button>
          </div>
        </div>
        <div id="mStatus" style="display:none">
          <div id="msWait" class="mstatus">
            <div style="font-size:11px;color:var(--muted);margin-bottom:4px">Posição na fila</div>
            <div class="bignum" id="mPos">—</div>
            <div class="stag sw">● Aguardando</div>
            <div style="font-size:10px;color:var(--muted);margin-top:10px">Você será chamado em breve.</div>
          </div>
          <div id="msCall" style="display:none;text-align:center;padding:16px 0">
            <div style="font-size:11px;color:var(--blue);font-weight:700;margin-bottom:8px">▲ SUA VEZ!</div>
            <div class="bigdoca" id="mDocaNum">—</div>
            <div style="font-size:10px;color:var(--muted);margin-top:4px">Dirija-se à doca indicada</div>
            <button class="mabtn g" onclick="M.iniciar()" style="margin-top:16px">▶ Iniciar Descarga</button>
          </div>
          <div id="msDesc" style="display:none">
            <div style="text-align:center;padding:12px 0">
              <div style="font-size:10px;color:var(--muted)">Em descarga</div>
              <div class="mtimer" id="mTimer">00:00:00</div>
              <div class="stag sd">● Descarregando</div>
            </div>
            <button class="mabtn r" onclick="M.finalizar()">■ Finalizar Descarga</button>
          </div>
          <div id="msFim" style="display:none;text-align:center;padding:20px 0">
            <div style="font-size:36px;color:var(--green)">✓</div>
            <div style="font-size:15px;font-weight:700;color:var(--green);margin-top:8px">Descarga Concluída!</div>
            <div style="font-size:11px;color:var(--muted);margin-top:6px">Obrigado pelo serviço.</div>
            <button class="mabtn x" onclick="M.novo()" style="margin-top:14px">← Novo Registro</button>
          </div>
          <div id="mNomeDisp" style="margin-top:10px;padding-top:10px;border-top:1px solid var(--b1);font-size:10px;color:var(--muted);text-align:center"></div>
          <button class="mabtn x" onclick="M.novo()" style="font-size:10px;padding:7px;margin-top:8px">↩ Sair</button>
        </div>
      </div>
    </div>
    <div style="text-align:center;margin-top:12px;font-size:10px;color:var(--muted)">
      Gestor? <a href="#" onclick="showLogin()" style="color:var(--acc)">Painel de controle</a>
    </div>
  </div>
</div>

<!-- MODAL -->
<div class="modal-bg" id="modalOk">
  <div class="modal">
    <h3>Confirmar exclusão</h3>
    <p>Todos os dados de operações serão apagados. Esta ação não pode ser desfeita.</p>
    <div class="mbtns">
      <button class="btn" onclick="closeModal()">Cancelar</button>
      <button class="btn r" onclick="D.confirmar()">Limpar tudo</button>
    </div>
  </div>
</div>

<script>
// ══════════════════════════════════════════════════════
//  ESTADO
// ══════════════════════════════════════════════════════
const S={
  // Docas 169-181
  docas: Array.from({length:13},(_,i)=>({
    id:169+i, status:'livre',
    dupla:null,   // {id,nomes,pkgs,color}
    motorista:null, // {id,nome,empresa,pacotes}
    inicio:null,  // horário que ficou ocupada
  })),
  pool_duplas:[],   // duplas disponíveis para drag
  pool_motos:[],    // motoristas disponíveis para drag
  duplas_all:[],    // todas as duplas cadastradas
  ops:[],           // operações concluídas
  filteredOps:[],
  fila:[],          // fila de espera (motoristas sem doca)
  meta:5000, desc:0,
  historico:{},
  log:[],
  teamColors:{}, cidx:0,
  motorista:null,   // motorista logado no portal
  empresas:['Prálog','Imediato','Hawk','Ontime'],
};

const PAL=['#e6a817','#58a6ff','#3fb950','#f85149','#bc8cff','#ffa657','#39d353','#ff7b72','#d2a8ff','#79c0ff','#56d364','#f0883e'];

// ── persistência ──
function sv(){
  try{
    ['docas','pool_duplas','pool_motos','duplas_all','ops','fila','historico','log','teamColors','empresas'].forEach(k=>localStorage.setItem('s_'+k,JSON.stringify(S[k])));
    ['meta','desc','cidx'].forEach(k=>localStorage.setItem('s_'+k,S[k]));
    localStorage.setItem('s_motorista',JSON.stringify(S.motorista));
  }catch(e){}
}
function ld(){
  try{
    ['docas','pool_duplas','pool_motos','duplas_all','ops','fila','historico','log','teamColors','empresas'].forEach(k=>{const v=localStorage.getItem('s_'+k);if(v)S[k]=JSON.parse(v);});
    ['meta','desc','cidx'].forEach(k=>{const v=localStorage.getItem('s_'+k);if(v!==null)S[k]=isNaN(+v)?S[k]:+v;});
    const mo=localStorage.getItem('s_motorista');if(mo)S.motorista=JSON.parse(mo);
  }catch(e){}
  // Garante docas com IDs 169-181 se veio sem dados
  if(!S.docas.length||S.docas[0]?.id<169){
    S.docas=Array.from({length:13},(_,i)=>({id:169+i,status:'livre',dupla:null,motorista:null,inicio:null}));
  }
  if(!S.ops.length) S.ops=genOps();
}
function genOps(){
  const names=["Carlos & Bruna","Lucas & Fernanda","Matheus & Gabriela","Rafael & Juliana","Vinicius & Patricia","Gustavo & Sandra"];
  const today=new Date().toISOString().slice(0,10);
  return names.map((n,i)=>{
    const h=13+(i%5),dur=60+Math.floor(Math.random()*90);
    const st=new Date(`${today}T${String(h).padStart(2,'0')}:${String((i*7)%60).padStart(2,'0')}`);
    const en=new Date(st.getTime()+dur*60000);
    return{id:i+1,team:n,tipo:'dupla',startTime:st.toISOString(),endTime:en.toISOString(),packages:1500+Math.floor(Math.random()*800)};
  });
}

// ── helpers ──
function addLog(tipo,msg){const h=new Date().toLocaleTimeString('pt-BR',{hour:'2-digit',minute:'2-digit',second:'2-digit'});S.log.unshift({tipo,msg,h});if(S.log.length>80)S.log.length=80;}
function color(t){if(!S.teamColors[t]){S.teamColors[t]=PAL[S.cidx%PAL.length];S.cidx++;}return S.teamColors[t];}
function initials(n){if(!n)return'??';const p=n.replace(/&/g,'e').split(/\s+e\s+/i).map(x=>x.trim());if(p.length>=2&&p[0]&&p[1])return(p[0][0]+p[1][0]).toUpperCase();const w=n.split(' ');if(w.length>=2)return(w[0][0]+w[w.length-1][0]).toUpperCase();return n.slice(0,2).toUpperCase();}
function fdt(d){return new Date(d).toLocaleString('pt-BR');}
function fmtMs(ms){const h=Math.floor(ms/3600000),m=Math.floor((ms%3600000)/60000),s=Math.floor((ms%60000)/1000);return`${String(h).padStart(2,'0')}:${String(m).padStart(2,'0')}:${String(s).padStart(2,'0')}`;}
function closeModal(){document.getElementById('modalOk').classList.remove('open');}
function fillEmpSels(){['mEmp','mfEmp'].forEach(id=>{const el=document.getElementById(id);if(!el)return;const cv=el.value;el.innerHTML=S.empresas.map(e=>`<option>${e}</option>`).join('');if(cv&&S.empresas.includes(cv))el.value=cv;});}

// ── nav ──
function showScreen(id){document.querySelectorAll('.screen').forEach(s=>s.classList.remove('active'));document.getElementById(id).classList.add('active');}
function showLogin(){showScreen('sc-login');}
function showMoto(){showScreen('sc-moto');M.render();}
function doLogin(){
  const u=document.getElementById('lu').value.trim();
  const p=document.getElementById('lp').value;
  const e=document.getElementById('lerr');
  if(u==='gestor'&&p==='1234'){e.style.display='none';showScreen('sc-app');G.render();D.filter('today');}
  else{e.style.display='block';}
}
function doLogout(){document.getElementById('lu').value='';document.getElementById('lp').value='';showScreen('sc-login');}
function setTab(id,btn){
  document.querySelectorAll('.ntab').forEach(b=>b.classList.remove('active'));btn.classList.add('active');
  document.getElementById('tab-gestao').style.display=id==='gestao'?'flex':'none';
  document.getElementById('tab-dash').style.display=id==='dash'?'flex':'none';
  if(id==='dash')D.filter(document.getElementById('fPeriod').value);
}

// ══════════════════════════════════════════════════════
//  CHARTS
// ══════════════════════════════════════════════════════
let CH={};
const TT={backgroundColor:'#1c2331',borderColor:'#30363d',borderWidth:1,titleColor:'#c9d1d9',bodyColor:'#8b949e',padding:10,cornerRadius:6};
const SCALES={x:{ticks:{color:'#8b949e',font:{size:10}},grid:{color:'rgba(255,255,255,0.04)'}},y:{ticks:{color:'#8b949e',font:{size:10}},grid:{color:'rgba(255,255,255,0.04)'},beginAtZero:true}};
function initCharts(){
  CH.hr=new Chart(document.getElementById('hrChart'),{type:'bar',data:{labels:Array.from({length:24},(_,i)=>`${i}h`),datasets:[{label:'Caminhões',data:Array(24).fill(0),backgroundColor:'#e6a81788',borderColor:'#e6a817',borderWidth:0,borderRadius:3}]},options:{responsive:true,maintainAspectRatio:true,plugins:{legend:{display:false},tooltip:TT},scales:SCALES}});
  function mkBar(id,horiz){return new Chart(document.getElementById(id),{type:'bar',data:{labels:[],datasets:[{data:[],backgroundColor:[],borderRadius:3,borderWidth:0}]},options:{responsive:true,maintainAspectRatio:true,indexAxis:horiz?'y':'x',plugins:{legend:{display:false},tooltip:TT},scales:SCALES}});}
  CH.vol=mkBar('cVol',true);CH.eff=mkBar('cEff',true);CH.trk=mkBar('cTrk',false);
  CH.dist=new Chart(document.getElementById('cDist'),{type:'doughnut',data:{labels:[],datasets:[{data:[],backgroundColor:[],borderWidth:2,borderColor:'#161b22',hoverBorderColor:'#1c2331'}]},options:{responsive:true,cutout:'65%',plugins:{legend:{position:'right',labels:{color:'#8b949e',font:{size:10},boxWidth:10,padding:8}},tooltip:TT}}});
  CH.line=new Chart(document.getElementById('cLine'),{type:'line',data:{labels:[],datasets:[]},options:{responsive:true,maintainAspectRatio:true,plugins:{legend:{display:false},tooltip:TT},scales:SCALES,elements:{line:{tension:.4,borderWidth:2},point:{radius:3,hoverRadius:5}}}});
}

// ══════════════════════════════════════════════════════
//  G — GESTÃO
// ══════════════════════════════════════════════════════
const G={
  toggleForm(which){
    const fd=document.getElementById('formDupla');
    const fm=document.getElementById('formMoto');
    if(which==='dupla'){fd.style.display=fd.style.display==='none'?'block':'none';fm.style.display='none';}
    else{fm.style.display=fm.style.display==='none'?'block':'none';fd.style.display='none';}
  },
  addDupla(){
    const nomes=document.getElementById('dNomes').value.trim();
    const pkgs=parseInt(document.getElementById('dPkgs').value)||0;
    if(!nomes){alert('Informe os nomes.');return;}
    const c=color(nomes);
    const d={id:Date.now(),nomes,pkgs,color:c,tipo:'dupla'};
    S.pool_duplas.push(d);
    S.duplas_all.push(d);
    document.getElementById('dNomes').value='';document.getElementById('dPkgs').value='';
    document.getElementById('formDupla').style.display='none';
    addLog('reg',`Dupla "${nomes}" cadastrada (${pkgs.toLocaleString()} pac.)`);
    sv();this.render();
  },
  addMoto(){
    const nome=document.getElementById('mNome').value.trim();
    const emp=document.getElementById('mEmp').value;
    const pkgs=parseInt(document.getElementById('mPkgs').value)||0;
    if(!nome||!pkgs){alert('Preencha nome e pacotes.');return;}
    const c=color(nome);
    const m={id:Date.now(),nome,empresa:emp,pacotes:pkgs,color:c,tipo:'motorista'};
    S.pool_motos.push(m);
    // também entra na fila de espera
    S.fila.push({id:m.id,nome:m.nome,empresa:m.empresa,pacotes:m.pacotes});
    document.getElementById('mNome').value='';document.getElementById('mPkgs').value='';
    document.getElementById('formMoto').style.display='none';
    addLog('reg',`${nome} (${emp}) cadastrado — ${pkgs.toLocaleString()} pac.`);
    sv();this.render();
  },
  addCarreta(){
    const doca=S.docas.find(d=>d.status==='livre');
    if(!doca){alert('Nenhuma doca livre.');return;}
    doca.status='carreta';doca.inicio=new Date().toISOString();
    addLog('sys',`Carreta → Doca ${doca.id}`);sv();this.render();
  },
  setMeta(){const v=parseInt(document.getElementById('metaInput').value);if(v>0){S.meta=v;sv();this.render();}},

  // ─── ALOCAR via drag-drop ───
  alocar(itemId,itemTipo,docaId){
    const doca=S.docas.find(d=>d.id===docaId);
    if(!doca||doca.status==='carreta')return;
    if(itemTipo==='dupla'){
      if(doca.dupla){alert('Esta doca já tem uma dupla alocada.');return;}
      const idx=S.pool_duplas.findIndex(x=>x.id===itemId);
      if(idx<0)return;
      const dup=S.pool_duplas.splice(idx,1)[0];
      doca.dupla=dup;
      if(!doca.inicio)doca.inicio=new Date().toISOString();
      doca.status=doca.motorista?'both':'dup';
      addLog('call',`Dupla "${dup.nomes}" → Doca ${docaId}`);
    } else {
      if(doca.motorista){alert('Esta doca já tem um motorista alocado.');return;}
      const idx=S.pool_motos.findIndex(x=>x.id===itemId);
      if(idx<0)return;
      const mot=S.pool_motos.splice(idx,1)[0];
      // remove da fila de espera
      S.fila=S.fila.filter(x=>x.id!==mot.id);
      doca.motorista=mot;
      if(!doca.inicio)doca.inicio=new Date().toISOString();
      doca.status=doca.dupla?'both':'mot';
      addLog('call',`${mot.nome} → Doca ${docaId}`);
      // sync portal motorista
      if(S.motorista&&S.motorista.id===mot.id){S.motorista.status='chamado';S.motorista.docaId=docaId;sv();M.render();}
    }
    sv();this.render();
  },

  // ─── FINALIZAR DOCA ───
  finalizarDoca(id){
    const d=S.docas.find(x=>x.id===id);if(!d)return;
    const ini=d.inicio||new Date().toISOString();
    const fim=new Date().toISOString();
    const hr=new Date(ini).getHours();
    S.historico[hr]=(S.historico[hr]||0)+1;
    // Registrar ops
    if(d.dupla){
      const pkgs=d.dupla.pkgs||0;
      S.ops.push({id:Date.now()+1,team:d.dupla.nomes,tipo:'dupla',startTime:ini,endTime:fim,packages:pkgs});
      S.desc+=pkgs;
      // Retornar dupla ao pool
      d.dupla.docaId=null;S.pool_duplas.push(d.dupla);
      addLog('done',`Dupla "${d.dupla.nomes}" — Doca ${id} — +${pkgs.toLocaleString()} pac.`);
    }
    if(d.motorista){
      const pkgs=d.motorista.pacotes||0;
      S.ops.push({id:Date.now()+2,team:d.motorista.nome,tipo:'motorista',startTime:ini,endTime:fim,packages:pkgs});
      S.desc+=pkgs;
      addLog('done',`${d.motorista.nome} — Doca ${id} — +${pkgs.toLocaleString()} pac.`);
      if(S.motorista&&S.motorista.id===d.motorista.id){S.motorista.status='fim';}
    }
    if(d.status==='carreta'){addLog('done',`Carreta finalizada — Doca ${id}.`);}
    d.status='livre';d.dupla=null;d.motorista=null;d.inicio=null;
    sv();this.render();D.filter(document.getElementById('fPeriod').value);
  },

  // ─── REMOVER SÓ DUPLA OU SÓ MOTO DA DOCA ───
  removerDupla(docaId){
    const d=S.docas.find(x=>x.id===docaId);if(!d||!d.dupla)return;
    S.pool_duplas.push(d.dupla);d.dupla=null;
    d.status=d.motorista?'mot':'livre';
    if(d.status==='livre')d.inicio=null;
    addLog('sys',`Dupla removida da Doca ${docaId}`);sv();this.render();
  },
  removerMoto(docaId){
    const d=S.docas.find(x=>x.id===docaId);if(!d||!d.motorista)return;
    S.pool_motos.push(d.motorista);
    S.fila.push({id:d.motorista.id,nome:d.motorista.nome,empresa:d.motorista.empresa,pacotes:d.motorista.pacotes});
    d.motorista=null;
    d.status=d.dupla?'dup':'livre';
    if(d.status==='livre')d.inicio=null;
    addLog('sys',`Motorista removido da Doca ${docaId}`);sv();this.render();
  },

  // ─── RENDER ───
  render(){
    // KPIs topo
    const ativos=S.docas.filter(d=>d.status!=='livre');
    const comDupla=S.docas.filter(d=>d.dupla).length;
    const comMoto=S.docas.filter(d=>d.motorista).length;
    const pct=Math.min(100,Math.round(S.desc/S.meta*100));
    document.getElementById('kDesc').textContent=S.desc.toLocaleString('pt-BR');
    document.getElementById('kAtivas').textContent=ativos.length;
    document.getElementById('kDuplas').textContent=comDupla;
    document.getElementById('kMoots').textContent=comMoto;
    document.getElementById('kFila').textContent=S.fila.length;
    document.getElementById('kPct').textContent=pct+'%';
    document.getElementById('kPctSub').textContent=`${S.desc.toLocaleString()} / ${S.meta.toLocaleString()}`;
    document.getElementById('kBar').style.width=pct+'%';
    document.getElementById('metaInput').value=S.meta;
    // Pill
    const pill=document.getElementById('tpill');
    if(ativos.length>0){pill.textContent='● OPERAÇÃO ATIVA';pill.className='tpill on';}
    else{pill.textContent='● INATIVO';pill.className='tpill off';}

    // Pool duplas
    document.getElementById('cntDuplas').textContent=S.pool_duplas.length;
    const idDuplas=document.getElementById('itemsDuplas');
    idDuplas.innerHTML=S.pool_duplas.length?S.pool_duplas.map(d=>`
      <div class="drag-item" draggable="true" data-id="${d.id}" data-tipo="dupla">
        <div class="drag-avatar" style="background:${d.color}">${initials(d.nomes)}</div>
        <div class="drag-info">
          <div class="drag-name">${d.nomes}</div>
          <div class="drag-sub">${(d.pkgs||0).toLocaleString()} pacotes</div>
        </div>
        <span class="drag-type-badge tb-dupla">Dupla</span>
        <span class="drag-handle">⠿</span>
      </div>`).join(''):'<div class="drag-empty">Nenhuma dupla<br>disponível</div>';

    // Pool motoristas
    document.getElementById('cntMotos').textContent=S.pool_motos.length;
    const idMotos=document.getElementById('itemsMotos');
    idMotos.innerHTML=S.pool_motos.length?S.pool_motos.map(m=>`
      <div class="drag-item" draggable="true" data-id="${m.id}" data-tipo="motorista">
        <div class="drag-avatar" style="background:${m.color}">${initials(m.nome)}</div>
        <div class="drag-info">
          <div class="drag-name">${m.nome}</div>
          <div class="drag-sub">${m.empresa} · ${(m.pacotes||0).toLocaleString()} pac.</div>
        </div>
        <span class="drag-type-badge tb-moto">Moto</span>
        <span class="drag-handle">⠿</span>
      </div>`).join(''):'<div class="drag-empty">Nenhum motorista<br>disponível</div>';

    // Grid de docas
    const gr=document.getElementById('docasGrid');
    gr.innerHTML='';
    S.docas.forEach(d=>{
      const el=document.createElement('div');
      const cls={livre:'',dup:'dup',mot:'mot',both:'both',carreta:'cart'}[d.status]||'';
      el.className=`doca ${cls}`;
      el.dataset.docaId=d.id;
      if(d.status==='livre'){el.dataset.droppable='true';}

      const num=`<div class="doca-num">DOCA ${d.id}</div>`;

      if(d.status==='livre'){
        el.innerHTML=`${num}<div class="doca-badge db-livre">Livre</div><div class="doca-slots"></div><div class="drop-hint">↓ arraste aqui</div>`;
      } else if(d.status==='carreta'){
        el.innerHTML=`${num}<div class="doca-badge db-cart">Carreta</div><div class="doca-timer">${fmtMs(Date.now()-new Date(d.inicio).getTime())}</div><div class="doca-actions"><button class="btn r" onclick="G.finalizarDoca(${d.id})">Finalizar</button></div>`;
      } else {
        const timer=d.inicio?fmtMs(Date.now()-new Date(d.inicio).getTime()):'00:00:00';
        const badge=d.status==='both'?'db-full':d.status==='dup'?'db-dupla':'db-moto';
        const label=d.status==='both'?'Dupla+Moto':d.status==='dup'?'Dupla':'Motorista';
        // Calcular pac/h ao vivo
        const hrsDecorridas=(Date.now()-(new Date(d.inicio).getTime()))/36e5;
        const pkgsTot=(d.dupla?.pkgs||0)+(d.motorista?.pacotes||0);
        const effLive=hrsDecorridas>0?(pkgsTot/hrsDecorridas).toFixed(0):'—';

        let slots='';
        if(d.dupla)slots+=`<div class="slot slot-dup"><span class="slot-dot" style="background:${d.dupla.color}"></span><span class="slot-name">${initials(d.dupla.nomes)}: ${d.dupla.nomes}</span><button onclick="G.removerDupla(${d.id})" style="background:none;border:none;color:var(--muted);cursor:pointer;font-size:10px;padding:0 2px;margin-left:auto" title="Remover dupla">✕</button></div>`;
        if(d.motorista)slots+=`<div class="slot slot-mot"><span class="slot-dot" style="background:${d.motorista.color}"></span><span class="slot-name">${initials(d.motorista.nome)}: ${d.motorista.nome}</span><button onclick="G.removerMoto(${d.id})" style="background:none;border:none;color:var(--muted);cursor:pointer;font-size:10px;padding:0 2px;margin-left:auto" title="Remover moto">✕</button></div>`;
        // Drop zone extra se doca tem só um
        const canDropDupla=!d.dupla;const canDropMoto=!d.motorista;
        const extraDrop=(canDropDupla||canDropMoto)&&d.status!=='carreta'?`<div class="drop-hint" style="font-size:8px">+ ${canDropDupla?'dupla':''}${canDropDupla&&canDropMoto?' / ':''}${canDropMoto?'moto':''}</div>`:'';

        el.innerHTML=`${num}<div class="doca-badge ${badge}">${label}</div><div class="doca-slots">${slots}</div><div class="doca-timer">${timer}</div><div class="doca-kpis">${effLive} pac/h · ${pkgsTot.toLocaleString()} pac</div>${extraDrop}<div class="doca-actions"><button class="btn r" onclick="G.finalizarDoca(${d.id})">✓ Finalizar</button></div>`;
        // Docas parcialmente ocupadas também aceitam drop
        if(canDropDupla||canDropMoto)el.dataset.droppable='true';
      }
      gr.appendChild(el);
    });

    // Setup drag-drop
    setTimeout(()=>{
      // Drag items
      document.querySelectorAll('.drag-item[draggable]').forEach(item=>{
        item.addEventListener('dragstart',e=>{
          e.dataTransfer.setData('id',item.dataset.id);
          e.dataTransfer.setData('tipo',item.dataset.tipo);
          item.classList.add('dragging');
        });
        item.addEventListener('dragend',()=>item.classList.remove('dragging'));
      });
      // Drop targets
      document.querySelectorAll('.doca[data-droppable]').forEach(doca=>{
        doca.addEventListener('dragover',e=>{e.preventDefault();doca.classList.add('over');});
        doca.addEventListener('dragleave',()=>doca.classList.remove('over'));
        doca.addEventListener('drop',e=>{
          e.preventDefault();doca.classList.remove('over');
          const id=parseInt(e.dataTransfer.getData('id'));
          const tipo=e.dataTransfer.getData('tipo');
          G.alocar(id,tipo,parseInt(doca.dataset.docaId));
        });
      });
    },0);

    // KPIs produtividade
    this.renderProd();

    // Fila
    const fw=document.getElementById('filaWrap');
    if(!S.fila.length){fw.innerHTML='<div style="padding:12px 14px;font-size:11px;color:var(--muted)">Nenhum motorista na fila.</div>';}
    else{fw.innerHTML=S.empresas.map(emp=>{const items=S.fila.filter(x=>x.empresa===emp);if(!items.length)return'';return`<div class="femp"><div class="femp-head">${emp}<span class="femp-cnt">${items.length}</span></div>${items.map(x=>{const pos=S.fila.findIndex(f=>f.id===x.id)+1;return`<div class="fitem"><div class="fpos">${pos}</div><div class="fnome">${x.nome}</div><div class="fpkgs">${x.pacotes.toLocaleString()}</div></div>`;}).join('')}</div>`;}).join('');}

    // Log
    const lc=document.getElementById('logWrap');
    const lco={sys:'#58a6ff',reg:'#3fb950',call:'#e6a817',done:'#bc8cff',err:'#f85149'};
    lc.innerHTML=S.log.length?S.log.map(e=>`<div class="lentry"><div class="lt">${e.h}</div><div class="ld" style="background:${lco[e.tipo]||'#6e7681'}"></div><div class="lm">${e.msg}</div></div>`).join(''):'<div style="font-size:11px;color:var(--muted)">Aguardando eventos...</div>';

    // Hr chart
    CH.hr.data.datasets[0].data=Array.from({length:24},(_,i)=>S.historico[i]||0);CH.hr.update();
  },

  renderProd(){
    const ativas=S.docas.filter(d=>d.status!=='livre'&&d.status!=='carreta');
    const pg=document.getElementById('prodGrid');
    if(!ativas.length){pg.innerHTML='<div style="padding:14px;font-size:11px;color:var(--muted);text-align:center;grid-column:1/-1">Nenhuma doca ativa no momento.</div>';return;}

    // Calcular max eficiência para barra relativa
    const effs=ativas.map(d=>{
      const ms=Date.now()-new Date(d.inicio).getTime();
      const h=ms/3600000;
      const p=(d.dupla?.pkgs||0)+(d.motorista?.pacotes||0);
      return h>0?p/h:0;
    });
    const maxEff=Math.max(...effs,1);
    // Média geral para comparar
    const mediaEff=effs.reduce((a,b)=>a+b,0)/effs.length;

    pg.innerHTML=ativas.map((d,i)=>{
      const ms=Date.now()-new Date(d.inicio).getTime();
      const hDecorridas=ms/3600000;
      const minutos=Math.floor(ms/60000);
      const pkgsTot=(d.dupla?.pkgs||0)+(d.motorista?.pacotes||0);
      const eff=hDecorridas>0?pkgsTot/hDecorridas:0;
      const pct=Math.min(100,(eff/maxEff)*100);
      // Cor da barra: verde se acima da média, amarelo se próximo, vermelho se abaixo
      const barColor=eff>=mediaEff*1.1?'var(--green)':eff>=mediaEff*0.8?'var(--acc)':'var(--red)';
      const ratingTxt=eff>=mediaEff*1.1?'Acima da média':eff>=mediaEff*0.8?'Na média':'Abaixo da média';
      const ratingColor=eff>=mediaEff*1.1?'var(--green)':eff>=mediaEff*0.8?'var(--acc)':'var(--red)';
      // Nome principal
      const nome=d.dupla?d.dupla.nomes:(d.motorista?d.motorista.nome:'Carreta');
      const avatarColor=d.dupla?d.dupla.color:d.motorista?.color||'#6e7681';
      const avInit=initials(nome);
      // Segunda linha (motorista se tiver dupla+moto)
      const sub2=d.dupla&&d.motorista?`<div style="font-size:9px;color:var(--blue);margin-top:2px">+ ${d.motorista.nome}</div>`:'';

      return`<div class="prod-card">
        <div class="prod-header">
          <div class="prod-avatar" style="background:${avatarColor}">${avInit}</div>
          <div style="flex:1;min-width:0">
            <div class="prod-name">${nome}</div>
            ${sub2}
          </div>
          <div class="prod-doca">D${d.id}</div>
        </div>
        <div class="prod-metrics">
          <div class="pm"><div class="pm-v" style="color:var(--acc)">${eff>0?Math.round(eff):'—'}</div><div class="pm-l">pac/h</div></div>
          <div class="pm"><div class="pm-v" style="color:var(--green)">${pkgsTot.toLocaleString()}</div><div class="pm-l">pacotes</div></div>
          <div class="pm"><div class="pm-v" style="color:var(--blue)">${minutos}m</div><div class="pm-l">tempo</div></div>
        </div>
        <div class="prod-bar-wrap">
          <div style="display:flex;justify-content:space-between;font-size:9px;color:var(--muted);margin-bottom:2px"><span>Eficiência relativa</span><span>${pct.toFixed(0)}%</span></div>
          <div class="prod-bar"><div class="prod-bar-fill" style="width:${pct}%;background:${barColor}"></div></div>
        </div>
        <div class="prod-status"><span class="pst-dot" style="background:${ratingColor}"></span>${ratingTxt}</div>
      </div>`;
    }).join('');
  }
};

// ══════════════════════════════════════════════════════
//  D — DASHBOARD
// ══════════════════════════════════════════════════════
const D={
  filter(val){
    const now=new Date(),td=new Date(now.getFullYear(),now.getMonth(),now.getDate());
    let st,en;
    switch(val){
      case'today':st=td;en=new Date(td.getTime()+864e5);break;
      case'yesterday':st=new Date(td.getTime()-864e5);en=td;break;
      case'week':st=new Date(td);st.setDate(st.getDate()-now.getDay());en=new Date(now.getTime()+864e5);break;
      case'month':st=new Date(now.getFullYear(),now.getMonth(),1);en=new Date(now.getTime()+864e5);break;
      default:S.filteredOps=[...S.ops];this.render();return;
    }
    S.filteredOps=S.ops.filter(o=>{const d=new Date(o.startTime);return d>=st&&d<en;});
    this.render();
  },
  addOp(){
    const t=document.getElementById('opTeam').value.trim();
    const s=document.getElementById('opStart').value;
    const e=document.getElementById('opEnd').value;
    const p=parseInt(document.getElementById('opPkgs').value)||0;
    if(!t||!s||!e||!p){alert('Preencha todos os campos.');return;}
    if(new Date(e)<=new Date(s)){alert('Término deve ser após o início.');return;}
    S.ops.push({id:Date.now(),team:t,tipo:'manual',startTime:new Date(s).toISOString(),endTime:new Date(e).toISOString(),packages:p});
    color(t);document.getElementById('opTeam').value='';document.getElementById('opPkgs').value='';
    sv();this.filter(document.getElementById('fPeriod').value);this.fillDL();
  },
  fillDL(){const dl=document.getElementById('teamsDL');dl.innerHTML=[...new Set(S.ops.map(o=>o.team))].map(t=>`<option value="${t}">`).join('');},
  metrics(){
    const teams=[...new Set(S.filteredOps.map(o=>o.team))];
    const m={};
    teams.forEach(t=>{
      const ops=S.filteredOps.filter(o=>o.team===t);
      const p=ops.reduce((a,o)=>a+o.packages,0);
      const h=ops.reduce((a,o)=>a+(new Date(o.endTime)-new Date(o.startTime))/36e5,0);
      const tipo=ops[0]?.tipo||'manual';
      m[t]={p,trucks:ops.length,h,eff:h>0?p/h:0,avgH:ops.length?h/ops.length:0,tipo};
    });
    return m;
  },
  render(){
    const m=this.metrics();const teams=Object.keys(m);const cols=teams.map(t=>color(t));
    const tot=Object.values(m).reduce((a,x)=>a+x.p,0);
    const trk=Object.values(m).reduce((a,x)=>a+x.trucks,0);
    const avg=teams.length?Object.values(m).reduce((a,x)=>a+x.eff,0)/teams.length:0;
    const avgT=teams.length?Object.values(m).reduce((a,x)=>a+x.avgH,0)/teams.length:0;
    document.getElementById('dTot').textContent=tot.toLocaleString('pt-BR');
    document.getElementById('dTrk').textContent=trk;
    document.getElementById('dTeams').textContent=teams.length;
    document.getElementById('dAvg').textContent=avg.toFixed(0);
    document.getElementById('dAvgT').textContent=avgT>0?`${(avgT*60).toFixed(0)}min`:'—';
    const srt=Object.entries(m).sort((a,b)=>b[1].eff-a[1].eff);
    if(srt.length>=2){
      document.getElementById('dFast').textContent=`▲ ${srt[0][0].split(' ')[0]} ${srt[0][1].eff.toFixed(0)}p/h`;
      document.getElementById('dSlow').textContent=`▼ ${srt[srt.length-1][0].split(' ')[0]} ${srt[srt.length-1][1].eff.toFixed(0)}p/h`;
    }else{document.getElementById('dFast').textContent='—';document.getElementById('dSlow').textContent='—';}

    const srtV=Object.entries(m).sort((a,b)=>b[1].p-a[1].p);
    const srtE=Object.entries(m).sort((a,b)=>b[1].eff-a[1].eff);
    const initsV=srtV.map(([t])=>initials(t));const initsE=srtE.map(([t])=>initials(t));const inits=teams.map(t=>initials(t));

    function upd(ch,labels,data,bgs){ch.data.labels=labels;ch.data.datasets[0].data=data;ch.data.datasets[0].backgroundColor=bgs;ch.update();}
    upd(CH.vol,initsV,srtV.map(([,v])=>v.p),srtV.map(([t])=>color(t)+'cc'));
    upd(CH.eff,initsE,srtE.map(([,v])=>+v.eff.toFixed(1)),srtE.map(([t])=>color(t)+'cc'));
    upd(CH.trk,inits,teams.map(t=>m[t].trucks),cols.map(c=>c+'cc'));
    CH.dist.data.labels=inits;CH.dist.data.datasets[0].data=teams.map(t=>m[t].p);CH.dist.data.datasets[0].backgroundColor=cols;CH.dist.update();

    const hrs=Array.from({length:11},(_,i)=>`${String(i+13).padStart(2,'0')}:00`);
    const ds=teams.map(t=>({label:initials(t),data:hrs.map((_,idx)=>{const h=idx+13;return S.filteredOps.filter(o=>o.team===t&&new Date(o.endTime).getHours()<h+1).reduce((a,o)=>a+o.packages,0);}),borderColor:color(t),backgroundColor:color(t)+'22',fill:false,pointBackgroundColor:color(t)}));
    CH.line.data.labels=hrs;CH.line.data.datasets=ds;CH.line.update();
    document.getElementById('lineLegend').innerHTML=teams.map(t=>`<div class="ch-li"><div class="ch-sw" style="background:${color(t)}"></div><span>${initials(t)}: ${t}</span></div>`).join('');

    const tb=document.querySelector('#opsTable tbody');
    if(!S.filteredOps.length){tb.innerHTML=`<tr><td colspan="8" style="text-align:center;color:var(--muted);padding:16px">Nenhuma operação encontrada.</td></tr>`;return;}
    const allEff=Object.values(m).map(x=>x.eff).sort((a,b)=>b-a);
    const q1=allEff[0]||1,q3=allEff[Math.floor(allEff.length*2/3)]||0;
    tb.innerHTML=[...S.filteredOps].sort((a,b)=>new Date(b.startTime)-new Date(a.startTime)).map(op=>{
      const dur=(new Date(op.endTime)-new Date(op.startTime))/36e5;
      const eff=dur>0?op.packages/dur:0;
      const c=color(op.team);
      const rating=eff>=q1*0.9?`<span class="badge g">A+</span>`:eff>=q3?`<span class="badge a">B</span>`:`<span class="badge r">C</span>`;
      const tipoBadge=op.tipo==='dupla'?`<span class="badge a">Dupla</span>`:op.tipo==='motorista'?`<span class="badge" style="background:rgba(88,166,255,.15);color:var(--blue);border:1px solid rgba(88,166,255,.2)">Motorista</span>`:`<span style="font-size:10px;color:var(--muted)">Manual</span>`;
      return`<tr><td><span style="width:18px;height:18px;border-radius:3px;background:${c};display:inline-flex;align-items:center;justify-content:center;font-size:8px;font-weight:700;color:#000;margin-right:6px;vertical-align:middle">${initials(op.team)}</span>${op.team}</td><td>${tipoBadge}</td><td class="mono">${fdt(op.startTime)}</td><td class="mono">${fdt(op.endTime)}</td><td class="mono">${op.packages.toLocaleString('pt-BR')}</td><td class="mono">${dur.toFixed(2)}</td><td><span class="badge g">${eff.toFixed(1)} p/h</span></td><td>${rating}</td></tr>`;
    }).join('');
    this.fillDL();
  },
  exportar(){
    const data=S.filteredOps.map(op=>{const dur=(new Date(op.endTime)-new Date(op.startTime))/36e5;return{'Nome':op.team,'Tipo':op.tipo||'—','Início':fdt(op.startTime),'Término':fdt(op.endTime),'Pacotes':op.packages,'Duração(h)':dur.toFixed(2),'Eficiência':dur>0?(op.packages/dur).toFixed(1):'0'};});
    const ws=XLSX.utils.json_to_sheet(data);const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,'Operações');
    XLSX.writeFile(wb,`operacoes-${new Date().toISOString().slice(0,10)}.xlsx`);
  },
  limpar(){document.getElementById('modalOk').classList.add('open');},
  confirmar(){S.ops=[];S.filteredOps=[];S.teamColors={};S.cidx=0;closeModal();sv();this.render();}
};

// ══════════════════════════════════════════════════════
//  M — MOTORISTA PORTAL
// ══════════════════════════════════════════════════════
const M={
  addEmpresa(){
    const n=document.getElementById('mfNovaEmp').value.trim();
    if(!n)return;
    if(!S.empresas.includes(n))S.empresas.push(n);
    document.getElementById('mfNovaEmp').value='';
    fillEmpSels();sv();document.getElementById('mfEmp').value=n;
  },
  registrar(){
    const nome=document.getElementById('mfNome').value.trim();
    const wpp=document.getElementById('mfWpp').value.trim();
    const emp=document.getElementById('mfEmp').value;
    const pkgs=parseInt(document.getElementById('mfPkgs').value)||0;
    if(!nome||!pkgs){alert('Preencha nome e pacotes.');return;}
    const c=color(nome);
    const mot={id:Date.now(),nome,wpp,empresa:emp,pacotes:pkgs,status:'fila',docaId:null,ini:null,color:c};
    S.motorista=mot;
    // Adiciona ao pool e à fila
    S.pool_motos.push({id:mot.id,nome:mot.nome,empresa:mot.empresa,pacotes:mot.pacotes,color:c,tipo:'motorista'});
    S.fila.push({id:mot.id,nome:mot.nome,empresa:mot.empresa,pacotes:mot.pacotes});
    addLog('reg',`${nome} registrou chegada (portal).`);
    sv();G.render();this.render();
  },
  iniciar(){
    const m=S.motorista;if(!m||m.status!=='chamado')return;
    m.status='desc';m.ini=new Date().toISOString();
    // Atualiza início da doca
    const doca=S.docas.find(d=>d.motorista&&d.motorista.id===m.id);
    if(doca&&!doca.inicio)doca.inicio=m.ini;
    addLog('sys',`${m.nome} iniciou — Doca ${m.docaId}.`);sv();this.render();G.render();
  },
  finalizar(){
    const m=S.motorista;if(!m||m.status!=='desc')return;
    m.status='fim';
    const doca=S.docas.find(d=>d.motorista&&d.motorista.id===m.id);
    if(doca)G.finalizarDoca(doca.id);
    else{
      if(m.ini){const fim=new Date().toISOString();S.ops.push({id:Date.now(),team:m.nome,tipo:'motorista',startTime:m.ini,endTime:fim,packages:m.pacotes});S.desc+=m.pacotes;}
    }
    addLog('done',`${m.nome} finalizou descarga.`);sv();this.render();
  },
  novo(){
    S.motorista=null;['mfNome','mfWpp','mfPkgs'].forEach(id=>document.getElementById(id).value='');sv();this.render();
  },
  updTimer(){
    if(!S.motorista||S.motorista.status!=='desc')return;
    const el=document.getElementById('mTimer');if(!el)return;
    el.textContent=fmtMs(Date.now()-new Date(S.motorista.ini).getTime());
  },
  render(){
    const m=S.motorista;
    document.getElementById('mForm').style.display=m?'none':'block';
    document.getElementById('mStatus').style.display=m?'block':'none';
    if(!m)return;
    document.getElementById('mNomeDisp').textContent=`${m.nome} · ${m.empresa} · ${m.pacotes.toLocaleString()} pac.`;
    ['msWait','msCall','msDesc','msFim'].forEach(id=>document.getElementById(id).style.display='none');
    if(m.status==='fila'){document.getElementById('msWait').style.display='block';const pos=S.fila.findIndex(x=>x.id===m.id)+1;document.getElementById('mPos').textContent=pos>0?pos:'—';}
    else if(m.status==='chamado'){document.getElementById('msCall').style.display='block';document.getElementById('mDocaNum').textContent=`DOCA ${m.docaId}`;}
    else if(m.status==='desc'){document.getElementById('msDesc').style.display='block';this.updTimer();}
    else if(m.status==='fim'){document.getElementById('msFim').style.display='block';}
  }
};

// ══════════════════════════════════════════════════════
//  INIT
// ══════════════════════════════════════════════════════
document.addEventListener('DOMContentLoaded',()=>{
  ld();initCharts();fillEmpSels();D.fillDL();
  setInterval(()=>document.getElementById('clock').textContent=new Date().toLocaleTimeString('pt-BR'),1000);
  setInterval(()=>{
    if(document.getElementById('sc-app').classList.contains('active'))G.render();
    if(document.getElementById('sc-moto').classList.contains('active')){
      M.updTimer();
      if(S.motorista?.status==='fila'){const pos=S.fila.findIndex(x=>x.id===S.motorista.id)+1;const el=document.getElementById('mPos');if(el)el.textContent=pos>0?pos:'—';}
    }
  },1000);
});
</script>
</body>
</html>

