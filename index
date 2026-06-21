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
  --bg:#0d1117;--s1:#161b22;--s2:#1c2331;
  --b1:#30363d;--b2:#444c56;--b3:#6e7681;
  --tx:#c9d1d9;--muted:#8b949e;
  --acc:#e6a817;--acc2:#b8820e;
  --green:#3fb950;--red:#f85149;--blue:#58a6ff;--purple:#bc8cff;--orange:#ffa657;
}
body{background:var(--bg);color:var(--tx);font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;font-size:13px;line-height:1.5;min-height:100vh}
.screen{display:none;min-height:100vh}.screen.active{display:flex;flex-direction:column}
/* LOGIN */
#sc-login{align-items:center;justify-content:center}
.lcard{background:var(--s1);border:1px solid var(--b1);border-radius:10px;padding:40px 44px;width:360px}
.lbrand{display:flex;align-items:center;gap:10px;margin-bottom:32px}
.lbrand .ico{width:36px;height:36px;background:var(--acc);border-radius:7px;display:flex;align-items:center;justify-content:center}
.lbrand .ico svg{width:18px;height:18px}
.lbrand h1{font-size:15px;font-weight:700;color:#fff}.lbrand p{font-size:11px;color:var(--muted)}
.lf{margin-bottom:14px}
.lf label{display:block;font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.07em;color:var(--muted);margin-bottom:5px}
.lf input{width:100%;background:var(--bg);border:1px solid var(--b1);border-radius:6px;padding:9px 12px;color:var(--tx);font-size:13px;outline:none;transition:border-color .15s}
.lf input:focus{border-color:var(--acc)}
.lerr{font-size:11px;color:var(--red);margin-top:6px;display:none}
.lbtn{width:100%;background:var(--acc);border:none;border-radius:6px;padding:11px;color:#000;font-weight:700;font-size:13px;cursor:pointer;margin-top:4px}
.lbtn:hover{background:var(--acc2)}
.lmoto{margin-top:20px;padding-top:16px;border-top:1px solid var(--b1);text-align:center;font-size:11px;color:var(--muted)}
.lmoto a{color:var(--acc);text-decoration:none}
/* TOPBAR */
.topbar{background:var(--s1);border-bottom:1px solid var(--b1);height:44px;padding:0 20px;display:flex;align-items:center;flex-shrink:0;position:sticky;top:0;z-index:100}
.tbrand{display:flex;align-items:center;gap:8px;padding-right:20px;border-right:1px solid var(--b1);margin-right:4px}
.tbrand .ico{width:24px;height:24px;background:var(--acc);border-radius:4px;display:flex;align-items:center;justify-content:center}
.tbrand .ico svg{width:12px;height:12px}
.tbrand span{font-size:13px;font-weight:700;color:#fff}
.navtabs{display:flex;flex:1;padding-left:4px;overflow-x:auto}
.ntab{padding:0 14px;height:44px;display:flex;align-items:center;gap:6px;font-size:12px;font-weight:500;color:var(--muted);cursor:pointer;border:none;border-bottom:2px solid transparent;background:none;transition:all .15s;white-space:nowrap;flex-shrink:0}
.ntab:hover{color:var(--tx)}.ntab.active{color:var(--acc);border-bottom-color:var(--acc)}
.tright{display:flex;align-items:center;gap:8px;margin-left:auto;flex-shrink:0}
.tclock{font-family:'SF Mono','Consolas',monospace;font-size:12px;color:var(--muted);background:var(--bg);border:1px solid var(--b1);padding:3px 9px;border-radius:4px}
.tpill{font-size:10px;font-weight:700;padding:3px 8px;border-radius:3px}
.tpill.on{background:rgba(63,185,80,.12);color:var(--green);border:1px solid rgba(63,185,80,.25)}
.tpill.off{background:rgba(248,81,73,.1);color:var(--red);border:1px solid rgba(248,81,73,.25)}
.tbtn{background:none;border:1px solid var(--b1);border-radius:4px;padding:4px 9px;color:var(--muted);font-size:11px;cursor:pointer}
.tbtn:hover{border-color:var(--b2);color:var(--tx)}
/* LAYOUT */
.mcontent{flex:1;overflow:auto;padding:14px 18px;display:flex;flex-direction:column;gap:12px}
/* PANEL */
.panel{background:var(--s1);border:1px solid var(--b1);border-radius:7px;overflow:hidden}
.phead{padding:9px 14px;border-bottom:1px solid var(--b1);display:flex;align-items:center;justify-content:space-between;gap:8px;flex-wrap:wrap}
.ptitle{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.08em;color:var(--muted);display:flex;align-items:center;gap:6px}
.dot{width:6px;height:6px;border-radius:50%;display:inline-block;flex-shrink:0}
.dg{background:var(--green)}.da{background:var(--acc)}.db{background:var(--blue)}.dr{background:var(--red)}.dm{background:var(--b3)}.dp{background:var(--purple)}
/* KPI */
.kstrip{display:grid;grid-template-columns:repeat(auto-fit,minmax(110px,1fr));gap:1px;background:var(--b1);border:1px solid var(--b1);border-radius:7px;overflow:hidden}
.kpi{background:var(--s1);padding:11px 14px}
.kl{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.07em;color:var(--muted);margin-bottom:3px}
.kv{font-size:22px;font-weight:800;color:#fff;font-family:'SF Mono','Consolas',monospace;line-height:1}
.kv.g{color:var(--green)!important}.kv.a{color:var(--acc)!important}.kv.b{color:var(--blue)!important}.kv.r{color:var(--red)!important}.kv.p{color:var(--purple)!important}
.ks{font-size:10px;color:var(--muted);margin-top:2px}
.pbar{height:5px;background:var(--b1);border-radius:3px;overflow:hidden;margin-top:5px}
.pfill{height:100%;background:linear-gradient(90deg,var(--acc2),var(--acc));border-radius:3px;transition:width .4s}
/* BUTTONS */
.btn{padding:5px 12px;border-radius:5px;border:1px solid var(--b1);background:var(--s2);color:var(--tx);font-size:11px;font-weight:600;cursor:pointer;transition:all .15s;white-space:nowrap;display:inline-flex;align-items:center;gap:5px}
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
/* DRAG LISTS */
.drag-list{border:1px solid var(--b1);border-radius:6px;background:var(--bg);overflow:hidden;display:flex;flex-direction:column}
.dlhead{padding:8px 12px;background:var(--s2);border-bottom:1px solid var(--b1);font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.07em;color:var(--muted);display:flex;align-items:center;justify-content:space-between}
.dlhead .cnt{font-family:'SF Mono','Consolas',monospace;font-size:11px;color:var(--acc);font-weight:800}
.dl-items{overflow-y:auto}
.di{padding:8px 12px;border-bottom:1px solid rgba(48,54,61,.5);display:flex;align-items:center;gap:9px;cursor:grab;user-select:none;transition:background .12s}
.di:last-child{border-bottom:none}
.di:hover{background:var(--s2)}
.di.dragging{opacity:.3;cursor:grabbing}
.di-av{width:28px;height:28px;border-radius:5px;display:flex;align-items:center;justify-content:center;font-size:9px;font-weight:700;color:#000;flex-shrink:0;font-family:'SF Mono','Consolas',monospace}
.di-info{flex:1;min-width:0}
.di-name{font-size:11px;font-weight:600;color:var(--tx);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.di-sub{font-size:9px;color:var(--muted);margin-top:1px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.di-handle{color:var(--b3);font-size:13px;flex-shrink:0}
.di-badge{font-size:8px;font-weight:700;padding:2px 6px;border-radius:3px;flex-shrink:0}
.bdup{background:rgba(230,168,23,.15);color:var(--acc)}.bmot{background:rgba(88,166,255,.15);color:var(--blue)}
.bcart{background:rgba(248,81,73,.15);color:var(--red)}
.dl-empty{padding:20px 12px;font-size:10px;color:var(--muted);text-align:center;font-style:italic;line-height:1.7}
/* ALOC */
.aloc-pools{display:grid;grid-template-columns:1fr 1fr 1fr;gap:10px;padding:12px 14px 8px}
.aloc-hint{padding:0 14px 10px;font-size:10px;color:var(--muted);display:flex;align-items:center;gap:14px;flex-wrap:wrap}
/* DOCAS */
.docas-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(130px,1fr));gap:6px;padding:0 14px 14px}
.doca{border:1px solid var(--b1);border-radius:6px;padding:9px 8px;min-height:138px;display:flex;flex-direction:column;align-items:center;text-align:center;gap:3px;transition:border-color .15s,background .15s;background:var(--bg);position:relative}
.doca.over{background:rgba(230,168,23,.07);border-color:var(--acc);border-style:dashed}
.doca.s-dup{background:rgba(63,185,80,.04);border-color:var(--green)}
.doca.s-mot{background:rgba(88,166,255,.04);border-color:var(--blue)}
.doca.s-both{background:rgba(188,140,255,.04);border-color:var(--purple)}
.doca.s-cart{background:rgba(248,81,73,.04);border-color:var(--red)}
@keyframes pg{0%,100%{border-color:var(--green)}50%{border-color:rgba(63,185,80,.25)}}
@keyframes pb{0%,100%{border-color:var(--blue)}50%{border-color:rgba(88,166,255,.25)}}
@keyframes pp{0%,100%{border-color:var(--purple)}50%{border-color:rgba(188,140,255,.25)}}
.doca.s-dup{animation:pg 2.5s infinite}.doca.s-mot{animation:pb 2.5s infinite}.doca.s-both{animation:pp 2.5s infinite}
.doca-num{font-family:'SF Mono','Consolas',monospace;font-size:9px;font-weight:700;color:var(--muted);letter-spacing:.04em}
.doca-badge{font-size:8px;font-weight:700;text-transform:uppercase;letter-spacing:.05em;padding:2px 6px;border-radius:3px}
.db-livre{background:rgba(110,118,129,.1);color:var(--muted)}
.db-dup{background:rgba(63,185,80,.15);color:var(--green)}.db-mot{background:rgba(88,166,255,.15);color:var(--blue)}
.db-both{background:rgba(188,140,255,.15);color:var(--purple)}.db-cart{background:rgba(248,81,73,.15);color:var(--red)}
.doca-slots{display:flex;flex-direction:column;gap:2px;width:100%}
.slot{font-size:9px;padding:3px 5px;border-radius:3px;display:flex;align-items:center;gap:4px;width:100%;text-align:left}
.slot-g{background:rgba(63,185,80,.1);border:1px solid rgba(63,185,80,.2);color:var(--green)}
.slot-b{background:rgba(88,166,255,.1);border:1px solid rgba(88,166,255,.2);color:var(--blue)}
.slot-r{background:rgba(248,81,73,.1);border:1px solid rgba(248,81,73,.2);color:var(--red)}
.slot-dot{width:5px;height:5px;border-radius:50%;flex-shrink:0}
.slot-name{flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;color:var(--tx)}
.slot-rm{background:none;border:none;color:var(--muted);cursor:pointer;font-size:10px;padding:0 1px;line-height:1;margin-left:auto}
.slot-rm:hover{color:var(--red)}
.doca-timer{font-family:'SF Mono','Consolas',monospace;font-size:11px;font-weight:700;color:var(--acc)}
.doca-eff{font-size:9px;color:var(--muted)}
.doca-drop-hint{font-size:8px;color:var(--b3);margin-top:2px}
.doca-actions{display:flex;flex-direction:column;gap:3px;width:100%;margin-top:auto}
.doca-actions .btn{font-size:9px;padding:3px 6px;width:100%;justify-content:center}
/* FILA */
.fila-scroll{overflow-x:auto;padding:14px 16px 18px;display:flex;gap:10px;min-height:130px;align-items:flex-start}
.fila-scroll::-webkit-scrollbar{height:5px}.fila-scroll::-webkit-scrollbar-thumb{background:var(--b2);border-radius:3px}
.fila-card{background:var(--bg);border:1px solid var(--b1);border-radius:7px;padding:12px 14px;min-width:190px;max-width:210px;flex-shrink:0;cursor:grab;user-select:none;display:flex;flex-direction:column;gap:8px;transition:background .12s,border-color .12s,transform .1s}
.fila-card:hover{background:var(--s2);border-color:var(--b2);transform:translateY(-2px)}
.fila-card.dragging{opacity:.3;cursor:grabbing}
.fc-top{display:flex;align-items:center;gap:10px}
.fc-av{width:32px;height:32px;border-radius:6px;display:flex;align-items:center;justify-content:center;font-size:10px;font-weight:700;color:#000;flex-shrink:0;font-family:'SF Mono','Consolas',monospace}
.fc-nome{font-size:12px;font-weight:700;color:var(--tx);line-height:1.3;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.fc-emp{font-size:10px;color:var(--muted)}
.fc-pkgs{font-family:'SF Mono','Consolas',monospace;font-size:12px;font-weight:700;color:var(--acc)}
.fc-pos{font-size:10px;color:var(--muted);display:flex;align-items:center;gap:5px}
.pos-num{background:var(--s2);border:1px solid var(--b1);border-radius:3px;padding:1px 6px;font-weight:700;font-family:'SF Mono','Consolas',monospace;color:var(--acc);font-size:10px}
.fc-handle{font-size:10px;color:var(--b3);text-align:center;border-top:1px solid var(--b1);padding-top:6px;margin-top:2px}
.fila-empty{display:flex;align-items:center;justify-content:center;width:100%;font-size:11px;color:var(--muted);font-style:italic;padding:10px 0}
/* BOTTOM GRID */
.bottom-grid{display:grid;grid-template-columns:1fr 1fr;gap:12px}
.cw{padding:14px 16px}.cw canvas{max-height:220px}
.log-wrap{padding:12px 16px;max-height:260px;overflow-y:auto}
.log-wrap::-webkit-scrollbar{width:4px}.log-wrap::-webkit-scrollbar-thumb{background:var(--b2);border-radius:2px}
.lentry{display:flex;gap:10px;padding:5px 0;border-bottom:1px solid rgba(48,54,61,.4);font-size:11px}
.lentry:last-child{border-bottom:none}
.lt{font-family:'SF Mono','Consolas',monospace;font-size:10px;color:var(--muted);white-space:nowrap;min-width:58px;padding-top:2px}
.ld{width:6px;height:6px;border-radius:50%;flex-shrink:0;margin-top:4px}
.lm{color:var(--tx);line-height:1.5}
/* PRODUTIVIDADE */
.prod-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(210px,1fr));gap:8px;padding:12px 14px}
.pcard{background:var(--bg);border:1px solid var(--b1);border-radius:6px;padding:10px 12px}
.pc-head{display:flex;align-items:center;gap:8px;margin-bottom:8px}
.pc-av{width:26px;height:26px;border-radius:4px;display:flex;align-items:center;justify-content:center;font-size:9px;font-weight:700;color:#000;flex-shrink:0;font-family:'SF Mono','Consolas',monospace}
.pc-name{font-size:11px;font-weight:700;color:var(--tx);flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.pc-doca{font-family:'SF Mono','Consolas',monospace;font-size:10px;color:var(--muted);flex-shrink:0}
.pc-metrics{display:grid;grid-template-columns:1fr 1fr;gap:4px;margin-bottom:7px}
.pm{background:var(--s2);border-radius:4px;padding:5px 7px;text-align:center}
.pm-v{font-size:14px;font-weight:800;font-family:'SF Mono','Consolas',monospace;line-height:1}
.pm-l{font-size:8px;color:var(--muted);margin-top:2px}
.pc-sub2{font-size:9px;color:var(--blue);margin-top:2px}
.prod-empty{padding:20px 14px;text-align:center;font-size:11px;color:var(--muted);grid-column:1/-1}
/* DASHBOARD */
.dash-g2{display:grid;grid-template-columns:1fr 1fr;gap:12px}
.dash-g3{display:grid;grid-template-columns:1fr 1fr 1fr;gap:12px}
.twrap{overflow-x:auto}
table{width:100%;border-collapse:collapse}
thead th{background:var(--bg);color:var(--muted);font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.06em;padding:7px 12px;text-align:left;border-bottom:1px solid var(--b1);white-space:nowrap}
tbody td{padding:7px 12px;border-bottom:1px solid rgba(48,54,61,.5);font-size:11px;vertical-align:middle}
tbody tr:last-child td{border-bottom:none}
tbody tr:hover td{background:rgba(255,255,255,.02)}
.mono{font-family:'SF Mono','Consolas',monospace;font-size:10px}
.badge{display:inline-block;padding:2px 7px;border-radius:3px;font-size:10px;font-weight:700}
.badge.g{background:rgba(63,185,80,.15);color:var(--green);border:1px solid rgba(63,185,80,.2)}
.badge.a{background:rgba(230,168,23,.15);color:var(--acc);border:1px solid rgba(230,168,23,.2)}
.badge.r{background:rgba(248,81,73,.1);color:var(--red);border:1px solid rgba(248,81,73,.2)}
.badge.b{background:rgba(88,166,255,.15);color:var(--blue);border:1px solid rgba(88,166,255,.2)}
.ch-legend{display:flex;flex-wrap:wrap;gap:12px;padding:8px 14px;border-top:1px solid var(--b1)}
.ch-li{display:flex;align-items:center;gap:5px;font-size:10px;color:var(--muted)}
.ch-sw{width:12px;height:3px;border-radius:2px}
/* HISTORICO */
.hist-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(280px,1fr));gap:10px;padding:12px 14px}
.hist-card{background:var(--bg);border:1px solid var(--b1);border-radius:6px;padding:12px 14px}
.hist-date{font-family:'SF Mono','Consolas',monospace;font-size:11px;font-weight:700;color:var(--acc);margin-bottom:8px}
.hist-row{display:flex;justify-content:space-between;font-size:11px;padding:3px 0;border-bottom:1px solid rgba(48,54,61,.4)}
.hist-row:last-child{border-bottom:none}
.hist-k{color:var(--muted)}.hist-v{font-weight:600;color:var(--tx)}
/* CARRETAS */
.cart-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(260px,1fr));gap:10px;padding:12px 14px}
.cart-card{background:var(--bg);border:1px solid var(--b1);border-radius:6px;padding:12px 14px;display:flex;flex-direction:column;gap:6px}
.cart-card.active{border-color:var(--red)}
.cart-title{display:flex;align-items:center;justify-content:space-between}
.cart-id{font-family:'SF Mono','Consolas',monospace;font-size:12px;font-weight:700;color:var(--tx)}
.cart-status{font-size:9px;font-weight:700;padding:2px 7px;border-radius:3px}
.cs-active{background:rgba(248,81,73,.15);color:var(--red)}
.cs-done{background:rgba(63,185,80,.15);color:var(--green)}
.cart-timer{font-family:'SF Mono','Consolas',monospace;font-size:20px;font-weight:800;color:var(--acc)}
.cart-row{display:flex;justify-content:space-between;font-size:11px;color:var(--muted);align-items:center}
.cart-row span:last-child{color:var(--tx);font-weight:600;font-family:'SF Mono','Consolas',monospace}
.cart-inp{background:var(--s2);border:1px solid var(--b1);border-radius:4px;padding:4px 7px;color:var(--tx);font-size:11px;outline:none;width:70px;text-align:center}
.cart-inp:focus{border-color:var(--acc)}
/* MODAL */
.modal-bg{position:fixed;inset:0;background:rgba(0,0,0,.75);z-index:500;display:none;align-items:center;justify-content:center}
.modal-bg.open{display:flex}
.modal{background:var(--s1);border:1px solid var(--b1);border-radius:8px;padding:22px;width:340px;max-width:95vw}
.modal h3{font-size:13px;font-weight:700;color:#fff;margin-bottom:7px}
.modal p{font-size:11px;color:var(--muted);line-height:1.6;margin-bottom:14px}
.mbtns{display:flex;gap:8px;justify-content:flex-end}
/* MOTORISTA PORTAL */
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
.sd{background:rgba(63,185,80,.1);color:var(--green);border:1px solid rgba(63,185,80,.25)}
.mtimer{font-family:'SF Mono','Consolas',monospace;font-size:28px;font-weight:700;color:var(--acc);text-align:center;margin:8px 0}
.mabtn{width:100%;padding:12px;border:none;border-radius:6px;font-size:13px;font-weight:700;cursor:pointer;display:flex;align-items:center;justify-content:center;gap:8px;margin-top:10px;transition:all .15s}
.mabtn.g{background:var(--green);color:#000}.mabtn.g:hover{background:#4dd860}
.mabtn.r{background:var(--red);color:#fff}.mabtn.r:hover{background:#ff6459}
.mabtn.x{background:var(--s2);color:var(--tx);border:1px solid var(--b1)}
.emp-add{display:flex;gap:6px;margin-top:6px}
.emp-add input{flex:1;background:var(--bg);border:1px solid var(--b1);border-radius:5px;padding:6px 9px;color:var(--tx);font-size:12px;outline:none}
.emp-add input:focus{border-color:var(--acc)}
@media(max-width:900px){.dash-g2,.dash-g3,.bottom-grid,.aloc-pools{grid-template-columns:1fr}}
</style>
</head>
<body>

<!-- LOGIN -->
<div id="sc-login" class="screen active">
  <div class="lcard">
    <div class="lbrand">
      <div class="ico"><svg viewBox="0 0 24 24" fill="none" stroke="#000" stroke-width="2.5"><path d="M3 9l9-7 9 7v11a2 2 0 01-2 2H5a2 2 0 01-2-2z"/><polyline points="9,22 9,12 15,12 15,22"/></svg></div>
      <div><h1>DocaOps</h1><p>Sistema de Gestão Operacional</p></div>
    </div>
    <div class="lf"><label>Usuário</label><input type="text" id="lu" placeholder="Seu usuário" autocomplete="off"></div>
    <div class="lf"><label>Senha</label><input type="password" id="lp" placeholder="••••••" onkeydown="if(event.key==='Enter')doLogin()"></div>
    <div id="lerr" class="lerr">Usuário ou senha incorretos.</div>
    <button class="lbtn" onclick="doLogin()">Acessar o Sistema</button>
    <div class="lmoto">Motorista? <a href="#" onclick="showMoto()">Acessar portal de chegada</a></div>
  </div>
</div>

<!-- APP -->
<div id="sc-app" class="screen">
  <div class="topbar">
    <div class="tbrand">
      <div class="ico"><svg viewBox="0 0 24 24" fill="none" stroke="#000" stroke-width="2.5"><path d="M3 9l9-7 9 7v11a2 2 0 01-2 2H5a2 2 0 01-2-2z"/><polyline points="9,22 9,12 15,12 15,22"/></svg></div>
      <span>DocaOps</span>
    </div>
    <div class="navtabs">
      <button class="ntab active" onclick="setTab('gestao',this)">
        <svg viewBox="0 0 16 16" fill="currentColor" width="12" height="12"><rect x="1" y="1" width="6" height="6" rx="1"/><rect x="9" y="1" width="6" height="6" rx="1"/><rect x="1" y="9" width="6" height="6" rx="1"/><rect x="9" y="9" width="6" height="6" rx="1"/></svg>
        Doca em Tempo Real
      </button>
      <button class="ntab" onclick="setTab('carretas',this)">
        <svg viewBox="0 0 16 16" fill="currentColor" width="12" height="12"><path d="M1 4h10v7H1zm10 2l4 2v5h-4z"/><circle cx="4" cy="12" r="1.5"/><circle cx="12" cy="12" r="1.5"/></svg>
        Carretas
      </button>
      <button class="ntab" onclick="setTab('historico',this)">
        <svg viewBox="0 0 16 16" fill="currentColor" width="12" height="12"><rect x="1" y="2" width="14" height="13" rx="2" fill="none" stroke="currentColor" stroke-width="1.5"/><path d="M5 1v3M11 1v3M1 6h14"/><rect x="4" y="9" width="2" height="2" rx=".5"/><rect x="7" y="9" width="2" height="2" rx=".5"/><rect x="10" y="9" width="2" height="2" rx=".5"/></svg>
        Histórico
      </button>
      <button class="ntab" onclick="setTab('dash',this)">
        <svg viewBox="0 0 16 16" fill="currentColor" width="12" height="12"><path d="M1 14V8h3v6H1zm4-8v8h3V6H5zm4 3v5h3V9H9zm4-5v10h2V4h-2z"/></svg>
        Dashboard & Analytics
      </button>
    </div>
    <div class="tright">
      <div id="tpill" class="tpill off">● INATIVO</div>
      <div class="tclock" id="clock">--:--:--</div>
      <button class="tbtn" onclick="doLogout()">Sair</button>
    </div>
  </div>

  <!-- TAB GESTÃO -->
  <div id="tab-gestao" class="mcontent">
    <div class="kstrip">
      <div class="kpi">
        <div class="kl">Meta de Pacotes</div>
        <div style="display:flex;align-items:center;gap:5px;margin-top:2px">
          <input type="number" id="metaInput" value="5000" style="width:82px;background:var(--bg);border:1px solid var(--b1);border-radius:4px;padding:2px 6px;color:#fff;font-size:20px;font-weight:800;font-family:'SF Mono','Consolas',monospace;outline:none">
          <button class="btn a" onclick="G.setMeta()" style="padding:3px 8px;font-size:10px">✓</button>
        </div>
      </div>
      <div class="kpi"><div class="kl">Descarregados Hoje</div><div class="kv g" id="kDesc">0</div><div class="ks">pacotes</div></div>
      <div class="kpi"><div class="kl">Docas Ativas</div><div class="kv b" id="kAtivas">0</div><div class="ks">de 13</div></div>
      <div class="kpi"><div class="kl">Duplas na Doca</div><div class="kv a" id="kDuplas">0</div></div>
      <div class="kpi"><div class="kl">Motoristas na Doca</div><div class="kv p" id="kMotos">0</div></div>
      <div class="kpi"><div class="kl">Na Fila</div><div class="kv" id="kFila">0</div></div>
      <div class="kpi" style="grid-column:span 2">
        <div class="kl">Progresso da Meta</div>
        <div style="display:flex;align-items:baseline;gap:8px;margin:2px 0 4px"><div class="kv" id="kPct">0%</div><div class="ks" id="kSub">0 / 5.000</div></div>
        <div class="pbar"><div class="pfill" id="kBar" style="width:0%"></div></div>
      </div>
      <div class="kpi">
        <div class="kl">Ações do Dia</div>
        <div style="display:flex;flex-direction:column;gap:4px;margin-top:4px">
          <button class="btn a" onclick="G.fecharDia()" style="font-size:10px;padding:4px 8px">💾 Fechar Dia</button>
          <button class="btn r" onclick="G.resetarDia()" style="font-size:10px;padding:4px 8px">↺ Resetar</button>
        </div>
      </div>
    </div>

    <!-- Duplas -->
    <div class="panel">
      <div class="phead">
        <div class="ptitle"><span class="dot da"></span>Duplas de Trabalho</div>
        <div style="display:flex;gap:6px">
          <button class="btn a" onclick="G.toggleForm()">+ Nova Dupla</button>
        </div>
      </div>
      <div id="formDupla" style="display:none;border-top:1px solid var(--b1)">
        <div class="frow">
          <div class="fld" style="flex:1;min-width:160px"><label>Nomes da Dupla</label><input type="text" id="dNomes" placeholder="ex: João e Maria"></div>
          <div class="fld" style="width:100px"><label>Pacotes</label><input type="number" id="dPkgs" placeholder="0" min="0"></div>
          <div class="fld" style="align-self:flex-end"><button class="btn g" onclick="G.addDupla()">Adicionar</button></div>
        </div>
      </div>
    </div>

    <!-- Alocação -->
    <div class="panel">
      <div class="phead">
        <div class="ptitle"><span class="dot da"></span>Alocar nas Docas — arraste para a doca desejada</div>
        <div style="display:flex;gap:10px;font-size:10px;color:var(--muted)">
          <span><span class="dot dg" style="display:inline-block"></span> Dupla</span>
          <span><span class="dot db" style="display:inline-block"></span> Motorista</span>
          <span><span class="dot dr" style="display:inline-block"></span> Carreta</span>
        </div>
      </div>
      <div class="aloc-pools">
        <div class="drag-list">
          <div class="dlhead">Duplas disponíveis <span class="cnt" id="cntDuplas">0</span></div>
          <div class="dl-items" id="poolDuplas" style="max-height:200px"></div>
        </div>
        <div class="drag-list">
          <div class="dlhead">Motoristas disponíveis <span class="cnt" id="cntMotos">0</span></div>
          <div class="dl-items" id="poolMotos" style="max-height:200px"></div>
        </div>
        <div class="drag-list">
          <div class="dlhead">Carretas <span class="cnt" id="cntCarretas">0</span></div>
          <div class="dl-items" id="poolCarretas" style="max-height:200px"></div>
          <div style="padding:8px 12px;border-top:1px solid var(--b1)">
            <button class="btn r" onclick="G.criarCarreta()" style="width:100%;justify-content:center;font-size:10px">+ Nova Carreta</button>
          </div>
        </div>
      </div>
      <div class="aloc-hint">
        <span>↑ Arraste qualquer item acima para uma doca abaixo</span>
        <span style="color:var(--acc)">✕ dentro da doca devolve ao pool</span>
      </div>
      <div class="docas-grid" id="docasGrid"></div>
    </div>

    <!-- Produtividade -->
    <div class="panel">
      <div class="phead"><div class="ptitle"><span class="dot dg"></span>Produtividade do Turno — dados reais de operações finalizadas</div></div>
      <div class="prod-grid" id="prodGrid"></div>
    </div>

    <!-- Fila -->
    <div class="panel">
      <div class="phead">
        <div class="ptitle"><span class="dot db"></span>Fila de Espera</div>
        <div style="display:flex;align-items:center;gap:12px">
          <span style="font-size:10px;color:var(--muted)">Arraste um card diretamente para uma doca acima</span>
          <span id="filaTotal" style="font-family:'SF Mono','Consolas',monospace;font-size:11px;color:var(--acc);font-weight:700">0</span>
        </div>
      </div>
      <div class="fila-scroll" id="filaScroll"></div>
    </div>

    <!-- Gráfico + Log -->
    <div class="bottom-grid">
      <div class="panel">
        <div class="phead"><div class="ptitle"><span class="dot dg"></span>Entradas por Hora</div></div>
        <div class="cw"><canvas id="hrChart"></canvas></div>
      </div>
      <div class="panel">
        <div class="phead"><div class="ptitle"><span class="dot dm"></span>Log de Eventos</div></div>
        <div class="log-wrap" id="logWrap"></div>
      </div>
    </div>
  </div>

  <!-- TAB CARRETAS -->
  <div id="tab-carretas" class="mcontent" style="display:none">
    <div class="kstrip">
      <div class="kpi"><div class="kl">Carretas Hoje</div><div class="kv a" id="cTot">0</div></div>
      <div class="kpi"><div class="kl">Em Doca Agora</div><div class="kv r" id="cAtivas">0</div></div>
      <div class="kpi"><div class="kl">Finalizadas</div><div class="kv g" id="cFin">0</div></div>
      <div class="kpi"><div class="kl">Tempo Médio</div><div class="kv" id="cAvgT" style="font-size:16px;margin-top:3px">—</div></div>
      <div class="kpi"><div class="kl">Total Paletes</div><div class="kv b" id="cTotPal">0</div></div>
      <div class="kpi"><div class="kl">Total Manga-paletes</div><div class="kv p" id="cTotMP">0</div></div>
    </div>
    <div style="display:flex;justify-content:flex-end;gap:8px;padding:0 2px">
      <button class="btn" onclick="CAR.exportar()">↓ Exportar Excel</button>
    </div>
    <div class="cart-grid" id="cartGrid">
      <div style="grid-column:1/-1;padding:20px;text-align:center;font-size:11px;color:var(--muted)">Nenhuma carreta registrada. Crie uma na aba "Doca em Tempo Real".</div>
    </div>
  </div>

  <!-- TAB HISTÓRICO -->
  <div id="tab-historico" class="mcontent" style="display:none">
    <div class="panel">
      <div class="phead">
        <div class="ptitle"><span class="dot da"></span>Histórico por Data</div>
        <div style="display:flex;gap:8px;align-items:center">
          <input type="month" id="histMes" style="background:var(--bg);border:1px solid var(--b1);border-radius:5px;padding:5px 9px;color:var(--tx);font-size:12px;outline:none" onchange="HIST.render()">
          <button class="btn r" onclick="HIST.limpar()">⊘ Apagar Tudo</button>
        </div>
      </div>
    </div>
    <div class="hist-grid" id="histGrid">
      <div style="grid-column:1/-1;padding:20px;text-align:center;font-size:11px;color:var(--muted)">Nenhum dia fechado ainda. Use "Fechar Dia" na aba Doca em Tempo Real.</div>
    </div>
  </div>

  <!-- TAB DASHBOARD -->
  <div id="tab-dash" class="mcontent" style="display:none">
    <div class="panel">
      <div class="phead"><div class="ptitle"><span class="dot da"></span>Registrar Operação Manual</div></div>
      <div class="frow">
        <div class="fld" style="flex:1;min-width:140px"><label>Nome/Dupla</label><input type="text" id="opTeam" list="teamsDL" placeholder="Nome"><datalist id="teamsDL"></datalist></div>
        <div class="fld"><label>Início</label><input type="datetime-local" id="opStart"></div>
        <div class="fld"><label>Fim</label><input type="datetime-local" id="opEnd"></div>
        <div class="fld" style="width:90px"><label>Pacotes</label><input type="number" id="opPkgs"></div>
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
      <div class="kpi"><div class="kl">Nomes Únicos</div><div class="kv" id="dTeams">0</div></div>
      <div class="kpi"><div class="kl">Mais Rápido(a)</div><div class="kv g" id="dFast" style="font-size:13px;margin-top:3px">—</div></div>
      <div class="kpi"><div class="kl">Mais Lento(a)</div><div class="kv r" id="dSlow" style="font-size:13px;margin-top:3px">—</div></div>
      <div class="kpi"><div class="kl">Média pac/h</div><div class="kv" id="dAvg">0</div><div class="ks">baseado em ops concluídas</div></div>
      <div class="kpi"><div class="kl">Tempo Médio/Op</div><div class="kv" id="dAvgT" style="font-size:16px;margin-top:3px">—</div></div>
    </div>
    <div class="dash-g3">
      <div class="panel"><div class="phead"><div class="ptitle"><span class="dot da"></span>Volume por Nome</div></div><div class="cw"><canvas id="cVol"></canvas></div></div>
      <div class="panel"><div class="phead"><div class="ptitle"><span class="dot db"></span>Eficiência pac/h <span style="font-size:9px;color:var(--muted);font-weight:400;margin-left:4px">— linha vermelha = média</span></div></div><div class="cw"><canvas id="cEff"></canvas></div></div>
      <div class="panel"><div class="phead"><div class="ptitle"><span class="dot dg"></span>Caminhões</div></div><div class="cw"><canvas id="cTrk"></canvas></div></div>
    </div>
    <div class="dash-g2">
      <div class="panel"><div class="phead"><div class="ptitle"><span class="dot dm"></span>Distribuição de Volume</div></div><div class="cw"><canvas id="cDist"></canvas></div></div>
      <div class="panel">
        <div class="phead"><div class="ptitle"><span class="dot da"></span>Acumulado por Hora</div></div>
        <div class="cw"><canvas id="cLine"></canvas></div>
        <div class="ch-legend" id="lineLegend"></div>
      </div>
    </div>
    <div class="panel">
      <div class="phead"><div class="ptitle"><span class="dot dm"></span>Detalhamento de Operações Concluídas</div></div>
      <div class="twrap">
        <table id="opsTable">
          <thead><tr><th>Nome / Dupla</th><th>Tipo</th><th>Início</th><th>Término</th><th>Pacotes</th><th>Duração (h)</th><th>Eficiência</th><th>Rating</th></tr></thead>
          <tbody></tbody>
        </table>
      </div>
    </div>
  </div>
</div>

<!-- PORTAL MOTORISTA -->
<div id="sc-moto" class="screen">
  <div class="mwrap">
    <div class="mcard">
      <div class="mhead">
        <div class="ico"><svg viewBox="0 0 24 24" fill="none" stroke="#000" stroke-width="2.5"><rect x="1" y="3" width="15" height="13" rx="2"/><path d="M16 8l5 3v5h-5V8z"/><circle cx="5.5" cy="18.5" r="2.5"/><circle cx="18.5" cy="18.5" r="2.5"/></svg></div>
        <div><h2>Portal do Motorista</h2><p>Registro de chegada e descarga</p></div>
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
                <button class="btn a" onclick="M.addEmpresa()">+ Adicionar</button>
              </div>
            </div>
            <div class="fld"><label>Volume de Pacotes</label><input type="number" id="mfPkgs" placeholder="Nº de pacotes" style="font-size:13px;padding:8px 10px"></div>
            <button class="mabtn g" onclick="M.registrar()">✓ ENTRAR NA FILA</button>
            <button class="mabtn x" onclick="showLogin()" style="font-size:11px;padding:8px;margin-top:2px">← Voltar ao login</button>
          </div>
        </div>
        <div id="mStatus" style="display:none">
          <div id="msWait" class="mstatus">
            <div style="font-size:11px;color:var(--muted);margin-bottom:4px">Sua posição na fila</div>
            <div class="bignum" id="mPos">—</div>
            <div class="stag sw">● Aguardando</div>
            <div style="font-size:10px;color:var(--muted);margin-top:12px">Você será chamado em breve pelo gestor.</div>
          </div>
          <div id="msCall" style="display:none;text-align:center;padding:16px 0">
            <div style="font-size:11px;color:var(--blue);font-weight:700;margin-bottom:8px">▲ SUA VEZ — DIRIJA-SE À:</div>
            <div class="bigdoca" id="mDocaNum">—</div>
            <button class="mabtn g" onclick="M.iniciar()" style="margin-top:16px">▶ Iniciar Descarga</button>
          </div>
          <div id="msDesc" style="display:none">
            <div style="text-align:center;padding:14px 0 8px">
              <div style="font-size:10px;color:var(--muted)">Descarga em andamento</div>
              <div class="mtimer" id="mTimer">00:00:00</div>
              <div class="stag sd">● Descarregando</div>
            </div>
            <button class="mabtn r" onclick="M.finalizar()">■ Finalizar Descarga</button>
          </div>
          <div id="msFim" style="display:none;text-align:center;padding:20px 0">
            <div style="font-size:40px;color:var(--green)">✓</div>
            <div style="font-size:16px;font-weight:700;color:var(--green);margin-top:8px">Descarga Concluída!</div>
            <div style="font-size:11px;color:var(--muted);margin-top:6px">Obrigado pelo serviço.</div>
            <button class="mabtn x" onclick="M.novo()" style="margin-top:16px">← Novo Registro</button>
          </div>
          <div id="mNomeDisp" style="margin-top:12px;padding-top:12px;border-top:1px solid var(--b1);font-size:10px;color:var(--muted);text-align:center"></div>
          <button class="mabtn x" onclick="M.novo()" style="font-size:10px;padding:8px;margin-top:8px">↩ Sair</button>
        </div>
      </div>
    </div>
    <div style="text-align:center;margin-top:14px;font-size:10px;color:var(--muted)">Gestor? <a href="#" onclick="showLogin()" style="color:var(--acc)">Painel de controle</a></div>
  </div>
</div>

<!-- MODAL RESET -->
<div class="modal-bg" id="modalReset">
  <div class="modal">
    <h3>Resetar o dia</h3>
    <p>Isso vai limpar as docas, a fila e os contadores de hoje. Os dados serão perdidos (use "Fechar Dia" primeiro para salvar). Continuar?</p>
    <div class="mbtns">
      <button class="btn" onclick="closeModal('modalReset')">Cancelar</button>
      <button class="btn r" onclick="G.confirmarReset()">Resetar</button>
    </div>
  </div>
</div>
<!-- MODAL CLEAR OPS -->
<div class="modal-bg" id="modalOk">
  <div class="modal">
    <h3>Confirmar exclusão</h3>
    <p>Todos os dados de operações do Dashboard serão apagados permanentemente.</p>
    <div class="mbtns">
      <button class="btn" onclick="closeModal('modalOk')">Cancelar</button>
      <button class="btn r" onclick="D.confirmar()">Limpar tudo</button>
    </div>
  </div>
</div>

<script>
// ═══════════════════════════════════════════════════
//  ESTADO
// ═══════════════════════════════════════════════════
const S = {
  docas: Array.from({length:13}, (_,i) => ({id:169+i, status:'livre', dupla:null, motorista:null, carreta:null, inicio:null})),
  pool_duplas: [],
  pool_motos: [],
  pool_carretas: [],  // carretas prontas para arrastar
  carretas_all: [],   // todas as carretas (ativas + finalizadas)
  ops: [],
  filteredOps: [],
  meta: 5000, desc: 0,
  historico: {}, // {hora: count}
  historico_dias: {}, // {YYYY-MM-DD: {desc, caminhoes, duplas, carretas}}
  log: [],
  teamColors: {}, cidx: 0,
  motorista: null,
  empresas: ['Prálog','Imediato','Hawk','Ontime'],
};

const PAL = ['#e6a817','#58a6ff','#3fb950','#f85149','#bc8cff','#ffa657','#39d353','#ff7b72','#d2a8ff','#79c0ff','#56d364','#f0883e'];
const KEYS_ARR = ['docas','pool_duplas','pool_motos','pool_carretas','carretas_all','ops','historico','historico_dias','log','teamColors','empresas'];

function sv(){
  try{
    KEYS_ARR.forEach(k => localStorage.setItem('dop_'+k, JSON.stringify(S[k])));
    ['meta','desc','cidx'].forEach(k => localStorage.setItem('dop_'+k, S[k]));
    localStorage.setItem('dop_motorista', JSON.stringify(S.motorista));
  }catch(e){}
}
function ld(){
  try{
    KEYS_ARR.forEach(k => { const v=localStorage.getItem('dop_'+k); if(v) S[k]=JSON.parse(v); });
    ['meta','desc','cidx'].forEach(k => { const v=localStorage.getItem('dop_'+k); if(v!==null) S[k]=+v; });
    const mo=localStorage.getItem('dop_motorista'); if(mo) S.motorista=JSON.parse(mo);
  }catch(e){}
  if(!S.docas.length || S.docas[0]?.id < 169)
    S.docas = Array.from({length:13}, (_,i) => ({id:169+i, status:'livre', dupla:null, motorista:null, carreta:null, inicio:null}));
  if(!S.ops.length) S.ops = genOps();
}

function genOps(){
  const names=["Carlos & Bruna","Lucas & Fernanda","Matheus & Gabriela","Rafael & Juliana","Vinicius & Patricia"];
  const td = new Date().toISOString().slice(0,10);
  return names.map((nm,i) => {
    const h=13+(i%5), dur=60+Math.floor(Math.random()*90);
    const st=new Date(`${td}T${String(h).padStart(2,'0')}:${String((i*7)%60).padStart(2,'0')}`);
    const en=new Date(st.getTime()+dur*60000);
    return {id:i+1, team:nm, tipo:'dupla', startTime:st.toISOString(), endTime:en.toISOString(), packages:1500+Math.floor(Math.random()*800)};
  });
}

// ═══════════════════════════════════════════════════
//  HELPERS
// ═══════════════════════════════════════════════════
function addLog(tipo,msg){ const h=new Date().toLocaleTimeString('pt-BR',{hour:'2-digit',minute:'2-digit',second:'2-digit'}); S.log.unshift({tipo,msg,h}); if(S.log.length>100)S.log.length=100; }
function clr(t){ if(!S.teamColors[t]){ S.teamColors[t]=PAL[S.cidx%PAL.length]; S.cidx++; } return S.teamColors[t]; }
function ini(n){ if(!n)return'??'; const p=n.replace(/&/g,'e').split(/\s+e\s+/i).map(x=>x.trim()); if(p.length>=2&&p[0]&&p[1])return(p[0][0]+p[1][0]).toUpperCase(); const w=n.split(' '); if(w.length>=2)return(w[0][0]+w[w.length-1][0]).toUpperCase(); return n.slice(0,2).toUpperCase(); }
function fdt(d){ return new Date(d).toLocaleString('pt-BR'); }
function fmtMs(ms){ const h=Math.floor(ms/3600000),m=Math.floor((ms%3600000)/60000),s=Math.floor((ms%60000)/1000); return`${String(h).padStart(2,'0')}:${String(m).padStart(2,'0')}:${String(s).padStart(2,'0')}`; }
function today(){ return new Date().toISOString().slice(0,10); }
function closeModal(id){ document.getElementById(id).classList.remove('open'); }
function fillEmpSels(){ ['mfEmp'].forEach(id=>{ const el=document.getElementById(id); if(!el)return; const cv=el.value; el.innerHTML=S.empresas.map(e=>`<option>${e}</option>`).join(''); if(cv&&S.empresas.includes(cv))el.value=cv; }); }

// ═══════════════════════════════════════════════════
//  NAV
// ═══════════════════════════════════════════════════
function showScreen(id){ document.querySelectorAll('.screen').forEach(s=>s.classList.remove('active')); document.getElementById(id).classList.add('active'); }
function showLogin(){ showScreen('sc-login'); }
function showMoto(){ showScreen('sc-moto'); M.render(); }
function doLogin(){
  const u=document.getElementById('lu').value.trim(), p=document.getElementById('lp').value;
  const e=document.getElementById('lerr');
  if(u==='gestor'&&p==='1234'){ e.style.display='none'; showScreen('sc-app'); G.render(); D.filter('today'); CAR.render(); HIST.init(); }
  else { e.style.display='block'; }
}
function doLogout(){ document.getElementById('lu').value=''; document.getElementById('lp').value=''; showScreen('sc-login'); }
function setTab(id, btn){
  document.querySelectorAll('.ntab').forEach(b=>b.classList.remove('active')); btn.classList.add('active');
  ['gestao','carretas','historico','dash'].forEach(t=>{ const el=document.getElementById('tab-'+t); if(el)el.style.display=t===id?'flex':'none'; });
  if(id==='dash') D.filter(document.getElementById('fPeriod').value);
  if(id==='carretas') CAR.render();
  if(id==='historico') HIST.render();
}

// ═══════════════════════════════════════════════════
//  CHARTS
// ═══════════════════════════════════════════════════
let CH = {};
const TT = {backgroundColor:'#1c2331',borderColor:'#30363d',borderWidth:1,titleColor:'#c9d1d9',bodyColor:'#8b949e',padding:10,cornerRadius:6};
const SC = {x:{ticks:{color:'#8b949e',font:{size:10}},grid:{color:'rgba(255,255,255,0.04)'}},y:{ticks:{color:'#8b949e',font:{size:10}},grid:{color:'rgba(255,255,255,0.04)'},beginAtZero:true}};

function initCharts(){
  CH.hr = new Chart(document.getElementById('hrChart'),{
    type:'bar',
    data:{labels:Array.from({length:24},(_,i)=>`${i}h`),datasets:[{label:'Entradas',data:Array(24).fill(0),backgroundColor:ctx=>{const v=ctx.raw||0,mx=15;return`rgba(230,168,23,${Math.max(0.2,Math.min(1,v/mx))})`;},borderColor:'#e6a817',borderWidth:0,borderRadius:3}]},
    options:{responsive:true,maintainAspectRatio:true,plugins:{legend:{display:false},tooltip:{...TT,callbacks:{label:ctx=>`${ctx.raw} caminhão(s)`}}},scales:SC}
  });

  // Volume — barra horizontal com gradiente de cor por valor
  CH.vol = new Chart(document.getElementById('cVol'),{
    type:'bar',
    data:{labels:[],datasets:[{data:[],backgroundColor:[],borderRadius:3,borderWidth:0}]},
    options:{responsive:true,maintainAspectRatio:true,indexAxis:'y',plugins:{legend:{display:false},tooltip:{...TT,callbacks:{label:ctx=>`${ctx.raw.toLocaleString('pt-BR')} pacotes`}}},scales:{...SC,x:{...SC.x,ticks:{...SC.x.ticks,callback:v=>v.toLocaleString('pt-BR')}}}}
  });

  // Eficiência — barra horizontal + linha de média
  CH.eff = new Chart(document.getElementById('cEff'),{
    type:'bar',
    data:{labels:[],datasets:[
      {label:'Eficiência',data:[],backgroundColor:[],borderRadius:3,borderWidth:0,order:2},
      {label:'Média',data:[],type:'line',borderColor:'rgba(248,81,73,0.7)',borderWidth:1.5,borderDash:[4,3],pointRadius:0,fill:false,order:1}
    ]},
    options:{responsive:true,maintainAspectRatio:true,indexAxis:'y',
      plugins:{legend:{display:false},tooltip:{...TT,callbacks:{label:ctx=>ctx.datasetIndex===0?`${ctx.raw} pac/h`:`Média: ${ctx.raw} pac/h`}}},
      scales:{...SC,x:{...SC.x,ticks:{...SC.x.ticks,callback:v=>`${v}p/h`}}}
    }
  });

  CH.trk = new Chart(document.getElementById('cTrk'),{
    type:'bar',
    data:{labels:[],datasets:[{data:[],backgroundColor:[],borderRadius:3,borderWidth:0}]},
    options:{responsive:true,maintainAspectRatio:true,plugins:{legend:{display:false},tooltip:{...TT,callbacks:{label:ctx=>`${ctx.raw} caminhão(s)`}}},scales:SC}
  });

  CH.dist = new Chart(document.getElementById('cDist'),{
    type:'doughnut',
    data:{labels:[],datasets:[{data:[],backgroundColor:[],borderWidth:2,borderColor:'#161b22'}]},
    options:{responsive:true,cutout:'65%',plugins:{legend:{position:'right',labels:{color:'#8b949e',font:{size:10},boxWidth:10,padding:8}},tooltip:{...TT,callbacks:{label:ctx=>`${ctx.label}: ${ctx.raw.toLocaleString('pt-BR')} pac`}}}}
  });

  CH.line = new Chart(document.getElementById('cLine'),{
    type:'line',
    data:{labels:[],datasets:[]},
    options:{responsive:true,maintainAspectRatio:true,plugins:{legend:{display:false},tooltip:TT},scales:SC,elements:{line:{tension:.4,borderWidth:2},point:{radius:3,hoverRadius:5}}}
  });
}

// ═══════════════════════════════════════════════════
//  G — GESTÃO
// ═══════════════════════════════════════════════════
const G = {
  toggleForm(){ const f=document.getElementById('formDupla'); f.style.display=f.style.display==='none'?'block':'none'; },

  addDupla(){
    const nomes=document.getElementById('dNomes').value.trim();
    const pkgs=parseInt(document.getElementById('dPkgs').value)||0;
    if(!nomes){ alert('Informe os nomes.'); return; }
    const c=clr(nomes);
    S.pool_duplas.push({id:Date.now(), nomes, pkgs, color:c, tipo:'dupla'});
    document.getElementById('dNomes').value=''; document.getElementById('dPkgs').value='';
    document.getElementById('formDupla').style.display='none';
    addLog('reg', `Dupla "${nomes}" cadastrada — ${pkgs.toLocaleString()} pac.`);
    sv(); this.render();
  },

  criarCarreta(){
    const id = Date.now();
    const label = `Carreta ${S.carretas_all.length+1}`;
    const c = {id, label, status:'pool', docaId:null, inicio:null, fim:null, paletes:0, mangapaletes:0};
    S.carretas_all.push(c);
    S.pool_carretas.push({id, label, tipo:'carreta'});
    addLog('reg', `${label} criada.`);
    sv(); this.render();
  },

  setMeta(){ const v=parseInt(document.getElementById('metaInput').value); if(v>0){ S.meta=v; sv(); this.render(); } },

  fecharDia(){
    const d = today();
    const caminhoes = S.ops.filter(o=>new Date(o.startTime).toISOString().slice(0,10)===d).length;
    const carretas = S.carretas_all.filter(c=>c.status==='done'&&c.fim&&c.fim.startsWith(d)).length;
    const duplas = [...new Set(S.ops.filter(o=>o.tipo==='dupla'&&o.startTime.startsWith(d)).map(o=>o.team))].length;
    S.historico_dias[d] = { desc: S.desc, caminhoes, duplas, carretas, meta: S.meta, date: d };
    addLog('sys', `Dia ${d} fechado e salvo no histórico.`);
    sv(); HIST.render();
    alert(`Dia ${d} salvo no histórico com sucesso!`);
  },

  resetarDia(){ document.getElementById('modalReset').classList.add('open'); },

  confirmarReset(){
    S.docas = Array.from({length:13}, (_,i) => ({id:169+i, status:'livre', dupla:null, motorista:null, carreta:null, inicio:null}));
    S.pool_duplas = []; S.pool_motos = []; S.pool_carretas = [];
    S.carretas_all = S.carretas_all.filter(c=>c.status==='done'); // mantém finalizadas
    S.desc = 0; S.historico = {}; S.log = [];
    closeModal('modalReset');
    addLog('sys', 'Dia resetado pelo gestor.');
    sv(); this.render();
  },

  alocar(itemId, itemTipo, docaId){
    const doca = S.docas.find(d=>d.id===docaId);
    if(!doca) return;

    if(itemTipo==='dupla'){
      // Não pode soltar dupla em doca que já tem carreta como status único
      if(doca.status==='carreta'&&!doca.dupla&&!doca.motorista){ alert('Esta doca está com carreta. Finalize-a primeiro.'); return; }
      if(doca.dupla){ alert('Doca já tem uma dupla.'); return; }
      const idx=S.pool_duplas.findIndex(x=>x.id===itemId); if(idx<0) return;
      const dup=S.pool_duplas.splice(idx,1)[0];
      doca.dupla=dup; if(!doca.inicio) doca.inicio=new Date().toISOString();
      doca.status=doca.motorista?'both':doca.carreta?'both':'dup';
      addLog('call', `Dupla "${dup.nomes}" → Doca ${docaId}`);

    } else if(itemTipo==='motorista'){
      if(doca.status==='carreta'&&!doca.dupla&&!doca.motorista){ alert('Esta doca está com carreta. Finalize-a primeiro.'); return; }
      if(doca.motorista){ alert('Doca já tem um motorista.'); return; }
      let mot=null;
      const idx=S.pool_motos.findIndex(x=>x.id===itemId);
      if(idx>=0){ mot=S.pool_motos.splice(idx,1)[0]; }
      else if(S.motorista&&S.motorista.id===itemId){
        mot={id:S.motorista.id, nome:S.motorista.nome, empresa:S.motorista.empresa, pacotes:S.motorista.pacotes, color:S.motorista.color||clr(S.motorista.nome), tipo:'motorista'};
      }
      if(!mot) return;
      doca.motorista=mot; if(!doca.inicio) doca.inicio=new Date().toISOString();
      doca.status=doca.dupla||doca.carreta?'both':'mot';
      addLog('call', `${mot.nome} → Doca ${docaId}`);
      if(S.motorista&&S.motorista.id===mot.id){ S.motorista.status='chamado'; S.motorista.docaId=docaId; M.render(); }

    } else if(itemTipo==='carreta'){
      if(doca.carreta){ alert('Doca já tem uma carreta alocada.'); return; }
      if(doca.dupla||doca.motorista){ alert('Remova a dupla/motorista antes de alocar uma carreta exclusiva, ou finalize a doca.'); return; }
      const idx=S.pool_carretas.findIndex(x=>x.id===itemId); if(idx<0) return;
      const cart=S.pool_carretas.splice(idx,1)[0];
      doca.carreta=cart; if(!doca.inicio) doca.inicio=new Date().toISOString();
      doca.status='carreta';
      const ca=S.carretas_all.find(c=>c.id===cart.id);
      if(ca){ ca.status='ativa'; ca.docaId=docaId; ca.inicio=doca.inicio; }
      addLog('call', `${cart.label} → Doca ${docaId}`);
    }
    sv(); this.render();
  },

  removerDupla(docaId){
    const d=S.docas.find(x=>x.id===docaId); if(!d||!d.dupla) return;
    S.pool_duplas.push(d.dupla); d.dupla=null;
    d.status=d.motorista?'mot':d.carreta?'carreta':'livre';
    if(d.status==='livre') d.inicio=null;
    addLog('sys', `Dupla removida da Doca ${docaId}`);
    sv(); this.render();
  },

  removerMoto(docaId){
    const d=S.docas.find(x=>x.id===docaId); if(!d||!d.motorista) return;
    S.pool_motos.push(d.motorista); d.motorista=null;
    d.status=d.dupla?'dup':d.carreta?'carreta':'livre';
    if(d.status==='livre') d.inicio=null;
    if(S.motorista&&S.motorista.docaId===docaId){ S.motorista.status='fila'; S.motorista.docaId=null; M.render(); }
    addLog('sys', `Motorista removido da Doca ${docaId}`);
    sv(); this.render();
  },

  finalizarDoca(id){
    const d=S.docas.find(x=>x.id===id); if(!d) return;
    const ini=d.inicio||new Date().toISOString(), fim=new Date().toISOString();
    const hr=new Date(ini).getHours();
    S.historico[hr]=(S.historico[hr]||0)+1;

    if(d.dupla){
      const p=d.dupla.pkgs||0;
      S.ops.push({id:Date.now()+1, team:d.dupla.nomes, tipo:'dupla', startTime:ini, endTime:fim, packages:p});
      S.desc+=p; S.pool_duplas.push(d.dupla);
      addLog('done', `Dupla "${d.dupla.nomes}" — Doca ${id} — +${p.toLocaleString()} pac.`);
    }
    if(d.motorista){
      const p=d.motorista.pacotes||0;
      S.ops.push({id:Date.now()+2, team:d.motorista.nome, tipo:'motorista', startTime:ini, endTime:fim, packages:p});
      S.desc+=p;
      addLog('done', `${d.motorista.nome} — Doca ${id} — +${p.toLocaleString()} pac.`);
      if(S.motorista&&S.motorista.id===d.motorista.id){ S.motorista.status='fim'; M.render(); }
    }
    if(d.carreta){
      const ca=S.carretas_all.find(c=>c.id===d.carreta.id);
      if(ca){ ca.status='done'; ca.fim=fim; ca.docaId=id; }
      S.pool_carretas.push(d.carreta);
      addLog('done', `${d.carreta.label} finalizada — Doca ${id}.`);
    }
    if(d.status==='carreta'&&!d.carreta) addLog('done', `Doca ${id} liberada.`);

    d.status='livre'; d.dupla=null; d.motorista=null; d.carreta=null; d.inicio=null;
    sv(); this.render(); D.filter(document.getElementById('fPeriod').value);
  },

  render(){
    const ativos=S.docas.filter(d=>d.status!=='livre');
    const pct=Math.min(100,Math.round(S.desc/S.meta*100));
    document.getElementById('kDesc').textContent=S.desc.toLocaleString('pt-BR');
    document.getElementById('kAtivas').textContent=ativos.length;
    document.getElementById('kDuplas').textContent=S.docas.filter(d=>d.dupla).length;
    document.getElementById('kMotos').textContent=S.docas.filter(d=>d.motorista).length;
    document.getElementById('kFila').textContent=S.pool_motos.length;
    document.getElementById('kPct').textContent=pct+'%';
    document.getElementById('kSub').textContent=`${S.desc.toLocaleString()} / ${S.meta.toLocaleString()}`;
    document.getElementById('kBar').style.width=pct+'%';
    document.getElementById('metaInput').value=S.meta;
    const pill=document.getElementById('tpill');
    if(ativos.length){pill.textContent='● OPERAÇÃO ATIVA';pill.className='tpill on';}
    else{pill.textContent='● INATIVO';pill.className='tpill off';}

    // Pool duplas
    document.getElementById('cntDuplas').textContent=S.pool_duplas.length;
    document.getElementById('poolDuplas').innerHTML=S.pool_duplas.length
      ?S.pool_duplas.map(d=>`<div class="di" draggable="true" data-id="${d.id}" data-tipo="dupla"><div class="di-av" style="background:${d.color}">${ini(d.nomes)}</div><div class="di-info"><div class="di-name">${d.nomes}</div><div class="di-sub">${(d.pkgs||0).toLocaleString()} pac.</div></div><span class="di-badge bdup">Dupla</span><span class="di-handle">⠿</span></div>`).join('')
      :'<div class="dl-empty">Nenhuma dupla<br>disponível</div>';

    // Pool motos
    document.getElementById('cntMotos').textContent=S.pool_motos.length;
    document.getElementById('poolMotos').innerHTML=S.pool_motos.length
      ?S.pool_motos.map(m=>`<div class="di" draggable="true" data-id="${m.id}" data-tipo="motorista"><div class="di-av" style="background:${m.color||'#58a6ff'}">${ini(m.nome)}</div><div class="di-info"><div class="di-name">${m.nome}</div><div class="di-sub">${m.empresa} · ${(m.pacotes||0).toLocaleString()} pac.</div></div><span class="di-badge bmot">Moto</span><span class="di-handle">⠿</span></div>`).join('')
      :'<div class="dl-empty">Nenhum motorista<br>disponível</div>';

    // Pool carretas
    document.getElementById('cntCarretas').textContent=S.pool_carretas.length;
    document.getElementById('poolCarretas').innerHTML=S.pool_carretas.length
      ?S.pool_carretas.map(c=>`<div class="di" draggable="true" data-id="${c.id}" data-tipo="carreta"><div class="di-av" style="background:#f85149;font-size:8px">🚛</div><div class="di-info"><div class="di-name">${c.label}</div><div class="di-sub">Pronta para alocar</div></div><span class="di-badge bcart">Carreta</span><span class="di-handle">⠿</span></div>`).join('')
      :'<div class="dl-empty">Crie carretas<br>abaixo</div>';

    // Docas
    const gr=document.getElementById('docasGrid');
    gr.innerHTML='';
    S.docas.forEach(d=>{
      const el=document.createElement('div');
      let sc='';
      if(d.status==='carreta')sc='s-cart';
      else if(d.status==='both')sc='s-both';
      else if(d.status==='dup')sc='s-dup';
      else if(d.status==='mot')sc='s-mot';
      el.className=`doca ${sc}`;
      el.dataset.docaId=d.id;
      const canDrop = d.status==='livre' || (d.status!=='carreta' && (!d.dupla||!d.motorista));
      if(canDrop) el.dataset.droppable='true';
      const num=`<div class="doca-num">DOCA ${d.id}</div>`;

      if(d.status==='livre'){
        el.innerHTML=`${num}<div class="doca-badge db-livre">Livre</div><div class="doca-drop-hint">↓ arraste aqui</div>`;
      } else {
        const ms=d.inicio?Date.now()-new Date(d.inicio).getTime():0;
        const hd=ms/3600000;
        let slots='';
        if(d.dupla)    slots+=`<div class="slot slot-g"><span class="slot-dot" style="background:${d.dupla.color}"></span><span class="slot-name">${d.dupla.nomes}</span><button class="slot-rm" onclick="G.removerDupla(${d.id})">✕</button></div>`;
        if(d.motorista)slots+=`<div class="slot slot-b"><span class="slot-dot" style="background:${d.motorista.color||'#58a6ff'}"></span><span class="slot-name">${d.motorista.nome}</span><button class="slot-rm" onclick="G.removerMoto(${d.id})">✕</button></div>`;
        if(d.carreta)  slots+=`<div class="slot slot-r"><span class="slot-dot" style="background:var(--red)"></span><span class="slot-name">${d.carreta.label}</span></div>`;
        const badge=d.status==='both'?'db-both':d.status==='dup'?'db-dup':d.status==='mot'?'db-mot':'db-cart';
        const label=d.status==='both'?'Ativo':d.status==='dup'?'Dupla':d.status==='mot'?'Motorista':'Carreta';
        const extraDrop=canDrop?`<div class="doca-drop-hint">+ ${!d.dupla?'dupla ':''}${!d.motorista?'moto ':''}${!d.carreta?'carreta':''}</div>`:'';
        // Produtividade simplificada: só mostra pacotes declarados e tempo
        const pkgsTot=(d.dupla?.pkgs||0)+(d.motorista?.pacotes||0);
        const effTxt=pkgsTot>0&&hd>0.016?`${Math.round(pkgsTot/hd)} pac/h est.`:'aguardando';
        el.innerHTML=`${num}<div class="doca-badge ${badge}">${label}</div><div class="doca-slots">${slots}</div><div class="doca-timer">${fmtMs(ms)}</div>${pkgsTot>0?`<div class="doca-eff">${effTxt}</div>`:''}<div class="doca-drop-hint" style="font-size:8px">${extraDrop?extraDrop.replace(/<[^>]+>/g,''):''}</div><div class="doca-actions"><button class="btn r" onclick="G.finalizarDoca(${d.id})">✓ Finalizar</button></div>`;
      }
      gr.appendChild(el);
    });

    // Drag-drop setup
    setTimeout(()=>{
      document.querySelectorAll('.di[draggable],.fila-card[draggable]').forEach(item=>{
        item.addEventListener('dragstart',e=>{ e.dataTransfer.setData('did',item.dataset.id); e.dataTransfer.setData('dtipo',item.dataset.tipo); item.classList.add('dragging'); });
        item.addEventListener('dragend',()=>item.classList.remove('dragging'));
      });
      document.querySelectorAll('.doca[data-droppable]').forEach(doca=>{
        doca.addEventListener('dragover',e=>{ e.preventDefault(); doca.classList.add('over'); });
        doca.addEventListener('dragleave',()=>doca.classList.remove('over'));
        doca.addEventListener('drop',e=>{ e.preventDefault(); doca.classList.remove('over'); G.alocar(parseInt(e.dataTransfer.getData('did')),e.dataTransfer.getData('dtipo'),parseInt(doca.dataset.docaId)); });
      });
    },0);

    // Produtividade — apenas ops concluídas do dia de hoje
    this.renderProd();

    // Fila
    const fs=document.getElementById('filaScroll');
    document.getElementById('filaTotal').textContent=S.pool_motos.length;
    if(!S.pool_motos.length){
      fs.innerHTML='<div class="fila-empty">Nenhum motorista na fila.</div>';
    } else {
      fs.innerHTML=S.pool_motos.map((m,i)=>`
        <div class="fila-card" draggable="true" data-id="${m.id}" data-tipo="motorista">
          <div class="fc-top">
            <div class="fc-av" style="background:${m.color||'#58a6ff'}">${ini(m.nome)}</div>
            <div style="flex:1;min-width:0"><div class="fc-nome">${m.nome}</div><div class="fc-emp">${m.empresa}</div></div>
          </div>
          <div class="fc-pkgs">${(m.pacotes||0).toLocaleString()} pacotes</div>
          <div class="fc-pos"><span class="pos-num">${i+1}º</span> na fila</div>
          <div class="fc-handle">⠿ arraste para uma doca</div>
        </div>`).join('');
    }

    // Log
    const lc=document.getElementById('logWrap');
    const lco={sys:'#58a6ff',reg:'#3fb950',call:'#e6a817',done:'#bc8cff',err:'#f85149'};
    lc.innerHTML=S.log.length?S.log.map(e=>`<div class="lentry"><div class="lt">${e.h}</div><div class="ld" style="background:${lco[e.tipo]||'#6e7681'}"></div><div class="lm">${e.msg}</div></div>`).join(''):'<div style="font-size:11px;color:var(--muted)">Aguardando eventos...</div>';

    CH.hr.data.datasets[0].data=Array.from({length:24},(_,i)=>S.historico[i]||0);
    CH.hr.update();
  },

  renderProd(){
    // Produtividade baseada APENAS em operações concluídas de hoje
    const todayStr=today();
    const opsHoje=S.ops.filter(o=>o.startTime&&o.startTime.startsWith(todayStr));
    const teams=[...new Set(opsHoje.map(o=>o.team))];
    const pg=document.getElementById('prodGrid');

    if(!teams.length){
      pg.innerHTML='<div class="prod-empty">Os dados aparecerão ao finalizar a primeira doca do dia.</div>';
      return;
    }

    const metrics=teams.map(t=>{
      const ops=opsHoje.filter(o=>o.team===t);
      const p=ops.reduce((a,o)=>a+o.packages,0);
      // Eficiência REAL: soma dos minutos efetivos de cada operação concluída
      const minTot=ops.reduce((a,o)=>a+(new Date(o.endTime)-new Date(o.startTime))/60000,0);
      const hTot=minTot/60;
      const eff=hTot>0?p/hTot:0;
      const avgMin=ops.length?minTot/ops.length:0;
      return {team:t, p, trucks:ops.length, hTot, eff, avgMin, ops};
    }).sort((a,b)=>b.p-a.p);

    // Linha de média para comparação
    const mediaEff=metrics.reduce((a,m)=>a+m.eff,0)/metrics.length||0;

    pg.innerHTML=metrics.map(m=>{
      const c=clr(m.team);
      const effStr=m.eff>0?Math.round(m.eff).toLocaleString('pt-BR'):'—';
      const rating=m.eff>0?(m.eff>=mediaEff*1.15?'↑ Acima da média':m.eff>=mediaEff*0.85?'→ Na média':'↓ Abaixo da média'):'';
      const ratingColor=m.eff>=mediaEff*1.15?'var(--green)':m.eff>=mediaEff*0.85?'var(--acc)':'var(--red)';
      return`<div class="pcard">
        <div class="pc-head">
          <div class="pc-av" style="background:${c}">${ini(m.team)}</div>
          <div class="pc-name" title="${m.team}">${m.team}</div>
          <div class="pc-doca" style="color:var(--acc)">${m.trucks}×🚛</div>
        </div>
        <div class="pc-metrics">
          <div class="pm">
            <div class="pm-v" style="color:var(--green)">${m.p.toLocaleString('pt-BR')}</div>
            <div class="pm-l">pacotes totais</div>
          </div>
          <div class="pm">
            <div class="pm-v" style="color:var(--acc)">${effStr}</div>
            <div class="pm-l">pac/h (real)</div>
          </div>
          <div class="pm">
            <div class="pm-v" style="color:var(--blue)">${m.avgMin>0?Math.round(m.avgMin)+'min':'—'}</div>
            <div class="pm-l">tempo médio/op</div>
          </div>
        </div>
        ${m.eff>0&&mediaEff>0?`<div style="font-size:9px;display:flex;align-items:center;gap:5px;color:${ratingColor};margin-top:2px">${rating} · média: ${Math.round(mediaEff).toLocaleString()} pac/h</div>`:'<div style="font-size:9px;color:var(--muted);margin-top:2px">Eficiência calculada após 1ª operação</div>'}
      </div>`;
    }).join('');
  }
};

// ═══════════════════════════════════════════════════
//  CAR — CARRETAS
// ═══════════════════════════════════════════════════
const CAR = {
  render(){
    const all=S.carretas_all;
    const ativas=all.filter(c=>c.status==='ativa');
    const fin=all.filter(c=>c.status==='done');
    document.getElementById('cTot').textContent=all.length;
    document.getElementById('cAtivas').textContent=ativas.length;
    document.getElementById('cFin').textContent=fin.length;
    const totPal=all.reduce((a,c)=>a+(c.paletes||0),0);
    const totMP=all.reduce((a,c)=>a+(c.mangapaletes||0),0);
    document.getElementById('cTotPal').textContent=totPal.toLocaleString();
    document.getElementById('cTotMP').textContent=totMP.toLocaleString();
    const avgMs=fin.length?fin.reduce((a,c)=>a+(new Date(c.fim)-new Date(c.inicio)),0)/fin.length:0;
    document.getElementById('cAvgT').textContent=avgMs?fmtMs(avgMs):'—';

    const gr=document.getElementById('cartGrid');
    if(!all.length){gr.innerHTML='<div style="grid-column:1/-1;padding:20px;text-align:center;font-size:11px;color:var(--muted)">Nenhuma carreta registrada. Crie uma na aba "Doca em Tempo Real".</div>';return;}
    gr.innerHTML=all.map(c=>{
      const isActive=c.status==='ativa';
      const dur=c.inicio?(isActive?Date.now()-new Date(c.inicio).getTime():new Date(c.fim)-new Date(c.inicio)):0;
      return`<div class="cart-card ${isActive?'active':''}">
        <div class="cart-title">
          <div class="cart-id">${c.label}</div>
          <div class="cart-status ${isActive?'cs-active':'cs-done'}">${isActive?'● Em doca':'✓ Finalizada'}</div>
        </div>
        ${c.inicio?`<div class="cart-timer">${fmtMs(dur)}</div>`:'<div style="font-size:11px;color:var(--muted)">Aguardando alocação</div>'}
        ${c.inicio?`<div class="cart-row"><span>Início</span><span>${fdt(c.inicio)}</span></div>`:''}
        ${c.fim?`<div class="cart-row"><span>Fim</span><span>${fdt(c.fim)}</span></div>`:''}
        ${c.docaId?`<div class="cart-row"><span>Doca</span><span>${c.docaId}</span></div>`:''}
        <div class="cart-row">
          <span>Paletes</span>
          <input type="number" class="cart-inp" value="${c.paletes||0}" min="0" onchange="CAR.updField(${c.id},'paletes',this.value)">
        </div>
        <div class="cart-row">
          <span>Manga-paletes</span>
          <input type="number" class="cart-inp" value="${c.mangapaletes||0}" min="0" onchange="CAR.updField(${c.id},'mangapaletes',this.value)">
        </div>
      </div>`;
    }).join('');
  },

  updField(id, field, val){
    const c=S.carretas_all.find(x=>x.id===id);
    if(c){ c[field]=parseInt(val)||0; sv(); this.render(); }
  },

  exportar(){
    const data=S.carretas_all.map(c=>{
      const dur=c.inicio&&c.fim?((new Date(c.fim)-new Date(c.inicio))/3600000).toFixed(2):'—';
      return{'Carreta':c.label,'Status':c.status==='done'?'Finalizada':'Ativa','Doca':c.docaId||'—','Início':c.inicio?fdt(c.inicio):'—','Fim':c.fim?fdt(c.fim):'—','Duração (h)':dur,'Paletes':c.paletes||0,'Manga-paletes':c.mangapaletes||0};
    });
    const ws=XLSX.utils.json_to_sheet(data);const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,'Carretas');
    XLSX.writeFile(wb,`carretas-${today()}.xlsx`);
  }
};

// ═══════════════════════════════════════════════════
//  HIST — HISTÓRICO
// ═══════════════════════════════════════════════════
const HIST = {
  init(){
    const el=document.getElementById('histMes');
    el.value=today().slice(0,7);
  },
  render(){
    const mes=document.getElementById('histMes')?.value||today().slice(0,7);
    const dias=Object.values(S.historico_dias).filter(d=>d.date.startsWith(mes)).sort((a,b)=>b.date.localeCompare(a.date));
    const gr=document.getElementById('histGrid');
    if(!dias.length){
      gr.innerHTML='<div style="grid-column:1/-1;padding:20px;text-align:center;font-size:11px;color:var(--muted)">Nenhum dia fechado neste mês.</div>';
      return;
    }
    gr.innerHTML=dias.map(d=>`
      <div class="hist-card">
        <div class="hist-date">${new Date(d.date+'T12:00:00').toLocaleDateString('pt-BR',{weekday:'long',day:'2-digit',month:'long',year:'numeric'})}</div>
        <div class="hist-row"><span class="hist-k">Pacotes descarregados</span><span class="hist-v">${(d.desc||0).toLocaleString('pt-BR')}</span></div>
        <div class="hist-row"><span class="hist-k">Meta do dia</span><span class="hist-v">${(d.meta||0).toLocaleString('pt-BR')}</span></div>
        <div class="hist-row"><span class="hist-k">Caminhões</span><span class="hist-v">${d.caminhoes||0}</span></div>
        <div class="hist-row"><span class="hist-k">Duplas ativas</span><span class="hist-v">${d.duplas||0}</span></div>
        <div class="hist-row"><span class="hist-k">Carretas</span><span class="hist-v">${d.carretas||0}</span></div>
        <div class="hist-row"><span class="hist-k">Atingimento da meta</span><span class="hist-v" style="color:${d.desc>=d.meta?'var(--green)':'var(--red)'}">${d.meta?Math.round((d.desc/d.meta)*100):0}%</span></div>
        <button class="btn r" onclick="HIST.remover('${d.date}')" style="margin-top:8px;font-size:10px;width:100%;justify-content:center">✕ Remover</button>
      </div>`).join('');
  },
  remover(date){ delete S.historico_dias[date]; sv(); this.render(); },
  limpar(){ if(confirm('Apagar todo o histórico?')){ S.historico_dias={}; sv(); this.render(); } }
};

// ═══════════════════════════════════════════════════
//  D — DASHBOARD
// ═══════════════════════════════════════════════════
const D = {
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
    S.filteredOps=S.ops.filter(o=>{ const d=new Date(o.startTime);return d>=st&&d<en; });
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
    clr(t);document.getElementById('opTeam').value='';document.getElementById('opPkgs').value='';
    sv();this.filter(document.getElementById('fPeriod').value);this.fillDL();
  },
  fillDL(){ const dl=document.getElementById('teamsDL'); dl.innerHTML=[...new Set(S.ops.map(o=>o.team))].map(t=>`<option value="${t}">`).join(''); },
  metrics(){
    const teams=[...new Set(S.filteredOps.map(o=>o.team))],m={};
    teams.forEach(t=>{
      const ops=S.filteredOps.filter(o=>o.team===t);
      const p=ops.reduce((a,o)=>a+o.packages,0);
      // Eficiência real: soma das horas efetivamente trabalhadas (startTime→endTime)
      const h=ops.reduce((a,o)=>a+(new Date(o.endTime)-new Date(o.startTime))/36e5,0);
      m[t]={p,trucks:ops.length,h,eff:h>0?p/h:0,avgH:ops.length?h/ops.length:0};
    });
    return m;
  },
  render(){
    const m=this.metrics();const teams=Object.keys(m);const cols=teams.map(t=>clr(t));
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
    } else { document.getElementById('dFast').textContent='—';document.getElementById('dSlow').textContent='—'; }

    const srtV=Object.entries(m).sort((a,b)=>b[1].p-a[1].p);
    const srtE=Object.entries(m).sort((a,b)=>b[1].eff-a[1].eff);
    const inits=teams.map(t=>ini(t));
    // Average efficiency for reference line
    const avgEff=avg;
    function upd(ch,labels,data,bgs){ ch.data.labels=labels;ch.data.datasets[0].data=data;ch.data.datasets[0].backgroundColor=bgs;ch.update(); }
    upd(CH.vol,srtV.map(([t])=>ini(t)),srtV.map(([,v])=>v.p),srtV.map(([t])=>clr(t)+'cc'));
    // Efficiency chart with average line
    CH.eff.data.labels=srtE.map(([t])=>ini(t));
    CH.eff.data.datasets[0].data=srtE.map(([,v])=>+v.eff.toFixed(1));
    CH.eff.data.datasets[0].backgroundColor=srtE.map(([t])=>clr(t)+'cc');
    CH.eff.data.datasets[1].data=srtE.map(()=>+avgEff.toFixed(1));
    CH.eff.update();
    upd(CH.trk,inits,teams.map(t=>m[t].trucks),cols.map(c=>c+'cc'));
    CH.dist.data.labels=inits;CH.dist.data.datasets[0].data=teams.map(t=>m[t].p);CH.dist.data.datasets[0].backgroundColor=cols;CH.dist.update();
    const hrs=Array.from({length:11},(_,i)=>`${String(i+13).padStart(2,'0')}:00`);
    CH.line.data.labels=hrs;
    CH.line.data.datasets=teams.map(t=>({label:ini(t),data:hrs.map((_,idx)=>{const h=idx+13;return S.filteredOps.filter(o=>o.team===t&&new Date(o.endTime).getHours()<h+1).reduce((a,o)=>a+o.packages,0);}),borderColor:clr(t),backgroundColor:clr(t)+'22',fill:false,pointBackgroundColor:clr(t)}));
    CH.line.update();
    document.getElementById('lineLegend').innerHTML=teams.map(t=>`<div class="ch-li"><div class="ch-sw" style="background:${clr(t)}"></div><span>${ini(t)}: ${t}</span></div>`).join('');
    const tb=document.querySelector('#opsTable tbody');
    if(!S.filteredOps.length){tb.innerHTML=`<tr><td colspan="8" style="text-align:center;color:var(--muted);padding:16px">Nenhuma operação encontrada.</td></tr>`;this.fillDL();return;}
    const allEff=Object.values(m).map(x=>x.eff).sort((a,b)=>b-a);
    const q1=allEff[0]||1, q3=allEff[Math.floor(allEff.length*2/3)]||0;
    tb.innerHTML=[...S.filteredOps].sort((a,b)=>new Date(b.startTime)-new Date(a.startTime)).map(op=>{
      const dur=(new Date(op.endTime)-new Date(op.startTime))/36e5;
      const eff=dur>0?op.packages/dur:0;
      const c=clr(op.team);
      const rating=eff>=q1*.9?'<span class="badge g">A+</span>':eff>=q3?'<span class="badge a">B</span>':'<span class="badge r">C</span>';
      const tb2=op.tipo==='dupla'?'<span class="badge a">Dupla</span>':op.tipo==='motorista'?'<span class="badge b">Motorista</span>':'<span style="font-size:10px;color:var(--muted)">Manual</span>';
      return`<tr><td><span style="width:18px;height:18px;border-radius:3px;background:${c};display:inline-flex;align-items:center;justify-content:center;font-size:8px;font-weight:700;color:#000;margin-right:6px;vertical-align:middle">${ini(op.team)}</span>${op.team}</td><td>${tb2}</td><td class="mono">${fdt(op.startTime)}</td><td class="mono">${fdt(op.endTime)}</td><td class="mono">${op.packages.toLocaleString('pt-BR')}</td><td class="mono">${dur.toFixed(2)}</td><td><span class="badge g">${eff.toFixed(1)} p/h</span></td><td>${rating}</td></tr>`;
    }).join('');
    this.fillDL();
  },
  exportar(){
    const data=S.filteredOps.map(op=>{ const dur=(new Date(op.endTime)-new Date(op.startTime))/36e5; return{'Nome':op.team,'Tipo':op.tipo||'—','Início':fdt(op.startTime),'Término':fdt(op.endTime),'Pacotes':op.packages,'Duração(h)':dur.toFixed(2),'Eficiência(pac/h)':dur>0?(op.packages/dur).toFixed(1):'0'}; });
    const ws=XLSX.utils.json_to_sheet(data);const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,'Operações');
    XLSX.writeFile(wb,`operacoes-${today()}.xlsx`);
  },
  limpar(){ document.getElementById('modalOk').classList.add('open'); },
  confirmar(){ S.ops=[];S.filteredOps=[];S.teamColors={};S.cidx=0;closeModal('modalOk');sv();this.render(); }
};

// ═══════════════════════════════════════════════════
//  M — PORTAL MOTORISTA
// ═══════════════════════════════════════════════════
const M = {
  addEmpresa(){
    const n=document.getElementById('mfNovaEmp').value.trim(); if(!n) return;
    if(!S.empresas.includes(n)) S.empresas.push(n);
    document.getElementById('mfNovaEmp').value='';
    fillEmpSels(); sv(); document.getElementById('mfEmp').value=n;
  },
  registrar(){
    const nome=document.getElementById('mfNome').value.trim();
    const wpp=document.getElementById('mfWpp').value.trim();
    const emp=document.getElementById('mfEmp').value;
    const pkgs=parseInt(document.getElementById('mfPkgs').value)||0;
    if(!nome||!pkgs){ alert('Preencha nome e pacotes.'); return; }
    const c=clr(nome);
    const mot={id:Date.now(),nome,wpp,empresa:emp,pacotes:pkgs,status:'fila',docaId:null,ini:null,color:c};
    S.motorista=mot;
    S.pool_motos.push({id:mot.id,nome:mot.nome,empresa:mot.empresa,pacotes:mot.pacotes,color:c,tipo:'motorista'});
    addLog('reg',`${nome} (${emp}) registrou chegada — ${pkgs.toLocaleString()} pac.`);
    sv(); G.render(); this.render();
  },
  iniciar(){
    const m=S.motorista; if(!m||m.status!=='chamado') return;
    m.status='desc'; m.ini=new Date().toISOString();
    const doca=S.docas.find(d=>d.motorista&&d.motorista.id===m.id);
    if(doca&&!doca.inicio) doca.inicio=m.ini;
    addLog('sys',`${m.nome} iniciou descarga — Doca ${m.docaId}.`);
    sv(); this.render(); G.render();
  },
  finalizar(){
    const m=S.motorista; if(!m||m.status!=='desc') return;
    m.status='fim';
    const doca=S.docas.find(d=>d.motorista&&d.motorista.id===m.id);
    if(doca) G.finalizarDoca(doca.id);
    else if(m.ini){ const fim=new Date().toISOString(); S.ops.push({id:Date.now(),team:m.nome,tipo:'motorista',startTime:m.ini,endTime:fim,packages:m.pacotes}); S.desc+=m.pacotes; }
    addLog('done',`${m.nome} finalizou descarga.`);
    sv(); this.render();
  },
  novo(){ S.motorista=null; ['mfNome','mfWpp','mfPkgs'].forEach(id=>document.getElementById(id).value=''); sv(); this.render(); },
  updTimer(){
    if(!S.motorista||S.motorista.status!=='desc') return;
    const el=document.getElementById('mTimer'); if(!el) return;
    el.textContent=fmtMs(Date.now()-new Date(S.motorista.ini).getTime());
  },
  render(){
    const m=S.motorista;
    document.getElementById('mForm').style.display=m?'none':'block';
    document.getElementById('mStatus').style.display=m?'block':'none';
    if(!m) return;
    document.getElementById('mNomeDisp').textContent=`${m.nome} · ${m.empresa} · ${m.pacotes.toLocaleString()} pac.`;
    ['msWait','msCall','msDesc','msFim'].forEach(id=>document.getElementById(id).style.display='none');
    if(m.status==='fila'){
      document.getElementById('msWait').style.display='block';
      const pos=S.pool_motos.findIndex(x=>x.id===m.id)+1;
      document.getElementById('mPos').textContent=pos>0?pos:'—';
    } else if(m.status==='chamado'){
      document.getElementById('msCall').style.display='block';
      document.getElementById('mDocaNum').textContent=`DOCA ${m.docaId}`;
    } else if(m.status==='desc'){
      document.getElementById('msDesc').style.display='block'; this.updTimer();
    } else if(m.status==='fim'){
      document.getElementById('msFim').style.display='block';
    }
  }
};

// ═══════════════════════════════════════════════════
//  INIT
// ═══════════════════════════════════════════════════
document.addEventListener('DOMContentLoaded',()=>{
  ld(); initCharts(); fillEmpSels(); D.fillDL();
  const histMes=document.getElementById('histMes');
  if(histMes) histMes.value=today().slice(0,7);
  setInterval(()=>document.getElementById('clock').textContent=new Date().toLocaleTimeString('pt-BR'),1000);
  setInterval(()=>{
    if(document.getElementById('sc-app').classList.contains('active')){
      G.render();
      // Atualizar carretas ativas
      S.docas.forEach(d=>{ if(d.carreta&&d.status==='carreta'){ const ca=S.carretas_all.find(c=>c.id===d.carreta.id); if(ca&&ca.status==='ativa') CAR.render(); } });
    }
    if(document.getElementById('sc-moto').classList.contains('active')){
      M.updTimer();
      if(S.motorista?.status==='fila'){ const pos=S.pool_motos.findIndex(x=>x.id===S.motorista.id)+1; const el=document.getElementById('mPos'); if(el)el.textContent=pos>0?pos:'—'; }
    }
  },1000);
});
</script>
</body>
</html>
