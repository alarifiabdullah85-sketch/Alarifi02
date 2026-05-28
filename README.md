<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>صندوق الأقربين — نظام الإدارة المتكامل</title>
<link href="https://fonts.googleapis.com/css2?family=Cairo:wght@300;400;600;700;900&family=Tajawal:wght@300;400;500;700;800&display=swap" rel="stylesheet">
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<style>
:root{
  --ink:#faf7f2; --surface:#ffffff; --surface-2:#f6f1ea; --surface-3:#ede4d8; --surface-4:#e2d3c0;
  --border:rgba(124,82,40,.13); --border-2:rgba(124,82,40,.24);
  --gold:#9c6817; --gold-d:#7a4e0c; --gold-l:#c08a30; --gold-p:rgba(156,104,23,.10); --gold-g:rgba(156,104,23,.2);
  --teal:#2f7d6b; --teal-p:rgba(47,125,107,.1);
  --blue:#3a6ea5; --blue-p:rgba(58,110,165,.1);
  --violet:#7a5ba0; --violet-p:rgba(122,91,160,.1);
  --rose:#b54a3a; --rose-p:rgba(181,74,58,.1);
  --green:#3f8a4a; --green-p:rgba(63,138,74,.1);
  --amber:#c4881e; --amber-p:rgba(196,136,30,.1);
  --text:#2a1c0c; --text-dim:#6a5238; --text-muted:#9d8568;
  --q1c:#2f7d6b; --q2c:#3a6ea5; --q3c:#7a5ba0; --q4c:#b54a3a;
  --shadow:0 4px 22px rgba(124,82,40,.13); --shadow-xl:0 14px 50px rgba(124,82,40,.2);
}
*{margin:0;padding:0;box-sizing:border-box}
html{scroll-behavior:smooth}
body{font-family:'Tajawal',sans-serif;background:var(--ink);color:var(--text);direction:rtl;min-height:100vh;overflow-x:hidden}
body::before{content:'';position:fixed;inset:0;pointer-events:none;z-index:0;
  background:radial-gradient(ellipse 50% 40% at 12% 4%,rgba(156,104,23,.04),transparent 60%),
            radial-gradient(ellipse 45% 45% at 88% 90%,rgba(156,104,23,.03),transparent 60%);}

/* ════ LOGIN ════ */
#login{position:fixed;inset:0;z-index:9999;display:flex;align-items:center;justify-content:center;
  background:linear-gradient(160deg,#2a1605,#3d2208 45%,#1c0f04);font-family:'Tajawal',sans-serif}
#login.gone{display:none}
.login-card{width:380px;max-width:92vw;padding:46px 40px 40px;border-radius:24px;position:relative;overflow:hidden;
  background:rgba(255,248,240,.05);border:1px solid rgba(156,104,23,.3);
  box-shadow:0 30px 90px rgba(0,0,0,.6);backdrop-filter:blur(22px)}
.login-card::before{content:'';position:absolute;inset:0;pointer-events:none;
  background:radial-gradient(ellipse 80% 55% at 50% 0%,rgba(156,104,23,.13),transparent 65%)}
.login-logo{width:74px;height:74px;border-radius:20px;margin:0 auto 16px;position:relative;
  background:linear-gradient(135deg,var(--gold-d),var(--gold-l));display:flex;align-items:center;justify-content:center;
  font-size:34px;box-shadow:0 10px 30px var(--gold-g)}
.login-h{text-align:center;font-size:1.3rem;font-weight:900;font-family:'Cairo',sans-serif;color:#fff;margin-bottom:4px;position:relative}
.login-s{text-align:center;font-size:.76rem;color:rgba(255,255,255,.5);margin-bottom:30px;position:relative}
.lf{margin-bottom:16px;position:relative}
.lf label{display:block;font-size:.7rem;font-weight:700;color:rgba(255,255,255,.55);text-transform:uppercase;letter-spacing:.5px;margin-bottom:7px}
.lf input{width:100%;background:rgba(255,255,255,.06);border:1.5px solid rgba(255,255,255,.12);border-radius:11px;
  padding:12px 14px;font-family:'Tajawal',sans-serif;font-size:.95rem;color:#fff;direction:rtl;text-align:right;transition:.2s}
.lf input::placeholder{color:rgba(255,255,255,.25)}
.lf input:focus{outline:none;border-color:var(--gold-l);box-shadow:0 0 0 3px rgba(156,104,23,.2)}
.login-btn{width:100%;padding:13px;margin-top:6px;border:none;border-radius:11px;cursor:pointer;position:relative;
  background:linear-gradient(135deg,var(--gold-d),var(--gold-l));color:#fff;
  font-family:'Cairo',sans-serif;font-size:1rem;font-weight:900;letter-spacing:.5px;transition:.2s}
.login-btn:hover{opacity:.92;transform:translateY(-1px)}
.login-err{text-align:center;color:#f0a088;font-size:.78rem;margin-top:12px;min-height:18px;position:relative}
.login-roles{margin-top:22px;padding-top:16px;border-top:1px solid rgba(255,255,255,.08);position:relative;display:flex;flex-direction:column;gap:6px}
.rh{display:flex;align-items:center;gap:9px;padding:7px 11px;background:rgba(255,255,255,.03);border-radius:8px;font-size:.7rem;color:rgba(255,255,255,.4)}
.rh b{padding:2px 9px;border-radius:14px;font-size:.66rem;white-space:nowrap}

/* ════ APP SHELL ════ */
.shell{display:flex;min-height:100vh;position:relative;z-index:1}
.shell.hidden{display:none}

/* ════ SIDEBAR ════ */
.sidebar{width:270px;flex-shrink:0;position:sticky;top:0;height:100vh;overflow-y:auto;z-index:200;transition:transform .3s;
  background:linear-gradient(180deg,#3a2208,#26150420 100%),linear-gradient(180deg,#3a2208,#1c0f04);
  border-left:1px solid var(--border-2);display:flex;flex-direction:column}
.sidebar::-webkit-scrollbar{width:5px}
.sidebar::-webkit-scrollbar-thumb{background:rgba(192,138,48,.3);border-radius:10px}
.sb-brand{padding:22px 20px;display:flex;align-items:center;gap:12px;border-bottom:1px solid rgba(255,255,255,.07)}
.sb-brand-ic{width:48px;height:48px;border-radius:13px;flex-shrink:0;background:linear-gradient(135deg,var(--gold-d),var(--gold-l));
  display:flex;align-items:center;justify-content:center;font-size:23px;box-shadow:0 0 18px var(--gold-g)}
.sb-brand h1{font-size:1rem;font-weight:900;font-family:'Cairo',sans-serif;color:#fff;line-height:1.2}
.sb-brand p{font-size:.66rem;color:rgba(255,255,255,.5);margin-top:2px}
.sb-nav{flex:1;padding:12px 12px;display:flex;flex-direction:column;gap:2px}
.sb-lbl{font-size:.62rem;font-weight:800;color:rgba(255,255,255,.3);text-transform:uppercase;letter-spacing:1px;padding:14px 14px 7px}
.sb-btn{display:flex;align-items:center;gap:11px;padding:11px 14px;border:none;background:none;cursor:pointer;width:100%;
  font-family:'Tajawal',sans-serif;font-size:.85rem;font-weight:600;color:rgba(255,255,255,.7);
  border-radius:10px;transition:.18s;text-align:right;position:relative}
.sb-btn .ic{font-size:1.05rem;width:22px;text-align:center;flex-shrink:0}
.sb-btn:hover{background:rgba(255,255,255,.06);color:#fff}
.sb-btn.active{background:linear-gradient(135deg,rgba(192,138,48,.22),rgba(192,138,48,.07));color:var(--gold-l);font-weight:800}
.sb-btn.active::before{content:'';position:absolute;right:0;top:50%;transform:translateY(-50%);width:3.5px;height:22px;border-radius:4px;background:var(--gold-l)}
.sb-chip{margin-right:auto;background:rgba(255,255,255,.1);border-radius:10px;padding:1px 8px;font-size:.64rem;color:rgba(255,255,255,.55)}
.sb-btn.fin{background:linear-gradient(135deg,var(--gold-d),var(--gold-l));color:#fff;font-weight:900;font-family:'Cairo',sans-serif;box-shadow:0 3px 12px var(--gold-g);margin-top:3px}
.sb-btn.fin:hover{opacity:.93;color:#fff}
.sb-btn.fin.active::before{display:none}
.sb-foot{padding:13px 12px;border-top:1px solid rgba(255,255,255,.07)}
.sb-user{display:flex;align-items:center;gap:10px;padding:10px 12px;background:rgba(255,255,255,.05);border-radius:10px;margin-bottom:8px}
.sb-av{width:36px;height:36px;border-radius:50%;flex-shrink:0;background:linear-gradient(135deg,var(--gold-d),var(--gold-l));display:flex;align-items:center;justify-content:center;font-size:16px;color:#fff}
.sb-uinfo{flex:1;min-width:0}
.sb-uname{font-size:.8rem;font-weight:700;color:#fff;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.sb-urole{font-size:.65rem;color:rgba(255,255,255,.5)}
.sb-logout{width:100%;padding:9px;border-radius:8px;cursor:pointer;border:1px solid rgba(181,74,58,.4);background:rgba(181,74,58,.12);color:#e8a088;font-family:'Tajawal',sans-serif;font-size:.78rem;font-weight:700;transition:.2s}
.sb-logout:hover{background:rgba(181,74,58,.24)}

/* mobile toggle */
.sb-tog{display:none;position:fixed;top:13px;right:13px;z-index:400;width:44px;height:44px;border-radius:11px;border:none;cursor:pointer;background:linear-gradient(135deg,var(--gold-d),var(--gold-l));color:#fff;font-size:20px;box-shadow:var(--shadow)}
.sb-ov{display:none;position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:150}
.sb-ov.show{display:block}

/* ════ CONTENT ════ */
.content{flex:1;min-width:0;display:flex;flex-direction:column}
.topbar{position:sticky;top:0;z-index:90;padding:14px 26px;display:flex;align-items:center;justify-content:space-between;gap:14px;
  background:rgba(255,248,240,.96);backdrop-filter:blur(20px);border-bottom:1px solid var(--border-2)}
.topbar-title{font-size:1.05rem;font-weight:900;font-family:'Cairo',sans-serif;color:var(--gold)}
.topbar-r{display:flex;align-items:center;gap:10px}
.ybadge{padding:5px 15px;border:1.5px solid var(--gold);border-radius:30px;color:var(--gold);font-size:.82rem;font-weight:800;font-family:'Cairo',sans-serif;letter-spacing:1px;white-space:nowrap}
.main{max-width:1480px;margin:0 auto;padding:24px 24px 60px;width:100%}
.sec{display:none;animation:fade .3s ease}
.sec.active{display:block}
@keyframes fade{from{opacity:0;transform:translateY(12px)}to{opacity:1;transform:translateY(0)}}

/* ════ GENERIC ════ */
.card{background:var(--surface);border:1px solid var(--border);border-radius:14px;padding:20px;box-shadow:var(--shadow)}
.card-t{font-size:.76rem;font-weight:800;color:var(--text-dim);text-transform:uppercase;letter-spacing:.6px;margin-bottom:15px;display:flex;align-items:center;gap:7px}
.pg-title{font-size:1.1rem;font-weight:900;font-family:'Cairo',sans-serif;color:var(--gold);margin-bottom:6px}
.pg-sub{font-size:.8rem;color:var(--text-muted);margin-bottom:20px}
.btn{padding:10px 18px;border-radius:9px;border:none;cursor:pointer;font-family:'Tajawal',sans-serif;font-size:.84rem;font-weight:800;display:inline-flex;align-items:center;gap:7px;justify-content:center;transition:.2s}
.btn-gold{background:linear-gradient(135deg,var(--gold-d),var(--gold-l));color:#fff}
.btn-gold:hover{opacity:.92;transform:translateY(-1px)}
.btn-teal{background:var(--teal-p);color:var(--teal);border:1px solid rgba(47,125,107,.3)}
.btn-teal:hover{background:rgba(47,125,107,.2)}
.btn-rose{background:var(--rose-p);color:var(--rose);border:1px solid rgba(181,74,58,.3)}
.btn-rose:hover{background:rgba(181,74,58,.2)}
.btn-blue{background:var(--blue-p);color:var(--blue);border:1px solid rgba(58,110,165,.3)}
.btn-blue:hover{background:rgba(58,110,165,.2)}
.btn-violet{background:var(--violet-p);color:var(--violet);border:1px solid rgba(122,91,160,.3)}
.btn-violet:hover{background:rgba(122,91,160,.2)}
.field{display:flex;flex-direction:column;gap:5px}
.field label{font-size:.7rem;font-weight:700;color:var(--text-muted);text-transform:uppercase;letter-spacing:.5px}
.ctrl{width:100%;background:var(--surface-3);border:1.5px solid var(--border-2);border-radius:8px;padding:9px 12px;font-family:'Tajawal',sans-serif;font-size:.85rem;color:var(--text);direction:rtl;text-align:right;transition:.2s}
.ctrl:focus{outline:none;border-color:var(--gold);box-shadow:0 0 0 3px var(--gold-p)}
textarea.ctrl{resize:vertical}
.badge{display:inline-flex;align-items:center;gap:4px;padding:3px 10px;border-radius:20px;font-size:.68rem;font-weight:700;white-space:nowrap}
.lock-banner{display:flex;align-items:center;gap:11px;padding:12px 18px;background:var(--rose-p);border:1px solid rgba(181,74,58,.2);border-radius:10px;margin-bottom:18px;font-size:.82rem;color:var(--rose);font-weight:600}

/* ════ WELCOME ════ */
.welcome{display:flex;flex-direction:column;align-items:center;padding:30px 16px}
.w-hero{width:100%;max-width:960px;border-radius:24px;padding:44px 38px;text-align:center;position:relative;overflow:hidden;
  background:linear-gradient(135deg,rgba(90,50,15,.06),rgba(156,104,23,.03));border:1px solid var(--border-2)}
.w-hero::before{content:'';position:absolute;inset:0;pointer-events:none;background:radial-gradient(ellipse 55% 50% at 50% 0%,var(--gold-p),transparent 65%)}
.w-crest{width:86px;height:86px;border-radius:24px;margin:0 auto 20px;background:linear-gradient(135deg,var(--gold-d),var(--gold-l));display:flex;align-items:center;justify-content:center;font-size:40px;box-shadow:0 12px 40px var(--gold-g);position:relative}
.w-h1{font-size:1.9rem;font-weight:900;font-family:'Cairo',sans-serif;line-height:1.25;margin-bottom:10px;position:relative}
.w-sub{font-size:.98rem;color:var(--text-dim);max-width:600px;margin:0 auto;line-height:1.7;position:relative}
.w-greet{font-size:.92rem;color:var(--gold);font-weight:700;margin-top:18px;position:relative}
.w-stats{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;width:100%;max-width:960px;margin-top:24px}
.w-stat{background:var(--surface);border:1px solid var(--border);border-radius:16px;padding:22px 16px;text-align:center;border-top:3px solid;transition:.2s}
.w-stat:hover{transform:translateY(-3px);box-shadow:var(--shadow-xl)}
.w-stat-v{font-size:1.6rem;font-weight:900;font-family:'Cairo',sans-serif;line-height:1}
.w-stat-l{font-size:.76rem;color:var(--text-dim);margin-top:6px}
.w-cards{display:grid;grid-template-columns:repeat(3,1fr);gap:16px;width:100%;max-width:960px;margin-top:24px}
.w-card{background:var(--surface);border:1px solid var(--border);border-radius:16px;padding:22px;text-align:right;cursor:pointer;transition:.2s}
.w-card:hover{transform:translateY(-3px);box-shadow:var(--shadow-xl);border-color:var(--border-2)}
.w-card-ic{font-size:1.8rem;margin-bottom:10px}
.w-card-t{font-size:1rem;font-weight:800;font-family:'Cairo',sans-serif;margin-bottom:6px}
.w-card-d{font-size:.79rem;color:var(--text-muted);line-height:1.6}
.w-card-a{margin-top:12px;font-size:.8rem;font-weight:700;color:var(--gold)}

/* ════ KPI ════ */
.kpis{display:grid;gap:14px;margin-bottom:20px}
.kpis.k5{grid-template-columns:repeat(5,1fr)}
.kpis.k4{grid-template-columns:repeat(4,1fr)}
.kpis.k3{grid-template-columns:repeat(3,1fr)}
.kpi{background:var(--surface);border:1px solid var(--border);border-radius:14px;padding:17px 17px 13px;border-top:3px solid;position:relative;overflow:hidden;transition:.2s}
.kpi:hover{transform:translateY(-2px);box-shadow:var(--shadow)}
.kpi-wm{position:absolute;bottom:-6px;left:6px;font-size:3rem;opacity:.05;user-select:none}
.kpi-v{font-size:1.45rem;font-weight:900;font-family:'Cairo',sans-serif;line-height:1;margin-bottom:4px}
.kpi-l{font-size:.72rem;color:var(--text-dim)}
.kpi-bar{height:4px;background:var(--border-2);border-radius:4px;margin-top:10px;overflow:hidden}
.kpi-bar-f{height:100%;border-radius:4px;transition:width .8s cubic-bezier(.34,1.56,.64,1)}

/* ════ CHARTS ════ */
.ch-row{display:grid;grid-template-columns:1.6fr 1fr;gap:14px;margin-bottom:18px}
.ch-3{display:grid;grid-template-columns:repeat(3,1fr);gap:14px;margin-bottom:18px}
.ch-box{background:var(--surface);border:1px solid var(--border);border-radius:14px;padding:18px}
.ch-t{font-size:.74rem;font-weight:800;color:var(--text-dim);text-transform:uppercase;letter-spacing:.5px;margin-bottom:14px}
.ch-wrap{position:relative;height:220px}

/* ════ TABLE ════ */
.tbl-wrap{overflow-x:auto;border-radius:10px;border:1px solid var(--border)}
.tbl{width:100%;border-collapse:collapse;font-size:.78rem}
.tbl thead th{background:var(--surface-3);color:var(--text-dim);padding:10px 12px;font-weight:700;text-align:right;font-size:.7rem;text-transform:uppercase;letter-spacing:.4px;white-space:nowrap}
.tbl tbody td{padding:9px 12px;text-align:right;border-bottom:1px solid var(--border)}
.tbl tbody tr:hover td{background:rgba(156,104,23,.03)}
.tbl tbody tr:last-child td{border-bottom:none}
.tbl-foot{padding:11px 14px;font-size:.76rem;color:var(--text-dim);border-top:1px solid var(--border);display:flex;justify-content:space-between;flex-wrap:wrap;gap:8px}
.edit{outline:none;cursor:text;border-radius:4px;padding:2px 5px;transition:background .15s}
.edit:hover{background:rgba(156,104,23,.08)}
.edit:focus{background:rgba(156,104,23,.14);box-shadow:inset 0 0 0 1.5px var(--gold)}
.pctbar{display:flex;align-items:center;gap:7px}
.pctbar-m{height:4px;flex:1;background:var(--border-2);border-radius:4px;overflow:hidden;min-width:46px}
.pctbar-f{height:100%;border-radius:4px}

/* ════ MODAL ════ */
.modal-bg{position:fixed;inset:0;z-index:8500;background:rgba(0,0,0,.55);backdrop-filter:blur(6px);display:none;align-items:center;justify-content:center}
.modal-bg.open{display:flex}
.modal{background:var(--surface);border:1px solid var(--border-2);border-radius:18px;padding:28px;width:600px;max-width:95vw;box-shadow:var(--shadow-xl);max-height:90vh;overflow-y:auto;direction:rtl}
.modal-t{font-size:1rem;font-weight:900;font-family:'Cairo',sans-serif;color:var(--gold);margin-bottom:20px;display:flex;align-items:center;gap:8px}
.modal-grid{display:grid;grid-template-columns:1fr 1fr;gap:13px;margin-bottom:14px}
.modal-act{display:flex;gap:10px;margin-top:6px}

/* ════ TOAST ════ */
.toast{position:fixed;bottom:22px;left:50%;transform:translateX(-50%) translateY(120px);z-index:9000;
  background:var(--surface-4);border:1px solid var(--border-2);color:var(--text);padding:11px 22px;border-radius:40px;
  box-shadow:var(--shadow-xl);font-size:.85rem;font-weight:600;display:flex;align-items:center;gap:10px;transition:transform .4s cubic-bezier(.34,1.56,.64,1)}
.toast.show{transform:translateX(-50%) translateY(0)}
.toast-save{background:var(--gold);color:#fff;border:none;padding:5px 14px;border-radius:20px;font-family:'Tajawal',sans-serif;font-size:.75rem;font-weight:900;cursor:pointer}

/* ════ PROGRAM CARDS ════ */
.prog-list{display:flex;flex-direction:column;gap:12px}
.pc{background:var(--surface);border:1px solid var(--border);border-radius:14px;overflow:hidden;transition:box-shadow .2s}
.pc:hover{box-shadow:var(--shadow)}
.pc-head{display:flex;align-items:center;gap:10px;padding:14px 18px;cursor:pointer;user-select:none;transition:.15s}
.pc.open .pc-head{background:rgba(156,104,23,.04);border-bottom:1px solid var(--border)}
.pc-tag{padding:4px 13px;border-radius:20px;font-size:.76rem;font-weight:800;white-space:nowrap;flex-shrink:0}
.pc-title{flex:1;font-size:.87rem;font-weight:600}
.pc-meta{display:flex;align-items:center;gap:8px;flex-shrink:0}
.pill{display:flex;align-items:center;gap:4px;padding:4px 11px;border-radius:20px;font-size:.73rem;font-weight:700;white-space:nowrap}
.pill-gold{background:var(--gold-p);color:var(--gold)}
.pill-dim{background:rgba(156,104,23,.06);color:var(--text-dim)}
.pc-chev{width:26px;height:26px;border-radius:7px;background:rgba(156,104,23,.07);display:flex;align-items:center;justify-content:center;font-size:.7rem;color:var(--text-muted);transition:.3s;flex-shrink:0}
.pc.open .pc-chev{transform:rotate(180deg);background:var(--gold-p);color:var(--gold)}
.pc-body{display:none;padding:18px}
.pc.open .pc-body{display:block}
.pc-grid{display:grid;grid-template-columns:1fr 1fr;gap:18px}
.meta-box{background:var(--surface-2);border:1px solid var(--border);border-radius:10px;overflow:hidden;margin-bottom:12px}
.meta-r{display:flex;border-bottom:1px solid var(--border)}
.meta-r:last-child{border-bottom:none}
.meta-k{width:110px;padding:9px 12px;font-size:.71rem;font-weight:700;color:var(--text-muted);background:rgba(156,104,23,.04);border-left:1px solid var(--border);flex-shrink:0;display:flex;align-items:center}
.meta-v{padding:9px 12px;font-size:.81rem;font-weight:600;flex:1;display:flex;align-items:center}
.target-sec{background:var(--surface-2);border:1px solid var(--border);border-radius:10px;padding:14px;margin-bottom:12px}
.ts-t{font-size:.71rem;font-weight:800;color:var(--text-muted);text-transform:uppercase;letter-spacing:.5px;margin-bottom:12px}
.annual-row{display:flex;align-items:center;gap:10px;margin-bottom:12px;padding-bottom:12px;border-bottom:1px solid var(--border)}
.annual-row label{font-size:.77rem;color:var(--text-dim);width:120px;flex-shrink:0}
.annual-in{flex:1;background:var(--surface-3);border:1.5px solid var(--border-2);border-radius:8px;padding:8px 12px;font-family:'Tajawal',sans-serif;font-size:.9rem;font-weight:800;color:var(--gold);text-align:right}
.annual-in:focus{outline:none;border-color:var(--gold);box-shadow:0 0 0 3px var(--gold-p)}
.q-splits{display:grid;grid-template-columns:repeat(4,1fr);gap:8px}
.qsp{border-radius:8px;padding:10px;border:1px solid var(--border);background:rgba(156,104,23,.03);position:relative;overflow:hidden}
.qsp::before{content:'';position:absolute;top:0;right:0;left:0;height:2.5px}
.qsp.q1::before{background:var(--q1c)} .qsp.q2::before{background:var(--q2c)} .qsp.q3::before{background:var(--q3c)} .qsp.q4::before{background:var(--q4c)}
.qsp-l{font-size:.67rem;color:var(--text-muted);margin-bottom:5px;display:flex;justify-content:space-between}
.qsp-l span{font-weight:700}
.qsp.q1 .qsp-l span{color:var(--q1c)} .qsp.q2 .qsp-l span{color:var(--q2c)} .qsp.q3 .qsp-l span{color:var(--q3c)} .qsp.q4 .qsp-l span{color:var(--q4c)}
.qsp-in{width:100%;background:var(--surface-3);border:1.5px solid var(--border-2);border-radius:6px;padding:6px 8px;font-family:'Tajawal',sans-serif;font-size:.82rem;font-weight:700;color:var(--text);text-align:center}
.qsp-in:focus{outline:none;border-color:var(--gold)}
.qsp-ach{width:100%;background:var(--green-p);border:1.5px solid rgba(63,138,74,.25);border-radius:6px;padding:5px 8px;font-family:'Tajawal',sans-serif;font-size:.82rem;font-weight:700;color:var(--green);text-align:center;margin-top:5px}
.qsp-ach:focus{outline:none;border-color:var(--teal)}
.qsp-cap{font-size:.62rem;color:var(--text-muted);margin-top:4px;text-align:center}
.qbar{margin-top:5px;background:rgba(156,104,23,.08);border-radius:20px;height:4px;overflow:hidden}
.qbar-f{height:100%;border-radius:20px;transition:width .6s}
.prog-bar{margin-top:12px;background:rgba(156,104,23,.08);border-radius:20px;height:6px;overflow:hidden}
.prog-bar-f{height:100%;border-radius:20px;transition:width .8s cubic-bezier(.34,1.56,.64,1)}
.sub-list{display:flex;flex-direction:column;gap:7px;margin-bottom:12px}
.sub-item{background:var(--surface-2);border:1px solid var(--border);border-radius:8px;padding:10px 12px;display:flex;gap:9px}
.sub-dot{width:24px;height:24px;border-radius:6px;display:flex;align-items:center;justify-content:center;font-size:.68rem;font-weight:800;flex-shrink:0}
.sub-desc{font-size:.8rem;line-height:1.5}
.sub-chips{display:flex;gap:6px;flex-wrap:wrap;margin-top:5px}
.chip{padding:2px 9px;border-radius:10px;font-size:.67rem;font-weight:700}
.chip-teal{background:var(--teal-p);color:var(--teal)} .chip-gold{background:var(--gold-p);color:var(--gold-d)} .chip-dim{background:rgba(156,104,23,.06);color:var(--text-dim)}
.notes-box{background:var(--surface-2);border:1px solid var(--border);border-radius:10px;padding:14px;margin-top:12px}
.notes-area{width:100%;background:var(--surface-3);border:1.5px solid var(--border-2);border-radius:8px;padding:10px 12px;font-family:'Tajawal',sans-serif;font-size:.83rem;color:var(--text);direction:rtl;text-align:right;resize:vertical;margin-top:6px}
.notes-area:focus{outline:none;border-color:var(--violet);box-shadow:0 0 0 3px var(--violet-p)}
.assignee-tag{display:inline-flex;align-items:center;gap:5px;padding:3px 11px;border-radius:20px;background:var(--blue-p);color:var(--blue);font-size:.72rem;font-weight:700}

/* ════ OVERVIEW ════ */
.ov-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin-bottom:20px}
.ov-card{background:var(--surface);border:1px solid var(--border);border-radius:14px;padding:18px;position:relative;overflow:hidden;border-right:4px solid;transition:.2s}
.ov-card:hover{transform:translateY(-2px);box-shadow:var(--shadow-xl)}
.ov-num{font-size:1.8rem;font-weight:900;font-family:'Cairo',sans-serif;line-height:1}
.ov-lbl{font-size:.77rem;color:var(--text-dim);margin-top:4px}
.ov-sub{margin-top:9px;font-size:.71rem;color:var(--text-muted);padding-top:9px;border-top:1px solid var(--border)}

/* ════ BYLAWS ════ */
.bylaws-doc{max-width:900px;margin:0 auto;background:var(--surface);border:1px solid var(--border);border-radius:14px;padding:34px 40px;box-shadow:var(--shadow);line-height:1.9}
.bylaws-doc h2{font-family:'Cairo',sans-serif;color:var(--gold);font-size:1.15rem;margin:26px 0 12px;padding-bottom:8px;border-bottom:2px solid var(--border-2)}
.bylaws-doc h3{font-family:'Cairo',sans-serif;color:var(--teal);font-size:.95rem;margin:18px 0 8px}
.bylaws-doc h4{font-family:'Cairo',sans-serif;color:var(--text);font-size:.88rem;font-weight:800;margin:14px 0 6px}
.bylaws-doc p{font-size:.86rem;color:var(--text-dim);margin-bottom:8px}
.bylaws-doc ul{margin:6px 24px 12px;font-size:.85rem;color:var(--text-dim)}
.bylaws-doc li{margin-bottom:5px}
.bylaws-hero{text-align:center;padding:30px 0 20px;border-bottom:2px solid var(--border-2);margin-bottom:20px}
.bylaws-hero .crest{width:70px;height:70px;border-radius:18px;margin:0 auto 14px;background:linear-gradient(135deg,var(--teal),var(--blue));display:flex;align-items:center;justify-content:center;font-size:32px}
.bylaws-meta{display:inline-flex;gap:18px;flex-wrap:wrap;justify-content:center;margin-top:12px;font-size:.78rem;color:var(--text-muted)}

/* responsive */
@media(max-width:1180px){.kpis.k5{grid-template-columns:repeat(3,1fr)}.ch-row{grid-template-columns:1fr}.ch-3{grid-template-columns:1fr 1fr}.ov-grid{grid-template-columns:repeat(2,1fr)}}
@media(max-width:980px){.w-stats{grid-template-columns:repeat(2,1fr)}.w-cards{grid-template-columns:1fr}.kpis.k4{grid-template-columns:repeat(2,1fr)}}
@media(max-width:860px){.sidebar{position:fixed;top:0;right:0;transform:translateX(100%)}.sidebar.open{transform:translateX(0);box-shadow:-10px 0 40px rgba(0,0,0,.4)}.sb-tog{display:flex;align-items:center;justify-content:center}.topbar{padding-right:68px}.pc-grid{grid-template-columns:1fr}.w-h1{font-size:1.5rem}.modal-grid{grid-template-columns:1fr}.kpis.k5,.kpis.k3,.ch-3{grid-template-columns:1fr}.q-splits{grid-template-columns:repeat(2,1fr)}}
</style>
</head>
<body>

<!-- ═══════ LOGIN ═══════ -->
<div id="login">
  <div class="login-card">
    <div class="login-logo">🏛️</div>
    <div class="login-h">صندوق الأقربين</div>
    <div class="login-s">نظام الإدارة المتكامل — عائلة العريفي</div>
    <div class="lf"><label>اسم المستخدم</label><input type="text" id="lu" placeholder="أدخل اسم المستخدم"></div>
    <div class="lf"><label>الرمز السري</label><input type="password" id="lp" placeholder="••••••••" onkeydown="if(event.key==='Enter')doLogin()"></div>
    <button class="login-btn" onclick="doLogin()">🔐 تسجيل الدخول</button>
    <div class="login-err" id="lerr"></div>
    <div class="login-roles">
      <div class="rh"><b style="background:rgba(156,104,23,.18);color:#e0b060">المدير</b> admin / admin2026 — تحكم كامل</div>
      <div class="rh"><b style="background:rgba(63,138,74,.18);color:#7cc78a">المحاسب</b> hisab / hisab2026 — القسم المالي</div>
      <div class="rh"><b style="background:rgba(58,110,165,.18);color:#7eb0e8">موظف</b> saleh / saleh2026 — برامجه فقط</div>
    </div>
  </div>
</div>

<!-- ═══════ TOAST ═══════ -->
<div class="toast" id="toast"><span id="toast-msg"></span><button class="toast-save" id="toast-save" onclick="persist()" style="display:none">💾 حفظ</button></div>

<!-- ═══════ MOBILE TOGGLE ═══════ -->
<button class="sb-tog" onclick="toggleSB()">☰</button>
<div class="sb-ov" id="sb-ov" onclick="toggleSB()"></div>

<!-- ═══════ SHELL ═══════ -->
<div class="shell hidden" id="shell">

  <!-- SIDEBAR -->
  <aside class="sidebar" id="sidebar">
    <div class="sb-brand">
      <div class="sb-brand-ic">🏛️</div>
      <div><h1>صندوق الأقربين</h1><p>عائلة العريفي — حائل</p></div>
    </div>
    <nav class="sb-nav" id="sb-nav">
      <button class="sb-btn active" data-tab="home" onclick="showTab('home',this)"><span class="ic">🏠</span> الرئيسية</button>

      <div class="sb-lbl" data-grp="plan">الخطة التشغيلية</div>
      <button class="sb-btn" data-tab="overview" data-grp="plan" onclick="showTab('overview',this)"><span class="ic">📊</span> نظرة عامة</button>
      <button class="sb-btn" data-tab="goal1" data-grp="plan" onclick="showTab('goal1',this)"><span class="ic">🤝</span> التكافل <span class="sb-chip" id="chip-g1">0</span></button>
      <button class="sb-btn" data-tab="goal2" data-grp="plan" onclick="showTab('goal2',this)"><span class="ic">🌐</span> التواصل <span class="sb-chip" id="chip-g2">0</span></button>
      <button class="sb-btn" data-tab="goal3" data-grp="plan" onclick="showTab('goal3',this)"><span class="ic">📚</span> التطوير <span class="sb-chip" id="chip-g3">0</span></button>
      <button class="sb-btn" data-tab="kpis" data-grp="plan" onclick="showTab('kpis',this)"><span class="ic">📈</span> ملخص الأداء</button>

      <div class="sb-lbl" data-grp="mytasks" style="display:none">مهامي</div>
      <button class="sb-btn" data-tab="mytasks" data-grp="mytasks" onclick="showTab('mytasks',this)" style="display:none"><span class="ic">✅</span> برامجي المسندة <span class="sb-chip" id="chip-mytasks">0</span></button>

      <div class="sb-lbl" data-grp="admin">الإدارة</div>
      <button class="sb-btn" data-tab="manage" data-grp="admin" onclick="showTab('manage',this)"><span class="ic">⚙️</span> إدارة البرامج</button>
      <button class="sb-btn" data-tab="team" data-grp="admin" onclick="showTab('team',this)"><span class="ic">👥</span> الموظفون والإسناد</button>

      <div class="sb-lbl" data-grp="fin">القسم المالي</div>
      <button class="sb-btn fin" data-tab="finance" data-grp="fin" onclick="showTab('finance',this)"><span class="ic">💰</span> لوحة المحاسب</button>

      <div class="sb-lbl">المستندات</div>
      <button class="sb-btn" data-tab="bylaws" onclick="showTab('bylaws',this)"><span class="ic">📜</span> اللائحة الأساسية</button>
    </nav>
    <div class="sb-foot">
      <div class="sb-user">
        <div class="sb-av" id="sb-av">👤</div>
        <div class="sb-uinfo"><div class="sb-uname" id="sb-uname">—</div><div class="sb-urole" id="sb-urole">—</div></div>
      </div>
      <button class="sb-logout" onclick="doLogout()">🚪 تسجيل الخروج</button>
    </div>
  </aside>

  <!-- CONTENT -->
  <div class="content">
    <div class="topbar">
      <div class="topbar-title" id="topbar-title">🏠 الرئيسية</div>
      <div class="topbar-r"><span class="ybadge">2026 م</span></div>
    </div>

    <div class="main">

      <!-- ══════ HOME ══════ -->
      <div id="tab-home" class="sec active">
        <div class="welcome">
          <div class="w-hero">
            <div class="w-crest">🏛️</div>
            <div class="w-h1">صندوق الأقربين لعائلة العريفي</div>
            <div class="w-sub">نظام الإدارة المتكامل لإدارة الخطة التشغيلية والشؤون المالية ومتابعة البرامج والمشاريع لعام 2026م</div>
            <div class="w-greet" id="w-greet">مرحباً بك 👋</div>
          </div>
          <div class="w-stats">
            <div class="w-stat" style="border-color:var(--gold)"><div class="w-stat-v" style="color:var(--gold)" id="w-progs">0</div><div class="w-stat-l">برنامج ومشروع</div></div>
            <div class="w-stat" style="border-color:var(--teal)"><div class="w-stat-v" style="color:var(--teal)" id="w-pct">0%</div><div class="w-stat-l">نسبة الإنجاز الكلي</div></div>
            <div class="w-stat" style="border-color:var(--blue)"><div class="w-stat-v" style="color:var(--blue)" id="w-revenue">0</div><div class="w-stat-l">إجمالي الإيرادات (ر.س)</div></div>
            <div class="w-stat" style="border-color:var(--green)"><div class="w-stat-v" style="color:var(--green)" id="w-surplus">0</div><div class="w-stat-l">الفائض المالي (ر.س)</div></div>
          </div>
          <div class="w-cards" id="w-cards"></div>
        </div>
      </div>

      <!-- ══════ OVERVIEW ══════ -->
      <div id="tab-overview" class="sec">
        <div class="pg-title">📊 نظرة عامة على الخطة التشغيلية</div>
        <div class="pg-sub">ملخص تقدم الأهداف والبرامج لعام 2026</div>
        <div class="ov-grid" id="ov-grid"></div>
        <div class="card"><div class="card-t">📊 تقدم جميع البرامج</div><div id="all-bars"></div></div>
      </div>

      <!-- ══════ GOAL TABS ══════ -->
      <div id="tab-goal1" class="sec"><div class="pg-title" id="g1-title">🤝 الهدف الأول</div><div class="pg-sub" id="g1-sub"></div><div class="prog-list" id="list1"></div></div>
      <div id="tab-goal2" class="sec"><div class="pg-title" id="g2-title">🌐 الهدف الثاني</div><div class="pg-sub" id="g2-sub"></div><div class="prog-list" id="list2"></div></div>
      <div id="tab-goal3" class="sec"><div class="pg-title" id="g3-title">📚 الهدف الثالث</div><div class="pg-sub" id="g3-sub"></div><div class="prog-list" id="list3"></div></div>

      <!-- ══════ MY TASKS ══════ -->
      <div id="tab-mytasks" class="sec">
        <div class="pg-title">✅ برامجي المسندة</div>
        <div class="pg-sub">البرامج التي أُسندت إليك — يمكنك تحديث نسب الإنجاز والملاحظات</div>
        <div class="prog-list" id="list-mytasks"></div>
        <div id="mytasks-empty" style="display:none;text-align:center;padding:50px;color:var(--text-muted)">لا توجد برامج مسندة إليك حالياً. تواصل مع المدير لإسناد المهام.</div>
      </div>

      <!-- ══════ KPIS ══════ -->
      <div id="tab-kpis" class="sec">
        <div class="pg-title">📈 ملخص مؤشرات الأداء</div>
        <div class="pg-sub">جدول تفصيلي لمؤشرات الأداء لكل برنامج</div>
        <div id="kpi-container"></div>
      </div>

      <!-- ══════ MANAGE (ADMIN) ══════ -->
      <div id="tab-manage" class="sec">
        <div style="display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:10px;margin-bottom:18px">
          <div><div class="pg-title">⚙️ إدارة البرامج</div><div class="pg-sub" style="margin-bottom:0">إضافة وتعديل وحذف البرامج وتعديل أسماء الأهداف</div></div>
          <button class="btn btn-gold" onclick="openProgModal()">➕ إضافة برنامج</button>
        </div>
        <div id="manage-lock" class="lock-banner" style="display:none">🔒 هذا القسم للمدير فقط</div>
        <div id="manage-content">
          <div class="card" style="margin-bottom:16px">
            <div class="card-t">✏️ أسماء الأهداف التشغيلية</div>
            <div style="display:grid;gap:12px">
              <div class="field"><label>الهدف الأول</label><input class="ctrl" id="gn-1"></div>
              <div class="field"><label>الهدف الثاني</label><input class="ctrl" id="gn-2"></div>
              <div class="field"><label>الهدف الثالث</label><input class="ctrl" id="gn-3"></div>
              <button class="btn btn-teal" onclick="saveGoalNames()" style="justify-self:start">💾 حفظ أسماء الأهداف</button>
            </div>
          </div>
          <div class="card">
            <div class="card-t">📋 جميع البرامج</div>
            <div class="tbl-wrap"><table class="tbl"><thead><tr><th>البرنامج</th><th>الهدف</th><th>الميزانية</th><th>المسؤول</th><th>الإنجاز</th><th>إجراء</th></tr></thead><tbody id="manage-tbody"></tbody></table></div>
          </div>
        </div>
      </div>

      <!-- ══════ TEAM (ADMIN) ══════ -->
      <div id="tab-team" class="sec">
        <div class="pg-title">👥 الموظفون وإسناد البرامج</div>
        <div class="pg-sub">إضافة الموظفين وإسناد البرامج لهم ومتابعة نسبة إنجاز كل موظف</div>
        <div id="team-lock" class="lock-banner" style="display:none">🔒 هذا القسم للمدير فقط</div>
        <div id="team-content">
          <div style="display:grid;grid-template-columns:1fr 1.4fr;gap:16px;margin-bottom:16px">
            <div class="card">
              <div class="card-t">➕ إضافة موظف</div>
              <div style="display:grid;gap:10px">
                <div class="field"><label>اسم الموظف</label><input class="ctrl" id="emp-name" placeholder="مثال: صالح العريفي"></div>
                <div class="field"><label>الجهة / المنصب</label><input class="ctrl" id="emp-role" placeholder="مثال: إدارة البرامج"></div>
                <div class="field"><label>اسم مستخدم للدخول (اختياري)</label><input class="ctrl" id="emp-user" placeholder="مثال: saleh"></div>
                <button class="btn btn-gold" onclick="addEmp()">➕ إضافة الموظف</button>
              </div>
            </div>
            <div class="card">
              <div class="card-t">🔗 إسناد برنامج لموظف</div>
              <div style="display:grid;gap:10px">
                <div class="field"><label>الموظف</label><select class="ctrl" id="assign-emp"></select></div>
                <div class="field"><label>البرنامج</label><select class="ctrl" id="assign-prog"></select></div>
                <button class="btn btn-blue" onclick="assignProg()">🔗 إسناد البرنامج</button>
                <div style="font-size:.74rem;color:var(--text-muted);line-height:1.6">عند الإسناد، يستطيع الموظف الدخول على برنامجه من قسم «برامجي المسندة» وتحديث نسبة الإنجاز، وتنعكس مباشرة على المؤشرات.</div>
              </div>
            </div>
          </div>
          <div class="card">
            <div class="card-t">👥 قائمة الموظفين والبرامج المسندة</div>
            <div id="team-list"></div>
          </div>
        </div>
      </div>

      <!-- ══════ FINANCE (ACCOUNTANT) ══════ -->
      <div id="tab-finance" class="sec">
        <div style="display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:12px;margin-bottom:18px">
          <div><div class="pg-title">💰 لوحة القيادة المالية</div><div class="pg-sub" style="margin-bottom:0" id="fin-role">—</div></div>
          <div style="display:flex;gap:8px;flex-wrap:wrap">
            <button class="btn btn-teal" id="fin-add-rev" onclick="openFinModal('rev')" style="display:none">➕ إيراد</button>
            <button class="btn btn-blue" id="fin-add-exp" onclick="openFinModal('exp')" style="display:none">➕ مصروف</button>
            <button class="btn btn-gold" onclick="finExport()">📥 تصدير Excel</button>
            <button class="btn btn-violet" onclick="persist()">💾 حفظ</button>
          </div>
        </div>
        <div class="lock-banner" id="fin-lock" style="display:none">🔒 وضع العرض فقط — صلاحية التعديل المالي للمحاسب فقط</div>

        <!-- KPI strip 1 -->
        <div class="kpis k4">
          <div class="kpi" style="border-color:var(--green)"><div class="kpi-wm">📥</div><div class="kpi-v" style="color:var(--green)" id="f-revenue">—</div><div class="kpi-l">إجمالي الإيرادات (ر.س)</div><div class="kpi-bar"><div class="kpi-bar-f" style="background:var(--green);width:100%"></div></div></div>
          <div class="kpi" style="border-color:var(--rose)"><div class="kpi-wm">📤</div><div class="kpi-v" style="color:var(--rose)" id="f-expense">—</div><div class="kpi-l">إجمالي المصروفات (ر.س)</div><div class="kpi-bar"><div class="kpi-bar-f" id="fbar-exp" style="background:var(--rose);width:0%"></div></div></div>
          <div class="kpi" style="border-color:var(--gold)"><div class="kpi-wm">💎</div><div class="kpi-v" style="color:var(--gold)" id="f-surplus">—</div><div class="kpi-l">الفائض / العجز المالي (ر.س)</div><div class="kpi-bar"><div class="kpi-bar-f" id="fbar-sur" style="background:var(--gold);width:0%"></div></div></div>
          <div class="kpi" style="border-color:var(--blue)"><div class="kpi-wm">🏦</div><div class="kpi-v" style="color:var(--blue)" id="f-totbal">—</div><div class="kpi-l">إجمالي الأرصدة (ر.س)</div><div class="kpi-bar"><div class="kpi-bar-f" style="background:var(--blue);width:100%"></div></div></div>
        </div>
        <!-- KPI strip 2 -->
        <div class="kpis k4">
          <div class="kpi" style="border-color:var(--teal)"><div class="kpi-wm">💵</div><div class="kpi-v" style="color:var(--teal)" id="f-cashfree">—</div><div class="kpi-l">الرصيد النقدي المتاح (ر.س)</div><div class="kpi-bar"><div class="kpi-bar-f" style="background:var(--teal);width:100%"></div></div></div>
          <div class="kpi" style="border-color:var(--amber)"><div class="kpi-wm">🔒</div><div class="kpi-v" style="color:var(--amber)" id="f-cashrestr">—</div><div class="kpi-l">الرصيد النقدي المقيّد (ر.س)</div><div class="kpi-bar"><div class="kpi-bar-f" style="background:var(--amber);width:100%"></div></div></div>
          <div class="kpi" style="border-color:var(--violet)"><div class="kpi-wm">📋</div><div class="kpi-v" style="color:var(--violet)" id="f-liabilities">—</div><div class="kpi-l">إجمالي الالتزامات (ر.س)</div><div class="kpi-bar"><div class="kpi-bar-f" style="background:var(--violet);width:100%"></div></div></div>
          <div class="kpi" style="border-color:var(--rose)"><div class="kpi-wm">📊</div><div class="kpi-v" style="color:var(--rose)" id="f-spendrate">—</div><div class="kpi-l">نسبة الصرف من الإيرادات</div><div class="kpi-bar"><div class="kpi-bar-f" id="fbar-rate" style="background:var(--rose);width:0%"></div></div></div>
        </div>

        <!-- Charts -->
        <div class="ch-row">
          <div class="ch-box"><div class="ch-t">📊 الإيرادات مقابل المصروفات شهرياً</div><div class="ch-wrap"><canvas id="fc-monthly"></canvas></div></div>
          <div class="ch-box"><div class="ch-t">🥧 توزيع المصروفات حسب البند</div><div class="ch-wrap"><canvas id="fc-expcat"></canvas></div></div>
        </div>
        <div class="ch-3">
          <div class="ch-box"><div class="ch-t">💰 مصادر الإيرادات</div><div class="ch-wrap"><canvas id="fc-revsrc"></canvas></div></div>
          <div class="ch-box"><div class="ch-t">🏦 الأرصدة: متاح مقابل مقيّد</div><div class="ch-wrap"><canvas id="fc-balance"></canvas></div></div>
          <div class="ch-box"><div class="ch-t">🎯 الفائض التراكمي</div><div class="ch-wrap"><canvas id="fc-cumulative"></canvas></div></div>
        </div>

        <!-- Revenues table -->
        <div class="card" style="margin-bottom:18px">
          <div class="card-t">📥 الإيرادات</div>
          <div class="tbl-wrap"><table class="tbl"><thead><tr><th>#</th><th>المصدر</th><th>التصنيف</th><th>المبلغ (ر.س)</th><th>النوع</th><th>الشهر</th><th>ملاحظات</th><th class="fin-act" style="display:none">إجراء</th></tr></thead><tbody id="rev-tbody"></tbody></table></div>
          <div class="tbl-foot" id="rev-foot"></div>
        </div>

        <!-- Expenses table -->
        <div class="card" style="margin-bottom:18px">
          <div class="card-t">📤 المصروفات</div>
          <div class="tbl-wrap"><table class="tbl"><thead><tr><th>#</th><th>البند</th><th>التصنيف</th><th>المبلغ (ر.س)</th><th>الشهر</th><th>الحالة</th><th>ملاحظات</th><th class="fin-act" style="display:none">إجراء</th></tr></thead><tbody id="exp-tbody"></tbody></table></div>
          <div class="tbl-foot" id="exp-foot"></div>
        </div>

        <!-- Balances & Liabilities -->
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:16px" id="bal-liab-grid">
          <div class="card">
            <div class="card-t" style="justify-content:space-between"><span>🏦 الأرصدة البنكية والنقدية</span><button class="btn btn-teal fin-add-extra" onclick="openFinModal('bal')" style="display:none;padding:4px 12px;font-size:.74rem">➕</button></div>
            <div class="tbl-wrap"><table class="tbl"><thead><tr><th>الحساب / الرصيد</th><th>النوع</th><th>المبلغ</th><th class="fin-act" style="display:none"></th></tr></thead><tbody id="bal-tbody"></tbody></table></div>
          </div>
          <div class="card">
            <div class="card-t" style="justify-content:space-between"><span>📋 الالتزامات المالية</span><button class="btn btn-violet fin-add-extra" onclick="openFinModal('liab')" style="display:none;padding:4px 12px;font-size:.74rem">➕</button></div>
            <div class="tbl-wrap"><table class="tbl"><thead><tr><th>الالتزام</th><th>المستحق</th><th>المبلغ</th><th class="fin-act" style="display:none"></th></tr></thead><tbody id="liab-tbody"></tbody></table></div>
          </div>
        </div>
      </div>

      <!-- ══════ BYLAWS ══════ -->
      <div id="tab-bylaws" class="sec">
        <div style="display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:10px;margin-bottom:18px">
          <div><div class="pg-title">📜 اللائحة الأساسية لصندوق الأقربين</div><div class="pg-sub" style="margin-bottom:0">مُرخّص برقم (7145) — معتمدة من المركز الوطني لتنمية القطاع غير الربحي</div></div>
          <a class="btn btn-gold" href="اللائحة_الأساسية_صندوق_الأقربين.pdf" target="_blank" download>📥 تحميل اللائحة (PDF)</a>
        </div>
        <div class="bylaws-doc" id="bylaws-doc"></div>
      </div>

    </div><!-- /main -->
  </div><!-- /content -->
</div><!-- /shell -->

<!-- ═══════ PROGRAM MODAL ═══════ -->
<div class="modal-bg" id="prog-modal">
  <div class="modal">
    <div class="modal-t" id="prog-modal-t">➕ إضافة برنامج جديد</div>
    <input type="hidden" id="pm-id">
    <div class="modal-grid">
      <div class="field"><label>اسم البرنامج</label><input class="ctrl" id="pm-name" placeholder="مثال: برنامج رعاية"></div>
      <div class="field"><label>الهدف التشغيلي</label><select class="ctrl" id="pm-goal"><option value="1">الهدف الأول: التكافل</option><option value="2">الهدف الثاني: التواصل</option><option value="3">الهدف الثالث: التطوير</option></select></div>
      <div class="field"><label>الميزانية (ر.س)</label><input class="ctrl" type="number" id="pm-budget" placeholder="0"></div>
      <div class="field"><label>جهة التنفيذ</label><input class="ctrl" id="pm-exec" placeholder="إدارة البرامج"></div>
      <div class="field"><label>الوقت الزمني</label><input class="ctrl" id="pm-time" placeholder="طوال العام"></div>
      <div class="field"><label>مؤشر الأداء KPI</label><input class="ctrl" id="pm-kpi" placeholder="عدد المستفيدين"></div>
      <div class="field"><label>الهدف السنوي الكمّي</label><input class="ctrl" type="number" id="pm-target" placeholder="12"></div>
      <div class="field"><label>الموظف المسؤول</label><select class="ctrl" id="pm-assignee"><option value="">— بدون —</option></select></div>
    </div>
    <div class="field" style="margin-bottom:14px"><label>وصف البرنامج</label><textarea class="ctrl" id="pm-desc" rows="2" placeholder="وصف البرنامج..."></textarea></div>
    <div style="display:grid;grid-template-columns:repeat(4,1fr);gap:8px;margin-bottom:16px">
      <div class="field"><label>ر١ %</label><input class="ctrl" type="number" id="pm-q1" value="25"></div>
      <div class="field"><label>ر٢ %</label><input class="ctrl" type="number" id="pm-q2" value="25"></div>
      <div class="field"><label>ر٣ %</label><input class="ctrl" type="number" id="pm-q3" value="25"></div>
      <div class="field"><label>ر٤ %</label><input class="ctrl" type="number" id="pm-q4" value="25"></div>
    </div>
    <div class="modal-act"><button class="btn btn-gold" style="flex:1" onclick="saveProg()">💾 حفظ البرنامج</button><button class="btn btn-rose" onclick="closeProgModal()">إلغاء</button></div>
  </div>
</div>

<!-- ═══════ FINANCE MODAL ═══════ -->
<div class="modal-bg" id="fin-modal">
  <div class="modal">
    <div class="modal-t" id="fin-modal-t">➕ إضافة</div>
    <input type="hidden" id="fm-kind">
    <div id="fm-fields"></div>
    <div class="modal-act"><button class="btn btn-gold" style="flex:1" onclick="saveFinItem()">💾 حفظ</button><button class="btn btn-rose" onclick="closeFinModal()">إلغاء</button></div>
  </div>
</div>

<script>
/* ════════════════════════════════════════════
   USERS & ROLES
   ════════════════════════════════════════════ */
const USERS=[
  {u:'admin',p:'admin2026',name:'المدير التنفيذي',role:'admin'},
  {u:'hisab',p:'hisab2026',name:'المحاسب المالي',role:'accountant'},
  {u:'saleh',p:'saleh2026',name:'صالح العريفي',role:'employee',empId:'emp_saleh'},
];
let CU=null;

/* ════════════════════════════════════════════
   STATE / STORAGE
   ════════════════════════════════════════════ */
let ST={};
function loadST(){try{const s=localStorage.getItem('aqrabin_v1');if(s)ST=JSON.parse(s);}catch(e){ST={};}}
function persist(){localStorage.setItem('aqrabin_v1',JSON.stringify(ST));toast('✅ تم حفظ جميع البيانات',false);}
let toastTmr;
function toast(msg,save=true){document.getElementById('toast-msg').textContent=msg;document.getElementById('toast-save').style.display=save?'inline-block':'none';document.getElementById('toast').classList.add('show');clearTimeout(toastTmr);toastTmr=setTimeout(()=>document.getElementById('toast').classList.remove('show'),3500);}
function fmt(n){return Math.round(n||0).toLocaleString('en-US');}

/* ════════════════════════════════════════════
   DEFAULT DATA
   ════════════════════════════════════════════ */
let GOAL_NAMES=['تعزيز التكافل العائلي','تحقيق التواصل العائلي','تطوير وتدريب أفراد العائلة'];

const DEFAULT_PROGRAMS=[
  // GOAL 1 - التكافل
  {id:'p_riaaya',goal:1,name:'رعاية',color:'#9c6817',budget:600000,exec:'إدارة البرامج',time:'طوال العام',kpi:'عدد المستفيدين',qPct:[25,25,25,25],target:120,assignee:'',desc:'برنامج رعاية الأسر المحتاجة من العائلة',subs:[{desc:'رعاية شهرية للأسر المحتاجة',target:'120 أسرة',kpi:'عدد المستفيدين',budget:600000,q:'طوال العام'}]},
  {id:'p_zawaj',goal:1,name:'إعانة الزواج',color:'#c4881e',budget:200000,exec:'إدارة البرامج',time:'طوال العام',kpi:'عدد المستفيدين',qPct:[25,25,25,25],target:20,assignee:'',desc:'دعم الشباب المقبلين على الزواج',subs:[{desc:'إعانة زواج للشباب',target:'20 مستفيد',kpi:'عدد المستفيدين',budget:200000,q:'طوال العام'}]},
  {id:'p_sehiya',goal:1,name:'الطوارئ الصحية',color:'#b54a3a',budget:150000,exec:'إدارة البرامج',time:'طوال العام',kpi:'عدد الحالات',qPct:[25,25,25,25],target:30,assignee:'',desc:'تغطية الطوارئ الصحية لأفراد العائلة',subs:[{desc:'دعم الحالات الصحية الطارئة',target:'30 حالة',kpi:'عدد الحالات',budget:150000,q:'طوال العام'}]},
  // GOAL 2 - التواصل
  {id:'p_liqaa',goal:2,name:'لقاءات العائلة',color:'#3a6ea5',budget:114000,exec:'إدارة البرامج',time:'الأرباع 1-3',kpi:'عدد الحاضرين',qPct:[13,68,19,0],target:8,assignee:'',desc:'لقاءات عائلية سنوية وإقليمية',subs:[{desc:'اللقاء العائلي السنوي + لقاءات إقليمية',target:'8 لقاءات',kpi:'عدد اللقاءات',budget:114000,q:'الأرباع 1-3'}]},
  {id:'p_ihsaa',goal:2,name:'إحصاء العائلة',color:'#2f7d6b',budget:22000,exec:'إدارة التقنية',time:'طوال العام',kpi:'نسبة المسجلين',qPct:[55,18,14,13],target:100,assignee:'',desc:'حصر أفراد العائلة وتطبيق الأقربين',subs:[{desc:'حصر الأفراد + إطلاق التطبيق',target:'إكمال المراحل',kpi:'نسبة المسجلين',budget:22000,q:'طوال العام'}]},
  {id:'p_eelam',goal:2,name:'إعلام',color:'#7a5ba0',budget:15000,exec:'إدارة الاتصال',time:'طوال العام',kpi:'عدد المحتوى',qPct:[25,25,25,25],target:48,assignee:'',desc:'إدارة قنوات التواصل ونشر المحتوى',subs:[{desc:'محتوى إعلامي منتظم',target:'محتوى أسبوعي',kpi:'عدد المنشورات',budget:15000,q:'طوال العام'}]},
  // GOAL 3 - التطوير
  {id:'p_taheel',goal:3,name:'تأهيل',color:'#7a5ba0',budget:10000,exec:'إدارة البرامج',time:'الربع الثالث',kpi:'عدد المشاركين',qPct:[0,0,100,0],target:25,assignee:'',desc:'تأهيل الباحثين عن عمل',subs:[{desc:'دورة تأهيلية للباحثين عن عمل',target:'25 مشارك',kpi:'عدد المشاركين',budget:10000,q:'الربع الثالث'}]},
  {id:'p_mustashar',goal:3,name:'مستشارك',color:'#3f8a4a',budget:10000,exec:'مختصون من العائلة',time:'من الربع الثاني',kpi:'عدد الاستشارات',qPct:[0,33,34,33],target:150,assignee:'',desc:'استشارات من مختصي العائلة',subs:[{desc:'استشارات أسبوعية متخصصة',target:'150 استشارة',kpi:'عدد الاستشارات',budget:10000,q:'من الربع الثاني'}]},
  {id:'p_mahara',goal:3,name:'مهارة',color:'#c4881e',budget:10000,exec:'إدارة البرامج',time:'الأرباع 2-3',kpi:'عدد المشاركين',qPct:[0,50,50,0],target:60,assignee:'',desc:'برامج تنمية المهارات',subs:[{desc:'برامج تطوير المهارات الحياتية',target:'60 مشارك',kpi:'عدد المشاركين',budget:10000,q:'الأرباع 2-3'}]},
];

let DEFAULT_EMPLOYEES=[
  {id:'emp_saleh',name:'صالح العريفي',role:'إدارة البرامج',user:'saleh'},
  {id:'emp_fahd',name:'فهد العريفي',role:'إدارة التقنية والاتصال',user:''},
  {id:'emp_nasser',name:'ناصر العريفي',role:'إدارة الشراكات',user:''},
];

/* Finance defaults */
const DEFAULT_REVENUES=[
  {id:'r1',source:'تبرعات أفراد العائلة',cat:'تبرعات',amount:850000,type:'متاح',month:'يناير',notes:''},
  {id:'r2',source:'اشتراكات الأعضاء (600 ر/عضو)',cat:'اشتراكات',amount:120000,type:'متاح',month:'يناير',notes:'200 عضو مشترك'},
  {id:'r3',source:'أوقاف الصندوق',cat:'أوقاف',amount:400000,type:'مقيّد',month:'فبراير',notes:'ريع الأوقاف'},
  {id:'r4',source:'أموال زكاة',cat:'زكوات',amount:300000,type:'مقيّد',month:'مارس',notes:'تصرف في مصارفها الشرعية'},
  {id:'r5',source:'عائدات استثمارات',cat:'استثمارات',amount:180000,type:'متاح',month:'أبريل',notes:''},
  {id:'r6',source:'هبات ووصايا',cat:'هبات',amount:90000,type:'متاح',month:'مايو',notes:''},
];
const DEFAULT_EXPENSES=[
  {id:'e1',item:'برامج التكافل الاجتماعي',cat:'برامج',amount:480000,month:'مارس',status:'قيد التنفيذ',notes:''},
  {id:'e2',item:'صرف زكاة في مصارفها',cat:'زكوات',amount:280000,month:'مارس',status:'مكتمل',notes:'من الرصيد المقيّد'},
  {id:'e3',item:'رواتب الجهاز التنفيذي',cat:'رواتب',amount:180000,month:'أبريل',status:'قيد التنفيذ',notes:''},
  {id:'e4',item:'لقاءات وفعاليات العائلة',cat:'برامج',amount:95000,month:'مايو',status:'قيد التنفيذ',notes:''},
  {id:'e5',item:'مصاريف تشغيلية وإدارية',cat:'إداري',amount:60000,month:'فبراير',status:'مكتمل',notes:''},
  {id:'e6',item:'تقنية وتطبيق الأقربين',cat:'تقنية',amount:48000,month:'يناير',status:'مكتمل',notes:''},
];
const DEFAULT_BALANCES=[
  {id:'b1',name:'الحساب الجاري — مصرف الراجحي',type:'متاح',amount:620000},
  {id:'b2',name:'حساب الزكاة المستقل',type:'مقيّد',amount:120000},
  {id:'b3',name:'حساب الأوقاف',type:'مقيّد',amount:340000},
  {id:'b4',name:'صندوق نقدي',type:'متاح',amount:45000},
];
const DEFAULT_LIABILITIES=[
  {id:'l1',name:'التزامات برامج معتمدة لم تُصرف',due:'الربع الثاني',amount:150000},
  {id:'l2',name:'مستحقات موردين',due:'الربع الأول',amount:35000},
];

/* Working data */
let PROGRAMS=[], EMPLOYEES=[], REVENUES=[], EXPENSES=[], BALANCES=[], LIABILITIES=[];

function initData(){
  loadST();
  PROGRAMS = ST.programs ? ST.programs : JSON.parse(JSON.stringify(DEFAULT_PROGRAMS));
  EMPLOYEES = ST.employees ? ST.employees : JSON.parse(JSON.stringify(DEFAULT_EMPLOYEES));
  REVENUES = ST.revenues ? ST.revenues : JSON.parse(JSON.stringify(DEFAULT_REVENUES));
  EXPENSES = ST.expenses ? ST.expenses : JSON.parse(JSON.stringify(DEFAULT_EXPENSES));
  BALANCES = ST.balances ? ST.balances : JSON.parse(JSON.stringify(DEFAULT_BALANCES));
  LIABILITIES = ST.liabilities ? ST.liabilities : JSON.parse(JSON.stringify(DEFAULT_LIABILITIES));
  if(ST.goalNames) GOAL_NAMES = ST.goalNames;
  if(ST.progState===undefined) ST.progState={};
}
function syncST(){ST.programs=PROGRAMS;ST.employees=EMPLOYEES;ST.revenues=REVENUES;ST.expenses=EXPENSES;ST.balances=BALANCES;ST.liabilities=LIABILITIES;ST.goalNames=GOAL_NAMES;}

/* program achievement state helpers */
function pg(id,k,d=''){const key=`${id}_${k}`;return ST.progState&&ST.progState[key]!==undefined?ST.progState[key]:d;}
function psv(id,k,v){if(!ST.progState)ST.progState={};ST.progState[`${id}_${k}`]=v;toast('💾 تعديل — اضغط حفظ',true);recalcAndRender();}

/* ════════════════════════════════════════════
   AUTH
   ════════════════════════════════════════════ */
function doLogin(){
  const u=document.getElementById('lu').value.trim(),p=document.getElementById('lp').value;
  const f=USERS.find(x=>x.u===u&&x.p===p);
  if(!f){document.getElementById('lerr').textContent='❌ بيانات الدخول غير صحيحة';document.getElementById('lp').value='';return;}
  CU=f;
  document.getElementById('lerr').textContent='';
  document.getElementById('login').classList.add('gone');
  document.getElementById('shell').classList.remove('hidden');
  initData();
  applyRoleVisibility();
  buildBylaws();
  recalcAndRender();
  // user chip
  document.getElementById('sb-uname').textContent=f.name;
  const roleLabel={admin:'المدير التنفيذي',accountant:'المحاسب المالي',employee:'موظف مسؤول'};
  document.getElementById('sb-urole').textContent=roleLabel[f.role];
  document.getElementById('sb-av').textContent=f.role==='accountant'?'🧮':f.role==='admin'?'👔':'👤';
  document.getElementById('w-greet').textContent=`مرحباً ${f.name} 👋`;
  buildWelcomeCards();
  goTab('home');
}
function doLogout(){if(!confirm('تسجيل الخروج؟'))return;CU=null;document.getElementById('lu').value='';document.getElementById('lp').value='';document.getElementById('login').classList.remove('gone');document.getElementById('shell').classList.add('hidden');}

function isAdmin(){return CU&&CU.role==='admin';}
function isAccountant(){return CU&&CU.role==='accountant';}
function isEmployee(){return CU&&CU.role==='employee';}
function canEditPlan(){return isAdmin();}

/* Show/hide sidebar groups based on role */
function applyRoleVisibility(){
  const show=(sel,on)=>document.querySelectorAll(sel).forEach(e=>e.style.display=on?'':'none');
  // groups
  const planOn = isAdmin()||isEmployee()||isAccountant(); // everyone sees plan summary
  // admin sees everything; accountant sees finance + plan(view); employee sees mytasks + plan(view)
  show('[data-grp="admin"]', isAdmin());
  show('[data-grp="fin"]', isAdmin()||isAccountant());
  show('[data-grp="mytasks"]', isEmployee());
  // finance edit controls
  const finEdit=isAccountant();
  document.getElementById('fin-add-rev').style.display=finEdit?'inline-flex':'none';
  document.getElementById('fin-add-exp').style.display=finEdit?'inline-flex':'none';
  document.querySelectorAll('.fin-add-extra').forEach(e=>e.style.display=finEdit?'inline-flex':'none');
  document.querySelectorAll('.fin-act').forEach(e=>e.style.display=finEdit?'':'none');
  document.getElementById('fin-lock').style.display=finEdit?'none':'flex';
  document.getElementById('fin-role').textContent=finEdit?('🧮 '+CU.name+' — صلاحية تعديل كاملة'):('👁️ '+CU.name+' — عرض فقط');
  // manage/team lock (only admin reaches them anyway)
}

/* ════════════════════════════════════════════
   NAVIGATION
   ════════════════════════════════════════════ */
const TAB_TITLES={home:'🏠 الرئيسية',overview:'📊 نظرة عامة',goal1:'🤝 الهدف الأول',goal2:'🌐 الهدف الثاني',goal3:'📚 الهدف الثالث',kpis:'📈 ملخص الأداء',mytasks:'✅ برامجي المسندة',manage:'⚙️ إدارة البرامج',team:'👥 الموظفون والإسناد',finance:'💰 لوحة المحاسب المالية',bylaws:'📜 اللائحة الأساسية'};
function showTab(name,btn){
  document.querySelectorAll('.sec').forEach(s=>s.classList.remove('active'));
  document.querySelectorAll('.sb-btn').forEach(b=>b.classList.remove('active'));
  const tab=document.getElementById('tab-'+name);if(tab)tab.classList.add('active');
  if(btn)btn.classList.add('active'); else {const b=document.querySelector(`.sb-btn[data-tab="${name}"]`);if(b)b.classList.add('active');}
  const t=document.getElementById('topbar-title');if(t)t.textContent=TAB_TITLES[name]||'';
  if(window.innerWidth<=860){document.getElementById('sidebar').classList.remove('open');document.getElementById('sb-ov').classList.remove('show');}
  if(name==='kpis')renderKPIs();
  if(name==='finance')renderFinance();
  if(name==='manage')renderManage();
  if(name==='team')renderTeam();
  if(name==='mytasks')renderMyTasks();
  if(name==='overview')renderOverview();
  window.scrollTo({top:0,behavior:'smooth'});
}
function goTab(name){const b=document.querySelector(`.sb-btn[data-tab="${name}"]`);showTab(name,b);}
function toggleSB(){document.getElementById('sidebar').classList.toggle('open');document.getElementById('sb-ov').classList.toggle('show');}

function buildWelcomeCards(){
  const cards=[];
  cards.push({ic:'📊',t:'الخطة التشغيلية',d:'متابعة تقدم البرامج وإنجاز الأهداف على مدار العام.',tab:'overview'});
  if(isEmployee()) cards.push({ic:'✅',t:'برامجي المسندة',d:'البرامج المسندة إليك لتحديث نسب الإنجاز.',tab:'mytasks'});
  if(isAdmin()){cards.push({ic:'⚙️',t:'إدارة البرامج',d:'إضافة وتعديل البرامج وإسنادها للموظفين.',tab:'manage'});}
  if(isAdmin()||isAccountant()) cards.push({ic:'💰',t:'لوحة المحاسب المالية',d:'الإيرادات والمصروفات والأرصدة والفائض المالي.',tab:'finance'});
  cards.push({ic:'📜',t:'اللائحة الأساسية',d:'الاطلاع على اللائحة المعتمدة وتحميلها.',tab:'bylaws'});
  document.getElementById('w-cards').innerHTML=cards.slice(0,6).map(c=>`<div class="w-card" onclick="goTab('${c.tab}')"><div class="w-card-ic">${c.ic}</div><div class="w-card-t">${c.t}</div><div class="w-card-d">${c.d}</div><div class="w-card-a">دخول ←</div></div>`).join('');
}

/* ════════════════════════════════════════════
   PROGRAM CALCULATIONS
   ════════════════════════════════════════════ */
function annualTarget(id){const p=PROGRAMS.find(x=>x.id===id);const v=pg(id,'annual_target',p?p.target:0);return parseFloat(v)||0;}
function qTarget(id,qi){const a=annualTarget(id);const p=PROGRAMS.find(x=>x.id===id);if(!p||!a)return 0;const pct=parseFloat(pg(id,`q${qi}_pct`,p.qPct[qi-1]))||0;return Math.round(a*pct/100);}
function qAch(id,qi){return parseFloat(pg(id,`q${qi}_ach`,0))||0;}
function totalAch(id){return [1,2,3,4].reduce((s,q)=>s+qAch(id,q),0);}
function progPct(id){const t=annualTarget(id);if(!t)return 0;return Math.min(100,Math.round(totalAch(id)/t*100));}
function goalPct(gn){const pp=PROGRAMS.filter(x=>x.goal===gn);if(!pp.length)return 0;return Math.round(pp.reduce((s,p)=>s+progPct(p.id),0)/pp.length);}
function totalPct(){const g=[1,2,3].filter(n=>PROGRAMS.some(p=>p.goal===n));if(!g.length)return 0;return Math.round(g.reduce((s,n)=>s+goalPct(n),0)/g.length);}
function pCol(p){return p>=80?'var(--green)':p>=50?'var(--amber)':p>0?'var(--rose)':'var(--text-muted)';}
function stLabel(p){if(p>=80)return{t:'🏆 ممتاز',bg:'var(--green-p)',c:'var(--green)'};if(p>=50)return{t:'✅ في المسار',bg:'var(--amber-p)',c:'var(--amber)'};if(p>0)return{t:'⚠️ يحتاج متابعة',bg:'var(--rose-p)',c:'var(--rose)'};return{t:'🔴 لم يبدأ',bg:'rgba(156,104,23,.06)',c:'var(--text-muted)'};}

function empName(id){const e=EMPLOYEES.find(x=>x.id===id);return e?e.name:'';}

/* ════════════════════════════════════════════
   RENDER MASTER
   ════════════════════════════════════════════ */
function recalcAndRender(){
  syncST();
  applyGoalNames();
  [1,2,3].forEach(renderGoal);
  renderOverview();
  renderWelcomeStats();
  updateChips();
}
function updateChips(){
  document.getElementById('chip-g1').textContent=PROGRAMS.filter(p=>p.goal===1).length;
  document.getElementById('chip-g2').textContent=PROGRAMS.filter(p=>p.goal===2).length;
  document.getElementById('chip-g3').textContent=PROGRAMS.filter(p=>p.goal===3).length;
  if(isEmployee()){const mine=PROGRAMS.filter(p=>p.assignee===CU.empId);document.getElementById('chip-mytasks').textContent=mine.length;}
}
function applyGoalNames(){
  document.getElementById('g1-title').textContent='🤝 الهدف الأول: '+GOAL_NAMES[0];
  document.getElementById('g2-title').textContent='🌐 الهدف الثاني: '+GOAL_NAMES[1];
  document.getElementById('g3-title').textContent='📚 الهدف الثالث: '+GOAL_NAMES[2];
  ['1','2','3'].forEach(i=>{const e=document.getElementById('gn-'+i);if(e&&!e.value)e.value=GOAL_NAMES[i-1];});
}

function renderWelcomeStats(){
  document.getElementById('w-progs').textContent=PROGRAMS.length;
  document.getElementById('w-pct').textContent=totalPct()+'%';
  document.getElementById('w-revenue').textContent=fmt(totRevenue());
  document.getElementById('w-surplus').textContent=fmt(surplus());
}

/* ════════════════════════════════════════════
   OVERVIEW
   ════════════════════════════════════════════ */
function renderOverview(){
  const goals=[{n:1,ic:'🤝',c:'var(--teal)'},{n:2,ic:'🌐',c:'var(--blue)'},{n:3,ic:'📚',c:'var(--violet)'}];
  let html=goals.map(g=>{const pp=PROGRAMS.filter(p=>p.goal===g.n);const b=pp.reduce((s,p)=>s+(+p.budget||0),0);return`<div class="ov-card" style="border-color:${g.c}"><div class="ov-num" style="color:${g.c}">${goalPct(g.n)}%</div><div class="ov-lbl">${GOAL_NAMES[g.n-1]}</div><div class="ov-sub">${pp.length} برامج · ${fmt(b)} ر.س</div></div>`;}).join('');
  const totB=PROGRAMS.reduce((s,p)=>s+(+p.budget||0),0);
  html+=`<div class="ov-card" style="border-color:var(--gold)"><div class="ov-num" style="color:var(--gold)">${totalPct()}%</div><div class="ov-lbl">إجمالي الإنجاز</div><div class="ov-sub">ميزانية البرامج: ${fmt(totB)} ر.س</div></div>`;
  document.getElementById('ov-grid').innerHTML=html;
  document.getElementById('all-bars').innerHTML=PROGRAMS.map(p=>{const pct=progPct(p.id);const col=pCol(pct);return`<div style="display:flex;align-items:center;gap:10px;margin-bottom:9px"><div style="width:150px;flex-shrink:0;font-size:.75rem;font-weight:600;color:var(--text-dim)">برنامج ${p.name}</div><div style="flex:1;background:rgba(156,104,23,.08);border-radius:20px;height:7px;overflow:hidden"><div style="height:100%;width:${pct}%;background:${col};border-radius:20px;transition:width .8s"></div></div><div style="width:36px;text-align:left;font-size:.75rem;font-weight:800;color:${col}">${pct}%</div><div style="width:90px;text-align:left;font-size:.68rem;color:var(--text-muted)">${fmt(p.budget)}</div></div>`;}).join('');
}

/* ════════════════════════════════════════════
   PROGRAM CARDS
   ════════════════════════════════════════════ */
function buildCard(p,editable){
  const pct=progPct(p.id),col=pCol(pct),sl=stLabel(pct);
  const qLabels=['الربع الأول','الربع الثاني','الربع الثالث','الربع الرابع'];
  const qHTML=[0,1,2,3].map(i=>{const qi=i+1;const pctVal=parseFloat(pg(p.id,`q${qi}_pct`,p.qPct[i]))||0;const ach=pg(p.id,`q${qi}_ach`,'');const tgt=qTarget(p.id,qi);const qp=tgt>0?Math.min(100,Math.round((parseFloat(ach)||0)/tgt*100)):0;const qcol=['var(--q1c)','var(--q2c)','var(--q3c)','var(--q4c)'][i];
    return`<div class="qsp q${qi}"><div class="qsp-l">${qLabels[i]} <span>${qp}%</span></div>
      <input class="qsp-in" type="number" value="${pctVal}" ${editable?`onchange="psv('${p.id}','q${qi}_pct',this.value)"`:'disabled'} title="نسبة المستهدف %">
      <div class="qsp-cap">المستهدف: ${tgt||'—'}</div>
      <input class="qsp-ach" type="number" placeholder="المحقق" value="${ach}" ${editable?`onchange="psv('${p.id}','q${qi}_ach',this.value)"`:'disabled'}>
      <div class="qsp-cap">تم تحقيقه ✅</div>
      <div class="qbar"><div class="qbar-f" style="width:${qp}%;background:${qcol}"></div></div></div>`;}).join('');
  const subs=p.subs.map((s,i)=>`<div class="sub-item"><div class="sub-dot" style="background:${p.color}1a;color:${p.color}">${i+1}</div><div style="flex:1"><div class="sub-desc">${s.desc}</div><div class="sub-chips"><span class="chip chip-dim">🎯 ${s.target}</span><span class="chip chip-teal">📊 ${s.kpi}</span>${s.budget?`<span class="chip chip-gold">💰 ${fmt(s.budget)}</span>`:''}</div></div></div>`).join('');
  const assigneeTag=p.assignee?`<span class="assignee-tag">👤 ${empName(p.assignee)}</span>`:'<span style="font-size:.72rem;color:var(--text-muted)">غير مسند</span>';
  return`<div class="pc" id="pc-${p.id}">
    <div class="pc-head" onclick="togglePC('${p.id}')">
      <div class="pc-tag" style="background:${p.color}1a;color:${p.color}">${p.name}</div>
      <div class="pc-title">${p.subs[0].desc.substring(0,55)}${p.subs[0].desc.length>55?'…':''}</div>
      <div class="pc-meta"><div class="pill pill-gold">💰 ${fmt(p.budget)}</div><div class="pill pill-dim">🕐 ${p.time}</div><div style="font-size:.85rem;font-weight:900;font-family:'Cairo',sans-serif;min-width:38px;text-align:center;color:${col}">${pct}%</div><div class="pc-chev">▼</div></div>
    </div>
    <div class="pc-body"><div class="pc-grid">
      <div>
        <div class="meta-box">
          <div class="meta-r"><div class="meta-k">جهة التنفيذ</div><div class="meta-v">${p.exec}</div></div>
          <div class="meta-r"><div class="meta-k">الموظف المسؤول</div><div class="meta-v">${assigneeTag}</div></div>
          <div class="meta-r"><div class="meta-k">الوقت الزمني</div><div class="meta-v">${p.time}</div></div>
          <div class="meta-r"><div class="meta-k">مؤشر الأداء</div><div class="meta-v" style="color:var(--teal)">${p.kpi}</div></div>
          <div class="meta-r"><div class="meta-k">الميزانية</div><div class="meta-v" style="color:var(--gold);font-weight:800">${fmt(p.budget)} ر.س</div></div>
        </div>
        <div class="target-sec">
          <div class="ts-t">🎯 المستهدف السنوي وتوزيعه على الأرباع</div>
          <div class="annual-row"><label>المستهدف السنوي</label><input class="annual-in" type="number" value="${annualTarget(p.id)||''}" ${editable?`onchange="psv('${p.id}','annual_target',this.value)"`:'disabled'} placeholder="الهدف السنوي"></div>
          <div class="q-splits">${qHTML}</div>
          <div class="prog-bar"><div class="prog-bar-f" style="width:${pct}%;background:${col}"></div></div>
          <div style="display:flex;justify-content:space-between;margin-top:7px"><div style="font-size:.75rem;color:var(--text-dim)">إجمالي الإنجاز: <strong style="color:${col}">${pct}%</strong></div><div class="badge" style="background:${sl.bg};color:${sl.c}">${sl.t}</div></div>
        </div>
      </div>
      <div>
        <div style="font-size:.7rem;font-weight:700;color:var(--text-muted);text-transform:uppercase;letter-spacing:.5px;margin-bottom:8px">المشاريع الفرعية</div>
        <div class="sub-list">${subs}</div>
        <div class="notes-box">
          <div style="font-size:.71rem;font-weight:800;color:var(--text-muted);text-transform:uppercase;letter-spacing:.5px">📝 ملاحظات البرنامج</div>
          <textarea class="notes-area" rows="3" placeholder="التحديات، المخاطر، الإنجازات..." ${editable?`onchange="psv('${p.id}','notes',this.value)"`:'disabled'}>${pg(p.id,'notes','')}</textarea>
        </div>
      </div>
    </div></div>
  </div>`;
}
function togglePC(id){const c=document.getElementById('pc-'+id);c.classList.toggle('open');c.querySelector('.pc-chev').textContent=c.classList.contains('open')?'▲':'▼';}

function renderGoal(gn){
  const c=document.getElementById('list'+gn);if(!c)return;
  const editable=canEditPlan();
  const pp=PROGRAMS.filter(p=>p.goal===gn);
  c.innerHTML=pp.map(p=>buildCard(p,editable)).join('')||'<div style="text-align:center;padding:40px;color:var(--text-muted)">لا توجد برامج بعد</div>';
  const b=pp.reduce((s,p)=>s+(+p.budget||0),0);
  const sub=document.getElementById('g'+gn+'-sub');if(sub)sub.textContent=`${pp.length} برامج ومشاريع · ${fmt(b)} ر.س ميزانية معتمدة`;
}

/* MY TASKS (employee) */
function renderMyTasks(){
  if(!isEmployee())return;
  const mine=PROGRAMS.filter(p=>p.assignee===CU.empId);
  const c=document.getElementById('list-mytasks');
  document.getElementById('mytasks-empty').style.display=mine.length?'none':'block';
  c.innerHTML=mine.map(p=>buildCard(p,true)).join('');
}

/* ════════════════════════════════════════════
   KPIS
   ════════════════════════════════════════════ */
function renderKPIs(){
  const c=document.getElementById('kpi-container');if(!c)return;
  const colors=['var(--teal)','var(--blue)','var(--violet)'];
  c.innerHTML=[1,2,3].map(gn=>{const pp=PROGRAMS.filter(x=>x.goal===gn);if(!pp.length)return'';
    return`<div class="card" style="margin-bottom:18px"><div style="font-size:.92rem;font-weight:800;color:${colors[gn-1]};margin-bottom:12px;font-family:'Cairo',sans-serif">الهدف ${['الأول','الثاني','الثالث'][gn-1]}: ${GOAL_NAMES[gn-1]}</div>
    <div class="tbl-wrap"><table class="tbl"><thead><tr><th>البرنامج</th><th>المسؤول</th><th>الميزانية</th><th>السنوي</th><th>المحقق</th><th>النسبة</th><th>الحالة</th></tr></thead><tbody>
    ${pp.map(p=>{const pct=progPct(p.id);const col=pCol(pct);const sl=stLabel(pct);return`<tr><td style="font-weight:700;color:${p.color}">برنامج ${p.name}</td><td style="font-size:.74rem;color:var(--text-dim)">${p.assignee?empName(p.assignee):'—'}</td><td style="color:var(--gold)">${fmt(p.budget)}</td><td style="font-weight:800">${annualTarget(p.id)||'—'}</td><td style="color:var(--teal);font-weight:700">${totalAch(p.id)||'—'}</td><td style="font-weight:900;color:${col};font-family:'Cairo',sans-serif">${pct}%</td><td><span class="badge" style="background:${sl.bg};color:${sl.c}">${sl.t}</span></td></tr>`;}).join('')}
    </tbody></table></div></div>`;}).join('');
}

/* ════════════════════════════════════════════
   MANAGE (ADMIN)
   ════════════════════════════════════════════ */
function renderManage(){
  applyGoalNames();
  const tb=document.getElementById('manage-tbody');
  tb.innerHTML=PROGRAMS.map(p=>`<tr>
    <td style="font-weight:700">برنامج ${p.name}</td>
    <td style="font-size:.74rem">${GOAL_NAMES[p.goal-1]}</td>
    <td style="color:var(--gold);font-weight:700">${fmt(p.budget)}</td>
    <td style="font-size:.74rem">${p.assignee?empName(p.assignee):'—'}</td>
    <td>${pctMini(progPct(p.id))}</td>
    <td style="white-space:nowrap"><button class="btn btn-teal" style="padding:4px 10px;font-size:.72rem" onclick="openProgModal('${p.id}')">✏️</button> <button class="btn btn-rose" style="padding:4px 10px;font-size:.72rem" onclick="delProg('${p.id}')">🗑️</button></td>
  </tr>`).join('');
}
function pctMini(pct){const col=pCol(pct);return`<div class="pctbar"><div class="pctbar-m"><div class="pctbar-f" style="width:${pct}%;background:${col}"></div></div><span style="font-size:.72rem;font-weight:800;color:${col}">${pct}%</span></div>`;}

/* Program modal */
function fillAssigneeOptions(sel,selected){sel.innerHTML='<option value="">— بدون —</option>'+EMPLOYEES.map(e=>`<option value="${e.id}" ${selected===e.id?'selected':''}>${e.name}</option>`).join('');}
function openProgModal(id){
  if(!isAdmin())return;
  fillAssigneeOptions(document.getElementById('pm-assignee'),'');
  if(id){const p=PROGRAMS.find(x=>x.id===id);document.getElementById('prog-modal-t').textContent='✏️ تعديل برنامج';
    document.getElementById('pm-id').value=p.id;document.getElementById('pm-name').value=p.name;document.getElementById('pm-goal').value=p.goal;document.getElementById('pm-budget').value=p.budget;document.getElementById('pm-exec').value=p.exec;document.getElementById('pm-time').value=p.time;document.getElementById('pm-kpi').value=p.kpi;document.getElementById('pm-target').value=annualTarget(p.id);document.getElementById('pm-desc').value=p.desc||p.subs[0].desc;document.getElementById('pm-q1').value=p.qPct[0];document.getElementById('pm-q2').value=p.qPct[1];document.getElementById('pm-q3').value=p.qPct[2];document.getElementById('pm-q4').value=p.qPct[3];
    fillAssigneeOptions(document.getElementById('pm-assignee'),p.assignee);
  }else{document.getElementById('prog-modal-t').textContent='➕ إضافة برنامج جديد';['pm-id','pm-name','pm-budget','pm-exec','pm-time','pm-kpi','pm-target','pm-desc'].forEach(i=>document.getElementById(i).value='');document.getElementById('pm-q1').value=25;document.getElementById('pm-q2').value=25;document.getElementById('pm-q3').value=25;document.getElementById('pm-q4').value=25;}
  document.getElementById('prog-modal').classList.add('open');
}
function closeProgModal(){document.getElementById('prog-modal').classList.remove('open');}
function saveProg(){
  const id=document.getElementById('pm-id').value;
  const name=document.getElementById('pm-name').value.trim();if(!name){alert('أدخل اسم البرنامج');return;}
  const goal=parseInt(document.getElementById('pm-goal').value);
  const budget=parseFloat(document.getElementById('pm-budget').value)||0;
  const exec=document.getElementById('pm-exec').value||'إدارة البرامج';
  const time=document.getElementById('pm-time').value||'طوال العام';
  const kpi=document.getElementById('pm-kpi').value||'عدد المستفيدين';
  const target=parseFloat(document.getElementById('pm-target').value)||0;
  const desc=document.getElementById('pm-desc').value||name;
  const assignee=document.getElementById('pm-assignee').value;
  const qPct=[1,2,3,4].map(i=>parseInt(document.getElementById('pm-q'+i).value)||0);
  if(id){const p=PROGRAMS.find(x=>x.id===id);p.name=name;p.goal=goal;p.budget=budget;p.exec=exec;p.time=time;p.kpi=kpi;p.desc=desc;p.assignee=assignee;p.qPct=qPct;p.subs[0]={desc,target:target+'',kpi,budget,q:time};if(!ST.progState)ST.progState={};ST.progState[`${id}_annual_target`]=target+'';
  }else{const colors=['#9c6817','#3a6ea5','#2f7d6b','#7a5ba0','#c4881e','#b54a3a','#3f8a4a'];const nid='p_'+Date.now();PROGRAMS.push({id:nid,goal,name,color:colors[PROGRAMS.length%colors.length],budget,exec,time,kpi,qPct,target,assignee,desc,subs:[{desc,target:target+'',kpi,budget,q:time}]});if(!ST.progState)ST.progState={};ST.progState[`${nid}_annual_target`]=target+'';}
  closeProgModal();recalcAndRender();renderManage();toast('✅ تم حفظ البرنامج',true);
}
function delProg(id){if(!confirm('حذف هذا البرنامج نهائياً؟'))return;PROGRAMS=PROGRAMS.filter(p=>p.id!==id);recalcAndRender();renderManage();toast('🗑️ تم الحذف',true);}

function saveGoalNames(){[1,2,3].forEach(i=>{const v=document.getElementById('gn-'+i).value.trim();if(v)GOAL_NAMES[i-1]=v;});recalcAndRender();toast('✅ تم حفظ أسماء الأهداف',true);}

/* ════════════════════════════════════════════
   TEAM (ADMIN)
   ════════════════════════════════════════════ */
function renderTeam(){
  // assign selects
  document.getElementById('assign-emp').innerHTML=EMPLOYEES.map(e=>`<option value="${e.id}">${e.name}</option>`).join('');
  document.getElementById('assign-prog').innerHTML=PROGRAMS.map(p=>`<option value="${p.id}">برنامج ${p.name}</option>`).join('');
  // team list with each employee's progress
  const c=document.getElementById('team-list');
  c.innerHTML=EMPLOYEES.map(e=>{
    const progs=PROGRAMS.filter(p=>p.assignee===e.id);
    const avg=progs.length?Math.round(progs.reduce((s,p)=>s+progPct(p.id),0)/progs.length):0;
    const col=pCol(avg);
    return`<div style="background:var(--surface-2);border:1px solid var(--border);border-radius:12px;padding:14px;margin-bottom:10px">
      <div style="display:flex;align-items:center;gap:12px;margin-bottom:${progs.length?'12px':'0'}">
        <div style="width:40px;height:40px;border-radius:50%;background:var(--gold-p);color:var(--gold);display:flex;align-items:center;justify-content:center;font-size:17px;flex-shrink:0">👤</div>
        <div style="flex:1"><div style="font-weight:800;font-size:.9rem">${e.name}</div><div style="font-size:.72rem;color:var(--text-muted)">${e.role}${e.user?' · مستخدم: '+e.user:''}</div></div>
        <div style="text-align:left"><div style="font-size:1.1rem;font-weight:900;font-family:'Cairo',sans-serif;color:${col}">${avg}%</div><div style="font-size:.65rem;color:var(--text-muted)">متوسط الإنجاز</div></div>
        <button class="btn btn-rose" style="padding:4px 10px;font-size:.72rem" onclick="delEmp('${e.id}')">🗑️</button>
      </div>
      ${progs.length?`<div style="display:flex;flex-direction:column;gap:6px">${progs.map(p=>{const pct=progPct(p.id);return`<div style="display:flex;align-items:center;gap:10px;font-size:.78rem"><span style="flex:1">📌 برنامج ${p.name}</span><div style="width:120px">${pctMini(pct)}</div><button onclick="unassign('${p.id}')" style="background:var(--rose-p);color:var(--rose);border:none;border-radius:5px;padding:2px 7px;cursor:pointer;font-size:.68rem">إلغاء</button></div>`;}).join('')}</div>`:'<div style="font-size:.74rem;color:var(--text-muted);padding-right:52px">لا توجد برامج مسندة</div>'}
    </div>`;
  }).join('')||'<div style="text-align:center;padding:30px;color:var(--text-muted)">لا يوجد موظفون</div>';
}
function addEmp(){const name=document.getElementById('emp-name').value.trim();if(!name){alert('أدخل اسم الموظف');return;}const role=document.getElementById('emp-role').value||'موظف';const user=document.getElementById('emp-user').value.trim();EMPLOYEES.push({id:'emp_'+Date.now(),name,role,user});['emp-name','emp-role','emp-user'].forEach(i=>document.getElementById(i).value='');syncST();renderTeam();toast('✅ تمت إضافة الموظف',true);}
function delEmp(id){if(!confirm('حذف هذا الموظف؟'))return;EMPLOYEES=EMPLOYEES.filter(e=>e.id!==id);PROGRAMS.forEach(p=>{if(p.assignee===id)p.assignee='';});syncST();renderTeam();recalcAndRender();toast('🗑️ تم الحذف',true);}
function assignProg(){const eid=document.getElementById('assign-emp').value;const pid=document.getElementById('assign-prog').value;if(!eid||!pid)return;const p=PROGRAMS.find(x=>x.id===pid);p.assignee=eid;syncST();renderTeam();recalcAndRender();toast('🔗 تم إسناد البرنامج',true);}
function unassign(pid){const p=PROGRAMS.find(x=>x.id===pid);p.assignee='';syncST();renderTeam();recalcAndRender();toast('تم إلغاء الإسناد',true);}

/* ════════════════════════════════════════════
   FINANCE CALCULATIONS
   ════════════════════════════════════════════ */
function totRevenue(){return REVENUES.reduce((s,r)=>s+(+r.amount||0),0);}
function totExpense(){return EXPENSES.reduce((s,e)=>s+(+e.amount||0),0);}
function surplus(){return totRevenue()-totExpense();}
function totBalances(){return BALANCES.reduce((s,b)=>s+(+b.amount||0),0);}
function cashFree(){return BALANCES.filter(b=>b.type==='متاح').reduce((s,b)=>s+(+b.amount||0),0);}
function cashRestr(){return BALANCES.filter(b=>b.type==='مقيّد').reduce((s,b)=>s+(+b.amount||0),0);}
function totLiab(){return LIABILITIES.reduce((s,l)=>s+(+l.amount||0),0);}
function spendRate(){const r=totRevenue();return r?Math.min(100,Math.round(totExpense()/r*100)):0;}

const MONTHS=['يناير','فبراير','مارس','أبريل','مايو','يونيو','يوليو','أغسطس','سبتمبر','أكتوبر','نوفمبر','ديسمبر'];
let FCH={};
function mkFC(id,cfg){if(FCH[id])FCH[id].destroy();const c=document.getElementById(id);if(!c)return;FCH[id]=new Chart(c,cfg);}

function renderFinance(){
  // KPIs
  document.getElementById('f-revenue').textContent=fmt(totRevenue());
  document.getElementById('f-expense').textContent=fmt(totExpense());
  const sur=surplus();const surEl=document.getElementById('f-surplus');surEl.textContent=fmt(sur);surEl.style.color=sur>=0?'var(--gold)':'var(--rose)';
  document.getElementById('f-totbal').textContent=fmt(totBalances());
  document.getElementById('f-cashfree').textContent=fmt(cashFree());
  document.getElementById('f-cashrestr').textContent=fmt(cashRestr());
  document.getElementById('f-liabilities').textContent=fmt(totLiab());
  document.getElementById('f-spendrate').textContent=spendRate()+'%';
  document.getElementById('fbar-exp').style.width=spendRate()+'%';
  document.getElementById('fbar-rate').style.width=spendRate()+'%';
  document.getElementById('fbar-sur').style.width=(totRevenue()?Math.max(0,Math.round(sur/totRevenue()*100)):0)+'%';
  renderFinCharts();
  renderRevTable();renderExpTable();renderBalTable();renderLiabTable();
  renderWelcomeStats();
}

function renderFinCharts(){
  const GOLD='#9c6817',ROSE='#b54a3a',TEAL='#2f7d6b',BLUE='#3a6ea5',VIO='#7a5ba0',AMB='#c4881e',GRN='#3f8a4a';
  const tick={color:'rgba(106,82,56,.85)',font:{family:"'Cairo',sans-serif",size:10}};
  const grid={color:'rgba(124,82,40,.08)'};
  const leg={labels:{color:'rgba(106,82,56,.9)',font:{family:"'Cairo',sans-serif",size:11},boxWidth:10,padding:12}};
  const yA={ticks:{...tick,callback:v=>v>=1e6?(v/1e6).toFixed(1)+'م':v>=1e3?Math.round(v/1e3)+'ك':v},grid};
  const xA={ticks:tick,grid};
  // monthly rev vs exp
  const revByM=MONTHS.map(m=>REVENUES.filter(r=>r.month===m).reduce((s,r)=>s+(+r.amount||0),0));
  const expByM=MONTHS.map(m=>EXPENSES.filter(e=>e.month===m).reduce((s,e)=>s+(+e.amount||0),0));
  const lastM=Math.max(...MONTHS.map((m,i)=>revByM[i]||expByM[i]?i:0))+1;
  const ml=MONTHS.slice(0,Math.max(6,lastM)),mr=revByM.slice(0,ml.length),me=expByM.slice(0,ml.length);
  mkFC('fc-monthly',{type:'bar',data:{labels:ml,datasets:[{label:'إيرادات',data:mr,backgroundColor:GRN+'30',borderColor:GRN,borderWidth:2,borderRadius:4},{label:'مصروفات',data:me,backgroundColor:ROSE+'30',borderColor:ROSE,borderWidth:2,borderRadius:4}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:leg},scales:{x:xA,y:yA}}});
  // exp by category
  const ecats=[...new Set(EXPENSES.map(e=>e.cat))];
  const ecols=[GOLD,BLUE,AMB,VIO,TEAL,ROSE,GRN];
  mkFC('fc-expcat',{type:'doughnut',data:{labels:ecats,datasets:[{data:ecats.map(c=>EXPENSES.filter(e=>e.cat===c).reduce((s,e)=>s+(+e.amount||0),0)),backgroundColor:ecols.map(c=>c+'b3'),borderColor:ecols,borderWidth:2,hoverOffset:5}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:leg},cutout:'58%'}});
  // rev sources
  const rcats=[...new Set(REVENUES.map(r=>r.cat))];
  const rcols=[GRN,TEAL,BLUE,AMB,VIO,GOLD];
  mkFC('fc-revsrc',{type:'doughnut',data:{labels:rcats,datasets:[{data:rcats.map(c=>REVENUES.filter(r=>r.cat===c).reduce((s,r)=>s+(+r.amount||0),0)),backgroundColor:rcols.map(c=>c+'b3'),borderColor:rcols,borderWidth:2,hoverOffset:5}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:leg},cutout:'58%'}});
  // balance free vs restricted
  mkFC('fc-balance',{type:'bar',data:{labels:['متاح','مقيّد'],datasets:[{label:'الرصيد',data:[cashFree(),cashRestr()],backgroundColor:[TEAL+'40',AMB+'40'],borderColor:[TEAL,AMB],borderWidth:2,borderRadius:6}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},scales:{x:xA,y:yA}}});
  // cumulative surplus
  let cum=0;const cumData=ml.map((m,i)=>{cum+=(mr[i]||0)-(me[i]||0);return cum;});
  mkFC('fc-cumulative',{type:'line',data:{labels:ml,datasets:[{label:'الفائض التراكمي',data:cumData,borderColor:GOLD,backgroundColor:GOLD+'20',borderWidth:2.5,fill:true,tension:.35,pointRadius:3,pointBackgroundColor:GOLD}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},scales:{x:xA,y:yA}}});
}

function revCatBadge(c){const m={تبرعات:'var(--green)',اشتراكات:'var(--teal)',أوقاف:'var(--blue)',زكوات:'var(--violet)',استثمارات:'var(--amber)',هبات:'var(--gold)'};const col=m[c]||'var(--text-dim)';return`<span class="badge" style="background:${col}1a;color:${col}">${c}</span>`;}
function expCatBadge(c){const m={برامج:'var(--blue)',رواتب:'var(--gold)',إداري:'var(--text-muted)',تقنية:'var(--violet)',زكوات:'var(--violet)',سفر:'var(--teal)',مشتريات:'var(--amber)'};const col=m[c]||'var(--text-dim)';return`<span class="badge" style="background:${col}1a;color:${col}">${c}</span>`;}
function typeBadge(t){const col=t==='مقيّد'?'var(--amber)':'var(--teal)';return`<span class="badge" style="background:${col}1a;color:${col}">${t==='مقيّد'?'🔒 مقيّد':'💵 متاح'}</span>`;}
function statusBadge(s){const m={'مكتمل':'var(--green)','قيد التنفيذ':'var(--amber)','معتمد':'var(--blue)','موقوف':'var(--rose)'};const col=m[s]||'var(--text-dim)';return`<span class="badge" style="background:${col}1a;color:${col}">${s}</span>`;}

function renderRevTable(){
  const acc=isAccountant();
  document.getElementById('rev-tbody').innerHTML=REVENUES.map((r,i)=>`<tr>
    <td style="color:var(--text-muted)">${i+1}</td>
    <td style="font-weight:700" ${acc?`class="edit" contenteditable onblur="updFin('rev','${r.id}','source',this.textContent)"`:''}>${r.source}</td>
    <td>${revCatBadge(r.cat)}</td>
    <td style="color:var(--green);font-weight:700" ${acc?`class="edit" contenteditable onblur="updFin('rev','${r.id}','amount',this.textContent)"`:''}>${fmt(r.amount)}</td>
    <td>${typeBadge(r.type)}</td>
    <td style="font-size:.74rem">${r.month}</td>
    <td style="font-size:.72rem;color:var(--text-muted)">${r.notes||'—'}</td>
    ${acc?`<td><button onclick="delFin('rev','${r.id}')" style="background:var(--rose-p);color:var(--rose);border:none;border-radius:6px;padding:3px 8px;cursor:pointer;font-size:.7rem">🗑️</button></td>`:'<td class="fin-act" style="display:none"></td>'}
  </tr>`).join('');
  document.getElementById('rev-foot').innerHTML=`<span>عدد البنود: <strong>${REVENUES.length}</strong></span><span>الإجمالي: <strong style="color:var(--green)">${fmt(totRevenue())} ر.س</strong></span>`;
}
function renderExpTable(){
  const acc=isAccountant();
  document.getElementById('exp-tbody').innerHTML=EXPENSES.map((e,i)=>`<tr>
    <td style="color:var(--text-muted)">${i+1}</td>
    <td style="font-weight:700" ${acc?`class="edit" contenteditable onblur="updFin('exp','${e.id}','item',this.textContent)"`:''}>${e.item}</td>
    <td>${expCatBadge(e.cat)}</td>
    <td style="color:var(--rose);font-weight:700" ${acc?`class="edit" contenteditable onblur="updFin('exp','${e.id}','amount',this.textContent)"`:''}>${fmt(e.amount)}</td>
    <td style="font-size:.74rem">${e.month}</td>
    <td>${statusBadge(e.status)}</td>
    <td style="font-size:.72rem;color:var(--text-muted)">${e.notes||'—'}</td>
    ${acc?`<td><button onclick="delFin('exp','${e.id}')" style="background:var(--rose-p);color:var(--rose);border:none;border-radius:6px;padding:3px 8px;cursor:pointer;font-size:.7rem">🗑️</button></td>`:'<td class="fin-act" style="display:none"></td>'}
  </tr>`).join('');
  document.getElementById('exp-foot').innerHTML=`<span>عدد البنود: <strong>${EXPENSES.length}</strong></span><span>الإجمالي: <strong style="color:var(--rose)">${fmt(totExpense())} ر.س</strong></span>`;
}
function renderBalTable(){
  const acc=isAccountant();
  document.getElementById('bal-tbody').innerHTML=BALANCES.map(b=>`<tr>
    <td style="font-weight:700" ${acc?`class="edit" contenteditable onblur="updFin('bal','${b.id}','name',this.textContent)"`:''}>${b.name}</td>
    <td>${typeBadge(b.type)}</td>
    <td style="color:var(--blue);font-weight:700" ${acc?`class="edit" contenteditable onblur="updFin('bal','${b.id}','amount',this.textContent)"`:''}>${fmt(b.amount)}</td>
    ${acc?`<td><button onclick="delFin('bal','${b.id}')" style="background:var(--rose-p);color:var(--rose);border:none;border-radius:6px;padding:3px 7px;cursor:pointer;font-size:.68rem">🗑️</button></td>`:'<td class="fin-act" style="display:none"></td>'}
  </tr>`).join('')+`<tr style="background:var(--surface-3)"><td colspan="2" style="font-weight:800">الإجمالي</td><td style="font-weight:900;color:var(--blue)">${fmt(totBalances())}</td>${acc?'<td></td>':'<td class="fin-act" style="display:none"></td>'}</tr>`;
}
function renderLiabTable(){
  const acc=isAccountant();
  document.getElementById('liab-tbody').innerHTML=LIABILITIES.map(l=>`<tr>
    <td style="font-weight:700" ${acc?`class="edit" contenteditable onblur="updFin('liab','${l.id}','name',this.textContent)"`:''}>${l.name}</td>
    <td style="font-size:.74rem" ${acc?`class="edit" contenteditable onblur="updFin('liab','${l.id}','due',this.textContent)"`:''}>${l.due}</td>
    <td style="color:var(--violet);font-weight:700" ${acc?`class="edit" contenteditable onblur="updFin('liab','${l.id}','amount',this.textContent)"`:''}>${fmt(l.amount)}</td>
    ${acc?`<td><button onclick="delFin('liab','${l.id}')" style="background:var(--rose-p);color:var(--rose);border:none;border-radius:6px;padding:3px 7px;cursor:pointer;font-size:.68rem">🗑️</button></td>`:'<td class="fin-act" style="display:none"></td>'}
  </tr>`).join('')+`<tr style="background:var(--surface-3)"><td colspan="2" style="font-weight:800">الإجمالي</td><td style="font-weight:900;color:var(--violet)">${fmt(totLiab())}</td>${acc?'<td></td>':'<td class="fin-act" style="display:none"></td>'}</tr>`;
}

function updFin(kind,id,field,val){
  if(!isAccountant())return;
  const arr={rev:REVENUES,exp:EXPENSES,bal:BALANCES,liab:LIABILITIES}[kind];
  const item=arr.find(x=>x.id===id);if(!item)return;
  if(field==='amount'){item.amount=parseFloat(val.replace(/,/g,'').trim())||0;}
  else item[field]=val.trim();
  syncST();toast('💾 تعديل مالي — اضغط حفظ',true);renderFinance();
}
function delFin(kind,id){
  if(!isAccountant()||!confirm('حذف هذا البند؟'))return;
  if(kind==='rev')REVENUES=REVENUES.filter(x=>x.id!==id);
  if(kind==='exp')EXPENSES=EXPENSES.filter(x=>x.id!==id);
  if(kind==='bal')BALANCES=BALANCES.filter(x=>x.id!==id);
  if(kind==='liab')LIABILITIES=LIABILITIES.filter(x=>x.id!==id);
  syncST();renderFinance();toast('🗑️ تم الحذف',true);
}

/* Finance modal */
function openFinModal(kind){
  if(!isAccountant())return;
  document.getElementById('fm-kind').value=kind;
  const f=document.getElementById('fm-fields');
  const titles={rev:'➕ إضافة إيراد',exp:'➕ إضافة مصروف',bal:'➕ إضافة رصيد',liab:'➕ إضافة التزام'};
  document.getElementById('fin-modal-t').textContent=titles[kind];
  const monthOpts=MONTHS.map(m=>`<option>${m}</option>`).join('');
  if(kind==='rev'){f.innerHTML=`<div class="modal-grid">
    <div class="field"><label>مصدر الإيراد</label><input class="ctrl" id="x-source" placeholder="مثال: تبرعات"></div>
    <div class="field"><label>التصنيف</label><select class="ctrl" id="x-cat"><option>تبرعات</option><option>اشتراكات</option><option>أوقاف</option><option>زكوات</option><option>استثمارات</option><option>هبات</option></select></div>
    <div class="field"><label>المبلغ (ر.س)</label><input class="ctrl" type="number" id="x-amount" placeholder="0"></div>
    <div class="field"><label>النوع</label><select class="ctrl" id="x-type"><option>متاح</option><option>مقيّد</option></select></div>
    <div class="field"><label>الشهر</label><select class="ctrl" id="x-month">${monthOpts}</select></div>
    <div class="field"><label>ملاحظات</label><input class="ctrl" id="x-notes"></div></div>`;}
  else if(kind==='exp'){f.innerHTML=`<div class="modal-grid">
    <div class="field"><label>بند المصروف</label><input class="ctrl" id="x-item" placeholder="مثال: رواتب"></div>
    <div class="field"><label>التصنيف</label><select class="ctrl" id="x-cat"><option>برامج</option><option>رواتب</option><option>إداري</option><option>تقنية</option><option>زكوات</option><option>سفر</option><option>مشتريات</option></select></div>
    <div class="field"><label>المبلغ (ر.س)</label><input class="ctrl" type="number" id="x-amount" placeholder="0"></div>
    <div class="field"><label>الشهر</label><select class="ctrl" id="x-month">${monthOpts}</select></div>
    <div class="field"><label>الحالة</label><select class="ctrl" id="x-status"><option>معتمد</option><option>قيد التنفيذ</option><option>مكتمل</option><option>موقوف</option></select></div>
    <div class="field"><label>ملاحظات</label><input class="ctrl" id="x-notes"></div></div>`;}
  else if(kind==='bal'){f.innerHTML=`<div class="modal-grid">
    <div class="field"><label>اسم الحساب / الرصيد</label><input class="ctrl" id="x-name" placeholder="مثال: حساب جاري"></div>
    <div class="field"><label>النوع</label><select class="ctrl" id="x-type"><option>متاح</option><option>مقيّد</option></select></div>
    <div class="field"><label>المبلغ (ر.س)</label><input class="ctrl" type="number" id="x-amount" placeholder="0"></div></div>`;}
  else if(kind==='liab'){f.innerHTML=`<div class="modal-grid">
    <div class="field"><label>الالتزام</label><input class="ctrl" id="x-name" placeholder="مثال: مستحقات موردين"></div>
    <div class="field"><label>تاريخ الاستحقاق</label><input class="ctrl" id="x-due" placeholder="مثال: الربع الثاني"></div>
    <div class="field"><label>المبلغ (ر.س)</label><input class="ctrl" type="number" id="x-amount" placeholder="0"></div></div>`;}
  document.getElementById('fin-modal').classList.add('open');
}
function closeFinModal(){document.getElementById('fin-modal').classList.remove('open');}
function saveFinItem(){
  if(!isAccountant())return;
  const kind=document.getElementById('fm-kind').value;
  const amount=parseFloat(document.getElementById('x-amount').value)||0;
  if(kind==='rev'){const source=document.getElementById('x-source').value.trim();if(!source){alert('أدخل المصدر');return;}REVENUES.push({id:'r'+Date.now(),source,cat:document.getElementById('x-cat').value,amount,type:document.getElementById('x-type').value,month:document.getElementById('x-month').value,notes:document.getElementById('x-notes').value.trim()});}
  else if(kind==='exp'){const item=document.getElementById('x-item').value.trim();if(!item){alert('أدخل البند');return;}EXPENSES.push({id:'e'+Date.now(),item,cat:document.getElementById('x-cat').value,amount,month:document.getElementById('x-month').value,status:document.getElementById('x-status').value,notes:document.getElementById('x-notes').value.trim()});}
  else if(kind==='bal'){const name=document.getElementById('x-name').value.trim();if(!name){alert('أدخل الاسم');return;}BALANCES.push({id:'b'+Date.now(),name,type:document.getElementById('x-type').value,amount});}
  else if(kind==='liab'){const name=document.getElementById('x-name').value.trim();if(!name){alert('أدخل الالتزام');return;}LIABILITIES.push({id:'l'+Date.now(),name,due:document.getElementById('x-due').value.trim()||'—',amount});}
  closeFinModal();syncST();renderFinance();toast('✅ تمت الإضافة',true);
}

/* EXPORT EXCEL */
function finExport(){
  if(typeof XLSX==='undefined'){alert('مكتبة Excel غير محملة');return;}
  const wb=XLSX.utils.book_new();
  // summary
  const sum=[['البند','المبلغ (ر.س)'],
    ['إجمالي الإيرادات',totRevenue()],['إجمالي المصروفات',totExpense()],['الفائض / العجز',surplus()],
    ['الرصيد النقدي المتاح',cashFree()],['الرصيد النقدي المقيّد',cashRestr()],['إجمالي الأرصدة',totBalances()],
    ['إجمالي الالتزامات',totLiab()],['نسبة الصرف',spendRate()+'%']];
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(sum),'الملخص المالي');
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([['#','المصدر','التصنيف','المبلغ','النوع','الشهر','ملاحظات'],...REVENUES.map((r,i)=>[i+1,r.source,r.cat,+r.amount,r.type,r.month,r.notes]),['','','الإجمالي',totRevenue(),'','','']]),'الإيرادات');
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([['#','البند','التصنيف','المبلغ','الشهر','الحالة','ملاحظات'],...EXPENSES.map((e,i)=>[i+1,e.item,e.cat,+e.amount,e.month,e.status,e.notes]),['','','الإجمالي',totExpense(),'','','']]),'المصروفات');
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([['الحساب','النوع','المبلغ'],...BALANCES.map(b=>[b.name,b.type,+b.amount]),['الإجمالي','',totBalances()]]),'الأرصدة');
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([['الالتزام','الاستحقاق','المبلغ'],...LIABILITIES.map(l=>[l.name,l.due,+l.amount]),['الإجمالي','',totLiab()]]),'الالتزامات');
  XLSX.writeFile(wb,`التقرير_المالي_صندوق_الأقربين_${new Date().toISOString().slice(0,10)}.xlsx`);
  toast('📥 تم تصدير التقرير المالي',false);
}

/* ════════════════════════════════════════════
   BYLAWS
   ════════════════════════════════════════════ */
function buildBylaws(){
  document.getElementById('bylaws-doc').innerHTML=`
  <div class="bylaws-hero">
    <div class="crest">📜</div>
    <div style="font-size:1.3rem;font-weight:900;font-family:'Cairo',sans-serif;color:var(--text)">اللائحة الأساسية لصندوق الأقربين</div>
    <div style="font-size:.85rem;color:var(--text-muted);margin-top:6px">عائلة العريفي — عرافا حائل</div>
    <div class="bylaws-meta"><span>📋 رقم الترخيص: 7145</span><span>📅 بتاريخ: 1445/5/29هـ</span><span>🏛️ معتمدة من المركز الوطني لتنمية القطاع غير الربحي</span></div>
  </div>

  <h2>الباب الأول: التعريفات والتأسيس والأهداف</h2>
  <h3>التعريفات والتأسيس</h3>
  <p>تأسس صندوق الأقربين بموجب نظام الجمعيات والمؤسسات الأهلية والقواعد التنظيمية للصناديق العائلية الصادرة عن المركز الوطني لتنمية القطاع غير الربحي. ونوع الصندوق: <b>صندوق عام</b>.</p>
  <p><b>العائلة:</b> أفراد ينتسبون بالاسم إلى عائلة العريفي، عرافا حائل. ويخدم الصندوق كل من كان حفيداً لمحمد الأول العريفي وحسن الأول العريفي.</p>
  <p><b>المقر الرئيس:</b> منطقة حائل، مدينة حائل، حي الزبارة / أبو جرف. ونطاق تقديم الخدمات: كافة مناطق المملكة.</p>

  <h3>الأهداف والأغراض (المادة الخامسة)</h3>
  <p>يهدف الصندوق — دون أن يكون من أغراضه الربح المادي — إلى تحقيق الآتي:</p>
  <ul>
    <li>تحقيق صلة الرحم بين أفراد العائلة.</li>
    <li>تحقيق التكافل الاجتماعي بين أفراد العائلة.</li>
    <li>بث روح التآلف والرحمة بين أفراد العائلة.</li>
    <li>إصلاح ذات البين في وسط العائلة.</li>
    <li>التعاون على البر والتقوى بين أفراد العائلة.</li>
    <li>تنظيم أوجه الإحسان بين أفراد العائلة.</li>
    <li>المساهمة في تعليم وتدريب وتطوير أفراد العائلة.</li>
  </ul>

  <h2>الباب الثاني: التنظيم الإداري ومجلس الأمناء</h2>
  <p>يتكون الصندوق من: مجلس الأمناء، والإدارة التنفيذية، واللجان الدائمة أو المؤقتة.</p>
  <h3>مجلس الأمناء</h3>
  <p>يدير الصندوق مجلس أمناء لا يقل عن ثلاثة أعضاء، ويُشترط في العضو أن يكون سعودياً كامل الأهلية، وألا يكون من العاملين في المركز، وموافقة المركز على تعيينه. ومدة الدورة الواحدة <b>أربع سنوات</b>، وتنعقد الاجتماعات بما لا يقل عن اجتماعين سنوياً.</p>
  <h4>أبرز اختصاصات مجلس الأمناء (المادة السادسة عشرة):</h4>
  <ul>
    <li>اعتماد السياسات العامة والخطة الاستراتيجية والتنفيذية ومتابعة تنفيذها.</li>
    <li>اعتماد الهياكل التنظيمية والأنظمة وضوابط الصرف.</li>
    <li>تزويد المركز بالحساب الختامي والتقارير المالية المدققة خلال أربعة أشهر من نهاية السنة المالية.</li>
    <li>الإشراف على إعداد الموازنة التقديرية واعتمادها.</li>
    <li>تعيين مسؤول تنفيذي للصندوق وتحديد صلاحياته.</li>
    <li>تعيين مراجع حسابات خارجي مرخص له.</li>
    <li>فتح الحسابات البنكية والإشراف على عملياتها.</li>
  </ul>

  <h3>العضوية والاشتراكات</h3>
  <p>تكون العضوية حصراً لمن ينتسب لأسرة العريفي عرافا حائل ويسجل بياناته لدى الصندوق. وهناك نوعان: <b>عضوية أساسية</b> و<b>عضوية مشترك</b> (يُشترط ألا يقل العمر عن 15 سنة والالتزام بالرسوم السنوية).</p>
  <p>يلتزم الأعضاء المشتركون بدفع رسوم سنوية قدرها <b>(600 ريال)</b>.</p>

  <h2>الباب الثالث: موارد الصندوق والسنة المالية</h2>
  <h3>موارد الصندوق المالية (المادة الرابعة والثلاثون)</h3>
  <ul>
    <li>ما يخصصه المؤسسون أو أفراد العائلة من أموال أو تبرعات أو هبات أو أوقاف أو وصايا أو زكوات.</li>
    <li>الاشتراكات الدورية لأفراد العائلة.</li>
    <li>إيرادات الأنشطة ذات العائد المالي.</li>
    <li>عائدات استثمار ممتلكات الصندوق الثابتة والمنقولة.</li>
    <li>الأموال المستقبلة من خارج أفراد العائلة بعد موافقة المركز.</li>
  </ul>
  <h3>أحكام مالية مهمة</h3>
  <ul>
    <li>ينحصر صرف أموال الصندوق في غايات تحقيق أهدافه فقط.</li>
    <li>في حال تلقي أموال زكاة يجب إيداعها في <b>حساب مستقل</b> وتُصرف في مصارفها الشرعية.</li>
    <li>تبدأ السنة المالية من تاريخ صدور الترخيص وتنتهي في ديسمبر، وكل سنة بعدها اثنا عشر شهراً ميلادياً.</li>
    <li>لا يُسحب من أموال الصندوق إلا بتوقيع مشترك بين رئيس مجلس الأمناء ونائبه.</li>
    <li>يلتزم الصندوق بالمعايير المحاسبية الصادرة عن الهيئة السعودية للمحاسبين القانونيين.</li>
  </ul>
  <h3>مكافحة غسل الأموال (المادة الثانية والأربعون)</h3>
  <p>يلتزم الصندوق بأحكام نظام مكافحة غسل الأموال ونظام مكافحة جرائم الإرهاب وتمويله، والاحتفاظ بالسجلات والمستندات المالية لمدة لا تقل عن عشر سنوات، والإبلاغ الفوري للإدارة العامة للتحريات المالية عند الاشتباه.</p>

  <h2>الباب الرابع: التعديل والدمج والحل</h2>
  <p>تُعدّل اللائحة بناءً على اقتراح مجلس الأمناء ودراسة المسؤول التنفيذي، ولا يدخل التعديل حيز النفاذ إلا بعد موافقة المركز.</p>
  <p>تكون إجراءات حل الصندوق الاختياري وفق دراسة مجلس الأمناء ورفع توصية للمؤسسين، ويجوز أن تؤول ممتلكات الصندوق بعد الحل إلى صندوق أو جمعية أو مؤسسة أهلية مسجلة لدى المركز.</p>

  <div style="margin-top:30px;padding:18px;background:var(--surface-2);border-radius:12px;text-align:center;font-size:.82rem;color:var(--text-muted)">
    📄 هذا ملخص لأبرز بنود اللائحة. للاطلاع على النص الكامل (21 صفحة) اضغط على زر <b>تحميل اللائحة (PDF)</b> أعلى الصفحة.
  </div>`;
}

/* ════════════════════════════════════════════
   INIT
   ════════════════════════════════════════════ */
// nothing runs until login
</script>
</body>
</html>
