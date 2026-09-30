<!doctype html>
<html lang="es">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<meta name="theme-color" content="#0b3d91">
<meta name="mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-capable" content="yes">
<title>Cortados v2</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<style>
:root{
  --bg:#f3f5f9; --surface:#ffffff; --surface2:#eef1f6; --text:#141a26; --muted:#5d6678; --line:#dde2ea;
  --brand:#0b3d91; --brand-ink:#ffffff; --accent:#e4002b;
  --s-none:#9aa3b2; --s-llame:#2f6fdb; --s-whatsapp:#2f6fdb; --s-nocontesto:#e8791a; --s-contesto:#7b4bd1;
  --s-promesa:#e3b400; --s-visitar:#8a5a36; --s-pago:#1f9d55; --wa:#1faa53;
  --radius:14px; --shadow:0 1px 2px rgba(16,24,40,.06),0 2px 8px rgba(16,24,40,.06);
}
@media (prefers-color-scheme: dark){
  :root:not([data-theme="light"]){
    --bg:#0e1118; --surface:#171b24; --surface2:#1f2531; --text:#e8ecf3; --muted:#9aa4b5; --line:#2a3140;
    --brand:#3d7bff; --shadow:none;
  }
}
*{box-sizing:border-box;-webkit-tap-highlight-color:transparent}
html,body{margin:0;background:var(--bg);color:var(--text);font:15px/1.4 system-ui,-apple-system,"Segoe UI",Roboto,sans-serif}
button,input,select,textarea{font:inherit;color:inherit}
button{cursor:pointer}
.app{max-width:640px;margin:0 auto;min-height:100vh;padding-bottom:90px}
header.top{background:var(--brand);color:var(--brand-ink);padding:12px 16px 10px;padding-top:max(12px,env(safe-area-inset-top))}
.top-row{display:flex;align-items:center;gap:8px}
.top-row h1{font-size:18px;margin:0;flex:1;letter-spacing:.2px}
.icon-btn{background:rgba(255,255,255,.14);border:0;color:inherit;border-radius:10px;padding:8px 10px;font-weight:600;font-size:14px}
.period-select{margin-top:8px;width:100%;background:rgba(255,255,255,.14);border:0;border-radius:10px;padding:9px 10px;color:#fff;font-weight:600}
.period-select option{color:#000}
.progress{margin-top:10px}
.progress .pnums{display:flex;justify-content:space-between;font-size:13px;opacity:.95}
.bar{height:8px;border-radius:8px;background:rgba(255,255,255,.2);overflow:hidden;display:flex;margin-top:6px}
.bar span{height:100%}
.counts{display:flex;gap:6px;flex-wrap:wrap;margin-top:8px}
.cnt{border:0;border-radius:999px;padding:3px 9px;font-size:12px;font-weight:600;background:rgba(255,255,255,.14);color:#fff;display:flex;align-items:center;gap:5px}
.cnt i{width:10px;height:10px;border-radius:50%;display:inline-block}
.cnt.on{background:#fff;color:#111}
.filters{padding:10px 16px;background:var(--surface);border-bottom:1px solid var(--line);position:sticky;top:0;z-index:4}
.search{width:100%;border:1px solid var(--line);background:var(--surface2);border-radius:10px;padding:10px 12px}
.frow{display:flex;gap:8px;margin-top:8px}
.frow select{flex:1;min-width:0;border:1px solid var(--line);background:var(--surface2);border-radius:10px;padding:8px}
.chips{display:flex;gap:6px;overflow-x:auto;margin-top:8px;padding-bottom:2px;scrollbar-width:none}
.chip{flex:none;border:1px solid var(--line);background:var(--surface2);border-radius:999px;padding:5px 11px;font-size:13px;white-space:nowrap}
.chip.on{background:var(--brand);border-color:var(--brand);color:#fff}
.list{padding:10px 12px;display:flex;flex-direction:column;gap:8px}
.card{position:relative;background:var(--surface);border-radius:var(--radius);box-shadow:var(--shadow);padding:12px 12px 12px 18px;border:1px solid var(--line);text-align:left;width:100%}
.card::before{content:"";position:absolute;left:0;top:0;bottom:0;width:6px;border-radius:var(--radius) 0 0 var(--radius);background:var(--sc,var(--s-none))}
.card .nm{font-weight:700;font-size:16px;display:flex;align-items:center;gap:6px}
.card .meta{color:var(--muted);font-size:13px;margin-top:2px}
.card .row2{display:flex;justify-content:space-between;align-items:center;margin-top:6px;gap:8px;flex-wrap:wrap}
.pill{font-size:12px;font-weight:600;border-radius:999px;padding:3px 9px;background:var(--surface2);border:1px solid var(--line)}
.state-tag{font-size:12px;font-weight:700;color:var(--sc,var(--s-none))}
.rep{font-size:11px;font-weight:700;background:#fff3cd;color:#7a5b00;border-radius:6px;padding:1px 6px}
.empty{text-align:center;color:var(--muted);padding:48px 24px}
.empty .big{font-size:44px}
.btn{border:0;border-radius:12px;padding:13px 16px;font-weight:700;background:var(--brand);color:#fff;width:100%;font-size:16px}
.btn.sec{background:var(--surface2);color:var(--text);border:1px solid var(--line)}
.fab{position:fixed;left:50%;transform:translateX(-50%);bottom:max(16px,env(safe-area-inset-bottom));z-index:6;display:flex;gap:8px;width:calc(100% - 32px);max-width:608px}
.fab .btn{box-shadow:0 6px 18px rgba(0,0,0,.18);white-space:nowrap}
@media (max-width:380px){.fab .btn{font-size:14px;padding:12px 10px}}
/* detalle */
.sheet{position:fixed;inset:0;z-index:20;background:var(--bg);overflow-y:auto;display:none}
.sheet.open{display:block}
.sheet .inner{max-width:640px;margin:0 auto;padding-bottom:40px}
.dhead{position:sticky;top:0;z-index:2;background:var(--surface);border-bottom:1px solid var(--line);padding:10px 12px;padding-top:max(10px,env(safe-area-inset-top));display:flex;align-items:center;gap:8px}
.dhead .pos{flex:1;text-align:center;color:var(--muted);font-size:14px}
.nav{border:1px solid var(--line);background:var(--surface2);border-radius:10px;padding:8px 12px;font-weight:700}
.sec{background:var(--surface);margin:10px 12px;border-radius:var(--radius);border:1px solid var(--line);padding:14px}
.sec h3{margin:0 0 10px;font-size:12px;letter-spacing:.8px;text-transform:uppercase;color:var(--muted)}
.hero{border-left:6px solid var(--sc,var(--s-none))}
.hero .nm{font-size:22px;font-weight:800}
.contract{display:flex;align-items:center;gap:8px;margin-top:6px}
.contract code{font-size:18px;font-weight:700;letter-spacing:.5px;font-family:ui-monospace,Menlo,monospace}
.mini{border:1px solid var(--line);background:var(--surface2);border-radius:8px;padding:5px 10px;font-weight:600;font-size:13px}
.kv{display:grid;grid-template-columns:auto 1fr;gap:6px 12px;font-size:14px}
.kv dt{color:var(--muted)} .kv dd{margin:0;text-align:right;font-weight:600;overflow-wrap:anywhere;min-width:0}
.warnbox{margin-top:12px;background:#fff4e0;color:#6b4200;border:1px solid #f0c77a;border-radius:10px;padding:10px;font-size:13px}
@media (prefers-color-scheme: dark){:root:not([data-theme="light"]) .warnbox{background:#3a2a0c;color:#ffd99a;border-color:#6b4f1d}}
.btn:disabled{opacity:.7}
.kv dd.big{font-size:18px;color:var(--accent)}
.promos{display:grid;grid-template-columns:1fr 1fr;gap:8px}
.promo{border:2px solid var(--line);background:var(--surface2);border-radius:12px;padding:10px;text-align:left}
.promo b{display:block;font-size:13px;color:var(--muted)}
.promo span{font-size:18px;font-weight:800}
.promo.on{border-color:var(--brand);background:color-mix(in srgb,var(--brand) 12%,var(--surface))}
.msg{background:#e7fbe9;color:#0f3d1c;border-radius:12px;padding:12px;font-size:14px;white-space:pre-wrap;border:1px solid #bfe7c8}
@media (prefers-color-scheme: dark){:root:not([data-theme="light"]) .msg{background:#123222;color:#d8f5e1;border-color:#1f5236}}
.msg.off{background:var(--surface2);color:var(--muted);border-color:var(--line);font-style:italic}
.nums{display:flex;flex-direction:column;gap:8px}
.numbtn{display:flex;align-items:center;gap:10px;border:0;border-radius:12px;padding:12px 14px;font-weight:700;color:#fff;text-decoration:none;font-size:15px}
.numbtn small{font-weight:500;opacity:.9;margin-left:auto;font-size:12px}
.numbtn.wa{background:var(--wa)} .numbtn.call{background:var(--brand)}
.numbtn.dis{opacity:.4;pointer-events:none}
.states{display:flex;flex-wrap:wrap;gap:8px}
.st{border:2px solid var(--c);color:var(--text);background:var(--surface);border-radius:999px;padding:7px 12px;font-weight:700;font-size:14px}
.st.on{background:var(--c);color:#fff}
.st[data-s="promesa"].on{color:#1a1400}
textarea{width:100%;min-height:80px;border:1px solid var(--line);background:var(--surface2);border-radius:10px;padding:10px;resize:vertical}
.hist{list-style:none;margin:0;padding:0;font-size:13px}
.hist li{padding:6px 0;border-bottom:1px dashed var(--line);display:flex;gap:8px}
.hist li time{color:var(--muted);flex:none;width:92px}
.muted{color:var(--muted)}
/* modal */
.modal{position:fixed;inset:0;z-index:40;background:rgba(0,0,0,.45);display:none;align-items:flex-end;justify-content:center}
.modal.open{display:flex}
.mbox{background:var(--surface);width:100%;max-width:640px;border-radius:18px 18px 0 0;padding:18px 16px;padding-bottom:max(18px,env(safe-area-inset-bottom));max-height:85vh;overflow:auto}
.mbox h2{margin:0 0 12px;font-size:18px}
.ok{color:var(--s-pago)} .warn{color:#b26a00} .err{color:var(--accent)}
.res{border:1px solid var(--line);border-radius:12px;padding:12px;margin-bottom:10px;font-size:14px}
.res b{font-size:15px}
.pbar{height:6px;background:var(--surface2);border-radius:6px;overflow:hidden;margin-top:6px}
.pbar span{display:block;height:100%;background:var(--brand);width:0;transition:width .15s}
.lst{display:flex;align-items:center;gap:10px;border:1px solid var(--line);border-radius:12px;padding:12px;margin-bottom:8px}
.lst div{flex:1}
.toast{position:fixed;left:50%;bottom:90px;transform:translateX(-50%);background:#111;color:#fff;padding:10px 16px;border-radius:999px;font-size:14px;z-index:60;opacity:0;transition:opacity .2s;pointer-events:none}
.toast.show{opacity:.92}
input[type=date]{width:100%;border:1px solid var(--line);background:var(--surface2);border-radius:10px;padding:10px;margin:6px 0 12px}
</style>
</head>
<body>
<div class="app">
  <header class="top">
    <div class="top-row">
      <h1>Cortados v2</h1>
      <button class="icon-btn" id="btnListados">Mis listados</button>
    </div>
    <select class="period-select" id="selPeriodo"></select>
    <div class="progress" id="progress"></div>
  </header>
  <div class="filters" id="filters">
    <input class="search" id="q" type="search" placeholder="🔍 Buscar nombre, suscriptor o contrato">
    <div class="frow">
      <select id="selPago"></select>
      <select id="selOrden">
        <option value="pdf">Orden del PDF</option>
        <option value="adeudo">Mayor adeudo</option>
        <option value="colonia">Colonia</option>
        <option value="pago">Último pago (antiguo)</option>
      </select>
    </div>
    <div class="chips" id="chipsCol"></div>
  </div>
  <div class="list" id="list"></div>
</div>

<div class="fab">
  <button class="btn" id="btnCargar">📄 Cargar listados</button>
  <button class="btn sec" id="btnExport" style="max-width:140px">📊 Excel</button>
</div>
<input type="file" id="file" accept="application/pdf,.pdf" multiple hidden>

<div class="sheet" id="sheet"><div class="inner" id="detail"></div></div>
<div class="modal" id="modal"><div class="mbox" id="mbox"></div></div>
<div class="toast" id="toast"></div>
<script>
// ===== Parser del Listado único para recuperación =====
// pages: [[{s,x,y}]]  (y crece hacia arriba, como pdf.js)
function parseListado(pages){
  const MONEY=/^\$\s*-?[\d ,]*\d\.\d\d$/;
  const money=s=>parseFloat(s.replace(/[$\s,]/g,''));
  const phone=s=>{s=(s||'').trim(); if(/^S\/T/i.test(s)) return ''; return s.replace(/\D/g,'');};
  const title=s=>s.toLowerCase().replace(/(^|[\s-])(\p{L})/gu,(m,a,b)=>a+b.toUpperCase());
  const PROMOS=['1969','1147','1396','989'];
  const SUSC=/^\d{1,3}(?:\s\d{3})*$/;              // "937", "71 937", "1 234 567"
  const FECHA='(\\d{1,2}-[A-Za-zÁÉÍÓÚáéíóú]{3,4}\\.?-\\d{4})';
  let periodo=null, saldoAl='', esListado=false, totalItems=0;
  const clients=[]; let cur=null;
  pages.forEach((items,pi)=>{
    totalItems+=items.filter(i=>i.s.trim()).length;
    // agrupar en líneas
    const sorted=items.filter(i=>i.s.trim()).sort((a,b)=>b.y-a.y||a.x-b.x);
    const lines=[];
    for(const it of sorted){
      const L=lines[lines.length-1];
      if(L && Math.abs(L.y-it.y)<2.5) L.items.push(it); else lines.push({y:it.y,items:[it]});
    }
    lines.forEach(L=>{L.items.sort((a,b)=>a.x-b.x); L.t=L.items.map(i=>i.s.trim()).join(' ').replace(/\s+/g,' ');});
    let colx=null,secx=null,ofx=null,domx=null;
    for(const L of lines){
      if(/Listado único para recuperación/i.test(L.t)) esListado=true;
      const m=L.t.match(new RegExp('entre\\s+'+FECHA+'\\s+y\\s+'+FECHA,'i'));
      if(m && !periodo) periodo={ini:m[1].toLowerCase(),fin:m[2].toLowerCase()};
      const s=L.t.match(/Saldo (?:calculado al|generado el) día\s*:\s*(\d{4}\/\d{2}\/\d{2})/);
      if(s && !saldoAl) saldoAl=s[1];
      for(const i of L.items){
        const t=i.s.trim();
        if(t==='Colonia') colx=i.x; if(t==='Sector') secx=i.x;
        if(t==='Teléfono oficina') ofx=i.x; if(t==='Domicilio') domx=i.x;
      }
    }
    if(!colx) return;
    if(!secx) secx=colx+150; if(!ofx) ofx=500; if(!domx) domx=70;
    for(let k=0;k<lines.length;k++){
      const L=lines[k];
      const first=L.items[0];
      const esInicio=first && first.x<45 && SUSC.test(first.s.trim())
        && L.items.some(i=>i.x>first.x+5 && i.x<colx-3)          // trae nombre
        && L.items.some(i=>i.x>=colx-3);                         // y columnas a la derecha
      if(esInicio){
        cur={hoja:pi+1,suscriptor:String(parseInt(first.s.replace(/\D/g,''),10)),raw:[]};
        clients.push(cur);
        cur.nombre_raw=L.items.filter(i=>i.x>first.x+5 && i.x<colx-3).map(i=>i.s.trim()).join(' ');
        cur.colonia=L.items.filter(i=>i.x>=colx-3 && i.x<secx-3).map(i=>i.s.trim()).join(' ');
        const N=lines[k+1];
        if(N){
          cur.tel_oficina=phone((N.items.find(i=>i.x>=ofx-40 && /\d{7,}/.test(i.s))||{s:''}).s);
          const dom=N.items.filter(i=>i.x>=domx-12 && i.x<colx-3).map(i=>i.s.trim()).join(' ');
          const ent=N.items.filter(i=>i.x>=colx-3 && i.x<ofx-40).map(i=>i.s.trim()).join(' ');
          cur.direccion=dom+(ent?' · entre '+ent:'');
        }
        continue;
      }
      if(cur) cur.raw.push(L);
    }
  });
  if(!periodo && esListado && saldoAl){ // respaldo: sin rango de fechas en el encabezado
    const [y,m,d]=saldoAl.split('/'); const MES=['ene','feb','mar','abr','may','jun','jul','ago','sep','oct','nov','dic'];
    const f=`${d}-${MES[+m-1]}-${y}`; periodo={ini:f,fin:f};
  }
  clients.forEach(c=>{
    const raw=c.raw, body=raw.map(l=>l.t).join('\n');
    // ---- teléfonos y referencia del domicilio ----
    const ci=raw.findIndex(l=>/^Casa:/.test(l.t));
    const casaL=ci>=0?raw[ci]:null;
    const casa=casaL?casaL.items.filter(i=>i.x<480).map(i=>i.s.trim()).join(' '):'';
    const g=re=>{const m=casa.match(re); return m?phone(m[1]):'';};
    c.casa=g(/Casa\s*:\s*(S\/T|\d+)/);
    c.celular=g(/(?:^|,\s*)Cel\s*:\s*(S\/T|\d+)/);
    c.megafon=g(/Megaf[oó]n\s*:\s*(S\/T|\d+)/i);
    c.whatsapp=g(/Whatsapp\s*:\s*(S\/T|\d+)/i);
    c.web_cel=g(/Web Cel\s*:\s*(S\/T|\d+)/i);
    c.ult_tel=g(/UltTel\s*:\s*(S\/T|\d+)/i);
    let ref=casa.replace(/(?:Casa|Cel|Megaf[oó]n|Whatsapp|Web Cel|UltTel)\s*:\s*(?:S\/T|\d+)\s*,?/gi,'').trim();
    if(ci>=0){ for(let j=ci+1;j<raw.length && !/^Servicio o paquete/.test(raw[j].t);j++) ref+=' '+raw[j].items.filter(i=>i.x<480).map(i=>i.s.trim()).join(' '); }
    // teléfonos extra escritos en las notas (ej. "Cel2: 9981171504")
    const extra=[]; ref=ref.replace(/\b(?:Cel|Tel)\s?\d\s*:\s*(\d{10})\s*,?/gi,(m,n)=>{extra.push(n); return ' ';});
    c.otros_tel=extra.join(',');
    c.referencia=ref.replace(/\s+/g,' ').replace(/^[,\s]+|[,\s]+$/g,'').trim();
    let m=body.match(/Servicios Móviles:\s*([\d,\s]+)/); c.serv_moviles=m?m[1].replace(/\s/g,'').replace(/,+$/,''):'';
    m=body.match(/Fec upa servicio\s*(\d{1,2}\/\d{1,2}\/\d{4})/); c.ultimo_pago=m?m[1]:'';
    // ---- paquete ----
    const pk=raw.findIndex(l=>/^Servicio o paquete/.test(l.t));
    if(pk>=0){
      const L=raw[pk]; const lab=L.items.findIndex(i=>/Servicio o paquete/.test(i.s));
      const parts=[]; for(let j=lab+1;j<L.items.length;j++){const s=L.items[j].s.trim(); if(s.startsWith('$')) break; parts.push(s);}
      const nx=raw[pk+1]; if(nx && !/^(Fecha|Adeudo|Servicios)/.test(nx.t)){ const lim=L.items[lab+1]?L.items[lab+1].x+250:300; parts.push(...nx.items.filter(i=>i.x<lim && !/Fecha|Fin plazo/.test(i.s)).map(i=>i.s.trim()));}
      c.paquete=parts.join(' ').replace(/\s+/g,' ').replace(/^:\s*/,'');
    } else c.paquete='';
    // ---- montos: columnas por posición (x) con respaldo por conteo ----
    const hk=raw.findIndex(l=>/^Adeudo\b/.test(l.t) && !/comodato/.test(l.t));
    const isl=raw.find(l=>/^Importe servicios/.test(l.t));
    let vals=isl?isl.items.filter(i=>MONEY.test(i.s.trim())).map(i=>({x:i.x,v:money(i.s)})):[];
    c.adeudo=null; c.promos={}; c.cols_ok=false;
    if(hk>=0 && raw[hk+1] && vals.length){
      const L1=raw[hk], L2=raw[hk+1];
      const cols=L1.items.map(i=>({x:i.x,t:i.s.trim()}));
      cols.forEach(col=>{ if(/^Importe$/.test(col.t)){ const lab=L2.items.filter(i=>Math.abs(i.x-col.x)<14).sort((a,b)=>Math.abs(a.x-col.x)-Math.abs(b.x-col.x))[0]; col.code=lab?lab.s.trim():''; } });
      const usadas=new Set(); let ok=cols.length>0 && /^Adeudo/.test(cols[0].t);
      const asignado=vals.map(v=>{ let best=-1,d=1e9; cols.forEach((col,ix)=>{const dd=Math.abs(v.x-col.x); if(dd<d){d=dd;best=ix;}}); if(d>20||usadas.has(best)) ok=false; usadas.add(best); return {...v,col:cols[best]}; });
      if(ok){
        const ad=asignado.find(a=>/^Adeudo/.test(a.col.t)); c.adeudo=ad?ad.v:null;
        asignado.forEach(a=>{ if(a.col.code && PROMOS.includes(a.col.code)) c.promos[a.col.code]=a.v; });
        c.cols_ok=c.adeudo!=null;
      } else {
        // respaldo: contar columnas (funciona cuando no hay celdas vacías)
        const l1=L1.t, l2=L2.t;
        const rest=l2.replace('Mes Ant','').replace('de Mes','').trim().split(/\s+/).filter(Boolean);
        const nimp=(l1.match(/Importe/g)||[]).length;
        const codes=nimp?rest.slice(-nimp):[];
        const ncols=1+(l1.includes('Último Día')?1:0)+(l1.includes('Imp Fin')?1:0)+nimp;
        let v=vals.map(a=>a.v); if(v.length<ncols) v=(isl.t.match(/\$\s?-?[\d ]*\d\.\d\d/g)||[]).map(money);
        c.cols_ok=(ncols===v.length); c.adeudo=v.length?v[0]:null;
        if(codes.length && v.length>=codes.length) codes.forEach((code,ix)=>{ if(PROMOS.includes(code)) c.promos[code]=v[v.length-codes.length+ix]; });
      }
    } else if(vals.length){ c.adeudo=vals[0].v; }
    // ---- nombre: 1 apellido y 1 nombre sin asteriscos (respeta "DE LA", "DEL", "Y"...) ----
    const PART=/^(DE|DEL|LA|LAS|LOS|Y|E|MC|MAC|SAN|SANTA|VAN|VON|DA|DI|DO|DOS)$/i;
    const toks=(c.nombre_raw||'').split(/\s+/).filter(Boolean), grupos=[]; let acc=[];
    for(const t of toks){ acc.push(t); if(!PART.test(t)){ grupos.push(acc.join(' ')); acc=[]; } }
    if(acc.length) grupos.push(acc.join(' '));
    const ok=g=>!g.includes('*');
    const ap=grupos.slice(0,2).find(ok), nm=grupos.slice(2).find(ok);
    c.apellido=ap?title(ap):''; c.nombre=nm?title(nm):'';
    c.contrato='325'+c.suscriptor.padStart(7,'0');
    c.saldo_al=saldoAl;
    delete c.raw;
  });
  return {periodo,esListado,totalItems,saldoAl,clients};
}
// ===== Estado y almacenamiento =====
const KEY='cortados_v2';
let db={periodos:{},ui:{periodo:null,q:'',pago:'',orden:'pdf',col:'',estado:'',promHoy:false}};
function load(){try{const s=localStorage.getItem(KEY); if(s){const d=JSON.parse(s); db.periodos=d.periodos||{}; Object.assign(db.ui,d.ui||{});}}catch(e){}}
function save(){try{localStorage.setItem(KEY,JSON.stringify(db));}catch(e){toast('⚠️ No se pudo guardar en este navegador');}}

const ESTADOS=[
 {k:'llame',t:'Llamé',c:'var(--s-llame)',ico:'🔵'},
 {k:'contesto',t:'Contestó',c:'var(--s-contesto)',ico:'🟣'},
 {k:'nocontesto',t:'No contestó',c:'var(--s-nocontesto)',ico:'🟠'},
 {k:'whatsapp',t:'WhatsApp enviado',c:'var(--s-whatsapp)',ico:'🔵'},
 {k:'promesa',t:'Promesa de pago',c:'var(--s-promesa)',ico:'🟡'},
 {k:'pago',t:'Pagó',c:'var(--s-pago)',ico:'✅'},
 {k:'visitar',t:'Visitar domicilio',c:'var(--s-visitar)',ico:'🏠'},
];
const PRIORIDAD=['pago','promesa','visitar','contesto','nocontesto','whatsapp','llame'];
// grupos del contador (Llamé y WhatsApp comparten azul)
const GRUPOS=[
 {k:'pago',t:'Pagó',c:'var(--s-pago)',m:['pago']},
 {k:'promesa',t:'Promesa',c:'var(--s-promesa)',m:['promesa']},
 {k:'visitar',t:'Visitar',c:'var(--s-visitar)',m:['visitar']},
 {k:'contesto',t:'Contestó',c:'var(--s-contesto)',m:['contesto']},
 {k:'nocontesto',t:'No contestó',c:'var(--s-nocontesto)',m:['nocontesto']},
 {k:'intento',t:'Llamé/WA',c:'var(--s-llame)',m:['llame','whatsapp']},
 {k:'none',t:'Sin contactar',c:'var(--s-none)',m:[null]},
];
const est=c=>c.est||(c.est={estados:[],promo:null,promesa:'',notas:'',hist:[]});
function principal(c){const e=est(c).estados; for(const k of PRIORIDAD) if(e.includes(k)) return k; return null;}
function colorDe(k){const e=ESTADOS.find(x=>x.k===k); return e?e.c:'var(--s-none)';}
function nombreDe(k){const e=ESTADOS.find(x=>x.k===k); return e?e.t:'Sin contactar';}

// ===== Utilidades =====
const $=s=>document.querySelector(s);
const esc=s=>String(s??'').replace(/[&<>"']/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m]));
const fmt=n=>n==null?'—':'$'+n.toLocaleString('es-MX',{minimumFractionDigits:2,maximumFractionDigits:2});
const MESES={ene:1,feb:2,mar:3,abr:4,may:5,jun:6,jul:7,ago:8,sep:9,sept:9,set:9,oct:10,nov:11,dic:12};
function fechaNum(s){ // '01-sep-2026' o '31/08/2026'
  if(!s) return 0; let m=s.match(/(\d{1,2})-([a-záéíóú]{3,4})\.?-(\d{4})/i); if(m) return +m[3]*10000+(MESES[m[2].toLowerCase()]||0)*100+ +m[1];
  m=s.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/); if(m) return +m[3]*10000+ +m[2]*100+ +m[1]; return 0;}
const pkey=p=>p.ini+'_'+p.fin;
const plabel=p=>`${p.ini} → ${p.fin}`;
const hoyISO=()=>{const d=new Date(); return d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0');};
function ahora(){const d=new Date(); return d.toLocaleDateString('es-MX',{day:'2-digit',month:'2-digit'})+' '+d.toLocaleTimeString('es-MX',{hour:'2-digit',minute:'2-digit',hour12:false});}
function toast(t){const e=$('#toast'); e.textContent=t; e.classList.add('show'); clearTimeout(e._t); e._t=setTimeout(()=>e.classList.remove('show'),1800);}

// ===== Teléfonos: sanitizar =====
function limpiarNum(n,contrato){
  let d=String(n||'').replace(/\D/g,'');
  if(d.length===13&&d.startsWith('521')) d=d.slice(3); else if(d.length===12&&d.startsWith('52')) d=d.slice(2);
  if(d.length!==10) return null;
  if(d===contrato) return null;               // Web Cel = contrato, no es teléfono
  if(/^(\d)\1{9}$/.test(d)) return null;      // 0000000000, etc.
  if(/^[01]/.test(d)) return null;            // en México ningún número empieza con 0 o 1
  return d;
}
function numeros(c){
  const fuentes=[['WhatsApp',c.whatsapp],['Cel',c.celular],['Megafon',c.megafon],['UltTel',c.ult_tel],['Web Cel',c.web_cel],['Casa',c.casa],['Oficina',c.tel_oficina]];
  (c.otros_tel||'').split(',').forEach(n=>fuentes.push(['Cel2',n]));
  (c.serv_moviles||'').split(',').forEach(n=>fuentes.push(['Móvil MEGA',n]));
  const map=new Map();
  for(const [lab,n] of fuentes){const d=limpiarNum(n,c.contrato); if(!d) continue; if(!map.has(d)) map.set(d,[]); const l=map.get(d); if(!l.includes(lab)) l.push(lab);}
  return [...map].map(([num,labs])=>({num,labs}));
}

// ===== Mensaje WhatsApp =====
const montoPromo=(c,code)=>code==='adeudo'?c.adeudo:(c.promos||{})[code];
const nombrePromo=code=>code==='adeudo'?'Adeudo total':code;
function mensaje(c){
  const code=est(c).promo; const m=code?montoPromo(c,code):null; if(m==null) return null;
  const monto=fmt(m);
  return `Hola ${c.nombre?c.nombre+' ':''}con contrato ${c.contrato} espero te encuentres bien, veo que tienes problemas con tu internet y el día de hoy con un pago de ${monto} te puedes poner al corriente. También si deseas ajustar más el precio de tu contrato puedo ayudarte con ello escríbeme 👌`;
}

// ===== Botón "atrás" del celular: cada pantalla abierta es una capa =====
const capas=[];
function abrirCapa(cerrarFn){ let pushed=false; try{history.pushState({capa:capas.length+1},''); pushed=true;}catch(e){} capas.push({cerrarFn,pushed}); }
function cerrarCapa(){ const top=capas[capas.length-1]; if(!top||top.cerrando) return; top.cerrando=true; if(top.pushed){ history.back(); } else { capas.pop(); top.cerrarFn(); } }
window.addEventListener('popstate',()=>{ const top=capas.pop(); if(top) top.cerrarFn(); });
// dentro de un marco (vista previa) los enlaces tel: necesitan abrirse aparte
const enMarco=(()=>{try{return window.self!==window.top;}catch(e){return true;}})();

// ===== Lectura de PDF =====
if(window.pdfjsLib) pdfjsLib.GlobalWorkerOptions.workerSrc='https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
async function leerPDF(file,onProg){
  if(!window.pdfjsLib) throw new Error('No cargó el lector de PDF. Revisa tu conexión a internet y vuelve a abrir la página.');
  const buf=await file.arrayBuffer();
  const doc=await pdfjsLib.getDocument({data:new Uint8Array(buf)}).promise;
  const pages=[];
  for(let i=1;i<=doc.numPages;i++){
    const p=await doc.getPage(i); const tc=await p.getTextContent();
    pages.push(tc.items.map(it=>({s:it.str,x:it.transform[4],y:it.transform[5]})));
    onProg(i,doc.numPages);
  }
  return {pages,numPages:doc.numPages};
}

function errorES(e){
  const n=(e&&e.name)||'', m=String((e&&e.message)||e);
  if(n==='PasswordException') return 'El PDF tiene contraseña. Pide el listado sin contraseña o quítasela antes de cargarlo.';
  if(n==='InvalidPDFException') return 'El archivo está dañado o no es un PDF válido.';
  if(n==='MissingPDFException'||n==='UnexpectedResponseException') return 'No se pudo abrir el archivo. Intenta descargarlo de nuevo.';
  return m;
}
let cargando=false;
async function cargarArchivos(files){
  if(cargando) return; cargando=true; $('#btnCargar').disabled=true; $('#btnCargar').textContent='⏳ Leyendo…';
  const box=$('#mbox'); abrirModal('<h2>Leyendo listados…</h2><p class="muted" style="margin-top:-6px">No cierres la app mientras termina.</p><div id="resList"></div>',{bloqueado:true});
  const resList=$('#resList'); let ultimo=null;
  for(const f of files){
    const div=document.createElement('div'); div.className='res';
    div.innerHTML=`<b>${esc(f.name)}</b><div class="muted" data-p>Abriendo…</div><div class="pbar"><span></span></div>`;
    resList.appendChild(div);
    const pt=div.querySelector('[data-p]'), pb=div.querySelector('.pbar span');
    try{
      if(!/\.pdf$/i.test(f.name)&&f.type!=='application/pdf') throw new Error('No es un archivo PDF.');
      const {pages,numPages}=await leerPDF(f,(i,n)=>{pt.textContent=`Leyendo hoja ${i} de ${n}…`; pb.style.width=(i/n*100)+'%';});
      const r=parseListado(pages);
      if(r.totalItems<20) throw new Error('El PDF no trae texto (parece escaneado o imagen). No se puede leer con precisión.');
      if(!r.esListado||!r.periodo) throw new Error('Este archivo no es un "Listado único para recuperación".');
      if(!r.clients.length) throw new Error('No se encontraron clientes en el listado.');
      const k=pkey(r.periodo);
      const P=db.periodos[k]||(db.periodos[k]={key:k,ini:r.periodo.ini,fin:r.periodo.fin,cargado:new Date().toISOString(),archivos:[],clientes:{},orden:[]});
      let nuevos=0,ya=0,act=0,viejos=0;
      for(const c of r.clients){
        const prev=P.clientes[c.suscriptor];
        if(prev){
          // mismo cliente en el mismo corte: se refrescan sus datos y se conserva tu avance
          if((c.saldo_al||'')<(prev.saldo_al||'')){viejos++; continue;}
          if(prev.saldo_al&&(c.saldo_al||'')>prev.saldo_al) act++; else ya++;
          const e=prev.est; Object.assign(prev,c); prev.est=e; continue;
        }
        c.est={estados:[],promo:null,promesa:'',notas:'',hist:[]};
        P.clientes[c.suscriptor]=c; P.orden.push(c.suscriptor); nuevos++;
      }
      if(!P.archivos.includes(f.name)) P.archivos.push(f.name);
      const sinTel=r.clients.filter(c=>!numeros(c).length).length;
      const reps=r.clients.filter(c=>Object.values(db.periodos).some(o=>o.key!==k&&o.clientes[c.suscriptor])).length;
      const malos=r.clients.filter(c=>!c.cols_ok).length;
      pt.innerHTML=`<span class="ok">✅ Corte ${esc(plabel(r.periodo))}</span><br>${r.clients.length} clientes · ${numPages} hojas`+
        (ya?`<br>↩︎ ${ya} ya estaban cargados (se conservó tu avance)`:'')+
        (act?`<br>🔄 ${act} actualizados con saldos más recientes (se conservó tu avance)`:'')+
        (viejos?`<br><span class="warn">⚠️ ${viejos} no se actualizaron: este PDF es más viejo que el que ya tenías</span>`:'')+
        (nuevos&&ya?`<br>➕ ${nuevos} nuevos`:'')+
        (reps?`<br>🔁 ${reps} también en otro periodo`:'')+
        (sinTel?`<br><span class="warn">⚠️ ${sinTel} sin teléfono válido</span>`:'')+
        (malos?`<br><span class="warn">⚠️ ${malos} con montos por revisar</span>`:'');
      ultimo=k; save();
    }catch(e){ pt.innerHTML=`<span class="err">❌ ${esc(errorES(e))}</span>`; pb.style.width='0'; }
  }
  cargando=false; $('#btnCargar').disabled=false; $('#btnCargar').textContent='📄 Cargar listados'; bloqueado=false;
  if(ultimo){db.ui.periodo=ultimo; db.ui.col='';db.ui.pago='';db.ui.estado='';db.ui.q='';db.ui.promHoy=false; save();}
  const b=document.createElement('button'); b.className='btn'; b.style.marginTop='6px'; b.textContent=ultimo?'Ver clientes':'Cerrar';
  b.onclick=()=>{cerrarModal(); render();}; box.appendChild(b);
  render();
}

// ===== Vista lista =====
function periodoActual(){const ks=Object.keys(db.periodos); if(!ks.length) return null; if(!db.periodos[db.ui.periodo]){db.ui.periodo=ks.sort((a,b)=>fechaNum(db.periodos[b].fin)-fechaNum(db.periodos[a].fin))[0];} return db.periodos[db.ui.periodo];}
function otrosPeriodos(sus,k){return Object.values(db.periodos).filter(p=>p.key!==k&&p.clientes[sus]).map(p=>p.ini);}
function grupoDe(c){const p=principal(c); return (GRUPOS.find(g=>g.m.includes(p))||GRUPOS[GRUPOS.length-1]).k;}
function filtrados(P){
  const u=db.ui, q=u.q.trim().toLowerCase(), hoy=hoyISO();
  let arr=P.orden.map(s=>P.clientes[s]);
  arr=arr.filter(c=>{
    if(u.col&&c.colonia!==u.col) return false;
    if(u.pago&&c.ultimo_pago!==u.pago) return false;
    if(u.estado&&grupoDe(c)!==u.estado) return false;
    if(u.promHoy&&!promesaPendiente(c)) return false;
    if(q){const h=(c.nombre+' '+c.apellido+' '+c.nombre_raw+' '+c.suscriptor+' '+c.contrato).toLowerCase(); if(!h.includes(q)) return false;}
    return true;});
  const o=u.orden;
  if(o==='adeudo') arr.sort((a,b)=>(b.adeudo||0)-(a.adeudo||0));
  else if(o==='colonia') arr.sort((a,b)=>String(a.colonia).localeCompare(String(b.colonia),'es',{numeric:true}));
  else if(o==='pago') arr.sort((a,b)=>fechaNum(a.ultimo_pago)-fechaNum(b.ultimo_pago));
  return arr;
}
let vista=[];
// promesa con fecha de hoy o ya vencida, y que todavía no pagó
function promesaPendiente(c){const E=est(c); return E.estados.includes('promesa')&&!E.estados.includes('pago')&&E.promesa&&E.promesa<=hoyISO();}
function etiquetaPromesa(c){const E=est(c); if(!E.estados.includes('promesa')||!E.promesa) return ''; const h=hoyISO();
  if(!E.estados.includes('pago')&&E.promesa<h) return ` · <span style="color:var(--accent)">vencida ${fechaCorta(E.promesa)}</span>`;
  if(E.promesa===h) return ' · hoy'; return ' · '+fechaCorta(E.promesa);}
function render(){
  const P=periodoActual(); const sel=$('#selPeriodo');
  const ps=Object.values(db.periodos).sort((a,b)=>fechaNum(b.fin)-fechaNum(a.fin));
  sel.style.display=ps.length?'':'none';
  sel.innerHTML=ps.map(p=>`<option value="${esc(p.key)}" ${P&&p.key===P.key?'selected':''}>Corte ${esc(plabel(p))}</option>`).join('');
  $('#filters').style.display=P?'':'none'; $('#btnExport').style.display=P?'':'none';
  if(!P){ $('#progress').innerHTML=''; $('#list').innerHTML=`<div class="empty"><div class="big">📄</div><h2>Carga tu primer listado</h2><p>Toca <b>Cargar listados</b> y elige uno o varios PDFs de cortados. Los encuentras en <b>Documentos → WhatsApp</b> o en Descargas.</p><p class="muted">Todo se lee y se guarda solo en este celular.</p></div>`; return;}
  const todos=P.orden.map(s=>P.clientes[s]); const tot=todos.length;
  const cnt={}; GRUPOS.forEach(g=>cnt[g.k]=0); todos.forEach(c=>cnt[grupoDe(c)]++);
  const trab=tot-cnt.none, u=db.ui;
  $('#progress').innerHTML=`<div class="pnums"><span>${tot} clientes · <b>${trab} trabajados</b></span><span>${tot?Math.round(trab/tot*100):0}%</span></div>
   <div class="bar">${GRUPOS.filter(g=>g.k!=='none').map(g=>`<span style="width:${cnt[g.k]/tot*100}%;background:${g.c}"></span>`).join('')}</div>
   <div class="counts">${GRUPOS.map(g=>`<button class="cnt ${u.estado===g.k?'on':''}" data-est="${g.k}"><i style="background:${g.c}"></i>${cnt[g.k]}<span style="font-weight:500">${g.t}</span></button>`).join('')}</div>`;
  // filtros
  const pagos=[...new Set(todos.map(c=>c.ultimo_pago))].sort((a,b)=>fechaNum(a)-fechaNum(b));
  $('#selPago').innerHTML=`<option value="">Último pago: todos</option>`+pagos.map(p=>`<option ${u.pago===p?'selected':''} value="${esc(p)}">Últ. pago ${esc(p)} (${todos.filter(c=>c.ultimo_pago===p).length})</option>`).join('');
  $('#selOrden').value=u.orden; if($('#q').value!==u.q) $('#q').value=u.q;
  const cols=[...new Set(todos.map(c=>c.colonia))].sort((a,b)=>String(a).localeCompare(String(b),'es',{numeric:true}));
  const promHoyN=todos.filter(promesaPendiente).length;
  $('#chipsCol').innerHTML=`<button class="chip ${u.promHoy?'on':''}" data-promhoy>📅 Promesas hoy y vencidas (${promHoyN})</button><button class="chip ${!u.col?'on':''}" data-col="">Todas las colonias</button>`+
    cols.map(c=>`<button class="chip ${u.col===c?'on':''}" data-col="${esc(c)}">Col. ${esc(c)} (${todos.filter(x=>x.colonia===c).length})</button>`).join('');
  vista=filtrados(P);
  $('#list').innerHTML=vista.length?vista.map((c,i)=>{
    const p=principal(c), otros=otrosPeriodos(c.suscriptor,P.key);
    const pr=Object.entries(c.promos).map(([k,v])=>`<span class="pill">${k} · ${fmt(v)}</span>`).join(' ');
    return `<button class="card" data-i="${i}" style="--sc:${colorDe(p)}">
      <div class="nm">${esc(c.nombre)} ${esc(c.apellido)} ${otros.length?`<span class="rep">🔁 ${esc(otros.join(', '))}</span>`:''}</div>
      <div class="meta">📋 ${c.contrato} · Col. ${esc(c.colonia)}</div>
      <div class="meta">Adeudo <b style="color:var(--text)">${fmt(c.adeudo)}</b> · Últ. pago ${esc(c.ultimo_pago)}</div>
      <div class="row2"><div style="display:flex;gap:4px;flex-wrap:wrap">${pr||'<span class="muted" style="font-size:12px">Sin promociones</span>'}</div>
      <span class="state-tag">${p?esc(nombreDe(p)):'⚪ Sin contactar'}${p==='promesa'?etiquetaPromesa(c):''}</span></div>
    </button>`;}).join(''):`<div class="empty"><div class="big">🔎</div><p>No hay clientes con estos filtros.</p></div>`;
}
const fechaCorta=iso=>{const m=iso.match(/(\d{4})-(\d{2})-(\d{2})/); return m?`${m[3]}/${m[2]}`:iso;};

// ===== Vista detalle =====
let idx=-1;
function abrir(i){idx=i; $('#sheet').classList.add('open'); document.body.style.overflow='hidden'; renderDetalle(); $('#sheet').scrollTop=0; abrirCapa(ocultarDetalle);}
function ocultarDetalle(){$('#sheet').classList.remove('open'); document.body.style.overflow=''; render();}
function cerrar(){ if($('#sheet').classList.contains('open')) cerrarCapa(); }
function registrar(c,accion,extra){est(c).hist.unshift({t:ahora(),iso:new Date().toISOString(),accion,...(extra||{})}); save();}
function marcar(c,k,on){const e=est(c).estados; const has=e.includes(k); if(on===undefined) on=!has; if(on&&!has){e.push(k); registrar(c,'Estado: '+nombreDe(k)+(k==='promesa'&&est(c).promesa?' ('+fechaCorta(est(c).promesa)+')':''));} if(!on&&has){e.splice(e.indexOf(k),1); registrar(c,'Quitó: '+nombreDe(k)); if(k==='promesa') est(c).promesa='';} save();}
function renderDetalle(){
  const P=periodoActual(); const c=vista[idx]; if(!c){cerrar();return;}
  const E=est(c), p=principal(c), nums=numeros(c), msg=mensaje(c), otros=otrosPeriodos(c.suscriptor,P.key);
  const promos=Object.entries(c.promos).sort((a,b)=>['1969','1147','1396','989'].indexOf(a[0])-['1969','1147','1396','989'].indexOf(b[0]));
  $('#detail').innerHTML=`
  <div class="dhead"><button class="nav" data-a="volver">← Volver</button><span class="pos">${idx+1} / ${vista.length}</span>
   <button class="nav" data-a="prev" ${idx<=0?'disabled style="opacity:.4"':''}>◀</button><button class="nav" data-a="next" ${idx>=vista.length-1?'disabled style="opacity:.4"':''}>▶</button></div>
  <div class="sec hero" style="--sc:${colorDe(p)}">
    <div class="nm">${esc(c.nombre)} ${esc(c.apellido)}</div>
    <div class="muted" style="font-size:12px">${esc(c.nombre_raw)} · Susc. ${esc(c.suscriptor)}</div>
    <div class="contract"><code>${c.contrato}</code><button class="mini" data-a="copiar">📋 Copiar</button></div>
    <div style="margin-top:8px;font-weight:700;color:${colorDe(p)}">${p?esc(nombreDe(p)):'⚪ Sin contactar'}${p==='promesa'?etiquetaPromesa(c):''}</div>
    ${otros.length?`<div style="margin-top:6px"><span class="rep">🔁 También en corte ${esc(otros.join(', '))}</span></div>`:''}
  </div>
  <div class="sec">
    <dl class="kv">
      <dt>Adeudo</dt><dd class="big">${fmt(c.adeudo)}</dd>
      <dt>Último pago</dt><dd>${esc(c.ultimo_pago)}</dd>
      <dt>Colonia</dt><dd>${esc(c.colonia)}</dd>
      <dt>Paquete</dt><dd style="font-weight:500;font-size:13px">${esc(c.paquete)||'—'}</dd>
      <dt>Dirección</dt><dd style="font-weight:500;font-size:13px">📍 ${esc(c.direccion)||'—'}</dd>
      ${c.referencia?`<dt>Referencia</dt><dd style="font-weight:500;font-size:13px">${esc(c.referencia)}</dd>`:''}
    </dl>
    ${c.cols_ok===false?`<div class="warnbox">⚠️ Los montos de este cliente no se pudieron leer con seguridad. Revísalos en el PDF (hoja ${esc(c.hoja)}) antes de ofrecer una promoción.</div>`:''}
  </div>
  <div class="sec"><h3>Promociones</h3>
    ${promos.length?`<div class="promos">${promos.map(([k,v])=>`<button class="promo ${E.promo===k?'on':''}" data-promo="${k}"><b>${k}</b><span>${fmt(v)}</span></button>`).join('')}</div>`
      :(c.adeudo!=null?`<p class="muted" style="margin-top:0">Este cliente no trae promociones en el listado. Puedes ofrecerle ponerse al corriente con su adeudo:</p><div class="promos"><button class="promo ${E.promo==='adeudo'?'on':''}" data-promo="adeudo"><b>Adeudo total</b><span>${fmt(c.adeudo)}</span></button></div>`:'<p class="muted">Este cliente no trae promociones en el listado.</p>')}
  </div>
  <div class="sec"><h3>💬 Mensaje de WhatsApp</h3>
    <div class="msg ${msg?'':'off'}">${msg?esc(msg):'Elige una promoción para armar el mensaje.'}</div>
    ${msg?'<button class="mini" data-a="copiarmsg" style="margin-top:8px">📋 Copiar mensaje</button>':''}
    <h3 style="margin-top:14px">Enviar a</h3>
    <div class="nums">${nums.length?nums.map(n=>`<a class="numbtn wa ${msg?'':'dis'}" data-wa="${n.num}" href="https://wa.me/52${n.num}?text=${msg?encodeURIComponent(msg):''}" target="_blank" rel="noopener">💬 ${n.num}<small>${esc(n.labs.join(' · '))}</small></a>`).join(''):'<p class="muted">Sin números válidos.</p>'}</div>
  </div>
  <div class="sec"><h3>📞 Llamar</h3>
    <div class="nums">${nums.length?nums.map(n=>`<a class="numbtn call" data-call="${n.num}" href="tel:${n.num}"${enMarco?' target="_blank"':''}>📞 ${n.num}<small>${esc(n.labs.join(' · '))}</small></a>`).join(''):'<p class="muted">Sin números válidos.</p>'}</div>
  </div>
  <div class="sec"><h3>Resultado</h3>
    <div class="states">${ESTADOS.map(s=>`<button class="st ${E.estados.includes(s.k)?'on':''}" style="--c:${s.c}" data-s="${s.k}">${s.t}${s.k==='promesa'&&E.promesa&&E.estados.includes('promesa')?' · '+fechaCorta(E.promesa):''}</button>`).join('')}</div>
    <h3 style="margin-top:14px">📝 Notas</h3>
    <textarea id="notas" placeholder="Ej. Dice que paga el viernes, hablarle en la tarde">${esc(E.notas)}</textarea>
  </div>
  <div class="sec"><h3>🕒 Historial</h3>
    ${E.hist.length?`<ul class="hist">${E.hist.map(h=>`<li><time>${esc(h.t)}</time><span>${esc(h.accion)}${h.num?' · '+esc(h.num):''}${h.promo?' · promo '+esc(nombrePromo(h.promo)):''}</span></li>`).join('')}</ul>`:'<p class="muted">Aún no hay acciones con este cliente.</p>'}
  </div>`;
  const ta=$('#notas'); ta.oninput=()=>{E.notas=ta.value; clearTimeout(ta._t); ta._t=setTimeout(save,400);};
  ta.onchange=()=>{ if(E.notas.trim()) registrar(c,'Nota actualizada'); };
}
$('#detail').addEventListener('click',e=>{
  const c=vista[idx]; if(!c) return; const E=est(c); const t=e.target.closest('[data-a],[data-promo],[data-wa],[data-call],[data-s]'); if(!t) return;
  if(t.dataset.a==='volver') return cerrar();
  if(t.dataset.a==='copiarmsg'){ const m=mensaje(c); if(m) copiar(m,'📋 Mensaje copiado'); return; }
  if(t.dataset.a==='prev'&&idx>0){idx--; renderDetalle(); $('#sheet').scrollTop=0; return;}
  if(t.dataset.a==='next'&&idx<vista.length-1){idx++; renderDetalle(); $('#sheet').scrollTop=0; return;}
  if(t.dataset.a==='copiar'){copiar(c.contrato,'📋 Contrato copiado: '+c.contrato); return;}
  if(t.dataset.promo){E.promo=E.promo===t.dataset.promo?null:t.dataset.promo; save(); renderDetalle(); return;}
  if(t.dataset.wa){ if(!mensaje(c)){e.preventDefault(); return;} registrar(c,'WhatsApp enviado',{num:t.dataset.wa,promo:E.promo}); if(!E.estados.includes('whatsapp')) E.estados.push('whatsapp'); save(); setTimeout(renderDetalle,300); return;}
  if(t.dataset.call){ registrar(c,'Llamada',{num:t.dataset.call}); if(!E.estados.includes('llame')) E.estados.push('llame'); save(); setTimeout(renderDetalle,300); return;}
  if(t.dataset.s){ const k=t.dataset.s;
    if(k==='promesa') return pedirPromesa(c);
    marcar(c,k); renderDetalle(); }
});
function pedirPromesa(c){
  const E=est(c), activa=E.estados.includes('promesa'); const d=new Date(); d.setDate(d.getDate()+1);
  const def=E.promesa||(d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0'));
  abrirModal(`<h2>🟡 Promesa de pago</h2><label class="muted">${activa?'Cambia la fecha o quita la promesa':'¿Qué día promete pagar?'}</label><input type="date" id="fProm" value="${esc(def)}">
   <div style="display:flex;gap:8px">${activa?'<button class="btn sec" id="pQuitar">Quitar promesa</button>':'<button class="btn sec" id="pCancel">Cancelar</button>'}<button class="btn" id="pOk">Guardar</button></div>`);
  if($('#pCancel')) $('#pCancel').onclick=cerrarModal;
  if($('#pQuitar')) $('#pQuitar').onclick=()=>{marcar(c,'promesa',false); cerrarModal(); renderDetalle();};
  $('#pOk').onclick=()=>{ const nueva=$('#fProm').value; if(!nueva){toast('Elige una fecha'); return;}
    if(activa){ if(nueva!==E.promesa){E.promesa=nueva; registrar(c,'Promesa cambiada al '+fechaCorta(nueva));} }
    else { E.promesa=nueva; marcar(c,'promesa',true); }
    save(); cerrarModal(); renderDetalle(); };
}
async function copiar(t,aviso){try{await navigator.clipboard.writeText(t); toast(aviso);}catch(e){const i=document.createElement('textarea'); i.value=t; document.body.appendChild(i); i.select(); try{document.execCommand('copy'); toast(aviso);}catch(_){toast('No se pudo copiar');} i.remove();}}
let bloqueado=false;
function abrirModal(html,opt){ $('#mbox').innerHTML=html; bloqueado=!!(opt&&opt.bloqueado); if(!$('#modal').classList.contains('open')){ $('#modal').classList.add('open'); abrirCapa(ocultarModal); } }
function ocultarModal(){ $('#modal').classList.remove('open'); bloqueado=false; }
function cerrarModal(){ if($('#modal').classList.contains('open')) cerrarCapa(); }
$('#modal').addEventListener('click',e=>{if(e.target.id==='modal'&&!bloqueado) cerrarModal();});

// ===== Mis listados =====
function misListados(){
  const ps=Object.values(db.periodos).sort((a,b)=>fechaNum(b.fin)-fechaNum(a.fin));
  abrirModal(`<h2>Mis listados</h2>${ps.length?ps.map(p=>{const n=p.orden.length, t=p.orden.filter(s=>principal(p.clientes[s])).length;
    return `<div class="lst"><div><b>Corte ${esc(plabel(p))}</b><br><span class="muted" style="font-size:13px">${n} clientes · ${t} trabajados · ${esc(p.archivos.join(', '))}</span></div><button class="mini" data-del="${esc(p.key)}">🗑️</button></div>`;}).join(''):'<p class="muted">Aún no has cargado listados.</p>'}
   <p class="muted" style="font-size:13px">Tu avance se guarda solo en este celular. Exporta a Excel de vez en cuando como respaldo.</p>
   <button class="btn sec" onclick="cerrarModal()">Cerrar</button>`);
  $('#mbox').querySelectorAll('[data-del]').forEach(b=>b.onclick=()=>{const p=db.periodos[b.dataset.del]; if(confirm(`¿Eliminar el corte ${plabel(p)} y todo su avance? Esto no se puede deshacer.`)){delete db.periodos[b.dataset.del]; save(); render(); misListados();}});
}

// ===== Exportar Excel =====
function exportar(){
  const P=periodoActual(); if(!P) return; if(!window.XLSX){toast('No cargó la librería de Excel, revisa tu internet'); return;}
  const per=plabel(P);
  const filas=P.orden.map(s=>{const c=P.clientes[s], E=est(c), p=principal(c);
    return {'Periodo de corte':per,'Contrato':c.contrato,'Suscriptor':c.suscriptor,'Nombre':c.nombre,'Apellido':c.apellido,'Nombre en listado':c.nombre_raw,'Colonia':c.colonia,'Dirección':c.direccion,'Referencia':c.referencia||'','Paquete':c.paquete,
      'Último pago':c.ultimo_pago,'Adeudo':c.adeudo,'1969':c.promos['1969']??'','1147':c.promos['1147']??'','1396':c.promos['1396']??'','989':c.promos['989']??'',
      'Celular':c.celular,'Megafon':c.megafon,'Ult cel':c.ult_tel,'WhatsApp':c.whatsapp,'Web Cel':c.web_cel,'Tel. oficina':c.tel_oficina,'Casa':c.casa||'','Otros tel.':c.otros_tel||'','Serv. móviles':c.serv_moviles,
      'Estado final':p?nombreDe(p):'Sin contactar','Todos los estados':E.estados.map(nombreDe).join(', '),'Promo ofrecida':E.promo?nombrePromo(E.promo):'','Monto promo':E.promo?(montoPromo(c,E.promo)??''):'','Fecha de promesa':E.estados.includes('promesa')&&E.promesa?E.promesa.split('-').reverse().join('/'):'',
      'Notas':E.notas,'Último contacto':E.hist[0]?new Date(E.hist[0].iso).toLocaleString('es-MX',{hour12:false}):'','Acciones':E.hist.length,'También en corte':otrosPeriodos(s,P.key).join(', ')};});
  const hist=[]; P.orden.forEach(s=>{const c=P.clientes[s]; est(c).hist.forEach(h=>hist.push({'Fecha y hora':new Date(h.iso).toLocaleString('es-MX',{hour12:false}),'Contrato':c.contrato,'Nombre':(c.nombre+' '+c.apellido).trim(),'Colonia':c.colonia,'Acción':h.accion,'Número usado':h.num||'','Promo':h.promo?nombrePromo(h.promo):'',_o:h.iso}));});
  hist.sort((a,b)=>a._o<b._o?1:-1); hist.forEach(h=>delete h._o);
  const wb=XLSX.utils.book_new();
  const H=['Periodo de corte','Contrato','Suscriptor','Nombre','Apellido','Nombre en listado','Colonia','Dirección','Referencia','Paquete','Último pago','Adeudo','1969','1147','1396','989','Celular','Megafon','Ult cel','WhatsApp','Web Cel','Tel. oficina','Casa','Otros tel.','Serv. móviles','Estado final','Todos los estados','Promo ofrecida','Monto promo','Fecha de promesa','Notas','Último contacto','Acciones','También en corte'];
  const ws1=XLSX.utils.aoa_to_sheet([H,...filas.map(f=>H.map(h=>f[h]??''))]); ws1['!cols']=H.map(k=>({wch:Math.min(40,Math.max(10,k.length+2))}));
  const ws2=XLSX.utils.json_to_sheet(hist.length?hist:[{'Fecha y hora':'','Contrato':'','Nombre':'','Colonia':'','Acción':'Sin acciones aún','Número usado':'','Promo':''}]); ws2['!cols']=[{wch:20},{wch:13},{wch:22},{wch:10},{wch:34},{wch:13},{wch:12}];
  XLSX.utils.book_append_sheet(wb,ws1,'Clientes'); XLSX.utils.book_append_sheet(wb,ws2,'Historial');
  XLSX.writeFile(wb,`cortados_${P.ini}_a_${P.fin}.xlsx`); toast('📊 Excel descargado');
}

// ===== Eventos de la lista =====
$('#btnCargar').onclick=()=>$('#file').click();
$('#file').onchange=e=>{const f=[...e.target.files]; e.target.value=''; if(f.length&&!cargando) cargarArchivos(f);};
$('#btnListados').onclick=misListados;
$('#btnExport').onclick=exportar;
$('#selPeriodo').onchange=e=>{db.ui.periodo=e.target.value; db.ui.col='';db.ui.pago='';db.ui.estado='';db.ui.promHoy=false; save(); render(); scrollTo(0,0);};
$('#selPago').onchange=e=>{db.ui.pago=e.target.value; save(); render();};
$('#selOrden').onchange=e=>{db.ui.orden=e.target.value; save(); render();};
$('#q').oninput=e=>{db.ui.q=e.target.value; clearTimeout(window._qt); window._qt=setTimeout(()=>{save(); render();},200);};
$('#chipsCol').onclick=e=>{const b=e.target.closest('button'); if(!b) return; if('promhoy' in b.dataset) db.ui.promHoy=!db.ui.promHoy; else db.ui.col=b.dataset.col; save(); render();};
$('#progress').onclick=e=>{const b=e.target.closest('[data-est]'); if(!b) return; db.ui.estado=db.ui.estado===b.dataset.est?'':b.dataset.est; save(); render();};
$('#list').onclick=e=>{const b=e.target.closest('.card'); if(b) abrir(+b.dataset.i);};
load(); render();
</script></body></html>
