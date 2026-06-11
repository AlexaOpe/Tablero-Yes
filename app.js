/* =========================================================================
   Yes Visas · Tablero de Ingresos
   Lee la tabla BaseDeDatos (SOLO LECTURA). No escribe nada en el Excel.
   Fuente de datos (en orden):
     1) window.__FACT__         -> doble clic / file:// (data/fact.js)
     2) fetch('data/fact.json') -> servidor local / SharePoint
     3) Office.js (futuro add-in) -> reemplaza loadData() leyendo la tabla viva
   ========================================================================= */

/* Paleta tomada del logo (molinillo): verde, naranja, amarillo, lima, rojo... */
const BRAND = {
  azul:'#3E6E23',   // verde hoja (base)
  azul2:'#78C030',  // verde lima
  naranja:'#F09018',// naranja del logo
  naranja2:'#F2A83D',
  naranja3:'#E0531C',// rojo-naranja
  dorado:'#E9CE2A', // amarillo
  dorado2:'#CBB028',// dorado
  vino:'#9C3415',   // terracota
};
const SERIE = [BRAND.azul, BRAND.naranja, BRAND.dorado, BRAND.azul2,
               BRAND.naranja3, BRAND.dorado2, BRAND.vino, BRAND.naranja2];

const MESES = ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'];

const fmtMXN  = new Intl.NumberFormat('es-MX',{style:'currency',currency:'MXN',maximumFractionDigits:0});
const fmtMXN2 = new Intl.NumberFormat('es-MX',{style:'currency',currency:'MXN',maximumFractionDigits:2});
const fmtUSD  = new Intl.NumberFormat('es-MX',{style:'currency',currency:'USD',maximumFractionDigits:0});
const fmtNum  = new Intl.NumberFormat('es-MX');

const state = { rows:[], anio:null, depto:'TODOS', oficina:'TODOS', source:'', inExcel:false };
const charts = {};

/* ---------- Carga de datos (modo archivo / fetch) ---------- */
async function loadData(){
  if (Array.isArray(window.__FACT__)) { state.source = 'fact.js (local)'; return window.__FACT__; }
  try{
    const r = await fetch('data/fact.json', {cache:'no-store'});
    if(!r.ok) throw new Error('HTTP '+r.status);
    state.source = 'fact.json (fetch)';
    return await r.json();
  }catch(e){
    throw new Error('No se pudo cargar la fuente de datos. '+
      'Abre con el archivo data/fact.js presente, o sirve la carpeta por HTTP.');
  }
}

/* ---------- Carga de datos EN VIVO (add-in dentro de Excel, SOLO LECTURA) ----------
   Lee la tabla/hoja BaseDeDatos con Office.js. No escribe nada en el libro:
   solo getHeaderRowRange / getDataBodyRange / getUsedRange + load + sync. */
const COLMAP = {
  'Mes':'mes','Fecha':'fecha','Oficina':'oficina','Cliente':'cliente',
  'Departamento':'departamento','Servicio':'servicio',
  'Descripción del servicio':'descripcion','TotalMXN':'totalMXN','Llave':'llave',
  'TipoTC':'tipoTC','TotalUSD':'totalUSD','Origen':'origen',
  'MontoMXN':'montoMXN','MontoUSD':'montoUSD','Grupo':'grupo'
};
function excelSerialToISO(n){
  // Excel guarda fechas como nº de serie (días desde 1899-12-30; 25569 = 1970-01-01)
  const d = new Date(Math.round((n - 25569) * 86400000));
  return isNaN(d) ? '' : d.toISOString().slice(0,10);
}
function normFecha(val, txt){
  if(typeof val === 'number' && isFinite(val)) return excelSerialToISO(val);
  const s = String(txt ?? val ?? '').trim();
  const m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/); // dd/mm/yyyy -> yyyy-mm-dd
  if(m){ let [,dd,mo,y]=m; if(y.length===2) y='20'+y;
    return `${y}-${mo.padStart(2,'0')}-${dd.padStart(2,'0')}`; }
  return s.slice(0,10);
}
async function loadDataLive(){
  return await Excel.run(async (ctx)=>{
    let headerRow, bodyVals, bodyTxt;
    // 1) Preferir la Tabla (ListObject) llamada BaseDeDatos
    const tables = ctx.workbook.tables; tables.load('items/name');
    await ctx.sync();
    const tbl = tables.items.find(t=>t.name==='BaseDeDatos');
    if(tbl){
      const hdr  = tbl.getHeaderRowRange().load('values');
      const body = tbl.getDataBodyRange().load('values, text');
      await ctx.sync();
      headerRow = hdr.values[0]; bodyVals = body.values; bodyTxt = body.text;
    }else{
      // 2) Caer a la hoja BaseDeDatos (encabezado en fila 1)
      const used = ctx.workbook.worksheets.getItem('BaseDeDatos').getUsedRange().load('values, text');
      await ctx.sync();
      headerRow = used.values[0];
      bodyVals = used.values.slice(1); bodyTxt = used.text.slice(1);
    }

    const headers = headerRow.map(h=>String(h||'').trim());
    const keys = headers.map(h=>COLMAP[h]||null);
    const faltan = ['fecha','oficina','departamento','totalMXN','grupo'].filter(k=>!keys.includes(k));
    if(faltan.length) throw new Error('CONTRATO ROTO: BaseDeDatos no tiene las columnas: '+faltan.join(', '));

    const rows=[];
    for(let i=0;i<bodyVals.length;i++){
      const o={};
      for(let c=0;c<keys.length;c++){
        const k=keys[c]; if(!k) continue;
        o[k] = (k==='fecha') ? normFecha(bodyVals[i][c], bodyTxt[i][c]) : bodyVals[i][c];
      }
      if(o.fecha || o.cliente || num(o.totalMXN)) rows.push(o); // descarta filas vacías
    }
    state.source = 'BaseDeDatos (Excel, en vivo) · solo lectura';
    return rows;
  });
}

/* ---------- Utilidades ---------- */
const yearOf  = r => (r.fecha||'').slice(0,4);
const monthOf = r => parseInt((r.fecha||'0000-00').slice(5,7),10) - 1; // 0..11
const isBlank = v => v===null || v===undefined || String(v).trim()==='';
const num = v => (typeof v==='number' && isFinite(v)) ? v : 0;

function sumBy(rows, keyFn, valFn){
  const m = new Map();
  for(const r of rows){ const k=keyFn(r); m.set(k,(m.get(k)||0)+valFn(r)); }
  return m;
}

/* ---------- Banderas de validación (quality-at-source) ----------
   Solo "servicios sin catalogar" (Grupo "Otros"), por decisión del usuario. */
function computeFlags(rows){
  const sinCatalogar=[];
  const flaggedRows = new Set();
  rows.forEach((r,i)=>{
    if((r.grupo||'')==='Otros'){ sinCatalogar.push(i); flaggedRows.add(i); }
  });

  const ex = idxs => idxs.slice(0,5).map(i=>{
    const r=rows[i];
    return `${(r.fecha||'').slice(0,10)} · ${r.cliente||'—'} · ${r.servicio||'—'} · ${fmtMXN.format(num(r.totalMXN))}`;
  });

  return {
    flaggedRows,
    cats:[
      {key:'otros', sev:'media', ico:'🏷️', title:'Servicios sin catalogar (Grupo "Otros")',
       n:sinCatalogar.length, ej:ex(sinCatalogar),
       desc:'Servicios que no entran en un grupo definido — revisar con la dueña para clasificar.'},
    ]
  };
}

/* ---------- Render: KPIs ---------- */
function deltaBadge(curr, prev){
  if(prev===0 || prev===null){ return `<span class="delta flat">s/d</span>`; }
  const d=(curr-prev)/Math.abs(prev)*100;
  const cls = d>0.5?'up':(d<-0.5?'down':'flat');
  const sign = d>0?'▲':(d<0?'▼':'■');
  return `<span class="delta ${cls}">${sign} ${Math.abs(d).toFixed(1)}%</span>`;
}
function renderKPIs(rowsAnio, allRows, matchDepto){
  const anio=state.anio, prev=String(+anio-1);
  const totalAnio = rowsAnio.reduce((s,r)=>s+num(r.totalMXN),0);
  const totalPrev = allRows.filter(r=>yearOf(r)===prev && matchDepto(r)).reduce((s,r)=>s+num(r.totalMXN),0);

  // mes de referencia = último mes con datos del año seleccionado
  const mesesConDatos = [...new Set(rowsAnio.map(monthOf))].sort((a,b)=>a-b);
  const mRef = mesesConDatos.length?mesesConDatos[mesesConDatos.length-1]:0;
  const rowsMes = rowsAnio.filter(r=>monthOf(r)===mRef);
  const totalMes = rowsMes.reduce((s,r)=>s+num(r.totalMXN),0);
  const rowsMesPrev = mRef>0 ? rowsAnio.filter(r=>monthOf(r)===mRef-1) : [];
  const totalMesPrev = rowsMesPrev.reduce((s,r)=>s+num(r.totalMXN),0);

  // TC vigente = TC de la transacción más reciente del año
  const ult = rowsAnio.slice().sort((a,b)=>(b.fecha||'').localeCompare(a.fecha||''))[0];
  const tcVig = ult?num(ult.tipoTC):0;

  const html = `
    <div class="kpi">
      <span class="k-label">Ingreso del año ${anio}</span>
      <span class="k-value">${fmtMXN.format(totalAnio)}</span>
      <span class="k-sub">vs ${prev}: ${deltaBadge(totalAnio,totalPrev)} <span style="opacity:.7">${fmtMXN.format(totalPrev)}</span></span>
    </div>
    <div class="kpi k-naranja">
      <span class="k-label">Ingreso ${MESES[mRef]} ${anio}</span>
      <span class="k-value">${fmtMXN.format(totalMes)}</span>
      <span class="k-sub">vs ${MESES[(mRef+11)%12]}: ${deltaBadge(totalMes,totalMesPrev)}</span>
    </div>
    <div class="kpi k-dorado">
      <span class="k-label">Tipo de cambio vigente</span>
      <span class="k-value">${tcVig>0?fmtMXN2.format(tcVig).replace('$',''):'—'}</span>
      <span class="k-sub">última transacción ${ult?('· '+(ult.fecha||'').slice(0,10)):''}</span>
    </div>
    <div class="kpi k-azul2">
      <span class="k-label"># Transacciones ${MESES[mRef]}</span>
      <span class="k-value">${fmtNum.format(rowsMes.length)}</span>
      <span class="k-sub">${fmtNum.format(rowsAnio.length)} en el año</span>
    </div>`;
  document.getElementById('kpis').innerHTML = html;
}

/* ---------- Render: gráficas ---------- */
function destroyChart(id){ if(charts[id]){ charts[id].destroy(); delete charts[id]; } }

function renderDepto(rowsAnio){
  // Siempre muestra TODOS los departamentos (es el desglose). Clic = filtra.
  const m = sumBy(rowsAnio, r=>r.departamento||'(sin depto)', r=>num(r.totalMXN));
  const labels=[...m.keys()], data=[...m.values()];
  const total=data.reduce((a,b)=>a+b,0)||1;
  // resalta el seleccionado; atenúa los demás
  const colors=labels.map((l,i)=>{
    const base=SERIE[i%SERIE.length];
    if(state.depto==='TODOS' || state.depto===l) return base;
    return base+'55'; // semitransparente
  });
  destroyChart('depto');
  charts.depto = new Chart(document.getElementById('chartDepto'),{
    type:'doughnut',
    data:{labels, datasets:[{data, backgroundColor:colors, borderWidth:2, borderColor:'#fff'}]},
    options:{responsive:true,maintainAspectRatio:false,cutout:'58%',
      onClick:(e,els)=>{
        if(!els.length) return;
        const l=labels[els[0].index];
        setDepto(state.depto===l?'TODOS':l);
      },
      plugins:{legend:{display:false},
        tooltip:{callbacks:{label:c=>` ${c.label}: ${fmtMXN.format(c.parsed)} (${(c.parsed/total*100).toFixed(1)}%)`}}}}
  });

  // Lista de montos por departamento (visible de inmediato, sin hover). Clic = filtra.
  const leg=document.getElementById('deptoLegend');
  const pares=labels.map((l,i)=>({l,v:data[i],c:SERIE[i%SERIE.length]}))
    .sort((a,b)=>b.v-a.v);
  leg.innerHTML=pares.map(p=>`
    <div class="dl-row ${state.depto===p.l?'sel':''}" data-dep="${p.l}">
      <span class="dl-dot" style="background:${p.c}"></span>
      <span class="dl-name">${p.l}</span>
      <span class="dl-amt">${fmtMXN.format(p.v)}</span>
      <span class="dl-pct">${(p.v/total*100).toFixed(1)}%</span>
    </div>`).join('');
  leg.querySelectorAll('.dl-row').forEach(row=>{
    row.onclick=()=>{ const d=row.dataset.dep; setDepto(state.depto===d?'TODOS':d); };
  });
}
function renderGrupo(rowsAnio){
  const m=[...sumBy(rowsAnio,r=>r.grupo||'(sin grupo)',r=>num(r.totalMXN))]
    .sort((a,b)=>b[1]-a[1]).slice(0,8);
  destroyChart('grupo');
  charts.grupo=new Chart(document.getElementById('chartGrupo'),{
    type:'bar',
    data:{labels:m.map(x=>x[0]),datasets:[{data:m.map(x=>x[1]),
      backgroundColor:BRAND.naranja,borderRadius:5}]},
    options:{indexAxis:'y',responsive:true,maintainAspectRatio:false,
      plugins:{legend:{display:false},
        tooltip:{callbacks:{label:c=>' '+fmtMXN.format(c.parsed.x)}}},
      scales:{x:{ticks:{callback:v=>'$'+(v/1000).toFixed(0)+'k'}}}}
  });
}
function renderComp(allRows, matchDepto){
  const anio=state.anio, prev=String(+anio-1);
  const arr = y => { const a=new Array(12).fill(0);
    allRows.filter(r=>yearOf(r)===y && matchDepto(r)).forEach(r=>{a[monthOf(r)]+=num(r.totalMXN);}); return a; };
  const cur=arr(anio), pre=arr(prev);
  // variación interanual por mes (monto y %)
  const varMonto = cur.map((v,i)=>v-pre[i]);
  const varPct   = cur.map((v,i)=>pre[i]>0?((v-pre[i])/pre[i]*100):null);
  destroyChart('comp');
  charts.comp=new Chart(document.getElementById('chartComp'),{
    type:'bar',
    data:{labels:MESES,datasets:[
      {label:`Ingreso ${anio}`,data:cur,backgroundColor:BRAND.azul,borderRadius:4,order:2},
      // línea de referencia = ingreso del año previo, pero etiquetada como Variación interanual;
      // el tooltip expresa la variación en % y monto (no el valor del año previo).
      {label:'Variación interanual',type:'line',data:pre,borderColor:BRAND.naranja,
       backgroundColor:BRAND.naranja,borderWidth:2,borderDash:[5,4],tension:.3,pointRadius:3,order:1}
    ]},
    options:{responsive:true,maintainAspectRatio:false,
      plugins:{legend:{position:'top',labels:{boxWidth:12}},
        tooltip:{callbacks:{
          label:c=>{
            if(c.dataset.label==='Variación interanual'){
              const m=varMonto[c.dataIndex], p=varPct[c.dataIndex];
              if(p===null) return ' s/d';
              const sgn=m>=0?'▲ +':'▼ ';
              return ` ${sgn}${p.toFixed(1)}%  ·  ${m>=0?'+':''}${fmtMXN.format(m)}`;
            }
            return ` ${c.dataset.label}: ${fmtMXN.format(c.parsed.y)}`;
          }
        }}},
      scales:{y:{ticks:{callback:v=>'$'+(v/1000).toFixed(0)+'k'}}}}
  });
}

/* ---------- Render: banderas ---------- */
function renderFlags(flags){
  const cont=document.getElementById('flags');
  const total=flags.cats.reduce((s,c)=>s+c.n,0);
  document.getElementById('flagsCount').textContent=fmtNum.format(total);
  if(total===0){
    cont.innerHTML=`<div class="flag ok"><span class="f-ico">✅</span>
      <div class="f-body"><span class="f-title">Sin observaciones</span>
      <span class="f-desc">La captura del año seleccionado no disparó banderas.</span></div></div>`;
    return;
  }
  cont.innerHTML=flags.cats.filter(c=>c.n>0).map(c=>`
    <div class="flag sev-${c.sev}">
      <span class="f-ico">${c.ico}</span>
      <div class="f-body">
        <span class="f-title">${c.title} · ${fmtNum.format(c.n)}</span>
        <span class="f-desc">${c.desc}</span>
        ${c.ej.map(e=>`<span class="f-desc">› ${e}</span>`).join('')}
      </div>
    </div>`).join('');
}

/* ---------- Render: últimas transacciones ---------- */
function renderTable(rowsAnio, flaggedRows, idxOf){
  const tbody=document.querySelector('#tblLast tbody');
  const ult=rowsAnio.map(r=>r).sort((a,b)=>(b.fecha||'').localeCompare(a.fecha||'')).slice(0,20);
  document.getElementById('lastCount').textContent='últimas '+ult.length;
  tbody.innerHTML=ult.map(r=>{
    const flagged = flaggedRows.has(idxOf.get(r));
    return `<tr class="${flagged?'flagged':''}">
      <td>${(r.fecha||'').slice(0,10)}</td>
      <td class="cell-clip" title="${(r.cliente||'').replace(/"/g,'')}">${r.cliente||'—'}</td>
      <td>${r.departamento||'—'}</td>
      <td class="cell-clip" title="${(r.servicio||'').replace(/"/g,'')}">${r.servicio||'—'}</td>
      <td>${r.origen||'—'}</td>
      <td class="num">${fmtMXN.format(num(r.totalMXN))}</td>
      <td class="num">${fmtUSD.format(num(r.totalUSD))}</td>
      <td class="num">${num(r.tipoTC)>0?num(r.tipoTC).toFixed(2):'—'}</td>
    </tr>`;
  }).join('');
}

/* ---------- Orquestación ---------- */
function renderAll(){
  const all=state.rows;
  const matchDepto   = r => state.depto==='TODOS'   || (r.departamento||'')===state.depto;
  const matchOficina = r => state.oficina==='TODOS' || (r.oficina||'')===state.oficina;
  const match        = r => matchDepto(r) && matchOficina(r);

  const rowsAnio   = all.filter(r=>yearOf(r)===state.anio);        // base del año
  const rowsDepto  = rowsAnio.filter(matchOficina);               // doughnut: deptos dentro de la oficina
  const rowsView   = rowsAnio.filter(match);                      // todo lo demás (ambos filtros)
  const idxOf=new Map(); all.forEach((r,i)=>idxOf.set(r,i));

  const flags=computeFlags(rowsView);
  const flaggedByGlobal=new Set([...flags.flaggedRows].map(i=>idxOf.get(rowsView[i])));

  const parts=[]; if(state.depto!=='TODOS') parts.push(state.depto);
  if(state.oficina!=='TODOS') parts.push(state.oficina);
  const suf = parts.length?' · '+parts.join(' · '):'';
  document.getElementById('lblAnioDep').textContent=state.anio +
    (state.oficina==='TODOS'?'':' · '+state.oficina) + (state.depto==='TODOS'?'':' · '+state.depto+' resaltado');
  document.getElementById('lblAnioGrupo').textContent=state.anio+suf;
  document.getElementById('lblComp').textContent=`${state.anio} vs ${+state.anio-1}${suf}`;

  renderKPIs(rowsView, all, match);
  renderDepto(rowsDepto);
  renderGrupo(rowsView);
  renderComp(all, match);
  renderFlags(flags);
  renderTable(rowsView, flaggedByGlobal, idxOf);

  const stamp=new Date();
  document.getElementById('updatedLabel').textContent='Actualizado '+stamp.toLocaleTimeString('es-MX');
  document.getElementById('sourceLabel').textContent=state.source;
  document.getElementById('footMeta').textContent=
    `${fmtNum.format(rowsView.length)} transacciones mostradas${suf} · fuente: BaseDeDatos (solo lectura)`;
}

function buildSelectors(){
  const years=[...new Set(state.rows.map(yearOf))].filter(Boolean).sort();
  const selA=document.getElementById('selAnio');
  selA.innerHTML=years.map(y=>`<option value="${y}">${y}</option>`).join('');
  state.anio=years[years.length-1];
  selA.value=state.anio;
  selA.onchange=()=>{ state.anio=selA.value; renderAll(); };

  const deptos=[...new Set(state.rows.map(r=>r.departamento).filter(Boolean))].sort();
  const selD=document.getElementById('selDepto');
  selD.innerHTML=`<option value="TODOS">Todos</option>`+
    deptos.map(d=>`<option value="${d}">${d}</option>`).join('');
  selD.value=state.depto;
  selD.onchange=()=>{ state.depto=selD.value; renderAll(); };

  // Pestañas por oficina (YES Todo / MID / CDMX)
  document.querySelectorAll('#tabs .tab').forEach(btn=>{
    btn.onclick=()=>setOficina(btn.dataset.ofi);
  });
}

function setOficina(o){
  state.oficina=o;
  document.querySelectorAll('#tabs .tab').forEach(b=>
    b.classList.toggle('active', b.dataset.ofi===o));
  renderAll();
}

// Cambia el filtro de departamento desde cualquier lado (select o clic en doughnut)
function setDepto(d){
  state.depto=d;
  const selD=document.getElementById('selDepto');
  if(selD) selD.value=d;
  renderAll();
}

async function refresh(){
  const btn=document.getElementById('btnRefresh');
  btn.classList.add('spinning');
  try{
    state.rows = state.inExcel ? await loadDataLive() : await loadData();
    if(!state.anio) buildSelectors();
    renderAll();
  }catch(e){
    document.getElementById('kpis').innerHTML=
      `<div class="state-msg err">${e.message}</div>`;
  }finally{
    setTimeout(()=>btn.classList.remove('spinning'),400);
  }
}

document.getElementById('btnRefresh').addEventListener('click',refresh);

/* ---------- Descargar FOTO (HTML autocontenido para compartir por WhatsApp) ----------
   Toma los datos YA leídos (state.rows, los mismos que ve el panel), los inyecta en
   la plantilla self-contained (Chart.js + estilos + logo ya incrustados) y dispara la
   descarga. NO escribe en el Excel: solo genera un archivo local en Descargas. */
function fechaHoyISO(){
  const d=new Date();
  return d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0');
}
async function downloadFoto(){
  const btn=document.getElementById('btnDownload');
  if(!btn) return;
  const original=btn.innerHTML;
  if(!Array.isArray(state.rows) || state.rows.length===0){
    alert('Aún no hay datos cargados. Pulsa "Actualizar" primero.');
    return;
  }
  btn.disabled=true; btn.innerHTML='<span class="ico">⏳</span> Generando…';
  try{
    const r=await fetch('foto-plantilla.html',{cache:'no-store'});
    if(!r.ok) throw new Error('No se pudo leer la plantilla (HTTP '+r.status+')');
    let tpl=await r.text();
    // El marcador va entre comillas en la plantilla; lo cambiamos por el arreglo real.
    const datos=JSON.stringify(state.rows);
    tpl=tpl.replace('"__FACT_PLACEHOLDER__"', datos);
    const blob=new Blob([tpl],{type:'text/html;charset=utf-8'});
    const url=URL.createObjectURL(blob);
    const a=document.createElement('a');
    a.href=url; a.download='Tablero-Yes-'+fechaHoyISO()+'.html';
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(()=>URL.revokeObjectURL(url),2000);
  }catch(e){
    alert('No se pudo generar la foto: '+e.message);
  }finally{
    btn.disabled=false; btn.innerHTML=original;
  }
}
const _btnDl=document.getElementById('btnDownload');
if(_btnDl) _btnDl.addEventListener('click',downloadFoto);

/* Arranque:
   - Dentro de Excel (add-in): espera a que Office esté listo y lee la tabla EN VIVO.
   - Fuera de Excel (navegador / doble clic): arranca directo con fact.js / fetch. */
if (typeof Office !== 'undefined' && Office.onReady) {
  Office.onReady(info => { state.inExcel = !!(info && info.host); refresh(); });
} else {
  refresh();
}
