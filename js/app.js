/* Интерфейс: ввод данных, рендер таблиц, импорт, распознавание через Claude, запуск */
/* ===================== Рендер: таблицы анализа ===================== */
const esc = s => String(s).replace(/[&<>"]/g, ch=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[ch]));
function tableHTML(t){
  const rows=allRows(t), grid=layout(rows,t.ncols);
  let h='<table>';
  rows.forEach((row,r)=>{
    const f=row[0]||{};
    const cls = f.hd?'' : f.sec?'sec' : (f.tt===2?'t2':f.tt===1?'t1':'') + (f.sub?' sub':'');
    h+=`<tr class="${cls}">`;
    for (let c=0;c<t.ncols;c++){
      const s=grid[r][c]; if(!s||s.r0!==r||s.c0!==c) continue;
      const cell=s.cell, span=(cell.cs>1?` colspan="${cell.cs}"`:'')+(cell.rs>1?` rowspan="${cell.rs}"`:'');
      if (cell.hd){ h+=`<th${span}>${esc(cell.t)}</th>`; continue; }
      const isNum = cell.v!==undefined;
      const k = isNum ? ('num'+(cell.v<0&&fmt(cell)!=='0'?' neg':'')+(cell.dev&&cell.v>0&&/[1-9]/.test(fmt(cell))?' pos':'')+(cell.x?' '+cell.x:'')) : (c===0?'name':(cell.cls||''));
      h+=`<td class="${k}"${span}>${esc(fmt(cell))}</td>`;
    }
    h+='</tr>';
  });
  return h+'</table>';
}
let sampleFn = null, downloadsNs = null, imagesOk = false;
// Одна строка с формулами под каждой таблицей первой вкладки (на странице; в Word и Excel не попадает)
const TABLE_FORMULAS = {
  t22:'Абсолютное отклонение = Xᵢ − Xᵢ₋₁; относительное отклонение = Xᵢ / Xᵢ₋₁ × 100%.',
  t23:'Абсолютное отклонение = Xᵢ − Xᵢ₋₁; относительное отклонение = Xᵢ / Xᵢ₋₁ × 100%.',
  t24:'Удельный вес = статья / Баланс (стр. 1600) × 100%; изменение удельного веса = уд. весᵢ − уд. весᵢ₋₁, п. п.',
  t25:'Удельный вес = статья / Баланс (стр. 1700) × 100%; изменение удельного веса = уд. весᵢ − уд. весᵢ₋₁, п. п.',
  t26:'Разница = КЗ (стр. 1520) − ДЗ (стр. 1230); КЗ на 1 руб. ДЗ = КЗ / ДЗ.',
  t27:'ВП = 2110 − 2120; Пп = ВП − 2210 − 2220; Пнобл = Пп + 2310 + 2320 − 2330 + 2340 − 2350; ПЧ = Пнобл − 2410 + 2460; отклонение = Xᵢ − Xᵢ₋₁.',
  t28:'Кабс = (ДС + КФВ) / КО; Ксроч = (ДС + КФВ + ДЗ) / КО; Кмоб = (З + НДС) / КО; Кобщ = (ДС + КФВ + ДЗ + З) / КО.',
  t29:'Кнез = СК / ВБ; Кнапр = ЗК / ВБ; Ксамоф = СК / ЗК; Кзад = ЗК / СК; Кобесп = СОК / ОА; Кманевр = СОК / СК; Кмоб/иммоб = ОА / ВОА; Кипн = (ВОА + З) / ВБ.',
  t210:'СОК = СК − ВОА; СДИ = СОК + ДКЗ; ОИ = СДИ + ККЗ; ΔСОК = СОК − З; ΔСДИ = СДИ − З; ΔОИ = ОИ − З.',
  t211:'А1 = ДС + КФВ; А2 = ДЗ + прочие ОА; А3 = З + НДС + активы к продаже; А4 = ВОА; П1 = КЗ + прочие КО; П2 = ККЗ; П3 = ДО; П4 = СК + ДБП + оценочные обязательства.',
};
function renderOut(){
  const tables = buildTables();
  document.getElementById('toc').innerHTML = tables.map(t=>`<a href="#${t.id}">Таблица ${t.num}</a>`).join('')+'<a href="#s4">Расчеты</a>';
  const out = document.getElementById('out');
  out.innerHTML = tables.map(t=>`
    <figure class="tbl" id="${t.id}">
      <figcaption>Таблица ${t.num} – ${esc(t.title)}</figcaption>
      <div class="scroll">${tableHTML(t)}</div>
      ${TABLE_FORMULAS[t.id]?`<p class="formula"><b>Формулы:</b> ${esc(TABLE_FORMULAS[t.id])}</p>`:''}
      <div class="concl">
        <textarea data-id="${t.id}" placeholder="Вывод по таблице ${t.num} — напишите сами или нажмите «Сгенерировать вывод»">${esc(state.conclusions[t.id]||'')}</textarea>
        <div class="row">
          <button class="btn ghost small" data-auto="${t.id}">Сгенерировать вывод</button>
          <button class="btn ghost small ai-only" data-ai="${t.id}" ${sampleFn?'':'hidden'}>Черновик вывода от Claude</button>
          <small data-aist="${t.id}"></small>
        </div>
      </div>
    </figure>${chartFiguresHTML(t.id)}`).join('');
  document.getElementById('calc').innerHTML = calcHTML();
  renderKpis();
  renderCtx();
  document.querySelectorAll("#out textarea").forEach(autosize);
  drawCharts();
  renderFR();
}
// Единая шапка: компания и выбранные годы (видна на обеих вкладках)
function renderCtx(){
  document.getElementById('ctx').textContent = (state.company||'Организация не указана')+' · '+state.years[0]+'–'+state.years[nYears()-1]+' гг.';
}

/* ===================== Вкладки: «Анализ отчетности» / «Анализ финансовых результатов» ===================== */
function showPane(p){
  if (p!=='fr') p='main';
  document.querySelectorAll('[data-pane]').forEach(el=>{
    if (el.closest('#mainTabs')) el.setAttribute('aria-selected', el.dataset.pane===p);
    else el.hidden = el.dataset.pane!==p;
  });
  try{ localStorage.setItem('fua-pane', p); }catch(e){}
  if (p==='main') drawCharts();   // графики Chart.js не рисуются в скрытом блоке — перерисовываем при возврате
}
document.getElementById('mainTabs').addEventListener('click',e=>{
  const b=e.target.closest('button[data-pane]'); if(!b) return;
  showPane(b.dataset.pane); window.scrollTo({top:0});
});

/* ===================== Ключевые показатели в шапке ===================== */
// Компактная сумма: тыс. руб. → «615,4 млрд руб.» / «12,3 млн руб.»
const money = x => Math.abs(x)>=1e6 ? fmtNum(x/1e6,1)+' млрд руб.' : Math.abs(x)>=1e3 ? fmtNum(x/1e3,1)+' млн руб.' : fmtNum(x,0)+' тыс. руб.';
function renderKpis(){
  const box=document.getElementById('kpis');
  const lastIdx = stmt => { const I=state.years.map((_,i)=>i).filter(i=>has(stmt,i)); return [I.at(-1), I.at(-2)]; };
  // изменение к предыдущему году: «▲ 17,1% к 2024»
  const delta = (code,z,p) => { if (p==null) return ''; const a=nz(v(code,p)), b=nz(v(code,z)); if (!a) return '';
    const d=(b-a)/Math.abs(a)*100; return `<div class="d ${d>=0?'up':'down'}">${d>=0?'▲':'▼'} ${fmtNum(Math.abs(d),1)}% к ${Y(p)}</div>`; };
  const card = (k,v,d='',cls='') => `<div class="kpi"><div class="k">${esc(k)}</div><div class="v ${cls}">${v}</div>${d}</div>`;
  const [zb,pb]=lastIdx('bal'), [zp,pp]=lastIdx('pl'), parts=[];
  if (zb!=null) parts.push(card(`Валюта баланса на 31.12.${Y(zb)}`, money(nz(v('1600',zb))), delta('1600',zb,pb)));
  if (zp!=null){
    parts.push(card(`Выручка за ${Y(zp)} г.`, money(nz(v('2110',zp))), delta('2110',zp,pp)));
    const np=nz(v('2400',zp)), npPrev=pp!=null?nz(v('2400',pp)):null;
    if (np<0 && npPrev!=null && npPrev<0){
      // убыток два года подряд: показываем размер убытка и сократился он или вырос
      const d=(-np-(-npPrev))/(-npPrev)*100;
      parts.push(card(`Чистый убыток за ${Y(zp)} г.`, money(-np), `<div class="d ${d<=0?'up':'down'}">убыток ${d<=0?'сократился':'вырос'} на ${fmtNum(Math.abs(d),1)}% к ${Y(pp)}</div>`));
    } else parts.push(card(`${np<0?'Чистый убыток':'Чистая прибыль'} за ${Y(zp)} г.`, money(Math.abs(np)), np<0?'':delta('2400',zp,pp)));
  }
  if (zb!=null) parts.push(card('Тип финансовой устойчивости', esc(stabType(base(zb)).t), `<div class="d">трехмерная модель ${stabType(base(zb)).m}</div>`, 'txt'));
  box.innerHTML=parts.join('');
}

/* ===================== Рендер: сетка ввода ===================== */
let curTab='asset';
function parseNum(s){
  s=String(s).trim(); if(!s || s==='-' || s==='—' || s==='–') return null;
  let neg=false; if(/^\(.*\)$/.test(s)){neg=true; s=s.slice(1,-1);}
  s=s.replace(/[\s\u00a0\u202f]/g,'').replace(/[−–]/g,'-').replace(',','.');
  const x=parseFloat(s); if(!Number.isFinite(x)) return null; return neg?-Math.abs(x):x;
}
function renderGrid(){
  const n=nYears(), lines=LINES[curTab];
  let h='<thead><tr><th>Код</th><th>Показатель</th>'+state.years.map((y,i)=>`<th><input data-yr="${i}" value="${y}" aria-label="Год ${i+1}" inputmode="numeric"></th>`).join('')+'</tr></thead><tbody>';
  let row=0;
  for (const l of lines){
    if (l.h){ h+=`<tr class="sec"><td colspan="${2+n}">${esc(l.h)}</td></tr>`; continue; }
    const cls=(l.t===2?'t2':l.t===1?'t1':'')+(l.sub?' sub':'');
    h+=`<tr class="${cls}"><td class="code">${l.c}</td><td class="name">${esc(l.n)}</td>`;
    if (l.sum||l.f){ h+=state.years.map((_,i)=>{const x=v(l.c,i); return `<td class="calc${x<0?' neg':''}">${esc(fmt(N(x)))}</td>`;}).join(''); }
    else { h+=state.years.map((_,i)=>{const x=(state.values[l.c]||[])[i]; return `<td><input data-c="${l.c}" data-i="${i}" data-r="${row}" value="${x==null?'':esc(Math.round(x)===x?x.toLocaleString('ru-RU'):x)}" inputmode="decimal" aria-label="${esc(l.n)}, ${state.years[i]}" class="${x<0?'neg':''}"></td>`;}).join(''); row++; }
    h+='</tr>';
  }
  document.getElementById('grid').innerHTML=h+'</tbody>';
  renderCheck();
}
function renderCheck(){
  const el=document.getElementById('check'); const parts=[];
  state.years.forEach((y,i)=>{
    if(!has('bal',i)) { parts.push(`${y}: баланса нет`); return; }
    const a=v('1600',i), p=v('1700',i), d=nz(a)-nz(p);
    parts.push(Math.abs(d)<1 ? `${y}: актив = пассив ✓` : `<span style="color:var(--red)">${y}: актив − пассив = ${fmt(N(d))}</span>`);
  });
  const noPl = state.years.filter((_,i)=>!has('pl',i));
  if (noPl.length) parts.push(`ОФР не заполнен за ${noPl.join(', ')} (возьмите из отчета за следующий год — там есть прошлый период)`);
  el.innerHTML=parts.join(' &nbsp;·&nbsp; ');
}
function setVal(code,i,x){
  if(!state.values[code]) state.values[code]=Array(nYears()).fill(null);
  state.values[code][i]=x;
}
function refreshAll(){ renderGrid(); renderOut(); save(); }
let outT; function refreshSoon(){ clearTimeout(outT); outT=setTimeout(()=>{ renderOut(); renderCheck(); updateCalcCells(); },200); save(); }
function updateCalcCells(){
  // перерисовать только итоговые ячейки в сетке, не трогая фокус
  const trs=[...document.querySelectorAll('#grid tbody tr')]; let li=0;
  for (const l of LINES[curTab]){
    const tr=trs[li++]; if(l.h||!(l.sum||l.f)) continue;
    [...tr.querySelectorAll('td.calc')].forEach((td,i)=>{const x=v(l.c,i); td.textContent=fmt(N(x)); td.classList.toggle('neg',x<0);});
  }
}

const gridEl=document.getElementById('grid');
gridEl.addEventListener('change',e=>{
  const t=e.target;
  if (t.dataset.yr!=null){ const y=parseInt(t.value,10); if(y>1900&&y<2200){ state.years[+t.dataset.yr]=y; refreshAll(); } else t.value=state.years[+t.dataset.yr]; return; }
  if (t.dataset.c){ const x=parseNum(t.value); setVal(t.dataset.c,+t.dataset.i,x); t.value=x==null?'':x.toLocaleString('ru-RU'); t.classList.toggle('neg',x<0); refreshSoon(); }
});
gridEl.addEventListener('paste',e=>{
  const t=e.target; if(!t.dataset||!t.dataset.c) return;
  const txt=(e.clipboardData||window.clipboardData).getData('text'); if(!/[\t\n]/.test(txt.trim())) return;
  e.preventDefault();
  const lines=txt.replace(/\r/g,'').replace(/\n$/,'').split('\n');
  const r0=+t.dataset.r, i0=+t.dataset.i;
  lines.forEach((ln,dr)=>ln.split('\t').forEach((cell,dc)=>{
    const inp=gridEl.querySelector(`input[data-r="${r0+dr}"][data-i="${i0+dc}"]`); if(!inp) return;
    setVal(inp.dataset.c,i0+dc,parseNum(cell));
  }));
  refreshAll();
});
gridEl.addEventListener('keydown',e=>{
  const t=e.target; if(e.key!=='Enter'||!t.dataset.c) return;
  e.preventDefault(); t.dispatchEvent(new Event('change',{bubbles:true}));
  const nx=gridEl.querySelector(`input[data-r="${+t.dataset.r+(e.shiftKey?-1:1)}"][data-i="${t.dataset.i}"]`); if(nx){nx.focus();nx.select();}
});
document.getElementById('tabs').addEventListener('click',e=>{
  const b=e.target.closest('button[data-tab]'); if(!b) return;
  curTab=b.dataset.tab; document.querySelectorAll('#tabs button').forEach(x=>x.setAttribute('aria-selected',x===b)); renderGrid();
});
document.getElementById('addYear').onclick=()=>{
  if(nYears()>=5) return status('Максимум 5 лет');
  state.years.unshift(state.years[0]-1); for(const k in state.values) state.values[k].unshift(null); refreshAll();
};
document.getElementById('delYear').onclick=()=>{
  if(nYears()<=2) return status('Минимум 2 года');
  state.years.shift(); for(const k in state.values) state.values[k].shift(); refreshAll();
};
document.getElementById('company').addEventListener('input',e=>{ state.company=e.target.value; renderCtx(); save(); });
document.getElementById('hideZero').addEventListener('change',renderOut);
// рисунок выручки и прибыли: в образце его нет, поэтому его можно выключить (номера рисунков пересчитаются)
document.getElementById('revChart').addEventListener('change',e=>{ state.revChart=e.target.checked; renderOut(); save(); });
// Кнопки примеров: data-example — номер в EXAMPLES (0 — НКНХ, 1 — Россети Северный Кавказ)
document.querySelectorAll('[data-example]').forEach(b=>b.onclick=()=>{ state=clone(EXAMPLES[+b.dataset.example]); document.getElementById('company').value=state.company; innIn.value=state.inn; updNalog(); refreshAll(); status('Загружен пример: '+state.company); });
document.getElementById('clearAll').onclick=()=>{ state=blank(); document.getElementById('company').value=''; innIn.value=''; updNalog(); refreshAll(); status('Данные очищены'); };

// поле вывода растягивается по тексту, без внутренней прокрутки
function autosize(ta){ ta.style.height='auto'; ta.style.height=(ta.scrollHeight+2)+'px'; }
document.getElementById('out').addEventListener('input',e=>{ const ta=e.target.closest('textarea[data-id]'); if(ta){ state.conclusions[ta.dataset.id]=ta.value; autosize(ta); save(); }});
// Шаблонный вывод (без ИИ): собирается по правилам из цифр таблицы — см. js/conclusions.js
document.getElementById('out').addEventListener('click',e=>{
  const b=e.target.closest('button[data-auto]'); if(!b) return;
  const id=b.dataset.auto, ta=document.querySelector(`textarea[data-id="${id}"]`), st=document.querySelector(`[data-aist="${id}"]`);
  if(ta.value.trim() && !confirm('Заменить текущий вывод шаблонным?')) return;
  ta.value=autoConclusion(id); state.conclusions[id]=ta.value; autosize(ta); save();
  st.textContent='Готово — проверьте и дополните своими словами.';
});
// Все пустые выводы сразу (уже написанные не трогаем)
document.getElementById('autoAll').onclick=()=>{
  let k=0;
  document.querySelectorAll('textarea[data-id]').forEach(ta=>{ if(ta.value.trim()) return; ta.value=autoConclusion(ta.dataset.id); state.conclusions[ta.dataset.id]=ta.value; autosize(ta); k++; });
  save(); status(k?`Заполнено выводов: ${k}`:'Все выводы уже заполнены');
};

let stT; function status(s){ const el=document.getElementById('status'); el.textContent=s; clearTimeout(stT); stT=setTimeout(()=>el.textContent='',4000); }
function msg(s,err){ const el=document.getElementById('msg'); el.textContent=s||''; el.classList.toggle('err',!!err); }

/* ===================== Импорт Excel ===================== */
document.getElementById('xlsxIn').addEventListener('change',async e=>{
  const f=e.target.files[0]; e.target.value=''; if(!f) return;
  try{
    await loadScript(XLSX_URL); const XL=window.XLSX;
    const wb=XL.read(await f.arrayBuffer());
    const ws=wb.Sheets['Исходные данные']||wb.Sheets[wb.SheetNames[0]];
    const rows=XL.utils.sheet_to_json(ws,{header:1,raw:true,defval:null});
    let yc=null, years=null;
    for(const r of rows){ const idx=r.map((x,j)=>/^(19|20)\d{2}$/.test(String(x??'').trim())?j:-1).filter(j=>j>=0); if(idx.length>=2){ yc=idx; years=idx.map(j=>parseInt(r[j],10)); break; } }
    if(!years) throw new Error('не нашел строку с годами (например 2023, 2024, 2025)');
    const vals={}, rep={}; let cnt=0;
    for(const r of rows){ const code=r.map(x=>String(x??'').trim()).find(x=>/^\d{4}$/.test(x)&&BYCODE[x]);
      if(code&&REPORTED_CODES.includes(code)) rep[code]=yc.map(j=>typeof r[j]==='number'?r[j]:parseNum(r[j]??''));
      if(!code||!INPUT_CODES.includes(code)) continue;
      vals[code]=yc.map(j=>{ const x=typeof r[j]==='number'?r[j]:parseNum(r[j]??''); if(x!=null) cnt++; return x; }); }
    applyImport(years, vals, null, rep);
    msg(`Из Excel загружено ${cnt} значений за ${years.join(', ')}. Проверьте цифры ниже.`);
  }catch(err){ msg('Не удалось прочитать Excel: '+err.message, true); }
});
// Итоговые строки ОФР, которые запоминаем «как в отчете» — для проверки цепочки прибыли (вкладка 2)
const REPORTED_CODES=['2100','2200','2300','2400'];
function applyImport(years, vals, company, rep){
  // vals: {code: [по years]}; rep: {code: [по years]} — итоги из отчета, хранятся по годам: state.reported[code][year]
  if(rep){ state.reported ||= {}; for(const c in rep){ state.reported[c] ||= {}; years.forEach((y,j)=>{ const x=rep[c][j]; if(x!=null) state.reported[c][y]=x; }); } }
  const all=[...new Set([...state.years, ...years])].sort((a,b)=>a-b).slice(-5);
  const nv={};
  for(const c of INPUT_CODES){
    nv[c]=all.map(y=>{
      const j=years.indexOf(y); if(j>=0 && vals[c] && vals[c][j]!=null) return vals[c][j];
      const k=state.years.indexOf(y); return k>=0 && state.values[c] ? state.values[c][k] ?? null : null;
    });
  }
  state.years=all; state.values=nv;
  if(company && (!state.company || EXAMPLES.some(e=>e.company===state.company))){ state.company=company; document.getElementById('company').value=company; }
  refreshAll();
}

/* ===================== Распознавание PDF / фото через Claude ===================== */
let pages=[]; // {label, thumb, render, sel}
const drop=document.getElementById('drop');
['dragenter','dragover'].forEach(ev=>drop.addEventListener(ev,e=>{e.preventDefault();drop.classList.add('over');}));
['dragleave','drop'].forEach(ev=>drop.addEventListener(ev,e=>{e.preventDefault();drop.classList.remove('over');}));
drop.addEventListener('drop',e=>{ const fs=[...e.dataTransfer.files]; const x=fs.find(f=>/\.(xlsx|xls|csv)$/i.test(f.name)); if(x){ const dt=new DataTransfer(); dt.items.add(x); const inp=document.getElementById('xlsxIn'); inp.files=dt.files; inp.dispatchEvent(new Event('change')); } else handleFiles(fs); });
document.getElementById('fileIn').addEventListener('change',e=>{ const fs=[...e.target.files]; e.target.value=''; handleFiles(fs); });
document.querySelectorAll('label.btn[for]').forEach(l=>l.addEventListener('keydown',e=>{ if(e.key==='Enter'||e.key===' '){ e.preventDefault(); document.getElementById(l.htmlFor).click(); }}));

function canvasToBlob(cv,q=0.85){ return new Promise(r=>cv.toBlob(r,'image/jpeg',q)); }
async function renderPdfPage(page, maxW){
  const vp0=page.getViewport({scale:1}); const scale=maxW/vp0.width; const vp=page.getViewport({scale});
  const cv=document.createElement('canvas'); cv.width=Math.round(vp.width); cv.height=Math.round(vp.height);
  const ctx=cv.getContext('2d'); ctx.fillStyle='#fff'; ctx.fillRect(0,0,cv.width,cv.height);
  await page.render({canvasContext:ctx, viewport:vp}).promise; return cv;
}
async function imageToCanvas(file,maxW){
  const url=URL.createObjectURL(file); const img=new Image(); img.src=url; await img.decode();
  const s=Math.min(1,maxW/img.naturalWidth); const cv=document.createElement('canvas'); cv.width=Math.round(img.naturalWidth*s); cv.height=Math.round(img.naturalHeight*s);
  const ctx=cv.getContext('2d'); ctx.fillStyle='#fff'; ctx.fillRect(0,0,cv.width,cv.height); ctx.drawImage(img,0,0,cv.width,cv.height); URL.revokeObjectURL(url); return cv;
}
async function handleFiles(files){
  if(!files.length) return;
  pages=[]; renderThumbs(); msg('Готовлю страницы…');
  try{
    for(const f of files){
      if(f.type==='application/pdf'||/\.pdf$/i.test(f.name)){
        await loadScript(PDF_URL); const lib=window.pdfjsLib; lib.GlobalWorkerOptions.workerSrc=PDFW_URL;
        const pdf=await lib.getDocument({data:await f.arrayBuffer()}).promise;
        const total=Math.min(pdf.numPages,80);
        for(let p=1;p<=total;p++){
          const page=await pdf.getPage(p);
          const cv=await renderPdfPage(page,170);
          let sel=false;
          try{ const tc=await page.getTextContent(); const tx=tc.items.map(it=>it.str).join(' '); sel=/0710001|0710002|Бухгалтерский баланс\s+на|Отч[её]т о финансовых результатах\s+за/i.test(tx) && /\b1600\b|\b2110\b/.test(tx); }catch(e){}
          pages.push({label:`стр. ${p}`, thumb:cv.toDataURL('image/jpeg',0.7), sel, render:async()=>canvasToBlob(await renderPdfPage(page,1700))});
          if(p%6===0){ renderThumbs(); msg(`Готовлю страницы… ${p} из ${total}`); }
        }
      } else if(f.type.startsWith('image/')){
        const cv=await imageToCanvas(f,200);
        pages.push({label:f.name.slice(0,18), thumb:cv.toDataURL('image/jpeg',0.7), sel:true, render:async()=>canvasToBlob(await imageToCanvas(f,1800))});
      }
    }
    renderThumbs();
    const auto=pages.filter(p=>p.sel).length;
    msg(pages.length ? (auto? `Нашел страницы с формами: ${auto}. Проверьте выбор и нажмите «Распознать».` : 'Отметьте страницы с балансом и отчетом о финансовых результатах.') : 'В файлах нет страниц PDF или изображений.', !pages.length);
  }catch(e){ msg('Не удалось открыть файл: '+e.message,true); }
}
function renderThumbs(){
  const box=document.getElementById('thumbs');
  box.innerHTML=pages.map((p,i)=>`<button type="button" class="thumb" data-p="${i}" aria-pressed="${p.sel}" aria-label="${esc(p.label)}"><img src="${p.thumb}" alt=""><span>${esc(p.label)}</span></button>`).join('');
  document.getElementById('recRow').hidden=!pages.length;
  const k=pages.filter(p=>p.sel).length, btn=document.getElementById('recognize');
  btn.textContent=k?`Распознать выбранные страницы (${k})`:'Распознать выбранные страницы';
  btn.disabled=!k||k>6||!sampleFn||!imagesOk;
  document.getElementById('recHint').textContent = !sampleFn||!imagesOk ? 'Распознавание работает, когда страница открыта в Claude. Иначе заполните шаблон Excel.' : k>6 ? 'Не больше 6 страниц за раз — снимите лишние.' : 'Отметьте страницы с балансом и отчетом о финрезультатах (до 6 за раз).';
}
document.getElementById('thumbs').addEventListener('click',e=>{ const b=e.target.closest('.thumb'); if(!b) return; const p=pages[+b.dataset.p]; p.sel=!p.sel; renderThumbs(); });

const EXTRACT_CODES = INPUT_CODES.concat(['1100','1200','1300','1400','1500','1600','1700','2100','2200','2300','2410','2400']);
document.getElementById('recognize').onclick=async()=>{
  const sel=pages.filter(p=>p.sel); if(!sel.length||!sampleFn) return;
  const btn=document.getElementById('recognize'); btn.disabled=true;
  try{
    msg('Готовлю изображения…');
    const images=[]; for(const p of sel) images.push(await p.render());
    msg('Claude читает страницы — обычно 30–90 секунд…');
    const prompt=`Перед тобой изображения страниц российской бухгалтерской отчетности: бухгалтерский баланс (форма по ОКУД 0710001) и/или отчет о финансовых результатах (0710002).
Извлеки значения по кодам строк из ВСЕХ числовых колонок (отчетная дата и прошлые даты/периоды).
Правила:
- Год колонки: для баланса — год даты «На 31 декабря YYYY»; для отчета о финрезультатах — год периода «За YYYY год» (или «За январь–декабрь YYYY»).
- Единицы — тыс. руб. Если в шапке указано «млн руб.», умножь на 1000.
- Число в круглых скобках — отрицательное (верни со знаком минус). Прочерк или пустая ячейка — null.
- Пробелы внутри чисел — разделители разрядов, не десятичные.
- Нужны только эти коды: ${EXTRACT_CODES.join(', ')}. Кодов, которых нет в отчете, не добавляй.
- Название организации возьми из строки «Организация».
Ответь ТОЛЬКО JSON без пояснений:
{"company": "полное название или null", "values": {"1150": {"2025": 447732708, "2024": 364067997}, "2120": {"2025": -204158810}}}`;
    const res=await sampleFn.json(prompt,{images, modelTier:'default'});
    const raw=res&&res.values||{}; const years=new Set();
    for(const c in raw) for(const y in raw[c]) if(/^\d{4}$/.test(y)) years.add(+y);
    const ys=[...years].sort((a,b)=>a-b); if(!ys.length) throw new Error('в ответе нет данных — выберите страницы с формами');
    const vals={}; let cnt=0;
    for(const c of INPUT_CODES){
      if(!raw[c]) continue; const l=BYCODE[c];
      vals[c]=ys.map(y=>{ let x=raw[c][y]; if(x==null||x==='') return null; x=typeof x==='number'?x:parseNum(x); if(x==null) return null;
        if(l.exp) x=Math.abs(x); else if(l.tax) x=-x; cnt++; return x; });
    }
    // если есть только итог налога 2410 — кладем его в текущий налог
    if(raw['2410']&&!raw['2411']&&!raw['2412']) vals['2411']=ys.map(y=>{const x=raw['2410'][y]; return x==null?null:-(typeof x==='number'?x:parseNum(x)); });
    const rep={}; for(const c of REPORTED_CODES) if(raw[c]) rep[c]=ys.map(y=>{ const x=raw[c][y]; return x==null||x===''?null:(typeof x==='number'?x:parseNum(x)); });
    applyImport(ys, vals, typeof res.company==='string'?res.company:null, rep);
    // сверка итогов
    const warn=[];
    for(const tc of ['1600','1700','2400']){ if(!raw[tc]) continue; for(const y of ys){ const i=state.years.indexOf(y); const got=raw[tc][y]; if(i<0||got==null) continue; const calc=v(tc,i); if(calc!=null&&Math.abs(calc-got)>1) warn.push(`${tc} за ${y}: в отчете ${fmt(N(got))}, по строкам ${fmt(N(calc))}`); } }
    msg(`Распознано ${cnt} значений за ${ys.join(', ')}. `+(warn.length?'Не сходятся итоги — проверьте: '+warn.join('; '):'Итоги сходятся с отчетом.'), warn.length>0);
    document.getElementById('s2').scrollIntoView({behavior:'smooth'});
  }catch(e){
    const code=e&&e.code;
    msg(code==='not_granted'?'Распознавание не разрешено. Можно заполнить шаблон Excel.':code==='rate_limited'?'Слишком много запросов, попробуйте через минуту.':'Не получилось распознать: '+(e&&e.message||e), true);
  }finally{ renderThumbs(); }
};

/* ===================== Выводы через Claude ===================== */
function tableText(t){
  const rows=allRows(t), grid=layout(rows,t.ncols);
  return rows.map((row,r)=>{ const out=[]; for(let c=0;c<t.ncols;c++){ const s=grid[r][c]; out.push(s&&s.r0===r&&s.c0===c?fmt(s.cell):''); } return out.join('\t'); }).join('\n');
}
document.getElementById('out').addEventListener('click',async e=>{
  const b=e.target.closest('button[data-ai]'); if(!b||!sampleFn) return;
  const t=buildTables().find(x=>x.id===b.dataset.ai); const ta=document.querySelector(`textarea[data-id="${t.id}"]`); const st=document.querySelector(`[data-aist="${t.id}"]`);
  b.disabled=true; st.textContent='Думаю…';
  try{
    const prompt=`Ты помогаешь студенту написать вывод к таблице в работе по финансовому учету и анализу.
Организация: ${state.company||'организация'}. Таблица ${t.num} – ${t.title}. Данные (тыс. руб., если не указано иное; «—» — нет данных):
${tableText(t)}

Напиши вывод на 4–7 предложений в официально-деловом стиле курсовой работы, на русском. Опирайся на конкретные цифры из таблицы, называй годы, сравнивай с нормативами, если они есть. Не выдумывай данных, которых нет в таблице. Без заголовков, списков и markdown — только связный текст.`;
    const r=await sampleFn(prompt,{modelTier:'default', onText:({text})=>{ ta.value=text; st.textContent=''; }});
    ta.value=r.text.trim(); state.conclusions[t.id]=ta.value; save(); st.textContent='Готово — отредактируйте под себя.';
  }catch(err){ st.textContent = err&&err.code==='not_granted'?'Claude не разрешен для этой страницы.':'Не получилось: '+(err&&err.message||err); if(err&&err.text){ ta.value=err.text; } }
  finally{ b.disabled=false; }
});

/* ===================== Ссылка на ГИР БО ===================== */
const innIn=document.getElementById('innIn'), nalogLink=document.getElementById('nalogLink');
function updNalog(){
  const q=(innIn.value||state.company||'').trim();
  nalogLink.href = q ? 'https://bo.nalog.gov.ru/search?query='+encodeURIComponent(q) : 'https://bo.nalog.gov.ru/';
}
innIn.addEventListener('input',()=>{ state.inn=innIn.value; updNalog(); save(); });
document.getElementById('company').addEventListener('input',updNalog);

/* ===================== Старт ===================== */
state=load();
document.getElementById('company').value=state.company||'';
innIn.value=state.inn||''; updNalog();
document.getElementById('revChart').checked = state.revChart!==false;
refreshAll();
showPane(location.hash==='#fr' ? 'fr' : (()=>{ try{ return localStorage.getItem('fua-pane'); }catch(e){ return null; } })());
(async()=>{
  if(!window.claude||!window.claude.use) { renderThumbs(); return; }
  try{ downloadsNs=await window.claude.use('downloads'); }catch(e){}
  try{ sampleFn=await window.claude.use('sample'); if(sampleFn){ const lim=await sampleFn.limits(); imagesOk=!!(lim&&lim.images); } }catch(e){ sampleFn=null; }
  if(!downloadsNs){ ['dlXlsx','dlDocx','dlTpl'].forEach(id=>document.getElementById(id).title='Скачивание доступно в Claude'); }
  document.querySelectorAll('.ai-only').forEach(b=>b.hidden=!sampleFn);
  renderThumbs();
})();
