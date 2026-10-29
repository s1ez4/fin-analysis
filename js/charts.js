/* Графики (Chart.js): рисунки на странице под таблицами и PNG-картинки для Word */
const CHART_URL='https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js';

/* ===================== Палитра ===================== */
// cat — цвета серий (порядок подобран так, чтобы соседние цвета различались и при дальтонизме);
// ramp — ступени одного синего для годов: чем новее год, тем заметнее столбик
const PALETTE={
  light:{cat:['#2a78d6','#eb6834','#1baf7a','#eda100','#e87ba4','#008300','#6250d6','#e34948'], ramp:['#86b6ef','#5598e7','#2a78d6','#1c5cab','#104281']},
  dark: {cat:['#3987e5','#d95926','#199e70','#c98500','#d55181','#008300','#9085e9','#e66767'], ramp:['#184f95','#256abf','#2a78d6','#5598e7','#86b6ef']},
};
// цвет подписи внутри цветного сегмента: черный на светло-желтом, иначе белый
const inkOn = c => c==='#eda100' ? '#0b0b0b' : '#ffffff';
// n цветов для n годов, равномерно по шкале (2 года — самый светлый и самый темный)
function yearColors(n, ramp){ return Array.from({length:n},(_,k)=> ramp[n===1?2:Math.round(k*(ramp.length-1)/(n-1))]); }

// Тема: для Word — всегда светлая на белом и Times New Roman; на странице — цвета из CSS-переменных
function chartTheme(forExport){
  if (forExport) return {mode:'light', surface:'#ffffff', ink:'#0b0b0b', soft:'#52514e', grid:'#e1e0d9', axis:'#c3c2b7', font:'"Times New Roman", Times, serif'};
  const cs=getComputedStyle(document.documentElement), g=k=>cs.getPropertyValue(k).trim();
  return {mode: cs.colorScheme==='dark'?'dark':'light', surface:g('--sheet-2'), ink:g('--ink'), soft:g('--ink-soft'), grid:g('--rule'), axis:g('--rule-strong'), font:'"Golos Text", system-ui, sans-serif'};
}

/* ===================== Что рисуем ===================== */
// Рисунки стоят как в образце: 2.2–2.3 — после горизонтального анализа, 2.4–2.5 «Доля активов/пассивов» —
// после вертикального, затем ликвидность и финансовая устойчивость. Рисунок выручки и прибыли в образце
// нет — его можно выключить галочкой, тогда номера совпадут с образцом (2.6 и 2.7). Номера ставятся по порядку.
// kind: 'stack100' — структура в % (столбики 100%), 'bars' — суммы по годам, 'ratios' — коэффициенты по годам
const cap = s => s[0].toUpperCase()+s.slice(1);
function chartDefs(){
  const hv=yearsHas('bal'), hp=yearsHas('pl');
  const yb=state.years.map((_,i)=>i).filter(i=>hv[i]), yp=state.years.map((_,i)=>i).filter(i=>hp[i]);
  const g=(c,i)=>nz(v(c,i)), per=I=>I.length?` за ${Y(I[0])}-${Y(I.at(-1))} гг.`:'';
  const defs=[
    {id:'fA',after:'t22',title:'Динамика структуры актива баланса'+per(yb),kind:'bars',unit:'млн руб.',idx:yb,series:[
      ['Внеоборотные активы',i=>v('1100',i)],
      ['Оборотные активы',i=>v('1200',i)]]},
    {id:'fP',after:'t23',title:'Динамика структуры пассива баланса'+per(yb),kind:'bars',unit:'млн руб.',idx:yb,series:[
      ['Собственный капитал',i=>v('1300',i)],
      ['Долгосрочные обязательства',i=>v('1400',i)],
      ['Краткосрочные обязательства',i=>v('1500',i)]]},
    {id:'fAs',after:'t24',title:'Доля активов баланса'+per(yb),kind:'stack100',idx:yb,total:b=>b.VB,series:[
      ['Внеоборотные активы',b=>b.VOA],
      ['Запасы',b=>b.Z],
      ['Дебиторская задолженность',b=>b.DZ],
      ['Денежные средства и финансовые вложения',b=>b.DS+b.KFV],
      ['Прочие оборотные активы',b=>b.OA-b.Z-b.DZ-b.DS-b.KFV]]},
    {id:'fPs',after:'t25',title:'Доля пассивов баланса'+per(yb),kind:'stack100',idx:yb,total:b=>b.VB,series:[
      ['Собственный капитал',b=>b.SK],
      ['Долгосрочные обязательства',b=>b.DO],
      ['Краткосрочные заемные средства',b=>b.KKZ],
      ['Кредиторская задолженность',(b,i)=>g('1520',i)],
      ['Прочие краткосрочные обязательства',(b,i)=>b.KO-b.KKZ-g('1520',i)]]},
    {id:'fRev',after:'t27',title:'Динамика выручки и прибыли'+per(yp),kind:'bars',unit:'млн руб.',idx:yp,series:[
      ['Выручка',i=>v('2110',i)],
      ['Валовая прибыль',i=>v('2100',i)],
      ['Прибыль от продаж',i=>v('2200',i)],
      ['Чистая прибыль',i=>v('2400',i)]]},
    {id:'fLiq',after:'t28',title:'Динамика коэффициентов ликвидности',kind:'ratios',idx:yb,defs:LIQ},
    {id:'fFin',after:'t29',title:'Изменения коэффициентов финансовой устойчивости',kind:'ratios',idx:yb,defs:FIN},
  ].filter(d=>d.idx.length && (d.id!=='fRev' || state.revChart!==false));
  defs.forEach((d,k)=>d.num='2.'+(k+2));   // сквозная нумерация: 2.2, 2.3, …
  return defs;
}

// длинную подпись по оси X режем на строки примерно по 16 символов
// очень длинное слово (> 17 букв) переносим по слогам: после гласной перед «согласная + гласная»,
// ближе к середине («Самофинанси-рования»), иначе соседние подписи слипаются
const isV = ch => /[аеёиоуыэюяАЕЁИОУЫЭЮЯ]/.test(ch||'');
function hyph(x){
  if(x.length<=17) return [x];
  let best=-1;
  for(let i=3;i<x.length-3;i++) if(isV(x[i-1])&&!isV(x[i])&&isV(x[i+1]) && (best<0||Math.abs(i-x.length/2)<Math.abs(best-x.length/2))) best=i;
  return best<0 ? [x] : [x.slice(0,best)+'-', x.slice(best)];
}
function wrap(s, w=16){ const out=[]; let cur='';
  const words=s.split(' ').flatMap(hyph);
  for(const word of words){ if(cur && (cur+' '+word).length>w){ out.push(cur); cur=word; } else cur=cur?cur+' '+word:word; } if(cur) out.push(cur); return out; }

/* ===================== Плагины ===================== */
// Заливка фона (иначе PNG в Word будет прозрачным)
const bgPlugin={id:'bgFill', beforeDraw(ch,_,o){ const c=ch.ctx; c.save(); c.fillStyle=o.color||'#fff'; c.fillRect(0,0,ch.width,ch.height); c.restore(); }};
// Подписи значений: в Word нет всплывающих подсказок, поэтому цифры пишем прямо на графике.
// mode 'inside' — внутри сегмента (если он не меньше min), 'top' — над столбиком (под ним, если значение < 0)
const labelPlugin={id:'valueLabels', afterDatasetsDraw(ch,_,o){
  if(!o.mode) return;
  const c=ch.ctx; c.save(); c.font=`${o.size||11}px ${o.font}`; c.textAlign='center';
  ch.data.datasets.forEach((ds,k)=>{
    const meta=ch.getDatasetMeta(k); if(meta.hidden) return;
    meta.data.forEach((el,i)=>{
      const val=ds.data[i]; if(val==null||!Number.isFinite(val)) return;
      const text=o.format(val);
      if(o.mode==='inside'){
        if(Math.abs(val)<(o.min||0)) return;
        c.fillStyle=inkOn(ds.backgroundColor); c.textBaseline='middle';
        c.fillText(text, el.x, (el.y+el.base)/2);
      } else {
        c.fillStyle=o.ink; c.textBaseline= val<0?'top':'bottom';
        c.fillText(text, el.x, val<0? el.y+3 : el.y-3);
      }
    });
  });
  c.restore();
}};

/* ===================== Конфигурация Chart.js ===================== */
function chartConfig(d, th, forExport){
  const pal=PALETTE[th.mode], labels=d.idx.map(i=>String(Y(i))), font={family:th.font, size:forExport?14:12};
  let datasets, xLabels=labels, valueLabels, yTicks, stacked=false, yMax, xFont=font;
  if(d.kind==='stack100'){
    stacked=true; yMax=100;
    datasets=d.series.map(([name,fn],k)=>({label:name, backgroundColor:pal.cat[k],
      data:d.idx.map(i=>{ const b=base(i), x=div(fn(b,i), d.total(b)); return x==null?null:x*100; }),
      borderColor:th.surface, borderWidth:{top:2}, borderSkipped:false}));
    valueLabels={mode:'inside', min:4, format:x=>fmtNum(x,1), size:forExport?14:11};
    yTicks=x=>x+'%';
  } else if(d.kind==='bars'){
    datasets=d.series.map(([name,fn],k)=>({label:name, backgroundColor:pal.cat[k], borderRadius:3,
      data:d.idx.map(i=>{ const x=fn(i); return x==null?null:x/1000; })}));
    valueLabels={mode:'top', format:x=>fmtNum(x,0), ink:th.soft, size:forExport?13:11};
    yTicks=x=>fmtNum(x,0);
  } else {
    // коэффициенты: по оси X — коэффициенты, серии — годы
    // если коэффициентов много (2.9 — восемь), шрифт мельче, иначе подписи наезжают друг на друга
    const dense=d.defs.length>5;
    if(dense) xFont={...font, size:forExport?12:11};
    xLabels=d.defs.map(([name])=>wrap(cap(name), dense?14:16));
    const cols=yearColors(d.idx.length, pal.ramp);
    datasets=d.idx.map((i,k)=>({label:String(Y(i)), backgroundColor:cols[k], borderRadius:3,
      data:d.defs.map(([,,fn])=>fn(base(i)))}));
    valueLabels={mode:'top', format:x=>fmtNum(x,2), ink:th.soft, size:dense?(forExport?10:10):(forExport?13:11)};
    yTicks=x=>fmtNum(x,1);
  }
  return {
    type:'bar',
    data:{labels:xLabels, datasets},
    plugins:[bgPlugin,labelPlugin],
    options:{
      responsive:!forExport, maintainAspectRatio:false, animation:false,
      devicePixelRatio: forExport?2:undefined,
      layout:{padding:{top:18,right:8,left:4,bottom:4}},
      categoryPercentage:.78, barPercentage:.92,
      scales:{
        x:{stacked, grid:{display:false}, border:{color:th.axis}, ticks:{color:th.soft, font:xFont, autoSkip:false, maxRotation:0}},
        y:{stacked, max:yMax, beginAtZero:true, grid:{color:th.grid}, border:{display:false}, ticks:{color:th.soft, font, callback:yTicks},
           title:{display:!!d.unit, text:d.unit||'', color:th.soft, font}},   // единицы измерения у оси (млн руб.)
      },
      plugins:{
        bgFill:{color:th.surface},
        valueLabels:{...valueLabels, font:th.font},
        legend:{position:'bottom', labels:{color:th.ink, font, boxWidth:12, boxHeight:12, padding:14}},
        tooltip:{titleFont:font, bodyFont:font, callbacks:{
          title:items=>{ const l=items[0].label; return Array.isArray(l)?l.join(' '):l; },
          label:it=>`${it.dataset.label}: ${it.raw==null?'—':valueLabels.format(it.raw)}${d.kind==='stack100'?'%':''}`}},
      },
    },
  };
}

/* ===================== Рисунки на странице ===================== */
let charts=[], chartGen=0;
// HTML-заготовки рисунков, которые стоят после таблицы tableId
function chartFiguresHTML(tableId){
  return chartDefs().filter(d=>d.after===tableId && d.idx.length).map(d=>`
    <figure class="chart" id="${d.id}">
      <div class="chartbox"><canvas role="img" aria-label="Рисунок ${d.num} – ${esc(d.title)}"></canvas></div>
      <figcaption>Рисунок ${d.num} – ${esc(d.title)}</figcaption>
      <div class="concl">
        <textarea data-id="${d.id}" placeholder="Вывод по рисунку ${d.num} — напишите сами или нажмите «Сгенерировать вывод»">${esc(state.conclusions[d.id]||'')}</textarea>
        <div class="row">
          <button class="btn ghost small" data-auto="${d.id}">Сгенерировать вывод</button>
          <small data-aist="${d.id}"></small>
        </div>
      </div>
    </figure>`).join('');
}
async function drawCharts(){
  const gen=++chartGen;
  try{ await loadScript(CHART_URL); }
  catch(e){ document.querySelectorAll('figure.chart .chartbox').forEach(el=>el.innerHTML='<p class="hint">График не загрузился: для Chart.js нужен интернет.</p>'); return; }
  if(gen!==chartGen) return; // за время загрузки страницу уже перерисовали
  charts.forEach(c=>c.destroy()); charts=[];
  const th=chartTheme(false);
  for(const d of chartDefs()){
    const cv=document.querySelector(`#${d.id} canvas`); if(!cv) continue;
    charts.push(new Chart(cv, chartConfig(d,th,false)));
  }
}
// при смене светлой/темной темы перерисовываем цвета
matchMedia('(prefers-color-scheme: dark)').addEventListener('change',()=>drawCharts());

/* ===================== Картинки для Word ===================== */
// Рисуем каждый график на скрытом холсте 800×400 (×2 для четкости) и забираем PNG
async function chartImages(){
  await loadScript(CHART_URL);
  const th=chartTheme(true), box=document.createElement('div'), out=[];
  box.style.cssText='position:fixed;left:-10000px;top:0;width:800px;height:400px;';
  document.body.appendChild(box);
  try{
    for(const d of chartDefs()){
      if(!d.idx.length) continue;
      const cv=document.createElement('canvas'); cv.width=800; cv.height=400; box.appendChild(cv);
      const ch=new Chart(cv, chartConfig(d,th,true));
      const blob=await new Promise(r=>cv.toBlob(r,'image/png'));
      out.push({id:d.id, num:d.num, title:d.title, after:d.after, png:await blob.arrayBuffer()});
      ch.destroy(); cv.remove();
    }
  } finally { box.remove(); }
  return out;
}
