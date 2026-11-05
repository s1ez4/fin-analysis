/* Таблицы для Word и Excel в оформлении как в образце «ФУиА задание 1.doc».
   Параметры каждой таблицы сняты прямо из XML образца: кегль, межстрочный интервал, рамки таблицы и ячеек,
   отступ, автоподбор, ширины колонок, высоты строк шапки, выравнивание каждой ячейки шапки и первой колонки,
   «Итого по разделу» — жирный курсив, «БАЛАНС» — жирный, числа без пробелов между разрядами, минус — дефисом.
   На странице таблицы остаются прежними (там удобнее читать), эти — только для скачивания. */

// Число как в образце: без разделителей разрядов, дефис для минуса, лишние нули после запятой убраны.
// plus — ставить «+» у положительных (в образце так в отклонениях 2.6 и 2.8). Пусто — «-».
function fx(x, d=0, plus=false){
  if (x==null || !Number.isFinite(x)) return '-';
  let r = Math.round(x*10**d)/10**d; if (r===0) r=0; // убираем «-0»
  const s = r.toLocaleString('ru-RU',{useGrouping:false, maximumFractionDigits:d}).replace(/[−–]/g,'-');
  return plus && r>0 ? '+'+s : s;
}

// Ячейки: C — текст, Nm — число, H — шапка (жирная, по центру).
// Свойства ячейки: jc — выравнивание по горизонтали, va — по вертикали (null — не задавать, как в образце),
// pind — отступы абзаца слева/справа в twips (в 2.4/2.5 у «удельный вес, %» −108), '\n' в тексте — новый абзац.
const C  = (t,o={}) => ({t:String(t),...o});
const Nm = (v,d=0,o={}) => ({v,d,...o});
const H  = (t,o={}) => ({t:String(t),hd:true,b:true,jc:'center',va:'center',...o});
// строка раздела («I. Внеоборотные активы»): название жирным, остальные клетки пустые, как в образце
const sectionRow = (text,n,va) => [C(text,{b:true,va}), ...Array.from({length:n-1},()=>C('',{jc:'right',va:'bottom'}))];

// Ширины колонок (twips, 1 см = 567): для 3 лет — ровно как в образце,
// для другого числа лет первая колонка та же, остальная ширина делится поровну
function widths(sample, ncols){
  if (ncols===sample.length) return sample.slice();
  const total=sample.reduce((a,b)=>a+b,0), first=sample[0], other=Math.floor((total-first)/(ncols-1));
  return [first, ...Array(ncols-1).fill(other)];
}
// Высота строки шапки: «точно» как в образце — только когда колонок столько же, сколько в образце
// (при 2, 4 или 5 годах колонки уже, текст может не влезть — тогда «не меньше»)
const hh = (h, exact) => ({h, exact});

/* Общие свойства таблицы (по умолчанию — как у 2.2–2.5 и 2.8–2.9 в образце):
   line — интервал 1,5 внутри ячеек; tblB/cellB — рамки таблицы и ячеек; ind — отступ таблицы слева;
   layout — fixed (фиксированные ширины) или auto (автоподбор); tblW — ширина таблицы, если задана;
   headH — высоты строк шапки; rowH/rowExact — высота строк тела; cont — оформление продолжений объединенных ячеек */
const BASE = {line:false, tblB:{val:'single',sz:4}, cellB:null, ind:108, layout:'fixed', tblW:null, headH:[], rowH:null, rowExact:false, cont:()=>null};

/* ---------- 2.2 / 2.3 Горизонтальный анализ (8 пт) ---------- */
function xHorizontal(id,num,title,lines,first,grid){
  const P=pairs(), hv=yearsHas('bal'), ncols=1+nYears()+2*P.length, same=ncols===grid.length, is22=id==='t22';
  // 2.2: заголовки групп отклонений без вертикального выравнивания, «2020/2019» — по левому краю внизу;
  // 2.3: всё по центру
  const grp = is22 ? {va:null} : {};
  const sub = is22 ? {jc:'left',va:'bottom'} : {};
  const head=[
    [H(first,{rs:2}), ...state.years.map(y=>H(y+'г.',{rs:2})), H('Абсолютное отклонение, тыс.руб',{cs:P.length,...grp}), H('Относительное отклонение, %',{cs:P.length,...grp})],
    [...P.map(i=>H(Y(i)+'/'+Y(i-1),sub)), ...P.map(i=>H(Y(i)+'/'+Y(i-1),sub))],
  ];
  const firstVa = is22 ? 'center' : 'bottom';   // первая колонка: в 2.2 по центру, в 2.3 — по нижнему краю
  const val=(l,i)=>hv[i]?v(l.c,i):null;
  const body=visible(lines,l=>state.years.map((_,i)=>val(l,i))).map(l=>{
    if (l.h) return sectionRow(l.h, ncols, firstVa);
    const x=state.years.map((_,i)=>val(l,i));
    const st = l.t===1?{b:true,i:true} : l.t===2?{b:true} : {};   // «Итого» — жирный курсив, «БАЛАНС» — жирный
    return [C(l.n,{...st,va:firstVa}), ...x.map(a=>Nm(a,0,st)),
      ...P.map(i=>Nm(x[i]!=null&&x[i-1]!=null?x[i]-x[i-1]:null,0,st)),
      ...P.map(i=>Nm(x[i]!=null&&x[i-1]?x[i]/x[i-1]*100:null,2))];
  });
  const prof = is22
    ? {line:true, headH:[hh(289,false)]}                                              // 2.2: интервал 1,5
    : {cellB:{val:'single',sz:4}, headH:[hh(384,same),hh(277,same)], rowH:284,       // 2.3: одинарный, шапка 384/277, строки не ниже 0,5 см
       cont:c=>c>0?{jc:'center',va:'center'}:null};
  return {...BASE,...prof,id,num,title,ncols,grid:widths(grid,ncols),sz:16,head,body};
}

/* ---------- 2.4 / 2.5 Вертикальный анализ (8 пт, одинарный интервал) ---------- */
function xVertical(id,num,title,lines,first,totalCode,grid){
  const P=pairs(), hv=yearsHas('bal'), ncols=1+2*nYears()+P.length, same=ncols===grid.length, is24=id==='t24';
  const head=[
    [H(first,{rs:2}), ...state.years.map(y=>H('31.12.'+y,{cs:2})), H('Отклонение, %',{cs:P.length})],
    [...state.years.flatMap(()=>[H('сумма, тыс. руб.'),H('удельный вес, %',{pind:-108})]),
     // в 2.4 «2020/ 2019» одной строкой, в 2.5 — «2020/» и «2019» отдельными абзацами
     ...P.map(i=>H(is24 ? Y(i)+'/ '+Y(i-1) : Y(i)+'/\n'+Y(i-1)))],
  ];
  const firstVa = is24 ? 'center' : 'bottom';
  const ls = is24 ? {line:true} : {};   // в 2.4 у чисел интервал 1,5 (у названий — одинарный), как в образце
  const val=(l,i)=>hv[i]?v(l.c,i):null;
  const body=visible(lines,l=>state.years.map((_,i)=>val(l,i))).map(l=>{
    if (l.h) return sectionRow(l.h, ncols, firstVa);
    const x=state.years.map((_,i)=>val(l,i));
    const sh=x.map((a,i)=>{ const s=a!=null?div(a, hv[i]?v(totalCode,i):null):null; return s==null?null:s*100; });
    const st = l.t===1?{b:true,i:true} : l.t===2?{b:true} : {};
    const dev = l.t===2 ? P.map(()=>C('-',{jc:'right',va:'bottom',...ls}))      // у строки «БАЛАНС» отклонения — прочерк
                        : P.map(i=>Nm(sh[i]!=null&&sh[i-1]!=null?sh[i]-sh[i-1]:null,2,ls));
    return [C(l.n,{...st,va:firstVa}), ...x.flatMap((a,i)=>[Nm(a,0,{...st,...ls}), Nm(sh[i],2,ls)]), ...dev];
  });
  const prof = is24
    ? {cellB:{val:'single',sz:4}, headH:[hh(284,same),hh(406,same)], rowH:284, cont:()=>({jc:'center',va:'center'})}
    : {cellB:{val:'single',sz:4}, cont:()=>({jc:'center',va:'center'})};             // 2.5: высоты строк не заданы
  return {...BASE,...prof,id,num,title,ncols,grid:widths(grid,ncols),sz:16,head,body};
}

/* ---------- 2.6 ДЗ и КЗ (12 пт; рамка таблицы пунктиром, у ячеек — сплошная; автоподбор) ---------- */
function xDebt(){
  const hv=yearsHas('bal'), ncols=2+nYears();
  const dz=state.years.map((_,i)=>hv[i]?v('1230',i):null), kz=state.years.map((_,i)=>hv[i]?v('1520',i):null);
  const head=[[H('Показатели'),H('Единица измерения'),...state.years.map(y=>H('на 31.12.'+String(y).slice(2)))]];
  const name=t=>C(t,{va:'center'}), unit=t=>C(t,{va:'center'});
  const body=[
    [name('Остаток дебиторской задолженности'),unit('тыс. руб.'),...dz.map(a=>Nm(a))],
    [name('Остаток кредиторской задолженности'),unit('тыс. руб.'),...kz.map(a=>Nm(a))],
    [name('Разница между кредиторской и дебиторской задолженностью'),unit('тыс. руб.'),...kz.map((a,i)=>Nm(a!=null&&dz[i]!=null?a-dz[i]:null,0,{plus:true}))],
    [name('Приходится кредиторской задолженности на 1 рубль дебиторской задолженности'),unit('руб.'),...kz.map((a,i)=>Nm(a!=null?div(a,dz[i]):null,2))],
  ];
  return {...BASE, tblB:{val:'dotted',sz:4}, cellB:{val:'single',sz:4}, ind:88, layout:'auto', tblW:9741, headH:[hh(435,false)], rowH:435,
    id:'t26',num:'2.6',title:'Соотношение дебиторской и кредиторской задолженности',ncols,grid:widths([4313,1368,1353,1353,1354],ncols),sz:24,head,body};
}

/* ---------- 2.7 Формирование прибыли (10 пт; рамки только у ячеек, толщиной 1 пт; строки ровно 0,5 см) ---------- */
const PL_NAMES={'2411':'в т.ч. текущий налог на прибыль'};
function xProfit(){
  const P=pairs(), hp=yearsHas('pl'), ncols=1+nYears()+P.length;
  // годы в шапке не объединены по вертикали: «2019 год» внизу верхней клетки (над ним два пустых абзаца), нижняя клетка пустая
  const head=[
    [H('Наименование',{rs:2}), ...state.years.map(y=>H('\n\n'+y+' год')), H('Отклонения (влияние основных факторов на сумму прибыли)',{cs:P.length})],
    [...state.years.map(()=>H('',{b:false})), ...P.map(i=>H(Y(i)+'г.'))],
  ];
  const val=(l,i)=>hp[i]?v(l.c,i):null;
  const body=visible(LINES.pl,l=>state.years.map((_,i)=>val(l,i))).map(l=>{
    const x=state.years.map((_,i)=>val(l,i));
    return [C(PL_NAMES[l.c]||l.n,{va:'center'}), ...x.map(a=>Nm(a,0,{va:'center'})), ...P.map(i=>Nm(x[i]!=null&&x[i-1]!=null?x[i]-x[i-1]:null))];
  });
  return {...BASE, tblB:null, cellB:{val:'single',sz:8}, ind:93, layout:'auto', tblW:9678, headH:[hh(574,false),hh(269,false)], rowH:284, rowExact:true,
    cont:()=>({jc:'center',va:'center'}),
    id:'t27',num:'2.7',title:'Анализ формирования прибыли',ncols,grid:widths([4410,1016,992,992,1134,1134],ncols),sz:20,head,body};
}

/* ---------- 2.8 / 2.9 Коэффициенты (11 пт, без столбца «Норматив» — как в образце) ---------- */
function xRatios(id,num,title,defs,plus){
  const P=pairs(), hv=yearsHas('bal'), ncols=1+nYears()+P.length, is28=id==='t28';
  // 2.8: шапка по левому краю, без вертикального выравнивания; 2.9: по центру, без вертикального выравнивания
  const hs = is28 ? {jc:'left',va:null} : {va:null};
  const head=[[H('Коэффициенты:',hs), ...state.years.map(y=>H(String(y),hs)), ...P.map(i=>H(Y(i)+'-'+Y(i-1),hs))]];
  const body=defs.map(([name,,fn,d])=>{
    const x=state.years.map((_,i)=>hv[i]?fn(base(i)):null);
    // разницу считаем по округленным значениям — чтобы в таблице «0,45 − 0,50 = -0,05» сходилось
    const r=a=>a==null?null:Math.round(a*10**d)/10**d;
    return [C(name,{va:null}), ...x.map(a=>Nm(a,d,is28?{va:null}:{})), ...P.map(i=>Nm(x[i]!=null&&x[i-1]!=null?r(x[i])-r(x[i-1]):null,d,{plus}))];
  });
  return {...BASE, layout:is28?'auto':'fixed', id,num,title,ncols,grid:widths([3969,993,992,992,1276,1240],ncols),sz:22,head,body};
}

/* ---------- 2.10 / 2.11 — в образце их нет, оформляем как 2.9 (11 пт) ---------- */
function xFromModel(t){
  const ncols=t.ncols;
  const head=t.head.map(r=>r.map(c=>H(c.t,{cs:c.cs,rs:c.rs,va:null})));
  const body=t.rows.map(r=>{
    if (r.sec) return sectionRow(r.cells[0].t, ncols, null);
    const st = r.t ? {b:true} : {};
    return r.cells.map((c,k)=> c.v!==undefined ? Nm(c.v, c.k==='num'?0:c.d, st) : C(c.t, k>0?{...st,jc:'center'}:{...st,va:null}));
  });
  return {...BASE, id:t.id,num:t.num,title:t.title,ncols,grid:widths([3969,...Array(ncols-1).fill(Math.floor(5493/(ncols-1)))],ncols),sz:22,head,body};
}

// Все таблицы для скачивания — в том же порядке и с теми же id, что на странице
function exportTables(){
  const yr = state.years[0]+' – '+state.years[nYears()-1]+' гг.';
  return [
    xHorizontal('t22','2.2','Горизонтальный анализ актива баланса за '+yr, LINES.asset,'Актив',[2410,992,993,992,992,992,993,992]),
    xHorizontal('t23','2.3','Горизонтальный анализ пассива баланса за '+yr, LINES.liab,'Пассив',[2977,992,993,992,942,901,941,901]),
    xVertical('t24','2.4','Вертикальный анализ актива баланса за '+yr, LINES.asset,'Активы','1600',[2410,1134,851,1134,850,992,851,709,708]),
    xVertical('t25','2.5','Вертикальный анализ пассива баланса за '+yr, LINES.liab,'Пассив','1700',[2694,992,850,993,850,992,851,709,708]),
    xDebt(),
    xProfit(),
    xRatios('t28','2.8','Сравнительная таблица коэффициентов ликвидности',LIQ,true),
    xRatios('t29','2.9','Сравнительная таблица показателей финансовой устойчивости',FIN,false),
    xFromModel(tblStability()),
    xFromModel(tblLiqGroups()),
  ];
}
