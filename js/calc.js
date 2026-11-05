/* Расчеты показателей и модель таблиц анализа (2.2–2.11) */
/* ===================== Расчеты ===================== */
const nz = x => x==null || Number.isNaN(x) ? 0 : x;
const div = (a,b) => (b && Number.isFinite(a) && Number.isFinite(b)) ? a/b : null;
function v(code, i){
  const l = BYCODE[code];
  if (l.sum){ const xs = l.sum.map(c=>v(c,i)); return xs.every(x=>x==null) ? null : xs.reduce((a,b)=>a+nz(b),0); }
  if (l.f){ const xs = l.deps.map(c=>v(c,i)); if (xs.every(x=>x==null)) return null; return l.f(c=>nz(v(c,i))); }
  const arr = state.values[code]; const x = arr ? arr[i] : null; return x==null ? null : x;
}
function has(stmt, i){ return INPUT_CODES.some(c => STMT_OF(c)===stmt && v(c,i)!=null); }
function nYears(){ return state.years.length; }
function pairs(){ const p=[]; for(let i=1;i<nYears();i++) p.push(i); return p; }
const Y = i => state.years[i];

function base(i){
  const g = c => nz(v(c,i));
  const SK=g('1300'), VB=g('1600'), DO=g('1400'), KO=g('1500'), ZK=DO+KO, VOA=g('1100'), OA=g('1200');
  const Z=g('1210'), NDS=g('1220'), DZ=g('1230'), KFV=g('1240'), DS=g('1250');
  const SOK=SK-VOA, DKZ=g('1410'), KKZ=g('1510');
  const SDI=SOK+DKZ, OI=SDI+KKZ;
  return {SK,VB,ZK,VOA,OA,Z,NDS,DZ,KFV,DS,KO,DO,SOK,DKZ,KKZ,SDI,OI,
    dSOK:SOK-Z, dSDI:SDI-Z, dOI:OI-Z,
    A1:DS+KFV, A2:DZ+g('1260'), A3:Z+NDS+g('1215'), A4:VOA,
    P1:g('1520')+g('1550'), P2:KKZ, P3:DO, P4:SK+g('1530')+g('1540')};
}
const LIQ = [
  ['абсолютной ликвидности','0,15–0,2',b=>div(b.DS+b.KFV,b.KO),3],
  ['срочной ликвидности','0,5–0,8',b=>div(b.DS+b.KFV+b.DZ,b.KO),2],
  ['ликвидности при мобилизации средств','0,5–0,7',b=>div(b.Z+b.NDS,b.KO),2],
  ['общей ликвидности','1–2',b=>div(b.DS+b.KFV+b.DZ+b.Z,b.KO),2],
];
const FIN = [
  ['финансовой независимости','≥ 0,5',b=>div(b.SK,b.VB),2],
  ['финансовой напряженности','≤ 0,5',b=>div(b.ZK,b.VB),2],
  ['самофинансирования','≥ 1',b=>div(b.SK,b.ZK),2],
  ['задолженности','≤ 0,67',b=>div(b.ZK,b.SK),2],
  ['обеспеченности собственными оборотными средствами','≥ 0,1',b=>div(b.SOK,b.OA),2],
  ['маневренности собственных оборотных средств','0,2–0,5',b=>div(b.SOK,b.SK),2],
  ['соотношения мобильных и иммобилизованных активов','индивидуально',b=>div(b.OA,b.VOA),2],
  ['имущества производственного назначения','≥ 0,5',b=>div(b.VOA+b.Z,b.VB),2],
];
function stabType(b){
  const m=[b.dSOK>=0?1:0, b.dSDI>=0?1:0, b.dOI>=0?1:0];
  const key=m.join('');
  const t = key==='111'?'Абсолютная финансовая устойчивость':key==='011'?'Нормальная финансовая устойчивость':key==='001'?'Неустойчивое финансовое состояние':key==='000'?'Кризисное финансовое состояние':'Нестандартное сочетание';
  return {m:'('+m.join(', ')+')', t};
}

// Число по-русски: пробелы между разрядами, запятая, длинный минус «−»; d — знаков после запятой
const fmtNum=(x,d=0)=>{ const s=x.toLocaleString('ru-RU',{minimumFractionDigits:d,maximumFractionDigits:d}); return /^-[0,\s ]*$/.test(s)?s.slice(1):s.replace('-','−'); };

// Сравнение коэффициента с нормативом из LIQ/FIN («0,15–0,2», «≥ 0,5», «≤ 0,67», «индивидуально»).
// Сравниваем уже округленное значение (d знаков) — как оно показано в таблице.
// Возвращает {ok, text} или null, если норматива нет.
function normCheck(x, norm, d=2){
  if (x==null || !Number.isFinite(x)) return null;
  const num = s => parseFloat(s.replace(',','.'));
  const r = Math.round(x*10**d)/10**d;
  let m;
  if ((m = norm.match(/^([\d,]+)\s*[–-]\s*([\d,]+)$/))){
    const lo=num(m[1]), hi=num(m[2]);
    if (r<lo) return {ok:false, text:`ниже рекомендуемого интервала ${norm}`};
    if (r>hi) return {ok:false, text:`выше рекомендуемого интервала ${norm}`};
    return {ok:true, text:`находится в пределах рекомендуемого интервала ${norm}`};
  }
  if ((m = norm.match(/^≥\s*([\d,]+)$/))){
    return r<num(m[1]) ? {ok:false, text:`ниже рекомендуемого значения ${m[1]}`} : {ok:true, text:`не ниже рекомендуемого значения ${m[1]}`};
  }
  if ((m = norm.match(/^≤\s*([\d,]+)$/))){
    return r>num(m[1]) ? {ok:false, text:`выше рекомендуемого значения ${m[1]}`} : {ok:true, text:`не выше рекомендуемого значения ${m[1]}`};
  }
  return null; // «индивидуально» — норматива нет
}
// Коэффициенты с собственным капиталом в знаменателе: при СК ≤ 0 формальное сравнение с нормативом
// дает бессмыслицу (задолженности −13,38 «не выше 0,67»; маневренности: минус на минус = плюс),
// поэтому при отрицательном капитале они считаются не соответствующими нормативу
const SK_DENOM=['задолженности','маневренности собственных оборотных средств'];
function ratioCheck(name, fn, norm, d, b){
  if (SK_DENOM.includes(name) && b.SK<=0) return {ok:false, text:'не соответствует нормативу, так как собственный капитал отрицательный'};
  return normCheck(fn(b), norm, d);
}

/* ===================== Модель таблиц ===================== */
// ячейка: {t:текст} | {v:число,k:'num'|'pct'|'r',d:знаки}; заголовки: {t,cs,rs}
const N = (val,k='num',d=0)=>({v:val,k,d});
const T = t=>({t});
function fmt(c){
  if (c.t!=null) return c.t;
  if (c.v==null || !Number.isFinite(c.v)) return '—';
  const s = c.k==='num' ? Math.round(c.v).toLocaleString('ru-RU') : c.v.toLocaleString('ru-RU',{minimumFractionDigits:c.d,maximumFractionDigits:c.d});
  return /^-[0,\s\u00a0]*$/.test(s) ? s.slice(1) : s.replace('-','−');
}
const visible = (lines, getVals) => lines.filter(l => {
  if (l.h || l.t || !document.getElementById('hideZero').checked) return true;
  return getVals(l).some(x => x!=null && Math.round(x)!==0);
});
function yearsHas(stmt){ return state.years.map((_,i)=>has(stmt,i)); }

function tblHorizontal(id, num, title, lines, first){
  const n=nYears(), P=pairs(), hv=yearsHas('bal');
  const head=[
    [{t:first,rs:2},...state.years.map(y=>({t:y+' г.',rs:2})),{t:'Абсолютное отклонение, тыс. руб.',cs:P.length},{t:'Относительное отклонение, %',cs:P.length}],
    [...P.map(i=>({t:Y(i)+'/'+Y(i-1)})),...P.map(i=>({t:Y(i)+'/'+Y(i-1)}))]
  ];
  const vals = (l,i)=> hv[i] ? v(l.c,i) : null;
  const rows = visible(lines, l=>state.years.map((_,i)=>vals(l,i))).map(l=>{
    if (l.h) return {sec:true, cells:[T(l.h)]};
    const x = state.years.map((_,i)=>vals(l,i));
    return {t:l.t, cells:[T(l.n), ...x.map(a=>N(a)),
      ...P.map(i=> N(x[i]!=null&&x[i-1]!=null ? x[i]-x[i-1] : null)),
      ...P.map(i=> N(x[i]!=null&&x[i-1] ? x[i]/x[i-1]*100 : null,'pct',2))]};
  });
  return {id,num,title,head,rows,ncols:1+n+2*P.length};
}
function tblVertical(id, num, title, lines, first, totalCode){
  const n=nYears(), P=pairs(), hv=yearsHas('bal');
  const head=[
    [{t:first,rs:2},...state.years.map(y=>({t:'31.12.'+y,cs:2})),{t:'Изменение удельного веса, п. п.',cs:P.length}],
    [...state.years.flatMap(()=>[{t:'сумма, тыс. руб.'},{t:'уд. вес, %'}]),...P.map(i=>({t:Y(i)+'/'+Y(i-1)}))]
  ];
  const vals=(l,i)=> hv[i] ? v(l.c,i) : null;
  const rows = visible(lines, l=>state.years.map((_,i)=>vals(l,i))).map(l=>{
    if (l.h) return {sec:true, cells:[T(l.h)]};
    const x=state.years.map((_,i)=>vals(l,i));
    const sh=state.years.map((_,i)=> x[i]!=null ? div(x[i], hv[i]?v(totalCode,i):null) : null).map(a=>a==null?null:a*100);
    return {t:l.t, cells:[T(l.n), ...x.flatMap((a,i)=>[N(a),N(sh[i],'pct',2)]),
      ...P.map(i=>N(sh[i]!=null&&sh[i-1]!=null ? sh[i]-sh[i-1] : null,'pct',2))]};
  });
  return {id,num,title,head,rows,ncols:1+2*n+P.length};
}
function tblDebtCred(){
  const hv=yearsHas('bal');
  const head=[[{t:'Показатели'},{t:'Ед. изм.'},...state.years.map(y=>({t:'на 31.12.'+y}))]];
  const dz=state.years.map((_,i)=>hv[i]?v('1230',i):null), kz=state.years.map((_,i)=>hv[i]?v('1520',i):null);
  const rows=[
    {cells:[T('Остаток дебиторской задолженности'),T('тыс. руб.'),...dz.map(a=>N(a))]},
    {cells:[T('Остаток кредиторской задолженности'),T('тыс. руб.'),...kz.map(a=>N(a))]},
    {cells:[T('Разница между кредиторской и дебиторской задолженностью'),T('тыс. руб.'),...kz.map((a,i)=>N(a!=null&&dz[i]!=null?a-dz[i]:null))]},
    {cells:[T('Приходится кредиторской задолженности на 1 рубль дебиторской'),T('руб.'),...kz.map((a,i)=>N(a!=null?div(a,dz[i]):null,'r',2))]},
  ];
  return {id:'t26',num:'2.6',title:'Соотношение дебиторской и кредиторской задолженности',head,rows,ncols:2+nYears()};
}
function tblProfit(){
  const P=pairs(), hp=yearsHas('pl');
  const head=[
    [{t:'Наименование',rs:2},...state.years.map(y=>({t:y+' год',rs:2})),{t:'Отклонение (влияние на сумму прибыли), тыс. руб.',cs:P.length}],
    [...P.map(i=>({t:Y(i)+' г.'}))]
  ];
  const vals=(l,i)=> hp[i] ? v(l.c,i) : null;
  const rows=visible(LINES.pl, l=>state.years.map((_,i)=>vals(l,i))).map(l=>{
    const x=state.years.map((_,i)=>vals(l,i));
    return {t:l.t, sub:l.sub, cells:[T(l.n),...x.map(a=>N(a)),...P.map(i=>N(x[i]!=null&&x[i-1]!=null?x[i]-x[i-1]:null))]};
  });
  return {id:'t27',num:'2.7',title:'Анализ формирования прибыли',head,rows,ncols:1+nYears()+P.length};
}
function tblRatios(id,num,title,defs){
  const P=pairs(), hv=yearsHas('bal');
  const head=[[{t:'Коэффициенты'},{t:'Норматив'},...state.years.map(y=>({t:String(y)})),...P.map(i=>({t:Y(i)+'–'+Y(i-1)}))]];
  const rows=defs.map(([name,norm,fn,d])=>{
    const x=state.years.map((_,i)=>hv[i]?fn(base(i)):null);
    return {cells:[T(name),T(norm),...x.map(a=>N(a,'r',d)),...P.map(i=>N(x[i]!=null&&x[i-1]!=null?x[i]-x[i-1]:null,'r',d))]};
  });
  return {id,num,title,head,rows,ncols:2+nYears()+P.length};
}
function tblStability(){
  const hv=yearsHas('bal');
  const B=state.years.map((_,i)=>hv[i]?base(i):null);
  const head=[[{t:'Показатель'},...state.years.map(y=>({t:'на 31.12.'+y}))]];
  const r=(name,key)=>({cells:[T(name),...B.map(b=>N(b?b[key]:null))]});
  const rows=[
    r('Собственный капитал (СК)','SK'), r('Внеоборотные активы (ВОА)','VOA'),
    {t:1,cells:[T('Собственный оборотный капитал, СОК = СК − ВОА'),...B.map(b=>N(b?b.SOK:null))]},
    r('Долгосрочные кредиты и займы (ДКЗ)','DKZ'),
    {t:1,cells:[T('Собственные и долгосрочные источники, СДИ = СОК + ДКЗ'),...B.map(b=>N(b?b.SDI:null))]},
    r('Краткосрочные кредиты и займы (ККЗ)','KKZ'),
    {t:1,cells:[T('Общая величина источников, ОИ = СДИ + ККЗ'),...B.map(b=>N(b?b.OI:null))]},
    r('Запасы (З)','Z'),
    {cells:[T('Излишек (+) / недостаток (−) СОК, ΔСОК = СОК − З'),...B.map(b=>N(b?b.dSOK:null))]},
    {cells:[T('Излишек (+) / недостаток (−) СДИ, ΔСДИ = СДИ − З'),...B.map(b=>N(b?b.dSDI:null))]},
    {cells:[T('Излишек (+) / недостаток (−) ОИ, ΔОИ = ОИ − З'),...B.map(b=>N(b?b.dOI:null))]},
    {t:1,cells:[T('Трехмерная модель'),...B.map(b=>T(b?stabType(b).m:'—'))]},
    {t:2,cells:[T('Тип финансовой устойчивости'),...B.map(b=>T(b?stabType(b).t:'—'))]},
  ];
  return {id:'t210',num:'2.10',title:'Абсолютные показатели и тип финансовой устойчивости',head,rows,ncols:1+nYears()};
}
function tblLiqGroups(){
  const hv=yearsHas('bal');
  const B=state.years.map((_,i)=>hv[i]?base(i):null);
  const head=[[{t:'Группа'},...state.years.map(y=>({t:'на 31.12.'+y}))]];
  const r=(name,key,t)=>({t,cells:[T(name),...B.map(b=>N(b?b[key]:null))]});
  const cmp=(name,a,p,ge)=>({cells:[T(name),...B.map(b=>{ if(!b) return T('—'); const ok= ge? b[a]>=b[p] : b[a]<=b[p]; return {t: ok?'выполняется':'не выполняется', cls: ok?'ok':'bad'}; })]});
  const rows=[
    {sec:true,cells:[T('Активы')]},
    r('А1 — наиболее ликвидные (ДС + КФВ)','A1'), r('А2 — быстрореализуемые (ДЗ + прочие ОА)','A2'),
    r('А3 — медленнореализуемые (запасы + НДС + активы к продаже)','A3'), r('А4 — труднореализуемые (ВОА)','A4'),
    {sec:true,cells:[T('Пассивы')]},
    r('П1 — наиболее срочные (КЗ + прочие краткосрочные)','P1'), r('П2 — краткосрочные (краткосрочные займы)','P2'),
    r('П3 — долгосрочные (раздел IV)','P3'), r('П4 — постоянные (раздел III + ДБП + оценочные)','P4'),
    {sec:true,cells:[T('Излишек (+) / недостаток (−)')]},
    ...[1,2,3,4].map(k=>({cells:[T(`А${k} − П${k}`),...B.map(b=>N(b?b['A'+k]-b['P'+k]:null))]})),
    {sec:true,cells:[T('Условия абсолютной ликвидности')]},
    cmp('А1 ≥ П1','A1','P1',true), cmp('А2 ≥ П2','A2','P2',true), cmp('А3 ≥ П3','A3','P3',true), cmp('А4 ≤ П4','A4','P4',false),
    {t:2,cells:[T('Вывод'),...B.map(b=>T(!b?'—':(b.A1>=b.P1&&b.A2>=b.P2&&b.A3>=b.P3&&b.A4<=b.P4)?'баланс абсолютно ликвиден':'баланс не является абсолютно ликвидным'))]},
  ];
  return {id:'t211',num:'2.11',title:'Группировка активов и пассивов по степени ликвидности',head,rows,ncols:1+nYears()};
}
function buildTables(){
  const yr = state.years[0]+' – '+state.years[nYears()-1]+' гг.';
  return [
    tblHorizontal('t22','2.2','Горизонтальный анализ актива баланса за '+yr, LINES.asset, 'Актив'),
    tblHorizontal('t23','2.3','Горизонтальный анализ пассива баланса за '+yr, LINES.liab, 'Пассив'),
    tblVertical('t24','2.4','Вертикальный анализ актива баланса за '+yr, LINES.asset, 'Актив','1600'),
    tblVertical('t25','2.5','Вертикальный анализ пассива баланса за '+yr, LINES.liab, 'Пассив','1700'),
    tblDebtCred(),
    tblProfit(),
    tblRatios('t28','2.8','Сравнительная таблица коэффициентов ликвидности',LIQ),
    tblRatios('t29','2.9','Сравнительная таблица показателей финансовой устойчивости',FIN),
    tblStability(),
    tblLiqGroups(),
  ];
}

/* ===================== Раскладка ячеек (для HTML / Excel / Word) ===================== */
function allRows(t){
  const rows = t.head.map(r=>r.map(c=>({...c, hd:true})));
  for (const r of t.rows){
    if (r.sec) rows.push([{...r.cells[0], cs:t.ncols, sec:true}]);
    else rows.push(r.cells.map(c=>({...c, tt:r.t, sub:r.sub})));
  }
  return rows;
}
function layout(rows, ncols){
  const grid = rows.map(()=>Array(ncols).fill(null));
  rows.forEach((row,r)=>{
    let c=0;
    for (const cell of row){
      while (c<ncols && grid[r][c]) c++;
      const cs=cell.cs||1, rs=cell.rs||1;
      for (let dr=0; dr<rs; dr++) for (let dc=0; dc<cs; dc++)
        if (grid[r+dr]) grid[r+dr][c+dc] = {cell, r0:r, c0:c};
      c+=cs;
    }
  });
  return grid;
}
