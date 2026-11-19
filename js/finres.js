/* Вкладка 2 «Анализ финансовых результатов»: таблицы 1–4, выводы и резервы, выгрузка в Excel.
   Данные те же, что на первой вкладке (state.values, коды строк форм 1 и 2).
   Расходы приводятся к единому виду: положительное число = расход (даже если введено в скобках или с минусом).
   Пустые строки считаются равными 0 — об этом пишется пометка над таблицами. */

// Строки таблицы 1 (код, название, итоговая строка: 1 — промежуточный итог, 2 — чистая прибыль)
const FR_ROWS = [
  {c:'2110', n:'Выручка'},
  {c:'2120', n:'Себестоимость продаж'},
  {c:'2100', n:'Валовая прибыль (убыток)', t:1},
  {c:'2210', n:'Коммерческие расходы'},
  {c:'2220', n:'Управленческие расходы'},
  {c:'2200', n:'Прибыль (убыток) от продаж', t:1},
  {c:'2310', n:'Доходы от участия в других организациях'},
  {c:'2320', n:'Проценты к получению'},
  {c:'2330', n:'Проценты к уплате'},
  {c:'2340', n:'Прочие доходы'},
  {c:'2350', n:'Прочие расходы'},
  {c:'2300', n:'Прибыль (убыток) до налогообложения', t:1},
  {c:'2410', n:'Налог на прибыль (текущий и отложенный)'},
  {c:'2400', n:'Чистая прибыль (убыток)', t:2},
];
const FR_EXP = ['2120','2210','2220','2330','2350'];          // расходные строки
const FR_TOTALS = ['2100','2200','2300','2400'];               // итоговые строки, которые бывают «по отчету»
const FR_NAME = Object.fromEntries([...FR_ROWS.map(r=>[r.c,r.n]), ['2411','Текущий налог на прибыль'], ['2412','Отложенный налог на прибыль'], ['2460','Прочее']]);

/* ===================== Данные за год ===================== */
// Возвращает null, если отчета о финрезультатах за год нет совсем
function frData(i){
  if (!has('pl', i)) return null;
  const raw = c => { const a = state.values[c]; return a ? a[i] ?? null : null; };
  const missing = [], flipped = [];
  const inp = c => {
    const x = raw(c);
    if (x == null){ missing.push(c); return 0; }
    if (FR_EXP.includes(c) && x < 0){ flipped.push(c); return -x; }   // расход в скобках / с минусом → положительный
    return x;
  };
  const d = {B:inp('2110'), C:inp('2120'), KR:inp('2210'), UR:inp('2220'),
    o2310:inp('2310'), o2320:inp('2320'), o2330:inp('2330'), o2340:inp('2340'), o2350:inp('2350')};
  // налог на прибыль 2410 = текущий (2411) + отложенный (2412); «+» — расход, «−» — доход
  const t1 = raw('2411'), t2 = raw('2412');
  if (t1 == null && t2 == null) missing.push('2410');
  d.TAX = nz(t1) + nz(t2);
  d.P2460 = nz(raw('2460'));
  // пересчет по формулам задания
  d.VP  = d.B - d.C;
  d.PP  = d.VP - d.KR - d.UR;
  d.PN  = d.PP + d.o2310 + d.o2320 - d.o2330 + d.o2340 - d.o2350;
  d.PCH = d.PN - d.TAX;
  // строки «по отчету»: итог из загруженного отчета, если он сохранен, иначе — итог формы по введенным строкам
  const form = {'2100':d.VP, '2200':d.PP, '2300':d.PN, '2400':d.PN - d.TAX + d.P2460};
  d.R = {}; d.src = {};
  for (const c of FR_TOTALS){
    const r = state.reported && state.reported[c] ? state.reported[c][Y(i)] : null;
    d.R[c] = r != null ? r : form[c]; d.src[c] = r != null ? 'rep' : 'form';
  }
  d.missing = missing; d.flipped = flipped;
  return d;
}
function frAll(){ return state.years.map((_, i) => frData(i)); }
// значение строки таблицы 1
function frVal(d, c){
  if (!d) return null;
  return ({'2110':d.B, '2120':d.C, '2210':d.KR, '2220':d.UR, '2310':d.o2310, '2320':d.o2320, '2330':d.o2330,
    '2340':d.o2340, '2350':d.o2350, '2410':d.TAX})[c] ?? d.R[c];
}
// баланс на 31.12 года i (null, если баланса за год нет)
const balAt = (code, i) => i >= 0 && i < nYears() && has('bal', i) ? v(code, i) : null;
// средняя величина за год i: (на начало + на конец) / 2; начало года = конец предыдущего года
function avgBal(code, i){
  const end = balAt(code, i); if (end == null) return {x:null, how:'none'};
  const beg = balAt(code, i - 1);
  return beg == null ? {x:end, how:'end'} : {x:(beg + end) / 2, how:'avg'};
}

const Dv = (val, k='num', d=0) => ({...N(val, k, d), dev:true});   // отклонение: «+» — зеленым, «−» — красным
const pct = (a, b) => (b ? a / b * 100 : null);
const pairLabel = i => Y(i) + '/' + Y(i - 1);

/* ===================== Таблица 1. Динамика и структура ===================== */
function frT1(D){
  const n = nYears(), P = pairs();
  const head = [
    [{t:'Показатель', rs:2}, {t:'Код', rs:2}, {t:'Сумма, тыс. руб.', cs:n}, {t:'Абсолютное отклонение, тыс. руб.', cs:P.length},
     {t:'Темп роста, %', cs:P.length}, {t:'Доля в выручке, %', cs:n}],
    [...state.years.map(y => ({t:y + ' г.'})), ...P.map(i => ({t:pairLabel(i)})), ...P.map(i => ({t:pairLabel(i)})), ...state.years.map(y => ({t:y + ' г.'}))],
  ];
  let signNote = false;
  const rows = FR_ROWS.map(r => {
    const x = D.map(d => frVal(d, r.c));
    const growth = P.map(i => {
      const a = x[i - 1], b = x[i];
      if (a == null || b == null || a === 0) return null;
      if ((a < 0) !== (b < 0)) { signNote = true; return null; }   // смена знака (прибыль ↔ убыток) — темп роста не имеет смысла
      return b / a * 100;
    });
    return {t:r.t, cells:[T(r.n), T(r.c), ...x.map(a => N(a)),
      ...P.map(i => Dv(x[i] != null && x[i - 1] != null ? x[i] - x[i - 1] : null)),
      ...growth.map(g => N(g, 'pct', 1)),
      ...x.map((a, i) => N(a != null && D[i] ? pct(a, D[i].B) : null, 'pct', 2))]};
  });
  const notes = [];
  if (signNote) notes.push('Темп роста не рассчитывается, если показатель сменил знак (прибыль ↔ убыток) или база равна нулю — стоит «—».');
  if (D.some(d => d && d.TAX < 0)) notes.push('Налог на прибыль со знаком «−» — это доход (отложенный налог уменьшил расход по налогу).');
  return {id:'fr1', num:'1', title:'Динамика и структура финансовых результатов', head, rows, ncols:2 + 2 * n + 2 * P.length,
    formula:'Абсолютное отклонение = Xᵢ − Xᵢ₋₁; темп роста = Xᵢ / Xᵢ₋₁ × 100%; доля в выручке = X / Выручка (стр. 2110) × 100%. Расходы и налог показаны положительными числами (расход).',
    notes};
}

/* ===================== Таблица 2. Проверка цепочки формирования прибыли ===================== */
function frT2(D){
  const n = nYears();
  const head = [
    [{t:'Показатель', rs:2}, {t:'Формула расчета', rs:2}, ...state.years.map(y => ({t:y + ' г.', cs:3}))],
    state.years.flatMap(() => [{t:'расчет'}, {t:'по отчету'}, {t:'расхождение'}]),
  ];
  const defs = [
    ['Валовая прибыль (ВП), стр. 2100', 'В − С', 'VP', '2100', 1],
    ['Прибыль от продаж (Пп), стр. 2200', 'ВП − КР − УР', 'PP', '2200', 1],
    ['Прибыль до налогообложения (Пнобл), стр. 2300', 'Пп + 2310 + 2320 − 2330 + 2340 − 2350', 'PN', '2300', 1],
    ['Чистая прибыль (ПЧ), стр. 2400', 'Пнобл − 2410', 'PCH', '2400', 2],
  ];
  const rows = defs.map(([name, f, key, code, t]) => ({t, cells:[T(name), T(f), ...D.flatMap(d => {
    if (!d) return [N(null), N(null), N(null)];
    const diff = d[key] - d.R[code];
    return [N(d[key]), N(d.R[code]), {...N(diff), x: Math.abs(diff) >= 1 ? 'warn' : 'okv'}];
  })]}));
  const notes = [];
  // расхождение по ПЧ ровно на строку 2460 «Прочее» — объясняем
  const by2460 = D.map((d, i) => d && Math.abs(d.PCH - d.R['2400']) >= 1 && Math.abs(d.PCH - d.R['2400'] + d.P2460) < 1 ? i : -1).filter(i => i >= 0);
  if (by2460.length) notes.push(`Расхождение по чистой прибыли за ${by2460.map(Y).join(', ')} г. равно строке 2460 «Прочее» (${by2460.map(i => fmt(N(D[i].P2460))).join('; ')} тыс. руб.): в отчете ПЧ = Пнобл − 2410 + 2460, а формула задания эту строку не учитывает.`);
  const formOnly = FR_TOTALS.filter(c => D.some(d => d && d.src[c] === 'form'));
  if (formOnly.length) notes.push(`Строки ${formOnly.join(', ')} из самого отчета не загружены — в колонке «по отчету» для них итог формы по введенным строкам (ПЧ — с учетом стр. 2460). Итоги из PDF или Excel подставляются сюда автоматически.`);
  return {id:'fr2', num:'2', title:'Проверка цепочки формирования прибыли', head, rows, ncols:2 + 3 * n,
    formula:'ВП = В − С; Пп = ВП − КР − УР; Пнобл = Пп + 2310 + 2320 − 2330 + 2340 − 2350; ПЧ = Пнобл − 2410 (текущий + отложенный налог); расхождение = расчет − по отчету.',
    notes};
}

/* ===================== Таблица 3. Факторный анализ прибыли от продаж ===================== */
function frFactors(d0, d1){
  if (!d0 || !d1 || !d0.B || !d1.B) return null;   // защита от деления на ноль
  const fB  = (d1.B - d0.B) * d0.PP / d0.B;
  const fC  = -(d1.C / d1.B - d0.C / d0.B) * d1.B;
  const fKR = -(d1.KR / d1.B - d0.KR / d0.B) * d1.B;
  const fUR = -(d1.UR / d1.B - d0.UR / d0.B) * d1.B;
  const sum = fB + fC + fKR + fUR, dPP = d1.PP - d0.PP;
  return {fB, fC, fKR, fUR, sum, dPP, check:sum - dPP};
}
const FACTOR_NAMES = {fB:'изменение выручки', fC:'изменение уровня себестоимости', fKR:'изменение уровня коммерческих расходов', fUR:'изменение уровня управленческих расходов'};
function frT3(D){
  const P = pairs();
  const F = P.map(i => frFactors(D[i - 1], D[i]));
  const head = [[{t:'Показатель'}, ...P.map(i => ({t:pairLabel(i)}))]];
  const r = (name, fn, o={}) => ({t:o.t, cells:[T(name), ...P.map((i, k) => {
    const val = fn(k, i); return o.dev ? Dv(val) : o.chk ? {...N(val), x: val != null && Math.abs(val) < 1 ? 'okv' : 'warn'} : N(val);
  })]});
  const g = (k, key) => F[k] ? F[k][key] : null;
  const rows = [
    r('Выручка прошлого года (В₀)', (k, i) => D[i - 1] ? D[i - 1].B : null),
    r('Выручка отчетного года (В₁)', (k, i) => D[i] ? D[i].B : null),
    r('Прибыль от продаж прошлого года (Пп₀)', (k, i) => D[i - 1] ? D[i - 1].PP : null),
    r('Прибыль от продаж отчетного года (Пп₁)', (k, i) => D[i] ? D[i].PP : null),
    r('Изменение прибыли от продаж (Пп₁ − Пп₀)', k => g(k, 'dPP'), {dev:1, t:1}),
    {sec:true, cells:[T('Влияние факторов на изменение прибыли от продаж, тыс. руб.')]},
    r('Влияние выручки', k => g(k, 'fB'), {dev:1}),
    r('Влияние себестоимости', k => g(k, 'fC'), {dev:1}),
    r('Влияние коммерческих расходов', k => g(k, 'fKR'), {dev:1}),
    r('Влияние управленческих расходов', k => g(k, 'fUR'), {dev:1}),
    r('Итого влияние факторов', k => g(k, 'sum'), {dev:1, t:2}),
    r('Проверка: итог − (Пп₁ − Пп₀), должно быть 0', k => g(k, 'check'), {chk:1}),
  ];
  const notes = [];
  const bad = P.filter((i, k) => !F[k] && D[i] && D[i - 1]);
  if (bad.length) notes.push(`Для ${bad.map(pairLabel).join(', ')} факторы не рассчитаны: выручка равна нулю (деление на ноль).`);
  notes.push('Пп здесь — прибыль от продаж, пересчитанная по строкам (В − С − КР − УР), поэтому сумма влияний совпадает с изменением прибыли.');
  return {id:'fr3', num:'3', title:'Факторный анализ прибыли от продаж', head, rows, ncols:1 + P.length, F,
    formula:'Влияние выручки = (В₁ − В₀) × Пп₀ / В₀; себестоимости = −(С₁/В₁ − С₀/В₀) × В₁; коммерческих расходов = −(КР₁/В₁ − КР₀/В₀) × В₁; управленческих расходов = −(УР₁/В₁ − УР₀/В₀) × В₁; итог = сумма влияний = Пп₁ − Пп₀.',
    notes};
}

/* ===================== Таблица 4. Рентабельность (8 показателей) ===================== */
function frRent(D){
  return D.map((d, i) => {
    if (!d) return null;
    // средние величины за год: (на начало + на конец) / 2
    const A = avgBal('1600', i), K = avgBal('1300', i), OA = avgBal('1200', i), VOA = avgBal('1100', i);
    const NP = d.R['2400'];
    return {
      roe:      K.x > 0 ? NP / K.x * 100 : null,          // при отрицательном капитале ROE не имеет смысла
      roa:      A.x ? NP / A.x * 100 : null,
      roca:     OA.x ? NP / OA.x * 100 : null,
      rnca:     VOA.x ? NP / VOA.x * 100 : null,
      prod:     d.C ? d.R['2200'] / d.C * 100 : null,
      netSales: d.B ? NP / d.B * 100 : null,
      sales:    d.B ? d.R['2200'] / d.B * 100 : null,
      roaEbt:   A.x ? d.R['2300'] / A.x * 100 : null,
      A, K, OA, VOA,
    };
  });
}
const RENT_DEFS = [
  ['roe',      'Рентабельность собственного капитала (ROE)', 'ЧП / ср. СК (1300)'],
  ['roa',      'Рентабельность авансированного капитала (всех активов, ROA)', 'ЧП / ср. активы (1600)'],
  ['roca',     'Рентабельность оборотных активов', 'ЧП / ср. оборотные активы (1200)'],
  ['rnca',     'Рентабельность внеоборотных активов', 'ЧП / ср. внеоборотные активы (1100)'],
  ['prod',     'Рентабельность производства', '2200 / 2120'],
  ['netSales', 'Рентабельность продаж по чистой прибыли', '2400 / 2110'],
  ['sales',    'Рентабельность продаж по прибыли от продаж', '2200 / 2110'],
  ['roaEbt',   'Рентабельность активов по прибыли до налогообложения', '2300 / ср. активы (1600)'],
];
function frT4(D){
  const P = pairs(), R = frRent(D);
  const head = [
    [{t:'Показатель', rs:2}, {t:'Формула', rs:2}, {t:'Значение, %', cs:nYears()}, {t:'Отклонение, п. п.', cs:P.length}],
    [...state.years.map(y => ({t:y + ' г.'})), ...P.map(i => ({t:pairLabel(i)}))],
  ];
  const rows = RENT_DEFS.map(([k, name, f]) => {
    const x = R.map(r => r ? r[k] : null);
    return {cells:[T(name), T(f), ...x.map(a => N(a, 'pct', 2)), ...P.map(i => Dv(x[i] != null && x[i - 1] != null ? x[i] - x[i - 1] : null, 'pct', 2))]};
  });
  const notes = [];
  const endOnly = R.map((r, i) => r && r.A.how === 'end' ? i : -1).filter(i => i >= 0);
  if (endOnly.length) notes.push(`За ${endOnly.map(Y).join(', ')} г. нет баланса на начало года (31.12 прошлого года) — средние величины активов, капитала, оборотных и внеоборотных активов взяты на конец года.`);
  const noBal = R.map((r, i) => r && r.A.how === 'none' ? i : -1).filter(i => i >= 0);
  if (noBal.length) notes.push(`За ${noBal.map(Y).join(', ')} г. нет баланса — показатели по активам и капиталу не рассчитаны.`);
  const negK = R.map((r, i) => r && r.K.x != null && r.K.x <= 0 ? i : -1).filter(i => i >= 0);
  if (negK.length) notes.push(`За ${negK.map(Y).join(', ')} г. собственный капитал отрицательный — ROE не рассчитывается.`);
  return {id:'fr4', num:'4', title:'Показатели рентабельности, %', head, rows, ncols:2 + nYears() + P.length, R,
    formula:'ROE = ЧП / ср. СК; ROA = ЧП / ср. активы; Rоа = ЧП / ср. оборотные активы; Rвна = ЧП / ср. внеоборотные активы; Rпроизв = Пп / С; Rпрод(ЧП) = ЧП / В; Rпрод(Пп) = Пп / В; Rа(Пнобл) = Пнобл / ср. активы; средняя = (на начало года + на конец года) / 2; всё × 100%, отклонение — в п. п.',
    notes};
}

/* ===================== Общий вид таблицы «по парам лет» (таблицы 5 и 7) ===================== */
// spec: [название, fn(j, i) → число, {dev, chk, t, k, d, tol}] или {sec:'текст'}
function pairTable(id, num, title, spec, formula, notes, extra={}){
  const P = pairs();
  const head = [[{t:'Показатель'}, ...P.map(i => ({t:pairLabel(i)}))]];
  const rows = spec.map(r => {
    if (r.sec) return {sec:true, cells:[T(r.sec)]};
    const [name, fn, o={}] = r, k = o.k || 'num', d = o.d || 0;
    return {t:o.t, cells:[T(name), ...P.map((i, j) => {
      const val = fn(j, i);
      if (o.chk) return {...N(val, k, d), x: val != null && Math.abs(val) < (o.tol || 1) ? 'okv' : 'warn'};
      return o.dev ? Dv(val, k, d) : N(val, k, d);
    })]};
  });
  return {id, num, title, head, rows, ncols:1 + P.length, formula, notes, ...extra};
}

/* ===================== Таблица 5. Факторы изменения чистой прибыли ===================== */
function frNpFactors(a, b){
  if (!a || !b) return null;
  const fin = d => d.o2310 + d.o2320 - d.o2330, oth = d => d.o2340 - d.o2350;
  const f = {
    fPP:   b.R['2200'] - a.R['2200'],
    fFin:  fin(b) - fin(a),
    fOth:  oth(b) - oth(a),
    fTax:  -(b.TAX - a.TAX),          // рост налога уменьшает прибыль — знак обратный
    f2460: b.P2460 - a.P2460,
  };
  f.sum = f.fPP + f.fFin + f.fOth + f.fTax + f.f2460;
  f.dNP = b.R['2400'] - a.R['2400'];
  f.check = f.sum - f.dNP;
  return f;
}
const NP_FACTOR_NAMES = {fPP:'изменение прибыли от продаж', fFin:'изменение сальдо доходов от участия и процентов', fOth:'изменение сальдо прочих доходов и расходов', fTax:'изменение налога на прибыль', f2460:'изменение прочих статей (стр. 2460)'};
function frT5(D){
  const F = pairs().map(i => frNpFactors(D[i - 1], D[i]));
  const g = (j, key) => F[j] ? F[j][key] : null;
  const spec = [
    ['Чистая прибыль прошлого года (ЧП₀)', (j, i) => D[i - 1] ? D[i - 1].R['2400'] : null],
    ['Чистая прибыль отчетного года (ЧП₁)', (j, i) => D[i] ? D[i].R['2400'] : null],
    ['Изменение чистой прибыли (Δ2400)', j => g(j, 'dNP'), {dev:1, t:1}],
    {sec:'Влияние факторов на изменение чистой прибыли, тыс. руб.'},
    ['Изменение прибыли от продаж (Δ2200)', j => g(j, 'fPP'), {dev:1}],
    ['Сальдо доходов от участия и процентов, Δ(2310 + 2320 − 2330)', j => g(j, 'fFin'), {dev:1}],
    ['Сальдо прочих доходов и расходов, Δ(2340 − 2350)', j => g(j, 'fOth'), {dev:1}],
    ['Налог на прибыль, −Δ2410', j => g(j, 'fTax'), {dev:1}],
    ['Прочие статьи, Δ2460', j => g(j, 'f2460'), {dev:1}],
    ['Итого влияние факторов', j => g(j, 'sum'), {dev:1, t:2}],
    ['Проверка: итог − Δ2400, должно быть 0', j => g(j, 'check'), {chk:1}],
  ];
  const notes = ['Строка 2460 «Прочее» добавлена к четырем факторам задания: в отчете ЧП = Пнобл − 2410 + 2460, без нее итог не совпал бы с Δ2400.',
    'Налог на прибыль 2410 учитывается целиком (текущий + отложенный); его рост уменьшает прибыль, поэтому влияние берется с обратным знаком.'];
  return pairTable('fr5', '5', 'Факторы изменения чистой прибыли', spec,
    'Влияние Пп = Δ2200; сальдо участия и процентов = Δ(2310 + 2320 − 2330); сальдо прочих = Δ(2340 − 2350); налога = −Δ2410; прочих статей = Δ2460; итог = сумма влияний = Δ2400.',
    notes, {F});
}

/* ===================== Таблицы 6–7. Модель Дюпона ===================== */
// ROE = (ЧП / В) × (В / ср. активы) × (ср. активы / ср. СК)
function frDupont(D){
  return D.map((d, i) => {
    if (!d) return null;
    const A = avgBal('1600', i), K = avgBal('1300', i);
    if (!d.B || !A.x || !(K.x > 0)) return {a:null, b:null, c:null, roe:null, negK: K.x != null && K.x <= 0};
    const a = d.R['2400'] / d.B * 100, b = d.B / A.x, c = A.x / K.x;
    return {a, b, c, roe:a * b * c};
  });
}
function frT6(D){
  const P = pairs(), M = frDupont(D);
  const head = [
    [{t:'Показатель', rs:2}, {t:'Формула', rs:2}, {t:'Значение', cs:nYears()}, {t:'Отклонение', cs:P.length}],
    [...state.years.map(y => ({t:y + ' г.'})), ...P.map(i => ({t:pairLabel(i)}))],
  ];
  const defs = [
    ['a', 'a — рентабельность продаж по чистой прибыли, %', 'ЧП / В × 100', 'pct', 2],
    ['b', 'b — оборачиваемость активов, раз', 'В / ср. активы', 'r', 3],
    ['c', 'c — мультипликатор капитала, раз', 'ср. активы / ср. СК', 'r', 3],
    ['roe', 'ROE = a × b × c, %', 'рентабельность × оборачиваемость × мультипликатор', 'pct', 2],
  ];
  const rows = defs.map(([key, name, f, k, d]) => {
    const x = M.map(m => m ? m[key] : null);
    return {t:key === 'roe' ? 2 : 0, cells:[T(name), T(f), ...x.map(a => N(a, k, d)), ...P.map(i => Dv(x[i] != null && x[i - 1] != null ? x[i] - x[i - 1] : null, k, d))]};
  });
  const notes = [];
  const neg = M.map((m, i) => m && m.negK ? i : -1).filter(i => i >= 0);
  if (neg.length) notes.push(`За ${neg.map(Y).join(', ')} г. собственный капитал отрицательный — модель Дюпона не рассчитывается.`);
  notes.push('ROE по модели совпадает с ROE из таблицы 4: средние активы в произведении сокращаются. Отклонение по ROE и рентабельности — в п. п., по оборачиваемости и мультипликатору — в разах.');
  return {id:'fr6', num:'6', title:'Модель Дюпона: факторы рентабельности собственного капитала', head, rows, ncols:2 + nYears() + P.length, M,
    formula:'ROE = (ЧП / В) × (В / ср. активы) × (ср. активы / ср. СК) = рентабельность продаж × оборачиваемость активов × мультипликатор капитала.',
    notes};
}
function frDupontChain(m0, m1){
  if (!m0 || !m1 || m0.roe == null || m1.roe == null) return null;
  // цепные подстановки: сначала меняем a, затем b, затем c
  const fa = (m1.a - m0.a) * m0.b * m0.c;
  const fb = m1.a * (m1.b - m0.b) * m0.c;
  const fc = m1.a * m1.b * (m1.c - m0.c);
  const sum = fa + fb + fc, dROE = m1.roe - m0.roe;
  return {fa, fb, fc, sum, dROE, check:sum - dROE};
}
const DUPONT_NAMES = {fa:'рентабельность продаж по чистой прибыли', fb:'оборачиваемость активов', fc:'мультипликатор капитала'};
function frT7(D, t6){
  const M = t6.M, F = pairs().map(i => frDupontChain(M[i - 1], M[i]));
  const g = (j, key) => F[j] ? F[j][key] : null, o = {k:'pct', d:2};
  const spec = [
    ['ROE прошлого года (ROE₀), %', (j, i) => M[i - 1] ? M[i - 1].roe : null, o],
    ['ROE отчетного года (ROE₁), %', (j, i) => M[i] ? M[i].roe : null, o],
    ['Изменение ROE (ΔROE), п. п.', j => g(j, 'dROE'), {...o, dev:1, t:1}],
    {sec:'Влияние факторов (цепные подстановки), п. п.'},
    ['Рентабельности продаж по ЧП: (a₁ − a₀) × b₀ × c₀', j => g(j, 'fa'), {...o, dev:1}],
    ['Оборачиваемости активов: a₁ × (b₁ − b₀) × c₀', j => g(j, 'fb'), {...o, dev:1}],
    ['Мультипликатора капитала: a₁ × b₁ × (c₁ − c₀)', j => g(j, 'fc'), {...o, dev:1}],
    ['Итого влияние факторов', j => g(j, 'sum'), {...o, dev:1, t:2}],
    ['Проверка: итог − ΔROE, должно быть 0', j => g(j, 'check'), {...o, chk:1, tol:0.005}],
  ];
  const notes = [];
  if (F.some(f => !f)) notes.push('Для пар лет, где модель Дюпона не рассчитана хотя бы за один год, разложение не выполняется.');
  return pairTable('fr7', '7', 'Разложение изменения ROE по модели Дюпона (цепные подстановки)', spec,
    'ΔROE(a) = (a₁ − a₀) × b₀ × c₀; ΔROE(b) = a₁ × (b₁ − b₀) × c₀; ΔROE(c) = a₁ × b₁ × (c₁ − c₀); итог = ΔROE(a) + ΔROE(b) + ΔROE(c) = ROE₁ − ROE₀ (a — рентабельность продаж по ЧП, b — оборачиваемость активов, c — мультипликатор капитала).',
    notes, {F});
}

/* ===================== Графики вкладки (Chart.js — та же библиотека и оформление, что на первой вкладке) ===================== */
let frCharts = [], frChartGen = 0;
function frChartData(D, R){
  const idx = state.years.map((_, i) => i).filter(i => D[i]);
  return {
    idx,
    bars:[['Выручка', i => D[i].B], ['Прибыль от продаж', i => D[i].R['2200']], ['Чистая прибыль', i => D[i].R['2400']]],
    lines:[['Рентабельность продаж по Пп', i => R[i].sales], ['Рентабельность продаж по ЧП', i => R[i].netSales],
      ['Рентабельность производства', i => R[i].prod], ['ROA', i => R[i].roa], ['ROE', i => R[i].roe]],
  };
}
function frLineConfig(idx, series, th){
  const pal = PALETTE[th.mode], font = {family:th.font, size:12};
  return {
    type:'line',
    data:{labels:idx.map(i => String(Y(i))), datasets:series.map(([name, fn], k) => ({label:name, data:idx.map(fn),
      borderColor:pal.cat[k], backgroundColor:pal.cat[k], borderWidth:2, pointRadius:4, pointHoverRadius:6, tension:0}))},
    plugins:[bgPlugin],
    options:{responsive:true, maintainAspectRatio:false, animation:false,
      layout:{padding:{top:18, right:12, left:4, bottom:4}},
      scales:{x:{offset:true, grid:{display:false}, border:{color:th.axis}, ticks:{color:th.soft, font}},
        y:{grid:{color:th.grid}, border:{display:false}, ticks:{color:th.soft, font, callback:x => fmtNum(x, 0) + '%'}}},
      plugins:{bgFill:{color:th.surface},
        legend:{position:'bottom', labels:{color:th.ink, font, boxWidth:12, boxHeight:12, padding:14}},
        tooltip:{titleFont:font, bodyFont:font, callbacks:{label:it => `${it.dataset.label}: ${it.raw == null ? '—' : fmtNum(it.raw, 2) + '%'}`}}}},
  };
}
async function drawFrCharts(){
  const sec = document.getElementById('fr'); if (!sec || sec.hidden) return;   // в скрытом блоке Chart.js рисует нулевой размер
  const gen = ++frChartGen;
  try{ await loadScript(CHART_URL); }
  catch(e){ document.querySelectorAll('#fr .chartbox').forEach(el => el.innerHTML = '<p class="hint">График не загрузился: для Chart.js нужен интернет.</p>'); return; }
  if (gen !== frChartGen) return;
  frCharts.forEach(c => c.destroy()); frCharts = [];
  const D = frAll(), C = frChartData(D, frRent(D)), th = chartTheme(false);
  if (!C.idx.length) return;
  const cb = document.querySelector('#frFigBars canvas'), cl = document.querySelector('#frFigLines canvas');
  if (cb) frCharts.push(new Chart(cb, chartConfig({kind:'bars', unit:'млн руб.', idx:C.idx, series:C.bars}, th, false)));
  if (cl) frCharts.push(new Chart(cl, frLineConfig(C.idx, C.lines, th)));
}

/* ===================== Пометки о данных ===================== */
function frDataNotes(D){
  const notes = [];
  const noPl = state.years.filter((_, i) => !D[i]);
  if (noPl.length) notes.push(`Отчет о финансовых результатах за ${noPl.join(', ')} г. не заполнен — эти годы в расчетах пропущены.`);
  D.forEach((d, i) => {
    if (!d) return;
    if (d.missing.length) notes.push(`${Y(i)} г.: не заполнены строки ${d.missing.map(c => c + ' «' + FR_NAME[c] + '»').join(', ')} — приняты равными 0.`);
    if (d.flipped.length) notes.push(`${Y(i)} г.: расходы по строкам ${d.flipped.join(', ')} были введены со знаком «−» (в скобках) — приведены к положительному виду.`);
  });
  return notes;
}

/* ===================== Выводы и резервы роста прибыли (по правилам, без ИИ) ===================== */
const F0 = x => fmtNum(Math.abs(x), 0);
const F1 = x => fmtNum(Math.abs(x), 1);
const sgn = x => (x > 0 ? '+' : x < 0 ? '−' : '') + F0(x);
// Изменение показателя с учетом смены знака (прибыль ↔ убыток)
function chgText(nm, loss, lossGen, x0, x1, y0, y1){
  const d = x1 - x0;
  if (x0 >= 0 && x1 >= 0){
    const p = x0 ? ` (на ${F1(d / x0 * 100)}%)` : '';
    return `${nm} в ${y1} г. по сравнению с ${y0} г. ${d >= 0 ? 'увеличилась' : 'снизилась'} на ${F0(d)} тыс. руб.${p} — с ${F0(x0)} до ${F0(x1)} тыс. руб.`;
  }
  if (x0 < 0 && x1 < 0)
    return `${loss[0].toUpperCase() + loss.slice(1)} в ${y1} г. ${-x1 < -x0 ? 'сократился' : 'вырос'} на ${F0(d)} тыс. руб. (на ${F1(d / x0 * 100)}%) — с ${F0(x0)} до ${F0(x1)} тыс. руб.`;
  if (x0 >= 0) return `В ${y1} г. вместо прибыли ${F0(x0)} тыс. руб. (${y0} г.) получен ${loss} ${F0(x1)} тыс. руб. — ухудшение на ${F0(d)} тыс. руб.`;
  return `В ${y1} г. получена прибыль ${F0(x1)} тыс. руб. вместо ${lossGen} ${F0(x0)} тыс. руб. в ${y0} г. — улучшение на ${F0(d)} тыс. руб.`;
}
function rentText(name, a, b){
  if (a == null || b == null) return null;
  const d = b - a;
  return `${name} ${Math.abs(d) < 0.005 ? 'не изменилась' : d > 0 ? 'выросла' : 'снизилась'} с ${fmtNum(a, 2)}% до ${fmtNum(b, 2)}%${Math.abs(d) < 0.005 ? '' : ` (на ${fmtNum(Math.abs(d), 2)} п. п.)`}`;
}
function frConclusions(D, t3, t4, t5, t7){
  const out = [];   // [{h, ps:[...]}]
  const P = pairs().filter(i => D[i] && D[i - 1]);
  if (!P.length){ out.push({h:'Выводы', ps:['Для выводов нужен отчет о финансовых результатах минимум за два соседних года.']}); return out; }
  const co = state.company || 'организации';

  // 1. Динамика
  const dyn = [];
  for (const i of P){
    const a = D[i - 1], b = D[i], y0 = Y(i - 1), y1 = Y(i);
    dyn.push([
      chgText('Выручка', '', '', a.B, b.B, y0, y1),
      chgText('Прибыль от продаж', 'убыток от продаж', 'убытка от продаж', a.R['2200'], b.R['2200'], y0, y1),
      chgText('Чистая прибыль', 'чистый убыток', 'чистого убытка', a.R['2400'], b.R['2400'], y0, y1),
    ].join(' '));
  }
  out.push({h:'Динамика финансовых результатов', ps:dyn});

  // 2–4. Факторы прибыли от продаж, чистой прибыли и модель Дюпона — те же тексты, что под таблицами 3, 5 и 7
  const fac = [], npf = [], dp = [];
  pairs().forEach((i, k) => {
    if (t3.F[k]) fac.push(ppFactorText(i, t3.F[k]));
    if (t5.F[k]) npf.push(npFactorText(i, t5.F[k]));
    if (D[i] && D[i - 1]) dp.push(dupontText(i, t7.F[k]));
  });
  if (fac.length) out.push({h:'Влияние факторов на прибыль от продаж', ps:fac});
  if (npf.length) out.push({h:'Влияние факторов на чистую прибыль', ps:npf});
  if (dp.length) out.push({h:'Модель Дюпона', ps:dp});

  // 5. Рентабельность — за последний период
  const L = P.at(-1), r0 = t4.R[L - 1], r1 = t4.R[L];
  const rs = [rentText('рентабельность продаж по прибыли от продаж', r0.sales, r1.sales), rentText('рентабельность продаж по чистой прибыли', r0.netSales, r1.netSales),
    rentText('рентабельность активов (ROA)', r0.roa, r1.roa), rentText('рентабельность собственного капитала (ROE)', r0.roe, r1.roe)].filter(Boolean);
  if (rs.length) out.push({h:'Рентабельность', ps:[`В ${Y(L)} г. по сравнению с ${Y(L - 1)} г. ` + rs.join('; ') + '.']});

  // 6. Резервы роста прибыли — за последний период
  const a = D[L - 1], b = D[L], res = []; let total = 0;
  if (a.B && b.B){
    for (const [key, name] of [['C', 'себестоимости'], ['KR', 'коммерческих расходов'], ['UR', 'управленческих расходов']]){
      const s0 = a[key] / a.B, s1 = b[key] / b.B;
      if (s1 > s0){
        const x = (s1 - s0) * b.B; total += x;
        res.push(`Резерв: снижение ${name} до уровня прошлого года (доля в выручке ${fmtNum(s0 * 100, 2)}% вместо ${fmtNum(s1 * 100, 2)}%) = ${F0(x)} тыс. руб.`);
      }
    }
  }
  if (!res.length) res.push(`Доли себестоимости, коммерческих и управленческих расходов в выручке в ${Y(L)} г. не выросли — резервов за счет снижения их уровня до прошлого года нет.`);
  else if (res.length > 1) res.push(`Итого резервы роста прибыли от продаж: ${F0(total)} тыс. руб.`);
  if (b.o2350 > b.o2340) res.push(`Зона потерь: прочие расходы (${F0(b.o2350)} тыс. руб.) превышают прочие доходы (${F0(b.o2340)} тыс. руб.) на ${F0(b.o2350 - b.o2340)} тыс. руб. — сокращение этого разрыва увеличит прибыль до налогообложения.`);
  else res.push(`Прочие доходы (${F0(b.o2340)} тыс. руб.) не меньше прочих расходов (${F0(b.o2350)} тыс. руб.) — зоны потерь по прочей деятельности в ${Y(L)} г. нет.`);
  out.push({h:`Резервы роста прибыли ${co} (${Y(L)} г.)`, ps:res});
  return out;
}

/* ===================== Выводы под каждой таблицей (по правилам, без ИИ) ===================== */
const lcFirst = s => s[0].toLowerCase() + s.slice(1);
const P2 = x => fmtNum(x, 2);
const growth = (x0, x1) => (x0 > 0 && x1 >= 0 ? x1 / x0 * 100 : null);
const shareOf = (d, x) => (d.B ? x / d.B * 100 : null);
const yearsWith = D => state.years.map((_, i) => i).filter(i => D[i]);
const FR_NEED2 = 'Для вывода нужен отчет о финансовых результатах минимум за два года.';

// Таблица 1 — динамика и структура
function tcFr1(M){
  const D = M.D, I = yearsWith(D); if (I.length < 2) return FR_NEED2;
  const a = D[I[0]], b = D[I.at(-1)], y0 = Y(I[0]), y1 = Y(I.at(-1)), ps = [];
  const gB = growth(a.B, b.B), gC = growth(a.C, b.C);
  ps.push(`За ${y0}–${y1} гг. выручка ${b.B >= a.B ? 'увеличилась' : 'снизилась'} с ${F0(a.B)} до ${F0(b.B)} тыс. руб.${gB != null ? ` (темп роста ${F1(gB)}%)` : ''}.`);
  if (gB != null && gC != null){
    const s0 = shareOf(a, a.C), s1 = shareOf(b, b.C);
    ps.push(`Себестоимость продаж росла ${gC > gB ? 'быстрее' : 'медленнее'} выручки (темп роста ${F1(gC)}% против ${F1(gB)}%), поэтому ее доля в выручке ${s1 > s0 ? 'выросла' : 'снизилась'} с ${P2(s0)}% до ${P2(s1)}%.`);
  }
  if (a.B && b.B){
    const pp0 = shareOf(a, a.R['2200']), pp1 = shareOf(b, b.R['2200']), np0 = shareOf(a, a.R['2400']), np1 = shareOf(b, b.R['2400']);
    const rub = x => fmtNum(Math.abs(x) / 100, 2) + ' руб.';
    const tail = np0 >= 0 && np1 >= 0
      ? `каждый рубль выручки в ${y1} г. приносит ${np1 >= np0 ? 'больше' : 'меньше'} чистой прибыли, чем в ${y0} г. (${rub(np1)} против ${rub(np0)})`
      : np0 < 0 && np1 < 0
        ? `убыток на каждый рубль выручки ${np1 >= np0 ? 'сократился' : 'вырос'} с ${rub(np0)} в ${y0} г. до ${rub(np1)} в ${y1} г.`
        : np1 >= 0 ? `в ${y1} г. каждый рубль выручки приносит ${rub(np1)} чистой прибыли вместо ${rub(np0)} убытка в ${y0} г.`
                   : `в ${y1} г. на каждый рубль выручки приходится ${rub(np1)} убытка вместо ${rub(np0)} прибыли в ${y0} г.`;
    ps.push(`Доля прибыли от продаж в выручке изменилась с ${P2(pp0)}% до ${P2(pp1)}%, чистой прибыли — с ${P2(np0)}% до ${P2(np1)}%: ${tail}${tail.endsWith('.') ? '' : '.'}`);
  }
  // какая из прочих статей расходов изменилась сильнее всего
  const items = [['коммерческие расходы', 'KR'], ['управленческие расходы', 'UR'], ['проценты к уплате', 'o2330'], ['прочие расходы', 'o2350']]
    .map(([n, k]) => [n, a[k], b[k], b[k] - a[k]]).sort((p, q) => Math.abs(q[3]) - Math.abs(p[3]));
  const [n, x0, x1, dx] = items[0];
  if (Math.abs(dx) >= 1) ps.push(`Из остальных статей расходов сильнее всего изменились ${n}: ${dx > 0 ? 'рост' : 'снижение'} с ${F0(x0)} до ${F0(x1)} тыс. руб. (${sgn(dx)} тыс. руб.).`);
  return ps.join(' ');
}

// Таблица 2 — проверка цепочки
function tcFr2(M){
  const D = M.D, I = yearsWith(D); if (!I.length) return 'Нет данных отчета о финансовых результатах.';
  const bad = [];
  for (const i of I) for (const [n, k, c] of [['валовой прибыли', 'VP', '2100'], ['прибыли от продаж', 'PP', '2200'], ['прибыли до налогообложения', 'PN', '2300']]){
    const x = D[i][k] - D[i].R[c]; if (Math.abs(x) >= 1) bad.push(`${n} за ${Y(i)} г. (${sgn(x)} тыс. руб.)`);
  }
  const ps = [bad.length
    ? `Пересчет выявил расхождения с отчетными строками: ${bad.join('; ')} — следует проверить введенные строки отчета.`
    : `Пересчет валовой прибыли, прибыли от продаж и прибыли до налогообложения по формулам полностью совпал со строками 2100, 2200 и 2300 отчета за ${I.map(Y).join(', ')} гг. — цепочка формирования прибыли отражена в отчетности корректно.`];
  const npDiff = I.filter(i => Math.abs(D[i].PCH - D[i].R['2400']) >= 1);
  const by2460 = npDiff.filter(i => Math.abs(D[i].PCH - D[i].R['2400'] + D[i].P2460) < 1);
  if (!npDiff.length) ps.push('Чистая прибыль, рассчитанная как Пнобл − 2410, совпадает с отчетной строкой 2400.');
  else if (by2460.length === npDiff.length) ps.push(`Расхождение по чистой прибыли (${npDiff.map(i => `${Y(i)} г.: ${sgn(D[i].PCH - D[i].R['2400'])}`).join('; ')} тыс. руб.) полностью объясняется строкой 2460 «Прочее», которую формула ПЧ = Пнобл − 2410 не учитывает; с учетом этой строки расчет совпадает с отчетом.`);
  else ps.push(`По чистой прибыли есть расхождение, не объясняемое строкой 2460 (${npDiff.filter(i => !by2460.includes(i)).map(Y).join(', ')} г.) — проверьте строки налога 2411/2412.`);
  const L = I.at(-1), d = D[L];
  if (d.PN > 0) ps.push(`Налог на прибыль в ${Y(L)} г. составил ${F0(d.TAX)} тыс. руб., или ${F1(d.TAX / d.PN * 100)}% прибыли до налогообложения.`);
  return ps.join(' ');
}

// Тексты по факторам (используются и под таблицами, и в общем блоке выводов)
function ppFactorText(i, F){
  const items = ['fB', 'fC', 'fKR', 'fUR'].map(key => [key, F[key]]);
  const pos = items.filter(x => x[1] >= 1).sort((a, b) => b[1] - a[1])[0];
  const neg = items.filter(x => x[1] <= -1).sort((a, b) => a[1] - b[1])[0];
  let s = `Прибыль от продаж в ${Y(i)} г. изменилась на ${sgn(F.dPP)} тыс. руб. `;
  s += pos ? `Наибольшее положительное влияние оказало ${FACTOR_NAMES[pos[0]]} (${sgn(pos[1])} тыс. руб.)` : 'Факторов с положительным влиянием не было';
  s += neg ? `, наибольшее отрицательное — ${FACTOR_NAMES[neg[0]]} (${sgn(neg[1])} тыс. руб.).` : ', факторов с отрицательным влиянием не было.';
  return s + ` Сумма влияний (${sgn(F.sum)} тыс. руб.) совпадает с изменением прибыли от продаж.`;
}
function npFactorText(i, F){
  const items = Object.keys(NP_FACTOR_NAMES).map(key => [key, F[key]]).sort((a, b) => Math.abs(b[1]) - Math.abs(a[1]));
  const top = items[0];
  const plus = items.filter(x => x[1] >= 1).map(([k, x]) => `${NP_FACTOR_NAMES[k]} (${sgn(x)})`);
  const minus = items.filter(x => x[1] <= -1).map(([k, x]) => `${NP_FACTOR_NAMES[k]} (${sgn(x)})`);
  return `Чистая прибыль в ${Y(i)} г. изменилась на ${sgn(F.dNP)} тыс. руб. Сильнее всего на нее повлияло ${NP_FACTOR_NAMES[top[0]]} (${sgn(top[1])} тыс. руб.). ` +
    (plus.length ? `Увеличили прибыль (тыс. руб.): ${plus.join(', ')}. ` : 'Факторов, увеличивших прибыль, не было. ') +
    (minus.length ? `Уменьшили прибыль (тыс. руб.): ${minus.join(', ')}.` : 'Факторов, уменьшивших прибыль, не было.') + ' Сумма влияний совпадает с изменением чистой прибыли.';
}
function dupontText(i, F){
  if (!F) return `Разложение ROE за ${Y(i)} г. по модели Дюпона не выполнено: собственный капитал отрицательный или нет данных баланса.`;
  const items = Object.keys(DUPONT_NAMES).map(key => [key, F[key]]).sort((a, b) => Math.abs(b[1]) - Math.abs(a[1]));
  const pp = x => `${x >= 0 ? '+' : '−'}${fmtNum(Math.abs(x), 2)} п. п.`;
  return `ROE в ${Y(i)} г. ${F.dROE >= 0 ? 'выросла' : 'снизилась'} на ${fmtNum(Math.abs(F.dROE), 2)} п. п. Сильнее всего ROE изменил фактор «${DUPONT_NAMES[items[0][0]]}» (${pp(items[0][1])}); влияние остальных: ` +
    items.slice(1).map(([key, x]) => `${DUPONT_NAMES[key]} ${pp(x)}`).join(', ');
}

// Таблица 3 — факторы прибыли от продаж
function tcFr3(M){
  const t = M.tables[2], ps = [];
  pairs().forEach((i, k) => { if (t.F[k]) ps.push(ppFactorText(i, t.F[k])); });
  if (!ps.length) return 'Факторный анализ не выполнен: нужны данные минимум за два соседних года с ненулевой выручкой.';
  const last = t.F.at(-1);
  if (last && last.fC <= -1) ps.push('Рост доли себестоимости в выручке — ключевой резерв: возврат ее уровня к прошлому году увеличил бы прибыль от продаж на ' + F0(-last.fC) + ' тыс. руб.');
  return ps.join(' ');
}

// Таблица 4 — рентабельность
function tcFr4(M){
  const D = M.D, R = M.tables[3].R, I = yearsWith(D); if (I.length < 2) return FR_NEED2;
  const L = I.at(-1), P = I.at(-2), r0 = R[P], r1 = R[L], up = [], down = [];
  for (const [k, name] of RENT_DEFS){
    const x0 = r0[k], x1 = r1[k]; if (x0 == null || x1 == null) continue;
    (x1 >= x0 ? up : down).push(`${lcFirst(name)} (${P2(x0)}% → ${P2(x1)}%)`);
  }
  const n = up.length + down.length, ps = [];
  if (!down.length) ps.push(`В ${Y(L)} г. по сравнению с ${Y(P)} г. выросли все рассчитанные показатели рентабельности (${n}): ${up.join('; ')}.`);
  else if (!up.length) ps.push(`В ${Y(L)} г. по сравнению с ${Y(P)} г. снизились все рассчитанные показатели рентабельности (${n}): ${down.join('; ')}.`);
  else ps.push(`В ${Y(L)} г. по сравнению с ${Y(P)} г. снизились ${down.length} из ${n} показателей рентабельности: ${down.join('; ')}. Выросли: ${up.join('; ')}.`);
  const vals = RENT_DEFS.map(([k, name]) => [name, r1[k]]).filter(x => x[1] != null).sort((a, b) => b[1] - a[1]);
  if (vals.length > 1) ps.push(`В ${Y(L)} г. наибольшее значение у показателя «${lcFirst(vals[0][0])}» (${P2(vals[0][1])}%), наименьшее — «${lcFirst(vals.at(-1)[0])}» (${P2(vals.at(-1)[1])}%).`);
  if (vals.some(x => x[1] < 0)) ps.push('Отрицательные значения означают, что компания работает с убытком: на каждый рубль капитала, активов или выручки приходится убыток.');
  if (r1.roe == null && r1.K && r1.K.x != null && r1.K.x <= 0) ps.push('ROE не рассчитывается: собственный капитал отрицательный.');
  if (I.length > 2 && R[I[0]].roa != null && r1.roa != null) ps.push(`За весь период ${Y(I[0])}–${Y(L)} гг. рентабельность активов изменилась с ${P2(R[I[0]].roa)}% до ${P2(r1.roa)}%${R[I[0]].roe != null && r1.roe != null ? `, рентабельность собственного капитала — с ${P2(R[I[0]].roe)}% до ${P2(r1.roe)}%` : ''}.`);
  return ps.join(' ');
}

// Таблица 5 — факторы чистой прибыли
function tcFr5(M){
  const t = M.tables[4], ps = [];
  pairs().forEach((i, k) => { if (t.F[k]) ps.push(npFactorText(i, t.F[k])); });
  return ps.length ? ps.join(' ') : FR_NEED2;
}

// Таблица 6 — модель Дюпона
function tcFr6(M){
  const Mm = M.tables[5].M, I = state.years.map((_, i) => i).filter(i => Mm[i] && Mm[i].roe != null);
  if (I.length < 1){
    return Mm.some(m => m && m.negK) ? 'Модель Дюпона не рассчитывается: собственный капитал отрицательный, поэтому ROE и мультипликатор капитала не имеют экономического смысла.' : 'Модель Дюпона не рассчитана: нет данных баланса или выручки.';
  }
  const a = Mm[I[0]], b = Mm[I.at(-1)], y0 = Y(I[0]), y1 = Y(I.at(-1));
  if (I.length === 1) return `В ${y1} г. ROE = ${P2(b.a)}% × ${fmtNum(b.b, 3)} × ${fmtNum(b.c, 3)} = ${P2(b.roe)}%.`;
  const ps = [`За ${y0}–${y1} гг. рентабельность продаж по чистой прибыли ${b.a >= a.a ? 'выросла' : 'снизилась'} с ${P2(a.a)}% до ${P2(b.a)}%.`,
    `Оборачиваемость активов ${b.b >= a.b ? 'выросла' : 'снизилась'} с ${fmtNum(a.b, 3)} до ${fmtNum(b.b, 3)} раза: каждый рубль активов в ${y1} г. приносит ${fmtNum(b.b, 2)} руб. выручки${b.b < a.b ? ' — имущество используется менее эффективно (активы растут быстрее выручки)' : ''}.`,
    `Мультипликатор капитала ${b.c >= a.c ? 'вырос' : 'снизился'} с ${fmtNum(a.c, 3)} до ${fmtNum(b.c, 3)}: ${b.c >= a.c ? 'доля заемных средств в финансировании активов растет — это повышает ROE, но увеличивает финансовый риск' : 'компания больше опирается на собственный капитал — финансовый риск снижается'}.`,
    `В результате ROE ${b.roe >= a.roe ? 'выросла' : 'снизилась'} с ${P2(a.roe)}% до ${P2(b.roe)}%.`];
  return ps.join(' ');
}

// Таблица 7 — разложение ROE
function tcFr7(M){
  const t = M.tables[6], D = M.D, ps = [];
  pairs().forEach((i, k) => { if (D[i] && D[i - 1]) ps.push(dupontText(i, t.F[k])); });
  if (!ps.length) return FR_NEED2;
  if (t.F.some(Boolean)) ps.push('Сумма влияний трех факторов совпадает с изменением ROE (строка проверки равна нулю).');
  return ps.join(' ');
}
const FR_TC = {fr1:tcFr1, fr2:tcFr2, fr3:tcFr3, fr4:tcFr4, fr5:tcFr5, fr6:tcFr6, fr7:tcFr7};
// текст вывода: свой (если правили вручную) или автоматический
const frTcText = (id, M) => (state.conclusions[id] || '').trim() || FR_TC[id](M);

/* ===================== Рендер вкладки ===================== */
function frBuild(){
  const D = frAll();
  const t3 = frT3(D), t4 = frT4(D), t5 = frT5(D), t6 = frT6(D), t7 = frT7(D, t6);
  const M = {D, tables:[frT1(D), frT2(D), t3, t4, t5, t6, t7], concl:frConclusions(D, t3, t4, t5, t7), dataNotes:frDataNotes(D)};
  M.tc = Object.fromEntries(M.tables.map(t => [t.id, frTcText(t.id, M)]));
  return M;
}
// Рисунки стоят после таблиц: 1 — после таблицы 1, 2 — после таблицы 4
const FR_FIGS = [
  {id:'frFigBars', after:'fr1', num:1, title:'Динамика выручки, прибыли от продаж и чистой прибыли, млн руб.'},
  {id:'frFigLines', after:'fr4', num:2, title:'Динамика основных показателей рентабельности, %'},
];
function renderFR(){
  const box = document.getElementById('frOut'); if (!box) return;
  const M = frBuild();
  const notesEl = document.getElementById('frNotes');
  notesEl.innerHTML = M.dataNotes.length ? '<b>Пометки к данным</b><ul>' + M.dataNotes.map(s => `<li>${esc(s)}</li>`).join('') + '</ul>' : '';
  document.getElementById('frToc').innerHTML = M.tables.map(t => `<a href="#${t.id}">Таблица ${t.num}</a>`).join('') +
    FR_FIGS.map(f => `<a href="#${f.id}">Рисунок ${f.num}</a>`).join('') + '<a href="#frConcl">Выводы и резервы</a>';
  const fig = id => FR_FIGS.filter(f => f.after === id).map(f => `
    <figure class="chart" id="${f.id}">
      <div class="chartbox"><canvas role="img" aria-label="Рисунок ${f.num} – ${esc(f.title)}"></canvas></div>
      <figcaption>Рисунок ${f.num} – ${esc(f.title)}</figcaption>
    </figure>`).join('');
  box.innerHTML = M.tables.map(t => `
    <figure class="tbl" id="${t.id}">
      <figcaption>Таблица ${t.num} – ${esc(t.title)}</figcaption>
      <div class="scroll">${tableHTML(t)}</div>
      <p class="formula"><b>Формулы:</b> ${esc(t.formula)}</p>
      ${t.notes.length ? `<ul class="tnotes">${t.notes.map(s => `<li>${esc(s)}</li>`).join('')}</ul>` : ''}
      <div class="concl">
        <label class="concl-h" for="tc_${t.id}">Вывод по таблице ${t.num}</label>
        <textarea id="tc_${t.id}" data-frid="${t.id}">${esc(M.tc[t.id])}</textarea>
        <div class="row">
          <button class="btn ghost small" data-frauto="${t.id}">Пересобрать вывод по цифрам</button>
          <small data-frst="${t.id}">${(state.conclusions[t.id] || '').trim() ? 'Вывод изменен вручную.' : 'Собран автоматически по цифрам таблицы — можно дополнить своими словами.'}</small>
        </div>
      </div>
    </figure>${fig(t.id)}`).join('') +
    `<figure class="tbl" id="frConcl"><figcaption>Выводы и резервы роста прибыли</figcaption>
      <div class="frconcl">${M.concl.map(b => `<h3>${esc(b.h)}</h3>${b.ps.map(p => `<p>${esc(p)}</p>`).join('')}`).join('')}</div>
      <p class="formula">Текст собирается автоматически по правилам из цифр таблиц 1–7 (без нейросетей) и обновляется при изменении данных.</p>
    </figure>`;
  box.querySelectorAll('textarea[data-frid]').forEach(autosize);
  drawFrCharts();
}
// Правка вывода вручную — сохраняем; кнопка возвращает автоматический текст
document.getElementById('frOut').addEventListener('input', e => {
  const ta = e.target.closest('textarea[data-frid]'); if (!ta) return;
  state.conclusions[ta.dataset.frid] = ta.value; autosize(ta); save();
  const st = document.querySelector(`[data-frst="${ta.dataset.frid}"]`); if (st) st.textContent = 'Вывод изменен вручную.';
});
document.getElementById('frOut').addEventListener('click', e => {
  const b = e.target.closest('button[data-frauto]'); if (!b) return;
  const id = b.dataset.frauto, ta = document.querySelector(`textarea[data-frid="${id}"]`);
  if ((state.conclusions[id] || '').trim() && !confirm('Заменить ваш текст автоматическим выводом?')) return;
  delete state.conclusions[id]; save();
  ta.value = FR_TC[id](frBuild()); autosize(ta);
  document.querySelector(`[data-frst="${id}"]`).textContent = 'Собран автоматически по цифрам таблицы — можно дополнить своими словами.';
});

/* ===================== Выгрузка вкладки в Excel ===================== */
function frSheet(t, tc){
  const XL = window.XLSX, rows = allRows(t), grid = layout(rows, t.ncols);
  const aoa = [[`Таблица ${t.num} - ${t.title}`]], merges = [], off = 1;
  rows.forEach((row, r) => {
    const line = [];
    for (let c = 0; c < t.ncols; c++){
      const g = grid[r][c];
      if (!g || g.r0 !== r || g.c0 !== c){ line.push(null); continue; }
      const cell = g.cell;
      if ((cell.cs || 1) > 1 || (cell.rs || 1) > 1) merges.push({s:{r:r + off, c}, e:{r:r + off + (cell.rs || 1) - 1, c:c + (cell.cs || 1) - 1}});
      if (cell.v !== undefined){
        if (cell.v == null || !Number.isFinite(cell.v)) line.push('-');
        else line.push({t:'n', v:Math.round(cell.v * 10 ** cell.d) / 10 ** cell.d, z:cell.d ? '#,##0.' + '0'.repeat(cell.d) : '#,##0'});
      } else line.push(cell.t);
    }
    aoa.push(line);
  });
  aoa.push([], ['Формулы: ' + t.formula]);
  for (const s of t.notes) aoa.push([s]);
  if (tc){ aoa.push([], ['Вывод']); for (const p of tc.split(/\n+/).filter(Boolean)) aoa.push([p]); }
  const ws = XL.utils.aoa_to_sheet(aoa); ws['!merges'] = merges;
  const pairOnly = ['fr3', 'fr5', 'fr7'].includes(t.id);   // у таблиц «по парам лет» нет колонки формулы/кода
  ws['!cols'] = [{wch:pairOnly ? 60 : 46}, ...Array(t.ncols - 1).fill(0).map((_, j) => ({wch:j === 0 && !pairOnly ? 24 : 14}))];
  return ws;
}
// Лист «Графики»: данные для обоих рисунков; сами рисунки вставляются ниже картинками (см. addImagesToXlsx)
function frChartSheet(M){
  const XL = window.XLSX, C = frChartData(M.D, M.tables[3].R);
  const num = (x, d) => x == null || !Number.isFinite(x) ? '-' : {t:'n', v:Math.round(x * 10 ** d) / 10 ** d, z:d ? '#,##0.' + '0'.repeat(d) : '#,##0'};
  const aoa = [['Данные для рисунка 1 — ' + FR_FIGS[0].title.replace(', млн руб.', '') + ', тыс. руб.'],
    ['Год', ...C.bars.map(s => s[0])], ...C.idx.map(i => [Y(i), ...C.bars.map(s => num(s[1](i), 0))]), [],
    ['Данные для рисунка 2 — ' + FR_FIGS[1].title],
    ['Год', ...C.lines.map(s => s[0])], ...C.idx.map(i => [Y(i), ...C.lines.map(s => num(s[1](i), 2))]), []];
  aoa.push(['Рисунки 1 и 2 вставлены ниже картинками (если их не видно — откройте файл в Excel; данные выше позволяют построить графики заново: Вставка → Диаграмма).']);
  const ws = XL.utils.aoa_to_sheet(aoa); ws['!cols'] = [{wch:12}, ...Array(6).fill({wch:26})];
  return ws;
}
// PNG рисунков в светлой теме (как для Word), 800×400 ×2
async function frChartPngs(){
  await loadScript(CHART_URL);
  const D = frAll(), C = frChartData(D, frRent(D)), th = chartTheme(true), out = [];
  if (!C.idx.length) return out;
  const box = document.createElement('div'); box.style.cssText = 'position:fixed;left:-10000px;top:0;width:800px;height:400px;';
  document.body.appendChild(box);
  try{
    const cfgs = [chartConfig({kind:'bars', unit:'млн руб.', idx:C.idx, series:C.bars}, th, true), frLineConfig(C.idx, C.lines, th)];
    const lc = cfgs[1].options; lc.responsive = false; lc.devicePixelRatio = 2;
    lc.scales.x.ticks.font = lc.scales.y.ticks.font = lc.plugins.legend.labels.font = {family:th.font, size:14};
    for (const cfg of cfgs){
      const cv = document.createElement('canvas'); cv.width = 800; cv.height = 400; box.appendChild(cv);
      const ch = new Chart(cv, cfg);
      const blob = await new Promise(r => cv.toBlob(r, 'image/png'));
      out.push(new Uint8Array(await blob.arrayBuffer()));
      ch.destroy(); cv.remove();
    }
  } finally { box.remove(); }
  return out;
}
// Дописывает картинки в готовый .xlsx (SheetJS бесплатной версии картинки не пишет): media + drawing + связи
async function addImagesToXlsx(buf, sheetNo, pngs, row0){
  await loadScript(JSZIP_URL);
  const zip = await JSZip.loadAsync(buf), EMU = 9525, W = 760 * EMU, H = 380 * EMU;
  pngs.forEach((png, k) => zip.file(`xl/media/frchart${k + 1}.png`, png));
  const anchors = pngs.map((_, k) => `<xdr:oneCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${row0 + k * 21}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="${W}" cy="${H}"/>` +
    `<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${k + 2}" name="Рисунок ${k + 1}"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>` +
    `<xdr:blipFill><a:blip r:embed="rIdImg${k + 1}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>` +
    `<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${W}" cy="${H}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>`).join('');
  zip.file('xl/drawings/drawing1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${anchors}</xdr:wsDr>`);
  zip.file('xl/drawings/_rels/drawing1.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    pngs.map((_, k) => `<Relationship Id="rIdImg${k + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/frchart${k + 1}.png"/>`).join('') + '</Relationships>');
  // связь листа с рисунком
  const relPath = `xl/worksheets/_rels/sheet${sheetNo}.xml.rels`, rel = '<Relationship Id="rIdDrw1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>';
  const oldRels = zip.file(relPath) ? await zip.file(relPath).async('string') : null;
  zip.file(relPath, oldRels ? oldRels.replace('</Relationships>', rel + '</Relationships>')
    : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rel}</Relationships>`);
  const shPath = `xl/worksheets/sheet${sheetNo}.xml`;
  let sh = await zip.file(shPath).async('string');
  if (!/xmlns:r=/.test(sh)) sh = sh.replace('<worksheet ', '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ');
  zip.file(shPath, sh.replace('</worksheet>', '<drawing r:id="rIdDrw1"/></worksheet>'));
  let ct = await zip.file('[Content_Types].xml').async('string');
  if (!/Extension="png"/.test(ct)) ct = ct.replace('<Default ', '<Default Extension="png" ContentType="image/png"/><Default ');
  ct = ct.replace('</Types>', '<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>');
  zip.file('[Content_Types].xml', ct);
  return zip.generateAsync({type:'blob', mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
}
async function buildFrXlsx(){
  await loadScript(XLSX_URL); const XL = window.XLSX, wb = XL.utils.book_new(), M = frBuild();
  const names = ['1 Динамика ФР', '2 Цепочка прибыли', '3 Факторы Пп', '4 Рентабельность', '5 Факторы ЧП', '6 Модель Дюпона', '7 Разложение ROE'];
  M.tables.forEach((t, k) => XL.utils.book_append_sheet(wb, frSheet(t, M.tc[t.id]), names[k]));
  const chartsWs = frChartSheet(M);
  XL.utils.book_append_sheet(wb, chartsWs, 'Графики');
  const chartSheetNo = wb.SheetNames.length;   // SheetJS нумерует листы по порядку: sheet1.xml, sheet2.xml, …
  const aoa = [[`Выводы и резервы роста прибыли — ${state.company || 'организация'}, ${state.years[0]}–${state.years.at(-1)} гг.`]];
  for (const b of M.concl){ aoa.push([], [b.h]); for (const p of b.ps) aoa.push([p]); }
  if (M.dataNotes.length){ aoa.push([], ['Пометки к данным']); for (const s of M.dataNotes) aoa.push([s]); }
  const ws = XL.utils.aoa_to_sheet(aoa); ws['!cols'] = [{wch:140}];
  XL.utils.book_append_sheet(wb, ws, 'Выводы');
  const buf = XL.write(wb, {bookType:'xlsx', type:'array'});
  // картинки графиков: если Chart.js или JSZip не загрузились — файл без картинок, но с данными
  try{
    const pngs = await frChartPngs();
    const row0 = XL.utils.decode_range(chartsWs['!ref']).e.r + 2;
    if (pngs.length) return await addImagesToXlsx(buf, chartSheetNo, pngs, row0);
  }catch(e){ status('Графики в Excel не добавлены: ' + e.message); }
  return new Blob([buf]);
}
document.getElementById('dlFr').onclick = async () => {
  try{ status('Собираю Excel…'); await offer(`Финрезультаты_${fileBase()}.xlsx`, await buildFrXlsx()); }
  catch(e){ msg('Ошибка при сборке Excel: ' + e.message, true); status('Ошибка при сборке Excel'); }
};
