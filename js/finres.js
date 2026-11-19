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

/* ===================== Таблица 4. Рентабельность ===================== */
function frRent(D){
  return D.map((d, i) => {
    if (!d) return null;
    const A = avgBal('1600', i), K = avgBal('1300', i);
    const costs = d.C + d.KR + d.UR;
    return {
      gross: d.B ? d.R['2100'] / d.B * 100 : null,
      sales: d.B ? d.R['2200'] / d.B * 100 : null,
      net:   d.B ? d.R['2400'] / d.B * 100 : null,
      cost:  costs ? d.R['2200'] / costs * 100 : null,
      roa:   A.x ? d.R['2400'] / A.x * 100 : null,
      roe:   K.x > 0 ? d.R['2400'] / K.x * 100 : null,   // при отрицательном капитале ROE не имеет смысла
      A, K,
    };
  });
}
const RENT_DEFS = [
  ['gross', 'Валовая рентабельность', 'ВП / В'],
  ['sales', 'Рентабельность продаж', 'Пп / В'],
  ['net',   'Рентабельность по чистой прибыли', 'ПЧ / В'],
  ['cost',  'Рентабельность затрат', 'Пп / (С + КР + УР)'],
  ['roa',   'Рентабельность активов (ROA)', 'ПЧ / ср. активы (1600)'],
  ['roe',   'Рентабельность собственного капитала (ROE)', 'ПЧ / ср. СК (1300)'],
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
  if (endOnly.length) notes.push(`За ${endOnly.map(Y).join(', ')} г. нет баланса на начало года (31.12 прошлого года) — для ROA и ROE взята величина активов и капитала на конец года.`);
  const noBal = R.map((r, i) => r && r.A.how === 'none' ? i : -1).filter(i => i >= 0);
  if (noBal.length) notes.push(`За ${noBal.map(Y).join(', ')} г. нет баланса — ROA и ROE не рассчитаны.`);
  const negK = R.map((r, i) => r && r.K.x != null && r.K.x <= 0 ? i : -1).filter(i => i >= 0);
  if (negK.length) notes.push(`За ${negK.map(Y).join(', ')} г. собственный капитал отрицательный — ROE не рассчитывается.`);
  return {id:'fr4', num:'4', title:'Показатели рентабельности, %', head, rows, ncols:2 + nYears() + P.length, R,
    formula:'Rвал = ВП / В; Rпрод = Пп / В; Rчп = ПЧ / В; Rзатр = Пп / (С + КР + УР); ROA = ПЧ / ((стр. 1600 на начало + на конец года) / 2); ROE = ПЧ / ((стр. 1300 на начало + на конец года) / 2); всё × 100%, отклонение — в процентных пунктах.',
    notes};
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
function frConclusions(D, t3, t4){
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

  // 2. Факторы
  const fac = [];
  pairs().forEach((i, k) => {
    const F = t3.F[k]; if (!F) return;
    const items = ['fB', 'fC', 'fKR', 'fUR'].map(key => [key, F[key]]);
    const pos = items.filter(x => x[1] >= 1).sort((a, b) => b[1] - a[1])[0];
    const neg = items.filter(x => x[1] <= -1).sort((a, b) => a[1] - b[1])[0];
    let s = `Прибыль от продаж в ${Y(i)} г. изменилась на ${sgn(F.dPP)} тыс. руб. `;
    s += pos ? `Наибольшее положительное влияние оказало ${FACTOR_NAMES[pos[0]]} (${sgn(pos[1])} тыс. руб.)` : 'Факторов с положительным влиянием не было';
    s += neg ? `, наибольшее отрицательное — ${FACTOR_NAMES[neg[0]]} (${sgn(neg[1])} тыс. руб.).` : ', факторов с отрицательным влиянием не было.';
    s += ` Сумма влияний (${sgn(F.sum)} тыс. руб.) совпадает с изменением прибыли от продаж.`;
    fac.push(s);
  });
  if (fac.length) out.push({h:'Влияние факторов на прибыль от продаж', ps:fac});

  // 3. Рентабельность — за последний период
  const L = P.at(-1), r0 = t4.R[L - 1], r1 = t4.R[L];
  const rs = [rentText('рентабельность продаж', r0.sales, r1.sales), rentText('рентабельность активов (ROA)', r0.roa, r1.roa),
    rentText('рентабельность собственного капитала (ROE)', r0.roe, r1.roe)].filter(Boolean);
  if (rs.length) out.push({h:'Рентабельность', ps:[`В ${Y(L)} г. по сравнению с ${Y(L - 1)} г. ` + rs.join('; ') + '.']});

  // 4. Резервы роста прибыли — за последний период
  const a = D[L - 1], b = D[L], res = []; let total = 0;
  if (a.B && b.B){
    for (const [key, name, gen] of [['C', 'себестоимости', 'себестоимости'], ['KR', 'коммерческих расходов', 'коммерческих расходов'], ['UR', 'управленческих расходов', 'управленческих расходов']]){
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

/* ===================== Рендер вкладки ===================== */
function frBuild(){
  const D = frAll();
  const t3 = frT3(D), t4 = frT4(D);
  return {D, tables:[frT1(D), frT2(D), t3, t4], concl:frConclusions(D, t3, t4), dataNotes:frDataNotes(D)};
}
function renderFR(){
  const box = document.getElementById('frOut'); if (!box) return;
  const M = frBuild();
  const notesEl = document.getElementById('frNotes');
  notesEl.innerHTML = M.dataNotes.length ? '<b>Пометки к данным</b><ul>' + M.dataNotes.map(s => `<li>${esc(s)}</li>`).join('') + '</ul>' : '';
  document.getElementById('frToc').innerHTML = M.tables.map(t => `<a href="#${t.id}">Таблица ${t.num}</a>`).join('') + '<a href="#frConcl">Выводы и резервы</a>';
  box.innerHTML = M.tables.map(t => `
    <figure class="tbl" id="${t.id}">
      <figcaption>Таблица ${t.num} – ${esc(t.title)}</figcaption>
      <div class="scroll">${tableHTML(t)}</div>
      <p class="formula"><b>Формулы:</b> ${esc(t.formula)}</p>
      ${t.notes.length ? `<ul class="tnotes">${t.notes.map(s => `<li>${esc(s)}</li>`).join('')}</ul>` : ''}
    </figure>`).join('') +
    `<figure class="tbl" id="frConcl"><figcaption>Выводы и резервы роста прибыли</figcaption>
      <div class="frconcl">${M.concl.map(b => `<h3>${esc(b.h)}</h3>${b.ps.map(p => `<p>${esc(p)}</p>`).join('')}`).join('')}</div>
      <p class="formula">Текст собирается автоматически по правилам из цифр таблиц 1–4 (без нейросетей) и обновляется при изменении данных.</p>
    </figure>`;
}

/* ===================== Выгрузка вкладки в Excel ===================== */
function frSheet(t){
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
  const ws = XL.utils.aoa_to_sheet(aoa); ws['!merges'] = merges;
  ws['!cols'] = [{wch:46}, ...Array(t.ncols - 1).fill(0).map((_, j) => ({wch:j === 0 && t.id !== 'fr3' ? 22 : 14}))];
  return ws;
}
document.getElementById('dlFr').onclick = async () => {
  try{
    await loadScript(XLSX_URL); const XL = window.XLSX, wb = XL.utils.book_new(), M = frBuild();
    const names = ['1 Динамика ФР', '2 Цепочка прибыли', '3 Факторный анализ', '4 Рентабельность'];
    M.tables.forEach((t, k) => XL.utils.book_append_sheet(wb, frSheet(t), names[k]));
    const aoa = [[`Выводы и резервы роста прибыли — ${state.company || 'организация'}, ${state.years[0]}–${state.years.at(-1)} гг.`]];
    for (const b of M.concl){ aoa.push([], [b.h]); for (const p of b.ps) aoa.push([p]); }
    if (M.dataNotes.length){ aoa.push([], ['Пометки к данным']); for (const s of M.dataNotes) aoa.push([s]); }
    const ws = XL.utils.aoa_to_sheet(aoa); ws['!cols'] = [{wch:140}];
    XL.utils.book_append_sheet(wb, ws, 'Выводы');
    await offer(`Финрезультаты_${fileBase()}.xlsx`, new Blob([XL.write(wb, {bookType:'xlsx', type:'array'})]));
  }catch(e){ msg('Ошибка при сборке Excel: ' + e.message, true); status('Ошибка при сборке Excel'); }
};
