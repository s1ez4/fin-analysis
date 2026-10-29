/* Экспорт: Excel (SheetJS) и Word (.docx вручную через JSZip) */
/* ===================== Загрузка библиотек ===================== */
const libs={};
function loadScript(src){ return libs[src] ||= new Promise((res,rej)=>{ const s=document.createElement('script'); s.src=src; s.onload=res; s.onerror=()=>{delete libs[src]; rej(new Error('Не удалось загрузить '+src));}; document.head.appendChild(s); }); }
const XLSX_URL='https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
const JSZIP_URL='https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js';
const PDF_URL='https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js';
const PDFW_URL='https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

/* ===================== Скачивание ===================== */
async function offer(filename, data){
  if (!downloadsNs){
    if (!window.claude){ // открыто как обычный файл в браузере
      const a=document.createElement('a'); a.href=URL.createObjectURL(data instanceof Blob?data:new Blob([data])); a.download=filename;
      document.body.appendChild(a); a.click(); a.remove(); setTimeout(()=>URL.revokeObjectURL(a.href),5000); status('Файл сохранен: '+filename); return;
    }
    msg('Скачивание недоступно в этом окне.', true); return;
  }
  try { await downloadsNs.save({filename, data}); status('Файл сохранен: '+filename); }
  catch(e){ if(e&&e.code==='declined') status('Сохранение отменено'); else msg('Не получилось сохранить файл: '+(e&&e.message||e), true); }
}
const fileBase = ()=> (state.company||'компания').replace(/[«»"\\/:*?<>|]/g,'').trim().replace(/\s+/g,'_').slice(0,60)||'компания';

function sourceAoa(withValues=true){
  const aoa=[['Код','Показатель',...state.years]];
  const push=(title,lines)=>{ aoa.push([title]); for(const l of lines){ if(l.h){aoa.push(['',l.h]);continue;}
    aoa.push([l.c,l.n,...state.years.map((_,i)=>{ if(!withValues) return null; const x=v(l.c,i); return x==null?null:x; })]); } };
  push('БУХГАЛТЕРСКИЙ БАЛАНС — АКТИВ (тыс. руб.)',LINES.asset);
  push('БУХГАЛТЕРСКИЙ БАЛАНС — ПАССИВ (тыс. руб.)',LINES.liab);
  push('ОТЧЕТ О ФИНАНСОВЫХ РЕЗУЛЬТАТАХ (тыс. руб., расходы положительными числами)',LINES.pl);
  return aoa;
}
// Лист Excel с таблицей в том же виде, что в Word (шапки и подписи как в образце); числа — числами
function tableSheet(s){
  const rows=[...s.head,...s.body], grid=layout(rows,s.ncols), XL=window.XLSX;
  const aoa=[[`Таблица ${s.num} - ${s.title}`]], merges=[], off=1;
  rows.forEach((row,r)=>{
    const line=[];
    for(let c=0;c<s.ncols;c++){
      const g=grid[r][c];
      if(!g||g.r0!==r||g.c0!==c){ line.push(null); continue; }
      const cell=g.cell;
      if((cell.cs||1)>1||(cell.rs||1)>1) merges.push({s:{r:r+off,c},e:{r:r+off+(cell.rs||1)-1,c:c+(cell.cs||1)-1}});
      if(cell.v!==undefined){
        if(cell.v==null||!Number.isFinite(cell.v)) line.push('-');
        else line.push({t:'n', v:Math.round(cell.v*10**cell.d)/10**cell.d, z:cell.plus?'+General;-General;0':'General'});
      } else line.push(cell.t);
    }
    aoa.push(line);
  });
  const concl=state.conclusions[s.id]; if(concl){ aoa.push([]); for(const p of concl.split(/\n+/).filter(Boolean)) aoa.push([p]); }
  const ws=XL.utils.aoa_to_sheet(aoa); ws['!merges']=merges;
  ws['!cols']=s.grid.map(w=>({wch:Math.max(8,Math.round(w/100))}));   // ширины из образца (twips → символы)
  return ws;
}
document.getElementById('dlXlsx').onclick=async()=>{
  try{
    await loadScript(XLSX_URL); const XL=window.XLSX; const wb=XL.utils.book_new();
    const src=XL.utils.aoa_to_sheet(sourceAoa()); src['!cols']=[{wch:8},{wch:60},...state.years.map(()=>({wch:16}))];
    XL.utils.book_append_sheet(wb,src,'Исходные данные');
    for(const t of exportTables()) XL.utils.book_append_sheet(wb,tableSheet(t),('Табл '+t.num).slice(0,31));
    const buf=XL.write(wb,{bookType:'xlsx',type:'array'});
    await offer(`Анализ_${fileBase()}.xlsx`, new Blob([buf]));
  }catch(e){ msg('Ошибка при сборке Excel: '+e.message,true); }
};
document.getElementById('dlTpl').onclick=async()=>{
  try{
    await loadScript(XLSX_URL); const XL=window.XLSX; const wb=XL.utils.book_new();
    const ws=XL.utils.aoa_to_sheet(sourceAoa(false)); ws['!cols']=[{wch:8},{wch:60},...state.years.map(()=>({wch:16}))];
    XL.utils.book_append_sheet(wb,ws,'Исходные данные');
    await offer('Шаблон_отчетности.xlsx', new Blob([XL.write(wb,{bookType:'xlsx',type:'array'})]));
  }catch(e){ msg('Ошибка: '+e.message,true); }
};

/* ---- Word (.docx) — оформление как в образце «ФУиА задание 1.doc» ---- */
const xe=s=>String(s).replace(/[&<>"]/g,ch=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[ch]));
// sub — подстрочный индекс («К₂₀₂₅»)
function wRun(text,{b,sz=28,i,sub}={}){ return `<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>${b?'<w:b/>':''}${i?'<w:i/>':''}${sub?'<w:vertAlign w:val="subscript"/>':''}<w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/></w:rPr><w:t xml:space="preserve">${xe(text)}</w:t></w:r>`; }
// Абзац текста как в образце: 14 пт, интервал 1,5, отступ первой строки 1,25 см, по ширине, без интервала после.
// pb — начать с новой страницы, keep — не отрывать от следующего абзаца
function wPPr({jc='both',ind=true,after=0,keep,pb}={}){
  return `<w:pPr><w:widowControl w:val="0"/>${keep?'<w:keepNext/>':''}${pb?'<w:pageBreakBefore/>':''}<w:spacing w:after="${after}" w:line="360" w:lineRule="auto"/>${ind?'<w:ind w:firstLine="709"/>':''}<w:jc w:val="${jc}"/></w:pPr>`;
}
function wPara(text,{b,sz=28,...o}={}){ return `<w:p>${wPPr(o)}${wRun(text,{b,sz})}</w:p>`; }
// абзац из нескольких кусков текста: runs = [{t, sub}]
function wParaRuns(runs,{sz=28,...o}={}){ return `<w:p>${wPPr(o)}${runs.map(r=>wRun(r.t,{sz,sub:r.sub})).join('')}</w:p>`; }
const wText = text => (text||'').split(/\n+/).filter(Boolean).map(p=>wPara(p)).join('');

// Таблица по спецификации из docstyle.js: {grid, sz, line, bsz, rowH, head, body}
function wTable(s){
  const rows=[...s.head,...s.body], grid=layout(rows,s.ncols), W=s.grid;
  const bd=['top','left','bottom','right','insideH','insideV'].map(k=>`<w:${k} w:val="single" w:sz="${s.bsz}" w:space="0" w:color="auto"/>`).join('');
  const lineSp = s.line ? 'w:line="360" w:lineRule="auto"' : 'w:line="240" w:lineRule="auto"';
  const markSz = `<w:rPr><w:sz w:val="${s.sz}"/><w:szCs w:val="${s.sz}"/></w:rPr>`;  // кегль знака абзаца как у текста — иначе пустые строки выше, чем в образце
  const small = rows.length<=16;   // небольшую таблицу держим целиком на одной странице
  let x=`<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblInd w:w="108" w:type="dxa"/><w:tblBorders>${bd}</w:tblBorders><w:tblLayout w:type="fixed"/><w:tblLook w:val="04A0"/></w:tblPr><w:tblGrid>${W.map(w=>`<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>`;
  rows.forEach((row,r)=>{
    const isHead=r<s.head.length;
    // строка не разрывается между страницами; шапка повторяется на новой странице
    const trPr='<w:cantSplit/>'+(isHead?'<w:tblHeader/>':'')+(!isHead&&s.rowH?`<w:trHeight w:val="${s.rowH}"/>`:'');
    x+=`<w:tr><w:trPr>${trPr}</w:trPr>`;
    const keep = (small||isHead) && r<rows.length-1 ? '<w:keepNext/>' : '';   // «не отрывать от следующей строки»
    let c=0;
    while(c<s.ncols){
      const g=grid[r][c];
      if(!g){ x+=`<w:tc><w:tcPr><w:tcW w:w="${W[c]}" w:type="dxa"/></w:tcPr><w:p/></w:tc>`; c++; continue; }
      const cell=g.cell, cs=cell.cs||1, w=W.slice(c,c+cs).reduce((a,b)=>a+b,0);
      const gs=cs>1?`<w:gridSpan w:val="${cs}"/>`:'';
      const pEmpty=`<w:p><w:pPr><w:widowControl w:val="0"/>${keep}<w:spacing w:after="0" ${lineSp}/>${markSz}</w:pPr></w:p>`;
      if(g.r0!==r){ x+=`<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/>${gs}<w:vMerge/></w:tcPr>${pEmpty}</w:tc>`; c+=cs; continue; }
      const vm=(cell.rs||1)>1?'<w:vMerge w:val="restart"/>':'';
      const isNum=cell.v!==undefined;
      const text=isNum ? fx(cell.v,cell.d,cell.plus) : cell.t;
      const jc=cell.jc || (isNum?'right':'left');           // числа — по правому краю, текст — по левому
      const va=cell.va || (isNum?'bottom':'center');
      x+=`<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/>${gs}${vm}<w:vAlign w:val="${va}"/></w:tcPr><w:p><w:pPr><w:widowControl w:val="0"/>${keep}<w:spacing w:after="0" ${lineSp}/><w:jc w:val="${jc}"/>${markSz}</w:pPr>${text?wRun(text,{b:cell.b,i:cell.i,sz:s.sz}):''}</w:p></w:tc>`;
      c+=cs;
    }
    x+='</w:tr>';
  });
  return x+'</w:tbl>';
}
// Картинка в тексте (DrawingML inline) в тонкой рамке, как диаграммы в образце. 1 см = 360 000 EMU; 16 × 8 см
function wImage(rid, n, name){
  const cx=16*360000, cy=8*360000;
  return `<w:p><w:pPr><w:keepNext/><w:spacing w:before="240" w:after="0"/><w:jc w:val="center"/></w:pPr><w:r><w:drawing>`+
    `<wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="6350" t="6350" r="6350" b="6350"/><wp:docPr id="${n}" name="${xe(name)}"/>`+
    `<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>`+
    `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>`+
    `<pic:nvPicPr><pic:cNvPr id="${n}" name="${xe(name)}"/><pic:cNvPicPr/></pic:nvPicPr>`+
    `<pic:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>`+
    `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:ln w="6350"><a:solidFill><a:srgbClr val="808080"/></a:solidFill></a:ln></pic:spPr>`+
    `</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`;
}

// Вводные фразы перед таблицами — как в образце; для 2.8–2.10 следом идут расчеты с подстановкой
function leads(){
  const co=state.company||'организации', y0=state.years[0], yN=state.years[nYears()-1];
  return {
    t22:`Проведем горизонтальный анализ актива бухгалтерского баланса предприятия ${co} за ${y0} - ${yN} гг.`,
    t23:`Проведем горизонтальный анализ пассива бухгалтерского баланса ${co} за период с ${y0} по ${yN} годы.`,
    t24:'Проведем вертикальный анализ бухгалтерского баланса. Для этого сумму баланса активов и пассивов принимаем за 100%. Каждую статью выражаем в процентах от общей суммы баланса, т.е. определяем удельный вес каждого показателя активов и пассивов.',
    t25:'Рассмотрим вертикальный анализ пассивов бухгалтерского баланса предприятия.',
    t26:`Проведем сравнительную оценку дебиторской и кредиторской задолженности за период с ${y0} по ${yN} гг. для оценки финансового состояния ${co}.`,
    t27:`Информационной базой для детального анализа прибыли является бухгалтерская отчетность, и в первую очередь форма № 2 «Отчет о финансовых результатах». Анализ формирования прибыли ${co} представлен в таблице 2.7.`,
    t28:`Проведем анализ относительных коэффициентов ликвидности ${co} за ${y0}-${yN} гг.`,
    t29:`Проведем оценку относительных показателей финансовой устойчивости ${co} за ${y0}-${yN} гг.`,
    t210:`Проведем оценку финансовой устойчивости ${co}, для этого сначала определим вспомогательные показатели:`,
    t211:'Для анализа ликвидности баланса сгруппируем активы по степени ликвидности, а пассивы — по срочности оплаты обязательств.',
  };
}
const CALC_BEFORE={t28:'liq', t29:'fin', t210:'abs'};

// Собирает .docx и возвращает Blob (отдельно от кнопки — удобно проверять)
async function buildDocx(){
  await loadScript(JSZIP_URL);
  const tables=exportTables(), L=leads(), co=state.company||'организации';
  // графики: если Chart.js не загрузился (нет интернета), документ соберется без рисунков
  let images=[];
  try{ images=await chartImages(); }catch(e){ status('Рисунки не добавлены: не загрузился Chart.js'); }
  const zip=new JSZip(), rels=[];
  let body=wPara('2.  Горизонтальный и вертикальный анализ бухгалтерской отчетности. Анализ отчетности на основе финансовых коэффициентов.',{jc:'left'});
  body+=wPara('');
  body+=wPara(`Целью анализа финансовой отчетности является получение наиболее информативных параметров, дающих объективную картину финансового состояния ${co}, его имущественного положения, изменений в структуре активов и пассивов, в расчетах с дебиторами и кредиторами, а также результатов деятельности за ряд отчетных периодов.`);
  for(const t of tables){
    // вводная фраза → (расчеты) → «Таблица 2.x - …» → таблица → вывод → рисунки с выводами
    if(L[t.id]) body+=wPara(L[t.id],{keep:!CALC_BEFORE[t.id]});
    if(CALC_BEFORE[t.id]) body+=calcXml(CALC_BEFORE[t.id]);
    body+=wPara('',{ind:false});
    body+=wPara(`Таблица ${t.num} - ${t.title}`,{ind:false,keep:true});
    body+=wTable(t);
    body+=wText(state.conclusions[t.id]);
    for(const im of images.filter(x=>x.after===t.id)){
      const n=rels.length+1, rid='rIdImg'+n, file=`media/chart${n}.png`;
      zip.file('word/'+file, im.png);   // PNG в word/media, связь rIdN — в document.xml.rels
      rels.push(`<Relationship Id="${rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="${file}"/>`);
      body+=wImage(rid, n, `Рисунок ${im.num}`);
      body+=wPara(`Рисунок ${im.num} - ${im.title}`,{ind:false,jc:'center'});
      body+=wText(state.conclusions[im.id]);
    }
  }
  const ns='xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"';
  // поля как в образце: верх/низ 2 см, левое 3 см, правое 1 см
  const doc=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${ns}><w:body>${body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="567" w:bottom="1134" w:left="1701" w:header="709" w:footer="709" w:gutter="0"/></w:sectPr></w:body></w:document>`;
  zip.file('[Content_Types].xml','<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('_rels/.rels','<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  zip.file('word/_rels/document.xml.rels',`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.join('')}</Relationships>`);
  zip.file('word/document.xml',doc);
  return zip.generateAsync({type:'blob'});
}
document.getElementById('dlDocx').onclick=async()=>{
  try{ status('Собираю Word…'); await offer(`Анализ_${fileBase()}.docx`, await buildDocx()); }
  catch(e){ msg('Ошибка при сборке Word: '+e.message,true); }
};
