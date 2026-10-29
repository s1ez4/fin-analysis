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
function tableSheet(t){
  const rows=allRows(t), grid=layout(rows,t.ncols), XL=window.XLSX;
  const aoa=[[`Таблица ${t.num} – ${t.title}`]], merges=[], off=1;
  rows.forEach((row,r)=>{
    const line=[];
    for(let c=0;c<t.ncols;c++){
      const s=grid[r][c];
      if(!s||s.r0!==r||s.c0!==c){ line.push(null); continue; }
      const cell=s.cell;
      if((cell.cs||1)>1||(cell.rs||1)>1) merges.push({s:{r:r+off,c},e:{r:r+off+(cell.rs||1)-1,c:c+(cell.cs||1)-1}});
      if(cell.v!==undefined){
        if(cell.v==null||!Number.isFinite(cell.v)) line.push('—');
        else { const val = cell.k==='num'?Math.round(cell.v):Math.round(cell.v*10**cell.d)/10**cell.d;
          line.push({t:'n',v:val,z:cell.k==='num'?'#,##0':(cell.d===3?'0.000':'0.00')}); }
      } else line.push(cell.t);
    }
    aoa.push(line);
  });
  const concl=state.conclusions[t.id]; if(concl){ aoa.push([]); aoa.push([concl]); }
  const ws=XL.utils.aoa_to_sheet(aoa); ws['!merges']=merges;
  ws['!cols']=[{wch:52},...Array(t.ncols-1).fill({wch:15})];
  return ws;
}
document.getElementById('dlXlsx').onclick=async()=>{
  try{
    await loadScript(XLSX_URL); const XL=window.XLSX; const wb=XL.utils.book_new();
    const src=XL.utils.aoa_to_sheet(sourceAoa()); src['!cols']=[{wch:8},{wch:60},...state.years.map(()=>({wch:16}))];
    XL.utils.book_append_sheet(wb,src,'Исходные данные');
    for(const t of buildTables()) XL.utils.book_append_sheet(wb,tableSheet(t),('Табл '+t.num).slice(0,31));
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

/* ---- Word (.docx) ---- */
const xe=s=>String(s).replace(/[&<>"]/g,ch=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[ch]));
function wRun(text,{b,sz=28,i}={}){ return `<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>${b?'<w:b/>':''}${i?'<w:i/>':''}<w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/></w:rPr><w:t xml:space="preserve">${xe(text)}</w:t></w:r>`; }
function wPara(text,{b,sz=28,jc='both',ind=true,after=120,keep}={}){
  return `<w:p><w:pPr>${keep?'<w:keepNext/>':''}<w:spacing w:after="${after}" w:line="360" w:lineRule="auto"/>${ind?'<w:ind w:firstLine="709"/>':''}<w:jc w:val="${jc}"/></w:pPr>${wRun(text,{b,sz})}</w:p>`;
}
function wTable(t){
  const rows=allRows(t), grid=layout(rows,t.ncols), W=9355;
  const first=Math.round(W*(t.ncols>7?0.26:0.36)), other=Math.floor((W-first)/(t.ncols-1));
  const widths=[first,...Array(t.ncols-1).fill(other)];
  const bd=['top','left','bottom','right','insideH','insideV'].map(s=>`<w:${s} w:val="single" w:sz="4" w:space="0" w:color="000000"/>`).join('');
  let x=`<w:tbl><w:tblPr><w:tblW w:w="${W}" w:type="dxa"/><w:tblBorders>${bd}</w:tblBorders><w:tblLayout w:type="fixed"/><w:tblCellMar><w:left w:w="57" w:type="dxa"/><w:right w:w="57" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid>${widths.map(w=>`<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>`;
  rows.forEach((row,r)=>{
    x+=`<w:tr>${row[0]&&row[0].hd?'<w:trPr><w:tblHeader/></w:trPr>':''}`;
    let c=0;
    while(c<t.ncols){
      const s=grid[r][c];
      if(!s){ x+=`<w:tc><w:tcPr><w:tcW w:w="${widths[c]}" w:type="dxa"/></w:tcPr><w:p/></w:tc>`; c++; continue; }
      const cell=s.cell, cs=cell.cs||1, w=widths.slice(c,c+cs).reduce((a,b)=>a+b,0);
      const gs=cs>1?`<w:gridSpan w:val="${cs}"/>`:'';
      if(s.r0!==r){ x+=`<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/>${gs}<w:vMerge/></w:tcPr><w:p/></w:tc>`; c+=cs; continue; }
      const vm=(cell.rs||1)>1?'<w:vMerge w:val="restart"/>':'';
      const isNum=cell.v!==undefined, bold=cell.hd||cell.sec||cell.tt;
      const jc=cell.hd?'center':(isNum||(c>0&&!cell.sec))?'center':'left';
      const shade=cell.hd||cell.sec?'<w:shd w:val="clear" w:color="auto" w:fill="EDEDED"/>':'';
      x+=`<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/>${gs}${vm}${shade}<w:vAlign w:val="center"/></w:tcPr><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/><w:jc w:val="${jc}"/></w:pPr>${wRun(fmt(cell),{b:bold,sz:t.ncols>7?18:20})}</w:p></w:tc>`;
      c+=cs;
    }
    x+='</w:tr>';
  });
  return x+'</w:tbl>';
}
document.getElementById('dlDocx').onclick=async()=>{
  try{
    await loadScript(JSZIP_URL);
    const tables=buildTables();
    let body=wPara(`Анализ бухгалтерской отчетности ${state.company||''}`.trim(),{b:true,jc:'center',ind:false,sz:32,after:240});
    for(const t of tables){
      body+=wPara(`Таблица ${t.num} – ${t.title}`,{ind:false,jc:'left',keep:true,after:60});
      body+=wTable(t);
      body+=wPara('',{after:60,ind:false});
      const c=state.conclusions[t.id];
      if(c) for(const p of c.split(/\n+/).filter(Boolean)) body+=wPara(p);
      body+=wPara('',{after:120,ind:false});
    }
    const doc=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="850" w:bottom="1134" w:left="1701" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>`;
    const zip=new JSZip();
    zip.file('[Content_Types].xml','<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
    zip.file('_rels/.rels','<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
    zip.file('word/document.xml',doc);
    const blob=await zip.generateAsync({type:'blob'});
    await offer(`Анализ_${fileBase()}.docx`, blob);
  }catch(e){ msg('Ошибка при сборке Word: '+e.message,true); }
};
