#!/usr/bin/env python3
# Применяет быструю сборку Excel-файлов к index.html (3 правки). Запуск: python3 patch_speed.py
# Положите скрипт в одну папку с index.html. Рядом создаётся копия index.html.bak.
import os, sys, shutil

HERE = os.path.dirname(os.path.abspath(__file__))
PATH = os.path.join(HERE, 'index.html')
if not os.path.exists(PATH):
    sys.exit('Не найден index.html рядом со скриптом')
src = open(PATH, encoding='utf-8').read()
if 'buildFileFast' in src:
    sys.exit('Правка уже применена — ничего не делаю')

def replace_once(text, old, new, name):
    n = text.count(old)
    if n != 1:
        sys.exit('Правка «%s»: нужный фрагмент найден %d раз(а), а ожидался 1. index.html отличается от версии, под которую сделан патч — файл НЕ изменён.' % (name, n))
    return text.replace(old, new)

# --- 1. parseXlsx: сохраняем книгу и признак «сложного» файла ---
OLD1 = "  return{buf,sheets,names,notes};"
NEW1 = r"""  const slow=sheets.some(S=>{const w=wb.getWorksheet(S.name);return !!((w.conditionalFormattings||[]).length||(w.getImages&&w.getImages().length))});   // условное форматирование и картинки быстрый способ не переносит
  return{buf,wb,sheets,names,notes,slow};"""
src = replace_once(src, OLD1, NEW1, '1: parseXlsx')

# --- 2. buildFile: быстрый способ + прежний как запасной ---
OLD2 = "async function buildFile(D,rec,hide){"
NEW2 = r"""async function buildFile(D,rec,hide){   // быстрый способ; при любой ошибке или «сложном» файле — прежний
  if(!D.slow){try{return await buildFileFast(D,rec,hide)}catch(e){console.warn('Быстрая сборка не удалась, использую прежний способ:',e)}}
  return buildFileSlow(D,rec,hide);
}
async function buildFileFast(D,rec,hide){   // книга D.wb читается один раз; для получателя создаётся новая книга только с его строками
  await EJ();const out=new ExcelJS.Workbook(),rng=/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/;
  for(const S of D.sheets){
    const mine=rec.get(S.name);if(!mine||!mine.length)continue;
    const src=D.wb.getWorksheet(S.name),nc=S.nc,map=[];let j=0;
    for(let c=1;c<=nc;c++)map[c]=hide.has(nk(S.heads[c]||''))?0:++j;
    const cm=c=>map[c]||0,v0=(src.views&&src.views[0])||null;
    const ws=out.addWorksheet(S.name,{views:v0&&v0.state==='frozen'&&v0.ySplit?[{state:'frozen',xSplit:0,ySplit:v0.ySplit}]:(v0&&v0.showGridLines===false?[{showGridLines:false}]:undefined)});
    for(let c=1;c<=nc;c++){if(!map[c])continue;const sc=src.getColumn(c),dc=ws.getColumn(map[c]);if(sc.width)dc.width=sc.width;if(sc.hidden)dc.hidden=true}
    const cp=(s,d,v)=>{d.value=v;const st=s.style;if(st&&Object.keys(st).length)d.style=Object.assign({},st)};
    for(let r=1;r<=S.hr;r++){const sr=src.getRow(r),dr=ws.getRow(r);if(sr.height)dr.height=sr.height;
      for(let c=1;c<=nc;c++){if(!map[c])continue;const s=sr.getCell(c);cp(s,dr.getCell(map[c]),s.value)}}
    Object.values(src._merges||{}).forEach(m=>{const g=m.model;if(g.bottom>S.hr)return;const ks=[];for(let c=g.left;c<=g.right;c++)if(map[c])ks.push(map[c]);if(!ks.length)return;
      const l=Math.min(...ks),r=Math.max(...ks);if(r>l||g.bottom>g.top){try{ws.mergeCells(g.top,l,g.bottom,r)}catch(e){}}});
    mine.forEach((row,i)=>{const o=row.n,w=S.hr+1+i,sr=src.getRow(o),dr=ws.getRow(w);if(sr.height)dr.height=sr.height;
      for(let c=1;c<=nc;c++){if(!map[c])continue;const s=sr.getCell(c);let v=s.value;
        if(s.type===ExcelJS.ValueType.Formula){const nf=shiftF(s.formula,o,w,S.hr,cm);v=nf?{formula:nf,result:s.result}:(s.result===undefined?null:s.result)}
        cp(s,dr.getCell(map[c]),v)}});
    const last=S.hr+mine.length;let hasTbl=false;
    Object.keys(src.tables||{}).forEach(name=>{const t=src.tables[name].table,m=rng.exec(t.tableRef||'');if(!m)return;
      const c1=colN(m[1]),r1=+m[2],c2=colN(m[3]),kept=[];for(let c=c1;c<=c2;c++)if(map[c])kept.push(c);
      if(r1!==S.hr||!kept.length)return;
      const had=!!t.autoFilterRef,rows=[];for(let r=S.hr+1;r<=last;r++){const dr=ws.getRow(r);rows.push(kept.map(c=>dr.getCell(map[c]).value))}
      ws.addTable({name:t.name||name,ref:colL(map[kept[0]])+r1,headerRow:true,totalsRow:false,style:t.style,
        columns:kept.map(c=>({name:String(t.columns[c-c1].name),filterButton:had?(t.columns[c-c1].filterButton!==false):false})),rows});hasTbl=true});
    const af=src.autoFilter;
    if(af&&!hasTbl){let c1,r1,c2;
      if(typeof af==='string'){const m=rng.exec(af);if(m){c1=colN(m[1]);r1=+m[2];c2=colN(m[3])}}
      else if(af.from&&af.to){c1=af.from.column;r1=af.from.row;c2=af.to.column}
      const kept=[];if(c1)for(let c=c1;c<=c2;c++)if(map[c])kept.push(c);
      if(c1&&r1===S.hr&&kept.length)ws.autoFilter=colL(map[kept[0]])+r1+':'+colL(map[kept[kept.length-1]])+last}
    if(ws.rowCount!==last)throw new Error('число строк не совпало');
  }
  return new Uint8Array(await out.xlsx.writeBuffer());
}
async function buildFileSlow(D,rec,hide){"""
src = replace_once(src, OLD2, NEW2, '2: buildFile')

# --- 3. mSave: полная проверка только для «сложных» файлов и каждого 25-го ---
OLD3 = "        try{await verifyFile(D,rec,hide,u8,a)}catch(e){throw"
NEW3 = "        if(D.slow||i%25===1)try{await verifyFile(D,rec,hide,u8,a)}catch(e){throw"
src = replace_once(src, OLD3, NEW3, '3: verifyFile')

shutil.copyfile(PATH, PATH + '.bak')
open(PATH, 'w', encoding='utf-8', newline='').write(src)
print('Готово: index.html обновлён (копия — index.html.bak). Не забудьте залить index.html и sw.js на GitHub.')
