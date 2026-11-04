'use strict';
(()=>{
 const NS='http://schemas.openxmlformats.org/spreadsheetml/2006/main';
 const REL='http://schemas.openxmlformats.org/officeDocument/2006/relationships';
 const RPKG='http://schemas.openxmlformats.org/package/2006/relationships';
 const el=id=>document.getElementById(id);
 const qs=(sel,root=document)=>root.querySelector(sel);
 const xmlChildren=(element,local)=>Array.from(element?.children||[]).filter(n=>n.localName===local);
 const first=(element,local)=>xmlChildren(element,local)[0]||null;
 const all=(element,local)=>element?Array.from(element.getElementsByTagNameNS(NS,local)):[];
 const escapeHtml=v=>String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
 const asText=x=>x==null?'':String(x);
 const colLetter=num=>{let s='';for(;num>0;num=Math.floor((num-1)/26))s=String.fromCharCode(65+(num-1)%26)+s;return s;};
 const colIndex=s=>{let n=0;for(const c of s.toUpperCase())n=n*26+c.charCodeAt(0)-64;return n;};
 const addr=(row,col)=>colLetter(col)+row;
 const colFromAddress=s=>colIndex((s.match(/^[A-Z]+/i)||[''])[0]);
 const rowFromAddress=s=>Number((s.match(/\d+/)||['0'])[0]);
 const niceNum=n=>new Intl.NumberFormat('vi-VN').format(n);
 const message=(text,kind='')=>{el('importStatus').textContent=text;el('importStatus').className='status '+kind;};
 const toast=(m)=>{const t=el('toast');t.textContent=m;t.classList.add('show');clearTimeout(toast.timer);toast.timer=setTimeout(()=>t.classList.remove('show'),3000);};
 window.efToast=toast;
 const sleep=()=>new Promise(r=>setTimeout(r,10));
 const namedBlob=(bytes,name)=>new Blob([bytes],{type:name.toLowerCase().endsWith('.csv')?'text/csv;charset=utf-8':'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
 const download=(blob,name)=>{const u=URL.createObjectURL(blob);const a=document.createElement('a');a.href=u;a.download=name;a.style.display='none';document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(u),20000);};
 window.efDownload=download;
 function xmlParse(text){const doc=new DOMParser().parseFromString(text,'application/xml');const err=doc.getElementsByTagName('parsererror')[0];if(err)throw Error('Tệp Excel có XML không hợp lệ: '+err.textContent.slice(0,140));return doc;}
 function normalizePath(p){const arr=[];for(const bit of p.replaceAll('\\','/').split('/')){if(bit==='..')arr.pop();else if(bit&&bit!=='.')arr.push(bit);}return arr.join('/');}
 function pathFromRel(target){const t=target.replace(/^\//,'');return normalizePath(target.startsWith('/')?t:'xl/'+t);}
 function normalizeHeader(s){return String(s??'').normalize('NFC').trim().replace(/\s+/g,' ').toLocaleLowerCase('vi-VN');}
 function readString(cell,shared){
  if(!cell)return null;
  const typ=cell.getAttribute('t')||'n';const v=first(cell,'v');
  const f=first(cell,'f');
  if(typ==='inlineStr') {const ins=first(cell,'is');return {kind:'s',raw:ins?all(ins,'t').map(t=>t.textContent).join(''):'',fromFormula:!!f};}
  if(!v||v.textContent==='') return null;
  if(typ==='s') return {kind:'s',raw:shared[Number(v.textContent)]??'',fromFormula:!!f};
  if(typ==='str'||typ==='d')return {kind:typ==='d'?'d':'s',raw:v.textContent,fromFormula:!!f};
  if(typ==='b')return {kind:'b',raw:v.textContent,fromFormula:!!f};
  if(typ==='e')return {kind:'e',raw:v.textContent,fromFormula:!!f};
  return {kind:'n',raw:v.textContent,fromFormula:!!f};
 }
 function formatValue(o){if(!o)return '(trống)';if(o.kind==='b')return o.raw==='1'?'ĐÚNG':'SAI';if(o.kind==='n'&&o.raw!==''&&Number.isFinite(Number(o.raw)))return String(Number(o.raw));return asText(o.raw);}
 function equalValues(a,b){if(!a&&!b)return true;if(!a||!b)return false;if(a.kind==='e'||b.kind==='e')return a.kind===b.kind&&a.raw===b.raw;if(a.kind==='n'&&b.kind==='n')return Number(a.raw)===Number(b.raw);return a.kind===b.kind&&a.raw===b.raw;}
 function extractStyleFormats(doc){
  const formats={};if(!doc)return [];
  const nfs=all(doc.documentElement,'numFmt');for(const x of nfs)formats[Number(x.getAttribute('numFmtId'))]=x.getAttribute('formatCode');
  const cellXfs=all(doc.documentElement,'cellXfs')[0];return cellXfs?xmlChildren(cellXfs,'xf').map(x=>formats[Number(x.getAttribute('numFmtId'))]||''):[];
 }
 function keyOf(value,cell,styles){if(!value||value.kind==='e')return null;let s=asText(value.raw).trim();if(!s)return null;
  if(value.kind==='n'&&Number.isInteger(Number(s))){s=String(Number(s));const styleNum=Number(cell?.getAttribute('s')||0),fmt=styles[styleNum]||'';if(/^0{2,}$/.test(fmt))s=s.padStart(fmt.length,'0');}
  return s.toLocaleLowerCase('vi-VN');
 }
 function extractLineCells(items, gapThreshold = 9) {
  if (!items.length) return [];
  const cells = [];
  let currentCell = { text: items[0].str, x: items[0].tx, xEnd: items[0].tx + (items[0].width || 0) };
  for (let i = 1; i < items.length; i++) {
    const it = items[i];
    const itWidth = it.width || (it.str.length * 6);
    const gap = it.tx - currentCell.xEnd;
    if (gap < gapThreshold) {
      const needSpace = gap > 2 && !currentCell.text.endsWith(' ') && !it.str.startsWith(' ');
      currentCell.text += (needSpace ? ' ' : '') + it.str;
      currentCell.xEnd = Math.max(currentCell.xEnd, it.tx + itWidth);
    } else {
      if (currentCell.text.trim()) {
        cells.push({ text: currentCell.text.trim(), x: currentCell.x, xEnd: currentCell.xEnd });
      }
      currentCell = { text: it.str, x: it.tx, xEnd: it.tx + itWidth };
    }
  }
  if (currentCell.text.trim()) {
    cells.push({ text: currentCell.text.trim(), x: currentCell.x, xEnd: currentCell.xEnd });
  }
  return cells;
 }

 function clusterColumns(allLines) {
  const allX = [];
  for (const line of allLines) {
    for (const cell of line) {
      allX.push(cell.x);
    }
  }
  allX.sort((a,b) => a - b);
  const clusters = [];
  for (const x of allX) {
    let found = false;
    for (const cl of clusters) {
      if (Math.abs(cl.avg - x) <= 20) {
        cl.points.push(x);
        cl.avg = cl.points.reduce((a,b)=>a+b,0) / cl.points.length;
        found = true;
        break;
      }
    }
    if (!found) {
      clusters.push({ points: [x], avg: x });
    }
  }
  clusters.sort((a,b) => a.avg - b.avg);
  const minPoints = Math.max(1, Math.floor(allLines.length * 0.05));
  const validClusters = clusters.filter(c => c.points.length >= minPoints || clusters.length <= 16);
  return (validClusters.length ? validClusters : clusters).map(c => c.avg);
 }

 function assembleCellText(items) {
  if (!items || !items.length) return '';
  const sorted = [...items].sort((a,b) => a.vy - b.vy || a.vx - b.vx);
  const sublines = [];
  let curSub = [];
  let curY = null;
  for (const it of sorted) {
    if (curY === null || Math.abs(it.vy - curY) > 3) {
      if (curSub.length) sublines.push(curSub);
      curSub = [it];
      curY = it.vy;
    } else {
      curSub.push(it);
    }
  }
  if (curSub.length) sublines.push(curSub);

  return sublines.map(line => {
    line.sort((a,b) => a.vx - b.vx);
    let lineStr = '';
    for (let i = 0; i < line.length; i++) {
      const it = line[i];
      if (i > 0) {
        const prev = line[i-1];
        const gap = it.vx - (prev.vx + (prev.w || 0));
        if (gap >= 2 && !lineStr.endsWith(' ') && !it.str.startsWith(' ')) {
          lineStr += ' ';
        }
      }
      lineStr += it.str;
    }
    return lineStr.replace(/\s+/g, ' ').trim();
  }).filter(Boolean).join(' ');
 }

 function sheetsToXlsxZip(sheetsInput) {
  const zip = new JSZip();
  const sheets = Array.isArray(sheetsInput) && sheetsInput.length && sheetsInput[0].grid
    ? sheetsInput
    : [{ name: 'Bảng PDF', grid: sheetsInput }];

  let contentTypesOverride = '';
  let relsEntries = '';
  let wbSheets = '';

  for (let sIdx = 0; sIdx < sheets.length; sIdx++) {
    const sId = sIdx + 1;
    const rId = 'rId' + sId;
    const sheetInfo = sheets[sIdx];
    const grid = sheetInfo.grid || [];

    const maxRow = Math.max(1, grid.length);
    const maxCol = Math.max(1, ...grid.map(r => r.length));

    let rowsXml = '';
    for (let r = 0; r < grid.length; r++) {
      const rowNum = r + 1;
      let cellsXml = '';
      const row = grid[r];
      for (let c = 0; c < row.length; c++) {
        const val = row[c];
        if (val == null || val === '') continue;
        const ref = colLetter(c + 1) + rowNum;
        const sVal = String(val).trim();
        const escVal = sVal.replace(/[&<>"']/g, ch => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[ch]));
        cellsXml += `<c r="${ref}" t="inlineStr"><is><t xml:space="preserve">${escVal}</t></is></c>`;
      }
      if (cellsXml) {
        rowsXml += `<row r="${rowNum}">${cellsXml}</row>`;
      }
    }

    zip.file(`xl/worksheets/sheet${sId}.xml`, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:${colLetter(maxCol)}${maxRow}"/><sheetData>${rowsXml}</sheetData></worksheet>`);
    contentTypesOverride += `<Override PartName="/xl/worksheets/sheet${sId}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`;
    relsEntries += `<Relationship Id="${rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${sId}.xml"/>`;
    wbSheets += `<sheet name="${sheetInfo.name.replace(/[&<>"']/g, '')}" sheetId="${sId}" r:id="${rId}"/>`;
  }

  const stylesRId = 'rId' + (sheets.length + 1);
  relsEntries += `<Relationship Id="${stylesRId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>`;
  contentTypesOverride += `<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>`;

  zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>${contentTypesOverride}</Types>`);
  zip.file('_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`);
  zip.file('xl/_rels/workbook.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relsEntries}</Relationships>`);
  zip.file('xl/workbook.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>${wbSheets}</sheets></workbook>`);
  zip.file('xl/styles.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs></styleSheet>`);

  return zip;
 }

 function gridToXlsxZip(grid) {
  return sheetsToXlsxZip([{ name: 'Bảng PDF', grid }]);
 }

 async function parsePdfToWorkbook(file) {
  if (typeof window.pdfjsLib === 'undefined') {
    throw Error('Chưa tải được thư viện PDF (pdf.min.js). Vui lòng làm mới trang.');
  }
  window.pdfjsLib.GlobalWorkerOptions.workerSrc = 'js/pdf.worker.min.js';
  const arrayBuffer = await file.arrayBuffer();
  const pdfDoc = await window.pdfjsLib.getDocument({ data: arrayBuffer }).promise;

  const pagesData = [];
  let fullPdfText = '';

  for (let p = 1; p <= pdfDoc.numPages; p++) {
    const page = await pdfDoc.getPage(p);
    const viewport = page.getViewport({ scale: 1.0 });
    const content = await page.getTextContent();
    const items = [];
    for (const it of (content.items || [])) {
      if (!it.str || !it.str.trim()) continue;
      // Lọc bỏ barcode ẩn và số trang chân trang
      if (it.str.includes('TCT-BARCODE') || /Trang\s+\d+\/\d+/i.test(it.str)) continue;
      const [vx, vy] = viewport.convertToViewportPoint(it.transform[4], it.transform[5]);
      if (vy < 30) continue; // Khu vực barcode lề trên
      const itemObj = { str: it.str, vx: Math.round(vx), vy: Math.round(vy), w: Math.round(it.width || (it.str.length * 6)) };
      items.push(itemObj);
      fullPdfText += ' ' + it.str;
    }
    pagesData.push({ p, items });
  }

  // Nhận diện tệp Tờ khai thuế GTGT có bảng kê bán ra (01-1) và mua vào (01-2)
  const hasTaxHeaders = /01-1\/GTGT|01-2\/GTGT|BẢNG KÊ HOÁ ĐƠN|BÁN RA|MUA VÀO/i.test(fullPdfText);

  if (hasTaxHeaders) {
    const sheetsResult = [];
    const muaVaoPages = [];
    const banRaPages = [];

    let currentSection = null;
    for (const pd of pagesData) {
      const pageTxt = pd.items.map(i => i.str).join(' ');
      // Trang 1 thường là Tờ khai chính (Mẫu 01/GTGT) không chứa bảng kê hoá đơn chi tiết -> bỏ qua trang tờ khai
      const isMainDeclarationPage = /TỜ KHAI THUẾ GIÁ TRỊ GIA TĂNG/i.test(pageTxt) && !/BẢNG KÊ HOÁ ĐƠN/i.test(pageTxt);
      if (isMainDeclarationPage) continue;

      if (/01-2\/GTGT|HÀNG HOÁ.*?MUA VÀO|BẢNG KÊ.*?MUA VÀO/i.test(pageTxt)) {
        currentSection = 'muavao';
      } else if (/01-1\/GTGT|HÀNG HOÁ.*?BÁN RA|BẢNG KÊ.*?BÁN RA/i.test(pageTxt)) {
        currentSection = 'banra';
      }
      if (currentSection === 'muavao') {
        muaVaoPages.push(pd);
      } else if (currentSection === 'banra') {
        banRaPages.push(pd);
      }
    }

    // 1. Trích xuất Bảng kê Mua Vào (Mẫu 01-2/GTGT)
    if (muaVaoPages.length > 0) {
      const colBoundsMuaVao = [
        { name: 'STT', min: 40, max: 76 },
        { name: 'Ký hiệu mẫu hóa đơn', min: 76, max: 125 },
        { name: 'Ký hiệu hóa đơn', min: 125, max: 185 },
        { name: 'Số hoá đơn', min: 185, max: 245 },
        { name: 'Ngày phát hành', min: 245, max: 304 },
        { name: 'Tên người bán', min: 304, max: 398 },
        { name: 'Mã số thuế người bán', min: 398, max: 462 },
        { name: 'Mặt hàng', min: 462, max: 545 },
        { name: 'Tiền chưa thuế', min: 545, max: 625 },
        { name: 'Thuế suất (%)', min: 625, max: 660 },
        { name: 'Thuế GTGT', min: 660, max: 725 },
        { name: 'Ghi chú', min: 725, max: 840 }
      ];
      const gridMV = [colBoundsMuaVao.map(c => c.name)];
      for (const pd of muaVaoPages) {
        const items = pd.items;
        const dateItems = items.filter(it => it.vx >= 245 && it.vx <= 304 && /^\d{2}\/\d{2}\/\d{4}$/.test(it.str.trim())).sort((a,b) => a.vy - b.vy);
        let baselines = dateItems.map(d => d.vy);
        if (!baselines.length) {
          const invItems = items.filter(it => it.vx >= 185 && it.vx <= 245 && /^\d+$/.test(it.str.trim())).sort((a,b) => a.vy - b.vy);
          baselines = invItems.map(d => d.vy);
        }
        for (let i = 0; i < baselines.length; i++) {
          const curY = baselines[i];
          const prevY = (i > 0) ? baselines[i - 1] : (curY - 22);
          const nextY = (i + 1 < baselines.length) ? baselines[i + 1] : (curY + 22);
          const rowMinY = (prevY + curY) / 2;
          const rowMaxY = (curY + nextY) / 2;
          const rowItems = items.filter(it => it.vy >= rowMinY && it.vy < rowMaxY);
          const row = new Array(colBoundsMuaVao.length).fill('');
          for (let c = 0; c < colBoundsMuaVao.length; c++) {
            const cellItems = rowItems.filter(it => it.vx >= colBoundsMuaVao[c].min && it.vx < colBoundsMuaVao[c].max);
            let val = assembleCellText(cellItems);
            if (c === 1) val = val.replace(/(\d)\s+(\d)/g, '$1$2');
            row[c] = val;
          }
          if (row[3] || row[4] || row[8]) gridMV.push(row);
        }
      }
      if (gridMV.length > 1) {
        sheetsResult.push({ name: 'Bảng kê mua vào (Mẫu 01-2)', grid: gridMV });
      }
    }

    // 2. Trích xuất Bảng kê Bán Ra (Mẫu 01-1/GTGT)
    if (banRaPages.length > 0) {
      const colBoundsBanRa = [
        { name: 'STT', min: 40, max: 76 },
        { name: 'Ký hiệu mẫu hóa đơn', min: 76, max: 135 },
        { name: 'Ký hiệu hóa đơn', min: 135, max: 195 },
        { name: 'Số hoá đơn', min: 195, max: 260 },
        { name: 'Ngày phát hành', min: 260, max: 325 },
        { name: 'Tên người mua', min: 325, max: 420 },
        { name: 'Mã số thuế người mua', min: 420, max: 475 },
        { name: 'Mặt hàng', min: 475, max: 560 },
        { name: 'Doanh số bán chưa có thuế', min: 560, max: 640 },
        { name: 'Thuế GTGT', min: 640, max: 720 },
        { name: 'Ghi chú', min: 720, max: 820 }
      ];
      const gridBR = [colBoundsBanRa.map(c => c.name)];
      for (const pd of banRaPages) {
        const items = pd.items;
        const dateItems = items.filter(it => it.vx >= 260 && it.vx <= 325 && /^\d{2}\/\d{2}\/\d{4}$/.test(it.str.trim())).sort((a,b) => a.vy - b.vy);
        let baselines = dateItems.map(d => d.vy);
        if (!baselines.length) {
          const invItems = items.filter(it => it.vx >= 195 && it.vx <= 260 && /^\d+$/.test(it.str.trim())).sort((a,b) => a.vy - b.vy);
          baselines = invItems.map(d => d.vy);
        }
        for (let i = 0; i < baselines.length; i++) {
          const curY = baselines[i];
          const prevY = (i > 0) ? baselines[i - 1] : (curY - 22);
          const nextY = (i + 1 < baselines.length) ? baselines[i + 1] : (curY + 22);
          const rowMinY = (prevY + curY) / 2;
          const rowMaxY = (curY + nextY) / 2;
          const rowItems = items.filter(it => it.vy >= rowMinY && it.vy < rowMaxY);
          const row = new Array(colBoundsBanRa.length).fill('');
          for (let c = 0; c < colBoundsBanRa.length; c++) {
            const cellItems = rowItems.filter(it => it.vx >= colBoundsBanRa[c].min && it.vx < colBoundsBanRa[c].max);
            let val = assembleCellText(cellItems);
            if (c === 1) val = val.replace(/(\d)\s+(\d)/g, '$1$2');
            row[c] = val;
          }
          if (row[3] || row[4] || row[8]) gridBR.push(row);
        }
      }
      if (gridBR.length > 1) {
        sheetsResult.push({ name: 'Bảng kê bán ra (Mẫu 01-1)', grid: gridBR });
      }
    }

    if (sheetsResult.length > 0) {
      const zip = sheetsToXlsxZip(sheetsResult);
      zip._sheets = sheetsResult;
      return zip;
    }
  }

  // 3. Fallback cho PDF sao kê ngân hàng hoặc tài liệu thông thường
  const allLines = [];
  for (const pd of pagesData) {
    const lineMap = new Map();
    for (const it of pd.items) {
      let matchedY = null;
      for (const y of lineMap.keys()) {
        if (Math.abs(y - it.vy) <= 4.0) {
          matchedY = y;
          break;
        }
      }
      if (matchedY === null) {
        lineMap.set(it.vy, [{ str: it.str, tx: it.vx, width: it.w }]);
      } else {
        lineMap.get(matchedY).push({ str: it.str, tx: it.vx, width: it.w });
      }
    }
    const sortedYs = Array.from(lineMap.keys()).sort((a,b) => a - b);
    for (const y of sortedYs) {
      const lineItems = lineMap.get(y).sort((a,b) => a.tx - b.tx);
      const cells = extractLineCells(lineItems, 9);
      if (cells.length) allLines.push(cells);
    }
  }

  if (!allLines.length) {
    throw Error('File PDF "' + file.name + '" không tìm thấy lớp văn bản. Hãy dùng file PDF xuất từ ngân hàng hoặc phần mềm.');
  }

  const cols = clusterColumns(allLines);
  const grid = [];
  for (const line of allLines) {
    const row = new Array(cols.length).fill('');
    for (const cell of line) {
      let bestCol = 0, minDist = Infinity;
      for (let c = 0; c < cols.length; c++) {
        const dist = Math.abs(cols[c] - cell.x);
        if (dist < minDist) {
          minDist = dist;
          bestCol = c;
        }
      }
      row[bestCol] = (row[bestCol] ? row[bestCol] + ' ' : '') + cell.text;
    }
    grid.push(row);
  }

  const fallbackSheets = [{ name: 'Bảng PDF', grid }];
  const zip = sheetsToXlsxZip(fallbackSheets);
  zip._sheets = fallbackSheets;
  return zip;
 }
 window.efParsePdfToWorkbook = parsePdfToWorkbook;

 async function openWorkbook(file){
  if(!file||!/\.(xlsx|xlsm|pdf)$/i.test(file.name))throw Error('Chỉ hỗ trợ .xlsx, .xlsm hoặc .pdf. Hãy chọn đúng định dạng.');
  let zip;
  if(/\.pdf$/i.test(file.name)){
    zip = await parsePdfToWorkbook(file);
  } else {
    const bytes=await file.arrayBuffer();
    zip=await JSZip.loadAsync(bytes);
  }
  const workbookFile=zip.file('xl/workbook.xml'),relnFile=zip.file('xl/_rels/workbook.xml.rels');
  if(!workbookFile||!relnFile)throw Error('Không đọc được cấu trúc file Excel (tệp có thể bị khóa mật khẩu).');
  const [workbookXml,relsXml,sharedXml,stylesXml]=await Promise.all([
   workbookFile.async('string'),relnFile.async('string'),zip.file('xl/sharedStrings.xml')?.async('string')||null,zip.file('xl/styles.xml')?.async('string')||null
  ]);
  const wbDoc=xmlParse(workbookXml),relationships=xmlParse(relsXml),relnames=new Map();
  for(const r of relationships.documentElement.children){if(r.localName==='Relationship'&&!r.getAttribute('TargetMode'))relnames.set(r.getAttribute('Id'),pathFromRel(r.getAttribute('Target')||''));}
  const sheets=[];for(const sh of all(wbDoc.documentElement,'sheet')){const path=relnames.get(sh.getAttributeNS(REL,'id')||sh.getAttribute('r:id'));if(path&&zip.file(path))sheets.push({name:sh.getAttribute('name'),path});}
  if(!sheets.length)throw Error('Không tìm thấy sheet dữ liệu nào.');
  const shared=sharedXml?all(xmlParse(sharedXml).documentElement,'si').map(si=>all(si,'t').map(t=>t.textContent).join('')):[];
  const styles=stylesXml?extractStyleFormats(xmlParse(stylesXml)):[];
  const wp=first(wbDoc.documentElement,'workbookPr');const date1904=['1','true'].includes((wp?.getAttribute('date1904')||'').toLowerCase());
  return {file,zip,wbDoc,sheets,shared,styles,date1904,sheetCache:new Map()};
 }
 async function sheetOf(book,name){
  if(book.sheetCache.has(name))return book.sheetCache.get(name);
  const info=book.sheets.find(s=>s.name===name);if(!info)throw Error('Không tìm thấy sheet: '+name);
  const raw=await book.zip.file(info.path).async('string');const doc=xmlParse(raw);const sheetData=all(doc.documentElement,'sheetData')[0];if(!sheetData)throw Error('Sheet không có vùng dữ liệu.');
  const rows=new Map(),cells=new Map(),colNums=new Set();let maxRow=0,maxCol=0;
  for(const row of xmlChildren(sheetData,'row')){
   const rowNum=Number(row.getAttribute('r'));if(!rowNum)continue;rows.set(rowNum,row);maxRow=Math.max(maxRow,rowNum);
   for(const c of xmlChildren(row,'c')){const ref=c.getAttribute('r')||'';const col=colFromAddress(ref);if(!col)continue;cells.set(ref.toUpperCase(),c);colNums.add(col);maxCol=Math.max(maxCol,col);}
  }
  const merges=[];for(const group of all(doc.documentElement,'mergeCell')){const bounds=(group.getAttribute('ref')||'').split(':');if(bounds.length<2)continue;const top=rowFromAddress(bounds[0]),bottom=rowFromAddress(bounds[1]),left=colFromAddress(bounds[0]),right=colFromAddress(bounds[1]);merges.push({top,bottom,left,right});}
  const result={name,path:info.path,doc,sheetData,rows,cells,maxRow,maxCol,merges,colNums};book.sheetCache.set(name,result);return result;
 }
 function cellAt(sheet,row,col){return sheet.cells.get(addr(row,col))||null;}
 function mergeAt(sheet,row,col){for(const m of sheet.merges){if(row>=m.top&&row<=m.bottom&&col>=m.left&&col<=m.right)return m;}return null;}
 function mergeRef(m){return m?addr(m.top,m.left)+':'+addr(m.bottom,m.right):'';}
 function readLocation(sheet,row,col){const m=mergeAt(sheet,row,col);return m?{row:m.top,col:m.left,merge:m}:{row,col,merge:null};}
 function readCellAt(sheet,row,col){const r=readLocation(sheet,row,col);return cellAt(sheet,r.row,r.col);}
 function cellValue(sheet,book,row,col){return readString(readCellAt(sheet,row,col),book.shared);}
 function getColumns(sheet,book,headerRow,limit=20){
  const count=Math.min(160,Math.max(limit,sheet.maxCol));const list=[];
  for(let c=1;c<=count;c++){
   const m=mergeAt(sheet,headerRow,c);
   if(m&&m.left!==c)continue;  // One visible choice per merged header range.
   const val=cellValue(sheet,book,headerRow,c),label=val?formatValue(val).trim().replace(/\s+/g,' '):'';
   const range=m&&m.right>m.left?colLetter(m.left)+'–'+colLetter(m.right):colLetter(c);
   list.push({num:c,letter:colLetter(c),header:label,merge:m,display:range+(label?' · '+label.slice(0,54):' · (cột '+colLetter(c)+')')});
  }
  return list;
 }
 function populateSelect(id,columns,keep){const sel=el(id),prev=keep?sel.value:'';sel.replaceChildren();for(const c of columns){const opt=document.createElement('option');opt.value=String(c.num);opt.textContent=c.display;sel.appendChild(opt);}if(prev&&columns.some(c=>String(c.num)===prev))sel.value=prev;}
 function selectSheets(id,book){const sel=el(id);sel.replaceChildren();for(const s of book.sheets){const opt=document.createElement('option');opt.value=s.name;opt.textContent=s.name;sel.appendChild(opt);}}
 function detectHeaderRow(sheet, book, maxScan = 35) {
  if (!sheet || !sheet.rows) return 1;
  const keywords = [
    'stt', 'serial', 'số', 'so', 'ngày', 'ngay', 'date', 'tên', 'ten', 'khách', 'khach',
    'mst', 'mã', 'ma', 'tiền', 'tien', 'thuế', 'thue', 'tổng', 'tong', 'nội dung', 'noi dung',
    'diễn giải', 'dien giai', 'hóa đơn', 'hoa don', 'mặt hàng', 'mat hang', 'tk', 'tài khoản',
    'đối tác', 'doi tac', 'đơn vị', 'don vi', 'doanh số', 'doanh so'
  ];
  let bestRow = 1, maxScore = -1;
  const limit = Math.min(sheet.maxRow || 35, maxScan);
  for (let r = 1; r <= limit; r++) {
    const cols = [];
    for (let c = 1; c <= Math.min(sheet.maxCol || 30, 40); c++) {
      const v = cellValue(sheet, book, r, c);
      const text = (v?.kind === 's' || typeof v?.raw === 'string') ? String(v.raw).trim().toLowerCase() : '';
      if (text) cols.push(text);
    }
    if (!cols.length) continue;
    let kwHits = 0;
    for (const text of cols) {
      if (keywords.some(kw => text.includes(kw))) kwHits++;
    }
    const score = kwHits * 4 + cols.length;
    if (kwHits >= 2 && score > maxScore) {
      maxScore = score;
      bestRow = r;
    } else if (cols.length >= 3 && score > maxScore && maxScore < 5) {
      maxScore = score;
      bestRow = r;
    }
  }
  return bestRow;
 }

 async function autoDetectAndSetTargetHeader(sheetName) {
  if (!dstBook) return;
  const name = sheetName || el('targetSheet').value;
  if (!name) return;
  try {
    const sh = await sheetOf(dstBook, name);
    const detected = detectHeaderRow(sh, dstBook);
    const hdrInput = el('targetHeaderRow');
    const tag = el('targetHeaderTag');
    if (hdrInput) {
      hdrInput.value = detected;
    }
    if (tag) {
      tag.textContent = `Tự nhận diện dòng ${detected}`;
      tag.style.display = 'inline-flex';
    }
    let startRow = detected + 1;
    if (isTotalRow(sh, dstBook, startRow, totalTerms(el('totalTerms').value))) {
      startRow++;
    }
    const startInput = el('targetStart');
    if (startInput && el('targetMode').value === 'auto') {
      startInput.value = startRow;
    }
  } catch (err) {
    console.warn('Cannot detect target header:', err);
  }
 }

 async function autoDetectAndSetSourceHeader(sheetName) {
  if (!srcBook) return;
  const name = sheetName || el('sourceSheet').value;
  if (!name) return;
  try {
    const sh = await sheetOf(srcBook, name);
    const isPdf = srcEntries.some(e => e.file && /\.pdf$/i.test(e.file.name));
    const detected = isPdf ? 1 : detectHeaderRow(sh, srcBook);
    const hdrInput = el('headerRow');
    const tag = el('sourceHeaderTag');
    if (hdrInput) {
      hdrInput.value = detected;
    }
    if (tag) {
      tag.textContent = isPdf ? 'Tự nhận diện dòng 1 (PDF)' : `Tự nhận diện dòng ${detected}`;
      tag.style.display = 'inline-flex';
    }
    const startInput = el('sourceStart');
    if (startInput) {
      startInput.value = detected + 1;
    }
  } catch (err) {
    console.warn('Cannot detect source header:', err);
  }
 }

 let srcBook=null,dstBook=null,plan=null,planSettings=null,lastErr=null;let loading=false;
 let srcEntries=[],sourceCounter=0,sourceQueue=Promise.resolve();
 async function refreshCols(){try{
  const hdr=num(el('headerRow')?.value || '1','Dòng tiêu đề nguồn'),dh=num(el('targetHeaderRow')?.value || '1','Dòng tiêu đề đích');
  if(srcBook){const sh=await sheetOf(srcBook,el('sourceSheet').value);const sc=getColumns(sh,srcBook,hdr);populateSelect('addSource',sc,true);populateSelect('sourceKey',sc,true);}
  if(dstBook){const sh=await sheetOf(dstBook,el('targetSheet').value);const tc=getColumns(sh,dstBook,dh);populateSelect('addTarget',tc,true);populateSelect('targetKey',tc,true);}
  renderMappingList();
  if(srcBook&&dstBook&&typeof window.efAutoMapColumns==='function'){window.efAutoMapColumns();}
 }catch(e){message(e.message,'err');}}
 function invalidate(){plan=null;planSettings=null;el('exportBtn').disabled=true;el('previewLabel').textContent='Chưa phân tích';el('statChanges').textContent='0';el('statRows').textContent='0';el('statSkipped').textContent='0';el('showSkipped').disabled=true;el('skipDetails').hidden=true;el('skipQuickText').textContent='Sau khi xem trước, bạn có thể biết chính xác dòng hoặc ô nào không được nhập.';el('downloadSkips').disabled=true;el('previewTable').style.display='none';el('noPreview').style.display='block';el('reconDownload').disabled=true;el('reconBadge').textContent='Cần phân tích lại';el('reconBadge').classList.remove('warn');el('reconTable').innerHTML='';el('reconDifferences').innerHTML='';el('reconRows').textContent='0';el('reconMismatches').textContent='0';el('reconPairs').textContent='0';el('reconStatus').textContent='Các quy tắc đã thay đổi. Hãy xem trước dữ liệu một lần nữa để tính đối chiếu chính xác.';el('reconStatus').className='status warn';el('fileResults').style.display='none';}
 function num(str,name){const n=Number(str);if(!Number.isInteger(n)||n<1||n>1048576)throw Error(name+' phải là số nguyên từ 1 đến 1048576.');return n;}
 function mappingColumn(token,columns,side){
  const clean=token.trim().replace(/^["']|["']$/g,'');const letter=/^([A-Z]{1,3})$/i.exec(clean);
  if(letter){const n=colIndex(clean);if(n>=1&&n<=16384)return n;}
  const matches=columns.filter(c=>c.header&&normalizeHeader(c.header)===normalizeHeader(clean));
  if(matches.length===1)return matches[0].num;
  if(matches.length>1)throw Error('Tiêu đề "'+clean+'" bị trùng tại '+side+'. Hãy dùng chữ cột.');
  throw Error('Không tìm thấy cột "'+clean+'" trong '+side+'. Hãy dùng chữ cột (A,B,C…) hoặc tên tiêu đề chính xác.');
 }
 async function extractSettings(){
  if(!srcEntries.length||!srcBook||!dstBook)throw Error('Hãy chọn ít nhất một file nguồn và một file đích.');
  if(srcEntries.some(x=>x.file===dstBook.file||x.file.name===dstBook.file.name&&x.file.size===dstBook.file.size&&x.file.lastModified===dstBook.file.lastModified))throw Error('File đích không được trùng với bất kỳ file nguồn nào.');
  const headerRow=num(el('headerRow').value,'Dòng tiêu đề nguồn'),targetHeaderRow=num(el('targetHeaderRow').value,'Dòng tiêu đề đích');
  const sourceStart=num(el('sourceStart').value,'Dòng đầu tiên nguồn');
  const sourceEnd=el('sourceEnd').value.trim()?num(el('sourceEnd').value,'Dòng cuối cùng nguồn'):1048576;
  if(sourceEnd<sourceStart)throw Error('Dòng cuối cùng của nguồn phải lớn hơn hoặc bằng dòng đầu tiên.');
  const targetMode=el('targetMode').value,targetStart=targetMode==='manual'?num(el('targetStart').value,'Dòng bắt đầu đích'):Math.min(1048576,targetHeaderRow+1);
  const src=srcBook.sheetCache.get(el('sourceSheet').value),dst=dstBook.sheetCache.get(el('targetSheet').value);
  if(!src||!dst)throw Error('Hãy chọn đúng sheet nguồn và sheet đích.');
  const sc=getColumns(src,srcBook,headerRow),tc=getColumns(dst,dstBook,targetHeaderRow);
  const mappings=[],dstUsed=new Set(),instructions=el('mapping').value.trim();
  if(!instructions)throw Error('Hãy ghép ít nhất một cặp cột bằng ô chọn ở bước 2.');
  for(const [i,line] of instructions.split('\n').entries()){
   if(!line.trim()||/^#/.test(line.trim()))continue;
   const match=line.match(/^\s*(.*?)\s*(?:->|=>|→|:|=)\s*(.*?)\s*$/);
   if(!match||!match[1]||!match[2])throw Error('Quy tắc cột dòng '+(i+1)+' chưa đúng. Ví dụ: M -> D');
   let source=mappingColumn(match[1],sc,'file nguồn'),target=mappingColumn(match[2],tc,'file đích');
   const hm=mergeAt(src,headerRow,source);if(hm)source=hm.left;
   const tm=mergeAt(dst,targetHeaderRow,target);if(tm)target=tm.left;
   if(dstUsed.has(target))throw Error('Cột đích '+colLetter(target)+' được chọn nhiều lần. Hãy kiểm tra phần ghép cột.');
   dstUsed.add(target);mappings.push({source,target});
  }
  if(!mappings.length)throw Error('Cần ít nhất một cặp cột nguồn → đích.');
  const sources=[];const autoRange=el('autoSourceRange').checked;
  for(const [index,entry] of srcEntries.entries()){
   const sourceBook=entry.book,sheetName=index===0?el('sourceSheet').value:entry.sheet;
   if(!sourceBook.sheets.some(s=>s.name===sheetName))throw Error('File '+entry.file.name+' chưa chọn được sheet nguồn hợp lệ.');
   const sheet=await sheetOf(sourceBook,sheetName),suggestion=autoRange?detectStatement(sheet,sourceBook):null;
   const h=suggestion?suggestion.header:headerRow;
   const start=suggestion?suggestion.start:sourceStart,end=suggestion?suggestion.end:Math.min(sourceEnd,sheet.maxRow);
   if(start>end)throw Error('Không tìm thấy dòng giao dịch hợp lệ trong '+entry.file.name+'; hãy kiểm tra sheet và vùng dữ liệu.');
   for(const m of mappings){
    const a=cellValue(src,srcBook,headerRow,m.source),b=cellValue(sheet,sourceBook,h,m.source);
    const headA=a?.kind==='s'?normalizeHeader(a.raw):'',headB=b?.kind==='s'?normalizeHeader(b.raw):'';
    if(headA&&headB&&headA!==headB)throw Error('Cột '+colLetter(m.source)+' ở "'+entry.file.name+'" có tiêu đề khác file mẫu ("'+headB+'" thay vì "'+headA+'"). Hãy kiểm tra lại sheet hoặc cấu trúc file.');
    if(headA&&!headB)throw Error('File "'+entry.file.name+'" không tìm thấy tiêu đề tại cột '+colLetter(m.source)+'. Hãy chọn đúng sheet cùng cấu trúc.');
   }
   sources.push({src:sheet,book:sourceBook,name:entry.file.name,index,sourceSheet:sheetName,start,end,header:h});
  }
  const matchMode=el('matchMode').value,sourceKey=Number(el('sourceKey').value),targetKey=Number(el('targetKey').value);
  if(matchMode==='key'&&(!sourceKey||!targetKey))throw Error('Hãy chọn 2 cột dùng để dò mã chung.');
  return {src,dst,sources,sourceSheet:src.name,targetSheet:dst.name,sourceStart,sourceEnd,autoRange,targetStart,targetMode,headerRow,targetHeaderRow,mappings,matchMode,sourceKey,targetKey,writeMode:el('writeMode').value,skipBlank:el('skipBlank').checked,skipEmpty:el('skipEmpty').checked,reverse:el('reverse').checked,
    skipTotals:el('skipTotals').checked,skipZeros:el('skipZeros')?.checked||false,totalTerms:totalTerms(el('totalTerms').value),rowFilter:el('rowFilter').value,formatRules:parseFormats(el('formatRules').value)};
 }

 // Import filters, explicit data types and Excel number formats.
 const deaccent=s=>String(s??'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/đ/g,'d').replace(/Đ/g,'D').replace(/\s+/g,' ').trim().toLowerCase();
 function totalTerms(raw){return raw.split(',').map(s=>deaccent(s)).filter(Boolean);}
 function isTotalRow(sheet,book,row,terms){
  if(!terms.length)return false;
  const rowNode=sheet.rows.get(row);if(!rowNode)return false;
  for(const c of xmlChildren(rowNode,'c')){
   const v=readString(c,book.shared);if(!v||v.kind!=='s')continue;
   const text=deaccent(v.raw);if(!text)continue;
   for(const term of terms){
    if(text===term||text===term+':'||text===term+' :'||text===term+' -')return true;
    if(term.length>=7&&text.startsWith(term)&&(text.length===term.length||/^[\s:\-–(]/.test(text.slice(term.length))))return true;
   }
  }
  return false;
 }
 function hasFormula(sheet,row){const node=sheet.rows.get(row);return !!node&&xmlChildren(node,'c').some(c=>!!first(c,'f'));}
 const FORMAT_INFO={auto:{type:'auto',code:null},text:{type:'text',code:'@'},number:{type:'number',code:'General'},integer:{type:'number',code:'#,##0'},decimal:{type:'number',code:'#,##0.00'},money:{type:'number',code:'#,##0 "₫"'},percent:{type:'percent',code:'0.00%'},date:{type:'date',code:'[$-42A]dd/mm/yyyy;@'},'date-iso':{type:'date',code:'yyyy-mm-dd'},datetime:{type:'datetime',code:'[$-42A]dd/mm/yyyy hh:mm:ss;@'}};
 function parseRange(scope){
  const v=scope.trim().toUpperCase().replaceAll('$','');let m;
  if((m=/^([A-Z]{1,3})$/.exec(v)))return {r1:1,r2:1048576,c1:colIndex(m[1]),c2:colIndex(m[1])};
  if((m=/^([A-Z]{1,3}):([A-Z]{1,3})$/.exec(v)))return {r1:1,r2:1048576,c1:colIndex(m[1]),c2:colIndex(m[2])};
  if((m=/^(\d+)$/.exec(v)))return {r1:Number(m[1]),r2:Number(m[1]),c1:1,c2:16384};
  if((m=/^(\d+):(\d+)$/.exec(v)))return {r1:Number(m[1]),r2:Number(m[2]),c1:1,c2:16384};
  if((m=/^([A-Z]{1,3})(\d+)$/.exec(v)))return {r1:Number(m[2]),r2:Number(m[2]),c1:colIndex(m[1]),c2:colIndex(m[1])};
  if((m=/^([A-Z]{1,3})(\d+):([A-Z]{1,3})(\d+)$/.exec(v)))return {r1:Number(m[2]),r2:Number(m[4]),c1:colIndex(m[1]),c2:colIndex(m[3])};
  throw Error('Vùng định dạng không hợp lệ: '+scope+'. Ví dụ D, 5:10 hoặc D5:F20.');
 }
 function parseFormats(raw){const result=[];
  for(const [i,line] of raw.split('\n').entries()){
   if(!line.trim()||line.trim().startsWith('#'))continue;
   const m=/^\s*(.*?)\s*(?:->|=>|→|=)\s*(.*?)\s*$/.exec(line);
   if(!m||!m[1]||!m[2])throw Error('Quy tắc định dạng dòng '+(i+1)+' sai. Ví dụ D -> date.');
   const bounds=parseRange(m[1]);if(!Number.isInteger(bounds.r1)||!Number.isInteger(bounds.r2)||bounds.r1<1||bounds.r2>1048576||bounds.r1>bounds.r2||bounds.c1<1||bounds.c2>16384||bounds.c1>bounds.c2)throw Error('Vùng Excel sai tại quy tắc '+(i+1));
   const rawSpec=m[2].trim(),spec=FORMAT_INFO[rawSpec.toLowerCase()]||(rawSpec.toLowerCase().startsWith('custom:')?{type:'custom',code:rawSpec.slice(7).trim()}:null);
   if(!spec||spec.type==='custom'&&(!spec.code||spec.code.length>140))throw Error('Kiểu định dạng không hợp lệ tại dòng '+(i+1)+': '+rawSpec);
   result.push({...bounds,scope:m[1],format:rawSpec,type:spec.type,code:spec.code});
  }return result;
 }
 function formatFor(rules,row,col){let selected=null;for(const x of rules)if(row>=x.r1&&row<=x.r2&&col>=x.c1&&col<=x.c2)selected=x;return selected;}
 function parseNumber(raw,reference){
  let s=String(raw).trim().replace(/[\s\u00a0\u202f]/g,'').replace(/(?:vnd|đ|₫)/ig,'');const negative=/^\(.*\)$/.test(s);if(negative)s=s.slice(1,-1);
  if(!/^[+-]?[\d.,]+%?$/.test(s))throw Error('Không thể đổi '+reference+' thành số: "'+String(raw)+'"');
  const percent=s.endsWith('%');if(percent)s=s.slice(0,-1);
  const dot=s.lastIndexOf('.'),comma=s.lastIndexOf(',');
  if(dot!==-1&&comma!==-1){const decimal=dot>comma?'.':',';s=s.replaceAll(decimal==='.'?',':'.','');s=s.replace(decimal,'.');}
  else if(dot!==-1||comma!==-1){const symbol=dot!==-1?'.':',';const esc=symbol==='.'?'\\.':',';
   const grouped=new RegExp('^\\d{1,3}('+esc+'\\d{3})+$').test(s);
   if(grouped&&!/^0[.,]/.test(s))s=s.replaceAll(symbol,'');
   else if(s.split(symbol).length===2)s=s.replace(symbol,'.');
   else s=s.replaceAll(symbol,'');
  }
  const n=Number(s)*(negative?-1:1);if(!Number.isFinite(n))throw Error('Số không hợp lệ tại '+reference);
  return percent?n/100:n;
 }
 function serialFromDate(y,m,d,dest1904,reference){
  const ms=Date.UTC(y,m-1,d),dt=new Date(ms);if(dt.getUTCFullYear()!==y||dt.getUTCMonth()+1!==m||dt.getUTCDate()!==d)throw Error('Ngày không hợp lệ tại '+reference);
  const serial=Math.round((ms-Date.UTC(1899,11,30))/86400000)-(dest1904?1462:0);
  if(serial<0||serial>2958465)throw Error('Ngày vượt phạm vi Excel ở '+reference);
  return serial;
 }
 function normalizeImport(v,rule,sourceCell,reference,sourceBook=srcBook){
  if(!rule||rule.type==='auto')return v;
  if(rule.type==='text'){
   let text=asText(v.raw);if(v.kind==='n'){const fmt=sourceBook.styles[Number(sourceCell?.getAttribute('s')||0)]||'';if(/^0{2,}$/.test(fmt))text=text.padStart(fmt.length,'0');}
   return {kind:'s',raw:text};
  }
  if(rule.type==='date'||rule.type==='datetime'){
    if(v.kind==='n'){const shift=(sourceBook.date1904?1462:0)-(dstBook.date1904?1462:0);const serial=Number(v.raw)+shift;return {kind:'n',raw:String(rule.type==='date'?Math.floor(serial):serial)};}
    const str=asText(v.raw).trim();let y,m,d,hh=0,mi=0,ss=0,mm;
    if((mm=/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})(?:[T ]+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/.exec(str))){y=+mm[1];m=+mm[2];d=+mm[3];hh=+(mm[4]||0);mi=+(mm[5]||0);ss=+(mm[6]||0);}
    else if((mm=/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2,4})(?:[ T]+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/.exec(str))){d=+mm[1];m=+mm[2];y=+mm[3];if(y<100)y+=y<50?2000:1900;hh=+(mm[4]||0);mi=+(mm[5]||0);ss=+(mm[6]||0);}
    else throw Error('Không đọc được ngày ở '+reference+': '+str+'. Ví dụ: 21/06/2026 11:40:57.');
    if(hh>23||mi>59||ss>59)throw Error('Giờ/phút/giây không hợp lệ tại '+reference);
    const serial=serialFromDate(y,m,d,dstBook.date1904,reference);
    return {kind:'n',raw:String(serial+(rule.type==='datetime'?(hh*3600+mi*60+ss)/86400:0))};
   }
  if(rule.type==='custom'&&rule.code==='@')return {kind:'s',raw:String(v.raw)};
  if(rule.type==='number'||rule.type==='percent'||rule.type==='custom'){
   if(rule.type==='custom'&&/[^#0.,%\s;_\-+"€₫đ$E\[\]a-zA-Z]/.test(rule.code))throw Error('Mã định dạng tùy chỉnh chứa ký tự không được hỗ trợ: '+rule.code);
   const percentSymbol=String(v.raw).includes('%'),value=v.kind==='n'?Number(v.raw):parseNumber(v.raw,reference);
   if(!Number.isFinite(value))throw Error('Số không hợp lệ ở '+reference);
   const numeric=rule.type==='percent'&&!percentSymbol&&Math.abs(value)>1?value/100:value;
   return {kind:'n',raw:String(numeric)};
  }
  return v;
 }
 function displayImport(v,rule){
  if(v&&rule&&(rule.type==='date'||rule.type==='datetime')&&v.kind==='n'&&Number.isFinite(Number(v.raw))){
    const serial=Number(v.raw),days=Math.floor(serial)+(dstBook.date1904?1462:0);
    const d=new Date(Date.UTC(1899,11,30)+days*86400000),dd=String(d.getUTCDate()).padStart(2,'0'),mm=String(d.getUTCMonth()+1).padStart(2,'0'),yyyy=d.getUTCFullYear();
    if(rule.type==='datetime'){
     const sec=Math.round((serial-Math.floor(serial))*86400)%86400;
     return dd+'/'+mm+'/'+yyyy+' '+String(Math.floor(sec/3600)).padStart(2,'0')+':'+String(Math.floor(sec%3600/60)).padStart(2,'0')+':'+String(sec%60).padStart(2,'0');
    }
    return rule.code==='yyyy-mm-dd'?yyyy+'-'+mm+'-'+dd:dd+'/'+mm+'/'+yyyy;
   }
  return formatValue(v);
 }
 function numericValue(v,rule){
  if(!v||v.kind==='e'||rule&&(['text','date','datetime'].includes(rule.type)))return null;
  try{const n=v.kind==='n'?Number(v.raw):parseNumber(v.raw,'đối chiếu');return Number.isFinite(n)?n:null;}catch{return null;}
 }
 function getStyleManager(zip){
  let doc=null,xfs=null,numFmts=null,formatMap=new Map(),styleCache=new Map(),nextFmt=164;
  async function prepare(){const file=zip.file('xl/styles.xml');if(!file)return;
   doc=xmlParse(await file.async('string'));const root=doc.documentElement;xfs=first(root,'cellXfs');
   numFmts=first(root,'numFmts');if(!numFmts){numFmts=doc.createElementNS(NS,'numFmts');root.insertBefore(numFmts,root.firstElementChild);numFmts.setAttribute('count','0');}
   for(const f of xmlChildren(numFmts,'numFmt')){
     const id=Number(f.getAttribute('numFmtId'));
     let code=f.getAttribute('formatCode')||'';
     if(/年|月|日|\[\$-411\]|\[\$-804\]|\[\$-0411\]|\[\$-0804\]/i.test(code)){
       code='[$-42A]dd/mm/yyyy;@';
       f.setAttribute('formatCode',code);
     }
     if(id>=164&&code)formatMap.set(code,id);
     nextFmt=Math.max(nextFmt,id+1);
   }
   const viDateCode='[$-42A]dd/mm/yyyy;@';
   let viDateFmtId=formatMap.get(viDateCode);
   if(!viDateFmtId){
     viDateFmtId=nextFmt++;
     const f=doc.createElementNS(NS,'numFmt');
     f.setAttribute('numFmtId',String(viDateFmtId));
     f.setAttribute('formatCode',viDateCode);
     numFmts.appendChild(f);
     numFmts.setAttribute('count',String(xmlChildren(numFmts,'numFmt').length));
     formatMap.set(viDateCode,viDateFmtId);
   }
   if(xfs){
     const cjkBuiltinDates=new Set([14,30,31,55,56,57,58]);
     for(const xf of xmlChildren(xfs,'xf')){
       const nid=Number(xf.getAttribute('numFmtId'));
       if(cjkBuiltinDates.has(nid)){
         xf.setAttribute('numFmtId',String(viDateFmtId));
         xf.setAttribute('applyNumberFormat','1');
       }
     }
   }
  }
  function assign(cell,code){if(!code||!doc||!xfs)return;
   const original=Number(cell.getAttribute('s')||0),key=original+'|'+code;
   if(styleCache.has(key)){cell.setAttribute('s',String(styleCache.get(key)));return;}
   const builtin=code==='General'?0:code==='@'?49:null;let fid=builtin??formatMap.get(code);
   if(fid==null){fid=nextFmt++;const f=doc.createElementNS(NS,'numFmt');f.setAttribute('numFmtId',String(fid));f.setAttribute('formatCode',code);numFmts.appendChild(f);numFmts.setAttribute('count',String(xmlChildren(numFmts,'numFmt').length));formatMap.set(code,fid);}
   const styles=xmlChildren(xfs,'xf');const xf=(styles[original]||styles[0]||doc.createElementNS(NS,'xf')).cloneNode(true);xf.setAttribute('numFmtId',String(fid));xf.setAttribute('applyNumberFormat','1');
   xfs.appendChild(xf);const index=styles.length;xfs.setAttribute('count',String(index+1));styleCache.set(key,index);cell.setAttribute('s',String(index));
  }
  return {prepare,assign,async save(){if(doc)zip.file('xl/styles.xml',new XMLSerializer().serializeToString(doc));}};
 }

 function isInsideMerged(sheet,row,col){for(const m of sheet.merges){if(row>=m.top&&row<=m.bottom&&col>=m.left&&col<=m.right)return !(row===m.top&&col===m.left);}return false;}

  const SKIP_REASONS={
    total_source:['Dòng tổng cộng ở file nguồn','Đúng quy tắc nếu bạn chỉ muốn nhập giao dịch.'],
    total_target:['Bỏ qua dòng tổng cộng ở file đích','Không ghi vào dòng tổng. Kiểm tra vùng đích nếu cần.'],
    empty_row:['Dòng nguồn không có dữ liệu ở các cột đã chọn','Kiểm tra sao kê nếu bạn cho rằng dòng này có giao dịch.'],
    source_blank:['Ô nguồn trống','Kiểm tra ô nguồn nếu cột này bắt buộc phải có dữ liệu.'],
    dest_formula:['Ô đích đang có công thức','Công thức được giữ nguyên. Hãy chọn cột đích khác nếu cần nhập.'],
    dest_merged:['Ô đích là phần tiếp theo của ô gộp nhiều dòng','Chọn ô đầu của vùng gộp hoặc chỉnh vùng đích.'],
    dest_exists:['Ô đích đã chứa dữ liệu','Đang bật chỉ nhập ô trống. Kiểm tra giá trị trước khi đổi lựa chọn.'],
    same_value:['Giá trị nguồn trùng với ô đích','Không cần nhập lại khi dữ liệu đã giống nhau.'],
    no_target:['Giao dịch chưa có dòng đích phù hợp','Tạo thêm dòng trống phía trên Tổng cộng, hoặc kiểm tra cách tìm dòng.'],
    no_key:['Không tìm thấy mã đối chiếu ở file nguồn','Kiểm tra cột mã giao dịch trong nguồn.'],
    dup_source_key:['Mã giao dịch bị trùng ở file nguồn','Kiểm tra giao dịch trùng trước khi nhập.'],
    dup_target_key:['Mã giao dịch bị trùng ở file đích','Làm rõ mã trùng trong file đích trước khi nhập.'],
    missing_target_key:['Không tìm thấy mã nguồn trong file đích','Kiểm tra mã hoặc bổ sung dòng cần ghép vào file đích.']
  };
  function skipRecord(category,unit,ctx={},extra={}){
    const detail=SKIP_REASONS[category]||[category,'Kiểm tra lại dữ liệu.'];
    return {category,unit,reason:detail[0],advice:detail[1],file:ctx.file||'',sheet:ctx.sheet||'',source:ctx.source||'',dest:ctx.dest||'',sourceValue:ctx.sourceValue??'',targetValue:ctx.targetValue??'',...extra};
  }
  function skipCsv(details){
    const rows=[['Loại','Nguyên nhân','File nguồn','Sheet nguồn','Dòng / ô nguồn','Dòng / ô đích','Giá trị nguồn','Giá trị đích cũ','Gợi ý kiểm tra']];
    for(const d of details)rows.push([d.unit==='cell'?'Ô':d.unit==='unmatched'?'Giao dịch chưa ghép':'Dòng',d.reason,d.file,d.sheet,d.source,d.dest,d.sourceValue,d.targetValue,d.advice]);
    return '\uFEFF'+rows.map(row=>row.map(csvCell).join(';')).join('\r\n');
  }
  function analyze(cfg){
   const {dst}=cfg;
   const summary={examined:0,matched:0,unmatched:0,dupSrc:0,dupDst:0,skippedFormula:0,skippedBlank:0,skippedExists:0,skippedSame:0,skippedMerge:0,skippedEmpty:0,skippedError:0,skippedTotalSource:0,skippedTotalTarget:0,skippedFormulaRows:0};
   const skipDetails=[];
   const fileStats=cfg.sources.map(x=>({name:x.name,sheet:x.sourceSheet,read:0,matched:0,unmatched:0,excludedTotals:0,range:x.start+'–'+x.end}));
   let sourceRows=[];
   for(const e of cfg.sources){
    let rows=Array.from(e.src.rows.keys()).filter(r=>r>=e.start&&r<=e.end).sort((a,b)=>a-b);
    if(cfg.skipTotals)rows=rows.filter(r=>{if(isTotalRow(e.src,e.book,r,cfg.totalTerms)){summary.skippedTotalSource++;fileStats[e.index].excludedTotals++;
     skipDetails.push(skipRecord('total_source','row',{file:e.name,sheet:e.sourceSheet,source:'Dòng '+r}));return false;}return true;});
    if(cfg.reverse&&cfg.matchMode==='position')rows.reverse();
    for(const sr of rows)sourceRows.push({sr,src:e.src,sourceBook:e.book,sourceName:e.name,sourceIndex:e.index});
   }
   const allowedTarget=r=>{
     if(cfg.skipTotals&&isTotalRow(dst,dstBook,r,cfg.totalTerms))return false;
     if(cfg.rowFilter==='withFormula'&&!hasFormula(dst,r))return false;
     if(cfg.rowFilter==='withoutFormula'&&hasFormula(dst,r))return false;
     return true;
   };
   let eligibleTargetRows=[];
   if(cfg.matchMode==='position'&&cfg.targetMode==='auto'){
     // Keep target totals intact, but do not let top summary rows (adjacent to header) block data insertion
     let effectiveStart = cfg.targetStart;
     while (effectiveStart <= dst.maxRow && !allowedTarget(effectiveStart) && isTotalRow(dst, dstBook, effectiveStart, cfg.totalTerms)) {
       effectiveStart++;
     }
     const boundaries = cfg.skipTotals
       ? Array.from(dst.rows.keys()).filter(r => r > effectiveStart && isTotalRow(dst, dstBook, r, cfg.totalTerms))
       : [];
     const firstTotal = boundaries.length ? Math.min(...boundaries) : 0;
     const desired = sourceRows.length;
     // In multi-section templates (like BKMV2014 with Q1, Q2, Q3, Q4 totals), boundaries.length > 1.
     // We allow filling template rows across sections while allowedTarget(r) protects each total & formula row!
     const multiSection = boundaries.length > 1;
     const limit = Math.min(1048576, (!multiSection && firstTotal) ? firstTotal - 1 : Math.max(dst.maxRow + desired + 5, effectiveStart + desired + 10));
     for(let r=effectiveStart;r<=limit&&eligibleTargetRows.length<desired;r++){
       if(!allowedTarget(r))continue;
       if(cfg.mappings.every(m=>{
         const loc=readLocation(dst,r,m.target);
         if(loc.row!==r)return false; // vertical merge continuation belongs to an earlier row
         const cell=cellAt(dst,loc.row,loc.col),v=cellValue(dst,dstBook,r,m.target);
         return (!cell||!first(cell,'f'))&&(!v||v.raw===''||v.raw==null);
       }))eligibleTargetRows.push(r);
     }
     summary.autoFound=eligibleTargetRows.length;
     summary.autoBlockedByTotal=(!multiSection && firstTotal > 0 && eligibleTargetRows.length < desired);
   }else if(cfg.matchMode==='position'&&cfg.rowFilter!=='all'){
     eligibleTargetRows=Array.from(dst.rows.keys()).filter(r=>r>=cfg.targetStart&&allowedTarget(r)).sort((a,b)=>a-b);
   }
   const targetKeys=new Map(),ambiguousDst=new Set(),srcCounts=new Map();
   if(cfg.matchMode==='key'){
     for(const {sr,src,sourceBook} of sourceRows){const v=cellValue(src,sourceBook,sr,cfg.sourceKey),k=keyOf(v,readCellAt(src,sr,cfg.sourceKey),sourceBook.styles);if(k)srcCounts.set(k,(srcCounts.get(k)||0)+1);}
     for(const [tr] of dst.rows){if(tr<cfg.targetStart||!allowedTarget(tr))continue;const v=cellValue(dst,dstBook,tr,cfg.targetKey),k=keyOf(v,readCellAt(dst,tr,cfg.targetKey),dstBook.styles);if(!k)continue;if(targetKeys.has(k))ambiguousDst.add(k);else targetKeys.set(k,tr);}
     for(const k of ambiguousDst)targetKeys.delete(k);
     summary.dupDst=ambiguousDst.size;summary.dupSrc=Array.from(srcCounts.values()).filter(x=>x>1).length;
   }
   const changes=[],reconcile=cfg.mappings.map(m=>({source:colLetter(m.source),target:colLetter(m.target),sourceSum:0,beforeSum:0,afterSum:0,numericRows:0,missingAfter:0,differentBefore:0,differentAfter:0,changed:0}));
   const differences=[];let position=0,nextTarget=cfg.targetStart;
   for(const {sr,src,sourceBook,sourceName,sourceIndex} of sourceRows){
     summary.examined++;fileStats[sourceIndex].read++;
     if(cfg.skipEmpty&&!cfg.mappings.some(m=>{const v=cellValue(src,sourceBook,sr,m.source);return v&&v.kind!=='e'&&v.raw!=='';})){
       summary.skippedEmpty++;skipDetails.push(skipRecord('empty_row','row',{file:sourceName,sheet:cfg.sources[sourceIndex].sourceSheet,source:'Dòng '+sr}));continue;}
      if(cfg.skipZeros){
        let isZero = false;
        for(const m of cfg.mappings){
          const raw = cellValue(src, sourceBook, sr, m.source);
          if(!raw || raw.raw === '' || raw.kind === 'e') continue;
          const s = String(raw.raw).trim();
          const clean = s.replace(/[^\d.-]/g, '');
          const valNum = Number(clean);
          const isNumZero = (!isNaN(valNum) && valNum === 0 && clean.length > 0) || s === '0' || s === '0.0' || s === '0,0' || s === '0.00' || s === '0,00' || s === '0 ₫' || s === '0 đ';
          if(isNumZero){
            const targetHeader = cellValue(dst, dstBook, cfg.targetHeaderRow, m.target);
            const sourceHeader = cellValue(src, sourceBook, cfg.headerRow, m.source);
            const combinedHeaders = (targetHeader?.raw || '') + ' ' + (sourceHeader?.raw || '');
            if(/tiền|amount|giá|phát\s*sinh|nợ|có|thu|chi/i.test(combinedHeaders) || cfg.mappings.length <= 3){
              isZero = true;
              break;
            }
          }
        }
        if(isZero){
          summary.skippedZeros = (summary.skippedZeros || 0) + 1;
          skipDetails.push(skipRecord('zero_value', 'row', { file: sourceName, sheet: cfg.sources[sourceIndex].sourceSheet, source: 'Dòng ' + sr, dest: 'Số tiền = 0' }));
          continue;
        }
      }
     let tr;
     if(cfg.matchMode==='position'){
       if(cfg.rowFilter!=='all'||cfg.targetMode==='auto'){
         if(position>=eligibleTargetRows.length){summary.unmatched++;fileStats[sourceIndex].unmatched++;summary.skippedFormulaRows++;
          skipDetails.push(skipRecord('no_target','unmatched',{file:sourceName,sheet:cfg.sources[sourceIndex].sourceSheet,source:'Dòng '+sr}));continue;}
         tr=eligibleTargetRows[position++];
       }else{
         tr=nextTarget++;
         while(tr<=1048576&&!allowedTarget(tr)){summary.skippedTotalTarget++;
          skipDetails.push(skipRecord('total_target','row',{file:sourceName,sheet:cfg.sources[sourceIndex].sourceSheet,source:'Dòng '+sr,dest:'Dòng '+tr}));tr=nextTarget++;}
       }
     }else{
       const k=keyOf(cellValue(src,sourceBook,sr,cfg.sourceKey),readCellAt(src,sr,cfg.sourceKey),sourceBook.styles);
       if(!k||srcCounts.get(k)>1||!targetKeys.has(k)){
         summary.unmatched++;fileStats[sourceIndex].unmatched++;
         const reason=!k?'no_key':srcCounts.get(k)>1?'dup_source_key':ambiguousDst.has(k)?'dup_target_key':'missing_target_key';
         skipDetails.push(skipRecord(reason,'unmatched',{file:sourceName,sheet:cfg.sources[sourceIndex].sourceSheet,source:'Dòng '+sr,sourceValue:k||''}));continue;}
       tr=targetKeys.get(k);
     }
     if(tr>1048576)throw Error('Vượt quá giới hạn số dòng của Excel.');
     summary.matched++;fileStats[sourceIndex].matched++;
     const targetAddresses=new Set();
     for(const [i,m] of cfg.mappings.entries()){
       const srcLoc=readLocation(src,sr,m.source);
       const raw=cellValue(src,sourceBook,sr,m.source),rec=reconcile[i];
       if(!raw||raw.raw===''||raw.kind==='e'){
         summary.skippedBlank++;
         skipDetails.push(skipRecord('source_blank','cell',{file:sourceName,sheet:cfg.sources[sourceIndex].sourceSheet,source:addr(srcLoc.row,srcLoc.col),dest:addr(tr,m.target)}));continue;}
       const loc=readLocation(dst,tr,m.target),destination=addr(loc.row,loc.col);
       if(targetAddresses.has(destination))throw Error('Hai cột bạn chọn đang trỏ vào cùng một ô gộp '+destination+' tại dòng '+tr+'. Hãy bỏ một cặp cột bị trùng.');
       targetAddresses.add(destination);
       const rule=formatFor(cfg.formatRules,tr,m.target)||formatFor(cfg.formatRules,loc.row,loc.col),reference=addr(srcLoc.row,srcLoc.col)+' → '+destination;
       const val=normalizeImport(raw,rule,cellAt(src,srcLoc.row,srcLoc.col),reference,sourceBook);
       const oldCell=cellAt(dst,loc.row,loc.col),oldVal=cellValue(dst,dstBook,tr,m.target);
       const merged=loc.row!==tr,formula=!!oldCell&&!!first(oldCell,'f');
       const exists=cfg.writeMode==='empty'&&oldVal&&oldVal.raw!=='';
       const same=equalValues(oldVal,val),needsFormat=!!rule&&rule.type!=='auto';
       let doWrite=false;
       let reason=null;
       if(merged){summary.skippedMerge++;reason='dest_merged';}
       else if(formula){summary.skippedFormula++;reason='dest_formula';}
       else if(exists){summary.skippedExists++;reason='dest_exists';}
       else if(same&&!needsFormat){summary.skippedSame++;reason='same_value';}
       if(reason)skipDetails.push(skipRecord(reason,'cell',{file:sourceName,sheet:cfg.sources[sourceIndex].sourceSheet,source:addr(srcLoc.row,srcLoc.col),dest:destination,sourceValue:displayImport(val,rule),targetValue:oldVal?formatValue(oldVal):''}));
       else {
         doWrite=true;
         changes.push({src:(sourceIndex+1)+". "+sourceName+" · "+addr(srcLoc.row,srcLoc.col),sourceFile:sourceName,sourceIndex,srcMerge:mergeRef(srcLoc.merge),dst:destination,dstMerge:mergeRef(loc.merge),row:loc.row,_origRow:loc.row,col:loc.col,old:oldVal?formatValue(oldVal):'',new:displayImport(val,rule),value:val,formatCode:rule?.code||null,formatType:rule?.type||'auto'});
         rec.changed++;
       }
       const after=doWrite?val:oldVal,sNum=numericValue(val,rule);
       if(sNum!==null){
         rec.sourceSum+=sNum;rec.numericRows++;
         const before=numericValue(oldVal,rule),afterNum=numericValue(after,rule);
         if(before!==null)rec.beforeSum+=before;
         if(afterNum!==null)rec.afterSum+=afterNum;else rec.missingAfter++;
         if(before===null||Math.abs(before-sNum)>0.000001)rec.differentBefore++;
         if(afterNum===null||Math.abs(afterNum-sNum)>0.000001){
           rec.differentAfter++;
           differences.push({sourceFile:sourceName,sourceRow:sr,targetRow:tr,source:colLetter(srcLoc.col),target:colLetter(loc.col),sourceValue:formatValue(val),before:before===null?'':formatValue(oldVal),after:afterNum===null?'':formatValue(after),reason:formula?'Ô đích có công thức':merged?'Ô gộp nhiều dòng':exists?'Ô đích có dữ liệu':'Khác số'});
         }
       }
     }
   }
   for(const rec of reconcile){rec.sourceSum=Number(rec.sourceSum.toFixed(6));rec.beforeSum=Number(rec.beforeSum.toFixed(6));rec.afterSum=Number(rec.afterSum.toFixed(6));rec.diffBefore=Number((rec.sourceSum-rec.beforeSum).toFixed(6));rec.diffAfter=Number((rec.sourceSum-rec.afterSum).toFixed(6));}
   return {settings:cfg,changes,summary,reconcile,differences,fileStats,skipDetails};
  }
 function setCell(sheet,row,col,val,formatCode,styles){
  let rowNode=sheet.rows.get(row);
  if(!rowNode){rowNode=sheet.doc.createElementNS(NS,'row');rowNode.setAttribute('r',String(row));
   const later=Array.from(sheet.rows.entries()).filter(([r])=>r>row).sort((a,b)=>a[0]-b[0])[0];
   if(later)sheet.sheetData.insertBefore(rowNode,later[1]);else sheet.sheetData.appendChild(rowNode);
   sheet.rows.set(row,rowNode);sheet.maxRow=Math.max(sheet.maxRow,row);
  }
  let cell=cellAt(sheet,row,col);
  if(!cell){cell=sheet.doc.createElementNS(NS,'c');cell.setAttribute('r',addr(row,col));
   let later=null;for(const item of xmlChildren(rowNode,'c')){if(colFromAddress(item.getAttribute('r')||'')>col){later=item;break;}}
   if(later)rowNode.insertBefore(cell,later);else rowNode.appendChild(cell);sheet.cells.set(addr(row,col),cell);
  }
  for(const child of Array.from(cell.children)){if(['v','is'].includes(child.localName))cell.removeChild(child);}
  if(val.kind==='s'||val.kind==='d'){
   cell.setAttribute('t','inlineStr');const ins=sheet.doc.createElementNS(NS,'is'),t=sheet.doc.createElementNS(NS,'t');
   if(/^\s|\s$|\n/.test(val.raw))t.setAttribute('xml:space','preserve');
   t.textContent=val.raw;ins.appendChild(t);cell.appendChild(ins);
  }else{cell.setAttribute('t',val.kind==='b'?'b':'n');const v=sheet.doc.createElementNS(NS,'v');v.textContent=val.raw;cell.appendChild(v);}
  sheet.maxCol=Math.max(sheet.maxCol,col);
  if(formatCode)styles.assign(cell,formatCode);
 }
 function fixDimension(sheet,changes){const dim=first(sheet.doc.documentElement,'dimension');if(!dim)return;const old=dim.getAttribute('ref')||'A1';const parts=old.split(':');const start=parts[0]||'A1',end=parts[1]||start;let minC=colFromAddress(start)||1,minR=rowFromAddress(start)||1;for(const change of changes){minC=Math.min(minC,change.col);minR=Math.min(minR,change.row);}const maxC=Math.max(sheet.maxCol,colFromAddress(end)||1),maxR=Math.max(sheet.maxRow,rowFromAddress(end)||1);dim.setAttribute('ref',addr(minR,minC)+':'+addr(maxR,maxC));}
 function markRecalc(book){
  const root=book.wbDoc.documentElement;let c=first(root,'calcPr');if(!c){c=book.wbDoc.createElementNS(NS,'calcPr');root.appendChild(c);}c.setAttribute('fullCalcOnLoad','1');c.setAttribute('forceFullCalc','1');c.setAttribute('calcMode','auto');
 }
 function csvCell(text){let s=String(text??'');if(/^[\s]*[=+@-]/.test(s))s="'"+s;return '"'+s.replaceAll('"','""')+'"';}
 function csvLog(entries){return '\uFEFF'+[['Nguồn','Đích','Giá trị trước','Giá trị sau'],...entries.map(x=>[x.src,x.dst,x.old,x.new])].map(row=>row.map(csvCell).join(';')).join('\r\n');}

 function fmtReconcile(v){return niceNum(Number(v.toFixed(4)));}
 function reconciliationTable(recs){
  return '<table class="data-table"><thead><tr><th>Cột nguồn → đích</th><th class="num">Dòng số</th><th class="num">Tổng nguồn</th><th class="num">Đích trước</th><th class="num">Dự kiến sau</th><th class="num">Lệch sau</th><th class="num">Dòng lệch</th></tr></thead><tbody>'+recs.map(x=>'<tr><td><strong>'+escapeHtml(x.source+' → '+x.target)+'</strong></td><td class="num">'+niceNum(x.numericRows)+'</td><td class="num">'+fmtReconcile(x.sourceSum)+'</td><td class="num">'+fmtReconcile(x.beforeSum)+'</td><td class="num">'+fmtReconcile(x.afterSum)+'</td><td class="num '+(Math.abs(x.diffAfter)>0.00001?'recon-warn':'recon-ok')+'">'+fmtReconcile(x.diffAfter)+'</td><td class="num">'+niceNum(x.differentAfter)+'</td></tr>').join('')+'</tbody></table>';
 }
 function showReconciliation(result){
  const r=result.reconcile,total=result.differences.length;
  el('reconRows').textContent=niceNum(result.summary.matched);el('reconMismatches').textContent=niceNum(total);el('reconPairs').textContent=niceNum(r.length);el('reconBadge').textContent=total?'Cần kiểm tra chênh lệch':'Đã đối chiếu';el('reconBadge').classList.toggle('warn',!!total);
  el('reconStatus').textContent='Đã kiểm tra '+result.fileStats.length+' file nguồn. Ghép '+niceNum(result.summary.matched)+' dòng; '+niceNum(result.summary.skippedTotalSource)+' dòng tổng cộng nguồn đã bỏ qua. Có '+niceNum(result.changes.length)+' ô sẽ nhập.'+(result.summary.unmatched?'\n[Lưu ý] '+niceNum(result.summary.unmatched)+' dòng nguồn chưa ghép được với đích.':'')+(total?'\n[Lưu ý] '+niceNum(total)+' giá trị số vẫn chênh lệch sau khi nhập theo các quy tắc hiện tại.':'\n✓ Không phát hiện chênh lệch số trên các ô đã ghép và có giá trị số.');
  el('reconStatus').className='status '+(total||result.summary.unmatched?'warn':'ok');
  el('reconTable').innerHTML=reconciliationTable(r);
  const diffs=result.differences.slice(0,150);
  el('reconDifferences').innerHTML=diffs.length?'<table class="data-table"><thead><tr><th>File / Ô nguồn</th><th>Ô đích</th><th>Số nguồn</th><th>Đích trước</th><th>Dự kiến sau</th><th>Lý do</th></tr></thead><tbody>'+diffs.map(d=>'<tr><td>'+escapeHtml((d.sourceFile?d.sourceFile+' · ':'')+d.source+d.sourceRow)+'</td><td>'+escapeHtml(d.target+d.targetRow)+'</td><td>'+escapeHtml(d.sourceValue)+'</td><td>'+escapeHtml(d.before)+'</td><td>'+escapeHtml(d.after)+'</td><td>'+escapeHtml(d.reason)+'</td></tr>').join('')+'</tbody></table>':'<div class="empty">Không có dòng số chênh lệch trong phần đã ghép.</div>';
  el('reconDownload').disabled=false;
 }
 function reconcileCSV(result){const lines=[['ĐỐI CHIẾU THEO CỘT'],['Cột nguồn','Cột đích','Số dòng số','Tổng nguồn','Đích trước','Dự kiến sau','Chênh trước','Chênh sau','Dòng lệch sau']];
  for(const r of result.reconcile)lines.push([r.source,r.target,r.numericRows,r.sourceSum,r.beforeSum,r.afterSum,r.diffBefore,r.diffAfter,r.differentAfter]);
  lines.push([],['CHI TIẾT CHÊNH LỆCH'],['Ô nguồn','Ô đích','Nguồn','Trước nhập','Dự kiến sau','Nguyên nhân']);
  for(const d of result.differences)lines.push([(d.sourceFile?d.sourceFile+' · ':'')+d.source+d.sourceRow,d.target+d.targetRow,d.sourceValue,d.before,d.after,d.reason]);
  return '\uFEFF'+lines.map(row=>row.map(csvCell).join(';')).join('\r\n');
 }

 function renderSkipDetailTable(details,limit=200){
   if(!details.length)return '<div class="empty">Không có mục nào thuộc nhóm đang chọn.</div>';
   return '<table><thead><tr><th style="width:12%">Loại</th><th style="width:20%">File / Sheet</th><th style="width:12%">Ô nguồn</th><th style="width:11%">Ô đích</th><th style="width:20%">Nguyên nhân</th><th style="width:25%">Cách kiểm tra</th></tr></thead><tbody>'+details.slice(0,limit).map(x=>'<tr><td><span class="skip-pill '+(x.unit==='unmatched'?'attention':'')+'">'+(x.unit==='unmatched'?'Chưa ghép':x.unit==='cell'?'Ô':'Dòng')+'</span></td><td><b>'+escapeHtml(x.file||'—')+'</b><span class="note">'+escapeHtml(x.sheet||'')+'</span></td><td>'+escapeHtml(x.source||'—')+(x.sourceValue!==''?'<span class="note">'+escapeHtml(x.sourceValue)+'</span>':'')+'</td><td>'+escapeHtml(x.dest||'—')+(x.targetValue!==''?'<span class="note">Trước: '+escapeHtml(x.targetValue)+'</span>':'')+'</td><td class="why '+(x.unit==='unmatched'?'important':'')+'">'+escapeHtml(x.reason)+'</td><td>'+escapeHtml(x.advice)+'</td></tr>').join('')+'</tbody></table>';
 }
 function refreshSkipDetails(){
   if(!plan)return;
   const all=plan.skipDetails||[],kind=el('skipReason').value,query=el('skipSearch').value.toLocaleLowerCase('vi');
   const rows=all.filter(x=>(kind==='all'||x.category===kind)&&(!query||[x.file,x.sheet,x.source,x.dest,x.reason,x.sourceValue,x.targetValue].join(' ').toLocaleLowerCase('vi').includes(query)));
   el('skipTable').innerHTML=renderSkipDetailTable(rows);
   el('skipFoot').textContent='Đang hiển thị '+niceNum(Math.min(rows.length,200))+' / '+niceNum(rows.length)+' mục phù hợp. Bấm “Tải CSV” để lưu đầy đủ '+niceNum(all.length)+' mục.';
 }
 function showSkipDetails(result){
   const all=result.skipDetails||[];
   const cells=all.filter(x=>x.unit==='cell').length,rows=all.filter(x=>x.unit==='row').length,unmatched=all.filter(x=>x.unit==='unmatched').length;
   el('statSkipped').textContent=niceNum(cells+rows);
   el('showSkipped').disabled=false;el('downloadSkips').disabled=!all.length;
   el('showSkipped').textContent='⌕ Xem chi tiết ('+niceNum(all.length)+')';
   el('skipQuickText').textContent='Bỏ qua '+niceNum(cells)+' ô và '+niceNum(rows)+' dòng.'+(unmatched?' Có '+niceNum(unmatched)+' giao dịch chưa ghép được; cần xử lý trước khi xuất.':' Không có giao dịch chưa ghép.');
   el('skipBreakdown').innerHTML='<div class="skip-stat"><b>'+niceNum(cells)+'</b><small>Ô bỏ qua (trống, có công thức, đã có dữ liệu…)</small></div><div class="skip-stat"><b>'+niceNum(rows)+'</b><small>Dòng bỏ qua (tổng cộng, dòng rỗng…)</small></div><div class="skip-stat '+(unmatched?'attention':'')+'"><b>'+niceNum(unmatched)+'</b><small>Giao dịch chưa tìm được dòng đích</small></div>';
   const categories=[...new Set(all.map(x=>x.category))];
   const select=el('skipReason');select.innerHTML='<option value="all">Tất cả nguyên nhân ('+niceNum(all.length)+')</option>'+categories.map(cat=>'<option value="'+cat+'">'+escapeHtml(SKIP_REASONS[cat]?.[0]||cat)+' ('+niceNum(all.filter(x=>x.category===cat).length)+')</option>').join('');
   el('skipSearch').value='';refreshSkipDetails();
 }
 function showPreview(result){window.dispatchEvent(new CustomEvent('excelflow:preview',{detail:result}));
  el('statChanges').textContent=niceNum(result.changes.length);el('statRows').textContent=niceNum(result.summary.matched);
  showSkipDetails(result);el('previewLabel').textContent='Đã phân tích';
  const s=result.summary;message('✓ Đã đối chiếu '+result.fileStats.length+' file nguồn, '+niceNum(s.examined)+' dòng; '+niceNum(s.matched)+' dòng khớp. Dự kiến nhập '+niceNum(result.changes.length)+' ô.'+(s.unmatched?'\n[Lưu ý] '+s.unmatched+(result.settings.matchMode==='key'?' dòng không tìm thấy khóa hợp lệ.':' dòng không tìm thấy vị trí đích phù hợp.'):'')+((s.dupSrc||s.dupDst)?'\n[Lưu ý] Khóa trùng: nguồn '+s.dupSrc+', đích '+s.dupDst+'. Các khóa trùng được bỏ qua.':'')+(s.skippedFormula?'\n✓ Đã bảo vệ '+s.skippedFormula+' ô có công thức.':'')+(s.skippedTotalSource||s.skippedTotalTarget?'\n✓ Đã bỏ '+s.skippedTotalSource+' dòng tổng cộng nguồn và '+s.skippedTotalTarget+' dòng tổng cộng đích.':'')+(s.skippedFormulaRows?'\n[Lưu ý] Không tìm thấy đủ dòng đích phù hợp: '+s.skippedFormulaRows+' giao dịch chưa ghép.':'')+(s.autoBlockedByTotal?'\n[Lưu ý] Bảng có dòng TỔNG CỘNG ở cuối; ứng dụng không tự ghi xuống dưới dòng tổng. Hãy tạo dòng trống phía trên dòng tổng.':'')+(s.unmatched?'\n[Lưu ý] Chưa thể xuất vì vẫn còn giao dịch chưa ghép. Hãy xử lý các dòng trên trước.':''),s.unmatched?'warn':'ok');
  const list=result.changes.slice(0,150);
  el('noPreview').style.display=list.length?'none':'block';el('previewTable').style.display=list.length?'block':'none';
  if(list.length){const html='<table class="data-table"><thead><tr><th>File / Ô nguồn</th><th>Ô đích</th><th>Trước khi nhập</th><th>Sau khi nhập</th></tr></thead><tbody>'+list.map(ch=>'<tr><td>'+escapeHtml(ch.src)+(ch.srcMerge?'<small style="display:block;color:#7690a7">Ô gộp '+escapeHtml(ch.srcMerge)+'</small>':'')+'</td><td><strong>'+escapeHtml(ch.dst)+'</strong>'+(ch.dstMerge?'<small style="display:block;color:#7690a7">Ô gộp '+escapeHtml(ch.dstMerge)+'</small>':'')+'</td><td>'+escapeHtml(ch.old||'(trống)')+'</td><td class="diff">'+escapeHtml(ch.new)+'</td></tr>').join('')+'</tbody></table>';el('previewTable').innerHTML=html;}
  const fileResults=el('fileResults');fileResults.style.display='block';
  fileResults.innerHTML='<table><thead><tr><th>File nguồn theo thứ tự nhập</th><th>Vùng dòng</th><th>Đã xét</th><th>Đã ghép</th><th>Chưa ghép</th></tr></thead><tbody>'+result.fileStats.map((x,i)=>'<tr><td><b>'+(i+1)+'. '+escapeHtml(x.name)+'</b><small style="display:block;color:#8192a5">'+escapeHtml(x.sheet)+'</small></td><td>'+escapeHtml(x.range)+'</td><td>'+niceNum(x.read)+'</td><td>'+niceNum(x.matched)+'</td><td>'+niceNum(x.unmatched)+'</td></tr>').join('')+'</tbody></table>';
  el('previewFoot').textContent='Hiển thị tối đa 150 ô đầu tiên. Bản Excel xuất sẽ bao gồm tất cả '+niceNum(result.changes.length)+' ô; nhật ký đầy đủ có thể tải trong Lịch sử.';
 }
 async function preview(){if(loading)return;try{loading=true;el('previewBtn').disabled=true;message('Đang kiểm tra dữ liệu…');await sleep();const cfg=await extractSettings();plan=analyze(cfg);planSettings=JSON.stringify(collectUI());showPreview(plan);showReconciliation(plan);el('exportBtn').disabled=plan.changes.length===0||plan.summary.unmatched>0;if(plan.summary.unmatched)toast('Còn '+plan.summary.unmatched+' giao dịch chưa có dòng đích. Hãy chuẩn bị đủ dòng trống hoặc kiểm tra mã trước khi xuất.');else if(!plan.changes.length)toast('Không phát hiện ô nào cần nhập.');}catch(e){plan=null;el('exportBtn').disabled=true;message('Lỗi: '+e.message,'err');toast(e.message);}finally{loading=false;el('previewBtn').disabled=false;}}
 function collectUI(){const ids=['autoSourceRange','headerRow','targetHeaderRow','sourceStart','sourceEnd','targetStart','targetMode','sourceSheet','targetSheet','mapping','matchMode','writeMode','sourceKey','targetKey','skipBlank','skipEmpty','reverse','skipTotals','skipZeros','totalTerms','rowFilter','formatRules'];return {fields:ids.map(id=>{const node=el(id);return [id,node.type==='checkbox'?node.checked:node.value];}),sources:srcEntries.map(x=>[x.id,x.sheet,x.file.name,x.file.size,x.file.lastModified])};}
 async function exportResult(){
  if(loading)return;
  if(!plan||!plan.changes.length){toast('Hãy xem trước dữ liệu trước khi xuất.');return;}
  if(plan.summary.unmatched){toast('Không thể xuất vì còn '+plan.summary.unmatched+' giao dịch chưa ghép. Hãy sửa file đích hoặc quy tắc và xem trước lại.');return;}
  if(planSettings!==JSON.stringify(collectUI())){invalidate();toast('Quy tắc đã đổi. Hãy xem trước lại.');return;}
  const approved=await window.efAsk('File đích: '+dstBook.file.name+'\nSố ô sẽ ghi: '+plan.changes.length.toLocaleString('vi-VN')+'\nSố giao dịch chưa ghép: '+plan.summary.unmatched+'\n\nỨng dụng sẽ tạo một file Excel mới. File gốc không bị thay đổi. Bạn có đồng ý xuất file không?', 'Xác nhận xuất Excel');
  if(!approved)return;
  if(loading||planSettings!==JSON.stringify(collectUI())){toast('Thiết lập đã thay đổi. Hãy xem trước lại.');return;}
  try{
   loading=true;el('exportBtn').disabled=true;message('Đang ghi dữ liệu vào bản sao Excel…');await sleep();
   const cfg=plan.settings,clone=await JSZip.loadAsync(await dstBook.file.arrayBuffer());
   const dstXml=xmlParse(await clone.file(cfg.dst.path).async('string'));
   const sheet={...cfg.dst,doc:dstXml,sheetData:all(dstXml.documentElement,'sheetData')[0],rows:new Map(),cells:new Map(),maxRow:cfg.dst.maxRow,maxCol:cfg.dst.maxCol};
   for(const row of xmlChildren(sheet.sheetData,'row')){const rn=Number(row.getAttribute('r'));if(rn)sheet.rows.set(rn,row);for(const c of xmlChildren(row,'c'))sheet.cells.set(c.getAttribute('r')?.toUpperCase(),c);}
   const styleManager=getStyleManager(clone);await styleManager.prepare();
   for(const change of plan.changes)setCell(sheet,change.row,change.col,change.value,change.formatCode,styleManager);
   await styleManager.save();
   fixDimension(sheet,plan.changes);
   clone.file(cfg.dst.path,new XMLSerializer().serializeToString(sheet.doc));
   const wbDoc=xmlParse(await clone.file('xl/workbook.xml').async('string'));
   const cp=first(wbDoc.documentElement,'calcPr');if(cp){cp.setAttribute('fullCalcOnLoad','1');cp.setAttribute('forceFullCalc','1');cp.setAttribute('calcMode','auto');}
   else {const newCalc=wbDoc.createElementNS(NS,'calcPr');newCalc.setAttribute('fullCalcOnLoad','1');newCalc.setAttribute('forceFullCalc','1');newCalc.setAttribute('calcMode','auto');wbDoc.documentElement.appendChild(newCalc);}
   clone.file('xl/workbook.xml',new XMLSerializer().serializeToString(wbDoc));
   const outputBlob=await clone.generateAsync({type:'blob',compression:'DEFLATE',compressionOptions:{level:4}});
   const outName=dstBook.file.name.replace(/\.(xlsx|xlsm)$/i,'')+'_DA_NHAP.'+(dstBook.file.name.match(/\.(xlsx|xlsm)$/i)||['','.xlsx'])[1];
   download(outputBlob,outName);
   let saved=false;try{
    let sourceBlob=srcEntries[0].file,sourceName=srcEntries[0].file.name;
    if(srcEntries.length>1){const archive=new JSZip();for(const [index,entry] of srcEntries.entries())archive.file('nguon/'+String(index+1).padStart(2,'0')+'_'+entry.file.name,entry.file);archive.file('DANH_SACH_FILE_NGUON.txt',srcEntries.map((e,i)=>(i+1)+'. '+e.file.name+' | Sheet: '+e.sheet).join('\n'));sourceBlob=await archive.generateAsync({type:'blob',compression:'DEFLATE'});sourceName='EXCELFLOW_DANH_SACH_FILE_NGUON.zip';}
    await insertRun({id:Date.now()+'-'+Math.random().toString(36).slice(2,8),time:Date.now(),sourceName,targetName:dstBook.file.name,sourceFiles:srcEntries.map(x=>({name:x.file.name,sheet:x.sheet})),targetNameSingle:dstBook.file.name,outputName:outName,sourceSheet:cfg.sourceSheet,targetSheet:cfg.targetSheet,mapping:el('mapping').value,matched:plan.summary.matched,count:plan.changes.length,summary:plan.summary,
     importOptions:collectUI(),reconciliation:plan.reconcile,reconDifferences:plan.differences,skipDetails:plan.skipDetails,skipCount:plan.skipDetails.length,
     changes:plan.changes.map(({src,dst,old:newOld,new:newNew})=>({src,dst,old:newOld,new:newNew})),sourceBlob,beforeBlob:dstBook.file,afterBlob:outputBlob,logBlob:new Blob([csvLog(plan.changes)],{type:'text/csv;charset=utf-8'}),skipBlob:new Blob([skipCsv(plan.skipDetails)],{type:'text/csv;charset=utf-8'})});saved=true;
   }catch(err){console.error(err);toast('Đã tải Excel nhưng không lưu được lịch sử: '+err.message);}
   message('✓ Đã xuất '+outName+' với '+niceNum(plan.changes.length)+' ô được nhập.'+(saved?'\n✓ Đã lưu cả dữ liệu trước và sau trong Lịch sử.':'\n[Lưu ý] Chưa lưu được lịch sử, hãy kiểm tra dung lượng trình duyệt.'),'ok');
   toast('Đã tạo file Excel'+(saved?' và lưu lịch sử':'')+'.');
  }catch(e){message('Không thể xuất file: '+e.message,'err');toast('Lỗi xuất: '+e.message);}
  finally{loading=false;el('exportBtn').disabled=!plan||!plan.changes.length||plan.summary.unmatched>0;}
 }
 // IndexedDB: an offline archive of source, original target, output, and the full changes log.
 const DB='ExcelFlow_2Files_LichSu',STORE='runs';let dbOpenPromise;
 function openDb(){if(!('indexedDB' in window))return Promise.reject(Error('Trình duyệt không hỗ trợ lưu lịch sử.'));
  if(!dbOpenPromise)dbOpenPromise=new Promise((resolve,reject)=>{const req=indexedDB.open(DB,1);req.onupgradeneeded=()=>{if(!req.result.objectStoreNames.contains(STORE))req.result.createObjectStore(STORE,{keyPath:'id'});};req.onsuccess=()=>resolve(req.result);req.onerror=()=>reject(req.error);});return dbOpenPromise;
 }
 async function dbRequest(method,arg){const db=await openDb();return new Promise((resolve,reject)=>{const tx=db.transaction(STORE,method==='getAll'||method==='get'?'readonly':'readwrite');let q;const s=tx.objectStore(STORE);if(method==='getAll')q=s.getAll();if(method==='get')q=s.get(arg);if(method==='put')q=s.put(arg);if(method==='delete')q=s.delete(arg);if(method==='clear')q=s.clear();q.onsuccess=()=>resolve(q.result);q.onerror=()=>reject(q.error);});}
 const insertRun=entry=>dbRequest('put',entry);
 let histories=[],activeHistory=null,historyView='changes',historyPage=0,loadedHistorySheet=null,historyLoading=false;
 async function showHistories(preferredId){
  try{histories=(await dbRequest('getAll')).sort((a,b)=>b.time-a.time);el('historyCountPill').textContent=histories.length+' lần nhập đã lưu';
   const list=el('historyList');list.replaceChildren();if(!histories.length){list.innerHTML='<div class="empty">Chưa có lịch sử. Hãy nhập và xuất một file trước.</div>';activeHistory=null;el('historyDetail').style.display='none';el('historyEmpty').style.display='block';return;}
   for(const h of histories){const b=document.createElement('button');b.className='history-item'+(h.id===(preferredId||activeHistory?.id)?' active':'');const when=new Date(h.time).toLocaleString('vi-VN');b.innerHTML='<strong>'+escapeHtml(h.outputName)+'</strong><small>◷ '+escapeHtml(when)+' • '+niceNum(h.count)+' ô được nhập</small><small>'+escapeHtml(h.sourceSheet)+' → '+escapeHtml(h.targetSheet)+'</small>';b.onclick=()=>chooseHistory(h.id);list.appendChild(b);}
   await chooseHistory(preferredId&&histories.some(h=>h.id===preferredId)?preferredId:(activeHistory&&histories.some(h=>h.id===activeHistory.id)?activeHistory.id:histories[0].id));
  }catch(e){el('historyList').innerHTML='<div class="status err">Không đọc được lịch sử: '+escapeHtml(e.message)+'</div>';}
 }
 async function chooseHistory(id){const h=histories.find(x=>x.id===id);if(!h)return;activeHistory=h;loadedHistorySheet=null;historyPage=0;historyView='changes';
  el('historyDetail').style.display='block';el('historyEmpty').style.display='none';el('historyTitle').textContent=h.outputName;el('historyDesc').textContent=new Date(h.time).toLocaleString('vi-VN')+' • Nguồn: '+(h.sourceFiles?.length?h.sourceFiles.length+' file ('+h.sourceFiles.map(x=>x.name).join(', ')+')':h.sourceName)+' • Đích: '+h.targetName;
  el('historyMeta').textContent='Đã nhập '+niceNum(h.count)+' ô trên '+niceNum(h.matched)+' dòng. Quy tắc đã sử dụng:\n'+h.mapping+(h.importOptions?'\nTùy chọn: '+h.importOptions.map(x=>x[0]+'='+x[1]).join(' ; '):'');
  document.querySelectorAll('[data-view]').forEach(b=>b.classList.toggle('active',b.dataset.view==='changes'));el('historyList').querySelectorAll('.history-item').forEach((b,i)=>b.classList.toggle('active',histories[i].id===id));
  await renderHistoryView();
 }
 async function renderHistoryView(){
  const h=activeHistory;if(!h)return;const body=el('historyView');body.innerHTML='<div class="empty">Đang tải dữ liệu...</div>';
  if(historyView==='reconcile'){
   const rec=h.reconciliation||[];body.innerHTML=rec.length?reconciliationTable(rec):'<div class="empty">Lần nhập này chưa có bảng đối chiếu.</div>';
   el('pageInfo').textContent=rec.length+' cặp cột';el('prevPage').disabled=true;el('nextPage').disabled=true;return;
  }
  if(historyView==='skips'){
   const rows=h.skipDetails||[],pages=Math.max(1,Math.ceil(rows.length/50));historyPage=Math.max(0,Math.min(historyPage,pages-1));
   body.innerHTML=rows.length?'<div class="skip-table">'+renderSkipDetailTable(rows.slice(historyPage*50,(historyPage+1)*50),50)+'</div>':'<div class="empty">Không có dữ liệu bỏ qua được lưu ở lần nhập này (hoặc đây là lịch sử từ phiên bản trước 3.5).</div>';
   el('pageInfo').textContent='Trang '+(historyPage+1)+'/'+pages+' • '+niceNum(rows.length)+' mục';el('prevPage').disabled=historyPage===0;el('nextPage').disabled=historyPage>=pages-1;return;
  }
  if(historyView==='changes'){
   const rows=h.changes||[],pages=Math.max(1,Math.ceil(rows.length/50));historyPage=Math.max(0,Math.min(historyPage,pages-1));
   const items=rows.slice(historyPage*50,(historyPage+1)*50);
   body.innerHTML='<table class="data-table"><thead><tr><th>File / Ô nguồn</th><th>Ô đích</th><th>Giá trị cũ</th><th>Đã nhập</th></tr></thead><tbody>'+items.map(c=>'<tr><td>'+escapeHtml(c.src)+'</td><td>'+escapeHtml(c.dst)+'</td><td>'+escapeHtml(c.old||'(trống)')+'</td><td class="diff">'+escapeHtml(c.new)+'</td></tr>').join('')+'</tbody></table>';
   el('pageInfo').textContent='Trang '+(historyPage+1)+'/'+pages+' • '+niceNum(rows.length)+' ô';el('prevPage').disabled=historyPage===0;el('nextPage').disabled=historyPage>=pages-1;return;
  }
  try{
   if(!loadedHistorySheet||loadedHistorySheet.type!==historyView){const blob=historyView==='before'?h.beforeBlob:h.afterBlob,book=await openWorkbook(new File([blob],h.targetName));const sheet=await sheetOf(book,h.targetSheet);loadedHistorySheet={type:historyView,book,sheet};}
   const {book,sheet}=loadedHistorySheet;const rows=Array.from(sheet.rows.keys()).sort((a,b)=>a-b);const pages=Math.max(1,Math.ceil(rows.length/50));historyPage=Math.max(0,Math.min(historyPage,pages-1));
   const shownCols=Math.min(Math.max(sheet.maxCol,5),26),head='<thead><tr><th>Dòng</th>'+Array.from({length:shownCols},(_,c)=>'<th>'+colLetter(c+1)+'</th>').join('')+'</tr></thead>';
   const trows=rows.slice(historyPage*50,(historyPage+1)*50);const rowsHtml=trows.map(r=>'<tr><th>'+r+'</th>'+Array.from({length:shownCols},(_,c)=>{const v=cellValue(sheet,book,r,c+1);return '<td title="'+escapeHtml(v?formatValue(v):'')+'">'+escapeHtml(v?formatValue(v):'')+'</td>';}).join('')+'</tr>').join('');
   body.innerHTML='<table class="data-table">'+head+'<tbody>'+rowsHtml+'</tbody></table>';el('pageInfo').textContent='Trang '+(historyPage+1)+'/'+pages+' • '+niceNum(rows.length)+' dòng • Hiện tối đa 26 cột';el('prevPage').disabled=historyPage===0;el('nextPage').disabled=historyPage>=pages-1;
  }catch(e){body.innerHTML='<div class="status err">Không đọc được bảng lưu cũ: '+escapeHtml(e.message)+'</div>';}
 }
 async function backupHistory(){try{if(!histories.length){toast('Chưa có lịch sử để sao lưu.');return;}const zip=new JSZip();const metadata=[];
  for(const h of histories){const base='lan-nhap/'+h.id+'/';const {sourceBlob,beforeBlob,afterBlob,logBlob,skipBlob,...meta}=h;metadata.push(meta);zip.file(base+(h.sourceName?.endsWith('.zip')?'source.zip':'source.xlsx'),sourceBlob);zip.file(base+'before.xlsx',beforeBlob);zip.file(base+'after.xlsx',afterBlob);zip.file(base+'changes.csv',logBlob);if(skipBlob)zip.file(base+'skipped.csv',skipBlob);}
  zip.file('metadata.json',JSON.stringify({app:'ExcelFlow',version:2,records:metadata}));const blob=await zip.generateAsync({type:'blob',compression:'STORE'});download(blob,'SAO_LUU_LICH_SU_EXCEL_'+new Date().toISOString().slice(0,10)+'.zip');toast('Đã xuất bản sao lưu lịch sử.');
 }catch(e){toast('Không thể sao lưu: '+e.message);}}
 async function restoreBackup(file){if(!file)return;if(!await window.efAsk('Khôi phục dữ liệu lịch sử từ file '+file.name+'? Các bản ghi trong gói sao lưu sẽ được thêm vào lịch sử trình duyệt.','Xác nhận khôi phục lịch sử')){el('restoreFile').value='';return;}try{const zip=await JSZip.loadAsync(await file.arrayBuffer());const metadata=JSON.parse(await zip.file('metadata.json')?.async('string')||'{}');if(metadata.app!=='ExcelFlow'||!Array.isArray(metadata.records))throw Error('Không đúng file sao lưu lịch sử của ứng dụng.');let count=0;
  for(const h of metadata.records){const base='lan-nhap/'+h.id+'/';const source=zip.file(base+'source.zip')||zip.file(base+'source.xlsx'),before=zip.file(base+'before.xlsx'),after=zip.file(base+'after.xlsx'),log=zip.file(base+'changes.csv');if(!source||!before||!after||!log)continue;
   const skipFile=zip.file(base+'skipped.csv');const bytes=await Promise.all([source.async('blob'),before.async('blob'),after.async('blob'),log.async('blob'),skipFile?skipFile.async('blob'):Promise.resolve(null)]);await insertRun({...h,id:Date.now()+'-restore-'+count+'-'+Math.random().toString(36).slice(2,6),sourceBlob:bytes[0],beforeBlob:bytes[1],afterBlob:bytes[2],logBlob:bytes[3],skipBlob:bytes[4]});count++;}
  toast('Đã khôi phục '+count+' bản lịch sử.');await showHistories();
 }catch(e){toast('Không thể khôi phục: '+e.message);}finally{el('restoreFile').value='';}}

  // Bank statement smart detection, merged headers, and human-friendly mappings.
  const recognizedBankWord=t=>deaccent(t).replace(/\s/g,'');
  function detectStatement(sheet,book){
    const maxScan=Math.min(65,sheet.maxRow),maxC=Math.min(35,Math.max(sheet.maxCol,12));
    let candidate=null;
    for(let r=1;r<=maxScan;r++){
      let hits=0,serialCol=null;const labels=[];
      for(let c=1;c<=maxC;c++){
        const loc=readLocation(sheet,r,c);if(loc.col!==c)continue;
        const v=cellValue(sheet,book,r,c);if(!v||v.kind!=='s')continue;
        const txt=deaccent(v.raw);if(!txt||txt.length>80)continue;
        labels.push({col:c,label:txt});
        if(/^(stt|so thu tu|no\.?|tt)$/.test(txt))serialCol=c;
        if(/ngay giao dich|ngay hach toan|phat sinh no|phat sinh co|noi dung|so du|transaction date|debit|credit|balance/.test(txt))hits++;
      }
      const score=hits+(serialCol?2:0);
      if(score>=4&&(!candidate||score>candidate.score))candidate={header:r,serialCol,score,labels};
    }
    if(!candidate)return null;
    let start=candidate.header+1,end=0,count=0;
    if(candidate.serialCol){
      let began=false;
      for(let r=candidate.header+1;r<=sheet.maxRow;r++){
        const v=cellValue(sheet,book,r,candidate.serialCol),txt=v?String(v.raw).trim():'';
        if(/^\d{1,6}$/.test(txt)){if(!began)start=r;end=r;count++;began=true;}
        else if(began)break;
        else if(r>candidate.header+5)break;
      }
    }
    // When there is no numbered STT sequence, stop before first total row.
    if(!count){
      start=candidate.header+1;
      const terms=totalTerms(el('totalTerms').value);
      end=sheet.maxRow;
      for(let r=start;r<=sheet.maxRow;r++){if(isTotalRow(sheet,book,r,terms)){end=r-1;break;}}
      count=Math.max(0,end-start+1);
    }
    return {...candidate,start,end,count};
  }
  function shortCell(sheet,book,row,col,max=64){
    const v=cellValue(sheet,book,row,col);if(!v)return '';
    const txt=formatValue(v);return txt.length>max?txt.slice(0,max-1)+'…':txt;
  }
  async function showSourceSummary({autoSet=true}={}){
    if(!srcBook)return;
    try{
      const sheet=await sheetOf(srcBook,el('sourceSheet').value);
      const suggestion=detectStatement(sheet,srcBook);
      if(autoSet){
        if(suggestion){el('headerRow').value=suggestion.header;el('sourceStart').value=suggestion.start;el('sourceEnd').value=suggestion.end;el('reverse').checked=true;el('skipTotals').checked=true;}
        else {el('headerRow').value=1;el('sourceStart').value=2;el('sourceEnd').value='';}
      }
      const start=Number(el('sourceStart').value)||2,end=Number(el('sourceEnd').value)||sheet.maxRow;
      const mergeCount=sheet.merges.length;
      el('sourceInsight').className='smart-info'+(suggestion?'':' neutral');
      if(suggestion){
        el('sourceInsight').innerHTML='<strong>✓ Đã nhận diện cấu trúc sao kê ngân hàng</strong><div>Tiêu đề: dòng <b>'+suggestion.header+'</b> · Giao dịch: từ dòng <b>'+suggestion.start+'</b> đến <b>'+suggestion.end+'</b> ('+suggestion.count+' giao dịch) · <b>'+mergeCount+'</b> vùng ô gộp.</div><div>Bạn có thể bật/tắt nhập ngược ở bước 3 để chọn thứ tự thời gian phù hợp.</div>';
      }else{
        el('sourceInsight').innerHTML='<strong>✓ Đã đọc sheet nguồn</strong><div>Tìm thấy '+sheet.maxRow+' dòng và '+mergeCount+' vùng ô gộp. Nếu dòng đầu/dòng cuối chưa đúng, mở “Chỉnh vùng dữ liệu”.</div>';
      }
      await showSourcePeek(sheet,srcBook,suggestion,start,end);
      await refreshCols();
    }catch(e){el('sourceInsight').textContent='Không xem được dữ liệu nguồn: '+e.message;}
  }
  async function showSourcePeek(sheet,book,suggestion,start,end){
    const box=el('sourcePeek');
    if(!sheet||!sheet.rows.size){box.style.display='none';return;}
    let columns=[];const header=Number(el('headerRow').value)||1;
    if(suggestion){
      const keys=[/^(stt|so thu tu)$/, /ngay giao dich/,/phat sinh no/,/phat sinh co/,/noi dung/,/but toan/];
      for(const k of keys){const f=suggestion.labels.find(x=>k.test(x.label));if(f)columns.push(f.col);}
    }
    if(columns.length<3){columns=getColumns(sheet,book,header,15).filter(c=>c.header).slice(0,6).map(c=>c.num);}
    columns=[...new Set(columns)].slice(0,6);
    if(!columns.length){box.style.display='none';return;}
    const rows=Array.from(sheet.rows.keys()).filter(r=>r>=start&&r<=end&&columns.some(c=>{const v=cellValue(sheet,book,r,c);return v&&v.raw!=='';})).sort((a,b)=>a-b).slice(0,4);
    if(!rows.length){box.style.display='none';return;}
    const th=columns.map(c=>{const m=mergeAt(sheet,header,c);return '<th>'+escapeHtml(m&&m.right>m.left?colLetter(m.left)+'–'+colLetter(m.right):colLetter(c))+'<span class="subcell">'+escapeHtml(shortCell(sheet,book,header,c,30))+'</span></th>';}).join('');
    const body=rows.map(r=>'<tr>'+columns.map(c=>{
       const m=mergeAt(sheet,r,c),loc=readLocation(sheet,r,c);
       return '<td>'+escapeHtml(shortCell(sheet,book,r,c,85))+(m?'<span class="subcell">'+escapeHtml(mergeRef(m))+'</span>':'')+'</td>';
    }).join('')+'</tr>').join('');
    box.innerHTML='<div style="padding:8px 11px;font-size:11px;color:#62809a;background:#f9fbfd"><b>Xem thử 4 dòng đầu</b> · chỉ xem, chưa thay đổi file</div><table><thead><tr>'+th+'</tr></thead><tbody>'+body+'</tbody></table>';box.style.display='block';
  }
  function readMappingPairs(){
    return el('mapping').value.split('\n').map(x=>x.trim()).filter(x=>x&&!x.startsWith('#')).map(line=>{
      const m=line.match(/^\s*(.*?)\s*(?:->|=>|→|:|=)\s*(.*?)\s*$/);
      return m?{s:m[1],t:m[2],raw:line}:{raw:line,error:true};
    });
  }
  function renderMappingList(){
    const box=el('mappingList');if(!box)return;
    const pairs=readMappingPairs();if(!pairs.length){box.innerHTML='<div class="mapping-empty">Chưa có cột nào được ghép. Chọn cột nguồn và cột đích ở phía trên.</div>';return;}
    const src=srcBook?.sheetCache.get(el('sourceSheet').value),dst=dstBook?.sheetCache.get(el('targetSheet').value);
    const sc=src?getColumns(src,srcBook,Number(el('headerRow').value)||1):[],tc=dst?getColumns(dst,dstBook,Number(el('targetHeaderRow').value)||1):[];
    box.replaceChildren();
    pairs.forEach((p,i)=>{
      let a=p.s,b=p.t,detail='';
      if(!p.error){
        try{const ca=mappingColumn(p.s,sc,'nguồn'),cb=mappingColumn(p.t,tc,'đích');
          const sa=sc.find(x=>x.num===ca),ta=tc.find(x=>x.num===cb);
          a=sa?.display||p.s;b=ta?.display||p.t;
          const fr=el('formatRules').value.split('\n').filter(Boolean).findLast(x=>x.trim().toUpperCase().startsWith(colLetter(cb)+' ->'));
          if(fr)detail=fr.split('->')[1]?.trim()||'';
        }catch{}
      }
      const row=document.createElement('div');row.className='mapping-row';
      row.innerHTML='<div class="mapping-name" title="'+escapeHtml(a)+'">'+escapeHtml(a)+'</div><div class="mapping-arrow">→</div><div class="mapping-name" title="'+escapeHtml(b)+'">'+escapeHtml(b)+(detail?'<small>Kiểu: '+escapeHtml(detail)+'</small>':'')+'</div><button class="remove-map" type="button" title="Bỏ cặp cột" aria-label="Xóa cặp cột '+(i+1)+'">×</button>';
      row.querySelector('button').onclick=()=>{
       const current=readMappingPairs();const removed=current.splice(i,1)[0];
       el('mapping').value=current.map(x=>x.raw).join('\n');
       if(removed&&removed.t){try{
        const target=mappingColumn(removed.t,tc,'đích');const col=colLetter(target);
        el('formatRules').value=el('formatRules').value.split('\n').filter(line=>!new RegExp('^\\s*'+col+'\\s*(?:->|=>|→|=)').test(line)).join('\n');
       }catch{}}
       invalidate();renderMappingList();
      };
      box.appendChild(row);
    });
  }

 // Events and offline initialization.
 function goPage(which){
   const history=which==='history',compare=which==='compare',home=which==='home',pdf2excel=which==='pdf2excel';
   const isImport=!history&&!compare&&!home&&!pdf2excel;
   el('homePage')?.classList.toggle('active',home);
   el('importPage')?.classList.toggle('active',isImport);
   el('comparePage')?.classList.toggle('active',compare);
   el('historyPage')?.classList.toggle('active',history);
   el('pdf2ExcelPage')?.classList.toggle('active',pdf2excel);

   // Ẩn thanh header hoàn toàn khi ở Trang chủ
   if(el('efHeader')) el('efHeader').style.display = home ? 'none' : 'flex';

   el('navHome')?.classList.toggle('active',home);
   el('navHomeTab')?.classList.toggle('active',home);
   el('navImport')?.classList.toggle('active',isImport);
   el('navPdf2Excel')?.classList.toggle('active',pdf2excel);
   el('navCompare')?.classList.toggle('active',compare);
   el('navHistory')?.classList.toggle('active',history);

   if(el('efSaveTemplate')) el('efSaveTemplate').style.display = isImport ? '' : 'none';
   if(el('exportBtn')) el('exportBtn').style.display = isImport ? '' : 'none';

   if(history)showHistories();
   if(compare&&plan)showReconciliation(plan);
   window.scrollTo({top:0,behavior:'smooth'});
 }
 window.efGoPage=goPage;
 el('navHome')?.addEventListener('click',()=>goPage('home'));
 el('navHomeTab')?.addEventListener('click',()=>goPage('home'));
 el('navImport')?.addEventListener('click',()=>goPage('import'));
 el('navPdf2Excel')?.addEventListener('click',()=>goPage('pdf2excel'));
 el('navCompare')?.addEventListener('click',()=>goPage('compare'));
 el('navHistory')?.addEventListener('click',()=>goPage('history'));
 document.querySelectorAll('[data-nav-home]').forEach(btn=>btn.addEventListener('click',()=>goPage('home')));
 document.querySelectorAll('[data-nav-import]').forEach(btn=>btn.addEventListener('click',()=>goPage('import')));
 document.querySelectorAll('[data-nav-pdf2excel]').forEach(btn=>btn.addEventListener('click',()=>goPage('pdf2excel')));
 // Source files are additive; only the destination remains a single file.
 const fileVersions={source:0,target:0};
 function fileIdentity(file){return [file.name,file.size,file.lastModified].join('::');}
 function refreshSourceUI(){
  const zone=el('sourceDrop'),label=el('srcName'),list=el('sourceList');
  el('sourceFileCount').textContent=srcEntries.length+' file nguồn';el('clearSources').disabled=!srcEntries.length;
  label.textContent=srcEntries.length?'✓ Đã chọn '+srcEntries.length+' file nguồn · Có thể thêm file':'Chưa chọn file nguồn';label.classList.toggle('ready',!!srcEntries.length);
  list.replaceChildren();if(!srcEntries.length){list.innerHTML='<div class="source-list-empty">Chưa có file nguồn. Kéo nhiều file Excel vào khung phía trên.</div>';return;}
  for(const [i,entry] of srcEntries.entries()){
   const row=document.createElement('div');row.className='source-entry';
   const numBox=document.createElement('div');numBox.className='source-number';numBox.textContent=String(i+1);
   const info=document.createElement('div');info.className='source-file-info';
   const label=document.createElement('strong');label.title=entry.file.name;label.textContent=entry.file.name;
   const detail=document.createElement('small');detail.textContent=(entry.file.size/1024).toFixed(0)+' KB · '+entry.book.sheets.length+' sheet';info.append(label,detail);
   const sel=document.createElement('select');sel.setAttribute('aria-label','Sheet của file '+entry.file.name);
   for(const sh of entry.book.sheets){const opt=document.createElement('option');opt.value=sh.name;opt.textContent=sh.name;sel.appendChild(opt);}sel.value=entry.sheet;
   sel.onchange=async()=>{entry.sheet=sel.value;invalidate();if(i===0){el('sourceSheet').value=entry.sheet;await showSourceSummary({autoSet:true});}else{refreshSourceUI();toast('Đã đổi sheet của '+entry.file.name);}};
   const actions=document.createElement('div');actions.className='source-item-actions';
   const makeBtn=(icon,title,fn,disabled=false,css='')=>{const b=document.createElement('button');b.type='button';b.className=css;b.title=title;b.setAttribute('aria-label',title+' '+entry.file.name);b.textContent=icon;b.disabled=disabled;b.onclick=fn;actions.appendChild(b);};
   makeBtn('↑','Đưa file lên trước',()=>moveSource(i,-1),i===0);
   makeBtn('↓','Đưa file xuống sau',()=>moveSource(i,1),i===srcEntries.length-1);
   makeBtn('×','Bỏ file khỏi danh sách',()=>removeSource(i),false,'remove-source');
   row.append(numBox,info,sel,actions);list.appendChild(row);
  }
 }
 async function afterSourceListChange(updateTemplate=true){
  srcBook=srcEntries[0]?.book||null;
  if(srcBook){
    selectSheets('sourceSheet',srcBook);
    const isPdf = srcEntries.some(e => e.file && /\.pdf$/i.test(e.file.name));
    if (isPdf && dstBook) {
      const dstName = el('targetSheet')?.value || '';
      if (/bkmv|mua\s*v[aà]o/i.test(dstName)) {
        const mv = srcBook.sheets.find(s => /mua\s*v[aà]o|01-2/i.test(s.name));
        if (mv) el('sourceSheet').value = mv.name;
      } else if (/bkbr|b[aá]n\s*ra/i.test(dstName)) {
        const br = srcBook.sheets.find(s => /b[aá]n\s*ra|01-1/i.test(s.name));
        if (br) el('sourceSheet').value = br.name;
      }
    }
    el('sourceSheet').value=srcEntries[0].sheet;
    await autoDetectAndSetSourceHeader();
  }
  else {
    el('sourceSheet').replaceChildren(new Option('Chọn file nguồn trước',''));
    el('sourceInsight').textContent='Sau khi chọn các sao kê, ứng dụng sẽ tự tìm các giao dịch và ô gộp của từng file.';
    el('sourcePeek').style.display='none';
    if(el('sourceHeaderTag')) el('sourceHeaderTag').style.display='none';
  }
  refreshSourceUI();invalidate();await refreshCols();
  if(srcBook)await showSourceSummary({autoSet:updateTemplate&&el('autoSourceRange').checked});
 }
 async function moveSource(i,delta){const ni=i+delta;if(ni<0||ni>=srcEntries.length)return;const [item]=srcEntries.splice(i,1);srcEntries.splice(ni,0,item);await afterSourceListChange(true);toast('Đã đổi thứ tự. Hãy xem trước lại để kiểm tra.');}
 async function removeSource(i){if(!await window.efAsk('Bỏ file “'+srcEntries[i].file.name+'” khỏi phiên làm việc? File trên máy vẫn được giữ nguyên.','Xác nhận gỡ file nguồn',true))return;srcEntries.splice(i,1);await afterSourceListChange(true);toast('Đã bỏ file khỏi danh sách nguồn.');}
 async function addSourceFiles(files){
  const zone=el('sourceDrop');zone.classList.add('is-loading');zone.classList.remove('is-dragging');
  let added=0,ignored=0,errors=[];const seen=new Set(srcEntries.map(e=>fileIdentity(e.file)));
  try{
   for(const file of files){
    if(!/\.(xlsx|xlsm|pdf)$/i.test(file.name)){errors.push(file.name+': chỉ nhận .xlsx, .xlsm hoặc .pdf');continue;}
    if(seen.has(fileIdentity(file))){ignored++;continue;}
    try{
     message('Đang đọc file '+file.name+' ('+(added+1)+'/'+files.length+')…');await sleep();const book=await openWorkbook(file);
     const sourceSheet=srcEntries[0]?.sheet||el('sourceSheet').value;
     const preferred=book.sheets.some(s=>s.name===sourceSheet)?sourceSheet:book.sheets[0].name;
     await sheetOf(book,preferred);
     srcEntries.push({id:++sourceCounter,file,book,sheet:preferred});seen.add(fileIdentity(file));added++;
    }catch(e){errors.push(file.name+': '+e.message);}
   }
   if(added){await afterSourceListChange(srcEntries.length===added);}
   else refreshSourceUI();
   message((added?'✓ Đã thêm '+added+' file nguồn. Hiện có '+srcEntries.length+' file. ':'')+(ignored?'Bỏ qua '+ignored+' file đã chọn trước đó. ':'')+(errors.length?'\n[Lưu ý] '+errors.join('\n'):'')+(dstBook?'\nBạn có thể ghép cột và xem trước.':'\nHãy chọn thêm một file đích.'),errors.length?'warn':'ok');
  }finally{zone.classList.remove('is-loading');}
 }
 function queueSourceFiles(files){sourceQueue=sourceQueue.then(()=>addSourceFiles(files)).catch(e=>{message('Không đọc được file nguồn: '+e.message,'err');toast(e.message);});return sourceQueue;}
 window.efAddSourceFiles = queueSourceFiles;
 async function loadExcelFile(file,kind){
  if(kind==='source')return queueSourceFiles([file]);
  if(!file)return;
  const zone=el('targetDrop'),label=el('dstName');
  if(!/\.(xlsx|xlsm)$/i.test(file.name)){message('Chỉ hỗ trợ Excel .xlsx/.xlsm','err');return;}
  const version=++fileVersions.target;zone.classList.add('is-loading');zone.classList.remove('is-dragging');
  try{
   message('Đang đọc '+file.name+'…');await sleep();const book=await openWorkbook(file);
   if(version!==fileVersions.target)return;
   dstBook=book;selectSheets('targetSheet',book);
   const srcSheetName = el('sourceSheet')?.value || '';
   if (/mua\s*v[aà]o|bkmv|01-2/i.test(srcSheetName)) {
     const matchSheet = book.sheets.find(s => /bkmv|mua\s*v[aà]o/i.test(s.name));
     if (matchSheet) el('targetSheet').value = matchSheet.name;
   } else if (/b[aá]n\s*ra|bkbr|01-1/i.test(srcSheetName)) {
     const matchSheet = book.sheets.find(s => /bkbr|b[aá]n\s*ra/i.test(s.name));
     if (matchSheet) el('targetSheet').value = matchSheet.name;
   }
   await autoDetectAndSetTargetHeader();
   invalidate();label.textContent='✓ '+file.name;label.classList.add('ready');
   await refreshCols();message('✓ Đã đọc file đích: '+file.name+'. '+(srcEntries.length?'Có '+srcEntries.length+' file nguồn đã sẵn sàng.':'Hãy thêm file nguồn.'),'ok');
  }catch(e){if(version===fileVersions.target){message('Không đọc được file đích: '+e.message,'err');toast(e.message);}}
  finally{if(version===fileVersions.target)zone.classList.remove('is-loading');}
 }
 const isFileDrag=e=>Array.from(e.dataTransfer?.types||[]).includes('Files');
 for(const eventName of ['dragover','drop'])document.addEventListener(eventName,e=>{if(isFileDrag(e))e.preventDefault();});
 for(const [field,kind,zoneId] of [['sourceFile','source','sourceDrop'],['targetFile','target','targetDrop']]){
  const picker=el(field),zone=el(zoneId);
  picker.addEventListener('change',()=>{
   const files=Array.from(picker.files||[]);
   if(kind==='source'){picker.value='';queueSourceFiles(files);}else loadExcelFile(files[0],kind);
  });
  zone.addEventListener('click',e=>{if(e.target!==picker&&!picker.contains(e.target))picker.click();});
  zone.addEventListener('keydown',e=>{if(e.target===zone&&(e.key==='Enter'||e.key===' ')){e.preventDefault();picker.click();}});
  zone.addEventListener('dragenter',e=>{if(!isFileDrag(e))return;e.preventDefault();zone.classList.add('is-dragging');});
  zone.addEventListener('dragover',e=>{if(!isFileDrag(e))return;e.preventDefault();e.dataTransfer.dropEffect='copy';zone.classList.add('is-dragging');});
  zone.addEventListener('dragleave',e=>{if(e.relatedTarget&&!zone.contains(e.relatedTarget))zone.classList.remove('is-dragging');});
  zone.addEventListener('drop',e=>{
   e.preventDefault();e.stopPropagation();zone.classList.remove('is-dragging');
   const files=Array.from(e.dataTransfer?.files||[]);
   if(!files.length)return;
   if(kind==='source')queueSourceFiles(files);
   else if(files.length===1)loadExcelFile(files[0],kind);
   else toast('File đích chỉ được chọn 1 file. Hãy thả một file đích.');
  });
 }
 el('clearSources').onclick=async()=>{if(!srcEntries.length)return;if(!await window.efAsk('Bạn muốn gỡ '+srcEntries.length+' file nguồn khỏi danh sách đang làm việc? File gốc trên máy sẽ không bị xóa.','Xác nhận xóa danh sách nguồn',true))return;srcEntries=[];await afterSourceListChange(true);toast('Đã xóa danh sách nguồn. File đích không thay đổi.');};
 document.addEventListener('dragend',()=>document.querySelectorAll('.dropzone').forEach(z=>z.classList.remove('is-dragging')));
  el('sourceSheet').addEventListener('change',async()=>{
    if(srcEntries[0])srcEntries[0].sheet=el('sourceSheet').value;
    if(dstBook){
      const srcName=el('sourceSheet').value;
      if(/mua\s*v[aà]o|bkmv|01-2/i.test(srcName)){
        const match=dstBook.sheets.find(s=>/bkmv|mua\s*v[aà]o/i.test(s.name));
        if(match&&el('targetSheet').value!==match.name){el('targetSheet').value=match.name;await autoDetectAndSetTargetHeader();}
      }else if(/b[aá]n\s*ra|bkbr|01-1/i.test(srcName)){
        const match=dstBook.sheets.find(s=>/bkbr|b[aá]n\s*ra/i.test(s.name));
        if(match&&el('targetSheet').value!==match.name){el('targetSheet').value=match.name;await autoDetectAndSetTargetHeader();}
      }
    }
    await autoDetectAndSetSourceHeader();
    refreshSourceUI();invalidate();await showSourceSummary({autoSet:true});
  });
  el('targetSheet').addEventListener('change',async()=>{
    await autoDetectAndSetTargetHeader();
    invalidate();
    await refreshCols();
  });
  ['headerRow','targetHeaderRow'].forEach(id=>{
    const input = el(id);
    if (input) {
      input.addEventListener('change',()=>{invalidate();refreshCols();if(id==='headerRow'&&srcBook)showSourceSummary({autoSet:false});});
    }
  });
  ['autoSourceRange','sourceStart','sourceEnd','targetStart','targetMode','mapping','matchMode','writeMode','sourceKey','targetKey','skipBlank','skipEmpty','reverse','skipTotals','skipZeros','totalTerms','rowFilter','formatRules'].forEach(id=>el(id).addEventListener('change',invalidate));
  el('mapping').addEventListener('input',()=>{invalidate();renderMappingList();});el('formatRules').addEventListener('input',()=>{invalidate();renderMappingList();});el('totalTerms').addEventListener('input',invalidate);
  el('sourceStart').addEventListener('change',()=>{el('autoSourceRange').checked=false;invalidate();srcBook&&showSourceSummary({autoSet:false});});
  el('sourceEnd').addEventListener('change',()=>{el('autoSourceRange').checked=false;invalidate();srcBook&&showSourceSummary({autoSet:false});});
  el('matchMode').addEventListener('change',()=>{el('keys').style.display=el('matchMode').value==='key'?'grid':'none';});
  el('targetMode').addEventListener('change',()=>{el('manualTargetField').style.display=el('targetMode').value==='manual'?'block':'none';});
  el('detectRange').onclick=()=>{if(!srcBook){toast('Hãy chọn file nguồn trước.');return;}el('autoSourceRange').checked=true;showSourceSummary({autoSet:true});invalidate();toast('Sẽ tự nhận diện vùng giao dịch và ô gộp riêng cho từng file.');};
  el('addPair').onclick=()=>{
    const s=el('addSource').value,t=el('addTarget').value,type=el('addType').value;
    if(!s||!t||!srcBook||!dstBook){toast('Hãy chọn đủ 2 file, cột nguồn và cột đích.');return;}
    const src=srcBook.sheetCache.get(el('sourceSheet').value),dst=dstBook.sheetCache.get(el('targetSheet').value);
    const sCol=mergeAt(src,Number(el('headerRow').value),Number(s))?.left||Number(s);
    const dCol=mergeAt(dst,Number(el('targetHeaderRow').value),Number(t))?.left||Number(t);
    const rule=colLetter(sCol)+' -> '+colLetter(dCol);
    const existing=readMappingPairs();if(existing.some(p=>p.raw.toUpperCase()===rule.toUpperCase())){toast('Cặp cột này đã được thêm.');return;}
    el('mapping').value+=(el('mapping').value.trim()?'\n':'')+rule;
    if(type!=='auto'){
      const spec=colLetter(dCol)+' -> '+type;
      const existingFormat=el('formatRules').value.split('\n').filter(x=>x.trim()&&!new RegExp('^\\s*'+colLetter(dCol)+'\\s*(?:->|=>|→|=)').test(x));
      existingFormat.push(spec);el('formatRules').value=existingFormat.join('\n');
    }
    invalidate();renderMappingList();toast('Đã thêm: '+rule+(type!=='auto'?' · '+type:''));
  };
  el('bankPreset').onclick=async()=>{if(!await window.efAsk('Áp dụng cấu hình sao kê ngân hàng? Thao tác này thay đổi các lựa chọn nhập ngược, bỏ dòng tổng cộng và tự tìm dòng đích.','Áp dụng thiết lập sao kê'))return;el('reverse').checked=true;el('skipTotals').checked=true;el('skipEmpty').checked=true;el('matchMode').value='position';el('keys').style.display='none';el('targetMode').value='auto';el('manualTargetField').style.display='none';invalidate();toast('Đã bật nhập ngược, bỏ tổng cộng và tự tìm dòng đích.');};
 el('fmtType').onchange=()=>{el('customFmtField').style.display=el('fmtType').value==='custom'?'block':'none';};
 el('addFormat').onclick=()=>{const where=el('fmtWhere').value.trim(),type=el('fmtType').value;if(!where){toast('Hãy ghi cột, dòng hoặc vùng đích, ví dụ D hoặc F5:F30.');return;}
  const fmt=type==='custom'?'custom:'+el('fmtCustom').value.trim():type;const line=where+' -> '+fmt;try{parseFormats(line);}catch(e){toast(e.message);return;}
  const target=el('formatRules');target.value+=(target.value.trim()?'\n':'')+line;invalidate();renderMappingList();toast('Đã thêm quy tắc '+line);};
 el('clearFormat').onclick=()=>{el('formatRules').value='';invalidate();renderMappingList();toast('Đã xóa quy tắc định dạng, file đích sẽ giữ kiểu dữ liệu gốc.');};
 el('goPreview').onclick=()=>{goPage('import');el('previewBtn').scrollIntoView({behavior:'smooth',block:'center'});};
 el('reconDownload').onclick=()=>{if(plan)download(new Blob([reconcileCSV(plan)],{type:'text/csv;charset=utf-8'}),'DOI_CHIEU_SO_LIEU_'+new Date().toISOString().slice(0,10)+'.csv');};
 el('previewBtn').onclick=preview;el('exportBtn').onclick=exportResult;
 el('showSkipped').onclick=()=>{const box=el('skipDetails');box.hidden=!box.hidden;if(!box.hidden)box.scrollIntoView({behavior:'smooth',block:'nearest'});};
 el('skipReason').onchange=refreshSkipDetails;el('skipSearch').oninput=refreshSkipDetails;
 el('downloadSkips').onclick=async()=>{if(!plan)return;if(!await window.efAsk('Xuất báo cáo CSV chi tiết các mục bỏ qua?','Xác nhận xuất báo cáo'))return;download(new Blob([skipCsv(plan.skipDetails)],{type:'text/csv;charset=utf-8'}),'EXCELFLOW_CHI_TIET_BO_QUA_'+new Date().toISOString().slice(0,10)+'.csv');};
 el('downloadSource').onclick=async()=>{if(!activeHistory)return;if(!await window.efAsk('Bạn muốn tải lại file nguồn từ lịch sử?','Xác nhận tải dữ liệu'))return;download(activeHistory.sourceBlob,activeHistory.sourceName);};
 el('downloadBefore').onclick=async()=>{if(!activeHistory)return;if(!await window.efAsk('Bạn muốn tải file trước khi nhập từ lịch sử?','Xác nhận tải dữ liệu'))return;download(activeHistory.beforeBlob,'TRUOC_NHAP_'+activeHistory.targetName);};
 el('downloadAfter').onclick=async()=>{if(!activeHistory)return;if(!await window.efAsk('Bạn muốn tải file sau khi nhập từ lịch sử?','Xác nhận tải dữ liệu'))return;download(activeHistory.afterBlob,activeHistory.outputName);};
 el('downloadLog').onclick=async()=>{if(!activeHistory)return;if(!await window.efAsk('Bạn muốn tải nhật ký nhập liệu từ lịch sử?','Xác nhận tải dữ liệu'))return;download(activeHistory.logBlob,'NHAT_KY_'+activeHistory.id+'.csv');};
 el('downloadSkippedHistory').onclick=async()=>{if(!activeHistory)return;if(!await window.efAsk('Bạn muốn tải báo cáo mục bỏ qua từ lịch sử?','Xác nhận tải dữ liệu'))return;download(activeHistory.skipBlob||new Blob([skipCsv(activeHistory.skipDetails||[])],{type:'text/csv;charset=utf-8'}),'CHI_TIET_BO_QUA_'+activeHistory.id+'.csv');};
 el('deleteHistory').onclick=async()=>{if(!activeHistory)return;if(!await window.efAsk('Xóa vĩnh viễn bản lịch sử đã chọn và các bản Excel được lưu trong lịch sử này? Thao tác không thể hoàn tác.','Xóa lịch sử',true))return;const id=activeHistory.id;try{await dbRequest('delete',id);activeHistory=null;await showHistories();toast('Đã xóa bản ghi đã chọn.');}catch(e){toast('Không xóa được: '+e.message);}};
 document.querySelectorAll('[data-view]').forEach(b=>b.onclick=async()=>{if(!activeHistory)return;historyView=b.dataset.view;historyPage=0;document.querySelectorAll('[data-view]').forEach(x=>x.classList.toggle('active',x===b));await renderHistoryView();});
 el('prevPage').onclick=()=>{if(historyPage>0){historyPage--;renderHistoryView();}};el('nextPage').onclick=()=>{historyPage++;renderHistoryView();};
 el('backupHistory').onclick=async()=>{if(!await window.efAsk('Tạo và tải file ZIP sao lưu toàn bộ lịch sử đang lưu trên trình duyệt này?','Xác nhận sao lưu lịch sử'))return;await backupHistory();};el('restoreHistory').onclick=()=>el('restoreFile').click();el('restoreFile').addEventListener('change',()=>restoreBackup(el('restoreFile').files?.[0]));
 openDb().catch(e=>{console.warn('Local history not available',e);el('importStatus').textContent='Trình duyệt không cho lưu lịch sử. Hãy sử dụng Microsoft Edge hoặc Chrome bản mới để xem lại dữ liệu cũ.';});
})();
