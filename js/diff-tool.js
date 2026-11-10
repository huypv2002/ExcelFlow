/**
 * Puddin Excel Flow — Module Đối chiếu Dữ liệu giữa 2 File Excel
 * Hỗ trợ: chọn cột / chọn dòng / đối chiếu cả 2 file, đọc ngược sao kê,
 * nhận diện dòng khớp hoàn toàn, ô lệch dữ liệu, dòng thiếu và dòng dư.
 * Chỉ đối chiếu và hiển thị sự khác biệt (không xuất Excel).
 */
(function() {
  'use strict';

  const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const NS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const PAGE_SIZE = 50;

  const state = {
    bookA: null,
    bookB: null,
    scope: 'all', // 'all' | 'cols' | 'rows'
    pairs: [],    // Array<{ colA: number, colB: number, label: string }> (1-based col index, 0 = none)
    results: null, // { rows: Array<DiffRow>, pairs, counts }
    filterStatus: 'all', // 'all' | 'matched' | 'changed' | 'onlyA' | 'onlyB'
    onlyDiff: false,
    searchText: '',
    currentPage: 1
  };

  const el = id => document.getElementById(id);

  function showToast(msg) {
    if (typeof window.efToast === 'function') {
      window.efToast(msg);
    } else {
      const t = el('toast');
      if (t) {
        t.textContent = msg;
        t.classList.add('show');
        clearTimeout(showToast.timer);
        showToast.timer = setTimeout(() => t.classList.remove('show'), 2800);
      }
    }
  }

  function escapeHtml(val) {
    return String(val ?? '').replace(/[&<>"']/g, c => ({
      '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
    }[c]));
  }

  function colLetter(num) {
    let s = '';
    for (let n = num; n > 0; n = Math.floor((n - 1) / 26)) {
      s = String.fromCharCode(65 + (n - 1) % 26) + s;
    }
    return s || 'A';
  }

  function colIndex(letters) {
    let n = 0;
    for (const ch of String(letters || '').toUpperCase()) {
      const code = ch.charCodeAt(0);
      if (code >= 65 && code <= 90) n = n * 26 + (code - 64);
    }
    return n;
  }

  function deaccent(str) {
    return String(str ?? '')
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .replace(/đ/g, 'd')
      .replace(/Đ/g, 'D')
      .replace(/\s+/g, ' ')
      .trim()
      .toLowerCase();
  }

  function normalizeHeaderText(str) {
    return deaccent(str);
  }

  // =========================================================================
  // 1. PARSE EXCEL (.xlsx / .xlsm) & CSV
  // =========================================================================

  function xmlParse(text) {
    const doc = new DOMParser().parseFromString(text, 'application/xml');
    const err = doc.getElementsByTagName('parsererror')[0];
    if (err) throw new Error('XML không hợp lệ trong file Excel.');
    return doc;
  }

  function xmlChildren(node, localName) {
    return Array.from(node?.children || []).filter(n => n.localName === localName);
  }

  function firstChild(node, localName) {
    return xmlChildren(node, localName)[0] || null;
  }

  function allByTag(node, localName) {
    return node ? Array.from(node.getElementsByTagNameNS(NS_MAIN, localName)) : [];
  }

  function normalizeRelPath(target) {
    const clean = String(target || '').replaceAll('\\', '/');
    const full = clean.startsWith('/') ? clean.slice(1) : (clean.startsWith('xl/') ? clean : 'xl/' + clean);
    const out = [];
    for (const part of full.split('/')) {
      if (part === '..') out.pop();
      else if (part && part !== '.') out.push(part);
    }
    return out.join('/');
  }

  function isDateFormatCode(numFmtId, formatCode) {
    // Built-in Excel date/datetime numFmtIds
    if ((numFmtId >= 14 && numFmtId <= 22) || (numFmtId >= 45 && numFmtId <= 47)) return true;
    if (!formatCode) return false;
    const cleaned = String(formatCode)
      .replace(/\[[^\]]*\]/g, '')
      .replace(/"[^"]*"/g, '')
      .toLowerCase();
    return /[ymd]/.test(cleaned) && !/[#0?]/.test(cleaned);
  }

  function formatExcelSerialDate(serial, date1904, formatCode) {
    const n = Number(serial);
    if (!Number.isFinite(n) || n < 0 || n > 2958465) return String(serial);
    const epochOffset = date1904 ? 24107 : 25569; // days between 1900/1904 and 1970-01-01
    const wholeDays = Math.floor(n);
    const fracDay = n - wholeDays;
    // Excel 1900 leap year bug adjustment (serial 60 = 1900-02-29)
    const adjDays = (!date1904 && wholeDays > 59) ? wholeDays : wholeDays;
    const utcMs = Math.round((adjDays - epochOffset) * 86400000);
    const d = new Date(utcMs);
    if (Number.isNaN(d.getTime())) return String(serial);
    const dd = String(d.getUTCDate()).padStart(2, '0');
    const mm = String(d.getUTCMonth() + 1).padStart(2, '0');
    const yyyy = d.getUTCFullYear();

    if (fracDay > 0.0001 || /h|s/i.test(formatCode || '')) {
      const totalSec = Math.round(fracDay * 86400);
      const hh = String(Math.floor(totalSec / 3600) % 24).padStart(2, '0');
      const mi = String(Math.floor((totalSec % 3600) / 60)).padStart(2, '0');
      const ss = String(totalSec % 60).padStart(2, '0');
      if (wholeDays === 0) return `${hh}:${mi}:${ss}`;
      if (ss !== '00') return `${dd}/${mm}/${yyyy} ${hh}:${mi}:${ss}`;
      return `${dd}/${mm}/${yyyy} ${hh}:${mi}`;
    }
    return `${dd}/${mm}/${yyyy}`;
  }

  function extractStyles(stylesXml) {
    if (!stylesXml) return [];
    const doc = xmlParse(stylesXml);
    const fmtMap = new Map();
    for (const nf of allByTag(doc.documentElement, 'numFmt')) {
      fmtMap.set(Number(nf.getAttribute('numFmtId')), nf.getAttribute('formatCode') || '');
    }
    const cellXfs = allByTag(doc.documentElement, 'cellXfs')[0];
    if (!cellXfs) return [];
    return xmlChildren(cellXfs, 'xf').map(xf => {
      const id = Number(xf.getAttribute('numFmtId') || 0);
      const code = fmtMap.get(id) || '';
      return { numFmtId: id, formatCode: code, isDate: isDateFormatCode(id, code) };
    });
  }

  function parseCsvGrid(text) {
    const clean = text.replace(/^\uFEFF/, '');
    const firstLine = clean.split(/\r?\n/)[0] || '';
    const commaCount = (firstLine.match(/,/g) || []).length;
    const semiCount = (firstLine.match(/;/g) || []).length;
    const tabCount = (firstLine.match(/\t/g) || []).length;
    const delim = tabCount > commaCount && tabCount > semiCount ? '\t' : (semiCount > commaCount ? ';' : ',');

    const rows = [];
    let row = [];
    let cur = '';
    let inQuotes = false;

    for (let i = 0; i < clean.length; i++) {
      const ch = clean[i];
      if (inQuotes) {
        if (ch === '"') {
          if (clean[i + 1] === '"') {
            cur += '"';
            i++;
          } else {
            inQuotes = false;
          }
        } else {
          cur += ch;
        }
      } else if (ch === '"') {
        inQuotes = true;
      } else if (ch === delim) {
        row.push(cur);
        cur = '';
      } else if (ch === '\r') {
        if (clean[i + 1] === '\n') i++;
        row.push(cur);
        rows.push(row);
        row = [];
        cur = '';
      } else if (ch === '\n') {
        row.push(cur);
        rows.push(row);
        row = [];
        cur = '';
      } else {
        cur += ch;
      }
    }
    if (cur.length || row.length) {
      row.push(cur);
      rows.push(row);
    }

    const rowMap = new Map();
    let maxCol = 0;
    rows.forEach((r, idx) => {
      const rNum = idx + 1;
      const colMap = new Map();
      r.forEach((val, cIdx) => {
        if (val !== '') {
          colMap.set(cIdx + 1, String(val));
          if (cIdx + 1 > maxCol) maxCol = cIdx + 1;
        }
      });
      rowMap.set(rNum, colMap);
    });

    return {
      fileName: '',
      sheets: [{
        name: 'CSV',
        maxRow: rows.length,
        maxCol,
        rows: rowMap,
        merges: []
      }]
    };
  }

  async function parseExcelBook(file) {
    if (!file) throw new Error('Vui lòng chọn file.');
    if (/\.csv$/i.test(file.name)) {
      const text = await file.text();
      const book = parseCsvGrid(text);
      book.fileName = file.name;
      return book;
    }
    if (!/\.(xlsx|xlsm)$/i.test(file.name)) {
      throw new Error(`File "${file.name}" không đúng định dạng (.xlsx, .xlsm hoặc .csv).`);
    }
    const JSZipLib = window.JSZip;
    if (!JSZipLib) throw new Error('Chưa tải được thư viện JSZip.');

    const zip = await JSZipLib.loadAsync(await file.arrayBuffer());
    const wbFile = zip.file('xl/workbook.xml');
    const relFile = zip.file('xl/_rels/workbook.xml.rels');
    if (!wbFile || !relFile) {
      throw new Error(`Không đọc được cấu trúc file "${file.name}" (có thể file bị hỏng hoặc đặt mật khẩu).`);
    }

    const [wbXml, relXml, sharedXml, stylesXml] = await Promise.all([
      wbFile.async('string'),
      relFile.async('string'),
      zip.file('xl/sharedStrings.xml')?.async('string') || Promise.resolve(null),
      zip.file('xl/styles.xml')?.async('string') || Promise.resolve(null)
    ]);

    const wbDoc = xmlParse(wbXml);
    const relDoc = xmlParse(relXml);
    const relMap = new Map();
    for (const r of Array.from(relDoc.documentElement.children)) {
      if (r.localName === 'Relationship' && !r.getAttribute('TargetMode')) {
        relMap.set(r.getAttribute('Id'), normalizeRelPath(r.getAttribute('Target') || ''));
      }
    }

    const shared = sharedXml
      ? allByTag(xmlParse(sharedXml).documentElement, 'si').map(si =>
          allByTag(si, 't').map(t => t.textContent || '').join('')
        )
      : [];

    const styles = extractStyles(stylesXml);
    const wp = firstChild(wbDoc.documentElement, 'workbookPr');
    const date1904 = ['1', 'true'].includes((wp?.getAttribute('date1904') || '').toLowerCase());

    const sheetNodes = allByTag(wbDoc.documentElement, 'sheet');
    const sheets = [];

    for (const shNode of sheetNodes) {
      const name = shNode.getAttribute('name') || `Sheet ${sheets.length + 1}`;
      const rId = shNode.getAttributeNS(NS_REL, 'id') || shNode.getAttribute('r:id');
      const path = relMap.get(rId);
      if (!path || !zip.file(path)) continue;

      const sXml = await zip.file(path).async('string');
      const sDoc = xmlParse(sXml);
      const sheetData = firstChild(sDoc.documentElement, 'sheetData');
      const rowMap = new Map();
      let maxRow = 0;
      let maxCol = 0;

      if (sheetData) {
        let seqRow = 0;
        for (const rNode of xmlChildren(sheetData, 'row')) {
          seqRow++;
          const rowNum = Number(rNode.getAttribute('r')) || seqRow;
          const colMap = new Map();
          let seqCol = 0;

          for (const cNode of xmlChildren(rNode, 'c')) {
            seqCol++;
            const ref = cNode.getAttribute('r') || '';
            const colNum = ref ? colIndex((ref.match(/^[A-Z]+/i) || [''])[0]) : seqCol;
            if (!colNum) continue;

            const typ = cNode.getAttribute('t') || 'n';
            let valStr = '';

            if (typ === 'inlineStr') {
              const isNode = firstChild(cNode, 'is');
              valStr = isNode ? allByTag(isNode, 't').map(t => t.textContent || '').join('') : '';
            } else {
              const vNode = firstChild(cNode, 'v');
              if (vNode && vNode.textContent !== '') {
                const raw = vNode.textContent;
                if (typ === 's') {
                  valStr = shared[Number(raw)] ?? '';
                } else if (typ === 'b') {
                  valStr = raw === '1' ? 'TRUE' : 'FALSE';
                } else if (typ === 'str' || typ === 'd' || typ === 'e') {
                  valStr = raw;
                } else {
                  const sIdx = Number(cNode.getAttribute('s') || 0);
                  const styleInfo = styles[sIdx];
                  if (styleInfo && styleInfo.isDate && Number.isFinite(Number(raw))) {
                    valStr = formatExcelSerialDate(raw, date1904, styleInfo.formatCode);
                  } else if (styleInfo && /^0{2,}$/.test(styleInfo.formatCode) && Number.isInteger(Number(raw))) {
                    valStr = String(Number(raw)).padStart(styleInfo.formatCode.length, '0');
                  } else if (Number.isFinite(Number(raw))) {
                    valStr = String(Number(raw));
                  } else {
                    valStr = raw;
                  }
                }
              }
            }

            if (valStr !== '') {
              colMap.set(colNum, valStr);
              if (colNum > maxCol) maxCol = colNum;
            }
          }

          if (colMap.size > 0) {
            rowMap.set(rowNum, colMap);
            if (rowNum > maxRow) maxRow = rowNum;
          } else if (rowNum > maxRow) {
            maxRow = rowNum;
          }
        }
      }

      const merges = [];
      for (const mNode of allByTag(sDoc.documentElement, 'mergeCell')) {
        const bounds = (mNode.getAttribute('ref') || '').split(':');
        if (bounds.length >= 2) {
          const r1 = Number((bounds[0].match(/\d+/) || ['0'])[0]);
          const r2 = Number((bounds[1].match(/\d+/) || ['0'])[0]);
          const c1 = colIndex((bounds[0].match(/^[A-Z]+/i) || [''])[0]);
          const c2 = colIndex((bounds[1].match(/^[A-Z]+/i) || [''])[0]);
          if (r1 && r2 && c1 && c2) merges.push({ top: r1, bottom: r2, left: c1, right: c2 });
        }
      }

      sheets.push({ name, maxRow, maxCol, rows: rowMap, merges });
    }

    if (!sheets.length) throw new Error(`File "${file.name}" không có sheet dữ liệu nào.`);
    return { fileName: file.name, sheets };
  }

  function getCellDisplay(sheet, rowNum, colNum) {
    const rMap = sheet.rows.get(rowNum);
    if (rMap && rMap.has(colNum)) return rMap.get(colNum);
    return '';
  }

  function getHeaderCellDisplay(sheet, rowNum, colNum) {
    const direct = getCellDisplay(sheet, rowNum, colNum);
    if (direct !== '') return direct;
    if (sheet.merges && sheet.merges.length) {
      for (const m of sheet.merges) {
        if (rowNum >= m.top && rowNum <= m.bottom && colNum >= m.left && colNum <= m.right) {
          return getCellDisplay(sheet, m.top, m.left);
        }
      }
    }
    return '';
  }

  function detectHeaderRow(sheet) {
    if (!sheet || !sheet.maxRow) return 1;
    const keywords = [
      'stt', 'ngay', 'date', 'so', 'ma', 'ten', 'khach', 'mst', 'noi dung', 'dien giai',
      'so tien', 'phat sinh', 'ghi no', 'ghi co', 'no', 'co', 'so du', 'tien', 'thue',
      'tong', 'hoa don', 'chung tu', 'tai khoan', 'ghi chu', 'mat hang', 'don gia', 'so luong'
    ];
    let bestRow = 1;
    let bestScore = -1;
    const limit = Math.min(sheet.maxRow, 30);

    for (let r = 1; r <= limit; r++) {
      const rMap = sheet.rows.get(r);
      if (!rMap || !rMap.size) continue;
      let kwHits = 0;
      let nonEmpty = 0;
      for (const val of rMap.values()) {
        const text = deaccent(val);
        if (!text) continue;
        nonEmpty++;
        if (keywords.some(kw => text.includes(kw))) kwHits++;
      }
      const score = kwHits * 5 + nonEmpty;
      if (nonEmpty >= 2 && score > bestScore) {
        bestScore = score;
        bestRow = r;
      }
    }
    return bestRow;
  }

  function getColumns(sheet, headerRow) {
    if (!sheet) return [];
    const maxCol = Math.max(sheet.maxCol || 0, 1);
    const cols = [];
    for (let c = 1; c <= maxCol; c++) {
      const letter = colLetter(c);
      const label = headerRow > 0 ? getHeaderCellDisplay(sheet, headerRow, c).trim().replace(/\s+/g, ' ') : '';
      cols.push({
        num: c,
        letter,
        label,
        display: label ? `${letter} · ${label.slice(0, 42)}` : `Cột ${letter}`
      });
    }
    return cols;
  }

  function getActiveSheet(side) {
    const book = side === 'A' ? state.bookA : state.bookB;
    if (!book) return null;
    const name = el(`diffSheet${side}`)?.value;
    return book.sheets.find(s => s.name === name) || book.sheets[0] || null;
  }

  function getHeaderRowNum(side) {
    const raw = parseInt(el(`diffHeader${side}`)?.value ?? '1', 10);
    return Number.isFinite(raw) && raw >= 0 ? raw : 0;
  }

  // =========================================================================
  // 2. CHỌN DÒNG, ĐỌC NGƯỢC & XEM TRƯỚC MINI
  // =========================================================================

  function parseRowPickExpression(expr, maxRow) {
    const clean = String(expr || '').trim();
    if (!clean) return null;
    const result = [];
    const seen = new Set();
    const tokens = clean.split(/[,;]+/).map(t => t.trim()).filter(Boolean);

    for (const tok of tokens) {
      const rangeMatch = tok.match(/^(\d+)\s*[-–:]\s*(\d+)$/);
      if (rangeMatch) {
        let a = parseInt(rangeMatch[1], 10);
        let b = parseInt(rangeMatch[2], 10);
        if (a > b) [a, b] = [b, a];
        a = Math.max(1, a);
        b = Math.min(maxRow, b);
        for (let r = a; r <= b; r++) {
          if (!seen.has(r)) { seen.add(r); result.push(r); }
        }
      } else if (/^\d+$/.test(tok)) {
        const r = parseInt(tok, 10);
        if (r >= 1 && r <= maxRow && !seen.has(r)) {
          seen.add(r);
          result.push(r);
        }
      } else {
        throw new Error(`Danh sách dòng "${tok}" không hợp lệ. Ví dụ đúng: 2-30 hoặc 5, 8, 12-15.`);
      }
    }
    return result;
  }

  function isTotalOrFooterRow(sheet, rowNum) {
    const rMap = sheet.rows.get(rowNum);
    if (!rMap || !rMap.size) return false;
    const terms = ['tong cong', 'tong so', 'tong tien', 'tong', 'cong', 'subtotal', 'grand total', 'printed on', 'ngay gio in'];
    for (const val of rMap.values()) {
      const txt = deaccent(val);
      if (!txt) continue;
      for (const term of terms) {
        if (txt === term || txt === term + ':' || (term.length >= 6 && txt.startsWith(term))) {
          return true;
        }
      }
    }
    return false;
  }

  function extractRowNumbers(side, opts = {}) {
    const sheet = getActiveSheet(side);
    if (!sheet) return [];
    const maxRow = Math.max(sheet.maxRow, 1);
    const pickExpr = el(`diffRowPick${side}`)?.value.trim() || '';
    const skipBlank = opts.skipBlank !== undefined ? opts.skipBlank : (el('diffSkipBlank')?.checked ?? true);
    const skipTotals = opts.skipTotals !== undefined ? opts.skipTotals : (el('diffSkipTotals')?.checked ?? false);
    const reverse = el(`diffReverse${side}`)?.checked ?? false;

    let rowNums = [];
    if (pickExpr) {
      rowNums = parseRowPickExpression(pickExpr, maxRow) || [];
    } else {
      const headerRow = getHeaderRowNum(side);
      const startInput = parseInt(el(`diffStart${side}`)?.value || '', 10);
      const endInput = parseInt(el(`diffEnd${side}`)?.value || '', 10);
      const start = Number.isFinite(startInput) && startInput >= 1 ? startInput : Math.min(maxRow, headerRow + 1);
      const end = Number.isFinite(endInput) && endInput >= 1 ? Math.min(maxRow, endInput) : maxRow;
      if (start > end) return [];
      for (let r = start; r <= end; r++) rowNums.push(r);
    }

    const filtered = [];
    for (const r of rowNums) {
      const rMap = sheet.rows.get(r);
      const isBlank = !rMap || Array.from(rMap.values()).every(v => String(v ?? '').trim() === '');
      if (skipBlank && isBlank) continue;
      if (skipTotals && isTotalOrFooterRow(sheet, r)) continue;
      filtered.push(r);
    }

    if (reverse) filtered.reverse();
    return filtered;
  }

  function renderFilePeek(side) {
    const box = el(`diffPeek${side}`);
    const sheet = getActiveSheet(side);
    if (!box || !sheet) {
      if (box) box.innerHTML = '';
      return;
    }

    let rowNums = [];
    try {
      rowNums = extractRowNumbers(side);
    } catch (e) {
      box.innerHTML = `<div class="diff-peek-err">${escapeHtml(e.message)}</div>`;
      return;
    }

    const isRev = el(`diffReverse${side}`)?.checked;
    const cols = getColumns(sheet, getHeaderRowNum(side)).slice(0, 4);

    if (!rowNums.length) {
      box.innerHTML = `<div class="diff-peek-summary">Không có dòng dữ liệu nào trong phạm vi đã chọn.</div>`;
      return;
    }

    const firstRow = rowNums[0];
    const lastRow = rowNums[rowNums.length - 1];
    const sampleRows = rowNums.length <= 4
      ? rowNums
      : [rowNums[0], rowNums[1], '...', rowNums[rowNums.length - 1]];

    const rowsHtml = sampleRows.map(r => {
      if (r === '...') {
        return `<tr class="diff-peek-dots"><td colspan="${cols.length + 1}">⋮ (${rowNums.length - 3} dòng ở giữa)</td></tr>`;
      }
      const cells = cols.map(c => `<td>${escapeHtml(getCellDisplay(sheet, r, c.num) || '')}</td>`).join('');
      return `<tr><td class="diff-peek-rnum">Dòng ${r}</td>${cells}</tr>`;
    }).join('');

    box.innerHTML = `
      <div class="diff-peek-summary">
        <span><strong>${rowNums.length}</strong> dòng dữ liệu</span>
        <span>Thứ tự: <strong>Dòng ${firstRow} → Dòng ${lastRow}</strong>${isRev ? ' <em>(Đã đảo ngược)</em>' : ''}</span>
      </div>
      <div class="diff-peek-table-wrap">
        <table class="diff-peek-table">
          <thead>
            <tr><th>Dòng</th>${cols.map(c => `<th>${escapeHtml(c.display)}</th>`).join('')}</tr>
          </thead>
          <tbody>${rowsHtml}</tbody>
        </table>
      </div>
    `;
  }

  // =========================================================================
  // 3. GHÉP CẶP CỘT ĐỐI CHIẾU
  // =========================================================================

  function populateColumnPickers() {
    const shA = getActiveSheet('A');
    const shB = getActiveSheet('B');
    const colsA = shA ? getColumns(shA, getHeaderRowNum('A')) : [];
    const colsB = shB ? getColumns(shB, getHeaderRowNum('B')) : [];

    const selA = el('diffAddColA');
    const selB = el('diffAddColB');
    if (selA) {
      const prevA = selA.value;
      selA.innerHTML = `<option value="0">(Không có ở File 1)</option>` +
        colsA.map(c => `<option value="${c.num}">${escapeHtml(c.display)}</option>`).join('');
      if (prevA && Array.from(selA.options).some(o => o.value === prevA)) selA.value = prevA;
      else if (colsA.length) selA.value = String(colsA[0].num);
    }
    if (selB) {
      const prevB = selB.value;
      selB.innerHTML = `<option value="0">(Không có ở File 2)</option>` +
        colsB.map(c => `<option value="${c.num}">${escapeHtml(c.display)}</option>`).join('');
      if (prevB && Array.from(selB.options).some(o => o.value === prevB)) selB.value = prevB;
      else if (colsB.length) selB.value = String(colsB[0].num);
    }
  }

  function makePairLabel(colsA, colsB, colA, colB) {
    const a = colsA.find(c => c.num === colA);
    const b = colsB.find(c => c.num === colB);
    const labelA = a ? (a.label || `Cột ${a.letter}`) : '(Trống)';
    const labelB = b ? (b.label || `Cột ${b.letter}`) : '(Trống)';
    if (a && b) {
      if (normalizeHeaderText(labelA) === normalizeHeaderText(labelB) && a.label) {
        return `${a.label} (${a.letter} ↔ ${b.letter})`;
      }
      return `${a.letter}: ${labelA} ↔ ${b.letter}: ${labelB}`;
    }
    if (a) return `Chỉ File 1 (${a.letter}: ${labelA})`;
    if (b) return `Chỉ File 2 (${b.letter}: ${labelB})`;
    return 'Cặp cột';
  }

  function autoMatchColumns(mode) {
    const shA = getActiveSheet('A');
    const shB = getActiveSheet('B');
    if (!shA || !shB) return;

    const colsA = getColumns(shA, getHeaderRowNum('A'));
    const colsB = getColumns(shB, getHeaderRowNum('B'));
    const pairs = [];

    if (mode === 'byName') {
      const usedB = new Set();
      for (const cA of colsA) {
        if (!cA.label) continue;
        const normA = normalizeHeaderText(cA.label);
        const matchB = colsB.find(cB => !usedB.has(cB.num) && cB.label && normalizeHeaderText(cB.label) === normA);
        if (matchB) {
          usedB.add(matchB.num);
          pairs.push({ colA: cA.num, colB: matchB.num, label: makePairLabel(colsA, colsB, cA.num, matchB.num) });
        }
      }
      if (!pairs.length) {
        showToast('Không tìm thấy tên cột trùng nhau; tự chuyển sang ghép theo vị trí cột.');
        return autoMatchColumns('byPosition');
      }
    } else {
      const maxCols = Math.max(colsA.length, colsB.length);
      for (let i = 0; i < maxCols; i++) {
        const cA = colsA[i] ? colsA[i].num : 0;
        const cB = colsB[i] ? colsB[i].num : 0;
        pairs.push({ colA: cA, colB: cB, label: makePairLabel(colsA, colsB, cA, cB) });
      }
    }

    state.pairs = pairs;
    renderPairsList();
  }

  function renderPairsList() {
    const list = el('diffPairsList');
    const countEl = el('diffColPairCount');
    const advBadge = el('diffAdvancedBadge');
    if (countEl) countEl.textContent = String(state.pairs.length);
    if (advBadge) advBadge.textContent = `${state.pairs.length} cặp cột`;
    if (!list) return;

    if (!state.bookA || !state.bookB) {
      list.innerHTML = `<div class="diff-pairs-empty">Hãy tải cả 2 file Excel để chọn cột đối chiếu.</div>`;
      return;
    }

    if (!state.pairs.length) {
      list.innerHTML = `<div class="diff-pairs-empty">Chưa có cặp cột nào. Hãy bấm <strong>Tự ghép theo tên tiêu đề</strong>, <strong>Ghép tất cả theo vị trí</strong> hoặc chọn ở bên dưới.</div>`;
      return;
    }

    const shA = getActiveSheet('A');
    const shB = getActiveSheet('B');
    const colsA = shA ? getColumns(shA, getHeaderRowNum('A')) : [];
    const colsB = shB ? getColumns(shB, getHeaderRowNum('B')) : [];

    list.innerHTML = '';
    state.pairs.forEach((p, idx) => {
      const cA = colsA.find(c => c.num === p.colA);
      const cB = colsB.find(c => c.num === p.colB);
      const row = document.createElement('div');
      row.className = 'diff-pair-chip';
      row.innerHTML = `
        <span class="diff-pair-idx">${idx + 1}</span>
        <span class="diff-pair-side is-a">${escapeHtml(cA ? cA.display : '(Không có ở File 1)')}</span>
        <span class="diff-pair-sep">↔</span>
        <span class="diff-pair-side is-b">${escapeHtml(cB ? cB.display : '(Không có ở File 2)')}</span>
        <button type="button" class="diff-pair-remove" title="Xóa cặp cột này" aria-label="Xóa cặp cột">×</button>
      `;
      row.querySelector('.diff-pair-remove').addEventListener('click', () => {
        state.pairs.splice(idx, 1);
        renderPairsList();
      });
      list.appendChild(row);
    });
  }

  // =========================================================================
  // 4. CHUẨN HÓA GIÁ TRỊ & THUẬT TOÁN ĐỐI CHIẾU
  // =========================================================================

  function parseFlexibleNumber(str) {
    const s = String(str).replace(/\u00a0/g, ' ').trim();
    if (!s) return null;

    // Tránh gộp mã có số 0 đứng đầu như "00123" thành "123"
    if (/^[+-]?0\d+$/.test(s)) return null;

    // Dạng chuẩn US/Excel: 1,234,567.89 hoặc 1,234,567
    if (/^[+-]?\d{1,3}(,\d{3})+(\.\d+)?$/.test(s)) {
      const n = Number(s.replace(/,/g, ''));
      return Number.isFinite(n) ? n : null;
    }
    // Dạng VN: 1.234.567,89 hoặc 1.234.567
    if (/^[+-]?\d{1,3}(\.\d{3})+(,\d+)?$/.test(s)) {
      const n = Number(s.replace(/\./g, '').replace(',', '.'));
      return Number.isFinite(n) ? n : null;
    }
    // Dạng số thuần có dấu chấm thập phân: 1234.50 hoặc -500
    if (/^[+-]?(0|[1-9]\d*)(\.\d+)?$/.test(s)) {
      const n = Number(s);
      return Number.isFinite(n) ? n : null;
    }
    // Dạng số có dấu phẩy thập phân đơn lẻ: 123,50
    if (/^[+-]?(0|[1-9]\d*)(,\d{1,2})$/.test(s)) {
      const n = Number(s.replace(',', '.'));
      return Number.isFinite(n) ? n : null;
    }
    return null;
  }

  function normalizeCellValue(raw, opts) {
    let s = String(raw ?? '');
    if (opts.trim) s = s.replace(/\s+/g, ' ').trim();
    if (opts.flexNum) {
      const num = parseFlexibleNumber(s);
      if (num !== null) return `__NUM__:${Object.is(num, -0) ? 0 : num}`;
    }
    if (!opts.caseSensitive) s = s.toLocaleLowerCase('vi-VN');
    return s;
  }

  function buildRowRecords(side, pairs, opts) {
    const sheet = getActiveSheet(side);
    const rowNums = extractRowNumbers(side, opts);
    return rowNums.map(rNum => {
      const rawVals = pairs.map(p => {
        const colNum = side === 'A' ? p.colA : p.colB;
        return colNum > 0 ? getCellDisplay(sheet, rNum, colNum) : '';
      });
      const normVals = rawVals.map(v => normalizeCellValue(v, opts));
      return {
        rowNum: rNum,
        rawVals,
        normVals,
        sig: normVals.join('\u0001')
      };
    });
  }

  // Hunt-Szymanski LCS cho 2 mảng ID số nguyên -> trả về mảng cặp [idxA, idxB] khớp hoàn toàn theo thứ tự
  function computeExactLcsAnchors(idsA, idsB) {
    const n = idsA.length;
    const m = idsB.length;
    let start = 0;
    while (start < n && start < m && idsA[start] === idsB[start]) start++;

    let endA = n - 1;
    let endB = m - 1;
    while (endA >= start && endB >= start && idsA[endA] === idsB[endB]) {
      endA--;
      endB--;
    }

    const anchors = [];
    for (let i = 0; i < start; i++) anchors.push([i, i]);

    const sliceLenA = endA - start + 1;
    const sliceLenB = endB - start + 1;

    if (sliceLenA > 0 && sliceLenB > 0) {
      const bPosMap = new Map();
      for (let j = endB; j >= start; j--) {
        const id = idsB[j];
        let arr = bPosMap.get(id);
        if (!arr) {
          arr = [];
          bPosMap.set(id, arr);
        }
        arr.push(j); // Lưu giảm dần để LIS giữ thứ tự tăng chặt
      }

      const tails = [];
      const links = [];

      for (let i = start; i <= endA; i++) {
        const positions = bPosMap.get(idsA[i]);
        if (!positions) continue;

        for (const j of positions) {
          // Binary search tìm vị trí nhỏ nhất k sao cho tails[k] >= j
          let lo = 0;
          let hi = tails.length;
          while (lo < hi) {
            const mid = (lo + hi) >> 1;
            if (tails[mid] < j) lo = mid + 1;
            else hi = mid;
          }
          tails[lo] = j;
          links[lo] = { i, j, prev: lo > 0 ? links[lo - 1] : null };
        }
      }

      if (links.length) {
        const midAnchors = [];
        let cur = links[links.length - 1];
        while (cur) {
          midAnchors.push([cur.i, cur.j]);
          cur = cur.prev;
        }
        midAnchors.reverse();
        anchors.push(...midAnchors);
      }
    }

    for (let i = endA + 1, j = endB + 1; i < n && j < m; i++, j++) {
      anchors.push([i, j]);
    }
    return anchors;
  }

  function pairGapRows(chunkA, chunkB) {
    const out = [];
    if (!chunkA.length && !chunkB.length) return out;
    if (chunkA.length && !chunkB.length) {
      for (const a of chunkA) out.push({ status: 'onlyA', a, b: null, diffCols: [] });
      return out;
    }
    if (!chunkA.length && chunkB.length) {
      for (const b of chunkB) out.push({ status: 'onlyB', a: null, b, diffCols: [] });
      return out;
    }

    // Nếu cả 2 bên đều có đúng 1 dòng trong khoảng lệch -> ghép trực tiếp để chỉ rõ ô lệch
    if (chunkA.length === 1 && chunkB.length === 1) {
      const a = chunkA[0];
      const b = chunkB[0];
      const diffCols = [];
      for (let c = 0; c < a.normVals.length; c++) {
        if (a.normVals[c] !== b.normVals[c]) diffCols.push(c);
      }
      out.push({ status: diffCols.length ? 'changed' : 'matched', a, b, diffCols });
      return out;
    }

    // Ghép các dòng tương đồng nhất trong khoảng lệch (khi có >=1 cột giống nhau)
    const usedB = new Set();
    const matchedPairs = new Map(); // idxInChunkA -> idxInChunkB

    for (let i = 0; i < chunkA.length; i++) {
      const a = chunkA[i];
      let bestJ = -1;
      let bestScore = 0;

      for (let j = 0; j < chunkB.length; j++) {
        if (usedB.has(j)) continue;
        const b = chunkB[j];
        let sameCount = 0;
        for (let c = 0; c < a.normVals.length; c++) {
          if (a.normVals[c] === b.normVals[c] && a.normVals[c] !== '') sameCount++;
        }
        if (sameCount > bestScore) {
          bestScore = sameCount;
          bestJ = j;
        }
      }

      if (bestJ !== -1 && bestScore >= 1) {
        usedB.add(bestJ);
        matchedPairs.set(i, bestJ);
      }
    }

    // Giữ thứ tự hiển thị tự nhiên của các dòng trong khoảng lệch
    let nextUnpairedB = 0;
    const flushUnpairedBUpTo = limitB => {
      while (nextUnpairedB < limitB) {
        if (!usedB.has(nextUnpairedB)) {
          out.push({ status: 'onlyB', a: null, b: chunkB[nextUnpairedB], diffCols: [] });
          usedB.add(nextUnpairedB);
        }
        nextUnpairedB++;
      }
    };

    for (let i = 0; i < chunkA.length; i++) {
      if (matchedPairs.has(i)) {
        const j = matchedPairs.get(i);
        flushUnpairedBUpTo(j);
        const a = chunkA[i];
        const b = chunkB[j];
        const diffCols = [];
        for (let c = 0; c < a.normVals.length; c++) {
          if (a.normVals[c] !== b.normVals[c]) diffCols.push(c);
        }
        out.push({ status: diffCols.length ? 'changed' : 'matched', a, b, diffCols });
        nextUnpairedB = Math.max(nextUnpairedB, j + 1);
      } else {
        out.push({ status: 'onlyA', a: chunkA[i], b: null, diffCols: [] });
      }
    }
    flushUnpairedBUpTo(chunkB.length);
    return out;
  }

  function runComparison() {
    if (!state.bookA || !state.bookB) {
      showToast('Vui lòng chọn đầy đủ File 1 và File 2 trước khi đối chiếu.');
      return;
    }

    if (!state.pairs.length) {
      autoMatchColumns(state.scope === 'all' ? 'byName' : 'byPosition');
    }
    if (!state.pairs.length) {
      showToast('Hãy chọn ít nhất 1 cặp cột để đối chiếu.');
      return;
    }

    const opts = {
      skipBlank: el('diffSkipBlank')?.checked ?? true,
      trim: el('diffTrim')?.checked ?? true,
      flexNum: el('diffFlexNum')?.checked ?? true,
      caseSensitive: el('diffCaseSensitive')?.checked ?? false,
      skipTotals: el('diffSkipTotals')?.checked ?? false
    };
    const mode = el('diffCompareMode')?.value || 'smart';

    let rowsA, rowsB;
    try {
      rowsA = buildRowRecords('A', state.pairs, opts);
      rowsB = buildRowRecords('B', state.pairs, opts);
    } catch (err) {
      showToast(err.message);
      return;
    }

    const diffRows = [];

    if (mode === 'positional') {
      const maxLen = Math.max(rowsA.length, rowsB.length);
      for (let i = 0; i < maxLen; i++) {
        const a = rowsA[i] || null;
        const b = rowsB[i] || null;
        if (a && !b) {
          diffRows.push({ status: 'onlyA', a, b: null, diffCols: [] });
        } else if (!a && b) {
          diffRows.push({ status: 'onlyB', a: null, b, diffCols: [] });
        } else {
          const diffCols = [];
          for (let c = 0; c < state.pairs.length; c++) {
            if (a.normVals[c] !== b.normVals[c]) diffCols.push(c);
          }
          diffRows.push({ status: diffCols.length ? 'changed' : 'matched', a, b, diffCols });
        }
      }
    } else if (mode === 'unordered') {
      const sigBucketsB = new Map();
      rowsB.forEach((b, idx) => {
        let q = sigBucketsB.get(b.sig);
        if (!q) { q = []; sigBucketsB.set(b.sig, q); }
        q.push(idx);
      });
      const matchedB = new Set();

      for (const a of rowsA) {
        const q = sigBucketsB.get(a.sig);
        if (q && q.length) {
          const idxB = q.shift();
          matchedB.add(idxB);
          diffRows.push({ status: 'matched', a, b: rowsB[idxB], diffCols: [] });
        } else {
          diffRows.push({ status: 'onlyA', a, b: null, diffCols: [] });
        }
      }
      rowsB.forEach((b, idx) => {
        if (!matchedB.has(idx)) {
          diffRows.push({ status: 'onlyB', a: null, b, diffCols: [] });
        }
      });
    } else {
      // Smart LCS sequence alignment
      const sigToId = new Map();
      let nextId = 1;
      const getId = sig => {
        let id = sigToId.get(sig);
        if (!id) { id = nextId++; sigToId.set(sig, id); }
        return id;
      };
      const idsA = rowsA.map(r => getId(r.sig));
      const idsB = rowsB.map(r => getId(r.sig));
      const anchors = computeExactLcsAnchors(idsA, idsB);

      let prevA = 0;
      let prevB = 0;
      for (const [ancA, ancB] of anchors) {
        const gapA = rowsA.slice(prevA, ancA);
        const gapB = rowsB.slice(prevB, ancB);
        diffRows.push(...pairGapRows(gapA, gapB));
        diffRows.push({ status: 'matched', a: rowsA[ancA], b: rowsB[ancB], diffCols: [] });
        prevA = ancA + 1;
        prevB = ancB + 1;
      }
      diffRows.push(...pairGapRows(rowsA.slice(prevA), rowsB.slice(prevB)));
    }

    const counts = { all: diffRows.length, matched: 0, changed: 0, onlyA: 0, onlyB: 0 };
    for (const r of diffRows) counts[r.status]++;

    state.results = {
      rows: diffRows,
      pairs: state.pairs.slice(),
      counts,
      totalA: rowsA.length,
      totalB: rowsB.length
    };
    state.filterStatus = 'all';
    state.onlyDiff = false;
    state.searchText = '';
    state.currentPage = 1;
    if (el('diffOnlyDiffToggle')) el('diffOnlyDiffToggle').checked = false;
    if (el('diffSearchInput')) el('diffSearchInput').value = '';

    renderResults();
    el('diffResultCard')?.scrollIntoView({ behavior: 'smooth', block: 'start' });
  }

  // =========================================================================
  // 5. HIỂN THỊ KẾT QUẢ ĐỐI CHIẾU
  // =========================================================================

  function renderResults() {
    const card = el('diffResultCard');
    if (!card || !state.results) return;
    card.style.display = 'block';

    const { rows, pairs, counts, totalA, totalB } = state.results;
    const banner = el('diffResultBanner');
    const isPerfect = counts.changed === 0 && counts.onlyA === 0 && counts.onlyB === 0 && counts.all > 0;

    if (banner) {
      if (isPerfect) {
        banner.className = 'diff-result-banner is-perfect';
        banner.innerHTML = `
          <div class="diff-banner-icon">✓</div>
          <div>
            <strong>Hai file khớp hoàn toàn 100%!</strong>
            <p>Đã đối chiếu ${counts.matched} dòng trên ${pairs.length} cặp cột — không có dòng thiếu, dòng dư hay sai lệch ô nào.</p>
          </div>
        `;
      } else if (!counts.all) {
        banner.className = 'diff-result-banner is-warn';
        banner.innerHTML = `
          <div class="diff-banner-icon">!</div>
          <div>
            <strong>Không có dòng dữ liệu nào để đối chiếu</strong>
            <p>Hãy kiểm tra lại phạm vi dòng hoặc sheet đã chọn ở File 1 và File 2.</p>
          </div>
        `;
      } else {
        banner.className = 'diff-result-banner is-diff';
        banner.innerHTML = `
          <div class="diff-banner-icon">⚠</div>
          <div>
            <strong>Phát hiện sự khác biệt giữa 2 file</strong>
            <p>File 1 có <b>${totalA}</b> dòng · File 2 có <b>${totalB}</b> dòng · Khớp: <b>${counts.matched}</b> · Lệch ô: <b>${counts.changed}</b> · Thiếu ở File 2: <b>${counts.onlyA}</b> · Dư ở File 2: <b>${counts.onlyB}</b></p>
          </div>
        `;
      }
    }

    el('diffCountAll').textContent = String(counts.all);
    el('diffCountMatched').textContent = String(counts.matched);
    el('diffCountChanged').textContent = String(counts.changed);
    el('diffCountOnlyA').textContent = String(counts.onlyA);
    el('diffCountOnlyB').textContent = String(counts.onlyB);

    document.querySelectorAll('.diff-kpi-btn[data-diff-filter]').forEach(btn => {
      btn.classList.toggle('active', btn.dataset.diffFilter === state.filterStatus);
    });

    const term = deaccent(state.searchText);
    const filtered = rows.filter(r => {
      if (state.onlyDiff && r.status === 'matched') return false;
      if (state.filterStatus !== 'all' && r.status !== state.filterStatus) return false;
      if (term) {
        const hay = [
          r.a ? `dong ${r.a.rowNum} ${r.a.rawVals.join(' ')}` : '',
          r.b ? `dong ${r.b.rowNum} ${r.b.rawVals.join(' ')}` : ''
        ].join(' ');
        if (!deaccent(hay).includes(term)) return false;
      }
      return true;
    });

    const totalPages = Math.max(1, Math.ceil(filtered.length / PAGE_SIZE));
    if (state.currentPage > totalPages) state.currentPage = totalPages;
    const startIdx = (state.currentPage - 1) * PAGE_SIZE;
    const pageRows = filtered.slice(startIdx, startIdx + PAGE_SIZE);

    el('diffResultMeta').textContent = `Đang hiện ${filtered.length} / ${counts.all} kết quả`;
    el('diffPageInfo').textContent = `Trang ${state.currentPage} / ${totalPages}`;
    if (el('diffPrevPage')) el('diffPrevPage').disabled = state.currentPage <= 1;
    if (el('diffNextPage')) el('diffNextPage').disabled = state.currentPage >= totalPages;

    const wrap = el('diffTableWrap');
    if (!wrap) return;

    if (!pageRows.length) {
      wrap.innerHTML = `<div class="empty"><div class="big">✓</div><strong>Không có dòng nào thuộc bộ lọc hiện tại</strong></div>`;
      return;
    }

    const statusBadge = status => {
      if (status === 'matched') return '<span class="diff-status-badge is-match">✓ Khớp</span>';
      if (status === 'changed') return '<span class="diff-status-badge is-changed">≠ Lệch ô</span>';
      if (status === 'onlyA') return '<span class="diff-status-badge is-only-a">− Thiếu ở F2 (Dư F1)</span>';
      return '<span class="diff-status-badge is-only-b">+ Thiếu ở F1 (Dư F2)</span>';
    };

    const headCols = pairs.map(p => `<th title="${escapeHtml(p.label)}">${escapeHtml(p.label)}</th>`).join('');
    const bodyHtml = pageRows.map((item, idx) => {
      const stt = startIdx + idx + 1;
      const rA = item.a ? `Dòng ${item.a.rowNum}` : '—';
      const rB = item.b ? `Dòng ${item.b.rowNum}` : '—';
      const diffSet = new Set(item.diffCols || []);

      const cellsHtml = pairs.map((p, cIdx) => {
        if (item.status === 'onlyA') {
          const v = item.a.rawVals[cIdx];
          return `<td>${v ? escapeHtml(v) : '<span class="muted">(trống)</span>'}</td>`;
        }
        if (item.status === 'onlyB') {
          const v = item.b.rawVals[cIdx];
          return `<td>${v ? escapeHtml(v) : '<span class="muted">(trống)</span>'}</td>`;
        }
        const vA = item.a.rawVals[cIdx];
        const vB = item.b.rawVals[cIdx];
        if (diffSet.has(cIdx)) {
          return `
            <td class="diff-cell-mismatch">
              <div class="diff-val-pair">
                <span class="diff-val-tag is-a"><b>F1:</b> ${vA !== '' ? escapeHtml(vA) : '<i>(trống)</i>'}</span>
                <span class="diff-val-tag is-b"><b>F2:</b> ${vB !== '' ? escapeHtml(vB) : '<i>(trống)</i>'}</span>
              </div>
            </td>
          `;
        }
        return `<td>${vA !== '' ? escapeHtml(vA) : (vB !== '' ? escapeHtml(vB) : '<span class="muted">(trống)</span>')}</td>`;
      }).join('');

      return `
        <tr class="diff-row is-${item.status}">
          <td class="rowid">${stt}</td>
          <td>${statusBadge(item.status)}</td>
          <td class="diff-rnum">${rA}</td>
          <td class="diff-rnum">${rB}</td>
          ${cellsHtml}
        </tr>
      `;
    }).join('');

    wrap.innerHTML = `
      <table class="data-table diff-result-table">
        <thead>
          <tr>
            <th>#</th>
            <th>Trạng thái</th>
            <th>File 1</th>
            <th>File 2</th>
            ${headCols}
          </tr>
        </thead>
        <tbody>${bodyHtml}</tbody>
      </table>
    `;
  }

  // =========================================================================
  // 6. SỰ KIỆN GIAO DIỆN
  // =========================================================================

  function updateRunBtnState() {
    const btn = el('diffRunBtn');
    if (btn) btn.disabled = !(state.bookA && state.bookB);
  }

  function applySheetDefaults(side) {
    const sheet = getActiveSheet(side);
    if (!sheet) return;
    const hdr = detectHeaderRow(sheet);
    const hdrInput = el(`diffHeader${side}`);
    const startInput = el(`diffStart${side}`);
    const endInput = el(`diffEnd${side}`);
    if (hdrInput) hdrInput.value = String(hdr);
    if (startInput) startInput.value = String(Math.min(Math.max(sheet.maxRow, 1), hdr + 1));
    if (endInput) endInput.value = '';
    if (el(`diffRowPick${side}`)) el(`diffRowPick${side}`).value = '';
    if (side === 'A' && el('diffQuickRowsA')) el('diffQuickRowsA').value = '';
    if (side === 'B' && el('diffQuickRowsB')) el('diffQuickRowsB').value = '';
  }

  async function handleFileSelect(side, file) {
    if (!file) return;
    const nameEl = el(`diffFileName${side}`);
    const dropEl = el(`diffDrop${side}`);
    if (dropEl) dropEl.classList.add('is-loading');

    try {
      const book = await parseExcelBook(file);
      if (side === 'A') state.bookA = book;
      else state.bookB = book;

      if (nameEl) {
        nameEl.textContent = `✓ ${book.fileName} (${book.sheets.length} sheet)`;
        nameEl.classList.add('ready');
      }

      const sel = el(`diffSheet${side}`);
      if (sel) {
        sel.innerHTML = book.sheets.map(s => `<option value="${escapeHtml(s.name)}">${escapeHtml(s.name)}</option>`).join('');
      }
      if (el(`diffControls${side}`)) el(`diffControls${side}`).style.display = 'block';

      applySheetDefaults(side);
      renderFilePeek(side);
      populateColumnPickers();

      if (state.bookA && state.bookB) {
        autoMatchColumns(state.scope === 'all' ? 'byName' : 'byPosition');
      } else {
        renderPairsList();
      }
      updateRunBtnState();
    } catch (err) {
      showToast(err.message || 'Không đọc được file Excel.');
    } finally {
      if (dropEl) dropEl.classList.remove('is-loading');
    }
  }

  function setupDropzone(side) {
    const drop = el(`diffDrop${side}`);
    const input = el(`diffInput${side}`);
    if (!drop || !input) return;

    input.addEventListener('change', e => {
      const file = e.target.files?.[0];
      if (file) handleFileSelect(side, file);
    });

    ['dragenter', 'dragover'].forEach(evt => {
      drop.addEventListener(evt, e => {
        e.preventDefault();
        drop.classList.add('is-dragging');
      });
    });

    ['dragleave', 'drop'].forEach(evt => {
      drop.addEventListener(evt, e => {
        e.preventDefault();
        drop.classList.remove('is-dragging');
      });
    });

    drop.addEventListener('drop', e => {
      const file = e.dataTransfer?.files?.[0];
      if (file) handleFileSelect(side, file);
    });

    ['diffSheetA', 'diffSheetB'].forEach(id => {
      if (id !== `diffSheet${side}`) return;
      el(id)?.addEventListener('change', () => {
        applySheetDefaults(side);
        renderFilePeek(side);
        populateColumnPickers();
        if (state.bookA && state.bookB) autoMatchColumns('byName');
      });
    });

    [`diffHeader${side}`, `diffStart${side}`, `diffEnd${side}`, `diffRowPick${side}`, `diffReverse${side}`].forEach(id => {
      const node = el(id);
      if (!node) return;
      const evt = node.type === 'checkbox' ? 'change' : 'input';
      node.addEventListener(evt, () => {
        if (id === `diffHeader${side}`) {
          const hdr = getHeaderRowNum(side);
          const startEl = el(`diffStart${side}`);
          if (startEl && hdr >= 0) startEl.value = String(hdr + 1);
          populateColumnPickers();
          renderPairsList();
        }
        renderFilePeek(side);
      });
    });
  }

  function setScope(scope) {
    state.scope = scope;
    document.querySelectorAll('.diff-scope-tab[data-diff-scope]').forEach(tab => {
      tab.classList.toggle('active', tab.dataset.diffScope === scope);
    });

    const adv = el('diffAdvancedDetails');
    const rowBox = el('diffRowScopeBox');
    const colBox = el('diffColMapper');
    if (adv) adv.open = scope === 'cols' || scope === 'rows';
    if (rowBox) rowBox.classList.toggle('is-emphasized', scope === 'rows');
    if (colBox) colBox.classList.toggle('is-emphasized', scope === 'cols');

    const titleEl = el('diffMapperTitle');
    if (titleEl) {
      if (scope === 'all') {
        titleEl.innerHTML = `Cặp cột đối chiếu toàn bộ 2 file (<strong id="diffColPairCount">${state.pairs.length}</strong> cặp)`;
      } else if (scope === 'cols') {
        titleEl.innerHTML = `Chọn các cột cần đối chiếu (<strong id="diffColPairCount">${state.pairs.length}</strong> cặp)`;
      } else {
        titleEl.innerHTML = `Cột đối chiếu trên các dòng đã chọn (<strong id="diffColPairCount">${state.pairs.length}</strong> cặp)`;
      }
    }
  }

  function resetAll() {
    state.bookA = null;
    state.bookB = null;
    state.pairs = [];
    state.results = null;
    state.filterStatus = 'all';
    state.onlyDiff = false;
    state.searchText = '';
    state.currentPage = 1;

    ['A', 'B'].forEach(side => {
      if (el(`diffInput${side}`)) el(`diffInput${side}`).value = '';
      if (el(`diffFileName${side}`)) {
        el(`diffFileName${side}`).textContent = 'Chưa chọn file';
        el(`diffFileName${side}`).classList.remove('ready');
      }
      if (el(`diffControls${side}`)) el(`diffControls${side}`).style.display = 'none';
      if (el(`diffReverse${side}`)) el(`diffReverse${side}`).checked = false;
      if (el(`diffRowPick${side}`)) el(`diffRowPick${side}`).value = '';
      if (el(`diffPeek${side}`)) el(`diffPeek${side}`).innerHTML = '';
    });

    if (el('diffResultCard')) el('diffResultCard').style.display = 'none';
    if (el('diffAdvancedDetails')) el('diffAdvancedDetails').open = false;
    setScope('all');
    populateColumnPickers();
    renderPairsList();
    updateRunBtnState();
  }

  function init() {
    if (!el('diffPage')) return;

    setupDropzone('A');
    setupDropzone('B');

    document.querySelectorAll('.diff-scope-tab[data-diff-scope]').forEach(tab => {
      tab.addEventListener('click', () => setScope(tab.dataset.diffScope));
    });

    el('diffAutoMatchByNameBtn')?.addEventListener('click', () => autoMatchColumns('byName'));
    el('diffAutoMatchByPosBtn')?.addEventListener('click', () => autoMatchColumns('byPosition'));
    el('diffClearPairsBtn')?.addEventListener('click', () => {
      state.pairs = [];
      renderPairsList();
    });

    el('diffAddPairBtn')?.addEventListener('click', () => {
      if (!state.bookA || !state.bookB) {
        showToast('Hãy chọn cả 2 file Excel trước.');
        return;
      }
      const colA = parseInt(el('diffAddColA')?.value || '0', 10);
      const colB = parseInt(el('diffAddColB')?.value || '0', 10);
      if (!colA && !colB) {
        showToast('Hãy chọn ít nhất 1 cột từ File 1 hoặc File 2.');
        return;
      }
      const shA = getActiveSheet('A');
      const shB = getActiveSheet('B');
      const colsA = shA ? getColumns(shA, getHeaderRowNum('A')) : [];
      const colsB = shB ? getColumns(shB, getHeaderRowNum('B')) : [];
      state.pairs.push({ colA, colB, label: makePairLabel(colsA, colsB, colA, colB) });
      renderPairsList();
    });

    ['diffSkipBlank', 'diffSkipTotals'].forEach(id => {
      el(id)?.addEventListener('change', () => {
        if (state.bookA) renderFilePeek('A');
        if (state.bookB) renderFilePeek('B');
      });
    });

    el('diffRunBtn')?.addEventListener('click', runComparison);
    el('diffResetBtn')?.addEventListener('click', resetAll);

    document.querySelectorAll('.diff-kpi-btn[data-diff-filter]').forEach(btn => {
      btn.addEventListener('click', () => {
        state.filterStatus = btn.dataset.diffFilter;
        state.currentPage = 1;
        renderResults();
      });
    });

    el('diffOnlyDiffToggle')?.addEventListener('change', e => {
      state.onlyDiff = e.target.checked;
      state.currentPage = 1;
      renderResults();
    });

    el('diffSearchInput')?.addEventListener('input', e => {
      state.searchText = e.target.value;
      state.currentPage = 1;
      renderResults();
    });

    el('diffPrevPage')?.addEventListener('click', () => {
      if (state.currentPage > 1) {
        state.currentPage--;
        renderResults();
      }
    });

    el('diffNextPage')?.addEventListener('click', () => {
      state.currentPage++;
      renderResults();
    });
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
})();
