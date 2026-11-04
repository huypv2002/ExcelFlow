/**
 * ExcelFlow Suite — Module Gộp nhiều file (Excel & PDF)
 * Tự động gom & ghép theo loại chứng từ: Bảng kê mua vào, bán ra, báo có, báo nợ, hoặc nối tiếp dòng.
 * Hỗ trợ ghép PDF với PDF (PDF-to-PDF) và xuất Excel (.xlsx) đa sheet.
 */
(function() {
  'use strict';

  // State
  let mergerFiles = []; // Array<{ id, file, name, size, type: 'excel'|'pdf', sheets: Array<{ name, grid }> }>
  let fileIdCounter = 0;
  let mergedResult = null; // { type: 'xlsx'|'pdf', blob, filename, sheets: Array<{ name, grid }>, pdfDoc, pdfStats }
  let activeSheetIndex = 0;
  let currentPage = 1;
  const rowsPerPage = 50;
  let searchFilter = '';

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
        showToast.timer = setTimeout(() => t.classList.remove('show'), 2500);
      }
    }
  }

  function formatBytes(bytes) {
    if (bytes === 0) return '0 B';
    const k = 1024;
    const sizes = ['B', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + ' ' + sizes[i];
  }

  function sanitizeDate(str) {
    if (!str || typeof str !== 'string') return str;
    // Chuyển đổi định dạng ngày tiếng Nhật YYYY年MM月DD日 -> DD/MM/YYYY
    return str.replace(/(\d{4})年(\d{1,2})月(\d{1,2})日/g, (match, y, m, d) => {
      const dd = d.padStart(2, '0');
      const mm = m.padStart(2, '0');
      return `${dd}/${mm}/${y}`;
    });
  }

  // Đọc nội dung file Excel thành danh sách sheets và grid 2D array
  async function parseExcelFile(file) {
    const JSZipLib = window.JSZip;
    if (!JSZipLib) throw new Error('Chưa tải được thư viện JSZip');
    const bytes = await file.arrayBuffer();
    const zip = await JSZipLib.loadAsync(bytes);

    // 1. Đọc sharedStrings.xml
    const sharedStrings = [];
    const sharedFile = zip.file('xl/sharedStrings.xml');
    if (sharedFile) {
      const sXml = await sharedFile.async('string');
      const parser = new DOMParser();
      const sDoc = parser.parseFromString(sXml, 'application/xml');
      const siNodes = sDoc.getElementsByTagName('si');
      for (let i = 0; i < siNodes.length; i++) {
        const tNodes = siNodes[i].getElementsByTagName('t');
        let text = '';
        for (let j = 0; j < tNodes.length; j++) text += tNodes[j].textContent || '';
        sharedStrings.push(text);
      }
    }

    // 2. Đọc danh sách sheet từ xl/workbook.xml & rels
    const relsFile = zip.file('xl/_rels/workbook.xml.rels');
    const relsMap = {};
    if (relsFile) {
      const rXml = await relsFile.async('string');
      const rDoc = new DOMParser().parseFromString(rXml, 'application/xml');
      const relNodes = rDoc.getElementsByTagName('Relationship');
      for (let i = 0; i < relNodes.length; i++) {
        const rId = relNodes[i].getAttribute('Id');
        let target = relNodes[i].getAttribute('Target') || '';
        if (target.startsWith('/')) target = target.slice(1);
        else if (!target.startsWith('xl/')) target = 'xl/' + target;
        relsMap[rId] = target;
      }
    }

    const wbFile = zip.file('xl/workbook.xml');
    if (!wbFile) throw new Error('Không tìm thấy xl/workbook.xml trong file Excel');
    const wbXml = await wbFile.async('string');
    const wbDoc = new DOMParser().parseFromString(wbXml, 'application/xml');
    const sheetNodes = wbDoc.getElementsByTagName('sheet');
    const sheets = [];

    for (let i = 0; i < sheetNodes.length; i++) {
      const sName = sheetNodes[i].getAttribute('name') || `Sheet ${i + 1}`;
      const rId = sheetNodes[i].getAttribute('r:id') || sheetNodes[i].getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id');
      const sheetPath = relsMap[rId];
      if (!sheetPath || !zip.file(sheetPath)) continue;

      const sheetXml = await zip.file(sheetPath).async('string');
      const sheetDoc = new DOMParser().parseFromString(sheetXml, 'application/xml');
      const rowNodes = sheetDoc.getElementsByTagName('row');

      const grid = [];
      for (let r = 0; r < rowNodes.length; r++) {
        const rowNode = rowNodes[r];
        const rowIdx = parseInt(rowNode.getAttribute('r') || String(r + 1), 10);
        const cNodes = rowNode.getElementsByTagName('c');
        const rowData = [];

        for (let c = 0; c < cNodes.length; c++) {
          const cNode = cNodes[c];
          const ref = cNode.getAttribute('r') || '';
          const colLetters = (ref.match(/^[A-Z]+/i) || ['A'])[0].toUpperCase();
          let colIdx = 0;
          for (let k = 0; k < colLetters.length; k++) {
            colIdx = colIdx * 26 + (colLetters.charCodeAt(k) - 64);
          }
          const zeroColIdx = colIdx - 1;

          const typ = cNode.getAttribute('t') || '';
          let val = '';
          if (typ === 's') {
            const v = cNode.getElementsByTagName('v')[0];
            if (v && v.textContent !== '') {
              const sIdx = parseInt(v.textContent, 10);
              val = sharedStrings[sIdx] ?? '';
            }
          } else if (typ === 'inlineStr') {
            const isNode = cNode.getElementsByTagName('is')[0];
            if (isNode) {
              const tNodes = isNode.getElementsByTagName('t');
              for (let k = 0; k < tNodes.length; k++) val += tNodes[k].textContent || '';
            }
          } else {
            const v = cNode.getElementsByTagName('v')[0];
            if (v) val = v.textContent || '';
          }

          val = sanitizeDate(val.trim());
          while (rowData.length < zeroColIdx) rowData.push('');
          rowData[zeroColIdx] = val;
        }

        // Lọc bỏ hàng rỗng hoàn toàn
        if (rowData.some(cell => cell !== '')) {
          while (grid.length < rowIdx - 1) grid.push([]);
          grid[rowIdx - 1] = rowData;
        }
      }

      // Chuẩn hóa grid bỏ các hàng rỗng ở giữa
      const cleanGrid = grid.filter(r => r && r.length && r.some(cell => cell !== ''));
      if (cleanGrid.length > 0) {
        sheets.push({ name: sName, grid: cleanGrid });
      }
    }

    return sheets;
  }

  // Đọc nội dung file PDF thành sheets và grid
  async function parsePdfFile(file) {
    if (typeof window.efParsePdfToWorkbook !== 'function') {
      throw new Error('Chưa nạp được chức năng phân tích PDF.');
    }
    const zip = await window.efParsePdfToWorkbook(file);
    if (zip && Array.isArray(zip._sheets) && zip._sheets.length) {
      return zip._sheets.map(s => ({
        name: s.name,
        grid: s.grid.map(row => row.map(sanitizeDate))
      }));
    }
    return [{ name: 'Bảng PDF', grid: [] }];
  }

  // Nhận diện phân loại kế toán cho một sheet
  function detectAccountingCategory(sheetName, grid) {
    const nameNorm = (sheetName || '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
    const firstRowsText = (grid.slice(0, 8).map(r => r.join(' ')).join(' ')).toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');

    if (/mua\s*vao|bkmv|01-2/.test(nameNorm) || /nguoi\s*ban|mst\s*nguoi\s*ban|mua\s*hang\s*hoa|hang\s*hoa.*?mua\s*vao/i.test(firstRowsText)) {
      return { id: 'muavao', title: 'Bảng kê mua vào', code: '01-2' };
    }
    if (/ban\s*ra|bkbr|01-1/.test(nameNorm) || /nguoi\s*mua|mst\s*nguoi\s*mua|doanh\s*so\s*ban|hang\s*hoa.*?ban\s*ra/i.test(firstRowsText)) {
      return { id: 'banra', title: 'Bảng kê bán ra', code: '01-1' };
    }
    if (/bao\s*co|phat\s*sinh\s*co|ghi\s*co|gbc|tien\s*vao|thu\s*tien/.test(nameNorm) || /bao\s*co|ghi\s*co|phat\s*sinh\s*co/i.test(firstRowsText)) {
      return { id: 'baoco', title: 'Giấy báo có', code: 'GBC' };
    }
    if (/bao\s*no|phat\s*sinh\s*no|ghi\s*no|gbn|tien\s*ra|chi\s*tien/.test(nameNorm) || /bao\s*no|ghi\s*no|phat\s*sinh\s*no/i.test(firstRowsText)) {
      return { id: 'baono', title: 'Giấy báo nợ', code: 'GBN' };
    }
    return { id: 'general', title: sheetName || 'Bảng tính', code: 'DATA' };
  }

  // Tìm dòng tiêu đề chuẩn của bảng
  function findHeaderRowIndex(grid) {
    if (!grid || !grid.length) return 0;
    const keywords = ['stt', 'serial', 'so', 'ngay', 'date', 'ten', 'khach', 'mst', 'ma', 'tien', 'thue', 'tong', 'noi dung', 'dien giai', 'hoa don', 'mat hang', 'tk'];
    let bestIdx = 0, maxScore = -1;
    const limit = Math.min(grid.length, 15);

    for (let r = 0; r < limit; r++) {
      const row = grid[r] || [];
      const texts = row.map(c => String(c || '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim()).filter(Boolean);
      if (!texts.length) continue;
      let hits = 0;
      for (const t of texts) {
        if (keywords.some(kw => t.includes(kw))) hits++;
      }
      const score = hits * 4 + texts.length;
      if (hits >= 2 && score > maxScore) {
        maxScore = score;
        bestIdx = r;
      } else if (texts.length >= 3 && score > maxScore && maxScore < 4) {
        maxScore = score;
        bestIdx = r;
      }
    }
    return bestIdx;
  }

  // Kiểm tra dòng tổng cộng
  function isSummaryRow(row) {
    if (!row || !row.length) return false;
    const text = row.slice(0, 5).join(' ').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
    return /^(cong|tong|tong\s*cong|tong\s*nam|tong\s*quy|so\s*du|total)/.test(text.trim()) ||
           /\b(tong\s*cong|tong\s*phat\s*sinh|so\s*du\s*cuoi)\b/.test(text);
  }

  // Kiểm tra dòng tiêu đề lặp lại
  function isHeaderRow(row, sampleHeader) {
    if (!row || !row.length) return false;
    const rowStr = row.map(c => String(c || '').trim().toLowerCase()).filter(Boolean).join('|');
    if (sampleHeader) {
      const headerStr = sampleHeader.map(c => String(c || '').trim().toLowerCase()).filter(Boolean).join('|');
      if (rowStr === headerStr) return true;
    }
    const textNorm = row.join(' ').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
    return /ky\s*hieu.*?mau|so\s*hoa\s*don|ngay\s*phat\s*hanh|ten\s*nguoi\s*ban|ten\s*nguoi\s*mua|ma\s*so\s*thue/i.test(textNorm);
  }

  // Thêm file vào hàng đợi
  async function addFiles(fileList) {
    const zone = el('mergerDrop');
    if (zone) zone.classList.add('is-loading');

    let addedCount = 0;
    for (const file of fileList) {
      const isExcel = /\.(xlsx|xlsm)$/i.test(file.name);
      const isPdf = /\.pdf$/i.test(file.name);
      if (!isExcel && !isPdf) continue;

      // Tránh thêm trùng cùng tên + cùng kích thước
      if (mergerFiles.some(f => f.name === file.name && f.size === file.size)) continue;

      try {
        let sheets = [];
        if (isExcel) {
          sheets = await parseExcelFile(file);
        } else if (isPdf) {
          sheets = await parsePdfFile(file);
        }

        const entry = {
          id: ++fileIdCounter,
          file,
          name: file.name,
          size: file.size,
          type: isPdf ? 'pdf' : 'excel',
          sheets
        };

        mergerFiles.push(entry);
        addedCount++;
      } catch (err) {
        console.error('Lỗi khi đọc file', file.name, err);
        showToast(`Không đọc được file "${file.name}": ${err.message}`);
      }
    }

    if (zone) zone.classList.remove('is-loading');
    renderFileList();
    updateMergerOptions();

    if (addedCount > 0) {
      showToast(`✓ Đã thêm ${addedCount} file vào danh sách gộp!`);
    }
  }

  // Cập nhật giao diện danh sách file
  function renderFileList() {
    const container = el('mergerFileList');
    const section = el('mergerFilesSection');
    const configCard = el('mergerConfigCard');
    const countEl = el('mergerFileCount');

    if (!mergerFiles.length) {
      if (section) section.style.display = 'none';
      if (configCard) configCard.style.display = 'none';
      if (countEl) countEl.textContent = '0';
      return;
    }

    if (section) section.style.display = 'block';
    if (configCard) configCard.style.display = 'block';
    if (countEl) countEl.textContent = String(mergerFiles.length);

    container.replaceChildren();

    mergerFiles.forEach((f, idx) => {
      const item = document.createElement('div');
      item.className = 'merger-file-item';

      // Phân tích các category có trong file
      const catBadges = f.sheets.map(s => {
        const cat = detectAccountingCategory(s.name, s.grid);
        const rows = Math.max(0, s.grid.length - 1);
        return `<span class="merger-cat-badge merger-cat-${cat.id}">${cat.title} (${rows} dòng)</span>`;
      }).join(' ');

      const typeBadge = f.type === 'pdf'
        ? '<span class="merger-type-tag pdf">PDF</span>'
        : '<span class="merger-type-tag xlsx">XLSX</span>';

      item.innerHTML = `
        <div class="merger-file-order">${idx + 1}</div>
        <div class="merger-file-info">
          <div class="merger-file-title-row">
            ${typeBadge}
            <strong class="merger-file-name" title="${f.name}">${f.name}</strong>
            <span class="merger-file-size">${formatBytes(f.size)}</span>
          </div>
          <div class="merger-file-cats">${catBadges || '<span class="muted">Chưa có bảng</span>'}</div>
        </div>
        <div class="merger-file-actions">
          <button type="button" class="btn small" data-action="up" title="Di chuyển lên" ${idx === 0 ? 'disabled' : ''}>↑</button>
          <button type="button" class="btn small" data-action="down" title="Di chuyển xuống" ${idx === mergerFiles.length - 1 ? 'disabled' : ''}>↓</button>
          <button type="button" class="btn small outline danger" data-action="remove" title="Xóa file này">✕</button>
        </div>
      `;

      item.querySelector('[data-action="up"]')?.addEventListener('click', () => {
        if (idx > 0) {
          const temp = mergerFiles[idx];
          mergerFiles[idx] = mergerFiles[idx - 1];
          mergerFiles[idx - 1] = temp;
          renderFileList();
        }
      });

      item.querySelector('[data-action="down"]')?.addEventListener('click', () => {
        if (idx < mergerFiles.length - 1) {
          const temp = mergerFiles[idx];
          mergerFiles[idx] = mergerFiles[idx + 1];
          mergerFiles[idx + 1] = temp;
          renderFileList();
        }
      });

      item.querySelector('[data-action="remove"]')?.addEventListener('click', () => {
        mergerFiles.splice(idx, 1);
        renderFileList();
        updateMergerOptions();
        showToast('Đã xóa file khỏi danh sách gộp.');
      });

      container.appendChild(item);
    });
  }

  // Cập nhật các tùy chọn gộp (ví dụ kiểm tra PDF-to-PDF)
  function updateMergerOptions() {
    const allPdf = mergerFiles.length > 0 && mergerFiles.every(f => f.type === 'pdf');
    const pdfOpt = el('mergerPdfOption');
    const formatSelect = el('mergerFormatSelect');

    if (pdfOpt) {
      pdfOpt.disabled = !allPdf;
      if (!allPdf && formatSelect && formatSelect.value === 'pdf') {
        formatSelect.value = 'xlsx';
      }
    }
  }

  // Thực hiện gộp dữ liệu
  async function executeMerge() {
    if (!mergerFiles.length) {
      showToast('Hãy chọn ít nhất một file để gộp.');
      return;
    }

    const format = el('mergerFormatSelect')?.value || 'xlsx';
    const mode = el('mergerModeSelect')?.value || 'category';
    const skipDupHeaders = el('mergerSkipHeaders')?.checked ?? true;
    const skipTotals = el('mergerSkipTotals')?.checked ?? true;
    const reindexStt = el('mergerReindexStt')?.checked ?? true;
    const addFileName = el('mergerAddFileNameCol')?.checked ?? false;

    const processBtn = el('mergerProcessBtn');
    if (processBtn) {
      processBtn.disabled = true;
      processBtn.textContent = 'Đang xử lý gộp…';
    }

    try {
      if (format === 'pdf') {
        // === GHÉP PDF-TO-PDF ===
        const PDFLib = window.PDFLib;
        if (!PDFLib) throw new Error('Chưa tải được thư viện PDFLib để ghép PDF.');

        const mergedPdf = await PDFLib.PDFDocument.create();
        let totalPages = 0;

        for (const fileObj of mergerFiles) {
          const bytes = await fileObj.file.arrayBuffer();
          const doc = await PDFLib.PDFDocument.load(bytes);
          const count = doc.getPageCount();
          const copied = await mergedPdf.copyPages(doc, doc.getPageIndices());
          copied.forEach(page => mergedPdf.addPage(page));
          totalPages += count;
        }

        const mergedBytes = await mergedPdf.save();
        const blob = new Blob([mergedBytes], { type: 'application/pdf' });
        const filename = `Merged_Documents_${new Date().toISOString().slice(0, 10)}.pdf`;

        mergedResult = {
          type: 'pdf',
          blob,
          filename,
          sheets: [],
          pdfStats: { fileCount: mergerFiles.length, totalPages }
        };

        renderResultPreview();
        showToast(`✓ Đã ghép xong ${mergerFiles.length} file PDF (${totalPages} trang)!`);
      } else {
        // === GỘP SANG EXCEL (.XLSX) ===
        const resultSheets = [];

        if (mode === 'category') {
          // GOM THEO LOẠI BẢNG KÊ (Smart Accounting Split)
          const catGroups = {
            muavao: { title: 'Bảng kê mua vào (Tổng hợp)', sheets: [] },
            banra: { title: 'Bảng kê bán ra (Tổng hợp)', sheets: [] },
            baoco: { title: 'Giấy báo có (Tổng hợp)', sheets: [] },
            baono: { title: 'Giấy báo nợ (Tổng hợp)', sheets: [] },
            general: { title: 'Dữ liệu chung (Tổng hợp)', sheets: [] }
          };

          // Gom các sheet từ tất cả file vào các nhóm
          for (const f of mergerFiles) {
            for (const s of f.sheets) {
              const cat = detectAccountingCategory(s.name, s.grid);
              catGroups[cat.id].sheets.push({ fileName: f.name, sheetName: s.name, grid: s.grid });
            }
          }

          // Ghép từng nhóm thành 1 sheet kết quả
          for (const key of Object.keys(catGroups)) {
            const group = catGroups[key];
            if (!group.sheets.length) continue;

            const mergedGrid = [];
            let masterHeader = null;
            let sttCounter = 1;

            for (const item of group.sheets) {
              const grid = item.grid;
              if (!grid || !grid.length) continue;

              const headerIdx = findHeaderRowIndex(grid);
              const curHeader = grid[headerIdx] || [];

              if (!masterHeader) {
                // Tiêu đề của file đầu tiên
                masterHeader = [...curHeader];
                if (addFileName) {
                  mergedGrid.push(['Tên file nguồn', ...masterHeader]);
                } else {
                  mergedGrid.push([...masterHeader]);
                }
              }

              // Duyệt các dòng dữ liệu sau dòng tiêu đề
              for (let r = headerIdx + 1; r < grid.length; r++) {
                const row = grid[r];
                if (!row || !row.some(c => c !== '')) continue;

                // Bỏ dòng tiêu đề lặp lại giữa chừng
                if (skipDupHeaders && isHeaderRow(row, masterHeader)) continue;

                // Bỏ dòng tổng cộng giữa chừng
                if (skipTotals && isSummaryRow(row)) continue;

                const newRow = [...row];
                // Chuẩn hóa độ dài dòng theo masterHeader
                while (newRow.length < masterHeader.length) newRow.push('');

                // Đánh lại số thứ tự (STT) nếu cột 0 là STT
                if (reindexStt && (/stt/i.test(masterHeader[0] || '') || /^\d+$/.test(newRow[0] || ''))) {
                  newRow[0] = String(sttCounter++);
                }

                if (addFileName) {
                  mergedGrid.push([item.fileName, ...newRow]);
                } else {
                  mergedGrid.push(newRow);
                }
              }
            }

            if (mergedGrid.length > 1) {
              resultSheets.push({ name: group.title, grid: mergedGrid });
            }
          }
        } else if (mode === 'append') {
          // GỘP NỐI TIẾP THEO DÒNG (Toàn bộ vào 1 bảng duy nhất)
          const mergedGrid = [];
          let masterHeader = null;
          let sttCounter = 1;

          for (const f of mergerFiles) {
            for (const s of f.sheets) {
              const grid = s.grid;
              if (!grid || !grid.length) continue;

              const headerIdx = findHeaderRowIndex(grid);
              const curHeader = grid[headerIdx] || [];

              if (!masterHeader) {
                masterHeader = [...curHeader];
                if (addFileName) {
                  mergedGrid.push(['Tên file nguồn', ...masterHeader]);
                } else {
                  mergedGrid.push([...masterHeader]);
                }
              }

              for (let r = headerIdx + 1; r < grid.length; r++) {
                const row = grid[r];
                if (!row || !row.some(c => c !== '')) continue;
                if (skipDupHeaders && isHeaderRow(row, masterHeader)) continue;
                if (skipTotals && isSummaryRow(row)) continue;

                const newRow = [...row];
                while (newRow.length < masterHeader.length) newRow.push('');

                if (reindexStt && (/stt/i.test(masterHeader[0] || '') || /^\d+$/.test(newRow[0] || ''))) {
                  newRow[0] = String(sttCounter++);
                }

                if (addFileName) {
                  mergedGrid.push([f.name, ...newRow]);
                } else {
                  mergedGrid.push(newRow);
                }
              }
            }
          }

          if (mergedGrid.length > 1) {
            resultSheets.push({ name: 'Dữ liệu gộp (Tổng hợp)', grid: mergedGrid });
          }
        } else if (mode === 'sheets') {
          // MỖI FILE / BẢNG THÀNH 1 SHEET RIÊNG
          const usedSheetNames = new Set();
          for (const f of mergerFiles) {
            for (const s of f.sheets) {
              let sheetTitle = `${f.name.replace(/\.[^/.]+$/, '')} - ${s.name}`;
              sheetTitle = sheetTitle.replace(/[\\/*?:[\]]/g, '').slice(0, 31);
              let uniqueTitle = sheetTitle;
              let suffix = 1;
              while (usedSheetNames.has(uniqueTitle.toLowerCase())) {
                uniqueTitle = `${sheetTitle.slice(0, 28)} (${suffix++})`;
              }
              usedSheetNames.add(uniqueTitle.toLowerCase());
              resultSheets.push({ name: uniqueTitle, grid: s.grid });
            }
          }
        }

        if (!resultSheets.length) {
          throw new Error('Không có dòng dữ liệu hợp lệ nào để gộp.');
        }

        // Tạo workbook .xlsx
        const zipGenerator = window.efSheetsToXlsxZip;
        if (!zipGenerator) throw new Error('Hệ thống chưa nạp trình tạo zip Excel.');

        const zip = zipGenerator(resultSheets);
        const xlsxBytes = await zip.generateAsync({ type: 'uint8array' });
        const blob = new Blob([xlsxBytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const filename = `ExcelFlow_Merged_${new Date().toISOString().slice(0, 10)}.xlsx`;

        mergedResult = {
          type: 'xlsx',
          blob,
          filename,
          sheets: resultSheets
        };

        renderResultPreview();
        const totalRows = resultSheets.reduce((sum, s) => sum + Math.max(0, s.grid.length - 1), 0);
        showToast(`✓ Đã gộp thành công ${resultSheets.length} sheet (${totalRows} dòng dữ liệu)!`);
      }
    } catch (err) {
      console.error('Lỗi khi gộp file:', err);
      showToast('Lỗi khi gộp file: ' + err.message);
    } finally {
      if (processBtn) {
        processBtn.disabled = false;
        processBtn.textContent = '⌕ Xem trước dữ liệu gộp';
      }
    }
  }

  // Hiển thị kết quả xem trước
  function renderResultPreview() {
    const card = el('mergerResultCard');
    if (!card || !mergedResult) return;
    card.style.display = 'block';

    const titleEl = el('mergerResultTitle');
    const badgeEl = el('mergerResultBadge');
    const summaryEl = el('mergerResultSummary');
    const downloadBtn = el('mergerDownloadBtn');
    const sendBtn = el('mergerSendToImportBtn');
    const tableContainer = el('mergerTableContainer');
    const pdfPreview = el('mergerPdfPreview');
    const sheetTabs = el('mergerSheetTabs');

    if (mergedResult.type === 'pdf') {
      if (titleEl) titleEl.textContent = mergedResult.filename;
      if (badgeEl) {
        badgeEl.textContent = 'Tài liệu PDF (.pdf)';
        badgeEl.className = 'merger-result-badge pdf';
      }
      if (summaryEl) summaryEl.textContent = `Ghép từ ${mergedResult.pdfStats.fileCount} file PDF · Tổng cộng ${mergedResult.pdfStats.totalPages} trang.`;
      if (downloadBtn) {
        downloadBtn.textContent = '⇩ Tải file PDF (.pdf)';
        downloadBtn.style.display = 'inline-flex';
      }
      if (sendBtn) sendBtn.style.display = 'none';
      if (tableContainer) tableContainer.style.display = 'none';
      if (sheetTabs) sheetTabs.style.display = 'none';
      if (pdfPreview) {
        pdfPreview.style.display = 'block';
        const statsEl = el('mergerPdfStats');
        if (statsEl) statsEl.textContent = `Tổng cộng: ${mergedResult.pdfStats.fileCount} file, ${mergedResult.pdfStats.totalPages} trang. File PDF đã sẵn sàng tải về.`;
      }
    } else {
      // Excel Result
      if (titleEl) titleEl.textContent = mergedResult.filename;
      if (badgeEl) {
        badgeEl.textContent = 'Bảng tính Excel (.xlsx)';
        badgeEl.className = 'merger-result-badge xlsx';
      }
      const totalRows = mergedResult.sheets.reduce((sum, s) => sum + Math.max(0, s.grid.length - 1), 0);
      if (summaryEl) summaryEl.textContent = `Gộp từ ${mergerFiles.length} file · ${mergedResult.sheets.length} bảng tính · ${totalRows} dòng dữ liệu.`;
      if (downloadBtn) {
        downloadBtn.textContent = '⇩ Tải file Excel (.xlsx)';
        downloadBtn.style.display = 'inline-flex';
      }
      if (sendBtn) sendBtn.style.display = 'inline-flex';
      if (pdfPreview) pdfPreview.style.display = 'none';
      if (tableContainer) tableContainer.style.display = 'block';

      // Render Sheet Tabs
      activeSheetIndex = 0;
      currentPage = 1;
      searchFilter = '';
      if (el('mergerSearchInput')) el('mergerSearchInput').value = '';

      renderSheetTabs();
      renderActiveTable();
    }

    card.scrollIntoView({ behavior: 'smooth', block: 'start' });
  }

  // Render các tab sheet
  function renderSheetTabs() {
    const nav = el('mergerSheetTabs');
    if (!nav || !mergedResult || !mergedResult.sheets) return;

    if (mergedResult.sheets.length <= 1) {
      nav.style.display = 'none';
      return;
    }

    nav.style.display = 'flex';
    nav.replaceChildren();

    mergedResult.sheets.forEach((sh, idx) => {
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'merger-tab-btn' + (idx === activeSheetIndex ? ' active' : '');
      const dataRows = Math.max(0, sh.grid.length - 1);
      btn.innerHTML = `<strong>${sh.name}</strong> <small>(${dataRows} dòng)</small>`;
      btn.addEventListener('click', () => {
        activeSheetIndex = idx;
        currentPage = 1;
        renderSheetTabs();
        renderActiveTable();
      });
      nav.appendChild(btn);
    });
  }

  // Render bảng dữ liệu hiện tại
  function renderActiveTable() {
    const tableArea = el('mergerTableArea');
    if (!tableArea || !mergedResult || !mergedResult.sheets.length) return;

    const currentSheet = mergedResult.sheets[activeSheetIndex];
    if (!currentSheet) return;

    const grid = currentSheet.grid || [];
    const headerRow = grid[0] || [];
    let dataRows = grid.slice(1);

    // Tìm kiếm / lọc
    if (searchFilter) {
      const q = searchFilter.toLowerCase();
      dataRows = dataRows.filter(r => r.some(c => String(c || '').toLowerCase().includes(q)));
    }

    // Thống kê
    const titleEl = el('mergerActiveSheetTitle');
    const statsEl = el('mergerActiveSheetStats');
    if (titleEl) titleEl.textContent = currentSheet.name;
    if (statsEl) {
      const totalCol = headerRow.length;
      statsEl.textContent = `${dataRows.length} dòng dữ liệu · ${totalCol} cột` + (searchFilter ? ` (lọc từ ${grid.length - 1} dòng gốc)` : '');
    }

    // Phân trang
    const totalPages = Math.max(1, Math.ceil(dataRows.length / rowsPerPage));
    if (currentPage > totalPages) currentPage = totalPages;
    const startIdx = (currentPage - 1) * rowsPerPage;
    const pageRows = dataRows.slice(startIdx, startIdx + rowsPerPage);

    // Render HTML table
    const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
    let tableHtml = '<table class="data-table"><thead><tr>';
    tableHtml += '<th style="width: 50px; text-align: center;">#</th>';
    headerRow.forEach((col, cIdx) => {
      tableHtml += `<th>${esc(col || `Cột ${cIdx + 1}`)}</th>`;
    });
    tableHtml += '</tr></thead><tbody>';

    if (!pageRows.length) {
      tableHtml += `<tr><td colspan="${headerRow.length + 1}" style="text-align: center; color: #94a3b8; padding: 30px;">Không tìm thấy dữ liệu phù hợp.</td></tr>`;
    } else {
      pageRows.forEach((row, rIdx) => {
        const absRowIdx = startIdx + rIdx + 1;
        tableHtml += `<tr><td style="text-align: center; color: #94a3b8; font-weight: 600;">${absRowIdx}</td>`;
        for (let c = 0; c < headerRow.length; c++) {
          const val = row[c] ?? '';
          tableHtml += `<td>${esc(val)}</td>`;
        }
        tableHtml += '</tr>';
      });
    }

    tableHtml += '</tbody></table>';
    tableArea.innerHTML = tableHtml;

    // Cập nhật phân trang
    const pageNumEl = el('mergerPageNumber');
    const prevBtn = el('mergerPrevPage');
    const nextBtn = el('mergerNextPage');

    if (pageNumEl) pageNumEl.textContent = `Trang ${currentPage} / ${totalPages}`;
    if (prevBtn) prevBtn.disabled = currentPage <= 1;
    if (nextBtn) nextBtn.disabled = currentPage >= totalPages;
  }

  // Khởi tạo các sự kiện giao diện
  function init() {
    const dropZone = el('mergerDrop');
    const fileInput = el('mergerFileInput');
    const browseBtn = el('mergerBrowseBtn');
    const clearAllBtn = el('mergerClearAllBtn');
    const processBtn = el('mergerProcessBtn');
    const downloadBtn = el('mergerDownloadBtn');
    const sendBtn = el('mergerSendToImportBtn');
    const resetBtn = el('mergerResetBtn');
    const searchInput = el('mergerSearchInput');
    const prevBtn = el('mergerPrevPage');
    const nextBtn = el('mergerNextPage');

    if (browseBtn && fileInput) {
      browseBtn.addEventListener('click', () => fileInput.click());
    }

    if (fileInput) {
      fileInput.addEventListener('change', e => {
        if (e.target.files && e.target.files.length) {
          addFiles(Array.from(e.target.files));
          e.target.value = '';
        }
      });
    }

    if (dropZone) {
      ['dragenter', 'dragover'].forEach(eventName => {
        dropZone.addEventListener(eventName, e => {
          e.preventDefault();
          e.stopPropagation();
          dropZone.classList.add('is-dragging');
        });
      });

      ['dragleave', 'drop'].forEach(eventName => {
        dropZone.addEventListener(eventName, e => {
          e.preventDefault();
          e.stopPropagation();
          dropZone.classList.remove('is-dragging');
        });
      });

      dropZone.addEventListener('drop', e => {
        const dt = e.dataTransfer;
        if (dt && dt.files && dt.files.length) {
          addFiles(Array.from(dt.files));
        }
      });
    }

    if (clearAllBtn) {
      clearAllBtn.addEventListener('click', () => {
        mergerFiles = [];
        mergedResult = null;
        renderFileList();
        const resCard = el('mergerResultCard');
        if (resCard) resCard.style.display = 'none';
        showToast('Đã xóa toàn bộ file trong danh sách.');
      });
    }

    if (processBtn) {
      processBtn.addEventListener('click', executeMerge);
    }

    if (downloadBtn) {
      downloadBtn.addEventListener('click', () => {
        if (!mergedResult || !mergedResult.blob) return;
        if (typeof window.efDownload === 'function') {
          window.efDownload(mergedResult.blob, mergedResult.filename);
        } else {
          const url = URL.createObjectURL(mergedResult.blob);
          const a = document.createElement('a');
          a.href = url;
          a.download = mergedResult.filename;
          a.click();
          setTimeout(() => URL.revokeObjectURL(url), 20000);
        }
        showToast(`✓ Đã tải file gộp "${mergedResult.filename}"!`);
      });
    }

    if (sendBtn) {
      sendBtn.addEventListener('click', async () => {
        if (!mergedResult || !mergedResult.blob || mergedResult.type !== 'xlsx') {
          showToast('Chỉ file gộp dạng Excel mới đưa được vào Ghép & Nhập liệu.');
          return;
        }
        try {
          const virtualFile = new File([mergedResult.blob], mergedResult.filename, {
            type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            lastModified: Date.now()
          });

          if (typeof window.efAddSourceFiles === 'function') {
            await window.efAddSourceFiles([virtualFile]);
          }

          if (typeof window.efGoPage === 'function') {
            window.efGoPage('import');
          }

          showToast(`✓ Đã chuyển file gộp "${mergedResult.filename}" sang làm File nguồn!`);
        } catch (err) {
          console.error('Lỗi khi chuyển sang Nhập liệu:', err);
          showToast('Không thể chuyển file sang Nhập liệu: ' + err.message);
        }
      });
    }

    if (resetBtn) {
      resetBtn.addEventListener('click', () => {
        mergedResult = null;
        const resCard = el('mergerResultCard');
        if (resCard) resCard.style.display = 'none';
        window.scrollTo({ top: 0, behavior: 'smooth' });
      });
    }

    if (searchInput) {
      searchInput.addEventListener('input', e => {
        searchFilter = e.target.value.trim();
        currentPage = 1;
        renderActiveTable();
      });
    }

    if (prevBtn) {
      prevBtn.addEventListener('click', () => {
        if (currentPage > 1) {
          currentPage--;
          renderActiveTable();
        }
      });
    }

    if (nextBtn) {
      nextBtn.addEventListener('click', () => {
        currentPage++;
        renderActiveTable();
      });
    }
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
})();
