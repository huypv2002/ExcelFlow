/**
 * ExcelFlow Suite — Module Chuyển đổi PDF sang Excel (.xlsx)
 * Trích xuất bảng kê thuế GTGT (01-1, 01-2), sao kê ngân hàng & tài liệu PDF sang file Excel
 */
(function() {
  'use strict';

  const $ = id => document.getElementById(id);
  const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

  let currentFile = null;
  let currentZip = null;
  let currentSheets = [];
  let currentSheetIdx = 0;
  let currentFilter = '';
  let currentPage = 0;
  const PAGE_SIZE = 50;
  let isEditMode = false;
  let selectedRowIdx = -1;

  function init() {
    const dropZone = $('pdf2ExcelDrop');
    const fileInput = $('pdf2ExcelFileInput');
    if (!dropZone || !fileInput) return;

    // Click to select
    dropZone.addEventListener('click', e => {
      if (e.target !== fileInput && !fileInput.contains(e.target)) {
        fileInput.click();
      }
    });

    fileInput.addEventListener('change', () => {
      const files = Array.from(fileInput.files || []);
      if (files.length) handlePdfFile(files[0]);
    });

    // Drag & drop
    ['dragenter', 'dragover'].forEach(name => {
      dropZone.addEventListener(name, e => {
        e.preventDefault();
        dropZone.classList.add('is-dragging');
      });
    });

    ['dragleave', 'drop'].forEach(name => {
      dropZone.addEventListener(name, e => {
        e.preventDefault();
        dropZone.classList.remove('is-dragging');
      });
    });

    dropZone.addEventListener('drop', e => {
      const files = Array.from(e.dataTransfer?.files || []);
      const pdf = files.find(f => /\.pdf$/i.test(f.name));
      if (pdf) {
        handlePdfFile(pdf);
      } else {
        alert('Vui lòng chọn file định dạng .pdf');
      }
    });

    // Action buttons
    $('pdf2ExcelDownloadBtn')?.addEventListener('click', downloadConvertedExcel);
    $('pdf2ExcelSendToImportBtn')?.addEventListener('click', sendToImportWorkspace);
    $('pdf2ExcelResetBtn')?.addEventListener('click', resetView);
    $('pdf2ExcelEditBtn')?.addEventListener('click', toggleEditMode);
    $('pdf2ExcelAddRowBtn')?.addEventListener('click', addRow);
    $('pdf2ExcelDeleteRowBtn')?.addEventListener('click', deleteSelectedRow);

    // Search filter
    $('pdf2ExcelSearchInput')?.addEventListener('input', e => {
      currentFilter = (e.target.value || '').trim().toLowerCase();
      currentPage = 0;
      renderCurrentSheetTable();
    });

    // Pagination
    $('pdf2ExcelPrevPage')?.addEventListener('click', () => {
      if (currentPage > 0) {
        currentPage--;
        renderCurrentSheetTable();
      }
    });

    $('pdf2ExcelNextPage')?.addEventListener('click', () => {
      const maxPages = getMaxPages();
      if (currentPage < maxPages - 1) {
        currentPage++;
        renderCurrentSheetTable();
      }
    });
  }

  async function handlePdfFile(file) {
    if (!file || !/\.pdf$/i.test(file.name)) {
      alert('Chỉ hỗ trợ file định dạng PDF (.pdf)');
      return;
    }

    currentFile = file;
    const dropZone = $('pdf2ExcelDrop');
    const loadingBox = $('pdf2ExcelLoading');
    const resultArea = $('pdf2ExcelResultArea');
    const fileNameEl = $('pdf2ExcelFileName');

    dropZone.classList.add('is-loading');
    if (loadingBox) loadingBox.style.display = 'flex';
    if (fileNameEl) fileNameEl.textContent = file.name;
    if (resultArea) resultArea.style.display = 'none';

    try {
      if (typeof window.efParsePdfToWorkbook !== 'function') {
        throw new Error('Chưa nạp xong module trích xuất PDF. Hãy thử lại.');
      }

      const zip = await window.efParsePdfToWorkbook(file);
      currentZip = zip;
      currentSheets = zip._sheets || [];

      if (!currentSheets.length) {
        throw new Error('Không trích xuất được bảng dữ liệu từ file PDF này.');
      }

      currentSheetIdx = 0;
      currentPage = 0;
      currentFilter = '';
      if ($('pdf2ExcelSearchInput')) $('pdf2ExcelSearchInput').value = '';

      renderResultView();
      if (typeof window.efToast === 'function') {
        window.efToast('✓ Đã trích xuất thành công ' + currentSheets.length + ' sheet từ ' + file.name + '!');
      }
    } catch (err) {
      alert('Không thể chuyển đổi file PDF: ' + err.message);
      if (fileNameEl) fileNameEl.textContent = 'Lỗi: ' + err.message;
    } finally {
      dropZone.classList.remove('is-loading');
      if (loadingBox) loadingBox.style.display = 'none';
    }
  }

  function renderResultView() {
    const resultArea = $('pdf2ExcelResultArea');
    if (!resultArea) return;
    resultArea.style.display = 'block';

    const titleEl = $('pdf2ExcelDocTitle');
    const subtitleEl = $('pdf2ExcelDocSubtitle');
    const outName = currentFile.name.replace(/\.pdf$/i, '') + '.xlsx';

    if (titleEl) titleEl.textContent = outName;

    const totalRows = currentSheets.reduce((sum, s) => sum + Math.max(0, (s.grid?.length || 1) - 1), 0);
    if (subtitleEl) {
      subtitleEl.textContent = 'Trích xuất ' + currentSheets.length + ' sheet · ' + totalRows.toLocaleString('vi-VN') + ' dòng dữ liệu · Dung lượng gốc: ' + (currentFile.size / 1024).toFixed(0) + ' KB';
    }

    // Render Sheet Nav Tabs nếu có nhiều sheet
    const navEl = $('pdf2ExcelSheetNav');
    if (navEl) {
      if (currentSheets.length > 1) {
        navEl.style.display = 'flex';
        navEl.innerHTML = currentSheets.map((sh, idx) => {
          const rowCount = Math.max(0, (sh.grid?.length || 1) - 1);
          const activeClass = idx === currentSheetIdx ? 'active' : '';
          return '<button type="button" class="pdf2excel-tab ' + activeClass + '" data-sheet-idx="' + idx + '">' + esc(sh.name) + ' <span class="badge-count">' + rowCount + ' dòng</span></button>';
        }).join('');

        navEl.querySelectorAll('.pdf2excel-tab').forEach(btn => {
          btn.addEventListener('click', () => {
            currentSheetIdx = Number(btn.dataset.sheetIdx);
            currentPage = 0;
            currentFilter = '';
            if ($('pdf2ExcelSearchInput')) $('pdf2ExcelSearchInput').value = '';
            navEl.querySelectorAll('.pdf2excel-tab').forEach(b => b.classList.toggle('active', b === btn));
            renderCurrentSheetTable();
          });
        });
      } else {
        navEl.style.display = 'none';
      }
    }

    renderCurrentSheetTable();
  }

  function getFilteredRows() {
    const sheet = currentSheets[currentSheetIdx];
    if (!sheet || !sheet.grid || sheet.grid.length <= 1) return { header: [], rows: [] };
    const header = sheet.grid[0] || [];
    let rows = sheet.grid.slice(1);

    if (currentFilter) {
      rows = rows.filter(r => r.some(c => String(c ?? '').toLowerCase().includes(currentFilter)));
    }
    return { header, rows };
  }

  function getMaxPages() {
    const { rows } = getFilteredRows();
    return Math.max(1, Math.ceil(rows.length / PAGE_SIZE));
  }

  function toggleEditMode() {
    isEditMode = !isEditMode;
    selectedRowIdx = -1;
    const btn = $('pdf2ExcelEditBtn');
    const toolbar = $('pdf2ExcelEditToolbar');

    if (isEditMode) {
      if (btn) {
        btn.textContent = '✓ Hoàn tất chỉnh sửa';
        btn.className = 'btn teal is-editing';
      }
      if (toolbar) toolbar.style.display = 'flex';
      if (typeof window.efToast === 'function') {
        window.efToast('Đã bật chế độ sửa: Bạn có thể bấm trực tiếp vào ô để sửa nội dung.');
      }
    } else {
      if (btn) {
        btn.textContent = '✎ Chỉnh sửa dữ liệu';
        btn.className = 'btn outline';
      }
      if (toolbar) toolbar.style.display = 'none';
      if (typeof window.efToast === 'function') {
        window.efToast('Đã lưu các sửa đổi vào bảng.');
      }
    }
    renderCurrentSheetTable();
  }

  function addRow() {
    const sheet = currentSheets[currentSheetIdx];
    if (!sheet) return;
    const header = sheet.grid[0] || [];
    const newRow = new Array(header.length).fill('');
    if (selectedRowIdx > 0 && selectedRowIdx < sheet.grid.length) {
      sheet.grid.splice(selectedRowIdx + 1, 0, newRow);
      selectedRowIdx++;
    } else {
      sheet.grid.push(newRow);
      selectedRowIdx = sheet.grid.length - 1;
    }
    renderCurrentSheetTable();
    if (typeof window.efToast === 'function') window.efToast('✓ Đã thêm 1 dòng mới vào bảng.');
  }

  function deleteSelectedRow() {
    const sheet = currentSheets[currentSheetIdx];
    if (!sheet) return;
    if (selectedRowIdx <= 0 || selectedRowIdx >= sheet.grid.length) {
      if (typeof window.efToast === 'function') window.efToast('Hãy bấm chọn một dòng trong bảng trước khi xóa.');
      else alert('Hãy bấm chọn một dòng trong bảng trước khi xóa.');
      return;
    }
    sheet.grid.splice(selectedRowIdx, 1);
    selectedRowIdx = -1;
    renderCurrentSheetTable();
    if (typeof window.efToast === 'function') window.efToast('✓ Đã xóa dòng được chọn.');
  }

  function renderCurrentSheetTable() {
    const sheet = currentSheets[currentSheetIdx];
    if (!sheet) return;

    const nameEl = $('pdf2ExcelCurrentSheetName');
    const statsEl = $('pdf2ExcelCurrentSheetStats');
    const tableArea = $('pdf2ExcelTableArea');
    const pageNumEl = $('pdf2ExcelPageNumber');
    const prevBtn = $('pdf2ExcelPrevPage');
    const nextBtn = $('pdf2ExcelNextPage');

    const { header, rows } = getFilteredRows();
    const totalDataRows = Math.max(0, (sheet.grid?.length || 1) - 1);

    if (nameEl) nameEl.textContent = sheet.name;
    if (statsEl) {
      statsEl.textContent = currentFilter
        ? 'Tìm thấy ' + rows.length + ' / ' + totalDataRows + ' dòng (Khớp từ khóa "' + currentFilter + '")'
        : totalDataRows + ' dòng dữ liệu · ' + header.length + ' cột';
    }

    const maxPages = getMaxPages();
    if (currentPage >= maxPages) currentPage = maxPages - 1;
    if (currentPage < 0) currentPage = 0;

    if (pageNumEl) pageNumEl.textContent = 'Trang ' + (currentPage + 1) + ' / ' + maxPages;
    if (prevBtn) prevBtn.disabled = currentPage === 0;
    if (nextBtn) nextBtn.disabled = currentPage >= maxPages - 1;

    if (!rows.length) {
      tableArea.innerHTML = '<div class="empty">Không tìm thấy dòng nào phù hợp với bộ lọc.</div>';
      return;
    }

    const startIdx = currentPage * PAGE_SIZE;
    const pageRows = rows.slice(startIdx, startIdx + PAGE_SIZE);

    const thead = '<thead><tr><th style="width:50px">#</th>' + header.map((h, i) => '<th>' + esc(h || 'Cột ' + (i + 1)) + '</th>').join('') + '</tr></thead>';
    const tbody = '<tbody>' + pageRows.map((r, rIdx) => {
      const absRowIdx = startIdx + rIdx + 1;
      const isSel = isEditMode && (absRowIdx === selectedRowIdx);
      const cells = header.map((_, cIdx) => {
        const val = esc(r[cIdx] || '');
        if (isEditMode) {
          return `<td contenteditable="true" data-row="${absRowIdx}" data-col="${cIdx}" title="Bấm để chỉnh sửa">${val}</td>`;
        }
        return `<td>${val}</td>`;
      }).join('');
      return `<tr data-row="${absRowIdx}" class="${isSel ? 'is-selected' : ''}"><th style="color:#64748b">${absRowIdx}</th>${cells}</tr>`;
    }).join('') + '</tbody>';

    tableArea.innerHTML = '<table class="data-table">' + thead + tbody + '</table>';

    if (isEditMode) {
      tableArea.querySelectorAll('td[contenteditable="true"]').forEach(td => {
        td.addEventListener('blur', () => {
          const r = Number(td.dataset.row);
          const c = Number(td.dataset.col);
          if (sheet.grid[r]) {
            sheet.grid[r][c] = td.textContent.trim();
          }
        });
      });

      tableArea.querySelectorAll('tbody tr').forEach(tr => {
        tr.addEventListener('click', () => {
          tableArea.querySelectorAll('tbody tr').forEach(t => t.classList.remove('is-selected'));
          tr.classList.add('is-selected');
          selectedRowIdx = Number(tr.dataset.row);
        });
      });
    }
  }

  async function generateCurrentBlob() {
    if (!currentSheets.length) return null;
    const zipGen = window.efSheetsToXlsxZip;
    const zip = zipGen ? zipGen(currentSheets) : currentZip;
    return await zip.generateAsync({
      type: 'blob',
      mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    });
  }

  async function downloadConvertedExcel() {
    if (!currentFile || !currentSheets.length) return;
    try {
      const blob = await generateCurrentBlob();
      const outName = currentFile.name.replace(/\.pdf$/i, '') + '.xlsx';
      if (typeof window.efDownload === 'function') {
        window.efDownload(blob, outName);
      }
      if (typeof window.efToast === 'function') {
        window.efToast('✓ Đã tải file Excel: ' + outName);
      }
    } catch (err) {
      alert('Không thể tạo file Excel: ' + err.message);
    }
  }

  async function sendToImportWorkspace() {
    if (!currentFile || !currentSheets.length) return;
    try {
      const blob = await generateCurrentBlob();
      const excelName = currentFile.name.replace(/\.pdf$/i, '') + '.xlsx';
      const excelFile = new File([blob], excelName, {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      });

      if (typeof window.efAddSourceFiles === 'function') {
        await window.efAddSourceFiles([excelFile]);
      }
      if (typeof window.efGoPage === 'function') {
        window.efGoPage('import');
      }
      if (typeof window.efToast === 'function') {
        window.efToast('✓ Đã đưa ' + excelName + ' vào danh sách file nguồn.');
      }
    } catch (err) {
      alert('Không thể chuyển file sang Nhập liệu: ' + err.message);
    }
  }

  function resetView() {
    currentFile = null;
    currentZip = null;
    currentSheets = [];
    currentSheetIdx = 0;
    currentFilter = '';
    currentPage = 0;
    isEditMode = false;
    selectedRowIdx = -1;

    const fileInput = $('pdf2ExcelFileInput');
    if (fileInput) fileInput.value = '';
    const fileNameEl = $('pdf2ExcelFileName');
    if (fileNameEl) fileNameEl.textContent = 'Chưa chọn file PDF';
    const resultArea = $('pdf2ExcelResultArea');
    if (resultArea) resultArea.style.display = 'none';
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
})();
