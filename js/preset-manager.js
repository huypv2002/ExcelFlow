/**
 * ExcelFlow Suite — Quản lý Bộ Mẫu Ghép Cột & Bộ Nhớ Tự Động Ghép Cột
 * Hỗ trợ: Thêm, Xóa, Sửa, Search, Lưu lịch sử, Auto-Map tự động
 */
(function() {
  'use strict';

  const $ = id => document.getElementById(id);
  const esc = v => String(v ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

  const PRESETS_STORAGE_KEY = 'excelflow_presets_v3';
  const MEMORY_STORAGE_KEY = 'excelflow_column_memory_v3';

  // Danh sách từ khóa ngân hàng & hóa đơn thuế phổ biến để auto-seed gợi ý ghép cột
  const BUILTIN_MAPPINGS = [
    // Bảng kê thuế GTGT (01-1 bán ra, 01-2 mua vào) & Sổ BKMV / BKBR
    { src: ['ký hiệu hóa đơn', 'ky hieu hoa don', 'ký hiệu hd', 'ky hieu hd', 'ký hiệu', 'serial', 'seri'], dst: ['Serial', 'Ký hiệu hóa đơn', 'Ký hiệu HĐ'], type: 'text' },
    { src: ['số hoá đơn', 'số hóa đơn', 'so hoa don', 'số hđ', 'so hd', 'inv no', 'invoice no'], dst: ['Số', 'Số hoá đơn', 'Số hóa đơn', 'Số HĐ'], type: 'text' },
    { src: ['ngày phát hành', 'ngay phat hanh', 'ngày lập', 'ngay lap', 'ngày hđ', 'ngày hóa đơn', 'ngay hd', 'ngay hoa don', 'inv date', 'invoice date'], dst: ['Ngày hóa đơn', 'Ngày phát hành', 'Ngày CT (*)', 'Ngày'], type: 'date' },
    { src: ['tên người bán', 'ten nguoi ban', 'người bán', 'tên người mua', 'ten nguoi mua', 'người mua', 'tên khách hàng', 'ten khach hang'], dst: ['Tên khách hàng', 'Tên người bán', 'Tên người mua', 'Khách hàng', 'Tên đối tượng'], type: 'text' },
    { src: ['mã số thuế người bán', 'ma so thue nguoi ban', 'mã số thuế người mua', 'ma so thue nguoi mua', 'mã số thuế', 'ma so thue', 'mst', 'mst người bán', 'mst người mua'], dst: ['MST', 'Mã số thuế', 'Mã số thuế người bán', 'Mã số thuế người mua'], type: 'text' },
    { src: ['tiền chưa thuế', 'tien chua thue', 'doanh số bán chưa có thuế', 'doanh số', 'tiền hàng', 'tien hang'], dst: ['Tiền hàng', 'Doanh số bán chưa có thuế', 'Tiền chưa thuế'], type: 'money' },
    { src: ['thuế gtgt', 'thue gtgt', 'tiền thuế gtgt', 'tiền thuế'], dst: ['Thuế', 'Thuế GTGT', 'Tiền thuế'], type: 'money' },
    { src: ['tổng tiền hàng', 'tong tien hang', 'tổng thanh toán', 'tong thanh toan', 'tổng tiền', 'tong tien'], dst: ['Tổng tiền hàng', 'Tổng tiền', 'Tổng thanh toán'], type: 'money' },

    // Sao kê ngân hàng phổ biến (Vietcombank, Techcombank, MB Bank, ACB, BIDV, VietinBank...)
    { src: ['ngày giao dịch', 'ngay giao dich', 'ngày giao dịch/transaction date', 'transaction date', 'ngày gd', 'ngay gd', 'ngày ct', 'ngày chứng từ', 'trans date'], dst: ['Ngày', 'Ngày CT (*)', 'Ngày giao dịch', 'Ngày hạch toán'], type: 'date' },
    { src: ['số tiền ghi có', 'so tien ghi co', 'phát sinh có', 'phat sinh co', 'có/credit', 'ghi có', 'credit', 'tiền vào', 'số tiền thu'], dst: ['Số tiền thu', 'Tiền (*)', 'Phát sinh có', 'Số tiền ghi có'], type: 'money' },
    { src: ['số tiền ghi nợ', 'so tien ghi no', 'phát sinh nợ', 'phat sinh no', 'nợ/debit', 'ghi nợ', 'debit', 'tiền ra', 'số tiền chi'], dst: ['Số tiền chi', 'Tiền (*)', 'Phát sinh nợ', 'Số tiền ghi nợ'], type: 'money' },
    { src: ['nội dung chi tiết', 'noi dung chi tiet', 'diễn giải/description', 'diễn giải', 'dien giai', 'nội dung', 'noi dung', 'nội dung gd', 'nội dung giao dịch', 'transactions in detail', 'description', 'remarks'], dst: ['Diễn giải', 'Nội dung', 'Nội dung chi tiết'], type: 'text' },
    { src: ['số ct', 'so ct', 'số chứng từ', 'so chung tu', 'số bút toán/reference number', 'số bút toán', 'so but toan', 'số tham chiếu', 'mã gd', 'mã giao dịch', 'reference number', 'doc no', 'ref no', 'trans no'], dst: ['Số tham chiếu', 'Số chứng từ (*)', 'Số CT', 'Số bút toán'], type: 'text' },
    { src: ['tên tài khoản đối ứng', 'ten tai khoan doi ung', 'remitter\'s account name', 'đơn vị thụ hưởng/đơn vị chuyển', 'tên đối tác', 'tên đối tượng'], dst: ['Tên đối tượng', 'Mã đối tượng (*)', 'Khách hàng'], type: 'text' },
    { src: ['stt', 'số thứ tự', 'so thu tu', 'order'], dst: ['Số chứng từ (*)', 'STT'], type: 'text' },
    { src: ['mã đối tượng', 'ma doi tuong', 'đối tượng'], dst: ['Mã đối tượng (*)', 'Mã đối tượng'], type: 'text' },
    { src: ['tài khoản nợ', 'tk nợ', 'tk no'], dst: ['Tk nợ'], type: 'text' },
    { src: ['tài khoản có', 'tk có', 'tk co'], dst: ['Tk có'], type: 'text' }
  ];

  function toast(msg) {
    if (typeof window.efToast === 'function') window.efToast(msg);
    else {
      const t = $('toast');
      if (t) {
        t.textContent = msg;
        t.classList.add('show');
        clearTimeout(toast.timer);
        toast.timer = setTimeout(() => t.classList.remove('show'), 2800);
      }
    }
  }

  // =========================================================================
  // 1. Quản lý Bộ Nhớ Ghép Cột (Column Mapping Memory)
  // =========================================================================
  function getMemory() {
    try {
      return JSON.parse(localStorage.getItem(MEMORY_STORAGE_KEY) || '{}');
    } catch {
      return {};
    }
  }

  function saveMappingToMemory(srcName, dstName, dataType) {
    if (!srcName || !dstName) return;
    const mem = getMemory();
    const cleanSrc = srcName.trim().toLowerCase();
    mem[cleanSrc] = {
      dst: dstName.trim(),
      type: dataType || 'auto',
      updatedAt: Date.now()
    };
    try {
      localStorage.setItem(MEMORY_STORAGE_KEY, JSON.stringify(mem));
    } catch {}
  }

  // =========================================================================
  // 2. Quản lý Danh Sách Bộ Mẫu (Presets Storage)
  // =========================================================================
  function seedDefaultPresetsIfEmpty() {
    const existing = getPresets();
    if (existing.length > 0) return;

    const defaults = [
      {
        id: 'preset_vcb_pdf_asp_default',
        name: 'Sao kê Vietcombank (PDF) → Mẫu Sao kê Ngân hàng (ASP/MISA)',
        description: 'Ghép chuẩn PDF Sao kê Vietcombank vào file mẫu nhập khẩu Sao_ke_ngan_hang.xlsx (Ngày, Số tiền thu, Số tiền chi, Diễn giải, Số tham chiếu)',
        updatedAt: Date.now(),
        mapping: 'B -> A\nE -> B\nD -> C\nG -> D\nC -> E',
        formats: 'A -> date\nB -> money\nC -> money\nD -> text\nE -> text',
        reverse: false,
        skipTotals: true,
        skipZeros: true
      },
      {
        id: 'preset_tcb_asp_default',
        name: 'Sao kê Techcombank (Excel) → Mẫu Sao kê Ngân hàng (ASP/MISA)',
        description: 'Ghép chuẩn Sao kê Techcombank vào file mẫu Sao_ke_ngan_hang.xlsx (Ngày giao dịch, Có/Credit, Nợ/Debit, Diễn giải, Số bút toán, Tên đối ứng)',
        updatedAt: Date.now() - 10000,
        mapping: 'B -> A\nI -> B\nH -> C\nG -> D\nC -> E\nF -> F',
        formats: 'A -> date\nB -> money\nC -> money\nD -> text\nE -> text\nF -> text',
        reverse: true,
        skipTotals: true,
        skipZeros: true
      },
      {
        id: 'preset_tax_invoice_default',
        name: 'Tờ khai Thuế GTGT (PDF) → Bảng kê Mua vào BKMV',
        description: 'Mẫu chuẩn trích xuất PDF Tờ khai thuế GTGT (01-2) điền vào Sổ BKMV (Serial, Số, Ngày, Tên KH, MST, Tiền hàng, Thuế)',
        updatedAt: Date.now() - 20000,
        mapping: 'C -> A\nD -> B\nE -> C\nF -> D\nG -> E\nI -> F\nK -> G',
        formats: 'C -> date\nF -> money\nG -> money',
        reverse: false,
        skipTotals: true,
        skipZeros: true
      },
      {
        id: 'preset_vcb_default',
        name: 'Sao kê Vietcombank → Báo có Freebird',
        description: 'Mẫu chuẩn sao kê ngân hàng VCB ghép vào Báo có (STT, Ngày CT, Tiền, Diễn giải)',
        updatedAt: Date.now() - 86400000,
        mapping: 'A -> A\nB -> B\nD -> R\nE -> V',
        formats: 'B -> date\nR -> money',
        reverse: false,
        skipTotals: true,
        skipZeros: true
      },
      {
        id: 'preset_mbbank_default',
        name: 'Sao kê MB Bank / Techcombank',
        description: 'Ghép cột theo Ngày giao dịch, Số tiền phát sinh có và Diễn giải thanh toán',
        updatedAt: Date.now() - 43200000,
        mapping: 'B -> B\nC -> R\nD -> V',
        formats: 'B -> date\nR -> money',
        reverse: false,
        skipTotals: true,
        skipZeros: true
      }
    ];

    try {
      localStorage.setItem(PRESETS_STORAGE_KEY, JSON.stringify(defaults));
    } catch {}
  }

  function getPresets() {
    try {
      const data = JSON.parse(localStorage.getItem(PRESETS_STORAGE_KEY) || '[]');
      if (Array.isArray(data)) return data;
      // Khôi phục nếu trước đó lưu dưới dạng object
      return Object.keys(data).map(key => ({
        id: 'preset_' + key.replace(/\s+/g, '_'),
        name: key,
        description: '',
        updatedAt: Date.now(),
        ...data[key]
      }));
    } catch {
      return [];
    }
  }

  function savePresets(presets) {
    try {
      localStorage.setItem(PRESETS_STORAGE_KEY, JSON.stringify(presets));
      updatePresetDropdown();
    } catch (e) {
      alert('Không thể lưu bộ mẫu vào bộ nhớ trình duyệt: ' + e.message);
    }
  }

  function addOrUpdatePreset(item) {
    const presets = getPresets();
    const idx = presets.findIndex(p => p.id === item.id);
    if (idx >= 0) {
      presets[idx] = { ...presets[idx], ...item, updatedAt: Date.now() };
    } else {
      presets.unshift({
        id: item.id || ('preset_' + Date.now()),
        name: item.name || 'Bộ mẫu mới',
        description: item.description || '',
        mapping: item.mapping || '',
        formats: item.formats || '',
        reverse: !!item.reverse,
        skipTotals: item.skipTotals !== false,
        skipZeros: item.skipZeros !== false,
        updatedAt: Date.now()
      });
    }
    savePresets(presets);
  }

  function deletePresetById(id) {
    let presets = getPresets();
    presets = presets.filter(p => p.id !== id);
    savePresets(presets);
  }

  // =========================================================================
  // 3. Tự Động Khớp Cột (Auto-Map) Dựa Trên Bộ Nhớ & Tiêu Đề
  // =========================================================================
  function autoMapColumns() {
    const srcSel = $('addSource');
    const dstSel = $('addTarget');
    const mappingEl = $('mapping');
    const formatsEl = $('formatRules');
    if (!srcSel || !dstSel || !mappingEl) return;

    // Chỉ tự động điền nếu ô quy tắc hiện tại đang trống
    if (mappingEl.value.trim().length > 0) return;

    const srcOptions = Array.from(srcSel.options).filter(o => o.value);
    const dstOptions = Array.from(dstSel.options).filter(o => o.value);
    if (!srcOptions.length || !dstOptions.length) return;

    const memory = getMemory();
    const cleanHeader = str => (str || '').replace(/^[A-Z]+(?:[–-][A-Z]+)?\s*[·.:\-]\s*/, '').trim().toLowerCase();
    const colLetterFromOpt = opt => (opt.textContent.match(/^[A-Z]+/i) || [''])[0].toUpperCase();

    const usedDst = new Set();
    const newMappings = [];
    const newFormats = [];

    // Duyệt qua từng cột nguồn để tìm cột đích phù hợp nhất
    for (const sOpt of srcOptions) {
      const sName = cleanHeader(sOpt.textContent);
      const sLetter = colLetterFromOpt(sOpt);
      if (!sName || !sLetter) continue;

      let matchedDst = null;
      let matchedType = 'auto';

      // 1. Kiểm tra trong bộ nhớ đã lưu
      if (memory[sName]) {
        const memDst = memory[sName].dst.toLowerCase();
        matchedDst = dstOptions.find(d => !usedDst.has(d.value) && cleanHeader(d.textContent) === memDst);
        if (matchedDst) matchedType = memory[sName].type || 'auto';
      }

      // 2. Khớp theo tên tiêu đề giống hệt nhau
      if (!matchedDst) {
        matchedDst = dstOptions.find(d => !usedDst.has(d.value) && cleanHeader(d.textContent) === sName);
      }

      // 3. Khớp theo danh sách từ khóa ngân hàng & hóa đơn thuế chuẩn
      if (!matchedDst) {
        const matchText = (header, term) => {
          if (!header || !term) return false;
          if (header === term) return true;
          if (term.length <= 4 || header.length <= 4) return false;
          return header.includes(term) || term.includes(header);
        };
        for (const rule of BUILTIN_MAPPINGS) {
          const matchSrc = rule.src.some(term => matchText(sName, term));
          if (matchSrc) {
            const dstTargets = Array.isArray(rule.dst) ? rule.dst : [rule.dst];
            // Ưu tiên khớp chính xác tên cột đích trước, sau đó mới khớp chuỗi con
            for (const target of dstTargets) {
              const tClean = cleanHeader(target);
              matchedDst = dstOptions.find(d => !usedDst.has(d.value) && cleanHeader(d.textContent) === tClean);
              if (matchedDst) break;
            }
            if (!matchedDst) {
              for (const target of dstTargets) {
                const tClean = cleanHeader(target);
                matchedDst = dstOptions.find(d => !usedDst.has(d.value) && matchText(cleanHeader(d.textContent), tClean));
                if (matchedDst) break;
              }
            }
            if (matchedDst) {
              matchedType = rule.type;
              break;
            }
          }
        }
      }

      if (matchedDst) {
        const dLetter = colLetterFromOpt(matchedDst);
        if (dLetter) {
          usedDst.add(matchedDst.value);
          newMappings.push(`${sLetter} -> ${dLetter}`);
          if (matchedType && matchedType !== 'auto') {
            newFormats.push(`${dLetter} -> ${matchedType}`);
          }
          saveMappingToMemory(sName, cleanHeader(matchedDst.textContent), matchedType);
        }
      }
    }

    if (newMappings.length > 0) {
      mappingEl.value = newMappings.join('\n');
      if (newFormats.length > 0 && formatsEl) {
        formatsEl.value = newFormats.join('\n');
        formatsEl.dispatchEvent(new Event('input', { bubbles: true }));
      }
      mappingEl.dispatchEvent(new Event('input', { bubbles: true }));
      toast(`Đã tự động khớp ${newMappings.length} cặp cột từ bộ nhớ đã học!`);
    }
  }
  window.efAutoMapColumns = autoMapColumns;

  // =========================================================================
  // 4. Áp Dụng Bộ Mẫu Vào Không Gian Làm Việc
  // =========================================================================
  function applyPreset(preset) {
    if (!preset) return;
    const mappingEl = $('mapping');
    const formatsEl = $('formatRules');
    const reverseEl = $('reverse');
    const skipTotalsEl = $('skipTotals');
    const skipZerosEl = $('skipZeros');

    if (mappingEl) mappingEl.value = preset.mapping || '';
    if (formatsEl) formatsEl.value = preset.formats || '';
    if (reverseEl && typeof preset.reverse !== 'undefined') reverseEl.checked = !!preset.reverse;
    if (skipTotalsEl && typeof preset.skipTotals !== 'undefined') skipTotalsEl.checked = !!preset.skipTotals;
    if (skipZerosEl && typeof preset.skipZeros !== 'undefined') skipZerosEl.checked = !!preset.skipZeros;

    if (mappingEl) mappingEl.dispatchEvent(new Event('input', { bubbles: true }));
    if (formatsEl) formatsEl.dispatchEvent(new Event('input', { bubbles: true }));

    // Ghi nhớ các cặp cột trong bộ mẫu vào memory
    if (preset.mapping) {
      const lines = preset.mapping.split('\n');
      for (const line of lines) {
        const parts = line.split(/(?:->|=>|→|=)/);
        if (parts.length === 2) {
          saveMappingToMemory(parts[0].trim(), parts[1].trim(), 'auto');
        }
      }
    }

    toast(`✓ Đã áp dụng bộ mẫu "${preset.name}"!`);
  }

  // =========================================================================
  // 5. Cập Nhật Dropdown Chọn Nhanh Bộ Mẫu Trên Màn Hình
  // =========================================================================
  function updatePresetDropdown() {
    const sel = $('efTemplateSelect');
    if (!sel) return;
    const presets = getPresets();
    const curVal = sel.value;

    sel.innerHTML = '<option value="">-- Chọn bộ mẫu để áp dụng --</option>' +
      presets.map(p => `<option value="${esc(p.id)}">${esc(p.name)}</option>`).join('');

    if (curVal && presets.some(p => p.id === curVal)) {
      sel.value = curVal;
    }
  }

  // =========================================================================
  // 6. UI Modal Quản Lý Bộ Mẫu (Search, Thêm, Xóa, Sửa)
  // =========================================================================
  let editingPresetId = null;

  function ensurePresetModalUI() {
    if (!$('efPresetModalBackdrop')) {
      const modalWrap = document.createElement('div');
      modalWrap.id = 'efPresetModalBackdrop';
      modalWrap.className = 'ef-preview-modal-backdrop';
      modalWrap.innerHTML = `
        <div class="ef-preview-modal ef-preset-modal-box" id="efPresetModal" role="dialog" aria-modal="true" aria-labelledby="efPresetModalTitle">
          <div class="ef-modal-header">
            <div class="ef-modal-title">
              <div class="ef-icon">▤</div>
              <div>
                <h3 id="efPresetModalTitle">Quản lý các bộ mẫu ghép cột</h3>
                <p>Lưu trữ, tìm kiếm và tái sử dụng các quy tắc ghép cột cho từng loại sao kê / đối tác.</p>
              </div>
            </div>
            <div class="ef-modal-actions">
              <button type="button" class="btn small primary" id="efPresetAddNewBtn">+ Thêm bộ mới</button>
              <button type="button" id="efPresetModalClose">Đóng</button>
            </div>
          </div>
          
          <div class="ef-modal-toolbar ef-preset-toolbar">
            <div class="field" style="flex:1">
              <label for="efPresetSearch">Tìm kiếm bộ mẫu</label>
              <input type="text" id="efPresetSearch" placeholder="Tìm theo tên bộ, mô tả, cặp cột ghép...">
            </div>
            <div class="ef-preset-stats-chips">
              <span class="chip">Tổng: <strong id="efPresetTotalCount">0</strong></span>
              <span class="chip">Hiển thị: <strong id="efPresetShownCount">0</strong></span>
            </div>
          </div>

          <div class="ef-modal-body ef-preset-modal-body">
            <div class="ef-preset-card-grid" id="efPresetCardGrid"></div>
          </div>

          <div class="ef-modal-footer">
            <span>Dữ liệu được lưu an toàn tại trình duyệt của bạn.</span>
            <strong>ExcelFlow Presets</strong>
          </div>
        </div>`;
      document.body.appendChild(modalWrap);
    }

    if (!$('efPresetEditBackdrop')) {
      const editWrap = document.createElement('div');
      editWrap.id = 'efPresetEditBackdrop';
      editWrap.className = 'ef-editor-backdrop';
      editWrap.innerHTML = `
        <div class="ef-editor ef-preset-editor-box" role="dialog" aria-modal="true" aria-labelledby="efPresetEditTitle">
          <h4 id="efPresetEditTitle">Thêm bộ mẫu mới</h4>
          <div class="field full">
            <label for="efPresetEditName">Tên bộ mẫu <span style="color:#e11d48">*</span></label>
            <input type="text" id="efPresetEditName" placeholder="Ví dụ: Sao kê Vietcombank → Báo có Freebird">
          </div>
          <div class="field full">
            <label for="efPresetEditDesc">Mô tả / Ghi chú</label>
            <input type="text" id="efPresetEditDesc" placeholder="Ví dụ: Ghép các cột STT, Ngày, Tiền và Diễn giải">
          </div>
          <div class="field full">
            <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:6px">
              <label for="efPresetEditMapping" style="margin:0">Quy tắc ghép cột (Nguồn → Đích)</label>
              <button type="button" class="btn small" id="efPresetGetFromWorkspace" style="font-size:11px;padding:3px 9px">Lấy từ bảng hiện tại</button>
            </div>
            <textarea id="efPresetEditMapping" placeholder="A -> A&#10;B -> B&#10;D -> R&#10;E -> V" rows="5"></textarea>
          </div>
          <div class="field full">
            <label for="efPresetEditFormats">Định dạng kiểu dữ liệu (tùy chọn)</label>
            <textarea id="efPresetEditFormats" placeholder="B -> date&#10;R -> money" rows="2"></textarea>
          </div>
          <div class="ef-editor-actions">
            <button type="button" id="efPresetEditCancel">Hủy</button>
            <button type="button" class="primary" id="efPresetEditSave">Lưu bộ mẫu</button>
          </div>
        </div>`;
      document.body.appendChild(editWrap);
    }
  }

  function renderPresetCards() {
    const grid = $('efPresetCardGrid');
    if (!grid) return;
    const query = ($('efPresetSearch')?.value || '').trim().toLowerCase();
    const allPresets = getPresets();

    const filtered = allPresets.filter(p => {
      if (!query) return true;
      const hay = [p.name, p.description, p.mapping, p.formats].join(' ').toLowerCase();
      return hay.includes(query);
    });

    $('efPresetTotalCount').textContent = String(allPresets.length);
    $('efPresetShownCount').textContent = String(filtered.length);

    if (!filtered.length) {
      grid.innerHTML = `
        <div class="ef-preset-empty">
          <div style="font-size:32px;margin-bottom:8px">▤</div>
          <strong>Không tìm thấy bộ mẫu nào phù hợp.</strong>
          <p style="color:#64748b;font-size:12px;margin-top:4px">Bạn có thể tạo bộ mới bằng nút "+ Thêm bộ mới" phía trên.</p>
        </div>`;
      return;
    }

    grid.innerHTML = filtered.map(p => {
      const mappingLines = (p.mapping || '').split('\n').filter(l => l.trim() && !l.trim().startsWith('#'));
      const pillsHtml = mappingLines.slice(0, 6).map(line => {
        const parts = line.split(/(?:->|=>|→|=)/);
        if (parts.length === 2) {
          return `<span class="ef-pair-pill"><b>${esc(parts[0].trim())}</b> → <b>${esc(parts[1].trim())}</b></span>`;
        }
        return `<span class="ef-pair-pill">${esc(line)}</span>`;
      }).join('');
      const moreCount = mappingLines.length > 6 ? `<span class="ef-pair-pill more">+${mappingLines.length - 6} cột</span>` : '';

      const dateStr = p.updatedAt ? new Date(p.updatedAt).toLocaleDateString('vi-VN') : 'Mới';

      return `
        <div class="ef-preset-card" data-preset-id="${esc(p.id)}">
          <div class="ef-preset-card-head">
            <div>
              <h4 class="ef-preset-card-title">${esc(p.name)}</h4>
              ${p.description ? `<p class="ef-preset-card-desc">${esc(p.description)}</p>` : ''}
            </div>
            <span class="ef-preset-badge">${mappingLines.length} cặp cột</span>
          </div>
          <div class="ef-preset-pills">${pillsHtml}${moreCount}</div>
          <div class="ef-preset-card-foot">
            <span class="ef-preset-time">Cập nhật: ${dateStr}</span>
            <div class="ef-preset-actions">
              <button type="button" class="btn small primary ef-btn-apply" data-action="apply">Áp dụng</button>
              <button type="button" class="btn small outline ef-btn-edit" data-action="edit">Sửa</button>
              <button type="button" class="btn small danger ef-btn-del" data-action="delete">Xóa</button>
            </div>
          </div>
        </div>`;
    }).join('');

    // Gắn sự kiện cho các nút trên từng Card
    grid.querySelectorAll('.ef-preset-card').forEach(card => {
      const id = card.dataset.presetId;
      card.querySelector('[data-action="apply"]')?.addEventListener('click', () => {
        const target = allPresets.find(p => p.id === id);
        if (target) {
          applyPreset(target);
          closePresetModal();
        }
      });
      card.querySelector('[data-action="edit"]')?.addEventListener('click', () => {
        openPresetEditor(id);
      });
      card.querySelector('[data-action="delete"]')?.addEventListener('click', () => {
        const target = allPresets.find(p => p.id === id);
        if (!target) return;
        if (confirm(`Bạn có chắc chắn muốn xóa bộ mẫu "${target.name}"?`)) {
          deletePresetById(id);
          renderPresetCards();
          toast(`Đã xóa bộ mẫu "${target.name}".`);
        }
      });
    });
  }

  function openPresetModal() {
    ensurePresetModalUI();
    $('efPresetModalBackdrop').classList.add('show');
    $('efPresetSearch').value = '';
    renderPresetCards();
  }

  function closePresetModal() {
    $('efPresetModalBackdrop')?.classList.remove('show');
  }

  function openPresetEditor(id = null) {
    ensurePresetModalUI();
    editingPresetId = id;
    if (id) {
      const item = getPresets().find(p => p.id === id);
      if (!item) return;
      $('efPresetEditTitle').textContent = 'Chỉnh sửa bộ mẫu';
      $('efPresetEditName').value = item.name || '';
      $('efPresetEditDesc').value = item.description || '';
      $('efPresetEditMapping').value = item.mapping || '';
      $('efPresetEditFormats').value = item.formats || '';
    } else {
      $('efPresetEditTitle').textContent = 'Thêm bộ mẫu mới';
      $('efPresetEditName').value = 'Bộ mẫu ' + (new Date()).toLocaleDateString('vi-VN');
      $('efPresetEditDesc').value = '';
      // Lấy quy tắc hiện tại trên màn hình nếu có
      $('efPresetEditMapping').value = $('mapping')?.value.trim() || '';
      $('efPresetEditFormats').value = $('formatRules')?.value.trim() || '';
    }
    $('efPresetEditBackdrop').classList.add('show');
  }

  function closePresetEditor() {
    $('efPresetEditBackdrop')?.classList.remove('show');
    editingPresetId = null;
  }

  function savePresetEditor() {
    const name = ($('efPresetEditName')?.value || '').trim();
    if (!name) {
      alert('Vui lòng nhập tên bộ mẫu.');
      $('efPresetEditName')?.focus();
      return;
    }
    const mapping = ($('efPresetEditMapping')?.value || '').trim();
    const formats = ($('efPresetEditFormats')?.value || '').trim();
    const desc = ($('efPresetEditDesc')?.value || '').trim();

    addOrUpdatePreset({
      id: editingPresetId || ('preset_' + Date.now()),
      name,
      description: desc,
      mapping,
      formats,
      reverse: $('reverse')?.checked || false,
      skipTotals: $('skipTotals')?.checked !== false,
      skipZeros: $('skipZeros')?.checked !== false
    });

    closePresetEditor();
    renderPresetCards();
    toast(`✓ Đã lưu bộ mẫu "${name}" thành công!`);
  }

  function quickSaveCurrent() {
    const currentMapping = ($('mapping')?.value || '').trim();
    if (!currentMapping) {
      alert('Bạn cần ghép ít nhất một cặp cột trước khi lưu bộ mẫu.');
      return;
    }
    const name = prompt('Đặt tên cho bộ mẫu ghép cột:', 'Bộ sao kê mới');
    if (!name || !name.trim()) return;

    addOrUpdatePreset({
      name: name.trim(),
      description: 'Lưu nhanh từ cấu hình hiện tại',
      mapping: currentMapping,
      formats: $('formatRules')?.value.trim() || '',
      reverse: $('reverse')?.checked || false,
      skipTotals: $('skipTotals')?.checked !== false,
      skipZeros: $('skipZeros')?.checked !== false
    });

    toast(`✓ Đã lưu bộ mẫu "${name.trim()}"!`);
  }

  function bindEvents() {
    ensurePresetModalUI();
    seedDefaultPresetsIfEmpty();
    updatePresetDropdown();

    // Sự kiện mở modal quản lý
    $('efOpenPresetManager')?.addEventListener('click', openPresetModal);
    $('efSaveCurrentPreset')?.addEventListener('click', quickSaveCurrent);
    $('efPresetModalClose')?.addEventListener('click', closePresetModal);
    $('efPresetModalBackdrop')?.addEventListener('click', e => {
      if (e.target.id === 'efPresetModalBackdrop') closePresetModal();
    });

    // Search trong popup
    $('efPresetSearch')?.addEventListener('input', renderPresetCards);

    // Thêm mới
    $('efPresetAddNewBtn')?.addEventListener('click', () => openPresetEditor(null));
    $('efPresetEditCancel')?.addEventListener('click', closePresetEditor);
    $('efPresetEditSave')?.addEventListener('click', savePresetEditor);
    $('efPresetEditBackdrop')?.addEventListener('click', e => {
      if (e.target.id === 'efPresetEditBackdrop') closePresetEditor();
    });

    // Nút lấy nhanh từ bảng hiện tại trong editor
    $('efPresetGetFromWorkspace')?.addEventListener('click', () => {
      const curMap = $('mapping')?.value.trim() || '';
      const curFmt = $('formatRules')?.value.trim() || '';
      if (!curMap) {
        alert('Hiện tại chưa có cặp cột nào được ghép trên màn hình.');
        return;
      }
      $('efPresetEditMapping').value = curMap;
      if (curFmt) $('efPresetEditFormats').value = curFmt;
      toast('Đã nạp quy tắc ghép từ bảng hiện tại!');
    });

    // Sự kiện thay đổi trên dropdown chọn nhanh
    $('efTemplateSelect')?.addEventListener('change', e => {
      const id = e.target.value;
      if (!id) return;
      const target = getPresets().find(p => p.id === id);
      if (target) {
        applyPreset(target);
      }
    });

    // Lắng nghe sự kiện thêm cặp cột thành công để ghi nhớ vào memory
    $('addPair')?.addEventListener('click', () => {
      setTimeout(() => {
        const srcSel = $('addSource');
        const dstSel = $('addTarget');
        const typeSel = $('addType');
        if (srcSel?.value && dstSel?.value) {
          const sText = (srcSel.options[srcSel.selectedIndex]?.textContent || '').replace(/^[A-Z]+(?:[–-][A-Z]+)?\s*[·.:\-]\s*/, '').trim();
          const dText = (dstSel.options[dstSel.selectedIndex]?.textContent || '').replace(/^[A-Z]+(?:[–-][A-Z]+)?\s*[·.:\-]\s*/, '').trim();
          const tType = typeSel?.value || 'auto';
          if (sText && dText) {
            saveMappingToMemory(sText, dText, tType);
          }
        }
      }, 100);
    });
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', bindEvents);
  } else {
    bindEvents();
  }
})();
