(function(){
  const $ = id => document.getElementById(id);
  const esc = v => String(v ?? '').replaceAll('&','&amp;').replaceAll('<','&lt;').replaceAll('>','&gt;').replaceAll('"','&quot;');
  let modalRows = [];
  let modalColumns = [];
  let selectedRowId = null;
  let editorMode = 'add';
  let editorEditingId = null;

  function ensurePopupUI(){
    const previewTop = document.querySelector('.ef-preview-top');
    if(previewTop && !document.getElementById('efOpenPreviewExpand')){
      const tools = document.createElement('div');
      tools.className = 'ef-preview-tools';
      tools.innerHTML = `
        <button type="button" class="ef-tool-btn" id="efOpenPreviewExpand" aria-label="Mở rộng bảng xem trước" title="Mở rộng bảng xem trước">
          <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" aria-hidden="true">
            <path d="M15 4H20V9" stroke="#FFFFFF" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"/>
            <path d="M9 20H4V15" stroke="#FFFFFF" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"/>
            <path d="M20 4L14 10" stroke="#FFFFFF" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"/>
            <path d="M4 20L10 14" stroke="#FFFFFF" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"/>
          </svg>
          <span class="screen-reader">Mở rộng bảng xem trước</span>
        </button>`;
      previewTop.appendChild(tools);
    }
    if(!document.getElementById('efPreviewModalBackdrop')){
      const wrap = document.createElement('div');
      wrap.id = 'efPreviewModalBackdrop';
      wrap.className = 'ef-preview-modal-backdrop';
      wrap.innerHTML = `
        <div class="ef-preview-modal" id="efPreviewModal" role="dialog" aria-modal="true" aria-labelledby="efPreviewModalTitle">
          <div class="ef-modal-header">
            <div class="ef-modal-title">
              <div class="ef-icon">▤</div>
              <div><h3 id="efPreviewModalTitle">Xem trước chi tiết</h3><p>Kiểm tra kỹ dữ liệu trước khi xuất file • Có search, lọc và chỉnh sửa tạm thời.</p></div>
            </div>
            <div class="ef-modal-actions">
              <button type="button" id="efModalClose">Đóng</button>
            </div>
          </div>
          <div class="ef-modal-toolbar">
            <div class="field"><label for="efModalSearch">Tìm kiếm</label><input id="efModalSearch" type="text" placeholder="Tìm theo diễn giải, file nguồn, ô đích..."></div>
            <div class="field"><label for="efModalFilter">Lọc trạng thái</label><select id="efModalFilter"><option value="all">Tất cả</option><option value="import">Sẽ nhập</option><option value="matched">Đã ghép</option><option value="skipped">Bỏ qua</option></select></div>
            <button type="button" class="toolbtn" id="efModalAdd">+ Thêm dòng</button>
            <button type="button" class="toolbtn" id="efModalEdit">Sửa dòng</button>
            <button type="button" class="toolbtn warn" id="efModalRemoveZeros" title="Tự động lọc và xóa các dòng có số tiền bằng 0 khỏi bảng xem trước">⊘ Loại số 0</button>
            <button type="button" class="toolbtn danger" id="efModalDelete">Xóa dòng</button>
          </div>
          <div class="ef-modal-stats" id="efModalStats"></div>
          <div class="ef-modal-body">
            <div class="ef-modal-tablewrap"><table class="ef-modal-table" id="efModalTable"></table></div>
          </div>
          <div class="ef-modal-footer"><span id="efModalFoot">Chưa có dữ liệu xem trước.</span><strong>Lưu ý: phần CRUD trong popup dùng để rà soát dữ liệu trước khi xuất file.</strong></div>
        </div>`;
      document.body.appendChild(wrap);
    }
    if(!document.getElementById('efEditorBackdrop')){
      const editor = document.createElement('div');
      editor.id = 'efEditorBackdrop';
      editor.className = 'ef-editor-backdrop';
      editor.innerHTML = `
        <div class="ef-editor" role="dialog" aria-modal="true" aria-labelledby="efEditorTitle">
          <h4 id="efEditorTitle">Chỉnh sửa dòng xem trước</h4>
          <div class="ef-editor-grid" id="efEditorFields"></div>
          <div class="ef-editor-actions">
            <button type="button" id="efEditorCancel">Hủy</button>
            <button type="button" class="primary" id="efEditorSave">Lưu dòng</button>
          </div>
        </div>`;
      document.body.appendChild(editor);
    }
  }

  function buildRowsFromPreview(result){
    const cfg = result?.settings || {};
    modalColumns = Array.isArray(cfg.mappings) ? cfg.mappings.map(m => {
      const target = String(m.target || '');
      const opt = $('addTarget') ? [...$('addTarget').options].find(o => String(o.value) === target) : null;
      return { id: target || ('col_' + Math.random().toString(36).slice(2,6)), name: (opt ? opt.textContent : 'Cột ' + target).replace(/^[A-Z]+(?:[–-][A-Z]+)?\s*[·.:\-]\s*/, '').trim() || ('Cột ' + target) };
    }) : [];

    // Lưu danh sách vị trí dòng đích ban đầu để khi dồn dòng không để lại khoảng trống
    if(result && Array.isArray(result.changes)){
      const rowSet = new Set(result.changes.map(c => Number(c.row)).filter(r => Number.isInteger(r) && r > 0));
      result._initialRowSlots = Array.from(rowSet).sort((a,b) => a - b);
    }

    const map = new Map();
    (result?.changes || []).forEach((ch, idx) => {
      if(!ch._origRow) ch._origRow = ch.row;
      const key = ch.sourceIndex + '|' + ch._origRow;
      if(!map.has(key)) map.set(key, {
        _id: 'row_' + key,
        kind: 'import',
        status: 'Sẽ nhập',
        file: ch.sourceFile || '—',
        dest: ch.dest || ('Dòng ' + (ch.row ?? '')),
        note: 'Nhập mới',
        values: {},
        row: ch.row ?? '',
        _origRow: ch._origRow,
        sourceIndex: ch.sourceIndex ?? '',
        matched: true
      });
      const row = map.get(key);
      row.values[String(ch.col)] = ch.new ?? '';
    });
    const rows = [...map.values()];
    const skipped = (result?.skipDetails || []).map((s, idx) => ({
      _id: 'skip_' + idx,
      kind: 'skipped',
      status: 'Bỏ qua',
      file: s.file || '—',
      dest: s.dest || s.source || '—',
      note: s.reason || 'Bỏ qua',
      values: {},
      matched: false
    }));
    return [...rows, ...skipped];
  }

  function filteredRows(){
    const q = ($('efModalSearch')?.value || '').trim().toLowerCase();
    const filter = $('efModalFilter')?.value || 'all';
    return modalRows.filter(r => {
      if(filter !== 'all' && r.kind !== filter) return false;
      if(!q) return true;
      const hay = [r.status,r.file,r.dest,r.note,...modalColumns.map(c => r.values?.[c.id])].map(v => String(v ?? '').toLowerCase()).join(' | ');
      return hay.includes(q);
    });
  }

  function renderModalStats(rows){
    const total = modalRows.length;
    const imports = modalRows.filter(r => r.kind === 'import').length;
    const matched = modalRows.filter(r => r.kind === 'matched').length;
    const skipped = modalRows.filter(r => r.kind === 'skipped').length;
    const shown = rows.length;
    $('efModalStats').innerHTML = `
      <span class="chip">Tổng: <strong>${total}</strong></span>
      <span class="chip">Sẽ nhập: <strong>${imports}</strong></span>
      <span class="chip">Đã ghép: <strong>${matched}</strong></span>
      <span class="chip">Bỏ qua: <strong>${skipped}</strong></span>
      <span class="chip">Đang hiển thị: <strong>${shown}</strong></span>`;
    $('efModalFoot').textContent = shown ? `Đang hiển thị ${shown} / ${total} dòng trong popup.` : 'Không có dòng nào phù hợp với điều kiện tìm kiếm/lọc.';
  }

  function renderModalTable(){
    const table = $('efModalTable');
    if(!table) return;
    const rows = filteredRows();
    renderModalStats(rows);
    if(!modalColumns.length && rows.length){
      const dynamic = new Set();
      rows.forEach(r => Object.keys(r.values || {}).forEach(k => dynamic.add(k)));
      modalColumns = [...dynamic].map(k => ({ id:k, name:k }));
    }
    table.innerHTML = `<thead><tr><th style="width:44px">#</th><th>Trạng thái</th>${modalColumns.map(c=>`<th>${esc(c.name)}</th>`).join('')}<th>File nguồn</th><th>Ô / dòng đích</th><th>Ghi chú</th></tr></thead><tbody></tbody>`;
    const tbody = table.querySelector('tbody');
    if(!rows.length){
      tbody.innerHTML = `<tr><td colspan="${modalColumns.length + 5}"><div class="empty" style="margin:14px 0">Không có dữ liệu phù hợp.</div></td></tr>`;
      return;
    }
    tbody.innerHTML = rows.map((r,i) => `
      <tr data-row-id="${esc(r._id)}" class="${selectedRowId===r._id?'is-selected':''}">
        <td>${i+1}</td>
        <td><span class="ef-mini-badge ${esc(r.kind)}">${esc(r.status || '')}</span></td>
        ${modalColumns.map(c=>`<td class="wrap" title="${esc(r.values?.[c.id] ?? '')}">${esc(r.values?.[c.id] ?? '—')}</td>`).join('')}
        <td title="${esc(r.file || '')}">${esc(r.file || '—')}</td>
        <td title="${esc(r.dest || '')}">${esc(r.dest || '—')}</td>
        <td class="wrap" title="${esc(r.note || '')}">${esc(r.note || '—')}</td>
      </tr>`).join('');
    [...tbody.querySelectorAll('tr[data-row-id]')].forEach(tr => tr.addEventListener('click', () => {
      selectedRowId = tr.dataset.rowId;
      renderModalTable();
    }));
  }

  function openModal(fullscreen=false){
    ensurePopupUI();
    const result = window.__excelFlowPreview;
    if(result && (!modalRows.length || !modalColumns.length)){
      modalRows = buildRowsFromPreview(result);
    }
    $('efPreviewModalBackdrop').classList.add('show');
    $('efPreviewModal').classList.toggle('fullscreen', !!fullscreen);
    renderModalTable();
  }
  function closeModal(){ $('efPreviewModalBackdrop')?.classList.remove('show'); }

  function openEditor(mode){
    ensurePopupUI();
    editorMode = mode;
    let row;
    if(mode === 'edit'){
      row = modalRows.find(r => r._id === selectedRowId);
      if(!row){ alert('Bạn hãy chọn một dòng trước khi sửa.'); return; }
      editorEditingId = row._id;
      $('efEditorTitle').textContent = 'Sửa dòng xem trước';
    } else {
      row = { _id:'manual_' + Date.now(), kind:'import', status:'Sẽ nhập', file:'Tự thêm', dest:'', note:'', values:{} };
      editorEditingId = row._id;
      $('efEditorTitle').textContent = 'Thêm dòng xem trước';
    }
    const fields = $('efEditorFields');
    fields.innerHTML = `
      <div class="field"><label>Trạng thái</label><select id="efEditKind"><option value="import">Sẽ nhập</option><option value="matched">Đã ghép</option><option value="skipped">Bỏ qua</option></select></div>
      <div class="field"><label>File nguồn</label><input id="efEditFile" type="text" value="${esc(row.file || '')}"></div>
      <div class="field"><label>Ô / dòng đích</label><input id="efEditDest" type="text" value="${esc(row.dest || '')}"></div>
      <div class="field"><label>Ghi chú</label><input id="efEditNote" type="text" value="${esc(row.note || '')}"></div>
      ${modalColumns.map(c=>`<div class="field ${String(c.name).length>22?'full':''}"><label>${esc(c.name)}</label><textarea data-col-id="${esc(c.id)}">${esc(row.values?.[c.id] ?? '')}</textarea></div>`).join('')}`;
    $('efEditKind').value = row.kind || 'import';
    $('efEditorBackdrop').classList.add('show');
  }
  function closeEditor(){ $('efEditorBackdrop')?.classList.remove('show'); }

  function saveEditor(){
    const kind = $('efEditKind').value;
    const row = {
      _id: editorEditingId,
      kind,
      status: kind==='matched'?'Đã ghép':kind==='skipped'?'Bỏ qua':'Sẽ nhập',
      file: $('efEditFile').value.trim(),
      dest: $('efEditDest').value.trim(),
      note: $('efEditNote').value.trim(),
      values: {}
    };
    [...document.querySelectorAll('#efEditorFields [data-col-id]')].forEach(el => row.values[el.dataset.colId] = el.value);
    const idx = modalRows.findIndex(r => r._id === editorEditingId);
    if(idx >= 0) modalRows[idx] = row; else modalRows.unshift(row);
    selectedRowId = row._id;
    closeEditor();
    renderModalTable();
  }

  function colLetter(num){
    let s = '';
    for(; num > 0; num = Math.floor((num - 1) / 26)) s = String.fromCharCode(65 + (num - 1) % 26) + s;
    return s;
  }

  function syncChangesAndShiftRows(){
    const preview = window.__excelFlowPreview;
    if(!preview) return;

    if(!preview._initialRowSlots || !preview._initialRowSlots.length){
      if(Array.isArray(preview.changes) && preview.changes.length){
        const rowSet = new Set(preview.changes.map(c => Number(c.row)).filter(r => Number.isInteger(r) && r > 0));
        preview._initialRowSlots = Array.from(rowSet).sort((a,b) => a - b);
      } else {
        preview._initialRowSlots = [];
      }
    }
    const initialSlots = preview._initialRowSlots;

    const remainingImportRows = modalRows.filter(r => r.kind === 'import');
    const remainingKeys = new Set(remainingImportRows.map(r => r._id));

    if(Array.isArray(preview.changes)){
      preview.changes = preview.changes.filter(ch => {
        const origRow = ch._origRow ?? ch.row;
        const key = 'row_' + ch.sourceIndex + '|' + origRow;
        return remainingKeys.has(key);
      });

      remainingImportRows.forEach((r, idx) => {
        const newTargetRow = idx < initialSlots.length 
          ? initialSlots[idx] 
          : (initialSlots.length ? (initialSlots[initialSlots.length - 1] + (idx - initialSlots.length + 1)) : (idx + 1));
        
        const oldRow = r.row;
        r.row = newTargetRow;
        r.dest = 'Dòng ' + newTargetRow;

        preview.changes.forEach(ch => {
          if(!ch._origRow) ch._origRow = oldRow;
          const key = 'row_' + ch.sourceIndex + '|' + ch._origRow;
          if(key === r._id){
            ch.row = newTargetRow;
            ch.dst = colLetter(ch.col) + newTargetRow;
          }
        });
      });
    }

    if(preview.summary){
      preview.summary.matched = remainingImportRows.length;
    }

    const statChanges = document.getElementById('statChanges');
    if(statChanges && preview.changes) statChanges.textContent = String(preview.changes.length);
    const statRows = document.getElementById('statRows');
    if(statRows) statRows.textContent = String(remainingImportRows.length);

    const previewTable = document.getElementById('previewTable');
    if(previewTable && preview.changes && preview.changes.length){
      const list = preview.changes.slice(0, 150);
      const escHtml = v => String(v ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
      const html = '<table class="data-table"><thead><tr><th>File / Ô nguồn</th><th>Ô đích</th><th>Trước khi nhập</th><th>Sau khi nhập</th></tr></thead><tbody>' +
        list.map(ch => '<tr><td>' + escHtml(ch.src) + '</td><td><strong>' + escHtml(ch.dst) + '</strong></td><td>' + escHtml(ch.old || '(trống)') + '</td><td class="diff">' + escHtml(ch.new) + '</td></tr>').join('') +
        '</tbody></table>';
      previewTable.innerHTML = html;
    }
  }

  function isZeroValue(v){
    if(v === 0 || v === '0') return true;
    if(typeof v === 'string'){
      const s = v.trim();
      if(s === '0' || s === '0.0' || s === '0,0' || s === '0.00' || s === '0,00' || s === '-0' || s === '0 ₫' || s === '0 đ') return true;
      const clean = s.replace(/[^\d.-]/g, '');
      if(clean === '0' || clean === '0.0' || clean === '0.00' || clean === '0,00') return true;
      const num = Number(clean);
      if(!isNaN(num) && num === 0 && clean.length > 0) return true;
    }
    return false;
  }

  function isZeroRow(r){
    if(!r || !r.values) return false;
    for(const col of modalColumns){
      const cName = (col.name || '').toLowerCase();
      const val = r.values[col.id];
      if(cName.includes('tiền') || cName.includes('amount') || cName.includes('giá') || cName.includes('st') || cName.includes('nợ') || cName.includes('có') || cName.includes('thu') || cName.includes('chi')){
        if(isZeroValue(val)) return true;
      }
    }
    for(const val of Object.values(r.values)){
      if(isZeroValue(val)) return true;
    }
    return false;
  }

  function removeZeroRows(){
    const zeroRows = modalRows.filter(r => isZeroRow(r));
    if(!zeroRows.length){
      if(typeof window.efToast === 'function') window.efToast('Không tìm thấy dòng nào có giá trị bằng 0.');
      else alert('Không tìm thấy dòng nào có giá trị bằng 0.');
      return;
    }
    const ok = window.confirm('Tìm thấy ' + zeroRows.length + ' dòng có số tiền bằng 0. Bạn có chắc muốn loại bỏ và tự động dồn các dòng còn lại lên liền kề?');
    if(!ok) return;

    modalRows = modalRows.filter(r => !isZeroRow(r));
    if(selectedRowId && zeroRows.some(r => r._id === selectedRowId)) selectedRowId = null;

    syncChangesAndShiftRows();
    renderModalTable();

    const msg = 'Đã loại bỏ ' + zeroRows.length + ' dòng số 0 và dồn các dòng lên liên tục!';
    if(typeof window.efToast === 'function') window.efToast(msg);
    else alert(msg);
  }

  function deleteSelected(){
    if(!selectedRowId){ alert('Bạn hãy chọn một dòng trước khi xóa.'); return; }
    const row = modalRows.find(r => r._id === selectedRowId);
    if(!row) return;
    const ok = window.confirm('Bạn muốn xóa dòng đang chọn và dồn các dòng phía sau lên?');
    if(!ok) return;
    modalRows = modalRows.filter(r => r._id !== selectedRowId);
    selectedRowId = null;

    syncChangesAndShiftRows();
    renderModalTable();

    const msg = 'Đã xóa dòng và dồn thứ tự các dòng đích!';
    if(typeof window.efToast === 'function') window.efToast(msg);
  }

  function bindPopupEvents(){
    ensurePopupUI();
    $('efOpenPreviewExpand')?.addEventListener('click', () => openModal(true));
    $('efModalClose')?.addEventListener('click', closeModal);
    $('efPreviewModalBackdrop')?.addEventListener('click', e => { if(e.target.id === 'efPreviewModalBackdrop') closeModal(); });
    $('efModalSearch')?.addEventListener('input', renderModalTable);
    $('efModalFilter')?.addEventListener('change', renderModalTable);
    $('efModalAdd')?.addEventListener('click', () => openEditor('add'));
    $('efModalEdit')?.addEventListener('click', () => openEditor('edit'));
    $('efModalRemoveZeros')?.addEventListener('click', removeZeroRows);
    $('efModalDelete')?.addEventListener('click', deleteSelected);
    $('efEditorCancel')?.addEventListener('click', closeEditor);
    $('efEditorSave')?.addEventListener('click', saveEditor);
    $('efEditorBackdrop')?.addEventListener('click', e => { if(e.target.id === 'efEditorBackdrop') closeEditor(); });
    document.addEventListener('keydown', e => {
      if(e.key === 'Escape'){
        if($('efEditorBackdrop')?.classList.contains('show')) closeEditor();
        else if($('efPreviewModalBackdrop')?.classList.contains('show')) closeModal();
      }
    });
  }

  document.addEventListener('DOMContentLoaded', () => {
    ensurePopupUI();
    bindPopupEvents();
    const previewBtn = $('previewBtn');
    if(previewBtn){
      previewBtn.addEventListener('click', () => { modalRows = []; modalColumns = []; selectedRowId = null; setTimeout(() => {
        if(window.__excelFlowPreview){ modalRows = buildRowsFromPreview(window.__excelFlowPreview); renderModalTable(); }
      }, 700); });
    }
    window.addEventListener('excelflow:preview', ev => { modalRows = buildRowsFromPreview(ev.detail || window.__excelFlowPreview || {}); selectedRowId = null; renderModalTable(); });
  });
})();
