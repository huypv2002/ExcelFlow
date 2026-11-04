(function(){
 const $=id=>document.getElementById(id);
 const esc=v=>String(v??'').replaceAll('&','&amp;').replaceAll('<','&lt;').replaceAll('>','&gt;').replaceAll('"','&quot;');
 const tabLabels={all:'Tất cả',import:'Sẽ nhập',matched:'Đã ghép',skipped:'Bỏ qua'};
 let previewRows=[],previewFilter='all',page=0;const pageSize=14;
 const fmtOptions=[['auto','Giữ nguyên'],['text','Văn bản'],['date','Ngày'],['datetime','Ngày giờ'],['money','Tiền tệ'],['integer','Số nguyên'],['decimal','Thập phân'],['percent','Phần trăm']];
 function syncSteps(){
  const hasSource=!!document.querySelector('#sourceList .source-entry'),hasTarget=!!$('dstName')?.classList.contains('ready');
  const mapped=!!$('mapping')?.value.trim(),previewed=!!window.__excelFlowPreview;
  [hasSource,hasTarget,mapped,previewed].forEach((ready,i)=>{
   const el=$('efStep'+(i+1));if(el){el.classList.toggle('complete',ready);el.classList.toggle('current',!ready && (i===0 || [hasSource,hasTarget,mapped][i-1]));}
  });
  if($('efStep3')?.classList.contains('complete'))$('efStep3').classList.remove('current');
 }
 function renderRows(){
  const box=$('efRowPreview');if(!box)return;
  let rows=previewRows;
  if(previewFilter==='import')rows=rows.filter(r=>r.kind==='import');
  if(previewFilter==='matched')rows=rows.filter(r=>r.matched);
  if(previewFilter==='skipped')rows=rows.filter(r=>r.kind==='skipped');
  let maxpage=Math.max(0,Math.ceil(rows.length/pageSize)-1);page=Math.min(maxpage,Math.max(0,page));
  if(!window.__excelFlowPreview){box.innerHTML='<div class="ef-blank"><span class="ef-blank-icon">▦</span><strong>Chưa có dữ liệu xem trước</strong><span>Thêm file, ghép cột và bấm <b>Xem trước</b> để kiểm tra kết quả.</span></div>';}
  else if(!rows.length){box.innerHTML='<div class="ef-blank"><span class="ef-blank-icon">▤</span><strong>Không có mục nào trong nhóm này</strong><span>Chọn bộ lọc khác để xem kết quả.</span></div>';}
  else{
   const cfg=window.__excelFlowPreview.settings||{};
   let mapping=Array.isArray(cfg.mappings)?cfg.mappings:[];
   const maxCols=7;
   if(!mapping.length){box.innerHTML='<div class="ef-blank">Chưa có cột dữ liệu.</div>';return;}
   mapping=mapping.slice(0,maxCols);
   const cols=mapping.map(m=>{const target=String(m.target||'');const opt=[...$('addTarget').options].find(o=>String(o.value)===target);return {id:target,name:(opt?opt.textContent:'Cột '+target).replace(/^[A-Z]+(?:[–-][A-Z]+)?\s*[·.:\-]\s*/, '').trim() || 'Cột '+target};});
   const slice=rows.slice(page*pageSize,(page+1)*pageSize);
   box.innerHTML='<table><thead><tr><th>#</th><th>Trạng thái</th>'+cols.map(c=>'<th>'+esc(c.name)+'</th>').join('')+'<th>File nguồn</th><th>Ô đích</th></tr></thead><tbody>'+slice.map((r,i)=>{
    const kind=r.kind==='skipped'?'skipped':r.kind==='import'?'import':'matched';
    const status=kind==='skipped'?'Bỏ qua':kind==='import'?'Sẽ nhập':'Đã ghép';
    const label=r.kind==='skipped'?r.reason||'Không ghi':(r.kind==='import'?'Nhập mới':'Đã có dòng đích');
    return '<tr><td class="rowid">'+(page*pageSize+i+1)+'</td><td><span class="ef-status status-'+kind+'">'+status+'</span></td>'+cols.map(c=>'<td title="'+esc(r.values[c.id]??'')+'">'+esc(r.values[c.id]??'—')+'</td>').join('')+'<td title="'+esc(r.file||'')+'">'+esc(r.file||'—')+'</td><td title="'+esc(label)+'">'+esc(r.dest||label)+'</td></tr>';}).join('')+'</tbody></table>';
  }
  $('efPreviewPagination').textContent=rows.length?'Hiển thị '+(page*pageSize+1)+'–'+Math.min((page+1)*pageSize,rows.length)+' / '+rows.length+' dòng':'0 dòng';
  $('efPreviewPage').textContent=(page+1)+' / '+Math.max(1,Math.ceil(rows.length/pageSize));
  $('efPrevPreview').disabled=page===0;$('efNextPreview').disabled=page===maxpage;
  document.querySelectorAll('[data-preview-filter]').forEach(b=>b.classList.toggle('active',b.dataset.previewFilter===previewFilter));
 }
 function updatePreview(result){
  window.__excelFlowPreview=result;
  const rows=new Map();
  (result.changes||[]).forEach(ch=>{
   const k=ch.sourceIndex+'|'+ch.row;
   if(!rows.has(k))rows.set(k,{kind:'import',matched:true,values:{},file:ch.sourceFile,dest:'Dòng '+ch.row,row:ch.row});
   rows.get(k).values[String(ch.col)]=ch.new;
  });
  const skipped=(result.skipDetails||[]).map(s=>({kind:'skipped',matched:false,values:{},reason:s.reason,file:s.file,dest:s.dest||s.source||'—'}));
  previewRows=[...rows.values(),...skipped];page=0;previewFilter='all';
  $('efCountAll').textContent=previewRows.length;
  $('efCountImport').textContent=rows.size;
  $('efCountMatched').textContent=rows.size;
  $('efCountSkipped').textContent=skipped.length;
  $('previewLabel').textContent='Đã kiểm tra';
  renderRows();syncSteps();
 }
 window.addEventListener('excelflow:preview',ev=>updatePreview(ev.detail));
 function renderMapTypes(){
  const list=$('mappingList');if(!list)return;
  const raw=$('mapping').value.split('\n').filter(x=>x.trim());
  const formats=$('formatRules').value.split('\n').map(x=>x.trim());
  list.querySelectorAll('.mapping-row').forEach((row,i)=>{
   if(row.querySelector('.mapping-type'))return;
   const parsed=raw[i]||'',m=parsed.match(/(?:->|=>|→|=)\s*([A-Z]+)/i),target=m?.[1]?.toUpperCase()||'';
   const spec=formats.find(x=>new RegExp('^'+target+'\\s*(?:->|=>|→|=)','i').test(x));
   const val=spec?.split(/->|=>|→|=/)[1]?.trim()||'auto';
   const wrap=document.createElement('div');wrap.className='mapping-type';
   const select=document.createElement('select');select.setAttribute('aria-label','Kiểu dữ liệu cho cột đích '+target);
   fmtOptions.forEach(([value,label])=>{const o=new Option(label,value);if(value===val)o.selected=true;select.add(o);});
   select.addEventListener('change',()=>{
    const next=$('formatRules').value.split('\n').filter(x=>x.trim()&&!new RegExp('^'+target+'\\s*(?:->|=>|→|=)','i').test(x));
    if(select.value!=='auto')next.push(target+' -> '+select.value);
    $('formatRules').value=next.join('\n');$('formatRules').dispatchEvent(new Event('input',{bubbles:true}));
   });
   wrap.append(select);
   const span=document.createElement('div');span.className='mapping-option';span.textContent=(val==='date'?'dd/MM/yyyy':val==='money'?'VND':'—');
   const del=row.querySelector('.remove-map');row.insertBefore(wrap,del);row.insertBefore(span,del);
  });
 }
 const mappingObs=new MutationObserver(()=>{if($('mappingList'))renderMapTypes();syncSteps();});
 mappingObs.observe($('mappingList'),{childList:true});
 renderMapTypes();
 $('onlyFormulaRowsToggle').addEventListener('change',function(){ $('rowFilter').value=this.checked?'withFormula':'all';$('rowFilter').dispatchEvent(new Event('change',{bubbles:true})); });
 $('rowFilter').addEventListener('change',()=>{$('onlyFormulaRowsToggle').checked=$('rowFilter').value==='withFormula';});
 document.querySelectorAll('[data-preview-filter]').forEach(b=>b.addEventListener('click',()=>{previewFilter=b.dataset.previewFilter;page=0;renderRows()}));
 $('efPrevPreview').addEventListener('click',()=>{page--;renderRows()});$('efNextPreview').addEventListener('click',()=>{page++;renderRows()});
 document.querySelectorAll('[data-go-step]').forEach(b=>b.addEventListener('click',()=>{const step=+b.dataset.goStep;const target=step===1?'efFiles':step===2?'efFiles':step===3?'efMapping':'efPreview';$(target).scrollIntoView({behavior:'smooth',block:'start'});if(step===2)$('targetDrop').classList.add('ef-highlight');setTimeout(()=>{if($('targetDrop'))$('targetDrop').classList.remove('ef-highlight')},1100)}));
 const sourceObs=new MutationObserver(syncSteps);sourceObs.observe($('sourceList'),{childList:true,subtree:true});
 const targetObs=new MutationObserver(syncSteps);targetObs.observe($('dstName'),{attributes:true,childList:true,subtree:true});
 const origPreview=$('previewBtn');origPreview.addEventListener('click',()=>{window.__excelFlowPreview=null;previewRows=[];renderRows()});
 document.querySelectorAll('.ef-files .dropzone').forEach(zone=>{zone.addEventListener('dragover',e=>e.preventDefault());});

 function resetPreviewUI(){
  window.__excelFlowPreview=null;previewRows=[];page=0;previewFilter='all';
  ['efCountAll','efCountImport','efCountMatched','efCountSkipped'].forEach(id=>$(id).textContent='0');
  renderRows();syncSteps();
 }
 ['mapping','formatRules','sourceSheet','targetSheet','rowFilter','reverse','skipTotals','onlyFormulaRowsToggle','sourceFile','targetFile'].forEach(id=>{
  const node=$(id);if(!node)return;node.addEventListener('change',resetPreviewUI);
  if(id==='mapping'||id==='formatRules')node.addEventListener('input',resetPreviewUI);
 });
 $('addPair').addEventListener('click',resetPreviewUI);
 syncSteps();renderRows();
})();
