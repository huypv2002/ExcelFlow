(function(){
 const overlay=document.getElementById('efApprovalOverlay');
 const title=document.getElementById('efApprovalTitle');
 const desc=document.getElementById('efApprovalDesc');
 const yes=document.getElementById('efApprovalYes');
 const cancel=document.getElementById('efApprovalCancel');
 let settle=null, lastFocus=null;
 window.efAsk=function(message, heading='Xác nhận thao tác', danger=false){
  if(settle)return Promise.resolve(false);
  return new Promise(resolve=>{
   lastFocus=document.activeElement;
   settle=resolve;
   title.textContent=heading;
   desc.textContent=message;
   yes.textContent=danger?'Vẫn thực hiện':'Đồng ý';
   yes.classList.toggle('danger',!!danger);
   overlay.hidden=false;
   document.body.style.overflow='hidden';
   cancel.focus();
  });
 };
 function finish(ok){
  if(!settle)return;
  const done=settle;settle=null;
  overlay.hidden=true;document.body.style.overflow='';
  if(lastFocus?.isConnected)lastFocus.focus();
  done(ok);
 }
 yes.addEventListener('click',()=>finish(true));
 cancel.addEventListener('click',()=>finish(false));
 overlay.addEventListener('click',e=>{if(e.target===overlay)finish(false)});
 document.addEventListener('keydown',e=>{
  if(!settle)return;
  if(e.key==='Escape'){e.preventDefault();finish(false)}
  if(e.key==='Tab'){
   const items=[cancel,yes];const index=items.indexOf(document.activeElement);
   if(e.shiftKey&&index===0){e.preventDefault();yes.focus()}
   else if(!e.shiftKey&&index===1){e.preventDefault();cancel.focus()}
  }
 });
 document.addEventListener('click',e=>{
  const btn=e.target.closest('[data-return-home]');
  if(!btn)return;
  document.getElementById('navImport')?.click();
  window.scrollTo({top:0,behavior:'smooth'});
 });
})();
