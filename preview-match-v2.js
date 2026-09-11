document.addEventListener('DOMContentLoaded',()=>{
  const h=document.querySelector('.inst-section-head h2');
  if(h && h.textContent.trim()==='O que pretende fazer?') h.textContent='Procedimentos disponíveis';
  const sub=document.querySelector('.inst-section-head p');
  if(sub) sub.textContent='';
});
