/* Plataforma Municipal do Associativismo — configuração pública
   Depois de publicar o Apps Script como Aplicação Web, cole o URL abaixo. */
window.ASSOCIATIVISMO_ENDPOINT = '';

(function(){
  var css=document.createElement('link');
  css.rel='stylesheet';
  css.href='preview-match-v2.css?v=2';
  document.head.appendChild(css);
  var js=document.createElement('script');
  js.src='preview-match-v2.js?v=2';
  js.defer=true;
  document.head.appendChild(js);
})();
