(function(){
  var nav = document.querySelector('.site-nav');
  var btn = document.querySelector('.nav-toggle');
  if(!btn) return;
  btn.addEventListener('click', function(){
    nav.classList.toggle('nav-open');
  });
  document.querySelectorAll('.nav-links a').forEach(function(a){
    a.addEventListener('click', function(){
      nav.classList.remove('nav-open');
    });
  });
})();
