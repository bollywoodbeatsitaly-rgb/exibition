(function(){
  // Hamburger nav toggle
  var nav = document.querySelector('.site-nav');
  var btn = document.querySelector('.nav-toggle');
  if(btn){
    btn.addEventListener('click', function(){
      nav.classList.toggle('nav-open');
    });
    document.querySelectorAll('.nav-links a').forEach(function(a){
      a.addEventListener('click', function(){
        nav.classList.remove('nav-open');
      });
    });
  }

  // Lazy-load all images that are NOT in the hero / page-hero (above the fold)
  document.querySelectorAll('img').forEach(function(img){
    var aboveFold = img.closest('.hero-media') || img.closest('.page-hero');
    if(!aboveFold) img.loading = 'lazy';
  });
})();
