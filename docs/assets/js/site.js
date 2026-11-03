(function(){
  'use strict';

  var doc = document;
  var html = doc.documentElement;
  var nav = doc.getElementById('nav');
  var progress = doc.getElementById('scrollProg');

  function onScroll(){
    if(progress){
      var max = Math.max(1, doc.documentElement.scrollHeight - window.innerHeight);
      progress.style.width = Math.min(100, Math.max(0, window.scrollY / max * 100)) + '%';
    }
    if(nav) nav.classList.toggle('scrolled', window.scrollY > 32);
  }
  window.addEventListener('scroll', onScroll, {passive:true});
  onScroll();

  var menuButton = doc.getElementById('mobMenuBtn') || doc.getElementById('mobBtn');
  var menu = doc.getElementById('mobMenu');
  var menuClose = doc.getElementById('mobMenuClose');

  function setMenu(open){
    if(!menu) return;
    menu.classList.toggle('open', open);
    html.style.overflow = open ? 'hidden' : '';
    if(menuButton){
      menuButton.textContent = open ? '✕' : '☰';
      menuButton.setAttribute('aria-expanded', String(open));
    }
    menu.setAttribute('aria-hidden', String(!open));
    if(open){
      var firstLink = menu.querySelector('a');
      if(firstLink) window.setTimeout(function(){ firstLink.focus(); }, 40);
    }
  }

  if(menuButton){
    menuButton.setAttribute('type','button');
    menuButton.setAttribute('aria-label','Open navigation menu');
    menuButton.setAttribute('aria-controls','mobMenu');
    menuButton.setAttribute('aria-expanded','false');
    menuButton.addEventListener('click', function(event){
      event.stopPropagation();
      setMenu(!(menu && menu.classList.contains('open')));
    });
  }
  if(menuClose){
    menuClose.setAttribute('type','button');
    menuClose.setAttribute('aria-label','Close navigation menu');
    menuClose.addEventListener('click', function(){ setMenu(false); });
  }
  if(menu){
    menu.setAttribute('aria-hidden','true');
    menu.querySelectorAll('a').forEach(function(link){
      link.addEventListener('click', function(){ setMenu(false); });
    });
  }
  doc.addEventListener('click', function(event){
    if(menu && menu.classList.contains('open') && !menu.contains(event.target) && !(menuButton && menuButton.contains(event.target))){
      setMenu(false);
    }
  });
  doc.addEventListener('keydown', function(event){
    if(event.key === 'Escape'){
      setMenu(false);
      closeLightbox();
    }
  });

  if('IntersectionObserver' in window){
    var observer = new IntersectionObserver(function(entries){
      entries.forEach(function(entry){
        if(entry.isIntersecting){
          entry.target.classList.add('in');
          observer.unobserve(entry.target);
        }
      });
    }, {threshold:.06, rootMargin:'0px 0px -20px 0px'});
    doc.querySelectorAll('.reveal').forEach(function(el){ observer.observe(el); });
  }else{
    doc.querySelectorAll('.reveal').forEach(function(el){ el.classList.add('in'); });
  }

  // Documentation sidebar and scroll spy.
  var sidebarToggle = doc.getElementById('sidebarToggle');
  var sidebar = doc.getElementById('docsSidebar');
  if(sidebar){
    function updateSidebar(){
      if(sidebarToggle) sidebarToggle.style.display = window.innerWidth <= 800 ? 'flex' : 'none';
      if(window.innerWidth > 800) sidebar.classList.remove('expanded');
    }
    updateSidebar();
    window.addEventListener('resize', updateSidebar, {passive:true});
    if(sidebarToggle){
      sidebarToggle.setAttribute('type','button');
      sidebarToggle.addEventListener('click', function(){ sidebar.classList.toggle('expanded'); });
    }
    sidebar.querySelectorAll('.sidebar-link').forEach(function(link){
      link.addEventListener('click', function(){
        if(window.innerWidth <= 800) sidebar.classList.remove('expanded');
      });
    });
    var sidebarLinks = Array.from(doc.querySelectorAll('.sidebar-link[href^="#"]'));
    var headings = Array.from(doc.querySelectorAll('.docs-content h1[id],.docs-content h2[id]'));
    if(sidebarLinks.length && headings.length){
      var spy = function(){
        var current = headings[0].id;
        headings.forEach(function(heading){ if(window.scrollY >= heading.offsetTop - 120) current = heading.id; });
        sidebarLinks.forEach(function(link){ link.classList.toggle('active', link.getAttribute('href') === '#' + current); });
      };
      window.addEventListener('scroll', spy, {passive:true});
      spy();
    }
  }

  // Lightweight accessible image lightbox.
  var lightbox;
  var lightboxImage;
  var previousFocus;
  function ensureLightbox(){
    if(lightbox) return;
    lightbox = doc.createElement('div');
    lightbox.className = 'ac-lightbox';
    lightbox.setAttribute('role','dialog');
    lightbox.setAttribute('aria-modal','true');
    lightbox.setAttribute('aria-label','Image preview');
    lightbox.innerHTML = '<button class="ac-lightbox-close" type="button" aria-label="Close image preview">✕</button><img alt=""/>';
    doc.body.appendChild(lightbox);
    lightboxImage = lightbox.querySelector('img');
    lightbox.querySelector('button').addEventListener('click', closeLightbox);
    lightbox.addEventListener('click', function(event){ if(event.target === lightbox) closeLightbox(); });
  }
  function openLightbox(img){
    ensureLightbox();
    previousFocus = doc.activeElement;
    lightboxImage.src = img.currentSrc || img.src;
    lightboxImage.alt = img.alt || 'AnchorCast screenshot';
    lightbox.classList.add('open');
    html.style.overflow = 'hidden';
    lightbox.querySelector('button').focus();
  }
  function closeLightbox(){
    if(!lightbox || !lightbox.classList.contains('open')) return;
    lightbox.classList.remove('open');
    lightboxImage.removeAttribute('src');
    html.style.overflow = '';
    if(previousFocus && previousFocus.focus) previousFocus.focus();
  }
  window.closeLightbox = closeLightbox;
  doc.querySelectorAll('img[data-lightbox]').forEach(function(img){
    img.setAttribute('tabindex','0');
    img.setAttribute('role','button');
    img.setAttribute('aria-label',(img.alt || 'Image') + ' — open larger view');
    img.addEventListener('click', function(){ openLightbox(img); });
    img.addEventListener('keydown', function(event){
      if(event.key === 'Enter' || event.key === ' '){ event.preventDefault(); openLightbox(img); }
    });
  });
})();
