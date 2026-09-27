document.addEventListener("DOMContentLoaded", async () => {
  const headerPlaceholder = document.getElementById("header-placeholder");
  if (headerPlaceholder) {
    try {
      const res = await fetch('/header');
      const data = await res.text();
      headerPlaceholder.outerHTML = data;
    } catch (e) { console.error("Error loading header:", e); }
  }

  const footerPlaceholder = document.getElementById("footer-placeholder");
  if (footerPlaceholder) {
    try {
      const res = await fetch('/footer');
      const data = await res.text();
      footerPlaceholder.outerHTML = data;
    } catch (e) { console.error("Error loading footer:", e); }
  }

  initGlobalUI();
});

function initGlobalUI() {

  // Active Link Highlighting
  const style = document.createElement('style');
  style.textContent = `
    .nav-link.active::after { width: 100%; }
    .nav-link.active { color: #7A1F2B; font-weight: 600; }
  `;
  document.head.appendChild(style);

  const currentPath = window.location.pathname;
  
  // Desktop Nav
  document.querySelectorAll('#navbar nav a').forEach(link => {
    try {
      const linkPath = new URL(link.href).pathname;
      if (linkPath === currentPath) {
        // Skip root anchor links unless we specifically want them
        const rawHref = link.getAttribute('href');
        if (!rawHref.startsWith('/#')) {
          link.classList.add('active');
        }
      }
    } catch(e) {}
  });

  // Mobile Nav
  document.querySelectorAll('#mobileBottomNav a').forEach(link => {
    try {
      const linkPath = new URL(link.href).pathname;
      if (linkPath === currentPath) {
        const rawHref = link.getAttribute('href');
        if (!rawHref.startsWith('/#')) {
          link.classList.remove('text-ink/70');
          link.classList.add('text-maroon', 'font-semibold');
        }
      }
    } catch(e) {}
  });

const yearEl = document.getElementById('year');
  if (yearEl) yearEl.textContent = new Date().getFullYear();

  const navbar = document.getElementById('navbar');
  if (navbar) {
    window.addEventListener('scroll', () => {
      if (window.scrollY > 8) navbar.classList.add('shadow-soft'); 
      else navbar.classList.remove('shadow-soft');
    });
  }

  document.querySelectorAll('a.coming-soon-link').forEach(a => {
    a.addEventListener('click', (e) => {
      e.preventDefault();
      showComingSoonPopup();
    });
  });
}

function showComingSoonPopup(){
  let overlay = document.getElementById('comingSoonOverlay');
  if (!overlay){
    overlay = document.createElement('div');
    overlay.id = 'comingSoonOverlay';
    overlay.style.cssText = 'position:fixed;inset:0;z-index:9999;display:flex;align-items:center;justify-content:center;background:rgba(28,26,36,0);padding:20px;transition:background .3s ease;pointer-events:none;';

    const modal = document.createElement('div');
    modal.id = 'comingSoonModal';
    modal.style.cssText = 'background:#FBF7EE;color:#1C1A24;font-family:Mukta,sans-serif;border-radius:24px;padding:32px 28px;max-width:340px;width:100%;text-align:center;box-shadow:0 25px 60px -15px rgba(28,26,36,.45);transform:scale(.85) translateY(10px);opacity:0;transition:transform .3s cubic-bezier(.34,1.56,.64,1), opacity .3s ease;';
    modal.innerHTML = '<div style="font-size:40px;line-height:1;margin-bottom:12px;">🎭</div>'
            + '<h3 style="font-family:\'Baloo Da 2\',sans-serif;font-weight:800;font-size:20px;color:#7A1F2B;margin:0 0 8px;">Coming Soon!</h3>'
            + '<p style="font-size:14px;color:rgba(28,26,36,.65);margin:0;line-height:1.5;">This page is on its way. Stay tuned \u2014 we\'re putting the final touches on it.</p>';
    overlay.appendChild(modal);
    document.body.appendChild(overlay);
  }

  const modal = document.getElementById('comingSoonModal');
  overlay.style.pointerEvents = 'auto';
  overlay.style.background = 'rgba(28,26,36,.45)';
  modal.style.opacity = '1';
  modal.style.transform = 'scale(1) translateY(0)';

  clearTimeout(overlay._hideTimer);
  overlay._hideTimer = setTimeout(() => {
    overlay.style.background = 'rgba(28,26,36,0)';
    overlay.style.pointerEvents = 'none';
    modal.style.opacity = '0';
    modal.style.transform = 'scale(.85) translateY(10px)';
  }, 2000);
}
