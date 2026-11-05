// Applies the saved colour theme before first paint (external file to satisfy the CSP).
try {
  var t = localStorage.getItem('salatic:theme');
  if (t === 'dark' || t === 'light') document.documentElement.dataset.theme = t;
} catch (e) {}
