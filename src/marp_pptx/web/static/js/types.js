// Live filter for the slide-type gallery.
(function () {
  const search = document.getElementById('gallery-search');
  const cards = [...document.querySelectorAll('.gcard')];
  const sections = [...document.querySelectorAll('.gcat')];
  const empty = document.getElementById('gallery-empty');
  if (!search) return;

  function apply() {
    const q = search.value.trim().toLowerCase();
    let total = 0;
    for (const card of cards) {
      const hit = !q || card.dataset.search.toLowerCase().includes(q);
      card.hidden = !hit;
      if (hit) total++;
    }
    // hide category sections that have no visible cards
    for (const sec of sections) {
      const any = [...sec.querySelectorAll('.gcard')].some(c => !c.hidden);
      sec.hidden = !any;
    }
    if (empty) empty.hidden = total !== 0;
  }

  search.addEventListener('input', apply);
})();
