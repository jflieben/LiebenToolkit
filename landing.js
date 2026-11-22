// Landing page: tool cards per group from the registry, plus a quick filter.
(() => {
  const { GROUPS, TOOLS, PRODUCTS } = window.ToolkitRegistry;
  const { h, productCard } = window.Toolkit;

  function render() {
    const host = document.getElementById('toolGroups');
    host.replaceChildren(...GROUPS.map((g) => h('section.tool-group', { dataset: { group: g.id } }, [
      h('h2.section-title', { text: g.label }),
      h('div.tool-grid', null, TOOLS.filter((t) => t.group === g.id).map((t) => {
        const card = productCard(t);
        card.dataset.search = `${t.id} ${t.tagline} ${t.description}`.toLowerCase();
        return card;
      })),
    ])));
    document.getElementById('productGrid').replaceChildren(...PRODUCTS.map((p) => productCard(p)));
    // Stagger the card entrance a little.
    document.querySelectorAll('.tool-card').forEach((c, i) => { c.style.animationDelay = `${Math.min(i, 16) * 30}ms`; });
  }

  function filter(q) {
    const words = q.toLowerCase().split(/\s+/).filter(Boolean);
    let any = false;
    document.querySelectorAll('.tool-group').forEach((section) => {
      let shown = 0;
      section.querySelectorAll('.tool-card').forEach((c) => {
        const hit = words.every((w) => c.dataset.search.includes(w));
        c.hidden = !hit;
        if (hit) shown++;
      });
      section.hidden = shown === 0;
      any = any || shown > 0;
    });
    document.getElementById('noMatch').hidden = any;
  }

  function boot() {
    render();
    const input = document.getElementById('toolSearch');
    input.addEventListener('input', () => filter(input.value));
    input.addEventListener('keydown', (e) => {
      if (e.key !== 'Enter') return;
      const first = document.querySelector('.tool-group .tool-card:not([hidden])');
      if (first) first.click();
    });
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', boot);
  else boot();
})();
