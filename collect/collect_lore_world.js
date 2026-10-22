/*
  Collect every review on LORE's own site (lore.world).

  How to run (desktop Chrome):
  1. Open any product page on lore.world and scroll until the reviews show.
  2. Open the console (Cmd+Option+J on Mac, Ctrl+Shift+J on Windows) and paste this file.
  3. Two files download: lore_world_reviews_raw.csv (move it into private/) and lore_world_site_stats.json
     (move it into private/verification/), which holds each listing's own totals and star breakdowns
     for analysis/verify_reviews.py.

  How it works: lore.world is a Shopify store whose reviews come from Yotpo. The script
  reads the store's public Yotpo key from the widget the page loaded, lists every product
  from Shopify's public products feed, and fetches each product's reviews.
  Yotpo groups reviews across related listings (a scent, its travel size and the sets),
  so the "listings" column records every listing a review appeared under. The analysis
  assigns a review to a scent only when it appears under exactly one scent's listing.
*/
(async () => {
  const loader = performance.getEntriesByType('resource').map(e => e.name)
    .find(u => /cdn-widgetsrepository\.yotpo\.com\/v1\/loader\//.test(u));
  if (!loader) { console.warn('Scroll to the reviews first, then run again.'); return; }
  const key = loader.split('/loader/')[1].split('?')[0];
  const pause = ms => new Promise(r => setTimeout(r, ms));
  const label = t => {
    const s = t.toLowerCase();
    for (const [k, v] of [['lil softie', 'Lil Softie'], ['sublimity', 'Sublimity'], ['lovely', 'Lovely'],
      ['disfruta', 'Disfruta']]) if (s.includes(k)) return v;
    if (s.includes('somewhere')) return s.includes('set') ? 'SomewhereSet' : 'Somewhere';
    if (s.includes('discovery')) return 'Discovery';
    if (s.includes('everything')) return 'EverythingSet';
    return t;
  };

  const products = (await fetch('/products.json?limit=250').then(r => r.json())).products;
  const seen = new Map(); const stats = []; const membership = {};
  for (const p of products) {
    const ids = []; let bottom = null;
    let page = 1, total = 1;
    while ((page - 1) * 50 < total) {
      const j = await fetch(`https://api-cdn.yotpo.com/v3/storefront/store/${key}/product/${p.id}/reviews?page=${page}&perPage=50&sort=date`)
        .then(r => r.json());
      total = j.pagination?.total || 0; bottom = bottom || j.bottomline;
      for (const r of j.reviews || []) {
        ids.push(String(r.id));
        const e = seen.get(r.id) || { r, listings: new Set() };
        e.listings.add(label(p.title)); seen.set(r.id, e);
      }
      page++; await pause(400);
    }
    stats.push({ id: String(p.id), title: p.title, handle: p.handle, total, avg: bottom?.averageScore,
      dist: bottom?.starDistribution, fetched: ids.length, unique: new Set(ids).size });
    membership[String(p.id)] = ids;
  }

  const cell = v => '"' + String(v ?? '').replace(/"/g, '""').replace(/[\r\n\t]+/g, ' ').trim() + '"';
  const lines = [['review_id','listings','date','rating','verified','incentivized','country','title','text'].join(',')];
  for (const { r, listings } of seen.values()) {
    lines.push([r.id, [...listings].join('|'), r.createdAt.slice(0, 10), r.score, r.verifiedBuyer,
      r.isIncentivized, r.country?.code, r.title, r.content].map(cell).join(','));
  }
  const a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob([lines.join('\n')], { type: 'text/csv' }));
  a.download = 'lore_world_reviews_raw.csv'; a.click();
  const reviews = [...seen.values()].map(({ r }) => [String(r.id), r.createdAt.slice(0, 10), r.score, r.verifiedBuyer,
    r.isIncentivized, (r.title || '').replace(/\s+/g, ' ').trim(), (r.content || '').replace(/\s+/g, ' ').trim()]);
  const b = document.createElement('a');
  b.href = URL.createObjectURL(new Blob([JSON.stringify({ pulled: new Date().toISOString(), stats, membership, reviews })], { type: 'application/json' }));
  b.download = 'lore_world_site_stats.json'; setTimeout(() => b.click(), 800);
  console.log(`Done: ${seen.size} unique reviews`);
})();
