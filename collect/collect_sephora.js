/*
  Collect every LORE review on Sephora US.

  How to run (desktop Chrome):
  1. Open any LORE product page on sephora.com, for example the Sublimity page.
  2. Scroll down until the "Ratings & Reviews" section has loaded.
  3. Open the console (Cmd+Option+J on Mac, Ctrl+Shift+J on Windows) and paste this file.
  4. Two files download: sephora_reviews_raw.csv (move it into private/) and sephora_site_stats.json
     (move it into private/verification/). The stats file holds Sephora's own totals, averages and
     star breakdowns, which analysis/verify_reviews.py checks the reviews against. To verify an
     earlier collection, run this again later and replace only the stats file.

  How it works: Sephora's review section is powered by Bazaarvoice. The script reuses
  the review request the page itself made (including the page's own public display
  key), looks up every product from the same brand, then pages through each product's
  reviews 100 at a time, oldest first, with a pause between requests.
  Full-size and travel-size listings often share one review pool, so reviews are
  de-duplicated by review ID.
*/
(async () => {
  const req = performance.getEntriesByType('resource').map(e => e.name)
    .find(u => /api\.bazaarvoice\.com\/data\/reviews\.json/.test(u));
  if (!req) { console.warn('Scroll to the Reviews section first, then run again.'); return; }
  const base = new URL(req);
  const passkey = base.searchParams.get('passkey');
  const pid = base.searchParams.getAll('Filter').find(f => f.startsWith('ProductId:')).split(':')[1];
  const api = (path, params) => {
    const u = new URL('https://api.bazaarvoice.com/data/' + path);
    Object.entries(params).forEach(([k, v]) => [].concat(v).forEach(x => u.searchParams.append(k, x)));
    u.searchParams.set('passkey', passkey); u.searchParams.set('apiversion', '5.4'); u.searchParams.set('Locale', 'en_US');
    return fetch(u).then(r => r.json());
  };
  const pause = ms => new Promise(r => setTimeout(r, ms));

  const brandId = (await api('products.json', { Filter: `Id:${pid}` })).Results[0].Brand.Id;
  const productResults = (await api('products.json', { Filter: `BrandId:${brandId}`, Limit: 100, Stats: 'Reviews' })).Results;
  const products = productResults.map(p => ({ id: p.Id, name: p.Name }));
  const stats = productResults.map(p => ({
    id: p.Id, name: p.Name, total: p.ReviewStatistics?.TotalReviewCount, avg: p.ReviewStatistics?.AverageOverallRating,
    dist: Object.fromEntries((p.ReviewStatistics?.RatingDistribution || []).map(d => [d.RatingValue, d.Count])),
    recommended: p.ReviewStatistics?.RecommendedCount, notRecommended: p.ReviewStatistics?.NotRecommendedCount,
    familyIds: p.FamilyIds || null }));
  console.log(`Found ${products.length} products`);

  const seen = new Map();
  for (const p of products) {
    let offset = 0, total = 1;
    while (offset < total) {
      const j = await api('reviews.json', {
        Filter: [`ProductId:${p.id}`], Sort: 'SubmissionTime:asc', Limit: 100, Offset: offset,
      });
      total = j.TotalResults || 0;
      for (const r of j.Results || []) if (!seen.has(r.Id)) seen.set(r.Id, { r, p });
      offset += 100; await pause(600);
    }
    console.log(`${p.name}: ${total} reviews`);
  }

  const cell = v => '"' + String(v ?? '').replace(/"/g, '""').replace(/[\r\n\t]+/g, ' ').trim() + '"';
  const header = ['review_id','product_id','product_name','date','rating','recommended','incentivized','staff',
    'helpful_yes','helpful_no','badges','title','text'];
  const lines = [header.join(',')];
  for (const { r, p } of seen.values()) {
    const ctx = r.ContextDataValues || {};
    lines.push([r.Id, r.ProductId || p.id, p.name, r.SubmissionTime.slice(0, 10), r.Rating, r.IsRecommended,
      ctx.IncentivizedReview?.Value ?? '', ctx.StaffContext?.Value ?? '', r.TotalPositiveFeedbackCount,
      r.TotalNegativeFeedbackCount, Object.keys(r.Badges || {}).join('|'), r.Title, r.ReviewText].map(cell).join(','));
  }
  const a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob([lines.join('\n')], { type: 'text/csv' }));
  a.download = 'sephora_reviews_raw.csv'; a.click();
  const meta = [...seen.values()].map(({ r }) => [r.Id, r.ProductId, r.OriginalProductName || '', r.Rating,
    r.SubmissionTime.slice(0, 10), r.IsSyndicated, r.IsRatingsOnly, r.ModerationStatus,
    r.ContextDataValues?.IncentivizedReview?.Value || '', Object.keys(r.Badges || {}).join('|'),
    (r.Title || '').replace(/\s+/g, ' ').trim(), (r.ReviewText || '').replace(/\s+/g, ' ').trim()]);
  const b = document.createElement('a');
  b.href = URL.createObjectURL(new Blob([JSON.stringify({ pulled: new Date().toISOString(), stats, meta })], { type: 'application/json' }));
  b.download = 'sephora_site_stats.json'; setTimeout(() => b.click(), 800);
  console.log(`Done: ${seen.size} unique reviews`);
})();
