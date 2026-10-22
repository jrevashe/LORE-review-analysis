"""
Step 4: cross-check every review and every number before publishing.

Run from the project folder, after code_reviews.py and make_results.py:
    python analysis/verify_reviews.py

It reads the collected reviews in private/, plus a later snapshot of each site's own figures
(private/verification/*_site_stats.json, saved by the collect/ scripts), and runs four groups of
checks:

  A. Completeness: for every listing, does the number of reviews we hold match the number the
     site reports? Does the star breakdown and average rating match?
  B. Accuracy: does every stored review still match the live site, word for word, along with
     its stars, date, listing and gifted flag?
  C. Consistency: is anything inside a review at odds with how it was filed or labeled?
     (gifted disclosure vs gifted flag, scent named vs listing,
     stars vs wording, labels vs their supporting quote, duplicates, dates.)
  D. Article: does every number quoted in the article reproduce from the data?

Anything that can't be confirmed automatically is written to the human-review list.

Outputs (all in private/, because they quote review text):
    private/verification/verification_report.md   pass/fail summary of every check
    private/verification/flags.csv                one row per flag
    private/needs_human_review.xlsx               flagged reviews with full text, to decide by hand
"""
from pathlib import Path
import json
import re
import warnings

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

warnings.filterwarnings("ignore", "This pattern is interpreted as a regular expression")
ROOT = Path(__file__).resolve().parent.parent
PRIV = ROOT / "private"
VER = PRIV / "verification"
LAUNCH = pd.Timestamp("2025-09-01")

# Every number the article states, so a change in the data can't silently break the text.
ARTICLE_CLAIMS = {
    "Sephora average stars": ("sephora_avg", 4.2, 0.05),
    "lore.world average stars": ("lore_avg", 4.7, 0.05),
    "Paying Sephora reviewers mentioning fading (%)": ("paid_fades", 18, 0.5),
    "Gifted Sephora reviewers mentioning fading (%)": ("gift_fades", 10, 0.5),
    "lore.world reviewers mentioning fading (%)": ("lore_fades", 15, 0.5),
    "Stars when a review mentions fading": ("stars_fades", 3.6, 0.05),
    "Stars otherwise": ("stars_other", 4.3, 0.05),
    "Original four scents, paying, fading before Apr 2026 (%)": ("trend_before", 11, 0.5),
    "Original four scents, paying, fading since Apr 2026 (%)": ("trend_since", 21, 0.5),
    "Reviews behind the trend, before": ("trend_before_n", 148, 0),
    "Reviews behind the trend, since": ("trend_since_n", 89, 0),
    "Share of Sephora reviews gifted (%)": ("gifted_share", 81, 0.5),
    "Gifted reviews in April 2026": ("gifted_apr", 532, 0),
    "Gifted reviewers' stars": ("gift_stars", 4.2, 0.05),
    "Paying reviewers' stars": ("paid_stars", 4.0, 0.05),
    "Disfruta stars": ("disfruta_stars", 3.8, 0.05),
    "Disfruta 3-star-and-below reviews mentioning cigarette-type smoke (%)": ("disfruta_smoke", 42, 0.5),
    "Unique reviews analyzed": ("n_total", 1914, 0),
    "Sephora reviews": ("n_sephora", 1760, 0),
    "lore.world reviews": ("n_lore", 154, 0),
    "Fragrantica reviews (cross-check)": ("n_frag", 350, 0),
}

SCENT_WORDS = {
    "Sublimity": r"sublimity",
    "Disfruta": r"disfruta",
    "Lovely": r"lovely and a little twisted|lovely & a little twisted|\bllt\b",
    "Somewhere": r"somewhere but nowhere|somewhere but no where|\bsbn\b",
    "Lil Softie": r"lil softie|lil' softie|little softie",
}
# Wording that says the reviewer has the travel or mini size
SMALL = (r"(got|bought|purchased|ordered|have|own|tried|picked up|grabbed|went with|using|use) (the |a |my )?"
         r"(travel|mini|small|10 ?ml|rollerball|sample)"
         r"|travel[- ]size|travel spray|the mini\b|a mini\b|mini size|mini version|10 ?ml")
# Wording that says the reviewer has the full bottle
FULL = r"(got|bought|purchased|ordered|have|own) (the |a |my )?(full[- ]size|full bottle|50 ?ml|big bottle)"
GIFT_TEXT = (r"\[i received this product|received (this|it) (complimentary|for free|free)|complimentary from"
             r"|influenster|in exchange for (my|an) honest|gifted (to me|this)|sent (to me )?(for free|complimentary)")
BOUGHT_TEXT = r"\bi (bought|purchased|paid for) (this|it)\b|not gifted|with my own money"
NEG = (r"\b(hate|hated|awful|disgusting|horrible|terrible|gross|returning|returned it|waste of money"
       r"|not for me|smells bad|can'?t stand|disappointed|disappointing)\b")
POS = r"\b(love it|loved it|obsessed|my favorite|favourite|amazing|perfect|holy grail|repurchase|new signature)\b"
FADES_KW = (r"(doesn't|does not|didn't|did not|don't) (seem to |really |even )?last|fades?|faint|weak|disappear"
            r"|longevity|projection|sillage|reapply|wish(ed)? (it|this) (lasted|was stronger)")


def norm(x):
    return re.sub(r"\s+", " ", str(x)).strip()


def low(x):
    return norm(x).lower().replace("’", "'").replace("‘", "'")


class Report:
    def __init__(self):
        self.checks, self.flags = [], []

    def check(self, group, name, ok, detail):
        self.checks.append({"group": group, "check": name, "result": "PASS" if ok else "FAIL", "detail": detail})

    def info(self, group, name, detail):
        self.checks.append({"group": group, "check": name, "result": "NOTE", "detail": detail})

    def flag(self, source, review_id, check, detail, level="Needs human review"):
        self.flags.append({"source": source, "review_id": str(review_id), "check": check, "level": level,
                           "detail": detail})


def load():
    s = pd.read_csv(PRIV / "sephora_reviews_raw.csv", dtype=str, keep_default_na=False)
    w = pd.read_csv(PRIV / "lore_world_reviews_raw.csv", dtype=str, keep_default_na=False)
    f = pd.read_csv(PRIV / "fragrantica_reviews_raw.csv", dtype=str, keep_default_na=False)
    sc = pd.read_csv(ROOT / "data/sephora_reviews_coded.csv", dtype={"review_id": str})
    wc = pd.read_csv(ROOT / "data/lore_world_reviews_coded.csv", dtype={"review_id": str})
    fc = pd.read_csv(ROOT / "data/fragrantica_reviews_coded.csv", dtype={"review_id": str})
    labels = pd.read_csv(VER / "full_read_labels_with_quotes.csv", dtype={"uid": str}, keep_default_na=False)
    fixes = pd.read_csv(ROOT / "data/label_corrections.csv", dtype=str)
    for fx in fixes.itertuples():  # corrected labels no longer need their old quote
        labels.loc[labels.uid == fx.uid, fx.field] = int(fx.value)
        labels.loc[labels.uid == fx.uid, "quote"] = ""
    return s, w, f, sc, wc, fc, labels


def pools_from_family(stats):
    """Sephora shows one review pool across a scent's full-size and travel listings. Join listings into pools."""
    parent = {st["id"]: st["id"] for st in stats}

    def find(x):
        while parent[x] != x:
            parent[x] = parent[parent[x]]
            x = parent[x]
        return x

    for st in stats:
        for fid in st.get("familyIds") or []:
            if fid in parent:
                parent[find(fid)] = find(st["id"])
    pools = {}
    for st in stats:
        pools.setdefault(find(st["id"]), []).append(st["id"])
    return list(pools.values())


def check_sephora(rep, s, sc):
    G = "A. Completeness"
    path = VER / "sephora_site_stats.json"
    if not path.exists():
        rep.info(G, "Sephora site snapshot", "private/verification/sephora_site_stats.json missing; skipped")
        return
    site = json.load(open(path))
    stats = {st["id"]: st for st in site["stats"]}
    cols = ["review_id", "product_id", "original_product_name", "rating", "date", "syndicated", "ratings_only",
            "moderation", "incent", "badges", "title", "text"]
    meta = pd.DataFrame(site["meta"], columns=cols[:len(site["meta"][0])]).astype({"review_id": str})
    snap = site["pulled"][:10]

    # A1-A3: per pool counts, star breakdown, average
    for pool in pools_from_family(site["stats"]):
        st = stats[pool[0]]
        names = " + ".join(sorted(stats[p]["name"] for p in pool))
        ours = s[s.product_id.isin(pool)]
        site_ids = set(meta[meta.product_id.isin(pool)].review_id)
        our_ids = set(ours.review_id)
        new, gone = site_ids - our_ids, our_ids - site_ids
        rep.check(G, f"Review count: {names}", len(ours) == st["total"] or (not gone and len(ours) + len(new) == st["total"]),
                  f"site {st['total']} on {snap}, ours {len(ours)}; new on site since collection {len(new)}, "
                  f"on file but no longer on site {len(gone)}")
        for rid in sorted(new):
            rep.flag("sephora", rid, "New review on site since collection",
                     f"posted after the data was collected; not in the analysis ({names})", "Info")
        for rid in sorted(gone):
            rep.flag("sephora", rid, "Review no longer on site", f"was collected but is gone from {names}")
        ours_now = ours[~ours.review_id.isin(gone)]
        site_dist = {int(k): v for k, v in (st.get("dist") or {}).items()}
        new_dist = meta[meta.review_id.isin(new)].rating.astype(int).value_counts().to_dict()
        our_dist = ours_now.rating.astype(int).value_counts().to_dict()
        expected = {k: our_dist.get(k, 0) + new_dist.get(k, 0) for k in range(1, 6)}
        got = {k: site_dist.get(k, 0) for k in range(1, 6)}
        rep.check(G, f"Star breakdown: {names}", expected == got,
                  f"site {got}; ours plus new reviews {expected}")
        all_ratings = pd.concat([ours_now.rating.astype(int), meta[meta.review_id.isin(new)].rating.astype(int)])
        rep.check(G, f"Average stars: {names}", abs(all_ratings.mean() - st["avg"]) < 0.005,
                  f"site {st['avg']:.3f}, recomputed {all_ratings.mean():.3f}")
    rep.check(G, "No syndicated, hidden or ratings-only reviews on Sephora",
              (~meta.syndicated).all() and (~meta.ratings_only).all() and meta.moderation.eq("APPROVED").all(),
              f"{int(meta.syndicated.sum())} syndicated, {int(meta.ratings_only.sum())} ratings-only, "
              f"{int((meta.moderation != 'APPROVED').sum())} not approved")

    # B: each review vs the live snapshot
    G = "B. Accuracy"
    m = s.merge(meta, on="review_id", how="inner", suffixes=("", "_site"))
    for col, a, b in [("stars", "rating", "rating_site"), ("date", "date", "date_site"),
                      ("listing", "product_id", "product_id_site")]:
        bad = m[m[a].astype(str) != m[b].astype(str)]
        rep.check(G, f"Sephora {col} matches live site, every review", bad.empty,
                  f"{len(m) - len(bad)} of {len(m)} match")
        for _, r in bad.iterrows():
            rep.flag("sephora", r.review_id, f"{col} differs from live site", f"ours {r[a]}, site {r[b]}")
    bad = m[m.incentivized.str.lower() != m.incent.astype(str).str.lower()]
    rep.check(G, "Sephora gifted flag matches live site, every review", bad.empty, f"{len(m) - len(bad)} of {len(m)} match")
    for _, r in bad.iterrows():
        rep.flag("sephora", r.review_id, "Gifted flag differs from live site", f"ours '{r.incentivized}', site '{r.incent}'")
    name_ok = m.product_name == m.original_product_name
    rep.check(G, "Sephora listing name matches the name the site records for each review", name_ok.all(),
              f"{int(name_ok.sum())} of {len(m)} match")

    fresh = VER / "sephora_fresh_pull.csv"
    fr = meta[["review_id", "title", "text"]] if "text" in meta else (
        pd.read_csv(fresh, dtype=str, keep_default_na=False) if fresh.exists() else None)
    if fr is not None:
        mm = s.merge(fr, on="review_id", suffixes=("", "_site"))
        for col in ["title", "text"]:
            bad = mm[mm[col].map(norm) != mm[f"{col}_site"].map(norm)]
            rep.check(G, f"Sephora review {col} word for word, every review", bad.empty,
                      f"{len(mm) - len(bad)} of {len(mm)} identical (second pull)")
            for _, r in bad.iterrows():
                rep.flag("sephora", r.review_id, f"{col} differs from live site", "text changed or copied wrong")


def check_lore(rep, w):
    G = "A. Completeness"
    path = VER / "lore_world_site_stats.json"
    if not path.exists():
        rep.info(G, "lore.world site snapshot", "missing; skipped")
        return
    site = json.load(open(path))
    snap = site["pulled"][:10]
    live = pd.DataFrame(site["reviews"], columns=["review_id", "date", "rating", "verified", "incent", "title", "text"]).astype(
        {"review_id": str, "rating": str})
    ours = set(w.review_id)
    new, gone = set(live.review_id) - ours, ours - set(live.review_id)
    rep.check(G, "lore.world total unique reviews", not gone,
              f"site {live.review_id.nunique()} unique on {snap}, ours {len(ours)}; new since collection {len(new)}, gone {len(gone)}")
    for rid in sorted(new):
        rep.flag("lore_world", rid, "New review on site since collection", "posted after collection; not in the analysis", "Info")
    for rid in sorted(gone):
        rep.flag("lore_world", rid, "Review no longer on site", "collected but gone")
    by_listing = {}
    for pid, ids in site["membership"].items():
        for rid in ids:
            by_listing.setdefault(rid, set()).add(pid)
    for st in site["stats"]:
        if not st["total"]:
            continue
        ids = set(site["membership"][st["id"]])
        held = ids & ours
        rep.check(G, f"lore.world listing count: {st['title']} ({st['handle']})",
                  st["fetched"] == st["total"] == st["unique"] and len(held) + len(ids & new) == st["total"],
                  f"site says {st['total']}, fetched {st['fetched']}, held {len(held)}, new {len(ids & new)}")
        dist = live[live.review_id.isin(ids)].rating.astype(int).value_counts().to_dict()
        site_dist = {int(k): v for k, v in (st.get("dist") or {}).items()}
        rep.check(G, f"lore.world star breakdown: {st['title']} ({st['handle']})",
                  all(dist.get(k, 0) == site_dist.get(k, 0) for k in range(1, 6)),
                  f"site {site_dist}, recomputed {dist}")

    G = "B. Accuracy"
    m = w.merge(live, on="review_id", suffixes=("", "_site"))
    for col, a, b in [("stars", "rating", "rating_site"), ("date", "date", "date_site"),
                      ("verified flag", "verified", "verified_site"), ("gifted flag", "incentivized", "incent")]:
        bad = m[m[a].astype(str).str.lower() != m[b].astype(str).str.lower()]
        rep.check(G, f"lore.world {col} matches live site, every review", bad.empty, f"{len(m) - len(bad)} of {len(m)} match")
        for _, r in bad.iterrows():
            rep.flag("lore_world", r.review_id, f"{col} differs from live site", f"ours {r[a]}, site {r[b]}")
    for col in ["title", "text"]:
        bad = m[m[col].map(norm) != m[f"{col}_site"].map(norm)]
        rep.check(G, f"lore.world review {col} word for word, every review", bad.empty, f"{len(m) - len(bad)} of {len(m)} identical")
        for _, r in bad.iterrows():
            rep.flag("lore_world", r.review_id, f"{col} differs from live site", "text changed or copied wrong")


def check_fragrantica(rep, f):
    path = VER / "fragrantica_site_stats.json"
    if not path.exists():
        return
    site = json.load(open(path))["stated_review_counts"]
    for scent, n in site.items():
        held = (f.scent == scent).sum()
        rep.check("A. Completeness", f"Fragrantica review count: {scent}", held == n, f"page says {n}, held {held}")
    dup = f[f.text.map(low).duplicated(keep=False)]
    rep.info("C. Consistency", "Fragrantica duplicate posts", f"{len(dup)} posts share identical text with another post")
    for _, r in dup.iterrows():
        rep.flag("fragrantica", r.review_id, "Duplicate text", "same text posted twice; counted twice in the cross-check", "Info")


def check_consistency(rep, s, w, f, sc, wc, labels):
    G = "C. Consistency"
    collected = pd.Timestamp("2026-10-21")

    # Basic integrity
    for name, d in [("Sephora", s), ("lore.world", w), ("Fragrantica", f)]:
        dates = pd.to_datetime(d.date, errors="coerce")
        rep.check(G, f"{name}: review IDs unique", d.review_id.is_unique, f"{d.review_id.duplicated().sum()} repeated IDs")
        rep.check(G, f"{name}: every date valid and between launch and collection",
                  dates.notna().all() and dates.min() >= LAUNCH - pd.Timedelta(days=60) and dates.max() <= collected,
                  f"{dates.min():%Y-%m-%d} to {dates.max():%Y-%m-%d}")
        if "rating" in d:
            rep.check(G, f"{name}: every star rating is a whole number 1 to 5",
                      d.rating.astype(int).between(1, 5).all(), f"values {sorted(d.rating.astype(int).unique())}")

    for name, d, src in [("Sephora", s, "sephora"), ("lore.world", w, "lore_world")]:
        body = d.text.map(low).str.replace(r"\[i received this product in exchange for my honest review\]", "", regex=True).str.strip()
        dup = d[(body.str.len() > 40) & body.duplicated(keep=False)]
        rep.info(G, f"{name}: identical review text under different IDs", f"{len(dup)} reviews")
        for _, r in dup.iterrows():
            rep.flag(src, r.review_id, "Duplicate text",
                     "same text appears under another review ID; check whether it was counted twice")

    # Gifted: the flag Sephora stores vs what the review itself says
    t = (s.title + " . " + s.text).map(low)
    gifted = s.incentivized.str.lower().eq("true")
    blank = s.incentivized.eq("")
    says_gift, says_bought = t.str.contains(GIFT_TEXT), t.str.contains(BOUGHT_TEXT)
    official = t.str.contains(r"\[i received this product in exchange for my honest review\]")
    rep.info(G, "Sephora disclosure line without the gifted flag",
             f"{int((~gifted & official).sum())} reviews carry Sephora's '[I received this product...]' line with the "
             "flag blank or 'no'; the analysis counts them as gifted")
    for _, r in s[~gifted & says_gift & ~official].iterrows():
        rep.flag("sephora", r.review_id, "Gifted flag vs text",
                 "text suggests a free product but the review isn't marked gifted, so it was counted as paying")
    for _, r in s[gifted & says_bought & ~says_gift].iterrows():
        rep.flag("sephora", r.review_id, "Gifted flag vs text", "marked gifted but the text says the reviewer bought it")
    rep.info(G, "Sephora gifted question left blank",
             f"{int(blank.sum())} reviews didn't answer 'I received this product free'; the analysis counts them as "
             f"paying. See the sensitivity check in section D.")

    # Size isn't used anywhere in the analysis (reviews are grouped by scent), so it isn't checked.

    # Scent named in the review vs listing
    scent = sc.set_index("review_id").scent
    for src, d, sc_map in [("sephora", s, scent), ("lore_world", w, wc.set_index("review_id").scent)]:
        tt = (d.title + " . " + d.text).map(low)
        for idx, r in d.iterrows():
            listed = sc_map.get(r.review_id, "Set")
            if listed == "Set":
                continue
            named = [k for k, p in SCENT_WORDS.items() if re.search(p, tt[idx])]
            others = [k for k in named if k != listed]
            if others and listed not in named:
                rep.flag(src, r.review_id, "Scent vs listing",
                         f"listed as {listed} but the text names only {', '.join(others)}")
            elif others:
                rep.flag(src, r.review_id, "Scent vs listing",
                         f"listed as {listed}; also discusses {', '.join(others)}, so labels may mix scents", "Info")

    # Stars vs wording
    for src, d in [("sephora", s), ("lore_world", w)]:
        tt = (d.title + " . " + d.text).map(low)
        stars = d.rating.astype(int)
        for _, r in d[(stars >= 4) & tt.str.contains(NEG) & ~tt.str.contains(POS)].iterrows():
            rep.flag(src, r.review_id, "Stars vs wording", f"{r.rating} stars but the wording reads negative")
        for _, r in d[(stars <= 2) & tt.str.contains(POS) & ~tt.str.contains(NEG)].iterrows():
            rep.flag(src, r.review_id, "Stars vs wording", f"{r.rating} stars but the wording reads positive")

    # lore.world: which scent a review belongs to
    for _, r in w.iterrows():
        scents = [x for x in r.listings.split("|") if x in SCENT_WORDS]
        if len(scents) > 1:
            rep.flag("lore_world", r.review_id, "Scent assignment", f"appears under several scents: {r.listings}")
    rep.info(G, "lore.world reviews only on sets",
             f"{int((wc.scent == 'Set').sum())} reviews appear only under sets and are left out of per-scent figures")

    # Labels: every label backed by a quote found in the review, and keyword tags as a second opinion
    lab = labels.assign(source=labels.uid.str.split(":").str[0], review_id=labels.uid.str.split(":", n=1).str[1])
    texts = pd.concat([s.assign(source="sephora")[["source", "review_id", "title", "text"]],
                       w.assign(source="lore_world")[["source", "review_id", "title", "text"]],
                       f.assign(source="fragrantica", title="")[["source", "review_id", "title", "text"]]])
    lab = lab.merge(texts, on=["source", "review_id"], how="left")
    rep.check(G, "Every review has a full-read label", lab.text.notna().all() and len(lab) == len(texts),
              f"{len(lab)} labels for {len(texts)} reviews")
    pos = lab[(lab.fades.astype(int) == 1) | (lab.smoke.astype(int) == 1)]
    alnum = lambda x: re.sub(r"[^a-z0-9]", "", low(x))
    def quote_found(r):
        body = alnum(r.title + " . " + r.text)
        parts = [alnum(p) for p in re.split(r" / |; |\.\.\.|…", str(r.quote)) if alnum(p)]
        words = lambda x: set(re.findall(r"[a-z0-9]+", low(x)))
        if bool(parts) and all(p[:30] in body for p in parts):
            return True
        # allow light paraphrase: most of the quote's words appear in the review
        qw, bw = words(r.quote), words(r.title + " " + r.text)
        return len(qw) > 0 and len(qw & bw) / len(qw) >= 0.85
    found = pos.apply(quote_found, axis=1)
    rep.check(G, "Every positive label has a supporting quote found in the review", found.all(),
              f"{int(found.sum())} of {len(pos)} quotes found word for word")
    for _, r in pos[~found].iterrows():
        rep.flag(r.source, r.review_id, "Label quote not found", f"quote '{r.quote}' not found in the review text")
    for src, d in [("sephora", sc), ("lore_world", wc)]:
        d = d.copy()
        dis = d[d.fades.astype(bool) != d.fades_read.astype(bool)]
        for _, r in dis.iterrows():
            who = "full read says yes, keyword tag says no" if r.fades_read else "keyword tag says yes, full read says no"
            rep.flag(src, r.review_id, "Fading label", f"{who}; borderline, confirm by reading")
        if "cigarette" in d:
            dis = d[d.cigarette.astype(bool) != d.smoke_read.astype(bool)]
            for _, r in dis.iterrows():
                who = "full read says yes, keyword tag says no" if r.smoke_read else "keyword tag says yes, full read says no"
                rep.flag(src, r.review_id, "Smoke label", f"{who}; confirm by reading")
    # Reviews where a fading word appears but neither label fired
    tt = (s.title + " . " + s.text).map(low)
    quiet = s[tt.str.contains(FADES_KW) & ~sc.set_index("review_id").reindex(s.review_id).fades_read.astype(bool).values
              & ~sc.set_index("review_id").reindex(s.review_id).fades.astype(bool).values]
    for _, r in quiet.iterrows():
        rep.flag("sephora", r.review_id, "Fading label",
                 "mentions longevity, projection or fading but both labels say no; likely positive, confirm", "Info")


def article_numbers(sc, wc, fc):
    ss, ws = sc[sc.scent != "Set"], wc[wc.scent != "Set"]
    paid, gift = ss[~ss.gifted], ss[ss.gifted]
    p4 = paid[paid.scent != "Lil Softie"]
    b, a = p4[p4.date < "2026-04-01"], p4[p4.date >= "2026-04-01"]
    low_d = ss[(ss.rating <= 3) & (ss.scent == "Disfruta")]
    pct = lambda x: x.mean() * 100
    return {
        "sephora_avg": sc.rating.mean(), "lore_avg": wc.rating.mean(),
        "paid_fades": pct(paid.fades_read), "gift_fades": pct(gift.fades_read), "lore_fades": pct(ws.fades_read),
        "stars_fades": ss[ss.fades_read == 1].rating.mean(), "stars_other": ss[ss.fades_read == 0].rating.mean(),
        "trend_before": pct(b.fades_read), "trend_since": pct(a.fades_read),
        "trend_before_n": len(b), "trend_since_n": len(a),
        "gifted_share": pct(sc.gifted), "gifted_apr": int(sc[sc.gifted & sc.date.str.startswith("2026-04")].shape[0]),
        "gift_stars": gift.rating.mean(), "paid_stars": paid.rating.mean(),
        "disfruta_stars": ss[ss.scent == "Disfruta"].rating.mean(), "disfruta_smoke": pct(low_d.smoke_read),
        "n_total": len(sc) + len(wc), "n_sephora": len(sc), "n_lore": len(wc), "n_frag": len(fc),
    }


def check_article(rep, s, sc, wc, fc):
    G = "D. Article numbers"
    sc = sc.copy()
    nums = article_numbers(sc, wc, fc)
    for label, (key, stated, tol) in ARTICLE_CLAIMS.items():
        val = nums[key]
        rep.check(G, label, abs(val - stated) <= tol, f"article says {stated}, data gives {val:.2f}")

    # Sensitivity: does the main finding survive the uncertain cases?
    raw = s.set_index("review_id")
    blank_ids = set(raw[raw.incentivized == ""].index)
    flagged_gift = set(raw[(raw.incentivized.str.lower() != "true") &
                           (raw.title + " . " + raw.text).map(low).str.contains(GIFT_TEXT)].index)
    ss = sc[sc.scent != "Set"]
    def gap(d):
        return (d[~d.gifted].fades_read.mean() * 100, d[d.gifted].fades_read.mean() * 100)
    p, g = gap(ss[~ss.review_id.isin(blank_ids)])
    rep.check(G, "Paid vs gifted gap holds without the blank-gifted reviews", p > g,
              f"paying {p:.0f}% vs gifted {g:.0f}% using only reviews that answered the question")
    alt = ss.copy()
    alt.loc[alt.review_id.isin(flagged_gift), "gifted"] = True
    p, g = gap(alt)
    rep.check(G, "Paid vs gifted gap holds if text-disclosed gifts are treated as gifted", p > g,
              f"paying {p:.0f}% vs gifted {g:.0f}%")

    # The gap within the same scent and quarter, so a shift in which scents got reviewed can't explain it
    cells = ss.assign(cell=ss.scent + "|" + pd.to_datetime(ss.date).dt.to_period("Q").astype(str))
    by = cells.groupby(["cell", "gifted"]).fades_read.agg(["mean", "size"]).unstack().dropna()
    wts = by[("size", False)]
    p, g = ((by[("mean", flag)] * wts).sum() / wts.sum() * 100 for flag in (False, True))
    rep.check(G, "Paid vs gifted gap holds within the same scent and quarter", p > g,
              f"paying {p:.0f}% vs gifted {g:.0f}% across {len(by)} scent-quarter groups")


def label_table(sc, wc, fc, labels):
    """What the analysis currently says about each review: scent, size, gifted, and both kinds of label."""
    yn = lambda x: "" if pd.isna(x) else ("Yes" if bool(x) else "No")
    q = labels.assign(source=labels.uid.str.split(":").str[0], review_id=labels.uid.str.split(":", n=1).str[1])
    parts = [
        sc.assign(source="sephora", size="", kw_f=sc.fades, kw_s=sc.cigarette),
        wc.assign(source="lore_world", size="", kw_f=wc.fades, kw_s=wc.cigarette),
        fc.assign(source="fragrantica", size="", gifted=pd.NA, kw_f=pd.NA, kw_s=pd.NA),
    ]
    out = pd.concat([d[["source", "review_id", "scent", "size", "gifted", "fades_read", "smoke_read", "kw_f", "kw_s"]]
                     for d in parts])
    out = out.merge(q[["source", "review_id", "quote"]], on=["source", "review_id"], how="left")
    return pd.DataFrame({
        "source": out.source, "review_id": out.review_id.astype(str), "a_scent": out.scent, "a_size": out["size"],
        "a_gifted": out.gifted.map(yn),
        "a_fades": out.fades_read.map(lambda x: "" if pd.isna(x) else int(x)),
        "a_smoke": out.smoke_read.map(lambda x: "" if pd.isna(x) else int(x)),
        "k_fades": out.kw_f.map(lambda x: "" if pd.isna(x) else int(bool(x))),
        "k_smoke": out.kw_s.map(lambda x: "" if pd.isna(x) else int(bool(x))),
        "quote": out.quote.fillna("")})


def write_outputs(rep, s, w, f, sc, wc, fc, labels):
    VER.mkdir(exist_ok=True)
    checks = pd.DataFrame(rep.checks)
    flags = pd.DataFrame(rep.flags).drop_duplicates()
    flags.to_csv(VER / "flags.csv", index=False)

    lines = ["# Verification report", "",
             f"{(checks.result == 'PASS').sum()} checks passed, {(checks.result == 'FAIL').sum()} failed, "
             f"{(checks.result == 'NOTE').sum()} notes. "
             f"{flags[flags.level == 'Needs human review'].review_id.nunique() if len(flags) else 0} reviews need human review; "
             f"{flags[flags.level == 'Info'].review_id.nunique() if len(flags) else 0} more carry an informational note.", ""]
    for group, d in checks.groupby("group", sort=True):
        lines += [f"## {group}", "", "| Result | Check | Detail |", "|---|---|---|"]
        lines += [f"| {r.result} | {r.check} | {r.detail} |" for r in d.itertuples()]
        lines.append("")
    if len(flags):
        lines += ["## Flags by type", "", "| Check | Level | Reviews |", "|---|---|---|"]
        for (c, lv), n in flags.groupby(["check", "level"]).review_id.nunique().sort_values(ascending=False).items():
            lines.append(f"| {c} | {lv} | {n} |")
    (VER / "verification_report.md").write_text("\n".join(lines))

    # Workbook of flagged reviews, full text beside the reason, for a person to decide
    texts = pd.concat([
        s.assign(source="sephora", where=s.product_name)[["source", "review_id", "where", "date", "rating", "title", "text"]],
        w.assign(source="lore_world", where=w.listings)[["source", "review_id", "where", "date", "rating", "title", "text"]],
        f.assign(source="fragrantica", where=f.scent, rating="", title="")[["source", "review_id", "where", "date", "rating", "title", "text"]],
    ])
    if not len(flags):
        return checks, flags
    grouped = flags.groupby(["source", "review_id"]).agg(
        level=("level", lambda x: "Needs human review" if (x == "Needs human review").any() else "Info"),
        checks=("check", lambda x: "; ".join(sorted(set(x)))),
        details=("detail", lambda x: " | ".join(x))).reset_index()
    out = grouped.merge(texts, on=["source", "review_id"], how="left")
    out = out.merge(label_table(sc, wc, fc, labels), on=["source", "review_id"], how="left")
    # Most consequential first: things that could move a headline number, then label borderlines
    priority = ["Review no longer on site", "differs from live site", "Gifted flag", "Scent",
                "Duplicate", "Label quote", "Smoke label", "Stars vs wording", "Fading label", "New review"]
    rank = lambda c: min([i for i, p in enumerate(priority) if p in c] or [len(priority)])
    out["order"] = out.level.map({"Needs human review": 0, "Info": 1})
    out["rank"] = out.checks.map(rank)
    out = out.sort_values(["order", "rank", "source", "date"]).drop(columns=["order", "rank"])
    cols = ["level", "checks", "details", "source", "review_id", "where", "date", "rating", "title", "text",
            "a_scent", "a_gifted", "a_fades", "a_smoke", "quote", "k_fades", "k_smoke"]
    heads = ["Level", "Why flagged", "Details", "Source", "Review ID", "Listing / scent", "Date", "Stars", "Title",
             "Review text", "Counted as scent", "Counted as gifted", "Fades / too faint (used)",
             "Cigarette-type smoke (used)", "Supporting quote", "Fades (keyword tag)", "Smoke (keyword tag)",
             "Your decision", "Your notes"]
    label_cols = {"Counted as scent", "Counted as gifted", "Fades / too faint (used)",
                  "Cigarette-type smoke (used)", "Supporting quote"}
    wb = Workbook()
    ws = wb.active
    ws.title = "Flagged reviews"
    ws.append(heads)
    for r in out[cols].itertuples(index=False):
        ws.append([("" if pd.isna(v) else v) for v in r] + ["", ""])
    widths = {"Why flagged": 26, "Details": 48, "Review text": 80, "Title": 26, "Listing / scent": 30,
              "Supporting quote": 34, "Fades / too faint (used)": 11, "Cigarette-type smoke (used)": 11,
              "Fades (keyword tag)": 10, "Smoke (keyword tag)": 10, "Counted as gifted": 10,
              "Counted as scent": 12}
    blue = PatternFill("solid", fgColor="E7F1FB")
    hf, body, yel = Font(name="Arial", size=10, bold=True, color="FFFFFF"), Font(name="Arial", size=10), PatternFill("solid", fgColor="FFF2CC")
    for i, h in enumerate(heads, 1):
        c = ws.cell(1, i)
        c.font, c.fill = hf, PatternFill("solid", fgColor="2F3B52")
        c.alignment = Alignment(wrap_text=True, vertical="center")
        ws.column_dimensions[get_column_letter(i)].width = widths.get(h, 13)
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.font, c.alignment = body, Alignment(wrap_text=True, vertical="top")
        row[-2].fill = row[-1].fill = yel
        for i, h in enumerate(heads):
            if h in label_cols:
                row[i].fill = blue
    dcol = get_column_letter(heads.index("Your decision") + 1)
    dv = DataValidation(type="list", formula1='"Fine as is,Fix the data,Fix the label,Exclude review,Unsure"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"{dcol}2:{dcol}{ws.max_row}")
    ws.freeze_panes, ws.auto_filter.ref = "C2", f"A1:{get_column_letter(len(heads))}{ws.max_row}"
    guide = wb.create_sheet("How to use", 0)
    for line in ["Flagged reviews: everything the verification script could not confirm on its own",
                 "",
                 "Each row is one review, with every reason it was flagged and its full text.",
                 "'Needs human review' rows could change a number; 'Info' rows are notes you can skim.",
                 "The blue columns show how the analysis currently counts each review: its scent and gifted status, "
                 "and the labels used in the article (1 = yes, 0 = no), with the quote behind each yes.",
                 "The keyword tags beside them are the earlier automatic tags, shown for comparison only.",
                 "Pick a value in the yellow 'Your decision' column and add a note if useful.",
                 "Fix the label = a fading or smoke label is wrong. Fix the data = the scent or gifted status is wrong. "
                 "Exclude review = it shouldn't be counted (for example a duplicate post).",
                 "Example: a review listed under Lil Softie that only talks about Sublimity -> Fix the data.",
                 "Rerun python analysis/verify_reviews.py after fixing anything; this file is rebuilt each run.",
                 "",
                 f"Flagged reviews needing a decision: {(out.level == 'Needs human review').sum()}",
                 f"Informational: {(out.level == 'Info').sum()}"]:
        guide.append([line])
    for row in guide.iter_rows():
        for c in row:
            c.font = Font(name="Arial", size=10)
    guide["A1"].font = Font(name="Arial", size=13, bold=True)
    guide.column_dimensions["A"].width = 110
    wb.save(PRIV / "needs_human_review.xlsx")
    return checks, flags


def main():
    rep = Report()
    s, w, f, sc, wc, fc, labels = load()
    check_sephora(rep, s, sc)
    check_lore(rep, w)
    check_fragrantica(rep, f)
    check_consistency(rep, s, w, f, sc, wc, labels)
    check_article(rep, s, sc, wc, fc)
    checks, flags = write_outputs(rep, s, w, f, sc, wc, fc, labels)
    print(f"{(checks.result == 'PASS').sum()} passed, {(checks.result == 'FAIL').sum()} failed, "
          f"{(checks.result == 'NOTE').sum()} notes")
    for r in checks[checks.result == "FAIL"].itertuples():
        print(f"  FAIL  {r.check}: {r.detail}")
    if len(flags):
        print(flags.groupby(["level", "check"]).review_id.nunique().to_string())
    print("Report: private/verification/verification_report.md · Review list: private/needs_human_review.xlsx")


if __name__ == "__main__":
    main()
