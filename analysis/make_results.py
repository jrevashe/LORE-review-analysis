"""
Step 3: turn the coded review files into the numbers and charts used in the article.

Run from the project folder after code_reviews.py:
    pip install pandas matplotlib
    python analysis/make_results.py

Writes results.md and the charts/ folder.
"""
from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
SCENTS = ["Sublimity", "Lil Softie", "Somewhere", "Lovely", "Disfruta"]
FULL_NAME = {"Lovely": "Lovely and a little twisted", "Somewhere": "Somewhere but nowhere"}
TEXT, MUTED, GRID = "#0b0b0b", "#52514e", "#e6e5e1"
PAID, LORE_SITE, GIFTED = "#eb6834", "#1baf7a", "#2a78d6"


def pct(x):
    return f"{x * 100:.0f}%"


def table(df: pd.DataFrame) -> str:
    cols = [str(c) for c in df.columns]
    rows = ["| " + " | ".join([df.index.name or ""] + cols) + " |", "|" + "---|" * (len(cols) + 1)]
    for idx, r in df.iterrows():
        rows.append("| " + " | ".join([str(idx)] + [str(v) for v in r.values]) + " |")
    return "\n".join(rows)


def style(ax, title):
    ax.set_title(title, loc="left", fontsize=12, color=TEXT, pad=32)
    for side in ("top", "right"):
        ax.spines[side].set_visible(False)
    ax.tick_params(colors=MUTED, length=0)


def main():
    out = ROOT / "charts"
    out.mkdir(exist_ok=True)
    s = pd.read_csv(ROOT / "data/sephora_reviews_coded.csv", parse_dates=["date"])
    w = pd.read_csv(ROOT / "data/lore_world_reviews_coded.csv", parse_dates=["date"])
    f = pd.read_csv(ROOT / "data/fragrantica_reviews_coded.csv", parse_dates=["date"])
    # The article uses the labels from reading every review in full; keyword tags are kept to compare.
    for d in (s, w):
        d["fades_keyword"], d["cigarette_keyword"] = d.fades, d.cigarette
        d["fades"], d["cigarette"] = d.fades_read.astype(bool), d.smoke_read.astype(bool)
    f["weak"] = f.fades_read.astype(bool)
    ss, ws = s[s.scent != "Set"], w[w.scent != "Set"]
    paid, gifted = ss[~ss.gifted], ss[ss.gifted]

    def agreement(d, a, b):
        return f"{(d[a] == d[b]).mean():.1%} agreement ({int(d[b].sum())} read vs {int(d[a].sum())} keyword)"

    lines = ["# Results", "",
             f"Generated from {len(s):,} Sephora reviews, {len(w):,} lore.world reviews and "
             f"{len(f):,} Fragrantica reviews. Every review was read in full and labeled; those labels drive "
             "every number below.", "",
             "## How well the keyword tags matched the full read", "",
             f"- Sephora, fades or too faint: {agreement(s, 'fades_keyword', 'fades')}",
             f"- Sephora, cigarette or similar smoke: {agreement(s, 'cigarette_keyword', 'cigarette')}",
             f"- lore.world, fades or too faint: {agreement(w, 'fades_keyword', 'fades')}",
             "- The keyword tags mostly missed milder complaints (\"longevity is medium\", \"had to reapply "
             "later in the day\"), so the full-read percentages run a few points higher.", ""]

    # Who writes Lore's Sephora reviews
    lines += ["## Sephora: who is reviewing", "",
              f"- Reviews: {len(s):,}, from {s.date.min():%b %d, %Y} to {s.date.max():%b %d, %Y}",
              f"- Received the product free: {pct(s.gifted.mean())}",
              f"- Average stars, single scents: gifted {gifted.rating.mean():.2f}, paying {paid.rating.mean():.2f}", ""]
    monthly = s.assign(month=s.date.dt.to_period("M")).groupby("month").agg(
        reviews=("review_id", "size"), gifted_share=("gifted", "mean"))
    monthly["gifted_share"] = (monthly.gifted_share * 100).round(0).astype(int).astype(str) + "%"
    lines += [table(monthly), ""]

    # Fade complaints by scent and source
    comp = pd.DataFrame({
        "Sephora, paying": paid.groupby("scent").fades.mean(),
        "Sephora, gifted": gifted.groupby("scent").fades.mean(),
        "lore.world buyers": ws.groupby("scent").fades.mean(),
        "Fragrantica": f.groupby("scent").weak.mean(),
    }).reindex(SCENTS)
    counts = pd.DataFrame({
        "Sephora, paying": paid.groupby("scent").size(), "Sephora, gifted": gifted.groupby("scent").size(),
        "lore.world buyers": ws.groupby("scent").size(), "Fragrantica": f.groupby("scent").size()}).reindex(SCENTS)
    shown = comp.map(pct) + " (n=" + counts.astype(str) + ")"
    shown.index.name = "Scent"
    lines += ["## Share of reviews saying the scent fades or is too faint", "", table(shown), "",
              f"- All single-scent reviews (sets left out), paying Sephora customers: {pct(paid.fades.mean())}; "
              f"gifted: {pct(gifted.fades.mean())}; lore.world: {pct(ws.fades.mean())}",
              f"- Average stars when a review mentions fading: {ss[ss.fades].rating.mean():.2f}, "
              f"otherwise {ss[~ss.fades].rating.mean():.2f}", ""]

    # Robustness: compare paying and gifted reviews within the same scent and quarter
    cells = ss.assign(cell=ss.scent + "|" + ss.date.dt.to_period("Q").astype(str))
    by = cells.groupby(["cell", "gifted"]).fades.agg(["mean", "size"]).unstack().dropna()
    wts = by[("size", False)]
    matched_paid = (by[("mean", False)] * wts).sum() / wts.sum()
    matched_gift = (by[("mean", True)] * wts).sum() / wts.sum()
    lines += ["## Check: same scent and quarter", "",
              "Gifted reviews re-weighted to match the paying reviews' mix of scent and quarter "
              f"({len(by)} groups with both kinds of review): paying {pct(matched_paid)}, gifted {pct(matched_gift)}.", ""]

    # Trend for the original four scents only (Lil Softie launched in 2026 and dominates later reviews)
    p4 = paid[paid.scent != "Lil Softie"]
    before, since = p4[p4.date < "2026-04-01"], p4[p4.date >= "2026-04-01"]
    lines += ["## Check: trend without Lil Softie", "",
              f"Paying reviews of the original four scents mentioning fading: {pct(before.fades.mean())} before "
              f"April 2026 ({int(before.fades.sum())} of {len(before)}), {pct(since.fades.mean())} since "
              f"({int(since.fades.sum())} of {len(since)}). The quarterly table below mixes in Lil Softie, which made up "
              "most paying reviews in July to September 2026.", ""]

    q = paid.assign(quarter=paid.date.dt.to_period("Q")).groupby("quarter").agg(
        reviews=("review_id", "size"), avg_stars=("rating", "mean"), fades=("fades", "mean"))
    q["avg_stars"] = q.avg_stars.round(2)
    q["fades"] = q.fades.map(pct)
    lines += ["## Paying Sephora customers over time", "", table(q), ""]

    # Sensitivity: does the headline survive different ways of handling the uncertain cases?
    raw = pd.read_csv(ROOT / "private/sephora_reviews_raw.csv", dtype={"review_id": str}, keep_default_na=False) \
        if (ROOT / "private/sephora_reviews_raw.csv").exists() else None
    def gap(d, col="fades"):
        return f"{pct(d[~d.gifted][col].mean())} vs {pct(d[d.gifted][col].mean())}"
    rows = [("Full-read labels (used in the article)", gap(ss)),
            ("Keyword tags instead of the full read", gap(ss, "fades_keyword")),
            ("Leaving out the 121 reviews where the gifted question was blank", gap(ss[~ss.gifted_flag_blank]))]
    flags_path = ROOT / "private/verification/flags.csv"
    if flags_path.exists():
        fl = pd.read_csv(flags_path, dtype={"review_id": str})
        unsure = set(fl[(fl.source == "sephora") & (fl.level == "Needs human review")].review_id)
        rows.append(("Leaving out every review the verification script flagged as borderline",
                     gap(ss[~ss.review_id.astype(str).isin(unsure)])))
    sens = pd.DataFrame(rows, columns=["Version", "Paying vs gifted, mentioning fading"]).set_index("Version")
    lines += ["## Sensitivity: paying vs gifted under different choices", "",
              "Borderline reviews (a neutral \"lasts about 4 hours\", say) take judgment, so the headline is "
              "recomputed several ways. Paying reviewers mention fading about twice as often in each.", "",
              table(sens), ""]

    low = ss[ss.rating <= 3].groupby("scent").agg(reviews=("review_id", "size"), cigarette=("cigarette", "mean"),
                                                   fades=("fades", "mean")).reindex(SCENTS)
    low[["cigarette", "fades"]] = low[["cigarette", "fades"]].map(pct)
    low.index.name = "Scent (3 stars or below)"
    lines += ["## What 3-star-and-below Sephora reviews mention", "", table(low), ""]
    (ROOT / "results.md").write_text("\n".join(lines))

    # Chart 1: fade complaints by scent and buyer group
    order = paid.groupby("scent").fades.mean().sort_values().index
    groups = [("Paying customers on Sephora", paid, PAID), ("Buyers on lore.world", ws, LORE_SITE),
              ("Got it free on Sephora", gifted, GIFTED)]
    fig, ax = plt.subplots(figsize=(8.5, 5.2))
    y, h = np.arange(len(order)), 0.26
    for k, (label, d, color) in enumerate(groups):
        v = d.groupby("scent").fades.mean().reindex(order) * 100
        n = d.groupby("scent").size().reindex(order)
        off = (1 - k) * h
        ax.barh(y + off, v, height=h, color=color, label=label, edgecolor="white", linewidth=2)
        for i, (vv, nn) in enumerate(zip(v, n)):
            ax.text(vv + 0.4, i + off, f"{vv:.0f}%  (n={nn})", va="center", fontsize=7.5, color=MUTED)
    ax.set_yticks(y, [FULL_NAME.get(x, x) for x in order])
    ax.set_xticks([]); ax.set_xlim(0, 30)
    ax.spines["bottom"].set_visible(False); ax.spines["left"].set_color(GRID)
    style(ax, "Reviews saying the scent fades or is too faint")
    ax.legend(frameon=False, labelcolor=MUTED, fontsize=8.5, loc="lower left", bbox_to_anchor=(0, 1.0), ncol=3)
    fig.text(0.01, 0.01, f"Source: {len(ss):,} Sephora and {len(ws):,} lore.world reviews of LORE's five scents",
             fontsize=7.5, color=MUTED)
    fig.tight_layout(rect=(0, 0.04, 1, 1)); fig.savefig(out / "fades_by_scent_and_buyer.png", dpi=200); plt.close(fig)

    # Chart 2: Sephora reviews per month, gifted vs paying
    m = s.assign(month=s.date.dt.to_period("M")).groupby(["month", "gifted"]).size().unstack(fill_value=0)
    fig, ax = plt.subplots(figsize=(8.5, 4.2))
    x = np.arange(len(m))
    ax.bar(x, m.get(False, 0), width=0.65, color=PAID, label="Paying customers", edgecolor="white", linewidth=2)
    ax.bar(x, m.get(True, 0), width=0.65, bottom=m.get(False, 0), color=GIFTED, label="Got it free",
           edgecolor="white", linewidth=2)
    ax.set_xticks(x, [p.strftime("%b %y") for p in m.index.to_timestamp()], rotation=45, ha="right")
    ax.spines["left"].set_visible(False); ax.spines["bottom"].set_color(GRID)
    ax.yaxis.grid(True, color=GRID); ax.set_axisbelow(True)
    style(ax, "LORE reviews on Sephora per month")
    ax.legend(frameon=False, labelcolor=MUTED, fontsize=9, loc="lower left", bbox_to_anchor=(0, 1.0), ncol=2)
    fig.tight_layout(); fig.savefig(out / "sephora_reviews_per_month.png", dpi=200); plt.close(fig)

    # Chart 3: Fragrantica fade complaints by scent
    g = f.groupby("scent").agg(n=("weak", "size"), weak=("weak", "mean")).sort_values("weak")
    fig, ax = plt.subplots(figsize=(8, 3.6))
    colors = [GIFTED if v > 0.2 else "#a8c6ec" for v in g.weak]
    ax.barh([FULL_NAME.get(x, x) for x in g.index], g.weak * 100, color=colors, height=0.55)
    for i, (v, n) in enumerate(zip(g.weak, g.n)):
        ax.text(v * 100 + 0.8, i, f"{v * 100:.0f}%  (of {n} reviews)", va="center", fontsize=9, color=MUTED)
    ax.set_xticks([]); ax.set_xlim(0, 42)
    ax.spines["bottom"].set_visible(False); ax.spines["left"].set_color(GRID)
    style(ax, "Fragrantica reviews saying the scent is faint or fades fast")
    fig.tight_layout(); fig.savefig(out / "fragrantica_fades_by_scent.png", dpi=200); plt.close(fig)
    print("Wrote results.md and charts/")


if __name__ == "__main__":
    main()
