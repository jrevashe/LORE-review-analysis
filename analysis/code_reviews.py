"""
Step 2: tag every review with themes.

Two kinds of tags end up in data/:
- fades_read / smoke_read: from reading every review in full (data/full_read_labels.csv).
  These are what the article uses.
- fades, cigarette and the other keyword columns: quick pattern matches, kept as a cross-check.

Reads the raw review files in private/ (they contain full review text, so they stay
out of the public repo) and writes coded files to data/ that keep only dates, stars,
gifted/verified flags and the theme tags.

Run from the project folder:
    pip install pandas
    python analysis/code_reviews.py
"""
from pathlib import Path
import re
import warnings

import pandas as pd

warnings.filterwarnings("ignore", "This pattern is interpreted as a regular expression")
ROOT = Path(__file__).resolve().parent.parent

# Each theme is a regular expression matched against "title . text", lowercased,
# with curly apostrophes straightened. Checked by reading random samples of matches.
THEMES = {
    # says the scent fades, is faint, or doesn't last
    "fades": (
        r"(doesn't|does not|didn't|did not|won't|wouldn't|don't) (seem to |really |even )?last"
        r"|fades? (so |really |very |too )?(fast|quick|quickly|away)|fades? (after|within|in) "
        r"|faded (away|fast|quickly|within|after)|disappear|vanish"
        r"|(barely|can't|cannot|could not|couldn't|can barely) (even )?(smell|detect)|\bweak\b|\bfaint\b"
        r"|(poor|bad|terrible|awful|no|zero|short|low|little) (longevity|projection|sillage|staying power)"
        r"|longevity (is|was) (poor|bad|terrible|not|lacking|short|meh|awful)|short[- ]lived"
        r"|not (very )?long[- ]?lasting"
        r"|wish(ed)? (it|this) (lasted|was stronger|were stronger|would last|stayed|projected|had more)"
        r"|(only|just) lasts? (about |like |maybe )?(an? |\d+|a few|one|two|couple)"
        r"|gone (in|after|within)|(dies|died) (down|off)|wears? off"
        r"|wish(ed)? (it|they|she|this) (\w+ )?(lasted|projected|stayed|lingered|were stronger|was stronger"
        r"|had (more|longer|better)|would linger|would last)"
        r"|(concern|complaint|downside|drawback|con) (is|was) (the )?(longevity|projection|staying power)"
        r"|improvement would be (better )?longevity|longevity (could|can) be (much )?better"
        r"|disappeared (soon|quickly|fast|after|within)"
    ),
    # praises how long it lasts
    "lasts_long": (
        r"long[- ]?lasting|lasts? all day|lasted all day|lasts? (for )?(hours|a long)|\b\d{1,2}\+? ?(hours|hrs)\b"
        r"|(great|good|amazing|insane|impressive|excellent|incredible|decent) (longevity|staying power|projection|sillage)"
        r"|longevity is (great|amazing|good|insane|impressive|decent)|still smell it"
    ),
    "skin_scent": r"skin scent|close to (the |my )?skin|intimate",
    "cigarette": r"cigarette|ashtray|ash tray|\bbbq\b|barbecue|stale smoke|burnt rubber",
    "story_memory": r"remind|memor|nostalg|transport|vacation|escape|childhood|takes me|feel like i'm|brings me back|takes you",
    "price": r"overpriced|too expensive|not worth|pricey|price point|for the price",
}


def scent_of(name: str) -> str:
    n = str(name).lower()
    for key, label in [("lil softie", "Lil Softie"), ("sublimity", "Sublimity"), ("lovely", "Lovely"),
                       ("disfruta", "Disfruta"), ("somewhere", "Somewhere")]:
        if key in n and "set" not in n:
            return label
    return "Set"


def tag(df: pd.DataFrame) -> pd.DataFrame:
    text = (df["title"].fillna("") + " . " + df["text"].fillna("")).str.lower()
    text = text.str.replace("’", "'").str.replace("‘", "'")
    for theme, pattern in THEMES.items():
        df[theme] = text.str.contains(pattern, regex=True)
    df["review_length"] = df["text"].fillna("").str.len()
    return df


def add_read_labels(df, source):
    """Attach the labels from reading every review in full (data/full_read_labels.csv)."""
    lab = pd.read_csv(ROOT / "data/full_read_labels.csv", dtype={"uid": str})
    fixes = pd.read_csv(ROOT / "data/label_corrections.csv", dtype=str)
    for fx in fixes.itertuples():
        lab.loc[lab.uid == fx.uid, fx.field] = int(fx.value)
    lab = lab[lab.uid.str.startswith(source + ":")].assign(review_id=lambda d: d.uid.str.split(":", n=1).str[1])
    lab = lab.rename(columns={"fades": "fades_read", "smoke": "smoke_read"})[["review_id", "fades_read", "smoke_read"]]
    return df.merge(lab, on="review_id", how="left")


def main():
    s = pd.read_csv(ROOT / "private/sephora_reviews_raw.csv", dtype=str, keep_default_na=False)
    s["scent"] = s["product_name"].map(scent_of)
    s["travel_size"] = s["product_name"].str.contains("Travel", case=False)
    # Gifted = Sephora's "received this product free" flag, or Sephora's own disclosure line in the text.
    # 31 reviews carry the disclosure line with the flag blank or "no" (found by verify_reviews.py).
    disclosed = s["text"].str.contains(r"\[I received this product in exchange for my honest review\]", case=False)
    s["gifted"] = s["incentivized"].str.lower().eq("true") | disclosed
    s["gifted_flag_blank"] = s["incentivized"].eq("")
    s = tag(s)
    keep = ["review_id", "product_id", "product_name", "scent", "travel_size", "date", "rating", "recommended",
            "gifted", "gifted_flag_blank", "helpful_yes", "helpful_no", "review_length", *THEMES]
    add_read_labels(s[keep], "sephora").to_csv(ROOT / "data/sephora_reviews_coded.csv", index=False)

    singles = {"Sublimity", "Somewhere", "Lovely", "Disfruta", "Lil Softie"}
    w = pd.read_csv(ROOT / "private/lore_world_reviews_raw.csv", dtype=str, keep_default_na=False)
    w["scent"] = w["listings"].map(lambda f: next((x for x in f.split("|") if x in singles), "Set"))
    w["gifted"] = w["incentivized"].str.lower().eq("true")
    w = tag(w)
    keep = ["review_id", "listings", "scent", "date", "rating", "verified", "gifted", "review_length", *THEMES]
    add_read_labels(w[keep], "lore_world").to_csv(ROOT / "data/lore_world_reviews_coded.csv", index=False)

    f = pd.read_csv(ROOT / "private/fragrantica_reviews_raw.csv", dtype={"review_id": str})
    add_read_labels(f[["review_id", "scent", "date", "tags"]], "fragrantica").to_csv(
        ROOT / "data/fragrantica_reviews_coded.csv", index=False)
    print(f"Coded {len(s)} Sephora, {len(w)} lore.world and {len(f)} Fragrantica reviews into data/")


if __name__ == "__main__":
    main()
