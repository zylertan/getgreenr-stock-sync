"""
matcher.py — Build the SKU registry by matching GetGreenr listings to the
Mister Mobile Masterlist (POS stock report).

Match key = (brand, model_family, capacity, color), all normalized.

Tabs produced:
  Locked Matches   exact key match; Seller Stock = SUM of Available Qty of all linked ML rows
  Match Review     model+capacity match but color differs, OR fuzzy model — needs human eyes
  Skipped (0 Qty)  locked match whose summed qty == 0  -> Seller Stock 0
  Accessories      GetGreenr AirPods/Watch/Buds/etc not matched -> Seller Stock 10
  Not on GetGreenr Yet   Masterlist products (qty>0) with no GetGreenr listing
"""
import re
import pandas as pd
from collections import defaultdict

UP = "/sessions/dreamy-hopeful-mayer/mnt/uploads"
GG_PATH = f"{UP}/GetGreenr Inventory - inventory (1).csv.csv"
ML_PATH = f"{UP}/stock_report (18).xlsx"

# --------------------------------------------------------------------------- #
# Normalisation helpers
# --------------------------------------------------------------------------- #
CAP_RE = re.compile(r"\b(\d+)\s*(TB|GB)\b", re.I)
RAM_RE = re.compile(r"\b\d+\s*GB\s*RAM\b", re.I)
SLASHRAM_RE = re.compile(r"/\s*\d+")           # "128GB/4" -> drop "/4"
NET_RE = re.compile(r"\b[45]G\b", re.I)
SIM_RE = re.compile(r"\b(\d+\s*)?(physical\s*)?(e?sim)\b", re.I)
CODE_RE = re.compile(r"\b[A-Z]{1,3}\d{3,}[A-Z]?\b", re.I)  # "A075F"/"G998B"/"F766"/"CPH2573"
REGION_RE = re.compile(r"\b(kr|hk|ind|us|eu|sg|jp|cn|tw|global|choice)\b", re.I)
PLUS_RE = re.compile(r"\bplus\b", re.I)
NOISE = ["1 physical sim + esim", "2 physical sim", "1 physical sim",
         "physical sim", "esim", "wifi", "wi-fi", "cell", "cellular",
         "(5g)", "5g", "4g", "dual sim"]

BRAND_MAP = {  # GetGreenr first word -> Masterlist Brand bucket
    "samsung": "SAMSUNG", "google": "GOOGLE", "oneplus": "ONE PLUS",
    "honor": "HONOR", "xiaomi": "XIAOMI", "nothing": "NOTHING", "oppo": "OPPO",
    "sony": "SONY",
}
COLOR_SYNONYMS = {  # light touch; extend as needed
    "space grey": "space gray", "grey": "gray",
}


def norm(s):
    s = "" if s is None else str(s)
    s = s.replace("–", "-").replace("—", "-")
    s = s.lower().strip()
    s = re.sub(r"\s+", " ", s)
    return s


def extract_capacity(text):
    # NB: do NOT strip spaces first — "16 256GB" would merge into "16256GB".
    m = CAP_RE.search(text)
    if not m:
        return ""
    return f"{int(m.group(1))}{m.group(2).upper()}"


def clean_model(text):
    """Strip capacity, RAM, network, SIM, model codes, punctuation."""
    t = " " + text + " "
    t = RAM_RE.sub(" ", t)
    t = CAP_RE.sub(" ", t)
    t = SLASHRAM_RE.sub(" ", t)
    t = NET_RE.sub(" ", t)
    t = CODE_RE.sub(" ", t)        # after net/cap removal so "5g-s901e" -> code stripped
    t = SIM_RE.sub(" ", t)
    t = REGION_RE.sub(" ", t)
    for n in NOISE:
        t = t.replace(n, " ")
    t = re.sub(r"[()#]", " ", t)
    t = PLUS_RE.sub("+", t)         # "s24 plus" -> "s24+"
    t = re.sub(r"[^a-z0-9+ ]", " ", t)
    t = re.sub(r"\s*\+\s*", "+", t)  # collapse spaces around +
    t = re.sub(r"\s+", " ", t).strip()
    return t


def norm_color(c):
    c = norm(c)
    c = re.sub(r"#\s*\d+", " ", c)          # drop duplicate-listing markers "#2"
    c = re.sub(r"[^a-z0-9 ]", " ", c)
    c = re.sub(r"\b\d+\b\s*$", " ", c)       # drop trailing standalone digit "graphite 2"
    c = re.sub(r"\s+", " ", c).strip()
    return COLOR_SYNONYMS.get(c, c)


def mk(s):
    """Despace a model/colour for the match key: 'z flip 7' == 'z flip7'."""
    return re.sub(r"\s+", "", str(s))


# --------------------------------------------------------------------------- #
# Parse GetGreenr sku -> (brand, model_family, capacity, color)
# --------------------------------------------------------------------------- #
ACCESSORY_HINTS = ["airpod", "watch", "buds", "case", "charger", "cable",
                   "band", "pencil", "keyboard", "cover", "adapter"]


def parse_gg(sku, title):
    raw = norm(sku)
    segs = [s.strip() for s in raw.split(" - ") if s.strip()]
    first = segs[0] if segs else ""
    words = first.split()
    brand_word = words[0] if words else ""

    # Determine brand bucket
    if brand_word == "apple":
        rest = first[len("apple"):].strip()
        if rest.startswith("iphone"):
            brand = "IPHONE"; model_src = rest[len("iphone"):].strip()
        elif rest.startswith("ipad"):
            brand = "IPAD"; model_src = rest[len("ipad"):].strip()
        elif rest.startswith("watch"):
            brand = "APPLE WATCH"; model_src = rest[len("watch"):].strip()
        else:
            brand = "APPLE"; model_src = rest  # AirPods etc.
    else:
        brand = BRAND_MAP.get(brand_word, brand_word.upper())
        model_src = " ".join(words[1:])
        # Samsung: drop leading "galaxy"
        if model_src.startswith("galaxy"):
            model_src = model_src[len("galaxy"):].strip()
        # OnePlus etc keep number

    capacity = extract_capacity(raw)
    # Color = last segment that is not capacity/ram/sim/wifi
    color = ""
    for seg in reversed(segs[1:]):
        if CAP_RE.search(seg) or RAM_RE.search(seg) or "sim" in seg or "wifi" in seg:
            continue
        color = seg
        break
    model = clean_model(model_src)
    model = fix_model_aliases(brand, model)
    is_acc = any(h in raw for h in ACCESSORY_HINTS)
    return brand, model, capacity, norm_color(color), is_acc


def fix_model_aliases(brand, model):
    """Reconcile naming conventions that differ between GetGreenr and the Masterlist."""
    m = model
    # ML bakes brand into the model for some brands
    if brand == "ONE PLUS":
        m = re.sub(r"^(one\+|1\+|oneplus)\s*", "", m).strip()
    elif brand == "XIAOMI":
        m = re.sub(r"^xiaomi\s+", "", m).strip()
    # iPhone SE: GetGreenr uses release year, Masterlist uses generation number
    if brand == "IPHONE":
        m = re.sub(r"\bse 2016\b", "se 1", m)
        m = re.sub(r"\bse 2020\b", "se 2", m)
        m = re.sub(r"\bse 2022\b", "se 3", m)
    return m


# --------------------------------------------------------------------------- #
# Load & aggregate Masterlist
# --------------------------------------------------------------------------- #
def load_masterlist(ml_file=ML_PATH):
    raw = pd.read_excel(ml_file, header=None)
    df = raw.iloc[2:].copy()
    df.columns = (["StockTypeID", "Category", "Brand", "Model", "Color",
                   "Total", "Quantity", "Reserved", "TransitRes", "TransitNonRes"]
                  + [f"branch_{i}" for i in range(len(raw.columns) - 10)])
    df = df.reset_index(drop=True)
    df = df[df["StockTypeID"].notna()]

    rows = []
    for _, r in df.iterrows():
        brand = str(r["Brand"]).strip().upper()
        capacity = extract_capacity(norm(r["Model"]))
        model = fix_model_aliases(brand, clean_model(norm(r["Model"])))
        color = norm_color(r["Color"])
        try:
            qty = int(float(r["Total"]))
        except (ValueError, TypeError):
            qty = 0
        rows.append({
            "StockTypeID": r["StockTypeID"], "Category": r["Category"],
            "brand": brand, "model": model, "capacity": capacity,
            "color": color, "qty": max(qty, 0),
            "raw_model": str(r["Model"]).strip(), "raw_color": str(r["Color"]).strip(),
        })
    return pd.DataFrame(rows)


def key(b, m, c, col):
    return (b, mk(m), c, mk(col))


def build(gg_file=GG_PATH, ml_file=ML_PATH):
    name = getattr(gg_file, "name", str(gg_file)).lower()
    gg = pd.read_csv(gg_file) if name.endswith(".csv") else pd.read_excel(gg_file)
    ml = load_masterlist(ml_file)

    # GetGreenr grade "Excellent" == Masterlist "Used" condition, so GetGreenr
    # stock syncs ONLY against Used inventory. New rows are irrelevant here.
    ml = ml[ml["Category"].astype(str).str.strip().str.lower() == "used"].reset_index(drop=True)

    # Index masterlist by full key and by (brand, model, capacity) for review fallback
    ml_by_key = defaultdict(list)
    ml_by_mc = defaultdict(list)   # brand+model+capacity -> rows (any color)
    ml_by_m = defaultdict(list)    # brand+model -> rows
    for _, r in ml.iterrows():
        ml_by_key[key(r["brand"], r["model"], r["capacity"], r["color"])].append(r)
        ml_by_mc[(r["brand"], mk(r["model"]), r["capacity"])].append(r)
        ml_by_m[(r["brand"], mk(r["model"]))].append(r)

    locked, review, accessories = [], [], []
    matched_ml_keys = set()

    for _, g in gg.iterrows():
        brand, model, cap, color, is_acc = parse_gg(g["sku"], g["title"])
        k = key(brand, model, cap, color)
        base = {
            "gg_sku": g["sku"], "gg_title": g["title"],
            "brand": brand, "model": model, "capacity": cap, "color": color,
            "gg_on_hand": g.get("on_hand", ""), "gg_available": g.get("available", ""),
        }

        if k in ml_by_key:                       # exact match -> LOCKED
            linked = ml_by_key[k]
            total = sum(r["qty"] for r in linked)
            matched_ml_keys.add(k)
            rec = {**base,
                   "linked_StockTypeIDs": ", ".join(str(r["StockTypeID"]) for r in linked),
                   "linked_ml_models": " | ".join(f"{r['raw_model']} [{r['Category']}] {r['raw_color']}={r['qty']}" for r in linked),
                   "ml_pipe": ", ".join(f"{r['StockTypeID']}:{r['raw_model']}|{str(r['raw_color']).upper()}" for r in linked),
                   "ml_cats": ", ".join(sorted({str(r["Category"]) for r in linked})),
                   "summed_qty": total, "n_links": len(linked)}
            locked.append(rec)   # keep even if total == 0 (Target Qty 0); no separate Skipped bucket
        elif (brand, mk(model), cap) in ml_by_mc:    # model+cap match, color differs -> REVIEW
            cands = ml_by_mc[(brand, mk(model), cap)]
            review.append({**base, "reason": "color mismatch",
                           "status": "Model + storage match, colour differs", "score": 80,
                           "ml_candidate_colors": ", ".join(sorted({r["raw_color"] for r in cands})),
                           "ml_pipe": ", ".join(f"{r['StockTypeID']}:{r['raw_model']}|{str(r['raw_color']).upper()}" for r in cands),
                           "ml_cats": ", ".join(sorted({str(r["Category"]) for r in cands})),
                           "candidate_StockTypeIDs": ", ".join(str(r["StockTypeID"]) for r in cands)})
        elif is_acc:                             # accessory, unmatched -> ACCESSORIES
            accessories.append({**base, "note": "accessory (not matched) -> set 10"})
        elif (brand, mk(model)) in ml_by_m:          # model match, capacity differs -> REVIEW
            cands = ml_by_m[(brand, mk(model))]
            review.append({**base, "reason": "capacity mismatch",
                           "status": "Model match, storage differs", "score": 70,
                           "ml_candidate_caps": ", ".join(sorted({r["capacity"] for r in cands})),
                           "ml_pipe": ", ".join(f"{r['StockTypeID']}:{r['raw_model']}|{str(r['raw_color']).upper()}" for r in cands),
                           "ml_cats": ", ".join(sorted({str(r["Category"]) for r in cands})),
                           "candidate_StockTypeIDs": ", ".join(str(r["StockTypeID"]) for r in cands)})
        else:                                    # nothing close -> REVIEW (no match)
            # fuzzy: any ML row same brand whose model contains/!contained
            fuzzy = [r for r in ml_by_m_all(ml, brand) if r["model"] and (r["model"] in model or model in r["model"])] if model else []
            review.append({**base, "reason": "no match" if not fuzzy else "fuzzy model only",
                           "status": "No Masterlist match" if not fuzzy else "Fuzzy model match",
                           "score": ("" if not fuzzy else 50),
                           "ml_candidate_caps": "", "ml_candidate_colors": "",
                           "ml_pipe": ", ".join(f"{r['StockTypeID']}:{r['raw_model']}|{str(r['raw_color']).upper()}" for r in fuzzy[:8]),
                           "ml_cats": ", ".join(sorted({str(r["Category"]) for r in fuzzy})) if fuzzy else "",
                           "candidate_StockTypeIDs": ", ".join(str(r["StockTypeID"]) for r in fuzzy[:8])})

    # Not on GetGreenr yet: ML keys with qty>0 never matched
    not_yet = []
    seen = set()
    for kk, rows in ml_by_key.items():
        if kk in matched_ml_keys:
            continue
        tot = sum(r["qty"] for r in rows)
        if tot <= 0:
            continue
        r0 = rows[0]
        if kk in seen:
            continue
        seen.add(kk)
        not_yet.append({"brand": r0["brand"], "model": r0["model"], "capacity": r0["capacity"],
                        "color": r0["color"], "raw_model": r0["raw_model"], "raw_color": r0["raw_color"],
                        "qty": tot, "StockTypeIDs": ", ".join(str(r["StockTypeID"]) for r in rows)})

    return dict(locked=pd.DataFrame(locked), review=pd.DataFrame(review),
                accessories=pd.DataFrame(accessories),
                not_yet=pd.DataFrame(not_yet), gg=gg, ml=ml)


def ml_by_m_all(ml, brand):
    return [r for _, r in ml.iterrows() if r["brand"] == brand]


if __name__ == "__main__":
    out = build()
    print("=== TALLY ===")
    for name in ["locked", "review", "accessories", "not_yet"]:
        print(f"  {name:12s}: {len(out[name])}")
    tot = len(out["locked"]) + len(out["review"]) + len(out["accessories"])
    print(f"  GG listings classified: {tot} / {len(out['gg'])}")

    print("\n=== MATCH REVIEW breakdown by reason ===")
    if not out["review"].empty:
        print(out["review"]["reason"].value_counts().to_string())

    print("\n=== MATCH REVIEW sample (first 25) ===")
    if not out["review"].empty:
        cols = [c for c in ["gg_sku", "reason", "ml_candidate_colors", "ml_candidate_caps"] if c in out["review"].columns]
        print(out["review"][cols].head(25).to_string(index=False, max_colwidth=48))

    print("\n=== LOCKED sample ===")
    if not out["locked"].empty:
        print(out["locked"][["gg_sku", "summed_qty", "n_links"]].head(12).to_string(index=False, max_colwidth=55))

    # Build Match Review with a clear suggested-action column, sorted for eyeballing
    rev = out["review"].copy()
    if not rev.empty:
        action = {"color mismatch": "Confirm colour / set 0 if not stocked",
                  "capacity mismatch": "Confirm capacity / set 0 if not stocked",
                  "no match": "Map manually or leave (likely not in POS)",
                  "fuzzy model only": "Verify model mapping"}
        rev["suggested_action"] = rev["reason"].map(action)
        front = ["gg_sku", "gg_title", "reason", "suggested_action", "brand",
                 "model", "capacity", "color"]
        rev["brand"] = rev["gg_sku"].str.split().str[0]
        rev = rev[[c for c in front if c in rev.columns] +
                  [c for c in rev.columns if c not in front]]
        rev = rev.sort_values(["brand", "reason", "model"])

    def _df(x, cols):
        return x if (x is not None and not x.empty) else pd.DataFrame(columns=cols)

    summary = pd.DataFrame({
        "Tab": ["Locked Matches", "Match Review",
                "Accessories (Not Matched)", "Not on GetGreenr Yet"],
        "Count": [len(out["locked"]), len(rev),
                  len(out["accessories"]), len(out["not_yet"])],
        "Seller Stock action": ["SUM of linked Masterlist (Used) Available Qty", "leave unchanged",
                                "set to 10", "n/a (not a GetGreenr listing)"],
    })

    reg = "/sessions/dreamy-hopeful-mayer/mnt/outputs/GetGreenr_SKU_Registry.xlsx"
    with pd.ExcelWriter(reg, engine="openpyxl") as xl:
        summary.to_excel(xl, sheet_name="Summary", index=False)
        _df(out["locked"], ["gg_sku"]).to_excel(xl, sheet_name="Locked Matches", index=False)
        _df(rev, ["gg_sku"]).to_excel(xl, sheet_name="Match Review", index=False)
        _df(out["accessories"], ["gg_sku"]).to_excel(xl, sheet_name="Accessories (Not Matched)", index=False)
        _df(None, ["gg_sku", "note"]).to_excel(xl, sheet_name="Not Selling in GetGreenr", index=False)
        _df(out["not_yet"], ["brand"]).to_excel(xl, sheet_name="Not on GetGreenr Yet", index=False)
    print(f"\nWrote {reg}")
