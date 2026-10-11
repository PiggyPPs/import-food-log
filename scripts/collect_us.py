"""
輸入食品・食器の違反ログ — 輸出側（アメリカで輸入を拒否された日本の食品）

FDA Data Dashboard API の import_refusals から、生産国が日本の「食品」「食品関連（食器など）」を取り出して
data/us_refusals.json にまとめます。会社名・住所は保存しません。

必要な環境変数（GitHub の Secrets から渡す）:
  FDA_USER  … Authorization-User（登録したメールアドレス）
  FDA_KEY   … Authorization-Key
"""
import csv
import io
import json
import os
import re
import sys
import urllib.request
from datetime import datetime, timedelta, timezone

API_URL = "https://api-datadashboard.fda.gov/v1/import_refusals"
CHARGES_URL = "https://datadashboard.fda.gov/oii/download/ACT_SECTION_CHARGES.CSV"
DATA_PATH = "data/us_refusals.json"
META_PATH = "data/us_meta.json"
CHARGES_PATH = "data/us_charges.json"
START_DATE = "2026-04-01"   # 輸入側とそろえて令和8年度から
JST = timezone(timedelta(hours=9))
UA = "import-food-log/1.0 (+https://piggypps.github.io/import-food-log/)"

COLUMNS = [
    "ShipmentID", "RefusalDate", "CountryName", "ProductCategory",
    "IndustryCodeDescription", "ProductCodeDescription", "RefusalCharges",
    "FDASampleAnalysis", "PrivateLabAnalysis",
]


def post(body):
    user, key = os.environ.get("FDA_USER", ""), os.environ.get("FDA_KEY", "")
    if not user or not key:
        raise SystemExit("FDA_USER / FDA_KEY が設定されていません")
    req = urllib.request.Request(
        API_URL,
        data=json.dumps(body).encode("utf-8"),
        headers={
            "Content-Type": "application/json",
            "Authorization-User": user,
            "Authorization-Key": key,
            "User-Agent": UA,
        },
        method="POST",
    )
    with urllib.request.urlopen(req, timeout=90) as res:
        return json.loads(res.read().decode("utf-8"))


def fetch_refusals():
    rows, start = [], 1
    while True:
        body = {
            "start": start,
            "rows": 5000,
            "sort": "RefusalDate",
            "sortorder": "desc",
            "filters": {"CountryName": ["Japan"], "RefusalDateFrom": [START_DATE]},
            "columns": COLUMNS,
        }
        res = post(body)
        result = res.get("result") or []
        print(f"start={start} status={res.get('statuscode')} message={res.get('message')} count={len(result)}")
        if not isinstance(result, list):
            print(json.dumps(res, ensure_ascii=False)[:800], file=sys.stderr)
            raise SystemExit("想定外の応答です")
        if res.get("invalid_columns") or res.get("invalid_filters") or res.get("missing_parameters"):
            print(json.dumps(res, ensure_ascii=False)[:800], file=sys.stderr)
            raise SystemExit("リクエストの項目にエラーがあります")
        rows.extend(result)
        if len(result) < 5000:
            break
        start += 5000
    return rows


def fetch_charges():
    """違反コード → 説明文（英語）の対応表。取れなければ前回の分を使う。"""
    try:
        req = urllib.request.Request(CHARGES_URL, headers={"User-Agent": UA})
        with urllib.request.urlopen(req, timeout=60) as res:
            raw = res.read()
        text = raw.decode("utf-8-sig", errors="replace")
        reader = list(csv.reader(io.StringIO(text)))
        head, body = reader[0], [r for r in reader[1:] if r]
        print("charges header:", head)
        # コードの列（短い値）と説明の列（長い値）を見つける
        avg = [sum(len(r[i]) for r in body if i < len(r)) / max(1, len(body)) for i in range(len(head))]
        code_i = next((i for i, h in enumerate(head) if re.search(r"charge.*code|^code$|charge$", h, re.I)), None)
        if code_i is None:
            code_i = min(range(len(head)), key=lambda i: avg[i] if avg[i] > 2 else 999)
        text_i = max(range(len(head)), key=lambda i: avg[i])
        table = {}
        for r in body:
            if len(r) > max(code_i, text_i):
                table[r[code_i].strip().upper()] = r[text_i].strip()
        print(f"charges: {len(table)} codes (code col={head[code_i]}, text col={head[text_i]})")
        return table
    except Exception as e:
        print(f"charges fetch failed: {e}", file=sys.stderr)
        if os.path.exists(CHARGES_PATH):
            with open(CHARGES_PATH, encoding="utf-8") as f:
                return json.load(f)
        return {}


# 違反の種類（英語のコードと説明文から判定）
RULES = [
    ("アレルギー表示", r"ALLERG|ALERGEN|MAJOR FOOD"),
    ("菌・微生物", r"SALMONEL|LISTERIA|E\.? ?COLI|BACTER|MICROB|PATHOGEN|BOTUL|CLOSTRID|STAPH"),
    ("有害物質", r"POISON|DELETERIOUS|HEAVY METAL|LEAD|MERCURY|ARSENIC|CADMIUM|MELAMINE|HISTAMINE|AFLATOX|TOXIN"),
    ("残留農薬・動物用医薬品", r"PESTIC|VET ?DRUG|DRUG RES|ANIMAL DRUG"),
    ("添加物・着色料", r"COLOR ADDITIVE|COLOUR|FOOD ADDITIVE|UNSAFE ADD|SULFIT|SULPHIT|UNSAFE COL"),
    ("食品として不適", r"UNFIT FOR FOOD"),
    ("汚れ・腐敗", r"FILTH|DECOMP|PUTRID|ROTTEN|SPOIL"),
    ("製造・衛生管理の基準", r"123\.\d|HACCP|GOOD MANUFACTURING|INSANITARY"),
    ("缶詰・びん詰の届け出", r"SCHEDULED PROCESS|LOW ACID|ACIDIFIED|LACF|108\.(25|35)|113\.|114\."),
    ("登録・書類", r"NOT LISTED|REGIST|PRIOR NOTICE|FSVP|REFUSED INSPECT|FOREIGN SUPPLIER"),
    ("医薬品・健康の表示", r"NEW DRUG|UNAPPROVED|DISEASE|CURE|MITIGATE"),
    ("ラベル・表示", r"MISBRAND|LABEL|ENGLISH|USUAL NAME|INGREDIENT|NUTRITION|FALSE|NET QUANTITY|DECLAR"),
]
ORDER = [n for n, _ in RULES] + ["その他"]
SAFETY = {"アレルギー表示", "菌・微生物", "有害物質", "残留農薬・動物用医薬品", "添加物・着色料", "食品として不適", "汚れ・腐敗"}

INDUSTRY_JA = [
    (r"FISH|SEAFOOD|SHELLFISH", "魚介類"),
    (r"CANDY|CHOCOLATE|CONFECTION|GUM", "菓子"),
    (r"BAKERY|BREAD|COOKIE|CRACKER|BISCUIT|DOUGH", "パン・焼き菓子"),
    (r"NOODLE|MACARONI|PASTA", "めん類"),
    (r"CEREAL|RICE|GRAIN|FLOUR", "穀物・米"),
    (r"SNACK", "スナック"),
    (r"SOFT DRINK|BEVERAGE|WATER|JUICE", "飲料"),
    (r"COFFEE|TEA", "お茶・コーヒー"),
    (r"SPICE|FLAVOR|SALT|SEASON", "調味料・香辛料"),
    (r"DRESSING|CONDIMENT|SAUCE|SOY", "調味料・ソース"),
    (r"SOUP", "スープ"),
    (r"VEGETABLE|VEG", "野菜加工品"),
    (r"FRUIT", "果物加工品"),
    (r"NUT|SEED", "ナッツ・種子"),
    (r"MILK|CHEESE|DAIRY|BUTTER|ICE CREAM", "乳製品"),
    (r"EGG", "卵"),
    (r"MEAT|POULTRY", "肉"),
    (r"VITAMIN|DIETARY|SUPPLEMENT|MINERAL|PROTEIN", "サプリメント"),
    (r"ADDITIVE|ENZYME", "食品添加物"),
    (r"MULTIPLE FOOD|DINNER|SANDWICH|PREPARED", "総菜・加工食品"),
    (r"FOOD SERVICE|HOUSEWARE|UTENSIL|CONTAINER|PACKAG|CERAMIC|GLASS", "食器・容器"),
]


def industry_ja(desc, category):
    d = (desc or "").upper()
    if "HOUSEWARE" in (category or "").upper():
        return "食器・容器"
    for pat, ja in INDUSTRY_JA:
        if re.search(pat, d):
            return ja
    return "その他の食品"


def split_codes(s):
    return [c.strip().upper() for c in re.split(r"[,;|]", s or "") if c.strip()]


def categorize(codes, charges):
    cats = []
    for c in codes:
        text = (c + " " + charges.get(c, "")).upper()
        for name, pat in RULES:
            if re.search(pat, text):
                if name not in cats:
                    cats.append(name)
                break
        else:
            if "その他" not in cats:
                cats.append("その他")
    return sorted(cats or ["その他"], key=ORDER.index)


def main():
    charges = fetch_charges()
    raw = fetch_refusals()
    print(f"raw rows: {len(raw)}")
    if raw:
        print("sample keys:", sorted(raw[0].keys()))
        print("sample:", {k: raw[0].get(k) for k in ("RefusalDate", "ProductCategory", "IndustryCodeDescription", "ProductCodeDescription", "RefusalCharges")})

    data = {}
    for r in raw:
        cat = r.get("ProductCategory") or ""
        if not re.search(r"Human Food|Housewares", cat, re.I):
            continue
        codes = split_codes(r.get("RefusalCharges"))
        kinds = categorize(codes, charges)
        date = (r.get("RefusalDate") or "")[:10]
        rid = r.get("ShipmentID") or f"{date}-{r.get('ProductCodeDescription')}-{','.join(codes)}"
        data[rid] = {
            "id": rid,
            "date": date,
            "month": date[:7],
            "product": r.get("ProductCodeDescription") or "",
            "industry": r.get("IndustryCodeDescription") or "",
            "industry_ja": industry_ja(r.get("IndustryCodeDescription"), cat),
            "codes": codes,
            "kinds": kinds,
            "safety": any(k in SAFETY for k in kinds),
            "sampled": (r.get("FDASampleAnalysis") or "").upper().startswith("Y")
                       or (r.get("PrivateLabAnalysis") or "").upper().startswith("Y"),
        }
    out = sorted(data.values(), key=lambda x: (x["date"], x["product"]), reverse=True)

    os.makedirs("data", exist_ok=True)
    used = sorted({c for r in out for c in r["codes"]})
    with open(CHARGES_PATH, "w", encoding="utf-8") as f:
        json.dump({c: charges.get(c, "") for c in used} if charges else {}, f, ensure_ascii=False, indent=1)
    with open(DATA_PATH, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=1)
    with open(META_PATH, "w", encoding="utf-8") as f:
        json.dump({"checked_at": datetime.now(JST).strftime("%Y-%m-%d %H:%M"), "count": len(out),
                   "source": "FDA Data Dashboard / Import Refusals"}, f, ensure_ascii=False, indent=1)
    print(f"saved {len(out)} records")
    from collections import Counter
    print("kinds:", Counter(k for r in out for k in r["kinds"]).most_common())
    print("unmatched codes:", sorted({c for r in out for c in r["codes"] if categorize([c], charges) == ["その他"]}))


if __name__ == "__main__":
    main()
