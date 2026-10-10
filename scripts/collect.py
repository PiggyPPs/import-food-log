"""
輸入食品・食器の違反ログ — データ収集スクリプト

厚生労働省「輸入食品等の食品衛生法違反事例」のページから Excel ファイルを取得し、
data/violations.json にまとめます。会社名（製造者・シッパー・輸入者）は保存しません。
"""
import hashlib
import io
import json
import os
import re
import sys
import time
import urllib.request
from datetime import date, datetime, timedelta, timezone
from html import unescape
from urllib.parse import urljoin

import openpyxl

INDEX_URL = "https://www.mhlw.go.jp/stf/seisakunitsuite/bunya/kenkou_iryou/shokuhin/yunyu_kanshi/ihan/index.html"
DATA_PATH = "data/violations.json"
META_PATH = "data/meta.json"
START_MONTH = "2025-04"   # これより前の事例は保存しない（令和7年度から）
MAX_FILES = 80
UA = "Mozilla/5.0 (compatible; import-food-log/1.0; +https://piggypps.github.io/import-food-log/)"
JST = timezone(timedelta(hours=9))

ZEN = str.maketrans("０１２３４５６７８９", "0123456789")


def fetch(url):
    req = urllib.request.Request(url, headers={"User-Agent": UA})
    with urllib.request.urlopen(req, timeout=60) as res:
        return res.read()


def clean(v):
    if v is None:
        return ""
    s = str(v).replace("　", " ").replace("\r", "")
    s = re.sub(r"[ \t]+", " ", s)
    s = re.sub(r"\s*\n\s*", " ", s)
    return s.strip()


def norm_header(v):
    return re.sub(r"\s", "", clean(v))


def to_date(v):
    if v is None or v == "":
        return None
    if isinstance(v, datetime):
        return v.date().isoformat()
    if isinstance(v, date):
        return v.isoformat()
    if isinstance(v, (int, float)) and 30000 < v < 80000:
        return (date(1899, 12, 30) + timedelta(days=int(v))).isoformat()
    m = re.search(r"(20\d\d)[/\-年.](\d{1,2})[/\-月.](\d{1,2})", clean(v).translate(ZEN))
    if m:
        return f"{int(m.group(1)):04d}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"
    return None


def reiwa_month(text):
    """「令和８年７月分」→ 2026-07"""
    t = clean(text).translate(ZEN)
    m = re.search(r"令和\s*(\d+)\s*年\s*(\d+)\s*月", t)
    if m:
        return f"{2018 + int(m.group(1)):04d}-{int(m.group(2)):02d}"
    return None


def fiscal_title_month(text):
    """「2026年度10月分」→ 2026-10（1〜3月は翌年）"""
    t = clean(text).translate(ZEN)
    m = re.search(r"(20\d\d)\s*年度\s*(\d+)\s*月", t)
    if m:
        y, mo = int(m.group(1)), int(m.group(2))
        if mo <= 3:
            y += 1
        return f"{y:04d}-{mo:02d}"
    return reiwa_month(t)


def categorize(product, reason):
    p, r = product, reason
    if re.search(r"器具|容器包装|おもちゃ|食器|合成樹脂製|陶磁器製|ガラス製|金属製|ホウロウ|ゴム製", p) \
            or re.search(r"材質別規格|材質試験|溶出試験|蒸発残留物|過マンガン酸カリウム消費量", r):
        return "器具・容器"
    if re.search(r"アフラトキシン|カビ毒|デオキシニバレノール|パツリン|オクラトキシン", r):
        return "カビ毒"
    if re.search(r"放射線照射", r):
        return "放射線照射"
    if re.search(r"遺伝子組換え|安全性未審査", r):
        return "遺伝子組換え"
    if re.search(r"残留|農薬|動物用医薬品|オキシテトラサイクリン|エンロフロキサシン|クロラムフェニコール|ニトロフラン|マラカイトグリーン", r):
        return "残留農薬・動物用医薬品"
    if re.search(r"添加物|使用基準|保存基準", r):
        return "添加物"
    if re.search(r"大腸菌|E\.?coli|細菌|サルモネラ|リステリア|腸炎ビブリオ|ノロウイルス|微生物|腸管出血", r, re.I):
        return "菌・微生物"
    if re.search(r"カビの発生|腐敗|変敗|異臭|異物|固化", r):
        return "腐敗・カビ・異物"
    return "その他"


COUNTRY_ALIASES = {
    "中華人民共和国": "中国",
    "アメリカ合衆国": "アメリカ",
    "大韓民国": "韓国",
    "英国": "イギリス",
    "グレートブリテン及び北アイルランド連合王国": "イギリス",
}


def norm_country(name):
    name = clean(name)
    return COUNTRY_ALIASES.get(name, name)


def make_id(*parts):
    return hashlib.sha1("|".join(parts).encode("utf-8")).hexdigest()[:16]


def header_map(row):
    cols = {}
    for i, v in enumerate(row):
        h = norm_header(v)
        if not h:
            continue
        if h.startswith("品名"):
            cols["product"] = i
        elif h.startswith("条文"):
            cols["law"] = i
        elif h.startswith("生産国"):
            cols["country"] = i
        elif h.startswith("不適格内容") or h.startswith("違反内容"):
            cols["reason"] = i
        elif h.startswith("措置"):
            cols["action"] = i
        elif h.startswith("備考"):
            cols["note"] = i
        elif h.startswith("公表日"):
            cols["published"] = i
    return cols if {"product", "country", "reason"} <= set(cols) else None


def cell(row, cols, key):
    i = cols.get(key)
    return row[i] if i is not None and i < len(row) else None


def parse_workbook(blob):
    wb = openpyxl.load_workbook(io.BytesIO(blob), data_only=True, read_only=True)
    records = []
    for ws in wb.worksheets:
        rows = [tuple(r) for r in ws.iter_rows(values_only=True)]
        if not rows:
            continue
        head_text = " ".join(clean(v) for r in rows[:4] for v in r if v)
        domestic = "国内における" in head_text
        title_month = fiscal_title_month(head_text)
        cols = None
        cur_month = title_month
        for r in rows:
            first = clean(r[0]) if r else ""
            if not any(v not in (None, "") for v in r):
                continue
            hm = header_map(r)
            if hm:
                cols = hm
                continue
            mm = reiwa_month(first) if re.search(r"月分", first) else None
            if mm:
                cur_month = mm
                cols = None if domestic else cols
                continue
            if not cols:
                continue
            product = clean(cell(r, cols, "product"))
            reason = clean(cell(r, cols, "reason"))
            if not product or not reason or "違反事例なし" in product:
                continue
            published = to_date(cell(r, cols, "published"))
            month = (published[:7] if published else None) or cur_month
            if not month:
                continue
            country = norm_country(cell(r, cols, "country"))
            source = "domestic" if domestic else "border"
            records.append({
                "id": make_id(source, product, country, reason, published or month),
                "source": source,
                "date": published or month,
                "month": month,
                "product": product,
                "country": country,
                "reason": reason,
                "category": categorize(product, reason),
                "inspection": clean(cell(r, cols, "note")) if not domestic else "",
                "action": clean(cell(r, cols, "action")),
                "law": clean(cell(r, cols, "law")),
            })
    return records


def find_xlsx_links(html):
    links = []
    for m in re.finditer(r'<a\b[^>]*href=["\']([^"\']+?\.xlsx)["\'][^>]*>(.*?)</a>', html, re.S | re.I):
        url = urljoin(INDEX_URL, unescape(m.group(1)))
        text = re.sub(r"<[^>]+>", "", unescape(m.group(2))).strip()
        if url not in [u for u, _ in links]:
            links.append((url, text))
    return links


def load_existing():
    if os.path.exists(DATA_PATH):
        with open(DATA_PATH, encoding="utf-8") as f:
            return json.load(f)
    return []


def main(local_files=None):
    blobs = []
    if local_files:
        for p in local_files:
            with open(p, "rb") as f:
                blobs.append((p, f.read()))
    else:
        html = fetch(INDEX_URL).decode("utf-8", errors="replace")
        links = find_xlsx_links(html)
        print(f"xlsx links: {len(links)}")
        for url, text in links[:MAX_FILES]:
            try:
                blobs.append((url, fetch(url)))
                time.sleep(1)
            except Exception as e:
                print(f"skip {url}: {e}", file=sys.stderr)

    new = []
    for name, blob in blobs:
        try:
            recs = parse_workbook(blob)
            print(f"{name}: {len(recs)} records")
            new.extend(recs)
        except Exception as e:
            print(f"parse error {name}: {e}", file=sys.stderr)

    merged = {r["id"]: r for r in load_existing()}
    for r in new:
        merged[r["id"]] = r
    data = [r for r in merged.values() if r["month"] >= START_MONTH]
    data.sort(key=lambda r: (r["date"], r["product"]), reverse=True)

    os.makedirs("data", exist_ok=True)
    with open(DATA_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=1)
    with open(META_PATH, "w", encoding="utf-8") as f:
        json.dump({
            "checked_at": datetime.now(JST).strftime("%Y-%m-%d %H:%M"),
            "count": len(data),
            "source": INDEX_URL,
        }, f, ensure_ascii=False, indent=1)
    print(f"total {len(data)} records")


if __name__ == "__main__":
    main(sys.argv[1:] or None)
