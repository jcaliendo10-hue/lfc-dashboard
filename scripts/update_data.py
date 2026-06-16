#!/usr/bin/env python3
"""
Liverpool FC dashboard — data updater.

Scrapes Liverpool news from multiple RSS feeds, scores each story by
(1) source credibility and (2) relevance, ranks them, and writes:
  - data/news.json          (consumed by the dashboard)
  - data/liverpool_data.xlsx (the Excel "backend" / source of truth)

Run locally:  python scripts/update_data.py
Runs daily on GitHub Actions (see .github/workflows/update.yml).

No API keys required — everything is public RSS.
"""

import json
import re
import sys
import datetime as dt
from pathlib import Path
from html import unescape

try:
    import feedparser
except ImportError:
    sys.exit("Missing dependency: pip install feedparser openpyxl")

ROOT = Path(__file__).resolve().parents[1]
DATA = ROOT / "data"
DATA.mkdir(exist_ok=True)

# ---------------------------------------------------------------------------
# SOURCE CREDIBILITY MODEL
# credibility: 1-10 accuracy/reliability weight for the outlet.
# lfc_specific: True if the whole feed is already Liverpool-only (so every
#               item is relevant); False feeds are filtered for "liverpool".
# ---------------------------------------------------------------------------
SOURCES = [
    # Tier 1 — most reliable
    {"name": "BBC Sport",        "credibility": 10, "lfc_specific": True,
     "url": "https://feeds.bbci.co.uk/sport/football/teams/liverpool/rss.xml"},
    {"name": "The Guardian",     "credibility": 10, "lfc_specific": True,
     "url": "https://www.theguardian.com/football/liverpool/rss"},
    {"name": "Liverpool FC (official)", "credibility": 10, "lfc_specific": True,
     "url": "https://www.liverpoolfc.com/news.rss"},
    # Tier 2 — strong national outlets (general feeds, filtered for Liverpool)
    {"name": "Sky Sports",       "credibility": 9, "lfc_specific": False,
     "url": "https://www.skysports.com/rss/12040"},
    {"name": "ESPN FC",          "credibility": 8, "lfc_specific": False,
     "url": "https://www.espn.com/espn/rss/soccer/news"},
    {"name": "Liverpool Echo",   "credibility": 7, "lfc_specific": True,
     "url": "https://www.liverpoolecho.co.uk/all-about/liverpool-fc?service=rss"},
    # Tier 3 — reputable fan / aggregator sites
    {"name": "This Is Anfield",  "credibility": 6, "lfc_specific": True,
     "url": "https://www.thisisanfield.com/feed/"},
    {"name": "Empire of the Kop","credibility": 5, "lfc_specific": True,
     "url": "https://www.empireofthekop.com/feed/"},
]

# Relevance keywords (lowercase). Hits add to the relevance score.
KEYWORDS = [
    "liverpool", "anfield", "iraola", "transfer", "signing", "signs", "deal",
    "contract", "injury", "fixture", "wirtz", "isak", "salah", "van dijk",
    "szoboszlai", "ekitike", "gakpo", "mac allister", "gravenberch", "frimpong",
    "konate", "robertson", "fsg", "premier league", "champions league",
]

MAX_ITEMS = 25          # how many ranked stories to keep
RECENCY_DAYS = 21       # ignore anything older than this


def clean(text):
    text = unescape(text or "")
    text = re.sub(r"<[^>]+>", "", text)        # strip HTML tags
    return re.sub(r"\s+", " ", text).strip()


def published_dt(entry):
    for key in ("published_parsed", "updated_parsed"):
        t = entry.get(key)
        if t:
            return dt.datetime(*t[:6])
    return dt.datetime.utcnow()


def score_item(title, summary, src, age_days):
    """Return (relevance 0-100, credibility 1-10, composite)."""
    hay = (title + " " + summary).lower()

    relevance = 0
    if src["lfc_specific"] or "liverpool" in hay:
        relevance += 40
    if "liverpool" in summary.lower():
        relevance += 10
    hits = sum(1 for k in KEYWORDS if k in hay)
    relevance += min(hits * 5, 30)
    # recency boost
    if age_days <= 1:
        relevance += 20
    elif age_days <= 3:
        relevance += 12
    elif age_days <= 7:
        relevance += 6
    relevance = min(relevance, 100)

    cred = src["credibility"]
    # Blend: relevance weighted by source quality so a strong source
    # outranks a weak one at equal relevance.
    composite = round(relevance * (0.5 + 0.5 * cred / 10), 1)
    return relevance, cred, composite


def fetch_news():
    items = []
    now = dt.datetime.utcnow()
    for src in SOURCES:
        try:
            feed = feedparser.parse(src["url"])
        except Exception as e:
            print(f"  ! {src['name']}: {e}")
            continue
        n = 0
        for e in feed.entries:
            title = clean(e.get("title", ""))
            summary = clean(e.get("summary", ""))
            if not title:
                continue
            hay = (title + " " + summary).lower()
            if not src["lfc_specific"] and "liverpool" not in hay:
                continue
            pub = published_dt(e)
            age = (now - pub).days
            if age > RECENCY_DAYS:
                continue
            rel, cred, comp = score_item(title, summary, src, age)
            items.append({
                "headline": title,
                "summary": summary[:240],
                "source": src["name"],
                "credibility": cred,
                "relevance": rel,
                "rank_score": comp,
                "date": pub.strftime("%Y-%m-%d"),
                "url": e.get("link", ""),
            })
            n += 1
        print(f"  {src['name']}: {n} items")

    # de-duplicate by normalised headline, keep the highest-scoring copy
    items.sort(key=lambda x: x["rank_score"], reverse=True)
    seen, deduped = set(), []
    for it in items:
        key = re.sub(r"[^a-z0-9]", "", it["headline"].lower())[:60]
        if key in seen:
            continue
        seen.add(key)
        deduped.append(it)

    for i, it in enumerate(deduped, 1):
        it["rank"] = i
    return deduped[:MAX_ITEMS]


def load_reference():
    """Static reference data (squad, transfers, etc.). Edit via Excel or here."""
    ref_path = DATA / "reference.json"
    if ref_path.exists():
        return json.loads(ref_path.read_text(encoding="utf-8"))
    return {}


def write_excel(news, ref):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment

    wb = Workbook()
    red = PatternFill("solid", fgColor="C8102E")
    head = Font(bold=True, color="FFFFFF")

    def sheet(title, headers, rows):
        ws = wb.create_sheet(title)
        ws.append(headers)
        for c in ws[1]:
            c.fill = red; c.font = head; c.alignment = Alignment(vertical="center")
        for r in rows:
            ws.append(r)
        for i, h in enumerate(headers, 1):
            width = max([len(str(h))] + [len(str(r[i-1])) for r in rows]) if rows else len(h)
            ws.column_dimensions[chr(64+i)].width = min(width + 3, 60)
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        return ws

    wb.remove(wb.active)

    sheet("News (ranked)",
          ["Rank", "Date", "Headline", "Source", "Credibility (1-10)",
           "Relevance (0-100)", "Rank score", "URL"],
          [[n["rank"], n["date"], n["headline"], n["source"], n["credibility"],
            n["relevance"], n["rank_score"], n["url"]] for n in news])

    sheet("Sources",
          ["Source", "Credibility (1-10)", "Liverpool-only feed", "Feed URL"],
          [[s["name"], s["credibility"], "Yes" if s["lfc_specific"] else "No", s["url"]]
           for s in sorted(SOURCES, key=lambda s: -s["credibility"])])

    if ref.get("squad"):
        sheet("Squad",
              ["#", "Player", "Position", "Role", "Age", "Nationality", "Contract"],
              [[p["num"], p["name"], p["posGroup"], p["role"], p["age"], p["nat"], p["contract"]]
               for p in ref["squad"]])
    if ref.get("loans"):
        sheet("Loanees", ["Player", "Position", "Loan club", "Notes"],
              [[l["name"], l["posGroup"], l["club"], l["notes"]] for l in ref["loans"]])
    if ref.get("transfers"):
        t = ref["transfers"]
        rows = []
        for c in t.get("in", []):    rows.append(["IN (confirmed)", c["who"], c["meta"], c["fee"], ""])
        for c in t.get("out", []):   rows.append(["OUT (confirmed)", c["who"], c["meta"], c["fee"], ""])
        for c in t.get("rumIn", []): rows.append(["Rumour IN", c["who"], c["meta"], c["fee"], c["p"]])
        for c in t.get("rumOut", []):rows.append(["Rumour OUT", c["who"], c["meta"], c["fee"], c["p"]])
        sheet("Transfers", ["Type", "Player", "Detail", "Fee", "Likelihood"], rows)

    out = DATA / "liverpool_data.xlsx"
    wb.save(out)
    print(f"  wrote {out}")


def main():
    print("Fetching Liverpool news feeds...")
    news = fetch_news()
    print(f"Ranked {len(news)} stories.")

    ref = load_reference()
    payload = {
        "updated": dt.datetime.utcnow().strftime("%d %b %Y %H:%M UTC"),
        "count": len(news),
        "items": news,
    }
    (DATA / "news.json").write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"  wrote {DATA/'news.json'}")

    try:
        write_excel(news, ref)
    except Exception as e:
        print(f"  ! Excel write failed: {e}")

    print("Done.")


if __name__ == "__main__":
    main()
