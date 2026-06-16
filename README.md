# Liverpool FC Dashboard

An automated, self-updating Liverpool FC dashboard hosted free on GitHub Pages.

- **Live news** scraped daily from multiple RSS feeds, **ranked by source credibility × relevance**.
- **Excel backend** (`data/liverpool_data.xlsx`) regenerated on every run — your editable source of truth.
- **Squad, loanees, transfers & rumours, analytics, club hierarchy** rendered with charts and an org diagram.
- Refreshes automatically via **GitHub Actions** — no servers, no API keys.

## How it works

```
scripts/update_data.py     scrapes RSS → scores & ranks → writes data/news.json + data/liverpool_data.xlsx
.github/workflows/update.yml   runs the script daily (07:00 UTC) and commits refreshed data
index.html                 reads data/reference.json + data/news.json and renders the dashboard
data/reference.json        squad / transfers / analytics (you edit these; mirrored into the Excel file)
```

### The ranking model

Each story gets two scores:

- **Credibility (1–10)** — how reliable the outlet is. Tier 1 (BBC, Guardian, official site) = 10; strong nationals (Sky, ESPN) = 8–9; reputable fan/aggregator sites (This Is Anfield, Empire of the Kop) = 5–6.
- **Relevance (0–100)** — Liverpool focus + keyword matches (players, transfers, Iraola, fixtures…) + a recency boost.

Final rank score `= relevance × (0.5 + 0.5 × credibility/10)`, so a strong source outranks a weak one at equal relevance. Edit the `SOURCES` list and `KEYWORDS` in `scripts/update_data.py` to tune it.

## One-time setup (account: jcaliendo10-hue)

1. **Create a repo** on GitHub named e.g. `lfc-dashboard` (public).
2. **Upload these files** (keep the folder structure). Either drag-and-drop in the GitHub web UI, or:
   ```bash
   git init
   git add .
   git commit -m "Initial Liverpool dashboard"
   git branch -M main
   git remote add origin https://github.com/jcaliendo10-hue/lfc-dashboard.git
   git push -u origin main
   ```
3. **Enable GitHub Pages**: repo → *Settings → Pages* → Source = *Deploy from a branch*, Branch = `main`, folder = `/ (root)`. Save.
   Your dashboard will be live at `https://jcaliendo10-hue.github.io/lfc-dashboard/`.
4. **Allow Actions to commit**: repo → *Settings → Actions → General → Workflow permissions* → select **Read and write permissions**. Save.
5. **Run it once now**: repo → *Actions* tab → "Update Liverpool dashboard data" → *Run workflow*. It will scrape live feeds and replace the seed data.

After that it runs by itself every day.

## Run locally

```bash
pip install -r requirements.txt
python scripts/update_data.py
# then open index.html via a local server (fetch() needs http, not file://):
python -m http.server 8000   # visit http://localhost:8000
```

## Editing the squad / transfer data

Open `data/liverpool_data.xlsx` (or edit `data/reference.json` directly) to change squad, loanees, transfers and analytics. The news section updates itself.

---

*Data current to 16 June 2026. Transfer rumours and likelihood ratings are interpretive and change frequently. Unofficial fan project — not affiliated with Liverpool FC.*
