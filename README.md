# Dr. Jitendra Kumar — Portfolio Website

Plain HTML/CSS/JS, no build step, no framework. White background, maroon (`#8a1f2b`) accent, Inter throughout. Content matches your original React site (real bio, real publications, real students, real timeline, real news/awards/teaching), rebuilt as a static site where **almost everything is Excel-driven** — no code editing needed for day-to-day updates.

## Pages

```
index.html                → Home: intro (name/role/photo), bio, News feed,
                             Experience timeline preview, citations/publications chart
experience.html            → Full career timeline (data/experience.xlsx)
teaching.html               → Courses taught, current + previous (data/teaching.xlsx)
publications.html           → Full searchable/filterable publication list (data/publications.xlsx)
students-current.html       → Currently supervised students (data/students.xlsx)
students-graduated.html     → Graduated students + past research interns (data/students.xlsx)
awards.html                  → Awards & recognitions (data/awards.xlsx)
news.html                    → Full news/updates list, grouped by year (data/news.xlsx)
resources.html                → Public tools/datasets/tutorials (hardcoded — rarely changes)
protected.html                 → "Protected" resources page with a simple password gate (NOT real security)
contact.html                    → Contact details

style.css                  → All styling (edit CSS variables at the top to re-theme)
script.js                   → Nav toggle + all Excel loading/parsing/rendering + chart + password gate
vendor/xlsx.full.min.js      → SheetJS library (reads Excel in-browser), bundled locally
data/*.xlsx                   → ★ Your content — see below
assets/                         → Photos, social icons, CV
```

## What's Excel-driven

| File | Powers | Key columns |
|---|---|---|
| `data/publications.xlsx` | Publications page + home preview | 5 sheets (Journals/Conferences/Books/Book Chapters/Reports), each with Authors, Title, type-specific venue fields, Year, Status, Indexing, Impact Factor, Quartile, Other Info, Link |
| `data/students.xlsx` | Both student pages | Category, Name, Program, Year, Research Area/Thesis Title, Status, Affiliation, Notes, Image, Thesis Link |
| `data/experience.xlsx` | Experience page + home timeline preview | Year, Title, Institution, Note |
| `data/awards.xlsx` | Awards page | Year, Title, Description, Medal |
| `data/teaching.xlsx` | Teaching page | Institution, Course, Status (Current/Previous) |
| `data/news.xlsx` | News page + home preview | Year, Type, Headline, Detail, Link |
| `data/metrics.xlsx` | Home page chart | Year, Citations, Publications |

Every one of these files has an **Instructions** tab explaining its columns. General pattern: open the file, add/edit/delete a row, save (keep the filename), re-upload to the `data/` folder in your repo replacing the old file, refresh the site.

### Publications — Status and visibility

`Status` controls what's public but is **never shown** on the site itself:
- `Published`, `Accepted`, `Preprint` → visible to visitors
- `In Review`, `Submitted`, `Revision Submitted` → hidden (track work-in-progress privately in the same file)

`Indexing` uses `/` as a separator (e.g. `SCI/WoS/Scopus`). `Quartile` (Q1–Q4) renders as its own colored tag. `Link`, if filled in, makes the title clickable (points to the DOI/publisher page).

### Students — moving someone from Current to Graduated

Just change their `Status` from `Ongoing` to `Completed` (or `Awarded` for a PhD). They'll automatically appear on the Graduated page instead, no other edit needed. `Affiliation` doubles as "current employer" for graduates and "home institute" for research interns — whichever makes sense for that row.

### Photos

Student `Image` column should contain just a filename (e.g. `RohitPhD.jpg`) — the actual file needs to be uploaded to `assets/students/`. Leave `Image` blank for a clean initials-avatar fallback (this is the default for most students right now — only a handful of real photos exist in the original source).

### The citations chart

Google Scholar doesn't allow a static site to pull your citation count live (no CORS support, and it blocks scripted access) — so `data/metrics.xlsx` is a small manually-updated table. Check your Scholar profile periodically and update the `Citations` column; `Publications` (papers per year) you could also just eyeball from your Publications page.

## What's NOT Excel-driven (edit the HTML directly)

- **Bio text** on the home page (and the "show more" expanded paragraph) — these are short enough that editing `index.html` directly is simpler than adding a spreadsheet for two paragraphs.
- **Resources page** (Tools/Datasets/Tutorials) — static reference links that rarely change.
- **Contact details, social links** — in `contact.html` and the social row on `index.html`.

## The "Protected Resources" page

`protected.html` has a password gate (current passwords: `123456`, `234567`, `345678`, matching your original site). **This is not real security** — the password check runs entirely in the visitor's browser, so anyone who views the page source can read the passwords or the gated content directly. It's a basic deterrent only, exactly like the original. A real backend would be needed for genuine access control, which isn't possible on free static hosting like GitHub Pages. To change the passwords, edit the `RESOURCE_PASSWORDS` array near the bottom of `script.js`.

## Important: Excel-driven pages only work over http(s)

Browsers block a local page from reading a local file via `fetch()` (CORS security). So:
- ✅ Works once deployed on GitHub Pages / Netlify / Vercel
- ✅ Works via a local test server (`python3 -m http.server` from this folder)
- ❌ Does **not** work by double-clicking `index.html` directly — you'll see a "Couldn't load" message. Expected, not a bug.

## Deploy on GitHub Pages (free)

1. Create a repo (e.g. `yourusername.github.io` for a root URL, or any name for a project site).
2. Upload the **contents** of this folder to the repo root — `index.html` must sit directly at the root, not inside a subfolder. Make sure to include the `data/`, `assets/`, and `vendor/` folders (drag the folders themselves so structure is preserved).
3. Repo **Settings → Pages** → Source: **Deploy from a branch** → Branch **main**, folder **/ (root)** → Save.
4. Live within a minute or two at the URL GitHub shows you.

### Using git
```bash
git init
git add .
git commit -m "Initial site"
git branch -M main
git remote add origin https://github.com/yourusername/yourusername.github.io.git
git push -u origin main
```

## Alternatives (also free)
Netlify (drag-and-drop at netlify.com/drop), Vercel, Cloudflare Pages — same static files, same result.
