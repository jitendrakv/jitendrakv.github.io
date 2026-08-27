# AGENTS.md — Repository guide for AI agents

> Orientation file for any AI coding agent working on this repo. Read this first.
> A human-facing overview lives in `README.md`; this file is the technical map:
> what each file does, how the pieces wire together, and the rules to respect
> when editing. If you change the architecture, update this file too.

## 0. Ground rules — read before every task

This is the **personal academic website of a professor**. It represents a
scholar to peers, students, reviewers, and hiring/collaboration contacts. Every
change must fit that context. Before writing any code, internalize these:

1. **Academic, not commercial.** No marketing tone, no hype, no salesy CTAs, no
   growth-hacky patterns (popups, banners, cookie nags, newsletter modals,
   countdowns). The voice is understated and credible. Content and clarity come
   first; the design gets out of the way.
2. **Minimalism is the default.** Prefer the simplest thing that works. Add
   nothing that isn't earning its place. Generous whitespace, a clear reading
   column, strong typographic hierarchy, few colors. When unsure between "add
   something" and "leave it clean," leave it clean.
3. **Modern, but restrained.** Contemporary and polished — good type, consistent
   spacing scale, subtle depth, tasteful hover/focus states, responsive by
   default — *without* trend-chasing. No heavy gradients, glassmorphism, neon,
   loud animation, parallax, or decorative clutter. Motion is subtle and
   purposeful; respect `prefers-reduced-motion`.
4. **Restraint in the palette & type.** Stay within the existing system: white
   background, maroon `#8a1f2b` accent, the current font families. Introduce a
   new color or typeface only if asked. Change theme via the CSS variables at the
   top of `style.css`, never by scattering literal values.
5. **Accessibility & readability are non-negotiable.** Sufficient contrast,
   legible sizes, real focus outlines, semantic HTML, alt text, keyboard-usable.
   An academic audience includes people on many devices and assistive tech.
6. **Preserve the zero-build, static, Excel-driven architecture.** Don't add
   frameworks, bundlers, npm, or backends to achieve a visual goal (see §2/§5).
   Achieve "modern" with plain HTML/CSS/JS.
7. **Consistency over novelty.** Match existing components, spacing, and naming.
   A new page/section should look like it always belonged. Don't redesign one
   page in isolation.
8. **When a request is ambiguous, bias toward less.** Ask, or choose the more
   minimal interpretation, rather than adding visual weight the owner didn't ask
   for.

If a requested change would conflict with these, flag it and propose the
minimal-design alternative before proceeding.

## 1. What this is

Personal academic portfolio website for **Dr. Jitendra Kumar** (Assistant
Professor, MANIT Bhopal). Deployed as a **GitHub Pages** site at the repo root
(`jitendrakv.github.io`).

**Stack: plain static HTML + CSS + vanilla JS. No framework, no build step, no
bundler, no package.json, no Node toolchain.** What you see on disk is what ships.
Do not introduce a build system, a framework, or npm dependencies unless the
owner explicitly asks — the whole point of this site is zero-maintenance static
hosting.

The defining design choice: **content is Excel-driven.** Page structure lives in
HTML; the actual data (publications, students, news, etc.) lives in `.xlsx`
files under `data/`, read in the browser at runtime by SheetJS. The owner
updates content by editing spreadsheets, not code.

## 2. Directory layout

```
/                             repo root == site root (index.html must stay here)
├── AGENTS.md                 ← this file
├── README.md                 human-facing docs (owner's update/deploy guide)
├── index.html                Home page
├── experience.html           Full career timeline
├── teaching.html             Courses taught
├── publications.html         Searchable/filterable publication list
├── students-current.html     Currently-supervised students
├── students-graduated.html   Graduated students + past interns
├── awards.html               Awards & recognitions
├── news.html                 News/updates, grouped by year
├── resources.html            Public tools/datasets/tutorials (hardcoded)
├── protected.html            Password-gated resources (client-side only!)
├── contact.html              Contact details
├── style.css                 ALL styling (single file; theme vars at top)
├── script.js                 ALL behaviour (single file; ~760 lines)
├── vendor/
│   └── xlsx.full.min.js       SheetJS, bundled locally (reads .xlsx in-browser)
├── data/                      ★ CONTENT lives here — one .xlsx per domain
│   ├── publications.xlsx      (5 sheets: Journals/Conferences/Books/Book Chapters/Reports)
│   ├── students.xlsx
│   ├── experience.xlsx
│   ├── awards.xlsx
│   ├── teaching.xlsx
│   ├── news.xlsx
│   └── metrics.xlsx           (Year, Citations, Publications → home chart)
└── assets/
    ├── profile.jpg            portrait
    ├── medal-gold.jpg
    ├── CV_Jitendra_Kumar.pdf
    ├── social/                social-icon PNGs (scholar, orcid, researchid, researchgate, irins)
    └── students/              per-student photos (referenced by filename from students.xlsx)
```

External runtime dependencies loaded from CDN (not vendored): **Google Fonts**
(Inter / IBM Plex Mono) and **Chart.js 4.4.4** (jsDelivr, used only on
the home page for the metrics chart). SheetJS is vendored locally, not from CDN.

## 3. How the JS is organized (`script.js`)

One file, no modules. Structure, top to bottom:

1. Mobile nav toggle (`DOMContentLoaded` listener).
2. **Publications** — `PUB_URL`, constants, `formatEntry()` (per-type citation
   formatting), `badgesHtml()`, `loadPublications()`, `renderPublicationsPage()`
   (search + type filter), `renderHomeHighlights()`, `initPublications()`.
3. **Students** — `STUDENTS_URL`, `loadStudents()`, `renderStudents(rows, mode)`,
   `initStudents()`.
4. A `DOMContentLoaded` listener runs `initPublications()` + `initStudents()`.
5. **Shared helpers** — `loadSheet(url, sheetName)`, `withLoading()` wrapper.
6. **Experience / Awards / Teaching / News** — a `render*` + `init*` pair each.
7. **Metrics chart** — `renderMetricsChart()` / `initMetrics()` (Chart.js).
8. `initBioToggle()` — "Show More" on the home bio.
9. `initProtectedGate()` + `RESOURCE_PASSWORDS` array.
10. A second `DOMContentLoaded` listener runs experience/awards/teaching/news/
    metrics/bio/protected inits.

### Page-detection pattern (important)

`script.js` is loaded on **every** page and every `init*()` runs on every page.
Each `init*()` **guards on the presence of its container element** and no-ops if
absent:

```js
const container = document.getElementById("awards-list");
if (!container) return;   // not the awards page → do nothing
```

So the page a script targets is determined by which container `id` exists in that
page's HTML. **To wire a data-driven section into a page, add the right container
`id`; to disable one, remove it.** No routing, no per-page script tags.

### Container id → data source map

| Page | Container id(s) | init fn | Data |
|---|---|---|---|
| index.html | `experience-preview`, `news-preview`, `pub-highlights`, `metrics-chart`, `bio-more` | init(Experience/News/Publications/Metrics/Bio) | experience, news, publications, metrics |
| experience.html | `experience-list` | initExperience | experience.xlsx |
| teaching.html | `teaching-list` | initTeaching | teaching.xlsx |
| publications.html | `pub-list`, `pub-search` | initPublications | publications.xlsx |
| students-current.html | `student-list` (`data-mode="current"` default) | initStudents | students.xlsx |
| students-graduated.html | `student-list` `data-mode="graduated"` | initStudents | students.xlsx |
| awards.html | `awards-list` | initAwards | awards.xlsx |
| news.html | `news-list` | initNews | news.xlsx |
| protected.html | `protected-form/-password/-error/-gate/-content` | initProtectedGate | (none) |
| resources.html / contact.html | — | — | static HTML |

The same `experience.xlsx`/`news.xlsx`/`publications.xlsx` drive both the home
previews and the full pages; the preview-vs-full distinction is just which
container id is present (`*-preview` vs `*-list`). Students uses one file for both
pages, split by the `data-mode` attribute on `#student-list`.

## 4. Data model (the `.xlsx` files)

Each workbook column becomes a JS object key via `XLSX.utils.sheet_to_json(...,
{defval:""})`, so **column headers are the API** — renaming a header in Excel
breaks rendering. Every workbook also carries an **Instructions** sheet (human
notes; ignored by code). Key columns per file are tabulated in `README.md §
"What's Excel-driven"` — consult it before touching a spreadsheet or the code
that reads one.

Two behaviours worth knowing before editing:

- **Publications visibility:** only rows whose `Status` is `Published`,
  `Accepted`, or `Preprint` render (`VISIBLE_STATUSES` in `script.js`). `Status`
  is a filter and is **never displayed**. In-review/submitted work stays private
  in the same file. `formatEntry()` builds a different citation string per sheet
  type, so per-type columns (Journal/Conference/Publisher/etc.) matter.
- **Students Current vs Graduated:** the same file feeds both pages; a row's
  `Status` (`Ongoing` vs `Completed`/`Awarded`) decides which page it lands on.

## 5. Hard constraints & gotchas

- **Must be served over http(s).** Pages `fetch()` local `.xlsx` files; browsers
  block that from `file://`. So double-clicking `index.html` shows "Couldn't
  load" — expected, not a bug. To test locally: `python3 -m http.server` from the
  repo root, then open `http://localhost:8000`. GitHub Pages/Netlify/Vercel all
  work.
- **`protected.html` is NOT real security.** `RESOURCE_PASSWORDS` and the gated
  content are both in client-side code — anyone can view-source. Treat it as a
  soft deterrent only; never put anything genuinely sensitive behind it.
- **`index.html` must stay at repo root** for the GitHub Pages root URL to work.
- **Column headers are load-bearing.** Renaming an `.xlsx` header without updating
  the matching key in `script.js` silently drops that field.
- **Keep filenames stable.** Data URLs (`data/publications.xlsx`, etc.) are
  hardcoded in `script.js`; asset paths are hardcoded in HTML.
- **Theming** is done through CSS custom properties at the top of `style.css`
  (maroon `#8a1f2b` accent). Prefer editing the variables over scattering literal
  colors.
- **No secrets, no backend, no server code** exists or is possible here — it's a
  static site on free hosting.

## 6. Common edit recipes

- *Add/edit content (papers, students, news, awards, timeline, teaching):* edit
  the relevant `data/*.xlsx`, keep the filename and headers, done. No code change.
- *Add a student photo:* drop the image in `assets/students/`, put its filename in
  the student's `Image` column. Blank `Image` → initials-avatar fallback.
- *Update citation counts / the home chart:* edit `data/metrics.xlsx` (Google
  Scholar can't be scraped from a static page, so this table is manual).
- *Change bio / social links / contact / resources:* edit the HTML directly
  (`index.html`, `contact.html`, `resources.html`) — these are intentionally not
  Excel-driven.
- *Change protected-page passwords:* edit `RESOURCE_PASSWORDS` in `script.js`.
- *Add a new data-driven section to a page:* add a container `<div id="...">`,
  then either reuse an existing `init*` (match its expected id) or add a
  `render*`/`init*` pair following the guard-on-container pattern and register it
  in the appropriate `DOMContentLoaded` listener.

## 7. Deploy

Push to the default branch; GitHub Pages serves the repo root (Settings → Pages →
Deploy from branch → main → `/root`). Live within a minute or two. No CI, no build
artifacts to generate. See `README.md § Deploy` for first-time setup.


## 8. Homepage design (current state)

The site was redesigned to a **minimalist academic style** (modeled on typical
Jekyll academic homepages). Key facts for agents working on it:

- **Palette:** single maroon accent `#8a1f2b` on a white/black/gray base. No
  amber/green, no rainbow category tags — all tags/badges are maroon. Theme via
  the CSS variables at the top of `style.css`.
- **Fonts:** **Inter** for everything (Poppins was removed) + IBM Plex Mono for
  the rare mono use. `--serif` is aliased to the Inter stack.
- **Layout:** `--wrap` is `800px`; the site is a narrow, centered, single column.
  No cards/shadows on the homepage; interior cards are flattened to hairlines.
- **Nav:** the shared `.site-nav` is restyled as a flat text nav (no pills);
  `.brand-tag` is hidden. It's still one nav block per page — the site remains
  **multi-page**, each nav item is its own `.html` file.
- **Homepage (`index.html`, `<body class="minimal">`):** intro (name / role /
  affiliation / email·CV / social icons + round photo) → bio with Show More →
  **News** (static list, `#news-preview`) → **Experience** (light timeline,
  `#experience-preview`, capped at 3 with Show more) → **Bibliometrics** chart.
  The old "Recent Publications" home section was **removed**.
- **script.js specifics:** experience preview `limit: 3`; `renderNewsPreview`
  outputs a static `.updates-list` (the old auto-scroll ticker is gone); the
  metrics chart uses maroon `#8a1f2b` + slate `#7d8590` (not orange/purple).
