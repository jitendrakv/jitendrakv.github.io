// Mobile nav toggle
document.addEventListener("DOMContentLoaded", () => {
  const toggle = document.querySelector(".nav-toggle");
  const links = document.querySelector(".nav-links");
  if (toggle && links){
    toggle.addEventListener("click", () => links.classList.toggle("open"));
    links.querySelectorAll("a").forEach(a => a.addEventListener("click", () => links.classList.remove("open")));
  }
});

/* ============================================================
   PUBLICATIONS
   Source: data/publications.xlsx — one sheet per type.
   Only Status = Published / Accepted / Preprint are ever shown.
   Status itself is never displayed — it's a filter only.
   ============================================================ */
const PUB_URL = "data/publications.xlsx";
const VISIBLE_STATUSES = ["Published", "Accepted", "Preprint"];
const PUB_TYPES = ["Journals", "Conferences", "Books", "Book Chapters", "Reports"];
const PUB_TYPE_SINGULAR = {
  "Journals": "Journal", "Conferences": "Conference", "Books": "Book",
  "Book Chapters": "Book Chapter", "Reports": "Report"
};
const PUB_TYPE_CLASS = {
  "Journals": "tag-journal", "Conferences": "tag-conference", "Books": "tag-book",
  "Book Chapters": "tag-bookchapter", "Reports": "tag-report"
};

function cleanNum(v){
  if (v === "" || v === null || v === undefined) return null;
  const n = Number(v);
  return isNaN(n) ? null : n;
}

function formatEntry(type, row){
  const year = row.Year || "";
  let authors, title, venueCore;

  if (type === "Journals"){
    authors = row.Authors; title = row.Title;
    const vnp = [];
    if (row.Volume) vnp.push("Vol. " + row.Volume);
    if (row.Number) vnp.push("No. " + row.Number);
    const rest = [];
    if (vnp.length) rest.push(vnp.join(", "));
    if (row.Pages) rest.push(row.Pages);
    venueCore = `<b>${row.Journal || ""}</b>` + (rest.length ? " · " + rest.join(" · ") : "");
  } else if (type === "Conferences"){
    authors = row.Authors; title = row.Title;
    const rest = [];
    if (row.Place) rest.push(row.Place);
    if (row.Pages) rest.push("pp. " + row.Pages);
    venueCore = `<b>${row.Conference || ""}</b>` + (rest.length ? " · " + rest.join(" · ") : "");
  } else if (type === "Books"){
    authors = row["Authors/Editors"]; title = row.Title;
    const rest = [];
    if (row.Edition) rest.push(row.Edition + " Edition");
    venueCore = `<b>${row.Publisher || ""}</b>` + (rest.length ? " · " + rest.join(" · ") : "");
  } else if (type === "Book Chapters"){
    authors = row.Authors; title = row["Chapter Title"];
    const rest = [];
    if (row.Pages) rest.push("pp. " + row.Pages);
    venueCore = "in " + `<b>${row["Book Title"] || ""}</b>` + (row.Publisher ? " · " + row.Publisher : "") + (rest.length ? " · " + rest.join(" · ") : "");
  } else if (type === "Reports"){
    authors = row.Authors; title = row.Title;
    const rest = [];
    if (row.Identifier) rest.push(row.Identifier);
    venueCore = `<b>${row.Repository || ""}</b>` + (rest.length ? " · " + rest.join(" · ") : "");
  }

  const link = row.Link ? String(row.Link).trim() : "";
  const titleHtml = link
    ? `<a href="${link}" target="_blank" rel="noopener">${title}</a>`
    : title;

  const venueHtml = venueCore + (year ? " · " + year : "");
  return { authors, title: titleHtml, venueHtml, venueCore, year: cleanNum(year) || 0 };
}

function badgesHtml(row){
  const badges = [];
  if (row.Indexing) badges.push(`<span class="badge idx">${row.Indexing}</span>`);
  const ifNum = cleanNum(row["Impact Factor"]);
  if (ifNum !== null) badges.push(`<span class="badge if">IF ${ifNum}</span>`);
  const q = String(row.Quartile || "").trim().toUpperCase();
  if (q === "Q1") badges.push(`<span class="badge q1">Q1</span>`);
  else if (q === "Q2") badges.push(`<span class="badge q2">Q2</span>`);
  else if (q === "Q3") badges.push(`<span class="badge q3">Q3</span>`);
  else if (q === "Q4") badges.push(`<span class="badge q4">Q4</span>`);
  return badges.join("");
}

async function loadPublications(){
  const res = await fetch(PUB_URL);
  if (!res.ok) throw new Error("Could not load " + PUB_URL + " (HTTP " + res.status + ")");
  const buf = await res.arrayBuffer();
  const wb = XLSX.read(buf, { type: "array" });

  const data = {};
  PUB_TYPES.forEach(type => {
    if (!wb.SheetNames.includes(type)) { data[type] = []; return; }
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[type], { defval: "" });
    data[type] = rows.filter(r => VISIBLE_STATUSES.includes(String(r.Status || "").trim()));
  });
  return data;
}

function renderPublicationsPage(data){
  const container = document.getElementById("pub-list");
  const searchInput = document.getElementById("pub-search");
  const filterBtns = document.querySelectorAll(".pub-filter");
  if (!container) return;

  let activeType = "all";
  let query = "";

  function matches(type, row){
    if (!query) return true;
    const f = formatEntry(type, row);
    const hay = (f.authors + " " + f.title + " " + (row.Journal || row.Conference || row.Publisher || row.Repository || "")).toLowerCase();
    return hay.includes(query);
  }

  function draw(){
    container.innerHTML = "";

    const types = activeType === "all" ? PUB_TYPES : [activeType];
    let combined = [];
    types.forEach(type => {
      (data[type] || []).filter(r => matches(type, r)).forEach(row => combined.push({ type, row }));
    });

    // Sort: newest year first; ties broken by publication type order
    // (Journals -> Conferences -> Books -> Book Chapters -> Reports)
    combined.sort((a, b) => {
      const ay = Number(a.row.Year) || 0, by = Number(b.row.Year) || 0;
      if (by !== ay) return by - ay;
      return PUB_TYPES.indexOf(a.type) - PUB_TYPES.indexOf(b.type);
    });

    if (combined.length === 0){
      const p = document.createElement("div");
      p.className = "no-results";
      p.textContent = "No publications match your search.";
      container.appendChild(p);
      return;
    }

    const countEl = document.createElement("div");
    countEl.className = "pub-total-count";
    countEl.textContent = `${combined.length} publication${combined.length === 1 ? "" : "s"}`;
    container.appendChild(countEl);

    combined.forEach(({ type, row }) => {
      const f = formatEntry(type, row);
      const rowEl = document.createElement("div");
      rowEl.className = "pub-card-v";
      const pvLink = row.Link ? String(row.Link).trim() : "";
      const pvLinkLabel = type === "Books" ? "Publisher" : (type === "Reports" ? "Link" : "DOI");
      const pvVenue = f.venueCore + (f.year ? " · " + f.year : "") + (row["Other Info"] ? " · " + row["Other Info"] : "");
      rowEl.innerHTML = `
        <div class="pv-spine ${PUB_TYPE_CLASS[type]}">${PUB_TYPE_SINGULAR[type]}</div>
        <div class="pv-main">
          <div class="pv-textcol">
            <div class="pv-title">${f.title}</div>
            <div class="pv-authors">${f.authors}</div>
            <div class="pv-venue">${pvVenue}</div>
          </div>
          ${pvLink ? `<a class="pv-doi" href="${pvLink}" target="_blank" rel="noopener">${pvLinkLabel} ↗</a>` : ""}
        </div>
      `;
      container.appendChild(rowEl);
    });
  }

  filterBtns.forEach(btn => {
    btn.addEventListener("click", () => {
      filterBtns.forEach(b => b.classList.remove("active"));
      btn.classList.add("active");
      activeType = btn.dataset.type;
      draw();
    });
  });

  if (searchInput){
    searchInput.addEventListener("input", (e) => {
      query = e.target.value.trim().toLowerCase();
      draw();
    });
  }

  draw();
}

function renderHomeHighlights(data){
  const container = document.getElementById("pub-highlights");
  if (!container) return;
  const all = [];
  PUB_TYPES.forEach(type => (data[type] || []).forEach(row => all.push({ type, row })));
  all.sort((a, b) => (Number(b.row.Year) || 0) - (Number(a.row.Year) || 0));
  const top = all.slice(0, 3);

  container.innerHTML = top.map(({ type, row }) => {
    const f = formatEntry(type, row);
    return `
      <div class="pub-card">
        <div class="pub-card-top">
          <div class="authors">${f.authors}</div>
          <span class="type-tag ${PUB_TYPE_CLASS[type]}">${PUB_TYPE_SINGULAR[type]}</span>
        </div>
        <h3>${f.title}</h3>
        <div class="meta">${f.venueHtml}</div>
        ${badgesHtml(row) ? `<div class="badges">${badgesHtml(row)}</div>` : ""}
      </div>
    `;
  }).join("");
}

async function initPublications(){
  const pubList = document.getElementById("pub-list");
  const highlights = document.getElementById("pub-highlights");
  if (!pubList && !highlights) return;

  const target = pubList || highlights;
  const loadingMsg = document.createElement("div");
  loadingMsg.className = "pub-loading";
  loadingMsg.textContent = "Loading publications…";
  target.parentNode.insertBefore(loadingMsg, target);

  try{
    const data = await loadPublications();
    loadingMsg.remove();
    if (pubList) renderPublicationsPage(data);
    if (highlights) renderHomeHighlights(data);
  }catch(err){
    loadingMsg.textContent = "Couldn't load publications data (" + err.message + "). This page needs to be served over http(s) — it won't work by double-clicking the file directly.";
    loadingMsg.classList.add("pub-error");
    console.error(err);
  }
}

/* ============================================================
   STUDENTS
   Source: data/students.xlsx — one sheet "Students".
   Status: Ongoing -> Current Students page; Completed/Awarded -> Graduated Students page.
   ============================================================ */
const STUDENTS_URL = "data/students.xlsx";
const STUDENT_CATEGORY_ORDER = ["PhD", "Masters", "Undergraduate", "Research Intern"];
const STUDENT_CATEGORY_LABELS = {
  "PhD": "PhD Students",
  "Masters": "Masters Students",
  "Undergraduate": "Undergraduate Students",
  "Research Intern": "Research Interns"
};

async function loadStudents(){
  const res = await fetch(STUDENTS_URL);
  if (!res.ok) throw new Error("Could not load " + STUDENTS_URL + " (HTTP " + res.status + ")");
  const buf = await res.arrayBuffer();
  const wb = XLSX.read(buf, { type: "array", cellDates: true });
  const SHEET_CATEGORY = { "PhD": "PhD", "PG": "Masters", "UG": "Undergraduate", "Interns": "Research Intern" };
  const out = [];
  Object.keys(SHEET_CATEGORY).forEach(sheet => {
    if (!wb.SheetNames.includes(sheet)) return;
    const cat = SHEET_CATEGORY[sheet];
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheet], { defval: "" });
    rows.forEach(r => {
      if (!String(r.Name || "").trim()) return;
      out.push({
        Category: cat,
        Name: r.Name,
        Program: cat === "PhD" ? "PhD" : (r.Program || ""),
        Status: String(r.Status || "").trim(),
        Thesis: r["Research Area / Thesis Title"] || "",
        ThesisLink: r["Thesis Link"] || "",
        Affiliation: r["Current Affiliation"] || r["Affiliation"] || "",
        Image: r.Image || "",
        Notes: r.Notes || "",
        SupervisionTags: r["Supervision Tags"] || "",
        Year: r.Year || "",
        AwardDate: r["Award Date"] || ""
      });
    });
  });
  return out;
}

function studentYear(s){
  if (s.Category === "PhD" && s.AwardDate){
    const d = s.AwardDate;
    if (d instanceof Date && !isNaN(d)) return d.getFullYear();
    const m = String(d).match(/(\d{4})/); return m ? +m[1] : 0;
  }
  const m = String(s.Year || "").match(/(\d{4})/); return m ? +m[1] : 0;
}

function studentYearLabel(s){
  if (s.Status === "Ongoing") return "Ongoing";
  if (s.Category === "PhD"){ const y = studentYear(s); return y ? String(y) : ""; }
  return s.Year ? String(s.Year) : "";
}

function escAttr(v){
  return String(v).replace(/&/g,"&amp;").replace(/"/g,"&quot;").replace(/</g,"&lt;").replace(/>/g,"&gt;");
}

function studentCardHtml(s, wide){
  const catSlug = String(s.Category||'').toLowerCase().replace(/\s+/g,'');
  const awarded = s.Status === "Awarded" ? `<span class="s-awd">Awarded</span>` : "";
  const topic = s.Thesis;
  const thesisLink = s.ThesisLink ? String(s.ThesisLink).trim() : "";
  const topicHtml = topic
    ? (thesisLink ? `<a href="${thesisLink}" target="_blank" rel="noopener">${topic}</a>` : topic)
    : "";
  const imgName = s.Image ? String(s.Image).trim() : "";
  const initials = String(s.Name || "").split(/\s+/).map(w => w[0]).slice(0, 2).join("").toUpperCase();
  const avatar = imgName
    ? `<img class="s-avatar" src="assets/students/${imgName}" alt="${s.Name}">`
    : `<div class="s-avatar s-avatar-fallback">${initials}</div>`;
  const sub = [s.Program, studentYearLabel(s)].filter(Boolean).join(" \u00b7 ");
  const tip = [s.SupervisionTags, s.Notes].filter(Boolean).join(" \u00b7 ");
  const titleAttr = tip ? ` title="${escAttr(tip)}"` : "";
  const affHtml = s.Affiliation ? `<div class="s-aff"><span class="lab">Next/Now at </span>${s.Affiliation}</div>` : "";

  if (wide){
    return `
      <div class="scard scard-wide cat-${catSlug}"${titleAttr}>
        ${avatar}
        <div class="s-body">
          <div class="s-name-row">
            <span class="s-name">${s.Name}${awarded}</span>
            ${sub ? `<span class="s-year">${sub}</span>` : ""}
          </div>
          ${topicHtml ? `<div class="s-thesis-wide">${topicHtml}</div>` : ""}
          ${affHtml}
        </div>
      </div>
    `;
  }
  return `
    <div class="scard cat-${catSlug}"${titleAttr}>
      ${avatar}
      <div class="s-body">
        <div class="s-name">${s.Name}${awarded}</div>
        ${sub ? `<div class="s-sub">${sub}</div>` : ""}
        ${topicHtml ? `<div class="s-thesis">${topicHtml}</div>` : ""}
        ${affHtml}
      </div>
    </div>
  `;
}

function renderStudents(rows, mode){
  const container = document.getElementById("student-list");
  if (!container) return;

  const wanted = mode === "current" ? ["Ongoing"] : ["Completed", "Awarded"];
  const filtered = rows.filter(r => wanted.includes(String(r.Status || "").trim()));

  container.innerHTML = "";
  let any = false;

  STUDENT_CATEGORY_ORDER.forEach(cat => {
    const items = filtered.filter(r => String(r.Category || "").trim() === cat);
    if (items.length === 0) return;
    any = true;

    items.sort((a, b) => {
      const ay = studentYear(a), by = studentYear(b);
      if (by !== ay) return by - ay;
      return String(a.Name).localeCompare(String(b.Name));
    });

    const h = document.createElement("div");
    h.className = "pub-group-title";
    h.innerHTML = `${STUDENT_CATEGORY_LABELS[cat]} <span class="pub-count">(${items.length})</span>`;
    container.appendChild(h);

    const wide = (mode !== "current" && cat === "PhD");
    const grid = document.createElement("div");
    grid.className = "student-grid" + (wide ? " one-col" : " two-col");
    grid.innerHTML = items.map(s => studentCardHtml(s, wide)).join("");
    container.appendChild(grid);
  });

  if (!any){
    container.innerHTML = `<div class="no-results">No students found.</div>`;
  }
}

async function initStudents(){
  const container = document.getElementById("student-list");
  if (!container) return;
  const mode = container.dataset.mode || "current";

  const loadingMsg = document.createElement("div");
  loadingMsg.className = "pub-loading";
  loadingMsg.textContent = "Loading students…";
  container.parentNode.insertBefore(loadingMsg, container);

  try{
    const rows = await loadStudents();
    loadingMsg.remove();
    renderStudents(rows, mode);
  }catch(err){
    loadingMsg.textContent = "Couldn't load student data (" + err.message + "). This page needs to be served over http(s) — it won't work by double-clicking the file directly.";
    loadingMsg.classList.add("pub-error");
    console.error(err);
  }
}

document.addEventListener("DOMContentLoaded", () => {
  initPublications();
  initStudents();
});

/* ============================================================
   Generic small-sheet loader (Experience, Awards, Teaching, News, Metrics)
   ============================================================ */
async function loadSheet(url, sheetName){
  const res = await fetch(url);
  if (!res.ok) throw new Error("Could not load " + url + " (HTTP " + res.status + ")");
  const buf = await res.arrayBuffer();
  const wb = XLSX.read(buf, { type: "array" });
  const name = wb.SheetNames.includes(sheetName) ? sheetName : wb.SheetNames.find(n => n !== "Instructions");
  return XLSX.utils.sheet_to_json(wb.Sheets[name], { defval: "" });
}

async function withLoading(containerEl, loadFn, renderFn, label){
  const loadingMsg = document.createElement("div");
  loadingMsg.className = "pub-loading";
  loadingMsg.textContent = `Loading ${label}…`;
  containerEl.parentNode.insertBefore(loadingMsg, containerEl);
  try{
    const rows = await loadFn();
    loadingMsg.remove();
    renderFn(rows);
  }catch(err){
    loadingMsg.textContent = `Couldn't load ${label} (` + err.message + "). This needs to be served over http(s).";
    loadingMsg.classList.add("pub-error");
    console.error(err);
  }
}

/* ---------------- Experience / Career Timeline ---------------- */
function renderExperience(rows, opts = {}){
  const container = document.getElementById(opts.containerId || "experience-list");
  if (!container) return;
  const limit = opts.limit;

  rows.sort((a, b) => (Number(b.Year) || 0) - (Number(a.Year) || 0));

  // group consecutive rows by year
  const years = [];
  rows.forEach(r => {
    const y = r.Year;
    let bucket = years.find(g => g.year === y);
    if (!bucket){ bucket = { year: y, items: [] }; years.push(bucket); }
    bucket.items.push(r);
  });

  const shown = limit ? years.slice(0, limit) : years;
  const hidden = limit ? years.slice(limit) : [];

  function itemsHtml(list){
    return list.map(g => `
      <div class="tl-item">
        <div class="tl-date">${g.year}</div>
        ${g.items.map(it => `
          <h3>${it.Title}</h3>
          ${it.Institution ? `<div class="tl-org">${it.Institution}</div>` : ""}
          ${it.Note ? `<div class="tl-detail">${it.Note}</div>` : ""}
        `).join("")}
      </div>
    `).join("");
  }

  container.innerHTML = `<div class="timeline">${itemsHtml(shown)}</div>`;

  if (hidden.length){
    const btn = document.createElement("button");
    btn.className = "show-more-btn";
    btn.textContent = "Show more...";
    let expanded = false;
    btn.addEventListener("click", () => {
      expanded = !expanded;
      container.querySelector(".timeline").innerHTML = itemsHtml(expanded ? years : shown);
      btn.textContent = expanded ? "Show less..." : "Show more...";
    });
    container.appendChild(btn);
  }
}

function initExperience(){
  const full = document.getElementById("experience-list");
  const preview = document.getElementById("experience-preview");
  if (!full && !preview) return;
  const target = full || preview;
  withLoading(target, () => loadSheet("data/experience.xlsx", "Experience"), (rows) => {
    if (full) renderExperience(rows, { containerId: "experience-list" });
    if (preview) renderExperience(rows, { containerId: "experience-preview", limit: 3 });
  }, "experience");
}

/* ---------------- Awards ---------------- */
function renderAwards(rows){
  const container = document.getElementById("awards-list");
  if (!container) return;
  rows.sort((a, b) => (Number(b.Year) || 0) - (Number(a.Year) || 0));
  container.innerHTML = rows.map(a => `
    <div class="award-box">
      <div class="award-medal">🏅</div>
      <div>
        <div class="award-year">${a.Year}</div>
        <h3>${a.Title}</h3>
        <p>${a.Description}</p>
      </div>
    </div>
  `).join("");
}

function initAwards(){
  const container = document.getElementById("awards-list");
  if (!container) return;
  withLoading(container, () => loadSheet("data/awards.xlsx", "Awards"), renderAwards, "awards");
}

/* ---------------- Teaching ---------------- */
function renderTeaching(rows){
  const container = document.getElementById("teaching-list");
  if (!container) return;

  const current = rows.filter(r => r.Status === "Current");
  const previous = rows.filter(r => r.Status !== "Current");

  const byInst = {};
  previous.forEach(r => {
    const inst = r.Institution || "Other";
    if (!byInst[inst]) byInst[inst] = [];
    byInst[inst].push(r.Course);
  });

  let html = "";
  if (current.length){
    html += `<div class="teaching-block"><h3>Current</h3><ul>${current.map(c => `<li>${c.Course}</li>`).join("")}</ul></div>`;
  }
  Object.keys(byInst).forEach(inst => {
    html += `<div class="teaching-block"><h3>@ ${inst}</h3><ul>${byInst[inst].map(c => `<li>${c}</li>`).join("")}</ul></div>`;
  });

  container.innerHTML = html;
}

function initTeaching(){
  const container = document.getElementById("teaching-list");
  if (!container) return;
  withLoading(container, () => loadSheet("data/teaching.xlsx", "Teaching"), renderTeaching, "teaching");
}

/* ---------------- News ---------------- */
function newsItemHtml(n){
  const headline = n.Link
    ? `<a href="${n.Link}" target="_blank" rel="noopener">${n.Headline}</a>`
    : n.Headline;
  const slug = String(n.Type||"").trim().toLowerCase();
  return `<li><span class="news-tag nt-${slug}">[${n.Type}]</span><span class="news-body"><i>${headline}</i> ${n.Detail}</span></li>`;
}

function renderNewsFull(rows){
  const container = document.getElementById("news-list");
  if (!container) return;
  rows.sort((a, b) => (Number(b.Year) || 0) - (Number(a.Year) || 0));
  const byYear = {};
  rows.forEach(n => { (byYear[n.Year] = byYear[n.Year] || []).push(n); });
  container.innerHTML = Object.keys(byYear).sort((a,b) => b - a).map(y => `
    <div class="pub-group-title">${y}</div>
    <ol class="news-ol">${byYear[y].map(newsItemHtml).join("")}</ol>
  `).join("");
}

function renderNewsPreview(rows, count){
  const container = document.getElementById("news-preview");
  if (!container) return;

  // Home page "Updates": auto-scrolling ticker (marquee, bottom-to-top) of the most recent items.
  rows.sort((a, b) => (Number(b.Year) || 0) - (Number(a.Year) || 0));
  const recent = rows.slice(0, count);
  if (!recent.length){
    container.innerHTML = `<div class="no-results">No updates yet.</div>`;
    return;
  }
  const itemHtml = (n) => {
    const headline = n.Link
      ? `<a href="${n.Link}" target="_blank" rel="noopener">${n.Headline}</a>`
      : `<span class="u-head">${n.Headline}</span>`;
    const slug = String(n.Type||"").trim().toLowerCase();
    return `<li class="update-item"><span class="u-type nt-${slug}">[${n.Type}]</span><div class="u-body">${headline} <span class="u-detail">${n.Detail}</span></div></li>`;
  };
  container.innerHTML = `
    <ul class="updates-list">${recent.map(itemHtml).join("")}</ul>
    <div class="page-links" style="margin-top:18px;"><a href="news.html">Show more...</a></div>
  `;
}

function initNews(){
  const full = document.getElementById("news-list");
  const preview = document.getElementById("news-preview");
  if (!full && !preview) return;
  const target = full || preview;
  withLoading(target, () => loadSheet("data/news.xlsx", "News"), (rows) => {
    if (full) renderNewsFull(rows.slice());
    if (preview) renderNewsPreview(rows.slice(), 4);
  }, "news");
}

/* ---------------- Metrics chart (citations / publications per year) ----------------
   Reads data/metrics.xlsx (Year, Citations, Publications) and draws a dual-axis
   line chart with Chart.js — Publications on the left axis, Citations on the right,
   matching the reference site's layout. Falls back to a plain table if Chart.js
   didn't load (e.g. offline / CDN blocked). */
function renderMetricsChart(rows){
  const container = document.getElementById("metrics-chart");
  if (!container) return;
  rows.sort((a, b) => (Number(a.Year) || 0) - (Number(b.Year) || 0));

  const years = rows.map(r => r.Year);
  const citations = rows.map(r => Number(r.Citations) || 0);
  const publications = rows.map(r => Number(r.Publications) || 0);

  if (typeof Chart === "undefined"){
    // Offline fallback: simple data table, still fully derived from the xlsx file.
    container.innerHTML = `
      <div class="no-results">Chart library unavailable — showing raw data from metrics.xlsx.</div>
      <table style="width:100%; border-collapse:collapse; margin-top:12px;">
        <tr><th style="text-align:left;">Year</th>${years.map(y => `<td style="text-align:center;">${y}</td>`).join("")}</tr>
        <tr><th style="text-align:left;">Publications</th>${publications.map(v => `<td style="text-align:center;">${v}</td>`).join("")}</tr>
        <tr><th style="text-align:left;">Citations</th>${citations.map(v => `<td style="text-align:center;">${v}</td>`).join("")}</tr>
      </table>`;
    return;
  }

  container.innerHTML = `
    <div class="metrics-chart-wrap">
      <canvas id="metrics-canvas" height="120"></canvas>
    </div>
    <div class="metrics-source-note">Based on Google Scholar data</div>
  `;
  const ctx = document.getElementById("metrics-canvas").getContext("2d");
  new Chart(ctx, {
    type: "line",
    data: {
      labels: years,
      datasets: [
        {
          label: "Publications",
          data: publications,
          borderColor: "#8a1f2b",
          backgroundColor: "#8a1f2b",
          pointBackgroundColor: "#8a1f2b",
          pointRadius: 4,
          tension: 0.25,
          yAxisID: "yPub"
        },
        {
          label: "Citations",
          data: citations,
          borderColor: "#7d8590",
          backgroundColor: "#7d8590",
          pointBackgroundColor: "#7d8590",
          pointRadius: 4,
          tension: 0.25,
          yAxisID: "yCit"
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      interaction: { mode: "index", intersect: false },
      plugins: {
        legend: { position: "bottom", labels: { boxWidth: 14, usePointStyle: true } }
      },
      scales: {
        x: { title: { display: true, text: "Year" }, grid: { display: false } },
        yPub: {
          type: "linear", position: "left", beginAtZero: true,
          title: { display: true, text: "Publication (#)" }
        },
        yCit: {
          type: "linear", position: "right", beginAtZero: true,
          title: { display: true, text: "Citation (#)" },
          grid: { drawOnChartArea: false }
        }
      }
    }
  });
}

function initMetrics(){
  const container = document.getElementById("metrics-chart");
  if (!container) return;
  withLoading(container, () => loadSheet("data/metrics.xlsx", "Metrics"), renderMetricsChart, "metrics");
}

/* ---------------- Bio "show more" toggle ---------------- */
function initBioToggle(){
  const btns = document.querySelectorAll(".bio-toggle");
  btns.forEach(btn => {
    btn.addEventListener("click", () => {
      const extra = document.getElementById(btn.dataset.target);
      if (!extra) return;
      const isOpen = extra.classList.toggle("open");
      btn.textContent = isOpen ? "Show less..." : "Show more...";
    });
  });
}

/* ---------------- Protected resources page (client-side deterrent, NOT real security) ---------------- */
const RESOURCE_PASSWORDS = ["123456", "234567", "345678"];
function initProtectedGate(){
  const form = document.getElementById("protected-form");
  if (!form) return;
  const input = document.getElementById("protected-password");
  const error = document.getElementById("protected-error");
  const gate = document.getElementById("protected-gate");
  const content = document.getElementById("protected-content");

  form.addEventListener("submit", (e) => {
    e.preventDefault();
    if (RESOURCE_PASSWORDS.includes(input.value.trim())){
      gate.style.display = "none";
      content.style.display = "block";
      error.style.display = "none";
    } else {
      error.style.display = "block";
    }
  });
}

document.addEventListener("DOMContentLoaded", () => {
  initExperience();
  initAwards();
  initTeaching();
  initNews();
  initMetrics();
  initBioToggle();
  initProtectedGate();
});
