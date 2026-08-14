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
  let authors, title, venueHtml;

  if (type === "Journals"){
    authors = row.Authors; title = row.Title;
    let vnp = [];
    if (row.Volume) vnp.push("Vol. " + row.Volume);
    if (row.Number) vnp.push("No. " + row.Number);
    const rest = [];
    if (vnp.length) rest.push(vnp.join(", "));
    if (row.Pages) rest.push(row.Pages);
    if (year) rest.push(year);
    venueHtml = `<b>${row.Journal || ""}</b>` + (rest.length ? " · " + rest.join(" · ") : "");
  } else if (type === "Conferences"){
    authors = row.Authors; title = row.Title;
    const rest = [];
    if (row.Place) rest.push(row.Place);
    if (row.Pages) rest.push("pp. " + row.Pages);
    if (year) rest.push(year);
    venueHtml = `<b>${row.Conference || ""}</b>` + (rest.length ? " · " + rest.join(" · ") : "");
  } else if (type === "Books"){
    authors = row["Authors/Editors"]; title = row.Title;
    const rest = [];
    if (row.Edition) rest.push(row.Edition + " Edition");
    if (year) rest.push(year);
    venueHtml = `<b>${row.Publisher || ""}</b>` + (rest.length ? " · " + rest.join(" · ") : "");
  } else if (type === "Book Chapters"){
    authors = row.Authors; title = row["Chapter Title"];
    const rest = [];
    if (row.Pages) rest.push("pp. " + row.Pages);
    if (year) rest.push(year);
    venueHtml = "in " + `<b>${row["Book Title"] || ""}</b>` + (row.Publisher ? " · " + row.Publisher : "") + (rest.length ? " · " + rest.join(" · ") : "");
  } else if (type === "Reports"){
    authors = row.Authors; title = row.Title;
    const rest = [];
    if (row.Identifier) rest.push(row.Identifier);
    if (year) rest.push(year);
    venueHtml = `<b>${row.Repository || ""}</b>` + (rest.length ? " · " + rest.join(" · ") : "");
  }

  const link = row.Link ? String(row.Link).trim() : "";
  const titleHtml = link
    ? `<a href="${link}" target="_blank" rel="noopener">${title}</a>`
    : title;

  return { authors, title: titleHtml, venueHtml, year: cleanNum(year) || 0 };
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
      rowEl.className = "pub-row";
      rowEl.innerHTML = `
        <div class="pub-row-top">
          <div class="authors">${f.authors}</div>
          <span class="type-tag ${PUB_TYPE_CLASS[type]}">${PUB_TYPE_SINGULAR[type]}</span>
        </div>
        <div class="titletext">${f.title}</div>
        <div class="venueline">${f.venueHtml}</div>
        ${badgesHtml(row) ? `<div class="badges">${badgesHtml(row)}</div>` : ""}
        ${row["Other Info"] ? `<div class="other-info">${row["Other Info"]}</div>` : ""}
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
  const wb = XLSX.read(buf, { type: "array" });
  const sheetName = wb.SheetNames.includes("Students") ? "Students" : wb.SheetNames[0];
  return XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { defval: "" });
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
      const ay = parseInt(String(a.Year).slice(0, 4)) || 0;
      const by = parseInt(String(b.Year).slice(0, 4)) || 0;
      if (by !== ay) return by - ay;
      return String(a.Name).localeCompare(String(b.Name));
    });

    const h = document.createElement("div");
    h.className = "pub-group-title";
    h.innerHTML = `${STUDENT_CATEGORY_LABELS[cat]} <span class="pub-count">(${items.length})</span>`;
    container.appendChild(h);

    items.forEach(s => {
      const row = document.createElement("div");
      row.className = "student-row";
      const statusTag = s.Status === "Awarded" ? `<span class="type-tag awarded">Awarded</span>` : "";
      const topic = s["Research Area / Thesis Title"];
      const thesisLink = s["Thesis Link"] ? String(s["Thesis Link"]).trim() : "";
      const topicHtml = topic
        ? (thesisLink ? `<a href="${thesisLink}" target="_blank" rel="noopener">${topic}</a>` : topic)
        : "";
      const imgName = s.Image ? String(s.Image).trim() : "";
      const initials = String(s.Name || "").split(/\s+/).map(w => w[0]).slice(0, 2).join("").toUpperCase();
      const photoHtml = imgName
        ? `<img class="student-photo" src="assets/students/${imgName}" alt="${s.Name}">`
        : `<div class="student-photo student-photo-fallback">${initials}</div>`;

      row.innerHTML = `
        ${photoHtml}
        <div class="student-body">
          <div class="student-row-top">
            <div class="student-name">${s.Name}</div>
            <div class="student-tags">
              <span class="type-tag tag-${String(s.Category||'').toLowerCase().replace(/\s+/g,'')}">${s.Program}</span>
              ${statusTag}
            </div>
          </div>
          ${topicHtml ? `<div class="student-topic">${topicHtml}</div>` : ""}
          <div class="student-meta">
            ${s.Year ? `<span>${s.Year}</span>` : ""}
            ${s.Affiliation ? `<span>${s.Affiliation}</span>` : ""}
            ${s.Notes ? `<span class="student-notes">${s.Notes}</span>` : ""}
          </div>
        </div>
      `;
      container.appendChild(row);
    });
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
    if (preview) renderExperience(rows, { containerId: "experience-preview", limit: 5 });
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
  return `<li><span class="news-type">${n.Type}</span> <i>${headline}</i> ${n.Detail}</li>`;
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
  rows.sort((a, b) => (Number(b.Year) || 0) - (Number(a.Year) || 0));
  container.innerHTML = `<ol class="news-ol">${rows.slice(0, count).map(newsItemHtml).join("")}</ol>`;
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

/* ---------------- Metrics chart (citations / publications per year) ---------------- */
function renderMetricsChart(rows){
  const container = document.getElementById("metrics-chart");
  if (!container) return;
  rows.sort((a, b) => (Number(a.Year) || 0) - (Number(b.Year) || 0));

  const W = 720, H = 260, PAD_L = 46, PAD_R = 16, PAD_T = 16, PAD_B = 30;
  const plotW = W - PAD_L - PAD_R, plotH = H - PAD_T - PAD_B;

  const maxC = Math.max(...rows.map(r => Number(r.Citations) || 0), 1);
  const maxP = Math.max(...rows.map(r => Number(r.Publications) || 0), 1);

  const x = i => PAD_L + (i / (rows.length - 1)) * plotW;
  const yC = v => PAD_T + plotH - (v / maxC) * plotH;
  const yP = v => PAD_T + plotH - (v / maxP) * plotH;

  const citLine = rows.map((r, i) => `${x(i)},${yC(Number(r.Citations) || 0)}`).join(" ");
  const pubLine = rows.map((r, i) => `${x(i)},${yP(Number(r.Publications) || 0)}`).join(" ");

  const labels = rows.map((r, i) => (i % 2 === 0 || rows.length < 10)
    ? `<text x="${x(i)}" y="${H - 8}" font-size="10" text-anchor="middle" fill="var(--faint)">${r.Year}</text>` : "").join("");

  container.innerHTML = `
    <svg viewBox="0 0 ${W} ${H}" class="metrics-svg">
      <line x1="${PAD_L}" y1="${PAD_T}" x2="${PAD_L}" y2="${H-PAD_B}" stroke="var(--line-2)" />
      <line x1="${PAD_L}" y1="${H-PAD_B}" x2="${W-PAD_R}" y2="${H-PAD_B}" stroke="var(--line-2)" />
      <polyline points="${citLine}" fill="none" stroke="var(--accent)" stroke-width="2.2" />
      <polyline points="${pubLine}" fill="none" stroke="#333" stroke-width="2" stroke-dasharray="4 3" />
      ${labels}
    </svg>
    <div class="metrics-legend">
      <span><i class="dot" style="background:var(--accent)"></i> Citations (cumulative)</span>
      <span><i class="dot dash"></i> Publications / year</span>
    </div>
  `;
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
