// Mobile nav toggle
/* ---------------- Recognition & Service ---------------- */
async function loadWorkbookSheets(url){
  const res = await fetch(url);
  if (!res.ok) throw new Error("Could not load " + url + " (HTTP " + res.status + ")");
  const buf = await res.arrayBuffer();
  const wb = XLSX.read(buf, { type: "array" });
  const out = {};
  wb.SheetNames.forEach(n => { out[n] = XLSX.utils.sheet_to_json(wb.Sheets[n], { defval: "" }); });
  return out;
}

function svcHide(id){ const el = document.getElementById(id); if (el) el.style.display = "none"; }

function renderSvcAwards(rows){
  const el = document.getElementById("svc-awards");
  if (!el) return;
  rows = (rows || []).filter(r => String(r.Title || "").trim());
  if (!rows.length){ svcHide("block-awards"); return; }
  rows.sort((a, b) => (Number(b.Year) || 0) - (Number(a.Year) || 0));
  document.getElementById("cnt-awards").textContent = `(${rows.length})`;
  el.innerHTML = rows.map(a => `
    <div class="svc-award">
      <div class="svc-medal">\ud83c\udfc5</div>
      <div>
        <div class="svc-aw-year">${a.Year || ""}</div>
        <div class="svc-aw-title">${a.Title || ""}</div>
        ${a.Description ? `<div class="svc-aw-sub">${a.Description}</div>` : ""}
      </div>
    </div>`).join("");
}

function renderSvcTalks(rows){
  const el = document.getElementById("svc-talks");
  if (!el) return;
  rows = (rows || []).filter(r => String(r.Title || "").trim());
  if (!rows.length){ svcHide("block-talks"); return; }
  rows.sort((a, b) => (Number(b.Year) || 0) - (Number(a.Year) || 0));
  document.getElementById("cnt-talks").textContent = `(${rows.length})`;
  el.innerHTML = rows.map(t => {
    const when = t.Date || t.Year || "";
    const link = t.Link ? String(t.Link).trim() : "";
    const title = link ? `<a href="${link}" target="_blank" rel="noopener">${t.Title}</a>` : t.Title;
    return `
    <div class="svc-card svc-talk">
      <div class="svc-c-top"><span class="svc-c-title">${title}</span>${when ? `<span class="svc-c-when">${when}</span>` : ""}</div>
      ${t.Venue ? `<div class="svc-c-sub">${t.Venue}</div>` : ""}
    </div>`;
  }).join("");
}

function renderSvcEditorial(rows){
  const el = document.getElementById("svc-editorial");
  if (!el) return;
  rows = (rows || []).filter(r => String(r.Title || "").trim());
  if (!rows.length){ svcHide("block-editorial"); return; }
  rows.sort((a, b) => (Number(b.Year) || 0) - (Number(a.Year) || 0));
  document.getElementById("cnt-editorial").textContent = `(${rows.length})`;
  el.innerHTML = rows.map(e => {
    const link = e.Link ? String(e.Link).trim() : "";
    const title = link ? `<a href="${link}" target="_blank" rel="noopener">${e.Title}</a>` : e.Title;
    return `
    <div class="svc-card svc-editorial">
      ${e.Role ? `<div class="svc-c-role">${e.Role}</div>` : ""}
      <div class="svc-c-title">${title}</div>
      ${e.Venue ? `<div class="svc-c-sub">${e.Venue}</div>` : ""}
    </div>`;
  }).join("");
}

function renderSvcPanels(editor, reviewer, members){
  const el = document.getElementById("svc-panels");
  if (!el) return;
  function chipPanel(cls, title, rows, highlight){
    rows = (rows || []).filter(r => String(r.Name || "").trim());
    if (!rows.length) return "";
    const chips = rows.map(r => {
      const strong = highlight && String(r.Highlight || "").trim().toLowerCase() === "yes";
      return `<span class="svc-chip${strong ? " strong" : ""}">${r.Name}</span>`;
    }).join("");
    return `<div class="svc-panel ${cls}"><h3>${title}</h3><div class="svc-chips">${chips}</div></div>`;
  }
  const html = chipPanel("svc-p-editor", "Editorial Board Member", editor, false)
             + chipPanel("svc-p-reviewer", "Technical Committee Member", reviewer, false)
             + chipPanel("svc-p-member", "SocietyMemberships", members, true);
  if (!html){ svcHide("block-service"); return; }
  el.innerHTML = html;
}

async function initServices(){
  const root = document.getElementById("services-root");
  if (!root) return;
  try{ renderSvcAwards(await loadSheet("data/awards.xlsx", "Awards")); }
  catch(err){ svcHide("block-awards"); console.error(err); }
  try{
    const sh = await loadWorkbookSheets("data/service.xlsx");
    renderSvcTalks(sh.Talks || []);
    renderSvcEditorial(sh.Editorial || []);
    renderSvcPanels(sh.Editor || [], sh.Reviewer || [], sh.Memberships || []);
  }catch(err){
    ["block-talks","block-editorial","block-service"].forEach(svcHide);
    console.error(err);
  }
}

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
const PUB_TYPES = ["Journals", "Conferences", "Books", "Chapters", "Technical Reports"];
const PUB_TYPE_SINGULAR = {
  "Journals": "Journal", "Conferences": "Conference", "Books": "Book",
  "Chapters": "Book Chapter", "Technical Reports": "Report"
};
const PUB_TYPE_CLASS = {
  "Journals": "tag-journal", "Conferences": "tag-conference", "Books": "tag-book",
  "Chapters": "tag-bookchapter", "Technical Reports": "tag-report"
};

// Show published/accepted/preprint (and Books/Reports which carry no status);
// hide in-progress statuses like "In Review" / "Revision Submitted".
function pubVisible(type, row){
  if (type === "Books" || type === "Technical Reports") return true;
  const s = String(row.Status || "").trim();
  if (!s) return true;
  return /^(published|accepted|preprint)/i.test(s);
}
function statusLabel(s){
  s = String(s || "").trim();
  if (!s || /^published/i.test(s)) return "";
  if (/^(accepted|preprint)/i.test(s)) return s;
  return "";
}

function cleanNum(v){
  if (v === "" || v === null || v === undefined) return null;
  const n = Number(v);
  return isNaN(n) ? null : n;
}

function formatEntry(type, row){
  const year = row.Year || "";
  const url = row.URL ? String(row.URL).trim() : "";
  const pages = (() => {
    const s = row.PageStart, e = row.PageEnd;
    if (s && e) return "pp. " + s + "\u2013" + e;
    if (s) return "pp. " + s;
    return "";
  })();
  let authors, title, venueCore, statusNote = "";

  if (type === "Journals"){
    authors = row.Authors; title = row.Title;
    const rest = [];
    if (row.Volume) rest.push("Vol. " + row.Volume);
    if (pages) rest.push(pages);
    venueCore = `<b>${row.Journal || ""}</b>` + (rest.length ? " \u00b7 " + rest.join(" \u00b7 ") : "");
    statusNote = statusLabel(row.Status);
  } else if (type === "Conferences"){
    authors = row.Authors; title = row.Title;
    const rest = [];
    if (pages) rest.push(pages);
    venueCore = `<b>${row.ConferenceName || ""}</b>` + (rest.length ? " \u00b7 " + rest.join(" \u00b7 ") : "");
    statusNote = statusLabel(row.Status);
  } else if (type === "Books"){
    const isEd = String(row.IsEditor || "").trim().toLowerCase() === "yes";
    authors = (row.Authors || "") + (isEd ? " (eds.)" : "");
    title = row.Title;
    const rest = [];
    if (row.Edition) rest.push(row.Edition + " Edition");
    venueCore = `<b>${row.Publisher || ""}</b>` + (rest.length ? " \u00b7 " + rest.join(" \u00b7 ") : "");
    statusNote = row["Status/Note"] ? String(row["Status/Note"]).trim() : "";
  } else if (type === "Chapters"){
    authors = row.Authors; title = row.Title;
    const rest = [];
    if (pages) rest.push(pages);
    venueCore = "in " + `<b>${row.BookTitle || ""}</b>` + (row.Publisher ? " \u00b7 " + row.Publisher : "") + (rest.length ? " \u00b7 " + rest.join(" \u00b7 ") : "");
    statusNote = statusLabel(row.Status);
  } else if (type === "Technical Reports"){
    authors = row.Authors; title = row.Title;
    venueCore = `<b>${row.Source || ""}</b>`;
  }

  const titleHtml = url ? `<a href="${url}" target="_blank" rel="noopener">${title}</a>` : title;
  const venueHtml = venueCore + (year ? " \u00b7 " + year : "");
  return { authors, title: titleHtml, venueCore, venueHtml, year: cleanNum(year) || 0, statusNote };
}

function badgesHtml(row){
  const badges = [];
  if (row.Indexing) badges.push(`<span class="badge idx">${row.Indexing}</span>`);
  const ifNum = cleanNum(row.ImpactFactor);
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
    data[type] = rows.filter(r => pubVisible(type, r));
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
    const hay = (f.authors + " " + f.title + " " + (row.Journal || row.ConferenceName || row.Publisher || row.Source || row.BookTitle || "")).toLowerCase();
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
      const pvLink = row.URL ? String(row.URL).trim() : "";
      const pvLinkLabel = type === "Books" ? "Publisher" : (type === "Technical Reports" ? "Link" : "DOI");
      const pvVenue = f.venueCore + (f.year ? " \u00b7 " + f.year : "") + (f.statusNote ? ` \u00b7 <span class="pv-status">${f.statusNote}</span>` : "");
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
        Institute: r.Institute || "",
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

function abbrevInstitute(v){
  const s = String(v || "").replace(/,\s*India$/i, "").trim();
  return ({ "NIT Tiruchirappalli": "NIT Trichy" })[s] || s;
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
  // Always render the initials avatar; overlay the photo when a filename is set.
  // If the file is missing or fails to load, onerror removes the img and the avatar shows through.
  const avatar = `<div class="s-avatar s-avatar-fallback">${initials}` +
    (imgName ? `<img class="s-avatar-img" src="assets/students/${imgName}" alt="${s.Name}" loading="lazy" onerror="this.remove()">` : "") +
    `</div>`;
  const prog = s.Category === "PhD" ? "" : s.Program;   // PhD already grouped by its section header
const inst = abbrevInstitute(s.Institute);
const meta = [prog, studentYearLabel(s)].filter(Boolean).join(" \u00b7 ");
const sub = [meta, inst ? `<span class="s-inst">${escAttr(inst)}</span>` : ""]
              .filter(Boolean).join(" \u00b7 ");
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

  const years = [];
  rows.forEach(r => {
    let bucket = years.find(g => g.year === r.Year);
    if (!bucket){ bucket = { year: r.Year, items: [] }; years.push(bucket); }
    bucket.items.push(r);
  });

  const shown = limit ? years.slice(0, limit) : years;
  const hidden = limit ? years.slice(limit) : [];

  function itemsHtml(list){
    return list.map(g => `
      <div class="xp-item">
        <span class="xp-dot"></span>
        <div class="xp-year">${g.year}</div>
        <div class="xp-card">
          ${g.items.map(it => `
            <div class="xp-entry">
              <div class="xp-role">${it.Title}</div>
              ${it.Institution ? `<div class="xp-org">${it.Institution}</div>` : ""}
              ${it.Note ? `<div class="xp-note">${it.Note}</div>` : ""}
            </div>`).join("")}
        </div>
      </div>`).join("");
  }

  container.innerHTML = `<div class="xp-timeline">${itemsHtml(shown)}</div>`;

  if (hidden.length){
    const btn = document.createElement("button");
    btn.className = "show-more-btn";
    btn.textContent = "Show more...";
    let expanded = false;
    btn.addEventListener("click", () => {
      expanded = !expanded;
      container.querySelector(".xp-timeline").innerHTML = itemsHtml(expanded ? years : shown);
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

  const lock = '<svg class="notes-lock" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><rect x="3" y="11" width="18" height="11" rx="2"/><path d="M7 11V7a5 5 0 0 1 10 0v4"/></svg>';

  // Aggregate one row per OFFERING into one entry per COURSE.
  const order = [], map = {};
  rows.forEach(r => {
    const key = String(r.Course || "").trim();
    if (!key) return;
    if (!map[key]){ map[key] = { course: key, levels: new Set(), offers: [], current: false, notes: [] }; order.push(key); }
    const c = map[key];
    const level = String(r.Level || "").trim().toUpperCase();
    if (level) c.levels.add(level);
    if (r.Status === "Current") c.current = true;
    c.offers.push(r);
    const url = String(r.Notes || "").trim();
    if (url && !c.notes.some(n => n.url === url)) c.notes.push({ url: url, level: level });
  });

  const levelTag = lv => {
    const cls = lv === "PG" ? "lvl-pg" : (lv === "UG" ? "lvl-ug" : "");
    return `<span class="course-level ${cls}">${lv}</span>`;
  };

  const card = (c) => {
    const lvls = ["UG", "PG"].filter(l => c.levels.has(l)).map(levelTag).join("")
      + [...c.levels].filter(l => l !== "UG" && l !== "PG").map(levelTag).join("");

    const byInst = {}, instOrder = [];
    c.offers.forEach(o => {
      const inst = o.Institution || "Other";
      if (!byInst[inst]){ byInst[inst] = []; instOrder.push(inst); }
      const bit = [String(o.Level || "").trim().toUpperCase(), String(o.Term || "").trim()].filter(Boolean).join(" · ");
      if (bit) byInst[inst].push(bit);
    });
    const offHtml = instOrder.map(inst => {
      const terms = byInst[inst].join(", ");
      return `<li><span class="off-inst">${inst}</span>${terms ? ` — <span class="off-terms">${terms}</span>` : ""}</li>`;
    }).join("");

    // One notes link by default; if the sheet supplies a second (distinct) link, show it too.
    let notesHtml = "";
    if (c.notes.length === 1){
      notesHtml = `<a class="course-notes" href="${c.notes[0].url}" target="_blank" rel="noopener" title="Sign-in required">${lock} Class notes &rarr;</a>`;
    } else if (c.notes.length > 1){
      const links = c.notes.map(n => {
        const label = n.level ? `${n.level} notes` : "Class notes";
        return `<a class="course-notes" href="${n.url}" target="_blank" rel="noopener" title="Sign-in required">${lock} ${label} &rarr;</a>`;
      }).join("");
      notesHtml = `<div class="course-notes-row">${links}</div>`;
    }

    return `<div class="course-card${c.current ? " course-current" : ""}">
      <!-- ${c.current ? '<span class="course-badge">Now teaching</span>' : ""} -->
      <div class="course-top">
        <span class="course-title">${c.course}</span>
        ${lvls}
      </div>
      <ul class="course-offerings">${offHtml}</ul>
      ${notesHtml}
    </div>`;
  };

  const courses = order.map(k => map[k]);
  const current = courses.filter(c => c.current);
  const past = courses.filter(c => !c.current);

  let html = "";
  if (current.length){
    html += `<div class="teach-sec"><h2 class="teach-h">Currently Teaching</h2><div class="course-grid">${current.map(card).join("")}</div></div>`;
  }
  if (past.length){
    html += `<div class="teach-sec"><h2 class="teach-h">Past Courses</h2><div class="course-grid">${past.map(card).join("")}</div></div>`;
  }
  container.innerHTML = html;
}

function initTeaching(){
  const container = document.getElementById("teaching-list");
  if (!container) return;
  withLoading(container, () => loadSheet("data/teaching.xlsx", "Teaching"), renderTeaching, "teaching");
}

/* ---------------- News ---------------- */
const NEWS_GROUPS = {
  publication: { types:["publication"], label:"Publication",
    svg:'<svg class="news-ico-svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8z"/><path d="M14 3v5h5"/><path d="M8.5 13h7"/><path d="M8.5 16.5h4.5"/></svg>' },
  editorial: { types:["editorial","call for papers"], label:"Editorial / Call for Papers",
    svg:'<svg class="news-ico-svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="m3 11 18-5v12L3 14z"/><path d="M11.6 16.8a3 3 0 1 1-5.8-1.6"/></svg>' },
  talk: { types:["invited talk","invited talk/lecture","talk","lecture"], label:"Invited Talk",
    svg:'<svg class="news-ico-svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M12 2a3 3 0 0 0-3 3v7a3 3 0 0 0 6 0V5a3 3 0 0 0-3-3z"/><path d="M19 10v2a7 7 0 0 1-14 0v-2"/><line x1="12" y1="19" x2="12" y2="22"/></svg>' },
  update: { types:["update","news update"], label:"News Update",
    svg:'<svg class="news-ico-svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M6 8a6 6 0 0 1 12 0c0 7 3 9 3 9H3s3-2 3-9"/><path d="M10.3 21a1.94 1.94 0 0 0 3.4 0"/></svg>' }
};
function newsGroupFor(type){
  const t = String(type||"").trim().toLowerCase();
  for (const key in NEWS_GROUPS){ if (NEWS_GROUPS[key].types.includes(t)) return key; }
  return "update";
}
function newsItemHtml(n){
  const headline = n.Link
    ? `<a href="${n.Link}" target="_blank" rel="noopener">${n.Headline}</a>`
    : n.Headline;
  const gk = newsGroupFor(n.Type);
  const g = NEWS_GROUPS[gk];
  const label = String(n.Type || g.label).trim();
  const detail = n.Detail ? ` ${n.Detail}` : "";
  return `<li><span class="news-tag nt-${gk}" role="img" aria-label="${label}" title="${label}">${g.svg}</span><span class="news-body"><i>${headline}</i>${detail}</span></li>`;
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

  // Home page "Recent Announcements": auto-scrolling marquee (bottom-to-top) of the most recent items.
  rows.sort((a, b) => (Number(b.Year) || 0) - (Number(a.Year) || 0));
  const recent = rows.slice(0, count);
  if (!recent.length){
    container.innerHTML = `<div class="no-results">No updates yet.</div>`;
    return;
  }
  const itemHtml = (n, dup) => {
    const headline = n.Link
      ? `<a href="${n.Link}" target="_blank" rel="noopener">${n.Headline}</a>`
      : `<span class="u-head">${n.Headline}</span>`;
    const gk = newsGroupFor(n.Type);
    const g = NEWS_GROUPS[gk];
    const label = String(n.Type || g.label).trim();
    const detail = n.Detail ? ` <span class="u-detail">${n.Detail}</span>` : "";
    const cls = dup ? "news-ticker-item news-ticker-item--dup" : "news-ticker-item";
    const extra = dup ? ' aria-hidden="true"' : ` role="img" aria-label="${label}"`;
    return `<li class="${cls}"${extra}><span class="news-tag nt-${gk}"${dup ? '' : ' role="img" aria-label="'+label+'"'} title="${label}">${g.svg}</span><div class="u-body">${headline}${detail}</div></li>`;
  };
  const items = recent.map(n => itemHtml(n, false)).join("");
  const itemsDup = recent.map(n => itemHtml(n, true)).join("");
  const dur = Math.max(20, recent.length * 5); // seconds per loop — slow, human-readable
  container.innerHTML = `
    <div class="news-ticker">
      <ul class="news-ticker-track" style="animation-duration:${dur}s">${items}${itemsDup}</ul>
    </div>
    <div class="page-links" style="margin-top:14px;"><a href="news.html">Show more...</a></div>
  `;
}

function initNews(){
  const full = document.getElementById("news-list");
  const preview = document.getElementById("news-preview");
  if (!full && !preview) return;
  const target = full || preview;
  withLoading(target, () => loadSheet("data/news.xlsx", "News"), (rows) => {
    if (full) renderNewsFull(rows.slice());
    if (preview) renderNewsPreview(rows.slice(), 5);
  }, "news");
}

/* ---------------- Metrics chart (citations / publications per year) ----------------
   Reads data/metrics.xlsx (Year, Citations, Publications) and draws a dual-axis
   line chart with Chart.js — Publications on the left axis, Citations on the right,
   matching the reference site's layout. Falls back to a plain table if Chart.js
   didn't load (e.g. offline / CDN blocked). */
let METRICS_ROWS = null, metricsFullChart = null;

function drawMetricsChart(canvas, rows, compact){
  const years = rows.map(r => r.Year);
  const citations = rows.map(r => Number(r.Citations) || 0);
  const publications = rows.map(r => Number(r.Publications) || 0);
  const last = years.length - 1;
  const MAROON = "#8c4a51", MAROON_SOFT = "rgba(138,31,43,0.30)", STEEL = "#1f5f7a";
  const f = compact ? 10 : 13;
  return new Chart(canvas.getContext("2d"), {
    type: "bar",
    data: {
      labels: years,
      datasets: [
        {
          type: "bar", label: "Publications", data: publications, yAxisID: "yPub",
          backgroundColor: publications.map((_, i) => i === last ? MAROON_SOFT : MAROON),
          borderColor: MAROON, borderWidth: publications.map((_, i) => i === last ? 1.5 : 0),
          borderRadius: compact ? 2 : 4, borderSkipped: false,
          maxBarThickness: compact ? 5 : 30,
          barPercentage: compact ? 0.55 : 0.9, categoryPercentage: compact ? 0.7 : 0.8, order: 2
        },
        {
          type: "line", label: "Citations", data: citations, yAxisID: "yCit",
          borderColor: STEEL, backgroundColor: STEEL, pointBackgroundColor: STEEL,
          pointRadius: publications.map((_, i) => compact ? (i === last ? 3 : 0) : (i === last ? 5 : 3)),
          pointBorderColor: "#fff", pointBorderWidth: 1,
          borderWidth: compact ? 1.8 : 2.4, tension: 0.35, order: 1,
          segment: { borderDash: c => c.p1DataIndex === last ? [5, 4] : undefined }
        }
      ]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      interaction: { mode: "index", intersect: false },
      plugins: {
        legend: { position: "bottom", labels: { boxWidth: compact ? 9 : 12, usePointStyle: true, padding: compact ? 8 : 16, font: { size: compact ? 10 : 13 } } },
        tooltip: { callbacks: { title: it => it[0].label + (it[0].dataIndex === last ? " (year-to-date)" : "") } }
      },
      scales: {
        x: { grid: { display: false }, ticks: { font: { size: f }, maxRotation: 0, autoSkip: true, autoSkipPadding: compact ? 8 : 12, color: "#6a6a6a" } },
        yPub: {
          type: "linear", position: "left", beginAtZero: true, suggestedMax: 14,
          title: { display: !compact, text: "Publications / year", color: MAROON, font: { weight: "600" } },
          grid: { color: "rgba(0,0,0,0.05)" }, ticks: { stepSize: 2, font: { size: f }, color: "#8a8a8a" }
        },
        yCit: {
          type: "linear", position: "right", beginAtZero: true,
          title: { display: !compact, text: "Citations / year", color: STEEL, font: { weight: "600" } },
          grid: { drawOnChartArea: false }, ticks: { font: { size: f }, color: "#8a8a8a", maxTicksLimit: compact ? 5 : 8 }
        }
      }
    }
  });
}

function setupMetricsModal(lastYear){
  if (document.getElementById("metrics-modal")) return;
  const m = document.createElement("div");
  m.id = "metrics-modal"; m.className = "metrics-modal"; m.hidden = true;
  m.innerHTML =
    '<div class="metrics-modal-backdrop"></div>' +
    '<div class="metrics-modal-panel" role="dialog" aria-modal="true" aria-label="Citations and publications by year">' +
      '<button class="metrics-modal-close" aria-label="Close">&times;</button>' +
      '<div class="metrics-modal-head"><h3>Citations &amp; Publications by Year</h3>' +
        '<span>Google Scholar · ' + lastYear + ' year-to-date</span></div>' +
      '<div class="metrics-modal-canvas"><canvas id="metrics-canvas-full"></canvas></div>' +
    '</div>';
  document.body.appendChild(m);
  m.querySelector(".metrics-modal-close").addEventListener("click", closeMetricsModal);
  m.querySelector(".metrics-modal-backdrop").addEventListener("click", closeMetricsModal);
  document.addEventListener("keydown", e => { if (e.key === "Escape") closeMetricsModal(); });
}

function openMetricsModal(){
  const m = document.getElementById("metrics-modal");
  if (!m || !METRICS_ROWS) return;
  m.hidden = false;
  document.body.style.overflow = "hidden";
  if (!metricsFullChart){
    metricsFullChart = drawMetricsChart(document.getElementById("metrics-canvas-full"), METRICS_ROWS, false);
  } else {
    metricsFullChart.resize();
  }
}

function closeMetricsModal(){
  const m = document.getElementById("metrics-modal");
  if (!m) return;
  m.hidden = true;
  document.body.style.overflow = "";
}

function renderMetricsChart(rows){
  const container = document.getElementById("metrics-chart");
  if (!container) return;
  rows.sort((a, b) => (Number(a.Year) || 0) - (Number(b.Year) || 0));
  METRICS_ROWS = rows;

  const years = rows.map(r => r.Year);
  const citations = rows.map(r => Number(r.Citations) || 0);
  const publications = rows.map(r => Number(r.Publications) || 0);

  if (typeof Chart === "undefined"){
    container.innerHTML = `
      <div class="no-results">Chart library unavailable — showing raw data from metrics.xlsx.</div>
      <table style="width:100%; border-collapse:collapse; margin-top:12px;">
        <tr><th style="text-align:left;">Year</th>${years.map(y => `<td style="text-align:center;">${y}</td>`).join("")}</tr>
        <tr><th style="text-align:left;">Publications</th>${publications.map(v => `<td style="text-align:center;">${v}</td>`).join("")}</tr>
        <tr><th style="text-align:left;">Citations</th>${citations.map(v => `<td style="text-align:center;">${v}</td>`).join("")}</tr>
      </table>`;
    return;
  }

  const lastYear = years[years.length - 1];
  container.innerHTML = `<canvas id="metrics-canvas"></canvas><span class="chart-expand" aria-hidden="true">&#10530; Enlarge</span>`;
  drawMetricsChart(document.getElementById("metrics-canvas"), rows, true);
  setupMetricsModal(lastYear);
  container.addEventListener("click", openMetricsModal);
  container.addEventListener("keydown", e => { if (e.key === "Enter" || e.key === " "){ e.preventDefault(); openMetricsModal(); } });
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
  initServices();
});
