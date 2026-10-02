/* ==========================================================================
   REELSIGHT v4.4 — Sub Divisi EXATA (AI/MT/SN) di Laporan & Planner + Multi-Divisi + Server CRUD + Excel Export + ARSIP SYSTEM
   ========================================================================== */

/* ==========================================================================
   1. STATE & KONSTANTA
   ========================================================================== */
const LS_DRAFTS = "reelsight_drafts_v4";
const LS_FORMDRAFT = "reelsight_formdraft_v4";
const LS_SHEETCACHE = "reelsight_sheetcache_v4";
const LS_PLANNER_DATA = "reelsight_planner_v2";
const LS_PLAN_DRAFT = "reelsight_plan_draft";

let drafts = loadJSON(LS_DRAFTS, []);
let sheetData = loadJSON(LS_SHEETCACHE, []);
let plannerData = loadJSON(LS_PLANNER_DATA, []);

let currentEditId = null;
let currentImageDataUrl = null;
let activeTab = "input";
let activeDivisiLaporan = null;
let activeDivisiPlanner = null;
let activeSubPlanner = null; // "AI" | "MT" | "SN" | "ALL" (planner, khusus EXATA)
let activeSubLaporan = null; // "AI" | "MT" | "SN" | "ALL" (hanya dipakai saat divisi = EXATA)
let plannerEditId = null;

const fieldsInsight = [
  "type", "divisi", "subdivisi", "week", "title", "link", "script", "isi_carousel", "posted",
  "downloaded", "caption", "views", "reach", "duration", "watchtime",
  "kunjungan", "mengikuti", "likes", "comments", "reposts", "shares", "saves"
];
const fieldsPlan = ["divisi", "subdivisi", "week", "title", "planDate", "format", "objective", "concept", "script"];

/* ==========================================================================
   2. UTILITAS DASAR
   ========================================================================== */
function loadJSON(key, fallback) {
  try {
    const raw = localStorage.getItem(key);
    return raw ? JSON.parse(raw) : fallback;
  } catch (e) {
    return fallback;
  }
}

function saveJSON(key, val) {
  try {
    localStorage.setItem(key, JSON.stringify(val));
  } catch (e) {
    /* storage penuh / diblokir, abaikan */
  }
}

function uid() {
  return "d_" + Date.now() + "_" + Math.random().toString(36).slice(2, 8);
}

function $(sel) {
  return document.querySelector(sel);
}

function $all(sel) {
  return Array.from(document.querySelectorAll(sel));
}

function escapeHtml(s) {
  const map = { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" };
  return String(s || "").replace(/[&<>"']/g, (c) => map[c]);
}

function fmtNum(n) {
  return (Number(n) || 0).toLocaleString("id-ID");
}

function fmtDate(d) {
  if (!d) return "-";
  try {
    return new Date(d + "T00:00:00").toLocaleDateString("id-ID", {
      day: "numeric", month: "short", year: "numeric"
    });
  } catch (e) {
    return d;
  }
}

function fmtPlanDate(d) {
  if (!d) return "-";
  try {
    return new Date(d + "T00:00:00").toLocaleDateString("id-ID", {
      weekday: "long",
      day: "numeric",
      month: "short",
      year: "numeric"
    });
  } catch (e) {
    return d;
  }
}

/* Label "EXATA · AI" untuk kartu */
function divisiLabel(d) {
  if (!d.divisi) return "";
  return d.divisi + (d.subdivisi ? " · " + d.subdivisi : "");
}

/* Urutan kartu: tanggal posting TERLAMA di awal (kiri/atas) -> TERBARU di akhir (kanan/bawah).
   Item tanpa tanggal ditaruh paling akhir; jika tanggal sama, urut berdasar waktu input. */
function compareByDateAsc(dateA, dateB, createdA, createdB) {
  const a = dateA || "";
  const b = dateB || "";
  if (a !== b) {
    if (!a) return 1;
    if (!b) return -1;
    return String(a).localeCompare(String(b));
  }
  return (createdA || 0) - (createdB || 0);
}

let toastTimer;
function toast(msg, type = "") {
  const el = $("#toast");
  if (!el) return;
  el.textContent = msg;
  el.className = "toast show " + type;
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => el.classList.remove("show"), 2600);
}

let flashSavedTimer;
function flashSaved() {
  const el = $("#autosaveStatus");
  if (!el) return;
  const dot = el.querySelector(".dot");

  if (dot) dot.style.background = "var(--amber)";
  el.lastChild.textContent = " Menyimpan...";

  clearTimeout(flashSavedTimer);
  flashSavedTimer = setTimeout(() => {
    if (dot) dot.style.background = "var(--teal)";
    el.lastChild.textContent = " Tersimpan";
  }, 500);
}

function isServerConfigured() {
  return typeof APPS_SCRIPT_URL !== "undefined" && APPS_SCRIPT_URL.includes("http");
}

/* ==========================================================================
   3. LIGHTBOX & EMBED VIDEO
   ========================================================================== */
const lightbox = $("#lightbox");
const lightboxImg = $("#lightboxImg");

function openLightbox(src) {
  if (!src || !lightboxImg || !lightbox) return;
  lightboxImg.src = src;
  lightbox.hidden = false;
}

$("#lightboxClose")?.addEventListener("click", () => {
  if (lightbox) lightbox.hidden = true;
});

lightbox?.addEventListener("click", (e) => {
  if (e.target.id === "lightbox") lightbox.hidden = true;
});

function buildEmbedUrl(url) {
  try {
    const u = new URL(url);
    if (u.hostname.includes("instagram.com")) {
      const path = u.pathname.endsWith("/") ? u.pathname : u.pathname + "/";
      return `https://www.instagram.com${path}embed`;
    }
  } catch (e) {
    /* URL tidak valid, pakai apa adanya */
  }
  return url;
}

function openEmbed(url) {
  if (!url) return;

  const eUrl = buildEmbedUrl(url);
  if ($("#embedNewTab")) $("#embedNewTab").href = url;
  if ($("#embedFrame")) $("#embedFrame").src = eUrl;
  if ($("#embedOverlay")) {
    $("#embedOverlay").hidden = false;
    document.body.style.overflow = "hidden";
  }
}

$("#embedCloseBtn")?.addEventListener("click", () => {
  if ($("#embedOverlay")) $("#embedOverlay").hidden = true;
  if ($("#embedFrame")) $("#embedFrame").src = "";
  document.body.style.overflow = "";
});

/* ==========================================================================
   4. NAVIGASI TAB & PEMILIHAN DIVISI
   ========================================================================== */
function switchTab(tabName) {
  activeTab = tabName;

  $all(".tab-btn").forEach((b) => b.classList.toggle("active", b.dataset.tab === tabName));
  $("#panel-input")?.classList.toggle("hidden", activeTab !== "input");
  $("#panel-lihat")?.classList.toggle("hidden", activeTab !== "lihat");
  $("#panel-planner")?.classList.toggle("hidden", activeTab !== "planner");
  $("#panel-arsip")?.classList.toggle("hidden", activeTab !== "arsip"); // <-- LOGIKA TAB ARSIP
  $("#appFooter")?.classList.toggle("hide", activeTab !== "input");

  if (activeTab === "lihat") {
    if (!activeDivisiLaporan) {
      if ($("#laporanDivisiMenu")) $("#laporanDivisiMenu").hidden = false;
      if ($("#laporanSubMenu")) $("#laporanSubMenu").hidden = true;
      if ($("#laporanContent")) $("#laporanContent").hidden = true;
    }
    if (sheetData.length === 0) fetchSheetData();
  }

  if (activeTab === "planner") {
    if (!activeDivisiPlanner) {
      if ($("#plannerDivisiMenu")) $("#plannerDivisiMenu").hidden = false;
      if ($("#plannerSubMenu")) $("#plannerSubMenu").hidden = true;
      if ($("#plannerContent")) $("#plannerContent").hidden = true;
    }
    if (plannerData.length === 0) fetchPlannerData();
  }

  if (activeTab === "arsip") {
    if (sheetData.length === 0) fetchSheetData();
    if (plannerData.length === 0) fetchPlannerData();
    if (typeof renderArsipGrid === "function") renderArsipGrid();
  }
}

$all(".tab-btn").forEach((btn) => {
  btn.addEventListener("click", () => switchTab(btn.dataset.tab));
});

function laporanLabel() {
  const sub = activeDivisiLaporan === "EXATA" && activeSubLaporan && activeSubLaporan !== "ALL"
    ? " · " + activeSubLaporan : "";
  return activeDivisiLaporan + sub;
}

function showLaporanContent() {
  if ($("#laporanDivisiTitle")) $("#laporanDivisiTitle").textContent = "Laporan: " + laporanLabel();
  if ($("#laporanDivisiMenu")) $("#laporanDivisiMenu").hidden = true;
  if ($("#laporanSubMenu")) $("#laporanSubMenu").hidden = true;
  if ($("#laporanContent")) $("#laporanContent").hidden = false;
  renderSheetGrid();
}

function resetLaporanMenu() {
  activeDivisiLaporan = null;
  activeSubLaporan = null;
  if ($("#laporanDivisiMenu")) $("#laporanDivisiMenu").hidden = false;
  if ($("#laporanSubMenu")) $("#laporanSubMenu").hidden = true;
  if ($("#laporanContent")) $("#laporanContent").hidden = true;
}

$all("#laporanDivisiMenu .div-card").forEach((btn) => {
  btn.addEventListener("click", () => {
    activeDivisiLaporan = btn.dataset.div;
    activeSubLaporan = null;
    if (activeDivisiLaporan === "EXATA") {
      // EXATA: pilih dulu bagian AI / MT / SN
      if ($("#laporanDivisiMenu")) $("#laporanDivisiMenu").hidden = true;
      if ($("#laporanSubMenu")) $("#laporanSubMenu").hidden = false;
      if ($("#laporanContent")) $("#laporanContent").hidden = true;
      return;
    }
    showLaporanContent();
  });
});

$all("#laporanSubMenu .div-card").forEach((btn) => {
  btn.addEventListener("click", () => {
    activeSubLaporan = btn.dataset.sub;
    showLaporanContent();
  });
});

$("#backLaporanSubBtn")?.addEventListener("click", resetLaporanMenu);

$("#backLaporanBtn")?.addEventListener("click", () => {
  if (activeDivisiLaporan === "EXATA") {
    // kembali ke pilihan AI / MT / SN
    activeSubLaporan = null;
    if ($("#laporanSubMenu")) $("#laporanSubMenu").hidden = false;
    if ($("#laporanContent")) $("#laporanContent").hidden = true;
  } else {
    resetLaporanMenu();
  }
});

function plannerLabel() {
  const sub = activeDivisiPlanner === "EXATA" && activeSubPlanner && activeSubPlanner !== "ALL"
    ? " · " + activeSubPlanner : "";
  return activeDivisiPlanner + sub;
}

function showPlannerContent() {
  if ($("#plannerDivisiTitle")) $("#plannerDivisiTitle").textContent = "Planner: " + plannerLabel();
  if ($("#plannerDivisiMenu")) $("#plannerDivisiMenu").hidden = true;
  if ($("#plannerSubMenu")) $("#plannerSubMenu").hidden = true;
  if ($("#plannerContent")) $("#plannerContent").hidden = false;
  renderPlannerGrid();
}

function resetPlannerMenu() {
  activeDivisiPlanner = null;
  activeSubPlanner = null;
  if ($("#plannerDivisiMenu")) $("#plannerDivisiMenu").hidden = false;
  if ($("#plannerSubMenu")) $("#plannerSubMenu").hidden = true;
  if ($("#plannerContent")) $("#plannerContent").hidden = true;
}

$all("#plannerDivisiMenu .div-card").forEach((btn) => {
  btn.addEventListener("click", () => {
    activeDivisiPlanner = btn.dataset.div;
    activeSubPlanner = null;
    if (activeDivisiPlanner === "EXATA") {
      // EXATA: pilih dulu bagian AI / MT / SN
      if ($("#plannerDivisiMenu")) $("#plannerDivisiMenu").hidden = true;
      if ($("#plannerSubMenu")) $("#plannerSubMenu").hidden = false;
      if ($("#plannerContent")) $("#plannerContent").hidden = true;
      return;
    }
    showPlannerContent();
  });
});

$all("#plannerSubMenu .div-card").forEach((btn) => {
  btn.addEventListener("click", () => {
    activeSubPlanner = btn.dataset.sub;
    showPlannerContent();
  });
});

$("#backPlannerSubBtn")?.addEventListener("click", resetPlannerMenu);

$("#backPlannerBtn")?.addEventListener("click", () => {
  if (activeDivisiPlanner === "EXATA") {
    activeSubPlanner = null;
    if ($("#plannerSubMenu")) $("#plannerSubMenu").hidden = false;
    if ($("#plannerContent")) $("#plannerContent").hidden = true;
  } else {
    resetPlannerMenu();
  }
});

/* ==========================================================================
   5. KARTU INSIGHT & DRAFT (RENDER)
   ========================================================================== */
function renderDrafts() {
  const grid = $("#draftGrid");
  if (!grid) return;

  grid.querySelectorAll(".insight-card").forEach((n) => n.remove());

  const sorted = [...drafts].sort((a, b) => (b._createdAt || 0) - (a._createdAt || 0));
  sorted.forEach((d, i) => {
    const card = buildCard(d, true, "draft");
    card.style.animationDelay = i * 0.04 + "s";
    grid.appendChild(card);
  });

  if ($("#pendingBadge")) {
    $("#pendingBadge").hidden = drafts.length === 0;
    $("#pendingBadge").textContent = drafts.length;
  }
  if ($("#footerCount")) $("#footerCount").textContent = drafts.length;
  if ($("#sendAllBtn")) $("#sendAllBtn").disabled = drafts.length === 0;
}

function cardImageHtml(d, editable) {
  if (d.image) {
    return `<img class="thumb" src="${d.image}" style="cursor:zoom-in;">`;
  }
  if (d.link && !editable) {
    return `<div class="thumb-fallback clickable-embed" data-link="${escapeHtml(d.link)}"
      style="cursor:pointer; background:var(--primary-light); color:var(--primary);" title="Putar Konten">
      <span style="font-size:28px;">▶️</span>
      <span style="font-size:12px; font-weight:600; margin-top:4px;">Putar Video</span>
    </div>`;
  }
  return `<div class="thumb-fallback">🎬</div>`;
}

function statRow(items) {
  return `<div class="card-stats">${items.map(([label, val]) =>
    `<div class="card-stat"><b>${val}</b><span>${label}</span></div>`
  ).join("")}</div>`;
}

function detailRow(items) {
  return `<div class="card-details-list">${items.map(([label, val]) =>
    `<div class="card-detail-item"><span>${label}</span><b>${val}</b></div>`
  ).join("")}</div>`;
}

function cardDetailsHtml(d, editable) {
  if (editable) {
    return statRow([
      ["Tayangan", d.views || "0"],
      ["Suka", fmtNum(d.likes)],
      ["Disimpan", fmtNum(d.saves)]
    ]);
  }

  if (d.type === "carousel") {
    const scriptHtml = d.isi_carousel ? `<div class="card-script">${escapeHtml(d.isi_carousel)}</div>` : "";
    const details = detailRow([
      ["Tayangan", d.views || "-"],
      ["Pemirsa", d.reach || "-"],
      ["Kunjungan Profil", d.kunjungan || "-"],
      ["Mengikuti", d.mengikuti || "-"],
      ["Suka", fmtNum(d.likes)],
      ["Komentar", fmtNum(d.comments)],
      ["Repost", fmtNum(d.reposts)],
      ["Share", fmtNum(d.shares)],
      ["Save", fmtNum(d.saves)]
    ]);
    return `${scriptHtml}${details}<button class="card-copy" data-act="copy">📋 Salin Prompt AI</button>`;
  }

  const scriptHtml = d.script ? `<div class="card-script">${escapeHtml(d.script)}</div>` : "";
  const details = detailRow([
    ["Durasi", d.duration ? d.duration + " dtk" : "-"],
    ["Tayangan", d.views || "-"],
    ["Pemirsa", d.reach || "-"],
    ["Waktu Tonton", d.watchtime ? d.watchtime + " dtk" : "-"],
    ["Suka", fmtNum(d.likes)],
    ["Komentar", fmtNum(d.comments)],
    ["Repost", fmtNum(d.reposts)],
    ["Disimpan", fmtNum(d.saves)]
  ]);
  return `${scriptHtml}${details}<button class="card-copy" data-act="copy">📋 Salin Prompt AI</button>`;
}

function cardActionsHtml() {
  return `<div class="card-actions">
    <button class="card-edit" data-act="edit">✎ Edit</button>
    <button class="card-delete" data-act="delete">🗑 Hapus</button>
  </div>`;
}

function buildCopyPrompt(d) {
  if (d.type === "carousel") {
    return `Konten carousel tanggal ${fmtDate(d.posted)}, judul ${d.title}, isi: ${d.isi_carousel}. `
      + `Insight (unduh ${fmtDate(d.downloaded)}):\n`
      + `Tayangan:${d.views}, Pemirsa:${d.reach}, Kunjungan:${d.kunjungan}, Follow:${d.mengikuti}, `
      + `Suka:${fmtNum(d.likes)}, Komentar:${fmtNum(d.comments)}, Repost:${fmtNum(d.reposts)}, `
      + `Share:${fmtNum(d.shares)}, Save:${fmtNum(d.saves)}.\nAnalisa mendalam dan beri penilaian konten ini.`;
  }
  return `Konten reels tanggal ${fmtDate(d.posted)}, judul ${d.title}, script: ${d.script}, durasi ${d.duration}s. `
    + `Insight (unduh ${fmtDate(d.downloaded)}):\n`
    + `Tayangan:${d.views}, Pemirsa:${d.reach}, Waktu Tonton:${d.watchtime}s, `
    + `Suka:${fmtNum(d.likes)}, Komentar:${fmtNum(d.comments)}, Repost:${fmtNum(d.reposts)}, `
    + `Share:${fmtNum(d.shares)}, Save:${fmtNum(d.saves)}.\nAnalisa mendalam dan beri penilaian konten ini.`;
}

async function deleteSheetCard(card, d) {
  const btn = card.querySelector('[data-act="delete"]');
  btn.textContent = "Menghapus...";
  btn.disabled = true;

  try {
    const res = await fetch(APPS_SCRIPT_URL, {
      method: "POST",
      body: JSON.stringify({ action: "delete_insight", id: d.id, divisi: d.divisi, type: d.type })
    });
    const json = await res.json();
    if (!json.ok) throw new Error("Gagal server");

    toast("Data terhapus dari server", "success");
    sheetData = sheetData.filter((x) => x.id !== d.id);
    saveJSON(LS_SHEETCACHE, sheetData);
    renderSheetGrid();
    if (typeof renderArsipGrid === "function") renderArsipGrid();
  } catch (e) {
    toast("Gagal menghapus", "error");
    btn.textContent = "🗑 Hapus";
    btn.disabled = false;
  }
}

function deleteDraftCard(d) {
  drafts = drafts.filter((x) => x.id !== d.id);
  saveJSON(LS_DRAFTS, drafts);
  renderDrafts();
}

function bindCardEvents(card, d, editable, context) {
  if (d.image) {
    card.querySelector(".thumb")?.addEventListener("click", () => openLightbox(d.image));
  }

  card.querySelector('[data-act="edit"]')?.addEventListener("click", () => {
    openModal(d.id, d.type, context);
  });

  card.querySelector('[data-act="delete"]')?.addEventListener("click", async () => {
    if (!confirm(`Hapus insight "${d.title}" ini?`)) return;
    if (context === "draft") {
      deleteDraftCard(d);
    } else if (context === "sheet") {
      await deleteSheetCard(card, d);
    }
  });

  card.querySelector('[data-act="copy"]')?.addEventListener("click", () => {
    navigator.clipboard.writeText(buildCopyPrompt(d)).then(() => toast("Prompt tersalin", "success"));
  });

  if (!editable) {
    card.querySelector(".clickable-embed")?.addEventListener("click", () => openEmbed(d.link));
  }
}

function buildCard(d, editable, context) {
  const card = document.createElement("div");
  card.className = `insight-card type-${d.type || "reels"}`;

  const imgHtml = cardImageHtml(d, editable);
  const detailsHtml = cardDetailsHtml(d, editable);
  const actionHtml = cardActionsHtml();
  const formatLabel = d.type === "carousel" ? "<span style='color:var(--teal)'>• Carousel</span>" : "• Reels";
  const divLabel = d.divisi
    ? `<span style="background:var(--gray-light); padding:2px 6px; border-radius:4px; margin-right:4px;">${escapeHtml(divisiLabel(d))}</span>`
    : "";

  card.innerHTML = `${imgHtml}<div class="card-body">
      <div class="card-date">${divLabel}${fmtDate(d.posted)} ${formatLabel}</div>
      <div class="card-title">${escapeHtml(d.title || "(Tanpa judul)")}</div>
      ${detailsHtml}
      ${actionHtml}
    </div>`;

  bindCardEvents(card, d, editable, context);
  return card;
}

/* ==========================================================================
   6. MODAL FORMULIR (INPUT / EDIT INSIGHT)
   ========================================================================== */
const overlay = $("#modalOverlay");

function fillFormFields(prefix, fields, source) {
  fields.forEach((f) => {
    const el = $(`#${prefix}_${f}`);
    if (el) el.value = source[f] ?? "";
  });
}

function readFormFields(prefix, fields) {
  const out = {};
  fields.forEach((f) => {
    const el = $(`#${prefix}_${f}`);
    if (el) out[f] = el.value.trim();
  });
  return out;
}

function syncSubDivisiField() {
  const isExata = $("#f_divisi")?.value === "EXATA";
  if ($("#f_subdivisi_wrap")) $("#f_subdivisi_wrap").hidden = !isExata;
}

$("#f_divisi")?.addEventListener("change", () => {
  syncSubDivisiField();
  if ($("#f_divisi").value !== "EXATA" && $("#f_subdivisi")) $("#f_subdivisi").value = "";
  autosaveInsightDraft();
});

function toggleTypeFields(activeType) {
  $all(".type-reels").forEach((el) => (el.hidden = activeType !== "reels"));
  $all(".type-carousel").forEach((el) => (el.hidden = activeType !== "carousel"));
}

function loadEditingRecord(editId, context) {
  const source = context === "draft" ? drafts.find((x) => x.id === editId) : sheetData.find((x) => x.id === editId);
  if (!source) return null;

  const saved = loadJSON(LS_FORMDRAFT, null);
  const useDraft = !!(saved && saved._editId === editId && saved._context === context);
  const d = useDraft ? { ...source, ...saved } : source;

  if ($("#modalTitle")) {
    $("#modalTitle").textContent = context === "sheet" ? "Update Data Laporan" : "Edit Insight Draft";
  }
  fillFormFields("f", fieldsInsight, d);
  setImage(d.image || null, false);
  if ($("#f_divisi")) $("#f_divisi").disabled = context === "sheet";
  if (useDraft) toast("Melanjutkan perubahan yang belum tersimpan", "success");

  return d.type || source.type || "reels";
}

function loadNewRecordForm(forceType) {
  if ($("#f_divisi")) $("#f_divisi").disabled = false;
  if ($("#f_week")) $("#f_week").value = "Minggu 1";

  const saved = loadJSON(LS_FORMDRAFT, null);
  if (saved && !saved._editId && saved._context === "draft" && saved.type === forceType) {
    fillFormFields("f", fieldsInsight, saved);
    setImage(saved.image || null, false);
    toast("Melanjutkan draft yang belum tersimpan", "success");
  } else {
    fillFormFields("f", fieldsInsight, {});
    if ($("#f_type")) $("#f_type").value = forceType;
    if ($("#f_divisi")) $("#f_divisi").value = "EXATA";
    setImage(null, false);
  }

  if ($("#modalTitle")) {
    $("#modalTitle").textContent = forceType === "carousel" ? "Tambah Insight Carousel" : "Tambah Insight Reels";
  }
}

function openModal(editId = null, forceType = "reels", context = "draft") {
  currentEditId = editId;
  if ($("#f_edit_context")) $("#f_edit_context").value = context;
  if (overlay) {
    overlay.hidden = false;
    document.body.style.overflow = "hidden";
  }

  let activeType = forceType;
  if (editId) {
    activeType = loadEditingRecord(editId, context) || forceType;
  } else {
    loadNewRecordForm(forceType);
  }

  toggleTypeFields(activeType);
  syncSubDivisiField();
}

function closeModal() {
  if (overlay) overlay.hidden = true;
  document.body.style.overflow = "";
  currentEditId = null;
  renderResumeBanners();
}

$("#openAddReelsModal")?.addEventListener("click", () => openModal(null, "reels", "draft"));
$("#openAddCarouselModal")?.addEventListener("click", () => openModal(null, "carousel", "draft"));
$("#modalCloseBtn")?.addEventListener("click", closeModal);
$("#cancelModalBtn")?.addEventListener("click", closeModal);

function autosaveInsightDraft() {
  const obj = {
    _editId: currentEditId,
    _context: $("#f_edit_context")?.value || "draft",
    image: currentImageDataUrl,
    ...readFormFields("f", fieldsInsight)
  };
  saveJSON(LS_FORMDRAFT, obj);
  flashSaved();
}

fieldsInsight.forEach((f) => {
  $(`#f_${f}`)?.addEventListener("input", autosaveInsightDraft);
});

/* ---- Upload gambar ---- */
const fileInput = $("#fileInput");

$("#dropzone")?.addEventListener("click", () => fileInput?.click());

fileInput?.addEventListener("change", (e) => {
  const file = e.target.files && e.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = () => setImage(reader.result, true);
  reader.readAsDataURL(file);
});

$("#maximizeBtn")?.addEventListener("click", (e) => {
  e.stopPropagation();
  openLightbox(currentImageDataUrl);
});
$("#replaceImgBtn")?.addEventListener("click", () => fileInput?.click());
$("#removeImgBtn")?.addEventListener("click", () => setImage(null, true));

function setImage(dataUrl, trig) {
  currentImageDataUrl = dataUrl;
  if ($("#previewImg")) $("#previewImg").src = dataUrl || "";
  if ($("#dropzoneEmpty")) $("#dropzoneEmpty").hidden = !!dataUrl;
  if ($("#dropzonePreview")) $("#dropzonePreview").hidden = !dataUrl;
  if ($("#dropzoneActions")) $("#dropzoneActions").hidden = !dataUrl;
  if ($("#scanBtn")) $("#scanBtn").disabled = !dataUrl;
  if (trig) flashSaved();
}

/* ==========================================================================
   7. OCR SCAN LOGIC
   ========================================================================== */
$("#scanBtn")?.addEventListener("click", async () => {
  if (!currentImageDataUrl) return;

  const pWrap = $("#scanProgress");
  const bar = $("#scanProgressBar");
  const btn = $("#scanBtn");

  btn.disabled = true;
  if (pWrap) pWrap.hidden = false;
  if (bar) bar.style.width = "4%";

  try {
    const res = await Tesseract.recognize(currentImageDataUrl, "ind+eng", {
      logger: (m) => {
        if (m.status === "recognizing text" && bar) {
          bar.style.width = Math.max(6, Math.round(m.progress * 100)) + "%";
        }
      }
    });
    if (bar) bar.style.width = "100%";
    applyParsedData(parseInsightText(res.data.text));
    toast("Pemindaian selesai", "success");
  } catch (err) {
    toast("Pemindaian gagal", "error");
  } finally {
    setTimeout(() => {
      if (pWrap) pWrap.hidden = true;
      btn.disabled = false;
    }, 500);
  }
});

function extractTopMetrics(lines) {
  const pattern = /^([\d.,]+[kmb]?)\s+([\d.,]+[kmb]?)\s+([\d.,]+[kmb]?)\s+([\d.,]+[kmb]?)\s+([\d.,]+[kmb]?)$/i;
  for (const line of lines) {
    const match = line.match(pattern);
    if (match) {
      return { likes: match[1], comments: match[2], reposts: match[3], shares: match[4], saves: match[5] };
    }
  }
  return null;
}

function extractCaption(lines) {
  const idx = lines.findIndex((l) => /keterangan/i.test(l));
  if (idx === -1 || !lines[idx + 1]) return null;

  let cap = lines[idx].replace(/keterangan\s*[:\-]?/i, "").trim();
  if (!cap) cap = lines[idx + 1];
  return cap;
}

function findDateNear(lines, labelRe, dateRe) {
  const months = {
    jan: 1, feb: 2, mar: 3, apr: 4, mei: 5, may: 5, jun: 6,
    jul: 7, agu: 8, aug: 8, sep: 9, okt: 10, oct: 10, nov: 11, des: 12, dec: 12
  };

  for (let i = 0; i < lines.length; i++) {
    if (!labelRe.test(lines[i])) continue;

    const str = `${lines[i]} ${lines[i + 1] || ""} ${lines[i + 2] || ""}`;
    const m = str.match(dateRe);
    if (!m) continue;

    let dMatch = m[0].match(/(\d{1,2})\s+([A-Za-z]+)\.?\s+(\d{4})/);
    if (dMatch) {
      const mo = months[dMatch[2].toLowerCase().slice(0, 3)];
      if (mo) return `${dMatch[3]}-${String(mo).padStart(2, "0")}-${String(dMatch[1]).padStart(2, "0")}`;
    }

    dMatch = m[0].match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/);
    if (dMatch) {
      const yr = dMatch[3].length === 2 ? "20" + dMatch[3] : dMatch[3];
      return `${yr}-${String(dMatch[2]).padStart(2, "0")}-${String(dMatch[1]).padStart(2, "0")}`;
    }

    return m[0];
  }
  return "";
}

function buildKnownLabels(foundTopMetrics) {
  const labels = [
    { key: "views", regex: /tayangan/i },
    { key: "reach", regex: /pemirsa/i },
    { key: "watchtime", regex: /waktu (tonton|menonton)|rata-?rata/i },
    { key: "kunjungan", regex: /kunjungan/i },
    { key: "mengikuti", regex: /mengikuti/i }
  ];
  if (!foundTopMetrics) {
    labels.push(
      { key: "likes", regex: /^suka\b|\bsuka$/i },
      { key: "comments", regex: /komentar/i },
      { key: "reposts", regex: /posting ulang/i },
      { key: "shares", regex: /dibagikan|bagikan/i },
      { key: "saves", regex: /disimpan|simpan/i }
    );
  }
  return labels;
}

function extractLabeledMetrics(lines, knownLabels, out) {
  const numRegex = /([\d]+[.,]?[\d]*)\s*(rb|jt|k|m)?/gi;

  for (let i = 0; i < lines.length; i++) {
    const foundLabels = knownLabels
      .map((lbl) => {
        const match = lines[i].match(lbl.regex);
        return match ? { key: lbl.key, index: match.index, regex: lbl.regex } : null;
      })
      .filter(Boolean)
      .sort((a, b) => a.index - b.index);

    if (foundLabels.length === 0) continue;

    const inlineNums = [...lines[i].matchAll(numRegex)].map((m) => m[0]);
    const nextLineNums = lines[i + 1] ? [...lines[i + 1].matchAll(numRegex)].map((m) => m[0]) : [];

    if (nextLineNums.length >= foundLabels.length) {
      foundLabels.forEach((lbl, idx) => { if (!out[lbl.key]) out[lbl.key] = nextLineNums[idx]; });
    } else if (inlineNums.length >= foundLabels.length) {
      foundLabels.forEach((lbl, idx) => { if (!out[lbl.key]) out[lbl.key] = inlineNums[idx]; });
    } else {
      foundLabels.forEach((lbl) => {
        if (out[lbl.key]) return;
        const after = lines[i].substring(lbl.index).replace(lbl.regex, "");
        let m = after.match(/([\d]+[.,]?[\d]*)\s*(rb|jt|k|m)?/i);
        if (m) {
          out[lbl.key] = m[0];
        } else if (lines[i + 1]) {
          m = lines[i + 1].match(/([\d]+[.,]?[\d]*)\s*(rb|jt|k|m)?/i);
          if (m) out[lbl.key] = m[0];
        }
      });
    }
  }
}

function parseInsightText(raw) {
  const lines = raw.replace(/\r/g, "").split("\n").map((l) => l.trim()).filter(Boolean);
  const out = {};

  const topMetrics = extractTopMetrics(lines);
  if (topMetrics) Object.assign(out, topMetrics);

  const datePattern = /(\d{1,2}\s+[a-zA-Z]+\.?\s+\d{4}|\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})/;
  out.posted = findDateNear(lines, /post(ing)?|diposting/i, datePattern);
  out.downloaded = findDateNear(lines, /diunduh|di ?unduh|unduh/i, datePattern);

  const caption = extractCaption(lines);
  if (caption) out.caption = caption;

  extractLabeledMetrics(lines, buildKnownLabels(!!topMetrics), out);

  return out;
}

function applyParsedData(data) {
  Object.keys(data).forEach((k) => {
    const el = $("#f_" + k);
    if (!el) return;

    let val = data[k];
    if (val === "" || val == null) return;
    if (el.type === "number") val = String(val).replace(/[^\d]/g, "");

    if (val !== "") {
      el.value = val;
      el.classList.add("filled");
      setTimeout(() => el.classList.remove("filled"), 1200);
    }
  });

  autosaveInsightDraft();
}

/* ==========================================================================
   8. SIMPAN DATA & KIRIM KE SERVER
   ========================================================================== */
async function updateSheetInsight(payload) {
  const btn = $("#saveCardBtn");
  btn.disabled = true;
  btn.textContent = "Updating...";

  try {
    const res = await fetch(APPS_SCRIPT_URL, {
      method: "POST",
      body: JSON.stringify({ action: "update_insight", data: payload })
    });
    const json = await res.json();
    if (!json.ok) throw new Error("Gagal Update");

    toast("Data Server Diperbarui", "success");
    const idx = sheetData.findIndex((x) => x.id === currentEditId);
    if (idx > -1) sheetData[idx] = { ...sheetData[idx], ...payload };
    saveJSON(LS_SHEETCACHE, sheetData);
    localStorage.removeItem(LS_FORMDRAFT);
    renderSheetGrid();
    if (typeof renderArsipGrid === "function") renderArsipGrid();
    closeModal();
  } catch (e) {
    toast("Gagal memperbarui server", "error");
  } finally {
    btn.disabled = false;
    btn.textContent = "Simpan";
  }
}

function saveInsightDraft(payload) {
  if (currentEditId) {
    const idx = drafts.findIndex((x) => x.id === currentEditId);
    if (idx !== -1) drafts[idx] = { ...drafts[idx], ...payload };
    toast("Draft diperbarui", "success");
  } else {
    payload.id = uid();
    payload._createdAt = Date.now();
    drafts.push(payload);
    toast("Insight ditambah ke draft", "success");
  }

  saveJSON(LS_DRAFTS, drafts);
  renderDrafts();
  localStorage.removeItem(LS_FORMDRAFT);
  closeModal();
}

$("#saveCardBtn")?.addEventListener("click", async () => {
  const titleOk = $("#f_title")?.value.trim();
  const divisiOk = $("#f_divisi")?.value;
  if (!titleOk || !divisiOk) {
    toast("Judul dan Divisi wajib diisi", "error");
    return;
  }

  const payload = { image: currentImageDataUrl, ...readFormFields("f", fieldsInsight) };
  const ctx = $("#f_edit_context")?.value || "draft";

  // Sub divisi hanya untuk EXATA (wajib untuk input baru; data lama boleh dikosongkan saat edit)
  if (payload.divisi !== "EXATA") {
    payload.subdivisi = "";
  } else if (!payload.subdivisi && ctx !== "sheet") {
    toast("Pilih bagian EXATA dulu (AI / MT / SN)", "error");
    return;
  }

  if (ctx === "sheet" && currentEditId) {
    payload.id = currentEditId;
    await updateSheetInsight(payload);
  } else {
    saveInsightDraft(payload);
  }
});

$("#sendAllBtn")?.addEventListener("click", async () => {
  if (drafts.length === 0) return;

  const btn = $("#sendAllBtn");
  btn.disabled = true;
  btn.textContent = "Mengirim...";

  try {
    const res = await fetch(APPS_SCRIPT_URL, {
      method: "POST",
      body: JSON.stringify({ action: "append", rows: drafts })
    });
    const json = await res.json();
    if (!json.ok) throw new Error("Gagal backend");

    toast(`${drafts.length} insight terkirim`, "success");
    drafts = [];
    saveJSON(LS_DRAFTS, drafts);
    renderDrafts();
    fetchSheetData();
  } catch (e) {
    toast("Gagal mengirim", "error");
  } finally {
    btn.disabled = drafts.length === 0;
    btn.innerHTML = "Kirim ke Spreadsheet";
  }
});

async function fetchSheetData() {
  if (!isServerConfigured()) return;
  if ($("#topLoadingBar")) $("#topLoadingBar").hidden = false;

  try {
    const res = await fetch(APPS_SCRIPT_URL + "?action=list");
    const json = await res.json();
    sheetData = json.rows || [];
    saveJSON(LS_SHEETCACHE, sheetData);
  } catch (e) {
    /* biarkan cache lama tetap dipakai */
  } finally {
    if ($("#topLoadingBar")) $("#topLoadingBar").hidden = true;
    renderSheetGrid();
    if (activeTab === "arsip" && typeof renderArsipGrid === "function") renderArsipGrid();
  }
}

function renderSheetGrid() {
  const grid = $("#sheetGrid");
  if (!grid) return;
  grid.innerHTML = "";
  if (!activeDivisiLaporan) return;

  const kw = $("#searchInput")?.value.toLowerCase();
  const from = $("#dateFrom")?.value;
  const to = $("#dateTo")?.value;
  const typeVal = $("#typeFilter")?.value;

  const list = sheetData.filter((d) => {
    if (d.divisi !== activeDivisiLaporan) return false;
    if (d.isArchived) return false;
    if (activeDivisiLaporan === "EXATA" && activeSubLaporan && activeSubLaporan !== "ALL"
        && d.subdivisi !== activeSubLaporan) return false;
    if (kw && !(d.title || "").toLowerCase().includes(kw)) return false;
    if (from && d.posted < from) return false;
    if (to && d.posted > to) return false;
    if (typeVal && d.type !== typeVal) return false;
    return true;
  });

  if ($("#sheetEmptyHint")) $("#sheetEmptyHint").hidden = list.length > 0;

  const weeks = ["Minggu 1", "Minggu 2", "Minggu 3", "Minggu 4"];
  weeks.forEach((w) => {
    const items = list
      .filter((item) => (item.week || "Minggu 1") === w)
      .sort((a, b) => compareByDateAsc(a.posted, b.posted, a._createdAt, b._createdAt));

    if (items.length === 0) return;

    const accordion = document.createElement("details");
    accordion.className = "week-accordion";
    accordion.open = true;
    accordion.innerHTML = `
      <summary class="week-summary">
        <div style="display:flex; align-items:center; gap:12px;">
          <input type="checkbox" class="archive-checkbox" data-week="${w}" title="Centang untuk arsip" style="transform:scale(1.2); cursor:pointer;">
          <span>📅 ${w}</span>
        </div>
        <span class="week-count-badge">${items.length} Insight</span>
      </summary>
      <div class="week-content"></div>
    `;

    // Cegah accordion menutup saat mencentang checkbox
    const checkbox = accordion.querySelector('.archive-checkbox');
    checkbox.addEventListener('click', (e) => e.stopPropagation());

    const contentGrid = accordion.querySelector(".week-content");
    items.forEach((d) => contentGrid.appendChild(buildCard(d, false, "sheet")));
    grid.appendChild(accordion);
  });
}

["searchInput", "dateFrom", "dateTo", "sortSelect", "typeFilter"].forEach((id) => {
  $("#" + id)?.addEventListener("input", renderSheetGrid);
  $("#" + id)?.addEventListener("change", renderSheetGrid);
});
$("#refreshSheetBtn")?.addEventListener("click", fetchSheetData);

/* ==========================================================================
   9. EKSPOR EXCEL (SheetJS)
   ========================================================================== */
$("#openExportModalBtn")?.addEventListener("click", () => {
  // Hanya mengekspor data yang BUKAN arsip
  const unarchivedData = sheetData.filter(d => !d.isArchived);
  if (unarchivedData.length === 0) {
    toast("Tidak ada laporan aktif (data kosong/sudah diarsipkan semua)", "error");
    return;
  }
  if ($("#exportModalOverlay")) $("#exportModalOverlay").hidden = false;
});

$("#closeExportBtn")?.addEventListener("click", () => {
  if ($("#exportModalOverlay")) $("#exportModalOverlay").hidden = true;
});

function insightTimestamp(d) {
  const date = new Date(d._createdAt);
  return !isNaN(date) ? date.toLocaleString("id-ID") : "";
}

function toReelsRow(d) {
  return {
    "Timestamp": insightTimestamp(d),
    "ID": d.id,
    "Divisi": d.divisi,
    "Minggu": d.week || "Minggu 1",
    "Judul Konten": d.title,
    "Script": d.script,
    "Reel Diposting": d.posted,
    "Insight Diunduh": d.downloaded,
    "Keterangan": d.caption,
    "Tayangan": d.views,
    "Pemirsa": d.reach,
    "Waktu Tonton Rata-rata": d.watchtime,
    "Suka": Number(d.likes) || 0,
    "Komentar": Number(d.comments) || 0,
    "Posting Ulang": Number(d.reposts) || 0,
    "Dibagikan": Number(d.shares) || 0,
    "Disimpan": Number(d.saves) || 0,
    "Gambar": d.image,
    "Durasi Video": d.duration,
    "Link Konten": d.link
  };
}

function toCarouselRow(d) {
  return {
    "Timestamp": insightTimestamp(d),
    "ID": d.id,
    "Divisi": d.divisi,
    "Minggu": d.week || "Minggu 1",
    "Judul Konten": d.title,
    "Isi Carousel": d.isi_carousel,
    "Tanggal Posting": d.posted,
    "Insight Diunduh": d.downloaded,
    "Tayangan": d.views,
    "Pemirsa": d.reach,
    "Kunjungan Profil": d.kunjungan,
    "Mengikuti": d.mengikuti,
    "Suka": Number(d.likes) || 0,
    "Komentar": Number(d.comments) || 0,
    "Posting Ulang": Number(d.reposts) || 0,
    "Dibagikan": Number(d.shares) || 0,
    "Disimpan": Number(d.saves) || 0,
    "Gambar": d.image,
    "Link Konten": d.link
  };
}

function buildSheetOrPlaceholder(rows) {
  return rows.length > 0
    ? XLSX.utils.json_to_sheet(rows)
    : XLSX.utils.json_to_sheet([{ "Timestamp": "", "ID": "", "Divisi": "", "Minggu": "", "Judul Konten": "" }]);
}

$("#confirmExportBtn")?.addEventListener("click", () => {
  const choice = $("#exportDivisiSelect")?.value || "all";
  let exportData = sheetData.filter(d => !d.isArchived); // Export data aktif saja
  
  if (choice !== "all") {
    exportData = exportData.filter((d) => d.divisi === choice);
  }

  if (exportData.length === 0) {
    toast("Tidak ada data aktif untuk filter divisi ini", "error");
    return;
  }

  const dataReels = exportData.filter((d) => d.type !== "carousel").map(toReelsRow).reverse();
  const dataCarousel = exportData.filter((d) => d.type === "carousel").map(toCarouselRow).reverse();

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, buildSheetOrPlaceholder(dataReels), "Insight Reels");
  XLSX.utils.book_append_sheet(wb, buildSheetOrPlaceholder(dataCarousel), "Insight Carousel");

  const fileName = `Laporan_${choice === "all" ? "Semua_Divisi" : choice}_Reelsight.xlsx`;
  XLSX.writeFile(wb, fileName);

  toast("Berhasil mengunduh Excel!", "success");
  if ($("#exportModalOverlay")) $("#exportModalOverlay").hidden = true;
});

/* ==========================================================================
   10. PLANNER LOGIC
   ========================================================================== */
const planOverlay = $("#plannerModalOverlay");

function planCardHtml(p) {
  const postingText = p.planDate 
    ? `posting tanggal : ${fmtPlanDate(p.planDate)}` 
    : `Dibuat: ${fmtDate(p.createdDate)}`;

  return `<div class="plan-header">
      <div class="plan-meta">
        <span class="plan-date">${postingText}</span>
        ${p.subdivisi ? `<span class="plan-format-badge">${escapeHtml(divisiLabel(p))}</span>` : ""}
        <span class="plan-format-badge">${p.format}</span>
      </div>
      <h3 class="plan-title">${escapeHtml(p.title)}</h3>
      <div class="plan-status-wrap">
        <select class="status-select" data-id="${p.id}">
          <option value="planned" ${p.status === "planned" ? "selected" : ""}>⏳ Planned</option>
          <option value="progress" ${p.status === "progress" ? "selected" : ""}>🔥 In Progress</option>
          <option value="done" ${p.status === "done" ? "selected" : ""}>✅ Done</option>
          <option value="cancel" ${p.status === "cancel" ? "selected" : ""}>❌ Cancelled</option>
        </select>
      </div>
    </div>
    <div class="plan-body">
      <details class="plan-details">
        <summary>Lihat Rincian Plan</summary>
        <div class="plan-details-content">
          <div class="plan-section"><strong>Objective</strong><p>${escapeHtml(p.objective) || "-"}</p></div>
          <div class="plan-section"><strong>Konsep / Visual</strong><p>${escapeHtml(p.concept) || "-"}</p></div>
          <div class="plan-section"><strong>Script</strong><p>${escapeHtml(p.script) || "-"}</p></div>
        </div>
      </details>
    </div>
    <div class="plan-footer">
      <button class="btn-edit-plan" data-act="edit-plan">✎ Edit</button>
      <button class="btn-edit-plan" data-act="del-plan" style="color:var(--pink);">🗑 Hapus</button>
    </div>`;
}

async function deletePlan(p) {
  if (!confirm(`Hapus plan "${p.title}"?`)) return;

  try {
    const res = await fetch(APPS_SCRIPT_URL, {
      method: "POST",
      body: JSON.stringify({ action: "delete_plan", id: p.id, divisi: p.divisi })
    });
    const json = await res.json();
    if (!json.ok) return;

    toast("Plan dihapus", "success");
    plannerData = plannerData.filter((x) => x.id !== p.id);
    saveJSON(LS_PLANNER_DATA, plannerData);
    renderPlannerGrid();
    if (typeof renderArsipGrid === "function") renderArsipGrid();
  } catch (e) {
    toast("Gagal menghapus plan", "error");
  }
}

function bindPlanCardEvents(card, p) {
  card.querySelector(".status-select")?.addEventListener("change", (e) => {
    const idx = plannerData.findIndex((x) => x.id === p.id);
    if (idx === -1) return;
    plannerData[idx].status = e.target.value;
    saveJSON(LS_PLANNER_DATA, plannerData);
    renderPlannerGrid();
    if (typeof renderArsipGrid === "function") renderArsipGrid();
    syncPlanToSheet(plannerData[idx]);
  });

  card.querySelector('[data-act="edit-plan"]')?.addEventListener("click", () => openPlanModal(p.id));
  card.querySelector('[data-act="del-plan"]')?.addEventListener("click", () => deletePlan(p));
}

function renderPlannerGrid() {
  const grid = $("#plannerGrid");
  if (!grid) return;
  grid.innerHTML = "";
  if (!activeDivisiPlanner) return;

  const kw = $("#planSearchInput")?.value.toLowerCase();
  const statFilt = $("#planStatusFilter")?.value;

  const list = plannerData.filter((p) => {
    if (p.divisi !== activeDivisiPlanner) return false;
    if (p.isArchived) return false;
    if (activeDivisiPlanner === "EXATA" && activeSubPlanner && activeSubPlanner !== "ALL"
        && p.subdivisi !== activeSubPlanner) return false;
    if (kw && !(p.title.toLowerCase().includes(kw) || p.objective.toLowerCase().includes(kw))) return false;
    if (statFilt && p.status !== statFilt) return false;
    return true;
  });

  if ($("#plannerEmptyHint")) $("#plannerEmptyHint").hidden = list.length > 0;

  const weeks = ["Minggu 1", "Minggu 2", "Minggu 3", "Minggu 4"];
  weeks.forEach((w) => {
    const items = list
      .filter((item) => (item.week || "Minggu 1") === w)
      .sort((a, b) => compareByDateAsc(a.planDate || a.createdDate, b.planDate || b.createdDate, a._createdAt, b._createdAt));

    if (items.length === 0) return;

    const accordion = document.createElement("details");
    accordion.className = "week-accordion";
    accordion.open = true;
    accordion.innerHTML = `
      <summary class="week-summary">
        <div style="display:flex; align-items:center; gap:12px;">
          <input type="checkbox" class="archive-checkbox" data-week="${w}" title="Centang untuk arsip" style="transform:scale(1.2); cursor:pointer;">
          <span>🗓️ ${w}</span>
        </div>
        <span class="week-count-badge">${items.length} Plan</span>
      </summary>
      <div class="week-content"></div>
    `;

    // Cegah accordion menutup saat mencentang checkbox
    const checkbox = accordion.querySelector('.archive-checkbox');
    checkbox.addEventListener('click', (e) => e.stopPropagation());

    const contentGrid = accordion.querySelector(".week-content");
    items.forEach((p) => {
      const card = document.createElement("div");
      card.className = `plan-card format-${p.format} status-${p.status}`;
      card.dataset.id = p.id;
      card.innerHTML = planCardHtml(p);
      bindPlanCardEvents(card, p);
      contentGrid.appendChild(card);
    });

    grid.appendChild(accordion);
  });
}

function syncPlanSubField() {
  const isExata = $("#p_divisi")?.value === "EXATA";
  if ($("#p_subdivisi_wrap")) $("#p_subdivisi_wrap").hidden = !isExata;
}

$("#p_divisi")?.addEventListener("change", () => {
  syncPlanSubField();
  if ($("#p_divisi").value !== "EXATA" && $("#p_subdivisi")) $("#p_subdivisi").value = "";
  autosavePlanDraft();
});

function openPlanModal(editId = null) {
  plannerEditId = editId;
  if (planOverlay) planOverlay.hidden = false;
  document.body.style.overflow = "hidden";
  
  if ($("#p_week")) $("#p_week").value = "Minggu 1";

  const saved = loadJSON(LS_PLAN_DRAFT, null);
  const useDraft = !!(saved && saved._editId === editId);

  if (editId) {
    const source = plannerData.find((x) => x.id === editId) || {};
    const d = useDraft ? { ...source, ...saved } : source;
    if ($("#planModalTitle")) $("#planModalTitle").textContent = "Edit Plan Konten";
    fillFormFields("p", fieldsPlan, d);
    if ($("#p_divisi")) $("#p_divisi").disabled = true;
    if (useDraft) toast("Melanjutkan perubahan yang belum tersimpan", "success");
  } else {
    if ($("#p_divisi")) $("#p_divisi").disabled = false;
    if ($("#planModalTitle")) $("#planModalTitle").textContent = "Tambah Plan Baru";
    if (useDraft) {
      fillFormFields("p", fieldsPlan, saved);
      toast("Melanjutkan draft yang belum tersimpan", "success");
    } else {
      fillFormFields("p", fieldsPlan, {});
      if ($("#p_format")) $("#p_format").value = "video";
      if ($("#p_divisi")) $("#p_divisi").value = activeDivisiPlanner || "EXATA";
      // otomatis pakai bagian yang sedang dibuka (AI/MT/SN), bisa diganti manual
      if ($("#p_subdivisi") && activeSubPlanner && activeSubPlanner !== "ALL") $("#p_subdivisi").value = activeSubPlanner;
    }
  }
  syncPlanSubField();
}

function closePlanModal() {
  if (planOverlay) planOverlay.hidden = true;
  document.body.style.overflow = "";
  plannerEditId = null;
  renderResumeBanners();
}

$("#openPlanModalBtn")?.addEventListener("click", () => openPlanModal(null));
$("#planModalCloseBtn")?.addEventListener("click", closePlanModal);
$("#cancelPlanModalBtn")?.addEventListener("click", closePlanModal);
planOverlay?.addEventListener("click", (e) => {
  if (e.target === planOverlay) closePlanModal();
});

function autosavePlanDraft() {
  const obj = { _editId: plannerEditId, ...readFormFields("p", fieldsPlan) };
  saveJSON(LS_PLAN_DRAFT, obj);
  flashSaved();
}

fieldsPlan.forEach((f) => {
  $(`#p_${f}`)?.addEventListener("input", autosavePlanDraft);
});

$("#savePlanBtn")?.addEventListener("click", () => {
  if (!$("#p_title")?.value.trim()) {
    toast("Judul plan wajib", "error");
    return;
  }

  const payload = readFormFields("p", fieldsPlan);
  let savedPlan = null;

  // Sub divisi hanya untuk EXATA (wajib untuk plan baru; plan lama boleh dikosongkan saat edit)
  if (payload.divisi !== "EXATA") {
    payload.subdivisi = "";
  } else if (!payload.subdivisi && !plannerEditId) {
    toast("Pilih bagian EXATA dulu (AI / MT / SN)", "error");
    return;
  }

  if (plannerEditId) {
    const idx = plannerData.findIndex((x) => x.id === plannerEditId);
    if (idx !== -1) {
      plannerData[idx] = { ...plannerData[idx], ...payload };
      savedPlan = plannerData[idx];
      toast("Plan diupdate", "success");
    }
  } else {
    payload.id = uid();
    payload.createdDate = new Date().toISOString().split("T")[0];
    payload.status = "planned";
    plannerData.unshift(payload);
    savedPlan = payload;
    toast("Plan ditambah", "success");
  }

  saveJSON(LS_PLANNER_DATA, plannerData);
  localStorage.removeItem(LS_PLAN_DRAFT);
  renderPlannerGrid();
  if (typeof renderArsipGrid === "function") renderArsipGrid();
  closePlanModal();
  if (savedPlan) syncPlanToSheet(savedPlan);
});

async function syncPlanToSheet(pObj) {
  if (!isServerConfigured()) return;
  try {
    await fetch(APPS_SCRIPT_URL, {
      method: "POST",
      body: JSON.stringify({ action: "sync_plan", plan: pObj })
    });
  } catch (e) {
    /* sync gagal, data lokal tetap tersimpan */
  }
}

async function fetchPlannerData() {
  if (!isServerConfigured()) return;
  if ($("#topLoadingBar")) $("#topLoadingBar").hidden = false;

  try {
    const res = await fetch(APPS_SCRIPT_URL + "?action=list_plans");
    const json = await res.json();
    if (json.ok) {
      plannerData = json.rows || [];
      saveJSON(LS_PLANNER_DATA, plannerData);
    }
  } catch (e) {
    /* biarkan cache lama tetap dipakai */
  } finally {
    if ($("#topLoadingBar")) $("#topLoadingBar").hidden = true;
    renderPlannerGrid();
    if (activeTab === "arsip" && typeof renderArsipGrid === "function") renderArsipGrid();
  }
}

$("#refreshPlannerBtn")?.addEventListener("click", fetchPlannerData);
["planSearchInput", "planStatusFilter"].forEach((id) => {
  $("#" + id)?.addEventListener("input", renderPlannerGrid);
  $("#" + id)?.addEventListener("change", renderPlannerGrid);
});

/* ==========================================================================
   11. FUNGSI ARCHIVE SYSTEM (BARU)
   ========================================================================== */

// Fungsi untuk Memindahkan Data (Minggu Terpilih) ke Arsip
async function archiveSelectedWeeks(type) {
  const activeDivisi = type === "sheet" ? activeDivisiLaporan : activeDivisiPlanner;
  
  if (!activeDivisi) {
    toast("Silahkan pilih divisi terlebih dahulu.", "error");
    return;
  }

  // 1. Ambil data checkbox yang sedang dicentang
  const gridId = type === "sheet" ? "#sheetGrid" : "#plannerGrid";
  const grid = document.querySelector(gridId);
  if (!grid) return;

  const checkedBoxes = Array.from(grid.querySelectorAll('.archive-checkbox:checked'));
  if (checkedBoxes.length === 0) {
    toast("Pilih minimal satu kotak minggu untuk diarsipkan.", "error");
    return;
  }

  // 2. Minta User menamai Arsip ini
  const customArchiveName = prompt("Masukkan nama untuk Arsip ini (Contoh: 'Arsip Campaign Lebaran' atau 'Bulan Maret'):");
  if (!customArchiveName || !customArchiveName.trim()) {
    toast("Nama arsip tidak boleh kosong. Proses dibatalkan.", "error");
    return;
  }

  const targetWeeks = checkedBoxes.map(cb => cb.dataset.week);
  const safeName = customArchiveName.trim();
  let count = 0;

  // 3. Eksekusi Pemindahan ke Arsip dengan nama custom
  if (type === "sheet") {
    sheetData.forEach(d => {
      const weekVal = d.week || "Minggu 1";
      const subOk = activeDivisi !== "EXATA" || !activeSubLaporan || activeSubLaporan === "ALL"
        || d.subdivisi === activeSubLaporan;
      if (d.divisi === activeDivisi && subOk && !d.isArchived && targetWeeks.includes(weekVal)) {
        d.isArchived = true;
        d.archiveName = safeName; // Simpan properti nama custom
        count++;
        if (isServerConfigured()) {
          fetch(APPS_SCRIPT_URL, { method: "POST", body: JSON.stringify({ action: "update_insight", data: d }) }).catch(e=>e);
        }
      }
    });
    if (count > 0) {
      saveJSON(LS_SHEETCACHE, sheetData);
      renderSheetGrid();
    }
  } else if (type === "planner") {
    plannerData.forEach(p => {
      const weekVal = p.week || "Minggu 1";
      const subOkP = activeDivisi !== "EXATA" || !activeSubPlanner || activeSubPlanner === "ALL"
        || p.subdivisi === activeSubPlanner;
      if (p.divisi === activeDivisi && subOkP && !p.isArchived && targetWeeks.includes(weekVal)) {
        p.isArchived = true;
        p.archiveName = safeName; // Simpan properti nama custom
        count++;
        if (isServerConfigured()) {
          fetch(APPS_SCRIPT_URL, { method: "POST", body: JSON.stringify({ action: "sync_plan", plan: p }) }).catch(e=>e);
        }
      }
    });
    if (count > 0) {
      saveJSON(LS_PLANNER_DATA, plannerData);
      renderPlannerGrid();
    }
  }

  if (count > 0) {
    toast(`${count} item berhasil diarsipkan sebagai "${safeName}".`, "success");
    if (activeTab === "arsip") renderArsipGrid();
  }
}

// Fungsi untuk Mengembalikan Data dari Arsip
function unarchiveItem(id, type) {
  if (type === "sheet") {
    const item = sheetData.find(x => x.id === id);
    if (item) {
      item.isArchived = false;
      saveJSON(LS_SHEETCACHE, sheetData);
      if (isServerConfigured()) {
        fetch(APPS_SCRIPT_URL, { method: "POST", body: JSON.stringify({ action: "update_insight", data: item }) }).catch(e=>e);
      }
    }
  } else {
    const item = plannerData.find(x => x.id === id);
    if (item) {
      item.isArchived = false;
      saveJSON(LS_PLANNER_DATA, plannerData);
      if (isServerConfigured()) {
        fetch(APPS_SCRIPT_URL, { method: "POST", body: JSON.stringify({ action: "sync_plan", plan: item }) }).catch(e=>e);
      }
    }
  }

  toast("Data berhasil dikembalikan dari arsip.", "success");
  renderArsipGrid();
  if (activeDivisiLaporan) renderSheetGrid();
  if (activeDivisiPlanner) renderPlannerGrid();
}

// Merender Isi Halaman Arsip
function renderArsipGrid() {
  const grid = $("#arsipGrid");
  if (!grid) return;
  grid.innerHTML = "";

  const kw = $("#arsipSearch")?.value.toLowerCase();
  const statusFilt = $("#arsipStatus")?.value;
  const dateFilt = $("#arsipDate")?.value;

  const archivedSheets = sheetData.filter(d => {
    if (!d.isArchived) return false;
    if (kw && !(d.title || "").toLowerCase().includes(kw)) return false;
    if (dateFilt && d.posted !== dateFilt) return false;
    return true;
  });

  const archivedPlans = plannerData.filter(p => {
    if (!p.isArchived) return false;
    if (kw && !(p.title.toLowerCase().includes(kw) || (p.objective && p.objective.toLowerCase().includes(kw)))) return false;
    if (statusFilt && p.status !== statusFilt) return false;
    if (dateFilt && p.planDate !== dateFilt) return false;
    return true;
  });

  if (archivedSheets.length === 0 && archivedPlans.length === 0) {
    grid.innerHTML = `<div class="empty-hint" style="text-align:center; padding: 40px; color:var(--gray-dark);">Belum ada data (atau data sesuai pencarian) yang diarsipkan.</div>`;
    return;
  }

  // Helper function untuk mengelompokkan arsip berdasar custom "archiveName"
  // Item diurutkan dulu (terlama -> terbaru) sehingga kartu di tiap folder arsip juga berurutan.
  const groupByArchiveName = (items, dateOf) => {
    const ordered = [...items].sort((x, y) => compareByDateAsc(dateOf(x), dateOf(y), x._createdAt, y._createdAt));
    return ordered.reduce((acc, item) => {
      const name = item.archiveName || `Arsip ${item.week || "Lama"}`; // Fallback jika data lama belum punya nama
      if (!acc[name]) acc[name] = [];
      acc[name].push(item);
      return acc;
    }, {});
  };

  // 1. Tampilkan Grid Laporan Arsip (Dikelompokkan berdasar Nama Custom)
  if (archivedSheets.length > 0) {
    const secSheet = document.createElement("div");
    secSheet.innerHTML = `<h3 style="margin: 16px 0 8px; color:var(--primary); font-size:16px;">📊 Laporan Insight Tersimpan</h3>`;
    
    const groupedSheets = groupByArchiveName(archivedSheets, (d) => d.posted);
    for (const [aName, items] of Object.entries(groupedSheets)) {
      const accordion = document.createElement("details");
      accordion.className = "week-accordion";
      accordion.open = true;
      accordion.innerHTML = `<summary class="week-summary"><span>📁 ${aName}</span><span class="week-count-badge">${items.length} Item</span></summary><div class="week-content"></div>`;
      
      const cGrid = accordion.querySelector(".week-content");
      items.forEach(d => {
        const card = buildCard(d, false, "sheet"); 
        const unarchBtn = document.createElement("button");
        unarchBtn.className = "btn-ghost";
        unarchBtn.style.cssText = "width:100%; margin-top:8px; border:1px solid var(--border);";
        unarchBtn.innerHTML = "↖️ Kembalikan (Unarchive)";
        unarchBtn.onclick = () => unarchiveItem(d.id, "sheet");
        card.querySelector(".card-body").appendChild(unarchBtn);
        cGrid.appendChild(card);
      });
      secSheet.appendChild(accordion);
    }
    grid.appendChild(secSheet);
  }

  // 2. Tampilkan Grid Planner Arsip (Dikelompokkan berdasar Nama Custom)
  if (archivedPlans.length > 0) {
    const secPlan = document.createElement("div");
    secPlan.innerHTML = `<h3 style="margin: 24px 0 8px; color:var(--primary); font-size:16px;">🗓️ Planner Tersimpan</h3>`;
    
    const groupedPlans = groupByArchiveName(archivedPlans, (p) => p.planDate || p.createdDate);
    for (const [aName, items] of Object.entries(groupedPlans)) {
      const accordion = document.createElement("details");
      accordion.className = "week-accordion";
      accordion.open = true;
      accordion.innerHTML = `<summary class="week-summary"><span>📁 ${aName}</span><span class="week-count-badge">${items.length} Item</span></summary><div class="week-content"></div>`;
      
      const cGrid = accordion.querySelector(".week-content");
      items.forEach(p => {
        const card = document.createElement("div");
        card.className = `plan-card format-${p.format} status-${p.status}`;
        card.dataset.id = p.id;
        card.innerHTML = planCardHtml(p);
        bindPlanCardEvents(card, p);

        const unarchBtn = document.createElement("button");
        unarchBtn.className = "btn-ghost";
        unarchBtn.style.cssText = "width:100%; margin-top:8px; border:1px solid var(--border);";
        unarchBtn.innerHTML = "↖️ Kembalikan (Unarchive)";
        unarchBtn.onclick = () => unarchiveItem(p.id, "planner");
        card.querySelector(".plan-footer").appendChild(unarchBtn);
        cGrid.appendChild(card);
      });
      secPlan.appendChild(accordion);
    }
    grid.appendChild(secPlan);
  }
}

/* ==========================================================================
   12. BANNER "LANJUTKAN DRAFT" (autosave lintas tab/refresh/restart)
   ========================================================================== */
function resumeInsightDraft() {
  const d = loadJSON(LS_FORMDRAFT, null);
  if (!d) return;

  if (d._context === "sheet" && d.divisi) {
    switchTab("lihat");
    activeDivisiLaporan = d.divisi;
    activeSubLaporan = d.divisi === "EXATA" ? (d.subdivisi || "ALL") : null;
    showLaporanContent();
  } else {
    switchTab("input");
  }
  openModal(d._editId || null, d.type || "reels", d._context || "draft");
}

function discardInsightDraft() {
  localStorage.removeItem(LS_FORMDRAFT);
  renderResumeBanners();
  toast("Draft dibuang", "");
}

function resumePlanDraft() {
  const d = loadJSON(LS_PLAN_DRAFT, null);
  if (!d) return;

  switchTab("planner");
  if (d.divisi) {
    activeDivisiPlanner = d.divisi;
    activeSubPlanner = d.divisi === "EXATA" ? (d.subdivisi || "ALL") : null;
    showPlannerContent();
  }
  openPlanModal(d._editId || null);
}

function discardPlanDraft() {
  localStorage.removeItem(LS_PLAN_DRAFT);
  renderResumeBanners();
  toast("Draft plan dibuang", "");
}

function renderResumeBanners() {
  const bar = $("#draftResumeBar");
  if (!bar) return;

  const items = [];
  const insightDraft = loadJSON(LS_FORMDRAFT, null);
  if (insightDraft) {
    items.push({
      key: "insight",
      label: insightDraft._editId ? "Mengedit insight" : "Insight baru",
      title: insightDraft.title || "(Tanpa judul)"
    });
  }
  const planDraft = loadJSON(LS_PLAN_DRAFT, null);
  if (planDraft) {
    items.push({
      key: "plan",
      label: planDraft._editId ? "Mengedit plan" : "Plan baru",
      title: planDraft.title || "(Tanpa judul)"
    });
  }

  if (items.length === 0) {
    bar.hidden = true;
    bar.innerHTML = "";
    return;
  }

  bar.hidden = false;
  bar.innerHTML = items.map((it) => `
    <div class="resume-bar-item">
      <span>📝 ${it.label}: <strong>${escapeHtml(it.title)}</strong> — belum tersimpan</span>
      <div class="resume-bar-actions">
        <button type="button" data-resume="${it.key}">Lanjutkan</button>
        <button type="button" data-discard="${it.key}">Buang</button>
      </div>
    </div>
  `).join("");

  bar.querySelector('[data-resume="insight"]')?.addEventListener("click", resumeInsightDraft);
  bar.querySelector('[data-discard="insight"]')?.addEventListener("click", discardInsightDraft);
  bar.querySelector('[data-resume="plan"]')?.addEventListener("click", resumePlanDraft);
  bar.querySelector('[data-discard="plan"]')?.addEventListener("click", discardPlanDraft);
}

/* ==========================================================================
   13. INIT
   ========================================================================== */
renderDrafts();
renderResumeBanners();
