"use strict";

/**
 * Equivalent van VBA-const N
 */
const N = 9;

const els = {
  userInput: document.getElementById("userInput"),
  code1: document.getElementById("code1"),
  code2: document.getElementById("code2"),
  statusCell: document.getElementById("statusCell"),

  substitutionInput: document.getElementById("substitutionInput"),
  saveSubBtn: document.getElementById("saveSubBtn"),
  resetSubBtn: document.getElementById("resetSubBtn"),

  clearBtn: document.getElementById("clearBtn"),
  exportBtn: document.getElementById("exportBtn"),
  clearArchiveBtn: document.getElementById("clearArchiveBtn"),

  archiveTableBody: document.querySelector("#archiveTable tbody"),
};

const STORAGE_KEYS = {
  user: "bm_user",
  substitution: "bm_substitution",
  archive: "bm_archive",
};

let clearTimer = null;

/**
 * === Helpers: VBA-equivalent ===
 */

// VBA: ExtractUserLetters(raw) -> enkel letters vóór spatie, uppercase
function extractUserLetters(raw) {
  const s = String(raw ?? "").trim();
  if (!s) return "";
  const token = s.split(" ")[0]; // vóór eerste spatie
  let out = "";
  for (const ch of token) {
    if (/[A-Za-z]/.test(ch)) out += ch;
  }
  return out.toUpperCase();
}

// VBA: GetCodePrefix(cell) -> eerste N tekens, uppercase
// In web: gewoon string trimmen, en eerste N nemen.
// (We doen geen numeric-format met voorloopnullen; scanners leveren meestal tekst.
//  Als je wél padding wil, kan dat hier.)
function getCodePrefix(value) {
  const raw = String(value ?? "").trim();
  return raw.slice(0, N).toUpperCase();
}

// Substitutie: parse textarea -> lijst van [expPrefix, subPrefix]
function parseSubstitution(text) {
  const lines = String(text ?? "")
    .split(/\r?\n/)
    .map(l => l.trim())
    .filter(l => l && !l.startsWith("#"));

  const pairs = [];
  for (const line of lines) {
    const parts = line.split(",").map(p => p.trim());
    if (parts.length < 2) continue;
    const exp = getCodePrefix(parts[0]);
    const sub = getCodePrefix(parts[1]);
    if (exp && sub) pairs.push([exp, sub]);
  }
  return pairs;
}

// VBA: IsAllowedCombination(firstN_1, firstN_2)
function isAllowedCombination(p1, p2, substitutionPairs) {
  if (!p1 || !p2) return false;

  // directe match
  if (p1.toUpperCase() === p2.toUpperCase()) return true;

  // substitutie match in beide richtingen
  for (const [exp, sub] of substitutionPairs) {
    if ((p1 === exp && p2 === sub) || (p2 === exp && p1 === sub)) return true;
  }
  return false;
}

// Status opmaak
function setStatus(status) {
  const normalized = String(status ?? "").trim().toLowerCase();
  els.statusCell.classList.remove("status-incomplete", "status-match", "status-nomatch");

  if (normalized === "match") {
    els.statusCell.classList.add("status-match");
    els.statusCell.textContent = "Match";
  } else if (normalized === "no match" || normalized === "geen match") {
    els.statusCell.classList.add("status-nomatch");
    els.statusCell.textContent = "No match";
  } else {
    els.statusCell.classList.add("status-incomplete");
    els.statusCell.textContent = "incomplete";
  }
}

// Archief opslag
function loadArchive() {
  try {
    const raw = localStorage.getItem(STORAGE_KEYS.archive);
    const arr = raw ? JSON.parse(raw) : [];
    return Array.isArray(arr) ? arr : [];
  } catch {
    return [];
  }
}

function saveArchive(items) {
  localStorage.setItem(STORAGE_KEYS.archive, JSON.stringify(items));
}

function addArchiveRow(row) {
  const archive = loadArchive();
  archive.unshift(row); // nieuwste bovenaan
  saveArchive(archive);
  renderArchive();
}

function renderArchive() {
  const archive = loadArchive();
  els.archiveTableBody.innerHTML = "";
  for (const item of archive) {
    const tr = document.createElement("tr");
    const cols = [
      item.datetime,
      item.user,
      item.full1,
      item.full2,
      item.prefix1,
      item.prefix2,
      item.status,
    ];
    for (const c of cols) {
      const td = document.createElement("td");
      td.textContent = c;
      tr.appendChild(td);
    }
    els.archiveTableBody.appendChild(tr);
  }
}

// CSV export
function downloadText(filename, text) {
  const blob = new Blob([text], { type: "text/plain;charset=utf-8" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(a.href);
}

function exportArchiveCsv() {
  const archive = loadArchive();
  const header = [
    "Datum/tijd",
    "Gebruiker",
    "Barcode 1 (volledig)",
    "Barcode 2 (volledig)",
    `Eerste ${N} (A2)`,
    `Eerste ${N} (B2)`,
    "Status",
  ];

  const escapeCsv = (v) => {
    const s = String(v ?? "");
    if (/[",\n\r]/.test(s)) return `"${s.replace(/"/g, '""')}"`;
    return s;
  };

  const lines = [header.map(escapeCsv).join(",")];
  for (const item of archive) {
    lines.push([
      item.datetime,
      item.user,
      item.full1,
      item.full2,
      item.prefix1,
      item.prefix2,
      item.status,
    ].map(escapeCsv).join(","));
  }

  const ts = new Date().toISOString().slice(0, 19).replace(/[:T]/g, "-");
  downloadText(`archief-${ts}.csv`, lines.join("\n"));
}

/**
 * === Main logic (Worksheet_Change equivalent) ===
 */
function normalizeUserIfNeeded() {
  const raw = els.userInput.value;
  const cleaned = extractUserLetters(raw);
  if (raw !== cleaned) {
    els.userInput.value = cleaned;
  }
  localStorage.setItem(STORAGE_KEYS.user, cleaned);
}

function resetToIncomplete() {
  setStatus("incomplete");
}

function clearCodesAndReset() {
  if (clearTimer) {
    clearTimeout(clearTimer);
    clearTimer = null;
  }
  els.code1.value = "";
  els.code2.value = "";
  resetToIncomplete();
  els.code1.focus();
}

function evaluateIfReady(source) {
  // 1) gebruiker verplicht
  const userName = String(els.userInput.value ?? "").trim();
  if (!userName) {
    alert("Gelieve eerst uw gebruikersbarcode te scannen.");
    els.userInput.focus();
    return;
  }

  // 2) status reset
  resetToIncomplete();

  // 3) barcodes lezen
  const full1 = String(els.code1.value ?? "").trim();
  const full2 = String(els.code2.value ?? "").trim();

  // Nog niet genoeg tekens?
  if (full1.length < N || full2.length < N) {
    if (source === "code1") els.code2.focus();
    return;
  }

  const prefix1 = getCodePrefix(full1);
  const prefix2 = getCodePrefix(full2);

  const substitutionPairs = parseSubstitution(els.substitutionInput.value);

  let status = "No match";
  let isMatch = false;

  if (isAllowedCombination(prefix1, prefix2, substitutionPairs)) {
    status = "Match";
    isMatch = true;
  }

  setStatus(status);

  // loggen
  const dt = new Date();
  const datetime = dt.toLocaleString("nl-BE", {
    year: "numeric", month: "2-digit", day: "2-digit",
    hour: "2-digit", minute: "2-digit", second: "2-digit",
  });

  addArchiveRow({
    datetime,
    user: userName,
    full1,
    full2,
    prefix1,
    prefix2,
    status,
  });

  // gedrag na resultaat
  if (isMatch) {
    if (clearTimer) clearTimeout(clearTimer);
    clearTimer = setTimeout(() => {
      clearCodesAndReset();
    }, 5000);
  } else {
    if (source === "code1") els.code2.focus();
  }
}

/**
 * === Init / UI wiring ===
 */
function loadPersisted() {
  const user = localStorage.getItem(STORAGE_KEYS.user);
  if (user) els.userInput.value = user;

  const subs = localStorage.getItem(STORAGE_KEYS.substitution);
  if (subs) {
    els.substitutionInput.value = subs;
  } else {
    // voorbeeld
    els.substitutionInput.value = [
      "# Voorbeeld:",
      "# EXPECTED,SUBSTITUTE",
      "ABCDEF123,XYZ987654",
      "111111111,222222222",
    ].join("\n");
  }
}

function persistSubstitution() {
  localStorage.setItem(STORAGE_KEYS.substitution, els.substitutionInput.value);
}

els.userInput.addEventListener("change", () => {
  normalizeUserIfNeeded();
});
els.userInput.addEventListener("input", () => {
  // scanners sturen vaak een volledige string in 1 keer; input is prima
  // we normaliseren pas bij change/blur om “springende cursor” te vermijden
});

els.code1.addEventListener("input", () => evaluateIfReady("code1"));
els.code2.addEventListener("input", () => evaluateIfReady("code2"));

els.clearBtn.addEventListener("click", clearCodesAndReset);

els.saveSubBtn.addEventListener("click", () => {
  persistSubstitution();
  alert("Substitutie opgeslagen.");
});

els.resetSubBtn.addEventListener("click", () => {
  els.substitutionInput.value = [
    "# Voorbeeld:",
    "# EXPECTED,SUBSTITUTE",
    "ABCDEF123,XYZ987654",
    "111111111,222222222",
  ].join("\n");
  persistSubstitution();
});

els.exportBtn.addEventListener("click", exportArchiveCsv);

els.clearArchiveBtn.addEventListener("click", () => {
  if (!confirm("Archief zeker wissen?")) return;
  saveArchive([]);
  renderArchive();
});

window.addEventListener("load", () => {
  loadPersisted();
  renderArchive();
  resetToIncomplete();
  // focus workflow: eerst user scannen, anders code1
  if (!String(els.userInput.value ?? "").trim()) els.userInput.focus();
  else els.code1.focus();
});

