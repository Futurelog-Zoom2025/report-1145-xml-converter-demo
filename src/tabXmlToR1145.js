// Tab module for "Convert XML to Report 1145".
//
// Flow:
//   1. User drops/selects a .xml file
//   2. Parser extracts rows + validity metadata
//   3. If filename matches the forward-generator pattern, decode and show
//      Division / Supplier / Date as an info block (no preview table)
//   4. Show Full Data button opens the shared modal
//   5. Download button writes the Report 1145 .xlsx

import { parseFuturelogXml } from "./xmlParser.js";
import { buildReport1145Xlsx } from "./r1145Writer.js";
import {
  $, escapeHtml, formatBytes, runWithLoading, downloadBlob,
} from "./shared.js";
import { openFullDataModal } from "./fullDataModal.js";

// Full Data modal columns — everything we extracted from the XML.
const FULL_COLS = [
  { key: "pos",          label: "#" },
  { key: "itemNo",       label: "Article no." },
  { key: "ean",          label: "EAN / GTIN" },
  { key: "manArtId",     label: "Mfg item no" },
  { key: "descDE",       label: "Name (DE)" },
  { key: "descFR",       label: "Name (FR)" },
  { key: "descIT",       label: "Name (IT)" },
  { key: "descGB",       label: "Name (GB)" },
  { key: "descExtra",    label: "Name (Local)" },
  { key: "ou",           label: "OU" },
  { key: "cu",           label: "CU" },
  { key: "cuou",         label: "CU/OU" },
  { key: "priceOU",      label: "Price (PRCOU)" },
  { key: "origin",       label: "Origin" },
  { key: "customsNo",    label: "Customs no" },
  { key: "availability", label: "Lead time (VLZ)" },
  { key: "specUrl",      label: "Spec URL" },
  { key: "offerStart",   label: "Offer start" },
  { key: "offerEnd",     label: "Offer end" },
  { key: "customerId",   label: "Customer ID" },
];

const MONTH_NAMES_EN = [
  "Jan", "Feb", "Mar", "Apr", "May", "Jun",
  "Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
];

// Decode a filename matching the forward XML generator's output. Accepted:
//   16911111120251202.cat.xml    ← exact forward-generator output
//   16911111120251202_cat.xml    ← underscore variant (some renames)
//   16911111120251202.xml        ← stripped variant
//
// Pattern: 3 digits (division) + 6 digits (supplier) + 8 digits (YYYYMMDD)
// followed by ".cat.xml", "_cat.xml", or ".xml". Returns null when the
// filename doesn't match (caller hides the decoded-info card in that case).
function decodeFilename(name) {
  if (typeof name !== "string" || name === "") return null;

  const m = name.match(/^(\d{17})(?:[._]cat)?\.xml$/i);
  if (!m) return null;

  const body = m[1];
  const division   = body.slice(0, 3);
  const supplierNo = body.slice(3, 9);
  const yyyy       = body.slice(9, 13);
  const mm         = body.slice(13, 15);
  const dd         = body.slice(15, 17);

  // Validate month/day ranges so a 17-digit string that isn't actually a
  // real date doesn't pretend to be one. Year range covers ongoing files.
  const monthIdx = parseInt(mm, 10) - 1;
  const day      = parseInt(dd, 10);
  const year     = parseInt(yyyy, 10);
  if (
    monthIdx < 0 || monthIdx > 11 ||
    day < 1 || day > 31 ||
    year < 2000 || year > 2200
  ) {
    return null;
  }

  return {
    division,
    supplierNo,
    yyyy, mm, dd,
    dateHuman: `${dd} ${MONTH_NAMES_EN[monthIdx]} ${yyyy}`,
  };
}

export function initXmlToR1145Tab() {
  const els = {
    dropZone:        $("#xmlDropZone"),
    fileInput:       $("#xmlFileInput"),
    fileInfo:        $("#xmlFileInfo"),
    status:          $("#xmlStatus"),
    resetBtn:        $("#xmlResetBtn"),
    decodedCard:     $("#xmlDecodedCard"),
    decodedDivision: $("#xmlDecodedDivision"),
    decodedSupplier: $("#xmlDecodedSupplier"),
    decodedDate:     $("#xmlDecodedDate"),
    showFullBtn:     $("#xmlShowFullBtn"),
    actionCard:      $("#xmlActionCard"),
    generateBtn:     $("#xmlGenerateBtn"),
    genHint:         $("#xmlGenHint"),
  };

  const state = {
    rows: [],
    validity: null,
    fileName: null,
    decoded: null,
  };

  // ─────────────── Helpers ───────────────

  function setStatus(kind, html) {
    els.status.className = `status ${kind}`;
    els.status.innerHTML = html;
    els.status.classList.remove("hidden");
  }
  function clearStatus() {
    els.status.className = "status hidden";
    els.status.innerHTML = "";
  }

  function setGenerateReady(kind, detail) {
    if (kind === "empty") {
      els.generateBtn.disabled = true;
      els.genHint.textContent = detail || "Upload an XML file first.";
      els.genHint.className = "gen-hint";
    } else if (kind === "ready") {
      els.generateBtn.disabled = false;
      els.genHint.textContent = detail || "Ready — click Download.";
      els.genHint.className = "gen-hint ready";
    }
  }

  function resetAll() {
    state.rows = [];
    state.validity = null;
    state.fileName = null;
    state.decoded = null;
    els.fileInput.value = "";
    els.fileInfo.classList.add("hidden");
    els.fileInfo.innerHTML = "";
    els.decodedCard.classList.add("hidden");
    els.actionCard.classList.add("hidden");
    clearStatus();
    setGenerateReady("empty");
  }

  function renderDecoded(decoded) {
    if (!decoded) {
      els.decodedCard.classList.add("hidden");
      return;
    }
    els.decodedDivision.textContent = decoded.division;
    els.decodedSupplier.textContent = decoded.supplierNo;
    els.decodedDate.textContent     = decoded.dateHuman;
    els.decodedCard.classList.remove("hidden");
  }

  // ─────────────── File handling ───────────────

  async function handleFile(file) {
    clearStatus();
    state.fileName = file.name;

    try {
      const text = await file.text();

      // Parse — wrapped in the loading overlay since DOMParser can take a
      // moment on multi-MB XML files.
      const { rows, validity } = await runWithLoading(
        "Parsing XML…",
        `Reading ${formatBytes(file.size)} from ${file.name}.`,
        () => parseFuturelogXml(text)
      );

      state.rows = rows;
      state.validity = validity;
      state.decoded = decodeFilename(file.name);

      els.fileInfo.innerHTML = `
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="width:18px;height:18px;color:var(--success)"><path d="M20 6L9 17l-5-5"/></svg>
        <span class="name">${escapeHtml(file.name)}</span>
        <span class="size">· ${formatBytes(file.size)} · ${rows.length} article${rows.length === 1 ? "" : "s"}</span>
      `;
      els.fileInfo.classList.remove("hidden");

      renderDecoded(state.decoded);
      els.actionCard.classList.remove("hidden");
      setGenerateReady("ready", `${rows.length} row${rows.length === 1 ? "" : "s"} ready.`);

      // Status message: surface XML validity only when filename didn't decode
      // (otherwise the decoded card already shows the date prominently).
      setStatus("info",
        `<strong>Parsed.</strong> ${rows.length} article${rows.length === 1 ? "" : "s"} extracted` +
        (validity.raw && !state.decoded
          ? ` · XML validity ${escapeHtml(validity.dd || "")}/${escapeHtml(validity.mm || "")}/${escapeHtml(validity.yyyy || "")}`
          : "") +
        `. Click <em>Download Report 1145</em> below.`
      );
    } catch (err) {
      console.error(err);
      state.rows = [];
      state.validity = null;
      state.decoded = null;
      els.fileInfo.classList.add("hidden");
      els.decodedCard.classList.add("hidden");
      els.actionCard.classList.add("hidden");
      setGenerateReady("empty");
      setStatus("error",
        `<h3>Could not parse the XML</h3>${escapeHtml(err.message || String(err))}`
      );
    }
  }

  // ─────────────── Generate ───────────────

  async function runGenerate() {
    if (state.rows.length === 0) return;

    const result = await runWithLoading(
      "Building Excel file…",
      `Writing ${state.rows.length.toLocaleString()} row${state.rows.length === 1 ? "" : "s"}.`,
      () => buildReport1145Xlsx(state.rows)
    );
    const { blob, suggestedFilename } = result;
    downloadBlob(blob, suggestedFilename,
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");

    setStatus("success",
      `<h3>Excel file downloaded</h3>File: <code>${escapeHtml(suggestedFilename)}</code> · ${state.rows.length} row${state.rows.length === 1 ? "" : "s"}`
    );
  }

  // ─────────────── Event wiring ───────────────

  els.dropZone.addEventListener("click", () => els.fileInput.click());
  els.dropZone.addEventListener("dragover", (e) => {
    e.preventDefault();
    els.dropZone.classList.add("dragging");
  });
  els.dropZone.addEventListener("dragleave", () => els.dropZone.classList.remove("dragging"));
  els.dropZone.addEventListener("drop", (e) => {
    e.preventDefault();
    els.dropZone.classList.remove("dragging");
    const f = e.dataTransfer.files[0];
    if (f) handleFile(f);
  });
  els.fileInput.addEventListener("change", (e) => {
    const f = e.target.files[0];
    if (f) handleFile(f);
  });

  els.generateBtn.addEventListener("click", runGenerate);
  els.resetBtn.addEventListener("click", resetAll);

  els.showFullBtn.addEventListener("click", () => {
    openFullDataModal({
      rows: state.rows,
      columns: FULL_COLS,
      invalidCells: new Map(),
      exportFilename: "XML_to_R1145_Preview",
      exportSheetName: "XML Rows",
    });
  });

  setGenerateReady("empty");
}
