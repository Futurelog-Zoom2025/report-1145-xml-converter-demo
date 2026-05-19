// Tab module for "Convert XML to Report 1145".
//
// Flow:
//   1. User drops/selects a .xml file
//   2. Parser extracts rows + validity metadata
//   3. Inline preview + Show Full Data modal
//   4. Download button writes the Report 1145 .xlsx

import { parseFuturelogXml } from "./xmlParser.js";
import { buildReport1145Xlsx } from "./r1145Writer.js";
import {
  $, escapeHtml, formatBytes, runWithLoading, downloadBlob,
} from "./shared.js";
import { openFullDataModal } from "./fullDataModal.js";

const PREVIEW_LIMIT = 10;

// Inline preview columns — small subset, chosen to fit on screen without
// horizontal scroll. The Full Data modal shows the complete set.
const PREVIEW_COLS = [
  { key: "pos",        label: "#",          cls: "c-pos" },
  { key: "itemNo",     label: "Article",    cls: "c-art" },
  { key: "descGB",     label: "Name (GB)",  cls: "c-en" },
  { key: "ou",         label: "OU",         cls: "c-unit" },
  { key: "priceOU",    label: "Price",      cls: "c-price" },
  { key: "availability", label: "Lead",     cls: "c-avail" },
  { key: "customerId", label: "Cust ID",    cls: "c-cust" },
];

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

export function initXmlToR1145Tab() {
  const els = {
    dropZone:        $("#xmlDropZone"),
    fileInput:       $("#xmlFileInput"),
    fileInfo:        $("#xmlFileInfo"),
    status:          $("#xmlStatus"),
    resetBtn:        $("#xmlResetBtn"),
    previewCard:     $("#xmlPreviewCard"),
    previewSummary:  $("#xmlPreviewSummary"),
    previewTable:    $("#xmlPreviewTable"),
    showFullBtn:     $("#xmlShowFullBtn"),
    actionCard:      $("#xmlActionCard"),
    generateBtn:     $("#xmlGenerateBtn"),
    genHint:         $("#xmlGenHint"),
  };

  const state = {
    rows: [],
    validity: null,
    fileName: null,
  };

  // ───────────────────────── Helpers ─────────────────────────

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
    els.fileInput.value = "";
    els.fileInfo.classList.add("hidden");
    els.fileInfo.innerHTML = "";
    els.previewCard.classList.add("hidden");
    els.actionCard.classList.add("hidden");
    clearStatus();
    setGenerateReady("empty");
  }

  // ─────────────────────── File handling ──────────────────────

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

      els.fileInfo.innerHTML = `
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="width:18px;height:18px;color:var(--success)"><path d="M20 6L9 17l-5-5"/></svg>
        <span class="name">${escapeHtml(file.name)}</span>
        <span class="size">· ${formatBytes(file.size)} · ${rows.length} article${rows.length === 1 ? "" : "s"}</span>
      `;
      els.fileInfo.classList.remove("hidden");

      renderPreview(rows);
      els.actionCard.classList.remove("hidden");
      setGenerateReady("ready");

      setStatus("info",
        `<strong>Parsed.</strong> ${rows.length} article${rows.length === 1 ? "" : "s"} extracted` +
        (validity.raw ? ` · validity ${escapeHtml(validity.dd || "")}/${escapeHtml(validity.mm || "")}/${escapeHtml(validity.yyyy || "")}` : "") +
        `. Click <em>Download Report 1145</em> below.`
      );
    } catch (err) {
      console.error(err);
      state.rows = [];
      state.validity = null;
      els.fileInfo.classList.add("hidden");
      els.previewCard.classList.add("hidden");
      els.actionCard.classList.add("hidden");
      setGenerateReady("empty");
      setStatus("error",
        `<h3>Could not parse the XML</h3>${escapeHtml(err.message || String(err))}`
      );
    }
  }

  // ─────────────────────── Preview ─────────────────────────────

  function displayValue(v) {
    if (v === null || v === undefined) return "";
    if (typeof v === "number") {
      return Number.isInteger(v) ? String(v) : v.toFixed(2);
    }
    return String(v);
  }

  function renderPreview(rows) {
    const showRows = rows.slice(0, PREVIEW_LIMIT);
    const head = `<thead><tr>${PREVIEW_COLS.map((c) => `<th class="${c.cls}">${c.label}</th>`).join("")}</tr></thead>`;
    const body = "<tbody>" + showRows.map((r) => {
      return "<tr>" + PREVIEW_COLS.map((c) => {
        const v = r[c.key];
        const shown = displayValue(v);
        return `<td class="${c.cls}" title="${escapeHtml(shown)}">${escapeHtml(shown)}</td>`;
      }).join("") + "</tr>";
    }).join("") + "</tbody>";
    els.previewTable.innerHTML = head + body;

    els.previewSummary.textContent =
      rows.length > PREVIEW_LIMIT
        ? `Showing first ${PREVIEW_LIMIT} of ${rows.length} rows.`
        : `Showing all ${rows.length} row${rows.length === 1 ? "" : "s"}.`;
    els.previewCard.classList.remove("hidden");
  }

  // ─────────────────────── Generate ────────────────────────────

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

  // ─────────────────────── Event wiring ────────────────────────

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
      invalidCells: new Map(),   // no validation on this tab — no error cells
      exportFilename: "XML_to_R1145_Preview",
      exportSheetName: "XML Rows",
    });
  });

  setGenerateReady("empty");
}
