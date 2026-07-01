// Merges parsed Makro price data into parsed Report 1145 rows, producing the
// Tabelle1-shaped rows the XML generator consumes.
//
// Ported from the VBA Get_Data_From_FileReport1145 (Module2). Key differences
// from the P2P merger:
//   - The DRIVING list is Report 1145 (every 1145 row appears in the output),
//     NOT the Makro file. Makro codes with no matching 1145 article are ignored
//     (counted in the summary only), exactly like the workbook's XLOOKUP does.
//   - The join key is Report 1145 "Article no." (itemNo) === Makro "รหัสสินค้า".
//   - Lead time is set UNCONDITIONALLY (no toggle): matched → "1", unmatched → "0".
//
// Per-row rule (VBA):
//   AA "New Price per order unit"      = XLOOKUP(Article no., รหัสสินค้า, L, "No Information")
//   AB "New Price per order unit(Adj)" = ROUND( XLOOKUP(..., L, <old 1145 price>), 2)
//   AC "Diff Price Between Old and New"= <old 1145 price> - AB
//   AF "Article lead time(Adj)"        = IIf(AA = "No Information", "0", "1")
//
// So:
//   matched   → price = Round(Makro L, 2)   (Makro.newPrice); lead time = "1"
//   unmatched → price = old 1145 price;      lead time = "0"; flagged "No Information"

import { vbaRound } from "./makroParser.js";

// Status labels carried on each merged row (single-status shape, like P2P).
export const MAKRO_STATUS = {
  MATCHED: "Price updated from Makro",
  NO_INFO: "No Information",
};

function r1145PriceOf(r) {
  const v = r?.priceOU;
  return typeof v === "number" && Number.isFinite(v) ? v : 0;
}

// Matched row: 1145 base with the Makro new price + lead time "1" overlaid.
function buildMatchedRow({ r1145Row, makroRow, pos }) {
  const oldPrice = r1145PriceOf(r1145Row);
  const newPrice = makroRow.newPrice;   // = ROUND(Makro L, 2)
  return {
    ...r1145Row,
    pos,
    priceOU:      newPrice,
    availability: "1",       // AF — matched ⇒ open lead time
    leadTimeRaw:  "1",       // keep in sync so validation reflects the XML output
    // The Makro price replaces the 1145 price entirely, so the 1145 parser's
    // scaled-price flags no longer describe this row — clear them to avoid
    // misleading "Scaled price…" warnings from the validator.
    __scaledPriceWasZero: false,
    __scaledPriceWasBlank: false,
    __priceBothBlank: false,
    // Drop the 1145 multi-pill status array so the single-status model (below)
    // is the one the modal filter / Excel export read.
    statuses: undefined,
    // Preview / full-data extras
    __oldPrice:   oldPrice,
    __newPrice:   newPrice,
    __diff:       oldPrice - newPrice,   // AC = old - new
    __makroName:      makroRow.itemName,
    __makroExVat:     makroRow.priceExVat,
    __makroVat:       makroRow.vatAmt,
    __makroInVat:     makroRow.inVat,
    __makroExVatAdj:  makroRow.priceExVatAdj,   // L (raw)
    __makroStatus:    makroRow.status,
    __makroArtGroup:  makroRow.artGroup,
    status:   MAKRO_STATUS.MATCHED,
    __source: "matched",
  };
}

// Unmatched 1145 row: keep its existing price, lead time "0", flag "No Information".
function buildNoInfoRow({ r1145Row, pos }) {
  const oldPrice = r1145PriceOf(r1145Row);
  return {
    ...r1145Row,
    pos,
    priceOU:      vbaRound(oldPrice, 2),   // VBA: AB = ROUND(Z, 2) when not found
    availability: "0",        // AF — not found ⇒ lead time 0
    leadTimeRaw:  "0",
    // Single-status model — drop the 1145 multi-pill array (see buildMatchedRow).
    statuses: undefined,
    __oldPrice:   oldPrice,
    __newPrice:   "",
    __diff:       "",
    __makroName:      "",
    __makroExVat:     "",
    __makroVat:       "",
    __makroInVat:     "",
    __makroExVatAdj:  "",
    __makroStatus:    "",
    __makroArtGroup:  "",
    status:   MAKRO_STATUS.NO_INFO,
    __source: "no-info",
  };
}

/**
 * @param {Array} r1145Rows  rows from parseReport1145(...)
 * @param {Array} makroRows  rows from parseMakroFile(...).rows
 * @returns {{rows: Array, summary: {total, matched, noInfo, makroOnly}}}
 */
export function mergeMakroAndReport1145(r1145Rows, makroRows) {
  // Makro dict keyed by product code (first occurrence wins).
  const makroByCode = new Map();
  for (const m of makroRows) {
    const key = String(m.artCode || "").trim();
    if (key && !makroByCode.has(key)) makroByCode.set(key, m);
  }

  const merged = [];
  const usedCodes = new Set();
  let pos = 1;

  // Drive by the Report 1145 list — every 1145 row appears in the output.
  for (const r of r1145Rows) {
    const key = String(r.itemNo || "").trim();
    const m = key ? makroByCode.get(key) : undefined;
    if (m) {
      merged.push(buildMatchedRow({ r1145Row: r, makroRow: m, pos: pos++ }));
      usedCodes.add(key);
    } else {
      merged.push(buildNoInfoRow({ r1145Row: r, pos: pos++ }));
    }
  }

  // Makro codes never matched to any 1145 article — informational only (these
  // do not appear in the output, matching the VBA's 1145-driven XLOOKUP).
  let makroOnly = 0;
  for (const key of makroByCode.keys()) {
    if (!usedCodes.has(key)) makroOnly++;
  }

  const summary = {
    total:    merged.length,
    matched:  merged.filter((m) => m.__source === "matched").length,
    noInfo:   merged.filter((m) => m.__source === "no-info").length,
    makroOnly,
  };

  return { rows: merged, summary };
}
