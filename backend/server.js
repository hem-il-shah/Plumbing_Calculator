/**
 * Plumbing Contractor Calculator - Backend API
 *
 * Data source: ./data/Plumbing_Calculator.xlsx
 *  - Sheet "CPVC uPVC WM "         -> Wall Mixer materials (HotX only / HotX+CoolX)
 *  - Sheet "CPVC uPVC SL Diverter" -> Single Lever Diverter materials
 *  - Sheet "SWR"                   -> Drainage (SWR) materials
 *
 * The workbook is parsed on startup into in-memory JSON. The C-Shape uplift
 * (+10%) is applied at quote time per BRD section 1 (Scope) and FR-01.
 */

const path = require("path");
const fs = require("fs");
const express = require("express");
const cors = require("cors");
const ExcelJS = require("exceljs");
const PDFDocument = require("pdfkit");

const XLSX_PATH = path.join(__dirname, "data", "Plumbing_Calculator.xlsx");
const LOGO_PATH = path.join(__dirname, "assets", "sintex-logo.jpg");
const BANNER_PATH = path.join(__dirname, "assets", "sintex-banner.jpg");
const PORT = process.env.PORT || 4000;

// ---------------------------------------------------------------------------
// Quotation ID generation
// ---------------------------------------------------------------------------
// Pattern: ddmmyyyyhhmmssSSS<IP-digits><XXXX>  (single continuous token, no separators)
//   ddmmyyyy   : date
//   hhmmssSSS  : hours, minutes, seconds, milliseconds (24-hr)
//   IP         : client device IP with non-alphanumerics stripped
//                (192.168.1.42 → 192168142)
//   XXXX       : 4 random alphanumeric characters (uppercase)

function clientIp(req) {
  // Prefer X-Forwarded-For (first hop) if behind a proxy, else socket
  const fwd = (req.headers["x-forwarded-for"] || "").split(",")[0].trim();
  let ip = fwd || req.socket?.remoteAddress || "0.0.0.0";
  // Normalize IPv6-mapped IPv4 like ::ffff:192.168.1.5 → 192.168.1.5
  if (ip.startsWith("::ffff:")) ip = ip.slice(7);
  if (ip === "::1") ip = "127.0.0.1";
  return ip;
}

function sanitizeIpForId(ip) {
  // Strip all non-alphanumerics (dots in IPv4, colons in IPv6) so the ID is a
  // single continuous token.
  return String(ip).replace(/[^A-Za-z0-9]/g, "");
}

function randomSuffix(n = 4) {
  const ALPHABET = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
  let s = "";
  for (let i = 0; i < n; i++) s += ALPHABET[Math.floor(Math.random() * ALPHABET.length)];
  return s;
}

function generateQuoteRef(ip) {
  const d = new Date();
  const pad = (n, w = 2) => String(n).padStart(w, "0");
  const datePart = pad(d.getDate()) + pad(d.getMonth() + 1) + d.getFullYear();
  const timePart = pad(d.getHours()) + pad(d.getMinutes()) + pad(d.getSeconds()) + pad(d.getMilliseconds(), 3);
  return `${datePart}${timePart}${sanitizeIpForId(ip)}${randomSuffix(4)}`;
}

// New product flow:
// System options now always include SWR (it's no longer standalone):
//   system:
//     - "CPVC + UPVC + SWR"  -> fixture sheet "Combination of HotX and CoolX" + SWR sheet
//     - "CPVC + SWR"         -> fixture sheet "Only HotX" + SWR sheet
//     - "UPVC + SWR"         -> fixture sheet "Combination of HotX and CoolX" with
//                                only UPVC-prefixed material groups + SWR sheet
//   pipingSystem (always required):
//     - "Single Lever Diverter / Flush Valve"
//     - "Wall Mixer"
//
// Inputs are now per-bathroom: the user supplies a list of bathrooms, each with
// its own shape, length, and width. Kitchens are taken as a count only (no
// dimensions); each kitchen is priced as one base 8 x 4 ft L-Shape BoQ.
//
// Pricing model (per bathroom)
//   base reference room: 8 ft (length) x 4 ft (width)
//   dim_uplift  = 1 + 0.015 * (L - 8) + 0.015 * (W - 4)
//   shape_mult  = 1.10 if "C Shape" else 1.00
//   per_room    = base_items * dim_uplift * shape_mult
//
// Grand total = sum(per_bathroom) + (kitchens * base_items_unscaled).
// There is no separate "bathrooms x kitchens" multiplier — rooms are added
// up explicitly.

const SYSTEMS = ["CPVC + UPVC + SWR", "CPVC + SWR", "UPVC + SWR"];
const PIPING_SYSTEMS = ["Single Lever Diverter / Flush Valve", "Wall Mixer"];

const BASE_LENGTH = 8; // ft (reference)
const BASE_WIDTH  = 4; // ft
const UPLIFT_PER_FT = 0.015; // 1.5% per ft per axis

const BATHROOM_SHAPES = ["L Shape", "C Shape"];
const C_SHAPE_UPLIFT = 1.10;

// Legacy constants kept so old buildBoq path (used by no current endpoint, but
// still referenced from a couple of tests) remains valid.
const FIXTURES = ["Wall Mixer", "Single Lever Diverter"];
const PIPE_SYSTEMS = ["HotX Only", "HotX + CoolX"];

function round2(n) {
  return Math.round(n * 100) / 100;
}

// ---------------------------------------------------------------------------
// Excel parsing
// ---------------------------------------------------------------------------

/**
 * The supply sheets have two stacked tables: "Only HotX" then
 * "Combination of HotX and CoolX". Each has a header row followed by
 * line items, terminated by a "Total" row. We scan rows, split on the
 * combination marker, and extract structured line items.
 */
/**
 * Normalize a cell value to a string. ExcelJS may return rich-text objects
 * `{ richText: [...] }`, formula objects `{ result: ..., formula: ... }`,
 * or hyperlink objects — we want the underlying text content for matching.
 */
function cellText(v) {
  if (v == null) return "";
  if (typeof v === "string") return v;
  if (typeof v === "number" || typeof v === "boolean") return String(v);
  if (typeof v === "object") {
    if (Array.isArray(v.richText)) return v.richText.map((r) => r.text || "").join("");
    if (v.text != null) return cellText(v.text);
    if (v.result != null) return String(v.result);
  }
  return "";
}

/** Numeric value from a cell, handling formula objects. */
function cellNum(v) {
  if (typeof v === "number") return v;
  if (v && typeof v === "object" && typeof v.result === "number") return v.result;
  return null;
}

function parseSupplySheet(ws, hasSlNo) {
  // hasSlNo: SL Diverter sheet has an extra SL:NO column shifting all fields right by 1.
  // Column layout (1-indexed):
  //   WM sheet (hasSlNo=false): B=code C=group D=desc E=sub F=type G=sizeIn H=sizeMm I=mrp J=qty K=value
  //   SL sheet (hasSlNo=true):  C=code D=group E=desc F=sub G=type H=sizeIn I=sizeMm J=mrp K=qty L=value
  const offset = hasSlNo ? 1 : 0;
  const col = {
    code:   2 + offset,
    group:  3 + offset,
    desc:   4 + offset,
    sub:    5 + offset,
    type:   6 + offset,
    sizeIn: 7 + offset,
    sizeMm: 8 + offset,
    mrp:    9 + offset,
    qty:    10 + offset,
    value:  11 + offset,
  };

  const sections = { "HotX Only": [], "HotX + CoolX": [] };
  let current = null;

  ws.eachRow({ includeEmpty: false }, (row) => {
    const cells = row.values; // 1-indexed
    const joined = cells.map(cellText).join(" ").toLowerCase();

    if (joined.includes("only hotx")) { current = "HotX Only"; return; }
    if (joined.includes("combination of hotx")) { current = "HotX + CoolX"; return; }
    if (current == null) return;

    const code = cellText(cells[col.code]).trim();
    const mrp = cellNum(cells[col.mrp]);
    const qty = cellNum(cells[col.qty]);

    if (!code) return;
    if (code.toLowerCase().includes("material code")) return;
    if (code.toLowerCase().includes("total")) return;
    if (mrp == null || qty == null) return;

    sections[current].push({
      materialCode: code,
      materialGroup: cellText(cells[col.group]).trim(),
      description: cellText(cells[col.desc]).trim(),
      subCategory: cellText(cells[col.sub]).trim(),
      pipeType: cellText(cells[col.type]).trim(),
      sizeInch: cellText(cells[col.sizeIn]).trim(),
      sizeMm: cellText(cells[col.sizeMm]).trim(),
      mrp,
      qty,
      assessableValue: round2(mrp * qty),
    });
  });

  return sections;
}

function parseSwrSheet(ws) {
  // SWR layout: A=SL B=code C=group D=desc E=sub F=type G=sizeIn H=sizeMm I=mrp J=qty K=value
  const items = [];
  ws.eachRow({ includeEmpty: false }, (row) => {
    const cells = row.values;
    const code = cellText(cells[2]).trim();
    const mrp = cellNum(cells[9]);
    const qty = cellNum(cells[10]);

    if (!code) return;
    if (code.toLowerCase().includes("material code")) return;
    if (code.toLowerCase().includes("total")) return;
    if (code.toLowerCase().includes("drainx")) return;
    if (mrp == null || qty == null) return;

    items.push({
      materialCode: code,
      materialGroup: cellText(cells[3]).trim(),
      description: cellText(cells[4]).trim(),
      subCategory: cellText(cells[5]).trim(),
      pipeType: cellText(cells[6]).trim(),
      sizeInch: cellText(cells[7]).trim(),
      sizeMm: cellText(cells[8]).trim(),
      mrp,
      qty,
      assessableValue: round2(mrp * qty),
    });
  });
  return items;
}

let MASTER = null;

async function loadMaster() {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(XLSX_PATH);

  const wmSheet = wb.getWorksheet("CPVC uPVC WM ");
  const slSheet = wb.getWorksheet("CPVC uPVC SL Diverter ");
  const swrSheet = wb.getWorksheet("SWR");

  if (!wmSheet || !slSheet || !swrSheet) {
    throw new Error("Missing expected sheet(s) in workbook.");
  }

  MASTER = {
    "Wall Mixer": parseSupplySheet(wmSheet, false),
    "Single Lever Diverter": parseSupplySheet(slSheet, true),
    SWR: parseSwrSheet(swrSheet),
  };

  console.log("[data] Master loaded:");
  console.log(`  Wall Mixer / HotX Only:        ${MASTER["Wall Mixer"]["HotX Only"].length} items`);
  console.log(`  Wall Mixer / HotX + CoolX:     ${MASTER["Wall Mixer"]["HotX + CoolX"].length} items`);
  console.log(`  SL Diverter / HotX Only:       ${MASTER["Single Lever Diverter"]["HotX Only"].length} items`);
  console.log(`  SL Diverter / HotX + CoolX:    ${MASTER["Single Lever Diverter"]["HotX + CoolX"].length} items`);
  console.log(`  SWR:                           ${MASTER.SWR.length} items`);
}

// ---------------------------------------------------------------------------
// BOQ generation
// ---------------------------------------------------------------------------

function applyUplift(items, multiplier) {
  return items.map((it) => {
    const qty = round2(it.qty * multiplier);
    return {
      ...it,
      qty,
      assessableValue: round2(it.mrp * qty),
    };
  });
}

function parseDim(v, name) {
  const n = typeof v === "number" ? v : parseFloat(v);
  if (!isFinite(n) || n <= 0) throw new Error(`Invalid ${name}: must be a positive number (in feet)`);
  return round2(n);
}

function buildBoq(body = {}) {
  const {
    shape, fixture, pipeSystem,
    length, width, height,
    includeSwr = true,
  } = body;

  if (!BATHROOM_SHAPES.includes(shape)) throw new Error(`Invalid shape: ${shape}`);
  if (!FIXTURES.includes(fixture)) throw new Error(`Invalid fixture: ${fixture}`);
  if (!PIPE_SYSTEMS.includes(pipeSystem)) throw new Error(`Invalid pipe system: ${pipeSystem}`);

  const L = parseDim(length, "length");
  const W = parseDim(width, "width");
  const H = parseDim(height, "height");

  const multiplier = shape === "C Shape" ? C_SHAPE_UPLIFT : 1.0;
  const supply = applyUplift(MASTER[fixture][pipeSystem], multiplier);
  const supplyTotal = round2(supply.reduce((s, x) => s + x.assessableValue, 0));

  let swr = null;
  let swrTotal = 0;
  if (includeSwr) {
    const swrItems = applyUplift(MASTER.SWR, multiplier);
    swrTotal = round2(swrItems.reduce((s, x) => s + x.assessableValue, 0));
    swr = { items: swrItems, total: swrTotal, label: "SWR Drainage (DrainX)" };
  }

  const grandTotal = round2(supplyTotal + swrTotal);

  return {
    selection: {
      shape, fixture, pipeSystem,
      length: L, width: W, height: H,
      dimensionsLabel: `${L} × ${W} × ${H} ft`,
      includeSwr: !!includeSwr,
    },
    multiplier,
    supply: { items: supply, total: supplyTotal, label: `${fixture} – ${pipeSystem}` },
    swr,
    grandTotal,
    generatedAt: new Date().toISOString(),
  };
}

// ---------------------------------------------------------------------------
// New flow: buildQuote
// ---------------------------------------------------------------------------

/** Canonical group image filenames (kept as the actual files on disk). */
const GROUP_IMAGE_MAP = {
  "CPVC FITTINGS":           "CPVC_Fittings.jpg",
  "CPVC BRASS FITTINGS":     "CPVC_Brass_Fittings.jpg",
  "CPVC PIPES":              "CPVC_Pipes.jpg",
  "CPVC SOLVENT":            "CPVC_Solvent.jpg",
  "UPVC FITTINGS":           "UPVC_Fittings.jpg",
  "UPVC FITTINGS-BRASS":     "UPVC_Fittings_Brass.jpg",
  "UPVC PIPES":              "UPVC_Pipes.jpg",
  "UPVC SOLVENT":            "UPVC_Solvent.jpeg",
  "SWR FITTINGS":            "SWR_Fittings.jpg",
  "SWR PIPES":               "SWR_Pipes.jpg",
  "SWR PIPES 3L":            "SWR_Pipes_3l.jpg",
  "SWR SOLVENT":             "SWR_Solvent.jpg",
  "SURFACE DRAINAGE":        "Surface_Drainage.jpg",
  "AGRI FITTINGS":           "AGRI_Fittings.jpg",
  "AGRI PIPES":              "AGRP_Pipes.jpg",
  "RECLAIM PIPES":           "Reclaim_Pipes.jpg",
  "FOAMCORE FITTINGS (UGD)": "Foamcore_Fittings_UGD.png",
  "FOAMCORE PIPES (UGD)":    "Foamcore_Pipes.jpg",
};

/** Normalize a group name (the workbook uses inconsistent case). */
function normalizeGroup(name) {
  return String(name || "").trim().toUpperCase().replace(/\s+/g, " ");
}

function imageForGroup(groupName) {
  return GROUP_IMAGE_MAP[normalizeGroup(groupName)] || null;
}

function parsePositiveInt(v, name) {
  const n = Number(v);
  if (!Number.isInteger(n) || n < 1) throw new Error(`Invalid ${name}: must be an integer >= 1`);
  return n;
}

function parsePositiveFloat(v, name) {
  const n = Number(v);
  if (!Number.isFinite(n) || n <= 0) throw new Error(`Invalid ${name}: must be a positive number (in feet)`);
  return round2(n);
}

/**
 * Build the new product quote.
 *
 * body = {
 *   system: "CPVC" | "CPVC + UPVC" | "SWR",
 *   bathrooms: int >= 1,
 *   kitchens: int >= 1                 (only for CPVC / CPVC+UPVC)
 *   floors:   int >= 1                 (only for SWR)
 *   length, width: positive numbers in ft (base reference 8 x 4)
 *   pipingSystem: "Single Lever Diverter / Flush Valve" | "Wall Mixer"
 *                                       (only for CPVC / CPVC+UPVC)
 * }
 */
/**
 * Resolve which line items to use for a given system + piping system.
 *
 *   "CPVC + UPVC + SWR" → Combination (HotX+CoolX) section + SWR sheet
 *   "CPVC + SWR"        → Only HotX section + SWR sheet
 *   "UPVC + SWR"        → Combination section filtered to UPVC-only groups + SWR sheet
 *
 * Returns the base item array (untouched, full qty/MRP from the workbook).
 */
function resolveBaseItems(system, pipingSystem) {
  const fixtureKey = pipingSystem === "Wall Mixer" ? "Wall Mixer" : "Single Lever Diverter";

  let fixtureItems = [];
  if (system === "CPVC + SWR") {
    fixtureItems = MASTER[fixtureKey]["HotX Only"] || [];
  } else if (system === "CPVC + UPVC + SWR") {
    fixtureItems = MASTER[fixtureKey]["HotX + CoolX"] || [];
  } else if (system === "UPVC + SWR") {
    // Take the combination section and strip out anything whose Material Group
    // does NOT contain "UPVC". So you're left with UPVC fittings/pipes/solvent
    // (the CPVC half of the dual-pipe kit is removed).
    const combo = MASTER[fixtureKey]["HotX + CoolX"] || [];
    fixtureItems = combo.filter((it) =>
      /UPVC/i.test(String(it.materialGroup || ""))
    );
  } else {
    throw new Error(`Unknown system: ${system}`);
  }

  // SWR is included in every system option.
  const swrItems = MASTER.SWR || [];
  return [...fixtureItems, ...swrItems];
}

/** Scale a list of base items by a single overall multiplier. */
function scaleItems(baseItems, mult) {
  return baseItems.map((it) => {
    const scaledQty = it.qty * mult;
    return {
      ...it,
      qty: round2(scaledQty),
      assessableValue: round2(scaledQty * it.mrp),
    };
  });
}

/** Group an item list by Material Group. */
function groupItems(items) {
  const map = new Map();
  for (const it of items) {
    const key = normalizeGroup(it.materialGroup);
    if (!map.has(key)) {
      map.set(key, {
        groupKey: key,
        groupName: it.materialGroup,
        image: imageForGroup(it.materialGroup),
        items: [],
        subtotal: 0,
      });
    }
    const g = map.get(key);
    g.items.push(it);
    g.subtotal = round2(g.subtotal + it.assessableValue);
  }
  return [...map.values()].filter((g) => g.subtotal > 0);
}

/** Compute the per-bathroom uplift (dimension * shape) from raw inputs. */
function computeRoomMultiplier(length, width, shape) {
  const lengthDelta = length - BASE_LENGTH;
  const widthDelta  = width  - BASE_WIDTH;
  const dimensionPct = (lengthDelta + widthDelta) * UPLIFT_PER_FT;
  const dimensionMult = Math.max(0.1, 1 + dimensionPct);
  const shapeMult = shape === "C Shape" ? C_SHAPE_UPLIFT : 1.0;
  return {
    dimensionPct,
    dimensionMult,
    shapeMult,
    overall: dimensionMult * shapeMult,
  };
}

/**
 * Build a quote under the new per-bathroom model.
 *
 * body = {
 *   system: one of SYSTEMS,
 *   pipingSystem: one of PIPING_SYSTEMS,
 *   bathrooms: [ { shape, length, width }, ... ]   -- one entry per bathroom
 *   kitchens: integer >= 1                         -- kitchens have no dims
 * }
 */
function buildQuote(body = {}) {
  const system = String(body.system || "").trim();
  if (!SYSTEMS.includes(system)) {
    throw new Error(`Invalid system: ${system}. Must be one of: ${SYSTEMS.join(", ")}`);
  }

  const pipingSystem = String(body.pipingSystem || "").trim();
  if (!PIPING_SYSTEMS.includes(pipingSystem)) {
    throw new Error(`Invalid pipingSystem: ${pipingSystem}. Must be one of: ${PIPING_SYSTEMS.join(", ")}`);
  }

  if (!Array.isArray(body.bathrooms) || body.bathrooms.length < 1) {
    throw new Error("Provide at least one bathroom.");
  }

  const bathroomsInput = body.bathrooms.map((b, idx) => {
    const shape = String(b.shape || "").trim();
    if (!BATHROOM_SHAPES.includes(shape)) {
      throw new Error(`Bathroom #${idx + 1}: invalid shape "${b.shape}". Must be one of: ${BATHROOM_SHAPES.join(", ")}`);
    }
    return {
      shape,
      length: parsePositiveFloat(b.length, `bathroom #${idx + 1} length`),
      width:  parsePositiveFloat(b.width,  `bathroom #${idx + 1} width`),
    };
  });

  const kitchens = parsePositiveInt(body.kitchens, "kitchens");

  const baseItems = resolveBaseItems(system, pipingSystem);
  if (!baseItems.length) throw new Error("No items for the selected combination");

  // --- Build per-bathroom sections (one per input bathroom) -----------------
  const bathroomSections = bathroomsInput.map((b, idx) => {
    const m = computeRoomMultiplier(b.length, b.width, b.shape);
    const scaled = scaleItems(baseItems, m.overall);
    const groups = groupItems(scaled);
    const subtotal = round2(groups.reduce((s, g) => s + g.subtotal, 0));
    return {
      index: idx + 1,
      shape: b.shape,
      length: b.length,
      width: b.width,
      dimensionsLabel: `${b.length} × ${b.width} ft`,
      multiplier: round2(m.overall * 1000) / 1000,
      dimensionPct: round2(m.dimensionPct * 100) / 100,
      shapeUpliftPct: b.shape === "C Shape" ? Math.round((C_SHAPE_UPLIFT - 1) * 100) : 0,
      groups,
      subtotal,
    };
  });

  // --- Kitchens section: one base BoQ per kitchen, no uplift ----------------
  // We model it as a single section with quantity = (base_qty * kitchens) so
  // the displayed line items represent the combined kitchen material list.
  const kitchenScaled = scaleItems(baseItems, kitchens);
  const kitchenGroups = groupItems(kitchenScaled);
  const kitchenSubtotal = round2(kitchenGroups.reduce((s, g) => s + g.subtotal, 0));
  // The cost of ONE kitchen (used for the per-cost summary). All base items at
  // multiplier 1 → just sum mrp * qty across the base list.
  const perKitchenCost = round2(baseItems.reduce((s, it) => s + it.mrp * it.qty, 0));

  const kitchenSection = {
    count: kitchens,
    perCost: perKitchenCost,
    groups: kitchenGroups,
    subtotal: kitchenSubtotal,
  };

  // --- Grand total = sum of bathroom subtotals + kitchen subtotal -----------
  const grandTotal = round2(
    bathroomSections.reduce((s, b) => s + b.subtotal, 0) + kitchenSubtotal
  );

  // --- Summary rows for the second PDF --------------------------------------
  // Bathrooms: group by (shape, length, width). Same dims → one row, count++.
  const bathroomSummaryMap = new Map();
  for (const b of bathroomSections) {
    const key = `${b.shape}|${b.length}|${b.width}`;
    if (!bathroomSummaryMap.has(key)) {
      bathroomSummaryMap.set(key, {
        kind: "Bathroom",
        shape: b.shape,
        length: b.length,
        width: b.width,
        dimensionsLabel: b.dimensionsLabel,
        count: 0,
        perCost: b.subtotal, // every bathroom in this group has the same subtotal by construction
      });
    }
    bathroomSummaryMap.get(key).count += 1;
  }
  const summaryRows = [...bathroomSummaryMap.values()].map((r) => ({
    ...r,
    total: round2(r.perCost * r.count),
  }));
  // Kitchens: single row
  summaryRows.push({
    kind: "Kitchen",
    count: kitchens,
    perCost: perKitchenCost,
    total: round2(perKitchenCost * kitchens),
  });

  return {
    selection: {
      system,
      pipingSystem,
      kitchens,
      bathrooms: bathroomsInput,
    },
    bathroomSections,
    kitchenSection,
    summaryRows,
    grandTotal,
    generatedAt: new Date().toISOString(),
  };
}

/**
 * Build a PDF quotation matching the on-screen layout: header, configuration
 * block, section subtotals, then the full itemized BOQ (Supply + optional SWR)
 * with every column shown in the UI.
 *
 * PDFKit's bundled Helvetica is WinAnsi-encoded and doesn't include the
 * Rupee glyph (₹, U+20B9); we use the "Rs." prefix and the "×" multiplication
 * sign which IS supported. This keeps the PDF font-embed-free and portable.
 */
function exportBoqPdf(boq, quoteRef) {
  return new Promise((resolve, reject) => {
    const doc = new PDFDocument({
      size: "A4",
      layout: "portrait",
      margins: { top: 96, bottom: 64, left: 36, right: 36 },
      bufferPages: true,
      info: {
        Title: `Plumbing Quotation - ${boq.selection.shape}`,
        Author: "Sintex by Welspun",
        Subject: "Plumbing Contractor Quotation",
      },
    });
    const chunks = [];
    doc.on("data", (c) => chunks.push(c));
    doc.on("end", () => resolve(Buffer.concat(chunks)));
    doc.on("error", reject);

    const PAGE_W = doc.page.width;
    const PAGE_H = doc.page.height;
    const M = doc.page.margins.left;
    const RIGHT = PAGE_W - doc.page.margins.right;
    const CONTENT_W = RIGHT - M;
    const TOP = doc.page.margins.top;
    const BOTTOM = PAGE_H - doc.page.margins.bottom;

    const COLOR = {
      brand: "#c8102e",        // Sintex red (primary)
      brandDark: "#7f0a1d",    // deep maroon for emphasis
      accent: "#f59e0b",       // warm gold accent stripe (complements red)
      ink: "#1f2937",
      muted: "#6b7280",
      border: "#e5d0d3",       // soft red-tinted border
      rowAlt: "#fff5f6",       // very light pink for alt rows
      headBg: "#c8102e",       // table header red
      headText: "#ffffff",
      totalBg: "#fef3c7",      // soft gold for grand total band
      totalText: "#7f0a1d",
      headerBg: "#7f0a1d",     // deep maroon header band
      footerBg: "#fff5f6",
    };

    const money = (n) => "Rs. " + Number(n || 0).toLocaleString("en-IN", { maximumFractionDigits: 2 });

    /**
     * Sanitize text for PDFKit's WinAnsi-encoded Helvetica. Several characters
     * present in the source workbook (ring-above ˚, curly quotes, en/em dashes,
     * etc.) are not in WinAnsi and would render as garbage. Map them to safe
     * ASCII equivalents. We keep the Latin-1 fractions (¾ ½ ¼) — those ARE in
     * WinAnsi and render fine.
     */
    const safe = (s) => {
      if (s == null) return "";
      return String(s)
        .replace(/\u02DA/g, "\u00B0")  // ring above ˚  → degree °
        .replace(/[\u2018\u2019]/g, "'")  // curly singles → '
        .replace(/[\u201C\u201D]/g, '"')  // curly doubles → "
        .replace(/[\u2013\u2014]/g, "-")  // en/em dash   → -
        .replace(/\u2026/g, "...")        // ellipsis
        .replace(/\u20B9/g, "Rs.")        // rupee sign   → Rs.
        .replace(/\u00A0/g, " ");         // nbsp         → space
    };

    // ---- Header / Footer drawn on every page ----
    const HEADER_H = 76; // band height; top margin is 96 so 20pt gap before content
    const FOOTER_H = 36;
    const generatedAt = new Date(boq.generatedAt);
    const generatedLabel = generatedAt.toLocaleString("en-IN", {
      day: "2-digit", month: "short", year: "numeric",
      hour: "2-digit", minute: "2-digit", hour12: true,
    });

    const drawHeader = () => {
      // Brand band (deep maroon)
      doc.rect(0, 0, PAGE_W, HEADER_H).fillColor(COLOR.headerBg).fill();
      // Gold accent stripe under the band
      doc.rect(0, HEADER_H, PAGE_W, 3).fillColor(COLOR.accent).fill();

      // ---- Logo card (top-left) ----
      // The source logo has a white background, so we place it on a small white
      // rounded card so it sits cleanly on the maroon band.
      const logoBoxW = 96;
      const logoBoxH = 50;
      const logoX = M;
      const logoY = (HEADER_H - logoBoxH) / 2;
      try {
        doc.roundedRect(logoX, logoY, logoBoxW, logoBoxH, 4).fillColor("#ffffff").fill();
        // Fit logo inside the card with 4pt padding
        doc.image(LOGO_PATH, logoX + 4, logoY + 4, {
          fit: [logoBoxW - 8, logoBoxH - 8],
          align: "center",
          valign: "center",
        });
      } catch (e) {
        // If logo isn't present for some reason, fall back to text-only branding.
      }

      // ---- Company / tag text (next to logo) ----
      const textX = logoX + logoBoxW + 14;
      doc.fillColor("#ffffff").font("Helvetica-Bold").fontSize(18)
        .text(safe("Sintex by Welspun"), textX, 22, { lineBreak: false });
      doc.fillColor("#fde2e6").font("Helvetica").fontSize(9)
        .text("PLUMBING CONTRACTOR QUOTATION", textX, 48, { lineBreak: false });

      // ---- Right side: Quotation ID + Date ----
      doc.fillColor("#fde2e6").font("Helvetica").fontSize(8)
        .text("QUOTATION ID", 0, 14, { width: CONTENT_W, align: "right", lineBreak: false });
      doc.fillColor("#ffffff").font("Helvetica-Bold").fontSize(10)
        .text(safe(quoteRef), 0, 26, { width: CONTENT_W, align: "right", lineBreak: false });
      doc.fillColor("#fde2e6").font("Helvetica").fontSize(8)
        .text("DATE", 0, 44, { width: CONTENT_W, align: "right", lineBreak: false });
      doc.fillColor("#ffffff").font("Helvetica-Bold").fontSize(9)
        .text(generatedLabel, 0, 55, { width: CONTENT_W, align: "right", lineBreak: false });
    };

    const drawFooter = (pageNum, pageCount) => {
      // Temporarily relax bottom margin so text() doesn't auto-paginate when
      // writing into the footer band (which sits below the normal text area).
      const savedBottom = doc.page.margins.bottom;
      doc.page.margins.bottom = 0;
      const fy = PAGE_H - FOOTER_H;
      // Top border line
      doc.strokeColor(COLOR.brand).lineWidth(1.2)
        .moveTo(M, fy).lineTo(RIGHT, fy).stroke();
      // Line 1: brand mark (left) + page indicator (right) — same baseline
      doc.fillColor(COLOR.muted).font("Helvetica").fontSize(8)
        .text("Sintex by Welspun  •  Plumbing Contractor Calculator", M, fy + 8, {
          width: CONTENT_W / 2, lineBreak: false,
        });
      doc.fillColor(COLOR.ink).font("Helvetica-Bold").fontSize(8)
        .text(`Page ${pageNum} of ${pageCount}`, M + CONTENT_W / 2, fy + 8, {
          width: CONTENT_W / 2, align: "right", lineBreak: false,
        });
      // Line 2: quote ref (full width, smaller)
      doc.fillColor(COLOR.muted).font("Helvetica").fontSize(7)
        .text(`Quotation ID: ${safe(quoteRef)}`, M, fy + 22, {
          width: CONTENT_W, align: "center", lineBreak: false,
        });
      doc.page.margins.bottom = savedBottom;
    };

    // Draw header for the first page now; remaining pages handled via pageAdded.
    drawHeader();
    doc.on("pageAdded", () => {
      drawHeader();
    });

    // y cursor starts just below the header band
    let yStart = HEADER_H + 18;
    // ---- Configuration block ----
    let y = yStart;
    doc.fillColor(COLOR.brandDark).font("Helvetica-Bold").fontSize(12)
      .text("Configuration", M, y);
    y = doc.y + 6;
    const config = [
      ["Bathroom Shape", boq.selection.shape],
      ["Dimensions (L x W x H)", `${boq.selection.length} x ${boq.selection.width} x ${boq.selection.height} ft`],
      ["Fixture Type", boq.selection.fixture],
      ["Pipe System", boq.selection.pipeSystem],
      ["SWR Drainage", boq.selection.includeSwr ? "Included" : "Not included"],
      ["Quantity Multiplier", "x" + boq.multiplier],
    ];
    const cfgCols = 3;
    const cfgColW = CONTENT_W / cfgCols;
    const cfgRowH = 30;
    doc.rect(M, y, CONTENT_W, cfgRowH * Math.ceil(config.length / cfgCols))
      .fillColor(COLOR.rowAlt).fill();
    config.forEach((entry, i) => {
      const col = i % cfgCols;
      const row = Math.floor(i / cfgCols);
      const cx = M + col * cfgColW + 10;
      const cy = y + row * cfgRowH + 6;
      doc.fillColor(COLOR.muted).font("Helvetica").fontSize(7).text(entry[0].toUpperCase(), cx, cy);
      doc.fillColor(COLOR.ink).font("Helvetica-Bold").fontSize(10).text(safe(String(entry[1])), cx, cy + 9);
    });
    y += cfgRowH * Math.ceil(config.length / cfgCols) + 14;

    // ---- Section subtotals ----
    doc.fillColor(COLOR.brandDark).font("Helvetica-Bold").fontSize(11).text("Summary", M, y);
    y = doc.y + 6;

    const sumCol1 = M;
    const sumCol2 = RIGHT - 130;
    const sumColW1 = sumCol2 - sumCol1 - 10;

    // header
    doc.rect(M, y, CONTENT_W, 22).fillColor(COLOR.headBg).fill();
    doc.fillColor(COLOR.headText).font("Helvetica-Bold").fontSize(9);
    doc.text("DESCRIPTION", sumCol1 + 8, y + 7);
    doc.text("AMOUNT", sumCol2, y + 7, { width: 130 - 8, align: "right" });
    y += 22;

    const writeSummaryRow = (label, sub, amount) => {
      const h = 24;
      doc.rect(M, y, CONTENT_W, h).fillColor("#ffffff").fill().strokeColor(COLOR.border).lineWidth(0.5);
      doc.moveTo(M, y + h).lineTo(RIGHT, y + h).stroke();
      doc.fillColor(COLOR.ink).font("Helvetica").fontSize(10).text(safe(label), sumCol1 + 8, y + 5);
      doc.fillColor(COLOR.muted).fontSize(7).text(sub, sumCol1 + 8, y + 16);
      doc.fillColor(COLOR.ink).font("Helvetica").fontSize(10)
        .text(money(amount), sumCol2, y + 9, { width: 130 - 8, align: "right" });
      y += h;
    };
    writeSummaryRow(`Supply piping - ${boq.supply.label}`, `${boq.supply.items.length} line items`, boq.supply.total);
    if (boq.swr) {
      writeSummaryRow(boq.swr.label, `${boq.swr.items.length} line items`, boq.swr.total);
    }

    // Grand Total
    const gtH = 28;
    doc.rect(M, y, CONTENT_W, gtH).fillColor(COLOR.totalBg).fill();
    doc.fillColor(COLOR.totalText).font("Helvetica-Bold").fontSize(12);
    doc.text("Grand Total", sumCol1 + 8, y + 9);
    doc.text(money(boq.grandTotal), sumCol2, y + 9, { width: 130 - 8, align: "right" });
    y += gtH + 16;

    // ---- Itemized BOQ ----
    // Columns: Sr, Code, Group, Description, Sub Category, Type, Size(in), Size(mm), MRP, Qty, Amount
    const COL_WIDTHS = [22, 78, 70, 142, 80, 42, 38, 38, 50, 28, 60]; // auto-scaled below to fit portrait A4 content width
    const totalColW = COL_WIDTHS.reduce((s, w) => s + w, 0);
    const colScale = CONTENT_W / totalColW;
    const W = COL_WIDTHS.map((w) => w * colScale);
    const HEADERS = ["Sr.", "Material Code", "Group", "Description", "Sub Category", "Type", "Size in", "Size mm", "MRP", "Qty", "Amount"];
    const NUM_COLS = new Set([0, 8, 9, 10]);

    const writeTableHeader = () => {
      doc.rect(M, y, CONTENT_W, 18).fillColor(COLOR.headBg).fill();
      doc.fillColor(COLOR.headText).font("Helvetica-Bold").fontSize(7.5);
      let x = M;
      HEADERS.forEach((h, i) => {
        doc.text(h, x + 4, y + 6, {
          width: W[i] - 8,
          align: NUM_COLS.has(i) ? "right" : "left",
          lineBreak: false,
        });
        x += W[i];
      });
      y += 18;
    };

    const writeTableRow = (cells, alt, bold = false) => {
      const safeCells = cells.map((c) => safe(c));
      // measure row height
      doc.font(bold ? "Helvetica-Bold" : "Helvetica").fontSize(7);
      let rowH = 0;
      safeCells.forEach((c, i) => {
        const h = doc.heightOfString(c, { width: W[i] - 8 });
        if (h > rowH) rowH = h;
      });
      rowH = Math.max(rowH + 6, 14);

      // page break if needed
      if (y + rowH > doc.page.height - doc.page.margins.bottom) {
        doc.addPage();
        y = doc.page.margins.top;
        writeTableHeader();
      }

      if (alt) {
        doc.rect(M, y, CONTENT_W, rowH).fillColor(COLOR.rowAlt).fill();
      }
      doc.strokeColor(COLOR.border).lineWidth(0.3)
        .moveTo(M, y + rowH).lineTo(RIGHT, y + rowH).stroke();

      doc.fillColor(COLOR.ink).font(bold ? "Helvetica-Bold" : "Helvetica").fontSize(7);
      let x = M;
      safeCells.forEach((c, i) => {
        doc.text(c, x + 4, y + 3, {
          width: W[i] - 8,
          align: NUM_COLS.has(i) ? "right" : "left",
        });
        x += W[i];
      });
      y += rowH;
    };

    const writeSubtotalRow = (label, total) => {
      const rowH = 18;
      if (y + rowH > doc.page.height - doc.page.margins.bottom) {
        doc.addPage();
        y = doc.page.margins.top;
        writeTableHeader();
      }
      doc.rect(M, y, CONTENT_W, rowH).fillColor("#f3f4f6").fill();
      doc.strokeColor(COLOR.brand).lineWidth(0.8)
        .moveTo(M, y).lineTo(RIGHT, y).stroke();
      doc.strokeColor(COLOR.brand).lineWidth(0.8)
        .moveTo(M, y + rowH).lineTo(RIGHT, y + rowH).stroke();

      // last column gets the amount; everything to the left of it gets the label
      const amountW = W[W.length - 1];
      const labelW = CONTENT_W - amountW;
      doc.fillColor(COLOR.brandDark).font("Helvetica-Bold").fontSize(8.5);
      doc.text(label, M + 4, y + 5, { width: labelW - 8, align: "right" });
      doc.text(safe(money(total)), M + labelW, y + 5, { width: amountW - 4, align: "right" });
      y += rowH;
    };

    const writeSection = (heading, items, total, totalLabel) => {
      // section heading
      if (y + 40 > doc.page.height - doc.page.margins.bottom) {
        doc.addPage();
        y = doc.page.margins.top;
      }
      doc.fillColor(COLOR.brandDark).font("Helvetica-Bold").fontSize(11).text(safe(heading), M, y);
      y = doc.y + 6;

      writeTableHeader();
      items.forEach((it, i) => {
        writeTableRow([
          i + 1,
          it.materialCode,
          it.materialGroup,
          it.description,
          it.subCategory,
          it.pipeType,
          it.sizeInch,
          it.sizeMm,
          money(it.mrp),
          it.qty,
          money(it.assessableValue),
        ], i % 2 === 1);
      });
      writeSubtotalRow(`${totalLabel} Subtotal`, total);
      y += 10;
    };

    doc.fillColor(COLOR.brandDark).font("Helvetica-Bold").fontSize(13)
      .text("Itemized Bill of Quantities", M, y);
    y = doc.y + 8;

    writeSection(`Supply BOQ - ${boq.supply.label}`, boq.supply.items, boq.supply.total, "Supply");
    if (boq.swr) {
      writeSection(`SWR Drainage - ${boq.swr.label}`, boq.swr.items, boq.swr.total, "SWR");
    }

    // ---- Footer note (closing paragraph; the page footer band is drawn per-page below) ----
    if (y + 60 > BOTTOM) {
      doc.addPage();
      y = HEADER_H + 18;
    } else {
      y += 8;
    }
    doc.strokeColor(COLOR.border).lineWidth(0.5).moveTo(M, y).lineTo(RIGHT, y).stroke();
    y += 8;
    doc.fillColor(COLOR.ink).font("Helvetica").fontSize(8).text(
      "Note: This is a tentative quotation. Prices are based on current SAP MRP reference data and are subject to revision. " +
      "All amounts are in Indian Rupees (Rs.), computed as unit MRP x quantity (assessable value).",
      M, y, { width: CONTENT_W }
    );

    // ---- Stamp footers on every page now that total count is known ----
    const range = doc.bufferedPageRange();
    for (let i = 0; i < range.count; i++) {
      doc.switchToPage(range.start + i);
      drawFooter(i + 1, range.count);
    }

    doc.end();
  });
}

// ---------------------------------------------------------------------------
// PDF export — shared helpers (header, footer, configuration grid)
// ---------------------------------------------------------------------------

const PDF_COLOR = {
  brand:     "#c8102e",
  brandDark: "#7f0a1d",
  accent:    "#f59e0b",
  ink:       "#1f2937",
  muted:     "#6b7280",
  border:    "#e5d0d3",
  rowAlt:    "#fff5f6",
  headBg:    "#c8102e",
  headText:  "#ffffff",
  totalBg:   "#fef3c7",
  totalText: "#7f0a1d",
  headerBg:  "#7f0a1d",
};

function pdfSafe(s) {
  if (s == null) return "";
  return String(s)
    .replace(/\u02DA/g, "\u00B0")    // ring above ˚ → degree °
    .replace(/[\u2018\u2019]/g, "'")  // curly singles
    .replace(/[\u201C\u201D]/g, '"')  // curly doubles
    .replace(/[\u2013\u2014]/g, "-")  // dashes
    .replace(/\u2026/g, "...")
    .replace(/\u20B9/g, "Rs. ")
    .replace(/\u00A0/g, " ")
    .replace(/\u00BC/g, "1/4")        // ¼
    .replace(/\u00BD/g, "1/2")        // ½
    .replace(/\u00BE/g, "3/4");       // ¾
}

const pdfInr = (n) => "Rs. " + Number(n).toLocaleString("en-IN", { maximumFractionDigits: 2 });

/**
 * Initialize a new PDFKit document for a Sintex BOQ/Summary export.
 * Returns { doc, geom, drawHeader, drawFooter, stampFooters } so the caller
 * just renders the body and calls stampFooters() at the end.
 */
function initPdf({ headerSubtitle, quoteRef, generatedAt, info = {} }) {
  const doc = new PDFDocument({
    size: "A4",
    layout: "portrait",
    margins: { top: 96, bottom: 64, left: 36, right: 36 },
    bufferPages: true,
    info: {
      Author: "Sintex by Welspun",
      ...info,
    },
  });

  const PAGE_W = doc.page.width;
  const PAGE_H = doc.page.height;
  const M = doc.page.margins.left;
  const RIGHT = PAGE_W - doc.page.margins.right;
  const CONTENT_W = RIGHT - M;
  const BOTTOM = PAGE_H - doc.page.margins.bottom;
  const HEADER_H = 76;
  const FOOTER_H = 36;

  const genAt = new Date(generatedAt || Date.now());
  const generatedLabel = genAt.toLocaleString("en-IN", {
    day: "2-digit", month: "short", year: "numeric",
    hour: "2-digit", minute: "2-digit", hour12: true,
  });

  const drawHeader = () => {
    doc.rect(0, 0, PAGE_W, HEADER_H).fillColor(PDF_COLOR.headerBg).fill();
    doc.rect(0, HEADER_H, PAGE_W, 3).fillColor(PDF_COLOR.accent).fill();

    const logoBoxW = 96, logoBoxH = 50;
    const logoX = M, logoY = (HEADER_H - logoBoxH) / 2;
    try {
      doc.roundedRect(logoX, logoY, logoBoxW, logoBoxH, 4).fillColor("#ffffff").fill();
      doc.image(LOGO_PATH, logoX + 4, logoY + 4, {
        fit: [logoBoxW - 8, logoBoxH - 8], align: "center", valign: "center",
      });
    } catch (e) { /* ignore */ }

    const textX = logoX + logoBoxW + 14;
    doc.fillColor("#ffffff").font("Helvetica-Bold").fontSize(18)
      .text(pdfSafe("Sintex by Welspun"), textX, 22, { lineBreak: false });
    doc.fillColor("#fde2e6").font("Helvetica").fontSize(9)
      .text(headerSubtitle, textX, 48, { lineBreak: false });

    doc.fillColor("#fde2e6").font("Helvetica").fontSize(8)
      .text("BOQ ID", 0, 14, { width: CONTENT_W, align: "right", lineBreak: false });
    doc.fillColor("#ffffff").font("Helvetica-Bold").fontSize(10)
      .text(pdfSafe(quoteRef), 0, 26, { width: CONTENT_W, align: "right", lineBreak: false });
    doc.fillColor("#fde2e6").font("Helvetica").fontSize(8)
      .text("DATE", 0, 44, { width: CONTENT_W, align: "right", lineBreak: false });
    doc.fillColor("#ffffff").font("Helvetica-Bold").fontSize(9)
      .text(generatedLabel, 0, 55, { width: CONTENT_W, align: "right", lineBreak: false });
  };

  const drawFooter = (pageNum, pageCount) => {
    const savedBottom = doc.page.margins.bottom;
    doc.page.margins.bottom = 0;
    const fy = PAGE_H - FOOTER_H;
    doc.strokeColor(PDF_COLOR.brand).lineWidth(1.2).moveTo(M, fy).lineTo(RIGHT, fy).stroke();
    doc.fillColor(PDF_COLOR.muted).font("Helvetica").fontSize(8)
      .text("Sintex by Welspun  •  Plumbing Contractor Calculator", M, fy + 8, {
        width: CONTENT_W / 2, lineBreak: false,
      });
    doc.fillColor(PDF_COLOR.ink).font("Helvetica-Bold").fontSize(8)
      .text(`Page ${pageNum} of ${pageCount}`, M + CONTENT_W / 2, fy + 8, {
        width: CONTENT_W / 2, align: "right", lineBreak: false,
      });
    doc.fillColor(PDF_COLOR.muted).font("Helvetica").fontSize(7)
      .text(`BOQ ID: ${pdfSafe(quoteRef)}`, M, fy + 22, {
        width: CONTENT_W, align: "center", lineBreak: false,
      });
    doc.page.margins.bottom = savedBottom;
  };

  drawHeader();
  doc.on("pageAdded", () => drawHeader());

  const stampFooters = () => {
    const range = doc.bufferedPageRange();
    for (let i = 0; i < range.count; i++) {
      doc.switchToPage(range.start + i);
      drawFooter(i + 1, range.count);
    }
  };

  return {
    doc, drawHeader, drawFooter, stampFooters,
    geom: { PAGE_W, PAGE_H, M, RIGHT, CONTENT_W, BOTTOM, HEADER_H, FOOTER_H, yStart: HEADER_H + 18 },
    COLOR: PDF_COLOR, safe: pdfSafe, inr: pdfInr,
  };
}

/**
 * Render a configuration block (rows of label/value cards). Returns the new y.
 * `cfg` is an array of [label, value] pairs.
 */
function drawConfigGrid(ctx, y, cfg) {
  const { doc, geom, COLOR, safe } = ctx;
  const { M, CONTENT_W } = geom;
  doc.fillColor(COLOR.brandDark).font("Helvetica-Bold").fontSize(12)
    .text("Configuration", M, y);
  y = doc.y + 6;

  const colsPerRow = 3;
  const cellW = CONTENT_W / colsPerRow;
  const cellPad = 8;
  const labelH = 12;
  const minCellH = 36;
  const cellGap = 4;

  const cellHeights = cfg.map(([, value]) => {
    doc.font("Helvetica-Bold").fontSize(10);
    const valueH = doc.heightOfString(safe(value), { width: cellW - 16 });
    return Math.max(minCellH, labelH + 6 + valueH + 8);
  });
  const rowHeights = [];
  for (let i = 0; i < cfg.length; i += colsPerRow) {
    rowHeights.push(Math.max(...cellHeights.slice(i, i + colsPerRow)));
  }

  let cy = y;
  for (let i = 0; i < cfg.length; i++) {
    const rowIdx = Math.floor(i / colsPerRow);
    const colIdx = i % colsPerRow;
    const cellH = rowHeights[rowIdx];
    const cx = M + colIdx * cellW;
    doc.roundedRect(cx + 2, cy, cellW - 4, cellH - cellGap, 4)
      .lineWidth(0.6).strokeColor(COLOR.border).stroke();
    doc.fillColor(COLOR.muted).font("Helvetica").fontSize(8)
      .text(cfg[i][0].toUpperCase(), cx + cellPad, cy + 5, { width: cellW - 16, lineBreak: false });
    doc.fillColor(COLOR.ink).font("Helvetica-Bold").fontSize(10)
      .text(safe(cfg[i][1]), cx + cellPad, cy + 17, { width: cellW - 16, height: cellH - 22 });
    if (colIdx === colsPerRow - 1 || i === cfg.length - 1) cy += cellH;
  }
  return cy + 8;
}

// ---------------------------------------------------------------------------
// Item table renderer (used by both BOQ sections)
// ---------------------------------------------------------------------------

function drawItemTable(ctx, y, sections) {
  // sections = [{ label, subtotal, groups: [{ groupName, subtotal, items: [...] }] }, ...]
  // Each top-level section gets a band header; groups within get inline separators.
  const { doc, geom, COLOR, safe, inr } = ctx;
  const { M, CONTENT_W, BOTTOM, yStart } = geom;

  const COLS = [
    { key: "sr",    label: "Sr.",          w: 22, align: "right" },
    { key: "code",  label: "Material Code",w: 96 },
    { key: "desc",  label: "Description",  w: 140 },
    { key: "sub",   label: "Sub Category", w: 72 },
    { key: "type",  label: "Type",         w: 42 },
    { key: "size",  label: "Size",         w: 50 },
    { key: "mrp",   label: "MRP",          w: 46, align: "right" },
    { key: "qty",   label: "Qty",          w: 28, align: "right" },
    { key: "amt",   label: "Amount",       w: 68, align: "right" },
  ];
  const totalColW = COLS.reduce((s, c) => s + c.w, 0);
  const scale = CONTENT_W / totalColW;
  COLS.forEach((c) => { c.w *= scale; });

  const rowH = 18;
  const headRowH = 20;
  const groupSepH = 18;
  const sectionBandH = 26;

  const drawColHeaders = () => {
    doc.rect(M, y, CONTENT_W, headRowH).fillColor(COLOR.headBg).fill();
    let cx = M;
    doc.fillColor(COLOR.headText).font("Helvetica-Bold").fontSize(8);
    for (const c of COLS) {
      doc.text(c.label, cx + 4, y + 6, { width: c.w - 8, align: c.align || "left", lineBreak: false });
      cx += c.w;
    }
    y += headRowH;
  };

  const drawSectionBand = (label, subtotal) => {
    if (y + sectionBandH + headRowH + rowH > BOTTOM - 20) { doc.addPage(); y = yStart; }
    doc.rect(M, y, CONTENT_W, sectionBandH).fillColor(COLOR.brandDark).fill();
    doc.fillColor("#ffffff").font("Helvetica-Bold").fontSize(11)
      .text(safe(label), M + 10, y + 8, { width: CONTENT_W - 180, lineBreak: false });
    if (subtotal != null) {
      doc.font("Helvetica-Bold").fontSize(10)
        .text(`Section Total: ${inr(subtotal)}`, M + CONTENT_W - 180, y + 9,
              { width: 170, align: "right", lineBreak: false });
    }
    y += sectionBandH;
  };

  const drawGroupSeparator = (g) => {
    if (y + groupSepH + rowH > BOTTOM - 20) { doc.addPage(); y = yStart; drawColHeaders(); }
    doc.rect(M, y, CONTENT_W, groupSepH).fillColor("#fff0f1").fill();
    doc.rect(M, y, 3, groupSepH).fillColor(COLOR.brand).fill();
    doc.fillColor(COLOR.brandDark).font("Helvetica-Bold").fontSize(9)
      .text(safe(g.groupName), M + 10, y + 5, { width: CONTENT_W - 160, lineBreak: false });
    doc.fillColor(COLOR.brandDark).font("Helvetica").fontSize(8)
      .text(`Group Subtotal: ${inr(g.subtotal)}`, M + CONTENT_W - 160, y + 5,
            { width: 150, align: "right", lineBreak: false });
    y += groupSepH;
  };

  const drawDataRow = (item, sr, alt) => {
    const descWidth = COLS[2].w - 8;
    const descHeight = doc.font("Helvetica").fontSize(8).heightOfString(safe(item.description), { width: descWidth });
    const thisRowH = Math.max(rowH, descHeight + 8);

    if (y + thisRowH > BOTTOM - 20) {
      doc.addPage();
      y = yStart;
      drawColHeaders();
    }

    if (alt) doc.rect(M, y, CONTENT_W, thisRowH).fillColor(COLOR.rowAlt).fill();
    doc.strokeColor(COLOR.border).lineWidth(0.4)
      .moveTo(M, y + thisRowH).lineTo(M + CONTENT_W, y + thisRowH).stroke();

    const sizeStr = [item.sizeInch, item.sizeMm].filter(Boolean).join(" / ");
    const values = {
      sr: String(sr),
      code: safe(item.materialCode),
      desc: safe(item.description),
      sub: safe(item.subCategory),
      type: safe(item.pipeType),
      size: safe(sizeStr),
      mrp: inr(item.mrp),
      qty: String(item.qty),
      amt: inr(item.assessableValue),
    };

    let cx = M;
    doc.fillColor(COLOR.ink).font("Helvetica").fontSize(8);
    for (const c of COLS) {
      doc.text(values[c.key] || "", cx + 4, y + 4, {
        width: c.w - 8, align: c.align || "left", height: thisRowH - 6, ellipsis: true,
      });
      cx += c.w;
    }
    y += thisRowH;
  };

  for (const section of sections) {
    drawSectionBand(section.label, section.subtotal);
    drawColHeaders();
    let sr = 1;
    let alt = false;
    for (const g of section.groups) {
      drawGroupSeparator(g);
      for (const it of g.items) {
        drawDataRow(it, sr++, alt);
        alt = !alt;
      }
    }
    y += 10; // gap between sections
  }

  return y;
}

// ---------------------------------------------------------------------------
// BOQ export — full itemized document with per-bathroom sections + kitchens
// ---------------------------------------------------------------------------

function exportBoqPdf2(quote, quoteRef) {
  return new Promise((resolve, reject) => {
    const ctx = initPdf({
      headerSubtitle: "BILL OF QUANTITIES (BOQ)",
      quoteRef,
      generatedAt: quote.generatedAt,
      info: { Title: `Sintex BOQ - ${quote.selection.system}`, Subject: "Plumbing BOQ" },
    });
    const { doc, geom, COLOR, safe, inr, stampFooters } = ctx;
    const { M, CONTENT_W, BOTTOM, yStart, RIGHT } = geom;

    const chunks = [];
    doc.on("data", (c) => chunks.push(c));
    doc.on("end", () => resolve(Buffer.concat(chunks)));
    doc.on("error", reject);

    let y = yStart;

    // ---- Configuration ----
    const sel = quote.selection;
    const cfg = [
      ["System", sel.system],
      ["Piping System", sel.pipingSystem],
      ["No. of Bathrooms", String(sel.bathrooms.length)],
      ["No. of Kitchens", String(sel.kitchens)],
      ["Grand Total", inr(quote.grandTotal)],
    ];
    y = drawConfigGrid(ctx, y, cfg);

    // ---- Section list ----
    const sections = quote.bathroomSections.map((b) => ({
      label: `Bathroom ${b.index} — ${b.shape}, ${b.dimensionsLabel}`,
      subtotal: b.subtotal,
      groups: b.groups,
    }));
    sections.push({
      label: `Kitchens (${quote.kitchenSection.count} × base BoQ)`,
      subtotal: quote.kitchenSection.subtotal,
      groups: quote.kitchenSection.groups,
    });

    y = drawItemTable(ctx, y, sections);

    // ---- Grand total band ----
    if (y + 28 > BOTTOM - 20) { doc.addPage(); y = yStart; }
    y += 6;
    doc.rect(M, y, CONTENT_W, 28).fillColor(COLOR.totalBg).fill();
    doc.fillColor(COLOR.totalText).font("Helvetica-Bold").fontSize(13)
      .text("Grand Total", M + 12, y + 8, { width: CONTENT_W - 200, lineBreak: false });
    doc.text(inr(quote.grandTotal), M + CONTENT_W - 200, y + 8, { width: 190, align: "right", lineBreak: false });
    y += 36;

    // ---- Note ----
    if (y + 60 > BOTTOM) { doc.addPage(); y = yStart; }
    doc.strokeColor(COLOR.border).lineWidth(0.5).moveTo(M, y).lineTo(RIGHT, y).stroke();
    y += 8;
    doc.fillColor(COLOR.ink).font("Helvetica").fontSize(8).text(
      "Note: This BOQ is tentative. Per-bathroom quantities are scaled from a base 8 × 4 ft reference room " +
      "(1.5% per ft per axis) with a 10% uplift for C Shape bathrooms. Kitchens use the base BoQ unscaled. " +
      "Prices are based on current SAP MRP reference data and are subject to revision. " +
      "All amounts are in Indian Rupees (Rs.).",
      M, y, { width: CONTENT_W }
    );

    stampFooters();
    doc.end();
  });
}

// ---------------------------------------------------------------------------
// Summary export — short overview document with the rollup table
// ---------------------------------------------------------------------------

function exportSummaryPdf(quote, quoteRef) {
  return new Promise((resolve, reject) => {
    const ctx = initPdf({
      headerSubtitle: "PLUMBING ESTIMATE - SUMMARY",
      quoteRef,
      generatedAt: quote.generatedAt,
      info: { Title: `Sintex Estimate Summary - ${quote.selection.system}`, Subject: "Plumbing Estimate Summary" },
    });
    const { doc, geom, COLOR, safe, inr, stampFooters } = ctx;
    const { M, CONTENT_W, BOTTOM, yStart, RIGHT } = geom;

    const chunks = [];
    doc.on("data", (c) => chunks.push(c));
    doc.on("end", () => resolve(Buffer.concat(chunks)));
    doc.on("error", reject);

    let y = yStart;

    // ---- Configuration ----
    const sel = quote.selection;
    const cfg = [
      ["System", sel.system],
      ["Piping System", sel.pipingSystem],
      ["No. of Bathrooms", String(sel.bathrooms.length)],
      ["No. of Kitchens", String(sel.kitchens)],
    ];
    y = drawConfigGrid(ctx, y, cfg);

    // ---- Summary heading ----
    doc.fillColor(COLOR.brandDark).font("Helvetica-Bold").fontSize(12)
      .text("Estimate Summary", M, y);
    y = doc.y + 8;

    // ---- Summary table ----
    // Columns: Item (50%), Number (15%), Per Cost (17.5%), Total (17.5%)
    const COL = {
      item:   { w: CONTENT_W * 0.50, label: "Item",      align: "left"  },
      num:    { w: CONTENT_W * 0.12, label: "Number",    align: "right" },
      per:    { w: CONTENT_W * 0.19, label: "Per Cost",  align: "right" },
      total:  { w: CONTENT_W * 0.19, label: "Total",     align: "right" },
    };
    const order = ["item", "num", "per", "total"];
    const headRowH = 22;
    const rowH = 26;

    // Header
    doc.rect(M, y, CONTENT_W, headRowH).fillColor(COLOR.headBg).fill();
    let cx = M;
    doc.fillColor(COLOR.headText).font("Helvetica-Bold").fontSize(10);
    for (const k of order) {
      doc.text(COL[k].label, cx + 10, y + 6,
        { width: COL[k].w - 20, align: COL[k].align, lineBreak: false });
      cx += COL[k].w;
    }
    y += headRowH;

    // Rows
    let alt = false;
    for (const row of quote.summaryRows) {
      if (y + rowH > BOTTOM - 40) { doc.addPage(); y = yStart; }
      if (alt) doc.rect(M, y, CONTENT_W, rowH).fillColor(COLOR.rowAlt).fill();
      doc.strokeColor(COLOR.border).lineWidth(0.4)
        .moveTo(M, y + rowH).lineTo(M + CONTENT_W, y + rowH).stroke();

      const label = row.kind === "Bathroom"
        ? `Bathroom — ${row.shape}, ${row.dimensionsLabel}`
        : "Kitchen";

      const values = {
        item: label,
        num: String(row.count),
        per: inr(row.perCost),
        total: inr(row.total),
      };

      cx = M;
      doc.fillColor(COLOR.ink).font("Helvetica").fontSize(10);
      for (const k of order) {
        doc.text(values[k], cx + 10, y + 8,
          { width: COL[k].w - 20, align: COL[k].align, lineBreak: false });
        cx += COL[k].w;
      }
      y += rowH;
      alt = !alt;
    }

    // Grand total row
    y += 6;
    doc.rect(M, y, CONTENT_W, rowH + 4).fillColor(COLOR.totalBg).fill();
    cx = M;
    doc.fillColor(COLOR.totalText).font("Helvetica-Bold").fontSize(12);
    const totalValues = { item: "Grand Total", num: "", per: "", total: inr(quote.grandTotal) };
    for (const k of order) {
      doc.text(totalValues[k], cx + 10, y + 10,
        { width: COL[k].w - 20, align: COL[k].align, lineBreak: false });
      cx += COL[k].w;
    }
    y += rowH + 12;

    // ---- Features callout (highlights ANTI RAT + ANTIMICROBIAL) ----
    // Drawn just above the banner so it frames the visual that follows.
    const SINTEX_URL = "https://www.sintexonline.com/plastic-pipes/";
    const LINK_TEXT  = "www.sintexonline.com/plastic-pipes/";

    if (y + 50 > BOTTOM) { doc.addPage(); y = yStart; }
    y += 4;
    doc.fillColor(COLOR.ink).font("Helvetica").fontSize(9.5)
       .text("Sintex Pipes come with ", M, y, { width: CONTENT_W, continued: true })
       .fillColor(COLOR.brand).font("Helvetica-Bold")
       .text("ANTI RAT", { continued: true })
       .fillColor(COLOR.ink).font("Helvetica")
       .text(" and ", { continued: true })
       .fillColor(COLOR.brand).font("Helvetica-Bold")
       .text("ANTIMICROBIAL", { continued: true })
       .fillColor(COLOR.ink).font("Helvetica")
       .text(" protection. To know more about these features, click the banner below or visit ", { continued: true })
       .fillColor(COLOR.brand).font("Helvetica-Bold")
       .text(LINK_TEXT, {
         link: SINTEX_URL,
         underline: true,
         continued: true,
       })
       .fillColor(COLOR.ink).font("Helvetica")
       .text(" for more details.", { link: null, underline: false });
    y = doc.y + 8;

    // ---- Banner image (clickable — opens Sintex Pipes website) ----
    // Banner is 2560x896 (aspect ~2.857:1). We fit it to the full content width
    // and compute the resulting height to keep proportions intact.
    const bannerW = CONTENT_W;
    const bannerH = bannerW * (896 / 2560);
    if (y + bannerH + 60 > BOTTOM) { doc.addPage(); y = yStart; }
    try {
      doc.image(BANNER_PATH, M, y, { width: bannerW });
      // Link annotation over the image — tapping the banner opens the site.
      doc.link(M, y, bannerW, bannerH, SINTEX_URL);
    } catch (e) {
      // If the banner asset isn't available, just continue without it.
    }
    y += bannerH;

    // ---- Note ----
    if (y + 60 > BOTTOM) { doc.addPage(); y = yStart; }
    y += 6;
    doc.strokeColor(COLOR.border).lineWidth(0.5).moveTo(M, y).lineTo(RIGHT, y).stroke();
    y += 8;
    doc.fillColor(COLOR.ink).font("Helvetica").fontSize(8).text(
      "Note: This summary is tentative. Bathrooms with identical shape and dimensions are grouped " +
      "into a single row; each kitchen is priced as one base 8 × 4 ft BoQ. All amounts in Indian Rupees (Rs.).",
      M, y, { width: CONTENT_W }
    );

    stampFooters();
    doc.end();
  });
}


const app = express();
app.set("trust proxy", true);
app.use(cors());
app.use(express.json());

app.get("/api/health", (req, res) => {
  res.json({ ok: true, loaded: MASTER != null });
});

app.get("/api/options", (req, res) => {
  res.json({
    // New flow
    systems: SYSTEMS,
    pipingSystems: PIPING_SYSTEMS,
    baseLength: BASE_LENGTH,
    baseWidth: BASE_WIDTH,
    upliftPerFt: UPLIFT_PER_FT,
    // Legacy (unused by the new UI but kept for any older client)
    shapes: BATHROOM_SHAPES,
    fixtures: FIXTURES,
    pipeSystems: PIPE_SYSTEMS,
    cShapeUplift: C_SHAPE_UPLIFT,
  });
});

// New flow: build a quote
app.post("/api/quote", (req, res) => {
  try {
    const quote = buildQuote(req.body || {});
    res.json(quote);
  } catch (e) {
    res.status(400).json({ error: e.message });
  }
});

// Serve material-group product images (whitelisted via GROUP_IMAGE_MAP)
app.get("/api/group-image/:group", (req, res) => {
  const file = imageForGroup(req.params.group);
  if (!file) return res.status(404).json({ error: "No image for group" });
  const full = path.join(__dirname, "assets", "groups", file);
  if (!fs.existsSync(full)) return res.status(404).json({ error: "Image file missing" });
  res.sendFile(full);
});

// Serve other product/system images by filename — restricted to the assets/groups directory
// so the route can never be abused to read arbitrary paths.
app.get("/api/asset/:file", (req, res) => {
  const name = String(req.params.file || "");
  // Only allow safe filenames (letters, digits, _ - . space)
  if (!/^[A-Za-z0-9_.\- ]+$/.test(name)) return res.status(400).json({ error: "Bad filename" });
  const full = path.join(__dirname, "assets", "groups", name);
  // Defensive: make sure resolved path is still inside the assets/groups directory.
  const base = path.join(__dirname, "assets", "groups");
  if (!full.startsWith(base + path.sep)) return res.status(400).json({ error: "Bad path" });
  if (!fs.existsSync(full)) return res.status(404).json({ error: "Not found" });
  res.sendFile(full);
});

// Legacy endpoint kept for backward compatibility (older UI)
app.post("/api/boq", (req, res) => {
  try {
    const boq = buildBoq(req.body || {});
    res.json(boq);
  } catch (e) {
    res.status(400).json({ error: e.message });
  }
});

app.get("/api/quote-ref", (req, res) => {
  const ip = clientIp(req);
  res.json({ quoteRef: generateQuoteRef(ip), ip });
});

// New flow: BOQ (itemized, partitioned by bathroom + kitchen)
app.post("/api/quote/boq-pdf", async (req, res) => {
  try {
    const { quoteRef: providedRef, ...body } = req.body || {};
    const quote = buildQuote(body);
    const ip = clientIp(req);
    const quoteRef = providedRef || generateQuoteRef(ip);
    const buf = await exportBoqPdf2(quote, quoteRef);
    const fname = `BOQ_${String(quoteRef).replace(/[^a-z0-9_.-]+/gi, "_")}.pdf`;
    res.setHeader("Content-Type", "application/pdf");
    res.setHeader("Content-Disposition", `attachment; filename="${fname}"`);
    res.send(buf);
  } catch (e) {
    res.status(400).json({ error: e.message });
  }
});

// New flow: Estimate summary (compact overview)
app.post("/api/quote/summary-pdf", async (req, res) => {
  try {
    const { quoteRef: providedRef, ...body } = req.body || {};
    const quote = buildQuote(body);
    const ip = clientIp(req);
    const quoteRef = providedRef || generateQuoteRef(ip);
    const buf = await exportSummaryPdf(quote, quoteRef);
    const fname = `Estimate_${String(quoteRef).replace(/[^a-z0-9_.-]+/gi, "_")}.pdf`;
    res.setHeader("Content-Type", "application/pdf");
    res.setHeader("Content-Disposition", `attachment; filename="${fname}"`);
    res.send(buf);
  } catch (e) {
    res.status(400).json({ error: e.message });
  }
});

// Legacy single-shot endpoint kept for backward compatibility — alias to BOQ.
app.post("/api/quote/pdf", async (req, res) => {
  try {
    const { quoteRef: providedRef, ...body } = req.body || {};
    const quote = buildQuote(body);
    const ip = clientIp(req);
    const quoteRef = providedRef || generateQuoteRef(ip);
    const buf = await exportBoqPdf2(quote, quoteRef);
    const fname = `BOQ_${String(quoteRef).replace(/[^a-z0-9_.-]+/gi, "_")}.pdf`;
    res.setHeader("Content-Type", "application/pdf");
    res.setHeader("Content-Disposition", `attachment; filename="${fname}"`);
    res.send(buf);
  } catch (e) {
    res.status(400).json({ error: e.message });
  }
});

// Legacy export kept for backward compatibility
app.post("/api/boq/export", async (req, res) => {
  try {
    const { quoteRef: providedRef, ...selection } = req.body || {};
    const boq = buildBoq(selection);
    const ip = clientIp(req);
    const quoteRef = providedRef || generateQuoteRef(ip);
    const buf = await exportBoqPdf(boq, quoteRef);
    const fname = String(quoteRef).replace(/[^a-z0-9_.-]+/gi, "_") + ".pdf";
    res.setHeader("Content-Type", "application/pdf");
    res.setHeader("Content-Disposition", `attachment; filename="${fname}"`);
    res.send(buf);
  } catch (e) {
    res.status(400).json({ error: e.message });
  }
});

// ---------------------------------------------------------------------------
// Static serving — when the React frontend has been built, serve the
// production bundle from the same Express process so the whole app is
// reachable on a single URL (one deployment, no CORS).
// ---------------------------------------------------------------------------

const FRONTEND_DIST  = path.join(__dirname, "..", "frontend", "dist");
const FRONTEND_INDEX = path.join(FRONTEND_DIST, "index.html");
const FRONTEND_BUILT = fs.existsSync(FRONTEND_INDEX);

console.log(`[api] frontend/dist exists: ${FRONTEND_BUILT}`);
console.log(`[api] frontend/dist path:   ${FRONTEND_DIST}`);

if (FRONTEND_BUILT) {
  // 1) Lightweight per-request log so it's obvious whether traffic is
  //    even reaching the static block. Set NO_REQ_LOG=1 to silence.
  if (!process.env.NO_REQ_LOG) {
    app.use((req, res, next) => {
      console.log(`[req] ${req.method} ${req.url}`);
      next();
    });
  }

  // 2) Serve the built bundle (index.html, /assets/*, etc.). express.static
  //    handles `/` -> index.html automatically.
  app.use(express.static(FRONTEND_DIST));

  // 3) SPA fallback — any GET that wasn't an API call and wasn't a real
  //    file under dist falls back to index.html so client-side routing
  //    works. Plain middleware (no regex routes) for maximum portability.
  app.use((req, res, next) => {
    if (req.method !== "GET") return next();
    if (req.path.startsWith("/api/")) return next();
    res.sendFile(FRONTEND_INDEX, (err) => {
      if (err) {
        console.error(`[api] sendFile error for ${req.url}:`, err.message);
        next(err);
      }
    });
  });

  console.log(`[api] Serving built frontend from ${FRONTEND_DIST}`);
} else {
  console.log("[api] No frontend build found — API only.");
  console.log("[api] Run `npm run build` in the project root to build the React app.");
}

// ---------------------------------------------------------------------------
// Boot
// ---------------------------------------------------------------------------

// In production hosts (Render, Railway, Fly, Heroku, etc.) the server must
// bind to 0.0.0.0 so the host's router can reach it. Locally we keep the
// safer 127.0.0.1 default.
const HOST = process.env.HOST
  || (process.env.NODE_ENV === "production" ? "0.0.0.0" : "127.0.0.1");

loadMaster()
  .then(() => {
    app.listen(PORT, HOST, () => {
      console.log(`[api] Plumbing Calculator backend on http://${HOST}:${PORT}`);
    });
  })
  .catch((e) => {
    console.error("Failed to load master data:", e);
    process.exit(1);
  });
