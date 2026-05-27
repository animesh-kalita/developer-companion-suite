console.time("DocumentEnd");

const XLSX = require("xlsx");
const fs = require("fs");
const dayjs = require("dayjs");

// ====== CONFIG ======
const INPUT_FILE = "Manual-CAN-Pioneer-Server-Response.xlsx";

// ====== READ EXCEL ======
const workbook = XLSX.readFile(INPUT_FILE);
const sheet = workbook.Sheets[workbook.SheetNames[0]];

// Get merges
const merges = sheet["!merges"] || [];

// Get raw 2D array (defval ensures empty cells are "" not undefined)
let data = XLSX.utils.sheet_to_json(sheet, {
  header: 1,
  raw: false,
  defval: "",
});

// ====== APPLY MERGES ======
// For every merge region, propagate the top-left value to ALL cells in the region,
// overwriting blanks AND already-equal values so every cell in the span is filled.
function applyMerges(data, merges) {
  merges.forEach((merge) => {
    const { s, e } = merge;
    const value = data[s.r]?.[s.c] ?? "";

    for (let r = s.r; r <= e.r; r++) {
      if (!data[r]) data[r] = [];
      for (let c = s.c; c <= e.c; c++) {
        // Always stamp the merge value — even if the cell already has it.
        // This guarantees every position in the span is consistent.
        if (data[r][c] === "" || data[r][c] === undefined) {
          data[r][c] = value;
        }
      }
    }
  });
  return data;
}

data = applyMerges(data, merges);

// ====== HELPERS ======
const clean = (v) =>
  typeof v === "string" ? v.replace(/\n/g, " ").trim() : String(v ?? "").trim();

// ====== COLUMN GROUPING ======
// Row 0 = header titles. Merged cells mean the title appears only in the first
// column of the group; the remaining columns of that group show the same value
// (after applyMerges). We detect group boundaries by watching where a NEW
// non-empty value begins — but we must ignore cells that are merely the
// propagated duplicate of the current group title.
function getColumnGroups(headerRow) {
  const groups = [];
  let current = null;

  headerRow.forEach((rawCell, colIndex) => {
    const cell = clean(rawCell);

    if (!cell) {
      // Empty column — extend current group if one is open
      if (current) current.end = colIndex;
      return;
    }

    if (!current) {
      // First group
      current = { title: cell, start: colIndex, end: colIndex };
      return;
    }

    if (cell === current.title) {
      // Same value propagated from the merge — still the same group
      current.end = colIndex;
    } else {
      // A genuinely new title — close previous group, open new one
      groups.push(current);
      current = { title: cell, start: colIndex, end: colIndex };
    }
  });

  if (current) groups.push(current);
  return groups;
}

// ====== BUILD DOCUMENT ======
// Layout assumption (0-indexed rows):
//   Row 0 → field titles  (merged across their byte-width)
//   Row 1 → hex spec      (e.g. "HEX (2 bytes)")
//   Row 2 → byte positions (e.g. 1, 2 for a 2-byte field)
//   Row 3 → hex values    (e.g. 0x25, 0x25)
//   Row 4+ → notes / free-text (may have gaps / empty rows)
function buildDocument(data) {
  const headerRow = data[0] || [];
  const groups = getColumnGroups(headerRow);

  const result = {};

  groups.forEach(({ title, start, end }) => {
    const column = {
      hex: "",
      positions: [],
      values: [],
      notes: [],
    };

    // ── HEX spec (row 1) ─────────────────────────────────────────────
    // Take the first non-empty cell across the group's column span.
    for (let c = start; c <= end; c++) {
      const val = clean(data[1]?.[c]);
      if (val) {
        column.hex = val;
        break;
      }
    }

    // ── POSITIONS (row 2) ─────────────────────────────────────────────
    // Collect ALL non-empty, non-duplicate cells across the span.
    // Because applyMerges only propagates the merge-source value and
    // individual position cells are NOT merged, every cell here is unique.
    const seenPos = new Set();
    for (let c = start; c <= end; c++) {
      const val = clean(data[2]?.[c]);
      if (val && !seenPos.has(val)) {
        seenPos.add(val);
        column.positions.push(val);
      }
    }

    // ── VALUES (row 3) ────────────────────────────────────────────────
    // Same logic — collect every distinct value across the span.
    // For a 2-byte field both bytes may differ (e.g. 0x00 and 0x57).
    for (let c = start; c <= end; c++) {
      const val = clean(data[3]?.[c]);
      if (val) {
        column.values.push(val); // keep duplicates — they carry meaning (e.g. 0x25 0x25)
      }
    }

    // ── NOTES (row 4 onwards) ─────────────────────────────────────────
    // Notes can appear in any cell within the group's column span across
    // any row from row 4 downward, separated by empty rows.
    // We collect every unique non-empty string found in the span,
    // regardless of gaps.
    const noteSet = new Set();

    // Determine the actual last row of data
    const lastRow = data.length;

    for (let r = 4; r < lastRow; r++) {
      const row = data[r];
      if (!row) continue;

      for (let c = start; c <= end; c++) {
        const val = clean(row[c]);
        if (val) {
          noteSet.add(val);
          // Do NOT break — collect every non-empty cell in this row's span
        }
      }
    }

    column.notes = Array.from(noteSet);

    result[title] = column;
  });

  return result;
}

// ====== MARKDOWN ======
function toMarkdown(doc) {
  let md = "## IoT Packet Document\n\n";
  md += "| Field | HEX | Positions | Values | Notes |\n";
  md += "|------|------|-----------|--------|-------|\n";

  for (const key in doc) {
    const col = doc[key];
    md += `| ${key} | ${col.hex} | ${col.positions.join(", ")} | ${col.values.join(", ")} | ${col.notes.join("; ")} |\n`;
  }

  return md;
}

// ====== EXECUTION ======
const doc = buildDocument(data);

// Timestamp
const timestamp = dayjs().format("YYYY-MM-DD_HH-mm-ss");

// Save JSON
const jsonFile = `iot_doc_${timestamp}.json`;
fs.writeFileSync(jsonFile, JSON.stringify(doc, null, 2));

// Save Markdown (uncomment to enable)
// const mdFile = `iot_doc_${timestamp}.md`;
// fs.writeFileSync(mdFile, toMarkdown(doc));

console.log("✅ Files generated:");
console.log("JSON:", jsonFile);
// console.log("Markdown:", mdFile);

console.timeEnd("DocumentEnd");
