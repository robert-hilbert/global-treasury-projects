/**
 * ═══════════════════════════════════════════════════════════════════════════
 * extractMexico.ts — Office Script  v1.0
 * Runs ON the Mexico Exposure Forecast workbook.
 * Returns canonical FX JSON schema with USDMXN (Finished Goods).
 *
 * VERIFIED CELL REFERENCES (April 2026):
 *   Sheet  : "SUMMARY KB"
 *   Row 5  = Column headers (full month+year labels, e.g. "Jan 2026")
 *   Row 15 = USDMXN Finished Goods exposure (foreign currency)
 *   Cols C(3)–Z(26) = 24 months, Jan 2026–Dec 2027, no gaps
 *
 * Schema contract : schema_version "1.0"
 * Called by Power Automate "Run script" action targeting the source file.
 * ═══════════════════════════════════════════════════════════════════════════
 */

// ── Canonical schema interfaces ──────────────────────────────────────────

interface ExposureRecord {
  transaction_type: string;
  period_label: string;       // "Jan", "Feb" …
  period_year: number;        // e.g. 2026
  period_date: string;        // ISO end-of-month "YYYY-MM-DD"
  exposure_fc: number | null; // foreign-currency amount; null = blank cell
  is_estimate: boolean;       // true = current month or future
}

interface ValidationEnvelope {
  header_hash: string;        // hash of header row — detects column shifts
  period_count: number;       // must equal 24
  row_count: number;          // total ExposureRecords written
  null_count: number;         // blank cells found — non-zero triggers review
  passed: boolean;            // master gate for updateMaster.ts
  failure_reason: string | null;
}

interface FxExtract {
  schema_version: string;
  currency_pair: string;
  exposed_currency: string;
  source_file: string;
  extracted_at: string;       // ISO 8601 UTC
  extraction_script: string;
  exposures: ExposureRecord[];
  validation: ValidationEnvelope;
}

// ── Main ─────────────────────────────────────────────────────────────────

function main(workbook: ExcelScript.Workbook): FxExtract {
  const timestamp = new Date().toISOString();
  const scriptName = "extractMexico.ts";
  const sourceFile = workbook.getName();

  // ── Sheet validation ──
  const ws = workbook.getWorksheet("SUMMARY KB");
  if (!ws) {
    console.log("ERROR: Sheet 'SUMMARY KB' not found");
    return {
      schema_version: "1.0",
      currency_pair: "USDMXN",
      exposed_currency: "MXN",
      source_file: sourceFile,
      extracted_at: timestamp,
      extraction_script: scriptName,
      exposures: [],
      validation: {
        header_hash: "",
        period_count: 0,
        row_count: 0,
        null_count: 0,
        passed: false,
        failure_reason: "Sheet 'SUMMARY KB' not found",
      },
    };
  }

  console.log("Sheet 'SUMMARY KB' found. Starting extraction...");

  // ── Constants ──
  const HEADER_ROW   = 5;   // row containing month+year labels
  const DATA_ROW     = 15;  // USDMXN Finished Goods
  const START_COL    = 3;   // col C (1-based)
  const END_COL      = 26;  // col Z (1-based)
  const NUM_PERIODS  = END_COL - START_COL + 1; // 24

  const MONTH_LABELS = ["Jan","Feb","Mar","Apr","May","Jun",
                         "Jul","Aug","Sep","Oct","Nov","Dec"];

  // ── Batch read header row and data row in two range calls ──
  // Header: row 5, cols C–Z  →  getCell is 0-based: row=4, col=2, width=24
  const headerValues = ws
    .getRangeByIndexes(HEADER_ROW - 1, START_COL - 1, 1, NUM_PERIODS)
    .getValues()[0];

  // Data: row 15, cols C–Z
  const dataValues = ws
    .getRangeByIndexes(DATA_ROW - 1, START_COL - 1, 1, NUM_PERIODS)
    .getValues()[0];

  // ── Header hash — detects pivot/column shifts ──
  let raw = "";
  for (let i = 0; i < headerValues.length; i++) {
    raw += String(headerValues[i] ?? "");
  }
  let hash = 0;
  for (let i = 0; i < raw.length; i++) {
    hash = (hash * 31 + raw.charCodeAt(i)) & 0xffffffff;
  }
  const headerHash = (hash >>> 0).toString(16).padStart(8, "0");

  // ── Parse year/month from header cell value ──
  // Cells display as "Jan 2026" but underlying value is an Excel serial date.
  // Excel serial: days since 1900-01-00 (with leap-year bug: 1900 counted as leap).
  // Conversion: subtract 25569 to get Unix days, multiply by 86400000 for ms.
  function serialToDate(serial: number): Date {
    return new Date((serial - 25569) * 86400000);
  }

  function parseHeaderYear(val: string | number | boolean): number {
    if (typeof val === "number" && val > 40000) {
      return serialToDate(val).getUTCFullYear();
    }
    const s = String(val ?? "");
    const match = s.match(/\d{4}/);
    return match ? parseInt(match[0], 10) : 0;
  }

  function parseHeaderMonth(val: string | number | boolean): number {
    if (typeof val === "number" && val > 40000) {
      return serialToDate(val).getUTCMonth() + 1; // 1-based
    }
    const s = String(val ?? "").trim().slice(0, 3);
    return MONTH_LABELS.indexOf(s) + 1; // 1-based; 0 if not found
  }

  // ── ISO end-of-month date ──
  function endOfMonth(year: number, month: number): string {
    const d = new Date(Date.UTC(year, month, 0));
    return d.toISOString().slice(0, 10);
  }

  // ── is_estimate: current month or future ──
  const now = new Date();
  const currentYearMonth = now.getFullYear() * 100 + (now.getMonth() + 1);

  // ── Build exposure records ──
  const exposures: ExposureRecord[] = [];
  let nullCount = 0;
  let failureReason: string | null = null;

  for (let i = 0; i < NUM_PERIODS; i++) {
    const year  = parseHeaderYear(headerValues[i]);
    const month = parseHeaderMonth(headerValues[i]);

    if (year === 0 || month === 0) {
      failureReason = `Could not parse header at column offset ${i}: "${headerValues[i]}"`;
      break;
    }

    const raw = dataValues[i];
    let exposure_fc: number | null;
    if (raw === null || raw === undefined || raw === "") {
      exposure_fc = null;
      nullCount++;
    } else {
      const n = typeof raw === "number" ? raw : Number(raw);
      exposure_fc = isNaN(n) ? null : n;
      if (exposure_fc === null) nullCount++;
    }

    exposures.push({
      transaction_type: "Finished Goods",
      period_label: MONTH_LABELS[month - 1],
      period_year: year,
      period_date: endOfMonth(year, month),
      exposure_fc,
      is_estimate: year * 100 + month >= currentYearMonth,
    });
  }

  // ── Validation ──
  if (!failureReason && exposures.length !== NUM_PERIODS) {
    failureReason = `Period count mismatch: expected ${NUM_PERIODS}, got ${exposures.length}`;
  }

  const passed = failureReason === null;

  const validation: ValidationEnvelope = {
    header_hash: headerHash,
    period_count: exposures.length,
    row_count: exposures.length,
    null_count: nullCount,
    passed,
    failure_reason: failureReason,
  };

  // ── Verification log — no loops, all inline ──
  console.log("=== EXTRACTION COMPLETE ===");
  console.log(`Source file   : ${sourceFile}`);
  console.log(`Extracted at  : ${timestamp}`);
  console.log(`Header hash   : ${headerHash}`);
  console.log(`Overall passed: ${passed}`);
  console.log("");
  console.log("--- USDMXN Finished Goods (Row 15) — first 3 months ---");
  console.log(`  ${exposures[0]?.period_date}  exposure_fc=${exposures[0]?.exposure_fc}  is_estimate=${exposures[0]?.is_estimate}`);
  console.log(`  ${exposures[1]?.period_date}  exposure_fc=${exposures[1]?.exposure_fc}  is_estimate=${exposures[1]?.is_estimate}`);
  console.log(`  ${exposures[2]?.period_date}  exposure_fc=${exposures[2]?.exposure_fc}  is_estimate=${exposures[2]?.is_estimate}`);
  console.log("");
  console.log("VERIFY: Cross-check values above against cells C15, D15, E15 in SUMMARY KB.");
  console.log(`null_count : ${nullCount}`);
  if (!passed) {
    console.log(`FAILED: ${failureReason}`);
  }

  return {
    schema_version: "1.0",
    currency_pair: "USDMXN",
    exposed_currency: "MXN",
    source_file: sourceFile,
    extracted_at: timestamp,
    extraction_script: scriptName,
    exposures,
    validation,
  };
}
