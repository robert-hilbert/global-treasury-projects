/**
 * ═══════════════════════════════════════════════════════════════════════════
 * extractEuropean.ts — Office Script  v2.0
 * Runs ON the European Exposure Forecast workbook.
 * Returns canonical FX JSON schema with EURUSD (FG + Royalties)
 * and EURPLN (Revenue).
 *
 * VERIFIED CELL REFERENCES (March 2026):
 *   Sheet : "Exposure FCST"
 *   Row 7  = EURUSD Finished Goods
 *   Row 10 = EURUSD Royalties
 *   Row 39 = EURPLN Revenue
 *   Cols Q(17)–AB(28) = Jan–Dec 2026
 *   Col  AC(29)        = TOTAL (skipped)
 *   Cols AD(30)–AO(41) = Jan–Dec 2027
 *   Total: 24 months
 *
 * Schema contract: schema_version "1.0"
 * Called by Power Automate "Run script" action targeting the source file.
 * ═══════════════════════════════════════════════════════════════════════════
 */

// ── Canonical schema interfaces ──────────────────────────────────────────

interface ExposureRecord {
  transaction_type: string;
  period_label: string;       // "Jan", "Feb" …
  period_year: number;        // 2026 | 2027
  period_date: string;        // ISO end-of-month "YYYY-MM-DD"
  exposure_fc: number | null; // foreign-currency amount; null = blank cell
  is_estimate: boolean;       // true = current month or future
}

interface ValidationEnvelope {
  header_hash: string;        // hash of header row — detects pivot collapse
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

interface EuropeanExtract {
  schema_version: string;
  source_file: string;
  extracted_at: string;
  extraction_script: string;
  pairs: {
    EURUSD: FxExtract;
    EURPLN: FxExtract;
  };
  overall_passed: boolean;
}

// ── Main ─────────────────────────────────────────────────────────────────

function main(workbook: ExcelScript.Workbook): EuropeanExtract {
  const timestamp = new Date().toISOString();
  const scriptName = "extractEuropean.ts";

  // Derive source filename from workbook name
  const sourceFile = workbook.getName();

  // ── Sheet validation ──
  const ws = workbook.getWorksheet("Exposure FCST");
  if (!ws) {
    console.log("ERROR: Sheet 'Exposure FCST' not found");
    const failEnvelope: ValidationEnvelope = {
      header_hash: "",
      period_count: 0,
      row_count: 0,
      null_count: 0,
      passed: false,
      failure_reason: "Sheet 'Exposure FCST' not found",
    };
    const failExtract: FxExtract = {
      schema_version: "1.0",
      currency_pair: "",
      exposed_currency: "",
      source_file: sourceFile,
      extracted_at: timestamp,
      extraction_script: scriptName,
      exposures: [],
      validation: failEnvelope,
    };
    return {
      schema_version: "1.0",
      source_file: sourceFile,
      extracted_at: timestamp,
      extraction_script: "extractEuropean.ts",
      pairs: { EURUSD: failExtract, EURPLN: { ...failExtract } },
      overall_passed: false,
    };
  }

  console.log("Sheet 'Exposure FCST' found. Starting extraction...");

  // ── Helpers ──

  // Column index (1-based) → letter(s)
  function colLetter(col: number): string {
    let result = "";
    let c = col;
    while (c > 0) {
      c--;
      result = String.fromCharCode(65 + (c % 26)) + result;
      c = Math.floor(c / 26);
    }
    return result;
  }

  // Safely read a numeric cell; returns null if blank
  function readNum(row: number, col: number): number | null {
    const val = ws.getCell(row - 1, col - 1).getValue();
    if (val === null || val === undefined || val === "") return null;
    const n = typeof val === "number" ? val : Number(val);
    return isNaN(n) ? null : n;
  }

  // ISO end-of-month date string for a given year and 1-based month
  function endOfMonth(year: number, month: number): string {
    // Day 0 of next month = last day of this month
    const d = new Date(Date.UTC(year, month, 0));
    return d.toISOString().slice(0, 10);
  }

  // Simple header hash: concatenate header cell values into a string and sum char codes
  function buildHeaderHash(headerRow: number, startCol: number, endCol: number): string {
    let raw = "";
    for (let c = startCol; c <= endCol; c++) {
      const v = ws.getCell(headerRow - 1, c - 1).getValue();
      raw += String(v ?? "");
    }
    let hash = 0;
    for (let i = 0; i < raw.length; i++) {
      hash = (hash * 31 + raw.charCodeAt(i)) & 0xffffffff;
    }
    return (hash >>> 0).toString(16).padStart(8, "0");
  }

  // Determine if a period is an estimate (current month or future)
  const now = new Date();
  const currentYearMonth = now.getFullYear() * 100 + (now.getMonth() + 1); // e.g. 202604

  function isEstimate(year: number, month: number): boolean {
    return year * 100 + month >= currentYearMonth;
  }

  // Month index (0-based) → short label
  const MONTH_LABELS = ["Jan","Feb","Mar","Apr","May","Jun",
                         "Jul","Aug","Sep","Oct","Nov","Dec"];

  // ── Column mapping: 24 months with gap at AC(29) ──
  interface ColMap {
    colIndex: number;
    year: number;
    month: number; // 1-based
  }

  const colMap: ColMap[] = [];

  // Jan–Dec 2026: cols 17–28
  for (let m = 0; m < 12; m++) {
    colMap.push({ colIndex: 17 + m, year: 2026, month: m + 1 });
  }
  // Jan–Dec 2027: cols 30–41 (skip col 29 = TOTAL)
  for (let m = 0; m < 12; m++) {
    colMap.push({ colIndex: 30 + m, year: 2027, month: m + 1 });
  }

  // ── Header hash (use row 6 as the header row — adjust if needed) ──
  const HEADER_ROW = 6;
  const headerHash = buildHeaderHash(HEADER_ROW, 17, 41);

  // ── Extraction helper: build FxExtract for one pair/row ──
  function extractPair(
    currencyPair: string,
    exposedCurrency: string,
    rows: { row: number; transaction_type: string }[]
  ): FxExtract {
    const exposures: ExposureRecord[] = [];
    let nullCount = 0;

    for (const { row, transaction_type } of rows) {
      for (const col of colMap) {
        const raw = readNum(row, col.colIndex);
        if (raw === null) nullCount++;

        exposures.push({
          transaction_type,
          period_label: MONTH_LABELS[col.month - 1],
          period_year: col.year,
          period_date: endOfMonth(col.year, col.month),
          exposure_fc: raw,
          is_estimate: isEstimate(col.year, col.month),
        });
      }
    }

    // Validation checks
    const expectedPeriods = 24;
    const expectedRows = expectedPeriods * rows.length;
    let failureReason: string | null = null;

    if (colMap.length !== expectedPeriods) {
      failureReason = `Period count mismatch: expected ${expectedPeriods}, got ${colMap.length}`;
    }

    const passed = failureReason === null;

    const validation: ValidationEnvelope = {
      header_hash: headerHash,
      period_count: colMap.length,
      row_count: exposures.length,
      null_count: nullCount,
      passed,
      failure_reason: failureReason,
    };

    return {
      schema_version: "1.0",
      currency_pair: currencyPair,
      exposed_currency: exposedCurrency,
      source_file: sourceFile,
      extracted_at: timestamp,
      extraction_script: scriptName,
      exposures,
      validation,
    };
  }

  // ── Extract both pairs ──
  const eurusd = extractPair("EURUSD", "EUR", [
    { row: 7,  transaction_type: "Finished Goods" },
    { row: 10, transaction_type: "Royalties" },
  ]);

  const eurpln = extractPair("EURPLN", "PLN", [
    { row: 39, transaction_type: "Revenue" },
  ]);

  const overall_passed = eurusd.validation.passed && eurpln.validation.passed;

  // ── Verification log ──
  console.log("=== EXTRACTION COMPLETE ===");
  console.log(`Source file   : ${sourceFile}`);
  console.log(`Extracted at  : ${timestamp}`);
  console.log(`Header hash   : ${headerHash}`);
  console.log(`Overall passed: ${overall_passed}`);
  console.log("");

  console.log("--- EURUSD Finished Goods (Row 7) — first 3 months ---");
  const fgRecords = eurusd.exposures.filter(e => e.transaction_type === "Finished Goods");
  for (let i = 0; i < 3 && i < fgRecords.length; i++) {
    console.log(`  ${fgRecords[i].period_date}  exposure_fc=${fgRecords[i].exposure_fc}  is_estimate=${fgRecords[i].is_estimate}`);
  }
  console.log("");

  console.log("--- EURUSD Royalties (Row 10) — first 3 months ---");
  const royRecords = eurusd.exposures.filter(e => e.transaction_type === "Royalties");
  for (let i = 0; i < 3 && i < royRecords.length; i++) {
    console.log(`  ${royRecords[i].period_date}  exposure_fc=${royRecords[i].exposure_fc}  is_estimate=${royRecords[i].is_estimate}`);
  }
  console.log("");

  console.log("--- EURPLN Revenue (Row 39) — first 3 months ---");
  const plnRecords = eurpln.exposures.filter(e => e.transaction_type === "Revenue");
  for (let i = 0; i < 3 && i < plnRecords.length; i++) {
    console.log(`  ${plnRecords[i].period_date}  exposure_fc=${plnRecords[i].exposure_fc}  is_estimate=${plnRecords[i].is_estimate}`);
  }
  console.log("");

  console.log("VERIFY: Cross-check values above against cells Q7, Q10, Q39 in the workbook.");
  console.log(`EURUSD null_count : ${eurusd.validation.null_count}`);
  console.log(`EURPLN null_count : ${eurpln.validation.null_count}`);
  if (!overall_passed) {
    console.log(`EURUSD failure: ${eurusd.validation.failure_reason}`);
    console.log(`EURPLN failure: ${eurpln.validation.failure_reason}`);
  }

  return {
    schema_version: "1.0",
    source_file: sourceFile,
    extracted_at: timestamp,
    extraction_script: "extractEuropean.ts",
    pairs: { EURUSD: eurusd, EURPLN: eurpln },
    overall_passed,
  };
}
