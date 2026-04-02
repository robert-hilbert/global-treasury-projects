/**
 * ═══════════════════════════════════════════════════════════════════════════
 * extractNorway.ts — Office Script  v1.0
 * Runs ON the Norway Exposure Forecast workbook.
 * Returns canonical FX JSON schema for 5 NOK pairs.
 *
 * VERIFIED CELL REFERENCES (April 2026):
 *   Sheet : "Hedging Impact - P&L vs Act"
 *   Row 2  = Month number (1–12)
 *   Row 3  = Year (2026)
 *   Cols E(5)–P(16) = Jan–Dec 2026, 12 months
 *   2027 periods = 2026 values duplicated, all marked is_estimate=true
 *
 *   SEKNOK  Row 6  — Revenue
 *   CADNOK  Row 7  — Revenue
 *   EURNOK  Row 8  — Revenue  |  Row 16 — COGS
 *   GBPNOK  Row 9  — Revenue
 *   USDNOK  Row 10 — Revenue  |  Row 18 — COGS
 *
 * Schema contract : schema_version "1.0"
 * Called by Power Automate "Run script" action targeting the source file.
 * ═══════════════════════════════════════════════════════════════════════════
 */

// ── Canonical schema interfaces ──────────────────────────────────────────

interface ExposureRecord {
  transaction_type: string;
  period_label: string;
  period_year: number;
  period_date: string;        // ISO end-of-month "YYYY-MM-DD"
  exposure_fc: number | null;
  is_estimate: boolean;
}

interface ValidationEnvelope {
  header_hash: string;
  period_count: number;
  row_count: number;
  null_count: number;
  passed: boolean;
  failure_reason: string | null;
}

interface FxExtract {
  schema_version: string;
  currency_pair: string;
  exposed_currency: string;
  source_file: string;
  extracted_at: string;
  extraction_script: string;
  exposures: ExposureRecord[];
  validation: ValidationEnvelope;
}

interface NorwayExtract {
  schema_version: string;
  source_file: string;
  extracted_at: string;
  extraction_script: string;
  pairs: {
    SEKNOK: FxExtract;
    CADNOK: FxExtract;
    EURNOK: FxExtract;
    GBPNOK: FxExtract;
    USDNOK: FxExtract;
  };
  overall_passed: boolean;
}

// ── Main ─────────────────────────────────────────────────────────────────

function main(workbook: ExcelScript.Workbook): NorwayExtract {
  const timestamp = new Date().toISOString();
  const scriptName = "extractNorway.ts";
  const sourceFile = workbook.getName();

  // ── Sheet validation ──
  const ws = workbook.getWorksheet("Hedging Impact - P&L vs Act");
  if (!ws) {
    console.log("ERROR: Sheet 'Hedging Impact - P&L vs Act' not found");
    const failExtract = (pair: string, ccy: string): FxExtract => ({
      schema_version: "1.0",
      currency_pair: pair,
      exposed_currency: ccy,
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
        failure_reason: "Sheet 'Hedging Impact - P&L vs Act' not found",
      },
    });
    return {
      schema_version: "1.0",
      source_file: sourceFile,
      extracted_at: timestamp,
      extraction_script: scriptName,
      pairs: {
        SEKNOK: failExtract("SEKNOK", "SEK"),
        CADNOK: failExtract("CADNOK", "CAD"),
        EURNOK: failExtract("EURNOK", "EUR"),
        GBPNOK: failExtract("GBPNOK", "GBP"),
        USDNOK: failExtract("USDNOK", "USD"),
      },
      overall_passed: false,
    };
  }

  console.log("Sheet 'Hedging Impact - P&L vs  Act' found. Starting extraction...");

  // ── Constants ──
  const MONTH_ROW  = 2;
  const YEAR_ROW   = 3;
  const START_COL  = 5;   // col E (1-based)
  const END_COL    = 16;  // col P (1-based)
  const NUM_COLS   = END_COL - START_COL + 1; // 12

  const MONTH_LABELS = ["Jan","Feb","Mar","Apr","May","Jun",
                         "Jul","Aug","Sep","Oct","Nov","Dec"];

  // ── Batch read header rows and all data rows in one pass ──
  // Rows needed: 2, 3, 6, 7, 8, 9, 10, 16, 18
  // Read a contiguous block rows 2–18 cols E–P to minimise range calls
  const BLOCK_START_ROW = 2;
  const BLOCK_END_ROW   = 18;
  const BLOCK_NUM_ROWS  = BLOCK_END_ROW - BLOCK_START_ROW + 1; // 17

  const block = ws
    .getRangeByIndexes(BLOCK_START_ROW - 1, START_COL - 1, BLOCK_NUM_ROWS, NUM_COLS)
    .getValues();

  // block[r][c] where r=0 → row 2, c=0 → col E
  const monthRow  = block[0];  // row 2: month numbers 1–12
  const yearRow   = block[1];  // row 3: year (2026)
  const sekRow    = block[4];  // row 6
  const cadRow    = block[5];  // row 7
  const eurRevRow = block[6];  // row 8
  const gbpRow    = block[7];  // row 9
  const usdRevRow = block[8];  // row 10
  const eurCogRow = block[14]; // row 16
  const usdCogRow = block[16]; // row 18

  // ── Header hash (month row + year row concatenated) ──
  let rawHeader = "";
  for (let i = 0; i < NUM_COLS; i++) {
    rawHeader += String(monthRow[i] ?? "") + String(yearRow[i] ?? "");
  }
  let hash = 0;
  for (let i = 0; i < rawHeader.length; i++) {
    hash = (hash * 31 + rawHeader.charCodeAt(i)) & 0xffffffff;
  }
  const headerHash = (hash >>> 0).toString(16).padStart(8, "0");

  // ── ISO end-of-month ──
  function endOfMonth(year: number, month: number): string {
    const d = new Date(Date.UTC(year, month, 0));
    return d.toISOString().slice(0, 10);
  }

  // ── is_estimate: current month or future ──
  const now = new Date();
  const currentYearMonth = now.getFullYear() * 100 + (now.getMonth() + 1);

  // ── Parse a cell value to number | null ──
  function toNum(val: string | number | boolean): number | null {
    if (val === null || val === undefined || val === "") return null;
    const n = typeof val === "number" ? val : Number(val);
    return isNaN(n) ? null : n;
  }

  // ── Build ColMap from header rows ──
  interface ColMap { month: number; year: number; }
  const colMap: ColMap[] = [];
  let headerParseError: string | null = null;

  for (let i = 0; i < NUM_COLS; i++) {
    const month = typeof monthRow[i] === "number" ? monthRow[i] as number : parseInt(String(monthRow[i]), 10);
    const year  = typeof yearRow[i]  === "number" ? yearRow[i]  as number : parseInt(String(yearRow[i]),  10);
    if (isNaN(month) || isNaN(year) || month < 1 || month > 12 || year < 2000) {
      headerParseError = `Could not parse header at column offset ${i}: month="${monthRow[i]}" year="${yearRow[i]}"`;
      break;
    }
    colMap.push({ month, year });
  }

  // ── Build FxExtract for one pair ──
  function buildExtract(
    currencyPair: string,
    exposedCurrency: string,
    rows: { data: (string | number | boolean)[]; transaction_type: string }[]
  ): FxExtract {
    if (headerParseError) {
      return {
        schema_version: "1.0",
        currency_pair: currencyPair,
        exposed_currency: exposedCurrency,
        source_file: sourceFile,
        extracted_at: timestamp,
        extraction_script: scriptName,
        exposures: [],
        validation: {
          header_hash: headerHash,
          period_count: 0,
          row_count: 0,
          null_count: 0,
          passed: false,
          failure_reason: headerParseError,
        },
      };
    }

    const exposures: ExposureRecord[] = [];
    let nullCount = 0;

    // 2026 records — read from source
    for (const { data, transaction_type } of rows) {
      for (let i = 0; i < NUM_COLS; i++) {
        const { month, year } = colMap[i];
        const val = toNum(data[i]);
        if (val === null) nullCount++;
        exposures.push({
          transaction_type,
          period_label: MONTH_LABELS[month - 1],
          period_year: year,
          period_date: endOfMonth(year, month),
          exposure_fc: val,
          is_estimate: year * 100 + month >= currentYearMonth,
        });
      }
    }

    // 2027 records — duplicate 2026 values, all is_estimate=true
    for (const { data, transaction_type } of rows) {
      for (let i = 0; i < NUM_COLS; i++) {
        const { month } = colMap[i];
        const year2027 = 2027;
        const val = toNum(data[i]);
        // null_count not incremented for duplicated nulls — already counted above
        exposures.push({
          transaction_type,
          period_label: MONTH_LABELS[month - 1],
          period_year: year2027,
          period_date: endOfMonth(year2027, month),
          exposure_fc: val,
          is_estimate: true,
        });
      }
    }

    const expectedRecords = 24 * rows.length;
    let failureReason: string | null = null;
    if (exposures.length !== expectedRecords) {
      failureReason = `Record count mismatch: expected ${expectedRecords}, got ${exposures.length}`;
    }

    return {
      schema_version: "1.0",
      currency_pair: currencyPair,
      exposed_currency: exposedCurrency,
      source_file: sourceFile,
      extracted_at: timestamp,
      extraction_script: scriptName,
      exposures,
      validation: {
        header_hash: headerHash,
        period_count: 24,
        row_count: exposures.length,
        null_count: nullCount,
        passed: failureReason === null,
        failure_reason: failureReason,
      },
    };
  }

  // ── Extract all 5 pairs ──
  const seknok = buildExtract("SEKNOK", "SEK", [{ data: sekRow,    transaction_type: "Revenue" }]);
  const cadnok = buildExtract("CADNOK", "CAD", [{ data: cadRow,    transaction_type: "Revenue" }]);
  const gbpnok = buildExtract("GBPNOK", "GBP", [{ data: gbpRow,    transaction_type: "Revenue" }]);
  const eurnok = buildExtract("EURNOK", "EUR", [
    { data: eurRevRow, transaction_type: "Revenue" },
    { data: eurCogRow, transaction_type: "COGS"    },
  ]);
  const usdnok = buildExtract("USDNOK", "USD", [
    { data: usdRevRow, transaction_type: "Revenue" },
    { data: usdCogRow, transaction_type: "COGS"    },
  ]);

  const overall_passed =
    seknok.validation.passed &&
    cadnok.validation.passed &&
    gbpnok.validation.passed &&
    eurnok.validation.passed &&
    usdnok.validation.passed;

  // ── Verification log ──
  console.log("=== EXTRACTION COMPLETE ===");
  console.log(`Source file   : ${sourceFile}`);
  console.log(`Extracted at  : ${timestamp}`);
  console.log(`Header hash   : ${headerHash}`);
  console.log(`Overall passed: ${overall_passed}`);
  console.log("");
  console.log("--- SEKNOK Revenue (Row 6) — first 3 months ---");
  console.log(`  ${seknok.exposures[0]?.period_date}  exposure_fc=${seknok.exposures[0]?.exposure_fc}  is_estimate=${seknok.exposures[0]?.is_estimate}`);
  console.log(`  ${seknok.exposures[1]?.period_date}  exposure_fc=${seknok.exposures[1]?.exposure_fc}  is_estimate=${seknok.exposures[1]?.is_estimate}`);
  console.log(`  ${seknok.exposures[2]?.period_date}  exposure_fc=${seknok.exposures[2]?.exposure_fc}  is_estimate=${seknok.exposures[2]?.is_estimate}`);
  console.log("");
  console.log("--- CADNOK Revenue (Row 7) — first 3 months ---");
  console.log(`  ${cadnok.exposures[0]?.period_date}  exposure_fc=${cadnok.exposures[0]?.exposure_fc}  is_estimate=${cadnok.exposures[0]?.is_estimate}`);
  console.log(`  ${cadnok.exposures[1]?.period_date}  exposure_fc=${cadnok.exposures[1]?.exposure_fc}  is_estimate=${cadnok.exposures[1]?.is_estimate}`);
  console.log(`  ${cadnok.exposures[2]?.period_date}  exposure_fc=${cadnok.exposures[2]?.exposure_fc}  is_estimate=${cadnok.exposures[2]?.is_estimate}`);
  console.log("");
  console.log("--- EURNOK Revenue (Row 8) — first 3 months ---");
  console.log(`  ${eurnok.exposures[0]?.period_date}  exposure_fc=${eurnok.exposures[0]?.exposure_fc}  is_estimate=${eurnok.exposures[0]?.is_estimate}`);
  console.log(`  ${eurnok.exposures[1]?.period_date}  exposure_fc=${eurnok.exposures[1]?.exposure_fc}  is_estimate=${eurnok.exposures[1]?.is_estimate}`);
  console.log(`  ${eurnok.exposures[2]?.period_date}  exposure_fc=${eurnok.exposures[2]?.exposure_fc}  is_estimate=${eurnok.exposures[2]?.is_estimate}`);
  console.log("");
  console.log("--- GBPNOK Revenue (Row 9) — first 3 months ---");
  console.log(`  ${gbpnok.exposures[0]?.period_date}  exposure_fc=${gbpnok.exposures[0]?.exposure_fc}  is_estimate=${gbpnok.exposures[0]?.is_estimate}`);
  console.log(`  ${gbpnok.exposures[1]?.period_date}  exposure_fc=${gbpnok.exposures[1]?.exposure_fc}  is_estimate=${gbpnok.exposures[1]?.is_estimate}`);
  console.log(`  ${gbpnok.exposures[2]?.period_date}  exposure_fc=${gbpnok.exposures[2]?.exposure_fc}  is_estimate=${gbpnok.exposures[2]?.is_estimate}`);
  console.log("");
  console.log("--- USDNOK Revenue (Row 10) — first 3 months ---");
  console.log(`  ${usdnok.exposures[0]?.period_date}  exposure_fc=${usdnok.exposures[0]?.exposure_fc}  is_estimate=${usdnok.exposures[0]?.is_estimate}`);
  console.log(`  ${usdnok.exposures[1]?.period_date}  exposure_fc=${usdnok.exposures[1]?.exposure_fc}  is_estimate=${usdnok.exposures[1]?.is_estimate}`);
  console.log(`  ${usdnok.exposures[2]?.period_date}  exposure_fc=${usdnok.exposures[2]?.exposure_fc}  is_estimate=${usdnok.exposures[2]?.is_estimate}`);
  console.log("");
  console.log("VERIFY: Cross-check Row 6 E-G, Row 7 E-G, Row 8 E-G, Row 9 E-G, Row 10 E-G against console output.");
  console.log(`SEKNOK null_count: ${seknok.validation.null_count}`);
  console.log(`CADNOK null_count: ${cadnok.validation.null_count}`);
  console.log(`EURNOK null_count: ${eurnok.validation.null_count}`);
  console.log(`GBPNOK null_count: ${gbpnok.validation.null_count}`);
  console.log(`USDNOK null_count: ${usdnok.validation.null_count}`);
  if (!overall_passed) {
    console.log(`SEKNOK failure: ${seknok.validation.failure_reason}`);
    console.log(`CADNOK failure: ${cadnok.validation.failure_reason}`);
    console.log(`EURNOK failure: ${eurnok.validation.failure_reason}`);
    console.log(`GBPNOK failure: ${gbpnok.validation.failure_reason}`);
    console.log(`USDNOK failure: ${usdnok.validation.failure_reason}`);
  }

  return {
    schema_version: "1.0",
    source_file: sourceFile,
    extracted_at: timestamp,
    extraction_script: scriptName,
    pairs: { SEKNOK: seknok, CADNOK: cadnok, EURNOK: eurnok, GBPNOK: gbpnok, USDNOK: usdnok },
    overall_passed,
  };
}
