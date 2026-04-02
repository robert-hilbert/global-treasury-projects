/**
 * ═══════════════════════════════════════════════════════════════════════════
 * extractCanada.ts — Office Script  v1.0
 * Runs ON the Canada Exposure Forecast workbook.
 * Returns canonical FX JSON schema with USDCAD (FG + I/C Mgmt Fees).
 *
 * VERIFIED CELL REFERENCES (April 2026):
 *   Sheet 2026 : "CDNvsUSD 2026"
 *   Sheet 2027 : "CDNvsUSD 2027"
 *   Row 28     = Month labels (both sheets)
 *   Row 32     = Finished Goods (both sheets)
 *   Row 34     = I/C Mgmt Fees (both sheets)
 *
 *   Verified column map (1-based) — irregular due to quarterly summary blocks:
 *   Jan : Forecast=D(4),   Actual=G(7)
 *   Feb : Forecast=H(8),   Actual=K(11)
 *   Mar : Forecast=L(12),  Actual=O(15)
 *   [Q1 summary gap: P–T (16–20)]
 *   Apr : Forecast=U(21),  Actual=X(24)
 *   May : Forecast=Y(25),  Actual=AB(28)
 *   Jun : Forecast=AC(29), Actual=AF(32)
 *   [Q2 summary gap: AG–AK (33–37)]
 *   Jul : Forecast=AL(38), Actual=AO(41)
 *   Aug : Forecast=AP(42), Actual=AS(45)
 *   Sep : Forecast=AT(46), Actual=AW(49)
 *   [Q3 summary gap: AX–BB (50–54)]
 *   Oct : Forecast=BC(55), Actual=BF(58)
 *   Nov : Forecast=BG(59), Actual=BJ(62)
 *   Dec : Forecast=BK(63), Actual=BN(66)
 *
 *   Value rule: use Actual if non-zero, otherwise Forecast.
 *   is_estimate: true if Forecast was used OR period is current/future.
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
  period_date: string;
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

// ── Main ─────────────────────────────────────────────────────────────────

function main(workbook: ExcelScript.Workbook): FxExtract {
  const timestamp = new Date().toISOString();
  const scriptName = "extractCanada.ts";
  const sourceFile = workbook.getName();

  const SHEET_2026 = "CDNvsUSD 2026";
  const SHEET_2027 = "CDNvsUSD 2027";

  // ── Sheet validation ──
  const ws2026 = workbook.getWorksheet(SHEET_2026);
  const ws2027 = workbook.getWorksheet(SHEET_2027);

  if (!ws2026 || !ws2027) {
    const missing = !ws2026 ? SHEET_2026 : SHEET_2027;
    console.log(`ERROR: Sheet '${missing}' not found`);
    return {
      schema_version: "1.0",
      currency_pair: "USDCAD",
      exposed_currency: "CAD",
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
        failure_reason: `Sheet '${missing}' not found`,
      },
    };
  }

  console.log("Both sheets found. Starting extraction...");

  // ── Constants ──
  const HEADER_ROW = 28;
  const FG_ROW     = 32;
  const IC_ROW     = 34;

  const MONTH_LABELS = ["Jan","Feb","Mar","Apr","May","Jun",
                         "Jul","Aug","Sep","Oct","Nov","Dec"];

  // Verified column map (1-based) — hardcoded, no pattern assumption.
  // Quarterly summary blocks are skipped entirely.
  interface MonthCols { forecast: number; actual: number; }
  const COL_MAP: MonthCols[] = [
    { forecast:  4, actual:  7 }, // Jan  D, G
    { forecast:  8, actual: 11 }, // Feb  H, K
    { forecast: 12, actual: 15 }, // Mar  L, O
    { forecast: 21, actual: 24 }, // Apr  U, X
    { forecast: 25, actual: 28 }, // May  Y, AB
    { forecast: 29, actual: 32 }, // Jun  AC, AF
    { forecast: 38, actual: 41 }, // Jul  AL, AO
    { forecast: 42, actual: 45 }, // Aug  AP, AS
    { forecast: 46, actual: 49 }, // Sep  AT, AW
    { forecast: 55, actual: 58 }, // Oct  BC, BF
    { forecast: 59, actual: 62 }, // Nov  BG, BJ
    { forecast: 63, actual: 66 }, // Dec  BK, BN
  ];

  // Read from col D(4) to col BN(66) = 63 columns wide
  const READ_START_COL = 4;  // col D
  const READ_WIDTH     = 63; // D through BN

  // ── Helpers ──

  function toNum(val: string | number | boolean): number | null {
    if (val === null || val === undefined || val === "") return null;
    const n = typeof val === "number" ? val : Number(val);
    return isNaN(n) ? null : n;
  }

  function endOfMonth(year: number, month: number): string {
    const d = new Date(Date.UTC(year, month, 0));
    return d.toISOString().slice(0, 10);
  }

  const now = new Date();
  const currentYearMonth = now.getFullYear() * 100 + (now.getMonth() + 1);

  // ── Batch read: header + data rows in one range call per sheet ──
  // Read rows 28–34 (7 rows), cols D–BN (63 cols)
  function readSheetBlock(ws: ExcelScript.Worksheet): (string | number | boolean)[][] {
    return ws
      .getRangeByIndexes(HEADER_ROW - 1, READ_START_COL - 1, IC_ROW - HEADER_ROW + 1, READ_WIDTH)
      .getValues();
  }

  const block2026 = readSheetBlock(ws2026);
  const block2027 = readSheetBlock(ws2027);

  // Row offsets within block (0-based from row 28)
  const HEADER_OFFSET = 0;                    // row 28
  const FG_OFFSET     = FG_ROW - HEADER_ROW; // row 32 → offset 4
  const IC_OFFSET     = IC_ROW - HEADER_ROW; // row 34 → offset 6

  // ── Header hash — both sheets combined ──
  function hashRow(row: (string | number | boolean)[]): string {
    let raw = "";
    for (let i = 0; i < row.length; i++) raw += String(row[i] ?? "");
    let hash = 0;
    for (let i = 0; i < raw.length; i++) {
      hash = (hash * 31 + raw.charCodeAt(i)) & 0xffffffff;
    }
    return (hash >>> 0).toString(16).padStart(8, "0");
  }

  const headerHash = `${hashRow(block2026[HEADER_OFFSET])}|${hashRow(block2027[HEADER_OFFSET])}`;

  // ── Build exposure records for one sheet/year ──
  function buildRecords(
    block: (string | number | boolean)[][],
    year: number
  ): { records: ExposureRecord[]; nullCount: number } {
    const records: ExposureRecord[] = [];
    let nullCount = 0;

    const fgRow = block[FG_OFFSET];
    const icRow = block[IC_OFFSET];

    const pairs: { data: (string | number | boolean)[]; transaction_type: string }[] = [
      { data: fgRow, transaction_type: "Finished Goods" },
      { data: icRow, transaction_type: "I/C Mgmt Fees"  },
    ];

    for (const { data, transaction_type } of pairs) {
      for (let m = 0; m < COL_MAP.length; m++) {
        // Convert 1-based col index to 0-based offset within block
        const forecastOffset = COL_MAP[m].forecast - READ_START_COL;
        const actualOffset   = COL_MAP[m].actual   - READ_START_COL;

        const actualVal   = toNum(data[actualOffset]);
        const forecastVal = toNum(data[forecastOffset]);

        let exposure_fc: number | null;
        let usedForecast: boolean;

        if (actualVal !== null && actualVal !== 0) {
          exposure_fc  = actualVal;
          usedForecast = false;
        } else {
          exposure_fc  = forecastVal;
          usedForecast = true;
        }

        if (exposure_fc === null) nullCount++;

        const month = m + 1;
        records.push({
          transaction_type,
          period_label: MONTH_LABELS[m],
          period_year: year,
          period_date: endOfMonth(year, month),
          exposure_fc,
          is_estimate: usedForecast || (year * 100 + month >= currentYearMonth),
        });
      }
    }

    return { records, nullCount };
  }

  const result2026 = buildRecords(block2026, 2026);
  const result2027 = buildRecords(block2027, 2027);

  const allExposures = [...result2026.records, ...result2027.records];
  const totalNulls   = result2026.nullCount + result2027.nullCount;

  // ── Validation ──
  const expectedRecords = 24 * 2; // 24 months × 2 transaction types
  let failureReason: string | null = null;
  if (allExposures.length !== expectedRecords) {
    failureReason = `Record count mismatch: expected ${expectedRecords}, got ${allExposures.length}`;
  }

  const passed = failureReason === null;

  // ── Verification log ──
  const fg2026 = allExposures.filter(e => e.transaction_type === "Finished Goods" && e.period_year === 2026);
  const ic2026 = allExposures.filter(e => e.transaction_type === "I/C Mgmt Fees"  && e.period_year === 2026);

  console.log("=== EXTRACTION COMPLETE ===");
  console.log(`Source file   : ${sourceFile}`);
  console.log(`Extracted at  : ${timestamp}`);
  console.log(`Header hash   : ${headerHash}`);
  console.log(`Overall passed: ${passed}`);
  console.log("");
  console.log("--- USDCAD Finished Goods (Row 32) — first 3 months ---");
  console.log(`  ${fg2026[0]?.period_date}  exposure_fc=${fg2026[0]?.exposure_fc}  is_estimate=${fg2026[0]?.is_estimate}`);
  console.log(`  ${fg2026[1]?.period_date}  exposure_fc=${fg2026[1]?.exposure_fc}  is_estimate=${fg2026[1]?.is_estimate}`);
  console.log(`  ${fg2026[2]?.period_date}  exposure_fc=${fg2026[2]?.exposure_fc}  is_estimate=${fg2026[2]?.is_estimate}`);
  console.log("");
  console.log("--- USDCAD I/C Mgmt Fees (Row 34) — first 3 months ---");
  console.log(`  ${ic2026[0]?.period_date}  exposure_fc=${ic2026[0]?.exposure_fc}  is_estimate=${ic2026[0]?.is_estimate}`);
  console.log(`  ${ic2026[1]?.period_date}  exposure_fc=${ic2026[1]?.exposure_fc}  is_estimate=${ic2026[1]?.is_estimate}`);
  console.log(`  ${ic2026[2]?.period_date}  exposure_fc=${ic2026[2]?.exposure_fc}  is_estimate=${ic2026[2]?.is_estimate}`);
  console.log("");
  console.log("VERIFY: Cross-check G32, K32, O32 (Jan-Mar Actuals FG) and G34, K34, O34 (Jan-Mar Actuals I/C) in CDNvsUSD 2026.");
  console.log(`Total null_count : ${totalNulls}`);
  if (!passed) {
    console.log(`FAILED: ${failureReason}`);
  }

  return {
    schema_version: "1.0",
    currency_pair: "USDCAD",
    exposed_currency: "CAD",
    source_file: sourceFile,
    extracted_at: timestamp,
    extraction_script: scriptName,
    exposures: allExposures,
    validation: {
      header_hash: headerHash,
      period_count: 24,
      row_count: allExposures.length,
      null_count: totalNulls,
      passed,
      failure_reason: failureReason,
    },
  };
}
