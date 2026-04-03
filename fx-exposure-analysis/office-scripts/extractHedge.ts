/**
 * ═══════════════════════════════════════════════════════════════════════════
 * extractHedge.ts — Office Script  v1.0
 * Runs ON the master FX exposure workbook.
 * Returns canonical hedge JSON schema for all 9 currency pair tabs.
 *
 * VERIFIED HEDGE REFERENCE MAP (April 2026):
 *
 *   EURUSD        : FG notional r49, FG coverage% r51, Roy notional r55, Roy coverage% r57  — Jan col BK(63)
 *   EURPLN        : notional r24, coverage% r26                                              — Jan col BJ(62)
 *   USDMXN        : notional r24, coverage% r26                                              — Jan col BJ(62)
 *   USDCAD        : FG notional r49, FG coverage% r50, IC notional r53, IC coverage% r55    — Jan col BK(63)
 *   SEKNOK        : notional r25, coverage% r27                                              — Jan col N(14)
 *   GBPNOK        : notional r25, coverage% r27                                              — Jan col N(14)
 *   CADNOK_Layers : notional r25, coverage% r27                                              — Jan col N(14)
 *   EURNOK_Layers : notional r25, coverage% r27                                              — Jan col N(14)
 *   USDNOK_Layers : notional r25, coverage% r27                                              — Jan col N(14)
 *
 * All tabs: 24 consecutive months from Jan 2026 starting column.
 *
 * Schema contract : schema_version "1.0"
 * Called by Power Automate processing flow after updateMaster.ts completes.
 * ═══════════════════════════════════════════════════════════════════════════
 */

// ── Interfaces ───────────────────────────────────────────────────────────

interface HedgeRecord {
  currency_pair: string;
  transaction_type: string;
  period_label: string;
  period_year: number;
  period_date: string;        // ISO end-of-month
  hedge_notional: number | null;
  hedge_coverage_pct: number | null;
  is_estimate: boolean;
}

interface HedgeExtract {
  schema_version: string;
  source_file: string;
  extracted_at: string;
  extraction_script: string;
  records: HedgeRecord[];
  record_count: number;
  passed: boolean;
  failure_reason: string | null;
}

// ── Tab config ────────────────────────────────────────────────────────────

interface HedgeRow {
  transaction_type: string;
  notional_row: number;
  coverage_row: number;
}

interface TabConfig {
  tabName: string;
  currency_pair: string;
  janCol: number;
  rows: HedgeRow[];
}

const TAB_CONFIG: TabConfig[] = [
  {
    tabName: "EURUSD", currency_pair: "EURUSD", janCol: 63,
    rows: [
      { transaction_type: "Finished Goods", notional_row: 49, coverage_row: 51 },
      { transaction_type: "Royalties",      notional_row: 55, coverage_row: 57 },
    ],
  },
  {
    tabName: "EURPLN", currency_pair: "EURPLN", janCol: 62,
    rows: [{ transaction_type: "Revenue", notional_row: 24, coverage_row: 26 }],
  },
  {
    tabName: "USDMXN", currency_pair: "USDMXN", janCol: 62,
    rows: [{ transaction_type: "Finished Goods", notional_row: 24, coverage_row: 26 }],
  },
  {
    tabName: "USDCAD", currency_pair: "USDCAD", janCol: 63,
    rows: [
      { transaction_type: "Finished Goods", notional_row: 49, coverage_row: 50 },
      { transaction_type: "I/C Mgmt Fees",  notional_row: 53, coverage_row: 55 },
    ],
  },
  {
    tabName: "SEKNOK", currency_pair: "SEKNOK", janCol: 14,
    rows: [{ transaction_type: "Revenue", notional_row: 25, coverage_row: 27 }],
  },
  {
    tabName: "GBPNOK", currency_pair: "GBPNOK", janCol: 14,
    rows: [{ transaction_type: "Revenue", notional_row: 25, coverage_row: 27 }],
  },
  {
    tabName: "CADNOK_Layers", currency_pair: "CADNOK", janCol: 14,
    rows: [{ transaction_type: "Revenue+COGS", notional_row: 25, coverage_row: 27 }],
  },
  {
    tabName: "EURNOK_Layers", currency_pair: "EURNOK", janCol: 14,
    rows: [{ transaction_type: "Revenue+COGS", notional_row: 25, coverage_row: 27 }],
  },
  {
    tabName: "USDNOK_Layers", currency_pair: "USDNOK", janCol: 14,
    rows: [{ transaction_type: "Revenue+COGS", notional_row: 25, coverage_row: 27 }],
  },
];

// ── 24-month period sequence ──────────────────────────────────────────────

const PERIODS: { year: number; month: number }[] = [];
for (let y = 2026; y <= 2027; y++) {
  for (let m = 1; m <= 12; m++) {
    PERIODS.push({ year: y, month: m });
  }
}

const MONTH_LABELS = ["Jan","Feb","Mar","Apr","May","Jun",
                       "Jul","Aug","Sep","Oct","Nov","Dec"];

// ── Main ─────────────────────────────────────────────────────────────────

function main(workbook: ExcelScript.Workbook): HedgeExtract {
  const timestamp = new Date().toISOString();
  const sourceFile = workbook.getName();
  const scriptName = "extractHedge.ts";

  const now = new Date();
  const currentYearMonth = now.getFullYear() * 100 + (now.getMonth() + 1);

  function endOfMonth(year: number, month: number): string {
    const d = new Date(Date.UTC(year, month, 0));
    return d.toISOString().slice(0, 10);
  }

  function toNum(val: string | number | boolean): number | null {
    if (val === null || val === undefined || val === "") return null;
    const n = typeof val === "number" ? val : Number(val);
    return isNaN(n) ? null : n;
  }

  const allRecords: HedgeRecord[] = [];
  const errors: string[] = [];

  for (const config of TAB_CONFIG) {
    const ws = workbook.getWorksheet(config.tabName);
    if (!ws) {
      errors.push("Tab '" + config.tabName + "' not found");
      console.log("ERROR: Tab '" + config.tabName + "' not found");
      continue;
    }

    for (const hedgeRow of config.rows) {
      // Batch read: notional row and coverage row, 24 cols from janCol
      const notionalValues = ws
        .getRangeByIndexes(hedgeRow.notional_row - 1, config.janCol - 1, 1, 24)
        .getValues()[0];

      const coverageValues = ws
        .getRangeByIndexes(hedgeRow.coverage_row - 1, config.janCol - 1, 1, 24)
        .getValues()[0];

      for (let i = 0; i < 24; i++) {
        const { year, month } = PERIODS[i];
        allRecords.push({
          currency_pair: config.currency_pair,
          transaction_type: hedgeRow.transaction_type,
          period_label: MONTH_LABELS[month - 1],
          period_year: year,
          period_date: endOfMonth(year, month),
          hedge_notional: toNum(notionalValues[i]),
          hedge_coverage_pct: toNum(coverageValues[i]),
          is_estimate: year * 100 + month >= currentYearMonth,
        });
      }
    }
  }

  const passed = errors.length === 0;

  // ── Verification log ──
  console.log("=== HEDGE EXTRACTION COMPLETE ===");
  console.log("Source file   : " + sourceFile);
  console.log("Extracted at  : " + timestamp);
  console.log("Record count  : " + allRecords.length);
  console.log("Passed        : " + passed);
  if (errors.length > 0) {
    console.log("Errors        : " + errors.join(" | "));
  }
  console.log("");
  console.log("--- EURUSD FG hedge (row 49) — first 3 months ---");
  const eurusdFg = allRecords.filter(r => r.currency_pair === "EURUSD" && r.transaction_type === "Finished Goods");
  console.log("  " + eurusdFg[0]?.period_date + "  notional=" + eurusdFg[0]?.hedge_notional + "  coverage%=" + eurusdFg[0]?.hedge_coverage_pct);
  console.log("  " + eurusdFg[1]?.period_date + "  notional=" + eurusdFg[1]?.hedge_notional + "  coverage%=" + eurusdFg[1]?.hedge_coverage_pct);
  console.log("  " + eurusdFg[2]?.period_date + "  notional=" + eurusdFg[2]?.hedge_notional + "  coverage%=" + eurusdFg[2]?.hedge_coverage_pct);
  console.log("");
  console.log("VERIFY: Cross-check BK49, BL49, BM49 and BK51, BL51, BM51 in EURUSD tab.");

  return {
    schema_version: "1.0",
    source_file: sourceFile,
    extracted_at: timestamp,
    extraction_script: scriptName,
    records: allRecords,
    record_count: allRecords.length,
    passed,
    failure_reason: passed ? null : errors.join(" | "),
  };
}
