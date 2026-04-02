/**
 * ═══════════════════════════════════════════════════════════════════════════
 * updateMaster.ts — Office Script  v1.0  [PRODUCTION]
 * Runs ON the master FX exposure workbook.
 * Receives combined JSON from all four extraction scripts via Power Automate
 * and writes exposure values to verified destination cells.
 *
 * VERIFIED DESTINATION MAP (April 2026):
 *   EURUSD        : FG → row 7, Royalties → row 11,  Jan 2026 = col BK(63)
 *   EURPLN        : Revenue → row 7,                  Jan 2026 = col BJ(62)
 *   USDMXN        : Finished Goods → row 6,           Jan 2026 = col BJ(62)
 *   USDCAD        : FG → row 7, I/C Mgmt Fees → row 9, Jan 2026 = col BK(63)
 *   SEKNOK        : Revenue → row 7,                  Jan 2026 = col N(14)
 *   GBPNOK        : Revenue → row 7,                  Jan 2026 = col N(14)
 *   CADNOK_Layers : Revenue+COGS summed → row 7,      Jan 2026 = col N(14)
 *   EURNOK_Layers : Revenue+COGS summed → row 7,      Jan 2026 = col N(14)
 *   USDNOK_Layers : Revenue+COGS summed → row 7,      Jan 2026 = col N(14)
 *
 * All tabs: 24 consecutive months from Jan 2026 starting column.
 *
 * Hard constraints:
 *   - Never writes if any extraction script passed=false (reconciliation gate)
 *   - Never overwrites formula rows
 *   - _De des tabs not touched in this version
 *   - Returns JSON result string for Power Automate condition branching
 *
 * Schema contract : schema_version "1.0"
 * Called by Power Automate after all four extraction scripts complete.
 * ═══════════════════════════════════════════════════════════════════════════
 */

// ── Interfaces ───────────────────────────────────────────────────────────

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

interface EuropeanExtract {
  schema_version: string;
  source_file: string;
  extracted_at: string;
  extraction_script: string;
  pairs: { EURUSD: FxExtract; EURPLN: FxExtract };
  overall_passed: boolean;
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

interface CombinedPayload {
  european: EuropeanExtract;
  mexico:   FxExtract;
  norway:   NorwayExtract;
  canada:   FxExtract;
}

// ── Destination map ───────────────────────────────────────────────────────

interface DestinationRow {
  transaction_type: string | null; // null = sum all transaction types
  row: number;
}

interface TabConfig {
  tabName: string;
  janCol: number;         // 1-based column index for Jan 2026
  destinations: DestinationRow[];
}

const TAB_CONFIG: TabConfig[] = [
  {
    tabName: "EURUSD", janCol: 63,
    destinations: [
      { transaction_type: "Finished Goods", row: 7  },
      { transaction_type: "Royalties",      row: 11 },
    ],
  },
  {
    tabName: "EURPLN", janCol: 62,
    destinations: [{ transaction_type: "Revenue", row: 7 }],
  },
  {
    tabName: "USDMXN", janCol: 62,
    destinations: [{ transaction_type: "Finished Goods", row: 6 }],
  },
  {
    tabName: "USDCAD", janCol: 63,
    destinations: [
      { transaction_type: "Finished Goods", row: 7 },
      { transaction_type: "I/C Mgmt Fees",  row: 9 },
    ],
  },
  {
    tabName: "SEKNOK", janCol: 14,
    destinations: [{ transaction_type: "Revenue", row: 7 }],
  },
  {
    tabName: "GBPNOK", janCol: 14,
    destinations: [{ transaction_type: "Revenue", row: 7 }],
  },
  {
    tabName: "CADNOK_Layers", janCol: 14,
    destinations: [{ transaction_type: null, row: 7 }],
  },
  {
    tabName: "EURNOK_Layers", janCol: 14,
    destinations: [{ transaction_type: null, row: 7 }],
  },
  {
    tabName: "USDNOK_Layers", janCol: 14,
    destinations: [{ transaction_type: null, row: 7 }],
  },
];

// ── Tab to pair key lookup ────────────────────────────────────────────────

const TAB_TO_PAIR: { [key: string]: string } = {
  "EURUSD":        "EURUSD",
  "EURPLN":        "EURPLN",
  "USDMXN":        "USDMXN",
  "USDCAD":        "USDCAD",
  "SEKNOK":        "SEKNOK",
  "GBPNOK":        "GBPNOK",
  "CADNOK_Layers": "CADNOK",
  "EURNOK_Layers": "EURNOK",
  "USDNOK_Layers": "USDNOK",
};

// ── 24-month period sequence ──────────────────────────────────────────────

const PERIODS: { year: number; month: number }[] = [];
for (let y = 2026; y <= 2027; y++) {
  for (let m = 1; m <= 12; m++) {
    PERIODS.push({ year: y, month: m });
  }
}

// ── Main ─────────────────────────────────────────────────────────────────

function main(workbook: ExcelScript.Workbook, payloadJson: string): string {
  const timestamp = new Date().toISOString();

  // ── Parse payload ──
  let payload: CombinedPayload;
  try {
    payload = JSON.parse(payloadJson) as CombinedPayload;
  } catch (e) {
    const msg = "FAILED: Could not parse payloadJson — check Compose step output in Power Automate";
    console.log(msg);
    return JSON.stringify({ success: false, error: msg, tabs_written: 0, tabs_failed: 0, errors: [msg], timestamp });
  }

  // ── Reconciliation gate — all extraction scripts must have passed ──
  const gates = [
    { name: "European", passed: payload.european?.overall_passed },
    { name: "Mexico",   passed: payload.mexico?.validation?.passed },
    { name: "Norway",   passed: payload.norway?.overall_passed },
    { name: "Canada",   passed: payload.canada?.validation?.passed },
  ];

  const failedGates = gates.filter(g => !g.passed).map(g => g.name);
  if (failedGates.length > 0) {
    const msg = "FAILED: Reconciliation gate blocked write. Failed scripts: " + failedGates.join(", ");
    console.log(msg);
    return JSON.stringify({ success: false, error: msg, tabs_written: 0, tabs_failed: 0, errors: [msg], timestamp });
  }

  console.log("Reconciliation gate passed. All extraction scripts verified.");
  console.log("Starting master workbook update...");

  // ── Flatten all pairs into lookup map ──
  const pairData: { [key: string]: ExposureRecord[] } = {
    EURUSD: payload.european.pairs.EURUSD.exposures,
    EURPLN: payload.european.pairs.EURPLN.exposures,
    USDMXN: payload.mexico.exposures,
    USDCAD: payload.canada.exposures,
    SEKNOK: payload.norway.pairs.SEKNOK.exposures,
    GBPNOK: payload.norway.pairs.GBPNOK.exposures,
    CADNOK: payload.norway.pairs.CADNOK.exposures,
    EURNOK: payload.norway.pairs.EURNOK.exposures,
    USDNOK: payload.norway.pairs.USDNOK.exposures,
  };

  // ── Get summed exposure value for one period ──
  function getValue(
    records: ExposureRecord[],
    year: number,
    month: number,
    transaction_type: string | null
  ): number {
    const matches = records.filter(r => {
      const d = new Date(r.period_date);
      const periodMatch = r.period_year === year && d.getUTCMonth() + 1 === month;
      return transaction_type === null
        ? periodMatch
        : periodMatch && r.transaction_type === transaction_type;
    });
    return matches.reduce((sum, r) => sum + (r.exposure_fc ?? 0), 0);
  }

  // ── Write to each tab ──
  let tabsWritten = 0;
  let tabsFailed  = 0;
  const errors: string[] = [];

  for (const config of TAB_CONFIG) {
    const ws = workbook.getWorksheet(config.tabName);
    if (!ws) {
      const msg = "Tab '" + config.tabName + "' not found — skipping";
      console.log(msg);
      errors.push(msg);
      tabsFailed++;
      continue;
    }

    const pairKey = TAB_TO_PAIR[config.tabName];
    const records = pairData[pairKey];

    if (!records || records.length === 0) {
      const msg = "No records for pair " + pairKey + " (tab " + config.tabName + ") — skipping";
      console.log(msg);
      errors.push(msg);
      tabsFailed++;
      continue;
    }

    for (const dest of config.destinations) {
      const values: number[] = PERIODS.map(p =>
        getValue(records, p.year, p.month, dest.transaction_type)
      );
      ws.getRangeByIndexes(dest.row - 1, config.janCol - 1, 1, 24).setValues([values]);
      const label = dest.transaction_type ?? "Revenue+COGS";
      console.log("  Written: " + config.tabName + " row " + dest.row + " (" + label + ") — 24 months");
    }

    tabsWritten++;
  }

  // ── Summary ──
  console.log("");
  console.log("=== UPDATE COMPLETE ===");
  console.log("Tabs written : " + tabsWritten);
  console.log("Tabs failed  : " + tabsFailed);
  console.log("Completed at : " + timestamp);
  if (errors.length > 0) {
    console.log("Errors       : " + errors.join(" | "));
  }

  return JSON.stringify({
    success: tabsFailed === 0,
    tabs_written: tabsWritten,
    tabs_failed: tabsFailed,
    errors,
    timestamp,
  });
}
