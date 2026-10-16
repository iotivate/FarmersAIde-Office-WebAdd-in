/* FarmersAIde task pane logic.
 * Builds three worksheets (Market Data, Farm Profile & Costs, Breakeven vs Market Price)
 * from the farmer's onboarding profile. ZAR/USD is fetched live; SAFEX crop prices and
 * diesel are sample data until a licensed feed / API key is in place (see README.md).
 */

interface FarmProfile {
  district: string;
  primaryCrop: string;
  rotationCrops: string[];
  livestock: "yes" | "no";
  hectares: number;
}

interface FxResult {
  value: number | null;
  rateDate: string;
  fetchedAt: string;
  error?: string;
}

const SARB_FX_URL = "https://allratestoday.com/api/open/central-bank/sarb";

const SAMPLE_CROP_PRICES: Array<{ crop: string; priceRPerTon: number }> = [
  { crop: "White Maize", priceRPerTon: 4036 },
  { crop: "Yellow Maize", priceRPerTon: 4060 },
  { crop: "Wheat", priceRPerTon: 6250 },
  { crop: "Sunflower", priceRPerTon: 9800 },
  { crop: "Soybean", priceRPerTon: 8700 },
];

const SAMPLE_DIESEL_PRICE_R_PER_LITRE = 23.5;

Office.onReady(() => {
  const button = document.getElementById("run-button") as HTMLButtonElement;
  button.addEventListener("click", onRunButtonClicked);
});

function setStatus(message: string, kind: "info" | "success" | "error" = "info"): void {
  const statusEl = document.getElementById("status") as HTMLDivElement;
  statusEl.textContent = message;
  statusEl.className = `status ${kind === "info" ? "" : kind}`.trim();
}

function readProfileFromForm(): FarmProfile {
  const districtEl = document.getElementById("district") as HTMLSelectElement;
  const hectaresEl = document.getElementById("hectares") as HTMLInputElement;
  const primaryCropEl = document.querySelector(
    'input[name="primaryCrop"]:checked'
  ) as HTMLInputElement;
  const livestockEl = document.querySelector(
    'input[name="livestock"]:checked'
  ) as HTMLInputElement;

  const rotationCrops: string[] = [];
  (["rotSunflower", "rotSoybean", "rotWheat"] as const).forEach((id) => {
    const el = document.getElementById(id) as HTMLInputElement;
    if (el.checked) rotationCrops.push(el.value);
  });

  return {
    district: districtEl.value,
    primaryCrop: primaryCropEl.value,
    rotationCrops,
    livestock: livestockEl.value as "yes" | "no",
    hectares: Number(hectaresEl.value) || 1,
  };
}

async function fetchZarUsdRate(): Promise<FxResult> {
  const fetchedAt = new Date().toISOString();
  try {
    const response = await fetch(SARB_FX_URL, { method: "GET" });
    if (!response.ok) {
      throw new Error(`HTTP ${response.status}`);
    }
    const data = await response.json();
    const usdZar = (data.rates as Array<{ base: string; quote: string; value: number }>).find(
      (r) => r.base === "USD" && r.quote === "ZAR"
    );
    if (!usdZar) {
      throw new Error("USD/ZAR rate not found in response");
    }
    return {
      value: usdZar.value,
      rateDate: data.rate_date ?? "unknown",
      fetchedAt,
    };
  } catch (err) {
    return {
      value: null,
      rateDate: "n/a",
      fetchedAt,
      error: err instanceof Error ? err.message : "FX fetch failed — check connection",
    };
  }
}

async function onRunButtonClicked(): Promise<void> {
  const button = document.getElementById("run-button") as HTMLButtonElement;
  button.disabled = true;
  setStatus("Fetching live ZAR/USD rate…");

  try {
    const profile = readProfileFromForm();
    const fx = await fetchZarUsdRate();

    if (fx.error) {
      setStatus(`FX fetch failed (${fx.error}) — continuing with sample data only.`, "error");
    } else {
      setStatus("Live FX rate retrieved. Building worksheets…");
    }

    await buildWorkbook(profile, fx);

    setStatus(
      fx.error
        ? "Worksheets built. Note: ZAR/USD fetch failed — see Market Data sheet."
        : "Worksheets built successfully. Live ZAR/USD rate applied.",
      fx.error ? "error" : "success"
    );
  } catch (err) {
    setStatus(
      `Something went wrong building the worksheets: ${
        err instanceof Error ? err.message : String(err)
      }`,
      "error"
    );
  } finally {
    button.disabled = false;
  }
}

async function getOrCreateSheet(
  context: Excel.RequestContext,
  name: string
): Promise<Excel.Worksheet> {
  const sheets = context.workbook.worksheets;
  sheets.load("items/name");
  await context.sync();

  const existing = sheets.items.find((s) => s.name === name);
  if (existing) {
    existing.delete();
    await context.sync();
  }

  const sheet = sheets.add(name);
  await context.sync();
  return sheet;
}

async function buildWorkbook(profile: FarmProfile, fx: FxResult): Promise<void> {
  await Excel.run(async (context) => {
    const marketSheet = await getOrCreateSheet(context, "Market Data");
    const profileSheet = await getOrCreateSheet(context, "Farm Profile & Costs");
    const breakevenSheet = await getOrCreateSheet(context, "Breakeven vs Market Price");

    buildMarketDataSheet(marketSheet, fx);
    buildFarmProfileSheet(profileSheet, profile);
    buildBreakevenSheet(breakevenSheet);

    breakevenSheet.activate();
    await context.sync();
  });
}

function setCell(sheet: Excel.Worksheet, address: string, value: string | number): void {
  sheet.getRange(address).values = [[value]];
}

function setBold(sheet: Excel.Worksheet, address: string): void {
  sheet.getRange(address).format.font.bold = true;
}

function buildMarketDataSheet(sheet: Excel.Worksheet, fx: FxResult): void {
  setCell(sheet, "A1", "FarmersAIde — Market Data");
  setBold(sheet, "A1");
  sheet.getRange("A1").format.font.size = 14;

  setCell(sheet, "A3", "SAFEX Settlement Prices (R/ton)");
  setBold(sheet, "A3");
  setCell(
    sheet,
    "A4",
    "Sample data — live SAFEX feed integration pending data partnership (see README)."
  );
  sheet.getRange("A4").format.font.italic = true;

  setCell(sheet, "A6", "Crop");
  setCell(sheet, "B6", "Price (R/ton)");
  setBold(sheet, "A6:B6");

  SAMPLE_CROP_PRICES.forEach((row, i) => {
    const r = 7 + i;
    setCell(sheet, `A${r}`, row.crop);
    setCell(sheet, `B${r}`, row.priceRPerTon);
    sheet.getRange(`B${r}`).numberFormat = [['"R" #,##0.00']];
  });

  setCell(sheet, "A13", "Diesel Price (R/litre)");
  setBold(sheet, "A13");
  setCell(
    sheet,
    "A14",
    "Sample data — live feed requires FuelPrice API key (not yet provisioned)."
  );
  sheet.getRange("A14").format.font.italic = true;
  setCell(sheet, "A15", "Diesel (50ppm)");
  setCell(sheet, "B15", SAMPLE_DIESEL_PRICE_R_PER_LITRE);
  sheet.getRange("B15").numberFormat = [['"R" #,##0.00']];

  setCell(sheet, "A17", "ZAR / USD Exchange Rate (LIVE)");
  setBold(sheet, "A17");
  setCell(sheet, "A18", "Rate (1 USD = x ZAR)");
  setCell(sheet, "B18", fx.value ?? "n/a — fetch failed");
  if (fx.value !== null) {
    sheet.getRange("B18").numberFormat = [["0.0000"]];
  }
  setCell(sheet, "A19", "SARB rate date");
  setCell(sheet, "B19", fx.rateDate);
  setCell(sheet, "A20", "Fetched at");
  setCell(sheet, "B20", fx.fetchedAt);
  if (fx.error) {
    setCell(sheet, "A21", `FX fetch failed — check connection (${fx.error})`);
    sheet.getRange("A21").format.font.color = "#C0362C"; // brand error (Terracotta 600)
  } else {
    setCell(
      sheet,
      "A21",
      "Source: Official rates published by SARB, served by AllRatesToday — https://allratestoday.com/central-bank-rates-api/sarb/"
    );
    sheet.getRange("A21").format.font.italic = true;
  }

  sheet.getRange("A1:B21").format.autofitColumns();
}

const FARM_PROFILE_COST_ROWS = {
  seed: 10,
  fertilizer: 11,
  chemicals: 12,
  labor: 13,
  dieselUsage: 14,
  yield: 15,
  storageCost: 16,
};

function buildFarmProfileSheet(sheet: Excel.Worksheet, profile: FarmProfile): void {
  setCell(sheet, "A1", "FarmersAIde — Farm Profile & Costs");
  setBold(sheet, "A1");
  sheet.getRange("A1").format.font.size = 14;

  setCell(sheet, "A3", "District");
  setCell(sheet, "B3", profile.district);
  setCell(sheet, "A4", "Primary Crop");
  setCell(sheet, "B4", profile.primaryCrop);
  setCell(sheet, "A5", "Rotation Crops");
  setCell(sheet, "B5", profile.rotationCrops.length ? profile.rotationCrops.join(", ") : "None");
  setCell(sheet, "A6", "Livestock");
  setCell(sheet, "B6", profile.livestock === "yes" ? "Yes" : "No");
  setCell(sheet, "A7", "Hectares Planted");
  setCell(sheet, "B7", profile.hectares);

  setCell(sheet, "A9", "Cost Inputs (editable — enter your own figures)");
  setBold(sheet, "A9");

  const costRows: Array<[number, string, number, string?]> = [
    [FARM_PROFILE_COST_ROWS.seed, "Seed (R/ha)", 1850, '"R" #,##0.00'],
    [FARM_PROFILE_COST_ROWS.fertilizer, "Fertilizer (R/ha)", 3200, '"R" #,##0.00'],
    [FARM_PROFILE_COST_ROWS.chemicals, "Chemicals — herbicide/pesticide (R/ha)", 950, '"R" #,##0.00'],
    [FARM_PROFILE_COST_ROWS.labor, "Labor (R/ha)", 650, '"R" #,##0.00'],
    [FARM_PROFILE_COST_ROWS.dieselUsage, "Diesel usage (litres/ha)", 45],
    [FARM_PROFILE_COST_ROWS.yield, "Expected yield (ton/ha)", 5.5],
    [FARM_PROFILE_COST_ROWS.storageCost, "Storage cost (R/ton/month)", 35, '"R" #,##0.00'],
  ];

  costRows.forEach(([row, label, value, format]) => {
    setCell(sheet, `A${row}`, label);
    setCell(sheet, `B${row}`, value);
    if (format) {
      sheet.getRange(`B${row}`).numberFormat = [[format]];
    }
  });

  sheet.getRange("A1:B16").format.autofitColumns();
}

function buildBreakevenSheet(sheet: Excel.Worksheet): void {
  const P = FARM_PROFILE_COST_ROWS;
  const money = '"R" #,##0.00';

  setCell(sheet, "A1", "FarmersAIde — Breakeven vs Market Price");
  setBold(sheet, "A1");
  sheet.getRange("A1").format.font.size = 14;

  setCell(sheet, "A3", "Primary Crop");
  sheet.getRange("B3").formulas = [["='Farm Profile & Costs'!B4"]];
  setCell(sheet, "A4", "Hectares Planted");
  sheet.getRange("B4").formulas = [["='Farm Profile & Costs'!B7"]];

  setCell(sheet, "A6", "Cost per Hectare (R/ha)");
  setBold(sheet, "A6");

  setCell(sheet, "A7", "Seed");
  sheet.getRange("B7").formulas = [[`='Farm Profile & Costs'!B${P.seed}`]];
  setCell(sheet, "A8", "Fertilizer");
  sheet.getRange("B8").formulas = [[`='Farm Profile & Costs'!B${P.fertilizer}`]];
  setCell(sheet, "A9", "Chemicals");
  sheet.getRange("B9").formulas = [[`='Farm Profile & Costs'!B${P.chemicals}`]];
  setCell(sheet, "A10", "Labor");
  sheet.getRange("B10").formulas = [[`='Farm Profile & Costs'!B${P.labor}`]];
  setCell(sheet, "A11", "Diesel (price × litres/ha)");
  sheet.getRange("B11").formulas = [
    [`='Market Data'!B15 * 'Farm Profile & Costs'!B${P.dieselUsage}`],
  ];
  setCell(sheet, "A12", "Total Cost / ha");
  sheet.getRange("B12").formulas = [["=SUM(B7:B11)"]];
  setBold(sheet, "A12:B12");

  ["B7", "B8", "B9", "B10", "B11", "B12"].forEach((addr) => {
    sheet.getRange(addr).numberFormat = [[money]];
  });

  setCell(sheet, "A14", "Expected Yield (ton/ha)");
  sheet.getRange("B14").formulas = [[`='Farm Profile & Costs'!B${P.yield}`]];

  setCell(sheet, "A15", "Breakeven Price (R/ton)");
  sheet.getRange("B15").formulas = [["=B12/B14"]];
  sheet.getRange("B15").numberFormat = [[money]];
  setBold(sheet, "A15:B15");

  setCell(sheet, "A17", "Current Market Price (R/ton)");
  sheet.getRange("B17").formulas = [
    ["=INDEX('Market Data'!$B$7:$B$11,MATCH(B3,'Market Data'!$A$7:$A$11,0))"],
  ];
  sheet.getRange("B17").numberFormat = [[money]];

  setCell(sheet, "A19", "Margin per ton (R/ton)");
  sheet.getRange("B19").formulas = [["=B17-B15"]];
  sheet.getRange("B19").numberFormat = [[money]];

  setCell(sheet, "A20", "Total Margin, all hectares (R)");
  sheet.getRange("B20").formulas = [["=B19*(B14*B4)"]];
  sheet.getRange("B20").numberFormat = [[money]];

  setCell(sheet, "A22", "Status");
  sheet.getRange("B22").formulas = [
    ['=IF(B19>=0,"PROFITABLE AT CURRENT PRICE","LOSS AT CURRENT PRICE")'],
  ];
  setBold(sheet, "A22:B22");

  applyProfitLossNumericFormat(sheet, "B19:B20");
  applyProfitLossTextFormat(sheet, "B22");

  sheet.getRange("A1:B22").format.autofitColumns();
}

// Brand semantic rule (BRAND-THEME.md): green = market up / profitable,
// terracotta = market down / loss. Fill/font pairs are the exact
// primaryContainer/onPrimaryContainer and errorContainer/onErrorContainer
// tokens from the brand's light ColorScheme.
const BRAND_PROFIT_FILL = "#C7EBC2"; // primaryContainer
const BRAND_PROFIT_FONT = "#072B0A"; // onPrimaryContainer
const BRAND_LOSS_FILL = "#FFDAD5"; // errorContainer
const BRAND_LOSS_FONT = "#410001"; // onErrorContainer

function applyProfitLossNumericFormat(sheet: Excel.Worksheet, address: string): void {
  const range = sheet.getRange(address);

  const positive = range.conditionalFormats.add(Excel.ConditionalFormatType.cellValue);
  positive.cellValue.format.fill.color = BRAND_PROFIT_FILL;
  positive.cellValue.format.font.color = BRAND_PROFIT_FONT;
  positive.cellValue.rule = {
    formula1: "0",
    operator: Excel.ConditionalCellValueOperator.greaterThanOrEqual,
  };

  const negative = range.conditionalFormats.add(Excel.ConditionalFormatType.cellValue);
  negative.cellValue.format.fill.color = BRAND_LOSS_FILL;
  negative.cellValue.format.font.color = BRAND_LOSS_FONT;
  negative.cellValue.rule = {
    formula1: "0",
    operator: Excel.ConditionalCellValueOperator.lessThan,
  };
}

function applyProfitLossTextFormat(sheet: Excel.Worksheet, address: string): void {
  const range = sheet.getRange(address);

  const positive = range.conditionalFormats.add(Excel.ConditionalFormatType.containsText);
  positive.textComparison.format.fill.color = BRAND_PROFIT_FILL;
  positive.textComparison.format.font.color = BRAND_PROFIT_FONT;
  positive.textComparison.rule = {
    operator: Excel.ConditionalTextOperator.contains,
    text: "PROFITABLE",
  };

  const negative = range.conditionalFormats.add(Excel.ConditionalFormatType.containsText);
  negative.textComparison.format.fill.color = BRAND_LOSS_FILL;
  negative.textComparison.format.font.color = BRAND_LOSS_FONT;
  negative.textComparison.rule = {
    operator: Excel.ConditionalTextOperator.contains,
    text: "LOSS",
  };
}
