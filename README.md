# Farmer's AIde — Excel Add-in (Prototype)

An Excel task pane add-in for Free State (South Africa) commercial maize farmers.
On request, it builds three worksheets in the active workbook — live market data,
an editable farm cost profile, and a breakeven-vs-market-price calculator — so a
farmer can see, at a glance, whether this season is profitable at today's price.

This is a prototype for internal review, not a production release — see "What's
live vs. sample data" below before showing it outside the team.

## What's live vs. sample data

| Data | Status | Why |
|---|---|---|
| ZAR/USD exchange rate | **Live** — fetched at runtime from SARB's published rate via [AllRatesToday](https://allratestoday.com/central-bank-rates-api/sarb/) (free, keyless) | Genuinely free public source, no account needed |
| SAFEX crop prices (white/yellow maize, wheat, sunflower, soybean) | Sample data | SAFEX data is available (e.g. via Grain SA's public page), but has no API and no stated reuse terms — needs a direct conversation with Grain SA or the underlying data vendor before wiring up live, especially for a public-facing version |
| Diesel price (R/litre) | Sample data | Live feeds exist (e.g. FuelPrice API) but require signing up for an API key — intentionally not provisioned without sign-off, since that creates an account tied to the business |

The Market Data sheet labels each section accordingly so this is never ambiguous
when showing the workbook to someone else.

## Project structure

```
manifest.xml              Office Add-in manifest (task pane, Excel/Workbook host)
src/taskpane/
  taskpane.html            Task pane UI: farm profile onboarding form
  taskpane.css             Styling — Farmer's AIde brand tokens (colors, type, radii)
  taskpane.ts               All logic: reads the form, fetches live FX, builds sheets
assets/                    App icons (16/32/64/80/128) derived from the brand emblem
webpack.config.js         Dev server (HTTPS, port 3000) + production build
```

Brand colors, type (Poppins/Inter), and the green-profit/terracotta-loss semantic
are taken directly from the original Farmer's AIde brand system
(`farmers-aide/docs/engineering/BRAND-THEME.md`), not invented fresh.

## Running it locally

```bash
npm install
npm run certs    # one-time: trust the local HTTPS dev cert if prompted
npm run dev-server    # serves the task pane at https://localhost:3000
```

Leave the dev server running — Excel loads the task pane from `localhost:3000`
while you're testing.

## Loading it into Excel (macOS)

The add-in isn't published anywhere yet, so it has to be sideloaded. Two options:

### Option A — Excel on the web (most reliable for a quick demo)

1. With the dev server running, go to [office.com](https://office.com) and open a
   new or existing workbook in Excel for the web.
2. **Insert** tab → **Add-ins** → **Upload My Add-in**.
3. Choose `manifest.xml` from this project folder.
4. A "FarmersAIde" button appears on the Home tab — click it to open the task pane.

### Option B — Excel desktop (Mac)

Desktop sideloading on Mac uses a shared manifest folder:

1. Copy `manifest.xml` into
   `~/Library/Containers/com.microsoft.Excel/Data/Documents/wef`
   (create the `wef` folder if it doesn't exist).
2. Quit and reopen Excel.
3. **Insert** tab → **Add-ins** → **My Add-ins** → the add-in should appear under
   "Developer Add-ins".

If it doesn't show up, Option A is the faster fallback for a demo — desktop
sideloading can be inconsistent across Excel/macOS versions.

## Using it

1. Open the task pane (see above).
2. Fill in the farm profile: district, primary crop, rotation crops, livestock,
   hectares planted.
3. Click **Pull Market Data & Build Sheets**.
4. Three sheets appear: **Market Data**, **Farm Profile & Costs** (edit the cost
   figures to match a real farm), and **Breakeven vs Market Price** — the last
   one recalculates live as you edit costs, and is color-coded green (profitable)
   or terracotta (loss) based on the current sample market price.

## Known gaps / next steps

- SAFEX and diesel feeds need a data-sourcing decision (partnership vs. API key)
  before this can show real prices.
- No automatic "run on workbook open" yet — this prototype is button-triggered
  for predictability during review; the original product vision has it firing
  automatically when Excel opens.
- Sell-now-vs-store and FX input-cost exposure calculators are scoped but not
  built — this prototype covers the MVP cut only (Market Data + Farm Profile +
  Breakeven calculator).
