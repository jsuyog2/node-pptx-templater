# Skill: Native Chart Updating & Dual Excel Synchronization

> **Domain**: DrawingML ChartML (`<c:chartSpace>`) & SpreadsheetML Integration  
> **Target Subsystems**: `ChartManager.js`, `ChartCacheGenerator.js`, `ChartWorkbookUpdater.js`

---

## 1. Dual-Store Architecture

PowerPoint native charts require synchronous updates across two distinct stores:
1. **DrawingML Visual Cache** (`ppt/charts/chartX.xml`): Graphic rendering pipeline.
2. **Embedded Excel Workbook** (`ppt/embeddings/Microsoft_Excel_WorksheetX.xlsx`): Underlying data sheet for "Edit Data in Excel".

Failure to synchronize both causes PowerPoint "Repair Presentation" prompts or mismatched values.

---

## 2. Step-by-Step Update Procedure

### Step 1: Query Target Chart
```javascript
const chartInfo = chartManager.findChart(slideIndex, chartIdOrName);
// Returns: { chartPartName: 'ppt/charts/chart1.xml', rId: 'rId2', workbookPart: 'ppt/embeddings/*.xlsx' }
```

### Step 2: Mutate DrawingML Cache (`ChartCacheGenerator.js`)
- Update series `<c:ser>` in `<c:plotArea>`:
  - Series title: `<c:tx>/<c:strRef>/<c:strCache>/<c:pt idx="0">`
  - Category points: `<c:cat>/<c:strRef>/<c:strCache>` (`<c:ptCount>` and `<c:pt idx="i">`)
  - Value points: `<c:val>/<c:numRef>/<c:numCache>` (`<c:ptCount>` and `<c:pt idx="i">`)

### Step 3: Mutate Embedded Excel Workbook (`ChartWorkbookUpdater.js`)
- Unpack embedded `.xlsx` via JSZip.
- Map data: `A1` empty, Row 1 (`B1..`) series titles, Column A (`A2..`) categories, Grid (`B2..`) values.
- **Table Cleansing (Critical)**: Strip `xl/tables/table*.xml`, delete worksheet `tableParts` relationships, and remove table overrides from workbook `[Content_Types].xml` to avoid stale dimension repair errors.
- Compress back with DEFLATE level 6 and commit to VFS.

### Step 4: Data Labels (`<c:dLbls>`)
- Configure `showVal`, `showCatName`, `showLegendKey`, `showPercent`, and `position` (`bestFit`, `ctr`, `inBase`, `inEnd`, `outEnd`) directly in `<c:dLbls>`.
