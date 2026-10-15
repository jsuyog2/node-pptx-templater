# Skill: Native Chart Updating & Dual Excel Synchronization

> **Domain**: DrawingML ChartML (`<c:chartSpace>`) & SpreadsheetML Integration  
> **Target Subsystems**: `ChartManager.js`, `ChartCacheGenerator.js`, `ChartWorkbookUpdater.js`

---

## 1. Context & Architecture

Native PowerPoint charts do not store series data in a single location. They operate on a dual-store architecture:
1. **The DrawingML XML Cache (`ppt/charts/chartX.xml`)**:
   Used by PowerPoint's graphics pipeline to render the chart immediately when opening the deck.
2. **The Embedded Excel Workbook (`ppt/embeddings/Microsoft_Excel_WorksheetX.xlsx`)**:
   Used when the user clicks "Edit Data in PowerPoint" or "Edit Data in Excel".

If the DrawingML cache and the embedded `.xlsx` workbook become desynchronized, PowerPoint will display mismatched figures or prompt the user with a repair error dialog.

```
                    ┌─────────────────────────┐
                    │    ppt.updateChart()    │
                    └───────────┬─────────────┘
                                │
        ┌───────────────────────┴───────────────────────┐
        ▼                                               ▼
┌──────────────────────────────┐        ┌──────────────────────────────┐
│  DrawingML Cache Sync        │        │ Embedded XLSX Workbook Sync  │
│  (ppt/charts/chart1.xml)     │        │ (ppt/embeddings/*.xlsx)      │
│  • <c:cat> categories cache  │        │ • xl/worksheets/sheet1.xml   │
│  • <c:val> series points     │        │ • xl/sharedStrings.xml       │
│  • <c:dLbls> data labels     │        │ • Remove xl/tables/ conflicts│
└──────────────────────────────┘        └──────────────────────────────┘
```

---

## 2. Playbook: Step-by-Step Chart Update Procedure

### Step 1: Locate the Target Chart
Query the slide to locate the chart by its element name or non-visual ID:
```javascript
const chartInfo = chartManager.findChart(slideIndex, chartIdOrName);
// Returns: { chartPartName: 'ppt/charts/chart1.xml', rId: 'rId2', workbookPart: 'ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx' }
```

### Step 2: Update the `<c:chartSpace>` DrawingML Cache
Use `ChartCacheGenerator` to rewrite the category and value nodes:
```javascript
// 1. Locate plot area <c:plotArea> and target chart type element (<c:barChart>, <c:lineChart>, etc.)
const chartTypeNode = plotArea[chartTypeName];
const seriesNodes = Array.isArray(chartTypeNode['c:ser']) ? chartTypeNode['c:ser'] : [chartTypeNode['c:ser']];

// 2. Iterate through series and categories
for (let sIdx = 0; sIdx < newSeries.length; sIdx++) {
  const ser = seriesNodes[sIdx];
  const sData = newSeries[sIdx];

  // Set series title in <c:tx>/<c:strRef>
  ser['c:tx']['c:strRef']['c:strCache']['c:pt'] = [{ '@_idx': '0', 'c:v': sData.name }];

  // Set category references in <c:cat>/<c:strRef>
  const catPoints = newCategories.map((cat, idx) => ({
    '@_idx': String(idx),
    'c:v': String(cat)
  }));
  ser['c:cat']['c:strRef']['c:strCache']['c:pt'] = catPoints;
  ser['c:cat']['c:strRef']['c:strCache']['c:ptCount'] = { '@_val': String(newCategories.length) };

  // Set value points in <c:val>/<c:numRef>
  const numPoints = sData.values.map((val, idx) => ({
    '@_idx': String(idx),
    'c:v': String(val)
  }));
  ser['c:val']['c:numRef']['c:numCache']['c:pt'] = numPoints;
  ser['c:val']['c:numRef']['c:numCache']['c:ptCount'] = { '@_val': String(sData.values.length) };
}
```

### Step 3: Synchronize the Embedded Excel Workbook
Pass the raw workbook binary buffer to `ChartWorkbookUpdater.updateWorkbook`:
```javascript
const updatedXlsxBuffer = await ChartWorkbookUpdater.updateWorkbook(rawXlsxBuffer, {
  categories: newCategories,
  series: newSeries
});
vfs.writeBytes(workbookPartName, updatedXlsxBuffer);
```

Within `ChartWorkbookUpdater`:
1. **Decompress Excel Archive**: Load with JSZip.
2. **Re-populate Cell Grid**:
   - `A1`: Empty header.
   - Row 1 (`B1`, `C1`, `D1`...): Series names.
   - Column A (`A2`, `A3`, `A4`...): Category names.
   - Grid (`B2`, `C2`, etc.): Numeric series point values.
3. **Table Cleansing (Critical)**:
   - When updating data row/column counts, any pre-existing Excel table objects (`xl/tables/table1.xml`) will have stale dimensions (`ref="A1:C5"`). If row counts expand or shrink, Excel throws a repair prompt.
   - **Fix**: Remove `xl/tables/` files, delete worksheet `tableParts` relationships, and strip `<Override PartName="/xl/tables/..." />` from the embedded workbook's `[Content_Types].xml`.

---

## 3. Data Labels Configuration

To toggle or customize data labels on charts:
```javascript
ppt.useSlide(1).updateDataLabels('SalesChart', {
  showVal: true,        // Show point numeric value
  showCatName: false,   // Hide category name
  showLegendKey: false, // Hide legend key
  showPercent: false,   // Show percentage (for pie/doughnut)
  position: 'outEnd'    // 'bestFit' | 'b' | 'ctr' | 'inBase' | 'inEnd' | 'l' | 'r' | 't' | 'outEnd'
});
```
This mutates `<c:dLbls>` in `ppt/charts/chartX.xml` without corrupting plot area layout settings.
