# Skill: Dual-Sync Native Chart Data

This skill explains how to update PowerPoint native charts (`<c:chartSpace>`) while keeping the embedded Excel cache synchronized to prevent PowerPoint Repair warnings.

---

## 1. Anatomy of a PowerPoint Chart

A chart in OpenXML comprises:
1. **Host Slide (`ppt/slides/slideX.xml`)**:
   Contains a `<p:graphicFrame>` whose `<a:graphicData>` has URI `http://schemas.openxmlformats.org/drawingml/2006/chart` and links via `r:id` to a chart part.
2. **Slide Relationship (`ppt/slides/_rels/slideX.xml.rels`)**:
   Maps the `rId` to `../charts/chartY.xml`.
3. **Chart Definition (`ppt/charts/chartY.xml`)**:
   DrawingML `<c:chartSpace>` describing the visual plot (series, colors, labels, categories, values).
4. **Chart Relationship (`ppt/charts/_rels/chartY.xml.rels`)**:
   Maps an `rId` to the embedded spreadsheet: `../embeddings/Microsoft_Excel_Worksheet1.xlsx`.
5. **Embedded Excel Workbook (`ppt/embeddings/Microsoft_Excel_Worksheet*.xlsx`)**:
   Binary OpenXML spreadsheet storing the backing data table.

---

## 2. The Dual-Sync Requirement

When PowerPoint opens a presentation:
- It renders the graphic using the cached numbers inside `ppt/charts/chartY.xml`.
- If a user clicks **"Edit Data in Excel"**, PowerPoint opens the embedded `.xlsx` workbook.
- If the values or dimensions in `chartY.xml` conflict with the `.xlsx` sheet, or if the formula references (e.g. `Sheet1!$B$2:$B$5`) point to non-existent rows/cols, PowerPoint flags the presentation as **corrupt** and prompts the user to "Repair" it.

---

## 3. Step-by-Step Chart Synchronization

### Step 1: Locate the Target Chart
Find the chart relationship ID from the slide:
```javascript
const chartRel = slideRels.find(rel => rel['@_Type'].includes('chart'));
const chartPath = `ppt/${chartRel['@_Target'].replace('../', '')}`;
const chartXml = vfs.readText(chartPath);
const chartDoc = xmlParser.parse(chartXml);
```

### Step 2: Update DrawingML Series Data (`chartY.xml`)
Navigate to the plot area:
```javascript
const plotArea = chartDoc['c:chartSpace']['c:chart']['c:plotArea'];
// Determine chart type: barChart, lineChart, pieChart, etc.
const chartTypeKey = Object.keys(plotArea).find(k => k.startsWith('c:') && k.endsWith('Chart'));
const chartObj = plotArea[chartTypeKey];
const seriesList = ensureArray(chartObj['c:ser']);
```

For each series in `seriesList`:
1. **Update Categories (`<c:cat>`)**:
   ```javascript
   const strRef = series['c:cat']['c:strRef'];
   strRef['c:f'] = `Sheet1!$A$2:$A$${categories.length + 1}`;
   strRef['c:strCache']['c:ptCount']['@_val'] = categories.length;
   strRef['c:strCache']['c:pt'] = categories.map((cat, idx) => ({
     '@_idx': idx,
     'c:v': String(cat)
   }));
   ```

2. **Update Values (`<c:val>`)**:
   ```javascript
   const numRef = series['c:val']['c:numRef'];
   numRef['c:f'] = `Sheet1!$B$2:$B$${values.length + 1}`;
   numRef['c:numCache']['c:ptCount']['@_val'] = values.length;
   numRef['c:numCache']['c:pt'] = values.map((val, idx) => ({
     '@_idx': idx,
     'c:v': String(val)
   }));
   ```

3. **Update Series Title (`<c:tx>`)**:
   ```javascript
   if (series['c:tx'] && series['c:tx']['c:strRef']) {
     series['c:tx']['c:strRef']['c:strCache']['c:pt']['c:v'] = seriesName;
   }
   ```

### Step 3: Update Embedded Excel Workbook (`.xlsx`)
Locate the embedded workbook from `ppt/charts/_rels/chartY.xml.rels`:
```javascript
const chartRelPath = `ppt/charts/_rels/${path.basename(chartPath)}.rels`;
const chartRels = xmlParser.parse(vfs.readText(chartRelPath));
const sheetRel = ensureArray(chartRels.Relationships.Relationship)
  .find(r => r['@_Type'].includes('package') || r['@_Target'].endsWith('.xlsx'));

if (sheetRel) {
  const xlsxPath = `ppt/${sheetRel['@_Target'].replace('../', '')}`;
  const xlsxBuffer = vfs.readBinary(xlsxPath);
  
  // Use ChartWorkbookUpdater to patch sheet1.xml inside the embedded xlsx zip:
  const updatedXlsxBuffer = await updateEmbeddedExcelData(xlsxBuffer, {
    categories,
    seriesData: [{ name: seriesName, values }]
  });
  
  vfs.writeBinary(xlsxPath, updatedXlsxBuffer);
}
```

---

## 4. Common Chart Sync Traps & How to Avoid Them

1. **Mismatched `ptCount` vs. `pt` length**:
   Always set `<c:ptCount @_val="N"/>` exactly equal to the length of the `c:pt` array.
2. **Skipped or Non-Contiguous `idx`**:
   The `idx` attribute in `<c:pt @_idx="0">` must start at `0` and increment by `1` without gaps.
3. **Invalid Formula Ranges**:
   If you expand the data from 4 items to 8 items, update `<c:f>` from `$A$2:$A$5` to `$A$2:$A$9`. Leaving a stale formula range causes Excel to prompt "Formula cell range does not match series data".
4. **Data Type Mismatches**:
   Categories are typically strings (`<c:strRef>` / `<c:strCache>`), whereas values are numeric (`<c:numRef>` / `<c:numCache>`). Do not put numbers in `<c:strCache>` unless the chart is using text categories.
