const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, ImageRun,
  Header, Footer, AlignmentType, LevelFormat,
  HeadingLevel, BorderStyle, WidthType, ShadingType,
  PageNumber, PageBreak
} = require("docx");

const border = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const borders = { top: border, bottom: border, left: border, right: border };
const cellMargins = { top: 80, bottom: 80, left: 120, right: 120 };

const DARK_BLUE = "1B3A5C";
const MED_BLUE = "2E75B6";
const LIGHT_GRAY = "F2F2F2";
const RED_BG = "FCE4EC";
const GREEN_BG = "E8F5E9";
const AMBER_BG = "FFF8E1";

function hCell(text, w) {
  return new TableCell({ borders, width: { size: w, type: WidthType.DXA },
    shading: { fill: DARK_BLUE, type: ShadingType.CLEAR }, margins: cellMargins,
    children: [new Paragraph({ children: [new TextRun({ text, bold: true, color: "FFFFFF", font: "Arial", size: 20 })] })] });
}
function dCell(text, w, o = {}) {
  return new TableCell({ borders, width: { size: w, type: WidthType.DXA },
    shading: o.shading ? { fill: o.shading, type: ShadingType.CLEAR } : undefined, margins: cellMargins,
    children: [new Paragraph({ alignment: o.align, children: [new TextRun({ text, font: "Arial", size: 20, bold: o.bold||false, color: o.color||"333333", italics: o.italics||false })] })] });
}
function h(text, level) { return new Paragraph({ heading: level, children: [new TextRun({ text, font: "Arial" })] }); }
function p(text, o = {}) { return new Paragraph({ spacing: { after: 120 }, children: [new TextRun({ text, font: "Arial", size: 22, ...o })] }); }
function mp(runs) { return new Paragraph({ spacing: { after: 120 }, children: runs.map(r => new TextRun({ font: "Arial", size: 22, ...r })) }); }
function bi(text, ref) { return new Paragraph({ numbering: { reference: ref, level: 0 }, spacing: { after: 80 }, children: [new TextRun({ text, font: "Arial", size: 22 })] }); }
function ni(text, ref) { return new Paragraph({ numbering: { reference: ref, level: 0 }, spacing: { after: 80 }, children: [new TextRun({ text, font: "Arial", size: 22 })] }); }

// Load images
const baseAP = fs.readFileSync("outputs/plots/actual_vs_predicted.png");
const baseRes = fs.readFileSync("outputs/plots/residuals.png");
const baseFI = fs.readFileSync("outputs/plots/feature_importance.png");
const strictAP = fs.readFileSync("outputs/strict/plots/actual_vs_predicted.png");
const strictRes = fs.readFileSync("outputs/strict/plots/residuals.png");
const strictFI = fs.readFileSync("outputs/strict/plots/feature_importance.png");

function img(data, w, ht, title) {
  return new Paragraph({ children: [new ImageRun({
    type: "png", data, transformation: { width: w, height: ht },
    altText: { title, description: title, name: title },
  })] });
}

// Helper for alternating row shading
function altRow(cells, i) { return new TableRow({ children: cells }); }

const doc = new Document({
  styles: {
    default: { document: { run: { font: "Arial", size: 22 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 36, bold: true, font: "Arial", color: DARK_BLUE },
        paragraph: { spacing: { before: 360, after: 200 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 28, bold: true, font: "Arial", color: MED_BLUE },
        paragraph: { spacing: { before: 240, after: 160 }, outlineLevel: 1 } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 24, bold: true, font: "Arial", color: "444444" },
        paragraph: { spacing: { before: 200, after: 120 }, outlineLevel: 2 } },
    ],
  },
  numbering: { config: [
    { reference: "b1", levels: [{ level: 0, format: LevelFormat.BULLET, text: "-", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
    { reference: "b2", levels: [{ level: 0, format: LevelFormat.BULLET, text: "-", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
    { reference: "b3", levels: [{ level: 0, format: LevelFormat.BULLET, text: "-", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
    { reference: "b4", levels: [{ level: 0, format: LevelFormat.BULLET, text: "-", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
    { reference: "b5", levels: [{ level: 0, format: LevelFormat.BULLET, text: "-", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
    { reference: "n1", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
  ] },
  sections: [{
    properties: {
      page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } },
    },
    headers: { default: new Header({ children: [new Paragraph({
      border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: MED_BLUE, space: 1 } },
      children: [
        new TextRun({ text: "Ascent Global Logistics", font: "Arial", size: 18, color: MED_BLUE, bold: true }),
        new TextRun({ text: "  |  Baseline vs Strict Comparison", font: "Arial", size: 18, color: "888888" }),
      ],
    })] }) },
    footers: { default: new Footer({ children: [new Paragraph({
      border: { top: { style: BorderStyle.SINGLE, size: 4, color: "CCCCCC", space: 1 } },
      alignment: AlignmentType.CENTER,
      children: [
        new TextRun({ text: "Confidential  |  Page ", font: "Arial", size: 16, color: "888888" }),
        new TextRun({ children: [PageNumber.CURRENT], font: "Arial", size: 16, color: "888888" }),
      ],
    })] }) },
    children: [
      // ===== TITLE =====
      new Paragraph({ spacing: { before: 3000 } }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 200 },
        children: [new TextRun({ text: "Freight Pricing Dataset", font: "Arial", size: 56, bold: true, color: DARK_BLUE })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 100 },
        children: [new TextRun({ text: "Baseline vs Strict Cleaning: Model Comparison & Data Guide", font: "Arial", size: 30, color: MED_BLUE })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: MED_BLUE, space: 1 } }, spacing: { after: 600 }, children: [] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 },
        children: [new TextRun({ text: "Date: March 26, 2026", font: "Arial", size: 22, color: "555555" })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 },
        children: [new TextRun({ text: "Version: 2.0 - Comparison Report", font: "Arial", size: 20, color: "888888", italics: true })] }),

      new Paragraph({ children: [new PageBreak()] }),

      // ===== 1. EXECUTIVE SUMMARY =====
      h("1. Executive Summary", HeadingLevel.HEADING_1),
      p("This document compares two data cleaning approaches for the Ascent Freight Pricing prediction task and serves as a data guide for anyone reviewing this project."),
      mp([
        { text: "Baseline Pipeline: ", bold: true },
        { text: "Caps outliers at thresholds and imputes missing values with median. Retains 9,934 of 9,991 rows (99.4%). Random Forest R-squared = 0.5746." },
      ]),
      mp([
        { text: "Strict Pipeline: ", bold: true },
        { text: "Removes all garbage data, negative values, invalid dates, NaN rows, and IQR outliers. Retains 6,028 of 9,991 rows (60.3%). Linear Regression R-squared = 0.5671, nearly matching Random Forest." },
      ]),
      mp([
        { text: "Key Finding: ", bold: true, color: "C62828" },
        { text: "Removing outliers dramatically improved Linear Regression (from R-squared = -0.36 to 0.57) — proving the baseline model was severely impacted by dirty data. Both models now perform comparably, with cleaner residuals and much lower error magnitudes." },
      ]),

      new Paragraph({ children: [new PageBreak()] }),

      // ===== 2. DATA GUIDE =====
      h("2. Data Guide: Understanding the Raw Dataset", HeadingLevel.HEADING_1),
      p("This section provides a comprehensive overview of the raw data for anyone new to this project."),

      h("2.1 Dataset At a Glance", HeadingLevel.HEADING_2),
      new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [3120, 6240], rows: [
        new TableRow({ children: [hCell("Property", 3120), hCell("Value", 6240)] }),
        new TableRow({ children: [dCell("Source File", 3120, { bold: true }), dCell("ascent freight pricing dataset.csv", 6240)] }),
        new TableRow({ children: [dCell("Total Rows", 3120, { bold: true, shading: LIGHT_GRAY }), dCell("9,991 (includes 1 duplicate header row)", 6240, { shading: LIGHT_GRAY })] }),
        new TableRow({ children: [dCell("Total Columns", 3120, { bold: true }), dCell("11 (1 ID + 1 date + 4 categorical + 5 numeric + 1 target)", 6240)] }),
        new TableRow({ children: [dCell("Target Variable", 3120, { bold: true, shading: LIGHT_GRAY }), dCell("winning_bid_price (continuous, USD)", 6240, { shading: LIGHT_GRAY })] }),
        new TableRow({ children: [dCell("Date Range", 3120, { bold: true }), dCell("Jan 2024 - Dec 2026 (with some invalid dates up to 2099)", 6240)] }),
        new TableRow({ children: [dCell("Task Type", 3120, { bold: true, shading: LIGHT_GRAY }), dCell("Regression: predict the winning carrier bid price for a shipment", 6240, { shading: LIGHT_GRAY })] }),
      ] }),
      new Paragraph({ spacing: { after: 200 } }),

      h("2.2 Column Dictionary", HeadingLevel.HEADING_2),
      p("Complete reference for every column in the dataset:"),
      new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [2000, 1100, 1100, 5160], rows: [
        new TableRow({ children: [hCell("Column", 2000), hCell("Type", 1100), hCell("Nulls", 1100), hCell("Description & Notes", 5160)] }),
        ...([
          ["shipment_id", "ID", "0", "Unique shipment identifier (format: SHP-XXXXXX). No duplicates found."],
          ["shipment_date", "Date", "8", "Shipment date. WARNING: 3 mixed formats (YYYY-MM-DD, MM/DD/YYYY, YYYY-MM-DD HH:MM:SS). Some dates go up to year 2099 (invalid)."],
          ["origin_region", "Categorical", "8", "Origin region code (26 unique values, e.g., MW-1, WC-2, SE-3). Format: Region-Number."],
          ["destination_region", "Categorical", "8", "Destination region code (24 unique values). Same format as origin_region."],
          ["equipment_type", "Categorical", "8", "Type of transport equipment. MAJOR ISSUE: 32 messy variants for 7 real types. See Equipment Type section below."],
          ["weight_lbs", "Numeric", "328", "Shipment weight in pounds. Contains garbage values (TBD, #REF!, pending, -). Range: -500 to 1,000,000 after coercion."],
          ["cubic_feet", "Numeric", "517", "Shipment volume in cubic feet. Contains garbage values (pending, #REF!). Some negative values."],
          ["lead_time_hours", "Numeric", "207", "Hours between booking and shipment. Contains negatives and extreme values (max: 9,999 hrs = 416 days)."],
          ["num_carrier_bids", "Numeric", "395", "Number of carrier bids received. Expect 1-20 typical range. Max found: 999 (likely data error)."],
          ["lane_volume", "Numeric", "8", "Historical shipment volume on this lane. Higher = more frequently used route."],
          ["winning_bid_price", "Numeric", "8", "TARGET: Winning carrier bid price in USD. Contains 33 zeros and 15 negatives (invalid). Mean: $1,750, Median: $1,057."],
        ]).map((row, i) => new TableRow({ children: row.map((val, j) =>
          dCell(val, [2000, 1100, 1100, 5160][j], { shading: i % 2 === 0 ? LIGHT_GRAY : undefined, bold: j === 0 })
        ) })),
      ] }),
      new Paragraph({ spacing: { after: 200 } }),

      h("2.3 Equipment Type Reference", HeadingLevel.HEADING_2),
      p("The 7 canonical equipment types and their raw data variants:"),
      new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [1800, 1200, 6360], rows: [
        new TableRow({ children: [hCell("Canonical Name", 1800), hCell("Count (Raw)", 1200), hCell("Variants in Data", 6360)] }),
        ...[
          ["Small Straight", "2,480", "Small Straight, small straight, Sm Straight, Small Str"],
          ["Sprinter Van", "2,002", "Sprinter Van, sprinter van, SPRINTER VAN, Sprinter"],
          ["Large Straight", "1,769", "Large Straight, LARGE STRAIGHT, Large Str, Lg Straight"],
          ["Cargo Van", "1,530", "Cargo Van, cargo van, CARGO VAN, CV, Cargo  Van, <b>Cargo Van</b>"],
          ["Tractor Trailer", "1,021", "Tractor Trailer, tractor trailer, Tractor, TT"],
          ["Flatbed", "672", "Flatbed, flatbed, FLATBED, Flat Bed"],
          ["Reefer", "483", "Reefer, reefer, REEFER, Refrigerated"],
        ].map((row, i) => new TableRow({ children: [
          dCell(row[0], 1800, { bold: true, shading: i % 2 === 0 ? LIGHT_GRAY : undefined }),
          dCell(row[1], 1200, { shading: i % 2 === 0 ? LIGHT_GRAY : undefined }),
          dCell(row[2], 6360, { shading: i % 2 === 0 ? LIGHT_GRAY : undefined }),
        ] })),
      ] }),
      new Paragraph({ spacing: { after: 100 } }),
      p("Additionally: 25 rows with 'Test Equipment' and 1 row with literal 'equipment_type' header value are dropped.", { italics: true }),

      h("2.4 Known Data Quality Issues", HeadingLevel.HEADING_2),
      ni("Garbage values in numeric columns: TBD, pending, #REF!, - (total: 18 values across weight and cubic_feet)", "n1"),
      ni("Negative values: weight_lbs (-500), cubic_feet (-10), lead_time_hours (4 rows), num_carrier_bids (-1)", "n1"),
      ni("Extreme outliers: weight up to 1,000,000 lbs, lead_time up to 9,999 hrs, num_bids up to 999", "n1"),
      ni("Invalid dates: some dates in year 2099 (8 rows with dates outside 2023-2026 range)", "n1"),
      ni("Zero/negative target prices: 33 zeros + 15 negatives = 48 invalid target rows", "n1"),
      ni("High missingness: cubic_feet (517 nulls, 5.2%), num_carrier_bids (395, 4.0%), weight_lbs (328, 3.3%)", "n1"),
      ni("1 duplicate header row embedded in the data", "n1"),

      new Paragraph({ children: [new PageBreak()] }),

      // ===== 3. CLEANING APPROACHES COMPARED =====
      h("3. Cleaning Approaches Compared", HeadingLevel.HEADING_1),

      h("3.1 Baseline Pipeline (Cap + Impute)", HeadingLevel.HEADING_2),
      p("Conservative approach: caps outliers at thresholds, fills missing values with median. Maximizes data retention."),
      new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [4000, 2680, 2680], rows: [
        new TableRow({ children: [hCell("Step", 4000), hCell("Rows Removed", 2680), hCell("Remaining", 2680)] }),
        ...[
          ["Drop header row", "1", "9,990"],
          ["Drop null rows", "8", "9,982"],
          ["Standardize equipment_type", "25", "9,957"],
          ["Remove invalid targets", "23", "9,934"],
          ["Cap outliers", "0 (capped)", "9,934"],
          ["Impute missing (median)", "0 (filled)", "9,934"],
        ].map((r, i) => new TableRow({ children: [
          dCell(r[0], 4000, { bold: true, shading: i%2===0?LIGHT_GRAY:undefined }),
          dCell(r[1], 2680, { shading: i%2===0?LIGHT_GRAY:undefined }),
          dCell(r[2], 2680, { bold: true, shading: i%2===0?LIGHT_GRAY:undefined }),
        ] })),
        new TableRow({ children: [
          dCell("TOTAL", 4000, { bold: true, shading: GREEN_BG }),
          dCell("57 removed", 2680, { shading: GREEN_BG }),
          dCell("9,934 (99.4%)", 2680, { bold: true, color: "2E7D32", shading: GREEN_BG }),
        ] }),
      ] }),
      new Paragraph({ spacing: { after: 300 } }),

      h("3.2 Strict Pipeline (Remove + No Imputation)", HeadingLevel.HEADING_2),
      p("Aggressive approach: removes all garbage, negative values, invalid dates, NaN rows, and IQR outliers. No imputation — only verified clean data."),
      new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [4000, 2680, 2680], rows: [
        new TableRow({ children: [hCell("Step", 4000), hCell("Rows Removed", 2680), hCell("Remaining", 2680)] }),
        ...[
          ["Drop header row", "1", "9,990"],
          ["Drop null rows", "8", "9,982"],
          ["Standardize equipment_type", "25", "9,957"],
          ["Remove invalid targets", "23", "9,934"],
          ["Remove negative values", "11", "9,923"],
          ["Remove invalid dates (outside 2023-2026)", "8", "9,915"],
          ["Drop rows with any NaN", "1,346", "8,569"],
          ["Remove IQR outliers (1.5x)", "2,541", "6,028"],
        ].map((r, i) => new TableRow({ children: [
          dCell(r[0], 4000, { bold: true, shading: i%2===0?LIGHT_GRAY:undefined }),
          dCell(r[1], 2680, { shading: i%2===0?LIGHT_GRAY:undefined }),
          dCell(r[2], 2680, { bold: true, shading: i%2===0?LIGHT_GRAY:undefined }),
        ] })),
        new TableRow({ children: [
          dCell("TOTAL", 4000, { bold: true, shading: AMBER_BG }),
          dCell("3,963 removed", 2680, { shading: AMBER_BG }),
          dCell("6,028 (60.3%)", 2680, { bold: true, color: "E65100", shading: AMBER_BG }),
        ] }),
      ] }),
      new Paragraph({ spacing: { after: 200 } }),

      h("3.3 IQR Outlier Removal Details (Strict)", HeadingLevel.HEADING_2),
      new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [1800, 1260, 1260, 1800, 1800, 1440], rows: [
        new TableRow({ children: [hCell("Column", 1800), hCell("Q1", 1260), hCell("Q3", 1260), hCell("Lower Bound", 1800), hCell("Upper Bound", 1800), hCell("Removed", 1440)] }),
        ...[
          ["weight_lbs", "513.7", "2,234.6", "-2,067.7", "4,816.0", "626"],
          ["cubic_feet", "3.7", "15.1", "-13.4", "32.2", "382"],
          ["lead_time_hours", "6.5", "23.5", "-19.0", "49.0", "864"],
          ["num_carrier_bids", "2.0", "6.0", "-4.0", "12.0", "64"],
          ["lane_volume", "74.0", "198.0", "-112.0", "384.0", "180"],
          ["winning_bid_price", "563.3", "1,731.6", "-1,189.3", "3,484.2", "425"],
        ].map((r, i) => new TableRow({ children: r.map((v, j) =>
          dCell(v, [1800,1260,1260,1800,1800,1440][j], { shading: i%2===0?LIGHT_GRAY:undefined, bold: j===0 || j===5 })
        ) })),
      ] }),

      new Paragraph({ children: [new PageBreak()] }),

      // ===== 4. STRICT DATA PROFILE =====
      h("4. Strict Cleaned Data Profile", HeadingLevel.HEADING_1),
      p("After strict cleaning, the dataset is free of nulls, negatives, and statistical outliers."),

      h("4.1 Numeric Feature Distributions (Post-Cleaning)", HeadingLevel.HEADING_2),
      new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [1800, 1080, 1080, 1080, 1080, 1080, 1080, 1080], rows: [
        new TableRow({ children: [hCell("Column", 1800), hCell("Min", 1080), hCell("Q1", 1080), hCell("Median", 1080), hCell("Q3", 1080), hCell("Max", 1080), hCell("Mean", 1080), hCell("Std", 1080)] }),
        ...[
          ["weight_lbs", "0.1", "435.3", "880.2", "1,549.6", "4,813.2", "1,119.4", "882.6"],
          ["cubic_feet", "0.0", "3.4", "6.9", "12.4", "32.1", "8.8", "6.9"],
          ["lead_time_hrs", "0.0", "6.1", "12.0", "20.1", "49.0", "14.7", "11.0"],
          ["num_bids", "0.0", "2.0", "4.0", "6.0", "12.0", "4.4", "2.6"],
          ["lane_volume", "0.0", "73.0", "121.0", "196.0", "372.0", "140.0", "83.9"],
          ["bid_price", "0.0", "538.2", "913.4", "1,513.9", "3,479.7", "1,120.4", "755.8"],
        ].map((r, i) => new TableRow({ children: r.map((v, j) =>
          dCell(v, j===0?1800:1080, { shading: i%2===0?LIGHT_GRAY:undefined, bold: j===0 })
        ) })),
      ] }),
      new Paragraph({ spacing: { after: 200 } }),

      h("4.2 Equipment Type Distribution (Strict)", HeadingLevel.HEADING_2),
      new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [3120, 3120, 3120], rows: [
        new TableRow({ children: [hCell("Equipment Type", 3120), hCell("Count", 3120), hCell("Percentage", 3120)] }),
        ...[
          ["Small Straight", "1,678", "27.8%"],
          ["Sprinter Van", "1,449", "24.0%"],
          ["Large Straight", "1,170", "19.4%"],
          ["Cargo Van", "1,118", "18.5%"],
          ["Flatbed", "257", "4.3%"],
          ["Reefer", "214", "3.5%"],
          ["Tractor Trailer", "142", "2.4%"],
        ].map((r, i) => new TableRow({ children: [
          dCell(r[0], 3120, { bold: true, shading: i%2===0?LIGHT_GRAY:undefined }),
          dCell(r[1], 3120, { shading: i%2===0?LIGHT_GRAY:undefined }),
          dCell(r[2], 3120, { shading: i%2===0?LIGHT_GRAY:undefined }),
        ] })),
      ] }),
      new Paragraph({ spacing: { after: 100 } }),
      p("Note: Tractor Trailer and Flatbed had more outliers removed proportionally (heavy/high-value shipments), reducing their share.", { italics: true }),

      new Paragraph({ children: [new PageBreak()] }),

      // ===== 5. MODEL COMPARISON =====
      h("5. Model Results Comparison", HeadingLevel.HEADING_1),

      h("5.1 Performance Summary", HeadingLevel.HEADING_2),
      new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [2340, 1170, 1170, 1170, 1170, 1170, 1170], rows: [
        new TableRow({ children: [
          hCell("", 2340), hCell("MAE", 1170), hCell("RMSE", 1170), hCell("R-sq", 1170),
          hCell("MAE", 1170), hCell("RMSE", 1170), hCell("R-sq", 1170),
        ] }),
        new TableRow({ children: [
          dCell("", 2340, { bold: true }),
          dCell("BASELINE", 1170, { bold: true, shading: LIGHT_GRAY }), dCell("", 1170, { shading: LIGHT_GRAY }), dCell("", 1170, { shading: LIGHT_GRAY }),
          dCell("STRICT", 1170, { bold: true, shading: GREEN_BG }), dCell("", 1170, { shading: GREEN_BG }), dCell("", 1170, { shading: GREEN_BG }),
        ] }),
        new TableRow({ children: [
          dCell("Linear Regression", 2340, { bold: true }),
          dCell("$831", 1170), dCell("$2,457", 1170), dCell("-0.36", 1170, { color: "C62828", bold: true }),
          dCell("$355", 1170, { shading: GREEN_BG }), dCell("$495", 1170, { shading: GREEN_BG }), dCell("0.567", 1170, { color: "2E7D32", bold: true, shading: GREEN_BG }),
        ] }),
        new TableRow({ children: [
          dCell("Random Forest", 2340, { bold: true, shading: LIGHT_GRAY }),
          dCell("$662", 1170, { shading: LIGHT_GRAY }), dCell("$1,374", 1170, { shading: LIGHT_GRAY }), dCell("0.575", 1170, { bold: true, shading: LIGHT_GRAY }),
          dCell("$360", 1170, { shading: GREEN_BG }), dCell("$502", 1170, { shading: GREEN_BG }), dCell("0.554", 1170, { bold: true, shading: GREEN_BG }),
        ] }),
      ] }),
      new Paragraph({ spacing: { after: 200 } }),

      h("5.2 Key Observations", HeadingLevel.HEADING_2),
      bi("Linear Regression improved dramatically: R-squared went from -0.36 (worse than mean) to 0.567 (meaningful). This proves outliers were severely distorting the linear model.", "b1"),
      bi("Both models now perform comparably (~0.55-0.57 R-squared). In clean data, the relationship is more linear than expected.", "b1"),
      bi("MAE dropped from $662-$831 to $355-$360 — prediction errors are now roughly half the size.", "b1"),
      bi("RMSE dropped even more (from $1,374-$2,457 to $495-$502) — eliminating large-error outliers was the main driver.", "b1"),
      bi("The gap between MAE and RMSE is now much smaller, indicating fewer extreme prediction errors.", "b1"),
      new Paragraph({ spacing: { after: 200 } }),

      h("5.3 Trade-Off Analysis", HeadingLevel.HEADING_2),
      new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [1560, 3900, 3900], rows: [
        new TableRow({ children: [hCell("Aspect", 1560), hCell("Baseline (Cap + Impute)", 3900), hCell("Strict (Remove Outliers)", 3900)] }),
        ...[
          ["Rows", "9,934 (99.4%)", "6,028 (60.3%)"],
          ["Data Loss", "Minimal — only truly invalid rows", "Significant — 40% of data removed"],
          ["Data Quality", "Contains capped values + imputed medians", "Only verified, clean data points"],
          ["LR R-sq", "-0.3609 (broken)", "0.5671 (usable)"],
          ["RF R-sq", "0.5746", "0.5541"],
          ["Best For", "When data volume matters most", "When model reliability matters most"],
        ].map((r, i) => new TableRow({ children: [
          dCell(r[0], 1560, { bold: true, shading: i%2===0?LIGHT_GRAY:undefined }),
          dCell(r[1], 3900, { shading: i%2===0?LIGHT_GRAY:undefined }),
          dCell(r[2], 3900, { shading: i%2===0?LIGHT_GRAY:undefined }),
        ] })),
      ] }),

      new Paragraph({ children: [new PageBreak()] }),

      // ===== 6. VISUAL COMPARISON =====
      h("6. Visual Comparison", HeadingLevel.HEADING_1),

      h("6.1 Actual vs Predicted: Baseline", HeadingLevel.HEADING_2),
      img(baseAP, 600, 257, "Baseline Actual vs Predicted"),
      new Paragraph({ spacing: { after: 100 } }),
      h("6.2 Actual vs Predicted: Strict", HeadingLevel.HEADING_2),
      img(strictAP, 600, 257, "Strict Actual vs Predicted"),

      new Paragraph({ children: [new PageBreak()] }),

      h("6.3 Residuals: Baseline", HeadingLevel.HEADING_2),
      img(baseRes, 600, 257, "Baseline Residuals"),
      new Paragraph({ spacing: { after: 100 } }),
      h("6.4 Residuals: Strict", HeadingLevel.HEADING_2),
      img(strictRes, 600, 257, "Strict Residuals"),

      new Paragraph({ children: [new PageBreak()] }),

      h("6.5 Feature Importance: Baseline (Random Forest)", HeadingLevel.HEADING_2),
      img(baseFI, 460, 368, "Baseline Feature Importance"),
      new Paragraph({ spacing: { after: 200 } }),
      h("6.6 Feature Importance: Strict (Random Forest)", HeadingLevel.HEADING_2),
      img(strictFI, 460, 368, "Strict Feature Importance"),

      new Paragraph({ children: [new PageBreak()] }),

      // ===== 7. DECISION LOG =====
      h("7. Decision Log", HeadingLevel.HEADING_1),
      new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [1200, 2700, 3260, 2200], rows: [
        new TableRow({ children: [hCell("Date", 1200), hCell("Decision", 2700), hCell("Rationale", 3260), hCell("Status", 2200)] }),
        ...[
          ["2026-03-26", "Build reusable pipeline", "Avoid rework for new data. CLI-driven with config.py.", "Implemented", GREEN_BG, "2E7D32"],
          ["2026-03-26", "Defer feature engineering", "Establish baseline first, then add complexity.", "Deferred", AMBER_BG, "E65100"],
          ["2026-03-26", "Create two cleaning modes", "Baseline (cap+impute) vs Strict (remove). Comparison reveals data quality impact.", "Implemented", GREEN_BG, "2E7D32"],
          ["2026-03-26", "Remove outliers via IQR (1.5x)", "Standard statistical method. Removes extreme values that distort models.", "Implemented", GREEN_BG, "2E7D32"],
          ["2026-03-26", "No imputation in strict mode", "Only use verified data points. Avoids bias from synthetic median values.", "Implemented", GREEN_BG, "2E7D32"],
          ["2026-03-26", "Remove invalid dates (>2026)", "Dates like 2099-01-01 are clearly data entry errors.", "Implemented", GREEN_BG, "2E7D32"],
          ["2026-03-26", "Document all decisions", "Living document for team visibility. Updated each phase.", "Ongoing", undefined, MED_BLUE],
        ].map((r, i) => new TableRow({ children: [
          dCell(r[0], 1200, { shading: i%2===0?LIGHT_GRAY:undefined }),
          dCell(r[1], 2700, { bold: true, shading: i%2===0?LIGHT_GRAY:undefined }),
          dCell(r[2], 3260, { shading: i%2===0?LIGHT_GRAY:undefined }),
          dCell(r[3], 2200, { bold: true, color: r[5], shading: r[4] }),
        ] })),
      ] }),

      new Paragraph({ children: [new PageBreak()] }),

      // ===== 8. RECOMMENDATIONS =====
      h("8. Recommendations & Next Steps", HeadingLevel.HEADING_1),

      h("8.1 Recommended Approach", HeadingLevel.HEADING_2),
      mp([
        { text: "Use the strict pipeline as the primary model training dataset. ", bold: true },
        { text: "The dramatic improvement in Linear Regression proves that data quality is the primary driver of model performance, not model complexity. Clean data with a simple model outperforms dirty data with a complex model." },
      ]),

      h("8.2 Immediate Next Steps", HeadingLevel.HEADING_2),
      bi("Investigate the 1,346 rows with NaN values — can they be recovered from the source system?", "b3"),
      bi("Review the IQR bounds with domain experts — some removed shipments may be valid edge cases (e.g., heavy Tractor Trailer loads)", "b3"),
      bi("Add feature engineering: date features, weight-per-cubic-foot density, lane-level price history", "b3"),
      bi("Try Gradient Boosting (XGBoost/LightGBM) on the strict cleaned data", "b3"),

      h("8.3 Data Source Improvements", HeadingLevel.HEADING_2),
      bi("Fix garbage values at source: TBD, #REF!, pending should not enter the dataset", "b4"),
      bi("Validate equipment_type at data entry to prevent free-text variants", "b4"),
      bi("Add date validation to prevent future dates (2099)", "b4"),
      bi("Add range validation for weight, lead_time, and bid counts", "b4"),

      h("8.4 Pipeline Usage", HeadingLevel.HEADING_2),
      p("python run_pipeline.py --mode baseline    # Conservative (cap + impute)", { font: "Consolas", size: 20 }),
      p("python run_pipeline.py --mode strict      # Aggressive (remove outliers)", { font: "Consolas", size: 20 }),
      p("python run_pipeline.py --mode strict --input new_data.csv", { font: "Consolas", size: 20 }),
    ],
  }],
});

Packer.toBuffer(doc).then((buffer) => {
  fs.writeFileSync("Ascent_Freight_Pricing_Comparison_Report.docx", buffer);
  console.log("Comparison report created successfully!");
});
