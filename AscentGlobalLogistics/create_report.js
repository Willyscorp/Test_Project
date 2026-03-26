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

// Colors
const DARK_BLUE = "1B3A5C";
const MED_BLUE = "2E75B6";
const LIGHT_GRAY = "F2F2F2";
const RED_BG = "FCE4EC";
const GREEN_BG = "E8F5E9";
const AMBER_BG = "FFF8E1";

function headerCell(text, width) {
  return new TableCell({
    borders,
    width: { size: width, type: WidthType.DXA },
    shading: { fill: DARK_BLUE, type: ShadingType.CLEAR },
    margins: cellMargins,
    verticalAlign: "center",
    children: [new Paragraph({ children: [new TextRun({ text, bold: true, color: "FFFFFF", font: "Arial", size: 20 })] })],
  });
}

function dataCell(text, width, opts = {}) {
  return new TableCell({
    borders,
    width: { size: width, type: WidthType.DXA },
    shading: opts.shading ? { fill: opts.shading, type: ShadingType.CLEAR } : undefined,
    margins: cellMargins,
    children: [new Paragraph({
      children: [new TextRun({ text, font: "Arial", size: 20, bold: opts.bold || false, color: opts.color || "333333", italics: opts.italics || false })],
    })],
  });
}

function heading(text, level) {
  return new Paragraph({ heading: level, children: [new TextRun({ text, font: "Arial" })] });
}

function para(text, opts = {}) {
  return new Paragraph({
    spacing: { after: 120 },
    children: [new TextRun({ text, font: "Arial", size: 22, ...opts })],
  });
}

function multiRunPara(runs) {
  return new Paragraph({
    spacing: { after: 120 },
    children: runs.map(r => new TextRun({ font: "Arial", size: 22, ...r })),
  });
}

function bulletItem(text, ref) {
  return new Paragraph({
    numbering: { reference: ref, level: 0 },
    spacing: { after: 80 },
    children: [new TextRun({ text, font: "Arial", size: 22 })],
  });
}

function numberItem(text, ref) {
  return new Paragraph({
    numbering: { reference: ref, level: 0 },
    spacing: { after: 80 },
    children: [new TextRun({ text, font: "Arial", size: 22 })],
  });
}

// Load images
const actualVsPredImg = fs.readFileSync("E:/MyProject/AscentGlobalLogistics/outputs/plots/actual_vs_predicted.png");
const featureImpImg = fs.readFileSync("E:/MyProject/AscentGlobalLogistics/outputs/plots/feature_importance.png");
const residualsImg = fs.readFileSync("E:/MyProject/AscentGlobalLogistics/outputs/plots/residuals.png");

const doc = new Document({
  styles: {
    default: { document: { run: { font: "Arial", size: 22 } } },
    paragraphStyles: [
      {
        id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 36, bold: true, font: "Arial", color: DARK_BLUE },
        paragraph: { spacing: { before: 360, after: 200 }, outlineLevel: 0 },
      },
      {
        id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 28, bold: true, font: "Arial", color: MED_BLUE },
        paragraph: { spacing: { before: 240, after: 160 }, outlineLevel: 1 },
      },
      {
        id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 24, bold: true, font: "Arial", color: "444444" },
        paragraph: { spacing: { before: 200, after: 120 }, outlineLevel: 2 },
      },
    ],
  },
  numbering: {
    config: [
      { reference: "b1", levels: [{ level: 0, format: LevelFormat.BULLET, text: "-", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "b2", levels: [{ level: 0, format: LevelFormat.BULLET, text: "-", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "b3", levels: [{ level: 0, format: LevelFormat.BULLET, text: "-", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "b4", levels: [{ level: 0, format: LevelFormat.BULLET, text: "-", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "b5", levels: [{ level: 0, format: LevelFormat.BULLET, text: "-", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "b6", levels: [{ level: 0, format: LevelFormat.BULLET, text: "-", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "n1", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "n2", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
    ],
  },
  sections: [
    {
      properties: {
        page: {
          size: { width: 12240, height: 15840 },
          margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
        },
      },
      headers: {
        default: new Header({
          children: [
            new Paragraph({
              border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: MED_BLUE, space: 1 } },
              children: [
                new TextRun({ text: "Ascent Global Logistics", font: "Arial", size: 18, color: MED_BLUE, bold: true }),
                new TextRun({ text: "  |  Freight Pricing Analysis", font: "Arial", size: 18, color: "888888" }),
              ],
            }),
          ],
        }),
      },
      footers: {
        default: new Footer({
          children: [
            new Paragraph({
              border: { top: { style: BorderStyle.SINGLE, size: 4, color: "CCCCCC", space: 1 } },
              alignment: AlignmentType.CENTER,
              children: [
                new TextRun({ text: "Confidential  |  Page ", font: "Arial", size: 16, color: "888888" }),
                new TextRun({ children: [PageNumber.CURRENT], font: "Arial", size: 16, color: "888888" }),
              ],
            }),
          ],
        }),
      },
      children: [
        // ===== TITLE PAGE =====
        new Paragraph({ spacing: { before: 3000 } }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 200 },
          children: [new TextRun({ text: "Freight Pricing Dataset", font: "Arial", size: 56, bold: true, color: DARK_BLUE })],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 100 },
          children: [new TextRun({ text: "Data Quality Assessment, Pipeline & Baseline Results", font: "Arial", size: 32, color: MED_BLUE })],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: MED_BLUE, space: 1 } },
          spacing: { after: 600 },
          children: [],
        }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 },
          children: [new TextRun({ text: "Prepared for: Ascent Global Logistics Team", font: "Arial", size: 22, color: "555555" })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 },
          children: [new TextRun({ text: "Date: March 26, 2026", font: "Arial", size: 22, color: "555555" })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 },
          children: [new TextRun({ text: "Objective: Predict winning_bid_price", font: "Arial", size: 22, color: "555555" })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 },
          children: [new TextRun({ text: "Version: 1.1 (Updated with actual pipeline results)", font: "Arial", size: 20, color: "888888", italics: true })] }),

        new Paragraph({ children: [new PageBreak()] }),

        // ===== 1. EXECUTIVE SUMMARY =====
        heading("1. Executive Summary", HeadingLevel.HEADING_1),
        para("This document presents the complete data quality assessment, cleaning pipeline architecture, and baseline model results for the Ascent Freight Pricing dataset. The goal is to predict winning_bid_price using shipment characteristics."),
        multiRunPara([
          { text: "Key Outcome: " , bold: true },
          { text: "A reusable data pipeline was built that cleans raw data and trains baseline models. The pipeline reduced 9,991 raw rows to 9,934 clean rows (99.4% retention). The Random Forest baseline achieved " },
          { text: "R-squared = 0.5746", bold: true },
          { text: " with MAE of $662.19, while Linear Regression performed poorly (R-squared = -0.36), confirming non-linear relationships in the data." },
        ]),
        multiRunPara([
          { text: "Decision: ", bold: true },
          { text: "Feature engineering is intentionally deferred to a future iteration. The baseline uses raw numeric features and one-hot encoded categoricals only. This establishes a performance floor before adding complexity." },
        ]),

        new Paragraph({ children: [new PageBreak()] }),

        // ===== 2. DATASET OVERVIEW =====
        heading("2. Dataset Overview", HeadingLevel.HEADING_1),
        new Table({
          width: { size: 9360, type: WidthType.DXA },
          columnWidths: [3120, 3120, 3120],
          rows: [
            new TableRow({ children: [headerCell("Metric", 3120), headerCell("Value", 3120), headerCell("Notes", 3120)] }),
            new TableRow({ children: [
              dataCell("Total Rows", 3120, { bold: true }), dataCell("9,991", 3120), dataCell("Including 1 duplicate header row", 3120)] }),
            new TableRow({ children: [
              dataCell("Total Columns", 3120, { bold: true }), dataCell("11", 3120), dataCell("1 target + 10 features", 3120)] }),
            new TableRow({ children: [
              dataCell("Target Variable", 3120, { bold: true }), dataCell("winning_bid_price", 3120), dataCell("Continuous numeric (regression)", 3120)] }),
            new TableRow({ children: [
              dataCell("Date Range", 3120, { bold: true }), dataCell("2024 - 2025", 3120), dataCell("Mixed formats across rows", 3120)] }),
            new TableRow({ children: [
              dataCell("Price Mean / Median", 3120, { bold: true }), dataCell("$1,749.75 / $1,056.78", 3120), dataCell("Right-skewed distribution", 3120)] }),
          ],
        }),
        new Paragraph({ spacing: { after: 200 } }),

        heading("2.1 Column Descriptions", HeadingLevel.HEADING_2),
        new Table({
          width: { size: 9360, type: WidthType.DXA },
          columnWidths: [2200, 1400, 1400, 4360],
          rows: [
            new TableRow({ children: [
              headerCell("Column", 2200), headerCell("Expected Type", 1400), headerCell("Nulls", 1400), headerCell("Description", 4360)] }),
            ...([
              ["shipment_id", "String", "0", "Unique identifier for each shipment"],
              ["shipment_date", "Date", "8", "Date of shipment (3 mixed formats)"],
              ["origin_region", "Categorical", "8", "Origin region code (e.g., MW-1, WC-2)"],
              ["destination_region", "Categorical", "8", "Destination region code"],
              ["equipment_type", "Categorical", "8", "Type of transport equipment (32 messy variants)"],
              ["weight_lbs", "Numeric", "328", "Shipment weight in pounds (has garbage values)"],
              ["cubic_feet", "Numeric", "517", "Shipment volume in cubic feet (has garbage values)"],
              ["lead_time_hours", "Numeric", "207", "Hours of lead time before shipment"],
              ["num_carrier_bids", "Numeric", "395", "Number of carrier bids received"],
              ["lane_volume", "Numeric", "8", "Historical volume on the shipping lane"],
              ["winning_bid_price", "Numeric", "8", "TARGET: Winning carrier bid price ($)"],
            ]).map((row, i) => new TableRow({
              children: row.map((val, j) => dataCell(val, [2200, 1400, 1400, 4360][j], {
                shading: i % 2 === 0 ? LIGHT_GRAY : undefined, bold: j === 0,
              })),
            })),
          ],
        }),
        new Paragraph({ spacing: { after: 200 } }),

        // ===== 3. DATA QUALITY ISSUES =====
        heading("3. Data Quality Issues Identified", HeadingLevel.HEADING_1),
        para("The following critical issues were found during data profiling. All columns were read as object (string) type due to embedded garbage values."),
        new Table({
          width: { size: 9360, type: WidthType.DXA },
          columnWidths: [2800, 4200, 2360],
          rows: [
            new TableRow({ children: [headerCell("Issue", 2800), headerCell("Details", 4200), headerCell("Impact", 2360)] }),
            ...[
              ["All columns read as object", "Garbage values (TBD, pending, #REF!, -) prevent type inference", "High"],
              ["Duplicate header row", "1 row repeats column names as values", "Low"],
              ["equipment_type inconsistencies", "32 variants for 7 types (case, abbreviations, HTML tags)", "High"],
              ["Missing values", "weight: 328, cubic_feet: 517, lead_time: 207, bids: 395", "High"],
              ["Invalid target values", "33 zero prices, 15 negative prices (min: -4,504)", "High"],
              ["Outliers", "weight: 1M lbs, price: 63K, lead_time: 9,999 hrs", "Medium"],
              ["Mixed date formats", "YYYY-MM-DD (6,082), MM/DD/YYYY (1,912), with time (1,011)", "Medium"],
              ["Negative lead_time_hours", "4 rows with negative values", "Low"],
            ].map((row, i) => {
              const color = row[2] === "High" ? "C62828" : row[2] === "Medium" ? "E65100" : "2E7D32";
              const bg = row[2] === "High" ? RED_BG : row[2] === "Medium" ? "FFF3E0" : GREEN_BG;
              return new TableRow({ children: [
                dataCell(row[0], 2800, { bold: true, shading: i % 2 === 0 ? LIGHT_GRAY : undefined }),
                dataCell(row[1], 4200, { shading: i % 2 === 0 ? LIGHT_GRAY : undefined }),
                dataCell(row[2], 2360, { bold: true, color, shading: bg }),
              ] });
            }),
          ],
        }),
        new Paragraph({ spacing: { after: 200 } }),

        heading("3.1 Equipment Type Variants (32 found)", HeadingLevel.HEADING_2),
        new Table({
          width: { size: 9360, type: WidthType.DXA },
          columnWidths: [2340, 7020],
          rows: [
            new TableRow({ children: [headerCell("Canonical Type", 2340), headerCell("Variants Found", 7020)] }),
            ...[
              ["Cargo Van", "Cargo Van, cargo van, CARGO VAN, CV, Cargo  Van, <b>Cargo Van</b>"],
              ["Sprinter Van", "Sprinter Van, sprinter van, SPRINTER VAN, Sprinter"],
              ["Small Straight", "Small Straight, small straight, Sm Straight, Small Str"],
              ["Large Straight", "Large Straight, LARGE STRAIGHT, Large Str, Lg Straight"],
              ["Tractor Trailer", "Tractor Trailer, tractor trailer, Tractor, TT"],
              ["Flatbed", "Flatbed, flatbed, FLATBED, Flat Bed"],
              ["Reefer", "Reefer, reefer, REEFER, Refrigerated"],
            ].map((row, i) => new TableRow({ children: [
              dataCell(row[0], 2340, { bold: true, shading: i % 2 === 0 ? LIGHT_GRAY : undefined }),
              dataCell(row[1], 7020, { shading: i % 2 === 0 ? LIGHT_GRAY : undefined }),
            ] })),
          ],
        }),
        new Paragraph({ spacing: { after: 200 } }),

        heading("3.2 Non-Numeric Garbage Values", HeadingLevel.HEADING_2),
        bulletItem("weight_lbs: TBD (3), - (2), #REF! (2), pending (1)", "b1"),
        bulletItem("cubic_feet: pending (5), #REF! (3)", "b1"),
        bulletItem("lead_time_hours, num_carrier_bids, lane_volume: 1 row each with column header as value", "b1"),

        new Paragraph({ children: [new PageBreak()] }),

        // ===== 4. PIPELINE ARCHITECTURE =====
        heading("4. Reusable Pipeline Architecture", HeadingLevel.HEADING_1),
        para("A modular, reusable pipeline was built so new data can be processed without rework. The pipeline is configured via config.py and can be run end-to-end or step-by-step."),

        heading("4.1 Project Structure", HeadingLevel.HEADING_2),
        bulletItem("config.py - Central configuration (paths, mappings, thresholds)", "b2"),
        bulletItem("pipeline/clean.py - 9 cleaning functions + orchestrator", "b2"),
        bulletItem("pipeline/features.py - Feature matrix preparation (minimal for baseline)", "b2"),
        bulletItem("pipeline/model.py - Model training, evaluation, plotting", "b2"),
        bulletItem("run_pipeline.py - Main entry point with CLI arguments", "b2"),
        new Paragraph({ spacing: { after: 100 } }),

        heading("4.2 Usage", HeadingLevel.HEADING_2),
        para("python run_pipeline.py                        # Full pipeline: clean + model", { font: "Consolas", size: 20 }),
        para("python run_pipeline.py --step clean            # Only cleaning", { font: "Consolas", size: 20 }),
        para("python run_pipeline.py --step model            # Only modeling", { font: "Consolas", size: 20 }),
        para("python run_pipeline.py --input new_data.csv    # Run on new data", { font: "Consolas", size: 20 }),
        new Paragraph({ spacing: { after: 100 } }),

        heading("4.3 Cleaning Pipeline Steps", HeadingLevel.HEADING_2),
        numberItem("drop_header_rows() - Remove embedded header rows", "n1"),
        numberItem("drop_null_rows() - Remove fully null rows", "n1"),
        numberItem("standardize_equipment_type() - Map 32 variants to 7 canonical types, strip HTML", "n1"),
        numberItem("coerce_numeric_columns() - Convert to numeric, garbage becomes NaN", "n1"),
        numberItem("remove_invalid_targets() - Drop zero/negative prices", "n1"),
        numberItem("parse_dates() - Unify 3 date formats into datetime", "n1"),
        numberItem("cap_outliers() - Cap extreme values at configurable thresholds", "n1"),
        numberItem("impute_missing() - Median imputation for numeric columns", "n1"),

        new Paragraph({ children: [new PageBreak()] }),

        // ===== 5. ACTUAL CLEANING RESULTS =====
        heading("5. Data Cleaning Results (Actual)", HeadingLevel.HEADING_1),
        para("The pipeline was executed on the raw dataset. Below are the actual results from each cleaning step."),

        heading("5.1 Row Impact Summary", HeadingLevel.HEADING_2),
        new Table({
          width: { size: 9360, type: WidthType.DXA },
          columnWidths: [4000, 2680, 2680],
          rows: [
            new TableRow({ children: [headerCell("Step", 4000), headerCell("Rows Removed", 2680), headerCell("Remaining", 2680)] }),
            ...[
              ["Original dataset", "-", "9,991"],
              ["Drop duplicate header row", "1", "9,990"],
              ["Drop fully null rows", "8", "9,982"],
              ["Standardize equipment_type", "25", "9,957"],
              ["Remove invalid targets (<=0)", "23", "9,934"],
            ].map((row, i) => new TableRow({ children: [
              dataCell(row[0], 4000, { bold: true, shading: i % 2 === 0 ? LIGHT_GRAY : undefined }),
              dataCell(row[1], 2680, { shading: i % 2 === 0 ? LIGHT_GRAY : undefined }),
              dataCell(row[2], 2680, { bold: true, shading: i % 2 === 0 ? LIGHT_GRAY : undefined }),
            ] })),
            new TableRow({ children: [
              dataCell("FINAL CLEANED DATASET", 4000, { bold: true, shading: GREEN_BG }),
              dataCell("57 total removed", 2680, { shading: GREEN_BG }),
              dataCell("9,934 rows (99.4%)", 2680, { bold: true, color: "2E7D32", shading: GREEN_BG }),
            ] }),
          ],
        }),
        new Paragraph({ spacing: { after: 200 } }),

        heading("5.2 Outlier Capping Applied", HeadingLevel.HEADING_2),
        new Table({
          width: { size: 9360, type: WidthType.DXA },
          columnWidths: [2340, 2340, 2340, 2340],
          rows: [
            new TableRow({ children: [headerCell("Column", 2340), headerCell("Threshold", 2340), headerCell("Values Capped", 2340), headerCell("Action", 2340)] }),
            ...[
              ["weight_lbs", "100,000 lbs", "10", "Capped at threshold"],
              ["cubic_feet", "5,000 cu ft", "4", "Capped at threshold"],
              ["lead_time_hours", "500 hrs", "3", "Capped at threshold"],
              ["winning_bid_price", "$50,000", "1", "Capped at threshold"],
              ["lead_time_hours", "< 0", "4", "Set to NaN (imputed)"],
            ].map((row, i) => new TableRow({ children: row.map((val, j) =>
              dataCell(val, 2340, { shading: i % 2 === 0 ? LIGHT_GRAY : undefined, bold: j === 0 })
            ) })),
          ],
        }),
        new Paragraph({ spacing: { after: 200 } }),

        heading("5.3 Missing Value Imputation", HeadingLevel.HEADING_2),
        new Table({
          width: { size: 9360, type: WidthType.DXA },
          columnWidths: [2340, 2340, 2340, 2340],
          rows: [
            new TableRow({ children: [headerCell("Column", 2340), headerCell("Missing Count", 2340), headerCell("Strategy", 2340), headerCell("Fill Value", 2340)] }),
            ...[
              ["weight_lbs", "328", "Median", "1,111.95"],
              ["cubic_feet", "517", "Median", "8.80"],
              ["lead_time_hours", "203", "Median", "13.70"],
              ["num_carrier_bids", "387", "Median", "4.00"],
            ].map((row, i) => new TableRow({ children: row.map((val, j) =>
              dataCell(val, 2340, { shading: i % 2 === 0 ? LIGHT_GRAY : undefined, bold: j === 0 })
            ) })),
          ],
        }),

        new Paragraph({ children: [new PageBreak()] }),

        // ===== 6. BASELINE MODEL RESULTS =====
        heading("6. Baseline Model Results (Actual)", HeadingLevel.HEADING_1),
        multiRunPara([
          { text: "Note: ", bold: true },
          { text: "Feature engineering was intentionally deferred for this baseline. Models use raw numeric features (weight_lbs, cubic_feet, lead_time_hours, num_carrier_bids, lane_volume) and one-hot encoded categoricals (origin_region, destination_region, equipment_type). Total features: 59." },
        ]),
        new Paragraph({ spacing: { after: 100 } }),

        para("Train/test split: 80/20 (7,947 train, 1,987 test). Random seed: 42."),

        heading("6.1 Performance Comparison", HeadingLevel.HEADING_2),
        new Table({
          width: { size: 9360, type: WidthType.DXA },
          columnWidths: [2340, 2340, 2340, 2340],
          rows: [
            new TableRow({ children: [headerCell("Model", 2340), headerCell("MAE", 2340), headerCell("RMSE", 2340), headerCell("R-squared", 2340)] }),
            new TableRow({ children: [
              dataCell("Linear Regression", 2340, { bold: true }),
              dataCell("$831.49", 2340),
              dataCell("$2,457.21", 2340),
              dataCell("-0.3609", 2340, { color: "C62828", bold: true }),
            ] }),
            new TableRow({ children: [
              dataCell("Random Forest", 2340, { bold: true, shading: GREEN_BG }),
              dataCell("$662.19", 2340, { shading: GREEN_BG }),
              dataCell("$1,373.84", 2340, { shading: GREEN_BG }),
              dataCell("0.5746", 2340, { color: "2E7D32", bold: true, shading: GREEN_BG }),
            ] }),
          ],
        }),
        new Paragraph({ spacing: { after: 200 } }),

        heading("6.2 Analysis", HeadingLevel.HEADING_2),
        bulletItem("Linear Regression: Negative R-squared indicates the model performs worse than predicting the mean. This confirms non-linear relationships between features and price.", "b3"),
        bulletItem("Random Forest: Explains 57.5% of variance with MAE of $662. This is a reasonable baseline given no feature engineering.", "b3"),
        bulletItem("The large gap between MAE ($662) and RMSE ($1,374) for Random Forest suggests some large prediction errors on high-value shipments.", "b3"),
        new Paragraph({ spacing: { after: 200 } }),

        heading("6.3 Actual vs Predicted", HeadingLevel.HEADING_2),
        new Paragraph({
          children: [new ImageRun({
            type: "png",
            data: actualVsPredImg,
            transformation: { width: 620, height: 266 },
            altText: { title: "Actual vs Predicted", description: "Scatter plot of actual vs predicted prices", name: "actual_vs_predicted" },
          })],
        }),
        new Paragraph({ spacing: { after: 200 } }),

        heading("6.4 Residual Analysis", HeadingLevel.HEADING_2),
        new Paragraph({
          children: [new ImageRun({
            type: "png",
            data: residualsImg,
            transformation: { width: 620, height: 266 },
            altText: { title: "Residuals", description: "Residual plots for both models", name: "residuals" },
          })],
        }),
        new Paragraph({ spacing: { after: 200 } }),

        heading("6.5 Feature Importance (Random Forest)", HeadingLevel.HEADING_2),
        new Paragraph({
          children: [new ImageRun({
            type: "png",
            data: featureImpImg,
            transformation: { width: 500, height: 400 },
            altText: { title: "Feature Importance", description: "Top 20 feature importances from Random Forest", name: "feature_importance" },
          })],
        }),

        new Paragraph({ children: [new PageBreak()] }),

        // ===== 7. DECISION LOG =====
        heading("7. Decision Log", HeadingLevel.HEADING_1),
        para("This section captures key decisions made during the project, along with rationale and status."),
        new Paragraph({ spacing: { after: 100 } }),

        new Table({
          width: { size: 9360, type: WidthType.DXA },
          columnWidths: [1200, 2700, 3260, 2200],
          rows: [
            new TableRow({ children: [
              headerCell("Date", 1200), headerCell("Decision", 2700), headerCell("Rationale", 3260), headerCell("Status", 2200),
            ] }),
            new TableRow({ children: [
              dataCell("2026-03-26", 1200),
              dataCell("Build reusable pipeline instead of one-off notebooks", 2700, { bold: true }),
              dataCell("Avoid rework when new data arrives. Pipeline is configurable via config.py and can be run via CLI.", 3260),
              dataCell("Implemented", 2200, { color: "2E7D32", bold: true, shading: GREEN_BG }),
            ] }),
            new TableRow({ children: [
              dataCell("2026-03-26", 1200, { shading: LIGHT_GRAY }),
              dataCell("Defer feature engineering to future iteration", 2700, { bold: true, shading: LIGHT_GRAY }),
              dataCell("Establish baseline performance first. Current model uses raw features + one-hot encoding only.", 3260, { shading: LIGHT_GRAY }),
              dataCell("Deferred", 2200, { color: "E65100", bold: true, shading: AMBER_BG }),
            ] }),
            new TableRow({ children: [
              dataCell("2026-03-26", 1200),
              dataCell("Cap outliers instead of removing them", 2700, { bold: true }),
              dataCell("Preserves data volume. Extreme values (weight=1M, lead_time=9999) capped at configurable thresholds.", 3260),
              dataCell("Implemented", 2200, { color: "2E7D32", bold: true, shading: GREEN_BG }),
            ] }),
            new TableRow({ children: [
              dataCell("2026-03-26", 1200, { shading: LIGHT_GRAY }),
              dataCell("Use median imputation for missing values", 2700, { bold: true, shading: LIGHT_GRAY }),
              dataCell("Robust to outliers. Configurable in config.py (can switch to mean).", 3260, { shading: LIGHT_GRAY }),
              dataCell("Implemented", 2200, { color: "2E7D32", bold: true, shading: GREEN_BG }),
            ] }),
            new TableRow({ children: [
              dataCell("2026-03-26", 1200),
              dataCell("Remove rows with winning_bid_price <= 0", 2700, { bold: true }),
              dataCell("Zero and negative prices are invalid for a bid price prediction task. 48 rows removed.", 3260),
              dataCell("Implemented", 2200, { color: "2E7D32", bold: true, shading: GREEN_BG }),
            ] }),
            new TableRow({ children: [
              dataCell("2026-03-26", 1200, { shading: LIGHT_GRAY }),
              dataCell("Document all conversations and decisions in this living document", 2700, { bold: true, shading: LIGHT_GRAY }),
              dataCell("Enables team visibility. Document is updated after each phase with actual results.", 3260, { shading: LIGHT_GRAY }),
              dataCell("Ongoing", 2200, { color: MED_BLUE, bold: true }),
            ] }),
          ],
        }),

        new Paragraph({ children: [new PageBreak()] }),

        // ===== 8. NEXT STEPS =====
        heading("8. Next Steps", HeadingLevel.HEADING_1),

        heading("8.1 Feature Engineering (Deferred)", HeadingLevel.HEADING_2),
        bulletItem("Extract date features: month, day_of_week, day_of_month from shipment_date", "b4"),
        bulletItem("Feature scaling/normalization for linear models", "b4"),
        bulletItem("Interaction features (e.g., weight * distance proxy)", "b4"),
        bulletItem("Lane-level aggregations (historical avg price per lane)", "b4"),
        new Paragraph({ spacing: { after: 100 } }),

        heading("8.2 Model Improvements", HeadingLevel.HEADING_2),
        bulletItem("Hyperparameter tuning for Random Forest (n_estimators, max_depth, min_samples_split)", "b5"),
        bulletItem("Try Gradient Boosting (XGBoost/LightGBM) for potential improvement", "b5"),
        bulletItem("Cross-validation instead of single train/test split", "b5"),
        bulletItem("Investigate high-error predictions (likely high-value shipments)", "b5"),
        new Paragraph({ spacing: { after: 100 } }),

        heading("8.3 Data Quality Improvements", HeadingLevel.HEADING_2),
        bulletItem("Investigate root cause of garbage values (TBD, #REF!, pending) at the data source", "b6"),
        bulletItem("Validate equipment_type mapping completeness when new data arrives", "b6"),
        bulletItem("Consider time-based train/test split instead of random split", "b6"),
      ],
    },
  ],
});

Packer.toBuffer(doc).then((buffer) => {
  fs.writeFileSync("E:/MyProject/AscentGlobalLogistics/Ascent_Freight_Pricing_Data_Assessment_v2.docx", buffer);
  console.log("Document created successfully!");
});
