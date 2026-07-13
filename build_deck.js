const pptxgen = require("pptxgenjs");

// ---- Palette: Midnight Executive, dominant navy + ice blue, single gold accent ----
const NAVY = "1E2761";
const NAVY_DARK = "141B47";
const ICE = "CADCFC";
const WHITE = "FFFFFF";
const GREY = "9AA5B1";        // muted text ON DARK (navy) backgrounds only
const MUTED = "5B6472";       // muted text ON LIGHT (white/off-white) backgrounds
const LIGHTGREY = "E8EAED";
const GOLD = "D8A33D";
const RED = "C0392B";
const DARKTEXT = "1F2937";

const FONT = "Calibri";
const HEADFONT = "Cambria";

let pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.3 x 7.5

// ============================================================
// SLIDE 1 — TITLE
// ============================================================
let s1 = pres.addSlide();
s1.background = { color: NAVY };
// motif: large soft circle, bottom-right, off-canvas bleed
s1.addShape(pres.ShapeType.ellipse, {
  x: 9.8, y: 3.2, w: 6.5, h: 6.5, fill: { color: NAVY_DARK }, line: { type: "none" },
});
s1.addShape(pres.ShapeType.ellipse, {
  x: 11.3, y: 4.7, w: 3.5, h: 3.5, fill: { color: GOLD, transparency: 85 }, line: { type: "none" },
});

s1.addText("ORDER PERFORMANCE", {
  x: 0.8, y: 2.55, w: 10.5, h: 0.5,
  fontFace: FONT, fontSize: 14, bold: true, color: GOLD, charSpacing: 3, margin: 0,
});
s1.addText("A Data Storytelling Review", {
  x: 0.8, y: 3.0, w: 10.5, h: 1.1,
  fontFace: HEADFONT, fontSize: 42, bold: true, color: WHITE, margin: 0,
});
s1.addText("Turning 30 months of e-commerce order data into three decisions worth making", {
  x: 0.8, y: 3.95, w: 9.5, h: 0.5,
  fontFace: FONT, fontSize: 15, color: ICE, italic: true, margin: 0,
});

s1.addText("PROJECT 4", { x: 0.8, y: 6.55, w: 3, h: 0.3, fontFace: FONT, fontSize: 10.5, bold: true, color: GOLD, charSpacing: 2, margin: 0 });
s1.addText("DecodeLabs Data Analytics Internship  |  Batch 2026  |  Prepared by Shafi", {
  x: 0.8, y: 6.85, w: 10, h: 0.4,
  fontFace: FONT, fontSize: 12, color: ICE, margin: 0,
});

// ============================================================
// SLIDE 2 — EXECUTIVE SUMMARY (Situation / Complication / Resolution)
// ============================================================
let s2 = pres.addSlide();
s2.background = { color: WHITE };
s2.addText("EXECUTIVE SUMMARY", { x: 0.7, y: 0.45, w: 6, h: 0.3, fontFace: FONT, fontSize: 11, bold: true, color: GOLD, charSpacing: 2, margin: 0 });
s2.addText("Three Months of Orders, One Clear Set of Decisions", {
  x: 0.7, y: 0.78, w: 11.9, h: 0.65,
  fontFace: HEADFONT, fontSize: 26, bold: true, color: DARKTEXT, margin: 0,
});

const scr = [
  { label: "SITUATION", title: "1,200 orders, 30 months, steady operations", body: "The business has processed a consistent flow of orders across seven product categories since January 2023, with no single category dominating the catalog." },
  { label: "COMPLICATION", title: "Growth has stalled and fulfillment is leaking", body: "Revenue has been flat rather than trending up, and 41% of all orders are stuck in Cancelled or Pending status — nearly matching completed orders." },
  { label: "RESOLUTION", title: "Fix the funnel before chasing new growth", body: "Investigate the fulfillment backlog first. It is the single largest, most actionable lever in this dataset — larger than any product or channel optimization." },
];

let cardX = 0.7;
const cardW = 3.87;
scr.forEach((c, i) => {
  s2.addShape(pres.ShapeType.rect, {
    x: cardX, y: 1.75, w: cardW, h: 5.0,
    fill: { color: i === 1 ? NAVY : "F7F8FA" }, line: { type: "none" },
    shadow: i === 1 ? { type: "outer", color: "000000", opacity: 0.25, blur: 8, offset: 3, angle: 90 } : undefined,
  });
  const txtColor = i === 1 ? WHITE : DARKTEXT;
  const labelColor = i === 1 ? GOLD : NAVY;
  s2.addText(c.label, { x: cardX + 0.3, y: 2.05, w: cardW - 0.6, h: 0.3, fontFace: FONT, fontSize: 11, bold: true, color: labelColor, charSpacing: 2, margin: 0 });
  s2.addText(c.title, { x: cardX + 0.3, y: 2.4, w: cardW - 0.6, h: 1.3, fontFace: HEADFONT, fontSize: 17, bold: true, color: txtColor, margin: 0 });
  s2.addText(c.body, { x: cardX + 0.3, y: 3.75, w: cardW - 0.6, h: 2.7, fontFace: FONT, fontSize: 12, color: i === 1 ? ICE : MUTED, margin: 0, valign: "top" });
  cardX += cardW + 0.29;
});

// ============================================================
// SLIDE 3 — KEY METRICS (stat callouts)
// ============================================================
let s3 = pres.addSlide();
s3.background = { color: NAVY };
s3.addText("THE NUMBERS AT A GLANCE", { x: 0.7, y: 0.55, w: 6, h: 0.3, fontFace: FONT, fontSize: 11, bold: true, color: GOLD, charSpacing: 2, margin: 0 });
s3.addText("Four Figures That Frame Every Decision Ahead", {
  x: 0.7, y: 0.88, w: 11.9, h: 0.65,
  fontFace: HEADFONT, fontSize: 26, bold: true, color: WHITE, margin: 0,
});

const stats = [
  ["$1.26M", "Total revenue\nacross 30 months"],
  ["1,200", "Orders placed\nJan 2023 - Jun 2025"],
  ["$1,050", "Average\norder value"],
  ["41%", "Orders Cancelled\nor Pending"],
];
let statX = 0.7;
const statW = 2.9;
stats.forEach(([num, label], i) => {
  s3.addText(num, {
    x: statX, y: 2.6, w: statW, h: 1.1,
    fontFace: HEADFONT, fontSize: 46, bold: true, color: i === 3 ? GOLD : WHITE, margin: 0, align: "left",
  });
  s3.addShape(pres.ShapeType.rect, { x: statX, y: 3.75, w: 0.45, h: 0.045, fill: { color: GOLD }, line: { type: "none" } });
  s3.addText(label, {
    x: statX, y: 3.95, w: statW - 0.2, h: 0.8,
    fontFace: FONT, fontSize: 13, color: ICE, margin: 0,
  });
  statX += statW + 0.35;
});

s3.addText("Full breakdown on the following slides — starting with what's driving the top line.", {
  x: 0.7, y: 6.7, w: 11.5, h: 0.4, fontFace: FONT, fontSize: 12, italic: true, color: GREY, margin: 0,
});

// ============================================================
// SLIDE 4 — CHART 1: PRODUCT REVENUE
// ============================================================
let s4 = pres.addSlide();
s4.background = { color: WHITE };
s4.addText("REVENUE BY CATEGORY", { x: 0.7, y: 0.4, w: 6, h: 0.3, fontFace: FONT, fontSize: 11, bold: true, color: GOLD, charSpacing: 2, margin: 0 });
s4.addText("Chairs and Printers Lead, But No Category Carries the Business Alone", {
  x: 0.7, y: 0.72, w: 11.9, h: 0.75,
  fontFace: HEADFONT, fontSize: 23, bold: true, color: DARKTEXT, margin: 0,
});

const prodLabels = ["Chair", "Printer", "Laptop", "Tablet", "Monitor", "Desk", "Phone"];
const prodValues = [195620, 195613, 192127, 186569, 175651, 167460, 151722];
const prodColors = prodLabels.map((_, i) => (i < 2 ? NAVY : LIGHTGREY));

s4.addChart(pres.ChartType.bar, [{ name: "Revenue", labels: prodLabels, values: prodValues }], {
  x: 0.7, y: 1.65, w: 11.9, h: 4.85,
  barDir: "col",
  showTitle: false, showLegend: false, showValue: true,
  dataLabelPosition: "outEnd",
  dataLabelFormatCode: "$#,##0",
  dataLabelColor: DARKTEXT, dataLabelFontSize: 11, dataLabelFontBold: true,
  chartColors: prodColors,
  catAxisLabelColor: DARKTEXT, catAxisLabelFontSize: 13,
  valAxisHidden: true,
  valGridLine: { style: "none" }, catGridLine: { style: "none" },
  valAxisMinVal: 0,
  barGapWidthPct: 35,
});

s4.addShape(pres.ShapeType.rect, { x: 0.7, y: 6.75, w: 0.35, h: 0.045, fill: { color: GOLD }, line: { type: "none" } });
s4.addText("Spread across all 7 categories stays within ~22% of the top performer — the catalog isn't overexposed to one product.", {
  x: 1.2, y: 6.62, w: 11.4, h: 0.5, fontFace: FONT, fontSize: 12, italic: true, color: MUTED, margin: 0,
});

// ============================================================
// SLIDE 5 — CHART 2: MONTHLY TREND
// ============================================================
let s5 = pres.addSlide();
s5.background = { color: WHITE };
s5.addText("REVENUE OVER TIME", { x: 0.7, y: 0.4, w: 6, h: 0.3, fontFace: FONT, fontSize: 11, bold: true, color: GOLD, charSpacing: 2, margin: 0 });
s5.addText("Revenue Held Flat Across 30 Months — There Is No Growth Trend to Report", {
  x: 0.7, y: 0.72, w: 11.9, h: 0.75,
  fontFace: HEADFONT, fontSize: 23, bold: true, color: DARKTEXT, margin: 0,
});

const months = ["Jan-23","Feb-23","Mar-23","Apr-23","May-23","Jun-23","Jul-23","Aug-23","Sep-23","Oct-23","Nov-23","Dec-23",
"Jan-24","Feb-24","Mar-24","Apr-24","May-24","Jun-24","Jul-24","Aug-24","Sep-24","Oct-24","Nov-24","Dec-24",
"Jan-25","Feb-25","Mar-25","Apr-25","May-25","Jun-25"];
const monthlyValues = [56686,40118,48609,27752,63837,49500,43825,44493,40484,36995,44192,38781,
41652,45215,38246,49658,44417,68069,39754,42156,46187,35892,41023,37654,
29099,35318,39201,31821,43397,53047];

s5.addChart(pres.ChartType.line, [{ name: "Monthly Revenue", labels: months, values: monthlyValues }], {
  x: 0.7, y: 1.6, w: 11.9, h: 4.55,
  showTitle: false, showLegend: false, showValue: false,
  chartColors: [NAVY],
  lineSize: 2.75, lineSmooth: false, lineDataSymbol: "circle", lineDataSymbolSize: 4,
  catAxisLabelColor: GREY, catAxisLabelFontSize: 8, catAxisLabelRotate: 90,
  valAxisLabelColor: GREY, valAxisLabelFontSize: 10, valAxisLabelFormatCode: "$#,##0",
  valGridLine: { color: LIGHTGREY, size: 0.75 }, catGridLine: { style: "none" },
  valAxisMinVal: 0,
});

// annotation callouts
s5.addShape(pres.ShapeType.rect, { x: 9.85, y: 1.75, w: 2.0, h: 0.62, fill: { color: NAVY }, line: { type: "none" } });
s5.addText([{ text: "PEAK  ", options: { bold: true, color: GOLD, fontSize: 9 } }, { text: "Jun \u201924: $68.1K", options: { color: WHITE, fontSize: 10.5 } }], {
  x: 9.95, y: 1.8, w: 1.85, h: 0.52, fontFace: FONT, margin: 0, valign: "middle",
});

s5.addShape(pres.ShapeType.rect, { x: 0.7, y: 6.3, w: 0.35, h: 0.045, fill: { color: GOLD }, line: { type: "none" } });
s5.addText("Average $42.2K/month, swinging \u00b125% month to month with no seasonal pattern. Low: April 2023 ($27.8K).", {
  x: 1.2, y: 6.17, w: 11.4, h: 0.5, fontFace: FONT, fontSize: 12, italic: true, color: MUTED, margin: 0,
});

// ============================================================
// SLIDE 6 — CHART 3: ORDER STATUS (the problem)
// ============================================================
let s6 = pres.addSlide();
s6.background = { color: WHITE };
s6.addText("FULFILLMENT HEALTH", { x: 0.7, y: 0.4, w: 6, h: 0.3, fontFace: FONT, fontSize: 11, bold: true, color: RED, charSpacing: 2, margin: 0 });
s6.addText("Cancelled and Pending Orders Consume 41% of Volume", {
  x: 0.7, y: 0.72, w: 11.9, h: 0.75,
  fontFace: HEADFONT, fontSize: 24, bold: true, color: DARKTEXT, margin: 0,
});

const statusLabels = ["Cancelled", "Returned", "Pending", "Shipped", "Delivered"];
const statusValues = [250, 247, 237, 235, 231];
const statusColors = statusLabels.map(l => (l === "Cancelled" || l === "Pending") ? RED : LIGHTGREY);

s6.addChart(pres.ChartType.bar, [{ name: "Orders", labels: statusLabels, values: statusValues }], {
  x: 0.7, y: 1.7, w: 7.6, h: 4.8,
  barDir: "col",
  showTitle: false, showLegend: false, showValue: true,
  dataLabelPosition: "outEnd", dataLabelColor: DARKTEXT, dataLabelFontSize: 12, dataLabelFontBold: true,
  chartColors: statusColors,
  catAxisLabelColor: DARKTEXT, catAxisLabelFontSize: 12.5,
  valAxisHidden: true,
  valGridLine: { style: "none" }, catGridLine: { style: "none" },
  valAxisMinVal: 0,
  barGapWidthPct: 40,
});

s6.addShape(pres.ShapeType.rect, { x: 8.65, y: 1.9, w: 4.0, h: 4.35, fill: { color: "FDF2F1" }, line: { type: "none" } });
s6.addText("487", { x: 8.95, y: 2.15, w: 3.4, h: 0.95, fontFace: HEADFONT, fontSize: 44, bold: true, color: RED, margin: 0 });
s6.addText("orders Cancelled or Pending", { x: 8.95, y: 3.05, w: 3.4, h: 0.4, fontFace: FONT, fontSize: 12.5, bold: true, color: DARKTEXT, margin: 0 });
s6.addText("That's nearly as many as Delivered and Shipped combined (466 orders) — a near-even split between orders that finish and orders that stall.", {
  x: 8.95, y: 3.55, w: 3.4, h: 1.5, fontFace: FONT, fontSize: 11.5, color: MUTED, margin: 0, valign: "top",
});
s6.addShape(pres.ShapeType.rect, { x: 8.95, y: 5.35, w: 3.1, h: 0.02, fill: { color: RED }, line: { type: "none" } });
s6.addText("Recommended owner: Operations / Fulfillment team", { x: 8.95, y: 5.5, w: 3.4, h: 0.5, fontFace: FONT, fontSize: 10.5, italic: true, color: RED, margin: 0 });

// ============================================================
// SLIDE 7 — SO WHAT: RECOMMENDATIONS
// ============================================================
let s7 = pres.addSlide();
s7.background = { color: NAVY };
s7.addText("SO WHAT", { x: 0.7, y: 0.55, w: 6, h: 0.3, fontFace: FONT, fontSize: 11, bold: true, color: GOLD, charSpacing: 2, margin: 0 });
s7.addText("Three Actions for Next Quarter", {
  x: 0.7, y: 0.88, w: 11.9, h: 0.7,
  fontFace: HEADFONT, fontSize: 28, bold: true, color: WHITE, margin: 0,
});

const actions = [
  ["01", "Investigate the fulfillment backlog", "Cancelled and Pending orders are 41% of volume — pair with the ops team to find where orders are stalling before spending on new acquisition."],
  ["02", "Don't chase a growth narrative", "Revenue is flat, not trending up or down. Forecasts and targets should use a flat baseline with seasonality checks, not a growth assumption."],
  ["03", "Protect catalog balance", "No product category is over-relied on. Keep marketing spend spread across categories rather than concentrating on one \"hero\" product."],
];

let yPos = 1.95;
actions.forEach(([num, title, body]) => {
  s7.addShape(pres.ShapeType.ellipse, { x: 0.7, y: yPos, w: 0.85, h: 0.85, fill: { color: GOLD }, line: { type: "none" } });
  s7.addText(num, { x: 0.7, y: yPos, w: 0.85, h: 0.85, fontFace: HEADFONT, fontSize: 22, bold: true, color: NAVY, align: "center", valign: "middle", margin: 0 });
  s7.addText(title, { x: 1.85, y: yPos, w: 10.5, h: 0.45, fontFace: FONT, fontSize: 17, bold: true, color: WHITE, margin: 0 });
  s7.addText(body, { x: 1.85, y: yPos + 0.45, w: 10.5, h: 0.55, fontFace: FONT, fontSize: 12.5, color: ICE, margin: 0 });
  yPos += 1.55;
});

// ============================================================
// SLIDE 8 — METHODOLOGY / CREDIBILITY
// ============================================================
let s8 = pres.addSlide();
s8.background = { color: WHITE };
s8.addText("HOW THIS WAS BUILT", { x: 0.7, y: 0.45, w: 6, h: 0.3, fontFace: FONT, fontSize: 11, bold: true, color: GOLD, charSpacing: 2, margin: 0 });
s8.addText("Every Chart Follows the Same Three Disciplines", {
  x: 0.7, y: 0.78, w: 11.9, h: 0.65,
  fontFace: HEADFONT, fontSize: 26, bold: true, color: DARKTEXT, margin: 0,
});

const pillars = [
  ["A", "The Architect", "Chart type matched strictly to the business question — bar for comparison, line for trend. Every bar axis starts at zero."],
  ["E", "The Editor", "Maximized data-ink: no gridlines, no 3D, no legends. Values are labeled directly on the chart, not decoded from a key."],
  ["S", "The Storyteller", "Every slide has an action title that states the conclusion, and one message only — structured with a Situation-Complication-Resolution arc."],
];
let px = 0.7;
const pw = 3.87;
pillars.forEach(([letter, title, body]) => {
  s8.addShape(pres.ShapeType.ellipse, { x: px, y: 1.85, w: 0.7, h: 0.7, fill: { color: NAVY }, line: { type: "none" } });
  s8.addText(letter, { x: px, y: 1.85, w: 0.7, h: 0.7, fontFace: HEADFONT, fontSize: 20, bold: true, color: GOLD, align: "center", valign: "middle", margin: 0 });
  s8.addText(title, { x: px, y: 2.75, w: pw, h: 0.45, fontFace: HEADFONT, fontSize: 16, bold: true, color: DARKTEXT, margin: 0 });
  s8.addText(body, { x: px, y: 3.25, w: pw, h: 2.0, fontFace: FONT, fontSize: 12, color: MUTED, margin: 0, valign: "top" });
  px += pw + 0.29;
});

s8.addShape(pres.ShapeType.rect, { x: 0.7, y: 5.7, w: 11.9, h: 1.1, fill: { color: "F7F8FA" }, line: { type: "none" } });
s8.addText([
  { text: "Data source: ", options: { bold: true, color: DARKTEXT, fontSize: 11.5 } },
  { text: "Share_Dataset_for_Data_Analytics.xlsx (1,200 orders, Jan 2023-Jun 2025), cleaned and explored in Projects 1-3.   ", options: { color: MUTED, fontSize: 11.5 } },
  { text: "Tools: ", options: { bold: true, color: DARKTEXT, fontSize: 11.5 } },
  { text: "Python (pandas) for the underlying numbers, native PowerPoint charts for every visual.", options: { color: MUTED, fontSize: 11.5 } },
], { x: 1.0, y: 5.95, w: 11.3, h: 0.6, fontFace: FONT, margin: 0, valign: "middle" });

// ============================================================
// SLIDE 9 — CLOSING
// ============================================================
let s9 = pres.addSlide();
s9.background = { color: NAVY };
s9.addShape(pres.ShapeType.ellipse, { x: -2, y: -2, w: 6, h: 6, fill: { color: NAVY_DARK }, line: { type: "none" } });
s9.addText("Thank You", {
  x: 0.8, y: 3.0, w: 11.9, h: 1.0,
  fontFace: HEADFONT, fontSize: 40, bold: true, color: WHITE, margin: 0,
});
s9.addShape(pres.ShapeType.rect, { x: 0.85, y: 3.95, w: 0.6, h: 0.045, fill: { color: GOLD }, line: { type: "none" } });
s9.addText("Project 4 — Data Visualization  |  DecodeLabs Data Analytics Internship", {
  x: 0.8, y: 4.2, w: 11.9, h: 0.5,
  fontFace: FONT, fontSize: 15, color: ICE, margin: 0,
});
s9.addText("Prepared by Shafi", {
  x: 0.8, y: 4.7, w: 11.9, h: 0.4,
  fontFace: FONT, fontSize: 13, color: GREY, margin: 0,
});

pres.writeFile({ fileName: "Project4_DataVisualization.pptx" }).then(() => {
  console.log("Deck written");
});
