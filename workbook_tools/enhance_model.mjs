import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const inputPath = "../assets/financial-model.xlsx";
const input = await FileBlob.load(inputPath);
const wb = await SpreadsheetFile.importXlsx(input);

const sheetInfo = await wb.inspect({ kind: "sheet", include: "id,name", maxChars: 4000 });
console.log(sheetInfo.ndjson);

if (process.argv.includes("--inspect")) {
  for (const name of ["Assumptions", "Revenue", "Cost Model", "P&L", "Cash Flow", "Sensitivity"]) {
    try {
      const report = await wb.inspect({ kind: "table", range: `${name}!A1:J30`, include: "values,formulas", tableMaxRows: 30, tableMaxCols: 10, maxChars: 8000 });
      console.log(report.ndjson);
      const preview = await wb.render({ sheetName: name, autoCrop: "all", scale: 1, format: "png" });
      await fs.writeFile(`../workbook_previews/${name.replaceAll(" ", "_").replaceAll("&", "and")}.png`, new Uint8Array(await preview.arrayBuffer()));
    } catch (error) {
      console.log(`SKIP ${name}: ${error.message}`);
    }
  }
  process.exit(0);
}

const navy = "#17365D", blue = "#4472C4", pale = "#F3F6FA", gold = "#FFF2CC", green = "#E2F0D9";
const white = "#FFFFFF", red = "#FCE4D6";

let analysis;
try { analysis = wb.worksheets.getItem("Decision Analysis"); } catch { analysis = wb.worksheets.add("Decision Analysis"); }
analysis.deleteAllDrawings();
analysis.showGridLines = false;
analysis.getRange("A1:H2").merge();
analysis.getRange("A1").values = [["Commercial Decision Analysis | Break-even, Unit Economics & Runway"]];
analysis.getRange("A1:H2").format = { fill: navy, font: { bold: true, color: white, size: 16 }, verticalAlignment: "center" };

analysis.getRange("A4:B10").values = [
  ["Decision inputs", "Base case"],
  ["Annual subscription / provider", 24000],
  ["Gross margin", 0.781],
  ["Annual fixed operating cost", 421000],
  ["Estimated CAC / provider", 18000],
  ["Annual provider churn", 0.10],
  ["Starting funding", 375000],
];
analysis.getRange("A4:B4").format = { fill: blue, font: { bold: true, color: white } };
analysis.getRange("B5:B10").format = { fill: gold, font: { color: "#0000FF" } };
analysis.getRange("B5").format.numberFormat = '£#,##0';
analysis.getRange("B6").format.numberFormat = "0.0%";
analysis.getRange("B7:B8").format.numberFormat = '£#,##0';
analysis.getRange("B9").format.numberFormat = "0.0%";
analysis.getRange("B10").format.numberFormat = '£#,##0';

analysis.getRange("D4:E10").values = [
  ["Key decision outputs", "Base case"],
  ["Gross profit / provider", null],
  ["Break-even providers", null],
  ["Illustrative provider LTV", null],
  ["LTV / CAC", null],
  ["CAC payback (months)", null],
  ["FY3 providers vs break-even", null],
];
analysis.getRange("D4:E4").format = { fill: blue, font: { bold: true, color: white } };
analysis.getRange("E5:E10").formulas = [["=B5*B6"], ["=ROUNDUP(B7/E5,0)"], ["=E5/B9"], ["=E7/B8"], ["=B8/E5*12"], ["=15-E6"]];
analysis.getRange("E5").format.numberFormat = '£#,##0';
analysis.getRange("E6").format.numberFormat = "0";
analysis.getRange("E7").format.numberFormat = '£#,##0';
analysis.getRange("E8").format.numberFormat = "0.0x";
analysis.getRange("E9:E10").format.numberFormat = "0.0";
analysis.getRange("D4:E10").format.borders = { preset: "all", style: "thin", color: "#D9E1F2" };

analysis.getRange("A13:D18").values = [
  ["Price scenario", "Annual price", "Break-even providers", "FY3 gap / surplus"],
  ["Downside", 12000, null, null],
  ["Base", 24000, null, null],
  ["Upside", 36000, null, null],
  ["Premium", 48000, null, null],
  ["Interpretation", null, null, null],
];
analysis.getRange("A13:D13").format = { fill: blue, font: { bold: true, color: white }, wrapText: true };
for (let r = 14; r <= 17; r++) {
  analysis.getRange(`C${r}`).formulas = [[`=ROUNDUP($B$7/(B${r}*$B$6),0)`]];
  analysis.getRange(`D${r}`).formulas = [[`=15-C${r}`]];
}
analysis.getRange("B14:B17").format.numberFormat = '£#,##0';
analysis.getRange("C14:D17").format.numberFormat = "0";
analysis.getRange("A18:D18").merge();
analysis.getRange("A18").values = [["At the £24K base price, FY3 adoption remains below operating break-even; pricing alone does not solve the gap without tighter fixed costs, higher provider count or phased hiring."]];
analysis.getRange("A18:D18").format = { fill: red, font: { italic: true, color: "#9C0006" }, wrapText: true };

analysis.getRange("F13:H18").values = [
  ["Risk register", "Trigger", "Mitigation"],
  ["Slow provider adoption", "<6 providers by FY2", "Stage hiring; extend pilots"],
  ["Price resistance", "Deals below £24K", "Tiered modules; evidence-led ROI"],
  ["High implementation cost", ">£6K per provider", "Standard onboarding and templates"],
  ["Regulatory scope creep", "Diagnostic/dosing features", "Maintain non-device boundary"],
  ["Cash runway pressure", "<9 months cash", "Raise earlier; gate discretionary spend"],
];
analysis.getRange("F13:H13").format = { fill: blue, font: { bold: true, color: white } };
analysis.getRange("F14:H18").format.wrapText = true;
analysis.getRange("F13:H18").format.borders = { preset: "all", style: "thin", color: "#D9E1F2" };

analysis.getRange("A21:E26").values = [
  ["Runway scenario", "Starting funding", "Annual cash burn", "Indicative runway (months)", "Decision"],
  ["Downside", 375000, 240000, null, "Raise before month 12"],
  ["Base", 375000, 180000, null, "Begin raise by month 15"],
  ["Upside", 375000, 120000, null, "Evidence before next round"],
  ["Lean launch", 375000, 90000, null, "Preserve optionality"],
  ["Note", null, null, null, "Illustrative planning scenarios; replace burn with approved operating plan"],
];
analysis.getRange("A21:E21").format = { fill: blue, font: { bold: true, color: white }, wrapText: true };
for (let r = 22; r <= 25; r++) analysis.getRange(`D${r}`).formulas = [[`=B${r}/C${r}*12`]];
analysis.getRange("B22:C25").format.numberFormat = '£#,##0';
analysis.getRange("D22:D25").format.numberFormat = "0.0";
analysis.getRange("A26:E26").format = { fill: pale, font: { italic: true }, wrapText: true };

analysis.getRange("A4:H26").format.borders = { preset: "all", style: "thin", color: "#D9E1F2" };
analysis.getRange("A:A").format.columnWidth = 28;
analysis.getRange("B:E").format.columnWidth = 20;
analysis.getRange("F:F").format.columnWidth = 24;
analysis.getRange("G:H").format.columnWidth = 28;
analysis.getRange("A1:H26").format.verticalAlignment = "center";
analysis.freezePanes.freezeRows(3);

const inspect = await wb.inspect({ kind: "table", range: "Decision Analysis!A4:H26", include: "values,formulas", tableMaxRows: 30, tableMaxCols: 8 });
console.log(inspect.ndjson);
const errors = await wb.inspect({ kind: "match", searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A", options: { useRegex: true, maxResults: 200 }, summary: "formula error scan" });
console.log(errors.ndjson);

const preview = await wb.render({ sheetName: "Decision Analysis", autoCrop: "all", scale: 1, format: "png" });
await fs.writeFile("../workbook_previews/Decision_Analysis.png", new Uint8Array(await preview.arrayBuffer()));
const out = await SpreadsheetFile.exportXlsx(wb);
await out.save(inputPath);
console.log(`Saved ${inputPath}`);
