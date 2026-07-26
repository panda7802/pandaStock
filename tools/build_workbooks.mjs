import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const ROOT = process.cwd();
const RESULT_JSON = path.join(ROOT, "策略回测输出", "回测中间结果", "backtests.json");
const OUTPUT_DIR = path.join(ROOT, "策略回测输出", "单标的Excel");
const PREVIEW_DIR = path.join(ROOT, "策略回测输出", "回测中间结果", "Excel预览");

const COLORS = {
  navy: "#17365D",
  blue: "#2E74B5",
  paleBlue: "#DCE6F1",
  paleGray: "#F2F4F7",
  midGray: "#667085",
  border: "#D0D5DD",
  white: "#FFFFFF",
  positive: "#067647",
  negative: "#B42318",
};

const METRIC_LABELS = [
  ["期末累计收益", "final_return"],
  ["年化收益（CAGR）", "cagr"],
  ["最大回撤", "max_drawdown"],
  ["年化波动", "volatility"],
  ["仓位占比", "invested_pct"],
  ["调仓次数", "trades"],
];

function a1Column(index) {
  let n = index + 1;
  let text = "";
  while (n > 0) {
    n -= 1;
    text = String.fromCharCode(65 + (n % 26)) + text;
    n = Math.floor(n / 26);
  }
  return text;
}

function sanitizeSheetText(value) {
  return String(value ?? "").replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, "");
}

function sourceText(ticker, metadata) {
  const leveraged = {
    TQQQ: ["QQQ", 3],
    SQQQ: ["QQQ", -3],
    UDOW: ["DIA", 3],
    SDOW: ["DIA", -3],
    UPRO: ["SPY", 3],
    SPXU: ["SPY", -3],
  };
  if (leveraged[ticker]) {
    const [underlying, leverage] = leveraged[ticker];
    return `基于 ${underlying} 的逐日总收益，按每日 ${leverage > 0 ? "+" : ""}${leverage} 倍复利模拟；底层来源：${metadata.sources[underlying]}`;
  }
  return metadata.sources[ticker];
}

async function buildWorkbook(record, metadata, renderPreview = false) {
  const workbook = Workbook.create();
  const summary = workbook.worksheets.add("收益对比");
  const notes = workbook.worksheets.add("数据与说明");
  summary.showGridLines = false;
  notes.showGridLines = false;

  summary.mergeCells("A1:N1");
  summary.getRange("A1").values = [[`${record.ticker}｜${record.start_year} 起点｜三种策略收益对比`]];
  summary.getRange("A1:N1").format = {
    fill: COLORS.navy,
    font: { bold: true, color: COLORS.white, size: 18 },
    horizontalAlignment: "left",
    verticalAlignment: "center",
  };
  summary.getRange("A1:N1").format.rowHeight = 34;

  summary.mergeCells("A2:N2");
  summary.getRange("A2").values = [[
    `实际起始交易日：${record.start_date}　数据截至：${record.end_date}　口径：月初首个交易日的累计收益`,
  ]];
  summary.getRange("A2:N2").format = {
    fill: COLORS.paleBlue,
    font: { color: COLORS.navy, size: 10 },
    verticalAlignment: "center",
  };
  summary.getRange("A2:N2").format.rowHeight = 24;

  summary.getRange("A4:G4").values = [[
    "策略", ...METRIC_LABELS.map(([label]) => label),
  ]];
  summary.getRange("A4:G4").format = {
    fill: COLORS.blue,
    font: { bold: true, color: COLORS.white },
    horizontalAlignment: "center",
    verticalAlignment: "center",
    wrapText: true,
    borders: { preset: "all", style: "thin", color: COLORS.border },
  };
  summary.getRange("A5:A7").values = [["A"], ["B"], ["C"]];
  summary.getRange("A5:A7").format = {
    fill: COLORS.paleGray,
    font: { bold: true, color: COLORS.navy },
    horizontalAlignment: "center",
    borders: { preset: "all", style: "thin", color: COLORS.border },
  };
  const summaryFormulas = [];
  for (let rowIndex = 0; rowIndex < 3; rowIndex += 1) {
    const formulas = [];
    for (let colIndex = 0; colIndex < METRIC_LABELS.length; colIndex += 1) {
      formulas.push(`='数据与说明'!${a1Column(colIndex + 1)}${12 + rowIndex}`);
    }
    summaryFormulas.push(formulas);
  }
  summary.getRange("B5:G7").formulas = summaryFormulas;
  summary.getRange("B5:F7").format.numberFormat = "0.0%";
  summary.getRange("G5:G7").format.numberFormat = "0";
  summary.getRange("B5:G7").format = {
    horizontalAlignment: "right",
    borders: { preset: "all", style: "thin", color: COLORS.border },
  };
  summary.getRange("B5:F7").conditionalFormats.add(
    "cellIs",
    { operator: "greaterThanOrEqual", formula: 0, format: { font: { color: COLORS.positive } } },
  );
  summary.getRange("B5:F7").conditionalFormats.add(
    "cellIs",
    { operator: "lessThan", formula: 0, format: { font: { color: COLORS.negative } } },
  );

  summary.mergeCells("A9:N9");
  summary.getRange("A9").values = [[
    "图表提示：若杠杆标的长期收益跨度极大，Excel 原生图采用线性坐标；同名 PNG 使用对称对数坐标以保留细节。",
  ]];
  summary.getRange("A9:N9").format = {
    fill: "#FFF7E6",
    font: { color: "#7A2E0E", italic: true, size: 9 },
    wrapText: true,
  };

  const monthlyStartRow = 12;
  const monthlyRows = record.monthly.map((row) => [
    row.date.slice(0, 7),
    row.A,
    row.B,
    row.C,
  ]);
  const monthlyEndRow = monthlyStartRow + monthlyRows.length;
  summary.getRange(`A${monthlyStartRow}:D${monthlyStartRow}`).values = [[
    "月份", "A 策略累计收益", "B 策略累计收益", "C 策略累计收益",
  ]];
  summary.getRange(`A${monthlyStartRow}:D${monthlyStartRow}`).format = {
    fill: COLORS.navy,
    font: { bold: true, color: COLORS.white },
    horizontalAlignment: "center",
    verticalAlignment: "center",
    borders: { preset: "all", style: "thin", color: COLORS.border },
  };
  summary.getRange(`A${monthlyStartRow + 1}:D${monthlyEndRow}`).values = monthlyRows;
  summary.getRange(`B${monthlyStartRow + 1}:D${monthlyEndRow}`).format.numberFormat = "0.0%";
  summary.getRange(`A${monthlyStartRow + 1}:D${monthlyEndRow}`).format.borders = {
    preset: "all",
    style: "thin",
    color: "#E4E7EC",
  };
  summary.getRange(`B${monthlyStartRow + 1}:D${monthlyEndRow}`).conditionalFormats.add(
    "cellIs",
    { operator: "greaterThanOrEqual", formula: 0, format: { font: { color: COLORS.positive } } },
  );
  summary.getRange(`B${monthlyStartRow + 1}:D${monthlyEndRow}`).conditionalFormats.add(
    "cellIs",
    { operator: "lessThan", formula: 0, format: { font: { color: COLORS.negative } } },
  );
  summary.freezePanes.freezeRows(monthlyStartRow);

  summary.getRange("A:A").format.columnWidth = 14;
  summary.getRange("B:F").format.columnWidth = 17;
  summary.getRange("G:G").format.columnWidth = 12;
  summary.getRange("H:N").format.columnWidth = 12;
  summary.getRange("A4:G7").format.rowHeight = 24;

  const chart = summary.charts.add(
    "line",
    summary.getRange(`A${monthlyStartRow}:D${monthlyEndRow}`),
  );
  chart.title = `${record.ticker} 三种策略累计收益`;
  chart.titleTextStyle.fontSize = 13;
  chart.hasLegend = true;
  chart.xAxis = { axisType: "textAxis", textStyle: { fontSize: 9 } };
  chart.yAxis = { numberFormatCode: "0.0%" };
  chart.setPosition("F12", "N31");

  notes.mergeCells("A1:G1");
  notes.getRange("A1").values = [["数据、口径与复核信息"]];
  notes.getRange("A1:G1").format = {
    fill: COLORS.navy,
    font: { bold: true, color: COLORS.white, size: 16 },
    verticalAlignment: "center",
  };
  notes.getRange("A1:G1").format.rowHeight = 32;
  notes.getRange("A3:B9").values = [
    ["项目", "说明"],
    ["标的", record.ticker],
    ["用户指定起始年", record.start_year],
    ["实际首个交易日", record.start_date],
    ["数据截止日", record.end_date],
    ["数据来源/模拟方法", sanitizeSheetText(sourceText(record.ticker, metadata))],
    ["月度观测", "每月第一个交易日；表内数值为相对起始净值的累计收益"],
  ];
  notes.getRange("A3:B3").format = {
    fill: COLORS.blue,
    font: { bold: true, color: COLORS.white },
    borders: { preset: "all", style: "thin", color: COLORS.border },
  };
  notes.getRange("A4:B9").format = {
    wrapText: true,
    verticalAlignment: "top",
    borders: { preset: "all", style: "thin", color: COLORS.border },
  };

  notes.getRange("A11:G11").values = [[
    "策略", ...METRIC_LABELS.map(([label]) => label),
  ]];
  notes.getRange("A11:G11").format = {
    fill: COLORS.blue,
    font: { bold: true, color: COLORS.white },
    horizontalAlignment: "center",
    wrapText: true,
    borders: { preset: "all", style: "thin", color: COLORS.border },
  };
  const metricRows = ["A", "B", "C"].map((strategy) => [
    strategy,
    ...METRIC_LABELS.map(([, key]) => record.metrics[strategy][key]),
  ]);
  notes.getRange("A12:G14").values = metricRows;
  notes.getRange("A12:A14").format = {
    fill: COLORS.paleGray,
    font: { bold: true, color: COLORS.navy },
    horizontalAlignment: "center",
  };
  notes.getRange("B12:F14").format.numberFormat = "0.0%";
  notes.getRange("G12:G14").format.numberFormat = "0";
  notes.getRange("A12:G14").format.borders = {
    preset: "all",
    style: "thin",
    color: COLORS.border,
  };

  notes.getRange("A16:B22").values = [
    ["规则/假设", "具体口径"],
    ["策略 A", metadata.strategies.A],
    ["策略 B", metadata.strategies.B],
    ["策略 C", metadata.strategies.C],
    ["信号执行", metadata.assumptions[1]],
    ["成本与现金", metadata.assumptions[2]],
    ["杠杆历史", metadata.assumptions[3]],
  ];
  notes.getRange("A16:B16").format = {
    fill: COLORS.blue,
    font: { bold: true, color: COLORS.white },
    borders: { preset: "all", style: "thin", color: COLORS.border },
  };
  notes.getRange("A17:B22").format = {
    wrapText: true,
    verticalAlignment: "top",
    borders: { preset: "all", style: "thin", color: COLORS.border },
  };
  notes.getRange("A:A").format.columnWidth = 18;
  notes.getRange("B:B").format.columnWidth = 86;
  notes.getRange("C:G").format.columnWidth = 16;
  notes.getRange("A3:G22").format.rowHeight = 24;
  notes.getRange("A6:A9").format.rowHeight = 42;
  notes.getRange("A17:B22").format.rowHeight = 42;
  notes.freezePanes.freezeRows(3);

  const formulaErrors = await workbook.inspect({
    kind: "match",
    searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
    options: { useRegex: true, maxResults: 100 },
    summary: `${record.ticker}-${record.start_year} formula error scan`,
  });
  if (formulaErrors.ndjson && !formulaErrors.ndjson.includes('"count":0')) {
    const lines = formulaErrors.ndjson.trim().split(/\r?\n/).filter(Boolean);
    const suspicious = lines.filter((line) => /#REF!|#DIV\/0!|#VALUE!|#NAME\?|#N\/A/.test(line));
    if (suspicious.length) {
      throw new Error(`Formula error in ${record.ticker}-${record.start_year}: ${suspicious[0]}`);
    }
  }

  if (renderPreview) {
    const previewSummary = await workbook.render({
      sheetName: "收益对比",
      range: "A1:N31",
      scale: 1,
      format: "png",
    });
    await fs.writeFile(
      path.join(PREVIEW_DIR, `${record.ticker}-${record.start_year}-收益对比.png`),
      new Uint8Array(await previewSummary.arrayBuffer()),
    );
    const previewNotes = await workbook.render({
      sheetName: "数据与说明",
      range: "A1:G22",
      scale: 1,
      format: "png",
    });
    await fs.writeFile(
      path.join(PREVIEW_DIR, `${record.ticker}-${record.start_year}-数据与说明.png`),
      new Uint8Array(await previewNotes.arrayBuffer()),
    );
  }

  const xlsx = await SpreadsheetFile.exportXlsx(workbook);
  await xlsx.save(path.join(OUTPUT_DIR, `${record.ticker}-${record.start_year}—三种策略收益对比.xlsx`));
}

await fs.mkdir(OUTPUT_DIR, { recursive: true });
await fs.mkdir(PREVIEW_DIR, { recursive: true });
const payload = JSON.parse(await fs.readFile(RESULT_JSON, "utf8"));
const previewKeys = new Set(["QQQ-2000", "TQQQ-2000", "SQQQ-2023"]);

let completed = 0;
for (const record of payload.results) {
  const key = `${record.ticker}-${record.start_year}`;
  await buildWorkbook(record, payload.metadata, previewKeys.has(key));
  completed += 1;
  if (completed % 11 === 0 || completed === payload.results.length) {
    process.stdout.write(`已生成 ${completed}/${payload.results.length} 份 Excel\n`);
  }
}

process.stdout.write(`完成：${completed} 份 Excel\n`);
