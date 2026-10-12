import fs from "node:fs/promises";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outputDir = "D:/MyPython/fill_inForm/test_templates";
const outputPath = `${outputDir}/多工作表申请材料.xlsx`;
const font = "Microsoft YaHei";
const colors = {
  navy: "#1F4E78",
  blue: "#D9EAF7",
  paleBlue: "#EDF5FB",
  input: "#FFF6D9",
  border: "#B7C9D6",
  note: "#666666",
};

function applyBase(sheet) {
  sheet.showGridLines = false;
}

function styleTitle(sheet, range, text) {
  sheet.mergeCells(range);
  const cell = sheet.getRange(range);
  cell.values = [[text]];
  cell.format = {
    font: { name: font, size: 16, bold: true, color: "#1F1F1F" },
    horizontalAlignment: "left",
    verticalAlignment: "center",
  };
  cell.format.rowHeight = 28;
}

function styleSection(sheet, range, text) {
  sheet.mergeCells(range);
  const cell = sheet.getRange(range);
  cell.values = [[text]];
  cell.format = {
    fill: colors.navy,
    font: { name: font, size: 10.5, bold: true, color: "#FFFFFF" },
    horizontalAlignment: "left",
    verticalAlignment: "center",
  };
  cell.format.rowHeight = 22;
}

function styleLabel(range) {
  range.format = {
    fill: colors.blue,
    font: { name: font, size: 10.5, bold: true, color: "#173B5D" },
    horizontalAlignment: "center",
    verticalAlignment: "center",
    borders: { preset: "all", style: "thin", color: colors.border },
  };
}

function styleInput(range) {
  range.format = {
    fill: colors.input,
    font: { name: font, size: 10.5, color: "#6E571D" },
    horizontalAlignment: "left",
    verticalAlignment: "center",
    wrapText: true,
    borders: { preset: "all", style: "thin", color: colors.border },
  };
}

function setWidths(sheet, widths) {
  for (const [column, width] of Object.entries(widths)) {
    sheet.getRange(`${column}:${column}`).format.columnWidth = width;
  }
}

const workbook = Workbook.create();
const guide = workbook.worksheets.add("填表说明");
const basic = workbook.worksheets.add("基本信息");
const work = workbook.worksheets.add("工作与项目");
const statement = workbook.worksheets.add("个人陈述");
const options = workbook.worksheets.add("选项测试");

for (const sheet of [guide, basic, work, statement, options]) applyBase(sheet);

// Sheet 1: no fill targets; it verifies that explanatory text remains unchanged.
setWidths(guide, { A: 5, B: 24, C: 72 });
styleTitle(guide, "A2:C2", "多工作表申请材料");
guide.getRange("A4:C4").values = [["序号", "工作表", "测试目的"]];
styleLabel(guide.getRange("A4:C4"));
guide.getRange("A5:C9").values = [
  [1, "基本信息", "左右标签 + 空白输入格 + 合并长文本格"],
  [2, "工作与项目", "字段在上方的项目经历表格"],
  [3, "个人陈述", "300 字和 500 字长文本生成"],
  [4, "选项测试", "下拉选项、日期、公式格和应跳过的说明文字"],
  [5, "填写规则", "浅黄色单元格为预期填写位置；蓝色区域为标签或表头"],
];
guide.getRange("A5:C9").format = {
  font: { name: font, size: 10.5 },
  verticalAlignment: "center",
  wrapText: true,
  borders: { preset: "all", style: "thin", color: colors.border },
};
guide.getRange("A5:A9").format.horizontalAlignment = "center";
guide.getRange("A5:C9").format.rowHeight = 28;
guide.getRange("B11").values = [["说明：本工作表只包含提示信息，不应被自动填写。"]];
guide.getRange("B11").format = { font: { name: font, size: 10, italic: true, color: colors.note } };

// Sheet 2: left/right labels, top labels, and a merged long-text target.
setWidths(basic, { A: 16, B: 18, C: 16, D: 18, E: 16, F: 18 });
styleTitle(basic, "A2:F2", "基本信息");
styleSection(basic, "A4:F4", "基础资料");
basic.getRange("A5:F8").values = [
  ["姓名", "", "性别", "", "出生日期", ""],
  ["联系电话", "", "电子邮箱", "", "常住城市", ""],
  ["工作单位", "", "职位", "", "工作年限", ""],
  ["最高学历", "", "专业方向", "", "是否接受出差", ""],
];
for (const range of ["A5", "C5", "E5", "A6", "C6", "E6", "A7", "C7", "E7", "A8", "C8", "E8"]) styleLabel(basic.getRange(range));
for (const range of ["B5", "D5", "F5", "B6", "D6", "F6", "B7", "D7", "F7", "B8", "D8", "F8"]) styleInput(basic.getRange(range));
basic.getRange("A5:F8").format.rowHeight = 24;
styleSection(basic, "A10:F10", "专业技能");
basic.getRange("A11").values = [["专业技能（100 字以内）"]];
styleLabel(basic.getRange("A11"));
basic.mergeCells("B11:F14");
basic.getRange("B11:F14").values = [[""]];
styleInput(basic.getRange("B11:F14"));
basic.getRange("B11:F14").format.verticalAlignment = "top";

// Sheet 3: headers above blank data cells.
setWidths(work, { A: 6, B: 22, C: 16, D: 18, E: 44 });
styleTitle(work, "A2:E2", "工作与项目经历");
work.getRange("A4:E4").values = [["序号", "项目名称", "项目时间", "担任角色", "核心贡献或成果"]];
styleLabel(work.getRange("A4:E4"));
work.getRange("A5:E7").values = [
  [1, "", "", "", ""],
  [2, "", "", "", ""],
  [3, "", "", "", ""],
];
work.getRange("A5:A7").format = {
  fill: colors.paleBlue,
  horizontalAlignment: "center",
  verticalAlignment: "center",
  borders: { preset: "all", style: "thin", color: colors.border },
};
styleInput(work.getRange("B5:E7"));
work.getRange("A5:E7").format.rowHeight = 42;
work.getRange("A9").values = [["说明"]];
styleLabel(work.getRange("A9"));
work.mergeCells("B9:E9");
work.getRange("B9:E9").values = [["请填写具有代表性的项目。此行是说明文字，不应覆盖。"]];
work.getRange("B9:E9").format = {
  font: { name: font, size: 10, italic: true, color: colors.note },
  fill: "#F7F9FB",
  borders: { preset: "all", style: "thin", color: colors.border },
};

// Sheet 4: two large AI-generated fields with explicit length instructions.
setWidths(statement, { A: 22, B: 18, C: 18, D: 18, E: 18, F: 18, G: 18, H: 18 });
styleTitle(statement, "A2:H2", "个人陈述");
styleSection(statement, "A4:H4", "个人工作简介");
statement.getRange("A5").values = [["个人工作简介（300 字以内）"]];
styleLabel(statement.getRange("A5"));
statement.mergeCells("B5:H11");
statement.getRange("B5:H11").values = [[""]];
styleInput(statement.getRange("B5:H11"));
statement.getRange("B5:H11").format.verticalAlignment = "top";
styleSection(statement, "A13:H13", "工作心得");
statement.getRange("A14").values = [["工作心得（500 字以内）"]];
styleLabel(statement.getRange("A14"));
statement.mergeCells("B14:H25");
statement.getRange("B14:H25").values = [[""]];
styleInput(statement.getRange("B14:H25"));
statement.getRange("B14:H25").format.verticalAlignment = "top";

// Sheet 5: data validation and formula cells should not be overwritten blindly.
setWidths(options, { A: 24, B: 26, C: 24, D: 24 });
styleTitle(options, "A2:D2", "选项与边界测试");
options.getRange("A4:D4").values = [["字段", "填写位置", "系统处理预期", "备注"]];
styleLabel(options.getRange("A4:D4"));
options.getRange("A5:D8").values = [
  ["是否接受出差", "", "可尝试填写，但应尊重下拉选项", "下拉：是 / 否 / 视情况"],
  ["可入职日期", "", "可按资料填写日期", "日期格式"],
  ["已填写字段数", "", "公式格，不应被覆盖", "依赖 B5:B6"],
  ["备注", "", "无明确标签时应提示人工确认", "边界用例"],
];
styleLabel(options.getRange("A5:A8"));
styleInput(options.getRange("B5:B6"));
styleInput(options.getRange("B8"));
options.getRange("C5:D8").format = {
  fill: "#F7F9FB",
  font: { name: font, size: 10, color: "#444444" },
  wrapText: true,
  verticalAlignment: "center",
  borders: { preset: "all", style: "thin", color: colors.border },
};
options.getRange("B7").formulas = [["=COUNTA(B5:B6)"]];
options.getRange("B7").format = {
  fill: "#EAF2F8",
  font: { name: font, size: 10.5, bold: true, color: "#173B5D" },
  horizontalAlignment: "center",
  verticalAlignment: "center",
  borders: { preset: "all", style: "thin", color: colors.border },
};
options.getRange("B5").dataValidation = { rule: { type: "list", values: ["是", "否", "视情况"] } };
options.getRange("B6").format.numberFormat = "yyyy-mm-dd";
options.getRange("A5:D8").format.rowHeight = 32;

for (const sheet of [guide, basic, work, statement, options]) {
  sheet.tabColor = colors.navy;
}

workbook.recalculate();
const check = await workbook.inspect({
  kind: "workbook,sheet,table",
  maxChars: 4000,
  tableMaxRows: 12,
  tableMaxCols: 10,
});
console.log(check.ndjson);

await fs.mkdir(outputDir, { recursive: true });
const preview = await workbook.render({ sheetName: "基本信息", range: "A2:F14", scale: 1.2, format: "png" });
await fs.writeFile(`${outputDir}/多工作表申请材料_基本信息预览.png`, new Uint8Array(await preview.arrayBuffer()));
const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);
console.log(outputPath);
