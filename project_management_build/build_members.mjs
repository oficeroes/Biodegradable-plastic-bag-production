import fs from 'node:fs/promises';
import assert from 'node:assert/strict';
import { FileBlob, SpreadsheetFile, Workbook } from '@oai/artifact-tool';

const outputDir = 'D:/UserData/Documents/小龙虾制成塑料袋/outputs/01a0f6b6-782d-7380-8c3a-668a2556b9b7';
let outputPath = `${outputDir}/项目成员表.xlsx`;
const previous = await SpreadsheetFile.importXlsx(await FileBlob.load(outputPath));
const oldValues = previous.worksheets.getItem('成员表').getRange('A1:K16').values;
const headerRow = oldValues.findIndex(row => row.includes('姓名'));
assert.ok(headerRow >= 0, 'Existing roster header was not found');
const fields = ['姓名', '班级', '学号'];
const columns = fields.map(field => oldValues[headerRow].indexOf(field));
const rows = oldValues.slice(headerRow + 1, headerRow + 10).map(row =>
  columns.map(col => col < 0 ? null : row[col] ?? null)
);

const workbook = Workbook.create();
const sheet = workbook.worksheets.add('成员表');
sheet.showGridLines = false;
sheet.getRange('A1:C10').format = {
  font: { name: 'Arial', size: 11, color: '#24292F' },
  verticalAlignment: 'center',
  rowHeight: 32,
  borders: { preset: 'all', style: 'thin', color: '#CCD1D6' },
};
sheet.getRange('A1:C1').values = [fields];
sheet.getRange('A1:C1').format = {
  fill: '#384452',
  font: { name: 'Arial', size: 11, bold: true, color: '#FFFFFF' },
  horizontalAlignment: 'center',
};
sheet.getRange('A2:C10').values = rows;
sheet.getRange('A2:C10').setNumberFormat('@');
[18, 22, 22].forEach((width, col) => {
  sheet.getRangeByIndexes(0, col, 10, 1).format.columnWidth = width;
});
workbook.recalculate();
const preview = await workbook.render({ sheetName: '成员表', range: 'A1:C10', scale: 1.5, format: 'png' });
await fs.writeFile(`${outputDir}/成员表预览.png`, new Uint8Array(await preview.arrayBuffer()));
const output = await SpreadsheetFile.exportXlsx(workbook);
try {
  await output.save(outputPath);
} catch (error) {
  if (error.code !== 'EBUSY') throw error;
  outputPath = `${outputDir}/成员名单.xlsx`;
  await output.save(outputPath);
}
const saved = await SpreadsheetFile.importXlsx(await FileBlob.load(outputPath));
assert.deepEqual(saved.worksheets.getItem('成员表').getRange('A1:C1').values[0], fields);
assert.deepEqual(saved.worksheets.getItem('成员表').getRange('A2:C10').values, rows);
console.log(JSON.stringify({ outputPath, columns: fields, rows: 9 }));
