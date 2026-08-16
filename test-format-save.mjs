// Self-check: verify align/bg/font/colWidths/rowHeights persist through readXlsx/writeXlsx.
import { readXlsx, writeXlsx } from 'file:///C:/Users/82019/.dsh/profiles/web/node_modules/dsh-excel-panel/lib/index.js'
import ExcelJS from 'file:///C:/Users/82019/.dsh/profiles/web/node_modules/exceljs/excel.js'

const file = 'D:/apps/DP专武/.dsh/self-check-format.xlsx'
const wb = new ExcelJS.Workbook()
const ws = wb.addWorksheet('测试')
ws.getCell('A1').value = 'hello'
await wb.xlsx.writeFile(file)

const payload = {
  sheets: [{
    name: '测试',
    rows: [[{ v: 'hello', align: 'center', bg: 'FFFFFF00', font: { bold: true, color: 'FFFF0000' } }]],
    colWidths: { 0: 120 },
    rowHeights: { 0: 55 },
  }],
}
await writeXlsx(file, payload)
const data = await readXlsx(file)
const cell = data.sheets[0].rows[0][0]
const checks = {
  align: cell.align === 'center',
  bg: cell.bg === 'FFFFFF00',
  bold: cell.font && cell.font.bold === true,
  color: cell.font && cell.font.color === 'FFFF0000',
  colWidth: data.sheets[0].colWidths && data.sheets[0].colWidths[0] === 120,
  rowHeight: data.sheets[0].rowHeights && data.sheets[0].rowHeights[0] === 55,
}
console.log(JSON.stringify(checks, null, 2))
const ok = Object.values(checks).every(Boolean)
if (!ok) process.exit(1)
console.log('ALL FORMAT SAVE CHECKS PASSED')
