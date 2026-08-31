import fs from 'node:fs/promises'
import path from 'node:path'
import ExcelJS from 'exceljs'

const workbookPath = path.resolve('workbook/Lutherska-Inventarier.xlsx')
const outputPath = path.resolve('test-results/workbook-preview.html')
const workbook = new ExcelJS.Workbook()
await workbook.xlsx.readFile(workbookPath)

const escapeHtml = (value) => String(value ?? '')
  .replaceAll('&', '&amp;')
  .replaceAll('<', '&lt;')
  .replaceAll('>', '&gt;')
  .replaceAll('"', '&quot;')

const columnLetter = (index) => {
  let result = ''
  for (let value = index; value > 0; value = Math.floor((value - 1) / 26)) {
    result = String.fromCharCode(65 + ((value - 1) % 26)) + result
  }
  return result
}

const renderSheet = (sheet) => {
  const columnCount = sheet.columnCount
  const rows = []
  for (let rowNumber = 1; rowNumber <= Math.max(sheet.rowCount, 2); rowNumber += 1) {
    const cells = []
    for (let column = 1; column <= columnCount; column += 1) {
      const cell = sheet.getCell(rowNumber, column)
      cells.push(`<td>${escapeHtml(cell.text)}</td>`)
    }
    rows.push(`<tr><th class="row-number">${rowNumber}</th>${cells.join('')}</tr>`)
  }

  const letters = Array.from({ length: columnCount }, (_, index) => `<th>${columnLetter(index + 1)}</th>`).join('')
  const tabs = workbook.worksheets.map((candidate) => `<span class="tab${candidate.name === sheet.name ? ' active' : ''}">${escapeHtml(candidate.name)}</span>`).join('')
  return `<section class="sheet" id="${escapeHtml(sheet.name)}">
    <header><span class="excel-mark">X</span><strong>Lutherska-Inventarier.xlsx</strong><span>Sparad arbetsboksmall</span></header>
    <div class="formula"><span>fx</span><span>${escapeHtml(sheet.name)}</span></div>
    <div class="grid"><table><thead><tr><th class="corner"></th>${letters}</tr></thead><tbody>${rows.join('')}</tbody></table></div>
    <footer>${tabs}</footer>
  </section>`
}

const html = `<!doctype html>
<html lang="sv"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<style>
*{box-sizing:border-box}body{margin:0;padding:40px;background:#dde1e5;color:#222;font-family:Calibri,"Segoe UI",sans-serif}.sheet{width:max-content;min-width:960px;margin:0 0 48px;background:#fff;box-shadow:0 14px 38px rgba(31,51,79,.2);overflow:hidden}header{height:48px;display:flex;align-items:center;gap:13px;padding:0 18px;color:#fff;background:#217346;font-size:16px}header span:last-child{margin-left:auto;font-size:13px;opacity:.85}.excel-mark{display:grid;width:28px;height:28px;place-items:center;background:#185c37;font-weight:700}.formula{height:36px;display:flex;align-items:center;border-bottom:1px solid #c9c9c9;background:#fafafa}.formula span:first-child{width:44px;text-align:center;color:#666;font-style:italic}.formula span:last-child{min-width:220px;padding:4px 10px;border-left:1px solid #ddd;font-size:13px}.grid{overflow:hidden}table{border-collapse:collapse;table-layout:fixed;font-size:13px}th,td{height:29px;padding:5px 9px;border-right:1px solid #d6d6d6;border-bottom:1px solid #d6d6d6;white-space:nowrap;text-align:left}thead th{height:24px;padding:2px 8px;text-align:center;color:#555;background:#f1f1f1;font-weight:400}.corner,.row-number{width:42px;min-width:42px;text-align:center!important;color:#555!important;background:#f1f1f1!important;font-weight:400!important}tbody tr:first-child td{color:#fff;background:#4472c4;font-weight:700}tbody tr:nth-child(odd):not(:first-child) td{background:#d9e2f3}td{min-width:150px}td:nth-child(2){min-width:240px}footer{height:42px;display:flex;align-items:end;gap:2px;padding:0 14px;background:#f3f3f3;border-top:1px solid #ccc}.tab{padding:10px 14px 8px;color:#555;font-size:12px}.tab.active{color:#185c37;background:#fff;border-bottom:3px solid #217346;font-weight:700}
</style></head><body>${workbook.worksheets.map(renderSheet).join('')}</body></html>`

await fs.mkdir(path.dirname(outputPath), { recursive: true })
await fs.writeFile(outputPath, html)
console.log(`Rendered ${outputPath}`)