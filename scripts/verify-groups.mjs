import fs from 'node:fs'
import ExcelJS from 'exceljs'

const expectedGroups = [
  'Styrelsen Lutherska missionsföreningen',
  'Styrelsen Hagaparken Ekonomisk förening',
  'Arbetsgrupp för lokalförändringar',
  'Ljudgruppen',
  'Bildvisning i gudstjänst',
  'Inköp för kök och städ',
  'Allergiinköp',
  'Husmor/far',
  'Utsmyckning storhelger',
  'Musikutskottet',
  'Gudstjänstutskottet',
  'Barn- och ungdomsutskottet',
  'Lutherska Missionskyrkans orkester',
  'Lutherska Missionskyrkans kör',
  'Lutherska Missionskyrkans körråd',
  'Söndagsskolan',
  'Skogsskolan',
  'Babyrytmik',
  'Klapp och klang',
  'Popkidz',
  'Tweenies',
  'LUX',
  'Gamla Barn',
]

const workbook = new ExcelJS.Workbook()
await workbook.xlsx.readFile('workbook/Lutherska-Inventarier.xlsx')
const sheet = workbook.getWorksheet('Groups')
const workbookGroups = []

for (let rowNumber = 2; rowNumber <= sheet.rowCount; rowNumber += 1) {
  const value = sheet.getCell(rowNumber, 2).text
  if (value) workbookGroups.push(value)
}

const source = fs.readFileSync('src/App.tsx', 'utf8')
const missingInApp = expectedGroups.filter((name) => !source.includes(`name: '${name}'`))
const workbookMatches = JSON.stringify(workbookGroups) === JSON.stringify(expectedGroups)

console.log(`App groups: ${expectedGroups.length - missingInApp.length}/${expectedGroups.length}`)
console.log(`Workbook groups: ${workbookGroups.length}/${expectedGroups.length}`)
console.log(`Exact order and spelling: ${workbookMatches ? 'yes' : 'no'}`)

if (missingInApp.length || !workbookMatches) {
  if (missingInApp.length) console.error('Missing in app:', missingInApp)
  if (!workbookMatches) console.error('Workbook groups:', workbookGroups)
  process.exit(1)
}