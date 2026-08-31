import fs from 'node:fs'
import path from 'node:path'
import ExcelJS from 'exceljs'

const workbook = new ExcelJS.Workbook()
workbook.creator = 'Lutherska Missionskyrkan'
workbook.title = 'Inventarieregister'

const addTableSheet = (name, columns, rows, widths) => {
  const sheet = workbook.addWorksheet(name, {
    views: [{ state: 'frozen', ySplit: 1 }],
    properties: { defaultRowHeight: 19 },
  })
  sheet.addTable({
    name,
    ref: 'A1',
    headerRow: true,
    totalsRow: false,
    style: { theme: 'TableStyleMedium2', showRowStripes: true },
    columns: columns.map((column) => ({ name: column, filterButton: true })),
    rows,
  })
  widths.forEach((width, index) => { sheet.getColumn(index + 1).width = width })
  sheet.getRow(1).height = 24
  sheet.getRow(1).font = { name: 'Alata', bold: true, color: { argb: 'FFFFFFFF' } }
  sheet.eachRow((row, rowNumber) => {
    if (rowNumber > 1) row.font = { name: 'Calibri', size: 11 }
  })
  return sheet
}

addTableSheet('Categories', ['Id', 'Name', 'Color', 'Prefix'], [
  ['cat-furniture', 'Möbler', '#6f8657', 'MOB'],
  ['cat-lighting', 'Belysning', '#f59b45', 'BEL'],
  ['cat-sound', 'Ljudutrustning', '#3688df', 'LJU'],
  ['cat-kitchen', 'Kök & fika', '#c73f3b', 'KOK'],
], [22, 28, 16, 14])

addTableSheet('Groups', ['Id', 'Name'], [
  ['grp-board-mission', 'Styrelsen Lutherska missionsföreningen'],
  ['grp-board-hagaparken', 'Styrelsen Hagaparken Ekonomisk förening'],
  ['grp-premises-change', 'Arbetsgrupp för lokalförändringar'],
  ['grp-sound', 'Ljudgruppen'],
  ['grp-service-visuals', 'Bildvisning i gudstjänst'],
  ['grp-kitchen-cleaning-purchases', 'Inköp för kök och städ'],
  ['grp-allergy-purchases', 'Allergiinköp'],
  ['grp-household', 'Husmor/far'],
  ['grp-holiday-decoration', 'Utsmyckning storhelger'],
  ['grp-music-committee', 'Musikutskottet'],
  ['grp-service-committee', 'Gudstjänstutskottet'],
  ['grp-children-youth-committee', 'Barn- och ungdomsutskottet'],
  ['grp-orchestra', 'Lutherska Missionskyrkans orkester'],
  ['grp-choir', 'Lutherska Missionskyrkans kör'],
  ['grp-choir-council', 'Lutherska Missionskyrkans körråd'],
  ['grp-sunday-school', 'Söndagsskolan'],
  ['grp-forest-school', 'Skogsskolan'],
  ['grp-baby-rhythmics', 'Babyrytmik'],
  ['grp-clap-and-sound', 'Klapp och klang'],
  ['grp-popkidz', 'Popkidz'],
  ['grp-tweenies', 'Tweenies'],
  ['grp-lux', 'LUX'],
  ['grp-gamla-barn', 'Gamla Barn'],
], [34, 48])

addTableSheet('Locations', ['Id', 'Name'], [
  ['loc-sanctuary', 'Kyrksalen'],
  ['loc-stage', 'Scenförrådet'],
  ['loc-kitchen', 'Köksförrådet'],
  ['loc-basement', 'Källarförrådet'],
], [22, 32])

const inventorySheet = addTableSheet('Inventory', ['Id', 'AssetTag', 'Name', 'CategoryId', 'PrimaryGroupId', 'SecondaryGroupIds', 'LocationId', 'Quantity', 'Notes'], [], [30, 16, 36, 22, 24, 30, 22, 12, 44])
inventorySheet.dataValidations.add('B2:B5000', {
  type: 'custom',
  allowBlank: true,
  formulae: ['=OR(COUNTA(A2:I2)=0,AND(LEN(B2)=7,MID(B2,4,1)="-",CODE(MID(B2,1,1))>=65,CODE(MID(B2,1,1))<=90,CODE(MID(B2,2,1))>=65,CODE(MID(B2,2,1))<=90,CODE(MID(B2,3,1))>=65,CODE(MID(B2,3,1))<=90,ISNUMBER(--RIGHT(B2,3)),COUNTIF($B:$B,B2)=1,D2<>"",LEFT(B2,3)=IFERROR(VLOOKUP(D2,Categories,4,FALSE),"")))'],
  showErrorMessage: true,
  errorStyle: 'error',
  errorTitle: 'Ogiltigt inventarienummer',
  error: 'Använd kategorins trebokstavskod, bindestreck och tre siffror. Numret måste vara unikt, till exempel MOB-014.',
})
addTableSheet('Loans', ['Id', 'ItemId', 'Borrower', 'BorrowerGroupId', 'RecordedBy', 'LentAt', 'DueAt', 'ReturnedAt'], [], [30, 30, 28, 24, 28, 14, 14, 14])

const outputDirectory = path.resolve('workbook')
fs.mkdirSync(outputDirectory, { recursive: true })
const outputPath = path.join(outputDirectory, 'Lutherska-Inventarier.xlsx')
await workbook.xlsx.writeFile(outputPath)
console.log(`Created ${outputPath}`)