import ExcelJS from 'exceljs'

const DEEP_BLUE = '1E3A5F'
const GOLD = 'C8952E'
const HEADER_FG = 'FFFFFF'
const ROW_EVEN = 'F4F6F9'
const ROW_ODD = 'FFFFFF'

const APP_COLS = [
  { header: 'رقم الطلب', key: 'requestNumber', width: 16 },
  { header: 'اسم الطالب', key: 'studentName', width: 22 },
  { header: 'كود الطالب', key: 'studentCode', width: 16 },
  { header: 'المدرسة السابقة', key: 'previousSchool', width: 24 },
  { header: 'رقم الطالب', key: 'studentPhone', width: 16 },
  { header: 'رقم ولي الأمر', key: 'parentPhone', width: 16 },
  { header: 'الحالة الاجتماعية', key: 'socialStatus', width: 16 },
  { header: 'حالة الطلب', key: 'appStatus', width: 16 },
  { header: 'تاريخ التقديم', key: 'submittedAt', width: 18 },
  { header: 'الاسم بالإنجليزية', key: 'fullNameEn', width: 20 },
  { header: 'تاريخ الميلاد', key: 'dateOfBirth', width: 14 },
  { header: 'الجنس', key: 'gender', width: 10 },
  { header: 'الجنسية', key: 'nationality', width: 14 },
  { header: 'الديانة', key: 'religion', width: 12 },
  { header: 'العنوان', key: 'address', width: 28 },
  { header: 'البريد الإلكتروني', key: 'email', width: 26 },
  { header: 'رقم بطاقة الطالب', key: 'studentId', width: 18 },
  { header: 'اسم ولي الأمر', key: 'parentName', width: 20 },
  { header: 'رقم بطاقة ولي الأمر', key: 'parentId', width: 18 },
  { header: 'بريد ولي الأمر', key: 'parentEmail', width: 26 },
]

async function main() {
  const wb = new ExcelJS.Workbook()
  wb.creator = 'WE Schools'
  wb.created = new Date()

  // ──── THEME ────
  const headerFont = { name: 'Calibri', bold: true, color: { argb: HEADER_FG }, size: 11 }
  const headerFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: DEEP_BLUE } }
  const headerAlign = { horizontal: 'center', vertical: 'middle' }
  const centerAlign = { horizontal: 'center', vertical: 'middle' }
  const borderStyle = {
    top: { style: 'thin', color: { argb: 'D0D5DD' } },
    left: { style: 'thin', color: { argb: 'D0D5DD' } },
    bottom: { style: 'thin', color: { argb: 'D0D5DD' } },
    right: { style: 'thin', color: { argb: 'D0D5DD' } },
  }

  // ──── Helper: style header row ────
  function styleHeader(ws, numCols) {
    const row = ws.getRow(1)
    row.height = 30
    for (let c = 1; c <= numCols; c++) {
      const cell = row.getCell(c)
      cell.font = headerFont
      cell.fill = headerFill
      cell.alignment = headerAlign
      cell.border = borderStyle
    }
  }

  // ──── Helper: style data rows ────
  function styleDataRows(ws, numCols, startRow) {
    ws.eachRow((row, rowNum) => {
      if (rowNum < startRow) return
      row.eachCell((cell, colNum) => {
        if (colNum > numCols) return
        cell.border = borderStyle
        cell.alignment = { ...centerAlign, wrapText: true }
        const bg = (rowNum - startRow) % 2 === 0 ? ROW_EVEN : ROW_ODD
        if (!cell.fill || cell.fill.fgColor?.argb !== DEEP_BLUE) {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } }
        }
      })
    })
  }

  // ════════════════════════════════════════
  // Sheet 1: Applications
  // ════════════════════════════════════════
  const app = wb.addWorksheet('Applications', {
    views: [{ state: 'frozen', ySplit: 1 }],
    properties: { tabColor: { argb: DEEP_BLUE } },
  })

  app.columns = APP_COLS
  styleHeader(app, APP_COLS.length)

  // Sample data row (empty)
  const sampleRow = {}
  APP_COLS.forEach(c => { sampleRow[c.key] = '' })
  app.addRow(sampleRow)
  styleDataRows(app, APP_COLS.length, 2)

  // Data validation: social status (G) and app status (H)
  app.getColumn(7).eachCell((cell, rowNum) => {
    if (rowNum > 1) {
      cell.dataValidation = {
        type: 'list',
        formulae: ['"أعزب,متزوج"'],
        allowBlank: true,
      }
    }
  })
  app.getColumn(8).eachCell((cell, rowNum) => {
    if (rowNum > 1) {
      cell.dataValidation = {
        type: 'list',
        formulae: ['"قيد المراجعة,مقبول,مرفوض"'],
        allowBlank: true,
      }
    }
  })

  // Conditional formatting for app status (column H)
  const statusCol = colLetter(8)
  app.addConditionalFormatting({
    ref: `${statusCol}2:${statusCol}1048576`,
    rules: [
      { type: 'expression', formulae: [`${statusCol}2="مقبول"`], style: { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'E4F4E4' } }, font: { color: { argb: '1E7E1E' } } } },
      { type: 'expression', formulae: [`${statusCol}2="مرفوض"`], style: { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FBE4E4' } }, font: { color: { argb: 'B82424' } } } },
      { type: 'expression', formulae: [`${statusCol}2="قيد المراجعة"`], style: { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FEF7DA' } }, font: { color: { argb: 'AB7F12' } } } },
    ],
  })

  // ════════════════════════════════════════
  // Sheet 2: Statistics
  // ════════════════════════════════════════
  const stats = wb.addWorksheet('Statistics', {
    views: [{ state: 'frozen', ySplit: 1 }],
    properties: { tabColor: { argb: GOLD } },
  })

  stats.columns = [
    { header: 'الإحصائية', key: 'stat', width: 30 },
    { header: 'القيمة', key: 'value', width: 16 },
  ]
  styleHeader(stats, 2)

  const statRows = [
    { stat: 'إجمالي المتقدمين', value: `=COUNTA(Applications!A2:A1048576)` },
    { stat: 'عدد المقبولين', value: `=COUNTIF(Applications!H2:H1048576,"مقبول")` },
    { stat: 'عدد المرفوضين', value: `=COUNTIF(Applications!H2:H1048576,"مرفوض")` },
    { stat: 'قيد المراجعة', value: `=COUNTIF(Applications!H2:H1048576,"قيد المراجعة")` },
    {},
    { stat: 'الطلاب حسب المدرسة' },
    { stat: 'المدرسة', value: 'العدد' },
  ]
  statRows.forEach(r => stats.addRow(r))

  // School distribution QUERY
  stats.getRow(9).getCell(1).value = {
    formula: `QUERY(Applications!D2:D1048576,"select D,count(D) where D is not null group by D order by count(D) desc label D 'المدرسة', count(D) 'العدد'")`,
    date1904: false,
  }
  stats.getRow(9).getCell(1).numFmt = '@'

  styleDataRows(stats, 2, 2)

  // Highlight stat values
  stats.getCell('B2').font = { bold: true, size: 16, color: { argb: DEEP_BLUE } }
  stats.getCell('B3').font = { bold: true, size: 14, color: { argb: '1E7E1E' } }
  stats.getCell('B4').font = { bold: true, size: 14, color: { argb: 'B82424' } }
  stats.getCell('B5').font = { bold: true, size: 14, color: { argb: 'AB7F12' } }

  // ════════════════════════════════════════
  // Sheet 3: Search
  // ════════════════════════════════════════
  const search = wb.addWorksheet('Search', {
    views: [{ state: 'frozen', ySplit: 2 }],
    properties: { tabColor: { argb: '2D5A8E' } },
  })

  const searchCols = APP_COLS.slice(0, 9)
  search.columns = searchCols
  styleHeader(search, searchCols.length)

  // Row 2: search instruction + input
  search.mergeCells('A2:A2')
  const labelCell = search.getCell('A2')
  labelCell.value = '🔍 اكتب رقم الطلب / اسم الطالب / كود الطالب'
  labelCell.font = { bold: true, color: { argb: DEEP_BLUE }, size: 10 }
  labelCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'EDF2F9' } }
  labelCell.alignment = { horizontal: 'center', vertical: 'middle' }
  labelCell.border = borderStyle

  const inputCell = search.getCell('B2')
  inputCell.value = ''
  inputCell.font = { size: 12 }
  inputCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF9E0' } }
  inputCell.alignment = { horizontal: 'center', vertical: 'middle' }
  inputCell.border = borderStyle
  search.getRow(2).height = 30

  // Row 3: results formula
  const r3 = search.getRow(3)
  r3.getCell(1).value = {
    formula: `IFERROR(FILTER(Applications!A:I, (Applications!A:A=B2)+(Applications!B:B=B2)+(Applications!C:C=B2)), "ابحث بالرقم أو الاسم أو الكود في الخلية B2")`,
    date1904: false,
  }

  styleDataRows(search, searchCols.length, 3)

  // ════════════════════════════════════════
  // Sheet 4: Documents
  // ════════════════════════════════════════
  const docs = wb.addWorksheet('Documents', {
    views: [{ state: 'frozen', ySplit: 1 }],
    properties: { tabColor: { argb: '1E7E1E' } },
  })

  const docCols = [
    { header: 'رقم الطلب', key: 'rn', width: 16 },
    { header: 'اسم الطالب', key: 'name', width: 22 },
    { header: 'شهادة الميلاد', key: 'birth', width: 16 },
    { header: 'بطاقة الطالب', key: 'studCard', width: 16 },
    { header: 'بطاقة ولي الأمر', key: 'parentCard', width: 16 },
    { header: 'الصور الشخصية', key: 'photos', width: 14 },
    { header: 'ملف التقديم', key: 'appFile', width: 14 },
  ]
  docs.columns = docCols
  styleHeader(docs, docCols.length)

  docs.addRow({})
  styleDataRows(docs, docCols.length, 2)

  // Data validation for document columns (C-G)
  for (let col = 3; col <= 7; col++) {
    docs.getColumn(col).eachCell((cell, rowNum) => {
      if (rowNum > 1) {
        cell.dataValidation = {
          type: 'list',
          formulae: ['"تم التسليم,غير مسلم"'],
          allowBlank: true,
        }
      }
    })
  }

  // Conditional formatting for document statuses
  for (let col = 3; col <= 7; col++) {
    const cl = colLetter(col)
    docs.addConditionalFormatting({
      ref: `${cl}2:${cl}1048576`,
      rules: [
        { type: 'expression', formulae: [`${cl}2="تم التسليم"`], style: { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'E4F4E4' } }, font: { color: { argb: '1E7E1E' } } } },
        { type: 'expression', formulae: [`${cl}2="غير مسلم"`], style: { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FBE4E4' } }, font: { color: { argb: 'B82424' } } } },
      ],
    })
  }

  // ════════════════════════════════════════
  // Sheet 5: Dashboard
  // ════════════════════════════════════════
  const db = wb.addWorksheet('Dashboard', {
    properties: { tabColor: { argb: '8B5CF6' } },
  })

  // Summary cards
  db.columns = [
    { header: 'المؤشر', key: 'indicator', width: 24 },
    { header: 'القيمة', key: 'val', width: 14 },
    { header: '', key: 'blank1', width: 4 },
    { header: 'التاريخ', key: 'date', width: 14 },
    { header: 'عدد المتقدمين', key: 'dailyCount', width: 16 },
    { header: '', key: 'blank2', width: 4 },
    { header: 'المدرسة', key: 'school', width: 22 },
    { header: 'العدد', key: 'schoolCount', width: 12 },
  ]
  styleHeader(db, 8)

  const dbData = [
    { indicator: 'إجمالي المتقدمين', val: `=Statistics!B2` },
    { indicator: 'المقبولين', val: `=Statistics!B3` },
    { indicator: 'المرفوضين', val: `=Statistics!B4` },
    { indicator: 'قيد المراجعة', val: `=Statistics!B5` },
  ]
  dbData.forEach(r => db.addRow(r))

  // Daily applicants table
  db.getRow(2).getCell(4).value = 'التاريخ'
  db.getRow(2).getCell(4).font = headerFont
  db.getRow(2).getCell(4).fill = headerFill
  db.getRow(2).getCell(4).alignment = headerAlign
  db.getRow(2).getCell(5).value = 'عدد المتقدمين'
  db.getRow(2).getCell(5).font = headerFont
  db.getRow(2).getCell(5).fill = headerFill
  db.getRow(2).getCell(5).alignment = headerAlign

  db.getCell('D3').value = { formula: 'UNIQUE(Applications!I2:I1048576)', date1904: false }
  db.getCell('E3').value = { formula: 'IF(D3="","",COUNTIF(Applications!I2:I1048576,D3))', date1904: false }

  // School distribution table
  db.getRow(2).getCell(7).value = 'المدرسة'
  db.getRow(2).getCell(7).font = headerFont
  db.getRow(2).getCell(7).fill = headerFill
  db.getRow(2).getCell(7).alignment = headerAlign
  db.getRow(2).getCell(8).value = 'العدد'
  db.getRow(2).getCell(8).font = headerFont
  db.getRow(2).getCell(8).fill = headerFill
  db.getRow(2).getCell(8).alignment = headerAlign

  db.getCell('G3').value = { formula: "Statistics!A9:A", date1904: false }
  db.getCell('H3').value = { formula: "Statistics!B9:B", date1904: false }

  styleDataRows(db, 2, 2)

  // ──── APPLY SHEET TAB COLORS ────
  app.properties.tabColor = { argb: DEEP_BLUE }
  stats.properties.tabColor = { argb: GOLD }
  search.properties.tabColor = { argb: '2D5A8E' }
  docs.properties.tabColor = { argb: '1E7E1E' }
  db.properties.tabColor = { argb: '8B5CF6' }

  // ──── SET RTL (bi-directional support) ────
  // ExcelJS doesn't easily set RTL on worksheets, but we can note it
  // Content will render as RTL due to Arabic text

  // ──── SAVE ────
  const filePath = 'WE_Schools_Admission_Portal.xlsx'
  await wb.xlsx.writeFile(filePath)
  console.log(`\n  ✅ Created: ${filePath}`)
  console.log(`  📁 ${process.cwd()}\\${filePath}\n`)
}

function colLetter(n) {
  let s = ''
  while (n > 0) {
    n--
    s = String.fromCharCode(65 + (n % 26)) + s
    n = Math.floor(n / 26)
  }
  return s
}

main().catch(err => {
  console.error('Error:', err.message)
  process.exit(1)
})
