import ExcelJS from 'exceljs'

const C = {
  blue: '1E3A5F',
  gold: 'C8952E',
  green: '1E7E1E',
  red: 'B82424',
  yellow: 'AB7F12',
  orange: 'CC6600',
  gray: '6B7280',
  purple: '8B5CF6',
  teal: '0D9488',
  white: 'FFFFFF',
  bgEven: 'F4F6F9',
  bgOdd: 'FFFFFF',
  border: 'D0D5DD',
}

const APP_COLS = [
  { header: 'رقم الطلب',    key: 'rn',           width: 14 },
  { header: 'اسم الطالب',    key: 'name',         width: 22 },
  { header: 'كود الطالب',    key: 'code',         width: 16 },
  { header: 'المدرسة السابقة', key: 'prevSchool',  width: 24 },
  { header: 'الصف الدراسي السابق', key: 'prevGrade', width: 18 },
  { header: 'المحافظة',      key: 'gov',          width: 16 },
  { header: 'المركز / المدينة', key: 'city',       width: 18 },
  { header: 'رقم الطالب',    key: 'phone',        width: 16 },
  { header: 'رقم ولي الأمر',  key: 'parentPhone',  width: 16 },
  { header: 'الحالة الاجتماعية', key: 'social',    width: 16 },
  { header: 'حالة الطلب',    key: 'status',       width: 18 },
  { header: 'تاريخ التقديم',  key: 'submitted',    width: 18 },
  { header: 'وقت آخر تعديل',  key: 'modified',     width: 18 },
  { header: 'الاسم بالإنجليزية', key: 'nameEn',    width: 20 },
  { header: 'تاريخ الميلاد',  key: 'dob',          width: 14 },
  { header: 'الجنس',         key: 'gender',       width: 10 },
  { header: 'الجنسية',       key: 'nationality',  width: 14 },
  { header: 'الديانة',       key: 'religion',     width: 12 },
  { header: 'مهنة الأب',     key: 'fatherJob',    width: 20 },
  { header: 'مهنة الأم',     key: 'motherJob',    width: 20 },
  { header: 'العنوان',       key: 'address',      width: 28 },
  { header: 'البريد الإلكتروني', key: 'email',     width: 26 },
  { header: 'رقم بطاقة الطالب', key: 'studentId',  width: 18 },
  { header: 'اسم ولي الأمر',  key: 'parentName',   width: 20 },
  { header: 'رقم بطاقة ولي الأمر', key: 'parentId', width: 18 },
  { header: 'بريد ولي الأمر',  key: 'parentEmail', width: 26 },
]

const STATUSES = ['قيد المراجعة', 'مقبول', 'مرفوض', 'ناقص مستندات', 'بانتظار التواصل']
const STATUS_COLORS = {
  'مقبول':        { bg: 'E4F4E4', fg: C.green },
  'مرفوض':        { bg: 'FBE4E4', fg: C.red },
  'قيد المراجعة':  { bg: 'FEF7DA', fg: C.yellow },
  'ناقص مستندات':  { bg: 'FFF0E0', fg: C.orange },
  'بانتظار التواصل': { bg: 'E8ECF0', fg: C.gray },
}

const DOC_COLS = [
  { header: 'رقم الطلب',       key: 'rn',        width: 16 },
  { header: 'اسم الطالب',       key: 'name',      width: 22 },
  { header: 'شهادة الميلاد',    key: 'birth',     width: 18 },
  { header: 'بطاقة الطالب',     key: 'studCard',  width: 18 },
  { header: 'بطاقة ولي الأمر',  key: 'parentCard', width: 18 },
  { header: 'الصور الشخصية',    key: 'photos',    width: 18 },
  { header: 'ملف التقديم',      key: 'appFile',   width: 18 },
]

// ──── Helpers ────
function colLetter(n) {
  let s = ''
  while (n > 0) { n--; s = String.fromCharCode(65 + (n % 26)) + s; n = Math.floor(n / 26) }
  return s
}

function hdrFont() { return { name: 'Calibri', bold: true, color: { argb: C.white }, size: 11 } }
function hdrFill()  { return { type: 'pattern', pattern: 'solid', fgColor: { argb: C.blue } } }
function hdrAlign() { return { horizontal: 'center', vertical: 'middle' } }
function cellBorders() {
  return { top: { style: 'thin', color: { argb: C.border } }, left: { style: 'thin', color: { argb: C.border } }, bottom: { style: 'thin', color: { argb: C.border } }, right: { style: 'thin', color: { argb: C.border } } }
}

function styleHeader(ws, numCols) {
  const row = ws.getRow(1)
  row.height = 32
  for (let c = 1; c <= numCols; c++) {
    const cell = row.getCell(c)
    cell.font = hdrFont()
    cell.fill = hdrFill()
    cell.alignment = hdrAlign()
    cell.border = cellBorders()
  }
  ws.views = [{ state: 'frozen', ySplit: 1 }]
  ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: numCols } }
}

function styleDataRows(ws, numCols, startRow) {
  ws.eachRow((row, rowNum) => {
    if (rowNum < startRow) return
    const bg = (rowNum - startRow) % 2 === 0 ? C.bgEven : C.bgOdd
    row.eachCell((cell, colNum) => {
      if (colNum > numCols) return
      cell.border = cellBorders()
      cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true }
      if (!cell.fill || cell.fill.fgColor?.argb !== C.blue) {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } }
      }
    })
  })
}

async function main() {
  const wb = new ExcelJS.Workbook()
  wb.creator = 'WE Schools'
  wb.created = new Date()

  // ════════════════════════════════════════════════
  // Sheet 1: Applications
  // ════════════════════════════════════════════════
  const app = wb.addWorksheet('Applications', { properties: { tabColor: { argb: C.blue } } })
  app.columns = APP_COLS
  styleHeader(app, APP_COLS.length)

  // Empty data row
  const emptyRow = {}
  APP_COLS.forEach(c => { emptyRow[c.key] = '' })
  app.addRow(emptyRow)
  styleDataRows(app, APP_COLS.length, 2)

  // ── Data validation: social status (col 10 = J) ──
  const socialCol = app.getColumn(10)
  socialCol.eachCell((cell, rn) => { if (rn > 1) cell.dataValidation = { type: 'list', formulae: ['"أعزب,متزوج"'], allowBlank: true } })

  // ── Data validation: app status (col 11 = K) ──
  const statusColLetter = colLetter(11)
  const statusCol = app.getColumn(11)
  statusCol.eachCell((cell, rn) => { if (rn > 1) cell.dataValidation = { type: 'list', formulae: [`"${STATUSES.join(',')}"`], allowBlank: true } })

  // ── Conditional formatting for status column ──
  for (const [status, colors] of Object.entries(STATUS_COLORS)) {
    app.addConditionalFormatting({
      ref: `${statusColLetter}2:${statusColLetter}1048576`,
      rules: [{
        type: 'expression',
        formulae: [`${statusColLetter}2="${status}"`],
        style: { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: colors.bg } }, font: { color: { argb: colors.fg }, bold: true } },
      }],
    })
  }

  // ── AutoFilter on all columns ──
  app.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: APP_COLS.length } }

  // ════════════════════════════════════════════════
  // Sheet 2: Statistics
  // ════════════════════════════════════════════════
  const stats = wb.addWorksheet('Statistics', { properties: { tabColor: { argb: C.gold } } })
  stats.columns = [
    { header: 'الإحصائية', key: 'stat', width: 30 },
    { header: 'القيمة', key: 'value', width: 18 },
  ]
  styleHeader(stats, 2)
  stats.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: 2 } }

  const allRange = 'Applications!A2:A1048576'
  const statusRange = `${colLetter(11)}2:${colLetter(11)}1048576`

  const statRows = [
    { stat: 'إجمالي المتقدمين',        value: `=COUNTA(${allRange})` },
    { stat: 'عدد المقبولين',           value: `=COUNTIF(${statusRange},"مقبول")` },
    { stat: 'عدد المرفوضين',          value: `=COUNTIF(${statusRange},"مرفوض")` },
    { stat: 'قيد المراجعة',            value: `=COUNTIF(${statusRange},"قيد المراجعة")` },
    { stat: 'ناقص مستندات',           value: `=COUNTIF(${statusRange},"ناقص مستندات")` },
    { stat: 'بانتظار التواصل',         value: `=COUNTIF(${statusRange},"بانتظار التواصل")` },
    {},
    { stat: 'الطلاب حسب المدرسة' },
    { stat: 'المدرسة', value: 'العدد' },
  ]
  statRows.forEach(r => stats.addRow(r))
  stats.getRow(10).getCell(1).value = { formula: `QUERY(Applications!D2:D1048576,"select D,count(D) where D is not null group by D order by count(D) desc label D 'المدرسة', count(D) 'العدد'")`, date1904: false }

  styleDataRows(stats, 2, 2)
  stats.getCell('B2').font = { bold: true, size: 16, color: { argb: C.blue } }
  for (const [status, colors] of Object.entries(STATUS_COLORS)) {
    const cell = stats.getCell(`B${2 + STATUSES.indexOf(status)}`)
    if (cell) { cell.font = { bold: true, size: 14, color: { argb: colors.fg } } }
  }

  // ── Governorate distribution ──
  stats.getCell('D7').value = 'الطلاب حسب المحافظة'
  stats.getCell('D7').font = { bold: true, color: { argb: C.blue }, size: 11 }
  stats.getCell('D8').value = 'المحافظة'
  stats.getCell('D8').font = hdrFont(); stats.getCell('D8').fill = hdrFill(); stats.getCell('D8').alignment = hdrAlign()
  stats.getCell('E8').value = 'العدد'
  stats.getCell('E8').font = hdrFont(); stats.getCell('E8').fill = hdrFill(); stats.getCell('E8').alignment = hdrAlign()
  stats.getCell('D9').value = { formula: `QUERY(Applications!F2:F1048576,"select F,count(F) where F is not null group by F order by count(F) desc label F 'المحافظة', count(F) 'العدد'")`, date1904: false }

  // ════════════════════════════════════════════════
  // Sheet 3: Search
  // ════════════════════════════════════════════════
  const search = wb.addWorksheet('Search', { properties: { tabColor: { argb: '2D5A8E' } } })
  search.columns = APP_COLS // all 26 columns
  styleHeader(search, APP_COLS.length)

  // Row 2: search instruction + input
  search.mergeCells('A2:A2')
  const labelCell = search.getCell('A2')
  labelCell.value = '🔍 اكتب رقم الطلب / اسم الطالب / كود الطالب'
  labelCell.font = { bold: true, color: { argb: C.blue }, size: 10 }
  labelCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'EDF2F9' } }
  labelCell.alignment = { horizontal: 'center', vertical: 'middle' }
  labelCell.border = cellBorders()

  const inputCell = search.getCell('B2')
  inputCell.value = ''
  inputCell.font = { size: 12 }
  inputCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF9E0' } }
  inputCell.alignment = { horizontal: 'center', vertical: 'middle' }
  inputCell.border = cellBorders()
  search.getRow(2).height = 30

  // Merge remaining cells in row 2 for visual consistency
  for (let c = 2; c <= APP_COLS.length; c++) {
    const cl = colLetter(c)
    if (cl === 'B') continue
    search.mergeCells(`${cl}2:${cl}2`)
    const cell = search.getCell(`${cl}2`)
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'EDF2F9' } }
    cell.border = cellBorders()
  }

  // Row 3: FILTER formula — search ALL columns (A through last column letter)
  const lastCol = colLetter(APP_COLS.length)
  const searchAllCols = []
  for (let i = 1; i <= APP_COLS.length; i++) {
    searchAllCols.push(`Applications!${colLetter(i)}:${colLetter(i)}`)
  }
  const searchFormula = `IFERROR(FILTER(Applications!A:${lastCol}, (Applications!A:A=B2)+(Applications!B:B=B2)+(Applications!C:C=B2)), "ابحث بالرقم أو الاسم أو الكود في الخلية B2")`

  const r3 = search.getRow(3)
  r3.getCell(1).value = { formula: searchFormula, date1904: false }

  search.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: APP_COLS.length } }

  // ════════════════════════════════════════════════
  // Sheet 4: Documents
  // ════════════════════════════════════════════════
  const docs = wb.addWorksheet('Documents', { properties: { tabColor: { argb: C.green } } })
  docs.columns = DOC_COLS
  styleHeader(docs, DOC_COLS.length)

  docs.addRow({})
  styleDataRows(docs, DOC_COLS.length, 2)
  docs.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: DOC_COLS.length } }

  // Checkbox simulation with ✓ / ✗ + conditional formatting
  for (let col = 3; col <= DOC_COLS.length; col++) {
    const cl = colLetter(col)
    docs.getColumn(col).eachCell((cell, rn) => {
      if (rn > 1) {
        cell.dataValidation = { type: 'list', formulae: ['"✓,✗"'], allowBlank: true }
        cell.value = '✗'
      }
    })

    docs.addConditionalFormatting({
      ref: `${cl}2:${cl}1048576`,
      rules: [
        { type: 'expression', formulae: [`${cl}2="✓"`], style: { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'E4F4E4' } }, font: { color: { argb: C.green }, bold: true, size: 14 } } },
        { type: 'expression', formulae: [`${cl}2="✗"`], style: { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FBE4E4' } }, font: { color: { argb: C.red }, bold: true, size: 14 } } },
      ],
    })
  }

  // ════════════════════════════════════════════════
  // Sheet 5: Dashboard
  // ════════════════════════════════════════════════
  const db = wb.addWorksheet('Dashboard', { properties: { tabColor: { argb: C.purple } } })
  db.columns = [
    { header: 'المؤشر', key: 'ind', width: 24 },
    { header: 'القيمة', key: 'val', width: 14 },
    { header: '', key: 's1', width: 3 },
    { header: 'التاريخ', key: 'dt', width: 14 },
    { header: 'عدد المتقدمين', key: 'daily', width: 18 },
    { header: '', key: 's2', width: 3 },
    { header: 'المدرسة', key: 'school', width: 22 },
    { header: 'العدد', key: 'sc', width: 12 },
    { header: '', key: 's3', width: 3 },
    { header: 'المحافظة', key: 'gov', width: 16 },
    { header: 'العدد', key: 'gc', width: 12 },
    { header: '', key: 's4', width: 3 },
    { header: 'الأسبوع', key: 'wk', width: 14 },
    { header: 'المتقدمين', key: 'wc', width: 14 },
  ]
  styleHeader(db, 14)

  // ── Summary Cards (A1:B7) ──
  const cardData = [
    { ind: '📋 إجمالي الطلبات',       val: '=Statistics!B2' },
    { ind: '🟢 المقبولين',            val: '=Statistics!B3' },
    { ind: '🔴 المرفوضين',            val: '=Statistics!B4' },
    { ind: '🟡 تحت المراجعة',         val: '=Statistics!B5' },
    { ind: '🟠 ناقص مستندات',         val: '=Statistics!B6' },
    { ind: 'بانتظار التواصل',         val: '=Statistics!B7' },
  ]
  cardData.forEach(r => db.addRow(r))

  // Card colors
  const cardColors = [C.blue, C.green, C.red, C.yellow, C.orange, C.gray]
  cardData.forEach((r, i) => {
    const cell = db.getCell(`A${i + 2}`)
    cell.font = { bold: true, color: { argb: cardColors[i] }, size: 12 }
    const valCell = db.getCell(`B${i + 2}`)
    valCell.font = { bold: true, size: 18, color: { argb: cardColors[i] } }
  })

  // ── KPI Cards (D1:E…) ──
  db.getRow(2).getCell(4).value = '📊 مؤشرات الأداء'
  db.getRow(2).getCell(4).font = { bold: true, size: 12, color: { argb: C.blue } }

  const kpis = [
    { label: 'آخر تسجيل',          formula: `=MAX(Applications!${colLetter(12)}2:${colLetter(12)}1048576)` },
    { label: 'أول تسجيل',          formula: `=MIN(Applications!${colLetter(12)}2:${colLetter(12)}1048576)` },
    { label: 'متوسط التسجيل يوميًا', formula: `=IFERROR(ROUND(Statistics!B2/MAX(DAYS(MAX(Applications!${colLetter(12)}2:${colLetter(12)}1048576),MIN(Applications!${colLetter(12)}2:${colLetter(12)}1048576)),1),1),0)` },
    { label: 'نسبة اكتمال البيانات', formula: `=IFERROR(ROUND(COUNTIF(Applications!A2:A1048576,"<>")/COUNTA(Applications!A2:A1048576)*100,1)&"%","0%")` },
  ]
  kpis.forEach((k, i) => {
    db.getCell(`D${3 + i}`).value = k.label
    db.getCell(`D${3 + i}`).font = { bold: true, color: { argb: C.blue }, size: 10 }
    db.getCell(`E${3 + i}`).value = { formula: k.formula, date1904: false }
    db.getCell(`E${3 + i}`).font = { bold: true, size: 12, color: { argb: C.teal } }
  })

  // ── Daily registrations (G1:H…) ──
  db.getRow(2).getCell(7).value = 'التاريخ'
  db.getRow(2).getCell(7).font = hdrFont(); db.getRow(2).getCell(7).fill = hdrFill(); db.getRow(2).getCell(7).alignment = hdrAlign()
  db.getRow(2).getCell(8).value = 'عدد المتقدمين'
  db.getRow(2).getCell(8).font = hdrFont(); db.getRow(2).getCell(8).fill = hdrFill(); db.getRow(2).getCell(8).alignment = hdrAlign()
  const dateCol = colLetter(12)
  db.getCell('G3').value = { formula: `UNIQUE(Applications!${dateCol}2:${dateCol}1048576)`, date1904: false }
  db.getCell('H3').value = { formula: `IF(G3="","",COUNTIF(Applications!${dateCol}2:${dateCol}1048576,G3))`, date1904: false }

  // ── School distribution (J1:K…) ──
  db.getRow(2).getCell(10).value = 'المدرسة'
  db.getRow(2).getCell(10).font = hdrFont(); db.getRow(2).getCell(10).fill = hdrFill(); db.getRow(2).getCell(10).alignment = hdrAlign()
  db.getRow(2).getCell(11).value = 'العدد'
  db.getRow(2).getCell(11).font = hdrFont(); db.getRow(2).getCell(11).fill = hdrFill(); db.getRow(2).getCell(11).alignment = hdrAlign()
  db.getCell('J3').value = { formula: "Statistics!A9:A", date1904: false }
  db.getCell('K3').value = { formula: "Statistics!B9:B", date1904: false }

  // ── Governorate distribution (M1:N…) ──
  db.getRow(2).getCell(13).value = 'المحافظة'
  db.getRow(2).getCell(13).font = hdrFont(); db.getRow(2).getCell(13).fill = hdrFill(); db.getRow(2).getCell(13).alignment = hdrAlign()
  db.getRow(2).getCell(14).value = 'العدد'
  db.getRow(2).getCell(14).font = hdrFont(); db.getRow(2).getCell(14).fill = hdrFill(); db.getRow(2).getCell(14).alignment = hdrAlign()
  db.getCell('M3').value = { formula: "Statistics!D9:D", date1904: false }
  db.getCell('N3').value = { formula: "Statistics!E9:E", date1904: false }

  // ── Weekly registrations (P1:Q…) ──
  db.getRow(2).getCell(16).value = 'الأسبوع'
  db.getRow(2).getCell(16).font = hdrFont(); db.getRow(2).getCell(16).fill = hdrFill(); db.getRow(2).getCell(16).alignment = hdrAlign()
  db.getRow(2).getCell(17).value = 'المتقدمين'
  db.getRow(2).getCell(17).font = hdrFont(); db.getRow(2).getCell(17).fill = hdrFill(); db.getRow(2).getCell(17).alignment = hdrAlign()
  db.getCell('P3').value = { formula: `SORT(UNIQUE(TEXT(Applications!${dateCol}2:${dateCol}1048576,"YYYY-ww")))`, date1904: false }
  db.getCell('Q3').value = { formula: `IF(P3="","",COUNTIF(TEXT(Applications!${dateCol}2:${dateCol}1048576,"YYYY-ww"),P3))`, date1904: false }

  // ── Freeze + Filter on Dashboard ──
  db.views = [{ state: 'frozen', ySplit: 1 }]

  // ── SAVE ──
  const filePath = 'WE_Schools_v2.xlsx'
  await wb.xlsx.writeFile(filePath)
  console.log(`\n  ✅ Created: ${filePath} (${APP_COLS.length} columns, 5 sheets with full formatting)\n`)
}

main().catch(err => { console.error('Error:', err.message); process.exit(1) })
