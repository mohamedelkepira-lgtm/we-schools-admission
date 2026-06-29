import ExcelJS from 'exceljs'

const C = {
  blue: '1E3A5F', gold: 'C8952E', green: '1E7E1E', red: 'B82424',
  yellow: 'AB7F12', orange: 'CC6600', gray: '6B7280', purple: '8B5CF6',
  teal: '0D9488', white: 'FFFFFF', bgEven: 'F4F6F9', bgOdd: 'FFFFFF', border: 'D0D5DD',
}

const APP_COLS = [
  { header: 'رقم الطلب',            key: 'rn',           width: 14 },
  { header: 'اسم الطالب',            key: 'name',         width: 22 },
  { header: 'كود الطالب',            key: 'code',         width: 16 },
  { header: 'المدرسة السابقة',        key: 'prevSchool',   width: 24 },
  { header: 'الصف الدراسي السابق',     key: 'prevGrade',    width: 18 },
  { header: 'المحافظة',              key: 'gov',          width: 16 },
  { header: 'المركز / المدينة',        key: 'city',         width: 18 },
  { header: 'رقم الطالب',            key: 'phone',        width: 16 },
  { header: 'رقم ولي الأمر',          key: 'parentPhone',  width: 16 },
  { header: 'الحالة الاجتماعية',       key: 'social',       width: 16 },
  { header: 'حالة الطلب',            key: 'status',       width: 18 },
  { header: 'تاريخ التقديم',          key: 'submitted',    width: 18 },
  { header: 'وقت آخر تعديل',          key: 'modified',     width: 18 },
  { header: 'الاسم بالإنجليزية',       key: 'nameEn',       width: 20 },
  { header: 'تاريخ الميلاد',          key: 'dob',          width: 14 },
  { header: 'الجنس',                key: 'gender',       width: 10 },
  { header: 'الجنسية',              key: 'nationality',  width: 14 },
  { header: 'الديانة',              key: 'religion',     width: 12 },
  { header: 'مهنة الأب',            key: 'fatherJob',    width: 20 },
  { header: 'مهنة الأم',            key: 'motherJob',    width: 20 },
  { header: 'العنوان',              key: 'address',      width: 28 },
  { header: 'البريد الإلكتروني',      key: 'email',        width: 26 },
  { header: 'رقم بطاقة الطالب',       key: 'studentId',    width: 18 },
  { header: 'اسم ولي الأمر',         key: 'parentName',   width: 20 },
  { header: 'رقم بطاقة ولي الأمر',    key: 'parentId',     width: 18 },
  { header: 'بريد ولي الأمر',        key: 'parentEmail',  width: 26 },
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

const FH = { name: 'Calibri', bold: true, color: { argb: C.white }, size: 11 }
const HB = () => ({ type: 'pattern', pattern: 'solid', fgColor: { argb: C.blue } })
const HA = () => ({ horizontal: 'center', vertical: 'middle' })
const BD = () => ({
  top: { style: 'thin', color: { argb: C.border } },
  left: { style: 'thin', color: { argb: C.border } },
  bottom: { style: 'thin', color: { argb: C.border } },
  right: { style: 'thin', color: { argb: C.border } },
})

function styleHdr(ws, n) {
  const r = ws.getRow(1)
  r.height = 32
  for (let c = 1; c <= n; c++) { const v = r.getCell(c); v.font = FH; v.fill = HB(); v.alignment = HA(); v.border = BD() }
  ws.views = [{ state: 'frozen', ySplit: 1 }]
  ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: n } }
}

function styleRows(ws, n, start) {
  ws.eachRow((row, rn) => {
    if (rn < start) return
    const bg = (rn - start) % 2 === 0 ? C.bgEven : C.bgOdd
    row.eachCell((c, cn) => {
      if (cn > n) return
      c.border = BD()
      c.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true }
      if (!c.fill || c.fill.fgColor?.argb !== C.blue) c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } }
    })
  })
}

function formula(v) { return { formula: v, date1904: false } }

async function main() {
  const wb = new ExcelJS.Workbook()
  wb.creator = 'WE Schools'
  wb.created = new Date()

  // ════════════════════════════════ Sheet 1: Applications ════════════════════
  const app = wb.addWorksheet('Applications', { properties: { tabColor: { argb: C.blue } } })
  app.columns = APP_COLS; styleHdr(app, APP_COLS.length)
  const e = {}; APP_COLS.forEach(c => { e[c.key] = '' }); app.addRow(e)
  styleRows(app, APP_COLS.length, 2)

  // Social status dropdown (col J)
  app.getColumn(10).eachCell((c, rn) => { if (rn > 1) c.dataValidation = { type: 'list', formulae: ['"أعزب,متزوج"'], allowBlank: true } })
  // App status dropdown (col K) — 5 values
  const kL = colLetter(11)
  app.getColumn(11).eachCell((c, rn) => { if (rn > 1) c.dataValidation = { type: 'list', formulae: [`"${STATUSES.join(',')}"`], allowBlank: true } })
  // Conditional formatting — status colors
  for (const [s, cl] of Object.entries(STATUS_COLORS))
    app.addConditionalFormatting({ ref: `${kL}2:${kL}1048576`, rules: [{ type: 'expression', formulae: [`${kL}2="${s}"`], style: { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: cl.bg } }, font: { color: { argb: cl.fg }, bold: true } } }] })

  // ════════════════════════════════ Sheet 2: Statistics ════════════════════════
  const st = wb.addWorksheet('Statistics', { properties: { tabColor: { argb: C.gold } } })
  st.columns = [{ header: 'الإحصائية', key: 'a', width: 30 }, { header: 'القيمة', key: 'b', width: 18 }]
  styleHdr(st, 2)

  const AR = 'Applications!A2:A1048576'       // request numbers
  const SR = `${kL}2:${kL}1048576`            // statuses
  const [SR2, SR3, SR4, SR5, SR6] = STATUSES.map(s => `"${s}"`)

  ;[
    { a: 'إجمالي المتقدمين',        b: `=COUNTA(${AR})` },
    { a: 'عدد المقبولين',           b: `=COUNTIF(${SR},${SR2})` },
    { a: 'عدد المرفوضين',          b: `=COUNTIF(${SR},${SR3})` },
    { a: 'قيد المراجعة',            b: `=COUNTIF(${SR},${SR4})` },
    { a: 'ناقص مستندات',           b: `=COUNTIF(${SR},${SR5})` },
    { a: 'بانتظار التواصل',         b: `=COUNTIF(${SR},${SR6})` },
    {},
    { a: 'الطلاب حسب المدرسة' },
    { a: 'المدرسة', b: 'العدد' },
  ].forEach(r => st.addRow(r))

  // Sorted unique schools (A11↓) — compatible array formula
  st.getCell('A11').value = formula(
    `IFERROR(INDEX(Applications!$D$2:$D$9999, MATCH(0, COUNTIF($A$10:$A10, Applications!$D$2:$D$9999), 0)), "")`
  )
  st.getCell('B11').value = formula(
    `IF(A11="","",COUNTIF(Applications!$D$2:$D$9999, A11))`
  )

  // Governorate distribution (D8:E8↓)
  st.getCell('D8').value = 'الطلاب حسب المحافظة'
  st.getCell('D8').font = FH; st.getCell('D8').fill = HB(); st.getCell('D8').alignment = HA()
  st.getCell('E8').value = 'العدد'
  st.getCell('E8').font = FH; st.getCell('E8').fill = HB(); st.getCell('E8').alignment = HA()
  st.getCell('D9').value = formula(
    `IFERROR(INDEX(Applications!$F$2:$F$9999, MATCH(0, COUNTIF($D$8:$D8, Applications!$F$2:$F$9999), 0)), "")`
  )
  st.getCell('E9').value = formula(
    `IF(D9="","",COUNTIF(Applications!$F$2:$F$9999, D9))`
  )

  styleRows(st, 2, 2)
  // Big number styling
  st.getCell('B2').font = { bold: true, size: 16, color: { argb: C.blue } }
  const statusIdx = { 'مقبول': 3, 'مرفوض': 4, 'قيد المراجعة': 5, 'ناقص مستندات': 6, 'بانتظار التواصل': 7 }
  for (const [s, idx] of Object.entries(statusIdx))
    st.getCell(`B${idx}`).font = { bold: true, size: 14, color: { argb: STATUS_COLORS[s].fg } }

  // ════════════════════════════════ Sheet 3: Search ════════════════════════════
  const sr = wb.addWorksheet('Search', { properties: { tabColor: { argb: '2D5A8E' } } })
  sr.columns = APP_COLS; styleHdr(sr, APP_COLS.length)

  sr.mergeCells('A2:A2')
  sr.getCell('A2').value = '🔍 اكتب رقم الطلب'
  sr.getCell('A2').font = { bold: true, color: { argb: C.blue }, size: 10 }
  sr.getCell('A2').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'EDF2F9' } }
  sr.getCell('A2').alignment = HA(); sr.getCell('A2').border = BD()

  sr.getCell('B2').value = ''
  sr.getCell('B2').font = { size: 12 }
  sr.getCell('B2').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF9E0' } }
  sr.getCell('B2').alignment = HA(); sr.getCell('B2').border = BD()
  sr.getRow(2).height = 30

  // VLOOKUP by request number — works in ALL Excel versions
  sr.getCell('A3').value = formula(
    `IFERROR(VLOOKUP($B$2, Applications!$A$2:$Z$9999, COLUMN(A1), FALSE), "")`
  )
  // Fill across all columns
  const srRow3 = sr.getRow(3)
  for (let c = 2; c <= APP_COLS.length; c++) {
    srRow3.getCell(c).value = formula(
      `IFERROR(VLOOKUP($B$2, Applications!$A$2:$Z$9999, ${c}, FALSE), "")`
    )
  }

  // ════════════════════════════════ Sheet 4: Documents ════════════════════════
  const dc = wb.addWorksheet('Documents', { properties: { tabColor: { argb: C.green } } })
  dc.columns = DOC_COLS; styleHdr(dc, DOC_COLS.length)
  dc.addRow({}); styleRows(dc, DOC_COLS.length, 2)
  dc.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: DOC_COLS.length } }

  for (let c = 3; c <= DOC_COLS.length; c++) {
    const cl = colLetter(c)
    dc.getColumn(c).eachCell((cell, rn) => {
      if (rn > 1) { cell.dataValidation = { type: 'list', formulae: ['"✓,✗"'], allowBlank: true }; cell.value = '✗' }
    })
    dc.addConditionalFormatting({ ref: `${cl}2:${cl}1048576`, rules: [
      { type: 'expression', formulae: [`${cl}2="✓"`], style: { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'E4F4E4' } }, font: { color: { argb: C.green }, bold: true, size: 14 } } },
      { type: 'expression', formulae: [`${cl}2="✗"`], style: { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FBE4E4' } }, font: { color: { argb: C.red }, bold: true, size: 14 } } },
    ]})
  }

  // ════════════════════════════════ Sheet 5: Dashboard ════════════════════════
  const db = wb.addWorksheet('Dashboard', { properties: { tabColor: { argb: C.purple } } })
  db.columns = [
    { header: 'المؤشر', key: 'i', width: 24 }, { header: 'القيمة', key: 'v', width: 14 },
    { header: '', key: 's1', width: 3 },
    { header: 'المؤشر', key: 'i2', width: 22 }, { header: 'القيمة', key: 'v2', width: 18 },
    { header: '', key: 's2', width: 3 },
    { header: 'التاريخ', key: 'd', width: 14 }, { header: 'عدد المتقدمين', key: 'dc', width: 18 },
    { header: '', key: 's3', width: 3 },
    { header: 'المدرسة', key: 'sc', width: 22 }, { header: 'العدد', key: 'scn', width: 12 },
    { header: '', key: 's4', width: 3 },
    { header: 'المحافظة', key: 'gv', width: 16 }, { header: 'العدد', key: 'gvn', width: 12 },
  ]
  styleHdr(db, 14)

  // Row 2: summary cards + KPI + chart data headers
  const cards = [
    '📋 إجمالي الطلبات', '🟢 المقبولين', '🔴 المرفوضين',
    '🟡 تحت المراجعة', '🟠 ناقص مستندات', 'بانتظار التواصل',
  ]
  cards.forEach((c, i) => {
    db.getCell(`A${i + 2}`).value = c
    db.getCell(`A${i + 2}`).font = { bold: true, color: { argb: [C.blue, C.green, C.red, C.yellow, C.orange, C.gray][i] }, size: 12 }
    db.getCell(`B${i + 2}`).value = formula(`=Statistics!B${i + 2}`)
    db.getCell(`B${i + 2}`).font = { bold: true, size: 18, color: { argb: [C.blue, C.green, C.red, C.yellow, C.orange, C.gray][i] } }
  })

  // KPI section (D2:E8)
  db.getCell('D2').value = '📊 مؤشرات الأداء'
  db.getCell('D2').font = { bold: true, size: 12, color: { argb: C.blue } }

  const kpis = [
    { label: 'آخر تسجيل',          f: `=MAX(Applications!${colLetter(12)}2:${colLetter(12)}1048576)` },
    { label: 'أول تسجيل',          f: `=MIN(Applications!${colLetter(12)}2:${colLetter(12)}1048576)` },
    { label: 'متوسط التسجيل يوميًا', f: `=IFERROR(ROUND(Statistics!B2/MAX(DAYS(MAX(Applications!${colLetter(12)}2:${colLetter(12)}1048576),MIN(Applications!${colLetter(12)}2:${colLetter(12)}1048576)),1),1),0)` },
    { label: 'نسبة اكتمال البيانات', f: `=IFERROR(ROUND(COUNTIF(Applications!A2:A1048576,"<>")/COUNTA(Applications!A2:A1048576)*100,1)&"%","0%")` },
  ]
  kpis.forEach((k, i) => {
    db.getCell(`D${3 + i}`).value = k.label
    db.getCell(`D${3 + i}`).font = { bold: true, color: { argb: C.blue }, size: 10 }
    db.getCell(`E${3 + i}`).value = formula(k.f)
    db.getCell(`E${3 + i}`).font = { bold: true, size: 12, color: { argb: C.teal } }
  })

  // E8 divider
  db.getCell('D8').value = ''; db.getCell('E8').value = ''

  // Chart data tables — compatible formulas only
  // School distribution (G2:H↓) — references Statistics
  db.getCell('G2').value = 'المدرسة'; db.getCell('G2').font = FH; db.getCell('G2').fill = HB(); db.getCell('G2').alignment = HA()
  db.getCell('H2').value = 'العدد';   db.getCell('H2').font = FH; db.getCell('H2').fill = HB(); db.getCell('H2').alignment = HA()
  db.getCell('G3').value = formula(`=Statistics!A11:A`)
  db.getCell('H3').value = formula(`=Statistics!B11:B`)

  // Governorate distribution (J2:K↓) — references Statistics
  db.getCell('J2').value = 'المحافظة'; db.getCell('J2').font = FH; db.getCell('J2').fill = HB(); db.getCell('J2').alignment = HA()
  db.getCell('K2').value = 'العدد';    db.getCell('K2').font = FH; db.getCell('K2').fill = HB(); db.getCell('K2').alignment = HA()
  db.getCell('J3').value = formula(`=Statistics!D9:D`)
  db.getCell('K3').value = formula(`=Statistics!E9:E`)

  // Daily data (M2:N↓) — manual note since UNIQUE not available in Excel 2019
  db.getCell('M2').value = '📈 بيانات الرسم البياني'
  db.getCell('M2').font = { bold: true, size: 10, color: { argb: C.blue } }
  db.getCell('M3').value = 'استخدم PivotTable للرسوم البيانية'
  db.getCell('M3').font = { color: { argb: C.gray }, italic: true, size: 9 }

  db.views = [{ state: 'frozen', ySplit: 1 }]

  // ════════════════════════════════ Save ════════════════════════════════
  const f = 'WE_Schools_v2.xlsx'
  await wb.xlsx.writeFile(f)
  console.log(`\n  ✅ Created: ${f} (${APP_COLS.length} cols · 5 sheets · Excel 2019+ compatible)\n`)
}

main().catch(e => { console.error('Error:', e.message); process.exit(1) })
