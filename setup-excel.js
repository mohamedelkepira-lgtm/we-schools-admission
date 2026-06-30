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
const STATUS_CLR = {
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
function cl(n) { let s = ''; while (n > 0) { n--; s = String.fromCharCode(65 + (n % 26)) + s; n = Math.floor(n / 26) } return s }

const fH   = { name: 'Calibri', bold: true, color: { argb: C.white }, size: 11 }
const hFill = () => ({ type: 'pattern', pattern: 'solid', fgColor: { argb: C.blue } })
const hA    = () => ({ horizontal: 'center', vertical: 'middle' })
const bd    = () => ({ top: { style: 'thin', color: { argb: C.border } }, left: { style: 'thin', color: { argb: C.border } }, bottom: { style: 'thin', color: { argb: C.border } }, right: { style: 'thin', color: { argb: C.border } } })

function hdr(ws, n) {
  const r = ws.getRow(1); r.height = 32
  for (let c = 1; c <= n; c++) { const v = r.getCell(c); v.font = fH; v.fill = hFill(); v.alignment = hA(); v.border = bd() }
  ws.views = [{ state: 'frozen', ySplit: 1 }]
  ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: n } }
}

function rows(ws, n, start) {
  ws.eachRow((row, rn) => {
    if (rn < start) return
    const bg = (rn - start) % 2 === 0 ? C.bgEven : C.bgOdd
    row.eachCell((c, cn) => {
      if (cn > n) return
      c.border = bd(); c.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true }
      if (!c.fill || c.fill.fgColor?.argb !== C.blue) c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } }
    })
  })
}

function fm(v) { return { formula: v, date1904: false } }

// ──── Unique-extraction formula (Excel 2019 compatible, no CSE required) ────
// Uses INDEX(COUNTIF(...),0) to force array evaluation without Ctrl+Shift+Enter
function uniqF(dataCol, headerRow) {
  return `IFERROR(INDEX(Applications!$${dataCol}$${2}:$${dataCol}$${9999}, MATCH(0, INDEX(COUNTIF($${cl(1)}$${headerRow}:$${cl(1)}${headerRow}, Applications!$${dataCol}$${2}:$${dataCol}$${9999}), 0), 0)), "")`
}
function countF(dataCol) {
  return `IF(A11="","",COUNTIF(Applications!$${dataCol}$${2}:$${dataCol}$${9999}, A11))`
}

async function main() {
  const wb = new ExcelJS.Workbook()
  wb.creator = 'WE Schools'; wb.created = new Date()

  // ════════════════════════ 1 · Applications ════════════════════════
  const app = wb.addWorksheet('Applications', { properties: { tabColor: { argb: C.blue } } })
  app.columns = APP_COLS; hdr(app, APP_COLS.length)
  const e = {}; APP_COLS.forEach(c => { e[c.key] = '' }); app.addRow(e)
  rows(app, APP_COLS.length, 2)

  const KL = cl(11) // status column letter
  app.getColumn(10).eachCell((c, rn) => { if (rn > 1) c.dataValidation = { type: 'list', formulae: ['"أعزب,متزوج"'], allowBlank: true } })
  app.getColumn(11).eachCell((c, rn) => { if (rn > 1) c.dataValidation = { type: 'list', formulae: [`"${STATUSES.join(',')}"`], allowBlank: true } })
  for (const [s, clr] of Object.entries(STATUS_CLR))
    app.addConditionalFormatting({ ref: `${KL}2:${KL}1048576`, rules: [{ type: 'expression', formulae: [`${KL}2="${s}"`], style: { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: clr.bg } }, font: { color: { argb: clr.fg }, bold: true } } }] })

  // ════════════════════════ 2 · Statistics ════════════════════════
  const st = wb.addWorksheet('Statistics', { properties: { tabColor: { argb: C.gold } } })
  st.columns = [
    { header: 'الإحصائية', key: 'a', width: 30 },
    { header: 'القيمة', key: 'b', width: 16 },
  ]
  hdr(st, 2)
  st.autoFilter = null  // only 1 header row is ok

  const AR = 'Applications!A2:A1048576'
  const SR = `${KL}2:${KL}1048576`

  // ── Main counts ──
  const counts = [
    ['📋 إجمالي المتقدمين',      `=COUNTA(${AR})`],
    ['🟢 عدد المقبولين',          `=COUNTIF(${SR},"مقبول")`],
    ['🔴 عدد المرفوضين',          `=COUNTIF(${SR},"مرفوض")`],
    ['🟡 قيد المراجعة',           `=COUNTIF(${SR},"قيد المراجعة")`],
    ['🟠 ناقص مستندات',          `=COUNTIF(${SR},"ناقص مستندات")`],
    ['⏳ بانتظار التواصل',        `=COUNTIF(${SR},"بانتظار التواصل")`],
  ]
  let r = 2
  for (const [label, f] of counts) {
    st.getCell(`A${r}`).value = label
    st.getCell(`A${r}`).font = { bold: true, size: 12, color: { argb: C.blue } }
    st.getCell(`B${r}`).value = fm(f)
    st.getCell(`B${r}`).font = { bold: true, size: 16, color: { argb: C.blue } }
    r++
  }

  // Color the status values
  const rowMap = { 'مقبول': 3, 'مرفوض': 4, 'قيد المراجعة': 5, 'ناقص مستندات': 6, 'بانتظار التواصل': 7 }
  for (const [s, ri] of Object.entries(rowMap))
    st.getCell(`B${ri}`).font = { bold: true, size: 16, color: { argb: STATUS_CLR[s].fg } }

  // ── School distribution ──
  r += 1  // blank row
  st.getCell(`A${r}`).value = '🏫 توزيع المدارس'
  st.getCell(`A${r}`).font = { bold: true, size: 12, color: { argb: C.blue } }
  r++
  st.getCell(`A${r}`).value = 'المدرسة'
  st.getCell(`A${r}`).font = fH; st.getCell(`A${r}`).fill = hFill(); st.getCell(`A${r}`).alignment = hA()
  st.getCell(`B${r}`).value = 'العدد'
  st.getCell(`B${r}`).font = fH; st.getCell(`B${r}`).fill = hFill(); st.getCell(`B${r}`).alignment = hA()
  const schoolStart = r
  r++
  // cell A{r}: unique schools — INDEX(…, INDEX(COUNTIF(…),0), …) avoids CSE
  st.getCell(`A${r}`).value = fm(
    `IFERROR(INDEX(Applications!$${cl(4)}$2:$${cl(4)}$9999, MATCH(0, INDEX(COUNTIF($${cl(1)}$${schoolStart}:$${cl(1)}${r-1}, Applications!$${cl(4)}$2:$${cl(4)}$9999), 0), 0)), "")`
  )
  st.getCell(`B${r}`).value = fm(`IF(A${r}="","",COUNTIF(Applications!$${cl(4)}$2:$${cl(4)}$9999, A${r}))`)
  const schoolDataRow = r  // first data row

  // ── Governorate distribution ──
  r += 2
  st.getCell(`A${r}`).value = '🌍 توزيع المحافظات'
  st.getCell(`A${r}`).font = { bold: true, size: 12, color: { argb: C.blue } }
  r++
  st.getCell(`A${r}`).value = 'المحافظة'
  st.getCell(`A${r}`).font = fH; st.getCell(`A${r}`).fill = hFill(); st.getCell(`A${r}`).alignment = hA()
  st.getCell(`B${r}`).value = 'العدد'
  st.getCell(`B${r}`).font = fH; st.getCell(`B${r}`).fill = hFill(); st.getCell(`B${r}`).alignment = hA()
  const govStart = r
  r++
  st.getCell(`A${r}`).value = fm(
    `IFERROR(INDEX(Applications!$${cl(6)}$2:$${cl(6)}$9999, MATCH(0, INDEX(COUNTIF($${cl(1)}$${govStart}:$${cl(1)}${r-1}, Applications!$${cl(6)}$2:$${cl(6)}$9999), 0), 0)), "")`
  )
  st.getCell(`B${r}`).value = fm(`IF(A${r}="","",COUNTIF(Applications!$${cl(6)}$2:$${cl(6)}$9999, A${r}))`)
  const govDataRow = r

  rows(st, 2, 2)

  // ════════════════════════ 3 · Search ════════════════════════
  const sr = wb.addWorksheet('Search', { properties: { tabColor: { argb: '2D5A8E' } } })
  sr.columns = APP_COLS; hdr(sr, APP_COLS.length)

  sr.mergeCells('A2:A2')
  sr.getCell('A2').value = '🔍 اكتب رقم الطلب في الخلية B2'
  sr.getCell('A2').font = { bold: true, color: { argb: C.blue }, size: 10 }
  sr.getCell('A2').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'EDF2F9' } }
  sr.getCell('A2').alignment = hA(); sr.getCell('A2').border = bd()

  sr.getCell('B2').value = ''
  sr.getCell('B2').font = { size: 12 }
  sr.getCell('B2').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF9E0' } }
  sr.getCell('B2').alignment = hA(); sr.getCell('B2').border = bd()
  sr.getRow(2).height = 30

  // VLOOKUP across all columns
  for (let c = 1; c <= APP_COLS.length; c++)
    sr.getRow(3).getCell(c).value = fm(`IFERROR(VLOOKUP($B$2, Applications!$A$2:$Z$9999, ${c}, FALSE), "")`)

  // ════════════════════════ 4 · Documents ════════════════════════
  const dc = wb.addWorksheet('Documents', { properties: { tabColor: { argb: C.green } } })
  dc.columns = DOC_COLS; hdr(dc, DOC_COLS.length)
  dc.addRow({}); rows(dc, DOC_COLS.length, 2)

  for (let c = 3; c <= DOC_COLS.length; c++) {
    const colL = cl(c)
    dc.getColumn(c).eachCell((cell, rn) => { if (rn > 1) { cell.dataValidation = { type: 'list', formulae: ['"✓,✗"'], allowBlank: true }; cell.value = '✗' } })
    dc.addConditionalFormatting({ ref: `${colL}2:${colL}1048576`, rules: [
      { type: 'expression', formulae: [`${colL}2="✓"`], style: { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'E4F4E4' } }, font: { color: { argb: C.green }, bold: true, size: 14 } } },
      { type: 'expression', formulae: [`${colL}2="✗"`], style: { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FBE4E4' } }, font: { color: { argb: C.red }, bold: true, size: 14 } } },
    ]})
  }

  // ════════════════════════ 5 · Dashboard ════════════════════════
  const db = wb.addWorksheet('Dashboard', { properties: { tabColor: { argb: C.purple } } })
  db.columns = [
    { header: 'بطاقات الأداء', key: 'a1', width: 24 }, { header: 'القيمة', key: 'b1', width: 14 },
    { header: '', key: 'sp1', width: 3 },
    { header: 'مؤشرات الأداء', key: 'a2', width: 22 }, { header: 'القيمة', key: 'b2', width: 16 },
    { header: '', key: 'sp2', width: 4 },
    { header: '📅 التسجيلات اليومية', key: 'a3', width: 14 }, { header: 'العدد', key: 'b3', width: 14 },
    { header: '', key: 'sp3', width: 4 },
    { header: '🏫 المدارس', key: 'a4', width: 18 }, { header: 'العدد', key: 'b4', width: 12 },
    { header: '', key: 'sp4', width: 4 },
    { header: '🌍 المحافظات', key: 'a5', width: 16 }, { header: 'العدد', key: 'b5', width: 12 },
  ]
  // Custom header — merge sections
  hdr(db, 16)
  // override header row to show section names
  db.getRow(1).getCell(1).value = '📊 بطاقات الأداء'
  db.getRow(1).getCell(4).value = '📈 مؤشرات الأداء'
  db.getRow(1).getCell(7).value = '📅 التسجيلات اليومية'
  db.getRow(1).getCell(10).value = '🏫 توزيع المدارس'
  db.getRow(1).getCell(13).value = '🌍 توزيع المحافظات'

  // ── Col A-B: Summary cards ──
  const cardColors = [C.blue, C.green, C.red, C.yellow, C.orange, C.gray]
  const cards = [
    '📋 إجمالي الطلبات', '🟢 المقبولين', '🔴 المرفوضين',
    '🟡 تحت المراجعة', '🟠 ناقص مستندات', '⏳ بانتظار التواصل',
  ]
  cards.forEach((c, i) => {
    const rowI = i + 2
    db.getCell(`A${rowI}`).value = c
    db.getCell(`A${rowI}`).font = { bold: true, size: 12, color: { argb: cardColors[i] } }
    db.getCell(`B${rowI}`).value = fm(`=Statistics!B${i + 2}`)
    db.getCell(`B${rowI}`).font = { bold: true, size: 18, color: { argb: cardColors[i] } }
  })

  // ── Col D-E: KPIs ──
  const dateCL = cl(12) // submission date column
  const kpis = [
    { lbl: 'آخر تسجيل',          f: `=MAX(Applications!${dateCL}2:${dateCL}1048576)` },
    { lbl: 'أول تسجيل',          f: `=MIN(Applications!${dateCL}2:${dateCL}1048576)` },
    { lbl: 'متوسط التسجيل يوميًا', f: `=IFERROR(ROUND(Statistics!B2/MAX(DAYS(MAX(Applications!${dateCL}2:${dateCL}1048576),MIN(Applications!${dateCL}2:${dateCL}1048576)),1),1),0)` },
    { lbl: 'نسبة اكتمال البيانات', f: `=IFERROR(ROUND(COUNTIF(Applications!A2:A1048576,"<>")/COUNTA(Applications!A2:A1048576)*100,1)&"%","0%")` },
  ]
  kpis.forEach((k, i) => {
    const rowI = i + 2
    db.getCell(`D${rowI}`).value = k.lbl
    db.getCell(`D${rowI}`).font = { bold: true, size: 11, color: { argb: C.blue } }
    db.getCell(`E${rowI}`).value = fm(k.f)
    db.getCell(`E${rowI}`).font = { bold: true, size: 13, color: { argb: C.teal } }
  })

  // ── Col G-H: Daily registrations (unique dates + counts, compatible) ──
  db.getRow(2).getCell(7).value = 'التاريخ'
  db.getRow(2).getCell(7).font = fH; db.getRow(2).getCell(7).fill = hFill(); db.getRow(2).getCell(7).alignment = hA()
  db.getRow(2).getCell(8).value = 'العدد'
  db.getRow(2).getCell(8).font = fH; db.getRow(2).getCell(8).fill = hFill(); db.getRow(2).getCell(8).alignment = hA()

  db.getCell('G3').value = fm(
    `IFERROR(INDEX(Applications!$${dateCL}$2:$${dateCL}$9999, MATCH(0, INDEX(COUNTIF($G$2:$G2, Applications!$${dateCL}$2:$${dateCL}$9999), 0), 0)), "")`
  )
  db.getCell('H3').value = fm(`IF(G3="","",COUNTIF(Applications!$${dateCL}$2:$${dateCL}$9999, G3))`)

  // ── Col J-K: School distribution (references Statistics) ──
  db.getRow(2).getCell(10).value = 'المدرسة'
  db.getRow(2).getCell(10).font = fH; db.getRow(2).getCell(10).fill = hFill(); db.getRow(2).getCell(10).alignment = hA()
  db.getRow(2).getCell(11).value = 'العدد'
  db.getRow(2).getCell(11).font = fH; db.getRow(2).getCell(11).fill = hFill(); db.getRow(2).getCell(11).alignment = hA()
  db.getCell(`J3`).value = fm(`=Statistics!A${schoolDataRow}:A1048576`)
  db.getCell(`K3`).value = fm(`=Statistics!B${schoolDataRow}:B1048576`)

  // ── Col M-N: Governorate distribution (references Statistics) ──
  db.getRow(2).getCell(13).value = 'المحافظة'
  db.getRow(2).getCell(13).font = fH; db.getRow(2).getCell(13).fill = hFill(); db.getRow(2).getCell(13).alignment = hA()
  db.getRow(2).getCell(14).value = 'العدد'
  db.getRow(2).getCell(14).font = fH; db.getRow(2).getCell(14).fill = hFill(); db.getRow(2).getCell(14).alignment = hA()
  db.getCell(`M3`).value = fm(`=Statistics!A${govDataRow}:A1048576`)
  db.getCell(`N3`).value = fm(`=Statistics!B${govDataRow}:B1048576`)

  db.views = [{ state: 'frozen', ySplit: 1 }]

  // ════════════════════════ Save ════════════════════════
  const f = 'WE_Schools_v2.xlsx'
  await wb.xlsx.writeFile(f)
  console.log(`\n  ✅ Created: ${f}`)
  console.log(`  ${APP_COLS.length} cols · 5 sheets · Excel 2019+ (no CSE, no QUERY/FILTER/UNIQUE)\n`)
}

main().catch(e => { console.error('Error:', e.message); process.exit(1) })
