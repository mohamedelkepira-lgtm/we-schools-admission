/**
 * WE Schools — Admission Portal Sheet Setup
 *
 * كيف تستخدمه:
 * 1. افتح Google Sheets → أنشأ جدول جديد (File > New Spreadsheet)
 * 2. Extensions > Apps Script
 * 3. إلصق كل محتوى هذا الملف
 * 4. احفظ (Ctrl+S) وسمّ المشروع "WE Schools Setup"
 * 5. شغّل الدالة setupWESheet() — أول مرة هتطلب أذونات، وافق
 * 6. بعد ما يخلص، ارجع للـ Sheet — هتلاقي 5 شيتات جاهزة
 */

function setupWESheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  ss.setName('WE Schools - Admission Portal')

  // ──── Rename default sheet ────
  const app = ss.getSheetByName('Sheet1')
  if (app) app.setName('Applications')

  // ──── Add sheets ────
  const names = ['Statistics', 'Search', 'Documents', 'Dashboard']
  const sheets = { Applications: app }
  for (const n of names) {
    const s = ss.getSheetByName(n)
    sheets[n] = s || ss.insertSheet(n)
  }

  // ──── Applications headers (20 cols) ────
  const appHeaders = [
    'رقم الطلب', 'اسم الطالب', 'كود الطالب', 'المدرسة السابقة',
    'رقم الطالب', 'رقم ولي الأمر', 'الحالة الاجتماعية', 'حالة الطلب', 'تاريخ التقديم',
    'الاسم بالإنجليزية', 'تاريخ الميلاد', 'الجنس', 'الجنسية', 'الديانة',
    'العنوان', 'البريد الإلكتروني', 'رقم بطاقة الطالب',
    'اسم ولي الأمر', 'رقم بطاقة ولي الأمر', 'بريد ولي الأمر',
  ]
  const a = sheets.Applications
  a.clear()
  a.getRange(1, 1, 1, appHeaders.length).setValues([appHeaders])
  a.setFrozenRows(1)

  // ──── Statistics ────
  const s = sheets.Statistics
  s.clear()
  s.getRange('A1:B1').setValues([['الإحصائية', 'القيمة']])
  s.getRange('A2').setValue('إجمالي المتقدمين')
  s.getRange('B2').setFormula('=COUNTA(Applications!A2:A)')
  s.getRange('A3').setValue('عدد المقبولين')
  s.getRange('B3').setFormula('=COUNTIF(Applications!H2:H, "مقبول")')
  s.getRange('A4').setValue('عدد المرفوضين')
  s.getRange('B4').setFormula('=COUNTIF(Applications!H2:H, "مرفوض")')
  s.getRange('A5').setValue('قيد المراجعة')
  s.getRange('B5').setFormula('=COUNTIF(Applications!H2:H, "قيد المراجعة")')

  s.getRange('A7').setValue('الطلاب حسب المدرسة')
  s.getRange('A8:B8').setValues([['المدرسة', 'العدد']])
  s.getRange('A9').setFormula(
    '=QUERY(Applications!D2:D, "select D, count(D) where D is not null group by D order by count(D) desc label D \'المدرسة\', count(D) \'العدد\'")'
  )

  s.getRange('D7').setValue('الطلاب حسب المحافظة')
  s.getRange('D8:E8').setValues([['المحافظة', 'العدد']])
  // Uses Applications!N:N (الديانة column as proxy — replace with actual gov column if added)
  // s.getRange('D9').setFormula('=QUERY(Applications!N2:N, "select N, count(N) where N is not null group by N order by count(N) desc label N \'المحافظة\', count(N) \'العدد\'")')

  s.getRange('D7').setNote('أضف عمود "المحافظة" في Applications ثم فعّل الصيغة في D9')
  s.setFrozenRows(1)

  // ──── Search ────
  const search = sheets.Search
  search.clear()
  search.getRange('A1:I1').setValues([appHeaders.slice(0, 9)])
  search.getRange('A2').setValue('🔍 اكتب رقم الطلب / اسم الطالب / كود الطالب').setFontSize(10)
  search.getRange('B2').setFontSize(12).setHorizontalAlignment('center')
  search.setFrozenRows(1)
  // Merged label cell
  search.getRange('A2:A2').merge().setHorizontalAlignment('center').setVerticalAlignment('middle')
  search.getRange('A3').setFormula(
    '=IFERROR(FILTER(Applications!A:I, (Applications!A:A=B2)+(Applications!B:B=B2)+(Applications!C:C=B2)), "لا توجد نتائج — أدخل رقم الطلب أو الاسم أو الكود في الخلية B2")'
  )

  // ──── Documents ────
  const docHeaders = [
    'رقم الطلب', 'اسم الطالب', 'شهادة الميلاد', 'بطاقة الطالب',
    'بطاقة ولي الأمر', 'الصور الشخصية', 'ملف التقديم',
  ]
  const d = sheets.Documents
  d.clear()
  d.getRange(1, 1, 1, docHeaders.length).setValues([docHeaders])
  d.setFrozenRows(1)

  // Document status dropdowns
  const docStatusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['تم التسليم', 'غير مسلم'], true)
    .setAllowInvalid(false)
    .build()

  for (let col = 3; col <= docHeaders.length; col++) {
    const range = d.getRange(2, col, 500)
    range.setDataValidation(docStatusRule)
  }

  // ──── Dashboard ────
  const db = sheets.Dashboard
  db.clear()

  // Summary cards
  const cardHeaders = ['المؤشر', 'القيمة']
  db.getRange('A1:B1').setValues([cardHeaders])
  db.getRange('A2').setValue('إجمالي المتقدمين')
  db.getRange('B2').setFormula('=Statistics!B2')
  db.getRange('A3').setValue('المقبولين')
  db.getRange('B3').setFormula('=Statistics!B3')
  db.getRange('A4').setValue('المرفوضين')
  db.getRange('B4').setFormula('=Statistics!B4')
  db.getRange('A5').setValue('قيد المراجعة')
  db.getRange('B5').setFormula('=Statistics!B5')

  // Daily applicants chart data
  db.getRange('D1:E1').setValues([['التاريخ', 'عدد المتقدمين']])
  db.getRange('D2').setFormula(
    '=UNIQUE(Applications!I2:I)'
  )
  db.getRange('E2').setFormula(
    '=ARRAYFORMULA(IF(D2:D="",,COUNTIF(Applications!I2:I, D2:D)))'
  )

  // School distribution chart data
  db.getRange('G1:H1').setValues([['المدرسة', 'العدد']])
  db.getRange('G2').setFormula('=Statistics!A9:A')
  db.getRange('H2').setFormula('=Statistics!B9:B')

  db.setFrozenRows(1)

  // ──── Formatting ────
  const headerBg = '#1e3a5f'
  const headerFg = '#ffffff'

  for (const [key, sheet] of Object.entries(sheets)) {
    const lastCol = sheet.getLastColumn() || 1
    const headerRange = sheet.getRange(1, 1, 1, lastCol)
    headerRange
      .setBackground(headerBg)
      .setFontColor(headerFg)
      .setFontWeight('bold')
      .setFontSize(11)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
    sheet.setRowHeight(1, 32)

    // Alternating rows
    const dataRange = sheet.getRange(2, 1, Math.max(sheet.getLastRow(), 500), lastCol)
    if (key !== 'Dashboard') {
      dataRange.setAlternatingBackgroundColors(['#f8f9fb', '#ffffff'])
    }
  }

  // Applications: column widths
  a.setColumnWidth(1, 130)  // request #
  a.setColumnWidth(2, 180)  // student name
  a.setColumnWidth(3, 140)  // code
  a.setColumnWidth(4, 200)  // school
  a.setColumnWidth(5, 130)  // student phone
  a.setColumnWidth(6, 130)  // parent phone
  a.setColumnWidth(7, 120)  // social status
  a.setColumnWidth(8, 120)  // app status
  a.setColumnWidth(9, 150)  // date
  for (let c = 10; c <= 20; c++) a.setColumnWidth(c, 150)

  // Statistics column widths
  s.setColumnWidth(1, 250)
  s.setColumnWidth(2, 150)

  // Search column widths
  search.setColumnWidth(1, 350) // merged label
  search.setColumnWidth(2, 400) // input cell

  // Documents column widths
  d.setColumnWidth(1, 130)
  d.setColumnWidth(2, 180)
  for (let c = 3; c <= docHeaders.length; c++) d.setColumnWidth(c, 130)

  // Dashboard column widths
  db.setColumnWidth(1, 200)
  db.setColumnWidth(2, 120)

  // ──── Conditional formatting: Applications & Documents ────
  // Applications — status colors (column H)
  const appStatusRange = a.getRange('H2:H')
  const accRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('مقبول')
    .setBackground('#e4f4e4')
    .setFontColor('#1e7e1e')
    .setRanges([a.getRange('A2:T')])
    .build()
  const rejRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('مرفوض')
    .setBackground('#fbe4e4')
    .setFontColor('#b82424')
    .setRanges([a.getRange('A2:T')])
    .build()
  const penRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('قيد المراجعة')
    .setBackground('#fef7da')
    .setFontColor('#ab7f12')
    .setRanges([a.getRange('A2:T')])
    .build()
  const rules = a.getConditionalFormatRules()
  rules.push(accRule, rejRule, penRule)
  a.setConditionalFormatRules(rules)

  // Documents — status colors
  const docRules = []
  for (let col = 3; col <= docHeaders.length; col++) {
    const r = d.getRange(2, col, 500)
    docRules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('تم التسليم')
        .setBackground('#e4f4e4')
        .setFontColor('#1e7e1e')
        .setRanges([r])
        .build()
    )
    docRules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('غير مسلم')
        .setBackground('#fbe4e4')
        .setFontColor('#b82424')
        .setRanges([r])
        .build()
    )
  }
  const existingDocRules = d.getConditionalFormatRules()
  d.setConditionalFormatRules(existingDocRules.concat(docRules))

  // ──── Data validation: Applications ────
  // Social status (G)
  const socialRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['أعزب', 'متزوج'], true)
    .setAllowInvalid(false)
    .build()
  a.getRange('G2:G').setDataValidation(socialRule)

  // Application status (H)
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['قيد المراجعة', 'مقبول', 'مرفوض'], true)
    .setAllowInvalid(false)
    .build()
  a.getRange('H2:H').setDataValidation(statusRule)

  // ──── Charts: Dashboard ────
  // Remove existing charts
  const charts = db.getCharts()
  for (const chart of charts) db.removeChart(chart)

  // 1. Daily applicants — line chart
  const dailyData = db.getRange('D1:E')
  const dailyChart = db.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(dailyData)
    .setPosition(7, 1, 0, 0)
    .setOption('title', '📈 عدد المتقدمين يوميًا')
    .setOption('width', 500)
    .setOption('height', 300)
    .setOption('legend', { position: 'none' })
    .setOption('colors', ['#1e3a5f'])
    .setOption('curveType', 'function')
    .build()
  db.insertChart(dailyChart)

  // 2. School distribution — pie chart
  const schoolData = db.getRange('G1:H')
  const schoolChart = db.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(schoolData)
    .setPosition(7, 6, 0, 0)
    .setOption('title', '🏫 أكثر المدارس المتقدم منها')
    .setOption('width', 450)
    .setOption('height', 300)
    .setOption('pieSliceText', 'label')
    .setOption('colors', ['#1e3a5f', '#c8952e', '#2d5a8e', '#e8c469', '#5a7fa0'])
    .build()
  db.insertChart(schoolChart)

  // 3. Application status — pie chart
  const statusData = db.getRange('A3:B5')
  const statusChart = db.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(statusData)
    .setPosition(18, 1, 0, 0)
    .setOption('title', '🟢 حالة الطلبات')
    .setOption('width', 400)
    .setOption('height', 250)
    .setOption('colors', ['#1e7e1e', '#b82424', '#ab7f12'])
    .build()
  db.insertChart(statusChart)

  // 4. Completion gauge — bar chart
  const compData = db.getRange('A1:B5')
  const compChart = db.newChart()
    .setChartType(Charts.ChartType.BAR)
    .addRange(compData)
    .setPosition(18, 6, 0, 0)
    .setOption('title', '📊 إحصائيات عامة')
    .setOption('width', 450)
    .setOption('height', 250)
    .setOption('colors', ['#1e3a5f'])
    .setOption('hAxis', { title: 'العدد' })
    .build()
  db.insertChart(compChart)

  // ──── Protection: prevent accidental header edits ────
  const allSheets = Object.values(sheets)
  for (const sheet of allSheets) {
    const proto = sheet.getRange(1, 1, 1, sheet.getLastColumn() || 1).protect()
    proto.setDescription('Header row — protected')
    const me = Session.getEffectiveUser()
    proto.addEditor(me)
    proto.removeEditors(proto.getEditors().filter(e => e.getEmail() !== me.getEmail()))
  }

  // ──── Done ────
  SpreadsheetApp.flush()
  ss.toast('✅ WE Schools Admission Portal جاهز!', 'تم', 5)
  console.log('Setup complete!')
}
