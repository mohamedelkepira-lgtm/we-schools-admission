/**
 * WE Schools — Admission Portal Sheet Setup
 *
 * 1. أنشأ جدول جديد (File > New Spreadsheet)
 * 2. Extensions > Apps Script
 * 3. إلصق المحتوى → احفظ → شغّل setupWESheet()
 * 4. خذ Sheet ID من الرابط (gid/XXXXX)
 * 5. حطه في GOOGLE_SHEET_ID (إعدادات Vercel)
 * 6. الموقع هيكتب تلقائي في Applications — باقي الشيتات بتتحسب لوحدها
 */

function setupWESheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  ss.setName('WE Schools - Admission Portal')

  // ──── Sheets ────
  const app = ss.getSheetByName('Sheet1'); if (app) app.setName('Applications')
  const names = ['Statistics', 'Search', 'Documents', 'Dashboard']
  const sheets = { Applications: app || ss.insertSheet('Applications') }
  for (const n of names) sheets[n] = ss.getSheetByName(n) || ss.insertSheet(n)

  // ═══════════════════════════ Applications (26 cols — matches API) ═══════════
  const a = sheets.Applications; a.clear()
  const appH = [
    'رقم الطلب','اسم الطالب','كود الطالب','المدرسة السابقة','الصف الدراسي السابق',
    'المحافظة','المركز / المدينة','رقم الطالب','رقم ولي الأمر','الحالة الاجتماعية',
    'حالة الطلب','تاريخ التقديم','وقت آخر تعديل','الاسم بالإنجليزية','تاريخ الميلاد',
    'الجنس','الجنسية','الديانة','مهنة الأب','مهنة الأم','العنوان','البريد الإلكتروني',
    'رقم بطاقة الطالب','اسم ولي الأمر','رقم بطاقة ولي الأمر','بريد ولي الأمر',
  ]
  a.getRange(1,1,1,26).setValues([appH]); a.setFrozenRows(1)
  for (let c = 1; c <= 26; c++) a.setColumnWidth(c, c <= 13 ? 140 : 160)

  // Data validation
  a.getRange('J2:J').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['أعزب','متزوج'],true).setAllowInvalid(false).build())
  a.getRange('K2:K').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['قيد المراجعة','مقبول','مرفوض','ناقص مستندات','بانتظار التواصل'],true).setAllowInvalid(false).build())

  // Conditional formatting — status (col K)
  const statusClr = [
    {v:'مقبول',bg:'#e4f4e4',fg:'#1e7e1e'},
    {v:'مرفوض',bg:'#fbe4e4',fg:'#b82424'},
    {v:'قيد المراجعة',bg:'#fef7da',fg:'#ab7f12'},
    {v:'ناقص مستندات',bg:'#fff0e0',fg:'#cc6600'},
    {v:'بانتظار التواصل',bg:'#e8ecf0',fg:'#6b7280'},
  ]
  const rules = []
  for (const s of statusClr) {
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(s.v).setBackground(s.bg).setFontColor(s.fg).setRanges([a.getRange('A2:Z')]).build())
  }
  a.setConditionalFormatRules(rules)

  // ═══════════════════════════ Statistics ═══════════════════════════
  const st = sheets.Statistics; st.clear()
  st.getRange('A1:B1').setValues([['الإحصائية','القيمة']])
  const stats = [
    ['📋 إجمالي المتقدمين',     '=COUNTA(Applications!A2:A)'],
    ['🟢 المقبولين',             '=COUNTIF(Applications!K2:K,"مقبول")'],
    ['🔴 المرفوضين',             '=COUNTIF(Applications!K2:K,"مرفوض")'],
    ['🟡 قيد المراجعة',          '=COUNTIF(Applications!K2:K,"قيد المراجعة")'],
    ['🟠 ناقص مستندات',         '=COUNTIF(Applications!K2:K,"ناقص مستندات")'],
    ['⏳ بانتظار التواصل',       '=COUNTIF(Applications!K2:K,"بانتظار التواصل")'],
  ]
  stats.forEach((r,i) => st.getRange(i+2,1,1,2).setValues([r]))
  st.getRange('B2').setFontSize(18).setFontWeight('bold').setFontColor('#1e3a5f')
  st.getRange('B3:B7').setFontSize(16).setFontWeight('bold')
  st.getRange('B3').setFontColor('#1e7e1e'); st.getRange('B4').setFontColor('#b82424')
  st.getRange('B5').setFontColor('#ab7f12'); st.getRange('B6').setFontColor('#cc6600')
  st.getRange('B7').setFontColor('#6b7280')

  // School distribution
  st.getRange('A9:A10').setValues([['🏫 توزيع المدارس'],['المدرسة']])
  st.getRange('B10').setValue('العدد')
  st.getRange('A11').setFormula('=QUERY(Applications!D2:D,"select D,count(D) where D is not null group by D order by count(D) desc label D \'المدرسة\', count(D) \'العدد\'")')

  // Governorate distribution
  st.getRange('D9:D10').setValues([['🌍 توزيع المحافظات'],['المحافظة']])
  st.getRange('E10').setValue('العدد')
  st.getRange('D11').setFormula('=QUERY(Applications!F2:F,"select F,count(F) where F is not null group by F order by count(F) desc label F \'المحافظة\', count(F) \'العدد\'")')

  st.setFrozenRows(1); st.setColumnWidth(1,250); st.setColumnWidth(2,150)

  // ═══════════════════════════ Search ═══════════════════════════
  const sr = sheets.Search; sr.clear()
  sr.getRange(1,1,1,26).setValues([appH]); sr.setFrozenRows(1)
  sr.getRange('A2').setValue('🔍 اكتب رقم الطلب / اسم الطالب / كود الطالب').setFontSize(10).setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle')
  sr.getRange('B2').setFontSize(12).setHorizontalAlignment('center').setBackground('#fff9e0')
  sr.getRow(2).setHeight(30)
  sr.getRange('A3').setFormula('=IFERROR(FILTER(Applications!A:Z, (Applications!A:A=B2)+(Applications!B:B=B2)+(Applications!C:C=B2)), "ابحث بالرقم أو الاسم أو الكود")')
  sr.setColumnWidth(1,300); sr.setColumnWidth(2,400)

  // ═══════════════════════════ Documents ═══════════════════════════
  const d = sheets.Documents; d.clear()
  const docH = ['رقم الطلب','اسم الطالب','شهادة الميلاد','بطاقة الطالب','بطاقة ولي الأمر','الصور الشخصية','ملف التقديم']
  d.getRange(1,1,1,7).setValues([docH]); d.setFrozenRows(1)
  const dv = SpreadsheetApp.newDataValidation().requireValueInList(['تم التسليم','غير مسلم'],true).setAllowInvalid(false).build()
  for (let c = 3; c <= 7; c++) d.getRange(2,c,500).setDataValidation(dv)

  // Conditional formatting
  const docRules = []
  for (let c = 3; c <= 7; c++) {
    docRules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('تم التسليم').setBackground('#e4f4e4').setFontColor('#1e7e1e').setRanges([d.getRange(2,c,500)]).build())
    docRules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('غير مسلم').setBackground('#fbe4e4').setFontColor('#b82424').setRanges([d.getRange(2,c,500)]).build())
  }
  d.setConditionalFormatRules(docRules)

  // ═══════════════════════════ Dashboard ═══════════════════════════
  const db = sheets.Dashboard; db.clear()

  // Summary cards
  db.getRange('A1:B1').setValues([['بطاقات الأداء','القيمة']])
  const cards = [
    ['📋 إجمالي الطلبات','=Statistics!B2'],['🟢 المقبولين','=Statistics!B3'],
    ['🔴 المرفوضين','=Statistics!B4'],['🟡 تحت المراجعة','=Statistics!B5'],
    ['🟠 ناقص مستندات','=Statistics!B6'],['⏳ بانتظار التواصل','=Statistics!B7'],
  ]
  cards.forEach((r,i) => db.getRange(i+2,1,1,2).setValues([r]))
  db.getRange('B2:B7').setFontSize(18).setFontWeight('bold')
  db.getRange('B2').setFontColor('#1e3a5f'); db.getRange('B3').setFontColor('#1e7e1e')
  db.getRange('B4').setFontColor('#b82424'); db.getRange('B5').setFontColor('#ab7f12')
  db.getRange('B6').setFontColor('#cc6600'); db.getRange('B7').setFontColor('#6b7280')

  // KPIs
  db.getRange('D1:E1').setValues([['مؤشرات الأداء','القيمة']])
  const kpis = [
    ['آخر تسجيل','=MAX(Applications!L2:L)'],
    ['أول تسجيل','=MIN(Applications!L2:L)'],
    ['متوسط التسجيل يوميًا','=IFERROR(ROUND(Statistics!B2/MAX(DAYS(MAX(Applications!L2:L),MIN(Applications!L2:L)),1),1),0)'],
    ['نسبة اكتمال البيانات','=IFERROR(ROUND(COUNTIF(Applications!A2:A,"<>")/COUNTA(Applications!A2:A)*100,1)&"%","0%")'],
  ]
  kpis.forEach((r,i) => db.getRange(i+2,4,1,2).setValues([r]))

  // Daily registrations chart data
  db.getRange('G1:H1').setValues([['التاريخ','عدد المتقدمين']])
  db.getRange('G2').setFormula('=UNIQUE(Applications!L2:L)')
  db.getRange('H2').setFormula('=ARRAYFORMULA(IF(G2:G="",,COUNTIF(Applications!L2:L,G2:G)))')

  // School distribution chart data
  db.getRange('J1:K1').setValues([['المدرسة','العدد']])
  db.getRange('J2').setFormula('=Statistics!A11:A')
  db.getRange('K2').setFormula('=Statistics!B11:B')

  // Governorate distribution chart data
  db.getRange('M1:N1').setValues([['المحافظة','العدد']])
  db.getRange('M2').setFormula('=Statistics!D11:D')
  db.getRange('N2').setFormula('=Statistics!E11:E')

  db.setFrozenRows(1)

  // ──── Charts ────
  const oldCharts = db.getCharts(); for (const c of oldCharts) db.removeChart(c)

  db.insertChart(db.newChart().setChartType(Charts.ChartType.LINE).addRange(db.getRange('G1:H')).setPosition(7,1,0,0)
    .setOption('title','📈 عدد المتقدمين يوميًا').setOption('width',500).setOption('height',300)
    .setOption('legend',{position:'none'}).setOption('colors',['#1e3a5f']).setOption('curveType','function').build())

  db.insertChart(db.newChart().setChartType(Charts.ChartType.PIE).addRange(db.getRange('J1:K')).setPosition(7,6,0,0)
    .setOption('title','🏫 أكثر المدارس المتقدم منها').setOption('width',450).setOption('height',300)
    .setOption('pieSliceText','label').setOption('colors',['#1e3a5f','#c8952e','#2d5a8e','#e8c469','#5a7fa0']).build())

  db.insertChart(db.newChart().setChartType(Charts.ChartType.PIE).addRange(db.getRange('A3:B5')).setPosition(18,1,0,0)
    .setOption('title','🟢 حالة الطلبات').setOption('width',400).setOption('height',250)
    .setOption('colors',['#1e7e1e','#b82424','#ab7f12']).build())

  // ──── Global formatting ────
  for (const [k,sh] of Object.entries(sheets)) {
    const last = sh.getLastColumn()||1
    sh.getRange(1,1,1,last).setBackground('#1e3a5f').setFontColor('#ffffff').setFontWeight('bold').setFontSize(11).setHorizontalAlignment('center').setVerticalAlignment('middle')
    sh.getRange(2,1,Math.max(sh.getLastRow(),500),last).setAlternatingBackgroundColors(['#f8f9fb','#ffffff'])
  }

  // ──── Done ────
  SpreadsheetApp.flush()
  ss.toast('✅ WE Schools جاهز — API هيكتب في Applications', 'تم', 5)
  console.log('Sheet ID:', ss.getId())
}
