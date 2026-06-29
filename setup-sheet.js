import 'dotenv/config'
import { createPrivateKey, sign } from 'crypto'

const HEADERS = {
  Applications: [
    'رقم الطلب', 'اسم الطالب', 'كود الطالب', 'المدرسة السابقة',
    'رقم الطالب', 'رقم ولي الأمر', 'الحالة الاجتماعية', 'حالة الطلب', 'تاريخ التقديم',
    'الاسم بالإنجليزية', 'تاريخ الميلاد', 'الجنس', 'الجنسية', 'الديانة',
    'العنوان', 'البريد الإلكتروني', 'رقم بطاقة الطالب',
    'اسم ولي الأمر', 'رقم بطاقة ولي الأمر', 'بريد ولي الأمر',
  ],
  Statistics: [
    'الإحصائية', 'القيمة',
  ],
  Search: [
    'رقم الطلب', 'اسم الطالب', 'كود الطالب', 'المدرسة السابقة',
    'رقم الطالب', 'رقم ولي الأمر', 'الحالة الاجتماعية', 'حالة الطلب', 'تاريخ التقديم',
  ],
  Documents: [
    'رقم الطلب', 'اسم الطالب', 'شهادة الميلاد', 'بطاقة الطالب',
    'بطاقة ولي الأمر', 'الصور الشخصية', 'ملف التقديم',
  ],
  Dashboard: [
    'المؤشر', 'القيمة', 'التاريخ',
  ],
}

function base64url(str) {
  return Buffer.from(str).toString('base64url')
}

async function getAccessToken(clientEmail, privateKey) {
  const now = Math.floor(Date.now() / 1000)
  const header = base64url(JSON.stringify({ alg: 'RS256', typ: 'JWT' }))
  const payload = base64url(JSON.stringify({
    iss: clientEmail,
    scope: 'https://www.googleapis.com/auth/spreadsheets',
    aud: 'https://oauth2.googleapis.com/token',
    exp: now + 3600,
    iat: now,
  }))

  const key = createPrivateKey(privateKey.replace(/\\n/g, '\n'))
  const sig = sign(null, Buffer.from(`${header}.${payload}`), key)
  const jwt = `${header}.${payload}.${base64url(sig)}`

  const res = await fetch('https://oauth2.googleapis.com/token', {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer',
      assertion: jwt,
    }),
  })

  const data = await res.json()
  if (!data.access_token) throw new Error('Google Auth failed: ' + JSON.stringify(data))
  return data.access_token
}

async function api(url, token, body) {
  const opts = { headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' } }
  if (body) opts.body = JSON.stringify(body)
  const res = await fetch(url, opts)
  const data = await res.json()
  if (data.error) throw new Error(data.error.message || JSON.stringify(data.error))
  return data
}

async function createSpreadsheet(token) {
  console.log('Creating spreadsheet...')
  const data = await api(
    'https://sheets.googleapis.com/v4/spreadsheets',
    token,
    { properties: { title: 'WE Schools - Admission Portal' } }
  )
  console.log('  Created:', data.spreadsheetUrl)
  return data.spreadsheetId
}

async function setupSheets(token, id) {
  console.log('Setting up sheet names...')

  const addSheet = (title) => ({
    addSheet: {
      properties: { title, index: 1 },
    },
  })

  const requests = [
    {
      updateSheetProperties: {
        properties: { sheetId: 0, title: 'Applications' },
        fields: 'title',
      },
    },
  ]

  const sheetTitles = ['Statistics', 'Search', 'Documents', 'Dashboard']
  for (const title of sheetTitles) {
    requests.push(addSheet(title))
  }

  await api(
    `https://sheets.googleapis.com/v4/spreadsheets/${id}:batchUpdate`,
    token,
    { requests }
  )
  console.log('  Sheets: Applications, Statistics, Search, Documents, Dashboard')
}

async function getSheetIds(token, id) {
  const data = await api(
    `https://sheets.googleapis.com/v4/spreadsheets/${id}?fields=sheets.properties`,
    token
  )
  const map = {}
  for (const s of data.sheets) {
    map[s.properties.title] = s.properties.sheetId
  }
  return map
}

async function writeHeaders(token, id) {
  console.log('Writing headers...')
  const requests = []

  for (const [name, headers] of Object.entries(HEADERS)) {
    const range = `'${name}'!A1:${String.fromCharCode(64 + headers.length)}1`
    await api(
      `https://sheets.googleapis.com/v4/spreadsheets/${id}/values/${encodeURIComponent(range)}?valueInputOption=USER_ENTERED`,
      token,
      { values: [headers] }
    )
  }
}

async function applyFormatting(token, id) {
  console.log('Applying formatting...')
  const sheetIds = await getSheetIds(token, id)

  const requests = []

  for (const [name, sheetId] of Object.entries(sheetIds)) {
    const numCols = HEADERS[name]?.length || 2

    requests.push({
      updateDimensionProperties: {
        range: { sheetId, dimension: 'COLUMNS', startIndex: 0, endIndex: numCols },
        properties: { pixelSize: 140 },
        fields: 'pixelSize',
      },
    })

    requests.push({
      repeatCell: {
        range: { sheetId, startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: numCols },
        cell: {
          userEnteredFormat: {
            backgroundColor: { red: 0.117, green: 0.227, blue: 0.373 },
            textFormat: { bold: true, foregroundColor: { red: 1, green: 1, blue: 1 }, fontSize: 11 },
            horizontalAlignment: 'CENTER',
            verticalAlignment: 'MIDDLE',
          },
        },
        fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
      },
    })
  }

  const sheetsWithData = ['Applications', 'Documents']
  for (const name of sheetsWithData) {
    const sid = sheetIds[name]
    const numCols = HEADERS[name].length

    requests.push({
      addConditionalFormatRule: {
        rule: {
          ranges: [{ sheetId: sid, startRowIndex: 1, startColumnIndex: 0, endColumnIndex: numCols }],
          booleanRule: {
            condition: { type: 'TEXT_EQ', values: [{ userEnteredValue: 'مقبول' }] },
            format: {
              backgroundColor: { red: 0.89, green: 0.96, blue: 0.9 },
              textFormat: { foregroundColor: { red: 0.12, green: 0.52, blue: 0.12 } },
            },
          },
        },
      },
    })

    requests.push({
      addConditionalFormatRule: {
        rule: {
          ranges: [{ sheetId: sid, startRowIndex: 1, startColumnIndex: 0, endColumnIndex: numCols }],
          booleanRule: {
            condition: { type: 'TEXT_EQ', values: [{ userEnteredValue: 'مرفوض' }] },
            format: {
              backgroundColor: { red: 0.98, green: 0.88, blue: 0.88 },
              textFormat: { foregroundColor: { red: 0.72, green: 0.15, blue: 0.15 } },
            },
          },
        },
      },
    })

    requests.push({
      addConditionalFormatRule: {
        rule: {
          ranges: [{ sheetId: sid, startRowIndex: 1, startColumnIndex: 0, endColumnIndex: numCols }],
          booleanRule: {
            condition: { type: 'TEXT_EQ', values: [{ userEnteredValue: 'قيد المراجعة' }] },
            format: {
              backgroundColor: { red: 0.99, green: 0.96, blue: 0.85 },
              textFormat: { foregroundColor: { red: 0.67, green: 0.5, blue: 0.07 } },
            },
          },
        },
      },
    })

    const docStatus = ['تم التسليم', 'غير مسلم']
    for (const status of docStatus) {
      requests.push({
        addConditionalFormatRule: {
          rule: {
            ranges: [{ sheetId: sid, startRowIndex: 1, startColumnIndex: 2, endColumnIndex: numCols }],
            booleanRule: {
              condition: { type: 'TEXT_EQ', values: [{ userEnteredValue: status }] },
              format: status === 'تم التسليم'
                ? { backgroundColor: { red: 0.89, green: 0.96, blue: 0.9 }, textFormat: { foregroundColor: { red: 0.12, green: 0.52, blue: 0.12 } } }
                : { backgroundColor: { red: 0.98, green: 0.88, blue: 0.88 }, textFormat: { foregroundColor: { red: 0.72, green: 0.15, blue: 0.15 } } },
            },
          },
        },
      })
    }
  }

  const appSid = sheetIds.Applications
  const appCols = HEADERS.Applications.length

  requests.push({
    setDataValidation: {
      range: { sheetId: appSid, startRowIndex: 1, startColumnIndex: 6, endColumnIndex: 7 },
      rule: {
        condition: { type: 'ONE_OF_LIST', values: [
          { userEnteredValue: 'أعزب' }, { userEnteredValue: 'متزوج' },
        ]},
        strict: true,
        showCustomUi: true,
      },
    },
  })

  requests.push({
    setDataValidation: {
      range: { sheetId: appSid, startRowIndex: 1, startColumnIndex: 7, endColumnIndex: 8 },
      rule: {
        condition: { type: 'ONE_OF_LIST', values: [
          { userEnteredValue: 'قيد المراجعة' }, { userEnteredValue: 'مقبول' }, { userEnteredValue: 'مرفوض' },
        ]},
        strict: true,
        showCustomUi: true,
      },
    },
  })

  const docSid = sheetIds.Documents
  const docCols = HEADERS.Documents.length
  for (let col = 2; col < docCols; col++) {
    requests.push({
      setDataValidation: {
        range: { sheetId: docSid, startRowIndex: 1, startColumnIndex: col, endColumnIndex: col + 1 },
        rule: {
          condition: { type: 'ONE_OF_LIST', values: [
            { userEnteredValue: 'تم التسليم' }, { userEnteredValue: 'غير مسلم' },
          ]},
          strict: true,
          showCustomUi: true,
        },
      },
    })
  }

  for (const [name, sheetId] of Object.entries(sheetIds)) {
    requests.push({
      updateSheetProperties: {
        properties: {
          sheetId,
          gridProperties: { frozenRowCount: 1 },
        },
        fields: 'gridProperties.frozenRowCount',
      },
    })
  }

  await api(
    `https://sheets.googleapis.com/v4/spreadsheets/${id}:batchUpdate`,
    token,
    { requests }
  )
}

async function writeFormulas(token, id) {
  console.log('Writing formulas...')

  const statsFormulas = [
    ['إجمالي المتقدمين', `=COUNTA(Applications!A2:A)`],
    ['عدد المقبولين', `=COUNTIF(Applications!H2:H, "مقبول")`],
    ['عدد المرفوضين', `=COUNTIF(Applications!H2:H, "مرفوض")`],
    ['قيد المراجعة', `=COUNTIF(Applications!H2:H, "قيد المراجعة")`],
    [],
    ['الطلاب حسب المدرسة'],
    ['المدرسة', 'العدد'],
  ]

  const schoolData = `=QUERY(Applications!D2:D, "select D, count(D) where D is not null group by D order by count(D) desc label D 'المدرسة', count(D) 'العدد'")`
  await api(
    `https://sheets.googleapis.com/v4/spreadsheets/${id}/values/'Statistics'!A1:B7?valueInputOption=USER_ENTERED`,
    token,
    { values: statsFormulas }
  )
  await api(
    `https://sheets.googleapis.com/v4/spreadsheets/${id}/values/'Statistics'!A8?valueInputOption=USER_ENTERED`,
    token,
    { values: [[schoolData]] }
  )

  const searchFormula = `=IFERROR(FILTER(Applications!A:I, (Applications!A:A=B2)+(Applications!B:B=B2)+(Applications!C:C=B2)), "لا توجد نتائج")`
  await api(
    `https://sheets.googleapis.com/v4/spreadsheets/${id}/values/'Search'!A3?valueInputOption=USER_ENTERED`,
    token,
    { values: [[searchFormula]] }
  )
}

async function applySearchFormatting(token, id) {
  const sheetIds = await getSheetIds(token, id)
  const sid = sheetIds.Search
  const numCols = HEADERS.Search.length

  const requests = [
    {
      updateDimensionProperties: {
        range: { sheetId: sid, dimension: 'COLUMNS', startIndex: 1, endIndex: 2 },
        properties: { pixelSize: 300 },
        fields: 'pixelSize',
      },
    },
    {
      mergeCells: {
        range: { sheetId: sid, startRowIndex: 1, endRowIndex: 2, startColumnIndex: 0, endColumnIndex: 1 },
        mergeType: 'MERGE_ALL',
      },
    },
    {
      repeatCell: {
        range: { sheetId: sid, startRowIndex: 1, endRowIndex: 2, startColumnIndex: 0, endColumnIndex: 1 },
        cell: {
          userEnteredValue: { stringValue: 'ابحث بالرقم أو الاسم أو الكود' },
          userEnteredFormat: {
            backgroundColor: { red: 0.95, green: 0.97, blue: 0.99 },
            textFormat: { bold: true, fontSize: 11, foregroundColor: { red: 0.117, green: 0.227, blue: 0.373 } },
            horizontalAlignment: 'CENTER',
            verticalAlignment: 'MIDDLE',
          },
        },
        fields: 'userEnteredValue,userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
      },
    },
    {
      repeatCell: {
        range: { sheetId: sid, startRowIndex: 1, endRowIndex: 2, startColumnIndex: 1, endColumnIndex: 2 },
        cell: {
          userEnteredFormat: {
            backgroundColor: { red: 0.95, green: 0.97, blue: 0.99 },
            textFormat: { fontSize: 12 },
            horizontalAlignment: 'CENTER',
            verticalAlignment: 'MIDDLE',
          },
        },
        fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
      },
    },
  ]

  await api(
    `https://sheets.googleapis.com/v4/spreadsheets/${id}:batchUpdate`,
    token,
    { requests }
  )
}

async function updateApiConfig(token, id) {
  console.log(`\n  GOOGLE_SHEET_ID=${id}`)
  console.log(`  URL: https://docs.google.com/spreadsheets/d/${id}\n`)
}

async function main() {
  const clientEmail = process.env.GOOGLE_CLIENT_EMAIL
  const privateKey = process.env.GOOGLE_PRIVATE_KEY

  if (!clientEmail || !privateKey) {
    console.error('Missing GOOGLE_CLIENT_EMAIL or GOOGLE_PRIVATE_KEY in .env')
    process.exit(1)
  }

  const token = await getAccessToken(clientEmail, privateKey)
  const id = await createSpreadsheet(token)
  await setupSheets(token, id)
  await writeHeaders(token, id)
  await writeFormulas(token, id)
  await applyFormatting(token, id)
  await applySearchFormatting(token, id)
  await updateApiConfig(token, id)

  console.log('Done!')
}

main().catch(err => {
  console.error('Error:', err.message)
  process.exit(1)
})
