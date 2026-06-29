import 'dotenv/config'
import express from 'express'
import { createPrivateKey, sign } from 'crypto'
import { readFileSync, writeFileSync, existsSync } from 'fs'
import { join, dirname } from 'path'
import { fileURLToPath } from 'url'

const __dirname = dirname(fileURLToPath(import.meta.url))
const DATA_FILE = join(__dirname, '_local_data.json')
const app = express()
app.use(express.json())

function base64url(str) {
  return Buffer.from(str).toString('base64url')
}

function generateRequestNumber() {
  const year = new Date().getFullYear()
  return `WE-${year}-${Math.floor(1000 + Math.random() * 9000)}`
}

async function getGoogleToken(clientEmail, privateKey) {
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
    body: new URLSearchParams({ grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer', assertion: jwt }),
  })
  const data = await res.json()
  if (!data.access_token) throw new Error('Google Auth failed')
  return data.access_token
}

async function appendToGoogleSheet(token, sheetId, headers, row) {
  const base = `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/Applications`
  const opts = { headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' } }

  const totalCols = headers.length
  const endCol = String.fromCharCode(64 + totalCols)
  const check = await fetch(`${base}!A1:${endCol}1`, { headers: opts.headers })
  const existing = await check.json()

  if (!existing.values || existing.values.length === 0) {
    await fetch(`${base}!A1:${endCol}1?valueInputOption=USER_ENTERED`, {
      method: 'PUT', ...opts, body: JSON.stringify({ values: [headers] }),
    })
  }

  await fetch(`${base}!A:${endCol}:append?valueInputOption=USER_ENTERED`, {
    method: 'POST', ...opts, body: JSON.stringify({ values: [row] }),
  })
}

async function sendEmail(apiKey, data) {
  await fetch('https://api.resend.com/emails', {
    method: 'POST',
    headers: { Authorization: `Bearer ${apiKey}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({
      from: process.env.EMAIL_FROM || 'WE Schools <noreply@weschools.com>',
      to: [data.email],
      subject: `تأكيد استلام طلب التسجيل - ${data.requestNumber}`,
      html: `<div dir="rtl" style="font-family:'Cairo',Arial,sans-serif;max-width:600px;margin:0 auto">
<div style="background:#1e3a5f;padding:30px;text-align:center;border-radius:16px 16px 0 0">
<h1 style="color:#fff;margin:0;font-size:24px">WE Schools</h1></div>
<div style="background:#f9fafb;padding:30px;border-radius:0 0 16px 16px">
<h2 style="color:#1f2937;margin-top:0">تم استلام طلبك بنجاح</h2>
<p style="color:#6b7280;line-height:1.8">شكراً لتسجيلك في WE Schools. تم استلام طلب التقديم بنجاح.</p>
<div style="background:#dbeafe;border:1px solid #bfdbfe;border-radius:12px;padding:20px;text-align:center;margin:20px 0">
<div style="color:#2563eb;font-size:14px;margin-bottom:4px">رقم الطلب</div>
<div style="color:#1e40af;font-size:28px;font-weight:bold;letter-spacing:2px" dir="ltr">${data.requestNumber}</div></div>
<p style="color:#6b7280;line-height:1.8">سنتواصل معك قريباً لإكمال إجراءات القبول. يرجى الاحتفاظ برقم الطلب لمتابعة حالة طلبك.</p>
<div style="border-top:1px solid #e5e7eb;padding-top:16px;margin-top:24px;text-align:center">
<p style="color:#9ca3af;font-size:12px">WE Schools &copy; ${new Date().getFullYear()} - جميع الحقوق محفوظة</p></div></div></div>`,
    }),
  })
}

app.post('/api/submit-application', async (req, res) => {
  try {
    const data = req.body
    data.requestNumber = generateRequestNumber()
    data.submittedAt = new Date().toISOString()

    const sheetId = process.env.GOOGLE_SHEET_ID
    const privateKey = process.env.GOOGLE_PRIVATE_KEY
    const clientEmail = process.env.GOOGLE_CLIENT_EMAIL

    if (sheetId && privateKey && clientEmail) {
      const token = await getGoogleToken(clientEmail, privateKey)
      const headers = [
        'رقم الطلب', 'اسم الطالب', 'كود الطالب', 'المدرسة السابقة',
        'الصف الدراسي السابق', 'المحافظة', 'المركز / المدينة',
        'رقم الطالب', 'رقم ولي الأمر', 'الحالة الاجتماعية', 'حالة الطلب', 'تاريخ التقديم',
        'وقت آخر تعديل',
        'الاسم بالإنجليزية', 'تاريخ الميلاد', 'الجنس', 'الجنسية', 'الديانة',
        'مهنة الأب', 'مهنة الأم',
        'العنوان', 'البريد الإلكتروني', 'رقم بطاقة الطالب',
        'اسم ولي الأمر', 'رقم بطاقة ولي الأمر', 'بريد ولي الأمر',
      ]
      await appendToGoogleSheet(token, sheetId, headers, [
        data.requestNumber, data.fullNameAr, data.requestNumber,
        data.previousSchool || '', data.previousGrade || '', data.governorate || '', data.city || '',
        data.phone, data.parentPhone,
        'قيد المراجعة', 'قيد المراجعة', data.submittedAt, data.submittedAt || '',
        data.fullNameEn, data.dateOfBirth, data.gender,
        data.nationality, data.religion,
        data.fatherOccupation || '', data.motherOccupation || '',
        data.address, data.email,
        data.studentNationalId, data.parentName, data.parentNationalId, data.parentEmail,
      ])
    } else {
      const records = existsSync(DATA_FILE) ? JSON.parse(readFileSync(DATA_FILE, 'utf-8')) : []
      records.push(data)
      writeFileSync(DATA_FILE, JSON.stringify(records, null, 2))
    }

    if (process.env.RESEND_API_KEY) {
      try { await sendEmail(process.env.RESEND_API_KEY, data) }
      catch (e) { console.error('Email failed:', e.message) }
    }

    res.json({ success: true, requestNumber: data.requestNumber })
  } catch (error) {
    console.error('Server error:', error)
    res.status(500).json({ error: error.message || 'Internal server error' })
  }
})

app.listen(3001, () => console.log(`\n  API Server running on http://localhost:3001\n`))
