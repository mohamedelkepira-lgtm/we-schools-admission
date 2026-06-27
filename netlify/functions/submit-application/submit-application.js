import { createPrivateKey, sign } from 'crypto'

export async function handler(event) {
  if (event.httpMethod !== 'POST') {
    return { statusCode: 405, body: JSON.stringify({ error: 'Method not allowed' }) }
  }

  try {
    const data = JSON.parse(event.body)

    const sheetId = process.env.GOOGLE_SHEET_ID
    const privateKey = process.env.GOOGLE_PRIVATE_KEY?.replace(/\\n/g, '\n')
    const clientEmail = process.env.GOOGLE_CLIENT_EMAIL

    const headersRow = [
      'رقم الطلب', 'تاريخ التقديم', 'الاسم بالعربية', 'الاسم بالإنجليزية',
      'تاريخ الميلاد', 'الجنس', 'الجنسية', 'الديانة', 'العنوان',
      'رقم الهاتف', 'البريد الإلكتروني', 'رقم بطاقة الطالب',
      'اسم ولي الأمر', 'رقم بطاقة ولي الأمر', 'هاتف ولي الأمر', 'بريد ولي الأمر',
      'المدرسة السابقة',
    ]

    const valuesRow = [
      data.requestNumber, data.submittedAt,
      data.fullNameAr, data.fullNameEn,
      data.dateOfBirth, data.gender,
      data.nationality, data.religion,
      data.address, data.phone, data.email,
      data.studentNationalId,
      data.parentName, data.parentNationalId,
      data.parentPhone, data.parentEmail,
      data.previousSchool,
    ]

    if (sheetId && privateKey && clientEmail) {
      const token = await getGoogleToken(clientEmail, privateKey)
      await appendToGoogleSheet(token, sheetId, headersRow, valuesRow)
    }

    if (process.env.RESEND_API_KEY) {
      try {
        await sendEmail(process.env.RESEND_API_KEY, data)
      } catch (e) {
        console.error('Email failed:', e)
      }
    }

    return {
      statusCode: 200,
      body: JSON.stringify({ success: true, requestNumber: data.requestNumber }),
    }
  } catch (error) {
    console.error('Submit error:', error)
    return {
      statusCode: 500,
      body: JSON.stringify({ error: error.message || 'Internal server error' }),
    }
  }
}

function base64url(str) {
  return Buffer.from(str).toString('base64url')
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

  const key = createPrivateKey(privateKey)
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
  if (!data.access_token) throw new Error('Google Auth failed')
  return data.access_token
}

async function appendToGoogleSheet(token, sheetId, headers, row) {
  const base = `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/Sheet1`
  const opts = { headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' } }

  const check = await fetch(`${base}!A1:Q1`, { headers: opts.headers })
  const existing = await check.json()

  if (!existing.values || existing.values.length === 0) {
    await fetch(`${base}!A1:Q1?valueInputOption=USER_ENTERED`, {
      method: 'PUT', ...opts, body: JSON.stringify({ values: [headers] }),
    })
  }

  await fetch(`${base}!A:Q:append?valueInputOption=USER_ENTERED`, {
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
