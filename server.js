import 'dotenv/config'
import express from 'express'
import { google } from 'googleapis'
import { Resend } from 'resend'
import { readFileSync, writeFileSync, existsSync } from 'fs'
import { join, dirname } from 'path'
import { fileURLToPath } from 'url'

const __dirname = dirname(fileURLToPath(import.meta.url))
const DATA_FILE = join(__dirname, '_local_data.json')

const app = express()
app.use(express.json())

function generateRequestNumber() {
  const year = new Date().getFullYear()
  const random = Math.floor(1000 + Math.random() * 9000)
  return `WE-${year}-${random}`
}

async function saveToGoogleSheets(data, sheetId, privateKey, clientEmail) {
  const auth = new google.auth.GoogleAuth({
    credentials: {
      type: 'service_account',
      private_key: privateKey.replace(/\\n/g, '\n'),
      client_email: clientEmail,
    },
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  })

  const sheets = google.sheets({ version: 'v4', auth })

  const headers = [
    'رقم الطلب', 'تاريخ التقديم', 'الاسم بالعربية', 'الاسم بالإنجليزية',
    'تاريخ الميلاد', 'الجنس', 'الجنسية', 'الديانة', 'العنوان',
    'رقم الهاتف', 'البريد الإلكتروني', 'رقم بطاقة الطالب',
    'اسم ولي الأمر', 'رقم بطاقة ولي الأمر', 'هاتف ولي الأمر', 'بريد ولي الأمر',
    'المدرسة السابقة',
  ]

  const existing = await sheets.spreadsheets.values.get({
    spreadsheetId: sheetId,
    range: 'Sheet1!A1:Q1',
  })

  if (!existing.data.values || existing.data.values.length === 0) {
    await sheets.spreadsheets.values.update({
      spreadsheetId: sheetId,
      range: 'Sheet1!A1:Q1',
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: [headers] },
    })
  }

  await sheets.spreadsheets.values.append({
    spreadsheetId: sheetId,
    range: 'Sheet1!A:Q',
    valueInputOption: 'USER_ENTERED',
    requestBody: {
      values: [[
        data.requestNumber, data.submittedAt,
        data.fullNameAr, data.fullNameEn,
        data.dateOfBirth, data.gender,
        data.nationality, data.religion,
        data.address, data.phone, data.email,
        data.studentNationalId,
        data.parentName, data.parentNationalId,
        data.parentPhone, data.parentEmail,
        data.previousSchool,
      ]],
    },
  })
}

async function sendConfirmationEmail(data, apiKey) {
  const resend = new Resend(apiKey)
  await resend.emails.send({
    from: process.env.EMAIL_FROM || 'WE Schools <noreply@weschools.com>',
    to: [data.email],
    subject: `تأكيد استلام طلب التسجيل - ${data.requestNumber}`,
    html: `
      <div dir="rtl" style="font-family: 'Cairo', Arial, sans-serif; max-width: 600px; margin: 0 auto;">
        <div style="background: #1e3a5f; padding: 30px; text-align: center; border-radius: 16px 16px 0 0;">
          <h1 style="color: white; margin: 0; font-size: 24px;">WE Schools</h1>
        </div>
        <div style="background: #f9fafb; padding: 30px; border-radius: 0 0 16px 16px;">
          <h2 style="color: #1f2937; margin-top: 0;">تم استلام طلبك بنجاح</h2>
          <p style="color: #6b7280; line-height: 1.8;">
            شكراً لتسجيلك في WE Schools. تم استلام طلب التقديم بنجاح.
          </p>
          <div style="background: #dbeafe; border: 1px solid #bfdbfe; border-radius: 12px; padding: 20px; text-align: center; margin: 20px 0;">
            <div style="color: #2563eb; font-size: 14px; margin-bottom: 4px;">رقم الطلب</div>
            <div style="color: #1e40af; font-size: 28px; font-weight: bold; letter-spacing: 2px;" dir="ltr">
              ${data.requestNumber}
            </div>
          </div>
          <p style="color: #6b7280; line-height: 1.8;">
            سنتواصل معك قريباً لإكمال إجراءات القبول. يرجى الاحتفاظ برقم الطلب لمتابعة حالة طلبك.
          </p>
          <div style="border-top: 1px solid #e5e7eb; padding-top: 16px; margin-top: 24px; text-align: center;">
            <p style="color: #9ca3af; font-size: 12px;">
              WE Schools &copy; ${new Date().getFullYear()} - جميع الحقوق محفوظة
            </p>
          </div>
        </div>
      </div>
    `,
  })
}

app.post('/.netlify/functions/submit-application', async (req, res) => {
  try {
    const data = req.body
    data.requestNumber = generateRequestNumber()
    data.submittedAt = new Date().toISOString()

    const hasGoogleConfig = process.env.GOOGLE_SHEET_ID && process.env.GOOGLE_PRIVATE_KEY && process.env.GOOGLE_CLIENT_EMAIL
    const hasResendKey = process.env.RESEND_API_KEY

    if (hasGoogleConfig) {
      await saveToGoogleSheets(data, process.env.GOOGLE_SHEET_ID, process.env.GOOGLE_PRIVATE_KEY, process.env.GOOGLE_CLIENT_EMAIL)
    } else {
      const records = existsSync(DATA_FILE) ? JSON.parse(readFileSync(DATA_FILE, 'utf-8')) : []
      records.push(data)
      writeFileSync(DATA_FILE, JSON.stringify(records, null, 2))
      console.log('[DEV] Data saved to _local_data.json')
    }

    if (hasResendKey) {
      try {
        await sendConfirmationEmail(data, process.env.RESEND_API_KEY)
      } catch (emailErr) {
        console.error('[DEV] Email send failed (expected if no API key):', emailErr.message)
      }
    } else {
      console.log('[DEV] Email not sent — no RESEND_API_KEY set. Data would be:', data.email)
    }

    res.json({ success: true, requestNumber: data.requestNumber })
  } catch (error) {
    console.error('Server error:', error)
    res.status(500).json({ error: error.message || 'Internal server error' })
  }
})

const PORT = 3001
app.listen(PORT, () => {
  console.log(`\n  ✓ API Server running on http://localhost:${PORT}\n`)
})
