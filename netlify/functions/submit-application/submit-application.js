import { google } from 'googleapis'
import { Resend } from 'resend'

const resend = new Resend(process.env.RESEND_API_KEY)

export async function handler(event) {
  if (event.httpMethod !== 'POST') {
    return { statusCode: 405, body: JSON.stringify({ error: 'Method not allowed' }) }
  }

  try {
    const data = JSON.parse(event.body)

    const sheetId = process.env.GOOGLE_SHEET_ID
    const privateKey = process.env.GOOGLE_PRIVATE_KEY?.replace(/\\n/g, '\n')
    const clientEmail = process.env.GOOGLE_CLIENT_EMAIL

    if (!sheetId || !privateKey || !clientEmail) {
      throw new Error('Missing Google Sheets configuration')
    }

    const auth = new google.auth.GoogleAuth({
      credentials: {
        type: 'service_account',
        private_key: privateKey,
        client_email: clientEmail,
      },
      scopes: ['https://www.googleapis.com/auth/spreadsheets'],
    })

    const sheets = google.sheets({ version: 'v4', auth })

    const headers = [
      'رقم الطلب',
      'تاريخ التقديم',
      'الاسم بالعربية',
      'الاسم بالإنجليزية',
      'تاريخ الميلاد',
      'الجنس',
      'الجنسية',
      'الديانة',
      'العنوان',
      'رقم الهاتف',
      'البريد الإلكتروني',
      'رقم بطاقة الطالب',
      'اسم ولي الأمر',
      'رقم بطاقة ولي الأمر',
      'هاتف ولي الأمر',
      'بريد ولي الأمر',
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

    const values = [[
      data.requestNumber,
      data.submittedAt,
      data.fullNameAr,
      data.fullNameEn,
      data.dateOfBirth,
      data.gender,
      data.nationality,
      data.religion,
      data.address,
      data.phone,
      data.email,
      data.studentNationalId,
      data.parentName,
      data.parentNationalId,
      data.parentPhone,
      data.parentEmail,
      data.previousSchool,
    ]]

    await sheets.spreadsheets.values.append({
      spreadsheetId: sheetId,
      range: 'Sheet1!A:Q',
      valueInputOption: 'USER_ENTERED',
      requestBody: { values },
    })

    if (process.env.RESEND_API_KEY) {
      try {
        await resend.emails.send({
          from: process.env.EMAIL_FROM || 'WE Schools <noreply@weschools.com>',
          to: [data.email],
          subject: `تأكيد استلام طلب التسجيل - ${data.requestNumber}`,
          html: `
            <div dir="rtl" style="font-family: 'Cairo', Arial, sans-serif; max-width: 600px; margin: 0 auto;">
              <div style="background: #1a56db; padding: 30px; text-align: center; border-radius: 16px 16px 0 0;">
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
      } catch (emailErr) {
        console.error('Email send failed:', emailErr)
      }
    }

    return {
      statusCode: 200,
      body: JSON.stringify({
        success: true,
        requestNumber: data.requestNumber,
      }),
    }
  } catch (error) {
    console.error('Submit error:', error)
    return {
      statusCode: 500,
      body: JSON.stringify({ error: error.message || 'Internal server error' }),
    }
  }
}
