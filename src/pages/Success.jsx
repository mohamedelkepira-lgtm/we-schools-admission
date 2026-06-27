import { useEffect, useState } from 'react'
import { Link } from 'react-router-dom'

export default function Success() {
  const [request, setRequest] = useState(null)

  useEffect(() => {
    const stored = sessionStorage.getItem('we-last-request')
    if (stored) {
      setRequest(JSON.parse(stored))
    }
  }, [])

  return (
    <div className="max-w-lg mx-auto px-4 py-20 text-center animate-scale-in">
      <div className="bg-white rounded-3xl shadow-sm border border-gray-100 p-8">
        <div className="w-16 h-16 bg-green-50 rounded-full flex items-center justify-center mx-auto mb-4 border border-green-100">
          <svg className="w-8 h-8 text-green-600" fill="none" viewBox="0 0 24 24" stroke="currentColor">
            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M5 13l4 4L19 7" />
          </svg>
        </div>
        <h1 className="text-2xl font-bold text-[var(--we-blue)] mb-2">تم استلام طلبك بنجاح</h1>
        <p className="text-gray-500 mb-6 leading-relaxed">
          شكراً لتسجيلك في WE Schools. سنتواصل معك قريباً.
        </p>

        {request && (
          <div className="bg-blue-50 border border-blue-200 rounded-xl p-4 mb-6 animate-fade-in">
            <div className="text-sm text-[var(--we-blue)] mb-1">رقم الطلب</div>
            <div className="text-2xl font-bold text-[var(--we-blue)] font-mono tracking-wider" dir="ltr">
              {request.requestNumber}
            </div>
          </div>
        )}

        {request?.email && (
          <p className="text-sm text-gray-500 mb-6">
            تم إرسال رسالة تأكيد إلى <span className="font-medium text-gray-700" dir="ltr">{request.email}</span>
          </p>
        )}

        <Link
          to="/"
          className="inline-block bg-[var(--we-blue)] hover:bg-[var(--we-blue-light)] text-white px-8 py-3 rounded-xl font-semibold transition-all duration-300 hover:shadow-lg hover:-translate-y-0.5"
        >
          العودة للرئيسية
        </Link>
      </div>
    </div>
  )
}
