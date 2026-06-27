import { Link } from 'react-router-dom'

const documents = [
  { name: 'شهادة ميلاد الطالب', desc: 'صورة واضحة من شهادة الميلاد', required: true },
  { name: 'صورة شخصية للطالب', desc: 'صورة شخصية حديثة بخلفية بيضاء', required: true },
  { name: 'شهادة التخرج من العام السابق', desc: 'شهادة نجاح من العام الدراسي السابق', required: true },
  { name: 'بيان نجاح آخر سنة دراسية', desc: 'كشف بدرجات آخر سنة دراسية', required: true },
  { name: 'صورة بطاقة ولي الأمر', desc: 'صورة من بطاقة الرقم القومي لولي الأمر', required: true },
  { name: 'شهادة تحليل نوع', desc: 'شهادة تحليل نوع للطالب', required: false },
  { name: 'الملف الدراسي', desc: 'ملف الطالب الدراسي من المدرسة السابقة', required: false },
  { name: 'إيصال سداد المصروفات', desc: 'إيصال سداد مصروفات التقديم', required: true },
]

export default function Documents() {
  return (
    <div className="max-w-4xl mx-auto px-4 py-12">
      <div className="animate-fade-up">
        <h1 className="text-3xl font-bold text-[var(--we-blue)] mb-2">الأوراق المطلوبة</h1>
        <p className="text-gray-500 mb-8">
          يرجى تجهيز الأوراق التالية قبل البدء في عملية التسجيل
        </p>
      </div>

      <div className="bg-white rounded-2xl shadow-sm border border-gray-100 overflow-hidden">
        {documents.map((doc, i) => (
          <div
            key={doc.name}
            className={`flex items-center gap-4 p-4 hover:bg-gray-50 transition-colors duration-200 animate-fade-up ${
              i < documents.length - 1 ? 'border-b border-gray-100' : ''
            }`}
            style={{ animationDelay: `${0.08 * (i + 1)}s` }}
          >
            <div className={`w-2.5 h-2.5 rounded-full flex-shrink-0 ${doc.required ? 'bg-[var(--we-gold)]' : 'bg-gray-300'}`} />
            <div className="flex-1 min-w-0">
              <div className="font-medium text-gray-900">{doc.name}</div>
              <div className="text-sm text-gray-500">{doc.desc}</div>
            </div>
            <span className={`text-xs font-medium px-2.5 py-1 rounded-full flex-shrink-0 ${
              doc.required
                ? 'bg-amber-50 text-[var(--we-gold)] border border-amber-200'
                : 'bg-gray-100 text-gray-500'
            }`}>
              {doc.required ? 'إجباري' : 'اختياري'}
            </span>
          </div>
        ))}
      </div>

      <div className="mt-8 text-center animate-fade-in">
        <Link
          to="/register"
          className="inline-block bg-[var(--we-blue)] hover:bg-[var(--we-blue-light)] text-white px-8 py-3 rounded-xl font-semibold transition-all duration-300 hover:shadow-lg hover:-translate-y-0.5"
        >
          انتقل إلى التسجيل
        </Link>
      </div>
    </div>
  )
}
