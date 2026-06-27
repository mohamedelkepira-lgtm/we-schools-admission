import { Link } from 'react-router-dom'

function FormIcon() {
  return (
    <svg className="w-8 h-8 mx-auto mb-3 text-[var(--we-blue-light)]" fill="none" viewBox="0 0 24 24" stroke="currentColor">
      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={1.5} d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
    </svg>
  )
}

function ShieldIcon() {
  return (
    <svg className="w-8 h-8 mx-auto mb-3 text-[var(--we-blue-light)]" fill="none" viewBox="0 0 24 24" stroke="currentColor">
      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={1.5} d="M9 12l2 2 4-4m5.618-4.016A11.955 11.955 0 0112 2.944a11.955 11.955 0 01-8.618 3.04A12.02 12.02 0 003 9c0 5.591 3.824 10.29 9 11.622 5.176-1.332 9-6.03 9-11.622 0-1.042-.133-2.052-.382-3.016z" />
    </svg>
  )
}

function MailIcon() {
  return (
    <svg className="w-8 h-8 mx-auto mb-3 text-[var(--we-blue-light)]" fill="none" viewBox="0 0 24 24" stroke="currentColor">
      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={1.5} d="M3 8l7.89 5.26a2 2 0 002.22 0L21 8M5 19h14a2 2 0 002-2V7a2 2 0 00-2-2H5a2 2 0 00-2 2v10a2 2 0 002 2z" />
    </svg>
  )
}

export default function Home() {
  return (
    <div className="max-w-6xl mx-auto px-4 py-12">
      <section className="text-center mb-16 animate-fade-up">
        <div className="inline-flex items-center gap-2 bg-[var(--we-blue)]/5 text-[var(--we-blue)] px-4 py-1.5 rounded-full text-sm font-medium mb-6 border border-[var(--we-blue)]/10">
          العام الدراسي 2025 - 2026
        </div>
        <h1 className="text-4xl md:text-5xl font-bold text-[var(--we-blue)] mb-4 leading-tight">
          مرحباً بك في <span className="text-[var(--we-gold)]">WE Schools</span>
        </h1>
        <p className="text-lg text-gray-600 max-w-2xl mx-auto mb-8 leading-relaxed">
          سجل الآن في مدارس WE للتميز الأكاديمي وبناء مستقبل مشرق لأبنائنا.
        </p>
        <div className="flex flex-col sm:flex-row gap-3 justify-center">
          <Link
            to="/register"
            className="bg-[var(--we-blue)] hover:bg-[var(--we-blue-light)] text-white px-8 py-3.5 rounded-xl text-lg font-semibold transition-all duration-300 shadow-lg hover:shadow-xl hover:-translate-y-0.5"
          >
            ابدأ التسجيل
          </Link>
          <Link
            to="/documents"
            className="border-2 border-gray-300 hover:border-[var(--we-blue)] text-gray-700 hover:text-[var(--we-blue)] px-8 py-3.5 rounded-xl text-lg font-semibold transition-all duration-300 hover:-translate-y-0.5"
          >
            الأوراق المطلوبة
          </Link>
        </div>
      </section>

      <section className="grid md:grid-cols-3 gap-6 mb-16">
        {[
          { icon: FormIcon, title: 'تسجيل سهل', desc: 'نموذج تسجيل مبسط يمكنك من إكمال طلبك في دقائق معدودة' },
          { icon: ShieldIcon, title: 'بيانات آمنة', desc: 'جميع بياناتك مشفرة ومحفوظة بأمان وفق أعلى معايير الخصوصية' },
          { icon: MailIcon, title: 'تأكيد فوري', desc: 'تصلك رسالة تأكيد على بريدك الإلكتروني فور اكتمال التسجيل' },
        ].map((item, i) => {
          const Icon = item.icon
          return (
            <div
              key={item.title}
              className="bg-white p-6 rounded-2xl shadow-sm border border-gray-100 text-center hover:shadow-md hover:-translate-y-1 transition-all duration-300 animate-fade-up"
              style={{ animationDelay: `${0.15 * (i + 1)}s` }}
            >
              <Icon />
              <h3 className="font-bold text-[var(--we-blue)] mb-1">{item.title}</h3>
              <p className="text-gray-500 text-sm leading-relaxed">{item.desc}</p>
            </div>
          )
        })}
      </section>

      <section className="bg-gradient-to-l from-[var(--we-blue)] to-[var(--we-blue-light)] rounded-3xl p-8 md:p-12 text-white text-center animate-fade-up">
        <h2 className="text-2xl md:text-3xl font-bold mb-3">استثمر في مستقبل أبنائك</h2>
        <p className="text-white/80 mb-6 max-w-xl mx-auto leading-relaxed">
          نوفر بيئة تعليمية متكاملة تجمع بين المناهج الأكاديمية العالمية والقيم العربية الأصيلة.
        </p>
        <Link
          to="/register"
          className="inline-block bg-[var(--we-gold)] hover:bg-[var(--we-gold-light)] text-white px-8 py-3 rounded-xl font-semibold transition-all duration-300 hover:shadow-lg hover:-translate-y-0.5"
        >
          سجل الآن
        </Link>
      </section>
    </div>
  )
}
