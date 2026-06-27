import { useState, useMemo } from 'react'
import { useForm } from 'react-hook-form'
import { zodResolver } from '@hookform/resolvers/zod'
import { useNavigate } from 'react-router-dom'
import { registrationSchema } from '../schema/registrationSchema'
import { useAutoSave, getSavedDraft, clearSavedDraft } from '../hooks/useAutoSave'
import { generateRequestNumber } from '../utils/generateRequestNumber'
import ProgressBar from '../components/ProgressBar'

export default function Register() {
  const navigate = useNavigate()
  const [submitting, setSubmitting] = useState(false)
  const [submitError, setSubmitError] = useState('')

  const {
    register,
    handleSubmit,
    watch,
    formState: { errors },
  } = useForm({
    resolver: zodResolver(registrationSchema),
    defaultValues: getSavedDraft() || {},
  })

  const watchedData = watch()
  useAutoSave(watchedData)

  const totalFields = useMemo(() => {
    return Object.keys(registrationSchema.shape).length
  }, [])

  const filledFields = useMemo(() => {
    let count = 0
    for (const key of Object.keys(registrationSchema.shape)) {
      const val = watchedData[key]
      if (val && String(val).trim().length > 0) count++
    }
    return count
  }, [watchedData])

  const progress = Math.round((filledFields / totalFields) * 100)

  async function onSubmit(data) {
    setSubmitting(true)
    setSubmitError('')

    try {
      const requestNumber = generateRequestNumber()
      const payload = { ...data, requestNumber, submittedAt: new Date().toISOString() }

      const res = await fetch('/.netlify/functions/submit-application', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload),
      })

      const result = await res.json()

      if (!res.ok) {
        throw new Error(result.error || 'حدث خطأ أثناء حفظ البيانات')
      }

      clearSavedDraft()
      sessionStorage.setItem('we-last-request', JSON.stringify({ requestNumber, email: data.email }))
      navigate('/success')
    } catch (err) {
      setSubmitError(err.message)
    } finally {
      setSubmitting(false)
    }
  }

  const inputClass = (fieldError) =>
    `w-full px-4 py-3 rounded-xl border transition-all duration-200 text-sm outline-none ${
      fieldError
        ? 'border-red-400 bg-red-50 focus:border-red-500'
        : 'border-gray-300 bg-white focus:border-[var(--we-blue)] focus:ring-1 focus:ring-[var(--we-blue)]/20'
    }`

  const labelClass = 'block text-sm font-medium text-gray-700 mb-1.5'

  return (
    <div className="max-w-3xl mx-auto px-4 py-12 animate-fade-up">
      <h1 className="text-3xl font-bold text-[var(--we-blue)] mb-2">نموذج التسجيل</h1>
      <p className="text-gray-500 mb-6">يرجى ملء جميع الحقول الإجبارية</p>

      <div className="mb-8">
        <ProgressBar percentage={progress} />
      </div>

      <form onSubmit={handleSubmit(onSubmit)} className="space-y-8" noValidate>
        <section className="bg-white p-6 rounded-2xl shadow-sm border border-gray-100 transition-shadow duration-300 hover:shadow-md">
          <h2 className="text-lg font-bold text-[var(--we-blue)] mb-5 pb-3 border-b border-gray-100">
            بيانات الطالب
          </h2>
          <div className="grid md:grid-cols-2 gap-4">
            <div>
              <label className={labelClass}>الاسم بالعربية</label>
              <input {...register('fullNameAr')} className={inputClass(errors.fullNameAr)} placeholder="الاسم الرباعي بالعربية" />
              {errors.fullNameAr && <p className="text-red-500 text-xs mt-1">{errors.fullNameAr.message}</p>}
            </div>
            <div>
              <label className={labelClass}>الاسم بالإنجليزية</label>
              <input {...register('fullNameEn')} className={inputClass(errors.fullNameEn)} placeholder="الاسم بالإنجليزية" />
              {errors.fullNameEn && <p className="text-red-500 text-xs mt-1">{errors.fullNameEn.message}</p>}
            </div>
            <div>
              <label className={labelClass}>تاريخ الميلاد</label>
              <input type="date" {...register('dateOfBirth')} className={inputClass(errors.dateOfBirth)} />
              {errors.dateOfBirth && <p className="text-red-500 text-xs mt-1">{errors.dateOfBirth.message}</p>}
            </div>
            <div>
              <label className={labelClass}>الجنس</label>
              <select {...register('gender')} className={inputClass(errors.gender)}>
                <option value="">اختر الجنس</option>
                <option value="ذكر">ذكر</option>
                <option value="أنثى">أنثى</option>
              </select>
              {errors.gender && <p className="text-red-500 text-xs mt-1">{errors.gender.message}</p>}
            </div>
            <div>
              <label className={labelClass}>الجنسية</label>
              <input {...register('nationality')} className={inputClass(errors.nationality)} placeholder="مصري" />
              {errors.nationality && <p className="text-red-500 text-xs mt-1">{errors.nationality.message}</p>}
            </div>
            <div>
              <label className={labelClass}>الديانة</label>
              <input {...register('religion')} className={inputClass(errors.religion)} placeholder="مسلم / مسيحي" />
              {errors.religion && <p className="text-red-500 text-xs mt-1">{errors.religion.message}</p>}
            </div>
            <div className="md:col-span-2">
              <label className={labelClass}>العنوان</label>
              <input {...register('address')} className={inputClass(errors.address)} placeholder="العنوان بالتفصيل" />
              {errors.address && <p className="text-red-500 text-xs mt-1">{errors.address.message}</p>}
            </div>
            <div>
              <label className={labelClass}>رقم الهاتف</label>
              <input {...register('phone')} className={inputClass(errors.phone)} placeholder="01234567890" dir="ltr" />
              {errors.phone && <p className="text-red-500 text-xs mt-1">{errors.phone.message}</p>}
            </div>
            <div>
              <label className={labelClass}>البريد الإلكتروني</label>
              <input type="email" {...register('email')} className={inputClass(errors.email)} placeholder="student@example.com" dir="ltr" />
              {errors.email && <p className="text-red-500 text-xs mt-1">{errors.email.message}</p>}
            </div>
            <div>
              <label className={labelClass}>رقم البطاقة</label>
              <input {...register('studentNationalId')} className={inputClass(errors.studentNationalId)} placeholder="14 رقم" maxLength={14} dir="ltr" inputMode="numeric" />
              {errors.studentNationalId && <p className="text-red-500 text-xs mt-1">{errors.studentNationalId.message}</p>}
            </div>
          </div>
        </section>

        <section className="bg-white p-6 rounded-2xl shadow-sm border border-gray-100 transition-shadow duration-300 hover:shadow-md">
          <h2 className="text-lg font-bold text-[var(--we-blue)] mb-5 pb-3 border-b border-gray-100">
            بيانات ولي الأمر
          </h2>
          <div className="grid md:grid-cols-2 gap-4">
            <div>
              <label className={labelClass}>اسم ولي الأمر</label>
              <input {...register('parentName')} className={inputClass(errors.parentName)} placeholder="الاسم الرباعي" />
              {errors.parentName && <p className="text-red-500 text-xs mt-1">{errors.parentName.message}</p>}
            </div>
            <div>
              <label className={labelClass}>رقم البطاقة</label>
              <input {...register('parentNationalId')} className={inputClass(errors.parentNationalId)} placeholder="14 رقم" maxLength={14} dir="ltr" inputMode="numeric" />
              {errors.parentNationalId && <p className="text-red-500 text-xs mt-1">{errors.parentNationalId.message}</p>}
            </div>
            <div>
              <label className={labelClass}>رقم الهاتف</label>
              <input {...register('parentPhone')} className={inputClass(errors.parentPhone)} placeholder="01234567890" dir="ltr" />
              {errors.parentPhone && <p className="text-red-500 text-xs mt-1">{errors.parentPhone.message}</p>}
            </div>
            <div className="md:col-span-2">
              <label className={labelClass}>البريد الإلكتروني</label>
              <input type="email" {...register('parentEmail')} className={inputClass(errors.parentEmail)} placeholder="parent@example.com" dir="ltr" />
              {errors.parentEmail && <p className="text-red-500 text-xs mt-1">{errors.parentEmail.message}</p>}
            </div>
          </div>
        </section>

        <section className="bg-white p-6 rounded-2xl shadow-sm border border-gray-100 transition-shadow duration-300 hover:shadow-md">
          <h2 className="text-lg font-bold text-[var(--we-blue)] mb-5 pb-3 border-b border-gray-100">
            البيانات الدراسية
          </h2>
          <div className="grid md:grid-cols-1 gap-4">
            <div>
              <label className={labelClass}>المدرسة السابقة</label>
              <input {...register('previousSchool')} className={inputClass(errors.previousSchool)} placeholder="اسم المدرسة السابقة" />
              {errors.previousSchool && <p className="text-red-500 text-xs mt-1">{errors.previousSchool.message}</p>}
            </div>
          </div>
        </section>

        {submitError && (
          <div className="bg-red-50 border border-red-200 text-red-700 px-4 py-3 rounded-xl text-sm animate-fade-in">
            {submitError}
          </div>
        )}

        <button
          type="submit"
          disabled={submitting}
          className="w-full bg-[var(--we-blue)] hover:bg-[var(--we-blue-light)] disabled:bg-gray-400 disabled:cursor-not-allowed text-white py-3.5 rounded-xl font-semibold text-lg transition-all duration-300 shadow-lg hover:shadow-xl hover:-translate-y-0.5 disabled:transform-none"
        >
          {submitting ? 'جاري إرسال البيانات...' : 'إرسال الطلب'}
        </button>
      </form>
    </div>
  )
}
