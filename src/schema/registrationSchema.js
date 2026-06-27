import { z } from 'zod'

export const registrationSchema = z.object({
  fullNameAr: z.string().min(3, 'الاسم بالعربية مطلوب'),
  fullNameEn: z.string().min(3, 'الاسم بالإنجليزية مطلوب'),
  dateOfBirth: z.string().min(1, 'تاريخ الميلاد مطلوب'),
  gender: z.enum(['ذكر', 'أنثى'], { required_error: 'الجنس مطلوب' }),
  nationality: z.string().min(2, 'الجنسية مطلوبة'),
  religion: z.string().min(2, 'الديانة مطلوبة'),
  address: z.string().min(5, 'العنوان مطلوب'),
  phone: z.string().min(10, 'رقم الهاتف مطلوب'),
  email: z.string().email('البريد الإلكتروني غير صحيح'),
  studentNationalId: z.string().regex(/^\d{14}$/, 'رقم البطاقة يجب أن يكون 14 رقمًا'),
  parentName: z.string().min(3, 'اسم ولي الأمر مطلوب'),
  parentNationalId: z.string().regex(/^\d{14}$/, 'رقم البطاقة يجب أن يكون 14 رقمًا'),
  parentPhone: z.string().min(10, 'رقم هاتف ولي الأمر مطلوب'),
  parentEmail: z.string().email('البريد الإلكتروني لولي الأمر غير صحيح'),
  previousSchool: z.string().min(2, 'اسم المدرسة السابقة مطلوب'),
})
