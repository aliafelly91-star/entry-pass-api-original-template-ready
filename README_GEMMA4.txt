Gemma 4 - زر G
================

الملفات:
1) main.py
   استبدل به main.py في سيرفر Render الحالي الخاص بـ:
   https://entry-pass-api-original-template-ready.onrender.com

2) stage_two_review_screen.dart
   استبدل به ملف الشاشة الحالي.

3) gemma4_passport_reader.dart
   ضعه بنفس مجلد stage_two_review_screen.dart وبقية ملفات readers.

إعداد Render:
- افتح خدمة Render الحالية.
- Environment / Environment Variables.
- أضف:
  GEMMA_API_KEY = مفتاح Google AI Studio
- Save / Deploy latest commit.

اختبار السيرفر:
GET /health
يجب أن يظهر:
"gemma_configured": true
"gemma_model": "gemma-4-26b-a4b-it"

بعدها افتح التطبيق وستجد زر قراءة جديد باسم G لكل جواز.
الزر يرسل الصورة إلى Render، وRender يستدعي Gemma 4، ثم يرجع البيانات للتطبيق.
المفتاح لا يوجد داخل Flutter.
