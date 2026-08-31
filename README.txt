سيرفر سمة الدخول - نسخة القالب الأصلي

القالب المستخدم: ملف Word الأصلي المرفوع من المستخدم، مع الحفاظ على الشعار والتوقيع والتنسيق.

Render Build Command:
pip install -r requirements.txt

Render Start Command:
uvicorn main:app --host 0.0.0.0 --port $PORT

POST /fill-entry-pass
الحقول:
nationality
count
first_name
last_name
entry_port
hotel
arrival_date
marketing_company
telegram_destination
document_date

إذا document_date فارغ يوضع تاريخ اليوم تلقائياً.
