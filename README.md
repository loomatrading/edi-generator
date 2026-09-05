# CUSCAR EDI / TXT Generator

تطبيق بلغة **Python** مزود بواجهة رسومية (GUI) لقراءة بيانات الشحنات والحاويات من ملفات **Excel** وتوليد ملفات المانيفست والجمارك القياسية **EDIFACT CUSCAR** وتصديرها بصيغة `.txt` أو `.edi` للتكامل مع الموانئ والهيئات الجمركية (مثل ميناء السخنة ومنصة نافذة Nafeza).

---

## 🌟 المميزات
- **واجهة رسومية (GUI):** مريحة وسهلة الاستخدام بدون الحاجة للتعامل مع سطر الأوامر.
- **دعم كافة صيغ الإكسيل:** يتوافق مع ملفات `.xlsx` و `.xls`.
- **معاينة سريعة:** عرض جدول للبيانات المقروءة قبل عملية الحفظ.
- **التوافق القياسي:** يولد الهيكل الرسمي القياسي لـ CUSCAR (`UNB`, `UNH`, `BGM`, `EQD`, `CNT`, `UNT`, `UNZ`).

---

## 📁 هيكل المستودع (Project Structure)
```text
cuscar-edi-generator/
│
├── main.py              # الكود البرمجي الواجهة والتشغيل
├── requirements.txt     # المكتبات المطلوبة
├── .gitignore           # ملف تجاهل الملفات المؤقتة
├── README.md            # دليل الاستخدام
└── samples/             # مجلد العينات التجريبية
```

---

## 🛠️ التثبيت والتشغيل

### 1. استคลون المستودع (Clone):
```bash
git clone https://github.com/YOUR-USERNAME/cuscar-edi-generator.git
cd cuscar-edi-generator
```

### 2. تثبيت المكتبات المطلوبة:
```bash
pip install -r requirements.txt
```

### 3. تشغيل البرنامج:
```bash
python main.py
```

---

## 💡 تحويل البرنامج إلى ملف `.exe` يعمل بدون Python
يمكنك تحويل البرنامج إلى تطبيق كمبيوتر تنفيذي بضغطة زر باستخدام `pyinstaller`:
```bash
pip install pyinstaller
pyinstaller --noconsole --onefile main.py
```
سيكون الملف التنفيذي جاهزاً داخل مجلد `dist/main.exe`.
