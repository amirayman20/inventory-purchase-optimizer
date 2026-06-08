# Inventory Purchase Optimizer

[![Python](https://img.shields.io/badge/Python-3.8%2B-blue?style=flat-square&logo=python)](https://www.python.org/)
[![Pandas](https://img.shields.io/badge/Pandas-Data%20Analysis-green?style=flat-square&logo=pandas)](https://pandas.pydata.org/)
[![Excel](https://img.shields.io/badge/Output-Excel%20Reports-darkgreen?style=flat-square&logo=microsoft-excel)](https://www.microsoft.com/en-us/microsoft-365/excel)
[![License](https://img.shields.io/badge/License-MIT-lightgrey?style=flat-square)](LICENSE)

---

## المشكلة: رصيد خادع

تخيل هذا السيناريو:

- **صنف معين** بتبيع 100 وحدة شهرياً
- **المستودع المركزي** عنده 50 وحدة
- **الفروع** عندهم 200 وحدة إجمالي

على السطح: لديك 250 وحدة ✓

**لكن الحقيقة:**
- 150 من الـ 200 محبوسة في **فروع ضعيفة الأداء**
- هذه الفروع بتبيع 2-3 وحدات فقط شهرياً
- الـ 150 هستفضل هناك 2-3 سنين على الأقل
- بينما المستودع والفروع القوية جوعى من المخزون

**النتيجة:** قرار شراء مبني على رقم وهم — وهذا ممكن يخرج عن الميزانية بـ 40-50%.

---

## الحل: Think Differently

هذا المشروع لا يحل المشكلة بـ "رقم أفضل" — بل بـ **منطق مختلف تماماً**.

**الفكرة الأساسية:** فصل الـ **useful stock** عن الـ **total stock**.

### الخطوات الثلاثة

```
Raw Stock → ABC Classification → Useful Stock → Purchase Decision
```

#### 1. ABC Classification الفروع
تصنيف الفروع حسب مساهمتها في المبيعات:
- **A**: أقوى 40% من فروعك (تمثل 40% من إجمالي المبيعات)
- **B**: الـ 35% التالية (من 40% إلى 75%)
- **C**: الفروع الضعيفة المتبقية (75%+)

#### 2. Useful Stock الفرع
لا كل رصيد محسوب بنفس الوزن:
- **Fروع A**: 100% من الرصيد مفيد ✓
- **Fروع B**: 50% من الرصيد فقط (قليل التحرك)
- **Fروع C**: 0% (محبوس فعلياً)

#### 3. Two Types of Dead Stock
هنا الـ insight الأهم — مش كل رصيد راكد بنفس المشكلة:

| | Slow Item | Branch Leakage |
|---|---|---|
| **التعريف** | صنف ضعيف في كل الفروع | صنف قوي لكن محبوس في فروع ضعيفة |
| **السبب** | مشكلة في الصنف نفسه | مشكلة في التوزيع |
| **القرار** | توقف الشراء / أرجع الستوك | أعد التوزيع بين الفروع |
| **المسؤول** | المشتريات | Supply Chain |

---

## المنطق في العمل

**مثال واقعي:**
- صنف بيبيع 500 وحدة شهرياً
- الفروع فيها 800 وحدة
- المستودع فيه 200 وحدة
- **إجمالي رصيد: 1000 ✓**

لما نطبق الـ algorithm:
- مبيعات الفروع الحقيقية = 500
- فروع A بتبيع 300 (useful stock = 300)
- فروع B بتبيع 150 (useful stock = 75)
- فروع C بتبيع 50 (useful stock = 0)
- **إجمالي useful stock = 375 فقط**

الفرق = 625 وحدة **محبوسة** → لا نشتري حتى نفرغ الـ 625.

---

## المخرجات: ماذا في الـ Excel

الكود يطلع **8 sheets**:

| Sheet | الفائدة |
|---|---|
| **Purchases** | كل صنف + احتياجه + أولويته بألوان |
| **NeedPurchase** | اللي تحتاج شراء بس (مرشح) |
| **Warehouse** | حالة المستودع + الرصيد الراكد |
| **SlowItems** | أصناف ضعيفة في كل مكان |
| **LeakageBranches** | الرصيد المحبوس في فروع C |
| **BranchClassification** | تصنيف كل فرع (A/B/C) |
| **Summary** | الأرقام العامة |

---

## التقنيات المستخدمة

- **Python 3.8+** — الـ core logic
- **Pandas** — معالجة البيانات
- **Excel (XlsxWriter)** — التقارير الملونة

---

## كيفية الاستخدام

### 1. تحضير البيانات

ملفين Excel:
- **ارصدة.xlsx**: ITEM_CODE, NAME_E, BAL
- **مبيعات.xlsx**: ITEM_CODE, ITEM_NAME, STORE_NAME, SALES_QTY, BALANCE, UNIT_COST, SALES_PRICE

### 2. تعديل المسارات

افتح `analysis.py` وحدث:
```python
WAREHOUSE_PATH = r"C:\Users\YourName\Desktop\ارصدة.xlsx"
BRANCHES_PATH  = r"C:\Users\YourName\Desktop\مبيعات.xlsx"
```

### 3. شغّل الكود

```bash
python analysis.py
```

### 4. استقبل التقرير

ملف Excel جديد على الـ Desktop بكل الـ sheets.

---

## الـ Settings (يمكن تعديلها)

```python
# فترة التقرير
START_DATE     = "2026-02-12"
END_DATE       = "2026-05-12"
COVERAGE_DAYS  = 90      # كم يوم نحتاج نغطيها من المخزون
SAFETY_DAYS    = 14      # احتياطي فوق الـ coverage

# ABC thresholds
ABC_A_THRESH   = 40      # أعلى 40% من الفروع
ABC_B_THRESH   = 75      # إلى 75%

# Useful stock weights
WEIGHT_A       = 1.00    # A branches = 100%
WEIGHT_B       = 0.50    # B branches = 50%
WEIGHT_C       = 0.00    # C branches = 0%

# Tuning
FAST_MULTIPLIER = 1.30   # ضرب كمية الأصناف السريعة بـ 1.3
```

---

## Output Colors

- 🟢 **Green (A / FAST)** — سليم / بيع سريع
- 🟡 **Yellow (B / YELLOW)** — احذر / شراء قليل
- 🟠 **Orange (C / SLOW)** — مشكلة / بيع بطيء
- 🔴 **Red (DEAD / RED)** — خطير / لا بيع

---

## أسئلة شائعة

### هل هذا يعني مش نشتري من الفروع C؟
لا — C مفيدة. لكن قبل نشتري، بنتأكد إن الـ useful stock كافي.

### إذا الفرع A ضعيف فجأة؟
الـ thresholds هتتحدث في الـ review التالي. الـ algorithm adaptive.

### ممكن نستخدمه في شركة غير صيدليات؟
أيوه — أي شركة عندها مستودع مركزي وفروع بيع.

---

## اتصل

- **LinkedIn**: [Amir Ayman](https://www.linkedin.com/in/amir-ayman-664513103/)
- **GitHub**: [amirayman20](https://github.com/amirayman20)
- **Email**: [amirayman20@gmail.com](mailto:amirayman20@gmail.com)

---

## License

MIT License — استخدم وعدّل براحتك.

---

**Last updated**: June 2026
