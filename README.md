# Inventory Purchase Optimizer

<img src="https://raw.githubusercontent.com/amirayman20/inventory-purchase-optimizer/main/assets/inventory_logic_flow.png" width="100%">

---

## Badges

![Python](https://img.shields.io/badge/Python-3.8+-blue?style=for-the-badge&logo=python)
![Pandas](https://img.shields.io/badge/Pandas-Data_Analysis-green?style=for-the-badge&logo=pandas)
![Excel](https://img.shields.io/badge/Excel-Automated_Reports-217346?style=for-the-badge&logo=microsoft-excel)
![Supply Chain](https://img.shields.io/badge/Supply_Chain-Optimization-orange?style=for-the-badge)
![Business Logic](https://img.shields.io/badge/Algorithm-Inventory_Intelligence-purple?style=for-the-badge)

---

## Contact Me

![LinkedIn](https://img.shields.io/badge/LinkedIn-Amir_Ayman-0A66C2?style=for-the-badge&logo=linkedin)
![GitHub](https://img.shields.io/badge/GitHub-amirayman20-181717?style=for-the-badge&logo=github)
![Email](https://img.shields.io/badge/Email-amirayman20@gmail.com-red?style=for-the-badge&logo=gmail)

---

# المشكلة: رصيد خادع (Misleading Stock)

تخيل هذا السيناريو:

- صنف بيبيع **100 وحدة شهرياً**
- المستودع المركزي فيه **50 وحدة**
- الفروع فيها **200 وحدة**

على الورق: عندك **250 وحدة**  
لكن الحقيقة مختلفة…

- 150 وحدة موجودة في **فروع C ضعيفة**
- الفروع دي بتبيع 2–3 وحدات شهرياً
- يعني المخزون محبوس **سنتين أو أكتر**
- بينما الفروع القوية والمستودع محتاجين مخزون

**النتيجة:** قرار شراء مبني على رقم وهمي → ممكن يضرب الميزانية بـ **40–50%**.

---

# الحل: Think Differently

المشروع دا مش بيحسّن رقم…  
هو بيغيّر **طريقة التفكير** في المخزون.

## الفكرة الأساسية
**نفصل الـ Useful Stock عن الـ Total Stock.**

### Pipeline

```
Raw Stock → Branch ABC Classification → Useful Stock → Purchase Decision
```

---

# 1) Branch ABC Classification

تصنيف الفروع حسب مساهمتها في المبيعات:

- **A** = أعلى 40% من الفروع  
- **B** = من 40% إلى 75%  
- **C** = أقل 25% (ضعيفة جداً)

---

# 2) Useful Stock Logic

مش كل رصيد له نفس القيمة:

| Branch Type | Useful % |
|-------------|----------|
| A | 100% |
| B | 50% |
| C | 0% |

---

# 3) Two Types of Dead Stock

| | Slow Item | Branch Leakage |
|---|---|---|
| التعريف | صنف ضعيف في كل الفروع | صنف قوي لكن محبوس في فروع C |
| السبب | مشكلة في الصنف | مشكلة توزيع |
| القرار | وقف شراء | إعادة توزيع |
| المسؤول | المشتريات | Supply Chain |

---

# مثال عملي

- مبيعات شهرية: **500**
- رصيد الفروع: **800**
- رصيد المستودع: **200**
- الإجمالي: **1000**

بعد تطبيق الـ Algorithm:

- Useful Stock = **375**
- الفرق = **625 وحدة محبوسة**

**القرار:** لا شراء حتى يتم تفريغ الـ 625.

---

# Excel Output (8 Sheets)

| Sheet | Description |
|-------|-------------|
| Purchases | الاحتياج النهائي + الأولويات |
| NeedPurchase | الأصناف المرشحة للشراء |
| Warehouse | حالة المستودع |
| SlowItems | الأصناف الضعيفة |
| LeakageBranches | الرصيد المحبوس |
| BranchClassification | تصنيف الفروع |
| Summary | ملخص عام |

---

# Tech Stack

- Python 3.8+
- Pandas
- XlsxWriter
- Excel Automation

---

# Usage

## 1) Prepare Data

- **ارصدة.xlsx** → ITEM_CODE, NAME_E, BAL  
- **مبيعات.xlsx** → ITEM_CODE, STORE_NAME, SALES_QTY, BALANCE…

## 2) Update Paths

```python
WAREHOUSE_PATH = r"C:\Users\YourName\Desktop\ارصدة.xlsx"
BRANCHES_PATH  = r"C:\Users\YourName\Desktop\مبيعات.xlsx"
```

## 3) Run

```bash
python analysis.py
```

## 4) Output

Excel file with all sheets.

---

# Settings

```python
START_DATE     = "2026-02-12"
END_DATE       = "2026-05-12"
COVERAGE_DAYS  = 90
SAFETY_DAYS    = 14

ABC_A_THRESH   = 40
ABC_B_THRESH   = 75

WEIGHT_A       = 1.00
WEIGHT_B       = 0.50
WEIGHT_C       = 0.00

FAST_MULTIPLIER = 1.30
```

---

# Colors

- 🟢 FAST  
- 🟡 Medium  
- 🟠 Slow  
- 🔴 Dead  

---

# FAQ

### هل نوقف شراء فروع C؟
لا — لكن لازم نتأكد إن useful stock كافي.

### هل ينفع لشركات غير صيدليات؟
نعم — أي شركة عندها مستودع + فروع.

---

**Last Updated:** June 2026
