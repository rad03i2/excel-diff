# Excel Diff

A small, deterministic command-line and Python toolkit for comparing `.xlsx`, `.xlsm`, and `.csv` spreadsheets cell by cell. It is designed for local audits, data-pipeline checks, review workflows, and CI where a machine-readable spreadsheet diff is more useful than opening two workbooks manually.

## Why it exists

Spreadsheet changes are often difficult to review reliably. Excel Diff turns workbook differences into explicit sheet/cell changes and uses meaningful exit codes, without uploading files anywhere.

## Features

- Compare XLSX/XLSM workbooks across all common sheets.
- Compare UTF-8 CSV files as a single logical sheet.
- Detect added/removed sheets and added/removed/changed cells.
- Compare cached formula results by default, or formula expressions with `--formulas`.
- Optional case-insensitive string comparison.
- Human-readable text and stable JSON output.
- Write reports to disk for CI artifacts or audit records.
- Exit codes: `0` identical, `1` differences, `2` invalid input/error.
- Offline operation; no telemetry or network requests.
- Python API for integration into other tools.

## Preview

```text
Excel Diff: before.xlsx -> after.xlsx
Result: differences found
Inventory!B2: 2 -> 5
Changed cells: 1
```

For screenshots, run the sample below and capture the terminal output. The project intentionally has no GUI.

## Requirements & installation

- Python 3.10+
- `openpyxl` 3.1+

```bash
git clone https://github.com/rad03i2/excel-diff.git
cd excel-diff
python -m venv .venv
# Windows: .venv\Scripts\activate
# macOS/Linux: source .venv/bin/activate
pip install -e .
```

## Usage

```bash
excel-diff before.xlsx after.xlsx
excel-diff before.xlsx after.xlsx --format json
excel-diff before.xlsx after.xlsx --format json --output diff.json
excel-diff old.csv new.csv --ignore-case
excel-diff before.xlsx after.xlsx --formulas
python -m excel_diff before.xlsx after.xlsx
```

Generate reproducible sample workbooks:

```bash
python examples/make_samples.py
excel-diff examples/generated/before.xlsx examples/generated/after.xlsx
```

Python API:

```python
from excel_diff import compare

report = compare("before.xlsx", "after.xlsx")
print(report.identical)
for change in report.changes:
    print(change.sheet, change.cell, change.before, change.after)
```

## Configuration

No environment variables, accounts, keys, or configuration files are required. CLI flags control behavior. `--limit` only limits text display; JSON output contains every detected cell change.

## Project structure

```text
src/excel_diff/       comparison engine and CLI
tests/                core and CLI tests
examples/             reproducible sample generator
.github/workflows/    cross-platform CI
```

## Testing

```bash
pip install -e . pytest
python -m compileall -q src tests
pytest -q
excel-diff --version
```

CI runs the same project on Ubuntu, Windows, and macOS with Python 3.10, 3.12, and 3.13.

## Limitations

- Legacy binary `.xls` is not supported.
- Formatting, comments, charts, images, macros, merged-cell geometry, and workbook metadata are not compared.
- Formula mode compares formula text, not semantic equivalence.
- Default formula-result comparison depends on cached values stored in the workbook; openpyxl does not calculate formulas.
- CSV comparison treats values as strings; XLSX preserves Excel value types.
- Large workbooks are read in openpyxl read-only mode, but the non-empty cell maps and final diff remain in memory.

## Security & privacy

Files are processed locally. The tool makes no network requests. Reports may reproduce cell values, so protect generated JSON/text reports just as you protect the source workbook. See [SECURITY.md](SECURITY.md).

## Optional roadmap

Possible future additions include style-aware comparison, selected-sheet/range filters, and HTML reports. They are not required for the current cell-data comparison workflow.

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md). Please use synthetic spreadsheets in tests and issues rather than confidential real-world files.

## License

MIT — see [LICENSE](LICENSE).

## Author

**Radwan Abdulhadi Ahmed**  
**رضوان عبدالهادي أحمد**  
GitHub: [@rad03i2](https://github.com/rad03i2)

---

# Excel Diff — العربية

أداة محلية صغيرة وموثوقة لسطر الأوامر وPython لمقارنة ملفات `.xlsx` و`.xlsm` و`.csv` خليةً بخلية. تناسب مراجعة تغييرات الجداول، وفحوصات خطوط معالجة البيانات، والأتمتة وCI عندما تحتاج تقرير فروقات واضحًا بدل فتح ملفين يدويًا.

## لماذا هذا المشروع؟

مراجعة تغييرات ملفات Excel يدويًا قد تكون صعبة، خصوصًا مع تعدد الأوراق. تحول الأداة الفروقات إلى تغييرات صريحة باسم الورقة وعنوان الخلية، مع رموز خروج مناسبة للأتمتة، وكل المعالجة تتم محليًا.

## الميزات

- مقارنة ملفات XLSX/XLSM في جميع الأوراق المشتركة.
- مقارنة CSV بترميز UTF-8 كورقة منطقية واحدة.
- اكتشاف الأوراق المضافة والمحذوفة والخلايا المضافة والمحذوفة والمتغيرة.
- مقارنة القيم المخزنة للمعادلات افتراضيًا أو نص المعادلة عبر `--formulas`.
- خيار لتجاهل حالة الأحرف في النصوص.
- تقارير نصية أو JSON منظمة.
- حفظ التقرير في ملف لاستخدامه في CI أو التدقيق.
- رموز الخروج: `0` للتطابق، `1` لوجود فروقات، `2` للخطأ.
- لا شبكة ولا Telemetry ولا رفع للملفات.
- واجهة Python للاستخدام البرمجي.

## معاينة

```text
Excel Diff: before.xlsx -> after.xlsx
Result: differences found
Inventory!B2: 2 -> 5
Changed cells: 1
```

المشروع أداة CLI ولا يحتوي واجهة رسومية؛ يمكن أخذ لقطة للشاشة بعد تشغيل المثال المرفق.

## المتطلبات والتثبيت

يتطلب Python 3.10 أو أحدث و`openpyxl` 3.1 أو أحدث.

```bash
git clone https://github.com/rad03i2/excel-diff.git
cd excel-diff
python -m venv .venv
pip install -e .
```

## الاستخدام

```bash
excel-diff before.xlsx after.xlsx
excel-diff before.xlsx after.xlsx --format json
excel-diff before.xlsx after.xlsx --format json --output diff.json
excel-diff old.csv new.csv --ignore-case
excel-diff before.xlsx after.xlsx --formulas
```

لتوليد ملفي تجربة حقيقيين:

```bash
python examples/make_samples.py
excel-diff examples/generated/before.xlsx examples/generated/after.xlsx
```

ومن Python:

```python
from excel_diff import compare
report = compare("before.xlsx", "after.xlsx")
print(report.to_json())
```

## الإعداد

لا تحتاج الأداة متغيرات بيئة أو مفاتيح API أو حسابات. جميع الخيارات عبر CLI. الخيار `--limit` يحد فقط عدد الفروقات المعروضة نصيًا، بينما JSON يحتوي كل الفروقات المكتشفة.

## بنية المشروع

`src/excel_diff/` للمحرك وCLI، و`tests/` للاختبارات، و`examples/` للأمثلة، و`.github/workflows/` للتكامل المستمر.

## الاختبارات

```bash
pip install -e . pytest
python -m compileall -q src tests
pytest -q
excel-diff --version
```

يشغّل CI الاختبارات على Ubuntu وWindows وmacOS مع Python 3.10 و3.12 و3.13.

## القيود

لا يدعم `.xls` القديم، ولا يقارن التنسيق أو التعليقات أو الرسوم أو الصور أو الماكرو أو بيانات المصنف الوصفية. وضع المعادلات يقارن النص لا التكافؤ الرياضي. الوضع الافتراضي يعتمد على القيمة المخزنة داخل الملف لأن openpyxl لا يحسب المعادلات. CSV يُعامل كنص، بينما XLSX يحتفظ بأنواع القيم. كما تبقى خرائط الخلايا غير الفارغة والتقرير النهائي في الذاكرة.

## الأمان والخصوصية

كل المعالجة محلية ولا توجد اتصالات شبكية. قد يتضمن التقرير قيم الخلايا نفسها، لذلك يجب حماية التقرير مثل حماية الملف الأصلي. راجع [SECURITY.md](SECURITY.md).

## تطوير اختياري مستقبلًا

يمكن إضافة مقارنة التنسيق، وتحديد أوراق/نطاقات بعينها، وتقارير HTML. هذه إضافات اختيارية وليست وعودًا بميزات موجودة حاليًا.

## المساهمة

راجع [CONTRIBUTING.md](CONTRIBUTING.md)، واستخدم بيانات اختبار مصطنعة بدل الملفات السرية الحقيقية.

## الترخيص

MIT — راجع [LICENSE](LICENSE).

## المؤلف

**Radwan Abdulhadi Ahmed**  
**رضوان عبدالهادي أحمد**  
GitHub: [@rad03i2](https://github.com/rad03i2)
