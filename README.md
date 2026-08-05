# 📊 Excel Dev Toolkit — Business-Grade Student Management System

> A complete, multi-language Excel development toolkit and student management system.
> Business-standard, deployable to multiple schools.

---

## What Is This?

This package combines:

1. **A full Excel Development Toolkit** — supporting ALL languages that work with Excel:
   - VBA, Python, DAX, Power Query M, Office Scripts (TypeScript), C#/.NET, SQL, LAMBDA, JavaScript
2. **A business-grade Student Management System** — a complete school grade tracking & analytics platform
3. **A Docker-based document lab** — headless LibreOffice, pandoc, LaTeX, PDF tools, databases

Designed for schools: create separate workbooks per school, track grades, generate reports, export PDFs.

---

## Quick Start

### 1. Build the Document Lab (Docker)

```bash
cd doclab/
docker build -t doclab:latest .
```

### 2. Generate a School Workbook

```bash
docker run --rm --entrypoint bash doclab:latest -c '
  cd /workspace/project/python
  python3 student_management.py --school "Lincoln International School" --output /workspace/output
'
```

### 3. Run Data Analysis

```bash
docker run --rm --entrypoint bash doclab:latest -c '
  cd /workspace/project/python
  python3 data_analysis.py --input /workspace/output/Lincoln_International_School_Student_Management.xlsx
'
```

### 4. Export to PDF

```bash
docker run --rm --entrypoint bash doclab:latest -c '
  libreoffice --headless --convert-to pdf /workspace/output/*.xlsx --outdir /workspace/output
'
```

---

## Repository Structure

```
excel-dev-toolkit/
├── README.md                          ← You are here
├── doclab/
│   └── Dockerfile                     ← Docker image with ALL tools
├── project/
│   ├── python/
│   │   ├── student_management.py      ← Main: generates full .xlsx workbook
│   │   ├── data_analysis.py           ← pandas + matplotlib analytics
│   │   └── requirements.txt           ← Python dependencies
│   ├── vba/
│   │   ├── Module_StudentCRUD.bas     ← VBA: Add/Update/Delete/Search students
│   │   ├── Module_GradeCalculator.bas ← VBA: Auto-calculate grades + charts
│   │   ├── Module_ReportGenerator.bas ← VBA: Report cards + PDF export
│   │   └── README_VBA.md              ← How to import VBA into Excel
│   ├── powerquery/
│   │   ├── StudentData.m             ← Power Query M: Import & clean student data
│   │   └── GradeAnalytics.m          ← Power Query M: Unpivot + analyze scores
│   ├── dax/
│   │   └── measures.dax              ← DAX: 40+ measures for Power Pivot
│   └── office-scripts/
│       ├── gradeBook.ts              ← Office Script: Automated grading (Excel Online)
│       └── attendanceTracker.ts      ← Office Script: Attendance tracking
├── docs/
│   └── Excel_Languages_Guide.md       ← Complete guide to ALL Excel languages
├── output/                            ← Generated workbooks, PDFs, charts
│   ├── Test_School_Student_Management.xlsx
│   ├── Test_School_Student_Management.pdf
│   └── charts/
└── tests/
    └── feature_test.py               ← Comprehensive 112-point test suite
```

---

## The Student Management System

### 7 Sheets Generated

| Sheet | Purpose | Features |
|---|---|---|
| **Dashboard** | KPI cards + charts | 6 KPIs (Total Students, Class Avg, Top Score, Pass Rate, A Grades, Attendance), Bar chart, Pie chart |
| **Students** | Full student database | 24 students, 20 columns, SUM/AVERAGE/IF formulas, auto-filter, freeze panes |
| **Grades** | Grade book | 5 subjects, Total/Average/Grade/Remarks formulas, conditional formatting (color scale) |
| **Boys** | Male student records | Filtered data, formulas, subject averages, bar chart |
| **Girls** | Female student records | Filtered data, formulas, subject averages, bar chart |
| **Analytics** | Subject stats, gender comparison, rankings | AVERAGE/MAX/MIN/COUNTIF/RANK formulas, bar chart |
| **Reports** | Report card layout | Cross-sheet references to Grades + Students |

### 24 Students (Ghanaian School Dataset)

- 12 Male, 12 Female
- 5 Subjects: English, Mathematics, Science, ICT, Social Studies
- Student IDs (ST001-ST024), Parent info, Attendance %
- Grade scale: A (≥80), B (≥70), C (≥60), D (≥50), E (≥40), F (<40)

### Multi-School Support

```bash
# Generate for different schools
python3 student_management.py --school "Accra Academy" --output ./schools/accra
python3 student_management.py --school "Kumasi High School" --output ./schools/kumasi
python3 student_management.py --school "Takoradi International" --output ./schools/takoradi
```

---

## Document Lab — Full Tool Inventory

### Office Suite
| Tool | Version | What It Does |
|---|---|---|
| LibreOffice | 25.2.3.2 | .docx .xlsx .pptx .odt .pdf conversion |
| pandoc | 3.1.11.1 | Markdown ↔ Word ↔ PDF ↔ HTML ↔ LaTeX |
| TeX Live | 2025 | Full LaTeX for dissertations, math papers |

### Excel / Data
| Tool | Version | What It Does |
|---|---|---|
| openpyxl | 3.1.5 | Create/edit .xlsx with formulas, charts, formatting |
| xlsxwriter | 3.2.9 | Advanced Excel generation |
| pandas | 3.0.5 | Data analysis and manipulation |
| matplotlib | 3.11.1 | Charts and visualizations |
| seaborn | — | Statistical visualizations |
| xlwings | — | Control live Excel instance (Windows/Mac) |

### Database
| Tool | What It Does |
|---|---|
| sqlite3 3.46 | Embedded database (replaces MS Access) |
| mdbtools | Read legacy MS Access .mdb files |
| psycopg2 | PostgreSQL connector |
| pymysql | MySQL/MariaDB connector |

### PDF
| Tool | What It Does |
|---|---|
| poppler-utils | pdftotext — extract text from PDFs |
| qpdf 12.2 | Inspect, merge, split, manipulate PDFs |
| pdftk 3.3 | PDF merge, split, rotate, stamp |
| reportlab 5.0 | Generate PDFs from scratch with Python |

### Development
| Tool | What It Does |
|---|---|
| Node.js + npm | TypeScript / Office Scripts development |
| TypeScript (global) | Type-check Office Scripts |
| Mono (C# / .NET) | Excel Add-in development |
| vim 9.1 | Terminal text editor |
| jq + xmlstarlet | Inspect xlsx internal JSON/XML |

---

## Excel Languages Supported

See `docs/Excel_Languages_Guide.md` for the complete 11-language reference with code examples.

| Language | Use Case | Where to Write |
|---|---|---|
| Excel Formulas | Simple calculations | Any Excel cell |
| LAMBDA | Reusable custom formulas | Name Manager (Excel 365) |
| VBA | Desktop automation | Alt+F11 (Windows/Mac) |
| Python | Data science, automation | `project/python/` |
| Power Query M | Data import/cleaning/ETL | `project/powerquery/` |
| DAX | Power Pivot measures | `project/dax/` |
| Office Scripts (TS) | Cloud Excel automation | `project/office-scripts/` |
| C# / .NET | Enterprise add-ins | Mono in doclab |
| JavaScript | Office Add-ins | Office.js |
| SQL | Database queries | sqlite3 in doclab |
| C++ | XLL high-performance add-ins | External (Visual Studio) |

---

## Test Results

The comprehensive test suite (`tests/feature_test.py`) validates 112 features:

```
RESULTS: 111/112 passed, 1 failed
  ❌ doclab: nodejs — (expected, new Dockerfile not yet rebuilt)
  ✅ All other 111 tests PASSED
```

---

*Atlas — Business-grade tools for academic excellence.* 📚
