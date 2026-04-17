# Claude Skills — Oncall Weekly Report

A collection of custom Claude skills for customer support operations.

---

## 📦 Skills

### oncall-weekly-report

Automatically generates a Chinese PDF analysis report from oncall ticket Excel data.

**Triggers:**
- Every Monday at noon (scheduled)
- When you say: `生成oncall报告` / `weekly report` / `oncall分析`

**What it does:**
1. Reads oncall ticket data from an Excel file
2. Filters out test data (OCIC department)
3. Parses nested JSON fields to extract issue categories, root causes, transfer reasons
4. Calculates objective metrics (duration, response time, satisfaction rate)
5. Uses Claude AI to generate Chinese analysis with findings, risks, and recommendations
6. Outputs a formatted PDF report

**Output preview:**

| Section | Content |
|---------|---------|
| 核心指标总览 | Total tickets, solved rate, over-48hr rate, avg duration |
| 处理时长分布 | Average, median (P50), P75, max |
| 工单问题分类 Top5 | By issue category |
| 根本原因分析 Top5 | By root cause |
| AI 智能分析与建议 | Key findings, risk points, recommendations |

---

## 🚀 How to Use

### In Claude.ai

1. Download the `oncall-weekly-report/` folder and zip it
2. Go to **Settings → Customize → Skills → Upload Skill**
3. Make sure **Code execution** is enabled in **Settings → Capabilities**
4. Upload your Excel file and say: `生成oncall报告`

### Required Excel columns

| Column | Type | Description |
|--------|------|-------------|
| `agent_department` | text | Agent's department (rows with "OCIC" are filtered as test data) |
| `decrypted_reference_extra_map` | JSON string | Contains issueCategoryName, transferReason, priority_level |
| `ticket_value_new` | JSON array string | Contains root_ause, issue_ategory_updated, key_issue |
| `Total solved rate` | number | 1 = solved, 0 = unsolved |
| `> 48 hr rate` | number | 1 = exceeded 48hrs |
| `Satisfaction rate` | number / empty | 1 = positive, 0 = negative, empty = not rated |
| `First Response duration/H` | float | First response time in hours |
| `Ticket Duration(hrs)` | float | Total ticket duration in hours |

### Dependencies

```bash
pip install pandas openpyxl reportlab requests
```

---

## 📁 Repository Structure

```
claude-skills/
└── oncall-weekly-report/
    ├── SKILL.md              # Skill definition and instructions for Claude
    └── generate_report.py    # Report generation script
```

---

## 🛠️ Development Notes

- **Chinese font**: Uses ReportLab's built-in `STSong-Light` CID font — do NOT use system `.ttc` fonts (postscript outlines not supported)
- **JSON parsing**: `ticket_value_new` is a JSON array; iterate `fieldCode` to extract values
- **Null handling**: `ticket_value_new` has ~10% null rows — always wrap JSON parsing in try/except

---

## 📄 License

MIT License — free to use and modify.
