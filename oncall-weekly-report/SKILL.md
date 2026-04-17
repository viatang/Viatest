name: oncall-weekly-report
description: >
  Use this skill to generate a weekly Oncall ticket analysis report.
  Trigger automatically every Monday at noon, or when the user says
  "生成oncall报告", "weekly report", "oncall分析", or "发送问卷前的报告".
  Input is an Excel file with oncall ticket data. Output is a Chinese PDF report
  with objective metrics and AI-powered analysis.
---

# Oncall 周报自动生成 Skill

## 概述

每周一中午自动触发。读取 oncall 工单 Excel 数据，生成包含客观指标和 AI 智能分析的中文 PDF 报告，供发送周报/问卷前参考。

## 触发时机

| 触发方式 | 说明 |
|----------|------|
| 定时触发 | 每周一 12:00 |
| 手动触发 | 用户说"生成oncall报告"、"weekly report"、"oncall分析" |

## 输入：Excel 字段说明

Excel 文件须包含以下列（列名需与此一致）：

| 列名 | 类型 | 说明 |
|------|------|------|
| `sales department` | 文本 | 销售所在部门 |
| `agent_department` | 文本 | 客服所在部门（含 OCIC 的行为测试数据，自动过滤）|
| `decrypted_reference_extra_map` | JSON 字符串 | 嵌套字段，含 issueCategoryName、transferReason、priority_level 等 |
| `ticket_value_new` | JSON 数组字符串 | 嵌套字段，含 root_ause、issue_ategory_updated、key_issue 等 |
| `Total agent ticket` | 数字 | 客服处理工单数 |
| `Total Ticket(without transfer)` | 数字 | 未转单工单数 |
| `Total solved rate` | 数字 | 已解决工单数（0/1） |
| `> 48 hr rate` | 数字 | 是否超过48小时（0/1） |
| `Satisfaction rate` | 数字/空 | 满意度评分（1=好评，0=差评，空=未评价）|
| `First Response duration/H` | 小数 | 首次响应时长（小时）|
| `Ticket Duration(hrs)` | 小数 | 工单总处理时长（小时）|

## 输出：PDF 报告内容

报告包含以下章节：

1. **核心指标总览** — 总工单数、解决率、超48小时率、平均处理时长、首次响应时长、满意度
2. **处理时长分布** — 平均值、中位数(P50)、P75、最大值
3. **工单问题分类 Top5** — 按 issueCategoryName 统计
4. **根本原因分析 Top5** — 按 root_ause 统计
5. **AI 智能分析与建议** — 核心发现、风险点、优化建议（中文，由 Claude 生成）

## 数据清洗规则

- **过滤测试数据**：`agent_department` 含 "OCIC" 的行自动忽略
- **JSON 解析**：`decrypted_reference_extra_map` 和 `ticket_value_new` 均为 JSON 字符串，需用 `json.loads()` 解析后提取字段
- **满意度处理**：只统计值为 0 或 1 的行，空值忽略
- **数值转换**：`Satisfaction rate` 列用 `pd.to_numeric(errors='coerce')` 处理，避免类型错误

## 踩坑记录

### ⚠️ 中文字体问题
reportlab 默认不支持中文。**不要**用系统 `.ttc` 字体文件（会报 postscript outlines 错误）。
正确做法：使用 reportlab 内置的 CID 字体：

```python
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.pdfbase import pdfmetrics
pdfmetrics.registerFont(UnicodeCIDFont('STSong-Light'))
font_name = 'STSong-Light'
```

### ⚠️ JSON 嵌套字段
两个 JSON 列结构不同：
- `decrypted_reference_extra_map`：普通 JSON 对象，直接 `json.loads()` 后取字段
- `ticket_value_new`：JSON 数组，每个元素有 `fieldCode` 和 `fieldValueAll`，需遍历匹配

```python
# ticket_value_new 解析示例
def parse_ticket(s):
    result = {}
    try:
        for item in json.loads(s):
            if item.get('fieldCode') in ['root_ause', 'issue_ategory_updated', 'key_issue']:
                result[item['fieldCode']] = item.get('fieldValueAll', '')
    except:
        pass
    return result
```

### ⚠️ ticket_value_new 有 71 个空值
部分行该字段为空，`json.loads()` 会报错，务必用 try/except 包裹。

## 完整代码

见同目录下 `generate_report.py`。

运行方式：
```bash
python3 generate_report.py <excel路径> <输出pdf路径>

# 示例
python3 generate_report.py ./data/oncall_data.xlsx ./output/report.pdf
```

## 依赖

```bash
pip install pandas openpyxl reportlab requests
```
