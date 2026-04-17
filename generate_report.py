"""
Oncall Report Generator
从 Excel 数据生成月度 Oncall 分析报告
"""

import pandas as pd
import json
import sys
import os
import requests
from datetime import datetime
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable
)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# ─────────────────────────────────────────────
# 1. 注册中文字体（使用系统字体）
# ─────────────────────────────────────────────
def register_fonts():
    from reportlab.pdfbase.cidfonts import UnicodeCIDFont
    pdfmetrics.registerFont(UnicodeCIDFont('STSong-Light'))
    return "STSong-Light"

# ─────────────────────────────────────────────
# 2. 读取并清洗数据
# ─────────────────────────────────────────────
def load_and_clean(filepath):
    df = pd.read_excel(filepath)

    # 过滤 OCIC 测试数据
    df = df[~df['agent_department'].str.contains('OCIC', na=False)].copy()
    df = df.reset_index(drop=True)

    # 解析 decrypted_reference_extra_map
    def parse_extra(s):
        try:
            return json.loads(s)
        except:
            return {}

    extra = df['decrypted_reference_extra_map'].apply(parse_extra).apply(pd.Series)

    # 解析 ticket_value_new（JSON数组）
    def parse_ticket(s):
        result = {}
        try:
            items = json.loads(s)
            for item in items:
                code = item.get('fieldCode', '')
                val = item.get('fieldValueAll', '')
                if code in ['root_ause', 'issue_ategory_updated', 'key_issue', 'transferNote']:
                    result[code] = val
        except:
            pass
        return result

    ticket = df['ticket_value_new'].apply(parse_ticket).apply(pd.Series)

    # 合并所有字段
    df = pd.concat([df, extra[['issueCategoryName', 'transferReason',
                                'priority_level', 'userDepartmentNameEn',
                                'accountSegment']].add_prefix('x_')], axis=1)
    for col in ['root_ause', 'issue_ategory_updated', 'key_issue']:
        if col in ticket.columns:
            df[col] = ticket[col]
        else:
            df[col] = ''

    return df


# ─────────────────────────────────────────────
# 3. 计算客观指标
# ─────────────────────────────────────────────
def compute_metrics(df):
    metrics = {}

    # 基础量
    metrics['total_tickets'] = len(df)
    metrics['total_solved'] = df['Total solved rate'].sum()
    metrics['solved_rate'] = round(metrics['total_solved'] / metrics['total_tickets'] * 100, 1)
    metrics['over_48h_count'] = df['> 48 hr rate'].sum()
    metrics['over_48h_rate'] = round(metrics['over_48h_count'] / metrics['total_tickets'] * 100, 1)

    # 处理时长（小时）
    dur = df['Ticket Duration(hrs)']
    metrics['avg_duration'] = round(dur.mean(), 1)
    metrics['median_duration'] = round(dur.median(), 1)
    metrics['max_duration'] = round(dur.max(), 1)
    metrics['p75_duration'] = round(dur.quantile(0.75), 1)

    # 首次响应时长
    frt = df['First Response duration/H']
    metrics['avg_frt'] = round(frt.mean(), 2)
    metrics['median_frt'] = round(frt.median(), 2)

    # 满意度
    sat = pd.to_numeric(df['Satisfaction rate'], errors='coerce').dropna()
    rated = sat[sat.isin([0, 1])]
    if len(rated) > 0:
        metrics['satisfaction_total'] = len(rated)
        metrics['satisfaction_good'] = int((rated == 1).sum())
        metrics['satisfaction_bad'] = int((rated == 0).sum())
        metrics['satisfaction_rate'] = round(metrics['satisfaction_good'] / len(rated) * 100, 1)
    else:
        metrics['satisfaction_total'] = 0
        metrics['satisfaction_good'] = 0
        metrics['satisfaction_bad'] = 0
        metrics['satisfaction_rate'] = 0

    # 问题分类 Top5
    if 'x_issueCategoryName' in df.columns:
        metrics['issue_top5'] = df['x_issueCategoryName'].value_counts().head(5).to_dict()
    else:
        metrics['issue_top5'] = {}

    # 根本原因 Top5
    if 'root_ause' in df.columns:
        metrics['root_cause_top5'] = df['root_ause'].replace('', pd.NA).dropna().value_counts().head(5).to_dict()
    else:
        metrics['root_cause_top5'] = {}

    # 转单原因 Top3（过滤空值）
    if 'x_transferReason' in df.columns:
        metrics['transfer_top3'] = df['x_transferReason'].replace('', pd.NA).dropna().value_counts().head(3).to_dict()
    else:
        metrics['transfer_top3'] = {}

    # 客服部门分布 Top5
    metrics['agent_dept_top5'] = df['agent_department'].value_counts().head(5).to_dict()

    return metrics


# ─────────────────────────────────────────────
# 4. 调用 Claude API 生成智能分析
# ─────────────────────────────────────────────
def get_ai_analysis(metrics):
    prompt = f"""
你是一位资深客服运营分析师。请根据以下oncall工单数据指标，用中文提供简洁的分析，包含以下三个部分：

1. **核心发现**（3-4条最重要的观察结论）
2. **风险点**（2-3个需要重点关注的风险）
3. **优化建议**（3-4条下一阶段可落地的行动建议）

每条保持1-2句话，要具体、有数据支撑。

--- DATA METRICS ---
Total tickets: {metrics['total_tickets']}
Solved rate: {metrics['solved_rate']}%
Over 48hr rate: {metrics['over_48h_rate']}% ({metrics['over_48h_count']} tickets)

Average ticket duration: {metrics['avg_duration']} hrs (median: {metrics['median_duration']} hrs, max: {metrics['max_duration']} hrs)
Average first response time: {metrics['avg_frt']} hrs (median: {metrics['median_frt']} hrs)

Satisfaction: {metrics['satisfaction_good']} positive / {metrics['satisfaction_bad']} negative out of {metrics['satisfaction_total']} rated tickets ({metrics['satisfaction_rate']}%)

Top issue categories:
{json.dumps(metrics['issue_top5'], indent=2)}

Top root causes:
{json.dumps(metrics['root_cause_top5'], indent=2)}

Top transfer reasons:
{json.dumps(metrics['transfer_top3'], indent=2)}
"""

    try:
        response = requests.post(
            "https://api.anthropic.com/v1/messages",
            headers={"Content-Type": "application/json"},
            json={
                "model": "claude-sonnet-4-20250514",
                "max_tokens": 1000,
                "messages": [{"role": "user", "content": prompt}]
            }
        )
        data = response.json()
        return data['content'][0]['text']
    except Exception as e:
        return f"[AI analysis unavailable: {e}]"


# ─────────────────────────────────────────────
# 5. 生成 PDF 报告
# ─────────────────────────────────────────────
def generate_pdf(metrics, ai_analysis, output_path, font_name):
    doc = SimpleDocTemplate(
        output_path,
        pagesize=A4,
        leftMargin=2*cm, rightMargin=2*cm,
        topMargin=2*cm, bottomMargin=2*cm
    )

    styles = getSampleStyleSheet()
    W = A4[0] - 4*cm  # 可用宽度

    # 自定义样式
    def S(name, **kw):
        base = kw.pop('parent', 'Normal')
        s = ParagraphStyle(name, parent=styles[base], fontName=font_name, **kw)
        return s

    s_title   = S('Title2',   fontSize=20, textColor=colors.HexColor('#1a1a2e'), spaceAfter=4, leading=26)
    s_sub     = S('Sub',      fontSize=11, textColor=colors.HexColor('#555'), spaceAfter=16)
    s_h2      = S('H2',       fontSize=13, textColor=colors.HexColor('#1a73e8'), spaceBefore=16, spaceAfter=6, leading=18)
    s_body    = S('Body',     fontSize=10, leading=16, spaceAfter=4)
    s_bullet  = S('Bullet',   fontSize=10, leading=16, leftIndent=12, spaceAfter=3)

    story = []

    # ── 标题区 ──
    now = datetime.now().strftime("%B %Y")
    story.append(Paragraph(f"Oncall 工单分析报告", s_title))
    story.append(Paragraph(f"报告周期：{now}  |  生成时间：{datetime.now().strftime('%Y-%m-%d')}", s_sub))
    story.append(HRFlowable(width=W, thickness=1.5, color=colors.HexColor('#1a73e8'), spaceAfter=16))

    # ── 核心指标卡片（用表格模拟） ──
    story.append(Paragraph("核心指标总览", s_h2))

    card_data = [
        ["总工单数", "解决率", "超48小时率", "平均处理时长"],
        [
            str(metrics['total_tickets']),
            f"{metrics['solved_rate']}%",
            f"{metrics['over_48h_rate']}%",
            f"{metrics['avg_duration']} 小时"
        ],
        ["平均首次响应", "满意度", "已评价工单", "最长处理时长"],
        [
            f"{metrics['avg_frt']} 小时",
            f"{metrics['satisfaction_rate']}%",
            str(metrics['satisfaction_total']),
            f"{metrics['max_duration']} 小时"
        ]
    ]

    col_w = W / 4
    card_table = Table(card_data, colWidths=[col_w]*4)
    card_table.setStyle(TableStyle([
        # 标签行
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#e8f0fe')),
        ('BACKGROUND', (0,2), (-1,2), colors.HexColor('#e8f0fe')),
        ('FONTNAME',   (0,0), (-1,-1), font_name),
        ('FONTSIZE',   (0,0), (-1,0), 8),
        ('FONTSIZE',   (0,2), (-1,2), 8),
        ('FONTSIZE',   (0,1), (-1,1), 16),
        ('FONTSIZE',   (0,3), (-1,3), 16),
        ('FONTWEIGHT', (0,1), (-1,1), 'BOLD'),
        ('FONTWEIGHT', (0,3), (-1,3), 'BOLD'),
        ('TEXTCOLOR',  (0,1), (-1,1), colors.HexColor('#1a73e8')),
        ('TEXTCOLOR',  (0,3), (-1,3), colors.HexColor('#1a73e8')),
        ('ALIGN',      (0,0), (-1,-1), 'CENTER'),
        ('VALIGN',     (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 8),
        ('BOTTOMPADDING', (0,0), (-1,-1), 8),
        ('GRID',       (0,0), (-1,-1), 0.5, colors.HexColor('#dadce0')),
        ('ROWBACKGROUNDS', (0,1), (-1,1), [colors.white]),
        ('ROWBACKGROUNDS', (0,3), (-1,3), [colors.white]),
    ]))
    story.append(card_table)
    story.append(Spacer(1, 0.4*cm))

    # ── 处理时长分布 ──
    story.append(Paragraph("处理时长分布", s_h2))
    dur_data = [
        ["指标", "数值"],
        ["平均值", f"{metrics['avg_duration']} 小时"],
        ["中位数 (P50)", f"{metrics['median_duration']} 小时"],
        ["P75", f"{metrics['p75_duration']} 小时"],
        ["最大值", f"{metrics['max_duration']} 小时"],
    ]
    dur_table = Table(dur_data, colWidths=[W*0.5, W*0.5])
    dur_table.setStyle(TableStyle([
        ('BACKGROUND',  (0,0), (-1,0), colors.HexColor('#1a73e8')),
        ('TEXTCOLOR',   (0,0), (-1,0), colors.white),
        ('FONTNAME',    (0,0), (-1,-1), font_name),
        ('FONTSIZE',    (0,0), (-1,-1), 10),
        ('ALIGN',       (1,0), (1,-1), 'CENTER'),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor('#f8f9fa')]),
        ('GRID',        (0,0), (-1,-1), 0.5, colors.HexColor('#dadce0')),
        ('TOPPADDING',  (0,0), (-1,-1), 7),
        ('BOTTOMPADDING', (0,0), (-1,-1), 7),
        ('LEFTPADDING', (0,0), (-1,-1), 10),
    ]))
    story.append(dur_table)

    # ── 问题分类 Top5 ──
    story.append(Paragraph("工单问题分类 Top5", s_h2))
    if metrics['issue_top5']:
        issue_data = [["问题分类", "数量"]] + [
            [cat.split('/')[-1][:55], str(cnt)]
            for cat, cnt in metrics['issue_top5'].items()
        ]
        issue_table = Table(issue_data, colWidths=[W*0.75, W*0.25])
        issue_table.setStyle(TableStyle([
            ('BACKGROUND',  (0,0), (-1,0), colors.HexColor('#34a853')),
            ('TEXTCOLOR',   (0,0), (-1,0), colors.white),
            ('FONTNAME',    (0,0), (-1,-1), font_name),
            ('FONTSIZE',    (0,0), (-1,-1), 10),
            ('ALIGN',       (1,0), (1,-1), 'CENTER'),
            ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor('#f0faf4')]),
            ('GRID',        (0,0), (-1,-1), 0.5, colors.HexColor('#dadce0')),
            ('TOPPADDING',  (0,0), (-1,-1), 7),
            ('BOTTOMPADDING', (0,0), (-1,-1), 7),
            ('LEFTPADDING', (0,0), (-1,-1), 10),
        ]))
        story.append(issue_table)

    # ── 根本原因 Top5 ──
    story.append(Paragraph("根本原因分析 Top5", s_h2))
    if metrics['root_cause_top5']:
        rc_data = [["根本原因", "数量"]] + [
            [rc[:60], str(cnt)]
            for rc, cnt in metrics['root_cause_top5'].items()
        ]
        rc_table = Table(rc_data, colWidths=[W*0.75, W*0.25])
        rc_table.setStyle(TableStyle([
            ('BACKGROUND',  (0,0), (-1,0), colors.HexColor('#fa7b17')),
            ('TEXTCOLOR',   (0,0), (-1,0), colors.white),
            ('FONTNAME',    (0,0), (-1,-1), font_name),
            ('FONTSIZE',    (0,0), (-1,-1), 10),
            ('ALIGN',       (1,0), (1,-1), 'CENTER'),
            ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor('#fff4ec')]),
            ('GRID',        (0,0), (-1,-1), 0.5, colors.HexColor('#dadce0')),
            ('TOPPADDING',  (0,0), (-1,-1), 7),
            ('BOTTOMPADDING', (0,0), (-1,-1), 7),
            ('LEFTPADDING', (0,0), (-1,-1), 10),
        ]))
        story.append(rc_table)

    # ── AI 智能分析 ──
    story.append(Paragraph("AI 智能分析与建议", s_h2))
    story.append(HRFlowable(width=W, thickness=0.5, color=colors.HexColor('#dadce0'), spaceAfter=8))

    for line in ai_analysis.split('\n'):
        line = line.strip()
        if not line:
            story.append(Spacer(1, 0.2*cm))
        elif line.startswith('**') and line.endswith('**'):
            story.append(Paragraph(line.replace('**',''), s_h2))
        elif line.startswith('- ') or line.startswith('* '):
            story.append(Paragraph('• ' + line[2:], s_bullet))
        elif line.startswith('#'):
            story.append(Paragraph(line.lstrip('#').strip(), s_h2))
        else:
            story.append(Paragraph(line, s_body))

    # ── 页脚 ──
    story.append(Spacer(1, 0.5*cm))
    story.append(HRFlowable(width=W, thickness=0.5, color=colors.HexColor('#dadce0')))
    story.append(Paragraph(
        f"由 Oncall Report Skill 生成  |  {datetime.now().strftime('%Y-%m-%d %H:%M')}",
        S('Footer', fontSize=8, textColor=colors.HexColor('#999'), alignment=1)
    ))

    doc.build(story)
    print(f"✅ Report saved: {output_path}")


# ─────────────────────────────────────────────
# 主流程
# ─────────────────────────────────────────────
if __name__ == "__main__":
    filepath = sys.argv[1] if len(sys.argv) > 1 else '/mnt/user-data/uploads/聚类分析.xlsx'
    output   = sys.argv[2] if len(sys.argv) > 2 else '/mnt/user-data/outputs/oncall_report.pdf'

    print("📂 Loading data...")
    df = load_and_clean(filepath)
    print(f"   {len(df)} valid tickets loaded")

    print("📊 Computing metrics...")
    metrics = compute_metrics(df)

    print("🤖 Generating AI analysis...")
    ai_analysis = get_ai_analysis(metrics)

    print("📄 Building PDF...")
    font_name = register_fonts()
    os.makedirs(os.path.dirname(output), exist_ok=True)
    generate_pdf(metrics, ai_analysis, output, font_name)
