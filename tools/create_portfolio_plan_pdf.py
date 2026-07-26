from __future__ import annotations

from datetime import date
from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    BaseDocTemplate,
    Frame,
    KeepTogether,
    PageBreak,
    PageTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
)


ROOT = Path(__file__).resolve().parents[1]
OUTPUT_DIR = ROOT / "output" / "pdf"
OUTPUT_PATH = OUTPUT_DIR / "10000美元一年期进取组合方案-不含BRK.B.pdf"

NAVY = colors.HexColor("#17365D")
BLUE = colors.HexColor("#2E74B5")
PALE_BLUE = colors.HexColor("#DCE6F1")
PALE_GRAY = colors.HexColor("#F2F4F7")
MID_GRAY = colors.HexColor("#667085")
GREEN = colors.HexColor("#067647")
RED = colors.HexColor("#B42318")
ORANGE_BG = colors.HexColor("#FFF4E5")
ORANGE_TEXT = colors.HexColor("#8A3B12")
BORDER = colors.HexColor("#D0D5DD")


def register_fonts() -> None:
    regular = Path(r"C:\Windows\Fonts\msyh.ttc")
    bold = Path(r"C:\Windows\Fonts\msyhbd.ttc")
    if regular.exists():
        pdfmetrics.registerFont(TTFont("YaHei", str(regular), subfontIndex=0))
    else:
        raise FileNotFoundError("Microsoft YaHei font not found")
    if bold.exists():
        pdfmetrics.registerFont(TTFont("YaHeiBold", str(bold), subfontIndex=0))
    else:
        pdfmetrics.registerFont(TTFont("YaHeiBold", str(regular), subfontIndex=0))


def page_decor(canvas, doc) -> None:
    canvas.saveState()
    width, height = A4
    canvas.setStrokeColor(colors.HexColor("#E4E7EC"))
    canvas.setLineWidth(0.5)
    canvas.line(18 * mm, height - 15 * mm, width - 18 * mm, height - 15 * mm)
    canvas.setFont("YaHei", 8)
    canvas.setFillColor(MID_GRAY)
    canvas.drawString(18 * mm, height - 11.5 * mm, "一年期进取组合研究方案｜不含 BRK.B")
    canvas.drawRightString(width - 18 * mm, 10 * mm, f"第 {doc.page} 页")
    canvas.restoreState()


def p(text: str, style: ParagraphStyle) -> Paragraph:
    return Paragraph(text, style)


def bullet(text: str, style: ParagraphStyle) -> Paragraph:
    return Paragraph(f"• {text}", style)


def money(value: float) -> str:
    return f"${value:,.2f}"


def build_pdf() -> Path:
    register_fonts()
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    doc = BaseDocTemplate(
        str(OUTPUT_PATH),
        pagesize=A4,
        leftMargin=18 * mm,
        rightMargin=18 * mm,
        topMargin=20 * mm,
        bottomMargin=17 * mm,
        title="10000美元一年期进取组合方案-不含BRK.B",
        author="OpenAI Codex",
        subject="基于历史滚动一年结果与交易成本约束的研究方案",
    )
    frame = Frame(
        doc.leftMargin,
        doc.bottomMargin,
        doc.width,
        doc.height,
        id="normal",
    )
    doc.addPageTemplates([PageTemplate(id="main", frames=[frame], onPage=page_decor)])

    base = getSampleStyleSheet()
    styles = {
        "title": ParagraphStyle(
            "TitleCN",
            parent=base["Title"],
            fontName="YaHeiBold",
            fontSize=25,
            leading=32,
            textColor=NAVY,
            alignment=TA_LEFT,
            spaceAfter=8,
        ),
        "subtitle": ParagraphStyle(
            "SubtitleCN",
            parent=base["Normal"],
            fontName="YaHei",
            fontSize=12,
            leading=19,
            textColor=colors.HexColor("#475467"),
            spaceAfter=12,
        ),
        "h1": ParagraphStyle(
            "H1CN",
            parent=base["Heading1"],
            fontName="YaHeiBold",
            fontSize=16,
            leading=22,
            textColor=BLUE,
            spaceBefore=12,
            spaceAfter=8,
        ),
        "h2": ParagraphStyle(
            "H2CN",
            parent=base["Heading2"],
            fontName="YaHeiBold",
            fontSize=12,
            leading=18,
            textColor=NAVY,
            spaceBefore=8,
            spaceAfter=5,
        ),
        "body": ParagraphStyle(
            "BodyCN",
            parent=base["BodyText"],
            fontName="YaHei",
            fontSize=10.2,
            leading=17,
            textColor=colors.HexColor("#101828"),
            spaceAfter=6,
        ),
        "bullet": ParagraphStyle(
            "BulletCN",
            parent=base["BodyText"],
            fontName="YaHei",
            fontSize=10,
            leading=16,
            leftIndent=5 * mm,
            firstLineIndent=-4 * mm,
            textColor=colors.HexColor("#101828"),
            spaceAfter=4,
        ),
        "small": ParagraphStyle(
            "SmallCN",
            parent=base["BodyText"],
            fontName="YaHei",
            fontSize=8.4,
            leading=13,
            textColor=MID_GRAY,
            spaceAfter=3,
        ),
        "callout": ParagraphStyle(
            "CalloutCN",
            parent=base["BodyText"],
            fontName="YaHeiBold",
            fontSize=11,
            leading=18,
            textColor=NAVY,
            alignment=TA_LEFT,
        ),
        "cell": ParagraphStyle(
            "CellCN",
            parent=base["BodyText"],
            fontName="YaHei",
            fontSize=8.5,
            leading=12,
            textColor=colors.HexColor("#101828"),
        ),
        "cell_bold": ParagraphStyle(
            "CellBoldCN",
            parent=base["BodyText"],
            fontName="YaHeiBold",
            fontSize=8.5,
            leading=12,
            textColor=colors.HexColor("#101828"),
        ),
        "cell_header": ParagraphStyle(
            "CellHeaderCN",
            parent=base["BodyText"],
            fontName="YaHeiBold",
            fontSize=8.5,
            leading=12,
            textColor=colors.white,
            alignment=TA_CENTER,
        ),
    }

    allocation = [
        ("TQQQ", 0.30, 2_997.60, "纳指三倍做多，主要收益引擎"),
        ("QQQ", 0.20, 1_998.40, "纳指核心仓"),
        ("UPRO", 0.10, 999.20, "标普三倍做多"),
        ("SPY", 0.20, 1_998.40, "标普核心仓"),
        ("DIA", 0.15, 1_498.80, "道指分散与价值暴露"),
        ("KO", 0.05, 499.60, "防御与分红"),
    ]

    story = []
    story.append(Spacer(1, 7 * mm))
    story.append(p("10,000 美元一年期进取组合方案", styles["title"]))
    story.append(
        p(
            "不含 BRK.B｜纳入每笔买入或卖出 1 美元手续费｜数据截至 2026-07-24",
            styles["subtitle"],
        )
    )

    decision = Table(
        [[p(
            "<b>结论：</b>采用 40% 三倍做多仓位与 60% 非杠杆核心仓。"
            "计划投入 9,992 美元，预留 8 美元完成两阶段共 8 笔买入。"
            "SQQQ、SDOW、SPXU、UDOW 均不纳入初始组合。",
            styles["callout"],
        )]],
        colWidths=[doc.width],
    )
    decision.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), PALE_BLUE),
                ("BOX", (0, 0), (-1, -1), 0.8, BLUE),
                ("LEFTPADDING", (0, 0), (-1, -1), 10),
                ("RIGHTPADDING", (0, 0), (-1, -1), 10),
                ("TOPPADDING", (0, 0), (-1, -1), 9),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 9),
            ]
        )
    )
    story.append(decision)
    story.append(Spacer(1, 5 * mm))

    story.append(p("一、目标配置", styles["h1"]))
    allocation_rows = [
        [
            p("标的", styles["cell_header"]),
            p("比例", styles["cell_header"]),
            p("目标金额", styles["cell_header"]),
            p("定位", styles["cell_header"]),
        ]
    ]
    for ticker, weight, amount, role in allocation:
        allocation_rows.append(
            [
                p(f"<b>{ticker}</b>", styles["cell_bold"]),
                p(f"{weight:.0%}", styles["cell"]),
                p(money(amount), styles["cell"]),
                p(role, styles["cell"]),
            ]
        )
    allocation_rows.append(
        [
            p("<b>合计</b>", styles["cell_bold"]),
            p("<b>100%</b>", styles["cell_bold"]),
            p("<b>$9,992.00</b>", styles["cell_bold"]),
            p("另预留 $8.00 买入手续费", styles["cell"]),
        ]
    )
    allocation_table = Table(
        allocation_rows,
        colWidths=[25 * mm, 21 * mm, 31 * mm, doc.width - 77 * mm],
        repeatRows=1,
    )
    allocation_table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), NAVY),
                ("BACKGROUND", (0, -1), (-1, -1), PALE_GRAY),
                ("GRID", (0, 0), (-1, -1), 0.45, BORDER),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("ALIGN", (1, 1), (2, -1), "RIGHT"),
                ("LEFTPADDING", (0, 0), (-1, -1), 6),
                ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                ("TOPPADDING", (0, 0), (-1, -1), 6),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
            ]
        )
    )
    story.append(allocation_table)
    story.append(Spacer(1, 3 * mm))
    story.append(
        p(
            "为什么不配置反向三倍产品：SQQQ、SDOW、SPXU追求的是底层指数的每日反向三倍，"
            "长期结果会受到每日复利和波动损耗影响。若需要降低风险，优先减少 TQQQ、UPRO，"
            "而不是额外买入反向产品并承担更多手续费。",
            styles["body"],
        )
    )

    story.append(p("二、历史一年期风险画像", styles["h1"]))
    comparison = [
        [
            p("方案", styles["cell_header"]),
            p("中位收益", styles["cell_header"]),
            p("较差10%情形", styles["cell_header"]),
            p("历史最差", styles["cell_header"]),
            p("正收益比例", styles["cell_header"]),
        ],
        [
            p("本进取方案", styles["cell_bold"]),
            p("+23.6%", styles["cell"]),
            p("-39.6%", styles["cell"]),
            p("-62.5%", styles["cell"]),
            p("75.5%", styles["cell"]),
        ],
        [
            p("100% TQQQ", styles["cell"]),
            p("+41.4%", styles["cell"]),
            p("-78.3%", styles["cell"]),
            p("-99.1%", styles["cell"]),
            p("72.5%", styles["cell"]),
        ],
        [
            p("非杠杆组合", styles["cell"]),
            p("+13.7%", styles["cell"]),
            p("-17.0%", styles["cell"]),
            p("-43.2%", styles["cell"]),
            p("79.8%", styles["cell"]),
        ],
    ]
    comparison_table = Table(
        comparison,
        colWidths=[36 * mm, 28 * mm, 32 * mm, 28 * mm, 31 * mm],
        repeatRows=1,
    )
    comparison_table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), NAVY),
                ("BACKGROUND", (0, 1), (-1, 1), colors.HexColor("#EEF4FF")),
                ("GRID", (0, 0), (-1, -1), 0.45, BORDER),
                ("ALIGN", (1, 1), (-1, -1), "RIGHT"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 5),
                ("RIGHTPADDING", (0, 0), (-1, -1), 5),
                ("TOPPADDING", (0, 0), (-1, -1), 6),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
            ]
        )
    )
    story.append(comparison_table)
    story.append(
        p(
            "统计基于历史序列中的滚动 252 交易日窗口。三倍产品的上市前历史按底层每日总收益的"
            " +3 倍复利模拟。表内已扣除各方案的计划买入和一年期卖出手续费，"
            "但未计真实基金费率、融资成本和跟踪误差。上述数据不是未来收益预测。",
            styles["small"],
        )
    )

    story.append(PageBreak())
    story.append(p("三、具体买入步骤", styles["h1"]))
    story.append(
        p(
            "截至 2026-07-24，QQQ、SPY、DIA及相应三倍做多序列略低于5日均线，KO略高于5日均线。"
            "因此采用“核心仓一次买入、杠杆仓分两次”的方式。",
            styles["body"],
        )
    )

    stage1 = [
        [
            p("第一阶段：现在执行", styles["cell_header"]),
            p("买入金额", styles["cell_header"]),
            p("手续费", styles["cell_header"]),
        ],
        [p("QQQ", styles["cell_bold"]), p("$1,998.40", styles["cell"]), p("$1", styles["cell"])],
        [p("SPY", styles["cell_bold"]), p("$1,998.40", styles["cell"]), p("$1", styles["cell"])],
        [p("DIA", styles["cell_bold"]), p("$1,498.80", styles["cell"]), p("$1", styles["cell"])],
        [p("KO", styles["cell_bold"]), p("$499.60", styles["cell"]), p("$1", styles["cell"])],
        [p("TQQQ", styles["cell_bold"]), p("$1,498.80", styles["cell"]), p("$1", styles["cell"])],
        [p("UPRO", styles["cell_bold"]), p("$499.60", styles["cell"]), p("$1", styles["cell"])],
        [p("阶段合计", styles["cell_bold"]), p("$7,993.60", styles["cell_bold"]), p("$6", styles["cell_bold"])],
    ]
    stage1_table = Table(stage1, colWidths=[72 * mm, 48 * mm, 35 * mm], repeatRows=1)
    stage1_table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), BLUE),
                ("BACKGROUND", (0, -1), (-1, -1), PALE_GRAY),
                ("GRID", (0, 0), (-1, -1), 0.45, BORDER),
                ("ALIGN", (1, 1), (-1, -1), "RIGHT"),
                ("LEFTPADDING", (0, 0), (-1, -1), 6),
                ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                ("TOPPADDING", (0, 0), (-1, -1), 5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
            ]
        )
    )
    story.append(stage1_table)
    story.append(
        p(
            "执行后账户剩余现金约 $2,000.40。建议使用按美元金额下单的碎股功能；若不支持碎股，"
            "各标的按整股向下取整，剩余资金保留为现金。",
            styles["small"],
        )
    )

    story.append(p("第二阶段：满足条件后完成杠杆仓", styles["h2"]))
    story.append(
        bullet(
            "QQQ收盘重新站上5日均线：买入 TQQQ $1,498.80，手续费 $1。",
            styles["bullet"],
        )
    )
    story.append(
        bullet(
            "SPY收盘重新站上5日均线：买入 UPRO $499.60，手续费 $1。",
            styles["bullet"],
        )
    )
    story.append(
        bullet(
            "两个条件可以在不同交易日触发；未触发前对应资金保留为现金，不追涨盘前或盘后价格。",
            styles["bullet"],
        )
    )

    fee_box = Table(
        [[p(
            "<b>手续费预算：</b>计划买入 8 笔，合计 $8；一年期结束若卖出 6 个标的，"
            "再产生 $6。最基本的一买一卖周期总成本 $14，即初始本金的 0.14%。",
            styles["body"],
        )]],
        colWidths=[doc.width],
    )
    fee_box.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), ORANGE_BG),
                ("BOX", (0, 0), (-1, -1), 0.6, colors.HexColor("#F4B183")),
                ("LEFTPADDING", (0, 0), (-1, -1), 9),
                ("RIGHTPADDING", (0, 0), (-1, -1), 9),
                ("TOPPADDING", (0, 0), (-1, -1), 7),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ]
        )
    )
    story.append(Spacer(1, 2 * mm))
    story.append(fee_box)

    story.append(p("四、降低交易次数的规则", styles["h1"]))
    low_frequency_rules = [
        "每周五收盘后检查一次，不按5日均线每天全仓切换。5日均线只用于第二阶段的建仓确认。",
        "每3个月检查配置；只有某个标的偏离目标比例超过5个百分点，或 TQQQ+UPRO 合计超过45%，才再平衡。",
        "单次调整金额低于250美元时暂不交易。1美元手续费对250美元交易相当于0.4%的单边成本。",
        "不同时持有做多三倍与反向三倍产品。需要降风险时直接减少 TQQQ、UPRO。",
    ]
    for item in low_frequency_rules:
        story.append(bullet(item, styles["bullet"]))

    risk_rows = [
        [
            p("触发条件", styles["cell_header"]),
            p("下一交易日操作", styles["cell_header"]),
            p("预计交易数", styles["cell_header"]),
        ],
        [
            p("账户相对历史最高值回撤15%", styles["cell"]),
            p("TQQQ、UPRO各卖出一半，资金转为现金", styles["cell"]),
            p("2笔 / $2", styles["cell"]),
        ],
        [
            p("账户回撤扩大至25%", styles["cell"]),
            p("卖出剩余 TQQQ、UPRO；核心仓继续持有", styles["cell"]),
            p("2笔 / $2", styles["cell"]),
        ],
        [
            p("对应底层连续2日站上20日均线", styles["cell"]),
            p("恢复一半 TQQQ 或 UPRO 目标仓位", styles["cell"]),
            p("每标的1笔", styles["cell"]),
        ],
        [
            p("对应底层重新站上50日均线", styles["cell"]),
            p("恢复至完整目标仓位，但杠杆总仓不超过40%", styles["cell"]),
            p("每标的1笔", styles["cell"]),
        ],
    ]
    risk_table = Table(
        risk_rows,
        colWidths=[48 * mm, 79 * mm, 28 * mm],
        repeatRows=1,
    )
    risk_table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), NAVY),
                ("GRID", (0, 0), (-1, -1), 0.45, BORDER),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("ALIGN", (2, 1), (2, -1), "RIGHT"),
                ("LEFTPADDING", (0, 0), (-1, -1), 5),
                ("RIGHTPADDING", (0, 0), (-1, -1), 5),
                ("TOPPADDING", (0, 0), (-1, -1), 6),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
            ]
        )
    )
    risk_note = p(
            "若完整触发一次“减仓-退出-恢复”周期，通常新增不超过8笔交易，即约8美元。"
            "连同初始买入和一年期最终卖出，全年约22美元，占本金约0.22%。",
            styles["small"],
        )
    story.append(
        KeepTogether(
            [
                p("五、风险控制与重新进入", styles["h1"]),
                risk_table,
                risk_note,
            ]
        )
    )
    story.append(p("六、实际下单注意事项", styles["h1"]))
    execution_notes = [
        "使用限价单，尽量在美股正常交易时段下单；避免开盘后前15分钟和盘后流动性较低时段。",
        "金额按券商成交金额与手续费机制微调，任何时候都不要让现金余额因手续费变成负数。",
        "不要因为某个标的一天大涨就临时提高目标比例，也不要在跌破风险阈值后立即补仓。",
        "每次交易记录日期、价格、金额、手续费、交易原因和交易后目标比例。",
        "一年期结束且确实需要现金时，每个标的一次性卖出即可，预计6笔、手续费6美元；没有用款需求时，核心仓是否继续持有应另行评估。",
    ]
    for item in execution_notes:
        story.append(bullet(item, styles["bullet"]))

    story.append(p("七、这份方案仍可能发生什么", styles["h1"]))
    warnings = [
        "组合仍有40%资金配置在每日三倍做多产品，极端下跌年份可能损失大部分本金。",
        "历史最差滚动一年约为 -62.4%，意味着10,000美元可能阶段性降至约3,760美元。",
        "止损规则只能限制部分风险，跳空下跌、快速反转和连续震荡都可能让实际结果更差。",
        "交易成本虽然只有每笔1美元，但频繁使用5日均线会累积大量成本并增加踏空概率。",
        "历史模拟不包含税务影响。若账户是应税账户，短期资本利得税可能远高于交易手续费。",
    ]
    for item in warnings:
        story.append(bullet(item, styles["bullet"]))

    warning_box = Table(
        [[p(
            "<b>适用前提：</b>只有在你能接受账户一度跌至约4,000美元，并且这10,000美元在未来一年没有刚性用途时，"
            "才适合使用本进取方案。若不能接受，应取消 TQQQ、UPRO，改用 QQQ、SPY、DIA、KO 的非杠杆组合。",
            styles["callout"],
        )]],
        colWidths=[doc.width],
    )
    warning_box.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#FEE4E2")),
                ("BOX", (0, 0), (-1, -1), 0.8, RED),
                ("LEFTPADDING", (0, 0), (-1, -1), 10),
                ("RIGHTPADDING", (0, 0), (-1, -1), 10),
                ("TOPPADDING", (0, 0), (-1, -1), 8),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
            ]
        )
    )
    story.append(Spacer(1, 3 * mm))
    story.append(warning_box)

    story.append(p("八、数据与资料来源", styles["h1"]))
    sources = [
        "本地回测数据：策略回测输出/历史数据；基础序列来自 Total Real Returns，数据截至2026-07-24。",
        "ProShares TQQQ 产品页：https://www.proshares.com/our-etfs/leveraged-and-inverse/tqqq",
        "ProShares SQQQ 产品页：https://www.proshares.com/our-etfs/leveraged-and-inverse/sqqq",
        "SEC ETF Investor Bulletin：https://www.sec.gov/file/etfspdf",
    ]
    for item in sources:
        story.append(p(item, styles["small"]))
    story.append(Spacer(1, 3 * mm))
    story.append(
        p(
            "本文件是基于指定标的、历史数据和手续费假设制作的研究方案，不保证未来收益，亦不构成受托投资建议。",
            styles["small"],
        )
    )

    doc.build(story)
    return OUTPUT_PATH


if __name__ == "__main__":
    print(build_pdf())
