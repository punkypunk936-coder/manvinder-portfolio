from __future__ import annotations

from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT
from reportlab.lib.pagesizes import LETTER
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import (
    HRFlowable,
    KeepTogether,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)


ROOT = Path(__file__).resolve().parents[1]
ASSETS = ROOT / "assets"

INK = colors.HexColor("#121619")
MUTED = colors.HexColor("#5C666D")
GREEN = colors.HexColor("#087B68")
PALE = colors.HexColor("#EAF1EF")
LINE = colors.HexColor("#D8DDDF")
WHITE = colors.white

CONTACT = (
    '<link href="mailto:mani.arora03@gmail.com" color="#087B68">Email</link>'
    '  |  <link href="https://x.com/0xgoodie" color="#087B68">X</link>'
    '  |  <link href="https://www.linkedin.com/in/manvinderarora/" color="#087B68">LinkedIn</link>'
    '  |  <link href="https://punkypunk936-coder.github.io/manvinder-portfolio/" color="#087B68">Portfolio</link>'
)

METRICS = [
    ("150K+", "audience scaled"),
    ("50K+", "Discord managed"),
    ("2,000+", "users activated"),
    ("$100M+", "campaign volume"),
    ("~$290K", "LP capital onboarded"),
]

CVS = [
    {
        "filename": "manvinder-arora-community-growth-cv.pdf",
        "title": "COMMUNITY & GROWTH",
        "summary": (
            "Crypto community and growth operator with five years across community ownership, "
            "GTM, creator coordination, product education, user activation and launch execution. "
            "I stay close to users, make difficult products easier to understand and measure work "
            "through activation, retention, liquidity and volume."
        ),
        "ownership": "Community operations  |  GTM campaigns  |  KOL and partner coordination  |  User activation  |  Launch execution  |  Sentiment and feedback",
        "fusion": [
            "Ran marketing and BD across EVM and SVM products, coordinating founders, chain teams, creators, partners and communities around launches.",
            "Helped scale SummitX to 150K+ followers and supported CobaltX campaigns tied to more than $100M in 30-day trading volume.",
            "Managed trader and community groups, launch messaging and product feedback loops across FusionX, SummitX and CobaltX.",
        ],
        "timeswap": [
            "Managed a 50K+ Discord and community/support activity across Discord, Telegram and X.",
            "Ran AMAs, education, campaigns and launch communication that made an oracle-less lending product easier to understand.",
            "Helped activate 2,000+ users in one month, contributed to roughly 50% TVL growth and personally onboarded about $290K in starting LP capital.",
            "Worked directly with KOLs, liquidity partners and users; converted recurring questions into clearer communication and product feedback.",
        ],
        "independent": [
            "Write about crypto markets, products and communities, and create social posts that have reached six-figure audiences.",
            "Build personal AI workflows for market analysis, community monitoring, content and publishing.",
        ],
        "links": [
            ("Community Manager", "https://x.com/0xgoodie/status/2075661908217364766"),
            ("Perp DEX research", "https://manvinder.substack.com/p/perp-dexs-cryptos-clearest-pmf"),
            ("Meme Scout", "https://github.com/punkypunk936-coder/meme-scout-agent"),
        ],
        "context": "Discord, Telegram, X, DeFi products, community analytics, campaign reporting, launch messaging and AI-assisted content workflows.",
    },
    {
        "filename": "manvinder-arora-customer-success-cv.pdf",
        "title": "CUSTOMER SUPPORT & SUCCESS",
        "summary": (
            "Crypto-native customer support and success operator with five years helping users "
            "understand products, solve problems and stay active. I connect support, community and "
            "product: answer clearly, spot repeated friction, escalate the right issue and make the "
            "next user's journey easier."
        ),
        "ownership": "Onboarding  |  Troubleshooting  |  Escalation  |  Product education  |  FAQs and knowledge bases  |  Voice of customer",
        "fusion": [
            "Worked with traders, communities and product teams across several EVM and SVM products during launches and campaigns.",
            "Built direct feedback loops with active users and turned product friction into clearer onboarding, communication and escalation context.",
            "Supported campaigns tied to more than $100M in 30-day volume while keeping user and trader concerns visible to the operating team.",
        ],
        "timeswap": [
            "Managed and moderated a 50K+ Discord while supporting users across Discord, Telegram and X.",
            "Explained complex lending mechanics through direct support, AMAs, educational content, FAQs and launch messages.",
            "Used repeated questions and community sentiment to improve support answers, product communication and the feedback sent to product teams.",
            "Helped activate 2,000+ users in one month and supported repeat LP relationships representing about $290K in starting capital.",
        ],
        "independent": [
            "Use wallets, DeFi protocols and perp venues as a real user, giving me first-hand context on onboarding, execution and support friction.",
            "Build searchable knowledge and publishing tools that make retrieval, explanation and follow-up faster.",
        ],
        "links": [
            ("Community operations essay", "https://x.com/0xgoodie/status/2075661908217364766"),
            ("Portfolio", "https://punkypunk936-coder.github.io/manvinder-portfolio/"),
            ("Telegram Brain", "https://github.com/punkypunk936-coder/telegram-brain"),
        ],
        "context": "Discord, Telegram, X, wallets, block explorers, transaction troubleshooting, product documentation, FAQs and searchable knowledge systems.",
    },
    {
        "filename": "manvinder-arora-trader-user-operations-cv.pdf",
        "title": "TRADER & USER OPERATIONS",
        "summary": (
            "Crypto-native trader and user operations professional with five years between users, "
            "community, growth and DeFi products. I understand wallets, liquidity, execution and "
            "perp venues from actual use, then turn that context into better onboarding, support, "
            "retention and product feedback."
        ),
        "ownership": "Trader onboarding  |  DeFi product support  |  Liquidity context  |  Execution and venue research  |  Retention  |  Voice of trader",
        "fusion": [
            "Coordinated product launches, KOLs, partners and trader communities across FusionX, SummitX and CobaltX.",
            "Supported CobaltX campaigns tied to more than $100M in 30-day trading volume and helped scale SummitX to 150K+ followers.",
            "Maintained direct feedback loops with traders, whales and communities so product and launch teams had useful operating context.",
        ],
        "timeswap": [
            "Helped users understand lending, borrowing and liquidity flows through support, education, AMAs and community operations.",
            "Personally onboarded repeat liquidity providers with about $290K in starting capital and contributed to roughly 50% TVL growth.",
            "Managed a 50K+ Discord and helped activate 2,000+ users in one month while staying close to product questions and liquidity friction.",
        ],
        "independent": [
            "Actively use wallets, DeFi protocols and perp venues including Hyperliquid, Lighter, Variational, Nado and Aster.",
            "Compare products through execution, liquidity, fees, funding, incentives and the full user journey; write product and market research from that perspective.",
            "Built a personal crypto trading dashboard that records trade theses, market context, position management and P&L reasoning.",
        ],
        "links": [
            ("Live trading dashboard", "https://punkypunk936-coder.github.io/crypto-trading-agent/"),
            ("Perp DEX research", "https://manvinder.substack.com/p/perp-dexs-cryptos-clearest-pmf"),
            ("Trading agent code", "https://github.com/punkypunk936-coder/crypto-trading-agent"),
        ],
        "context": "Hyperliquid, Lighter, Variational, Nado, Aster, wallets, DeFi protocols, perp markets, liquidity, funding, execution and AI-assisted research.",
    },
]


def styles():
    base = getSampleStyleSheet()
    return {
        "name": ParagraphStyle(
            "Name",
            parent=base["Normal"],
            fontName="Helvetica-Bold",
            fontSize=24,
            leading=25,
            textColor=INK,
            spaceAfter=1,
        ),
        "title": ParagraphStyle(
            "RoleTitle",
            parent=base["Normal"],
            fontName="Helvetica-Bold",
            fontSize=10.4,
            leading=12,
            textColor=GREEN,
            spaceAfter=4,
        ),
        "contact": ParagraphStyle(
            "Contact",
            parent=base["Normal"],
            fontName="Helvetica",
            fontSize=8,
            leading=10,
            textColor=MUTED,
            spaceAfter=5,
        ),
        "summary": ParagraphStyle(
            "Summary",
            parent=base["Normal"],
            fontName="Helvetica",
            fontSize=9.15,
            leading=12,
            textColor=INK,
            spaceBefore=6,
            spaceAfter=7,
        ),
        "section": ParagraphStyle(
            "Section",
            parent=base["Normal"],
            fontName="Helvetica-Bold",
            fontSize=9.1,
            leading=11,
            textColor=GREEN,
            spaceBefore=8,
            spaceAfter=3,
        ),
        "role": ParagraphStyle(
            "Role",
            parent=base["Normal"],
            fontName="Helvetica-Bold",
            fontSize=9.25,
            leading=11,
            textColor=INK,
            spaceBefore=4,
            spaceAfter=2,
        ),
        "bullet": ParagraphStyle(
            "Bullet",
            parent=base["Normal"],
            fontName="Helvetica",
            fontSize=8.45,
            leading=10.55,
            leftIndent=11,
            firstLineIndent=-8,
            textColor=INK,
            bulletIndent=0,
            spaceAfter=1.8,
        ),
        "small": ParagraphStyle(
            "Small",
            parent=base["Normal"],
            fontName="Helvetica",
            fontSize=8.2,
            leading=10.3,
            textColor=INK,
            spaceAfter=2,
        ),
    }


def metric_table(style_map):
    data = []
    for value, label in METRICS:
        data.append(
            Paragraph(
                f'<font color="#087B68" size="11.5"><b>{value}</b></font><br/>'
                f'<font color="#5C666D" size="7">{label}</font>',
                style_map["small"],
            )
        )
    table = Table([data], colWidths=[1.085 * inch] * 5, rowHeights=[.55 * inch])
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), PALE),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("LEFTPADDING", (0, 0), (-1, -1), 4),
                ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                ("LINEBEFORE", (1, 0), (-1, -1), .4, WHITE),
            ]
        )
    )
    return table


def role_block(style_map, company, role, dates, bullets):
    heading = Paragraph(
        f"{company}  |  {role}  <font color='#5C666D' size='7.2'>{dates}</font>",
        style_map["role"],
    )
    parts = [heading]
    for text in bullets:
        parts.append(Paragraph(text, style_map["bullet"], bulletText="-"))
    return KeepTogether(parts)


def link_line(style_map, links):
    pieces = []
    for label, url in links:
        pieces.append(f'<link href="{url}" color="#087B68"><b>{label}</b></link>')
    return Paragraph("  |  ".join(pieces), style_map["small"])


def footer(canvas, doc):
    canvas.saveState()
    canvas.setStrokeColor(LINE)
    canvas.setLineWidth(.45)
    canvas.line(doc.leftMargin, 24, LETTER[0] - doc.rightMargin, 24)
    canvas.setFillColor(MUTED)
    canvas.setFont("Helvetica", 6.7)
    canvas.drawString(doc.leftMargin, 13, "Portfolio: punkypunk936-coder.github.io/manvinder-portfolio")
    canvas.drawRightString(LETTER[0] - doc.rightMargin, 13, "Manvinder Arora")
    canvas.linkURL(
        "https://punkypunk936-coder.github.io/manvinder-portfolio/",
        (doc.leftMargin, 10, doc.leftMargin + 245, 21),
        relative=0,
    )
    canvas.restoreState()


def build_cv(cv):
    output = ASSETS / cv["filename"]
    doc = SimpleDocTemplate(
        str(output),
        pagesize=LETTER,
        leftMargin=.53 * inch,
        rightMargin=.53 * inch,
        topMargin=.48 * inch,
        bottomMargin=.5 * inch,
        title=f"Manvinder Arora - {cv['title'].title()} CV",
        author="Manvinder Arora",
        subject="One-page professional CV",
    )
    s = styles()
    story = [
        Paragraph("MANVINDER ARORA", s["name"]),
        Paragraph(cv["title"], s["title"]),
        Paragraph("India (UTC+5:30)  |  " + CONTACT, s["contact"]),
        HRFlowable(width="100%", thickness=.8, color=GREEN, spaceBefore=0, spaceAfter=2),
        Paragraph(cv["summary"], s["summary"]),
        metric_table(s),
        Paragraph("WHAT I CAN OWN", s["section"]),
        Paragraph(cv["ownership"], s["small"]),
        Paragraph("EXPERIENCE", s["section"]),
        role_block(s, "FusionX Finance", "Marketing & BD Head", "Jul 2024 - Mar 2026", cv["fusion"]),
        role_block(s, "Timeswap", "Marketing & Community Head", "Jul 2021 - Jun 2024", cv["timeswap"]),
        role_block(s, "Independent", "Trader, Writer & AI Builder", "2021 - Present", cv["independent"]),
        Paragraph("SELECTED WORK", s["section"]),
        link_line(s, cv["links"]),
        Paragraph("PRODUCT & TOOL CONTEXT", s["section"]),
        Paragraph(cv["context"], s["small"]),
    ]
    doc.build(story, onFirstPage=footer, onLaterPages=footer)
    return output


def main():
    ASSETS.mkdir(parents=True, exist_ok=True)
    for cv in CVS:
        print(build_cv(cv))


if __name__ == "__main__":
    main()
