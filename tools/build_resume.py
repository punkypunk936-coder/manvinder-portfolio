from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor


ROOT = Path(__file__).resolve().parents[1]
OUTPUT = ROOT / "assets" / "manvinder-arora-360-resume.docx"

INK = RGBColor(20, 25, 29)
MUTED = RGBColor(82, 92, 100)
ACCENT = RGBColor(8, 123, 104)
LINE = RGBColor(207, 214, 216)
WHITE = RGBColor(255, 255, 255)
FONT = "Arial"


def set_cell_shading(cell, fill):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:fill"), fill)


def mark_row_as_header(row):
    tr_pr = row._tr.get_or_add_trPr()
    tbl_header = OxmlElement("w:tblHeader")
    tbl_header.set(qn("w:val"), "true")
    tr_pr.append(tbl_header)


def set_font(run, size=9.4, bold=False, color=INK, italic=False):
    run.font.name = FONT
    run._element.get_or_add_rPr().rFonts.set(qn("w:ascii"), FONT)
    run._element.get_or_add_rPr().rFonts.set(qn("w:hAnsi"), FONT)
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    run.font.color.rgb = color


def add_hyperlink(paragraph, text, url, color=ACCENT, bold=False, size=8.7):
    part = paragraph.part
    rel_id = part.relate_to(
        url,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True,
    )
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), rel_id)
    run = OxmlElement("w:r")
    r_pr = OxmlElement("w:rPr")
    color_el = OxmlElement("w:color")
    color_el.set(qn("w:val"), str(color))
    r_pr.append(color_el)
    r_fonts = OxmlElement("w:rFonts")
    r_fonts.set(qn("w:ascii"), FONT)
    r_fonts.set(qn("w:hAnsi"), FONT)
    r_pr.append(r_fonts)
    size_el = OxmlElement("w:sz")
    size_el.set(qn("w:val"), str(int(size * 2)))
    r_pr.append(size_el)
    if bold:
        r_pr.append(OxmlElement("w:b"))
    run.append(r_pr)
    text_el = OxmlElement("w:t")
    text_el.set(qn("xml:space"), "preserve")
    text_el.text = text
    run.append(text_el)
    hyperlink.append(run)
    paragraph._p.append(hyperlink)


def add_bottom_border(paragraph, color="087B68", size="12", space="5"):
    p_pr = paragraph._p.get_or_add_pPr()
    p_bdr = p_pr.find(qn("w:pBdr"))
    if p_bdr is None:
        p_bdr = OxmlElement("w:pBdr")
        p_pr.append(p_bdr)
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), size)
    bottom.set(qn("w:space"), space)
    bottom.set(qn("w:color"), color)
    p_bdr.append(bottom)


def add_page_number(paragraph):
    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = paragraph.add_run("Page ")
    set_font(run, size=8, color=MUTED)
    fld_char_1 = OxmlElement("w:fldChar")
    fld_char_1.set(qn("w:fldCharType"), "begin")
    instr_text = OxmlElement("w:instrText")
    instr_text.set(qn("xml:space"), "preserve")
    instr_text.text = "PAGE"
    fld_char_2 = OxmlElement("w:fldChar")
    fld_char_2.set(qn("w:fldCharType"), "end")
    run._r.extend([fld_char_1, instr_text, fld_char_2])


def section_heading(doc, text):
    p = doc.add_paragraph(style="Heading 1")
    p.paragraph_format.keep_with_next = True
    run = p.add_run(text.upper())
    set_font(run, size=10.4, bold=True, color=ACCENT)
    return p


def body_paragraph(doc, text, after=4, size=9.4):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(after)
    p.paragraph_format.line_spacing = 1.08
    run = p.add_run(text)
    set_font(run, size=size, color=INK)
    return p


def bullet(doc, text, after=2.5):
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.left_indent = Inches(0.19)
    p.paragraph_format.first_line_indent = Inches(-0.12)
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(after)
    p.paragraph_format.line_spacing = 1.07
    p.paragraph_format.keep_together = True
    run = p.add_run(text)
    set_font(run, size=9.25, color=INK)
    return p


def role_header(doc, company, title, dates):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.keep_with_next = True
    left = p.add_run(f"{company} | {title}")
    set_font(left, size=10, bold=True, color=INK)
    right = p.add_run(f"    {dates}")
    set_font(right, size=8.8, bold=True, color=MUTED)
    return p


def label_detail(doc, label, detail, after=2.5):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(after)
    p.paragraph_format.line_spacing = 1.06
    label_run = p.add_run(f"{label}: ")
    set_font(label_run, size=9.15, bold=True, color=INK)
    detail_run = p.add_run(detail)
    set_font(detail_run, size=9.15, color=INK)
    return p


def configure_styles(doc):
    normal = doc.styles["Normal"]
    normal.font.name = FONT
    normal._element.rPr.rFonts.set(qn("w:ascii"), FONT)
    normal._element.rPr.rFonts.set(qn("w:hAnsi"), FONT)
    normal.font.size = Pt(9.4)
    normal.font.color.rgb = INK
    normal.paragraph_format.space_after = Pt(4)
    normal.paragraph_format.line_spacing = 1.08

    heading = doc.styles["Heading 1"]
    heading.font.name = FONT
    heading._element.rPr.rFonts.set(qn("w:ascii"), FONT)
    heading._element.rPr.rFonts.set(qn("w:hAnsi"), FONT)
    heading.font.size = Pt(10.4)
    heading.font.bold = True
    heading.font.color.rgb = ACCENT
    heading.paragraph_format.space_before = Pt(8)
    heading.paragraph_format.space_after = Pt(4)

    list_style = doc.styles["List Bullet"]
    list_style.font.name = FONT
    list_style._element.rPr.rFonts.set(qn("w:ascii"), FONT)
    list_style._element.rPr.rFonts.set(qn("w:hAnsi"), FONT)
    list_style.font.size = Pt(9.25)


def build_resume():
    doc = Document()
    doc.core_properties.title = "Manvinder Arora - Crypto-Native Operator and AI Builder"
    doc.core_properties.subject = "360-degree professional resume"
    doc.core_properties.author = "Manvinder Arora"
    doc.core_properties.keywords = "crypto, DeFi, growth, marketing, community, business development, trading, AI"

    section = doc.sections[0]
    section.page_width = Inches(8.5)
    section.page_height = Inches(11)
    section.top_margin = Inches(0.48)
    section.bottom_margin = Inches(0.48)
    section.left_margin = Inches(0.58)
    section.right_margin = Inches(0.58)
    section.header_distance = Inches(0.25)
    section.footer_distance = Inches(0.25)
    configure_styles(doc)

    footer = section.footer
    add_page_number(footer.paragraphs[0])

    # Resume masthead: memo_masthead adapted for a two-page professional resume.
    name = doc.add_paragraph()
    name.paragraph_format.space_before = Pt(0)
    name.paragraph_format.space_after = Pt(1)
    run = name.add_run("MANVINDER ARORA")
    set_font(run, size=23, bold=True, color=INK)

    title = doc.add_paragraph()
    title.paragraph_format.space_before = Pt(0)
    title.paragraph_format.space_after = Pt(4)
    run = title.add_run("CRYPTO-NATIVE OPERATOR | GROWTH, MARKETING, COMMUNITY, BD & AI BUILDER")
    set_font(run, size=9.5, bold=True, color=ACCENT)

    contact = doc.add_paragraph()
    contact.paragraph_format.space_before = Pt(0)
    contact.paragraph_format.space_after = Pt(6)
    set_font(contact.add_run("India (UTC+5:30)  |  "), size=8.55, color=MUTED)
    add_hyperlink(contact, "mani.arora03@gmail.com", "mailto:mani.arora03@gmail.com", color=MUTED, size=8.55)
    set_font(contact.add_run("  |  "), size=8.55, color=MUTED)
    add_hyperlink(contact, "X", "https://x.com/0xgoodie", color=MUTED, size=8.55)
    set_font(contact.add_run("  |  "), size=8.55, color=MUTED)
    add_hyperlink(contact, "LinkedIn", "https://www.linkedin.com/in/manvinderarora/", color=MUTED, size=8.55)
    set_font(contact.add_run("  |  "), size=8.55, color=MUTED)
    add_hyperlink(contact, "Substack", "https://manvinder.substack.com", color=MUTED, size=8.55)
    set_font(contact.add_run("  |  "), size=8.55, color=MUTED)
    add_hyperlink(contact, "Portfolio", "https://punkypunk936-coder.github.io/manvinder-portfolio/", color=MUTED, size=8.55)
    add_bottom_border(contact)

    section_heading(doc, "Profile")
    body_paragraph(
        doc,
        "Crypto-native operator with five years across DeFi community, content, growth, marketing, business development, liquidity and go-to-market. My arc began in community and product education at Timeswap, expanded into marketing and BD leadership at FusionX, and now includes active onchain trading, independent research, long-form writing and AI-assisted personal tools. I combine operator execution with a real user lens: I understand how products are explained, distributed, funded, used and improved.",
        after=4,
    )

    arc = doc.add_paragraph()
    arc.paragraph_format.space_before = Pt(1)
    arc.paragraph_format.space_after = Pt(5)
    arc.paragraph_format.line_spacing = 1.0
    run = arc.add_run("COMMUNITY & EDUCATION  ->  CONTENT & MARKETING  ->  GROWTH & BD  ->  LIQUIDITY & GTM  ->  MARKETS, RESEARCH & AI SYSTEMS")
    set_font(run, size=8.15, bold=True, color=MUTED)

    section_heading(doc, "Selected Impact")
    impact_table = doc.add_table(rows=2, cols=3)
    impact_table.autofit = False
    impact_table.allow_autofit = False
    mark_row_as_header(impact_table.rows[0])
    widths = [Inches(2.43), Inches(2.43), Inches(2.42)]
    metrics = [
        ("~50%", "TVL growth supported at Timeswap"),
        ("~$290K", "Initial LP allocations sourced"),
        ("$100M+", "30-day CobaltX volume supported"),
        ("150K+", "SummitX followers scaled"),
        ("50K + 50K", "Timeswap Discord and X reach"),
        ("2,000+", "Active users added in one month"),
    ]
    for idx, cell in enumerate(impact_table._cells):
        cell.width = widths[idx % 3]
        set_cell_shading(cell, "F2F5F4")
        cell.vertical_alignment = 1
        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.line_spacing = 1.0
        value, label = metrics[idx]
        set_font(p.add_run(f"{value}\n"), size=11, bold=True, color=ACCENT)
        set_font(p.add_run(label), size=7.65, color=MUTED)

    section_heading(doc, "Professional Experience")
    role_header(doc, "FusionX Finance", "Marketing & BD Head", "Jul 2024 - Mar 2026")
    bullet(doc, "Led marketing, business development and go-to-market across EVM and SVM products, coordinating founders, chain teams, KOLs, communities and launch partners from outreach through execution.")
    bullet(doc, "Supported CobaltX campaigns that generated more than $100M in 30-day trading volume; shaped launch narratives, partner communication and ecosystem distribution across SOON, svmBASE and svmBNB.")
    bullet(doc, "Helped scale SummitX to 150,000+ followers through positioning, content testing, product-led storytelling, memes and ecosystem-native distribution.")
    bullet(doc, "Built direct communication loops with whales, liquidity providers, traders and power users through X, Telegram and private groups; carried onboarding friction and product feedback back to growth and product teams.")
    bullet(doc, "Developed co-marketing narratives and partner positioning while writing launch threads, announcements, FAQs, campaign instructions and community updates.")

    role_header(doc, "Timeswap", "Marketing & Community Head", "Jul 2021 - Jun 2024")
    bullet(doc, "Helped build the brand, community and user education for an early oracle-less lending protocol across X, Discord and Telegram, translating technical mechanics into clear user actions.")
    bullet(doc, "Supported approximately 50% TVL growth through LP onboarding, educational content, organic demand generation and demand-led launches without relying on unsustainable token emissions.")
    bullet(doc, "Personally sourced and onboarded repeat LPs with approximate initial allocations of $200K, $60K and $30K; reactivated users after fixed-term pools expired when new pools offered compelling opportunities.")
    bullet(doc, "Grew and moderated a 50,000+ member Discord and helped scale X to roughly 50,000 followers while running daily support, product communication and community programming.")
    bullet(doc, "Designed and ran Galxe, QuestN and Phi campaigns that added 2,000+ active users (about 15%) in one month; managed eligibility questions, feedback and post-campaign retention.")
    bullet(doc, "Worked directly with KOLs and community leaders, ensuring they understood the product before promotion and turning recurring user questions into explainers, FAQs and product feedback.")

    doc.add_page_break()
    continuation = doc.add_paragraph()
    continuation.paragraph_format.space_before = Pt(0)
    continuation.paragraph_format.space_after = Pt(7)
    continuation.paragraph_format.keep_with_next = True
    set_font(continuation.add_run("MANVINDER ARORA  |  CRYPTO-NATIVE OPERATOR & AI BUILDER"), size=8.2, bold=True, color=MUTED)
    add_bottom_border(continuation, color="CFD6D8", size="6", space="4")

    role_header(doc, "Independent", "Onchain Trader, Researcher, Writer & AI Builder", "2021 - Present")
    bullet(doc, "Active user of Hyperliquid, Lighter, Variational, Nado, Aster and emerging venues; received a top-100 allocation in the Lighter airdrop through sustained product usage.")
    bullet(doc, "Trade crypto, onchain equities, indices, commodities and other RWA markets; evaluate venues across onboarding, collateral, execution, liquidity, market coverage, fees, funding, points, withdrawals and support.")
    bullet(doc, "Track market narratives, catalysts, leverage, liquidity and trader attention across X, Telegram, protocol documentation and public/onchain data, then form independent market views.")
    bullet(doc, "Write long-form articles and product commentary about perp DEXs, stablecoins, tokenized markets, community operations, incentives and trader behaviour in simple, crypto-native language.")

    section_heading(doc, "Selected Personal Systems")
    label_detail(doc, "Crypto Trading Agent", "Probability-first research and execution dashboard for setup ranking, triggers, trade review, risk context and learning from missed moves.")
    label_detail(doc, "X-to-Substack Pipeline", "Personal writing workflow that turns long-form X articles into reviewable Substack drafts and reconciles published posts.")
    label_detail(doc, "Market Bubble Live Desk", "Personal production view that combines Twitch, Kick and X signals into questions, clips and show decisions.")
    label_detail(doc, "Substack AI Digest", "Daily reader that ranks selected AI/Codex posts, removes duplicates and remembers what has already been surfaced.")
    projects = doc.add_paragraph()
    projects.paragraph_format.space_before = Pt(0)
    projects.paragraph_format.space_after = Pt(4)
    add_hyperlink(projects, "View portfolio and public repositories", "https://punkypunk936-coder.github.io/manvinder-portfolio/", bold=True, size=8.7)

    section_heading(doc, "Selected Writing")
    writing_1 = bullet(doc, "Crypto didn't become TradFi. TradFi became crypto - market reflexivity, leverage, narratives and forced selling across global markets.")
    add_hyperlink(writing_1, "  Read", "https://manvinder.substack.com/p/crypto-didnt-become-tradfi-tradfi", bold=True, size=8.5)
    writing_2 = bullet(doc, "The casino should always be open - tokenized assets, stock perpetuals and the move toward always-on markets.")
    add_hyperlink(writing_2, "  Read", "https://manvinder.substack.com/p/the-casino-should-always-be-open", bold=True, size=8.5)
    writing_3 = bullet(doc, "The Community Manager Is Dead. Long Live the Community Manager - a practical view of community work as product feedback, distribution and trust.")
    add_hyperlink(writing_3, "  More on Substack", "https://manvinder.substack.com", bold=True, size=8.5)

    section_heading(doc, "Core Capabilities")
    label_detail(doc, "Growth & GTM", "Organic acquisition, product-led campaigns, activation, retention, reactivation, A/B content testing, launch strategy and ecosystem expansion.")
    label_detail(doc, "Business Development", "Founder and partner outreach, LP/whale onboarding, KOL relationships, ecosystem co-marketing, value propositions, follow-up and liquidity activation.")
    label_detail(doc, "Community & Content", "X, Discord, Telegram, AMAs, moderation, crisis communication, threads, explainers, FAQs, memes, campaign copy and technical-to-simple writing.")
    label_detail(doc, "Markets & Product", "DeFi lending, perp DEXs, stablecoins, RWAs, liquidity, fees, funding, incentives, trader psychology, user journeys and product feedback.")
    label_detail(doc, "AI & Building", "ChatGPT, Codex, AI-assisted research, writing and rapid prototyping; Python, Node.js, dashboards, browser automation, RSS and lightweight data workflows.")

    section_heading(doc, "Working Style")
    body_paragraph(
        doc,
        "Crypto-native, first-principles and hands-on. Strongest when a role needs someone who can understand the product, speak to users and partners, write clearly, operate inside fast-moving markets, and build a practical tool when the existing workflow is not enough.",
        after=0,
    )

    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    doc.save(OUTPUT)
    print(OUTPUT)


if __name__ == "__main__":
    build_resume()
