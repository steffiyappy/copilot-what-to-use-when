from __future__ import annotations

import re
from html import unescape
from pathlib import Path
from typing import Iterable

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Inches, Pt

ROOT = Path(__file__).resolve().parent
SLIDES = ROOT / "slides"

COWORK_TASKS = {
    "Light": {
        "All teams": [
            ("Rebalance your week and protect focus time", "Rebalance the week, resolve conflicts, and protect focus blocks."),
            ("Wrap up projects and organize all related work", "Close the project with a clean, shareable OneDrive archive."),
            ("Catch up on messages and send replies automatically", "Handle routine replies and draft higher-stakes responses for review."),
        ],
        "Sales": [
            ("Pricing screenshot to customer presentation", "Turn approved pricing into a customer-ready deck and draft email."),
            ("Weekly external customer email review", "Surface action items, draft replies, and send a Teams reminder."),
        ],
        "Marketing": [
            ("Research and insights alignment recap", "Post decisions, open issues, and named next steps to Teams."),
        ],
    },
    "Medium": {
        "All teams": [
            ("Turn inbox noise into a curated intelligence brief", "Create a weekly trend brief from the publications you follow."),
        ],
        "Sales": [],
        "Marketing": [
            ("Partner and channel activation kit", "Produce partner-ready and channel-ready kits with a routing plan."),
        ],
    },
    "Heavy": {
        "All teams": [
            ("Prepare a complete out-of-office handoff", "Build a handoff from projects, owners, deadlines, and files."),
            ("Analyze and optimize your OneDrive at scale", "Create a catalog with cleanup decisions queued for approval."),
            ("Audit and surface against a standard", "Audit files against a policy and return a sortable fix list."),
            ("Onboard a new hire with a complete 30-60-90 plan", "Create the plan, getting started dashboard, and check-ins."),
        ],
        "Sales": [
            ("Product launch customer pitch", "Create the comparison, value prop, and customer-ready pitch deck."),
            ("New customer onboarding automation", "Set up onboarding artifacts, team message, and welcome email."),
            ("Customer adoption materials", "Build persona-tailored learning paths and rollout order."),
            ("ROI and value selling artifact", "Produce an executive deck and microsite with use cases and ROI model."),
        ],
        "Marketing": [
            ("Analyst briefing prep and rehearsal routing", "Build a briefing dashboard and lock rehearsal time."),
            ("Messaging drift audit and remediation routing", "Flag inconsistencies with severity, fix, and owner."),
            ("Scheduled customer signal activation", "Create a Friday brief with themes and campaign adjustments."),
        ],
    },
}

SCHEDULED_PROMPT_TIP = (
    'Any of these Cowork tasks can be scheduled to run automatically. '
    'In Microsoft 365 Copilot, hover a saved prompt and pick "Schedule this prompt", '
    'then choose a daily or weekly cadence (requires a Microsoft 365 Copilot licence).'
)

PAGES = [
    ("business.html", "business.pptx", "business"),
    ("business-lite.html", "business-lite.pptx", "business"),
    ("technical.html", "technical.pptx", "technical"),
    ("technical-lite.html", "technical-lite.pptx", "technical"),
]

THEMES = {
    "business": {
        "bg": RGBColor(255, 250, 245),
        "panel": RGBColor(255, 255, 255),
        "ink": RGBColor(27, 27, 44),
        "muted": RGBColor(90, 96, 120),
        "accent": RGBColor(26, 111, 214),
        "accent2": RGBColor(62, 142, 34),
        "accent3": RGBColor(199, 62, 120),
        "soft": RGBColor(246, 247, 251),
    },
    "technical": {
        "bg": RGBColor(14, 20, 48),
        "panel": RGBColor(27, 34, 65),
        "ink": RGBColor(245, 247, 255),
        "muted": RGBColor(174, 184, 224),
        "accent": RGBColor(0, 183, 195),
        "accent2": RGBColor(115, 184, 46),
        "accent3": RGBColor(227, 80, 143),
        "soft": RGBColor(22, 30, 69),
    },
}

TIER_COLORS = {
    "Light": RGBColor(62, 142, 34),
    "Medium": RGBColor(14, 142, 146),
    "Heavy": RGBColor(199, 62, 120),
}

SECTION_STOP_WORDS = {"System", "Light", "Dark", "References"}


def clean_text(value: str) -> str:
    value = re.sub(r"<script[\s\S]*?</script>", " ", value, flags=re.I)
    value = re.sub(r"<style[\s\S]*?</style>", " ", value, flags=re.I)
    value = re.sub(r"<[^>]+>", " ", value)
    value = unescape(value)
    value = value.replace("\u2013", ",").replace("\u2014", ",")
    value = re.sub(r"\s+", " ", value).strip()
    return value


def find_first(pattern: str, text: str) -> str:
    match = re.search(pattern, text, flags=re.I | re.S)
    return clean_text(match.group(1)) if match else ""


def split_sections(html: str) -> list[str]:
    return re.findall(r"<section[\s\S]*?</section>", html, flags=re.I)


def extract_bullets(section_html: str, title: str, limit: int = 7) -> list[str]:
    candidates = []
    selectors = [
        r'<div class="nm">([\s\S]*?)</div>',
        r'<div class="task">([\s\S]*?)</div>',
        r'<div class="t">([\s\S]*?)</div>',
        r'<div class="title">([\s\S]*?)</div>',
        r'<h3>([\s\S]*?)</h3>',
        r'<div class="lab">([\s\S]*?)</div>\s*<h3>([\s\S]*?)</h3>',
        r'<td class="task">([\s\S]*?)</td>',
    ]
    for pattern in selectors:
        for match in re.finditer(pattern, section_html, flags=re.I | re.S):
            if len(match.groups()) == 2:
                text = clean_text(match.group(1) + ": " + match.group(2))
            else:
                text = clean_text(match.group(1))
            if text and text not in candidates and text not in SECTION_STOP_WORDS and text != title:
                candidates.append(text)
    if len(candidates) < 3:
        for match in re.finditer(r"<p[^>]*>([\s\S]*?)</p>", section_html, flags=re.I | re.S):
            text = clean_text(match.group(1))
            if 20 <= len(text) <= 170 and text not in candidates:
                candidates.append(text)
    return candidates[:limit]


def parse_page(path: Path) -> dict:
    html = path.read_text(encoding="utf-8")
    title = find_first(r"<h1[^>]*>([\s\S]*?)</h1>", html) or clean_text(find_first(r"<title>([\s\S]*?)</title>", html))
    lede = find_first(r'<p class="lede">([\s\S]*?)</p>', html) or find_first(r'<div class="top-sub">([\s\S]*?)</div>', html)
    sections = []
    for section in split_sections(html):
        heading = find_first(r"<h2[^>]*>([\s\S]*?)</h2>", section)
        if not heading:
            continue
        if heading == "Cowork tasks in action":
            sections.append({"title": heading, "special": "cowork"})
            continue
        subtitle = find_first(r'<p class="sec-sub">([\s\S]*?)</p>', section) or find_first(r'<span class="hint">([\s\S]*?)</span>', section)
        bullets = extract_bullets(section, heading)
        sections.append({"title": heading, "subtitle": subtitle, "bullets": bullets, "special": None})
    refs = re.findall(r'<a href="([^"]+)"[^>]*>([\s\S]*?)</a>', html, flags=re.I | re.S)
    ref_titles = [clean_text(label) for _, label in refs]
    return {"title": title, "lede": lede, "sections": sections, "refs": ref_titles[:8]}


def set_background(slide, color: RGBColor) -> None:
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = color


def add_textbox(slide, text: str, x: float, y: float, w: float, h: float, size: int, color: RGBColor, bold: bool = False, align=None):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.margin_left = Inches(0.03)
    tf.margin_right = Inches(0.03)
    tf.margin_top = Inches(0.02)
    tf.margin_bottom = Inches(0.02)
    p = tf.paragraphs[0]
    p.text = text
    if align:
        p.alignment = align
    run = p.runs[0]
    run.font.name = "Segoe UI"
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    return box


def add_bullets(slide, items: Iterable[str], x: float, y: float, w: float, h: float, theme: dict, size: int = 15) -> None:
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.clear()
    tf.margin_left = Inches(0.08)
    tf.margin_right = Inches(0.08)
    tf.margin_top = Inches(0.04)
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = f"• {item}"
        p.level = 0
        p.font.name = "Segoe UI"
        p.font.size = Pt(size)
        p.font.color.rgb = theme["ink"]
        p.space_after = Pt(5)


def add_header(slide, title: str, theme: dict, kicker: str | None = None) -> None:
    if kicker:
        add_textbox(slide, kicker.upper(), 0.55, 0.25, 7.0, 0.25, 9, theme["muted"], True)
    add_textbox(slide, title, 0.55, 0.48, 8.7, 0.72, 22, theme["ink"], True)
    line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.55), Inches(1.22), Inches(1.2), Inches(0.05))
    line.fill.solid()
    line.fill.fore_color.rgb = theme["accent"]
    line.line.fill.background()


def add_panel(slide, x: float, y: float, w: float, h: float, theme: dict, accent: RGBColor | None = None):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    shape.fill.solid()
    shape.fill.fore_color.rgb = theme["panel"]
    shape.line.color.rgb = theme["soft"]
    if accent:
        bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(0.07))
        bar.fill.solid()
        bar.fill.fore_color.rgb = accent
        bar.line.fill.background()
    return shape


def add_title_slide(prs: Presentation, page: dict, theme: dict, kind: str) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_background(slide, theme["bg"])
    accent = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(8.4), Inches(-0.4), Inches(2.2), Inches(2.2))
    accent.fill.solid()
    accent.fill.fore_color.rgb = theme["accent"]
    accent.line.fill.background()
    add_textbox(slide, page["title"], 0.7, 0.95, 7.65, 1.35, 26, theme["ink"], True)
    add_textbox(slide, page["lede"], 0.72, 2.42, 8.45, 1.1, 14, theme["muted"])
    label = "Warm customer guide" if kind == "business" else "Technical decision guide"
    add_textbox(slide, label, 0.72, 4.8, 4.5, 0.3, 12, theme["accent"], True)


def add_section_slide(prs: Presentation, section: dict, theme: dict, idx: int) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_background(slide, theme["bg"])
    add_header(slide, section["title"], theme, f"Section {idx:02d}")
    if section.get("subtitle"):
        add_textbox(slide, section["subtitle"], 0.6, 1.35, 8.7, 0.45, 13, theme["muted"])
    bullets = section.get("bullets") or ["Use this page section as the decision checkpoint for the workflow."]
    left = bullets[:4]
    right = bullets[4:8]
    add_panel(slide, 0.6, 1.95, 4.25, 2.85, theme, theme["accent"])
    add_bullets(slide, left, 0.85, 2.22, 3.75, 2.25, theme, 13)
    add_panel(slide, 5.15, 1.95, 4.25, 2.85, theme, theme["accent2"])
    if right:
        add_bullets(slide, right, 5.4, 2.22, 3.75, 2.25, theme, 13)
    else:
        add_textbox(slide, "Key takeaway", 5.45, 2.25, 3.4, 0.3, 15, theme["ink"], True)
        add_textbox(slide, "Pick the simplest Copilot surface that can finish the job with the right level of control.", 5.45, 2.75, 3.3, 0.8, 14, theme["muted"])


def add_cowork_summary(prs: Presentation, theme: dict) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_background(slide, theme["bg"])
    add_header(slide, "Cowork tasks: light, medium, heavy", theme, "Copilot Cowork")
    notes = [
        ("Light", "Quick everyday handoffs, minutes to finish, a source or two."),
        ("Medium", "Recurring multi-step work that reasons across several sources."),
        ("Heavy", "Large multi-part jobs that produce a full set of artifacts."),
    ]
    for i, (tier, body) in enumerate(notes):
        x = 0.65 + i * 3.1
        add_panel(slide, x, 1.65, 2.75, 2.55, theme, TIER_COLORS[tier])
        add_textbox(slide, tier, x + 0.25, 1.95, 2.2, 0.38, 22, TIER_COLORS[tier], True, PP_ALIGN.CENTER)
        add_textbox(slide, body, x + 0.3, 2.55, 2.15, 0.9, 14, theme["muted"], False, PP_ALIGN.CENTER)
    add_textbox(slide, "Classification follows the source task packs: Lightweight, Everyday workflows, and Hard Problems.", 0.8, 4.65, 8.4, 0.35, 12, theme["muted"], False, PP_ALIGN.CENTER)


def add_cowork_tier_slide(prs: Presentation, tier: str, theme: dict) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_background(slide, theme["bg"])
    add_header(slide, f"Cowork tasks: {tier.lower()}", theme, "Grouped by audience")
    audiences = [(k, v) for k, v in COWORK_TASKS[tier].items() if v]
    col_w = 8.8 / len(audiences)
    for i, (audience, items) in enumerate(audiences):
        x = 0.6 + i * (col_w + 0.12)
        w = col_w - 0.05
        add_panel(slide, x, 1.38, w, 3.78, theme, TIER_COLORS[tier])
        add_textbox(slide, audience, x + 0.15, 1.6, w - 0.3, 0.3, 13, TIER_COLORS[tier], True)
        y = 2.0
        for title, outcome in items:
            if tier == "Heavy":
                add_textbox(slide, title, x + 0.18, y, w - 0.36, 0.42, 8, theme["ink"], True)
                add_textbox(slide, outcome, x + 0.18, y + 0.34, w - 0.36, 0.34, 7, theme["muted"])
                y += 0.76
            else:
                add_textbox(slide, title, x + 0.18, y, w - 0.36, 0.25, 10, theme["ink"], True)
                add_textbox(slide, outcome, x + 0.18, y + 0.24, w - 0.36, 0.33, 9, theme["muted"])
                y += 0.72


def add_cowork_schedule_tip(prs: Presentation, theme: dict) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_background(slide, theme["bg"])
    add_header(slide, "Add-on tip: schedule Cowork prompts", theme, "Copilot Cowork")
    add_panel(slide, 1.0, 1.75, 8.0, 2.45, theme, theme["accent2"])
    add_textbox(slide, "Tip", 1.35, 2.08, 1.2, 0.45, 22, theme["accent2"], True)
    add_textbox(slide, SCHEDULED_PROMPT_TIP, 1.35, 2.75, 7.2, 0.95, 17, theme["ink"])


def add_refs_slide(prs: Presentation, refs: list[str], theme: dict) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_background(slide, theme["bg"])
    add_header(slide, "References", theme, "Microsoft sources")
    clean_refs = [r for r in refs if r and r not in SECTION_STOP_WORDS][:8]
    add_bullets(slide, clean_refs, 0.8, 1.45, 8.4, 3.4, theme, 12)


def build_deck(page_file: str, out_file: str, kind: str) -> None:
    page = parse_page(ROOT / page_file)
    theme = THEMES[kind]
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(5.625)
    add_title_slide(prs, page, theme, kind)
    section_number = 1
    for section in page["sections"]:
        if section.get("special") == "cowork":
            add_cowork_summary(prs, theme)
            for tier in ["Light", "Medium", "Heavy"]:
                add_cowork_tier_slide(prs, tier, theme)
            add_cowork_schedule_tip(prs, theme)
        else:
            add_section_slide(prs, section, theme, section_number)
        section_number += 1
    add_refs_slide(prs, page["refs"], theme)
    SLIDES.mkdir(exist_ok=True)
    prs.save(SLIDES / out_file)
    print(f"Generated {SLIDES / out_file} with {len(prs.slides)} slides")


def main() -> None:
    for page_file, out_file, kind in PAGES:
        build_deck(page_file, out_file, kind)


if __name__ == "__main__":
    main()
