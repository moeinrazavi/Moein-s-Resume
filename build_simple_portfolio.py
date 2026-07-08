#!/usr/bin/env python3
"""Short, simple, story-first portfolio for the Ian Jiang interview.

Built around five real production projects from Ford:
  1. ESA-FR Chatbot           — multi-agent EV analytics chatbot
  2. HOMELOAD                 — multi-agent utility data extraction
  3. ESS Routine Finder       — time-series forecasting for EV charging
  4. CPA Action Prioritizer   — Google ADK agent app
  5. Code Complexity Analyzer — LLM-powered code & GCP cost analysis

Each story gets two slides:
  · a plain-English "Problem -> What I built -> What changed" story slide
  · a follow-up detail slide with the agents/workflow + tech stack
"""

from __future__ import annotations

from pathlib import Path

import requests
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Emu, Inches, Pt

ROOT = Path(__file__).resolve().parent
ASSETS = ROOT / "presentation-assets"
OUTPUT = ROOT / "Moein-Razavi-Simple-Portfolio.pptx"

# ---------------------------------------------------------------------------
# Palette — clean, friendly, distinct from the technical deck
# ---------------------------------------------------------------------------
NAVY = RGBColor(20, 42, 78)
NAVY_DARK = RGBColor(13, 27, 51)
INK = RGBColor(28, 36, 54)
BODY = RGBColor(70, 82, 102)
MUTED = RGBColor(130, 140, 158)
LINE = RGBColor(220, 226, 236)
CREAM = RGBColor(250, 248, 244)
PAPER = RGBColor(255, 254, 252)
CORAL = RGBColor(232, 96, 80)
GOLD = RGBColor(232, 178, 70)
TEAL = RGBColor(52, 152, 162)
PLUM = RGBColor(150, 90, 158)
WHITE = RGBColor(255, 255, 255)
SOFT_WHITE = RGBColor(238, 244, 252)

# Per-project accent color (one color = one project, consistent on both slides)
ACCENT_BY_PROJECT = {
    1: CORAL,   # Chatbot
    2: NAVY,    # HOMELOAD
    3: TEAL,    # Routine Finder
    4: PLUM,    # CPA Action Prioritizer
    5: GOLD,    # Code Complexity Analyzer
}

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)

IMAGE_URLS = {
    "hero": "https://images.unsplash.com/photo-1454165804606-c3d57bc86b40?w=1920&q=80",
    "ai": "https://images.unsplash.com/photo-1677442136019-21780ecad995?w=1920&q=80",
    "cloud": "https://images.unsplash.com/photo-1451187580459-43490279c0fa?w=1920&q=80",
    "data": "https://images.unsplash.com/photo-1551288049-bebda4e38f71?w=1920&q=80",
    "pipeline": "https://images.unsplash.com/photo-1518432031352-d6fc5c10da5a?w=1920&q=80",
    "logistics": "https://images.unsplash.com/photo-1586528116311-ad8dd3c8310d?w=1920&q=80",
    "team": "https://images.unsplash.com/photo-1522071820081-009f0129c71c?w=1920&q=80",
    "architecture": "https://images.unsplash.com/photo-1558494949-ef010cbdcc31?w=1920&q=80",
}


def load_images() -> dict[str, Path]:
    ASSETS.mkdir(exist_ok=True)
    paths: dict[str, Path] = {}
    s = requests.Session()
    s.headers["User-Agent"] = "Moein-Simple-Portfolio/1.0"
    for key, url in IMAGE_URLS.items():
        dest = ASSETS / f"{key}.jpg"
        if not dest.exists():
            try:
                r = s.get(url, timeout=30)
                r.raise_for_status()
                dest.write_bytes(r.content)
            except Exception as exc:
                print(f"[warn] {key}: {exc}")
                continue
        paths[key] = dest
    return paths


# ---------------------------------------------------------------------------
# Primitives
# ---------------------------------------------------------------------------
def blank(prs): return prs.slides.add_slide(prs.slide_layouts[6])


def bg(slide, color):
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = color


def fill(shape, color, transparency=None):
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()
    if transparency is not None:
        shape.fill.transparency = transparency


def rect(slide, l, t, w, h, color, transparency=None):
    s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, l, t, w, h)
    fill(s, color, transparency)
    return s


def round_rect(slide, l, t, w, h, color, corner=0.05, transparency=None):
    s = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, l, t, w, h)
    fill(s, color, transparency)
    try:
        s.adjustments[0] = corner
    except Exception:
        pass
    return s


def text(slide, l, t, w, h, content, *, size=14, bold=False, color=INK,
         align=PP_ALIGN.LEFT, anchor=None, italic=False, font=None, wrap=True):
    box = slide.shapes.add_textbox(l, t, w, h)
    tf = box.text_frame
    tf.word_wrap = wrap
    tf.margin_left = Emu(0)
    tf.margin_right = Emu(0)
    tf.margin_top = Emu(0)
    tf.margin_bottom = Emu(0)
    if anchor is not None:
        tf.vertical_anchor = anchor
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = content
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.italic = italic
    r.font.color.rgb = color
    if font:
        r.font.name = font
    return box


def bullets(slide, l, t, w, h, items, *, size=15, color=BODY,
            bullet=CORAL, space=10):
    box = slide.shapes.add_textbox(l, t, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = Emu(0)
    tf.margin_right = Emu(0)
    tf.margin_top = Emu(0)
    tf.margin_bottom = Emu(0)
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(space)
        r1 = p.add_run()
        r1.text = "•  "
        r1.font.size = Pt(size)
        r1.font.bold = True
        r1.font.color.rgb = bullet
        r2 = p.add_run()
        r2.text = item
        r2.font.size = Pt(size)
        r2.font.color.rgb = color


def image_overlay(slide, img: Path, transparency=0.55, color=NAVY_DARK):
    slide.shapes.add_picture(str(img), 0, 0, width=SLIDE_W, height=SLIDE_H)
    rect(slide, 0, 0, SLIDE_W, SLIDE_H, color, transparency=transparency)


def notes(slide, content):
    slide.notes_slide.notes_text_frame.text = content


# ---------------------------------------------------------------------------
# Page chrome
# ---------------------------------------------------------------------------
def page(prs, *, title, subtitle=None, eyebrow=None, eyebrow_color=CORAL):
    s = blank(prs)
    bg(s, CREAM)
    # Small accent above the title
    rect(s, Inches(0.7), Inches(0.6), Inches(0.6), Inches(0.07), eyebrow_color)
    if eyebrow:
        text(s, Inches(0.7), Inches(0.72), Inches(11.5), Inches(0.3),
             eyebrow.upper(), size=11, bold=True, color=eyebrow_color)
        t_top = Inches(1.02)
    else:
        t_top = Inches(0.78)
    text(s, Inches(0.7), t_top, Inches(12), Inches(0.9),
         title, size=32, bold=True, color=NAVY)
    if subtitle:
        text(s, Inches(0.7), t_top + Inches(0.95), Inches(12), Inches(0.45),
             subtitle, size=14, color=BODY)
    return s


def footer(slide, idx, total):
    rect(slide, Inches(0.7), Inches(7.05), Inches(11.93), Inches(0.012), LINE)
    text(slide, Inches(0.7), Inches(7.13), Inches(8), Inches(0.25),
         "Moein Razavi  ·  Portfolio for Ian Jiang",
         size=9, color=MUTED)
    text(slide, Inches(11.4), Inches(7.13), Inches(1.4), Inches(0.25),
         f"{idx} / {total}", size=9, color=MUTED, align=PP_ALIGN.RIGHT)


# ---------------------------------------------------------------------------
# Pipeline / agent flow visual
# ---------------------------------------------------------------------------
def agent_flow(slide, steps, *, top, left=Inches(0.7),
               total_width=Inches(11.93), height=Inches(1.4),
               accent=CORAL):
    """Horizontal numbered flow: each step has a title + 1-line description."""
    n = len(steps)
    gap = Inches(0.16)
    box_w = (total_width - gap * (n - 1)) / n
    for i, (title, desc) in enumerate(steps):
        x = left + i * (box_w + gap)
        round_rect(slide, x, top, box_w, height, WHITE, corner=0.07)
        rect(slide, x, top, box_w, Inches(0.07), accent)
        # Number badge
        round_rect(slide, x + Inches(0.13), top + Inches(0.2),
                   Inches(0.38), Inches(0.38), accent, corner=0.5)
        text(slide, x + Inches(0.13), top + Inches(0.2), Inches(0.38), Inches(0.38),
             str(i + 1), size=11, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        # Title
        text(slide, x + Inches(0.6), top + Inches(0.2),
             box_w - Inches(0.75), Inches(0.35),
             title, size=11, bold=True, color=NAVY)
        # Description
        text(slide, x + Inches(0.18), top + Inches(0.65),
             box_w - Inches(0.36), height - Inches(0.75),
             desc, size=8.5, color=BODY)
        # Arrow to next box
        if i < n - 1:
            arrow = slide.shapes.add_shape(
                MSO_SHAPE.RIGHT_ARROW,
                x + box_w + Inches(0.005),
                top + height / 2 - Inches(0.1),
                gap - Inches(0.01),
                Inches(0.2),
            )
            fill(arrow, accent)


# ---------------------------------------------------------------------------
# Slide builders
# ---------------------------------------------------------------------------
def slide_title(prs, images):
    s = blank(prs)
    if "hero" in images:
        image_overlay(s, images["hero"], transparency=0.58, color=NAVY_DARK)
    else:
        bg(s, NAVY_DARK)

    rect(s, 0, Inches(2.55), Inches(0.45), Inches(0.18), CORAL)

    text(s, Inches(0.85), Inches(2.4), Inches(11), Inches(0.4),
         "PORTFOLIO  ·  PREPARED FOR IAN JIANG", size=13, bold=True, color=GOLD)
    text(s, Inches(0.85), Inches(2.95), Inches(11.5), Inches(1.3),
         "Hi, I'm Moein.", size=56, bold=True, color=WHITE)
    text(s, Inches(0.85), Inches(4.25), Inches(11.5), Inches(0.6),
         "I build AI systems that work in the real world.",
         size=22, color=SOFT_WHITE)
    text(s, Inches(0.85), Inches(5.0), Inches(11.5), Inches(0.45),
         "Today I'd love to walk you through five of them.",
         size=16, italic=True, color=GOLD)
    text(s, Inches(0.85), Inches(6.5), Inches(11.5), Inches(0.35),
         "razavi.moein94@gmail.com  ·  linkedin.com/in/moein-razavi",
         size=11, color=SOFT_WHITE)

    notes(s, "Open warm. Thank Ian. Tell him you'd like to walk through five real systems.")


def slide_about(prs, images, idx, total):
    s = page(prs, title="A little about me",
             subtitle="Short version, in plain English.")
    points = [
        "I'm a Senior AI/ML Engineer with a Ph.D. from Texas A&M.",
        "For the last few years I've been at Ford building production AI systems.",
        "Today I'd like to walk you through five of them.",
        "Each is a real system in production. Each tells a clear story.",
    ]
    bullets(s, Inches(0.7), Inches(2.4), Inches(7.5), Inches(4.4),
            points, size=17, color=INK, bullet=CORAL, space=14)
    if "team" in images:
        s.shapes.add_picture(str(images["team"]),
                             Inches(8.6), Inches(2.4),
                             width=Inches(4.1), height=Inches(4.4))
    footer(s, idx, total)
    notes(s, "Keep this short. Speak the spirit; don't read the bullets.")


def slide_numbers(prs, images, idx, total):
    s = page(prs, title="A few numbers, fast",
             subtitle="Things I've actually shipped to production at Ford.")
    metrics = [
        ("5", "Systems", "End-to-end AI systems I've built and own at Ford.", CORAL),
        ("10", "US States", "Where my chatbot is in production today.", NAVY),
        ("35%", "Faster", "End-to-end latency cut on the chatbot.", TEAL),
        ("99.9%", "Uptime", "On my production AI platform on GCP.", GOLD),
    ]
    gap = Inches(0.22)
    card_w = (Inches(11.93) - gap * 3) / 4
    top = Inches(2.6)
    for i, (val, label, desc, accent) in enumerate(metrics):
        x = Inches(0.7) + i * (card_w + gap)
        round_rect(s, x, top, card_w, Inches(3.4), WHITE, corner=0.06)
        rect(s, x, top, card_w, Inches(0.08), accent)
        text(s, x + Inches(0.3), top + Inches(0.4), card_w - Inches(0.6), Inches(1.1),
             val, size=42, bold=True, color=accent)
        text(s, x + Inches(0.3), top + Inches(1.55), card_w - Inches(0.6), Inches(0.4),
             label, size=14, bold=True, color=INK)
        text(s, x + Inches(0.3), top + Inches(2.0), card_w - Inches(0.6), Inches(1.3),
             desc, size=12, color=BODY)
    footer(s, idx, total)


def slide_five_stories(prs, images, idx, total):
    s = page(prs, title="Five stories I'd love to share",
             subtitle="Each one is a real Ford system. Each has a clear before and after.")

    stories = [
        ("Story 1", "ESA-FR Chatbot",
         "Multi-agent chatbot for Ford EV analytics (router + SQL + docs + checker).", CORAL),
        ("Story 2", "HOMELOAD",
         "Turns any utility file (CSV/Excel/PDF/XML) into one clean format.", NAVY),
        ("Story 3", "V2G Dispatch",
         "Real-time system that uses Ford EVs as a fleet power source for the grid.", TEAL),
        ("Story 4", "CPA Action Prioritizer",
         "4-stage ADK agent that scores actions against a rubric.", PLUM),
        ("Story 5", "Code Complexity Analyzer",
         "Reads a GitHub repo, rates Big-O, estimates Vertex AI cost.", GOLD),
    ]
    gap = Inches(0.18)
    card_w = (Inches(11.93) - gap * 4) / 5
    top = Inches(2.55)
    h = Inches(3.9)
    for i, (label, title, desc, accent) in enumerate(stories):
        x = Inches(0.7) + i * (card_w + gap)
        round_rect(s, x, top, card_w, h, WHITE, corner=0.06)
        rect(s, x, top, card_w, Inches(0.08), accent)
        text(s, x + Inches(0.25), top + Inches(0.3), card_w - Inches(0.5), Inches(0.4),
             label.upper(), size=10, bold=True, color=accent)
        text(s, x + Inches(0.25), top + Inches(0.75), card_w - Inches(0.5), Inches(1.4),
             title, size=15, bold=True, color=NAVY)
        text(s, x + Inches(0.25), top + Inches(2.1), card_w - Inches(0.5), h - Inches(2.25),
             desc, size=12, color=BODY)
    footer(s, idx, total)
    notes(s, "Preview the five stories. Tell Ian he can stop me on any.")


# ---------------------------------------------------------------------------
# Story slide (Problem -> Built -> Changed) — uses correct total story count
# ---------------------------------------------------------------------------
def story_slide(prs, *, idx, total, story_num, total_stories,
                story_title, problem, solution, result,
                accent):
    s = page(prs,
             eyebrow=f"Story {story_num} of {total_stories}",
             eyebrow_color=accent,
             title=story_title,
             subtitle=None)

    col_top = Inches(2.3)
    col_h = Inches(3.7)
    gap = Inches(0.25)
    col_w = (Inches(11.93) - gap * 2) / 3

    steps = [
        ("THE PROBLEM", problem, CORAL),
        ("WHAT I BUILT", solution, NAVY),
        ("WHAT CHANGED", result, TEAL),
    ]
    for i, (label, body, color) in enumerate(steps):
        x = Inches(0.7) + i * (col_w + gap)
        round_rect(s, x, col_top, col_w, col_h, WHITE, corner=0.05)
        rect(s, x, col_top, col_w, Inches(0.08), color)
        round_rect(s, x + Inches(0.35), col_top + Inches(0.3),
                   Inches(0.45), Inches(0.45), color, corner=0.5)
        text(s, x + Inches(0.35), col_top + Inches(0.3), Inches(0.45), Inches(0.45),
             str(i + 1), size=14, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        text(s, x + Inches(0.95), col_top + Inches(0.34),
             col_w - Inches(1.1), Inches(0.4),
             label, size=11, bold=True, color=color)
        text(s, x + Inches(0.35), col_top + Inches(0.95),
             col_w - Inches(0.7), col_h - Inches(1.1),
             body, size=13, color=INK)

    # Continue indicator
    text(s, 0, Inches(6.2), SLIDE_W, Inches(0.35),
         "Next slide: how it actually works  ↓",
         size=11, italic=True, color=accent, align=PP_ALIGN.CENTER)
    footer(s, idx, total)


# ---------------------------------------------------------------------------
# Detail slide (workflow + tech stack)
# ---------------------------------------------------------------------------
def detail_slide(prs, *, idx, total, story_num, total_stories,
                 story_title, flow_steps, components, stack_groups,
                 accent):
    """One slide with workflow pipeline + components + tech stack."""
    s = page(prs,
             eyebrow=f"Story {story_num} · Under the hood",
             eyebrow_color=accent,
             title=story_title,
             subtitle="How it works inside, in plain English.")

    # Workflow pipeline at the top
    text(s, Inches(0.7), Inches(2.15), Inches(8), Inches(0.3),
         "THE WORKFLOW", size=11, bold=True, color=accent)
    agent_flow(s, flow_steps, top=Inches(2.45), height=Inches(1.45),
               accent=accent)

    # Left card: what each part does
    left_x = Inches(0.7)
    bot_top = Inches(4.2)
    bot_h = Inches(2.65)
    left_w = Inches(7.2)
    round_rect(s, left_x, bot_top, left_w, bot_h, WHITE, corner=0.05)
    rect(s, left_x, bot_top, Inches(0.08), bot_h, accent)
    text(s, left_x + Inches(0.3), bot_top + Inches(0.2),
         left_w - Inches(0.5), Inches(0.35),
         "WHAT EACH PART DOES", size=11, bold=True, color=accent)
    # Component list
    box = s.shapes.add_textbox(left_x + Inches(0.3), bot_top + Inches(0.6),
                                left_w - Inches(0.5), bot_h - Inches(0.75))
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = Emu(0); tf.margin_right = Emu(0)
    tf.margin_top = Emu(0); tf.margin_bottom = Emu(0)
    for i, (name, desc) in enumerate(components):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.space_after = Pt(4)
        r1 = p.add_run()
        r1.text = f"{name}  "
        r1.font.size = Pt(11)
        r1.font.bold = True
        r1.font.color.rgb = NAVY
        r2 = p.add_run()
        r2.text = desc
        r2.font.size = Pt(10.5)
        r2.font.color.rgb = BODY

    # Right card: tech stack
    right_x = left_x + left_w + Inches(0.25)
    right_w = Inches(11.93) - left_w - Inches(0.25)
    round_rect(s, right_x, bot_top, right_w, bot_h, WHITE, corner=0.05)
    rect(s, right_x, bot_top, Inches(0.08), bot_h, accent)
    text(s, right_x + Inches(0.3), bot_top + Inches(0.2),
         right_w - Inches(0.5), Inches(0.35),
         "TECH STACK", size=11, bold=True, color=accent)
    sb = s.shapes.add_textbox(right_x + Inches(0.3), bot_top + Inches(0.6),
                               right_w - Inches(0.5), bot_h - Inches(0.75))
    stf = sb.text_frame
    stf.word_wrap = True
    stf.margin_left = Emu(0); stf.margin_right = Emu(0)
    stf.margin_top = Emu(0); stf.margin_bottom = Emu(0)
    for i, (group, items) in enumerate(stack_groups):
        if i == 0:
            p = stf.paragraphs[0]
        else:
            p = stf.add_paragraph()
        p.space_after = Pt(5)
        r1 = p.add_run()
        r1.text = f"{group}\n"
        r1.font.size = Pt(10)
        r1.font.bold = True
        r1.font.color.rgb = accent
        r2 = p.add_run()
        r2.text = items
        r2.font.size = Pt(10)
        r2.font.color.rgb = INK

    footer(s, idx, total)


# ---------------------------------------------------------------------------
# CPA walked-through example slide
# ---------------------------------------------------------------------------
def slide_cpa_example(prs, *, idx, total):
    """A concrete walked-through example of the CPA agent on a real-shape problem."""
    accent = ACCENT_BY_PROJECT[4]  # PLUM
    s = page(prs,
             eyebrow="Story 4 · A walked-through example",
             eyebrow_color=accent,
             title="Ranking 200 Ford dealers for a $5M EV incentive",
             subtitle="What you'd upload, what the agent does, what you'd get back.")

    LEFT = Inches(0.7)
    TOTAL_W = Inches(11.93)
    gap = Inches(0.18)

    # ---- SECTION 1: WHAT YOU UPLOAD ----
    text(s, LEFT, Inches(2.2), TOTAL_W, Inches(0.25),
         "1.  WHAT YOU UPLOAD", size=10, bold=True, color=accent)

    inputs = [
        ("dealer_data_q4.xlsx",
         "3 tables on one sheet — EV sales, service capacity, CSAT (200 dealers)"),
        ("ev_incentive_rubric.xlsx",
         "4 dimensions with weights and plain-English scoring rules"),
        ("ev_strategy_2025.docx",
         "3-page memo on which states + investments to prioritize"),
    ]
    card_w = (TOTAL_W - gap * 2) / 3
    card_top = Inches(2.50)
    card_h = Inches(0.95)
    for i, (name, desc) in enumerate(inputs):
        x = LEFT + i * (card_w + gap)
        round_rect(s, x, card_top, card_w, card_h, WHITE, corner=0.06)
        rect(s, x, card_top, Inches(0.08), card_h, accent)
        text(s, x + Inches(0.22), card_top + Inches(0.12),
             card_w - Inches(0.4), Inches(0.3),
             name, size=11, bold=True, color=NAVY)
        text(s, x + Inches(0.22), card_top + Inches(0.42),
             card_w - Inches(0.4), card_h - Inches(0.5),
             desc, size=9.5, color=BODY)

    # ---- SECTION 2: THE PLAN + APPROVAL ----
    sec2_label_top = Inches(3.65)
    text(s, LEFT, sec2_label_top, TOTAL_W, Inches(0.25),
         "2.  THE AGENT WRITES A PLAN AND WAITS FOR YOU TO SAY 'GO'",
         size=10, bold=True, color=accent)
    plan_top = Inches(3.95)
    plan_h = Inches(1.55)
    round_rect(s, LEFT, plan_top, TOTAL_W, plan_h, WHITE, corner=0.05)
    rect(s, LEFT, plan_top, Inches(0.08), plan_h, accent)

    box = s.shapes.add_textbox(LEFT + Inches(0.3), plan_top + Inches(0.18),
                                TOTAL_W - Inches(0.5), plan_h - Inches(0.3))
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = Emu(0); tf.margin_right = Emu(0)
    tf.margin_top = Emu(0); tf.margin_bottom = Emu(0)

    p = tf.paragraphs[0]
    p.space_after = Pt(3)
    r = p.add_run(); r.text = "Agent:  "
    r.font.size = Pt(10); r.font.bold = True; r.font.color.rgb = accent
    r = p.add_run()
    r.text = ("\"Here's my plan: join 3 tables on dealer_id. "
              "ev_momentum = EV ÷ total units. service_ready = 1 if EV techs ≥ 2 AND chargers ≥ 4. "
              "cx_ok = 1 if CSAT ≥ 4.0 AND complaints < 10. regional_bonus = 0.15 if state is CA/TX/FL/NY. "
              "Final score = 0.35·ev_momentum + 0.30·service_ready + 0.20·cx_ok + 0.15·regional_bonus. "
              "Return top 30 with per-dimension breakdown. OK?\"")
    r.font.size = Pt(10); r.font.color.rgb = INK

    p = tf.add_paragraph()
    p.space_after = Pt(3)
    r = p.add_run(); r.text = "You:  "
    r.font.size = Pt(10); r.font.bold = True; r.font.color.rgb = accent
    r = p.add_run(); r.text = "\"Looks good. Proceed.\""
    r.font.size = Pt(10); r.font.color.rgb = INK

    p = tf.add_paragraph()
    r = p.add_run(); r.text = "Agent:  "
    r.font.size = Pt(10); r.font.bold = True; r.font.color.rgb = accent
    r = p.add_run()
    r.text = ("walks the plan step by step. For each step, calls the Pandas helper — "
              "PandasAI translates the English into pandas code, runs it on the local "
              "dataframes, and reports a snippet of the partial result.")
    r.font.size = Pt(10); r.font.italic = True; r.font.color.rgb = BODY

    # ---- SECTION 3: WHAT YOU GET BACK ----
    sec3_label_top = Inches(5.7)
    text(s, LEFT, sec3_label_top, TOTAL_W, Inches(0.25),
         "3.  WHAT YOU GET BACK", size=10, bold=True, color=accent)
    outs = [
        ("ev_incentive_ranking.csv",
         "200 dealers sorted descending, with per-dimension scores + final rank."),
        ("ev_incentive_report.md",
         "9 sections: exec summary, methodology, top 30, findings, caveats. Delivered as a Google Cloud Storage download link via the MCP server."),
    ]
    out_top = Inches(6.0)
    out_h = Inches(1.0)
    out_w = (TOTAL_W - gap) / 2
    for i, (name, desc) in enumerate(outs):
        x = LEFT + i * (out_w + gap)
        round_rect(s, x, out_top, out_w, out_h, WHITE, corner=0.06)
        rect(s, x, out_top, Inches(0.08), out_h, accent)
        text(s, x + Inches(0.25), out_top + Inches(0.13),
             out_w - Inches(0.45), Inches(0.3),
             name, size=11, bold=True, color=NAVY)
        text(s, x + Inches(0.25), out_top + Inches(0.43),
             out_w - Inches(0.45), out_h - Inches(0.53),
             desc, size=9.5, color=BODY)

    footer(s, idx, total)
    notes(s, "Use if Ian asks 'show me a concrete example'. Walk left → middle → right.")


# ---------------------------------------------------------------------------
# Closing slides
# ---------------------------------------------------------------------------
def slide_lessons(prs, idx, total):
    s = page(prs, title="What ties these projects together",
             subtitle="A few simple ideas behind every system I build.")
    lessons = [
        ("Start with the user's problem",
         "Not with the model. Models are a tool, not the point.",
         CORAL),
        ("Ship the whole thing",
         "A model in a notebook is not a product. I own end-to-end.",
         NAVY),
        ("Plan for things to go wrong",
         "Retries, fallbacks, monitoring, and clear ways to recover.",
         TEAL),
        ("Measure what really matters",
         "Latency, cost, accuracy — and whether the business wins.",
         GOLD),
    ]
    gap_x = Inches(0.25)
    gap_y = Inches(0.22)
    card_w = (Inches(11.93) - gap_x) / 2
    card_h = Inches(2.0)
    top = Inches(2.4)
    for i, (title, desc, accent) in enumerate(lessons):
        col = i % 2
        row = i // 2
        x = Inches(0.7) + col * (card_w + gap_x)
        y = top + row * (card_h + gap_y)
        round_rect(s, x, y, card_w, card_h, WHITE, corner=0.06)
        rect(s, x, y, Inches(0.08), card_h, accent)
        text(s, x + Inches(0.35), y + Inches(0.3),
             card_w - Inches(0.7), Inches(0.5),
             title, size=17, bold=True, color=NAVY)
        text(s, x + Inches(0.35), y + Inches(0.9),
             card_w - Inches(0.7), card_h - Inches(1.0),
             desc, size=13, color=BODY)
    footer(s, idx, total)


def slide_sgws_fit(prs, idx, total):
    s = page(prs, title="Why this matters for Southern Glazer's",
             subtitle="The same five patterns map cleanly to your business.")
    items = [
        ("Sales assistant", "Like the Ford ESA-FR chatbot — answer reps' questions about accounts, products, and pricing using SQL + document agents, with a checker reviewing every answer.", CORAL),
        ("Clean partner data", "Like HOMELOAD — turn supplier and customer files of any format into one clean, standard source. Plain code first, AI as backup, self-healing retries.", NAVY),
        ("Real-time decisions under uncertainty", "Like the V2G dispatch system — combine live signals, an ensemble forecast, and Monte Carlo to commit to a number with confidence and adapt in real time when things change.", TEAL),
        ("Action prioritization", "Like the CPA Action Prioritizer — an agent reads a rubric and your data, proposes a plan, waits for your approval, then runs it.", PLUM),
        ("Engineering productivity", "Like the code analyzer — let your engineers paste a repo URL and get Big-O analysis + Google Cloud cost estimates in minutes.", GOLD),
    ]
    top = Inches(2.4)
    row_h = Inches(0.85)
    for i, (title, desc, accent) in enumerate(items):
        y = top + i * (row_h + Inches(0.05))
        round_rect(s, Inches(0.7), y, Inches(11.93), row_h, WHITE, corner=0.05)
        rect(s, Inches(0.7), y, Inches(0.6), row_h, accent)
        text(s, Inches(0.7), y, Inches(0.6), row_h, str(i + 1),
             size=18, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        text(s, Inches(1.55), y + Inches(0.12), Inches(2.9), Inches(0.45),
             title, size=14, bold=True, color=NAVY)
        text(s, Inches(4.5), y + Inches(0.18), Inches(8.0), row_h - Inches(0.3),
             desc, size=12, color=BODY)
    footer(s, idx, total)
    notes(s, "Keep humble. These are patterns from work I've done, not claims about SGWS internals.")


def slide_questions(prs, idx, total):
    s = page(prs, title="A few things I'd love to ask you",
             subtitle="To learn how your team works.")
    qs = [
        "What are the hardest technical problems your team is solving right now?",
        "What does your AI setup look like today?",
        "What does a great first few months in this role look like?",
        "What kinds of projects would I likely jump into first?",
        "How does the team balance trying new things with keeping things stable?",
    ]
    top = Inches(2.4)
    row_h = Inches(0.78)
    for i, q in enumerate(qs):
        y = top + i * (row_h + Inches(0.06))
        round_rect(s, Inches(0.7), y, Inches(11.93), row_h, WHITE, corner=0.06)
        rect(s, Inches(0.7), y, Inches(0.7), row_h, CORAL if i % 2 == 0 else TEAL)
        text(s, Inches(0.7), y, Inches(0.7), row_h, str(i + 1),
             size=18, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        text(s, Inches(1.65), y + Inches(0.2), Inches(10.1), row_h - Inches(0.3),
             q, size=15, color=INK)
    footer(s, idx, total)


def slide_thanks(prs, images):
    s = blank(prs)
    if "hero" in images:
        image_overlay(s, images["hero"], transparency=0.62, color=NAVY_DARK)
    else:
        bg(s, NAVY_DARK)
    rect(s, Inches(6.0), Inches(2.0), Inches(1.3), Inches(0.18), CORAL)
    text(s, 0, Inches(2.4), SLIDE_W, Inches(1.2),
         "Thank you.", size=58, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    text(s, 0, Inches(3.7), SLIDE_W, Inches(0.5),
         "I'd love to hear what you think — and answer any questions.",
         size=18, color=GOLD, align=PP_ALIGN.CENTER)
    card_w = Inches(8.0)
    cx = (SLIDE_W - card_w) / 2
    round_rect(s, cx, Inches(4.7), card_w, Inches(1.85), NAVY_DARK, corner=0.04)
    text(s, cx, Inches(4.88), card_w, Inches(0.4), "MOEIN RAZAVI",
         size=12, bold=True, color=GOLD, align=PP_ALIGN.CENTER)
    text(s, cx, Inches(5.25), card_w, Inches(0.4),
         "razavi.moein94@gmail.com", size=14, color=WHITE, align=PP_ALIGN.CENTER)
    text(s, cx, Inches(5.65), card_w, Inches(0.4),
         "(979) 676-7486", size=13, color=SOFT_WHITE, align=PP_ALIGN.CENTER)
    text(s, cx, Inches(6.05), card_w, Inches(0.4),
         "linkedin.com/in/moein-razavi  ·  github.com/moeinrazavi",
         size=12, color=SOFT_WHITE, align=PP_ALIGN.CENTER)
    notes(s, "Close warm. Invite questions. Pause and listen.")


# ---------------------------------------------------------------------------
# The five projects — content
# ---------------------------------------------------------------------------
PROJECT_1_CHATBOT = dict(
    accent=ACCENT_BY_PROJECT[1],
    story_title="An AI assistant for Ford EV teams",
    problem=(
        "Ford EV analytics teams across 10 US states had to dig through "
        "different BigQuery tables and PDF manuals just to answer simple "
        "questions about EV pilot programs, rate plans, or F-150 Lightning "
        "specs. Every new question meant an engineering ticket."
    ),
    solution=(
        "I built a multi-agent chatbot. A smart router (rules + LLM) reads "
        "the question and picks 1 of 6 routes. A SQL Agent writes BigQuery "
        "for analytics questions. A Document Agent searches the F-150 manual + "
        "specs + EV docs. A Small Talk Agent handles greetings. A Checker "
        "agent reviews every answer for domain rules before it goes out."
    ),
    result=(
        "End-to-end latency cut by 35%. "
        "In daily use across 10 US states. "
        "Teams self-serve — no more tickets, and every answer is grounded "
        "in real data, not invented."
    ),
    flow_steps=[
        ("Small Talk", "Greetings handled by templates (no AI)"),
        ("Intent Proc.", "Translates 'my truck' to 'Ford F-150'"),
        ("Query Proc.", "Normalizes EV synonyms (V2H, V2G, etc.)"),
        ("LLM Decider", "Rules + AI pick 1 of 6 routes"),
        ("SQL or Doc Agent", "Specialist runs the query"),
        ("Checker", "Grades the answer for domain rules"),
    ],
    components=[
        ("LLM Decider", " — hybrid rules + LLM router. Picks 1 of 6 labels (table, F-150 specs, F-150 manual, general doc, small talk, irrelevant)."),
        ("SQL Agent", " — turns the question into a BigQuery query using the live schema + 15+ example queries. Applies the weekly→monthly 4.33 rule."),
        ("Document Agent", " — keyword search (BM25) over 3 documents: F-150 Lightning specs, 2023 F-150 Owner's Manual, and a general EV doc."),
        ("Small Talk Agent", " — handles 'hi' and 'thanks' with text templates. No AI call needed."),
        ("Checker", " — a separate LLM grades the answer against domain rules. Returns VERIFIED or NEEDS_CORRECTION."),
        ("Context Manager", " — remembers the last 20 messages, the user's state, treatment, and rate."),
    ],
    stack_groups=[
        ("Language & framework", "Python · LangChain · Dash UI"),
        ("Models", "Ford LLM API gateway (GPT-4o + GPT-4o-mini)"),
        ("Data & retrieval", "BigQuery · BM25 keyword search over PDFs"),
        ("Cloud", "GCP · Cloud Run · SAML/ADFS auth"),
    ],
)

PROJECT_2_HOMELOAD = dict(
    accent=ACCENT_BY_PROJECT[2],
    story_title="HOMELOAD — turning messy utility files into clean data",
    problem=(
        "Utility companies (DTE, PG&E, ConEd, and many more) send Ford "
        "customer energy data in completely different file formats — CSV, "
        "Excel, PDF, Green Button XML, URL downloads. Each utility used to "
        "need its own hand-written parser (~500 lines of brittle code). "
        "Adding a new utility took weeks."
    ),
    solution=(
        "I built a multi-agent FastAPI service. File Reader → Intake → "
        "Schema Discovery → Extraction → Validation. Plain code tries to "
        "read the file first; the AI only steps in if plain code fails. "
        "If confidence is low, the orchestrator feeds the problems back as "
        "hints and retries (up to 2 times). Output is always the same 7 "
        "standard columns, written safely to BigQuery."
    ),
    result=(
        "Replaced ~500 lines of utility-specific parsers with one universal "
        "self-healing pipeline. Ingestion time cut 25%. Adding a new utility "
        "is now days, not weeks. Same clean format every time."
    ),
    flow_steps=[
        ("File Reader", "Reads CSV/Excel/PDF/XML/URL into raw text"),
        ("Intake", "Decides: is this electric? timezone?"),
        ("Schema", "Plans which columns to extract"),
        ("Extract", "Plain code first; AI as backup"),
        ("Validate", "Scores result; lists what's wrong"),
        ("Reflection", "Retries with problems fed back (max 2)"),
    ],
    components=[
        ("File Reader", " — plain code. Handles CSV, Excel, PDF, XML, HTML, URLs. No AI call."),
        ("Intake Agent", " — Gemini reads a smart sample. Decides: is it electric? what utility? timezone? interval?"),
        ("Schema Discovery Agent", " — Gemini builds an extraction plan (column mappings, units, date format, ESPI config for Green Button XML)."),
        ("Extraction Agent", " — tries pandas/lxml first. Only if row count is too low does it ask the AI for a sample — then pandas still parses the bulk."),
        ("Validation Agent", " — deterministic checks (no AI). Scores schema, timestamps, ranges, coverage."),
        ("Orchestrator", " — runs up to 3 attempts. Always keeps the best attempt — never throws it away."),
    ],
    stack_groups=[
        ("Language & framework", "Python · FastAPI · Pydantic"),
        ("AI", "Vertex AI · Gemini 2.0 Flash (google-genai SDK)"),
        ("Storage", "Cloud Storage · BigQuery (MERGE = safe re-runs)"),
        ("Deployment", "Docker · Cloud Run · Terraform"),
    ],
)

PROJECT_3_ROUTINE = dict(
    accent=ACCENT_BY_PROJECT[3],
    story_title="V2G — using Ford EVs as a fleet power source for the grid",
    problem=(
        "When the power grid needs extra electricity (peak hours, hot "
        "evenings), Ford's EVs could sell some back from their batteries "
        "(this is V2G — vehicle-to-grid). The hard part: how do you commit "
        "a fixed amount of power to the grid hours ahead, when you don't "
        "know which cars will be home, plugged in, and charged enough? "
        "Guess wrong and Ford pays a penalty. Guess too conservatively and "
        "you leave money on the table."
    ),
    solution=(
        "I built an end-to-end real-time V2G dispatch system. It pulls live "
        "ERCOT electricity prices, predicts which cars will be home using "
        "a forecasting ensemble I trained (Chronos-2 + iTransformer + "
        "Copy-Last-Week, best model picked per car), runs ~3,000 Monte "
        "Carlo scenarios per dispatch cycle, and picks the smallest car "
        "pool that meets the commitment at 90% confidence — with a "
        "Dropout Compensator that instantly redistributes load if a car leaves."
    ),
    result=(
        "Sub-second dispatch decisions. Commitment held even when cars "
        "drop out. Forecasts 10% better than baseline. Training 40% faster. "
        "Answers 'I need 30 kWh/hr from 6-10pm — which cars?' with a "
        "ranked VIN list and the reasoning."
    ),
    flow_steps=[
        ("Market Feed", "Live ERCOT prices via gridstatus"),
        ("Forecast", "Predict which cars are home (Chronos-2 + iTransformer + TFT)"),
        ("Monte Carlo", "~3,000 scenarios → P50/P90 capacity"),
        ("Bid Optimizer", "Pick smallest VIN pool that meets target at 90%"),
        ("Power Control", "Send per-car kW commands"),
        ("Compensate", "Redistribute instantly when a car drops out"),
    ],
    components=[
        ("MarketFeed", " — pulls live ERCOT prices via the gridstatus library. Synthetic-prices fallback for offline runs."),
        ("Forecasting ensemble", " — Chronos-2 + iTransformer + Copy-Last-Week + my trained TFT (TFT/Autoformer trained on 444 cars with Optuna on Vertex AI V100s). Picks the best model per car. Outputs 9 quantile forecasts, 168 hours ahead."),
        ("BidOptimizer", " — turns 'I need 30 kWh from 6-10pm at 90% confidence' into a ranked car list. Greedy selection + Monte Carlo scoring (~3k scenarios per cycle)."),
        ("PowerController + FleetMonitor", " — sends per-VIN kW commands; watches aggregated + per-car output every tick; flags deviation."),
        ("DropoutCompensator", " — when a car unexpectedly leaves, instantly redistributes load or pulls in reserve cars. Keeps the grid commitment alive."),
    ],
    stack_groups=[
        ("Language & framework", "Python · PyTorch · CUDA"),
        ("Forecasting models", "Chronos-2 · iTransformer · GluonTS TFT · Autoformer"),
        ("Market & decisioning", "gridstatus (ERCOT) · Monte Carlo (~3k/cycle) · P50/P90"),
        ("Cloud & data", "Vertex AI Workbench (V100 GPU) · BigQuery · Optuna"),
    ],
)

PROJECT_4_CPA = dict(
    accent=ACCENT_BY_PROJECT[4],
    story_title="CPA Action Prioritizer — a smart business assistant",
    problem=(
        "When you've got a business spreadsheet and a scoring rubric, "
        "deciding which actions to take first is hard. Doing it in Excel "
        "is slow, inconsistent, and easy to get wrong. Teams need a way "
        "to score actions against a rubric in plain English — and they "
        "need to approve the plan before any analysis runs."
    ),
    solution=(
        "I built a 4-stage agent app on Google's ADK. The user uploads "
        "files (Excel, CSV, Word) and a rubric. The Input agent reads "
        "them. The Planning agent proposes a scoring plan and WAITS for "
        "user approval. The Execution agent walks the plan as a checklist, "
        "calling a Pandas helper agent (PandasAI on local dataframes) for "
        "the actual analysis. The Report agent writes a 9-section markdown "
        "report and uploads it to Cloud Storage via an MCP server."
    ),
    result=(
        "Teams get a ranked list of actions with the reasoning behind "
        "every score. Works on any rubric. Hard user-approval gate before "
        "anything runs — nothing happens until you say go."
    ),
    flow_steps=[
        ("Input", "Reads uploaded Excel, CSV, Word files"),
        ("Planning", "Proposes a scoring plan from the rubric"),
        ("Approval", "Hard gate — user must say 'proceed'"),
        ("Execution", "Walks the plan as a checklist"),
        ("Pandas Helper", "English → pandas code (PandasAI)"),
        ("Report", "9-section markdown + CSV → Cloud Storage"),
    ],
    components=[
        ("Input Agent", " — ingests Excel/CSV via a multi-table parser and Word via the mammoth library. Stores as in-memory tables."),
        ("Planning Agent", " — reads the rubric + data, proposes scoring metrics, exact tables/columns, and step dependencies. Does NOT run anything."),
        ("Execution Agent", " — runs the approved plan as a hierarchical to-do list with status markers (open / in-progress / done)."),
        ("Pandas Agent", " — child agent. English → pandas code via PandasAI on the LOCAL uploaded dataframes (not BigQuery). Saves working code as snippets."),
        ("Report Agent + MCP", " — writes a 9-section markdown report and uploads outputs via a small MCP server (one tool: create_download_file → signed URL)."),
    ],
    stack_groups=[
        ("Agent framework", "Google ADK · MCP"),
        ("Models", "Gemini 2.5 Flash (root + Pandas) · gemini-2.5-flash-lite for PandasAI"),
        ("Data", "Local pandas dataframes (from uploads) · Cloud Storage for output"),
        ("UI & infra", "Angular (in google-adk) · FastAPI · SQLite session DB"),
    ],
)

PROJECT_5_HACKATHON = dict(
    accent=ACCENT_BY_PROJECT[5],
    story_title="Code Complexity Analyzer + Cloud Cost Estimator",
    problem=(
        "Ford engineers wanted two things fast: is my code efficient, and "
        "how much would it cost to run on Vertex AI? Doing this by hand "
        "for a whole repo was painful — and there was no easy way to "
        "compare different LLM models for the same code."
    ),
    solution=(
        "I built a FastAPI web app. Paste a GitHub URL, pick from 21 LLM "
        "models (all routed through Ford's LLM API gateway). The app clones "
        "the repo, uses Python's AST module to find each function, class, "
        "and loop, sends each block to the AI for Big-O analysis + "
        "improvement suggestions, and streams results to the browser live "
        "via Server-Sent Events. A second AI call estimates Vertex AI cost."
    ),
    result=(
        "A whole repo is analyzed in minutes. Engineers compare models on "
        "the same code. Cost estimator (hybrid AI + static pricing table) "
        "predicts current vs optimized cost on Google Cloud. Inefficient "
        "code gets caught before it ships."
    ),
    flow_steps=[
        ("Clone", "GitPython pulls the GitHub repo"),
        ("AST Walk", "Python ast module finds functions/loops"),
        ("LLM Analyze", "Per-block Big-O + suggestions"),
        ("Summary", "AI writes JSON repo-level summary"),
        ("SSE Stream", "Live results to the browser"),
        ("Cost", "AI picks GCP machine; Python prices it"),
    ],
    components=[
        ("Repo Cloner", " — GitPython safely clones any repo to a temp folder."),
        ("AST Extractor", " — Python's ast module finds every function, class, and for/while loop."),
        ("LLM Analyzer", " — per-block call to Ford LLM API (default: gpt-4.1-mini, temperature 0). Returns 6 templated lines."),
        ("Repo Summarizer", " — second AI call. Returns JSON: purpose, tech stack, complexity, compute hours, confidence."),
        ("SSE Streamer", " — FastAPI + async queue + thread pool. Streams every result to the browser as it lands."),
        ("Cost Estimator", " — hybrid: AI picks GCP machine type; Python multiplies by a static 2024 pricing table. Not a real billing API call."),
    ],
    stack_groups=[
        ("Language & framework", "Python · FastAPI · LangChain · Jinja2"),
        ("Models (21)", "GPT-4.1 · Gemini 2.5 · Claude · Llama · DeepSeek · Qwen · all via one endpoint"),
        ("Backend", "Ford LLM API gateway (model picked per request)"),
        ("Streaming & UI", "Server-Sent Events · async queue · thread pool"),
    ],
)


# ---------------------------------------------------------------------------
# Build
# ---------------------------------------------------------------------------
def build(images):
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H

    TOTAL = 19
    TOTAL_STORIES = 5

    # 1-4: framing
    slide_title(prs, images)
    slide_about(prs, images, 2, TOTAL)
    slide_numbers(prs, images, 3, TOTAL)
    slide_five_stories(prs, images, 4, TOTAL)

    # 5-15: five projects × (story + detail), plus a worked example after CPA
    projects = [
        PROJECT_1_CHATBOT,
        PROJECT_2_HOMELOAD,
        PROJECT_3_ROUTINE,
        PROJECT_4_CPA,
        PROJECT_5_HACKATHON,
    ]
    slide_idx = 5
    for n, proj in enumerate(projects, start=1):
        story_slide(
            prs,
            idx=slide_idx, total=TOTAL,
            story_num=n, total_stories=TOTAL_STORIES,
            story_title=proj["story_title"],
            problem=proj["problem"],
            solution=proj["solution"],
            result=proj["result"],
            accent=proj["accent"],
        )
        slide_idx += 1
        detail_slide(
            prs,
            idx=slide_idx, total=TOTAL,
            story_num=n, total_stories=TOTAL_STORIES,
            story_title=proj["story_title"],
            flow_steps=proj["flow_steps"],
            components=proj["components"],
            stack_groups=proj["stack_groups"],
            accent=proj["accent"],
        )
        slide_idx += 1
        # After CPA (Story 4), drop in a concrete worked example
        if n == 4:
            slide_cpa_example(prs, idx=slide_idx, total=TOTAL)
            slide_idx += 1

    # 16-19: closing
    slide_lessons(prs, 16, TOTAL)
    slide_sgws_fit(prs, 17, TOTAL)
    slide_questions(prs, 18, TOTAL)
    slide_thanks(prs, images)

    return prs


def main():
    print("Loading images...")
    images = load_images()
    print(f"Images: {len(images)}")
    prs = build(images)
    prs.save(OUTPUT)
    print(f"Saved: {OUTPUT}  ({len(prs.slides)} slides)")


if __name__ == "__main__":
    main()
