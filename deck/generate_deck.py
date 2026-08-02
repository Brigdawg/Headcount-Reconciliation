#!/usr/bin/env python3
"""BambooHR case study deck — clean product story + market evidence."""

from pathlib import Path
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

ROOT = Path(__file__).resolve().parent
SHOTS = ROOT / "screenshots"
LOGO = SHOTS / "bamboohr-logo.png"

# Bamboo-inspired but quieter
GREEN = RGBColor(0x73, 0xC4, 0x1D)
GREEN_DK = RGBColor(0x4F, 0x8F, 0x12)
INK = RGBColor(0x1B, 0x1F, 0x17)
MUTED = RGBColor(0x6A, 0x72, 0x64)
LINE = RGBColor(0xE6, 0xEA, 0xE1)
SOFT = RGBColor(0xF3, 0xF8, 0xEA)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
WASH = RGBColor(0xFA, 0xFB, 0xF8)

TOTAL = 13
prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)


def font(run, size=18, bold=False, color=INK):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = "Arial"


def text(shape, content, size=18, bold=False, color=INK, align=PP_ALIGN.LEFT):
    tf = shape.text_frame
    tf.clear()
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = content
    font(r, size, bold, color)
    return tf


def para(tf, content, size=14, bold=False, color=INK, before=8, align=PP_ALIGN.LEFT):
    p = tf.add_paragraph()
    p.alignment = align
    p.space_before = Pt(before)
    r = p.add_run()
    r.text = content
    font(r, size, bold, color)
    return p


def slide():
    return prs.slides.add_slide(prs.slide_layouts[6])


def box(s, l, t, w, h, fill=WHITE, line=None):
    sh = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, l, t, w, h)
    sh.fill.solid()
    sh.fill.fore_color.rgb = fill
    if line is None:
        sh.line.fill.background()
    else:
        sh.line.color.rgb = line
    try:
        sh.adjustments[0] = 0.06
    except Exception:
        pass
    return sh


def bar(s):
    sh = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.06))
    sh.fill.solid()
    sh.fill.fore_color.rgb = GREEN
    sh.line.fill.background()


def logo(s, left=Inches(0.75), top=Inches(0.28), width=Inches(1.55)):
    if LOGO.exists():
        s.shapes.add_picture(str(LOGO), left, top, width=width)


def foot(s, n):
    t = s.shapes.add_textbox(Inches(0.75), Inches(7.15), Inches(10), Inches(0.25))
    text(t, "BambooHR interview case study", size=11, color=MUTED)
    p = s.shapes.add_textbox(Inches(12.2), Inches(7.15), Inches(0.8), Inches(0.25))
    text(p, f"{n}/{TOTAL}", size=11, color=MUTED, align=PP_ALIGN.RIGHT)


def notes(s, body):
    s.notes_slide.notes_text_frame.text = body


def head(s, kicker, title, sub=None):
    bar(s)
    logo(s)
    k = s.shapes.add_textbox(Inches(2.55), Inches(0.25), Inches(10), Inches(0.25))
    text(k, kicker.upper(), size=11, bold=True, color=GREEN_DK)
    t = s.shapes.add_textbox(Inches(2.55), Inches(0.5), Inches(10.2), Inches(0.55))
    text(t, title, size=28, bold=True, color=INK)
    if sub:
        u = s.shapes.add_textbox(Inches(2.55), Inches(1.05), Inches(10.2), Inches(0.35))
        text(u, sub, size=15, color=MUTED)


# ---------- 1 Title ----------
s = slide()
wash = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
wash.fill.solid()
wash.fill.fore_color.rgb = WASH
wash.line.fill.background()
accent = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(0.12), Inches(7.5))
accent.fill.solid()
accent.fill.fore_color.rgb = GREEN
accent.line.fill.background()
if LOGO.exists():
    s.shapes.add_picture(str(LOGO), Inches(0.9), Inches(1.7), width=Inches(2.4))
t = s.shapes.add_textbox(Inches(0.9), Inches(2.7), Inches(11), Inches(1.5))
tf = text(t, "Headcount Control Tower", size=42, bold=True, color=INK)
para(tf, "Product discovery case study", size=20, color=MUTED, before=14)
m = s.shapes.add_textbox(Inches(0.9), Inches(4.6), Inches(10), Inches(1.0))
tf = text(m, "Make planned vs actual headcount understandable in 30 seconds.", size=18, color=INK)
para(tf, "AI-augmented discovery  ·  Interactive prototype  ·  2–3 hour timebox", size=14, color=MUTED, before=10)
notes(s, "Lead with the outcome. This deck is process; the live app is the deliverable.")


# ---------- 2 Problem ----------
s = slide()
head(s, "Problem", "Reconciliation is critical. The tools are unreadable.")
# three pain blocks
pains = [
    ("Many sources", "Board plan, roll-forward, HRIS actuals, and approvals live in separate tabs and systems."),
    ("No shared language", "Backfill, net-new, contractor, and “approved” mean different things to Finance and HR."),
    ("No trusted as-of", "Leaders cannot answer headcount at a date without a rebuild — so trust erodes every close."),
]
for i, (h, d) in enumerate(pains):
    left = Inches(0.75 + i * 4.1)
    box(s, left, Inches(1.7), Inches(3.85), Inches(3.6), WHITE, LINE)
    num = s.shapes.add_textbox(left + Inches(0.3), Inches(2.0), Inches(3.2), Inches(0.35))
    text(num, f"0{i+1}", size=14, bold=True, color=GREEN_DK)
    title = s.shapes.add_textbox(left + Inches(0.3), Inches(2.5), Inches(3.2), Inches(0.5))
    text(title, h, size=22, bold=True, color=INK)
    body = s.shapes.add_textbox(left + Inches(0.3), Inches(3.2), Inches(3.2), Inches(1.7))
    text(body, d, size=15, color=MUTED)
foot(s, 2)
notes(s, "Problem is comprehension and alignment — not missing data. Comes from a real Utah Head of Finance workflow.")


# ---------- 3 Why — stakes ----------
s = slide()
head(s, "Why it matters", "People cost is the budget. Misaligned headcount is a strategy tax.")
stats = [
    ("60–80%", "of operating expense\nis people cost", "Aleph / FP&A industry framing, 2026"),
    ("~70%", "of operating costs\nare labor at most firms", "HR Bench on HR–Finance headcount alignment"),
    ("99%", "of FP&A pros use\nspreadsheets monthly", "AFP FP&A Benchmarking Survey, 2025"),
]
for i, (big, line2, src) in enumerate(stats):
    left = Inches(0.75 + i * 4.1)
    box(s, left, Inches(1.7), Inches(3.85), Inches(4.5), WHITE, LINE)
    b = s.shapes.add_textbox(left + Inches(0.3), Inches(2.2), Inches(3.25), Inches(1.0))
    text(b, big, size=44, bold=True, color=GREEN_DK)
    l = s.shapes.add_textbox(left + Inches(0.3), Inches(3.4), Inches(3.25), Inches(1.4))
    text(l, line2, size=16, color=INK)
    c = s.shapes.add_textbox(left + Inches(0.3), Inches(5.3), Inches(3.25), Inches(0.6))
    text(c, src, size=11, color=MUTED)
foot(s, 3)
notes(s, "Cite sources aloud. Point: if people cost dominates opex and spreadsheets dominate FP&A, headcount recon is a first-class product problem for BambooHR’s finance buyers.")


# ---------- 4 Why — who benefits ----------
s = slide()
head(s, "Why it matters", "Who wins when the number is trustworthy")
rows = [
    ("Finance", "One explained variance story at close", "Fewer recon threads · faster board answers"),
    ("HR", "Clear ticket types and approval path", "Backfills don’t masquerade as net-new"),
    ("Managers", "Visible request status", "Less tribal knowledge"),
    ("BambooHR", "A reason for finance to care about HRIS", "“Sell the finance department easier”"),
]
for i, (who, need, out) in enumerate(rows):
    y = Inches(1.65 + i * 1.2)
    box(s, Inches(0.75), y, Inches(2.4), Inches(1.0), SOFT if i % 2 == 0 else WHITE, LINE)
    w = s.shapes.add_textbox(Inches(0.95), y + Inches(0.3), Inches(2.0), Inches(0.45))
    text(w, who, size=16, bold=True, color=INK)
    n = s.shapes.add_textbox(Inches(3.4), y + Inches(0.2), Inches(5.2), Inches(0.7))
    text(n, need, size=16, color=INK)
    o = s.shapes.add_textbox(Inches(8.8), y + Inches(0.2), Inches(3.8), Inches(0.7))
    text(o, out, size=15, bold=True, color=GREEN_DK)
foot(s, 4)
notes(s, "North star for every later decision. If it doesn’t help these people trust or decide, cut it.")


# ---------- 5 Evidence — customer ----------
s = slide()
head(s, "Evidence", "Primary research — Head of Finance")
box(s, Inches(0.75), Inches(1.7), Inches(12), Inches(2.2), SOFT)
q1 = s.shapes.add_textbox(Inches(1.15), Inches(1.95), Inches(11.2), Inches(1.7))
tf = text(q1, "“Reconciliation is critical, but the presentation tools are complicated and nobody understands them”", size=20, color=INK)
para(tf, "— Head of Finance", size=14, bold=True, color=GREEN_DK, before=14)

box(s, Inches(0.75), Inches(4.2), Inches(12), Inches(2.2), WHITE, LINE)
q2 = s.shapes.add_textbox(Inches(1.15), Inches(4.45), Inches(11.2), Inches(1.7))
tf = text(q2, "“If this product existed, you would sell the finance department much easier”", size=20, color=INK)
para(tf, "— Head of Finance  ·  Direct GTM signal for BambooHR", size=14, bold=True, color=GREEN_DK, before=14)
foot(s, 5)
notes(s, "Treat as reasons to believe. Pain = presentation failure. Opportunity = finance buying motion.")


# ---------- 6 Evidence — market ----------
s = slide()
head(s, "Evidence", "Market pattern — this is not one company’s pain")
points = [
    ("72%", "HR & Finance lack shared systems\nfor workforce planning", "Cited via Kinnect (Gartner)"),
    ("45%", "reduction in reconciliation time\nafter HRIS–FP&A integration", "IJIRMPS healthcare case, 2025"),
    ("20–30%", "faster forecasting cycles with\nintegrated planning models", "McKinsey, cited in IJIRMPS 2025"),
]
for i, (big, mid, src) in enumerate(points):
    left = Inches(0.75 + i * 4.1)
    box(s, left, Inches(1.7), Inches(3.85), Inches(4.5), WHITE, LINE)
    b = s.shapes.add_textbox(left + Inches(0.3), Inches(2.15), Inches(3.25), Inches(0.9))
    text(b, big, size=40, bold=True, color=GREEN_DK)
    m = s.shapes.add_textbox(left + Inches(0.3), Inches(3.3), Inches(3.25), Inches(1.6))
    text(m, mid, size=15, color=INK)
    c = s.shapes.add_textbox(left + Inches(0.3), Inches(5.3), Inches(3.25), Inches(0.55))
    text(c, src, size=11, color=MUTED)
foot(s, 6)
notes(s, "Be precise on attribution. Pattern: shared systems scarce; integration cuts recon time and speeds forecasts. Supports product thesis.")


# ---------- 7 Hypotheses ----------
s = slide()
head(s, "Hypotheses", "Five beliefs we designed the prototype to test")
hyps = [
    ("H1", "Cognitive load", "Multi-tab recon exceeds working memory", "Story + one control tower"),
    ("H2", "Definition debt", "Backfill / new / contractor undefined", "Type tags + shared glossary"),
    ("H3", "No decision object", "Approvals live outside the recon", "Ticket queue in-product"),
    ("H4", "Time blindness", "No history or as-of snapshots", "Ranges + audit trail"),
    ("H5", "Wrong default viz", "Grids serve authors, not leaders", "Bridge > Sankey"),
]
for i, (tag, title, belief, test) in enumerate(hyps):
    y = Inches(1.55 + i * 0.95)
    box(s, Inches(0.75), y, Inches(1.0), Inches(0.75), GREEN)
    tg = s.shapes.add_textbox(Inches(0.75), y + Inches(0.2), Inches(1.0), Inches(0.4))
    text(tg, tag, size=14, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    a = s.shapes.add_textbox(Inches(2.0), y + Inches(0.1), Inches(3.5), Inches(0.6))
    text(a, title, size=16, bold=True, color=INK)
    b = s.shapes.add_textbox(Inches(5.6), y + Inches(0.1), Inches(3.8), Inches(0.6))
    text(b, belief, size=14, color=MUTED)
    c = s.shapes.add_textbox(Inches(9.5), y + Inches(0.1), Inches(3.2), Inches(0.6))
    text(c, test, size=14, bold=True, color=GREEN_DK)
foot(s, 7)
notes(s, "AI helped generate; human kept five and mapped each to a UI test.")


# ---------- 8 Interviews ----------
s = slide()
head(s, "Interviews", "What we have — and what we would learn next")
box(s, Inches(0.75), Inches(1.65), Inches(5.9), Inches(4.7), WHITE, LINE)
t1 = s.shapes.add_textbox(Inches(1.1), Inches(1.95), Inches(5.2), Inches(0.4))
text(t1, "In hand", size=14, bold=True, color=GREEN_DK)
b1 = s.shapes.add_textbox(Inches(1.1), Inches(2.5), Inches(5.2), Inches(3.5))
tf = text(b1, "Head of Finance conversation (exact quotes)", size=15, color=INK)
para(tf, "Live spreadsheet: board / roll-forward / HC recon / net-new", size=15, color=INK, before=14)
para(tf, "Observed friction: definitions, as-of, approvals, unexplained variance", size=15, color=INK, before=14)

box(s, Inches(6.9), Inches(1.65), Inches(5.7), Inches(4.7), SOFT)
t2 = s.shapes.add_textbox(Inches(7.25), Inches(1.95), Inches(5.1), Inches(0.4))
text(t2, "Next five interviews", size=14, bold=True, color=GREEN_DK)
b2 = s.shapes.add_textbox(Inches(7.25), Inches(2.5), Inches(5.1), Inches(3.5))
tf = text(b2, "FP&A lead — close ritual & variance questions", size=15, color=INK)
para(tf, "HR ops — ticket classification & SLA", size=15, color=INK, before=12)
para(tf, "Hiring manager — request status needs", size=15, color=INK, before=12)
para(tf, "Controller / CFO — board report bar", size=15, color=INK, before=12)
para(tf, "Eng / ATS owner — source-of-truth constraints", size=15, color=INK, before=12)
foot(s, 8)
notes(s, "Honesty builds credibility: one deep signal + artifact analysis now; structured interviews before scale.")


# ---------- 9 Process ----------
s = slide()
head(s, "Product process", "AI accelerates. Cross-functional partners decide.")
steps = ["Artifact", "AI summarize", "Validate", "Hypothesize", "Prioritize", "Prototype", "Review"]
for i, label in enumerate(steps):
    left = Inches(0.55 + i * 1.8)
    fill = GREEN if i in (2, 4) else SOFT
    box(s, left, Inches(1.65), Inches(1.65), Inches(0.9), fill)
    t = s.shapes.add_textbox(left, Inches(1.85), Inches(1.65), Inches(0.5))
    text(t, label, size=12, bold=True, color=INK if i not in (2, 4) else WHITE, align=PP_ALIGN.CENTER)

people = [
    ("Product", "Frame, prioritize, decision log"),
    ("Finance", "Board, variance, close needs"),
    ("HR", "Ticket semantics & path"),
    ("Design", "30-second comprehension"),
    ("Engineering", "Source of truth & audit"),
    ("Analytics", "Time-to-answer & SLAs"),
]
for i, (h, d) in enumerate(people):
    col = i % 3
    row = i // 3
    left = Inches(0.75 + col * 4.1)
    top = Inches(3.0 + row * 1.75)
    box(s, left, top, Inches(3.85), Inches(1.5), WHITE, LINE)
    a = s.shapes.add_textbox(left + Inches(0.25), top + Inches(0.3), Inches(3.35), Inches(0.4))
    text(a, h, size=16, bold=True, color=INK)
    b = s.shapes.add_textbox(left + Inches(0.25), top + Inches(0.8), Inches(3.35), Inches(0.45))
    text(b, d, size=14, color=MUTED)
foot(s, 9)
notes(s, "Highlight human nodes in green path. Weekly triad + glossary + source-of-truth agreement.")


# ---------- 10 Iterations ----------
s = slide()
head(s, "Iterations", "Ship → show → simplify")
shots = [
    (SHOTS / "01-home-outlook-tile.png", "Outlook", "Story + timeline + history ranges"),
    (SHOTS / "02-bridge-tile.png", "Bridge", "Readable gap explanation"),
    (SHOTS / "03-approvals-tile.png", "Approvals", "Manager → HR → Finance"),
    (SHOTS / "04-export-tile.png", "Export", "Close-ready workbook preview"),
]
for i, (path, label, sub) in enumerate(shots):
    col = i % 2
    row = i // 2
    left = Inches(0.6 + col * 6.35)
    top = Inches(1.5 + row * 2.7)
    box(s, left, top, Inches(6.1), Inches(0.4), SOFT)
    cap = s.shapes.add_textbox(left + Inches(0.2), top + Inches(0.05), Inches(5.7), Inches(0.3))
    text(cap, f"{label}  ·  {sub}", size=12, bold=True, color=INK)
    if path.exists():
        pic = s.shapes.add_picture(str(path), left + Inches(0.05), top + Inches(0.45), width=Inches(6.0))
        max_h = Inches(2.05)
        if pic.height > max_h:
            pic.width = int(pic.width * (max_h / pic.height))
            pic.height = max_h
foot(s, 10)
notes(s, "Key cuts: quotes off product UI; waterfall → bridge; working filters; history ranges; real logo; export preview.")


# ---------- 11 Success ----------
s = slide()
head(s, "Success metrics", "Measure understanding — tie every metric to why")
metrics = [
    ("Time to understand HC @ date", "< 30 seconds", "Comprehension"),
    ("Month-end recon time", "Hours → minutes", "Finance speed"),
    ("Backfill decision SLA", "Median days ↓", "HR clarity"),
    ("Spreadsheet steps removed", "Checklist shrinks", "Ops cost"),
    ("Variance follow-ups / close", "Fewer threads", "Trust"),
    ("Weekly active Finance + HR", "Adoption", "Real use"),
]
for i, (h, target, why) in enumerate(metrics):
    col = i % 3
    row = i // 3
    left = Inches(0.75 + col * 4.1)
    top = Inches(1.65 + row * 2.45)
    box(s, left, top, Inches(3.85), Inches(2.2), WHITE, LINE)
    green = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, Inches(3.85), Inches(0.08))
    green.fill.solid()
    green.fill.fore_color.rgb = GREEN
    green.line.fill.background()
    a = s.shapes.add_textbox(left + Inches(0.25), top + Inches(0.35), Inches(3.35), Inches(0.7))
    text(a, h, size=15, bold=True, color=INK)
    b = s.shapes.add_textbox(left + Inches(0.25), top + Inches(1.1), Inches(3.35), Inches(0.4))
    text(b, target, size=18, bold=True, color=GREEN_DK)
    c = s.shapes.add_textbox(left + Inches(0.25), top + Inches(1.6), Inches(3.35), Inches(0.35))
    text(c, why, size=13, color=MUTED)
foot(s, 11)
notes(s, "North star: time to understand headcount at a given date.")


# ---------- 12 Illustrative impact ----------
s = slide()
head(s, "Illustrative impact", "Modeled for a ~300 FTE company — labeled assumptions")
# honest framing
note = s.shapes.add_textbox(Inches(0.75), Inches(1.5), Inches(12), Inches(0.4))
text(note, "Not a customer ROI claim — a defendable planning model for interview discussion.", size=13, bold=True, color=MUTED)

cards = [
    ("8–12 hrs", "FP&A time / month on recon\nbefore (assumption)", "Based on multi-tab close ritual"),
    ("< 30 min", "Target time to trusted view\nafter Control Tower", "North-star product goal"),
    ("~90%", "Time saved on recon work\nin the model", "Aligns with automation vendor claims; validate in pilot"),
]
for i, (big, mid, src) in enumerate(cards):
    left = Inches(0.75 + i * 4.1)
    box(s, left, Inches(2.1), Inches(3.85), Inches(4.0), WHITE, LINE)
    b = s.shapes.add_textbox(left + Inches(0.3), Inches(2.5), Inches(3.25), Inches(0.8))
    text(b, big, size=36, bold=True, color=GREEN_DK)
    m = s.shapes.add_textbox(left + Inches(0.3), Inches(3.5), Inches(3.25), Inches(1.3))
    text(m, mid, size=15, color=INK)
    c = s.shapes.add_textbox(left + Inches(0.3), Inches(5.2), Inches(3.25), Inches(0.6))
    text(c, src, size=12, color=MUTED)
foot(s, 12)
notes(s, "Be explicit: illustrative. Validate hours in interviews. Pattern matches industry automation claims (TeamOhana/Drivetrain narrative of hours→minutes).")


# ---------- 13 Close ----------
s = slide()
accent = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(0.12), Inches(7.5))
accent.fill.solid()
accent.fill.fore_color.rgb = GREEN
accent.line.fill.background()
if LOGO.exists():
    s.shapes.add_picture(str(LOGO), Inches(0.9), Inches(1.4), width=Inches(2.2))
t = s.shapes.add_textbox(Inches(0.9), Inches(2.4), Inches(11), Inches(1.6))
tf = text(t, "Make headcount reconciliation\nunderstandable in 30 seconds.", size=32, bold=True, color=INK)
b = s.shapes.add_textbox(Inches(0.9), Inches(4.4), Inches(11), Inches(1.4))
tf = text(b, "Problem: comprehension. Why: finance trust + HR decisions + BambooHR GTM.", size=16, color=INK)
para(tf, "Evidence → hypotheses → cross-functional build → iterate → measure.", size=16, color=INK, before=10)
para(tf, "AI accelerated discovery. Judgment drove the product calls.", size=16, bold=True, color=GREEN_DK, before=10)
c = s.shapes.add_textbox(Inches(0.9), Inches(6.3), Inches(11), Inches(0.4))
text(c, "Live prototype link in README  ·  Next: 5 interviews + one close pilot", size=14, color=MUTED)
notes(s, "End on north star. Offer demo + pilot conversation.")

out = ROOT / "Headcount_Control_Tower_Case_Study.pptx"
prs.save(out)
print("Wrote", out, "slides", TOTAL)
