#!/usr/bin/env python3
"""Builds the Cortex Atlas client-ready presentation (.pptx)."""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

# ---- brand palette ----
INK     = RGBColor(0x09,0x13,0x1B)
INK2    = RGBColor(0x0E,0x1A,0x22)
ORANGE  = RGBColor(0xF8,0x97,0x39)
SURFACE = RGBColor(0xF4,0xF4,0xF4)
CARD    = RGBColor(0xFF,0xFF,0xFF)
LINE    = RGBColor(0xD6,0xDC,0xE4)
TEXT2   = RGBColor(0x59,0x59,0x59)
MUTE    = RGBColor(0xA9,0xAF,0xB6)
GREEN   = RGBColor(0x1E,0x7F,0x4F)
RED     = RGBColor(0xE8,0x45,0x1C)
BLUE    = RGBColor(0x05,0x63,0xC1)
GOLD    = RGBColor(0xC9,0xA2,0x27)
WHITE   = RGBColor(0xFF,0xFF,0xFF)
FONT    = "Calibri"   # PPT-friendly: PowerPoint default, renders on any machine

prs = Presentation()
prs.slide_width  = Inches(13.333)
prs.slide_height = Inches(7.5)
SW, SH = prs.slide_width, prs.slide_height
BLANK = prs.slide_layouts[6]
MX = Inches(0.75)            # left margin
CW = Inches(13.333-1.5)      # content width

def slide():
    return prs.slides.add_slide(BLANK)

def bg(s, color):
    s.background.fill.solid()
    s.background.fill.fore_color.rgb = color

def no_shadow(shp):
    # remove default preset shadow
    el = shp._element.spPr
    existing = el.find(qn('a:effectLst'))
    if existing is None:
        el.append(el.makeelement(qn('a:effectLst'), {}))

def rect(s, x,y,w,h, fill=None, line=None, lw=0.75, shape=MSO_SHAPE.RECTANGLE, radius=None):
    shp = s.shapes.add_shape(shape, x,y,w,h)
    if fill is None:
        shp.fill.background()
    else:
        shp.fill.solid(); shp.fill.fore_color.rgb = fill
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line; shp.line.width = Pt(lw)
    no_shadow(shp)
    if radius is not None and shape==MSO_SHAPE.ROUNDED_RECTANGLE:
        try: shp.adjustments[0] = radius
        except Exception: pass
    return shp

def tb(s, x,y,w,h, anchor=MSO_ANCHOR.TOP):
    box = s.shapes.add_textbox(x,y,w,h)
    tf = box.text_frame; tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left=Pt(0); tf.margin_right=Pt(0); tf.margin_top=Pt(0); tf.margin_bottom=Pt(0)
    return tf

def par(tf, text, size=14, bold=False, color=INK, align=PP_ALIGN.LEFT,
        before=0, after=4, name=FONT, first=False, italic=False):
    p = tf.paragraphs[0] if first else tf.add_paragraph()
    p.text = text
    p.alignment = align
    p.space_before = Pt(before); p.space_after = Pt(after)
    for r in p.runs:
        r.font.size=Pt(size); r.font.bold=bold; r.font.italic=italic
        r.font.color.rgb=color; r.font.name=name
    return p

def bullets(tf, items, size=13.5, color=TEXT2, after=7, gap_color=ORANGE):
    """items: list of (lead, rest) -> lead bold colored, rest normal."""
    first=True
    for it in items:
        if isinstance(it, tuple):
            lead, rest = it
        else:
            lead, rest = None, it
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first=False
        p.space_before=Pt(0); p.space_after=Pt(after); p.alignment=PP_ALIGN.LEFT
        # bullet dash
        r0=p.add_run(); r0.text="—  "; r0.font.size=Pt(size); r0.font.bold=True
        r0.font.color.rgb=gap_color; r0.font.name=FONT
        if lead:
            r1=p.add_run(); r1.text=lead; r1.font.size=Pt(size); r1.font.bold=True
            r1.font.color.rgb=INK; r1.font.name=FONT
            r2=p.add_run(); r2.text=" "+rest; r2.font.size=Pt(size); r2.font.color.rgb=color; r2.font.name=FONT
        else:
            r1=p.add_run(); r1.text=rest; r1.font.size=Pt(size); r1.font.color.rgb=color; r1.font.name=FONT

def header(s, kicker, title, sub=None):
    rect(s, MX, Inches(0.62), Inches(0.34), Inches(0.09), fill=ORANGE)
    t=tb(s, MX+Inches(0.45), Inches(0.5), CW-Inches(0.45), Inches(0.35))
    par(t, kicker.upper(), 12.5, True, ORANGE, first=True, after=0)
    t2=tb(s, MX, Inches(0.92), CW, Inches(0.7))
    par(t2, title, 29, True, INK, first=True, after=0)
    if sub:
        t3=tb(s, MX, Inches(1.62), CW, Inches(0.5))
        par(t3, sub, 14.5, False, TEXT2, first=True, after=0)

def footer(s, n):
    rect(s, MX, Inches(7.02), CW, Pt(0.75), fill=LINE)
    t=tb(s, MX, Inches(7.08), Inches(8), Inches(0.3))
    par(t, "Cortex Atlas  ·  Illustrative prototype, synthetic data  ·  Confidential", 9, False, MUTE, first=True, after=0)
    t2=tb(s, SW-MX-Inches(1.0), Inches(7.08), Inches(1.0), Inches(0.3))
    par(t2, str(n), 9, False, MUTE, align=PP_ALIGN.RIGHT, first=True, after=0)

def notes(s, text):
    s.notes_slide.notes_text_frame.text = text

def card(s, x,y,w,h, fill=CARD, line=LINE):
    return rect(s, x,y,w,h, fill=fill, line=line, lw=1.0, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.06)

def badge(s, x,y, text, d=Inches(0.42), fill=ORANGE, fg=INK):
    rect(s, x,y, d,d, fill=fill, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.25)
    t=tb(s, x, y, d, d, anchor=MSO_ANCHOR.MIDDLE)
    par(t, text, 14, True, fg, align=PP_ALIGN.CENTER, first=True, after=0)

def chip(s, x, y, text, color=TEXT2, fill=SURFACE):
    w = Inches(0.16+0.092*len(text))
    rect(s, x, y, w, Inches(0.32), fill=fill, line=LINE, lw=0.75, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.5)
    t=tb(s, x, y, w, Inches(0.32), anchor=MSO_ANCHOR.MIDDLE)
    par(t, text, 10, True, color, align=PP_ALIGN.CENTER, first=True, after=0)
    return x+w+Inches(0.12)

PAGE=0
def pg():
    global PAGE; PAGE+=1; return PAGE

# ============================================================ 1 TITLE
s=slide(); bg(s, INK)
rect(s, 0,0, SW, Inches(0.16), fill=ORANGE)
# brand chip
rect(s, MX, Inches(1.5), Inches(2.5), Inches(0.5), fill=INK2, line=RGBColor(0x2a,0x3a,0x44), lw=1.0, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.5)
t=tb(s, MX, Inches(1.5), Inches(2.5), Inches(0.5), anchor=MSO_ANCHOR.MIDDLE)
par(t, "C O R T E X   S U I T E", 11, True, ORANGE, align=PP_ALIGN.CENTER, first=True, after=0)
t=tb(s, MX, Inches(2.45), Inches(11.5), Inches(1.4))
par(t, "Cortex Atlas", 60, True, WHITE, first=True, after=2)
t=tb(s, MX, Inches(3.7), Inches(11.5), Inches(0.9))
par(t, "Commercial data governance, connected.", 26, False, RGBColor(0xC7,0xCC,0xD3), first=True, after=0)
t=tb(s, MX, Inches(4.7), Inches(11.5), Inches(1.0))
par(t, "Five interlinked tools — requests, catalogue, KPIs, vendors and agreements — rebuilt as one decision layer for Commercial Insights & Data Procurement.", 15, False, MUTE, first=True, after=0)
# footer band
rect(s, MX, Inches(6.4), CW, Pt(0.75), fill=RGBColor(0x2a,0x3a,0x44))
t=tb(s, MX, Inches(6.55), Inches(11.5), Inches(0.5))
par(t, "Prepared for: NovaCura Pharmaceuticals (illustrative)   ·   Working prototype walkthrough   ·   Synthetic data", 11.5, False, MUTE, first=True, after=0)
notes(s, "Open with the one-liner: 'We took the five SharePoint tools your team uses to buy and govern commercial data, and rebuilt them as one connected platform — Cortex Atlas.' Set expectations: this is a working, clickable prototype on synthetic data, not slides of screenshots. Today I'll show what it does, why it matters commercially, and how we'd deliver it. Note it shares the exact design language and architecture as the other Cortex modules — Compass, Vantage, Meridian — so it slots into the suite.")

# ============================================================ 2 AGENDA
s=slide(); bg(s, SURFACE)
header(s, "Agenda", "What we'll cover today")
items=[
 ("01","The challenge","Five tools, five silos — and what that costs"),
 ("02","The solution","Cortex Atlas: one connected estate"),
 ("03","Inside the platform","A guided walk through all five tools"),
 ("04","The decision layer","Automation and intelligence built in"),
 ("05","Value & outcomes","What the business gets"),
 ("06","Delivery & next steps","Architecture, roadmap, and how we start"),
]
y=Inches(1.95); colw=Inches(5.85)
for i,(n,tl,ds) in enumerate(items):
    col=i%2; rowi=i//2
    x=MX+ (col*(colw+Inches(0.3)))
    yy=y+rowi*Inches(1.45)
    card(s, x, yy, colw, Inches(1.2))
    badge(s, x+Inches(0.3), yy+Inches(0.36), n, d=Inches(0.5))
    t=tb(s, x+Inches(1.05), yy+Inches(0.22), colw-Inches(1.3), Inches(0.9), anchor=MSO_ANCHOR.MIDDLE)
    par(t, tl, 17, True, INK, first=True, after=2)
    par(t, ds, 12.5, False, TEXT2, after=0)
footer(s, pg())
notes(s, "Quick roadmap of the conversation. Keep it to 20 seconds. Signal that the bulk of the time (sections 3-4) is the tool itself, and that we'll close on commercials and how we'd deliver. Invite them to stop you with questions at any point — this is meant to be interactive, ideally alongside the live prototype.")

# ============================================================ 3 DIVIDER: CHALLENGE
def divider(num, kicker, title, sub):
    s=slide(); bg(s, INK)
    rect(s, MX, Inches(2.5), Inches(0.34), Inches(0.09), fill=ORANGE)
    t=tb(s, MX+Inches(0.45), Inches(2.38), Inches(10), Inches(0.4))
    par(t, kicker.upper(), 13, True, ORANGE, first=True, after=0)
    t=tb(s, MX, Inches(2.8), Inches(11.8), Inches(1.4))
    par(t, title, 40, True, WHITE, first=True, after=0)
    if sub:
        t=tb(s, MX, Inches(4.2), Inches(11.0), Inches(1.0))
        par(t, sub, 16, False, MUTE, first=True, after=0)
    t=tb(s, SW-MX-Inches(2.2), Inches(0.6), Inches(2.2), Inches(0.5))
    par(t, num, 15, True, RGBColor(0x3a,0x4a,0x54), align=PP_ALIGN.RIGHT, first=True, after=0)
    return s

s=divider("01 / 06","The challenge","Five tools. Five silos.","The pieces are all there today — they just don't talk to each other.")
notes(s, "Transition: 'Before the solution, let's be clear on the problem we keep hearing.' The team already has good tools — this isn't a criticism of what exists. The issue is they're disconnected SharePoint forms and lists, so the organisation can't see across them. That gap is where money leaks and risk hides. Keep this empathetic — we're solving their pain, not selling them something new for its own sake.")

# ============================================================ 4 TODAY: five tools
s=slide(); bg(s, SURFACE)
header(s, "The situation today", "The estate runs on five disconnected tools")
tools=[
 ("DPRH","Data Purchase Request Hub","Where data buys are raised and approved"),
 ("Catalogue","Data Catalogue","The record of everything purchased"),
 ("KPI","KPI Catalogue","The metrics built on that data"),
 ("Vendor","Vendor Hub","Who supplies what"),
 ("TPA","TPA Hub","The agreements that govern sharing"),
]
x=MX; cw=Inches(2.28); y=Inches(2.0)
for i,(tag,tl,ds) in enumerate(tools):
    xx=MX+i*(cw+Inches(0.1))
    card(s, xx, y, cw, Inches(1.9))
    tt=tb(s, xx+Inches(0.18), y+Inches(0.22), cw-Inches(0.36), Inches(1.5))
    par(tt, tag.upper(), 10.5, True, ORANGE, first=True, after=3)
    par(tt, tl, 13.5, True, INK, after=4)
    par(tt, ds, 11, False, TEXT2, after=0)
    if i<4:
        ar=tb(s, xx+cw-Inches(0.02), y+Inches(0.7), Inches(0.14), Inches(0.4), anchor=MSO_ANCHOR.MIDDLE)
        par(ar, "·", 18, True, MUTE, align=PP_ALIGN.CENTER, first=True, after=0)
t=tb(s, MX, Inches(4.3), CW, Inches(2.3))
par(t, "The problem isn't the tools — it's the gaps between them", 17, True, INK, first=True, after=8)
bullets(t, [
 ("No single line of sight.","To answer 'do we already buy this?' someone manually checks the catalogue — if they remember to."),
 ("Manual, reactive renewals.","Expiry tracking lives in spreadsheets and inboxes; agreements lapse or get renewed in a rush, losing leverage."),
 ("Value is invisible.","Spend, quality and dependency sit in separate lists, so no one sees vendor concentration or what a metric actually rests on."),
 ("Governance is effort, not default.","PII/PHI classification and lineage are done by hand, late, and inconsistently."),
], size=13.5)
footer(s, pg())
notes(s, "Walk left to right: this is the lifecycle of a data asset — it's requested (DPRH), recorded (Catalogue), measured (KPIs), supplied by a vendor (Vendor Hub), and governed by an agreement (TPA). Each tool is fine on its own. But because they're separate SharePoint forms, the connections — the dots between them — only exist in people's heads. Then land the four bullets as the lived pain. These are the exact problems Atlas is built to remove, so foreshadow that.")

# ============================================================ 5 COST of status quo
s=slide(); bg(s, SURFACE)
header(s, "Why it matters", "What the disconnect costs — four ways money and risk leak")
cards=[
 ("Duplicate buys","Teams re-purchase data the organisation already owns because there's no check at the point of request.","Avoidable spend"),
 ("Lapsed / rushed renewals","Agreements expire without warning, or renew under time pressure with no negotiating room.","Lost leverage + risk"),
 ("Vendor over-dependence","Spend quietly concentrates in a few vendors with no substitutes — a single point of failure.","Concentration risk"),
 ("Audit exposure","Sensitive data isn't consistently classified and lineage is unclear when the auditor asks.","Compliance risk"),
]
cw=Inches(2.85); y=Inches(2.1)
for i,(tl,ds,tag) in enumerate(cards):
    xx=MX+i*(cw+Inches(0.083))
    card(s, xx, y, cw, Inches(3.2))
    rect(s, xx, y, cw, Inches(0.12), fill=RED if i in(1,3) else ORANGE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.5)
    tt=tb(s, xx+Inches(0.2), y+Inches(0.4), cw-Inches(0.4), Inches(2.7))
    par(tt, tl, 15.5, True, INK, first=True, after=8)
    par(tt, ds, 12, False, TEXT2, after=0)
    tg=tb(s, xx+Inches(0.2), y+Inches(2.75), cw-Inches(0.4), Inches(0.35))
    par(tg, tag.upper(), 10, True, RED if i in(1,3) else RGBColor(0x9a,0x5c,0x12), first=True, after=0)
t=tb(s, MX, Inches(5.6), CW, Inches(1.0))
par(t, "None of this is a people problem. It's a structure problem — and structure is exactly what a connected platform fixes.", 14, True, INK, first=True, after=0)
footer(s, pg())
notes(s, "This is your 'cost of inaction' slide — it earns the right to propose change. You don't need hard client numbers yet (that's the baseline audit later); frame these as the categories of leakage every CI function recognises. Tie each to a P&L or risk owner: duplicate buys = procurement budget; renewals = legal/continuity; concentration = supply risk; audit = compliance. The closing line reframes it as fixable structure, not blame — that lowers defensiveness and sets up the solution.")

# ============================================================ 6 DIVIDER: SOLUTION
s=divider("02 / 06","The solution","One connected estate","A purchase request becomes a catalogued asset, measured by KPIs, supplied by a vendor, governed by an agreement — and Atlas sees all of it.")
notes(s, "Pivot to the upside. The core idea in one sentence: 'one asset, five tools.' Everything from here shows how connecting the five tools turns a set of forms into a decision layer. Say the product name deliberately — Cortex Atlas — and that 'Atlas' is about seeing the whole map, not another dashboard.")

# ============================================================ 7 Vision / one estate
s=slide(); bg(s, SURFACE)
header(s, "The big idea", "Cortex Atlas: the five tools as one estate")
# flow row
flow=[("Data Requests","DPRH"),("Data Catalogue","Asset"),("KPI Catalogue","Metric"),("Vendor Hub","Vendor"),("TPA Hub","Agreement")]
cw=Inches(2.18); y=Inches(2.15)
for i,(tl,tag) in enumerate(flow):
    xx=MX+i*(cw+Inches(0.16))
    c=card(s, xx, y, cw, Inches(1.5))
    tt=tb(s, xx+Inches(0.18), y+Inches(0.24), cw-Inches(0.36), Inches(1.1))
    par(tt, tag.upper(), 10, True, ORANGE, first=True, after=4)
    par(tt, tl, 14.5, True, INK, after=0)
    if i<4:
        ar=tb(s, xx+cw+Inches(0.0), y+Inches(0.5), Inches(0.16), Inches(0.5), anchor=MSO_ANCHOR.MIDDLE)
        par(ar, "→", 16, True, ORANGE, align=PP_ALIGN.CENTER, first=True, after=0)
t=tb(s, MX, Inches(4.0), CW, Inches(0.5))
par(t, "Real links, not labels — every connection is clickable", 16, True, INK, first=True, after=10)
cols=[
 ("From a request","jump to the catalogue entry it creates"),
 ("From an asset","jump to its KPIs, its vendor, its agreement"),
 ("From a KPI","trace the exact asset and vendor it depends on"),
 ("From an agreement","see the asset and the strategic KPI behind it"),
]
cw2=Inches(2.85); y2=Inches(4.6)
for i,(a,b) in enumerate(cols):
    xx=MX+i*(cw2+Inches(0.083))
    card(s, xx, y2, cw2, Inches(1.5))
    tt=tb(s, xx+Inches(0.18), y2+Inches(0.22), cw2-Inches(0.36), Inches(1.1))
    par(tt, a, 13, True, INK, first=True, after=5)
    par(tt, b, 11.5, False, TEXT2, after=0)
footer(s, pg())
notes(s, "Make the flow tangible: trace one asset end to end out loud — 'OncoNova's prescription data is requested in DPRH, recorded in the Catalogue, feeds the New-to-Brand Rx Share KPI, comes from vendor Helix, and is governed by a TPA expiring in three weeks.' That single sentence is impossible to say with today's tools and trivial with Atlas. The bottom row is the payoff: the cross-links are live navigation in the prototype — offer to click any of them during the demo.")

# ============================================================ 8 Three layers
s=slide(); bg(s, SURFACE)
header(s, "How it works", "Three layers, one tool")
layers=[
 ("01","Estate","The five tools, faithful to how they work today","Combined role-gated request form · full catalogue record · KPI registry · vendor inventory · agreement register. The trustworthy foundation.", BLUE),
 ("02","Automation","The work the platform does for you","30/60/90-day renewal reminders · approved request auto-writes to the catalogue · renewals pre-fill from it · PII/PHI auto-classification · duplicate-buy detection at intake.", ORANGE),
 ("03","Intelligence","The decision layer on top","Cross-tool signals · renewal prediction · vendor concentration & value analytics (Vendor Hub Plus) · natural-language Ask Atlas across the estate.", GREEN),
]
cw=Inches(3.83); y=Inches(2.05)
for i,(n,tl,sub,ds,col) in enumerate(layers):
    xx=MX+i*(cw+Inches(0.1))
    card(s, xx, y, cw, Inches(4.3))
    rect(s, xx, y, cw, Inches(0.14), fill=col, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.5)
    badge(s, xx+Inches(0.25), y+Inches(0.4), n, d=Inches(0.5), fill=col, fg=WHITE)
    tt=tb(s, xx+Inches(0.25), y+Inches(1.1), cw-Inches(0.5), Inches(3.0))
    par(tt, tl, 20, True, INK, first=True, after=2)
    par(tt, sub, 12.5, True, col if col!=ORANGE else RGBColor(0x9a,0x5c,0x12), after=10)
    par(tt, ds, 12.5, False, TEXT2, after=0)
footer(s, pg())
notes(s, "This is the mental model to anchor the whole demo. Stress that it's ONE tool delivered together, not three products — the layers describe depth, not phases the client has to buy separately. Layer 1 earns trust (it mirrors their reality). Layer 2 removes manual effort. Layer 3 is where it stops being a system of record and becomes a decision layer. If a client is sceptical of 'AI', reassure: every intelligence output is advisory, with a confidence read, and a human always approves — we'll see that cue on every panel.")

# ============================================================ 9 DIVIDER: INSIDE
s=divider("03 / 06","Inside the platform","A guided walk through the five tools","Each tool keeps its real-world job — then carries the automation and intelligence that belong to it.")
notes(s, "Set up the walkthrough. If you have the live prototype, this is where you switch to it and use the slides as a backstop. Tell them you'll go in the order data actually flows: request, catalogue, KPI, vendor, value layer, agreements — then the cross-estate intelligence that ties it together.")

# ============================================================ 10 Module: DPRH
def module_slide(kicker, title, sub, features, surfaces=None, feeds=None):
    s=slide(); bg(s, SURFACE)
    header(s, kicker, title, sub)
    y=Inches(2.05)
    n=len(features)
    cw=Inches((12.0-(n-1)*0.15)/n) if n<=3 else Inches(3.83)
    for i,(bn,tl,ds) in enumerate(features):
        xx=MX+i*(cw+Inches(0.15))
        card(s, xx, y, cw, Inches(3.05))
        badge(s, xx+Inches(0.22), y+Inches(0.28), bn, d=Inches(0.46))
        tt=tb(s, xx+Inches(0.22), y+Inches(0.95), cw-Inches(0.44), Inches(2.0))
        par(tt, tl, 15, True, INK, first=True, after=6)
        par(tt, ds, 12, False, TEXT2, after=0)
    if surfaces:
        t=tb(s, MX, Inches(5.45), CW, Inches(0.4))
        par(t, "WHAT THE PROTOTYPE SURFACES", 10.5, True, ORANGE, first=True, after=0)
        t2=tb(s, MX, Inches(5.8), CW, Inches(0.8))
        par(t2, surfaces, 13.5, True, INK, first=True, after=0)
    if feeds:
        x=MX
        ty=tb(s, MX, Inches(6.5), Inches(2.2), Inches(0.32)); par(ty,"Production feeds:",9.5,True,TEXT2,first=True,after=0)
        x=MX+Inches(1.35)
        for f in feeds:
            x=chip(s, x, Inches(6.46), f)
    footer(s, pg())
    return s

s=module_slide("Tool 1 — Workflow","Data Requests (DPRH)",
  "One combined, role-gated form governs every data buy before money is committed.",
  [("1","Status-driven & role-gated","Status is set automatically on submit from national/sub-national + cost threshold (e.g. National above → In data sourcing pre review). Reviewer tabs open progressively: Admin → Regional (region-scoped) → CDS committee (regional + global) → Admin final."),
   ("2","Automation built in","Renewals pre-fill from the catalogue; on approval, Move to Catalogue lands the record there as a draft to complete — no re-keying, no drift. Power Automate notifications throughout."),
   ("3","Duplicate-buy detection","At intake, Atlas matches the request against the catalogue on vendor, therapeutic area and data type — and flags overlap before spend."),],
  surfaces="2 in-flight requests flagged as overlapping data already owned — roughly $1.09M of exposure to review before approval.",
  feeds=["SharePoint","Power Automate","Veeva CRM"])
notes(s, "This is the front door — the highest-value place to stop waste, because it's before the money is spent. Land three things: (1) the form is status-driven — on submit it auto-routes by national/sub-national and the divisional cost threshold, and each reviewer tab (Admin, region-scoped Regional, CDS committee of regional+global reviewers, Admin final) opens progressively and is visible only to that role — demo the Requester/Admin/Regional/Global toggle. (2) Automation kills the re-keying that causes catalogue drift; on approval the record moves to the catalogue as a draft to complete. (3) The duplicate flag is the headline — show DP-2042 overlapping the existing Oncology EMR asset. That one flag can pay for the platform. The flag is advisory; a human decides. (This mirrors the live CDSP-DPRH portal.)")

# ============================================================ 11 Module: Catalogue
s=module_slide("Tool 2 — System of record","Data Catalogue",
  "The central record of every asset purchased — the single source everything links to.",
  [("1","One rich record","The full asset profile — identity, classification, coverage, quality, commercials, access — with Gold / Silver / Bronze quality and PII / PHI flags."),
   ("2","Auto-classification","On a new entry, granularity and archetype drive automatic PII/PHI detection and a suggested quality tier — governance by default, not by afterthought."),
   ("3","Three ways in","An approved request, a renewal read-back, or a manual entry — all land in the same trusted record. 'Do we already own this?' finally has one answer."),],
  surfaces="A Bronze-tier asset is feeding a Global, strategic KPI — surfaced as a quality flag before the number is used externally.",
  feeds=["SharePoint","Microsoft Purview","Databricks","Model N"])
notes(s, "Frame the catalogue as the heart — it's the system of record the other four tools point at. Two demo moments: open a full asset record and scroll the grouped fields so they see it's complete, not a stub; then hit 'New manual entry', pick a patient-level claims archetype, and watch PHI flip to Yes with a Gold suggestion in real time — that's the active-metadata idea (Purview thinking) made concrete. Close on the quality flag: this is where the catalogue starts giving advice, not just storing facts.")

# ============================================================ 12 Module: KPI
s=module_slide("Tool 3 — Lineage","KPI Catalogue",
  "The registry of metrics built on the data — each tied back to its source.",
  [("1","Full KPI metadata","Definition, methodology, market definition, cadence, deliverables and contacts — the institutional memory of every metric, in one place."),
   ("2","Lineage to source","Every KPI links to the catalogue asset (and therefore the vendor) that feeds it — so you can trace any number to its origin."),
   ("3","Quality risk, visible","When a strategic KPI rests on a low-tier asset, Atlas flags it — you learn how trustworthy a metric is, not just what it says."),],
  surfaces="Trace 'Diagnosed-but-Untreated Funnel' in two clicks: KPI → Bronze registry asset → vendor Quanta — with the quality caveat attached.",
  feeds=["SharePoint","KPI Catalogue","Microsoft Purview"])
notes(s, "The question this answers: 'when leadership challenges a number in a meeting, can we say where it came from and how good it is?' Today that's tribal knowledge; here it's two clicks. Demo the lineage chain on a KPI drawer. The strategic point: lineage isn't just governance hygiene — it's confidence. Knowing a metric leans on Bronze data changes how you use it. That's the difference between a report and a decision layer.")

# ============================================================ 13 Module: Vendor Hub
s=module_slide("Tool 4 — Inventory","Vendor Hub",
  "The vendor-centric view: who they are, what they hold, and what we actually buy.",
  [("1","Vendor inventory","Each vendor as a first-class record — the data they offer, how to access it, contacts and guidance, in one place."),
   ("2","Purchased vs available","For every vendor, what we buy versus everything they offer — turning coverage into a negotiating fact."),
   ("3","Built to roll up","Assets link up to the vendor, so quality, spend and dependency aggregate cleanly — the foundation for the value layer next."),],
  surfaces="8 vendors supplying 12 assets — and a clear read on how much of each vendor's catalogue we've actually licensed.",
  feeds=["SharePoint","Vendor Hub","Veeva CRM"])
notes(s, "Keep this one short — it's the setup for the star slide that follows. The key design decision to mention: we made each vendor a first-class record that assets link UP to. That sounds technical, but it's what makes the value analytics possible. Today the Vendor Hub is a flat list; tee up that 'flat list becomes a decision tool' in the next slide.")

# ============================================================ 14 Module: Vendor Hub Plus (HERO)
s=slide(); bg(s, INK)
rect(s, MX, Inches(0.62), Inches(0.34), Inches(0.09), fill=ORANGE)
t=tb(s, MX+Inches(0.45), Inches(0.5), CW-Inches(0.45), Inches(0.35)); par(t,"INTELLIGENCE — THE DIFFERENTIATOR",12.5,True,ORANGE,first=True,after=0)
t=tb(s, MX, Inches(0.92), CW, Inches(0.7)); par(t,"Vendor Hub Plus — the value layer",29,True,WHITE,first=True,after=0)
t=tb(s, MX, Inches(1.62), CW, Inches(0.5)); par(t,"The same inventory, rolled up into the five things a vendor decision actually needs.",14.5,False,MUTE,first=True,after=0)
plus=[
 ("Quality score","Asset-level Gold/Silver/Bronze rolled up to a composite vendor grade — who delivers quality, not just who exists."),
 ("Spend concentration","Each vendor's share of spend, with an estate-wide HHI — single-vendor dependency made measurable."),
 ("Irreplaceability","What breaks if a vendor walks: assets and downstream KPIs at stake versus available substitutes."),
 ("Cost-per-quality","Spend normalised against quality — are we overpaying for low-tier data? Renegotiation candidates surface."),
 ("Leverage","Purchased vs available becomes a negotiation lever — coverage gaps as consolidation opportunities."),
]
cw=Inches(2.3); y=Inches(2.4)
for i,(tl,ds) in enumerate(plus):
    xx=MX+i*(cw+Inches(0.1))
    rect(s, xx, y, cw, Inches(3.0), fill=INK2, line=RGBColor(0x2a,0x3a,0x44), lw=1.0, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.06)
    badge(s, xx+Inches(0.2), y+Inches(0.24), str(i+1), d=Inches(0.44))
    tt=tb(s, xx+Inches(0.2), y+Inches(0.9), cw-Inches(0.4), Inches(2.0))
    par(tt, tl, 13.5, True, WHITE, first=True, after=6)
    par(tt, ds, 10.8, False, RGBColor(0xC7,0xCC,0xD3), after=0)
t=tb(s, MX, Inches(5.7), CW, Inches(0.9))
par(t, "Prototype reads: HHI ≈ 1,825 (moderate) · top two vendors hold ~50% of spend · cost-per-quality outlier flagged for renegotiation.", 13, True, ORANGE, first=True, after=4)
par(t, "All advisory — it informs the vendor and renewal decision; a human makes the call.", 11.5, False, MUTE, after=0)
footer(s, pg())
notes(s, "Slow down — this is the slide that separates Atlas from a generic catalogue (Collibra/Atlan) and from their status quo. Borrow the Data Value Tracker framing: instead of only knowing WHO our vendors are, we know what each is WORTH, what depends on them, and where our leverage is. Walk the five layers as a story: quality → concentration → irreplaceability → efficiency → leverage. Use the concrete reads (HHI, top-two 50%, the cost-per-quality outlier) but label them illustrative. This is the slide procurement leadership will remember — it directly arms their next renewal negotiation.")

# ============================================================ 15 Module: TPA Hub
s=module_slide("Tool 5 — Continuous control","TPA Hub",
  "The agreement register — and an active reminder engine that already runs.",
  [("1","30/60/90 reminders","Every agreement is tracked to expiry; reminders fire automatically at 90, 60 and 30 days, escalating to the owner inside 30."),
   ("2","Runway at a glance","A live view of what's expiring when — moving renewals from reactive scramble to planned, leverage-preserving negotiation."),
   ("3","Cross-tool warning","Atlas connects expiry to impact: it flags when an expiring agreement sits behind a strategic KPI, and predicts renew vs retire."),],
  surfaces="6 agreements expire within 90 days (3 within 30) — including the sole agreement behind a Tier-1 OncoNova KPI, due in ~3 weeks.",
  feeds=["SharePoint","Power Automate","Microsoft Purview"])
notes(s, "This is the tool with the clearest, most defensible ROI because the reminder automation already exists in their world — we're making it intelligent. The leap from today: today a reminder says 'TPA expires in 30 days.' Atlas says 'this expires in 21 days AND it's the only agreement feeding your New-to-Brand Rx Share KPI for OncoNova — renew now.' That's continuous, risk-tiered monitoring, the direction third-party-risk tooling has gone. Show the runway view and the red cross-tool warning. Renewal prediction is advisory — a recommendation, not an auto-renew.")

# ============================================================ Enhancements matrix (all tools)
s=slide(); bg(s, SURFACE)
header(s, "Enhancements", "Automation and intelligence — built into every tool")
ORANGE_D=RGBColor(0x9a,0x5c,0x12)
colA=MX+Inches(2.55); colI=MX+Inches(7.3); wA=Inches(4.6); wI=Inches(4.6)
th=tb(s, MX, Inches(1.92), Inches(2.5), Inches(0.3)); par(th,"TOOL",10.5,True,TEXT2,first=True,after=0)
th2=tb(s, colA, Inches(1.92), wA, Inches(0.3)); par(th2,"AUTOMATION (Layer 2)",10.5,True,BLUE,first=True,after=0)
th3=tb(s, colI, Inches(1.92), wI, Inches(0.3)); par(th3,"INTELLIGENCE & NEXT-GEN (Layer 3)",10.5,True,ORANGE_D,first=True,after=0)
rowsM=[
 ("Data Requests","Auto-status routing · renewal pre-fill · move-to-catalogue (draft) · Power Automate alerts","Duplicate-buy detection · approval-likelihood · status dashboard · Ask Atlas"),
 ("Data Catalogue","PII/PHI auto-classification · approved-request sync · renewal read-back","Quality-risk flags · overlap detection · quality dashboard · Ask Atlas"),
 ("KPI Catalogue","Lineage auto-resolve · retirement propagation","Source quality-risk · impact analysis · scope dashboard · Ask Atlas"),
 ("Vendor Hub","Asset-to-vendor rollup · spend aggregation","Vendor Hub Plus: value, HHI, irreplaceability, leverage · Ask Atlas"),
 ("TPA Hub","30/60/90 reminders · owner escalation · email notifications","Renewal prediction · KPI-impact warning · runway dashboard · Ask Atlas"),
]
y=Inches(2.32); rh=Inches(0.82)
for i,(t,a,intel) in enumerate(rowsM):
    yy=y+i*rh
    card(s, MX, yy, Inches(12.0), Inches(0.72))
    rect(s, MX, yy, Inches(0.1), Inches(0.72), fill=ORANGE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.5)
    tt=tb(s, MX+Inches(0.24), yy+Inches(0.08), Inches(2.3), Inches(0.58), anchor=MSO_ANCHOR.MIDDLE); par(tt,t,12.5,True,INK,first=True,after=0)
    ta=tb(s, colA, yy+Inches(0.08), wA, Inches(0.58), anchor=MSO_ANCHOR.MIDDLE); par(ta,a,10.3,False,TEXT2,first=True,after=0)
    ti=tb(s, colI, yy+Inches(0.08), wI, Inches(0.58), anchor=MSO_ANCHOR.MIDDLE); par(ti,intel,10.3,False,TEXT2,first=True,after=0)
t=tb(s, MX, Inches(6.5), CW, Inches(0.5))
par(t,"Cross-connectivity throughout: every record links to its related asset, KPI, vendor and agreement; Ask Atlas spans all five; one connected dashboard layer.",12,True,INK,first=True,after=0)
footer(s, pg())
notes(s, "This is the slide that answers 'what have you enhanced?' across the board. Read it as: every one of the five tools keeps its as-is job (Layer 1) and now carries Layer 2 automation and Layer 3 intelligence — not just one or two tools. Point out the consistent pattern: each has automation that removes manual effort, an AI/analytics capability, a dashboard view, and Ask Atlas. The bottom line is the next-gen integration story — cross-connectivity between tools, a unified dashboard layer, and AI (Azure OpenAI) woven in. Reassure: every AI output stays advisory with a human in the loop.")

# ============================================================ Decision layer: signals + Ask
s=slide(); bg(s, SURFACE)
header(s, "The decision layer", "Cross-tool intelligence and a natural-language ask")
# left: signals
card(s, MX, Inches(2.0), Inches(6.0), Inches(4.4))
tt=tb(s, MX+Inches(0.3), Inches(2.2), Inches(5.4), Inches(0.5)); par(tt,"Cross-tool signals — ranked by what's at stake",14.5,True,INK,first=True,after=0)
sig=[
 (RED,"Agreement at risk feeds a Tier-1 KPI","Helix TPA expires in ~21 days — sole source behind New-to-Brand Rx Share."),
 (ORANGE,"Duplicate-buy risk at intake","DP-2042 overlaps a catalogued Oncology EMR asset — ~$640K avoidable."),
 (ORANGE,"Low-quality asset feeds a strategic KPI","Bronze registry data behind the Diagnosed-but-Untreated Funnel."),
 (BLUE,"Vendor concentration","Top two vendors hold ~50% of spend; estate HHI ≈ 1,825."),
]
yy=Inches(2.75)
for col,tl,ds in sig:
    rect(s, MX+Inches(0.3), yy+Inches(0.05), Inches(0.12), Inches(0.78), fill=col, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.5)
    t=tb(s, MX+Inches(0.55), yy, Inches(5.2), Inches(0.9))
    par(t, tl, 12.5, True, INK, first=True, after=2)
    par(t, ds, 11, False, TEXT2, after=0)
    yy=yy+Inches(0.9)
# right: Ask
rect(s, MX+Inches(6.3), Inches(2.0), Inches(5.5), Inches(4.4), fill=INK2, line=None, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.05)
tt=tb(s, MX+Inches(6.6), Inches(2.25), Inches(4.9), Inches(0.5)); par(tt,"ASK ATLAS",12,True,ORANGE,first=True,after=0)
rect(s, MX+Inches(6.6), Inches(2.7), Inches(4.9), Inches(0.55), fill=RGBColor(0x16,0x24,0x2e), line=RGBColor(0x2a,0x3a,0x44), lw=1.0, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.2)
t=tb(s, MX+Inches(6.8), Inches(2.7), Inches(4.5), Inches(0.55), anchor=MSO_ANCHOR.MIDDLE); par(t,"Which agreements expire soon and feed a Tier-1 KPI?",11.5,False,RGBColor(0xC7,0xCC,0xD3),first=True,after=0)
t=tb(s, MX+Inches(6.6), Inches(3.5), Inches(4.9), Inches(2.4))
par(t,"Atlas answers across all five tools — grounded in the catalogue, KPI lineage, vendor spend and agreement dates — and links you straight to the record.",12.5,False,RGBColor(0xE8,0xEA,0xEE),first=True,after=12)
par(t,"Every answer carries a confidence read and an advisory cue:",11.5,False,MUTE,after=8)
par(t,"\"Confidence: High · grounded in source rows · Advisory only — a human approves any action.\"",11.5,True,RGBColor(0x7F,0xD1,0xA8),after=0)
footer(s, pg())
notes(s, "This is where the pieces become a decision layer. Left: Atlas doesn't make you hunt — it pushes the handful of cross-tool issues that need a decision, ranked by stakes, each linking to the record. Right: for everything else, ask in plain English and it reads across all five tools. Demo one chip live. Crucial trust message for a pharma audience: it's grounded (no hallucinated numbers — it cites the source rows), it shows its confidence, and it's advisory only — a human always approves. No autonomous decisions. That sentence often unlocks the room.")

# ============================================================ 17 DIVIDER: VALUE
s=divider("04 / 06","Value & delivery","From report to decision layer","What the business gets — and how we build it.")
notes(s, "Transition from 'what it does' to 'why it pays and how we deliver.' Keep momentum — the demo has done the convincing; now make it easy to say yes.")

# ============================================================ 18 Business value
s=slide(); bg(s, SURFACE)
header(s, "Business value", "Five outcomes leadership can point to")
val=[
 ("Stop duplicate spend","Overlap is caught at the point of request, before commitment — direct, recurring savings.","Procurement"),
 ("Never miss a renewal","Automated 30/60/90 reminders plus impact-aware warnings end lapses and rushed, low-leverage renewals.","Legal / Continuity"),
 ("Negotiate from strength","Vendor value, concentration and coverage analytics turn renewals into informed negotiations.","Procurement / CI"),
 ("Audit-ready by default","Automatic PII/PHI classification and full lineage make governance the default, not a project.","Compliance"),
 ("Faster, confident decisions","One source of truth and a natural-language ask cut manual lookups and let teams trust the number.","Commercial Insights"),
]
cw=Inches(5.85);
for i,(tl,ds,owner) in enumerate(val):
    col=i%2; rowi=i//2
    xx=MX+col*(cw+Inches(0.3)); yy=Inches(2.0)+rowi*Inches(1.18)
    if i==4:
        xx=MX; yy=Inches(2.0)+2*Inches(1.18);
    card(s, xx, yy, cw if i!=4 else Inches(12.0), Inches(1.05))
    badge(s, xx+Inches(0.22), yy+Inches(0.3), str(i+1), d=Inches(0.46), fill=GREEN, fg=WHITE)
    tt=tb(s, xx+Inches(0.95), yy+Inches(0.16), (cw if i!=4 else Inches(12.0))-Inches(3.0), Inches(0.8), anchor=MSO_ANCHOR.MIDDLE)
    par(tt, tl, 14.5, True, INK, first=True, after=2)
    par(tt, ds, 11.3, False, TEXT2, after=0)
    ow=tb(s, xx+(cw if i!=4 else Inches(12.0))-Inches(1.9), yy+Inches(0.3), Inches(1.7), Inches(0.45), anchor=MSO_ANCHOR.MIDDLE)
    par(ow, owner.upper(), 9.5, True, GREEN, align=PP_ALIGN.RIGHT, first=True, after=0)
footer(s, pg())
notes(s, "Translate features into outcomes and name the owner who cares about each — that's how you build a coalition of sponsors in the room. Lead with 'stop duplicate spend' and 'never miss a renewal' because they're the easiest to quantify in a baseline audit. If asked for ROI numbers, be honest: the prototype shows the mechanism; the baseline audit (next steps) quantifies it against their actual spend and agreements. Avoid over-claiming hard figures on synthetic data.")

# ============================================================ 19 Architecture
s=slide(); bg(s, SURFACE)
header(s, "Design & architecture", "Built to drop into your Microsoft estate — and the Cortex suite")
# left card: prototype
card(s, MX, Inches(2.0), Inches(5.85), Inches(4.4))
tt=tb(s, MX+Inches(0.3), Inches(2.2), Inches(5.25), Inches(4.0))
par(tt,"The prototype you're seeing",15,True,INK,first=True,after=8)
bullets(tt, [
 ("One self-contained file.","Runs in any browser, fully offline — no install to evaluate it."),
 ("Cortex design system.","Same look and behaviour as Compass, Vantage and Meridian."),
 ("Lift-and-shift ready.","Drops into the cortex-commercial-suite with near-zero rework."),
 ("Synthetic data only.","No real client or patient data anywhere in the prototype."),
], size=12.5)
# right card: target stack
card(s, MX+Inches(6.15), Inches(2.0), Inches(5.85), Inches(4.4))
tt=tb(s, MX+Inches(6.45), Inches(2.2), Inches(5.25), Inches(0.5)); par(tt,"Target production stack — Microsoft-native",15,True,INK,first=True,after=8)
stack=[
 ("SharePoint","Storage & lists — the records system"),
 ("Power Automate","Workflow & the 30/60/90 reminders"),
 ("Microsoft Purview","Catalogue, classification & lineage"),
 ("Azure OpenAI","The intelligence & Ask Atlas layer"),
 ("Databricks / Model N / SAP","Data, spend and contract feeds"),
]
yy=Inches(2.75)
for nm,ds in stack:
    rect(s, MX+Inches(6.45), yy+Inches(0.04), Inches(0.1), Inches(0.5), fill=ORANGE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.5)
    t=tb(s, MX+Inches(6.7), yy, Inches(5.0), Inches(0.6))
    par(t, nm, 12.5, True, INK, first=True, after=1)
    par(t, ds, 10.8, False, TEXT2, after=0)
    yy=yy+Inches(0.68)
footer(s, pg())
notes(s, "Reassure IT and architecture stakeholders. Two messages. Left: the prototype is deliberately a single offline file so anyone can evaluate it without procurement or installs — and it reuses the proven Cortex pattern, so it's not bespoke risk. Right: the production target maps cleanly onto tools they already own and trust — SharePoint, Power Automate, Purview, Azure OpenAI. We're not asking them to adopt a new platform; we're orchestrating their existing Microsoft estate. That lowers security, procurement and change-management friction enormously.")

# ============================================================ 20 Roadmap
s=slide(); bg(s, SURFACE)
header(s, "Delivery roadmap", "A phased rollout — value at every step")
phases=[
 ("Phase 1","Foundation","Migrate the five tools, the data model and the cross-links onto SharePoint. The as-is estate, connected and trustworthy.",BLUE),
 ("Phase 2","Automation","Power Automate reminders, approved-request → catalogue sync, renewal pre-fill, Purview PII/PHI classification, duplicate detection.",ORANGE),
 ("Phase 3","Intelligence","Vendor Hub Plus, renewal prediction, cross-tool signals and Ask Atlas on Azure OpenAI — the full decision layer.",GREEN),
]
cw=Inches(3.83); y=Inches(2.2)
for i,(ph,tl,ds,col) in enumerate(phases):
    xx=MX+i*(cw+Inches(0.1))
    card(s, xx, y, cw, Inches(3.2))
    rect(s, xx, y, cw, Inches(0.14), fill=col, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.5)
    tt=tb(s, xx+Inches(0.25), y+Inches(0.4), cw-Inches(0.5), Inches(2.7))
    par(tt, ph.upper(), 11, True, col if col!=ORANGE else RGBColor(0x9a,0x5c,0x12), first=True, after=2)
    par(tt, tl, 18, True, INK, after=8)
    par(tt, ds, 12, False, TEXT2, after=0)
    if i<2:
        ar=tb(s, xx+cw-Inches(0.02), y+Inches(1.4), Inches(0.16), Inches(0.5), anchor=MSO_ANCHOR.MIDDLE)
        par(ar,"→",16,True,MUTE,align=PP_ALIGN.CENTER,first=True,after=0)
t=tb(s, MX, Inches(5.7), CW, Inches(0.8))
par(t,"The prototype already demonstrates all three layers — so each phase de-risks the next, and you see working value from Phase 1 onward.",14,True,INK,first=True,after=0)
footer(s, pg())
notes(s, "Phasing makes a big programme feel safe and fundable. Key reframe: although we're showing all three layers working today in the prototype, we deliver them in sequence so the client gets usable value early and each phase proves the ground for the next. Phase 1 alone — connected tools with real cross-links — already beats today. Don't commit to calendar durations here unless you've scoped them; talk in phases and let the next-steps/scoping conversation set dates.")

# ============================================================ 21 Why this approach
s=slide(); bg(s, SURFACE)
header(s, "Why this approach", "Lower risk, faster value, proven pattern")
why=[
 ("Proven in the suite","Atlas reuses the exact design and structure as Compass, Vantage and Meridian — a pattern already delivered, not invented here."),
 ("Your stack, orchestrated","Built on Microsoft tools you already own — minimal new procurement, security or training overhead."),
 ("Evaluate it today","A working, offline prototype means stakeholders can click through and react now — no leap of faith."),
 ("Built to grow","Clean module separation means automation and intelligence bolt on without a rewrite."),
]
cw=Inches(5.85)
for i,(tl,ds) in enumerate(why):
    col=i%2; rowi=i//2
    xx=MX+col*(cw+Inches(0.3)); yy=Inches(2.1)+rowi*Inches(2.05)
    card(s, xx, yy, cw, Inches(1.85))
    badge(s, xx+Inches(0.28), yy+Inches(0.3), "✓", d=Inches(0.5), fill=GREEN, fg=WHITE)
    tt=tb(s, xx+Inches(1.0), yy+Inches(0.28), cw-Inches(1.3), Inches(1.3))
    par(tt, tl, 16, True, INK, first=True, after=6)
    par(tt, ds, 12.5, False, TEXT2, after=0)
footer(s, pg())
notes(s, "Pre-empt the 'why you / why this way' questions. The strongest card is the first — this isn't a science project; it's the same pattern we've already shipped as other Cortex modules, so the delivery risk is low. Pair that with 'it runs on what you already own' and 'you can try it today' and you've removed the three usual objections: risk, cost of change, and uncertainty. Keep this brisk; it's reinforcement before the ask.")

# ============================================================ 22 Next steps
s=slide(); bg(s, INK)
rect(s, 0,0, SW, Inches(0.16), fill=ORANGE)
t=tb(s, MX+Inches(0.45), Inches(0.7), Inches(10), Inches(0.4)); par(t,"NEXT STEPS",13,True,ORANGE,first=True,after=0)
t=tb(s, MX, Inches(1.1), Inches(11.8), Inches(0.9)); par(t,"How we start",36,True,WHITE,first=True,after=0)
steps=[
 ("1","Walk the prototype live","A hands-on session with your CI and procurement leads — react to the real thing."),
 ("2","Baseline data audit","Point Atlas at your real estate to quantify duplication, renewal risk and concentration — the disclaimer becomes real numbers."),
 ("3","Confirm scope & architecture","Agree Phase 1 boundaries and the SharePoint / Power Automate / Purview target with your IT team."),
 ("4","Stand up Phase 1","Migrate the connected five-tool foundation — value from day one."),
]
y=Inches(2.3)
for n,tl,ds in steps:
    badge(s, MX, y, n, d=Inches(0.5), fill=ORANGE, fg=INK)
    t=tb(s, MX+Inches(0.8), y-Inches(0.02), Inches(11.0), Inches(0.95))
    par(t, tl, 17, True, WHITE, first=True, after=2)
    par(t, ds, 12.5, False, RGBColor(0xC7,0xCC,0xD3), after=0)
    y=y+Inches(1.02)
t=tb(s, MX, Inches(6.7), Inches(11.8), Inches(0.5))
par(t,"Thank you.  Let's open the prototype.",15,True,ORANGE,first=True,after=0)
notes(s, "Close with a clear, low-commitment ask: the next step is just a live walkthrough with their people — easy to say yes to. The baseline audit is the real hook: it converts every illustrative number in this deck into their actual exposure, which is what unlocks budget. End by inviting them into the prototype — finish on the product, not a slide. Hand over to the live demo or open for questions.")

# ============================================================ 23 Appendix / disclaimer
s=slide(); bg(s, SURFACE)
header(s, "Appendix", "Notes, scope and disclaimer")
t=tb(s, MX, Inches(2.0), CW, Inches(4.5))
par(t,"On the data",14,True,INK,first=True,after=6)
par(t,"All figures, companies, brands and vendors in this prototype and deck are synthetic and illustrative (NovaCura Pharmaceuticals; brands OncoNova, Immunexa, CardioZen, Neurovia, RareGene; invented vendors such as Helix Health Data and Cardinal Claims Co). Public vendor names (Veeva, IQVIA, Databricks, Model N, SAP, SharePoint, Microsoft Purview) appear only as integration labels.",12.5,False,TEXT2,after=12)
par(t,"On the figures",14,True,INK,after=6)
par(t,"Numbers shown — spend, HHI, expiry counts, duplicate exposure — are directional and exist to demonstrate the mechanics. A client baseline audit is required before any figure is used externally or for decisions.",12.5,False,TEXT2,after=12)
par(t,"On the intelligence",14,True,INK,after=6)
par(t,"Every AI and analytics output in Atlas is advisory only, carries a confidence read, and requires human approval. There are no autonomous decisions or actions.",12.5,False,TEXT2,after=0)
footer(s, pg())
notes(s, "Leave-behind / governance slide — you may not present it, but it protects you and answers the compliance reviewer who reads the deck afterwards. The three points: data is synthetic; figures are directional pending a baseline audit; AI is advisory with human approval. If anyone asks 'are these our numbers?' this is your reference.")

prs.save("/home/user/datasharing/Cortex-Atlas-Client-Deck.pptx")
print("Saved with", len(prs.slides._sldIdLst), "slides")
