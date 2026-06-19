#!/usr/bin/env python3
"""Builds the Cortex Atlas EXECUTIVE CUT (~10 slides)."""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

INK=RGBColor(0x09,0x13,0x1B); INK2=RGBColor(0x0E,0x1A,0x22); ORANGE=RGBColor(0xF8,0x97,0x39)
SURFACE=RGBColor(0xF4,0xF4,0xF4); CARD=RGBColor(0xFF,0xFF,0xFF); LINE=RGBColor(0xD6,0xDC,0xE4)
TEXT2=RGBColor(0x59,0x59,0x59); MUTE=RGBColor(0xA9,0xAF,0xB6); GREEN=RGBColor(0x1E,0x7F,0x4F)
RED=RGBColor(0xE8,0x45,0x1C); BLUE=RGBColor(0x05,0x63,0xC1); WHITE=RGBColor(0xFF,0xFF,0xFF)
ORANGE_D=RGBColor(0x9a,0x5c,0x12); STROKE=RGBColor(0x2a,0x3a,0x44)
FONT="Calibri"

prs=Presentation(); prs.slide_width=Inches(13.333); prs.slide_height=Inches(7.5)
SW,SH=prs.slide_width,prs.slide_height; BLANK=prs.slide_layouts[6]
MX=Inches(0.75); CW=Inches(13.333-1.5)

def slide(): return prs.slides.add_slide(BLANK)
def bg(s,c): s.background.fill.solid(); s.background.fill.fore_color.rgb=c
def no_shadow(shp):
    el=shp._element.spPr
    if el.find(qn('a:effectLst')) is None: el.append(el.makeelement(qn('a:effectLst'),{}))
def rect(s,x,y,w,h,fill=None,line=None,lw=0.75,shape=MSO_SHAPE.RECTANGLE,radius=None):
    shp=s.shapes.add_shape(shape,x,y,w,h)
    if fill is None: shp.fill.background()
    else: shp.fill.solid(); shp.fill.fore_color.rgb=fill
    if line is None: shp.line.fill.background()
    else: shp.line.color.rgb=line; shp.line.width=Pt(lw)
    no_shadow(shp)
    if radius is not None and shape==MSO_SHAPE.ROUNDED_RECTANGLE:
        try: shp.adjustments[0]=radius
        except Exception: pass
    return shp
def tb(s,x,y,w,h,anchor=MSO_ANCHOR.TOP):
    b=s.shapes.add_textbox(x,y,w,h); tf=b.text_frame; tf.word_wrap=True; tf.vertical_anchor=anchor
    tf.margin_left=Pt(0); tf.margin_right=Pt(0); tf.margin_top=Pt(0); tf.margin_bottom=Pt(0); return tf
def par(tf,text,size=14,bold=False,color=INK,align=PP_ALIGN.LEFT,before=0,after=4,name=FONT,first=False):
    p=tf.paragraphs[0] if first else tf.add_paragraph()
    p.text=text; p.alignment=align; p.space_before=Pt(before); p.space_after=Pt(after)
    for r in p.runs: r.font.size=Pt(size); r.font.bold=bold; r.font.color.rgb=color; r.font.name=name
    return p
def bullets(tf,items,size=13.5,color=TEXT2,after=7):
    first=True
    for it in items:
        lead,rest=it if isinstance(it,tuple) else (None,it)
        p=tf.paragraphs[0] if first else tf.add_paragraph(); first=False
        p.space_before=Pt(0); p.space_after=Pt(after); p.alignment=PP_ALIGN.LEFT
        r0=p.add_run(); r0.text="—  "; r0.font.size=Pt(size); r0.font.bold=True; r0.font.color.rgb=ORANGE; r0.font.name=FONT
        if lead:
            r1=p.add_run(); r1.text=lead; r1.font.size=Pt(size); r1.font.bold=True; r1.font.color.rgb=INK; r1.font.name=FONT
            r2=p.add_run(); r2.text=" "+rest; r2.font.size=Pt(size); r2.font.color.rgb=color; r2.font.name=FONT
        else:
            r1=p.add_run(); r1.text=rest; r1.font.size=Pt(size); r1.font.color.rgb=color; r1.font.name=FONT
def badge(s,x,y,text,d=Inches(0.42),fill=ORANGE,fg=INK):
    rect(s,x,y,d,d,fill=fill,shape=MSO_SHAPE.ROUNDED_RECTANGLE,radius=0.25)
    t=tb(s,x,y,d,d,anchor=MSO_ANCHOR.MIDDLE); par(t,text,14,True,fg,align=PP_ALIGN.CENTER,first=True,after=0)
def card(s,x,y,w,h,fill=CARD,line=LINE): return rect(s,x,y,w,h,fill=fill,line=line,lw=1.0,shape=MSO_SHAPE.ROUNDED_RECTANGLE,radius=0.06)
def notes(s,text): s.notes_slide.notes_text_frame.text=text
def header(s,kicker,title,sub=None):
    rect(s,MX,Inches(0.62),Inches(0.34),Inches(0.09),fill=ORANGE)
    t=tb(s,MX+Inches(0.45),Inches(0.5),CW-Inches(0.45),Inches(0.35)); par(t,kicker.upper(),12.5,True,ORANGE,first=True,after=0)
    t2=tb(s,MX,Inches(0.92),CW,Inches(0.7)); par(t2,title,29,True,INK,first=True,after=0)
    if sub:
        t3=tb(s,MX,Inches(1.62),CW,Inches(0.5)); par(t3,sub,14.5,False,TEXT2,first=True,after=0)
def footer(s,n):
    rect(s,MX,Inches(7.02),CW,Pt(0.75),fill=LINE)
    t=tb(s,MX,Inches(7.08),Inches(9),Inches(0.3)); par(t,"Cortex Atlas  ·  Executive summary  ·  Illustrative prototype, synthetic data  ·  Confidential",9,False,MUTE,first=True,after=0)
    t2=tb(s,SW-MX-Inches(1.0),Inches(7.08),Inches(1.0),Inches(0.3)); par(t2,str(n),9,False,MUTE,align=PP_ALIGN.RIGHT,first=True,after=0)
PAGE=0
def pg():
    global PAGE; PAGE+=1; return PAGE

# ---------- 1 Title
s=slide(); bg(s,INK)
rect(s,0,0,SW,Inches(0.16),fill=ORANGE)
rect(s,MX,Inches(1.45),Inches(2.7),Inches(0.5),fill=INK2,line=STROKE,lw=1.0,shape=MSO_SHAPE.ROUNDED_RECTANGLE,radius=0.5)
t=tb(s,MX,Inches(1.45),Inches(2.7),Inches(0.5),anchor=MSO_ANCHOR.MIDDLE); par(t,"EXECUTIVE SUMMARY",11,True,ORANGE,align=PP_ALIGN.CENTER,first=True,after=0)
t=tb(s,MX,Inches(2.35),Inches(11.5),Inches(1.3)); par(t,"Cortex Atlas",58,True,WHITE,first=True,after=2)
t=tb(s,MX,Inches(3.55),Inches(11.5),Inches(0.8)); par(t,"Commercial data governance, connected.",25,False,RGBColor(0xC7,0xCC,0xD3),first=True,after=0)
t=tb(s,MX,Inches(4.5),Inches(11.5),Inches(1.0)); par(t,"Five disconnected tools rebuilt as one decision layer — stopping duplicate spend, ending missed renewals, and turning vendor data into negotiating leverage.",15,False,MUTE,first=True,after=0)
rect(s,MX,Inches(6.3),CW,Pt(0.75),fill=STROKE)
t=tb(s,MX,Inches(6.45),Inches(11.5),Inches(0.5)); par(t,"For Commercial Insights & Data Procurement leadership  ·  NovaCura Pharmaceuticals (illustrative)",11.5,False,MUTE,first=True,after=0)
notes(s,"Executive cut — designed for a 10-15 minute leadership conversation or as a pre-read. Same story as the full deck, compressed to the decisions a sponsor cares about. Open on the outcome, not the tech: 'we connected five tools so the business stops paying for data twice, never gets surprised by a renewal, and walks into vendor negotiations with the facts.'")

# ---------- 2 Challenge (combined)
s=slide(); bg(s,SURFACE)
header(s,"The challenge","Five disconnected tools — and what the gaps cost")
t=tb(s,MX,Inches(1.75),CW,Inches(0.4)); par(t,"Today the request hub, catalogue, KPI registry, vendor hub and agreement register are separate SharePoint forms — no line of sight across them.",13,False,TEXT2,first=True,after=0)
cards=[("Duplicate buys","Teams re-purchase data already owned — no check at the point of request.","Avoidable spend"),
 ("Lapsed / rushed renewals","Agreements expire unseen or renew under pressure, losing leverage.","Lost leverage"),
 ("Vendor over-dependence","Spend quietly concentrates with no substitutes — single points of failure.","Concentration risk"),
 ("Audit exposure","Sensitive data isn't consistently classified; lineage is unclear.","Compliance risk")]
cw=Inches(2.85); y=Inches(2.5)
for i,(tl,ds,tag) in enumerate(cards):
    xx=MX+i*(cw+Inches(0.083)); card(s,xx,y,cw,Inches(3.4))
    rect(s,xx,y,cw,Inches(0.12),fill=RED if i in(1,3) else ORANGE,shape=MSO_SHAPE.ROUNDED_RECTANGLE,radius=0.5)
    tt=tb(s,xx+Inches(0.2),y+Inches(0.4),cw-Inches(0.4),Inches(2.7)); par(tt,tl,15.5,True,INK,first=True,after=8); par(tt,ds,12,False,TEXT2,after=0)
    tg=tb(s,xx+Inches(0.2),y+Inches(2.95),cw-Inches(0.4),Inches(0.35)); par(tg,tag.upper(),10,True,RED if i in(1,3) else ORANGE_D,first=True,after=0)
t=tb(s,MX,Inches(6.15),CW,Inches(0.6)); par(t,"It's not a people problem — it's a structure problem. And structure is what a connected platform fixes.",14,True,INK,first=True,after=0)
footer(s,pg())
notes(s,"One slide for the whole problem. Name the risk owner for each card: procurement (duplicate spend), legal/continuity (renewals), supply risk (concentration), compliance (audit). You're assembling a coalition of sponsors. The closing line reframes blame as fixable structure.")

# ---------- 3 Solution (estate + 3 layers)
s=slide(); bg(s,SURFACE)
header(s,"The solution","One connected estate, in three layers")
flow=[("Requests","DPRH"),("Catalogue","Asset"),("KPIs","Metric"),("Vendors","Vendor"),("Agreements","TPA")]
cw=Inches(2.18); y=Inches(1.85)
for i,(tl,tag) in enumerate(flow):
    xx=MX+i*(cw+Inches(0.16)); card(s,xx,y,cw,Inches(1.15))
    tt=tb(s,xx+Inches(0.16),y+Inches(0.2),cw-Inches(0.32),Inches(0.85)); par(tt,tag.upper(),9.5,True,ORANGE,first=True,after=3); par(tt,tl,13.5,True,INK,after=0)
    if i<4:
        ar=tb(s,xx+cw,y+Inches(0.35),Inches(0.16),Inches(0.5),anchor=MSO_ANCHOR.MIDDLE); par(ar,"→",15,True,ORANGE,align=PP_ALIGN.CENTER,first=True,after=0)
t=tb(s,MX,Inches(3.05),CW,Inches(0.35)); par(t,'"One asset, five tools" — and every link is clickable, not just labelled.',12.5,False,TEXT2,first=True,after=0)
layers=[("01","Estate","The five tools as they work today — the trustworthy foundation.",BLUE),
 ("02","Automation","30/60/90 renewal reminders · request↔catalogue sync · PII/PHI auto-classification · duplicate detection.",ORANGE),
 ("03","Intelligence","Cross-tool signals · renewal prediction · vendor value analytics · natural-language Ask Atlas.",GREEN)]
cw=Inches(3.83); y=Inches(3.55)
for i,(n,tl,ds,col) in enumerate(layers):
    xx=MX+i*(cw+Inches(0.1)); card(s,xx,y,cw,Inches(2.9))
    rect(s,xx,y,cw,Inches(0.14),fill=col,shape=MSO_SHAPE.ROUNDED_RECTANGLE,radius=0.5)
    badge(s,xx+Inches(0.25),y+Inches(0.35),n,d=Inches(0.5),fill=col,fg=WHITE)
    tt=tb(s,xx+Inches(0.25),y+Inches(1.05),cw-Inches(0.5),Inches(1.7)); par(tt,tl,19,True,INK,first=True,after=6); par(tt,ds,12,False,TEXT2,after=0)
footer(s,pg())
notes(s,"The mental model. Trace one asset end to end out loud to make the flow real. Then the three layers: foundation earns trust, automation removes effort, intelligence makes it a decision layer. Stress it's ONE tool delivered together — the layers are depth, not separate purchases.")

# ---------- 4 Platform at a glance
s=slide(); bg(s,SURFACE)
header(s,"Inside the platform","Every tool keeps its as-is job and now carries automation, intelligence, dashboards and Ask Atlas")
rows=[("1","Data Requests (DPRH)","Status-driven, role-gated form; auto-routed on submit.","Auto-status · duplicate detection (AI) · dashboard · Ask Atlas"),
 ("2","Data Catalogue","System of record — asset profile, quality, sensitivity.","PII/PHI auto-classify · quality flags (AI) · Ask Atlas"),
 ("3","KPI Catalogue","Every metric, with lineage to its source asset.","Lineage auto-resolve · source quality-risk (AI) · Ask Atlas"),
 ("4","Vendor Hub","Vendor inventory — purchased vs available.","Rollups · Vendor Hub Plus value & HHI (AI) · Ask Atlas"),
 ("5","TPA Hub","Agreement register with a live expiry engine.","30/60/90 reminders · renewal prediction (AI) · dashboard")]
y=Inches(1.95); rh=Inches(0.94)
for i,(n,tl,ds,tag) in enumerate(rows):
    yy=y+i*rh; card(s,MX,yy,CW,Inches(0.82))
    badge(s,MX+Inches(0.2),yy+Inches(0.21),n,d=Inches(0.42))
    t=tb(s,MX+Inches(0.85),yy+Inches(0.11),Inches(3.4),Inches(0.62),anchor=MSO_ANCHOR.MIDDLE); par(t,tl,13.5,True,INK,first=True,after=0)
    t2=tb(s,MX+Inches(4.3),yy+Inches(0.11),Inches(3.5),Inches(0.62),anchor=MSO_ANCHOR.MIDDLE); par(t2,ds,10.8,False,TEXT2,first=True,after=0)
    t3=tb(s,MX+Inches(7.95),yy+Inches(0.11),Inches(3.85),Inches(0.62),anchor=MSO_ANCHOR.MIDDLE); par(t3,tag,9.5,True,ORANGE_D,first=True,after=0)
footer(s,pg())
notes(s,"The whole platform on one slide. Read down the left for the lifecycle of an asset; the right column shows the automation/intelligence each tool carries. If you're demoing live, this is your map before you click in. Keep it to a minute and move to the differentiator.")

# ---------- 5 Vendor Hub Plus (hero)
s=slide(); bg(s,INK)
rect(s,MX,Inches(0.62),Inches(0.34),Inches(0.09),fill=ORANGE)
t=tb(s,MX+Inches(0.45),Inches(0.5),CW-Inches(0.45),Inches(0.35)); par(t,"THE DIFFERENTIATOR",12.5,True,ORANGE,first=True,after=0)
t=tb(s,MX,Inches(0.92),CW,Inches(0.7)); par(t,"Vendor Hub Plus — the value layer",29,True,WHITE,first=True,after=0)
t=tb(s,MX,Inches(1.62),CW,Inches(0.5)); par(t,"The same inventory, rolled up into the five things a vendor decision actually needs.",14.5,False,MUTE,first=True,after=0)
plus=[("Quality score","Asset Gold/Silver/Bronze rolled up to a composite vendor grade."),
 ("Concentration","Each vendor's share of spend, with an estate-wide HHI."),
 ("Irreplaceability","What breaks if a vendor walks — assets and KPIs at stake vs substitutes."),
 ("Cost-per-quality","Spend normalised against quality — overpaying for low-tier data?"),
 ("Leverage","Purchased vs available as a negotiation lever.")]
cw=Inches(2.3); y=Inches(2.4)
for i,(tl,ds) in enumerate(plus):
    xx=MX+i*(cw+Inches(0.1)); rect(s,xx,y,cw,Inches(3.0),fill=INK2,line=STROKE,lw=1.0,shape=MSO_SHAPE.ROUNDED_RECTANGLE,radius=0.06)
    badge(s,xx+Inches(0.2),y+Inches(0.24),str(i+1),d=Inches(0.44))
    tt=tb(s,xx+Inches(0.2),y+Inches(0.9),cw-Inches(0.4),Inches(2.0)); par(tt,tl,13.5,True,WHITE,first=True,after=6); par(tt,ds,10.8,False,RGBColor(0xC7,0xCC,0xD3),after=0)
t=tb(s,MX,Inches(5.7),CW,Inches(0.9)); par(t,"Prototype reads: HHI ≈ 1,825 (moderate) · top two vendors ~50% of spend · a cost-per-quality outlier flagged for renegotiation.",13,True,ORANGE,first=True,after=4)
par(t,"All advisory — it informs the decision; a human makes the call.",11.5,False,MUTE,after=0)
notes(s,"The slide procurement leadership remembers. Instead of only knowing WHO our vendors are, we know what each is WORTH, what depends on them, and where the leverage is. This directly arms their next renewal. Use the concrete reads but label them illustrative pending the baseline audit.")

# ---------- 6 Decision layer
s=slide(); bg(s,SURFACE)
header(s,"The decision layer","Cross-tool signals and a natural-language ask")
card(s,MX,Inches(2.0),Inches(6.0),Inches(4.4))
tt=tb(s,MX+Inches(0.3),Inches(2.2),Inches(5.4),Inches(0.5)); par(tt,"Signals — ranked by what's at stake",14.5,True,INK,first=True,after=0)
sig=[(RED,"Agreement at risk feeds a Tier-1 KPI","Sole source behind New-to-Brand Rx Share expires in ~21 days."),
 (ORANGE,"Duplicate-buy risk at intake","A request overlaps a catalogued asset — ~$640K avoidable."),
 (ORANGE,"Low-quality asset feeds a strategic KPI","Bronze data behind a Global funnel KPI."),
 (BLUE,"Vendor concentration","Top two vendors ~50% of spend; HHI ≈ 1,825.")]
yy=Inches(2.75)
for col,tl,ds in sig:
    rect(s,MX+Inches(0.3),yy+Inches(0.05),Inches(0.12),Inches(0.78),fill=col,shape=MSO_SHAPE.ROUNDED_RECTANGLE,radius=0.5)
    t=tb(s,MX+Inches(0.55),yy,Inches(5.2),Inches(0.9)); par(t,tl,12.5,True,INK,first=True,after=2); par(t,ds,11,False,TEXT2,after=0); yy=yy+Inches(0.9)
rect(s,MX+Inches(6.3),Inches(2.0),Inches(5.5),Inches(4.4),fill=INK2,shape=MSO_SHAPE.ROUNDED_RECTANGLE,radius=0.05)
tt=tb(s,MX+Inches(6.6),Inches(2.25),Inches(4.9),Inches(0.5)); par(tt,"ASK ATLAS",12,True,ORANGE,first=True,after=0)
rect(s,MX+Inches(6.6),Inches(2.7),Inches(4.9),Inches(0.55),fill=RGBColor(0x16,0x24,0x2e),line=STROKE,lw=1.0,shape=MSO_SHAPE.ROUNDED_RECTANGLE,radius=0.2)
t=tb(s,MX+Inches(6.8),Inches(2.7),Inches(4.5),Inches(0.55),anchor=MSO_ANCHOR.MIDDLE); par(t,"Which agreements expire soon and feed a Tier-1 KPI?",11.5,False,RGBColor(0xC7,0xCC,0xD3),first=True,after=0)
t=tb(s,MX+Inches(6.6),Inches(3.5),Inches(4.9),Inches(2.4))
par(t,"Atlas answers across all five tools — grounded in the data, with a link straight to the record.",12.5,False,RGBColor(0xE8,0xEA,0xEE),first=True,after=12)
par(t,"Every answer shows its confidence and an advisory cue:",11.5,False,MUTE,after=8)
par(t,'"Confidence: High · grounded in source rows · Advisory only — a human approves any action."',11.5,True,RGBColor(0x7F,0xD1,0xA8),after=0)
footer(s,pg())
notes(s,"Where the pieces become a decision layer. Left: it pushes the few issues that need a decision; right: ask anything in plain English. The trust message for pharma: grounded (cites sources, no invented numbers), shows confidence, advisory only — human approves. No autonomous actions.")

# ---------- 7 Business value
s=slide(); bg(s,SURFACE)
header(s,"Business value","Five outcomes leadership can point to")
val=[("Stop duplicate spend","Overlap caught before commitment — direct, recurring savings.","Procurement"),
 ("Never miss a renewal","Automated reminders plus impact-aware warnings end lapses.","Legal / Continuity"),
 ("Negotiate from strength","Vendor value and concentration analytics inform every renewal.","Procurement / CI"),
 ("Audit-ready by default","Automatic classification and lineage make governance the default.","Compliance"),
 ("Faster, confident decisions","One source of truth and a plain-English ask cut manual lookups.","Commercial Insights")]
cw=Inches(5.85)
for i,(tl,ds,owner) in enumerate(val):
    col=i%2; rowi=i//2; xx=MX+col*(cw+Inches(0.3)); yy=Inches(2.0)+rowi*Inches(1.18)
    w=cw
    if i==4: xx=MX; yy=Inches(2.0)+2*Inches(1.18); w=Inches(12.0)
    card(s,xx,yy,w,Inches(1.05)); badge(s,xx+Inches(0.22),yy+Inches(0.3),str(i+1),d=Inches(0.46),fill=GREEN,fg=WHITE)
    tt=tb(s,xx+Inches(0.95),yy+Inches(0.16),w-Inches(3.0),Inches(0.8),anchor=MSO_ANCHOR.MIDDLE); par(tt,tl,14.5,True,INK,first=True,after=2); par(tt,ds,11.3,False,TEXT2,after=0)
    ow=tb(s,xx+w-Inches(1.9),yy+Inches(0.3),Inches(1.7),Inches(0.45),anchor=MSO_ANCHOR.MIDDLE); par(ow,owner.upper(),9.5,True,GREEN,align=PP_ALIGN.RIGHT,first=True,after=0)
footer(s,pg())
notes(s,"Outcomes, each tied to the executive who owns it. Lead with the two most quantifiable — duplicate spend and renewals. If asked for ROI numbers, be honest: the prototype shows the mechanism; the baseline audit quantifies it against their real estate.")

# ---------- 8 Architecture + roadmap
s=slide(); bg(s,SURFACE)
header(s,"Delivery","Your Microsoft estate, orchestrated — rolled out in phases")
card(s,MX,Inches(2.0),Inches(5.6),Inches(4.4))
tt=tb(s,MX+Inches(0.3),Inches(2.2),Inches(5.0),Inches(0.5)); par(tt,"Target stack — Microsoft-native",15,True,INK,first=True,after=8)
stack=[("SharePoint","Storage & lists — the records system"),("Power Automate","Workflow & the 30/60/90 reminders"),
 ("Microsoft Purview","Catalogue, classification & lineage"),("Azure OpenAI","Intelligence & Ask Atlas"),
 ("Databricks / Model N / SAP","Data, spend and contract feeds")]
yy=Inches(2.8)
for nm,ds in stack:
    rect(s,MX+Inches(0.3),yy+Inches(0.04),Inches(0.1),Inches(0.5),fill=ORANGE,shape=MSO_SHAPE.ROUNDED_RECTANGLE,radius=0.5)
    t=tb(s,MX+Inches(0.55),yy,Inches(4.7),Inches(0.6)); par(t,nm,12.5,True,INK,first=True,after=1); par(t,ds,10.8,False,TEXT2,after=0); yy=yy+Inches(0.66)
card(s,MX+Inches(5.9),Inches(2.0),Inches(6.1),Inches(4.4))
tt=tb(s,MX+Inches(6.2),Inches(2.2),Inches(5.5),Inches(0.5)); par(tt,"Phased rollout — value at every step",15,True,INK,first=True,after=8)
ph=[("PHASE 1","Foundation","Connect the five tools and the data model on SharePoint.",BLUE),
 ("PHASE 2","Automation","Reminders, catalogue sync, classification, duplicate detection.",ORANGE),
 ("PHASE 3","Intelligence","Vendor Hub Plus, prediction and Ask Atlas on Azure OpenAI.",GREEN)]
yy=Inches(2.8)
for tag,tl,ds,col in ph:
    rect(s,MX+Inches(6.2),yy+Inches(0.02),Inches(0.1),Inches(0.85),fill=col,shape=MSO_SHAPE.ROUNDED_RECTANGLE,radius=0.5)
    t=tb(s,MX+Inches(6.45),yy,Inches(5.3),Inches(1.0)); par(t,tag,10,True,col if col!=ORANGE else ORANGE_D,first=True,after=1); par(t,tl,15,True,INK,after=2); par(t,ds,11,False,TEXT2,after=0); yy=yy+Inches(1.05)
footer(s,pg())
notes(s,"Reassure IT and de-risk the programme in one slide. Left: we orchestrate tools you already own — minimal new procurement, security or training. Right: phased delivery means working value from Phase 1, each phase proving the next. Talk phases, not fixed dates, until scoping.")

# ---------- 9 Next steps
s=slide(); bg(s,INK)
rect(s,0,0,SW,Inches(0.16),fill=ORANGE)
t=tb(s,MX+Inches(0.45),Inches(0.7),Inches(10),Inches(0.4)); par(t,"NEXT STEPS",13,True,ORANGE,first=True,after=0)
t=tb(s,MX,Inches(1.1),Inches(11.8),Inches(0.9)); par(t,"How we start",36,True,WHITE,first=True,after=0)
steps=[("1","Walk the prototype live","A hands-on session with your CI and procurement leads."),
 ("2","Baseline data audit","Point Atlas at your real estate — turn illustrative figures into your numbers."),
 ("3","Confirm scope & architecture","Agree Phase 1 boundaries and the Microsoft target with IT."),
 ("4","Stand up Phase 1","Migrate the connected five-tool foundation — value from day one.")]
y=Inches(2.3)
for n,tl,ds in steps:
    badge(s,MX,y,n,d=Inches(0.5),fill=ORANGE,fg=INK)
    t=tb(s,MX+Inches(0.8),y-Inches(0.02),Inches(11.0),Inches(0.95)); par(t,tl,17,True,WHITE,first=True,after=2); par(t,ds,12.5,False,RGBColor(0xC7,0xCC,0xD3),after=0); y=y+Inches(1.02)
t=tb(s,MX,Inches(6.55),Inches(11.8),Inches(0.5)); par(t,"Synthetic prototype · figures directional pending a baseline audit · AI advisory, human-approved.",10.5,False,MUTE,first=True,after=0)
notes(s,"Close on a low-commitment ask: the next step is just a live walkthrough — easy yes. The baseline audit is the hook that unlocks budget by converting illustrative numbers into their real exposure. End on the product: open the prototype.")

prs.save("/home/user/datasharing/Cortex-Atlas-Executive-Cut.pptx")
print("Exec cut saved with", len(prs.slides._sldIdLst), "slides")
