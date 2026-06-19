#!/usr/bin/env python3
"""Builds the Cortex Atlas User Training Guide (.docx)."""
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

INK=RGBColor(0x09,0x13,0x1B); ORANGE=RGBColor(0xC8,0x6A,0x12); ORANGE_BAR=RGBColor(0xF8,0x97,0x39)
GREY=RGBColor(0x59,0x59,0x59); BLUE=RGBColor(0x05,0x63,0xC1)
FONT="Calibri"
doc=Document()

# base font
st=doc.styles['Normal']; st.font.name=FONT; st.font.size=Pt(10.5); st.font.color.rgb=RGBColor(0x20,0x2a,0x33)
for nm,clr,sz in [('Heading 1',INK,17),('Heading 2',ORANGE,13.5),('Heading 3',INK,11.5)]:
    s=doc.styles[nm]; s.font.name=FONT; s.font.color.rgb=clr; s.font.size=Pt(sz)

def shade(p,color="09131B"):
    pPr=p._p.get_or_add_pPr(); sh=OxmlElement('w:shd'); sh.set(qn('w:val'),'clear'); sh.set(qn('w:fill'),color); pPr.append(sh)

def h1(t): doc.add_heading(t,level=1)
def h2(t): doc.add_heading(t,level=2)
def h3(t): doc.add_heading(t,level=3)
def p(t,bold=False,italic=False,color=None,size=None):
    par=doc.add_paragraph(); r=par.add_run(t); r.font.bold=bold; r.font.italic=italic; r.font.name=FONT
    if color: r.font.color.rgb=color
    if size: r.font.size=Pt(size)
    return par
def bullet(t):
    par=doc.add_paragraph(style='List Bullet'); par.add_run(t); return par
def steps(items):
    for it in items:
        par=doc.add_paragraph(style='List Number'); par.add_run(it)
def note(t):
    par=doc.add_paragraph(); par.paragraph_format.left_indent=Inches(0.15)
    r=par.add_run("Note  "); r.font.bold=True; r.font.color.rgb=ORANGE; r.font.name=FONT
    r2=par.add_run(t); r2.font.italic=True; r2.font.color.rgb=GREY; r2.font.name=FONT
def table(headers,rows):
    t=doc.add_table(rows=1,cols=len(headers)); t.style='Light Grid Accent 1'; t.alignment=WD_TABLE_ALIGNMENT.LEFT
    hc=t.rows[0].cells
    for i,htxt in enumerate(headers):
        hc[i].text=''; r=hc[i].paragraphs[0].add_run(htxt); r.font.bold=True; r.font.size=Pt(9.5); r.font.name=FONT
    for row in rows:
        cells=t.add_row().cells
        for i,val in enumerate(row):
            cells[i].text=''; r=cells[i].paragraphs[0].add_run(val); r.font.size=Pt(9.5); r.font.name=FONT
    doc.add_paragraph()

# ===== Title page =====
bar=doc.add_paragraph(); shade(bar,"09131B"); bar.paragraph_format.space_after=Pt(0)
rb=bar.add_run("  CORTEX ATLAS"); rb.font.bold=True; rb.font.size=Pt(26); rb.font.color.rgb=RGBColor(0xFF,0xFF,0xFF); rb.font.name=FONT
sb=doc.add_paragraph(); shade(sb,"09131B"); sb.paragraph_format.space_before=Pt(0)
rs=sb.add_run("  User Training Guide"); rs.font.size=Pt(15); rs.font.color.rgb=RGBColor(0xF8,0x97,0x39); rs.font.name=FONT
doc.add_paragraph()
p("A complete, step-by-step guide to using Cortex Atlas - the connected workspace for Commercial Insights & Data Procurement.",size=12,color=GREY)
p("Audience: requesters, reviewers (regional / global / admin), data-catalogue stewards, KPI owners, and vendor & agreement managers.",size=10.5,color=GREY)
doc.add_paragraph()
p("Version 1.0  -  Illustrative prototype (synthetic data)  -  Confidential",size=9.5,color=GREY,italic=True)

# ===== Contents =====
doc.add_page_break(); h1("Contents")
for i,t in enumerate([
 "1. Introduction","2. Getting started","3. Key concepts you need","4. Command Centre",
 "5. Data Requests (DPRH)","6. Data Catalogue","7. KPI Catalogue","8. Vendor Hub",
 "9. Vendor Hub Plus","10. TPA Hub","11. Using Ask Atlas","12. Tips & good practice",
 "13. Frequently asked questions","14. Glossary","15. Data & deployment notes"]):
    bullet(t)

# ===== 1 Introduction =====
doc.add_page_break(); h1("1. Introduction")
h2("1.1  What Cortex Atlas is")
p("Cortex Atlas is a single, connected workspace for the data your team buys and governs. It brings together the five tools you use today - the data purchase request hub, the data catalogue, the KPI catalogue, the vendor hub and the third-party-agreement (TPA) register - so you can request, record, measure, source and govern commercial data in one place instead of moving between separate SharePoint forms and lists.")
h2("1.2  What makes it different")
bullet("It is connected. A purchase request becomes a catalogue asset, which is measured by KPIs, supplied by a vendor and governed by an agreement - and every link is clickable.")
bullet("It does work for you. Requests route themselves, renewals pre-fill, sensitivity is auto-classified, and agreement reminders fire automatically.")
bullet("It gives you answers. It flags duplicate purchases and quality risks, analyses vendor value, predicts renewals, and answers plain-English questions - all advisory, with a human approving any action.")
h2("1.3  The five tools at a glance")
table(["Module","What it is","You use it to"],[
 ["Data Requests (DPRH)","The purchase-request workflow","Raise and approve a data purchase"],
 ["Data Catalogue","The record of everything purchased","Check what exists; add and publish assets"],
 ["KPI Catalogue","The registry of metrics","See which data each KPI relies on"],
 ["Vendor Hub","The vendor inventory","Look up vendors and what we buy"],
 ["Vendor Hub Plus","The vendor value layer","Assess vendor quality, risk and leverage"],
 ["TPA Hub","The agreement register","Track agreements and renewals"],
])

# ===== 2 Getting started =====
doc.add_page_break(); h1("2. Getting started")
h2("2.1  Opening Atlas")
p("Open the tool in any modern web browser. It loads straight to the Command Centre. (In this prototype it runs entirely in the browser on synthetic data; on deployment it will sit in your Microsoft environment.)")
h2("2.2  The screen explained")
p("Every screen has the same four areas plus the assistant:")
bullet("Top bar (dark): the module title, a Live badge, and the Division and Financial-Year filters that apply across the whole tool. The Suite button returns to the wider Cortex suite.")
bullet("Left navigation (dark): grouped into Command Centre, The Estate (the five tools) and Intelligence (Vendor Hub Plus). Click an item to open that module.")
bullet("Working surface (light): the content of the module you are in - headline tiles, charts and tables.")
bullet("Detail drawer: slides in from the right when you click a table row. It holds the full record, its tabs and any actions.")
bullet("Ask Atlas: a question box on the Command Centre that reads across all five tools.")
h2("2.3  Global filters")
p("Two selectors in the top bar filter the tables and charts in every module:")
bullet("Division - All divisions, Oncology, Immunology, Cardiology, Neurology or Rare Disease. Choosing a division narrows every module to that area; cross-brand items always remain visible.")
bullet("Financial Year - the reporting year context.")
note("The filters are global. If a table looks shorter than expected, check whether a Division filter is applied.")
h2("2.4  Moving around and closing the drawer")
steps([
 "Click a left-nav item to switch module.",
 "Click any table row to open its detail drawer.",
 "Use the blue links and clickable rows inside a drawer to jump to a related record (the drawer closes as you navigate).",
 "Close the drawer with the X button, by clicking the dark area outside it, or by pressing the Escape key.",
])

# ===== 3 Key concepts =====
doc.add_page_break(); h1("3. Key concepts you need")
h2("3.1  The connected estate")
p("Think of one data asset moving through five tools: it is requested (DPRH), recorded (Catalogue), measured (KPIs), supplied by a vendor (Vendor Hub) and governed by an agreement (TPA). Atlas keeps these connected, so you can trace any item to everything related to it.")
h2("3.2  The three layers")
bullet("Estate - the faithful as-is records and workflow (your source of truth).")
bullet("Automation - the work the tool does for you (routing, pre-fill, classification, reminders).")
bullet("Intelligence - the decision support (duplicate and quality flags, vendor analytics, renewal prediction, Ask Atlas).")
h2("3.3  Roles - who sees and does what")
p("In Data Requests, what you can see and do depends on your role. Use the role switch at the top of a request to view it as each role (in the live system your role is set by your access group).")
table(["Role","Sees","Acts on"],[
 ["Requester","The request and data-detail sections","Raises and submits the request"],
 ["Admin","The whole request","Routes it, approves/rejects, and final approval"],
 ["Regional reviewer","Requests for their own region","Submits the regional review"],
 ["Global reviewer","Requests at committee stage","Takes part in the CDS committee review"],
])
note("The CDS committee review is done by regional reviewers, global reviewers and admin together.")
h2("3.4  Request statuses and what they mean")
table(["Status","Meaning"],[
 ["Data sourcing draft","Saved but not yet submitted"],
 ["Submitted for offline review","A sub-national request, handled offline by local teams"],
 ["Sub-national data for contracting by local teams","Sub-national request routed to local contracting"],
 ["In data sourcing pre review","National request above the cost threshold, awaiting admin"],
 ["Submitted for CILT review","National request below the cost threshold"],
 ["Submitted for validation","Admin has routed it into the review chain"],
 ["In CILT CDSP review","With the CDS committee"],
 ["Recommendation available","Reviews complete, awaiting admin final decision"],
 ["Approved in contracting","Approved; contracting under way"],
 ["Approved purchase completed","Approved and purchased"],
 ["Rejected","Not approved"],
])
h2("3.5  Routing rules and cost thresholds")
p("When a request is submitted, Atlas sets its status automatically from two things:")
bullet("National or sub-national. Sub-national requests go to offline review with local teams and skip the committee chain.")
bullet("Anticipated cost versus the division's cost threshold. Above the threshold a national request becomes In data sourcing pre review; below it becomes Submitted for CILT review.")
p("Thresholds differ by region type:")
table(["Region type","Cost threshold (USD)"],[
 ["APUS established","250,000"],["Global","500,000"],["China","150,000"],["International markets","300,000"],
])
h2("3.6  Quality tiers and sensitivity")
bullet("Quality tier - each catalogue asset is graded Gold (high), Silver (medium) or Bronze (low).")
bullet("PII - personally identifiable information (e.g. HCP-level data).")
bullet("PHI - protected health information (e.g. patient-level data). PHI assets carry handling controls.")
h2("3.7  Advisory AI and human approval")
p("Every AI or analytics output in Atlas - duplicate alerts, quality flags, vendor analytics, renewal predictions and Ask Atlas answers - is advisory. Each shows a confidence read and an 'a human approves any action' cue. Atlas never approves, buys or changes anything on its own.")

# ===== 4 Command Centre =====
doc.add_page_break(); h1("4. Command Centre")
p("The Command Centre is your landing page. Use it to see the state of the estate and what needs attention.")
h3("What it shows")
bullet("Headline tiles - catalogued assets, contracted spend, in-flight requests and agreements expiring within 90 days.")
bullet("The connected estate - the five tools as clickable tiles.")
bullet("Cross-tool signals - the most important issues that span more than one tool, ranked by what is at stake.")
bullet("Ask Atlas - a question box across the whole estate.")
h3("Reading a signal and acting on it")
steps([
 "Scan the signals; each has a severity tag (High / Medium / Watch).",
 "Read the one-line explanation - for example, an agreement expiring soon that feeds a strategic KPI.",
 "Click the signal (or the blue link inside it) to jump straight to the record involved.",
])

# ===== 5 Data Requests =====
doc.add_page_break(); h1("5. Data Requests (DPRH)")
p("This is the front door for buying data. One combined, role-gated form raises and governs every purchase.")
h2("5.1  Raise a new request")
steps([
 "Open Data Requests from the left nav.",
 "Click New data request.",
 "Choose the Commercial division and Region type (the region type sets the cost threshold).",
 "Choose National or Sub-national, and the Purchase type (New or Continuing).",
 "Enter the Data name, choose the Vendor, the Data type and the Anticipated cost.",
 "Watch the Live routing & checks panel: it shows the status the request will get, and flags any possible duplicate against the catalogue.",
 "Add a short business rationale and click Submit request.",
])
note("If the panel shows a possible duplicate, check the existing asset before continuing - you may already own what you are about to buy.")
h2("5.2  Renew an existing asset")
steps([
 "Click New data request and set Purchase type to Continuing Purchase.",
 "Choose the Data being renewed from the list of catalogue assets.",
 "The form pre-fills the vendor, prior cost and data type from the catalogue record.",
 "Adjust the anticipated cost if needed and submit.",
])
h2("5.3  Understand the request record")
p("Open any request to see its full record. It has a role switch, an Actions area, an auto-routing summary, the approval chain, and tabbed sections:")
bullet("Notes - requester notes and the latest update.")
bullet("Request - DPRH - submitter, division, data details, vendor, cadence, threshold, dates and rationale.")
bullet("Scope, cost & alternatives - market scope, costs, licences and any alternative data considered.")
bullet("Admin review, Regional reviewer, CDS committee review, Admin final - each visible to its role, and each opens only when the request reaches that step.")
h2("5.4  Review and act on a request")
p("Use the role switch at the top of the request to act as the relevant role. The Actions area shows only the buttons valid for your role at the current status.")
h3("As Admin (first decision)")
steps([
 "Open a national request that is In data sourcing pre review or Submitted for CILT review.",
 "Switch the role to Admin.",
 "Choose an action: Route to validation (sends it into the regional/committee chain), Approve in contracting, Approve purchase completed, or Reject.",
])
h3("As Regional reviewer")
steps([
 "Open a request that is Submitted for validation.",
 "Switch the role to Regional.",
 "Click Submit regional review to advance it to the CDS committee.",
])
h3("As the CDS committee (Global / Regional / Admin)")
steps([
 "Open a request that is In CILT CDSP review.",
 "Switch the role to Global (or Regional/Admin).",
 "Click Submit committee recommendation to advance it to Recommendation available.",
])
h3("Admin final decision")
steps([
 "Open a request that is Recommendation available.",
 "Switch the role to Admin.",
 "Click Approve in contracting or Approve purchase completed (or Reject).",
])
note("If you see 'no actions for this role at this status', switch to the role that owns the current step - usually Admin.")
h2("5.5  Move an approved request to the catalogue")
steps([
 "Open an approved request (Approved in contracting or Approved purchase completed).",
 "Click Move to Data Catalogue.",
 "Atlas creates a draft asset in the Data Catalogue, carrying over the captured details.",
 "Go to the Data Catalogue to complete and publish it (see section 6.4).",
])
h2("5.6  Sub-national requests")
p("Sub-national requests are submitted for offline review and handled by local teams for contracting. They do not enter the regional or committee chain. Admin can route them to local contracting or approve/reject them.")
h2("5.7  Duplicate-buy alerts")
p("When a new request overlaps an asset already in the catalogue (same vendor, therapeutic area and data type), Atlas flags it - on the request list, on the Command Centre, and inside the request - so you can review before any spend is committed.")

# ===== 6 Data Catalogue =====
doc.add_page_break(); h1("6. Data Catalogue")
p("The central record of every data asset purchased - and the place to check before buying.")
h2("6.1  Check whether data already exists")
steps([
 "Open the Data Catalogue.",
 "Scan or filter the asset table (use the Division filter to narrow it).",
 "Open a candidate asset to confirm vendor, scope and quality before raising a new request.",
])
h2("6.2  Read an asset record")
p("Click a row to open the full record, grouped into Identity, Classification, Coverage, Quality & sensitivity, Commercial and Access & usage. The Lineage strip at the top links to the vendor, the KPIs it feeds and its agreement.")
h2("6.3  Add a manual entry")
steps([
 "In the Data Catalogue, click New manual entry.",
 "Enter the asset name and choose the vendor, therapeutic area, data archetype and lowest granularity.",
 "Watch the Auto-classification panel: Atlas detects PII/PHI and suggests a quality tier from what you enter.",
 "Add the purchase cost and origin, then click Save to Catalogue.",
])
note("Auto-classification is a suggestion to speed you up; a steward should confirm sensitivity and quality before the asset is relied upon.")
h2("6.4  Complete and publish a draft")
p("Assets that arrive from an approved request appear with a Draft tag.")
steps([
 "Open a draft asset (it shows a Draft banner).",
 "Review the carried-over details.",
 "Click Complete & publish to add it to the live catalogue, KPI lineage and vendor rollups.",
])
h2("6.5  Quality flags")
p("If a low-tier (Bronze) asset feeds a strategic (Global) KPI, the catalogue surfaces a quality flag. Treat outputs from that KPI as directional and consider a quality uplift before using the figure externally.")

# ===== 7 KPI Catalogue =====
doc.add_page_break(); h1("7. KPI Catalogue")
p("The registry of KPIs and metrics built on the data, each linked back to its source.")
h2("7.1  Read a KPI and its lineage")
steps([
 "Open the KPI Catalogue and click a KPI.",
 "Read its definition, methodology, market definition, cadence and contacts.",
 "Use the Lineage strip to see the source asset(s) and vendor(s) behind the number, and follow the links to those records.",
])
h2("7.2  Add a KPI or metric")
steps([
 "In the KPI Catalogue, click Add KPI / metric.",
 "Enter the name and choose KPI or Metric, the scope (Global / Regional / Local) and the brand.",
 "Choose the Source asset - the Live lineage check shows the vendor behind it and flags any quality risk.",
 "Set the cadence, add a description, and click Add to KPI Catalogue.",
])
h2("7.3  Quality risk")
p("If you give a Global KPI a Bronze source asset, Atlas flags the quality risk as you build it and on the KPI list. This tells everyone how much to trust the metric.")

# ===== 8 Vendor Hub =====
doc.add_page_break(); h1("8. Vendor Hub")
p("The vendor-centric inventory: who supplies your data and what you buy from each.")
h2("8.1  Look up a vendor")
steps([
 "Open Vendor Hub and click a vendor.",
 "Review what they hold, how to access it, contacts and guidance.",
])
h2("8.2  Purchased vs available")
p("Each vendor shows what we have purchased versus everything they offer. Untapped offerings are a consolidation or negotiation opportunity. For the full value and risk analytics, click Open Vendor Hub Plus.")

# ===== 9 Vendor Hub Plus =====
doc.add_page_break(); h1("9. Vendor Hub Plus")
p("The vendor value and risk layer. It rolls the catalogue up to the vendor level and recomputes automatically as the catalogue changes. All output is advisory.")
table(["Layer","What it tells you","How to read it"],[
 ["Quality score","Each vendor's Gold/Silver/Bronze mix and a composite grade","Higher grade = more reliable data"],
 ["Spend concentration","Each vendor's share of spend, plus an estate-wide HHI","Higher HHI = more concentrated, more single-vendor risk (above ~2,500 is highly concentrated)"],
 ["Irreplaceability","Assets and KPIs that depend on a vendor vs available substitutes","Critical = high dependency, few substitutes"],
 ["Cost-per-quality","Spend normalised against quality points","A high value on low-tier data is a renegotiation candidate"],
 ["Leverage","Purchased vs available offerings","Large gap = room to consolidate or negotiate"],
])
p("Click any vendor row to open its full record. Use these as inputs to a renewal or negotiation decision - a person makes the call.")

# ===== 10 TPA Hub =====
doc.add_page_break(); h1("10. TPA Hub")
p("The third-party-agreement register, with an automatic 30/60/90-day reminder engine.")
h2("10.1  The expiry runway and reminders")
p("The runway shows how many agreements expire within 30, 31-60 and 61-90 days, and beyond. Each agreement shows which reminders (90/60/30) have fired; inside 30 days it escalates to the owner.")
h2("10.2  Read an agreement")
steps([
 "Open TPA Hub and click an agreement.",
 "Review the agreement, vendor and data, scope and brand, third party, dates and ownership.",
 "Use the Lineage strip to see the asset and any KPIs that depend on it.",
])
h2("10.3  Renewal prediction and cross-tool warnings")
p("Each agreement carries an advisory renewal recommendation (Renew / Review). A cross-tool warning appears when an expiring agreement sits behind a strategic KPI - so you act before leverage or continuity is lost.")
h2("10.4  Create a renewal request")
steps([
 "Open an agreement that is due for renewal.",
 "Click Create renewal request.",
 "Atlas raises a pre-filled continuing-purchase request in Data Requests, carrying the vendor and asset across.",
 "Switch to Data Requests to review and submit it through the workflow.",
])

# ===== 11 Ask Atlas =====
doc.add_page_break(); h1("11. Using Ask Atlas")
p("Ask Atlas, on the Command Centre, answers questions across all five tools.")
steps([
 "Go to the Command Centre.",
 "Type a question in plain English, or tap one of the suggested questions.",
 "Read the answer; follow the blue links to the records it references.",
 "Check the source chips (the feeds the answer draws on) and the confidence read.",
])
p("Example questions you can ask:")
bullet("Which agreements expire soon and feed a Tier-1 KPI?")
bullet("Where might we be buying data we already own?")
bullet("Which vendor carries the most concentration risk?")
bullet("Are any low-quality assets feeding strategic KPIs?")
note("Ask Atlas is advisory. It cites its sources and shows its confidence; it does not act on its own, and a human approves any action.")

# ===== 12 Tips =====
doc.add_page_break(); h1("12. Tips & good practice")
bullet("Always check the catalogue (or raise the request and read the duplicate panel) before buying new data.")
bullet("Submit renewals from the TPA Hub or as Continuing purchases so the form pre-fills and stays consistent.")
bullet("When reviewing, set your role first - the available actions follow your role and the request's status.")
bullet("Complete and publish catalogue drafts promptly so KPIs and vendor analytics stay accurate.")
bullet("Treat any KPI with a quality flag as directional until the source data is improved.")
bullet("Use the Division filter to focus on your area; clear it to see the whole estate.")

# ===== 13 FAQ =====
doc.add_page_break(); h1("13. Frequently asked questions")
faq=[
 ("Why can I not see any action buttons on a request?","The actions depend on your role and the request's current status. Switch to the role that owns the current step (often Admin)."),
 ("A table looks empty or short - why?","A Division filter is probably applied in the top bar. Set it back to All divisions."),
 ("What is the difference between PII and PHI?","PII is personally identifiable information (e.g. HCP-level); PHI is protected health information (e.g. patient-level). PHI carries extra handling controls."),
 ("What does the HHI number mean?","It measures how concentrated your spend is across vendors. Higher means more reliance on a few vendors; above about 2,500 is highly concentrated."),
 ("Is anything decided automatically?","No. Routing and classification are automated to save effort, but all approvals and any AI-suggested actions require a person."),
 ("My changes disappeared after refreshing the page.","In this prototype, data is held for your session and resets on refresh. On deployment, records persist in the backend."),
]
for q,a in faq:
    p(q,bold=True,color=INK); p(a,color=GREY); doc.add_paragraph()

# ===== 14 Glossary =====
doc.add_page_break(); h1("14. Glossary")
table(["Term","Meaning"],[
 ["DPRH / CDSP","The Data Purchase Request Hub - the request workflow"],
 ["Catalogue asset","A purchased data asset recorded in the Data Catalogue"],
 ["Draft asset","A catalogue entry created from an approved request, awaiting completion and publishing"],
 ["KPI / Metric","A measure built on purchased data"],
 ["Lineage","The chain from a KPI to its source asset and vendor"],
 ["TPA","Third-Party Agreement governing data use and sharing"],
 ["Runway","The view of agreements by time-to-expiry (30/60/90 days)"],
 ["HHI","Herfindahl index - a measure of spend concentration across vendors"],
 ["Gold / Silver / Bronze","Data quality tiers (high / medium / low)"],
 ["CDS committee","Central review by regional + global reviewers + admin"],
])

# ===== 15 Data & deployment =====
doc.add_page_break(); h1("15. Data & deployment notes")
p("This guide describes a working prototype that runs in the browser on synthetic, illustrative data (the fictional company NovaCura Pharmaceuticals; invented vendors; public vendor names appear only as integration labels). No real client or patient data is used, and figures are directional.")
p("On deployment, Atlas is designed to run on your Microsoft environment: SharePoint for records, Power Automate for workflow and the 30/60/90-day reminders, Microsoft Purview for classification and lineage, and Azure OpenAI for the Ask Atlas assistant. At that point, records persist centrally and roles are governed by your access groups.")
p("All AI and analytics output remains advisory, with a confidence read and human approval - by design.",italic=True,color=GREY)

doc.save("/home/user/datasharing/Cortex-Atlas-User-Training-Guide.docx")
print("Training guide saved")
