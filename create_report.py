#!/usr/bin/env python3
"""
Generate LoopForgeLab Competitor Report as Word document — English version.
"""

import io
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ── Brand colours ─────────────────────────────
DARK_BLUE  = RGBColor(0x0D, 0x2B, 0x4A)
MID_BLUE   = RGBColor(0x1A, 0x5F, 0x8A)
ACCENT     = RGBColor(0x00, 0xA8, 0xCC)
TEXT_DARK  = RGBColor(0x1A, 0x1A, 0x2E)
WHITE      = RGBColor(0xFF, 0xFF, 0xFF)
PIE_COLS   = ["#1A5F8A", "#00A8CC", "#F0A500", "#2ECC71"]


# ── Cell helpers ──────────────────────────────
def shade_cell(cell, hex_color: str):
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd  = OxmlElement('w:shd')
    shd.set(qn('w:val'),   'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'),  hex_color.lstrip('#'))
    tcPr.append(shd)


def hdr_cell(cell, text, bg='0D2B4A', fs=8, center=True):
    cell.text = text
    shade_cell(cell, bg)
    for para in cell.paragraphs:
        if center:
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in para.runs:
            run.font.bold = True
            run.font.color.rgb = WHITE
            run.font.size = Pt(fs)


def data_cell(cell, text, bg='FFFFFF', fs=8, bold=False, color=None, center=False):
    cell.text = text
    shade_cell(cell, bg)
    for para in cell.paragraphs:
        if center:
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in para.runs:
            run.font.size  = Pt(fs)
            run.font.bold  = bold
            if color:
                run.font.color.rgb = color


# ── Paragraph helpers ─────────────────────────
def add_para(doc, text, bold=False, size=10, color=None,
             align=None, sb=0, sa=3):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(sb)
    p.paragraph_format.space_after  = Pt(sa)
    if align:
        p.alignment = align
    r = p.add_run(text)
    r.bold = bold
    r.font.size = Pt(size)
    r.font.color.rgb = color or TEXT_DARK
    return p


def add_h1(doc, text):
    p = doc.add_heading(text, level=1)
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after  = Pt(4)
    for run in p.runs:
        run.font.color.rgb = DARK_BLUE
        run.font.size = Pt(14)


def add_h2(doc, text, color=None):
    p = doc.add_heading(text, level=2)
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after  = Pt(3)
    for run in p.runs:
        run.font.color.rgb = color or MID_BLUE
        run.font.size = Pt(11)


def add_h3(doc, text):
    p = doc.add_heading(text, level=3)
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after  = Pt(2)
    for run in p.runs:
        run.font.color.rgb = MID_BLUE
        run.font.size = Pt(10)


def add_bullet(doc, text, size=9):
    p = doc.add_paragraph(style='List Bullet')
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after  = Pt(1)
    r = p.add_run(text)
    r.font.size = Pt(size)
    r.font.color.rgb = TEXT_DARK


def thin_gap(doc, pts=3):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(pts)


# ── Pie chart ─────────────────────────────────
def make_pie():
    labels  = ['Engineering\n270 (72.8%)',
               'Data & Req. Mgmt\n40 (10.8%)',
               'Sales & Quotes\n35 (9.4%)',
               'Sustainability\n26 (7.0%)']
    sizes   = [270, 40, 35, 26]
    explode = [0.02, 0.02, 0.07, 0.02]

    fig, ax = plt.subplots(figsize=(5.5, 3.8), dpi=160)
    wedges, texts, autotexts = ax.pie(
        sizes, labels=labels, colors=PIE_COLS,
        explode=explode, autopct='%1.1f%%',
        startangle=140, pctdistance=0.76,
        textprops={'fontsize': 8, 'color': '#1A1A2E'},
        wedgeprops={'linewidth': 1.5, 'edgecolor': 'white'}
    )
    for at in autotexts:
        at.set_fontsize(7.5)
        at.set_color('white')
        at.set_fontweight('bold')

    ax.set_title('Company Distribution by Category  (CI v1.7 | 371 Companies)',
                 fontsize=9.5, fontweight='bold', color='#0D2B4A', pad=10)
    plt.tight_layout(pad=0.4)
    buf = io.BytesIO()
    plt.savefig(buf, format='png', bbox_inches='tight')
    plt.close()
    buf.seek(0)
    return buf


# ── Generic 3-col provider table ─────────────
def provider_table(doc, rows_data, bg_hdr='1A5F8A'):
    tbl = doc.add_table(rows=1 + len(rows_data), cols=3)
    tbl.style = 'Table Grid'
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    widths = [Cm(2.8), Cm(1.4), Cm(13.1)]
    for j, (h, w) in enumerate(zip(['Company', 'Type', 'Solution / USP'], widths)):
        c = tbl.cell(0, j)
        c.width = w
        hdr_cell(c, h, bg=bg_hdr)
    for i, (name, typ, desc) in enumerate(rows_data):
        bg = 'FFFFFF' if i % 2 == 0 else 'F2F5F8'
        data_cell(tbl.cell(i+1, 0), name, bg=bg, bold=True, color=DARK_BLUE)
        tc = RGBColor(0xC0,0x39,0x2B) if typ == 'Large' else MID_BLUE
        data_cell(tbl.cell(i+1, 1), typ,  bg=bg, color=tc, center=True)
        data_cell(tbl.cell(i+1, 2), desc, bg=bg)
    return tbl


# ══════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════
def build():
    doc = Document()

    # ── Tight margins ─────────────────────────
    for sec in doc.sections:
        sec.top_margin    = Cm(1.8)
        sec.bottom_margin = Cm(1.8)
        sec.left_margin   = Cm(2.4)
        sec.right_margin  = Cm(2.4)

    # ── Default style ─────────────────────────
    ns = doc.styles['Normal']
    ns.font.name  = 'Calibri'
    ns.font.size  = Pt(10)
    ns.font.color.rgb = TEXT_DARK
    ns.paragraph_format.space_before = Pt(0)
    ns.paragraph_format.space_after  = Pt(3)

    # ══ TITLE PAGE ════════════════════════════
    thin_gap(doc, 18)
    add_para(doc, 'LoopForgeLab', bold=True, size=26,
             color=DARK_BLUE, align=WD_ALIGN_PARAGRAPH.CENTER, sb=0, sa=4)
    add_para(doc, 'Competitive Intelligence Report', bold=True, size=18,
             color=MID_BLUE, align=WD_ALIGN_PARAGRAPH.CENTER, sb=0, sa=4)
    add_para(doc, 'Market & Competitive Analysis  ·  Focus: Sales & Quotes', size=12,
             color=ACCENT, align=WD_ALIGN_PARAGRAPH.CENTER, sb=0, sa=6)
    add_para(doc, 'Database: Competitive Intelligence DB v1.7  |  Date: March 2026',
             size=9, color=RGBColor(0x80,0x80,0x90),
             align=WD_ALIGN_PARAGRAPH.CENTER, sb=0, sa=2)
    add_para(doc, 'Prepared for: LoopForgeLab GbR · Berlin',
             size=9, color=RGBColor(0x80,0x80,0x90),
             align=WD_ALIGN_PARAGRAPH.CENTER, sb=0, sa=2)
    doc.add_page_break()

    # ══ 1  EXECUTIVE SUMMARY ══════════════════
    add_h1(doc, '1  Executive Summary & Market Overview')
    add_para(doc, (
        'This Competitive Intelligence Report analyses the competitive landscape for LoopForgeLab — '
        'a Berlin-based deep-tech start-up building an AI-powered Product Intelligence Engine for '
        'mechanical product development. The database covers version 1.7 of the internal CI database, '
        'built on 44 qualitative expert interviews and systematic market research.'
    ), sa=5)

    # KPI table
    add_h2(doc, 'Database at a Glance')
    kpi = doc.add_table(rows=2, cols=4)
    kpi.alignment = WD_TABLE_ALIGNMENT.CENTER
    kpi.style = 'Table Grid'
    for j, (h, v) in enumerate(zip(
            ['Unique Companies', 'Products / Solutions', 'Large Enterprises', 'Start-ups'],
            ['371', '541', '76  (20 %)', '295  (80 %)'])):
        hdr_cell(kpi.cell(0, j), h, fs=9)
        c = kpi.cell(1, j)
        data_cell(c, v, bg='E8F4F8', bold=True, color=DARK_BLUE, center=True)
        for para in c.paragraphs:
            for run in para.runs:
                run.font.size = Pt(13)
    thin_gap(doc, 5)

    # Category table
    add_h2(doc, 'Market Split by Category')
    cat_rows = [
        ('Engineering',              '270','72.8%','380','70.2%','54','216','80%',
         'Factory Intelligence & IoT, BIM & Construction,\nEngineering Copilots & AI Assistants, Generative Design'),
        ('Data & Requirements Mgmt', '40', '10.8%','65', '12.0%','7', '33', '82%',
         'MBSE & Systems Engineering, PLM & PDM Platforms,\nEngineering Collaboration, CAD & Modeling'),
        ('Sales & Quotes',           '35', '9.4%', '39', '7.2%', '13','22', '63%',
         'Supply Chain & Procurement, CPQ, RfP/RfQ Mgmt,\nE-Sourcing, Supply Chain Planning'),
        ('Sustainability',           '26', '7.0%', '57', '10.5%','2', '24', '92%',
         'LCA Software & Platforms, Carbon & Eco-Design,\nCircular Economy, Regulatory & Compliance (DPP)'),
        ('TOTAL',                    '371','100%',  '541','100%', '76','295','80%',''),
    ]
    cat_hdrs = ['Category','# Co.','Share','# Prod.','Share','# Large','# Startup','Startup%','Focus Areas']
    cat_w    = [Cm(3.4),Cm(1.1),Cm(1.2),Cm(1.2),Cm(1.2),Cm(1.2),Cm(1.4),Cm(1.4),Cm(5.4)]
    ctbl = doc.add_table(rows=1+len(cat_rows), cols=9)
    ctbl.style = 'Table Grid'
    ctbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    for j,(h,w) in enumerate(zip(cat_hdrs, cat_w)):
        c = ctbl.cell(0,j); c.width=w; hdr_cell(c,h)
    bgs = ['FFFFFF','F2F5F8','FFFFFF','F2F5F8','E8F4F8']
    for i,row in enumerate(cat_rows):
        for j,val in enumerate(row):
            c = ctbl.cell(i+1,j)
            bold = (i==4)
            col  = DARK_BLUE if bold else None
            data_cell(c, val, bg=bgs[i], bold=bold, color=col,
                      center=(j not in (0,8)), fs=8)
    thin_gap(doc, 5)

    # Pie chart
    add_h2(doc, 'Company Distribution by Category')
    pie_buf = make_pie()
    pp = doc.add_paragraph()
    pp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pp.paragraph_format.space_before = Pt(2)
    pp.paragraph_format.space_after  = Pt(4)
    pp.add_run().add_picture(pie_buf, width=Cm(10))

    add_para(doc, (
        'The Engineering category dominates with 72.8% — reflecting the broad wave of digitisation in CAD, CAE, IoT '
        'and AI-assisted development tools. Sales & Quotes (9.4%) and Sustainability (7.0%) are smaller by headcount '
        'but show high start-up density, signalling intense innovation dynamics and unresolved core problems in these areas.'
    ), sa=4)
    doc.add_page_break()

    # ══ 2  SEGMENT PROFILES ═══════════════════
    add_h1(doc, '2  Market Segments: Short Profile of the Four Categories')
    segments = [
        {
            'title': '2.1  Engineering  (270 Companies | 380 Products)',
            'color': MID_BLUE,
            'body': (
                'The Engineering category is by far the largest and most diverse. It covers all tools that '
                'directly intervene in the product development and manufacturing process: from classical '
                'CAD/CAE/CAM software to AI-powered design copilots, Industrial IoT platforms, BIM solutions '
                'and factory intelligence systems. 80% of the 270 companies are start-ups — a sign of '
                'intense innovation. Established giants (Autodesk, Siemens, Ansys, Dassault) compete '
                'with agile newcomers occupying niches in generative design, digital twins and AI assistance.'
            ),
            'bullets': [
                'Sub-categories: Factory Intelligence & IoT · BIM & Construction · Engineering Copilots & AI Assistants · Generative & Computational Design',
                'LFL relevance: Direct overlap with LoopForgeLab\'s Design Copilot approach; competitors like Autodesk Fusion and Siemens NX are rapidly adding AI capabilities',
            ]
        },
        {
            'title': '2.2  Data & Requirements Management  (40 Companies | 65 Products)',
            'color': MID_BLUE,
            'body': (
                'This category addresses the core problem of fragmented data in product development: PLM '
                'platforms, MBSE tools, engineering collaboration solutions and requirements management '
                'software. With only 7 large companies out of 40 (82% start-ups), dynamism is high. '
                'Established players such as PTC (Windchill), Dassault (ENOVIA) and SAP PLM dominate the '
                'enterprise market, while start-ups pioneer new approaches for digital thread concepts, '
                'collaborative requirements management and AI-driven data aggregation.'
            ),
            'bullets': [
                'Sub-categories: MBSE & Systems Engineering · PLM & PDM Platforms · Engineering Collaboration & Data · CAD & Modeling',
                'LFL relevance: The "Data Desert" finding from the White Paper (fragmented PLM/ERP data) is the central gap that LFL\'s Product Intelligence Engine aims to close',
            ]
        },
        {
            'title': '2.3  Sales & Quotes  (35 Companies | 39 Products)',
            'color': MID_BLUE,
            'body': (
                'Sales & Quotes is the category with the most direct relevance to LoopForgeLab\'s core solution. '
                'It covers solutions around quoting, procurement and supply chain: RfP/Bid-Response Management, '
                'e-Sourcing & RfQ platforms, CPQ systems for complex products and supply chain planning. '
                'At 63% start-up share, the category is more balanced between large and small players. '
                'Despite comparatively fewer companies, the depth of problem-solving is high — and '
                'measurable customer pain is substantial.'
            ),
            'bullets': [
                'Sub-categories: A — RfP/Bid Response Management · B — E-Sourcing/RfQ · C — CPQ (Configure Price Quote) · D — Supply Chain Planning',
                'LFL relevance: LFL\'s RFQ module (in development from Q2 2026) and quote cost analysis directly address problems in sub-clusters B and C',
                'Notable: 39 products from only 35 companies — many providers offer multiple product lines',
            ]
        },
        {
            'title': '2.4  Sustainability  (26 Companies | 57 Products)',
            'color': MID_BLUE,
            'body': (
                'The smallest category by company count but with the highest product density per company '
                '(avg. 2.2 products/company). Sustainability covers LCA software, CO₂ calculation tools, '
                'circular economy platforms and compliance solutions for CSRD/DPP requirements. '
                'With 92% start-up share, this is the most innovative and fastest-growing segment — '
                'driven by EU regulatory pressure. Only 2 established large companies (CarbonBright, '
                'Sphera) have gained a foothold so far.'
            ),
            'bullets': [
                'Sub-categories: LCA Software & Platforms · Carbon & Eco-Design · Circular Economy & R-Strategies · Regulatory & Compliance (DPP)',
                'LFL relevance: LFL\'s "Carbon Case" (10.5M tCO₂e savings by year 10) and real-time CO₂ feedback position LFL as a bridge between Engineering and Sustainability',
                'Regulatory driver: CSRD obligation from 2025 for ~50,000 EU companies creates immense pressure to act',
            ]
        },
    ]

    for seg in segments:
        add_h2(doc, seg['title'], color=seg['color'])
        add_para(doc, seg['body'], sa=3)
        for b in seg['bullets']:
            add_bullet(doc, b)
        thin_gap(doc, 4)
    doc.add_page_break()

    # ══ 3  DEEP DIVE ══════════════════════════
    add_h1(doc, '3  Deep Dive: Sales & Quotes')
    add_para(doc, (
        'The following in-depth analysis of the Sales & Quotes category combines structured market data '
        'from CI database v1.7 with qualitative insights from the LoopForgeLab White Paper '
        '("Optimized for Performance, Approved for Price", Feb. 2026) and the Pitch Deck (March 2026). '
        'The objective is to quantify customer problem fields and show where and how '
        'LoopForgeLab\'s solution enters the market.'
    ), sa=5)

    # 3.1 Sub-clusters
    add_h2(doc, '3.1  Market Structure: The Four Sub-Clusters')
    add_para(doc, 'The Sales & Quotes category is divided into four functional sub-clusters with distinct customers, technologies and problems:', sa=3)

    cl_data = [
        ('A','RfP / Bid Response Management (AI-assisted)','8',
         'Generative AI (GPT-4/Claude), RAG, AI Agents, CRM Integration',
         'B2B Technology, Engineering Services, SaaS, Defence, Consulting'),
        ('B','E-Sourcing / RfQ Management (Procurement)','8',
         'Autonomous AI Sourcing Bots, Reverse Auctions, ERP Integration (SAP/Oracle), Spend Analytics',
         'Enterprise/Mid-Market Procurement: Manufacturing, Pharma, Automotive, Public Sector'),
        ('C','CPQ – Configure Price Quote','3',
         'Constraint-based AI Configuration, CAD Integration, ERP/CRM Coupling, CO₂ Calculation',
         'ETO Manufacturers: Industrial Machinery, Automotive Suppliers, High-Tech'),
        ('D','Supply Chain Planning & Optimisation','14',
         'AI/ML Demand Forecasting, Digital Twins, Concurrent Planning, Graph Neural Networks, Maths Optimisation',
         'Automotive, Consumer Goods, Pharma, Aerospace, Retail, High-Tech'),
    ]
    cl_hdrs = ['ID','Sub-Cluster','# Co.','Key Technologies','Primary Target Sectors']
    cl_w    = [Cm(0.7),Cm(4.3),Cm(0.9),Cm(5.6),Cm(5.8)]
    cltbl = doc.add_table(rows=1+len(cl_data), cols=5)
    cltbl.style = 'Table Grid'
    cltbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    for j,(h,w) in enumerate(zip(cl_hdrs,cl_w)):
        c=cltbl.cell(0,j); c.width=w; hdr_cell(c,h,bg='1A5F8A')
    for i,row in enumerate(cl_data):
        bg='FFFFFF' if i%2==0 else 'F2F5F8'
        for j,val in enumerate(row):
            data_cell(cltbl.cell(i+1,j), val, bg=bg,
                      bold=(j==0), color=(MID_BLUE if j==0 else None),
                      center=(j in (0,2)), fs=8)
    thin_gap(doc, 5)

    # 3.2 Customer problems
    add_h2(doc, '3.2  Customer Problem Fields — and the LFL Context')
    add_para(doc, (
        'LoopForgeLab\'s White Paper documents four structural barriers to profitable and sustainable '
        'product development, based on 44 expert interviews. These barriers are directly mirrored in '
        'the measured customer problems of the Sales & Quotes category:'
    ), sa=4)

    framing = [
        ('The "Economic Gatekeeper"',
         '"The price tag decides everything" — the near-universal sentiment from 44 LFL expert interviews. '
         'This is directly reflected in the CPQ and RfQ world: engineers spend 30–40% of their time on '
         'quoting rather than design, because configuration and pricing are manual, error-prone and '
         'completely decoupled from the CAD model.'),
        ('The "Data Desert"',
         'Critical product data — CO₂ footprints, supplier prices, manufacturing constraints — are scattered '
         'across PLM, ERP, Excel and local documents. The result: RfP teams rewrite 30–40% of all responses '
         'from scratch even though the information exists somewhere in the organisation. '
         'The Loopio benchmark measures 17+ hours per RfP response.'),
        ('The "Redesign Trap"',
         'Because quoting and CAD design are fully decoupled, 50–70% of all engineering changes originate '
         'from inconsistency between quote and CAD (DriveWorks studies). Every late change multiplies cost — '
         'typically 5–15% of order value in error correction after contract award.'),
        ('The "Tribal Knowledge" Erosion',
         'When experienced engineers retire, the implicit knowledge of why certain configurations work — '
         'or why certain suppliers are preferred — disappears. CPQ systems can only scale this knowledge '
         'if it is explicitly codified. Today, that codification largely does not exist.'),
    ]

    for title, body in framing:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(5)
        p.paragraph_format.space_after  = Pt(1)
        r = p.add_run(f'■  {title}')
        r.bold = True; r.font.color.rgb = DARK_BLUE; r.font.size = Pt(10)
        add_para(doc, body, sa=2)
    thin_gap(doc, 4)

    # Problem table
    add_h3(doc, 'Customer Problem Fields by Sub-Cluster')
    prob_data = [
        ('A – RfP/Bid Response','Time burden',      '17+ h/RfP response (Loopio benchmark)',                '68 h/month = ~1.7 FTE solely for RfP work',                 '★★★★★'),
        ('A – RfP/Bid Response','Knowledge silos',  '30–40% of answers rewritten from scratch',             'Massive duplication; no institutional memory',               '★★★★★'),
        ('A – RfP/Bid Response','Bid/No-Bid',       'Gut-feeling decision, no win-probability score',        '80–85% of effort spent on losing bids',                     '★★★★☆'),
        ('A – RfP/Bid Response','Win-theme gap',    'No access to CRM/call data for positioning',           'Generic proposals → low win rate',                          '★★★★☆'),
        ('B – E-Sourcing/RfQ',  'Excel/email chaos','RfX processes via Excel without versioning',            '3–5 day lead time; frequent errors from version confusion',  '★★★★★'),
        ('B – E-Sourcing/RfQ',  'CSRD/Scope-3',     '~50,000 EU companies must report Scope-3 emissions',   'Manual supplier surveys for hundreds of Tier-1 suppliers',   '★★★★★'),
        ('B – E-Sourcing/RfQ',  'SME access barrier','Enterprise suites cost €100K+/year',                  '99.8% of European manufacturers are SMEs without digital sourcing','★★★★★'),
        ('C – CPQ',             'Engineer bottleneck','Every quote for custom products requires engineer',   'ETO lead time: 2–4 weeks; engineers spend 30–40% on quoting','★★★★★'),
        ('C – CPQ',             'No CAD-quote link', 'Quote and CAD fully decoupled',                       '50–70% engineering changes from quote/CAD inconsistency',    '★★★★☆'),
        ('C – CPQ',             'ETO without solution','No standard CPQ for true ETO products',             '~60% of EU machinery orders are custom-engineered',          '★★★★★'),
        ('D – SC Planning',     'Forecast errors',   'Excel + legacy APS, no real-time updates',             'Avg. 30–40% MAPE forecast error; excessive inventory costs', '★★★★★'),
        ('D – SC Planning',     'Multi-tier blindness','Only Tier-1 suppliers visible',                     '77% of disruptions originate at Tier-2+ suppliers',          '★★★★☆'),
    ]
    pr_hdrs = ['Sub-Cluster','Problem','Description','Measurable Impact','Priority']
    pr_w    = [Cm(3.0),Cm(2.6),Cm(5.4),Cm(5.2),Cm(1.5)]
    prtbl = doc.add_table(rows=1+len(prob_data), cols=5)
    prtbl.style = 'Table Grid'
    prtbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    for j,(h,w) in enumerate(zip(pr_hdrs,pr_w)):
        c=prtbl.cell(0,j); c.width=w; hdr_cell(c,h)
    for i,row in enumerate(prob_data):
        bg='FFFFFF' if i%2==0 else 'F2F5F8'
        for j,val in enumerate(row):
            data_cell(prtbl.cell(i+1,j), val, bg=bg,
                      bold=(j==0), color=(MID_BLUE if j==0 else None),
                      center=(j==4), fs=8)
    thin_gap(doc, 5)

    # 3.3 LFL strategic fit
    add_h2(doc, '3.3  Relevance for LoopForgeLab — the Strategic Fit')
    add_para(doc, (
        'LoopForgeLab positions its Product Intelligence Engine as the economic integration layer '
        'across the engineering ecosystem. The Pitch Deck identifies four direct competitive fields: '
        'Engineering, Data & Requirements Management, Sustainability and Sales & Quotes. '
        'From the measured customer problems, a clear strategic fit emerges:'
    ), sa=4)

    fit_data = [
        ('RFQ Module (from Q2 2026)', 'Sub-cluster B & C',
         'LFL\'s first core module automates quote calculation for mechanical products. '
         'It resolves the "engineer bottleneck" (30–40% of engineering time on quoting) and closes '
         'the gap between CAD model and quote configuration — the core problem of CPQ for ETO.'),
        ('Manufacturing Cost Intelligence', 'Sub-cluster B, C',
         'Real-time manufacturing costs derived directly from engineering data — without Excel. '
         'Addresses the "Economic Gatekeeper" finding: cost-based decisions can be made earlier '
         'in the design process rather than at the final stage.'),
        ('Compliance & Risk', 'Sub-cluster B, D',
         'Integration of CSRD/Scope-3 requirements into the design and procurement process. '
         'Addresses the growing regulatory pressure that applies to ~50,000 EU companies from 2025.'),
        ('Life Cycle Cost & Operator Insights', 'Sub-cluster C, D',
         'Total Cost of Ownership, repairability and circular economy potential as native design '
         'parameters — making sustainability an innovation driver rather than a compliance burden.'),
    ]
    for title, cluster, body in fit_data:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(5)
        p.paragraph_format.space_after  = Pt(1)
        r1 = p.add_run(f'▶  {title}')
        r1.bold=True; r1.font.color.rgb=DARK_BLUE; r1.font.size=Pt(10)
        r2 = p.add_run(f'   [{cluster}]')
        r2.font.color.rgb=ACCENT; r2.font.size=Pt(9)
        add_para(doc, body, sa=2)
    thin_gap(doc, 3)

    add_para(doc,
        'Market size (Pitch Deck):  TAM € 16.6B · SAM € 5.6B · SOM € 98M  '
        '(415,000 mechanical engineering companies, engineering software CAGR 15%)',
        bold=True, color=DARK_BLUE, sa=2)
    add_para(doc,
        'Primary target markets: Packaging Machinery producers · Material Handling Machinery · '
        'Engineering Services — OEM in DE/AT/CH/PL/CZ/NL/BE with ≥ 30 employees', sa=4)
    doc.add_page_break()

    # ══ 4  SOLUTION PROVIDERS ═════════════════
    add_h1(doc, '4  Solution Providers & Technologies in Sales & Quotes')

    # 4.1 Provider profiles
    add_h2(doc, '4.1  Provider Profiles by Sub-Cluster')

    # Cluster A
    add_h3(doc, 'Sub-Cluster A — RfP / Bid Response Management')
    add_para(doc, (
        'This sub-cluster focuses on AI-assisted automation of responses to tenders (RfP, RfI, DDQ). '
        'The market exploded in 2023–2025, disrupted by a wave of AI-native start-ups challenging '
        'legacy providers such as Loopio.'
    ), sa=3)
    a_rows = [
        ('Loopio',        'Large',  'RfP market leader, 1,500+ customers. AI trained on 10+ years of proposal data. Strengths: UX, support (9.7/10 G2). Weakness: library-based AI, no live connections, slower than newer AI-native tools.'),
        ('Responsive.io', 'Large',  '2,000+ customers, 100+ integrations (Salesforce, Slack, MS Teams), 50+ languages. Strengths: broad integration ecosystem, analytics. Weakness: steep learning curve, complex pricing.'),
        ('Arphie',        'Startup','AI-native with live KB connections (Google Drive, SharePoint, Confluence). 80%+ time saving vs. Loopio benchmark, 84% AI acceptance rate. Flat-rate pricing (unlimited users). Zero data retention as trust feature.'),
        ('Altura',        'Startup','Autonomous "Bid Companion" AI agent that independently completes RfP tasks and proactively flags compliance risks. European focus, real-time bid/no-bid intelligence.'),
        ('AutogenAI',     'Startup','$36.1M revenue (2025), 200+ enterprise customers. Generates complete bid documents for complex government and engineering RfPs. Best for narrative-heavy bids; less suited for structured RfQ forms.'),
        ('Inventive AI',  'Startup','Identifies "Win Themes" from CRM + Slack + sales calls for strategically superior RfP responses. Deep Salesforce/CRM integration. $1.7M ARR bootstrapped.'),
        ('DeepRFP',       'Startup','AI agent platform with 28-language support and dedicated compliance automation agents. Most comprehensive multilingual approach in the cluster.'),
        ('AutoRFP.ai',    'Startup','Browser-based RfP automation via Chrome Extension — works directly inside SAP Ariba and government portals. Unique workflow for procurement-driven organisations.'),
    ]
    provider_table(doc, a_rows)
    thin_gap(doc, 5)

    # Cluster B
    add_h3(doc, 'Sub-Cluster B — E-Sourcing / RfQ Management (Procurement)')
    add_para(doc, (
        'E-sourcing platforms digitise the procurement process. The market is divided: enterprise suites '
        'for large corporations (Coupa, Ivalua, Jaggaer) and affordable alternatives for mid-market companies. '
        'The most pressing issue remains the unaffordability of large suites for SMEs.'
    ), sa=3)
    b_rows = [
        ('Coupa',       'Large',  '$619.4M revenue (FY2025), 3,500 employees. Community intelligence from $6T+ transactions for benchmarking and risk detection. End-to-end BSM. Weakness: premium pricing, unsuitable for SMEs.'),
        ('Ivalua',      'Large',  'Global S2P leader on a unified platform. Strong EU presence, CSRD reporting integration. High implementation costs.'),
        ('Jaggaer',     'Large',  'Autonomous procurement agents (JAI) for automated source-to-pay. Strong manufacturing industry focus. Enterprise-only.'),
        ('Keelvar',     'Startup','Only vendor in Gartner Market Guide 2025 for BOTH Advanced Sourcing Optimization AND Autonomous Sourcing. $90B+ spend managed. Autonomous AI bots run complete RfQ cycles without manual intervention.'),
        ('ProcurePort', 'Startup','Affordable full-suite eProcurement since 2011 (bootstrapped). Handles 1,000–15,000-item RfX events. Significantly cheaper than Coupa/Ivalua/Jaggaer — ideal SME entry point.'),
        ('QLM Sourcing', 'Startup','Custom eRfQ template engine: per-product-category customisable RfQ forms for engineering procurement. Strong supplier collaboration features.'),
        ('Bonfire',     'Startup','600+ public sector procurement teams, $18.2M revenue. Digital scorecards, weighted evaluation, what-if analysis. Weakness: primarily public sector.'),
        ('Vendorful',   'Startup','API-first e-sourcing with deep ERP integration, replaces Excel/SharePoint RfX workflows. Suitable for beginners and experts alike.'),
    ]
    provider_table(doc, b_rows)
    thin_gap(doc, 5)

    # Cluster C
    add_h3(doc, 'Sub-Cluster C — CPQ: Configure Price Quote')
    add_para(doc, (
        'CPQ is the smallest sub-cluster (3 companies) but holds the strongest strategic fit to LoopForgeLab. '
        'The core task: configure complex, customer-specific products (Engineer-to-Order) so that a correct '
        'quote is generated in minutes rather than weeks — including automatic CAD generation. '
        'Around 60% of all orders in European mechanical engineering are custom; a scalable standard solution is missing.'
    ), sa=3)
    c_rows = [
        ('Tacton',    'Large',  'Stockholm, 301 employees. Market leader CPQ for complex manufactured products. Only CPQ solution with integrated CO₂ footprint calculation and EPD generation directly in product configuration. Strong CAD and ERP/CRM coupling.'),
        ('DriveWorks','Startup','Sheffield, UK. Design automation + CPQ native to SolidWorks. Rule-based CAD generation directly from the quote — unique workflow: quote → CAD model automatically. Dominant in the SolidWorks ecosystem.'),
        ('Elfsquad',  'Startup','Groningen, NL. CPQ/configuration software specifically for manufacturing and industrial companies. Focus on visual product configuration and seamless ERP integration for DACH/Benelux mid-market.'),
    ]
    provider_table(doc, c_rows)
    thin_gap(doc, 5)

    # Cluster D
    add_h3(doc, 'Sub-Cluster D — Supply Chain Planning & Optimisation')
    add_para(doc, (
        'With 14 companies, this is the largest sub-cluster. Supply chain planning solutions address '
        'forecast errors, real-time responsiveness and multi-tier supplier visibility. The market is '
        'dominated by a few large platform providers, complemented by specialised AI start-ups for '
        'risk management and optimisation.'
    ), sa=3)
    d_rows = [
        ('Kinaxis',      'Large',  'Ottawa, CA. Concurrent planning platform "Maestro" — the only AI-infused end-to-end SC orchestration platform. Real-time plan adjustment instead of sequential cycles. Dominant in automotive and high-tech.'),
        ('Blue Yonder',  'Large',  'Scottsdale, USA. AI/ML-powered supply chain planning & execution. Broad industry coverage. Strong demand-sensing capabilities.'),
        ('o9 Solutions', 'Startup','Dallas, TX. Proprietary Enterprise Knowledge Graph (EKG) for end-to-end SC modelling. Neurosymbolic AI agents + GenAI/LLM composite agents for cross-functional planning.'),
        ('Resilinc',     'Startup','Milpitas, CA. Agentic AI trained on 15+ years of SC data. Autonomous risk detection and mitigation. Specialist in multi-tier visibility and disruption response.'),
        ('Scoutbee',     'Startup','Würzburg, DE. AI-driven supplier discovery and risk management using LLMs. European player, strong ties to German manufacturing industry.'),
        ('Cosmo Tech',   'Startup','Lyon, FR. Simulation-based supply chain optimisation with digital twins. Strength: complex scenario analysis for non-linear SC problems.'),
    ]
    provider_table(doc, d_rows)
    thin_gap(doc, 5)
    doc.add_page_break()

    # 4.2 Technology radar
    add_h2(doc, '4.2  Technology Radar: Key Technologies in Sales & Quotes')
    add_para(doc, 'The technology radar below shows the most important technologies in the Sales & Quotes market, their maturity and relevance to LoopForgeLab:', sa=3)

    tech_data = [
        ('Generative AI (LLM) for Proposal Text', '★★★★★ Production', 'A',
         'Arphie, AutogenAI, Inventive AI, Loopio, DeepRFP, Altura',
         'Decisive — quality, speed, personalisation',
         'Directly relevant for LFL\'s RFQ text generation'),
        ('RAG (Retrieval-Augmented Generation)', '★★★★★ Production', 'A+B',
         'Arphie (live KB), Inventive AI, Responsive.io',
         'Foundation for all AI-native tools — knowledge base access',
         'Core technology for LFL\'s Data Desert solution'),
        ('Autonomous AI Agents', '★★★★☆ Rising', 'A+B',
         'Keelvar (Sourcing Bots), Altura (Bid Companion), Arphie',
         'Next evolution: from assist to autonomous',
         'Roadmap-relevant for LFL\'s Design Copilot'),
        ('Constraint-based Configuration Engine', '★★★★★ Production', 'C',
         'Tacton, DriveWorks, Elfsquad',
         'Core IP for CPQ in ETO — differentiation through complexity depth',
         'Direct competition for LFL\'s RFQ module for ETO'),
        ('CAD-CPQ Integration', '★★★★☆ Production', 'C',
         'DriveWorks (SolidWorks-native)',
         'Unique workflow: quote → CAD model automatically',
         'Core feature of LFL\'s differentiation approach'),
        ('ML for Instant Quoting', '★★★★★ Production', 'B (Mfg)',
         'Xometry, Protolabs',
         'ML trained on millions of parts — price in <60 sec',
         'Long-term goal for LFL in manufacturing network'),
        ('Concurrent Planning / Digital Twin SC', '★★★★☆ Rising', 'D',
         'Kinaxis (RapidResponse), Cosmo Tech, o9 Solutions',
         'Real-time plan adjustment without sequential cycles',
         'Indirectly relevant for LFL\'s supply chain data layer'),
        ('ESG/CSRD Data in Sourcing Workflow', '★★★☆☆ Growing', 'B+C',
         'Tacton (CO₂ Configurator), Jaggaer, Ivalua',
         'Mandatory feature for EU compliance from 2025',
         'Core element of LFL Sustainability positioning'),
        ('Bid/No-Bid AI Scoring', '★★★☆☆ Rising', 'A',
         'Altura, Inventive AI, AutogenAI',
         'Optimise resource allocation: prioritise the right bids',
         'Relevant for LFL\'s Technical Sales target group'),
    ]
    th_hdrs = ['Technology','Maturity','Cluster','Example Providers','Competitive Differentiation','LFL Relevance']
    th_w    = [Cm(3.4),Cm(2.0),Cm(1.0),Cm(3.8),Cm(3.9),Cm(3.2)]
    ttbl = doc.add_table(rows=1+len(tech_data), cols=6)
    ttbl.style = 'Table Grid'
    ttbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    for j,(h,w) in enumerate(zip(th_hdrs,th_w)):
        c=ttbl.cell(0,j); c.width=w; hdr_cell(c,h)
    for i,row in enumerate(tech_data):
        bg='FFFFFF' if i%2==0 else 'F2F5F8'
        for j,val in enumerate(row):
            clr = MID_BLUE if j==5 else None
            data_cell(ttbl.cell(i+1,j), val, bg=bg, color=clr,
                      fs=8, center=(j==2))
            if j==5:
                for para in ttbl.cell(i+1,j).paragraphs:
                    for run in para.runs:
                        run.font.italic=True
    thin_gap(doc, 5)
    doc.add_page_break()

    # ══ 5  POSITIONING & WHITESPACE ═══════════
    add_h1(doc, '5  Strategic Positioning & Whitespace for LoopForgeLab')
    add_para(doc, (
        'The analysis of 35 companies and 39 solutions in the Sales & Quotes category reveals a clear '
        'picture of market gaps that LoopForgeLab can target:'
    ), sa=5)

    ws_items = [
        ('Engineering-native RFQ process is missing',
         'All existing CPQ and e-sourcing solutions start after the design process. No solution anchors '
         'quote calculation natively in the CAD-based engineering workflow. Tacton and DriveWorks come '
         'closest, but are not positioned as "Engineering-First" tools. LFL closes this gap with its '
         'Product Intelligence Engine — cost, compliance and CO₂ become native design parameters, '
         'not downstream calculations.'),
        ('SME-ready ETO-CPQ is missing',
         'The three CPQ providers (Tacton, DriveWorks, Elfsquad) are either enterprise-only or '
         'SolidWorks-exclusive. For the European mid-market (mechanical engineering, material handling), '
         'there is no affordable, CAD-agnostic CPQ solution for true ETO products. LFL\'s target segment '
         '(OEM >30 employees in DACH/Benelux) is precisely this gap.'),
        ('CO₂ + cost in a single tool',
         'Only Tacton has begun integrating a CO₂ calculator into CPQ. The combination of real-time '
         'cost calculation, CO₂ feedback and quote generation in an engineer-centric interface does not '
         'yet exist. LFL\'s "Carbon Case" (up to 10.5M tCO₂e savings in year 10 at 14% market share) '
         'positions this integration as a strategic differentiator.'),
        ('Tribal knowledge as competitive advantage',
         'No provider in the Sales & Quotes category explicitly addresses the codification of '
         'institutional engineering knowledge. LFL\'s "Operator Insights" module (repairability, '
         'circular economy) and Design Copilot approach create scalable digital knowledge infrastructure '
         '— a genuine whitespace.'),
    ]
    for i,(title,body) in enumerate(ws_items, 1):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(6)
        p.paragraph_format.space_after  = Pt(1)
        r = p.add_run(f'{i}.  {title}')
        r.bold=True; r.font.color.rgb=DARK_BLUE; r.font.size=Pt(11)
        add_para(doc, body, sa=3)
    thin_gap(doc, 6)

    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    r = p.add_run(
        '⟶  LoopForgeLab is the only provider occupying the whitespace of "economic integration" — '
        'connecting engineering physics with economic logic at the moment of maximum influence: '
        'the early design phase. This is not a feature competitors can copy overnight — it requires '
        'deep domain expertise in mechanical engineering, product economics and AI engineering simultaneously.'
    )
    r.bold=True; r.font.color.rgb=DARK_BLUE; r.font.size=Pt(11)
    thin_gap(doc, 6)

    add_para(doc, (
        'Sources: All market data from Competitive Intelligence Database v1.7 (13 March 2026). '
        'Qualitative insights from LoopForgeLab White Paper (February 2026, 44 expert interviews) '
        'and LoopForgeLab Pitch Deck (March 2026). External benchmarks: Loopio RfP Benchmark Report; '
        'Resilinc Supply Chain Disruption Study 2023; DriveWorks Engineering Change Studies; '
        'McKinsey Supply Chain Disruption Cost Analysis 2023; Center for Automotive Research 2026; '
        'EU Commission CSRD Regulation.'
    ), size=7.5, color=RGBColor(0x80,0x80,0x90), sa=2)

    out = '/home/user/LFL_Competitor_Report/260313_LFL_Competitor_Report_EN.docx'
    doc.save(out)
    print(f'✅  Report saved: {out}')

if __name__ == '__main__':
    build()
