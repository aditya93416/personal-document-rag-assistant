import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

def build_presentation():
    prs = Presentation()
    # 16:9 widescreen layout: 13.333 x 7.5 inches
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    # Professional academic & corporate color scheme
    NAVY_DARK = RGBColor(16, 44, 87)       # #102C57
    NAVY_LIGHT = RGBColor(53, 89, 142)     # #35598E
    SLATE_TEXT = RGBColor(30, 41, 59)      # #1E293B
    MUTED_TEXT = RGBColor(100, 116, 139)   # #64748B
    BG_LIGHT = RGBColor(248, 250, 252)     # #F8FAFC
    CARD_BG = RGBColor(255, 255, 255)      # White
    ACCENT_TEAL = RGBColor(14, 165, 233)   # #0EA5E9
    BORDER_COLOR = RGBColor(203, 213, 225) # #CBD5E1

    def add_slide_header(slide, title_text, category_text="B.TECH. INTERNSHIP PRESENTATION (SESSION 2025–2026)"):
        # Header banner shape
        hdr_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.733), Inches(1.1))
        tf = hdr_box.text_frame
        tf.word_wrap = True
        tf.margin_top = tf.margin_bottom = tf.margin_left = tf.margin_right = 0
        
        p_cat = tf.paragraphs[0]
        p_cat.text = category_text.upper()
        p_cat.font.name = "Calibri"
        p_cat.font.size = Pt(11)
        p_cat.font.bold = True
        p_cat.font.color.rgb = ACCENT_TEAL
        p_cat.space_after = Pt(2)
        
        p_title = tf.add_paragraph()
        p_title.text = title_text
        p_title.font.name = "Calibri"
        p_title.font.size = Pt(24)
        p_title.font.bold = True
        p_title.font.color.rgb = NAVY_DARK

        # Top decorative rule line
        line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.5), Inches(11.733), Inches(0.04))
        line.fill.solid()
        line.fill.fore_color.rgb = NAVY_LIGHT
        line.line.color.rgb = NAVY_LIGHT

    def add_slide_footer(slide):
        ftr_box = slide.shapes.add_textbox(Inches(0.8), Inches(6.9), Inches(11.733), Inches(0.4))
        tf = ftr_box.text_frame
        tf.word_wrap = True
        tf.margin_top = tf.margin_bottom = tf.margin_left = tf.margin_right = 0
        p = tf.paragraphs[0]
        p.text = "Haridwar University, Roorkee | Aditya Kumar (024101010012) | InternCourse (Patna) | IC26/11873"
        p.font.name = "Calibri"
        p.font.size = Pt(9.5)
        p.font.color.rgb = MUTED_TEXT

    def add_card(slide, left, top, width, height, bg_color=CARD_BG, border_color=BORDER_COLOR):
        shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
        shape.fill.solid()
        shape.fill.fore_color.rgb = bg_color
        if border_color:
            shape.line.color.rgb = border_color
            shape.line.width = Pt(1)
        else:
            shape.line.fill.background()
        return shape

    # =========================================================================
    # SLIDE 1: TITLE SLIDE
    # =========================================================================
    s1 = prs.slides.add_slide(blank_layout)
    bg1 = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
    bg1.fill.solid()
    bg1.fill.fore_color.rgb = NAVY_DARK
    bg1.line.fill.background()

    # Inner decorative container
    inner = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(0.8), Inches(11.733), Inches(5.9))
    inner.fill.solid()
    inner.fill.fore_color.rgb = RGBColor(24, 52, 98)
    inner.line.color.rgb = ACCENT_TEAL
    inner.line.width = Pt(1.5)

    tb1 = s1.shapes.add_textbox(Inches(1.2), Inches(1.1), Inches(10.933), Inches(3.2))
    tf1 = tb1.text_frame
    tf1.word_wrap = True
    
    p = tf1.paragraphs[0]
    p.text = "HARIDWAR UNIVERSITY, ROORKEE"
    p.font.name = "Calibri"
    p.font.size = Pt(14)
    p.font.bold = True
    p.font.color.rgb = ACCENT_TEAL
    p.space_after = Pt(2)

    p = tf1.add_paragraph()
    p.text = "Department of Computer Science & Engineering | B.Tech. 5th Semester"
    p.font.name = "Calibri"
    p.font.size = Pt(12)
    p.font.color.rgb = RGBColor(203, 213, 225)
    p.space_after = Pt(18)

    p = tf1.add_paragraph()
    p.text = "Basic Analysis of RAG Systems & Development of a\nPersonal Document Q&A Assistant"
    p.font.name = "Calibri"
    p.font.size = Pt(28)
    p.font.bold = True
    p.font.color.rgb = RGBColor(255, 255, 255)
    p.space_after = Pt(12)

    p = tf1.add_paragraph()
    p.text = "Summer Internship Project Presentation (15 June 2026 – 18 July 2026)"
    p.font.name = "Calibri"
    p.font.size = Pt(13)
    p.font.italic = True
    p.font.color.rgb = RGBColor(148, 163, 184)

    # 2 Column Meta Box
    # Left: Student Details
    c_left = add_card(s1, Inches(1.2), Inches(4.5), Inches(5.2), Inches(1.8), bg_color=NAVY_DARK, border_color=NAVY_LIGHT)
    tb_sl = s1.shapes.add_textbox(Inches(1.4), Inches(4.6), Inches(4.8), Inches(1.6))
    tfl = tb_sl.text_frame
    tfl.word_wrap = True
    p = tfl.paragraphs[0]
    p.text = "CANDIDATE PROFILE:"
    p.font.bold = True
    p.font.size = Pt(11)
    p.font.color.rgb = ACCENT_TEAL
    p = tfl.add_paragraph()
    p.text = "Name: Aditya Kumar\nRoll No / Enrolment No: 024101010012\nFaculty Mentor: Pooja Parmar | HOD: Dr. Himanshu Verma"
    p.font.size = Pt(10)
    p.font.color.rgb = RGBColor(241, 245, 249)

    # Right: Organisation Details
    c_right = add_card(s1, Inches(6.9), Inches(4.5), Inches(5.2), Inches(1.8), bg_color=NAVY_DARK, border_color=NAVY_LIGHT)
    tb_sr = s1.shapes.add_textbox(Inches(7.1), Inches(4.6), Inches(4.8), Inches(1.6))
    tfr = tb_sr.text_frame
    tfr.word_wrap = True
    p = tfr.paragraphs[0]
    p.text = "ORGANISATION PROFILE:"
    p.font.bold = True
    p.font.size = Pt(11)
    p.font.color.rgb = ACCENT_TEAL
    p = tfr.add_paragraph()
    p.text = "InternCourse (interncourse.online), Patna, Bihar\nManaging Director: Ankit Pandey\nAccreditation: ISO 9001:2015, AICTE, MCA | Cert ID: IC26/11873"
    p.font.size = Pt(10)
    p.font.color.rgb = RGBColor(241, 245, 249)

    # =========================================================================
    # SLIDE 2: PRESENTATION OUTLINE / AGENDA
    # =========================================================================
    s2 = prs.slides.add_slide(blank_layout)
    add_slide_header(s2, "Executive Presentation Outline")
    add_slide_footer(s2)

    agenda_items = [
        ("01", "Introduction & Problem Statement", "Parametric LLM limitations, hallucinations, and privacy vulnerabilities in enterprise document search."),
        ("02", "Host Organisation Overview", "InternCourse (Patna), ISO 9001:2015 quality frameworks, AICTE training divisions, and mentorship."),
        ("03", "Internship Objectives & Scope", "Vector algebra, dense semantic representations, chunking heuristics, and grounding architecture."),
        ("04", "RAG Pipeline Architecture", "7-stage workflow: PyPDF parsing -> 500-token chunking -> MiniLM embeddings -> ChromaDB -> Grounded LLM."),
        ("05", "Experimental Benchmarks", "Empirical evaluations of chunk size vs. retrieval accuracy, sub-15ms vector lookup latencies."),
        ("06", "Challenges, Solutions & Demo", "Boundary context loss remedies, token optimization, UI dashboard, and live GitHub repository.")
    ]

    for idx, (num, title, desc) in enumerate(agenda_items):
        col = idx % 2
        row = idx // 2
        l = Inches(0.8 + col * 5.95)
        t = Inches(1.8 + row * 1.6)
        
        add_card(s2, l, t, Inches(5.75), Inches(1.4), bg_color=BG_LIGHT)
        
        # Number badge
        nb = s2.shapes.add_shape(MSO_SHAPE.RECTANGLE, l + Inches(0.2), t + Inches(0.2), Inches(0.6), Inches(0.6))
        nb.fill.solid()
        nb.fill.fore_color.rgb = NAVY_DARK
        nb.line.fill.background()
        tf_nb = nb.text_frame
        p = tf_nb.paragraphs[0]
        p.text = num
        p.font.bold = True
        p.font.size = Pt(14)
        p.font.color.rgb = RGBColor(255, 255, 255)
        p.alignment = PP_ALIGN.CENTER
        
        # Text
        tb = s2.shapes.add_textbox(l + Inches(0.95), t + Inches(0.15), Inches(4.6), Inches(1.1))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = title
        p.font.bold = True
        p.font.size = Pt(13)
        p.font.color.rgb = NAVY_DARK
        p.space_after = Pt(2)
        p2 = tf.add_paragraph()
        p2.text = desc
        p2.font.size = Pt(10)
        p2.font.color.rgb = MUTED_TEXT

    # =========================================================================
    # SLIDE 3: PROBLEM STATEMENT & MOTIVATION
    # =========================================================================
    s3 = prs.slides.add_slide(blank_layout)
    add_slide_header(s3, "Problem Statement: Limitations of Foundational LLMs")
    add_slide_footer(s3)

    # 3 Comparison Columns
    cards_s3 = [
        ("Parametric Knowledge Cutoffs", 
         "• LLM neural weights are frozen post-pretraining.\n• Foundational models have zero awareness of new world events or private domain datasets.\n• Fine-tuning is computationally expensive and quickly obsolescent."),
        ("Hallucination Vulnerabilities", 
         "• When queried about unobserved domain documents, LLMs generate factually inaccurate yet fluent answers.\n• Intolerable risk in academic, engineering, and legal workflows requiring absolute factual fidelity."),
        ("Data Privacy & Enterprise Security", 
         "• Transmitting proprietary university records or enterprise intellectual property to commercial third-party cloud APIs violates compliance.\n• Mandates fully offline, local, privacy-preserving retrieval architectures.")
    ]

    for idx, (title, content) in enumerate(cards_s3):
        l = Inches(0.8 + idx * 3.95)
        add_card(s3, l, Inches(1.8), Inches(3.8), Inches(4.7), bg_color=BG_LIGHT)
        
        # Top banner
        tb_hdr = s3.shapes.add_shape(MSO_SHAPE.RECTANGLE, l, Inches(1.8), Inches(3.8), Inches(0.6))
        tb_hdr.fill.solid()
        tb_hdr.fill.fore_color.rgb = NAVY_DARK if idx != 1 else NAVY_LIGHT
        tb_hdr.line.fill.background()
        p = tb_hdr.text_frame.paragraphs[0]
        p.text = title
        p.font.bold = True
        p.font.size = Pt(12)
        p.font.color.rgb = RGBColor(255, 255, 255)
        p.alignment = PP_ALIGN.CENTER
        
        tb = s3.shapes.add_textbox(l + Inches(0.2), Inches(2.6), Inches(3.4), Inches(3.7))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = content
        p.font.size = Pt(11)
        p.font.color.rgb = SLATE_TEXT

    # =========================================================================
    # SLIDE 4: ABOUT THE ORGANISATION (INTERNCOURSE)
    # =========================================================================
    s4 = prs.slides.add_slide(blank_layout)
    add_slide_header(s4, "Host Organisation: InternCourse (interncourse.online)")
    add_slide_footer(s4)

    # Left: Overview Card
    add_card(s4, Inches(0.8), Inches(1.8), Inches(6.8), Inches(4.7), bg_color=BG_LIGHT)
    tb_org = s4.shapes.add_textbox(Inches(1.1), Inches(2.0), Inches(6.2), Inches(4.3))
    tf = tb_org.text_frame
    tf.word_wrap = True
    
    p = tf.paragraphs[0]
    p.text = "Corporate Profile & Mission:"
    p.font.bold = True
    p.font.size = Pt(14)
    p.font.color.rgb = NAVY_DARK
    p.space_after = Pt(6)

    p = tf.add_paragraph()
    p.text = (
        "InternCourse is an accredited industrial training provider registered under the Ministry of Corporate Affairs (MCA), "
        "headquartered in Patna, Bihar. The organization bridges academia and the AI industry through intensive, production-focused engineering internships."
    )
    p.font.size = Pt(11)
    p.font.color.rgb = SLATE_TEXT
    p.space_after = Pt(10)

    p = tf.add_paragraph()
    p.text = "Key Institutional Accreditations:"
    p.font.bold = True
    p.font.size = Pt(12)
    p.font.color.rgb = NAVY_DARK
    p.space_after = Pt(4)

    bullets = [
        "ISO 9001:2015 Certified for Quality Management in technical education.",
        "Partnered with AICTE (All India Council for Technical Education) initiatives.",
        "Ministry of Corporate Affairs (MCA) registered and compliant corporate entity.",
        "Executive Leadership: Managing Director Ankit Pandey."
    ]
    for b in bullets:
        p = tf.add_paragraph()
        p.text = "✓  " + b
        p.font.size = Pt(10.5)
        p.font.color.rgb = SLATE_TEXT

    # Right: Internship Verification Details Card
    add_card(s4, Inches(7.9), Inches(1.8), Inches(4.6), Inches(4.7), bg_color=NAVY_DARK, border_color=ACCENT_TEAL)
    tb_v = s4.shapes.add_textbox(Inches(8.2), Inches(2.0), Inches(4.0), Inches(4.3))
    tfv = tb_v.text_frame
    tfv.word_wrap = True
    
    p = tfv.paragraphs[0]
    p.text = "OFFICIAL CREDENTIALS"
    p.font.bold = True
    p.font.size = Pt(14)
    p.font.color.rgb = ACCENT_TEAL
    p.space_after = Pt(12)

    cred_rows = [
        ("Student Intern", "Aditya Kumar"),
        ("Roll Number", "024101010012"),
        ("Certificate ID", "IC26/11873"),
        ("Training Duration", "15 June 2026 – 18 July 2026 (6 Weeks)"),
        ("Technical Domain", "Artificial Intelligence & NLP"),
        ("University Mentor", "Pooja Parmar"),
        ("Head of Department", "Dr. Himanshu Verma")
    ]
    for label, val in cred_rows:
        p = tfv.add_paragraph()
        p.text = f"{label}:"
        p.font.bold = True
        p.font.size = Pt(10)
        p.font.color.rgb = RGBColor(148, 163, 184)
        p = tfv.add_paragraph()
        p.text = val
        p.font.size = Pt(10.5)
        p.font.color.rgb = RGBColor(255, 255, 255)
        p.space_after = Pt(4)

    # =========================================================================
    # SLIDE 5: RAG SYSTEM ARCHITECTURE & DATAFLOW
    # =========================================================================
    s5 = prs.slides.add_slide(blank_layout)
    add_slide_header(s5, "RAG Pipeline: End-to-End System Architecture")
    add_slide_footer(s5)

    stages = [
        ("1. Ingestion", "PyPDF / PyMuPDF\nText extraction & noise cleaning"),
        ("2. Chunking", "Recursive Splitter\n500 tokens, 50 sliding overlap"),
        ("3. Embedding", "all-MiniLM-L6-v2\n384-dimensional dense vectors"),
        ("4. Vector Store", "ChromaDB\nHNSW graph indexing & search"),
        ("5. Retrieval", "Cosine Similarity\nTop-K (K=4) candidate search"),
        ("6. Augmentation", "Prompt Engine\nContext injection + citations"),
        ("7. Inference", "Ollama / Llama-3\nGrounded factual answering")
    ]

    for idx, (stitle, sdesc) in enumerate(stages):
        l = Inches(0.8 + idx * 1.68)
        add_card(s5, l, Inches(2.0), Inches(1.6), Inches(3.2), bg_color=BG_LIGHT)
        
        # Header strip
        hs = s5.shapes.add_shape(MSO_SHAPE.RECTANGLE, l, Inches(2.0), Inches(1.6), Inches(0.5))
        hs.fill.solid()
        hs.fill.fore_color.rgb = NAVY_DARK
        hs.line.fill.background()
        p = hs.text_frame.paragraphs[0]
        p.text = stitle
        p.font.bold = True
        p.font.size = Pt(9.5)
        p.font.color.rgb = RGBColor(255, 255, 255)
        p.alignment = PP_ALIGN.CENTER
        
        tb = s5.shapes.add_textbox(l + Inches(0.08), Inches(2.6), Inches(1.44), Inches(2.4))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = sdesc
        p.font.size = Pt(9)
        p.font.color.rgb = SLATE_TEXT

    # Lower Flow Formula Card
    add_card(s5, Inches(0.8), Inches(5.5), Inches(11.733), Inches(1.1), bg_color=NAVY_DARK, border_color=NAVY_LIGHT)
    tb_f = s5.shapes.add_textbox(Inches(1.0), Inches(5.6), Inches(11.333), Inches(0.9))
    tff = tb_f.text_frame
    tff.word_wrap = True
    p = tff.paragraphs[0]
    p.text = "Mathematical Principle: Dense Cosine Similarity Matching in R^384 Space"
    p.font.bold = True
    p.font.size = Pt(11)
    p.font.color.rgb = ACCENT_TEAL
    p = tff.add_paragraph()
    p.text = "CosineSimilarity(Query, Chunk_i) = (E(Query) · E(Chunk_i)) / ( ||E(Query)|| * ||E(Chunk_i)|| )  — Matches conceptual semantic meaning even with 0% lexical keyword overlap."
    p.font.name = "Consolas"
    p.font.size = Pt(10)
    p.font.color.rgb = RGBColor(241, 245, 249)

    # =========================================================================
    # SLIDE 6: TECHNOLOGIES & TOOLS STACK
    # =========================================================================
    s6 = prs.slides.add_slide(blank_layout)
    add_slide_header(s6, "Technology Stack & Development Environment")
    add_slide_footer(s6)

    tech_categories = [
        ("AI & Orchestration", [
            ("Python 3.11+", "Primary language for high-performance ML pipelines."),
            ("LangChain", "Decoupled pipeline chaining, loaders, and splitters."),
            ("Sentence-Transformers", "all-MiniLM-L6-v2 transformer model (384-dim).")
        ]),
        ("Database & Storage", [
            ("ChromaDB", "Embedded vector database with HNSW graph indexing."),
            ("PyPDF & PyMuPDF", "High-fidelity text & table stream extraction."),
            ("SQLite3", "Session state & chat interaction metadata persistence.")
        ]),
        ("Front-End & Deployment", [
            ("Streamlit", "Full-stack reactive web dashboard with live citation cards."),
            ("Git & GitHub", "Collaborative version control and artifact hosting."),
            ("WordCraft", "Native engine for professional report verification & conversion.")
        ])
    ]

    for idx, (cat_name, items) in enumerate(tech_categories):
        l = Inches(0.8 + idx * 3.95)
        add_card(s6, l, Inches(1.8), Inches(3.8), Inches(4.7), bg_color=BG_LIGHT)
        
        hb = s6.shapes.add_shape(MSO_SHAPE.RECTANGLE, l, Inches(1.8), Inches(3.8), Inches(0.6))
        hb.fill.solid()
        hb.fill.fore_color.rgb = NAVY_DARK
        hb.line.fill.background()
        p = hb.text_frame.paragraphs[0]
        p.text = cat_name
        p.font.bold = True
        p.font.size = Pt(13)
        p.font.color.rgb = RGBColor(255, 255, 255)
        p.alignment = PP_ALIGN.CENTER

        tb = s6.shapes.add_textbox(l + Inches(0.2), Inches(2.6), Inches(3.4), Inches(3.7))
        tf = tb.text_frame
        tf.word_wrap = True
        for i_idx, (t_name, t_desc) in enumerate(items):
            p = tf.paragraphs[0] if i_idx == 0 else tf.add_paragraph()
            p.text = t_name
            p.font.bold = True
            p.font.size = Pt(11.5)
            p.font.color.rgb = NAVY_LIGHT
            p2 = tf.add_paragraph()
            p2.text = t_desc
            p2.font.size = Pt(10)
            p2.font.color.rgb = SLATE_TEXT
            p2.space_after = Pt(8)

    # =========================================================================
    # SLIDE 7: EMPIRICAL BENCHMARKS & EVALUATION
    # =========================================================================
    s7 = prs.slides.add_slide(blank_layout)
    add_slide_header(s7, "Experimental Evaluation: Chunking Benchmarks")
    add_slide_footer(s7)

    # Benchmark Table
    add_card(s7, Inches(0.8), Inches(1.8), Inches(11.733), Inches(3.2), bg_color=CARD_BG)
    
    rows, cols = 5, 5
    table_shape = s7.shapes.add_table(rows, cols, Inches(1.0), Inches(2.0), Inches(11.333), Inches(2.8))
    table = table_shape.table
    
    # Column widths
    table.columns[0].width = Inches(2.2)
    table.columns[1].width = Inches(1.8)
    table.columns[2].width = Inches(1.8)
    table.columns[3].width = Inches(1.8)
    table.columns[4].width = Inches(3.733)

    headers = ["Chunk Configuration", "Sliding Overlap", "Retrieval Latency", "Hit Rate @ K=4", "Qualitative Failure Analysis"]
    for i, h in enumerate(headers):
        cell = table.cell(0, i)
        cell.text = h
        cell.fill.solid()
        cell.fill.fore_color.rgb = NAVY_DARK
        p = cell.text_frame.paragraphs[0]
        p.font.bold = True
        p.font.size = Pt(10)
        p.font.color.rgb = RGBColor(255, 255, 255)

    data = [
        ("Small (250 tokens)", "25 tokens", "11.2 ms", "68.4%", "Frequent sentence fragmentation; loses semantic context."),
        ("Medium (500 tokens) ★", "50 tokens", "14.8 ms", "89.6%", "Optimal balance; high semantic precision, complete concepts."),
        ("Large (1000 tokens)", "100 tokens", "22.6 ms", "74.2%", "Diluted semantic density; vector includes extraneous noise."),
        ("Extreme (2000 tokens)", "200 tokens", "38.1 ms", "59.8%", "High latency; prompt token limits exceeded, attention dispersion.")
    ]

    for r_idx, row_vals in enumerate(data):
        for c_idx, val in enumerate(row_vals):
            cell = table.cell(r_idx + 1, c_idx)
            cell.text = val
            cell.fill.solid()
            if r_idx == 1:
                cell.fill.fore_color.rgb = RGBColor(238, 242, 246) # highlight optimal
            else:
                cell.fill.fore_color.rgb = RGBColor(255, 255, 255) if r_idx % 2 == 0 else BG_LIGHT
            p = cell.text_frame.paragraphs[0]
            p.font.size = Pt(9.5)
            if r_idx == 1 and c_idx == 0:
                p.font.bold = True
                p.font.color.rgb = NAVY_DARK
            else:
                p.font.color.rgb = SLATE_TEXT

    # Takeaway card
    add_card(s7, Inches(0.8), Inches(5.3), Inches(11.733), Inches(1.3), bg_color=NAVY_DARK, border_color=ACCENT_TEAL)
    tb_t = s7.shapes.add_textbox(Inches(1.1), Inches(5.4), Inches(11.1), Inches(1.1))
    tft = tb_t.text_frame
    tft.word_wrap = True
    p = tft.paragraphs[0]
    p.text = "EMPIRICAL TAKEAWAY:"
    p.font.bold = True
    p.font.size = Pt(11)
    p.font.color.rgb = ACCENT_TEAL
    p = tft.add_paragraph()
    p.text = (
        "500-token chunks with 50-token sliding overlap achieved the peak Hit Rate (89.6%) while maintaining "
        "sub-15ms vector lookup latency on standard consumer hardware. This proved superior to larger windows where semantic precision degraded."
    )
    p.font.size = Pt(10.5)
    p.font.color.rgb = RGBColor(241, 245, 249)

    # =========================================================================
    # SLIDE 8: CHALLENGES FACED & SOLUTIONS
    # =========================================================================
    s8 = prs.slides.add_slide(blank_layout)
    add_slide_header(s8, "Key Engineering Challenges & Solutions")
    add_slide_footer(s8)

    challenges = [
        ("Chunk Boundary Context Loss",
         "Problem: Fixed-size character boundaries split crucial sentences across adjacent chunks, causing retrieval failures.\n\n"
         "Solution: Configured recursive hierarchical splitting (['\\n\\n', '\\n', '. ', ' ']) with a 10% sliding overlap window (50 tokens) to guarantee contextual continuity."),
        ("Token Limit & Cost Management",
         "Problem: Retrieving high K values inflated context payloads, degraded response generation speed, and caused 'lost-in-the-middle' attention loss.\n\n"
         "Solution: Tuned retrieval to Top-K = 4 and introduced a cosine relevance score threshold (0.65) to strictly discard noisy background chunks."),
        ("Tabular Data Extraction in PDFs",
         "Problem: Standard document extractors linearized matrix tables into continuous incoherent word streams, destroying relational structure.\n\n"
         "Solution: Integrated specialized table extractors that serialize PDF tables into Markdown syntax before embedding, enabling exact numerical reasoning.")
    ]

    for idx, (ctitle, cdesc) in enumerate(challenges):
        l = Inches(0.8 + idx * 3.95)
        add_card(s8, l, Inches(1.8), Inches(3.8), Inches(4.7), bg_color=BG_LIGHT)
        
        hb = s8.shapes.add_shape(MSO_SHAPE.RECTANGLE, l, Inches(1.8), Inches(3.8), Inches(0.6))
        hb.fill.solid()
        hb.fill.fore_color.rgb = NAVY_DARK
        hb.line.fill.background()
        p = hb.text_frame.paragraphs[0]
        p.text = ctitle
        p.font.bold = True
        p.font.size = Pt(11.5)
        p.font.color.rgb = RGBColor(255, 255, 255)
        p.alignment = PP_ALIGN.CENTER

        tb = s8.shapes.add_textbox(l + Inches(0.2), Inches(2.6), Inches(3.4), Inches(3.7))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = cdesc
        p.font.size = Pt(10)
        p.font.color.rgb = SLATE_TEXT

    # =========================================================================
    # SLIDE 9: USER INTERFACE & LIVE DEMO PREVIEW
    # =========================================================================
    s9 = prs.slides.add_slide(blank_layout)
    add_slide_header(s9, "Personal Document Q&A Assistant: UI Dashboard")
    add_slide_footer(s9)

    # Left: Screenshot 1
    add_card(s9, Inches(0.8), Inches(1.8), Inches(5.75), Inches(4.7), bg_color=BG_LIGHT)
    tb_img1_lbl = s9.shapes.add_textbox(Inches(1.0), Inches(1.9), Inches(5.35), Inches(0.4))
    p = tb_img1_lbl.text_frame.paragraphs[0]
    p.text = "Document Ingestion & Grounded Q&A Interface"
    p.font.bold = True
    p.font.size = Pt(11)
    p.font.color.rgb = NAVY_DARK
    
    img1_path = "artifacts/screenshots/ui_artifact_1_ingestion_chat.png"
    if os.path.exists(img1_path):
        s9.shapes.add_picture(img1_path, Inches(1.0), Inches(2.4), width=Inches(5.35))

    # Right: Screenshot 2
    add_card(s9, Inches(6.75), Inches(1.8), Inches(5.75), Inches(4.7), bg_color=BG_LIGHT)
    tb_img2_lbl = s9.shapes.add_textbox(Inches(6.95), Inches(1.9), Inches(5.35), Inches(0.4))
    p = tb_img2_lbl.text_frame.paragraphs[0]
    p.text = "Verified Citation Inspector & Relevance Scoring"
    p.font.bold = True
    p.font.size = Pt(11)
    p.font.color.rgb = NAVY_DARK

    img2_path = "artifacts/screenshots/ui_artifact_2_citations_inspector.png"
    if os.path.exists(img2_path):
        s9.shapes.add_picture(img2_path, Inches(6.95), Inches(2.4), width=Inches(5.35))

    # =========================================================================
    # SLIDE 10: CONCLUSION & FUTURE SCOPE
    # =========================================================================
    s10 = prs.slides.add_slide(blank_layout)
    add_slide_header(s10, "Conclusion & Future Research Directions")
    add_slide_footer(s10)

    # Left: Conclusion Card
    add_card(s10, Inches(0.8), Inches(1.8), Inches(5.75), Inches(4.7), bg_color=BG_LIGHT)
    tb_c = s10.shapes.add_textbox(Inches(1.1), Inches(2.1), Inches(5.15), Inches(4.2))
    tf = tb_c.text_frame
    tf.word_wrap = True
    
    p = tf.paragraphs[0]
    p.text = "Internship Conclusions:"
    p.font.bold = True
    p.font.size = Pt(14)
    p.font.color.rgb = NAVY_DARK
    p.space_after = Pt(8)

    conc_bullets = [
        "Eliminated Hallucination: The RAG assistant demonstrated that local factual grounding eliminates neural hallucinations.",
        "High Performance: Achieved an 89.6% retrieval hit rate with sub-15ms vector lookup latency using all-MiniLM-L6-v2 and ChromaDB.",
        "Privacy-Preserving: Entire pipeline operates locally without streaming sensitive enterprise/academic data to commercial cloud APIs.",
        "Comprehensive Deliverables: Successfully submitted the 23-page formal report, source code repository, and interactive UI."
    ]
    for b in conc_bullets:
        p = tf.add_paragraph()
        p.text = "• " + b
        p.font.size = Pt(10.5)
        p.font.color.rgb = SLATE_TEXT
        p.space_after = Pt(6)

    # Right: Future Scope Card
    add_card(s10, Inches(6.75), Inches(1.8), Inches(5.75), Inches(4.7), bg_color=BG_LIGHT)
    tb_fs = s10.shapes.add_textbox(Inches(7.05), Inches(2.1), Inches(5.15), Inches(4.2))
    tffs = tb_fs.text_frame
    tffs.word_wrap = True
    
    p = tffs.paragraphs[0]
    p.text = "Future Research Directions:"
    p.font.bold = True
    p.font.size = Pt(14)
    p.font.color.rgb = NAVY_DARK
    p.space_after = Pt(8)

    future_bullets = [
        "Hybrid Search (BM25 + Dense Vectors): Integrate Reciprocal Rank Fusion (RRF) to capture exact keyword acronyms alongside semantic concepts.",
        "Cross-Encoder Re-Rankers: Deploy BAAI/bge-reranker-large to re-score Top-20 retrieved candidates, maximizing Top-3 precision.",
        "Multimodal OCR Parsing: Integrate PaddleOCR/Tesseract to seamlessly index scanned historical documents and architectural schematics.",
        "Multi-Hop Agentic Workflows: Equip the assistant with autonomous multi-step planning to decompose complex cross-document inquiries."
    ]
    for b in future_bullets:
        p = tffs.add_paragraph()
        p.text = "➔ " + b
        p.font.size = Pt(10.5)
        p.font.color.rgb = SLATE_TEXT
        p.space_after = Pt(6)

    # =========================================================================
    # SLIDE 11: CLOSING & Q&A SLIDE
    # =========================================================================
    s11 = prs.slides.add_slide(blank_layout)
    bg11 = s11.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
    bg11.fill.solid()
    bg11.fill.fore_color.rgb = NAVY_DARK
    bg11.line.fill.background()

    tb_end = s11.shapes.add_textbox(Inches(1.5), Inches(1.5), Inches(10.333), Inches(4.5))
    tfe = tb_end.text_frame
    tfe.word_wrap = True
    
    p = tfe.paragraphs[0]
    p.text = "THANK YOU!"
    p.font.name = "Calibri"
    p.font.size = Pt(36)
    p.font.bold = True
    p.font.color.rgb = RGBColor(255, 255, 255)
    p.alignment = PP_ALIGN.CENTER
    p.space_after = Pt(8)

    p = tfe.add_paragraph()
    p.text = "Questions & Academic Discussion"
    p.font.name = "Calibri"
    p.font.size = Pt(20)
    p.font.color.rgb = ACCENT_TEAL
    p.alignment = PP_ALIGN.CENTER
    p.space_after = Pt(28)

    p = tfe.add_paragraph()
    p.text = (
        "Aditya Kumar | B.Tech. Computer Science & Engineering (5th Semester)\n"
        "University Roll Number: 024101010012 | Enrolment Number: 024101010012\n"
        "Department of Computer Science & Engineering, Haridwar University, Roorkee\n\n"
        "Official GitHub Repository:\n"
        "https://github.com/aditya93416/personal-document-rag-assistant"
    )
    p.font.name = "Calibri"
    p.font.size = Pt(13)
    p.font.color.rgb = RGBColor(226, 232, 240)
    p.alignment = PP_ALIGN.CENTER

    output_pptx = "Aditya_Kumar_Internship_Presentation.pptx"
    prs.save(output_pptx)
    print(f"[SUCCESS] PowerPoint presentation saved: {output_pptx}")

if __name__ == "__main__":
    build_presentation()
