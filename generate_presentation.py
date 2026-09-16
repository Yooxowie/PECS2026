import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml import parse_xml

def create_presentation():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    
    # Brand Color Palette
    DARK_GREEN = RGBColor(27, 67, 50)      # #1B4332 Coastal Forest Green
    DEEP_BG = RGBColor(16, 42, 31)         # Darker Green for Cover & Conclusion
    HARVEST_GOLD = RGBColor(212, 175, 55)  # #D4AF37 Harvest Gold
    LIGHT_GOLD = RGBColor(243, 222, 138)   # Light Gold Accent
    OCEAN_AQUA = RGBColor(42, 157, 143)    # #2A9D8F Ocean Aqua
    SOFT_SAND = RGBColor(248, 249, 250)    # #F8F9FA Soft Sand
    WHITE = RGBColor(255, 255, 255)
    CARD_BG = RGBColor(255, 255, 255)
    CARD_BORDER = RGBColor(220, 228, 224)
    TEXT_DARK = RGBColor(30, 41, 35)       # Dark Charcoal
    TEXT_MUTED = RGBColor(95, 110, 100)    # Muted Green-Gray
    PILL_BG = RGBColor(232, 243, 239)

    blank_slide_layout = prs.slide_layouts[6]

    def apply_slide_transition(slide, transition_type="fade"):
        if transition_type == "fade":
            trans_xml = parse_xml('<p:transition xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" spd="med"><p:fade/></p:transition>')
        elif transition_type == "push":
            trans_xml = parse_xml('<p:transition xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" spd="fast"><p:push dir="r"/></p:transition>')
        elif transition_type == "wipe":
            trans_xml = parse_xml('<p:transition xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" spd="med"><p:wipe dir="r"/></p:transition>')
        else:
            trans_xml = parse_xml('<p:transition xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" spd="med"><p:fade/></p:transition>')
        slide._element.append(trans_xml)

    def add_header(slide, title_text, category_text="PECS 2026 | BAM 205 EVENT PROPOSAL"):
        top_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(1.15))
        top_bar.fill.solid()
        top_bar.fill.fore_color.rgb = DARK_GREEN
        top_bar.line.fill.background()
        
        accent_line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(1.15), Inches(13.333), Inches(0.06))
        accent_line.fill.solid()
        accent_line.fill.fore_color.rgb = HARVEST_GOLD
        accent_line.line.fill.background()

        cat_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.12), Inches(11.7), Inches(0.3))
        tf_cat = cat_box.text_frame
        tf_cat.word_wrap = True
        p_cat = tf_cat.paragraphs[0]
        p_cat.text = category_text.upper()
        p_cat.font.size = Pt(10)
        p_cat.font.bold = True
        p_cat.font.color.rgb = HARVEST_GOLD
        p_cat.font.name = "Arial"

        title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.38), Inches(11.7), Inches(0.65))
        tf_title = title_box.text_frame
        tf_title.word_wrap = True
        p_title = tf_title.paragraphs[0]
        p_title.text = title_text
        p_title.font.size = Pt(21)
        p_title.font.bold = True
        p_title.font.color.rgb = WHITE
        p_title.font.name = "Georgia"

        footer_line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(7.05), Inches(11.733), Inches(0.02))
        footer_line.fill.solid()
        footer_line.fill.fore_color.rgb = CARD_BORDER
        footer_line.line.fill.background()

        footer_box = slide.shapes.add_textbox(Inches(0.8), Inches(7.08), Inches(11.733), Inches(0.35))
        tf_foot = footer_box.text_frame
        p_foot = tf_foot.paragraphs[0]
        p_foot.text = "PHINMA University of Pangasinan • College of Hospitality and Tourism Management • BAM 205: MICE"
        p_foot.font.size = Pt(9)
        p_foot.font.color.rgb = TEXT_MUTED
        p_foot.font.name = "Arial"

    def set_slide_bg(slide, color=SOFT_SAND):
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5))
        bg.fill.solid()
        bg.fill.fore_color.rgb = color
        bg.line.fill.background()
        return bg

    def create_card(slide, left, top, width, height, bg_color=CARD_BG, border_color=CARD_BORDER):
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
        card.fill.solid()
        card.fill.fore_color.rgb = bg_color
        if border_color:
            card.line.color.rgb = border_color
            card.line.width = Pt(1)
        else:
            card.line.fill.background()
        return card

    # ==========================================
    # SLIDE 1: TITLE SLIDE (Dark Theme)
    # ==========================================
    s1 = prs.slides.add_slide(blank_slide_layout)
    set_slide_bg(s1, DEEP_BG)
    apply_slide_transition(s1, "fade")

    accent_bar = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(0.3), Inches(7.5))
    accent_bar.fill.solid()
    accent_bar.fill.fore_color.rgb = HARVEST_GOLD
    accent_bar.line.fill.background()

    tag = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.2), Inches(0.8), Inches(4.5), Inches(0.45))
    tag.fill.solid()
    tag.fill.fore_color.rgb = OCEAN_AQUA
    tag.line.fill.background()
    tf_tag = tag.text_frame
    tf_tag.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_tag = tf_tag.paragraphs[0]
    p_tag.alignment = PP_ALIGN.CENTER
    p_tag.text = "BAM 205 • MICE EVENT PROPOSAL"
    p_tag.font.size = Pt(11)
    p_tag.font.bold = True
    p_tag.font.color.rgb = WHITE
    p_tag.font.name = "Arial"

    t_box = s1.shapes.add_textbox(Inches(1.2), Inches(1.45), Inches(11.0), Inches(1.9))
    tf = t_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "PECS 2026"
    p.font.size = Pt(44)
    p.font.bold = True
    p.font.color.rgb = HARVEST_GOLD
    p.font.name = "Georgia"

    p2 = tf.add_paragraph()
    p2.text = "Pangasinan Eco-Haven & Culinary Showcase"
    p2.font.size = Pt(28)
    p2.font.bold = True
    p2.font.color.rgb = WHITE
    p2.font.name = "Georgia"
    p2.space_before = Pt(5)

    p3 = tf.add_paragraph()
    p3.text = "Regional Conference on Sustainable Tourism & Heritage Culinary Innovation"
    p3.font.size = Pt(15)
    p3.font.italic = True
    p3.font.color.rgb = LIGHT_GOLD
    p3.font.name = "Georgia"
    p3.space_before = Pt(8)

    div = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1.2), Inches(3.7), Inches(10.8), Inches(0.04))
    div.fill.solid()
    div.fill.fore_color.rgb = HARVEST_GOLD
    div.line.fill.background()

    card_data = [
        ("📅 DATE & TIME", "Thursday, September 17, 2026\n8:30 AM – 3:30 PM (7-Hour Program)"),
        ("📍 VENUE", "PHINMA-University of Pangasinan\nGymnasium, Dagupan City"),
        ("👥 AUDIENCE", "500+ Seated Delegates\nBSHM, BSTM, Faculty & Leaders")
    ]
    for i, (head, sub) in enumerate(card_data):
        c_left = Inches(1.2 + i * 3.7)
        c = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, c_left, Inches(4.0), Inches(3.4), Inches(1.6))
        c.fill.solid()
        c.fill.fore_color.rgb = DARK_GREEN
        c.line.color.rgb = OCEAN_AQUA
        c.line.width = Pt(1)
        
        tb = s1.shapes.add_textbox(c_left + Inches(0.15), Inches(4.1), Inches(3.1), Inches(1.4))
        ctf = tb.text_frame
        ctf.word_wrap = True
        cp1 = ctf.paragraphs[0]
        cp1.text = head
        cp1.font.size = Pt(12)
        cp1.font.bold = True
        cp1.font.color.rgb = HARVEST_GOLD
        cp1.font.name = "Arial"
        
        cp2 = ctf.add_paragraph()
        cp2.text = sub
        cp2.font.size = Pt(11)
        cp2.font.color.rgb = SOFT_SAND
        cp2.font.name = "Arial"
        cp2.space_before = Pt(4)

    inst_box = s1.shapes.add_textbox(Inches(1.2), Inches(6.0), Inches(11.0), Inches(0.8))
    itf = inst_box.text_frame
    ip = itf.paragraphs[0]
    ip.text = "College of Hospitality and Tourism Management (CHTM)"
    ip.font.size = Pt(13)
    ip.font.bold = True
    ip.font.color.rgb = WHITE
    ip.font.name = "Georgia"
    
    ip2 = itf.add_paragraph()
    ip2.text = "PHINMA University of Pangasinan — Dagupan City, Pangasinan"
    ip2.font.size = Pt(11)
    ip2.font.color.rgb = TEXT_MUTED
    ip2.font.name = "Arial"

    # ==========================================
    # SLIDE 2: EXECUTIVE SUMMARY
    # ==========================================
    s2 = prs.slides.add_slide(blank_slide_layout)
    set_slide_bg(s2)
    apply_slide_transition(s2, "push")
    add_header(s2, "Executive Summary & Event Classification", "Section 1 & 2 • Event Identity")

    create_card(s2, Inches(0.8), Inches(1.5), Inches(5.6), Inches(5.3))
    t1 = s2.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.5), Inches(5.6), Inches(0.55))
    t1.fill.solid()
    t1.fill.fore_color.rgb = DARK_GREEN
    t1.line.fill.background()
    t1_p = t1.text_frame.paragraphs[0]
    t1_p.text = "EVENT ESSENTIALS"
    t1_p.font.bold = True
    t1_p.font.size = Pt(12)
    t1_p.font.color.rgb = HARVEST_GOLD
    t1_p.font.name = "Arial"

    tb_left = s2.shapes.add_textbox(Inches(1.0), Inches(2.2), Inches(5.2), Inches(4.4))
    ltf = tb_left.text_frame
    ltf.word_wrap = True

    items = [
        ("Full Event Title:", "PECS 2026: Pangasinan Eco-Haven & Culinary Showcase"),
        ("Official Subtitle:", "Regional Conference on Sustainable Tourism & Heritage Culinary Innovation"),
        ("Course Alignment:", "BAM 205: Meetings, Incentives, Conferences, and Exhibitions (MICE)"),
        ("Academic Host:", "College of Hospitality and Tourism Management (CHTM), PHINMA University of Pangasinan"),
        ("Event Classification:", "Academic & Professional Conference (MICE Industry Standard)")
    ]
    for i, (label, val) in enumerate(items):
        p = ltf.paragraphs[0] if i == 0 else ltf.add_paragraph()
        p.text = label
        p.font.bold = True
        p.font.size = Pt(12)
        p.font.color.rgb = DARK_GREEN
        p.font.name = "Arial"
        if i > 0:
            p.space_before = Pt(8)
            
        pv = ltf.add_paragraph()
        pv.text = val
        pv.font.size = Pt(11)
        pv.font.color.rgb = TEXT_DARK
        pv.font.name = "Arial"

    create_card(s2, Inches(6.8), Inches(1.5), Inches(5.7), Inches(5.3))
    t2 = s2.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(6.8), Inches(1.5), Inches(5.7), Inches(0.55))
    t2.fill.solid()
    t2.fill.fore_color.rgb = DARK_GREEN
    t2.line.fill.background()
    t2_p = t2.text_frame.paragraphs[0]
    t2_p.text = "STRATEGIC PURPOSE & VALUE"
    t2_p.font.bold = True
    t2_p.font.size = Pt(12)
    t2_p.font.color.rgb = HARVEST_GOLD
    t2_p.font.name = "Arial"

    tb_right = s2.shapes.add_textbox(Inches(7.0), Inches(2.2), Inches(5.3), Inches(4.4))
    rtf = tb_right.text_frame
    rtf.word_wrap = True

    right_items = [
        ("🌿 Eco-Haven Focus", "Addressing coastal environmental preservation, carrying capacity, and eco-tourism frameworks in Lingayen Gulf and Pangasinan municipalities."),
        ("🍲 Heritage Culinary Innovation", "Safeguarding authentic regional gastronomic treasures (Bangus, Alaminos Longganisa, Kakanin) while integrating modern culinary and hospitality techniques."),
        ("🎓 Academic & Industry Synergy", "Bridging classroom theoretical knowledge with live industry insights, research presentations, and municipal tourism governance.")
    ]
    for i, (title, desc) in enumerate(right_items):
        p = rtf.paragraphs[0] if i == 0 else rtf.add_paragraph()
        p.text = title
        p.font.bold = True
        p.font.size = Pt(13)
        p.font.color.rgb = OCEAN_AQUA
        p.font.name = "Georgia"
        if i > 0:
            p.space_before = Pt(12)
            
        pd = rtf.add_paragraph()
        pd.text = desc
        pd.font.size = Pt(11)
        pd.font.color.rgb = TEXT_DARK
        pd.font.name = "Arial"
        pd.space_before = Pt(3)

    # ==========================================
    # SLIDE 3: THEME & CONCEPTUAL FRAMEWORK
    # ==========================================
    s3 = prs.slides.add_slide(blank_slide_layout)
    set_slide_bg(s3)
    apply_slide_transition(s3, "push")
    add_header(s3, "Theme & Conference Concept", "Section 3 • Vision & Direction")

    theme_banner = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.45), Inches(11.733), Inches(1.3))
    theme_banner.fill.solid()
    theme_banner.fill.fore_color.rgb = DARK_GREEN
    theme_banner.line.color.rgb = HARVEST_GOLD
    theme_banner.line.width = Pt(1.5)

    tb_th = s3.shapes.add_textbox(Inches(1.0), Inches(1.5), Inches(11.3), Inches(1.2))
    th_tf = tb_th.text_frame
    th_tf.word_wrap = True
    th_p1 = th_tf.paragraphs[0]
    th_p1.text = "OFFICIAL CONFERENCE THEME"
    th_p1.font.bold = True
    th_p1.font.size = Pt(10)
    th_p1.font.color.rgb = HARVEST_GOLD
    th_p1.font.name = "Arial"

    th_p2 = th_tf.add_paragraph()
    th_p2.text = "“Sustaining the Coastline: Strategies for Eco-Tourism Development and Heritage Culinary Preservation in Pangasinan”"
    th_p2.font.bold = True
    th_p2.font.size = Pt(15)
    th_p2.font.color.rgb = WHITE
    th_p2.font.name = "Georgia"
    th_p2.space_before = Pt(3)

    pillars = [
        ("01. PLENARY KEYNOTES", "Expert Dialogues", "Gathering municipal tourism officers, resort general managers, and local culinary historians to present frameworks on coastal carrying capacity and indigenous gastronomic preservation.", DARK_GREEN),
        ("02. RESEARCH COLLOQUIUM", "Student Innovation", "Oral presentations of 4 shortlisted student research papers addressing empirical resort models, community eco-tourism, and sustainable heritage practices.", OCEAN_AQUA),
        ("03. MULTI-STAKEHOLDER Q&A", "Open Academic Forum", "Direct interactive floor microphone sessions bridging BSHM/BSTM students with practitioners, fostering active inquiry and policy recommendations.", DARK_GREEN)
    ]

    for i, (p_title, p_sub, p_body, color) in enumerate(pillars):
        x = Inches(0.8 + i * 4.0)
        c = create_card(s3, x, Inches(2.95), Inches(3.733), Inches(3.85))
        
        strip = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, Inches(2.95), Inches(3.733), Inches(0.6))
        strip.fill.solid()
        strip.fill.fore_color.rgb = color
        strip.line.fill.background()
        strip_p = strip.text_frame.paragraphs[0]
        strip_p.text = p_title
        strip_p.font.bold = True
        strip_p.font.size = Pt(11)
        strip_p.font.color.rgb = WHITE if color != HARVEST_GOLD else DARK_GREEN
        strip_p.font.name = "Arial"
        strip_p.alignment = PP_ALIGN.CENTER
        
        tb_b = s3.shapes.add_textbox(x + Inches(0.2), Inches(3.7), Inches(3.333), Inches(2.9))
        btf = tb_b.text_frame
        btf.word_wrap = True
        
        sub_p = btf.paragraphs[0]
        sub_p.text = p_sub
        sub_p.font.bold = True
        sub_p.font.size = Pt(13)
        sub_p.font.color.rgb = DARK_GREEN
        sub_p.font.name = "Georgia"
        
        bp = btf.add_paragraph()
        bp.text = p_body
        bp.font.size = Pt(11)
        bp.font.color.rgb = TEXT_DARK
        bp.font.name = "Arial"
        bp.space_before = Pt(8)

    # ==========================================
    # SLIDE 4: TARGET AUDIENCE & STAKEHOLDERS
    # ==========================================
    s4 = prs.slides.add_slide(blank_slide_layout)
    set_slide_bg(s4)
    apply_slide_transition(s4, "push")
    add_header(s4, "Target Audience & Stakeholder Profile", "Section 4 • Delegate & Speaker Segmentation")

    create_card(s4, Inches(0.8), Inches(1.5), Inches(5.65), Inches(5.3))
    hdr_p = s4.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.5), Inches(5.65), Inches(0.6))
    hdr_p.fill.solid()
    hdr_p.fill.fore_color.rgb = DARK_GREEN
    hdr_p.line.fill.background()
    hp_text = hdr_p.text_frame.paragraphs[0]
    hp_text.text = "PRIMARY AUDIENCE — DELEGATES"
    hp_text.font.bold = True
    hp_text.font.size = Pt(12)
    hp_text.font.color.rgb = HARVEST_GOLD
    hp_text.font.name = "Arial"

    tb_p = s4.shapes.add_textbox(Inches(1.05), Inches(2.25), Inches(5.15), Inches(4.3))
    ptf = tb_p.text_frame
    ptf.word_wrap = True

    p_data = [
        ("👤 Delegate Profile:", "Undergraduate BSHM (Hospitality) and BSTM (Tourism) students, student researchers, and CHTM academic faculty."),
        ("📊 Capacity & Volume:", "500+ Seated Delegates inside the UPang Gymnasium main floor."),
        ("🎯 Critical Needs & Expectations:", "• Practical industry insights beyond textbooks\n• Peer research validation & academic exchange\n• Professional networking opportunities\n• Complete conference kits (ID badge, notepad, program, bag, e-certificate)")
    ]
    for i, (l, v) in enumerate(p_data):
        p = ptf.paragraphs[0] if i == 0 else ptf.add_paragraph()
        p.text = l
        p.font.bold = True
        p.font.size = Pt(12)
        p.font.color.rgb = DARK_GREEN
        p.font.name = "Arial"
        if i > 0:
            p.space_before = Pt(10)
        
        pv = ptf.add_paragraph()
        pv.text = v
        pv.font.size = Pt(11)
        pv.font.color.rgb = TEXT_DARK
        pv.font.name = "Arial"
        pv.space_before = Pt(2)

    create_card(s4, Inches(6.88), Inches(1.5), Inches(5.65), Inches(5.3))
    hdr_s = s4.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(6.88), Inches(1.5), Inches(5.65), Inches(0.6))
    hdr_s.fill.solid()
    hdr_s.fill.fore_color.rgb = OCEAN_AQUA
    hdr_s.line.fill.background()
    hs_text = hdr_s.text_frame.paragraphs[0]
    hs_text.text = "SECONDARY AUDIENCE — VIPs & SPEAKERS"
    hs_text.font.bold = True
    hs_text.font.size = Pt(12)
    hs_text.font.color.rgb = WHITE
    hs_text.font.name = "Arial"

    tb_s = s4.shapes.add_textbox(Inches(7.13), Inches(2.25), Inches(5.15), Inches(4.3))
    stf = tb_s.text_frame
    stf.word_wrap = True

    s_data = [
        ("🎖️ VIP & Speaker Profile:", "Municipal Tourism Officers (MTOs), Resort General Managers, Local Culinary Historians, and Senior Faculty Reactors."),
        ("📊 Capacity & Volume:", "8–10 Invited Plenary Keynote Speakers, Panelists, and Reactors."),
        ("🎯 Critical Needs & Expectations:", "• Strict program pacing and structured timing\n• Clear stage AV, roving wireless mics, and confidence monitors\n• Dedicated VIP holding lounge with private plated catering\n• Formal honoraria, local tokens of appreciation, and plaques")
    ]
    for i, (l, v) in enumerate(s_data):
        p = stf.paragraphs[0] if i == 0 else stf.add_paragraph()
        p.text = l
        p.font.bold = True
        p.font.size = Pt(12)
        p.font.color.rgb = DARK_GREEN
        p.font.name = "Arial"
        if i > 0:
            p.space_before = Pt(10)
        
        pv = stf.add_paragraph()
        pv.text = v
        pv.font.size = Pt(11)
        pv.font.color.rgb = TEXT_DARK
        pv.font.name = "Arial"
        pv.space_before = Pt(2)

    # ==========================================
    # SLIDE 5: SMART OBJECTIVES
    # ==========================================
    s5 = prs.slides.add_slide(blank_slide_layout)
    set_slide_bg(s5)
    apply_slide_transition(s5, "wipe")
    add_header(s5, "SMART Event Objectives", "Section 5 • Measurable Goals & Outcomes")

    smart_items = [
        ("S", "SPECIFIC", "Host a regional academic conference featuring 3 keynote sessions and 4 student research panel presentations on Pangasinan tourism development and culinary heritage.", DARK_GREEN),
        ("M", "MEASURABLE", "Achieve attendance of at least 500 registered delegates and collect completed session feedback forms from at least 90% of attendees.", OCEAN_AQUA),
        ("A", "ACHIEVABLE", "Secure regional tourism officials and faculty experts as session speakers and panel reactors using the university's central facility.", DARK_GREEN),
        ("R", "RELEVANT", "Fulfill BAM 205 conference management competencies including large-scale program structuring, speaker management, main-stage AV logistics, and delegate registration.", OCEAN_AQUA),
        ("T", "TIME-BOUND", "Execute the complete plenary schedule, panel sessions, open forum, and awarding ceremony within a 7-hour operational timeframe on September 17, 2026.", DARK_GREEN)
    ]

    card_height = Inches(0.96)
    spacing = Inches(0.12)
    start_top = Inches(1.45)

    for i, (letter, label, desc, color) in enumerate(smart_items):
        y = start_top + i * (card_height + spacing)
        create_card(s5, Inches(0.8), y, Inches(11.733), card_height)
        
        badge = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.95), y + Inches(0.1), Inches(0.76), Inches(0.76))
        badge.fill.solid()
        badge.fill.fore_color.rgb = color
        badge.line.fill.background()
        bp = badge.text_frame.paragraphs[0]
        bp.text = letter
        bp.font.bold = True
        bp.font.size = Pt(22)
        bp.font.color.rgb = HARVEST_GOLD if color == DARK_GREEN else WHITE
        bp.font.name = "Georgia"
        bp.alignment = PP_ALIGN.CENTER
        
        tb = s5.shapes.add_textbox(Inches(1.85), y + Inches(0.08), Inches(10.5), card_height - Inches(0.16))
        tf = tb.text_frame
        tf.word_wrap = True
        
        p1 = tf.paragraphs[0]
        p1.text = label
        p1.font.bold = True
        p1.font.size = Pt(12)
        p1.font.color.rgb = DARK_GREEN
        p1.font.name = "Arial"
        
        p2 = tf.add_paragraph()
        p2.text = desc
        p2.font.size = Pt(10.5)
        p2.font.color.rgb = TEXT_DARK
        p2.font.name = "Arial"
        p2.space_before = Pt(2)

    # ==========================================
    # SLIDE 6: PROPOSED DATE & VENUE RATIONALE
    # ==========================================
    s6 = prs.slides.add_slide(blank_slide_layout)
    set_slide_bg(s6)
    apply_slide_transition(s6, "push")
    add_header(s6, "Date, Venue & Facilities Analysis", "Section 6 • Event Timing & Location")

    create_card(s6, Inches(0.8), Inches(1.5), Inches(4.5), Inches(5.3))
    v_hdr = s6.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.5), Inches(4.5), Inches(0.6))
    v_hdr.fill.solid()
    v_hdr.fill.fore_color.rgb = DARK_GREEN
    v_hdr.line.fill.background()
    v_hp = v_hdr.text_frame.paragraphs[0]
    v_hp.text = "EVENT LOGISTICS SNAPSHOT"
    v_hp.font.bold = True
    v_hp.font.size = Pt(12)
    v_hp.font.color.rgb = HARVEST_GOLD
    v_hp.font.name = "Arial"

    tb_v = s6.shapes.add_textbox(Inches(1.0), Inches(2.25), Inches(4.1), Inches(4.3))
    vtf = tb_v.text_frame
    vtf.word_wrap = True

    v_details = [
        ("🗓️ Proposed Date:", "Thursday, September 17, 2026"),
        ("⏰ Timeframe:", "8:30 AM – 3:30 PM (7 Hours Operational)"),
        ("📍 Official Venue:", "PHINMA University of Pangasinan Gymnasium"),
        ("📌 Location:", "Arellano Street, Dagupan City, Pangasinan"),
        ("🏛️ Facility Type:", "University Indoor Sports & Assembly Complex")
    ]
    for i, (l, val) in enumerate(v_details):
        p = vtf.paragraphs[0] if i == 0 else vtf.add_paragraph()
        p.text = l
        p.font.bold = True
        p.font.size = Pt(12)
        p.font.color.rgb = DARK_GREEN
        p.font.name = "Arial"
        if i > 0:
            p.space_before = Pt(8)
            
        pv = vtf.add_paragraph()
        pv.text = val
        pv.font.size = Pt(11)
        pv.font.color.rgb = TEXT_DARK
        pv.font.name = "Arial"
        pv.space_before = Pt(2)

    rat_cards = [
        ("👥 500+ Seating Capacity", "Ample floor and bleacher capacity to seat 500+ delegates comfortably while maintaining clear aisleways and safety egress.", Inches(5.6), Inches(1.5)),
        ("🎤 Elevated Stage & Backing", "Dedicated high main stage providing complete sightlines for plenary keynote lectures, slide projections, and ceremonies.", Inches(9.4), Inches(1.5)),
        ("🔊 Acoustic & AV Setup", "Integrated university sound reinforcement system, wireless lapel microphones, and LED/projector connectivity for crystal-clear talks.", Inches(5.6), Inches(4.2)),
        ("🚪 Dedicated Foyer & VIP Zone", "Spacious foyer for orderly multi-lane kit distribution, and secure air-conditioned VIP lounge for guest speaker dining.", Inches(9.4), Inches(4.2))
    ]

    for title, desc, rx, ry in rat_cards:
        create_card(s6, rx, ry, Inches(3.6), Inches(2.6))
        tb = s6.shapes.add_textbox(rx + Inches(0.2), ry + Inches(0.2), Inches(3.2), Inches(2.2))
        tf = tb.text_frame
        tf.word_wrap = True
        
        p1 = tf.paragraphs[0]
        p1.text = title
        p1.font.bold = True
        p1.font.size = Pt(12)
        p1.font.color.rgb = DARK_GREEN
        p1.font.name = "Georgia"
        
        p2 = tf.add_paragraph()
        p2.text = desc
        p2.font.size = Pt(10.5)
        p2.font.color.rgb = TEXT_DARK
        p2.font.name = "Arial"
        p2.space_before = Pt(6)

    # ==========================================
    # SLIDE 7: MORNING PLENARY PROGRAM FLOW
    # ==========================================
    s7 = prs.slides.add_slide(blank_slide_layout)
    set_slide_bg(s7)
    apply_slide_transition(s7, "push")
    add_header(s7, "Program Flow: Morning Plenary Sessions", "Section 7 (Part 1) • Schedule of Activities (08:30 AM – 12:00 PM)")

    rows = 6
    cols = 4
    left = Inches(0.8)
    top = Inches(1.5)
    width = Inches(11.733)
    height = Inches(5.2)

    table_shape = s7.shapes.add_table(rows, cols, left, top, width, height)
    table = table_shape.table

    table.columns[0].width = Inches(2.0)
    table.columns[1].width = Inches(3.2)
    table.columns[2].width = Inches(4.2)
    table.columns[3].width = Inches(2.333)

    morning_data = [
        ["Time", "Session Title", "Description / Content", "Person / Group in Charge"],
        ["08:30 AM – 09:00 AM", "Delegate Registration & Kit Distribution", "Sign-in at Gym foyer registration desks; distribution of conference badges, notepad kits, and session programs.", "Secretariat & Registration Team"],
        ["09:00 AM – 09:30 AM", "Opening Plenary & Welcome Address", "Doxology, National Anthem, and Opening Remarks by CHTM Leadership.", "Conference Master of Ceremonies"],
        ["09:30 AM – 10:30 AM", "Keynote 1: Coastal Eco-Tourism Frameworks", "Presentation on sustainable destination management across Pangasinan municipalities.", "Invited Municipal Tourism Officer"],
        ["10:30 AM – 11:30 AM", "Keynote 2: Preserving Heritage Flavor Profiles", "Industry talk on preserving regional culinary recipes and integrating them into modern hospitality.", "Guest Chef / Culinary Historian"],
        ["11:30 AM – 12:00 PM", "Open Forum & Delegate Q&A", "Floor microphones open for delegate inquiries directed to Keynote Speakers 1 and 2.", "Stage Moderator"]
    ]

    for r_idx, row_content in enumerate(morning_data):
        for c_idx, cell_value in enumerate(row_content):
            cell = table.cell(r_idx, c_idx)
            cell.text = cell_value
            p = cell.text_frame.paragraphs[0]
            p.font.name = "Arial"
            
            if r_idx == 0:
                cell.fill.solid()
                cell.fill.fore_color.rgb = DARK_GREEN
                p.font.bold = True
                p.font.size = Pt(11)
                p.font.color.rgb = HARVEST_GOLD
            else:
                cell.fill.solid()
                cell.fill.fore_color.rgb = WHITE if r_idx % 2 == 1 else SOFT_SAND
                p.font.size = Pt(10)
                p.font.color.rgb = TEXT_DARK
                if c_idx == 0:
                    p.font.bold = True
                    p.font.color.rgb = DARK_GREEN
                elif c_idx == 1:
                    p.font.bold = True

    # ==========================================
    # SLIDE 8: AFTERNOON COLLOQUIUM PROGRAM FLOW
    # ==========================================
    s8 = prs.slides.add_slide(blank_slide_layout)
    set_slide_bg(s8)
    apply_slide_transition(s8, "push")
    add_header(s8, "Program Flow: Afternoon Colloquium & Closing", "Section 7 (Part 2) • Schedule of Activities (12:00 PM – 03:30 PM)")

    table_shape_aft = s8.shapes.add_table(5, 4, left, top, width, Inches(4.3))
    table_aft = table_shape_aft.table

    table_aft.columns[0].width = Inches(2.0)
    table_aft.columns[1].width = Inches(3.2)
    table_aft.columns[2].width = Inches(4.2)
    table_aft.columns[3].width = Inches(2.333)

    aft_data = [
        ["Time", "Session Title", "Description / Content", "Person / Group in Charge"],
        ["12:00 PM – 01:00 PM", "Lunch Break & Networking", "Networking lunch for delegates; VIP dining in the designated Gym VIP Lounge.", "Logistics & Catering Team"],
        ["01:00 PM – 02:30 PM", "Panel Session: Student Research Colloquium", "Oral presentations of 4 selected student research papers on local resort models and eco-tourism.", "Student Presenters & Reactor Panel"],
        ["02:30 PM – 03:15 PM", "Awarding of Best Paper & Certificates", "Conferment of Plaque of Appreciation to speakers and Best Student Research Presentation.", "Conference Director & CHTM Dean"],
        ["03:15 PM – 03:30 PM", "Closing Remarks & Photo Session", "Formal adjournment of PECS 2026 Conference followed by group photo on the main stage.", "Conference Director"]
    ]

    for r_idx, row_content in enumerate(aft_data):
        for c_idx, cell_value in enumerate(row_content):
            cell = table_aft.cell(r_idx, c_idx)
            cell.text = cell_value
            p = cell.text_frame.paragraphs[0]
            p.font.name = "Arial"
            
            if r_idx == 0:
                cell.fill.solid()
                cell.fill.fore_color.rgb = DARK_GREEN
                p.font.bold = True
                p.font.size = Pt(11)
                p.font.color.rgb = HARVEST_GOLD
            else:
                cell.fill.solid()
                cell.fill.fore_color.rgb = WHITE if r_idx % 2 == 1 else SOFT_SAND
                p.font.size = Pt(10)
                p.font.color.rgb = TEXT_DARK
                if c_idx == 0:
                    p.font.bold = True
                    p.font.color.rgb = DARK_GREEN
                elif c_idx == 1:
                    p.font.bold = True

    create_card(s8, Inches(0.8), Inches(6.0), Inches(11.733), Inches(0.85), bg_color=PILL_BG, border_color=OCEAN_AQUA)
    tb_n = s8.shapes.add_textbox(Inches(1.0), Inches(6.05), Inches(11.3), Inches(0.75))
    ntf = tb_n.text_frame
    ntf.word_wrap = True
    np = ntf.paragraphs[0]
    np.text = "⏱️ Time Management Protocol:"
    np.font.bold = True
    np.font.size = Pt(11)
    np.font.color.rgb = DARK_GREEN
    np.font.name = "Arial"

    np2 = ntf.add_paragraph()
    np2.text = "Each student presenter is allocated 15 minutes presentation + 5 minutes reactor critique. Stage timer monitors will be visible to enforce punctual program execution."
    np2.font.size = Pt(10)
    np2.font.color.rgb = TEXT_DARK
    np2.font.name = "Arial"

    # ==========================================
    # SLIDE 9: BUDGET & FINANCIAL PLAN
    # ==========================================
    s9 = prs.slides.add_slide(blank_slide_layout)
    set_slide_bg(s9)
    apply_slide_transition(s9, "push")
    add_header(s9, "Estimated Budget & Financial Plan", "Section 8 • Cost Breakdown & Funding Mechanism")

    table_shape_b = s9.shapes.add_table(6, 3, Inches(0.8), Inches(1.5), Inches(7.8), Inches(5.2))
    table_b = table_shape_b.table

    table_b.columns[0].width = Inches(2.6)
    table_b.columns[1].width = Inches(1.3)
    table_b.columns[2].width = Inches(3.9)

    budget_data = [
        ["Expense Item", "Estimated Cost", "Notes / Details"],
        ["Keynote Speaker Honoraria & Tokens", "₱15,000", "Tokens of appreciation, local produce gift baskets, and honoraria for guest speakers."],
        ["Delegate Kits & Printed Materials", "₱20,000", "Lanyards, printed ID badges, notepad kits, pens, and certificates for 500 delegates."],
        ["Venue Audio-Visual & Stage Framing", "₱15,000", "Gymnasium stage LED wall/projector backdrop, wireless lapels, roving floor mics, and podium setup."],
        ["VIP & Speaker Catering", "₱10,000", "Plated lunch and AM/PM snacks for invited guest speakers and panel reactors."],
        ["TOTAL ESTIMATED BUDGET", "₱60,000", "Complete budget covering all 500 delegates and 10 VIP speakers."]
    ]

    for r_idx, row_content in enumerate(budget_data):
        for c_idx, cell_value in enumerate(row_content):
            cell = table_b.cell(r_idx, c_idx)
            cell.text = cell_value
            p = cell.text_frame.paragraphs[0]
            p.font.name = "Arial"
            
            if r_idx == 0:
                cell.fill.solid()
                cell.fill.fore_color.rgb = DARK_GREEN
                p.font.bold = True
                p.font.size = Pt(11)
                p.font.color.rgb = HARVEST_GOLD
            elif r_idx == 5:
                cell.fill.solid()
                cell.fill.fore_color.rgb = DARK_GREEN
                p.font.bold = True
                p.font.size = Pt(11)
                p.font.color.rgb = HARVEST_GOLD if c_idx == 1 else WHITE
            else:
                cell.fill.solid()
                cell.fill.fore_color.rgb = WHITE if r_idx % 2 == 1 else SOFT_SAND
                p.font.size = Pt(10)
                p.font.color.rgb = TEXT_DARK
                if c_idx == 0:
                    p.font.bold = True
                elif c_idx == 1:
                    p.font.bold = True
                    p.font.color.rgb = DARK_GREEN

    create_card(s9, Inches(8.8), Inches(1.5), Inches(3.733), Inches(2.4), bg_color=DARK_GREEN)
    tb_tot = s9.shapes.add_textbox(Inches(9.0), Inches(1.7), Inches(3.333), Inches(2.0))
    tot_tf = tb_tot.text_frame
    tot_tf.word_wrap = True
    
    tp1 = tot_tf.paragraphs[0]
    tp1.text = "TOTAL BUDGET"
    tp1.font.size = Pt(11)
    tp1.font.bold = True
    tp1.font.color.rgb = HARVEST_GOLD
    tp1.font.name = "Arial"
    
    tp2 = tot_tf.add_paragraph()
    tp2.text = "₱60,000"
    tp2.font.size = Pt(32)
    tp2.font.bold = True
    tp2.font.color.rgb = WHITE
    tp2.font.name = "Georgia"
    tp2.space_before = Pt(4)
    
    tp3 = tot_tf.add_paragraph()
    tp3.text = "Cost per delegate: ~₱120 (Highly cost-efficient academic staging)"
    tp3.font.size = Pt(10)
    tp3.font.color.rgb = LIGHT_GOLD
    tp3.font.name = "Arial"
    tp3.space_before = Pt(4)

    create_card(s9, Inches(8.8), Inches(4.1), Inches(3.733), Inches(2.6))
    tb_fund = s9.shapes.add_textbox(Inches(9.0), Inches(4.25), Inches(3.333), Inches(2.3))
    ftf = tb_fund.text_frame
    ftf.word_wrap = True
    
    fp1 = ftf.paragraphs[0]
    fp1.text = "FUNDING MECHANISM"
    fp1.font.size = Pt(11)
    fp1.font.bold = True
    fp1.font.color.rgb = DARK_GREEN
    fp1.font.name = "Arial"
    
    fp2 = ftf.add_paragraph()
    fp2.text = "1. Student Conference Fees\nCovering physical kits, lanyards, and certificates.\n\n2. CHTM Departmental Support\nSubsidizing speaker tokens, stage backdrops, and technical AV assets."
    fp2.font.size = Pt(10)
    fp2.font.color.rgb = TEXT_DARK
    fp2.font.name = "Arial"
    fp2.space_before = Pt(6)

    # ==========================================
    # SLIDE 10: MARKETING STRATEGY
    # ==========================================
    s10 = prs.slides.add_slide(blank_slide_layout)
    set_slide_bg(s10)
    apply_slide_transition(s10, "push")
    add_header(s10, "Marketing & Promotional Strategy", "Section 9 • Delegate Acquisition & Awareness")

    mkt_cards = [
        ("📱 Social Media Teaser Campaign", "Digital Engagement", "Launch targeted countdown graphics and keynote speaker spot-highlights across CHTM departmental Facebook & Instagram pages starting September 7, 2026 (10 days prior).", "Timeline: Sept 7 – Sept 17, 2026", DARK_GREEN),
        ("📌 Campus Bulletin Displays", "Physical Touchpoints", "Mount full-color promotional posters, QR code registration infographics, and program agendas across CHTM bulletin boards and high-traffic gym walkways.", "Coverage: High-foot-traffic halls", OCEAN_AQUA),
        ("🗣️ Class-to-Class Caravan", "Direct Outreach", "Conduct synchronized 3-minute promotional walkthroughs across all BSHM and BSTM lecture classes to explain conference learning outcomes and drive registration.", "Audience: 500+ BSHM/BSTM Students", DARK_GREEN)
    ]

    for i, (title, sub, body, stat, col) in enumerate(mkt_cards):
        x = Inches(0.8 + i * 4.0)
        create_card(s10, x, Inches(1.5), Inches(3.733), Inches(5.3))
        
        b = s10.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, Inches(1.5), Inches(3.733), Inches(0.6))
        b.fill.solid()
        b.fill.fore_color.rgb = col
        b.line.fill.background()
        bp = b.text_frame.paragraphs[0]
        bp.text = sub.upper()
        bp.font.bold = True
        bp.font.size = Pt(11)
        bp.font.color.rgb = HARVEST_GOLD if col == DARK_GREEN else WHITE
        bp.font.name = "Arial"
        bp.alignment = PP_ALIGN.CENTER
        
        tb = s10.shapes.add_textbox(x + Inches(0.2), Inches(2.3), Inches(3.333), Inches(3.6))
        tf = tb.text_frame
        tf.word_wrap = True
        
        p1 = tf.paragraphs[0]
        p1.text = title
        p1.font.bold = True
        p1.font.size = Pt(13)
        p1.font.color.rgb = DARK_GREEN
        p1.font.name = "Georgia"
        
        p2 = tf.add_paragraph()
        p2.text = body
        p2.font.size = Pt(10.5)
        p2.font.color.rgb = TEXT_DARK
        p2.font.name = "Arial"
        p2.space_before = Pt(8)
        
        pill = s10.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x + Inches(0.2), Inches(5.95), Inches(3.333), Inches(0.55))
        pill.fill.solid()
        pill.fill.fore_color.rgb = PILL_BG
        pill.line.color.rgb = OCEAN_AQUA
        pill.line.width = Pt(1)
        pp = pill.text_frame.paragraphs[0]
        pp.text = stat
        pp.font.bold = True
        pp.font.size = Pt(9.5)
        pp.font.color.rgb = DARK_GREEN
        pp.font.name = "Arial"
        pp.alignment = PP_ALIGN.CENTER

    # ==========================================
    # SLIDE 11: ORGANIZATIONAL ROLES
    # ==========================================
    s11 = prs.slides.add_slide(blank_slide_layout)
    set_slide_bg(s11)
    apply_slide_transition(s11, "push")
    add_header(s11, "Organizational Structure & Committee Roles", "Section 10 • BAM 205 Organizing Committee")

    roles = [
        ("🏆 Conference Director", "Executive Leadership", "Leads overall event management, approves presentation decks, enforces strict time control, and coordinates institutional approvals with CHTM Administration."),
        ("🎬 Stage & Program Lead", "Production Control", "Manages main-stage cues, slide transitions, stage timer displays, audio prompts, and coordinates floor moderators during plenary Q&A sessions."),
        ("📋 Secretariat & Registration Lead", "Delegate Experience", "Oversees delegate check-in at Gym foyer desks, physical kit distribution, issuance of digital certificates, and digital feedback form consolidation."),
        ("🤝 Speaker Relations Lead", "VIP Management", "Coordinates travel logistics, VIP holding room hospitality, honoraria distribution, and dedicated stage escorting for keynote speakers and panelists."),
        ("🎛️ Technical & AV Lead", "Infrastructure & Tech", "Manages Gymnasium audio levels, main LED backdrop screen inputs, stage lighting controls, livestreaming, and session recording gear.")
    ]

    for i, (title, sub, desc) in enumerate(roles):
        y = start_top + i * (card_height + spacing)
        create_card(s11, Inches(0.8), y, Inches(11.733), card_height)
        
        rt_box = s11.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.95), y + Inches(0.12), Inches(3.2), Inches(0.72))
        rt_box.fill.solid()
        rt_box.fill.fore_color.rgb = DARK_GREEN
        rt_box.line.fill.background()
        
        rtf = rt_box.text_frame
        rtf.word_wrap = True
        rp1 = rtf.paragraphs[0]
        rp1.text = title
        rp1.font.bold = True
        rp1.font.size = Pt(11)
        rp1.font.color.rgb = HARVEST_GOLD
        rp1.font.name = "Georgia"
        
        rp2 = rtf.add_paragraph()
        rp2.text = sub.upper()
        rp2.font.size = Pt(8.5)
        rp2.font.color.rgb = WHITE
        rp2.font.name = "Arial"
        
        desc_box = s11.shapes.add_textbox(Inches(4.3), y + Inches(0.1), Inches(8.0), card_height - Inches(0.2))
        dtf = desc_box.text_frame
        dtf.word_wrap = True
        dp = dtf.paragraphs[0]
        dp.text = desc
        dp.font.size = Pt(10.5)
        dp.font.color.rgb = TEXT_DARK
        dp.font.name = "Arial"

    # ==========================================
    # SLIDE 12: VISUAL IDENTITY
    # ==========================================
    s12 = prs.slides.add_slide(blank_slide_layout)
    set_slide_bg(s12)
    apply_slide_transition(s12, "push")
    add_header(s12, "Visual Identity & Environmental Design", "Section 11 • Brand Aesthetic & Stage Setup")

    swatches = [
        ("Coastal Forest Green", "#1B4332", "Primary Brand Color", "Represents Pangasinan eco-tourism; used for stage backdrop, staff polos, lanyards.", DARK_GREEN, WHITE),
        ("Harvest Gold", "#D4AF37", "Primary Accent Color", "Signifies culinary heritage (kakanin, golden bangus); used for titles, awards, VIP badges.", HARVEST_GOLD, DARK_GREEN),
        ("Ocean Aqua", "#2A9D8F", "Secondary Accent Color", "Symbolizes Lingayen Gulf and coastal management; used for icons, subheaders, post frames.", OCEAN_AQUA, WHITE),
        ("Soft Sand", "#F8F9FA", "Neutral Foundation", "High-contrast clean backdrop for legible slide presentations and printed collaterals.", SOFT_SAND, TEXT_DARK)
    ]

    for i, (name, hex_code, role, desc, col, text_col) in enumerate(swatches):
        x = Inches(0.8 + i * 3.0)
        cbox = s12.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, Inches(1.45), Inches(2.75), Inches(1.8))
        cbox.fill.solid()
        cbox.fill.fore_color.rgb = col
        cbox.line.color.rgb = CARD_BORDER
        cbox.line.width = Pt(1)
        
        ctf = cbox.text_frame
        ctf.word_wrap = True
        cp1 = ctf.paragraphs[0]
        cp1.text = name
        cp1.font.bold = True
        cp1.font.size = Pt(12)
        cp1.font.color.rgb = text_col
        cp1.font.name = "Georgia"
        
        cp2 = ctf.add_paragraph()
        cp2.text = f"{hex_code} • {role}"
        cp2.font.size = Pt(9)
        cp2.font.color.rgb = text_col
        cp2.font.name = "Arial"
        cp2.space_before = Pt(3)

    create_card(s12, Inches(0.8), Inches(3.45), Inches(5.65), Inches(3.35))
    e_hdr = s12.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(3.45), Inches(5.65), Inches(0.5))
    e_hdr.fill.solid()
    e_hdr.fill.fore_color.rgb = DARK_GREEN
    e_hdr.line.fill.background()
    eh_p = e_hdr.text_frame.paragraphs[0]
    eh_p.text = "OFFICIAL EVENT EMBLEM & LOGO CONCEPT"
    eh_p.font.bold = True
    eh_p.font.size = Pt(11)
    eh_p.font.color.rgb = HARVEST_GOLD
    eh_p.font.name = "Arial"

    tb_e = s12.shapes.add_textbox(Inches(1.0), Inches(4.05), Inches(5.25), Inches(2.65))
    etf = tb_e.text_frame
    etf.word_wrap = True
    
    e_items = [
        ("Emblem Structure:", "A circular crest featuring a stylized green leaf merging seamlessly with an ocean wave, framing a minimalist golden chef’s toque at the center."),
        ("Symbolic Meaning:", "Unifies environmental coastal conservation (leaf & wave) with regional gastronomic innovation (chef's toque)."),
        ("Typography Hierarchy:", "High-impact serif typeface for 'PECS 2026' paired with a modern sans-serif subtext for regional clarity.")
    ]
    for i, (l, v) in enumerate(e_items):
        p = etf.paragraphs[0] if i == 0 else etf.add_paragraph()
        p.text = l
        p.font.bold = True
        p.font.size = Pt(11)
        p.font.color.rgb = DARK_GREEN
        p.font.name = "Arial"
        if i > 0:
            p.space_before = Pt(6)
        pv = etf.add_paragraph()
        pv.text = v
        pv.font.size = Pt(10)
        pv.font.color.rgb = TEXT_DARK
        pv.font.name = "Arial"

    create_card(s12, Inches(6.88), Inches(3.45), Inches(5.65), Inches(3.35))
    s_hdr = s12.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(6.88), Inches(3.45), Inches(5.65), Inches(0.5))
    s_hdr.fill.solid()
    s_hdr.fill.fore_color.rgb = DARK_GREEN
    s_hdr.line.fill.background()
    sh_p = s_hdr.text_frame.paragraphs[0]
    sh_p.text = "STAGE SETUP & DELEGATE COLLATERAL"
    sh_p.font.bold = True
    sh_p.font.size = Pt(11)
    sh_p.font.color.rgb = HARVEST_GOLD
    sh_p.font.name = "Arial"

    tb_stg = s12.shapes.add_textbox(Inches(7.08), Inches(4.05), Inches(5.25), Inches(2.65))
    stgtf = tb_stg.text_frame
    stgtf.word_wrap = True

    stg_items = [
        ("Gymnasium Stage Setup:", "16:9 LED wall backdrop in Coastal Forest Green with Harvest Gold accents; sleek acrylic podium with official emblem, flanked by native potted palms."),
        ("Badges & Lanyards:", "Dark green lanyards paired with color-coded ID badges (Green for Delegates, Gold for Committee, Blue for Keynote Speakers/VIPs)."),
        ("Eco Conference Tote Kits:", "Canvas bags containing branded notepads, pens, program booklets, and QR access cards for digital verified certificates.")
    ]
    for i, (l, v) in enumerate(stg_items):
        p = stgtf.paragraphs[0] if i == 0 else stgtf.add_paragraph()
        p.text = l
        p.font.bold = True
        p.font.size = Pt(11)
        p.font.color.rgb = OCEAN_AQUA
        p.font.name = "Arial"
        if i > 0:
            p.space_before = Pt(6)
        pv = stgtf.add_paragraph()
        pv.text = v
        pv.font.size = Pt(10)
        pv.font.color.rgb = TEXT_DARK
        pv.font.name = "Arial"

    # ==========================================
    # SLIDE 13: LETTER 1 - GYMNASIUM RESERVATION
    # ==========================================
    s13 = prs.slides.add_slide(blank_slide_layout)
    set_slide_bg(s13)
    apply_slide_transition(s13, "push")
    add_header(s13, "Letter 1: Gymnasium Reservation & Facility Access", "Section 12 (Part 1) • Administrative Communication Letters")

    create_card(s13, Inches(0.8), Inches(1.5), Inches(4.2), Inches(5.3))
    l1_hdr = s13.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.5), Inches(4.2), Inches(0.55))
    l1_hdr.fill.solid()
    l1_hdr.fill.fore_color.rgb = DARK_GREEN
    l1_hdr.line.fill.background()
    l1_hp = l1_hdr.text_frame.paragraphs[0]
    l1_hp.text = "TRANSMITTAL DETAILS"
    l1_hp.font.bold = True
    l1_hp.font.size = Pt(11)
    l1_hp.font.color.rgb = HARVEST_GOLD
    l1_hp.font.name = "Arial"

    tb_l1_m = s13.shapes.add_textbox(Inches(0.95), Inches(2.15), Inches(3.9), Inches(4.4))
    l1_mtf = tb_l1_m.text_frame
    l1_mtf.word_wrap = True

    l1_meta = [
        ("📅 Date of Transmittal:", "September 2, 2026"),
        ("📬 Recipient Office:", "The Office of the VPAA / Campus Facilities & Athletics Office, PHINMA-UPang"),
        ("📌 Subject Matter:", "Venue Reservation, Technical Support & Facility Access for PECS 2026"),
        ("✍️ Signatory Authority:", "Conference Director, PECS 2026 (CHTM BAM 205)")
    ]
    for i, (l, v) in enumerate(l1_meta):
        p = l1_mtf.paragraphs[0] if i == 0 else l1_mtf.add_paragraph()
        p.text = l
        p.font.bold = True
        p.font.size = Pt(11)
        p.font.color.rgb = DARK_GREEN
        p.font.name = "Arial"
        if i > 0:
            p.space_before = Pt(8)
        pv = l1_mtf.add_paragraph()
        pv.text = v
        pv.font.size = Pt(10.5)
        pv.font.color.rgb = TEXT_DARK
        pv.font.name = "Arial"
        pv.space_before = Pt(2)

    create_card(s13, Inches(5.2), Inches(1.5), Inches(7.333), Inches(5.3), bg_color=WHITE, border_color=CARD_BORDER)
    l1_body_hdr = s13.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(5.2), Inches(1.5), Inches(7.333), Inches(0.55))
    l1_body_hdr.fill.solid()
    l1_body_hdr.fill.fore_color.rgb = OCEAN_AQUA
    l1_body_hdr.line.fill.background()
    lbh_p = l1_body_hdr.text_frame.paragraphs[0]
    lbh_p.text = "OFFICIAL TRANSMITTAL TEXT"
    lbh_p.font.bold = True
    lbh_p.font.size = Pt(11)
    lbh_p.font.color.rgb = WHITE
    lbh_p.font.name = "Arial"

    tb_l1_b = s13.shapes.add_textbox(Inches(5.4), Inches(2.15), Inches(6.9), Inches(4.5))
    l1_btf = tb_l1_b.text_frame
    l1_btf.word_wrap = True

    p = l1_btf.paragraphs[0]
    p.text = "Dear Administration,"
    p.font.bold = True
    p.font.size = Pt(11)
    p.font.color.rgb = DARK_GREEN
    p.font.name = "Georgia"

    p_body1 = l1_btf.add_paragraph()
    p_body1.text = "The CHTM BAM 205 class respectfully requests permission to reserve and utilize the PHINMA-University of Pangasinan Gymnasium for PECS 2026 on September 17, 2026, from 7:00 AM to 5:00 PM (inclusive of setup, operational hours, and teardown)."
    p_body1.font.size = Pt(10)
    p_body1.font.color.rgb = TEXT_DARK
    p_body1.font.name = "Arial"
    p_body1.space_before = Pt(4)

    p_body2 = l1_btf.add_paragraph()
    p_body2.text = "We also humbly request after-hours setup authorization for Sept 16, 2026 (until 7:00 PM) and technical support for:"
    p_body2.font.size = Pt(10)
    p_body2.font.bold = True
    p_body2.font.color.rgb = DARK_GREEN
    p_body2.font.name = "Arial"
    p_body2.space_before = Pt(4)

    items_support = [
        "1. Main stage projector/LED screen integration with HDMI connectivity.",
        "2. Sound system setup: 1 podium mic, 4 wireless handheld mics, and 2 lapel mics.",
        "3. Stage podium, panel chairs, and 500 spectator chairs arranged theater-style.",
        "4. Full facility ventilation, cooling, and lighting support during operational hours."
    ]
    for item in items_support:
        pi = l1_btf.add_paragraph()
        pi.text = item
        pi.font.size = Pt(9.5)
        pi.font.color.rgb = TEXT_DARK
        pi.font.name = "Arial"
        pi.space_before = Pt(2)

    p_close = l1_btf.add_paragraph()
    p_close.text = "Respectfully yours,\n[Your Name] — Conference Director, PECS 2026"
    p_close.font.bold = True
    p_close.font.size = Pt(10)
    p_close.font.color.rgb = DARK_GREEN
    p_close.font.name = "Arial"
    p_close.space_before = Pt(6)

    # ==========================================
    # SLIDE 14: LETTER 2 - KEYNOTE SPEAKER INVITATION
    # ==========================================
    s14 = prs.slides.add_slide(blank_slide_layout)
    set_slide_bg(s14)
    apply_slide_transition(s14, "push")
    add_header(s14, "Letter 2: Official Keynote Speaker Invitation", "Section 12 (Part 2) • Administrative Communication Letters")

    create_card(s14, Inches(0.8), Inches(1.5), Inches(4.2), Inches(5.3))
    l2_hdr = s14.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.5), Inches(4.2), Inches(0.55))
    l2_hdr.fill.solid()
    l2_hdr.fill.fore_color.rgb = DARK_GREEN
    l2_hdr.line.fill.background()
    l2_hp = l2_hdr.text_frame.paragraphs[0]
    l2_hp.text = "INVITATION PARAMETERS"
    l2_hp.font.bold = True
    l2_hp.font.size = Pt(11)
    l2_hp.font.color.rgb = HARVEST_GOLD
    l2_hp.font.name = "Arial"

    tb_l2_m = s14.shapes.add_textbox(Inches(0.95), Inches(2.15), Inches(3.9), Inches(4.4))
    l2_mtf = tb_l2_m.text_frame
    l2_mtf.word_wrap = True

    l2_meta = [
        ("📅 Date of Invitation:", "September 2, 2026"),
        ("📬 Target Honoree:", "[Speaker's Name / Title], [Organization Name]"),
        ("⏱️ Session Allocation:", "45-Minute Plenary Presentation + 15-Minute Open Floor Q&A"),
        ("🎁 VIP Provision:", "Private holding lounge, plated catering, formal honoraria, parking pass & local token basket"),
        ("✍️ Signatory Authority:", "Speaker Relations Lead, PECS 2026")
    ]
    for i, (l, v) in enumerate(l2_meta):
        p = l2_mtf.paragraphs[0] if i == 0 else l2_mtf.add_paragraph()
        p.text = l
        p.font.bold = True
        p.font.size = Pt(11)
        p.font.color.rgb = DARK_GREEN
        p.font.name = "Arial"
        if i > 0:
            p.space_before = Pt(6)
        pv = l2_mtf.add_paragraph()
        pv.text = v
        pv.font.size = Pt(10)
        pv.font.color.rgb = TEXT_DARK
        pv.font.name = "Arial"
        pv.space_before = Pt(2)

    create_card(s14, Inches(5.2), Inches(1.5), Inches(7.333), Inches(5.3), bg_color=WHITE, border_color=CARD_BORDER)
    l2_body_hdr = s14.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(5.2), Inches(1.5), Inches(7.333), Inches(0.55))
    l2_body_hdr.fill.solid()
    l2_body_hdr.fill.fore_color.rgb = OCEAN_AQUA
    l2_body_hdr.line.fill.background()
    lbh2_p = l2_body_hdr.text_frame.paragraphs[0]
    lbh2_p.text = "OFFICIAL INVITATION LETTER TEXT"
    lbh2_p.font.bold = True
    lbh2_p.font.size = Pt(11)
    lbh2_p.font.color.rgb = WHITE
    lbh2_p.font.name = "Arial"

    tb_l2_b = s14.shapes.add_textbox(Inches(5.4), Inches(2.15), Inches(6.9), Inches(4.5))
    l2_btf = tb_l2_b.text_frame
    l2_btf.word_wrap = True

    p = l2_btf.paragraphs[0]
    p.text = "Dear [Speaker's Name],"
    p.font.bold = True
    p.font.size = Pt(11)
    p.font.color.rgb = DARK_GREEN
    p.font.name = "Georgia"

    p_body1 = l2_btf.add_paragraph()
    p_body1.text = "Greetings from PHINMA University of Pangasinan!\n\nThe College of Hospitality and Tourism Management cordially invites you to serve as a Keynote Speaker for the Pangasinan Eco-Haven & Culinary Showcase 2026 (PECS 2026) Conference on Thursday, September 17, 2026, at the PHINMA-UPang Gymnasium."
    p_body1.font.size = Pt(10)
    p_body1.font.color.rgb = TEXT_DARK
    p_body1.font.name = "Arial"
    p_body1.space_before = Pt(4)

    p_body2 = l2_btf.add_paragraph()
    p_body2.text = "Our conference theme, “Sustaining the Coastline: Strategies for Eco-Tourism Development and Heritage Culinary Preservation in Pangasinan,” aims to equip 500+ undergraduate delegates with actionable strategies for sustainable hospitality growth. Given your distinguished expertise, we would be honored to have you deliver a 45-minute presentation followed by a 15-minute open Q&A session."
    p_body2.font.size = Pt(10)
    p_body2.font.color.rgb = TEXT_DARK
    p_body2.font.name = "Arial"
    p_body2.space_before = Pt(6)

    p_body3 = l2_btf.add_paragraph()
    p_body3.text = "Dedicated VIP hospitality, honorarium, and campus parking will be provided. We look forward to confirming your participation.\n\nWarm regards,\n[Your Name] — Speaker Relations Lead, PECS 2026"
    p_body3.font.size = Pt(10)
    p_body3.font.color.rgb = TEXT_DARK
    p_body3.font.name = "Arial"
    p_body3.space_before = Pt(6)

    # ==========================================
    # SLIDE 15: LETTER 3 - STUDENT CLASS EXCUSE
    # ==========================================
    s15 = prs.slides.add_slide(blank_slide_layout)
    set_slide_bg(s15)
    apply_slide_transition(s15, "push")
    add_header(s15, "Letter 3: Student Class Excuse Authorization", "Section 12 (Part 3) • Administrative Communication Letters")

    create_card(s15, Inches(0.8), Inches(1.5), Inches(4.2), Inches(5.3))
    l3_hdr = s15.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.5), Inches(4.2), Inches(0.55))
    l3_hdr.fill.solid()
    l3_hdr.fill.fore_color.rgb = DARK_GREEN
    l3_hdr.line.fill.background()
    l3_hp = l3_hdr.text_frame.paragraphs[0]
    l3_hp.text = "EXCUSAL SPECIFICATIONS"
    l3_hp.font.bold = True
    l3_hp.font.size = Pt(11)
    l3_hp.font.color.rgb = HARVEST_GOLD
    l3_hp.font.name = "Arial"

    tb_l3_m = s15.shapes.add_textbox(Inches(0.95), Inches(2.15), Inches(3.9), Inches(4.4))
    l3_mtf = tb_l3_m.text_frame
    l3_mtf.word_wrap = True

    l3_meta = [
        ("📅 Date of Notice:", "September 14, 2026"),
        ("📬 Recipients:", "CHTM Faculty Members and Department Chairs"),
        ("👥 Covered Cohort:", "All participating BSHM & BSTM student organizers and registered delegates"),
        ("⏰ Excusal Duration:", "Thursday, September 17, 2026 (8:00 AM – 4:00 PM)"),
        ("⚖️ Academic Policy:", "Students remain accountable for all missed lectures and course deliverables.")
    ]
    for i, (l, v) in enumerate(l3_meta):
        p = l3_mtf.paragraphs[0] if i == 0 else l3_mtf.add_paragraph()
        p.text = l
        p.font.bold = True
        p.font.size = Pt(11)
        p.font.color.rgb = DARK_GREEN
        p.font.name = "Arial"
        if i > 0:
            p.space_before = Pt(6)
        pv = l3_mtf.add_paragraph()
        pv.text = v
        pv.font.size = Pt(10)
        pv.font.color.rgb = TEXT_DARK
        pv.font.name = "Arial"
        pv.space_before = Pt(2)

    create_card(s15, Inches(5.2), Inches(1.5), Inches(7.333), Inches(5.3), bg_color=WHITE, border_color=CARD_BORDER)
    l3_body_hdr = s15.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(5.2), Inches(1.5), Inches(7.333), Inches(0.55))
    l3_body_hdr.fill.solid()
    l3_body_hdr.fill.fore_color.rgb = OCEAN_AQUA
    l3_body_hdr.line.fill.background()
    lbh3_p = l3_body_hdr.text_frame.paragraphs[0]
    lbh3_p.text = "OFFICIAL FACULTY TRANSMITTAL TEXT"
    lbh3_p.font.bold = True
    lbh3_p.font.size = Pt(11)
    lbh3_p.font.color.rgb = WHITE
    lbh3_p.font.name = "Arial"

    tb_l3_b = s15.shapes.add_textbox(Inches(5.4), Inches(2.15), Inches(6.9), Inches(4.5))
    l3_btf = tb_l3_b.text_frame
    l3_btf.word_wrap = True

    p = l3_btf.paragraphs[0]
    p.text = "Dear Professors,"
    p.font.bold = True
    p.font.size = Pt(11)
    p.font.color.rgb = DARK_GREEN
    p.font.name = "Georgia"

    p_body1 = l3_btf.add_paragraph()
    p_body1.text = "The organizing committee of PECS 2026 respectfully requests class excuse for participating BSHM and BSTM student organizers and registered delegates on Thursday, September 17, 2026, from 8:00 AM to 4:00 PM."
    p_body1.font.size = Pt(10.5)
    p_body1.font.color.rgb = TEXT_DARK
    p_body1.font.name = "Arial"
    p_body1.space_before = Pt(6)

    p_body2 = l3_btf.add_paragraph()
    p_body2.text = "These students will be actively managing conference logistics, main-stage operations, secretariat desks, and participating in session discussions at the PHINMA-UPang Gymnasium to fulfill BAM 205 course requirements. Participating students remain accountable for all academic responsibilities and missed lectures."
    p_body2.font.size = Pt(10.5)
    p_body2.font.color.rgb = TEXT_DARK
    p_body2.font.name = "Arial"
    p_body2.space_before = Pt(8)

    p_body3 = l3_btf.add_paragraph()
    p_body3.text = "Thank you for your cooperation and support.\n\nSincerely,\n[Your Name] — Conference Director, PECS 2026"
    p_body3.font.size = Pt(10.5)
    p_body3.font.color.rgb = TEXT_DARK
    p_body3.font.name = "Arial"
    p_body3.space_before = Pt(8)

    # ==========================================
    # SLIDE 16: FORM 4 - PARENT CONSENT
    # ==========================================
    s16 = prs.slides.add_slide(blank_slide_layout)
    set_slide_bg(s16)
    apply_slide_transition(s16, "push")
    add_header(s16, "Form 4: Parent’s / Guardian’s Official Consent Form", "Section 12 (Part 4) • Risk Management & Academic Compliance")

    create_card(s16, Inches(0.8), Inches(1.45), Inches(11.733), Inches(5.4), bg_color=WHITE, border_color=CARD_BORDER)
    
    f_hdr = s16.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.45), Inches(11.733), Inches(0.75))
    f_hdr.fill.solid()
    f_hdr.fill.fore_color.rgb = DARK_GREEN
    f_hdr.line.fill.background()
    f_hp = f_hdr.text_frame.paragraphs[0]
    f_hp.text = "PHINMA UNIVERSITY OF PANGASINAN • COLLEGE OF HOSPITALITY AND TOURISM MANAGEMENT"
    f_hp.font.bold = True
    f_hp.font.size = Pt(11)
    f_hp.font.color.rgb = HARVEST_GOLD
    f_hp.font.name = "Arial"
    f_hp.alignment = PP_ALIGN.CENTER

    f_hp2 = f_hdr.text_frame.add_paragraph()
    f_hp2.text = "PARENT’S / GUARDIAN’S OFFICIAL CONSENT FORM"
    f_hp2.font.bold = True
    f_hp2.font.size = Pt(12)
    f_hp2.font.color.rgb = WHITE
    f_hp2.font.name = "Georgia"
    f_hp2.alignment = PP_ALIGN.CENTER

    tb_form = s16.shapes.add_textbox(Inches(1.1), Inches(2.35), Inches(11.133), Inches(4.3))
    ftf = tb_form.text_frame
    ftf.word_wrap = True

    p1 = ftf.paragraphs[0]
    p1.text = "I, _______________________________________, parent/guardian of _______________________________________,\nstudent of BS [HM / TM]  Year & Section: ____________, hereby grant permission for my child to attend and participate in the official campus conference:"
    p1.font.size = Pt(11)
    p1.font.color.rgb = TEXT_DARK
    p1.font.name = "Arial"

    bullet_items = [
        ("• Event Name:", "PECS 2026: Regional Conference on Sustainable Tourism & Culinary Innovation"),
        ("• Date & Operational Time:", "September 17, 2026 | 7:00 AM – 5:00 PM (Including ingress, plenary sessions, and egress)"),
        ("• Official Venue:", "PHINMA-University of Pangasinan Gymnasium, Arellano Street, Dagupan City")
    ]
    for b_label, b_val in bullet_items:
        p = ftf.add_paragraph()
        p.text = f"{b_label} {b_val}"
        p.font.size = Pt(10.5)
        p.font.color.rgb = DARK_GREEN
        p.font.bold = True
        p.font.name = "Arial"
        p.space_before = Pt(4)

    p_ack = ftf.add_paragraph()
    p_ack.text = "I understand that this event fulfills academic criteria for BAM 205 (MICE) and that university safety and hall protocols will be strictly maintained by faculty supervisors."
    p_ack.font.size = Pt(10.5)
    p_ack.font.italic = True
    p_ack.font.color.rgb = TEXT_DARK
    p_ack.font.name = "Georgia"
    p_ack.space_before = Pt(8)

    p_sig = ftf.add_paragraph()
    p_sig.text = "\n__________________________________________________               Date: ________________________\nParent / Guardian Signature over Printed Name"
    p_sig.font.size = Pt(11)
    p_sig.font.bold = True
    p_sig.font.color.rgb = TEXT_DARK
    p_sig.font.name = "Arial"
    p_sig.space_before = Pt(10)

    # ==========================================
    # SLIDE 17: CONCLUSION (Dark Theme)
    # ==========================================
    s17 = prs.slides.add_slide(blank_slide_layout)
    set_slide_bg(s17, DEEP_BG)
    apply_slide_transition(s17, "fade")

    ab = s17.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(0.3), Inches(7.5))
    ab.fill.solid()
    ab.fill.fore_color.rgb = HARVEST_GOLD
    ab.line.fill.background()

    ct_tag = s17.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.2), Inches(0.8), Inches(3.6), Inches(0.45))
    ct_tag.fill.solid()
    ct_tag.fill.fore_color.rgb = OCEAN_AQUA
    ct_tag.line.fill.background()
    ct_p = ct_tag.text_frame.paragraphs[0]
    ct_p.alignment = PP_ALIGN.CENTER
    ct_p.text = "CONCLUSION & NEXT STEPS"
    ct_p.font.size = Pt(11)
    ct_p.font.bold = True
    ct_p.font.color.rgb = WHITE
    ct_p.font.name = "Arial"

    t_end = s17.shapes.add_textbox(Inches(1.2), Inches(1.4), Inches(11.0), Inches(1.2))
    etf2 = t_end.text_frame
    etf2.word_wrap = True
    ep1 = etf2.paragraphs[0]
    ep1.text = "Ready to Deliver an Exemplary MICE Experience"
    ep1.font.size = Pt(28)
    ep1.font.bold = True
    ep1.font.color.rgb = HARVEST_GOLD
    ep1.font.name = "Georgia"

    ep2 = etf2.add_paragraph()
    ep2.text = "PECS 2026 bridges academic rigor with regional hospitality excellence, providing BAM 205 students with hands-on event execution mastery."
    ep2.font.size = Pt(14)
    ep2.font.color.rgb = SOFT_SAND
    ep2.font.name = "Georgia"
    ep2.space_before = Pt(4)

    concl_data = [
        ("🏛️ Institutional Impact", "Positions PHINMA University of Pangasinan CHTM as a regional center of excellence for eco-tourism research and culinary heritage preservation."),
        ("📈 Student Learning Outcomes", "Demonstrates end-to-end conference management competencies: finance, logistics, AV production, speaker hospitality, and secretariat workflows."),
        ("🤝 Community & Industry Link", "Creates actionable partnerships between academia, local government units (LGUs), municipal tourism officers, and regional hospitality stakeholders.")
    ]

    for i, (head, text) in enumerate(concl_data):
        x = Inches(1.2 + i * 3.7)
        c = s17.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, Inches(2.9), Inches(3.4), Inches(3.0))
        c.fill.solid()
        c.fill.fore_color.rgb = DARK_GREEN
        c.line.color.rgb = HARVEST_GOLD
        c.line.width = Pt(1)
        
        tb = s17.shapes.add_textbox(x + Inches(0.15), Inches(3.1), Inches(3.1), Inches(2.6))
        tf = tb.text_frame
        tf.word_wrap = True
        
        p = tf.paragraphs[0]
        p.text = head
        p.font.size = Pt(13)
        p.font.bold = True
        p.font.color.rgb = HARVEST_GOLD
        p.font.name = "Georgia"
        
        p_body = tf.add_paragraph()
        p_body.text = text
        p_body.font.size = Pt(10.5)
        p_body.font.color.rgb = SOFT_SAND
        p_body.font.name = "Arial"
        p_body.space_before = Pt(6)

    c_foot = s17.shapes.add_textbox(Inches(1.2), Inches(6.2), Inches(11.0), Inches(0.6))
    cf_tf = c_foot.text_frame
    cfp = cf_tf.paragraphs[0]
    cfp.text = "Thank you! We welcome questions, recommendations, and formal approval."
    cfp.font.size = Pt(13)
    cfp.font.bold = True
    cfp.font.color.rgb = WHITE
    cfp.font.name = "Georgia"

    # Save presentation
    paths = [
        r"C:\Users\Admin\.gemini\antigravity\scratch\PECS_2026_Animated_Presentation.pptx",
        r"C:\Users\Admin\Downloads\PECS_2026_Animated_Presentation.pptx",
        r"C:\Users\Admin\Documents\PECS_2026_Animated_Presentation.pptx",
        r"C:\Users\Admin\Downloads\PECS_2026_Full_Proposal_with_Letters.pptx",
        r"C:\Users\Admin\.gemini\antigravity\scratch\PECS_2026_Full_Proposal_with_Letters.pptx"
    ]
    for p_out in paths:
        try:
            os.makedirs(os.path.dirname(p_out), exist_ok=True)
            prs.save(p_out)
            print(f"Animated PPTX successfully saved to: {p_out}")
        except Exception as e:
            print(f"Notice: Could not write to {p_out}: {e}")

if __name__ == "__main__":
    create_presentation()
