"""
build_presentation.py
Generates the consolidated, ultra-professional 7-slide PowerPoint presentation (.pptx)
for Aashi Srivastava's Portfolio: 'The Lucknow Gazette & Archive'.
Focuses on high-impact visual storytelling, candidate identity (Student & Tech Learner,
Aspiring Data Analyst), 4-year Java foundation, active Python learning, and all 5 Android Nova live photos.
"""

import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# ==============================================================================
# PRESENTATION INITIALIZATION & THEME CONFIGURATION
# ==============================================================================
prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

# Color Palette Constants (Vintage Editorial Meets Modern Luxury)
COLOR_DARK_BG        = RGBColor(28, 21, 16)     # #1C1510 - Rich Espresso
COLOR_DARK_CARD      = RGBColor(38, 29, 23)     # #261D17 - Dark Espresso Card
COLOR_PARCHMENT      = RGBColor(250, 246, 238)  # #FAF6EE - Warm Antique Parchment
COLOR_CARD_BG        = RGBColor(245, 239, 229)  # #F5EFE5 - Slightly darker card parchment
COLOR_BURGUNDY       = RGBColor(107, 23, 36)    # #6B1724 - Royal Deep Burgundy
COLOR_BURGUNDY_LIGHT = RGBColor(243, 230, 232)  # #F3E6E8 - Light burgundy tint
COLOR_OLIVE          = RGBColor(74, 90, 63)     # #4A5A3F - Antique Olive / Moss
COLOR_OLIVE_LIGHT    = RGBColor(234, 238, 230)  # #EAEEE6 - Light olive tint
COLOR_GOLD           = RGBColor(200, 155, 60)   # #C89B3C - Antique Gold
COLOR_TEXT_DARK      = RGBColor(35, 24, 18)     # #231812 - Deep Espresso Ink
COLOR_TEXT_MUTED     = RGBColor(115, 102, 92)   # #73665C - Muted Charcoal / Brown
COLOR_TEXT_LIGHT     = RGBColor(246, 241, 232)  # #F6F1E8 - Parchment white for dark slides
COLOR_BORDER         = RGBColor(213, 199, 178)  # #D5C7B2 - Subtle Rule Border
COLOR_WHITE          = RGBColor(255, 255, 255)

FONT_SERIF = "Georgia"
FONT_SANS  = "Calibri"
FONT_MONO  = "Consolas"

blank_layout = prs.slide_layouts[6]

# ==============================================================================
# HELPER FUNCTIONS FOR VISUAL CARDS & ACCENTS
# ==============================================================================
def set_slide_background(slide, color):
    bg_shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    bg_shape.fill.solid()
    bg_shape.fill.fore_color.rgb = color
    bg_shape.line.fill.background()
    return bg_shape

def add_header(slide, chapter_tag, title_text, subtitle_text=""):
    # Header container
    header_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.733), Inches(1.1))
    tf = header_box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
    
    # Tag
    p_tag = tf.paragraphs[0]
    p_tag.text = chapter_tag.upper()
    p_tag.font.name = FONT_MONO
    p_tag.font.size = Pt(9.5)
    p_tag.font.bold = True
    p_tag.font.color.rgb = COLOR_BURGUNDY
    p_tag.space_after = Pt(2)
    
    # Title
    p_title = tf.add_paragraph()
    p_title.text = title_text
    p_title.font.name = FONT_SERIF
    p_title.font.size = Pt(22)
    p_title.font.bold = True
    p_title.font.color.rgb = COLOR_TEXT_DARK
    p_title.space_after = Pt(2)
    
    # Subtitle
    if subtitle_text:
        p_sub = tf.add_paragraph()
        p_sub.text = subtitle_text
        p_sub.font.name = FONT_SANS
        p_sub.font.size = Pt(11)
        p_sub.font.italic = True
        p_sub.font.color.rgb = COLOR_TEXT_MUTED
    
    # Accent Rule Line
    rule = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.55), Inches(11.733), Inches(0.02))
    rule.fill.solid()
    rule.fill.fore_color.rgb = COLOR_BORDER
    rule.line.fill.background()

def add_card_box(slide, left, top, width, height, bg_color=COLOR_CARD_BG, border_color=COLOR_BORDER, border_width=Pt(1.2)):
    card = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    card.fill.solid()
    card.fill.fore_color.rgb = bg_color
    if border_color:
        card.line.color.rgb = border_color
        card.line.width = border_width
    else:
        card.line.fill.background()
    return card

def add_accent_card(slide, left, top, width, height, accent_color, bg_color=COLOR_CARD_BG, border_color=COLOR_BORDER):
    # Base Card
    card = add_card_box(slide, left, top, width, height, bg_color=bg_color, border_color=border_color)
    # Top Accent Strip
    strip = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, Inches(0.07))
    strip.fill.solid()
    strip.fill.fore_color.rgb = accent_color
    strip.line.fill.background()
    return card

def add_image_card(slide, left, top, width, height, img_path, caption="", bg_color=COLOR_WHITE, border_color=COLOR_BORDER):
    # Outer frame
    frame = add_card_box(slide, left, top, width, height, bg_color=bg_color, border_color=border_color)
    margin = Inches(0.08)
    img_left = left + margin
    img_top = top + margin
    cap_h = Inches(0.32) if caption else 0
    img_w = width - (margin * 2)
    img_h = height - (margin * 2) - cap_h
    
    if os.path.exists(img_path):
        slide.shapes.add_picture(img_path, img_left, img_top, width=img_w, height=img_h)
    
    if caption:
        cap_box = slide.shapes.add_textbox(left + margin, top + height - Inches(0.34), width - (margin * 2), Inches(0.28))
        tf = cap_box.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
        p = tf.paragraphs[0]
        p.text = caption
        p.font.name = FONT_MONO
        p.font.size = Pt(9)
        p.font.bold = True
        p.font.color.rgb = COLOR_BURGUNDY
        p.alignment = PP_ALIGN.CENTER

def add_badge(slide, left, top, width, height, text, bg_color, text_color, font_size=Pt(9)):
    b = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    b.fill.solid()
    b.fill.fore_color.rgb = bg_color
    b.line.fill.background()
    tf = b.text_frame
    tf.word_wrap = False
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.text = text
    p.font.name = FONT_MONO
    p.font.size = font_size
    p.font.bold = True
    p.font.color.rgb = text_color
    p.alignment = PP_ALIGN.CENTER
    return b

def add_metric_tile(slide, left, top, width, height, value_text, label_text, bg_color=COLOR_CARD_BG, value_color=COLOR_BURGUNDY):
    card = add_card_box(slide, left, top, width, height, bg_color=bg_color, border_color=COLOR_BORDER)
    tb = slide.shapes.add_textbox(left + Inches(0.05), top + Inches(0.08), width - Inches(0.1), height - Inches(0.16))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    p1 = tf.paragraphs[0]
    p1.text = value_text
    p1.font.name = FONT_SERIF
    p1.font.size = Pt(17)
    p1.font.bold = True
    p1.font.color.rgb = value_color
    p1.alignment = PP_ALIGN.CENTER
    
    p2 = tf.add_paragraph()
    p2.text = label_text
    p2.font.name = FONT_MONO
    p2.font.size = Pt(8.5)
    p2.font.color.rgb = COLOR_TEXT_MUTED
    p2.alignment = PP_ALIGN.CENTER
    p2.space_before = Pt(1)
    return card

top_pos = Inches(1.8)


# ==============================================================================
# SLIDE 1: EXECUTIVE COVER & PROFILE COMPENDIUM (DARK LUXURY ESPRESSO)
# ==============================================================================
slide1 = prs.slides.add_slide(blank_layout)
set_slide_background(slide1, COLOR_DARK_BG)

# Decorative double gold/burgundy borders
outer_border = slide1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(0.5), Inches(12.333), Inches(6.5))
outer_border.fill.background()
outer_border.line.color.rgb = COLOR_GOLD
outer_border.line.width = Pt(1.5)

inner_border = slide1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.6), Inches(0.6), Inches(12.133), Inches(6.3))
inner_border.fill.background()
inner_border.line.color.rgb = COLOR_BURGUNDY
inner_border.line.width = Pt(0.75)

# Broadside Masthead Header Ribbon
tb_top = slide1.shapes.add_textbox(Inches(1.0), Inches(0.8), Inches(11.333), Inches(0.35))
tf_top = tb_top.text_frame
tf_top.margin_left = tf_top.margin_right = tf_top.margin_top = tf_top.margin_bottom = 0
p = tf_top.paragraphs[0]
p.text = "THE OFFICIAL ARCHIVE & PORTFOLIO COMPENDIUM • VOL. XXVI • ISSUE 01"
p.font.name = FONT_MONO
p.font.size = Pt(10.5)
p.font.bold = True
p.font.color.rgb = COLOR_GOLD
p.alignment = PP_ALIGN.CENTER

# Headline Titles
tb_title = slide1.shapes.add_textbox(Inches(1.0), Inches(1.2), Inches(11.333), Inches(2.2))
tf_title = tb_title.text_frame
tf_title.word_wrap = True
tf_title.margin_left = tf_title.margin_right = tf_title.margin_top = tf_title.margin_bottom = 0

p1 = tf_title.paragraphs[0]
p1.text = "AASHI SRIVASTAVA"
p1.font.name = FONT_SERIF
p1.font.size = Pt(44)
p1.font.bold = True
p1.font.color.rgb = COLOR_TEXT_LIGHT
p1.alignment = PP_ALIGN.CENTER

p2 = tf_title.add_paragraph()
p2.text = "The Lucknow Gazette & Curated Engineering Dossier"
p2.font.name = FONT_SERIF
p2.font.size = Pt(20)
p2.font.bold = True
p2.font.color.rgb = COLOR_GOLD
p2.alignment = PP_ALIGN.CENTER
p2.space_before = Pt(3)

p3 = tf_title.add_paragraph()
p3.text = "Undergraduate Student & Tech Learner (AI & DS) • Aspiring Data Analyst"
p3.font.name = FONT_SANS
p3.font.size = Pt(13)
p3.font.bold = True
p3.font.color.rgb = RGBColor(195, 185, 175)
p3.alignment = PP_ALIGN.CENTER
p3.space_before = Pt(5)

# Center Profile & Credential Box
center_card = add_card_box(slide1, Inches(1.5), Inches(3.55), Inches(10.333), Inches(2.65), bg_color=COLOR_DARK_CARD, border_color=COLOR_GOLD)

# Portrait inside card
portrait_path = "assets/images/profile/portrait.jpg"
if os.path.exists(portrait_path):
    slide1.shapes.add_picture(portrait_path, Inches(1.8), Inches(3.75), width=Inches(1.75), height=Inches(2.25))

# Profile Narrative & Details
tb_meta = slide1.shapes.add_textbox(Inches(3.8), Inches(3.75), Inches(4.7), Inches(2.25))
tf_meta = tb_meta.text_frame
tf_meta.word_wrap = True
tf_meta.margin_left = tf_meta.margin_right = tf_meta.margin_top = tf_meta.margin_bottom = 0

pm1 = tf_meta.paragraphs[0]
pm1.text = "ACADEMIC & PROFESSIONAL STANDING"
pm1.font.name = FONT_MONO
pm1.font.size = Pt(10)
pm1.font.bold = True
pm1.font.color.rgb = COLOR_GOLD

pm2 = tf_meta.add_paragraph()
pm2.text = ("• B.Tech in CSE (Artificial Intelligence & Data Science)\n"
            "• Shri Ramswaroop Memorial University (SRMU), Lucknow\n"
            "• Academic Standing: 8.29 CGPA • Class of 2029\n"
            "• 4-Year Java Foundations (Grades 9–12) + Active Python Stack")
pm2.font.name = FONT_SANS
pm2.font.size = Pt(11)
pm2.font.color.rgb = COLOR_TEXT_LIGHT
pm2.space_before = Pt(4)

pm3 = tf_meta.add_paragraph()
pm3.text = "Live Portfolio: https://aashi230.github.io/portfolio/"
pm3.font.name = FONT_MONO
pm3.font.size = Pt(9.5)
pm3.font.color.rgb = RGBColor(165, 205, 165)
pm3.space_before = Pt(6)

# Right 4-Metric Tiles inside the center card
m_left = Inches(8.7)
m_top = Inches(3.75)
m_w = Inches(1.4)
m_h = Inches(1.05)
m_gap = Inches(0.12)

add_metric_tile(slide1, m_left, m_top, m_w, m_h, "8.29", "SRMU CGPA", bg_color=RGBColor(48, 38, 30), value_color=COLOR_GOLD)
add_metric_tile(slide1, m_left + m_w + m_gap, m_top, m_w, m_h, "4+", "PROJECTS", bg_color=RGBColor(48, 38, 30), value_color=COLOR_TEXT_LIGHT)
add_metric_tile(slide1, m_left, m_top + m_h + m_gap, m_w, m_h, "SIH '26", "HACKATHON", bg_color=RGBColor(48, 38, 30), value_color=COLOR_GOLD)
add_metric_tile(slide1, m_left + m_w + m_gap, m_top + m_h + m_gap, m_w, m_h, "HOST", "ANDRO NOVA", bg_color=RGBColor(48, 38, 30), value_color=COLOR_TEXT_LIGHT)

# Footer Colophon
tb_foot = slide1.shapes.add_textbox(Inches(1.0), Inches(6.4), Inches(11.333), Inches(0.35))
p_foot = tb_foot.text_frame.paragraphs[0]
p_foot.text = "Aesthetic Doctrine: Vintage Editorial Luxury Meets Computational Rigor • Designed & Engineered with Pride"
p_foot.font.name = FONT_MONO
p_foot.font.size = Pt(9.5)
p_foot.font.color.rgb = COLOR_GOLD
p_foot.alignment = PP_ALIGN.CENTER


# ==============================================================================
# SLIDE 2: CANDIDATE PROFILE & ACADEMIC DOSSIER (PARCHMENT)
# ==============================================================================
slide2 = prs.slides.add_slide(blank_layout)
set_slide_background(slide2, COLOR_PARCHMENT)
add_header(slide2, "CHAPTER 01 / CANDIDATE PROFILE", "The Academic Dossier & Scholarly Standing",
           "Verified undergraduate trajectory, dual engineering identity, and academic distinction.")

left_w2 = Inches(5.6)
add_accent_card(slide2, Inches(0.8), top_pos, left_w2, Inches(5.1), COLOR_BURGUNDY)

t_abt = slide2.shapes.add_textbox(Inches(1.05), top_pos + Inches(0.2), left_w2 - Inches(0.5), Inches(4.7))
f_abt = t_abt.text_frame
f_abt.word_wrap = True

p = f_abt.paragraphs[0]
p.text = "UNDERGRADUATE PROFILE"
p.font.name = FONT_MONO
p.font.size = Pt(10)
p.font.bold = True
p.font.color.rgb = COLOR_BURGUNDY

p = f_abt.add_paragraph()
p.text = "Rigorous Foundation in AI & Data Science"
p.font.name = FONT_SERIF
p.font.size = Pt(16)
p.font.bold = True
p.font.color.rgb = COLOR_TEXT_DARK
p.space_before = Pt(2)

p = f_abt.add_paragraph()
p.text = ("• Collegiate Distinction: Second-year undergraduate pursuing B.Tech in Computer Science Engineering (Artificial Intelligence & Data Science) at SRMU Lucknow, holding a stellar 8.29 CGPA.\n\n"
          "• Four-Year Java Heritage: Cultivated foundational programming excellence through 4 years of formal Java instruction from Class 9 to 12, cementing core object-oriented architecture, classes, and algorithmic problem-solving.\n\n"
          "• Active Python Stack: Actively expanding into modern Python workflows for data science, exploratory data analysis (EDA), data visualization, and applied machine learning.\n\n"
          "• Professional Target: Combining technical curiosity as a Student & Tech Learner with industry-ready aspirations for Aspiring Data Analyst roles.")
p.font.name = FONT_SANS
p.font.size = Pt(10.5)
p.font.color.rgb = COLOR_TEXT_MUTED
p.space_before = Pt(4)

# Academic Highlight Banner inside left card
add_card_box(slide2, Inches(1.05), top_pos + Inches(3.9), left_w2 - Inches(0.5), Inches(0.95), bg_color=COLOR_BURGUNDY_LIGHT, border_color=COLOR_BURGUNDY)
tb_hilight = slide2.shapes.add_textbox(Inches(1.15), top_pos + Inches(3.95), left_w2 - Inches(0.7), Inches(0.85))
tf_hi = tb_hilight.text_frame
tf_hi.word_wrap = True
p = tf_hi.paragraphs[0]
p.text = "CORE TRAJECTORY SUMMARY"
p.font.name = FONT_MONO
p.font.size = Pt(9)
p.font.bold = True
p.font.color.rgb = COLOR_BURGUNDY
p = tf_hi.add_paragraph()
p.text = "B.Tech CSE (AI & DS) • Class of 2029 • Lucknow, India • 8.29 CGPA • Java 9–12 • Python Active"
p.font.name = FONT_SANS
p.font.size = Pt(9.5)
p.font.bold = True
p.font.color.rgb = COLOR_TEXT_DARK
p.space_before = Pt(2)

# Right Column: Screenshot Card
add_image_card(slide2, Inches(6.7), top_pos, Inches(5.833), Inches(5.1),
               "presentation_assets/02_about_dossier.png",
               "ARCHIVAL DOSSIER: Verified Credentials, University Standing & Passports")


# ==============================================================================
# SLIDE 3: TECHNICAL ARSENAL & DATA ANALYTICS STACK (PARCHMENT)
# ==============================================================================
slide3 = prs.slides.add_slide(blank_layout)
set_slide_background(slide3, COLOR_PARCHMENT)
add_header(slide3, "CHAPTER 02 / TECHNICAL COMPETENCIES", "Technical Arsenal & Data Analytics Stack",
           "Structured multi-tier proficiencies across languages, analytical libraries, and engineering tools.")

col_w3 = Inches(3.7)
gap3 = Inches(0.316)
c1_left = Inches(0.8)
c2_left = c1_left + col_w3 + gap3
c3_left = c2_left + col_w3 + gap3

# Pillar 1: Programming Languages
add_accent_card(slide3, c1_left, top_pos, col_w3, Inches(5.1), COLOR_BURGUNDY)
tb_p1 = slide3.shapes.add_textbox(c1_left + Inches(0.2), top_pos + Inches(0.2), col_w3 - Inches(0.4), Inches(4.7))
tf_p1 = tb_p1.text_frame
tf_p1.word_wrap = True

p = tf_p1.paragraphs[0]
p.text = "01 / CORE LANGUAGES"
p.font.name = FONT_MONO
p.font.size = Pt(9.5)
p.font.bold = True
p.font.color.rgb = COLOR_BURGUNDY

p = tf_p1.add_paragraph()
p.text = "Code Foundations"
p.font.name = FONT_SERIF
p.font.size = Pt(16)
p.font.bold = True
p.font.color.rgb = COLOR_TEXT_DARK
p.space_before = Pt(2)

p = tf_p1.add_paragraph()
p.text = ("• Java (Class 9–12 Foundation):\n"
          "  4-year academic background establishing strong object-oriented principles, classes, methods, and algorithmic structure.\n\n"
          "• Python (Actively Learning):\n"
          "  Focused adoption for data analytics, scientific computation, data cleaning, and AI workflows.\n\n"
          "• C (Procedural Fundamentals):\n"
          "  Solid grounding in memory concepts, procedural logic, pointers, and data representations.\n\n"
          "• SQL (Database Querying):\n"
          "  Relational schema design, filtering, aggregate queries, and data extraction.")
p.font.name = FONT_SANS
p.font.size = Pt(10)
p.font.color.rgb = COLOR_TEXT_MUTED
p.space_before = Pt(4)

# Pillar 2: AI & Data Science Stack
add_accent_card(slide3, c2_left, top_pos, col_w3, Inches(5.1), COLOR_OLIVE)
tb_p2 = slide3.shapes.add_textbox(c2_left + Inches(0.2), top_pos + Inches(0.2), col_w3 - Inches(0.4), Inches(4.7))
tf_p2 = tb_p2.text_frame
tf_p2.word_wrap = True

p = tf_p2.paragraphs[0]
p.text = "02 / DATA & MACHINE LEARNING"
p.font.name = FONT_MONO
p.font.size = Pt(9.5)
p.font.bold = True
p.font.color.rgb = COLOR_OLIVE

p = tf_p2.add_paragraph()
p.text = "Analytical Toolkit"
p.font.name = FONT_SERIF
p.font.size = Pt(16)
p.font.bold = True
p.font.color.rgb = COLOR_TEXT_DARK
p.space_before = Pt(2)

p = tf_p2.add_paragraph()
p.text = ("• Exploratory Data Analysis (EDA):\n"
          "  Statistical summary extraction, data distributions, correlation analysis, and data cleansing.\n\n"
          "• Data Science Libraries:\n"
          "  Hands-on learning with Pandas (DataFrames) and NumPy (numerical matrices & vectors).\n\n"
          "• Data Visualization:\n"
          "  Matplotlib and Seaborn for publication-grade graphs, histograms, and scatter plots.\n\n"
          "• Generative AI & Prompting:\n"
          "  Be10x Certified in prompt engineering, AI tool integration, and accelerated workflows.")
p.font.name = FONT_SANS
p.font.size = Pt(10)
p.font.color.rgb = COLOR_TEXT_MUTED
p.space_before = Pt(4)

# Pillar 3: Tooling & Platform
add_accent_card(slide3, c3_left, top_pos, col_w3, Inches(5.1), COLOR_GOLD)
tb_p3 = slide3.shapes.add_textbox(c3_left + Inches(0.2), top_pos + Inches(0.2), col_w3 - Inches(0.4), Inches(4.7))
tf_p3 = tb_p3.text_frame
tf_p3.word_wrap = True

p = tf_p3.paragraphs[0]
p.text = "03 / TOOLING & PLATFORM"
p.font.name = FONT_MONO
p.font.size = Pt(9.5)
p.font.bold = True
p.font.color.rgb = COLOR_GOLD

p = tf_p3.add_paragraph()
p.text = "Engineering Workflow"
p.font.name = FONT_SERIF
p.font.size = Pt(16)
p.font.bold = True
p.font.color.rgb = COLOR_TEXT_DARK
p.space_before = Pt(2)

p = tf_p3.add_paragraph()
p.text = ("• Version Control:\n"
          "  Git and GitHub proficiency for branching, commits, release tagging, and GitHub Pages CI/CD.\n\n"
          "• Development Environments:\n"
          "  VS Code, Jupyter Notebooks, and Google Colab for interactive computational exploration.\n\n"
          "• Web Interface:\n"
          "  Semantic HTML5 and modern modular CSS3 (CSS Grid, Flexbox, responsive typography).\n\n"
          "• Command Shell:\n"
          "  Windows PowerShell, shell scripts, and reproducible archival documentation.")
p.font.name = FONT_SANS
p.font.size = Pt(10)
p.font.color.rgb = COLOR_TEXT_MUTED
p.space_before = Pt(4)


# ==============================================================================
# SLIDE 4: FLAGSHIP ENGINEERING PROJECTS (PARCHMENT)
# ==============================================================================
slide4 = prs.slides.add_slide(blank_layout)
set_slide_background(slide4, COLOR_PARCHMENT)
add_header(slide4, "CHAPTER 03 / PROJECT ARCHIVE", "Flagship Projects & Computational Solutions",
           "End-to-end engineering works merging algorithmic logic with real-world user utility.")

left_w4 = Inches(5.6)

# Project Card 1: ARS Tourist Guide
add_accent_card(slide4, Inches(0.8), top_pos, left_w4, Inches(2.45), COLOR_BURGUNDY)
tb_pj1 = slide4.shapes.add_textbox(Inches(1.05), top_pos + Inches(0.15), left_w4 - Inches(0.5), Inches(2.2))
tf_pj1 = tb_pj1.text_frame
tf_pj1.word_wrap = True

p = tf_pj1.paragraphs[0]
p.text = "SIH 2026 SELECTION • SMART TRAVEL & HERITAGE"
p.font.name = FONT_MONO
p.font.size = Pt(9)
p.font.bold = True
p.font.color.rgb = COLOR_BURGUNDY

p = tf_pj1.add_paragraph()
p.text = "ARS Tourist Guide & Heritage Explorer"
p.font.name = FONT_SERIF
p.font.size = Pt(15)
p.font.bold = True
p.font.color.rgb = COLOR_TEXT_DARK
p.space_before = Pt(2)

p = tf_pj1.add_paragraph()
p.text = ("• AI-assisted cultural navigation and dynamic itinerary planner engineered for domestic and global travelers.\n"
          "• Selected as premier Smart India Hackathon (SIH 2026) internal college innovation.\n"
          "• Features smart monument discovery, localized transit tips, and crowd-informed schedules.\n"
          "• Built through agile team collaboration and rapid hackathon sprint cycles.")
p.font.name = FONT_SANS
p.font.size = Pt(9.5)
p.font.color.rgb = COLOR_TEXT_MUTED
p.space_before = Pt(3)

# Project Card 2: AI Mental Health Journaling Platform
add_accent_card(slide4, Inches(0.8), top_pos + Inches(2.65), left_w4, Inches(2.45), COLOR_OLIVE)
tb_pj2 = slide4.shapes.add_textbox(Inches(1.05), top_pos + Inches(2.8), left_w4 - Inches(0.5), Inches(2.2))
tf_pj2 = tb_pj2.text_frame
tf_pj2.word_wrap = True

p = tf_pj2.paragraphs[0]
p.text = "NATURAL LANGUAGE PROCESSING • WELLNESS AI"
p.font.name = FONT_MONO
p.font.size = Pt(9)
p.font.bold = True
p.font.color.rgb = COLOR_OLIVE

p = tf_pj2.add_paragraph()
p.text = "AI Mental Health Journaling Companion"
p.font.name = FONT_SERIF
p.font.size = Pt(15)
p.font.bold = True
p.font.color.rgb = COLOR_TEXT_DARK
p.space_before = Pt(2)

p = tf_pj2.add_paragraph()
p.text = ("• Sentiment-aware reflective journaling application with emotional wellness tracking and mood trend detection.\n"
          "• Incorporates natural language processing to deliver empathetic wellness prompts and trend summaries.\n"
          "• Built with strict user privacy, secure local storage, and structured reflection history.\n"
          "• Demonstrates human-centered AI utility and empathetic interface design.")
p.font.name = FONT_SANS
p.font.size = Pt(9.5)
p.font.color.rgb = COLOR_TEXT_MUTED
p.space_before = Pt(3)

# Right Column: Screenshot Card
add_image_card(slide4, Inches(6.7), top_pos, Inches(5.833), Inches(5.1),
               "presentation_assets/03_projects_work.png",
               "PROJECT ARCHIVE: Interactive Modals, Tech Badges & Case Studies")


# ==============================================================================
# SLIDE 5: CAMPUS LEADERSHIP & PUBLIC SPEAKING (ALL 5 LIVE STAGE FRAMES)
# ==============================================================================
slide5 = prs.slides.add_slide(blank_layout)
set_slide_background(slide5, COLOR_PARCHMENT)
add_header(slide5, "CHAPTER 04 / LEADERSHIP & PUBLIC SPEAKING", "Android Nova 2.0 — Symposium Stage Host & Coordinator",
           "Master of ceremonies leading SRMU's premier technical symposium across all 5 live event frames.")

# Left Card: Executive Narrative & Milestone Details
left_nova_w = Inches(3.5)
add_accent_card(slide5, Inches(0.8), top_pos, left_nova_w, Inches(5.1), COLOR_BURGUNDY)
t_nv = slide5.shapes.add_textbox(Inches(1.0), top_pos + Inches(0.18), left_nova_w - Inches(0.4), Inches(4.7))
f_nv = t_nv.text_frame
f_nv.word_wrap = True

p = f_nv.paragraphs[0]
p.text = "SYMPOSIUM HOST & COORDINATOR"
p.font.name = FONT_MONO
p.font.size = Pt(9.5)
p.font.bold = True
p.font.color.rgb = COLOR_BURGUNDY

p = f_nv.add_paragraph()
p.text = "Stage Leadership & Poise"
p.font.name = FONT_SERIF
p.font.size = Pt(15)
p.font.bold = True
p.font.color.rgb = COLOR_TEXT_DARK
p.space_before = Pt(2)

p = f_nv.add_paragraph()
p.text = ("• Selected as stage anchor for Android Nova 2.0 before hundreds of attendees, faculty, and industry speakers.\n"
          "• Commanded the auditorium stage with poise, speaker introductions, and dignitary felicitations.\n"
          "• Orchestrated backstage logistics, synchronization across technical teams, and audio-visual flow.\n"
          "• Demonstrated high-stakes public speaking, rapid adaptability, and large-audience engagement.")
p.font.name = FONT_SANS
p.font.size = Pt(10)
p.font.color.rgb = COLOR_TEXT_MUTED
p.space_before = Pt(4)

p = f_nv.add_paragraph()
p.text = "EVENT DOSSIER & IMPACT"
p.font.name = FONT_MONO
p.font.size = Pt(9.5)
p.font.bold = True
p.font.color.rgb = COLOR_OLIVE
p.space_before = Pt(6)

p = f_nv.add_paragraph()
p.text = ("• Role: Stage Anchor & Event Host\n"
          "• Institution: SRMU Lucknow\n"
          "• Scope: Flagship University Symposium\n"
          "• Documentation: 5 Live Stage Frames Below")
p.font.name = FONT_SANS
p.font.size = Pt(9.5)
p.font.color.rgb = COLOR_TEXT_DARK
p.space_before = Pt(2)

# Right Area: Prominent 5 Live Photo Spread (No Certificate)
# Top Row: 3 Stage Frames across width 8.033 in
add_image_card(slide5, Inches(4.5), top_pos, Inches(2.6), Inches(2.45),
               "presentation_assets/nova_1.jpg", "Frame 01: Stage Anchoring")
add_image_card(slide5, Inches(7.22), top_pos, Inches(2.8), Inches(2.45),
               "presentation_assets/nova_2.jpg", "Frame 02: Keynote & Audience")
add_image_card(slide5, Inches(10.14), top_pos, Inches(2.393), Inches(2.45),
               "presentation_assets/nova_3.jpg", "Frame 03: Organizing Desk")

# Bottom Row: 2 Wide Stage Frames across width 8.033 in
add_image_card(slide5, Inches(4.5), top_pos + Inches(2.65), Inches(3.95), Inches(2.45),
               "presentation_assets/nova_4.jpg", "Frame 04: Dignitary Felicitation")
add_image_card(slide5, Inches(8.583), top_pos + Inches(2.65), Inches(3.95), Inches(2.45),
               "presentation_assets/nova_5.jpg", "Frame 05: Host Commemorative Portrait")


# ==============================================================================
# SLIDE 6: DIGITAL PLATFORM ARCHITECTURE & CURRICULUM VITAE (PARCHMENT)
# ==============================================================================
slide6 = prs.slides.add_slide(blank_layout)
set_slide_background(slide6, COLOR_PARCHMENT)
add_header(slide6, "CHAPTER 05 / DIGITAL ARCHITECTURE & CV", "The Portfolio Platform & Executive Curriculum Vitae",
           "Handcrafted vintage editorial luxury, full-screen interactive inspection, and verified resume.")

col_w6 = Inches(5.7)
gap6 = Inches(0.333)

# Column 1: Web Platform Architecture
add_image_card(slide6, Inches(0.8), top_pos, col_w6, Inches(3.2),
               "presentation_assets/01_hero_landing.png", "Landing Broadside Architecture & Metrics Ticker")

add_card_box(slide6, Inches(0.8), top_pos + Inches(3.35), col_w6, Inches(1.75))
tb_w1 = slide6.shapes.add_textbox(Inches(0.95), top_pos + Inches(3.45), col_w6 - Inches(0.3), Inches(1.55))
tf_w1 = tb_w1.text_frame
tf_w1.word_wrap = True

p = tf_w1.paragraphs[0]
p.text = "VINTAGE EDITORIAL LUXURY PHILOSOPHY"
p.font.name = FONT_MONO
p.font.size = Pt(9.5)
p.font.bold = True
p.font.color.rgb = COLOR_BURGUNDY

p = tf_w1.add_paragraph()
p.text = ("• Bespoke 19th-century broadside aesthetics crafted in semantic HTML5 & modular CSS3.\n"
          "• Dynamic full-screen image lightbox inspection and interactive project detail modals.\n"
          "• Live contact dispatch form powered by FormSubmit with instant inbox delivery.")
p.font.name = FONT_SANS
p.font.size = Pt(9.5)
p.font.color.rgb = COLOR_TEXT_MUTED
p.space_before = Pt(3)

# Column 2: Executive Curriculum Vitae
add_image_card(slide6, Inches(0.8) + col_w6 + gap6, top_pos, col_w6, Inches(3.2),
               "presentation_assets/04_resume_cv.png", "Curriculum Vitae & One-Click PDF Resume Download")

add_card_box(slide6, Inches(0.8) + col_w6 + gap6, top_pos + Inches(3.35), col_w6, Inches(1.75))
tb_w2 = slide6.shapes.add_textbox(Inches(0.95) + col_w6 + gap6, top_pos + Inches(3.45), col_w6 - Inches(0.3), Inches(1.55))
tf_w2 = tb_w2.text_frame
tf_w2.word_wrap = True

p = tf_w2.paragraphs[0]
p.text = "ARCHIVAL CV & DIRECT PDF ACCESS"
p.font.name = FONT_MONO
p.font.size = Pt(9.5)
p.font.bold = True
p.font.color.rgb = COLOR_OLIVE

p = tf_w2.add_paragraph()
p.text = ("• Pristine cream parchment card presentation with double-border burgundy trim.\n"
          "• Cleaned and polished visual layout with zero distracting background stripes.\n"
          "• Direct one-click 'Download Official PDF' button linked to Aashi_Srivastava_Resume.pdf.")
p.font.name = FONT_SANS
p.font.size = Pt(9.5)
p.font.color.rgb = COLOR_TEXT_MUTED
p.space_before = Pt(3)


# ==============================================================================
# SLIDE 7: VERIFIED CREDENTIALS, HORIZON & CONNECT (DARK LUXURY ESPRESSO)
# ==============================================================================
slide7 = prs.slides.add_slide(blank_layout)
set_slide_background(slide7, COLOR_DARK_BG)

# Decorative double gold/burgundy borders
outer_border7 = slide7.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(0.5), Inches(12.333), Inches(6.5))
outer_border7.fill.background()
outer_border7.line.color.rgb = COLOR_GOLD
outer_border7.line.width = Pt(1.5)

inner_border7 = slide7.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.6), Inches(0.6), Inches(12.133), Inches(6.3))
inner_border7.fill.background()
inner_border7.line.color.rgb = COLOR_BURGUNDY
inner_border7.line.width = Pt(0.75)

# Broadside Header Ribbon
tb_top7 = slide7.shapes.add_textbox(Inches(1.0), Inches(0.8), Inches(11.333), Inches(0.35))
tf_top7 = tb_top7.text_frame
tf_top7.margin_left = tf_top7.margin_right = tf_top7.margin_top = tf_top7.margin_bottom = 0
p = tf_top7.paragraphs[0]
p.text = "THE OFFICIAL ARCHIVE & DISPATCH COMPENDIUM • VOL. XXVI • CHAPTER 06"
p.font.name = FONT_MONO
p.font.size = Pt(10)
p.font.bold = True
p.font.color.rgb = COLOR_GOLD
p.alignment = PP_ALIGN.CENTER

# Main Title Box
tb_title7 = slide7.shapes.add_textbox(Inches(1.0), Inches(1.15), Inches(11.333), Inches(1.2))
tf_title7 = tb_title7.text_frame
tf_title7.word_wrap = True
tf_title7.margin_left = tf_title7.margin_right = tf_title7.margin_top = tf_title7.margin_bottom = 0

p1 = tf_title7.paragraphs[0]
p1.text = "Verified Credentials & The Horizon Ahead"
p1.font.name = FONT_SERIF
p1.font.size = Pt(28)
p1.font.bold = True
p1.font.color.rgb = COLOR_TEXT_LIGHT
p1.alignment = PP_ALIGN.CENTER

p2 = tf_title7.add_paragraph()
p2.text = "National hackathon selections, technical certifications, and direct recruitment correspondence."
p2.font.name = FONT_SANS
p2.font.size = Pt(12)
p2.font.color.rgb = RGBColor(185, 175, 165)
p2.alignment = PP_ALIGN.CENTER
p2.space_before = Pt(3)

# Two Big Dark Luxury Cards
c_w7 = Inches(5.45)
c_h7 = Inches(3.9)
c_top7 = Inches(2.4)

# Left Card: Verified Credentials
add_card_box(slide7, Inches(1.1), c_top7, c_w7, c_h7, bg_color=COLOR_DARK_CARD, border_color=COLOR_GOLD)
tb_c1 = slide7.shapes.add_textbox(Inches(1.3), c_top7 + Inches(0.2), c_w7 - Inches(0.4), c_h7 - Inches(0.4))
tf_c1 = tb_c1.text_frame
tf_c1.word_wrap = True

p = tf_c1.paragraphs[0]
p.text = "VERIFIED MILESTONE PASSPORTS"
p.font.name = FONT_MONO
p.font.size = Pt(10)
p.font.bold = True
p.font.color.rgb = COLOR_GOLD

p = tf_c1.add_paragraph()
p.text = ("• Smart India Hackathon (SIH 2026 Selection):\n"
          "  Premier national innovation initiative selection for ARS Tourist Guide.\n\n"
          "• Google Developer Groups (GDG DevFest 2025):\n"
          "  Technical builder and attendee across high-impact developer summits.\n\n"
          "• Generative AI Certification (Be10x Certified):\n"
          "  Certified in prompt engineering, LLM tools, and AI workflows.\n\n"
          "• Nerds Room Buildethon & Hackathon Sprint:\n"
          "  Rapid prototype sprint collaboration and mentor-guided development.")
p.font.name = FONT_SANS
p.font.size = Pt(10)
p.font.color.rgb = COLOR_TEXT_LIGHT
p.space_before = Pt(5)

# Right Card: Contact & Dispatch Channels
add_card_box(slide7, Inches(6.783), c_top7, c_w7, c_h7, bg_color=COLOR_DARK_CARD, border_color=COLOR_GOLD)
tb_c2 = slide7.shapes.add_textbox(Inches(6.983), c_top7 + Inches(0.2), c_w7 - Inches(0.4), c_h7 - Inches(0.4))
tf_c2 = tb_c2.text_frame
tf_c2.word_wrap = True

p = tf_c2.paragraphs[0]
p.text = "DISPATCH & CORRESPONDENCE"
p.font.name = FONT_MONO
p.font.size = Pt(10)
p.font.bold = True
p.font.color.rgb = COLOR_GOLD

p = tf_c2.add_paragraph()
p.text = "Open for Data Analyst & AI Opportunities"
p.font.name = FONT_SERIF
p.font.size = Pt(16)
p.font.bold = True
p.font.color.rgb = COLOR_TEXT_LIGHT
p.space_before = Pt(3)

p = tf_c2.add_paragraph()
p.text = ("• Direct Email: aashimaahi23@gmail.com\n"
          "• Live Website: https://aashi230.github.io/portfolio/\n"
          "• GitHub Repository: github.com/Aashi230/portfolio\n"
          "• Academic Institution: SRMU Lucknow (Class of 2029)\n"
          "• Academic Standing: 8.29 CGPA Distinction\n"
          "• Location: Lucknow, Uttar Pradesh, India")
p.font.name = FONT_SANS
p.font.size = Pt(10)
p.font.color.rgb = RGBColor(200, 190, 180)
p.space_before = Pt(5)

p = tf_c2.add_paragraph()
p.text = "Seeking Internship & Entry Roles: Aspiring Data Analyst • AI / DS Tech Learner"
p.font.name = FONT_MONO
p.font.size = Pt(9)
p.font.bold = True
p.font.color.rgb = RGBColor(165, 205, 165)
p.space_before = Pt(6)

# Footer Colophon
tb_foot7 = slide7.shapes.add_textbox(Inches(1.0), Inches(6.45), Inches(11.333), Inches(0.35))
p_foot7 = tb_foot7.text_frame.paragraphs[0]
p_foot7.text = "The Lucknow Gazette & Curated Engineering Dossier • Designed & Handcrafted with Pride • 2026"
p_foot7.font.name = FONT_MONO
p_foot7.font.size = Pt(9.5)
p_foot7.font.color.rgb = COLOR_GOLD
p_foot7.alignment = PP_ALIGN.CENTER


# ==============================================================================
# SAVE THE COMPLETE PRESENTATION
# ==============================================================================
output_filename = "Aashi_Srivastava_Portfolio_Presentation.pptx"
output_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), output_filename)
prs.save(output_path)
print(f"Presentation saved successfully to: {output_path}")
print(f"Total slides in consolidated deck: {len(prs.slides)}")
