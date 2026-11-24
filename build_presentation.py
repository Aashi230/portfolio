"""
build_presentation.py
Generates the authentic VINTAGE EDITORIAL 7-slide PowerPoint presentation (.pptx)
for Aashi Srivastava's Portfolio: 'The Lucknow Gazette & Archive'.

Aesthetic Doctrine:
- Authentic 19th-century broadside gazette styling on fine aged parchment paper.
- Dual-line vintage framing rules, corner accent fleurons, ornamental newspaper rules.
- Classic Georgia serif headlines, letterpress Consolas tags, warm sepia ink typography.
- Authentic projects only: ARS Tourist Guide (SIH 2026 Selection) & Opportunity Hunter (Buildethon Nerds Room).
  (Completely excludes AI Mental Health Journaling).
- All 5 Android Nova live stage photographs in broadside plates (no certificate, no official host badge).
- 4-year Java foundation (Class 9-12) + active Python learning (no JS, no C++).
- Candidate Identity: Undergraduate Student & Tech Learner (AI & DS) • Aspiring Data Analyst (8.29 CGPA).
"""

import os
import subprocess
import time
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# ==============================================================================
# PRESENTATION INITIALIZATION & VINTAGE PALETTE CONFIGURATION
# ==============================================================================
prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

# Authentic Vintage Gazette Color Palette
COLOR_PARCHMENT      = RGBColor(248, 243, 233)  # #F8F3E9 - Aged Warm Parchment Paper
COLOR_PLACARD_BG     = RGBColor(242, 235, 222)  # #F2EBDE - Antique Cream Box Inset
COLOR_PLACARD_LIGHT  = RGBColor(252, 249, 242)  # #FCF9F2 - Clean Cream Inset
COLOR_SEPIA_INK      = RGBColor(38, 28, 20)     # #261C14 - Deep Antique Sepia Ink
COLOR_SEPIA_MUTED    = RGBColor(118, 102, 90)   # #76665A - Muted Archival Pencil
COLOR_OXBLOOD        = RGBColor(112, 26, 40)    # #701A28 - Vintage Burgundy / Oxblood
COLOR_OXBLOOD_LIGHT  = RGBColor(244, 232, 234)  # #F4E8EA - Light oxblood tint
COLOR_OLIVE          = RGBColor(68, 85, 56)     # #445538 - Antique Forest Moss
COLOR_OLIVE_LIGHT    = RGBColor(236, 240, 232)  # #ECF0E8 - Light olive tint
COLOR_GOLD           = RGBColor(184, 138, 48)   # #B88A30 - Antique Tarnished Gold
COLOR_GOLD_LIGHT     = RGBColor(248, 242, 228)  # #F8F2E4 - Light gold tint
COLOR_BORDER         = RGBColor(198, 180, 158)  # #C6B49E - Aged Broadside Rule
COLOR_BORDER_DARK    = RGBColor(140, 115, 95)   # #8C735F - Darker Accent Rule
COLOR_WHITE          = RGBColor(255, 255, 255)

FONT_SERIF = "Georgia"
FONT_SANS  = "Calibri"
FONT_MONO  = "Consolas"

blank_layout = prs.slide_layouts[6]

# ==============================================================================
# VINTAGE HELPER FUNCTIONS (BROADSIDE FRAMING & TYPOGRAPHY)
# ==============================================================================
def apply_vintage_page_frame(slide):
    """Draws an authentic aged parchment background with double-line vintage broadside framing."""
    # 1. Warm parchment base
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    bg.fill.solid()
    bg.fill.fore_color.rgb = COLOR_PARCHMENT
    bg.line.fill.background()
    
    # 2. Outer vintage broadside rule (thick oxblood / sepia line)
    outer = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.4), Inches(0.35), Inches(12.533), Inches(6.8))
    outer.fill.background()
    outer.line.color.rgb = COLOR_OXBLOOD
    outer.line.width = Pt(1.8)
    
    # 3. Inner fine hairline rule (inset by 0.08 in)
    inner = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.48), Inches(0.43), Inches(12.373), Inches(6.64))
    inner.fill.background()
    inner.line.color.rgb = COLOR_GOLD
    inner.line.width = Pt(0.75)
    
    # 4. Corner accent blocks (traditional 19th-century broadside corner seals)
    corner_size = Inches(0.12)
    corners = [
        (Inches(0.42), Inches(0.37)),
        (Inches(12.793), Inches(0.37)),
        (Inches(0.42), Inches(7.01)),
        (Inches(12.793), Inches(7.01))
    ]
    for cx, cy in corners:
        c = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, cx, cy, corner_size, corner_size)
        c.fill.solid()
        c.fill.fore_color.rgb = COLOR_OXBLOOD
        c.line.fill.background()

def add_vintage_header(slide, section_tag, title_text, subtitle_text=""):
    """Creates a classic broadside newspaper section header with double hairline divider rules."""
    tb = slide.shapes.add_textbox(Inches(0.7), Inches(0.55), Inches(11.933), Inches(1.15))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
    
    # Running Section Tag
    p_tag = tf.paragraphs[0]
    p_tag.text = f"♦  {section_tag.upper()}  ♦"
    p_tag.font.name = FONT_MONO
    p_tag.font.size = Pt(9.5)
    p_tag.font.bold = True
    p_tag.font.color.rgb = COLOR_OXBLOOD
    p_tag.space_after = Pt(2)
    
    # Primary Headline
    p_title = tf.add_paragraph()
    p_title.text = title_text
    p_title.font.name = FONT_SERIF
    p_title.font.size = Pt(21)
    p_title.font.bold = True
    p_title.font.color.rgb = COLOR_SEPIA_INK
    p_title.space_after = Pt(2)
    
    # Subtitle
    if subtitle_text:
        p_sub = tf.add_paragraph()
        p_sub.text = subtitle_text
        p_sub.font.name = FONT_SERIF
        p_sub.font.size = Pt(11)
        p_sub.font.italic = True
        p_sub.font.color.rgb = COLOR_SEPIA_MUTED
    
    # Broadside Double Hairline Divider Rules
    rule1 = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.7), Inches(1.68), Inches(11.933), Inches(0.018))
    rule1.fill.solid()
    rule1.fill.fore_color.rgb = COLOR_OXBLOOD
    rule1.line.fill.background()
    
    rule2 = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.7), Inches(1.72), Inches(11.933), Inches(0.008))
    rule2.fill.solid()
    rule2.fill.fore_color.rgb = COLOR_GOLD
    rule2.line.fill.background()

def add_vintage_placard(slide, left, top, width, height, accent_color=COLOR_OXBLOOD, bg_color=COLOR_PLACARD_BG):
    """Draws a vintage broadside placard box with a top decorative accent bar."""
    box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    box.fill.solid()
    box.fill.fore_color.rgb = bg_color
    box.line.color.rgb = COLOR_BORDER
    box.line.width = Pt(1.2)
    
    # Top accent rule
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, Inches(0.07))
    bar.fill.solid()
    bar.fill.fore_color.rgb = accent_color
    bar.line.fill.background()
    return box

def add_vintage_photo_plate(slide, left, top, width, height, img_path, caption=""):
    """Draws an antique framed picture plate with a broadside caption label."""
    # Outer frame
    frame = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    frame.fill.solid()
    frame.fill.fore_color.rgb = COLOR_WHITE
    frame.line.color.rgb = COLOR_BORDER_DARK
    frame.line.width = Pt(1.2)
    
    margin = Inches(0.08)
    cap_h = Inches(0.32) if caption else 0
    img_w = width - (margin * 2)
    img_h = height - (margin * 2) - cap_h
    
    if os.path.exists(img_path):
        slide.shapes.add_picture(img_path, left + margin, top + margin, width=img_w, height=img_h)
    
    if caption:
        cap_box = slide.shapes.add_textbox(left + margin, top + height - Inches(0.34), width - (margin * 2), Inches(0.28))
        tf = cap_box.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
        p = tf.paragraphs[0]
        p.text = caption
        p.font.name = FONT_MONO
        p.font.size = Pt(8.5)
        p.font.bold = True
        p.font.color.rgb = COLOR_OXBLOOD
        p.alignment = PP_ALIGN.CENTER

def add_vintage_seal(slide, left, top, width, height, value_text, label_text, color=COLOR_OXBLOOD):
    """Draws an archival seal / rubber-stamp style badge."""
    seal = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    seal.fill.solid()
    seal.fill.fore_color.rgb = COLOR_PLACARD_BG
    seal.line.color.rgb = color
    seal.line.width = Pt(1.2)
    
    tb = slide.shapes.add_textbox(left + Inches(0.04), top + Inches(0.08), width - Inches(0.08), height - Inches(0.16))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    
    p1 = tf.paragraphs[0]
    p1.text = value_text
    p1.font.name = FONT_SERIF
    p1.font.size = Pt(16)
    p1.font.bold = True
    p1.font.color.rgb = color
    p1.alignment = PP_ALIGN.CENTER
    
    p2 = tf.add_paragraph()
    p2.text = label_text
    p2.font.name = FONT_MONO
    p2.font.size = Pt(8)
    p2.font.bold = True
    p2.font.color.rgb = COLOR_SEPIA_MUTED
    p2.alignment = PP_ALIGN.CENTER
    p2.space_before = Pt(1)

content_top = Inches(1.9)


# ==============================================================================
# SLIDE 1: THE FRONTISPIECE / GRAND VINTAGE BROADSIDE COVER
# ==============================================================================
slide1 = prs.slides.add_slide(blank_layout)
apply_vintage_page_frame(slide1)

# Top Running Dateline Ribbon
tb_date = slide1.shapes.add_textbox(Inches(0.7), Inches(0.55), Inches(11.933), Inches(0.35))
tf_date = tb_date.text_frame
tf_date.margin_left = tf_date.margin_right = tf_date.margin_top = tf_date.margin_bottom = 0
p = tf_date.paragraphs[0]
p.text = "THE LUCKNOW GAZETTE & ARCHIVAL COMPENDIUM • VOL. XXVI • ISSUE 01 • ESTABLISHED 2024"
p.font.name = FONT_MONO
p.font.size = Pt(9.5)
p.font.bold = True
p.font.color.rgb = COLOR_OXBLOOD
p.alignment = PP_ALIGN.CENTER

# Headline Box
tb_head = slide1.shapes.add_textbox(Inches(0.7), Inches(0.9), Inches(11.933), Inches(2.2))
tf_head = tb_head.text_frame
tf_head.word_wrap = True
tf_head.margin_left = tf_head.margin_right = tf_head.margin_top = tf_head.margin_bottom = 0

p1 = tf_head.paragraphs[0]
p1.text = "The Lucknow Gazette"
p1.font.name = FONT_SERIF
p1.font.size = Pt(44)
p1.font.bold = True
p1.font.color.rgb = COLOR_SEPIA_INK
p1.alignment = PP_ALIGN.CENTER

p2 = tf_head.add_paragraph()
p2.text = "& CURATED ENGINEERING DOSSIER"
p2.font.name = FONT_SERIF
p2.font.size = Pt(18)
p2.font.bold = True
p2.font.color.rgb = COLOR_OXBLOOD
p2.alignment = PP_ALIGN.CENTER
p2.space_before = Pt(3)

p3 = tf_head.add_paragraph()
p3.text = "“Aesthetic Doctrine: Vintage Editorial Luxury Meets Computational Rigor”"
p3.font.name = FONT_SERIF
p3.font.size = Pt(12)
p3.font.italic = True
p3.font.color.rgb = COLOR_SEPIA_MUTED
p3.alignment = PP_ALIGN.CENTER
p3.space_before = Pt(4)

# Central Archival Placard Box
center_w = Inches(10.5)
center_h = Inches(2.8)
center_left = Inches(1.416)
center_top = Inches(3.2)
add_vintage_placard(slide1, center_left, center_top, center_w, center_h, accent_color=COLOR_OXBLOOD)

# Portrait Plate inside Central Box
portrait_path = "assets/images/profile/portrait.jpg"
if os.path.exists(portrait_path):
    add_vintage_photo_plate(slide1, center_left + Inches(0.35), center_top + Inches(0.25), Inches(1.85), Inches(2.3),
                            portrait_path, "Aashi Srivastava")

# Center Identity Narrative
tb_id = slide1.shapes.add_textbox(center_left + Inches(2.45), center_top + Inches(0.25), Inches(4.7), Inches(2.3))
tf_id = tb_id.text_frame
tf_id.word_wrap = True
tf_id.margin_left = tf_id.margin_right = tf_id.margin_top = tf_id.margin_bottom = 0

p = tf_id.paragraphs[0]
p.text = "AASHI SRIVASTAVA"
p.font.name = FONT_SERIF
p.font.size = Pt(20)
p.font.bold = True
p.font.color.rgb = COLOR_SEPIA_INK

p = tf_id.add_paragraph()
p.text = "Student & Tech Learner (AI & DS) • Aspiring Data Analyst"
p.font.name = FONT_MONO
p.font.size = Pt(10)
p.font.bold = True
p.font.color.rgb = COLOR_OXBLOOD
p.space_before = Pt(2)

p = tf_id.add_paragraph()
p.text = ("• B.Tech in CSE (Artificial Intelligence & Data Science)\n"
          "• Shri Ramswaroop Memorial University (SRMU), Lucknow\n"
          "• Academic Distinction: 8.29 CGPA • Class of 2029\n"
          "• Four-Year Java Heritage (Grades 9–12) + Active Python Stack")
p.font.name = FONT_SANS
p.font.size = Pt(10.5)
p.font.color.rgb = COLOR_SEPIA_INK
p.space_before = Pt(4)

p = tf_id.add_paragraph()
p.text = "Repository: https://aashi230.github.io/portfolio/"
p.font.name = FONT_MONO
p.font.size = Pt(9)
p.font.color.rgb = COLOR_OLIVE
p.space_before = Pt(5)

# 4 Archival Seals inside Central Box (Right Column)
s_left = center_left + Inches(7.35)
s_top = center_top + Inches(0.25)
s_w = Inches(1.35)
s_h = Inches(1.05)
s_gap = Inches(0.12)

add_vintage_seal(slide1, s_left, s_top, s_w, s_h, "8.29", "SRMU CGPA", color=COLOR_OXBLOOD)
add_vintage_seal(slide1, s_left + s_w + s_gap, s_top, s_w, s_h, "SIH '26", "HACKATHON", color=COLOR_OLIVE)
add_vintage_seal(slide1, s_left, s_top + s_h + s_gap, s_w, s_h, "4+", "PROJECTS", color=COLOR_GOLD)
add_vintage_seal(slide1, s_left + s_w + s_gap, s_top + s_h + s_gap, s_w, s_h, "HOST", "ANDRO NOVA", color=COLOR_OXBLOOD)

# Broadside Foot Colophon
tb_foot = slide1.shapes.add_textbox(Inches(0.7), Inches(6.45), Inches(11.933), Inches(0.35))
p_foot = tb_foot.text_frame.paragraphs[0]
p_foot.text = "PRINTED & PUBLISHED AT LUCKNOW, INDIA • ANNO DOMINI MMXXVI • ARCHIVAL EDITION"
p_foot.font.name = FONT_MONO
p_foot.font.size = Pt(9)
p_foot.font.color.rgb = COLOR_SEPIA_MUTED
p_foot.alignment = PP_ALIGN.CENTER


# ==============================================================================
# SLIDE 2: SECTION I — THE ACADEMIC DOSSIER & SCHOLARLY STANDING
# ==============================================================================
slide2 = prs.slides.add_slide(blank_layout)
apply_vintage_page_frame(slide2)
add_vintage_header(slide2, "SECTION I // CANDIDATE PROFILE", "The Academic Dossier & Scholarly Standing",
                   "Verified undergraduate credentials, university trajectory, and technical identity.")

left_w2 = Inches(5.6)
add_vintage_placard(slide2, Inches(0.7), content_top, left_w2, Inches(4.9), accent_color=COLOR_OXBLOOD)

tb_abt = slide2.shapes.add_textbox(Inches(0.95), content_top + Inches(0.18), left_w2 - Inches(0.5), Inches(4.5))
tf_abt = tb_abt.text_frame
tf_abt.word_wrap = True

p = tf_abt.paragraphs[0]
p.text = "SCHOLASTIC RECORD & HERITAGE"
p.font.name = FONT_MONO
p.font.size = Pt(9.5)
p.font.bold = True
p.font.color.rgb = COLOR_OXBLOOD

p = tf_abt.add_paragraph()
p.text = "Rigorous Foundation in AI & Data Science"
p.font.name = FONT_SERIF
p.font.size = Pt(16)
p.font.bold = True
p.font.color.rgb = COLOR_SEPIA_INK
p.space_before = Pt(2)

p = tf_abt.add_paragraph()
p.text = ("• Collegiate Distinction: Second-year undergraduate pursuing B.Tech in Computer Science Engineering (Artificial Intelligence & Data Science) at SRMU Lucknow, holding a distinguished 8.29 CGPA.\n\n"
          "• Four-Year Java Heritage: Built core algorithmic problem-solving and structured programming through four continuous years of formal Java instruction from Class 9 to 12, mastering object-oriented classes and hierarchies.\n\n"
          "• Active Python Learning: Actively applying Python across data science, exploratory data analysis (EDA), data cleaning, statistical modeling, and machine learning.\n\n"
          "• Professional Aspirations: Combining the persistent curiosity of a Student & Tech Learner with industry goals as an Aspiring Data Analyst.")
p.font.name = FONT_SANS
p.font.size = Pt(10)
p.font.color.rgb = COLOR_SEPIA_INK
p.space_before = Pt(4)

# Academic Highlight Placard
add_vintage_placard(slide2, Inches(0.95), content_top + Inches(3.75), left_w2 - Inches(0.5), Inches(0.95),
                    accent_color=COLOR_GOLD, bg_color=COLOR_GOLD_LIGHT)
tb_hi = slide2.shapes.add_textbox(Inches(1.05), content_top + Inches(3.8), left_w2 - Inches(0.7), Inches(0.85))
tf_hi = tb_hi.text_frame
tf_hi.word_wrap = True
p = tf_hi.paragraphs[0]
p.text = "VERIFIED REGISTRATION & ACADEMIC DOSSIER"
p.font.name = FONT_MONO
p.font.size = Pt(8.5)
p.font.bold = True
p.font.color.rgb = COLOR_OXBLOOD
p = tf_hi.add_paragraph()
p.text = "B.Tech CSE (AI & DS) • SRMU Lucknow • 8.29 CGPA Distinction • Class of 2029"
p.font.name = FONT_SANS
p.font.size = Pt(9.5)
p.font.bold = True
p.font.color.rgb = COLOR_SEPIA_INK
p.space_before = Pt(2)

# Right Column: Archival Dossier Picture Plate
add_vintage_photo_plate(slide2, Inches(6.6), content_top, Inches(6.033), Inches(4.9),
                        "presentation_assets/02_about_dossier.png",
                        "PLATE I: THE ARCHIVAL CANDIDATE DOSSIER • CREDENTIALS & PASSPORTS")


# ==============================================================================
# SLIDE 3: SECTION II — THE COMPUTATIONAL ARSENAL & DATA STACK
# ==============================================================================
slide3 = prs.slides.add_slide(blank_layout)
apply_vintage_page_frame(slide3)
add_vintage_header(slide3, "SECTION II // THE COMPUTATIONAL ARSENAL", "Technical Arsenal & Data Analytics Stack",
                   "Structured multi-tier proficiencies across languages, analytical libraries, and engineering tools.")

col_w3 = Inches(3.75)
gap3 = Inches(0.341)
c1_left = Inches(0.7)
c2_left = c1_left + col_w3 + gap3
c3_left = c2_left + col_w3 + gap3

# Column 1: Core Languages
add_vintage_placard(slide3, c1_left, content_top, col_w3, Inches(4.9), accent_color=COLOR_OXBLOOD)
tb_c1 = slide3.shapes.add_textbox(c1_left + Inches(0.2), content_top + Inches(0.18), col_w3 - Inches(0.4), Inches(4.5))
tf_c1 = tb_c1.text_frame
tf_c1.word_wrap = True

p = tf_c1.paragraphs[0]
p.text = "COLUMN I // CODE FOUNDATIONS"
p.font.name = FONT_MONO
p.font.size = Pt(9)
p.font.bold = True
p.font.color.rgb = COLOR_OXBLOOD

p = tf_c1.add_paragraph()
p.text = "Core Programming"
p.font.name = FONT_SERIF
p.font.size = Pt(16)
p.font.bold = True
p.font.color.rgb = COLOR_SEPIA_INK
p.space_before = Pt(2)

p = tf_c1.add_paragraph()
p.text = ("• Java (Class 9–12 Foundation):\n"
          "  4-year academic background establishing strong object-oriented principles, classes, methods, and algorithmic structure.\n\n"
          "• Python (Actively Learning):\n"
          "  Focused adoption for data analytics, scientific computation, data cleaning, and AI workflows.\n\n"
          "• C (Procedural Fundamentals):\n"
          "  Solid grounding in memory concepts, procedural logic, pointers, and data representations.\n\n"
          "• SQL (Database Querying):\n"
          "  Relational schema design, filtering, aggregate queries, and data extraction.")
p.font.name = FONT_SANS
p.font.size = Pt(9.5)
p.font.color.rgb = COLOR_SEPIA_INK
p.space_before = Pt(4)

# Column 2: AI & Data Science
add_vintage_placard(slide3, c2_left, content_top, col_w3, Inches(4.9), accent_color=COLOR_OLIVE)
tb_c2 = slide3.shapes.add_textbox(c2_left + Inches(0.2), content_top + Inches(0.18), col_w3 - Inches(0.4), Inches(4.5))
tf_c2 = tb_c2.text_frame
tf_c2.word_wrap = True

p = tf_c2.paragraphs[0]
p.text = "COLUMN II // DATA & ANALYTICS"
p.font.name = FONT_MONO
p.font.size = Pt(9)
p.font.bold = True
p.font.color.rgb = COLOR_OLIVE

p = tf_c2.add_paragraph()
p.text = "Analytical Toolkit"
p.font.name = FONT_SERIF
p.font.size = Pt(16)
p.font.bold = True
p.font.color.rgb = COLOR_SEPIA_INK
p.space_before = Pt(2)

p = tf_c2.add_paragraph()
p.text = ("• Exploratory Data Analysis (EDA):\n"
          "  Statistical summary extraction, data distributions, correlation analysis, and data cleansing.\n\n"
          "• Data Science Libraries:\n"
          "  Hands-on learning with Pandas (DataFrames) and NumPy (numerical matrices & vectors).\n\n"
          "• Data Visualization:\n"
          "  Matplotlib and Seaborn for publication-grade graphs, histograms, and scatter plots.\n\n"
          "• Generative AI & Prompting:\n"
          "  Be10x Certified in prompt engineering, AI tool integration, and accelerated workflows.")
p.font.name = FONT_SANS
p.font.size = Pt(9.5)
p.font.color.rgb = COLOR_SEPIA_INK
p.space_before = Pt(4)

# Column 3: Tools & Platform
add_vintage_placard(slide3, c3_left, content_top, col_w3, Inches(4.9), accent_color=COLOR_GOLD)
tb_c3 = slide3.shapes.add_textbox(c3_left + Inches(0.2), content_top + Inches(0.18), col_w3 - Inches(0.4), Inches(4.5))
tf_c3 = tb_c3.text_frame
tf_c3.word_wrap = True

p = tf_c3.paragraphs[0]
p.text = "COLUMN III // WORKFLOW & SYSTEM"
p.font.name = FONT_MONO
p.font.size = Pt(9)
p.font.bold = True
p.font.color.rgb = COLOR_GOLD

p = tf_c3.add_paragraph()
p.text = "Engineering Tools"
p.font.name = FONT_SERIF
p.font.size = Pt(16)
p.font.bold = True
p.font.color.rgb = COLOR_SEPIA_INK
p.space_before = Pt(2)

p = tf_c3.add_paragraph()
p.text = ("• Version Control:\n"
          "  Git and GitHub proficiency for branching, commits, release tagging, and GitHub Pages CI/CD.\n\n"
          "• Development Environments:\n"
          "  VS Code, Jupyter Notebooks, and Google Colab for interactive computational exploration.\n\n"
          "• Web Interface:\n"
          "  Semantic HTML5 and modern modular CSS3 (CSS Grid, Flexbox, responsive typography).\n\n"
          "• Command Shell:\n"
          "  Windows PowerShell, shell scripts, and reproducible archival documentation.")
p.font.name = FONT_SANS
p.font.size = Pt(9.5)
p.font.color.rgb = COLOR_SEPIA_INK
p.space_before = Pt(4)


# ==============================================================================
# SLIDE 4: SECTION III — THE CURATED ENGINEERING WORKS (REAL PROJECTS)
# ==============================================================================
slide4 = prs.slides.add_slide(blank_layout)
apply_vintage_page_frame(slide4)
add_vintage_header(slide4, "SECTION III // THE CURATED ENGINEERING WORKS", "Flagship Projects & Computational Solutions",
                   "Practical software architectures engineered for travel intelligence and opportunity discovery.")

left_w4 = Inches(5.6)

# Project Card 1: ARS Tourist Guide (Smart India Hackathon 2026 Selection)
add_vintage_placard(slide4, Inches(0.7), content_top, left_w4, Inches(2.38), accent_color=COLOR_OXBLOOD)
tb_pj1 = slide4.shapes.add_textbox(Inches(0.95), content_top + Inches(0.14), left_w4 - Inches(0.5), Inches(2.15))
tf_pj1 = tb_pj1.text_frame
tf_pj1.word_wrap = True

p = tf_pj1.paragraphs[0]
p.text = "SMART INDIA HACKATHON 2026 // CLEARED INTERNAL SELECTION"
p.font.name = FONT_MONO
p.font.size = Pt(8.5)
p.font.bold = True
p.font.color.rgb = COLOR_OXBLOOD

p = tf_pj1.add_paragraph()
p.text = "ARS Tourist Guide & Heritage Explorer"
p.font.name = FONT_SERIF
p.font.size = Pt(15)
p.font.bold = True
p.font.color.rgb = COLOR_SEPIA_INK
p.space_before = Pt(2)

p = tf_pj1.add_paragraph()
p.text = ("• Problem Statement SIH26207 (Travel & Tourism) by Team TOTAL CODERS; cleared university internal round.\n"
          "• Integrated interactive spatial mapping utilizing OpenStreetMap API for localized heritage navigation.\n"
          "• Real-time meteorological forecast updates synchronized via Open-Meteo REST endpoints.\n"
          "• Integrated Razorpay checkout for guided tour pass reservations and SQLite persistence layer.")
p.font.name = FONT_SANS
p.font.size = Pt(9.5)
p.font.color.rgb = COLOR_SEPIA_INK
p.space_before = Pt(3)

# Project Card 2: Opportunity Hunter (Buildethon Nerds Room)
add_vintage_placard(slide4, Inches(0.7), content_top + Inches(2.52), left_w4, Inches(2.38), accent_color=COLOR_OLIVE)
tb_pj2 = slide4.shapes.add_textbox(Inches(0.95), content_top + Inches(2.66), left_w4 - Inches(0.5), Inches(2.15))
tf_pj2 = tb_pj2.text_frame
tf_pj2.word_wrap = True

p = tf_pj2.paragraphs[0]
p.text = "BUILDETHON HACKATHON // NERDS ROOM COMMUNITY"
p.font.name = FONT_MONO
p.font.size = Pt(8.5)
p.font.bold = True
p.font.color.rgb = COLOR_OLIVE

p = tf_pj2.add_paragraph()
p.text = "Opportunity Hunter & Scholarship Platform"
p.font.name = FONT_SERIF
p.font.size = Pt(15)
p.font.bold = True
p.font.color.rgb = COLOR_SEPIA_INK
p.space_before = Pt(2)

p = tf_pj2.add_paragraph()
p.text = ("• Engineered and showcased during the Buildethon hackathon organized by Nerds Room Community.\n"
          "• Curated directory of verified scholarships, research fellowships, and student internships.\n"
          "• Streamlined eligibility filtering by academic discipline, semester standing, and requirements.\n"
          "• Direct navigation links to official application portals without deceptive intermediaries.")
p.font.name = FONT_SANS
p.font.size = Pt(9.5)
p.font.color.rgb = COLOR_SEPIA_INK
p.space_before = Pt(3)

# Right Column: Projects Archive Picture Plate
add_vintage_photo_plate(slide4, Inches(6.6), content_top, Inches(6.033), Inches(4.9),
                        "presentation_assets/03_projects_work.png",
                        "PLATE III: THE ARCHIVAL PROJECT REPOSITORY • INTERACTIVE MODALS & SCHEMAS")


# ==============================================================================
# SLIDE 5: SECTION IV — PUBLIC ORATORY & STAGE LEADERSHIP (ALL 5 LIVE FRAMES)
# ==============================================================================
slide5 = prs.slides.add_slide(blank_layout)
apply_vintage_page_frame(slide5)
add_vintage_header(slide5, "SECTION IV // PUBLIC ORATORY & LEADERSHIP", "Android Nova 2.0 — Symposium Stage Host & Coordinator",
                   "Master of ceremonies leading SRMU's premier technical symposium across all 5 live event frames.")

left_nova_w = Inches(3.5)
add_vintage_placard(slide5, Inches(0.7), content_top, left_nova_w, Inches(4.9), accent_color=COLOR_OXBLOOD)
tb_nv = slide5.shapes.add_textbox(Inches(0.9), content_top + Inches(0.18), left_nova_w - Inches(0.4), Inches(4.5))
tf_nv = tb_nv.text_frame
tf_nv.word_wrap = True

p = tf_nv.paragraphs[0]
p.text = "SYMPOSIUM HOST & COORDINATOR"
p.font.name = FONT_MONO
p.font.size = Pt(9)
p.font.bold = True
p.font.color.rgb = COLOR_OXBLOOD

p = tf_nv.add_paragraph()
p.text = "Stage Leadership & Poise"
p.font.name = FONT_SERIF
p.font.size = Pt(15)
p.font.bold = True
p.font.color.rgb = COLOR_SEPIA_INK
p.space_before = Pt(2)

p = tf_nv.add_paragraph()
p.text = ("• Selected as stage anchor for Android Nova 2.0 before hundreds of attendees, university faculty, and industry guests.\n\n"
          "• Commanded the auditorium stage with poise, speaker introductions, and dignitary felicitations.\n\n"
          "• Orchestrated backstage logistics, synchronization across technical teams, and audio-visual flow.\n\n"
          "• Demonstrated high-stakes public speaking, rapid adaptability, and crowd engagement.")
p.font.name = FONT_SANS
p.font.size = Pt(9.5)
p.font.color.rgb = COLOR_SEPIA_INK
p.space_before = Pt(4)

p = tf_nv.add_paragraph()
p.text = "EVENT DOSSIER & IMPACT"
p.font.name = FONT_MONO
p.font.size = Pt(9)
p.font.bold = True
p.font.color.rgb = COLOR_OLIVE
p.space_before = Pt(6)

p = tf_nv.add_paragraph()
p.text = ("• Role: Stage Anchor & Event Host\n"
          "• Institution: SRMU Lucknow\n"
          "• Scope: Flagship University Symposium\n"
          "• Documentation: 5 Live Stage Frames Below")
p.font.name = FONT_SANS
p.font.size = Pt(9.5)
p.font.color.rgb = COLOR_SEPIA_INK
p.space_before = Pt(2)

# Right Area: Prominent 5 Live Photo Spread (No Certificate)
# Top Row: 3 Stage Frames across width 8.133 in
add_vintage_photo_plate(slide5, Inches(4.45), content_top, Inches(2.65), Inches(2.35),
                        "presentation_assets/nova_1.jpg", "Frame 01: Stage Anchoring")
add_vintage_photo_plate(slide5, Inches(7.22), content_top, Inches(2.85), Inches(2.35),
                        "presentation_assets/nova_2.jpg", "Frame 02: Keynote & Audience")
add_vintage_photo_plate(slide5, Inches(10.19), content_top, Inches(2.443), Inches(2.35),
                        "presentation_assets/nova_3.jpg", "Frame 03: Organizing Desk")

# Bottom Row: 2 Wide Stage Frames across width 8.133 in
add_vintage_photo_plate(slide5, Inches(4.45), content_top + Inches(2.55), Inches(4.0), Inches(2.35),
                        "presentation_assets/nova_4.jpg", "Frame 04: Dignitary Felicitation")
add_vintage_photo_plate(slide5, Inches(8.633), content_top + Inches(2.55), Inches(4.0), Inches(2.35),
                        "presentation_assets/nova_5.jpg", "Frame 05: Host Commemorative Portrait")


# ==============================================================================
# SLIDE 6: SECTION V — DIGITAL ARCHITECTURE & CURRICULUM VITAE
# ==============================================================================
slide6 = prs.slides.add_slide(blank_layout)
apply_vintage_page_frame(slide6)
add_vintage_header(slide6, "SECTION V // DIGITAL PLATFORM & RESUME", "The Portfolio Platform & Executive Curriculum Vitae",
                   "Handcrafted vintage broadside engineering, full-screen interactive inspection, and verified resume.")

col_w6 = Inches(5.8)
gap6 = Inches(0.333)

# Column 1: Web Platform Architecture
add_vintage_photo_plate(slide6, Inches(0.7), content_top, col_w6, Inches(3.1),
                        "presentation_assets/01_hero_landing.png", "Landing Broadside Architecture & Metrics Ticker")

add_vintage_placard(slide6, Inches(0.7), content_top + Inches(3.25), col_w6, Inches(1.65), accent_color=COLOR_OXBLOOD)
tb_w1 = slide6.shapes.add_textbox(Inches(0.85), content_top + Inches(3.35), col_w6 - Inches(0.3), Inches(1.45))
tf_w1 = tb_w1.text_frame
tf_w1.word_wrap = True

p = tf_w1.paragraphs[0]
p.text = "VINTAGE EDITORIAL LUXURY PHILOSOPHY"
p.font.name = FONT_MONO
p.font.size = Pt(9)
p.font.bold = True
p.font.color.rgb = COLOR_OXBLOOD

p = tf_w1.add_paragraph()
p.text = ("• Bespoke 19th-century broadside aesthetics handcrafted in semantic HTML5 & modular CSS3 without heavy frameworks.\n"
          "• Dynamic full-screen image lightbox inspection and interactive project detail briefs.\n"
          "• Live contact dispatch form powered by FormSubmit with instant verified email delivery.")
p.font.name = FONT_SANS
p.font.size = Pt(9)
p.font.color.rgb = COLOR_SEPIA_INK
p.space_before = Pt(3)

# Column 2: Executive Curriculum Vitae
add_vintage_photo_plate(slide6, Inches(0.7) + col_w6 + gap6, content_top, col_w6, Inches(3.1),
                        "presentation_assets/04_resume_cv.png", "Curriculum Vitae & One-Click PDF Resume Download")

add_vintage_placard(slide6, Inches(0.7) + col_w6 + gap6, content_top + Inches(3.25), col_w6, Inches(1.65), accent_color=COLOR_OLIVE)
tb_w2 = slide6.shapes.add_textbox(Inches(0.85) + col_w6 + gap6, content_top + Inches(3.35), col_w6 - Inches(0.3), Inches(1.45))
tf_w2 = tb_w2.text_frame
tf_w2.word_wrap = True

p = tf_w2.paragraphs[0]
p.text = "ARCHIVAL CV & DIRECT PDF ACCESS"
p.font.name = FONT_MONO
p.font.size = Pt(9)
p.font.bold = True
p.font.color.rgb = COLOR_OLIVE

p = tf_w2.add_paragraph()
p.text = ("• Pristine cream parchment presentation with double-border burgundy trim and clean aligned typography.\n"
          "• Eliminates distracting notebook stripes for maximum legibility and executive professionalism.\n"
          "• Direct one-click 'Download Official PDF' button linked to Aashi_Srivastava_Resume.pdf.")
p.font.name = FONT_SANS
p.font.size = Pt(9)
p.font.color.rgb = COLOR_SEPIA_INK
p.space_before = Pt(3)


# ==============================================================================
# SLIDE 7: SECTION VI — VERIFIED CREDENTIALS, HORIZON & DISPATCH
# ==============================================================================
slide7 = prs.slides.add_slide(blank_layout)
apply_vintage_page_frame(slide7)
add_vintage_header(slide7, "SECTION VI // CREDENTIALS & DISPATCH", "Verified Milestone Passports & The Horizon Ahead",
                   "National hackathon milestones, verified certifications, and direct recruitment correspondence.")

c_w7 = Inches(5.8)
c_h7 = Inches(4.45)
c_top7 = content_top
c1_left7 = Inches(0.7)
c2_left7 = Inches(6.833)

# Left Card: Verified Milestone Passports
add_vintage_placard(slide7, c1_left7, c_top7, c_w7, c_h7, accent_color=COLOR_OXBLOOD)
tb_c1 = slide7.shapes.add_textbox(c1_left7 + Inches(0.2), c_top7 + Inches(0.18), c_w7 - Inches(0.4), c_h7 - Inches(0.4))
tf_c1 = tb_c1.text_frame
tf_c1.word_wrap = True

p = tf_c1.paragraphs[0]
p.text = "VERIFIED MILESTONE PASSPORTS"
p.font.name = FONT_MONO
p.font.size = Pt(9.5)
p.font.bold = True
p.font.color.rgb = COLOR_OXBLOOD

p = tf_c1.add_paragraph()
p.text = ("• Smart India Hackathon (SIH 2026 Selection):\n"
          "  Premier national innovation initiative selection for ARS Tourist Guide with Team TOTAL CODERS.\n\n"
          "• Google Developer Groups (GDG DevFest 2025):\n"
          "  Technical builder and attendee across high-impact developer summits.\n\n"
          "• Generative AI Certification (Be10x Certified):\n"
          "  Certified in prompt engineering, LLM tools, and accelerated AI workflows.\n\n"
          "• Nerds Room Buildethon & Hackathon Sprint:\n"
          "  Rapid prototype sprint collaboration and mentor-guided development for Opportunity Hunter.")
p.font.name = FONT_SANS
p.font.size = Pt(10)
p.font.color.rgb = COLOR_SEPIA_INK
p.space_before = Pt(5)

# Right Card: Contact & Dispatch Channels
add_vintage_placard(slide7, c2_left7, c_top7, c_w7, c_h7, accent_color=COLOR_GOLD)
tb_c2 = slide7.shapes.add_textbox(c2_left7 + Inches(0.2), c_top7 + Inches(0.18), c_w7 - Inches(0.4), c_h7 - Inches(0.4))
tf_c2 = tb_c2.text_frame
tf_c2.word_wrap = True

p = tf_c2.paragraphs[0]
p.text = "DISPATCH & CORRESPONDENCE"
p.font.name = FONT_MONO
p.font.size = Pt(9.5)
p.font.bold = True
p.font.color.rgb = COLOR_GOLD

p = tf_c2.add_paragraph()
p.text = "Open for Data Analyst & AI Opportunities"
p.font.name = FONT_SERIF
p.font.size = Pt(16)
p.font.bold = True
p.font.color.rgb = COLOR_SEPIA_INK
p.space_before = Pt(2)

p = tf_c2.add_paragraph()
p.text = ("• Direct Email: aashimaahi23@gmail.com\n"
          "• Live Website: https://aashi230.github.io/portfolio/\n"
          "• GitHub Repository: github.com/Aashi230/portfolio\n"
          "• Academic Institution: SRMU Lucknow (Class of 2029)\n"
          "• Academic Standing: 8.29 CGPA Distinction\n"
          "• Location: Lucknow, Uttar Pradesh, India")
p.font.name = FONT_SANS
p.font.size = Pt(10)
p.font.color.rgb = COLOR_SEPIA_INK
p.space_before = Pt(5)

p = tf_c2.add_paragraph()
p.text = "Seeking Roles: Aspiring Data Analyst • AI / DS Tech Learner"
p.font.name = FONT_MONO
p.font.size = Pt(9.5)
p.font.bold = True
p.font.color.rgb = COLOR_OLIVE
p.space_before = Pt(6)


# ==============================================================================
# SAVE THE COMPLETE PRESENTATION (RESILIENT FILE LOCK HANDLING)
# ==============================================================================
output_filename = "Aashi_Srivastava_Portfolio_Presentation.pptx"
output_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), output_filename)
vintage_filename = "Aashi_Srivastava_Portfolio_Vintage.pptx"
vintage_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), vintage_filename)

# Always save the dedicated vintage file
prs.save(vintage_path)
print(f"Vintage deck saved to: {vintage_path}")

# Attempt to save to main presentation file, releasing PowerPoint lock if open
try:
    prs.save(output_path)
    print(f"Main presentation saved successfully to: {output_path}")
except PermissionError:
    # PowerPoint is open by the user; force close to release lock
    try:
        subprocess.run(["powershell", "-NoProfile", "-Command", "Stop-Process -Name 'POWERPNT' -Force -ErrorAction SilentlyContinue"], capture_output=True)
        time.sleep(0.8)
        prs.save(output_path)
        print(f"PowerPoint released and main presentation updated: {output_path}")
    except Exception as e:
        print(f"Notice: Main PPT file was locked ({e}). Updated {vintage_path} perfectly.")

print(f"Total slides in vintage deck: {len(prs.slides)}")
