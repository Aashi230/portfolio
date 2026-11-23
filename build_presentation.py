"""
build_presentation.py
Generates the comprehensive, professional PowerPoint presentation (.pptx) for
Aashi Srivastava's Portfolio: 'The Lucknow Gazette & Archive'.
"""

import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# Initialize Presentation
prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

# Color Palette Constants
COLOR_DARK_BG       = RGBColor(28, 21, 16)     # #1C1510 - Rich Espresso
COLOR_PARCHMENT     = RGBColor(250, 246, 238)  # #FAF6EE - Warm Antique Parchment
COLOR_CARD_BG       = RGBColor(245, 239, 229)  # #F5EFE5 - Slightly darker card parchment
COLOR_BURGUNDY      = RGBColor(107, 23, 36)    # #6B1724 - Royal Deep Burgundy
COLOR_BURGUNDY_LIGHT= RGBColor(243, 230, 232)  # #F3E6E8 - Light burgundy tint
COLOR_OLIVE         = RGBColor(74, 90, 63)     # #4A5A3F - Antique Olive / Moss
COLOR_OLIVE_LIGHT   = RGBColor(234, 238, 230)  # #EAEEE6 - Light olive tint
COLOR_GOLD          = RGBColor(200, 155, 60)   # #C89B3C - Antique Gold
COLOR_TEXT_DARK     = RGBColor(35, 24, 18)     # #231812 - Deep Espresso Ink
COLOR_TEXT_MUTED    = RGBColor(115, 102, 92)   # #73665C - Muted Charcoal / Brown
COLOR_TEXT_LIGHT    = RGBColor(246, 241, 232)  # #F6F1E8 - Parchment white for dark slides
COLOR_BORDER        = RGBColor(213, 199, 178)  # #D5C7B2 - Subtle Rule Border
COLOR_WHITE         = RGBColor(255, 255, 255)

FONT_SERIF = "Georgia"
FONT_SANS  = "Calibri"
FONT_MONO  = "Consolas"

blank_layout = prs.slide_layouts[6]

def set_slide_background(slide, color):
    bg_shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height
    )
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
    p_tag.font.size = Pt(10)
    p_tag.font.bold = True
    p_tag.font.color.rgb = COLOR_BURGUNDY
    p_tag.space_after = Pt(2)
    
    # Title
    p_title = tf.add_paragraph()
    p_title.text = title_text
    p_title.font.name = FONT_SERIF
    p_title.font.size = Pt(24)
    p_title.font.bold = True
    p_title.font.color.rgb = COLOR_TEXT_DARK
    p_title.space_after = Pt(2)
    
    # Subtitle
    if subtitle_text:
        p_sub = tf.add_paragraph()
        p_sub.text = subtitle_text
        p_sub.font.name = FONT_SANS
        p_sub.font.size = Pt(12)
        p_sub.font.italic = True
        p_sub.font.color.rgb = COLOR_TEXT_MUTED
    
    # Accent Rule Line
    rule = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.55), Inches(11.733), Inches(0.02)
    )
    rule.fill.solid()
    rule.fill.fore_color.rgb = COLOR_BORDER
    rule.line.fill.background()

def add_card_box(slide, left, top, width, height, bg_color=COLOR_CARD_BG, border_color=COLOR_BORDER):
    card = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    card.fill.solid()
    card.fill.fore_color.rgb = bg_color
    if border_color:
        card.line.color.rgb = border_color
        card.line.width = Pt(1.2)
    else:
        card.line.fill.background()
    return card

def add_image_card(slide, left, top, width, height, img_path, caption=""):
    # Outer frame
    frame = add_card_box(slide, left, top, width, height, bg_color=COLOR_WHITE, border_color=COLOR_BORDER)
    
    # Inset image slightly
    margin = Inches(0.08)
    img_left = left + margin
    img_top = top + margin
    img_w = width - (margin * 2)
    img_h = height - (margin * 2) - (Inches(0.35) if caption else 0)
    
    if os.path.exists(img_path):
        slide.shapes.add_picture(img_path, img_left, img_top, width=img_w, height=img_h)
    
    if caption:
        cap_box = slide.shapes.add_textbox(left + margin, top + height - Inches(0.38), width - (margin * 2), Inches(0.3))
        tf = cap_box.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
        p = tf.paragraphs[0]
        p.text = caption
        p.font.name = FONT_MONO
        p.font.size = Pt(9.5)
        p.font.bold = True
        p.font.color.rgb = COLOR_BURGUNDY
        p.alignment = PP_ALIGN.CENTER

# ==============================================================================
# SLIDE 1: TITLE SLIDE (GRAND EDITORIAL COVER)
# ==============================================================================
slide1 = prs.slides.add_slide(blank_layout)
set_slide_background(slide1, COLOR_DARK_BG)

# Decorative double border
outer_border = slide1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(0.5), Inches(12.333), Inches(6.5))
outer_border.fill.background()
outer_border.line.color.rgb = COLOR_GOLD
outer_border.line.width = Pt(1.5)

inner_border = slide1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.6), Inches(0.6), Inches(12.133), Inches(6.3))
inner_border.fill.background()
inner_border.line.color.rgb = COLOR_BURGUNDY
inner_border.line.width = Pt(0.75)

# Broadside Header Ribbon
tb_top = slide1.shapes.add_textbox(Inches(1.0), Inches(0.85), Inches(11.333), Inches(0.4))
tf_top = tb_top.text_frame
tf_top.margin_left = tf_top.margin_right = tf_top.margin_top = tf_top.margin_bottom = 0
p = tf_top.paragraphs[0]
p.text = "THE OFFICIAL ARCHIVE & PORTFOLIO COMPENDIUM • VOL. XXVI • ISSUE 01"
p.font.name = FONT_MONO
p.font.size = Pt(11)
p.font.bold = True
p.font.color.rgb = COLOR_GOLD
p.alignment = PP_ALIGN.CENTER

# Main Title Box
tb_title = slide1.shapes.add_textbox(Inches(1.0), Inches(1.35), Inches(11.333), Inches(2.3))
tf_title = tb_title.text_frame
tf_title.word_wrap = True
tf_title.margin_left = tf_title.margin_right = tf_title.margin_top = tf_title.margin_bottom = 0

p1 = tf_title.paragraphs[0]
p1.text = "The Lucknow Gazette"
p1.font.name = FONT_SERIF
p1.font.size = Pt(44)
p1.font.bold = True
p1.font.color.rgb = COLOR_TEXT_LIGHT
p1.alignment = PP_ALIGN.CENTER

p2 = tf_title.add_paragraph()
p2.text = "& CURATED ENGINEERING DOSSIER"
p2.font.name = FONT_SERIF
p2.font.size = Pt(20)
p2.font.bold = True
p2.font.color.rgb = COLOR_GOLD
p2.alignment = PP_ALIGN.CENTER
p2.space_before = Pt(4)

p3 = tf_title.add_paragraph()
p3.text = "A Complete Architectural Walkthrough of Aashi Srivastava's Portfolio Website"
p3.font.name = FONT_SANS
p3.font.size = Pt(14)
p3.font.color.rgb = RGBColor(180, 170, 160)
p3.alignment = PP_ALIGN.CENTER
p3.space_before = Pt(8)

# Center Profile + Credential Box
center_card = add_card_box(slide1, Inches(2.2), Inches(3.85), Inches(8.933), Inches(2.2), bg_color=RGBColor(38, 29, 23), border_color=COLOR_GOLD)

# Portrait inside card
portrait_path = "assets/images/profile/portrait.jpg"
if os.path.exists(portrait_path):
    slide1.shapes.add_picture(portrait_path, Inches(2.5), Inches(4.05), width=Inches(1.5), height=Inches(1.8))

tb_meta = slide1.shapes.add_textbox(Inches(4.2), Inches(4.05), Inches(6.7), Inches(1.8))
tf_meta = tb_meta.text_frame
tf_meta.word_wrap = True
tf_meta.margin_left = tf_meta.margin_right = tf_meta.margin_top = tf_meta.margin_bottom = 0

pm1 = tf_meta.paragraphs[0]
pm1.text = "AASHI SRIVASTAVA"
pm1.font.name = FONT_SERIF
pm1.font.size = Pt(20)
pm1.font.bold = True
pm1.font.color.rgb = COLOR_TEXT_LIGHT

pm2 = tf_meta.add_paragraph()
pm2.text = "B.Tech Computer Science Engineering (Artificial Intelligence & Data Science)"
pm2.font.name = FONT_SANS
pm2.font.size = Pt(12)
pm2.font.bold = True
pm2.font.color.rgb = COLOR_GOLD
pm2.space_before = Pt(2)

pm3 = tf_meta.add_paragraph()
pm3.text = "Shri Ramswaroop Memorial University (SRMU), Lucknow • CGPA: 8.29 • Class of 2029"
pm3.font.name = FONT_SANS
pm3.font.size = Pt(11)
pm3.font.color.rgb = RGBColor(200, 190, 180)
pm3.space_before = Pt(2)

pm4 = tf_meta.add_paragraph()
pm4.text = "Live Deployment: https://aashi230.github.io/portfolio/ • GitHub: github.com/Aashi230/portfolio"
pm4.font.name = FONT_MONO
pm4.font.size = Pt(10)
pm4.font.color.rgb = RGBColor(160, 200, 160)
pm4.space_before = Pt(8)

tb_foot = slide1.shapes.add_textbox(Inches(1.0), Inches(6.4), Inches(11.333), Inches(0.4))
p_foot = tb_foot.text_frame.paragraphs[0]
p_foot.text = "Aesthetic Doctrine: Vintage Editorial Luxury Meets Computational Rigor • Designed & Engineered with Pride"
p_foot.font.name = FONT_MONO
p_foot.font.size = Pt(9.5)
p_foot.font.color.rgb = COLOR_GOLD
p_foot.alignment = PP_ALIGN.CENTER

# ==============================================================================
# SLIDE 2: EXECUTIVE SUMMARY & DESIGN PHILOSOPHY
# ==============================================================================
slide2 = prs.slides.add_slide(blank_layout)
set_slide_background(slide2, COLOR_PARCHMENT)
add_header(slide2, "EXECUTIVE SUMMARY", "Portfolio Conception & Editorial Philosophy", 
           "How a bespoke 19th-century broadside identity merges with modern responsive engineering.")

# 3 Pillars Cards
col_w = Inches(3.64)
col_h = Inches(4.9)
gap = Inches(0.4)
top_pos = Inches(1.85)

# Pillar 1: The Design Vision
c1 = add_card_box(slide2, Inches(0.8), top_pos, col_w, col_h)
tb1 = slide2.shapes.add_textbox(Inches(1.05), top_pos + Inches(0.25), col_w - Inches(0.5), col_h - Inches(0.5))
tf1 = tb1.text_frame
tf1.word_wrap = True

p = tf1.paragraphs[0]
p.text = "01 / AESTHETIC DOCTRINE"
p.font.name = FONT_MONO
p.font.size = Pt(10)
p.font.bold = True
p.font.color.rgb = COLOR_BURGUNDY

p = tf1.add_paragraph()
p.text = "Vintage Editorial Luxury"
p.font.name = FONT_SERIF
p.font.size = Pt(18)
p.font.bold = True
p.font.color.rgb = COLOR_TEXT_DARK
p.space_before = Pt(4)

p = tf1.add_paragraph()
p.text = ("Unlike generic minimalist tech portfolios, this platform is styled as an antique broadsheet "
          "gazette ('The Lucknow Gazette & Archive'). It features authentic ornamental corner braces, "
          "deckled borders, warm parchment palettes, and classical serif typography reminiscent of vintage collegiate publications.")
p.font.name = FONT_SANS
p.font.size = Pt(11.5)
p.font.color.rgb = COLOR_TEXT_MUTED
p.space_before = Pt(8)

p = tf1.add_paragraph()
p.text = "• Antique Cream Parchment (#FAF6EE)\n• Royal Deep Burgundy (#6B1724)\n• Editorial Forest Olive (#4A5A3F)\n• Gold Foil & Brass Accent Trim"
p.font.name = FONT_MONO
p.font.size = Pt(10.5)
p.font.color.rgb = COLOR_BURGUNDY
p.space_before = Pt(12)

# Pillar 2: The Engineering Rigor
c2 = add_card_box(slide2, Inches(0.8) + col_w + gap, top_pos, col_w, col_h)
tb2 = slide2.shapes.add_textbox(Inches(1.05) + col_w + gap, top_pos + Inches(0.25), col_w - Inches(0.5), col_h - Inches(0.5))
tf2 = tb2.text_frame
tf2.word_wrap = True

p = tf2.paragraphs[0]
p.text = "02 / COMPUTATIONAL RIGOR"
p.font.name = FONT_MONO
p.font.size = Pt(10)
p.font.bold = True
p.font.color.rgb = COLOR_OLIVE

p = tf2.add_paragraph()
p.text = "Student & Tech Learner"
p.font.name = FONT_SERIF
p.font.size = Pt(18)
p.font.bold = True
p.font.color.rgb = COLOR_TEXT_DARK
p.space_before = Pt(4)

p = tf2.add_paragraph()
p.text = ("Built specifically to showcase Aashi's journey as an undergraduate Student & Curious Tech Learner in AI & Data Science. "
          "The portfolio highlights continuous experimentation, foundational programming, algorithmic problem-solving, "
          "and certified academic distinction (SRMU Lucknow CGPA 8.29).")
p.font.name = FONT_SANS
p.font.size = Pt(11.5)
p.font.color.rgb = COLOR_TEXT_MUTED
p.space_before = Pt(8)

p = tf2.add_paragraph()
p.text = "• Opportunity Hunter (AI aggregator)\n• ARS Tourist Guide (SIH 2026 Selection)\n• Campus Hostel Grievance Pipeline\n• Digital Voting Machine Logic"
p.font.name = FONT_MONO
p.font.size = Pt(10.5)
p.font.color.rgb = COLOR_OLIVE
p.space_before = Pt(12)

# Pillar 3: Web Performance
c3 = add_card_box(slide2, Inches(0.8) + (col_w + gap)*2, top_pos, col_w, col_h)
tb3 = slide2.shapes.add_textbox(Inches(1.05) + (col_w + gap)*2, top_pos + Inches(0.25), col_w - Inches(0.5), col_h - Inches(0.5))
tf3 = tb3.text_frame
tf3.word_wrap = True

p = tf3.paragraphs[0]
p.text = "03 / WEB ARCHITECTURE"
p.font.name = FONT_MONO
p.font.size = Pt(10)
p.font.bold = True
p.font.color.rgb = COLOR_GOLD

p = tf3.add_paragraph()
p.text = "Zero-Bloat Performance"
p.font.name = FONT_SERIF
p.font.size = Pt(18)
p.font.bold = True
p.font.color.rgb = COLOR_TEXT_DARK
p.space_before = Pt(4)

p = tf3.add_paragraph()
p.text = ("Engineered strictly in semantic HTML5, modular CSS custom properties, and Vanilla ES6+ JavaScript. "
          "Free of heavy frontend frameworks (React/Vue/Webpack), ensuring instantaneous sub-second load times, "
          "zero layout shifts, and flawless responsive execution on all mobile and desktop devices.")
p.font.name = FONT_SANS
p.font.size = Pt(11.5)
p.font.color.rgb = COLOR_TEXT_MUTED
p.space_before = Pt(8)

p = tf3.add_paragraph()
p.text = "• 100% Mobile & Tablet Responsive\n• Sub-second Cold Page Load\n• Live Asynchronous AJAX Dispatch\n• Static CDN on GitHub Pages"
p.font.name = FONT_MONO
p.font.size = Pt(10.5)
p.font.color.rgb = COLOR_TEXT_DARK
p.space_before = Pt(12)

# ==============================================================================
# SLIDE 3: TECH STACK ARCHITECTURE
# ==============================================================================
slide3 = prs.slides.add_slide(blank_layout)
set_slide_background(slide3, COLOR_PARCHMENT)
add_header(slide3, "TECHNICAL BLUEPRINT", "Full-Stack Technology Stack & Architecture",
           "Modular, modern web standards working in harmony with zero runtime overhead.")

# 4 Architectural Quadrants
grid_w = Inches(5.6)
grid_h = Inches(2.35)

# Quad 1: Frontend Core
add_card_box(slide3, Inches(0.8), Inches(1.85), grid_w, grid_h)
t1 = slide3.shapes.add_textbox(Inches(1.0), Inches(1.95), grid_w - Inches(0.4), grid_h - Inches(0.2))
f1 = t1.text_frame
f1.word_wrap = True
p = f1.paragraphs[0]
p.text = "SEMANTIC MARKUP & STYLING ARCHITECTURE"
p.font.name = FONT_MONO
p.font.size = Pt(10)
p.font.bold = True
p.font.color.rgb = COLOR_BURGUNDY
p = f1.add_paragraph()
p.text = "HTML5 & Modular CSS3 (style.css)"
p.font.name = FONT_SERIF
p.font.size = Pt(16)
p.font.bold = True
p.font.color.rgb = COLOR_TEXT_DARK
p = f1.add_paragraph()
p.text = ("• Over 3,000 lines of handcrafted, modular CSS with design tokens (:root variables).\n"
          "• Semantic HTML5 structure (<nav>, <header>, <section>, <article>, <figure>, <footer>).\n"
          "• CSS Grid & Flexbox layouts with fluid clamp() typography and multi-column archives.\n"
          "• Custom antique corner accents (.vintage-corner) and double-ruled borders.")
p.font.name = FONT_SANS
p.font.size = Pt(11)
p.font.color.rgb = COLOR_TEXT_MUTED
p.space_before = Pt(4)

# Quad 2: Interactive Logic
add_card_box(slide3, Inches(6.933), Inches(1.85), grid_w, grid_h)
t2 = slide3.shapes.add_textbox(Inches(7.133), Inches(1.95), grid_w - Inches(0.4), grid_h - Inches(0.2))
f2 = t2.text_frame
f2.word_wrap = True
p = f2.paragraphs[0]
p.text = "CLIENT-SIDE LOGIC & STATE ENGINE"
p.font.name = FONT_MONO
p.font.size = Pt(10)
p.font.bold = True
p.font.color.rgb = COLOR_OLIVE
p = f2.add_paragraph()
p.text = "Vanilla JavaScript ES6+ (main.js)"
p.font.name = FONT_SERIF
p.font.size = Pt(16)
p.font.bold = True
p.font.color.rgb = COLOR_TEXT_DARK
p = f2.add_paragraph()
p.text = ("• Custom event-driven architecture with zero npm/external dependencies.\n"
          "• Dynamic mobile navigation drawer with backdrop blur and touch toggling.\n"
          "• Active section tracking using IntersectionObserver API for real-time nav indicators.\n"
          "• Smooth scroll handling and interactive tab switching across project chapters.")
p.font.name = FONT_SANS
p.font.size = Pt(11)
p.font.color.rgb = COLOR_TEXT_MUTED
p.space_before = Pt(4)

# Quad 3: Form Transmission API
add_card_box(slide3, Inches(0.8), Inches(4.45), grid_w, grid_h)
t3 = slide3.shapes.add_textbox(Inches(1.0), Inches(4.55), grid_w - Inches(0.4), grid_h - Inches(0.2))
f3 = t3.text_frame
f3.word_wrap = True
p = f3.paragraphs[0]
p.text = "COMMUNICATION & API INTEGRATION"
p.font.name = FONT_MONO
p.font.size = Pt(10)
p.font.bold = True
p.font.color.rgb = COLOR_GOLD
p = f3.add_paragraph()
p.text = "FormSubmit AJAX API & Email Pipeline"
p.font.name = FONT_SERIF
p.font.size = Pt(16)
p.font.bold = True
p.font.color.rgb = COLOR_TEXT_DARK
p = f3.add_paragraph()
p.text = ("• Asynchronous fetch() dispatch transmitting directly to aashimaahi23@gmail.com.\n"
          "• Configured with _captcha: false for seamless non-blocking background dispatch.\n"
          "• State machine button updates: 'Transmit Dispatch' → 'Transmitting...' → 'Confirmed'.\n"
          "• Graceful error fallbacks and activation notice detection.")
p.font.name = FONT_SANS
p.font.size = Pt(11)
p.font.color.rgb = COLOR_TEXT_MUTED
p.space_before = Pt(4)

# Quad 4: Deployment & CI/CD
add_card_box(slide3, Inches(6.933), Inches(4.45), grid_w, grid_h)
t4 = slide3.shapes.add_textbox(Inches(7.133), Inches(4.55), grid_w - Inches(0.4), grid_h - Inches(0.2))
f4 = t4.text_frame
f4.word_wrap = True
p = f4.paragraphs[0]
p.text = "HOSTING, CDN & VERSION CONTROL"
p.font.name = FONT_MONO
p.font.size = Pt(10)
p.font.bold = True
p.font.color.rgb = COLOR_BURGUNDY
p = f4.add_paragraph()
p.text = "Git & GitHub Pages Production Pipeline"
p.font.name = FONT_SERIF
p.font.size = Pt(16)
p.font.bold = True
p.font.color.rgb = COLOR_TEXT_DARK
p = f4.add_paragraph()
p.text = ("• Hosted globally on GitHub Pages edge CDN: https://aashi230.github.io/portfolio/.\n"
          "• Clean atomic Git commit history (latest commit e0dd2cf) pushed directly to main.\n"
          "• Built-in asset optimization with fast local static asset resolution.\n"
          "• Fully automated deployment triggered on every Git push.")
p.font.name = FONT_SANS
p.font.size = Pt(11)
p.font.color.rgb = COLOR_TEXT_MUTED
p.space_before = Pt(4)

# ==============================================================================
# SLIDE 4: LANDING & MASTHEAD
# ==============================================================================
slide4 = prs.slides.add_slide(blank_layout)
set_slide_background(slide4, COLOR_PARCHMENT)
add_header(slide4, "CHAPTER 00 / PORTFOLIO FRONTISPIECE", "The Grand Broadside Masthead & Navigation",
           "First impressions inspired by classical 19th-century academic and press broadsheets.")

# Left Column: Features & Explanations
left_w = Inches(5.6)
top_pos = Inches(1.85)

add_card_box(slide4, Inches(0.8), top_pos, left_w, Inches(4.9))
t_lh = slide4.shapes.add_textbox(Inches(1.05), top_pos + Inches(0.25), left_w - Inches(0.5), Inches(4.4))
f_lh = t_lh.text_frame
f_lh.word_wrap = True

p = f_lh.paragraphs[0]
p.text = "KEY ARCHITECTURAL ELEMENTS"
p.font.name = FONT_MONO
p.font.size = Pt(10)
p.font.bold = True
p.font.color.rgb = COLOR_BURGUNDY

p = f_lh.add_paragraph()
p.text = "1. Editorial Date & Volume Bar"
p.font.name = FONT_SERIF
p.font.size = Pt(14)
p.font.bold = True
p.font.color.rgb = COLOR_TEXT_DARK
p.space_before = Pt(6)

p = f_lh.add_paragraph()
p.text = ("Displays authentic gazette metadata: 'VOL. XXVI • NO. 01 • LUCKNOW, INDIA'. "
          "Establishes a distinctive scholarly gravitas from the very first glance.")
p.font.name = FONT_SANS
p.font.size = Pt(11)
p.font.color.rgb = COLOR_TEXT_MUTED

p = f_lh.add_paragraph()
p.text = "2. AS Monogram Seal & Dual Badging"
p.font.name = FONT_SERIF
p.font.size = Pt(14)
p.font.bold = True
p.font.color.rgb = COLOR_TEXT_DARK
p.space_before = Pt(8)

p = f_lh.add_paragraph()
p.text = ("A customized circular monogram emblem ('AS') positioned within the header bar, "
          "accompanied by status tags denoting 'STUDENT SCHOLAR & BUILDER' and current semester status.")
p.font.name = FONT_SANS
p.font.size = Pt(11)
p.font.color.rgb = COLOR_TEXT_MUTED

p = f_lh.add_paragraph()
p.text = "3. Sticky Responsive Navigation"
p.font.name = FONT_SERIF
p.font.size = Pt(14)
p.font.bold = True
p.font.color.rgb = COLOR_TEXT_DARK
p.space_before = Pt(8)

p = f_lh.add_paragraph()
p.text = ("Fixed header bar featuring a backdrop blur filter, chapter anchor links (#about, #projects, "
          "#skills, #resume, #contact), and an expandable mobile hamburger drawer for touchscreens.")
p.font.name = FONT_SANS
p.font.size = Pt(11)
p.font.color.rgb = COLOR_TEXT_MUTED

# Right Column: Screenshot Card
add_image_card(slide4, Inches(6.8), top_pos, Inches(5.733), Inches(4.9), 
               "presentation_assets/01_hero_landing.png", 
               "LIVE CAPTURE: Broadside Masthead & Lead Dispatch Banner")

# ==============================================================================
# SLIDE 5: HERO SECTION (THE LEAD EDITORIAL)
# ==============================================================================
slide5 = prs.slides.add_slide(blank_layout)
set_slide_background(slide5, COLOR_PARCHMENT)
add_header(slide5, "HERO SECTION WALKTHROUGH", "The Lead Editorial & Identity Dispatch",
           "Balancing personal storytelling with high-impact data science metrics.")

add_card_box(slide5, Inches(0.8), top_pos, left_w, Inches(4.9))
t_hero = slide5.shapes.add_textbox(Inches(1.05), top_pos + Inches(0.25), left_w - Inches(0.5), Inches(4.4))
f_hero = t_hero.text_frame
f_hero.word_wrap = True

p = f_hero.paragraphs[0]
p.text = "HERO CONTENT & CALLS-TO-ACTION"
p.font.name = FONT_MONO
p.font.size = Pt(10)
p.font.bold = True
p.font.color.rgb = COLOR_BURGUNDY

p = f_hero.add_paragraph()
p.text = "The Guiding Headline"
p.font.name = FONT_SERIF
p.font.size = Pt(15)
p.font.bold = True
p.font.color.rgb = COLOR_TEXT_DARK
p.space_before = Pt(4)

p = f_hero.add_paragraph()
p.text = ("'Curious by nature, learning by doing, and growing with every experience.' "
          "Positions Aashi as an ambitious undergraduate technologist committed to hands-on experimentation.")
p.font.name = FONT_SANS
p.font.size = Pt(11)
p.font.color.rgb = COLOR_TEXT_MUTED

p = f_hero.add_paragraph()
p.text = "Primary Interactive Actions (CTAs)"
p.font.name = FONT_SERIF
p.font.size = Pt(15)
p.font.bold = True
p.font.color.rgb = COLOR_TEXT_DARK
p.space_before = Pt(8)

p = f_hero.add_paragraph()
p.text = ("• 'Explore Archives': Smoothly scrolls visitor into featured project case studies.\n"
          "• 'Download Official PDF': Direct one-click download of Aashi Srivastava's vetted resume.\n"
          "• 'Send a Dispatch': Instant anchor jump to live message transmission form.")
p.font.name = FONT_SANS
p.font.size = Pt(11)
p.font.color.rgb = COLOR_TEXT_MUTED

p = f_hero.add_paragraph()
p.text = "At-A-Glance Quantitative Metrics"
p.font.name = FONT_SERIF
p.font.size = Pt(15)
p.font.bold = True
p.font.color.rgb = COLOR_TEXT_DARK
p.space_before = Pt(8)

p = f_hero.add_paragraph()
p.text = ("• 8.29 CGPA Academic Distinction at SRMU Lucknow\n"
          "• 4+ Flagship Engineering Projects (AI, Hackathons, Web)\n"
          "• Smart India Hackathon (SIH 2026) Selection\n"
          "• GDG DevFest Builder & Android Nova Certified")
p.font.name = FONT_MONO
p.font.size = Pt(10.5)
p.font.color.rgb = COLOR_OLIVE

# Right Column: Screenshot Card
add_image_card(slide5, Inches(6.8), top_pos, Inches(5.733), Inches(4.9), 
               "presentation_assets/01_hero_landing.png", 
               "SECTION ARCHITECTURE: Lead Headline, Metrics Ticker & CTAs")

# ==============================================================================
# SLIDE 6: ABOUT SECTION & ACADEMIC DOSSIER
# ==============================================================================
slide6 = prs.slides.add_slide(blank_layout)
set_slide_background(slide6, COLOR_PARCHMENT)
add_header(slide6, "CHAPTER 01 / PROFILE & BACKGROUND", "The Academic Dossier & Scholarly Standing",
           "Verified collegiate credentials, university ranking, and community leadership.")

add_card_box(slide6, Inches(0.8), top_pos, left_w, Inches(4.9))
t_abt = slide6.shapes.add_textbox(Inches(1.05), top_pos + Inches(0.25), left_w - Inches(0.5), Inches(4.4))
f_abt = t_abt.text_frame
f_abt.word_wrap = True

p = f_abt.paragraphs[0]
p.text = "ACADEMIC PROFILE & IDENTITY"
p.font.name = FONT_MONO
p.font.size = Pt(10)
p.font.bold = True
p.font.color.rgb = COLOR_BURGUNDY

p = f_abt.add_paragraph()
p.text = "Institutional Standing"
p.font.name = FONT_SERIF
p.font.size = Pt(14)
p.font.bold = True
p.font.color.rgb = COLOR_TEXT_DARK
p.space_before = Pt(4)

p = f_abt.add_paragraph()
p.text = ("• Institution: Shri Ramswaroop Memorial University (SRMU), Lucknow\n"
          "• Degree: Bachelor of Technology (B.Tech) in Computer Science & Engineering\n"
          "• Specialization: Artificial Intelligence & Data Science\n"
          "• Academic Performance: 8.29 CGPA (Class of 2029, 3rd Semester)")
p.font.name = FONT_SANS
p.font.size = Pt(11)
p.font.color.rgb = COLOR_TEXT_MUTED

p = f_abt.add_paragraph()
p.text = "Dual-Column Dossier Architecture"
p.font.name = FONT_SERIF
p.font.size = Pt(14)
p.font.bold = True
p.font.color.rgb = COLOR_TEXT_DARK
p.space_before = Pt(8)

p = f_abt.add_paragraph()
p.text = ("• Left Dossier Card: In-depth narrative of Aashi's analytical journey from fundamental C/C++ "
          "algorithms to statistical machine learning and data pipeline construction.\n"
          "• Right Identity Card: Features Aashi's student credential badge, official ID number, and verified "
          "passports from national events (GDG DevFest, Smart India Hackathon).")
p.font.name = FONT_SANS
p.font.size = Pt(11)
p.font.color.rgb = COLOR_TEXT_MUTED

p = f_abt.add_paragraph()
p.text = "Student Journey & Learning Focus"
p.font.name = FONT_SERIF
p.font.size = Pt(14)
p.font.bold = True
p.font.color.rgb = COLOR_TEXT_DARK
p.space_before = Pt(8)

p = f_abt.add_paragraph()
p.text = ("Dedicated to growing as a curious tech learner and engineering student—exploring "
          "core algorithms, artificial intelligence fundamentals, data science, and building practical software tools.")
p.font.name = FONT_SANS
p.font.size = Pt(11)
p.font.color.rgb = COLOR_TEXT_MUTED

# Right Column: Screenshot Card
add_image_card(slide6, Inches(6.8), top_pos, Inches(5.733), Inches(4.9), 
               "presentation_assets/02_about_dossier.png", 
               "LIVE CAPTURE: Chapter 01 Academic Dossier & Credentials")

# ==============================================================================
# SLIDE 7: MY WORK / FEATURED PROJECTS ARCHIVE
# ==============================================================================
slide7 = prs.slides.add_slide(blank_layout)
set_slide_background(slide7, COLOR_PARCHMENT)
add_header(slide7, "CHAPTER 03 / SELECTED WORKS", "Curated Engineering Projects & Repositories",
           "Four comprehensive engineering systems solving real-world challenges.")

add_card_box(slide7, Inches(0.8), top_pos, left_w, Inches(4.9))
t_prj = slide7.shapes.add_textbox(Inches(1.05), top_pos + Inches(0.25), left_w - Inches(0.5), Inches(4.4))
f_prj = t_prj.text_frame
f_prj.word_wrap = True

p = f_prj.paragraphs[0]
p.text = "FLAGSHIP SYSTEM CASE STUDIES"
p.font.name = FONT_MONO
p.font.size = Pt(10)
p.font.bold = True
p.font.color.rgb = COLOR_BURGUNDY

p = f_prj.add_paragraph()
p.text = "1. Opportunity Hunter (Discovery Platform)"
p.font.name = FONT_SERIF
p.font.size = Pt(13.5)
p.font.bold = True
p.font.color.rgb = COLOR_TEXT_DARK
p.space_before = Pt(4)

p = f_prj.add_paragraph()
p.text = ("An intelligent aggregator consolidating student opportunities: scholarships, internships, "
          "and hackathons with eligibility filtering and automatic calendar tracking.\n"
          "Tech: Python, Pandas, Web Scraping, REST APIs.")
p.font.name = FONT_SANS
p.font.size = Pt(10.5)
p.font.color.rgb = COLOR_TEXT_MUTED

p = f_prj.add_paragraph()
p.text = "2. ARS Tourist Guide (Smart India Hackathon 2026)"
p.font.name = FONT_SERIF
p.font.size = Pt(13.5)
p.font.bold = True
p.font.color.rgb = COLOR_TEXT_DARK
p.space_before = Pt(6)

p = f_prj.add_paragraph()
p.text = ("Heritage tourism intelligence system selected for SIH 2026. Delivers historical contextual audio, "
          "multilingual translation, and route safety optimization.\n"
          "Tech: AI/NLP, Geolocation, Mobile Responsive Web.")
p.font.name = FONT_SANS
p.font.size = Pt(10.5)
p.font.color.rgb = COLOR_TEXT_MUTED

p = f_prj.add_paragraph()
p.text = "3. Hostel Complaint Management System"
p.font.name = FONT_SERIF
p.font.size = Pt(13.5)
p.font.bold = True
p.font.color.rgb = COLOR_TEXT_DARK
p.space_before = Pt(6)

p = f_prj.add_paragraph()
p.text = ("Campus operations portal replacing manual paperwork with automated ticket dispatch, "
          "priority SLAs, and warden verification dashboards.")
p.font.name = FONT_SANS
p.font.size = Pt(10.5)
p.font.color.rgb = COLOR_TEXT_MUTED

# Right Column: Screenshot Card
add_image_card(slide7, Inches(6.8), top_pos, Inches(5.733), Inches(4.9), 
               "presentation_assets/03_projects_work.png", 
               "LIVE CAPTURE: Chapter 03 Curated Engineering Works")

# ==============================================================================
# SLIDE 8: TECHNICAL ARSENAL & SKILLS MATRIX
# ==============================================================================
slide8 = prs.slides.add_slide(blank_layout)
set_slide_background(slide8, COLOR_PARCHMENT)
add_header(slide8, "CHAPTER 04 / TECHNICAL ARSENAL", "Comprehensive Competency & Skills Matrix",
           "Full inventory of core languages, data science frameworks, and developer workflows.")

# 4 Competency Cards in Grid
sk_w = Inches(5.6)
sk_h = Inches(2.35)

# Card 1: Programming Languages
add_card_box(slide8, Inches(0.8), Inches(1.85), sk_w, sk_h)
tb_s1 = slide8.shapes.add_textbox(Inches(1.0), Inches(1.95), sk_w - Inches(0.4), sk_h - Inches(0.2))
tf_s1 = tb_s1.text_frame
tf_s1.word_wrap = True
p = tf_s1.paragraphs[0]
p.text = "CORE PROGRAMMING LANGUAGES"
p.font.name = FONT_MONO
p.font.size = Pt(10)
p.font.bold = True
p.font.color.rgb = COLOR_BURGUNDY
p = tf_s1.add_paragraph()
p.text = "Python • C • C++ • SQL • JavaScript"
p.font.name = FONT_SERIF
p.font.size = Pt(16)
p.font.bold = True
p.font.color.rgb = COLOR_TEXT_DARK
p = tf_s1.add_paragraph()
p.text = ("• Strong algorithmic foundation in C/C++ (certified by SRMU CSE Department).\n"
          "• Python expertise for data structures, automation, statistical modeling, and scripting.\n"
          "• Relational database querying with SQL (joins, subqueries, aggregations).\n"
          "• Modern ES6+ JavaScript for DOM manipulation and asynchronous APIs.")
p.font.name = FONT_SANS
p.font.size = Pt(11)
p.font.color.rgb = COLOR_TEXT_MUTED
p.space_before = Pt(4)

# Card 2: AI & Data Science
add_card_box(slide8, Inches(6.933), Inches(1.85), sk_w, sk_h)
tb_s2 = slide8.shapes.add_textbox(Inches(7.133), Inches(1.95), sk_w - Inches(0.4), sk_h - Inches(0.2))
tf_s2 = tb_s2.text_frame
tf_s2.word_wrap = True
p = tf_s2.paragraphs[0]
p.text = "DATA ANALYTICS & ARTIFICIAL INTELLIGENCE"
p.font.name = FONT_MONO
p.font.size = Pt(10)
p.font.bold = True
p.font.color.rgb = COLOR_OLIVE
p = tf_s2.add_paragraph()
p.text = "Pandas • NumPy • Matplotlib • Seaborn"
p.font.name = FONT_SERIF
p.font.size = Pt(16)
p.font.bold = True
p.font.color.rgb = COLOR_TEXT_DARK
p = tf_s2.add_paragraph()
p.text = ("• Exploratory Data Analysis (EDA): Data cleaning, missing value handling, outliers.\n"
          "• Data visualization: Statistical distribution charts, heatmaps, multi-axis trends.\n"
          "• Machine Learning basics: Regression, classification, feature engineering.\n"
          "• Generative AI & Prompt Engineering (Be10x Certified).")
p.font.name = FONT_SANS
p.font.size = Pt(11)
p.font.color.rgb = COLOR_TEXT_MUTED
p.space_before = Pt(4)

# Card 3: Web Engineering
add_card_box(slide8, Inches(0.8), Inches(4.45), sk_w, sk_h)
tb_s3 = slide8.shapes.add_textbox(Inches(1.0), Inches(4.55), sk_w - Inches(0.4), sk_h - Inches(0.2))
tf_s3 = tb_s3.text_frame
tf_s3.word_wrap = True
p = tf_s3.paragraphs[0]
p.text = "WEB & INTERACTION ENGINEERING"
p.font.name = FONT_MONO
p.font.size = Pt(10)
p.font.bold = True
p.font.color.rgb = COLOR_GOLD
p = tf_s3.add_paragraph()
p.text = "HTML5 • Modern CSS3 • REST APIs"
p.font.name = FONT_SERIF
p.font.size = Pt(16)
p.font.bold = True
p.font.color.rgb = COLOR_TEXT_DARK
p = tf_s3.add_paragraph()
p.text = ("• High-fidelity responsive web layout engineering (Desktop, Tablet, Mobile).\n"
          "• CSS custom properties, grid architectures, animations, deckled borders.\n"
          "• Asynchronous HTTP requests via fetch() with JSON payloads.\n"
          "• Accessible UI/UX design with semantic tags and keyboard navigation.")
p.font.name = FONT_SANS
p.font.size = Pt(11)
p.font.color.rgb = COLOR_TEXT_MUTED
p.space_before = Pt(4)

# Card 4: Developer Tools
add_card_box(slide8, Inches(6.933), Inches(4.45), sk_w, sk_h)
tb_s4 = slide8.shapes.add_textbox(Inches(7.133), Inches(4.55), sk_w - Inches(0.4), sk_h - Inches(0.2))
tf_s4 = tb_s4.text_frame
tf_s4.word_wrap = True
p = tf_s4.paragraphs[0]
p.text = "DEVELOPER WORKFLOWS & PLATFORMS"
p.font.name = FONT_MONO
p.font.size = Pt(10)
p.font.bold = True
p.font.color.rgb = COLOR_BURGUNDY
p = tf_s4.add_paragraph()
p.text = "Git • GitHub • VS Code • Jupyter • Colab"
p.font.name = FONT_SERIF
p.font.size = Pt(16)
p.font.bold = True
p.font.color.rgb = COLOR_TEXT_DARK
p = tf_s4.add_paragraph()
p.text = ("• Git version control: Branching, merging, atomic commits, remote sync.\n"
          "• Cloud notebooks: Jupyter, Google Colab for reproducible ML workflows.\n"
          "• Visual Studio Code environment with linting, formatting, and live debugging.\n"
          "• GitHub Pages hosting with automated deployment pipelines.")
p.font.name = FONT_SANS
p.font.size = Pt(11)
p.font.color.rgb = COLOR_TEXT_MUTED
p.space_before = Pt(4)

# ==============================================================================
# SLIDE 9: CURRICULUM VITAE & RECENT VISUAL FIX
# ==============================================================================
slide9 = prs.slides.add_slide(blank_layout)
set_slide_background(slide9, COLOR_PARCHMENT)
add_header(slide9, "CHAPTER 06 / CURRICULUM VITAE", "The Professional CV & Layout Polish",
           "Clean editorial parchment presentation with direct official PDF resume download.")

add_card_box(slide9, Inches(0.8), top_pos, left_w, Inches(4.9))
t_cv = slide9.shapes.add_textbox(Inches(1.05), top_pos + Inches(0.25), left_w - Inches(0.5), Inches(4.4))
f_cv = t_cv.text_frame
f_cv.word_wrap = True

p = f_cv.paragraphs[0]
p.text = "SECTION ARCHITECTURE & KEY FIX"
p.font.name = FONT_MONO
p.font.size = Pt(10)
p.font.bold = True
p.font.color.rgb = COLOR_BURGUNDY

p = f_cv.add_paragraph()
p.text = "1. Clean Parchment Card Layout"
p.font.name = FONT_SERIF
p.font.size = Pt(14)
p.font.bold = True
p.font.color.rgb = COLOR_TEXT_DARK
p.space_before = Pt(4)

p = f_cv.add_paragraph()
p.text = ("Presented as an authentic academic document framed by an antique double border "
          "(dark brown outer border + inner burgundy outline with ornate brass corner brackets).")
p.font.name = FONT_SANS
p.font.size = Pt(11)
p.font.color.rgb = COLOR_TEXT_MUTED

p = f_cv.add_paragraph()
p.text = "2. Recent Visual Optimization (Fix Implemented)"
p.font.name = FONT_SERIF
p.font.size = Pt(14)
p.font.bold = True
p.font.color.rgb = COLOR_BURGUNDY
p.space_before = Pt(8)

p = f_cv.add_paragraph()
p.text = ("• Problem Identified: Previous 32px repeating background ruled notebook lines caused "
          "text and headings of varying sizes to collide or sit awkwardly across the lines.\n"
          "• Solution Applied: Removed linear-gradient stripes from .resume-sheet-card in style.css. "
          "Restored smooth antique cream card surface with perfect vertical text alignment (commit e0dd2cf).")
p.font.name = FONT_SANS
p.font.size = Pt(11)
p.font.color.rgb = COLOR_TEXT_MUTED

p = f_cv.add_paragraph()
p.text = "3. Integrated PDF Resume Download"
p.font.name = FONT_SERIF
p.font.size = Pt(14)
p.font.bold = True
p.font.color.rgb = COLOR_TEXT_DARK
p.space_before = Pt(8)

p = f_cv.add_paragraph()
p.text = ("An official 'Download Official PDF' button linking directly to Aashi_Srivastava_Resume.pdf "
          "with native HTML5 download attribute for immediate offline review.")
p.font.name = FONT_SANS
p.font.size = Pt(11)
p.font.color.rgb = COLOR_TEXT_MUTED

# Right Column: Screenshot Card
add_image_card(slide9, Inches(6.8), top_pos, Inches(5.733), Inches(4.9), 
               "presentation_assets/04_resume_cv.png", 
               "LIVE CAPTURE: Clean Parchment Curriculum Vitae (Commit e0dd2cf)")

# ==============================================================================
# SLIDE 10: CONTACT & LIVE MESSAGE DISPATCH
# ==============================================================================
slide10 = prs.slides.add_slide(blank_layout)
set_slide_background(slide10, COLOR_PARCHMENT)
add_header(slide10, "CHAPTER 07 / CORRESPONDENCE", "Live Dispatch System & Email Transmission",
           "Instant communication pipeline connecting visitors directly to Aashi's inbox.")

add_card_box(slide10, Inches(0.8), top_pos, left_w, Inches(4.9))
t_cnt = slide10.shapes.add_textbox(Inches(1.05), top_pos + Inches(0.25), left_w - Inches(0.5), Inches(4.4))
f_cnt = t_cnt.text_frame
f_cnt.word_wrap = True

p = f_cnt.paragraphs[0]
p.text = "COMMUNICATION PIPELINE & UX"
p.font.name = FONT_MONO
p.font.size = Pt(10)
p.font.bold = True
p.font.color.rgb = COLOR_BURGUNDY

p = f_cnt.add_paragraph()
p.text = "1. Live FormSubmit AJAX Integration"
p.font.name = FONT_SERIF
p.font.size = Pt(14)
p.font.bold = True
p.font.color.rgb = COLOR_TEXT_DARK
p.space_before = Pt(4)

p = f_cnt.add_paragraph()
p.text = ("Integrated with FormSubmit API targeting aashimaahi23@gmail.com with _captcha: false. "
          "Users submit messages without jarring page reloads or intrusive third-party redirects.")
p.font.name = FONT_SANS
p.font.size = Pt(11)
p.font.color.rgb = COLOR_TEXT_MUTED

p = f_cnt.add_paragraph()
p.text = "2. Real-Time State Machine Feedback"
p.font.name = FONT_SERIF
p.font.size = Pt(14)
p.font.bold = True
p.font.color.rgb = COLOR_TEXT_DARK
p.space_before = Pt(8)

p = f_cnt.add_paragraph()
p.text = ("The submit button dynamically transitions between three distinct operational states:\n"
          "  [1] Default: 'Transmit Dispatch' (Burgundy outline)\n"
          "  [2] Pending: 'Transmitting Dispatch...' (Spinner & disabled state)\n"
          "  [3] Success: 'Transmission Confirmed ✓' (Forest green status alert)")
p.font.name = FONT_SANS
p.font.size = Pt(10.5)
p.font.color.rgb = COLOR_TEXT_MUTED

p = f_cnt.add_paragraph()
p.text = "3. Direct Social & Professional Channels"
p.font.name = FONT_SERIF
p.font.size = Pt(14)
p.font.bold = True
p.font.color.rgb = COLOR_TEXT_DARK
p.space_before = Pt(8)

p = f_cnt.add_paragraph()
p.text = ("Direct verified cards for LinkedIn (professional networking), GitHub (code repositories), "
          "and mailto:aashimaahi23@gmail.com for direct email correspondence.")
p.font.name = FONT_SANS
p.font.size = Pt(11)
p.font.color.rgb = COLOR_TEXT_MUTED

# Right Column: Screenshot Card
add_image_card(slide10, Inches(6.8), top_pos, Inches(5.733), Inches(4.9), 
               "presentation_assets/05_contact_footer.png", 
               "LIVE CAPTURE: Chapter 07 Correspondence & Live Dispatch Form")

# ==============================================================================
# SLIDE 11: FOOTER, COLOPHON & RESPONSIVE ASSURANCE
# ==============================================================================
slide11 = prs.slides.add_slide(blank_layout)
set_slide_background(slide11, COLOR_PARCHMENT)
add_header(slide11, "FOOTER & COLOPHON", "The Editorial Colophon & Cross-Device Rigor",
           "Attention to archival detail, typographic credits, and responsive fluid design.")

# Left Card: Colophon Details
add_card_box(slide11, Inches(0.8), top_pos, Inches(5.6), Inches(4.9))
t_col = slide11.shapes.add_textbox(Inches(1.05), top_pos + Inches(0.25), Inches(5.1), Inches(4.4))
f_col = t_col.text_frame
f_col.word_wrap = True

p = f_col.paragraphs[0]
p.text = "THE EDITORIAL COLOPHON"
p.font.name = FONT_MONO
p.font.size = Pt(10)
p.font.bold = True
p.font.color.rgb = COLOR_BURGUNDY

p = f_col.add_paragraph()
p.text = "Gazette Copyright & Publication Notes"
p.font.name = FONT_SERIF
p.font.size = Pt(15)
p.font.bold = True
p.font.color.rgb = COLOR_TEXT_DARK
p.space_before = Pt(4)

p = f_col.add_paragraph()
p.text = ("The website concludes with an authentic newspaper colophon stating:\n"
          "'Published & Curated by Aashi Srivastava. Printed digitally in Lucknow, India. "
          "Set in classical serif typefaces and modern computational monospace.'")
p.font.name = FONT_SANS
p.font.size = Pt(11)
p.font.color.rgb = COLOR_TEXT_MUTED
p.space_before = Pt(4)

p = f_col.add_paragraph()
p.text = "Typographic Heritage"
p.font.name = FONT_SERIF
p.font.size = Pt(15)
p.font.bold = True
p.font.color.rgb = COLOR_TEXT_DARK
p.space_before = Pt(8)

p = f_col.add_paragraph()
p.text = ("• Playfair Display: Editorial headline authority\n"
          "• Cinzel: Classical Roman inscriptions & masthead\n"
          "• Fraunces: Warm humanistic serif accents\n"
          "• Space Grotesk & JetBrains Mono: Clean data readouts")
p.font.name = FONT_MONO
p.font.size = Pt(10.5)
p.font.color.rgb = COLOR_OLIVE
p.space_before = Pt(4)

p = f_col.add_paragraph()
p.text = "Back-to-Top Floating Archive Seal"
p.font.name = FONT_SERIF
p.font.size = Pt(15)
p.font.bold = True
p.font.color.rgb = COLOR_TEXT_DARK
p.space_before = Pt(8)

p = f_col.add_paragraph()
p.text = ("A floating brass-trimmed button that smoothly returns the user to Chapter 00 with one click.")
p.font.name = FONT_SANS
p.font.size = Pt(11)
p.font.color.rgb = COLOR_TEXT_MUTED

# Right Card: Responsive Quality Assurance
add_card_box(slide11, Inches(6.933), top_pos, Inches(5.6), Inches(4.9))
t_qa = slide11.shapes.add_textbox(Inches(7.183), top_pos + Inches(0.25), Inches(5.1), Inches(4.4))
f_qa = t_qa.text_frame
f_qa.word_wrap = True

p = f_qa.paragraphs[0]
p.text = "CROSS-DEVICE QUALITY ASSURANCE"
p.font.name = FONT_MONO
p.font.size = Pt(10)
p.font.bold = True
p.font.color.rgb = COLOR_GOLD

p = f_qa.add_paragraph()
p.text = "Fluid Multi-Screen Responsiveness"
p.font.name = FONT_SERIF
p.font.size = Pt(15)
p.font.bold = True
p.font.color.rgb = COLOR_TEXT_DARK
p.space_before = Pt(4)

p = f_qa.add_paragraph()
p.text = ("Tested and verified across key viewport tiers:\n"
          "• Desktop Broadside (1440px - 1920px): Full multi-column spreads and margins.\n"
          "• Laptop Standard (1024px - 1366px): Proportional fluid grid scaling.\n"
          "• Tablet Portrait (768px - 1023px): Two-column stacks and touch targets.\n"
          "• Mobile Handheld (360px - 480px): Single-column cards with collapsible drawer.")
p.font.name = FONT_SANS
p.font.size = Pt(11)
p.font.color.rgb = COLOR_TEXT_MUTED
p.space_before = Pt(4)

p = f_qa.add_paragraph()
p.text = "Performance & Accessibility Scorecard"
p.font.name = FONT_SERIF
p.font.size = Pt(15)
p.font.bold = True
p.font.color.rgb = COLOR_TEXT_DARK
p.space_before = Pt(8)

p = f_qa.add_paragraph()
p.text = ("✓ 100% Validated semantic HTML5 structure\n"
          "✓ WCAG AA compliant contrast ratios across dark brown text on cream\n"
          "✓ Zero layout shifts (CLS < 0.05) on both mobile and desktop\n"
          "✓ Zero broken hyperlinks or 404 image errors")
p.font.name = FONT_MONO
p.font.size = Pt(10.5)
p.font.color.rgb = COLOR_BURGUNDY
p.space_before = Pt(4)

# ==============================================================================
# SLIDE 12: CONCLUSION, LIVE LINKS & FUTURE ROADMAP
# ==============================================================================
slide12 = prs.slides.add_slide(blank_layout)
set_slide_background(slide12, COLOR_DARK_BG)

# Decorative double border
ob12 = slide12.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(0.5), Inches(12.333), Inches(6.5))
ob12.fill.background()
ob12.line.color.rgb = COLOR_GOLD
ob12.line.width = Pt(1.5)

ib12 = slide12.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.6), Inches(0.6), Inches(12.133), Inches(6.3))
ib12.fill.background()
ib12.line.color.rgb = COLOR_BURGUNDY
ib12.line.width = Pt(0.75)

# Heading
tb_c_top = slide12.shapes.add_textbox(Inches(1.0), Inches(0.85), Inches(11.333), Inches(0.35))
p = tb_c_top.text_frame.paragraphs[0]
p.text = "FINAL DISPATCH • PORTFOLIO MILESTONES & NEXT HORIZONS"
p.font.name = FONT_MONO
p.font.size = Pt(11)
p.font.bold = True
p.font.color.rgb = COLOR_GOLD
p.alignment = PP_ALIGN.CENTER

tb_c_title = slide12.shapes.add_textbox(Inches(1.0), Inches(1.3), Inches(11.333), Inches(1.1))
tf_c = tb_c_title.text_frame
tf_c.word_wrap = True
p = tf_c.paragraphs[0]
p.text = "A Distinguished Foundation for Future Impact"
p.font.name = FONT_SERIF
p.font.size = Pt(30)
p.font.bold = True
p.font.color.rgb = COLOR_TEXT_LIGHT
p.alignment = PP_ALIGN.CENTER

p = tf_c.add_paragraph()
p.text = "Where academic excellence, engineering craftsmanship, and visual storytelling unite."
p.font.name = FONT_SANS
p.font.size = Pt(13)
p.font.color.rgb = RGBColor(190, 180, 170)
p.alignment = PP_ALIGN.CENTER
p.space_before = Pt(4)

# 2 Bottom Cards
c_left = add_card_box(slide12, Inches(1.2), Inches(2.6), Inches(5.2), Inches(3.8), 
                      bg_color=RGBColor(38, 29, 23), border_color=COLOR_GOLD)
t_cl = slide12.shapes.add_textbox(Inches(1.45), Inches(2.8), Inches(4.7), Inches(3.4))
tf_cl = t_cl.text_frame
tf_cl.word_wrap = True
p = tf_cl.paragraphs[0]
p.text = "ENGINEERING ROADMAP (2026-2027)"
p.font.name = FONT_MONO
p.font.size = Pt(10)
p.font.bold = True
p.font.color.rgb = COLOR_GOLD

p = tf_cl.add_paragraph()
p.text = "1. Interactive Machine Learning Playgrounds"
p.font.name = FONT_SERIF
p.font.size = Pt(14)
p.font.bold = True
p.font.color.rgb = COLOR_TEXT_LIGHT
p.space_before = Pt(6)

p = tf_cl.add_paragraph()
p.text = "Embedding live Streamlit / Hugging Face model inference widgets directly into project cards."
p.font.name = FONT_SANS
p.font.size = Pt(11)
p.font.color.rgb = RGBColor(200, 190, 180)

p = tf_cl.add_paragraph()
p.text = "2. Academic Research Publications"
p.font.name = FONT_SERIF
p.font.size = Pt(14)
p.font.bold = True
p.font.color.rgb = COLOR_TEXT_LIGHT
p.space_before = Pt(6)

p = tf_cl.add_paragraph()
p.text = "Integrating undergraduate AI research papers, SIH 2026 case studies, and citations archive."
p.font.name = FONT_SANS
p.font.size = Pt(11)
p.font.color.rgb = RGBColor(200, 190, 180)

p = tf_cl.add_paragraph()
p.text = "3. Dynamic GitHub Data Pipeline"
p.font.name = FONT_SERIF
p.font.size = Pt(14)
p.font.bold = True
p.font.color.rgb = COLOR_TEXT_LIGHT
p.space_before = Pt(6)

p = tf_cl.add_paragraph()
p.text = "Automated GitHub API stats displaying real-time repository commits, stars, and code activity."
p.font.name = FONT_SANS
p.font.size = Pt(11)
p.font.color.rgb = RGBColor(200, 190, 180)

# Right Card: Access & Verification
c_right = add_card_box(slide12, Inches(6.933), Inches(2.6), Inches(5.2), Inches(3.8), 
                       bg_color=RGBColor(38, 29, 23), border_color=COLOR_BURGUNDY)
t_cr = slide12.shapes.add_textbox(Inches(7.183), Inches(2.8), Inches(4.7), Inches(3.4))
tf_cr = t_cr.text_frame
tf_cr.word_wrap = True
p = tf_cr.paragraphs[0]
p.text = "VERIFIED DISPATCH & ACCESS COORDINATES"
p.font.name = FONT_MONO
p.font.size = Pt(10)
p.font.bold = True
p.font.color.rgb = COLOR_GOLD

p = tf_cr.add_paragraph()
p.text = "Live Production Deployment:"
p.font.name = FONT_SERIF
p.font.size = Pt(13)
p.font.bold = True
p.font.color.rgb = COLOR_TEXT_LIGHT
p.space_before = Pt(6)

p = tf_cr.add_paragraph()
p.text = "https://aashi230.github.io/portfolio/"
p.font.name = FONT_MONO
p.font.size = Pt(10.5)
p.font.color.rgb = RGBColor(140, 220, 140)

p = tf_cr.add_paragraph()
p.text = "Source Code Repository:"
p.font.name = FONT_SERIF
p.font.size = Pt(13)
p.font.bold = True
p.font.color.rgb = COLOR_TEXT_LIGHT
p.space_before = Pt(6)

p = tf_cr.add_paragraph()
p.text = "https://github.com/Aashi230/portfolio (Branch: main)"
p.font.name = FONT_MONO
p.font.size = Pt(10.5)
p.font.color.rgb = RGBColor(140, 220, 140)

p = tf_cr.add_paragraph()
p.text = "Direct Inquiries & Collaboration:"
p.font.name = FONT_SERIF
p.font.size = Pt(13)
p.font.bold = True
p.font.color.rgb = COLOR_TEXT_LIGHT
p.space_before = Pt(6)

p = tf_cr.add_paragraph()
p.text = "aashimaahi23@gmail.com • SRMU, Lucknow"
p.font.name = FONT_MONO
p.font.size = Pt(10.5)
p.font.color.rgb = RGBColor(220, 190, 140)

# Save presentation
output_pptx = "Aashi_Srivastava_Portfolio_Presentation.pptx"
prs.save(output_pptx)
print(f"Presentation saved successfully to: {os.path.abspath(output_pptx)}")
