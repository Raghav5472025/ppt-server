from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
import json, io, requests
from PIL import Image

app = Flask(__name__)
CORS(app)

def hex_to_rgb(hex_color):
    hex_color = hex_color.lstrip('#')
    return RGBColor(
        int(hex_color[0:2], 16),
        int(hex_color[2:4], 16),
        int(hex_color[4:6], 16)
    )

def add_rect(slide, x, y, w, h, color, transparency=0):
    shape = slide.shapes.add_shape(1, Inches(x), Inches(y), Inches(w), Inches(h))
    shape.fill.solid()
    shape.fill.fore_color.rgb = hex_to_rgb(color)
    shape.line.fill.background()
    return shape

def add_text(slide, text, x, y, w, h, font_size=14, bold=False, color='FFFFFF', align='left', italic=False):
    txBox = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = str(text)
    if align == 'center': p.alignment = PP_ALIGN.CENTER
    elif align == 'right': p.alignment = PP_ALIGN.RIGHT
    else: p.alignment = PP_ALIGN.LEFT
    run = p.runs[0]
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = hex_to_rgb(color)
    return txBox

def add_oval(slide, x, y, w, h, color):
    shape = slide.shapes.add_shape(9, Inches(x), Inches(y), Inches(w), Inches(h))
    shape.fill.solid()
    shape.fill.fore_color.rgb = hex_to_rgb(color)
    shape.line.fill.background()
    return shape

# ── SLIDE RENDERERS ──────────────────────────────────────────────────────────

def render_title(slide, data, T):
    W, H = 10, 5.625
    # Full background
    add_rect(slide, 0, 0, W, H, T['bg'])
    # Decorative circles
    add_oval(slide, 6.5, -1.5, 5, 5, T['accent'])
    add_oval(slide, 7.5, -0.5, 3, 3, T['accent2'])
    add_oval(slide, -0.8, 3.5, 2.5, 2.5, T['accent'])
    # Left accent bar
    add_rect(slide, 0, 0, 0.12, H, T['accent'])
    # Tag pill
    add_rect(slide, 0.5, 0.45, 3.0, 0.38, T['accent'])
    add_text(slide, 'AI GENERATED PRESENTATION', 0.5, 0.45, 3.0, 0.38, 8, True, 'FFFFFF', 'center')
    # Main title
    title = data.get('heading', 'Presentation')
    size = 28 if len(title) > 40 else 36 if len(title) > 25 else 46
    add_text(slide, title, 0.5, 1.0, 8.5, 2.2, size, True, 'FFFFFF', 'left')
    # Subheading
    if data.get('subheading'):
        add_text(slide, data['subheading'], 0.5, 3.3, 7.5, 0.65, 16, False, T['accent'], 'left', True)
    # Bullets as pills
    bullets = data.get('bullets', [])[:3]
    for i, b in enumerate(bullets):
        add_rect(slide, 0.5, 4.0 + i * 0.35, 7.0, 0.28, T['accent'])
        add_text(slide, '• ' + b, 0.65, 4.0 + i * 0.35, 6.8, 0.28, 10, False, 'FFFFFF', 'left')
    # Footer
    add_rect(slide, 0, H - 0.48, W, 0.48, T['dark'])
    add_text(slide, 'Powered by HackMate AI', 0.5, H - 0.48, W - 1, 0.48, 9, False, T['accent'], 'center')

def render_bullets(slide, data, T, n):
    W, H = 10, 5.625
    add_rect(slide, 0, 0, W, H, 'F8FAFC')
    # Header
    add_rect(slide, 0, 0, W, 1.0, T['bg'])
    add_rect(slide, 0, 0, 0.12, 1.0, T['accent'])
    add_text(slide, data.get('heading', ''), 0.28, 0, 9.2, 1.0, 20, True, 'FFFFFF', 'left')
    add_text(slide, str(n), 9.3, 0.32, 0.5, 0.36, 10, False, '94A3B8', 'right')
    if data.get('subheading'):
        add_rect(slide, 0.35, 1.06, 9.3, 0.36, T['accent'])
        add_text(slide, data['subheading'], 0.5, 1.06, 9.0, 0.36, 11, False, 'FFFFFF', 'left')
    bullets = data.get('bullets', [])[:6]
    y_start = 1.55 if data.get('subheading') else 1.12
    total_h = H - y_start - 0.5
    bh = min(total_h / max(len(bullets), 1) - 0.06, 0.9)
    for i, b in enumerate(bullets):
        y = y_start + i * (bh + 0.06)
        add_rect(slide, 0.35, y, 9.3, bh, 'FFFFFF')
        add_rect(slide, 0.35, y, 0.1, bh, T['accent'])
        add_oval(slide, 0.58, y + bh/2 - 0.15, 0.3, 0.3, T['accent'])
        add_text(slide, str(i+1), 0.58, y + bh/2 - 0.15, 0.3, 0.3, 9, True, 'FFFFFF', 'center')
        add_text(slide, b, 1.05, y, 8.5, bh, 12, False, '1E293B', 'left')

def render_big_stat(slide, data, T, n):
    W, H = 10, 5.625
    add_rect(slide, 0, 0, W, H, T['bg'])
    add_rect(slide, 0, 0, W, 1.0, T['dark'])
    add_rect(slide, 0, 0, 0.12, 1.0, T['accent'])
    add_text(slide, data.get('heading', ''), 0.28, 0, 9.2, 1.0, 20, True, 'FFFFFF', 'left')
    add_text(slide, str(n), 9.3, 0.32, 0.5, 0.36, 10, False, '94A3B8', 'right')
    stats = data.get('stats', [])[:3]
    sw = 2.88 if len(stats) == 3 else 4.2
    start_x = (W - sw * len(stats) - 0.18 * (len(stats)-1)) / 2
    for i, stat in enumerate(stats):
        x = start_x + i * (sw + 0.18)
        add_rect(slide, x, 1.1, sw, 1.95, T['accent'])
        add_rect(slide, x, 1.1, sw, 0.08, T['accent'])
        add_text(slide, stat.get('number',''), x, 1.18, sw, 1.0, 38, True, 'FFFFFF', 'center')
        add_text(slide, stat.get('label',''), x+0.1, 2.22, sw-0.2, 0.44, 11, True, 'FFFFFF', 'center')
        if stat.get('context'):
            add_text(slide, stat['context'], x+0.08, 2.68, sw-0.16, 0.32, 9, False, 'CBD5E1', 'center')
    bullets = data.get('bullets', [])[:3]
    if bullets:
        add_rect(slide, 0.35, 3.15, W-0.7, 0.04, T['accent'])
        for i, b in enumerate(bullets):
            add_oval(slide, 0.38, 3.25 + i*0.56, 0.28, 0.28, T['accent'])
            add_text(slide, b, 0.78, 3.22 + i*0.56, 8.9, 0.5, 12, False, 'E2E8F0', 'left')

def render_three_cards(slide, data, T, n):
    W, H = 10, 5.625
    add_rect(slide, 0, 0, W, H, 'F8FAFC')
    add_rect(slide, 0, 0, W, 1.0, T['bg'])
    add_rect(slide, 0, 0, 0.12, 1.0, T['accent'])
    add_text(slide, data.get('heading', ''), 0.28, 0, 9.2, 1.0, 20, True, 'FFFFFF', 'left')
    add_text(slide, str(n), 9.3, 0.32, 0.5, 0.36, 10, False, '94A3B8', 'right')
    cards = data.get('cards', [])[:3]
    cw = (W - 0.7) / 3
    for i, card in enumerate(cards):
        x = 0.35 + i * cw
        add_rect(slide, x, 1.08, cw-0.08, H-1.55, 'FFFFFF')
        add_rect(slide, x, 1.08, cw-0.08, 0.06, T['accent'])
        # Emoji circle
        add_oval(slide, x + (cw-0.08)/2 - 0.38, 1.22, 0.76, 0.76, T['accent'])
        add_text(slide, card.get('emoji','💡'), x + (cw-0.08)/2 - 0.38, 1.22, 0.76, 0.76, 20, False, 'FFFFFF', 'center')
        add_text(slide, card.get('title',''), x+0.08, 2.1, cw-0.24, 0.52, 12, True, '1E293B', 'center')
        add_rect(slide, x + (cw-0.08)*0.28, 2.66, (cw-0.08)*0.44, 0.04, T['accent'])
        points = card.get('points', [])[:3]
        for j, pt in enumerate(points):
            add_text(slide, '• ' + pt, x+0.12, 2.78 + j*0.68, cw-0.32, 0.62, 10, False, '374151', 'left')

def render_timeline(slide, data, T, n):
    W, H = 10, 5.625
    add_rect(slide, 0, 0, W, H, T['bg'])
    add_rect(slide, 0, 0, W, 1.0, T['dark'])
    add_rect(slide, 0, 0, 0.12, 1.0, T['accent'])
    add_text(slide, data.get('heading', ''), 0.28, 0, 9.2, 1.0, 20, True, 'FFFFFF', 'left')
    add_text(slide, str(n), 9.3, 0.32, 0.5, 0.36, 10, False, '94A3B8', 'right')
    steps = data.get('steps', [])[:4]
    sw = (W - 0.6) / len(steps) if steps else 2.5
    for i, step in enumerate(steps):
        x = 0.3 + i * sw
        is_last = i == len(steps) - 1
        bg_color = T['accent'] if is_last else 'FFFFFF'
        txt_color = 'FFFFFF' if is_last else '1E293B'
        add_rect(slide, x, 1.1, sw-0.1, H-1.55, bg_color)
        add_oval(slide, x + (sw-0.1)/2 - 0.38, 1.22, 0.76, 0.76, T['dark'] if is_last else T['bg'])
        add_text(slide, step.get('number', str(i+1)), x + (sw-0.1)/2 - 0.38, 1.22, 0.76, 0.76, 18, True, T['accent'] if not is_last else 'FFFFFF', 'center')
        add_text(slide, step.get('title',''), x+0.1, 2.1, sw-0.3, 0.55, 11, True, txt_color, 'center')
        add_rect(slide, x + (sw-0.1)*0.25, 2.7, (sw-0.1)*0.5, 0.04, T['accent'])
        add_text(slide, step.get('description',''), x+0.1, 2.82, sw-0.3, H-3.4, 10, False, txt_color, 'center')
        if not is_last:
            add_text(slide, '▶', x+sw-0.18, 1.5, 0.2, 0.4, 12, False, T['accent'], 'center')

def render_two_column(slide, data, T, n):
    W, H = 10, 5.625
    add_rect(slide, 0, 0, W, H, 'F8FAFC')
    add_rect(slide, 0, 0, W, 1.0, T['bg'])
    add_rect(slide, 0, 0, 0.12, 1.0, T['accent'])
    add_text(slide, data.get('heading', ''), 0.28, 0, 9.2, 1.0, 20, True, 'FFFFFF', 'left')
    add_text(slide, str(n), 9.3, 0.32, 0.5, 0.36, 10, False, '94A3B8', 'right')
    cw = 4.45
    # Left
    add_rect(slide, 0.3, 1.08, cw, H-1.55, 'FFFFFF')
    add_rect(slide, 0.3, 1.08, cw, 0.48, T['bg'])
    add_text(slide, data.get('left_heading','Left'), 0.45, 1.08, cw-0.25, 0.48, 13, True, 'FFFFFF', 'left')
    for i, b in enumerate(data.get('left_bullets',[])[:5]):
        add_rect(slide, 0.42, 1.68 + i*0.7, 0.08, 0.55, T['accent'])
        add_text(slide, b, 0.6, 1.65 + i*0.7, cw-0.42, 0.62, 11, False, '1E293B', 'left')
    # Right
    add_rect(slide, 5.25, 1.08, cw, H-1.55, T['bg'])
    add_rect(slide, 5.25, 1.08, cw, 0.48, T['accent'])
    add_text(slide, data.get('right_heading','Right'), 5.4, 1.08, cw-0.25, 0.48, 13, True, 'FFFFFF', 'left')
    for i, b in enumerate(data.get('right_bullets',[])[:5]):
        add_rect(slide, 5.35, 1.68 + i*0.7, 0.08, 0.55, T['accent'])
        add_text(slide, b, 5.52, 1.65 + i*0.7, cw-0.42, 0.62, 11, False, 'E2E8F0', 'left')

def render_comparison(slide, data, T, n):
    W, H = 10, 5.625
    add_rect(slide, 0, 0, W, H, 'F8FAFC')
    add_rect(slide, 0, 0, W, 1.0, T['bg'])
    add_rect(slide, 0, 0, 0.12, 1.0, T['accent'])
    add_text(slide, data.get('heading', ''), 0.28, 0, 9.2, 1.0, 20, True, 'FFFFFF', 'left')
    add_text(slide, str(n), 9.3, 0.32, 0.5, 0.36, 10, False, '94A3B8', 'right')
    cw = 4.45
    has_verdict = bool(data.get('verdict'))
    col_h = H - 1.98 if has_verdict else H - 1.55
    add_rect(slide, 0.3, 1.08, cw, col_h, 'FFFFFF')
    add_rect(slide, 0.3, 1.08, cw, 0.5, T['bg'])
    add_text(slide, data.get('left_heading','Option A'), 0.45, 1.08, cw-0.3, 0.5, 13, True, 'FFFFFF', 'left')
    for i, b in enumerate(data.get('left_bullets',[])[:5]):
        add_oval(slide, 0.42, 1.7 + i*0.65, 0.24, 0.24, T['accent'])
        add_text(slide, b, 0.75, 1.66 + i*0.65, cw-0.55, 0.6, 11, False, '1E293B', 'left')
    add_rect(slide, 5.25, 1.08, cw, col_h, T['bg'])
    add_rect(slide, 5.25, 1.08, cw, 0.5, T['accent'])
    add_text(slide, data.get('right_heading','Option B'), 5.4, 1.08, cw-0.3, 0.5, 13, True, 'FFFFFF', 'left')
    for i, b in enumerate(data.get('right_bullets',[])[:5]):
        add_oval(slide, 5.35, 1.7 + i*0.65, 0.24, 0.24, T['accent'])
        add_text(slide, b, 5.68, 1.66 + i*0.65, cw-0.55, 0.6, 11, False, 'E2E8F0', 'left')
    if has_verdict:
        add_rect(slide, 0.3, H-0.78, W-0.6, 0.44, T['accent'])
        add_text(slide, '💡  ' + data['verdict'], 0.45, H-0.78, W-0.9, 0.44, 11, False, 'FFFFFF', 'left')

def render_quote(slide, data, T, n):
    W, H = 10, 5.625
    add_rect(slide, 0, 0, W, H, T['bg'])
    add_oval(slide, -2, -2, 6, 6, T['accent'])
    add_oval(slide, 7, 2, 5, 5, T['accent'])
    add_rect(slide, 0, 0, 0.12, H, T['accent'])
    add_rect(slide, 0, 0, W, 0.52, T['dark'])
    add_text(slide, data.get('heading',''), 0.28, 0, W-0.6, 0.52, 13, True, T['accent'], 'left')
    add_text(slide, str(n), 9.3, 0.15, 0.5, 0.36, 10, False, '94A3B8', 'right')
    add_text(slide, '\u201C', 0.3, 0.55, 1.2, 1.1, 72, True, T['accent'], 'left')
    quote = data.get('quote','')
    size = 16 if len(quote) > 100 else 20
    add_text(slide, quote, 0.8, 1.1, 8.4, 2.0, size, True, 'FFFFFF', 'center')
    add_rect(slide, 3.5, 3.25, 3.0, 0.06, T['accent'])
    if data.get('author'):
        add_text(slide, '— ' + data['author'], 0.5, 3.38, W-1, 0.38, 12, False, T['accent'], 'center')
    if data.get('explanation'):
        add_rect(slide, 0.5, 3.88, W-1, 0.55, T['accent'])
        add_text(slide, data['explanation'], 0.65, 3.88, W-1.3, 0.55, 11, False, 'FFFFFF', 'center')

def render_checklist(slide, data, T, n):
    W, H = 10, 5.625
    add_rect(slide, 0, 0, W, H, 'F8FAFC')
    add_rect(slide, 0, 0, W, 1.0, T['bg'])
    add_rect(slide, 0, 0, 0.12, 1.0, T['accent'])
    add_text(slide, data.get('heading',''), 0.28, 0, 9.2, 1.0, 20, True, 'FFFFFF', 'left')
    add_text(slide, str(n), 9.3, 0.32, 0.5, 0.36, 10, False, '94A3B8', 'right')
    cw = 4.45
    add_rect(slide, 0.3, 1.08, cw, H-1.55, 'F0FDF4')
    add_rect(slide, 0.3, 1.08, cw, 0.48, '16A34A')
    add_text(slide, data.get('left_heading','✅ Do This'), 0.45, 1.08, cw-0.3, 0.48, 12, True, 'FFFFFF', 'left')
    for i, item in enumerate(data.get('left_items',[])[:5]):
        add_text(slide, '✓', 0.42, 1.68 + i*0.65, 0.3, 0.55, 14, True, '16A34A', 'center')
        add_text(slide, item, 0.76, 1.65 + i*0.65, cw-0.6, 0.58, 11, False, '14532D', 'left')
    add_rect(slide, 5.25, 1.08, cw, H-1.55, 'FFF1F2')
    add_rect(slide, 5.25, 1.08, cw, 0.48, 'DC2626')
    add_text(slide, data.get('right_heading','❌ Avoid'), 5.4, 1.08, cw-0.3, 0.48, 12, True, 'FFFFFF', 'left')
    for i, item in enumerate(data.get('right_items',[])[:5]):
        add_text(slide, '✗', 5.37, 1.68 + i*0.65, 0.3, 0.55, 14, True, 'DC2626', 'center')
        add_text(slide, item, 5.7, 1.65 + i*0.65, cw-0.6, 0.58, 11, False, '7F1D1D', 'left')

def render_case_study(slide, data, T, n):
    W, H = 10, 5.625
    add_rect(slide, 0, 0, W, H, 'F8FAFC')
    add_rect(slide, 0, 0, W, 1.0, T['bg'])
    add_rect(slide, 0, 0, 0.12, 1.0, T['accent'])
    add_text(slide, data.get('heading','Case Study'), 0.28, 0, 9.2, 1.0, 20, True, 'FFFFFF', 'left')
    add_text(slide, str(n), 9.3, 0.32, 0.5, 0.36, 10, False, '94A3B8', 'right')
    add_rect(slide, 0.3, 1.08, W-0.6, 0.44, T['accent'])
    add_text(slide, '📌  ' + data.get('case_name',''), 0.45, 1.08, W-0.9, 0.44, 13, True, 'FFFFFF', 'left')
    sections = [
        ('🔍 SITUATION', data.get('situation',''), 'EFF6FF', '93C5FD', '1E3A8A'),
        ('⚡ ACTION', data.get('action',''), 'F0FDF4', '86EFAC', '14532D'),
        ('📈 RESULT', data.get('result',''), 'FEF9C3', 'FDE047', '713F12'),
        ('💡 LESSON', data.get('lesson',''), 'FAE8FF', 'E879F9', '701A75'),
    ]
    sw = (W - 0.7) / 2
    for i, (label, content, bg, border, text_c) in enumerate(sections):
        x = 0.3 + (i % 2) * (sw + 0.1)
        y = 1.62 + (i // 2) * 1.82
        add_rect(slide, x, y, sw, 1.7, bg)
        add_text(slide, label, x+0.1, y+0.06, sw-0.2, 0.32, 9, True, text_c, 'left')
        add_rect(slide, x+0.1, y+0.4, sw-0.2, 0.03, border)
        add_text(slide, content, x+0.1, y+0.48, sw-0.2, 1.15, 10, False, text_c, 'left')

def render_closing(slide, data, T):
    W, H = 10, 5.625
    add_rect(slide, 0, 0, W, H, T['bg'])
    add_oval(slide, -2, -2, 6.5, 6.5, T['accent'])
    add_oval(slide, 7.5, 2.0, 6, 6, T['accent'])
    add_rect(slide, 0, 0, 0.12, H, T['accent'])
    add_text(slide, data.get('heading','Thank You!'), 0.5, 0.7, 9, 1.6, 56, True, 'FFFFFF', 'center')
    add_rect(slide, 3.0, 2.5, 4.0, 0.07, T['accent'])
    if data.get('subheading'):
        add_text(slide, data['subheading'], 0.5, 2.65, 9, 0.6, 16, False, T['accent'], 'center')
    takeaways = data.get('key_takeaways', [])[:3]
    if takeaways:
        add_text(slide, 'KEY TAKEAWAYS', 0.8, 3.38, 8.4, 0.28, 9, True, '94A3B8', 'center')
        for i, pt in enumerate(takeaways):
            x = 0.8 + i * 3.1
            add_rect(slide, x, 3.74, 2.85, 1.35, T['dark'])
            add_rect(slide, x, 3.74, 2.85, 1.35, T['accent'])
            add_text(slide, pt, x+0.1, 3.82, 2.65, 1.18, 10, False, 'E5E7EB', 'center')
    add_rect(slide, 0, H-0.48, W, 0.48, T['dark'])
    add_text(slide, 'Powered by HackMate AI', 0.5, H-0.48, W-1, 0.48, 9, False, T['accent'], 'center')

# ── MAIN ROUTE ───────────────────────────────────────────────────────────────

SLIDE_RENDERERS = {
    'title': render_title,
    'bullets': render_bullets,
    'big_stat': render_big_stat,
    'three_cards': render_three_cards,
    'timeline': render_timeline,
    'two_column': render_two_column,
    'comparison': render_comparison,
    'quote_focus': render_quote,
    'checklist': render_checklist,
    'case_study': render_case_study,
    'closing': render_closing,
}

@app.route('/generate', methods=['POST'])
def generate_ppt():
    try:
        body = request.json
        slides_data = body.get('slides', [])
        template = body.get('template', {})

        T = {
            'bg': template.get('bg', '09061E'),
            'accent': template.get('accent', '8B5CF6'),
            'accent2': template.get('accent2', '7C3AED'),
            'dark': template.get('dark', '040210'),
        }

        prs = Presentation()
        prs.slide_width = Inches(10)
        prs.slide_height = Inches(5.625)

        blank_layout = prs.slide_layouts[6]

        for idx, slide_data in enumerate(slides_data):
            slide = prs.slides.add_slide(blank_layout)
            layout = slide_data.get('layout', 'bullets')
            renderer = SLIDE_RENDERERS.get(layout, render_bullets)
            if layout in ['title', 'closing']:
                renderer(slide, slide_data, T)
            else:
                renderer(slide, slide_data, T, idx + 1)

        output = io.BytesIO()
        prs.save(output)
        output.seek(0)

        filename = body.get('title', 'Presentation').replace(' ', '_')[:40] + '.pptx'

        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation',
            as_attachment=True,
            download_name=filename
        )

    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/health', methods=['GET'])
def health():
    return jsonify({'status': 'ok', 'service': 'HackMate PPT Server'})

if __name__ == '__main__':
    app.run(debug=False, host='0.0.0.0', port=10000)
