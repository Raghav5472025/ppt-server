from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
import io
import os
import json
import requests

app = Flask(__name__)
CORS(app)

# ── Themes ────────────────────────────────────────────────────
THEMES = {
    'purple': {'primary': '7c3aed', 'bg': 'faf5ff', 'bg2': 'f3f0ff', 'text': '4c1d95', 'body': '374151'},
    'blue':   {'primary': '1d4ed8', 'bg': 'eff6ff', 'bg2': 'dbeafe', 'text': '1e3a8a', 'body': '374151'},
    'dark':   {'primary': '6d28d9', 'bg': '1f2937', 'bg2': '111827', 'text': 'f9fafb', 'body': 'd1d5db'},
    'green':  {'primary': '059669', 'bg': 'f0fdf4', 'bg2': 'dcfce7', 'text': '064e3b', 'body': '374151'},
    'orange': {'primary': 'ea580c', 'bg': 'fff7ed', 'bg2': 'fed7aa', 'text': '7c2d12', 'body': '374151'},
}

def hex_to_rgb(h):
    h = h.lstrip('#')
    return RGBColor(int(h[0:2],16), int(h[2:4],16), int(h[4:6],16))

def add_textbox(slide, text, x, y, w, h, size=18, bold=False, color='000000', align=PP_ALIGN.LEFT, italic=False):
    txBox = slide.shapes.add_textbox(Emu(x), Emu(y), Emu(w), Emu(h))
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = hex_to_rgb(color)
    return txBox

def add_rect(slide, x, y, w, h, fill_color, border_color=None):
    shape = slide.shapes.add_shape(1, Emu(x), Emu(y), Emu(w), Emu(h))
    shape.fill.solid()
    shape.fill.fore_color.rgb = hex_to_rgb(fill_color)
    if border_color:
        shape.line.color.rgb = hex_to_rgb(border_color)
    else:
        shape.line.fill.background()
    return shape

def build_slide_title(slide, data, T):
    # Background
    bg = slide.shapes.add_shape(1, Emu(0), Emu(0), Emu(9144000), Emu(6858000))
    bg.fill.solid()
    bg.fill.fore_color.rgb = hex_to_rgb(T['primary'])
    bg.line.fill.background()

    # Emoji
    if data.get('emoji'):
        add_textbox(slide, data['emoji'], 3800000, 700000, 1400000, 800000, size=48, align=PP_ALIGN.CENTER)

    # Title
    add_textbox(slide, data.get('title',''), 500000, 1700000, 8144000, 1200000,
                size=44, bold=True, color='FFFFFF', align=PP_ALIGN.CENTER)

    # Subtitle
    if data.get('subtitle'):
        add_textbox(slide, data['subtitle'], 500000, 3100000, 8144000, 700000,
                    size=20, color='E0D7FF', align=PP_ALIGN.CENTER, italic=True)

    # Footer
    add_textbox(slide, 'HackMate · Built to Win', 500000, 5800000, 8144000, 400000,
                size=11, color='C4B5FD', align=PP_ALIGN.CENTER)

def build_slide_impact(slide, data, T):
    # Top bar
    add_rect(slide, 0, 0, 9144000, 180000, T['primary'])
    # Title
    add_textbox(slide, data.get('title',''), 457200, 280000, 8229600, 700000,
                size=26, bold=True, color=T['text'])
    # Emoji
    if data.get('emoji'):
        add_textbox(slide, data['emoji'], 120000, 220000, 600000, 600000, size=28, align=PP_ALIGN.CENTER)

    stats = data.get('stats', [])
    cw = 2500000
    sx = (9144000 - cw * 3 - 300000 * 2) // 2

    for i, s in enumerate(stats[:3]):
        x = sx + i * (cw + 300000)
        add_rect(slide, x, 1400000, cw, 2000000, T['primary'])
        add_textbox(slide, s.get('number',''), x+100000, 1700000, cw-200000, 900000,
                    size=38, bold=True, color='FFFFFF', align=PP_ALIGN.CENTER)
        add_textbox(slide, s.get('label',''), x+100000, 2700000, cw-200000, 500000,
                    size=14, color='E0D7FF', align=PP_ALIGN.CENTER)

def build_slide_how(slide, data, T, has_diagram=False):
    add_rect(slide, 0, 0, 9144000, 180000, T['primary'])
    add_textbox(slide, data.get('title',''), 457200, 280000, 8229600, 700000,
                size=26, bold=True, color=T['text'])
    if data.get('emoji'):
        add_textbox(slide, data['emoji'], 120000, 220000, 600000, 600000, size=28, align=PP_ALIGN.CENTER)

    steps = data.get('steps') or data.get('content') or []

    if has_diagram:
        # Steps on left, diagram area on right
        for i, s in enumerate(steps[:4]):
            y = 1300000 + i * 900000
            add_rect(slide, 457200, y, 600000, 600000, T['primary'])
            add_textbox(slide, str(i+1), 457200, y, 600000, 600000,
                        size=22, bold=True, color='FFFFFF', align=PP_ALIGN.CENTER)
            add_textbox(slide, s, 1200000, y+80000, 3300000, 500000,
                        size=14, color=T['body'])
        # Diagram area placeholder box
        diag_box = slide.shapes.add_shape(1, Emu(4900000), Emu(1200000), Emu(3900000), Emu(4500000))
        diag_box.fill.solid()
        diag_box.fill.fore_color.rgb = hex_to_rgb(T['bg2'] if T['bg2'] else 'f3f0ff')
        diag_box.line.color.rgb = hex_to_rgb(T['primary'])
        add_textbox(slide, '[ Diagram ]', 4900000, 3200000, 3900000, 500000,
                    size=14, color=T['primary'], align=PP_ALIGN.CENTER)
    else:
        bw, bh = 1850000, 1800000
        for i, s in enumerate(steps[:4]):
            x = 400000 + i * (bw + 200000)
            add_rect(slide, x, 1500000, bw, bh, T['primary'] + '20' if len(T['primary'])==6 else T['primary'])
            add_textbox(slide, str(i+1), x, 1600000, bw, 500000,
                        size=30, bold=True, color=T['primary'], align=PP_ALIGN.CENTER)
            add_textbox(slide, s, x+80000, 2250000, bw-160000, 900000,
                        size=14, color=T['body'], align=PP_ALIGN.CENTER)
            if i < 3:
                add_textbox(slide, '→', x+bw+30000, 2150000, 140000, 400000,
                            size=22, bold=True, color=T['primary'], align=PP_ALIGN.CENTER)

def build_slide_team(slide, data, T):
    add_rect(slide, 0, 0, 9144000, 180000, T['primary'])
    add_textbox(slide, data.get('title',''), 457200, 280000, 8229600, 700000,
                size=26, bold=True, color=T['text'])
    if data.get('emoji'):
        add_textbox(slide, data['emoji'], 120000, 220000, 600000, 600000, size=28, align=PP_ALIGN.CENTER)

    members = data.get('members', [])
    cw = 2000000
    sx = (9144000 - cw * len(members) - 300000 * (len(members)-1)) // 2

    for i, m in enumerate(members[:4]):
        x = sx + i * (cw + 300000)
        # Avatar circle
        circle = slide.shapes.add_shape(9, Emu(x+600000), Emu(1600000), Emu(800000), Emu(800000))
        circle.fill.solid()
        circle.fill.fore_color.rgb = hex_to_rgb(T['primary'])
        circle.line.fill.background()
        init = ''.join([n[0] for n in (m.get('name') or 'TM').split()])[:2].upper()
        add_textbox(slide, init, x+600000, 1700000, 800000, 600000,
                    size=24, bold=True, color='FFFFFF', align=PP_ALIGN.CENTER)
        add_textbox(slide, m.get('name',''), x, 2600000, cw, 450000,
                    size=15, bold=True, color=T['text'], align=PP_ALIGN.CENTER)
        add_textbox(slide, m.get('role',''), x, 3100000, cw, 400000,
                    size=12, color=T['body'], align=PP_ALIGN.CENTER)

def build_slide_default(slide, data, T, has_diagram=False):
    add_rect(slide, 0, 0, 9144000, 180000, T['primary'])
    add_textbox(slide, data.get('title',''), 457200, 280000, 8229600, 700000,
                size=26, bold=True, color=T['text'])
    if data.get('emoji'):
        add_textbox(slide, data['emoji'], 120000, 220000, 600000, 600000, size=28, align=PP_ALIGN.CENTER)

    y = 1200000
    content_w = 4300000 if has_diagram else 8229600

    if data.get('headline'):
        add_textbox(slide, data['headline'], 457200, 1100000, content_w, 450000,
                    size=16, bold=True, color=T['primary'])
        add_rect(slide, 457200, 1620000, 900000, 35000, T['primary'])
        y = 1800000

    if data.get('content'):
        body = '\n'.join(['• ' + c for c in data['content']])
        add_textbox(slide, body, 457200, y, content_w, 6858000-y-200000,
                    size=18, color=T['body'])

    if has_diagram:
        diag_box = slide.shapes.add_shape(1, Emu(5000000), Emu(1200000), Emu(3700000), Emu(4500000))
        diag_box.fill.solid()
        diag_box.fill.fore_color.rgb = hex_to_rgb(T['bg2'] if T['bg2'] else 'f3f0ff')
        diag_box.line.color.rgb = hex_to_rgb(T['primary'])
        add_textbox(slide, '[ Diagram ]', 5000000, 3200000, 3700000, 500000,
                    size=14, color=T['primary'], align=PP_ALIGN.CENTER)

# ── AI Slide Generation ───────────────────────────────────────
def generate_slides_with_ai(topic, profile, diagram_mode):
    api_key = os.environ.get('ANTHROPIC_API_KEY')
    if not api_key:
        return get_fallback_slides(topic, profile)

    diagram_instructions = {
        'none': 'Do NOT add svg_diagram to any slide.',
        'simple': 'Add "has_diagram": true to ONE most relevant slide only (how it works or core concept).',
        'full': 'Add "has_diagram": true to slides where diagrams help (how it works, architecture, process).',
    }.get(diagram_mode, '')

    prompt = f"""Create a hackathon pitch deck for: "{topic}"
Creator: {profile.get('full_name','Student')} | {profile.get('role','Developer')} | {profile.get('college','')}

{diagram_instructions}

Return ONLY valid JSON:
{{"slides":[
  {{"layout":"title","title":"<name>","subtitle":"<value prop>","emoji":"🚀"}},
  {{"layout":"problem","title":"The Problem","headline":"<stat>","content":["<point1>","<point2>","<point3>"],"emoji":"😤"}},
  {{"layout":"solution","title":"Our Solution","headline":"<solution>","content":["<f1>","<f2>","<f3>"],"emoji":"💡"}},
  {{"layout":"how","title":"How It Works","steps":["<s1>","<s2>","<s3>","<s4>"],"emoji":"⚙️","has_diagram":true}},
  {{"layout":"tech","title":"Tech Stack","content":["Frontend: <tech>","Backend: <tech>","AI: <tech>","DB: <tech>"],"emoji":"🛠️"}},
  {{"layout":"impact","title":"Impact","stats":[{{"number":"<n>","label":"<l>"}},{{"number":"<n>","label":"<l>"}},{{"number":"<n>","label":"<l>"}}],"emoji":"📈"}},
  {{"layout":"demo","title":"Key Features","content":["<f1>","<f2>","<f3>","<f4>"],"emoji":"✨"}},
  {{"layout":"team","title":"Team","members":[{{"name":"{profile.get('full_name','Lead')}","role":"{profile.get('role','Developer')}"}}],"emoji":"👥"}}
]}}"""

    try:
        r = requests.post('https://api.anthropic.com/v1/messages',
            headers={'x-api-key': api_key, 'anthropic-version': '2023-06-01', 'Content-Type': 'application/json'},
            json={'model': 'claude-haiku-4-5-20251001', 'max_tokens': 2000,
                  'messages': [{'role': 'user', 'content': prompt}]},
            timeout=30)
        data = r.json()
        text = data.get('content', [{}])[0].get('text', '')
        clean = text.replace('```json','').replace('```','').strip()
        parsed = json.loads(clean)
        return parsed.get('slides', get_fallback_slides(topic, profile))
    except Exception as e:
        print('AI error:', e)
        return get_fallback_slides(topic, profile)

def get_fallback_slides(topic, profile):
    return [
        {'layout':'title','title':topic or 'My Project','subtitle':'An innovative hackathon solution','emoji':'🚀'},
        {'layout':'problem','title':'The Problem','headline':'A critical challenge','content':['Current solutions are inefficient','Users face real pain points','No affordable solution exists'],'emoji':'😤'},
        {'layout':'solution','title':'Our Solution','headline':f'Introducing {topic}','content':['AI-powered approach','Simple design','Scalable from day one'],'emoji':'💡'},
        {'layout':'how','title':'How It Works','steps':['User inputs data','AI analyzes','Results generated','User acts'],'emoji':'⚙️','has_diagram':True},
        {'layout':'tech','title':'Tech Stack','content':['Frontend: React','Backend: FastAPI','AI: Claude','Database: PostgreSQL'],'emoji':'🛠️'},
        {'layout':'impact','title':'Impact','stats':[{'number':'10M+','label':'Users'},{'number':'₹500Cr','label':'Market'},{'number':'80%','label':'Efficiency'}],'emoji':'📈'},
        {'layout':'demo','title':'Key Features','content':['Real-time processing','Intuitive UI','Offline support','Privacy first'],'emoji':'✨'},
        {'layout':'team','title':'Team','members':[{'name':profile.get('full_name','Lead'),'role':profile.get('role','Developer')}],'emoji':'👥'},
    ]

# ── Main Route ────────────────────────────────────────────────
@app.route('/generate-ppt', methods=['POST', 'OPTIONS'])
def generate_ppt():
    if request.method == 'OPTIONS':
        return '', 200

    try:
        body = request.get_json()
        topic = body.get('topic', 'My Project')
        theme_name = body.get('theme', 'purple')
        profile = body.get('profile', {})
        provided_slides = body.get('slides', [])
        generate_only = body.get('generateOnly', False)

        # Diagram mode detection
        t = topic.lower()
        if any(w in t for w in ['no diagram','no image','no chart','text only','plain']):
            diagram_mode = 'none'
        elif any(w in t for w in ['diagram','flowchart','chart','graph','visual']):
            diagram_mode = 'full'
        else:
            diagram_mode = 'simple'

        # If only JSON needed (preview)
        if generate_only:
            slides = provided_slides or generate_slides_with_ai(topic, profile, diagram_mode)
            return jsonify({'slides': slides})

        # Generate full PPTX
        slides = provided_slides or generate_slides_with_ai(topic, profile, diagram_mode)
        T = THEMES.get(theme_name, THEMES['purple'])

        prs = Presentation()
        prs.slide_width = Emu(9144000)
        prs.slide_height = Emu(6858000)

        blank_layout = prs.slide_layouts[6]  # Completely blank

        for i, slide_data in enumerate(slides):
            slide = prs.slides.add_slide(blank_layout)
            layout = slide_data.get('layout', 'default')
            has_diagram = slide_data.get('has_diagram', False)

            if layout == 'title':
                build_slide_title(slide, slide_data, T)
            elif layout == 'impact':
                build_slide_impact(slide, slide_data, T)
            elif layout == 'how':
                build_slide_how(slide, slide_data, T, has_diagram)
            elif layout == 'team':
                build_slide_team(slide, slide_data, T)
            else:
                build_slide_default(slide, slide_data, T, has_diagram)

        # Save to buffer
        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)

        filename = (topic[:30] + '.pptx').replace(' ', '_')
        return send_file(buf, as_attachment=True,
                        download_name=filename,
                        mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation')

    except Exception as e:
        print('Error:', e)
        return jsonify({'error': str(e)}), 500

@app.route('/health', methods=['GET'])
def health():
    return jsonify({'status': 'ok', 'service': 'HackMate PPT Generator'})

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
