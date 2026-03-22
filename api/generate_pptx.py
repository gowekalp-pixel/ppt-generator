import json
import io
import base64
from http.server import BaseHTTPRequestHandler
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN


def hex_to_rgb(hex_color):
    try:
        h = str(hex_color).replace('#', '').strip()
        if len(h) == 3:
            h = h[0]*2 + h[1]*2 + h[2]*2
        if len(h) != 6:
            return RGBColor(0x1A, 0x3C, 0x6E)
        return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))
    except Exception:
        return RGBColor(0x1A, 0x3C, 0x6E)


def safe_list(brand, key, default):
    val = brand.get(key, [])
    return val if isinstance(val, list) and val else [default]


def safe_font(brand, key):
    f = brand.get(key, {})
    return f.get('family', 'Calibri') if isinstance(f, dict) else 'Calibri'


def clean_hex(val):
    return str(val).replace('#', '').strip() or '1A3C6E'


def add_text(slide, text, x, y, w, h, size=16, bold=False,
             color='1A1A1A', font='Calibri', align=PP_ALIGN.LEFT):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = str(text)
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.name = font
    run.font.color.rgb = hex_to_rgb(color)


def add_rect(slide, x, y, w, h, color='1A3C6E'):
    s = slide.shapes.add_shape(1, Inches(x), Inches(y), Inches(w), Inches(h))
    s.fill.solid()
    s.fill.fore_color.rgb = hex_to_rgb(color)
    s.line.fill.background()


def build_title_slide(slide, spec, brand):
    p  = clean_hex(safe_list(brand, 'primary_colors',   '#1A3C6E')[0])
    s  = clean_hex(safe_list(brand, 'secondary_colors', '#F4A300')[0])
    tf = safe_font(brand, 'title_font')
    bf = safe_font(brand, 'body_font')

    add_rect(slide, 0, 0, 10, 7.5, p)
    add_rect(slide, 0, 4.6, 10, 0.1, s)
    add_text(slide, spec.get('title', 'Presentation'),
             0.8, 1.6, 8.4, 2.0, size=40, bold=True, color='FFFFFF', font=tf)
    sub = spec.get('subtitle', '')
    if sub:
        add_text(slide, sub, 0.8, 3.7, 8.4, 0.8, size=20, color='DDDDDD', font=bf)
    from datetime import date
    add_text(slide, date.today().strftime('%B %Y'),
             6.5, 5.1, 3.0, 0.4, size=11, color='AAAAAA', font=bf, align=PP_ALIGN.RIGHT)


def build_divider_slide(slide, spec, brand):
    p  = clean_hex(safe_list(brand, 'primary_colors',   '#1A3C6E')[0])
    s  = clean_hex(safe_list(brand, 'secondary_colors', '#F4A300')[0])
    tf = safe_font(brand, 'title_font')

    add_rect(slide, 0, 0, 10, 7.5, p)
    add_rect(slide, 0, 0, 0.15, 7.5, s)
    add_text(slide, 'SECTION', 0.5, 2.8, 9, 0.5, size=12, bold=True, color=s, font=tf)
    add_text(slide, spec.get('title', ''), 0.5, 3.4, 9, 1.6,
             size=34, bold=True, color='FFFFFF', font=tf)


def build_bullets(slide, bullets, accent, font):
    if not bullets:
        return
    row_h = min(0.75, 5.6 / max(len(bullets), 1))
    for i, b in enumerate(bullets):
        y = 1.1 + i * row_h
        add_rect(slide, 0.5, y + 0.2, 0.12, 0.12, accent)
        add_text(slide, str(b), 0.75, y, 8.8, row_h, size=15, color='1A1A1A', font=font)


def build_three_col(slide, bullets, accent, font):
    bullets = list(bullets or [])
    while len(bullets) < 3:
        bullets.append('—')
    for i in range(3):
        x = 0.5 + i * 3.05
        card = slide.shapes.add_shape(1, Inches(x), Inches(1.1), Inches(2.8), Inches(6.0))
        card.fill.solid()
        card.fill.fore_color.rgb = hex_to_rgb('F9FAFB')
        card.line.color.rgb = hex_to_rgb('E5E7EB')
        add_rect(slide, x, 1.1, 2.8, 0.1, accent)
        add_text(slide, str(i+1).zfill(2), x+0.15, 1.28, 0.6, 0.5,
                 size=22, bold=True, color=accent, font=font)
        add_text(slide, str(bullets[i]), x+0.15, 1.9, 2.5, 4.9,
                 size=13, color='1A1A1A', font=font)


def build_two_col(slide, bullets, primary, font):
    mid   = len(bullets) // 2 + len(bullets) % 2
    left  = bullets[:mid]
    right = bullets[mid:]
    add_rect(slide, 5.0, 1.0, 0.03, 6.2, 'E5E7EB')

    def col(items, x):
        for idx, b in enumerate(items):
            y = 1.1 + idx * 0.95
            add_rect(slide, x, y+0.12, 0.07, 0.55, primary)
            add_text(slide, str(b), x+0.22, y, 4.0, 0.9, size=14, color='1A1A1A', font=font)

    col(left, 0.5)
    col(right, 5.3)


def build_quote(slide, bullets, secondary, title_font, body_font):
    quote = str(bullets[0]) if bullets else 'Key insight.'
    add_text(slide, '\u201c', 0.5, 0.9, 1.2, 1.5, size=80, bold=True,
             color=secondary, font=title_font)
    add_text(slide, quote, 1.3, 1.6, 8.2, 3.2, size=22, color='1A1A1A', font=title_font)
    add_rect(slide, 1.3, 5.0, 3.5, 0.07, secondary)
    if len(bullets) > 1:
        add_text(slide, str(bullets[1]), 1.3, 5.2, 8.0, 0.5,
                 size=12, color='555555', font=body_font)


def build_content_slide(slide, spec, brand):
    p   = clean_hex(safe_list(brand, 'primary_colors',   '#1A3C6E')[0])
    s   = clean_hex(safe_list(brand, 'secondary_colors', '#F4A300')[0])
    tf  = safe_font(brand, 'title_font')
    bf  = safe_font(brand, 'body_font')
    bul = spec.get('bullets', [])
    vt  = str(spec.get('visual_type', 'text')).lower().replace('-', '_').replace(' ', '_')

    add_rect(slide, 0, 0, 10, 0.1, p)
    add_text(slide, spec.get('title', ''), 0.5, 0.18, 9.0, 0.7,
             size=24, bold=True, color=p, font=tf)
    add_rect(slide, 0.5, 0.93, 9.0, 0.03, 'E5E7EB')

    if vt in ('three_column', 'icons'):
        build_three_col(slide, bul, s, bf)
    elif vt == 'two_column':
        build_two_col(slide, bul, p, bf)
    elif vt == 'quote':
        build_quote(slide, bul, s, tf, bf)
    else:
        build_bullets(slide, bul, s, bf)

    add_text(slide, str(spec.get('slide_number', '')),
             9.2, 7.1, 0.6, 0.3, size=10, color='AAAAAA', font=bf, align=PP_ALIGN.RIGHT)

    note = spec.get('speaker_note', '')
    if note:
        try:
            slide.notes_slide.notes_text_frame.text = str(note)
        except Exception:
            pass


def generate_pptx(final_spec, brand_rulebook):
    prs = Presentation()
    prs.slide_width  = Inches(10)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    for spec in final_spec:
        slide = prs.slides.add_slide(blank)

        bg = spec.get('background_color') or \
             (safe_list(brand_rulebook, 'background_colors', '#FFFFFF')[0])
        fill = slide.background.fill
        fill.solid()
        fill.fore_color.rgb = hex_to_rgb(bg)

        stype = str(spec.get('type', 'content')).lower()
        if stype == 'title':
            build_title_slide(slide, spec, brand_rulebook)
        elif stype == 'divider':
            build_divider_slide(slide, spec, brand_rulebook)
        else:
            build_content_slide(slide, spec, brand_rulebook)

    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf


class handler(BaseHTTPRequestHandler):

    def do_OPTIONS(self):
        self.send_response(200)
        self._cors()
        self.end_headers()

    def do_GET(self):
        self._json(405, {'error': 'Use POST'})

    def do_POST(self):
        try:
            length = int(self.headers.get('Content-Length', 0))
            body   = self.rfile.read(length)
            data   = json.loads(body)

            final_spec     = data.get('finalSpec', [])
            brand_rulebook = data.get('brandRulebook', {})

            if not final_spec:
                self._json(400, {'error': 'finalSpec is required'})
                return

            buf     = generate_pptx(final_spec, brand_rulebook)
            encoded = base64.b64encode(buf.read()).decode('utf-8')

            self._json(200, {
                'success':  True,
                'filename': 'presentation.pptx',
                'data':     encoded,
                'slides':   len(final_spec)
            })

        except Exception as e:
            import traceback
            self._json(500, {'error': str(e), 'details': traceback.format_exc()})

    def _cors(self):
        self.send_header('Access-Control-Allow-Origin',  '*')
        self.send_header('Access-Control-Allow-Methods', 'POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')

    def _json(self, status, obj):
        body = json.dumps(obj).encode()
        self.send_response(status)
        self._cors()
        self.send_header('Content-Type',   'application/json')
        self.send_header('Content-Length', str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def log_message(self, format, *args):
        pass
