import json
import io
import os
from http.server import BaseHTTPRequestHandler
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt


# ─── HELPERS ─────────────────────────────────────────────────────────────────

def hex_to_rgb(hex_color):
    """Convert hex string to RGBColor. Handles # prefix and bad values."""
    try:
        h = str(hex_color).replace('#', '').strip()
        if len(h) == 3:
            h = h[0]*2 + h[1]*2 + h[2]*2
        if len(h) != 6:
            return RGBColor(0x1A, 0x3C, 0x6E)
        return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))
    except Exception:
        return RGBColor(0x1A, 0x3C, 0x6E)


def safe_get(d, *keys, default=''):
    """Safely get nested dict values."""
    for key in keys:
        if isinstance(d, dict):
            d = d.get(key, default)
        else:
            return default
    return d if d is not None else default


def add_text_box(slide, text, x, y, w, h, font_size=18, bold=False,
                 color='1A3C6E', font_name='Calibri', align=PP_ALIGN.LEFT,
                 wrap=True):
    """Add a styled text box to a slide."""
    txBox = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = txBox.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = str(text)
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.name = font_name
    run.font.color.rgb = hex_to_rgb(color)
    return txBox


def add_rect(slide, x, y, w, h, fill_color='1A3C6E', line_color=None):
    """Add a filled rectangle shape."""
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    shape = slide.shapes.add_shape(
        1,  # MSO_SHAPE_TYPE.RECTANGLE
        Inches(x), Inches(y), Inches(w), Inches(h)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = hex_to_rgb(fill_color)
    if line_color:
        shape.line.color.rgb = hex_to_rgb(line_color)
    else:
        shape.line.fill.background()
    return shape


# ─── SLIDE BUILDERS ──────────────────────────────────────────────────────────

def build_title_slide(slide, spec, brand):
    primary   = safe_get(brand, 'primary_colors',   default=['#1A3C6E'])[0]
    secondary = safe_get(brand, 'secondary_colors', default=['#F4A300'])[0]
    title_font = safe_get(brand, 'title_font', 'family', default='Calibri')
    body_font  = safe_get(brand, 'body_font',  'family', default='Calibri')

    # Full background
    bg = add_rect(slide, 0, 0, 10, 7.5, fill_color=primary.replace('#',''))

    # Accent bar
    add_rect(slide, 0, 4.8, 10, 0.08, fill_color=secondary.replace('#',''))

    # Title
    add_text_box(
        slide, spec.get('title', 'Presentation'),
        x=0.8, y=1.8, w=8.4, h=1.8,
        font_size=40, bold=True,
        color='FFFFFF', font_name=title_font
    )

    # Subtitle
    subtitle = spec.get('subtitle', '')
    if subtitle:
        add_text_box(
            slide, subtitle,
            x=0.8, y=3.7, w=8.4, h=0.8,
            font_size=20, bold=False,
            color='DDDDDD', font_name=body_font
        )

    # Date
    from datetime import date
    date_str = date.today().strftime('%B %Y')
    add_text_box(
        slide, date_str,
        x=6.5, y=5.1, w=3, h=0.4,
        font_size=12, color='AAAAAA',
        font_name=body_font, align=PP_ALIGN.RIGHT
    )


def build_divider_slide(slide, spec, brand):
    primary   = safe_get(brand, 'primary_colors',   default=['#1A3C6E'])[0]
    secondary = safe_get(brand, 'secondary_colors', default=['#F4A300'])[0]
    title_font = safe_get(brand, 'title_font', 'family', default='Calibri')

    # Background
    add_rect(slide, 0, 0, 10, 7.5, fill_color=primary.replace('#',''))

    # Left accent bar
    add_rect(slide, 0, 0, 0.12, 7.5, fill_color=secondary.replace('#',''))

    # Section label
    add_text_box(
        slide, 'SECTION',
        x=0.4, y=2.8, w=9, h=0.5,
        font_size=13, bold=True,
        color=secondary.replace('#',''), font_name=title_font
    )

    # Section title
    add_text_box(
        slide, spec.get('title', ''),
        x=0.4, y=3.3, w=9, h=1.6,
        font_size=36, bold=True,
        color='FFFFFF', font_name=title_font
    )


def build_content_slide(slide, spec, brand):
    primary    = safe_get(brand, 'primary_colors',    default=['#1A3C6E'])[0]
    secondary  = safe_get(brand, 'secondary_colors',  default=['#F4A300'])[0]
    title_color = spec.get('title_color', primary)
    title_font  = safe_get(brand, 'title_font', 'family', default='Calibri')
    body_font   = safe_get(brand, 'body_font',  'family', default='Calibri')
    bullets     = spec.get('bullets', [])
    visual_type = spec.get('visual_type', 'text').lower()

    # Top accent bar
    add_rect(slide, 0, 0, 10, 0.08, fill_color=primary.replace('#',''))

    # Slide title
    add_text_box(
        slide, spec.get('title', ''),
        x=0.5, y=0.15, w=9, h=0.65,
        font_size=24, bold=True,
        color=title_color.replace('#',''), font_name=title_font
    )

    # Divider line under title
    add_rect(slide, 0.5, 0.88, 9, 0.03, fill_color='E5E7EB')

    # Route to visual type
    if visual_type in ('three-column', 'three_column', 'icons'):
        build_three_column(slide, bullets, secondary, body_font)
    elif visual_type in ('two-column', 'two_column'):
        build_two_column(slide, bullets, primary, secondary, body_font)
    elif visual_type == 'table':
        build_table_layout(slide, bullets, primary, secondary, body_font)
    elif visual_type == 'quote':
        build_quote_layout(slide, bullets, primary, secondary, title_font, body_font)
    else:
        build_bullet_list(slide, bullets, secondary, body_font)

    # Slide number
    add_text_box(
        slide, str(spec.get('slide_number', '')),
        x=9.2, y=7.1, w=0.6, h=0.3,
        font_size=10, color='AAAAAA',
        font_name=body_font, align=PP_ALIGN.RIGHT
    )

    # Speaker note
    note = spec.get('speaker_note', '')
    if note:
        notes_slide = slide.notes_slide
        notes_slide.notes_text_frame.text = note


def build_bullet_list(slide, bullets, accent_color, body_font):
    if not bullets:
        return
    y_start = 1.1
    row_h   = min(0.7, 5.8 / max(len(bullets), 1))

    for i, bullet in enumerate(bullets):
        y = y_start + i * row_h

        # Bullet dot
        add_rect(slide, 0.5, y + 0.18, 0.12, 0.12,
                 fill_color=accent_color.replace('#',''))

        # Bullet text
        add_text_box(
            slide, str(bullet),
            x=0.75, y=y, w=8.8, h=row_h,
            font_size=16, color='1A1A1A', font_name=body_font
        )


def build_three_column(slide, bullets, accent_color, body_font):
    bullets = bullets or ['Point 1', 'Point 2', 'Point 3']
    col_w   = 2.8
    gap     = 0.25
    start_x = 0.5

    for i in range(min(3, len(bullets))):
        x = start_x + i * (col_w + gap)

        # Card background
        card = slide.shapes.add_shape(1, Inches(x), Inches(1.1), Inches(col_w), Inches(5.9))
        card.fill.solid()
        card.fill.fore_color.rgb = hex_to_rgb('F9FAFB')
        card.line.color.rgb = hex_to_rgb('E5E7EB')

        # Accent top bar
        add_rect(slide, x, 1.1, col_w, 0.08, fill_color=accent_color.replace('#',''))

        # Number
        add_text_box(
            slide, str(i + 1).zfill(2),
            x=x+0.15, y=1.25, w=0.6, h=0.5,
            font_size=22, bold=True,
            color=accent_color.replace('#',''), font_name=body_font
        )

        # Content
        add_text_box(
            slide, str(bullets[i]),
            x=x+0.15, y=1.85, w=col_w-0.3, h=4.9,
            font_size=14, color='1A1A1A', font_name=body_font
        )


def build_two_column(slide, bullets, primary, secondary, body_font):
    mid   = len(bullets) // 2 + len(bullets) % 2
    left  = bullets[:mid]
    right = bullets[mid:]

    def add_col(items, x):
        for idx, b in enumerate(items):
            y = 1.1 + idx * 0.9
            add_rect(slide, x, y+0.1, 0.06, 0.5, fill_color=primary.replace('#',''))
            add_text_box(slide, str(b), x=x+0.2, y=y, w=4.0, h=0.85,
                         font_size=15, color='1A1A1A', font_name=body_font)

    add_col(left, 0.5)
    add_col(right, 5.2)

    # Vertical divider
    add_rect(slide, 4.95, 1.0, 0.02, 6.0, fill_color='E5E7EB')


def build_table_layout(slide, bullets, primary, secondary, body_font):
    if not bullets:
        return

    from pptx.util import Inches, Pt
    rows   = len(bullets) + 1
    cols   = 2
    table  = slide.shapes.add_table(rows, cols, Inches(0.5), Inches(1.1),
                                     Inches(9), Inches(5.8)).table

    # Header row
    for c, hdr in enumerate(['Item', 'Details']):
        cell = table.cell(0, c)
        cell.text = hdr
        cell.fill.solid()
        cell.fill.fore_color.rgb = hex_to_rgb(primary.replace('#',''))
        p = cell.text_frame.paragraphs[0]
        run = p.add_run()
        run.font.bold  = True
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        run.font.size  = Pt(14)
        run.font.name  = body_font

    # Data rows
    for i, bullet in enumerate(bullets):
        parts = str(bullet).split(':', 1)
        left  = parts[0].strip()
        right = parts[1].strip() if len(parts) > 1 else ''

        for c, val in enumerate([left, right]):
            cell = table.cell(i + 1, c)
            cell.text = val
            if i % 2 == 0:
                cell.fill.solid()
                cell.fill.fore_color.rgb = hex_to_rgb('F9FAFB')
            p = cell.text_frame.paragraphs[0]
            run = p.add_run()
            run.font.size = Pt(13)
            run.font.name = body_font
            run.font.color.rgb = RGBColor(0x1A, 0x1A, 0x1A)


def build_quote_layout(slide, bullets, primary, secondary, title_font, body_font):
    quote = bullets[0] if bullets else 'Key insight goes here.'

    # Large quote mark
    add_text_box(slide, '\u201c', x=0.5, y=0.9, w=1.2, h=1.5,
                 font_size=80, bold=True,
                 color=secondary.replace('#',''), font_name=title_font)

    # Quote text
    add_text_box(slide, quote, x=1.2, y=1.5, w=8.3, h=3.5,
                 font_size=24, bold=False,
                 color='1A1A1A', font_name=title_font, align=PP_ALIGN.LEFT)

    # Accent line
    add_rect(slide, 1.2, 5.1, 4, 0.06, fill_color=secondary.replace('#',''))

    # Attribution
    if len(bullets) > 1:
        add_text_box(slide, bullets[1], x=1.2, y=5.3, w=8, h=0.5,
                     font_size=14, color='555555', font_name=body_font)


# ─── MAIN GENERATOR ──────────────────────────────────────────────────────────

def generate_pptx(final_spec, brand_rulebook):
    prs = Presentation()
    prs.slide_width  = Inches(10)
    prs.slide_height = Inches(7.5)

    blank_layout = prs.slide_layouts[6]  # completely blank layout

    for spec in final_spec:
        slide      = prs.slides.add_slide(blank_layout)
        slide_type = spec.get('type', 'content').lower()

        # Set slide background color
        bg_color = spec.get('background_color', '')
        if not bg_color:
            bg_list = brand_rulebook.get('background_colors', ['#FFFFFF'])
            bg_color = bg_list[0] if bg_list else '#FFFFFF'

        bg = slide.background
        fill = bg.fill
        fill.solid()
        fill.fore_color.rgb = hex_to_rgb(bg_color.replace('#', ''))

        if slide_type == 'title':
            build_title_slide(slide, spec, brand_rulebook)
        elif slide_type == 'divider':
            build_divider_slide(slide, spec, brand_rulebook)
        else:
            build_content_slide(slide, spec, brand_rulebook)

    # Save to bytes buffer
    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf


# ─── VERCEL HANDLER ──────────────────────────────────────────────────────────

class handler(BaseHTTPRequestHandler):

    def do_OPTIONS(self):
        self.send_response(200)
        self._send_cors_headers()
        self.end_headers()

    def do_POST(self):
        try:
            content_length = int(self.headers.get('Content-Length', 0))
            body           = self.rfile.read(content_length)
            data           = json.loads(body)

            final_spec     = data.get('finalSpec',     [])
            brand_rulebook = data.get('brandRulebook', {})

            if not final_spec:
                self._send_json(400, {'error': 'finalSpec is required'})
                return

            pptx_buf = generate_pptx(final_spec, brand_rulebook)
            pptx_bytes = pptx_buf.read()

            self.send_response(200)
            self._send_cors_headers()
            self.send_header('Content-Type',
                             'application/vnd.openxmlformats-officedocument.presentationml.presentation')
            self.send_header('Content-Disposition',
                             'attachment; filename="presentation.pptx"')
            self.send_header('Content-Length', str(len(pptx_bytes)))
            self.end_headers()
            self.wfile.write(pptx_bytes)

        except Exception as e:
            self._send_json(500, {'error': str(e)})

    def do_GET(self):
        self._send_json(405, {'error': 'Method not allowed'})

    def _send_cors_headers(self):
        self.send_header('Access-Control-Allow-Origin',  '*')
        self.send_header('Access-Control-Allow-Methods', 'POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')

    def _send_json(self, status, obj):
        body = json.dumps(obj).encode()
        self.send_response(status)
        self._send_cors_headers()
        self.send_header('Content-Type',   'application/json')
        self.send_header('Content-Length', str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def log_message(self, format, *args):
        pass  # silence default logging