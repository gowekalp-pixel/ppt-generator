"""
/api/generate-pptx.py
Vercel Python serverless function.
Receives Agent 5 finalSpec JSON → builds PPTX using python-pptx → returns base64.

Input (POST body JSON):
  {
    "finalSpec": [...],       # Agent 5 designed spec array
    "brandRulebook": {...}    # Agent 2 brand rulebook (for fallback values)
  }

Output (JSON):
  {
    "success": true,
    "data": "<base64 pptx>",
    "slides": 12,
    "filename": "presentation_20241115.pptx"
  }
"""

import json
import base64
import io
import traceback
from datetime import datetime
from http.server import BaseHTTPRequestHandler

# python-pptx imports
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE
from pptx.enum.chart import XL_CHART_TYPE
from pptx.chart.data import ChartData
from pptx.util import Inches, Pt
import pptx.oxml.ns as nsmap
from lxml import etree


# Renderer-side spacing / fit rules. These must not override Agent 5 placement,
# only how content is laid out inside the frames Agent 5 already chose.
INTERNAL_PADDING = 0.14
MIN_TEXT_MARGIN = 0.08
HEADER_TO_ARTIFACT = 0.12
ARTIFACT_TO_ARTIFACT = 0.18
ZONE_TOP_OFFSET = 0.08
HEADER_HEIGHT = 0.30
INSIGHT_LEFT_PADDING = 0.15
INSIGHT_RIGHT_PADDING = 0.12
INSIGHT_TOP_PADDING = 0.10
INSIGHT_BOTTOM_PADDING = 0.10
CARD_GAP = 0.15
CARD_INNER_PADDING = 0.15
TITLE_TO_SUBTITLE = 0.08
SUBTITLE_TO_BODY = 0.10
TABLE_MIN_ROW_HEIGHT = 0.32
CELL_PADDING = 0.09
CHART_HEADER_GAP = 0.11
OPTICAL_NUDGE = 0.02


# ─── TEMPLATE HELPERS ─────────────────────────────────────────────────────────

def clear_slides(prs):
    """Remove all content slides, preserving slide masters and layouts."""
    sldIdLst = prs.slides._sldIdLst
    for sldId in list(sldIdLst):
        rId = sldId.get(nsmap.qn('r:id'))
        prs.part.drop_rel(rId)
        sldIdLst.remove(sldId)


def find_blank_layout(prs):
    """Return the most minimal slide layout — prefers 'Blank', falls back by placeholder count."""
    for layout in prs.slide_layouts:
        if (layout.name or '').lower().strip() == 'blank':
            return layout
    for layout in prs.slide_layouts:
        if 'blank' in (layout.name or '').lower():
            return layout
    # Last resort: layout with fewest placeholders
    try:
        return min(prs.slide_layouts, key=lambda l: len(list(l.placeholders)))
    except Exception:
        return prs.slide_layouts[0]


def load_template_prs(template_b64):
    """
    Decode a base64 PPTX template, strip all content slides, and return the
    Presentation object. Slide masters and layouts are preserved so the brand
    master (background, fonts, logo, decorative shapes) carries through to
    every new slide.
    """
    template_bytes = base64.b64decode(template_b64)
    prs = Presentation(io.BytesIO(template_bytes))
    clear_slides(prs)
    return prs


# ─── HELPERS ──────────────────────────────────────────────────────────────────

def hex_to_rgb(hex_color):
    """Convert #RRGGBB or #RGB to RGBColor."""
    if not hex_color or not isinstance(hex_color, str):
        return RGBColor(0x11, 0x11, 0x11)
    h = hex_color.lstrip('#')
    if len(h) == 3:
        h = h[0]*2 + h[1]*2 + h[2]*2
    if len(h) != 6:
        return RGBColor(0x11, 0x11, 0x11)
    try:
        return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))
    except Exception:
        return RGBColor(0x11, 0x11, 0x11)

def inches(val):
    """Safe inches conversion — handles None and strings."""
    try:
        return Inches(float(val or 0))
    except Exception:
        return Inches(0)

def pt(val):
    """Safe Pt conversion."""
    try:
        return Pt(float(val or 12))
    except Exception:
        return Pt(12)

def set_font(run_or_para, font_family, font_size, bold=False, italic=False, color_hex=None):
    """Apply font properties to a run or paragraph."""
    try:
        font = run_or_para.font
        if font_family:
            font.name = str(font_family)
        if font_size:
            font.size = pt(font_size)
        font.bold   = bool(bold)
        font.italic = bool(italic)
        if color_hex:
            font.color.rgb = hex_to_rgb(color_hex)
    except Exception:
        pass


def enable_text_fit(text_frame):
    """Keep text inside the shape whenever possible."""
    try:
        text_frame.word_wrap = True
        text_frame.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
        text_frame.margin_left = 0
        text_frame.margin_right = 0
        text_frame.margin_top = 0
        text_frame.margin_bottom = 0
    except Exception:
        pass


def estimate_fit_font_size(text, width_in, height_in, max_size=12, min_size=7):
    """Cheap heuristic for table cells and dense text."""
    text = str(text or '')
    if not text.strip():
        return max_size

    width_pts = max(1, width_in * 72)
    height_pts = max(1, height_in * 72)
    usable_chars = max(6, int(width_pts / 5.4))
    lines = 0
    for chunk in text.split('\n'):
        lines += max(1, int(len(chunk) / usable_chars) + 1)

    if lines <= 1:
        return max_size

    size = min(max_size, int(height_pts / max(lines * 1.45, 1)))
    return max(min_size, size)


def estimate_wrapped_lines(text, width_in, font_size):
    """Cheap line-wrap estimate for short headings and labels."""
    text = str(text or '').strip()
    if not text:
        return 1
    width_pts = max(1, width_in * 72)
    char_w = max(1.0, font_size * 0.52)
    chars_per_line = max(4, int(width_pts / char_w))
    lines = 0
    for chunk in text.split('\n'):
        chunk = chunk.strip()
        lines += max(1, int(len(chunk) / chars_per_line) + 1)
    return max(1, lines)


def estimate_header_block_height(text, width_in, font_size):
    """Estimate rendered height for an artifact header, allowing wrapping."""
    lines = estimate_wrapped_lines(text, width_in, font_size)
    line_h = (font_size / 72.0) * 1.18
    return max(HEADER_HEIGHT, lines * line_h + 0.02)


def infer_slide_header_style(slide_spec):
    """Choose one consistent header language per content slide."""
    # Keep artifact headers visually consistent across all content slides.
    # Agent 5 controls placement; Agent 6 should not switch header language
    # slide-by-slide based on artifact sentiment.
    return 'underline'


def add_image_box(slide, image_b64, x, y, w, h):
    """Render an image from base64."""
    if not image_b64:
        return None
    try:
        blob = base64.b64decode(image_b64)
        return slide.shapes.add_picture(io.BytesIO(blob), inches(x), inches(y), inches(w), inches(h))
    except Exception:
        return None

def add_text_box(slide, x, y, w, h, text, font_family='Arial', font_size=12,
                 bold=False, color_hex='#111111', align='left', valign='top',
                 wrap=True, italic=False, hanging_in=0.0):
    """Add a text box to a slide with full formatting.

    hanging_in — if > 0, applies a hanging indent so that wrapped lines align
                 with the character after the bullet marker.
    """
    txBox = slide.shapes.add_textbox(inches(x), inches(y), inches(w), inches(h))
    tf    = txBox.text_frame
    tf.word_wrap = wrap
    enable_text_fit(tf)

    # Vertical alignment
    valign_map = { 'middle': MSO_ANCHOR.MIDDLE, 'bottom': MSO_ANCHOR.BOTTOM, 'top': MSO_ANCHOR.TOP }
    tf.vertical_anchor = valign_map.get(valign, MSO_ANCHOR.TOP)

    # Clear default paragraph and set text
    p   = tf.paragraphs[0]
    run = p.add_run()
    run.text = str(text or '')

    # Paragraph alignment
    align_map = { 'left': PP_ALIGN.LEFT, 'center': PP_ALIGN.CENTER, 'right': PP_ALIGN.RIGHT }
    p.alignment = align_map.get(align, PP_ALIGN.LEFT)

    set_font(run, font_family, font_size, bold, italic, color_hex)

    # Hanging indent: all continuation lines align with text after bullet
    if hanging_in > 0:
        _mar = int(Emu(inches(hanging_in)))
        pPr  = p._p.get_or_add_pPr()
        pPr.set('marL',   str(_mar))
        pPr.set('indent', str(-_mar))

    return txBox

def add_filled_rect(slide, x, y, w, h, fill_hex=None, border_hex=None,
                    border_pt=0, corner_radius=0):
    """Add a filled rectangle shape (used for backgrounds, cards, etc.)"""
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    shape = slide.shapes.add_shape(
        1,  # MSO_SHAPE.RECTANGLE
        inches(x), inches(y), inches(w), inches(h)
    )

    # Fill
    if fill_hex:
        shape.fill.solid()
        shape.fill.fore_color.rgb = hex_to_rgb(fill_hex)
    else:
        shape.fill.background()

    # Border
    if border_hex and border_pt and border_pt > 0:
        shape.line.color.rgb = hex_to_rgb(border_hex)
        shape.line.width     = pt(border_pt)
    else:
        shape.line.fill.background()

    # Corner radius — set via XML
    if corner_radius and corner_radius > 0:
        try:
            sp  = shape._element
            prstGeom = sp.find('.//' + '{http://schemas.openxmlformats.org/drawingml/2006/main}prstGeom')
            if prstGeom is not None:
                prstGeom.set('prst', 'roundRect')
                avLst = prstGeom.find('{http://schemas.openxmlformats.org/drawingml/2006/main}avLst')
                if avLst is None:
                    avLst = etree.SubElement(prstGeom, '{http://schemas.openxmlformats.org/drawingml/2006/main}avLst')
                # corner radius as percentage of shape size (max 50000 = 50%)
                radius_pct = min(int(corner_radius * 5000), 50000)
                gd = etree.SubElement(avLst, '{http://schemas.openxmlformats.org/drawingml/2006/main}gd')
                gd.set('name', 'adj')
                gd.set('fmla', 'val ' + str(radius_pct))
        except Exception:
            pass

    return shape

def rgb_tuple(hex_color):
    """Return (r, g, b) tuple from hex."""
    h = (hex_color or '#000000').lstrip('#')
    if len(h) == 3: h = h[0]*2 + h[1]*2 + h[2]*2
    try:
        return (int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))
    except Exception:
        return (0, 0, 0)


def is_dark_color(hex_color):
    """Return True if perceived luminance is < 0.5 (i.e. color is dark)."""
    r, g, b = rgb_tuple(hex_color)
    return (0.299 * r + 0.587 * g + 0.114 * b) / 255 < 0.5


def find_layout_by_name(prs, name):
    """Find a slide layout by exact name, then partial match."""
    if not name:
        return None
    name_lower = name.lower().strip()
    for layout in prs.slide_layouts:
        if (layout.name or '').lower().strip() == name_lower:
            return layout
    for layout in prs.slide_layouts:
        if name_lower in (layout.name or '').lower():
            return layout
    return None


def place_in_placeholder(slide, ph_idx, text, style_spec, bt):
    """
    Write text into the placeholder at ph_idx on the slide.
    Falls back to a free-form text box if the placeholder is not found.
    """
    if not text:
        return
    style_spec = style_spec or {}
    bt         = bt or {}

    try:
        for ph in slide.placeholders:
            if ph.placeholder_format.idx == ph_idx:
                tf = ph.text_frame
                tf.clear()
                p   = tf.paragraphs[0]
                run = p.add_run()
                run.text = str(text)
                font_family = style_spec.get('font_family') or bt.get('title_font_family', 'Arial')
                font_size   = style_spec.get('font_size', 18 if ph_idx == 0 else 14)
                bold        = style_spec.get('font_weight', '') in ('bold', 'semibold')
                color_hex   = style_spec.get('color') or bt.get('title_color', '#111111')
                align       = style_spec.get('align', 'left')
                set_font(run, font_family, font_size, bold, False, color_hex)
                align_map = {'left': PP_ALIGN.LEFT, 'center': PP_ALIGN.CENTER, 'right': PP_ALIGN.RIGHT}
                p.alignment = align_map.get(align, PP_ALIGN.LEFT)
                return
    except Exception as e:
        print(f'place_in_placeholder({ph_idx}) lookup error:', e)

    # Fallback: free-form text box using default positions
    if ph_idx == 0:
        add_text_box(slide, 0.4, 0.15, 9.0, 0.8, text,
                     style_spec.get('font_family', bt.get('title_font_family', 'Arial')),
                     style_spec.get('font_size', 20),
                     style_spec.get('font_weight', 'bold') in ('bold', 'semibold'),
                     style_spec.get('color', bt.get('title_color', '#1A3C8F')),
                     style_spec.get('align', 'left'), 'middle')
    else:
        add_text_box(slide, 0.4, 1.0, 9.0, 0.5, text,
                     style_spec.get('font_family', bt.get('body_font_family', 'Arial')),
                     style_spec.get('font_size', 14),
                     False,
                     style_spec.get('color', bt.get('body_color', '#333333')),
                     style_spec.get('align', 'left'), 'middle')


def render_header_block(slide, header_block, bt, header_style='underline'):
    """
    Render an artifact header label.
    style='underline' — text + thin brand-colour rule below.
    style='brand_fill' — filled rectangle with contrasting text.
    """
    if not header_block:
        return
    x     = header_block.get('x', 0)
    y     = header_block.get('y', 0)
    w     = header_block.get('w', 4)
    h     = header_block.get('h', 0.3)
    text  = header_block.get('text', '')
    style = header_style or 'underline'

    if not text:
        return

    primary   = bt.get('primary_color', '#1A3C8F')
    font_fam  = bt.get('title_font_family', 'Arial')
    font_size = header_block.get('font_size', 11)
    text_h = estimate_header_block_height(text, w, font_size)
    rule_gap = 0.02
    rule_h = 0.03

    if style == 'brand_fill':
        fill_color = header_block.get('accent_color') or primary
        add_filled_rect(slide, x, y, w, h, fill_hex=fill_color)
        text_color = '#FFFFFF' if is_dark_color(fill_color) else '#111111'
        add_text_box(slide, x + 0.08, y, w - 0.16, h,
                     text, font_fam, font_size, True, text_color, 'left', 'middle')
        return y + h
    else:
        # Underline style
        add_text_box(slide, x, y, w, text_h,
                     text, font_fam, font_size, True, primary, 'left', 'top')
        rule_y = y + text_h + rule_gap
        add_filled_rect(slide, x, rule_y, w, rule_h, fill_hex=primary)
        return rule_y + rule_h


# ─── SLIDE RENDERERS ──────────────────────────────────────────────────────────

def render_insight_text(slide, artifact, bt, suppress_heading=False):
    """Render an insight_text artifact.

    suppress_heading — when True, the heading was already written to a layout header
                       placeholder; skip the inline heading rendering.
    """
    x = artifact.get('x', 0)
    y = artifact.get('y', 0)
    w = artifact.get('w', 4)
    h = artifact.get('h', 2)

    style   = artifact.get('style', {})
    hs      = artifact.get('heading_style', {})
    bs      = artifact.get('body_style', {})
    heading = artifact.get('heading', '') if not suppress_heading else ''
    points  = artifact.get('points', [])

    # Background fill + border (skip in layout mode — template provides styling)
    if not suppress_heading:
        fill   = style.get('fill_color')
        border = style.get('border_color')
        bw     = style.get('border_width', 0)
        cr     = style.get('corner_radius', 0)
        if fill or (border and bw):
            add_filled_rect(slide, x, y, w, h, fill_hex=fill,
                            border_hex=border, border_pt=bw, corner_radius=cr)

    # ── Heading & body layout ────────────────────────────────────────────────
    TOP_PAD   = INSIGHT_TOP_PADDING
    content_y = y + TOP_PAD

    # Collect heading params; rendering is deferred until after b_size is known
    _h_params = None
    if heading:
        _h_params = {
            'font':  hs.get('font_family', bt.get('title_font_family', 'Arial')),
            'size':  hs.get('font_size', 12),
            'color': hs.get('color', bt.get('primary_color', '#1A3C8F')),
            'bold':  hs.get('font_weight', 'bold') in ('bold', 'semibold'),
        }
        content_y = y + 0.40   # heading row + gap

    # ── Native OOXML bullet helpers (defined early — used in sizing estimate too) ──
    BULLET_HANG = 0.22   # inches — bullet hangs left; text wraps at this indent

    def _strip_marker(s):
        s = str(s)
        return s[1:].lstrip() if s[:1] in ('✓', '✗', '→', '•') else s

    # Body points — compute sizes first so heading can scale to match
    if points:
        b_font       = bs.get('font_family', bt.get('body_font_family', 'Arial'))
        b_size       = bs.get('font_size', 10)
        b_color      = bs.get('color', '#111111')
        list_style   = bs.get('list_style', 'bullet')
        indent_in    = float(bs.get('indent_inches', INSIGHT_LEFT_PADDING))
        vert_dist    = bs.get('vertical_distribution', '')
        space_bef_pt = max(1, int(bs.get('space_before_pt', 4)))

        body_h = max(0.1, h - (content_y - y) - INSIGHT_BOTTOM_PADDING)
        text_w = max(0.5, w - indent_in - INSIGHT_RIGHT_PADDING - BULLET_HANG)

        # Estimate wrapped line count per point at a given font size
        def _wrapped_lines(text, font_pt, avail_w_in):
            # Average character width ≈ 50% of em (conservative)
            char_w = (font_pt / 72.0) * 0.50
            chars_per_line = max(1, int(avail_w_in / char_w))
            stripped = _strip_marker(str(text))
            words = stripped.split()
            if not words:
                return 1
            lines, line_len = 1, 0
            for word in words:
                wlen = len(word)
                if line_len == 0:
                    line_len = wlen
                elif line_len + 1 + wlen <= chars_per_line:
                    line_len += 1 + wlen
                else:
                    lines += 1
                    line_len = wlen
            return lines

        # Estimate total rendered height using actual wrap count — ONLY scale DOWN, never up.
        # If text already fits at the spec font_size, keep it as-is.
        def _total_h(font_pt, gap_pt):
            lh = (font_pt / 72.0) * 1.30   # line height with leading
            gh = gap_pt / 72.0
            return sum(_wrapped_lines(p, font_pt, text_w) * lh + gh for p in points)

        total_h_est = _total_h(b_size, space_bef_pt)
        if total_h_est > body_h * 0.95:
            # Binary search for the largest font_pt that still fits
            lo, hi = 7, b_size
            while hi - lo > 0.5:
                mid = (lo + hi) / 2.0
                if _total_h(mid, space_bef_pt) <= body_h * 0.95:
                    lo = mid
                else:
                    hi = mid
            b_size = max(7, int(lo))
        # else: text fits → keep b_size unchanged (no upscaling)

        # Inter-bullet gap: use spec value; only increase if there is clear excess space
        total_lines = sum(_wrapped_lines(p, b_size, text_w) for p in points)
        if total_lines > 12:
            b_size = max(7, b_size - 1)
            space_bef_pt = max(2, space_bef_pt - 1)
        available_per_slot = body_h / max(len(points), 1)
        used_per_slot      = _total_h(b_size, space_bef_pt) / max(len(points), 1)
        slack_pt           = (available_per_slot - used_per_slot) * 72
        if vert_dist == 'spread' and slack_pt > 4:
            space_bef_pt = min(space_bef_pt + int(slack_pt * 0.5), 20)

        # Scale heading font proportionally to body size (but never smaller than 10pt)
        if _h_params:
            spec_body          = bs.get('font_size', 10)
            _h_params['size']  = max(10, round(_h_params['size'] * (b_size / max(spec_body, 1))))

    # Render heading now (font size finalised)
    if _h_params:
        add_text_box(slide, x, y + TOP_PAD * 0.25, w, 0.34,
                     str(heading), _h_params['font'], _h_params['size'],
                     _h_params['bold'], _h_params['color'], 'left', 'middle')

    def _bullet_char(point, idx):
        s = str(point)
        if list_style == 'tick_cross':
            if s[:1] in ('✓', '✗', '→'):
                return s[:1]
            return '✗' if any(neg in s.lower() for neg in
                ['not ', 'no ', 'declin', 'risk', 'fail', 'loss', 'below', 'miss']) else '✓'
        if list_style == 'numbered':
            return str(idx + 1) + '.'
        return '•'

    def _set_bullet(pPr, bchar, hang_emu):
        """Apply native OOXML hanging-indent bullet to an <a:pPr> element."""
        for tag in [nsmap.qn('a:buNone'), nsmap.qn('a:buChar'),
                    nsmap.qn('a:buAutoNum'), nsmap.qn('a:buFont')]:
            for el in pPr.findall(tag):
                pPr.remove(el)
        pPr.set('marL',   str(hang_emu))
        pPr.set('indent', str(-hang_emu))
        buFont = etree.SubElement(pPr, nsmap.qn('a:buFont'))
        buFont.set('typeface', b_font)
        buChar = etree.SubElement(pPr, nsmap.qn('a:buChar'))
        buChar.set('char', bchar)

    if points:
        _hang_emu = int(Emu(inches(BULLET_HANG)))

        # ALWAYS use a single text box for all points.
        # For "spread" distribution: use larger space_before on each paragraph so
        # points fill the available height — this keeps the text in one container
        # so the renderer can auto-fit the font if needed.
        txBox = slide.shapes.add_textbox(
            inches(x + indent_in), inches(content_y),
            inches(w - indent_in - INSIGHT_RIGHT_PADDING), inches(body_h)
        )
        tf = txBox.text_frame
        enable_text_fit(tf)

        for i, point in enumerate(points):
            if i == 0:
                p_obj = tf.paragraphs[0]
                # First paragraph: no space before (avoid top gap)
                p_obj.space_before = pt(0)
            else:
                p_obj = tf.add_paragraph()
                p_obj.space_before = pt(space_bef_pt)
            pPr = p_obj._p.get_or_add_pPr()
            _set_bullet(pPr, _bullet_char(point, i), _hang_emu)
            run = p_obj.add_run()
            run.text = _strip_marker(point)
            set_font(run, b_font, b_size, False, False, b_color)


def _apply_dual_axis(chart, secondary_series_names):
    """Post-process a clustered column chart XML to move named series onto a secondary
    line-chart plot with its own right-hand Y axis.

    python-pptx does not natively support combo charts, so we manipulate the XML directly.
    After this call the chart has two plot elements: barChart (primary) + lineChart (secondary).
    """
    plot_area = chart._element.find(nsmap.qn('c:plotArea'))
    if plot_area is None:
        return

    bar_chart = plot_area.find(nsmap.qn('c:barChart'))
    if bar_chart is None:
        return

    # Grab existing axId values (catAx id, primary valAx id)
    ax_id_els = bar_chart.findall(nsmap.qn('c:axId'))
    if len(ax_id_els) < 2:
        return
    cat_ax_id     = ax_id_els[0].get('val')
    prim_val_ax_id = ax_id_els[1].get('val')
    sec_val_ax_id  = str(int(prim_val_ax_id) + 1)   # unique secondary ID

    # Pull matching series elements out of barChart
    sec_ser_els = []
    for ser_el in list(bar_chart.findall(nsmap.qn('c:ser'))):
        tx = ser_el.find('.//' + nsmap.qn('c:v'))
        if tx is not None and tx.text in secondary_series_names:
            bar_chart.remove(ser_el)
            sec_ser_els.append(ser_el)

    if not sec_ser_els:
        return   # nothing matched — leave chart unchanged

    # Re-index remaining bar series
    for idx, ser_el in enumerate(bar_chart.findall(nsmap.qn('c:ser'))):
        for tag in (nsmap.qn('c:idx'), nsmap.qn('c:order')):
            el = ser_el.find(tag)
            if el is not None:
                el.set('val', str(idx))

    # Build lineChart element for secondary series
    line_chart = etree.SubElement(plot_area, nsmap.qn('c:lineChart'))
    etree.SubElement(line_chart, nsmap.qn('c:grouping')).set('val', 'standard')
    etree.SubElement(line_chart, nsmap.qn('c:varyColors')).set('val', '0')

    prim_bar_count = len(bar_chart.findall(nsmap.qn('c:ser')))
    for idx, ser_el in enumerate(sec_ser_els):
        for tag in (nsmap.qn('c:idx'), nsmap.qn('c:order')):
            el = ser_el.find(tag)
            if el is not None:
                el.set('val', str(prim_bar_count + idx))
        line_chart.append(ser_el)

    etree.SubElement(line_chart, nsmap.qn('c:marker')).set('val', '1')
    etree.SubElement(line_chart, nsmap.qn('c:smooth')).set('val', '0')
    etree.SubElement(line_chart, nsmap.qn('c:axId')).set('val', cat_ax_id)
    etree.SubElement(line_chart, nsmap.qn('c:axId')).set('val', sec_val_ax_id)

    # Add secondary valAx after existing valAx
    sec_val_ax = etree.SubElement(plot_area, nsmap.qn('c:valAx'))
    etree.SubElement(sec_val_ax, nsmap.qn('c:axId')).set('val', sec_val_ax_id)
    scaling = etree.SubElement(sec_val_ax, nsmap.qn('c:scaling'))
    etree.SubElement(scaling, nsmap.qn('c:orientation')).set('val', 'minMax')
    etree.SubElement(sec_val_ax, nsmap.qn('c:delete')).set('val', '0')
    etree.SubElement(sec_val_ax, nsmap.qn('c:axPos')).set('val', 'r')
    etree.SubElement(sec_val_ax, nsmap.qn('c:crossAx')).set('val', cat_ax_id)
    etree.SubElement(sec_val_ax, nsmap.qn('c:crosses')).set('val', 'max')
    etree.SubElement(sec_val_ax, nsmap.qn('c:crossBetween')).set('val', 'between')


def render_chart(slide, artifact, bt, suppress_heading=False):
    """Render a chart artifact using python-pptx native charts.

    suppress_heading — when True the artifact heading was already written to a layout
                       placeholder; the internal chart title is suppressed.
    """
    x  = artifact.get('x', 0)
    y  = artifact.get('y', 0)
    w  = artifact.get('w', 5)
    h  = artifact.get('h', 3)

    chart_type_str  = artifact.get('chart_type', 'bar')
    categories      = artifact.get('categories', [])
    series_data     = artifact.get('series', [])
    chart_title     = artifact.get('chart_title', '')
    show_labels     = artifact.get('show_data_labels', True)
    show_legend     = artifact.get('show_legend', False)
    series_styles   = artifact.get('series_style', [])
    cs              = artifact.get('chart_style', {})
    dual_axis       = artifact.get('dual_axis', False)
    secondary_names = set(artifact.get('secondary_series', []))
    chart_palette   = bt.get('chart_palette', ['#0F2FB5', '#FF8E00', '#2D962D', '#D60202'])

    # Map chart type string to XL_CHART_TYPE
    chart_type_map = {
        'bar':           XL_CHART_TYPE.COLUMN_CLUSTERED,
        'clustered_bar': XL_CHART_TYPE.COLUMN_CLUSTERED,
        'horizontal_bar':XL_CHART_TYPE.BAR_CLUSTERED,
        'line':          XL_CHART_TYPE.LINE,
        'pie':           XL_CHART_TYPE.PIE,
        'waterfall':     XL_CHART_TYPE.COLUMN_CLUSTERED,
    }
    xl_type = chart_type_map.get(chart_type_str, XL_CHART_TYPE.COLUMN_CLUSTERED)

    if not categories or not series_data:
        add_filled_rect(slide, x, y, w, h, fill_hex='#F0F4FF', border_hex='#C0C8E0', border_pt=1)
        add_text_box(slide, x + 0.1, y + h / 2 - 0.2, w - 0.2, 0.4,
                     'Chart: ' + chart_type_str + ' (no data)', 'Arial', 10,
                     False, '#888888', 'center', 'middle')
        return

    # Build ChartData
    cd = ChartData()
    cd.categories = [str(c) for c in categories]
    for si, ser in enumerate(series_data):
        ser_name   = ser.get('name', 'Series ' + str(si + 1))
        ser_values = [float(v) if v is not None else 0 for v in (ser.get('values') or [])]
        while len(ser_values) < len(categories): ser_values.append(0)
        cd.add_series(ser_name, ser_values[:len(categories)])

    # Add chart
    chart = slide.shapes.add_chart(
        xl_type, inches(x), inches(y), inches(w), inches(h), cd
    ).chart

    # Apply dual-axis layout (barChart primary + lineChart secondary with right Y axis)
    if dual_axis and secondary_names and chart_type_str not in ('pie', 'line'):
        try:
            _apply_dual_axis(chart, secondary_names)
        except Exception as e:
            print('dual axis error:', e)

    # Title — suppress when the zone header placeholder already carries the label
    show_title = bool(chart_title) and not suppress_heading
    chart.has_title = show_title
    if show_title:
        chart.chart_title.text_frame.text = str(chart_title)
        try:
            tf_run = chart.chart_title.text_frame.paragraphs[0].runs[0]
            set_font(tf_run,
                     cs.get('title_font_family', 'Arial'),
                     cs.get('title_font_size', 12),
                     True, False,
                     cs.get('title_color', '#000000'))
        except Exception:
            pass

    # Legend — position based on frame dimensions
    # Vertically stretched (tall, narrow): legend on top; everything else: right
    SLIDE_W, SLIDE_H = 10.0, 7.5
    legend_pos = 4   # right (XL_LEGEND_POSITION.RIGHT = 4)
    if h > SLIDE_H * 0.6 and w <= SLIDE_W * 0.6:
        legend_pos = 1   # top (XL_LEGEND_POSITION.TOP = 1)

    chart.has_legend = show_legend
    if show_legend and chart.has_legend:
        try:
            chart.legend.position = legend_pos
            chart.legend.include_in_layout = False
        except Exception:
            pass

    # Series colors + data labels
    try:
        for si, ser_obj in enumerate(chart.series):
            color_hex = None
            if si < len(series_styles):
                color_hex = series_styles[si].get('fill_color')
            if not color_hex:
                color_hex = chart_palette[si % len(chart_palette)]
            try:
                ser_obj.format.fill.solid()
                ser_obj.format.fill.fore_color.rgb = hex_to_rgb(color_hex)
            except Exception:
                pass
            if show_labels:
                try:
                    ser_obj.data_labels.show_value = True
                    lbl_color = (series_styles[si].get('data_label_color', '#000000')
                                 if si < len(series_styles) else '#000000')
                    for lbl in ser_obj.data_labels:
                        try:
                            lbl.font.color.rgb = hex_to_rgb(lbl_color)
                            lbl.font.size = Pt(8)
                        except Exception:
                            pass
                except Exception:
                    pass
    except Exception:
        pass

    # Pie: color each slice individually
    if chart_type_str == 'pie':
        try:
            for ser_obj in chart.series:
                for pi, point_obj in enumerate(ser_obj.points):
                    c_hex = (series_styles[pi].get('fill_color')
                             if pi < len(series_styles) else None) or \
                            chart_palette[pi % len(chart_palette)]
                    point_obj.format.fill.solid()
                    point_obj.format.fill.fore_color.rgb = hex_to_rgb(c_hex)
        except Exception:
            pass

    # Axis font sizes + category label rotation for many categories
    try:
        ax_font_size = cs.get('axis_font_size', 9)
        ax_color     = cs.get('axis_color', '#000000')
        for ax in (chart.category_axis, chart.value_axis):
            try:
                ax.tick_labels.font.size      = Pt(ax_font_size)
                ax.tick_labels.font.color.rgb = hex_to_rgb(ax_color)
            except Exception:
                pass

        # Rotate category labels when many categories to prevent overlap
        n_cats = len(categories)
        if chart_type_str in ('bar', 'line', 'clustered_bar') and n_cats > 6:
            try:
                # -2700000 EMU = -270 degrees = 45° slanted (OOXML uses 1/60000 degree units * -1)
                # rotation: -45 degrees = -2700000 in OOXML tickLblSkip/txPr
                cat_ax = chart.category_axis
                txPr = cat_ax.tick_labels._txPr
                if txPr is None:
                    # Create txPr if not present
                    txPr_elem = etree.SubElement(
                        cat_ax.tick_labels._element,
                        nsmap.qn('c:txPr')
                    )
                    etree.SubElement(txPr_elem, nsmap.qn('a:bodyPr')).set('rot', '-2700000')
                    etree.SubElement(txPr_elem, nsmap.qn('a:lstStyle'))
                    p = etree.SubElement(txPr_elem, nsmap.qn('a:p'))
                    pr = etree.SubElement(p, nsmap.qn('a:pPr'))
                    etree.SubElement(pr, nsmap.qn('a:defRPr'))
                else:
                    bodyPr = txPr.find(nsmap.qn('a:bodyPr'))
                    if bodyPr is not None:
                        bodyPr.set('rot', '-2700000')
            except Exception:
                pass
    except Exception:
        pass


def render_cards(slide, artifact, bt):
    """Render a cards artifact — each card as a styled rectangle + text."""
    cs       = artifact.get('card_style', {})
    ts       = artifact.get('title_style', {})
    sub_s    = artifact.get('subtitle_style', {})
    bs       = artifact.get('body_style', {})
    frames   = artifact.get('card_frames', [])
    cards    = artifact.get('cards', [])

    fill_hex   = cs.get('fill_color', '#F5F5F5')
    border_hex = cs.get('border_color') or '#E0E0E0'
    border_w   = cs.get('border_width', 0.75)
    corner_r   = cs.get('corner_radius', 4)
    padding    = cs.get('internal_padding', CARD_INNER_PADDING)

    t_font  = ts.get('font_family', bt.get('title_font_family', 'Arial'))
    t_size  = ts.get('font_size', 10)
    t_color = ts.get('color', bt.get('primary_color', '#1A3C8F'))
    t_bold  = ts.get('font_weight', 'bold') in ('bold', 'semibold')

    su_font  = sub_s.get('font_family', bt.get('body_font_family', 'Arial'))
    su_size  = sub_s.get('font_size', 20)
    su_color = sub_s.get('color', '#000000')

    b_font  = bs.get('font_family', bt.get('body_font_family', 'Arial'))
    b_size  = bs.get('font_size', 9)
    b_color = bs.get('color', '#111111')

    # Sentinel fill colors for sentiment (override card_style.fill_color per card)
    _SENTIMENT_FILL = {
        'positive': '#EEF7F0',   # very light green
        'negative': '#FDF3F1',   # very light red
        'neutral':  None,        # use card_style fill as-is
    }

    for fi, frame in enumerate(frames):
        fx = frame.get('x', 0)
        fy = frame.get('y', 0)
        fw = frame.get('w', 2)
        fh = frame.get('h', 1)

        card = cards[fi] if fi < len(cards) else {}

        # Derive per-card fill based on sentiment (if fill not explicitly set in card_style)
        sentiment  = card.get('sentiment', 'neutral')
        card_fill  = _SENTIMENT_FILL.get(sentiment) or fill_hex

        # Card background
        add_filled_rect(slide, fx, fy, fw, fh,
                        fill_hex=card_fill,
                        border_hex=border_hex,
                        border_pt=border_w,
                        corner_radius=corner_r)

        # Top accent strip (4px) — color matches sentiment
        _SENTIMENT_ACCENT = {
            'positive': '#2D8A4E',   # green
            'negative': '#C0392B',   # red
            'neutral':  None,
        }
        accent = _SENTIMENT_ACCENT.get(sentiment) or ts.get('color', bt.get('primary_color', '#1A3C8F'))
        accent_h = 0.055
        add_filled_rect(slide, fx, fy, fw, accent_h, fill_hex=accent)

        # Proportional inner layout — works for any card height
        inner_top = fy + accent_h + padding
        inner_h   = fh - accent_h - 2 * padding

        card_title = card.get('title', '')
        card_sub   = card.get('subtitle', '')
        card_body  = card.get('body', '')

        has_sub  = bool(card_sub)
        has_body = bool(card_body)

        if inner_h > 0.05:
            if has_sub and has_body:
                title_h = inner_h * 0.20
                sub_h   = inner_h * 0.40
                body_h  = max(0.12, inner_h - title_h - sub_h - TITLE_TO_SUBTITLE - SUBTITLE_TO_BODY)
            elif has_sub:
                title_h = inner_h * 0.28
                sub_h   = inner_h - title_h - TITLE_TO_SUBTITLE
                body_h  = 0
            elif has_body:
                title_h = inner_h * 0.25
                sub_h   = 0
                body_h  = inner_h - title_h - TITLE_TO_SUBTITLE
            else:
                title_h = inner_h
                sub_h   = 0
                body_h  = 0

            actual_title_size = estimate_fit_font_size(str(card_title), max(0.3, fw - padding*2), max(0.14, title_h), t_size, 8) if card_title else t_size
            actual_su_size = estimate_fit_font_size(str(card_sub), max(0.3, fw - padding*2), max(0.18, sub_h), su_size, 12) if sub_h > 0 else su_size
            actual_body_size = estimate_fit_font_size(str(card_body), max(0.3, fw - padding*2), max(0.12, body_h), b_size, 7) if body_h > 0 else b_size

            cur_y = inner_top
            if card_title:
                add_text_box(slide, fx + padding, cur_y, fw - padding*2, title_h,
                             str(card_title), t_font, actual_title_size, t_bold, t_color, 'left', 'top')
            cur_y += title_h + (TITLE_TO_SUBTITLE if has_sub else 0)

            if card_sub and sub_h > 0:
                add_text_box(slide, fx + padding, cur_y, fw - padding*2, sub_h,
                             str(card_sub), su_font, actual_su_size, False, su_color, 'left', 'top')
            cur_y += sub_h + (SUBTITLE_TO_BODY if has_body else 0)

            if card_body and body_h > 0.05:
                add_text_box(slide, fx + padding, cur_y, fw - padding*2, body_h,
                             str(card_body), b_font, actual_body_size, False, b_color, 'left', 'top')


def render_workflow(slide, artifact, bt):
    """Render a workflow artifact — nodes as rectangles, connections as lines."""
    ws    = artifact.get('workflow_style', {})
    nodes = artifact.get('nodes', [])
    conns = artifact.get('connections', [])
    flow_direction = str(artifact.get('flow_direction', '') or '').lower()
    workflow_type = str(artifact.get('workflow_type', '') or '').lower()

    def _num(v, default=0.0):
        try:
            return float(v)
        except Exception:
            return default

    def _has_rect(rect):
        return isinstance(rect, dict) and all(rect.get(k) is not None for k in ('x', 'y', 'w', 'h'))

    def _workflow_bounds(nodes_, conns_):
        xs, ys, xe, ye = [], [], [], []
        for node in nodes_:
            nx = _num(node.get('x'))
            ny = _num(node.get('y'))
            nw = max(0.01, _num(node.get('w'), 2.0))
            nh = max(0.01, _num(node.get('h'), 0.8))
            xs.append(nx); ys.append(ny); xe.append(nx + nw); ye.append(ny + nh)
        for conn in conns_:
            for pt in conn.get('path', []):
                px = _num(pt.get('x'))
                py = _num(pt.get('y'))
                xs.append(px); ys.append(py); xe.append(px); ye.append(py)
        if not xs:
            return None
        return {
            'x': min(xs), 'y': min(ys),
            'w': max(0.01, max(xe) - min(xs)),
            'h': max(0.01, max(ye) - min(ys))
        }

    def _normalize_workflow_to_rect(nodes_, conns_, target_rect):
        source_rect = _workflow_bounds(nodes_, conns_)
        if not source_rect or not _has_rect(target_rect):
            return nodes_, conns_

        pad_x = min(0.12, max(0.03, target_rect['w'] * 0.04))
        pad_y = min(0.12, max(0.03, target_rect['h'] * 0.06))
        usable_w = max(0.05, target_rect['w'] - 2 * pad_x)
        usable_h = max(0.05, target_rect['h'] - 2 * pad_y)
        sx = usable_w / max(source_rect['w'], 0.01)
        sy = usable_h / max(source_rect['h'], 0.01)

        def _map_x(v):
            return target_rect['x'] + pad_x + (_num(v) - source_rect['x']) * sx

        def _map_y(v):
            return target_rect['y'] + pad_y + (_num(v) - source_rect['y']) * sy

        mapped_nodes = []
        for node in nodes_:
            mapped_nodes.append({
                **node,
                'x': _map_x(node.get('x')),
                'y': _map_y(node.get('y')),
                'w': max(0.20, _num(node.get('w'), 2.0) * sx),
                'h': max(0.20, _num(node.get('h'), 0.8) * sy),
            })

        mapped_conns = []
        for conn in conns_:
            mapped_conns.append({
                **conn,
                'path': [
                    {'x': _map_x(pt.get('x')), 'y': _map_y(pt.get('y'))}
                    for pt in conn.get('path', [])
                ]
            })

        return mapped_nodes, mapped_conns

    def _layout_horizontal_timeline(nodes_, target_rect):
        if not nodes_ or not _has_rect(target_rect):
            return nodes_, []

        count = max(1, len(nodes_))
        pad_x = min(0.10, max(0.04, target_rect['w'] * 0.025))
        pad_y = min(0.10, max(0.04, target_rect['h'] * 0.04))
        gap = min(0.12, max(0.05, target_rect['w'] * 0.02))
        usable_w = max(0.3, target_rect['w'] - 2 * pad_x - gap * (count - 1))
        node_w = max(0.75, usable_w / count)
        node_h = max(0.70, target_rect['h'] - 2 * pad_y)
        start_x = target_rect['x'] + pad_x
        node_y = target_rect['y'] + max(0.02, (target_rect['h'] - node_h) / 2.0)

        laid_out = []
        for i, node in enumerate(nodes_):
            laid_out.append({
                **node,
                'x': start_x + i * (node_w + gap),
                'y': node_y,
                'w': node_w,
                'h': node_h
            })

        conn_y = node_y + node_h / 2.0
        laid_conns = []
        for i in range(len(laid_out) - 1):
            cur = laid_out[i]
            nxt = laid_out[i + 1]
            laid_conns.append({
                'from': cur.get('id'),
                'to': nxt.get('id'),
                'type': 'arrow',
                'path': [
                    {'x': cur['x'] + cur['w'], 'y': conn_y},
                    {'x': nxt['x'], 'y': conn_y}
                ]
            })

        return laid_out, laid_conns

    target_rect = None
    container = artifact.get('container')
    if _has_rect(container):
        target_rect = {
            'x': _num(container.get('x')),
            'y': _num(container.get('y')),
            'w': max(0.05, _num(container.get('w'))),
            'h': max(0.05, _num(container.get('h')))
        }
    elif all(artifact.get(k) is not None for k in ('x', 'y', 'w', 'h')):
        target_rect = {
            'x': _num(artifact.get('x')),
            'y': _num(artifact.get('y')),
            'w': max(0.05, _num(artifact.get('w'))),
            'h': max(0.05, _num(artifact.get('h')))
        }
    if target_rect:
        if flow_direction in ('left_to_right', 'horizontal') or workflow_type in ('timeline', 'roadmap'):
            nodes, conns = _layout_horizontal_timeline(nodes, target_rect)
        else:
            nodes, conns = _normalize_workflow_to_rect(nodes, conns, target_rect)

    node_fill   = ws.get('node_fill_color', bt.get('primary_color', '#1A3C8F'))
    node_border = ws.get('node_border_color', '#FFFFFF')
    node_bw     = ws.get('node_border_width', 1.5)
    node_cr     = ws.get('node_corner_radius', 4)
    conn_color  = ws.get('connector_color', bt.get('primary_color', '#1A3C8F'))

    t_font  = ws.get('node_title_font_family', bt.get('title_font_family', 'Arial'))
    t_size  = ws.get('node_title_font_size', 11)
    t_color = ws.get('node_title_color', '#FFFFFF')
    t_bold  = ws.get('node_title_font_weight', 'bold') in ('bold', 'semibold')

    v_font  = ws.get('node_value_font_family', bt.get('body_font_family', 'Arial'))
    v_size  = ws.get('node_value_font_size', 10)
    v_color = ws.get('node_value_color', '#FFFFFF')

    # Draw connections first (behind nodes)
    for conn in conns:
        path = conn.get('path', [])
        if len(path) >= 2:
            try:
                from pptx.util import Emu
                from pptx.oxml.ns import qn
                # Use a connector shape for simple lines
                x1 = float(path[0].get('x', 0))
                y1 = float(path[0].get('y', 0))
                x2 = float(path[-1].get('x', 0))
                y2 = float(path[-1].get('y', 0))

                connector = slide.shapes.add_connector(
                    1,  # MSO_CONNECTOR_TYPE.STRAIGHT
                    inches(x1), inches(y1), inches(x2), inches(y2)
                )
                connector.line.color.rgb = hex_to_rgb(conn_color)
                connector.line.width     = Pt(ws.get('connector_width', 2))
            except Exception:
                pass

    # Draw nodes
    node_map = {n.get('id'): n for n in nodes}
    for node in nodes:
        nx = node.get('x', 0)
        ny = node.get('y', 0)
        nw = node.get('w', 2)
        nh = node.get('h', 0.8)

        add_filled_rect(slide, nx, ny, nw, nh,
                        fill_hex=node_fill,
                        border_hex=node_border,
                        border_pt=node_bw,
                        corner_radius=node_cr)

        # Node label (primary text)
        label = node.get('label', node.get('id', ''))
        has_desc = bool(node.get('description', ''))
        title_h = nh * (0.28 if has_desc else 0.34)
        value_h = nh * (0.22 if has_desc else 0.26)
        desc_h  = max(0.16, nh - title_h - value_h - 0.12)
        if label:
            label_size = estimate_fit_font_size(str(label), max(0.3, nw - 0.2), max(0.18, title_h), t_size, 8)
            add_text_box(slide, nx + 0.1, ny + 0.06, nw - 0.2, title_h,
                         str(label), t_font, label_size, t_bold, t_color, 'center', 'middle')

        # Node value (secondary metric)
        value = node.get('value', '')
        if value:
            value_size = estimate_fit_font_size(str(value), max(0.3, nw - 0.2), max(0.16, value_h), v_size, 7)
            add_text_box(slide, nx + 0.1, ny + 0.06 + title_h, nw - 0.2, value_h,
                         str(value), v_font, value_size, False, v_color, 'center', 'middle')

        # Description (small text below value)
        desc = node.get('description', '')
        if desc and desc_h > 0.14:
            desc_size = estimate_fit_font_size(str(desc), max(0.28, nw - 0.16), desc_h, max(7, v_size - 2), 6)
            add_text_box(slide, nx + 0.08, ny + 0.08 + title_h + value_h, nw - 0.16, desc_h,
                         str(desc), v_font, desc_size, False, v_color, 'center', 'top')


def render_table(slide, artifact, bt):
    """Render a table artifact using python-pptx native table."""
    x = artifact.get('x', 0)
    y = artifact.get('y', 0)
    w = artifact.get('w', 6)
    h = artifact.get('h', 2)

    ts      = artifact.get('table_style', {})
    headers = artifact.get('headers', [])
    rows    = artifact.get('rows', [])
    col_ws  = artifact.get('column_widths', [])
    row_hs  = artifact.get('row_heights', [])
    hl_rows = artifact.get('highlight_rows', [])

    if not headers:
        add_text_box(slide, x, y, w, h, 'Table (no headers)', 'Arial', 10,
                     False, '#888888', 'center', 'middle')
        return

    n_cols  = len(headers)
    n_rows  = len(rows) + 1   # +1 for header row
    data_rows = rows

    def _norm_sizes(values, count, total):
        nums = []
        for i in range(count):
            try:
                v = float(values[i]) if i < len(values) else 0
            except Exception:
                v = 0
            nums.append(max(0, v))
        summed = sum(nums)
        if summed > 0:
            return [(v / summed) * total for v in nums]
        return [total / max(count, 1)] * count

    def _auto_col_widths():
        weights = []
        for ci, hdr in enumerate(headers):
            max_len = len(str(hdr or ''))
            for row in data_rows:
                if ci < len(row):
                    max_len = max(max_len, len(str(row[ci] or '')))
            weights.append(max(6, min(max_len, 40)))
        total_weight = sum(weights) or 1
        return [w * (wt / total_weight) for wt in weights]

    norm_col_ws = _auto_col_widths()
    norm_row_hs = [max(TABLE_MIN_ROW_HEIGHT, rh) for rh in _norm_sizes(row_hs, n_rows, h)]
    row_total = sum(norm_row_hs)
    if row_total > h:
        norm_row_hs = [rh * (h / row_total) for rh in norm_row_hs]

    try:
        table_shape = slide.shapes.add_table(n_rows, n_cols, inches(x), inches(y), inches(w), inches(h))
        table       = table_shape.table

        # Column widths
        for ci, cw in enumerate(norm_col_ws):
            table.columns[ci].width = inches(cw)

        # Row heights
        for ri, rh in enumerate(norm_row_hs):
            table.rows[ri].height = inches(rh)

        # Header row styles
        h_fill  = ts.get('header_fill_color', bt.get('primary_color', '#1A3C8F'))
        h_text  = ts.get('header_text_color', '#FFFFFF')
        h_font  = ts.get('header_font_family', bt.get('title_font_family', 'Arial'))
        h_size  = ts.get('header_font_size', 11)

        b_fill  = ts.get('body_fill_color', '#FFFFFF')
        b_alt   = ts.get('body_alt_fill_color')
        b_text  = ts.get('body_text_color', '#111111')
        b_font  = ts.get('body_font_family', bt.get('body_font_family', 'Arial'))
        b_size  = ts.get('body_font_size', 10)
        hl_fill = ts.get('highlight_fill_color')
        grid_c  = ts.get('grid_color', '#CCCCCC')
        cell_p  = max(0.08, min(0.10, ts.get('cell_padding', CELL_PADDING) or CELL_PADDING))

        # Detect which columns are numeric (right-align) vs text (left-align)
        # A column is "numeric" if > 50% of its body values look like numbers/currency
        def _is_numeric_col(col_idx):
            import re
            _num_pat = re.compile(r'^[\s₹$€£¥\-\+]?[\d,\.]+[%KMBCr\s]*$')
            hits = sum(1 for r in data_rows
                       if col_idx < len(r) and _num_pat.match(str(r[col_idx]).strip()))
            return (hits / max(len(data_rows), 1)) > 0.5

        col_is_numeric = [_is_numeric_col(ci) for ci in range(n_cols)]
        # First column is always text (entity/label column)
        if n_cols > 0:
            col_is_numeric[0] = False

        cell_pad_emu = int(Emu(cell_p * 914400))  # cell_p in inches → EMU

        def _apply_cell_margin(cell, pad_emu):
            """Apply uniform cell margin via OOXML tcPr."""
            try:
                tc = cell._tc
                tcPr = tc.get_or_add_tcPr()
                for attr in ('marL', 'marR', 'marT', 'marB'):
                    tcPr.set(attr, str(pad_emu))
            except Exception:
                pass

        for ci, hdr in enumerate(headers):
            cell = table.cell(0, ci)
            cell.text = str(hdr)
            enable_text_fit(cell.text_frame)
            cell.fill.solid()
            cell.fill.fore_color.rgb = hex_to_rgb(h_fill)
            _apply_cell_margin(cell, cell_pad_emu)
            try:
                run = cell.text_frame.paragraphs[0].runs[0]
                fit_size = estimate_fit_font_size(
                    hdr,
                    norm_col_ws[ci],
                    norm_row_hs[0],
                    h_size,
                    8
                )
                set_font(run, h_font, fit_size, True, False, h_text)
                # Header row: center-align all columns
                cell.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
            except Exception:
                pass

        # Body rows
        for ri, row in enumerate(data_rows):
            row_idx = ri + 1
            use_alt = (b_alt and ri % 2 == 1)
            use_hl  = (hl_fill and ri in hl_rows)
            row_fill = hl_fill if use_hl else (b_alt if use_alt else b_fill)

            for ci in range(n_cols):
                cell = table.cell(row_idx, ci)
                cell_text = str(row[ci]) if ci < len(row) else ''
                cell.text = cell_text
                enable_text_fit(cell.text_frame)
                cell.fill.solid()
                cell.fill.fore_color.rgb = hex_to_rgb(row_fill or b_fill)
                _apply_cell_margin(cell, cell_pad_emu)
                try:
                    run = cell.text_frame.paragraphs[0].runs[0]
                    fit_size = estimate_fit_font_size(
                        cell_text,
                        norm_col_ws[ci],
                        norm_row_hs[row_idx],
                        b_size,
                        7
                    )
                    set_font(run, b_font, fit_size, False, False, b_text)
                    # Numeric columns → right-align; text columns → left-align
                    align = PP_ALIGN.RIGHT if col_is_numeric[ci] else PP_ALIGN.LEFT
                    cell.text_frame.paragraphs[0].alignment = align
                except Exception:
                    pass

    except Exception as e:
        # Fallback — plain text box
        all_text = ' | '.join(headers) + '\n' + '\n'.join([' | '.join([str(c) for c in r]) for r in rows])
        add_text_box(slide, x, y, w, h, all_text, 'Arial', 9, False, '#111111')


def _write_heading_to_header_ph(slide, heading_text, header_ph_idx, bt, header_style='underline'):
    """Write heading text into the layout's paired header placeholder."""
    if not heading_text or header_ph_idx is None:
        return None
    try:
        for ph in slide.placeholders:
            if ph.placeholder_format.idx == header_ph_idx:
                ph_x = ph.left / 914400
                ph_y = ph.top / 914400
                ph_w = ph.width / 914400
                ph_h = ph.height / 914400
                ph.text_frame.text = str(heading_text)
                try:
                    run = ph.text_frame.paragraphs[0].runs[0]
                    from pptx.util import Pt
                    from pptx.dml.color import RGBColor
                    run.font.name = bt.get('title_font_family', 'Arial')
                    run.font.size = Pt(10)
                    run.font.bold = True
                    color_hex = bt.get('primary_color', '#1A3C8F').lstrip('#')
                    run.font.color.rgb = RGBColor(
                        int(color_hex[0:2], 16),
                        int(color_hex[2:4], 16),
                        int(color_hex[4:6], 16)
                    )
                except Exception:
                    pass
                try:
                    font_size = 10
                    text_h = estimate_header_block_height(heading_text, ph_w, font_size)
                    if header_style == 'brand_fill':
                        add_filled_rect(
                            slide,
                            ph_x,
                            ph_y,
                            ph_w,
                            HEADER_HEIGHT,
                            fill_hex=bt.get('primary_color', '#1A3C8F')
                        )
                        return ph_y + HEADER_HEIGHT
                    else:
                        rule_y = ph_y + text_h + 0.02
                        add_filled_rect(
                            slide,
                            ph_x,
                            rule_y,
                            ph_w,
                            0.03,
                            fill_hex=bt.get('primary_color', '#1A3C8F')
                        )
                        return rule_y + 0.03
                except Exception:
                    pass
                return ph_y + max(text_h, HEADER_HEIGHT)
    except Exception as e:
        print(f'_write_heading_to_header_ph error (idx={header_ph_idx}):', e)
    return None


def render_artifact(slide, artifact, bt, ph_frame=None, header_ph_idx=None, header_style='underline'):
    """
    Dispatch to the correct renderer based on artifact type.

    ph_frame      — dict with x/y/w/h keys (inches) pre-read from the layout placeholder
                    BEFORE that placeholder was removed from the slide.  When set, these
                    bounds override the artifact's own x/y/w/h (e.g. when Agent 5 left
                    them null in layout mode, trusting the placeholder for positioning).
    header_ph_idx — layout mode: write artifact heading into the paired header placeholder
                    that was preserved in the slide (not removed).
    """
    t = (artifact.get('type') or '').lower()

    artifact = dict(artifact)   # shallow copy — don't mutate the spec

    # Apply pre-saved placeholder bounds when the artifact's own coords are null/missing
    if ph_frame is not None:
        artifact['x'] = ph_frame['x']
        artifact['y'] = ph_frame['y']
        artifact['w'] = ph_frame['w']
        artifact['h'] = ph_frame['h']

    # Route artifact heading into the layout's paired header placeholder when available
    heading_handled = False
    placeholder_header_bottom = None
    inline_header_rendered = False
    rendered_header_bottom = None
    header_block = artifact.get('header_block') or {}
    use_placeholder_header = bool(header_block.get('placeholder_ref'))
    if header_ph_idx is not None and t != 'cards' and use_placeholder_header:
        if t == 'insight_text':
            heading_text = artifact.get('heading') or artifact.get('insight_header', '')
        elif t == 'chart':
            heading_text = artifact.get('chart_header', '')
        elif t == 'table':
            heading_text = artifact.get('table_header', '')
        elif t == 'workflow':
            heading_text = artifact.get('workflow_header', '')
        else:
            heading_text = ''
        if heading_text:
            placeholder_header_bottom = _write_heading_to_header_ph(slide, heading_text, header_ph_idx, bt, header_style=header_style)
            heading_handled = placeholder_header_bottom is not None

    # Render header_block only when the heading wasn't routed to a layout placeholder
    if t != 'cards' and not heading_handled:
        hb = dict(header_block)
        if hb and artifact.get('x') is not None and artifact.get('w') is not None:
            hb['x'] = artifact.get('x')
            hb['w'] = artifact.get('w')
        if hb and not hb.get('placeholder_ref'):
            rendered_header_bottom = render_header_block(slide, hb, bt, header_style=header_style)
            inline_header_rendered = True

    suppress_internal_heading = heading_handled or inline_header_rendered
    if inline_header_rendered and header_block:
        gap = 0.08
        header_bottom = float(rendered_header_bottom or 0) + gap
        art_y = float(artifact.get('y', 0) or 0)
        art_h = float(artifact.get('h', 0) or 0)
        if art_h > 0 and header_bottom > art_y:
            delta = header_bottom - art_y
            artifact['y'] = header_bottom
            artifact['h'] = max(0.2, art_h - delta)
    elif heading_handled:
        # Only introduce the minimum gap required after the actual rendered header.
        art_y = float(artifact.get('y', 0) or 0)
        art_h = float(artifact.get('h', 0) or 0)
        required_top = float(placeholder_header_bottom or art_y) + HEADER_TO_ARTIFACT
        if art_h > 0 and required_top > art_y:
            delta = required_top - art_y
            artifact['y'] = required_top
            artifact['h'] = max(0.2, art_h - delta)

    try:
        if   t == 'insight_text': render_insight_text(slide, artifact, bt,
                                                       suppress_heading=suppress_internal_heading)
        elif t == 'chart':        render_chart(slide, artifact, bt,
                                               suppress_heading=suppress_internal_heading)
        elif t == 'cards':        render_cards(slide, artifact, bt)
        elif t == 'workflow':     render_workflow(slide, artifact, bt)
        elif t == 'table':        render_table(slide, artifact, bt)
    except Exception as e:
        print(f'render_artifact error ({t}):', e)
        traceback.print_exc()


# ─── SLIDE BUILDER ─────────────────────────────────────────────────────────────

def _write_speaker_note(slide, note_text):
    if note_text:
        try:
            slide.notes_slide.notes_text_frame.text = str(note_text)
        except Exception:
            pass


def _remove_content_placeholders(slide, keep_indices=None):
    """
    Upfront removal of non-system content/body placeholders from the slide XML.

    Layout templates often have descriptive default text baked into their
    content placeholders (e.g. "z1 · summary", "primary") for design guidance.
    Setting text = '' is unreliable because python-pptx may fall through to the
    layout's inherited txBody.  The only guaranteed fix is to remove the
    placeholder shape element from the slide's spTree entirely.

    keep_indices — set of placeholder.idx values to PRESERVE (these will be
                   written to by _write_heading_to_header_ph for heading text).
                   All other non-system placeholders are removed.
    """
    from pptx.enum.shapes import PP_PLACEHOLDER as _PH
    _SYSTEM_TYPES = {_PH.TITLE, _PH.CENTER_TITLE, _PH.SUBTITLE,
                     _PH.DATE, _PH.FOOTER, _PH.SLIDE_NUMBER}
    keep_indices = keep_indices or set()

    sp_tree = slide.shapes._spTree
    to_remove = []
    for ph in list(slide.placeholders):
        try:
            ph_type = ph.placeholder_format.type
            ph_idx  = ph.placeholder_format.idx
            if ph_type not in _SYSTEM_TYPES and ph_idx not in keep_indices:
                to_remove.append(ph._element)
        except Exception:
            pass
    for elem in to_remove:
        try:
            sp_tree.remove(elem)
        except Exception:
            pass


def _remove_empty_placeholders(slide):
    """
    Remove placeholder shapes that carry no content.
    After rendering, any layout placeholder that was not explicitly written to
    would appear as a 'Click to add text' ghost box overlaid on the slide.
    This is safe because we render all content as free shapes or via
    place_in_placeholder — so a placeholder is only needed when it has text.
    """
    sp_tree = slide.shapes._spTree
    to_remove = []
    for ph in list(slide.placeholders):
        try:
            if not ph.text_frame.text.strip():
                to_remove.append(ph._element)
        except Exception:
            # Non-text placeholder (picture / object type) — treat as unfilled
            to_remove.append(ph._element)
    for elem in to_remove:
        try:
            sp_tree.remove(elem)
        except Exception:
            pass


def _sanitise_system_placeholders(slide, slide_number):
    """
    After add_slide(), fix placeholders that carry stale template text:
      - dt  (date/time)   → blank — presentation-level date field or master handles it
      - ftr (footer)      → blank — confidentiality text is on the master; don't duplicate
      - sldNum (slide #)  → write actual slide number so it shows correctly in PDF/export
    """
    try:
        from pptx.enum.shapes import PP_PLACEHOLDER as _PH
        for ph in slide.placeholders:
            ph_type = ph.placeholder_format.type
            if ph_type in (_PH.DATE, _PH.FOOTER):
                try:
                    tf = ph.text_frame
                    for para in tf.paragraphs:
                        for run in para.runs:
                            run.text = ''
                    # Also clear via XML to strip any fld elements carrying old date
                    for txBody in ph._element.iter('{http://schemas.openxmlformats.org/drawingml/2006/main}txBody'):
                        for p_el in list(txBody):
                            tag = p_el.tag.split('}')[-1]
                            if tag == 'p':
                                for child in list(p_el):
                                    ctag = child.tag.split('}')[-1]
                                    if ctag in ('r', 'fld'):
                                        p_el.remove(child)
                except Exception:
                    pass
            elif ph_type == _PH.SLIDE_NUMBER:
                try:
                    ph.text_frame.text = str(slide_number)
                except Exception:
                    pass
    except Exception:
        pass


def build_slide(prs, slide_spec, blank_layout, use_template=False,
                title_layout_name=None, divider_layout_name=None):
    """
    Build one slide from its spec object.

    Three rendering paths:
    1. Title / divider  — template mode: text into placeholders only.
    2. Content          — template mode: title/subtitle into placeholders (idx 0/1);
                          artifacts rendered at coordinates OR at placeholder bounds
                          when layout_mode=True and placeholder_idx is set.
    3. Any type         — scratch mode (use_template=False): full coordinates for
                          everything including title, subtitle, logo, footer, page num.
    """
    slide_type           = slide_spec.get('slide_type', 'content')
    layout_mode          = slide_spec.get('layout_mode', False)
    selected_layout_name = slide_spec.get('selected_layout_name', '')
    bt                   = slide_spec.get('brand_tokens', {})
    cvs                  = slide_spec.get('canvas', {})
    tb                   = slide_spec.get('title_block') or {}
    sb                   = slide_spec.get('subtitle_block') or {}
    slide_header_style   = infer_slide_header_style(slide_spec)

    # ── Choose layout ────────────────────────────────────────────────────────
    layout = blank_layout
    if use_template:
        if selected_layout_name:
            # Content slide with a named layout chosen by Agent 4
            named = find_layout_by_name(prs, selected_layout_name)
            if named:
                layout = named
                print(f'  Slide {slide_spec.get("slide_number","?")}: layout="{named.name}"')

        elif slide_type == 'title':
            # Use the explicit name from Agent 2's rulebook first (most reliable),
            # then fall back to OOXML type="title" matching, then name heuristics.
            if title_layout_name:
                layout = find_layout_by_name(prs, title_layout_name) or blank_layout
            else:
                for lyt in prs.slide_layouts:
                    # python-pptx exposes the OOXML type via the internal XML
                    try:
                        import pptx.oxml.ns as _ns
                        lyt_type = lyt._element.get('type', '')
                    except Exception:
                        lyt_type = ''
                    if lyt_type == 'title':
                        layout = lyt; break
                else:
                    for lyt in prs.slide_layouts:
                        lname = (lyt.name or '').lower()
                        if 'title' in lname and 'section' not in lname and 'content' not in lname:
                            layout = lyt; break

        elif slide_type == 'divider':
            # Same priority: explicit rulebook name → OOXML secHead → name heuristics.
            if divider_layout_name:
                layout = find_layout_by_name(prs, divider_layout_name) or blank_layout
            else:
                for lyt in prs.slide_layouts:
                    try:
                        import pptx.oxml.ns as _ns
                        lyt_type = lyt._element.get('type', '')
                    except Exception:
                        lyt_type = ''
                    if lyt_type == 'secHead':
                        layout = lyt; break
                else:
                    for lyt in prs.slide_layouts:
                        lname = (lyt.name or '').lower()
                        if 'section' in lname or 'divider' in lname:
                            layout = lyt; break

    slide = prs.slides.add_slide(layout)
    if use_template:
        from pptx.enum.shapes import PP_PLACEHOLDER as _PH
        _sanitise_system_placeholders(slide, slide_spec.get('slide_number', ''))
        # ── Snapshot placeholder bounds BEFORE any removal ───────────────────
        # render_artifact needs the bounds of content placeholders (12, 13, …) for
        # positioning, but those placeholders must be removed to prevent ghost text.
        # We read all bounds now, then remove, then pass the snapshot to renderers.
        _ph_bounds = {}   # idx → {'x','y','w','h'} in inches
        _content_ph_frames = []  # ordered fallback slots for layout_mode when placeholder_idx is missing
        _SYSTEM_TYPES = {_PH.TITLE, _PH.CENTER_TITLE, _PH.SUBTITLE,
                         _PH.DATE, _PH.FOOTER, _PH.SLIDE_NUMBER}
        for _ph in slide.placeholders:
            try:
                _ph_idx = _ph.placeholder_format.idx
                _ph_type = _ph.placeholder_format.type
                _bounds = {
                    'x': _ph.left   / 914400,
                    'y': _ph.top    / 914400,
                    'w': _ph.width  / 914400,
                    'h': _ph.height / 914400,
                }
                _ph_bounds[_ph_idx] = _bounds
                # Some Agent 5 / 5.1 outputs reach Agent 6 without placeholder_idx.
                # Keep an ordered list of likely content slots so layout-mode content
                # can still render into the correct body area instead of disappearing.
                if (_ph_type not in _SYSTEM_TYPES and
                    _bounds['w'] >= 1.0 and _bounds['h'] >= 0.5):
                    _content_ph_frames.append({
                        'idx': _ph_idx,
                        **_bounds
                    })
            except Exception:
                pass
        _content_ph_frames.sort(key=lambda p: (round(p['y'] * 2), p['x'], p['idx']))

        # Collect header placeholder indices we will explicitly write heading text to.
        # These must survive removal so _write_heading_to_header_ph can use them.
        _header_ph_keep = set()
        for _z in slide_spec.get('zones', []):
            _hpi = _z.get('header_ph_idx')
            if _hpi is not None:
                _header_ph_keep.add(int(_hpi))
        _remove_content_placeholders(slide, keep_indices=_header_ph_keep)

    # ── Background (scratch only — master handles it in template mode) ────────
    if not use_template:
        bg_color = (cvs.get('background') or {}).get('color', '#FFFFFF')
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = hex_to_rgb(bg_color)

    # ════════════════════════════════════════════════════════════════════════
    # PATH 1 — TITLE & DIVIDER
    # Only write text content; the layout provides all visual decoration.
    # ════════════════════════════════════════════════════════════════════════
    if slide_type in ('title', 'divider'):
        if use_template:
            if tb.get('text'):
                place_in_placeholder(slide, 0, tb['text'], tb, bt)
            if sb.get('text'):
                place_in_placeholder(slide, 1, sb['text'], sb, bt)
        else:
            # Scratch mode — free-form text boxes
            if tb.get('text'):
                add_text_box(slide,
                    tb.get('x', 0.4), tb.get('y', 0.15),
                    tb.get('w', 9.0), tb.get('h', 0.8),
                    tb['text'],
                    tb.get('font_family', bt.get('title_font_family', 'Arial')),
                    tb.get('font_size', 28),
                    tb.get('font_weight', 'bold') in ('bold', 'semibold'),
                    tb.get('color', bt.get('title_color', '#1A3C8F')),
                    tb.get('align', 'center'), 'middle')
            if sb.get('text'):
                add_text_box(slide,
                    sb.get('x', 0.4), sb.get('y', 1.2),
                    sb.get('w', 9.0), sb.get('h', 0.6),
                    sb['text'],
                    sb.get('font_family', bt.get('body_font_family', 'Arial')),
                    sb.get('font_size', 16),
                    False,
                    sb.get('color', bt.get('body_color', '#333333')),
                    sb.get('align', 'center'), 'middle')
        if use_template:
            _remove_empty_placeholders(slide)
        _write_speaker_note(slide, slide_spec.get('speaker_note', ''))
        return slide

    # ════════════════════════════════════════════════════════════════════════
    # PATH 2 & 3 — CONTENT SLIDES
    # ════════════════════════════════════════════════════════════════════════

    # Title & subtitle: placeholders in template mode, text boxes in scratch
    if use_template:
        if tb.get('text'):
            place_in_placeholder(slide, 0, tb['text'], tb, bt)
        if sb.get('text'):
            place_in_placeholder(slide, 1, sb['text'], sb, bt)
    else:
        if tb.get('text'):
            add_text_box(slide,
                tb.get('x', 0.4), tb.get('y', 0.15),
                tb.get('w', 9.0), tb.get('h', 0.8),
                tb['text'],
                tb.get('font_family', bt.get('title_font_family', 'Arial')),
                tb.get('font_size', 20),
                tb.get('font_weight', 'bold') in ('bold', 'semibold'),
                tb.get('color', bt.get('title_color', '#1A3C8F')),
                tb.get('align', 'left'), 'middle')
        if sb.get('text'):
            add_text_box(slide,
                sb.get('x', 0.4), sb.get('y', 1.0),
                sb.get('w', 9.0), sb.get('h', 0.5),
                sb['text'],
                sb.get('font_family', bt.get('body_font_family', 'Arial')),
                sb.get('font_size', 14),
                sb.get('font_weight', 'semibold') in ('bold', 'semibold'),
                sb.get('color', bt.get('body_color', '#333333')),
                sb.get('align', 'left'), 'middle')

    # Zones → Artifacts
    _layout_ph_bounds = _ph_bounds if (layout_mode and use_template) else {}
    for zone_idx, zone in enumerate(slide_spec.get('zones', [])):
        # In layout mode, each zone may have a paired header placeholder
        hdr_ph_idx = zone.get('header_ph_idx') if layout_mode else None
        zone_artifacts = zone.get('artifacts', [])
        for art_idx, artifact in enumerate(zone_artifacts):
            ph_frame = None
            if layout_mode:
                ph_idx_spec = artifact.get('placeholder_idx')
                # Only apply placeholder bounds when:
                # 1. The artifact has a placeholder_idx set
                # 2. AND the artifact's own x/y/w/h are null/missing (Agent 5 deferred to layout)
                # 3. AND the zone has exactly 1 artifact (1:1 mapping to placeholder slot)
                # For multi-artifact zones, Agent 5's explicit coordinates are authoritative.
                needs_bounds = (
                    ph_idx_spec is not None
                    and len(zone_artifacts) == 1
                    and artifact.get('x') is None
                )
                if needs_bounds and ph_idx_spec in _layout_ph_bounds:
                    ph_frame = _layout_ph_bounds[ph_idx_spec]
                elif artifact.get('x') is None and _content_ph_frames:
                    # Fallback for specs that lost placeholder_idx but still rely on
                    # layout_mode with null coordinates. Prefer zone order for 1:1
                    # layouts; for multi-artifact zones, step through slots by artifact.
                    fallback_slot = zone_idx if len(zone_artifacts) == 1 else (zone_idx + art_idx)
                    fallback_slot = min(fallback_slot, len(_content_ph_frames) - 1)
                    ph_frame = _content_ph_frames[fallback_slot]
            render_artifact(slide, artifact, bt, ph_frame=ph_frame, header_ph_idx=hdr_ph_idx, header_style=slide_header_style)

    # ── Global elements (scratch only) ────────────────────────────────────────
    if not use_template:
        ge = slide_spec.get('global_elements', {})

        logo = ge.get('logo', {})
        if logo.get('show') and logo.get('image_base64'):
            add_image_box(slide, logo['image_base64'],
                logo.get('x', 0.0), logo.get('y', 0.0),
                logo.get('w', 1.2), logo.get('h', 0.4))

        footer = ge.get('footer', {})
        if footer.get('show'):
            add_text_box(slide,
                footer.get('x', 0.4), footer.get('y', 7.5),
                footer.get('w', 5),   footer.get('h', 0.25),
                'Confidential',
                footer.get('font_family', 'Arial'),
                footer.get('font_size', 8),
                False, footer.get('color', '#AAAAAA'),
                footer.get('align', 'left'), 'middle')

        pn = ge.get('page_number', {})
        if pn.get('show'):
            add_text_box(slide,
                pn.get('x', 9.5), pn.get('y', 7.5),
                pn.get('w', 0.8), pn.get('h', 0.25),
                str(slide_spec.get('slide_number', '')),
                pn.get('font_family', 'Arial'),
                pn.get('font_size', 8),
                False, pn.get('color', '#AAAAAA'),
                'right', 'middle')

    if use_template:
        _remove_empty_placeholders(slide)
    _write_speaker_note(slide, slide_spec.get('speaker_note', ''))
    return slide


# ─── MAIN BUILDER ─────────────────────────────────────────────────────────────

def build_presentation(final_spec, brand_rulebook, template_b64=None):
    """
    Build a complete Presentation from the Agent 5 final spec.

    When template_b64 is provided the function:
      1. Loads the brand PPTX template (preserving slide masters)
      2. Strips all existing content slides from it
      3. Adds new slides backed by the template's blank layout

    This keeps the full brand identity (master background, decorative shapes,
    registered fonts, theme colours) without the need for Agent 5 to manually
    specify every design token.
    """
    if not final_spec:
        raise ValueError('finalSpec is empty')

    use_template = bool(template_b64)

    if use_template:
        print('build_presentation: loading brand template…')
        try:
            prs = load_template_prs(template_b64)
            print(f'  Template loaded — {len(prs.slide_masters)} master(s), {len(prs.slide_layouts)} layout(s)')
        except Exception as e:
            print('  Template load failed, falling back to blank presentation:', e)
            prs = Presentation()
            use_template = False
    else:
        prs = Presentation()

    # Slide dimensions — use template's native size if template is loaded,
    # otherwise take from the first slide's canvas spec
    if not use_template:
        first_canvas = (final_spec[0].get('canvas') or {})
        w_in = float(first_canvas.get('width_in')  or 11.02)
        h_in = float(first_canvas.get('height_in') or 8.29)
        prs.slide_width  = Inches(w_in)
        prs.slide_height = Inches(h_in)

    blank_layout = find_blank_layout(prs)
    print(f'  Using fallback layout: "{blank_layout.name}" (use_template={use_template})')

    # Extract explicit layout names from Agent 2 rulebook so build_slide can
    # select title / divider layouts reliably without keyword heuristics.
    title_layout_name   = brand_rulebook.get('title_layout_name')
    divider_layout_name = brand_rulebook.get('divider_layout_name')
    if title_layout_name:
        print(f'  Title layout:   "{title_layout_name}"')
    if divider_layout_name:
        print(f'  Divider layout: "{divider_layout_name}"')

    for slide_spec in final_spec:
        try:
            build_slide(prs, slide_spec, blank_layout, use_template,
                        title_layout_name=title_layout_name,
                        divider_layout_name=divider_layout_name)
        except Exception as e:
            print(f"Error building slide {slide_spec.get('slide_number', '?')}:", e)
            traceback.print_exc()
            prs.slides.add_slide(blank_layout)

    return prs


# ─── VERCEL HANDLER ──────────────────────────────────────────────────────────

class handler(BaseHTTPRequestHandler):

    def do_POST(self):
        try:
            length = int(self.headers.get('Content-Length', 0))
            body   = self.rfile.read(length)
            data   = json.loads(body)

            final_spec     = data.get('finalSpec', [])
            brand_rulebook = data.get('brandRulebook', {})
            template_b64   = data.get('templateB64') or None

            if not final_spec:
                self._json(400, {'success': False, 'error': 'finalSpec is missing or empty'})
                return

            # Build the presentation (optionally from a brand template)
            prs = build_presentation(final_spec, brand_rulebook, template_b64)

            # Serialize to bytes
            buf = io.BytesIO()
            prs.save(buf)
            buf.seek(0)
            pptx_bytes = buf.read()

            # Encode as base64
            b64 = base64.b64encode(pptx_bytes).decode('utf-8')

            # Filename
            date_str = datetime.now().strftime('%Y%m%d')
            title_raw = (final_spec[0].get('title') or 'presentation')
            title_slug = ''.join(c if c.isalnum() else '_' for c in title_raw.lower())[:30]
            filename = title_slug + '_' + date_str + '.pptx'

            self._json(200, {
                'success':  True,
                'data':     b64,
                'slides':   len(final_spec),
                'filename': filename
            })

        except Exception as e:
            traceback.print_exc()
            self._json(500, {'success': False, 'error': str(e)})

    def _json(self, status, payload):
        body = json.dumps(payload).encode('utf-8')
        self.send_response(status)
        self.send_header('Content-Type',   'application/json')
        self.send_header('Content-Length', str(len(body)))
        self.send_header('Access-Control-Allow-Origin', '*')
        self.end_headers()
        self.wfile.write(body)

    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header('Access-Control-Allow-Origin',  '*')
        self.send_header('Access-Control-Allow-Methods', 'POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.end_headers()

    def log_message(self, format, *args):
        pass  # suppress default request logging
