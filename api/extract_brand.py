import json
import io
import base64
import zipfile
import re
import posixpath
from http.server import BaseHTTPRequestHandler


# ─── HELPERS ─────────────────────────────────────────────────────────────────

def emu_to_inches(emu):
    """Convert English Metric Units to inches."""
    try:
        return round(int(emu) / 914400, 2)
    except:
        return 0


def emu_to_pt(emu):
    """Convert EMU to points."""
    try:
        return round(int(emu) / 12700, 1)
    except:
        return 0


def half_pt_to_pt(half_pt):
    """Convert half-points to points."""
    try:
        return round(int(half_pt) / 2, 1)
    except:
        return 0


def clean_hex(val):
    """Normalize hex color value."""
    if not val:
        return None
    h = str(val).replace('#', '').strip().upper()
    if len(h) == 6 and all(c in '0123456789ABCDEF' for c in h):
        return '#' + h
    return None


def read_relationships(z, rel_path):
    """Read OpenXML relationship file and return rid -> normalized target path."""
    rels = {}
    try:
        rel_xml = z.read(rel_path).decode('utf-8', errors='ignore')
    except Exception:
        return rels

    base_dir = posixpath.dirname(rel_path).replace('/_rels', '')
    for rid, target in re.findall(r'Id="([^"]+)"[^>]*Target="([^"]+)"', rel_xml):
        rels[rid] = posixpath.normpath(posixpath.join(base_dir, target))
    return rels


def extract_shape_style(block):
    """Extract best-effort text and fill style from a placeholder shape block."""
    style = {}

    sz_match = re.search(r'<a:rPr[^>]*sz="(\d+)"', block)
    if sz_match:
        style['font_size_pt'] = half_pt_to_pt(sz_match.group(1))

    latin_match = re.search(r'<a:latin[^>]*typeface="([^"]+)"', block)
    if latin_match and not latin_match.group(1).startswith('+'):
        style['font_family'] = latin_match.group(1)

    color_match = re.search(r'<a:(?:srgbClr|schemeClr) val="([0-9A-Fa-f]{6}|[A-Za-z0-9]+)"', block)
    if color_match:
        style['font_color'] = clean_hex(color_match.group(1)) or color_match.group(1)

    bold_match = re.search(r'<a:rPr[^>]*b="1"', block)
    if bold_match:
        style['font_weight'] = 'bold'

    align_match = re.search(r'<a:pPr[^>]*algn="([^"]+)"', block)
    if align_match:
        style['align'] = align_match.group(1)

    fill_match = re.search(r'<a:solidFill>\s*<a:(?:srgbClr|schemeClr) val="([0-9A-Fa-f]{6}|[A-Za-z0-9]+)"', block)
    if fill_match:
        style['fill_color'] = clean_hex(fill_match.group(1)) or fill_match.group(1)

    return style


def extract_placeholders_from_xml(xml_text):
    """Extract placeholder geometry + style from layout/master XML."""
    placeholders = []
    ph_blocks = re.findall(r'<p:sp>(.*?)</p:sp>', xml_text, re.DOTALL)

    for block in ph_blocks:
        ph_match = re.search(r'<p:ph([^>]*)/>|<p:ph([^>]*)>', block)
        if not ph_match:
            continue

        ph_attrs = ph_match.group(1) or ph_match.group(2) or ''
        type_match = re.search(r'type="([^"]+)"', ph_attrs)
        idx_match = re.search(r'idx="(\d+)"', ph_attrs)
        ph_type = type_match.group(1) if type_match else 'body'
        ph_idx = int(idx_match.group(1)) if idx_match else 0

        off_match = re.search(r'<a:off x="(-?\d+)" y="(-?\d+)"', block)
        ext_match = re.search(r'<a:ext cx="(\d+)" cy="(\d+)"', block)
        if not (off_match and ext_match):
            continue

        ph = {
            'type': ph_type,
            'idx': ph_idx,
            'x_in': emu_to_inches(off_match.group(1)),
            'y_in': emu_to_inches(off_match.group(2)),
            'w_in': emu_to_inches(ext_match.group(1)),
            'h_in': emu_to_inches(ext_match.group(2))
        }

        style = extract_shape_style(block)
        if style:
            ph['style'] = style
            if style.get('font_size_pt'):
                ph['font_size_pt'] = style['font_size_pt']

        text_match = re.findall(r'<a:t>([^<]+)</a:t>', block)
        if text_match:
            ph['default_text'] = ' '.join(t.strip() for t in text_match if t.strip())

        placeholders.append(ph)

    return placeholders


def cluster_axis(values, tolerance=0.45):
    groups = []
    for v in sorted(values):
        if not groups or abs(v - groups[-1][-1]) > tolerance:
            groups.append([v])
        else:
            groups[-1].append(v)
    return [round(sum(g) / len(g), 2) for g in groups]


def analyze_placeholder_grid(placeholders):
    """Summarize row/column grouping so Agent 5 can reason about grids."""
    content = [p for p in placeholders if p.get('type') not in ('title', 'ctrTitle', 'subTitle', 'dt', 'ftr', 'sldNum')]
    if not content:
        return {'rows': 0, 'columns': 0, 'content_blocks': 0}

    x_centers = [round(p['x_in'] + p['w_in'] / 2, 2) for p in content]
    y_centers = [round(p['y_in'] + p['h_in'] / 2, 2) for p in content]
    cols = cluster_axis(x_centers)
    rows = cluster_axis(y_centers)

    return {
        'rows': len(rows),
        'columns': len(cols),
        'content_blocks': len(content),
        'grid_detected': len(rows) >= 2 and len(cols) >= 2,
        'header_body_pairs_likely': len(content) >= 4 and len(rows) >= 2
    }


def summarize_text_styles(placeholders):
    """Summarize dominant text styles by placeholder role."""
    summary = {
        'title': {},
        'body': {},
        'other': {}
    }

    for ph in placeholders:
        style = ph.get('style') or {}
        role = 'other'
        if ph.get('type') in ('title', 'ctrTitle', 'subTitle'):
            role = 'title'
        elif ph.get('type') == 'body':
            role = 'body'

        target = summary[role]
        if not target.get('font_family') and style.get('font_family'):
            target['font_family'] = style.get('font_family')
        if not target.get('font_size_pt') and style.get('font_size_pt'):
            target['font_size_pt'] = style.get('font_size_pt')
        if not target.get('font_color') and style.get('font_color'):
            target['font_color'] = style.get('font_color')
        if not target.get('font_weight') and style.get('font_weight'):
            target['font_weight'] = style.get('font_weight')
        if not target.get('align') and style.get('align'):
            target['align'] = style.get('align')

    return summary


def summarize_regions(placeholders):
    """Summarize major regions on a master/layout."""
    title = next((p for p in placeholders if p.get('type') in ('title', 'ctrTitle')), None)
    subtitle = next((p for p in placeholders if p.get('type') == 'subTitle'), None)
    bodies = [p for p in placeholders if p.get('type') == 'body']

    return {
        'title_region': title,
        'subtitle_region': subtitle,
        'body_regions': bodies[:12],
        'body_region_count': len(bodies)
    }


def extract_media_refs(z, rel_path):
    """Return master/layout media targets."""
    refs = []
    rels = read_relationships(z, rel_path)
    for rid, target in rels.items():
        if target.startswith('ppt/media/'):
            refs.append({
                'rel_id': rid,
                'target': target,
                'name': posixpath.basename(target)
            })
    return refs


def extract_master_info(master_xml, filename, media_refs=None):
    """Extract master-level layout guidance from a slide master."""
    name_match = re.search(r'<p:cSld[^>]*name="([^"]+)"', master_xml)
    master_name = name_match.group(1) if name_match else posixpath.basename(filename)

    placeholders = extract_placeholders_from_xml(master_xml)
    bg_match = re.search(r'<a:solidFill>\s*<a:srgbClr val="([0-9A-Fa-f]{6})"', master_xml)
    background_color = clean_hex(bg_match.group(1)) if bg_match else None

    title_ph = next((p for p in placeholders if p.get('type') in ('title', 'ctrTitle')), None)
    body_ph = next((p for p in placeholders if p.get('type') == 'body'), None)

    return {
        'name': master_name,
        'path': filename,
        'background_color': background_color,
        'placeholder_count': len(placeholders),
        'grid_summary': analyze_placeholder_grid(placeholders),
        'text_style_summary': summarize_text_styles(placeholders),
        'regions': summarize_regions(placeholders),
        'media_refs': media_refs or [],
        'title_placeholder': title_ph,
        'body_placeholder': body_ph,
        'placeholders': placeholders,
        'layout_paths': [],
        'layout_names': []
    }


# ─── EXTRACT FROM ZIP ─────────────────────────────────────────────────────────

def extract_brand_from_pptx(pptx_bytes):
    """
    Main extractor. Reads PPTX as a zip file and pulls:
    - Slide dimensions
    - Theme colors
    - Theme fonts
    - All slide layout names and their placeholder structures
    """
    result = {
        'slide_width_inches':  0,
        'slide_height_inches': 0,
        'slide_width_pt':      0,
        'slide_height_pt':     0,
        'color_scheme_name':   '',
        'font_scheme_name':    '',
        'primary_colors':      [],
        'secondary_colors':    [],
        'background_colors':   [],
        'text_colors':         [],
        'accent_colors':       [],
        'chart_colors':        [],
        'all_colors':          {},
        'title_font':          {},
        'body_font':           {},
        'slide_layouts':       [],
        'slide_masters':       [],
        'raw_fonts':           {},
        'errors':              []
    }

    try:
        with zipfile.ZipFile(io.BytesIO(pptx_bytes)) as z:
            files = z.namelist()

            # ── 1. Slide dimensions from presentation.xml ─────────────────
            if 'ppt/presentation.xml' in files:
                prs_xml = z.read('ppt/presentation.xml').decode('utf-8', errors='ignore')
                w_match = re.search(r'cx="(\d+)"', prs_xml)
                h_match = re.search(r'cy="(\d+)"', prs_xml)
                # More specific search for sldSz
                sz_match = re.search(r'<p:sldSz[^>]*cx="(\d+)"[^>]*cy="(\d+)"', prs_xml)
                if sz_match:
                    cx = int(sz_match.group(1))
                    cy = int(sz_match.group(2))
                    result['slide_width_inches']  = emu_to_inches(cx)
                    result['slide_height_inches'] = emu_to_inches(cy)
                    result['slide_width_pt']      = round(emu_to_inches(cx) * 72, 1)
                    result['slide_height_pt']     = round(emu_to_inches(cy) * 72, 1)

            # ── 2. Theme colors and fonts ─────────────────────────────────
            theme_files = [f for f in files if re.match(r'ppt/theme/theme\d+\.xml', f)]

            # Use theme1.xml as primary (it belongs to the main slide master)
            primary_theme = 'ppt/theme/theme1.xml'
            if primary_theme not in theme_files and theme_files:
                primary_theme = theme_files[0]

            if primary_theme in files:
                theme_xml = z.read(primary_theme).decode('utf-8', errors='ignore')
                extract_theme_colors(theme_xml, result)
                extract_theme_fonts(theme_xml, result)

            # ── 3. Slide masters + layouts ───────────────────────────────
            master_files = sorted(
                [f for f in files if re.match(r'ppt/slideMasters/slideMaster\d+\.xml$', f)],
                key=lambda x: int(re.search(r'(\d+)', x).group(1))
            )
            masters_by_path = {}
            layout_to_master = {}

            for mf in master_files:
                master_xml = z.read(mf).decode('utf-8', errors='ignore')
                media_refs = extract_media_refs(
                    z,
                    posixpath.join(posixpath.dirname(mf), '_rels', posixpath.basename(mf) + '.rels')
                )
                master_info = extract_master_info(master_xml, mf, media_refs)
                result['slide_masters'].append(master_info)
                masters_by_path[mf] = master_info

                rel_path = posixpath.join(posixpath.dirname(mf), '_rels', posixpath.basename(mf) + '.rels')
                rels = read_relationships(z, rel_path)
                for target in rels.values():
                    if target.startswith('ppt/slideLayouts/'):
                        layout_to_master[target] = mf

            layout_files = sorted(
                [f for f in files if re.match(r'ppt/slideLayouts/slideLayout\d+\.xml$', f)],
                key=lambda x: int(re.search(r'(\d+)', x).group(1))
            )

            for lf in layout_files:
                layout_xml = z.read(lf).decode('utf-8', errors='ignore')
                master_path = layout_to_master.get(lf)
                layout_info = extract_layout_info(layout_xml, lf, masters_by_path.get(master_path))
                if layout_info:
                    result['slide_layouts'].append(layout_info)
                    if master_path in masters_by_path:
                        masters_by_path[master_path]['layout_paths'].append(lf)
                        masters_by_path[master_path]['layout_names'].append(layout_info.get('name', lf))

    except Exception as e:
        result['errors'].append(str(e))

    return result


def extract_theme_colors(theme_xml, result):
    """Extract color scheme from theme XML."""

    # Get color scheme name
    name_match = re.search(r'<a:clrScheme[^>]*name="([^"]+)"', theme_xml)
    if name_match:
        result['color_scheme_name'] = name_match.group(1)

    # Color role mapping
    color_roles = {
        'dk1':     'dark1',
        'lt1':     'light1',
        'dk2':     'dark2',
        'lt2':     'light2',
        'accent1': 'accent1',
        'accent2': 'accent2',
        'accent3': 'accent3',
        'accent4': 'accent4',
        'accent5': 'accent5',
        'accent6': 'accent6',
        'hlink':   'hyperlink',
        'folHlink':'followed_hyperlink'
    }

    all_colors = {}

    for tag, role in color_roles.items():
        # Match <a:TAG> block and extract srgbClr or sysClr lastClr
        pattern = r'<a:' + tag + r'>.*?</a:' + tag + r'>'
        block_match = re.search(pattern, theme_xml, re.DOTALL)
        if block_match:
            block = block_match.group(0)
            # Direct srgbClr
            rgb_match = re.search(r'<a:srgbClr val="([0-9A-Fa-f]{6})"', block)
            if rgb_match:
                color = clean_hex(rgb_match.group(1))
                if color:
                    all_colors[role] = color
            # sysClr with lastClr fallback
            sys_match = re.search(r'<a:sysClr[^>]*lastClr="([0-9A-Fa-f]{6})"', block)
            if sys_match and role not in all_colors:
                color = clean_hex(sys_match.group(1))
                if color:
                    all_colors[role] = color

    result['all_colors'] = all_colors

    # Categorise colors
    if 'accent1' in all_colors:
        result['primary_colors'] = [all_colors['accent1']]
    if 'accent2' in all_colors:
        result['secondary_colors'] = [all_colors['accent2']]

    accent_keys = ['accent3', 'accent4', 'accent5', 'accent6']
    result['accent_colors'] = [all_colors[k] for k in accent_keys if k in all_colors]

    # Chart colors = accent1 through accent6 in order
    chart_keys = ['accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6']
    result['chart_colors'] = [all_colors[k] for k in chart_keys if k in all_colors]

    bg_keys = ['light1', 'light2']
    result['background_colors'] = [all_colors[k] for k in bg_keys if k in all_colors]

    text_keys = ['dark1', 'dark2']
    result['text_colors'] = [all_colors[k] for k in text_keys if k in all_colors]


def extract_theme_fonts(theme_xml, result):
    """Extract font scheme from theme XML."""

    name_match = re.search(r'<a:fontScheme[^>]*name="([^"]+)"', theme_xml)
    if name_match:
        result['font_scheme_name'] = name_match.group(1)

    raw_fonts = {}

    # Major font (headings/titles)
    major_match = re.search(r'<a:majorFont>(.*?)</a:majorFont>', theme_xml, re.DOTALL)
    if major_match:
        latin = re.search(r'<a:latin typeface="([^"]+)"', major_match.group(1))
        if latin and not latin.group(1).startswith('+'):
            raw_fonts['major'] = latin.group(1)

    # Minor font (body text)
    minor_match = re.search(r'<a:minorFont>(.*?)</a:minorFont>', theme_xml, re.DOTALL)
    if minor_match:
        latin = re.search(r'<a:latin typeface="([^"]+)"', minor_match.group(1))
        if latin and not latin.group(1).startswith('+'):
            raw_fonts['minor'] = latin.group(1)

    result['raw_fonts'] = raw_fonts

    major_font = raw_fonts.get('major', 'Arial')
    minor_font = raw_fonts.get('minor', 'Arial')

    result['title_font'] = {
        'family': major_font,
        'size':   '28pt',
        'weight': 'bold',
        'color':  result['primary_colors'][0] if result['primary_colors'] else '#000000'
    }
    result['body_font'] = {
        'family': minor_font,
        'size':   '14pt',
        'weight': 'regular',
        'color':  result['text_colors'][0] if result['text_colors'] else '#000000'
    }
    result['caption_font'] = {
        'family': minor_font,
        'size':   '9pt',
        'weight': 'regular',
        'color':  '#666666'
    }


def extract_layout_info(layout_xml, filename, master_info=None):
    """
    Extract name and placeholder structure from a slide layout XML.
    Returns a dict describing the layout.
    """

    # Layout name
    name_match = re.search(r'<p:cSld[^>]*name="([^"]+)"', layout_xml)
    layout_name = name_match.group(1) if name_match else filename

    # Layout type
    type_match = re.search(r'<p:sldLayout[^>]*type="([^"]+)"', layout_xml)
    layout_type = type_match.group(1) if type_match else 'unknown'

    placeholders = extract_placeholders_from_xml(layout_xml)

    # Determine column structure from placeholder positions
    structure = determine_structure(placeholders, layout_name)
    grid_summary = analyze_placeholder_grid(placeholders)
    master_name = (master_info or {}).get('name', '')
    matched_title = next((p for p in placeholders if p.get('type') in ('title', 'ctrTitle')), None)
    matched_body = next((p for p in placeholders if p.get('type') == 'body'), None)

    return {
        'name':         layout_name,
        'type':         layout_type,
        'placeholders': placeholders,
        'structure':    structure,
        'ph_count':     len(placeholders),
        'grid_summary': grid_summary,
        'master_name':  master_name,
        'master_summary': {
            'background_color': (master_info or {}).get('background_color'),
            'title_placeholder': (master_info or {}).get('title_placeholder'),
            'body_placeholder': (master_info or {}).get('body_placeholder')
        } if master_info else None,
        'title_placeholder': matched_title,
        'body_placeholder': matched_body
    }


def determine_structure(placeholders, name):
    """
    Infer the visual structure of a layout from placeholder positions.
    Returns a human-readable description.
    """
    name_lower = name.lower()

    # Name-based detection first
    if 'title slide' in name_lower or 'title' == name_lower:
        return 'Full-page title with subtitle — used for opening slide'
    if 'section' in name_lower or 'divider' in name_lower:
        return 'Section divider — full-page section break with title'
    if 'appendix' in name_lower:
        return 'Appendix divider — marks appendix sections'
    if 'blank' in name_lower:
        return 'Blank — no placeholders, fully custom'
    if 'contents' in name_lower or 'agenda' in name_lower:
        return 'Contents / agenda layout'

    # Count content placeholders (not title)
    content_phs = [p for p in placeholders if p.get('type') not in ('title', 'ctrTitle', 'subTitle', 'dt', 'ftr', 'sldNum')]

    if not content_phs:
        return 'Title only — large title, no content area'

    if len(content_phs) == 1:
        return 'Single content area — title + one main content block'

    # Check horizontal layout (columns) vs vertical
    if len(content_phs) >= 2:
        # Sort by x position
        sorted_by_x = sorted(content_phs, key=lambda p: p.get('x_in', 0))
        x_positions = [p.get('x_in', 0) for p in sorted_by_x]

        # Check if they are side by side (similar y, different x)
        y_positions = [p.get('y_in', 0) for p in content_phs]
        y_range = max(y_positions) - min(y_positions) if y_positions else 0
        x_range = max(x_positions) - min(x_positions) if x_positions else 0

        if x_range > 2 and y_range < 1.5:
            cols = len(content_phs)
            return f'{cols}-column layout — title + {cols} side-by-side content areas'
        elif y_range > 1.5 and x_range < 2:
            return f'Stacked layout — title + {len(content_phs)} vertical content sections'
        else:
            return f'Grid layout — title + {len(content_phs)} content areas in mixed arrangement'

    return f'Custom layout with {len(placeholders)} placeholders'


# ─── VERCEL HANDLER ──────────────────────────────────────────────────────────

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

            b64_data  = data.get('pptxBase64', '')
            file_type = data.get('fileType', 'pptx')

            if not b64_data:
                self._json(400, {'error': 'pptxBase64 is required'})
                return

            pptx_bytes = base64.b64decode(b64_data)

            if file_type not in ('pptx', 'ppt'):
                # For non-PPTX files return a minimal structure
                self._json(200, {
                    'success': True,
                    'source':  file_type,
                    'message': 'Non-PPTX file — Claude vision will extract brand rules',
                    'extracted': None
                })
                return

            extracted = extract_brand_from_pptx(pptx_bytes)

            self._json(200, {
                'success':   True,
                'source':    'pptx_extraction',
                'extracted': extracted
            })

        except Exception as e:
            import traceback
            self._json(500, {
                'error':   str(e),
                'details': traceback.format_exc()
            })

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
