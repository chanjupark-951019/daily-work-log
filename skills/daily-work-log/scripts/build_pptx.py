"""
A4 landscape PPT 생성 (python-pptx).
status + gantt 두 종류 모두 PPT 파일로 출력.
"""
from pathlib import Path

from pptx import Presentation
from pptx.util import Mm, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from lxml import etree


# Default: A4 portrait (status). Gantt은 별도 landscape 설정
STATUS_W = Mm(210)
STATUS_H = Mm(297)
GANTT_W = Mm(297)
GANTT_H = Mm(210)

# Backward-compat aliases (status 기준)
SLIDE_W = STATUS_W
SLIDE_H = STATUS_H

MARGIN_L = Mm(12)
MARGIN_R = Mm(12)
MARGIN_T = Mm(12)
MARGIN_B = Mm(10)

FONT_FACE = 'Malgun Gothic'

COLOR_BLACK = RGBColor(0, 0, 0)
COLOR_WHITE = RGBColor(255, 255, 255)
COLOR_LABEL_BG = RGBColor(242, 242, 242)
COLOR_HEADER_BG = COLOR_LABEL_BG
COLOR_HEADER_FG = COLOR_BLACK
COLOR_ALT_BG = COLOR_WHITE
COLOR_TASK_BG = COLOR_LABEL_BG
COLOR_DUE_FG = COLOR_BLACK
COLOR_ALT_TASK_BG = COLOR_LABEL_BG


def _a4_prs(orientation='portrait'):
    prs = Presentation()
    if orientation == 'landscape':
        prs.slide_width = GANTT_W
        prs.slide_height = GANTT_H
    else:
        prs.slide_width = STATUS_W
        prs.slide_height = STATUS_H
    return prs


def _set_run(run, text, *, size=Pt(9), bold=False, color=COLOR_BLACK, face=FONT_FACE):
    run.text = text
    run.font.name = face
    run.font.size = size
    run.font.bold = bold
    run.font.color.rgb = color


def _clear_paragraph(p):
    for r in list(p.runs):
        r._r.getparent().remove(r._r)


def _set_cell_text(cell, text_or_lines, *, size=Pt(9), bold=False, color=COLOR_BLACK,
                   align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, bullet=False, bg=None):
    tf = cell.text_frame
    tf.word_wrap = True
    tf.margin_left = Mm(2.2)
    tf.margin_right = Mm(1.5)
    tf.margin_top = Mm(1.0)
    tf.margin_bottom = Mm(1.0)
    cell.vertical_anchor = anchor
    if bg is not None:
        cell.fill.solid()
        cell.fill.fore_color.rgb = bg

    # Clear all paragraphs
    if isinstance(text_or_lines, str):
        lines = [text_or_lines]
    else:
        lines = list(text_or_lines) if text_or_lines else [""]

    # Remove extra paragraphs
    for _ in range(len(tf.paragraphs) - 1):
        p = tf.paragraphs[-1]
        p._p.getparent().remove(p._p)

    # First paragraph reuse
    first = tf.paragraphs[0]
    _clear_paragraph(first)
    first.alignment = align

    def _add_text(p, txt):
        run = p.add_run()
        prefix = ''
        if bullet:
            prefix = '· '
        _set_run(run, prefix + txt, size=size, bold=bold, color=color)

    _add_text(first, lines[0])
    for line in lines[1:]:
        p = tf.add_paragraph()
        p.alignment = align
        _add_text(p, line)


def _set_cell_border_color(cell, side, color=COLOR_BLACK, width_emu=12700):
    """side: 'T','B','L','R'. width_emu: default 12700 (1pt). """
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tag = f'a:ln{side}'
    ln = tcPr.find(qn(tag))
    if ln is None:
        ln = etree.SubElement(tcPr, qn(tag))
    ln.set('w', str(width_emu))
    ln.set('cap', 'flat')
    ln.set('cmpd', 'sng')
    ln.set('algn', 'ctr')
    # solidFill child
    for child in list(ln):
        ln.remove(child)
    sf = etree.SubElement(ln, qn('a:solidFill'))
    sc = etree.SubElement(sf, qn('a:srgbClr'))
    sc.set('val', '{:02X}{:02X}{:02X}'.format(color[0], color[1], color[2]))
    pd = etree.SubElement(ln, qn('a:prstDash'))
    pd.set('val', 'solid')
    etree.SubElement(ln, qn('a:round'))


def _row_height(text_lines, *, per_line_mm=4.0, min_mm=7.0):
    n = len(text_lines) if isinstance(text_lines, (list, tuple)) else 1
    return Mm(max(min_mm, n * per_line_mm + 2.0))


# ---------- Status PPT ----------

def build_status_pptx(out_path: Path, categories: list, year_short: str, month: int, week_num: int, author: str = ''):
    prs = _a4_prs('portrait')
    blank = prs.slide_layouts[6]  # Blank
    slide = prs.slides.add_slide(blank)

    content_w = STATUS_W - MARGIN_L - MARGIN_R
    col_w = content_w

    # Title
    left = MARGIN_L
    top = MARGIN_T
    tb = slide.shapes.add_textbox(left, top, Mm(150), Mm(9))
    tf = tb.text_frame
    tf.margin_left = tf.margin_right = Emu(0)
    tf.margin_top = tf.margin_bottom = Emu(0)
    p = tf.paragraphs[0]
    _clear_paragraph(p)
    r = p.add_run()
    _set_run(r, '주간 업무 현황 _ 기획', size=Pt(16), bold=True)

    # Week label (right)
    tb2 = slide.shapes.add_textbox(STATUS_W - MARGIN_R - Mm(50), top + Mm(1.5), Mm(50), Mm(7))
    tf2 = tb2.text_frame
    tf2.margin_left = tf2.margin_right = Emu(0)
    tf2.margin_top = tf2.margin_bottom = Emu(0)
    p = tf2.paragraphs[0]
    _clear_paragraph(p)
    p.alignment = PP_ALIGN.RIGHT
    r = p.add_run()
    _set_run(r, f"’{year_short}년 {month}월 {week_num}주차", size=Pt(10), bold=True)

    # Single column stacked tables
    col_data = [categories]

    body_top = MARGIN_T + Mm(10)
    body_h = STATUS_H - body_top - MARGIN_B

    for col_idx, cats in enumerate(col_data):
        col_left = MARGIN_L
        # Calculate total rows + heights
        # Each category table: 업무명 + 개요 + 실적 + [계획]
        # Heights per row: 6mm / 7mm / auto / auto
        tables_info = []
        for cat in cats:
            rows = []
            rows.append(('task', [cat['name']], _row_height(cat['name'], per_line_mm=4, min_mm=6.5)))
            rows.append(('overview', [cat['overview']], _row_height(cat['overview'], per_line_mm=3.8, min_mm=6.0)))
            rows.append(('done', cat['done'], _row_height(cat['done'], per_line_mm=4.0, min_mm=6.0)))
            if cat.get('plan'):
                rows.append(('plan', cat['plan'], _row_height(cat['plan'], per_line_mm=4.0, min_mm=6.0)))
            tables_info.append(rows)

        # Compute total height needed, and scaling if overflow
        gap_between_tables = Mm(2.0)
        total_rows_h = sum(sum(r[2] for r in tbl) for tbl in tables_info)
        total_gaps = gap_between_tables * max(0, len(tables_info) - 1)
        total_h_needed = total_rows_h + total_gaps
        scale = 1.0
        if total_h_needed > body_h:
            scale = body_h / total_h_needed

        current_top = body_top
        for tbl_idx, rows in enumerate(tables_info):
            n_rows = len(rows)
            row_heights = [int(r[2] * scale) for r in rows]
            table_h = sum(row_heights)
            tbl_shape = slide.shapes.add_table(n_rows, 2, col_left, current_top, col_w, table_h).table

            # Set column widths
            label_col_w = Mm(17)
            tbl_shape.columns[0].width = label_col_w
            tbl_shape.columns[1].width = col_w - label_col_w

            # Set row heights
            for i, h in enumerate(row_heights):
                tbl_shape.rows[i].height = h

            for r_idx, (kind, content, _h) in enumerate(rows):
                label_cell = tbl_shape.cell(r_idx, 0)
                value_cell = tbl_shape.cell(r_idx, 1)

                is_first = (r_idx == 0)
                is_last = (r_idx == n_rows - 1)

                # Label cell
                label_text = {'task': '업무명', 'overview': '개요', 'done': '실적', 'plan': '계획'}[kind]
                _set_cell_text(label_cell, label_text,
                               size=Pt(8.5), bold=True, align=PP_ALIGN.CENTER,
                               anchor=MSO_ANCHOR.MIDDLE, bg=COLOR_LABEL_BG)

                # Value cell
                if kind == 'task':
                    _set_cell_text(value_cell, content,
                                   size=Pt(10), bold=True,
                                   anchor=MSO_ANCHOR.MIDDLE)
                elif kind == 'overview':
                    _set_cell_text(value_cell, content,
                                   size=Pt(8.5),
                                   anchor=MSO_ANCHOR.MIDDLE)
                else:
                    _set_cell_text(value_cell, content,
                                   size=Pt(8.8),
                                   bullet=True,
                                   anchor=MSO_ANCHOR.TOP)

                # Borders: thick top on first row, thick bottom on last row
                if is_first:
                    _set_cell_border_color(label_cell, 'T', COLOR_BLACK, width_emu=15875)
                    _set_cell_border_color(value_cell, 'T', COLOR_BLACK, width_emu=15875)
                if is_last:
                    _set_cell_border_color(label_cell, 'B', COLOR_BLACK, width_emu=15875)
                    _set_cell_border_color(value_cell, 'B', COLOR_BLACK, width_emu=15875)

            current_top = current_top + table_h + gap_between_tables

    # Footer
    foot = slide.shapes.add_textbox(STATUS_W - MARGIN_R - Mm(30), STATUS_H - Mm(7), Mm(30), Mm(4))
    tf = foot.text_frame
    tf.margin_left = tf.margin_right = Emu(0)
    tf.margin_top = tf.margin_bottom = Emu(0)
    p = tf.paragraphs[0]
    _clear_paragraph(p)
    p.alignment = PP_ALIGN.RIGHT
    r = p.add_run()
    _set_run(r, author or '', size=Pt(8.5), color=RGBColor(85, 85, 85))

    prs.save(str(out_path))
    return out_path


# ---------- Gantt PPT ----------

def build_gantt_pptx(out_path: Path, sheets: list, author: str = ''):
    """sheets: [{'title', 'dates': {'mon','tue'..}, 'rows': [...]}, ...]"""
    prs = _a4_prs('landscape')
    blank = prs.slide_layouts[6]

    GANTT_ML = Mm(10)
    GANTT_MR = Mm(10)
    GANTT_MT = Mm(9)
    GANTT_MB = Mm(8)

    for sheet in sheets:
        slide = prs.slides.add_slide(blank)
        content_w = GANTT_W - GANTT_ML - GANTT_MR

        # Title
        tb = slide.shapes.add_textbox(GANTT_ML, GANTT_MT, Mm(200), Mm(9))
        tf = tb.text_frame
        tf.margin_left = tf.margin_right = Emu(0)
        tf.margin_top = tf.margin_bottom = Emu(0)
        p = tf.paragraphs[0]
        _clear_paragraph(p)
        r = p.add_run()
        _set_run(r, sheet['title'], size=Pt(16), bold=True)

        # Author
        tb2 = slide.shapes.add_textbox(GANTT_W - GANTT_MR - Mm(30), GANTT_MT + Mm(1.5), Mm(30), Mm(7))
        tf2 = tb2.text_frame
        tf2.margin_left = tf2.margin_right = Emu(0)
        tf2.margin_top = tf2.margin_bottom = Emu(0)
        p = tf2.paragraphs[0]
        _clear_paragraph(p)
        p.alignment = PP_ALIGN.RIGHT
        r = p.add_run()
        _set_run(r, author or '', size=Pt(10), bold=True)

        # Table
        table_top = GANTT_MT + Mm(11)
        table_h = GANTT_H - table_top - GANTT_MB
        n_rows = 1 + len(sheet['rows'])
        n_cols = 7

        tbl = slide.shapes.add_table(n_rows, n_cols, GANTT_ML, table_top, content_w, table_h).table

        # Column widths: task 10%, due 6%, day 16.8% × 5 = 84%
        total = content_w
        task_w = int(total * 0.10)
        due_w = int(total * 0.06)
        day_w = int((total - task_w - due_w) / 5)
        tbl.columns[0].width = task_w
        tbl.columns[1].width = due_w
        for i in range(2, 7):
            tbl.columns[i].width = day_w

        # Row heights
        header_h = Mm(9)
        data_h = (table_h - header_h) // len(sheet['rows'])
        tbl.rows[0].height = header_h
        for i in range(1, n_rows):
            tbl.rows[i].height = data_h

        # Header row
        headers = [
            '업무', '일정',
            f"월\n{sheet['dates']['mon']}",
            f"화\n{sheet['dates']['tue']}",
            f"수\n{sheet['dates']['wed']}",
            f"목\n{sheet['dates']['thu']}",
            f"금\n{sheet['dates']['fri']}",
        ]
        for c_idx, h in enumerate(headers):
            cell = tbl.cell(0, c_idx)
            lines = h.split('\n')
            _set_cell_text(cell, lines,
                           size=Pt(9.5), bold=True, color=COLOR_HEADER_FG,
                           align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
                           bg=COLOR_HEADER_BG)

        # Data rows
        day_keys = ['mon', 'tue', 'wed', 'thu', 'fri']
        for r_idx, row in enumerate(sheet['rows'], start=1):
            is_alt = (r_idx % 2 == 0)
            alt_bg = COLOR_ALT_BG if is_alt else RGBColor(255, 255, 255)
            task_bg = COLOR_ALT_TASK_BG if is_alt else COLOR_LABEL_BG

            # Task cell
            _set_cell_text(tbl.cell(r_idx, 0), row['task'],
                           size=Pt(9), bold=True,
                           align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
                           bg=task_bg)
            # Due cell
            _set_cell_text(tbl.cell(r_idx, 1), row.get('due', ''),
                           size=Pt(9), bold=True, color=COLOR_DUE_FG,
                           align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
                           bg=task_bg)
            # Day cells
            for c_idx, key in enumerate(day_keys, start=2):
                items = row.get(key, [])
                _set_cell_text(tbl.cell(r_idx, c_idx), items if items else "",
                               size=Pt(8.5), bullet=bool(items),
                               anchor=MSO_ANCHOR.TOP,
                               bg=alt_bg)

    prs.save(str(out_path))
    return out_path
