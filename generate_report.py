import os
import copy
from datetime import datetime
import pandas as pd
from pptx import Presentation
from pptx.oxml.ns import qn
from groq import Groq

# ── Config ────────────────────────────────────────────────────────────────────
CSV_PATH      = os.environ.get('CSV_PATH', '/data/prod-log.csv')
TEMPLATE_PATH = os.environ.get('TEMPLATE_PATH', os.path.join(os.path.dirname(__file__), 'weekly-report-template.pptx'))
OUTPUT_DIR    = os.environ.get('OUTPUT_DIR', '/output')
GROQ_KEY      = os.environ.get('GROQ_API_KEY', '')

REPORT_MONTH  = os.environ.get('REPORT_MONTH', '05_May')
REPORT_WEEK   = os.environ.get('REPORT_WEEK', 'Week 3')
MONTH_DISPLAY = os.environ.get('MONTH_DISPLAY', 'May')
WEEK_NUM      = os.environ.get('WEEK_NUM', '3')
PRESENTER     = os.environ.get('PRESENTER', 'Kevin Mensah')
REPORT_DATE   = os.environ.get('REPORT_DATE', datetime.now().strftime('%m/%d/%Y'))

# CSV column names — adjust if your headers differ
CSV_CATEGORY   = 'Category'
CSV_PROJECT    = 'Project'
CSV_TASK       = 'Activity/Task'
CSV_PLANNED_H  = 'Planned Hours'
CSV_ACTUAL_H   = 'Actual Hours Consumed'
CSV_REMARKS    = 'Remarks (Optional)'
CSV_PU         = 'Planned/Unplanned'   # set to None if column doesn't exist

# Slide indexes (0-based)
SLIDE_TITLE    = 0
SLIDE_TABLE_1  = 1   # 7-col table
SLIDE_TABLE_2  = 2   # 8-col table (extra Time col)
SLIDE_TABLE_3  = 3   # last table + total row
SLIDE_ANALYSIS = 4   # AI analysis text

# Rows per table slide (excluding header and total row)
TABLE_CAPS = [7, 7, 3]


# ── Helpers ───────────────────────────────────────────────────────────────────

def set_cell_text(cell, text):
    """Write text into a table cell while preserving existing run formatting."""
    tf = cell.text_frame
    # Remove extra paragraphs
    while len(tf.paragraphs) > 1:
        p = tf.paragraphs[-1]._p
        p.getparent().remove(p)

    para = tf.paragraphs[0]
    runs = para.runs

    if runs:
        runs[0].text = str(text)
        runs[0].font.name = 'Montserrat'
        runs[0].font.size = 100800  # 8pt in EMUs (1pt = 12600 EMUs)
        for r in runs[1:]:
            r._r.getparent().remove(r._r)
    else:
        from pptx.oxml import parse_xml
        text_escaped = str(text).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        r_xml = (
            '<a:r xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
            f'<a:t>{text_escaped}</a:t></a:r>'
        )
        para._p.append(parse_xml(r_xml))
        para.runs[0].font.name = 'Montserrat'
        para.runs[0].font.size = 100800  # 8pt in EMUs


def add_table_row(table, template_row_idx=1):
    """Append a new row to a table by cloning an existing data row."""
    tbl = table._tbl
    rows = tbl.findall(qn('a:tr'))
    new_tr = copy.deepcopy(rows[template_row_idx])
    for tc in new_tr.findall(qn('a:tc')):
        txBody = tc.find(qn('a:txBody'))
        if txBody is not None:
            for para in txBody.findall(qn('a:p')):
                for r in para.findall(qn('a:r')):
                    t = r.find(qn('a:t'))
                    if t is not None:
                        t.text = ''
    tbl.append(new_tr)


def fill_table(table, data_rows, has_time_col=False):
    """Fill a table's data rows (skipping header row 0)."""
    tbl = table._tbl
    tr_list = tbl.findall(qn('a:tr'))
    data_trs = tr_list[1:]  # skip header

    # Add rows if we need more capacity
    while len(data_trs) < len(data_rows):
        add_table_row(table, template_row_idx=1)
        data_trs = tbl.findall(qn('a:tr'))[1:]

    for i, row_data in enumerate(data_rows):
        tr = data_trs[i]
        tcs = tr.findall(qn('a:tc'))
        for j, val in enumerate(row_data):
            if j < len(tcs):
                set_cell_text(table.cell(i + 1, j), str(val) if val else '')



def remove_empty_rows(table):
    """Delete rows where all cells are empty."""
    tbl = table._tbl
    tr_list = list(tbl.findall(qn('a:tr')))[1:]  # skip header

    for tr in tr_list:
        tcs = tr.findall(qn('a:tc'))
        is_empty = True
        for tc in tcs:
            # Check if cell has any text
            txBody = tc.find(qn('a:txBody'))
            if txBody is not None:
                for para in txBody.findall(qn('a:p')):
                    for r in para.findall(qn('a:r')):
                        t = r.find(qn('a:t'))
                        if t is not None and t.text and t.text.strip():
                            is_empty = False
                            break
            if not is_empty:
                break

        if is_empty:
            tr.getparent().remove(tr)


def get_table_from_slide(slide):
    for shape in slide.shapes:
        if shape.shape_type == 19:
            return shape.table
    return None


def set_text_box(slide, shape_name, paragraphs):
    """Replace text in a text box shape by name, one string per paragraph."""
    for shape in slide.shapes:
        if shape.name == shape_name and shape.has_text_frame:
            tf = shape.text_frame
            # Clear all paragraph runs
            for i, para in enumerate(tf.paragraphs):
                for r in para.runs:
                    r._r.getparent().remove(r._r)

            for i, text in enumerate(paragraphs):
                if i < len(tf.paragraphs):
                    para = tf.paragraphs[i]
                else:
                    from pptx.oxml import parse_xml
                    p_xml = '<a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>'
                    tf._txBody.append(parse_xml(p_xml))
                    para = tf.paragraphs[i]

                from pptx.oxml import parse_xml
                text_escaped = text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
                r_xml = (
                    '<a:r xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
                    f'<a:t>{text_escaped}</a:t></a:r>'
                )
                para._p.append(parse_xml(r_xml))
            return


# ── Steps ─────────────────────────────────────────────────────────────────────

def load_data():
    df = pd.read_csv(CSV_PATH)
    df = df[(df['Month'] == REPORT_MONTH) & (df['Week No'] == REPORT_WEEK)]
    df[CSV_PLANNED_H] = df[CSV_PLANNED_H].fillna(0)
    df[CSV_ACTUAL_H]  = df[CSV_ACTUAL_H].fillna(0)
    df[CSV_REMARKS]   = df[CSV_REMARKS].fillna('')
    if CSV_PU and CSV_PU in df.columns:
        df[CSV_PU] = df[CSV_PU].fillna('Planned')
    return df


def build_row(row):
    """Convert a DataFrame row into an ordered list matching template columns."""
    pu = row[CSV_PU] if (CSV_PU and CSV_PU in row.index) else 'Planned'
    return [
        row[CSV_CATEGORY],
        row[CSV_PROJECT],
        row[CSV_TASK],
        pu,
        f"{row[CSV_PLANNED_H]:.2f}",
        f"{row[CSV_ACTUAL_H]:.2f}",
        row[CSV_REMARKS],
    ]


def update_title_slide(prs):
    slide = prs.slides[SLIDE_TITLE]
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        tf = shape.text_frame
        for para in tf.paragraphs:
            text = para.text
            if text.startswith('Month:'):
                if para.runs:
                    para.runs[0].text = f'Month: {MONTH_DISPLAY}'
            elif text.startswith('Week No:'):
                if para.runs:
                    para.runs[0].text = f'Week No: {WEEK_NUM}'
            elif text.startswith('Reporting Date:'):
                if para.runs:
                    para.runs[0].text = f'Reporting Date: {REPORT_DATE}'
            elif text.startswith('Presented By:'):
                if para.runs:
                    para.runs[0].text = f'Presented By: {PRESENTER}'


def update_data_tables(prs, df):
    all_rows = [build_row(row) for _, row in df.iterrows()]
    total_planned = df[CSV_PLANNED_H].sum()
    total_actual  = df[CSV_ACTUAL_H].sum()

    slide_indexes = [SLIDE_TABLE_1, SLIDE_TABLE_2, SLIDE_TABLE_3]
    offset = 0

    for slide_idx, cap in zip(slide_indexes, TABLE_CAPS):
        slide  = prs.slides[slide_idx]
        table  = get_table_from_slide(slide)
        if table is None:
            continue

        chunk = all_rows[offset: offset + cap]
        is_last = (slide_idx == SLIDE_TABLE_3)

        if is_last:
            fill_table(table, chunk)
            # Update total row (last row in the table)
            last_row_idx = len(table.rows) - 1
            set_cell_text(table.cell(last_row_idx, 3), 'Total:')
            set_cell_text(table.cell(last_row_idx, 4), f'{total_planned:.2f}')
            set_cell_text(table.cell(last_row_idx, 5), f'{total_actual:.2f}')
            remove_empty_rows(table)
        else:
            fill_table(table, chunk)
            remove_empty_rows(table)

        offset += cap
        if offset >= len(all_rows):
            break


def generate_analysis(df):
    if not GROQ_KEY:
        return ['[Add GROQ_API_KEY to .env to enable AI analysis]']

    client = Groq(api_key=GROQ_KEY)

    summary_lines = []
    for _, row in df.iterrows():
        summary_lines.append(
            f"- {row[CSV_CATEGORY]} | {row[CSV_PROJECT]} | {row[CSV_TASK]} | "
            f"Planned: {row[CSV_PLANNED_H]:.2f}h | Actual: {row[CSV_ACTUAL_H]:.2f}h | "
            f"Remarks: {row[CSV_REMARKS]}"
        )
    data_text = '\n'.join(summary_lines)
    total_p = df[CSV_PLANNED_H].sum()
    total_a = df[CSV_ACTUAL_H].sum()

    prompt = f"""You are writing a professional weekly productivity report analysis.

Week: {MONTH_DISPLAY} {REPORT_WEEK}
Total Planned: {total_p:.2f}h | Total Actual: {total_a:.2f}h | Variance: {total_a - total_p:.2f}h

Task breakdown:
{data_text}

Write exactly 3 short sentences — one per line:
1. Total planned vs actual hours and what the variance means
2. Which category had the most work and highlight one over/under-run if any
3. One short takeaway for the week

Plain language, no jargon, no filler. Like explaining to your manager in 30 seconds."""

    response = client.chat.completions.create(
        model='llama-3.3-70b-versatile',
        messages=[{'role': 'user', 'content': prompt}],
    )
    text = response.choices[0].message.content.strip()
    paragraphs = [p.strip() for p in text.split('\n\n') if p.strip()]
    return paragraphs


def update_analysis_slide(prs, paragraphs):
    slide = prs.slides[SLIDE_ANALYSIS]
    text_boxes = [s for s in slide.shapes if s.shape_type == 17 and s.has_text_frame]
    if not text_boxes:
        print('Warning: no text box found on analysis slide')
        return
    # Use the largest text box (the analysis one)
    target = max(text_boxes, key=lambda s: s.width * s.height)
    tf = target.text_frame

    # Clear existing content
    for para in tf.paragraphs:
        for r in para.runs:
            r._r.getparent().remove(r._r)
    while len(tf.paragraphs) > 1:
        p = tf.paragraphs[-1]._p
        p.getparent().remove(p)

    from pptx.oxml import parse_xml
    for i, text in enumerate(paragraphs):
        if i == 0:
            para = tf.paragraphs[0]
        else:
            p_xml = '<a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>'
            tf._txBody.append(parse_xml(p_xml))
            para = tf.paragraphs[i]
        text_escaped = text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        r_xml = (
            '<a:r xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
            f'<a:t>{text_escaped}</a:t></a:r>'
        )
        para._p.append(parse_xml(r_xml))


def main():
    print(f'Loading data: {REPORT_MONTH} / {REPORT_WEEK}')
    df = load_data()
    print(f'  {len(df)} rows found')
    if df.empty:
        print('No data found — check REPORT_MONTH and REPORT_WEEK values match your CSV.')
        return

    print('Opening template...')
    prs = Presentation(TEMPLATE_PATH)

    print('Updating title slide...')
    update_title_slide(prs)

    print('Populating data tables...')
    update_data_tables(prs, df)

    print('Generating AI analysis...')
    paragraphs = generate_analysis(df)
    update_analysis_slide(prs, paragraphs)

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    out_file = os.path.join(OUTPUT_DIR, f'weekly-report-{REPORT_MONTH}-{REPORT_WEEK.replace(" ", "_")}.pptx')
    prs.save(out_file)
    print(f'Saved: {out_file}')


if __name__ == '__main__':
    main()
