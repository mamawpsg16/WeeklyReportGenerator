import io
import os
import tempfile

import pandas as pd
import streamlit as st

from generate_report import (
    TEMPLATE_PATH,
    CSV_CATEGORY,
    CSV_PROJECT,
    CSV_TASK,
    CSV_PLANNED_H,
    CSV_ACTUAL_H,
    CSV_REMARKS,
    CSV_PU,
    TABLE_CAPS,
    update_title_slide,
    update_data_tables,
    generate_analysis,
    update_analysis_slide,
)
from pptx import Presentation

st.set_page_config(page_title='Weekly Report Generator', layout='centered')
st.title('Weekly Report Generator')

# ── Upload CSV ────────────────────────────────────────────────────────────────
uploaded = st.file_uploader('Upload your prod-log CSV', type='csv')

if uploaded:
    df_full = pd.read_csv(uploaded)

    months = sorted(df_full['Month'].dropna().unique().tolist())
    weeks  = sorted(df_full['Week No'].dropna().unique().tolist())

    col1, col2 = st.columns(2)
    with col1:
        selected_month = st.selectbox('Month', months)
    with col2:
        selected_week = st.selectbox('Week', weeks)

    month_display = selected_month.split('_')[-1] if '_' in selected_month else selected_month
    week_num      = ''.join(filter(str.isdigit, selected_week))

    presenter = st.text_input('Presented By', value=os.environ.get('PRESENTER', 'Kevin Mensah'))
    report_date = st.text_input('Reporting Date', value=pd.Timestamp.now().strftime('%m/%d/%Y'))

    df_filtered = df_full[
        (df_full['Month'] == selected_month) &
        (df_full['Week No'] == selected_week)
    ].copy()

    df_filtered[CSV_PLANNED_H] = df_filtered[CSV_PLANNED_H].fillna(0)
    df_filtered[CSV_ACTUAL_H]  = df_filtered[CSV_ACTUAL_H].fillna(0)
    df_filtered[CSV_REMARKS]   = df_filtered[CSV_REMARKS].fillna('')
    if CSV_PU and CSV_PU in df_filtered.columns:
        df_filtered[CSV_PU] = df_filtered[CSV_PU].fillna('Planned')

    st.divider()
    st.subheader(f'Preview — {month_display} {selected_week}')

    if df_filtered.empty:
        st.warning('No rows found for the selected month and week.')
    else:
        preview_cols = [CSV_CATEGORY, CSV_PROJECT, CSV_TASK, CSV_PLANNED_H, CSV_ACTUAL_H, CSV_REMARKS]
        existing_cols = [c for c in preview_cols if c in df_filtered.columns]
        st.dataframe(df_filtered[existing_cols], use_container_width=True)

        total_p = df_filtered[CSV_PLANNED_H].sum()
        total_a = df_filtered[CSV_ACTUAL_H].sum()
        c1, c2, c3 = st.columns(3)
        c1.metric('Planned Hours', f'{total_p:.2f}')
        c2.metric('Actual Hours',  f'{total_a:.2f}')
        c3.metric('Variance',      f'{total_a - total_p:.2f}')

        st.divider()
        if st.button('Generate Report', type='primary', use_container_width=True):
            with st.spinner('Building your PowerPoint report...'):
                import generate_report as gr
                gr.REPORT_MONTH  = selected_month
                gr.REPORT_WEEK   = selected_week
                gr.MONTH_DISPLAY = month_display
                gr.WEEK_NUM      = week_num
                gr.PRESENTER     = presenter
                gr.REPORT_DATE   = report_date

                prs = Presentation(TEMPLATE_PATH)
                update_title_slide(prs)
                update_data_tables(prs, df_filtered)

                with st.status('Generating AI analysis...'):
                    paragraphs = generate_analysis(df_filtered)
                update_analysis_slide(prs, paragraphs)

                buf = io.BytesIO()
                prs.save(buf)
                buf.seek(0)

            st.success('Report generated!')
            filename = f'weekly-report-{selected_month}-{selected_week.replace(" ", "_")}.pptx'
            st.download_button(
                label='Download PPTX',
                data=buf,
                file_name=filename,
                mime='application/vnd.openxmlformats-officedocument.presentationml.presentation',
                use_container_width=True,
            )
