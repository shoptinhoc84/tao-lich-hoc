import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from io import BytesIO
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

st.set_page_config(page_title="Tạo Lịch Học Tự Động", layout="wide")

# Danh sách khung giờ cố định theo file mẫu
TIME_SLOTS = [
    "07:30\n-\n09:30",
    "09:30\n-\n11:30",
    "13:00\n-\n15:00",
    "15:00\n-\n17:00"
]

TIME_SLOTS_FLAT = [
    "07:30 - 09:30",
    "09:30 - 11:30",
    "13:00 - 15:00",
    "15:00 - 17:00"
]

DAYS_HEADER = ["THỨ HAI", "THỨ BA", "THỨ TƯ", "THỨ NĂM", "THỨ SÁU", "THỨ BẢY", "CHỦ NHẬT"]
STUDENTS_PER_SLOT = 4  # Mỗi ca 4 hàng (4 học viên)


def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)


def set_cell_margins(cell, top=30, bottom=30, left=60, right=60):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)


def format_cell_text(cell, text, bold=False, font_size=10, color_rgb=(0, 0, 0), align=WD_ALIGN_PARAGRAPH.CENTER):
    cell.text = ""
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    p = cell.paragraphs[0]
    p.alignment = align
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after = Pt(1)
    run = p.add_run(str(text))
    run.font.name = "Times New Roman"
    run.font.size = Pt(font_size)
    run.bold = bold
    run.font.color.rgb = RGBColor(*color_rgb)


def create_empty_week_df():
    rows = []
    for slot in TIME_SLOTS_FLAT:
        for st_idx in range(1, STUDENTS_PER_SLOT + 1):
            row = {"GIỜ HỌC": slot, "HV": f"HV {st_idx}"}
            for day in DAYS_HEADER:
                row[day] = ""
            rows.append(row)
    return pd.DataFrame(rows)


def generate_docx(title_text, subtitle_text, weeks_data):
    doc = Document()

    # Khổ giấy A4 ngang (Landscape)
    for section in doc.sections:
        section.orientation = 1
        section.page_width = Inches(11.69)
        section.page_height = Inches(8.27)
        section.top_margin = Inches(0.4)
        section.bottom_margin = Inches(0.4)
        section.left_margin = Inches(0.4)
        section.right_margin = Inches(0.4)

    if title_text.strip():
        p_title = doc.add_paragraph()
        p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_title.paragraph_format.space_after = Pt(2)
        r_title = p_title.add_run(title_text.upper())
        r_title.font.name = "Times New Roman"
        r_title.font.size = Pt(14)
        r_title.bold = True

    if subtitle_text.strip():
        p_sub = doc.add_paragraph()
        p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_sub.paragraph_format.space_after = Pt(6)
        r_sub = p_sub.add_run(subtitle_text)
        r_sub.font.name = "Times New Roman"
        r_sub.font.size = Pt(10.5)
        r_sub.italic = True

    # Độ rộng 8 cột (Tổng ~ 10.89 inches)
    col_widths = [Inches(0.95)] + [Inches(1.42)] * 7

    for w_idx, week in enumerate(weeks_data):
        # Tổng số hàng = 2 hàng tiêu đề + (4 ca * 4 học viên) = 18 hàng, 8 cột
        total_rows = 2 + len(TIME_SLOTS) * STUDENTS_PER_SLOT
        table = doc.add_table(rows=total_rows, cols=8)
        table.style = "Table Grid"
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = False

        # Hàng 0: GIỜ HỌC | THỨ HAI ... CHỦ NHẬT
        row0 = table.rows[0]
        set_cell_background(row0.cells[0], "D9E1F2")
        for d_idx, day_name in enumerate(DAYS_HEADER):
            cell = row0.cells[d_idx + 1]
            format_cell_text(cell, day_name, bold=True, font_size=10)
            set_cell_background(cell, "D9E1F2")

        # Hàng 1: Gộp GIỜ HỌC | Ngày tháng
        row1 = table.rows[1]
        row0.cells[0].merge(row1.cells[0])
        format_cell_text(row0.cells[0], "GIỜ HỌC", bold=True, font_size=10)

        for d_idx, date_str in enumerate(week["dates"]):
            cell = row1.cells[d_idx + 1]
            format_cell_text(cell, date_str, bold=True, font_size=10, color_rgb=(192, 0, 0))
            set_cell_background(cell, "F2F2F2")

        # Các hàng ca học (mỗi ca 4 hàng con kẻ ô ngang)
        df_schedule = week["df"]
        for s_idx, slot_label in enumerate(TIME_SLOTS):
            start_row = 2 + s_idx * STUDENTS_PER_SLOT
            end_row = start_row + STUDENTS_PER_SLOT - 1

            # Gộp 4 hàng ở cột 0 (GIỜ HỌC)
            top_time_cell = table.rows[start_row].cells[0]
            bottom_time_cell = table.rows[end_row].cells[0]
            top_time_cell.merge(bottom_time_cell)
            format_cell_text(top_time_cell, slot_label, bold=True, font_size=9.5)
            set_cell_background(top_time_cell, "F9FAFB")

            # Điền dữ liệu cho từng học viên (1..4) trong ca
            for st_idx in range(STUDENTS_PER_SLOT):
                current_row_idx = start_row + st_idx
                df_row_idx = s_idx * STUDENTS_PER_SLOT + st_idx
                row = table.rows[current_row_idx]
                row.height = Inches(0.24)  # Chiều cao vừa đủ 1 dòng tên học viên

                for d_idx, day_name in enumerate(DAYS_HEADER):
                    cell = row.cells[d_idx + 1]
                    val = df_schedule.iloc[df_row_idx][day_name]
                    val_clean = "" if pd.isna(val) else str(val).strip()
                    format_cell_text(
                        cell,
                        val_clean,
                        bold=False,
                        font_size=9.5,
                        align=WD_ALIGN_PARAGRAPH.LEFT if val_clean else WD_ALIGN_PARAGRAPH.CENTER
                    )

        # Căn chỉnh độ rộng và lề toàn bảng
        for row in table.rows:
            for c_idx, cell in enumerate(row.cells):
                cell.width = col_widths[c_idx]
                set_cell_margins(cell, top=25, bottom=25, left=50, right=50)

        # Sang trang mới nếu in nhiều tuần (mỗi trang 1 tuần cho rộng rãi, dễ dán bảng)
        if w_idx < len(weeks_data) - 1:
            doc.add_page_break()

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


# ================= GIAO DIỆN STREAMLIT =================
st.title("📅 Tạo Lịch Học Tự Động (Kẻ 4 Ô / Ca Chuẩn Mẫu)")

with st.sidebar:
    st.header("⚙️ Cấu hình chung")
    doc_title = st.text_input("Tiêu đề lịch học", value="LỊCH HỌC THỰC HÀNH LÁI XE - KHÓA K6")
    doc_subtitle = st.text_input("Ghi chú phụ (tùy chọn)", value="Học viên có mặt đúng giờ quy định")

    start_date = st.date_input(
        "Ngày bắt đầu (Thứ Hai đầu tiên)",
        value=datetime(2026, 8, 10)
    )
    monday_start = start_date - timedelta(days=start_date.weekday())
    if monday_start != start_date:
        st.info(f"Đã tự động căn về Thứ Hai đầu tuần: **{monday_start.strftime('%d/%m/%Y')}**")

    num_weeks = st.number_input("Số tuần cần tạo", min_value=1, max_value=12, value=3, step=1)

    st.markdown("---")
    use_template = st.checkbox("Áp dụng lịch mẫu chung cho tất cả các tuần", value=False)

if use_template:
    st.markdown("### 🛠️ Thiết lập lịch mẫu (4 học viên / ca - Áp dụng cho mọi tuần)")
    template_df = create_empty_week_df()
    edited_template = st.data_editor(
        template_df,
        use_container_width=True,
        hide_index=True,
        disabled=["GIỜ HỌC", "HV"],
        height=600,
        key="template_editor"
    )

st.markdown("### 🗓️ Chi tiết lịch học từng tuần (Mỗi ca có sẵn 4 dòng kẻ ô)")
weeks_payload = []

for w in range(int(num_weeks)):
    week_monday = monday_start + timedelta(days=w * 7)
    week_dates = [
        f"{(week_monday + timedelta(days=d)).day}/{(week_monday + timedelta(days=d)).month}/{(week_monday + timedelta(days=d)).year}"
        for d in range(7)
    ]

    with st.expander(f"📌 Tuần {w + 1}: {week_dates[0]} – {week_dates[-1]}", expanded=(w == 0)):
        if use_template:
            base_df = edited_template.copy()
        else:
            base_df = create_empty_week_df()

        col_rename = {DAYS_HEADER[i]: f"{DAYS_HEADER[i]} ({week_dates[i]})" for i in range(7)}
        display_df = base_df.rename(columns=col_rename)

        edited_week_df = st.data_editor(
            display_df,
            use_container_width=True,
            hide_index=True,
            disabled=["GIỜ HỌC", "HV"],
            height=600,
            key=f"week_editor_{w}"
        )

        reverse_rename = {v: k for k, v in col_rename.items()}
        clean_week_df = edited_week_df.rename(columns=reverse_rename)

        weeks_payload.append({
            "dates": week_dates,
            "df": clean_week_df
        })

st.markdown("---")
if st.button("📥 Xuất file Word (.docx) kẻ ô chuẩn mẫu", type="primary", use_container_width=True):
    docx_file = generate_docx(doc_title, doc_subtitle, weeks_payload)
    file_name = f"LICH_HOC_{monday_start.strftime('%d_%m_%Y')}.docx"
    st.download_button(
        label=f"⬇️ Tải xuống {file_name}",
        data=docx_file,
        file_name=file_name,
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True
    )
