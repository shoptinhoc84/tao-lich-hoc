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

DAYS_HEADER = ["THỨ HAI", "THỨ BA", "THỨ TƯ", "THỨ NĂM", "THỨ SÁU", "THỨ BẢY", "CHỦ NHẬT"]


def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)


def set_cell_margins(cell, top=60, bottom=60, left=80, right=80):
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


def generate_docx(title_text, subtitle_text, weeks_data):
    doc = Document()

    # Thiết lập khổ giấy A4 ngang (Landscape) để bảng 8 cột rộng rãi
    for section in doc.sections:
        section.orientation = 1  # Landscape
        section.page_width = Inches(11.69)
        section.page_height = Inches(8.27)
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)

    # Tiêu đề văn bản
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
        p_sub.paragraph_format.space_after = Pt(8)
        r_sub = p_sub.add_run(subtitle_text)
        r_sub.font.name = "Times New Roman"
        r_sub.font.size = Pt(11)
        r_sub.italic = True

    # Độ rộng từng cột (Tổng ~ 10.69 inches)
    col_widths = [Inches(1.09)] + [Inches(1.37)] * 7

    for w_idx, week in enumerate(weeks_data):
        # Tạo bảng: 2 hàng tiêu đề (Thứ + Ngày) + 4 hàng ca học = 6 hàng, 8 cột
        table = doc.add_table(rows=6, cols=8)
        table.style = "Table Grid"
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = False

        # Hàng 0: GIỜ HỌC | THỨ HAI ... CHỦ NHẬT
        row0 = table.rows[0]
        format_cell_text(row0.cells[0], "GIỜ HỌC", bold=True, font_size=10)
        set_cell_background(row0.cells[0], "D9E1F2")

        for d_idx, day_name in enumerate(DAYS_HEADER):
            cell = row0.cells[d_idx + 1]
            format_cell_text(cell, day_name, bold=True, font_size=10)
            set_cell_background(cell, "D9E1F2")

        # Hàng 1: Trống (gộp với GIỜ HỌC) | Ngày tháng (10/8/2026 ...)
        row1 = table.rows[1]
        row0.cells[0].merge(row1.cells[0])
        format_cell_text(row0.cells[0], "GIỜ HỌC", bold=True, font_size=10)

        for d_idx, date_str in enumerate(week["dates"]):
            cell = row1.cells[d_idx + 1]
            format_cell_text(cell, date_str, bold=True, font_size=10, color_rgb=(192, 0, 0))
            set_cell_background(cell, "F2F2F2")

        # Các hàng ca học (Hàng 2 -> 5)
        df_schedule = week["df"]
        for s_idx, slot_label in enumerate(TIME_SLOTS):
            row = table.rows[s_idx + 2]
            row.height = Inches(0.85)  # Chiều cao mỗi ô ca học đủ chứa 3-4 dòng nội dung

            # Cột 0: Giờ học
            cell_time = row.cells[0]
            format_cell_text(cell_time, slot_label, bold=True, font_size=9.5)
            set_cell_background(cell_time, "F9FAFB")

            # Cột 1 -> 7: Nội dung từng thứ
            for d_idx, day_name in enumerate(DAYS_HEADER):
                cell = row.cells[d_idx + 1]
                val = df_schedule.iloc[s_idx][day_name]
                val_clean = "" if pd.isna(val) else str(val).strip()
                format_cell_text(cell, val_clean, bold=False, font_size=9.5)

        # Áp dụng độ rộng và lề cho toàn bộ ô
        for row in table.rows:
            for c_idx, cell in enumerate(row.cells):
                cell.width = col_widths[c_idx]
                set_cell_margins(cell, top=50, bottom=50, left=60, right=60)

        # Khoảng cách giữa các tuần
        if w_idx < len(weeks_data) - 1:
            p_space = doc.add_paragraph()
            p_space.paragraph_format.space_before = Pt(4)
            p_space.paragraph_format.space_after = Pt(4)

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


# ================= GIAO DIỆN STREAMLIT =================
st.title("📅 Tạo Lịch Học Tự Động & Xuất File Word (.docx)")

with st.sidebar:
    st.header("⚙️ Cấu hình chung")
    doc_title = st.text_input("Tiêu đề lịch học", value="LỊCH HỌC LÝ THUYẾT & THỰC HÀNH - KHÓA K6")
    doc_subtitle = st.text_input("Ghi chú phụ (tùy chọn)", value="Học viên có mặt trước giờ học 10 phút để điểm danh")

    start_date = st.date_input(
        "Ngày bắt đầu (Chọn Thứ Hai đầu tiên)",
        value=datetime(2026, 8, 10)
    )
    # Tự động điều chỉnh về Thứ Hai của tuần đó nếu người dùng chọn nhầm thứ khác
    monday_start = start_date - timedelta(days=start_date.weekday())
    if monday_start != start_date:
        st.info(f"Đã tự động căn về Thứ Hai đầu tuần: **{monday_start.strftime('%d/%m/%Y')}**")

    num_weeks = st.number_input("Số tuần cần tạo", min_value=1, max_value=12, value=3, step=1)

    st.markdown("---")
    st.subheader("📋 Điền nhanh mẫu lặp lại")
    use_template = st.checkbox("Áp dụng lịch mẫu chung cho tất cả các tuần", value=False)

# Khởi tạo dữ liệu mẫu nếu người dùng bật chế độ lặp lại
if use_template:
    st.markdown("### 🛠️ Thiết lập khung lịch mẫu (Áp dụng cho mọi tuần)")
    st.caption("Nhập nội dung vào các ô bên dưới (dùng Shift + Enter để xuống dòng trong 1 ô). Nội dung này sẽ tự động điền cho tất cả các tuần.")
    template_df = pd.DataFrame(
        [[""] * 7 for _ in range(4)],
        columns=DAYS_HEADER
    )
    template_df.insert(0, "GIỜ HỌC", ["07:30 - 09:30", "09:30 - 11:30", "13:00 - 15:00", "15:00 - 17:00"])
    edited_template = st.data_editor(
        template_df,
        use_container_width=True,
        hide_index=True,
        disabled=["GIỜ HỌC"],
        key="template_editor"
    )

st.markdown("### 🗓️ Chi tiết lịch học từng tuần")
weeks_payload = []

for w in range(int(num_weeks)):
    week_monday = monday_start + timedelta(days=w * 7)
    week_dates = [(week_monday + timedelta(days=d)).strftime("%-d/%-m/%Y") if hasattr(datetime, "strftime") else "" for d in range(7)]
    # Format ngày chuẩn d/m/yyyy tương thích cả Windows và Linux
    week_dates = [f"{(week_monday + timedelta(days=d)).day}/{(week_monday + timedelta(days=d)).month}/{(week_monday + timedelta(days=d)).year}" for d in range(7)]

    st.markdown(f"**Tuần {w + 1}** (`{week_dates[0]}` – `{week_dates[-1]}`)")

    if use_template:
        base_df = edited_template.copy()
    else:
        base_df = pd.DataFrame(
            [[""] * 7 for _ in range(4)],
            columns=DAYS_HEADER
        )
        base_df.insert(0, "GIỜ HỌC", ["07:30 - 09:30", "09:30 - 11:30", "13:00 - 15:00", "15:00 - 17:00"])

    # Đổi tên cột hiển thị kèm ngày tháng cho trực quan
    col_rename = {DAYS_HEADER[i]: f"{DAYS_HEADER[i]} ({week_dates[i]})" for i in range(7)}
    display_df = base_df.rename(columns=col_rename)

    edited_week_df = st.data_editor(
        display_df,
        use_container_width=True,
        hide_index=True,
        disabled=["GIỜ HỌC"],
        key=f"week_editor_{w}"
    )

    # Đổi lại tên cột gốc để đưa vào hàm xuất Word
    reverse_rename = {v: k for k, v in col_rename.items()}
    clean_week_df = edited_week_df.rename(columns=reverse_rename)

    weeks_payload.append({
        "dates": week_dates,
        "df": clean_week_df
    })

st.markdown("---")
if st.button("📥 Xuất file Word (.docx)", type="primary", use_container_width=True):
    docx_file = generate_docx(doc_title, doc_subtitle, weeks_payload)
    file_name = f"LICH_HOC_{monday_start.strftime('%d_%m_%Y')}.docx"
    st.download_button(
        label=f"⬇️ Tải xuống {file_name}",
        data=docx_file,
        file_name=file_name,
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True
    )