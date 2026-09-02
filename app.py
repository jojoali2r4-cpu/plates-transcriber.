import os
import re
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import pandas as pd
import streamlit as st
import whisper

# ضبط واجهة التطبيق
st.set_page_config(
    page_title="تفريغ لوحات السيارات", page_icon="🚗", layout="centered"
)


# 1. تنظيف حروف وأرقام السيارة (العمود A)
def normalize_plate(text):
    text = re.sub(r"[أإآ]", "ا", text)
    text = text.replace("هـ", "ه")
    letters = "".join(re.findall(r"[\u0600-\u06FF]", text))
    numbers = "".join(re.findall(r"\d+", text))
    return letters + numbers


# 2. استخراج الملاحظات والتصنيفات (العمود C)
def parse_classification(text):
    notes = []
    keywords = [
        ("نقل", "ن"),
        ("تاكسي", "ت"),
        ("حرف الباء", "ب"),
        ("حرف الميم", "م"),
        ("حرف الفاء", "ف"),
        ("حرف الراء", "ر"),
        ("مربع", "مربع"),
        ("شقق", "شقق"),
    ]
    for kw, symbol in keywords:
        if kw in text:
            notes.append(symbol)
    return " ".join(notes) if notes else None


# 3. معالجة النص المفرغ وترتيب البيانات
def process_text_data(raw_text):
    rows = []
    current_site = None
    site_written = False

    for line in raw_text.split("\n"):
        line = line.strip()
        if not line:
            continue

        site_match = re.search(r"موقع\s*(?:رقم)?\s*(\d+)", line)
        if site_match:
            current_site = site_match.group(1)
            site_written = False
            continue

        plate = normalize_plate(line)
        if plate:
            classification = parse_classification(line)
            site_val = ""
            if current_site and not site_written:
                site_val = current_site
                site_written = True

            rows.append({
                "plate": plate,
                "site": site_val,
                "classification": classification,
            })

    return pd.DataFrame(rows)


# 4. بناء ملف Excel المنسق (اتجاه RTL وحدود وألوان)
def generate_excel(df, output_filename="تفريغ_اللوحات.xlsx"):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "تفريغ اللوحات"
    ws.views.sheetView[0].rightToLeft = True

    headers = ["حروف وأرقام السيارة", "رقم الموقع", "التصنيف والملاحظات"]
    ws.append(headers)

    for _, row in df.iterrows():
        ws.append([row["plate"], row["site"], row["classification"]])

    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill(
        start_color="1F4E78", end_color="1F4E78", fill_type="solid"
    )
    data_font = Font(name="Calibri", size=11)
    center_align = Alignment(horizontal="center", vertical="center")
    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )

    for col_num in range(1, 4):
        cell = ws.cell(row=1, column=col_num)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center_align

    for row_num in range(2, len(df) + 2):
        for col_num in range(1, 4):
            cell = ws.cell(row=row_num, column=col_num)
            cell.font = data_font
            cell.alignment = center_align
            cell.border = thin_border
            if cell.value == "" or cell.value is None:
                cell.value = None

    ws.column_dimensions["A"].width = 25
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 22

    wb.save(output_filename)
    return output_filename


# --- الواجهة التفاعلية للمستخدم ---
st.title("🚗 تطبيق تفريغ لوحات السيارات")
st.write("ارفعي ملف التسجيل الصوتي للحصول على ملف Excel جاهز ومنسق.")

# السماح باختيار أي نوع ملف صوتي أو ريكورد من الجوال
uploaded_file = st.file_uploader(
    "اختاري ملف الصوت أو الريكورد من الجوال:", type=None
)

if uploaded_file is not None:
    # حفظ التسجيل مؤقتاً للمعالجة
    with open("temp_audio_file", "wb") as f:
        f.write(uploaded_file.getbuffer())

    if st.button("بدء التفريغ والاستخراج"):
        with st.spinner("🎤 جاري الاستماع للصوت وتحويله إلى نص..."):
            # استخدام نموذج tiny الخفيف لتجنب حظر المعالج
            model = whisper.load_model("tiny")
            result = model.transcribe("temp_audio_file", language="ar")

        with st.spinner("📊 جاري ترتيب البيانات وبناء ملف Excel..."):
            df = process_text_data(result["text"])
            excel_file = generate_excel(df, "تفريغ_اللوحات.xlsx")

        st.success("✅ تمت المعالجة بنجاح!")
        st.dataframe(df)

        with open(excel_file, "rb") as f:
            st.download_button(
                label="📥 تحميل ملف Excel الجاهز",
                data=f,
                file_name="تفريغ_اللوحات.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
