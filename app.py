import os
import re
from groq import Groq
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import pandas as pd
import streamlit as st

# ضبط واجهة التطبيق
st.set_page_config(
    page_title="تفريغ لوحات السيارات", page_icon="🚗", layout="centered"
)


# معالجة وتنظيم النص برمجيًا بدون نماذج خارجية
def parse_raw_text(raw_text):
    rows = []
    current_site = ""
    last_written_site = None

    # تنظيف وتجهيز النص لتقسيمه إلى أجزاء واضحة
    cleaned_text = re.sub(
        r"(بسم الله الرحمن الرحيم|السلام عليكم|مرحباً)", "", raw_text
    )
    lines = re.split(r"[\n\.\،]", cleaned_text)

    for line in lines:
        line = line.strip()
        if not line:
            continue

        # البحث عن رقم الموقع
        site_match = re.search(r"(?:موقع|رقم)\s*(\d+)", line)
        if site_match:
            current_site = site_match.group(1)
            if current_site != last_written_site:
                last_written_site = None
            # إذا كان السطر يحتوي على الموقع فقط، ننتقل للسطر التالي
            if len(line.replace(site_match.group(0), "").strip()) < 3:
                continue

        # استخراج الأرقام وحدها
        numbers = "".join(re.findall(r"\d+", line))

        # استخراج وتنقيه الحروف الأصلية للوحة
        letters_text = re.sub(
            r"(موقع|رقم|نقل|تاكسي|شقق|مربع|حرف|الباء|الميم|الفاء|الراء|باء|ميم|فاء|راء|\d+)",
            " ",
            line,
        )
        letters_text = re.sub(r"[أإآ]", "ا", letters_text)
        letters_text = letters_text.replace("هـ", "ه")
        letters = "".join(re.findall(r"[\u0600-\u06FF]", letters_text))

        plate = letters + numbers if (letters or numbers) else ""

        # استخراج التصنيفات والملاحظات
        notes = []
        if "نقل" in line or " ن " in line:
            notes.append("ن")
        if "تاكسي" in line or " ت " in line:
            notes.append("ت")
        if "باء" in line or "حرف الباء" in line:
            notes.append("ب")
        if "ميم" in line or "حرف الميم" in line:
            notes.append("م")
        if "فاء" in line or "حرف الفاء" in line:
            notes.append("ف")
        if "راء" in line or "حرف الراء" in line:
            notes.append("ر")
        if "مربع" in line:
            notes.append("مربع")
        if "شقق" in line:
            notes.append("شقق")

        classification = " ".join(notes) if notes else ""

        # اشتراط أن تحتوي اللوحة على أرقام واضحة
        if plate and len(numbers) >= 2:
            site_val = ""
            if current_site and last_written_site != current_site:
                site_val = current_site
                last_written_site = current_site  # كتابة الموقع مرة واحدة فقط

            rows.append({
                "plate": plate,
                "site": site_val,
                "classification": classification,
            })

    return pd.DataFrame(rows)


# بناء ملف Excel المنسق (يمين لليسار، 3 أعمدة)
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

groq_api_key = st.text_input("أدخلي مفتاح Groq API الخاص بك:", type="password")

uploaded_file = st.file_uploader(
    "اختاري ملف الصوت أو الريكورد من الجوال:", type=None
)

if uploaded_file is not None and groq_api_key:
    with open("temp_audio_file.m4a", "wb") as f:
        f.write(uploaded_file.getbuffer())

    if st.button("بدء التفريغ والاستخراج"):
        with st.spinner("🎤 جاري تفريغ وتحليل الصوت بدقة عالية..."):
            try:
                client = Groq(api_key=groq_api_key)
                with open("temp_audio_file.m4a", "rb") as file:
                    transcription = client.audio.transcriptions.create(
                        file=("audio.m4a", file.read()),
                        model="whisper-large-v3",
                        language="ar",
                        response_format="text",
                    )
                raw_text = transcription
            except Exception as e:
                st.error(f"حدث خطأ أثناء الاتصال: {e}")
                raw_text = ""

        if raw_text:
            with st.spinner("📊 جاري فرز اللوحات وبناء جدول Excel الاحترافي..."):
                df = parse_raw_text(raw_text)
                excel_file = generate_excel(df, "تفريغ_اللوحات.xlsx")

            st.success("✅ تمت المعالجة بنجاح! حملي الملف من الزر بالأسفل:")
            st.dataframe(df)

            with open(excel_file, "rb") as f:
                st.download_button(
                    label="📥 تحميل ملف Excel الجاهز",
                    data=f,
                    file_name="تفريغ_اللوحات.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
elif uploaded_file is not None and not groq_api_key:
    st.warning("⚠️ الرجاء إدخال مفتاح Groq API في الخانة المخصصة بالأعلى لبدء العمل.")
