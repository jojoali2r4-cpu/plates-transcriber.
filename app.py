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


# 1. استخراج وتنقية حروف وأرقام اللوحة (العمود الأول - أقصى اليمين)
def extract_plate(text):
    # إزالة كلمات المواقع والملاحظات الشائعة لترك الحروف والأرقام فقط
    cleaned = re.sub(
        r"(موقع|رقم|نقل|تاكسي|شقق|مربع|حرف|الباء|الميم|الفاء|الراء|باء|ميم|فاء|راء)",
        " ",
        text,
    )

    # توحيد الأشكال (الألفات والهاءات)
    cleaned = re.sub(r"[أإآ]", "ا", cleaned)
    cleaned = cleaned.replace("هـ", "ه")

    # استخراج الحروف العربية المفردة وتكوينها
    letters = "".join(re.findall(r"[\u0600-\u06FF]", cleaned))
    # استخراج الأرقام
    numbers = "".join(re.findall(r"\d+", text))

    if letters or numbers:
        return letters + numbers
    return None


# 2. استخراج التصنيفات والملاحظات بدقة (العمود الثالث)
def extract_classification(text):
    notes = []
    if "نقل" in text:
        notes.append("ن")
    if "تاكسي" in text:
        notes.append("ت")
    if "باء" in text or "حرف الباء" in text:
        notes.append("ب")
    if "ميم" in text or "حرف الميم" in text:
        notes.append("م")
    if "فاء" in text or "حرف الفاء" in text:
        notes.append("ف")
    if "راء" in text or "حرف الراء" in text:
        notes.append("ر")
    if "مربع" in text:
        notes.append("مربع")
    if "شقق" in text:
        notes.append("شقق")

    return " ".join(notes) if notes else ""


# 3. معالجة النص المفرغ وترتيب البيانات حسب الضوابط الدقيقة
def process_text_data(raw_text):
    rows = []
    current_site = ""
    last_written_site = None

    lines = raw_text.split("\n")
    for line in lines:
        line = line.strip()
        if not line:
            continue

        # البحث عن رقم الموقع الجديد
        site_match = re.search(r"موقع\s*(?:رقم)?\s*(\d+)", line)
        if site_match:
            current_site = site_match.group(1)
            last_written_site = (
                None  # لإعادة كتابة الموقع الجديد أول سيارة فقط
            )
            continue

        # التحقق إذا كانت السطر يحتوي على سيارة/لوحة
        plate = extract_plate(line)
        if plate:
            classification = extract_classification(line)

            # كتابة رقم الموقع مرة واحدة فقط لأول سيارة في الموقع
            site_val = ""
            if current_site and last_written_site != current_site:
                site_val = current_site
                last_written_site = current_site  # عدم تكراره للسيارات التالية

            rows.append({
                "plate": plate,
                "site": site_val,
                "classification": classification,
            })

    return pd.DataFrame(rows)


# 4. بناء ملف Excel المنسق (يمين لليسار، 3 أعمدة بالعربي)
def generate_excel(df, output_filename="تفريغ_اللوحات.xlsx"):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "تفريغ اللوحات"
    ws.views.sheetView[0].rightToLeft = True

    # الأعمدة بالترتيب الصحيح من اليمين لليسار
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
        with st.spinner("🎤 جاري تفريغ الصوت وتحليله بدقة عالية..."):
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
            with st.spinner("📊 جاري فرز اللوحات والمواقع وتجهيز ملف Excel..."):
                df = process_text_data(raw_text)
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
