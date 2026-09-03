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


# دالة معالجة برمجية ذكية للنص المفرغ بدون الحاجة لنموذج دردشة (لتجنب أخطاء API تماماً)
def smart_parse_transcription(raw_text):
    rows = []
    current_site = ""
    last_written_site = None

    # تقسيم النص إلى أسطر أو جمل بناءً على النقاط أو الفواصل أو الأسطر الجديدة
    lines = re.split(r"[\n\.\،؛]", raw_text)

    for line in lines:
        line = line.strip()
        if not line:
            continue

        # استخراج رقم الموقع (مثل موقع 5609 أو رقم 5609)
        site_match = re.search(r"(?:موقع|رقم)?\s*(\d{3,5})", line)
        if "موقع" in line or "رقم" in line or "مربع" in line:
            if site_match:
                current_site = site_match.group(1)
                if current_site != last_written_site:
                    last_written_site = None
                # إذا كان السطر يمثل إعلان الموقع فقط
                if len(line.replace(site_match.group(0), "").strip()) < 4:
                    continue

        # استخراج الأرقام الخاصة باللوحة
        numbers_found = re.findall(r"\d+", line)
        plate_numbers = ""
        for num in numbers_found:
            if len(num) >= 2 and num != current_site:
                plate_numbers = num
                break

        # استخراج وتصفية الحروف العربية للوحة
        cleaned_letters = re.sub(
            r"(موقع|رقم|نقل|تاكسي|شقق|مربع|حرف|الباء|الميم|الفاء|الراء|باء|ميم|فاء|راء|\d+)",
            " ",
            line,
        )
        cleaned_letters = re.sub(r"[أإآ]", "ا", cleaned_letters)
        cleaned_letters = cleaned_letters.replace("هـ", "ه")

        # محاولة استخلاص الكلمات العربية الواضحة كحروف للوحة
        arabic_words = re.findall(r"[\u0600-\u06FF]{2,}", cleaned_letters)
        letters = " ".join(arabic_words) if arabic_words else ""

        # استخراج الملاحظات والتصنيفات
        notes = []
        if any(w in line for w in ["نقل", " ن ", " حرف النون"]):
            notes.append("ن")
        if any(w in line for w in ["تاكسي", " ت ", "تاكسى", " حرف التاء"]):
            notes.append("ت")
        if "باء" in line or " حرف الباء" in line:
            notes.append("ب")
        if "ميم" in line or " حرف الميم" in line:
            notes.append("م")
        if "فاء" in line or " حرف الفاء" in line:
            notes.append("ف")
        if "راء" in line or " حرف الراء" in line:
            notes.append("ر")
        if "مربع" in line:
            notes.append("مربع")
        if "شقق" in line:
            notes.append("شقق")

        classification = " ".join(notes) if notes else ""

        # تجميع اللوحة النهائية
        plate = ""
        if letters and plate_numbers:
            plate = f"{letters} {plate_numbers}"
        elif plate_numbers:
            plate = plate_numbers
        elif letters:
            plate = letters

        if plate:
            site_val = ""
            if current_site and last_written_site != current_site:
                site_val = current_site
                last_written_site = current_site

            rows.append({
                "plate": plate,
                "site": site_val,
                "classification": classification,
            })

    # إذا لم يستخرج النظام أسطر كافية، نقوم بتقسيم النص البسيط كلمة بكلمة
    if not rows:
        words = raw_text.split()
        temp_plate = ""
        for word in words:
            if re.match(r"^\d{2,4}$", word):
                temp_plate = word
            elif re.match(r"^[\u0600-\u06FF]{1,4}$", word) and temp_plate:
                rows.append({
                    "plate": f"{word} {temp_plate}",
                    "site": current_site,
                    "classification": "",
                })
                temp_plate = ""
        if not rows:
            # حل أخير: وضع النص المفرغ كاملاً في السطر الأول إذا تعذر الفرز التلقائي
            rows.append({
                "plate": raw_text[:50],
                "site": current_site,
                "classification": "",
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
                st.error(f"حدث خطأ أثناء تفريغ الصوت: {e}")
                raw_text = ""

        if raw_text:
            with st.spinner("📊 جاري فرز البيانات واستخراج ملف Excel..."):
                df = smart_parse_transcription(raw_text)
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
        else:
            st.error("تعذر إتمام العملية. تأكدي من صحة مفتاح API أو جربي ملفاً آخر.")
elif uploaded_file is not None and not groq_api_key:
    st.warning("⚠️ الرجاء إدخال مفتاح Groq API في الخانة المخصصة بالأعلى لبدء العمل.")
