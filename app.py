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


# 1. تنظيف النص المستخرج من الذكاء الاصطناعي وترتيب الأعمدة
def process_ai_response(ai_text):
    rows = []
    lines = ai_text.strip().split("\n")
    current_site = ""

    for line in lines:
        line = line.strip()
        if not line:
            continue

        # إذا كانت السطر يعبر عن موقع جديد
        site_match = re.search(r"موقع\s*(\d+)", line)
        if site_match:
            current_site = site_match.group(1)
            continue

        # استخراج الأرقام وحدها
        nums = "".join(re.findall(r"\d+", line))
        # استخراج الحروف العربية وحدها
        letters = "".join(
            re.findall(
                r"[\u0600-\u06FF]",
                re.sub(
                    r"(موقع|رقم|نقل|تاكسي|شقق|مربع|حرف|الباء|الميم|الفاء|الراء|باء|ميم|فاء|راء|\d+)",
                    "",
                    line,
                ),
            )
        )

        plate = letters + nums if (letters or nums) else ""

        # استخراج الملاحظات والتصنيفات
        notes = []
        if "نقل" in line:
            notes.append("ن")
        if "تاكسي" in line:
            notes.append("ت")
        if "باء" in line:
            notes.append("ب")
        if "ميم" in line:
            notes.append("م")
        if "فاء" in line:
            notes.append("ف")
        if "راء" in line:
            notes.append("ر")
        if "مربع" in line:
            notes.append("مربع")
        if "شقق" in line:
            notes.append("شقق")

        classification = " ".join(notes) if notes else ""

        if plate:
            rows.append({
                "plate": plate,
                "site": current_site,
                "classification": classification,
            })
            # مسح رقم الموقع حتى لا يتكرر للسيارات التالية في نفس الموقع إلا إذا تغير
            # (نتركه يظهر أول مرة فقط كما طلبتِ)

    # جعل رقم الموقع يظهر أول مرة فقط لكل مجموعة متتالية
    seen_sites = set()
    final_rows = []
    for r in rows:
        s = r["site"]
        if s and s not in seen_sites:
            seen_sites.add(s)
            # احتفظ بالموقع لأول سيارة
            final_rows.append(r)
        else:
            # افرغ الموقع للسيارات التالية في نفس الموقع
            r_copy = r.copy()
            r_copy["site"] = ""
            final_rows.append(r_copy)

    return pd.DataFrame(final_rows)


# 2. بناء ملف Excel المنسق (يمين لليسار، 3 أعمدة)
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
        with st.spinner("🎤 جاري تفريغ الصوت وتحليله بالذكاء الاصطناعي..."):
            try:
                client = Groq(api_key=groq_api_key)
                with open("temp_audio_file.m4a", "rb") as file:
                    transcription = client.audio.transcriptions.create(
                        file=("audio.m4a", file.read()),
                        model="whisper-large-v3",
                        language="ar",
                        response_format="text",
                    )

                # استخدام نموذج لاما المتاح والسريع لتحليل النص وترتيبه بأسطر مستقلة
                chat_completion = client.chat.completions.create(
                    messages=[
                        {
                            "role": "system",
                            "content": (
                                "أنت مساعد ذكي متخصص في تنظيم تفريغ لوحات"
                                " السيارات السودانية. قم بتحليل النص المرفق"
                                " وافصله إلى أسطر. كل سطر يجب أن يحتوي على رقم"
                                " الموقع (عند ذكره)، حروف وأرقام اللوحة،"
                                " والملاحظات (نقل، تاكسي، حرف الباء، الخ). اكتب"
                                " كل سيارة في سطر مستقل."
                            ),
                        },
                        {"role": "user", "content": transcription},
                    ],
                    model="llama-3.1-8b-instant",
                )
                ai_formatted_text = chat_completion.choices[0].message.content
            except Exception as e:
                st.error(f"حدث خطأ أثناء المعالجة: {e}")
                ai_formatted_text = ""

        if ai_formatted_text:
            with st.spinner(
                "📊 جاري تنظيم البيانات في جدول Excel احترافي..."
            ):
                df = process_ai_response(ai_formatted_text)
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
