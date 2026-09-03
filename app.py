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


# تحليل النص المفرغ وتنظيمه بذكاء الاصطناعي لاستخراج اللوحات بدقة
def parse_text_with_llm(client, raw_text):
    try:
        response = client.chat.completions.create(
            model="llama-3.1-8b-instant",
            messages=[
                {
                    "role": "system",
                    "content": (
                        "أنت خبير في تنظيم وتصحيح تفريغ لوحات السيارات السودانية."
                        " النص القادم هو تفريغ صوتي يحتوي على أرقام مواقع ولوحات"
                        " سيارات وملاحظات (مثل نقل، تاكسي، حرف الباء، ميم، فاء،"
                        " راء).\nمهمتك:\n1. استخراج كل سيارة في سطر مستقل."
                        "\n2. تصحيح حروف وأرقام اللوحة بحيث تتكون من الحروف"
                        " (مثلا: ب م ر أو ج د ك) متلاصقة مع الأرقام (أربعة"
                        " أرقام) بدون أي كلام زائد.\n3. تحديد رقم الموقع (يكتب"
                        " أول مرة فقط للموقع أو عند تغييره).\n4. استخراج"
                        " الملاحظات (ن لنقل، ت لتاكسي، ب، م، ف، ر) في عمود"
                        " التصنيف.\nأعطني النتيجة مباشرة كقائمة مفصولة بالرمز"
                        " | بالترتيب التالي لكل سطر:\nرقم الموقع | حروف وأرقام"
                        " اللوحة | التصنيف والملاحظات"
                    ),
                },
                {"role": "user", "content": raw_text},
            ],
            temperature=0.1,
        )
        return response.choices[0].message.content
    except Exception as e:
        return str(e)


# تحويل مخرجات الذكاء الاصطناعي إلى جدول بيانات
def create_dataframe_from_ai(ai_output):
    rows = []
    lines = ai_output.strip().split("\n")
    current_site = ""

    for line in lines:
        if "|" in line:
            parts = [p.strip() for p in line.split("|")]
            if len(parts) >= 3:
                site_raw = parts[0]
                plate = parts[1]
                classification = parts[2]

                # تحديث الموقع فقط إذا وجد ونظفه من الحروف الزائدة
                site_num = "".join(re.findall(r"\d+", site_raw))
                if site_num:
                    current_site = site_num
                    site_val = current_site
                else:
                    site_val = ""

                if plate and len(re.findall(r"\d+", plate)) >= 2:
                    rows.append({
                        "plate": plate,
                        "site": site_val,
                        "classification": (
                            classification if classification != "-" else ""
                        ),
                    })

    # ضمان ظهور رقم الموقع لأول سيارة فقط في كل مجموعة متتالية
    seen_sites = set()
    final_rows = []
    for r in rows:
        s = r["site"]
        if s:
            if s not in seen_sites:
                seen_sites.add(s)
                final_rows.append(r)
            else:
                r_copy = r.copy()
                r_copy["site"] = ""
                final_rows.append(r_copy)
        else:
            final_rows.append(r)

    return pd.DataFrame(final_rows)


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

                # تحليل النص عبر نموذج Llama 3.1 8B السريع والمستقر
                ai_output = parse_text_with_llm(client, transcription)
            except Exception as e:
                st.error(f"حدث خطأ أثناء الاتصال: {e}")
                ai_output = ""

        if ai_output and "Error" not in ai_output:
            with st.spinner("📊 جاري ترتيب البيانات واستخراج ملف Excel..."):
                df = create_dataframe_from_ai(ai_output)
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
            st.error(
                "تعذر تحليل النص بالذكاء الاصطناعي. تأكدي من صحة مفتاح API."
            )
elif uploaded_file is not None and not groq_api_key:
    st.warning("⚠️ الرجاء إدخال مفتاح Groq API في الخانة المخصصة بالأعلى لبدء العمل.")
