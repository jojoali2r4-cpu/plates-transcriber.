import os
import re
from google import genai
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="تفريغ لوحات السيارات", page_icon="🚗", layout="centered"
)


def process_with_gemini(api_key, raw_text):
    client = genai.Client(api_key=api_key)
    prompt = (
        "أنت مساعد ذكاء اصطناعي متخصص في تنسيق واستخراج لوحات السيارات السودانية بدقة تامة.\n"
        "لديك نص مفرغ من تسجيل صوتي يحتوي على أرقام مواقع، أرقام لوحات، حروف متداخلة، وملاحظات (مثل نقل، تاكسي).\n"
        "مهمتك:\n"
        "1. فك التداخل واستخراج كل لوحة سيارة بشكل منظم (مثال: أ ب 4567).\n"
        "2. استخراج رقم الموقع الصحيح.\n"
        "3. استخراج التصنيف والملاحظات.\n"
        "أعطني النتيجة حصراً في أسطر مفصولة بالرمز | بهذا الترتيب لكل سيارة:\n"
        "رقم الموقع | حروف وأرقام اللوحة | التصنيف والملاحظات\n"
        "لا تضف أي مقدمات أو شرح، فقط الأسطر المطلوبة.\n\n"
        f"النص:\n{raw_text}"
    )
    response = client.models.generate_content(
        model="gemini-2.5-flash", contents=prompt
    )
    return response.text if response and response.text else ""


def create_dataframe(ai_output, raw_text=""):
    rows = []
    lines = ai_output.strip().split("\n")
    current_site = ""

    for line in lines:
        if "|" in line:
            parts = [p.strip() for p in line.split("|")]
            if len(parts) >= 2:
                site_raw = parts[0]
                plate = parts[1]
                classification = parts[2] if len(parts) > 2 else ""

                site_num = "".join(re.findall(r"\d+", site_raw))
                if site_num:
                    current_site = site_num
                    site_val = current_site
                else:
                    site_val = current_site

                if plate and plate != "-":
                    rows.append({
                        "plate": plate,
                        "site": site_val,
                        "classification": classification,
                    })

    if not rows and raw_text:
        for line in raw_text.split("\n"):
            if line.strip():
                rows.append({
                    "plate": line.strip(),
                    "site": "1",
                    "classification": "",
                })

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


st.title("🚗 تفريغ لوحات السيارات (مدعوم بـ Gemini)")
gemini_api_key = st.text_input(
    "أدخلي مفتاح Gemini API الخاص بك:", type="password"
)
uploaded_text = st.text_area(
    "أو الصقي النص الخام هنا مباشرة للمعالجة السريعة:"
)

if gemini_api_key and uploaded_text:
    if st.button("معالجة وتوليد Excel"):
        with st.spinner("جاري التحليل والترتيب..."):
            ai_out = process_with_gemini(gemini_api_key, uploaded_text)
            df = create_dataframe(ai_out, uploaded_text)
            file_path = generate_excel(df)

        st.success("تم بنجاح!")
        st.dataframe(df)
        with open(file_path, "rb") as f:
            st.download_button(
                "تحميل ملف Excel",
                f,
                file_name="تفريغ_اللوحات.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
