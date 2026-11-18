import streamlit as st
import openpyxl
from google import genai
from google.genai import types
import json
import io
import time

st.set_page_config(page_title="Նկարից Excel լրացնող գործիք", page_icon="📊")
st.title("📊 Փաստաթղթերի ընթերցում և Excel-ի լրացում")

api_key = st.sidebar.text_input("Մուտքագրեք Gemini API Key-ը", type="password")

excel_file = st.file_uploader("1. Վերբեռնեք ձեր Excel ֆորման (.xlsx)", type=["xlsx"])
doc_file = st.file_uploader("2. Վերբեռնեք նկարը կամ PDF-ը", type=["png", "jpg", "jpeg", "pdf"])
header_row = st.number_input("Ո՞ր տողում են գտնվում սյունակների վերնագրերը", min_value=1, value=1)

if st.button("Կարդալ և լրացնել Excel-ը"):
    if not api_key:
        st.error("Խնդրում ենք ձախ դաշտում մուտքագրել Gemini API Key-ը։")
    elif not excel_file or not doc_file:
        st.error("Խնդրում ենք վերբեռնել երկու ֆայլերն էլ։")
    else:
        with st.spinner("AI-ը կարդում է փաստաթուղթը և լրացնում աղյուսակը..."):
            try:
                wb = openpyxl.load_workbook(excel_file)
                sheet = wb.active
                
                headers = []
                for col in range(1, sheet.max_column + 1):
                    val = sheet.cell(row=header_row, column=col).value
                    headers.append(str(val).strip() if val else f"Column_{col}")

                client = genai.Client(api_key=api_key)
                file_bytes = doc_file.read()
                mime_type = doc_file.type

                prompt = f"""
                Վերլուծիր այս փաստաթուղթը և հանիր տվյալները՝ ըստ հետևյալ սյունակների անունների՝ {headers}։
                Վերադարձրու արդյունքը ՄԻԱՅՆ որպես JSON զանգված (list of objects), որտեղ յուրաքանչյուր օբյեկտի բանալիները (keys) ճշգրիտ համընկնում են տրված սյունակների անունների հետ։
                Եթե որևէ սյունակի համար տվյալ չկա, դիր դատարկ տող ""։
                """

                # Մինչև 3 փորձ՝ սերվերի ծանրաբեռնվածության դեպքում
                response = None
                for attempt in range(3):
                    try:
                        response = client.models.generate_content(
                            model="gemini-3.8-flash",
                            contents=[
                                types.Part.from_bytes(data=file_bytes, mime_type=mime_type),
                                prompt
                            ],
                            config=types.GenerateContentConfig(
                                response_mime_type="application/json"
                            )
                        )
                        break
                    except Exception as e:
                        if "503" in str(e) and attempt < 2:
                            time.sleep(3)
                            continue
                        raise e

                extracted_rows = json.loads(response.text)
                
                start_row = header_row + 1
                while sheet.cell(row=start_row, column=1).value is not None:
                    start_row += 1

                for row_idx, row_data in enumerate(extracted_rows):
                    current_row = start_row + row_idx
                    for col_idx, header_name in enumerate(headers, start=1):
                        if header_name in row_data:
                            sheet.cell(row=current_row, column=col_idx, value=row_data[header_name])

                output = io.BytesIO()
                wb.save(output)
                output.seek(0)

                st.success(f"Պատրաստ է։ Ավելացվեց {len(extracted_rows)} տող։")
                st.download_button(
                    label="📥 Ներբեռնել լրացված Excel-ը",
                    data=output,
                    file_name="filled_form.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.error(f"Տեղի ունեցավ սխալ. {e}")
