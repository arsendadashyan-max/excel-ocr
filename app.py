import streamlit as st
import openpyxl
from google import genai
from google.genai import types
import json
import io
import time

st.set_page_config(page_title="Խելացի Excel լրացնող գործիք", page_icon="📊")
st.title("📊 Փաստաթղթերի վերլուծություն՝ նախկին փորձի հիման վրա")

api_key = st.sidebar.text_input("Մուտքագրեք Gemini API Key-ը", type="password")

# 1. Դատարկ ֆորման
excel_file = st.file_uploader("1. Վերբեռնեք ձեր դատարկ Excel ֆորման (.xlsx)", type=["xlsx"])

# 2. Նախկին գործարքների օրինակներ (Ուսուցման համար)
history_files = st.file_uploader(
    "2. Վերբեռնեք նախկին ճիշտ լրացված Excel ֆայլեր (որպես օրինակ AI-ի համար)", 
    type=["xlsx"],
    accept_multiple_files=True
)

# 3. Նոր փաստաթղթեր
doc_files = st.file_uploader(
    "3. Վերբեռնեք նոր փաստաթղթերը (նկարներ, PDF-ներ կամ Excel ֆայլեր)", 
    type=["png", "jpg", "jpeg", "pdf", "xlsx"],
    accept_multiple_files=True
)

header_row = st.number_input("Ո՞ր տողում են գտնվում սյունակների վերնագրերը", min_value=1, value=1)

extra_instructions = st.text_area(
    "Լրացուցիչ ցուցումներ AI-ին (ոչ պարտադիր)", 
    placeholder="Օրինակ՝ Եթե կազմակերպությունը կա նախկին գործարքներում, վերցրու անունն ու կոդը ճիշտ այնպես, ինչպես այնտեղ է գրված..."
)

def extract_text_from_excel(file_obj, max_rows=150):
    source_wb = openpyxl.load_workbook(file_obj, data_only=True)
    text_lines = []
    for sheet_name in source_wb.sheetnames:
        ws = source_wb[sheet_name]
        text_lines.append(f"--- Էջ՝ {sheet_name} ---")
        row_count = 0
        for row in ws.iter_rows(values_only=True):
            row_vals = [str(cell).strip() if cell is not None else "" for cell in row]
            if any(row_vals):
                text_lines.append(" | ".join(row_vals))
                row_count += 1
                if row_count >= max_rows:
                    break
    return "\n".join(text_lines)

def safe_parse_json(raw_text):
    try:
        return json.loads(raw_text)
    except json.JSONDecodeError:
        last_brace = raw_text.rfind("}")
        if last_brace != -1:
            repaired = raw_text[:last_brace + 1] + "]"
            return json.loads(repaired)
        raise

if st.button("Ուսումնասիրել նախկին փորձը և լրացնել Excel-ը"):
    if not api_key:
        st.error("Խնդրում ենք ձախ դաշտում մուտքագրել Gemini API Key-ը։")
    elif not excel_file or not doc_files:
        st.error("Խնդրում ենք վերբեռնել ձեր դատարկ ֆորման և նոր փաստաթղթերը։")
    else:
        with st.spinner("AI-ը ուսումնասիրում է նախկին գործարքները և վերլուծում նոր փաստաթղթերը..."):
            try:
                wb = openpyxl.load_workbook(excel_file)
                sheet = wb.active
                
                headers = []
                for col in range(1, sheet.max_column + 1):
                    val = sheet.cell(row=header_row, column=col).value
                    headers.append(str(val).strip() if val else f"Column_{col}")

                client = genai.Client(api_key=api_key)
                request_contents = []

                # Եթե կան նախկին գործարքների օրինակներ, նախ դրանք ենք տալիս AI-ին
                if history_files:
                    request_contents.append("=== ԲԱԺԻՆ 1. ՆԱԽԿԻՆ ՃԻՇՏ ԼՐԱՑՎԱԾ ԳՈՐԾԱՐՔՆԵՐԻ ՕՐԻՆԱԿՆԵՐ (ՈՒՍՈՒՑՄԱՆ ՀԱՄԱՐ) ===")
                    for idx, hist_file in enumerate(history_files, start=1):
                        hist_text = extract_text_from_excel(hist_file, max_rows=100)
                        request_contents.append(f"--- Նախկին օրինակ #{idx} ({hist_file.name}) ---\n{hist_text}")

                # Այնուհետև ավելացնում ենք նոր փաստաթղթերը
                request_contents.append("=== ԲԱԺԻՆ 2. ՆՈՐ ՓԱՍՏԱԹՂԹԵՐ (ՈՐՈՆՔ ՊԵՏՔ Է ՄՇԱԿԵԼ ԵՎ ԼՐԱՑՆԵԼ) ===")
                for idx, current_file in enumerate(doc_files, start=1):
                    request_contents.append(f"--- Նոր փաստաթուղթ #{idx} ({current_file.name}) ---")
                    if current_file.name.lower().endswith(".xlsx"):
                        excel_text = extract_text_from_excel(current_file, max_rows=300)
                        request_contents.append(excel_text)
                    else:
                        file_bytes = current_file.read()
                        mime_type = current_file.type
                        request_contents.append(
                            types.Part.from_bytes(data=file_bytes, mime_type=mime_type)
                        )

                prompt = f"""
                Դու բարձրակարգ հաշվապահական և տվյալների վերլուծության մասնագետ ես։ 
                
                ՔՈ ԱՌԱՋԱԴՐԱՆՔԸ.
                1. **Սովորիր նախկին փորձից.** Եթե վերևում տրված են «ԲԱԺԻՆ 1. ՆԱԽԿԻՆ ՃԻՇՏ ԼՐԱՑՎԱԾ ԳՈՐԾԱՐՔՆԵՐԻ ՕՐԻՆԱԿՆԵՐ», ուշադիր ուսումնասիրիր դրանք։ Տես, թե ինչպես են այնտեղ գրված կազմակերպությունների անունները, ինչ ստանդարտ ձևաչափեր, հապավումներ, կատեգորիաներ կամ հաշվապահական մոտեցումներ են կիրառվել։
                2. **Իդենտիֆիկացրու անունները.** Նոր փաստաթղթերում (ԲԱԺԻՆ 2) հանդիպող անունները համեմատիր նախկին գործարքների բազայի հետ։ Եթե նույն կազմակերպությունը կամ անձը արդեն հանդիպել է նախկին օրինակներում (նույնիսկ եթե նոր փաստաթղթում գրված է այլ լեզվով, տրանսլիտով կամ փոքր տարբերությամբ), օգտագործիր նախկին ճիշտ լրացված անվանաձևն ու կապակցված տվյալները։
                3. **Լրացրու նոր տվյալները.** Մշակիր ՄԻԱՅՆ «ԲԱԺԻՆ 2»-ի նոր փաստաթղթերը և կազմիր տվյալները՝ ըստ հետևյալ սյունակների անունների՝ {headers}։
                4. **Կոմպակտ JSON.** Վերադարձրու կոմպակտ JSON՝ առանց ավելորդ բացատների և տողադարձերի։
                5. **Օգտատիրոջ հատուկ ցուցումներ.** {extra_instructions if extra_instructions else "Չկան"}

                Վերադարձրու արդյունքը ՄԻԱՅՆ որպես JSON զանգված (list of objects), որտեղ յուրաքանչյուր օբյեկտի բանալիները (keys) ճշգրիտ համընկնում են տրված սյունակների անունների հետ՝ {headers}։ Եթե որևէ դաշտի տվյալ չկա, դիր դատարկ տող ""։
                """
                request_contents.append(prompt)

                available_models = []
                for m in client.models.list():
                    name = m.name.replace("models/", "")
                    if "gemini" in name and "embedding" not in name and "imagen" not in name and "lite" not in name:
                        available_models.append(name)

                available_models.sort(key=lambda x: ("pro" not in x, x), reverse=False)

                response = None
                last_error = None

                for model_name in available_models:
                    try:
                        response = client.models.generate_content(
                            model=model_name,
                            contents=request_contents,
                            config=types.GenerateContentConfig(
                                response_mime_type="application/json",
                                temperature=0.0,
                                max_output_tokens=65536
                            )
                        )
                        st.info(f"Օգտագործվեց մոդել՝ {model_name}")
                        break
                    except Exception as e:
                        last_error = e
                        time.sleep(1)
                        continue

                if response is None:
                    raise Exception(f"Չհաջողվեց կապվել որևէ մոդելի հետ։ Սխալ՝ {last_error}")

                extracted_rows = safe_parse_json(response.text)
                
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

                st.success(f"Պատրաստ է։ Նախկին փորձի հիման վրա ավելացվեց {len(extracted_rows)} տող։")
                st.download_button(
                    label="📥 Ներբեռնել լրացված Excel-ը",
                    data=output,
                    file_name="filled_form.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.error(f"Տեղի ունեցավ սխալ. {e}")
