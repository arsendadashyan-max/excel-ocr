import streamlit as st
import openpyxl
from google import genai
from google.genai import types
import json
import io
import time

st.set_page_config(page_title="Նկարից Excel լրացնող գործիք", page_icon="📊")
st.title("📊 Փաստաթղթերի խորը վերլուծություն և Excel-ի լրացում")

api_key = st.sidebar.text_input("Մուտքագրեք Gemini API Key-ը", type="password")

excel_file = st.file_uploader("1. Վերբեռնեք ձեր Excel ֆորման (.xlsx)", type=["xlsx"])
doc_files = st.file_uploader(
    "2. Վերբեռնեք նկարներ, PDF-ներ կամ Excel ֆայլեր (միաժամանակյա վերլուծության համար)", 
    type=["png", "jpg", "jpeg", "pdf", "xlsx"],
    accept_multiple_files=True
)
header_row = st.number_input("Ո՞ր տողում են գտնվում սյունակների վերնագրերը", min_value=1, value=1)

extra_instructions = st.text_area(
    "Լրացուցիչ ցուցումներ AI-ին (ոչ պարտադիր)", 
    placeholder="Օրինակ՝ Կազմակերպությունների անունները համեմատիր ըստ ՀՎՀՀ-ի, ՍՊԸ-ն և LLC-ն համարիր նույնը, կամ գումարները գրիր առանց ԱԱՀ-ի..."
)

def extract_text_from_excel(file_obj):
    source_wb = openpyxl.load_workbook(file_obj, data_only=True)
    text_lines = []
    for sheet_name in source_wb.sheetnames:
        ws = source_wb[sheet_name]
        text_lines.append(f"--- Էջ՝ {sheet_name} ---")
        for row in ws.iter_rows(values_only=True):
            row_vals = [str(cell).strip() if cell is not None else "" for cell in row]
            if any(row_vals):
                text_lines.append(" | ".join(row_vals))
    return "\n".join(text_lines)

# Օգնող ֆունկցիա՝ շատ մեծ JSON-ը անվտանգ կարդալու համար
def safe_parse_json(raw_text):
    try:
        return json.loads(raw_text)
    except json.JSONDecodeError:
        # Եթե տեքստը շատ երկար լինելու պատճառով վերջում կտրվել է, վերականգնում ենք մինչև վերջին ամբողջական օբյեկտը
        last_brace = raw_text.rfind("}")
        if last_brace != -1:
            repaired = raw_text[:last_brace + 1] + "]"
            return json.loads(repaired)
        raise

if st.button("Խորը անալիզ անել և լրացնել Excel-ը"):
    if not api_key:
        st.error("Խնդրում ենք ձախ դաշտում մուտքագրել Gemini API Key-ը։")
    elif not excel_file or not doc_files:
        st.error("Խնդրում ենք վերբեռնել ձեր ֆորման և առնվազն մեկ փաստաթուղթ։")
    else:
        with st.spinner(f"AI-ը խորը վերլուծում է {len(doc_files)} փաստաթուղթ (մեծ ֆայլերի դեպքում կարող է տևել 30-60 վայրկյան)..."):
            try:
                wb = openpyxl.load_workbook(excel_file)
                sheet = wb.active
                
                headers = []
                for col in range(1, sheet.max_column + 1):
                    val = sheet.cell(row=header_row, column=col).value
                    headers.append(str(val).strip() if val else f"Column_{col}")

                client = genai.Client(api_key=api_key)

                request_contents = []
                for idx, current_file in enumerate(doc_files, start=1):
                    request_contents.append(f"=== Փաստաթուղթ #{idx} (Ֆայլի անունը՝ {current_file.name}) ===")
                    if current_file.name.lower().endswith(".xlsx"):
                        excel_text = extract_text_from_excel(current_file)
                        request_contents.append(excel_text)
                    else:
                        file_bytes = current_file.read()
                        mime_type = current_file.type
                        request_contents.append(
                            types.Part.from_bytes(data=file_bytes, mime_type=mime_type)
                        )

                prompt = f"""
                Դու բարձրակարգ հաշվապահական և տվյալների վերլուծության մասնագետ ես։ Քո խնդիրն է ԱՌԱՎԵԼԱԳՈՒՅՆ ՃՇԳՐՏՈՒԹՅԱՄԲ վերլուծել վերևում տրված բոլոր {len(doc_files)} փաստաթղթերը միասին և լրացնել հետևյալ սյունակները՝ {headers}։

                ԱՆՈՒՆՆԵՐԻ ԵՎ ՏՎՅԱԼՆԵՐԻ ԻԴԵՆՏԻՖԻԿԱՑՄԱՆ ԽԻՍՏ ԿԱՆՈՆՆԵՐ.
                1. **Անունների նույնականացում (Fuzzy Matching).** Տարբեր ֆայլերում նույն անձի կամ կազմակերպության անունը կարող է գրված լինել փոքր-ինչ տարբեր (օրինակ՝ հայերեն, ռուսերեն կամ լատինատառ տրանսլիտով, «ՍՊԸ»-ով կամ առանց դրա, բառերի տեղափոխությամբ, կամ փոքր տառասխալներով)։ Ուշադիր համեմատիր տառ առ տառ, ստուգիր նաև ՀՎՀՀ-ները, հեռախոսահամարները, հասցեները կամ կոդերը՝ համոզվելու համար, որ ճիշտ ես կապում նույն սուբյեկտին։
                2. **Առանց հորինելու.** Երբեք մի հորինիր տվյալներ։ Եթե անունը կամ թիվը դժվարընթեռնելի է, ուշադիր ուսումնասիրիր համատեքստը և մյուս փաստաթղթերը՝ ճիշտ տարբերակը հաստատելու համար։
                3. **Տվյալների միավորում.** Բոլոր ֆայլերից հավաքիր տվյալները և միավորիր մեկ ամբողջական տողի մեջ յուրաքանչյուր եզակի գրառման համար։
                4. **Կոմպակտ JSON.** Վերադարձրու կոմպակտ JSON՝ առանց ավելորդ բացատների և տողադարձերի, որպեսզի բոլոր տվյալները տեղավորվեն։
                5. **Օգտատիրոջ հատուկ ցուցումներ.** {extra_instructions if extra_instructions else "Չկան"}

                Վերադարձրու արդյունքը ՄԻԱՅՆ որպես JSON զանգված (list of objects), որտեղ յուրաքանչյուր օբյեկտի բանալիները (keys) ճշգրիտ համընկնում են տրված սյունակների անունների հետ՝ {headers}։ Եթե որևէ դաշտի տվյալ ընդհանրապես չկա, դիր դատարկ տող ""։
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
                        st.info(f"Օգտագործվեց բարձր ճշգրտության մոդել՝ {model_name}")
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

                st.success(f"Պատրաստ է։ Ավելացվեց {len(extracted_rows)} ճշգրտված տող։")
                st.download_button(
                    label="📥 Ներբեռնել լրացված Excel-ը",
                    data=output,
                    file_name="filled_form.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.error(f"Տեղի ունեցավ սխալ. {e}")
