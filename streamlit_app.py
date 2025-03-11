import os
import time
import openai
import PyPDF2
import docx
import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from dotenv import load_dotenv

load_dotenv()

COMPANY_NAME = os.getenv("COMPANY_NAME")  

try:
    from api_key import api_key
except ImportError:
    api_key = ""

PRICE_INPUT_PER_MILLION = 0.075
PRICE_OUTPUT_PER_MILLION = 0.30
AVG_TOKENS_PER_REQUEST = 780.5
DEFAULT_PROCESS_TIME_PER_FILE = 1.62  

ROOT_DIR = "uploaded_files"
os.makedirs(ROOT_DIR, exist_ok=True)

st.title("📄 Administrative Document Classification")

company_name = st.text_input("🔑 What is your company working for?", type="password")
if company_name != COMPANY_NAME:
    st.warning("⚠️ Access denied! Please enter the correct company name.")
    st.stop()

if not api_key:
    api_key = st.text_input("🔑 Enter your OpenAI API Key", type="password")
    if not api_key:
        st.warning("⚠️ Please enter your OpenAI API key to proceed.")
        st.stop()

openai.api_key = api_key

uploaded_files = st.file_uploader("📂 Upload files", type=["pdf", "txt", "docx"], accept_multiple_files=True)

if uploaded_files:
    total_files = len(uploaded_files)
    st.write(f"📂 **Total files uploaded:** {total_files}")

    saved_file_paths = []
    for uploaded_file in uploaded_files:
        file_path = os.path.join(ROOT_DIR, uploaded_file.name)
        with open(file_path, "wb") as f:
            f.write(uploaded_file.read())
        saved_file_paths.append(file_path)

    batch_size = st.number_input("🔢 Choose number of files per batch", min_value=1, max_value=total_files, value=min(10, total_files), step=1)
    delay_time = st.number_input("⏳ Choose delay time per batch (seconds)", min_value=1, max_value=60, value=5, step=1)

    estimated_total_tokens = total_files * AVG_TOKENS_PER_REQUEST
    estimated_cost_input = (estimated_total_tokens / 1_000_000) * PRICE_INPUT_PER_MILLION
    estimated_cost_output = (estimated_total_tokens * 0.05 / 1_000_000) * PRICE_OUTPUT_PER_MILLION
    estimated_total_cost = estimated_cost_input + estimated_cost_output

    st.write(f"💰 **Estimated Total Cost**: ${estimated_total_cost:.4f}")

    estimated_time_full = total_files * DEFAULT_PROCESS_TIME_PER_FILE
    estimated_time_batches = ((total_files / batch_size) * delay_time) + estimated_time_full

    st.write(f"⏳ **Estimated Time (Full Process, No Delay)**: {estimated_time_full:.2f} sec (~{estimated_time_full/60:.2f} min)")
    st.write(f"⏳ **Estimated Time (Batch Processing, With Delay)**: {estimated_time_batches:.2f} sec (~{estimated_time_batches/60:.2f} min)")

    process_all = st.button("🔄 Process All Files Now")
    process_batches = st.button("⏳ Process in Batches")

    if "classification_results" not in st.session_state:
        st.session_state.classification_results = []

    def extract_text(file_path):
        file_extension = file_path.split(".")[-1].lower()
        text = ""
        pages = 1

        if file_extension == "pdf":
            with open(file_path, "rb") as file:
                reader = PyPDF2.PdfReader(file)
                pages = len(reader.pages)
                for page in reader.pages:
                    text += page.extract_text() + "\n"

        elif file_extension == "txt":
            with open(file_path, "r", encoding="utf-8") as file:
                text = file.read()
                pages = text.count("\n") // 30 + 1

        elif file_extension == "docx":
            doc = docx.Document(file_path)
            text = "\n".join([para.text for para in doc.paragraphs])
            pages = len(doc.paragraphs) // 30 + 1

        return text.strip(), pages

    def classify_document(text, file_name):
        prompt = f"""
        You are an AI that classifies administrative documents.
        Given the document text, classify it into:
        - Main Category (e.g., Report, Proposal, Decision)
        - Subcategory (e.g., Inspection Report, Investment Proposal, Research Evaluation). If multiple, separate with `/`
        - Domain/Industry (e.g., Government Administration, Science & Technology). If multiple, separate with `/`

        Respond strictly in CSV format with ',' as a delimiter:
        File name, Main category, Subcategory, Domain/ Industry, Pages

        Example:
        1_bc_ket_qua_thuc_hien_kl_than202408100849157_Signed.pdf, Report, Inspection Report, Government Administration, 4

        Here is the document text:
        {text[:2000]}
        """
        
        try:
            response = openai.chat.completions.create(
                model="gpt-4o-mini",
                messages=[
                    {"role": "system", "content": "You classify administrative documents into structured categories."},
                    {"role": "user", "content": prompt}
                ]
            )

            result = response.choices[0].message.content.strip()
            lines = result.split("\n")
            return lines[1] if len(lines) > 1 else result

        except openai.AuthenticationError:
            st.error("❌ Invalid API Key! Please enter a valid OpenAI API key.")
            st.stop()

        except openai.OpenAIError as e:
            st.error(f"⚠️ OpenAI API Error: {str(e)}")
            st.stop()

    if process_all:
        start_time = time.time()
        progress_bar = st.progress(0)

        with st.spinner("Processing all files... Please wait."):
            for index, file_path in enumerate(saved_file_paths):
                text, pages = extract_text(file_path)
                classification = classify_document(text, os.path.basename(file_path))

                classification_data = classification.split(", ")
                if len(classification_data) >= 4:
                    file_name, main_category, subcategory, domain = classification_data[:4]
                    st.session_state.classification_results.append([file_name, main_category, subcategory, domain, pages])

                progress_bar.progress((index + 1) / total_files)

        total_time = round(time.time() - start_time, 2)
        st.success(f"✅ **All files processed in {total_time} sec (~{total_time/60:.2f} min)**")

    if process_batches:
        start_time = time.time()
        file_batches = [saved_file_paths[i:i + batch_size] for i in range(0, total_files, batch_size)]

        for batch_index, batch in enumerate(file_batches):
            batch_start_time = time.time()
            st.write(f"⚙️ Processing batch {batch_index + 1} of {len(file_batches)}...")

            for file_path in batch:
                text, pages = extract_text(file_path)
                classification = classify_document(text, os.path.basename(file_path))

                classification_data = classification.split(", ")
                if len(classification_data) >= 4:
                    file_name, main_category, subcategory, domain = classification_data[:4]
                    st.session_state.classification_results.append([file_name, main_category, subcategory, domain, pages])

            batch_time = round(time.time() - batch_start_time, 2)
            st.success(f"✅ **Batch {batch_index + 1} processed in {batch_time} sec**")

            if batch_index < len(file_batches) - 1:
                st.write(f"⏳ Waiting {delay_time} sec before next batch...")
                time.sleep(delay_time)

        total_time = round(time.time() - start_time, 2)
        st.success(f"✅ **All batches processed in {total_time} sec (~{total_time/60:.2f} min)**")

    if st.session_state.classification_results:
        df = pd.DataFrame(st.session_state.classification_results, columns=["File name", "Main category", "Subcategory", "Domain/ Industry", "Pages"])
        st.write(df)

        chart_type = st.selectbox("📊 Choose Chart Type", ["Bar", "Pie"])
        metric = st.selectbox("📌 Choose Metric", ["Main category", "Subcategory", "Domain/ Industry"])

        if chart_type == "Bar":
            plt.figure(figsize=(10, 5))
            sns.countplot(data=df, x=metric, order=df[metric].value_counts().index)
            plt.xticks(rotation=45)
            st.pyplot(plt)

        elif chart_type == "Pie":
            df[metric].value_counts().plot.pie(autopct="%1.1f%%", figsize=(6, 6))
            st.pyplot(plt)

        csv_file = "document_classification_results.csv"
        df.to_csv(csv_file, index=False)
        st.download_button("📥 Download CSV", data=open(csv_file, "rb"), file_name=csv_file, mime="text/csv")
