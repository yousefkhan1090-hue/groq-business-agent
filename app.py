import streamlit as st
from pypdf import PdfReader
from PIL import Image
import pytesseract
from groq import Groq
import json
import pandas as pd
import io
import base64
from streamlit_mic_recorder import mic_recorder

# Point pytesseract to your local Windows installation path (fallback for local scans if needed)
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# Page Layout Configuration
st.set_page_config(
    page_title="Autonomous AI Business Agent (Enterprise Edition)",
    page_icon="⚡",
    layout="wide"
)

st.title("⚡ Autonomous AI Business Agent (Enterprise Multi-File Suite)")
st.markdown("Upload **multiple PDFs, PNGs, or JPGs** at once. Native vision processing guarantees every price, quantity, and line item is read accurately.")

# Sidebar Configuration with secure cloud secrets support
st.sidebar.header("Configuration")
if "GROQ_API_KEY" in st.secrets:
    api_key = st.secrets["GROQ_API_KEY"]
    st.sidebar.success("API Key loaded securely from Cloud Secrets!")
else:
    api_key = st.sidebar.text_input(
        "Enter Groq API Key (Free):", 
        value="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", 
        type="password"
    )
st.sidebar.markdown("[Get a free Groq API key here](https://console.groq.com/)")

def process_single_file(uploaded_file, api_key):
    """Processes either PDF text or sends image directly to Groq Vision for perfect table reading."""
    client = Groq(api_key=api_key)
    file_extension = uploaded_file.name.split('.')[-1].lower()
    
    system_prompt = (
        "You are an expert business analyst and data extraction AI. "
        "Analyze the document thoroughly. Read every single word, number, line item, unit price, tax value, and total amount. "
        "Do NOT return null for prices or totals if they are present in the document. "
        "Return a clean JSON object containing keys like: document_type, date, vendor_name, client_name, total_amount, "
        "currency, summary, and an array called line_items (each containing item_description, quantity, unit_price, total)."
    )

    if file_extension == 'pdf':
        pdf_reader = PdfReader(uploaded_file)
        text = ""
        for page in pdf_reader.pages:
            extracted = page.extract_text()
            if extracted:
                text += extracted + "\n"
                
        response = client.chat.completions.create(
            model="openai/gpt-oss-120b",
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": f"Document Text:\n{text}"}
            ],
            response_format={"type": "json_object"}
        )
        return json.loads(response.choices[0].message.content)
        
    elif file_extension in ['png', 'jpg', 'jpeg']:
        image_bytes = uploaded_file.read()
        base64_image = base64.b64encode(image_bytes).decode('utf-8')
        
        response = client.chat.completions.create(
            model="qwen/qwen3.8-27b",
            messages=[
                {
                    "role": "user",
                    "content": [
                        {"type": "text", "text": system_prompt + "\nExtract all table rows, unit prices, and totals accurately from this invoice image."},
                        {"type": "image_url", "image_url": {"url": f"data:image/{file_extension};base64,{base64_image}"}}
                    ]
                }
            ],
            response_format={"type": "json_object"}
        )
        return json.loads(response.choices[0].message.content)

def transcribe_audio_with_groq(audio_bytes, api_key):
    """Transcribes user voice notes using Groq's Whisper model."""
    client = Groq(api_key=api_key)
    try:
        audio_file = io.BytesIO(audio_bytes)
        audio_file.name = "audio.wav"
        transcript = client.audio.transcriptions.create(
            model="whisper-large-v3",
            file=audio_file,
            response_format="text"
        )
        return transcript
    except Exception as e:
        return f"Transcription error: {e}"

def answer_question(document_texts, question, api_key):
    """Answers queries across processed documents."""
    client = Groq(api_key=api_key)
    context_str = "\n\n--- NEXT DOCUMENT ---\n\n".join([f"File: {name}\nContent/Data: {json.dumps(data, indent=2)}" for name, data in document_texts.items()])
    
    response = client.chat.completions.create(
        model="openai/gpt-oss-120b",
        messages=[
            {
                "role": "system",
                "content": "You are an AI business assistant answering questions accurately based strictly on the uploaded document data."
            },
            {
                "role": "user",
                "content": f"Documents Data:\n{context_str}\n\nQuestion: {question}"
            }
        ]
    )
    return response.choices[0].message.content

# --- Main Application UI ---
uploaded_files = st.file_uploader(
    "📂 Drag and drop multiple business documents (PDFs, PNGs, JPGs) here", 
    type=["pdf", "png", "jpg", "jpeg"],
    accept_multiple_files=True
)

if uploaded_files:
    st.success(f"📄 Loaded {len(uploaded_files)} file(s) successfully!")
    
    if "processed_data" not in st.session_state:
        st.session_state["processed_data"] = {}

    # Batch Processing Button
    if st.button("🚀 Process All Documents"):
        if not api_key:
            st.error("Please enter a valid Groq API Key.")
        else:
            with st.spinner("Analyzing documents with Groq Vision & Llama engines..."):
                results = {}
                for file in uploaded_files:
                    file.seek(0)
                    try:
                        extracted_json = process_single_file(file, api_key)
                        results[file.name] = extracted_json
                    except Exception as e:
                        results[file.name] = {"error": str(e)}
                st.session_state["processed_data"] = results
                st.success("Batch Processing Complete!")

    # Display Results & Multi-Tab Excel Generator
    if st.session_state["processed_data"]:
        st.subheader("📊 Extracted Structured Data per Document")
        
        tabs = st.tabs([file.name for file in uploaded_files])
        for idx, file in enumerate(uploaded_files):
            with tabs[idx]:
                file_data = st.session_state["processed_data"].get(file.name, {})
                st.json(file_data)

        # Generate Multi-Tab Excel Workbook
        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
            summary_rows = []
            for fname, fcontent in st.session_state["processed_data"].items():
                if "error" not in fcontent:
                    flat_dict = {k: v for k, v in fcontent.items() if not isinstance(v, (list, dict))}
                    flat_dict["file_name"] = fname
                    summary_rows.append(flat_dict)
                    
                    df_single = pd.json_normalize(fcontent)
                    safe_sheet_name = fname[:31].replace('/', '_').replace('\\', '_')
                    df_single.to_excel(writer, sheet_name=safe_sheet_name, index=False)
                else:
                    summary_rows.append({"file_name": fname, "error": fcontent["error"]})
            
            if summary_rows:
                df_summary = pd.DataFrame(summary_rows)
                df_summary.to_excel(writer, sheet_name="Master_Summary", index=False)
                
        excel_buffer.seek(0)

        st.markdown("---")
        st.subheader("📥 Export Master Excel Workbook")
        st.download_button(
            label="📊 Download Multi-Tab Excel Report (.xlsx)",
            data=excel_buffer,
            file_name="Autonomous_Business_Agent_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.divider()

        # Voice & Text Q&A Section
        st.subheader("💬 Chat or Speak with Your Documents")
        
        col1, col2 = st.columns([3, 1])
        with col2:
            st.markdown("**🎙️ Voice Input:**")
            audio_info = mic_recorder(start_prompt="Record Question", stop_prompt="Stop Recording", key='mic')

        user_question = ""
        if audio_info and 'bytes' in audio_info:
            with st.spinner("Transcribing voice via Groq Whisper..."):
                transcribed_text = transcribe_audio_with_groq(audio_info['bytes'], api_key)
                if transcribed_text and not transcribed_text.startswith("Transcription error"):
                    user_question = transcribed_text
                    st.info(f"**Transcribed Voice:** {user_question}")

        text_question = col1.text_input("Or type your question here:", value=user_question)
        final_query = text_question if text_question else user_question

        if final_query:
            if not api_key:
                st.error("Please enter your API key to use the Q&A feature.")
            else:
                with st.spinner("Searching documents..."):
                    answer = answer_question(st.session_state["processed_data"], final_query, api_key)
                    st.markdown(f"**Answer:** {answer}")