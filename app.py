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
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Point pytesseract to your local Windows installation path (fallback for local scans if needed)
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# Page Layout Configuration
st.set_page_config(
    page_title="Autonomous AI Business Agent (Enterprise Edition)",
    page_icon="⚡",
    layout="wide"
)

st.title("⚡ Autonomous AI Business Agent (Enterprise Multi-File Suite)")
st.markdown("Upload **multiple PDFs, PNGs, or JPGs** at once. Native vision processing guarantees every detail is extracted into a gorgeous, professional multi-tab Excel workbook.")

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
    """Processes either PDF text or sends image directly to Groq Vision for precise document reading."""
    client = Groq(api_key=api_key)
    file_extension = uploaded_file.name.split('.')[-1].lower()
    
    system_prompt = (
        "You are an expert financial auditor and data extraction AI. "
        "Analyze the document thoroughly. Read every single word, number, date, vendor name, client name, invoice number, "
        "payment method, subtotal, tax, total amount, and line item. "
        "Do NOT return null for prices, dates, or totals if they are present in the document. "
        "Return a clean JSON object containing exact keys: "
        "document_type, date, invoice_number, vendor_name, vendor_address, vendor_email, "
        "client_name, client_address, payment_method, currency, subtotal, tax_amount, total_amount, summary, "
        "and an array called line_items (each containing item_description, quantity, unit_price, total)."
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
                        {"type": "text", "text": system_prompt + "\nRead every word precisely from this invoice/document image and extract all details."},
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
            with st.spinner("Extracting precise details with Groq Vision & Llama engines..."):
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

        # Generate Single Multi-Tab Excel Workbook
        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
            summary_rows = []
            
            for fname, fcontent in st.session_state["processed_data"].items():
                if "error" not in fcontent:
                    # 1. Summary row for master sheet
                    flat_dict = {k: v for k, v in fcontent.items() if not isinstance(v, (list, dict))}
                    flat_dict["file_name"] = fname
                    summary_rows.append(flat_dict)
                    
                    # Determine clean sheet name based on document type or file name
                    doc_type = str(fcontent.get("document_type", "Document")).strip().title()
                    base_name = fname.rsplit('.', 1)[0]
                    raw_sheet_name = f"{doc_type}_{base_name}"
                    # Excel sheet name restrictions: max 31 chars, no invalid characters like / \ ? * : [ ]
                    safe_sheet_name = "".join(c for c in raw_sheet_name if c not in '/\\?*:[ ]')[:31]
                    
                    # 2. Build structured layout inside individual sheet
                    meta_rows = [
                        {"Field Name": "Document File Name", "Extracted Value": fname},
                        {"Field Name": "Document Type", "Extracted Value": fcontent.get("document_type", "N/A")},
                        {"Field Name": "Invoice Number", "Extracted Value": fcontent.get("invoice_number", "N/A")},
                        {"Field Name": "Date", "Extracted Value": fcontent.get("date", "N/A")},
                        {"Field Name": "Vendor Name", "Extracted Value": fcontent.get("vendor_name", "N/A")},
                        {"Field Name": "Vendor Address", "Extracted Value": fcontent.get("vendor_address", "N/A")},
                        {"Field Name": "Client Name", "Extracted Value": fcontent.get("client_name", "N/A")},
                        {"Field Name": "Payment Method", "Extracted Value": fcontent.get("payment_method", "N/A")},
                        {"Field Name": "Currency", "Extracted Value": fcontent.get("currency", "USD")},
                        {"Field Name": "Total Amount", "Extracted Value": fcontent.get("total_amount", "N/A")},
                        {"Field Name": "Summary / Notes", "Extracted Value": fcontent.get("summary", "N/A")},
                        {"Field Name": "", "Extracted Value": ""}, # Spacer row
                        {"Field Name": "--- LINE ITEMS BREAKDOWN ---", "Extracted Value": ""}
                    ]
                    df_meta = pd.DataFrame(meta_rows)
                    df_meta.to_excel(writer, sheet_name=safe_sheet_name, index=False)
                    
                    # Append line items table underneath metadata
                    line_items = fcontent.get("line_items", [])
                    if line_items and isinstance(line_items, list):
                        df_items = pd.DataFrame(line_items)
                    else:
                        df_items = pd.DataFrame([{"item_description": "No line items parsed", "quantity": "", "unit_price": "", "total": ""}])
                    
                    # Write line items starting below metadata
                    df_items.to_excel(writer, sheet_name=safe_sheet_name, index=False, startrow=len(meta_rows) + 1)
                else:
                    summary_rows.append({"file_name": fname, "error": fcontent["error"]})
            
            # Write master Summary tab first
            if summary_rows:
                df_summary = pd.DataFrame(summary_rows)
                df_summary.to_excel(writer, sheet_name="Master_Summary", index=False)

        # Post-process openpyxl workbook for professional executive styling & gorgeous header colors
        excel_buffer.seek(0)
        from openpyxl import load_workbook
        wb = load_workbook(excel_buffer)
        
        header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        section_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
        header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        section_font = Font(name="Calibri", size=11, bold=True, color="1F4E78")
        regular_font = Font(name="Calibri", size=11)
        thin_border = Border(
            left=Side(style='thin', color='D9D9D9'),
            right=Side(style='thin', color='D9D9D9'),
            top=Side(style='thin', color='D9D9D9'),
            bottom=Side(style='thin', color='D9D9D9')
        )
        alignment_left = Alignment(horizontal='left', vertical='center', wrap_text=True)

        for sheet in wb.sheetnames:
            ws = wb[sheet]
            if sheet == "Master_Summary":
                # Style Master Summary Header with bold dark blue background
                for col_num in range(1, ws.max_column + 1):
                    cell = ws.cell(row=1, column=col_num)
                    cell.fill = header_fill
                    cell.font = header_font
                    cell.alignment = Alignment(horizontal='center', vertical='center')
            else:
                # Style individual file sheet metadata and tables with bold headers
                for row in range(1, ws.max_row + 1):
                    val_a = ws.cell(row=row, column=1).value
                    if val_a == "Field Name":
                        # Metadata table header
                        for col_num in range(1, ws.max_column + 1):
                            cell = ws.cell(row=row, column=col_num)
                            cell.fill = header_fill
                            cell.font = header_font
                            cell.alignment = Alignment(horizontal='center', vertical='center')
                    elif val_a == "--- LINE ITEMS BREAKDOWN ---":
                        ws.cell(row=row, column=1).fill = section_fill
                        ws.cell(row=row, column=1).font = section_font
                    elif row > 1 and ws.cell(row=row-1, column=1).value == "--- LINE ITEMS BREAKDOWN ---":
                        # Line items table header row
                        for col_num in range(1, ws.max_column + 1):
                            cell = ws.cell(row=row, column=col_num)
                            cell.fill = header_fill
                            cell.font = header_font
                            cell.alignment = Alignment(horizontal='center', vertical='center')
                    else:
                        for col_num in range(1, ws.max_column + 1):
                            cell = ws.cell(row=row, column=col_num)
                            cell.font = regular_font
                            cell.border = thin_border
                            cell.alignment = alignment_left

            # Auto-adjust column widths cleanly with padding and wrap text enabled
            for col in ws.columns:
                max_len = 0
                col_letter = get_column_letter(col[0].column)
                for cell in col:
                    if cell.value:
                        max_len = max(max_len, len(str(cell.value)))
                ws.column_dimensions[col_letter].width = max(max_len + 5, 24)

        final_excel_buffer = io.BytesIO()
        wb.save(final_excel_buffer)
        final_excel_buffer.seek(0)

        st.markdown("---")
        st.subheader("📥 Export Master Excel Workbook")
        st.download_button(
            label="📊 Download Professional Multi-Tab Excel Report (.xlsx)",
            data=final_excel_buffer,
            file_name="Autonomous_Business_Agent_Professional_Report.xlsx",
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
