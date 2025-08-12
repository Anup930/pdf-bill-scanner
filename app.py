import os
import sys
import streamlit as st
import pdfplumber
import pytesseract
from pdf2image import convert_from_path
import tempfile
import google.generativeai as genai
import json
import pandas as pd
import re
import io

# -------- Clear history when app starts --------
EXCEL_FILE = "bill_data.xlsx"
if os.path.exists(EXCEL_FILE):
    os.remove(EXCEL_FILE)

# ----------- TESSERACT PATH SETUP -----------
if getattr(sys, 'frozen', False):
    base_path = sys._MEIPASS
    tesseract_path = os.path.join(base_path, "tesseract", "tesseract.exe")
else:
    tesseract_path = r"C:\\Program Files\\Tesseract-OCR\\tesseract.exe"

pytesseract.pytesseract.tesseract_cmd = tesseract_path

# ----------- CONFIG -----------
GEMINI_API_KEY = "AIzaSyCDEGN1ZXXVda9yhp2bHhpzT5yncr66CKY"
genai.configure(api_key=GEMINI_API_KEY)

MODEL_NAME = "models/gemini-2.5-pro"

DEFAULT_PROMPT = """
Scan this bill and tell me the following details in JSON format:
1. Name of the Vendor
2. Name of the company on which this bill has been raise
3. Nature of expense
4. Is TDS Applicable on this bill? If yes, what is the rate and amount?
5. Is GST under RCM applicable?
6. Is GST Input included in the bill?
7. Is the nature of IGST or CGST and SGST as per Place of Supply in GST correct?
8. What is the final amount payable to the vendor?
9. Are there any remarks mentioned in the bill?
Return ONLY valid JSON without any extra text, explanation, or markdown formatting.
"""

st.title("📄 PDF Bill Scanner + Gemini AI Data Extractor (Excel Auto-Save)")

# -------- Edit Prompt Option --------
edit_prompt = st.checkbox("✏️ Edit Extraction Prompt", value=False)
if edit_prompt:
    DATA_EXTRACTION_PROMPT = st.text_area(
        "Modify the extraction prompt as needed:",
        value=DEFAULT_PROMPT,
        height=200
    )
else:
    DATA_EXTRACTION_PROMPT = DEFAULT_PROMPT

uploaded_file = st.file_uploader("Upload your PDF bill", type=["pdf"])

if uploaded_file is not None:
    # -------- Save temp file --------
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_pdf:
        temp_pdf.write(uploaded_file.read())
        temp_pdf_path = temp_pdf.name

    extracted_text = ""

    # Step 1: Direct text extraction
    with pdfplumber.open(temp_pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                extracted_text += text + "\n"

    # Step 2: OCR fallback if no text found
    if not extracted_text.strip():
        st.warning("⚠ No text found in the PDF. Running OCR...")
        images = convert_from_path(temp_pdf_path)
        for img in images:
            text = pytesseract.image_to_string(img, lang="eng")
            extracted_text += text + "\n"

    if extracted_text.strip():
        st.subheader("📜 Extracted Text")
        st.text_area("Output", extracted_text, height=300)

        # -------- Manual extra fields --------
        st.subheader("📝 Additional Bill Details (Manual Entry)")
        bill_source = st.text_input("Bill Source")
        bill_given_by = st.text_input("Bill Given By")
        hod_approval = st.text_input("HOD for Approval")
        final_approval = st.text_input("Final Approval")

        st.subheader("🤖 Gemini AI Extracted Data")
        if st.button("Get Data from Gemini"):
            if not bill_source or not bill_given_by or not hod_approval or not final_approval:
                st.error("⚠ Please fill all manual fields before proceeding.")
            else:
                with st.spinner("Processing with Gemini..."):
                    try:
                        model = genai.GenerativeModel(MODEL_NAME)
                        response = model.generate_content(f"{DATA_EXTRACTION_PROMPT}\n\nBill Text:\n{extracted_text}")

                        # Get response text safely
                        if hasattr(response, "text") and response.text:
                            cleaned_text = response.text.strip()
                        elif hasattr(response, "candidates") and len(response.candidates) > 0:
                            cleaned_text = response.candidates[0].content.strip()
                        else:
                            cleaned_text = str(response)

                        # Extract only JSON part using regex
                        match = re.search(r"\{.*\}", cleaned_text, re.DOTALL)
                        if match:
                            cleaned_text = match.group(0)

                        try:
                            parsed_data = json.loads(cleaned_text)

                            # Add manual fields to JSON
                            parsed_data["Bill Source"] = bill_source
                            parsed_data["Bill Given By"] = bill_given_by
                            parsed_data["HOD Approval"] = hod_approval
                            parsed_data["Final Approval"] = final_approval

                            st.success("✅ Data Extracted from Gemini")

                            # Convert to DataFrame and display
                            df = pd.json_normalize(parsed_data)
                            st.dataframe(df)

                            # Append to Excel file for current session
                            if os.path.exists(EXCEL_FILE):
                                existing_df = pd.read_excel(EXCEL_FILE)
                                final_df = pd.concat([existing_df, df], ignore_index=True)
                            else:
                                final_df = df

                            final_df.to_excel(EXCEL_FILE, index=False)
                            st.info(f"💾 Data saved to '{EXCEL_FILE}'")

                            # Provide Excel download button
                            towrite = io.BytesIO()
                            final_df.to_excel(towrite, index=False, engine='openpyxl')
                            towrite.seek(0)
                            st.download_button(
                                label="Download Excel File",
                                data=towrite,
                                file_name=EXCEL_FILE,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )

                        except json.JSONDecodeError:
                            st.warning("⚠ Could not parse as JSON, showing raw output.")
                            st.text(cleaned_text)

                    except Exception as e:
                        st.error(f"Gemini API Error: {e}")

    else:
        st.error("No text could be extracted, even with OCR.")
