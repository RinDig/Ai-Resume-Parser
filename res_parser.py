import zipfile
import os
import fitz  # PyMuPDF
from docx import Document as DocxDocument # python-docx
import pandas as pd
import openai # Added for OpenAI API
import json # Added for parsing JSON response
from io import BytesIO # To read files from zip without extracting all to disk

# --- CONFIGURATION ---
ZIP_FILE_PATH = "R-data.zip"  # <--- PUT YOUR ZIP FILE NAME HERE
OUTPUT_EXCEL_PATH = "parsed_resumes.xlsx"
TEMP_EXTRACTION_DIR = "temp_resumes" # Optional: if you want to extract all first

# --- OpenAI API Key ---
# IMPORTANT: Replace "YOUR_OPENAI_API_KEY" with your actual OpenAI API key
OPENAI_API_KEY = "YOUR_OPENAI_API_KEY"
if OPENAI_API_KEY == "YOUR_OPENAI_API_KEY":
    print("WARNING: Please replace 'YOUR_OPENAI_API_KEY' with your actual OpenAI API key in the script.")
openai.api_key = OPENAI_API_KEY

# --- HELPER FUNCTIONS ---

def extract_text_from_pdf(file_content_bytes):
    """Extracts text from PDF file content (bytes)."""
    text = ""
    try:
        with fitz.open(stream=file_content_bytes, filetype="pdf") as doc:
            for page in doc:
                text += page.get_text()
    except Exception as e:
        print(f"Error reading PDF: {e}")
    return text

def extract_text_from_docx(file_content_bytes):
    """Extracts text from DOCX file content (bytes)."""
    text = ""
    try:
        doc = DocxDocument(BytesIO(file_content_bytes))
        for para in doc.paragraphs:
            text += para.text + "\\n"
    except Exception as e:
        print(f"Error reading DOCX: {e}")
    return text

def extract_text_from_doc(file_content_bytes, filename):
    """
    Extracts text from DOC files. This is more complex and less reliable.
    It often requires external tools like 'antiword' (Linux/macOS) or 'textract'.
    For simplicity, this example will skip .doc or suggest manual conversion.
    You could integrate 'textract' if needed: pip install textract
    import textract
    text = textract.process(filepath_or_bytes).decode('utf-8')
    """
    print(f"Skipping .doc file: {filename}. Please convert to DOCX or PDF for better results.")
    return ""

def extract_text_from_txt(file_content_bytes):
    """Extracts text from TXT file content (bytes)."""
    try:
        return file_content_bytes.decode('utf-8', errors='ignore')
    except Exception as e:
        print(f"Error reading TXT: {e}")
    return ""


def parse_info_from_text(text, filename):
    """
    Parses extracted text to find name, email, phone, location, and job title
    using the OpenAI API.
    """
    data = {
        "filename": filename,
        "name": None,
        "email": None,
        "phone": None,
        "location": None,
        "job_title_guess": None,
        "full_text_preview": text[:200].replace("\\n", " ") + "..."
    }

    if not text.strip():
        print(f"  No text to parse for {filename}.")
        return data

    if openai.api_key == "YOUR_OPENAI_API_KEY":
        print("  OpenAI API key not set. Skipping parsing.")
        return data

    prompt = f"""
    Extract the following information from the resume text provided below:
    - Name of the candidate
    - Email address
    - Phone number
    - Location (City, State/Country if available)
    - Most recent or prominent job title (job_title_guess)

    Return the information in a JSON format with the following keys: "name", "email", "phone", "location", "job_title_guess".
    If a piece of information is not found, use a null value for that key.

    Resume Text:
    ---
    {text[:4000]} # Limiting text length to manage token usage, adjust as needed
    ---

    JSON Output:
    """

    try:
        # For openai library v1.0.0 and later
        client = openai.OpenAI(api_key=OPENAI_API_KEY)
        response = client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[
                {"role": "system", "content": "You are a helpful assistant that extracts information from resumes and returns it in JSON format."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.2,
            max_tokens=300 
        )
        
        extracted_str = response.choices[0].message.content.strip()
        
        # The response might be inside a ```json ... ``` block
        if extracted_str.startswith("```json"):
            extracted_str = extracted_str[7:]
            if extracted_str.endswith("```"):
                extracted_str = extracted_str[:-3]
        
        extracted_json = json.loads(extracted_str)

        data["name"] = extracted_json.get("name")
        data["email"] = extracted_json.get("email")
        data["phone"] = extracted_json.get("phone")
        data["location"] = extracted_json.get("location")
        data["job_title_guess"] = extracted_json.get("job_title_guess")

    except json.JSONDecodeError as e:
        print(f"  Error parsing JSON response from OpenAI for {filename}: {e}")
        print(f"  Raw response: {extracted_str}")
    except Exception as e:
        print(f"  Error calling OpenAI API for {filename}: {e}")
        # If there's an API error, we keep the default data with None values for parsed fields
        # but still include the filename and preview.

    return data

# --- MAIN SCRIPT ---
def main():
    all_resume_data = []

    if not os.path.exists(ZIP_FILE_PATH):
        print(f"Error: ZIP file not found at {ZIP_FILE_PATH}")
        return

    print(f"Processing resumes from: {ZIP_FILE_PATH}")

    with zipfile.ZipFile(ZIP_FILE_PATH, 'r') as zf:
        file_list = zf.namelist()
        print(f"Found {len(file_list)} files in the zip.")

        for i, filename_in_zip in enumerate(file_list):
            print(f"\\nProcessing file {i+1}/{len(file_list)}: {filename_in_zip} ...")
            
            # Skip macOS metadata files or directories
            if filename_in_zip.startswith('__MACOSX/') or filename_in_zip.endswith('/'):
                print("Skipping metadata file or directory.")
                continue

            try:
                file_content_bytes = zf.read(filename_in_zip)
                text = ""
                file_ext = os.path.splitext(filename_in_zip)[1].lower()

                if file_ext == ".pdf":
                    text = extract_text_from_pdf(file_content_bytes)
                elif file_ext == ".docx":
                    text = extract_text_from_docx(file_content_bytes)
                elif file_ext == ".doc":
                    text = extract_text_from_doc(file_content_bytes, filename_in_zip) # Will likely skip
                elif file_ext == ".txt":
                    text = extract_text_from_txt(file_content_bytes)
                else:
                    print(f"Unsupported file type: {file_ext} for {filename_in_zip}")
                    all_resume_data.append({
                        "filename": filename_in_zip, "name": "UNSUPPORTED TYPE",
                        "email": None, "phone": None, "location": None,
                        "job_title_guess": None, "full_text_preview": "Unsupported file type"
                    })
                    continue

                if text.strip():
                    parsed_info = parse_info_from_text(text, filename_in_zip)
                    all_resume_data.append(parsed_info)
                    print(f"  Extracted: Name='{parsed_info['name']}', Email='{parsed_info['email']}'")
                else:
                    print(f"  Could not extract text or text was empty for {filename_in_zip}.")
                    all_resume_data.append({
                        "filename": filename_in_zip, "name": "EXTRACTION FAILED",
                        "email": None, "phone": None, "location": None,
                        "job_title_guess": None, "full_text_preview": "Failed to extract text"
                    })

            except Exception as e:
                print(f"  Error processing file {filename_in_zip}: {e}")
                all_resume_data.append({
                    "filename": filename_in_zip, "name": "PROCESSING ERROR",
                    "email": None, "phone": None, "location": None,
                    "job_title_guess": None, "full_text_preview": f"Error: {str(e)[:100]}"
                })

    # Create DataFrame and save to Excel
    if all_resume_data:
        df = pd.DataFrame(all_resume_data)
        df.to_excel(OUTPUT_EXCEL_PATH, index=False)
        print(f"\\nSuccessfully parsed {len(all_resume_data)} resumes.")
        print(f"Results saved to: {OUTPUT_EXCEL_PATH}")
    else:
        print("\\nNo data was parsed.")

if __name__ == "__main__":
    main()