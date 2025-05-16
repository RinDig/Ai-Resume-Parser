# Resume Parsing Script

This Python script automates the process of extracting information from a collection of resumes. It reads resumes from a ZIP archive, supports PDF, DOCX, and TXT formats, extracts key details (name, email, phone number, location, and a guessed job title) using the OpenAI API, and then compiles this information into an Excel spreadsheet.


![image](https://github.com/user-attachments/assets/644164e8-5c69-459b-af41-b2bc98f7de74)

## Features

*   Processes resumes from a `.zip` file.
*   Supports multiple resume formats:
    *   `.pdf`
    *   `.docx`
    *   `.txt`
*   Utilizes OpenAI's GPT model (specifically `gpt-3.5-turbo`) to extract:
    *   Candidate Name
    *   Email Address
    *   Phone Number
    *   Location (City, State/Country)
    *   Most recent/prominent Job Title (guessed)
*   Outputs parsed data into an Excel file (`.xlsx`).
*   Handles common errors and unsupported file types gracefully.
*   Includes a preview of the extracted text in the output for quick reference.

## Prerequisites

*   Python 3.x
*   `pip` (Python package installer)

## Setup

1.  **Clone or Download the Script:**
    Ensure you have the `res_parser.py` script.

2.  **Install Dependencies:**
    Open your terminal or command prompt and run the following command to install the necessary Python libraries:
    ```bash
    pip install PyMuPDF python-docx pandas openai
    ```
    *(Note: `fitz` is the import name for `PyMuPDF`)*

3.  **Configure the Script (`res_parser.py`):**
    Open `res_parser.py` in a text editor and make the following changes:

    *   **OpenAI API Key:**
        Locate the line:
        ```python
        OPENAI_API_KEY = "your key here"
        ```
        Replace `"sk-proj-..."` with your **actual OpenAI API key**. If you don't have one, you'll need to sign up at [OpenAI](https://openai.com/).

    *   **ZIP File Path (Required):**
        Locate the line:
        ```python
        ZIP_FILE_PATH = "R-data.zip"
        ```
        Change `"R-data.zip"` to the name of your ZIP file containing the resumes. Place this ZIP file in the same directory as the script, or provide the full path.

    *   **Output Excel Path (Optional):**
        You can change the name of the output Excel file if desired:
        ```python
        OUTPUT_EXCEL_PATH = "parsed_resumes.xlsx"
        ```

    *   **Temporary Extraction Directory (Optional):**
        This setting is mentioned in the script but not actively used by default for individual file processing from the zip.
        ```python
        TEMP_EXTRACTION_DIR = "temp_resumes"
        ```

## How to Run

1.  Ensure you have completed the **Setup** steps, especially configuring the API key and `ZIP_FILE_PATH`.
2.  Place your ZIP file (e.g., `R-data.zip`) in the same directory as the `res_parser.py` script (or update `ZIP_FILE_PATH` accordingly).
3.  Open your terminal or command prompt.
4.  Navigate to the directory where the script is saved.
5.  Run the script using Python:
    ```bash
    python res_parser.py
    ```
6.  The script will process the files in the ZIP archive. You will see output in the console indicating the progress and any errors.
7.  Once completed, an Excel file (default: `parsed_resumes.xlsx`) will be created in the same directory, containing the extracted information.

## Input

*   A single `.zip` file containing resumes.
    *   Supported resume formats: `.pdf`, `.docx`, `.txt`.
    *   `.doc` files will be skipped, and a message will be printed suggesting conversion to DOCX or PDF.

## Output

*   An Excel file (e.g., `parsed_resumes.xlsx`) with the following columns:
    *   `filename`: The original name of the file within the ZIP archive.
    *   `name`: Extracted candidate name.
    *   `email`: Extracted email address.
    *   `phone`: Extracted phone number.
    *   `location`: Extracted location.
    *   `job_title_guess`: OpenAI's best guess for the most recent or prominent job title.
    *   `full_text_preview`: A short preview (first 200 characters) of the text extracted from the resume.

## Important Notes

*   **OpenAI API Costs:** Using the OpenAI API incurs costs based on token usage. Be mindful of the number and size of resumes you process. The script limits the text sent to the API to the first 4000 characters to help manage this.
*   **Extraction Accuracy:** The accuracy of the extracted information (name, email, phone, location, job title) depends on the resume's formatting and the OpenAI model's ability to interpret it. Some manual review of the output may be necessary.
*   **`.doc` Files:** The script currently skips `.doc` files due to the complexity of reliable text extraction without external dependencies like `antiword` or `textract`. It is recommended to convert `.doc` files to `.docx` or `.pdf` before processing.
*   **Error Handling:** The script includes basic error handling for file processing and API calls. Errors will be noted in the console output and in the `name` field of the Excel sheet for the affected file (e.g., "UNSUPPORTED TYPE", "EXTRACTION FAILED", "PROCESSING ERROR").

## Troubleshooting

*   **`Error: ZIP file not found`**: Ensure the `ZIP_FILE_PATH` in the script is correct and the ZIP file exists at that location.
*   **`WARNING: Please replace 'YOUR_OPENAI_API_KEY'`**: You have not replaced the placeholder API key with your actual OpenAI API key.
*   **OpenAI API Errors (e.g., authentication, rate limits)**:
    *   Double-check that your `OPENAI_API_KEY` is correct and active.
    *   You might be exceeding your API rate limits or quota. Check your OpenAI account dashboard.
*   **Low Extraction Quality:**
    *   The resume format might be complex or image-based (scanned PDF without OCR).
    *   Consider improving the prompt sent to the OpenAI API within the `parse_info_from_text` function for more targeted extraction if needed.
*   **Dependencies Not Found:** Make sure you have run `pip install -r requirements.txt` (if a `requirements.txt` is provided) or `pip install PyMuPDF python-docx pandas openai`.

