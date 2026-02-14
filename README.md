# Gemini-Powered Financial Data OCR Automation

## Overview
This workflow automates the digitization of historical financial data from PDF scans into structured Excel files. It leverages **Google Gemini 3.0 Pro (Thinking Mode)** to intelligently map year sets and extract granular data, followed by a **Python** script for final conversion.

## Prerequisites

### 1. Environment & Tools
* **Python**: Ensure Python is installed on your system.
* **Required Libraries**: Install via terminal/command prompt:
    ```bash
    pip install pandas openpyxl
    ```
* **AI Model**: Subscription to **Gemini Advanced** (using **Gemini 3.0 Pro with Thinking Mode** is highly recommended).

### 2. File Assets
* **`Step1_Prompt_Mapping.txt`**: System prompt for file mapping and year set optimization.
* **`Step2_Prompt_Extraction.txt`**: System prompt for data extraction (JSON generation).
* **`Step3_Python_Create Excel.py`**: Script to convert JSON output to Excel.

---

## Workflow Steps

### Step 1. Mapping (Preparation & Selection)
**Goal:** Capture images and utilize Gemini to determine the optimal "Year Set" based on scan quality.

1.  **Image Capture**
    * Capture **IS (Income Statement)**, **BS (Balance Sheet)**, and **WC (Working Capital)** tables for every Company-Year.
    * *Tip:* Include the **Cover Page** in the capture to easily identify the year later.
    * *Note:* You may proactively skip years that are visibly illegible or if a clear copy for that year has already been secured.
2.  **Renaming (Standardization)**
    * Rename captured images to match the original PDF filename format.
    * **Bulk Renaming Shortcut:** Select all images for a specific book/year $\rightarrow$ Press `F2` $\rightarrow$ Paste the filename (e.g., `Kayser__Julius____Company_22372_1946`) $\rightarrow$ Press `Ctrl + Enter`.
    * **Result:**
        * `Kayser__Julius____Company_22372_1946 (1).png`
        * `Kayser__Julius____Company_22372_1946 (2).png`
3.  **Gemini Processing**
    * Input the content of **`Step1.txt`** into the Gemini chat.
    * **Upload Protocol:** Upload images in **batches of 10**.
        1.  Upload 10 images $\rightarrow$ Press Enter $\rightarrow$ Wait for Gemini to request the next batch.
        2.  Repeat until all files are uploaded.
        3.  After the final batch, type **"Done"**.
    * Gemini will analyze the batch and output the **Optimal Year Set**.

### Step 2. Extracting (Data Extraction)
**Goal:** Extract financial data from the selected images into JSON format.

1.  **Initiate Extraction**
    * In the **same chat session** used for Step 1, input the content of **`Step2.txt`**.
2.  **Segmentation Strategy**
    * If the dataset contains many years or rows, Gemini may hit token limits. It will suggest **segmenting the task**.
    * Type **"Yes"** to proceed with segmentation.
3.  **Output & Saving**
    * Copy the generated JSON code block.
    * Paste it into a local text file named **`output_gemini.txt`** and save.
4.  **Handling Token Limits (Truncation)**
    * **Normal Segmentation:** If segmentation is active, run **Step 3 (Excel Creation)** for the current part, then return to the chat and type **"Yes"** to generate the next part.
    * **Forced Cut-off:** If the JSON ends abruptly (missing the closing `}` bracket), ask Gemini to *"continue generating the rest"* and manually merge the text in your editor.

### Step 3. Create Excel (Conversion)
**Goal:** Convert the extracted JSON text into an Excel workbook.

1.  **Execution**
    * Ensure `output_gemini.txt` contains the JSON data.
    * Run the Python script:
    ```bash
    python "Step3_Python_Create Excel.py"
    ```
2.  **Result**
    * The generated Excel files will be saved in the **`output_excel`** folder.

### Step 4. Verification
**Goal:** Ensure data integrity.

1.  Cross-reference the generated Excel files with the original PDF images to verify accuracy.
