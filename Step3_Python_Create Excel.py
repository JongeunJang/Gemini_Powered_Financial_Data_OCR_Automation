import pandas as pd
import json
import io
import os
import re

def generate_organized_excels_smart(file_path, folder_name='output_excels'):
    # 1. Create output folder
    if not os.path.exists(folder_name):
        os.makedirs(folder_name)
        print(f"📂 Output folder verified: {folder_name}")

    # 2. Read the text file
    if not os.path.exists(file_path):
        print(f"❌ File not found: {file_path}")
        return

    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            raw_content = f.read().strip()

        # Pre-process JSON parsing
        clean_content = raw_content.replace('```json', '').replace('```', '').strip()

        # Attempt JSON parsing (handling potential syntax errors)
        try:
            data = json.loads(clean_content)
        except json.JSONDecodeError:
            # Auto-correct missing commas
            fixed_content = re.sub(r'(?<=[^\\]")\s+(?=")', ',\n', clean_content)
            try:
                data = json.loads(fixed_content)
            except:
                # Worst case: Attempt to process as a raw text chunk even if not valid JSON
                data = {"merged_data": clean_content}

    except Exception as e:
        print(f"❌ File read error: {e}")
        return

    # 3. Merge all data into one (Create Master DataFrame)
    all_dfs = []

    for key, text_data in data.items():
        try:
            # Logic to detect the header row
            lines = str(text_data).strip().split('\n')
            start_index = -1
            for i, line in enumerate(lines):
                if 'pdffile' in line and '\t' in line:
                    start_index = i
                    break

            if start_index != -1:
                clean_text = '\n'.join(lines[start_index:])
                df_chunk = pd.read_csv(io.StringIO(clean_text), sep='\t')
                all_dfs.append(df_chunk)
        except Exception as e:
            print(f"⚠️ Warning during data merge ({key}): {e}")

    if not all_dfs:
        print("❌ No data to process.")
        return

    # Merge all data chunks into a single DataFrame
    master_df = pd.concat(all_dfs, ignore_index=True)

    # 4. Split by 'pdffile' column and save as Excel (Core Logic)
    if 'pdffile' not in master_df.columns:
        print("❌ The 'pdffile' column is missing. Cannot split data.")
        return

    unique_files = master_df['pdffile'].unique()
    print(f"🔍 Found {len(unique_files)} unique files. Starting split and save process...")

    success_count = 0
    for pdf_filename in unique_files:
        try:
            # Filter data for the specific file
            file_df = master_df[master_df['pdffile'] == pdf_filename]

            # Clean filename (remove .pdf, etc.)
            base_name = str(pdf_filename).replace('.pdf', '').replace('.png', '').strip()
            # Remove/replace special characters invalid in Excel filenames
            base_name = re.sub(r'[\\/*?:"<>|]', "_", base_name)

            excel_filename = base_name + ".xlsx"
            excel_path = os.path.join(folder_name, excel_filename)

            # Save to Excel
            with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
                categories = {
                    'Income Statement': 'Income Statement',
                    'Balance Sheet': 'Balance Sheet',
                    'Working Capital': 'Working Capital'
                }

                found_any_sheet = False
                for cat_key, sheet_name in categories.items():
                    # Filter data by category
                    sheet_df = file_df[file_df['table'].str.contains(cat_key, case=False, na=False)]

                    if not sheet_df.empty:
                        sheet_df.to_excel(writer, sheet_name=sheet_name, index=False)
                        found_any_sheet = True

                if found_any_sheet:
                    print(f"✅ Successfully created: {excel_filename}")
                    success_count += 1
                else:
                    print(f"⚠️ Skipped due to insufficient data: {excel_filename}")

        except Exception as e:
            print(f"❌ Failed to save '{pdf_filename}': {e}")

    print(f"\n🎉 Successfully created {success_count} Excel files!")

if __name__ == "__main__":
    generate_organized_excels_smart('output_gemini.txt')