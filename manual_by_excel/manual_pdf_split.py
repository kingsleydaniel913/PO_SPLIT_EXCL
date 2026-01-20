import pandas as pd

from pypdf import PdfReader, PdfWriter

def process_excel_data(df_raw):
    # Standard splitting logic using index-based selection
    page_ranges_split = df_raw.iloc[:, 2].str.split('-', expand=True)
    df_raw['StartPage'] = pd.to_numeric(page_ranges_split[0], errors='coerce').astype('Int64')
    df_raw['EndPage'] = pd.to_numeric(page_ranges_split[1], errors='coerce').astype('Int64')
    df_raw['FileName'] = df_raw.iloc[:, 3]
    return df_raw[['FileName', 'StartPage', 'EndPage']].dropna()

def split_pdf_with_debug(input_pdf_path, df):
    reader = PdfReader(input_pdf_path)
    total_pages = len(reader.pages)
    print(f"--- Loaded PDF: {input_pdf_path} ({total_pages} total pages) ---")

    for index, row in df.iterrows():
        # DEBUG: View exactly what data is being used for this row

        print(f"\n[Row {index}] Processing: {row['FileName']} | Range: {row['StartPage']}-{row['EndPage']}")
        
        try:
            writer = PdfWriter()
            # Convert 1-based Excel to 0-based Python indices
            start = int(row['StartPage']) - 1
            end = int(row['EndPage'])

            # Boundary Catch: Check if requested pages exist in PDF
            if start < 0 or end > total_pages:
                print(f"  ERROR: Range {row['StartPage']}-{row['EndPage']} exceeds PDF limits (1-{total_pages}). Skipping.")
                continue

            for page_num in range(start, end):
                writer.add_page(reader.pages[page_num])

            output_name = f"{row['FileName']}.pdf"
            with open(output_name, "wb") as f:
                writer.write(f)
            print(f"  SUCCESS: Created {output_name}")

        except Exception as e:
            # Catch: Any unexpected error (e.g., file permissions, corrupt page)
            print(f"  CRITICAL ERROR on Row {index}: {str(e)}")

# --- Execution ---
try:
    df_raw = pd.read_excel("data.xlsx", header=None)
    df_final = process_excel_data(df_raw)
    split_pdf_with_debug("exp.pdf", df_final)
except Exception as e:
    print(f"Failed to start: {e}")
