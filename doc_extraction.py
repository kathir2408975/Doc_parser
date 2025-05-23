import os
import fitz  # PyMuPDF
from pdf2image import convert_from_path
import docx2txt
from docx2pdf import convert
import pandas as pd

# Change this to your poppler bin folder path
POPPLER_PATH = r"C:\Users\UTHRAVFST\AppData\Roaming\Microsoft\Windows\Network Shortcuts\Release-24.08.0-0\poppler-24.08.0\Library\bin"

input_folder = r"C:\kathir\all_file_1"


def extract_from_pdf(pdf_path, file_name):
    doc = fitz.open(pdf_path)

    # Create folder for this PDF
    output_path = os.path.join(input_folder, "extracted_output_1", file_name)
    os.makedirs(output_path, exist_ok=True)

    full_text = ""

    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        text = page.get_text().strip()

        if page_num == 12:
            tables = page.find_tables()
            df_table = pd.DataFrame(tables.tables[0].extract())
            df_table = df_table.dropna(axis=1, how="all")
            df_table = df_table.loc[:, ~(df_table.isna() | (df_table == "")).all()]
            print(df_table)

        if len(text) > 20:
            full_text += f"\n--- Page {page_num + 1} ---\n{text}\n"

            images = page.get_images(full=True)
            for img_index, img in enumerate(images, start=1):
                xref = img[0]
                base_image = doc.extract_image(xref)
                image_bytes = base_image["image"]
                image_ext = base_image["ext"]
                image_filename = (
                    f"{file_name}_page{page_num + 1}_img{img_index}.{image_ext}"
                )
                with open(os.path.join(output_path, image_filename), "wb") as img_file:
                    img_file.write(image_bytes)
        else:
            # Convert full page to image if text is minimal
            images = convert_from_path(
                pdf_path,
                first_page=page_num + 1,
                last_page=page_num + 1,
                poppler_path=POPPLER_PATH,
            )
            for i, img in enumerate(images):
                image_filename = f"{file_name}_page{page_num + 1}_converted.png"
                img.save(os.path.join(output_path, image_filename), "PNG")

    doc.close()

    # Save all collected text in one file
    text_file_path = os.path.join(output_path, f"{file_name}.txt")
    with open(text_file_path, "w", encoding="utf-8") as f:
        f.write(full_text)


def extract_from_docx(docx_path, file_name):
    output_path = os.path.join(input_folder, "extracted_output_1", file_name)
    os.makedirs(output_path, exist_ok=True)

    # Extract text and images
    text = docx2txt.process(docx_path, output_path)

    full_text = ""
    if len(text.strip()) > 20:
        full_text += f"\n--- Page 1 ---\n{text}\n"
    else:
        # Text is very short => possibly scanned docx (mostly images)
        # Convert docx to PDF for further processing if needed
        pdf_output_path = os.path.join(output_path, f"{file_name}.pdf")
        try:
            convert(docx_path, pdf_output_path)
            print(f"Converted scanned docx to PDF: {pdf_output_path}")
        except Exception as e:
            print(f"Error converting docx to PDF: {e}")

    # Save text file if any text extracted
    if full_text:
        text_filename = f"{file_name}.txt"
        with open(os.path.join(output_path, text_filename), "w", encoding="utf-8") as f:
            f.write(full_text)

    # Rename images extracted by docx2txt
    img_counter = 1
    for img_file in os.listdir(output_path):
        if img_file.lower().startswith("image") and img_file.lower().endswith(
            (".png", ".jpg", ".jpeg")
        ):
            old_path = os.path.join(output_path, img_file)
            ext = os.path.splitext(img_file)[1]
            new_name = f"{file_name}_page{img_counter}{ext}"
            new_path = os.path.join(output_path, new_name)
            os.rename(old_path, new_path)
            img_counter += 1


for file in os.listdir(input_folder):
    file_path = os.path.join(input_folder, file)
    file_name, ext = os.path.splitext(file)
    ext = ext.lower()

    if os.path.isfile(file_path):
        if ext == ".pdf":
            extract_from_pdf(file_path, file_name)
        elif ext == ".docx":
            extract_from_docx(file_path, file_name)

print("✅ All data extracted and saved page-wise.")
