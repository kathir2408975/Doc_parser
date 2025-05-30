import os
import fitz  # PyMuPDF
from pdf2image import convert_from_path
import docx2txt
from docx2pdf import convert
import pdfplumber

# import pandas as pd
# from docling.document_converter import DocumentConverter
from pathlib import Path

POPPLER_PATH = r"C:\Users\UTHRAVFST\AppData\Roaming\Microsoft\Windows\Network Shortcuts\Release-24.08.0-0\poppler-24.08.0\Library\bin"
input_folder = r"C:\kathir\all_file_1"


def extract_from_pdf(pdf_path, file_name, markdown_file):

    output_path = os.path.join(input_folder, "extracted_output_1", file_name)
    os.makedirs(output_path, exist_ok=True)

    images_output_path = os.path.join(
        input_folder, "extracted_output_1", file_name, "images"
    )
    os.makedirs(images_output_path, exist_ok=True)
    data = extract_text_and_tables(pdf_path)

    for page, content in data.items():

        text = content["text"]
        tables = content["tables"]

        table_texts = [
            cell for table in tables for row in table for cell in row if cell
        ]

        for txt in table_texts:
            text = text.replace(txt, "")

        table_presence = True if len(table_texts) > 0 else False

        data[page]["plain_text"] = text
        data[page]["table_presence"] = table_presence
        # print("page : ", page)
        # print("text : ", text)
        # print("tables : ", tables)

        # print("debugging")

    for page, content in data.items():

        # print("text : ", content["text"])
        # print("plain_text : ", content["plain_text"])
        if content["table_presence"]:
            print("table present on page : ", page)
        else:
            print("table not present on page : ", page)

    print("debugging")

    # # for docling conversion
    # converter = DocumentConverter()
    # result = converter.convert(pdf_path, max_num_pages=10)

    # text_file_path = os.path.join(output_path, f"{file_name}.txt")
    # table_file_path = os.path.join(output_path, f"{file_name}_table.txt")

    # with open(markdown_file, "r", encoding="utf-8") as file:
    #     markdown_content = file.read()

    # text_from_pdf, tables_from_pdf = separate_text_and_tables_from_markdown(markdown_content)

    # with open(text_file_path, "w", encoding="utf-8") as f:
    #     f.write(text_from_pdf)

    # with open(table_file_path, "w", encoding="utf-8") as f:
    #     for table in tables_from_pdf:
    #         f.write(table)

    # # pdfplumber
    # with pdfplumber.open(pdf_path) as pdf:
    #     for page_num, page in enumerate(pdf.pages, start=1):
    #         text = page.extract_text()
    #         tables = page.extract_tables()

    #         for table_num, table in enumerate(tables, start=1):

    #             for row in table:
    #                 for txt in row:

    #                     if txt in text and txt != "":
    #                         text = text.replace(txt, " ")

    #         print(f"\nPage {page_num}\n")
    #         print(f"\nText:\n{text}\n")
    #         print(f"\nTables:\n{tables}\n")
    #         print("testing")

    doc = fitz.open(pdf_path)
    full_text = ""

    for page_num in range(len(doc)):

        page = doc.load_page(page_num)
        text = page.get_text().strip()
        tables = page.find_tables()

        # if page_num == 12:
        #     tables = page.find_tables()
        #     df_table = pd.DataFrame(tables.tables[0].extract())
        #     df_table = df_table.dropna(axis=1, how="all")
        #     df_table = df_table.loc[:, ~(df_table.isna() | (df_table == "")).all()]
        #     print(df_table)

        if len(text) > 20:

            # full_text += f"\n--- Page {page_num + 1} ---\n{text}\n"

            images = page.get_images(full=True)
            for img_index, img in enumerate(images, start=1):

                xref = img[0]
                base_image = doc.extract_image(xref)
                image_bytes = base_image["image"]
                image_ext = base_image["ext"]
                image_filename = (
                    f"{file_name}_page{page_num + 1}_img{img_index}.{image_ext}"
                )

                with open(
                    os.path.join(images_output_path, image_filename), "wb"
                ) as img_file:
                    img_file.write(image_bytes)

        else:

            images = convert_from_path(
                pdf_path,
                first_page=page_num + 1,
                last_page=page_num + 1,
                poppler_path=POPPLER_PATH,
            )

            for i, img in enumerate(images):
                image_filename = f"{file_name}_page{page_num + 1}_converted.png"
                img.save(os.path.join(images_output_path, image_filename), "PNG")

    doc.close()

    # Save all collected text in one file
    # text_file_path = os.path.join(output_path, f"{file_name}.txt")
    # with open(text_file_path, "w", encoding="utf-8") as f:
    #     f.write(full_text)


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


def separate_text_and_tables_from_markdown(markdown_content):

    lines = markdown_content.split("\n")
    text = []
    tables = []
    current_table = []

    for line in lines:

        if line.strip().startswith("|"):
            current_table.append(line)

        else:

            if current_table:  # End of a table block
                tables.append("\n".join(current_table))
                current_table = []
            text.append(line)

    # Add any remaining table at the end
    if current_table:
        tables.append("\n".join(current_table))

    return "\n".join(text), tables


def extract_text_and_tables(pdf_path):
    page_data = {}

    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages):
            if page_num == 11:  # Limit to first 5 pages
                break

            text = page.extract_text()  # Extract text using pdfplumber
            tables = page.extract_tables()  # Extract tables using pdfplumber

            page_data[page_num + 1] = {"text": text, "tables": tables}

    return page_data


for file in os.listdir(input_folder):

    file_path = os.path.join(input_folder, file)
    file_path = Path(file_path).as_posix()

    file_name, ext = os.path.splitext(file)
    ext = ext.lower()
    markdown_file = "docling_output_full_file.txt"

    if os.path.isfile(file_path):

        if ext == ".pdf":
            extract_from_pdf(file_path, file_name, markdown_file)
        elif ext == ".docx":
            extract_from_docx(file_path, file_name)

print("✅ All data extracted and saved page-wise.")
