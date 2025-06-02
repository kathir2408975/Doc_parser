import os, fitz, docx2txt, pdfplumber, base64, json
from pdf2image import convert_from_path
from docx2pdf import convert
from mimetypes import guess_type

from pathlib import Path
from langchain.chat_models import AzureChatOpenAI
from langchain_core.messages import HumanMessage

POPPLER_PATH = r"C:\Users\UTHRAVFST\AppData\Roaming\Microsoft\Windows\Network Shortcuts\Release-24.08.0-0\poppler-24.08.0\Library\bin"
input_folder = r"C:\kathir\all_file_1"

azure_openai_creds_json_path = r"C:\Pranav\azure_openai_creds.json"

with open(azure_openai_creds_json_path, "r") as file:
    openai_creds = json.load(file)

print(openai_creds)


az_llm = AzureChatOpenAI(
    deployment_name=openai_creds["deployment_name"],
    openai_api_key=openai_creds["OPENAI_API_KEY"],
    azure_endpoint=openai_creds["azure_endpoint"],
    openai_api_version=openai_creds["OPENAI_API_VERSION"],
)


def extract_from_pdf(pdf_path, file_name, markdown_file):

    output_path = os.path.join(input_folder, "extracted_output_1", file_name)
    os.makedirs(output_path, exist_ok=True)

    images_output_path = os.path.join(
        input_folder, "extracted_output_1", file_name, "images"
    )
    os.makedirs(images_output_path, exist_ok=True)

    tables_as_images_output_path = os.path.join(
        input_folder, "extracted_output_1", file_name, "tables_as_images"
    )
    os.makedirs(tables_as_images_output_path, exist_ok=True)

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

    for page, content in data.items():

        if content["table_presence"]:
            print("Table present on page:", page)

            base64_image = capture_page_as_base64(
                pdf_path, page, tables_as_images_output_path
            )

            messages_base64 = [
                HumanMessage(
                    content=[
                        {
                            "type": "text",
                            "text": "Summarize all the tables in the image. Remember to summarize only tables as it is and no information should be left mentioned in the tables.",
                        },
                        {"type": "image_url", "image_url": {"url": base64_image}},
                    ]
                )
            ]

            response_base64 = az_llm.invoke(messages_base64)
            summary = response_base64.content
            print(f"Summary for page {page}: {summary}")
        else:
            print("Table not present on page:", page)

    print("debugging")

    doc = fitz.open(pdf_path)

    for page_num in range(len(doc)):

        page = doc.load_page(page_num)
        text = page.get_text().strip()
        tables = page.find_tables()

        if len(text) > 20:

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


def extract_from_docx(docx_path, file_name):

    output_path = os.path.join(input_folder, "extracted_output_1", file_name)
    os.makedirs(output_path, exist_ok=True)

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

            text = page.extract_text()
            tables = page.extract_tables()

            page_data[page_num + 1] = {"text": text, "tables": tables}

    return page_data


def capture_page_as_base64(pdf_path, page_num, tables_as_images_output_path):

    images = convert_from_path(
        pdf_path, first_page=page_num, last_page=page_num, dpi=100
    )
    image_path = os.path.join(tables_as_images_output_path, f"page_{page_num}.png")
    image = images[0].resize((400, 400)).convert("L")
    image.save(image_path, "WEBP", quality=50)

    with open(image_path, "rb") as image_file:
        
        base64_string = base64.b64encode(image_file.read()).decode("utf-8")
        mime_type, _ = guess_type(image_path)
        image_to_return = f"data:{mime_type};base64,{base64_string}"

    return image_to_return


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
