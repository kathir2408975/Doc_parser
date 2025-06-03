import os, docx2txt, pdfplumber, base64, json, io, json
from pdf2image import convert_from_path
from docx2pdf import convert
from mimetypes import guess_type
from PIL import Image

from pathlib import Path
from langchain.chat_models import AzureChatOpenAI
from langchain_core.messages import HumanMessage

POPPLER_PATH = r"C:\Users\UTHRAVFST\AppData\Roaming\Microsoft\Windows\Network Shortcuts\Release-24.08.0-0\poppler-24.08.0\Library\bin"
input_folder = r"C:\kathir\all_file_1"

azure_openai_creds_json_path = r"C:\Pranav\azure_openai_creds.json"
with open(azure_openai_creds_json_path, "r") as file:
    openai_creds = json.load(file)

az_llm = AzureChatOpenAI(
    deployment_name=openai_creds["deployment_name"],
    openai_api_key=openai_creds["OPENAI_API_KEY"],
    azure_endpoint=openai_creds["azure_endpoint"],
    openai_api_version=openai_creds["OPENAI_API_VERSION"],
)


def extract_from_pdf(pdf_path, file_name):

    output_path = os.path.join(input_folder, "extracted_output_1", file_name)
    os.makedirs(output_path, exist_ok=True)

    images_output_path = os.path.join(
        input_folder, "extracted_output_1", file_name, "images"
    )
    tables_as_images_output_path = os.path.join(
        input_folder, "extracted_output_1", file_name, "tables_as_images"
    )
    os.makedirs(images_output_path, exist_ok=True)
    os.makedirs(tables_as_images_output_path, exist_ok=True)

    data = extract_text_tables_images_from_pdf(pdf_path)
    for page, content in data.items():

        print(f"Extracting full text on page : {page}")
        data[page]["page_number"] = content["page"].page_number

        text, tables, images, page_data = (
            content["text"],
            content["tables"],
            content["images"],
            content["page"],
        )
        table_texts = [
            cell for table in tables for row in table for cell in row if cell
        ]

        for txt in table_texts:
            text = text.replace(txt, "")

        data[page]["page_plain_text"] = text
        page_full_text = text

        if images:
            for i, img in enumerate(images):

                image_path = os.path.join(images_output_path, f"page_{page}_{i}.png")
                x0, top, x1, bottom = img["x0"], img["top"], img["x1"], img["bottom"]
                image = page_data.to_image().original

                cropped_img = image.crop((x0, top, x1, bottom))
                base64_image = get_image_as_base64(cropped_img, image_path)

                messages_base64 = [
                    HumanMessage(
                        content=[
                            {
                                "type": "text",
                                "text": "Check whether the image is a logo or if there's any logo present in the image. If there is no logo present, then consider it as a valid image. If the image is valid then summarize the content in the image. Remember to summarize the entire content of the image as it is and no information should be left to be mentioned in the image ONLY IF it is a valid image. Return only None and nothing else if the image is not valid.",
                            },
                            {"type": "image_url", "image_url": {"url": base64_image}},
                        ]
                    )
                ]

                try:
                    response_base64 = az_llm.invoke(messages_base64)
                    image_summary = response_base64.content

                    if eval(image_summary) is not None:

                        print("Image present on page:", page)
                        # print(
                        #     f"Image summary for page {page}: and image {i} :  {image_summary}"
                        # )
                        page_full_text += "\n" + image_summary + "\n"

                except Exception as e:

                    print("LLM exception occured for image on page : ", page)

                # print("page_full_text : ", page_full_text)

        if table_texts:

            print("Table present on page:", page)
            images = convert_from_path(
                pdf_path, first_page=page, last_page=page, dpi=50
            )
            image_path = os.path.join(tables_as_images_output_path, f"page_{page}.png")

            base64_image = get_image_as_base64(images[0], image_path)
            messages_base64 = [
                HumanMessage(
                    content=[
                        {
                            "type": "text",
                            "text": "Summarize all the tables in the image. Remember to summarize only tables as it is and no information should be left to be mentioned in the tables. DO NOT change the named entities names or short forms or add your own interpretation. Return only summary and no extra text.",
                        },
                        {"type": "image_url", "image_url": {"url": base64_image}},
                    ]
                )
            ]

            try:

                response_base64 = az_llm.invoke(messages_base64)
                table_summary = response_base64.content
                # print(f"Table summary for page {page}: {table_summary}")

                page_full_text += "\n" + table_summary + "\n"
                # print("page_full_text : ", page_full_text)

            except Exception as e:

                print("LLM exception occured for table on page : ", page)

        data[page]["page_full_text"] = page_full_text

        if len(data[page]["text"]) > len(data[page]["page_full_text"]):

            data[page]["page_full_text"] = data[page]["text"]
            # print("page : ", page)

    with open(
        os.path.join(output_path, f"{file_name}.txt"), "w", encoding="utf-8"
    ) as file:

        for page, page_content in data.items():
            file.write("\n" + page_content["page_full_text"] + "\n")

    for key in data:
        data[key] = {
            "page_number": data[key]["page_number"],
            "page_full_text": data[key]["page_full_text"],
            "text": data[key]["text"],
            "page_plain_text": data[key]["page_plain_text"],
        }

    with open(os.path.join(output_path, f"{file_name}.json"), "w") as json_file:
        json.dump(data, json_file, indent=4)

    path_of_saved_full_text = output_path, f"{file_name}.txt"
    # return path_of_saved_full_text
    # print("debugging")

    # doc = fitz.open(pdf_path)

    # for page_num in range(len(doc)):

    #     page = doc.load_page(page_num)
    #     text = page.get_text().strip()
    #     tables = page.find_tables()

    #     # if text is present in page, save all images in the page to a folder
    #     if len(text) > 20:

    #         images = page.get_images(full=True)
    #         for img_index, img in enumerate(images, start=1):

    #             xref = img[0]
    #             base_image = doc.extract_image(xref)
    #             image_bytes = base_image["image"]
    #             image_ext = base_image["ext"]
    #             image_filename = (
    #                 f"{file_name}_page{page_num + 1}_img{img_index}.{image_ext}"
    #             )

    #             with open(
    #                 os.path.join(images_output_path, image_filename), "wb"
    #             ) as img_file:
    #                 img_file.write(image_bytes)

    #     else:
    #         # if text is not present in page, convert page to image and save it in the folder
    #         images = convert_from_path(
    #             pdf_path,
    #             first_page=page_num + 1,
    #             last_page=page_num + 1,
    #             poppler_path=POPPLER_PATH,
    #         )

    #         for i, img in enumerate(images):
    #             image_filename = f"{file_name}_page{page_num + 1}_converted.png"
    #             img.save(os.path.join(images_output_path, image_filename), "PNG")

    # doc.close()


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

    print("done")


def get_image_as_base64(image, image_path):

    image = image.resize((600, 600)).convert("L")
    image.save(image_path, "PNG", quality=50)

    with open(image_path, "rb") as image_file:

        base64_string = base64.b64encode(image_file.read()).decode("utf-8")
        mime_type, _ = guess_type(image_path)
        image_to_return = f"data:{mime_type};base64,{base64_string}"

    return image_to_return


def extract_text_tables_images_from_pdf(pdf_path):

    page_data = {}
    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages):

            # if page_num == 2:
            #     break

            text, tables, images = (
                page.extract_text(),
                page.extract_tables(),
                page.images,
            )

            page_data[page_num + 1] = {
                "page": page,
                "text": text,
                "tables": tables,
                "images": images,
            }

    return page_data


for file in os.listdir(input_folder):

    file_path = os.path.join(input_folder, file)
    file_path = Path(file_path).as_posix()

    file_name, ext = os.path.splitext(file)
    ext = ext.lower()

    if os.path.isfile(file_path):

        if ext == ".pdf":
            extract_from_pdf(file_path, file_name)
        elif ext == ".docx":
            extract_from_docx(file_path, file_name)

print("✅ All data extracted and saved page-wise.")
