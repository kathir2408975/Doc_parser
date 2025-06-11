import fitz, os, json, io, base64
from PIL import Image
from mimetypes import guess_type

from langchain.chat_models import AzureChatOpenAI
from langchain_core.messages import HumanMessage

azure_openai_creds_json_path = r"C:\Pranav\azure_openai_creds.json"
with open(azure_openai_creds_json_path, "r") as file:
    openai_creds = json.load(file)

pdf_path = r"C:\kathir\all_file_1\Global Economic Prospects.pdf"
output_dir = r"C:\kathir\all_file_1\output_new_approach"
images_dir = os.path.join(output_dir, "images")
text_output = os.path.join(output_dir, "extracted_text.txt")

os.makedirs(output_dir, exist_ok=True)
os.makedirs(images_dir, exist_ok=True)

az_llm = AzureChatOpenAI(
    deployment_name=openai_creds["deployment_name"],
    openai_api_key=openai_creds["OPENAI_API_KEY"],
    azure_endpoint=openai_creds["azure_endpoint"],
    openai_api_version=openai_creds["OPENAI_API_VERSION"],
)

# === PROCESS WITH PyMuPDF ===
doc = fitz.open(pdf_path)
output_text = []

for page_num in range(len(doc)):

    actual_page_number = page_num + 1
    output_text.append(f"\n---------Page {actual_page_number}---------\n")
    page = doc[page_num]
    blocks = page.get_text("dict")["blocks"]

    # --- TEXT BLOCKS (Column-wise) ---
    left_column = []
    right_column = []
    column_split_x = page.rect.width / 2

    for block in blocks:
        if block["type"] != 0:
            continue
        x0 = block["bbox"][0]
        y0 = block["bbox"][1]
        text = ""
        for line in block["lines"]:
            for span in line["spans"]:
                text += span["text"] + " "
        text = text.strip()

        if x0 < column_split_x:
            left_column.append((y0, text))
        else:
            right_column.append((y0, text))

    left_column.sort(key=lambda x: x[0])
    right_column.sort(key=lambda x: x[0])
    sorted_blocks = left_column + right_column

    for _, text in sorted_blocks:
        output_text.append(text)

    # --- IMAGE EXTRACTION AND PROCESSING ---
    image_list = page.get_images(full=True)
    for img_index, img in enumerate(image_list):
        xref = img[0]
        base_image = doc.extract_image(xref)
        image_bytes = base_image["image"]
        image_filename = f"page_{page_num + 1}_img_{img_index + 1}.png"
        image_path = os.path.join(images_dir, image_filename)

        # Convert bytes to PIL Image for processing
        image = Image.open(io.BytesIO(image_bytes))
        # Resize to 600x600 and convert to grayscale
        image = image.resize((200, 200)).convert("L")
        # Save as PNG with quality=50
        image.save(image_path, "PNG", quality=50)

        # Convert to base64
        with open(image_path, "rb") as image_file:
            base64_string = base64.b64encode(image_file.read()).decode("utf-8")
        mime_type, _ = guess_type(image_path)
        base64_image = f"data:{mime_type};base64,{base64_string}"

        messages_base64 = [
            HumanMessage(
                content=[
                    {
                        "type": "text",
                        "text": "Check whether the image is a logo or if there's any logo present in the image or if there's any actual content present in the image. If there is any actual content present in the image, then only consider it as a valid image. If the image is valid then summarize the whole content in the image. Remember to summarize the entire content of the image as it is and no information should be left to be mentioned ONLY IF it is a valid image. Return only None and nothing else if the image is not valid.",
                    },
                    {"type": "image_url", "image_url": {"url": base64_image}},
                ]
            )
        ]

        try:
            response_base64 = az_llm.invoke(messages_base64)
            image_summary = response_base64.content

            if image_summary != "None":

                print(f"Image summary for page {actual_page_number} :  {image_summary}")
                output_text.append(f"{image_summary}")

            else:

                print(f"Logo summary for page {actual_page_number} :  {image_summary}")

        except Exception as e:

            print("LLM exception occured for image on page : ", actual_page_number)

        # print("testing")

        # output_text.append(f"_{image_to_return}_")

# --- SAVE TEXT OUTPUT ---
with open(text_output, "w", encoding="utf-8") as f:
    f.write("\n".join(output_text))

doc.close()

print("done")
