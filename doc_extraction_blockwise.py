import fitz  # PyMuPDF
import os

# === CONFIGURATION ===
pdf_path = r"C:\kathir\all_file_1\temp.pdf"
output_dir = r"C:\kathir\all_file_1\output_new_approach"
images_dir = os.path.join(output_dir, "images")
text_output = os.path.join(output_dir, "extracted_text.txt")

os.makedirs(output_dir, exist_ok=True)
os.makedirs(images_dir, exist_ok=True)

# === PROCESS WITH PyMuPDF ===
doc = fitz.open(pdf_path)
output_text = []

for page_num in range(len(doc)):
    output_text.append(f"\n----------Page : {page_num + 1}----------\n")
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

    # --- IMAGE EXTRACTION ---
    image_list = page.get_images(full=True)
    for img_index, img in enumerate(image_list):
        xref = img[0]
        base_image = doc.extract_image(xref)
        image_bytes = base_image["image"]
        image_ext = base_image["ext"]
        image_filename = f"page_{page_num + 1}_img_{img_index + 1}.{image_ext}"
        image_path = os.path.join(images_dir, image_filename)

        with open(image_path, "wb") as f:
            f.write(image_bytes)

        output_text.append(f"{image_filename}")

# --- SAVE TEXT OUTPUT ---
with open(text_output, "w", encoding="utf-8") as f:
    f.write("\n".join(output_text))

doc.close()
