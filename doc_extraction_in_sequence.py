import os
import pdfplumber
import fitz  # PyMuPDF
from pathlib import Path

def table_to_string(table):
    """
    Convert a table (list of lists) to a readable string format.
    """
    if not table:
        return ""
    # Join rows with newlines, and cells within a row with tabs
    return "\n".join("\t".join(str(cell) if cell is not None else "" for cell in row) for row in table)

def is_text_in_table(text_bbox, table_bboxes):
    """
    Check if a text item's bounding box is inside any table's bounding box.
    
    Args:
        text_bbox (tuple): (x0, y0, x1, y1) of the text item.
        table_bboxes (list): List of table bounding boxes [(x0, y0, x1, y1), ...].
    
    Returns:
        bool: True if text is inside a table, False otherwise.
    """
    tx0, ty0, tx1, ty1 = text_bbox
    for tbx0, tby0, tbx1, tby1 in table_bboxes:
        if (tx0 >= tbx0 and tx1 <= tbx1 and ty0 >= tby0 and ty1 <= tby1):
            return True
    return False

def group_words_into_lines(words):
    """
    Group words into lines based on y-coordinate proximity.
    
    Args:
        words (list): List of word dictionaries with text, x0, y0, x1, y1.
    
    Returns:
        list: List of line dictionaries with combined text and position.
    """
    if not words:
        return []
    
    # Sort words by y0 (top) and x0 (left) to process in reading order
    words = sorted(words, key=lambda w: (w["top"], w["x0"]))
    
    lines = []
    current_line = []
    last_y0 = words[0]["top"]
    line_y0 = last_y0
    line_x0 = words[0]["x0"]
    
    # Threshold for grouping words into the same line (adjust based on PDF)
    y_threshold = 5  # Pixels to consider words on the same line
    
    for word in words:
        y0 = word["top"]
        # If the word is close enough to the current line's y-coordinate
        if abs(y0 - last_y0) <= y_threshold:
            current_line.append(word["text"])
        else:
            # Finish the current line and start a new one
            if current_line:
                lines.append({
                    "type": "text",
                    "data": " ".join(current_line),  # Join words with spaces
                    "y0": line_y0,
                    "x0": line_x0
                })
            current_line = [word["text"]]
            line_y0 = y0
            line_x0 = word["x0"]
        last_y0 = y0
    
    # Append the last line
    if current_line:
        lines.append({
            "type": "text",
            "data": " ".join(current_line),
            "y0": line_y0,
            "x0": line_x0
        })
    
    return lines

def extract_pdf_content(pdf_path, output_dir):
    """
    Extract text, tables, and images from a PDF and save as a text file in sequence.
    
    Args:
        pdf_path (str): Path to the PDF file.
        output_dir (str): Directory to save extracted images and text output.
    """
    # Create output directory for images if it doesn't exist
    Path(output_dir).mkdir(parents=True, exist_ok=True)
    
    # Initialize list to store all content
    output_lines = []
    
    # Open PDF with pdfplumber and PyMuPDF
    with pdfplumber.open(pdf_path) as pdf, fitz.open(pdf_path) as doc:
        # Ensure both libraries report the same number of pages
        num_pages = min(len(pdf.pages), doc.page_count)
        
        for page_num in range(num_pages):
            # Add page header
            output_lines.append(f"\n--- Page {page_num + 1} ---\n")
            
            # Initialize content items for this page
            content_items = []
            
            # Extract text and tables with pdfplumber
            plumber_page = pdf.pages[page_num]
            
            # Get table bounding boxes to filter text
            table_bboxes = [table.bbox for table in plumber_page.find_tables()]
            
            # Extract text words, excluding those inside tables
            text_words = []
            for text_obj in plumber_page.extract_words():
                text_bbox = (text_obj["x0"], text_obj["top"], text_obj["x1"], text_obj["bottom"])
                if not is_text_in_table(text_bbox, table_bboxes):
                    text_words.append(text_obj)
            
            # Group words into lines
            content_items.extend(group_words_into_lines(text_words))
            
            # Extract tables with bounding boxes
            for table_idx, table_obj in enumerate(plumber_page.extract_tables()):
                table_bbox = plumber_page.find_tables()[table_idx].bbox if plumber_page.find_tables() else None
                if table_bbox:
                    content_items.append({
                        "type": "table",
                        "data": table_to_string(table_obj),  # Convert table to string
                        "y0": table_bbox[1],  # Top y-coordinate
                        "x0": table_bbox[0]   # Left x-coordinate
                    })
            
            # Extract images with PyMuPDF
            fitz_page = doc[page_num]
            image_list = fitz_page.get_images(full=True)
            for img_index, img in enumerate(image_list):
                xref = img[0]
                base_image = doc.extract_image(xref)
                image_bytes = base_image["image"]
                image_ext = base_image["ext"]
                image_filename = f"{output_dir}/page_{page_num + 1}_img_{img_index + 1}.{image_ext}"
                
                # Save image to file
                with open(image_filename, "wb") as img_file:
                    img_file.write(image_bytes)
                
                # Approximate image position
                img_info = fitz_page.get_image_info()
                y0 = img_info[img_index]["bbox"][1] if img_index < len(img_info) else 0
                x0 = img_info[img_index]["bbox"][0] if img_index < len(img_info) else 0
                
                content_items.append({
                    "type": "image",
                    "data": f"[Image: {os.path.basename(image_filename)}]",
                    "y0": y0,
                    "x0": x0
                })
            
            # Sort content by y-coordinate (top to bottom) and x-coordinate (left to right for columns)
            y_threshold = 10  # Adjust for grouping items in same row
            content_items.sort(key=lambda x: (round(x["y0"] / y_threshold), x["x0"]))
            
            # Convert content items to text lines
            for item in content_items:
                output_lines.append(item["data"])
                output_lines.append("")  # Add blank line for readability
            
            output_lines.append("")  # Extra blank line between pages
        
        # Save output as text file
        output_txt_path = os.path.join(output_dir, f"{os.path.splitext(os.path.basename(pdf_path))[0]}_output.txt")
        with open(output_txt_path, "w", encoding="utf-8") as f:
            f.write("\n".join(output_lines))
        
        return output_txt_path

def process_folder(input_folder, output_folder):
    """
    Process all PDFs in a folder and extract their content.
    
    Args:
        input_folder (str): Path to folder containing PDFs.
        output_folder (str): Path to save extracted content.
    """
    # Create output folder if it doesn't exist
    Path(output_folder).mkdir(parents=True, exist_ok=True)
    
    # Get all PDF files in the input folder
    pdf_files = [f for f in os.listdir(input_folder) if f.lower().endswith(".pdf")]
    
    for pdf_file in pdf_files:
        pdf_path = os.path.join(input_folder, pdf_file)
        print(f"Processing {pdf_file}...")
        output_txt = extract_pdf_content(pdf_path, output_folder)
        print(f"Saved output to {output_txt}")


if __name__ == "__main__":

    input_folder = r"C:\kathir\all_file_1"  # Replace with your PDF folder path
    output_folder = r"C:\kathir\all_file_1\output_new_approach"   # Replace with your output folder path
    process_folder(input_folder, output_folder)
    print("done")