import openpyxl
from openpyxl.drawing.image import Image as OpenpyxlImage
import json
import os
from PIL import Image
from io import BytesIO

EXCEL_FILE = 'input/cartoon.xlsx'
OUTPUT_FOLDER = 'extracted_images'
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# Column mapping (adjust as needed)
COL_INDEX = {
    'productId': 1,
    'frontImage': 2,
    'backImage': 3,
    'frontSide': 4
}

# Load workbook
wb = openpyxl.load_workbook(EXCEL_FILE)
ws = wb.active

# Initialize map
product_map = {}

# Extract image data with anchor
for img in ws._images:
    # Get cell position from anchor (e.g., B2)
    anchor = img.anchor._from
    col, row = anchor.col + 1, anchor.row + 1  # 1-based indexing

    # Determine image type
    if col == COL_INDEX['frontImage']:
        image_type = 'frontImage'
    elif col == COL_INDEX['backImage']:
        image_type = 'backImage'
    elif col == COL_INDEX['frontSide']:
        image_type = 'frontSide'
    else:
        continue  # not a relevant column

    product_id = ws.cell(row=row, column=COL_INDEX['productId']).value
    if not product_id:
        continue

    if product_id not in product_map:
        product_map[product_id] = {}

    # Save image to disk
    img_data = img._data()
    image_filename = f"{product_id}_{image_type}.png"
    image_path = os.path.join(OUTPUT_FOLDER, image_filename)
    with open(image_path, "wb") as f:
        f.write(img_data)

    product_map[product_id][image_type] = image_filename

# Write JSON mapping
with open("product_image_mapping.json", "w") as f:
    json.dump(product_map, f, indent=2)

print("✅ Image mapping written to product_image_mapping.json")
