import os
import xlsxwriter
from PIL import Image
import io

# --- Configuration ---
# Prompt the user for the directory
image_dir = input("Enter the path to your image directory (or press Enter for './my_images'): ").strip().strip('"').strip("'")
if not image_dir:
    image_dir = './my_images'
    print(f"Defaulting to: {image_dir}")

# Verify the directory actually exists before proceeding
if not os.path.isdir(image_dir):
    print(f"Error: The directory '{image_dir}' does not exist. Exiting.")
    exit(1)
    
# image_dir = r'G:\Meu Drive\_TRABALHO\_07 PRODUTOS DIGITAIS\ICONIC ASSETS LOJA ETSY\Infoprodutos\PRINTS ORIGINAIS SKU'
excel_file = 'image_data.xlsx' # The name of the output file
valid_extensions = ('.png', '.jpg', '.jpeg', '.bmp', 'webp', '.svg')

# Create an Excel workbook and add a worksheet
workbook = xlsxwriter.Workbook(excel_file)
worksheet = workbook.add_worksheet()

# Write the headers
worksheet.write('A1', 'File Name')
worksheet.write('B1', 'Image')

# Widen the columns so the text and images fit better
worksheet.set_column('A:A', 30)
worksheet.set_column('B:B', 20)

row = 1
file_count = 0
for filename in sorted(os.listdir(image_dir)):
    if filename.lower().endswith(valid_extensions):
        try:
            filepath = os.path.join(image_dir, filename)
            file_count += 1

            if filename.lower().endswith('.svg'):
                # SVGs are vectors (tiny files) and Pillow can't read them.
                # We bypass the Pillow resize and insert them directly into Excel.
                worksheet.write(row, 0, filename)
                worksheet.set_row(row, 100)
                worksheet.insert_image(row, 1, filepath, {'x_scale': 0.3, 'y_scale': 0.3})
                print(f"\n#{file_count}) Added -> {filename}")
                row += 1
                continue

            # --- Resize the image before embedding ---
            with Image.open(filepath) as img:
                # Convert to RGB to prevent issues with transparent PNGs when saving as JPEG
                if img.mode in ('RGBA', 'P', 'LA'):
                    img = img.convert('RGB')
                    
                # Resize to a maximum of 1485x2100 pixels (maintains aspect ratio)
                img.thumbnail((1485, 2100))
                
                # Save to a memory buffer instead of the hard drive
                image_buffer = io.BytesIO()
                img.save(image_buffer, format='JPEG', quality=85)

            # Write the filename in column A
            worksheet.write(row, 0, filename)

            # Set row height to fit a thumbnail (adjust the 100 as needed)
            worksheet.set_row(row, 100)

            # Insert the image in column B
            # The x_scale and y_scale shrink the image to fit inside the cell
            worksheet.insert_image(row, 1, filepath, {'image_data': image_buffer})

            print(f"\n#{file_count}) Added -> {filename}")
            row +=1
        except Exception as e:
            # If an error occurs, print a warning and skip to the next file
            print(f"\nSkipping '{filename}' due to an error: {e}")
            continue

workbook.close()
print(f"\nSuccessfully created {excel_file} with embedded images! {file_count} file(s).")