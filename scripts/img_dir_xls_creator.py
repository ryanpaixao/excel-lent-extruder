import os
import xlsxwriter

# --- Configuration ---
image_dir = './image_dir' # Replace with the path to your image directory
excel_file = 'image_data.xlsx' # The name of the output file
valid_extensions = ('.png', '.jpg', '.jpeg', '.bmp')

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
for filename in os.listdir(image_dir):
    if filename.lower().endswith(valid_extensions):
        filepath = os.path.join(image_dir, filename)

        # Write the filename in column A
        worksheet.write(row, 0, filename)

        # Set row height to fit a thumbnail (adjust the 100 as needed)
        worksheet.set_row(row, 100)

        # Insert the image in column B
        # The x_scale and y_scale shrink the image to fit inside the cell
        worksheet.insert_image(row, 1, filepath, {'x_scale': 0.1, 'y_scale': 0.1})

        print(f"\nAdded -> {filename}")
        row +=1

workbook.close()
print(f"\nSuccessfully created {excel_file} with embedded images!")