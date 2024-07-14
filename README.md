# PDFs to Excel Converter

This project extracts data from unstructured from PDF files, processes the images to recognize text, and converts the text into an Excel spreadsheet.

### Overview
The script performs the following tasks:

1. **Extract Images from PDF**: Extracts images from each page of a given PDF file using PyMuPDF.
2. **Text Recognition**: Uses `easyocr` to detect and extract text from the images.
3. **Convert to Table**: Converts recognized text into a structured table format.
4. **Export to Excel**: Saves the table data into an Excel file using `pandas`.

### Functions

- `extract_images_from_pdf(pdf_file, path)`: 
  Extracts images from the specified PDF file and processes each image to convert it into table format.
  
  **Parameters:**
  - `pdf_file`: The name of the PDF file.
  - `path`: Directory path where the PDF file is located.
  
  **Returns:**
  - A list of tables extracted from the images in the PDF.

- `draw_boxes(image, bounds, color='yellow', width=2)`:
  Draws bounding boxes around detected text regions in an image.

  **Parameters:**
  - `image`: The image to draw boxes on.
  - `bounds`: List of bounding box coordinates.
  - `color`: Color of the bounding boxes.
  - `width`: Width of the bounding box lines.
  
  **Returns:**
  - The image with drawn bounding boxes.

- `image_to_table(image, table)`:
  Converts the detected text in an image into a structured table format.

  **Parameters:**
  - `image`: The image with detected text.
  - `table`: The existing table structure to append data to.
  
  **Returns:**
  - A structured table containing the text from the image.

- `table_to_excel(table, pdf_file)`:
  Exports the table data to an Excel file.

  **Parameters:**
  - `table`: The table data to export.
  - `pdf_file`: The original PDF file name to use in the output Excel file name.
  
  **Returns:**
  - A DataFrame containing the table data.

### System Design

1. **Data Extraction**:
   - Uses PyMuPDF to extract images from PDF files.

2. **Text Recognition**:
   - Utilizes `easyocr` to recognize and extract text from images.

3. **Data Conversion**:
   - Converts recognized text into a structured table format.

4. **Export to Excel**:
   - Uses `pandas` to save the table data into an Excel file.

### Conclusion

The PDF to Excel Converter effectively transforms images extracted from PDF files into structured tables and exports them to Excel. The combination of PyMuPDF for image extraction and easyocr for text recognition ensures accurate and efficient processing of textual data from PDF files.
