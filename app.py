import streamlit as st
import openpyxl
from openpyxl.styles import PatternFill
from PIL import Image
import io
import os

# --- Core Image to Excel Conversion Logic ---
# (Slightly modified to work with in-memory objects and provide progress updates)

def image_to_excel_pixel_art(image_data, max_size=100):
    """
    Convert an image to Excel pixel art.
    
    Args:
        image_data (BytesIO or file-like object): The image data from the user upload.
        max_size (int): Maximum width/height in pixels.
        
    Returns:
        BytesIO: An in-memory Excel file.
    """
    
    # Open and process the image
    img = Image.open(image_data)
    
    # Convert to RGB if necessary
    if img.mode != 'RGB':
        img = img.convert('RGB')
    
    # Resize image if it's too large
    width, height = img.size
    if width > max_size or height > max_size:
        ratio = min(max_size/width, max_size/height)
        new_width = int(width * ratio)
        new_height = int(height * ratio)
        # Use Image.Resampling.NEAREST for a more "pixelated" look
        img = img.resize((new_width, new_height), Image.Resampling.NEAREST)
        st.info(f"Image was large, resized from {width}x{height} to {new_width}x{new_height}")

    width, height = img.size
    
    # Create a new workbook and select the active sheet
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Pixel Art"
    
    # Set cell dimensions to be square-ish
    # Note: Column width is in "character units", row height is in "points".
    # A ratio of ~1:6 (width:height) often looks roughly square.
    for col_idx in range(1, width + 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(col_idx)].width = 2.0
    
    for row_idx in range(1, height + 1):
        ws.row_dimensions[row_idx].height = 12.0
    
    # --- Progress reporting for Streamlit ---
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    # Process each pixel and color the corresponding Excel cell
    for y in range(height):
        for x in range(width):
            # Get pixel color
            r, g, b = img.getpixel((x, y))
            
            # Convert RGB to hex format for Excel
            hex_color = f"{r:02x}{g:02x}{b:02x}"
            
            # Create a fill pattern with the pixel color
            fill = PatternFill(start_color=hex_color, end_color=hex_color, fill_type="solid")
            
            # Apply the fill to the corresponding Excel cell (1-indexed)
            cell = ws.cell(row=y + 1, column=x + 1)
            cell.fill = fill
        
        # Update progress bar
        progress_percentage = (y + 1) / height
        progress_bar.progress(progress_percentage)
        status_text.text(f"Processing row {y+1}/{height}")
            
    # Remove gridlines for a cleaner look
    ws.sheet_view.showGridLines = False
    
    status_text.text("Saving to memory...")

    # Save the workbook to an in-memory buffer
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0) # Rewind the buffer to the beginning
    
    progress_bar.empty()
    status_text.success("Conversion Complete!")
    
    return buffer

# --- Streamlit User Interface ---

st.set_page_config(page_title="Image to Excel Pixel Art", layout="centered")

st.title("🖼️ Image to Excel Pixel Art Converter")

st.markdown("""
Upload an image, and this app will convert it into an Excel spreadsheet where each cell is colored 
to represent a pixel.
""")

# 1. Image Upload
uploaded_file = st.file_uploader(
    "Choose an image file",
    type=['png', 'jpg', 'jpeg', 'bmp', 'gif']
)

# 2. Settings
st.sidebar.header("⚙️ Settings")
max_size = st.sidebar.slider(
    "Max Resolution (pixels)", 
    min_value=32, 
    max_value=256, 
    value=100,
    help="Higher values take longer to process but produce more detailed art. The image will be resized to fit within this square dimension."
)

if uploaded_file is not None:
    # Display the uploaded image
    st.image(uploaded_file, caption="Your Uploaded Image", use_column_width=True)
    
    # 3. Conversion Button
    if st.button("🎨 Convert to Excel Pixel Art"):
        with st.spinner("Converting your image... This may take a moment."):
            try:
                # Call the conversion function
                excel_buffer = image_to_excel_pixel_art(uploaded_file, max_size)
                
                # Generate a filename for the download
                # Takes the original filename and replaces the extension
                original_filename = uploaded_file.name
                base_filename = os.path.splitext(original_filename)[0]
                excel_filename = f"{base_filename}_pixel_art.xlsx"

                # 4. Download Button
                st.download_button(
                    label="📥 Download Excel File",
                    data=excel_buffer,
                    file_name=excel_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.error(f"An error occurred: {e}")
                st.error("Please try a different image or adjust the settings.")

else:
    st.info("Please upload an image to get started.")

st.sidebar.markdown("---")
st.sidebar.markdown("Made with ❤️ using [Streamlit](https://streamlit.io) and [OpenPyXL](https://openpyxl.readthedocs.io).")