import streamlit as st
import openpyxl
from openpyxl.styles import PatternFill
from PIL import Image
import io
import os

# --- Core Image to Excel Conversion Logic ---

def image_to_excel_pixel_art(image_data, should_resize, max_size, resampling_method, num_colors):
    """
    Convert an image to Excel pixel art with color quantization to prevent corruption.
    """
    img = Image.open(image_data)
    
    if img.mode != 'RGB':
        img = img.convert('RGB')
    
    # --- STEP 1: RESIZE (if requested) ---
    if should_resize:
        width, height = img.size
        if width > max_size or height > max_size:
            ratio = min(max_size / width, max_size / height)
            new_width = int(width * ratio)
            new_height = int(height * ratio)
            img = img.resize((new_width, new_height), resampling_method)
            st.info(f"Resized image from {width}x{height} to {new_width}x{new_height}.")
    else:
        st.warning("Processing image at original size. This may be very slow.")

    # --- STEP 2: QUANTIZE COLORS (Crucial for preventing corruption) ---
    # Reduce the number of unique colors in the image.
    # This prevents hitting Excel's style limit.
    st.info(f"Reducing image to a palette of {num_colors} colors...")
    # The 'dither=Image.Dither.NONE' option can be faster but looks less natural.
    # FLOYDSTEINBERG is higher quality.
    quantized_img = img.quantize(colors=num_colors, method=Image.MAXCOVERAGE, dither=Image.Dither.FLOYDSTEINBERG)
    # quantize() returns a 'P' (palette) mode image. We need to convert it back to RGB
    # so we can get the (R, G, B) value of each pixel.
    img = quantized_img.convert('RGB')
    st.info("Color reduction complete.")

    width, height = img.size
    
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Pixel Art"
    
    for col_idx in range(1, width + 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(col_idx)].width = 2.0
    
    for row_idx in range(1, height + 1):
        ws.row_dimensions[row_idx].height = 12.0
    
    fill_cache = {}
    
    progress_bar = st.progress(0, text="Processing pixels...")
    
    for y in range(height):
        for x in range(width):
            r, g, b = img.getpixel((x, y))
            hex_color = f"{r:02x}{g:02x}{b:02x}"
            
            if hex_color in fill_cache:
                fill = fill_cache[hex_color]
            else:
                fill = PatternFill(start_color=hex_color, end_color=hex_color, fill_type="solid")
                fill_cache[hex_color] = fill
            
            cell = ws.cell(row=y + 1, column=x + 1)
            cell.fill = fill
        
        progress_percentage = (y + 1) / height
        progress_bar.progress(progress_percentage)
            
    ws.sheet_view.showGridLines = False
    
    progress_bar.progress(1.0, text="Finalizing Excel file...")
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    
    st.success("Conversion Complete!")
    return buffer

# --- Streamlit User Interface ---

st.set_page_config(page_title="Image to Excel Art", layout="centered")
st.title("🖼️ Image to Excel Pixel Art Converter")

st.sidebar.header("⚙️ Settings")

# Color Palette Size
num_colors = st.sidebar.slider(
    "🎨 Number of Colors", 
    min_value=8, 
    max_value=256, 
    value=128,
    help="Fewer colors = faster processing and smaller files. More colors = higher fidelity. Recommended to keep below 256."
)

st.sidebar.markdown("---")

# Resizing Options
resize_image = st.sidebar.checkbox("Resize image for performance", value=True)
max_size = st.sidebar.slider("Resolution (if resizing)", min_value=32, max_value=512, value=128, disabled=not resize_image)
resampling_options = {"Nearest (Blocky)": Image.Resampling.NEAREST, "Lanczos (Smooth)": Image.Resampling.LANCZOS}
resampling_choice = st.sidebar.selectbox("Resizing Quality", list(resampling_options.keys()), disabled=not resize_image)
selected_method = resampling_options[resampling_choice]

uploaded_file = st.file_uploader("Choose an image file", type=['png', 'jpg', 'jpeg', 'bmp', 'gif'])

if uploaded_file is not None:
    if not resize_image:
        img_peek = Image.open(uploaded_file)
        width, height = img_peek.size
        uploaded_file.seek(0)
        st.error(f"**WARNING:** You are processing at original size ({width}x{height}). This will be VERY slow and may fail. Resizing is highly recommended.")

    st.image(uploaded_file, caption="Your Uploaded Image", use_column_width=True)
    
    if st.button("🎨 Convert to Excel Art"):
        with st.spinner("Processing... This may take several minutes."):
            try:
                excel_buffer = image_to_excel_pixel_art(
                    uploaded_file, 
                    should_resize=resize_image, 
                    max_size=max_size, 
                    resampling_method=selected_method,
                    num_colors=num_colors
                )
                
                original_filename = uploaded_file.name
                base_filename = os.path.splitext(original_filename)[0]
                size_str = f"{max_size}px" if resize_image else "original"
                excel_filename = f"{base_filename}_{size_str}_{num_colors}colors_art.xlsx"

                st.download_button("📥 Download Excel File", excel_buffer, excel_filename, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            except Exception as e:
                st.error(f"An error occurred: {e}")
                st.error("This can happen if the image is too large. Try enabling resizing or reducing the color count.")