import pandas as pd
import tkinter as tk
import os
import sys
import difflib
from tkinter import filedialog, messagebox
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib import units
inch = units.inch
from reportlab.platypus import Image
from reportlab.lib.colors import black  # To draw the border
from PIL import Image as PILImage
import subprocess  # For cross-platform PDF opening

def read_excel_data_with_header_row(file_path):
    """Reads data from an Excel file, treating the second row as the header
    and capturing the value from the first row."""
    try:
        # Read the first row to get the special value
        special_value_df = pd.read_excel(file_path, header=None, nrows=1)
        special_value = special_value_df.iloc[0, 0]  # Get the value from the first cell (row 0, column 0)
        # Read the data starting from the second row, using the second row as the header
        df = pd.read_excel(file_path, header=1)
        return special_value, df
    except FileNotFoundError:
        print(f"Error: File not found at '{file_path}'")
        return None, None
    except Exception as e:
        print(f"An error occurred while reading the Excel file: {e}")
        return None, None

def generate_labels(title, data, image_folder_name="logos", similarity_threshold=0.8, border_width=0.5, border_image_path="logos/Border.png"):
    """Generates styled PDF labels (4x8) with controlled vertical spacing and top/bottom image borders."""
    if data is None or data.empty:
        print("No data to generate labels from.")
        return
    
    if getattr(sys, 'frozen', False):
        # If the application is run as a bundled executable,
        # the script directory is the directory of the executable
        script_directory = os.path.dirname(sys.executable)
        border_directory = sys._MEIPASS

    else:
        # If the application is run as a Python script
        border_directory = os.path.dirname(os.path.abspath(__file__))
        script_directory = os.path.dirname(os.path.abspath(__file__))

    output_filename = f"{title}_labels.pdf"  # Corrected to use an f-string for string formatting
    c = canvas.Canvas(output_filename, pagesize=letter)

    # --- Label Layout Configuration ---
    num_cols = 4
    num_rows = 8
    horizontal_margin = 0.05 * inch
    vertical_margin = 0.05 * inch
    effective_width = (letter[0] - 2 * horizontal_margin) / num_cols
    effective_height = (letter[1] - 2 * vertical_margin) / num_rows
    label_width = effective_width - 2 * 0.015 * inch
    label_height = effective_height - 2 * 0.015 * inch
    label_inner_margin = 0.12 * inch  # Internal margin from top

    # --- Font Sizes & Styles ---
    code_font_name = "Times-Bold"  # Updated to Times Roman Bold
    code_font_size = 12
    name_font_name = "Helvetica"
    name_font_size = 14
    size_font_name = "Helvetica-Bold"
    size_font_size = 12
    line_spacing = 0.10 * inch

    # --- Image Handling ---
    image_width = 1.2 * inch
    image_height = 0.3 * inch
    image_folder = os.path.join(script_directory, image_folder_name)
    border_path = os.path.join(border_directory, border_image_path)
    border_image_height = 0.15 * inch  # Adjust as needed

    current_row_on_page = 0
    current_col_on_page = 0

    for index, row in data.iterrows():
        x_offset = horizontal_margin + current_col_on_page * effective_width + (effective_width - label_width) / 2
        y_offset = letter[1] - vertical_margin - (current_row_on_page + 1) * effective_height + (effective_height - label_height) / 2
        x_center = x_offset + label_width / 2
        current_y = y_offset + label_height - label_inner_margin  # Start drawing from the top with inner margin

        # --- Draw Label Border ---
        c.setStrokeColor(black)
        c.setLineWidth(border_width)
        c.rect(x_offset, y_offset, label_width, label_height)

        # --- Draw Top Border Image ---
        try:
            img_top_border = Image(border_path, width=label_width, height=border_image_height)
            img_top_border.drawOn(c, x_offset, y_offset + label_height - border_image_height)
        except Exception as e:
            print(f"Error loading top border image '{border_path}': {e}")

        # --- Draw Bottom Border Image ---
        try:
            img_bottom_border = Image(border_path, width=label_width, height=border_image_height)
            c.saveState()
            c.translate(x_offset + label_width / 2, y_offset + border_image_height / 2)  # Move to the center of the image
            c.rotate(180)  # Rotate the canvas 180 degrees
            img_bottom_border.drawOn(c, -label_width / 2, -border_image_height / 2)  # Draw the image centered
            c.restoreState()
        except Exception as e:
            print(f"Error loading bottom border image '{border_path}': {e}")

        # --- Draw Product Code (#) ---
        y_code = current_y - code_font_size  # Position code from the top
        if "#" in row:
            c.setFont(code_font_name, code_font_size)
            c.drawCentredString(x_center, y_code, str(row["#"]))
            current_y -= code_font_size + line_spacing  # Move down for the next element

        # --- Draw Product Name (Nombre) ---
        y_name = current_y - name_font_size  # Position name below code
        if "Nombre" in row:
            c.setFont(name_font_name, name_font_size)
            c.drawCentredString(x_center, y_name, str(row["Nombre"]))
            current_y -= name_font_size + line_spacing  # Move down for the next element

        # --- Draw Logo Image/Cursive Title ---
        found_image = False
        actual_img_height = 0
        image_y_top = current_y  # Top position for the image/title
        intended_image_height = 0.3 * inch  # The default image height we intend to use
        cursive_font_size = name_font_size * 1.2 # Use a consistent cursive font size
        title_vertical_offset = 0  # Initialize vertical offset for the title

        if os.path.exists(image_folder):
            for filename in os.listdir(image_folder):
                if filename.lower().endswith(".png"):
                    similarity = difflib.SequenceMatcher(None, title.lower().replace(' ', ''), filename.lower().replace(' ', '').replace('.png', '')).ratio()
                    if similarity > 0 and similarity >= similarity_threshold:
                        best_match_path = os.path.join(image_folder, filename)
                        try:
                            pil_img = PILImage.open(best_match_path)
                            original_width, original_height = pil_img.size
                            aspect_ratio = original_width / original_height
                            calculated_height = image_width / aspect_ratio
                            actual_img_height = calculated_height
                            img = Image(best_match_path, width=image_width, height=calculated_height)
                            img_y_draw = image_y_top - actual_img_height  # Position image from the top
                            img.drawOn(c, x_center - image_width / 2, img_y_draw)
                            current_y -= actual_img_height + line_spacing  # Move current_y below the image
                            found_image = True
                        except Exception as e:
                            print(f"Error loading image '{best_match_path}': {e}")
                            c.setFont("Times-Italic", cursive_font_size)
                            title_height = c.stringWidth(title, "Times-Italic", cursive_font_size) / 72 * (cursive_font_size * 1.5 / cursive_font_size) # Approximate height
                            c.drawCentredString(x_center, (image_y_top - (.02 * inch)) - (intended_image_height / 2), title)
                            current_y -= intended_image_height + line_spacing  # Move current_y by the intended height
                            actual_img_height = intended_image_height
                        break
            if not found_image:
                c.setFont("Times-Italic", cursive_font_size)
                title_height = c.stringWidth(title, "Times-Italic", cursive_font_size) / 72 * (cursive_font_size * 1.5 / cursive_font_size) # Approximate height
                c.drawCentredString(x_center, (image_y_top - (.02 * inch)) - (intended_image_height / 2), title)
                current_y -= intended_image_height + line_spacing  # Move current_y by the intended height
                actual_img_height = intended_image_height
        else:
            c.setFont("Times-Italic", cursive_font_size)
            title_height = c.stringWidth(title, "Times-Italic", cursive_font_size) / 72 * (cursive_font_size * 1.5 / cursive_font_size) # Approximate height
            c.drawCentredString(x_center, (image_y_top - (.02 * inch)) - (intended_image_height / 2), title)
            current_y -= intended_image_height + line_spacing  # Move current_y by the intended height
            actual_img_height = intended_image_height
            print(f"Warning: Image folder not found.")
            
            os.makedirs(image_folder)
            print(f"Created folder: {image_folder}")

        # --- Draw Size (Lugar) ---
        if "Lugar" in row:
            c.setFont(size_font_name, size_font_size)
            lugar_value = row["Lugar"]
            if pd.notna(lugar_value):
                try:
                    lugar_int = int(float(lugar_value))
                    # Adjust the vertical position of 'lugar' based on 'actual_img_height'
                    if not found_image: 
                        lugar_y = current_y - size_font_size + (actual_img_height / 2)
                    else:
                        lugar_y = current_y - size_font_size + (.06 * inch)

                    c.drawCentredString(x_center, lugar_y, str(lugar_int))
                    current_y -= size_font_size + line_spacing  # Update current_y below lugar
                except ValueError:
                    print(f"Warning: Invalid value '{lugar_value}' in 'Lugar' column, skipping.")

        # --- Move to the next label ---
        current_col_on_page += 1
        if current_col_on_page == num_cols:
            current_col_on_page = 0
            current_row_on_page += 1
            if current_row_on_page == num_rows and index < len(data) - 1:
                c.showPage()
                current_row_on_page = 0

    c.save()
    print(f"Styled labels with controlled vertical spacing saved to '{output_filename}'")
    return output_filename  # Return the filename of the generated PDF

def browse_file():
    file_path = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[("Excel files", "*.xlsx;*.xls"), ("All files", "*.*")]
    )
    if file_path:
        excel_path_entry.delete(0, tk.END)
        excel_path_entry.insert(0, file_path)

def open_pdf(filepath):
    """Opens the PDF file using the default viewer."""
    if os.name == 'nt':  # Windows
        os.startfile(filepath)
    elif os.name == 'posix':  # macOS and Linux
        subprocess.run(['open', filepath], check=True) # For macOS
        # For Linux, you might need to try 'xdg-open' instead of 'open'
        # subprocess.run(['xdg-open', filepath], check=True)

def generate_labels_from_ui():
    excel_file_path = excel_path_entry.get()
    if excel_file_path:
        title, data = read_excel_data_with_header_row(excel_file_path)
        if data is not None:
            pdf_filepath = generate_labels(title, data)
            if pdf_filepath:
                open_pdf(pdf_filepath)
    else:
        messagebox.showerror("Error", "Please select an Excel file.")

def setup_ui():
    """Sets up the main Tkinter user interface."""
    window = tk.Tk()
    window.title("Label Generator")

    # Label and Entry for Excel File Path
    excel_path_label = tk.Label(window, text="Excel File:")
    excel_path_label.grid(row=0, column=0, padx=10, pady=10, sticky="w")

    global excel_path_entry  # Make it accessible to other functions if needed
    excel_path_entry = tk.Entry(window, width=50)
    excel_path_entry.grid(row=0, column=1, padx=10, pady=10, sticky="ew")

    browse_button = tk.Button(window, text="Browse", command=browse_file)
    browse_button.grid(row=0, column=2, padx=10, pady=10, sticky="e")

    # Generate Labels Button
    generate_button = tk.Button(window, text="Generate Labels", command=generate_labels_from_ui)
    generate_button.grid(row=1, column=0, columnspan=3, padx=10, pady=10)

    window.grid_columnconfigure(1, weight=1) # Make the entry widget expand

    window.mainloop()

if __name__ == "__main__":
    setup_ui()