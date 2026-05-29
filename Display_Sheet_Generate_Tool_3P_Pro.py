import os
import sys
import glob
import customtkinter as ctk
import tkinter as tk
import xml.etree.ElementTree as ET
from tkcalendar import DateEntry
from tkinter import ttk, filedialog, simpledialog, messagebox
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_BREAK
from datetime import datetime
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


# 🔴 SET EXPIRY DATE HERE
EXPIRY_DATE = "2026-12-31"   # YYYY-MM-DD

def check_expiry():
    today = datetime.now().date()
    expiry = datetime.strptime(EXPIRY_DATE, "%Y-%m-%d").date()

    if today > expiry:
        messagebox.showerror("Error", "This .EXE has expired. Contact developer - abhishek.verma@genus.in")
        input("Press Enter to exit...")
        sys.exit()

# Call before GUI starts
check_expiry()


try:
    import pypandoc
    pandoc_installed = True
except ImportError:
    pandoc_installed = False

try:
    from docx2pdf import convert as docx2pdf_convert
    docx2pdf_installed = True
except ImportError:
    docx2pdf_installed = False

# # === Predefined mapping of decimal value to description ===

def read_dict_from_section(root, section):
    data_dict = {}
    for item in root.find(section).findall("Item"):
        key = int(item.get("id"))
        value = item.text.strip()
        data_dict[key] = value
    return data_dict

# Load XML
tree = ET.parse("predefined_names_3p_pro.xml")
root = tree.getroot()

# Read each dictionary
predefined_names_new = read_dict_from_section(root, "PredefinedNames")
diag_predefined_names_new = read_dict_from_section(root, "DiagPredefinedNames")
test_predefined_names_new = read_dict_from_section(root, "TestPredefinedNames")


# Global variables for folder paths
base_folder = ""
image_folders = {}
# template_path = ""

def add_page_number(paragraph):
    """Insert 'Page: X of Y' field into a paragraph with 'Page:' in bold."""
    # --- "Page:" text (bold) ---
    run = paragraph.add_run("Page: ")
    run.bold = True   # make only "Page:" bold

    # --- PAGE field ---
    fldChar1 = OxmlElement('w:fldChar')
    fldChar1.set(qn('w:fldCharType'), 'begin')

    instrText1 = OxmlElement('w:instrText')
    instrText1.text = "PAGE"

    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'separate')

    text_elem = OxmlElement('w:t')
    text_elem.text = "1"  # Placeholder before update
    fldChar2.append(text_elem)

    fldChar3 = OxmlElement('w:fldChar')
    fldChar3.set(qn('w:fldCharType'), 'end')

    r_element = paragraph.add_run()._r
    r_element.append(fldChar1)
    r_element.append(instrText1)
    r_element.append(fldChar2)
    r_element.append(fldChar3)

    # --- " of " text ---
    paragraph.add_run(" of ")

    # --- NUMPAGES field ---
    fldChar1 = OxmlElement('w:fldChar')
    fldChar1.set(qn('w:fldCharType'), 'begin')

    instrText2 = OxmlElement('w:instrText')
    instrText2.text = "NUMPAGES"

    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'separate')

    text_elem2 = OxmlElement('w:t')
    text_elem2.text = "1"
    fldChar2.append(text_elem2)

    fldChar3 = OxmlElement('w:fldChar')
    fldChar3.set(qn('w:fldCharType'), 'end')

    r_element = paragraph.add_run()._r
    r_element.append(fldChar1)
    r_element.append(instrText2)
    r_element.append(fldChar2)
    r_element.append(fldChar3)


# === Function to set base folder dynamically ===
def choose_base_folder():
    global base_folder, image_folders
    # , template_path

    folder = filedialog.askdirectory(title="Select Base Folder")
    if not folder:
        return
    
    base_folder = folder

    image_folders = {
        "3P-Pro WC 10-60A" : os.path.join(base_folder, "3P Pro WC Display"),
        # "3P-Pro WC 20-100A" : os.path.join(base_folder, "3P Pro WC Display"),
    }

    # Update meter type combobox values
    meter_type_combo.configure(values=list(image_folders.keys()))
    meter_type_var.set("Select Type")

    # Update default output path
    output_var.set(os.path.join(base_folder, "standards_report.docx"))

    messagebox.showinfo("Base Folder Selected", f"Base folder set to:\n{base_folder}")

def choose_output():
    # Ask user to select folder
    folder = filedialog.askdirectory(title="Select Output Folder")
    if not folder:
        return  # User cancelled
    
    # Ask user for file name
    file_name = simpledialog.askstring("File Name", "Enter Word file name (without extension):")
    if not file_name:  # Cancel or empty
        messagebox.showwarning("Invalid Name", "File name cannot be empty!")
        return
    
    # Create full path with .docx extension
    path = os.path.join(folder, file_name + ".docx")
    
    output_var.set(path)  # Set to variable


# === Function to collect available images ===
def get_image_list(folder):
    images = []
    for file in glob.glob(os.path.join(folder, "*.[jp][pn]g")):
        fname = os.path.basename(file)  # "000_LCD-Check.jpg or .png"
        name = os.path.splitext(fname)[0]  # "000_LCD-Check"
        images.append(name)
    return sorted(images)


# === Table creation in docx ===
def create_table(doc, values, meter_type):
    table = doc.add_table(rows=1, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.allow_autofit = False
    table.style = 'Table Grid'

    hdr_cells = table.rows[0].cells
    hdr_texts = ['S.No', 'Parameters Display']
    column_widths = [Inches(0.5), Inches(7.0)]
    for i, w in enumerate(column_widths):
        hdr_cells[i].width = w
        run = hdr_cells[i].paragraphs[0].add_run(hdr_texts[i])
        run.bold = True
        run.font.size = Pt(12)

    folder = image_folders[meter_type]
    # for idx, dec in enumerate(values, start=1):
    idx = 1
    for dec in values:
        filename = f"{dec:04d}"

        # # Get parameter name from XML mapping
        # param_name = predefined_names_new.get(dec, "").lower()

        # # If parameter contains "energy" use energy_decimal logic
        # if energy_decimal in ["1", "2"] and "energy" in param_name and "hr" not in param_name:
        #     matches = glob.glob(os.path.join(folder, f"{filename}@{energy_decimal}_*.[jp][pn]g"))
            
        # else:
        matches = glob.glob(os.path.join(folder, f"{filename}_*.[jp][pn]g"))

        if matches:
            for img_path in matches:
                row = table.add_row().cells
                row[0].text = str(idx)
                run = row[1].paragraphs[0].add_run()
                run.add_picture(img_path, width=Inches(6.5))
                idx += 1
        else:
            row = table.add_row().cells
            row[0].text = str(idx)
            row[1].text = f"Image {filename} not found"
            idx += 1


# === Generate document === 
def generate_doc():
    global auto_string, push_string, diag_string, test_string
    if not base_folder:
        messagebox.showerror("Error", "Please select a Base Folder first")
        return

    meter_type = meter_type_var.get()
    if meter_type not in image_folders:
        messagebox.showerror("Error", "Please select a valid Meter Type")
        return 
    
    auto_vals = [int(item.split(" - ")[0]) for item in auto_selected.get(0, tk.END)]
    push_vals = [int(item.split(" - ")[0]) for item in push_selected.get(0, tk.END)]
    diag_vals = [int(item.split(" - ")[0]) for item in diag_selected.get(0, tk.END)]
    test_vals = [int(item.split(" - ")[0]) for item in test_selected.get(0, tk.END)]

    auto_input_string = auto_input_hex_string.get().strip()
    auto_input_string_len = len(auto_input_string) % 4
    if(auto_input_string_len != 0): 
        messagebox.showerror("Error", "Input Auto hex String is Invalid, Check Length.")
        return
    
    push_input_string = push_input_hex_string.get().strip()
    push_input_string_len = len(push_input_string) % 4
    if(push_input_string_len != 0): 
        messagebox.showerror("Error", "Input Push hex String is Invalid, Check Length.")
        return

    diag_input_string = diag_input_hex_string.get().strip()
    diag_input_string_len = len(diag_input_string) % 4
    if(diag_input_string_len != 0): 
        messagebox.showerror("Error", "Input Diagnostic hex String is Invalid, Check Length.")
        return
    
    test_input_string = test_input_hex_string.get().strip()
    test_input_string_len = len(test_input_string) % 4  
    if(test_input_string_len != 0): 
        messagebox.showerror("Error", "Input Test hex String is Invalid, Check Length.")
        return

    if auto_input_string and auto_vals:
        messagebox.showerror("Error", "Please select either parameters or insert hex string in Auto Scroll, not both")
        return
    elif auto_input_string:
        auto_vals = [int(auto_input_string[i:i+4], 16) for i in range(0, len(auto_input_string), 4)]

    if push_input_string and push_vals:
        messagebox.showerror("Error", "Please select either parameters or insert hex string in Push Scroll, not both")
        return
    elif push_input_string:
        push_vals = [int(push_input_string[i:i+4], 16) for i in range(0, len(push_input_string), 4)]

    if diag_input_string and diag_vals:
        messagebox.showerror("Error", "Please select either parameters or insert hex string in Diagnostic Scroll, not both")
        return
    elif diag_input_string:
        diag_vals = [int(diag_input_string[i:i+4], 16) for i in range(0, len(diag_input_string), 4)]

    if test_input_string and test_vals:
        messagebox.showerror("Error", "Please select either parameters or insert hex string in Test Scroll, not both")
        return
    elif test_input_string:
        test_vals = [int(test_input_string[i:i+4], 16) for i in range(0, len(test_input_string), 4)]


    if not auto_vals:
        # elif not auto_input_hex_string:
        messagebox.showerror("Error", "Please select at least one parameter or insert hex string in Auto Scroll")
        return
    if not push_vals:
    # elif not push_input_hex_string:
        messagebox.showerror("Error", "Please select at least one parameter or insert hex string in Push Scroll")
        return
    if not diag_vals:
        diag_vals = [int(item.split(" - ")[0]) for item in ["92", "3101", "3103", "3106", "3107", "3102", "3105", "3104", "3108"]]
        

    auto_hex_str = "".join(f"{val:04X}" for val in auto_vals)
    auto_hex_str_len = len(auto_hex_str) // 2
    if(auto_hex_str_len <= 127):
        auto_hex_str_temp = f"{auto_hex_str_len:02X}" + auto_hex_str
    elif(auto_hex_str_len >127 and auto_hex_str_len <= 255):
        auto_hex_str_temp = "81" + f"{auto_hex_str_len:02X}" + auto_hex_str
    else:
        auto_hex_str_temp = "82" + f"{auto_hex_str_len:04X}" + auto_hex_str
    auto_hex_string = "09" + auto_hex_str_temp
    auto_string.set(auto_hex_string)
      
    push_hex_str = "".join(f"{val:04X}" for val in push_vals)
    push_hex_str_len = len(push_hex_str) // 2
    if(push_hex_str_len <= 127):
        push_hex_str_temp = f"{push_hex_str_len:02X}" + push_hex_str
    elif(push_hex_str_len >127 and push_hex_str_len <= 255):
        push_hex_str_temp = "81" + f"{push_hex_str_len:02X}" + push_hex_str
    else:
        push_hex_str_temp = "82" + f"{push_hex_str_len:04X}" + push_hex_str
    push_hex_string = "09" + push_hex_str_temp
    push_string.set(push_hex_string)

    diag_hex_str = "".join(f"{val:04X}" for val in diag_vals)
    diag_hex_str_len = len(diag_hex_str) // 2
    if(diag_hex_str_len <= 127):
        diag_hex_str_temp = f"{diag_hex_str_len:02X}" + diag_hex_str
    elif(diag_hex_str_len >127 and diag_hex_str_len <= 255):
        diag_hex_str_temp = "81" + f"{diag_hex_str_len:02X}" + diag_hex_str
    else:
        diag_hex_str_temp = "82" + f"{diag_hex_str_len:04X}" + diag_hex_str
    diag_hex_string = "09" + diag_hex_str_temp
    diag_string.set(diag_hex_string)
      
    test_hex_str = "".join(f"{val:04X}" for val in test_vals)
    test_hex_str_len = len(test_hex_str) // 2
    if(test_hex_str_len <= 127):
        test_hex_str_temp = f"{test_hex_str_len:02X}" + test_hex_str
    elif(test_hex_str_len >127 and test_hex_str_len <= 255):
        test_hex_str_temp = "81" + f"{test_hex_str_len:02X}" + test_hex_str
    else:
        test_hex_str_temp = "82" + f"{test_hex_str_len:04X}" + test_hex_str
    test_hex_string = "09" + test_hex_str_temp
    test_string.set(test_hex_string)
    
    # ==== Create Template in Docx ===
    # template_path = os.path.join(base_folder, "template.docx")
    # Auto-create template.docx if it does not exist
    # if not os.path.exists(template_path):
    doc = Document()
    
    section = doc.sections[0]
    
    # Page size (A4 example)
    section.page_height = Inches(11.69)   # A4 height in inches
    section.page_width  = Inches(8.27)    # A4 width in inches

    # Orientation (Landscape)
    # Swap width and height
    # section.orientation = WD_ORIENT.LANDSCAPE
    # section.page_width, section.page_height = section.page_height, section.page_width

    # Page margins
    section.top_margin    = Inches(0.5)
    section.bottom_margin = Inches(0.5)
    section.left_margin   = Inches(0.5)
    section.right_margin  = Inches(0.5)

    # Distance between top of page and start of header
    section.header_distance = Inches(0.5)   # default is around 0.5"

    # Distance between bottom of page and start of footer
    section.footer_distance = Inches(0.5)   # default is around 0.5"

    # Add header (appears on all pages)
    header = section.header

    # Create table with 2 columns
    table = header.add_table(rows=3, cols=3, width=Inches(7.7))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = True
    table.style = 'Table Grid'

    # Set column widths (optional, so it looks neat)
    col_widths = [Inches(1.4), Inches(4.2), Inches(2.0)]
    for col, width in zip(table.columns, col_widths):
        for cell in col.cells:
            cell.width = width

    # --- Row height function ---
    def set_row_height(row, height_in_inches):
        tr = row._tr
        trPr = tr.get_or_add_trPr()

        trHeight = OxmlElement('w:trHeight')
        trHeight.set(qn('w:val'), str(int(height_in_inches * 1440)))  # height in twips
        trHeight.set(qn('w:hRule'), 'exact')  # enforce exact height
        trPr.append(trHeight)

    # Set heights for header rows
    set_row_height(table.rows[0], 0.5)
    set_row_height(table.rows[1], 0.36)
    set_row_height(table.rows[2], 0.36)

    # === Insert company logo into [0,0] cell ===
    logo_path = os.path.join(base_folder, "Template images", "Company Logo.jpg")

    if os.path.exists(logo_path):
        cell = table.cell(0, 0)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        para = cell.paragraphs[0]
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = para.add_run()
        run.add_picture(logo_path, width=Inches(1.0))  # adjust width as needed
    else:
        table.cell(0, 0).text = "Logo Missing"

    cell = table.cell(0, 1)
    para = cell.paragraphs[0]
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = para.add_run("Standard Formats")
    run.bold = True
    run.font.size = Pt(20)

    cell = table.cell(0, 2)
    para = cell.paragraphs[0]
    # para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = para.add_run("Rev. No. :  ")
    run.bold = True
    run.font.size = Pt(12)
    run = para.add_run(revision_var.get())
    run.font.size = Pt(12)

    # Merge cell(1,0) and cell(2,0)
    cell_to_merge_1 = table.cell(1, 0)
    cell_to_merge_2 = table.cell(2, 0)

    merged_cell = cell_to_merge_1.merge(cell_to_merge_2)
    merged_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    para = merged_cell.paragraphs[0]
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run  = para.add_run("EF-RND-J-022")
    run.bold = True
    run.font.size = Pt(12)

    cell = table.cell(1, 1)
    para = cell.paragraphs[0]
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = para.add_run("Title :  ")
    run.bold = True
    run = para.add_run("Display Parameter")
    run.font.size = Pt(12)

    cell = table.cell(1, 2)   # Row=1, Col=2
    para = cell.paragraphs[0]
    # para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = para.add_run("Date: ")
    run.bold = True

    # Add selected date from calendar
    run = para.add_run(date_var.get())
    run.font.size = Pt(12)

    
    cell = table.cell(2, 1)
    para = cell.paragraphs[0]
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = para.add_run("Reference :  ")
    run.bold = True
    run = para.add_run("RND Procedure")
    run.font.size = Pt(12)
    
    cell = table.cell(2, 2)   # Row 2, Column 2
    para = cell.paragraphs[0]
    # para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_page_number(para)
    
    # Add footer (appears on all pages)
    footer = section.footer

    para = footer.paragraphs[0]
    para.alignment = WD_ALIGN_PARAGRAPH.LEFT
    add_page_number(para)

    # doc.save(template_path)

    output_path = output_var.get()
    
    # Set default font (for normal text)
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Calibri'
    font.size = Pt(13)  # Optional: default size

    for i in range(1, 4):  # Assuming you use heading levels 1-3
        style = doc.styles[f'Heading {i}']
        style.font.name = 'Calibri'
        style.font.size = Pt(16 if i==1 else 14)  # Adjust sizes as needed
        style.font.bold = True
        style.font.color.rgb = RGBColor(0, 0, 0)  # Dark Black color

    # doc = Document(template_path)   
    heading = doc.add_heading(f"{meter_type} Display Parameters", level=1)
    heading.paragraph_format.space_before = Pt(0)  # Remove space before
    heading.paragraph_format.space_after = Pt(0)   # Optional: remove space after

    if auto_vals:
        heading1 = doc.add_heading("Table 1: Auto Scroll Display :\n", level=2)
        # Apply underline to the heading text
        run = heading1.runs[0]
        run.underline = True  # Simple underline
        # Or for more options:
        # run.underline = WD_UNDERLINE.SINGLE
        create_table(doc, auto_vals, meter_type)
    
    # Insert a page break
    paragraph = doc.add_paragraph()
    run = paragraph.add_run()
    run.add_break(WD_BREAK.PAGE)  # <-- page break here

    if push_vals:
        heading2 = doc.add_heading("Table 2: Push Button Display :\n", level=2)
        # Apply underline to the heading text
        run = heading2.runs[0]
        run.underline = True  # Simple underline
        # Or for more options:
        # run.underline = WD_UNDERLINE.SINGLE
        create_table(doc, push_vals, meter_type)
    
    # Insert a page break
    paragraph = doc.add_paragraph()
    run = paragraph.add_run()
    run.add_break(WD_BREAK.PAGE)  # <-- page break here

    if diag_vals: 
        heading3 = doc.add_heading("Table 3:  Diagnostic Mode Display :\n", level=2)
        # Apply underline to the heading text
        run = heading3.runs[0]
        run.underline = True  # Simple underline
        # Or for more options:
        # run.underline = WD_UNDERLINE.SINGLE
        create_table(doc, diag_vals, meter_type)
    
    # Insert a page break
    paragraph = doc.add_paragraph()
    run = paragraph.add_run()
    run.add_break(WD_BREAK.PAGE)  # <-- page break here

    if test_vals:
        heading3 = doc.add_heading("Table 4: Test Mode Display :\n", level=2)
        # Apply underline to the heading text
        run = heading3.runs[0]
        run.underline = True  # Simple underline
        # Or for more options:
        # run.underline = WD_UNDERLINE.SINGLE
        create_table(doc, test_vals, meter_type)
       
    doc.add_heading("Note: ", level=3)
    para = doc.add_paragraph("1.  In the Push Mode Display table, present tamper status display parameter was normally hidden and if tamper present continuously for 10 seconds then status showing as above table for respective tamper.")
    para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    para = doc.add_paragraph("2.")  
    para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    run = para.add_run()
    run = para.add_run(" If net-meter mode configured in meter, all forwarded parameters converted into Import parameters and distinguish with ")
    run.add_picture(os.path.join(base_folder, "Template images", "import.jpg"), width=Inches(0.25))
    run = para.add_run(" on display.")
    para = doc.add_paragraph("3.  In the Auto Mode Display and Push Mode Display default here show all 8 zones for zone wise display parameters otherwise meter will show zones wise display parameters as per the configured zones in meter.")
    para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY


    doc.add_heading("\nDisplay Modes : ", level=2)
    para = doc.add_paragraph("Display parameters are displayed in four different modes called Auto scrolling mode, push button mode, diagnostic mode and test mode.")
    para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    doc.add_heading("\nAUTO MODE-- (Auto Scrolling, Refer TABLE-1): ", level=3)
    para = doc.add_paragraph(auto_txt.get())
    para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    doc.add_heading("\nPUSH MODE-- (Scroll with Push Button, Refer TABLE-2): ", level=3)
    para = doc.add_paragraph(push_txt.get())
    para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    # if diag_vals:
    doc.add_heading("\nDIAGNOSTIC MODE-- (High resolution display, Refer TABLE-3): ", level=3)
    para = doc.add_paragraph(diag_txt.get())
    para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    if test_vals:
        doc.add_heading("\nTEST MODE—(Test Mode Display, Refer TABLE-4): ", level=3)
    else:
        doc.add_heading("\nTEST MODE—(Test Mode Display): ", level=3)
    para = doc.add_paragraph(test_txt.get())
    para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    doc.add_heading("\nUnder Power OFF condition: ", level=3)
    para = doc.add_paragraph(batt_txt.get())
    para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    if test_vals:
        # Insert a page break
        paragraph = doc.add_paragraph()
        run = paragraph.add_run()
        run.add_break(WD_BREAK.PAGE)  # <-- page break here
        
        heading3 = doc.add_heading("Table 5: Relay Disconnect Reason (Present Status):\n", level=2)
        test_vals_temp = [3301, 3302, 3303, 3304, 3305]
        # Apply underline to the heading text
        run = heading3.runs[0]
        run.underline = True  # Simple underline
        # Or for more options:
        # run.underline = WD_UNDERLINE.SINGLE
        create_table(doc, test_vals_temp, meter_type)

        # Insert a page break
        paragraph = doc.add_paragraph()
        run = paragraph.add_run()
        run.add_break(WD_BREAK.PAGE)  # <-- page break here

        if 3115 in test_vals:
            heading3 = doc.add_heading("Table 6: Error Details for GPRS :\n", level=2)
            # Apply underline to the heading text
            run = heading3.runs[0]
            run.underline = True  # Simple underline
            # Or for more options:
            # run.underline = WD_UNDERLINE.SINGLE

            table = doc.add_table(rows=20, cols=3)
            table.alignment = WD_TABLE_ALIGNMENT.LEFT
            table.autofit = True
            table.style = 'Table Grid'

            # Set column widths (optional, so it looks neat)
            col_widths = [Inches(1.0), Inches(4.0), Inches(1.5)]
            for col, width in zip(table.columns, col_widths):
                for cell in col.cells:
                    cell.width = width

            cell = table.cell(0, 0)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("Sr. No.")
            run.bold = True
            run.font.size = Pt(14)

            cell = table.cell(0, 1)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("Error Description")
            run.bold = True
            run.font.size = Pt(14)
            
            cell = table.cell(0, 2)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("Meter Display")
            run.bold = True
            run.font.size = Pt(14)
            
            cell = table.cell(1, 0)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("1.")

            cell = table.cell(1, 1)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = para.add_run("All Good")

            cell = table.cell(1, 2)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("Err 00")

            cell = table.cell(2, 0)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("2.")

            cell = table.cell(2, 1)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = para.add_run("Meter NIC communication failure")

            cell = table.cell(2, 2)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("Err 01")

            cell = table.cell(3, 0)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("3.")

            cell = table.cell(3, 1)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = para.add_run("Modem initialization failure")

            cell = table.cell(3, 2)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("Err 02")

            cell = table.cell(4, 0)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("4.")

            cell = table.cell(4, 1)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = para.add_run("SIM not detected")

            cell = table.cell(4, 2)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("Err 03")

            cell = table.cell(5, 0)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("5.")

            cell = table.cell(5, 1)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = para.add_run("SIM invalid")

            cell = table.cell(5, 2)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("Err 04")

            cell = table.cell(6, 0)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("6.")

            cell = table.cell(6, 1)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = para.add_run("No GSM Network coverage")

            cell = table.cell(6, 2)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("Err 05")

            cell = table.cell(7, 0)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("7.")

            cell = table.cell(7, 1)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = para.add_run("GPRS Network registration failure")

            cell = table.cell(7, 2)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("Err 06")

            cell = table.cell(8, 0)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("8.")

            cell = table.cell(8, 1)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = para.add_run("GPRS registration denied")

            cell = table.cell(8, 2)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("Err 07")

            cell = table.cell(9, 0)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("9.")

            cell = table.cell(9, 1)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = para.add_run("No APN configured")

            cell = table.cell(9, 2)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("Err 08")

            cell = table.cell(10, 0)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("10.")

            cell = table.cell(10, 1)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = para.add_run("GPRS connection not establish")

            cell = table.cell(10, 2)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("Err 09")

            cell = table.cell(11, 0)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("11.")

            cell = table.cell(11, 1)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = para.add_run("HES IP/Port not configured")

            cell = table.cell(11, 2)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("Err 10")

            cell = table.cell(12, 0)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("12.")

            cell = table.cell(12, 1)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = para.add_run("HES port not open")

            cell = table.cell(12, 2)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("Err 11")

            cell = table.cell(13, 0)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("13.")

            cell = table.cell(13, 1)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = para.add_run("Internal Diagnostic Purpose")

            cell = table.cell(13, 2)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("Err 12")

            cell = table.cell(14, 0)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("14.")

            cell = table.cell(14, 1)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = para.add_run("Internal Diagnostic Purpose")

            cell = table.cell(14, 2)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("Err 13")

            cell = table.cell(15, 0)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("15.")

            cell = table.cell(15, 1)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = para.add_run("Internal Diagnostic Purpose")

            cell = table.cell(15, 2)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("Err 14")

            cell = table.cell(16, 0)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("16.")

            cell = table.cell(16, 1)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = para.add_run("Internal Diagnostic Purpose")

            cell = table.cell(16, 2)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("Err 15")

            cell = table.cell(17, 0)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("17.")

            cell = table.cell(17, 1)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = para.add_run("Internal Diagnostic Purpose")

            cell = table.cell(17, 2)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("Err 16")

            cell = table.cell(18, 0)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("18.")

            cell = table.cell(18, 1)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = para.add_run("Internal Diagnostic Purpose")

            cell = table.cell(18, 2)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("Err 17")

            cell = table.cell(19, 0)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("19.")

            cell = table.cell(19, 1)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = para.add_run("Internal Diagnostic Purpose")

            cell = table.cell(19, 2)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("Err 18")

        elif 3125 in test_vals:
            heading3 = doc.add_heading("Table 6: Error Details for RF :\n", level=2)
            # Apply underline to the heading text
            run = heading3.runs[0]
            run.underline = True  # Simple underline
            # Or for more options:
            # run.underline = WD_UNDERLINE.SINGLE

            table = doc.add_table(rows=8, cols=3)
            table.alignment = WD_TABLE_ALIGNMENT.LEFT
            table.autofit = True
            table.style = 'Table Grid'

            # Set column widths (optional, so it looks neat)
            col_widths = [Inches(1.0), Inches(4.0), Inches(1.5)]
            for col, width in zip(table.columns, col_widths):
                for cell in col.cells:
                    cell.width = width

            cell = table.cell(0, 0)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("Sr. No.")
            run.bold = True
            run.font.size = Pt(14)

            cell = table.cell(0, 1)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("Error Description")
            run.bold = True
            run.font.size = Pt(14)
            
            cell = table.cell(0, 2)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("Meter Display")
            run.bold = True
            run.font.size = Pt(14)
            
            cell = table.cell(1, 0)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("1.")

            cell = table.cell(1, 1)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = para.add_run("All Good")

            cell = table.cell(1, 2)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("Err 00")

            cell = table.cell(2, 0)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("2.")

            cell = table.cell(2, 1)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = para.add_run("Meter NIC communication failure")

            cell = table.cell(2, 2)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("Err 01")

            cell = table.cell(3, 0)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("3.")

            cell = table.cell(3, 1)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = para.add_run("NIC Not Joined with Gateway")

            cell = table.cell(3, 2)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("Err 02")

            cell = table.cell(4, 0)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("4.")

            cell = table.cell(4, 1)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = para.add_run("Network Joining")

            cell = table.cell(4, 2)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("Err 03")

            cell = table.cell(5, 0)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("5.")

            cell = table.cell(5, 1)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = para.add_run("NIC Joined with Gateway but not connected with HES")

            cell = table.cell(5, 2)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("Err 05")

            cell = table.cell(6, 0)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("6.")

            cell = table.cell(6, 1)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = para.add_run("Channel Change")

            cell = table.cell(6, 2)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("Err 06")

            cell = table.cell(7, 0)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("7.")

            cell = table.cell(7, 1)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = para.add_run("Searching Network")

            cell = table.cell(7, 2)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("Err 07")

        else:
            heading3 = doc.add_heading("Table 6: Error Details for GPRS :\n", level=2)
            # Apply underline to the heading text
            run = heading3.runs[0]
            run.underline = True  # Simple underline
            # Or for more options:
            # run.underline = WD_UNDERLINE.SINGLE

            table = doc.add_table(rows=20, cols=3)
            table.alignment = WD_TABLE_ALIGNMENT.LEFT
            table.autofit = True
            table.style = 'Table Grid'

            # Set column widths (optional, so it looks neat)
            col_widths = [Inches(1.0), Inches(4.0), Inches(1.5)]
            for col, width in zip(table.columns, col_widths):
                for cell in col.cells:
                    cell.width = width

            cell = table.cell(0, 0)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("Sr. No.")
            run.bold = True
            run.font.size = Pt(14)

            cell = table.cell(0, 1)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("Error Description")
            run.bold = True
            run.font.size = Pt(14)
            
            cell = table.cell(0, 2)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("Meter Display")
            run.bold = True
            run.font.size = Pt(14)
            
            cell = table.cell(1, 0)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("1.")

            cell = table.cell(1, 1)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = para.add_run("All Good")

            cell = table.cell(1, 2)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("Err 00")

            cell = table.cell(2, 0)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("2.")

            cell = table.cell(2, 1)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = para.add_run("Meter NIC communication failure")

            cell = table.cell(2, 2)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("Err 01")

            cell = table.cell(3, 0)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("3.")

            cell = table.cell(3, 1)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = para.add_run("Modem initialization failure")

            cell = table.cell(3, 2)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("Err 02")

            cell = table.cell(4, 0)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("4.")

            cell = table.cell(4, 1)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = para.add_run("SIM not detected")

            cell = table.cell(4, 2)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("Err 03")

            cell = table.cell(5, 0)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("5.")

            cell = table.cell(5, 1)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = para.add_run("SIM invalid")

            cell = table.cell(5, 2)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("Err 04")

            cell = table.cell(6, 0)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("6.")

            cell = table.cell(6, 1)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = para.add_run("No GSM Network coverage")

            cell = table.cell(6, 2)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("Err 05")

            cell = table.cell(7, 0)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("7.")

            cell = table.cell(7, 1)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = para.add_run("GPRS Network registration failure")

            cell = table.cell(7, 2)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("Err 06")

            cell = table.cell(8, 0)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("8.")

            cell = table.cell(8, 1)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = para.add_run("GPRS registration denied")

            cell = table.cell(8, 2)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("Err 07")

            cell = table.cell(9, 0)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("9.")

            cell = table.cell(9, 1)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = para.add_run("No APN configured")

            cell = table.cell(9, 2)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("Err 08")

            cell = table.cell(10, 0)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("10.")

            cell = table.cell(10, 1)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = para.add_run("GPRS connection not establish")

            cell = table.cell(10, 2)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("Err 09")

            cell = table.cell(11, 0)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("11.")

            cell = table.cell(11, 1)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = para.add_run("HES IP/Port not configured")

            cell = table.cell(11, 2)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("Err 10")

            cell = table.cell(12, 0)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("12.")

            cell = table.cell(12, 1)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = para.add_run("HES port not open")

            cell = table.cell(12, 2)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("Err 11")

            cell = table.cell(13, 0)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("13.")

            cell = table.cell(13, 1)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = para.add_run("Internal Diagnostic Purpose")

            cell = table.cell(13, 2)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("Err 12")

            cell = table.cell(14, 0)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("14.")

            cell = table.cell(14, 1)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = para.add_run("Internal Diagnostic Purpose")

            cell = table.cell(14, 2)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("Err 13")

            cell = table.cell(15, 0)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("15.")

            cell = table.cell(15, 1)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = para.add_run("Internal Diagnostic Purpose")

            cell = table.cell(15, 2)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("Err 14")

            cell = table.cell(16, 0)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("16.")

            cell = table.cell(16, 1)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = para.add_run("Internal Diagnostic Purpose")

            cell = table.cell(16, 2)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("Err 15")

            cell = table.cell(17, 0)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("17.")

            cell = table.cell(17, 1)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = para.add_run("Internal Diagnostic Purpose")

            cell = table.cell(17, 2)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("Err 16")

            cell = table.cell(18, 0)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("18.")

            cell = table.cell(18, 1)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = para.add_run("Internal Diagnostic Purpose")

            cell = table.cell(18, 2)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("Err 17")

            cell = table.cell(19, 0)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("19.")

            cell = table.cell(19, 1)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = para.add_run("Internal Diagnostic Purpose")

            cell = table.cell(19, 2)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("Err 18")

            heading3 = doc.add_heading("Table 7: Error Details for RF :\n", level=2)
            # Apply underline to the heading text
            run = heading3.runs[0]
            run.underline = True  # Simple underline
            # Or for more options:
            # run.underline = WD_UNDERLINE.SINGLE

            table = doc.add_table(rows=8, cols=3)
            table.alignment = WD_TABLE_ALIGNMENT.LEFT
            table.autofit = True
            table.style = 'Table Grid'

            # Set column widths (optional, so it looks neat)
            col_widths = [Inches(1.0), Inches(4.0), Inches(1.5)]
            for col, width in zip(table.columns, col_widths):
                for cell in col.cells:
                    cell.width = width

            cell = table.cell(0, 0)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("Sr. No.")
            run.bold = True
            run.font.size = Pt(14)

            cell = table.cell(0, 1)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("Error Description")
            run.bold = True
            run.font.size = Pt(14)
            
            cell = table.cell(0, 2)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("Meter Display")
            run.bold = True
            run.font.size = Pt(14)
            
            cell = table.cell(1, 0)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("1.")

            cell = table.cell(1, 1)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = para.add_run("All Good")

            cell = table.cell(1, 2)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("Err 00")

            cell = table.cell(2, 0)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("2.")

            cell = table.cell(2, 1)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = para.add_run("Meter NIC communication failure")

            cell = table.cell(2, 2)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("Err 01")

            cell = table.cell(3, 0)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("3.")

            cell = table.cell(3, 1)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = para.add_run("NIC Not Joined with Gateway")

            cell = table.cell(3, 2)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("Err 02")

            cell = table.cell(4, 0)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("4.")

            cell = table.cell(4, 1)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = para.add_run("Network Joining")

            cell = table.cell(4, 2)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("Err 03")

            cell = table.cell(5, 0)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("5.")

            cell = table.cell(5, 1)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = para.add_run("NIC Joined with Gateway but not connected with HES")

            cell = table.cell(5, 2)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("Err 05")

            cell = table.cell(6, 0)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("6.")

            cell = table.cell(6, 1)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = para.add_run("Channel Change")

            cell = table.cell(6, 2)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("Err 06")

            cell = table.cell(7, 0)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("7.")

            cell = table.cell(7, 1)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = para.add_run("Searching Network")

            cell = table.cell(7, 2)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run("Err 07")




    doc.add_heading("\nLCD Glass Description: ", level=3)
    para = doc.add_picture(os.path.join(base_folder, "Template images", "LCD_Glass_Description_3P-Pro.png"), width=Inches(5.0))
    # Center align the picture
    last_paragraph = doc.paragraphs[-1]  # the paragraph where the picture was added
    last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    doc.add_heading("\nIcons on LCD Glass:- ", level=3) 
    para1 = doc.add_paragraph("1. ")
    run = para1.add_run()
    run.add_picture(os.path.join(base_folder, "Template images", "magnet.jpg"), width=Inches(0.25))
    para1.add_run(" Sign indicates the instantaneous presence of Magnet Tamper.")

    para2 = doc.add_paragraph("2.  ")
    run = para2.add_run()
    run.add_picture(os.path.join(base_folder, "Template images", "comm.jpg"), width=Inches(0.25))
    para2.add_run(" Blinking sign indicates the communication is going on with meter.")

    para3 = doc.add_paragraph("3.  ")
    run = para3.add_run()
    para3.add_run("1 2 3 ").bold = True
    para3.add_run(" sign are the voltage and current indication for R phase, Y phase and B phase, respectively. Continuous display of ")
    para3.add_run("1 2 3 ").bold = True
    para3.add_run(" shows the presence of voltage and current in respective phase.Blinking of ")
    para3.add_run("1 2 3 ").bold = True
    para3.add_run(" shows absence of current and presence of voltage in the respective phase. The absence of ")
    para3.add_run("1 2 3 ").bold = True
    para3.add_run(" shows the absence of voltage in respective phase.")

    para4 = doc.add_paragraph("4.  ")
    run = para4.add_run()
    run.add_picture(os.path.join(base_folder, "Template images", "reverse_arrow.jpg"), width=Inches(0.3))
    para4.add_run(" Sign indicates reverse current.")

    para5 = doc.add_paragraph("5.  ")
    run = para5.add_run()
    run.add_picture(os.path.join(base_folder, "Template images", "battery.jpg"), width=Inches(0.35))
    para5.add_run(" Sign indicates internal backup low battery detection.")

    para6 = doc.add_paragraph("6.  ")
    run = para6.add_run()
    run.add_picture(os.path.join(base_folder, "Template images", "lock.jpg"), width=Inches(0.2))
    run = para6.add_run(" Sign indicates top cover open tamper and module cover open tamper present in meter. Continuous display of  ")
    run.add_picture(os.path.join(base_folder, "Template images", "lock.jpg"), width=Inches(0.2))
    run = para6.add_run(" shows top cover open present in meter. Blinking of  ")
    run.add_picture(os.path.join(base_folder, "Template images", "lock.jpg"), width=Inches(0.2))
    run = para6.add_run(" shows the module cover open tamper present in the meter. If Top cover and module cover both open then the top cover priority will be higher (Continuous display of  ")
    run.add_picture(os.path.join(base_folder, "Template images", "lock.jpg"), width=Inches(0.2))
    para6.add_run(" ).")

    para7 = doc.add_paragraph("7.  ")
    run = para7.add_run()
    run.add_picture(os.path.join(base_folder, "Template images", "bell.jpg"), width=Inches(0.3))
    para7.add_run(" Sign indicates phase sequence reverse in meter.")

    if(meter_type == "3P-Pro WC 10-60A" or meter_type == "3P-Pro WC 20-100A"):
        para8 = doc.add_paragraph("8.  ")
        run = para8.add_run()
        run.add_picture(os.path.join(base_folder, "Template images", "relay.jpg"), width=Inches(0.3))
        para8.add_run(" Sign indicates relay connect/disconnect status.")


    # === Save DOCX ===
    doc.save(output_path)
    messagebox.showinfo("Success", f"Report generated Word file created \n{output_path}")

    # # === Save DOCX ===
    # doc.save(output_path)
    # pdf_file_path = output_path.replace(".docx", ".pdf")
    
    # try:
    #     # Try using Pandoc first
    #     if pandoc_installed:
    #         try:
    #             pypandoc.convert_file(
    #                 output_path,
    #                 "pdf",
    #                 outputfile=pdf_file_path,
    #                 extra_args=["--standalone"]
    #             )
    #             messagebox.showinfo("Success", f"Report generated using Pandoc:\n\nWord: {output_path}\nPDF: {pdf_file_path}")
    #         except Exception as pandoc_error:
    #             # Pandoc failed; fallback to docx2pdf
    #             if docx2pdf_installed:
    #                 docx2pdf_convert(output_path, pdf_file_path)
    #                 messagebox.showinfo("Success", f"Pandoc failed but PDF generated using docx2pdf:\n\nWord: {output_path}\nPDF: {pdf_file_path}")
    #             else:
    #                 raise RuntimeError(f"PDF conversion failed with Pandoc:\n{pandoc_error}\nAlso, docx2pdf is not installed.")
    #     # If Pandoc not installed, try docx2pdf
    #     elif docx2pdf_installed:
    #         docx2pdf_convert(output_path, pdf_file_path)
    #         messagebox.showinfo("Success", f"PDF generated using docx2pdf:\n\nWord: {output_path}\nPDF: {pdf_file_path}")
    #     else:
    #         raise RuntimeError("Neither Pandoc nor docx2pdf is installed. Cannot generate PDF.")

    # except Exception as e:
    #     messagebox.showerror("PDF Conversion Failed", f"Word file created but PDF failed:\n{e}")

    # # === Save PDF ===
    # try:
    #     pdf_path = output_path.replace(".docx", ".pdf")
    #     convert(output_path, pdf_path)
    #     messagebox.showinfo("Success", f"Report generated:\n\n{output_path}\n{pdf_path}")
    # except Exception as e:
    #     messagebox.showerror("PDF Conversion Failed", f"Word file created but PDF failed:\n{e}")


# === Tkinter GUI ===
ctk.set_appearance_mode("light")   # Options: "light", "dark", "system"
ctk.set_default_color_theme("blue")  # Try "green", "dark-blue", etc.
# === Main Window ===
root = ctk.CTk()
root.iconbitmap("logo.ico")
root.title("3P Pro Meter Display Report Generator V1.0")
root.geometry("1200x760")
root.resizable(False, False)

# === Top Frame ===
top_frame = ctk.CTkFrame(root, corner_radius=15)
top_frame.pack(pady=15, padx=15, fill="x")

ctk.CTkLabel(top_frame, text="Base Folder:", text_color="#1a1a1a").grid(row=0, column=0, padx=5, sticky="w")
ctk.CTkButton(top_frame, text="Browse Base Folder", fg_color="#1f6aa5", hover_color="#155a8a", command= choose_base_folder).grid(row=0, column=1, padx=5, sticky="w")
ctk.CTkLabel(top_frame, text="Meter Type:").grid(row=0, column=2, padx=20, sticky="w")
meter_type_var = tk.StringVar(value="Select Type")
meter_type_combo = ctk.CTkComboBox(top_frame, variable=meter_type_var, values=[])
meter_type_combo.grid(row=0, column=3, padx=5, sticky="w")

# Date input with calendar picker (right side)
ctk.CTkLabel(top_frame, text="Date:").grid(row=0, column=4, padx=30, sticky="w")
date_var = tk.StringVar()
date_entry = DateEntry(top_frame, textvariable=date_var, date_pattern="dd-mm-yyyy", width=30)
date_entry.grid(row=0, column=5, padx=5, sticky="w")

# Revision Number (right side)
ctk.CTkLabel(top_frame, text="Rev. No:").grid(row=0, column=6, padx=40, sticky="w")
revision_var = tk.StringVar(value="00")
revision_combo = ctk.CTkComboBox(top_frame, variable=revision_var, values=["00", "01", "02", "03"], width=80)
revision_combo.grid(row=0, column=7, padx=5, sticky="w")


# === Dual listbox UI ===
def create_dual_listbox(frame, title, items):
    ttk.Label(frame, text=title, font=("Arial", 16, "bold")).grid(row=0, column=0, columnspan=3, pady=5)

    # --- Search Entry ---
    search_var = tk.StringVar()
    search_entry = ttk.Entry(frame, textvariable=search_var, width=60)
    search_entry.grid(row=1, column=0, padx=10, pady=5)

    ttk.Label(frame, text="Input String:").grid(row=1, column=1, padx=5, sticky="e")
    
    input_hex_string = tk.StringVar()
    input_string = ttk.Entry(frame, textvariable=input_hex_string, width=90)
    input_string.grid(row=1, column=2, padx=10, pady=5)
    input_string.insert(0, "")
    

    # --- Available Listbox + Scrollbar ---
    available_frame = tk.Frame(frame)
    available_frame.grid(row=2, column=0, padx=10, pady=5)
    available = tk.Listbox(available_frame, selectmode=tk.SINGLE, width=70, height=12)
    available_scroll = tk.Scrollbar(available_frame, orient=tk.VERTICAL, command=available.yview)
    available.config(yscrollcommand=available_scroll.set)
    available.grid(row=0, column=0, sticky="nsew")
    available_scroll.grid(row=0, column=1, sticky="ns")

    # --- Selected Listbox + Scrollbar ---
    selected_frame = tk.Frame(frame)
    selected_frame.grid(row=2, column=2, padx=10, pady=5)
    selected = tk.Listbox(selected_frame, selectmode=tk.SINGLE, width=70, height=12)
    selected_scroll = tk.Scrollbar(selected_frame, orient=tk.VERTICAL, command=selected.yview)
    selected.config(yscrollcommand=selected_scroll.set)
    selected.grid(row=0, column=0, sticky="nsew")
    selected_scroll.grid(row=0, column=1, sticky="ns")

    # --- Populate Available ---
    filtered_items = items.copy()
    for item in filtered_items:
        available.insert(tk.END, item)

    # --- Button Functions ---
    def add_item():
        sel = available.curselection()
        if sel:
            idx = sel[0]
            text = available.get(idx)
            if text not in selected.get(0, tk.END):
                selected.insert(tk.END, text)
            # Move selection to next item
            next_idx = idx + 1
            available.selection_clear(0, tk.END)
            if next_idx < available.size():
                available.selection_set(next_idx)
                available.see(next_idx)

    def remove_item():
        sel = selected.curselection()
        if sel:
            selected.delete(sel[0])

    def move_up():
        sel = selected.curselection()
        if sel and sel[0] > 0:
            idx = sel[0]
            item = selected.get(idx)
            selected.delete(idx)
            selected.insert(idx - 1, item)
            selected.select_set(idx - 1)

    def move_down():
        sel = selected.curselection()
        if sel and sel[0] < selected.size() - 1:
            idx = sel[0]
            item = selected.get(idx)
            selected.delete(idx)
            selected.insert(idx + 1, item)
            selected.select_set(idx + 1)

    # --- Buttons Frame ---
    btn_frame = tk.Frame(frame)
    btn_frame.grid(row=2, column=1)
    ttk.Button(btn_frame, text=">>", command=add_item).pack(pady=5)
    ttk.Button(btn_frame, text="🗑", command=remove_item).pack(pady=5)
    ttk.Button(btn_frame, text="↑", command=move_up).pack(pady=5)
    ttk.Button(btn_frame, text="↓", command=move_down).pack(pady=5)

    # --- Search Function ---
    def update_filter(*args):
        search_term = search_var.get().lower()
        available.delete(0, tk.END)
        for item in filtered_items:
            if search_term in str(item).lower():
                available.insert(tk.END, item)

    search_var.trace_add("write", update_filter)

    return available, selected, input_hex_string



# Load default items from predefined_names
default_items = [f"{dec:04d} - {desc}" for dec, desc in predefined_names_new.items()]
diag_items = [f"{dec:04d} - {desc}" for dec, desc in diag_predefined_names_new.items()]
test_items = [f"{dec:04d} - {desc}" for dec, desc in test_predefined_names_new.items()]

# Create a frame with fixed size for the notebook
notebook_frame = tk.Frame(root, width=1080, height=350)
notebook_frame.pack(padx=10, pady=(0,5))  # adjust pady to reduce gap

# Prevent frame from shrinking to fit notebook
notebook_frame.pack_propagate(False)

# Tabs for Auto / Push
notebook = ttk.Notebook(notebook_frame)
notebook.pack(pady=20, expand=True, fill="both")

auto_tab = ttk.Frame(notebook)
notebook.add(auto_tab, text="Auto Scroll Selection")
auto_available, auto_selected, auto_input_hex_string = create_dual_listbox(auto_tab, "Auto Scroll Selection", default_items)
# frame_auto = tk.LabelFrame(root, text="Auto Mode")
# frame_auto.pack(pady=10, fill="x")
# auto_available, auto_selected = create_dual_listbox(frame_auto, "Auto Scroll Selection", default_items)

push_tab = ttk.Frame(notebook)
notebook.add(push_tab, text="Push Scroll Selection")
push_available, push_selected, push_input_hex_string  = create_dual_listbox(push_tab, "Push Scroll Selection", default_items)
# frame_push = tk.LabelFrame(root, text="Push Mode")
# frame_push.pack(pady=10, fill="x")
# push_available, push_selected = create_dual_listbox(frame_push, "Push Scroll Selection", default_items)

diag_tab = ttk.Frame(notebook)
notebook.add(diag_tab, text="Diagnostic Selection")
diag_available, diag_selected, diag_input_hex_string  = create_dual_listbox(diag_tab, "Diagnostic Selection", diag_items)

test_tab = ttk.Frame(notebook)
notebook.add(test_tab, text="Test Mode Selection")
test_available, test_selected, test_input_hex_string  = create_dual_listbox(test_tab, "Test Mode Selection", test_items)

# Output file
output_var = tk.StringVar(value=os.path.join(base_folder, "Display_Parameters_report.docx"))

auto_txt = tk.StringVar(value="These parameters displayed on the LCD continuously with auto scrolling. All auto mode Parameters shall display for 10 seconds with auto scrolling.")

push_txt = tk.StringVar(value="This mode will activate when pressing the Scroll up or Scroll down push button. Meter will continuously scroll through Push Button as Display parameters (Refer TABLE-2). If push button not pressed for 10 seconds then display will switch over to AUTO display mode.")

diag_txt = tk.StringVar(value="This mode will activate when both Scroll UP & Scroll Down button will press then released & continuously scroll through push button as diagnostics mode display parameters (Refer TABLE-3). The diagnostic mode display will switch over to AUTO display mode when buttons (scroll up/scroll down) not pressed for 15 minutes or again press and released the both buttons.")

test_txt = tk.StringVar(value="Test mode can be enabled by continuously pressing the scroll down button for 8 seconds then released. The test mode options are accessed by navigating the options using the push buttons. The Test mode display will switch over to AUTO display mode when buttons (scroll up/scroll down) is not pressed for 1 minutes or at 'exit' parameter, pressing scroll up button for 4 sec then release it. Test mode parameters used for diagnostic purpose only.")

batt_txt = tk.StringVar(value="Meter display parameters with the help of internal battery activated by pressing the scroll down button and parameter shall scroll one after another using press of push buttons, if push button left unattended, then the display switch-off after 20 seconds.")


# Create a frame for output file controls
input_frame = ttk.Frame(root)
input_frame.pack(pady=20, padx=10, fill="x")   # This uses pack()

ttk.Label(input_frame, text="Auto Mode:").grid(row=0, column=0, padx=5, sticky="w")
ttk.Entry(input_frame, textvariable=auto_txt, width=170).grid(row=0, column=1, padx=30, sticky="w")

ttk.Label(input_frame, text="Push Mode:").grid(row=1, column=0, padx=5, sticky="w")
ttk.Entry(input_frame, textvariable=push_txt, width=170).grid(row=1, column=1, padx=30, sticky="w")

ttk.Label(input_frame, text="Diag. Mode:").grid(row=2, column=0, padx=5, sticky="w")
ttk.Entry(input_frame, textvariable=diag_txt, width=170).grid(row=2, column=1, padx=30, sticky="w")

ttk.Label(input_frame, text="Test Mode:").grid(row=3, column=0, padx=5, sticky="w")
ttk.Entry(input_frame, textvariable=test_txt, width=170).grid(row=3, column=1, padx=30, sticky="w")

ttk.Label(input_frame, text="Battery Mode:").grid(row=4, column=0, padx=5, sticky="w")
ttk.Entry(input_frame, textvariable=batt_txt, width=170).grid(row=4, column=1, padx=30, sticky="w")

# Create a frame for output file controls
output_frame = ctk.CTkFrame(root, corner_radius=15)
output_frame.pack(pady=5, fill="x")   # This uses pack()

ttk.Label(output_frame, text="Output File:").grid(row=0, column=0, padx=5, sticky="w")
ttk.Entry(output_frame, textvariable=output_var, width=130).grid(row=0, column=1, padx=30, sticky="w")
ctk.CTkButton(output_frame, text="Browse", command=choose_output, fg_color="#2196f3").grid(row=0, column=2, padx=30, sticky="w")

# Generate button
ctk.CTkButton(output_frame, text="Generate Report", command=generate_doc, fg_color="#388e3c").grid(row=2, column=2, padx=30, sticky="w")

auto_string = tk.StringVar()

ctk.CTkLabel(output_frame, text="").grid(row=1, column=3, padx=5, sticky="w")

ttk.Label(output_frame, text="Auto Display String:").grid(row=2, column=0, padx=5, sticky="w")
ttk.Entry(output_frame, textvariable=auto_string, width=130).grid(row=2, column=1, padx=30, sticky="e")

push_string = tk.StringVar()

ttk.Label(output_frame, text="Push Display String:").grid(row=3, column=0, padx=5, sticky="w")
ttk.Entry(output_frame, textvariable=push_string, width=130).grid(row=3, column=1, padx=30, sticky="e")

diag_string = tk.StringVar()

ttk.Label(output_frame, text="Diag Display String:").grid(row=4, column=0, padx=5, sticky="w")
ttk.Entry(output_frame, textvariable=diag_string, width=130).grid(row=4, column=1, padx=30, sticky="e")

test_string = tk.StringVar()

ttk.Label(output_frame, text="Test Display String:").grid(row=5, column=0, padx=5, sticky="w")
ttk.Entry(output_frame, textvariable=test_string, width=130).grid(row=5, column=1, padx=30, sticky="e")


# === Footer ===
footer_frame = ctk.CTkFrame(root, height=30)
footer_frame.pack(side="bottom", fill="x")

footer_label = ctk.CTkLabel(
    footer_frame,
    text="Disclaimer - This is internal property of Genus Power Infrastructure Limited. Do not use without permission of developers.                                CopyRight - abhishek.Verma@genus.in (All Rights Reserved)"
)
footer_label.pack(pady=5)

root.mainloop()
