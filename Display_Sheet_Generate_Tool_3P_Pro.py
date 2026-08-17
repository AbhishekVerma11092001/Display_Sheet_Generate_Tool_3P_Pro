import os
import sys
import glob
import customtkinter as ctk
import tkinter as tk
import xml.etree.ElementTree as ET
from tkcalendar import DateEntry
from tkinter import ttk, filedialog, simpledialog, messagebox
from docx import Document
from docx.shared import Cm, Inches, Pt, RGBColor
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_BREAK
from datetime import datetime
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import subprocess
import runpy


def run_embedded_script(script_name):
    """Run a Python child application bundled inside the same PyInstaller EXE."""
    if getattr(sys, "frozen", False):
        # PyInstaller --onefile extracts --add-data files to _MEIPASS.
        script_path = os.path.join(sys._MEIPASS, script_name)
        working_dir = os.path.dirname(sys.executable)
    else:
        script_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), script_name)
        working_dir = os.path.dirname(script_path)

    if not os.path.exists(script_path):
        messagebox.showerror(
            "Error",
            f"Embedded file not found:\n{script_path}"
        )
        return

    try:
        old_argv = sys.argv[:]
        old_cwd = os.getcwd()
        sys.argv = [script_path]
        os.chdir(working_dir)
        try:
            runpy.run_path(script_path, run_name="__main__")
        finally:
            sys.argv = old_argv
            os.chdir(old_cwd)
    except Exception as e:
        messagebox.showerror(
            "Error",
            f"Unable to open {script_name}:\n\n{e}"
        )


def handle_embedded_mode():
    """Start the requested child GUI when the main EXE is launched with a mode argument."""
    if len(sys.argv) < 2:
        return False

    mode_map = {
        "--profile-object-selection": "Profile_Object_Selection.py",
        "--top-level-string-selection": "Top_Level_String_Selection.py",
    }

    script_name = mode_map.get(sys.argv[1])
    if not script_name:
        return False

    run_embedded_script(script_name)
    return True


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

# If this EXE was launched in an embedded-child mode, run that child GUI
# and stop here so the main Display Sheet GUI is not created as well.
if handle_embedded_mode():
    sys.exit(0)


try:
    import pypandoc
    pypandoc_installed = True
except ImportError:
    pypandoc_installed = False

try:
    from docx2pdf import convert as docx2pdf_convert
    docx2pdf_installed = True
except ImportError:
    docx2pdf_installed = False

# # === Predefined mapping of decimal value to description ===

def read_dict_from_section(root, section):
    data_dict = {}
    section_node = root.find(section)
    if section_node is None:
        return data_dict

    for item in section_node.findall("Item"):
        item_id = item.get("id")
        if item_id is None:
            continue
        try:
            key = int(item_id)
        except ValueError:
            continue
        value = (item.text or "").strip()
        data_dict[key] = value
    return data_dict


def read_text_mapping_from_file(xml_path, section="FORWARDDisplayTexts"):
    if not os.path.exists(xml_path):
        return {}

    try:
        tree = ET.parse(xml_path)
        root = tree.getroot()
    except ET.ParseError:
        return {}

    return read_dict_from_section(root, section)


def load_xml_mappings():
    global predefined_names_new, diag_predefined_names_new, test_predefined_names_new, tamper_predefined_names_new, forward_display_text_mapping, net_display_text_mapping

    xml_path = os.path.join(base_folder, "RefXML",  "predefined_names_3p_pro.xml")
    display_text_xml_path = os.path.join(base_folder, "RefXML", "display_parameter_description.xml")

    if os.path.exists(xml_path):
        tree = ET.parse(xml_path)
        root = tree.getroot()
        predefined_names_new = read_dict_from_section(root, "PredefinedNames")
        diag_predefined_names_new = read_dict_from_section(root, "DiagPredefinedNames")
        test_predefined_names_new = read_dict_from_section(root, "TestPredefinedNames")
        tamper_predefined_names_new = read_dict_from_section(root, "TamperPredefinedNames")
    else:
        predefined_names_new = {}
        diag_predefined_names_new = {}
        test_predefined_names_new = {}
        tamper_predefined_names_new = {}

    if os.path.exists(display_text_xml_path):
        forward_display_text_mapping = read_text_mapping_from_file(display_text_xml_path, "FORWARDDisplayTexts")
        net_display_text_mapping = read_text_mapping_from_file(display_text_xml_path, "NETDisplayTexts")


# Determine app folder and default base folder
if getattr(sys, 'frozen', False):
    base_folder = os.path.dirname(sys.executable)
else:
    base_folder = os.path.dirname(os.path.abspath(__file__))

# Load XML relative to exe/script location
load_xml_mappings()


DEFAULT_OUTPUT_FILE = "Display_Parameters_report.docx"
DISPLAY_STRINGS_XML = "display_strings.xml"

# Global variables for folder paths
image_folders = {
    "3P-Pro WC 10-60A": os.path.join(base_folder, "3P Pro WC Display Forward Mode"),
    "3P-Pro WC (Torrent Specific) 10-60A": os.path.join(base_folder, "3P Pro WC Display Torrent Specific Forward Mode"),
    "3P Pro LTCT 5-10A": os.path.join(base_folder, "3P Pro LTCT Display Forward Mode")
}

image_folders_net = {
    "3P-Pro WC 10-60A": os.path.join(base_folder, "3P Pro WC Display NET Mode"),
    "3P-Pro WC (Torrent Specific) 10-60A": os.path.join(base_folder, "3P Pro WC Display Torrent Specific NET Mode"),
    # "3P-Pro WC 20-100A" : os.path.join(base_folder, "3P Pro WC Display"),
    "3P Pro LTCT 5-10A": os.path.join(base_folder, "3P Pro LTCT Display NET Mode")
}

def get_display_strings_xml_path():
    return os.path.join(base_folder, "RefXML", DISPLAY_STRINGS_XML)


def load_display_strings():
    xml_path = get_display_strings_xml_path()
    if not os.path.exists(xml_path):
        return {}

    try:
        tree = ET.parse(xml_path)
        root = tree.getroot()
        return {
            "auto_string": root.findtext("auto_string") or "",
            "push_string": root.findtext("push_string") or "",
            "tamper_string": root.findtext("tamper_string") or "",
            "batt_string": root.findtext("batt_string") or "",
            "diag_string": root.findtext("diag_string") or "",
            "test_string": root.findtext("test_string") or "",
        }
    except ET.ParseError:
        return {}


def save_display_strings(strings):
    xml_path = get_display_strings_xml_path()
    root = ET.Element("DisplayStrings")

    for key, value in strings.items():
        elem = ET.SubElement(root, key)
        elem.text = value or ""

    tree = ET.ElementTree(root)
    tree.write(xml_path, encoding="utf-8", xml_declaration=True)


def set_column_widths(table, widths):
    for col_idx, width in enumerate(widths):
        col = table.columns[col_idx]
        col.width = width
        for cell in col.cells:
            cell.width = width

        for cell in col.cells:
            tc_pr = cell._tc.get_or_add_tcPr()
            tc_w = tc_pr.find(qn('w:tcW'))
            if tc_w is not None:
                tc_pr.remove(tc_w)
            new_tc_w = OxmlElement('w:tcW')
            new_tc_w.set(qn('w:w'), str(int(width.inches * 1440)))
            new_tc_w.set(qn('w:type'), 'dxa')
            tc_pr.append(new_tc_w)


def add_underlined_heading(doc, text, level=2):
    heading = doc.add_heading(text, level=level)
    if heading.runs:
        heading.runs[0].underline = True
    return heading


def add_page_break(doc):
    paragraph = doc.add_paragraph()
    paragraph.add_run().add_break(WD_BREAK.PAGE)


def get_selected_values(listbox):
    return [int(item.split(" - ")[0]) for item in listbox.get(0, tk.END)]


def build_hex_payload(values, replacement_payload=""):
    parts = []
    for val in values:
        if val == 9999:
            if replacement_payload.startswith("09"):
                replacement_payload = replacement_payload[2:]  # Remove the 09 prefix
            
                if len(replacement_payload) >= 2:
                    if replacement_payload.startswith("81"):
                        replacement_payload = replacement_payload[4:]  # Remove the 81 and Length prefix
                    elif replacement_payload.startswith("82"):
                        replacement_payload = replacement_payload[6:]  # Remove the 82 and Length prefix
                    else:
                        replacement_payload = replacement_payload[2:]  # Remove the Length prefix
            parts.append(replacement_payload)
        else:
            parts.append(f"{val:04X}")

    payload = "".join(parts)
    length = len(payload) // 2

    if length <= 127:
        length_prefix = f"{length:02X}"
    elif length <= 255:
        length_prefix = "81" + f"{length:02X}"
    else:
        length_prefix = "82" + f"{length:04X}"

    return "09" + length_prefix + payload


def validate_hex_input_string(value, label):
    if not value:
        return True
    if len(value.strip()) % 4 != 0:
        messagebox.showerror("Error", f"Input {label} hex string is invalid. Check length.")
        return False
    return True


def parse_payload_to_values(payload, available_items):
    if not payload:
        return []

    payload = payload.strip()
    if payload.startswith("09"):
        payload = payload[2:]  # Remove the 09 prefix

    if len(payload) >= 2:
        if payload.startswith("81"):
            payload = payload[2:]  # Remove the 81 prefix
            length_bytes = payload[:2]
        elif payload.startswith("82"):
            payload = payload[2:]  # Remove the 82 prefix
            length_bytes = payload[:4]
        else:
            length_bytes = payload[:2]

        try:
            length = int(length_bytes, 16)
        except ValueError:
            return []

        if length <= 127:
            body = payload[2:2 + length * 2]
            body = payload[2:]  # Remove the Length
        elif length <= 255:
            body = payload[2:2 + length * 2]
            body = payload[2:]  # Remove the Length
        else:
            body = payload[4:4 + length * 2]
            body = payload[4:]  # Remove the Length
    else:
        body = ""

    values = []
    for i in range(0, len(body), 4):
        chunk = body[i:i+4]
        if len(chunk) != 4:
            continue
        try:
            values.append(int(chunk, 16))
        except ValueError:
            continue

    if not available_items:
        return values

    available_map = {f"{value:04d}": item for value, item in [(int(item.split(" - ", 1)[0]), item) for item in available_items]}
    result = []
    for value in values:
        item = available_map.get(f"{value:04d}")
        if item:
            result.append(item)
    return result


def extract_payload_body(payload):
    if not payload:
        return ""

    payload = payload.strip()
    if payload.startswith("09"):
        payload = payload[2:]  # Remove the 09 prefix

    if len(payload) >= 2:
        if payload.startswith("81"):
            payload = payload[2:]  # Remove the 81 prefix
            length_bytes = payload[:2]
        elif payload.startswith("82"):
            payload = payload[2:]  # Remove the 82 prefix
            length_bytes = payload[:4]
        else:
            length_bytes = payload[:2]

        try:
            length = int(length_bytes, 16)
        except ValueError:
            return []

    if length <= 127:
        body_start = 2
    elif length <= 255:
        body_start = 2
    else:
        body_start = 4

    if length <= 127:
        body_end = body_start + length * 2
    elif length <= 255:
        body_end = body_start + length * 2
    else:
        body_end = body_start + length * 2

    return payload[body_start:body_end]


def add_error_details_table(doc, title, rows):
    add_underlined_heading(doc, title, level=2)
    table = doc.add_table(rows=1, cols=3)
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    table.autofit = True
    table.style = 'Table Grid'

    set_column_widths(table, [Inches(1.0), Inches(4.0), Inches(1.5)])

    hdr_cells = table.rows[0].cells
    headers = ["Sr. No.", "Error Description", "Meter Display"]
    for cell, text in zip(hdr_cells, headers):
        run = cell.paragraphs[0].add_run(text)
        run.bold = True
        run.font.size = Pt(14)
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    for number, description, display in rows:
        row = table.add_row().cells
        row[0].text = number
        row[1].text = description
        row[2].text = display
        row[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        row[2].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    return table

GPRS_ERROR_ROWS = [
    ("1.", "All Good", "Err 00"),
    ("2.", "Meter NIC communication failure", "Err 01"),
    ("3.", "Modem initialization failure", "Err 02"),
    ("4.", "SIM not detected", "Err 03"),
    ("5.", "SIM invalid", "Err 04"),
    ("6.", "No GSM Network coverage", "Err 05"),
    ("7.", "GPRS Network registration failure", "Err 06"),
    ("8.", "GPRS registration denied", "Err 07"),
    ("9.", "No APN configured", "Err 08"),
    ("10.", "GPRS connection not establish", "Err 09"),
    ("11.", "HES IP/Port not configured", "Err 10"),
    ("12.", "HES port not open", "Err 11"),
    ("13.", "Internal Diagnostic Purpose", "Err 12"),
    ("14.", "Internal Diagnostic Purpose", "Err 13"),
    ("15.", "Internal Diagnostic Purpose", "Err 14"),
    ("16.", "Internal Diagnostic Purpose", "Err 15"),
    ("17.", "Internal Diagnostic Purpose", "Err 16"),
    ("18.", "Internal Diagnostic Purpose", "Err 17"),
    ("19.", "Internal Diagnostic Purpose", "Err 18"),
    ("20.", "Internal Diagnostic Purpose", "Err 98"),
]

RF_ERROR_ROWS = [
    ("1.", "All Good", "Err 00"),
    ("2.", "Meter NIC communication failure", "Err 01"),
    ("3.", "NIC Not Joined with Gateway", "Err 02"),
    ("4.", "Network Joining", "Err 03"),
    ("5.", "NIC Joined with Gateway but not connected with HES", "Err 05"),
    ("6.", "Channel Change", "Err 06"),
    ("7.", "Searching Network", "Err 07"),
]


def add_table_section(doc, title, values, meter_type, energy_decimal, display_table_mode, batt_var, add_break=True):
    add_underlined_heading(doc, title, level=2)
    create_table(doc, values, meter_type, energy_decimal, display_table_mode, batt_var)
    if add_break:
        add_page_break(doc)

def forward_add_table_section(doc, title, values, meter_type, energy_decimal, display_table_mode, batt_var, add_break=True):
    add_underlined_heading(doc, title, level=2)
    forward_create_table(doc, values, meter_type, energy_decimal, display_table_mode, batt_var)
    if add_break:
        add_page_break(doc)


def insert_error_tables(doc, test_vals, table_seq_no):
    if 3115 in test_vals and 3125 in test_vals:
        add_error_details_table(doc, f"Table {table_seq_no}: Error Details for GPRS :", GPRS_ERROR_ROWS)
        table_seq_no += 1
        add_error_details_table(doc, f"Table {table_seq_no}: Error Details for RF :", RF_ERROR_ROWS)
        return table_seq_no + 1
    elif 3115 in test_vals:
        add_error_details_table(doc, f"Table {table_seq_no}: Error Details for GPRS :", GPRS_ERROR_ROWS)
        return table_seq_no + 1
    elif 3125 in test_vals:
        add_error_details_table(doc, f"Table {table_seq_no}: Error Details for RF :", RF_ERROR_ROWS)
        return table_seq_no + 1
    else:
        return table_seq_no


def get_pdf_path(docx_path):
    root, _ = os.path.splitext(docx_path)
    return root + ".pdf"


def convert_docx_to_pdf(docx_path, pdf_path):
    last_error = None

    if docx2pdf_installed and sys.platform == "win32":
        try:
            docx2pdf_convert(docx_path, pdf_path)
            return "docx2pdf"
        except Exception as exc:
            last_error = exc

    if pypandoc_installed:
        try:
            pypandoc.convert_file(docx_path, "pdf", outputfile=pdf_path)
            return "pypandoc"
        except Exception as exc:
            last_error = exc

    if last_error:
        raise RuntimeError(f"PDF conversion failed: {last_error}")

    raise RuntimeError(
        "No PDF converter is available. Install docx2pdf (Windows) or pypandoc with Pandoc."
    )


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
    global base_folder, image_folders, image_folders_net
    # , template_path

    folder = filedialog.askdirectory(title="Select Base Folder")
    if not folder:
        return
    
    base_folder = folder
    load_xml_mappings()

    image_folders = {
        "3P-Pro WC 10-60A": os.path.join(base_folder, "3P Pro WC Display Forward Mode"),
        "3P-Pro WC (Torrent Specific) 10-60A": os.path.join(base_folder, "3P Pro WC Display Torrent Specific Forward Mode"),
        "3P Pro LTCT 5-10A": os.path.join(base_folder, "3P Pro LTCT Display Forward Mode")
    }

    image_folders_net = {
        "3P-Pro WC 10-60A": os.path.join(base_folder, "3P Pro WC Display NET Mode"),
        "3P-Pro WC (Torrent Specific) 10-60A": os.path.join(base_folder, "3P Pro WC Display Torrent Specific NET Mode"),
        # "3P-Pro WC 20-100A" : os.path.join(base_folder, "3P Pro WC Display"),
        "3P Pro LTCT 5-10A": os.path.join(base_folder, "3P Pro LTCT Display NET Mode")
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

def add_formatted_runs(paragraph, text):
    if not text:
        return

    paragraph.text = ""
    import re

    default_font_size = Pt(10)
    pattern = re.compile(r"\[(font|size)=([^\]]+)\](.*?)\[/\1\]", re.IGNORECASE)
    last_end = 0

    for match in pattern.finditer(text):
        if match.start() > last_end:
            run = paragraph.add_run(text[last_end:match.start()])
            run.font.size = default_font_size

        run = paragraph.add_run(match.group(3))
        if match.group(1).lower() == "font":
            run.font.name = match.group(2)
            run.font.size = default_font_size
        else:
            try:
                run.font.size = Pt(int(match.group(2)))
            except ValueError:
                run.font.size = default_font_size
        last_end = match.end()

    if last_end < len(text):
        run = paragraph.add_run(text[last_end:])
        run.font.size = default_font_size


# === Table creation in docx ===
def create_table(doc, values, meter_type, energy_decimal, display_table_mode, batt_var):
    table = doc.add_table(rows=1, cols=3)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.allow_autofit = False
    table.style = 'Table Grid'

    hdr_cells = table.rows[0].cells
    hdr_texts = ['S.No', 'Parameters Display', 'Parameters Description']
    column_widths = [Cm(1.3), Cm(10.8), Cm(6.6)]
    for i, w in enumerate(column_widths):
        for cell in table.columns[i].cells:
            cell.width = w
        para = hdr_cells[i].paragraphs[0]
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = para.add_run(hdr_texts[i])
        run.bold = True
        run.font.size = Pt(14)

    if display_table_mode == "1":
        if batt_var == "1":
            folder = os.path.join(image_folders_net[meter_type], "Battery Mode")
        else:
            folder = image_folders_net[meter_type]
    elif display_table_mode == "0":
        if batt_var == "1":
            folder = os.path.join(image_folders[meter_type], "Battery Mode")
        else:
            folder = image_folders[meter_type]
    idx = 1
    for dec in values:
        filename = f"{dec:04d}"

        # Skip creating a row for NET individual mode when the file name is 0137
        if display_table_mode == "1" and filename == "0137":
            continue

        # Get parameter name from XML mapping
        param_name = predefined_names_new.get(dec, "").lower()
        forward_display_text = forward_display_text_mapping.get(dec, "")
        net_display_text = net_display_text_mapping.get(dec, "")

        # If parameter contains "energy" use energy_decimal logic
        if energy_decimal in ["1", "2"] and "energy" in param_name and "hr" not in param_name and "tamper" not in param_name:
            matches = glob.glob(os.path.join(folder, f"{filename}@{energy_decimal}_*.[jp][pn]g"))
        else:
            matches = glob.glob(os.path.join(folder, f"{filename}_*.[jp][pn]g"))

        if matches:
            for img_path in matches:
                row = table.add_row().cells
                row[0].text = str(idx)
                row[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                for i, w in enumerate(column_widths):
                    row[i].width = w
                run = row[1].paragraphs[0].add_run()
                run.add_picture(img_path, width=Inches(4.0))
                desc_para = row[2].paragraphs[0]
                desc_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                if display_table_mode == "1":
                    add_formatted_runs(desc_para, net_display_text)
                elif display_table_mode == "0":
                    add_formatted_runs(desc_para, forward_display_text)
            idx += 1
        else:
            row = table.add_row().cells
            row[0].text = str(idx)
            row[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            for i, w in enumerate(column_widths):
                row[i].width = w
            row[1].text = f"Image {filename} not found"
            desc_para = row[2].paragraphs[0]
            desc_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            if display_table_mode == "1":
                add_formatted_runs(desc_para, net_display_text)
            elif display_table_mode == "0":
                add_formatted_runs(desc_para, forward_display_text)
            idx += 1


def forward_create_table(doc, values, meter_type, energy_decimal, display_table_mode, batt_var):
    table = doc.add_table(rows=1, cols=3)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.allow_autofit = False
    table.style = 'Table Grid'

    hdr_cells = table.rows[0].cells
    hdr_texts = ['S.No', 'Parameters Display', 'Parameters Description']
    column_widths = [Cm(1.3), Cm(10.8), Cm(6.6)]
    for i, w in enumerate(column_widths):
        for cell in table.columns[i].cells:
            cell.width = w
        para = hdr_cells[i].paragraphs[0]
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = para.add_run(hdr_texts[i])
        run.bold = True
        run.font.size = Pt(14)

    if batt_var == "1":
        folder = os.path.join(image_folders[meter_type], "Battery Mode")
    else:
        folder = image_folders[meter_type]
    idx = 1
    for dec in values:
        filename = f"{dec:04d}"

        # Get parameter name from XML mapping
        param_name = predefined_names_new.get(dec, "").lower()
        forward_display_text = forward_display_text_mapping.get(dec, "")
        net_display_text = net_display_text_mapping.get(dec, "")

        # If parameter contains "energy" use energy_decimal logic
        if energy_decimal in ["1", "2"] and "energy" in param_name and "hr" not in param_name and "tamper" not in param_name:
            matches = glob.glob(os.path.join(folder, f"{filename}@{energy_decimal}_*.[jp][pn]g"))
        else:
            matches = glob.glob(os.path.join(folder, f"{filename}_*.[jp][pn]g"))

        if display_table_mode == "1":
            if "ex" not in param_name and "net active" not in param_name and "net apparent" not in param_name:
                if matches:
                    for img_path in matches:
                        row = table.add_row().cells
                        row[0].text = str(idx)
                        row[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                        for i, w in enumerate(column_widths):
                            row[i].width = w
                        run = row[1].paragraphs[0].add_run()
                        run.add_picture(img_path, width=Inches(4.0))
                        desc_para = row[2].paragraphs[0]
                        desc_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        add_formatted_runs(desc_para, forward_display_text)
                    idx += 1
                else:
                    row = table.add_row().cells
                    row[0].text = str(idx)
                    row[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                    for i, w in enumerate(column_widths):
                        row[i].width = w
                    row[1].text = f"Image {filename} not found"
                    desc_para = row[2].paragraphs[0]
                    desc_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    add_formatted_runs(desc_para, forward_display_text)
                    idx += 1

class InputDialog(simpledialog.Dialog):
    def body(self, master):
        tk.Label(master, text="Energy Decimal :").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.energy_var = tk.StringVar(value="1 Decimal Digit")
        self.energy_combo = ttk.Combobox(master, textvariable=self.energy_var, values=["1 Decimal Digit", "0 Decimal Digit", "2 Decimal Digit"], state="readonly", width=20)
        self.energy_combo.grid(row=0, column=1, padx=5, pady=5)

        tk.Label(master, text="Display Lock Feature:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.display_lock_var = tk.StringVar(value="0-No Mode lock")
        self.display_lock_combo = ttk.Combobox(master, textvariable=self.display_lock_var, values=["0-No Mode lock", "1-Auto Mode lock", "2-All Mode lock"], state="readonly", width=20)
        self.display_lock_combo.grid(row=1, column=1, padx=5, pady=5)

        tk.Label(master, text="Display Required for Lock:").grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.display_lock_show_var = tk.StringVar(value="0-Only C-Open Show")
        self.display_lock_show_combo = ttk.Combobox(master, textvariable=self.display_lock_show_var, values=["0-Only C-Open Show", "1-C-Open Shows With D&T"], state="disabled", width=20)
        self.display_lock_show_combo.grid(row=2, column=1, padx=5, pady=5)

        def update_lock_show_state(*args):
            selected_lock = self.display_lock_var.get()
            if selected_lock in {"1-Auto Mode lock", "2-All Mode lock"}:
                self.display_lock_show_combo.configure(state="readonly")
            else:
                self.display_lock_show_combo.configure(state="disabled")
                self.display_lock_show_var.set("0-Only C-Open Show")

        self.display_lock_var.trace_add("write", update_lock_show_state)
        update_lock_show_state()

        tk.Label(master, text="Bell Icon Display :").grid(row=3, column=0, padx=5, pady=5, sticky="w")
        self.bell_icon_var = tk.StringVar(value="0-PSEQ Only")
        self.bell_icon_combo = ttk.Combobox(master, textvariable=self.bell_icon_var, values=["0-PSEQ Only", "1-All Tampers"], state="readonly", width=20)
        self.bell_icon_combo.grid(row=3, column=1, padx=5, pady=5)

        tk.Label(master, text="Tamper Detect Time for Display :").grid(row=4, column=0, padx=5, pady=5, sticky="w")
        self.tamper_time_value = tk.StringVar(value="10 seconds")
        self.tamper_time_var = ttk.Entry(master, textvariable=self.tamper_time_value, width=22)
        self.tamper_time_var.grid(row=4, column=1, padx=5, pady=5)

        tk.Label(master, text="Display Table Required :").grid(row=5, column=0, padx=5, pady=5, sticky="w")
        self.display_table_mode_var = tk.StringVar(value="Forward & NET Combined Mode")
        self.display_table_mode_combo = ttk.Combobox(master, textvariable=self.display_table_mode_var, values=["Forward & NET Combined Mode", "Forward & NET Individual Mode"], state="readonly", width=20)
        self.display_table_mode_combo.grid(row=5, column=1, padx=5, pady=5)

        return self.energy_combo

    def apply(self):
        self.result = (
            self.energy_var.get().strip(),
            self.display_lock_var.get().strip(),
            self.display_lock_show_var.get().strip(),
            self.bell_icon_var.get().strip(),
            self.tamper_time_value.get().strip(),
            self.display_table_mode_var.get().strip()
        )

def get_user_inputs(root):
    dialog = InputDialog(root, title="Meter Configuration")

    energy_decimal, display_lock, display_lock_show, bell_icon, tamper_time, display_table_mode = dialog.result

    if energy_decimal == "0 Decimal Digit":
        energy_decimal = "0"
    elif energy_decimal == "1 Decimal Digit":
        energy_decimal = "1"
    elif energy_decimal == "2 Decimal Digit":
        energy_decimal = "2"

    if display_lock == "0-No Mode lock":
        display_lock = "0"
    elif display_lock == "1-Auto Mode lock":
        display_lock = "1"
    elif display_lock == "2-All Mode lock":
        display_lock = "2"

    if display_lock_show == "0-Only C-Open Show":
        display_lock_show = "0"
    elif display_lock_show == "1-C-Open Shows With D&T":
        display_lock_show = "1"

    if bell_icon == "0-PSEQ Only":
        bell_icon = "0"
    elif bell_icon == "1-All Tampers":
        bell_icon = "1"

    if display_table_mode == "Forward & NET Combined Mode":
        display_table_mode = "0"
    elif display_table_mode == "Forward & NET Individual Mode":
        display_table_mode = "1"


    return energy_decimal, display_lock, display_lock_show, bell_icon, tamper_time, display_table_mode

def generate_hex_string():
    global auto_string, push_string, tamper_string, batt_string, diag_string, test_string
        
    auto_vals = [int(item.split(" - ")[0]) for item in auto_selected.get(0, tk.END)]
    push_vals = [int(item.split(" - ")[0]) for item in push_selected.get(0, tk.END)]
    tamper_vals = [int(item.split(" - ")[0]) for item in tamper_selected.get(0, tk.END)]
    batt_vals = [int(item.split(" - ")[0]) for item in battery_selected.get(0, tk.END)]
    diag_vals = [int(item.split(" - ")[0]) for item in diag_selected.get(0, tk.END)]
    test_vals = [int(item.split(" - ")[0]) for item in test_selected.get(0, tk.END)]

    auto_input_string = auto_input_hex_string.get().strip()
    if not validate_hex_input_string(auto_input_string, "Auto"):
        return

    push_input_string = push_input_hex_string.get().strip()
    if not validate_hex_input_string(push_input_string, "Push"):
        return

    tamper_input_string = tamper_input_hex_string.get().strip()
    if not validate_hex_input_string(tamper_input_string, "Tamper"):
        return

    battery_input_string = battery_input_hex_string.get().strip()
    if not validate_hex_input_string(battery_input_string, "Battery"):
        return

    diag_input_string = diag_input_hex_string.get().strip()
    if not validate_hex_input_string(diag_input_string, "Diagnostic"):
        return

    test_input_string = test_input_hex_string.get().strip()
    if not validate_hex_input_string(test_input_string, "Test"):
        return

    tamper_payload = build_hex_payload(tamper_vals)
    auto_string.set(build_hex_payload(auto_vals, replacement_payload=tamper_payload))
    push_string.set(build_hex_payload(push_vals, replacement_payload=tamper_payload))
    tamper_string.set(tamper_payload)
    batt_string.set(build_hex_payload(batt_vals))
    diag_string.set(build_hex_payload(diag_vals))
    test_string.set(build_hex_payload(test_vals))

    save_display_strings({
        "auto_string": auto_string.get(),
        "push_string": push_string.get(),
        "tamper_string": tamper_string.get(),
        "batt_string": batt_string.get(),
        "diag_string": diag_string.get(),
        "test_string": test_string.get(),
    })

# === Generate document === 
def generate_doc():
    global auto_string, push_string, tamper_string, batt_string, diag_string, test_string
    table_seq_no = 1
    refer_table_seq_no = 1
    batt_var = "0"  # Initialize batt_var to "0" (not in battery mode)
    if not base_folder:
        messagebox.showerror("Error", "Please select a Base Folder first")
        return

    meter_type = meter_type_var.get()
    if meter_type not in image_folders:
        messagebox.showerror("Error", "Please select a valid Meter Type")
        return 
    
    auto_vals = [int(item.split(" - ")[0]) for item in auto_selected.get(0, tk.END)]
    push_vals = [int(item.split(" - ")[0]) for item in push_selected.get(0, tk.END)]
    tamper_vals = [int(item.split(" - ")[0]) for item in tamper_selected.get(0, tk.END)]
    batt_vals = [int(item.split(" - ")[0]) for item in battery_selected.get(0, tk.END)]
    diag_vals = [int(item.split(" - ")[0]) for item in diag_selected.get(0, tk.END)]
    test_vals = [int(item.split(" - ")[0]) for item in test_selected.get(0, tk.END)]

    auto_input_string = auto_input_hex_string.get().strip()
    if not validate_hex_input_string(auto_input_string, "Auto"):
        return

    push_input_string = push_input_hex_string.get().strip()
    if not validate_hex_input_string(push_input_string, "Push"):
        return

    tamper_input_string = tamper_input_hex_string.get().strip()
    if not validate_hex_input_string(tamper_input_string, "Tamper"):
        return

    battery_input_string = battery_input_hex_string.get().strip()
    if not validate_hex_input_string(battery_input_string, "Battery"):
        return

    diag_input_string = diag_input_hex_string.get().strip()
    if not validate_hex_input_string(diag_input_string, "Diagnostic"):
        return

    test_input_string = test_input_hex_string.get().strip()
    if not validate_hex_input_string(test_input_string, "Test"):
        return

    if not auto_vals:
        messagebox.showerror("Error", "Please select at least one parameter or insert hex string in Auto Scroll")
        return
    if not push_vals:
        messagebox.showerror("Error", "Please select at least one parameter or insert hex string in Push Scroll")
        return
    if not diag_vals:
        diag_vals = [92, 3101, 3103, 3106, 3107, 3102, 3105, 3104, 3108]

    energy_decimal, display_lock, display_lock_show, bell_icon, tamper_time, display_table_mode = get_user_inputs(root)
    if energy_decimal is None or display_lock is None or display_lock_show is None or bell_icon is None or tamper_time is None or display_table_mode is None:
        return  # User cancelled or invalid input

    tamper_payload = build_hex_payload(tamper_vals)
    auto_string.set(build_hex_payload(auto_vals, replacement_payload=tamper_payload))
    push_string.set(build_hex_payload(push_vals, replacement_payload=tamper_payload))
    tamper_string.set(tamper_payload)
    batt_string.set(build_hex_payload(batt_vals))
    diag_string.set(build_hex_payload(diag_vals))
    test_string.set(build_hex_payload(test_vals))

    save_display_strings({
        "auto_string": auto_string.get(),
        "push_string": push_string.get(),
        "tamper_string": tamper_string.get(),
        "batt_string": batt_string.get(),
        "diag_string": diag_string.get(),
        "test_string": test_string.get(),
    })
    
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
    run = para.add_run("Display Parameter ")
    run = para.add_run(project_code_var.get())
    run.font.size = Pt(13)

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

    if display_table_mode == "0":
        if auto_vals:
            add_table_section(doc, f"Table {table_seq_no}: Auto Scroll Display :", auto_vals, meter_type, energy_decimal, display_table_mode, batt_var)
            table_seq_no += 1

        if push_vals:
            add_table_section(doc, f"Table {table_seq_no}: Push Button Display :", push_vals, meter_type, energy_decimal, display_table_mode, batt_var)
            table_seq_no += 1

        if batt_vals:
            batt_var = "1"
            add_table_section(doc, f"Table {table_seq_no}: Auto and Push Mode Display (Battery)", batt_vals, meter_type, energy_decimal, display_table_mode, batt_var)
            batt_var = "0"  # Reset batt_var to "0" after battery mode table
            table_seq_no += 1

        if diag_vals:
            add_table_section(doc, f"Table {table_seq_no}: Diagnostic Mode Display :", diag_vals, meter_type, energy_decimal, display_table_mode, batt_var)
            table_seq_no += 1

        if test_vals:
            add_table_section(doc, f"Table {table_seq_no}: Test Mode Display :", test_vals, meter_type, energy_decimal, display_table_mode, batt_var)
            table_seq_no += 1

        if tamper_vals:
            add_table_section(doc, f"Table {table_seq_no}: Present Tamper Status Display :", tamper_vals, meter_type, energy_decimal, display_table_mode, batt_var, add_break=False)
            table_seq_no += 1

    elif display_table_mode == "1":
        if auto_vals:
            forward_add_table_section(doc, f"Table {table_seq_no}: Auto Scroll Display : Forwarded Mode", auto_vals, meter_type, energy_decimal, display_table_mode, batt_var)
            table_seq_no += 1

        if push_vals:
            forward_add_table_section(doc, f"Table {table_seq_no}: Push Button Display : Forwarded Mode", push_vals, meter_type, energy_decimal, display_table_mode, batt_var)
            table_seq_no += 1

        if batt_vals:
            batt_var = "1"
            forward_add_table_section(doc, f"Table {table_seq_no}: Auto and Push Mode Display (Battery) : Forwarded Mode", batt_vals, meter_type, energy_decimal, display_table_mode, batt_var)
            batt_var = "0"  # Reset batt_var to "0" after battery mode table
            table_seq_no += 1

        if diag_vals:
            forward_add_table_section(doc, f"Table {table_seq_no}: Diagnostic Mode Display : Forwarded Mode", diag_vals, meter_type, energy_decimal, display_table_mode, batt_var)
            table_seq_no += 1

        if auto_vals:
            add_table_section(doc, f"Table {table_seq_no}: Auto Scroll Display : NET Metering Mode", auto_vals, meter_type, energy_decimal, display_table_mode, batt_var)
            table_seq_no += 1

        if push_vals:
            add_table_section(doc, f"Table {table_seq_no}: Push Button Display : NET Metering Mode", push_vals, meter_type, energy_decimal, display_table_mode, batt_var)
            table_seq_no += 1

        if batt_vals:
            batt_var = "1"
            add_table_section(doc, f"Table {table_seq_no}: Auto and Push Mode Display (Battery) : NET Metering Mode", batt_vals, meter_type, energy_decimal, display_table_mode, batt_var)
            batt_var = "0"  # Reset batt_var to "0" after battery mode table
            table_seq_no += 1

        if diag_vals:
            add_table_section(doc, f"Table {table_seq_no}: Diagnostic Mode Display : NET Metering Mode", diag_vals, meter_type, energy_decimal, display_table_mode, batt_var)
            table_seq_no += 1

        if test_vals:
            add_table_section(doc, f"Table {table_seq_no}: Test Mode Display :", test_vals, meter_type, energy_decimal, display_table_mode, batt_var)
            table_seq_no += 1

        if tamper_vals:
            add_table_section(doc, f"Table {table_seq_no}: Present Tamper Status Display :", tamper_vals, meter_type, energy_decimal, display_table_mode, batt_var, add_break=False)
            table_seq_no += 1
            

    doc.add_heading("Note: ", level=2)
    if 9999 in auto_vals and 9999 in push_vals:
        para = doc.add_paragraph(f"1.  In the Auto Mode Display and Push Mode Display tables, present tamper status display parameter was normally hidden and if tamper present continuously for {tamper_time} then status showing as per the above table (Present Tamper Status Display) for respective tamper.")    
    elif 9999 in auto_vals:
        para = doc.add_paragraph(f"1.  In the Auto Mode Display table, present tamper status display parameter was normally hidden and if tamper present continuously for {tamper_time} then status showing as per the above table (Present Tamper Status Display) for respective tamper.")
    elif 9999 in push_vals:
        para = doc.add_paragraph(f"1.  In the Push Mode Display table, present tamper status display parameter was normally hidden and if tamper present continuously for {tamper_time} then status showing as per the above table (Present Tamper Status Display) for respective tamper.")
    else:
        para = doc.add_paragraph(f"1.  Present tamper status display parameter was normally hidden and if tamper present continuously for {tamper_time} then status showing as above table for respective tamper.")
    
    para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    para = doc.add_paragraph("2.")  
    para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    run = para.add_run()
    run = para.add_run(" If net-meter mode configured in meter, all forwarded parameters converted into Import parameters and distinguish with ")
    run.add_picture(os.path.join(base_folder, "Template images", "import.jpg"), width=Inches(0.25))
    run = para.add_run(" on display.")
    para = doc.add_paragraph("3.  In the Auto Mode Display and Push Mode Display default here show all 8 zones for zone wise display parameters otherwise meter will show zones wise display parameters as per the configured zones in meter.")
    para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    if display_lock in["1", "2"]:
        para = doc.add_paragraph("4.")   
        para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        run = para.add_run()
        run = para.add_run(" If top cover open tamper occurred in the meter then only the ") 
        run.add_picture(os.path.join(base_folder, "Template images", "C-Open.jpg"), width=Inches(0.70))
        if display_lock == "1":
            if display_lock_show == "0":
                run = para.add_run("display continuously at LCD when meter power ON at supply or battery. During this time normal auto scroll mode will get stopped.However push parameters will be available in push mode as usual.")
            elif display_lock_show == "1":        
                run = para.add_run(" with its Occurrence Date & Time display continuously auto scrolling with respective scrolling time at LCD when meter power ON at supply or battery. During this time normal auto scroll mode will get stopped.However push parameters will be available in push mode as usual.")
        elif display_lock == "2":
            if display_lock_show == "0":
                run = para.add_run("display continuously at LCD when meter power ON at supply or battery, and all other display modes will be disabled.")
            elif display_lock_show == "1":        
                run = para.add_run(" with its Occurrence Date & Time display continuously auto scrolling with respective scrolling time at LCD when meter power ON at supply or battery, and all other display modes will be disabled.")
        if display_lock_show == "0":
            para = doc.add_picture(os.path.join(base_folder, "Template images", "C-Open-Status.png"), width=Inches(5.0))
            # Center align the picture
            last_paragraph = doc.paragraphs[-1]  # the paragraph where the picture was added
            last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif display_lock_show == "1":
            table = doc.add_table(rows=1, cols=2)
            table.alignment = WD_TABLE_ALIGNMENT.CENTER
            table.autofit = True
            table.style = 'Table Grid'

            set_column_widths(table, [Cm(1.3), Cm(12.8)])

            first_row = table.rows[0].cells
            first_row[0].text = "1"
            first_row[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            first_row[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            first_row[1].paragraphs[0].add_run().add_picture(os.path.join(base_folder, "Template images", "C-Open-Status.png"),width=Inches(5.0))

            second_row = table.add_row().cells
            second_row[0].text = "2"
            second_row[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            second_row[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            second_row[1].paragraphs[0].add_run().add_picture(os.path.join(base_folder, "Template images", "C-Open-Status-Date.png"),width=Inches(5.0))
            

            third_row = table.add_row().cells
            third_row[0].text = "3"
            third_row[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            third_row[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            third_row[1].paragraphs[0].add_run().add_picture(os.path.join(base_folder, "Template images", "C-Open-Status-Time.png"),width=Inches(5.0))    

    doc.add_heading("\nDisplay Modes : ", level=2)
    para = doc.add_paragraph("Display parameters are displayed in four different modes called Auto scrolling mode, push button mode, diagnostic mode and test mode.")
    para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY


    if display_table_mode == "0":
        doc.add_heading(f"\nAUTO MODE-- (Auto Scrolling, Refer TABLE-{refer_table_seq_no}): ", level=3)
    elif display_table_mode == "1":
        if batt_vals:
            new_refer_table_seq_no = refer_table_seq_no + 4
        else:
            new_refer_table_seq_no = refer_table_seq_no + 3
        doc.add_heading(f"\nAUTO MODE-- (Auto Scrolling, Refer TABLE-{refer_table_seq_no} & TABLE-{new_refer_table_seq_no}): ", level=3)
    para = doc.add_paragraph(auto_txt.get())
    para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    refer_table_seq_no += 1

    if display_table_mode == "0":
        doc.add_heading(f"\nPUSH MODE-- (Scroll with Push Button, Refer TABLE-{refer_table_seq_no}): ", level=3)
    elif display_table_mode == "1":
        if batt_vals:
            new_refer_table_seq_no = refer_table_seq_no + 4
        else:
            new_refer_table_seq_no = refer_table_seq_no + 3
        doc.add_heading(f"\nPUSH MODE-- (Scroll with Push Button, Refer TABLE-{refer_table_seq_no} & TABLE-{new_refer_table_seq_no}): ", level=3)
    para = doc.add_paragraph(push_txt.get())
    para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    refer_table_seq_no += 1

    if batt_vals:
        refer_table_seq_no_temp = refer_table_seq_no
        refer_table_seq_no += 1

    if display_table_mode == "0":
        doc.add_heading(f"\nDIAGNOSTIC MODE-- (High resolution display, Refer TABLE-{refer_table_seq_no}): ", level=3)
    elif display_table_mode == "1":
        if batt_vals:
            new_refer_table_seq_no = refer_table_seq_no + 4
        else:
            new_refer_table_seq_no = refer_table_seq_no + 3
        doc.add_heading(f"\nDIAGNOSTIC MODE-- (High resolution display, Refer TABLE-{refer_table_seq_no} & TABLE-{new_refer_table_seq_no}): ", level=3)
    para = doc.add_paragraph(diag_txt.get())
    para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    refer_table_seq_no += 1

    if test_vals:
        if display_table_mode == "0":
            doc.add_heading(f"\nTEST MODE—(Test Mode Display, Refer TABLE-{refer_table_seq_no}): ", level=3)
        elif display_table_mode == "1":
            if batt_vals:
                new_refer_table_seq_no = refer_table_seq_no + 4
            else:
                new_refer_table_seq_no = refer_table_seq_no + 3
            doc.add_heading(f"\nTEST MODE—(Test Mode Display, Refer TABLE-{new_refer_table_seq_no}): ", level=3)
        refer_table_seq_no += 1
    else:
        doc.add_heading(f"\nTEST MODE—(Test Mode Display): ", level=3)
    para = doc.add_paragraph(test_txt.get())
    para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    if batt_vals: 
        if display_table_mode == "0":
            doc.add_heading(f"\nUnder Power OFF condition—(Battery Mode Display: Refer TABLE-{refer_table_seq_no_temp}): ", level=3)
        elif display_table_mode == "1":
            new_refer_table_seq_no = refer_table_seq_no_temp + 4
            doc.add_heading(f"\nUnder Power OFF condition—(Battery Mode Display: Refer TABLE-{refer_table_seq_no_temp} & TABLE-{new_refer_table_seq_no}): ", level=3)
    else: 
        doc.add_heading(f"\nUnder Power OFF condition: ", level=3)
    para = doc.add_paragraph(batt_txt.get())
    para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    if 3132 in test_vals:
        add_page_break(doc)
        add_table_section(doc, f"Table {table_seq_no}: Relay Disconnect Reason (Present Status):", [3301, 3302, 3303, 3304, 3305], meter_type, energy_decimal, display_table_mode, batt_var)
        table_seq_no += 1
    table_seq_no = insert_error_tables(doc, test_vals, table_seq_no)

    doc.add_heading("\nLCD Glass Description: ", level=2)
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
    if bell_icon == "0":
        para7.add_run(" Sign indicates phase sequence reverse in meter.")
    elif bell_icon == "1":
        para7.add_run(" Sign indicates any tamper exist in the meter.")

    if(meter_type == "3P-Pro WC 10-60A" or meter_type == "3P-Pro WC (Torrent Specific) 10-60A" or meter_type == "3P-Pro WC 20-100A"):
        para8 = doc.add_paragraph("8.  ")
        run = para8.add_run()
        run.add_picture(os.path.join(base_folder, "Template images", "relay.jpg"), width=Inches(0.3))
        para8.add_run(" Sign indicates relay connect/disconnect status.")

    doc.add_heading("\n Revision History: ", level=2)
    table = doc.add_table(rows=1, cols=3)
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    table.autofit = True
    table.style = 'Table Grid'

    set_column_widths(table, [Inches(1.0), Inches(1.5), Inches(4.0)])
    hdr_cells = table.rows[0].cells
    headers = ["Rev. No.", "Date", "Revision Description"]
    for cell, text in zip(hdr_cells, headers):
        run = cell.paragraphs[0].add_run(text)
        run.bold = True
        run.font.size = Pt(14)
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    rows = [(revision_var.get(), date_var.get(), "Initial")]

    for rev_number, date, description in rows:
        row = table.add_row().cells
        row[0].text = rev_number
        row[1].text = date
        row[2].text = description
        row[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        row[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        row[2].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER


    # === Save DOCX ===
    if not output_path.lower().endswith(".docx"):
        output_path = output_path + ".docx"
        output_var.set(output_path)

    doc.save(output_path)

    # pdf_path = get_pdf_path(output_path)
    # pdf_message = ""
    # try:
    #     converter = convert_docx_to_pdf(output_path, pdf_path)
    #     pdf_message = f"\nPDF created using {converter}:\n{pdf_path}"
    # except Exception as exc:
    #     pdf_message = f"\nPDF conversion failed:\n{exc}"

    messagebox.showinfo(
        "Success",
        # f"Word file created:\n{output_path}{pdf_message}"
        f"Word file created:\n{output_path}"
    )


# === Tkinter GUI ===
ctk.set_appearance_mode("light")   # Options: "light", "dark", "system"
ctk.set_default_color_theme("blue")  # Try "green", "dark-blue", etc.
# === Main Window ===
root = ctk.CTk()
# root.iconbitmap("logo.ico")
root.title("3P Pro Meter Display Report Generator V4.0")
root.geometry("1200x780")
root.resizable(False, False)

# === Top Frame ===
top_frame = ctk.CTkFrame(root, corner_radius=15)
top_frame.pack(pady=15, padx=15, fill="x")

# Base Folder (left side)
ctk.CTkLabel(top_frame, text="Base Folder:", text_color="#1a1a1a").grid(row=0, column=0, padx=5, sticky="w")
ctk.CTkButton(top_frame, text="Browse Base Folder", fg_color="#a843ec", hover_color="#155a8a", command= choose_base_folder).grid(row=0, column=1, padx=5, sticky="w")

# Meter Type (middle side)
ctk.CTkLabel(top_frame, text="Meter Type:").grid(row=0, column=2, padx=20, sticky="w")
meter_type_var = tk.StringVar(value="Select Type")
meter_type_combo = ctk.CTkComboBox(top_frame, variable=meter_type_var, values=list(image_folders.keys()))
meter_type_combo.grid(row=0, column=3, padx=5, sticky="w")

# Date input with calendar picker (right side)
ctk.CTkLabel(top_frame, text="Date:").grid(row=0, column=4, padx=30, sticky="w")
date_var = tk.StringVar()
date_entry = DateEntry(top_frame, textvariable=date_var, date_pattern="dd-mm-yyyy", width=15)
date_entry.grid(row=0, column=5, padx=5, sticky="w")

# Revision Number (right side)
ctk.CTkLabel(top_frame, text="Rev. No:").grid(row=0, column=6, padx=40, sticky="w")
revision_var = tk.StringVar(value="00")
revision_combo = ctk.CTkComboBox(top_frame, variable=revision_var, values=["00", "01", "02", "03"], width=65)
revision_combo.grid(row=0, column=7, padx=5, sticky="w")

ctk.CTkLabel(top_frame, text="Project Code").grid(row=0, column=8, padx=30, sticky="w")
project_code_var = tk.StringVar(value="G3P_DC21")
project_code_entry = ctk.CTkEntry(top_frame, textvariable=project_code_var, width=120)
project_code_entry.grid(row=0, column=9, padx=5, sticky="w")


# === Dual listbox UI ===

def auto_fill_selected(input_var, available_listbox, selected_listbox):
    """Auto-fill the selected listbox from a HEX entry string.

    The input is treated as a continuous HEX stream, split into 4-character chunks.
    Each valid chunk is converted to decimal and formatted as a 4-digit string.
    Matching available items are inserted into selected_listbox, duplicates avoided.
    """
    hex_text = input_var.get().strip().upper()
    selected_listbox.delete(0, tk.END)

    if not hex_text:
        return

    values = []
    for i in range(0, len(hex_text), 4):
        chunk = hex_text[i:i+4]
        if len(chunk) != 4:
            continue
        try:
            values.append(int(chunk, 16))
        except ValueError:
            continue

    if not values:
        return

    available_items = getattr(available_listbox, "full_items", available_listbox.get(0, tk.END))
    available_map = {item.split(" - ", 1)[0]: item for item in available_items}

    for value in values:
        formatted = f"{value:04d}"
        match = available_map.get(formatted)
        if match and match not in selected_listbox.get(0, tk.END):
            selected_listbox.insert(tk.END, match)


def create_dual_listbox(frame, title, items):
    ttk.Label(frame, text=title, font=("Arial", 16, "bold")).grid(row=0, column=0, columnspan=3, pady=5)

    # --- Search Entry ---
    search_var = tk.StringVar()
    search_entry = ttk.Entry(frame, textvariable=search_var, width=60)
    search_entry.grid(row=1, column=0, padx=10, pady=5)

    ttk.Label(frame, text="Input String:").grid(row=1, column=1, padx=5, sticky="e")
    
    input_hex_string = tk.StringVar()
    input_string = ttk.Entry(frame, textvariable=input_hex_string, width=80)
    input_string.grid(row=1, column=2, padx=10, pady=5)
    input_string.insert(0, "")
    

    # --- Available Listbox + Scrollbar ---
    available_frame = tk.Frame(frame)
    available_frame.grid(row=2, column=0, padx=10, pady=5)
    available = tk.Listbox(available_frame, selectmode=tk.EXTENDED, width=70, height=12)
    available_scroll = tk.Scrollbar(available_frame, orient=tk.VERTICAL, command=available.yview)
    available.config(yscrollcommand=available_scroll.set)
    available.grid(row=0, column=0, sticky="nsew")
    available_scroll.grid(row=0, column=1, sticky="ns")

    # --- Selected Listbox + Scrollbar ---
    selected_frame = tk.Frame(frame)
    selected_frame.grid(row=2, column=2, padx=10, pady=5)
    selected = tk.Listbox(selected_frame, selectmode=tk.EXTENDED, width=70, height=12)
    selected_scroll = tk.Scrollbar(selected_frame, orient=tk.VERTICAL, command=selected.yview)
    selected.config(yscrollcommand=selected_scroll.set)
    selected.grid(row=0, column=0, sticky="nsew")
    selected_scroll.grid(row=0, column=1, sticky="ns")

    # --- Populate Available ---
    filtered_items = items.copy()
    available.full_items = filtered_items
    for item in filtered_items:
        available.insert(tk.END, item)

    # --- Button Functions ---
    def add_item():
        selections = available.curselection()

        if not selections:
            return

        existing = set(selected.get(0, tk.END))

        for idx in selections:
            text = available.get(idx)
            if text not in existing:
                selected.insert(tk.END, text)
                existing.add(text)

        # Keep focus on the last selected item
        available.selection_clear(0, tk.END)

        next_idx = selections[-1] + 1
        if next_idx < available.size():
            available.selection_set(next_idx)
            available.activate(next_idx)
            available.see(next_idx)

    def remove_item():
        selections = selected.curselection()

        if not selections:
            return

        first = selections[0]

        for idx in reversed(selections):
            selected.delete(idx)

        if selected.size() > 0:
            new_idx = min(first, selected.size() - 1)
            selected.selection_set(new_idx)
            selected.activate(new_idx)
            selected.see(new_idx)

    def move_up():
        selections = list(selected.curselection())

        if not selections or selections[0] == 0:
            return

        # Move each selected item up by one
        for idx in selections:
            item = selected.get(idx)
            selected.delete(idx)
            selected.insert(idx - 1, item)

        # Restore selection
        selected.selection_clear(0, tk.END)
        for idx in selections:
            selected.selection_set(idx - 1)

        selected.activate(selections[0] - 1)
        selected.see(selections[0] - 1)

    def move_down():
        selections = list(selected.curselection())

        if not selections or selections[-1] == selected.size() - 1:
            return

        # Move from bottom to top
        for idx in reversed(selections):
            item = selected.get(idx)
            selected.delete(idx)
            selected.insert(idx + 1, item)

        # Restore selection
        selected.selection_clear(0, tk.END)
        for idx in selections:
            selected.selection_set(idx + 1)

        selected.activate(selections[-1] + 1)
        selected.see(selections[-1] + 1)

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
    input_hex_string.trace_add("write", lambda *args: auto_fill_selected(input_hex_string, available, selected))

    return available, selected, input_hex_string



# Load default items from predefined_names
default_items = [f"{dec:04d} - {desc}" for dec, desc in predefined_names_new.items()]
diag_items = [f"{dec:04d} - {desc}" for dec, desc in diag_predefined_names_new.items()]
test_items = [f"{dec:04d} - {desc}" for dec, desc in test_predefined_names_new.items()]
tamper_items = [f"{dec:04d} - {desc}" for dec, desc in tamper_predefined_names_new.items()]

saved_display_strings = load_display_strings()


def restore_saved_selection(input_var, listbox, payload, available_items, payload_name=None):
    if payload_name in ("auto", "push"):
        tamper_payload = saved_display_strings.get("tamper_string", "")
        if tamper_payload:
            tamper_body = extract_payload_body(tamper_payload)
            if tamper_body and tamper_body in payload:
                payload = payload.replace(tamper_body, "270F")

    body = extract_payload_body(payload)
    input_var.set(body)

    values = parse_payload_to_values(payload, available_items)
    listbox.delete(0, tk.END)
    for item in values:
        if item not in listbox.get(0, tk.END):
            listbox.insert(tk.END, item)

# Create a frame with fixed size for the notebook
notebook_frame = tk.Frame(root, width=1080, height=335)
notebook_frame.pack(padx=10, pady=(0,5))  # adjust pady to reduce gap

# Prevent frame from shrinking to fit notebook
notebook_frame.pack_propagate(False)

# Tabs for Auto / Push
notebook = ttk.Notebook(notebook_frame)
notebook.pack(pady=20, expand=True, fill="both")

auto_tab = ttk.Frame(notebook)
notebook.add(auto_tab, text="Auto Scroll Selection")
auto_available, auto_selected, auto_input_hex_string = create_dual_listbox(auto_tab, "Auto Scroll Selection", default_items)

push_tab = ttk.Frame(notebook)
notebook.add(push_tab, text="Push Scroll Selection")
push_available, push_selected, push_input_hex_string  = create_dual_listbox(push_tab, "Push Scroll Selection", default_items)

tamper_tab = ttk.Frame(notebook)
notebook.add(tamper_tab, text="Tamper Selection")
tamper_available, tamper_selected, tamper_input_hex_string  = create_dual_listbox(tamper_tab, "Present Tamper Status Selection", tamper_items)

battery_tab = ttk.Frame(notebook)
notebook.add(battery_tab, text="Battery Mode Selection")
battery_available, battery_selected, battery_input_hex_string  = create_dual_listbox(battery_tab, "Battery Mode Selection", default_items)

diag_tab = ttk.Frame(notebook)
notebook.add(diag_tab, text="Diagnostic Selection")
diag_available, diag_selected, diag_input_hex_string  = create_dual_listbox(diag_tab, "Diagnostic Selection", diag_items)

test_tab = ttk.Frame(notebook)
notebook.add(test_tab, text="Test Mode Selection")
test_available, test_selected, test_input_hex_string  = create_dual_listbox(test_tab, "Test Mode Selection", test_items)

restore_saved_selection(auto_input_hex_string, auto_selected, saved_display_strings.get("auto_string", ""), default_items, "auto")
restore_saved_selection(push_input_hex_string, push_selected, saved_display_strings.get("push_string", ""), default_items, "push")
restore_saved_selection(tamper_input_hex_string, tamper_selected, saved_display_strings.get("tamper_string", ""), tamper_items, "tamper")
restore_saved_selection(battery_input_hex_string, battery_selected, saved_display_strings.get("batt_string", ""), default_items, "battery")
restore_saved_selection(diag_input_hex_string, diag_selected, saved_display_strings.get("diag_string", ""), diag_items, "diag")
restore_saved_selection(test_input_hex_string, test_selected, saved_display_strings.get("test_string", ""), test_items, "test")

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
ctk.CTkButton(output_frame, text="Browse", command=choose_output, fg_color="#2dbe9f").grid(row=0, column=2, padx=30, sticky="w")

# Generate Report button
ctk.CTkButton(output_frame, text="Generate Report", command=generate_doc, fg_color="#37b63d").grid(row=2, column=2, padx=30, sticky="w")

# Generate Hex String button
ctk.CTkButton(output_frame, text="Generate Hex String", command=generate_hex_string, fg_color="#4946f1").grid(row=3, column=2, padx=30, sticky="w")
child_processes = [] 

def open_profile_object_selection():
    """Open Profile_Object_Selection from the same EXE."""
    if getattr(sys, "frozen", False):
        process = subprocess.Popen([sys.executable, "--profile-object-selection"], cwd=os.path.dirname(sys.executable)) 
        child_processes.append(process)
    else:
        run_embedded_script("Profile_Object_Selection.py")

ctk.CTkButton(output_frame, text="WC Profile Objects String", command=open_profile_object_selection, fg_color="#b341a3").grid(row=4, column=2, padx=30, sticky="w")

def open_top_level_string_selection():
    """Open Top_Level_String_Selection from the same EXE."""
    if getattr(sys, "frozen", False):
        process = subprocess.Popen([sys.executable, "--top-level-string-selection"], cwd=os.path.dirname(sys.executable))
        child_processes.append(process)
    else:
        run_embedded_script("Top_Level_String_Selection.py")

ctk.CTkButton(output_frame, text="Top Level String", command=open_top_level_string_selection, fg_color="#A34528").grid(row=5, column=2, padx=30, sticky="w")

def close_main_application():
    """Close all child EXE windows first, then close the main window."""
    for process in child_processes[:]:
        try:
            if process.poll() is None:
                process.terminate()

                try:
                    process.wait(timeout=2)
                except subprocess.TimeoutExpired:
                    process.kill()

        except (OSError, ValueError):
            pass

        finally:
            try:
                child_processes.remove(process)
            except ValueError:
                pass

    root.destroy()

root.protocol("WM_DELETE_WINDOW", close_main_application) 

saved_display_strings = load_display_strings()

auto_string = tk.StringVar(value="")

ctk.CTkLabel(output_frame, text="").grid(row=1, column=3, padx=5, sticky="w")

ttk.Label(output_frame, text="Auto Display String:").grid(row=2, column=0, padx=5, sticky="w")
ttk.Entry(output_frame, textvariable=auto_string, width=130).grid(row=2, column=1, padx=30, sticky="e")

push_string = tk.StringVar(value="")

ttk.Label(output_frame, text="Push Display String:").grid(row=3, column=0, padx=5, sticky="w")
ttk.Entry(output_frame, textvariable=push_string, width=130).grid(row=3, column=1, padx=30, sticky="e")

tamper_string = tk.StringVar(value="")

batt_string = tk.StringVar(value="")

ttk.Label(output_frame, text="Battery Display String:").grid(row=6, column=0, padx=5, sticky="w")
ttk.Entry(output_frame, textvariable=batt_string, width=130).grid(row=6, column=1, padx=30, sticky="e")

diag_string = tk.StringVar(value="")

ttk.Label(output_frame, text="Diag Display String:").grid(row=4, column=0, padx=5, sticky="w")
ttk.Entry(output_frame, textvariable=diag_string, width=130).grid(row=4, column=1, padx=30, sticky="e")

test_string = tk.StringVar(value="")

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
