import os
import re
import sys
import tkinter as tk
import xml.etree.ElementTree as ET
from tkinter import filedialog, messagebox
from tkinter import ttk

TAB_NAMES = [
    "Load-Survey-Profile",
    "Daily-Energy-Profile",
    "Billing-Profile",
    "Event-Profile-1",
    "Event-Profile-2",
    "Nameplate-Profile",
    "Instant-Profile",
]

Pro_3P_WC_Load_Survey_Profile = []
Pro_3P_WC_Daily_Energy_Profile = []
Pro_3P_WC_Billing_Profile = []
Pro_3P_WC_Event_Profile_1 = []
Pro_3P_WC_Event_Profile_2 = []
Pro_3P_WC_Nameplate_Profile = []
Pro_3P_WC_Instantaneous_Profile = []

OBJECT_LISTS = {
        "Load-Survey-Profile": [],
        "Daily-Energy-Profile": [],
        "Billing-Profile": [],
        "Event-Profile-1": [],
        "Event-Profile-2": [],
        "Nameplate-Profile": [],
        "Instant-Profile": [],
}

# Determine app folder and default base folder
if getattr(sys, 'frozen', False):
    base_folder = os.path.dirname(sys.executable)
else:
    base_folder = os.path.dirname(os.path.abspath(__file__))

def parse_profile_xml(path):
    tree = ET.parse(path)
    return tree.getroot()


def read_items_from_section(root, section_name):
    aliases = {
        "Load-Survey-Profile": ["Profile_3PProWCLoadSurvey", "3P-Pro-WC-Load-Survey-Profile"],
        "Daily-Energy-Profile": ["Profile_3PProWCDailyEnergy", "3P-Pro-WC-Daily-Energy-Profile"],
        "Billing-Profile": ["Profile_3PProWCBillingEnergy", "3P-Pro-WC-Billing-Profile"],
        "Event-Profile-1": ["Profile_3PProWCEvent1", "3P-Pro-WC-Event-Profile-1"],
        "Event-Profile-2": ["Profile_3PProWCEvent2", "3P-Pro-WC-Event-Profile-2"],
        "Nameplate-Profile": ["Profile_3PProWCNameplate", "3P-Pro-WC-Nameplate-Profile"],
        "Instant-Profile": ["Profile_3PProWCInstantaneous", "3P-Pro-WC-Instantaneous-Profile"],
    }

    for tag_name in aliases.get(section_name, [section_name]):
        section_node = root.find(tag_name)
        if section_node is None:
            continue

        values = []
        for item in section_node.findall("Item"):
            text = (item.text or "").strip()
            if text:
                values.append(text)
        return values

    return []


def load_xml_mappings():
    global Pro_3P_WC_Load_Survey_Profile, Pro_3P_WC_Daily_Energy_Profile, Pro_3P_WC_Billing_Profile, Pro_3P_WC_Event_Profile_1, Pro_3P_WC_Event_Profile_2, Pro_3P_WC_Nameplate_Profile, Pro_3P_WC_Instantaneous_Profile, OBJECT_LISTS

    object_xml_path = os.path.join(base_folder, "RefXML",  "profile_objects_3p_pro.xml")

    if os.path.exists(object_xml_path):
        root = parse_profile_xml(object_xml_path)
        Pro_3P_WC_Load_Survey_Profile = read_items_from_section(root, "Profile_3PProWCLoadSurvey")
        Pro_3P_WC_Daily_Energy_Profile = read_items_from_section(root, "Profile_3PProWCDailyEnergy")
        Pro_3P_WC_Billing_Profile = read_items_from_section(root, "Profile_3PProWCBillingEnergy")
        Pro_3P_WC_Event_Profile_1 = read_items_from_section(root, "Profile_3PProWCEvent1")
        Pro_3P_WC_Event_Profile_2 = read_items_from_section(root, "Profile_3PProWCEvent2")
        Pro_3P_WC_Nameplate_Profile = read_items_from_section(root, "Profile_3PProWCNameplate")
        Pro_3P_WC_Instantaneous_Profile = read_items_from_section(root, "Profile_3PProWCInstantaneous")
    else:
        Pro_3P_WC_Load_Survey_Profile = []
        Pro_3P_WC_Daily_Energy_Profile = []
        Pro_3P_WC_Billing_Profile = []
        Pro_3P_WC_Event_Profile_1 = []
        Pro_3P_WC_Event_Profile_2 = []
        Pro_3P_WC_Nameplate_Profile = []
        Pro_3P_WC_Instantaneous_Profile = []

    OBJECT_LISTS = {
        "Load-Survey-Profile": Pro_3P_WC_Load_Survey_Profile,
        "Daily-Energy-Profile": Pro_3P_WC_Daily_Energy_Profile,
        "Billing-Profile": Pro_3P_WC_Billing_Profile,
        "Event-Profile-1": Pro_3P_WC_Event_Profile_1,
        "Event-Profile-2": Pro_3P_WC_Event_Profile_2,
        "Nameplate-Profile": Pro_3P_WC_Nameplate_Profile,
        "Instant-Profile": Pro_3P_WC_Instantaneous_Profile,
    }

class CheckboxPanel(ttk.Frame):
    def __init__(self, parent, items, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)

        # Scrollable canvas to hold a large number of checkbuttons
        canvas = tk.Canvas(self)
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        self.inner = ttk.Frame(canvas)

        self.inner_id = canvas.create_window((0, 0), window=self.inner, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        def _on_mousewheel(event):
            if event.delta < 0:
                canvas.yview_scroll(1, "units")
            elif event.delta > 0:
                canvas.yview_scroll(-1, "units")
            return "break"

        def _on_shift_mousewheel(event):
            if event.delta < 0:
                canvas.xview_scroll(1, "units")
            elif event.delta > 0:
                canvas.xview_scroll(-1, "units")
            return "break"

        canvas.bind("<MouseWheel>", _on_mousewheel)
        canvas.bind("<Shift-MouseWheel>", _on_shift_mousewheel)
        self.bind("<MouseWheel>", _on_mousewheel)
        self.bind("<Shift-MouseWheel>", _on_shift_mousewheel)

        for child in self.winfo_children():
            child.bind("<MouseWheel>", _on_mousewheel)
            child.bind("<Shift-MouseWheel>", _on_shift_mousewheel)

        self.vars = []
        for idx, label in enumerate(items):
            var = tk.IntVar(value=0)
            cb = ttk.Checkbutton(self.inner, text=label, variable=var)
            cb.grid(row=idx, column=0, sticky="w", padx=4, pady=2)
            self.vars.append((label, var))

        # update scrollregion when needed
        self.inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

        # make the widget receive focus for wheel events
        canvas.bind("<Enter>", lambda e: canvas.focus_set())
        canvas.bind("<Button-1>", lambda e: canvas.focus_set())

    def get_bits(self):
        # Return list of 0/1 for each checkbox in order
        return [int(var.get()) for _, var in self.vars]

    def get_labels(self):
        return [label for label, _ in self.vars]


class ProfileTabbedApp(tk.Tk):
    def __init__(self):
        super().__init__()
        load_xml_mappings()
        self.title("3P Pro Profile Objects Selection V1.0")
        self.geometry("800x620")
        self.resizable(False, False)

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=8, pady=8)

        self.panels = {}
        self.output_texts = {}
        self.note_texts = {}
        self.all_profile_objects_text = None
        for name in TAB_NAMES:
            frame = ttk.Frame(self.notebook)
            self.notebook.add(frame, text=name)
            # build inner layout: checkbox panel + actions panel
            left = ttk.Frame(frame)
            left.pack(side="left", fill="both", expand=True, padx=(8, 4), pady=8)

            items = OBJECT_LISTS.get(name, [])
            panel = CheckboxPanel(left, items)
            panel.pack(fill="both", expand=True)
            self.panels[name] = panel

            right = ttk.Frame(frame, width=320)
            right.pack(side="right", fill="y", padx=(4, 8), pady=8)

            ttk.Label(right, text="Actions", font=(None, 10, "bold")).pack(pady=(4, 8))
            gen_btn = ttk.Button(right, text="Generate String", command=self.generate_for_all_tabs)
            gen_btn.pack(fill="x", padx=6, pady=4)

            ttk.Label(right, text="Output:").pack(anchor="w", padx=6, pady=(12, 2))
            output_text = tk.Text(right, height=10, width=40)
            output_text.pack(fill="both", padx=6, pady=(0, 6), expand=False)
            self.output_texts[name] = output_text

            if name == "Billing-Profile":
                ttk.Label(right, text="Note:", font=(None, 10, "bold")).pack(anchor="w", padx=6, pady=(14, 2))
                ttk.Label(right, text="Maximum Phase wise THD with Date & Time for Voltage").pack(anchor="w", padx=6, pady=(0, 2))
                ttk.Label(right, text="and Current Objects will be shows value only if it is also").pack(anchor="w", padx=6, pady=(0, 2))
                ttk.Label(right, text="enabled from Firmware also.").pack(anchor="w", padx=6, pady=(0, 2))
                

        bottom_frame = ttk.Frame(self)
        bottom_frame.pack(fill="x", padx=8, pady=(0, 8))

        ttk.Label(bottom_frame, text="All Profile Objects String:", font=(None, 10, "bold")).pack(anchor="w", pady=(0, 4))
        self.all_profile_objects_text = tk.Text(bottom_frame, height=2, width=80)
        self.all_profile_objects_text.pack(fill="both", expand=False)

        convert_btn = ttk.Button(bottom_frame, text="Convert String to Bits", command=self.convert_string_to_bits)
        convert_btn.pack(anchor="e", pady=(6, 0))

        # status bar
        self.status = ttk.Label(self, text="Ready", anchor="w")
        self.status.pack(fill="x", padx=4, pady=(0, 4))


    def _build_output_for_tab(self, tab_name, panel):
        bits = panel.get_bits()
        # bits[0] -> MSB or LSB? We'll treat first item as MSB (leftmost) for display.
        bitstring = "".join(str(b) for b in bits)
        # also produce hex for convenience
        # pad to multiple of 8
        pad = (-len(bitstring)) % 8
        bitstring_padded = ("0" * pad) + bitstring
        hex_chunks = []
        for i in range(0, len(bitstring_padded), 8):
            chunk = bitstring_padded[i:i+8]
            hex_chunks.append(f"{int(chunk, 2):02X}")
        hex_string_temp = "".join(hex_chunks)
        length = len(bitstring_padded)  # number of bits, which is also the length of the bitstring
        if length <= 127:
            length_prefix = f"{length:02X}"
        elif length <= 255:
            length_prefix = "81" + f"{length:02X}"
        else:
            length_prefix = "82" + f"{length:04X}"

        hex_string = f"04{length_prefix}{hex_string_temp}"  # prepend length as 1 byte
        display_text = f"Tab: {tab_name}\nBits: {bitstring}\nHex: {hex_string}\n"
        return display_text, hex_string

    def generate_for_all_tabs(self):
        tab_hex_strings = {}
        for tab_name, panel in self.panels.items():
            out, hex_string = self._build_output_for_tab(tab_name, panel)
            tab_hex_strings[tab_name] = hex_string

        combined_profile_string = "02081101" + "".join(tab_hex_strings.values())

        for tab_name in self.panels:
            output_text = self.output_texts.get(tab_name)

            if output_text is not None:
                display_text, _ = self._build_output_for_tab(tab_name, self.panels[tab_name])
                output_text.delete("1.0", tk.END)
                output_text.insert(tk.END, display_text)

        if self.all_profile_objects_text is not None:
            self.all_profile_objects_text.delete("1.0", tk.END)
            self.all_profile_objects_text.insert(tk.END, combined_profile_string)

        self.status.config(text="Generated strings for all tabs")

    def convert_string_to_bits(self):
        if self.all_profile_objects_text is None:
            return

        all_profile_objects_string = self.all_profile_objects_text.get("1.0", tk.END).strip()
        if not all_profile_objects_string:
            messagebox.showerror("Error", "Please generate or enter an All Profile Objects string first.")
            return

        length = len(all_profile_objects_string)
        if length != 146:
            messagebox.showerror("Error", "Input All Profile ObjectsHex string is invalid. Check length.")
            return

        if all_profile_objects_string.startswith("02081101"):
            all_profile_objects_string = all_profile_objects_string[8:]

        if all_profile_objects_string.startswith("0420"):
            all_profile_objects_string = all_profile_objects_string[4:]
            load_survey_profile_object_string = all_profile_objects_string[:8]
            all_profile_objects_string = all_profile_objects_string[8:]

        if all_profile_objects_string.startswith("0438"):
            all_profile_objects_string = all_profile_objects_string[4:]
            daily_energy_profile_object_string = all_profile_objects_string[:14]
            all_profile_objects_string = all_profile_objects_string[14:]

        if all_profile_objects_string.startswith("0481B8"):
            all_profile_objects_string = all_profile_objects_string[6:]
            billing_profile_object_string = all_profile_objects_string[:46]
            all_profile_objects_string = all_profile_objects_string[46:]

        if all_profile_objects_string.startswith("0418"):
            all_profile_objects_string = all_profile_objects_string[4:]
            event_profile_1_object_string = all_profile_objects_string[:6]
            all_profile_objects_string = all_profile_objects_string[6:]

        if all_profile_objects_string.startswith("0408"):
            all_profile_objects_string = all_profile_objects_string[4:]
            event_profile_2_object_string = all_profile_objects_string[:2]
            all_profile_objects_string = all_profile_objects_string[2:]

        if all_profile_objects_string.startswith("0410"):
            all_profile_objects_string = all_profile_objects_string[4:]
            nameplate_profile_object_string = all_profile_objects_string[:4]
            all_profile_objects_string = all_profile_objects_string[4:]

        if all_profile_objects_string.startswith("0470"):
            all_profile_objects_string = all_profile_objects_string[4:]
            instant_profile_object_string = all_profile_objects_string[:28]
            all_profile_objects_string = all_profile_objects_string[28:]

        if len(all_profile_objects_string) > 0:
            messagebox.showerror("Error", "Input All Profile ObjectsHex string is invalid. Extra data found.")
            return

        load_survey_profile_object_bits = hex_string_to_bits(load_survey_profile_object_string)
        daily_energy_profile_object_bits = hex_string_to_bits(daily_energy_profile_object_string)
        billing_profile_object_bits = hex_string_to_bits(billing_profile_object_string)
        event_profile_1_object_bits = hex_string_to_bits(event_profile_1_object_string)
        event_profile_2_object_bits = hex_string_to_bits(event_profile_2_object_string)
        nameplate_profile_object_bits = hex_string_to_bits(nameplate_profile_object_string)
        instant_profile_object_bits = hex_string_to_bits(instant_profile_object_string)

        self.apply_bits_to_panel("Load-Survey-Profile", load_survey_profile_object_bits)
        self.apply_bits_to_panel("Daily-Energy-Profile", daily_energy_profile_object_bits)
        self.apply_bits_to_panel("Billing-Profile", billing_profile_object_bits)
        self.apply_bits_to_panel("Event-Profile-1", event_profile_1_object_bits)
        self.apply_bits_to_panel("Event-Profile-2", event_profile_2_object_bits)
        self.apply_bits_to_panel("Nameplate-Profile", nameplate_profile_object_bits)
        self.apply_bits_to_panel("Instant-Profile", instant_profile_object_bits)

        self.status.config(text="Loaded bits for all profile tabs")

    def apply_bits_to_panel(self, tab_name, bits_string):
        panel = self.panels.get(tab_name)
        if panel is None:
            return False

        cleaned_bits = "".join(ch for ch in str(bits_string).strip() if ch in "01")
        if cleaned_bits.startswith("0b"):
            cleaned_bits = cleaned_bits[2:]

        expected_count = len(panel.vars)
        if not cleaned_bits:
            for _, var in panel.vars:
                var.set(0)
            return True

        if len(cleaned_bits) < expected_count:
            cleaned_bits = cleaned_bits.zfill(expected_count)
        elif len(cleaned_bits) > expected_count:
            cleaned_bits = cleaned_bits[-expected_count:]

        for index, (_, var) in enumerate(panel.vars):
            var.set(1 if cleaned_bits[index] == "1" else 0)

        return True


def hex_string_to_bits(hex_string):
    # Remove any whitespace and convert to uppercase
    hex_string = hex_string.replace(" ", "").upper()

    # Convert hex string to binary string without the '0b' prefix
    if not hex_string:
        return ""

    binary_string = bin(int(hex_string, 16))[2:]

    return binary_string



if __name__ == "__main__":
    app = ProfileTabbedApp()
    app.mainloop()
