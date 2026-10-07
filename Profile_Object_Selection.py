import os
import sys
import tkinter as tk
import xml.etree.ElementTree as ET
from tkinter import messagebox, ttk

TAB_NAMES = [
    "Load-Survey-Profile",
    "Daily-Energy-Profile",
    "Billing-Profile",
    "Event-Profile-1",
    "Event-Profile-2",
    "Nameplate-Profile",
    "Instant-Profile",
]

PROFILE_XML_ALIASES = {
    "Load-Survey-Profile": {
        "3P-WC": ["Profile_3PProWCLoadSurvey", "3P-Pro-WC-Load-Survey-Profile"],
        "3P-LTCT": ["Profile_3PProLTCTLoadSurvey", "3P-Pro-LTCT-Load-Survey-Profile"],
    },
    "Daily-Energy-Profile": {
        "3P-WC": ["Profile_3PProWCDailyEnergy", "3P-Pro-WC-Daily-Energy-Profile"],
        "3P-LTCT": ["Profile_3PProLTCTDailyEnergy", "3P-Pro-LTCT-Daily-Energy-Profile"],
    },
    "Billing-Profile": {
        "3P-WC": ["Profile_3PProWCBillingEnergy", "3P-Pro-WC-Billing-Profile"],
        "3P-LTCT": ["Profile_3PProLTCTBillingEnergy", "3P-Pro-LTCT-Billing-Profile"],
    },
    "Event-Profile-1": {
        "3P-WC": ["Profile_3PProWCEvent1", "3P-Pro-WC-Event-Profile-1"],
        "3P-LTCT": ["Profile_3PProLTCTEvent1", "3P-Pro-LTCT-Event-Profile-1"],
    },
    "Event-Profile-2": {
        "3P-WC": ["Profile_3PProWCEvent2", "3P-Pro-WC-Event-Profile-2"],
        "3P-LTCT": ["Profile_3PProLTCTEvent2", "3P-Pro-LTCT-Event-Profile-2"],
    },
    "Nameplate-Profile": {
        "3P-WC": ["Profile_3PProWCNameplate", "3P-Pro-WC-Nameplate-Profile"],
        "3P-LTCT": ["Profile_3PProLTCTNameplate", "3P-Pro-LTCT-Nameplate-Profile"],
    },
    "Instant-Profile": {
        "3P-WC": ["Profile_3PProWCInstantaneous", "3P-Pro-WC-Instantaneous-Profile"],
        "3P-LTCT": ["Profile_3PProLTCTInstantaneous", "3P-Pro-LTCT-Instantaneous-Profile"],
    },
}

PROFILE_FIELDS = {
    "Load-Survey-Profile": ("Load_Survey_Profile", "LoadSurvey"),
    "Daily-Energy-Profile": ("Daily_Energy_Profile", "DailyEnergy"),
    "Billing-Profile": ("Billing_Profile", "BillingEnergy"),
    "Event-Profile-1": ("Event_Profile_1", "Event1"),
    "Event-Profile-2": ("Event_Profile_2", "Event2"),
    "Nameplate-Profile": ("Nameplate_Profile", "Nameplate"),
    "Instant-Profile": ("Instantaneous_Profile", "Instantaneous"),
}

BASE_FOLDER = os.path.dirname(os.path.abspath(__file__)) if not getattr(sys, "frozen", False) else os.path.dirname(sys.executable)


def parse_profile_xml(path):
    return ET.parse(path).getroot()


def choose_profile_type():
    prompt_root = tk.Tk()
    prompt_root.title("Profile Objects Selection")
    prompt_root.resizable(False, False)

    selected_profile = {"value": None}
    content = ttk.Frame(prompt_root, padding=20)
    content.pack(fill="both", expand=True)
    ttk.Label(
        content,
        text="Please select the meter type for profile-object-selection string generation",
    ).pack(pady=(0, 16))

    buttons = ttk.Frame(content)
    buttons.pack()

    def select_profile(profile_type):
        selected_profile["value"] = profile_type
        prompt_root.destroy()

    ttk.Button(buttons, text="3P-WC", command=lambda: select_profile("3P-WC")).pack(side="left", padx=6)
    ttk.Button(buttons, text="3P-LTCT", command=lambda: select_profile("3P-LTCT")).pack(side="left", padx=6)

    prompt_root.protocol("WM_DELETE_WINDOW", prompt_root.destroy)
    prompt_root.update_idletasks()
    width = prompt_root.winfo_reqwidth()
    height = prompt_root.winfo_reqheight()
    x = (prompt_root.winfo_screenwidth() - width) // 2
    y = (prompt_root.winfo_screenheight() - height) // 2
    prompt_root.geometry(f"{width}x{height}+{x}+{y}")
    prompt_root.grab_set()
    prompt_root.mainloop()
    return selected_profile["value"]


def read_items_from_section(root, section_name, profile_type=None):
    candidates = []
    if profile_type:
        candidates.extend(PROFILE_XML_ALIASES.get(section_name, {}).get(profile_type, []))
    if not candidates:
        candidates.append(section_name)

    for tag_name in candidates:
        section_node = root.find(tag_name)
        if section_node is None:
            continue

        values = []
        for item in section_node.findall("Item"):
            text = (item.text or "").strip()
            if text:
                values.append(text)
        if values:
            return values

    return []


class ProfileXMLLoader:
    def __init__(self, base_folder):
        self.base_folder = base_folder

    def load_xml_mappings(self, profile_type):
        object_xml_path = os.path.join(self.base_folder, "RefXML", "profile_objects_3p_pro.xml")
        profile_code = profile_type.removeprefix("3P-")
        object_lists = {tab_name: [] for tab_name in TAB_NAMES}

        if not os.path.exists(object_xml_path):
            return object_lists

        root = parse_profile_xml(object_xml_path)
        for tab_name in TAB_NAMES:
            values = []
            aliases = PROFILE_XML_ALIASES.get(tab_name, {}).get(profile_type, [])
            for tag_name in aliases or [f"Profile_3PPro{profile_code}{PROFILE_FIELDS[tab_name][1]}"]:
                section_node = root.find(tag_name)
                if section_node is None:
                    continue
                values = [
                    (item.text or "").strip()
                    for item in section_node.findall("Item")
                    if (item.text or "").strip()
                ]
                if values:
                    break
            object_lists[tab_name] = values

        return object_lists


class CheckboxPanel(ttk.Frame):
    def __init__(self, parent, items, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)

        canvas = tk.Canvas(self)
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        self.inner = ttk.Frame(canvas)
        canvas.create_window((0, 0), window=self.inner, anchor="nw")
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
            checkbutton = ttk.Checkbutton(self.inner, text=label, variable=var)
            checkbutton.grid(row=idx, column=0, sticky="w", padx=4, pady=2)
            self.vars.append((label, var))

        self.inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Enter>", lambda e: canvas.focus_set())
        canvas.bind("<Button-1>", lambda e: canvas.focus_set())

    def get_bits(self):
        return [int(var.get()) for _, var in self.vars]


class ProfileTabbedApp(tk.Tk):
    def __init__(self, profile_type, object_lists):
        super().__init__()
        self.profile_type = profile_type
        self.object_lists = object_lists

        if not any(object_lists.values()):
            self.withdraw()
            messagebox.showerror(
                "Profile Objects Unavailable",
                f"No {profile_type} profile-object mappings were found in RefXML/profile_objects_3p_pro.xml.",
                parent=self,
            )
            self.destroy()
            return

        self.title(f"{profile_type} Profile Objects Selection V1.0")
        self.geometry("800x720")
        self.resizable(False, False)

        self.update_idletasks()
        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()
        win_w = 800
        win_h = 720
        pos_x = max((screen_w - win_w) // 2, 0)
        pos_y = max((screen_h - win_h) // 2, 0)
        self.geometry(f"{win_w}x{win_h}+{pos_x}+{pos_y}")
        self.state("normal")
        self.deiconify()
        self.lift()
        self.focus_force()
        self.after(100, lambda: self.attributes("-topmost", True))
        self.after(300, lambda: self.attributes("-topmost", False))

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=8, pady=8)

        self.panels = {}
        self.output_texts = {}
        self.all_profile_objects_text = None

        for name in TAB_NAMES:
            frame = ttk.Frame(self.notebook)
            self.notebook.add(frame, text=name)

            left = ttk.Frame(frame)
            left.pack(side="left", fill="both", expand=True, padx=(8, 4), pady=8)

            items = object_lists.get(name, [])
            panel = CheckboxPanel(left, items)
            panel.pack(fill="both", expand=True)
            self.panels[name] = panel

            right = ttk.Frame(frame, width=320)
            right.pack(side="right", fill="y", padx=(4, 8), pady=8)

            ttk.Label(right, text="Actions", font=(None, 10, "bold")).pack(pady=(4, 8))
            ttk.Button(right, text="Generate String", command=self.generate_for_all_tabs).pack(fill="x", padx=6, pady=4)

            ttk.Label(right, text="Individual Tab Output:", font=(None, 10, "bold")).pack(anchor="w", padx=6, pady=(12, 2))
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

        ttk.Button(bottom_frame, text="Convert String to Bits", command=self.convert_string_to_bits).pack(anchor="e", pady=(6, 0))
        self.status = ttk.Label(self, text="Ready", anchor="w")
        self.status.pack(fill="x", padx=4, pady=(0, 4))

    def _build_output_for_tab(self, tab_name, panel):
        bits = panel.get_bits()
        bitstring = "".join(str(bit) for bit in bits)
        pad = (-len(bitstring)) % 8
        bitstring_padded = ("0" * pad) + bitstring

        hex_chunks = [f"{int(bitstring_padded[i:i+8], 2):02X}" for i in range(0, len(bitstring_padded), 8)]
        hex_string_temp = "".join(hex_chunks)
        length = len(bitstring_padded)

        if length <= 127:
            length_prefix = f"{length:02X}"
        elif length <= 255:
            length_prefix = "81" + f"{length:02X}"
        else:
            length_prefix = "82" + f"{length:04X}"

        hex_string = f"04{length_prefix}{hex_string_temp}"
        display_text = f"Tab: {tab_name}\nBits: {bitstring}\nHex: {hex_string}\n"
        return display_text, hex_string

    def generate_for_all_tabs(self):
        tab_hex_strings = {}
        for tab_name, panel in self.panels.items():
            _, hex_string = self._build_output_for_tab(tab_name, panel)
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
        messagebox.showinfo("Success", "Strings generated for all tabs.")

    def convert_string_to_bits(self):
        if self.all_profile_objects_text is None:
            return

        all_profile_objects_string = self.all_profile_objects_text.get("1.0", tk.END).strip()
        if not all_profile_objects_string:
            messagebox.showerror("Error", "Please generate or enter an All Profile Objects string first.")
            return

        if len(all_profile_objects_string) != 146:
            messagebox.showerror("Error", "Input All Profile ObjectsHex string is invalid. Check length.")
            return

        if all_profile_objects_string.startswith("02081101"):
            all_profile_objects_string = all_profile_objects_string[8:]

        values = {}
        if all_profile_objects_string.startswith("0420"):
            all_profile_objects_string = all_profile_objects_string[4:]
            values["Load-Survey-Profile"] = all_profile_objects_string[:8]
            all_profile_objects_string = all_profile_objects_string[8:]

        if all_profile_objects_string.startswith("0438"):
            all_profile_objects_string = all_profile_objects_string[4:]
            values["Daily-Energy-Profile"] = all_profile_objects_string[:14]
            all_profile_objects_string = all_profile_objects_string[14:]

        if all_profile_objects_string.startswith("0481B8"):
            all_profile_objects_string = all_profile_objects_string[6:]
            values["Billing-Profile"] = all_profile_objects_string[:46]
            all_profile_objects_string = all_profile_objects_string[46:]

        if all_profile_objects_string.startswith("0418"):
            all_profile_objects_string = all_profile_objects_string[4:]
            values["Event-Profile-1"] = all_profile_objects_string[:6]
            all_profile_objects_string = all_profile_objects_string[6:]

        if all_profile_objects_string.startswith("0408"):
            all_profile_objects_string = all_profile_objects_string[4:]
            values["Event-Profile-2"] = all_profile_objects_string[:2]
            all_profile_objects_string = all_profile_objects_string[2:]

        if all_profile_objects_string.startswith("0410"):
            all_profile_objects_string = all_profile_objects_string[4:]
            values["Nameplate-Profile"] = all_profile_objects_string[:4]
            all_profile_objects_string = all_profile_objects_string[4:]

        if all_profile_objects_string.startswith("0470"):
            all_profile_objects_string = all_profile_objects_string[4:]
            values["Instant-Profile"] = all_profile_objects_string[:28]
            all_profile_objects_string = all_profile_objects_string[28:]

        if all_profile_objects_string:
            messagebox.showerror("Error", "Input All Profile ObjectsHex string is invalid. Extra data found.")
            return

        for tab_name, hex_string in values.items():
            bits_string = hex_string_to_bits(hex_string)
            self.apply_bits_to_panel(tab_name, bits_string)

        self.status.config(text="Loaded bits for all profile tabs")
        messagebox.showinfo("Success", "Strings converted to bits and applied to all tabs.")

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
    hex_string = (hex_string or "").replace(" ", "").upper()
    if not hex_string:
        return ""
    return bin(int(hex_string, 16))[2:]


def launch_profile_selection():
    profile_type = sys.argv[1] if len(sys.argv) > 1 else None
    if profile_type not in {"3P-WC", "3P-LTCT"}:
        profile_type = choose_profile_type()
    if profile_type is None:
        return

    loader = ProfileXMLLoader(BASE_FOLDER)
    object_lists = loader.load_xml_mappings(profile_type)
    app = ProfileTabbedApp(profile_type, object_lists)
    app.mainloop()


if __name__ == "__main__":
    launch_profile_selection()
