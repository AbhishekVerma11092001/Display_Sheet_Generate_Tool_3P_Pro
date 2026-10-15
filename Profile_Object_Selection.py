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
    """Modern startup selector for the supported meter/profile variants."""
    prompt_root = tk.Tk()
    prompt_root.title("Profile Object Selection")
    prompt_root.resizable(False, False)
    prompt_root.configure(bg="#F3F6F9")

    selected_profile = {"value": None}

    # Professional header
    header = tk.Frame(prompt_root, bg="#4B70E7", height=82)
    header.pack(fill="x")
    header.pack_propagate(False)

    tk.Label(header, text="PROFILE OBJECT SELECTION", bg="#4B70E7", fg="white", font=("Segoe UI Semibold", 17) ).pack(anchor="w", padx=24, pady=(14, 0))
    tk.Label(header, text="Select the smart-meter profile variant", bg="#4B70E7", fg="#BFD3E5", font=("Segoe UI", 10)).pack(anchor="w", padx=24)

    content = tk.Frame(prompt_root, bg="#FFFFFF", padx=24, pady=22)
    content.pack(fill="both", expand=True)

    tk.Label(content,
        text="Meter configuration", bg="#FFFFFF", fg="#4B70E7", font=("Segoe UI Semibold", 13)).pack(anchor="w")

    tk.Label(content, text="Choose the profile-object mapping to load from the XML configuration.", bg="#FFFFFF", fg="#4B70E7", font=("Segoe UI", 10)).pack(anchor="w", pady=(3, 18))

    card_frame = tk.Frame(content, bg="#FFFFFF")
    card_frame.pack(fill="x")

    def select_profile(profile_type):
        selected_profile["value"] = profile_type
        prompt_root.destroy()

    for profile_type, subtitle in (
        ("3P-WC", "3-phase • Whole Current"),
        ("3P-LTCT", "3-phase • LTCT"),
    ):
        card = tk.Frame(card_frame, bg="#F7F9FB", highlightbackground="#D8E0E8", highlightthickness=1)
        card.pack(side="left", fill="both", expand=True, padx=5)

        tk.Label(card, text=profile_type, bg="#F7F9FB", fg="#3D4EAF", font=("Segoe UI Semibold", 16) ).pack(anchor="w", padx=15, pady=(14, 2))
        tk.Label(card, text=subtitle, bg="#F7F9FB", fg="#4E67D6", font=("Segoe UI", 12)).pack(anchor="w", padx=15, pady=(0, 8))

        tk.Button(card, text="Select", command=lambda p=profile_type: select_profile(p), bg="#3D4EAF", fg="white", activebackground="#2A3B8C", activeforeground="white", relief="flat", bd=0, font=("Segoe UI Semibold", 12), padx=18, pady=7, cursor="hand2").pack(anchor="w", padx=15, pady=14)

    prompt_root.protocol("WM_DELETE_WINDOW", prompt_root.destroy)
    prompt_root.update_idletasks()
    width, height = 560, 360
    x = max((prompt_root.winfo_screenwidth() - width) // 2, 0)
    y = max((prompt_root.winfo_screenheight() - height) // 2, 0)
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
        self.required_items = {tab_name: set() for tab_name in TAB_NAMES}

    def load_xml_mappings(self, profile_type):
        object_xml_path = os.path.join(self.base_folder, "RefXML", "profile_objects_3p_pro.xml")
        # Development-friendly fallback when the XML is kept beside the EXE/script.
        if not os.path.exists(object_xml_path):
            fallback_xml = os.path.join(self.base_folder, "profile_objects_3p_pro.xml")
            if os.path.exists(fallback_xml):
                object_xml_path = fallback_xml

        profile_code = profile_type.removeprefix("3P-")
        object_lists = {tab_name: [] for tab_name in TAB_NAMES}
        self.required_items = {tab_name: set() for tab_name in TAB_NAMES}

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
                required_indices = set()
                for item in section_node.findall("Item"):
                    text = (item.text or "").strip()
                    if not text:
                        continue
                    if item.get("IS_15959_PARA", "").lower() == "true":
                        required_indices.add(len(values))
                    values.append(text)
                if values:
                    self.required_items[tab_name] = required_indices
                    break
            object_lists[tab_name] = values

        return object_lists


class CheckboxPanel(ttk.Frame):
    """Modern searchable multi-select checklist used by each profile tab."""

    def __init__(self, parent, items, required_indices=(), *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        self.items = list(items)
        self.required_indices = set(required_indices)
        self.vars = []
        self.filtered_items = list(enumerate(self.items))

        # Each list position gets its own variable. Duplicate parameter
        # names must remain independently selectable.
        self._item_vars = [
            tk.IntVar(value=int(index in self.required_indices))
            for index in range(len(self.items))
        ]

        # Search / selection toolbar
        toolbar = ttk.Frame(self, style="Toolbar.TFrame")
        toolbar.pack(fill="x", pady=(0, 8))

        ttk.Label(toolbar, text="Filter objects", style="ToolbarLabel.TLabel").pack(side="left", padx=(0, 8))

        self.search_var = tk.StringVar()
        search = ttk.Entry(toolbar, textvariable=self.search_var, width=32)
        search.pack(side="left", fill="x", expand=True, ipady=3)
        search.bind("<KeyRelease>", lambda e: self._filter())

        ttk.Button(toolbar, text="Clear", style="Small.TButton", command=lambda: self.search_var.set("") or self._filter()).pack(side="left", padx=(6, 0))
        ttk.Button(toolbar, text="Check All", style="Small.TButton", command=self.check_all).pack(side="left", padx=(6, 0))
        ttk.Button(toolbar, text="Uncheck All", style="Small.TButton", command=self.uncheck_all).pack(side="left", padx=(6, 0))

        self.count_label = ttk.Label(toolbar, text="", style="Count.TLabel")
        self.count_label.pack(side="right", padx=(10, 0))

        # Scrollable list
        list_frame = ttk.Frame(self, style="Panel.TFrame")
        list_frame.pack(fill="both", expand=True)

        self.canvas = tk.Canvas(list_frame, highlightthickness=0, bd=0, background=STYLE["surface"])
        scrollbar = ttk.Scrollbar( list_frame, orient="vertical", command=self.canvas.yview)
        self.inner = ttk.Frame(self.canvas, style="Panel.TFrame")
        self.window_id = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.canvas.configure(yscrollcommand=scrollbar.set)

        self.canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        self.canvas.bind("<Configure>", self._resize_inner)
        self.inner.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self._bind_wheel(self.inner)

        self._render()
        self._bind_wheel(self.canvas)
        self._update_count()

    def _resize_inner(self, event):
        self.canvas.itemconfigure(self.window_id, width=event.width)

    def _bind_wheel(self, widget):
        """Bind mouse-wheel scrolling to the widget and checklist children."""
        widget.bind("<MouseWheel>", self._on_mousewheel, add="+")
        widget.bind("<Button-4>", lambda e: self._scroll_units(-3), add="+")
        widget.bind("<Button-5>", lambda e: self._scroll_units(3), add="+")

    def _bind_wheel_to_item(self, widget):
        self._bind_wheel(widget)

    def _scroll_units(self, units):
        self.canvas.yview_scroll(units, "units")
        return "break"

    def _on_mousewheel(self, event):
        if event.delta:
            return self._scroll_units(-3 if event.delta > 0 else 3)
        return "break"

    def _filter(self):
        query = self.search_var.get().strip().lower()
        self.filtered_items = [
            (index, item)
            for index, item in enumerate(self.items)
            if query in item.lower()
        ]
        self._render()

    def _render(self):
        for child in self.inner.winfo_children():
            child.destroy()

        self.vars = []
        for row, (item_index, label) in enumerate(self.filtered_items):
            var = self._item_vars[item_index]

            row_frame = tk.Frame(self.inner, bg=STYLE["row"], highlightthickness=2)
            row_frame.grid(row=row, column=0, sticky="ew", padx=6, pady=2)
            row_frame.columnconfigure(0, weight=1)
            row_frame.configure(background=STYLE["row"])

            cb = tk.Checkbutton(
                row_frame,
                text=label,
                variable=var,
                selectcolor="white",
                bg=STYLE["row"],
                fg=STYLE["text"],
                font=("Segoe UI", 9),
                activebackground="white",
                activeforeground=STYLE["primary"],
                bd=0,
                anchor="w",
                state=tk.DISABLED if item_index in self.required_indices else tk.NORMAL,
                command=lambda index=item_index: self._item_changed(index),
            )
            cb.grid(row=0, column=0, sticky="ew", padx=8, pady=4)

            row_frame.bind("<Enter>", lambda event, frame=row_frame: self._set_hover_row(frame, True))
            row_frame.bind("<Leave>", lambda event, frame=row_frame: self._set_hover_row(frame, False))
            cb.bind("<Enter>", lambda event, frame=row_frame: self._set_hover_row(frame, True))
            cb.bind("<Leave>", lambda event, frame=row_frame: self._set_hover_row(frame, False))

            self._bind_wheel_to_item(row_frame)
            self._bind_wheel_to_item(cb)

            self.vars.append((item_index, label, var))

        self.inner.columnconfigure(0, weight=1)
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _set_hover_row(self, frame, is_hovered):
        frame.configure(background="#EAF9ED" if is_hovered else STYLE["row"])
        for child in frame.winfo_children():
            if hasattr(child, "configure"):
                child.configure(bg="#EAF9ED" if is_hovered else STYLE["row"])

    def check_all(self):
        for item_index, _item in self.filtered_items:
            self._item_vars[item_index].set(1)
        self._render()
        self._update_count()

    def uncheck_all(self):
        for item_index, _item in self.filtered_items:
            if item_index not in self.required_indices:
                self._item_vars[item_index].set(0)
        self._render()
        self._update_count()

    def set_bits(self, bits):
        for index, var in enumerate(self._item_vars):
            var.set(
                1 if index in self.required_indices else int(bool(bits[index]))
            )

    def get_bits(self):
        return [
            1 if index in self.required_indices else var.get()
            for index, var in enumerate(self._item_vars)
        ]

    def _item_changed(self, item_index):
        if item_index in self.required_indices:
            self._item_vars[item_index].set(1)
        self._update_count()

    def _update_count(self):
        selected = sum(var.get() for var in self._item_vars)
        visible = len(self.filtered_items)
        self.count_label.config(text=f"{selected} Selected  •  {visible} Shown")


class ProfileTabbedApp(tk.Tk):
    def __init__(self, profile_type, object_lists, required_items=None):
        super().__init__()
        self.profile_type = profile_type
        self.object_lists = object_lists
        required_items = required_items or {}

        if not any(object_lists.values()):
            self.withdraw()
            messagebox.showerror(
                "Profile Objects Unavailable",
                f"No {profile_type} profile-object mappings were found in RefXML/profile_objects_3p_pro.xml.",
                parent=self,
            )
            self.destroy()
            return

        self._configure_style()

        self.title(f"Profile Object Selection  |  {profile_type}")
        self.geometry("1180x780")
        self.minsize(1050, 600)
        self.resizable(False, False)

        self._center_window()
        self._build_header()

        # Main content
        content = ttk.Frame(self, style="App.TFrame")
        content.pack(fill="both", expand=True, padx=14, pady=(0, 10))

        self.notebook = ttk.Notebook(content, style="Modern.TNotebook")
        self.notebook.pack(fill="both", expand=True)

        self.panels = {}
        self.output_texts = {}
        self.all_profile_objects_text = None

        for name in TAB_NAMES:
            frame = ttk.Frame(self.notebook, style="Tab.TFrame", padding=12)
            self.notebook.add(frame, text=self._short_tab_name(name))

            items = object_lists.get(name, [])

            # Left: object selector
            selector = ttk.Frame(frame, style="Panel.TFrame", padding=10)
            selector.pack(side="left", fill="both", expand=True, padx=(0, 8))
            selector.columnconfigure(0, weight=1)
            selector.rowconfigure(1, weight=1)

            top = ttk.Frame(selector, style="Panel.TFrame")
            top.grid(row=0, column=0, sticky="ew", pady=(0, 8))
            ttk.Label(top, text=self._tab_heading(name), style="SectionTitle.TLabel").pack(side="left")
            ttk.Label(top, text=f"{len(items)} objects", style="Muted.TLabel").pack(side="right")

            panel = CheckboxPanel(
                selector, items, required_items.get(name, set())
            )
            panel.grid(row=1, column=0, sticky="nsew")
            self.panels[name] = panel

            # Right: generated output
            right = ttk.Frame(frame, style="Panel.TFrame", padding=12)
            right.pack(side="right", fill="y", padx=(8, 0))
            right.configure(width=355)
            right.pack_propagate(False)

            ttk.Label(right, text="OUTPUT", style="SectionTitle.TLabel").pack(anchor="w")
            ttk.Label(right, text="Individual profile-object string", style="SectionCaption.TLabel",).pack(anchor="w", pady=(2, 10))

            output_text = tk.Text(right, height=8, wrap="word", font=("Consolas", 11), bg=STYLE["input"], fg=STYLE["text"], insertbackground=STYLE["text"], relief="flat", bd=0, padx=10, pady=10, state="disabled",)
            output_text.pack(fill="x", pady=(0, 10))
            self.output_texts[name] = output_text

            note_all = (
                "Selection of IS 15959 parameter objects is mandatory; therefore, the checkboxes for these objects are disabled."
            )
            ttk.Label(right, text="NOTE", style="SectionTitle.TLabel").pack(anchor="w", pady=(14, 3))
            ttk.Label(right, text=note_all, wraplength=300, style="Note.TLabel", justify="left").pack(anchor="w")

            if name == "Billing-Profile":
                note = (
                    "Maximum Phase wise THD with Date & Time for Voltage and "
                    "Current Objects will show value only if enabled from Firmware also."
                )
                ttk.Label(right, text="NOTE", style="SectionTitle.TLabel").pack(anchor="w", pady=(14, 3))
                ttk.Label(right, text=note, wraplength=300, style="Note.TLabel", justify="left").pack(anchor="w")

        self._build_bottom_panel()
        self._build_statusbar()

    def _configure_style(self):
        global STYLE
        STYLE = {
            "bg": "#EAEEF3", "surface": "#FFFFFF", "panel": "#FFFFFF",
            "input": "#F3F6F9", "text": "#17212B", "muted": "#667381",
            "primary": "#1D55CE", "border": "#D8E0E8", "row": "#F8FAFC",
        }

        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        self.configure(bg=STYLE["bg"])

        style.configure(".", font=("Segoe UI", 10), foreground=STYLE["text"])
        style.configure("App.TFrame", background=STYLE["bg"])
        style.configure("Tab.TFrame", background=STYLE["bg"])
        style.configure("Panel.TFrame", background=STYLE["surface"])
        style.configure("Toolbar.TFrame", background=STYLE["surface"])
        style.configure("ToolbarLabel.TLabel", background=STYLE["surface"], foreground=STYLE["muted"], font=("Segoe UI Semibold", 10))
        style.configure("SectionCaption.TLabel", background=STYLE["surface"], foreground=STYLE["primary"], font=("Segoe UI Semibold", 10))
        style.configure("SectionTitle.TLabel", background=STYLE["surface"], foreground=STYLE["primary"], font=("Segoe UI Semibold", 13))
        style.configure("Muted.TLabel", background=STYLE["surface"], foreground=STYLE["muted"], font=("Segoe UI", 10))
        style.configure("Count.TLabel", background=STYLE["surface"], foreground=STYLE["primary"], font=("Segoe UI Semibold", 10))
        style.configure("Note.TLabel", background=STYLE["surface"], foreground=STYLE["muted"], font=("Segoe UI", 10))
        style.configure("Modern.TNotebook", background=STYLE["bg"], borderwidth=0, tabmargins=[2, 4, 2, 0])
        style.configure("Modern.TNotebook.Tab", padding=[14, 8], background="#E7EDF3", foreground="#40505F", font=("Segoe UI Semibold", 12), borderwidth=0)
        style.map("Modern.TNotebook.Tab", background=[("selected", STYLE["surface"])], foreground=[("selected", STYLE["primary"])])
        style.configure("CheckRow.TFrame", background=STYLE["row"], relief="flat")
        style.configure("Modern.TCheckbutton", background=STYLE["row"], foreground=STYLE["text"], font=("Segoe UI", 12))
        style.map("Modern.TCheckbutton", background=[("active", "#EEF5FB")], foreground=[("active", STYLE["primary"])])
        style.configure("Small.TButton", padding=[9, 5], background=STYLE["primary"], foreground="white", font=("Segoe UI Semibold", 9))
        style.map("Small.TButton", background=[("active", "#6A91E7"), ("pressed", "#1D55CE")])
        style.configure("Primary.TButton", padding=[12, 8], background=STYLE["primary"], foreground="white", font=("Segoe UI Semibold", 10), borderwidth=0,)
        style.map("Primary.TButton", background=[("active", "#6A91E7"), ("pressed", "#1D55CE")])
        style.configure("Action.TButton", padding=[10, 7], background="#5E8BEC", font=("Segoe UI Semibold", 10))
        style.map("Action.TButton",  background=[("active", "#9EB6E9"), ("pressed", "#5E8BEC")])

    def _center_window(self):
        self.update_idletasks()
        w, h = 1180, 780
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        x, y = max((sw - w) // 2, 0), max((sh - h) // 2, 0)
        self.geometry(f"{w}x{h}+{x}+{y}")
        self.deiconify()
        self.lift()
        self.focus_force()

    def _build_header(self):
        header = tk.Frame(self, bg="#1D55CE", height=74)
        header.pack(fill="x")
        header.pack_propagate(False)

        title_box = tk.Frame(header, bg="#1D55CE")
        title_box.pack(side="left", padx=20, pady=10)

        tk.Label(title_box, text="PROFILE OBJECT SELECTION", bg="#1D55CE", fg="white", font=("Segoe UI Semibold", 17)).pack(anchor="w")
        tk.Label(title_box, text=f"3P Smart Meter  •  {self.profile_type}  •  V2.0", bg="#1D55CE", fg="#BFD3E5", font=("Segoe UI", 10)).pack(anchor="w", pady=(2, 0))

        badge = tk.Label(header, text="READY", bg="#6389DB", fg="white", font=("Segoe UI Semibold", 10), padx=12, pady=5)
        badge.pack(side="right", padx=20)
        self.header_badge = badge

    def _build_bottom_panel(self):
        outer = ttk.Frame(self, style="Panel.TFrame", padding=12)
        outer.pack(fill="x", padx=14, pady=(0, 8))

        ttk.Label(outer, text="ALL PROFILE OBJECTS COMBINED STRING", style="SectionTitle.TLabel").pack(anchor="w")

        row = ttk.Frame(outer, style="Panel.TFrame")
        row.pack(fill="x")

        self.all_profile_objects_text = tk.Text(row, height=2, wrap="word", font=("Consolas", 11), bg=STYLE["input"], fg=STYLE["text"], insertbackground=STYLE["text"], relief="flat", bd=0, padx=10, pady=8)
        self.all_profile_objects_text.pack(side="left", fill="x", expand=True, padx=(0, 10))

        buttons = ttk.Frame(row, style="Panel.TFrame")
        buttons.pack(side="right", fill="y")
        ttk.Button(buttons, text="Generate String", style="Primary.TButton", command=self.generate_for_all_tabs).pack(fill="x", pady=(0, 6))
        ttk.Button(buttons, text="Convert String → Bits", style="Action.TButton", command=self.convert_string_to_bits).pack(fill="x")

    def _build_statusbar(self):
        status_frame = tk.Frame(self, bg="#4B70E7", height=28)
        status_frame.pack(fill="x", side="bottom")
        status_frame.pack_propagate(False)

        self.status = tk.Label(status_frame, text="Ready • Select profile objects and generate the string", bg="#4B70E7", fg="#D8E5F0", font=("Segoe UI", 11), anchor="w")
        self.status.pack(fill="both", padx=12)

    @staticmethod
    def _short_tab_name(name):
        return {
            "Load-Survey-Profile": "Load Survey",
            "Daily-Energy-Profile": "Daily Energy",
            "Billing-Profile": "Billing",
            "Event-Profile-1": "Event 1",
            "Event-Profile-2": "Event 2",
            "Nameplate-Profile": "Nameplate",
            "Instant-Profile": "Instantaneous",
        }.get(name, name)

    @staticmethod
    def _tab_heading(name):
        return {
            "Load-Survey-Profile": "Load Survey Objects",
            "Daily-Energy-Profile": "Daily Energy Objects",
            "Billing-Profile": "Billing Profile Objects",
            "Event-Profile-1": "Event Profile 1 Objects",
            "Event-Profile-2": "Event Profile 2 Objects",
            "Nameplate-Profile": "Nameplate Objects",
            "Instant-Profile": "Instantaneous Objects",
        }.get(name, name)

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
                output_text.configure(state="normal")
                output_text.delete("1.0", tk.END)
                output_text.insert(tk.END, display_text)
                output_text.configure(state="disabled")

        if self.all_profile_objects_text is not None:
            self.all_profile_objects_text.delete("1.0", tk.END)
            self.all_profile_objects_text.insert(tk.END, combined_profile_string)

        self.status.config(text="Generated strings for all profile tabs • Output updated")
        self.header_badge.config(text="GENERATED")
        messagebox.showinfo("Generation Complete", "Profile-object strings generated successfully.")

    def convert_string_to_bits(self):
        if self.all_profile_objects_text is None:
            return

        all_profile_objects_string = self.all_profile_objects_text.get("1.0", tk.END).strip()
        if not all_profile_objects_string:
            messagebox.showerror("Input Required", "Please generate or enter an All Profile Objects string first.")
            return

        if len(all_profile_objects_string) != 146:
            messagebox.showerror("Invalid String", "Input All Profile Objects Hex string is invalid. Check length.")
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
            messagebox.showerror("Invalid String", "Input All Profile Objects Hex string contains extra data.")
            return

        for tab_name, hex_string in values.items():
            bits_string = hex_string_to_bits(hex_string)
            self.apply_bits_to_panel(tab_name, bits_string)

        self.status.config(text="Loaded hexadecimal string into all profile tabs")
        self.header_badge.config(text="LOADED")
        messagebox.showinfo("Load Complete", "Hex string converted to bits and applied to all tabs.")

    def apply_bits_to_panel(self, tab_name, bits_string):
        panel = self.panels.get(tab_name)
        if panel is None:
            return False

        cleaned_bits = "".join(ch for ch in str(bits_string).strip() if ch in "01")
        if cleaned_bits.startswith("0b"):
            cleaned_bits = cleaned_bits[2:]

        expected_count = len(panel.items)
        if not cleaned_bits:
            panel.set_bits([0] * expected_count)
            panel._render()
            panel._update_count()
            return True

        if len(cleaned_bits) < expected_count:
            cleaned_bits = cleaned_bits.zfill(expected_count)
        elif len(cleaned_bits) > expected_count:
            cleaned_bits = cleaned_bits[-expected_count:]

        panel.set_bits([bit == "1" for bit in cleaned_bits])
        panel._render()
        panel._update_count()
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
    app = ProfileTabbedApp(profile_type, object_lists, loader.required_items)
    app.mainloop()


if __name__ == "__main__":
    launch_profile_selection()
