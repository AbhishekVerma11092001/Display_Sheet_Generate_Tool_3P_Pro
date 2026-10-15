import os
import re
import sys
from tkinter import simpledialog
from annotated_types import Len
import customtkinter as ctk
import tkinter as tk
import xml.etree.ElementTree as ET
from tkinter import filedialog, messagebox
from tkinter import ttk

# === Dialog and Function Definitions ===
class InputDialog(simpledialog.Dialog):
    def body(self, master):
        bit0 = tk.IntVar(value=0)
        chk_btn0 = ctk.CTkCheckBox(master, text="Metering Mode", variable=bit0)
        chk_btn0.grid(row=0, column=0, padx=30, sticky="w")

        bit1 = tk.IntVar(value=0)
        chk_btn1 = ctk.CTkCheckBox(master, text="Payment Mode", variable=bit1)
        chk_btn1.grid(row=1, column=0, padx=30, sticky="w")

        bit2 = tk.IntVar(value=0)
        chk_btn2 = ctk.CTkCheckBox(master, text="MD-IP", variable=bit2)
        chk_btn2.grid(row=2, column=0, padx=30, sticky="w")

        bit3 = tk.IntVar(value=0)
        chk_btn3 = ctk.CTkCheckBox(master, text="LS-IP", variable=bit3)
        chk_btn3.grid(row=3, column=0, padx=30, sticky="w")

        bit4 = tk.IntVar(value=0)
        chk_btn4 = ctk.CTkCheckBox(master, text="Firmware Upgrade", variable=bit4)
        chk_btn4.grid(row=4, column=0, padx=30, sticky="w")

        bit5 = tk.IntVar(value=0)
        chk_btn5 = ctk.CTkCheckBox(master, text="Apparent Cal Method", variable=bit5)
        chk_btn5.grid(row=5, column=0, padx=30, sticky="w")

        bit6 = tk.IntVar(value=0)
        # chk_btn6 = ctk.CTkCheckBox(master, text="Reserved", variable=bit6)
        # chk_btn6.grid(row=6, column=0, padx=30, sticky="w")

        bit7 = tk.IntVar(value=0)
        # chk_btn7 = ctk.CTkCheckBox(master, text="Reserved", variable=bit7)
        # chk_btn7.grid(row=7, column=0, padx=30, sticky="w")

        self.bit_vars = [bit7, bit6, bit5, bit4, bit3, bit2, bit1, bit0]
        return chk_btn0

    def apply(self):
        bill_string = "".join(str(var.get()) for var in self.bit_vars)
        self.result = bill_string


def bits_to_hex(bits_string):
    """Convert an 8-bit binary string to a 2-digit hex value."""
    if len(bits_string) == 8 and all(c in '01' for c in bits_string):
        hex_value = format(int(bits_string, 2), '04X')
        return hex_value
    return "00"


def get_bill_due_config():
    dialog = InputDialog(root, title="Bill due to Configuration")
    bill_string = dialog.result
    if bill_string is not None:
        bill_config_var.set(bill_string)


def generate_string():
    # Collect values from GUI elements
    tamper_log_time_type = tamper_log_time_type_var.get()
    tamper_detect_time = tamper_detect_time_var.get()
    if not tamper_detect_time.isdigit() or not (0 <= int(tamper_detect_time) <= 255):
        messagebox.showerror("Error", "Tamper Detect Time must be a number between 0 and 255.")
        return
    LS_Max_Entry = LS_Max_Entry_var.get()
    if not LS_Max_Entry.isdigit() or not (0 <= int(LS_Max_Entry) <= 4800):
        messagebox.showerror("Error", "Load Survey Max Entries must be a number between 0 and 4800.")
        return
    DE_Max_Entry = DE_Max_Entry_var.get()
    if not DE_Max_Entry.isdigit() or not (0 <= int(DE_Max_Entry) <= 100):
        messagebox.showerror("Error", "Daily Energy Max Entries must be a number between 0 and 100.")
        return
    BP_Max_Entry = BP_Max_Entry_var.get()
    if not BP_Max_Entry.isdigit() or not (int(BP_Max_Entry) == 12):
        messagebox.showerror("Error", "Bill Profile Max Entries must be a number equal to 12.")
        return
    Vol_Max_Entry = Vol_Max_Entry_var.get()
    if not Vol_Max_Entry.isdigit() or not (0 <= int(Vol_Max_Entry) <= 100):
        messagebox.showerror("Error", "Voltage Related Event Max Entries must be a number between 0 and 100.")
        return
    Curr_Max_Entry = Curr_Max_Entry_var.get()
    if not Curr_Max_Entry.isdigit() or not (0 <= int(Curr_Max_Entry) <= 150):
        messagebox.showerror("Error", "Current Related Event Max Entries must be a number between 0 and 150.")
        return
    Non_Roll_Max_Entry = Non_Roll_Max_Entry_var.get()
    if not Non_Roll_Max_Entry.isdigit() or not (0 <= int(Non_Roll_Max_Entry) <= 6):
        messagebox.showerror("Error", "Non-Rollover Event Max Entries must be a number between 0 and 6.")
        return
    Other_Max_Entry = Other_Max_Entry_var.get()
    if not Other_Max_Entry.isdigit() or not (0 <= int(Other_Max_Entry) <= 250):
        messagebox.showerror("Error", "Other Event Max Entries must be a number between 0 and 250.")
        return
    Trans_Max_Entry = Trans_Max_Entry_var.get()
    if not Trans_Max_Entry.isdigit() or not (0 <= int(Trans_Max_Entry) <= 100):
        messagebox.showerror("Error", "Transaction Event Max Entries must be a number between 0 and 100.")
        return
    Power_Max_Entry = Power_Max_Entry_var.get()
    if not Power_Max_Entry.isdigit() or not (0 <= int(Power_Max_Entry) <= 300):
        messagebox.showerror("Error", "Power Event Max Entries must be a number between 0 and 300.")
        return
    Control_Max_Entry = Control_Max_Entry_var.get()
    if not Control_Max_Entry.isdigit() or not (0 <= int(Control_Max_Entry) <= 100):
        messagebox.showerror("Error", "Control Event Max Entries must be a number between 0 and 100.")
        return
    display_lock = display_lock_var.get()
    relay_config = Relay_Config_var.get()
    bill_due_config = bill_config_var.get()
    active_led = Active_LED_var.get()
    app_calc = App_Calc_var.get()
    nd_defraud = ND_Defraud_var.get()
    push_timeout = push_timeout_var.get()
    if not push_timeout.isdigit() or not (0 <= int(push_timeout) <= 65535):
        messagebox.showerror("Error", "Push Mode Timeout must be a number between 0 and 65535.")
        return
    diagnostic_timeout = diagnostic_timeout_var.get()
    if not diagnostic_timeout.isdigit() or not (0 <= int(diagnostic_timeout) <= 65535):
        messagebox.showerror("Error", "Diagnostic Mode Timeout must be a number between 0 and 65535.")
        return
    test_timeout = test_timeout_var.get()
    if not test_timeout.isdigit() or not (0 <= int(test_timeout) <= 65535):
        messagebox.showerror("Error", "Test Mode Timeout must be a number between 0 and 65535.")
        return
    battery_mode = battery_mode_var.get()
    low_battery_threshold = low_battery_threshold_var.get()
    if not low_battery_threshold.isdigit() or not (0 <= int(low_battery_threshold) <= 65535):
        messagebox.showerror("Error", "Low Battery Detection Threshold must be a number between 0 and 65535.")
        return
    lrcf_postpaid = lrcf_postpaid_var.get()
    auto_mode_push = auto_mode_push_var.get()
    String = "0219"

    if tamper_log_time_type == "0-T2_T4":
        String += "1100"
    elif tamper_log_time_type == "1-T1_T3":
        String += "1101"

    if tamper_detect_time.isdigit() and 0 <= int(tamper_detect_time) <= 255:
        String = String + "11" 
        String += format(int(tamper_detect_time), '02X')

    if LS_Max_Entry.isdigit() and 0 <= int(LS_Max_Entry) <= 4800:
        String = String + "12" 
        String += format(int(LS_Max_Entry), '04X')

    if DE_Max_Entry.isdigit() and 0 <= int(DE_Max_Entry) <= 100:
        String = String + "12" 
        String += format(int(DE_Max_Entry), '04X')

    if BP_Max_Entry.isdigit() and 0 <= int(BP_Max_Entry) <= 12:
        String = String + "12" 
        String += format(int(BP_Max_Entry), '04X')

    if Vol_Max_Entry.isdigit() and 0 <= int(Vol_Max_Entry) <= 100:
        String = String + "12" 
        String += format(int(Vol_Max_Entry), '04X')

    if Curr_Max_Entry.isdigit() and 0 <= int(Curr_Max_Entry) <= 150:
        String = String + "12" 
        String += format(int(Curr_Max_Entry), '04X')

    if Non_Roll_Max_Entry.isdigit() and 0 <= int(Non_Roll_Max_Entry) <= 6:
        String = String + "12" 
        String += format(int(Non_Roll_Max_Entry), '04X')

    if Other_Max_Entry.isdigit() and 0 <= int(Other_Max_Entry) <= 250:
        String = String + "12" 
        String += format(int(Other_Max_Entry), '04X')

    if Trans_Max_Entry.isdigit() and 0 <= int(Trans_Max_Entry) <= 100:
        String = String + "12" 
        String += format(int(Trans_Max_Entry), '04X')

    if Power_Max_Entry.isdigit() and 0 <= int(Power_Max_Entry) <= 300:
        String = String + "12" 
        String += format(int(Power_Max_Entry), '04X')

    if Control_Max_Entry.isdigit() and 0 <= int(Control_Max_Entry) <= 100:
        String = String + "12" 
        String += format(int(Control_Max_Entry), '04X')

    if display_lock == "0-No Mode lock":
        String += "1100"
    elif display_lock == "1-Auto Mode lock":
        String += "1101"
    elif display_lock == "2-All Mode lock":
        String += "1102"

    if relay_config == "0-YMPL Lab config":
        String += "1100"
    elif relay_config == "1-All Projects Config":
        String += "1101"

    String += "12"
    bits_string = bits_to_hex(bill_due_config)
    String += bits_string

    if active_led == "0-Total(Import+Export)":
        String += "1100"
    elif active_led == "1-NET(Import-Export)":
        String += "1101"

    if app_calc == "0-Square Root":
        String += "1100"
    elif app_calc == "1-VI Method":
        String += "1101"

    if nd_defraud == "0-Disabled":
        String += "1100"
    elif nd_defraud == "1-Enabled":
        String += "1101"

    if push_timeout.isdigit() and 0 <= int(push_timeout) <= 65535:
        String = String + "12" 
        String += format(int(push_timeout), '04X')

    if diagnostic_timeout.isdigit() and 0 <= int(diagnostic_timeout) <= 65535:
        String = String + "12" 
        String += format(int(diagnostic_timeout), '04X')

    if test_timeout.isdigit() and 0 <= int(test_timeout) <= 65535:
        String = String + "12" 
        String += format(int(test_timeout), '04X')

    if battery_mode == "0-Disabled":
        String += "120000"
    elif battery_mode == "1-Enabled":
        String += "120001"

    if low_battery_threshold.isdigit() and 0 <= int(low_battery_threshold) <= 65535:
        String = String + "12" 
        String += format(int(low_battery_threshold), '04X')

    if lrcf_postpaid == "0-Disabled":
        String += "1100"
    elif lrcf_postpaid == "1-Enabled":
        String += "1101"

    if auto_mode_push == "0-Disabled":
        String += "1100"
    elif auto_mode_push == "1-Enabled":
        String += "1101"

    output_textbox.delete("1.0", "end")
    output_textbox.insert("1.0", String)
    messagebox.showinfo("Success", "String generated successfully.")


def convert_string_to_bits():
    Payload = output_textbox.get("1.0", "end-1c").strip()
    if not Payload:
        messagebox.showerror("Error", "Please enter the string first.")
        return

    Length = len(Payload)

    if Length != 136:
        messagebox.showerror("Error", "Please enter the valid string.")
        return

    if Payload.startswith("0219"):
        Payload = Payload[4:]

    if Payload.startswith("1100"):
        tamper_log_time_type_var.set("0-T2_T4")
        Payload = Payload[4:]
    elif Payload.startswith("1101"):
        tamper_log_time_type_var.set("1-T1_T3")
        Payload = Payload[4:]

    if Payload.startswith("11"):
        tamper_detect_time_hex = Payload[2:4]
        tamper_detect_time_var.set(str(int(tamper_detect_time_hex, 16)))
        Payload = Payload[4:]

    if Payload.startswith("12"):
        LS_Max_Entry_hex = Payload[2:6]
        LS_Max_Entry_var.set(str(int(LS_Max_Entry_hex, 16)))
        Payload = Payload[6:]

    if Payload.startswith("12"):
        DE_Max_Entry_hex = Payload[2:6]
        DE_Max_Entry_var.set(str(int(DE_Max_Entry_hex, 16)))
        Payload = Payload[6:]

    if Payload.startswith("12"):
        BP_Max_Entry_hex = Payload[2:6]
        BP_Max_Entry_var.set(str(int(BP_Max_Entry_hex, 16)))
        Payload = Payload[6:]

    if Payload.startswith("12"):
        Vol_Max_Entry_hex = Payload[2:6]
        Vol_Max_Entry_var.set(str(int(Vol_Max_Entry_hex, 16)))
        Payload = Payload[6:]

    if Payload.startswith("12"):
        Curr_Max_Entry_hex = Payload[2:6]
        Curr_Max_Entry_var.set(str(int(Curr_Max_Entry_hex, 16)))
        Payload = Payload[6:]

    if Payload.startswith("12"):
        Non_Roll_Max_Entry_hex = Payload[2:6]
        Non_Roll_Max_Entry_var.set(str(int(Non_Roll_Max_Entry_hex, 16)))
        Payload = Payload[6:]

    if Payload.startswith("12"):
        Other_Max_Entry_hex = Payload[2:6]
        Other_Max_Entry_var.set(str(int(Other_Max_Entry_hex, 16)))
        Payload = Payload[6:]

    if Payload.startswith("12"):
        Trans_Max_Entry_hex = Payload[2:6]
        Trans_Max_Entry_var.set(str(int(Trans_Max_Entry_hex, 16)))
        Payload = Payload[6:]

    if Payload.startswith("12"):
        Power_Max_Entry_hex = Payload[2:6]
        Power_Max_Entry_var.set(str(int(Power_Max_Entry_hex, 16)))
        Payload = Payload[6:]

    if Payload.startswith("12"):
        Control_Max_Entry_hex = Payload[2:6]
        Control_Max_Entry_var.set(str(int(Control_Max_Entry_hex, 16)))
        Payload = Payload[6:]

    if Payload.startswith("1100"):
        display_lock_var.set("0-No Mode lock")
        Payload = Payload[4:]
    elif Payload.startswith("1101"):
        display_lock_var.set("1-Auto Mode lock")
        Payload = Payload[4:]

    if Payload.startswith("1100"):
        Relay_Config_var.set("0-YMPL Lab config")
        Payload = Payload[4:]
    elif Payload.startswith("1101"):
        Relay_Config_var.set("1-All Projects Config")
        Payload = Payload[4:]

    if Payload.startswith("12"):
        bill_due_config_hex = Payload[4:6]
        bill_due_config_bits = format(int(bill_due_config_hex, 16), '08b')
        bill_config_var.set(bill_due_config_bits)
        Payload = Payload[6:]

    if Payload.startswith("1100"):
        Active_LED_var.set("0-Total(Import+Export)")
        Payload = Payload[4:]
    elif Payload.startswith("1101"):
        Active_LED_var.set("1-NET(Import-Export)")
        Payload = Payload[4:]

    if Payload.startswith("1100"):
        App_Calc_var.set("0-Square Root")
        Payload = Payload[4:]
    elif Payload.startswith("1101"):
        App_Calc_var.set("1-VI Method")
        Payload = Payload[4:]

    if Payload.startswith("1100"):
        ND_Defraud_var.set("0-Disabled")
        Payload = Payload[4:]
    elif Payload.startswith("1101"):
        ND_Defraud_var.set("1-Enabled")
        Payload = Payload[4:]

    if Payload.startswith("12"):
        push_timeout_hex = Payload[2:6]
        push_timeout_var.set(str(int(push_timeout_hex, 16)))
        Payload = Payload[6:]

    if Payload.startswith("12"):
        diagnostic_timeout_hex = Payload[2:6]
        diagnostic_timeout_var.set(str(int(diagnostic_timeout_hex, 16)))
        Payload = Payload[6:]

    if Payload.startswith("12"):
        test_timeout_hex = Payload[2:6]
        test_timeout_var.set(str(int(test_timeout_hex, 16)))
        Payload = Payload[6:]

    if Payload.startswith("120000"):
        battery_mode_var.set("0-Disabled")
        Payload = Payload[6:]
    elif Payload.startswith("120001"):
        battery_mode_var.set("1-Enabled")
        Payload = Payload[6:]

    if Payload.startswith("12"):
        low_battery_threshold_hex = Payload[2:6]
        low_battery_threshold_var.set(str(int(low_battery_threshold_hex, 16)))
        Payload = Payload[6:]

    if Payload.startswith("1100"):
        lrcf_postpaid_var.set("0-Disabled")
        Payload = Payload[4:]
    elif Payload.startswith("1101"):
        lrcf_postpaid_var.set("1-Enabled")
        Payload = Payload[4:]

    if Payload.startswith("1100"):
        auto_mode_push_var.set("0-Disabled")
        Payload = Payload[4:]
    elif Payload.startswith("1101"):
        auto_mode_push_var.set("1-Enabled")
        Payload = Payload[4:]

    messagebox.showinfo("Success", "String converted to bits successfully.")


# === Modern Industrial GUI ===
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

COLORS = {
    "bg": "#F3F6F9", "surface": "#FFFFFF", "input": "#E0EAF5",
    "text": "#17212B", "muted": "#667381", "primary": "#1D55CE",
    "primary_hover": "#6A91E7", "border": "#989A9C", "success": "#5E8BEC",
    "success_hover": "#9EB6E9",
}

root = ctk.CTk()
root.title("3P Pro Top Level String Selection | V2.0")
root.configure(fg_color=COLORS["bg"])
root.resizable(False, False)

WIN_W, WIN_H = 1100, 720
root.update_idletasks()
pos_x = max((root.winfo_screenwidth() - WIN_W) // 2, 0)
pos_y = max((root.winfo_screenheight() - WIN_H) // 2, 0)
root.geometry(f"{WIN_W}x{WIN_H}+{pos_x}+{pos_y}")
root.lift(); root.focus_force()
root.after(100, lambda: root.attributes("-topmost", True))
root.after(300, lambda: root.attributes("-topmost", False))

# Header
header = ctk.CTkFrame(root, fg_color="#1D55CE", corner_radius=0, height=35)
header.pack(fill="x")
header.pack_propagate(False)
ctk.CTkLabel(header, text="3P PRO TOP LEVEL STRING SELECTION", text_color="white", font=ctk.CTkFont(size=18, weight="bold")).pack(anchor="w", padx=20, pady=(10, 0))

content = ctk.CTkFrame(root, fg_color=COLORS["bg"], corner_radius=0)
content.pack(fill="both", expand=True, padx=18, pady=14)

left_frame = ctk.CTkFrame(content, fg_color=COLORS["surface"], border_width=1, border_color=COLORS["border"], corner_radius=12)
left_frame.pack(side="left", fill="both", expand=True, padx=(0, 7))
right_frame = ctk.CTkFrame(content, fg_color=COLORS["surface"], border_width=1, border_color=COLORS["border"], corner_radius=12)
right_frame.pack(side="right", fill="both", expand=True, padx=(7, 0))

ctk.CTkLabel(left_frame, text="Profile & Event Configuration", text_color=COLORS["primary"], font=ctk.CTkFont(size=16, weight="bold")).grid(row=0, column=0, columnspan=2, padx=20, pady=(8, 0), sticky="w")
ctk.CTkLabel(right_frame, text="Meter & Operating Configuration", text_color=COLORS["primary"], font=ctk.CTkFont(size=16, weight="bold")).grid(row=0, column=0, columnspan=3, padx=20, pady=(8, 0), sticky="w")

LABEL_FONT = ctk.CTkFont(size=14)
ENTRY_FONT = ctk.CTkFont(size=14)


def add_entry(parent, row, label, variable, width=180):
    ctk.CTkLabel(parent, text=label, text_color=COLORS["text"], font=LABEL_FONT).grid(row=row, column=0, padx=20, pady=5, sticky="w")
    w = ctk.CTkEntry(parent, textvariable=variable, width=width, height=20, font=ENTRY_FONT, fg_color=COLORS["input"], border_color=COLORS["border"], text_color=COLORS["text"])
    w.grid(row=row, column=1, padx=(5, 20), pady=5, sticky="w")
    return w


def add_combo(parent, row, label, variable, values, width=180):
    ctk.CTkLabel(parent, text=label, text_color=COLORS["text"], font=LABEL_FONT).grid(row=row, column=0, padx=20, pady=5, sticky="w")
    w = ctk.CTkComboBox(parent, variable=variable, values=values, state="readonly", width=width, height=20, font=ENTRY_FONT, fg_color=COLORS["input"], border_color=COLORS["border"], button_color=COLORS["primary_hover"], button_hover_color=COLORS["primary"], text_color=COLORS["text"])
    w.grid(row=row, column=1, padx=(5, 20), pady=5, sticky="w")
    return w

# Variables + controls (same names/values used by generation logic)
tamper_log_time_type_var = tk.StringVar(value="1-T1_T3")
tamper_log_time_type_combo = add_combo(left_frame, 1, "Tamper Log Time Type", tamper_log_time_type_var, ["0-T2_T4", "1-T1_T3"])
tamper_detect_time_var = tk.StringVar(value="10"); tamper_detect_time_combo = add_entry(left_frame, 2, "Tamper Detect Time (0-255)", tamper_detect_time_var)
LS_Max_Entry_var = tk.StringVar(value="2880"); LS_Max_Entry_combo = add_entry(left_frame, 3, "Load Survey Max Entries (≤4800)", LS_Max_Entry_var)
DE_Max_Entry_var = tk.StringVar(value="60"); DE_Max_Entry_combo = add_entry(left_frame, 4, "Daily Energy Max Entries (≤100)", DE_Max_Entry_var)
BP_Max_Entry_var = tk.StringVar(value="12"); BP_Max_Entry_combo = add_entry(left_frame, 5, "Bill Profile Max Entries (12)", BP_Max_Entry_var)
Vol_Max_Entry_var = tk.StringVar(value="80"); Vol_Max_Entry_combo = add_entry(left_frame, 6, "Voltage Event Max Entries (≤100)", Vol_Max_Entry_var)
Curr_Max_Entry_var = tk.StringVar(value="80"); Curr_Max_Entry_combo = add_entry(left_frame, 7, "Current Event Max Entries (≤150)", Curr_Max_Entry_var)
Non_Roll_Max_Entry_var = tk.StringVar(value="2"); Non_Roll_Max_Entry_combo = add_entry(left_frame, 8, "Non-Rollover Event Max Entries (≤6)", Non_Roll_Max_Entry_var)
Other_Max_Entry_var = tk.StringVar(value="40"); Other_Max_Entry_combo = add_entry(left_frame, 9, "Other Event Max Entries (≤250)", Other_Max_Entry_var)
Trans_Max_Entry_var = tk.StringVar(value="24"); Trans_Max_Entry_combo = add_entry(left_frame, 10, "Transaction Event Max Entries (≤100)", Trans_Max_Entry_var)
Power_Max_Entry_var = tk.StringVar(value="288"); Power_Max_Entry_combo = add_entry(left_frame, 11, "Power Event Max Entries (≤300)", Power_Max_Entry_var)
Control_Max_Entry_var = tk.StringVar(value="64"); Control_Max_Entry_combo = add_entry(left_frame, 12, "Control Event Max Entries (≤100)", Control_Max_Entry_var)
display_lock_var = tk.StringVar(value="0-No Mode lock")
display_lock_combo = add_combo(left_frame, 13, "C-Open Lock Display Configuration", display_lock_var, ["0-No Mode lock", "1-Auto Mode lock", "2-All Mode lock"])

Relay_Config_var = tk.StringVar(value="1-All Projects Config"); Relay_Config_combo = add_combo(right_frame, 1, "Relay Configuration", Relay_Config_var, ["0-YMPL Lab config", "1-All Projects Config"])
bill_config_var = tk.StringVar(value="00000000")
ctk.CTkLabel(right_frame, text="Bill due to Configuration", text_color=COLORS["text"], font=LABEL_FONT).grid(row=2, column=0, padx=20, pady=5, sticky="w")
bill_config_entry = ctk.CTkEntry(right_frame, textvariable=bill_config_var, state="readonly", width=180, height=20, font=ENTRY_FONT, fg_color=COLORS["input"], border_color=COLORS["border"], text_color=COLORS["text"]); bill_config_entry.grid(row=2,column=1,padx=(5,5),pady=5,sticky="w")
bill_config_button = ctk.CTkButton(right_frame, text="Select", width=78, height=20, command=get_bill_due_config, font=ENTRY_FONT, fg_color=COLORS["primary"], hover_color=COLORS["primary_hover"]); bill_config_button.grid(row=2,column=2,padx=(0,20),pady=5)
Active_LED_var = tk.StringVar(value="0-Total(Import+Export)"); Active_LED_combo = add_combo(right_frame, 3, "Active LED Configuration", Active_LED_var, ["0-Total(Import+Export)", "1-NET(Import-Export)"])
App_Calc_var = tk.StringVar(value="0-Square Root"); App_Calc_combo = add_combo(right_frame, 4, "Apparent Calculation Method", App_Calc_var, ["0-Square Root", "1-VI Method"])
ND_Defraud_var = tk.StringVar(value="0-Disabled"); ND_Defraud_combo = add_combo(right_frame, 5, "Defraud at ND Tamper", ND_Defraud_var, ["0-Disabled", "1-Enabled"])
push_timeout_var = tk.StringVar(value="10"); push_timeout_combo = add_entry(right_frame, 6, "Push Mode Timeout", push_timeout_var)
diagnostic_timeout_var = tk.StringVar(value="900"); diagnostic_timeout_combo = add_entry(right_frame, 7, "Diagnostic Mode Timeout", diagnostic_timeout_var)
test_timeout_var = tk.StringVar(value="60"); test_timeout_combo = add_entry(right_frame, 8, "Test Mode Timeout", test_timeout_var)
battery_mode_var = tk.StringVar(value="0-Disabled"); battery_mode_combo = add_combo(right_frame, 9, "Battery Mode Requirement", battery_mode_var, ["0-Disabled", "1-Enabled"])
low_battery_threshold_var = tk.StringVar(value="1900"); low_battery_threshold_combo = add_entry(right_frame, 10, "Low Battery Detection Threshold", low_battery_threshold_var)
lrcf_postpaid_var = tk.StringVar(value="0-Disabled"); lrcf_postpaid_combo = add_combo(right_frame, 11, "LRCF Required in Postpaid Mode", lrcf_postpaid_var, ["0-Disabled", "1-Enabled"])
auto_mode_push_var = tk.StringVar(value="0-Disabled"); auto_mode_push_combo = add_combo(right_frame, 12, "Auto Mode with Push Functionality", auto_mode_push_var, ["0-Disabled", "1-Enabled"])

# Output/action card
bottom = ctk.CTkFrame(root, fg_color=COLORS["surface"], border_width=1, border_color=COLORS["border"], corner_radius=12)
bottom.pack(fill="x", padx=18, pady=(0, 15))
ctk.CTkLabel(bottom, text="Generated Configuration String", text_color=COLORS["primary"], font=ctk.CTkFont(size=14, weight="bold")).grid(row=0,column=0,padx=18,pady=(12,5),sticky="w")
output_textbox = ctk.CTkTextbox(bottom, height=50, font=("Consolas", 14), wrap="word", fg_color=COLORS["input"], border_width=1, border_color=COLORS["border"])
output_textbox.grid(row=1,column=0,columnspan=2,padx=18,pady=(0,12),sticky="ew")
output_textbox.insert("1.0", "02191101110A120B4012003C12000C1200501200501200021200281200181201201200401100110112000011001100110012000A12038412003C12000012076C11001100")
btns = ctk.CTkFrame(bottom, fg_color="transparent"); btns.grid(row=0,column=2,rowspan=2,padx=(5,18),pady=10)
convert_btn = ctk.CTkButton(btns, text="Generate String",font=ctk.CTkFont(size=14), command=generate_string, width=150, height=28, fg_color=COLORS["primary"], hover_color=COLORS["primary_hover"]); convert_btn.pack(pady=4)
convert_btn2 = ctk.CTkButton(btns, text="Convert String → Bits", font=ctk.CTkFont(size=14), command=convert_string_to_bits, width=150, height=28, fg_color=COLORS["success"], hover_color=COLORS["success_hover"]); convert_btn2.pack(pady=4)
bottom.grid_columnconfigure(0, weight=1)

if __name__ == "__main__":
    root.mainloop()
