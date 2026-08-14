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


# === Tkinter GUI ===
ctk.set_appearance_mode("light")   # Options: "light", "dark", "system"
ctk.set_default_color_theme("blue")  # Try "green", "dark-blue", etc.
# === Main Window ===
root = ctk.CTk()
# root.iconbitmap("logo.ico")
root.title("3P Pro Top Level Selection V1.0")
root.geometry("1100x600")
root.resizable(False, False)

# === Top Frame ===
top_frame = ctk.CTkFrame(root, corner_radius=15)
top_frame.pack(pady=15, padx=15, fill="x")

top_label = ctk.CTkLabel(top_frame, text="3P Pro Top Level String Selection", font=ctk.CTkFont(size=20, weight="bold"))
top_label.pack(pady=10, padx=10)

left_frame = ctk.CTkFrame(top_frame)
left_frame.pack(side="left", fill="both", expand=True, padx=(8, 4), pady=8)

right_frame = ctk.CTkFrame(top_frame)
right_frame.pack(side="right", fill="both", expand=True, padx=(4, 8), pady=8)

bottom_frame = ctk.CTkFrame(root)
bottom_frame.pack(fill="x", padx=8, pady=(0, 8))

bottom_left_frame = ctk.CTkFrame(bottom_frame)
bottom_left_frame.pack(side="left", fill="both", expand=True, padx=(8, 6), pady=8)

bottom_right_frame = ctk.CTkFrame(bottom_frame)
bottom_right_frame.pack(side="right", fill="both", expand=True, padx=(6, 8), pady=8)

# === GUI Elements ===
ctk.CTkLabel(left_frame, text="Tamper Log Time Type:").grid(row=0, column=0, padx=20, sticky="w")
tamper_log_time_type_var = tk.StringVar(value="1-T1_T3")
tamper_log_time_type_combo = ctk.CTkComboBox(left_frame, variable=tamper_log_time_type_var, values=["0-T2_T4", "1-T1_T3"], state="readonly", width=150)
tamper_log_time_type_combo.grid(row=0, column=1, padx=5, sticky="w")

ctk.CTkLabel(left_frame, text="Tamper Detect Time for Display: range(0-255)").grid(row=1, column=0, padx=20, sticky="w")
tamper_detect_time_var = tk.StringVar(value="10")
tamper_detect_time_combo = ctk.CTkEntry(left_frame, textvariable=tamper_detect_time_var, width=100)
tamper_detect_time_combo.grid(row=1, column=1, padx=5, sticky="w")

ctk.CTkLabel(left_frame, text="Load Survey Max Entries (upto 4800):").grid(row=2, column=0, padx=20, sticky="w")
LS_Max_Entry_var = tk.StringVar(value="2880")
LS_Max_Entry_combo = ctk.CTkEntry(left_frame, textvariable=LS_Max_Entry_var, width=100)
LS_Max_Entry_combo.grid(row=2, column=1, padx=5, sticky="w")

ctk.CTkLabel(left_frame, text="Daily Energy Max Entries (upto 100):").grid(row=3, column=0, padx=20, sticky="w")
DE_Max_Entry_var = tk.StringVar(value="60")
DE_Max_Entry_combo = ctk.CTkEntry(left_frame, textvariable=DE_Max_Entry_var, width=100)
DE_Max_Entry_combo.grid(row=3, column=1, padx=5, sticky="w")

ctk.CTkLabel(left_frame, text="Bill Profile  Max Entries (12):").grid(row=4, column=0, padx=20, sticky="w")
BP_Max_Entry_var = tk.StringVar(value="12")
BP_Max_Entry_combo = ctk.CTkEntry(left_frame, textvariable=BP_Max_Entry_var, width=100)
BP_Max_Entry_combo.grid(row=4, column=1, padx=5, sticky="w")

ctk.CTkLabel(left_frame, text="Voltage Related Event Max Entries (upto 100):").grid(row=5, column=0, padx=20, sticky="w")
Vol_Max_Entry_var = tk.StringVar(value="80")
Vol_Max_Entry_combo = ctk.CTkEntry(left_frame, textvariable=Vol_Max_Entry_var, width=100)
Vol_Max_Entry_combo.grid(row=5, column=1, padx=5, sticky="w")

ctk.CTkLabel(left_frame, text="Current Related Event Max Entries (upto 150):").grid(row=6, column=0, padx=20, sticky="w")
Curr_Max_Entry_var = tk.StringVar(value="80")
Curr_Max_Entry_combo = ctk.CTkEntry(left_frame, textvariable=Curr_Max_Entry_var, width=100)
Curr_Max_Entry_combo.grid(row=6, column=1, padx=5, sticky="w")

ctk.CTkLabel(left_frame, text="Non-Rollover Event Max Entries (upto 6):").grid(row=7, column=0, padx=20, sticky="w")
Non_Roll_Max_Entry_var = tk.StringVar(value="2")
Non_Roll_Max_Entry_combo = ctk.CTkEntry(left_frame, textvariable=Non_Roll_Max_Entry_var, width=100)
Non_Roll_Max_Entry_combo.grid(row=7, column=1, padx=5, sticky="w")

ctk.CTkLabel(left_frame, text="Other Event Max Entries (upto 250):").grid(row=8, column=0, padx=20, sticky="w")
Other_Max_Entry_var = tk.StringVar(value="40")
Other_Max_Entry_combo = ctk.CTkEntry(left_frame, textvariable=Other_Max_Entry_var, width=100)
Other_Max_Entry_combo.grid(row=8, column=1, padx=5, sticky="w")

ctk.CTkLabel(left_frame, text="Transaction Event Max Entries (upto 100):").grid(row=9, column=0, padx=20, sticky="w")
Trans_Max_Entry_var = tk.StringVar(value="24")
Trans_Max_Entry_combo = ctk.CTkEntry(left_frame, textvariable=Trans_Max_Entry_var, width=100)
Trans_Max_Entry_combo.grid(row=9, column=1, padx=5, sticky="w")

ctk.CTkLabel(left_frame, text="Power Event Max Entries (upto 300):").grid(row=10, column=0, padx=20, sticky="w")
Power_Max_Entry_var = tk.StringVar(value="288")
Power_Max_Entry_combo = ctk.CTkEntry(left_frame, textvariable=Power_Max_Entry_var, width=100)
Power_Max_Entry_combo.grid(row=10, column=1, padx=5, sticky="w")

ctk.CTkLabel(left_frame, text="Control Event Max Entries (upto 100):").grid(row=11, column=0, padx=20, sticky="w")
Control_Max_Entry_var = tk.StringVar(value="64")
Control_Max_Entry_combo = ctk.CTkEntry(left_frame, textvariable=Control_Max_Entry_var, width=100)
Control_Max_Entry_combo.grid(row=11, column=1, padx=5, sticky="w")

ctk.CTkLabel(left_frame, text="C-Open Lock Display Configuration:").grid(row=12, column=0, padx=20, sticky="w")
display_lock_var = tk.StringVar(value="0-No Mode lock")
display_lock_combo = ctk.CTkComboBox(left_frame, variable=display_lock_var, values=["0-No Mode lock", "1-Auto Mode lock", "2-All Mode lock"], state="readonly", width=150)
display_lock_combo.grid(row=12, column=1, padx=5, sticky="w")

ctk.CTkLabel(right_frame, text="Relay Configuration:").grid(row=0, column=0, padx=20, sticky="w")
Relay_Config_var = tk.StringVar(value="1-All Projects Config")
Relay_Config_combo = ctk.CTkComboBox(right_frame, variable=Relay_Config_var, values=["0-YMPL Lab config", "1-All Projects Config"], state="readonly", width=200)
Relay_Config_combo.grid(row=0, column=1, padx=5, sticky="w")

ctk.CTkLabel(right_frame, text="Bill due to Configuration:").grid(row=1, column=0, padx=20, sticky="w")
bill_config_var = tk.StringVar(value="00000000")
bill_config_entry = ctk.CTkEntry(right_frame, textvariable=bill_config_var, state="readonly", width=100)
bill_config_entry.grid(row=1, column=1, padx=5, sticky="w")
bill_config_button = ctk.CTkButton(right_frame, text="Select", width=60, height=20, command=get_bill_due_config, fg_color="#782E86")
bill_config_button.grid(row=1, column=2, padx=5, sticky="w")


ctk.CTkLabel(right_frame, text="Active LED Configuration:").grid(row=2, column=0, padx=20, sticky="w")
Active_LED_var = tk.StringVar(value="0-Total(Import+Export)")
Active_LED_combo = ctk.CTkComboBox(right_frame, variable=Active_LED_var, values=["0-Total(Import+Export)", "1-NET(Import-Export)"], state="readonly", width=200)
Active_LED_combo.grid(row=2, column=1, padx=5, sticky="w")

ctk.CTkLabel(right_frame, text="Apparent Calculation Method:").grid(row=3, column=0, padx=20, sticky="w")
App_Calc_var = tk.StringVar(value="0-Square Root")
App_Calc_combo = ctk.CTkComboBox(right_frame, variable=App_Calc_var, values=["0-Square Root", "1-VI Method"], state="readonly", width=200)
App_Calc_combo.grid(row=3, column=1, padx=5, sticky="w")

ctk.CTkLabel(right_frame, text="Defraud at ND Tamper:").grid(row=4, column=0, padx=20, sticky="w")
ND_Defraud_var = tk.StringVar(value="0-Disabled")
ND_Defraud_combo = ctk.CTkComboBox(right_frame, variable=ND_Defraud_var, values=["0-Disabled", "1-Enabled"], state="readonly", width=200)
ND_Defraud_combo.grid(row=4, column=1, padx=5, sticky="w")

ctk.CTkLabel(right_frame, text="Push Mode Timeout:").grid(row=5, column=0, padx=20, sticky="w")
push_timeout_var = tk.StringVar(value="10")
push_timeout_combo = ctk.CTkEntry(right_frame, textvariable=push_timeout_var, width=100)
push_timeout_combo.grid(row=5, column=1, padx=5, sticky="w")

ctk.CTkLabel(right_frame, text="Diagnostic Mode Timeout:").grid(row=6, column=0, padx=20, sticky="w")
diagnostic_timeout_var = tk.StringVar(value="900")
diagnostic_timeout_combo = ctk.CTkEntry(right_frame, textvariable=diagnostic_timeout_var, width=100)
diagnostic_timeout_combo.grid(row=6, column=1, padx=5, sticky="w")

ctk.CTkLabel(right_frame, text="Test Mode Timeout:").grid(row=7, column=0, padx=20, sticky="w")
test_timeout_var = tk.StringVar(value="60")
test_timeout_combo = ctk.CTkEntry(right_frame, textvariable=test_timeout_var, width=100)
test_timeout_combo.grid(row=7, column=1, padx=5, sticky="w")

ctk.CTkLabel(right_frame, text="Battery Mode Requirement :").grid(row=8, column=0, padx=20, sticky="w")
battery_mode_var = tk.StringVar(value="0-Disabled")
battery_mode_combo = ctk.CTkComboBox(right_frame, variable=battery_mode_var, values=["0-Disabled", "1-Enabled"], state="readonly", width=200)
battery_mode_combo.grid(row=8, column=1, padx=5, sticky="w")

ctk.CTkLabel(right_frame, text="Low Battery Detection Threshold:").grid(row=9, column=0, padx=20, sticky="w")
low_battery_threshold_var = tk.StringVar(value="1900")
low_battery_threshold_combo = ctk.CTkEntry(right_frame, textvariable=low_battery_threshold_var, width=100)
low_battery_threshold_combo.grid(row=9, column=1, padx=5, sticky="w")

ctk.CTkLabel(right_frame, text="LRCF required in Postpaid Mode:").grid(row=10, column=0, padx=20, sticky="w")
lrcf_postpaid_var = tk.StringVar(value="0-Disabled")
lrcf_postpaid_combo = ctk.CTkComboBox(right_frame, variable=lrcf_postpaid_var, values=["0-Disabled", "1-Enabled"], state="readonly", width=200)
lrcf_postpaid_combo.grid(row=10, column=1, padx=5, sticky="w")

ctk.CTkLabel(right_frame, text="Auto Mode with Push Functionality:").grid(row=11, column=0, padx=20, sticky="w")
auto_mode_push_var = tk.StringVar(value="0-Disabled")
auto_mode_push_combo = ctk.CTkComboBox(right_frame, variable=auto_mode_push_var, values=["0-Disabled", "1-Enabled"], state="readonly", width=200)
auto_mode_push_combo.grid(row=11, column=1, padx=5, sticky="w")


ctk.CTkLabel(bottom_left_frame, text="Output String:", font=(None, 15, "bold")).grid(row=0, column=0, padx=20, sticky="w")
output_textbox = ctk.CTkTextbox(bottom_left_frame, height=50, width=800, font=(None, 15), wrap="word")
output_textbox.grid(row=1, column=0, padx=20, pady=(0, 10), sticky="w")  

convert_btn = ctk.CTkButton(bottom_right_frame, text="Generate String", command=generate_string, fg_color="#4946f1")
convert_btn.grid(row=1, column=1, padx=20, pady=(6, 0), sticky="e")


convert_btn = ctk.CTkButton(bottom_right_frame, text="Convert String to Bits", command=convert_string_to_bits, fg_color="#317a41")
convert_btn.grid(row=2, column=1, padx=20, pady=(6, 0), sticky="e")

if __name__ == "__main__":
    root.mainloop()