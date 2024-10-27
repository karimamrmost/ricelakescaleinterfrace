import serial
import pandas as pd
from time import sleep
import datetime
import logging
import tkinter as tk
from tkinter import messagebox, filedialog
from PIL import Image, ImageTk

# Setup logging
logging.basicConfig(filename='scale_errors.log', level=logging.ERROR)

# Function to initialize the serial port
def init_serial(port, baudrate=9600, timeout=1):
    try:
        ser = serial.Serial(port, baudrate, timeout=timeout)
        return ser
    except Exception as e:
        logging.error(f"Error opening serial port: {e}")
        return None

# Function to send a command and read the response
def send_command(ser, command, retries=3):
    for _ in range(retries):
        try:
            ser.write(command.encode())
            sleep(1)
            response = ser.readline().decode().strip()
            if response:
                return response
            else:
                logging.warning(f"Empty response for command: {command}")
        except Exception as e:
            logging.error(f"Error sending command {command}: {e}")
    return None

# Function to zero the scale using the serial command SC.REZERO#1
def zero_scale(ser):
    zero_command = "SC.REZERO#1\r\n"
    response = send_command(ser, zero_command)
    if response == "OK":
        messagebox.showinfo("Success", "Scale zeroed successfully.")
    else:
        logging.error(f"Failed to zero scale: {response}")
        messagebox.showerror("Error", f"Failed to zero scale: {response}")

# Function to read the weight using the serial command SC.GROSS#1
def read_weight(ser):
    weight_command = "SC.GROSS#1\r\n"
    response = send_command(ser, weight_command)
    if response:
        try:
            weight = float(response)
            return weight
        except ValueError:
            logging.error(f"Invalid weight response: {response}")
            messagebox.showerror("Error", f"Invalid weight response: {response}")
            return None
    return None

# Function to log data into an Excel file
def log_to_excel(first_measurement, second_measurement, filename):
    data = {
        'Timestamp': [datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")],
        'First Measurement': [first_measurement],
        'Second Measurement': [second_measurement]
    }
    df = pd.DataFrame(data)
    # Append to an existing file or create a new one
    try:
        with pd.ExcelWriter(filename, mode='a', if_sheet_exists='overlay', engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Log', index=False, header=not writer.sheets)
    except FileNotFoundError:
        df.to_excel(filename, sheet_name='Log', index=False)
    messagebox.showinfo("Success", f"Logged measurements to {filename}")

# GUI Application
class ScaleApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Scale Controller")

        # Load a PNG image for both the window and taskbar icon
        icon_image = ImageTk.PhotoImage(file='window_icon.png')  # Use .png or .gif file
        self.root.iconphoto(False, icon_image)  # Set window and taskbar icon

        # Set window size
        window_width = 450  # Adjusted width
        window_height = 450
        self.root.geometry(f"{window_width}x{window_height}")

        # Center the window on the screen
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        position_top = int(screen_height / 2 - window_height / 2)
        position_right = int(screen_width / 2 - window_width / 2)
        self.root.geometry(f"+{position_right}+{position_top}")

        # Configure the grid layout to have equal weights for both columns
        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_columnconfigure(1, weight=1)

        # Load and display the top logo (adjusted for PNG)
        logo_image = Image.open('top_logo.png')
        logo_photo = ImageTk.PhotoImage(logo_image)
        logo_label = tk.Label(root, image=logo_photo)
        logo_label.image = logo_photo  # Keep a reference to avoid garbage collection
        logo_label.grid(row=0, column=0, columnspan=2, pady=10)  # Centered using columnspan

        # Port selection
        tk.Label(root, text="Serial Port:").grid(row=1, column=0, padx=10, pady=10, sticky="e")  # Align label to the right
        self.port_entry = tk.Entry(root)
        self.port_entry.grid(row=1, column=1, padx=10, pady=10, sticky="w")  # Align entry to the left
        self.port_entry.insert(0, "COM3")  # Default value, adjust based on system

        # Connect Button
        self.connect_btn = tk.Button(root, text="Connect", command=self.connect_to_scale)
        self.connect_btn.grid(row=2, column=0, columnspan=2, pady=10)  # Centered using columnspan

        # Zero Scale Button
        self.zero_btn = tk.Button(root, text="Zero Scale", command=self.zero_scale, state=tk.DISABLED)
        self.zero_btn.grid(row=3, column=0, columnspan=2, pady=10)  # Centered using columnspan

        # First Measurement Button
        self.first_measure_btn = tk.Button(root, text="First Measurement", command=self.first_measurement, state=tk.DISABLED)
        self.first_measure_btn.grid(row=4, column=0, pady=10)
        self.first_measure_label = tk.Label(root, text="First: Not Measured")
        self.first_measure_label.grid(row=4, column=1, pady=10)

        # Second Measurement Button
        self.second_measure_btn = tk.Button(root, text="Second Measurement", command=self.second_measurement, state=tk.DISABLED)
        self.second_measure_btn.grid(row=5, column=0, pady=10)
        self.second_measure_label = tk.Label(root, text="Second: Not Measured")
        self.second_measure_label.grid(row=5, column=1, pady=10)

        # Save Excel Button
        self.save_btn = tk.Button(root, text="Save to Excel", command=self.save_to_excel, state=tk.DISABLED)
        self.save_btn.grid(row=6, column=0, columnspan=2, pady=10)  # Centered using columnspan

        # Credit label with hyperlink
        self.credit_label = tk.Label(root, text="Designed by Eng. Karim Arafa", fg="blue", cursor="hand2")
        self.credit_label.grid(row=7, column=0, columnspan=2, pady=10)  # Centered using columnspan
        self.credit_label.bind("<Button-1>", lambda e: self.open_link("https://www.linkedin.com/in/karimarafa/"))

        self.ser = None
        self.first_measurement_value = None
        self.second_measurement_value = None

    def connect_to_scale(self):
        port = self.port_entry.get()
        self.ser = init_serial(port)
        if self.ser:
            messagebox.showinfo("Connected", f"Connected to {port}")
            self.zero_btn.config(state=tk.NORMAL)
            self.first_measure_btn.config(state=tk.NORMAL)
            self.second_measure_btn.config(state=tk.NORMAL)
            self.save_btn.config(state=tk.NORMAL)
        else:
            messagebox.showerror("Error", "Failed to connect to the scale")

    def zero_scale(self):
        zero_scale(self.ser)

    def first_measurement(self):
        weight = read_weight(self.ser)
        if weight is not None:
            self.first_measurement_value = weight
            self.first_measure_label.config(text=f"First: {weight} kg")

    def second_measurement(self):
        weight = read_weight(self.ser)
        if weight is not None:
            self.second_measurement_value = weight
            self.second_measure_label.config(text=f"Second: {weight} kg")

    def save_to_excel(self):
        if self.first_measurement_value is not None and self.second_measurement_value is not None:
            filename = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
            if filename:
                log_to_excel(self.first_measurement_value, self.second_measurement_value, filename)
        else:
            messagebox.showwarning("Missing Measurements", "Please record both measurements before saving.")

    def open_link(self, url):
        import webbrowser
        webbrowser.open_new(url)

# Main function
if __name__ == "__main__":
    root = tk.Tk()
    app = ScaleApp(root)
    root.mainloop()
