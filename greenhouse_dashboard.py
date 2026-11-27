
import tkinter as tk
from tkinter import messagebox
import serial
import threading
import time
import os
from datetime import datetime
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# ==================================================
# SETTINGS
# ==================================================

SERIAL_PORT = "COM14"
BAUD_RATE = 9600
LOG_INTERVAL = 300  # 5 minutes

EXCEL_FILE = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "greenhouse_data.xlsx"
)

arduino = None
reading = False
reader_thread = None
last_save_time = None

data_lock = threading.Lock()
save_lock = threading.Lock()

# Sensor readings collected during each interval
samples = {
    "soil": [],
    "temperature": [],
    "humidity": [],
    "mq": [],
    "ldr_analog": [],
    "ldr_digital": [],
    "pump": [],
    "fan": []
}

latest_values = {
    "soil": "--",
    "temperature": "--",
    "humidity": "--",
    "mq": "--",
    "ldr_analog": "--",
    "ldr_digital": "--",
    "pump": "--",
    "fan": "--"
}


# ==================================================
# EXCEL FILE CREATION
# ==================================================

def setup_excel():

    if os.path.exists(EXCEL_FILE):
        wb = load_workbook(EXCEL_FILE)
    else:
        wb = Workbook()

    if "Readings" not in wb.sheetnames:

        ws = wb.create_sheet("Readings")

        ws.append([
            "Date & Time",
            "Average Soil Moisture",
            "Average Temperature (C)",
            "Average Humidity (%)",
            "Average MQ-135",
            "Average LDR Analog",
            "LDR Digital",
            "Pump Status",
            "Fan Status"
        ])

        for cell in ws[1]:
            cell.font = Font(
                bold=True,
                color="FFFFFF"
            )
            cell.fill = PatternFill(
                "solid",
                fgColor="218838"
            )
            cell.alignment = Alignment(
                horizontal="center"
            )

        ws.freeze_panes = "A2"

    if "Average Summary" not in wb.sheetnames:

        summary = wb.create_sheet("Average Summary")

        summary.append([
            "Sensor",
            "Overall Average"
        ])

        sensors = [
            ("Soil Moisture", "B"),
            ("Temperature (C)", "C"),
            ("Humidity (%)", "D"),
            ("MQ-135", "E"),
            ("LDR Analog", "F")
        ]

        for row, (name, column) in enumerate(
            sensors, start=2
        ):

            summary.cell(
                row=row,
                column=1,
                value=name
            )

            summary.cell(
                row=row,
                column=2,
                value=(
                    f'=IFERROR(AVERAGE('
                    f'Readings!{column}2:{column}1048576),'
                    f'"No data")'
                )
            )

        for cell in summary[1]:
            cell.font = Font(
                bold=True,
                color="FFFFFF"
            )
            cell.fill = PatternFill(
                "solid",
                fgColor="218838"
            )

        summary.column_dimensions["A"].width = 25
        summary.column_dimensions["B"].width = 22

    try:
        wb.calculation.fullCalcOnLoad = True
        wb.calculation.forceFullCalc = True
        wb.calculation.calcMode = "auto"
    except AttributeError:
        pass

    wb.save(EXCEL_FILE)
    wb.close()


# ==================================================
# SAVE SENSOR AVERAGES TO EXCEL
# ==================================================

def save_average_to_excel():

    with save_lock:

        # Take a snapshot without losing new incoming data.
        with data_lock:

            current_samples = {
                key: list(values)
                for key, values in samples.items()
            }

            for key in samples:
                samples[key].clear()

            current_values = latest_values.copy()

        numeric_keys = [
            "soil",
            "temperature",
            "humidity",
            "mq",
            "ldr_analog"
        ]

        # Do not save a row if no numeric readings exist.
        if not any(
            current_samples[key]
            for key in numeric_keys
        ):
            return

        averages = {}

        for key in numeric_keys:

            values = current_samples[key]

            averages[key] = (
                sum(values) / len(values)
                if values else None
            )

        def most_common_status(values, fallback):

            if not values:
                return fallback

            return max(
                set(values),
                key=values.count
            )

        try:

            setup_excel()

            wb = load_workbook(EXCEL_FILE)
            ws = wb["Readings"]

            ws.append([
                datetime.now().strftime(
                    "%Y-%m-%d %H:%M:%S"
                ),
                averages["soil"],
                averages["temperature"],
                averages["humidity"],
                averages["mq"],
                averages["ldr_analog"],
                current_values["ldr_digital"],
                most_common_status(
                    current_samples["pump"],
                    current_values["pump"]
                ),
                most_common_status(
                    current_samples["fan"],
                    current_values["fan"]
                )
            ])

            row_number = ws.max_row

            for column in range(2, 7):

                ws.cell(
                    row=row_number,
                    column=column
                ).number_format = "0.00"

            for column in range(1, 10):

                ws.column_dimensions[
                    get_column_letter(column)
                ].width = 23

            wb.save(EXCEL_FILE)
            wb.close()

            print(
                "Excel data saved:",
                datetime.now().strftime(
                    "%Y-%m-%d %H:%M:%S"
                )
            )

            print("File location:", EXCEL_FILE)

            root.after(
                0,
                status_var.set,
                "Sensor averages saved successfully!"
            )

        except Exception as error:

            # Restore samples if saving failed, so data is not lost.
            with data_lock:
                for key in samples:
                    samples[key] = (
                        current_samples[key] + samples[key]
                    )

            root.after(
                0,
                status_var.set,
                "Excel save error: " + str(error)
            )

            print("Excel save error:", error)


# ==================================================
# GUI WINDOW
# ==================================================

root = tk.Tk()

root.title("Smart Greenhouse Farming System")
root.geometry("850x700")
root.configure(bg="#eaf5e9")


# ==================================================
# GUI VARIABLES
# ==================================================

mode_var = tk.StringVar(value="Disconnected")
status_var = tk.StringVar(value="Connect Arduino to begin")
connection_var = tk.StringVar(value="Disconnected")

soil_var = tk.StringVar(value="--")
temp_var = tk.StringVar(value="--")
humidity_var = tk.StringVar(value="--")
mq_var = tk.StringVar(value="--")
ldr_analog_var = tk.StringVar(value="--")
ldr_digital_var = tk.StringVar(value="--")
pump_var = tk.StringVar(value="--")
fan_var = tk.StringVar(value="--")
load_var = tk.StringVar(value="--")


# ==================================================
# COMMAND FUNCTION
# ==================================================

def send_command(command):

    global arduino

    if arduino is None or not arduino.is_open:

        messagebox.showwarning(
            "Not connected",
            "Please connect Arduino first."
        )

        return

    try:

        arduino.write(
            (command + "\n").encode("utf-8")
        )

        arduino.flush()

        status_var.set(
            "Command sent: " + command
        )

    except serial.SerialException as error:

        status_var.set(
            "Serial error: " + str(error)
        )


# ==================================================
# NUMBER CONVERSION
# ==================================================

def get_number(value):

    try:
        return float(
            value.strip().split()[0]
        )

    except (ValueError, IndexError, AttributeError):
        return None


# ==================================================
# PROCESS ARDUINO DATA
# ==================================================

def process_line(line):

    line = line.strip()

    if not line:
        return

    if line.startswith("MODE:"):

        mode = line.split(":", 1)[1].strip()

        root.after(
            0,
            mode_var.set,
            mode
        )

    elif line.startswith("Status:"):

        status = line.split(":", 1)[1].strip()

        root.after(
            0,
            status_var.set,
            status
        )

    elif line.startswith("Soil Moisture Value:"):

        value = line.split(":", 1)[1].strip()
        number = get_number(value)

        if number is not None:

            with data_lock:
                samples["soil"].append(number)
                latest_values["soil"] = value

        root.after(0, soil_var.set, value)

    elif line.startswith("Temperature:"):

        value = line.split(":", 1)[1].strip()
        number = get_number(value)

        if number is not None:

            with data_lock:
                samples["temperature"].append(number)
                latest_values["temperature"] = value

        root.after(0, temp_var.set, value)

    elif line.startswith("Humidity:"):

        value = line.split(":", 1)[1].strip()
        number = get_number(value)

        if number is not None:

            with data_lock:
                samples["humidity"].append(number)
                latest_values["humidity"] = value

        root.after(0, humidity_var.set, value)

    elif line.startswith("MQ-135 raw value:"):

        value = line.split(":", 1)[1].strip()
        number = get_number(value)

        if number is not None:

            with data_lock:
                samples["mq"].append(number)
                latest_values["mq"] = value

        root.after(0, mq_var.set, value)

    elif line.startswith("LDR analog value:"):

        value = line.split(":", 1)[1].strip()
        number = get_number(value)

        if number is not None:

            with data_lock:
                samples["ldr_analog"].append(number)
                latest_values["ldr_analog"] = value

        root.after(0, ldr_analog_var.set, value)

    elif line.startswith("LDR digital output:"):

        value = line.split(":", 1)[1].strip()

        with data_lock:
            latest_values["ldr_digital"] = value

        root.after(
            0,
            ldr_digital_var.set,
            value
        )

    elif line.startswith("Pump Status:"):

        value = line.split(":", 1)[1].strip()

        with data_lock:
            samples["pump"].append(value)
            latest_values["pump"] = value

        root.after(0, pump_var.set, value)

    elif line.startswith("Fan Status:"):

        value = line.split(":", 1)[1].strip()

        with data_lock:
            samples["fan"].append(value)
            latest_values["fan"] = value

        root.after(0, fan_var.set, value)
        root.after(0, load_var.set, value)


# ==================================================
# SERIAL READING THREAD
# ==================================================

def read_serial():

    global arduino, reading, last_save_time

    while reading:

        try:

            if arduino and arduino.is_open:

                if arduino.in_waiting:

                    line = arduino.readline().decode(
                        "utf-8",
                        errors="replace"
                    ).strip()

                    if line:
                        process_line(line)

                else:
                    time.sleep(0.02)

                # Automatic 5-minute saving
                if (
                    time.monotonic() - last_save_time
                    >= LOG_INTERVAL
                ):

                    save_average_to_excel()
                    last_save_time = time.monotonic()

            else:
                time.sleep(0.1)

        except (
            serial.SerialException,
            OSError
        ) as error:

            root.after(
                0,
                status_var.set,
                "Serial reading error: " + str(error)
            )

            break


# ==================================================
# CONNECT ARDUINO
# ==================================================

def connect_arduino():

    global arduino, reading, last_save_time, reader_thread

    if arduino is not None and arduino.is_open:

        messagebox.showinfo(
            "Arduino",
            "Arduino is already connected."
        )

        return

    try:

        setup_excel()

        arduino = serial.Serial(
            port=SERIAL_PORT,
            baudrate=BAUD_RATE,
            timeout=1
        )

        # Arduino may reset when serial connects.
        time.sleep(2)

        with data_lock:

            for key in samples:
                samples[key].clear()

        last_save_time = time.monotonic()
        reading = True

        connection_var.set(
            "Connected - " + SERIAL_PORT
        )

        status_var.set(
            "Arduino connected. Excel logging started."
        )

        reader_thread = threading.Thread(
            target=read_serial,
            daemon=True
        )

        reader_thread.start()

        print("Excel file location:", EXCEL_FILE)

    except Exception as error:

        arduino = None
        connection_var.set("Disconnected")

        messagebox.showerror(
            "Connection Error",
            f"Could not connect to {SERIAL_PORT}.\n\n{error}"
        )


# ==================================================
# DISCONNECT AND SAVE REMAINING DATA
# ==================================================

def disconnect_arduino():

    global arduino, reading, reader_thread

    if not reading and arduino is None:
        return

    reading = False

    # Wait for the serial thread to stop before saving.
    if (
        reader_thread is not None
        and reader_thread.is_alive()
        and threading.current_thread() != reader_thread
    ):

        reader_thread.join(timeout=2)

    # Save any readings remaining in the current interval.
    save_average_to_excel()

    if arduino is not None:

        try:

            if arduino.is_open:
                arduino.close()

        except serial.SerialException:
            pass

    arduino = None
    reader_thread = None

    connection_var.set("Disconnected")

    status_var.set(
        "Arduino disconnected. Remaining data saved."
    )


# ==================================================
# MODE CONTROLS
# ==================================================

def switch_to_auto():
    send_command("AUTO")


def switch_to_manual():
    send_command("MANUAL")


# ==================================================
# SHARED PUMP AND FAN CONTROLS
# ==================================================

def load_on():
    send_command("PUMP_ON")


def load_off():
    send_command("PUMP_OFF")


# ==================================================
# GUI DESIGN
# ==================================================

title = tk.Label(
    root,
    text="SMART GREENHOUSE FARMING SYSTEM",
    font=("Arial", 20, "bold"),
    bg="#eaf5e9",
    fg="#176b36"
)

title.pack(pady=15)


# CONNECTION FRAME

connection_frame = tk.Frame(
    root,
    bg="#eaf5e9"
)

connection_frame.pack(pady=5)

tk.Label(
    connection_frame,
    text="Connection:",
    font=("Arial", 11, "bold"),
    bg="#eaf5e9"
).pack(side="left", padx=5)

tk.Label(
    connection_frame,
    textvariable=connection_var,
    font=("Arial", 11),
    bg="#eaf5e9",
    fg="#176b36"
).pack(side="left", padx=5)

tk.Button(
    connection_frame,
    text="Connect Arduino",
    command=connect_arduino,
    bg="#218838",
    fg="white",
    width=16
).pack(side="left", padx=8)

tk.Button(
    connection_frame,
    text="Disconnect",
    command=disconnect_arduino,
    bg="#b02a37",
    fg="white",
    width=12
).pack(side="left", padx=5)


# SENSOR DISPLAY

sensor_frame = tk.LabelFrame(
    root,
    text="Live Sensor Readings",
    font=("Arial", 13, "bold"),
    bg="white",
    padx=20,
    pady=15
)

sensor_frame.pack(
    fill="x",
    padx=25,
    pady=15
)


def add_sensor_row(parent, row, label, variable):

    tk.Label(
        parent,
        text=label,
        font=("Arial", 11),
        bg="white",
        width=25,
        anchor="w"
    ).grid(
        row=row,
        column=0,
        sticky="w",
        pady=6
    )

    tk.Label(
        parent,
        textvariable=variable,
        font=("Arial", 11, "bold"),
        bg="white",
        fg="#176b36",
        width=20,
        anchor="w"
    ).grid(
        row=row,
        column=1,
        sticky="w",
        pady=6
    )


add_sensor_row(sensor_frame, 0, "Soil Moisture Value", soil_var)
add_sensor_row(sensor_frame, 1, "Temperature", temp_var)
add_sensor_row(sensor_frame, 2, "Humidity", humidity_var)
add_sensor_row(sensor_frame, 3, "MQ-135 Raw Value", mq_var)
add_sensor_row(sensor_frame, 4, "LDR Analog Value", ldr_analog_var)
add_sensor_row(sensor_frame, 5, "LDR Digital Output", ldr_digital_var)


# MODE CONTROL

mode_frame = tk.LabelFrame(
    root,
    text="Operating Mode",
    font=("Arial", 13, "bold"),
    bg="white",
    padx=15,
    pady=12
)

mode_frame.pack(
    fill="x",
    padx=25,
    pady=8
)

tk.Label(
    mode_frame,
    text="Current Mode:",
    font=("Arial", 11),
    bg="white"
).pack(side="left", padx=10)

tk.Label(
    mode_frame,
    textvariable=mode_var,
    font=("Arial", 12, "bold"),
    fg="#176b36",
    bg="white"
).pack(side="left", padx=10)

tk.Button(
    mode_frame,
    text="AUTO MODE",
    command=switch_to_auto,
    bg="#218838",
    fg="white",
    width=15
).pack(side="left", padx=15)

tk.Button(
    mode_frame,
    text="MANUAL MODE",
    command=switch_to_manual,
    bg="#d39e00",
    fg="white",
    width=15
).pack(side="left", padx=5)


# LOAD CONTROL

control_frame = tk.LabelFrame(
    root,
    text="Pump + Fan Control (Shared Relay)",
    font=("Arial", 13, "bold"),
    bg="white",
    padx=15,
    pady=12
)

control_frame.pack(
    fill="x",
    padx=25,
    pady=8
)

tk.Label(
    control_frame,
    text="Shared Load Status:",
    font=("Arial", 11),
    bg="white"
).pack(side="left", padx=8)

tk.Label(
    control_frame,
    textvariable=load_var,
    font=("Arial", 12, "bold"),
    fg="#176b36",
    bg="white"
).pack(side="left", padx=8)

tk.Button(
    control_frame,
    text="LOAD ON",
    command=load_on,
    bg="#218838",
    fg="white",
    width=13
).pack(side="left", padx=12)

tk.Button(
    control_frame,
    text="LOAD OFF",
    command=load_off,
    bg="#b02a37",
    fg="white",
    width=13
).pack(side="left", padx=5)


# EXCEL INFORMATION

excel_frame = tk.LabelFrame(
    root,
    text="Excel Data Logging",
    font=("Arial", 12, "bold"),
    bg="white",
    padx=15,
    pady=10
)

excel_frame.pack(
    fill="x",
    padx=25,
    pady=8
)

tk.Label(
    excel_frame,
    text="Automatic sensor averages saved every 5 minutes",
    font=("Arial", 10),
    bg="white",
    fg="#176b36"
).pack(anchor="w")

tk.Label(
    excel_frame,
    text="File: greenhouse_data.xlsx",
    font=("Arial", 10),
    bg="white"
).pack(anchor="w")


# STATUS DISPLAY

status_frame = tk.Frame(
    root,
    bg="#eaf5e9"
)

status_frame.pack(
    fill="x",
    padx=25,
    pady=12
)

tk.Label(
    status_frame,
    text="System Status:",
    font=("Arial", 10, "bold"),
    bg="#eaf5e9"
).pack(side="left")

tk.Label(
    status_frame,
    textvariable=status_var,
    font=("Arial", 10),
    bg="#eaf5e9",
    fg="#176b36",
    wraplength=600
).pack(side="left", padx=8)


# ==================================================
# CLOSE WINDOW
# ==================================================

def on_closing():

    disconnect_arduino()
    root.destroy()


root.protocol(
    "WM_DELETE_WINDOW",
    on_closing
)

# Create Excel file when the program starts.
try:
    setup_excel()
    print("Excel file created/ready:")
    print(EXCEL_FILE)

except Exception as error:
    messagebox.showerror(
        "Excel Error",
        str(error)
    )

root.mainloop()
