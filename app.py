import wmi
import psutil
import platform
import pandas as pd
import customtkinter as ctk
from datetime import datetime
from tkinter import filedialog, messagebox

# Inicializa WMI
c = wmi.WMI()

def get_full_system_info():
    info = {}

    # Información básica del sistema
    info.update({
        "Processor": platform.processor(),
        "System": platform.system(),
        "Release": platform.release(),
        "Version": platform.version(),
        "Machine": platform.machine(),
        "Architecture": platform.architecture()[0],
        "CPU Cores (Physical)": psutil.cpu_count(logical=False),
        "CPU Cores (Logical)": psutil.cpu_count(logical=True),
        "Memory": f"{round(psutil.virtual_memory().total / (1024.0 ** 3))} GB",
        "Available Memory": f"{round(psutil.virtual_memory().available / (1024.0 ** 3))} GB",
        "Used Memory": f"{round(psutil.virtual_memory().used / (1024.0 ** 3))} GB",
        "Disk": f"{round(psutil.disk_usage('/').total / (1024.0 ** 3))} GB",
        "Disk Used": f"{round(psutil.disk_usage('/').used / (1024.0 ** 3))} GB",
        "Disk Free": f"{round(psutil.disk_usage('/').free / (1024.0 ** 3))} GB",
        "Boot Time": datetime.fromtimestamp(psutil.boot_time()).strftime("%Y-%m-%d %H:%M:%S"),
    })

    # Información del CPU
    for processor in c.Win32_Processor():
        info.update({
            "CPU Name": processor.Name,
            "CPU Manufacturer": processor.Manufacturer,
            "CPU Max Clock Speed": f"{processor.MaxClockSpeed} MHz",
            "CPU Number of Cores": processor.NumberOfCores,
            "CPU Threads": processor.ThreadCount,
            "CPU Socket": processor.SocketDesignation,
        })

    # Información de la memoria RAM
    for i, mem in enumerate(c.Win32_PhysicalMemory()):
        info.update({
            f"RAM Slot {i + 1} Capacity": f"{round(int(mem.Capacity) / (1024.0 ** 3))} GB",
            f"RAM Slot {i + 1} Manufacturer": mem.Manufacturer,
            f"RAM Slot {i + 1} Speed": f"{mem.Speed} MHz",
            f"RAM Slot {i + 1} Part Number": mem.PartNumber.strip(),
        })

    # Información de la placa base (motherboard)
    motherboard = c.Win32_BaseBoard()[0]
    info.update({
        "Motherboard Manufacturer": motherboard.Manufacturer,
        "Motherboard Model": motherboard.Product,
        "Motherboard Serial Number": motherboard.SerialNumber,
    })

    # Información de las unidades de disco
    for i, disk in enumerate(c.Win32_DiskDrive()):
        info.update({
            f"Disk {i + 1} Model": disk.Model,
            f"Disk {i + 1} Serial Number": disk.SerialNumber.strip(),
            f"Disk {i + 1} Size": f"{round(int(disk.Size) / (1024.0 ** 3))} GB",
            f"Disk {i + 1} Interface": disk.InterfaceType,
        })

    # Información de la tarjeta gráfica
    for i, gpu in enumerate(c.Win32_VideoController()):
        info.update({
            f"GPU {i + 1} Name": gpu.Name,
            f"GPU {i + 1} Driver Version": gpu.DriverVersion,
            f"GPU {i + 1} Status": gpu.Status,
            f"GPU {i + 1} Video Mode": gpu.VideoModeDescription,
            f"GPU {i + 1} Adapter RAM": f"{round(int(gpu.AdapterRAM) / (1024.0 ** 3))} GB" if gpu.AdapterRAM else "N/A",
        })

    # Información de los adaptadores de red
    for i, net_adapter in enumerate(c.Win32_NetworkAdapter()):
        info.update({
            f"Network Adapter {i + 1} Name": net_adapter.Name,
            f"Network Adapter {i + 1} MAC Address": net_adapter.MACAddress if net_adapter.MACAddress else "N/A",
            f"Network Adapter {i + 1} Speed": f"{net_adapter.Speed} bps" if net_adapter.Speed else "N/A",
            f"Network Adapter {i + 1} Status": net_adapter.Status,
        })

    # Información sobre sensores de temperatura (si están disponibles)
    try:
        for i, temp in enumerate(c.MSAcpi_ThermalZoneTemperature()):
            info.update({
                f"Temperature Sensor {i + 1}": f"{temp.CurrentTemperature / 10 - 273.15:.2f} °C",
            })
    except Exception as e:
        info.update({"Temperature Sensors": "Not available"})

    return info

def show_system_info():
    info = get_full_system_info()
    info_text = "\n".join(f"{key}: {value}" for key, value in info.items())
    output_textbox.delete("1.0", ctk.END)  # Limpiar el cuadro de texto
    output_textbox.insert(ctk.INSERT, info_text)

def save_to_excel():
    info = get_full_system_info()
    df = pd.DataFrame.from_dict(info, orient='index', columns=["Value"])
    df.reset_index(inplace=True)
    df.columns = ["Component", "Value"]

    # Guardar en archivo Excel
    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        df.to_excel(file_path, index=False)
        messagebox.showinfo("Success", f"Information saved to {file_path}")

# Configuración de la ventana principal
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("dark-blue")

root = ctk.CTk()
root.title("Advanced System Information")
root.geometry("700x700")

# Crear un frame para organizar el contenido
frame = ctk.CTkFrame(root)  
frame.pack(pady=20, padx=20, fill="both", expand=True)

# Crear widgets dentro del frame
title_label = ctk.CTkLabel(frame, text="System Information", font=("Arial", 24))
title_label.pack(pady=10)

# Barra de progreso (se puede usar para mostrar mientras se recoge la información)
progress_bar = ctk.CTkProgressBar(frame)
progress_bar.pack(pady=10)
progress_bar.set(0.0)  # Valor inicial

output_textbox = ctk.CTkTextbox(frame, wrap=ctk.WORD, width=600, height=400, font=("Arial", 12))
output_textbox.pack(pady=10)

info_button = ctk.CTkButton(frame, text="Get System Info", command=show_system_info, width=200, height=40)
info_button.pack(pady=10)

save_button = ctk.CTkButton(frame, text="Save to Excel", command=save_to_excel, width=200, height=40)
save_button.pack(pady=10)

# Ejecutar la aplicación
root.mainloop()
