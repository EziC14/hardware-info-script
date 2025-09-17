import subprocess
import csv
import os
import re

def run_cmd(cmd, use_powershell=False):
    if use_powershell:
        cmd = f'powershell -Command "{cmd}"'
    result = subprocess.run(cmd, capture_output=True, text=True, shell=True)
    return result.stdout.strip()

def bytes_to_gb(size_str):
    """Convierte bytes a GB, si es número válido"""
    try:
        size_int = int(size_str)
        gb = round(size_int / (1024 ** 3), 2)  # 2 decimales
        return f"{gb} GB"
    except:
        return size_str.strip()

# ------------------- COMANDOS BÁSICOS -------------------
marca = run_cmd("wmic computersystem get manufacturer").splitlines()[-1].strip()
modelo = run_cmd("wmic computersystem get model").splitlines()[-1].strip()
serie = run_cmd("wmic bios get serialnumber").splitlines()[-1].strip()
cpu = run_cmd("wmic cpu get name").splitlines()[-1].strip()
placa = run_cmd("wmic baseboard get product").splitlines()[-1].strip()
gpu = run_cmd("wmic path win32_VideoController get name").splitlines()[-1].strip()
windows = run_cmd("wmic os get caption").splitlines()[-1].strip()

# ------------------- RAM DETALLADA -------------------
ram_raw = run_cmd(
    "Get-CimInstance -ClassName Win32_PhysicalMemory | "
    "Select-Object Manufacturer, Speed, Capacity, SMBIOSMemoryType | Format-Table -HideTableHeaders",
    use_powershell=True
)

dict_type_ram = {
    "20": "DDR",
    "21": "DDR2",
    "22": "DDR2 FB-DIMM",
    "24": "DDR3",
    "26": "DDR4",
    "27": "LPDDR",
    "28": "LPDDR2",
    "29": "LPDDR3",
    "30": "LPDDR4",
    "31": "Logical non-volatile device"
}

ram_list = []
for line in ram_raw.splitlines():
    parts = re.split(r"\s+", line.strip())
    if len(parts) >= 4:
        manufacturer = parts[0]
        speed = parts[1] + "MHz"
        capacity = bytes_to_gb(parts[2])
        mem_code = parts[3]
        mem_type = dict_type_ram.get(mem_code, f"Desconocido({mem_code})")  
        ram_list.append(f"{manufacturer} {speed} {capacity} {mem_type}")

ram_str = " | ".join(ram_list)

# ------------------- DISCO DETALLADO -------------------
disco_raw = run_cmd(
    "Get-PhysicalDisk | Select-Object MediaType,Model,Size | Format-Table -HideTableHeaders",
    use_powershell=True
)

disco_list = []
for line in disco_raw.splitlines():
    parts = re.split(r"\s+", line.strip())
    if len(parts) >= 3:
        media_type = parts[0]
        model = " ".join(parts[1:-1])
        size = bytes_to_gb(parts[-1])
        disco_list.append(f"{media_type} {model} {size}")

disco_str = " | ".join(disco_list)

# ------------------- DATOS DEL USUARIO -------------------
nombre = input("Nombre del Usuario: ")
apellido = input("Apellido: ")
usurio = input("Usuario: ")
area = input("Área: ")
cargo = input("Cargo: ")
office = input("Microsoft Office: ")

# ------------------- DICCIONARIO FINAL -------------------
info = {
    "NOMBRE": nombre,
    "APELLIDO": apellido,
    "USUARIO": usurio,
    "AREA": area,
    "CARGO": cargo,
    "MICROSOFT OFFICE": office,
    "Marca": marca,
    "Modelo": modelo,
    "Serie": serie,
    "CPU": cpu,
    "Placa": placa,
    "RAM": ram_str,
    "Disco": disco_str,
    "GPU": gpu,
    "Windows": windows
}

# ------------------- GUARDAR EN CSV -------------------
archivo_csv = r"\\192.168.1.20\Aplicaciones\Inventario\inventario_pcs.csv"
existe = os.path.isfile(archivo_csv)

with open(archivo_csv, mode='a', newline='', encoding='utf-8') as file:
    writer = csv.DictWriter(file, fieldnames=info.keys())
    if not existe:
        writer.writeheader()
    writer.writerow(info)

print(f"\n✅ Información guardada correctamente en '{archivo_csv}'")
