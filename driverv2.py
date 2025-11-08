import wmi
import json
import os
import tkinter as tk
from tkinter import ttk
import webbrowser


def list_drivers(save_to_file=False):
    try:
        computer = wmi.WMI()  # eskiden _wmi.WMI()
        drivers = []
        for device in computer.Win32_PnPEntity():
            if device.Name and device.DeviceID:
                status = "YÜKLÜ" if getattr(device, "Status", "") == "OK" else "YÜKLÜ DEĞİL"
                driver_info = {
                    "Cihaz": device.Name,
                    "Device ID": device.DeviceID,
                    "Üretici": getattr(device, "Manufacturer", "Bilinmiyor"),
                    "Durum": status,
                    "Sürücü Linki": "Sürücüyü Araştır"
                }
                drivers.append(driver_info)

        if save_to_file:
            file_path = os.path.join(os.getcwd(), "driver_info.json")
            with open(file_path, "w", encoding="utf-8") as file:
                json.dump(drivers, file, ensure_ascii=False, indent=4)
            print(f"Sürücü bilgiler {file_path} dosyasına kaydedildi!")

        print(f"Toplam sürücü bulundu: {len(drivers)}")
        return drivers

    except Exception as e:
        print(f"Hata: {str(e)}")
        return []


def display_Drivers_in_table():
    drivers = list_drivers()
    if not drivers:
        print("Hiç sürücü bulunamadı!")
        return

    root = tk.Tk()
    root.title("Bilgisayar Sürücüleri")
    root.geometry("1000x500")

    # Arama çubuğu
    search_frame = tk.Frame(root)
    search_frame.pack(fill="x", padx=10, pady=5)

    search_label = tk.Label(search_frame, text="Ara:")
    search_label.pack(side="left", padx=5)

    search_entry = tk.Entry(search_frame)
    search_entry.pack(side="left", fill="x", expand=True, padx=5)

    # Sütunları oluştur
    columns = ("Cihaz", "Device ID", "Üretici", "Durum", "Sürücü Linki")
    tree = ttk.Treeview(root, columns=columns, show="headings")

    for col in columns:
        tree.heading(col, text=col)
        tree.column(col, width=200)

    # Tabloya veri dolduran fonksiyon
    def populate_treeview(data):
        for item in tree.get_children():
            tree.delete(item)
        for driver in data:
            item_id = tree.insert("", tk.END, values=(
                driver.get("Cihaz", "Bilinmiyor"),
                driver.get("Device ID", "Bilinmiyor"),
                driver.get("Üretici", "Bilinmiyor"),
                driver.get("Durum", "Bilinmiyor"),
                "Sürücüyü Araştır"
            ))
            # Duruma göre renk ver
            if driver.get("Durum") == "YÜKLÜ":
                tree.item(item_id, tags=("yuklu",))
            else:
                tree.item(item_id, tags=("yuklu_degil",))

    # Durum etiketleri
    tree.tag_configure("yuklu", background="lightgreen")
    tree.tag_configure("yuklu_degil", background="lightcoral")

    # Başlangıçta tüm sürücüleri doldur
    populate_treeview(drivers)
    tree.pack(expand=True, fill="both")

    # Arama fonksiyonu
    def search_drivers(event=None):
        query = search_entry.get().lower()
        filtered_drivers = [
            driver for driver in drivers
            if (driver.get("Cihaz") and query in driver["Cihaz"].lower()) or
               (driver.get("Device ID") and query in driver["Device ID"].lower()) or
               (driver.get("Üretici") and query in driver["Üretici"].lower()) or
               (driver.get("Durum") and query in driver["Durum"].lower())
        ]
        populate_treeview(filtered_drivers)

    search_entry.bind("<KeyRelease>", search_drivers)

    # Çift tıklamada sürücü arama
    def on_treview_double_click(event):
        selection = tree.selection()
        if selection:
            item = selection[0]
            device_name = tree.item(item, "values")[0]
            driver_link = f"https://www.google.com/search?q={device_name} driver"
            webbrowser.open(driver_link)

    tree.bind("<Double-1>", on_treview_double_click)

    root.mainloop()

if __name__ == "__main__":
    display_Drivers_in_table()
