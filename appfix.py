import pandas as pd
import customtkinter as ctk
from tkinter import filedialog, messagebox
from tkinter import ttk
import pywhatkit as pwk
import time
from PIL import Image, ImageTk
import threading
from flask import Flask, render_template_string, redirect
import os

# === GLOBAL VARIABEL ===
df = pd.DataFrame()
file_path = ''
stop_broadcast_flag = False

# === FLASK APP UNTUK STOP BROADCAST VIA BROWSER ===
app = Flask(__name__)

@app.route('/')
def index():
    return render_template_string('''
        <h2>Kontrol Broadcast WhatsApp</h2>
        <form action="/stop" method="post">
            <button type="submit" style="font-size:20px; padding:10px 20px; background-color:#dc3545; color:white; border:none; border-radius:5px;">Stop Broadcast Sekarang</button>
        </form>
    ''')

@app.route('/stop', methods=['POST'])
def stop():
    global stop_broadcast_flag
    stop_broadcast_flag = True
    return redirect('/')

def run_flask():
    app.run(port=5000, debug=False, use_reloader=False)

# === FUNGSI STOP DARI GUI ===
def stop_broadcast_gui():
    global stop_broadcast_flag
    stop_broadcast_flag = True
    messagebox.showinfo("Dihentikan", "Broadcast dihentikan lewat tombol di aplikasi.")

# === FUNGSI-FUNGSI UTAMA ===
def load_excel():
    global df, file_path
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        df = pd.read_excel(file_path)
        show_data()

def show_data():
    for i in tree.get_children():
        tree.delete(i)
    for index, row in df.iterrows():
        tree.insert('', 'end', values=list(row))

def save_data():
    if file_path:
        df.to_excel(file_path, index=False)
        messagebox.showinfo("Sukses", "Data berhasil disimpan!")
    else:
        messagebox.showwarning("Peringatan", "Tidak ada file yang dimuat!")

def edit_row():
    try:
        selected_item = tree.selection()[0]
        values = tree.item(selected_item, 'values')
        index = tree.index(selected_item)

        edit_window = ctk.CTkToplevel(root)
        edit_window.title("Edit Data")
        edit_window.geometry("500x400")
        edit_window.grab_set()

        ctk.CTkLabel(edit_window, text="Nama:").grid(row=0, column=0, padx=10, pady=10, sticky='w')
        nama_entry = ctk.CTkEntry(edit_window, width=300)
        nama_entry.grid(row=0, column=1, padx=10, pady=10)
        nama_entry.insert(0, values[0])

        ctk.CTkLabel(edit_window, text="Nomor WhatsApp:").grid(row=1, column=0, padx=10, pady=10, sticky='w')
        nomor_entry = ctk.CTkEntry(edit_window, width=300)
        nomor_entry.grid(row=1, column=1, padx=10, pady=10)
        nomor_entry.insert(0, values[1])

        ctk.CTkLabel(edit_window, text="Pesan:").grid(row=2, column=0, padx=10, pady=10, sticky='w')
        pesan_entry = ctk.CTkTextbox(edit_window, width=300, height=100)
        pesan_entry.grid(row=2, column=1, padx=10, pady=10)
        pesan_entry.insert("0.0", values[2])

        def save_changes():
            nama = nama_entry.get().strip()
            nomor = nomor_entry.get().strip()
            pesan = pesan_entry.get("0.0", "end").strip()
            if not nama or not nomor or not pesan:
                messagebox.showwarning("Peringatan", "Semua field harus diisi!")
                return
            df.at[index, 'Nama'] = nama
            df.at[index, 'Nomor WhatsApp'] = nomor
            df.at[index, 'Pesan'] = pesan
            tree.item(selected_item, values=(nama, nomor, pesan))
            edit_window.destroy()
            messagebox.showinfo("Sukses", "Data berhasil diedit!")

        ctk.CTkButton(edit_window, text="Simpan", command=save_changes, fg_color="#2E8B57").grid(row=3, columnspan=2, pady=20)

    except IndexError:
        messagebox.showwarning("Peringatan", "Pilih baris yang ingin diedit!")

def add_row():
    add_window = ctk.CTkToplevel(root)
    add_window.title("Tambah Data Baru")
    add_window.geometry("500x400")
    add_window.grab_set()

    ctk.CTkLabel(add_window, text="Nama:").grid(row=0, column=0, padx=10, pady=10, sticky='w')
    nama_entry = ctk.CTkEntry(add_window, width=300)
    nama_entry.grid(row=0, column=1, padx=10, pady=10)

    ctk.CTkLabel(add_window, text="Nomor WhatsApp:").grid(row=1, column=0, padx=10, pady=10, sticky='w')
    nomor_entry = ctk.CTkEntry(add_window, width=300)
    nomor_entry.grid(row=1, column=1, padx=10, pady=10)

    ctk.CTkLabel(add_window, text="Pesan:").grid(row=2, column=0, padx=10, pady=10, sticky='w')
    pesan_entry = ctk.CTkTextbox(add_window, width=300, height=100)
    pesan_entry.grid(row=2, column=1, padx=10, pady=10)

    def save_new_data():
        nama = nama_entry.get().strip()
        nomor = nomor_entry.get().strip()
        pesan = pesan_entry.get("0.0", "end").strip()
        if not nama or not nomor or not pesan:
            messagebox.showwarning("Peringatan", "Semua field harus diisi!")
            return
        df.loc[len(df)] = {'Nama': nama, 'Nomor WhatsApp': nomor, 'Pesan': pesan}
        show_data()
        add_window.destroy()
        messagebox.showinfo("Sukses", "Data berhasil ditambahkan!")

    ctk.CTkButton(add_window, text="Simpan", command=save_new_data, fg_color="#2E8B57").grid(row=3, columnspan=2, pady=20)

def delete_row():
    try:
        selected_item = tree.selection()[0]
        index = tree.index(selected_item)
        confirm = messagebox.askyesno("Konfirmasi", "Yakin ingin hapus data ini?")
        if confirm:
            df.drop(df.index[index], inplace=True)
            show_data()
            messagebox.showinfo("Sukses", "Data berhasil dihapus!")
    except IndexError:
        messagebox.showwarning("Peringatan", "Pilih data yang ingin dihapus!")

# === FUNGSI BROADCAST ===
def broadcast_messages():
    global stop_broadcast_flag
    for index, row in df.iterrows():
        if stop_broadcast_flag:
            print("Broadcast dihentikan.")
            messagebox.showinfo("Berhenti", "Broadcast dihentikan.")
            return
        nomor = str(row['Nomor WhatsApp']).strip()
        if not nomor.startswith('+'):
            nomor = '+' + nomor
        pesan = str(row['Pesan']).strip()
        if not nomor or not pesan:
            print(f"Data kosong pada baris {index+1}. Lewat...")
            continue
        try:
            print(f"Mengirim ke {nomor}")
            pwk.sendwhatmsg_instantly(nomor, pesan, 20, True, 10)
            time.sleep(10)
        except Exception as e:
            print(f"Error saat kirim ke {nomor}: {e}")
    messagebox.showinfo("Selesai", "Broadcast selesai!")

def start_broadcast():
    global stop_broadcast_flag
    stop_broadcast_flag = False
    if df.empty:
        messagebox.showwarning("Peringatan", "Tidak ada data untuk dikirim!")
        return
    confirm = messagebox.askyesno("Konfirmasi", "Yakin ingin mulai broadcast?")
    if confirm:
        threading.Thread(target=broadcast_messages).start()

# === GUI CUSTOMTKINTER (PERBAIKAN FINAL) ===
ctk.set_appearance_mode("Light")  # Bisa: "Light", "Dark"
ctk.set_default_color_theme("blue")  # Bisa: "blue", "green", "dark-blue"

root = ctk.CTk()
root.title("Broadcast WhatsApp Fallqyun.dev")
root.state('zoomed')  # Full screen

# === LOGO DAN JUDUL ===
top_frame = ctk.CTkFrame(root)
top_frame.pack(pady=20, fill='x')

try:
    # Ganti path jika kamu simpan logo di folder proyek, misal: assets/fallqyun_dev.png
    logo_path = r'C:\Users\62851\Documents\Python Ujicoba\BCWA\fallqyun_dev.png'
    logo_image = Image.open(logo_path)

    # Pertahankan rasio aspek
    width = 250
    aspect_ratio = logo_image.height / logo_image.width
    height = int(width * aspect_ratio)
    
    logo_image = logo_image.resize((width, height), Image.LANCZOS)
    logo_photo = ImageTk.PhotoImage(logo_image)

    logo_label = ctk.CTkLabel(top_frame, image=logo_photo, text="")
    logo_label.image = logo_photo  # Simpan referensi
    logo_label.pack(pady=10)
except Exception as e:
    print(f"Gagal memuat logo: {e}")
    ctk.CTkLabel(top_frame, text="fallqyun.dev", font=("Helvetica", 24, "bold")).pack(pady=10)

company_label = ctk.CTkLabel(top_frame, text="fallqyun.dev", font=("Helvetica", 16, "bold"))
company_label.pack(pady=5)

# === TABLE FRAME ===
main_frame = ctk.CTkFrame(root)
main_frame.pack(pady=10, padx=20, expand=True, fill='both')

tree = ttk.Treeview(main_frame, columns=("Nama", "Nomor WhatsApp", "Pesan"), show='headings')
tree.heading("Nama", text="Nama")
tree.heading("Nomor WhatsApp", text="Nomor WhatsApp")
tree.heading("Pesan", text="Pesan")
tree.column("Nama", anchor='center', width=150)
tree.column("Nomor WhatsApp", anchor='center', width=150)
tree.column("Pesan", anchor='w', width=600)
tree.pack(pady=20, expand=True, fill='both')

# === BUTTON FRAME (RATA TENGAH) ===
button_frame = ctk.CTkFrame(root)
button_frame.pack(fill='x', pady=20)

buttons = [
    ("Muat Data Excel", load_excel),
    ("Tambah Data", add_row),
    ("Edit Data", edit_row),
    ("Hapus Data", delete_row),
    ("Simpan Perubahan", save_data),
    ("Mulai Broadcast", start_broadcast),
    ("Stop Broadcast", stop_broadcast_gui)
]

# Konfigurasi kolom agar seimbang
for i in range(len(buttons)):
    button_frame.grid_columnconfigure(i, weight=1)

for i, (text, command) in enumerate(buttons):
    if text == "Stop Broadcast":
        btn = ctk.CTkButton(button_frame, text=text, command=command, fg_color="#DC3545", hover_color="#C82333", height=40)
    else:
        btn = ctk.CTkButton(button_frame, text=text, command=command, height=40)
    btn.grid(row=0, column=i, padx=5, pady=10, sticky="ew")

# === MULAI FLASK SERVER DALAM THREAD LAIN ===
threading.Thread(target=run_flask, daemon=True).start()

# === MULAI APLIKASI ===
root.mainloop()