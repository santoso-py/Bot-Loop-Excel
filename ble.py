import pandas as pd
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from tkinter import ttk


def process_excel (file_path, output_path, loop_column):
	# Membaca file excel
	df = pd.read_excel (file_path)
	
	rows = []
	
	for index, row in df.iterrows ():
		loop_count = int (row [loop_column])  # Mengambil nilai loop dari kolom yang dipilih
		for _ in range (loop_count):  # Loop sebanyak nilai di kolom tersebut
			rows.append (row)
	
	new_df = pd.DataFrame (rows, columns = df.columns)
	
	# Menyimpan hasil ke file baru
	new_df.to_excel (output_path, index = False)
	
	messagebox.showinfo ("Selesai", "File Selesai di looping dan disimpen")


def select_file ():
	# Pilih file Excel
	file_path = filedialog.askopenfilename (
		title = "Pilih File Excel",
		filetypes = [("Excel files", "*.xlsx *.xls")]
	)
	file_entry.delete (0, tk.END)
	file_entry.insert (0, file_path)
	
	# Setelah file dipilih, load kolom dari file tersebut
	if file_path:
		df = pd.read_excel (file_path)
		columns = df.columns.tolist ()  # Ambil nama kolom
		column_combobox ['values'] = columns  # Masukkan ke combobox
		column_combobox.set ('')  # Reset combobox


def save_file ():
	# Pilih tempat untuk menyimpan file hasil proses
	output_path = filedialog.asksaveasfilename (
		defaultextension = ".xlsx",
		filetypes = [("Excel files", "*.xlsx *.xls")],
		title = "Simpan File Sebagai"
	)
	output_entry.delete (0, tk.END)
	output_entry.insert (0, output_path)


def start_processing ():
	# Ambil nilai input
	file_path = file_entry.get ()
	output_path = output_entry.get ()
	loop_column = column_combobox.get ()
	
	if not file_path or not output_path or not loop_column:
		messagebox.showerror ("Maaf", "Pilih file atau kolom terlebih dulu")
		return
	
	try:
		# Proses file Excel
		process_excel (file_path, output_path, loop_column)
	except Exception as e:
		messagebox.showerror ("Maaf", f"Error: {e}")


root = tk.Tk ()
root.title ("Bot Loop")
root.geometry ("620x200")
root.resizable (False, False)
style = ttk.Style (root)
style.theme_use ('clam')
frame = ttk.Frame (root, padding = "10")
frame.grid (row = 0, column = 0, sticky = (tk.W, tk.E, tk.N, tk.S))

file_label = ttk.Label (frame, text = "Silakan Pilih File:")
file_label.grid (row = 0, column = 0, pady = 5, sticky = tk.W)
file_entry = ttk.Entry (frame, width = 40)
file_entry.grid (row = 0, column = 1, pady = 5, sticky = tk.W)
file_button = ttk.Button (frame, text = "Pilih", command = select_file)
file_button.grid (row = 0, column = 2, pady = 5, sticky = tk.W)

output_label = ttk.Label (frame, text = "Simpan File Ke:")
output_label.grid (row = 1, column = 0, pady = 5, sticky = tk.W)
output_entry = ttk.Entry (frame, width = 40)
output_entry.grid (row = 1, column = 1, pady = 5, sticky = tk.W)
output_button = ttk.Button (frame, text = "Pilih", command = save_file)
output_button.grid (row = 1, column = 2, pady = 5, sticky = tk.W)

# Ganti Entry dengan Combobox untuk memilih kolom
column_label = ttk.Label (frame, text = "Acuan kolom looping:")
column_label.grid (row = 2, column = 0, pady = 5, sticky = tk.W)

column_combobox = ttk.Combobox (frame, width = 20)  # Gunakan Combobox untuk memilih kolom
column_combobox.grid (row = 2, column = 1, pady = 5, sticky = tk.W)

process_button = ttk.Button (frame, text = "Proses", command = start_processing)
process_button.grid (row = 3, column = 0, columnspan = 3, pady = 20)

root.mainloop ()
