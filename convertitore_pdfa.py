# MIT License

# Copyright (c) 2023 NATALE AMATO

# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:

# The above copyright notice and this permission notice shall be included in all
# copies or substantial portions of the Software.

# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
# SOFTWARE.

import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
from ttkthemes import ThemedTk
import subprocess
import threading
import os

# info
SOFTWARE_NAME = "Convertitore PDF/A"
VERSION = "1.0.0"
DEVELOPER = "NATALE AMATO"
YEAR = "2023"

def show_info():
    messagebox.showinfo("Info", f"{SOFTWARE_NAME}\nVersione: {VERSION}\nSviluppatore: {DEVELOPER}\nAnno: {YEAR}")

def convert_to_pdfa(input_path, output_path):
    if not os.path.isfile(input_path):
        log_message(f"Il file di input {input_path} non esiste.")
        return

    if not input_path.lower().endswith('.pdf'):
        log_message(f"Il file di input {input_path} non è un file PDF.")
        return

    if os.path.exists(output_path):
        result = messagebox.askquestion("Sovrascrivi?", f"Il file di output {output_path} esiste già. Vuoi sovrascriverlo?")
        if result == 'no':
            log_message(f"Operazione annullata dall'utente.")
            convert_button.config(state=tk.NORMAL, text="Converti in PDF/A")
            progress_bar.stop()
            return

    try:
        command = f'gswin64c.exe -dPDFA -dBATCH -dNOPAUSE -sColorConversionStrategy=UseDeviceIndependentColor -sDEVICE=pdfwrite -sOutputFile="{output_path}" "{input_path}"'

        def run_conversion():
            try:
                subprocess.run(command, shell=True, check=True)
                log_message(f"Il file è stato convertito con successo e salvato come {output_path}")
            except subprocess.CalledProcessError as e:
                log_message(f"Errore durante la conversione: {e}")
            except Exception as e:
                log_message(f"Si è verificato un errore durante la conversione: {e}")
            finally:
                convert_button.config(state=tk.NORMAL, text="Converti in PDF/A")
                progress_bar.stop()

        conversion_thread = threading.Thread(target=run_conversion)
        conversion_thread.start()

    except Exception as e:
        log_message(f"Si è verificato un errore durante la conversione: {e}")
        convert_button.config(state=tk.NORMAL, text="Converti in PDF/A")
        progress_bar.stop()

def browse_input():
    input_file_path = filedialog.askopenfilename(title="Seleziona il file di input", filetypes=[("PDF Files", "*.pdf")])
    if input_file_path:
        input_path_var.set(input_file_path)
        output_path_var.set(os.path.splitext(input_file_path)[0] + "PDF-A.pdf")

def browse_output():
    output_file_path = filedialog.asksaveasfilename(title="Salva il file PDF/A come", defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")])
    if output_file_path:
        output_path_var.set(output_file_path)

def start_conversion():
    input_path = input_path_var.get()
    output_path = output_path_var.get()

    if not input_path or not output_path:
        log_message("Seleziona il file di input e l'output prima di convertire.")
        return

    convert_button.config(state=tk.DISABLED, text="Conversione in corso...")
    progress_bar.start()
    convert_to_pdfa(input_path, output_path)

def log_message(message):
    log_text.insert(tk.END, message + "\n")
    log_text.see(tk.END)

root = ThemedTk(theme="arc")  # Usa un tema moderno
root.title("Convertitore PDF/A")

# Crea un menu nella barra del titolo di sistema
menu_bar = tk.Menu(root)
root.config(menu=menu_bar)

# Aggiungi un'opzione 'Info' per mostrare le informazioni sul software
menu_bar.add_command(label="Info", command=show_info)

input_path_var = tk.StringVar()
output_path_var = tk.StringVar()

input_label = ttk.Label(root, text="Seleziona il file di input:")
input_label.pack(pady=5)
input_entry = ttk.Entry(root, textvariable=input_path_var, state="readonly", width=40)
input_entry.pack(pady=5)
input_button = ttk.Button(root, text="Sfoglia", command=browse_input)
input_button.pack(pady=5)

output_label = ttk.Label(root, text="Salva il file PDF/A come:")
output_label.pack(pady=5)
output_entry = ttk.Entry(root, textvariable=output_path_var, state="readonly", width=40)
output_entry.pack(pady=5)
output_button = ttk.Button(root, text="Sfoglia", command=browse_output)
output_button.pack(pady=5)

convert_button = ttk.Button(root, text="Converti in PDF/A", command=start_conversion)
convert_button.pack(pady=20)

progress_bar = ttk.Progressbar(root, mode='indeterminate', length=300)
progress_bar.pack(pady=10)

log_frame = ttk.Frame(root)
log_frame.pack(pady=10)
log_text = tk.Text(log_frame, height=10, width=50)
log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
log_scroll = ttk.Scrollbar(log_frame, command=log_text.yview)
log_scroll.pack(side=tk.RIGHT, fill=tk.Y)
log_text.config(yscrollcommand=log_scroll.set)

root.mainloop()
