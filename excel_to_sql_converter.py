import pandas as pd
import os
import logging
import tkinter as tk
from tkinter import filedialog, ttk, messagebox

logger = None

def setup_logging(file_path):
    base = os.path.splitext(os.path.basename(file_path))[0]
    dir_path = os.path.dirname(file_path)
    log_file = os.path.join(dir_path, f"{base}_log.log")
    global logger
    logger = logging.getLogger()
    if logger.hasHandlers():
        logger.handlers.clear()
    logging.basicConfig(
        filename=log_file,
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s'
    )
    logger.info(f"File di log creato: {log_file}")
    return log_file

def format_insert(db_type, schema, table, df):
    statements = []
    columns = df.columns.tolist()
    for _, row in df.iterrows():
        values = []
        for val in row:
            if pd.isnull(val):
                values.append("NULL")
            else:
                values.append(f"'{str(val).replace("'", "''")}'")
        cols = ", ".join([f'"{col}"' for col in columns]) if db_type == 'postgres' else ", ".join(columns)
        vals = ", ".join(values)
        statements.append(f'INSERT INTO {schema}.{table} ({cols}) VALUES ({vals});')
    logging.info(f"Generati {len(statements)} statements INSERT")
    return "\n".join(statements)

def convert_file(file_path, db_type, schema, table, database=None):
    setup_logging(file_path)
    ext = os.path.splitext(file_path)[1].lower()
    try:
        if ext == '.csv':
            df = pd.read_csv(file_path)
        else:
            df = pd.read_excel(file_path)
        logging.info(f"Dati caricati correttamente da {file_path}")
    except Exception as e:
        logging.error(f"Errore nel caricamento dati: {e}")
        return f"{os.path.basename(file_path)} -> Errore nel caricamento dati: {e}"
    try:
        sql_insert = format_insert(db_type, schema, table, df)
        base = os.path.splitext(os.path.basename(file_path))[0]
        dir_path = os.path.dirname(file_path)
        out_file = os.path.join(dir_path, f"{base}.sql")
        with open(out_file, "w", encoding="utf-8") as f:
            if db_type == "sqlserver" and database:
                f.write(f"USE {database}\nGO\n\n")
            f.write(f"DELETE FROM {schema}.{table};\nGO\n\n")
            f.write(sql_insert)
        logging.info(f"Conversione terminata correttamente. File SQL generato: {out_file}")
        return f"{os.path.basename(file_path)} -> OK (Generato: {out_file})"
    except Exception as e:
        logging.error(f"Errore conversione: {e}")
        return f"{os.path.basename(file_path)} -> Errore conversione: {e}"

def browse_files():
    files = filedialog.askopenfilenames(
        filetypes=[
            ("Excel files", "*.xlsx;*.xls;*.csv"),
            ("All files", "*.*")
        ]
    )
    files_listbox.delete(0, tk.END)
    for f in files:
        files_listbox.insert(tk.END, f)

def on_db_change(event):
    db_type = db_combobox.get().lower()
    if db_type == "sqlserver":
        db_label.grid(row=2, column=7, sticky='e', padx=2, pady=5)
        db_entry.grid(row=2, column=8, sticky='ew', padx=(0,15), pady=5)
    else:
        db_label.grid_remove()
        db_entry.grid_remove()

def start_conversion():
    files = [files_listbox.get(i) for i in range(files_listbox.size())]
    db_type = db_combobox.get().lower()
    schema = schema_entry.get()
    table = table_entry.get()
    database = db_entry.get() if db_type == "sqlserver" else None
    if not files or not db_type or not schema or not table or not files[0].strip():
        messagebox.showwarning("Attenzione", "Completa tutti i campi obbligatori!")
        return
    reports = []
    for file_path in files:
        file_path = file_path.strip()
        if not file_path:
            continue
        reports.append(convert_file(file_path, db_type, schema, table, database))
    messagebox.showinfo("Risultato Conversioni", "\n".join(reports))

# Costruzione GUI
root = tk.Tk()
root.title("Excel to SQL Converter")

# Grid configuration for padding/espansione
for i in range(0, 9):
    root.grid_columnconfigure(i, weight=1, minsize=20)

# Selezione file
tk.Label(root, text="File Excel (multi):").grid(row=0, column=0, columnspan=2, sticky="e", pady=(10,5), padx=(10,3))
files_listbox = tk.Listbox(root, width=70, height=4, selectmode=tk.EXTENDED)
files_listbox.grid(row=0, column=2, columnspan=5, sticky='ew', pady=(10,5), padx=(0,5))
browse_btn = tk.Button(root, text="Sfoglia", command=browse_files)
browse_btn.grid(row=0, column=7, columnspan=2, padx=(3,10), pady=(10,5), sticky='w')

# Riga parametri in linea, tutti sufficientemente larghi
tk.Label(root, text="Tipo Database:").grid(row=2, column=0, sticky='e', padx=(10,2), pady=7)
db_combobox = ttk.Combobox(root, values=["Postgres", "SQLServer", "Oracle"], state="readonly", width=12)
db_combobox.grid(row=2, column=1, sticky='ew', padx=(0,8), pady=7)
db_combobox.current(0)

tk.Label(root, text="Schema:").grid(row=2, column=2, sticky='e', padx=(8,2), pady=7)
schema_entry = tk.Entry(root, width=18)
schema_entry.grid(row=2, column=3, sticky='ew', padx=(0,8), pady=7)

tk.Label(root, text="Tabella:").grid(row=2, column=4, sticky='e', padx=(8,2), pady=7)
table_entry = tk.Entry(root, width=18)
table_entry.grid(row=2, column=5, sticky='ew', padx=(0,8), pady=7)

db_label = tk.Label(root, text="Database:")
db_entry = tk.Entry(root, width=18)

# Bottone converti centrato, largo sull'intera griglia
convert_btn = tk.Button(root, text="Converti", command=start_conversion)
convert_btn.grid(row=3, column=0, columnspan=9, pady=(18,16), sticky='n')

db_combobox.bind("<<ComboboxSelected>>", on_db_change)

# Centra la finestra al lancio di default (facoltativo, puoi togliere se non ti serve)
root.update_idletasks()
w = root.winfo_screenwidth()
h = root.winfo_screenheight()
size = tuple(int(_) for _ in root.geometry().split('+')[0].split('x'))
x = w//2 - size[0]//2
y = h//2 - size[1]//2
root.geometry(f"{size[0]}x{size[1]}+{x}+{y}")

root.mainloop()
