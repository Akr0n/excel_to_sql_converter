import pandas as pd
import os
import logging
import tkinter as tk
from tkinter import filedialog, ttk, messagebox

logger = None

def setup_logging(file_path):
    global logger
    base = os.path.splitext(os.path.basename(file_path))[0]
    dir_path = os.path.dirname(file_path)
    log_file = os.path.join(dir_path, f"{base}_log.log")
    logger = logging.getLogger()
    # Gestione logging multiplo (reset handler)
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
        statement = f'INSERT INTO {schema}.{table} ({cols}) VALUES ({vals});'
        statements.append(statement)
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
        logging.error(f"Errore caricando i dati: {e}")
        return f"{os.path.basename(file_path)} -> Errore caricamento dati: {e}"
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
        logging.info(f"Conversione OK. File SQL generato: {out_file}")
        return f"{os.path.basename(file_path)} -> OK (Generato: {out_file})"
    except Exception as e:
        logging.error(f"Errore conversione: {e}")
        return f"{os.path.basename(file_path)} -> Errore conversione: {e}"

def browse_files():
    files = filedialog.askopenfilenames(filetypes=[
        ("Excel files", "*.xlsx;*.xls;*.csv"),
        ("All files", "*.*")
    ])
    # Mostra nella field (solo per info visuale) i nomi selezionati separati da ;
    file_entry.delete(0, tk.END)
    file_entry.insert(0, "; ".join(files))

def on_db_change(event):
    db_type = db_combobox.get().lower()
    if db_type == "sqlserver":
        db_label.grid(row=4, column=0, sticky='e')
        db_entry.grid(row=4, column=1, columnspan=2, sticky='w')
    else:
        db_label.grid_remove()
        db_entry.grid_remove()

def start_conversion():
    files = file_entry.get().split("; ")
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

# Crea GUI
root = tk.Tk()
root.title("Excel to SQL Converter")

tk.Label(root, text="File Excel (multi):").grid(row=0, column=0, sticky='e')
file_entry = tk.Entry(root, width=65)
file_entry.grid(row=0, column=1)
browse_btn = tk.Button(root, text="Sfoglia", command=browse_files)
browse_btn.grid(row=0, column=2)

tk.Label(root, text="Tipo Database:").grid(row=1, column=0, sticky='e')
db_combobox = ttk.Combobox(root, values=["Postgres", "SQLServer", "Oracle"], state="readonly")
db_combobox.grid(row=1, column=1, columnspan=2, sticky='w')
db_combobox.current(0)

tk.Label(root, text="Schema:").grid(row=2, column=0, sticky='e')
schema_entry = tk.Entry(root, width=20)
schema_entry.grid(row=2, column=1, columnspan=2, sticky='w')

tk.Label(root, text="Tabella:").grid(row=3, column=0, sticky='e')
table_entry = tk.Entry(root, width=20)
table_entry.grid(row=3, column=1, columnspan=2, sticky='w')

db_label = tk.Label(root, text="Database:")
db_entry = tk.Entry(root, width=20)
# campo nascosto fino alla selezione

convert_btn = tk.Button(root, text="Converti", command=start_conversion)
convert_btn.grid(row=5, column=1, pady=8)

db_combobox.bind("<<ComboboxSelected>>", on_db_change)

root.mainloop()
