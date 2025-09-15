import pandas as pd
import os
import logging
import tkinter as tk
from tkinter import filedialog, ttk, messagebox

# Configura logging
logging.basicConfig(
    filename='conversion.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

def format_insert(db_type, schema, table, df):
    statements = []
    columns = df.columns.tolist()
    for _, row in df.iterrows():
        values = []
        for val in row:
            if pd.isnull(val):
                values.append("NULL")
            else:
                values.append(f"'{str(val).replace(\"'\", \"''\")}'"
        cols = ", ".join([f'"{col}"' for col in columns]) if db_type == 'postgres' else ", ".join(columns)
        vals = ", ".join(values)
        statement = f'INSERT INTO {schema}.{table} ({cols}) VALUES ({vals});'
        statements.append(statement)
    logging.info(f"Generati {len(statements)} statements INSERT")
    return "\n".join(statements)

def convert_file(file_path, db_type, schema, table):
    ext = os.path.splitext(file_path)[27].lower()
    try:
        if ext == '.csv':
            df = pd.read_csv(file_path)
        else:
            df = pd.read_excel(file_path)
        logging.info(f"Dati caricati correttamente da {file_path}")
    except Exception as e:
        logging.error(f"Errore caricando i dati: {e}")
        return f"Errore caricamento dati: {e}"
    try:
        sql = format_insert(db_type, schema, table, df)
        out_file = "output_inserts.sql"
        with open(out_file, "w", encoding="utf-8") as f:
            f.write(sql)
        logging.info(f"Conversione OK. File SQL generato: {out_file}")
        return f"File SQL generato: {out_file}"
    except Exception as e:
        logging.error(f"Errore conversione: {e}")
        return f"Errore conversione: {e}"

def browse_file():
    file_path = filedialog.askopenfilename(filetypes=[
        ("Excel files", "*.xlsx;*.xls;*.csv"),
        ("All files", "*.*")
    ])
    file_entry.delete(0, tk.END)
    file_entry.insert(0, file_path)

def start_conversion():
    file_path = file_entry.get()
    db_type = db_combobox.get().lower()
    schema = schema_entry.get()
    table = table_entry.get()
    if not file_path or not db_type or not schema or not table:
        messagebox.showwarning("Attenzione", "Completa tutti i campi!")
        return
    result = convert_file(file_path, db_type, schema, table)
    messagebox.showinfo("Risultato", result)

# Crea GUI
root = tk.Tk()
root.title("Excel to SQL Converter")

tk.Label(root, text="File Excel:").grid(row=0, column=0, sticky='e')
file_entry = tk.Entry(root, width=40)
file_entry.grid(row=0, column=1)
browse_btn = tk.Button(root, text="Sfoglia", command=browse_file)
browse_btn.grid(row=0, column=2)

tk.Label(root, text="Database:").grid(row=1, column=0, sticky='e')
db_combobox = ttk.Combobox(root, values=["Postgres", "SQLServer", "Oracle"], state="readonly")
db_combobox.grid(row=1, column=1, columnspan=2, sticky='w')
db_combobox.current(0)

tk.Label(root, text="Schema:").grid(row=2, column=0, sticky='e')
schema_entry = tk.Entry(root, width=20)
schema_entry.grid(row=2, column=1, columnspan=2, sticky='w')

tk.Label(root, text="Tabella:").grid(row=3, column=0, sticky='e')
table_entry = tk.Entry(root, width=20)
table_entry.grid(row=3, column=1, columnspan=2, sticky='w')

convert_btn = tk.Button(root, text="Converti", command=start_conversion)
convert_btn.grid(row=4, column=1, pady=8)

root.mainloop()
