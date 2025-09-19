import pandas as pd
import os
import logging
import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk

logger = None

ICON_SIZE = (24, 24)
IMAGES_PATH = "images"

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

class MainApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Excel to SQL Converter")
        self.geometry("665x260")
        self.resizable(False, False)
        self.db_type = None
        self.icon_images = {}
        self.create_db_selection()
        # Precarica le icone dalla cartella 'images'
        for dbname, filename in [
            ('oracle', 'oracle.png'),
            ('postgres', 'postgres.png'),
            ('sqlserver', 'sqlserver.png')
        ]:
            try:
                path_icon = os.path.join(IMAGES_PATH, filename)
                img = Image.open(path_icon).resize(ICON_SIZE, Image.ANTIALIAS)
                self.icon_images[dbname] = ImageTk.PhotoImage(img)
            except Exception as e:
                self.icon_images[dbname] = None

    def create_db_selection(self):
        self.clean_widgets()
        title = tk.Label(self, text="Scegli il database di destinazione:", font=("Segoe UI", 12, 'bold'))
        title.pack(pady=24)
        frame = tk.Frame(self)
        frame.pack()

        btn_oracle = tk.Button(
            frame,
            text=" Oracle",
            font=("Segoe UI", 11, 'bold'),
            compound="left",
            image=self.icon_images.get('oracle'),
            width=140,
            anchor='w',
            command=lambda: self.show_main_form("oracle")
        )
        btn_oracle.grid(row=0, column=0, padx=16)
        btn_postgres = tk.Button(
            frame,
            text=" Postgres",
            font=("Segoe UI", 11, 'bold'),
            compound="left",
            image=self.icon_images.get('postgres'),
            width=140,
            anchor='w',
            command=lambda: self.show_main_form("postgres")
        )
        btn_postgres.grid(row=0, column=1, padx=16)
        btn_sqlserver = tk.Button(
            frame,
            text=" SQL Server",
            font=("Segoe UI", 11, 'bold'),
            compound="left",
            image=self.icon_images.get('sqlserver'),
            width=160,
            anchor='w',
            command=lambda: self.show_main_form("sqlserver")
        )
        btn_sqlserver.grid(row=0, column=2, padx=16)

    def clean_widgets(self):
        for widget in self.winfo_children():
            widget.destroy()

    def show_main_form(self, db_type):
        self.db_type = db_type
        self.clean_widgets()
        file_label = tk.Label(self, text="File Excel (multi):", font=("Segoe UI", 10))
        file_label.grid(row=0, column=0, sticky="e", padx=(16,2), pady=(16,5))
        self.files_listbox = tk.Listbox(self, width=60, height=4, selectmode=tk.EXTENDED)
        self.files_listbox.grid(row=0, column=1, columnspan=3, sticky='ew', padx=(0,5), pady=(16,5))
        browse_btn = tk.Button(self, text="Sfoglia", width=10, command=self.browse_files)
        browse_btn.grid(row=0, column=4, padx=(2,16), pady=(16,5))

        tk.Label(self, text="Schema:", font=("Segoe UI", 10)).grid(row=2, column=0, sticky='e', padx=(16,2), pady=7)
        self.schema_entry = tk.Entry(self, width=18, font=("Segoe UI", 10))
        self.schema_entry.grid(row=2, column=1, sticky='ew', padx=(0,8), pady=7)

        tk.Label(self, text="Tabella:", font=("Segoe UI", 10)).grid(row=2, column=2, sticky='e', padx=(8,2), pady=7)
        self.table_entry = tk.Entry(self, width=18, font=("Segoe UI", 10))
        self.table_entry.grid(row=2, column=3, sticky='ew', padx=(0,8), pady=7)

        if db_type == "sqlserver":
            self.db_label = tk.Label(self, text="Database:", font=("Segoe UI", 10))
            self.db_label.grid(row=2, column=4, sticky='e', padx=(8,2), pady=7)
            self.db_entry = tk.Entry(self, width=18, font=("Segoe UI", 10))
            self.db_entry.grid(row=2, column=5, sticky='ew', padx=(0,16), pady=7)
        else:
            self.db_label = None
            self.db_entry = None

        convert_btn = tk.Button(self, text="Converti", font=("Segoe UI", 10), command=self.start_conversion)
        convert_btn.grid(row=4, column=0, columnspan=6, pady=(18,16))

    def browse_files(self):
        files = filedialog.askopenfilenames(
            filetypes=[
                ("Excel files", "*.xlsx;*.xls;*.csv"),
                ("All files", "*.*")
            ]
        )
        self.files_listbox.delete(0, tk.END)
        for f in files:
            self.files_listbox.insert(tk.END, f)

    def start_conversion(self):
        files = [self.files_listbox.get(i) for i in range(self.files_listbox.size())]
        db_type = self.db_type
        schema = self.schema_entry.get() if self.schema_entry else ""
        table = self.table_entry.get() if self.table_entry else ""
        database = self.db_entry.get() if self.db_entry else None
        if not files or not db_type or not schema or not table or not files[0].strip():
            messagebox.showwarning("Attenzione", "Completa tutti i campi obbligatori!")
            return
        reports = []
        for file_path in files:
            file_path = file_path.strip()
            if not file_path:
                continue
            reports.append(convert_file(file_path, db_type, schema, table, database))
        messagebox.showinfo("Risultato Conversione", "\n".join(reports))

if __name__ == '__main__':
    app = MainApp()
    app.mainloop()
