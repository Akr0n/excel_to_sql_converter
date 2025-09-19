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
        self.geometry("430x420")
        self.resizable(False, False)
        self.db_type = None
        self.icon_images = {}
        self.arrow_icon = None
        # Carica icone database
        for dbname, filename in [
            ('oracle', 'oracle.png'),
            ('postgres', 'postgres.png'),
            ('sqlserver', 'sqlserver.png')
        ]:
            try:
                path_icon = os.path.join(IMAGES_PATH, filename)
                img = Image.open(path_icon).resize(ICON_SIZE, Image.LANCZOS)
                self.icon_images[dbname] = ImageTk.PhotoImage(img)
            except Exception as e:
                self.icon_images[dbname] = None
        # Carica icona Indietro
        try:
            arrow_path = os.path.join(IMAGES_PATH, "barrow.png")
            arrow_img = Image.open(arrow_path).resize(ICON_SIZE, Image.LANCZOS)
            self.arrow_icon = ImageTk.PhotoImage(arrow_img)
        except Exception as e:
            self.arrow_icon = None
        self.show_db_menu()

    def clean_widgets(self):
        for widget in self.winfo_children():
            widget.destroy()

    def show_db_menu(self):
        self.clean_widgets()
        title = tk.Label(self, text="Scegli il database di destinazione:", font=("Segoe UI", 12, 'bold'))
        title.pack(pady=18)
        frame = tk.Frame(self)
        frame.pack(pady=8)
        # Oracle - primo (in alto)
        btn_oracle = tk.Button(
            frame,
            text=" Oracle",
            font=("Segoe UI", 12, 'bold'),
            compound="left",
            image=self.icon_images.get('oracle'),
            width=165,
            anchor='w',
            command=lambda: self.show_main_form("oracle")
        )
        btn_oracle.grid(row=0, column=0, pady=7)
        # Postgres - secondo
        btn_postgres = tk.Button(
            frame,
            text=" Postgres",
            font=("Segoe UI", 12, 'bold'),
            compound="left",
            image=self.icon_images.get('postgres'),
            width=165,
            anchor='w',
            command=lambda: self.show_main_form("postgres")
        )
        btn_postgres.grid(row=1, column=0, pady=7)
        # SQL Server - terzo (in basso)
        btn_sqlserver = tk.Button(
            frame,
            text=" SQL Server",
            font=("Segoe UI", 12, 'bold'),
            compound="left",
            image=self.icon_images.get('sqlserver'),
            width=165,
            anchor='w',
            command=lambda: self.show_main_form("sqlserver")
        )
        btn_sqlserver.grid(row=2, column=0, pady=7)

    def show_main_form(self, db_type):
        self.db_type = db_type
        self.clean_widgets()
        # bottone "Indietro" (in alto a sinistra)
        back_frame = tk.Frame(self)
        back_frame.grid(row=0, column=0, columnspan=2, sticky="nw", padx=(14,0), pady=(14,0))
        back_btn = tk.Button(
            back_frame,
            text=" Indietro",
            font=("Segoe UI", 10),
            width=90,
            height=36,
            image=self.arrow_icon,
            compound="left",      # icona a sinistra, testo a destra
            padx=12,              # padding orizzontale interno
            anchor="center",
            command=self.show_db_menu
        )
        back_btn.pack(fill="both", expand=True)

        # FILE
        file_label = tk.Label(self, text="File Excel:", font=("Segoe UI", 10))
        file_label.grid(row=1, column=0, sticky="w", padx=(18,2), pady=10)
        self.file_entry = tk.Entry(self, width=36, font=("Segoe UI", 10))
        self.file_entry.grid(row=2, column=0, sticky='w', padx=(18,2))
        browse_btn = tk.Button(self, text="Sfoglia", width=10, command=self.browse_file)
        browse_btn.grid(row=2, column=1, padx=(2,14), pady=4, sticky="w")

        # SCHEMA
        schema_label = tk.Label(self, text="Schema:", font=("Segoe UI", 10))
        schema_label.grid(row=3, column=0, sticky="w", padx=(18,2), pady=10)
        self.schema_entry = tk.Entry(self, width=36, font=("Segoe UI", 10))
        self.schema_entry.grid(row=4, column=0, sticky='w', padx=(18,2))

        # TABELLA
        table_label = tk.Label(self, text="Tabella:", font=("Segoe UI", 10))
        table_label.grid(row=5, column=0, sticky="w", padx=(18,2), pady=10)
        self.table_entry = tk.Entry(self, width=36, font=("Segoe UI", 10))
        self.table_entry.grid(row=6, column=0, sticky='w', padx=(18,2))

        # DATABASE (solo SQLServer)
        if db_type == "sqlserver":
            db_label = tk.Label(self, text="Database:", font=("Segoe UI", 10))
            db_label.grid(row=7, column=0, sticky="w", padx=(18,2), pady=10)
            self.db_entry = tk.Entry(self, width=36, font=("Segoe UI", 10))
            self.db_entry.grid(row=8, column=0, sticky='w', padx=(18,2))
        else:
            self.db_entry = None

        # Pulsante Converti in fondo centrale
        convert_btn = tk.Button(self, text="Converti", font=("Segoe UI", 10), width=21, command=self.start_conversion)
        convert_btn.grid(row=9, column=0, columnspan=2, pady=(26,12))

    def browse_file(self):
        file_path = filedialog.askopenfilename(
            filetypes=[
                ("Excel files", "*.xlsx;*.xls;*.csv"),
                ("All files", "*.*")
            ]
        )
        if file_path:
            self.file_entry.delete(0, tk.END)
            self.file_entry.insert(0, file_path)

    def start_conversion(self):
        file_path = self.file_entry.get()
        db_type = self.db_type
        schema = self.schema_entry.get() if self.schema_entry else ""
        table = self.table_entry.get() if self.table_entry else ""
        database = self.db_entry.get() if self.db_entry else None
        if not file_path or not db_type or not schema or not table or not file_path.strip():
            messagebox.showwarning("Attenzione", "Completa tutti i campi obbligatori!")
            return
        result = convert_file(file_path, db_type, schema, table, database)
        messagebox.showinfo("Risultato Conversione", result)

if __name__ == '__main__':
    app = MainApp()
    app.mainloop()
