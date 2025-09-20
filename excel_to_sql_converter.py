import pandas as pd
import os
import logging
import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk

#versione corrente
__version__ = "2.1.4"

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
        self.geometry("430x460")
        self.resizable(False, False)
        self.db_type = None
        self.icon_images = {}
        self.arrow_icon = None
        for dbname, filename in [
            ('oracle', 'oracle.png'),
            ('postgres', 'postgres.png'),
            ('sqlserver', 'sqlserver.png')
        ]:
            try:
                path_icon = os.path.join(IMAGES_PATH, filename)
                img = Image.open(path_icon).resize(ICON_SIZE, Image.LANCZOS)
                self.icon_images[dbname] = ImageTk.PhotoImage(img)
            except Exception:
                self.icon_images[dbname] = None
        try:
            arrow_path = os.path.join(IMAGES_PATH, "barrow.png")
            arrow_img = Image.open(arrow_path).resize(ICON_SIZE, Image.LANCZOS)
            self.arrow_icon = ImageTk.PhotoImage(arrow_img)
        except Exception:
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
        btn_oracle = tk.Button(
            frame,
            text=" Oracle",
            font=("Segoe UI", 12, 'bold'),
            compound="left",
            image=self.icon_images.get('oracle'),
            width=130, height=38,
            padx=12,
            anchor="center",
            command=lambda: self.show_main_form("oracle")
        )
        btn_oracle.grid(row=0, column=0, pady=7)
        btn_postgres = tk.Button(
            frame,
            text=" Postgres",
            font=("Segoe UI", 12, 'bold'),
            compound="left",
            image=self.icon_images.get('postgres'),
            width=130, height=38,
            padx=12,
            anchor="center",
            command=lambda: self.show_main_form("postgres")
        )
        btn_postgres.grid(row=1, column=0, pady=7)
        btn_sqlserver = tk.Button(
            frame,
            text=" SQL Server",
            font=("Segoe UI", 12, 'bold'),
            compound="left",
            image=self.icon_images.get('sqlserver'),
            width=130, height=38,
            padx=12,
            anchor="center",
            command=lambda: self.show_main_form("sqlserver")
        )
        btn_sqlserver.grid(row=2, column=0, pady=7)

    def show_main_form(self, db_type):
        self.db_type = db_type
        self.clean_widgets()

        # Bottone "Indietro" sempre a sinistra sopra il frame centrale
        back_btn = tk.Button(
            self,
            text=" Indietro",
            font=("Segoe UI", 10), width=90, height=38,
            compound="left", image=self.arrow_icon,
            padx=12,
            anchor="center",
            command=self.show_db_menu
        )
        back_btn.place(x=8, y=8)  # posizionato sempre in alto a sinistra

        # Central frame per tutto il resto, centrato nella finestra
        center_frame = tk.Frame(self)
        center_frame.place(relx=0.5, rely=0.15, anchor="n")


        spacing = 13  # Spazio verticale tra i campi

        # FILE
        file_label = tk.Label(center_frame, text="File Excel:", font=("Segoe UI", 10))
        file_label.pack(pady=(6,2))
        self.file_entry = tk.Entry(center_frame, width=36, font=("Segoe UI", 10), justify="center")
        self.file_entry.pack(pady=(0,spacing))
        browse_btn = tk.Button(center_frame, text="Sfoglia", width=16, command=self.browse_file)
        browse_btn.pack(pady=(0,spacing))

        # SCHEMA
        schema_label = tk.Label(center_frame, text="Schema:", font=("Segoe UI", 10))
        schema_label.pack(pady=(2,2))
        self.schema_entry = tk.Entry(center_frame, width=36, font=("Segoe UI", 10), justify="center")
        self.schema_entry.pack(pady=(0,spacing))

        # TABELLA
        table_label = tk.Label(center_frame, text="Tabella:", font=("Segoe UI", 10))
        table_label.pack(pady=(2,2))
        self.table_entry = tk.Entry(center_frame, width=36, font=("Segoe UI", 10), justify="center")
        self.table_entry.pack(pady=(0,spacing))

        # DATABASE (solo SQLServer)
        if db_type == "sqlserver":
            db_label = tk.Label(center_frame, text="Database:", font=("Segoe UI", 10))
            db_label.pack(pady=(2,2))
            self.db_entry = tk.Entry(center_frame, width=36, font=("Segoe UI", 10), justify="center")
            self.db_entry.pack(pady=(0,spacing))
        else:
            self.db_entry = None

        # Converti in fondo centrale nel frame
        convert_btn = tk.Button(center_frame, text="Converti", font=("Segoe UI", 10), width=23, command=self.start_conversion)
        convert_btn.pack(pady=(18,6))

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
