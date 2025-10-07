import pandas as pd
import os
import sys
import logging
import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk

#versione corrente
APP_VERSION = "1.0.22"

# Costanti UI
DEFAULT_FONT_FAMILY = "Segoe UI"

logger = None

class CSVLoadError(Exception):
    """Eccezione personalizzata per errori di caricamento CSV"""
    pass

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

ICON_SIZE = (24, 24)
IMAGES_PATH = resource_path("images")

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
                values.append(f"'{{str(val).replace(\"'\", \"''\")}}'")
        cols = ", ".join([f'"{col}"' for col in columns]) if db_type == 'postgres' else ", ".join(columns)
        vals = ", ".join(values)
        statements.append(f'INSERT INTO [{schema}].[{table}] ({cols}) VALUES ({vals});')
    logging.info(f"Generati {len(statements)} statements INSERT")
    return "\n".join(statements)

def load_csv_robust(file_path):
    """
    Carica un file CSV provando automaticamente diverse combinazioni di separatori e codifiche.
    Restituisce il DataFrame o solleva un'eccezione se tutti i tentativi falliscono.
    """
    # Combinazioni da provare: (separatore, codifica)
    combinations = [
        (',', 'utf-8'),          # Standard internazionale
        (';', 'utf-8'),          # Standard europeo UTF-8
        (',', 'latin-1'),        # Standard internazionale con codifica europea
        (';', 'latin-1'),        # Standard europeo con codifica europea
        (',', 'cp1252'),         # Windows encoding
        (';', 'cp1252'),         # Windows encoding con punto e virgola
        ('\t', 'utf-8'),         # Tab-separated UTF-8
        ('\t', 'latin-1'),       # Tab-separated latin-1
        ('|', 'utf-8'),          # Pipe-separated UTF-8
        ('|', 'latin-1'),        # Pipe-separated latin-1
    ]
    
    best_df = None
    best_score = -1
    best_combination = None
    errors = []
    
    def score_dataframe(df):
        # Heuristic scoring for DataFrame quality
        num_cols = len(df.columns)
        non_empty_rows = len(df.dropna(how='all'))
        col_names = df.columns.tolist()
        # Penalize if all columns are unnamed or empty
        num_unnamed = sum(
            (
                (isinstance(col, str) and (col.startswith("Unnamed") or col.strip() == "")) or
                not isinstance(col, str)
            )
            for col in col_names
        )
        unique_names = len(set(col_names))
        # Penalize if all column names are the same
        col_name_quality = (unique_names / num_cols) if num_cols > 0 else 0
        # Penalize if most columns are unnamed
        unnamed_penalty = num_unnamed / num_cols if num_cols > 0 else 1
        # Data consistency: fraction of rows with at least half non-null columns
        if num_cols > 0 and len(df) > 0:
            sufficient_data_rows = (df.notnull().sum(axis=1) >= (num_cols // 2)).sum()
            data_consistency = sufficient_data_rows / len(df)
        else:
            data_consistency = 0
        # Final score: weighted sum
        score = num_cols * 0.5 + non_empty_rows * 0.2 + col_name_quality * 10 + data_consistency * 10 - unnamed_penalty * 5
        return score, num_cols, non_empty_rows
    
    for sep, encoding in combinations:
        try:
            df = pd.read_csv(file_path, sep=sep, encoding=encoding, dtype=str)
            score, num_cols, non_empty_rows = score_dataframe(df)
            logging.info(f"Tentativo {sep}|{encoding}: {num_cols} colonne, {non_empty_rows} righe con dati, score={score:.2f}")
            if score > best_score:
                best_df = df
                best_score = score
                best_combination = (sep, encoding)
        except Exception as e:
            errors.append(f"{sep}|{encoding}: {str(e)}")
            continue
    
    if best_df is not None:
        sep, encoding = best_combination
        logging.info(f"CSV caricato con successo usando separatore '{sep}' e codifica '{encoding}'")
        logging.info(f"Colonne rilevate: {list(best_df.columns)}")
        return best_df
    else:
        error_msg = "Impossibile caricare il CSV con nessuna combinazione. Errori: " + "; ".join(errors)
        logging.error(error_msg)
        raise CSVLoadError(error_msg)

def convert_file(file_path, db_type, schema, table, database=None):
    setup_logging(file_path)
    ext = os.path.splitext(file_path)[1].lower()
    try:
        if ext == '.csv':
            df = load_csv_robust(file_path)
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
                f.write(f"USE [{database}]\nGO\n\n")
            f.write(f"DELETE FROM [{schema}].[{table}];\nGO\n\n")
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
        ico_path = resource_path("images/icon.ico")
        if os.path.exists(ico_path):
            self.iconbitmap(ico_path)
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
        self.version_label = None
        self.show_db_menu()

    def clean_widgets(self):
        for widget in self.winfo_children():
            widget.destroy()
        self._show_version()

    # Mostra la versione nell'angolo in basso a destra
    def _show_version(self):
        if self.version_label:
            self.version_label.destroy()
        self.version_label = tk.Label(
            self,
            text=f"Versione: {APP_VERSION}",
            font=(DEFAULT_FONT_FAMILY, 8),
            fg="grey",
            anchor="se"
        )
        self.version_label.place(relx=1.0, rely=1.0, x=-12, y=-8, anchor="se")

    def show_db_menu(self):
        self.clean_widgets()
        title = tk.Label(self, text="Scegli il database di destinazione:", font=(DEFAULT_FONT_FAMILY, 12, 'bold'))
        title.pack(pady=18)
        frame = tk.Frame(self)
        frame.pack(pady=8)
        btn_oracle = tk.Button(
            frame,
            text=" Oracle",
            font=(DEFAULT_FONT_FAMILY, 12, 'bold'),
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
            font=(DEFAULT_FONT_FAMILY, 12, 'bold'),
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
            font=(DEFAULT_FONT_FAMILY, 12, 'bold'),
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
        back_btn = tk.Button(
            self,
            text=" Indietro",
            font=(DEFAULT_FONT_FAMILY, 10), width=90, height=38,
            compound="left", image=self.arrow_icon,
            padx=12,
            anchor="center",
            command=self.show_db_menu
        )
        back_btn.place(x=8, y=8)  # posizionato sempre in alto a sinistra

        center_frame = tk.Frame(self)
        center_frame.place(relx=0.5, rely=0.13, anchor="n")
        spacing = 13

        file_label = tk.Label(center_frame, text="File Excel:", font=(DEFAULT_FONT_FAMILY, 10))
        file_label.pack(pady=(6,2))
        self.file_entry = tk.Entry(center_frame, width=36, font=(DEFAULT_FONT_FAMILY, 10), justify="center")
        self.file_entry.pack(pady=(0,spacing))
        browse_btn = tk.Button(center_frame, text="Sfoglia", width=16, command=self.browse_file)
        browse_btn.pack(pady=(0,spacing))

        schema_label = tk.Label(center_frame, text="Schema:", font=(DEFAULT_FONT_FAMILY, 10))
        schema_label.pack(pady=(2,2))
        self.schema_entry = tk.Entry(center_frame, width=36, font=(DEFAULT_FONT_FAMILY, 10), justify="center")
        self.schema_entry.pack(pady=(0,spacing))

        table_label = tk.Label(center_frame, text="Tabella:", font=(DEFAULT_FONT_FAMILY, 10))
        table_label.pack(pady=(2,2))
        self.table_entry = tk.Entry(center_frame, width=36, font=(DEFAULT_FONT_FAMILY, 10), justify="center")
        self.table_entry.pack(pady=(0,spacing))

        if db_type == "sqlserver":
            db_label = tk.Label(center_frame, text="Database:", font=(DEFAULT_FONT_FAMILY, 10))
            db_label.pack(pady=(2,2))
            self.db_entry = tk.Entry(center_frame, width=36, font=(DEFAULT_FONT_FAMILY, 10), justify="center")
            self.db_entry.pack(pady=(0,spacing))
        else:
            self.db_entry = None

        convert_btn = tk.Button(center_frame, text="Converti", font=(DEFAULT_FONT_FAMILY, 10), width=23, command=self.start_conversion)
        convert_btn.pack(pady=(2,18))

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
        messagebox.showinfo("Risultato della Conversione", result)

if __name__ == '__main__':
    app = MainApp()
    app.mainloop()
