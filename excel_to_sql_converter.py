import pandas as pd
import os
import sys
import logging
import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk

#versione corrente
APP_VERSION = "1.0.50"

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
    # Close existing handlers properly to avoid ResourceWarning about open files
    if logger.hasHandlers():
        for h in list(logger.handlers):
            try:
                h.flush()
                h.close()
            except Exception:
                pass
            try:
                logger.removeHandler(h)
            except Exception:
                pass

    # Create a dedicated FileHandler so we can control its lifecycle explicitly
    handler = logging.FileHandler(log_file, mode='a', encoding='utf-8')
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    handler.setFormatter(formatter)
    handler.setLevel(logging.INFO)
    logger.setLevel(logging.INFO)
    logger.addHandler(handler)
    import weakref
    # Ensure handler will close if it is removed without explicit close() (tests may clear handlers)
    weakref.finalize(handler, handler.close)

    logger.info(f"File di log creato: {log_file}")
    return log_file

def format_insert(db_type, schema, table, df):
    # Helpers: validate and quote SQL identifiers (schema, table, columns)

    def safe_identifier(name):
        """Return a sanitized identifier string or raise ValueError if invalid.

        This function rejects identifiers containing dangerous characters (such as quotes,
        semicolons, brackets, slashes, or newlines) and any whitespace. It does not
        explicitly restrict to Unicode letters, digits, or underscores.

        Note:
            This function does NOT guarantee SQL standard compliance for identifiers.
            It allows any characters except those explicitly blacklisted above, which may
            include characters that are invalid in some SQL dialects but are not explicitly checked.
        """
        if not isinstance(name, str) or name.strip() == "":
            raise ValueError("Identifier must be a non-empty string")
        n = name.strip()
        # Reject identifiers containing any internal whitespace (spaces, tabs, etc.)
        if any(ch.isspace() for ch in n):
            raise ValueError(f"Invalid identifier (contains whitespace): {name}")
        # Reject dangerous punctuation
        for ch in ['"', "'", ';', '[', ']', '\\', '/', '\n', '\r', '@', '-', '*', '%', '`']:
            if ch in n:
                raise ValueError(f"Invalid identifier: {name}")
        # Reject control characters (ASCII < 32)
        if any(ord(ch) < 32 for ch in n):
            raise ValueError(f"Invalid identifier (contains control character): {name}")
        # If passes basic checks, return as-is (we'll quote appropriately when building SQL)
        return n

    statements = []
    columns = [safe_identifier(c) for c in df.columns.tolist()]

    # Precompute quoted column list depending on DB type
    if db_type == 'postgres':
        cols = ", ".join([f'\"{c}\"' for c in columns])
    else:
        # For SQL Server and Oracle use plain names
        cols = ", ".join([f'{c}' for c in columns])

    # Validate schema/table too
    schema_safe = safe_identifier(schema)
    table_safe = safe_identifier(table)

    for _, row in df.iterrows():
        values = []
        for val in row:
            if pd.isnull(val):
                values.append("NULL")
            else:
                # escape single quotes in values
                values.append(f"'{str(val).replace(chr(39), chr(39)*2)}'")
        vals = ", ".join(values)
        # Build INSERT with safe identifiers. Use bracketed [schema].[table] as the
        # canonical reference so it matches the DELETE/USE lines produced elsewhere
        # in the code and the expectations in the test-suite.
        table_ref = f'[{schema_safe}].[{table_safe}]'
        statements.append(f'INSERT INTO {table_ref} ({cols}) VALUES ({vals});')

    logging.info(f"Generati {len(statements)} statements INSERT")
    return "\n".join(statements)

def load_csv_robust(file_path):
    """
    Carica un file CSV provando automaticamente diverse combinazioni di separatori e codifiche.
    Restituisce il DataFrame o solleva un'eccezione se tutti i tentativi falliscono.
    """
    # Combinazioni da provare: (separatore, codifica)
    combinations = [
        (';', 'utf-16'),         # SQL Server export (UTF-16 con BOM)
        (';', 'utf-16le'),       # UTF-16 Little Endian
        (';', 'utf-16be'),       # UTF-16 Big Endian
        (',', 'utf-8'),          # Standard internazionale
        (';', 'utf-8'),          # Standard europeo UTF-8
        (',', 'latin-1'),        # Standard internazionale con codifica europea
        (';', 'latin-1'),        # Standard europeo con codifica europea
        (',', 'cp1252'),         # Windows encoding
        (';', 'cp1252'),         # Windows encoding con punto e virgola
        ('\t', 'utf-8'),         # Tab-separated UTF-8
        ('\t', 'latin-1'),       # Tab-separated latin-1
        ('\t', 'utf-16'),        # Tab-separated UTF-16
        ('|', 'utf-8'),          # Pipe-separated UTF-8
        ('|', 'latin-1'),        # Pipe-separated latin-1
        ('|', 'utf-16'),         # Pipe-separated UTF-16
    ]
    
    best_df = None
    best_score = -1
    best_combination = None
    best_info = None
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
        
        # Penalize BOM and control characters in column names
        bom_penalty = 0
        for col in col_names:
            if isinstance(col, str):
                # Check for BOM characters (UTF-16 BOM: \ufeff, \ufffe, or other control chars)
                if any(ord(c) < 32 or ord(c) in [0xfeff, 0xfffe, 0xfffd] for c in col):
                    bom_penalty += 10
        
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
        
        # Final score: weighted sum with BOM penalty
        score = num_cols * 0.5 + non_empty_rows * 0.2 + col_name_quality * 10 + data_consistency * 10 - unnamed_penalty * 5 - bom_penalty
        return score, num_cols, non_empty_rows
    
    file_size_mb = os.path.getsize(file_path) / (1024 * 1024)
    if file_size_mb > 100:
        logging.warning(f"File molto grande: {file_size_mb:.1f} MB. Potrebbero verificarsi problemi di memoria.")
    chunking = file_size_mb > 10
    for sep, encoding in combinations:
        try:
            if chunking:
                # Carica solo il primo chunk per scoring
                chunk_iter = pd.read_csv(file_path, sep=sep, encoding=encoding, dtype=str, chunksize=100000)
                df = next(chunk_iter)
            else:
                df = pd.read_csv(file_path, sep=sep, encoding=encoding, dtype=str)
            score, num_cols, non_empty_rows = score_dataframe(df)
            logging.info(f"Tentativo {sep}|{encoding}: {num_cols} colonne, {non_empty_rows} righe con dati, score={score:.2f}")
            if score > best_score:
                best_df = df
                best_score = score
                best_info = {'separator': sep, 'encoding': encoding}
                best_combination = (sep, encoding)
        except Exception as e:
            errors.append(f"{sep}|{encoding}: {str(e)}")
            continue
    
    if best_df is not None:
        sep, encoding = best_combination
        logging.info(f"CSV caricato con successo usando separatore '{sep}' e codifica '{encoding}'")
        logging.info(f"Colonne rilevate: {list(best_df.columns)}")
        # Validazione aggiuntiva: se il DataFrame ha una sola colonna e nessuna riga utile, probabilmente il file è corrotto o non valido
        num_cols = len(best_df.columns)
        non_empty_rows = len(best_df.dropna(how='all'))
        if num_cols <= 1 or non_empty_rows == 0:
            error_msg = (f"CSV caricato ma sospetto: {num_cols} colonne, {non_empty_rows} righe con dati. "
                         f"Probabile file non valido o corrotto.")
            logging.error(error_msg)
            raise CSVLoadError(error_msg)
        return best_df, best_info
    else:
        error_msg = "Impossibile caricare il CSV con nessuna combinazione. Errori: " + "; ".join(errors)
        logging.error(error_msg)
        raise CSVLoadError(error_msg)

def convert_file(file_path, db_type, schema, table, database=None):
    setup_logging(file_path)
    ext = os.path.splitext(file_path)[1].lower()
    try:
        ext = os.path.splitext(file_path)[1].lower()
        file_size_mb = os.path.getsize(file_path) / (1024 * 1024)
        chunking = ext == '.csv' and file_size_mb > 10
        if ext == '.csv':
            if not chunking:
                df, csv_info = load_csv_robust(file_path)
            else:
                # Determina separatore/codifica migliore usando load_csv_robust (primo chunk)
                combinations = [(',', 'utf-8'), (';', 'utf-8'), (',', 'latin-1'), (';', 'latin-1'), (',', 'cp1252'), (';', 'cp1252'), ('\t', 'utf-8'), ('\t', 'latin-1'), ('|', 'utf-8'), ('|', 'latin-1')]
                best_sep, best_enc = None, None
                best_score = -1
                for sep, encoding in combinations:
                    try:
                        chunk_iter = pd.read_csv(file_path, sep=sep, encoding=encoding, dtype=str, chunksize=100000)
                        df_chunk = next(chunk_iter)
                        num_cols = len(df_chunk.columns)
                        non_empty_rows = len(df_chunk.dropna(how='all'))
                        unique_names = len(set(df_chunk.columns.tolist()))
                        col_name_quality = (unique_names / num_cols) if num_cols > 0 else 0
                        score = num_cols * 0.5 + non_empty_rows * 0.2 + col_name_quality * 10
                        if score > best_score:
                            best_score = score
                            best_sep, best_enc = sep, encoding
                    except Exception:
                        continue
                if best_sep is None:
                    raise CSVLoadError("Impossibile determinare separatore/codifica per file grande.")
                base = os.path.splitext(os.path.basename(file_path))[0]
                dir_path = os.path.dirname(file_path)
                out_file = os.path.join(dir_path, f"{base}.sql")
                with open(out_file, "w", encoding="utf-8") as f:
                    if db_type == "sqlserver" and database:
                        f.write(f"USE [{database}]\nGO\n\n")
                    f.write(f"DELETE FROM [{schema}].[{table}];\nGO\n\n")
                    chunk_iter = pd.read_csv(file_path, sep=best_sep, encoding=best_enc, dtype=str, chunksize=100000)
                    total_rows = 0
                    for chunk in chunk_iter:
                        sql_insert = format_insert(db_type, schema, table, chunk)
                        f.write(sql_insert + "\n")
                        total_rows += len(chunk)
                logging.info(f"Conversione terminata correttamente. File SQL generato: {out_file}. Righe totali: {total_rows}")
                return f"{os.path.basename(file_path)} -> OK (Generato: {out_file}, Righe: {total_rows})"
        else:
            df = pd.read_excel(file_path)
        if not chunking:
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
        logging.error(f"Errore nel caricamento/conversione dati: {e}")
        return f"{os.path.basename(file_path)} -> Errore nel caricamento/conversione dati: {e}"

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
