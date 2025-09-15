import sys
import os
import logging
import pandas as pd
from PySide6 import QtWidgets, QtCore

# ==================
# Helper functions
# ==================

def setup_logging(file_path):
    base = os.path.splitext(os.path.basename(file_path))[0]
    dir_path = os.path.dirname(file_path)
    log_file = os.path.join(dir_path, f"{base}_log.log")
    logger = logging.getLogger()
    # Remove handlers (important for multiple files!)
    if logger.hasHandlers():
        logger.handlers.clear()
    logging.basicConfig(
        filename=log_file,
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s'
    )
    logger.info(f"File di log creato: {log_file}")
    return logger, log_file

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
        if db_type == 'postgres':
            cols = ', '.join([f'"{col}"' for col in columns])
        else:
            cols = ', '.join(columns)
        vals = ', '.join(values)
        statements.append(f'INSERT INTO {schema}.{table} ({cols}) VALUES ({vals});')
    logging.info(f"Generati {len(statements)} statements INSERT")
    return "\n".join(statements)

def convert_file(file_path, db_type, schema, table, database=None):
    logger, log_file = setup_logging(file_path)
    ext = os.path.splitext(file_path)[1].lower()
    try:
        if ext == '.csv':
            df = pd.read_csv(file_path)
        else:
            df = pd.read_excel(file_path)
        logger.info(f"Dati caricati correttamente da {file_path}")
    except Exception as e:
        logger.error(f"Errore caricando i dati: {e}")
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
        logger.info(f"Conversione OK. File SQL generato: {out_file}")
        return f"{os.path.basename(file_path)} -> OK (SQL: {base}.sql, Log: {base}_log.log)"
    except Exception as e:
        logger.error(f"Errore conversione: {e}")
        return f"{os.path.basename(file_path)} -> Errore conversione: {e}"

# ==================
# PySide6 UI Class
# ==================

class ExcelToSQLConverter(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel to SQL Converter - PySide6 Edition")
        self.resize(580, 260)
        self.setup_ui()

    def setup_ui(self):
        layout = QtWidgets.QVBoxLayout(self)
        # File selection
        file_layout = QtWidgets.QHBoxLayout()
        self.file_edit = QtWidgets.QLineEdit()
        self.file_edit.setPlaceholderText("Nessun file selezionato")
        file_btn = QtWidgets.QPushButton("Sfoglia file Excel/CSV")
        file_btn.clicked.connect(self.select_files)
        file_layout.addWidget(self.file_edit)
        file_layout.addWidget(file_btn)
        layout.addLayout(file_layout)

        # DB and inputs
        form_layout = QtWidgets.QFormLayout()
        self.db_combo = QtWidgets.QComboBox()
        self.db_combo.addItems(["Postgres", "SQLServer", "Oracle"])
        self.db_combo.currentTextChanged.connect(self.on_db_changed)
        form_layout.addRow("Tipo Database:", self.db_combo)

        self.schema_edit = QtWidgets.QLineEdit()
        form_layout.addRow("Schema:", self.schema_edit)

        self.table_edit = QtWidgets.QLineEdit()
        form_layout.addRow("Tabella:", self.table_edit)

        self.dbname_edit = QtWidgets.QLineEdit()
        self.dbname_label = QtWidgets.QLabel("Database:")
        form_layout.addRow(self.dbname_label, self.dbname_edit)
        self.dbname_edit.hide()
        self.dbname_label.hide()
        layout.addLayout(form_layout)

        # Convert button
        self.convert_btn = QtWidgets.QPushButton("Converti")
        self.convert_btn.clicked.connect(self.start_conversion)
        layout.addWidget(self.convert_btn)

        # Results output
        self.results_box = QtWidgets.QTextEdit()
        self.results_box.setReadOnly(True)
        layout.addWidget(self.results_box)

    def select_files(self):
        files, _ = QtWidgets.QFileDialog.getOpenFileNames(
            self,
            "Seleziona uno o più file Excel/CSV",
            "",
            "Excel/CSV (*.xlsx *.xls *.csv)"
        )
        if files:
            self.file_edit.setText("; ".join(files))
        else:
            self.file_edit.setText("")

    def on_db_changed(self, dbtext):
        if dbtext.lower() == "sqlserver":
            self.dbname_edit.show()
            self.dbname_label.show()
        else:
            self.dbname_edit.hide()
            self.dbname_label.hide()

    def start_conversion(self):
        files = self.file_edit.text().split("; ")
        db_type = self.db_combo.currentText().lower()
        schema = self.schema_edit.text().strip()
        table = self.table_edit.text().strip()
        database = self.dbname_edit.text().strip() if db_type == "sqlserver" else None
        if not files or not files[0].strip() or not schema or not table:
            QtWidgets.QMessageBox.warning(self, "Attenzione", "Completa tutti i campi obbligatori e seleziona almeno un file!")
            return
        result_lines = []
        for file_path in files:
            file_path = file_path.strip()
            if not file_path:
                continue
            result = convert_file(file_path, db_type, schema, table, database)
            result_lines.append(result)
        self.results_box.setPlainText("\n".join(result_lines))
        QtWidgets.QMessageBox.information(self, "Risultato", "Conversione completata.\nControlla la cartella dei file per .sql e .log.")

# ==================
# Entrypoint
# ==================

if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    win = ExcelToSQLConverter()
    win.show()
    sys.exit(app.exec())
