"""
Test suite per excel_to_sql_converter.py

Esegui i test con: python -m pytest test_excel_to_sql_converter.py -v
O semplicemente: python test_excel_to_sql_converter.py
"""

import unittest
import tempfile
import os
import pandas as pd
import logging
from unittest.mock import patch, MagicMock
import sys

# Importa le funzioni da testare
from excel_to_sql_converter import (
    load_csv_robust, 
    format_insert, 
    convert_file, 
    setup_logging,
    CSVLoadError
)


class TestCSVLoading(unittest.TestCase):
    """Test per la funzione load_csv_robust"""
    
    def setUp(self):
        """Setup per ogni test"""
        self.temp_dir = tempfile.mkdtemp()
        
    def tearDown(self):
        """Cleanup dopo ogni test"""
        import shutil
        shutil.rmtree(self.temp_dir, ignore_errors=True)
    
    def create_test_csv(self, content, filename="test.csv", encoding="utf-8"):
        """Helper per creare file CSV di test"""
        filepath = os.path.join(self.temp_dir, filename)
        with open(filepath, 'w', encoding=encoding, newline='') as f:
            f.write(content)
        return filepath
    
    def test_load_csv_comma_utf8(self):
        """Test caricamento CSV standard con virgola e UTF-8"""
        csv_content = "nome,età,città\nMario,30,Roma\nLucia,25,Milano"
        filepath = self.create_test_csv(csv_content)
        
        df = load_csv_robust(filepath)
        
        self.assertEqual(len(df), 2)
        self.assertEqual(list(df.columns), ['nome', 'età', 'città'])
        self.assertEqual(df.iloc[0]['nome'], 'Mario')
    
    def test_load_csv_semicolon_utf8(self):
        """Test caricamento CSV europeo con punto e virgola"""
        csv_content = "nome;età;città\nMario;30;Roma\nLucia;25;Milano"
        filepath = self.create_test_csv(csv_content)
        
        df = load_csv_robust(filepath)
        
        self.assertEqual(len(df), 2)
        self.assertEqual(list(df.columns), ['nome', 'età', 'città'])
        self.assertEqual(df.iloc[0]['nome'], 'Mario')
    
    def test_load_csv_semicolon_latin1(self):
        """Test caricamento CSV con encoding latin-1"""
        csv_content = "nome;età;città\nMàrio;30;Ròma\nLùcia;25;Milàno"
        filepath = self.create_test_csv(csv_content, encoding="latin-1")
        
        df = load_csv_robust(filepath)
        
        self.assertEqual(len(df), 2)
        self.assertEqual(df.iloc[0]['nome'], 'Màrio')
    
    def test_load_csv_tab_separated(self):
        """Test caricamento CSV separato da tab"""
        csv_content = "nome\tetà\tcittà\nMario\t30\tRoma\nLucia\t25\tMilano"
        filepath = self.create_test_csv(csv_content)
        
        df = load_csv_robust(filepath)
        
        self.assertEqual(len(df), 2)
        self.assertEqual(list(df.columns), ['nome', 'età', 'città'])
    
    def test_load_csv_pipe_separated(self):
        """Test caricamento CSV separato da pipe"""
        csv_content = "nome|età|città\nMario|30|Roma\nLucia|25|Milano"
        filepath = self.create_test_csv(csv_content)
        
        df = load_csv_robust(filepath)
        
        self.assertEqual(len(df), 2)
        self.assertEqual(list(df.columns), ['nome', 'età', 'città'])
    
    def test_load_csv_invalid_file(self):
        """Test gestione file CSV invalido"""
        # File che non può essere parsato con nessuna combinazione
        csv_content = "dati corrotti \x00\x01\x02 non parsabili"
        filepath = self.create_test_csv(csv_content)
        
        with self.assertRaises(CSVLoadError):
            load_csv_robust(filepath)
    
    def test_load_csv_empty_file(self):
        """Test gestione file CSV vuoto"""
        filepath = self.create_test_csv("")
        
        with self.assertRaises(CSVLoadError):
            load_csv_robust(filepath)
    
    def test_load_csv_best_separator_selection(self):
        """Test selezione automatica del miglior separatore"""
        # CSV che potrebbe essere interpretato in modi diversi
        csv_content = "col1,col2;col3\nval1,val2;val3\nval4,val5;val6"
        filepath = self.create_test_csv(csv_content)
        
        df = load_csv_robust(filepath)
        
        # Dovrebbe scegliere il separatore che dà più colonne (punto e virgola = 2 colonne)
        self.assertEqual(len(df.columns), 2)


class TestSQLFormatting(unittest.TestCase):
    """Test per la funzione format_insert"""
    
    def setUp(self):
        """Setup dati di test"""
        self.sample_df = pd.DataFrame({
            'id': [1, 2, 3],
            'nome': ['Mario', "Luc'ia", 'Giuseppe'],
            'età': [30, 25, None]
        })
    
    def test_format_insert_postgres(self):
        """Test generazione SQL per PostgreSQL"""
        result = format_insert("postgres", "public", "utenti", self.sample_df)
        
        self.assertIn('"id", "nome", "età"', result)  # Colonne quotate
        self.assertIn("INSERT INTO [public].[utenti]", result)
        self.assertIn("'Luc''ia'", result)  # Escape delle virgolette
        self.assertIn("NULL", result)  # Gestione valori nulli
    
    def test_format_insert_sqlserver(self):
        """Test generazione SQL per SQL Server"""
        result = format_insert("sqlserver", "dbo", "utenti", self.sample_df)
        
        self.assertIn("id, nome, età", result)  # Colonne non quotate per SQL Server
        self.assertIn("INSERT INTO [dbo].[utenti]", result)
        self.assertIn("'Luc''ia'", result)  # Escape delle virgolette
    
    def test_format_insert_oracle(self):
        """Test generazione SQL per Oracle"""
        result = format_insert("oracle", "HR", "EMPLOYEES", self.sample_df)
        
        self.assertIn("INSERT INTO [HR].[EMPLOYEES]", result)
        self.assertIn("'1', 'Mario', '30.0'", result)  # Pandas converte int in float
    
    def test_format_insert_empty_dataframe(self):
        """Test con DataFrame vuoto"""
        empty_df = pd.DataFrame()
        result = format_insert("postgres", "public", "test", empty_df)
        
        self.assertEqual(result, "")
    
    def test_format_insert_special_characters(self):
        """Test gestione caratteri speciali"""
        special_df = pd.DataFrame({
            'text': ["Test's 'quote'", 'Normal text', "Another 'test'"]
        })
        
        result = format_insert("postgres", "test", "table", special_df)
        
        self.assertIn("'Test''s ''quote'''", result)
        self.assertIn("'Another ''test'''", result)


class TestConvertFile(unittest.TestCase):
    """Test per la funzione convert_file"""
    
    def setUp(self):
        """Setup per ogni test"""
        self.temp_dir = tempfile.mkdtemp()
        
    def tearDown(self):
        """Cleanup dopo ogni test"""
        import shutil
        shutil.rmtree(self.temp_dir, ignore_errors=True)
    
    def create_test_file(self, content, filename, encoding="utf-8"):
        """Helper per creare file di test"""
        filepath = os.path.join(self.temp_dir, filename)
        if filename.endswith('.csv'):
            with open(filepath, 'w', encoding=encoding, newline='') as f:
                f.write(content)
        else:
            # Per file Excel, usa pandas
            df = pd.DataFrame({'col1': ['val1', 'val2'], 'col2': ['val3', 'val4']})
            df.to_excel(filepath, index=False)
        return filepath
    
    @patch('excel_to_sql_converter.logging')
    def test_convert_csv_file_success(self, mock_logging):
        """Test conversione CSV con successo"""
        csv_content = "nome,età\nMario,30\nLucia,25"
        csv_path = self.create_test_file(csv_content, "test.csv")
        
        result = convert_file(csv_path, "postgres", "public", "utenti")
        
        self.assertIn("OK", result)
        self.assertIn("test.sql", result)
        
        # Verifica che il file SQL sia stato creato
        sql_path = os.path.join(self.temp_dir, "test.sql")
        self.assertTrue(os.path.exists(sql_path))
        
        # Verifica contenuto SQL
        with open(sql_path, 'r', encoding='utf-8') as f:
            sql_content = f.read()
            self.assertIn("DELETE FROM [public].[utenti]", sql_content)
            self.assertIn("INSERT INTO [public].[utenti]", sql_content)
    
    @patch('excel_to_sql_converter.logging')
    def test_convert_excel_file_success(self, mock_logging):
        """Test conversione file Excel con successo"""
        excel_path = self.create_test_file("", "test.xlsx")
        
        result = convert_file(excel_path, "sqlserver", "dbo", "test_table", "TestDB")
        
        self.assertIn("OK", result)
        
        # Verifica contenuto SQL per SQL Server
        sql_path = os.path.join(self.temp_dir, "test.sql")
        with open(sql_path, 'r', encoding='utf-8') as f:
            sql_content = f.read()
            self.assertIn("USE [TestDB]", sql_content)
            self.assertIn("GO", sql_content)
    
        @patch('excel_to_sql_converter.setup_logging')
        def test_convert_file_not_found(self, mock_setup_logging):
            """Test gestione file non esistente"""
            mock_setup_logging.return_value = "/fake/log/path.log"
        result = convert_file("/path/non/esistente.csv", "postgres", "public", "test")
        
        self.assertIn("Errore nel caricamento dati", result)
    
    @patch('excel_to_sql_converter.load_csv_robust')
    def test_convert_csv_load_error(self, mock_load_csv):
        """Test gestione errore caricamento CSV"""
        mock_load_csv.side_effect = CSVLoadError("Errore test")
        
        csv_path = self.create_test_file("invalid", "test.csv")
        result = convert_file(csv_path, "postgres", "public", "test")
        
        self.assertIn("Errore nel caricamento dati", result)
        self.assertIn("Errore test", result)


class TestLogging(unittest.TestCase):
    """Test per la funzione setup_logging"""
    
    def setUp(self):
        """Setup per ogni test"""
        self.temp_dir = tempfile.mkdtemp()
        
    def tearDown(self):
        """Cleanup dopo ogni test"""
        import shutil
        shutil.rmtree(self.temp_dir, ignore_errors=True)
        # Reset logging
        logging.getLogger().handlers.clear()
    
    def test_setup_logging_creates_log_file(self):
        """Test creazione file di log"""
        test_file = os.path.join(self.temp_dir, "test_file.csv")
        
        log_file = setup_logging(test_file)
        
        expected_log = os.path.join(self.temp_dir, "test_file_log.log")
        self.assertEqual(log_file, expected_log)
        self.assertTrue(os.path.exists(expected_log))
    
    def test_setup_logging_different_extensions(self):
        """Test setup logging con diverse estensioni"""
        test_cases = [
            ("data.xlsx", "data_log.log"),
            ("info.csv", "info_log.log"),
            ("file.xls", "file_log.log")
        ]
        
        for input_file, expected_log in test_cases:
            test_path = os.path.join(self.temp_dir, input_file)
            log_file = setup_logging(test_path)
            expected_path = os.path.join(self.temp_dir, expected_log)
            
            self.assertEqual(log_file, expected_path)


class TestIntegration(unittest.TestCase):
    """Test di integrazione end-to-end"""
    
    def setUp(self):
        """Setup per test di integrazione"""
        self.temp_dir = tempfile.mkdtemp()
        
    def tearDown(self):
        """Cleanup dopo ogni test"""
        import shutil
        shutil.rmtree(self.temp_dir, ignore_errors=True)
    
    def test_complete_csv_to_sql_workflow(self):
        """Test workflow completo CSV -> SQL"""
        # Prepara CSV di test con diversi tipi di dati
        csv_content = """id;nome;età;stipendio;attivo
1;Mario Rossi;30;1500.50;true
2;Lucia Bianc'hi;25;2000.00;false
3;Giuseppe;35;;true"""
        
        csv_path = os.path.join(self.temp_dir, "dipendenti.csv")
        with open(csv_path, 'w', encoding='utf-8', newline='') as f:
            f.write(csv_content)
        
        # Esegui conversione
        result = convert_file(csv_path, "postgres", "hr", "dipendenti")
        
        # Verifica successo
        self.assertIn("OK", result)
        
        # Verifica file SQL generato
        sql_path = os.path.join(self.temp_dir, "dipendenti.sql")
        self.assertTrue(os.path.exists(sql_path))
        
        # Verifica contenuto SQL
        with open(sql_path, 'r', encoding='utf-8') as f:
            sql_content = f.read()
            
            # Controlla struttura SQL
            self.assertIn("DELETE FROM [hr].[dipendenti]", sql_content)
            self.assertIn("GO", sql_content)
            self.assertIn('"id", "nome", "età", "stipendio", "attivo"', sql_content)
            
            # Controlla dati specifici
            self.assertIn("'Mario Rossi'", sql_content)
            self.assertIn("'Lucia Bianc''hi'", sql_content)  # Escape virgolette
            self.assertIn("NULL", sql_content)  # Valore nullo per età di Giuseppe
        
        # Verifica file di log creato
        log_path = os.path.join(self.temp_dir, "dipendenti_log.log")
        self.assertTrue(os.path.exists(log_path))


if __name__ == '__main__':
    # Configura logging per i test
    logging.basicConfig(level=logging.INFO)
    
    # Esegue tutti i test
    unittest.main(verbosity=2)