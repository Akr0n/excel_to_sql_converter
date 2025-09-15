# Copilot Instructions for excel_to_sql_converter

## Project Overview
This project is a GUI tool to convert Excel or CSV files into SQL `INSERT` statements, supporting Postgres, SQL Server, and Oracle. The main script is `excel_to_sql_converter.py`, which provides both the conversion logic and a Tkinter-based GUI.

## Key Files
- `excel_to_sql_converter.py`: Main application. Contains all logic for file selection, conversion, and GUI.
- `conversion.log`: Log file generated at runtime for conversion and error events.
- `README.md`: Basic usage and build instructions.

## Architecture & Data Flow
- The user selects an Excel/CSV file and database type via the GUI.
- The script loads the file into a pandas DataFrame, then generates SQL `INSERT` statements.
- For SQL Server, a `USE <database>` and `DELETE FROM` statement is prepended.
- Output is written to `output_inserts.sql`.
- Logging is used for all major events and errors.

## Developer Workflows
- **Build executable:** Use `pyinstaller --onefile --noconsole excel_to_sql_converter.py` to create a standalone `.exe`.
- **Run locally:** Execute `excel_to_sql_converter.py` with Python 3. GUI will launch.
- **Dependencies:** Requires `pandas` and `tkinter` (standard in most Python installs).
- **Debugging:** Check `conversion.log` for errors and info.

## Project-Specific Conventions
- All output SQL is written to `output_inserts.sql` in the working directory.
- For SQL Server, the script adds `USE <database>` and `DELETE FROM` before inserts, followed by `GO`.
- For Postgres, column names are quoted with double quotes.
- The GUI hides the database name field unless SQL Server is selected.
- Logging is always to `conversion.log`.

## Patterns & Examples
- To add support for a new DB, update `format_insert` and GUI dropdown.
- To change output file, modify the `out_file` variable in `convert_file()`.
- To customize logging, edit the `logging.basicConfig` call at the top.

## External Integrations
- No network or external service dependencies.
- All file dialogs and GUI are handled via Tkinter.

## References
- See `README.md` for build/run instructions.
- See `excel_to_sql_converter.py` for all logic and patterns.

---

**AI agents:**
- Follow the above conventions for new features or bugfixes.
- Keep all logic in `excel_to_sql_converter.py` unless refactoring for modularity.
- Update this file if project structure or conventions change.
