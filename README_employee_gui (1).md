# Employee Management System (GUI)

A Tkinter-based desktop application backed by SQLite.

## Setup

```bash
pip install matplotlib openpyxl
python employee_management_gui.py
```

## First-time use

1. Click "Create an account" to register (any username/password).
2. Log in with those same credentials.
3. Use the sidebar to add employees, view records, search, filter by report, view charts, and export to Excel.

## Notes

- Data is stored in `employees.db` (created automatically).
- Exported Excel files are saved in the same folder, named `employees_export_<timestamp>.xlsx`.
