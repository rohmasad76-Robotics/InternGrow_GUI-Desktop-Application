"""
Employee Management System - GUI/Desktop Application
--------------------------------------------------------
A Tkinter-based desktop application backed by SQLite.

Features:
- Login
- Dashboard
- CRUD Operations
- Database Integration
- Search
- Reports
- Charts (upgrade)
- Export to Excel (upgrade)
- Modern UI styling (upgrade)

Skills used: Tkinter, SQLite, OOP, GUI Development
"""

import tkinter as tk
from tkinter import ttk, messagebox
import sqlite3
import hashlib
import os
from datetime import datetime

import matplotlib
matplotlib.use("TkAgg")
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

from openpyxl import Workbook

DB_FILE = "employees.db"

# ---------- Color palette (modern UI) ----------
COLOR_BG = "#f4f6fb"
COLOR_SIDEBAR = "#1e3a8a"
COLOR_SIDEBAR_TEXT = "#ffffff"
COLOR_ACCENT = "#2563eb"
COLOR_CARD = "#ffffff"
COLOR_TEXT = "#1f2937"
COLOR_DANGER = "#dc2626"


# ---------- Database layer ----------
def get_connection():
    return sqlite3.connect(DB_FILE)


def initialize_db():
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            password_hash TEXT NOT NULL
        )
    """)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS employees (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            emp_id TEXT UNIQUE NOT NULL,
            name TEXT NOT NULL,
            department TEXT,
            designation TEXT,
            salary REAL
        )
    """)
    conn.commit()
    conn.close()


def hash_password(password):
    return hashlib.sha256(password.encode()).hexdigest()


# ---------- Main Application ----------
class EmployeeManagementApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Employee Management System")
        self.geometry("1000x650")
        self.configure(bg=COLOR_BG)
        self.minsize(900, 600)

        self.current_user = None

        # ttk styling — ensures buttons render their colors reliably across platforms
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Accent.TButton", background=COLOR_ACCENT, foreground="white",
                        font=("Segoe UI", 11, "bold"), padding=10, borderwidth=0)
        style.map("Accent.TButton", background=[("active", "#1e40af")])

        style.configure("Danger.TButton", background=COLOR_DANGER, foreground="white",
                        font=("Segoe UI", 11, "bold"), padding=10, borderwidth=0)
        style.map("Danger.TButton", background=[("active", "#991b1b")])

        style.configure("Sidebar.TButton", background=COLOR_SIDEBAR, foreground=COLOR_SIDEBAR_TEXT,
                        font=("Segoe UI", 11), padding=12, borderwidth=0, anchor="w")
        style.map("Sidebar.TButton", background=[("active", COLOR_ACCENT)])

        style.configure("Link.TButton", background=COLOR_CARD, foreground=COLOR_ACCENT,
                        font=("Segoe UI", 9), borderwidth=0)
        style.map("Link.TButton", background=[("active", COLOR_CARD)])

        # Container that holds every screen (frame)
        self.container = tk.Frame(self, bg=COLOR_BG)
        self.container.pack(fill="both", expand=True)

        self.frames = {}
        for F in (LoginScreen, RegisterScreen):
            frame = F(self.container, self)
            self.frames[F.__name__] = frame
            frame.place(relwidth=1, relheight=1)

        self.show_frame("LoginScreen")

    def show_frame(self, name):
        frame = self.frames[name]
        frame.tkraise()
        if hasattr(frame, "on_show"):
            frame.on_show()

    def launch_dashboard(self, username):
        self.current_user = username
        if "DashboardScreen" not in self.frames:
            frame = DashboardScreen(self.container, self)
            self.frames["DashboardScreen"] = frame
            frame.place(relwidth=1, relheight=1)
        self.show_frame("DashboardScreen")

    def logout(self):
        self.current_user = None
        if "DashboardScreen" in self.frames:
            self.frames["DashboardScreen"].destroy()
            del self.frames["DashboardScreen"]
        self.show_frame("LoginScreen")


# ---------- Login Screen ----------
class LoginScreen(tk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent, bg=COLOR_BG)
        self.app = app

        card = tk.Frame(self, bg=COLOR_CARD, padx=40, pady=40)
        card.place(relx=0.5, rely=0.5, anchor="center")

        tk.Label(card, text="📋 Employee Management System", font=("Segoe UI", 18, "bold"),
                 bg=COLOR_CARD, fg=COLOR_SIDEBAR).pack(pady=(0, 20))

        tk.Label(card, text="Username", font=("Segoe UI", 10, "bold"), bg=COLOR_CARD,
                 fg=COLOR_TEXT, anchor="w").pack(fill="x")
        self.username_entry = tk.Entry(card, font=("Segoe UI", 11), width=30)
        self.username_entry.pack(pady=(4, 12))

        tk.Label(card, text="Password", font=("Segoe UI", 10, "bold"), bg=COLOR_CARD,
                 fg=COLOR_TEXT, anchor="w").pack(fill="x")
        self.password_entry = tk.Entry(card, font=("Segoe UI", 11), width=30, show="*")
        self.password_entry.pack(pady=(4, 20))

        login_btn = ttk.Button(card, text="Login", style="Accent.TButton", command=self.attempt_login)
        login_btn.pack(fill="x")

        register_btn = ttk.Button(card, text="Create an account", style="Link.TButton",
                                   command=lambda: app.show_frame("RegisterScreen"))
        register_btn.pack(pady=(10, 0))

    def on_show(self):
        self.username_entry.delete(0, tk.END)
        self.password_entry.delete(0, tk.END)

    def attempt_login(self):
        username = self.username_entry.get().strip()
        password = self.password_entry.get().strip()

        if not username or not password:
            messagebox.showerror("Login Failed", "Please enter both username and password.")
            return

        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT password_hash FROM users WHERE username = ?", (username,))
        row = cursor.fetchone()
        conn.close()

        if row and row[0] == hash_password(password):
            self.app.launch_dashboard(username)
        else:
            messagebox.showerror("Login Failed", "Invalid username or password.")


# ---------- Register Screen ----------
class RegisterScreen(tk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent, bg=COLOR_BG)
        self.app = app

        card = tk.Frame(self, bg=COLOR_CARD, padx=40, pady=40)
        card.place(relx=0.5, rely=0.5, anchor="center")

        tk.Label(card, text="Create Account", font=("Segoe UI", 18, "bold"),
                 bg=COLOR_CARD, fg=COLOR_SIDEBAR).pack(pady=(0, 20))

        tk.Label(card, text="Username", font=("Segoe UI", 10, "bold"), bg=COLOR_CARD,
                 fg=COLOR_TEXT, anchor="w").pack(fill="x")
        self.username_entry = tk.Entry(card, font=("Segoe UI", 11), width=30)
        self.username_entry.pack(pady=(4, 12))

        tk.Label(card, text="Password", font=("Segoe UI", 10, "bold"), bg=COLOR_CARD,
                 fg=COLOR_TEXT, anchor="w").pack(fill="x")
        self.password_entry = tk.Entry(card, font=("Segoe UI", 11), width=30, show="*")
        self.password_entry.pack(pady=(4, 20))

        register_btn = ttk.Button(card, text="Register", style="Accent.TButton", command=self.register)
        register_btn.pack(fill="x")

        back_btn = ttk.Button(card, text="Back to login", style="Link.TButton",
                               command=lambda: app.show_frame("LoginScreen"))
        back_btn.pack(pady=(10, 0))

    def on_show(self):
        self.username_entry.delete(0, tk.END)
        self.password_entry.delete(0, tk.END)

    def register(self):
        username = self.username_entry.get().strip()
        password = self.password_entry.get().strip()

        if not username or not password:
            messagebox.showerror("Error", "Username and password cannot be empty.")
            return

        conn = get_connection()
        cursor = conn.cursor()
        try:
            cursor.execute(
                "INSERT INTO users (username, password_hash) VALUES (?, ?)",
                (username, hash_password(password)),
            )
            conn.commit()
            messagebox.showinfo("Success", "Account created! You can now log in.")
            self.app.show_frame("LoginScreen")
        except sqlite3.IntegrityError:
            messagebox.showerror("Error", "That username is already taken.")
        finally:
            conn.close()


# ---------- Dashboard Screen ----------
class DashboardScreen(tk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent, bg=COLOR_BG)
        self.app = app

        # Sidebar
        sidebar = tk.Frame(self, bg=COLOR_SIDEBAR, width=220)
        sidebar.pack(side="left", fill="y")
        sidebar.pack_propagate(False)

        tk.Label(sidebar, text="📋 EMS", font=("Segoe UI", 16, "bold"),
                 bg=COLOR_SIDEBAR, fg=COLOR_SIDEBAR_TEXT, pady=20).pack(fill="x")

        nav_buttons = [
            ("➕ Add Employee", self.show_add),
            ("📋 View All", self.show_view),
            ("✏️ Update", self.show_update),
            ("🗑️ Delete", self.show_delete),
            ("🔍 Search", self.show_search),
            ("📊 Reports", self.show_reports),
            ("📤 Export to Excel", self.export_to_excel),
        ]
        for text, cmd in nav_buttons:
            btn = ttk.Button(sidebar, text=text, style="Sidebar.TButton", command=cmd)
            btn.pack(fill="x")

        tk.Frame(sidebar, bg=COLOR_SIDEBAR).pack(fill="both", expand=True)  # spacer
        logout_btn = ttk.Button(sidebar, text="🚪 Logout", style="Danger.TButton", command=app.logout)
        logout_btn.pack(fill="x", side="bottom")

        # Main content area
        self.content = tk.Frame(self, bg=COLOR_BG, padx=30, pady=30)
        self.content.pack(side="left", fill="both", expand=True)

        self.show_view()  # default screen

    def clear_content(self):
        for widget in self.content.winfo_children():
            widget.destroy()

    # ----- Add Employee -----
    def show_add(self):
        self.clear_content()
        tk.Label(self.content, text="Add Employee", font=("Segoe UI", 18, "bold"),
                 bg=COLOR_BG, fg=COLOR_SIDEBAR).pack(anchor="w", pady=(0, 20))

        form = tk.Frame(self.content, bg=COLOR_CARD, padx=30, pady=30)
        form.pack(fill="x")

        fields = {}
        for label_text in ["Employee ID", "Name", "Department", "Designation", "Salary"]:
            tk.Label(form, text=label_text, font=("Segoe UI", 10, "bold"), bg=COLOR_CARD,
                     fg=COLOR_TEXT, anchor="w").pack(fill="x", pady=(8, 2))
            entry = tk.Entry(form, font=("Segoe UI", 11))
            entry.pack(fill="x")
            fields[label_text] = entry

        def save():
            emp_id = fields["Employee ID"].get().strip()
            name = fields["Name"].get().strip()
            department = fields["Department"].get().strip()
            designation = fields["Designation"].get().strip()
            salary_text = fields["Salary"].get().strip()

            if not emp_id or not name:
                messagebox.showerror("Error", "Employee ID and Name are required.")
                return
            try:
                salary = float(salary_text) if salary_text else 0.0
            except ValueError:
                messagebox.showerror("Error", "Salary must be a number.")
                return

            conn = get_connection()
            cursor = conn.cursor()
            try:
                cursor.execute(
                    "INSERT INTO employees (emp_id, name, department, designation, salary) VALUES (?, ?, ?, ?, ?)",
                    (emp_id, name, department, designation, salary),
                )
                conn.commit()
                messagebox.showinfo("Success", f"Employee '{name}' added successfully.")
                self.show_view()
            except sqlite3.IntegrityError:
                messagebox.showerror("Error", "An employee with this ID already exists.")
            finally:
                conn.close()

        ttk.Button(form, text="Save Employee", style="Accent.TButton", command=save).pack(fill="x", pady=(20, 0))

    # ----- View All -----
    def show_view(self):
        self.clear_content()
        tk.Label(self.content, text="All Employees", font=("Segoe UI", 18, "bold"),
                 bg=COLOR_BG, fg=COLOR_SIDEBAR).pack(anchor="w", pady=(0, 20))

        tree = self._build_employee_table(self.content)
        self._load_all_employees(tree)

    def _build_employee_table(self, parent):
        columns = ("emp_id", "name", "department", "designation", "salary")
        tree = ttk.Treeview(parent, columns=columns, show="headings", height=15)
        headings = {"emp_id": "Emp ID", "name": "Name", "department": "Department",
                    "designation": "Designation", "salary": "Salary"}
        for col in columns:
            tree.heading(col, text=headings[col])
            tree.column(col, width=140)
        tree.pack(fill="both", expand=True)
        return tree

    def _load_all_employees(self, tree):
        for row in tree.get_children():
            tree.delete(row)
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT emp_id, name, department, designation, salary FROM employees")
        for row in cursor.fetchall():
            tree.insert("", "end", values=row)
        conn.close()

    # ----- Update -----
    def show_update(self):
        self.clear_content()
        tk.Label(self.content, text="Update Employee", font=("Segoe UI", 18, "bold"),
                 bg=COLOR_BG, fg=COLOR_SIDEBAR).pack(anchor="w", pady=(0, 20))

        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT emp_id, name FROM employees")
        employees = cursor.fetchall()
        conn.close()

        if not employees:
            tk.Label(self.content, text="No employees available yet.", font=("Segoe UI", 11),
                      bg=COLOR_BG, fg="#9ca3af").pack()
            return

        form = tk.Frame(self.content, bg=COLOR_CARD, padx=30, pady=30)
        form.pack(fill="x")

        tk.Label(form, text="Select Employee", font=("Segoe UI", 10, "bold"), bg=COLOR_CARD,
                 fg=COLOR_TEXT, anchor="w").pack(fill="x", pady=(0, 2))
        options = [f"{emp_id} - {name}" for emp_id, name in employees]
        selected = tk.StringVar()
        dropdown = ttk.Combobox(form, textvariable=selected, values=options, state="readonly", font=("Segoe UI", 11))
        dropdown.pack(fill="x", pady=(0, 12))

        fields = {}
        for label_text in ["Name", "Department", "Designation", "Salary"]:
            tk.Label(form, text=f"New {label_text} (leave blank to keep current)", font=("Segoe UI", 10, "bold"),
                      bg=COLOR_CARD, fg=COLOR_TEXT, anchor="w").pack(fill="x", pady=(8, 2))
            entry = tk.Entry(form, font=("Segoe UI", 11))
            entry.pack(fill="x")
            fields[label_text] = entry

        def save_update():
            if not selected.get():
                messagebox.showerror("Error", "Please select an employee.")
                return
            emp_id = selected.get().split(" - ")[0]

            conn = get_connection()
            cursor = conn.cursor()
            cursor.execute("SELECT name, department, designation, salary FROM employees WHERE emp_id = ?", (emp_id,))
            current = cursor.fetchone()

            name = fields["Name"].get().strip() or current[0]
            department = fields["Department"].get().strip() or current[1]
            designation = fields["Designation"].get().strip() or current[2]
            salary_text = fields["Salary"].get().strip()
            try:
                salary = float(salary_text) if salary_text else current[3]
            except ValueError:
                messagebox.showerror("Error", "Salary must be a number.")
                conn.close()
                return

            cursor.execute(
                "UPDATE employees SET name=?, department=?, designation=?, salary=? WHERE emp_id=?",
                (name, department, designation, salary, emp_id),
            )
            conn.commit()
            conn.close()
            messagebox.showinfo("Success", "Employee updated successfully.")
            self.show_view()

        ttk.Button(form, text="Update Employee", style="Accent.TButton", command=save_update).pack(fill="x", pady=(20, 0))

    # ----- Delete -----
    def show_delete(self):
        self.clear_content()
        tk.Label(self.content, text="Delete Employee", font=("Segoe UI", 18, "bold"),
                 bg=COLOR_BG, fg=COLOR_SIDEBAR).pack(anchor="w", pady=(0, 20))

        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT emp_id, name FROM employees")
        employees = cursor.fetchall()
        conn.close()

        if not employees:
            tk.Label(self.content, text="No employees available to delete.", font=("Segoe UI", 11),
                      bg=COLOR_BG, fg="#9ca3af").pack()
            return

        form = tk.Frame(self.content, bg=COLOR_CARD, padx=30, pady=30)
        form.pack(fill="x")

        tk.Label(form, text="Select Employee", font=("Segoe UI", 10, "bold"), bg=COLOR_CARD,
                 fg=COLOR_TEXT, anchor="w").pack(fill="x", pady=(0, 2))
        options = [f"{emp_id} - {name}" for emp_id, name in employees]
        selected = tk.StringVar()
        dropdown = ttk.Combobox(form, textvariable=selected, values=options, state="readonly", font=("Segoe UI", 11))
        dropdown.pack(fill="x", pady=(0, 20))

        def delete():
            if not selected.get():
                messagebox.showerror("Error", "Please select an employee.")
                return
            emp_id = selected.get().split(" - ")[0]
            if not messagebox.askyesno("Confirm", f"Delete employee {selected.get()}?"):
                return
            conn = get_connection()
            cursor = conn.cursor()
            cursor.execute("DELETE FROM employees WHERE emp_id = ?", (emp_id,))
            conn.commit()
            conn.close()
            messagebox.showinfo("Deleted", "Employee deleted successfully.")
            self.show_view()

        ttk.Button(form, text="Delete Employee", style="Danger.TButton", command=delete).pack(fill="x")

    # ----- Search -----
    def show_search(self):
        self.clear_content()
        tk.Label(self.content, text="Search Employee", font=("Segoe UI", 18, "bold"),
                 bg=COLOR_BG, fg=COLOR_SIDEBAR).pack(anchor="w", pady=(0, 20))

        search_bar = tk.Frame(self.content, bg=COLOR_BG)
        search_bar.pack(fill="x", pady=(0, 15))
        search_entry = tk.Entry(search_bar, font=("Segoe UI", 11))
        search_entry.pack(side="left", fill="x", expand=True, ipady=6)

        tree = self._build_employee_table(self.content)

        def do_search():
            term = search_entry.get().strip().lower()
            for row in tree.get_children():
                tree.delete(row)
            conn = get_connection()
            cursor = conn.cursor()
            cursor.execute(
                "SELECT emp_id, name, department, designation, salary FROM employees "
                "WHERE LOWER(emp_id) = ? OR LOWER(name) LIKE ?",
                (term, f"%{term}%"),
            )
            for row in cursor.fetchall():
                tree.insert("", "end", values=row)
            conn.close()

        ttk.Button(search_bar, text="Search", style="Accent.TButton", command=do_search).pack(side="left", padx=(10, 0))

    # ----- Reports (with chart) -----
    def show_reports(self):
        self.clear_content()
        tk.Label(self.content, text="Reports", font=("Segoe UI", 18, "bold"),
                 bg=COLOR_BG, fg=COLOR_SIDEBAR).pack(anchor="w", pady=(0, 20))

        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT COUNT(*), AVG(salary), MAX(salary), MIN(salary) FROM employees")
        count, avg_salary, max_salary, min_salary = cursor.fetchone()

        cursor.execute("SELECT department, COUNT(*) FROM employees GROUP BY department")
        dept_counts = cursor.fetchall()
        conn.close()

        summary = tk.Frame(self.content, bg=COLOR_BG)
        summary.pack(fill="x", pady=(0, 20))

        stats = [
            ("Total Employees", count or 0),
            ("Average Salary", f"{avg_salary:.2f}" if avg_salary else "0"),
            ("Highest Salary", max_salary or 0),
            ("Lowest Salary", min_salary or 0),
        ]
        for label, value in stats:
            card = tk.Frame(summary, bg=COLOR_CARD, padx=20, pady=15)
            card.pack(side="left", fill="both", expand=True, padx=5)
            tk.Label(card, text=label, font=("Segoe UI", 9), bg=COLOR_CARD, fg="#6b7280").pack()
            tk.Label(card, text=str(value), font=("Segoe UI", 16, "bold"), bg=COLOR_CARD,
                      fg=COLOR_SIDEBAR).pack()

        if dept_counts:
            fig = Figure(figsize=(6, 3.2), dpi=100)
            ax = fig.add_subplot(111)
            departments = [d if d else "Unspecified" for d, _ in dept_counts]
            counts = [c for _, c in dept_counts]
            ax.bar(departments, counts, color="#2563eb")
            ax.set_title("Employees per Department")
            ax.set_ylabel("Count")
            fig.tight_layout()

            canvas = FigureCanvasTkAgg(fig, master=self.content)
            canvas.draw()
            canvas.get_tk_widget().pack(fill="both", expand=True)
        else:
            tk.Label(self.content, text="No data available for a chart yet.", font=("Segoe UI", 11),
                      bg=COLOR_BG, fg="#9ca3af").pack()

    # ----- Export to Excel -----
    def export_to_excel(self):
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT emp_id, name, department, designation, salary FROM employees")
        rows = cursor.fetchall()
        conn.close()

        if not rows:
            messagebox.showinfo("Export", "No employee data available to export.")
            return

        wb = Workbook()
        ws = wb.active
        ws.title = "Employees"
        ws.append(["Employee ID", "Name", "Department", "Designation", "Salary"])
        for row in rows:
            ws.append(row)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"employees_export_{timestamp}.xlsx"
        wb.save(filename)
        messagebox.showinfo("Export Successful", f"Employee data exported to '{filename}'.")


if __name__ == "__main__":
    initialize_db()
    app = EmployeeManagementApp()
    app.mainloop()
