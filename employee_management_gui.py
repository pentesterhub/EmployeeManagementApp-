import os
import platform
import subprocess
import tkinter as tk
import tkinter.ttk as ttk
import tkinter.messagebox as messagebox
import sqlite3
import matplotlib
matplotlib.use("Agg")  # واجهة الرسم غير التفاعلية (لا تعرض نافذة خارجية)
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from collections import Counter

# ------------------------------------------------------------------------------------
# دالة لطباعة ملف PDF مباشرة على الطابعة الافتراضية
# ------------------------------------------------------------------------------------
def direct_print_pdf(pdf_path):
    """
    طباعة ملف PDF مباشرة على الطابعة الافتراضية (إن أمكن).
    يختلف السلوك حسب نظام التشغيل:
      - Windows: os.startfile(pdf_path, 'print')
      - Linux/Mac: lp أو lpr
    """
    current_os = platform.system()
    if current_os == "Windows":
        # على ويندوز
        try:
            os.startfile(pdf_path, "print")
        except Exception as e:
            print(f"تعذرت الطباعة على ويندوز: {e}")
    elif current_os == "Linux":
        # على لينكس
        try:
            subprocess.run(["lp", pdf_path])
        except Exception as e:
            print(f"تعذرت الطباعة على لينكس: {e}")
    elif current_os == "Darwin":
        # على ماك
        try:
            subprocess.run(["lp", pdf_path])
        except Exception as e:
            print(f"تعذرت الطباعة على ماك: {e}")
    else:
        print("نظام تشغيل غير مدعوم للطباعة المباشرة.")


# ------------------------------------------------------------------------------------
# دوال قاعدة البيانات
# ------------------------------------------------------------------------------------
DB_NAME = "employees.db"

def create_connection():
    """إنشاء اتصال بقاعدة البيانات SQLite."""
    return sqlite3.connect(DB_NAME)

def create_table():
    """إنشاء جدول الموظفين في قاعدة البيانات (إن لم يكن موجودًا)."""
    conn = create_connection()
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS employees (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            position TEXT,
            salary REAL,
            department TEXT,
            phone TEXT
        )
    """)
    conn.commit()
    conn.close()

# ------------------------------------------------------------------------------------
# التطبيق الرئيسي بواجهة Tkinter
# ------------------------------------------------------------------------------------
class EmployeeManagementApp(tk.Tk):
    """
    تطبيق إدارة شؤون الموظفين بواجهة رسومية متقدمة.
    يشمل:
    - شريط قوائم (File, Help)
    - شريط أدوات (أزرار سريعة)
    - تبويب (Notebook): (1) عرض الموظفين، (2) إحصائيات
    - بحث بالاسم + تصفية بالقسم
    - إحصائيات توزيع الموظفين حسب القسم (مخطط بياني)
    - تصدير CSV
    - طباعة PDF مباشرة (ورقيًا) بعد إنشاء الملف
    """
    def __init__(self):
        super().__init__()
        self.title("إدارة شؤون الموظفين)")
        self.geometry("1000x600")

        # إنشاء الجدول إذا لم يكن موجودًا
        create_table()

        # --------------------------------------------------------------------------------
        # شريط القوائم (Menu Bar)
        # --------------------------------------------------------------------------------
        menubar = tk.Menu(self)
        self.config(menu=menubar)

        menu_file = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="ملف", menu=menu_file)
        menu_file.add_command(label="تصدير إلى CSV", command=self.export_to_csv)
        menu_file.add_command(label="طباعة قائمة الموظفين (PDF)", command=self.print_employees_pdf)
        menu_file.add_separator()
        menu_file.add_command(label="خروج", command=self.quit_app)

        menu_help = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="مساعدة", menu=menu_help)
        menu_help.add_command(label="عن البرنامج", command=self.show_about)

        # --------------------------------------------------------------------------------
        # شريط أدوات علوي (Toolbar)
        # --------------------------------------------------------------------------------
        toolbar = tk.Frame(self, bd=1, relief=tk.RAISED, bg="#f0f0f0")
        toolbar.pack(side=tk.TOP, fill=tk.X)

        btn_add = tk.Button(toolbar, text="إضافة موظف", command=self.add_employee_window, bg="#4CAF50", fg="white")
        btn_add.pack(side=tk.LEFT, padx=2, pady=2)

        btn_edit = tk.Button(toolbar, text="تعديل موظف", command=self.edit_employee_window, bg="#2196F3", fg="white")
        btn_edit.pack(side=tk.LEFT, padx=2, pady=2)

        btn_delete = tk.Button(toolbar, text="حذف موظف", command=self.delete_employee, bg="#F44336", fg="white")
        btn_delete.pack(side=tk.LEFT, padx=2, pady=2)

        btn_refresh = tk.Button(toolbar, text="تحديث", command=self.populate_treeview, bg="#9C27B0", fg="white")
        btn_refresh.pack(side=tk.LEFT, padx=2, pady=2)

        # --------------------------------------------------------------------------------
        # إطار البحث والتصفية
        # --------------------------------------------------------------------------------
        search_frame = tk.Frame(self, bd=2, relief=tk.GROOVE)
        search_frame.pack(side=tk.TOP, fill=tk.X, padx=5, pady=5)

        tk.Label(search_frame, text="بحث بالاسم:").pack(side=tk.LEFT, padx=5)
        self.search_var = tk.StringVar()
        search_entry = tk.Entry(search_frame, textvariable=self.search_var)
        search_entry.pack(side=tk.LEFT, padx=5)

        tk.Label(search_frame, text="تصفية بالقسم:").pack(side=tk.LEFT, padx=5)
        self.dept_filter_var = tk.StringVar()
        self.dept_filter_var.set("الكل")
        dept_filter_cb = ttk.Combobox(search_frame, textvariable=self.dept_filter_var, values=["الكل"], width=10)
        dept_filter_cb.pack(side=tk.LEFT, padx=5)

        btn_do_filter = tk.Button(search_frame, text="تصفية/بحث", command=self.apply_filters)
        btn_do_filter.pack(side=tk.LEFT, padx=5)

        # --------------------------------------------------------------------------------
        # دفتر التبويب (Notebook)
        # --------------------------------------------------------------------------------
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        # التبويب الأول: عرض الموظفين
        self.tab_employees = tk.Frame(self.notebook)
        self.notebook.add(self.tab_employees, text="قائمة الموظفين")

        # التبويب الثاني: إحصائيات
        self.tab_stats = tk.Frame(self.notebook)
        self.notebook.add(self.tab_stats, text="إحصائيات")

        # --------------------------------------------------------------------------------
        # شجرة عرض الموظفين (Treeview)
        # --------------------------------------------------------------------------------
        columns = ("ID", "Name", "Position", "Salary", "Dept", "Phone")
        self.tree = ttk.Treeview(self.tab_employees, columns=columns, show="headings")
        self.tree.heading("ID", text="الرقم")
        self.tree.heading("Name", text="الاسم")
        self.tree.heading("Position", text="الوظيفة")
        self.tree.heading("Salary", text="الراتب")
        self.tree.heading("Dept", text="القسم")
        self.tree.heading("Phone", text="الهاتف")

        self.tree.column("ID", width=50, anchor=tk.CENTER)
        self.tree.column("Name", width=150)
        self.tree.column("Position", width=120)
        self.tree.column("Salary", width=80)
        self.tree.column("Dept", width=100)
        self.tree.column("Phone", width=120)

        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(self.tab_employees, orient=tk.VERTICAL, command=self.tree.yview)
        scrollbar.pack(side=tk.LEFT, fill=tk.Y)
        self.tree.configure(yscroll=scrollbar.set)

        # --------------------------------------------------------------------------------
        # إطار الإحصائيات في التبويب الثاني
        # --------------------------------------------------------------------------------
        self.stats_frame = tk.Frame(self.tab_stats)
        self.stats_frame.pack(fill=tk.BOTH, expand=True)

        # --------------------------------------------------------------------------------
        # شريط الحالة (Status Bar)
        # --------------------------------------------------------------------------------
        self.status_var = tk.StringVar()
        self.status_bar = tk.Label(self, textvariable=self.status_var, bd=1, relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)

        # --------------------------------------------------------------------------------
        # تعبئة البيانات مبدئيًا
        # --------------------------------------------------------------------------------
        self.refresh_dept_combobox()
        self.populate_treeview()
        self.update_stats()

    # ================================================================================
    # وظائف القائمة الرئيسية
    # ================================================================================
    def quit_app(self):
        """إغلاق التطبيق."""
        self.destroy()

    def show_about(self):
        """مربع حوار 'عن البرنامج'."""
        messagebox.showinfo("عن البرنامج", "برنامج إدارة شؤون الموظفين- \nمثال تعليمي ببايثون + Tkinter")

    def export_to_csv(self):
        """
        تصدير بيانات الموظفين إلى ملف CSV.
        """
        import csv
        file_path = tk.filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV Files", "*.csv")])
        if not file_path:
            return

        conn = create_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT id, name, position, salary, department, phone FROM employees")
        rows = cursor.fetchall()
        conn.close()

        with open(file_path, mode="w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(["ID", "Name", "Position", "Salary", "Department", "Phone"])
            for row in rows:
                writer.writerow(row)

        messagebox.showinfo("نجاح", f"تم تصدير البيانات إلى {file_path}")

    def print_employees_pdf(self):
        """
        توليد ملف PDF يحتوي على قائمة الموظفين بتنسيق جميل (جدول)،
        ثم إرساله مباشرةً للطابعة الافتراضية.
        """
        try:
            from reportlab.pdfgen import canvas
            from reportlab.lib.pagesizes import A4
            from reportlab.lib.units import cm
        except ImportError:
            messagebox.showerror("خطأ", "يرجى تثبيت مكتبة reportlab:\n\npip install reportlab")
            return

        file_path = tk.filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF Files", "*.pdf")],
            initialfile="employee_list.pdf"
        )
        if not file_path:
            return

        # جلب بيانات الموظفين
        conn = create_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT id, name, position, salary, department, phone FROM employees")
        rows = cursor.fetchall()
        conn.close()

        # إنشاء كائن Canvas من reportlab
        c = canvas.Canvas(file_path, pagesize=A4)
        width, height = A4

        # عنوان الصفحة
        c.setFont("Helvetica-Bold", 16)
        c.drawString(2 * cm, height - 2 * cm, "قائمة الموظفين")

        # تحديد أعمدة الجدول وأحجامها
        col_widths = [1.5*cm, 4*cm, 3*cm, 2.5*cm, 3*cm, 3*cm]
        headers = ["ID", "الاسم", "الوظيفة", "الراتب", "القسم", "الهاتف"]
        row_height = 0.8 * cm

        x_start = 2 * cm
        y_start = height - 3 * cm
        x = x_start
        y = y_start

        # رسم ترويسة الجدول
        c.setFont("Helvetica-Bold", 10)
        for i, header in enumerate(headers):
            c.drawString(x, y, header)
            x += col_widths[i]
        y -= row_height
        x = x_start

        # رسم الصفوف
        c.setFont("Helvetica", 9)
        for row_data in rows:
            for i, cell in enumerate(row_data):
                text = str(cell) if cell is not None else ""
                c.drawString(x, y, text)
                x += col_widths[i]
            x = x_start
            y -= row_height

            # إذا انتهت الصفحة
            if y < 2 * cm:
                c.showPage()
                c.setFont("Helvetica-Bold", 16)
                c.drawString(2 * cm, height - 2 * cm, "قائمة الموظفين (تابع)")
                y = height - 3 * cm
                c.setFont("Helvetica-Bold", 10)
                # إعادة رسم ترويسة الجدول
                xx = x_start
                for i, header in enumerate(headers):
                    c.drawString(xx, y, header)
                    xx += col_widths[i]
                y -= row_height
                c.setFont("Helvetica", 9)

        c.showPage()
        c.save()

        # تأكيد الإنشاء
        messagebox.showinfo("تم الحفظ", f"تم إنشاء ملف PDF: {file_path}\nسيتم إرساله للطابعة الآن.")
        
        # طباعة الملف PDF ورقيًا
        direct_print_pdf(file_path)

    # ================================================================================
    # دوال البحث والتصفية
    # ================================================================================
    def refresh_dept_combobox(self):
        """تحديث قائمة الأقسام في Combobox."""
        conn = create_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT DISTINCT department FROM employees WHERE department IS NOT NULL AND department <> ''")
        depts = cursor.fetchall()
        conn.close()

        dept_list = ["الكل"] + [d[0] for d in depts]
        self.dept_filter_var.set("الكل")

        # إيجاد الـ Combobox وتحديثه
        for child in self.children.values():
            if isinstance(child, tk.Frame):
                for subchild in child.children.values():
                    if isinstance(subchild, ttk.Combobox):
                        subchild["values"] = dept_list

    def apply_filters(self):
        """تطبيق البحث بالاسم والتصفية بالقسم."""
        self.populate_treeview()

    # ================================================================================
    # عرض الموظفين في Treeview
    # ================================================================================
    def populate_treeview(self):
        """
        جلب بيانات الموظفين من قاعدة البيانات وعرضها في Treeview
        مع الأخذ بالبحث والتصفية.
        """
        # تنظيف الشجرة
        for row in self.tree.get_children():
            self.tree.delete(row)

        search_text = self.search_var.get().strip()
        selected_dept = self.dept_filter_var.get()

        # بناء جملة SQL ديناميكية
        query = "SELECT id, name, position, salary, department, phone FROM employees WHERE 1=1"
        params = []

        if search_text:
            query += " AND name LIKE ?"
            params.append(f"%{search_text}%")

        if selected_dept != "الكل":
            query += " AND department=?"
            params.append(selected_dept)

        conn = create_connection()
        cursor = conn.cursor()
        cursor.execute(query, tuple(params))
        rows = cursor.fetchall()
        conn.close()

        for row_data in rows:
            self.tree.insert("", tk.END, values=row_data)

        # تحديث شريط الحالة (عدد الموظفين)
        self.status_var.set(f"عدد الموظفين: {len(rows)}")

        # تحديث الإحصائيات
        self.update_stats()

    # ================================================================================
    # إنشاء المخطط البياني للإحصائيات
    # ================================================================================
    def update_stats(self):
        """إنشاء/تحديث مخطط بياني بسيط لتوزيع الموظفين حسب القسم."""
        # إزالة أي رسم قديم في تبويب الإحصائيات
        for widget in self.stats_frame.winfo_children():
            widget.destroy()

        # جلب البيانات
        conn = create_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT department FROM employees WHERE department IS NOT NULL AND department <> ''")
        rows = cursor.fetchall()
        conn.close()

        if not rows:
            label_no_data = tk.Label(self.stats_frame, text="لا توجد بيانات لعرض الإحصائيات", fg="red")
            label_no_data.pack(pady=20)
            return

        departments = [r[0] for r in rows]
        counter = Counter(departments)
        labels = list(counter.keys())
        values = list(counter.values())

        fig, ax = plt.subplots(figsize=(5,4), dpi=100)
        ax.bar(labels, values, color="#2196F3")
        ax.set_title("توزيع الموظفين حسب القسم")
        ax.set_xlabel("القسم")
        ax.set_ylabel("عدد الموظفين")
        plt.xticks(rotation=30, ha="right")

        canvas = FigureCanvasTkAgg(fig, master=self.stats_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

    # ================================================================================
    # نوافذ إدخال/تعديل بيانات الموظف
    # ================================================================================
    def add_employee_window(self):
        """
        فتح نافذة منبثقة لإضافة موظف جديد.
        """
        window = tk.Toplevel(self)
        window.title("إضافة موظف جديد")
        window.geometry("400x300")

        lbl_name = tk.Label(window, text="الاسم:")
        lbl_name.pack(pady=5)
        entry_name = tk.Entry(window)
        entry_name.pack()

        lbl_position = tk.Label(window, text="الوظيفة:")
        lbl_position.pack(pady=5)
        entry_position = tk.Entry(window)
        entry_position.pack()

        lbl_salary = tk.Label(window, text="الراتب:")
        lbl_salary.pack(pady=5)
        entry_salary = tk.Entry(window)
        entry_salary.pack()

        lbl_dept = tk.Label(window, text="القسم:")
        lbl_dept.pack(pady=5)
        entry_dept = tk.Entry(window)
        entry_dept.pack()

        lbl_phone = tk.Label(window, text="الهاتف:")
        lbl_phone.pack(pady=5)
        entry_phone = tk.Entry(window)
        entry_phone.pack()

        def on_submit():
            name = entry_name.get().strip()
            if not name:
                messagebox.showwarning("تحذير", "اسم الموظف مطلوب!")
                return

            position = entry_position.get().strip()
            salary = entry_salary.get().strip()
            dept = entry_dept.get().strip()
            phone = entry_phone.get().strip()

            conn = create_connection()
            cursor = conn.cursor()
            cursor.execute("""
                INSERT INTO employees (name, position, salary, department, phone)
                VALUES (?, ?, ?, ?, ?)
            """, (name, position, salary, dept, phone))
            conn.commit()
            conn.close()

            messagebox.showinfo("نجاح", "تمت إضافة الموظف بنجاح.")
            window.destroy()
            self.refresh_dept_combobox()
            self.populate_treeview()

        btn_save = tk.Button(window, text="حفظ", command=on_submit, bg="#4CAF50", fg="white")
        btn_save.pack(pady=10)

    def edit_employee_window(self):
        """
        فتح نافذة لتعديل بيانات موظف موجود.
        """
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("تحذير", "الرجاء اختيار موظف من القائمة.")
            return

        item = self.tree.item(selected[0])
        emp_id, name, position, salary, dept, phone = item["values"]

        window = tk.Toplevel(self)
        window.title("تعديل بيانات الموظف")
        window.geometry("400x300")

        lbl_name = tk.Label(window, text="الاسم:")
        lbl_name.pack(pady=5)
        entry_name = tk.Entry(window)
        entry_name.insert(0, name)
        entry_name.pack()

        lbl_position = tk.Label(window, text="الوظيفة:")
        lbl_position.pack(pady=5)
        entry_position = tk.Entry(window)
        entry_position.insert(0, position)
        entry_position.pack()

        lbl_salary = tk.Label(window, text="الراتب:")
        lbl_salary.pack(pady=5)
        entry_salary = tk.Entry(window)
        entry_salary.insert(0, salary)
        entry_salary.pack()

        lbl_dept = tk.Label(window, text="القسم:")
        lbl_dept.pack(pady=5)
        entry_dept = tk.Entry(window)
        entry_dept.insert(0, dept)
        entry_dept.pack()

        lbl_phone = tk.Label(window, text="الهاتف:")
        lbl_phone.pack(pady=5)
        entry_phone = tk.Entry(window)
        entry_phone.insert(0, phone)
        entry_phone.pack()

        def on_submit_edit():
            new_name = entry_name.get().strip()
            new_position = entry_position.get().strip()
            new_salary = entry_salary.get().strip()
            new_dept = entry_dept.get().strip()
            new_phone = entry_phone.get().strip()

            conn = create_connection()
            cursor = conn.cursor()
            cursor.execute("""
                UPDATE employees
                SET name=?, position=?, salary=?, department=?, phone=?
                WHERE id=?
            """, (new_name, new_position, new_salary, new_dept, new_phone, emp_id))
            conn.commit()
            conn.close()

            messagebox.showinfo("نجاح", "تم تحديث بيانات الموظف.")
            window.destroy()
            self.refresh_dept_combobox()
            self.populate_treeview()

        btn_save = tk.Button(window, text="حفظ التغييرات", command=on_submit_edit, bg="#2196F3", fg="white")
        btn_save.pack(pady=10)

    def delete_employee(self):
        """
        حذف موظف محدد من القائمة ومن قاعدة البيانات.
        """
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("تحذير", "الرجاء اختيار موظف من القائمة.")
            return

        item = self.tree.item(selected[0])
        emp_id = item["values"][0]

        confirm = messagebox.askyesno("تأكيد الحذف", "هل أنت متأكد من حذف الموظف؟")
        if confirm:
            conn = create_connection()
            cursor = conn.cursor()
            cursor.execute("DELETE FROM employees WHERE id=?", (emp_id,))
            conn.commit()
            conn.close()

            messagebox.showinfo("نجاح", "تم حذف الموظف بنجاح.")
            self.refresh_dept_combobox()
            self.populate_treeview()

# ------------------------------------------------------------------------------------
# الدالة الرئيسية لتشغيل التطبيق
# ------------------------------------------------------------------------------------
def main():
    app = EmployeeManagementApp()
    app.mainloop()

if __name__ == "__main__":
    main()
