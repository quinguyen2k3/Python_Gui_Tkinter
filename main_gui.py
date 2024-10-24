import sqlite3
import tkinter as tk
from pprint import pprint
from tkinter import messagebox
from tkinter import ttk
from email.mime.base import MIMEBase
from email import encoders
import schedule
import time
import os
import threading
import smtplib
import email
from email.header import decode_header
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import datetime
import imaplib
from tkcalendar import Calendar
from datetime import datetime, timedelta
from app import ChatApplication

import pandas as pd

"""Thực hiện kết nối tạo DB"""
def connect_db():
    conn = sqlite3.connect("students.db")
    cursor = conn.cursor()

    cursor.executescript('''
         CREATE TABLE IF NOT EXISTS students (
            ID INTEGER PRIMARY KEY AUTOINCREMENT,
            MSSV TEXT NOT NULL,
            HoTen TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS courses (
            ID INTEGER PRIMARY KEY AUTOINCREMENT,
            MonHoc TEXT NOT NULL,
            Dot TEXT NOT NULL,
            Lop TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS student_courses (
            ID INTEGER PRIMARY KEY AUTOINCREMENT,
            StudentID INTEGER,
            CourseID INTEGER,
            VangCoPhep INTEGER NOT NULL DEFAULT 0,
            VangKhongPhep INTEGER NOT NULL DEFAULT 0,
            TyLeVang DOUBLE NOT NULL DEFAULT 0,
            FOREIGN KEY (StudentID) REFERENCES students(ID),
            FOREIGN KEY (CourseID) REFERENCES courses(ID)
        );

        CREATE TABLE IF NOT EXISTS absences (
            ID INTEGER PRIMARY KEY AUTOINCREMENT,
            StudentCourseID INTEGER,
            NgayNghi DATE,
            CoPhep BOOLEAN NOT NULL DEFAULT 0,
            FOREIGN KEY (StudentCourseID) REFERENCES student_courses(ID)
        );

            CREATE TABLE IF NOT EXISTS report_statuses (
            ID INTEGER PRIMARY KEY AUTOINCREMENT,
            CourseID INTEGER,
            SubmissionDate DATE,
            SubmissionTime TIME, 
            Status TEXT NOT NULL, 
            Email TEXT NOT NULL,    
            FOREIGN KEY (CourseID) REFERENCES courses(ID)
        );
        
            CREATE TABLE IF NOT EXISTS questions (
            ID INTEGER PRIMARY KEY AUTOINCREMENT,
            StudentID INTEGER,
            QuestionText TEXT NOT NULL,
            SubmissionDateTime DATETIME NOT NULL,
            Status TEXT NOT NULL DEFAULT 'Pending',  -- 'Pending', 'Resolved', 'Closed'
            Email TEXT NOT NULL,
            FOREIGN KEY (StudentID) REFERENCES students(ID)
        );
        ''')
    conn.commit()
    return conn

"""Chức năng đăng nhập"""
logged_in = False

def login():
    dialog = tk.Toplevel(root)
    dialog.title("Đăng nhập")
    dialog.geometry("300x150")

    tk.Label(dialog, text="Tên người dùng:", anchor="w").grid(row=0, column=0, padx=10, pady=10, sticky="w")
    username_entry = tk.Entry(dialog, width=25)
    username_entry.grid(row=0, column=1, padx=10, pady=10)

    tk.Label(dialog, text="Mật khẩu:", anchor="w").grid(row=1, column=0, padx=10, pady=10, sticky="w")
    password_entry = tk.Entry(dialog, show="*", width=25)
    password_entry.grid(row=1, column=1, padx=10, pady=10)

    def check_login():
        global logged_in
        username = username_entry.get()
        password = password_entry.get()
        if username == "admin" and password == "admin":
            logged_in = True
            messagebox.showinfo("Đăng nhập", "Đăng nhập thành công!")
            dialog.destroy()
        else:
            messagebox.showerror("Đăng nhập", "Sai tên người dùng hoặc mật khẩu.")

    login_button = tk.Button(dialog, text="Đăng nhập", width=15, command=check_login)
    login_button.grid(row=2, column=0, columnspan=2, pady=20)

    dialog.grid_columnconfigure(0, weight=1)
    dialog.grid_columnconfigure(1, weight=1)

def check_login(func):
    def wrapper(*args, **kwargs):
        if logged_in:
            return func(*args, **kwargs)
        else:
            messagebox.showerror("Quyền truy cập", "Bạn cần đăng nhập trước.")

    return wrapper

"""Đọc file Excel và lưu vào cơ sở dữ liệu"""
def read_and_save_data(path):
    file_path = path
    conn = sqlite3.connect("students.db")
    cursor = conn.cursor()
    try:
        df = pd.read_excel(file_path, header=None)

        class_code = df.iloc[9, 2] if pd.notnull(df.iloc[9, 2]) else 'N/A'
        batch = df.iloc[5, 2] if pd.notnull(df.iloc[5, 2]) else 'N/A'
        subject = df.iloc[8, 2] if pd.notnull(df.iloc[8, 2]) else 'N/A'

        cursor.execute("SELECT ID, Lop, MonHoc, Dot FROM courses")
        courses = cursor.fetchall()
        course_dict = {(row[1], row[2], row[3]): row[0] for row in courses}

        cursor.execute("SELECT ID, MSSV FROM students")
        students = cursor.fetchall()
        student_dict = {row[1]: row[0] for row in students}

        course_id = course_dict.get((class_code, subject, batch))
        if course_id is None:
            cursor.execute(""" 
                INSERT INTO courses (Lop, MonHoc, Dot) VALUES (?, ?, ?)
            """, (class_code, subject, batch))
            course_id = cursor.lastrowid

        student_course_ids = []
        for index in range(13, len(df)):
            student_id = df.iloc[index, 1]
            student_name = df.iloc[index, 2] + " " + df.iloc[index, 3]
            vang_co_phep = df.iloc[index, 24] if pd.notnull(df.iloc[index, 24]) else 0
            vang_khong_phep = df.iloc[index, 25] if pd.notnull(df.iloc[index, 25]) else 0
            ty_le_vang = df.iloc[index, 27] if pd.notnull(df.iloc[index, 27]) else '0'
            ty_le_vang = float(str(ty_le_vang).replace(',', '.'))

            if student_id not in student_dict:
                cursor.execute("""
                    INSERT INTO students (MSSV, HoTen) VALUES (?, ?)
                """, (student_id, student_name))
                student_dict[student_id] = cursor.lastrowid

                cursor.execute("""
                    INSERT INTO student_courses (StudentID, CourseID, VangCoPhep, VangKhongPhep, TyLeVang)
                    VALUES (?, ?, ?, ?, ?)
                """, (student_dict[student_id], course_id, vang_co_phep, vang_khong_phep, ty_le_vang))
                student_course_ids.append(cursor.lastrowid)
            else:
                student_id_db = student_dict[student_id]
                cursor.execute(""" 
                                    SELECT ID, VangCoPhep, VangKhongPhep, TyLeVang FROM student_courses 
                                    WHERE StudentID = ? AND CourseID = ?
                                """, (student_id_db, course_id))
                existing_course = cursor.fetchone()

                if existing_course:

                    cursor.execute(""" 
                                        UPDATE student_courses 
                                        SET VangCoPhep = ?, VangKhongPhep = ?, TyLeVang = ? 
                                        WHERE ID = ?
                                    """, (vang_co_phep, vang_khong_phep, ty_le_vang, existing_course[0]))
                    student_course_ids.append(existing_course[0])
                else:
                    cursor.execute("""
                                       INSERT INTO student_courses (StudentID, CourseID, VangCoPhep, VangKhongPhep, TyLeVang)
                                       VALUES (?, ?, ?, ?, ?)
                                   """,
                                   (student_dict[student_id], course_id, vang_co_phep, vang_khong_phep, ty_le_vang))
                    student_course_ids.append(cursor.lastrowid)

        save_absence_dates(cursor, df, student_course_ids)
        conn.commit()
        print("Đã nhập sinh viên thành công!")
    except Exception as e:
        print(f"Không thể nhập sinh viên: {str(e)}")
    finally:
        conn.close()

def save_absence_dates(cursor, df, student_course_ids):
    row_date = 11
    start_column = 6
    start_row = 13
    index = 0

    for row in range(start_row, len(df)):
        if index >= len(student_course_ids):
            break

        student_course_id = student_course_ids[index]

        for column_index in range(start_column, len(df.columns), 3):
            date_value = df.iloc[row_date, column_index]

            if isinstance(df.iloc[row_date, column_index], str) and "Tổng cộng" in df.iloc[row_date, column_index]:
                break

            status = df.iloc[row, column_index]

            if pd.notnull(status) and status != "":
                co_phep = 1 if status == 'P' else 0

                cursor.execute(
                    "SELECT COUNT(*) FROM absences WHERE StudentCourseID = ? AND NgayNghi = ?",
                    (student_course_id, date_value)
                )
                exists = cursor.fetchone()[0]

                if exists:
                    cursor.execute(
                        "UPDATE absences SET CoPhep = ? WHERE StudentCourseID = ? AND NgayNghi = ?",
                        (co_phep, student_course_id, date_value)
                    )
                else:
                    cursor.execute(
                        "INSERT INTO absences (StudentCourseID, NgayNghi, CoPhep) VALUES (?, ?, ?)",
                        (student_course_id, date_value, co_phep)
                    )
        index += 1

@check_login
def upload_excel_file():
    dialog = tk.Toplevel(root)
    dialog.title("Nhập đường dẫn file Excel")
    dialog.geometry("300x150")

    tk.Label(dialog, text="Đường dẫn file Excel:").pack(pady=10)
    file_path_entry = tk.Entry(dialog, width=40)
    file_path_entry.pack(pady=5)

    def save_file_path():
        file_path = file_path_entry.get()
        if file_path:
            read_and_save_data(file_path)
            dialog.destroy()
            display_students()
        else:
            messagebox.showwarning("Cảnh báo", "Vui lòng nhập đường dẫn file Excel!")

    tk.Button(dialog, text="Lưu", command=save_file_path).pack(pady=20)

"""Chức năng phân loại, sắp xếp và hiển thị thông tin sinh viên"""
def get_classes():
    conn = sqlite3.connect("students.db")
    cursor = conn.cursor()
    cursor.execute("SELECT DISTINCT Lop FROM courses")
    classes = [row[0] for row in cursor.fetchall()]
    conn.close()
    return classes

def get_subjects():
    conn = sqlite3.connect("students.db")
    cursor = conn.cursor()
    cursor.execute("SELECT DISTINCT MonHoc FROM courses")
    subjects = [row[0] for row in cursor.fetchall()]
    conn.close()
    return subjects

def insert_value_combobox():
    classes = get_classes()
    subjects = get_subjects()
    class_combobox['values'] = classes
    subject_combobox['values']= subjects

def get_students_grouped_by_class(class_name):
    conn = sqlite3.connect("students.db")
    cursor = conn.cursor()
    cursor.execute("""
        SELECT s.MSSV, s.HoTen FROM students s
        JOIN student_courses sc ON s.ID = sc.StudentID
        JOIN courses c ON sc.CourseID = c.ID
        WHERE c.Lop = ?
    """, (class_name,))
    students = cursor.fetchall()
    conn.close()
    return students

def get_students_grouped_by_subject(subject_name):
    conn = sqlite3.connect("students.db")
    cursor = conn.cursor()
    cursor.execute("""
        SELECT s.MSSV, s.HoTen FROM students s
        JOIN student_courses sc ON s.ID = sc.StudentID
        JOIN courses c ON sc.CourseID = c.ID
        WHERE c.MonHoc = ?
    """, (subject_name,))
    students = cursor.fetchall()
    conn.close()
    return students

def get_students_sorted_by_attendance():
    conn = sqlite3.connect("students.db")
    cursor = conn.cursor()

    cursor.execute("""
        SELECT s.MSSV, s.HoTen,
            SUM(sc.VangCoPhep) + SUM(sc.VangKhongPhep) AS total_absences
        FROM students s
        JOIN student_courses sc ON s.ID = sc.StudentID
        GROUP BY s.ID
        ORDER BY total_absences DESC
    """)

    students = cursor.fetchall()
    conn.close()
    return students

def get_students_sorted_by_name():
    conn = sqlite3.connect("students.db")
    cursor = conn.cursor()
    cursor.execute("SELECT MSSV, HoTen FROM students")
    students = cursor.fetchall()

    conn.close()

    sorted_students = sorted(students, key=lambda student: student[1].split()[-1])

    return sorted_students

def get_students():
    conn = sqlite3.connect("students.db")
    cursor = conn.cursor()

    cursor.execute(''' 
    SELECT s.MSSV, s.HoTen
    FROM students AS s 
    ''')
    students = cursor.fetchall()
    conn.close()
    return students


@check_login
def display_students(group_by=None, value=None, order_by=None):
    for item in tree.get_children():
        tree.delete(item)

    try:

        insert_value_combobox()

        if group_by == 'Lop' and value:
            students = get_students_grouped_by_class(value)
        elif group_by == 'MonHoc' and value:
            students = get_students_grouped_by_subject(value)
        elif order_by == 'Vang':
            students = get_students_sorted_by_attendance()
        elif order_by == 'HoTen':
            students = get_students_sorted_by_name()
        else:
            students = get_students()

        if students:
            for student in students:
                tree.insert('', tk.END, values=(
                    student[0],  # MSSV
                    student[1],  # Họ tên
                    "View"
                ))
        else:
            tree.insert('', tk.END, values=("Không có sinh viên nào trong danh sách.", "", ""))
    except Exception as e:
        print("Đã xảy ra lỗi:", e)
        tree.insert('', tk.END, values=("Lỗi khi lấy dữ liệu.", "", ""))


def get_student_details(mssv):
    conn = sqlite3.connect("students.db")
    cursor = conn.cursor()

    query = '''
    SELECT 
        c.MonHoc, 
        sc.VangCoPhep, 
        sc.VangKhongPhep, 
        sc.TyLeVang, 
        a.NgayNghi, 
        a.CoPhep,
        c.Lop
    FROM 
        students s
    JOIN 
        student_courses sc ON s.ID = sc.StudentID
    JOIN 
        courses c ON sc.CourseID = c.ID
    LEFT JOIN 
        absences a ON sc.ID = a.StudentCourseID
    WHERE 
        s.MSSV = ?;
    '''

    cursor.execute(query, (mssv,))
    results = cursor.fetchall()

    student_details = {
        "MonHocs": {}
    }

    for row in results:
        mon_hoc = row[0]
        if mon_hoc not in student_details["MonHocs"]:
            student_details["MonHocs"][mon_hoc] = {
                "VangCoPhep": row[1],
                "VangKhongPhep": row[2],
                "TyLeVang": row[3],
                "Lop": row[6],
                "NgayNghi": []
            }

        if row[4]:
            student_details["MonHocs"][mon_hoc]["NgayNghi"].append({
                "Ngay": row[4],
                "CoPhep": row[5]
            })

    conn.close()

    return student_details


def show_student_details(event):
    selected_item = tree.selection()[0]
    mssv = tree.item(selected_item)['values'][0]
    student_details = get_student_details(mssv)

    details_window = tk.Toplevel(root)
    details_window.title("Thông tin chi tiết môn học")
    details_window.geometry("800x400")

    frame = tk.Frame(details_window)
    frame.pack(fill=tk.BOTH, expand=True)

    details_tree = ttk.Treeview(frame, columns=("MonHoc", "Lop", "VangCoPhep", "VangKhongPhep", "TyLeVang"),
                                show='headings')
    details_tree.heading("MonHoc", text="Môn học")
    details_tree.heading("Lop", text="Lớp")
    details_tree.heading("VangCoPhep", text="Vắng có phép")
    details_tree.heading("VangKhongPhep", text="Vắng không phép")
    details_tree.heading("TyLeVang", text="Tỷ lệ vắng")

    details_tree.column("MonHoc", width=200)
    details_tree.column("Lop", width=75)
    details_tree.column("VangCoPhep", width=150)
    details_tree.column("VangKhongPhep", width=150)
    details_tree.column("TyLeVang", width=150)

    details_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    tree_scroll = tk.Scrollbar(frame, orient=tk.VERTICAL, command=details_tree.yview)
    details_tree.configure(yscrollcommand=tree_scroll.set)
    tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)

    ngay_nghi_frame = tk.Frame(details_window)
    ngay_nghi_frame.pack(fill=tk.BOTH, expand=True)

    ngay_nghi_text = tk.Text(ngay_nghi_frame, wrap=tk.WORD, height=10)
    ngay_nghi_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    ngay_nghi_scroll = tk.Scrollbar(ngay_nghi_frame, orient=tk.VERTICAL, command=ngay_nghi_text.yview)
    ngay_nghi_text.configure(yscrollcommand=ngay_nghi_scroll.set)
    ngay_nghi_scroll.pack(side=tk.RIGHT, fill=tk.Y)

    if student_details and "MonHocs" in student_details:
        for mon_hoc, details in student_details["MonHocs"].items():
            details_tree.insert('', tk.END, values=(
                mon_hoc,
                details["Lop"],
                details["VangCoPhep"],
                details["VangKhongPhep"],
                details["TyLeVang"]
            ))
    else:
        details_tree.insert('', tk.END, values=("Không có thông tin chi tiết cho sinh viên này.", "", "", ""))
        ngay_nghi_text.insert(tk.END, "Không có thông tin ngày nghỉ cho sinh viên này.")

    ngay_nghi_text.config(state=tk.DISABLED)

    def on_tree_select(event):
        selected_item = details_tree.selection()
        if selected_item:
            index = selected_item[0]
            mon_hoc = details_tree.item(index)['values'][0]
            details = student_details["MonHocs"].get(mon_hoc)
            lop = details_tree.item(index)['values'][1]

            ngay_nghi_text.config(state=tk.NORMAL)
            ngay_nghi_text.delete(1.0, tk.END)

            if details:
                ngay_nghi_text.insert(tk.END, f"Môn: {mon_hoc}\n")
                ngay_nghi_text.insert(tk.END, f"Lớp: {lop}\n")
                ngay_nghi_text.insert(tk.END, "Ngày nghỉ: \n" + "\n".join(
                    ngay['Ngay'] for ngay in details["NgayNghi"]) + "\n\n")
            else:
                ngay_nghi_text.insert(tk.END, "Không có thông tin ngày nghỉ cho môn này.")

            ngay_nghi_text.config(state=tk.DISABLED)

    details_tree.bind("<<TreeviewSelect>>", on_tree_select)

"""Chức năng thêm một sinh viên"""

def add_student_to_db(mssv, ho_ten, course_id):
    conn = sqlite3.connect("students.db")
    cursor = conn.cursor()
    cursor.execute('''
            INSERT INTO students (MSSV, HoTen) 
            VALUES (?, ?)
        ''', (mssv, ho_ten))

    student_id = cursor.lastrowid

    cursor.execute('''
           INSERT INTO student_courses (StudentID, CourseID, VangCoPhep, VangKhongPhep, TyLeVang)
           VALUES (?, ?, ?, ?, ?)
       ''', (student_id, course_id, 0, 0, 0))

    conn.commit()
    conn.close()

@check_login
def add_student():
    dialog = tk.Toplevel(root)
    dialog.title("Thêm sinh viên")
    dialog.geometry("300x300")

    tk.Label(dialog, text="MSSV:").grid(row=1, column=0, padx=10, pady=10, sticky='e')
    mssv_entry = tk.Entry(dialog, width=30)
    mssv_entry.grid(row=1, column=1, padx=10, pady=10)

    tk.Label(dialog, text="Họ tên:").grid(row=2, column=0, padx=10, pady=10, sticky='e')
    ho_ten_entry = tk.Entry(dialog, width=30)
    ho_ten_entry.grid(row=2, column=1, padx=10, pady=10)

    tk.Label(dialog, text="Khóa học:").grid(row=3, column=0, padx=10, pady=10, sticky='e')
    course_combobox = ttk.Combobox(dialog, width=30)
    course_combobox.grid(row=3, column=1, padx=10, pady=10)

    course_list = []
    def load_courses():
        global course_list
        conn = sqlite3.connect("students.db")
        cursor = conn.cursor()
        cursor.execute('''
            SELECT ID, MonHoc, Lop, Dot FROM courses
        ''')
        courses = cursor.fetchall()
        course_list =  [(course[0], f"{course[1]} - {course[2]} - {course[3]}") for course in courses]
        course_combobox['values'] = [course[1] for course in course_list]
        conn.close()

    load_courses()

    def save_student():
        global course_list
        mssv = mssv_entry.get()
        ho_ten = ho_ten_entry.get()
        selected_course = course_combobox.get()

        course_id = None
        for course in course_list:
            if course[1] == selected_course:
                course_id = course[0]
                break

        if  mssv and ho_ten:
            try:
                add_student_to_db(mssv, ho_ten,  course_id)
                messagebox.showinfo("Thành công", "Đã thêm sinh viên!")
                dialog.destroy()
                display_students()
            except Exception as e:
                messagebox.showerror("Lỗi", f"Có lỗi xảy ra: {e}")
        else:
            messagebox.showwarning("Cảnh báo", "Vui lòng nhập đầy đủ thông tin!")

    tk.Button(dialog, text="Lưu", command=save_student).grid(row=5, columnspan=2, pady=20)


"""Xóa sinh viên"""
def delete_student_db(mssv):
    conn = sqlite3.connect("students.db")
    cursor = conn.cursor()

    cursor.execute("SELECT ID FROM students WHERE MSSV = ?", (mssv,))
    student_id = cursor.fetchone()

    if student_id:
        student_id = student_id[0]

        cursor.execute(
            "DELETE FROM absences WHERE StudentCourseID IN (SELECT ID FROM student_courses WHERE StudentID = ?)",
            (student_id,))
        cursor.execute("DELETE FROM student_courses WHERE StudentID = ?", (student_id,))

        cursor.execute("DELETE FROM students WHERE MSSV = ?", (mssv,))

    conn.commit()
    rows_deleted = cursor.rowcount
    conn.close()

    return rows_deleted > 0

@check_login
def delete_student():
    selected_item = tree.selection()
    if not selected_item:
        messagebox.showwarning("Cảnh báo", "Vui lòng chọn sinh viên cần xóa từ danh sách.")
        return

    student_info = tree.item(selected_item, 'values')
    mssv = student_info[0]

    confirm = messagebox.askyesno("Xác nhận xóa",
                                  f"Bạn có chắc chắn muốn xóa sinh viên MSSV: {mssv}")
    if confirm:
        result = delete_student_db(mssv)
        if result:
            messagebox.showinfo("Xóa sinh viên", "Xóa sinh viên thành công!")
        else:
            messagebox.showwarning("Xóa sinh viên", "Không tìm thấy sinh viên với MSSV và lớp đã nhập.")
        display_students()

"""Tìm kiếm một sinh viên"""
def find_students(search_value):
    conn = sqlite3.connect("students.db")
    cursor = conn.cursor()

    cursor.execute(''' 
        SELECT s.MSSV, s.HoTen
        FROM students AS s 
        WHERE s.MSSV = ? OR s.HoTen LIKE ?
    ''', (search_value, f"%{search_value}%"))

    results = cursor.fetchall()
    conn.close()
    return results

def search_student():
    search_value = search_entry.get().lower()
    results = find_students(search_value)

    for item in tree.get_children():
        tree.delete(item)

    if results:
        for student in results:

            tree.insert('', tk.END, values=(
                student[0],  # MSSV
                student[1],  # Họ tên
                "View"
            ))
    else:
        tree.insert('', tk.END, values=("Không tìm thấy sinh viên nào.",) + ("",) * 2)

"""Chức năng gửi đi Mail cảnh báo"""
def send_email(receiver_email, subject, body):
    sender_email = "nguyenddqui@gmail.com"
    password = "tpng wimy hsxv xkku"

    message = MIMEMultipart()
    message["From"] = sender_email
    message["To"] = receiver_email
    message["Subject"] = subject
    message.attach(MIMEText(body, "plain"))

    server = smtplib.SMTP("smtp.gmail.com", 587)

    try:

        server.starttls()
        server.login(sender_email, password)
        server.sendmail(sender_email, receiver_email, message.as_string())
        print(f"Email sent to {receiver_email}")
    except Exception as e:
        print(f"Failed to send email to {receiver_email}. Error: {e}")
    finally:
        server.quit()


def get_students_above_50_absence():
    conn = sqlite3.connect("students.db")
    cursor = conn.cursor()

    query = """
        SELECT s.MSSV, s.HoTen, c.MonHoc, c.Lop, sc.TyLeVang, sc.VangCoPhep + sc.VangKhongPhep as TongBuoiVang, a.NgayNghi, a.CoPhep 
        FROM students s
        JOIN student_courses sc ON s.ID = sc.StudentID
        JOIN courses c ON sc.CourseID = c.ID
        LEFT JOIN absences a ON sc.ID = a.StudentCourseID 
        WHERE sc.TyLeVang >= 50
        """

    cursor.execute(query)
    students = cursor.fetchall()
    conn.close()

    return students


def get_students_with_20_absence():
    conn = sqlite3.connect("students.db")
    cursor = conn.cursor()
    cursor.execute("""
        SELECT s.MSSV, s.HoTen, c.MonHoc, c.Lop, sc.TyLeVang, sc.VangCoPhep + sc.VangKhongPhep as TongBuoiVang ,a.NgayNghi, a.CoPhep 
        FROM students s
        JOIN student_courses sc ON s.ID = sc.StudentID
        JOIN courses c ON sc.CourseID = c.ID
        LEFT JOIN absences a ON sc.ID = a.StudentCourseID 
        WHERE sc.TyLeVang >= 20
        """)
    students = cursor.fetchall()
    conn.close()
    return students


def aggregate_students_by_mssv(students):
    student_data = {}
    pprint(students)

    for student in students:
        mssv, ho_ten, mon_hoc,lop, ty_le_vang, tong_buoi, ngay_nghi, co_phep = student

        if mssv not in student_data:

            student_data[mssv] = {
                "HoTen": ho_ten,  # Tên sinh viên
                "MonHocs": {}  # Môn học sẽ được lưu ở đây
            }


        if mon_hoc not in student_data[mssv]["MonHocs"]:

            student_data[mssv]["MonHocs"][mon_hoc] = {
                "Lop": lop,
                "TyLeVang": ty_le_vang,
                "TongBuoiVang": tong_buoi,
                "NgayNghi": []
            }

        student_data[mssv]["MonHocs"][mon_hoc]["NgayNghi"].append({
            "Ngay": ngay_nghi,
            "CoPhep": co_phep
        })
    return student_data

@check_login
def send_warning_emails_thread(recipient_type):
    students = None
    if recipient_type == "parents":
        students = get_students_above_50_absence()
        if not students:
            messagebox.showwarning("Không có dữ liệu", "Không có sinh viên nào cần gửi cảnh báo đến phụ huynh.")
            return
    elif recipient_type == "students":
        students = get_students_with_20_absence()
        if not students:
            messagebox.showwarning("Không có dữ liệu", "Không có sinh viên nào cần gửi cảnh báo.")
            return

    student_data = aggregate_students_by_mssv(students)

    for mssv, info in student_data.items():
        ho_ten = info["HoTen"]
        mon_hocs = info["MonHocs"]

        if mon_hocs:
            email = "bichqui1212@gmail.com"

            if recipient_type == "parents":
                subject = "Cảnh báo vắng mặt nhiều môn học"
                body = f"Kính gửi phụ huynh của {ho_ten} MSSV: {mssv}\n\n"
                body += "Sinh viên đã vắng mặt hơn 50% số buổi học trong các môn học sau:\n\n"
            else:
                subject = "Cảnh báo vắng mặt"
                body = f"Kính gửi {ho_ten} MSSV: {mssv}\n\n"
                body += "Bạn đã vắng mặt hơn 20% số buổi học trong các môn học sau:\n\n"

            for mon_hoc, details in mon_hocs.items():
                lop = details["Lop"]
                ty_le_vang = details["TyLeVang"]
                ngay_nghi = details["NgayNghi"]

                body += f"Môn học: {mon_hoc}, Lớp: {lop}, Tỷ lệ vắng: {ty_le_vang}%\n"

                if ngay_nghi:
                    body += "Các ngày nghỉ:\n"
                    for ngay in ngay_nghi:
                        co_phep = "Có phép" if ngay['CoPhep'] == 1 else "Không có phép"
                        body += f"- Ngày: {ngay['Ngay']} ({co_phep})\n"
                body += "\n"

            body += "Vui lòng xem xét và có biện pháp hỗ trợ sinh viên trong việc học tập.\n\nTrân trọng."
            send_email(email, subject, body)

    messagebox.showinfo("Thành công",
                        f"Đã gửi email cảnh báo cho {'phụ huynh' if recipient_type == 'parents' else 'sinh viên'}.")


def send_warning_emails(recipient_type):
    email_thread = threading.Thread(target=send_warning_emails_thread, args=(recipient_type,))
    email_thread.start()

"""Chức năng gửi Mail báo cáo"""
def get_high_absence_students():

    conn = sqlite3.connect('students.db')
    cursor = conn.cursor()


    query = """
    WITH Total_Students_Per_Class AS (
        SELECT 
            c.Lop,
            COUNT(sc.StudentID) AS TotalStudents
        FROM 
            student_courses sc
        JOIN 
            courses c ON sc.CourseID = c.ID
        GROUP BY 
            c.Lop
    ),
    High_Absence_Students_Per_Class AS (
        SELECT 
            c.Lop,
            COUNT(sc.StudentID) AS HighAbsenceStudents
        FROM 
            student_courses sc
        JOIN 
            courses c ON sc.CourseID = c.ID
        WHERE 
            sc.TyLeVang >= 10
        GROUP BY 
            c.Lop
    ),
    Classes_With_High_Absence_Rate AS (
        SELECT 
            TSPC.Lop
        FROM 
            Total_Students_Per_Class TSPC
        JOIN 
            High_Absence_Students_Per_Class HAPC ON TSPC.Lop = HAPC.Lop
        WHERE 
            HAPC.HighAbsenceStudents > TSPC.TotalStudents * 0.3
    )
    SELECT 
        c.Lop, 
        s.MSSV, 
        s.HoTen, 
        c.MonHoc, 
        sc.TyLeVang, 
        sc.VangCoPhep + sc.VangKhongPhep AS TongBuoiVang, 
        a.NgayNghi, 
        a.CoPhep
    FROM 
        students s
    JOIN 
        student_courses sc ON s.ID = sc.StudentID
    JOIN 
        courses c ON sc.CourseID = c.ID
    JOIN 
        absences a ON sc.ID = a.StudentCourseID
    JOIN 
        Classes_With_High_Absence_Rate hr ON c.Lop = hr.Lop
    WHERE 
        sc.TyLeVang >= 10;
    """

    cursor.execute(query)
    results = cursor.fetchall()

    conn.close()

    return results

def aggregate_students_by_class():
    class_data = {}
    students = get_high_absence_students()

    for student in students:
        lop, mssv, ho_ten, mon_hoc, ty_le_vang, tong_buoi, ngay_nghi, co_phep = student


        if lop not in class_data:
            class_data[lop] = {}

        if mssv not in class_data[lop]:
            class_data[lop][mssv] = {
                "HoTen": ho_ten,
                "MonHocs": {}
            }

        if mon_hoc not in class_data[lop][mssv]["MonHocs"]:
            class_data[lop][mssv]["MonHocs"][mon_hoc] = {
                "TyLeVang": ty_le_vang,
                "TongBuoiVang": tong_buoi,
                "NgayNghi": []
            }

        class_data[lop][mssv]["MonHocs"][mon_hoc]["NgayNghi"].append({
            "Ngay": ngay_nghi,
            "CoPhep": co_phep
        })

    return class_data


def create_report_file(file_name):
    report_data = aggregate_students_by_class()

    with pd.ExcelWriter(file_name) as writer:

        for lop, students_info in report_data.items():
            class_report_data = []

            for mssv, info in students_info.items():
                ho_ten = info["HoTen"]
                for mon_hoc, details in info["MonHocs"].items():
                    ty_le_vang = details["TyLeVang"]
                    tong_buoi_vang = details["TongBuoiVang"]
                    ngay_nghi = details["NgayNghi"]

                    ngay_nghi_list = [f"{ngay_info['Ngay']} (Có phép: {ngay_info['CoPhep']})" for ngay_info in ngay_nghi]
                    ngay_nghi_string = "\n".join(ngay_nghi_list)


                    class_report_data.append({
                        "MSSV": mssv,
                        "HoTen": ho_ten,
                        "MonHoc": mon_hoc,
                        "TyLeVang": ty_le_vang,
                        "TongBuoiVang": tong_buoi_vang,
                        "NgayNghi": ngay_nghi_string,
                    })

            df = pd.DataFrame(class_report_data)
            df.to_excel(writer, sheet_name=lop, index=False)

    print(f"Đã xuất file báo cáo thành công: {file_name}")

def send_file_by_email(receiver_email, subject, body, file_path):
    sender_email = "nguyenddqui@gmail.com"
    password = "tpng wimy hsxv xkku"


    message = MIMEMultipart()
    message["From"] = sender_email
    message["To"] = receiver_email
    message["Subject"] = subject

    message.attach(MIMEText(body, "plain"))

    if file_path:
        try:
            with open(file_path, "rb") as attachment:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(attachment.read())
                encoders.encode_base64(part)
                part.add_header(
                    "Content-Disposition",
                    f"attachment; filename={os.path.basename(file_path)}"
                )
                message.attach(part)
        except Exception as e:
            print(f"Lỗi khi đọc file: {e}")
            return

    server = smtplib.SMTP("smtp.gmail.com", 587)
    try:
        server.starttls()
        server.login(sender_email, password)
        server.sendmail(sender_email, receiver_email, message.as_string())
        print(f"Email sent successfully to {receiver_email}")
    except Exception as e:
        print(f"Failed to send email to {receiver_email}. Error: {e}")
    finally:
        server.quit()

def schedule_send_report():
    print("Job is running")
    file_name = "report_students_absences.xlsx"
    create_report_file(file_name)
    subject = "Báo cáo vắng mặt sinh viên"
    body = "Đây là báo cáo tổng hợp vắng mặt của sinh viên."
    to_emails = ["bichqui1212@gmail.com"]
    for email in to_emails:
        send_file_by_email(email, subject, body, file_name)


"""Chức năng đặt Deadline"""
def fetch_recent_dots():
    conn = sqlite3.connect('students.db')
    cursor = conn.cursor()
    query = "SELECT DISTINCT Dot FROM courses ORDER BY ID DESC LIMIT 5"
    cursor.execute(query)
    dots = [row[0] for row in cursor.fetchall()]
    conn.close()
    return dots


def open_deadline_panel():
    deadline_window = tk.Toplevel(root)
    deadline_window.title("Đặt Deadline")
    deadline_window.geometry("400x500")
    deadline_window.configure(bg="#f0f0f0")

    tk.Label(deadline_window, text="Đặt Deadline Báo Cáo", font=("Arial", 14, "bold"), bg="#f0f0f0").pack(pady=10)

    date_frame = tk.Frame(deadline_window, bg="#e0f7fa", padx=10, pady=10)
    date_frame.pack(pady=10, fill=tk.X)

    tk.Label(date_frame, text="Chọn ngày:", font=("Arial", 12), bg="#e0f7fa").pack(side=tk.LEFT, padx=5)
    cal = Calendar(date_frame, selectmode='day', date_pattern='yyyy-mm-dd')
    cal.pack(side=tk.LEFT, padx=5)

    time_frame = tk.Frame(deadline_window, bg="#e0f7fa", padx=10, pady=10)
    time_frame.pack(pady=10, fill=tk.X)

    tk.Label(time_frame, text="Chọn giờ:", font=("Arial", 12), bg="#e0f7fa").pack(side=tk.LEFT, padx=5)

    hour_combobox = ttk.Combobox(time_frame, values=[f"{str(i).zfill(2)}" for i in range(24)], state="readonly",
                                 font=("Arial", 12), width=5)
    hour_combobox.set("Giờ")
    hour_combobox.pack(side=tk.LEFT, padx=5)


    minute_combobox = ttk.Combobox(time_frame, values=[f"{str(i).zfill(2)}" for i in range(60)], state="readonly",
                                   font=("Arial", 12), width=5)
    minute_combobox.set("Phút")
    minute_combobox.pack(side=tk.LEFT, padx=5)

    dot_frame = tk.Frame(deadline_window, bg="#e0f7fa", padx=10, pady=10)
    dot_frame.pack(pady=10, fill=tk.X)

    tk.Label(dot_frame, text="Chọn học kỳ:", font=("Arial", 12), bg="#e0f7fa").pack(side=tk.LEFT, padx=5)
    dot_combobox = ttk.Combobox(dot_frame, values=fetch_recent_dots(), state="readonly", font=("Arial", 12), width=12)
    dot_combobox.pack(side=tk.LEFT, padx=5)


    def save_deadline(date_str, time_str, dot):
        if not date_str or not time_str or time_str == "Giờ" or time_str == "Phút":
            messagebox.showerror("Lỗi", "Vui lòng chọn ngày và giờ hợp lệ!")
            return

        conn = sqlite3.connect('students.db')
        cursor = conn.cursor()
        try:

            query = """
                INSERT INTO report_statuses (CourseID, SubmissionDate, SubmissionTime, Status, Email)
                SELECT ID, ?, ?, 'NO', '' FROM courses WHERE Dot = ?
                """
            cursor.execute(query, (date_str, time_str, dot))
            conn.commit()

            messagebox.showinfo("Thông báo", "Deadline đã được đặt thành công!")
            deadline_window.destroy()
        except ValueError as ve:
            messagebox.showerror("Lỗi", f"Định dạng ngày hoặc giờ không hợp lệ! {ve}")
        except Exception as e:
            messagebox.showerror("Lỗi", f"Có lỗi xảy ra khi lưu deadline: {str(e)}")
        finally:
            cursor.close()
            conn.close()

    save_button = tk.Button(deadline_window, text="Lưu Deadline", command=lambda: save_deadline(cal.get_date(),f"{hour_combobox.get()}:{minute_combobox.get()}",
                                                                                                dot_combobox.get()), font=("Arial", 12), bg="#4caf50", fg="white")
    save_button.pack(pady=20)

"""Chức năng gửi mail nếu trễ hạn"""
def get_current_month_deadlines():
    conn = sqlite3.connect('students.db')
    cursor = conn.cursor()
    current_month = datetime.now().month
    current_year = datetime.now().year

    query = """
    SELECT c.MonHoc, r.SubmissionDate, r.SubmissionTime 
    FROM report_statuses r
    JOIN courses c ON r.CourseID = c.ID
    WHERE strftime('%Y', r.SubmissionDate) = ? 
    AND strftime('%m', r.SubmissionDate) = ?
    """
    cursor.execute(query, (str(current_year), str(current_month).zfill(2)))
    deadlines = cursor.fetchall()

    cursor.close()
    conn.close()

    return deadlines

def fetch_email_data(username, password):
    mail = imaplib.IMAP4_SSL("imap.gmail.com")
    mail.login(username, password)

    mail.select("inbox")

    status, messages = mail.search(None, 'ALL')
    email_ids = messages[0].split()

    email_data_list = []

    for e_id in email_ids:
        status, msg_data = mail.fetch(e_id, '(RFC822)')
        msg = email.message_from_bytes(msg_data[0][1])

        subject, encoding = decode_header(msg['Subject'])[0]
        if isinstance(subject, bytes):
            subject = subject.decode(encoding if encoding else 'utf-8')

        sender = msg['From']

        date_str = msg['Date']
        date_sent = email.utils.parsedate_to_datetime(date_str)

        email_data_list.append({
            'subject': subject,
            'sender': sender,
            'date_sent': date_sent
        })

    mail.logout()

    return email_data_list

def notify_deadline_reminder(course_name, submission_date, submission_time):
    print(f"Bắt đầu kiểm tra email cho deadline của khóa học {course_name}.")

    username = "nguyenddqui@gmail.com"
    password = "tpng wimy hsxv xkku"

    emails = fetch_email_data(username, password)

    submission_datetime = datetime.combine(submission_date, submission_time)

    current_datetime = datetime.now()

    sender_email = None
    for email_info in emails:
        if f"Report for: {course_name}" in email_info['subject']:
            sender_email = email_info['sender']

    time_difference = submission_datetime - current_datetime

    if sender_email:
        print(f"Có email từ khóa học {course_name}: {sender_email}")
        update_status_in_database(course_name, sender_email)

    else:

        if timedelta(minutes=0) <= time_difference <= timedelta(minutes=3):
            subject = f"Thông báo: Sắp đến deadline cho khóa học {course_name}"
            body = f"Xin chào,\n\nDeadline cho khóa học {course_name} sắp đến. Vui lòng kiểm tra."
            send_email("bichqui1212@gmail.com", subject, body)
            print(f"Đã gửi thông báo sắp đến deadline cho khóa học {course_name}.")

        elif time_difference < timedelta(minutes=3):
            subject = f"Thông báo: Deadline đã qua cho khóa học {course_name}"
            body = f"Xin chào,\n\nDeadline cho khóa học {course_name} đã qua. Vui lòng liên hệ để biết thêm thông tin."
            send_email("bichqui1212@gmail.com", subject, body)
            print(f"Đã gửi thông báo muộn cho khóa học {course_name}.")

def update_status_in_database(course_name, sender_email):
    conn = sqlite3.connect('students.db')
    cursor = conn.cursor()

    try:
        query = """
            UPDATE report_statuses
            SET Status = 'YES', Email = ?
            WHERE CourseID = (SELECT ID FROM courses WHERE MonHoc = ?)
        """
        cursor.execute(query, (sender_email, course_name))
        conn.commit()
        print(f"Đã cập nhật trạng thái cho khóa học {course_name} với email {sender_email}.")
    except Exception as e:
        print(f"Có lỗi xảy ra khi cập nhật trạng thái: {e}")
    finally:
        cursor.close()
        conn.close()


def schedule_email_check():

    deadlines = get_current_month_deadlines()

    if not deadlines:
        schedule.every(2).minutes.do(schedule_email_check)
        return

    for course_name, date_str, time_str in deadlines:
        try:

            submission_date = datetime.strptime(date_str, '%Y-%m-%d').date()
            submission_time = datetime.strptime(time_str, '%H:%M').time()

            deadline_datetime = datetime.combine(submission_date, submission_time)

            before_deadline = deadline_datetime - timedelta(minutes=3)
            after_deadline = deadline_datetime + timedelta(minutes=3)

            schedule.every().day.at(before_deadline.strftime("%H:%M")).do(
                notify_deadline_reminder, course_name, submission_date, submission_time
            )

            schedule.every().day.at(after_deadline.strftime("%H:%M")).do(
                notify_deadline_reminder, course_name, submission_date, submission_time
            )

        except Exception as e:
            print(f"Error processing course '{course_name}': {e}")


def delete_all_data():
    conn = sqlite3.connect('students.db')
    cursor = conn.cursor()
    cursor.execute("DELETE FROM report_statuses")
    conn.commit()
    cursor.close()
    conn.close()

"""Chức năng kiểm tra xem Email được phản hồi hay chưa"""
def get_recent_questions():

    conn = sqlite3.connect('students.db')
    cursor = conn.cursor()
    query = '''
           SELECT *
           FROM questions
           WHERE Status = 'Pending'
       '''
    cursor.execute(query)
    questions = cursor.fetchall()
    cursor.close()
    conn.close()

    return questions


def check_recent_sent_email(username, password, submission_date_time, question_id, student_id):

    mail = imaplib.IMAP4_SSL("imap.gmail.com")
    mail.login(username, password)

    mail.select('"[Gmail]/Sent Mail"')

    next_day = submission_date_time + timedelta(days=1)


    subject_keyword = f"Response for question: {question_id} - {student_id}"
    query = f'SINCE {submission_date_time.strftime("%d-%b-%Y")} BEFORE {next_day.strftime("%d-%b-%Y")} SUBJECT "{subject_keyword}"'

    status, messages = mail.search(None, query)
    if status != 'OK' or not messages[0]:
        mail.logout()
        return False

    email_ids = messages[0].split()
    mail.logout()
    return bool(email_ids)

def update_question_status(question_id, status):
    conn = sqlite3.connect('students.db')
    cursor = conn.cursor()

    query = '''
        UPDATE questions
        SET Status =  ?
        WHERE ID = ?;
    '''
    cursor.execute(query, (status, question_id))
    conn.commit()
    cursor.close()
    conn.close()

def check_receivers_and_update():
    username = "nguyenddqui@gmail.com"
    password = "tpng wimy hsxv xkku"
    questions = get_recent_questions()


    for question in questions:
        question_id = question[0]
        student_id = question[1]
        content = question[2]
        submission_date_time = question[3]

        submission_date_time = datetime.strptime(submission_date_time, "%Y-%m-%d %H:%M:%S")
        current_time = datetime.now()
        time_limit = submission_date_time + timedelta(minutes=6)

        if current_time >= time_limit:

            sent_email = check_recent_sent_email(username, password, submission_date_time, question_id, student_id)

            if sent_email:
                update_question_status(question_id, "Resolved")
            else:
                update_question_status(question_id, "Escalated")
                send_email("ducphu625@gmail.com", f"New question: {question_id} - {student_id}", content)

def run_scheduler():
    while True:
        schedule.run_pending()
        time.sleep(1)

def start_scheduler():
    threading.Thread(target=run_scheduler, daemon=True).start()

connect_db()

root = tk.Tk()
root.title("Quản lý sinh viên")
root.geometry("950x600")
root.configure(bg="#f0f0f0")

details_frame = tk.Frame(root, bg="#e0f7fa", padx=10, pady=10, relief=tk.RIDGE, bd=2)
details_frame.pack(pady=10, fill=tk.X, padx=20)

"""Menu"""
menu = tk.Menu(root)
root.config(menu=menu)

account_menu = tk.Menu(menu, tearoff=0)
menu.add_cascade(label="Tài khoản", menu=account_menu)
account_menu.add_command(label="Đăng nhập", command=login)

manage_menu = tk.Menu(menu, tearoff=0)
menu.add_cascade(label="Quản lý sinh viên", menu=manage_menu)
manage_menu.add_command(label="Thêm sinh viên", command=add_student)
manage_menu.add_command(label="Xóa sinh viên", command=delete_student)
manage_menu.add_command(label="Hiển thị danh sách", command=display_students)
manage_menu.add_command(label="Thêm từ file Excel", command=upload_excel_file)

email_menu = tk.Menu(menu, tearoff=0)
menu.add_cascade(label="Gửi Email", menu=email_menu)
email_menu.add_command(label="Gửi Email cảnh báo cho phụ huynh", command=lambda: send_warning_emails("parents"))
email_menu.add_command(label="Gửi Email cảnh báo cho sinh viên", command=lambda: send_warning_emails("students"))

deadline_menu = tk.Menu(menu, tearoff=0)
menu.add_cascade(label="Đặt Deadline", menu=deadline_menu)
deadline_menu.add_command(label="Đặt thời hạn", command=open_deadline_panel)

chatbot_menu = tk.Menu(menu, tearoff=0)
menu.add_cascade(label="Chatbot", menu=chatbot_menu)
chatbot_menu.add_command(label="Bắt đầu trò chuyện", command=lambda: ChatApplication(root))

"""Tìm kiếm"""
search_frame = tk.Frame(root, bg="#e0f7fa", padx=10, pady=10, relief=tk.RIDGE, bd=2)
search_frame.pack(pady=10, fill=tk.X, padx=20)

tk.Label(search_frame, text="Tra cứu sinh viên (Tên hoặc MSSV):", font=("Arial", 12, "bold"), bg="#e0f7fa").pack(side=tk.LEFT, padx=10)
search_entry = tk.Entry(search_frame, font=("Arial", 12))
search_entry.pack(side=tk.LEFT, padx=10)
search_button = tk.Button(search_frame, text="🔎", font=("Arial", 12), bg="#4caf50", fg="white", command=search_student)
search_button.pack(side=tk.LEFT)

"""Phân loại sắp xếp"""
sort_group_frame = tk.Frame(root, bg="#e0f7fa", padx=10, pady=10, relief=tk.RIDGE, bd=2)
sort_group_frame.pack(pady=10, fill=tk.X, padx=20)

tk.Label(sort_group_frame, text="Phân loại theo lớp:", bg="#e0f7fa", font=("Arial", 12)).pack(side=tk.LEFT, padx=(0, 5))

class_combobox = ttk.Combobox(sort_group_frame, values=get_classes(), state="readonly")
class_combobox.pack(side=tk.LEFT, padx=5)
class_combobox.set("Chọn lớp")

btn_group_by_class = tk.Button(sort_group_frame, text="Hiển thị theo lớp", command=lambda: display_students(group_by='Lop', value=class_combobox.get()))
btn_group_by_class.pack(side=tk.LEFT, padx=5)

tk.Label(sort_group_frame, text="Phân loại theo môn học:", bg="#e0f7fa", font=("Arial", 12)).pack(side=tk.LEFT, padx=(20, 5))

subject_combobox = ttk.Combobox(sort_group_frame, values=get_subjects(), state="readonly")
subject_combobox.pack(side=tk.LEFT, padx=5)
subject_combobox.set("Chọn môn học")

btn_group_by_subject = tk.Button(sort_group_frame, text="Hiển thị theo môn", command=lambda: display_students(group_by='MonHoc', value=subject_combobox.get()))
btn_group_by_subject.pack(side=tk.LEFT, padx=5)

sort_frame = tk.Frame(root, bg="#e0f7fa", padx=10, pady=10, relief=tk.RIDGE, bd=2)
sort_frame.pack(pady=10, fill=tk.X, padx=20)

tk.Label(sort_frame, text="Sắp xếp theo:", bg="#e0f7fa", font=("Arial", 12)).pack(side=tk.LEFT, padx=(0, 5))

btn_sort_by_absences = tk.Button(sort_frame, text="Sắp xếp theo vắng", command=lambda: display_students(order_by='Vang'))
btn_sort_by_absences.pack(side=tk.LEFT, padx=5)

btn_sort_by_name = tk.Button(sort_frame, text="Sắp xếp theo tên", command=lambda: display_students(order_by='HoTen'))
btn_sort_by_name.pack(side=tk.LEFT, padx=5)

"""Hiển thị"""
tree = ttk.Treeview(root, columns=("MSSV", "HoTen", "HanhDong"), show='headings')

tree.heading("MSSV", text="MSSV")
tree.heading("HoTen", text="Họ tên")
tree.heading("HanhDong", text="Hành Động")

column_widths = {
    "MSSV": 100,
    "HoTen": 150,
    "HanhDong": 50
}

for col, width in column_widths.items():
    tree.column(col, width=width, anchor="center")

tree.pack(pady=10, padx=20, fill=tk.BOTH, expand=True)

tree.bind("<Double-1>", show_student_details)

delete_all_data()
start_scheduler()
schedule_email_check()
schedule.every(5).minutes.do(schedule_send_report)
schedule.every(1).minutes.do(check_receivers_and_update)
root.mainloop()



