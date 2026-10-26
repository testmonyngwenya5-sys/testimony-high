import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import sqlite3
import hashlib
import datetime
import re
import os

DB_NAME = "testimony_high.db"

def hash_pwd(p): return hashlib.sha256(p.encode()).hexdigest()

def init_db():
    # Auto-fix old DB with 7 columns error
    if os.path.exists(DB_NAME):
        try:
            con=sqlite3.connect(DB_NAME); cur=con.cursor()
            cur.execute("PRAGMA table_info(users)")
            cols=cur.fetchall()
            # If users table has 7 columns or wrong structure, delete DB to recreate fresh
            if cols and len(cols) != 6:
                con.close()
                os.remove(DB_NAME)
                print("Old DB with wrong columns removed - creating new one")
        except:
            pass

    con = sqlite3.connect(DB_NAME)
    cur = con.cursor()
    cur.execute("CREATE TABLE IF NOT EXISTS users (id INTEGER PRIMARY KEY AUTOINCREMENT, username TEXT UNIQUE, password TEXT, full_name TEXT, role TEXT, reg_number TEXT)")
    cur.execute("CREATE TABLE IF NOT EXISTS subjects (code TEXT PRIMARY KEY, name TEXT)")
    cur.execute("CREATE TABLE IF NOT EXISTS grades (id INTEGER PRIMARY KEY AUTOINCREMENT, reg_number TEXT, subject_code TEXT, mark INTEGER, grade TEXT, comment TEXT, term TEXT, teacher TEXT)")
    cur.execute("CREATE TABLE IF NOT EXISTS fees (id INTEGER PRIMARY KEY AUTOINCREMENT, reg_number TEXT, term TEXT, amount REAL, paid REAL, balance REAL)")
    cur.execute("CREATE TABLE IF NOT EXISTS announcements (id INTEGER PRIMARY KEY AUTOINCREMENT, title TEXT, message TEXT, date TEXT)")
    cur.execute("CREATE TABLE IF NOT EXISTS password_requests (id INTEGER PRIMARY KEY AUTOINCREMENT, username TEXT, new_password TEXT, time TEXT, approved INTEGER DEFAULT 0)")
    subs = [("4008","English Language"),("4028","Mathematics"),("4024","Combined Science"),("5008","Biology"),("5054","Chemistry"),("5055","Physics"),("4044","History"),("4022","Geography"),("7100","Commerce"),("4050","Accounting"),("4027","Computer Science"),("3159","Shona")]
    for c,n in subs: cur.execute("INSERT OR IGNORE INTO subjects VALUES (?,?)",(c,n))
    con.commit(); con.close()

def has_users():
    con=sqlite3.connect(DB_NAME); cur=con.cursor(); cur.execute("SELECT COUNT(*) FROM users"); count=cur.fetchone()[0]; con.close(); return count>0

def get_next_staff_id(role):
    year=datetime.datetime.now().year
    prefix_map={"DIRECTOR":f"DIR{year}","HEADMASTER":f"HM{year}","DEPUTY HEAD":f"DHM{year}","ADMIN":f"ADM{year}","BURSAR":f"BUR{year}","TEACHER":f"TCH{year}","STUDENT":f"THS{year}"}
    prefix=prefix_map.get(role, f"THS{year}")
    con=sqlite3.connect(DB_NAME); cur=con.cursor()
    cur.execute("SELECT reg_number FROM users WHERE reg_number LIKE ?", (f"{prefix}%",))
    rows=cur.fetchall(); con.close()
    max_num=0
    for (reg,) in rows:
        if not reg: continue
        m=re.search(rf"{re.escape(prefix)}(\d+)", reg)
        if m:
            try: max_num=max(max_num, int(m.group(1)))
            except: pass
    return f"{prefix}{max_num+1:04d}"

def export_excel(headers, rows, default_name):
    try:
        import openpyxl
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    except: messagebox.showerror("Missing","pip install openpyxl"); return
    wb=openpyxl.Workbook(); ws=wb.active
    ws.merge_cells('A1:G1'); ws['A1']="TESTIMONY HIGH SCHOOL - Official Report"; ws['A1'].font=Font(bold=True,size=14); ws['A1'].alignment=Alignment(horizontal="center")
    for col,h in enumerate(headers,1):
        c=ws.cell(row=4,column=col,value=h); c.font=Font(bold=True,color="FFFFFF"); c.fill=PatternFill(start_color="D97706",end_color="D97706",fill_type="solid"); c.alignment=Alignment(horizontal="center"); c.border=Border(left=Side(style='thin'),right=Side(style='thin'),top=Side(style='thin'),bottom=Side(style='thin'))
    for r_idx,row in enumerate(rows,5):
        for c_idx,val in enumerate(row,1):
            cell=ws.cell(row=r_idx,column=c_idx,value=val); cell.border=Border(left=Side(style='thin'),right=Side(style='thin'),top=Side(style='thin'),bottom=Side(style='thin')); cell.alignment=Alignment(horizontal="center")
    path=filedialog.asksaveasfilename(defaultextension=".xlsx",initialfile=default_name,filetypes=[("Excel","*.xlsx")])
    if path: wb.save(path); messagebox.showinfo("Exported",f"Saved to {path}")

def export_pdf(headers, rows, title):
    try:
        from reportlab.lib import colors
        from reportlab.lib.pagesizes import A4, landscape
        from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
        from reportlab.lib.styles import getSampleStyleSheet
    except: messagebox.showerror("Missing","pip install reportlab"); return
    path=filedialog.asksaveasfilename(defaultextension=".pdf",initialfile=f"{title}.pdf",filetypes=[("PDF","*.pdf")])
    if not path: return
    doc=SimpleDocTemplate(path,pagesize=landscape(A4),topMargin=30)
    styles=getSampleStyleSheet(); story=[Paragraph(f"<b>TESTIMONY HIGH SCHOOL</b> - {title}",styles['Title']), Spacer(1,12)]
    data=[headers]+rows; table=Table(data,repeatRows=1)
    table.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.HexColor("#D97706")),('TEXTCOLOR',(0,0),(-1,0),colors.whitesmoke),('ALIGN',(0,0),(-1,-1),'CENTER'),('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),('GRID',(0,0),(-1,-1),0.5,colors.black)]))
    story.append(table); doc.build(story); messagebox.showinfo("Exported",f"Saved to {path}")

def export_word(headers, rows, title):
    try:
        from docx import Document
        from docx.enum.text import WD_ALIGN_PARAGRAPH
    except: messagebox.showerror("Missing","pip install python-docx"); return
    path=filedialog.asksaveasfilename(defaultextension=".docx",initialfile=f"{title}.docx",filetypes=[("Word","*.docx")])
    if not path: return
    doc=Document(); doc.add_heading('TESTIMONY HIGH SCHOOL',1).alignment=WD_ALIGN_PARAGRAPH.CENTER; doc.add_heading(title,2).alignment=WD_ALIGN_PARAGRAPH.CENTER
    table=doc.add_table(rows=1,cols=len(headers)); table.style='Light Shading Accent 1'
    for i,h in enumerate(headers): table.rows[0].cells[i].text=str(h)
    for row in rows:
        cells=table.add_row().cells
        for i,val in enumerate(row): cells[i].text=str(val)
    doc.save(path); messagebox.showinfo("Exported",f"Saved to {path}")

class SetupWindow:
    def __init__(self, root, on_done):
        self.root=root; self.on_done=on_done
        root.title("First Time Setup"); root.geometry("500x480"); root.configure(bg="#fef3c7")
        tk.Label(root, text="TESTIMONY HIGH SCHOOL", font=("Segoe UI Bold",20), bg="#fef3c7", fg="#92400e").pack(pady=(30,5))
        tk.Label(root, text="First Time - Create ADMIN", font=("Segoe UI Bold",12), bg="#fef3c7").pack()
        card=tk.Frame(root, bg="white"); card.pack(pady=15, padx=30, fill="x")
        tk.Label(card, text="Full Name", bg="white").pack(anchor="w", padx=20, pady=(15,2)); self.fn=tk.Entry(card, font=("Segoe UI",11)); self.fn.pack(fill="x", padx=20, ipady=5)
        auto_id=get_next_staff_id("ADMIN")
        tk.Label(card, text="Staff ID (Auto Generated)", bg="white").pack(anchor="w", padx=20, pady=(10,2))
        self.un_var=tk.StringVar(value=auto_id)
        tk.Entry(card, font=("Segoe UI Bold",11), textvariable=self.un_var, state="readonly").pack(fill="x", padx=20, ipady=5)
        tk.Label(card, text="Password", bg="white").pack(anchor="w", padx=20, pady=(10,2)); self.pw=tk.Entry(card, show="*", font=("Segoe UI",11)); self.pw.pack(fill="x", padx=20, ipady=5)
        tk.Label(card, text="Confirm", bg="white").pack(anchor="w", padx=20, pady=(10,2)); self.cpw=tk.Entry(card, show="*", font=("Segoe UI",11)); self.cpw.pack(fill="x", padx=20, ipady=5)
        tk.Button(card, text="CREATE ADMIN & START", bg="#16a34a", fg="white", font=("Segoe UI Bold",12), bd=0, command=self.create).pack(fill="x", padx=20, pady=20, ipady=8)

    def create(self):
        if not self.fn.get() or not self.un_var.get() or not self.pw.get(): messagebox.showwarning("Missing","Fill all"); return
        if self.pw.get()!=self.cpw.get(): messagebox.showwarning("Error","Passwords don't match"); return
        con=sqlite3.connect(DB_NAME); cur=con.cursor()
        try:
            # FIXED - uses column names
            cur.execute("INSERT INTO users (username, password, full_name, role, reg_number) VALUES (?,?,?,?,?)",
                        (self.un_var.get().strip().upper(), hash_pwd(self.pw.get()), self.fn.get().strip(), "ADMIN", self.un_var.get().strip().upper()))
            con.commit(); con.close(); messagebox.showinfo("Success",f"Admin {self.un_var.get()} created!"); self.root.destroy(); self.on_done()
        except Exception as e: messagebox.showerror("Error",str(e)); con.close()

class LoginWindow:
    def __init__(self, root, on_login):
        self.root=root; self.on_login=on_login
        root.title("Testimony High School - Login"); root.geometry("480x520"); root.configure(bg="#fef3c7")
        tk.Label(root, text="TESTIMONY HIGH SCHOOL", font=("Segoe UI Bold",22), bg="#fef3c7", fg="#92400e").pack(pady=(35,5))
        card=tk.Frame(root, bg="white"); card.pack(pady=10, padx=35, fill="x")
        tk.Label(card, text="Reg No / Staff ID", bg="white").pack(anchor="w", padx=20, pady=(20,5))
        self.user_entry=tk.Entry(card, font=("Segoe UI",12), bd=1, relief="solid"); self.user_entry.pack(fill="x", padx=20, ipady=7)
        tk.Label(card, text="Password", bg="white").pack(anchor="w", padx=20, pady=(15,5))
        row=tk.Frame(card, bg="white"); row.pack(fill="x", padx=20)
        self.pass_entry=tk.Entry(row, font=("Segoe UI",12), show="*", bd=1, relief="solid"); self.pass_entry.pack(side="left", fill="x", expand=True, ipady=7)
        self.show_var=tk.BooleanVar(value=False)
        tk.Checkbutton(row, text="👁", bg="white", variable=self.show_var, command=lambda: self.pass_entry.config(show="" if self.show_var.get() else "*")).pack(side="left", padx=8)
        tk.Button(card, text="LOGIN", bg="#d97706", fg="white", font=("Segoe UI Bold",12), bd=0, cursor="hand2", command=self.do_login).pack(fill="x", padx=20, pady=(15,5), ipady=8)
        tk.Button(card, text="Forgot Admin Password? Click to Recover", bg="white", fg="#2563eb", bd=0, font=("Segoe UI",9,"underline"), cursor="hand2", command=self.recover_admin).pack(pady=(5,20))

    def do_login(self):
        u=self.user_entry.get().strip(); p=self.pass_entry.get().strip()
        if not u or not p: messagebox.showwarning("Missing","Enter ID and Password"); return
        con=sqlite3.connect(DB_NAME); cur=con.cursor()
        cur.execute("SELECT id, username, password, full_name, role, reg_number FROM users WHERE username=?", (u,))
        r=cur.fetchone(); con.close()
        if r and r[2]==hash_pwd(p):
            user={'id':r[0],'username':r[1],'full_name':r[3],'role':r[4],'reg_number':r[5]}
            self.root.destroy(); self.on_login(user)
        else: messagebox.showerror("Failed","Invalid ID or Password")

    def recover_admin(self):
        win=tk.Toplevel(self.root); win.title("Admin Recovery"); win.geometry("400x380"); win.configure(bg="white")
        tk.Label(win, text="ADMIN RECOVERY", font=("Segoe UI Bold",14), bg="white", fg="#dc2626").pack(pady=15)
        tk.Label(win, text="Enter Master Recovery Key:", bg="white").pack(anchor="w", padx=20, pady=(10,2))
        key_e=tk.Entry(win, show="*", font=("Segoe UI",11)); key_e.pack(fill="x", padx=20, ipady=6)
        tk.Label(win, text="Master Key: Testimony@Recover2026", bg="#fef3c7", fg="#92400e", font=("Segoe UI",8)).pack(padx=20, pady=5, fill="x")
        tk.Label(win, text="Select ADMIN to Reset:", bg="white").pack(anchor="w", padx=20, pady=(15,2))
        con=sqlite3.connect(DB_NAME); cur=con.cursor(); cur.execute("SELECT username, full_name FROM users WHERE role='ADMIN'"); admins=cur.fetchall(); con.close()
        admin_var=tk.StringVar()
        ttk.Combobox(win, textvariable=admin_var, values=[f"{a[0]} - {a[1]}" for a in admins], width=35).pack(padx=20, pady=5)
        tk.Label(win, text="New Password:", bg="white").pack(anchor="w", padx=20, pady=(10,2)); np=tk.Entry(win, show="*", font=("Segoe UI",11)); np.pack(fill="x", padx=20, ipady=6)
        tk.Label(win, text="Confirm Password:", bg="white").pack(anchor="w", padx=20, pady=(5,2)); cp=tk.Entry(win, show="*", font=("Segoe UI",11)); cp.pack(fill="x", padx=20, ipady=6)
        def do_reset():
            if key_e.get()!="Testimony@Recover2026": messagebox.showerror("Failed","Wrong Master Key"); return
            if not admin_var.get() or not np.get(): messagebox.showwarning("Missing","Fill all"); return
            if np.get()!=cp.get(): messagebox.showwarning("Error","Passwords don't match"); return
            username=admin_var.get().split(" - ")[0].strip()
            con=sqlite3.connect(DB_NAME); cur=con.cursor(); cur.execute("UPDATE users SET password=? WHERE username=?",(hash_pwd(np.get()), username)); con.commit(); con.close()
            messagebox.showinfo("Recovered",f"Admin {username} password reset!"); win.destroy()
        tk.Button(win, text="RESET ADMIN PASSWORD", bg="#dc2626", fg="white", font=("Segoe UI Bold",11), bd=0, command=do_reset).pack(fill="x", padx=20, pady=20, ipady=6)

class App:
    def __init__(self, root, user):
        self.root=root; self.user=user
        root.title(f"Testimony High School - {user['full_name']}"); root.geometry("1350x780"); root.configure(bg="#f8fafc")
        top=tk.Frame(root, bg="white", height=55); top.pack(fill="x", side="top")
        tk.Label(top, text="🎓 TESTIMONY HIGH SCHOOL", bg="white", fg="#0f172a", font=("Segoe UI Bold",13)).pack(side="left", padx=20)
        tk.Button(top, text="🔄 Switch User", bg="#e2e8f0", bd=0, font=("Segoe UI Bold",10), command=self.switch_user).pack(side="right", padx=10, pady=10)
        tk.Label(top, text=f"👤 {user['full_name']} ({user['role']})", bg="white", fg="#475569").pack(side="right", padx=10)
        self.sidebar=tk.Frame(root, bg="#0f172a", width=250); self.sidebar.pack(side="left", fill="y")
        if user['role']=="DIRECTOR": menus=[("📊 Dashboard & Stats",self.dashboard),("👥 Manage Users",self.manage_users),("👥 Teachers",self.teachers_list),("👥 Students",self.students_list),("📊 All Results",self.all_results),("💰 Financial Summary",self.fin_report),("📢 Announcements",self.announcements),("📚 Subjects",self.subjects)]
        elif user['role']=="HEADMASTER": menus=[("📊 Dashboard & Stats",self.dashboard),("👥 Students",self.students_list),("👥 Teachers",self.teachers_list),("📊 All Results",self.all_results),("📢 Announcements",self.announcements),("📚 Subjects",self.subjects),("💰 Fees Overview",self.fin_report)]
        elif user['role']=="DEPUTY HEAD": menus=[("📊 Dashboard & Stats",self.dashboard),("👥 Students",self.students_list),("👥 Teachers",self.teachers_list),("📊 All Results",self.all_results),("📢 Announcements",self.announcements),("📚 Subjects",self.subjects)]
        elif user['role']=="ADMIN": menus=[("📊 Dashboard",self.dashboard),("👥 Manage Users",self.manage_users),("📚 Subjects",self.subjects),("📢 Announcements",self.announcements),("🔑 Password Requests",self.pwd_req)]
        elif user['role']=="BURSAR": menus=[("📊 Dashboard",self.dashboard),("💰 Fees Management",self.fees_bursar),("📈 Financial Summary",self.fin_report),("👥 Students",self.students_list)]
        elif user['role']=="TEACHER": menus=[("📊 Dashboard",self.dashboard),("📝 Enter Marks",self.enter_marks),("📊 All Results",self.all_results),("👥 Students",self.students_list),("📢 Announcements",self.announcements)]
        else: menus=[("📊 Dashboard",self.dashboard),("📊 My Results",self.my_results),("💰 My Fees",self.my_fees),("📚 Subjects",self.subjects),("📢 Announcements",self.announcements)]
        menus.append(("👤 Profile",self.profile))
        for txt,cmd in menus: tk.Button(self.sidebar, text=txt, bg="#0f172a", fg="#e2e8f0", anchor="w", bd=0, font=("Segoe UI",11), padx=20, pady=12, command=cmd).pack(fill="x")
        tk.Button(self.sidebar, text="🚪 Logout", bg="#dc2626", fg="white", bd=0, command=root.destroy).pack(side="bottom", fill="x", padx=20, pady=5)
        tk.Button(self.sidebar, text="🔄 Switch User", bg="#334155", fg="white", bd=0, command=self.switch_user).pack(side="bottom", fill="x", padx=20, pady=(0,20))
        self.content=tk.Frame(root, bg="#f8fafc"); self.content.pack(side="left", fill="both", expand=True)
        self.dashboard()

    def switch_user(self):
        self.root.destroy(); root=tk.Tk()
        def on_login(user): dash=tk.Tk(); App(dash, user); dash.mainloop()
        LoginWindow(root, on_login); root.mainloop()

    def clear(self):
        for w in self.content.winfo_children(): w.destroy()

    def export_bar(self, headers, get_rows_func, title):
        bar=tk.Frame(self.content, bg="#f8fafc"); bar.pack(fill="x", padx=20, pady=5)
        tk.Label(bar, text="Export:", bg="#f8fafc", font=("Segoe UI Bold",10)).pack(side="left")
        tk.Button(bar, text="📗 Excel", bg="#16a34a", fg="white", bd=0, command=lambda: export_excel(headers, get_rows_func(), title)).pack(side="left", padx=5)
        tk.Button(bar, text="📕 PDF", bg="#dc2626", fg="white", bd=0, command=lambda: export_pdf(headers, get_rows_func(), title)).pack(side="left", padx=5)
        tk.Button(bar, text="📘 Word", bg="#2563eb", fg="white", bd=0, command=lambda: export_word(headers, get_rows_func(), title)).pack(side="left", padx=5)

    def dashboard(self):
        self.clear()
        tk.Label(self.content, text=f"Summary Statistics - {self.user['role']} View", font=("Segoe UI Bold",18), bg="#f8fafc").pack(anchor="w", padx=25, pady=(20,5))
        con=sqlite3.connect(DB_NAME); cur=con.cursor()
        cur.execute("SELECT COUNT(*) FROM users WHERE role='STUDENT'"); s=cur.fetchone()[0]
        cur.execute("SELECT COUNT(*) FROM users WHERE role='TEACHER'"); t=cur.fetchone()[0]
        cur.execute("SELECT COUNT(*) FROM subjects"); sub=cur.fetchone()[0]
        cur.execute("SELECT COUNT(*) FROM grades"); tot=cur.fetchone()[0]
        cur.execute("SELECT COUNT(*) FROM grades WHERE grade IN ('A','B','C')"); passes=cur.fetchone()[0]
        cur.execute("SELECT SUM(amount), SUM(paid), SUM(balance) FROM fees"); fees=cur.fetchone()
        con.close()
        billed=fees[0] or 0; paid=fees[1] or 0; bal=fees[2] or 0
        pass_rate=round(passes/tot*100 if tot else 0,1)
        row1=tk.Frame(self.content, bg="#f8fafc"); row1.pack(fill="x", padx=20)
        for title,val,color in [("👨‍🎓 Students",s,"#d97706"),("👨‍🏫 Teachers",t,"#2563eb"),("📚 ZIMSEC Subjects",sub,"#16a34a"),("📝 Results",tot,"#7c3aed")]:
            f=tk.Frame(row1, bg=color, width=200, height=90); f.pack(side="left", padx=10); f.pack_propagate(False)
            tk.Label(f, text=title, bg=color, fg="white").pack(pady=(15,2)); tk.Label(f, text=str(val), bg=color, fg="white", font=("Segoe UI Bold",22)).pack()
        row2=tk.Frame(self.content, bg="#f8fafc"); row2.pack(fill="x", padx=20, pady=10)
        for title in [f"Billed ${billed}",f"Paid ${paid}",f"Balance ${bal}",f"Pass Rate {pass_rate}%"]:
            f=tk.Frame(row2, bg="white", width=200, height=70, highlightbackground="#e2e8f0", highlightthickness=1); f.pack(side="left", padx=10); f.pack_propagate(False)
            tk.Label(f, text=title, bg="white", font=("Segoe UI Bold",12)).pack(pady=15)
        self.export_bar(["Metric","Value"], lambda: [["Students",s],["Teachers",t],["Subjects",sub],["Results",tot],["Pass Rate",f"{pass_rate}%"],["Billed",billed],["Paid",paid],["Balance",bal]], "Dashboard_Summary")

    def manage_users(self):
        self.clear()
        tk.Label(self.content, text="Manage Users - Auto Reg Number THS + Year + Sequence", font=("Segoe UI Bold",14), bg="#f8fafc").pack(anchor="w", padx=20, pady=10)
        form=tk.Frame(self.content, bg="white", highlightbackground="#e2e8f0", highlightthickness=1); form.pack(fill="x", padx=20, pady=5)
        tk.Label(form, text="Full Name", bg="white").grid(row=0,column=0,padx=8, pady=8, sticky="w")
        fn_entry=tk.Entry(form, width=24, font=("Segoe UI",11)); fn_entry.grid(row=0,column=1,padx=8, pady=8)
        tk.Label(form, text="Role", bg="white").grid(row=0,column=2,padx=8, pady=8, sticky="w")
        role_var=tk.StringVar(value="STUDENT")
        ttk.Combobox(form, textvariable=role_var, values=["STUDENT","TEACHER","BURSAR","ADMIN","DIRECTOR","HEADMASTER","DEPUTY HEAD"], width=16).grid(row=0,column=3,padx=8, pady=8)
        tk.Label(form, text="Reg No (Auto)", bg="white", font=("Segoe UI Bold",10)).grid(row=1,column=0,padx=8, pady=8, sticky="w")
        reg_var=tk.StringVar(value=get_next_staff_id("STUDENT"))
        tk.Entry(form, width=24, textvariable=reg_var, state="readonly", font=("Segoe UI Bold",11), readonlybackground="#fef3c7").grid(row=1,column=1,padx=8, pady=8)
        tk.Label(form, text="Password", bg="white").grid(row=1,column=2,padx=8, pady=8, sticky="w")
        pw_entry=tk.Entry(form, show="*", width=18, font=("Segoe UI",11)); pw_entry.grid(row=1,column=3,padx=8, pady=8)
        btn_frame=tk.Frame(form, bg="white"); btn_frame.grid(row=2,column=0,columnspan=5, pady=10)

        def update_reg(*args): reg_var.set(get_next_staff_id(role_var.get()))
        def refresh_id(): reg_var.set(get_next_staff_id(role_var.get())); messagebox.showinfo("Refreshed", f"Next ID: {reg_var.get()}")

        def create_user():
            full=fn_entry.get().strip(); reg=reg_var.get().strip().upper(); pwd=pw_entry.get().strip(); role=role_var.get()
            if not full or not reg or not pwd: messagebox.showwarning("Missing","Fill Full Name and Password"); return
            con=sqlite3.connect(DB_NAME); cur=con.cursor()
            try:
                # FIXED INSERT WITH COLUMN NAMES - SOLVES 7 columns error
                cur.execute("INSERT INTO users (username, password, full_name, role, reg_number) VALUES (?,?,?,?,?)",(reg, hash_pwd(pwd), full, role, reg))
                con.commit(); con.close()
                messagebox.showinfo("Created", f"✅ Created!\nID: {reg}\nName: {full}\nRole: {role}")
                fn_entry.delete(0, tk.END); pw_entry.delete(0, tk.END); load_users()
            except sqlite3.IntegrityError:
                messagebox.showerror("Exists", f"ID {reg} already exists!"); con.close()
            except Exception as e: messagebox.showerror("Error", str(e)); con.close()

        def delete_user():
            sel=tree.selection()
            if not sel: messagebox.showwarning("Select","Select a user"); return
            uname=tree.item(sel[0])['values'][0]
            if uname==self.user['username']: messagebox.showwarning("Stop","Can't delete yourself!"); return
            if messagebox.askyesno("Confirm", f"Delete {uname}?"):
                con=sqlite3.connect(DB_NAME); cur=con.cursor(); cur.execute("DELETE FROM users WHERE username=?",(uname,)); con.commit(); con.close(); load_users()

        role_var.trace_add("write", update_reg)
        tk.Button(btn_frame, text="➕ CREATE USER", bg="#16a34a", fg="white", font=("Segoe UI Bold",11), bd=0, padx=20, pady=6, cursor="hand2", command=create_user).pack(side="left", padx=8)
        tk.Button(btn_frame, text="🔄 Refresh ID", bg="#2563eb", fg="white", font=("Segoe UI Bold",10), bd=0, padx=15, pady=6, cursor="hand2", command=refresh_id).pack(side="left", padx=8)
        tk.Button(btn_frame, text="🗑️ Delete Selected", bg="#dc2626", fg="white", font=("Segoe UI Bold",10), bd=0, padx=15, pady=6, cursor="hand2", command=delete_user).pack(side="left", padx=8)

        tree=ttk.Treeview(self.content, columns=("id","name","role"), show="headings", height=18)
        for c,w in [("id",150),("name",300),("role",150)]: tree.heading(c, text=c.upper()); tree.column(c, width=w)
        def load_users():
            for i in tree.get_children(): tree.delete(i)
            con=sqlite3.connect(DB_NAME); cur=con.cursor(); cur.execute("SELECT username, full_name, role FROM users ORDER BY id DESC")
            for r in cur.fetchall(): tree.insert("","end",values=r)
            con.close(); reg_var.set(get_next_staff_id(role_var.get()))
        self.export_bar(["Username","Full Name","Role"], lambda: [tree.item(i)['values'] for i in tree.get_children()], "Users_List")
        tree.pack(fill="both", expand=True, padx=20, pady=10); load_users()

    def enter_marks(self):
        if self.user['role']!="TEACHER": messagebox.showwarning("Denied","Teachers only"); return
        self.clear(); tk.Label(self.content, text="Enter Marks", font=("Segoe UI Bold",14), bg="#f8fafc").pack(anchor="w", padx=20, pady=10)
        form=tk.Frame(self.content, bg="white"); form.pack(fill="x", padx=20)
        tk.Label(form, text="Student Reg", bg="white").grid(row=0,column=0,padx=5); reg_e=tk.Entry(form); reg_e.grid(row=0,column=1,padx=5)
        con=sqlite3.connect(DB_NAME); cur=con.cursor(); cur.execute("SELECT code || ' - ' || name FROM subjects"); subs=[r[0] for r in cur.fetchall()]; con.close()
        tk.Label(form, text="Subject", bg="white").grid(row=0,column=2,padx=5); sub_var=tk.StringVar(); ttk.Combobox(form, textvariable=sub_var, values=subs, width=28).grid(row=0,column=3,padx=5)
        tk.Label(form, text="Mark", bg="white").grid(row=1,column=0,padx=5); mark_e=tk.Entry(form, width=10); mark_e.grid(row=1,column=1,padx=5)
        tk.Label(form, text="Term", bg="white").grid(row=1,column=2,padx=5); term_var=tk.StringVar(value="Term 1 2026"); ttk.Combobox(form, textvariable=term_var, values=["Term 1 2026","Term 2 2026","Term 3 2026"], width=12).grid(row=1,column=3,padx=5)
        tk.Label(form, text="Comment", bg="white").grid(row=2,column=0,padx=5); comm_e=tk.Entry(form, width=45); comm_e.grid(row=2,column=1,columnspan=3,padx=5,pady=5,sticky="w")
        def save():
            try:
                code=sub_var.get().split(" - ")[0]; mk=int(mark_e.get()); grade="A" if mk>=80 else "B" if mk>=70 else "C" if mk>=60 else "D" if mk>=50 else "E" if mk>=40 else "U"
                con=sqlite3.connect(DB_NAME); cur=con.cursor()
                cur.execute("SELECT username FROM users WHERE username=? AND role='STUDENT'",(reg_e.get().strip().upper(),))
                if not cur.fetchone(): messagebox.showerror("Error","Student not found"); con.close(); return
                cur.execute("INSERT INTO grades (reg_number, subject_code, mark, grade, comment, term, teacher) VALUES (?,?,?,?,?,?,?)",(reg_e.get().strip().upper(), code, mk, grade, comm_e.get(), term_var.get(), self.user['full_name'])); con.commit(); con.close(); messagebox.showinfo("Saved",f"Saved {mk} {grade}"); self.all_results()
            except Exception as e: messagebox.showerror("Error",str(e))
        tk.Button(form, text="Save Result", bg="#d97706", fg="white", bd=0, command=save).grid(row=3,column=0,columnspan=4, pady=10, ipadx=20)

    def all_results(self):
        self.clear(); tk.Label(self.content, text="All Results", font=("Segoe UI Bold",14), bg="#f8fafc").pack(anchor="w", padx=20, pady=10)
        headers=["Reg No","Subject","Mark","Grade","Term","Comment","Teacher"]
        def get_rows(): con=sqlite3.connect(DB_NAME); cur=con.cursor(); cur.execute("SELECT reg_number, subject_code, mark, grade, term, comment, teacher FROM grades ORDER BY reg_number"); rows=cur.fetchall(); con.close(); return rows
        self.export_bar(headers, get_rows, "All_Results_ZIMSEC")
        tree=ttk.Treeview(self.content, columns=headers, show="headings")
        for c in headers: tree.heading(c, text=c); tree.pack(fill="both", expand=True, padx=20)
        for r in get_rows(): tree.insert("","end",values=r)

    def my_results(self):
        self.clear(); tk.Label(self.content, text=f"My Results - {self.user['reg_number']}", font=("Segoe UI Bold",14), bg="#f8fafc").pack(anchor="w", padx=20, pady=10)
        headers=["Subject","Mark","Grade","Term","Comment","Teacher"]
        def get_rows(): con=sqlite3.connect(DB_NAME); cur=con.cursor(); cur.execute("SELECT subject_code, mark, grade, term, comment, teacher FROM grades WHERE reg_number=?",(self.user['reg_number'],)); rows=cur.fetchall(); con.close(); return rows
        self.export_bar(headers, get_rows, f"Results_{self.user['reg_number']}")
        tree=ttk.Treeview(self.content, columns=headers, show="headings")
        for c in headers: tree.heading(c, text=c); tree.pack(fill="both", expand=True, padx=20)
        for r in get_rows(): tree.insert("","end",values=r)

    def fees_bursar(self):
        self.clear(); tk.Label(self.content, text="Bursar - Fees", font=("Segoe UI Bold",14), bg="#f8fafc").pack(anchor="w", padx=20, pady=10)
        form=tk.Frame(self.content, bg="white"); form.pack(fill="x", padx=20)
        tk.Label(form, text="Student Reg", bg="white").grid(row=0,column=0,padx=5); reg_e=tk.Entry(form); reg_e.grid(row=0,column=1,padx=5)
        tk.Label(form, text="Term", bg="white").grid(row=0,column=2,padx=5); term_e=tk.Entry(form); term_e.insert(0,"Term 1 2026"); term_e.grid(row=0,column=3,padx=5)
        tk.Label(form, text="Amount", bg="white").grid(row=1,column=0,padx=5); amt_e=tk.Entry(form); amt_e.grid(row=1,column=1,padx=5)
        tk.Label(form, text="Paid", bg="white").grid(row=1,column=2,padx=5); paid_e=tk.Entry(form); paid_e.grid(row=1,column=3,padx=5)
        headers=["Reg No","Term","Amount","Paid","Balance"]
        def get_rows(): con=sqlite3.connect(DB_NAME); cur=con.cursor(); cur.execute("SELECT reg_number, term, amount, paid, balance FROM fees"); rows=cur.fetchall(); con.close(); return rows
        def save():
            try: amt=float(amt_e.get()); paid=float(paid_e.get()); bal=amt-paid; con=sqlite3.connect(DB_NAME); cur=con.cursor(); cur.execute("INSERT INTO fees (reg_number, term, amount, paid, balance) VALUES (?,?,?,?,?)",(reg_e.get().strip().upper(), term_e.get(), amt, paid, bal)); con.commit(); con.close(); load()
            except Exception as e: messagebox.showerror("Error",str(e))
        def load():
            for i in tree.get_children(): tree.delete(i)
            for r in get_rows(): tree.insert("","end",values=r)
        tk.Button(form, text="Save", bg="#16a34a", fg="white", bd=0, command=save).grid(row=1,column=4,padx=10)
        self.export_bar(headers, get_rows, "Fees_Report")
        tree=ttk.Treeview(self.content, columns=headers, show="headings")
        for c in headers: tree.heading(c, text=c); tree.pack(fill="both", expand=True, padx=20, pady=10); load()

    def my_fees(self):
        self.clear(); tk.Label(self.content, text="My Fees", font=("Segoe UI Bold",14), bg="#f8fafc").pack(anchor="w", padx=20, pady=10)
        headers=["Term","Amount","Paid","Balance"]
        def get_rows(): con=sqlite3.connect(DB_NAME); cur=con.cursor(); cur.execute("SELECT term, amount, paid, balance FROM fees WHERE reg_number=?",(self.user['reg_number'],)); rows=cur.fetchall(); con.close(); return rows
        self.export_bar(headers, get_rows, f"Fees_{self.user['reg_number']}")
        tree=ttk.Treeview(self.content, columns=headers, show="headings")
        for c in headers: tree.heading(c, text=c); tree.pack(fill="both", expand=True, padx=20)
        for r in get_rows(): tree.insert("","end",values=r)

    def fin_report(self):
        self.clear(); con=sqlite3.connect(DB_NAME); cur=con.cursor(); cur.execute("SELECT SUM(amount), SUM(paid), SUM(balance) FROM fees"); row=cur.fetchone(); con.close()
        tk.Label(self.content, text="Financial Summary", font=("Segoe UI Bold",18), bg="#f8fafc").pack(anchor="w", padx=20, pady=20)
        tk.Label(self.content, text=f"Billed: ${row[0] or 0} | Paid: ${row[1] or 0} | Balance: ${row[2] or 0}", font=("Segoe UI",14), bg="#f8fafc").pack(anchor="w", padx=20)
        self.export_bar(["Type","Amount"], lambda: [["Billed",row[0] or 0],["Paid",row[1] or 0],["Balance",row[2] or 0]], "Financial_Summary")

    def subjects(self):
        self.clear(); tk.Label(self.content, text="ZIMSEC Subjects", font=("Segoe UI Bold",14), bg="#f8fafc").pack(anchor="w", padx=20, pady=10)
        headers=["Code","Subject Name"]
        def get_rows(): con=sqlite3.connect(DB_NAME); cur=con.cursor(); cur.execute("SELECT code, name FROM subjects"); rows=cur.fetchall(); con.close(); return rows
        self.export_bar(headers, get_rows, "ZIMSEC_Subjects")
        tree=ttk.Treeview(self.content, columns=headers, show="headings")
        for c in headers: tree.heading(c, text=c); tree.pack(fill="both", expand=True, padx=20)
        for r in get_rows(): tree.insert("","end",values=r)

    def students_list(self):
        self.clear(); tk.Label(self.content, text="Students", font=("Segoe UI Bold",14), bg="#f8fafc").pack(anchor="w", padx=20, pady=10)
        headers=["Reg No","Full Name"]
        def get_rows(): con=sqlite3.connect(DB_NAME); cur=con.cursor(); cur.execute("SELECT reg_number, full_name FROM users WHERE role='STUDENT'"); rows=cur.fetchall(); con.close(); return rows
        self.export_bar(headers, get_rows, "Students_List")
        tree=ttk.Treeview(self.content, columns=headers, show="headings")
        for c in headers: tree.heading(c, text=c); tree.pack(fill="both", expand=True, padx=20)
        for r in get_rows(): tree.insert("","end",values=r)

    def teachers_list(self):
        self.clear(); tk.Label(self.content, text="Teachers / Staff", font=("Segoe UI Bold",14), bg="#f8fafc").pack(anchor="w", padx=20, pady=10)
        headers=["Staff ID","Full Name","Role"]
        def get_rows(): con=sqlite3.connect(DB_NAME); cur=con.cursor(); cur.execute("SELECT reg_number, full_name, role FROM users WHERE role!='STUDENT'"); rows=cur.fetchall(); con.close(); return rows
        self.export_bar(headers, get_rows, "Staff_List")
        tree=ttk.Treeview(self.content, columns=headers, show="headings")
        for c in headers: tree.heading(c, text=c); tree.pack(fill="both", expand=True, padx=20)
        for r in get_rows(): tree.insert("","end",values=r)

    def announcements(self):
        self.clear(); tk.Label(self.content, text="Announcements", font=("Segoe UI Bold",14), bg="#f8fafc").pack(anchor="w", padx=20, pady=10)
        tree=ttk.Treeview(self.content, columns=["title","msg","date"], show="headings")
        for c in ["title","msg","date"]: tree.heading(c, text=c.upper()); tree.pack(fill="both", expand=True, padx=20)
        con=sqlite3.connect(DB_NAME); cur=con.cursor(); cur.execute("SELECT title, message, date FROM announcements")
        for r in cur.fetchall(): tree.insert("","end",values=r); con.close()
        if self.user['role'] in ["ADMIN","DIRECTOR","HEADMASTER","DEPUTY HEAD"]:
            form=tk.Frame(self.content, bg="white"); form.pack(fill="x", padx=20, pady=5)
            t1=tk.Entry(form, width=18); t1.grid(row=0,column=0,padx=5); t2=tk.Entry(form, width=35); t2.grid(row=0,column=1,padx=5)
            def post(): con=sqlite3.connect(DB_NAME); cur=con.cursor(); cur.execute("INSERT INTO announcements (title, message, date) VALUES (?,?,?)",(t1.get(), t2.get(), datetime.datetime.now().date().isoformat())); con.commit(); con.close(); self.announcements()
            tk.Button(form, text="Post", bg="#d97706", fg="white", bd=0, command=post).grid(row=0,column=2,padx=10)

    def pwd_req(self):
        self.clear(); tk.Label(self.content, text="Password Requests", font=("Segoe UI Bold",14), bg="#f8fafc").pack(anchor="w", padx=20, pady=10)
        tree=ttk.Treeview(self.content, columns=["id","user","time"], show="headings")
        for c in ["id","user","time"]: tree.heading(c, text=c.upper()); tree.pack(fill="both", expand=True, padx=20)
        def load():
            for i in tree.get_children(): tree.delete(i)
            con=sqlite3.connect(DB_NAME); cur=con.cursor(); cur.execute("SELECT id, username, time FROM password_requests WHERE approved=0")
            for r in cur.fetchall(): tree.insert("","end",values=r); con.close()
        def approve():
            sel=tree.selection()
            if not sel: return
            rid, uname, _ = tree.item(sel[0])['values']
            con=sqlite3.connect(DB_NAME); cur=con.cursor(); cur.execute("SELECT new_password FROM password_requests WHERE id=?",(rid,)); np=cur.fetchone()[0]
            cur.execute("UPDATE users SET password=? WHERE username=?",(np, uname)); cur.execute("UPDATE password_requests SET approved=1 WHERE id=?",(rid,)); con.commit(); con.close(); load()
        tk.Button(self.content, text="Approve", bg="#16a34a", fg="white", bd=0, command=approve).pack(pady=10); load()

    def profile(self):
        self.clear(); tk.Label(self.content, text="Profile", font=("Segoe UI Bold",14), bg="#f8fafc").pack(anchor="w", padx=20, pady=10)
        card=tk.Frame(self.content, bg="white"); card.pack(padx=20, pady=20, fill="x")
        tk.Label(card, text=f"{self.user['full_name']}\n{self.user['reg_number']} - {self.user['role']}", bg="white", font=("Segoe UI",11)).pack(padx=20, pady=10, anchor="w")
        tk.Label(card, text="New Password", bg="white").pack(anchor="w", padx=20); np=tk.Entry(card, show="*"); np.pack(fill="x", padx=20, pady=5)
        tk.Label(card, text="Confirm", bg="white").pack(anchor="w", padx=20); cp=tk.Entry(card, show="*"); cp.pack(fill="x", padx=20, pady=5)
        def req():
            if np.get()!=cp.get() or not np.get(): messagebox.showwarning("Error","Mismatch"); return
            con=sqlite3.connect(DB_NAME); cur=con.cursor(); cur.execute("INSERT INTO password_requests (username, new_password, time, approved) VALUES (?,?,?,0)",(self.user['username'], hash_pwd(np.get()), datetime.datetime.now().isoformat())); con.commit(); con.close(); messagebox.showinfo("Sent","Request sent to Admin")
        tk.Button(card, text="Request Change", bg="#2563eb", fg="white", bd=0, command=req).pack(padx=20, pady=15, fill="x")

def main():
    init_db()
    if not has_users():
        root=tk.Tk()
        def on_setup_done():
            root2=tk.Tk()
            def on_login(user): dash=tk.Tk(); App(dash, user); dash.mainloop()
            LoginWindow(root2, on_login); root2.mainloop()
        SetupWindow(root, on_setup_done); root.mainloop()
    else:
        root=tk.Tk()
        def on_login(user): dash=tk.Tk(); App(dash, user); dash.mainloop()
        LoginWindow(root, on_login); root.mainloop()

if __name__=="__main__": main()
