import re, csv, sqlite3
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime


class VisitorManagementSystem(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Visitor Management System")
        self.geometry("1425x775")
        self.resizable(False, False)
        self.style = ttk.Style()
        self.tk.call("source", "forest-dark.tcl")
        self.style.theme_use("forest-dark")
        self.name = tk.StringVar()
        self.phone = tk.StringVar()
        self.address = tk.StringVar()
        self.visiting = tk.StringVar()
        self.purpose = tk.StringVar()
        self.search_by = tk.StringVar()
        self.search_text = tk.StringVar()
        self.cnx = sqlite3.connect("visitors.db", timeout=10)
        self.cur = self.cnx.cursor()
        self.cur.execute(
            """
                CREATE TABLE IF NOT EXISTS visitor_data (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                phone_number TEXT NOT NULL,
                address TEXT NOT NULL,
                visiting TEXT NOT NULL,
                purpose TEXT NOT NULL,
                arrived TEXT NOT NULL,
                departed TEXT);
                """
        )

        Registration_Frame = tk.Frame(bd=4)
        Registration_Frame.place(x=25, y=20, width=800, height=900)

        m_title = tk.Label(
            Registration_Frame,
            text="Registration",
            font=("Montserrat", 25, "bold"),
            fg="white",
        )
        m_title.grid(
            row=0,
            columnspan=2,
            pady=20,
        )

        lbl_name = tk.Label(
            Registration_Frame, text="Name", font=("Montserrat", 12), fg="white"
        )
        lbl_name.grid(row=1, column=0, pady=10, padx=20, sticky="w")

        txt_name = ttk.Entry(
            Registration_Frame,
            textvariable=self.name,
            font=("Montserrat", 10),
            width=20,
        )
        txt_name.grid(row=1, column=1, pady=20, sticky="w")

        lbl_phone = tk.Label(
            Registration_Frame, text="Phone number", font=("Montserrat", 12), fg="white"
        )
        lbl_phone.grid(row=2, column=0, pady=10, padx=20, sticky="w")

        txt_phone = ttk.Entry(
            Registration_Frame,
            textvariable=self.phone,
            font=("Montserrat", 10),
            width=20,
        )
        txt_phone.grid(row=2, column=1, pady=20, sticky="w")

        lbl_address = tk.Label(
            Registration_Frame, text="Address", font=("Montserrat", 12), fg="white"
        )

        lbl_address.grid(row=3, column=0, pady=10, padx=18, sticky="w")

        global txt_address
        txt_address = tk.Text(
            Registration_Frame,
            width=20,
            height=4,
            relief="solid",
            highlightthickness=2,
            highlightbackground="#595959",
            highlightcolor="#595959",
        )
        txt_address.configure(font=("Montserrat", 10))
        txt_address.bind(
            "<FocusIn>",
            lambda e: txt_address.config(
                highlightbackground="#217346", highlightcolor="#217346"
            ),
        )
        txt_address.bind(
            "<FocusOut>",
            lambda e: txt_address.config(
                highlightbackground="#595959", highlightcolor="#595959"
            ),
        )

        self.address = txt_address.get("1.0", "end-1c")

        txt_address.grid(row=3, column=1, padx=6, pady=20, sticky="w")

        lbl_visitor = tk.Label(
            Registration_Frame, text="Visiting", font=("Montserrat", 12), fg="white"
        )
        lbl_visitor.grid(row=4, column=0, pady=10, padx=20, sticky="w")

        txt_visitor = ttk.Entry(
            Registration_Frame,
            textvariable=self.visiting,
            font=("Montserrat", 10),
            width=20,
        )
        txt_visitor.grid(row=4, column=1, pady=20, sticky="w")

        lbl_purpose = tk.Label(
            Registration_Frame, text="Purpose", font=("Montserrat", 12), fg="white"
        )
        lbl_purpose.grid(row=5, column=0, pady=10, padx=20, sticky="w")

        txt_purpose = ttk.Entry(
            Registration_Frame,
            textvariable=self.purpose,
            font=("Montserrat", 10),
            width=20,
        )
        txt_purpose.grid(row=5, column=1, pady=20, sticky="w")

        btn_frame = tk.Frame(Registration_Frame, bd=4)
        btn_frame.place(x=10, y=600, width=450)
        btn_frame.grid_columnconfigure(0, weight=1)

        self.style.configure("Accent.TButton", font=("Montserrat", 10))
        accentbutton = ttk.Button(
            btn_frame,
            text="Register",
            style="Accent.TButton",
            command=self.register_user,
        )

        accentbutton.grid(row=6, column=0, padx=5, pady=10, sticky="nsew")

        title = tk.Label(
            text="Developed By Adish and Harsh",
            bd=3,
            font=("Montserrat", 12),
            fg="grey",
        )
        title.pack(side=tk.BOTTOM, fill=tk.X)

        self.style.configure(
            "TLabelframe.Label", font=("Montserrat", 10), foreground="#fff"
        )
        Detail_Frame = ttk.LabelFrame(text="Details")
        Detail_Frame.place(x=525, y=80, width=850, height=600)

        lbl_search = tk.Label(
            Detail_Frame, text="Search By", font=("Montserrat", 12), fg="white"
        )
        lbl_search.grid(row=0, column=0, pady=10, padx=20, sticky="w")

        self.style.configure("Accent.TCombobox", font=("Montserrat", 10), height=3)

        combo_search = ttk.Combobox(
            Detail_Frame,
            state="readonly",
            width=14,
            textvariable=self.search_by,
            values=["Name", "Phone number", "Address", "Visiting", "Purpose"],
        )

        combo_search.current(0)
        combo_search.grid(row=0, column=1, padx=20, pady=10)

        txt_search = ttk.Entry(
            Detail_Frame,
            textvariable=self.search_text,
            font=("Montserrat", 10),
            width=12,
        )
        txt_search.grid(row=0, column=2, pady=10, padx=20, sticky="w")

        searchbtn = ttk.Button(
            Detail_Frame,
            text="Search",
            width=6,
            command=self.search_data,
            style="Accent.TButton",
        ).grid(row=0, column=3, padx=10, pady=10)

        self.style.configure("TButton", font=("Montserrat", 10))

        refreshbtn = ttk.Button(
            Detail_Frame,
            text="Refresh",
            width=6,
            command=self.fetch_data,
        ).grid(row=0, column=4, padx=5, pady=10)

        exportbtn = ttk.Button(
            Detail_Frame,
            text="Export",
            width=5,
            command=self.export_table,
        ).grid(row=0, column=5, padx=5, pady=10)

        Table_Frame = tk.Frame(Detail_Frame, bd=4)
        Table_Frame.place(x=10, y=80, width=830, height=480)

        scroll_x = ttk.Scrollbar(Table_Frame, orient=tk.HORIZONTAL)
        scroll_y = ttk.Scrollbar(Table_Frame, orient=tk.VERTICAL)
        self.Visitor_table = ttk.Treeview(
            Table_Frame,
            columns=(
                "id",
                "name",
                "phone_number",
                "address",
                "visiting",
                "purpose",
                "arrived",
                "departed",
            ),
            xscrollcommand=scroll_x.set,
            yscrollcommand=scroll_y.set,
        )

        scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        scroll_x.config(command=self.Visitor_table.xview)
        scroll_y.config(command=self.Visitor_table.yview)
        self.Visitor_table.heading("id", text="Id")
        self.Visitor_table.heading("name", text="Name")
        self.Visitor_table.heading("phone_number", text="Phone")
        self.Visitor_table.heading("address", text="Address")
        self.Visitor_table.heading("visiting", text="Visiting")
        self.Visitor_table.heading("purpose", text="Purpose")
        self.Visitor_table.heading("arrived", text="Arrived")
        self.Visitor_table.heading("departed", text="Departed")
        self.Visitor_table["show"] = "headings"
        self.Visitor_table.column("id", width=30, anchor="center")
        self.Visitor_table.column("name", width=100, anchor="center")
        self.Visitor_table.column("phone_number", width=100, anchor="center")
        self.Visitor_table.column("address", width=100, anchor="center")
        self.Visitor_table.column("visiting", width=80, anchor="center")
        self.Visitor_table.column("purpose", width=80, anchor="center")
        self.Visitor_table.column("arrived", width=100, anchor="center")
        self.Visitor_table.column("departed", width=100, anchor="center")
        self.Visitor_table.pack(fill=tk.BOTH, expand=1)
        self.Visitor_table.bind("<ButtonRelease-1>", self.mark_departed)

        self.fetch_data()

    def register_user(self):
        self.address = txt_address.get("1.0", "end-1c")
        if (
            not self.name.get()
            or not self.phone.get()
            or not self.address
            or not self.visiting.get()
            or not self.purpose.get()
        ):
            messagebox.showerror("Error", "All the fields are required!")
        elif not re.match(r"^\d{10}$", self.phone.get()):
            messagebox.showerror("Error", "Please enter a valid phone number.")

        else:
            arrived = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            self.cur.execute(
                "INSERT INTO visitor_data (name, phone_number, address, visiting, purpose, arrived, departed) values(?,?,?,?,?,?,?)",
                (
                    self.name.get(),
                    self.phone.get(),
                    self.address,
                    self.visiting.get(),
                    self.purpose.get(),
                    arrived,
                    None,
                ),
            )
            self.cnx.commit()
            self.fetch_data()
            messagebox.showinfo("Successful", "Record has been inserted")
            self.clear()

    def fetch_data(self):
        try:
            self.cur.execute("SELECT * FROM visitor_data")
            rows = self.cur.fetchall()
            if rows:
                self.Visitor_table.delete(*self.Visitor_table.get_children())
                for row in rows:
                    self.Visitor_table.insert("", tk.END, values=row)
                self.cnx.commit()
        except Exception as e:
            pass

    def search_data(self):
        search_by = self.search_by.get()

        if not self.search_text.get():
            messagebox.showerror("Error", "Search value is required!")
        else:
            substitution = {
                "Name": "name",
                "Phone number": "phone_number",
                "Address": "address",
                "Visiting": "visiting",
                "Purpose": "purpose",
            }

            if search_by in substitution:
                search_by = substitution[search_by]

            try:
                self.cur.execute(
                    f"SELECT * FROM visitor_data WHERE {search_by} LIKE '%{str(self.search_text.get())}%'"
                )

                rows = self.cur.fetchall()
                if rows:
                    self.Visitor_table.delete(*self.Visitor_table.get_children())
                    for row in rows:
                        self.Visitor_table.insert("", tk.END, values=row)
                    self.cnx.commit()
                else:
                    messagebox.showerror("Not Found", "Record not Found")
            except Exception as e:
                pass

    def clear(self):
        self.name.set("")
        self.phone.set("")
        txt_address.delete("1.0", "end-1c")
        self.visiting.set("")
        self.purpose.set("")

    def mark_departed(self, ev):
        try:
            selected_item = self.Visitor_table.item(self.Visitor_table.focus())
            name = selected_item["values"][1]
            is_departed = selected_item["values"][7]

            if is_departed == "None":
                departed = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                confirm = messagebox.askyesno(
                    "Confirmation", "Mark the selected visitor as departed?"
                )
                if confirm:
                    self.cur.execute(
                        f"UPDATE visitor_data SET departed = '{departed}' WHERE name = '{name}'"
                    )
                    self.cnx.commit()
                    messagebox.showinfo("Successful", "Marked visitor as departed")
                    self.fetch_data()
        except Exception as e:
            pass

    def export_table(self):
        try:
            file_path = filedialog.asksaveasfilename(
                defaultextension=".csv", filetypes=[("CSV File", "*.csv")]
            )
            if file_path:
                with open(file_path, "w", newline="") as csv_file:
                    csv_writer = csv.writer(csv_file)
                    headings = self.Visitor_table["columns"]
                    column_names = [
                        self.Visitor_table.heading(heading)["text"]
                        for heading in headings
                    ]
                    csv_writer.writerow(column_names)

                    for item in self.Visitor_table.get_children():
                        row = self.Visitor_table.item(item, "values")
                        csv_writer.writerow(row)
        except Exception as e:
            pass


if __name__ == "__main__":
    VisitorManagementSystem().mainloop()
