import os
import datetime
import pickle
from tkinter import *
from tkinter import ttk
from tkinter import simpledialog, messagebox
import webbrowser
import subprocess

class IMS:
    def __init__(self, root):
        self.root = root
        self.root.geometry("1350x700+0+0")
        self.root.title("OptionMetrics Inventory | Developed by Aditya Das")
        self.root.config(bg="white")

        self.data_file = "users_data.pkl"
        self.monitor_data_file = "monitor_data.pkl"
        self.pc_data_file = "pc_data.pkl"
        self.users_data = self.load_users_data()
        self.monitor_data = self.load_device_data(self.monitor_data_file)
        self.pc_data = self.load_device_data(self.pc_data_file)

        self.login_screen()

    def load_users_data(self):
        if os.path.exists(self.data_file):
            with open(self.data_file, "rb") as file:
                return pickle.load(file)
        return [
            (1, "OPS", "adas", "adas", "L-OPS-ADAS")
        ]

    def save_users_data(self, users_data):
        with open(self.data_file, "wb") as file:
            pickle.dump(users_data, file)

    def load_device_data(self, filename):
        if os.path.exists(filename):
            with open(filename, "rb") as file:
                return pickle.load(file)
        return []

    def save_device_data(self, filename, data):
        with open(filename, "wb") as file:
            pickle.dump(data, file)

    def login_screen(self):
        self.clear_screen()
        self.root.geometry("400x300+500+200")

        Label(self.root, text="Login", font=("times new roman", 30, "bold"), bg="#00165a", fg="#fac900").pack(side=TOP, fill=X)

        frame_login = Frame(self.root, bd=2, relief=RIDGE, bg="white")
        frame_login.place(x=50, y=100, width=300, height=150)

        Label(frame_login, text="Username", font=("sans mono", 15, "bold"), bg="white").place(x=20, y=20)
        self.txt_user = Entry(frame_login, font=("sans mono", 15), bg="lightgray")
        self.txt_user.place(x=20, y=50, width=250)

        Label(frame_login, text="Password", font=("sans mono", 15, "bold"), bg="white").place(x=20, y=80)
        self.txt_pass = Entry(frame_login, font=("sans mono", 15), bg="lightgray", show="*")
        self.txt_pass.place(x=20, y=110, width=250)

        self.root.bind('<Return>', self.check_login)

        btn_login = Button(self.root, text="Login", font=("sans mono", 15, "bold"), bg="#fac900", fg="#00165a", cursor="hand2", command=self.check_login)
        btn_login.place(x=150, y=250, width=100)

    def check_login(self, event=None):
        username = self.txt_user.get()
        password = self.txt_pass.get()

        if username == "adas" and password == "adas":
            self.main_screen()
        else:
            Label(self.root, text="Invalid Username/Password", font=("sans mono", 15, "bold"), bg="white", fg="red").place(x=80, y=200)

    def main_screen(self):
        self.clear_screen()
        self.root.geometry("1350x700+0+0")
        self.root.config(bg="white")

        self.icon_title = PhotoImage(file="Black Logo.png")
        title = Label(self.root, text="Option Metrics Inventory", image=self.icon_title, compound=LEFT, font=("times new roman", 40, "bold"), bg="#00165a", fg="#fac900", anchor="w", padx=20)
        title.place(x=0, y=0, relwidth=1, height=70)

        btn_logout = Button(self.root, text="Logout", font=("sans mono", 15, "bold"), bg="#fac900", fg="#00165a", cursor="hand2", command=self.logout)
        btn_logout.place(x=1150, y=10, height=50, width=150)

        self.lbl_clock = Label(self.root, text="", font=("sans mono", 15), bg="#fac900", fg="#00165a")
        self.lbl_clock.place(x=0, y=70, relwidth=1, height=30)
        self.update_clock()

        LeftMenu = Frame(self.root, bd=2, relief=RIDGE, bg="white")
        LeftMenu.place(x=0, y=102, width=200, height=565)

        def open_dropbox_link():
            webbrowser.open("https://www.dropbox.com/scl/fi/j0ah8j3tyisi10u5n7wqn/OM-Inventory-2024.xlsx?rlkey=1907d62969luja128t7br3nul&st=x46wnqlu&dl=0")

        lbl_menu = Label(LeftMenu, text="Menu", font=("times new roman", 20, "bold"), bg="#00165a", fg="white").pack(side=TOP, fill=X)

        lbl_users = Button(LeftMenu, text="Users", compound=LEFT, padx=5, anchor="w", font=("sans mono", 20, "bold"), bg="white", bd=3, cursor="hand2", command=self.users_screen).pack(side=TOP, fill=X)
        lbl_serialnumbers = Button(LeftMenu, text="SN", compound=LEFT, padx=5, anchor="w", font=("sans mono", 20, "bold"), bg="white", bd=3, cursor="hand2", command=self.show_sn_table).pack(side=TOP, fill=X)
        lbl_hostnamewarranty = Button(LeftMenu, text="Hostname", compound=LEFT, padx=5, anchor="w", font=("sans mono", 20, "bold"), bg="white", bd=3, cursor="hand2", command=self.show_hostname_table).pack(side=TOP, fill=X)
        lbl_specs = Button(LeftMenu, text="Specs", compound=LEFT, padx=5, anchor="w", font=("sans mono", 20, "bold"), bg="white", bd=3, cursor="hand2", command=self.show_specs_table).pack(side=TOP, fill=X)
        lbl_excel = Button(LeftMenu, text="Excel Sheet", compound=LEFT, padx=5, anchor="w", font=("sans mono", 20, "bold"), bg="white", bd=3, cursor="hand2", command=open_dropbox_link).pack(side=TOP, fill=X)
        lbl_request = Button(LeftMenu, text="Request", compound=LEFT, padx=5, anchor="w", font=("sans mono", 20, "bold"), bg="white", bd=3, cursor="hand2", command=self.create_outlook_draft).pack(side=TOP, fill=X)

        dashboard_frame = Frame(self.root, bg="white")
        dashboard_frame.place(x=200, y=102, width=1150, height=565)

        total_users = len(self.users_data)

        Label(dashboard_frame, text="Total Users", font=("sans mono", 20, "bold"), bg="#00165a", fg="#fac900").place(x=50, y=50)
        Label(dashboard_frame, text=total_users, font=("sans mono", 30, "bold"), bg="#00165a", fg="white").place(x=130, y=120, anchor="center")

        self.icon_monitor = PhotoImage(file="monitor picture.png")
        self.icon_pc = PhotoImage(file="case picture.png")

        lbl_monitor = Label(dashboard_frame, text="Monitors", font=("sans mono", 20, "bold"), bg="#00165a", fg="#fac900")
        lbl_monitor.place(x=400, y=50)
        lbl_monitor_image = Label(dashboard_frame, image=self.icon_monitor, bg="#00165a", cursor="hand2")
        lbl_monitor_image.place(x=425, y=100)
        lbl_monitor_image.bind("<Button-1>", self.show_monitor_page)

        lbl_pc = Label(dashboard_frame, text="PCs", font=("sans mono", 20, "bold"), bg="#00165a", fg="#fac900")
        lbl_pc.place(x=750, y=50)
        lbl_pc_image = Label(dashboard_frame, image=self.icon_pc, bg="#00165a", cursor="hand2")
        lbl_pc_image.place(x=750, y=100)
        lbl_pc_image.bind("<Button-1>", self.show_pc_page)

    def logout(self):
        self.login_screen()

    def clear_screen(self):
        for widget in self.root.winfo_children():
            widget.destroy()

    def update_clock(self):
        now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.lbl_clock.config(text=f"Welcome to OptionMetrics Inventory\t\t Date: {now.split()[0]}\t\t Time: {now.split()[1]}")
        self.root.after(1000, self.update_clock)

    def users_screen(self):
        self.clear_screen()
        self.root.geometry("1350x700+0+0")
        self.root.config(bg="white")

        frame_users = Frame(self.root, bd=2, relief=RIDGE, bg="white")
        frame_users.place(x=0, y=102, width=1350, height=595)

        lbl_users = Label(frame_users, text="Users List", font=("sans mono", 20, "bold"), bg="#00165a", fg="#fac900")
        lbl_users.pack(side=TOP, fill=X)

        columns = ("ID", "Team", "Name", "Email", "Host")
        self.users_tree = ttk.Treeview(frame_users, columns=columns, show="headings")

        for col in columns:
            self.users_tree.heading(col, text=col)
            self.users_tree.column(col, width=100)

        for user in self.users_data:
            self.users_tree.insert("", END, values=user)

        self.users_tree.pack(fill=BOTH, expand=True)

        btn_add_user = Button(frame_users, text="Add User", font=("sans mono", 15, "bold"), bg="#fac900", fg="#00165a", cursor="hand2", command=self.add_user)
        btn_add_user.pack(side=LEFT, padx=10, pady=10)

        btn_edit_user = Button(frame_users, text="Edit User", font=("sans mono", 15, "bold"), bg="#fac900", fg="#00165a", cursor="hand2", command=self.edit_user)
        btn_edit_user.pack(side=LEFT, padx=10, pady=10)

        btn_remove_user = Button(frame_users, text="Remove User", font=("sans mono", 15, "bold"), bg="#fac900", fg="#00165a", cursor="hand2", command=self.remove_user)
        btn_remove_user.pack(side=LEFT, padx=10, pady=10)

        btn_back = Button(frame_users, text="Back", font=("sans mono", 15, "bold"), bg="#fac900", fg="#00165a", cursor="hand2", command=self.main_screen)
        btn_back.pack(side=LEFT, padx=10, pady=10)

    def add_user(self):
        new_id = len(self.users_data) + 1
        new_team = simpledialog.askstring("Input", "Enter Team:")
        new_name = simpledialog.askstring("Input", "Enter Name:")
        new_email = simpledialog.askstring("Input", "Enter Email:")
        new_host = simpledialog.askstring("Input", "Enter Host:")

        if new_team and new_name and new_email and new_host:
            new_user = (new_id, new_team, new_name, new_email, new_host)
            self.users_data.append(new_user)
            self.save_users_data(self.users_data)
            self.users_tree.insert("", END, values=new_user)

    def edit_user(self):
        selected_item = self.users_tree.selection()
        if selected_item:
            user_values = self.users_tree.item(selected_item)["values"]
            new_team = simpledialog.askstring("Input", "Enter Team:", initialvalue=user_values[1])
            new_name = simpledialog.askstring("Input", "Enter Name:", initialvalue=user_values[2])
            new_email = simpledialog.askstring("Input", "Enter Email:", initialvalue=user_values[3])
            new_host = simpledialog.askstring("Input", "Enter Host:", initialvalue=user_values[4])

            if new_team and new_name and new_email and new_host:
                updated_user = (user_values[0], new_team, new_name, new_email, new_host)
                self.users_data[user_values[0] - 1] = updated_user
                self.save_users_data(self.users_data)
                self.users_tree.item(selected_item, values=updated_user)

    def remove_user(self):
        selected_item = self.users_tree.selection()
        if selected_item:
            user_values = self.users_tree.item(selected_item)["values"]
            del self.users_data[user_values[0] - 1]
            for i in range(len(self.users_data)):
                self.users_data[i] = (i + 1, *self.users_data[i][1:])
            self.save_users_data(self.users_data)
            self.users_tree.delete(selected_item)

    def show_sn_table(self):
        self.device_screen("Serial Numbers", self.monitor_data, "Serial Number")

    def show_hostname_table(self):
        self.device_screen("Hostnames", self.pc_data, "Hostname")

    def show_specs_table(self):
        self.device_screen("Specifications", self.pc_data, "Specs")

    def device_screen(self, title, data, column_name):
        self.clear_screen()
        self.root.geometry("1350x700+0+0")
        self.root.config(bg="white")

        frame_devices = Frame(self.root, bd=2, relief=RIDGE, bg="white")
        frame_devices.place(x=0, y=102, width=1350, height=595)

        lbl_devices = Label(frame_devices, text=title, font=("sans mono", 20, "bold"), bg="#00165a", fg="#fac900")
        lbl_devices.pack(side=TOP, fill=X)

        columns = ("ID", column_name)
        self.devices_tree = ttk.Treeview(frame_devices, columns=columns, show="headings")

        for col in columns:
            self.devices_tree.heading(col, text=col)
            self.devices_tree.column(col, width=100)

        for i, device in enumerate(data):
            self.devices_tree.insert("", END, values=(i + 1, device))

        self.devices_tree.pack(fill=BOTH, expand=True)

        btn_add_device = Button(frame_devices, text=f"Add {column_name}", font=("sans mono", 15, "bold"), bg="#fac900", fg="#00165a", cursor="hand2", command=lambda: self.add_device(data, column_name))
        btn_add_device.pack(side=LEFT, padx=10, pady=10)

        btn_edit_device = Button(frame_devices, text=f"Edit {column_name}", font=("sans mono", 15, "bold"), bg="#fac900", fg="#00165a", cursor="hand2", command=lambda: self.edit_device(data, column_name))
        btn_edit_device.pack(side=LEFT, padx=10, pady=10)

        btn_remove_device = Button(frame_devices, text=f"Remove {column_name}", font=("sans mono", 15, "bold"), bg="#fac900", fg="#00165a", cursor="hand2", command=lambda: self.remove_device(data, column_name))
        btn_remove_device.pack(side=LEFT, padx=10, pady=10)

        btn_back = Button(frame_devices, text="Back", font=("sans mono", 15, "bold"), bg="#fac900", fg="#00165a", cursor="hand2", command=self.main_screen)
        btn_back.pack(side=LEFT, padx=10, pady=10)

    def add_device(self, data, column_name):
        new_device = simpledialog.askstring("Input", f"Enter {column_name}:")

        if new_device:
            data.append(new_device)
            self.save_device_data(self.monitor_data_file if column_name == "Serial Number" else self.pc_data_file, data)
            self.devices_tree.insert("", END, values=(len(data), new_device))

    def edit_device(self, data, column_name):
        selected_item = self.devices_tree.selection()
        if selected_item:
            device_values = self.devices_tree.item(selected_item)["values"]
            new_device = simpledialog.askstring("Input", f"Enter {column_name}:", initialvalue=device_values[1])

            if new_device:
                data[device_values[0] - 1] = new_device
                self.save_device_data(self.monitor_data_file if column_name == "Serial Number" else self.pc_data_file, data)
                self.devices_tree.item(selected_item, values=(device_values[0], new_device))

    def remove_device(self, data, column_name):
        selected_item = self.devices_tree.selection()
        if selected_item:
            device_values = self.devices_tree.item(selected_item)["values"]
            del data[device_values[0] - 1]
            self.save_device_data(self.monitor_data_file if column_name == "Serial Number" else self.pc_data_file, data)
            self.devices_tree.delete(selected_item)

    def create_outlook_draft(self):
        outlook_path = r"C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE"
        recipient_email = "jmackey@optionmetrics.com"
        subject = "Purchase Request"
        body = "Hello, Julissa,\n\nI would like to request the purchase of the following items:\n\n[List of items]\n\nThank you."
        
        subprocess.Popen([outlook_path, '/c', 'ipm.note', '/m', f'mailto:{recipient_email}?subject={subject}&body={body}'])

    def show_monitor_page(self, event):
        self.device_screen("Monitor Details", self.monitor_data, "Serial Number")

    def show_pc_page(self, event):
        self.device_screen("PC Details", self.pc_data, "Hostname")

if __name__ == "__main__":
    root = Tk()
    app = IMS(root)
    root.mainloop()
