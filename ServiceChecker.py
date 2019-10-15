import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import servicetest
from lxml import etree as ET
import user_data_handler as udata
import defs, sendEmail, os, psutil, time, smtplib
from email.message import EmailMessage
import concurrent.futures

class MainApplication(ttk.Frame):
    def __init__(self, parent, *args, **kwargs):
        ttk.Frame.__init__(self, parent, *args, **kwargs)
        self.parent = parent
        self.notebook = ttk.Notebook(parent)
        self.notebook.grid(row=0, column=0, padx=20, pady=20)

        # Tab Setup
        self.email_tab = ttk.Frame(self.notebook)
        self.services_tab = ttk.Frame(self.notebook)
        self.task_scheduler = ttk.Frame(self.notebook)
        self.finished_buttons = ttk.Frame(parent)

        self.email_tab.grid(row=0, column=0, padx=20, pady=20)
        self.services_tab.grid(row=0, column=0, padx=20, pady=20)
        self.task_scheduler.grid(row=0, column=0, padx=20, pady=20)
        self.finished_buttons.grid(row=1, column=0, padx=20, pady=20)

        # Email Tab Variables
        self.display_email_smtp = tk.StringVar()
        self.display_email_smtp.set(defs.email_smtp)
        self.display_email_port = tk.IntVar()
        self.display_email_port.set(defs.email_port)
        self.display_email_from = tk.StringVar()
        self.display_email_from.set(defs.email_from)
        self.display_email_login = tk.StringVar()
        self.display_email_login.set(defs.email_user)
        self.display_email_passwd = tk.StringVar()
        self.display_email_passwd.set(defs.email_passwd)
        self.display_email_recipient = defs.email_recipient
        self.hidden = "\u25cf"

        # Email Tab If Empty strings
        if self.display_email_smtp.get() == "None":
            self.display_email_smtp.set("smtp.example.com")

        if self.display_email_from.get() == "None":
            self.display_email_from.set("Somebody or someone@example.com")

        if self.display_email_login.get() == "None":
            self.display_email_login.set("somebody@example.com")

        if self.display_email_passwd.get() == "None":
            self.display_email_passwd.set("")

        # Service Tab Variables
        self.display_service_list = defs.services

        # Task Scheduler Variables

        # Email Tab
        self.email_instructions_lbl = ttk.LabelFrame(self.email_tab, text="Intructions:", width=36)
        self.email_instructions_txt = ttk.Label(self.email_instructions_lbl, text="Please enter email settings.\n\nSMTP server example: smtp.gmail.com\nPort: 457\n\nThe \"From Sender\" field can be any unique name, such as \"Service Notifier\"\n\nUse the \"Login Email Address\" to enter your login user name. For example: \"someone@gmail.com\"\n\nEnter the password for the account in the provided field.\n\nAdd recipient\'s email addresses in the \"Add Recipient\" field and then click \"Add\". You can remove email address by selecting the email address you wish to remove and click \"Remove\"\n\nOnce completed, click \"Test Email Settings\" to confirm everything is configured correctly. Any email address in the recipient list will recieve a test email.\n\nOnce everything is setup correctly, click \"Update Configuration\" to save changes.\n\n*Note: Changes are saved in an encrypted file called \"ServiceChecker.conf\". Your password is only visible to this software and is never saved in plain-text.", wraplength=350)
        self.email_smtp_lbl = ttk.Label(self.email_tab, text="SMTP Server:")
        self.email_smtp_entry = ttk.Entry(self.email_tab, width=36, textvariable=self.display_email_smtp)
        self.email_port_lbl = ttk.Label(self.email_tab, text="Port:")
        self.email_port_entry = ttk.Entry(self.email_tab, textvariable=self.display_email_port, width=15)
        self.email_from_lbl = ttk.Label(self.email_tab, text="From Sender:")
        self.email_from_entry = ttk.Entry(self.email_tab, width=36, textvariable=self.display_email_from)
        self.email_login_lbl = ttk.Label(self.email_tab, text="Login Email Address:")
        self.email_login_entry = ttk.Entry(self.email_tab, width=36, textvariable=self.display_email_login)
        self.email_passwd_lbl = ttk.Label(self.email_tab, text="Login Password:")
        self.email_passwd_entry = ttk.Entry(self.email_tab, show=self.hidden, width=36, textvariable=self.display_email_passwd)
        self.email_passwd_show = ttk.Button(self.email_tab, width=15, text="Show")
        self.email_passwd_show.bind("<ButtonPress>", lambda btn: self.reveal_passwd(event="<ButtonPress>", btn=self.email_passwd_entry))
        self.email_passwd_show.bind("<ButtonRelease>", lambda btn: self.conceal_passwd(event="<ButtonRelease>", btn=self.email_passwd_entry))
        self.email_recipient_lbl = ttk.Label(self.email_tab, text="Add Recipient:")
        self.email_recipient_entry = ttk.Entry(self.email_tab, width=36, text="")
        self.email_recipient_add = ttk.Button(self.email_tab, width=15, text="Add", command=self.email_recipient_add_button)
        self.email_recipient_list = tk.Listbox(self.email_tab, width=36, height=5, bg="white")
        self.email_recipient_remove = ttk.Button(self.email_tab, width=15, text="Remove", command=self.email_recipient_remove_button)
        self.email_test_button = ttk.Button(self.email_tab, width=36, text="Test Email Settings", command=self.check_email)

        self.email_instructions_lbl.grid(row=0, column=0, rowspan=20, padx=(20, 0), pady=20, sticky="n")
        self.email_instructions_txt.grid(row=0, column=0)
        self.email_smtp_lbl.grid(row=0, column=1, padx=20, pady=(20, 0), sticky="w")
        self.email_smtp_entry.grid(row=1, column=1, padx=20, pady=(0, 10))
        self.email_port_lbl.grid(row=0, column=2, padx=0, pady=(20, 0), sticky="w")
        self.email_port_entry.grid(row=1, column=2, padx=(0, 20), pady=(0, 10))
        self.email_from_lbl.grid(row=2, column=1, padx=20, pady=(10, 0), sticky="w")
        self.email_from_entry.grid(row=3, column=1, padx=20, pady=(0, 10))
        self.email_login_lbl.grid(row=4, column=1, padx=20, pady=(10, 0), sticky="w")
        self.email_login_entry.grid(row=5, column=1, padx=20, pady=(0, 10))
        self.email_passwd_lbl.grid(row=6, column=1, padx=20, pady=(10, 0), stick="w")
        self.email_passwd_entry.grid(row=7, column=1, padx=20, pady=(0, 10))
        self.email_passwd_show.grid(row=7, column=2, padx=(0, 20), pady=(0, 10))
        self.email_recipient_lbl.grid(row=8, column=1, padx=20, pady=(10, 0), sticky="w")
        self.email_recipient_entry.grid(row=9, column=1, padx=20, pady=(0, 10))
        self.email_recipient_add.grid(row=9, column=2, padx=(0, 20), pady=(0, 10))
        self.email_recipient_list.grid(row=10, column=1, padx=20, pady=10)
        self.email_recipient_remove.grid(row=10, column=2, padx=(0, 20), pady=10, sticky="n")
        self.email_test_button.grid(row=11, column=1, padx=20, pady=5)

        # Populate Email Recipient List
        for x in self.display_email_recipient:
            self.email_recipient_list.insert(tk.END, x)

        # Service Tab
        self.service_instructions_lbl = ttk.LabelFrame(self.services_tab, text="Instructions:")
        self.service_instructions_txt = ttk.Label(self.service_instructions_lbl, text="Find the service you need to monitor under the \"Available Services\" list. This software polls Windows' currently installed services and lists them here and shows its current status. You can expand the service see its display name.\n\nOnce you have found the service, click on its name and click \"Add\".\n\nYou can remove services by selecting them from the \"Monitored Services\" list and clicking \"Remove\".\n\nClick \"Update Configuration\" to save changes to the configuration file.", wraplength=350)
        self.service_available_lbl = ttk.Label(self.services_tab, text="Available Services:")
        self.service_available_tree = ttk.Treeview(self.services_tab)
        self.service_control_lbl = ttk.Label(self.services_tab, text="Add or remove monitored services:")
        self.service_add_btn = ttk.Button(self.services_tab, text="Add", width=15, command=self.service_add_btn_clicked)
        self.service_remove_btn = ttk.Button(self.services_tab, text="Remove", width=15, command=self.service_remove_btn_clicked)
        self.service_monitored_lbl = ttk.Label(self.services_tab, text="Monitored Services:")
        self.service_monitored_list = tk.Listbox(self.services_tab, width=52, height=8, bg="white")

        self.service_instructions_lbl.grid(row=0, column=0, rowspan=20, padx=(20, 0), pady=20, sticky="n")
        self.service_instructions_txt.grid(row=0, column=0)
        self.service_available_lbl.grid(row=0, column=1, padx=20, pady=(20, 0), sticky="w")
        self.service_available_tree.grid(row=1, column=1, columnspan=2, padx=20, pady=(0, 10), sticky="w")
        self.service_control_lbl.grid(row=2, column=1, padx=20, pady=(10, 0), sticky="w")
        self.service_add_btn.grid(row=3, column=1, padx=(20, 0), pady=(0, 10), sticky="w")
        self.service_remove_btn.grid(row=3, column=2, padx=(0, 10), pady=(0, 10), sticky="w")
        self.service_monitored_lbl.grid(row=4, column=1, padx=20, pady=(10, 0), sticky="w")
        self.service_monitored_list.grid(row=5, column=1, columnspan=2, padx=20, pady=(0, 20), sticky="w")

        # Service Available Definitions
        self.service_available_tree.bind("<<TreeviewSelect>>", self.service_available_event)
        self.service_available_tree["columns"]=("status")
        self.service_available_tree.column("#0", width=250, minwidth=36)
        self.service_available_tree.column("status", width=50, minwidth=36)

        self.service_available_tree.heading("#0", text="Name", anchor=tk.W)
        self.service_available_tree.heading("status", text="Status", anchor=tk.W)

        # Populate Service Available Tree
        self.inum = 0
        for serv in psutil.win_service_iter():
            self.inum += 1
            i = str(serv.name())
            i = self.service_available_tree.insert("", self.inum, text=serv.name(), values=serv.status())
            self.service_available_tree.insert(i, "end", text=serv.display_name())

        # Populate Services Monitored List
        for x in self.display_service_list:
            self.service_monitored_list.insert(tk.END, x)

        # Task Scheduler Tab
        self.scheduler_instructions_lbl = ttk.LabelFrame(self.task_scheduler, text="Instructions", width=36)
        self.scheduler_instructions_txt = ttk.Label(self.scheduler_instructions_lbl, text="Use this tab to create a scheduled task.\n\nYou can set the program to run on \"User Logon\" (recommended) as well as other options, such as once per day, every hour, etc.\n\nOnce the task has been created, you can edit its settings in the standard Windows Task Scheduler.\n\n", wraplength=350)
        self.scheduler_general_lbl = ttk.LabelFrame(self.task_scheduler, text="General").pack()
        self.scheduler_general_name_lbl = ttk.Label(self.scheduler_general_lbl, text="Name: ", width=10)
        self.scheduler_general_name_txt = ttk.Label(self.scheduler_general_lbl, text="Service Monitor", width=16)
        self.scheduler_general_location_lbl = ttk.Label(self.scheduler_general_lbl, text="Location: ", width=10)
        self.scheduler_general_location_txt = ttk.Label(self.scheduler_general_lbl, text="\\", width=16)
        self.scheduler_general_author_lbl = ttk.Label(self.scheduler_general_lbl, text="Author: ", width=10)
        self.scheduler_general_author_txt = ttk.Label(self.scheduler_general_lbl, text="Tim Finley", width=16)
        self.scheduler_general_description_lbl = ttk.Label(self.scheduler_general_lbl, text="Description: ", width=10)
        self.scheduler_general_description_txt = ttk.Label(self.scheduler_general_lbl, text="This task monitors specified services to ensure their installation is still present in case of Windows Updates or Anti-Virus software unintentional uninstallations. The task will email a designated user or users in case any services are not found or are not running on the system.", width=16)
        self.scheduler_triggers_lbl = ttk.LabelFrame(self.task_scheduler, text="Triggers", width=36)
        self.scheduler_triggers_at_logon = ttk.Label(self.scheduler_triggers_lbl, text="filler")

        self.scheduler_instructions_lbl.grid(row=0, column=0, rowspan=20, padx=(20, 0), pady=20, sticky="n")
        self.scheduler_instructions_txt.grid(row=0, column=0)
        self.scheduler_general_lbl.grid(row=0, column=1, padx=20, pady=20)
        self.scheduler_general_name_lbl.grid(row=0, column=0, padx=0, pady=0)
        self.scheduler_general_name_txt.grid(row=0, column=1, padx=0, pady=0)
        self.scheduler_general_location_lbl.grid(row=1, column=0, padx=0, pady=0)
        self.scheduler_general_location_txt.grid(row=1, column=1, padx=0, pady=0)
        self.scheduler_general_author_lbl.grid(row=2, column=0, padx=0, pady=0)
        self.scheduler_general_author_txt.grid(row=2, column=1, padx=0, pady=0)
        self.scheduler_general_description_lbl.grid(row=3, column=0, padx=0, pady=0)
        self.scheduler_general_description_txt.grid(row=3, column=1, padx=0, pady=0)
        self.scheduler_triggers_lbl.grid(row=1, column=1, padx=0, pady=0)
        self.scheduler_triggers_at_logon.grid(row=0, column=0)

        # Populate Notebook
        self.notebook.add(self.email_tab, text="Email Configuration")
        self.notebook.add(self.services_tab, text="Services Configuration")
        self.notebook.add(self.task_scheduler, text="Task Scheduler")

        self.update_xml = ttk.Button(self.finished_buttons, width=36, text="Update Configuration", command=self.update_xml_click)
        self.update_xml.grid(row=0, column=0, padx=20, pady=20)

    def update_xml_click(self):
        try:
            root = ET.Element("main")
            email = ET.SubElement(root, "email")
            services = ET.SubElement(root, "services")

            write_smtp = ET.SubElement(email, "smtp").text=str(self.display_email_smtp.get())
            write_port = root[0][0].set("port", str(self.display_email_port.get()))
            write_sender_login = ET.SubElement(email, "sender_login").text=str(self.display_email_login.get())
            write_sender_passwd = ET.SubElement(email, "sender_passwd").text=str(self.display_email_passwd.get())
            write_from_sender = ET.SubElement(email, "from_addr").text=str(self.display_email_from.get())

            for x in range(len(self.display_service_list)):
                service = ET.SubElement(services, "service")
                service.text = self.display_service_list[x]

            for x in range(len(self.display_email_recipient)):
                recipient = ET.SubElement(email, "recipient")
                recipient.text = self.display_email_recipient[x]
            finale = ET.ElementTree(root)
            # print(ET.tostring(finale, pretty_print=True))
            with open(defs.udef, "wb") as file:
                file.write(udata.fernet.encrypt(ET.tostring(finale, pretty_print=True)))

            messagebox.showinfo("Write Successfull", "Configuration Settings have been saved")
        except Exception as ex:
            if "expected floating-point number but got \"\"" in str(ex):
                messagebox.showerror("Write Error", "Please enter a Port Number")
            else:
                print(ex)

    def service_available_event(self, event):
        iid = self.service_available_tree.focus()
        # print(iid)

    def service_add_btn_clicked(self):
        selected_service = self.service_available_tree.focus()
        selected_dict = self.service_available_tree.item(selected_service)
        service_name = selected_dict["text"]
        self.display_service_list.append(service_name)
        self.service_monitored_list.insert(tk.END, service_name)

    def service_remove_btn_clicked(self):
        service_selection = self.service_monitored_list.curselection()
        for index in reversed(service_selection):
            self.service_monitored_list.delete(index)
        self.display_service_list = []
        for item in reversed(self.service_monitored_list.get(0, tk.END)):
            self.display_service_list.append(item)

    def reveal_passwd(self, event, btn):
        btn.config(show="")

    def conceal_passwd(self, event, btn):
        btn.config(show=self.hidden)

    def email_recipient_add_button(self):
        recipient_entry = self.email_recipient_entry.get()
        self.display_email_recipient.append(recipient_entry)
        self.email_recipient_list.insert(tk.END, recipient_entry)
        self.email_recipient_entry.delete(0, tk.END)

    def email_recipient_remove_button(self):
        recipient_selection = self.email_recipient_list.curselection()
        for index in reversed(recipient_selection):
            self.email_recipient_list.delete(index)
        self.display_email_recipient = []
        for item in reversed(self.email_recipient_list.get(0, tk.END)):
            self.display_email_recipient.append(item)

    def check_email_waitbox(self):
        popup = tk.Toplevel()
        popup_label = ttk.Label(popup, text="Message being sent").grid(row=0, column=0)
        progress = 0
        progress_var = tk.DoubleVar()
        progress_bar = ttk.Progressbar(popup, variable=progress_var, maximum=100)
        progress_bar.grid(row=1, column=0)

    def check_email_backend(self):
        t1=time.perf_counter()
        print(t1)
        sender_name = self.display_email_from.get()
        test = "testlog.txt"
        for x in self.display_email_recipient:
            recipients = [x]
        with open(test, "w+") as f:
            f.write(f"From: {sender_name}\n")
            f.write(f"To: {recipients}\n")
            f.write("Subject: Testing\n\n")
            f.write("This is a test mailer.")
        try:
            server = smtplib.SMTP_SSL(self.display_email_smtp.get(), self.display_email_port.get())
            server.ehlo()
            server.login(self.display_email_login.get(), self.display_email_passwd.get())
            with open(test) as fp:
                msg = EmailMessage()
                msg.set_content(fp.read())
            msg["Subject"] = "Test"
            msg["From"] = self.display_email_from.get()
            msg["To"] = recipients
            server.send_message(msg)
            server.close()
        except Exception as ex:
            if "getaddrinfo" in str(ex):
                messagebox.showerror("Something Went Wrong", "Check your settings or your internet connection before continuing.")
            if "A connection attempt failed" in str(ex):
                messagebox.showerror("Something Went Wrong", f"Check your settings.\n\nError Returned: \n{ex}")
            print(ex)
        t2=time.perf_counter()
        print(t2 - t1)

    def check_email(self):
        self.check_email_backend()
        # with concurrent.futures.ThreadPoolExecutor() as executor:
        #     g1 = executor.submit(self.check_email_waitbox)
        #     g2 = executor.submit(self.check_email_backend)

if __name__ == "__main__":
    root = tk.Tk()
    MainApplication(root)
    root.mainloop()
