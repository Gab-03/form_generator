import pandas as pd
from datetime import datetime
from docxtpl import DocxTemplate
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import pandas as pd
from tkinter import ttk
import threading
import time
import xlwings as xw
from datetime import datetime
import numpy as np
import warnings
warnings.filterwarnings('ignore')

class FileSelectorApp:
    def __init__(self, master):
        self.master = master
        self.master.title("Prepare Form")
        self.file1_path = tk.StringVar()
        self.loading_window = None
        self.progress_var = tk.DoubleVar()
        self.create_widgets()

    def create_widgets(self):
        self.create_file_entry(0, "User Database (Excel)", self.file1_path, self.browse_file1)
        tk.Button(self.master, text="Prepare Form", command=self.submit).grid(row=3, column=1, pady=20)

    def create_file_entry(self, row, label_text, var, command):
        tk.Label(self.master, text=label_text).grid(row=row, column=0, pady=10)
        tk.Entry(self.master, textvariable=var, width=50, state='disabled').grid(row=row, column=1, pady=10)
        tk.Button(self.master, text="Browse", command=command).grid(row=row, column=2, pady=10)

    def browse_file1(self):
        self.file1_path.set(filedialog.askopenfilename())

    def submit(self):
        file1 = self.file1_path.get()

        if file1:
            try:
                # Show loading screen before reading files
                self.show_loading_screen()

                # Start background task
                threading.Thread(target=self.read_files_and_process, args=(file1,)).start()
                messagebox.showinfo("Process Complete!", "files downloaded successfully.")


            except Exception as e:
                # Display an error message if an error occurs
                messagebox.showerror("Error", f"An error occurred: {str(e)}")

        else:
            messagebox.showerror("Error", "Please upload all the required files.")

    def read_files_and_process(self, file1):

        doc = DocxTemplate("Form_Template.docx")
        # doc = DocxTemplate(file2)

        my_name = "Frank Andrade"
        my_phone = "(123) 456-789"
        my_email = "frank@gmail.com"
        my_address = "123 Main Street, NY"
        today_date = datetime.today().strftime("%d %b, %Y")

        my_context = {'my_name': my_name, 'my_phone': my_phone, 'my_email': my_email, 'my_address': my_address,'today_date': today_date}

        df = pd.read_excel(file1, sheet_name='Form1')

        
        for index, row in df.iterrows():
            isr = ""
            do = ""
            cdt = ""
            fss = ""
            cst= ""
            ds = ""
            role = row["Role"].split(";")[:-1]
            print(role)
            if "CSP/Planners (Reports)" in role:
                isr = "true"

            if "Field Partner" in role:
                cdt = "true"
                isr = "true"

            if "ISR Supervisor" in role:
                ds = "true"
                isr = "true"

            if "FSS/FSM" in role:
                fss = "true"

            if "Encoder" in role:
                do = "true"

            site = row["Site (e.g. 4001,4002,4003)"]
            print(type(site))

            if str(row["Site (e.g. 4001,4002,4003)"])[0].isdigit() and isinstance(site,int):
                site = str(row["Site (e.g. 4001,4002,4003)"])
                site = ','.join([site[i:i+4] for i in range(0, len(site), 4)])
                    
            context = {'first_name': row['First Name'],
                            'last_name': row['Last Name'],
                            'email': row["Email2"],
                            'site' : site,
                            'date': datetime.now().strftime("%Y-%m-%d"),
                            'role' : row["Role"],
                            'isr' : isr,
                            'do' : do,
                            'cdt' : cdt,
                            'fss': fss,
                            'isr': isr,
                            'cst': cst,
                            'ds': ds
                }

            context.update(my_context)
            doc.render(context)
            doc.core_properties.read_only = "false"
            doc.save(f"UserAccessForm_doc_{index}.docx")

        # Start of Excel
        master_wb = xw.Book(r"Role Mapping for PH_PAV.xlsx")
        # master_wb = xw.Book(file3)
        master_sheets = master_wb.sheets
        newdata_wb = xw.Book(file1)
        # print(newdata_wb.sheets, "<----")
        # print(newdata_wb.sheets[0], "<----")
        # print(newdata_wb.sheets[1], "<----")
        newdata_wb.sheets[1].range('A2').expand()
        master_wb.sheets[0].range('A1').end('down') ## ctrl + down
        new_data_raw = newdata_wb.sheets[0].range('G2').expand().value

        temp_data = [i[0:5] + i[5:] for i in new_data_raw]

        new_data_raw = newdata_wb.sheets[0].range('G2').expand().value
        temp_data = [i[0:5] + i[5:] for i in new_data_raw]
        if len(temp_data) == 5:
            temp_data = [temp_data]


        for x in range(0,len(temp_data)):
            master_wb.sheets["USERS"].range((6 + x, 3)).value = "ULP"
            master_wb.sheets["USERS"].range((6 + x, 4)).value = "PHILIPPINES"
            master_wb.sheets["USERS"].range((6 + x, 5)).value = "DT"
            master_wb.sheets["USERS"].range((6 + x, 6)).value = temp_data[x][0] + " " + temp_data[x][1]
            master_wb.sheets["USERS"].range((6 + x, 7)).value = "DT"
            master_wb.sheets["USERS"].range((6 + x, 8)).value = "DT"
            master_wb.sheets["USERS"].range((6 + x, 12)).value = "ULP"
            master_wb.sheets["USERS"].range((6 + x, 9)).value = temp_data[x][3] #email

            site = temp_data[x][2]
            if str(site)[0].isdigit() and isinstance(site, float) :
                site = str(site).rstrip(".0")
                print(site)
                site = ','.join([site[i:i+4] for i in range(0, len(site), 4)])
            
            master_wb.sheets["USERS"].range((6 + x, 13)).value = site

            #for multiple roles
            roles = temp_data[x][4].replace(";",",")[:-1].split(",")
            roles_x = []
            for i in roles:
                if i == "Encoder":
                    roles_x.append("Distributor Operator")

                elif i == "CSP/Planners (Reports)":
                    roles_x.append("ISR Reports Supervisor")

                elif i == "Field Partner":
                    roles_x.append("CD DT Executive")
                    roles_x.append("ISR Reports Supervisor")

                elif i == "FSS/FSM":
                    roles_x.append("FSS/FSM")

                elif i == "ISR Supervisor":
                    roles_x.append("Distributor Supervisor")
                    roles_x.append("ISR Reports Supervisor")

            print(roles_x)
            print(list(dict.fromkeys(roles_x)))
            my_list = list(dict.fromkeys(roles_x))
            result_string = ', '.join(my_list)
            print(result_string)

            # master_wb.sheets["USERS"].range((6 + x, 14)).value = temp_data[x][4].replace(";",", ")[:-2]
            master_wb.sheets["USERS"].range((6 + x, 14)).value = result_string
            print(temp_data[x][4])
            
            
            roles = temp_data[x][4].replace(";",",")[:-1].split(",")
            if "Encoder" in roles:
                master_wb.sheets["USERS"].range((6 + x, 15)).value = "x"

            if "CSP/Planners (Reports)" in roles:
                master_wb.sheets["USERS"].range((6 + x, 20)).value = "x"

            if "ISR Supervisor" in roles:
                master_wb.sheets["USERS"].range((6 + x, 17)).value = "x"
                master_wb.sheets["USERS"].range((6 + x, 20)).value = "x"        

            if "Field Partner" in roles:
                master_wb.sheets["USERS"].range((6 + x, 20)).value = "x"
                master_wb.sheets["USERS"].range((6 + x, 28)).value = "x"

            if "FSS/FSM" in roles:
                master_wb.sheets["USERS"].range((6 + x, 19)).value = "x"

        # Create a new workbook
        new_wb = xw.Book()

        # Copy the sheets from the existing workbook to the new workbook
        for sheet in master_wb.sheets:
            sheet.copy(before=new_wb.sheets[0])

        # Save the new workbook to a file
        new_filename = 'path_to_new_workbook.xlsx'
        new_wb.save(new_filename)

        # Close both workbooks
        master_wb.close()
        new_wb.close()

        self.close_loading_screen()



    def show_loading_screen(self):
        self.loading_window = tk.Toplevel(self.master)
        self.loading_window.title("Loading...")

        # Label to display progress text
        progress_label = tk.Label(self.loading_window, text="Reading all files: 0%", padx=20, pady=20)
        progress_label.pack()

        # Progress bar
        progress_bar = ttk.Progressbar(self.loading_window, variable=self.progress_var, maximum=100, length=200, mode='determinate')
        progress_bar.pack(padx=20, pady=10)

        self.master.update()

        # Store progress label and bar in instance variables
        self.progress_label = progress_label
        self.progress_bar = progress_bar


    def update_progress(self, value):
        self.progress_var.set(value)
        self.loading_window.update()

        # Update progress label text
        self.progress_label.config(text=f"Loading: {int(value)}%")

    def close_loading_screen(self):
        if self.loading_window:
            self.loading_window.destroy()
            self.loading_window = None

    def process_files(self, file1):
        
        self.update_progress(0)
        self.close_loading_screen()

if __name__ == "__main__":
    root = tk.Tk()
    app = FileSelectorApp(root)
    root.mainloop()
