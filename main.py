# AKASHIC - ATTENDANCE MONITORING

import tkinter
from tkinter import messagebox
from tkinter import ttk
from tkinter import *
from tkinter import filedialog as fd
import customtkinter
import sqlite3 
import openpyxl
import os
from openpyxl.styles.alignment import Alignment
from pathlib import Path
from datetime import date

customtkinter.set_appearance_mode("Dark") # Set to Dark mode
customtkinter.set_default_color_theme("dark-blue")  # Set 'dark-blue' theme


class App(customtkinter.CTk, tkinter.Tk):
    def __init__(self):
        super().__init__()

        # Main Window
        self.title("AKASHIC - Attendance Monitoring")
        self.geometry(f"{1050}x{600}")

        # Configure grid layout (4x4)
        self.grid_columnconfigure(1, weight=1)
        self.grid_columnconfigure((2, 3), weight=0)
        self.grid_rowconfigure((0, 1, 2, 3), weight=1)

        # Sidebar Frame with Widgets
        self.sidebar_frame = customtkinter.CTkFrame(self, width=150, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, rowspan=5, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(4, weight=1)
        self.logo_label = customtkinter.CTkLabel(self.sidebar_frame, text="AKASHIC", font=customtkinter.CTkFont(family="Impact", size=40, weight="normal"))
        self.logo_label.grid(row=0, column=0, padx=20, pady=(20, 10))
        self.sidebar_button_1 = customtkinter.CTkButton(self.sidebar_frame, text="Masterlist", command=self.masterlist)
        self.sidebar_button_1.grid(row=1, column=0, padx=20, pady=10)
        self.sidebar_button_2 = customtkinter.CTkButton(self.sidebar_frame, text="Record", command=self.records)
        self.sidebar_button_2.grid(row=2, column=0, padx=20, pady=10)
        self.sidebar_button_3 = customtkinter.CTkOptionMenu(self.sidebar_frame, values=["BSCOE 2-6"])
        self.sidebar_button_3.grid(row=3, column=0, padx=20, pady=40)
        self.appearance_mode_label = customtkinter.CTkLabel(self.sidebar_frame, text="Appearance Mode:", anchor="w")
        self.appearance_mode_label.grid(row=5, column=0, padx=20, pady=(10, 0))
        self.appearance_mode_optionemenu = customtkinter.CTkOptionMenu(self.sidebar_frame, values=["Dark", "Light", "System"], command=self.change_appearance_mode_event)
        self.appearance_mode_optionemenu.grid(row=6, column=0, padx=20, pady=(10, 10))
        self.scaling_label = customtkinter.CTkLabel(self.sidebar_frame, text="UI Scaling:", anchor="w")
        self.scaling_label.grid(row=7, column=0, padx=20, pady=(10, 0))
        self.scaling_optionemenu = customtkinter.CTkOptionMenu(self.sidebar_frame, values=["80%", "90%", "100%", "110%", "120%"], command=self.change_scaling_event)
        self.scaling_optionemenu.grid(row=8, column=0, padx=20, pady=(10, 20))


        # ------------------------------- FIRST PANEL -----------------------------------------

        # Frames for Label and Layout purposes
        self.masterlist_frame = customtkinter.CTkFrame(self)
        self.masterlist_frame.grid(row=0, column=1, padx=(15, 15), pady=(12, 0), columnspan=2, sticky="nsew")
        self.masterlist_frame.grid_columnconfigure(0, weight=1)

        self.masterlLabel_frame = customtkinter.CTkFrame(self)
        self.masterlLabel_frame.grid(row=1, column=1, padx=(15, 15), pady=(10, 0), columnspan=2, sticky="nsew")
        self.masterlLabel_frame.grid_columnconfigure(0, weight=1)

        self.ml_label = customtkinter.CTkLabel(self.masterlist_frame, text=f"MASTERLIST\n{self.sidebar_button_3.get().upper()}  -  DSA")
        self.ml_label.grid(row=0, column=0, padx=(15, 15), pady=(8, 0), sticky="n")


        # Treeview: Displays data from database
        self.terminal_tree = ttk.Treeview(self)
        self.terminal_tree.grid(row=1, column=1, padx=(15, 15), pady=(5, 0), columnspan=2, sticky=tkinter.NSEW)
        self.terminal_tree["columns"] = ("1", "2", "3", "4", "5")
        self.terminal_tree['show'] = 'headings'
        self.terminal_tree.column("1", width=10, anchor='c')
        self.terminal_tree.heading("1", text="Student No.")
        self.terminal_tree.column("2", width=200, anchor='w')
        self.terminal_tree.heading("2", text="Name")
        self.terminal_tree.column("3", width=100, anchor='c')
        self.terminal_tree.heading("3", text="Course Year & Section")
        self.terminal_tree.column("4", width=100, anchor='c')
        self.terminal_tree.heading("4")
        self.terminal_tree.column("5", width=100, anchor='c')
        self.terminal_tree.heading("5", text="Status")

        self.columnconfigure(2, weight=1) # column with treeview
        self.rowconfigure(2, weight=1) # row with treeview  


        # Frame for Treeview editing tools (Left Side)
        self.edit_masterlist_frame1 = customtkinter.CTkFrame(self)
        self.edit_masterlist_frame1.grid(row=2, column=1, padx=(15, 5), pady=(12, 0), sticky="nsew")
        self.edit_masterlist_frame1.grid_columnconfigure(1, weight=1)

        self.add_button = customtkinter.CTkButton(self.edit_masterlist_frame1, text="Add", fg_color="#059142" ,hover_color="#03692f", command=self.add_student)
        self.add_button.grid(row=1, column=1, padx=(20, 20), pady=13, sticky="w")

        self.clear_button = customtkinter.CTkButton(self.edit_masterlist_frame1, text="Clear", command=self.clear_entry)
        self.clear_button.grid(row=1, column=2, padx=(20, 20), pady=13, sticky="e")


        # Entry for Adding Treeview Contents
        self.entry_frame = customtkinter.CTkFrame(self.edit_masterlist_frame1)
        self.entry_frame.grid(row=2, column=1, padx=(15, 15), pady=(10, 0), columnspan=2, sticky="nsew")
        self.entry_frame.grid_columnconfigure(0, weight=1)

        self.entry_label = customtkinter.CTkLabel(self.entry_frame, text="Student Information")
        self.entry_label.place(relx=0.5, rely=0.1, anchor="center")

        self.name_entry = customtkinter.CTkEntry(self.entry_frame, placeholder_text="Name")
        self.name_entry.grid(row=0, column=0, padx=(20, 20), pady=(40, 5), sticky="nsew")

        self.stnum_entry = customtkinter.CTkEntry(self.entry_frame, placeholder_text="Student Number")
        self.stnum_entry.grid(row=1, column=0, padx=(20, 20), pady=5, sticky="nsew")

        self.section_entry = customtkinter.CTkEntry(self.entry_frame, placeholder_text="Course Year & Section")
        self.section_entry.grid(row=2, column=0, padx=(20, 20), pady=(5, 10), sticky="nsew")

        self.status_option = customtkinter.CTkOptionMenu(self.entry_frame, values=["Regular", "Irregular", "Withdrawn", "Dropped", "Transferee"])
        self.status_option.grid(row=3, column=0, padx=20, pady=(5, 20), sticky="nsew")


        # Frame for Treeview editing tools (Right Side)
        self.edit_masterlist_frame2 = customtkinter.CTkFrame(self)
        self.edit_masterlist_frame2.grid(row=2, column=2, padx=(5, 15), pady=(12, 0), sticky="nsew")
        self.edit_masterlist_frame2.grid_columnconfigure(2, weight=1)

        # Buttons for Organizing Treeview Contents
        self.sort_button = customtkinter.CTkButton(self.edit_masterlist_frame2, text="Sort", width=90, command=self.sort_data_entries)
        self.sort_button.grid(row=1, column=0, padx=(15, 5), pady=13, sticky="w")

        self.update_button = customtkinter.CTkButton(self.edit_masterlist_frame2, text="Update", width=90, command=self.update_panel)
        self.update_button.grid(row=1, column=2, padx=(5, 5), pady=13, sticky="e")

        self.delete_button = customtkinter.CTkButton(self.edit_masterlist_frame2, text="Delete", fg_color= "dark red", width=90, hover_color="#4c0303" , command=self.delete_student)
        self.delete_button.grid(row=1, column=3, padx=(5, 15), pady=13, sticky="e")

        # Text Summary of Treeview Data
        self.summary_details = customtkinter.CTkFrame(self.edit_masterlist_frame2)
        self.summary_details.grid(row=2, column=0, padx=(15, 15), pady=(10, 0), columnspan=4, sticky="nsew")
        self.edit_masterlist_frame2.grid_columnconfigure(0, weight=1)

        self.summary_label = customtkinter.CTkLabel(self.summary_details, text="List Summary")
        self.summary_label.place(relx=0.5, rely=0.1, anchor="center") 

        self.number_of_students = customtkinter.CTkLabel(self.summary_details, text="Number of Students Enrolled: ")
        self.number_of_students.grid(row=1, column=0, padx=(15, 15), pady=(50, 0), sticky="nw")

        self.blank_space = customtkinter.CTkLabel(self.summary_details, text="           ")
        self.blank_space.grid(row=1, column=1, padx=(15, 15), pady=(50, 0), sticky="nsew") 

        self.total_numeric = customtkinter.CTkLabel(self.summary_details, text="0")
        self.total_numeric.grid(row=1, column=3, padx=(15, 15), pady=(50, 0), sticky="nsew") 

        self.regular_students = customtkinter.CTkLabel(self.summary_details, text="Regularly Enrolled Students: ")
        self.regular_students.grid(row=2, column=0, padx=(15, 15), pady=(30, 0), sticky="nw")
        self.regular_numeric = customtkinter.CTkLabel(self.summary_details, text="0")
        self.regular_numeric.grid(row=2, column=3, padx=(15, 15), pady=(30, 0), sticky="nsew")  

        self.irreg_students = customtkinter.CTkLabel(self.summary_details, text="Irregular Students: ")
        self.irreg_students.grid(row=3, column=0, padx=(15, 15), pady=(0, 0), sticky="nw") 
        self.irreg_numeric = customtkinter.CTkLabel(self.summary_details, text="0")
        self.irreg_numeric.grid(row=3, column=3, padx=(15, 15), pady=(0, 0), sticky="nsew") 

        self.transferee_students = customtkinter.CTkLabel(self.summary_details, text="Transferee Students: ")
        self.transferee_students.grid(row=4, column=0, padx=(15, 15), pady=(0, 18), sticky="nw")
        self.transferee_numeric = customtkinter.CTkLabel(self.summary_details, text="0")
        self.transferee_numeric.grid(row=4, column=3, padx=(15, 15), pady=(0, 18), sticky="nsew")  


        # Frame for Update panel (contains entry with data from databs)
        self.update_panel_frame = customtkinter.CTkFrame(self.edit_masterlist_frame2)
        self.update_panel_frame.grid(row=2, column=0, padx=(15, 15), pady=(5, 0), columnspan=4, sticky="nsew")
        self.update_panel_frame.grid_columnconfigure(0, weight=1)

        # Entry Labels
        self.update_name_label = customtkinter.CTkLabel(self.update_panel_frame, text="Name: ")
        self.update_name_label.place(relx=0.05, rely=0.15, anchor="w")
        self.update_stnum_label = customtkinter.CTkLabel(self.update_panel_frame, text="Student No: ")
        self.update_stnum_label.place(relx=0.05, rely=0.34, anchor="w")
        self.update_section_label = customtkinter.CTkLabel(self.update_panel_frame, text="Section: ")
        self.update_section_label.place(relx=0.05, rely=0.51, anchor="w")
        self.update_status_label = customtkinter.CTkLabel(self.update_panel_frame, text="Status: ")
        self.update_status_label.place(relx=0.05, rely=0.69, anchor="w")


        self.update_name_entry = customtkinter.CTkEntry(self.update_panel_frame, placeholder_text="Name")
        self.update_name_entry.grid(row=0, column=0, padx=(95, 10), pady=(20, 5), columnspan=2, sticky="nsew")

        self.update_stnum_entry = customtkinter.CTkEntry(self.update_panel_frame, placeholder_text="Student Number")
        self.update_stnum_entry.grid(row=1, column=0, padx=(95, 10), pady=(5, 5), columnspan=2, sticky="nsew")

        self.update_section_entry = customtkinter.CTkEntry(self.update_panel_frame, placeholder_text="Course Year & Section")
        self.update_section_entry.grid(row=2, column=0, padx=(95, 10), pady=(5, 5),columnspan=2, sticky="nsew")

        self.update_status_option = customtkinter.CTkOptionMenu(self.update_panel_frame, values=["Regular", "Irregular", "Withdrawn", "Dropped", "Transferee"])
        self.update_status_option.grid(row=3, column=0, padx=(95, 10), pady=(5, 15), columnspan=2, sticky="nsew")

        # Buttons for Confirmation of edit
        self.confirm_update_yes = customtkinter.CTkButton(self.update_panel_frame, text="Confirm", width=100, command=self.confirm_update)
        self.confirm_update_yes.grid(row=4, column=0, padx=(20, 10), pady=(0, 10), sticky="nsew")

        self.confirm_update_no = customtkinter.CTkButton(self.update_panel_frame, text="Cancel", command=self.cancel_update)
        self.confirm_update_no.grid(row=4, column=1, padx=(10, 20), pady=(0, 10), sticky="nsew")


        # ------------------------------- SECOND PANEL -----------------------------------------

        self.attendance_tool_frame = customtkinter.CTkFrame(self, height=50)
        self.attendance_tool_frame.grid(row=0, column=3, padx=(5, 15), pady=(12, 0), columnspan=2, rowspan=1, sticky="nsew")
        self.attendance_tool_frame.grid_columnconfigure(0, weight=0)
        self.attendance_tool_frame.grid_rowconfigure(0, weight=1)

        self.frame_label = customtkinter.CTkLabel(self.attendance_tool_frame, text="Record Details")
        self.frame_label.place(relx=0.5, rely=0.08, anchor="center")
        self.date_label = customtkinter.CTkLabel(self.attendance_tool_frame, text="Date:")
        self.date_label.place(relx=0.05, rely=0.22, anchor="w")
        self.section_label = customtkinter.CTkLabel(self.attendance_tool_frame, text="Section:")
        self.section_label.place(relx=0.05, rely=0.39, anchor="w")
        self.prof_label = customtkinter.CTkLabel(self.attendance_tool_frame, text="Instructor:")
        self.prof_label.place(relx=0.05, rely=0.56, anchor="w")
        self.class_label = customtkinter.CTkLabel(self.attendance_tool_frame, text="Course Description:")
        self.class_label.place(relx=0.5, rely=0.73, anchor="center")

        self.date_entry = customtkinter.CTkEntry(self.attendance_tool_frame, placeholder_text="Date")
        self.date_entry.grid(row=1, column=0, padx=(80, 15), pady=(15, 10), sticky="nsew")
        self.sect_entry = customtkinter.CTkEntry(self.attendance_tool_frame, placeholder_text="Section")
        self.sect_entry.grid(row=2, column=0, padx=(80, 15), pady=(0, 10), sticky="nsew")
        self.instructor_name_entry = customtkinter.CTkEntry(self.attendance_tool_frame, placeholder_text="Prof. Name")
        self.instructor_name_entry.grid(row=3, column=0, padx=(80, 15), pady=(0, 15), sticky="nsew")
        self.course_entry = customtkinter.CTkOptionMenu(self.attendance_tool_frame, values=["Data Structure and Algorithms"])
        self.course_entry.grid(row=4, column=0, padx=(20, 15), pady=(30, 20), sticky="nsew")

        self.notebook = customtkinter.CTkTextbox(self, width=250, activate_scrollbars=True, border_spacing=15)
        self.notebook.grid(row=2, column=3, padx=(5, 15), pady=(0, 0), sticky="nsew")

        self.link_file = customtkinter.CTkButton(self, text="Open Excel file", command=self.select_file)        
        self.link_file.grid(row=1, column=3, padx=(7, 15), pady=(7, 7), sticky="ew")

        self.attendance_frame = customtkinter.CTkScrollableFrame(self, label_text="Attendance Checklist")
        self.attendance_frame.grid(row=0, column=1, padx=(15, 5), pady=(12, 0), columnspan=2, rowspan=3, sticky="nsew")
        self.attendance_frame.grid_columnconfigure(0, weight=1)

        self.generate_report_button = customtkinter.CTkButton(self, text="GENERATE REPORT", fg_color="#059142", hover_color="#03692f", command=self.write_in_excel)
        self.generate_report_button.grid(row=3, column=1, padx=(15, 15), pady=(12, 10), columnspan=3, sticky="nsew")

        
        # CONFIGURATIONS
        self.appearance_mode_optionemenu.set("Dark")
        self.scaling_optionemenu.set("100%")
        self.display_data_treeview()
        self.update_panel_frame.grid_remove()
        self.date_entry.insert(0, date.today())
        self.sect_entry.insert(0, "BSCOE 2-6")
        self.sect_entry.configure(state="disabled", fg_color="#2b2c2e")
        self.sidebar_button_1.configure(fg_color="#14375e")

        if (self.fetchfilerecord()[0][0] != ""):
            if os.path.isfile(self.fetchfilerecord()[0][0]):
                if (".xlsx" in self.fetchfilerecord()[0][0]) or (".xls" in self.fetchfilerecord()[0][0]):
                    self.link_file.configure(fg_color="#059142", hover_color="#03692f", text=f"{Path(self.fetchfilerecord()[0][0]).name}")
                else:
                    self.link_file.configure(fg_color="#F05316", hover_color="#b52d21")
            else:
                self.link_file.configure(fg_color="#1f538d", hover_color="#14375e", text="Open Excel file..")
        else:
            self.link_file.configure(fg_color="#1f538d", hover_color="#14375e", text="Open Excel file..")
        
        self.notebook.insert("0.0", "Attendance Notepad\n\n" + "Late Comers:\n  - \n  - \n  - \n  - \n\nClass Notes:")
        self.attendance_frame.grid_remove()
        self.attendance_tool_frame.grid_remove()
        self.notebook.grid_remove()
        self.link_file.grid_remove()
        self.generate_report_button.grid_remove()




    # FUNCTIONS
    def get_checkbox_values(self):
        children_widgets = self.attendance_frame.winfo_children()
        checkbox_values = []
        checkbox_names = []
        for child in children_widgets:
            if "checkbox" in child.winfo_name():
                checkbox_names.append(child.cget("text"))
                checkbox_values.append(child.get())
        return checkbox_names, checkbox_values

    def select_file(self):
        previous_file = self.fetchfilerecord()
        if (len(previous_file) != 0) or (previous_file != []):
            cursor_1.execute("DELETE FROM FILELOG WHERE Filename=?", [previous_file[0][0],])
            tempdata.commit()

        filetypes = (
            ("Excel files", "*.xlsx"), 
            ("Excel files", "*.xls"), 
            ("all files", "*.*")
        )

        filename = fd.askopenfilename(
            title='Open a file',
            initialdir='/',
            filetypes= filetypes)
        
        now_record = self.fetchfilerecord()
        if len(now_record) == 0:
            cursor_1.execute("INSERT INTO FILELOG VALUES(?,?)", [filename, f"{date.today()}"])
            tempdata.commit()
        
        if (self.fetchfilerecord()[0][0] != ""):
            if (".xlsx" in self.fetchfilerecord()[0][0]) or (".xls" in self.fetchfilerecord()[0][0]):
                self.link_file.configure(fg_color="#059142", hover_color="#03692f", text=f"{Path(self.fetchfilerecord()[0][0]).name}")
            else:
                self.link_file.configure(fg_color="#F05316", hover_color="#b52d21")
        else:
            self.link_file.configure(fg_color="#1f538d", hover_color="#14375e", text="Open Excel file..")
        

    def incr_chr(self, c):
        return chr(ord(c) + 1) if c != 'Z' else 'A'

    def incr_str(self, s):
        lpart = s.rstrip('Z')
        num_replacements = len(s) - len(lpart)
        new_s = lpart[:-1] + self.incr_chr(lpart[-1]) if lpart else 'A'
        new_s += 'A' * num_replacements
        return new_s

    def increment_column(self, str):
        if (len(str) < 2):
            if (str != 'Z'):
                new_char = self.incr_chr(str)
                return new_char
            elif (str == 'Z'):
                new_str = self.incr_str(str)
                return new_str
        else:
            increase = self.incr_str(str)
            return increase

    def check_for_available_column(self, sheet, row):
        if row >= 2:
            counter = "B"
        else:
            counter = "A"

        while True:
            sheet_cell = sheet[f"{counter}{row}"]
            if sheet_cell.value == None:
                return counter
            else:
                counter = self.increment_column(counter)

    def open_excel(self):
        fetch_file_logs = self.fetchfilerecord()
        selected_file = fetch_file_logs[0][0]

        if os.path.isfile(selected_file):
            if (".xlsx" in selected_file) or (".xls" in selected_file):
                workbook_obj = openpyxl.load_workbook(selected_file)
                sheet_obj = workbook_obj.active
                
                return selected_file, sheet_obj, workbook_obj
            else:
                return "not an excel file", "not an excel file", "not an excel file"
        else:
            return "not an excel file", "not an excel file", "not an excel file"

    def check_name_column(self, active_sheet, reference, file):
        ask_for_name_update = False
        for data in reference:
            if file in data:
                ask_for_name_update = True

        if (active_sheet["A3"].value != None) or (active_sheet["A3"].value != ""):
            if ask_for_name_update:
                if (messagebox.askyesno(title="AKASHIC", message="Students are being updated. Proceed?")):
                    return True
                else:
                    return False
            else:
                return True
        else:
            return True
    
    def check_date_row(self, _sheet, data):
        column = "B"

        while True:
            if _sheet[f"{column}2"].value == None:
                break
            elif _sheet[f"{column}2"].value == data:
                break
            else:
                column = self.increment_column(column)

        if (_sheet[f"{column}2"].value != None) or (_sheet[f"{column}2"].value != ""):
            if _sheet[f"{column}2"].value == data:
                if (messagebox.askyesno(title="AKASHIC", message=f"Do you want to update records of '{data}'?")):
                    return column, True
                else:
                    return column, False
            else:
                return column, False
        else:
            return column, False


    def write_in_excel(self):
        filename, sheet, book = self.open_excel()
        if filename != "not an excel file":
            input_date = self.date_entry.get()
            date_column, update_or_not = self.check_date_row(sheet, input_date)


            attendance_column = self.check_for_available_column(sheet, 3)
            acquired_names, acquired_attendance = self.get_checkbox_values()

            previous_dataset = self.fetchnamerec()
            
            if (self.check_name_column(sheet, previous_dataset, filename)):
                sheet["B1"].value = f"Attendance Record for {self.course_entry.get()}"
                sheet["B1"].alignment = Alignment(horizontal="center")
                sheet.merge_cells("B1:G1")
                sheet["I1"].value = f"Instructor: {self.instructor_name_entry.get()}"
                for num in range(len(acquired_names)):
                    cell = sheet[f"A{num + 3}"]
                    cell.value = acquired_names[num]

            is_new_file = True
            for stored in previous_dataset:
                if input_date in stored:
                    is_new_file = False
                    break
            
            if is_new_file:
                sheet[f"{date_column}2"].value = input_date
                for item in range(len(acquired_attendance)):
                    sheet[f"{date_column}{item + 3}"].alignment = Alignment(horizontal="center")
                    if acquired_attendance[item] == "present":
                        sheet[f"{date_column}{item + 3}"].value = "/"
                    elif acquired_attendance[item] == "absent":
                        sheet[f"{date_column}{item + 3}"].value = "A"
                # Message
            else:
                if update_or_not:
                    sheet[f"{date_column}2"].value = input_date
                    for item in range(len(acquired_attendance)):
                        sheet[f"{date_column}{item + 3}"].alignment = Alignment(horizontal="center")
                        if acquired_attendance[item] == "present":
                            sheet[f"{date_column}{item + 3}"].value = "/"
                        elif acquired_attendance[item] == "absent":
                            sheet[f"{date_column}{item + 3}"].value = "A"
                else:
                    sheet[f"{attendance_column}2"].value = input_date
                    for item in range(len(acquired_attendance)):
                        sheet[f"{attendance_column}{item + 3}"].alignment = Alignment(horizontal="center")
                        if acquired_attendance[item] =="present":
                            sheet[f"{attendance_column}{item + 3}"].value = "/"
                        elif acquired_attendance[item] =="absent":
                            sheet[f"{attendance_column}{item + 3}"].value = "A"
            
            if update_or_not == False:
                cursor_1.execute("INSERT INTO RECORDDATE VALUES(?,?,?)", [input_date, date_column, filename])
                tempdata.commit()

            book.save(filename)

            print("DONE")
        else:
            messagebox.showerror(title="AKASHIC", message="Please select an excel file before generating a report")
        
    
    # View First Panel (Masterlist)
    def masterlist(self):
        self.terminal_tree.grid()
        self.masterlist_frame.grid()
        self.edit_masterlist_frame1.grid()
        self.edit_masterlist_frame2.grid()
        self.masterlLabel_frame.grid()

        self.attendance_frame.grid_remove()
        self.attendance_tool_frame.grid_remove()
        self.notebook.grid_remove()
        self.link_file.grid_remove()
        self.generate_report_button.grid_remove()

        self.sidebar_button_1.configure(fg_color="#14375e")
        self.sidebar_button_2.configure(fg_color="#1f538d", hover_color="#14375e")

    # View Second Panel (Checklist)
    def records(self):
        self.terminal_tree.grid_remove()
        self.masterlist_frame.grid_remove()
        self.edit_masterlist_frame1.grid_remove()
        self.edit_masterlist_frame2.grid_remove()
        self.masterlLabel_frame.grid_remove()

        self.attendance_frame.grid()
        self.attendance_tool_frame.grid()
        self.notebook.grid()
        self.link_file.grid()
        self.generate_report_button.grid()

        self.sidebar_button_1.configure(fg_color="#1f538d", hover_color="#14375e")
        self.sidebar_button_2.configure(fg_color="#14375e")

        # Generate checklist
        self.student_roll =  self.fetchdb()
        self.student_rows = -1

        self.attendance_roll = []
        for stnum, name, section, space, status in self.student_roll:
            self.student_rows += 1
            student = customtkinter.CTkCheckBox(self.attendance_frame, text=f"  {name}", border_color="red", border_width=1, onvalue="present", offvalue="absent")
            student.grid(row=self.student_rows, column=0, padx=(15, 0), pady=(0, 15), sticky="w")
            st_num = customtkinter.CTkLabel(self.attendance_frame, text=f"{stnum}")
            st_num.grid(row=self.student_rows, column=1, padx=(0, 30), pady=(0, 15), sticky="w")
            st_section = customtkinter.CTkLabel(self.attendance_frame, text=f"{section}")
            st_section.grid(row=self.student_rows, column=2, padx=(5, 50), pady=(0, 15), sticky="w")
            empty_desc1 = customtkinter.CTkLabel(self.attendance_frame, text=f"--")
            empty_desc1.grid(row=self.student_rows, column=3, padx=(15, 50), pady=(0, 15), sticky="nsew")
            self.attendance_roll.append(student)
            self.attendance_roll.append(st_num)
            self.attendance_roll.append(st_section)
            self.attendance_roll.append(empty_desc1)
    
    def clear_entry(self):
        self.name_entry.delete(0, END)
        self.section_entry.delete(0, END)
        self.stnum_entry.delete(0, END)
        self.status_option.set(self.status_option._values[0])
    
    def add_student(self):
        if (self.name_entry.get()=="" or self.section_entry.get()=="" or self.stnum_entry.get()==""):
            messagebox.showerror(title="Error", message="Please complete the form to proceed")
        else:
            student_data = [self.stnum_entry.get(), self.name_entry.get(), self.section_entry.get(), "", self.status_option.get()]
            if (messagebox.askyesno(title="AKASHIC", message="Student Numbers will be unchangeable. Do you want to create this profile?")):
                cursor.execute("INSERT INTO ATTENDANCE VALUES(?,?,?,?,?)", student_data)
                databs.commit()
                messagebox.showinfo(title="AKASHIC", message="Student has been listed")
                self.clear_entry()
                self.display_data_treeview()
    
    # Fetch data from data bases
    def fetchdb(self):
        cursor.execute("SELECT * FROM ATTENDANCE")
        datalist = cursor.fetchall()
        return datalist
    
    def fetchupdates(self):
        cursor_1.execute("SELECT * FROM UPDATES")
        updatelist = cursor_1.fetchall()
        return updatelist
    
    def fetchfilerecord(self):
        cursor_1.execute("SELECT * FROM FILELOG")
        updatelist = cursor_1.fetchall()
        return updatelist
    
    def fetchnamerec(self):
        cursor_1.execute("SELECT * FROM RECORDDATE")
        updatelist = cursor_1.fetchall()
        return updatelist
    
    def display_data_treeview(self):
        self.terminal_tree.delete(*self.terminal_tree.get_children())
        for item in self.fetchdb():
            self.terminal_tree.insert("", END, values=item)
        
        summary_list = self.fetchdb()
        self.total_students = 0
        self.total_regular = 0
        self.total_irregular = 0
        self.total_transferee = 0
        for data in summary_list:
            if 'Regular' in data:
                self.total_students += 1
                self.total_regular += 1
            elif 'Irregular' in data:
                self.total_students += 1
                self.total_irregular += 1
            elif 'Transferee' in data:
                self.total_students += 1
                self.total_transferee += 1
        
        self.total_numeric.configure(text = self.total_students)
        self.regular_numeric.configure(text=self.total_regular)
        self.irreg_numeric.configure(text=self.total_irregular)
        self.transferee_numeric.configure(text=self.total_transferee)

    def get_focused_data(self):
        self.selected_row = self.terminal_tree.focus()
        self.treeview_data = self.terminal_tree.item(self.selected_row)
        self.rows = self.treeview_data["values"]
        return self.rows

    def convert_int_to_str(self, value):
        new_list = []
        for item in value:
            new_list.append(str(item))
        return new_list

    def delete_student(self):
        self.delete_item = self.get_focused_data()
        if (self.delete_item != ""):
            self.convert_list = self.convert_int_to_str(self.delete_item)
            if (messagebox.askyesno(title="AKASHIC", message=f"Student '{self.convert_list[1]}' will be deleted. Do want to proceed?")) == True:
                cursor.execute("DELETE FROM ATTENDANCE WHERE StudentNum=?", [self.convert_list[0],])
                databs.commit()
                self.display_data_treeview()
        else:
            messagebox.showwarning(title="AKASHIC", message="Tip: Click on an item you want to delete on the table above")

    def selection_sort_data(self, itemsList):
        n = len(itemsList)
        for i in range(n - 1): 
            minValueIndex = i

            for j in range( i + 1, n ):
                if itemsList[j] < itemsList[minValueIndex] :
                    minValueIndex = j

            if minValueIndex != i :
                temp = itemsList[i]
                itemsList[i] = itemsList[minValueIndex]
                itemsList[minValueIndex] = temp

        return itemsList
    
    # Implemented Selection Sort in sorting entries
    def sort_data_entries(self):
        self.unordered_dataset = self.fetchdb()
        self.key_unordered_data = []
        for data in self.unordered_dataset:
            for specific in data:
                if specific == data[1]:
                    self.key_unordered_data.append(specific)

        self.key_ordered_data = self.selection_sort_data(self.key_unordered_data)
        self.ordered_dataset = []
        for item in self.key_ordered_data:
            for group in self.unordered_dataset:
                if item in group:
                    self.ordered_dataset.append(group)

        # Order permanence
        cursor.execute("DELETE FROM ATTENDANCE")
        cursor.execute("VACUUM")
        databs.commit()
        self.terminal_tree.delete(*self.terminal_tree.get_children())

        for item in self.ordered_dataset:
            cursor.execute("INSERT INTO ATTENDANCE VALUES(?,?,?,?,?)", item)
            databs.commit()
            self.terminal_tree.insert("", END, values=item)

    def update_panel(self):
        self.update_item = self.get_focused_data()
        if (self.update_item != ""):
            self.summary_details.grid_remove()
            self.update_panel_frame.grid()

            # Disable all buttons except update
            self.add_button.configure(state="disabled", fg_color="#059142")
            self.clear_button.configure(state="disabled", fg_color="#14375e")
            self.sort_button.configure(state="disabled", fg_color="#14375e")
            self.delete_button.configure(state="disabled")

            self.name_entry.configure(state="disabled", fg_color="#2b2c2e")
            self.stnum_entry.configure(state="disabled", fg_color="#2b2c2e")
            self.section_entry.configure(state="disabled", fg_color="#2b2c2e")
            self.status_option.configure(state="disabled", fg_color="#14375e")

            self.update_stnum_entry.configure(state="normal")

            if self.update_item != [self.update_stnum_entry.get(), self.update_name_entry.get(), self.update_section_entry.get(), '', self.update_status_option.get()]:
                self.update_details = self.fetchupdates()

                # If focus is changed, remove temporary data from database
                if (self.update_details != []) or (self.update_details != None):
                    cursor_1.execute("DELETE FROM UPDATES WHERE Row=?", [1,])
                    tempdata.commit()
                self.remove_view_content()
                self.update_name_entry.insert(0, self.update_item[1])
                self.update_stnum_entry.insert(0, self.update_item[0])
                self.update_section_entry.insert(0, self.update_item[2])
                self.update_status_option.set(self.update_item[4])

                self.update_stnum_entry.configure(state="disabled", fg_color="#181818")
                
                # Store temporary data every attempt of update
                cursor_1.execute("INSERT INTO UPDATES VALUES(?,?,?,?,?,?)", [1, self.update_stnum_entry.get(), self.update_name_entry.get(), self.update_section_entry.get(), '', self.update_status_option.get()])
                tempdata.commit()

        else:
            messagebox.showwarning(title="AKASHIC", message="Tip: Click on an item you want to update on the table above")
    
    def confirm_update(self):
        self.updated_entries = [self.update_name_entry.get(), self.update_section_entry.get(), '', self.update_status_option.get(), self.update_stnum_entry.get()]
        self.get_temp_data = self.fetchupdates()
        self.previous_data = list(self.get_temp_data[0][2:]) + [self.get_temp_data[0][1],]

        if self.updated_entries == self.previous_data:
            messagebox.showinfo(title="AKASHIC", message="No changes have been made")
            self.cancel_update()
        else:
            messagebox.showinfo(title="AKASHIC", message="Profile has been updated")
            cursor.execute("UPDATE ATTENDANCE SET Name=?, CourseYS=?, Space=?, Status=? WHERE StudentNum=?", self.updated_entries)
            databs.commit()
            self.cancel_update()
            self.display_data_treeview()
        
        cursor_1.execute("DELETE FROM UPDATES WHERE Row=?", [1,])
        tempdata.commit()


    def cancel_update(self):
        self.update_panel_frame.grid_remove()
        self.summary_details.grid()
        self.remove_view_content()

        # Reenable other button features
        self.add_button.configure(state="normal", fg_color="#059142" ,hover_color="#03692f")
        self.clear_button.configure(state="normal", fg_color="#1f538d")
        self.sort_button.configure(state="normal", fg_color="#1f538d")
        self.delete_button.configure(state="normal")

        self.name_entry.configure(state="normal", fg_color="#343638")
        self.stnum_entry.configure(state="normal", fg_color="#343638")
        self.section_entry.configure(state="normal", fg_color="#343638")
        self.status_option.configure(state="normal", fg_color="#1f538d")

    def remove_view_content(self):
        self.update_name_entry.delete(0, END)
        self.update_section_entry.delete(0, END)
        self.update_stnum_entry.delete(0, END)
        self.update_status_option.set(self.status_option._values[0])

    def open_input_dialog_event(self):
        dialog = customtkinter.CTkInputDialog(text="Type in a number:", title="CTkInputDialog")
        print("CTkInputDialog:", dialog.get_input())

    def change_appearance_mode_event(self, new_appearance_mode: str):
        customtkinter.set_appearance_mode(new_appearance_mode)

    def change_scaling_event(self, new_scaling: str):
        new_scaling_float = int(new_scaling.replace("%", "")) / 100
        customtkinter.set_widget_scaling(new_scaling_float)


# Data bases
databs = sqlite3.connect("Course_Attendance.db", isolation_level=None)
cursor = databs.cursor()
cursor.execute("CREATE TABLE IF NOT EXISTS ATTENDANCE (StudentNum Integer, Name Text, CourseYS Text, Space Text, Status Text)")

tempdata = sqlite3.connect("Updates.db")
cursor_1 = tempdata.cursor()
cursor_1.execute("CREATE TABLE IF NOT EXISTS UPDATES (Row Integer, StudentNum Integer, Name Text, CourseYS Text, Space Text, Status Text)")
cursor_1.execute("CREATE TABLE IF NOT EXISTS RECORDDATE (Date Text, Column Text, Filename Text)")
cursor_1.execute("CREATE TABLE IF NOT EXISTS FILELOG (Filename Text, DateAccessed Text)")


# Main
if __name__ == "__main__":
    app = App()
    app.mainloop()