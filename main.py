import tkinter
from tkinter import messagebox
from tkinter import ttk
from tkinter import *
import customtkinter
import sqlite3 

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
        self.sidebar_button_1 = customtkinter.CTkButton(self.sidebar_frame, text="Records", command=self.masterlist)
        self.sidebar_button_1.grid(row=1, column=0, padx=20, pady=10)
        self.sidebar_button_2 = customtkinter.CTkButton(self.sidebar_frame, text="Schedule", command=self.records)
        self.sidebar_button_2.grid(row=2, column=0, padx=20, pady=10)
        self.sidebar_button_3 = customtkinter.CTkButton(self.sidebar_frame, text="Import", command=self.get_focused_data)
        self.sidebar_button_3.grid(row=3, column=0, padx=20, pady=40)
        self.appearance_mode_label = customtkinter.CTkLabel(self.sidebar_frame, text="Appearance Mode:", anchor="w")
        self.appearance_mode_label.grid(row=5, column=0, padx=20, pady=(10, 0))
        self.appearance_mode_optionemenu = customtkinter.CTkOptionMenu(self.sidebar_frame, values=["Dark", "Light", "System"], command=self.change_appearance_mode_event)
        self.appearance_mode_optionemenu.grid(row=6, column=0, padx=20, pady=(10, 10))
        self.scaling_label = customtkinter.CTkLabel(self.sidebar_frame, text="UI Scaling:", anchor="w")
        self.scaling_label.grid(row=7, column=0, padx=20, pady=(10, 0))
        self.scaling_optionemenu = customtkinter.CTkOptionMenu(self.sidebar_frame, values=["80%", "90%", "100%", "110%", "120%"], command=self.change_scaling_event)
        self.scaling_optionemenu.grid(row=8, column=0, padx=20, pady=(10, 20))


        # Frames for Label and Layout purposes
        self.masterlist_frame = customtkinter.CTkFrame(self)
        self.masterlist_frame.grid(row=0, column=1, padx=(15, 15), pady=(12, 0), columnspan=2, sticky="nsew")
        self.masterlist_frame.grid_columnconfigure(0, weight=1)

        self.masterlLabel_frame = customtkinter.CTkFrame(self)
        self.masterlLabel_frame.grid(row=1, column=1, padx=(15, 15), pady=(10, 0), columnspan=2, sticky="nsew")
        self.masterlLabel_frame.grid_columnconfigure(0, weight=1)

        self.ml_label = customtkinter.CTkLabel(self.masterlist_frame, text="MASTERLIST\nCOURSE YEAR SECTION - CMPE101")
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

        self.add_button = customtkinter.CTkButton(self.edit_masterlist_frame1, text="Add", command=self.add_student)
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
        self.update_button = customtkinter.CTkButton(self.edit_masterlist_frame2, text="Sort", width=90, command=self.sort_data_entries)
        self.update_button.grid(row=1, column=0, padx=(15, 5), pady=13, sticky="w")

        self.update_button = customtkinter.CTkButton(self.edit_masterlist_frame2, text="Update", width=90, command=self.update_panel)
        self.update_button.grid(row=1, column=2, padx=(5, 5), pady=13, sticky="e")

        self.delete_button = customtkinter.CTkButton(self.edit_masterlist_frame2, text="Delete", fg_color= "dark red", width=90, command=self.delete_student)
        self.delete_button.grid(row=1, column=3, padx=(5, 15), pady=13, sticky="e")

        # Text Summary of Treeview Data
        self.summary_details = customtkinter.CTkFrame(self.edit_masterlist_frame2)
        self.summary_details.grid(row=2, column=0, padx=(15, 15), pady=(10, 0), columnspan=4, sticky="nsew")
        self.edit_masterlist_frame2.grid_columnconfigure(0, weight=1)

        self.summary_label = customtkinter.CTkLabel(self.summary_details, text="List Summary")
        self.summary_label.place(relx=0.5, rely=0.1, anchor="center") 

        self.number_of_students = customtkinter.CTkLabel(self.summary_details, text="Number of Students Enrolled: ")
        self.number_of_students.grid(row=1, column=0, padx=(15, 15), pady=(50, 0), sticky="nw")

        self.blank_space = customtkinter.CTkLabel(self.summary_details, text="                                           ")
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
        self.transferee_students.grid(row=4, column=0, padx=(15, 15), pady=(0, 15), sticky="nw")
        self.transferee_numeric = customtkinter.CTkLabel(self.summary_details, text="0")
        self.transferee_numeric.grid(row=4, column=3, padx=(15, 15), pady=(0, 15), sticky="nsew")  


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
        self.confirm_update_yes = customtkinter.CTkButton(self.update_panel_frame, text="Confirm", width=100)
        self.confirm_update_yes.grid(row=4, column=0, padx=(20, 10), pady=(0, 10), sticky="nsew")

        self.confirm_update_no = customtkinter.CTkButton(self.update_panel_frame, text="Cancel")
        self.confirm_update_no.grid(row=4, column=1, padx=(10, 20), pady=(0, 10), sticky="nsew")

        # self.sidebar_button_1.configure(state="disabled")

        # self.tree = ttk.Treeview(self.masterlist_frame, columns=(1, 2, 3, 4), show="headings")
        # self.tree.heading("1", text="No.")
        # self.tree.column("1", width=75)
        # self.tree.heading("2", text="Name")
        # self.tree.column("2", width=75)
        # self.tree.heading("3", text="Section")
        # self.tree.column("3", width=75)
        # self.tree.heading("4", text="Status")
        # self.tree.column("4", width=75)

        # # create main entry and button
        # self.entry = customtkinter.CTkEntry(self, placeholder_text="CTkEntry")
        # self.entry.grid(row=3, column=1, columnspan=2, padx=(20, 0), pady=(20, 20), sticky="nsew")

        # self.main_button_1 = customtkinter.CTkButton(master=self, fg_color="transparent", border_width=2, text_color=("gray10", "#DCE4EE"))
        # self.main_button_1.grid(row=3, column=3, padx=(20, 20), pady=(20, 20), sticky="nsew")

        # # create textbox
        # self.textbox = customtkinter.CTkTextbox(self, width=250)
        # self.textbox.grid(row=0, column=1, padx=(20, 0), pady=(20, 0), sticky="nsew")

        # # create tabview
        # self.tabview = customtkinter.CTkTabview(self, width=250)
        # self.tabview.grid(row=0, column=2, padx=(20, 0), pady=(20, 0), sticky="nsew")
        # self.tabview.add("CTkTabview")
        # self.tabview.add("Tab 2")
        # self.tabview.add("Tab 3")
        # self.tabview.tab("CTkTabview").grid_columnconfigure(0, weight=1)  # configure grid of individual tabs
        # self.tabview.tab("Tab 2").grid_columnconfigure(0, weight=1)

        # self.optionmenu_1 = customtkinter.CTkOptionMenu(self.tabview.tab("CTkTabview"), dynamic_resizing=False,
        #                                                 values=["Value 1", "Value 2", "Value Long Long Long"])
        # self.optionmenu_1.grid(row=0, column=0, padx=20, pady=(20, 10))
        # self.combobox_1 = customtkinter.CTkComboBox(self.tabview.tab("CTkTabview"),
        #                                             values=["Value 1", "Value 2", "Value Long....."])
        # self.combobox_1.grid(row=1, column=0, padx=20, pady=(10, 10))
        # self.string_input_button = customtkinter.CTkButton(self.tabview.tab("CTkTabview"), text="Open CTkInputDialog",
        #                                                    command=self.open_input_dialog_event)
        # self.string_input_button.grid(row=2, column=0, padx=20, pady=(10, 10))
        # self.label_tab_2 = customtkinter.CTkLabel(self.tabview.tab("Tab 2"), text="CTkLabel on Tab 2")
        # self.label_tab_2.grid(row=0, column=0, padx=20, pady=20)

        # # create radiobutton frame
        # self.radiobutton_frame = customtkinter.CTkFrame(self)
        # self.radiobutton_frame.grid(row=0, column=3, padx=(20, 20), pady=(20, 0), sticky="nsew")
        # self.radio_var = tkinter.IntVar(value=0)
        # self.label_radio_group = customtkinter.CTkLabel(master=self.radiobutton_frame, text="CTkRadioButton Group:")
        # self.label_radio_group.grid(row=0, column=2, columnspan=1, padx=10, pady=10, sticky="")
        # self.radio_button_1 = customtkinter.CTkRadioButton(master=self.radiobutton_frame, variable=self.radio_var, value=0)
        # self.radio_button_1.grid(row=1, column=2, pady=10, padx=20, sticky="n")
        # self.radio_button_2 = customtkinter.CTkRadioButton(master=self.radiobutton_frame, variable=self.radio_var, value=1)
        # self.radio_button_2.grid(row=2, column=2, pady=10, padx=20, sticky="n")
        # self.radio_button_3 = customtkinter.CTkRadioButton(master=self.radiobutton_frame, variable=self.radio_var, value=2)
        # self.radio_button_3.grid(row=3, column=2, pady=10, padx=20, sticky="n")

        # # create slider and progressbar frame
        # self.slider_progressbar_frame = customtkinter.CTkFrame(self, fg_color="transparent")
        # self.slider_progressbar_frame.grid(row=1, column=1, padx=(20, 0), pady=(20, 0), sticky="nsew")
        # self.slider_progressbar_frame.grid_columnconfigure(0, weight=1)
        # self.slider_progressbar_frame.grid_rowconfigure(4, weight=1)
        # self.seg_button_1 = customtkinter.CTkSegmentedButton(self.slider_progressbar_frame)
        # self.seg_button_1.grid(row=0, column=0, padx=(20, 10), pady=(10, 10), sticky="ew")
        # self.progressbar_1 = customtkinter.CTkProgressBar(self.slider_progressbar_frame)
        # self.progressbar_1.grid(row=1, column=0, padx=(20, 10), pady=(10, 10), sticky="ew")
        # self.progressbar_2 = customtkinter.CTkProgressBar(self.slider_progressbar_frame)
        # self.progressbar_2.grid(row=2, column=0, padx=(20, 10), pady=(10, 10), sticky="ew")
        # self.slider_1 = customtkinter.CTkSlider(self.slider_progressbar_frame, from_=0, to=1, number_of_steps=4)
        # self.slider_1.grid(row=3, column=0, padx=(20, 10), pady=(10, 10), sticky="ew")
        # self.slider_2 = customtkinter.CTkSlider(self.slider_progressbar_frame, orientation="vertical")
        # self.slider_2.grid(row=0, column=1, rowspan=5, padx=(10, 10), pady=(10, 10), sticky="ns")
        # self.progressbar_3 = customtkinter.CTkProgressBar(self.slider_progressbar_frame, orientation="vertical")
        # self.progressbar_3.grid(row=0, column=2, rowspan=5, padx=(10, 20), pady=(10, 10), sticky="ns")

        # # create scrollable frame
        # self.scrollable_frame = customtkinter.CTkScrollableFrame(self, label_text="CTkScrollableFrame")
        # self.scrollable_frame.grid(row=1, column=2, padx=(20, 0), pady=(20, 0), sticky="nsew")
        # self.scrollable_frame.grid_columnconfigure(0, weight=1)
        # self.scrollable_frame_switches = []
        # for i in range(100):
        #     switch = customtkinter.CTkSwitch(master=self.scrollable_frame, text=f"CTkSwitch {i}")
        #     switch.grid(row=i, column=0, padx=10, pady=(0, 20))
        #     self.scrollable_frame_switches.append(switch)

        # # create checkbox and switch frame
        # self.checkbox_slider_frame = customtkinter.CTkFrame(self)
        # self.checkbox_slider_frame.grid(row=1, column=3, padx=(20, 20), pady=(20, 0), sticky="nsew")
        # self.checkbox_1 = customtkinter.CTkCheckBox(master=self.checkbox_slider_frame)
        # self.checkbox_1.grid(row=1, column=0, pady=(20, 0), padx=20, sticky="n")
        # self.checkbox_2 = customtkinter.CTkCheckBox(master=self.checkbox_slider_frame)
        # self.checkbox_2.grid(row=2, column=0, pady=(20, 0), padx=20, sticky="n")
        # self.checkbox_3 = customtkinter.CTkCheckBox(master=self.checkbox_slider_frame)
        # self.checkbox_3.grid(row=3, column=0, pady=20, padx=20, sticky="n")

        # # set default values
        # self.checkbox_3.configure(state="disabled")
        # self.checkbox_1.select()
        # self.scrollable_frame_switches[0].select()
        # self.scrollable_frame_switches[4].select()
        # self.radio_button_3.configure(state="disabled")
        self.appearance_mode_optionemenu.set("Dark")
        self.scaling_optionemenu.set("100%")
        self.display_data_treeview()
        self.summary_details.grid_remove()
        # self.optionmenu_1.set("CTkOptionmenu")
        # self.combobox_1.set("CTkComboBox")
        # self.slider_1.configure(command=self.progressbar_2.set)
        # self.slider_2.configure(command=self.progressbar_3.set)
        # self.progressbar_1.configure(mode="indeterminnate")
        # self.progressbar_1.start()
        # self.textbox.insert("0.0", "CTkTextbox\n\n" + "Lorem ipsum dolor sit amet, consetetur sadipscing elitr, sed diam nonumy eirmod tempor invidunt ut labore et dolore magna aliquyam erat, sed diam voluptua.\n\n" * 20)
        # self.seg_button_1.configure(values=["CTkSegmentedButton", "Value 2", "Value 3"])
        # self.seg_button_1.set("Value 2")


    def masterlist(self):
        self.terminal_tree.grid()
        self.masterlist_frame.grid()
        self.edit_masterlist_frame1.grid()
        self.edit_masterlist_frame2.grid()
        self.masterlLabel_frame.grid()

    def records(self):
        self.terminal_tree.grid_remove()
        self.masterlist_frame.grid_remove()
        self.edit_masterlist_frame1.grid_remove()
        self.edit_masterlist_frame2.grid_remove()
        self.masterlLabel_frame.grid_remove()
    
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
            cursor.execute("INSERT INTO ATTENDANCE VALUES(?,?,?,?,?)", student_data)
            databs.commit()
            messagebox.showinfo(title="AKASHIC", message="Student has been listed")
            self.clear_entry()
            self.display_data_treeview()
    
    def fetchdb(self):
        cursor.execute("SELECT * FROM ATTENDANCE")
        datalist = cursor.fetchall()
        return datalist
    
    def display_data_treeview(self):
        self.terminal_tree.delete(*self.terminal_tree.get_children())
        for item in self.fetchdb():
            self.terminal_tree.insert("", END, values=item)
        
        self.summary_list = self.fetchdb()
        self.total_students = 0
        self.total_regular = 0
        self.total_irregular = 0
        self.total_transferee = 0
        for data in self.summary_list:
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

    def int_to_str(self, value):
        new_list = []
        for item in value:
            new_list.append(str(item))
        return new_list

    def delete_student(self):
        self.delete_item = self.get_focused_data()
        if (self.delete_item != ""):
            self.convert_list = self.int_to_str(self.delete_item)
            if (messagebox.askyesno(title="AKASHIC", message=f"Student {self.convert_list[1]} will be deleted. Do want to proceed?")) == True:
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

        self.terminal_tree.delete(*self.terminal_tree.get_children())
        for item in self.ordered_dataset:
            self.terminal_tree.insert("", END, values=item)

    def update_panel(self):
        self.update_item = self.get_focused_data()
        if (self.update_item != ""):
            self.update_name_entry.insert(0, self.update_item[1])
            self.update_stnum_entry.insert(0, self.update_item[0])
            self.update_section_entry.insert(0, self.update_item[2])
            self.update_status_option.set(self.update_item[4])
        else:
            messagebox.showwarning(title="AKASHIC", message="Tip: Click on an item you want to update on the table above")

    def open_input_dialog_event(self):
        dialog = customtkinter.CTkInputDialog(text="Type in a number:", title="CTkInputDialog")
        print("CTkInputDialog:", dialog.get_input())

    def change_appearance_mode_event(self, new_appearance_mode: str):
        customtkinter.set_appearance_mode(new_appearance_mode)

    def change_scaling_event(self, new_scaling: str):
        new_scaling_float = int(new_scaling.replace("%", "")) / 100
        customtkinter.set_widget_scaling(new_scaling_float)


databs = sqlite3.connect("Course_Attendance.db")
cursor = databs.cursor()
cursor.execute("CREATE TABLE IF NOT EXISTS ATTENDANCE (StudentNum Integer, Name Text, CourseYS Text, Space Text, Status Text)")


if __name__ == "__main__":
    app = App()
    app.mainloop()