import customtkinter as ctk
import pandas as pd
import tkinter as tk
from tkinter import filedialog
import os
import getpass  # For getting the OS user name
from datetime import datetime  # For getting the current date
import locale
import calendar
import re

ctk.set_appearance_mode("Dark") 
ctk.set_default_color_theme("green")

locale.setlocale(locale.LC_TIME, 'ro_RO') #FORMAT ROMANIA DATA


#######################################################################################
# CLASA INTERFATA
#######################################################################################

class FileConverterApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("ESCOBAR RAPCOS")
        self.geometry("1310x600")
        self.checkbox_states = {}


    ###########################################
    #Definire frame TOP
    ###########################################

        # Top frame for buttons and user info
        self.top_frame = ctk.CTkFrame(self)
        self.top_frame.pack(pady=10, padx=10, fill="x")

        # Left section of the top frame for buttons
        self.top_left_section = ctk.CTkFrame(self.top_frame)
        self.top_left_section.pack(side="left", anchor="n", padx=5)

        # Right section of the top frame for user info and dates
        self.top_right_section = ctk.CTkFrame(self.top_frame)
        self.top_right_section.pack(side="left", fill="both", expand=True, padx=5)

        #right section has 3 sides
        #left side
        self.top_right_section_left = ctk.CTkFrame(self.top_right_section)
        self.top_right_section_left.pack(side="left", fill="both", expand=True, padx=5)

        #right side
        self.top_right_section_right = ctk.CTkFrame(self.top_right_section)
        self.top_right_section_right.pack(side="left", fill="both", expand=True, padx=5)


    ###########################################
    #Definire frame BOT
    ###########################################

        # Bottom frame for scroll canvas and buttons
        self.bot_frame = ctk.CTkFrame(self)
        self.bot_frame.pack(pady=10, padx=10, fill="x")

        #Bottom frame left
        self.bot_frame_left=ctk.CTkFrame(self.bot_frame)
        self.bot_frame_left.pack(side="left", anchor="n", padx=5)

        #Bottom frame mid
        self.bot_frame_mid=ctk.CTkFrame(self.bot_frame)
        self.bot_frame_mid.pack(side="left", anchor="n", padx=5)

        #Bottom frame right
        self.bot_frame_right=ctk.CTkFrame(self.bot_frame)
        self.bot_frame_right.pack(side="left", fill="both", padx=5)  


    ###########################################
    #Definire butoane
    ###########################################

        # Folder path selection button
        self.folder_path_button = ctk.CTkButton(self.top_left_section, text="Click aici pentru selectie folder baze de date", command=self.select_folder, width=500, height=50)
        self.folder_path_button.pack(pady=2, fill="x")

        # Load file names button
        self.load_files_btn = ctk.CTkButton(self.top_left_section, text="Incarca numele bazelor de date", state="disabled", fg_color="green", command=self.load_files, width=500, height=50)
        self.load_files_btn.pack(pady=2, fill="x")

        # Placeholders for "Luna raport"
        self.luna_raport_label = ctk.CTkLabel(self.top_right_section_left, text="Luna raport:")
        self.luna_raport_label.pack(padx=10,pady=2,side="top", anchor="w")

        # Placeholders for "Data rezerva"
        self.data_rezerva_label = ctk.CTkLabel(self.top_right_section_left, text="Data rezerva:")
        self.data_rezerva_label.pack(padx=10,pady=2,side="top", anchor="w")
        
        # Button convert
        self.convert_btn = ctk.CTkButton(self.bot_frame_mid, text="CONVERT", command=self.convert_selected_files, width=100, height=50)
        self.bot_frame_mid.grid_columnconfigure(0, weight=1)  # This configures the column to expand and fill the space
        self.convert_btn.grid(row=1, column=1)  # Place the button in the first row and column

        # Button switch merge
        switch_var = ctk.StringVar(value="on")
        self.switch_1 = ctk.CTkSwitch(master=self.bot_frame_mid, text="Concatenare rapoarte",variable=switch_var, onvalue="on", offvalue="off")
        self.switch_1.grid(row=2,column=1)

    ###########################################
    #Initializare CONTAINERE principale
    ###########################################
        
        # User info and current date
        self.setup_user_info()

        # Scrollable container for files
        self.setup_files_container()

        #buton conversie mijloc
        self.setup_files_btn_conversie()

        # Scrollable container for CSV files
        self.setup_CSV_files_container()

    ###########################################
    #FUNCTIE GENERARE USER NAME SI DATA CURENTA
    ###########################################
    def setup_user_info(self):
        user_name = getpass.getuser()
        current_date = datetime.now().strftime("%d-%B-%Y")

        # "User:" label
        ctk.CTkLabel(self.top_right_section_right, text="User: ").pack(side="top", anchor="e")
        # user_name label in red (bold style not applied due to limitations)
        ctk.CTkLabel(self.top_right_section_right, text=user_name, text_color="red").pack(side="top", anchor="e")

        # "Data curenta:" label
        ctk.CTkLabel(self.top_right_section_right, text="Data curenta: ").pack(side="top", anchor="e")
        # current_date label in red (bold style not applied)
        ctk.CTkLabel(self.top_right_section_right, text=current_date, text_color="red").pack(side="top", anchor="e")


    ###########################################
    #FUNCTIE CONTAINER FISIERE INITIALE (ST)
    ###########################################
    def setup_files_container(self):

        # Adjust the parent of the canvas and scrollbar to self.bot_frame
        self.canvas = tk.Canvas(self.bot_frame_left, bg="#2e2e2e", width=700, height=500)
        

        self.canvas.create_line(100, 100, 300, 200, fill="blue")
   
        self.scrollbar = tk.Scrollbar(self.bot_frame_left, orient="vertical", command=self.canvas.yview)

        header_frame = ctk.CTkFrame(self.bot_frame_left)
        header_frame.pack(fill="x")

        select_width = 10
        filename_width = 350
        filetype_width = 100
        #key_width = 50

        # Header labels with specific width settings
        ctk.CTkLabel(header_frame, text="Select", width=select_width).pack(side="left", padx=10)
        ctk.CTkLabel(header_frame, text="Nume fisier", width=filename_width).pack(side="left", padx=10)
        ctk.CTkLabel(header_frame, text="File_type", width=filetype_width).pack(side="left", padx=10)
        #ctk.CTkLabel(header_frame, text="KEY", width=key_width).pack(side="left", padx=10)

        # The initial_files frame will be inside the canvas, so it remains unchanged
        self.initial_files = ctk.CTkFrame(self.canvas)

        self.initial_files.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.create_window((0, 0), window=self.initial_files, anchor="nw")

        self.scrollbar = tk.Scrollbar(self.bot_frame_left, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        # Since these widgets are now children of self.bot_frame, their packing remains the same
        self.canvas.pack(side="left", fill="both", expand=False)
        self.scrollbar.pack(side="left", fill="y")

        #mouse bind
        self.canvas.bind("<Enter>", lambda e: self.bind_mousewheel(self.canvas))
        self.canvas.bind("<Leave>", lambda e: self.unbind_mousewheel(self.canvas))


    ###########################################
    #FUNCTIE CONTAINER BUTON CONVERSIE (MID)
    ###########################################
    def setup_files_btn_conversie(self):

        # Adjust the parent of the canvas and scrollbar to self.bot_frame
        self.canvas_buton = tk.Canvas(self.bot_frame_mid, bg="#2e2e2e", width=700, height=50)
        self.canvas_buton.create_line(25, 50, 50, 50, fill="green")
        self.canvas_buton.create_window((0, 0), anchor="nw")

    ###########################################
    #FUNCTIE CONTAINER FISIERE  CONVERTITE (DR)
    ###########################################
    def setup_CSV_files_container(self):
        # Adjust the parent of the canvas and scrollbar to self.bot_frame
        self.canvas_csv = tk.Canvas(self.bot_frame_right, bg="#2e2e2e", width=700, height=500)
        
        self.canvas_csv.create_line(100, 100, 600, 400, fill="red")
   
        self.scrollbar = tk.Scrollbar(self.bot_frame_right, orient="vertical", command=self.canvas_csv.yview)
        # The initial_files frame will be inside the canvas, so it remains unchanged
        self.initial_files_CSV = ctk.CTkFrame(self.canvas_csv)

        self.initial_files_CSV.bind("<Configure>", lambda e: self.canvas_csv.configure(scrollregion=self.canvas_csv.bbox("all")))
        self.canvas_csv.create_window((0, 0), window=self.initial_files_CSV, anchor="nw")
        self.canvas_csv.configure(yscrollcommand=self.scrollbar.set)


        # Since these widgets are now children of self.bot_frame, their packing remains the same
        self.canvas_csv.pack(side="left", fill="both", expand=False)
        self.scrollbar.pack(side="left", fill="y")

        self.canvas_csv.bind("<Enter>", lambda e: self.bind_mousewheel(self.canvas_csv))
        self.canvas_csv.bind("<Leave>", lambda e: self.unbind_mousewheel(self.canvas_csv))

    ###########################################
    #FUNCTIE BIND MOUSE
    ###########################################
    def bind_mousewheel(self, widget):
        widget.bind_all("<MouseWheel>", lambda e: self.on_mouse_scroll(e, widget))
        widget.bind_all("<Button-4>", lambda e: self.on_mouse_scroll(e, widget))
        widget.bind_all("<Button-5>", lambda e: self.on_mouse_scroll(e, widget))
    
    ###########################################
    #FUNCTIE UN_BIND MOUSE
    ###########################################   
    def unbind_mousewheel(self, widget):
        widget.unbind_all("<MouseWheel>")
        widget.unbind_all("<Button-4>")
        widget.unbind_all("<Button-5>")

    ###########################################
    #FUNCTIE SCROLL MOUSE
    ###########################################
    def on_mouse_scroll(self, event, widget):
        if event.delta:
            widget.yview_scroll(int(-1 * (event.delta / 120)), "units")
        else:
            if event.num == 4:
                widget.yview_scroll(-1, "units")
            elif event.num == 5:
                widget.yview_scroll(1, "units")

    ###########################################
    #FUNCTIE SELECTIE FOLDER - BUTON
    ###########################################
    def select_folder(self):
        folder_path = filedialog.askdirectory()
        if folder_path:
            self.folder_path = folder_path
            self.folder_path_button.configure(text=folder_path)
            self.load_files_btn.configure(state="normal")  # Enable button

    ###########################################
    #FUNCTIE CITIRE FILE EXTENSION
    ###########################################
    def get_file_extension(self,file_name):
        _, file_extension = os.path.splitext(file_name)
        return file_extension or 'Unknown'

    ###########################################
    #FUNCTIE INCARCARE NUME FISIERE (ST)
    ###########################################
    def load_files(self):

        for widget in self.initial_files.winfo_children():
            widget.destroy()  # Clear existing content

        select_width = 10
        filename_width = 350
        filetype_width = 100
        #key_width = 50

        files = [f for f in os.listdir(self.folder_path) if os.path.isfile(os.path.join(self.folder_path, f))]
        key="[x]"

        for file_name in files:
            # Create a new frame for each file row
            file_name_without_extension = os.path.splitext(file_name)[0]
            file_extension=self.get_file_extension(file_name)
            file_row = ctk.CTkFrame(self.initial_files)
            file_row.pack(fill="x")



            # Store checkbox state in a dictionary
            chk_state = ctk.BooleanVar()
            chk_box = ctk.CTkCheckBox(file_row, text="", width=select_width, variable=chk_state)
            chk_box.pack(side="left", padx=(10, 10))
            self.checkbox_states[file_name] = chk_state

            ctk.CTkLabel(file_row, text=file_name_without_extension, width=filename_width).pack(side="left", padx=(10, 10))
            ctk.CTkLabel(file_row, text=file_extension, width=filetype_width).pack(side="left", padx=(10, 10))

            # Update the canvas scroll region to account for the new file rows
            self.canvas.update_idletasks()  # Process all pending events which can change the size of widgets
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))

        month_to_num = {
            'ian': 1, 'feb': 2, 'mar': 3, 'apr': 4, 'mai': 5, 'iun': 6,
            'iul': 7, 'aug': 8, 'sep': 9, 'oct': 10, 'nov': 11, 'dec': 12
        }

        month_names_ro = {
            'ian': 'IANUARIE', 'feb': 'FEBRUARIE', 'mar': 'MARTIE',
            'apr': 'APRILIE', 'mai': 'MAI', 'iun': 'IUNIE',
            'iul': 'IULIE', 'aug': 'AUGUST', 'sep': 'SEPTEMBRIE',
            'oct': 'OCTOMBRIE', 'nov': 'NOIEMBRIE', 'dec': 'DECEMBRIE'
        }

        folder_name = os.path.basename(self.folder_path)
        #print(f"Folder Name: '{folder_name}'") 
        year_month_match = re.search(r'(\d+)\s+(\w+)\s+(\d{4})$', folder_name)

        if year_month_match:
            #print("this part is called")
            # Correcting the group indexes according to the new regex pattern
            year = int(year_month_match.group(3))
            month_abbr = year_month_match.group(2).lower()

            # Use the dictionary for month number; remove the redundant datetime.strptime call
            month_number = month_to_num.get(month_abbr.lower(), 0)
            month_full_name = month_names_ro.get(month_abbr, "UNKNOWN")

            # Calculate the last day of the month
            last_day = calendar.monthrange(year, month_number)[1]

            # Update labels
            self.luna_raport_label.configure(text=f"Luna rapoarte: {month_full_name} {year}")
            self.data_rezerva_label.configure(text=f"Data rezerva: {last_day} {month_full_name} {year}")

            self.csv_file_path_extracted = f"Baza CSV {month_full_name} {year}"

        else:
            print("Match not found")   


    ###########################################
    #FUNCTIE CONVERSI EXCEL TO CSV DOAR 1 SHEET
    ###########################################
    def convert_excel_to_csv_single(self, excel_file_path, csv_file_path):
        """
        Converts a single-sheet Excel file to CSV.
        """
        df = pd.read_excel(excel_file_path)
        df.to_csv(csv_file_path, index=False)
        print(f"Converted single-sheet Excel file to CSV: {os.path.basename(csv_file_path)}")

    ###########################################
    #FUNCTIE CONVERSI EXCEL TO CSV MULTI SHEET
    ###########################################


    def group_sheets_by_headers(self, excel_file_path):
        """
        Groups sheets by their headers.

        Returns a dictionary where each key is a tuple representing the headers of a group of sheets,
        and the value is a list of sheet names that have those headers.
        """
        xls = pd.ExcelFile(excel_file_path)
        grouped_sheets = {}

        for sheet_name in xls.sheet_names:
            df = pd.read_excel(excel_file_path, sheet_name=sheet_name)
            headers_tuple = tuple(df.columns.tolist())

            if headers_tuple not in grouped_sheets:
                grouped_sheets[headers_tuple] = [sheet_name]
            else:
                grouped_sheets[headers_tuple].append(sheet_name)

        return grouped_sheets

    def convert_excel_to_csv_by_group(self, excel_file_path, output_folder):
        """
        Converts Excel sheets to CSV files, grouping sheets by matching headers.
        Sheets with matching headers are merged into a single CSV.
        Unique sheets are exported to their own CSV.
        """
        grouped_sheets = self.group_sheets_by_headers(excel_file_path)
        group_number = 1

        for headers, sheet_names in grouped_sheets.items():
            if len(sheet_names) > 1:
                # Combine sheets with matching headers into a single CSV
                combined_df = pd.concat([pd.read_excel(excel_file_path, sheet_name=sheet_name) for sheet_name in sheet_names], ignore_index=True)
                combined_csv_path = os.path.join(output_folder, f"{os.path.splitext(os.path.basename(excel_file_path))[0]}_group{group_number}.csv")
                combined_df.to_csv(combined_csv_path, index=False)
                print(f"Combined sheets {', '.join(sheet_names)} into {os.path.basename(combined_csv_path)}")
                group_number += 1
            else:
                # Export unique sheets to their own CSV
                sheet_name = sheet_names[0]  # There's only one sheet in this group
                df = pd.read_excel(excel_file_path, sheet_name=sheet_name)
                csv_file_path = os.path.join(output_folder, f"{os.path.splitext(os.path.basename(excel_file_path))[0]}_{sheet_name}.csv")
                df.to_csv(csv_file_path, index=False)
                print(f"Exported sheet {sheet_name} to CSV: {os.path.basename(csv_file_path)}")


    
    ###########################################
    #FUNCTIE CONVERSIE ACESS TO CSV
    ###########################################    
            


    
    ###########################################
    #FUNCTIE CONVERSIE TOTAL
    ###########################################  

    def convert_selected_files(self):
        """
        Converts selected files to CSV based on their file type and checkbox state.
        Only files with a checked state are processed.
        """
        # Use os.path.normpath to ensure the path is normalized
        new_folder_path = os.path.normpath(os.path.join(self.folder_path, self.csv_file_path_extracted))
        os.makedirs(new_folder_path, exist_ok=True)  # Create the directory if it does not exist

        # Loop through each file and its checkbox state in the self.checkbox_states dictionary
        for file_name, chk_state in self.checkbox_states.items():
            # Check if the checkbox for this file is selected
            if chk_state.get():  # Returns True if the checkbox is checked
                file_path = os.path.join(self.folder_path, file_name)
                file_extension = os.path.splitext(file_name)[1].lower()
                #csv_file_path = os.path.splitext(file_path)[0] + '.csv'
                csv_file_path = os.path.join(new_folder_path, os.path.splitext(file_name)[0] + '.csv')


                # Process only if it's an Excel file; adjust as needed for other types
                if file_extension in ['.xlsx', '.xls']:
                    try:
                        xls = pd.ExcelFile(file_path)
                        if len(xls.sheet_names) > 1:
                            self.convert_excel_to_csv_by_group(file_path, new_folder_path)
                        else:
                            self.convert_excel_to_csv_single(file_path, csv_file_path)
                    except Exception as e:
                        print(f"Error converting {file_name}: {e}")
                # Add other file types and their respective conversion logic here as needed





if __name__ == "__main__":
    app = FileConverterApp()
    app.mainloop()


