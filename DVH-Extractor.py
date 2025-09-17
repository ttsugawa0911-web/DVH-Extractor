# --- Module Imports ---
# Import necessary libraries to build the application.

import tkinter  # Core module for creating GUI elements (windows, buttons, etc.)
from tkinter import ttk, filedialog, messagebox  # Advanced GUI widgets and dialog boxes
from pathlib import Path  # Module for handling file and folder paths easily, regardless of OS
import re  # Module for using regular expressions (complex string searching/replacing)
import csv  # Module for reading and writing CSV files
import threading  # Module for running time-consuming tasks in the background to prevent the app from freezing
from collections import defaultdict  # A dictionary subclass that provides a default value for a nonexistent key
import queue  # Module for safely exchanging data between threads
import json  # Module for saving and restoring Python data (like dictionaries) to a file

# --- Main Application Class ---
# This class acts as the blueprint for the application, defining its appearance and behavior.

class DvhConverterApp:
    # --- Initialization Method ---
    # This is the first method that is called when an object of the class is created.
    def __init__(self, root):
        # self: A reference to the current instance of the class.
        # root: The main window of the application.
        self.root = root
        
        # --- Define the path for the configuration file ---
        # Path.home() gets the user's home directory (e.g., C:\Users\YourName).
        # We create a file path named .dvh_converter_config.json.
        self.config_file = Path.home() / ".dvh_converter_config.json"
        
        # --- Basic Window Configuration ---
        root.title("DVH Data Extractor")  # Text displayed in the window's title bar
        root.geometry("600x400")  # Initial window size (width x height)
        root.columnconfigure(0, weight=1)  # Make the column expandable when the window is resized
        root.rowconfigure(0, weight=1)

        # --- Create the Main Frame ---
        # This frame serves as the base for placing other widgets.
        main_frame = ttk.Frame(root, padding=10)
        main_frame.grid(sticky=tkinter.NSEW)  # Expand the frame to fill the entire window
        main_frame.columnconfigure(1, weight=1)  # Make the column inside the frame expandable

        # --- Special Variables for GUI Interaction ---
        # When the content of these variables changes, the linked GUI elements automatically update.
        self.input_folder_path = tkinter.StringVar()  # String variable to store the input folder path
        self.output_folder_path = tkinter.StringVar()  # String variable to store the output folder path
        self.is_patient_wise = tkinter.BooleanVar(value=True)  # State of the "patient-wise" checkbox (ON/OFF)
        self.is_structure_wise = tkinter.BooleanVar(value=True)  # State of the "structure-wise" checkbox (ON/OFF)
        
        # --- Create and Place GUI Widgets ---
        # ttk.Label: Text label
        # ttk.Entry: Single-line text input box
        # ttk.Button: Clickable button
        # ttk.Combobox: Drop-down list
        # ttk.Checkbutton: Checkbox
        # .grid() is used to determine where to place the created widgets on the screen.

        # 1. Input Folder Selection Area
        ttk.Label(main_frame, text="1. Select the folder containing DVH text files").grid(column=0, row=0, columnspan=3, sticky=tkinter.W, pady=(0, 5))
        input_folder_box = ttk.Entry(main_frame, textvariable=self.input_folder_path, width=60)
        input_folder_box.grid(column=0, row=1, columnspan=2, sticky=tkinter.EW, padx=(0, 5))
        ttk.Button(main_frame, text="Browse...", command=self.ask_input_folder).grid(column=2, row=1)
        
        # 2. Output Folder Selection Area
        ttk.Label(main_frame, text="2. Select the destination folder for CSV files").grid(column=0, row=2, columnspan=3, sticky=tkinter.W, pady=(10, 5))
        output_folder_box = ttk.Entry(main_frame, textvariable=self.output_folder_path, width=60)
        output_folder_box.grid(column=0, row=3, columnspan=2, sticky=tkinter.EW, padx=(0, 5))
        ttk.Button(main_frame, text="Browse...", command=self.ask_output_folder).grid(column=2, row=3)

        # 3. Option Settings Area
        ttk.Label(main_frame, text="3. Configure Options").grid(column=0, row=4, columnspan=3, sticky=tkinter.W, pady=(10, 5))
        options_frame = ttk.Frame(main_frame)
        options_frame.grid(column=0, row=5, columnspan=3, sticky=tkinter.W)

        ttk.Label(options_frame, text="Sampling Interval:").pack(side=tkinter.LEFT, padx=(0, 5))
        self.order_comb = ttk.Combobox(options_frame, values=[0.1, 0.5, 1, 5, 10], state='readonly', width=5)
        self.order_comb.current(1)  # Set the initial value to the second item (0.5)
        self.order_comb.pack(side=tkinter.LEFT, padx=5)

        ttk.Label(options_frame, text="Dose Type:").pack(side=tkinter.LEFT, padx=(20, 5))
        self.type_comb = ttk.Combobox(options_frame, values=["%", "Gy"], state='readonly', width=5)
        self.type_comb.current(0)  # Set the initial value to the first item (%)
        self.type_comb.pack(side=tkinter.LEFT, padx=5)
        
        # 4. Output Format Selection Area
        ttk.Label(main_frame, text="4. Select Output Format (multiple allowed)").grid(column=0, row=6, columnspan=3, sticky=tkinter.W, pady=(10, 5))
        output_type_frame = ttk.Frame(main_frame)
        output_type_frame.grid(column=0, row=7, columnspan=3, sticky=tkinter.W)
        ttk.Checkbutton(output_type_frame, text="Create one CSV per patient", variable=self.is_patient_wise).pack(side=tkinter.LEFT)
        ttk.Checkbutton(output_type_frame, text="Create one CSV per structure", variable=self.is_structure_wise).pack(side=tkinter.LEFT, padx=20)

        # 5. Run Button and Progress Bar
        self.run_button = ttk.Button(main_frame, text="Run", command=self.run_conversion)
        self.run_button.grid(column=0, row=8, columnspan=3, pady=20)

        self.progress_var = tkinter.DoubleVar()  # Numeric variable to store progress (0-100)
        self.progress_bar = ttk.Progressbar(main_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.grid(column=0, row=9, columnspan=3, sticky=tkinter.EW, pady=10)
        self.status_label = ttk.Label(main_frame, text="Ready")
        self.status_label.grid(column=0, row=10, columnspan=3, sticky=tkinter.W)
        
        # --- Prepare for Inter-Thread Communication ---
        # Create a "pipe" to safely pass messages from the background thread to the GUI.
        self.thread_queue = queue.Queue()
        # Start a mechanism to periodically monitor this pipe.
        self.process_queue()
        
        # --- Load settings from the last session on startup ---
        self.load_settings()

    # --- Method to save settings to a file ---
    def save_settings(self):
        # Create a dictionary with the current input and output folder paths.
        settings = {
            'input_folder': self.input_folder_path.get(),
            'output_folder': self.output_folder_path.get()
        }
        try:
            # Open the settings file in write mode.
            with open(self.config_file, 'w') as f:
                # Write the dictionary to the file in JSON format.
                json.dump(settings, f, indent=4)
        except IOError:
            # If writing fails, print an error but don't stop the application.
            print("Failed to write to the settings file.")

    # --- Method to load settings from a file ---
    def load_settings(self):
        try:
            # Check if the configuration file exists.
            if self.config_file.exists():
                # Open the file in read mode.
                with open(self.config_file, 'r') as f:
                    # Read the data from the JSON file back into a dictionary.
                    settings = json.load(f)
                    # Set the paths in the GUI variables.
                    self.input_folder_path.set(settings.get('input_folder', ''))
                    self.output_folder_path.set(settings.get('output_folder', ''))
        except (json.JSONDecodeError, IOError):
            # If the file doesn't exist or is corrupted, do nothing.
            print("Failed to read the settings file.")
            
    # --- Method to monitor the queue ---
    # This processes messages received from the background thread.
    def process_queue(self):
        try:
            # Try to get one message from the queue (will raise an error if empty).
            message_type, message_content = self.thread_queue.get(block=False)
            
            # Display a dialog based on the message type (completion or error).
            if message_type == 'INFO':
                messagebox.showinfo("Completed", message_content)
            elif message_type == 'ERROR':
                messagebox.showerror("Error", message_content)
        except queue.Empty:
            # If the queue is empty, do nothing.
            pass
        finally:
            # Call this method again after 100 milliseconds to continuously monitor the queue.
            self.root.after(100, self.process_queue)

    # --- Method called when the "Run" button is pressed ---
    def run_conversion(self):
        # Get the paths entered by the user.
        input_dir = self.input_folder_path.get()
        output_dir = self.output_folder_path.get()

        # --- Input validation ---
        # Check for missing paths and invalid folders, and show an error if necessary.
        if not input_dir or not output_dir:
            messagebox.showerror("Error", "Please select both an input and an output folder.")
            return  # Exit the method
        if not Path(input_dir).is_dir() or not Path(output_dir).is_dir():
            messagebox.showerror("Error", "The specified path does not exist or is not a folder.")
            return
        if not self.is_patient_wise.get() and not self.is_structure_wise.get():
            messagebox.showerror("Error", "Please select at least one output format.")
            return

        # --- Prepare for the process to start ---
        self.run_button.config(state="disabled")  # Disable the button to prevent multiple clicks
        self.status_label.config(text="Processing...")  # Update the status label
        self.progress_var.set(0)  # Reset the progress bar

        # --- Start the background process ---
        # Run the conversion_thread method in a separate "thread" from the main GUI.
        # This prevents the GUI from freezing during the intensive task.
        thread = threading.Thread(target=self.conversion_thread, args=(input_dir, output_dir))
        thread.start()  # Start the thread
        
    # --- Main processing method run in the background thread ---
    def conversion_thread(self, input_dir, output_dir):
        # try...except...finally block:
        # It runs the code in the 'try' block. If an error (Exception) occurs, it jumps to 'except'.
        # The 'finally' block is always executed at the end, whether an error occurred or not.
        try:
            # Get the selected option values from the GUI.
            d_interval = float(self.order_comb.get())
            dose_type = self.type_comb.get()
            dose_col_index = 0 if dose_type == "%" else 1

            # 1. Parse all files
            self.status_label.config(text="Parsing files...")
            all_patients_data, txt_files = self.parse_folder(Path(input_dir), d_interval, dose_col_index)
            
            if not all_patients_data:
                raise ValueError("No .txt files found or no data was extracted.")

            # 2. Create CSV files
            total_tasks = self.is_patient_wise.get() + self.is_structure_wise.get()
            task_count = 0

            # If "patient-wise" is checked, call the corresponding CSV creation function.
            if self.is_patient_wise.get():
                task_count += 1
                self.status_label.config(text=f"({task_count}/{total_tasks}) Creating patient-wise CSVs...")
                self.write_patient_csvs(all_patients_data, Path(output_dir), dose_type, self.progress_var, 50.0 / total_tasks)

            # If "structure-wise" is checked, call the corresponding CSV creation function.
            if self.is_structure_wise.get():
                task_count += 1
                self.status_label.config(text=f"({task_count}/{total_tasks}) Creating structure-wise CSVs...")
                self.write_structure_csvs(all_patients_data, Path(output_dir), dose_type, self.progress_var, 50.0 / total_tasks * task_count)

            # --- Success ---
            # Save the folder settings for the next launch.
            self.save_settings()
            # Put a completion message into the queue (the GUI thread will pick it up and show a dialog).
            self.thread_queue.put(('INFO', 'The process has been completed successfully.'))

        except Exception as e:
            # --- Failure ---
            # Put an error message into the queue.
            self.thread_queue.put(('ERROR', f"An error occurred during processing:\n{e}"))
        finally:
            # --- After the process completes ---
            # Re-enable the button and update the status label, whether successful or not.
            self.run_button.config(state="normal")
            self.status_label.config(text="Done")
            self.progress_var.set(100)

    # --- Method for the "Browse..." button ---
    def ask_input_folder(self):
        # Open a folder selection dialog.
        path = filedialog.askdirectory()
        if path:  # If a folder was selected (not canceled)
            self.input_folder_path.set(path)  # Set the path to the GUI variable

    def ask_output_folder(self):
        path = filedialog.askdirectory()
        if path:
            self.output_folder_path.set(path)
            
    # --- Method to parse all text files in a folder ---
    def parse_folder(self, input_dir, d_interval, dose_col_index):
        all_patients_data = {}  # An empty dictionary to store all patients' data
        txt_files = list(input_dir.glob('*.txt'))  # List all .txt files in the folder
        
        if not txt_files:  # If no files are found, return empty data
            return {}, []
            
        # Loop through the list of files
        for i, file_path in enumerate(txt_files):
            # Call the method to parse each file individually
            patient_data = self.parse_dvh_file(file_path, d_interval, dose_col_index)
            if patient_data and 'Patient ID' in patient_data:
                # Store the parsing result in the dictionary, using the patient ID as the key
                all_patients_data[patient_data['Patient ID']] = patient_data
            # Update the progress bar
            self.progress_var.set((i + 1) / len(txt_files) * 50)
        
        return all_patients_data, txt_files

    # --- Method to parse a single text file ---
    def parse_dvh_file(self, file_path, d_interval, dose_col_index):
        patient_data = {'structures': {}}  # Dictionary to store data for one patient
        current_structure = None  # Variable to temporarily store the name of the structure being parsed

        # Open the file and read it line by line
        with open(file_path, 'r', encoding='utf-8-sig') as f:
            for line in f:
                line = line.strip()  # Remove leading/trailing whitespace
                if not line:  # Skip empty lines
                    continue
                
                # Use regex to split the line into a "key" and "value" (e.g., "Patient Name : hoge" -> "Patient Name", "hoge")
                parts = re.split(r'\s*:\s*|\s{2,}', line, 1)
                key = parts[0].strip()
                val = parts[1].strip() if len(parts) > 1 else ''

                # Store the data in the dictionary based on the key
                if key == 'Patient Name': patient_data['Patient Name'] = val
                elif key == 'Patient ID': patient_data['Patient ID'] = val
                elif key == 'Prescribed dose [Gy]': patient_data['Prescribed dose [Gy]'] = val
                elif key == 'Structure':
                    current_structure = val  # Record that a new structure section has started
                    patient_data['structures'][current_structure] = {'dvh': {}}
                elif current_structure:  # Process lines inside a structure section
                    if key in ['Volume [cm³]', 'Min Dose [%]', 'Max Dose [%]', 'Mean Dose [%]']:
                        patient_data['structures'][current_structure][key] = val
                    elif re.match(r'^[0-9.]+$', key):  # If the line starts with a number, assume it's DVH data
                        dvh_values = re.split(r'\s+', line.strip())
                        try:
                            dose = float(dvh_values[dose_col_index])
                            volume = float(dvh_values[-1])
                            # Extract only the data that matches the sampling interval
                            if abs(dose % d_interval) < 1e-9 or abs(dose % d_interval - d_interval) < 1e-9:
                                patient_data['structures'][current_structure]['dvh'][dose] = volume
                        except (ValueError, IndexError):
                            # Ignore lines that cannot be converted to numbers
                            pass
        return patient_data

    # --- Method to create patient-wise CSVs (wide format) ---
    def write_patient_csvs(self, all_patients_data, output_dir, dose_type, progress_var, progress_offset):
        dose_unit = "[%]" if dose_type == "%" else "[Gy]"

        # Loop through each patient's data
        for patient_id, data in all_patients_data.items():
            filename = output_dir / f"{patient_id}.csv"
            
            structures = data.get('structures', {})
            if not structures:  # Skip if there's no structure data
                continue

            # --- Prepare data for CSV writing ---
            all_doses = set()  # Set to collect all unique dose points
            structure_names = sorted(structures.keys())  # Sort structure names alphabetically

            # Loop through all structures to collect all dose points
            for str_name in structure_names:
                all_doses.update(structures[str_name].get('dvh', {}).keys())

            # Sort the collected dose points in ascending order
            sorted_doses = sorted(list(all_doses))

            # --- Start writing to the CSV file ---
            with open(filename, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)

                # 1. Write patient information
                writer.writerow(['Patient Name', data.get('Patient Name', '')])
                writer.writerow(['Patient ID', data.get('Patient ID', '')])
                writer.writerow(['Prescribed dose [Gy]', data.get('Prescribed dose [Gy]', '')])
                writer.writerow([])  # Empty row

                # 2. Write structure statistics (summary)
                stats_header = ['Structure Name', 'Volume [cm³]', 'Min Dose [%]', 'Max Dose [%]', 'Mean Dose [%]']
                writer.writerow(stats_header)
                for str_name in structure_names:
                    str_data = structures[str_name]
                    stats_row = [
                        str_name,
                        str_data.get('Volume [cm³]', ''),
                        str_data.get('Min Dose [%]', ''),
                        str_data.get('Max Dose [%]', ''),
                        str_data.get('Mean Dose [%]', '')
                    ]
                    writer.writerow(stats_row)
                writer.writerow([])  # Empty row

                # 3. Write the main DVH data (wide format)
                # Create the header row (e.g., Dose [%], BLAD_W_Volume [%], PTV_PSV_Volume [%], ...)
                dvh_header = [f'Dose {dose_unit}']
                for str_name in structure_names:
                    dvh_header.append(f'{str_name}_Volume [%]')
                writer.writerow(dvh_header)

                # Create a row for each dose point
                for dose in sorted_doses:
                    row = [dose]  # The first item is the dose value
                    # For each structure, get the volume data at that dose and add it to the row
                    for str_name in structure_names:
                        # .get(dose, '') retrieves the data or an empty string if it doesn't exist
                        volume = structures[str_name].get('dvh', {}).get(dose, '')
                        row.append(volume)
                    writer.writerow(row)

    # --- Method to create structure-wise CSVs ---
    def write_structure_csvs(self, all_patients_data, output_dir, dose_type, progress_var, progress_offset):
        dose_unit = "[%]" if dose_type == "%" else "[Gy]"
        
        # defaultdict is a convenient dictionary that automatically creates a default value (an empty dict here) if a key doesn't exist
        structure_data = defaultdict(lambda: defaultdict(dict))
        all_doses = set()
        patient_ids = sorted(list(all_patients_data.keys()))

        # Loop through all patient data to reorganize it by structure
        for pid, data in all_patients_data.items():
            for str_name, str_data in data.get('structures', {}).items():
                if 'dvh' in str_data:
                    for dose, volume in str_data['dvh'].items():
                        # Store volume data in the order: (structure name -> patient ID -> dose)
                        structure_data[str_name][pid][dose] = volume
                        all_doses.add(dose)

        sorted_doses = sorted(list(all_doses))

        # Use the reorganized data to create one CSV file per structure
        for str_name, patients in structure_data.items():
            # Replace characters that are invalid for filenames with "_"
            safe_str_name = re.sub(r'[\\/:*?"<>|]', '_', str_name)
            filename = output_dir / f"structure_{safe_str_name}.csv"
            
            with open(filename, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                
                # Header row (Dose, Patient ID 1, Patient ID 2, ...)
                header = [f'Dose {dose_unit}'] + patient_ids
                writer.writerow(header)
                
                # Create one row for each dose point
                for dose in sorted_doses:
                    row = [dose]
                    # For each patient, get the volume data at that dose and add it to the row
                    for pid in patient_ids:
                        volume = patients.get(pid, {}).get(dose, '')
                        row.append(volume)
                    writer.writerow(row)

# --- Program Execution Entry Point ---
# The code below is only executed when this file is run directly.
if __name__ == "__main__":
    main_window = tkinter.Tk()  # Create the main window
    app = DvhConverterApp(main_window)  # Create an instance of the DvhConverterApp class
    main_window.mainloop()  # Start the GUI application and wait for user interaction
