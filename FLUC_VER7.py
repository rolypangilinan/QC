import pandas as pd
import tkinter as tk
from tkinter import ttk
import os
import time
from datetime import datetime
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
import matplotlib.ticker as ticker

class DatabaseSelection:
    def __init__(self, root):
        self.root = root
        self.root.title("Settings")
        self.root.geometry("400x300")
        main_container = ttk.Frame(self.root)
        main_container.pack(fill=tk.BOTH, expand=True)
        self.canvas = tk.Canvas(main_container)
        self.scrollbar = ttk.Scrollbar(main_container, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas)
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(
                scrollregion=self.canvas.bbox("all")
            )
        )
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")
        ttk.Label(self.scrollable_frame, text="Choose Database Location:", font=('Arial', 12)).pack(pady=10)
        self.location_var = tk.StringVar()
        locations = [
            ("SWITCH 1 - ATU006", "HP1"),
            ("SWITCH 2 - ATU003", "HP2"),
            ("FAST-LINE - ATU007", "FASTLINE"),
            ("MULTILINE - ATU005", "MULTILINE"),
            ("TESTING", "TESTING")
        ]
        for text, location in locations:
            ttk.Radiobutton(self.scrollable_frame, text=text, variable=self.location_var,
                          value=location).pack(anchor='w', padx=40)
        ttk.Label(self.scrollable_frame, text="Select CSV Database:", font=('Arial', 12)).pack(pady=10)
        self.csv_var = tk.StringVar()
        base_dirs = [
            r"\\192.168.2.19\general\INSPECTION-MACHINE\HP1",
            r"\\192.168.2.19\general\INSPECTION-MACHINE\HP2",
            r"\\192.168.2.19\general\INSPECTION-MACHINE\FAST-LINE",
            r"\\192.168.2.19\general\INSPECTION-MACHINE\HP3",
            r"\\192.168.2.19\ai_team\INDIVIDUAL FOLDER\June-San\p2LTG\p2LTG_TransferData\OTHER PROJECT"
        ]
        csv_files = []
        for d in base_dirs:
            try:
                for f in os.listdir(d):
                    if f.startswith("log") and f.endswith(".csv"):
                        csv_files.append(os.path.join(d, f))
            except:
                pass
# DROPDOWN LIST COMBOBOX 
        self.csv_combo = ttk.Combobox(self.scrollable_frame, textvariable=self.csv_var, values=csv_files, width=90)
        self.csv_combo.pack(pady=5)
        ttk.Label(self.scrollable_frame, text="TOLERANCE:", font=('Arial', 12)).pack(pady=10)
        self.tolerance_var = tk.StringVar(value="5")
        tolerances = [
            ("3%", "3"),
            ("5%", "5"),
            ("OTHERS", "others")
        ]
        for text, val in tolerances:
            ttk.Radiobutton(self.scrollable_frame, text=text, variable=self.tolerance_var, value=val).pack(anchor='w', padx=40)
        self.other_frame = ttk.Frame(self.scrollable_frame)
        ttk.Label(self.other_frame, text="Enter %:").pack(side=tk.LEFT)
        self.other_entry = ttk.Entry(self.other_frame, width=5)
        self.other_entry.insert(0, "5")
        self.other_entry.pack(side=tk.LEFT)
        self.tolerance_var.trace("w", self.on_tolerance_change)
        ttk.Label(self.scrollable_frame, text="LOGICAL COMPUTATION:", font=('Arial', 12)).pack(pady=10)
        self.computation_var = tk.StringVar(value="FIND THE NEAREST GOOD")
        computations = [
            ("FIND THE NEAREST GOOD", "NEAREST"),
            ("ACCU AVG (TOL 5 Pts.)", "ACCU_AVG"),
            ("AKH (DOUBLE NOZZLE)", "AKH"),
            ("DUO (SINGLE NOZZLE)", "DUO")
        ]
        for text, val in computations:
            ttk.Radiobutton(self.scrollable_frame, text=text, variable=self.computation_var, value=val).pack(anchor='w', padx=40)
        ttk.Label(self.scrollable_frame, text="Generate CSV File:", font=('Arial', 12)).pack(pady=10)
        self.generate_csv_var = tk.StringVar()
        ttk.Radiobutton(self.scrollable_frame, text="YES", variable=self.generate_csv_var,
                       value="YES").pack(anchor='w', padx=40)
        ttk.Radiobutton(self.scrollable_frame, text="NO", variable=self.generate_csv_var,
                       value="NO").pack(anchor='w', padx=40)
        ttk.Button(self.scrollable_frame, text="Confirm", command=self.confirm_selection).pack(pady=10)
    def on_tolerance_change(self, *args):
        if self.tolerance_var.get() == "others":
            self.other_frame.pack(anchor='w', padx=40, pady=5)
            self.other_entry.focus()
            self.root.update_idletasks()
        else:
            self.other_frame.pack_forget()
            self.root.update_idletasks()
    def confirm_selection(self):
        location = self.location_var.get()
        generate_csv = self.generate_csv_var.get()
        file_path = self.csv_var.get()
        tolerance_str = self.tolerance_var.get()
        computation = self.computation_var.get()
        if tolerance_str == "others":
            tolerance_str = self.other_entry.get().strip()
        try:
            tolerance = float(tolerance_str)
        except ValueError:
            tolerance = 5.0
        if location and generate_csv and file_path:
            self.root.destroy()
            root = tk.Tk()
            app = FluctuationMonitor(root, location, generate_csv, tolerance, file_path, computation)
            root.protocol("WM_DELETE_WINDOW", app.on_closing)
            root.mainloop()

class FileChangeHandler(FileSystemEventHandler):
    def __init__(self, callback):
        self.callback = callback
    def on_modified(self, event):
        if event.src_path.endswith('.csv'):
            self.callback()

class FluctuationMonitor:
    def __init__(self, root, location, generate_csv, tolerance, file_path, computation):
        self.root = root
        self.root.title("Fluctuation Status Monitor - VER6")
        self.root.geometry("1000x800")
        self.zoom_level = 1.0
        self.generate_csv = generate_csv
        self.computation = computation
        self.main_container = ttk.Frame(root)
        self.main_container.pack(fill=tk.BOTH, expand=True)
        self.canvas = tk.Canvas(self.main_container)
        self.scrollbar = ttk.Scrollbar(self.main_container, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas)
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(
                scrollregion=self.canvas.bbox("all")
            )
        )
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")
        zoom_frame = ttk.Frame(self.scrollable_frame)
        zoom_frame.pack(pady=2, anchor='ne')
        ttk.Button(zoom_frame, text="Zoom In (+)", command=lambda: self.zoom(1.1)).pack(side=tk.LEFT, padx=2)
        ttk.Button(zoom_frame, text="Zoom Out (-)", command=lambda: self.zoom(0.9)).pack(side=tk.LEFT, padx=2)
        ttk.Button(zoom_frame, text="Reset Zoom", command=lambda: self.zoom(1.0)).pack(side=tk.LEFT, padx=2)
        self.status_vars = {}
        self.status_labels = {}
        self.current_model = None
        self.last_date = None
        self.last_good_values = {}
        self.last_good_serial = None
        self.accumulated_averages = {}
        self.run_counts = {}
        self.fluctuation_count = 0
        self.threshold = tolerance / 100
        self.file_path = file_path
        if location == "SWITCH 1 - ATU006":
            self.output_path = r"\\192.168.2.19\general\INSPECTION-MACHINE\HP2\fluctuatedQC_CSV.csv"
        elif location == "SWITCH 2 - ATU003":
            self.output_path = r"\\192.168.2.19\general\INSPECTION-MACHINE\HP1\fluctuatedQC_CSV.csv"
        elif location == "TESTING":
            self.output_path = r"\\192.168.2.19\ai_team\INDIVIDUAL FOLDER\June-San\p2LTG\p2LTG_TransferData\OTHER PROJECT\fluctuatedQC_CSV.csv"
        elif location == "MULTILINE":
            self.output_path = r"\\192.168.2.19\general\INSPECTION-MACHINE\HP3\FLUCTUATED PROGRAM\NEW VERSION\fluctuatedQC_CSV.csv"
        else:  # FASTLINE     
            self.output_path = r"\\192.168.2.19\general\INSPECTION-MACHINE\FAST-LINE\fluctuatedQC_CSV.csv"
        self.main_frame = ttk.Frame(self.scrollable_frame)
        self.main_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.top_frame = ttk.Frame(self.main_frame)
        self.top_frame.pack(fill=tk.X)
        self.left_frame = ttk.Frame(self.top_frame)
        self.left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        ttk.Label(self.left_frame, text="Fluctuation Status Monitor", 
                 font=('Arial', 14, 'bold')).pack(pady=5)
        ttk.Label(self.left_frame, text=f"Location: {location}", 
                 font=('Arial', 12)).pack(pady=2)
        ttk.Label(self.left_frame, text=f"Tolerance: {tolerance}%", 
                 font=('Arial', 12)).pack(pady=2)
        self.model_display = tk.Label(self.left_frame, text="MODEL CODE: N/A", 
                                    font=('Arial', 12), fg="purple")
        self.model_display.pack(pady=2)
        self.serial_display = tk.Label(self.left_frame, text="SERIAL No.", 
                                      font=('Arial', 16, 'bold'), fg="blue")
        self.serial_display.pack(pady=2)
        self.ref_serial_display = tk.Label(self.left_frame, text="PREVIOUS REF SN: N/A", 
                                         font=('Arial', 12), fg="green")
        self.ref_serial_display.pack(pady=2)
        self.status_box = tk.Canvas(self.left_frame, width=300, height=150,
                                   bg="gray", highlightthickness=2, 
                                   highlightbackground="black")
        self.status_box.pack(pady=10)
        self.status_text = self.status_box.create_text(150, 50, text="NO FLUCTUATION DETECTED", 
                                                     font=('Arial', 14, 'bold'), fill="white", anchor="center")
        self.counter_text = self.status_box.create_text(150, 100, text="Fluctuations: 0/8", 
                                                      font=('Arial', 10), fill="white", anchor="center")
        self.serial_log = tk.Text(self.left_frame, height=4, width=40)
        self.serial_log.pack(pady=5)
        self.serial_log.insert(tk.END, "Serial Number Log:\n")
        self.serial_log.config(state='disabled')
        self.details_frame = ttk.Frame(self.top_frame)
        self.details_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.create_section("50Hz Measurements", [
            "50Hz WATTAGE FLUCTUATED",
            "50Hz AIR VOLUME FLUCTUATED",
            "50Hz CLOSED PRESSURE FLUCTUATED",
            "50Hz AMPERAGE FLUCTUATED"
        ])
        ttk.Separator(self.details_frame, orient='horizontal').pack(fill='x', pady=5)
        self.create_section("60Hz Measurements", [
            "60Hz WATTAGE FLUCTUATED",
            "60Hz AIR VOLUME FLUCTUATED",
            "60Hz CLOSED PRESSURE FLUCTUATED",
            "60Hz AMPERAGE FLUCTUATED"
        ])
        ttk.Separator(self.details_frame, orient='horizontal').pack(fill='x', pady=5)
        self.right_frame = ttk.Frame(self.top_frame)
        self.right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        self.create_bar_graph()
        self.line_graph_frame = ttk.Frame(self.main_frame)
        self.line_graph_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        self.create_line_graph()
        button_frame = ttk.Frame(self.main_frame)
        button_frame.pack(pady=5)
        ttk.Button(button_frame, text="Refresh Values", 
                  command=self.process_and_update).pack(side=tk.LEFT, padx=2)
        ttk.Button(button_frame, text="Reset All", 
                  command=self.reset_all_fluctuations).pack(side=tk.LEFT, padx=2)
        ttk.Button(button_frame, text="Focus", 
                  command=self.open_focus_selection).pack(side=tk.LEFT, padx=2)
        self.last_modified = os.path.getmtime(self.file_path)
        event_handler = FileChangeHandler(self.check_file_update)
        self.observer = Observer()
        self.observer.schedule(event_handler, os.path.dirname(self.file_path))
        self.observer.start()
        self.process_and_update()
        self.root.after(1000, self.periodic_check)

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")

    def zoom(self, factor):
        if factor != 1.0:
            self.zoom_level *= factor
        else:
            self.zoom_level = 1.0
        self.zoom_level = max(0.5, min(self.zoom_level, 2.0))
        widgets = [
            (self.serial_display, 16),
            (self.ref_serial_display, 12),
            (self.model_display, 12),
        ]
        for widget, base_size in widgets:
            current_font = widget.cget("font")
            font_name = 'Arial'
            if isinstance(current_font, str):
                font_name = current_font.split()[0]
            new_size = int(base_size * self.zoom_level)
            widget.config(font=(font_name, new_size, 'bold' if 'bold' in str(current_font) else ''))
        new_width = int(300 * self.zoom_level)
        new_height = int(150 * self.zoom_level)
        self.status_box.config(width=new_width, height=new_height)
        status_y = int(50 * self.zoom_level)
        counter_y = int(100 * self.zoom_level)
        min_spacing = int(30 * self.zoom_level)
        if counter_y - status_y < min_spacing:
            counter_y = status_y + min_spacing
        current_font = self.status_box.itemcget(self.status_text, "font")
        font_parts = current_font.split()
        font_name = 'Arial'
        if len(font_parts) > 0:
            font_name = font_parts[0]
        new_size = int(14 * self.zoom_level)
        font_weight = 'bold' if 'bold' in current_font else ''
        self.status_box.itemconfig(self.status_text, font=(font_name, new_size, font_weight))
        self.status_box.coords(self.status_text, new_width / 2, status_y)
        current_font = self.status_box.itemcget(self.counter_text, "font")
        font_parts = current_font.split()
        font_name = 'Arial'
        if len(font_parts) > 0:
            font_name = font_parts[0]
        new_size = int(10 * self.zoom_level)
        font_weight = '' if 'bold' in current_font else ''
        self.status_box.itemconfig(self.counter_text, font=(font_name, new_size, font_weight))
        self.status_box.coords(self.counter_text, new_width / 2, counter_y)
        for column in self.status_vars:
            new_size = int(10 * self.zoom_level)
            self.status_labels[column].config(font=('Arial', new_size, 'bold'))
        for frame in self.details_frame.winfo_children():
            for widget in frame.winfo_children():
                if isinstance(widget, ttk.Label) and widget.cget('text') in ["50Hz Measurements", "60Hz Measurements"]:
                    new_size = int(12 * self.zoom_level)
                    widget.config(font=('Arial', new_size, 'bold'))
        if hasattr(self, 'fig'):
            self.fig.set_size_inches(6 * self.zoom_level, 4 * self.zoom_level)
            self.canvas_graph.draw()
        if hasattr(self, 'line_fig'):
            self.line_fig.set_size_inches(8 * self.zoom_level, 3 * self.zoom_level)
            self.line_canvas.draw()

    def create_bar_graph(self):
        """Create the matplotlib bar graph for displaying current fluctuation values"""
        self.fig, self.ax = plt.subplots(figsize=(6, 3.9))
        self.ax.set_title('Current Fluctuation Values')
        self.ax.set_ylabel('Fluctuation Amount (%)')
        self.ax.grid(True)
        self.fluctuation_measurements = [
            '50 WAT', '50 VOL', '50 CloP.', '50 AMP',
            '60 WAT', '60 VOL', '60 CloP.', '60 AMP'
        ]
        self.fluctuation_values = [0] * len(self.fluctuation_measurements)
        self.bars = self.ax.bar(self.fluctuation_measurements, self.fluctuation_values, color='skyblue')
        # Set fixed ticks to avoid UserWarning
        self.ax.set_xticks(range(len(self.fluctuation_measurements)))
        self.ax.set_xticklabels(self.fluctuation_measurements, rotation=45, ha='right', fontsize=7)
        self.ax.axhline(y=self.threshold * 100, color='r', linestyle='--', linewidth=1)
        self.ax.text(1.01, self.threshold * 100, 'Threshold (%.1f%%)' % (self.threshold * 100), color='r', va='center', ha='left', 
                     transform=self.ax.get_yaxis_transform(), fontsize=8, 
                     bbox=dict(facecolor='white', edgecolor='none', alpha=0.8))
        # Adjust layout with larger margins to avoid tight_layout warning
        self.fig.subplots_adjust(left=0.1, right=0.85, top=0.9, bottom=0.25)
        self.canvas_graph = FigureCanvasTkAgg(self.fig, master=self.right_frame)
        self.canvas_graph.draw()
        self.canvas_graph.get_tk_widget().pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    def create_line_graph(self):
        """Create the matplotlib line graph for displaying historical trends"""
        self.line_fig = Figure(figsize=(8, 2.5), dpi=100)
        self.line_ax = self.line_fig.add_subplot(111)
        self.line_ax.set_title('Historical Measurement SPC')
        self.line_ax.set_ylabel('Value')
        self.line_ax.grid(True)
        self.line_ax.plot([], [], label='50Hz WATTAGE', color='blue', linewidth=1.5)
        self.line_ax.plot([], [], label='50Hz AIR VOLUME', color='green', linewidth=1.5)
        self.line_ax.plot([], [], label='50Hz CLOSED PRESSURE', color='red', linewidth=1.5)
        self.line_ax.plot([], [], label='50Hz AMPERAGE', color='cyan', linewidth=1.5)
        self.line_ax.plot([], [], label='60Hz WATTAGE', color='magenta', linewidth=1.5)
        self.line_ax.plot([], [], label='60Hz AIR VOLUME', color='yellow', linewidth=1.5)
        self.line_ax.plot([], [], label='60Hz CLOSED PRESSURE', color='black', linewidth=1.5)
        self.line_ax.plot([], [], label='60Hz AMPERAGE', color='orange', linewidth=1.5)
        legend = self.line_ax.legend(loc='center left', bbox_to_anchor=(1, 0.5))
        for line in legend.get_lines():
            line.set_linewidth(4.0)  # Thicker legend lines
        self.line_ax.tick_params(axis='x', rotation=45, labelsize=8)
        self.line_fig.subplots_adjust(left=0.1, right=0.75, bottom=0.25, top=0.9)
        self.line_canvas = FigureCanvasTkAgg(self.line_fig, master=self.line_graph_frame)
        self.line_canvas.draw()
        self.line_canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

    def downsample_data(self, df, max_points=50):
        if len(df) <= max_points:
            return df
        step = len(df) // max_points
        return df.iloc[::step]

    def update_line_graph(self):
        """Update the line graph with historical data"""
        try:
            if not hasattr(self, 'compiledFrame') or self.compiledFrame.empty:
                return
            self.line_ax.clear()
            df = self.compiledFrame.tail(50)
            df = self.downsample_data(df)
            self.line_ax.plot(df['DATETIME'], df['50Hz WATTAGE'], label='50Hz WATTAGE', color='blue', linewidth=1.5)
            self.line_ax.plot(df['DATETIME'], df['50Hz AIR VOLUME'], label='50Hz AIR VOLUME', color='green', linewidth=1.5)
            self.line_ax.plot(df['DATETIME'], df['50Hz CLOSED PRESSURE'], label='50Hz CLOSED PRESSURE', color='red', linewidth=1.5)
            self.line_ax.plot(df['DATETIME'], df['50Hz AMPERAGE'], label='50Hz AMPERAGE', color='cyan', linewidth=1.5)
            self.line_ax.plot(df['DATETIME'], df['60Hz WATTAGE'], label='60Hz WATTAGE', color='magenta', linewidth=1.5)
            self.line_ax.plot(df['DATETIME'], df['60Hz AIR VOLUME'], label='60Hz AIR VOLUME', color='yellow', linewidth=1.5)
            self.line_ax.plot(df['DATETIME'], df['60Hz CLOSED PRESSURE'], label='60Hz CLOSED PRESSURE', color='black', linewidth=1.5)
            self.line_ax.plot(df['DATETIME'], df['60Hz AMPERAGE'], label='60Hz AMPERAGE', color='orange', linewidth=1.5)
            self.line_ax.set_title('Historical Measurement SPC')
            self.line_ax.set_ylabel('Value')
            self.line_ax.grid(True)
            self.line_ax.tick_params(axis='x', rotation=45, labelsize=8)
            legend = self.line_ax.legend(loc='center left', bbox_to_anchor=(1, 0.5))
            for line in legend.get_lines():
                line.set_linewidth(4.0)  # Thicker legend lines
            self.line_fig.subplots_adjust(left=0.1, right=0.75, bottom=0.25, top=0.9)
            self.line_canvas.draw()
        except Exception as e:
            print(f"Error updating line graph: {e}")

    def update_bar_graph(self, last_row):
        """Update the bar graph with new fluctuation values"""
        try:
            self.ax.clear()
            self.fluctuation_values = [
                last_row['50Hz WATTAGE FLUCTUATED'] * 100,
                last_row['50Hz AIR VOLUME FLUCTUATED'] * 100,
                last_row['50Hz CLOSED PRESSURE FLUCTUATED'] * 100,
                last_row['50Hz AMPERAGE FLUCTUATED'] * 100,
                last_row['60Hz WATTAGE FLUCTUATED'] * 100,
                last_row['60Hz AIR VOLUME FLUCTUATED'] * 100,
                last_row['60Hz CLOSED PRESSURE FLUCTUATED'] * 100,
                last_row['60Hz AMPERAGE FLUCTUATED'] * 100
            ]
            self.bars = self.ax.bar(self.fluctuation_measurements, self.fluctuation_values, color='skyblue')
            self.ax.set_title('Current Fluctuation Values')
            self.ax.set_ylabel('Fluctuation Amount (%)')
            self.ax.grid(True)
            self.ax.set_xticks(range(len(self.fluctuation_measurements)))
            self.ax.set_xticklabels(self.fluctuation_measurements, rotation=45, ha='right', fontsize=7)
            max_value = max(self.fluctuation_values) if max(self.fluctuation_values) > 0 else 10
            self.ax.set_ylim(0, max_value * 1.1)
            for bar in self.bars:
                height = bar.get_height()
                self.ax.text(bar.get_x() + bar.get_width()/2., height,
                            f'{height:.2f}',
                            ha='center', va='bottom')
            self.ax.axhline(y=self.threshold * 100, color='r', linestyle='--', linewidth=1)
            self.ax.text(1.01, self.threshold * 100, 'Threshold (%.1f%%)' % (self.threshold * 100), color='r', va='center', ha='left', 
                         transform=self.ax.get_yaxis_transform(), fontsize=8, 
                         bbox=dict(facecolor='white', edgecolor='none', alpha=0.8))
            self.fig.subplots_adjust(left=0.1, right=0.85, top=0.9, bottom=0.25)
            self.canvas_graph.draw()
        except Exception as e:
            print(f"Error updating bar graph: {e}")

    def update_status_box(self, has_fluctuation, count=0):
        if has_fluctuation:
            self.status_box.configure(bg="red")
            self.status_box.itemconfig(self.status_text, text="FLUCTUATION DETECTED!")
            self.status_box.itemconfig(self.counter_text, text=f"Fluctuations: {count}/8")
        else:
            self.status_box.configure(bg="green")
            self.status_box.itemconfig(self.status_text, text="NO FLUCTUATION DETECTED")
            self.status_box.itemconfig(self.counter_text, text="Fluctuations: 0/8")

    def periodic_check(self):
        self.check_file_update()
        self.root.after(1000, self.periodic_check)

    def check_file_update(self):
        try:
            current_modified = os.path.getmtime(self.file_path)
            if current_modified > self.last_modified:
                self.last_modified = current_modified
                self.process_and_update()
        except:
            pass

    def process_and_update(self):
        global compiledFrame
        try:
            dataList = []
            pd.set_option('display.max_columns', None)
            try:
                df = pd.read_csv(self.file_path, encoding='latin1')
            except UnicodeDecodeError:
                df = pd.read_csv(self.file_path, encoding='ISO-8859-1', errors='replace')
            df.columns = df.columns.str.strip()
            def fix_time(time_str):
                try:
                    hours, minutes, seconds = map(int, time_str.split(':'))
                    if hours >= 24:
                        hours = hours % 24
                        return f"{hours:02d}:{minutes:02d}:{seconds:02d}"
                    return time_str
                except:
                    return time_str
            df['TIME'] = df['TIME'].apply(fix_time)
            df['DATETIME'] = pd.to_datetime(
                df['DATE'] + ' ' + df['TIME'],
                dayfirst=True,
                format='mixed',
                errors='coerce'
            )
            df = df.dropna(subset=['DATETIME'])
            df = df.sort_values('DATETIME')
            emptyColumn = [
                "DATE", "TIME", "MODEL CODE", "TYPE", "BARCODE", "SERIAL No.", "PASS/NG",
                "50Hz WATTAGE", "50Hz WATTAGE FLUCTUATED",
                "50Hz AIR VOLUME", "50Hz AIR VOLUME FLUCTUATED",
                "50Hz CLOSED PRESSURE", "50Hz CLOSED PRESSURE FLUCTUATED",
                "50Hz AMPERAGE", "50Hz AMPERAGE FLUCTUATED",
                "60Hz WATTAGE", "60Hz WATTAGE FLUCTUATED",
                "60Hz AIR VOLUME", "60Hz AIR VOLUME FLUCTUATED",
                "60Hz CLOSED PRESSURE", "60Hz CLOSED PRESSURE FLUCTUATED",
                "60Hz AMPERAGE", "60Hz AMPERAGE FLUCTUATED",
                "REFERENCE SERIAL", "DATETIME"
            ]
            # Define dtypes for compiledFrame to ensure consistency
            dtype_dict = {
                "DATE": str,
                "TIME": str,
                "MODEL CODE": str,
                "TYPE": str,
                "BARCODE": str,
                "SERIAL No.": str,
                "PASS/NG": str,
                "50Hz WATTAGE": float,
                "50Hz WATTAGE FLUCTUATED": float,
                "50Hz AIR VOLUME": float,
                "50Hz AIR VOLUME FLUCTUATED": float,
                "50Hz CLOSED PRESSURE": float,
                "50Hz CLOSED PRESSURE FLUCTUATED": float,
                "50Hz AMPERAGE": float,
                "50Hz AMPERAGE FLUCTUATED": float,
                "60Hz WATTAGE": float,
                "60Hz WATTAGE FLUCTUATED": float,
                "60Hz AIR VOLUME": float,
                "60Hz AIR VOLUME FLUCTUATED": float,
                "60Hz CLOSED PRESSURE": float,
                "60Hz CLOSED PRESSURE FLUCTUATED": float,
                "60Hz AMPERAGE": float,
                "60Hz AMPERAGE FLUCTUATED": float,
                "REFERENCE SERIAL": str,
                "DATETIME": "datetime64[ns]"
            }
            if not hasattr(self, 'compiledFrame') or self.compiledFrame.empty:
                self.compiledFrame = pd.DataFrame(columns=emptyColumn).astype(dtype_dict)
            df = df[(~df["MODEL CODE"].isin(['120HP1000M', '60CAT0203M']))]
            df = df[(~df["TYPE"].isin(['T', 'D', 'A']))]
            df = df[(~df["PASS/NG"].isin([0]))]
            previous_values = {
                '50Hz_WATTAGE': None,
                '50Hz_AIR_VOLUME': None,
                '50Hz_CLOSED_PRESSURE': None,
                '50Hz_AMPERAGE': None,
                '60Hz_WATTAGE': None,
                '60Hz_AIR_VOLUME': None,
                '60Hz_CLOSED_PRESSURE': None,
                '60Hz_AMPERAGE': None
            }
            previous_date = None
            previous_model = None
            for a in range(len(df)):
                tempdf = df.iloc[[a]]
                current_date = tempdf["DATE"].values[0]
                model_code = tempdf["MODEL CODE"].values[0]
                serial_no = tempdf["SERIAL No."].values[0]
                self.serial_display.config(text=f"SERIAL No.: {serial_no}")
                self.model_display.config(text=f"MODEL CODE: {model_code}")
                self.serial_log.config(state='normal')
                self.serial_log.insert(tk.END, f"{serial_no}\n")
                self.serial_log.see(tk.END)
                self.serial_log.config(state='disabled')
                model_changed = (self.current_model is not None and model_code != self.current_model)
                is_new_date = (current_date != self.last_date) if self.last_date is not None else True
                is_new_model_in_day = (model_code != previous_model) if previous_model is not None else False
                if model_changed:
                    self.current_model = model_code
                    self.last_good_values = {}
                    self.last_good_serial = None
                    self.accumulated_averages = {}
                    self.run_counts = {}
                if is_new_date:
                    self.last_date = current_date
                    self.accumulated_averages = {}
                    self.run_counts = {}
                    previous_model = None
                current_values = {
                    '50Hz_WATTAGE': tempdf["50Hz WATTAGE"].values[0],
                    '50Hz_AIR_VOLUME': tempdf["50Hz AIR VOLUME"].values[0],
                    '50Hz_CLOSED_PRESSURE': tempdf["50Hz CLOSED PRESSURE"].values[0],
                    '50Hz_AMPERAGE': tempdf["50Hz AMPERAGE"].values[0],
                    '60Hz_WATTAGE': tempdf["60Hz WATTAGE"].values[0],
                    '60Hz_AIR_VOLUME': tempdf["60Hz AIR VOLUME"].values[0],
                    '60Hz_CLOSED_PRESSURE': tempdf["60Hz CLOSED PRESSURE"].values[0],
                    '60Hz_AMPERAGE': tempdf["60Hz AMPERAGE"].values[0]
                }
                fluctuations = {}
                ref_serial = "N/A"
                if self.computation == "ACCU_AVG":
                    # Initialize accumulated averages and run counts for the model and date
                    model_date_key = f"{model_code}_{current_date}"
                    if model_date_key not in self.accumulated_averages:
                        self.accumulated_averages[model_date_key] = {
                            '50Hz_WATTAGE': [],
                            '50Hz_AIR_VOLUME': [],
                            '50Hz_CLOSED_PRESSURE': [],
                            '50Hz_AMPERAGE': [],
                            '60Hz_WATTAGE': [],
                            '60Hz_AIR_VOLUME': [],
                            '60Hz_CLOSED_PRESSURE': [],
                            '60Hz_AMPERAGE': []
                        }
                        self.run_counts[model_date_key] = 0
                    self.run_counts[model_date_key] += 1
                    # First run or new day: set fluctuation to 5%
                    if self.run_counts[model_date_key] == 1 or is_new_date:
                        for key in current_values:
                            fluctuations[key] = 0.05 if current_values[key] != 0 else 0
                        # Store current values for averaging
                        for key in current_values:
                            if current_values[key] != 0:
                                self.accumulated_averages[model_date_key][key].append(current_values[key])
                    else:
                        # Calculate average of all previous runs for this model and date
                        for key in current_values:
                            if current_values[key] == 0:
                                fluctuations[key] = 0
                            else:
                                avg = sum(self.accumulated_averages[model_date_key][key]) / len(self.accumulated_averages[model_date_key][key])
                                fluctuations[key] = abs((avg - current_values[key]) / avg) if avg != 0 else 0
                                self.accumulated_averages[model_date_key][key].append(current_values[key])
                    ref_serial = "N/A"  # No reference serial for ACCU_AVG
                elif self.computation in ["AKH", "DUO"]:
                    # Placeholder for AKH and DUO logic
                    for key in current_values:
                        fluctuations[key] = 0.05 if current_values[key] != 0 else 0
                    ref_serial = serial_no
                else:  # NEAREST
                    if model_changed or is_new_date or not self.last_good_values or is_new_model_in_day:
                        for key in current_values:
                            fluctuations[key] = 0.05 if current_values[key] != 0 else 0
                        if all(v == 0 for v in fluctuations.values()):
                            self.last_good_values = current_values.copy()
                            self.last_good_serial = serial_no
                            ref_serial = serial_no
                    else:
                        for key in current_values:
                            if current_values[key] == 0:
                                fluctuations[key] = 0
                            else:
                                fluctuations[key] = abs((self.last_good_values[key] - current_values[key]) / self.last_good_values[key]) if self.last_good_values[key] != 0 else 0
                        ref_serial = self.last_good_serial
                    if all(v <= self.threshold for v in fluctuations.values()):
                        self.last_good_values = current_values.copy()
                        self.last_good_serial = serial_no
                if self.computation != "ACCU_AVG":
                    self.ref_serial_display.config(text=f"PREVIOUS REF SN: {ref_serial}")
                else:
                    self.ref_serial_display.config(text="PREVIOUS REF SN: N/A")
                previous_model = model_code
                dataFrame = {
                    "DATE": current_date,
                    "TIME": tempdf["TIME"].values[0],
                    "MODEL CODE": model_code,
                    "TYPE": tempdf["TYPE"].values[0],
                    "BARCODE": tempdf["BARCODE"].values[0],
                    "SERIAL No.": serial_no,
                    "PASS/NG": tempdf["PASS/NG"].values[0],
                    "50Hz WATTAGE": current_values['50Hz_WATTAGE'],
                    "50Hz WATTAGE FLUCTUATED": fluctuations['50Hz_WATTAGE'],
                    "50Hz AIR VOLUME": current_values['50Hz_AIR_VOLUME'],
                    "50Hz AIR VOLUME FLUCTUATED": fluctuations['50Hz_AIR_VOLUME'],
                    "50Hz CLOSED PRESSURE": current_values['50Hz_CLOSED_PRESSURE'],
                    "50Hz CLOSED PRESSURE FLUCTUATED": fluctuations['50Hz_CLOSED_PRESSURE'],
                    "50Hz AMPERAGE": current_values['50Hz_AMPERAGE'],
                    "50Hz AMPERAGE FLUCTUATED": fluctuations['50Hz_AMPERAGE'],
                    "60Hz WATTAGE": current_values['60Hz_WATTAGE'],
                    "60Hz WATTAGE FLUCTUATED": fluctuations['60Hz_WATTAGE'],
                    "60Hz AIR VOLUME": current_values['60Hz_AIR_VOLUME'],
                    "60Hz AIR VOLUME FLUCTUATED": fluctuations['60Hz_AIR_VOLUME'],
                    "60Hz CLOSED PRESSURE": current_values['60Hz_CLOSED_PRESSURE'],
                    "60Hz CLOSED PRESSURE FLUCTUATED": fluctuations['60Hz_CLOSED_PRESSURE'],
                    "60Hz AMPERAGE": current_values['60Hz_AMPERAGE'],
                    "60Hz AMPERAGE FLUCTUATED": fluctuations['60Hz_AMPERAGE'],
                    "REFERENCE SERIAL": ref_serial,
                    "DATETIME": tempdf["DATETIME"].values[0]
                }
                previous_values = current_values.copy()
                previous_date = current_date
                dataList.append(dataFrame)
            if dataList:
                new_data = pd.DataFrame(dataList)
                # Ensure new_data has all columns from emptyColumn with correct dtypes
                for col in emptyColumn:
                    if col not in new_data.columns:
                        new_data[col] = pd.Series([None] * len(new_data), dtype=dtype_dict[col])
                    else:
                        new_data[col] = new_data[col].astype(dtype_dict[col], errors='ignore')
                # Filter out any all-NA columns from new_data before concatenation
                new_data = new_data.loc[:, ~new_data.isna().all()]
                # Concatenate with compiledFrame, ensuring dtypes are preserved
                self.compiledFrame = pd.concat([self.compiledFrame, new_data], ignore_index=True)
            if self.generate_csv == "YES":
                self.compiledFrame.to_csv(self.output_path, index=False, encoding='utf-8-sig')
            self.update_display()
        except Exception as e:
            print(f"Error processing file: {e}")

    def update_display(self):
        try:
            last_row = self.compiledFrame.iloc[-1]
            has_fluctuation = False
            fluctuation_count = 0
            for column in self.status_vars:
                if column in last_row:
                    value = last_row[column]
                    status = 1 if value > self.threshold else 0
                    self.status_vars[column].set(f"= {status}")
                    color = "red" if status == 1 else "green"
                    self.status_labels[column].configure(foreground=color)
                    if status == 1:
                        has_fluctuation = True
                        fluctuation_count += 1
            self.update_status_box(has_fluctuation, fluctuation_count)
            if "REFERENCE SERIAL" in last_row and self.computation != "ACCU_AVG":
                ref_serial = last_row["REFERENCE SERIAL"]
                self.ref_serial_display.config(text=f"PREVIOUS REF SN: {ref_serial}")
            else:
                self.ref_serial_display.config(text="PREVIOUS REF SN: N/A")
            if "MODEL CODE" in last_row:
                model_code = last_row["MODEL CODE"]
                self.model_display.config(text=f"MODEL CODE: {model_code}")
            self.update_bar_graph(last_row)
            self.update_line_graph()
        except Exception as e:
            print(f"Error updating display: {e}")

    def create_section(self, title, columns):
        frame = ttk.Frame(self.details_frame)
        frame.pack(fill=tk.X, pady=2)
        ttk.Label(frame, text=title, font=('Arial', 12, 'bold')).pack(anchor='w')
        for col in columns:
            self.create_status_row(frame, col)

    def create_status_row(self, parent, column_name):
        row_frame = ttk.Frame(parent)
        row_frame.pack(fill=tk.X, pady=1)
        ttk.Label(row_frame, text=f"{column_name.replace(' FLUCTUATED', '')}:", 
                 width=20, anchor='w').pack(side=tk.LEFT)
        status_frame = ttk.Frame(row_frame)
        status_frame.pack(side=tk.LEFT, padx=2)
        self.status_vars[column_name] = tk.StringVar()
        self.status_labels[column_name] = ttk.Label(
            status_frame, 
            textvariable=self.status_vars[column_name],
            font=('Arial', 10, 'bold'),
            width=8
        )
        self.status_labels[column_name].pack(side=tk.LEFT)
        ttk.Button(
            status_frame, 
            text="Reset", 
            width=5,
            command=lambda: self.reset_fluctuation(column_name)
        ).pack(side=tk.LEFT, padx=2)

    def reset_fluctuation(self, column_name):
        try:
            self.compiledFrame.at[self.compiledFrame.index[-1], column_name] = 0
            self.status_vars[column_name].set("= 0")
            self.status_labels[column_name].configure(foreground="green")
            last_row = self.compiledFrame.iloc[-1]
            if all(last_row[col] <= self.threshold for col in self.status_vars if col in last_row):
                self.last_good_values = {
                    '50Hz_WATTAGE': last_row["50Hz WATTAGE"],
                    '50Hz_AIR_VOLUME': last_row["50Hz AIR VOLUME"],
                    '50Hz_CLOSED_PRESSURE': last_row["50Hz CLOSED PRESSURE"],
                    '50Hz_AMPERAGE': last_row["50Hz AMPERAGE"],
                    '60Hz_WATTAGE': last_row["60Hz WATTAGE"],
                    '60Hz_AIR_VOLUME': last_row["60Hz AIR VOLUME"],
                    '60Hz_CLOSED_PRESSURE': last_row["60Hz CLOSED PRESSURE"],
                    '60Hz_AMPERAGE': last_row["60Hz AMPERAGE"]
                }
                self.last_good_serial = last_row["SERIAL No."]
                if self.computation != "ACCU_AVG":
                    self.ref_serial_display.config(text=f"PREVIOUS REF SN: {self.last_good_serial}")
                    self.compiledFrame.at[self.compiledFrame.index[-1], "REFERENCE SERIAL"] = self.last_good_serial
            if self.generate_csv == "YES":
                self.compiledFrame.to_csv(self.output_path, index=False, encoding='utf-8-sig')
            self.update_display()
        except Exception as e:
            print(f"Error resetting fluctuation: {e}")

    def reset_all_fluctuations(self):
        try:
            for column in self.status_vars:
                self.compiledFrame.at[self.compiledFrame.index[-1], column] = 0
                self.status_vars[column].set("= 0")
                self.status_labels[column].configure(foreground="green")
            last_row = self.compiledFrame.iloc[-1]
            self.last_good_values = {
                '50Hz_WATTAGE': last_row["50Hz WATTAGE"],
                '50Hz_AIR_VOLUME': last_row["50Hz AIR VOLUME"],
                '50Hz_CLOSED_PRESSURE': last_row["50Hz CLOSED PRESSURE"],
                '50Hz_AMPERAGE': last_row["50Hz AMPERAGE"],
                '60Hz_WATTAGE': last_row["60Hz WATTAGE"],
                '60Hz_AIR_VOLUME': last_row["60Hz AIR VOLUME"],
                '60Hz_CLOSED_PRESSURE': last_row["60Hz CLOSED PRESSURE"],
                '60Hz_AMPERAGE': last_row["60Hz AMPERAGE"]
            }
            self.last_good_serial = last_row["SERIAL No."]
            if self.computation != "ACCU_AVG":
                self.ref_serial_display.config(text=f"PREVIOUS REF SN: {self.last_good_serial}")
                self.compiledFrame.at[self.compiledFrame.index[-1], "REFERENCE SERIAL"] = self.last_good_serial
            else:
                self.ref_serial_display.config(text="PREVIOUS REF SN: N/A")
                self.compiledFrame.at[self.compiledFrame.index[-1], "REFERENCE SERIAL"] = "N/A"
            if self.generate_csv == "YES":
                self.compiledFrame.to_csv(self.output_path, index=False, encoding='utf-8-sig')
            self.update_status_box(False)
            self.update_display()
        except Exception as e:
            print(f"Error resetting all fluctuations: {e}")

    def open_focus_selection(self):
        focus_window = tk.Toplevel(self.root)
        focus_window.title("Select Measurement to Focus")
        focus_window.geometry("300x400")
        ttk.Label(focus_window, text="Select measurement:", font=('Arial', 12)).pack(pady=10)
        self.focus_var = tk.StringVar()
        measurements = [
            "50Hz WATTAGE",
            "50Hz AIR VOLUME",
            "50Hz CLOSED PRESSURE",
            "50Hz AMPERAGE",
            "60Hz WATTAGE",
            "60Hz AIR VOLUME",
            "60Hz CLOSED PRESSURE",
            "60Hz AMPERAGE"
        ]
        for meas in measurements:
            ttk.Radiobutton(focus_window, text=meas, variable=self.focus_var, value=meas).pack(anchor='w', padx=20)
        ttk.Button(focus_window, text="Show Graph", command=self.show_focused_graph).pack(pady=10)

    def show_focused_graph(self):
        selected = self.focus_var.get()
        if not selected:
            return
        graph_window = tk.Toplevel(self.root)
        graph_window.title(f"Focused: {selected}")
        graph_window.geometry("1200x800")
        fig = Figure(figsize=(12, 8), dpi=100)
        ax = fig.add_subplot(111)
        ax.set_title(f'{selected} SPC')
        ax.set_ylabel('Value')
        ax.grid(True)
        if hasattr(self, 'compiledFrame') and not self.compiledFrame.empty:
            df = self.compiledFrame
            df = self.downsample_data(df, max_points=100)
            color_map = {
                "50Hz WATTAGE": 'blue',
                "50Hz AIR VOLUME": 'green',
                "50Hz CLOSED PRESSURE": 'red',
                "50Hz AMPERAGE": 'cyan',
                "60Hz WATTAGE": 'magenta',
                "60Hz AIR VOLUME": 'yellow',
                "60Hz CLOSED PRESSURE": 'black',
                "60Hz AMPERAGE": 'orange'
            }
            ax.plot(df['DATETIME'], df[selected], label=selected, color=color_map.get(selected, 'blue'), linewidth=2)
            ax.legend()
            ax.tick_params(axis='x', rotation=45, labelsize=10)
        fig.subplots_adjust(left=0.1, right=0.9, bottom=0.15, top=0.9)
        canvas = FigureCanvasTkAgg(fig, master=graph_window)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

    def on_closing(self):
        self.observer.stop()
        self.observer.join()
        self.root.destroy()

# Initialize compiledFrame with explicit dtypes
dtype_dict = {
    "DATE": str,
    "TIME": str,
    "MODEL CODE": str,
    "TYPE": str,
    "BARCODE": str,
    "SERIAL No.": str,
    "PASS/NG": str,
    "50Hz WATTAGE": float,
    "50Hz WATTAGE FLUCTUATED": float,
    "50Hz AIR VOLUME": float,
    "50Hz AIR VOLUME FLUCTUATED": float,
    "50Hz CLOSED PRESSURE": float,
    "50Hz CLOSED PRESSURE FLUCTUATED": float,
    "50Hz AMPERAGE": float,
    "50Hz AMPERAGE FLUCTUATED": float,
    "60Hz WATTAGE": float,
    "60Hz WATTAGE FLUCTUATED": float,
    "60Hz AIR VOLUME": float,
    "60Hz AIR VOLUME FLUCTUATED": float,
    "60Hz CLOSED PRESSURE": float,
    "60Hz CLOSED PRESSURE FLUCTUATED": float,
    "60Hz AMPERAGE": float,
    "60Hz AMPERAGE FLUCTUATED": float,
    "REFERENCE SERIAL": str,
    "DATETIME": "datetime64[ns]"
}
compiledFrame = pd.DataFrame(columns=[
    "DATE", "TIME", "MODEL CODE", "TYPE", "BARCODE", "SERIAL No.", "PASS/NG",
    "50Hz WATTAGE", "50Hz WATTAGE FLUCTUATED",
    "50Hz AIR VOLUME", "50Hz AIR VOLUME FLUCTUATED",
    "50Hz CLOSED PRESSURE", "50Hz CLOSED PRESSURE FLUCTUATED",
    "50Hz AMPERAGE", "50Hz AMPERAGE FLUCTUATED",
    "60Hz WATTAGE", "60Hz WATTAGE FLUCTUATED",
    "60Hz AIR VOLUME", "60Hz AIR VOLUME FLUCTUATED",
    "60Hz CLOSED PRESSURE", "60Hz CLOSED PRESSURE FLUCTUATED",
    "60Hz AMPERAGE", "60Hz AMPERAGE FLUCTUATED",
    "REFERENCE SERIAL", "DATETIME"
]).astype(dtype_dict)
root = tk.Tk()
DatabaseSelection(root)
root.mainloop()