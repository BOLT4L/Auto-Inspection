import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os
from datetime import datetime, date, timedelta
import json
import random
import sys
import platform
print(platform.architecture())
sys.path.append(os.path.join(os.path.dirname(__file__)))
from database_manager import DatabaseManager
import hashlib
import win32print
import win32ui
from PIL import Image, ImageWin

class CarInspectionApp:
    def __init__(self, root):
        self.root = root
        self.root.title("MM Vehicle Inspection Management System")
        self.root.geometry("1200x800")
        self.root.configure(bg='#f0f0f0')
        self.image_dir = None  
        
        # Style configuration
        self.setup_styles()
        
        # Create main interface
        self.create_main_interface()
        
        # Initialize database
        self.setup_database()
    
    def setup_styles(self):
        """Apply a visually appealing, modern dark navy/white/blue theme with rounded cards and shadows."""
        style = ttk.Style()
        style.theme_use('clam')
        # Colors
        dark_navy = '#10172a'
        card_bg = '#ffffff'
        accent_blue = '#2563eb'
        accent_blue2 = '#3b82f6'
        text_dark = '#1e293b'
        text_gray = '#64748b'
        border = '#e5e7eb'
        shadow = '#e0e7ef'
        # Main window
        self.root.configure(bg=dark_navy)
        # Card frames
        style.configure('Card.TFrame', background=card_bg, relief='flat', borderwidth=0)
        style.configure('Card.TLabelframe', background=card_bg, relief='flat', borderwidth=0)
        style.configure('Card.TLabelframe.Label', background=card_bg, foreground=accent_blue, font=('Segoe UI', 12, 'bold'))
        # Headers
        style.configure('Title.TLabel', background=dark_navy, foreground=accent_blue, font=('Segoe UI', 24, 'bold'))
        style.configure('Header.TLabel', background=card_bg, foreground=text_dark, font=('Segoe UI', 16, 'bold'))
        style.configure('SubHeader.TLabel', background=card_bg, foreground=text_gray, font=('Segoe UI', 11))
        # Labels
        style.configure('TLabel', background=card_bg, foreground=text_dark, font=('Segoe UI', 11))
        # Entry, Combobox
        style.configure('TEntry', fieldbackground=card_bg, background=card_bg, foreground=text_dark, bordercolor=border, borderwidth=1, font=('Segoe UI', 11))
        style.configure('TCombobox', fieldbackground=card_bg, background=card_bg, foreground=text_dark, bordercolor=border, borderwidth=1, font=('Segoe UI', 11))
        # Buttons
        style.configure('TButton', background=accent_blue, foreground=card_bg, font=('Segoe UI', 11, 'bold'), borderwidth=0, padding=8)
        style.map('TButton', background=[('active', accent_blue2)], foreground=[('active', card_bg)])
        # Treeview
        style.configure('Treeview', background=card_bg, foreground=text_dark, fieldbackground=card_bg, bordercolor=border, borderwidth=1, font=('Segoe UI', 11))
        style.configure('Treeview.Heading', background=accent_blue, foreground=card_bg, font=('Segoe UI', 11, 'bold'))
        style.map('Treeview', background=[('selected', accent_blue2)], foreground=[('selected', card_bg)])
        style.layout('Treeview', [('Treeview.treearea', {'sticky': 'nswe'})])
        # Alternate row color
        style.configure('Treeview', rowheight=28)
        style.map('Treeview', background=[('alternate', shadow)])
        # Add more padding to all frames
        style.configure('TFrame', padding=12)
        style.configure('TLabelframe', padding=12)
        # For rounded corners and shadow, use a canvas or images for cards if you want to go further.

    def create_main_interface(self):
        """Responsive, modern main interface with grid weights and padding. Fix overlapping."""
        main_frame = ttk.Frame(self.root, padding=20, style='Card.TFrame')
        main_frame.grid(row=0, column=0, sticky="nsew")
        self.root.rowconfigure(0, weight=1)
        self.root.columnconfigure(0, weight=1)
        main_frame.rowconfigure(2, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        self.db_status_var = tk.StringVar(value="No database connected")
        if not hasattr(self, 'db_status_label'):
            self.db_status_label = ttk.Label(main_frame, textvariable=self.db_status_var, style='Header.TLabel')
        self.db_status_label.grid(row=0, column=0, sticky="w", pady=(0, 10))
        self.change_db_btn = ttk.Button(main_frame, text="Connect to Database", command=self.select_database)
        self.change_db_btn.grid(row=0, column=1, sticky="e", padx=(10,0), pady=(0, 10))
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.grid(row=2, column=0, columnspan=2, sticky="nsew", pady=(10, 0))
        main_frame.rowconfigure(2, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        self.create_add_inspection_tab()
        self.create_report_tab()

    def create_add_inspection_tab(self):
        """Add New Inspection tab: Vehicle Info and Technician on the left, photo on the right, all other sections in the center. Disable checkboxes deactivate fields."""
        add_frame = ttk.Frame(self.notebook, style='Card.TFrame')
        self.notebook.add(add_frame, text="Add New Inspection")
        add_frame.grid_rowconfigure(0, weight=1)
        add_frame.grid_columnconfigure(0, weight=1)
        add_frame.grid_columnconfigure(1, weight=0)
        canvas = tk.Canvas(add_frame, borderwidth=0, highlightthickness=0, bg='#ffffff')
        scroll_y = ttk.Scrollbar(add_frame, orient="vertical", command=canvas.yview)
        scroll_x = ttk.Scrollbar(add_frame, orient="horizontal", command=canvas.xview)
        scrollable_frame = ttk.Frame(canvas, style='Card.TFrame')
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        canvas.grid(row=0, column=0, sticky="nsew")
        scroll_y.grid(row=0, column=1, sticky='ns')
        scroll_x.grid(row=1, column=0, sticky='ew')
        add_frame.grid_rowconfigure(0, weight=1)
        add_frame.grid_columnconfigure(0, weight=1)
        # --- Three-column layout: left (vehicle/tech), center (all forms), right (photo) ---
        left_col = ttk.Frame(scrollable_frame)
        left_col.grid(row=1, column=0, sticky='nsew', padx=(0,8), pady=4)
        center_col = ttk.Frame(scrollable_frame)
        center_col.grid(row=1, column=1, sticky='nsew', padx=4, pady=4)
        right_col = ttk.Frame(scrollable_frame)
        right_col.grid(row=1, column=2, sticky='nsew', padx=(8,0), pady=4)
        scrollable_frame.grid_columnconfigure(0, weight=1)
        scrollable_frame.grid_columnconfigure(1, weight=2)
        scrollable_frame.grid_columnconfigure(2, weight=1)
        scrollable_frame.grid_rowconfigure(1, weight=1)
        left_col.grid_rowconfigure(0, weight=1)
        left_col.grid_rowconfigure(1, weight=1)
        center_col.grid_rowconfigure(0, weight=1)
        right_col.grid_rowconfigure(0, weight=1)
        # --- Left: Vehicle Info at top, Technician directly under ---
        vehicle_info_frame = ttk.LabelFrame(left_col, text="Vehicle Info", style='Card.TLabelframe', padding=8)
        vehicle_info_frame.grid(row=0, column=0, sticky="nsew", padx=(0,8))
        vehicle_info_frame.grid_columnconfigure(1, weight=1)
        vehicle_info_frame.grid_rowconfigure(99, weight=1)  # ensure last row expands
        # --- Center and Right columns: ensure all major frames expand ---
        # (Repeat similar grid_rowconfigure/grid_columnconfigure for all major frames in center_col and right_col)
        # --- For all forms, use sticky="ew" for labels/entries and configure columns ---
        # (Repeat for all forms and tables)

        # --- Refresh Button ---
        def refresh_fields(event=None):
            import random
            # Suspension System Standards
            if hasattr(self, 'SuspenstionFrontLeftEfficieny_entry'):
                self.SuspenstionFrontLeftEfficieny_entry.delete(0, tk.END)
                self.SuspenstionFrontLeftEfficieny_entry.insert(0, str(random.randint(41, 100)))
            if hasattr(self, 'SuspenstionFrontRightEfficieny_entry'):
                self.SuspenstionFrontRightEfficieny_entry.delete(0, tk.END)
                self.SuspenstionFrontRightEfficieny_entry.insert(0, str(random.randint(41, 100)))
            if hasattr(self, 'SuspensionFrontRandLDifference_entry'):
                self.SuspensionFrontRandLDifference_entry.delete(0, tk.END)
                self.SuspensionFrontRandLDifference_entry.insert(0, str(random.randint(0, 30)))
            if hasattr(self, 'SuspensionRareLeftEfficency_entry'):
                self.SuspensionRareLeftEfficency_entry.delete(0, tk.END)
                self.SuspensionRareLeftEfficency_entry.insert(0, str(random.randint(41, 100)))
            if hasattr(self, 'SuspenstionRareRightEfficency_entry'):
                self.SuspenstionRareRightEfficency_entry.delete(0, tk.END)
                self.SuspenstionRareRightEfficency_entry.insert(0, str(random.randint(41, 100)))
            if hasattr(self, 'SuspensionrareRandLDifference_entry'):
                self.SuspensionrareRandLDifference_entry.delete(0, tk.END)
                self.SuspensionrareRandLDifference_entry.insert(0, str(random.randint(0, 30)))
            # Braking System Standards
            if hasattr(self, 'MaximumServiceBrakeForceFrontLeft_entry'):
                self.MaximumServiceBrakeForceFrontLeft_entry.delete(0, tk.END)
                self.MaximumServiceBrakeForceFrontLeft_entry.insert(0, str(random.randint(1500, 4000)))
            if hasattr(self, 'MaximumServiceBrakeForceFrontRight_entry'):
                self.MaximumServiceBrakeForceFrontRight_entry.delete(0, tk.END)
                self.MaximumServiceBrakeForceFrontRight_entry.insert(0, str(random.randint(1500, 4000)))
            if hasattr(self, 'MaximumServiceBrakeForceFrontDifference_entry'):
                self.MaximumServiceBrakeForceFrontDifference_entry.delete(0, tk.END)
                self.MaximumServiceBrakeForceFrontDifference_entry.insert(0, str(random.randint(0, 30)))
            if hasattr(self, 'MaximumServiceBrakeForceRareLeft_entry'):
                self.MaximumServiceBrakeForceRareLeft_entry.delete(0, tk.END)
                self.MaximumServiceBrakeForceRareLeft_entry.insert(0, str(random.randint(1000, 3000)))
            if hasattr(self, 'MaximumServiceBrakeForceRareRight_entry'):
                self.MaximumServiceBrakeForceRareRight_entry.delete(0, tk.END)
                self.MaximumServiceBrakeForceRareRight_entry.insert(0, str(random.randint(1000, 3000)))
            if hasattr(self, 'MaximumServiceBrakeForceRareDifference_entry'):
                self.MaximumServiceBrakeForceRareDifference_entry.delete(0, tk.END)
                self.MaximumServiceBrakeForceRareDifference_entry.insert(0, str(random.randint(0, 30)))
            if hasattr(self, 'TotalServiceBrakeEfficiency_entry'):
                self.TotalServiceBrakeEfficiency_entry.delete(0, tk.END)
                self.TotalServiceBrakeEfficiency_entry.insert(0, str(random.randint(50, 100)))
            if hasattr(self, 'TotalParkingBrakeEfficeny_entry'):
                self.TotalParkingBrakeEfficeny_entry.delete(0, tk.END)
                self.TotalParkingBrakeEfficeny_entry.insert(0, str(random.randint(25, 100)))
            if hasattr(self, 'ParkingBrakeLeftRightdifference_entry'):
                self.ParkingBrakeLeftRightdifference_entry.delete(0, tk.END)
                self.ParkingBrakeLeftRightdifference_entry.insert(0, str(random.randint(0, 35)))
            # Lighting System Standards
            if hasattr(self, 'headlight_high_beam_intensity_left_entry'):
                self.headlight_high_beam_intensity_left_entry.delete(0, tk.END)
                self.headlight_high_beam_intensity_left_entry.insert(0, str(random.randint(10000, 20000)))
            if hasattr(self, 'headlight_high_beam_intensity_right_entry'):
                self.headlight_high_beam_intensity_right_entry.delete(0, tk.END)
                self.headlight_high_beam_intensity_right_entry.insert(0, str(random.randint(10000, 20000)))
            if hasattr(self, 'headlight_low_beam_intensity_left_entry'):
                self.headlight_low_beam_intensity_left_entry.delete(0, tk.END)
                self.headlight_low_beam_intensity_left_entry.insert(0, str(random.randint(7000, 15000)))
            if hasattr(self, 'headlight_low_beam_intensity_right_entry'):
                self.headlight_low_beam_intensity_right_entry.delete(0, tk.END)
                self.headlight_low_beam_intensity_right_entry.insert(0, str(random.randint(7000, 15000)))
            
            # Headlight Horizontal and Vertical (2nd row onwards): Range 10-30 with "+0" prefix, right slightly lower
            if hasattr(self, 'headlight_high_beam_horizontal_left_entry'):
                left_value = random.randint(10, 30)
                right_value = max(10, left_value - random.randint(1, 5))  # Right slightly lower
                self.headlight_high_beam_horizontal_left_entry.delete(0, tk.END)
                self.headlight_high_beam_horizontal_left_entry.insert(0, f"+0{left_value}")
                if hasattr(self, 'headlight_high_beam_horizontal_right_entry'):
                    self.headlight_high_beam_horizontal_right_entry.delete(0, tk.END)
                    self.headlight_high_beam_horizontal_right_entry.insert(0, f"+0{right_value}")
            
            if hasattr(self, 'headlight_high_beam_vertical_left_entry'):
                left_value = random.randint(10, 30)
                right_value = max(10, left_value - random.randint(1, 5))  # Right slightly lower
                self.headlight_high_beam_vertical_left_entry.delete(0, tk.END)
                self.headlight_high_beam_vertical_left_entry.insert(0, f"+0{left_value}")
                if hasattr(self, 'headlight_high_beam_vertical_right_entry'):
                    self.headlight_high_beam_vertical_right_entry.delete(0, tk.END)
                    self.headlight_high_beam_vertical_right_entry.insert(0, f"+0{right_value}")
            
            if hasattr(self, 'headlight_low_beam_horizontal_left_entry'):
                left_value = random.randint(10, 30)
                right_value = max(10, left_value - random.randint(1, 5))  # Right slightly lower
                self.headlight_low_beam_horizontal_left_entry.delete(0, tk.END)
                self.headlight_low_beam_horizontal_left_entry.insert(0, f"+0{left_value}")
                if hasattr(self, 'headlight_low_beam_horizontal_right_entry'):
                    self.headlight_low_beam_horizontal_right_entry.delete(0, tk.END)
                    self.headlight_low_beam_horizontal_right_entry.insert(0, f"+0{right_value}")
            
            if hasattr(self, 'headlight_low_beam_vertical_left_entry'):
                left_value = random.randint(10, 30)
                right_value = max(10, left_value - random.randint(1, 5))  # Right slightly lower
                self.headlight_low_beam_vertical_left_entry.delete(0, tk.END)
                self.headlight_low_beam_vertical_left_entry.insert(0, f"+0{left_value}")
                if hasattr(self, 'headlight_low_beam_vertical_right_entry'):
                    self.headlight_low_beam_vertical_right_entry.delete(0, tk.END)
                    self.headlight_low_beam_vertical_right_entry.insert(0, f"+0{right_value}")
            # Emissions Standards (Gasoline)
            # Determine catalyst state
            catalyst = getattr(self, 'benzine_radio', None)
            with_catalyst = catalyst and catalyst.get() == 'withcc'
            # HC
            hc_min, hc_max = (300, 1200) if with_catalyst else (800, 1600)
            for entry in [getattr(self, 'gas_hc_left_entry', None), getattr(self, 'gas_hc_right_entry', None)]:
                if entry:
                    entry.delete(0, tk.END)
                    entry.insert(0, str(random.randint(hc_min, hc_max)))
            # CO
            co_min, co_max = (0.1, 0.75) if with_catalyst else (2.0, 4.5)
            for entry in [getattr(self, 'gas_co_left_entry', None), getattr(self, 'gas_co_right_entry', None)]:
                if entry:
                    entry.delete(0, tk.END)
                    entry.insert(0, f"{random.uniform(co_min, co_max):.2f}")
            # CO2
            for entry in [getattr(self, 'gas_co2_left_entry', None), getattr(self, 'gas_co2_right_entry', None)]:
                if entry:
                    entry.delete(0, tk.END)
                    entry.insert(0, f"{random.uniform(10, 20):.2f}")
            # O2
            for entry in [getattr(self, 'gas_o2_left_entry', None), getattr(self, 'gas_o2_right_entry', None)]:
                if entry:
                    entry.delete(0, tk.END)
                    entry.insert(0, f"{random.uniform(0.5, 5):.2f}")
            # Lambda
            for entry in [getattr(self, 'gas_lambda_left_entry', None), getattr(self, 'gas_lambda_right_entry', None)]:
                if entry:
                    entry.delete(0, tk.END)
                    entry.insert(0, f"{random.uniform(0.9, 1.1):.2f}")
            # Diesel N (Opacity)
            turbo = getattr(self, 'diesel_radio', None)
            with_turbo = turbo and turbo.get() == 'with_turbo'
            n_min, n_max = (1.0, 4.5) if with_turbo else (1.0, 3.8)
            n_entry = getattr(self, 'gas_n_entry', None)
            if n_entry:
                n_entry.delete(0, tk.END)
                n_entry.insert(0, f"{random.uniform(n_min, n_max):.2f}")
            # Diesel K (Absorption)
            k_entry = getattr(self, 'gas_k_entry', None)
            if k_entry:
                k_entry.delete(0, tk.END)
                k_entry.insert(0, f"{random.uniform(0, 1.5):.2f}")
            # Weight & Friction Standards
            if hasattr(self, 'front_axle_weight_entry'):
                self.front_axle_weight_entry.delete(0, tk.END)
                self.front_axle_weight_entry.insert(0, str(random.randint(600, 1200)))
            if hasattr(self, 'rear_axle_weight_entry'):
                self.rear_axle_weight_entry.delete(0, tk.END)
                self.rear_axle_weight_entry.insert(0, str(random.randint(400, 1000)))
            if hasattr(self, 'total_vehicle_weight_entry'):
                self.total_vehicle_weight_entry.delete(0, tk.END)
                self.total_vehicle_weight_entry.insert(0, str(random.randint(1000, 2200)))
            
            # Synchronize vehicle info weight with brake result total weight
            vehicle_weight = str(random.randint(950, 1100))
            if hasattr(self, 'weight_entry'):
                self.weight_entry.delete(0, tk.END)
                self.weight_entry.insert(0, vehicle_weight)
            if 'brake_total_weight_col1_entry' in self.brake_entries:
                self.brake_entries['brake_total_weight_col1_entry'].delete(0, tk.END)
                self.brake_entries['brake_total_weight_col1_entry'].insert(0, vehicle_weight)
            # Synchronize the 4th column of Total Weight in brake result with vehicle info weight
            if hasattr(self, 'weight_entry'):
                vehicle_weight = self.weight_entry.get()
                entry_name = 'brake_total_weight_col4_entry'
                if entry_name in self.brake_entries:
                    self.brake_entries[entry_name].delete(0, tk.END)
                    self.brake_entries[entry_name].insert(0, vehicle_weight)
            # Remove old individual assignments for brake_roller_friction_col1_entry, col3, col6
            # Add new loop for all columns
            for c in range(1, 11):
                entry_name = f"brake_roller_friction_col{c}_entry"
                if entry_name in self.brake_entries:
                    self.brake_entries[entry_name].delete(0, tk.END)
                    if c == 6:
                        self.brake_entries[entry_name].insert(0, str(random.randint(30, 35)))
                    elif c ==4 :
                        self.brake_entries[entry_name].insert(0, str(" "))

                    else:
                        # Higher probability for 2 and 3
                        weights = [0.1, 0.4, 0.4, 0.1]  # Probability for 1, 2, 3, 4
                        value = random.choices([1, 2, 3, 4], weights=weights)[0]
                        self.brake_entries[entry_name].insert(0, str(value))
            
            # --- Brake Result Standards ---
            # Brake Force (N): Front: 1500-4000 | Rear: 1000-3000
            svc_brake_front_left = random.randint(140, 250)
            svc_brake_front_right = random.randint(85, min(170, svc_brake_front_left - 1))
            svc_front_difference = int(100 - (svc_brake_front_right / svc_brake_front_left) * 100)
            svc_brake_rear_left = random.randint(60, 130)
            svc_brake_rear_right = random.randint(35, 105)
            svc_brake_differnce = int(100 - (svc_brake_rear_right / svc_brake_rear_left) * 100)
            for c in [1]:  # Front Left, Front Right
                entry_name = f"brake_brake_force_col{c}_entry"
                if entry_name in self.brake_entries:
                    self.brake_entries[entry_name].delete(0, tk.END)
                    self.brake_entries[entry_name].insert(0, str(svc_brake_front_left))
            for c in [2]:  # Front Left, Front Right
                entry_name = f"brake_brake_force_col{c}_entry"
                if entry_name in self.brake_entries:
                    self.brake_entries[entry_name].delete(0, tk.END)
                    self.brake_entries[entry_name].insert(0, svc_front_difference)
            for c in [3]:  # Front Left, Front Right
                entry_name = f"brake_brake_force_col{c}_entry"
                if entry_name in self.brake_entries:
                    self.brake_entries[entry_name].delete(0, tk.END)
                    self.brake_entries[entry_name].insert(0, str(svc_brake_front_right))
            
            # ServiceBrakeForceFrontLeft: 140-250
            for c in [5]:  # Rear Left, Rear Right
                entry_name = f"brake_brake_force_col{c}_entry"
                if entry_name in self.brake_entries:
                    self.brake_entries[entry_name].delete(0, tk.END)
                    self.brake_entries[entry_name].insert(0, str(svc_brake_rear_left))
            for c in [6]:  # Rear Left, Rear Right
                entry_name = f"brake_brake_force_col{c}_entry"
                if entry_name in self.brake_entries:
                    self.brake_entries[entry_name].delete(0, tk.END)
                    self.brake_entries[entry_name].insert(0, str(svc_brake_differnce))
            for c in [7]:  # Rear Left, Rear Right
                entry_name = f"brake_brake_force_col{c}_entry"
                if entry_name in self.brake_entries:
                    self.brake_entries[entry_name].delete(0, tk.END)
                    self.brake_entries[entry_name].insert(0, str(svc_brake_rear_right))
            for c in [8, 10 ]:  # Rear Left, Rear Right
                entry_name = f"brake_brake_force_col{c}_entry"
                if entry_name in self.brake_entries:
                    self.brake_entries[entry_name].delete(0, tk.END)
                    self.brake_entries[entry_name].insert(0, str(random.randint(3,7)))
            max_svc_brake_front_left = random.randint(190, 290)
            max_svc_brake_front_right = random.randint(110, min(175, max_svc_brake_front_left - 1))
            mx_front_di = int(100 - (max_svc_brake_front_right / max_svc_brake_front_left) * 100)
            max_svc_brake_rear_left = random.randint(130, 205)
            max_svc_brake_rear_right = random.randint(90, min(120, max_svc_brake_rear_left - 1))
            mx_rear_di = int(100 - (max_svc_brake_rear_right / max_svc_brake_rear_left) * 100)
            max_hn_brake_rear_left = random.randint(75, 95)
            max_hn_brake_rear_right = random.randint(55, min(78, max_svc_brake_rear_left - 1))
            mx_hn_di = int(100 - (max_svc_brake_rear_right / max_svc_brake_rear_left) * 100)
            for c in [1]:  # All brake force columns
                max_brake_force_entry = f"brake_max_brake_force_col{c}_entry"
                if max_brake_force_entry in self.brake_entries:
                    self.brake_entries[max_brake_force_entry].delete(0, tk.END)
                    self.brake_entries[max_brake_force_entry].insert(0, str(max_svc_brake_front_left))
            for c in [2]:  # All brake force columns               
                max_brake_force_entry = f"brake_max_brake_force_col{c}_entry"
                if  max_brake_force_entry in self.brake_entries:
                    self.brake_entries[max_brake_force_entry].delete(0, tk.END)
                    self.brake_entries[max_brake_force_entry].insert(0, str(mx_front_di))
            for c in [3]:  # All brake force columns
                max_brake_force_entry = f"brake_max_brake_force_col{c}_entry"
                if max_brake_force_entry in self.brake_entries:
                    self.brake_entries[max_brake_force_entry].delete(0, tk.END)
                    self.brake_entries[max_brake_force_entry].insert(0, str(max_svc_brake_front_right))
            for c in [5]:  # All brake force columns
                max_brake_force_entry = f"brake_max_brake_force_col{c}_entry"
                if max_brake_force_entry in self.brake_entries:
                    self.brake_entries[max_brake_force_entry].delete(0, tk.END)
                    self.brake_entries[max_brake_force_entry].insert(0, str(max_svc_brake_rear_left))
            for c in [6]:  # All brake force columns
                max_brake_force_entry = f"brake_max_brake_force_col{c}_entry"
                if max_brake_force_entry in self.brake_entries:
                    self.brake_entries[max_brake_force_entry].delete(0, tk.END)
                    self.brake_entries[max_brake_force_entry].insert(0, str(mx_rear_di))
            for c in [7]:
                max_brake_force_entry = f"brake_max_brake_force_col{c}_entry"
                if max_brake_force_entry in self.brake_entries:
                    self.brake_entries[max_brake_force_entry].delete(0, tk.END)
                    self.brake_entries[max_brake_force_entry].insert(0, str(max_svc_brake_rear_right))
            for c in [8]:  
                max_brake_force_entry = f"brake_max_brake_force_col{c}_entry"
                if max_brake_force_entry in self.brake_entries:
                    self.brake_entries[max_brake_force_entry].delete(0, tk.END)
                    self.brake_entries[max_brake_force_entry].insert(0, str(max_hn_brake_rear_left))
            for c in [9]: 
                max_brake_force_entry = f"brake_max_brake_force_col{c}_entry"
                if max_brake_force_entry in self.brake_entries:
                    self.brake_entries[max_brake_force_entry].delete(0, tk.END)
                    self.brake_entries[max_brake_force_entry].insert(0, str(mx_hn_di))
            for c in [10]:  
                max_brake_force_entry = f"brake_max_brake_force_col{c}_entry"
                if max_brake_force_entry in self.brake_entries:
                    self.brake_entries[max_brake_force_entry].delete(0, tk.END)
                    self.brake_entries[max_brake_force_entry].insert(0, str(max_hn_brake_rear_right))            
            for c in [1, 3, 5, 7]:  # All ovality columns
                entry_name = f"brake_ovality_col{c}_entry"
                if entry_name in self.brake_entries:
                    self.brake_entries[entry_name].delete(0, tk.END)
                    self.brake_entries[entry_name].insert(0, f"{random.uniform(0.0, 0.5):.2f}")
          
            # Axle Weight (kg): Front: 600-1200 | Rear: 400-1000
            front = random.randint(450, 600)
            rear = random.randint(300 ,449)
            for c in [4]:  # Front Left, Front Right
                entry_name = f"brake_axle_weight_col{c}_entry"
                if entry_name in self.brake_entries:
                    self.brake_entries[entry_name].delete(0, tk.END)
                    self.brake_entries[entry_name].insert(0, str(front))
            for c in [7]:  # Front Left, Front Right
                entry_name = f"brake_axle_weight_col{c}_entry"
                if entry_name in self.brake_entries:
                    self.brake_entries[entry_name].delete(0, tk.END)
                    self.brake_entries[entry_name].insert(0, str(rear))
           
            # Axle Deceleration (m/s²): 4.0-7.0 (≥3.92 m/s² for 50% efficiency)
            fronts = random.randint(55, 65)
            rears = random.randint(65 ,79)
            for c in [4]:  # All deceleration columns
                entry_name = f"brake_axle_deceleration_col{c}_entry"
                if entry_name in self.brake_entries:
                    self.brake_entries[entry_name].delete(0, tk.END)
                    self.brake_entries[entry_name].insert(0, str(fronts))
            for c in [7]:  # All deceleration columns
                entry_name = f"brake_axle_deceleration_col{c}_entry"
                if entry_name in self.brake_entries:
                    self.brake_entries[entry_name].delete(0, tk.END)
                    self.brake_entries[entry_name].insert(0, str(rears))
            
            # Total Brake Force (N): Sum of axles (≥50% of total weight * 9.81)
            total_brake_force = 0
            for c in [1, 3, 5, 7]:  # Sum all brake forces
                entry_name = f"brake_brake_force_col{c}_entry"
                if entry_name in self.brake_entries:
                    total_brake_force += int(self.brake_entries[entry_name].get() or 0)
            
            # Ensure minimum efficiency (≥50% of total weight * 9.81)
            vehicle_weight = int(self.weight_entry.get() or 1500)
            min_required_force = int(vehicle_weight * 9.81 * 0.5)
            if total_brake_force < min_required_force:
                total_brake_force = min_required_force
            
            # Set total brake force in column 1
            total_brake_force_entry = f"brake_total_brake_force_col1_entry"
            if total_brake_force_entry in self.brake_entries:
                self.brake_entries[total_brake_force_entry].delete(0, tk.END)
                self.brake_entries[total_brake_force_entry].insert(0, str(total_brake_force))
            
            # Side Slip (m/km): -5.0 to 5.0 (alignment standard)
            front_side = random.randint(-6, 6)
            rear_side = random.randint(-6,6)
            for c in [4]:  # Front and Rear axle side slip
                entry_name = f"brake_side_slip_col{c}_entry"
                if entry_name in self.brake_entries:
                    self.brake_entries[entry_name].delete(0, tk.END)
                    self.brake_entries[entry_name].insert(0, str(front_side))
            for c in [7]:  # Front and Rear axle side slip
                entry_name = f"brake_side_slip_col{c}_entry"
                if entry_name in self.brake_entries:
                    self.brake_entries[entry_name].delete(0, tk.END)
                    self.brake_entries[entry_name].insert(0, str(rear_side))

            # Set the 4th column of Total Brake Force in brake result to the sum of Brake Force fields
            brake_force_sum = 0
            for c in [1, 3, 5, 7]:
                entry_name = f'brake_brake_force_col{c}_entry'
                if entry_name in self.brake_entries:
                    try:
                        brake_force_sum += int(self.brake_entries[entry_name].get())
                    except ValueError:
                        pass
            entry_name = 'brake_total_brake_force_col4_entry'
            if entry_name in self.brake_entries:
                self.brake_entries[entry_name].delete(0, tk.END)
                self.brake_entries[entry_name].insert(0, str(brake_force_sum))

            # --- SHOCK/SUSPENSION SECTION RANDOMIZATION ---
            # Amplitude (mm): 5-15 (columns 1,3,5,7)
            for c in [1, 3, 5, 7]:
                entry_name = f"shock_amplitude_col{c}_entry"
                if entry_name in self.shock_entries:
                    self.shock_entries[entry_name].delete(0, tk.END)
                    self.shock_entries[entry_name].insert(0, str(random.randint(5, 15)))
            # Effect (%): 41-100 (columns 2,4,6,8)
            for c in [2, 4, 6, 8]:
                entry_name = f"shock_effect_col{c}_entry"
                if entry_name in self.shock_entries:
                    self.shock_entries[entry_name].delete(0, tk.END)
                    self.shock_entries[entry_name].insert(0, str(random.randint(41, 100)))
            # Wheel Weights (kg):
            # Front Axle (columns 1,3): 600-1200
            for c in [1, 3]:
                entry_name = f"shock_wheel_weights_col{c}_entry"
                if entry_name in self.shock_entries:
                    self.shock_entries[entry_name].delete(0, tk.END)
                    self.shock_entries[entry_name].insert(0, str(random.randint(180, 300)))
            # Rear Axle (columns 5,7): 400-1000
            for c in [5, 7]:
                entry_name = f"shock_wheel_weights_col{c}_entry"
                if entry_name in self.shock_entries:
                    self.shock_entries[entry_name].delete(0, tk.END)
                    self.shock_entries[entry_name].insert(0, str(random.randint(180, 300)))

        refresh_btn = ttk.Button(scrollable_frame, text="Refresh", command=refresh_fields, width=8)
        refresh_btn.grid(row=0, column=0, sticky='ew', pady=4)
        def reset_form():
            self.clear_form()
        reset_btn = ttk.Button(scrollable_frame, text="Reset Form", command=reset_form, width=10)
        reset_btn.grid(row=0, column=1, sticky='ew', pady=4)
        self.root.bind('<Control-F1>', refresh_fields)

        # --- Three-column layout: left (vehicle/tech), center (all forms), right (photo) ---
        left_col = ttk.Frame(scrollable_frame)
        left_col.grid(row=1, column=0, sticky='nsew', padx=(0,8), pady=4)
        center_col = ttk.Frame(scrollable_frame)
        center_col.grid(row=1, column=1, sticky='nsew', padx=4, pady=4)
        right_col = ttk.Frame(scrollable_frame)
        right_col.grid(row=1, column=2, sticky='nsew', padx=(8,0), pady=4)
        scrollable_frame.grid_columnconfigure(0, weight=1)
        scrollable_frame.grid_columnconfigure(1, weight=2)
        scrollable_frame.grid_columnconfigure(2, weight=1)
        scrollable_frame.grid_rowconfigure(1, weight=1)
        left_col.grid_rowconfigure(0, weight=1)
        left_col.grid_rowconfigure(1, weight=1)
        center_col.grid_rowconfigure(0, weight=1)
        right_col.grid_rowconfigure(0, weight=1)

        # --- Left: Vehicle Info at top, Technician directly under ---
        vehicle_info_frame = ttk.LabelFrame(left_col, text="Vehicle Info", style='Card.TLabelframe', padding=8)
        vehicle_info_frame.grid(row=0, column=0, sticky="nsew", padx=(0,8))
        vehicle_info_frame.grid_columnconfigure(1, weight=1)
        vehicle_info_frame.grid_rowconfigure(99, weight=1)  # ensure last row expands
        # --- Center and Right columns: ensure all major frames expand ---
        # (Repeat similar grid_rowconfigure/grid_columnconfigure for all major frames in center_col and right_col)
        # --- For all forms, use sticky="ew" for labels/entries and configure columns ---
        # (Repeat for all forms and tables)

        # Add reg.date label at the top
        self.reg_date_str = tk.StringVar()
        self.reg_date_str.set(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        ttk.Label(vehicle_info_frame, text="reg.date:", style='TLabel').grid(row=0, column=0, sticky='w', pady=1, padx=(0,2))
        self.reg_date_label = ttk.Label(vehicle_info_frame, textvariable=self.reg_date_str, style='TLabel')
        self.reg_date_label.grid(row=0, column=1, sticky='ew', pady=1)
        row_idx = 1
        # Plate No.
        ttk.Label(vehicle_info_frame, text="Plate No.:", style='TLabel').grid(row=row_idx, column=0, sticky='w', pady=1, padx=(0,2))
        self.plate_no_entry = ttk.Entry(vehicle_info_frame, width=18)
        self.plate_no_entry.grid(row=row_idx, column=1, sticky='ew', pady=1)
        self.plate_no_entry.bind('<Return>', lambda event: self.load_plate_image())
        row_idx += 1
        # Chan
        ttk.Label(vehicle_info_frame, text="cust:", style='TLabel').grid(row=row_idx, column=0, sticky='w', pady=1, padx=(0,2))
        self.cust_entry = ttk.Entry(vehicle_info_frame, width=18)
        self.cust_entry.grid(row=row_idx, column=1, sticky='ew', pady=1)
        row_idx += 1
        ttk.Label(vehicle_info_frame, text="Chan.:", style='TLabel').grid(row=row_idx, column=0, sticky='w', pady=1, padx=(0,2))
        self.chan_entry = ttk.Entry(vehicle_info_frame, width=18)
        self.chan_entry.grid(row=row_idx, column=1, sticky='ew', pady=1)
        row_idx += 1
        ttk.Label(vehicle_info_frame, text="Eng:", style='TLabel').grid(row=row_idx, column=0, sticky='w', pady=1, padx=(0,2))
        self.eng_entry = ttk.Entry(vehicle_info_frame, width=18)
        self.eng_entry.grid(row=row_idx, column=1, sticky='ew', pady=1)
        row_idx += 1
        ttk.Label(vehicle_info_frame, text="Model:", style='TLabel').grid(row=row_idx, column=0, sticky='w', pady=1, padx=(0,2))
        self.model_entry = ttk.Entry(vehicle_info_frame, width=18)
        self.model_entry.grid(row=row_idx, column=1, sticky='ew', pady=1)
        row_idx += 1
        ttk.Label(vehicle_info_frame, text="libre:", style='TLabel').grid(row=row_idx, column=0, sticky='w', pady=1, padx=(0,2))
        self.libre_entry = ttk.Entry(vehicle_info_frame, width=18)
        self.libre_entry.grid(row=row_idx, column=1, sticky='ew', pady=1)
        row_idx += 1
        ttk.Label(vehicle_info_frame, text="Fuel:", style='TLabel').grid(row=row_idx, column=0, sticky='w', pady=1, padx=(0,2))
        self.fuel_var = tk.StringVar()
        self.fuel_entry = ttk.Combobox(vehicle_info_frame, textvariable=self.fuel_var, values=["BENZINE", "DIESEL", "ELECTRIC"], state="readonly", width=16)
        self.fuel_entry.current(0)
        self.fuel_entry.grid(row=row_idx, column=1, sticky='ew', pady=1)
        row_idx += 1
        ttk.Label(vehicle_info_frame, text="px/sx:", style='TLabel').grid(row=row_idx, column=0, sticky='w', pady=1, padx=(0,2))
        self.pxsx1_var = tk.StringVar()
        pxsx1_options = ["1", "2", "3", "4", "5", "AU", "U.N", "ተላላፊ", "ፖሊስ"]
        self.pxsx1_entry = ttk.Combobox(vehicle_info_frame, textvariable=self.pxsx1_var, values=pxsx1_options, state="readonly", width=8)
        self.pxsx1_entry.current(0)
        self.pxsx1_entry.grid(row=row_idx, column=1, sticky='w', pady=1)
        self.pxsx2_var = tk.StringVar()
        pxsx2_options = ["ET", "AA", "OR", "AF", "AM", "BG", "DR", "GM", "HR", "SM"]
        self.pxsx2_entry = ttk.Combobox(vehicle_info_frame, textvariable=self.pxsx2_var, values=pxsx2_options, state="readonly", width=8)
        self.pxsx2_entry.current(0)
        self.pxsx2_entry.grid(row=row_idx, column=1, sticky='e', pady=1)
        row_idx += 1
        ttk.Label(vehicle_info_frame, text="MADE IN:", style='TLabel').grid(row=row_idx, column=0, sticky='w', pady=1, padx=(0,2))
        self.made_in_var = tk.StringVar()
        made_in_options = ["JAPAN", "ETHIOPIA", "CHINA", "USA", "GERMANY", "UK", "RUSSIA", "ITALY","INDIA","SOUTH AFRICA","KOREA","OTHER"]
        self.made_in_as_entry = ttk.Combobox(vehicle_info_frame, textvariable=self.made_in_var, values=made_in_options, state="readonly", width=16)
        self.made_in_as_entry.current(0)
        self.made_in_as_entry.grid(row=row_idx, column=1, sticky='ew', pady=1)
        row_idx += 1
        ttk.Label(vehicle_info_frame, text="Year:", style='TLabel').grid(row=row_idx, column=0, sticky='w', pady=1, padx=(0,2))
        self.year_var = tk.StringVar()
        year_options = [str(y) for y in range(1950, 2051)]
        self.year_entry = ttk.Combobox(vehicle_info_frame, textvariable=self.year_var, values=year_options, state="readonly", width=16)
        self.year_entry.current(0)
        self.year_entry.grid(row=row_idx, column=1, sticky='ew', pady=1)
        row_idx += 1
        # F. No.
        ttk.Label(vehicle_info_frame, text="File Place:", style='TLabel').grid(row=row_idx, column=0, sticky='w', pady=1, padx=(0,2))
        self.file_place_entry = ttk.Entry(vehicle_info_frame, width=18)
        self.file_place_entry.grid(row=row_idx, column=1, sticky='ew', pady=1)
        self.file_place_entry.insert(0, 'ADDIS')
        row_idx += 1
        ttk.Label(vehicle_info_frame, text="Operator:", style='TLabel').grid(row=row_idx, column=0, sticky='w', pady=1, padx=(0,2))
        self.operator_entry = ttk.Entry(vehicle_info_frame, width=18)
        self.operator_entry.grid(row=row_idx, column=1, sticky='ew', pady=1)
        row_idx += 1
        ttk.Label(vehicle_info_frame, text="TYPE:", style='TLabel').grid(row=row_idx, column=0, sticky='w', pady=1, padx=(0,2))
        self.type_var = tk.StringVar()
        type_options = ["AUTOMOBILE", "YEMESK/DIRIB", "DEREK CHINET", "YEHIZIB"]
        self.type_entry = ttk.Combobox(vehicle_info_frame, textvariable=self.type_var, values=type_options, state="readonly", width=16)
        self.type_entry.current(0)
        self.type_entry.grid(row=row_idx, column=1, sticky='ew', pady=1)
        row_idx += 1
        ttk.Label(vehicle_info_frame, text="Weight:", style='TLabel').grid(row=row_idx, column=0, sticky='w', pady=1, padx=(0,2))
        self.weight_entry = ttk.Entry(vehicle_info_frame, width=18)
        self.weight_entry.grid(row=row_idx, column=1, sticky='ew', pady=1)
        # Fuel (dropdown)
        
       
        # Year (dropdown)
        
        # MADE IN (dropdown)
        
        # TYPE (dropdown)
        
        
        
        
        
        # Add event handler to sync vehicle weight with brake total weight
        def sync_vehicle_weight_to_brake(*args):
            if hasattr(self, 'brake_entries') and 'brake_total_weight_col1_entry' in self.brake_entries:
                brake_entry = self.brake_entries['brake_total_weight_col1_entry']
                vehicle_weight = self.weight_entry.get()
                brake_entry.delete(0, tk.END)
                brake_entry.insert(0, vehicle_weight)
        
        self.weight_entry.bind('<KeyRelease>', sync_vehicle_weight_to_brake)
        self.weight_entry.bind('<FocusOut>', sync_vehicle_weight_to_brake)
        row_idx += 1
        
        
        
        
        
       
        
        # Pass/Fail dropdown
        ttk.Label(vehicle_info_frame, text="Result:", style='TLabel').grid(row=row_idx, column=0, sticky='w', pady=1, padx=(0,2))
        self.result_var = tk.StringVar()
        result_combo = ttk.Combobox(vehicle_info_frame, textvariable=self.result_var, values=["Pass", "Fail"], state="readonly", width=8)
        result_combo.grid(row=row_idx, column=1, sticky='ew', pady=1)
        row_idx += 1
        ttk.Button(vehicle_info_frame, text="Save", command=self.save_inspection).grid(row=row_idx, column=0, columnspan=2, sticky='ew', pady=4)
        row_idx += 1
        # Add Technician & Checklist fields under Save button
        ttk.Label(vehicle_info_frame, text="Technician:", style='TLabel').grid(row=row_idx, column=0, sticky='w', pady=1, padx=(0,2))
        self.technician_entry = ttk.Entry(vehicle_info_frame, width=16)
        self.technician_entry.grid(row=row_idx, column=1, sticky='ew', pady=1)
        row_idx += 1
        ttk.Label(vehicle_info_frame, text="Person:", style='TLabel').grid(row=row_idx, column=0, sticky='w', pady=1, padx=(0,2))
        self.person_entry = ttk.Entry(vehicle_info_frame, width=16)
        self.person_entry.grid(row=row_idx, column=1, sticky='ew', pady=1)
        row_idx += 1
        ttk.Label(vehicle_info_frame, text="Kuntal:", style='TLabel').grid(row=row_idx, column=0, sticky='w', pady=1, padx=(0,2))
        self.kuntal_entry = ttk.Entry(vehicle_info_frame, width=16)
        self.kuntal_entry.grid(row=row_idx, column=1, sticky='ew', pady=1)
        row_idx += 1
        ttk.Label(vehicle_info_frame, text="Litter:", style='TLabel').grid(row=row_idx, column=0, sticky='w', pady=1, padx=(0,2))
        self.litter_entry = ttk.Entry(vehicle_info_frame, width=16)
        self.litter_entry.grid(row=row_idx, column=1, sticky='ew', pady=1)
        row_idx += 1
        ttk.Button(vehicle_info_frame, text="Checklist", command=self.clear_form).grid(row=row_idx, column=0, columnspan=2, pady=1, sticky='ew')
        row_idx += 1
        ttk.Label(vehicle_info_frame, text="Total Pay:", style='TLabel').grid(row=row_idx, column=0, sticky='w', pady=1, padx=(0,2))
        self.total_pay_entry = ttk.Entry(vehicle_info_frame, width=16)
        self.total_pay_entry.grid(row=row_idx, column=1, sticky='ew', pady=1)
        vehicle_info_frame.grid_rowconfigure(row_idx+1, weight=1)
        ttk.Label(vehicle_info_frame, text="Defects/Notes:", style='TLabel').grid(row=row_idx, column=0, sticky='nw', pady=1, padx=(0,2))
        self.defects_text = tk.Text(vehicle_info_frame, width=18, height=3)
        self.defects_text.grid(row=row_idx, column=1, sticky='ew', pady=1)
        row_idx += 1

        # --- Center: All other sections stacked ---
        self.random_entries = []
        # --- GAS/SMOKE and HEADLIGHT side by side ---
        top_center_frame = ttk.Frame(center_col)
        top_center_frame.grid(row=0, column=0, sticky="nsew", pady=(0,8))
        top_center_frame.grid_columnconfigure(0, weight=1)
        top_center_frame.grid_columnconfigure(1, weight=1)
        # --- GAS/SMOKE ---
        gas_frame = ttk.LabelFrame(top_center_frame, style='Card.TLabelframe', padding=8)
        gas_frame.grid(row=0, column=0, sticky="nsew", padx=(0,4))
        gas_frame.grid_columnconfigure(0, weight=1)
        gas_title_frame = ttk.Frame(gas_frame)
        gas_title_frame.grid(row=0, column=0, columnspan=3, sticky='w')
        ttk.Label(gas_title_frame, text="Gas/Smoke Result", style='Card.TLabelframe.Label').pack(side='left')
        self.disable_gas_var = tk.BooleanVar()
        def toggle_gas_fields():
            state = 'disabled' if self.disable_gas_var.get() else 'normal'
            for w in [self.gas_hc_left_entry, self.gas_hc_right_entry, self.gas_co_left_entry, self.gas_co_right_entry, self.gas_co2_left_entry, self.gas_co2_right_entry, self.gas_o2_left_entry, self.gas_o2_right_entry, self.gas_lambda_left_entry, self.gas_lambda_right_entry, self.gas_n_entry, self.gas_k_entry]:
                w.config(state=state)
        ttk.Checkbutton(gas_title_frame, text="Disable", variable=self.disable_gas_var, command=toggle_gas_fields).pack(side='left', padx=(8,0))
        ttk.Label(gas_frame, text="Benzine..", style='TLabel').grid(row=1, column=0, sticky='w')
        # Benzine fields
        ttk.Label(gas_frame, text="HC", style='TLabel').grid(row=2, column=0, sticky='w', padx=(10,2))
        self.gas_hc_left_entry = ttk.Entry(gas_frame, width=5)
        self.gas_hc_left_entry.grid(row=2, column=1, padx=1)
        self.gas_hc_right_entry = ttk.Entry(gas_frame, width=5)
        self.gas_hc_right_entry.grid(row=2, column=2, padx=1)
        ttk.Label(gas_frame, text="CO", style='TLabel').grid(row=3, column=0, sticky='w', padx=(10,2))
        self.gas_co_left_entry = ttk.Entry(gas_frame, width=5)
        self.gas_co_left_entry.grid(row=3, column=1, padx=1)
        self.gas_co_right_entry = ttk.Entry(gas_frame, width=5)
        self.gas_co_right_entry.grid(row=3, column=2, padx=1)
        ttk.Label(gas_frame, text="CO2", style='TLabel').grid(row=4, column=0, sticky='w', padx=(10,2))
        self.gas_co2_left_entry = ttk.Entry(gas_frame, width=5)
        self.gas_co2_left_entry.grid(row=4, column=1, padx=1)
        self.gas_co2_right_entry = ttk.Entry(gas_frame, width=5)
        self.gas_co2_right_entry.grid(row=4, column=2, padx=1)
        ttk.Label(gas_frame, text="O2", style='TLabel').grid(row=5, column=0, sticky='w', padx=(10,2))
        self.gas_o2_left_entry = ttk.Entry(gas_frame, width=5)
        self.gas_o2_left_entry.grid(row=5, column=1, padx=1)
        self.gas_o2_right_entry = ttk.Entry(gas_frame, width=5)
        self.gas_o2_right_entry.grid(row=5, column=2, padx=1)
        ttk.Label(gas_frame, text="Lambda", style='TLabel').grid(row=6, column=0, sticky='w', padx=(10,2))
        self.gas_lambda_left_entry = ttk.Entry(gas_frame, width=5)
        self.gas_lambda_left_entry.grid(row=6, column=1, padx=1)
        self.gas_lambda_right_entry = ttk.Entry(gas_frame, width=5)
        self.gas_lambda_right_entry.grid(row=6, column=2, padx=1)
        self.benzine_radio = tk.StringVar(value='wocc')
        self.gas_benzine_radio1 = ttk.Radiobutton(gas_frame, text="W.o.C.C HC<=1600 CO<=4.5", variable=self.benzine_radio, value='wocc')
        self.gas_benzine_radio2 = ttk.Radiobutton(gas_frame, text="With.C.C HC<=1200 CO<=0.75", variable=self.benzine_radio, value='withcc')
        self.gas_benzine_radio1.grid(row=8, column=0, columnspan=3, sticky='w', padx=(10,2))
        self.gas_benzine_radio2.grid(row=9, column=0, columnspan=3, sticky='w', padx=(10,2))
        ttk.Label(gas_frame, text="Diesel..", style='TLabel').grid(row=10, column=0, sticky='w', pady=(6,0))
        ttk.Label(gas_frame, text="N", style='TLabel').grid(row=11, column=0, sticky='w', padx=(10,2))
        self.gas_n_entry = ttk.Entry(gas_frame, width=5)
        self.gas_n_entry.grid(row=11, column=1, padx=1)
        ttk.Label(gas_frame, text="K", style='TLabel').grid(row=12, column=0, sticky='w', padx=(10,2))
        self.gas_k_entry = ttk.Entry(gas_frame, width=5)
        self.gas_k_entry.grid(row=12, column=1, padx=1)
        self.diesel_radio = tk.StringVar(value='no_turbo')
        self.gas_diesel_radio1 = ttk.Radiobutton(gas_frame, text="Without Turbo Charger<=3.8", variable=self.diesel_radio, value='no_turbo')
        self.gas_diesel_radio2 = ttk.Radiobutton(gas_frame, text="With Turbo Charger<=4.5", variable=self.diesel_radio, value='with_turbo')
        self.gas_diesel_radio1.grid(row=13, column=0, columnspan=2, sticky='w', padx=(10,2))
        self.gas_diesel_radio2.grid(row=14, column=0, columnspan=2, sticky='w', padx=(10,2))
        gas_frame.grid_rowconfigure(15, weight=1)

        # --- HEADLIGHT ---
        headlight_frame = ttk.LabelFrame(top_center_frame, style='Card.TLabelframe', padding=8)
        headlight_frame.grid(row=0, column=1, sticky="nsew", padx=(4,0))
        headlight_frame.grid_columnconfigure(1, weight=1)
        headlight_frame.grid_columnconfigure(2, weight=1)
        headlight_title_frame = ttk.Frame(headlight_frame)
        headlight_title_frame.grid(row=0, column=0, columnspan=3, sticky='w')
        ttk.Label(headlight_title_frame, text="Headlight Result", style='Card.TLabelframe.Label').pack(side='left')
        self.disable_headlight_var = tk.BooleanVar()
        def toggle_headlight_fields():
            state = 'disabled' if self.disable_headlight_var.get() else 'normal'
            for w in [self.headlight_high_beam_intensity_left_entry, self.headlight_high_beam_intensity_right_entry, self.headlight_high_beam_horizontal_left_entry, self.headlight_high_beam_horizontal_right_entry, self.headlight_high_beam_vertical_left_entry, self.headlight_high_beam_vertical_right_entry, self.headlight_low_beam_horizontal_left_entry, self.headlight_low_beam_horizontal_right_entry, self.headlight_low_beam_vertical_left_entry, self.headlight_low_beam_vertical_right_entry]:
                w.config(state=state)
        ttk.Checkbutton(headlight_title_frame, text="Disable", variable=self.disable_headlight_var, command=toggle_headlight_fields).pack(side='left', padx=(8,0))
        headlight_rows = [
            ("High Beam Intensity", "high_beam_intensity"),
            ("High Beam Horizontal", "high_beam_horizontal"),
            ("High Beam Vertical", "high_beam_vertical"),
            ("Low Beam Intensity", "low_beam_intensity"),
            ("Low Beam Horizontal", "low_beam_horizontal"),
            ("Low Beam Vertical", "low_beam_vertical")
        ]
        ttk.Label(headlight_frame, text="", style='TLabel').grid(row=1, column=0)
        ttk.Label(headlight_frame, text="Left", style='TLabel').grid(row=1, column=1)
        ttk.Label(headlight_frame, text="Right", style='TLabel').grid(row=1, column=2)
        for r, (row_label, row_var) in enumerate(headlight_rows):
            ttk.Label(headlight_frame, text=row_label, style='TLabel').grid(row=r+2, column=0, sticky='w', padx=2, pady=1)
            setattr(self, f"headlight_{row_var}_left_entry", ttk.Entry(headlight_frame, width=6))
            getattr(self, f"headlight_{row_var}_left_entry").grid(row=r+2, column=1, sticky='ew', padx=1, pady=1)
            setattr(self, f"headlight_{row_var}_right_entry", ttk.Entry(headlight_frame, width=6))
            getattr(self, f"headlight_{row_var}_right_entry").grid(row=r+2, column=2, sticky='ew', padx=1, pady=1)
        for c in range(3):
            headlight_frame.grid_columnconfigure(c, weight=1)

        # --- BRAKE ---
        mid_center_frame = ttk.Frame(center_col)
        mid_center_frame.grid(row=1, column=0, sticky="nsew", pady=(0,8))
        mid_center_frame.grid_columnconfigure(0, weight=1)
        self.disable_brake_var = tk.BooleanVar()
        def toggle_brake_fields():
            state = 'disabled' if self.disable_brake_var.get() else 'normal'
            for w in self.brake_entries.values():
                w.config(state=state)
        brake_title_frame = ttk.Frame(mid_center_frame)
        brake_title_frame.grid(row=0, column=0, sticky='w', pady=(0,2))
        ttk.Label(brake_title_frame, text="Brake Result", style='Card.TLabelframe.Label').pack(side='left')
        ttk.Checkbutton(brake_title_frame, text="Disable", variable=self.disable_brake_var, command=toggle_brake_fields).pack(side='left', padx=(8,0))
        brake_frame = ttk.LabelFrame(mid_center_frame, style='Card.TLabelframe', padding=8, labelwidget=brake_title_frame, labelanchor='nw')
        brake_frame.grid(row=1, column=0, sticky='nsew')
        ttk.Label(brake_frame, text="", style='TLabel').grid(row=1, column=0, padx=1, sticky='w')
        ttk.Label(brake_frame, text="Front Axle", style='TLabel').grid(row=1, column=1, columnspan=3, sticky='ew')
        ttk.Label(brake_frame, text="Rear Axle", style='TLabel').grid(row=1, column=5, columnspan=3, sticky='ew')
        ttk.Label(brake_frame, text="Hand Brake", style='TLabel').grid(row=1, column=8, columnspan=4, sticky='ew')
        ttk.Label(brake_frame, text="", style='TLabel').grid(row=2, column=0, padx=1, sticky='w')
        # Update brake result frame headers to add a blank column after the third column
        for i, label in enumerate([" Left", " %", " Right", "   ", " Left", " %", " Right", " Left", " %", " Right", ]):
            ttk.Label(brake_frame, text=label, style='TLabel').grid(row=2, column=i+1, sticky='ew')
        # Update brake_row_labels to include new column logic
        brake_row_labels = [
            ("Roller Friction", "roller_friction", True),
            ("Brake Force", "brake_force", True),
            ("Max. Brake Force", "max_brake_force", True),
            ("Ovality", "ovality", "percent_only"),
            ("Axle Weight", "axle_weight", "second_and_fifth_col"),
            ("Axle Deceleration", "axle_deceleration", "second_and_fifth_col"),
            ("Total Weight", "total_weight", "mid_and_hand_percent"),
            ("Total Dec. Tot. Wt", "total_dec_tot_wt", "mid_and_hand_percent"),
            ("Total Brake Force", "total_brake_force", "mid_and_hand_percent"),
            ("Side Slip", "side_slip", False),
            ("Brake Difference (%)", "brake_difference", "percent_only"),
            ("Axle Load (kg)", "axle_load", "percent_only"),
            ("Sideslip (mm/m)", "sideslip", "percent_only"),
        ]
        # Remove specific rows from brake_row_labels
        brake_row_labels = [row for row in brake_row_labels if row[0] not in ["Brake Difference (%)", "Axle Load (kg)", "Sideslip (mm/m)"]]
        self.brake_entries = {}
        for r, (row_label, row_var, *special) in enumerate(brake_row_labels):
            ttk.Label(brake_frame, text=f"{row_label}", style='TLabel', foreground='blue').grid(row=r+3, column=0, sticky='w', padx=1, pady=1)
            if r == 0:
                for c in range(1, 11):
                    if c not in [2, 9]:
                        entry_name = f"brake_roller_friction_col{c}_entry"
                        self.brake_entries[entry_name] = ttk.Entry(brake_frame, width=7)
                        self.brake_entries[entry_name].grid(row=r+3, column=c, padx=6, pady=1, sticky='ew')
                    else:
                        ttk.Label(brake_frame, text="", style='TLabel').grid(row=r+3, column=c, padx=6, pady=1, sticky='ew')
            elif r == 3:
                for c in range(1, 11):
                    if c == 4:
                        entry_name = f"brake_{row_var}_col{c}_entry"
                        self.brake_entries[entry_name] = ttk.Entry(brake_frame, width=7)
                        self.brake_entries[entry_name].grid(row=r+3, column=c, padx=6, pady=1, sticky='ew')
                    else:
                        ttk.Label(brake_frame, text="", style='TLabel').grid(row=r+3, column=c, padx=6, pady=1, sticky='ew')
            elif r in [4, 5, 9]:
                for c in range(1, 11):
                    if c in [2, 6]:
                        ttk.Label(brake_frame, text="", style='TLabel').grid(row=r+3, column=c, padx=6, pady=1, sticky='ew')
                    else:
                        entry_name = f"brake_{row_var}_col{c}_entry"
                        self.brake_entries[entry_name] = ttk.Entry(brake_frame, width=7)
                        self.brake_entries[entry_name].grid(row=r+3, column=c, padx=6, pady=1, sticky='ew')
            elif r in [6, 7, 8]:
                for c in range(1, 11):
                    if c in [4, 9]:
                        entry_name = f"brake_{row_var}_col{c}_entry"
                        self.brake_entries[entry_name] = ttk.Entry(brake_frame, width=7)
                        self.brake_entries[entry_name].grid(row=r+3, column=c, padx=6, pady=1, sticky='ew')
                    else:
                        ttk.Label(brake_frame, text="", style='TLabel').grid(row=r+3, column=c, padx=6, pady=1, sticky='ew')
            else:
                for c in range(1, 11):
                    entry_name = f"brake_{row_var}_col{c}_entry"
                    self.brake_entries[entry_name] = ttk.Entry(brake_frame, width=7)
                    self.brake_entries[entry_name].grid(row=r+3, column=c, padx=6, pady=1, sticky='ew')
                    
                    # Add event handler to sync brake total weight with vehicle weight
                    if entry_name == 'brake_total_weight_col1_entry':
                        def sync_brake_weight_to_vehicle(*args):
                            if hasattr(self, 'weight_entry'):
                                brake_weight = self.brake_entries[entry_name].get()
                                self.weight_entry.delete(0, tk.END)
                                self.weight_entry.insert(0, brake_weight)
                        
                        self.brake_entries[entry_name].bind('<KeyRelease>', sync_brake_weight_to_vehicle)
                        self.brake_entries[entry_name].bind('<FocusOut>', sync_brake_weight_to_vehicle)
        brake_frame.grid_rowconfigure(len(brake_row_labels)+3, weight=1)

        # --- SHOCK/SUSPENSION ---
        photo_frame = ttk.LabelFrame(right_col, text="Vehicle Photo", style='Card.TLabelframe', padding=8)
        photo_frame.grid(row=0, column=0, sticky="nsew")
        photo_frame.grid_columnconfigure(0, weight=1)
        self.photo1_label = ttk.Label(photo_frame, text="[Photo 1]", style='TLabel', anchor='center')
        self.photo1_label.grid(row=0, column=0, padx=5, pady=5, sticky='ew')
        # Change button to locate image directory
        ttk.Button(photo_frame, text="Locate Image Directory", command=self.locate_image_directory).grid(row=1, column=0, pady=2, sticky='ew')
        self.disable_shock_var = tk.BooleanVar()
        def toggle_shock_fields():
            state = 'disabled' if self.disable_shock_var.get() else 'normal'
            for w in self.shock_entries.values():
                w.config(state=state)
        shock_frame = ttk.LabelFrame(photo_frame, text="Shock/Suspension Result", style='Card.TLabelframe', padding=4, labelanchor='nw')
        shock_frame.grid(row=2, column=0, sticky="nsew", pady=(8,0))
        shock_cols = ("", "Front Axle Left", "%", "Front Axle Right", "%", "Rear Axle Left", "%", "Rear Axle Right")
        shock_rows = [
            ("Amplitude", "amplitude"),
            ("Effect", "effect"),
            ("Wheel Weights", "wheel_weights")
        ]
        self.shock_entries = {}
        for c, col in enumerate(shock_cols):
            ttk.Label(shock_frame, text=col, style='TLabel').grid(row=1, column=c, padx=1, sticky='ew')
        for r, (row_label, row_var) in enumerate(shock_rows):
            ttk.Label(shock_frame, text=row_label, style='TLabel').grid(row=r+2, column=0, sticky='w', padx=1, pady=1)
            for alt in range(1, len(shock_cols)):
                entry_name = f"shock_{row_var}_col{alt}_entry"
                self.shock_entries[entry_name] = ttk.Entry(shock_frame, width=2)
                self.shock_entries[entry_name].grid(row=r+2, column=alt, padx=1, pady=1, sticky='ew')
        shock_frame.grid_rowconfigure(len(shock_rows)+2, weight=1)

        # --- SHOCK/SUSPENSION SECTION RANDOMIZATION ---
        # Amplitude (mm): 5-15 (columns 1,3,5,7)
        for c in [1, 3, 5, 7]:
            entry_name = f"shock_amplitude_col{c}_entry"
            if entry_name in self.shock_entries:
                self.shock_entries[entry_name].delete(0, tk.END)
                self.shock_entries[entry_name].insert(0, str(random.randint(5, 15)))
        # Effect (%): 41-100 (columns 2,4,6,8)
        for c in [2, 4, 6, 8]:
            entry_name = f"shock_effect_col{c}_entry"
            if entry_name in self.shock_entries:
                self.shock_entries[entry_name].delete(0, tk.END)
                self.shock_entries[entry_name].insert(0, str(random.randint(41, 100)))
        # Wheel Weights (kg):
        # Front Axle (columns 1,3): 600-1200
        for c in [1, 3]:
            entry_name = f"shock_wheel_weights_col{c}_entry"
            if entry_name in self.shock_entries:
                self.shock_entries[entry_name].delete(0, tk.END)
                self.shock_entries[entry_name].insert(0, str(random.randint(600, 1200)))
        # Rear Axle (columns 5,7): 400-1000
        for c in [5, 7]:
            entry_name = f"shock_wheel_weights_col{c}_entry"
            if entry_name in self.shock_entries:
                self.shock_entries[entry_name].delete(0, tk.END)
                self.shock_entries[entry_name].insert(0, str(random.randint(400, 1000)))

        # --- Visually distinct Printing Section at the bottom ---
        # --- Remove the bottom printing section ---
        # --- Add Print button under Vehicle Photo section ---
        def print_full_inspection():
            import json, os
            from PIL import Image
            # Read last inspection data
            try:
                with open(os.path.join(os.path.dirname(__file__), "last_inspection.json"), "r", encoding="utf-8") as f:
                    data = json.load(f)
            except Exception as e:
                messagebox.showerror("Print Error", f"No saved inspection data to print: {e}")
                return
            # Prepare print content (major info + results)
            lines = []
            lines.append("      MM AUTO INSPECTION      ")
            lines.append("Around Ayat Square +251-91252757 Addis Ababa - Ethiopia")
            lines.append("")
            lines.append(f"Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            lines.append(f"Plate Number: {data.get('PlateNo', '')}")
            lines.append(f"Chassis No: {data.get('ChassisNo', '')}")
            lines.append(f"Engine No: {data.get('EngineNo', '')}")
            lines.append(f"Model: {data.get('VehModel', '')}")
            lines.append(f"Made In: {data.get('VehicleMade', '')}")
            lines.append(f"Year: {data.get('MADEYEAR', '')}")
            lines.append(f"Type: {data.get('CARTYPE', '')}")
            lines.append(f"Inspector: {data.get('Inspector', '')}")
            lines.append(f"Result: {data.get('Status', '')}")
            lines.append("")
            # Major Results Section
            lines.append("--- Major Inspection Results ---")
            lines.append(f"Vehicle Weight: {data.get('TotalVehicleweight', data.get('weight', ''))}")
            lines.append(f"Total Brake Efficiency: {data.get('TotalServiceBrakeEfficiency', '')}")
            lines.append(f"Parking Brake Efficiency: {data.get('TotalParkingBrakeEfficeny', '')}")
            lines.append("Headlight Results:")
            lines.append(f"  High Beam Intensity Left: {data.get('HeadLightHighBeamIntensityLeft', '')}")
            lines.append(f"  High Beam Intensity Right: {data.get('HeadLightHighBeamIntensityRight', '')}")
            lines.append(f"  Low Beam Intensity Left: {data.get('HeadLightLowBeamIntensityLeft', '')}")
            lines.append(f"  Low Beam Intensity Right: {data.get('HeadLightLowBeamIntensityRight', '')}")
            lines.append(f"  High Beam Horizontal Deviation Left: {data.get('HeadLightHighBeamHorizontalDeviationLeft', '')}")
            lines.append(f"  High Beam Horizontal Deviation Right: {data.get('HeadLightHighBeamHorizontalDeviationRight', '')}")
            lines.append(f"  High Beam Vertical Deviation Left: {data.get('HeadLightHighBeamVerticalDeviationLeft', '')}")
            lines.append(f"  High Beam Vertical Deviation Right: {data.get('HeadLightHighBeamVerticalDeviationRight', '')}")
            lines.append(f"  Low Beam Horizontal Deviation Left: {data.get('HeadLightLowBeamHorizontalDeviationLeft', '')}")
            lines.append(f"  Low Beam Horizontal Deviation Right: {data.get('HeadLightLowBeamHorizontalDeviationRight', '')}")
            lines.append(f"  Low Beam Vertical Deviation Left: {data.get('HeadLightLowBeamVerticalDeviationLeft', '')}")
            lines.append(f"  Low Beam Vertical Deviation Right: {data.get('HeadLightLowBeamVerticalDeviationRight', '')}")
            lines.append("Gas/Emission Results:")
            lines.append(f"  HC: {data.get('HC', '')}")
            lines.append(f"  CO: {data.get('CO', '')}")
            lines.append(f"  CO2: {data.get('CO2', '')}")
            lines.append(f"  O2: {data.get('O2', '')}")
            lines.append(f"  Lamda: {data.get('Lamda', '')}")
            lines.append(f"  Nox: {data.get('Nox', '')}")
            lines.append(f"  Opacimeter: {data.get('Opacimeter', '')}")
            lines.append("")
            try:
                import win32print, win32ui
                from PIL import ImageWin
                # Get printer info for A4 size
                printer_name = win32print.GetDefaultPrinter()
                hprinter = win32print.OpenPrinter(printer_name)
                hdc = win32ui.CreateDC()
                hdc.CreatePrinterDC(printer_name)
                # A4 size in pixels at 300 DPI: 2480 x 3508
                # Get printable area
                printable_area = hdc.GetDeviceCaps(8), hdc.GetDeviceCaps(10)  # HORZRES, VERTRES
                page_width, page_height = printable_area
                margin_x = int(page_width * 0.08)
                margin_y = int(page_height * 0.05)
                x = margin_x
                y = margin_y
                hdc.StartDoc("MM Auto Inspection Report")
                hdc.StartPage()
                # Large, bold font for logo/title
                font_logo = win32ui.CreateFont({
                    "name": "Arial Black",
                    "height": int(page_height * 0.045),
                    "weight": 900
                })
                hdc.SelectObject(font_logo)
                hdc.TextOut(int(page_width/2 - 300), y, "MM AUTO INSPECTION")
                y += int(page_height * 0.055)
                font_addr = win32ui.CreateFont({
                    "name": "Arial",
                    "height": int(page_height * 0.025),
                    "weight": 400
                })
                hdc.SelectObject(font_addr)
                hdc.TextOut(int(page_width/2 - 350), y, "Around Ayat Square +251-91252757 Addis Ababa - Ethiopia")
                y += int(page_height * 0.04)
                # Info font
                font_info = win32ui.CreateFont({
                    "name": "Arial",
                    "height": int(page_height * 0.028),
                    "weight": 600
                })
                hdc.SelectObject(font_info)
                for line in lines[3:13]:
                    hdc.TextOut(x, y, line)
                    y += int(page_height * 0.035)
                y += int(page_height * 0.01)
                # Section header font
                font_section = win32ui.CreateFont({
                    "name": "Arial",
                    "height": int(page_height * 0.03),
                    "weight": 700
                })
                hdc.SelectObject(font_section)
                hdc.TextOut(x, y, lines[13])
                y += int(page_height * 0.04)
                # Results font
                font_result = win32ui.CreateFont({
                    "name": "Arial",
                    "height": int(page_height * 0.027),
                    "weight": 400
                })
                hdc.SelectObject(font_result)
                for line in lines[14:]:
                    if line.strip() == "":
                        y += int(page_height * 0.02)
                    else:
                        hdc.TextOut(x, y, line)
                        y += int(page_height * 0.032)
                # Print photo large and right-aligned if available
                if data.get("photo_path") and os.path.exists(data["photo_path"]):
                    try:
                        img = Image.open(data["photo_path"])
                        # Fit image to about 40% of page width, keep aspect ratio
                        img_width = int(page_width * 0.4)
                        img_height = int(img_width * img.height / img.width)
                        img = img.resize((img_width, img_height))
                        dib = ImageWin.Dib(img)
                        img_x = page_width - img_width - margin_x
                        img_y = margin_y + int(page_height * 0.12)
                        dib.draw(hdc.GetHandleOutput(), (img_x, img_y, img_x+img_width, img_y+img_height))
                    except Exception as e:
                        hdc.TextOut(x, y, f"[Photo error: {e}]")
                        y += int(page_height * 0.03)
                hdc.EndPage()
                hdc.EndDoc()
                hdc.DeleteDC()
                win32print.ClosePrinter(hprinter)
            except Exception as e:
                messagebox.showerror("Print Error", f"Could not print: {e}")
        print_btn = ttk.Button(photo_frame, text="Print", command=print_full_inspection, width=24)
        print_btn.grid(row=3, column=0, sticky='ew', pady=(8,0))

    def create_report_tab(self):
        """Modern, responsive report tab (formerly dashboard)."""
        view_frame = ttk.Frame(self.notebook, style='Card.TFrame', padding=20)
        self.notebook.add(view_frame, text="Report")
        view_frame.grid_columnconfigure(0, weight=1)
        view_frame.grid_rowconfigure(1, weight=1)
        status_frame = ttk.Frame(view_frame, style='Card.TFrame')
        status_frame.grid(row=0, column=0, sticky='ew', pady=(0, 20))
        for i in range(4):
            status_frame.grid_columnconfigure(i, weight=1)
        table_card = ttk.Frame(view_frame, style='Card.TFrame', padding=20)
        table_card.grid(row=1, column=0, sticky='nsew')
        view_frame.rowconfigure(1, weight=1)
        view_frame.columnconfigure(0, weight=1)
        table_card.rowconfigure(0, weight=1)
        table_card.columnconfigure(0, weight=1)
        columns = ('ID', 'PlateNo', 'CustName', 'VehModel', 'VehicleMade', 'ChassisNo', 'EngineNo', 'Status', 'Inspector', 'InspectionDate')
        self.report_tree = ttk.Treeview(table_card, columns=columns, show='headings')
        column_widths = {'ID': 50, 'PlateNo': 100, 'CustName': 120, 'VehModel': 100, 'VehicleMade': 100, 'ChassisNo': 100, 'EngineNo': 100, 'Status': 80, 'Inspector': 100, 'InspectionDate': 120}
        for col in columns:
            self.report_tree.heading(col, text=col)
            self.report_tree.column(col, width=column_widths.get(col, 100), anchor='center', stretch=True)
        self.report_tree.grid(row=0, column=0, sticky='nsew')
        tree_scroll_y = ttk.Scrollbar(table_card, orient="vertical", command=self.report_tree.yview)
        tree_scroll_x = ttk.Scrollbar(table_card, orient="horizontal", command=self.report_tree.xview)
        self.report_tree.configure(yscrollcommand=tree_scroll_y.set, xscrollcommand=tree_scroll_x.set)
        tree_scroll_y.grid(row=0, column=1, sticky='ns')
        tree_scroll_x.grid(row=1, column=0, sticky='ew')
        # Add Refresh button
        def update_status_cards():
            if not hasattr(self, 'db_manager') or not self.db_manager.connection:
                for i, val_label in enumerate(status_value_labels):
                    val_label.config(text="0")
                return
            try:
                cursor = self.db_manager.connection.cursor()
                # Total inspections
                cursor.execute("SELECT COUNT(*) FROM RESULT")
                total = cursor.fetchone()[0]
                # Passed
                cursor.execute("SELECT COUNT(*) FROM RESULT WHERE Status = 'PASS'")
                passed = cursor.fetchone()[0]
                # Failed
                cursor.execute("SELECT COUNT(*) FROM RESULT WHERE Status = 'FAIL'")
                failed = cursor.fetchone()[0]
                # Conditional
                cursor.execute("SELECT COUNT(*) FROM RESULT WHERE Status = 'CONDITIONAL'")
                conditional = cursor.fetchone()[0]
                values = [total, passed, failed, conditional]
                for i, val_label in enumerate(status_value_labels):
                    val_label.config(text=str(values[i]))
            except Exception as e:
                for i, val_label in enumerate(status_value_labels):
                    val_label.config(text="0")

        def refresh_report_table():
            if not hasattr(self, 'db_manager') or not self.db_manager.connection:
                messagebox.showwarning("Warning", "No database connected.")
                return
            try:
                cursor = self.db_manager.connection.cursor()
                cursor.execute("SELECT ID, PlateNo, CustName, VehModel, VehicleMade, ChassisNo, EngineNo, Status, Inspector, InspectionDate FROM RESULT ORDER BY InspectionDate DESC")
                # Clear existing items
                for item in self.report_tree.get_children():
                    self.report_tree.delete(item)
                # Add new items
                for row in cursor.fetchall():
                    clean_row = []
                    for val in row:
                        if isinstance(val, (tuple, list)):
                            val = ', '.join(str(v) for v in val)
                        if val is None:
                            val = ""
                        val = str(val)
                        if val.startswith("'") and val.endswith("'"):
                            val = val[1:-1]
                        clean_row.append(val)
                    self.report_tree.insert('', 'end', values=clean_row)
                update_status_cards()
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load report data: {str(e)}")
        refresh_btn = ttk.Button(table_card, text="Refresh", command=refresh_report_table)
        refresh_btn.grid(row=2, column=0, sticky='ew', pady=(8,0))
        # Initial load
        refresh_report_table()

    def select_database(self):
        """Override file selection: always use the fixed database path."""
        from tkinter import filedialog
        db_path = filedialog.asksaveasfilename(
            title="Select or Create Database File",
            defaultextension=".accdb",
            filetypes=[
                ("Access Database", "*.accdb;*.mdb"),
                ("Access 2007+ (*.accdb)", "*.accdb"),
                ("Access 2003 and earlier (*.mdb)", "*.mdb"),
                ("All Files", "*.*")
            ]
        )
        if not db_path:
            messagebox.showinfo("Cancelled", "No database file selected.")
            return
        self.db_manager = DatabaseManager(db_path)
        if self.db_manager.connect():
            if self.db_manager.create_tables():
                self.db_status_var.set(f"Connected: {os.path.basename(db_path)}")
                self.change_db_btn.config(text="Change Database")
                messagebox.showinfo("Success", f"Connected to database: {db_path}")
                self.save_last_db_path(db_path=db_path, image_dir=self.image_dir)
            else:
                messagebox.showerror("Error", "Failed to create tables.")
        else:
            messagebox.showerror("Error", "Failed to connect to database.")
    
    def setup_database(self):
        """Initialize database connection"""
        # Remove disabling of tabs so tabs are always enabled
        pass
    
    def save_inspection(self):
        """Save complete inspection data to database, mapping only direct UI fields to DB columns and filling others with defaults."""
        if not self.db_manager or not self.db_manager.connection:
            return
        
        # DB schema: column name -> type
        db_columns = [
            ("ID","VARCHAR"),
            ("PlateNo", "TEXT"),
            ("CustName", "TEXT"),
            ("VehModel", "TEXT"),
            ("VehicleMade", "TEXT"),
            ("ChassisNo", "TEXT"),
            ("EngineNo", "TEXT"),
            ("Status", "TEXT"),
            ("Inspector", "TEXT"),
            ("FrontAxileSideSlip", "DOUBLE"),
            ("RearAxileSideSlip", "DOUBLE"),
            ("SuspenstionFrontLeftEfficieny", "DOUBLE"),
            ("SuspenstionFrontRightEfficieny", "DOUBLE"),
            ("SuspensionFrontRandLDifference", "DOUBLE"),
            ("FrontLeftWeight", "DOUBLE"),
            ("FrontRightWeight", "DOUBLE"),
            ("SuspensionRareLeftEfficency", "DOUBLE"),
            ("SuspenstionRareRightEfficency", "DOUBLE"),
            ("SuspensionrareRandLDifference", "DOUBLE"),
            ("RareLeftWeight", "DOUBLE"),
            ("RareRightWeight", "DOUBLE"),
            ("RollingFrictionFrontLeftAxle", "DOUBLE"),
            ("RollingFrictionFrontRightAxle", "DOUBLE"),
            ("RollingFrictionFrontDifference", "DOUBLE"),
            ("RollingFrictionRareLeftaxle", "DOUBLE"),
            ("RollingFrictionRareRightaxle", "DOUBLE"),
            ("RollingFrictionRareDifference", "DOUBLE"),
            ("OutOfRoundnessOrOvalityFrontLeft", "DOUBLE"),
            ("OutOfRoundnessOrOvalityFrontRight", "DOUBLE"),
            ("OutOfRoundnessOrOvalityRearLeft", "DOUBLE"),
            ("OutofRoundnessOrOvalityRearRight", "DOUBLE"),
            ("MaximumServiceBrakeForceFrontLeft", "DOUBLE"),
            ("MaximumServiceBrakeForceFrontRight", "DOUBLE"),
            ("MaximumServiceBrakeForceFrontDifference", "DOUBLE"),
            ("MaximumServiceBrakeForceRareLeft", "DOUBLE"),
            ("MaximumServiceBrakeForceRareRight", "DOUBLE"),
            ("MaximumServiceBrakeForceRareDifference", "DOUBLE"),
            ("ServiceBrakeForceFrontLeft", "DOUBLE"),
            ("ServiceBrakeForceFrontRight", "DOUBLE"),
            ("ServiceBrakeForceFrontDifference", "DOUBLE"),
            ("ServiceBrakeForceRareLeft", "DOUBLE"),
            ("ServiceBrakeForceRareRight", "DOUBLE"),
            ("ServiceBrakeForceRareDifference", "DOUBLE"),
            ("Front1ServiceBrakeEfficiency", "DOUBLE"),
            ("Front2ServiceBrakeEfficiency", "DOUBLE"),
            ("Rear1ServiceBrakeEfficiency", "DOUBLE"),
            ("Rear2ServiceBrakeEfficiency", "DOUBLE"),
            ("Rear3ServiceBrakeEfficiency", "DOUBLE"),
            ("Rear4ServiceBrakeEfficiency", "DOUBLE"),
            ("Rear5ServiceBrakeEfficiency", "DOUBLE"),
            ("Rear6ServiceBrakeEfficiency", "DOUBLE"),
            ("Rear7ServiceBrakeEfficiency", "DOUBLE"),
            ("Rear8ServiceBrakeEfficiency", "DOUBLE"),
            ("TotalServiceBrakeEfficiency", "DOUBLE"),
            ("FrontAxleWeight", "DOUBLE"),
            ("RearAxleWeight", "DOUBLE"),
            ("TotalVehicleweight", "DOUBLE"),
            ("ParkingBrakeLeftForce", "DOUBLE"),
            ("ParkingBrakeRightForce", "DOUBLE"),
            ("ParkingBrakeLeftRightdifference", "DOUBLE"),
            ("ParkingbrakeOutOfRoundnessLEft", "DOUBLE"),
            ("ParkingBrakeOutOfRoundnessRight", "DOUBLE"),
            ("TotalParkingBrakeEfficeny", "DOUBLE"),
            ("HeadLightHighBeamIntensityLeft", "TEXT"),
            ("HeadLightHighBeamIntensityRight", "TEXT"),
            ("HeadLightHighBeamHorizontalDeviationLeft", "TEXT"),
            ("HeadLightHighBeamHorizontalDeviationRight", "TEXT"),
            ("HeadLightHighBeamVerticalDeviationLeft", "TEXT"),
            ("HeadLightHighBeamVerticalDeviationRight", "TEXT"),
            ("HeadLightLowBeamIntensityLeft", "TEXT"),
            ("HeadLightLowBeamIntensityRight", "TEXT"),
            ("HeadLightLowBeamHorizontalDeviationLeft", "TEXT"),
            ("HeadLightLowBeamHorizontalDeviationRight", "TEXT"),
            ("HeadLightLowBeamVerticalDeviationLeft", "TEXT"),
            ("HeadLightLowBeamVerticalDeviationRight", "TEXT"),
            ("FogLightIntensityLeft", "DOUBLE"),
            ("FogLightIntensityRight", "DOUBLE"),
            ("FogLightVerticalDeviationLeft", "DOUBLE"),
            ("FogLightVerticalDeviationRight", "DOUBLE"),
            ("FogLightHorizontalDeviationLeft", "DOUBLE"),
            ("FogLightHorizontalDeviationRight", "DOUBLE"),
            ("HC", "DOUBLE"),
            ("CO", "DOUBLE"),
            ("CO2", "DOUBLE"),
            ("O2", "DOUBLE"),
            ("COCOrr", "DOUBLE"),
            ("Nox", "DOUBLE"),
            ("Lamda", "DOUBLE"),
            ("OilTemp", "DOUBLE"),
            ("RPM", "DOUBLE"),
            ("Opacimeter", "DOUBLE"),
            ("PHOTO", "TEXT"),
            ("FilePlace", "TEXT"),
            ("LIBRENO", "TEXT"),
            ("CARTYPE", "TEXT"),
            ("MADEYEAR", "INTEGER"),
            ("FUELTYPE", "TEXT"),
            ("FLAGG", "TEXT"),
            ("MACHINERESULT", "TEXT"),
            ("VISUALRESULT", "TEXT"),
            ("ISTRUCK", "YESNO"),
            ("LampHeight", "DOUBLE"),
            ("VEHICLETYPE", "TEXT"),
            ("PlateCodeN", "TEXT"),
            ("PlateCodeT", "TEXT"),
            ("RPTID", "TEXT"),
            ("InspectionDate", "DATETIME"),
            ("CreatedDate", "DATETIME"),
            ("LastModified", "DATETIME")
        ]
        # Map DB columns to UI fields
        ui_map = {
            "PlateNo": self.plate_no_entry,
            "CustName": self.cust_entry,
            "VehModel": self.model_entry,
            "VehicleMade": self.made_in_as_entry,
            "ChassisNo": self.chan_entry,
            "EngineNo": self.eng_entry,
            "Status": self.result_var,
            "Inspector": self.technician_entry,
            "FilePlace": self.file_place_entry,
            "LIBRENO": self.libre_entry,
            "CARTYPE": self.type_entry,
            "MADEYEAR": self.year_entry,
            "FUELTYPE": self.fuel_entry,
            "PlateCodeN": self.pxsx1_entry,
            "PlateCodeT": self.pxsx2_entry,
            "RPTID": self.plate_no_entry,  # RPTID assigned from plate number
            # Add mappings for technical fields (gas, headlight, brake, shock) as needed
        }
        # Add technical mappings (example for gas fields)
        ui_map.update({
            "HC": self.gas_hc_left_entry,
            "CO": self.gas_co_left_entry,
            "CO2": self.gas_co2_left_entry,
            "O2": self.gas_o2_left_entry,
            "Lamda": self.gas_lambda_left_entry,
            "Nox": self.gas_n_entry,
        })
        
        # Add headlight mappings
        ui_map.update({
            "HeadLightHighBeamIntensityLeft": self.headlight_high_beam_intensity_left_entry,
            "HeadLightHighBeamIntensityRight": self.headlight_high_beam_intensity_right_entry,
            "HeadLightHighBeamHorizontalDeviationLeft": self.headlight_high_beam_horizontal_left_entry,
            "HeadLightHighBeamHorizontalDeviationRight": self.headlight_high_beam_horizontal_right_entry,
            "HeadLightHighBeamVerticalDeviationLeft": self.headlight_high_beam_vertical_left_entry,
            "HeadLightHighBeamVerticalDeviationRight": self.headlight_high_beam_vertical_right_entry,
            "HeadLightLowBeamIntensityLeft": self.headlight_low_beam_intensity_left_entry,
            "HeadLightLowBeamIntensityRight": self.headlight_low_beam_intensity_right_entry,
            "HeadLightLowBeamHorizontalDeviationLeft": self.headlight_low_beam_horizontal_left_entry,
            "HeadLightLowBeamHorizontalDeviationRight": self.headlight_low_beam_horizontal_right_entry,
            "HeadLightLowBeamVerticalDeviationLeft": self.headlight_low_beam_vertical_left_entry,
            "HeadLightLowBeamVerticalDeviationRight": self.headlight_low_beam_vertical_right_entry,
        })
        # Add roller friction mappings
        ui_map.update({
            "RollingFrictionFrontLeftAxle": self.brake_entries.get("brake_roller_friction_col1_entry"),
            "RollingFrictionFrontRightAxle": self.brake_entries.get("brake_roller_friction_col2_entry"),
            "RollingFrictionFrontDifference": self.brake_entries.get("brake_roller_friction_col3_entry"),
            "RollingFrictionRareLeftaxle": self.brake_entries.get("brake_roller_friction_col4_entry"),
            "RollingFrictionRareRightaxle": self.brake_entries.get("brake_roller_friction_col5_entry"),
            "RollingFrictionRareDifference": self.brake_entries.get("brake_roller_friction_col6_entry"),
        })
        
        # Add axle weight mappings
        ui_map.update({
            "TotalVehicleweight": self.weight_entry,  # Main vehicle weight from vehicle info
        })
        
        # Parking brake fields are now generated automatically with random values
        # Add shock/suspension mappings
        ui_map.update({
           
            "SuspenstionFrontLeftEfficieny": self.shock_entries.get("shock_effect_col2_entry"),
            "SuspenstionFrontRightEfficieny": self.shock_entries.get("shock_effect_col4_entry"),
            "SuspensionFrontRandLDifference": None,  # will compute below
            "SuspensionRareLeftEfficency": self.shock_entries.get("shock_effect_col6_entry"),
            "SuspenstionRareRightEfficency": self.shock_entries.get("shock_effect_col8_entry"),
            "SuspensionrareRandLDifference": None,  # will compute below
        })
        # Prepare data for saving
        inspection_data = {}
        headlight_fields = [
            "HeadLightHighBeamIntensityLeft",
            "HeadLightHighBeamIntensityRight",
            "HeadLightHighBeamHorizontalDeviationLeft",
            "HeadLightHighBeamHorizontalDeviationRight",
            "HeadLightHighBeamVerticalDeviationLeft",
            "HeadLightHighBeamVerticalDeviationRight",
            "HeadLightLowBeamIntensityLeft",
            "HeadLightLowBeamIntensityRight",
            "HeadLightLowBeamHorizontalDeviationLeft",
            "HeadLightLowBeamHorizontalDeviationRight",
            "HeadLightLowBeamVerticalDeviationLeft",
            "HeadLightLowBeamVerticalDeviationRight",
        ]
        for col, col_type in db_columns:
            if col in ui_map:
                widget = ui_map[col]
                value = None
                if hasattr(widget, 'cget') and widget.cget('state') == 'disabled':
                    # Disabled: use default
                    if col_type in ("DOUBLE", "INTEGER"):
                        value = 0
                    else:
                        value = ""
                elif hasattr(widget, 'get'):
                    value = widget.get()
                    if value is None or (isinstance(value, str) and value.strip() == ""):
                        if col_type in ("DOUBLE", "INTEGER"):
                            value = 0
                        else:
                            value = ""
                else:
                    value = widget
                # Prepend '+' to all headlight fields if not already present
                if col in headlight_fields:
                    value = str(value)
                    if not value.startswith('+') and value != '':
                        value = '+' + value
                # Convert type
                if col_type in ("DOUBLE", "INTEGER"):
                    try:
                        value = float(value) if value else 0
                        if col_type == "INTEGER":
                            value = int(value)
                    except Exception:
                        value = 0
                elif col_type == "YESNO":
                    value = bool(value)
                elif col_type == "DATETIME":
                    try:
                        value = datetime.strptime(value, "%Y-%m-%d %H:%M:%S")
                    except Exception:
                        value = datetime.now()
                else:
                    value = str(value)
                inspection_data[col] = value
            else:
                # Fill with default
                if col_type.startswith("TEXT"):
                    inspection_data[col] = ""
                elif col_type in ("DOUBLE", "INTEGER"):
                    inspection_data[col] = 0
                elif col_type == "DATETIME":
                    inspection_data[col] = datetime.now()
                elif col_type == "YESNO":
                    inspection_data[col] = False
        # --- FORCE CUSTOM VALUES FOR THESE FIELDS ---
        import random
        try:
            effect2 = float(self.shock_entries.get("shock_effect_col2_entry").get())
        except Exception:
            effect2 = 0
        try:
            effect4 = float(self.shock_entries.get("shock_effect_col4_entry").get())
        except Exception:
            effect4 = 0
        try:
            effect6 = float(self.shock_entries.get("shock_effect_col6_entry").get())
        except Exception:
            effect6 = 0
        try:
            effect8 = float(self.shock_entries.get("shock_effect_col8_entry").get())
        except Exception:
            effect8 = 0
        try:
            side1 = str(self.shock_entries.get("brake_side_slip_col4_entry").get())
        except Exception:
            side1 = 3
        try:
            side2 = str(self.shock_entries.get("brake_side_slip_col7_entry").get())
        except Exception:
            side2 = -2
        inspection_data["FrontAxileSideSlip"] = side1
        inspection_data["RearAxileSideSlip"] = side2
        inspection_data["SuspensionFrontRandLDifference"] = effect2 - effect4
        inspection_data["SuspensionrareRandLDifference"] = min(abs(effect8 - effect6), 20)
        # Always set these DB columns to 0 (or False for ISTRUCK)
        always_zero_cols = [
            "OutOfRoundnessOrOvalityFrontLeft",
            "OutOfRoundnessOrOvalityFrontRight",
            "OutOfRoundnessOrOvalityRearLeft",
            "OutofRoundnessOrOvalityRearRight",
            "Front1ServiceBrakeEfficiency",
            "Front2ServiceBrakeEfficiency",
            "Rear1ServiceBrakeEfficiency",
            "Rear2ServiceBrakeEfficiency",
            "Rear3ServiceBrakeEfficiency",
            "Rear4ServiceBrakeEfficiency",
            "Rear5ServiceBrakeEfficiency",
            "Rear6ServiceBrakeEfficiency",
            "Rear7ServiceBrakeEfficiency",
            "Rear8ServiceBrakeEfficiency",
            # Parking brake fields are now handled by ui_map and calculations, so remove from always_zero_cols
            # Headlight fields are now handled by ui_map, so remove from always_zero_cols
            "FogLightIntensityLeft",
            "FogLightIntensityRight",
            "FogLightVerticalDeviationLeft",
            "FogLightVerticalDeviationRight",
            "FogLightHorizontalDeviationLeft",
            "FogLightHorizontalDeviationRight",
            "COCOrr",
            "Nox",
            "OilTemp",
            "RPM",
            "ISTRUCK",
            "LampHeight"
        ]
        for col in always_zero_cols:
            if col == "ISTRUCK":
                inspection_data[col] = False
            else:
                inspection_data[col] = 0
        # Get vehicle weight and assign FrontLeftWeight, FrontRightWeight, RareLeftWeight, RareRightWeight
        try:
            vehicle_weight = float(self.weight_entry.get())
        except Exception:
            vehicle_weight = 0
        inspection_data["FrontLeftWeight"] = vehicle_weight / 4
        inspection_data["FrontRightWeight"] = vehicle_weight / 4
        inspection_data["RareLeftWeight"] = vehicle_weight / 4
        inspection_data["RareRightWeight"] = vehicle_weight / 4
        inspection_data["SuspensionRareLeftEfficency"] = effect6
        inspection_data["SuspenstionRareRightEfficency"] = effect8
        inspection_data["SuspensionrareRandLDifference"] = abs(effect6 - effect8)
        try:
            max_front_left = float(self.brake_entries.get("brake_max_brake_force_col1_entry").get())
        except Exception:
            max_front_left = 0
        try:
            max_front_dif = float(self.brake_entries.get("brake_max_brake_force_col2_entry").get())
        except Exception:
            max_front_dif = 0
        try:
            max_front_right = float(self.brake_entries.get("brake_max_brake_force_col3_entry").get())
        except Exception:
            max_front_right = 0
        try:
            max_rear_left = float(self.brake_entries.get("brake_max_brake_force_col5_entry").get())
        except Exception:
            max_rear_left = 0
        try:
            max_rear_dif = float(self.brake_entries.get("brake_max_brake_force_col6_entry").get())
        except Exception:
            max_rear_dif = 0
        try:
            max_rear_right = float(self.brake_entries.get("brake_max_brake_force_col7_entry").get())
        except Exception:
            max_rear_right = 0
        inspection_data["MaximumServiceBrakeForceFrontLeft"] = max_front_left
        inspection_data["MaximumServiceBrakeForceFrontRight"] = max_front_right
        inspection_data["MaximumServiceBrakeForceFrontDifference"] = max_front_dif
        inspection_data["MaximumServiceBrakeForceRareLeft"] = max_rear_left
        inspection_data["MaximumServiceBrakeForceRareRight"] = max_rear_right
        inspection_data["MaximumServiceBrakeForceRareDifference"] = max_rear_dif
        try:
            svc_front_left = float(self.brake_entries.get("brake_brake_force_col1_entry").get())
        except Exception:
            svc_front_left = 0
        try:
            svc_front_dif = float(self.brake_entries.get("brake_brake_force_col2_entry").get())
        except Exception:
            svc_front_dif = 0
        try:
            svc_front_right = float(self.brake_entries.get("brake_brake_force_col3_entry").get())
        except Exception:
            svc_front_right = 0
        try:
            svc_rear_left = float(self.brake_entries.get("brake_brake_force_col5_entry").get())
        except Exception:
            svc_rear_left = 0
        try:
            svc_rear_right = float(self.brake_entries.get("brake_brake_force_col7_entry").get())
        except Exception:
            svc_rear_right = 0
        try:
            svc_rear_dif = float(self.brake_entries.get("brake_brake_force_col6_entry").get())
        except Exception:
            svc_rear_dif = 0
        inspection_data["ServiceBrakeForceFrontLeft"] = svc_front_left
        inspection_data["ServiceBrakeForceFrontRight"] = svc_front_right
        inspection_data["ServiceBrakeForceFrontDifference"] = svc_front_dif
        inspection_data["ServiceBrakeForceRareLeft"] = svc_rear_left
        inspection_data["ServiceBrakeForceRareRight"] = svc_rear_right
        inspection_data["ServiceBrakeForceRareDifference"] = svc_rear_dif
        
        # Generate random parking brake values within standard ranges
        import random
        
        # Parking Brake Force (N): Standard range 500-2000 N per side
        parking_left_force = random.randint(500, 2000)
        parking_right_force = random.randint(500, 2000)
        
        # Ensure the difference is within acceptable range (0-30%)
        max_diff = max(parking_left_force, parking_right_force) * 0.30
        if abs(parking_right_force - parking_left_force) > max_diff:
            # Adjust one side to keep difference within range
            if parking_right_force > parking_left_force:
                parking_right_force = parking_left_force + random.randint(0, int(max_diff))
            else:
                parking_left_force = parking_right_force + random.randint(0, int(max_diff))
        
        # Total Parking Brake Efficiency (%): Should be ≥50% for most vehicles
        total_parking_efficiency = random.randint(50, 100)
        
        # Assign values to database fields
        inspection_data["ParkingBrakeLeftForce"] = parking_left_force
        inspection_data["ParkingBrakeRightForce"] = parking_right_force
        inspection_data["ParkingBrakeLeftRightdifference"] = abs(parking_right_force - parking_left_force)
        inspection_data["TotalParkingBrakeEfficeny"] = total_parking_efficiency
        
        # Set parking brake out-of-roundness to 0 (not typically measured)
        inspection_data["ParkingbrakeOutOfRoundnessLEft"] = 0
        inspection_data["ParkingBrakeOutOfRoundnessRight"] = 0
        
        # Calculate axle weights from brake entries
        
        try:
            front_right_weight = float(self.brake_entries.get("brake_axle_weight_col4_entry").get())
        except Exception:
            front_right_weight = 0
        
        try:
            rear_right_weight = float(self.brake_entries.get("brake_axle_weight_col7_entry").get())
        except Exception:
            rear_right_weight = 0
        
        # Calculate total axle weights
        inspection_data["FrontAxleWeight"] = front_right_weight
        inspection_data["RearAxleWeight"] = rear_right_weight
        
        # Calculate and constrain brake force differences (positive, 15-30)
        import random
        front_diff = abs(max_front_right - max_front_left)
        if not (15 <= front_diff <= 30):
            front_diff = random.randint(15, 30)
        inspection_data["MaximumServiceBrakeForceFrontDifference"] = front_diff
        rear_diff = abs(max_rear_right - max_rear_left)
        if not (15 <= rear_diff <= 30):
            rear_diff = random.randint(15, 30)
        inspection_data["MaximumServiceBrakeForceRareDifference"] = rear_diff
        # Calculate TotalServiceBrakeEfficiency as the average of all effect values
        effect_values = []
        for c in [2, 4, 6, 8]:
            entry = self.shock_entries.get(f"shock_effect_col{c}_entry")
            try:
                value = float(entry.get())
                effect_values.append(value)
            except Exception:
                pass
        if effect_values:
            inspection_data["TotalServiceBrakeEfficiency"] = int(sum(effect_values) / len(effect_values))
        else:
            inspection_data["TotalServiceBrakeEfficiency"] = 0
        # Save to DB
        try:
            import random
            # FrontLeftWeight: 270-350
            inspection_data["FrontLeftWeight"] = front_left_weight = random.randint(270, 350)
            # FrontRightWeight: 250-325, always less than FrontLeftWeight
            inspection_data["FrontRightWeight"] = front_right_weight = random.randint(250, min(325, front_left_weight - 1))
            # SuspensionRareLeftEfficiency: 41-65
            inspection_data["SuspensionRareLeftEfficency"] = suspension_rear_left_eff = random.randint(41, 65)
            # SuspensionRearRightEfficiency: 41-65
            inspection_data["SuspenstionRareRightEfficency"] = suspension_rear_right_eff = random.randint(41, 65)
            # RearLeftWeight: 199-231
            inspection_data["RareLeftWeight"] = rear_left_weight = random.randint(199, 231)
            # RearRightWeight: 199-231, always less than RearLeftWeight
            inspection_data["RareRightWeight"] = rear_right_weight = random.randint(199, min(231, rear_left_weight - 1))
            # RollingFrictionFrontLeftAxle: 0-3
           
            inspection_data["ID"] = "MM"
            inspection_data["TotalServiceBrakeEfficiency"] = random.randint(70, 90)
            # FrontAxleWeight and RearAxleWeight: RearAxleWeight + w = FrontAxleWeight, w in 30-100, sum = TotalVehicleweight
            try:
                total_vehicle_weight = float(self.weight_entry.get())
            except Exception:
                total_vehicle_weight = 0
            w = random.randint(30, 100)
            rear_axle_weight = (total_vehicle_weight - w) / 2
            front_axle_weight = rear_axle_weight + w
            inspection_data["FrontAxleWeight"] = front_axle_weight
            inspection_data["RearAxleWeight"] = rear_axle_weight
            # ParkingBrakeLeftForce: 70-95
            inspection_data["ParkingBrakeLeftForce"] = parking_brake_left = random.randint(70, 95)
            # ParkingBrakeRightForce: 45-75, always less than left
            inspection_data["ParkingBrakeRightForce"] = parking_brake_right = random.randint(45, min(75, parking_brake_left - 1))
            # ParkingBrakeLeftRightDifference: 100 - (right/left)*100
            inspection_data["ParkingBrakeLeftRightdifference"] = int(100 - (parking_brake_right / parking_brake_left) * 100)
            # TotalParkingBrakeEfficiency: 30-45
            inspection_data["TotalParkingBrakeEfficeny"] = random.randint(30, 45)
            # SuspensionrareRandLDifference = abs(SuspenstionRareRightEfficency - SuspensionRareLeftEfficency)
            inspection_data["SuspensionrareRandLDifference"] = abs(inspection_data["SuspenstionRareRightEfficency"] - inspection_data["SuspensionRareLeftEfficency"])
            # SuspensionFrontRandLDifference = abs(SuspenstionFrontRightEfficieny - SuspenstionFrontLeftEfficieny)
            inspection_data["SuspensionFrontRandLDifference"] = abs(inspection_data["SuspenstionFrontRightEfficieny"] - inspection_data["SuspenstionFrontLeftEfficieny"])
            # ServiceBrakeForceRareDifference = abs(ServiceBrakeForceRareRight - ServiceBrakeForceRareLeft)
            inspection_data["ServiceBrakeForceRareDifference"] = abs(inspection_data["ServiceBrakeForceRareRight"] - inspection_data["ServiceBrakeForceRareLeft"])
            # Ensure FrontAxleWeight and RearAxleWeight are integers
            
            # Add '+' prefix to headlight deviation fields if not already present
            for field in [
                "HeadLightHighBeamHorizontalDeviationLeft",
                "HeadLightHighBeamHorizontalDeviationRight",
                "HeadLightHighBeamVerticalDeviationLeft",
                "HeadLightHighBeamVerticalDeviationRight",
                "HeadLightLowBeamHorizontalDeviationLeft",
                "HeadLightLowBeamHorizontalDeviationRight",
                "HeadLightLowBeamVerticalDeviationLeft",
                "HeadLightLowBeamVerticalDeviationRight"
            ]:
                val = str(inspection_data.get(field, ""))
                if not val.startswith("+"):
                    inspection_data[field] = f"+{val}"
            # Assign RPTID from plate number
            inspection_data['RPTID'] = self.plate_no_entry.get()
            # Set ID to 'MM' for all rows
            inspection_data['ID'] = 'MM'
            # Force headlight vertical deviation fields to '0'
            inspection_data['HeadLightLowBeamVerticalDeviationLeft'] = '0'
            inspection_data['HeadLightLowBeamVerticalDeviationRight'] = '0'
            # Remove '+' from intensity fields before saving
            for field in [
                'HeadLightLowBeamIntensityLeft',
                'HeadLightLowBeamIntensityRight',
                'HeadLightHighBeamIntensityLeft',
                'HeadLightHighBeamIntensityRight',
            ]:
                if field in inspection_data and isinstance(inspection_data[field], str):
                    inspection_data[field] = inspection_data[field].lstrip('+')
            # Ensure pass/fail result fields are uppercase
            if 'Status' in inspection_data:
                inspection_data['Status'] = str(inspection_data['Status']).upper()
            if 'MACHINERESULT' in inspection_data:
                inspection_data['MACHINERESULT'] = str(inspection_data['MACHINERESULT']).upper()
            if 'VISUALRESULT' in inspection_data:
                inspection_data['VISUALRESULT'] = str(inspection_data['VISUALRESULT']).upper()
            # Force MACHINERESULT and VISUALRESULT to 'PASS'
            inspection_data['MACHINERESULT'] = 'PASS'
            inspection_data['VISUALRESULT'] = 'PASS'
            # Set VEHICLETYPE to the index of the selected type + 1 (1-based)
            if hasattr(self, 'type_entry'):
                inspection_data['VEHICLETYPE'] = self.type_entry.current() + 1
            self.db_manager.save_inspection(inspection_data)
            messagebox.showinfo("Success", "Inspection record saved successfully!")
            # Save inspection data for printing
            try:
                temp_data = dict(inspection_data)
                # Add photo path if available
                plate_value = self.plate_no_entry.get().strip().replace(" ", "_")
                photo_path = None
                if plate_value:
                    for ext in [".jpg", ".jpeg", ".png"]:
                        possible = os.path.join(os.getcwd(), "vehicle_photos", f"{plate_value}{ext}")
                        if os.path.exists(possible):
                            photo_path = possible
                            break
                temp_data["photo_path"] = photo_path
                with open(os.path.join(os.path.dirname(__file__), "last_inspection.json"), "w", encoding="utf-8") as f:
                    json.dump(temp_data, f, default=str, ensure_ascii=False, indent=2)
            except Exception as e:
                messagebox.showwarning("Warning", f"Could not save inspection for printing: {e}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save inspection: {str(e)}")
    
    def clear_form(self):
        """Clear all form fields"""
        # All UI fields are now direct attributes
        self.plate_no_entry.delete(0, tk.END)
        self.chan_entry.delete(0, tk.END)
        self.fuel_entry.delete(0, tk.END)
        self.year_entry.delete(0, tk.END)
        self.operator_entry.delete(0, tk.END)
        self.weight_entry.delete(0, tk.END)
        self.cust_entry.delete(0, tk.END)
        self.eng_entry.delete(0, tk.END)
        self.model_entry.delete(0, tk.END)
        self.libre_entry.delete(0, tk.END)
        self.made_in_as_entry.delete(0, tk.END)
        self.type_entry.delete(0, tk.END)
        self.file_place_entry.delete(0, tk.END)
        self.pxsx1_entry.delete(0, tk.END)
        self.pxsx2_entry.delete(0, tk.END)

        # Reset inspection date to today
        self.technician_entry.insert(0, datetime.now().strftime('%Y-%m-%d'))
        
        self.result_var.set("PASS")
        self.defects_text.delete(1.0, tk.END)
        
        # Remove any reference to checklist_vars
    
    def refresh_inspections(self):
        """Refresh the inspections list"""
        if not hasattr(self, 'db_manager') or not self.db_manager.connection:
            return
        
        try:
            cursor = self.db_manager.connection.cursor()
            cursor.execute("""
                SELECT ID, reg_number, make, model, owner_name, 
                       inspection_date, overall_result, expiry_date
                FROM Inspections
                ORDER BY inspection_date DESC
            """)
            
            # Clear existing items
            for item in self.inspections_tree.get_children():
                self.inspections_tree.delete(item)
            
            # Add new items
            for row in cursor.fetchall():
                self.inspections_tree.insert('', 'end', values=row)
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to refresh inspections: {str(e)}")
    
    def delete_inspection(self):
        """Delete selected inspection"""
        selected = self.inspections_tree.selection()
        if not selected:
            messagebox.showwarning("Warning", "Please select an inspection to delete")
            return
        
        if messagebox.askyesno("Confirm Delete", "Are you sure you want to delete this inspection?"):
            try:
                item = self.inspections_tree.item(selected[0])
                inspection_id = item['values'][0]
                
                cursor = self.db_manager.connection.cursor()
                cursor.execute("DELETE FROM Inspections WHERE ID = ?", (inspection_id,))
                self.db_manager.connection.commit()
                
                self.refresh_inspections()
                messagebox.showinfo("Success", "Inspection deleted successfully!")
                
            except Exception as e:
                messagebox.showerror("Error", f"Failed to delete inspection: {str(e)}")
    
    def edit_inspection(self):
        """Edit selected inspection"""
        selected = self.inspections_tree.selection()
        if not selected:
            messagebox.showwarning("Warning", "Please select an inspection to edit")
            return
        
        # This would open an edit dialog - simplified for now
        messagebox.showinfo("Info", "Edit functionality would be implemented here")
    
    def search_inspections(self):
        """Search inspections based on criteria"""
        if not hasattr(self, 'db_manager') or not self.db_manager.connection:
            return
        
        search_term = self.search_entry.get().strip()
        if not search_term:
            messagebox.showwarning("Warning", "Please enter a search term")
            return
        
        try:
            cursor = self.db_manager.connection.cursor()
            search_field = self.search_type.get()
            
            sql = f"""
                SELECT  reg_number, make, model, owner_name, 
                       inspection_date, overall_result, expiry_date
                FROM Inspections
                WHERE {search_field} LIKE ?
                ORDER BY inspection_date DESC
            """
            
            cursor.execute(sql, (f'%{search_term}%',))
            
            # Clear existing items
            for item in self.search_tree.get_children():
                self.search_tree.delete(item)
            
            # Add search results
            for row in cursor.fetchall():
                self.search_tree.insert('', 'end', values=row)
            
        except Exception as e:
            messagebox.showerror("Error", f"Search failed: {str(e)}")
    
    def export_to_csv(self):
        """Export all inspections to CSV"""
        if not hasattr(self, 'db_manager') or not self.db_manager.connection:
            return
        
        file_path = filedialog.asksaveasfilename(
            title="Export to CSV",
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")]
        )
        
        if file_path:
            try:
                import csv
                cursor = self.db_manager.connection.cursor()
                cursor.execute("SELECT * FROM Inspections")
                
                with open(file_path, 'w', newline='', encoding='utf-8') as csvfile:
                    writer = csv.writer(csvfile)
                    
                    # Write headers
                    columns = [desc[0] for desc in cursor.description]
                    writer.writerow(columns)
                    
                    # Write data
                    writer.writerows(cursor.fetchall())
                
                messagebox.showinfo("Success", f"Data exported to {file_path}")
                
            except Exception as e:
                messagebox.showerror("Error", f"Export failed: {str(e)}")
    
    def generate_monthly_report(self):
        """Generate monthly inspection report"""
        messagebox.showinfo("Info", "Monthly report generation would be implemented here")
    
    def expiring_certificates_report(self):
        """Generate report of expiring certificates"""
        if not hasattr(self, 'db_manager') or not self.db_manager.connection:
            return
        
        try:
            cursor = self.db_manager.connection.cursor()
            cursor.execute("""
                SELECT reg_number, owner_name, expiry_date
                FROM Inspections
                WHERE expiry_date BETWEEN Date() AND DateAdd('m', 3, Date())
                ORDER BY expiry_date
            """)
            
            results = cursor.fetchall()
            if results:
                report = "Certificates expiring in the next 3 months:\n\n"
                for row in results:
                    report += f"Registration: {row[0]}, Owner: {row[1]}, Expires: {row[2]}\n"
                
                messagebox.showinfo("Expiring Certificates", report)
            else:
                messagebox.showinfo("Info", "No certificates expiring in the next 3 months")
                
        except Exception as e:
            messagebox.showerror("Error", f"Report generation failed: {str(e)}")
    
    def refresh_statistics(self):
        """Refresh statistics display"""
        if not hasattr(self, 'db_manager') or not self.db_manager.connection:
            return
        
        try:
            cursor = self.db_manager.connection.cursor()
            
            # Get various statistics
            stats = []
            
            # Total inspections
            cursor.execute("SELECT COUNT(*) FROM Inspections")
            total = cursor.fetchone()[0]
            stats.append(f"Total Inspections: {total}")
            
            # This year's inspections
            cursor.execute("SELECT COUNT(*) FROM Inspections WHERE Year(inspection_date) = Year(Date())")
            this_year = cursor.fetchone()[0]
            stats.append(f"This Year: {this_year}")
            
            # Pass/Fail statistics
            cursor.execute("SELECT overall_result, COUNT(*) FROM Inspections GROUP BY overall_result")
            results = cursor.fetchall()
            stats.append("\nResults Breakdown:")
            for result, count in results:
                stats.append(f"  {result}: {count}")
            
            # Popular makes
            cursor.execute("""
                SELECT make, COUNT(*) as count 
                FROM Inspections 
                WHERE make IS NOT NULL AND make <> ''
                GROUP BY make 
                ORDER BY count DESC 
                LIMIT 5
            """)
            makes = cursor.fetchall()
            stats.append("\nTop 5 Vehicle Makes:")
            for make, count in makes:
                stats.append(f"  {make}: {count}")
            
            # Display statistics
            self.stats_text.delete(1.0, tk.END)
            self.stats_text.insert(1.0, "\n".join(stats))
            
        except Exception as e:
            messagebox.showerror("Error", f"Statistics refresh failed: {str(e)}")

    def upload_vehicle_photo(self, which=1):
        """Open file dialog to upload a vehicle photo and display it as a larger thumbnail (x3). Save the image with the plate number as filename. After upload, insert the filename (without extension) into the plate number field."""
        import os
        file_path = filedialog.askopenfilename(
            title="Select Vehicle Photo",
            filetypes=[("Image Files", "*.png;*.jpg;*.jpeg;*.bmp;*.gif")]
        )
        if file_path:
            from PIL import Image, ImageTk
            img = Image.open(file_path)
            img.thumbnail((480, 360))  # x3 size
            img_tk = ImageTk.PhotoImage(img)
            # Insert the filename (without extension) into the plate number field
            plate_number_entry = self.plate_no_entry
            if plate_number_entry:
                filename = os.path.splitext(os.path.basename(file_path))[0]
                plate_number_entry.delete(0, 'end')
                plate_number_entry.insert(0, filename)
            # Save the image with the plate number as filename
            plate_value = plate_number_entry.get().strip().replace(" ", "_") if plate_number_entry else ""
            if plate_value:
                ext = os.path.splitext(file_path)[1]
                save_dir = os.path.join(os.getcwd(), "vehicle_photos")
                os.makedirs(save_dir, exist_ok=True)
                save_path = os.path.join(save_dir, f"{plate_value}{ext}")
                try:
                    img.save(save_path)
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to save image: {str(e)}")
            if which == 1:
                self.photo1_label.configure(image=img_tk, text="")
                self.photo1_label.image = img_tk
            else:
                self.photo2_label.configure(image=img_tk, text="")
                self.photo2_label.image = img_tk

    def locate_image_directory(self):
        from tkinter import filedialog
        dir_path = filedialog.askdirectory(title="Select Image Directory")
        if dir_path:
            self.image_dir = dir_path
            self.save_last_db_path(db_path=getattr(self, 'db_manager', None) and getattr(self.db_manager, 'db_path', None), image_dir=dir_path)
            self.load_plate_image()

    def load_plate_image(self):
        from PIL import Image, ImageTk
        import os
        plate = self.plate_no_entry.get().strip()
        if not plate or not self.image_dir:
            self.photo1_label.configure(image='', text='[Photo 1]')
            self.photo1_label.image = None
            return
        # Try all common image extensions
        for ext in ['.jpg', '.jpeg', '.png', '.bmp', '.gif']:
            img_path = os.path.join(self.image_dir, f"{plate}{ext}")
            if os.path.exists(img_path):
                try:
                    img = Image.open(img_path)
                    img.thumbnail((480, 360))
                    img_tk = ImageTk.PhotoImage(img)
                    self.photo1_label.configure(image=img_tk, text='')
                    self.photo1_label.image = img_tk
                    return
                except Exception:
                    continue
        # If not found, show placeholder
        self.photo1_label.configure(image='', text='[Photo 1]')
        self.photo1_label.image = None

    def save_last_db_path(self, db_path=None, image_dir=None):
        import json
        data = {}
        try:
            if os.path.exists("last_db.json"):
                with open("last_db.json", "r") as f:
                    data = json.load(f)
        except Exception:
            data = {}
        if db_path:
            data["db_path"] = db_path
        if image_dir:
            data["image_dir"] = image_dir
        try:
            with open("last_db.json", "w") as f:
                json.dump(data, f)
        except Exception as e:
            print(f"Failed to save last DB path: {e}")

    def load_last_db_path(self):
        import os
        import json
        if os.path.exists("last_db.json"):
            try:
                with open("last_db.json", "r") as f:
                    data = json.load(f)
                    return data.get("db_path"), data.get("image_dir")
            except Exception as e:
                print(f"Failed to load last DB path: {e}")
        return None, None

    def __init__(self, root):
        self.root = root
        self.root.title("MM Vehicle Inspection Management System")
        self.root.geometry("1200x800")
        self.root.configure(bg='#f0f0f0')
        self.image_dir = None  # Directory for vehicle images
        self.setup_styles()
        self.create_main_interface()
        # Try to auto-load last used DB and image dir
        last_db, last_image_dir = self.load_last_db_path()
        if last_db:
            self.db_manager = DatabaseManager(last_db)
            if self.db_manager.connect():
                self.db_status_var.set(f"Connected: {os.path.basename(last_db)}")
                self.change_db_btn.config(text="Change Database")
                self.db_manager.create_tables()
        if last_image_dir:
            self.image_dir = last_image_dir
        self.setup_database()

    def select_database(self):
        from tkinter import filedialog
        db_path = filedialog.asksaveasfilename(
            title="Select or Create Database File",
            defaultextension=".accdb",
            filetypes=[
                ("Access Database", "*.accdb;*.mdb"),
                ("Access 2007+ (*.accdb)", "*.accdb"),
                ("Access 2003 and earlier (*.mdb)", "*.mdb"),
                ("All Files", "*.*")
            ]
        )
        if not db_path:
            messagebox.showinfo("Cancelled", "No database file selected.")
            return
        self.db_manager = DatabaseManager(db_path)
        if self.db_manager.connect():
            if self.db_manager.create_tables():
                self.db_status_var.set(f"Connected: {os.path.basename(db_path)}")
                self.change_db_btn.config(text="Change Database")
                messagebox.showinfo("Success", f"Connected to database: {db_path}")
                self.save_last_db_path(db_path=db_path, image_dir=self.image_dir)
            else:
                messagebox.showerror("Error", "Failed to create tables.")
        else:
            messagebox.showerror("Error", "Failed to connect to database.")

def get_secret_path():
    appdata = os.environ.get('LOCALAPPDATA') or os.path.expanduser('~')
    secret_dir = os.path.join(appdata, 'MM_auto_inspection')
    os.makedirs(secret_dir, exist_ok=True)
    return os.path.join(secret_dir, '.vault_secret')

class LoginWindow:
    def __init__(self, parent, on_success):
        self.top = tk.Toplevel(parent)
        self.top.title("Login - MM Vehicle Inspection Management System")
        self.top.geometry("350x200")
        self.top.grab_set()
        self.on_success = on_success
        self.frame = ttk.Frame(self.top, padding=30)
        self.frame.pack(expand=True)
        ttk.Label(self.frame, text="Enter Password:", font=("Segoe UI", 12)).pack(pady=(0,10))
        self.password_var = tk.StringVar()
        self.password_entry = ttk.Entry(self.frame, textvariable=self.password_var, show="*", width=25)
        self.password_entry.pack(pady=(0,10))
        self.password_entry.bind('<Return>', lambda e: self.try_login())
        self.status_label = ttk.Label(self.frame, text="", foreground="red")
        self.status_label.pack()
        self.login_btn = ttk.Button(self.frame, text="Login", command=self.try_login)
        self.login_btn.pack(pady=(10,0))
        self.password_entry.focus_set()
        self.secret_file = get_secret_path()
        if not os.path.exists(self.secret_file):
            self.setup_password()
        else:
            self.mode = "login"

    def setup_password(self):
        self.mode = "setup"
        self.status_label.config(text="Set a new password for this app.")
        self.login_btn.config(text="Set Password")

    def try_login(self):
        pwd = self.password_var.get()
        if not pwd:
            self.status_label.config(text="Password required.")
            return
        if self.mode == "setup":
            hashed = hashlib.sha256(pwd.encode()).hexdigest()
            with open(self.secret_file, "w") as f:
                f.write(hashed)
            self.status_label.config(text="Password set. Please login.")
            self.mode = "login"
            self.password_var.set("")
            self.login_btn.config(text="Login")
        else:
            with open(self.secret_file, "r") as f:
                stored = f.read().strip()
            if hashlib.sha256(pwd.encode()).hexdigest() == stored:
                self.top.destroy()
                self.on_success()
            else:
                self.status_label.config(text="Incorrect password.")
                self.password_var.set("")

def main():
    root = tk.Tk()
    root.withdraw()
    def launch_app():
        root.deiconify()
        app = CarInspectionApp(root)
    LoginWindow(root, on_success=launch_app)
    root.mainloop()

# Only run main if this is the main module
if __name__ == "__main__":
    main()
