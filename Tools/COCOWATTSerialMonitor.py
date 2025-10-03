import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog, colorchooser, font
import serial
import serial.tools.list_ports
import threading
import time
import queue
import csv
from datetime import datetime
import re
import matplotlib
matplotlib.use("TkAgg")
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
from matplotlib.figure import Figure
from collections import deque
import math
from PIL import Image, ImageTk

class SerialTerminalApp:
    def __init__(self, root):
        self.root = root
        self.root.title("COCOWATT Serial Monitor")
        self.root.geometry("1100x750")
        self.root.configure(bg="white")
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        # Set window icon
        try:
            logo_path = "cocowatt_logo.png"
            original_logo = Image.open(logo_path)
            resized_logo = original_logo.resize((32, 32), Image.LANCZOS)
            self.icon_image = ImageTk.PhotoImage(resized_logo)
            self.root.iconphoto(True, self.icon_image)
        except Exception:
            pass
        # State
        self.serial_conn = None
        self.running = False
        self.data_queue = queue.Queue()
        self.display_data = []
        self._last_display_count = 0
        self.last_activity = time.time()
        self.auto_scroll = True
        self.buffer = ""
        self.current_theme = "light"
        self.auto_clear_on_connect = False
        # New features
        self.command_history = []
        self.history_index = -1
        self.session_log_file = None
        self._last_rx_count = 0
        # Theme colors
        self.theme_colors = {
            "light": {
                "bg": "white", "fg": "black", "accent": "#4CAF50", "entry_bg": "white",
                "output_bg": "white", "output_fg": "black", "header_bg": "white",
                "header_fg": "black", "tagline_fg": "#4CAF50", "status_disconnected": "red",
                "status_connected": "green", "button_bg": "lightgray", "button_fg": "black",
                "graph_bg": "white", "graph_fg": "black",
            },
            "dark": {
                "bg": "#0D1B2A", "fg": "white", "accent": "#4CAF50", "entry_bg": "#1B263B",
                "output_bg": "#1B263B", "output_fg": "white", "header_bg": "#0D1B2A",
                "header_fg": "white", "tagline_fg": "#4CAF50", "status_disconnected": "red",
                "status_connected": "green", "button_bg": "#4CAF50", "button_fg": "white",
                "graph_bg": "#0D1B2A", "graph_fg": "white",
            }
        }
        # Parser config
        self.parser_entries = []
        self.parser_labels = []
        self.parser_values = {}
        self.parser_history = {}
        self.max_history = 2000
        self.default_parsers = ["ADC Volt:", "ADC Volt2:"]
        self.current_font = ("Consolas", 10)
        self.parser_colors = ['#4CAF50', '#2196F3', '#FF9800', '#9C27B0', '#FF5722', '#00BCD4', '#8BC34A', '#E91E63']
        # Session log
        self.session_log_file = None
        self.session_log_writer = None
        self.tx_append_newline = True
        self.graph_split_mode = False  # <<< NEW: Split mode flag

        self.create_widgets()
        self.processor_thread = threading.Thread(target=self.process_queue, daemon=True)
        self.processor_thread.start()

    def create_widgets(self):
        main_container = tk.Frame(self.root, bg=self.theme_colors[self.current_theme]["bg"])
        main_container.pack(fill="both", expand=True, padx=10, pady=10)
        # Header
        header_frame = tk.Frame(main_container, bg=self.theme_colors[self.current_theme]["header_bg"])
        header_frame.pack(fill="x", pady=(0, 10))
        self.title_label = tk.Label(header_frame, text="COCOWATT Serial Monitor", font=("Segoe UI", 16, "bold"), 
                                   fg=self.theme_colors[self.current_theme]["header_fg"], bg=self.theme_colors[self.current_theme]["header_bg"])
        self.title_label.pack(pady=(0, 5))
        self.tagline_label = tk.Label(header_frame, text="We Innovate, Educate and Boost Tomorrow!", font=("Segoe UI", 12, "italic"), 
                                     fg=self.theme_colors[self.current_theme]["tagline_fg"], bg=self.theme_colors[self.current_theme]["header_bg"])
        self.tagline_label.pack(pady=(0, 10))
        status_frame = tk.Frame(header_frame, bg=self.theme_colors[self.current_theme]["header_bg"])
        status_frame.pack(fill="x", pady=(0, 5))
        self.status_label = tk.Label(status_frame, text="❌ DISCONNECTED", fg=self.theme_colors[self.current_theme]["status_disconnected"], 
                                    bg=self.theme_colors[self.current_theme]["header_bg"], font=("Segoe UI", 9, "bold"))
        self.status_label.pack(side=tk.LEFT, padx=10)
        self.connect_btn = tk.Button(status_frame, text="🔗 Connect", font=("Segoe UI", 9, "bold"), 
                                    bg=self.theme_colors[self.current_theme]["button_bg"], fg=self.theme_colors[self.current_theme]["button_fg"], 
                                    command=self.connect_serial, relief="raised", padx=5, pady=2)
        self.connect_btn.pack(side=tk.LEFT, padx=5)
        self.disconnect_btn = tk.Button(status_frame, text="🔌 Disconnect", font=("Segoe UI", 9, "bold"), 
                                       bg=self.theme_colors[self.current_theme]["button_bg"], fg=self.theme_colors[self.current_theme]["button_fg"], 
                                       command=self.disconnect_serial, state="disabled", relief="raised", padx=5, pady=2)
        self.disconnect_btn.pack(side=tk.LEFT, padx=5)
        self.import_csv_btn = tk.Button(status_frame, text="📥 Import CSV", font=("Segoe UI", 9, "bold"),
                                bg=self.theme_colors[self.current_theme]["button_bg"],
                                fg=self.theme_colors[self.current_theme]["button_fg"],
                                command=self.import_csv_data, relief="raised", padx=5, pady=2)
        self.import_csv_btn.pack(side=tk.LEFT, padx=5)
        # Serial Configuration
        config_frame = tk.Frame(main_container, bg=self.theme_colors[self.current_theme]["bg"])
        config_frame.pack(fill="x", pady=(0, 10))
        ttk.Label(config_frame, text="Port:").pack(side=tk.LEFT, padx=(10, 0))
        self.port_var = tk.StringVar()
        self.port_combo = ttk.Combobox(config_frame, textvariable=self.port_var, state="readonly", width=10)
        self.port_combo.pack(side=tk.LEFT, padx=(0, 10))
        self.refresh_ports()
        ttk.Label(config_frame, text="Baud Rate:").pack(side=tk.LEFT, padx=(10, 0))
        self.baud_var = tk.StringVar(value="115200")
        self.baud_combo = ttk.Combobox(config_frame, textvariable=self.baud_var, values=["9600", "19200", "38400", "57600", "115200", "230400", "460800", "921600"], state="readonly", width=10)
        self.baud_combo.pack(side=tk.LEFT, padx=(0, 10))
        ttk.Label(config_frame, text="Data Size:").pack(side=tk.LEFT, padx=(10, 0))
        self.bytesize_var = tk.StringVar(value="8")
        bytesize_combo = ttk.Combobox(config_frame, textvariable=self.bytesize_var, values=["5", "6", "7", "8"], state="readonly", width=5)
        bytesize_combo.pack(side=tk.LEFT, padx=(0, 10))
        ttk.Label(config_frame, text="Parity:").pack(side=tk.LEFT, padx=(10, 0))
        self.parity_var = tk.StringVar(value="None")
        parity_combo = ttk.Combobox(config_frame, textvariable=self.parity_var, values=["None", "Even", "Odd", "Mark", "Space"], state="readonly", width=8)
        parity_combo.pack(side=tk.LEFT, padx=(0, 10))
        ttk.Label(config_frame, text="Stop Bits:").pack(side=tk.LEFT, padx=(10, 0))
        self.stopbits_var = tk.StringVar(value="1")
        stopbits_combo = ttk.Combobox(config_frame, textvariable=self.stopbits_var, values=["1", "1.5", "2"], state="readonly", width=5)
        stopbits_combo.pack(side=tk.LEFT, padx=(0, 10))
        self.refresh_btn = ttk.Button(config_frame, text="🔄 Refresh Ports", command=self.refresh_ports)
        self.refresh_btn.pack(side=tk.LEFT, padx=(10, 0))
        # Notebook
        notebook = ttk.Notebook(main_container)
        notebook.pack(fill="both", expand=True)
        # Clear All button
        clear_btn = tk.Button(main_container, text="🧹 Clear All", font=("Segoe UI", 10, "bold"), 
                              bg=self.theme_colors[self.current_theme]["accent"], fg="white",
                              command=self.clear_all_data, relief="raised", padx=10, pady=5)
        clear_btn.pack(pady=(0, 10))
        # Tabs
        terminal_frame = ttk.Frame(notebook)
        notebook.add(terminal_frame, text="Terminal")
        self.build_terminal_tab(terminal_frame)
        parser_config_frame = ttk.Frame(notebook)
        notebook.add(parser_config_frame, text="Parser")
        self.build_parser_config_tab(parser_config_frame)
        graph_frame = ttk.Frame(notebook)
        notebook.add(graph_frame, text="📈 Live Graph")
        self.build_graph_tab(graph_frame)
        settings_frame = ttk.Frame(notebook)
        notebook.add(settings_frame, text="⚙️ Settings")
        self.build_settings_tab(settings_frame)
        # Keyboard shortcuts
        self.root.bind('<Control-l>', lambda e: self.clear_all_data())
        self.root.bind('<Control-L>', lambda e: self.clear_all_data())

    def build_terminal_tab(self, parent):
        main_frame = ttk.Frame(parent)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        left_frame = ttk.Frame(main_frame)
        left_frame.pack(side=tk.LEFT, fill="y", padx=(0, 10))
        right_frame = ttk.Frame(main_frame)
        right_frame.pack(side=tk.RIGHT, fill="both", expand=True)
        # Send area
        send_frame = ttk.LabelFrame(left_frame, text="📤 Send Data")
        send_frame.pack(fill="x", pady=(0, 10))
        self.send_entry = ttk.Entry(send_frame)
        self.send_entry.pack(fill="x", pady=(0, 5))
        self.send_entry.bind("<Return>", lambda e: self.send_data())
        self.send_entry.bind("<Up>", self.history_up)
        self.send_entry.bind("<Down>", self.history_down)
        self.send_btn = ttk.Button(send_frame, text="Send", command=self.send_data)
        self.send_btn.pack(fill="x")
        # Quick Commands
        cmd_frame = ttk.LabelFrame(left_frame, text="🚀 Quick Commands")
        cmd_frame.pack(fill="x", pady=(0, 10))
        for cmd in ["TEST", "AT", "PING", "HELLO"]:
            btn = tk.Button(cmd_frame, text=cmd, font=("Segoe UI", 10, "bold"),
                            bg=self.theme_colors[self.current_theme]["accent"], fg="white",
                            relief="raised", padx=10, pady=5,
                            command=lambda c=cmd: self.send_data(c))
            btn.pack(fill="x", pady=2)
        ttk.Separator(cmd_frame, orient='horizontal').pack(fill='x', pady=5)
        digits = ["1", "2", "3", "4", "5", "6", "7", "8", "9", "0"]
        digits_frame = ttk.Frame(cmd_frame)
        digits_frame.pack(fill="x")
        for i, digit in enumerate(digits):
            if i % 3 == 0:
                row_frame = ttk.Frame(digits_frame)
                row_frame.pack(fill="x", pady=2)
            btn = tk.Button(row_frame, text=digit, font=("Segoe UI", 10, "bold"),
                            bg=self.theme_colors[self.current_theme]["accent"], fg="white",
                            relief="raised", padx=8, pady=4,
                            command=lambda d=digit: self.send_data(d))
            btn.pack(side=tk.LEFT, expand=True, fill="x", padx=1)
        tools_frame = ttk.Frame(left_frame)
        tools_frame.pack(fill="x", pady=(10, 0))
        ttk.Button(tools_frame, text="Clear RX", command=self.clear_rx).pack(side=tk.LEFT, fill="x", expand=True, padx=(0, 5))
        ttk.Button(tools_frame, text="Clear TX", command=self.clear_tx).pack(side=tk.RIGHT, fill="x", expand=True, padx=(5, 0))
        # === DUAL PANE: RX and TX ===
        rx_frame = ttk.LabelFrame(right_frame, text="📡 Received Data (RX)")
        rx_frame.pack(fill="both", expand=True, pady=(0, 5))
        self.rx_text = scrolledtext.ScrolledText(rx_frame, wrap=tk.WORD, font=("Consolas", 10))
        self.rx_text.pack(fill="both", expand=True, padx=5, pady=5)
        self.rx_text.config(state="disabled")
        tx_frame = ttk.LabelFrame(right_frame, text="📤 Sent Commands (TX)")
        tx_frame.pack(fill="both", expand=True, pady=(5, 0))
        self.tx_text = scrolledtext.ScrolledText(tx_frame, wrap=tk.WORD, font=("Consolas", 10, "bold"))
        self.tx_text.pack(fill="both", expand=True, padx=5, pady=5)
        self.tx_text.config(state="disabled")
        self.output_text = self.rx_text

    def build_parser_config_tab(self, parent):
        info_label = ttk.Label(parent, text="Add parser patterns (e.g., 'ADC Volt:', 'Temp:')")
        info_label.pack(anchor="w", padx=10, pady=10)
        self.parser_container = ttk.Frame(parent)
        self.parser_container.pack(fill="both", expand=True, padx=10, pady=10)
        for pattern in self.default_parsers:
            self.add_parser_row(pattern)
        btn_frame = ttk.Frame(parent)
        btn_frame.pack(fill="x", padx=10, pady=5)
        ttk.Button(btn_frame, text="➕ Add Parser", command=self.add_parser_row).pack(side=tk.LEFT)
        ttk.Button(btn_frame, text="🗑️ Clear All", command=self.clear_parsers).pack(side=tk.RIGHT)
        output_frame = ttk.LabelFrame(parent, text="Parsed Values (Latest)")
        output_frame.pack(fill="x", padx=10, pady=10)
        self.parsed_output = scrolledtext.ScrolledText(output_frame, height=8, font=self.current_font)
        self.parsed_output.pack(fill="x", padx=5, pady=5)
        self.parsed_output.config(state="disabled")

    def add_parser_row(self, default_text=""):
        row = ttk.Frame(self.parser_container)
        row.pack(fill="x", pady=2)
        label = ttk.Label(row, text="Pattern:")
        label.pack(side=tk.LEFT)
        entry = ttk.Entry(row, width=30)
        entry.pack(side=tk.LEFT, padx=5)
        entry.insert(0, default_text)
        remove_btn = ttk.Button(row, text="❌", width=3, command=lambda: self.remove_parser_row(row, entry))
        remove_btn.pack(side=tk.RIGHT)
        self.parser_entries.append(entry)
        self.parser_labels.append(row)

    def remove_parser_row(self, row, entry):
        if len(self.parser_entries) <= 1:
            messagebox.showinfo("Info", "Keep at least one parser.")
            return
        row.destroy()
        if entry in self.parser_entries:
            self.parser_entries.remove(entry)
        self.update_parsers()

    def clear_parsers(self):
        if len(self.parser_entries) > 1:
            for row in self.parser_labels[1:]:
                row.destroy()
            self.parser_entries = [self.parser_entries[0]]
            self.parser_labels = [self.parser_labels[0]]
        self.parser_entries[0].delete(0, tk.END)
        self.update_parsers()

    def update_parsers(self):
        patterns = []
        for entry in self.parser_entries:
            text = entry.get().strip()
            if text and text not in patterns:
                patterns.append(text)
        self.current_patterns = patterns
        self.parser_history = {p: deque(maxlen=self.max_history) for p in patterns}
        self.parser_values = {p: None for p in patterns}

    def import_csv_data(self):
        filename = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        if not filename:
            return
        try:
            with open(filename, 'r', encoding='utf-8-sig') as f:
                reader = csv.DictReader(f)
                rows = list(reader)
            if not rows:
                messagebox.showinfo("Info", "CSV is empty")
                return
            patterns = [e.get().strip() for e in self.parser_entries if e.get().strip()]
            if not patterns:
                patterns = ["ADC Volt:", "ADC Volt2:"]
            self.parser_history = {p: deque(maxlen=self.max_history) for p in patterns}
            self.parser_values = {p: None for p in patterns}
            self.display_data.clear()
            self.rx_text.config(state="normal")
            self.rx_text.delete(1.0, tk.END)
            base_time = None
            for row in rows:
                data = row['Data']
                ts_str = row['Timestamp']
                try:
                    if '.' in ts_str:
                        dt = datetime.strptime(ts_str, "%H:%M:%S.%f")
                    else:
                        dt = datetime.strptime(ts_str, "%H:%M:%S")
                    current_ts_numeric = dt.timestamp()
                    if base_time is None:
                        base_time = current_ts_numeric
                except Exception:
                    continue
                entry = {'timestamp': ts_str, 'data': data}
                self.display_data.append(entry)
                parsed = self.parse_line_for_patterns(data, patterns)
                for pat, val in parsed.items():
                    if isinstance(val, (int, float)) and not math.isnan(val):
                        if pat in self.parser_history:
                            self.parser_history[pat].append((current_ts_numeric, val))
                self.rx_text.insert(tk.END, f"{ts_str} ← {data}\n", "received")
            self.rx_text.config(state="disabled")
            self.rx_text.see(tk.END)
            self.update_graph()
            self.parsed_output.config(state="normal")
            self.parsed_output.delete(1.0, tk.END)
            for pattern in patterns:
                val = self.parser_values.get(pattern, "—")
                self.parsed_output.insert(tk.END, f"{pattern} {val}\n")
            self.parsed_output.config(state="disabled")
            messagebox.showinfo("Success", f"Imported {len(rows)} records")
        except Exception as e:
            messagebox.showerror("Import Error", str(e))    

    def export_all_terminal_data(self):
        filename = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")],
            initialfile=f"full_terminal_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        )
        if not filename:
            return
        try:
            with open(filename, 'w', newline='', encoding='utf-8-sig') as f:
                writer = csv.writer(f)
                writer.writerow(['Timestamp', 'Data', 'Direction'])
                for entry in self.display_data:
                    clean = entry['data'].replace('\x00', '').replace('\r', '').replace('\n', ' ')
                    writer.writerow([entry['timestamp'], clean, 'Received'])
                tx_content = self.tx_text.get(1.0, tk.END).strip()
                if tx_content:
                    for line in tx_content.split('\n'):
                        if ' → ' in line:
                            parts = line.split(' → ', 1)
                            if len(parts) == 2:
                                ts_part = parts[0].strip()
                                cmd = parts[1].strip()
                                writer.writerow([ts_part, cmd, 'Sent'])
            messagebox.showinfo("Success", f"Full terminal data exported to:\n{filename}")
        except Exception as e:
            messagebox.showerror("Export Error", f"Failed to export:\n{str(e)}")

    def build_graph_tab(self, parent):
        bg = self.theme_colors[self.current_theme]["graph_bg"]
        fg = self.theme_colors[self.current_theme]["graph_fg"]
        control_frame = tk.Frame(parent, bg=bg)
        control_frame.pack(fill="x", padx=5, pady=(5, 5))
        ttk.Button(control_frame, text="📥 Import CSV", command=self.import_csv_data).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(control_frame, text="📤 Export All Data", command=self.export_all_terminal_data).pack(side=tk.LEFT, padx=5)
        self.graph_auto_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(control_frame, text="Auto-update", variable=self.graph_auto_var).pack(side=tk.LEFT, padx=5)
        # <<< NEW: SplitOptions Toggle >>>
        self.graph_split_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(control_frame, text="SplitOptions", variable=self.graph_split_var, command=self.toggle_split_mode).pack(side=tk.LEFT, padx=5)
        ttk.Button(control_frame, text="🔄 Refresh", command=self.update_graph).pack(side=tk.RIGHT)
        ttk.Button(control_frame, text="💾 Export Graph", command=self.export_graph).pack(side=tk.RIGHT, padx=5)

        graph_frame = tk.Frame(parent, bg=bg)
        graph_frame.pack(fill="both", expand=True, padx=5, pady=0)
        self.figure = Figure(figsize=(10, 6), dpi=100, facecolor=bg)
        self.canvas = FigureCanvasTkAgg(self.figure, graph_frame)
        toolbar_frame = tk.Frame(graph_frame, bg=bg)
        toolbar_frame.pack(side=tk.TOP, fill="x", pady=2)
        toolbar = NavigationToolbar2Tk(self.canvas, toolbar_frame)
        toolbar.config(bg=bg)
        toolbar._message_label.config(bg=bg, fg=fg)
        for child in toolbar.winfo_children():
            if isinstance(child, tk.Button):
                child.config(bg=bg, fg=fg)
            elif isinstance(child, tk.Label):
                child.config(bg=bg, fg=fg)
        toolbar.update()
        toolbar.pack(side=tk.TOP, anchor="center")
        self.canvas.get_tk_widget().pack(fill="both", expand=True, pady=(2, 0))

    def toggle_split_mode(self):
        self.graph_split_mode = self.graph_split_var.get()
        self.update_graph()

    def build_settings_tab(self, parent):
        settings_frame = ttk.Frame(parent)
        settings_frame.pack(fill="both", expand=True, padx=10, pady=10)
        theme_label = ttk.Label(settings_frame, text="🎨 Theme:")
        theme_label.pack(anchor="w", pady=(5, 0))
        theme_var = tk.StringVar(value=self.current_theme)
        theme_combo = ttk.Combobox(settings_frame, textvariable=theme_var, values=["light", "dark"], state="readonly")
        theme_combo.pack(fill="x", pady=(0, 10))
        theme_combo.bind("<<ComboboxSelected>>", lambda e: self.switch_theme(theme_var.get()))
        font_label = ttk.Label(settings_frame, text="🔤 Font:")
        font_label.pack(anchor="w", pady=(5, 0))
        font_var = tk.StringVar(value=self.current_font[0])
        font_combo = ttk.Combobox(settings_frame, textvariable=font_var, values=list(font.families()), state="readonly")
        font_combo.pack(fill="x", pady=(0, 5))
        font_size_var = tk.StringVar(value=str(self.current_font[1]))
        font_size_combo = ttk.Combobox(settings_frame, textvariable=font_size_var, values=[str(i) for i in range(8, 21)], state="readonly", width=5)
        font_size_combo.pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(settings_frame, text="Apply Font", command=lambda: self.change_font(font_var.get(), int(font_size_var.get()))).pack(side=tk.LEFT)
        accent_label = ttk.Label(settings_frame, text="💚 Accent Color:")
        accent_label.pack(anchor="w", pady=(10, 0))
        self.accent_color_var = tk.StringVar(value=self.theme_colors[self.current_theme]["accent"])
        accent_entry = ttk.Entry(settings_frame, textvariable=self.accent_color_var, width=10)
        accent_entry.pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(settings_frame, text="Choose", command=self.choose_accent_color).pack(side=tk.LEFT)
        auto_clear_label = ttk.Label(settings_frame, text="🔄 Auto-Clear on Connect:")
        auto_clear_label.pack(anchor="w", pady=(10, 0))
        self.auto_clear_var = tk.BooleanVar(value=self.auto_clear_on_connect)
        auto_clear_check = ttk.Checkbutton(settings_frame, text="Clear all data when connecting", variable=self.auto_clear_var)
        auto_clear_check.pack(anchor="w", pady=(0, 10))
        tx_newline_label = ttk.Label(settings_frame, text="📤 Append \\n to TX Data:")
        tx_newline_label.pack(anchor="w", pady=(10, 0))
        self.tx_newline_var = tk.BooleanVar(value=self.tx_append_newline)
        tx_newline_check = ttk.Checkbutton(
            settings_frame,
            text="Add \\n after \\r when sending",
            variable=self.tx_newline_var
        )
        tx_newline_check.pack(anchor="w", pady=(0, 10))
        self.session_log_var = tk.BooleanVar(value=True)
        session_log_check = ttk.Checkbutton(
            settings_frame,
            text="Enable Session Logging (to CSV)",
            variable=self.session_log_var,
            command=self.toggle_session_logging
        )
        session_log_check.pack(anchor="w", pady=(0, 10))
        ttk.Button(settings_frame, text="✅ Apply All Settings", command=self.apply_settings).pack(fill="x", pady=(20, 0))

    def toggle_session_logging(self):
        enabled = self.session_log_var.get()
        if enabled and (not self.session_log_file or self.session_log_file.closed):
            log_filename = f"session_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
            self.session_log_file = open(log_filename, 'w', newline='', encoding='utf-8-sig')
            self.session_log_writer = csv.writer(self.session_log_file)
            self.session_log_writer.writerow(['Timestamp', 'Data', 'Direction'])
        elif not enabled and self.session_log_file and not self.session_log_file.closed:
            self.session_log_file.close()
            self.session_log_file = None

    def choose_accent_color(self):
        color = colorchooser.askcolor(title="Choose Accent Color")[1]
        if color:
            self.accent_color_var.set(color)

    def change_font(self, family, size):
        self.current_font = (family, size)
        self.output_text.config(font=self.current_font)
        self.parsed_output.config(font=self.current_font)

    def apply_settings(self):
        self.current_theme = "light" if self.theme_colors[self.current_theme]["bg"] == "white" else "dark"
        self.apply_theme()
        self.output_text.config(font=self.current_font)
        self.parsed_output.config(font=self.current_font)
        accent_color = self.accent_color_var.get()
        self.tx_append_newline = self.tx_newline_var.get()
        if accent_color:
            self.theme_colors["light"]["accent"] = accent_color
            self.theme_colors["dark"]["accent"] = accent_color
            self.apply_theme()
        self.auto_clear_on_connect = self.auto_clear_var.get()

    def apply_theme(self):
        colors = self.theme_colors[self.current_theme]
        self.root.configure(bg=colors["bg"])
        for widget in self.root.winfo_children():
            if isinstance(widget, tk.Frame) and widget != self.root:
                widget.config(bg=colors["header_bg"])
                for child in widget.winfo_children():
                    if isinstance(child, tk.Label):
                        child.config(bg=colors["header_bg"], fg=colors["header_fg"])
                    elif isinstance(child, tk.Button):
                        child.config(bg=colors["button_bg"], fg=colors["button_fg"])
        self.status_label.config(fg=colors["status_disconnected"] if not self.running else colors["status_connected"])
        self.connect_btn.config(bg=colors["button_bg"], fg=colors["button_fg"])
        self.disconnect_btn.config(bg=colors["button_bg"], fg=colors["button_fg"])
        self.output_text.config(bg=colors["output_bg"], fg=colors["output_fg"], insertbackground=colors["output_fg"])
        self.parsed_output.config(bg=colors["output_bg"], fg=colors["output_fg"], insertbackground=colors["output_fg"])
        if hasattr(self, 'figure'):
            self.figure.set_facecolor(colors["graph_bg"])
            if hasattr(self, 'canvas'):
                self.canvas.draw()

    def refresh_ports(self):
        ports = [p.device for p in serial.tools.list_ports.comports()]
        self.port_combo['values'] = ports
        if ports:
            self.port_var.set(ports[0])

    def get_parity(self, parity_str):
        mapping = {"None": serial.PARITY_NONE, "Even": serial.PARITY_EVEN, "Odd": serial.PARITY_ODD, "Mark": serial.PARITY_MARK, "Space": serial.PARITY_SPACE}
        return mapping.get(parity_str, serial.PARITY_NONE)

    def get_stopbits(self, stopbits_str):
        mapping = {"1": serial.STOPBITS_ONE, "1.5": serial.STOPBITS_ONE_POINT_FIVE, "2": serial.STOPBITS_TWO}
        return mapping.get(stopbits_str, serial.STOPBITS_ONE)

    def get_bytesize(self, size_str):
        mapping = {"5": serial.FIVEBITS, "6": serial.SIXBITS, "7": serial.SEVENBITS, "8": serial.EIGHTBITS}
        return mapping.get(size_str, serial.EIGHTBITS)

    def connect_serial(self):
        port = self.port_var.get()
        if not port:
            messagebox.showwarning("Warning", "Select a COM port")
            return
        try:
            baud = int(self.baud_var.get())
            bytesize = self.get_bytesize(self.bytesize_var.get())
            parity = self.get_parity(self.parity_var.get())
            stopbits = self.get_stopbits(self.stopbits_var.get())
        except Exception as e:
            messagebox.showerror("Error", f"Invalid setting: {e}")
            return
        try:
            if self.serial_conn and self.serial_conn.is_open:
                self.serial_conn.close()
            self.serial_conn = serial.Serial(
                port=port,
                baudrate=baud,
                timeout=0,
                bytesize=bytesize,
                parity=parity,
                stopbits=stopbits,
                rtscts=False,
                dsrdtr=False
            )
            self.serial_conn.reset_input_buffer()
            self.serial_conn.reset_output_buffer()
            if self.auto_clear_on_connect:
                self.clear_all_data(confirm=False)
            self.running = True
            self.last_activity = time.time()
            self.status_label.config(text="✅ CONNECTED", foreground=self.theme_colors[self.current_theme]["status_connected"])
            self.connect_btn.config(state="disabled")
            self.disconnect_btn.config(state="normal")
            self.reader_thread = threading.Thread(target=self.read_serial, daemon=True)
            self.reader_thread.start()
        except Exception as e:
            messagebox.showerror("Error", f"Connection failed:\n{e}")

    def disconnect_serial(self):
        self.running = False
        if self.serial_conn and self.serial_conn.is_open:
            self.serial_conn.close()
        self.serial_conn = None
        self.status_label.config(text="❌ DISCONNECTED", foreground=self.theme_colors[self.current_theme]["status_disconnected"])
        self.connect_btn.config(state="normal")
        self.disconnect_btn.config(state="disabled")

    def read_serial(self):
        last_data_time = time.time()
        while self.running:
            if self.serial_conn and self.serial_conn.is_open and self.serial_conn.in_waiting > 0:
                try:
                    data = self.serial_conn.read(self.serial_conn.in_waiting)
                    decoded = data.decode('iso-8859-1', errors='replace')
                    self.buffer += decoded
                    last_data_time = time.time()
                    self.buffer = self.buffer.replace('\r\n', '\n').replace('\r', '\n')
                    if '\n' in self.buffer:
                        lines = self.buffer.split('\n')
                        self.buffer = lines[-1]
                        for line in lines[:-1]:
                            if line.strip():
                                timestamp = datetime.now().strftime("%H:%M:%S.%f")[:-3]
                                self.data_queue.put({
                                    'timestamp': timestamp,
                                    'data': line.strip()
                                })
                except Exception as e:
                    if self.running:
                        self.data_queue.put({
                            'timestamp': datetime.now().strftime("%H:%M:%S"),
                            'data': f"[ERROR] {str(e)}"
                        })
            else:
                if self.buffer.strip() and (time.time() - last_data_time) > 0.1:
                    timestamp = datetime.now().strftime("%H:%M:%S.%f")[:-3]
                    self.data_queue.put({
                        'timestamp': timestamp,
                        'data': self.buffer.strip()
                    })
                    self.buffer = ""
                time.sleep(0.01)

    def send_data(self, text=None):
        if not self.running or not self.serial_conn or not self.serial_conn.is_open:
            messagebox.showwarning("Warning", "Not connected")
            return
        if text is None:
            text = self.send_entry.get().strip()
            if not text:
                return
        if self.tx_append_newline:
            data_to_send = text + '\r\n'
        else:
            data_to_send = text + '\r'
        try:
            encoded = data_to_send.encode('iso-8859-1')
            self.serial_conn.write(encoded)
            self.serial_conn.flush()
            if self.session_log_var.get() and self.session_log_writer:
                timestamp = datetime.now().strftime("%H:%M:%S.%f")[:-3]
                self.session_log_writer.writerow([timestamp, text, 'Sent'])
            timestamp = datetime.now().strftime("%H:%M:%S")
            self.tx_text.config(state="normal")
            self.tx_text.tag_config("tx_timestamp", foreground="#FF5722", font=("Consolas", 10, "bold"))
            self.tx_text.tag_config("tx_data", foreground="#03A9F4", font=("Consolas", 10))
            self.tx_text.insert(tk.END, f"{timestamp} ", "tx_timestamp")
            self.tx_text.insert(tk.END, f"→ {text}\n", "tx_data")
            self.tx_text.see(tk.END)
            self.tx_text.config(state="disabled")
            if text and (not self.command_history or self.command_history[-1] != text):
                self.command_history.append(text)
                self.history_index = len(self.command_history)
        except Exception as e:
            messagebox.showerror("Error", f"Send failed:\n{e}")

    def history_up(self, event=None):
        if self.command_history and self.history_index > 0:
            self.history_index -= 1
            self.send_entry.delete(0, tk.END)
            self.send_entry.insert(0, self.command_history[self.history_index])

    def history_down(self, event=None):
        if self.command_history and self.history_index < len(self.command_history) - 1:
            self.history_index += 1
            self.send_entry.delete(0, tk.END)
            self.send_entry.insert(0, self.command_history[self.history_index])
        else:
            self.history_index = len(self.command_history)
            self.send_entry.delete(0, tk.END)

    def parse_line_for_patterns(self, line, patterns):
        results = {}
        for pattern in patterns:
            if pattern in line:
                start_idx = line.find(pattern)
                after = line[start_idx + len(pattern):]
                match = re.search(r'[\d.]+', after)
                if match:
                    try:
                        value = float(match.group())
                        results[pattern] = value
                    except:
                        results[pattern] = match.group()
        return results

    def process_queue(self):
        while True:
            try:
                entry = self.data_queue.get(timeout=0.1)
                self.display_data.append(entry)
                if self.session_log_var.get() and self.session_log_writer:
                    self.session_log_writer.writerow([entry['timestamp'], entry['data'], 'Received'])
                if len(self.display_data) > 500:
                    self.display_data = self.display_data[-500:]
                self.update_display()
                self.update_parsers_and_graph()
            except:
                continue

    def update_parsers_and_graph(self):
        patterns = [e.get().strip() for e in self.parser_entries if e.get().strip()]
        if not patterns:
            patterns = ["ADC Volt:", "ADC Volt2:"]
        if self.display_data:
            latest_line = self.display_data[-1]['data']
            parsed = self.parse_line_for_patterns(latest_line, patterns)
            current_time = time.time()
            for pattern in patterns:
                value = parsed.get(pattern, None)
                self.parser_values[pattern] = value
                if value is not None and isinstance(value, (int, float)) and not math.isnan(value):
                    if pattern not in self.parser_history:
                        self.parser_history[pattern] = deque(maxlen=self.max_history)
                    self.parser_history[pattern].append((current_time, value))
            self.parsed_output.config(state="normal")
            self.parsed_output.delete(1.0, tk.END)
            for pattern in patterns:
                val = self.parser_values.get(pattern, "—")
                self.parsed_output.insert(tk.END, f"{pattern} {val}\n")
            self.parsed_output.config(state="disabled")
            if self.graph_auto_var.get():
                self.update_graph()

    def clear_rx(self):
        self.display_data = []
        self.buffer = ""
        self.rx_text.config(state="normal")
        self.rx_text.delete(1.0, tk.END)
        self.rx_text.config(state="disabled")
        self._last_rx_count = 0

    def clear_tx(self):
        self.tx_text.config(state="normal")
        self.tx_text.delete(1.0, tk.END)
        self.tx_text.config(state="disabled")

    def update_display(self):
        self.rx_text.config(state="normal")
        self.rx_text.tag_config("rx_timestamp", foreground="#2196F3", font=("Consolas", 10, "bold"))
        self.rx_text.tag_config("rx_data", foreground="#4CAF50", font=("Consolas", 10))
        if self.auto_scroll:
            start_idx = getattr(self, '_last_rx_count', 0)
            for i in range(start_idx, len(self.display_data)):
                entry = self.display_data[i]
                self.rx_text.insert(tk.END, f"{entry['timestamp']} ", "rx_timestamp")
                self.rx_text.insert(tk.END, f"← {entry['data']}\n", "rx_data")
            self._last_rx_count = len(self.display_data)
            self.rx_text.see(tk.END)
        else:
            self.rx_text.delete(1.0, tk.END)
            for entry in self.display_data:
                self.rx_text.insert(tk.END, f"{entry['timestamp']} ", "rx_timestamp")
                self.rx_text.insert(tk.END, f"← {entry['data']}\n", "rx_data")
            self.rx_text.see(tk.END)
        self.rx_text.config(state="disabled")

    def update_graph(self):
        bg = self.theme_colors[self.current_theme]["graph_bg"]
        fg = self.theme_colors[self.current_theme]["graph_fg"]
        self.figure.clear()

        patterns = list(self.parser_history.keys())
        valid_patterns = [p for p in patterns if self.parser_history[p]]

        if not valid_patterns:
            ax = self.figure.add_subplot(111, facecolor=bg)
            ax.text(0.5, 0.5, "No data to plot", transform=ax.transAxes, ha="center", color=fg, fontsize=12)
            ax.set_facecolor(bg)
            self.figure.set_facecolor(bg)
            self.canvas.draw()
            return

        all_times = []
        for hist in self.parser_history.values():
            if hist:
                all_times.extend([t for t, v in hist])
        if not all_times:
            ax = self.figure.add_subplot(111, facecolor=bg)
            ax.text(0.5, 0.5, "No numeric data", transform=ax.transAxes, ha="center", color=fg, fontsize=12)
            self.figure.set_facecolor(bg)
            self.canvas.draw()
            return

        t0 = min(all_times)

        if self.graph_split_mode:
            n = len(valid_patterns)
            # FIXED: Removed facecolor from subplots()
            axes = self.figure.subplots(n, 1, sharex=True)
            if n == 1:
                axes = [axes]
            self.figure.set_facecolor(bg)
            self.figure.subplots_adjust(hspace=0.3)

            for i, pattern in enumerate(valid_patterns):
                ax = axes[i]
                ax.set_facecolor(bg)  # Set facecolor per axis
                hist = self.parser_history[pattern]
                times = [(t - t0) for t, v in hist]
                values = [v for t, v in hist]
                color = self.parser_colors[i % len(self.parser_colors)]
                ax.plot(times, values, marker='o', linestyle='-', color=color, linewidth=1.5, label=pattern)
                ax.set_ylabel(pattern.strip(': '), color=fg, fontsize=9)
                ax.tick_params(colors=fg, labelsize=8)
                ax.grid(True, alpha=0.3, color=fg, linestyle='--')
                ax.legend(loc='upper right', facecolor=bg, edgecolor=fg, fontsize=8)

            axes[-1].set_xlabel("Time (s)", color=fg, fontsize=10)

        else:
            ax = self.figure.add_subplot(111, facecolor=bg)
            for i, pattern in enumerate(valid_patterns):
                hist = self.parser_history[pattern]
                times = [(t - t0) for t, v in hist]
                values = [v for t, v in hist]
                color = self.parser_colors[i % len(self.parser_colors)]
                ax.plot(times, values, marker='o', linestyle='-', label=pattern, linewidth=2, color=color)

            ax.set_title("Live Parsed Values", color=fg, fontsize=12)
            ax.set_xlabel("Time (s)", color=fg, fontsize=10)
            ax.set_ylabel("Value", color=fg, fontsize=10)
            ax.grid(True, alpha=0.4, color=fg, linestyle='--')
            ax.tick_params(colors=fg, labelsize=9)
            ax.legend(facecolor=bg, edgecolor=fg, fontsize=10, loc='upper right')

        self.figure.set_facecolor(bg)
        self.canvas.draw()

    def clear_display(self):
        self.display_data = []
        self.buffer = ""
        self.output_text.config(state="normal")
        self.output_text.delete(1.0, tk.END)
        self.output_text.config(state="disabled")

    def clear_all_data(self, confirm=True):
        if confirm and len(self.display_data) > 100:
            result = messagebox.askyesno(
                "Confirm Clear",
                f"You have {len(self.display_data)} messages.\nAre you sure you want to clear all data?",
                icon="warning"
            )
            if not result:
                return
        self.display_data = []
        self.buffer = ""
        self.output_text.config(state="normal")
        self.output_text.delete(1.0, tk.END)
        self.output_text.config(state="disabled")
        self.parsed_output.config(state="normal")
        self.parsed_output.delete(1.0, tk.END)
        self.parsed_output.config(state="disabled")
        self.parser_values = {p: None for p in self.parser_values}
        self.parser_history = {p: deque(maxlen=self.max_history) for p in self.parser_history}
        self.update_graph()
        if not confirm:
            return
        messagebox.showinfo("Cleared", "All data cleared successfully!")

    def export_data(self):
        if not self.display_data:
            messagebox.showinfo("Info", "No data to export")
            return
        filename = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")],
            initialfile=f"serial_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        )
        if not filename:
            return
        try:
            with open(filename, 'w', newline='', encoding='utf-8-sig') as f:
                writer = csv.writer(f)
                writer.writerow(['Timestamp', 'Data'])
                for entry in self.display_data:
                    clean_data = entry['data'].replace('\x00', '').replace('\r', '').replace('\n', ' ')
                    writer.writerow([entry['timestamp'], clean_data])
            messagebox.showinfo("Success", f"Exported to:\n{filename}")
        except Exception as e:
            messagebox.showerror("Export Error", f"Failed to export:\n{str(e)}")

    def export_graph(self):
        if not self.parser_history:
            messagebox.showinfo("Info", "No graph data to export")
            return
        filename = filedialog.asksaveasfilename(
            defaultextension=".png",
            filetypes=[("PNG files", "*.png")],
            initialfile=f"graph_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
        )
        if filename:
            try:
                self.figure.savefig(filename, dpi=150, bbox_inches='tight')
                messagebox.showinfo("Success", f"Graph exported to:\n{filename}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to export graph:\n{str(e)}")

    def switch_theme(self, theme):
        self.current_theme = theme
        self.apply_theme()

    def on_closing(self):
        self.running = False
        if self.serial_conn and self.serial_conn.is_open:
            self.serial_conn.close()
        if self.session_log_file and not self.session_log_file.closed:
            self.session_log_file.close()
        self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = SerialTerminalApp(root)
    root.mainloop()