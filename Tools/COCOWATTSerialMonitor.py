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
        self.root.configure(bg="white")  # Light theme by default
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

        # Set window icon (your logo)
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
        self.last_activity = time.time()
        self.auto_scroll = True
        self.buffer = ""
        self.current_theme = "light"  # Default: light

        # Theme colors
        self.theme_colors = {
            "light": {
                "bg": "white",
                "fg": "black",
                "accent": "#4CAF50",
                "entry_bg": "white",
                "output_bg": "white",
                "output_fg": "black",
                "header_bg": "white",
                "header_fg": "black",
                "tagline_fg": "#4CAF50",
                "status_disconnected": "red",
                "status_connected": "green",
                "button_bg": "lightgray",
                "button_fg": "black",
                "graph_bg": "white",
                "graph_fg": "black",
            },
            "dark": {
                "bg": "#0D1B2A",
                "fg": "white",
                "accent": "#4CAF50",
                "entry_bg": "#1B263B",
                "output_bg": "#1B263B",
                "output_fg": "white",
                "header_bg": "#0D1B2A",
                "header_fg": "white",
                "tagline_fg": "#4CAF50",
                "status_disconnected": "red",
                "status_connected": "green",
                "button_bg": "#4CAF50",
                "button_fg": "white",
                "graph_bg": "#0D1B2A",
                "graph_fg": "white",
            }
        }

        # Parser config
        self.parser_entries = []
        self.parser_labels = []
        self.parser_values = {}
        self.parser_history = {}
        self.max_history = 2000

        # Default parsers
        self.default_parsers = ["ADC Volt:", "ADC Volt2:"]

        # Font settings
        self.current_font = ("Consolas", 10)

        # Color palette for graph lines
        self.parser_colors = [
            '#4CAF50',  # Green (your brand)
            '#2196F3',  # Blue
            '#FF9800',  # Orange
            '#9C27B0',  # Purple
            '#FF5722',  # Deep Orange
            '#00BCD4',  # Cyan
            '#8BC34A',  # Light Green
            '#E91E63',  # Pink
        ]

        self.create_widgets()
        self.processor_thread = threading.Thread(target=self.process_queue, daemon=True)
        self.processor_thread.start()

    def create_widgets(self):
        main_container = tk.Frame(self.root, bg=self.theme_colors[self.current_theme]["bg"])
        main_container.pack(fill="both", expand=True, padx=10, pady=10)

        # Header
        header_frame = tk.Frame(main_container, bg=self.theme_colors[self.current_theme]["header_bg"])
        header_frame.pack(fill="x", pady=(0, 10))

        # Title
        self.title_label = tk.Label(header_frame, text="COCOWATT Serial Monitor", font=("Segoe UI", 16, "bold"), fg=self.theme_colors[self.current_theme]["header_fg"], bg=self.theme_colors[self.current_theme]["header_bg"])
        self.title_label.pack(pady=(0, 5))

        # Tagline
        self.tagline_label = tk.Label(header_frame, text="We Innovate, Educate and Boost Tomorrow!", font=("Segoe UI", 12, "italic"), fg=self.theme_colors[self.current_theme]["tagline_fg"], bg=self.theme_colors[self.current_theme]["header_bg"])
        self.tagline_label.pack(pady=(0, 10))

        # Status and buttons
        status_frame = tk.Frame(header_frame, bg=self.theme_colors[self.current_theme]["header_bg"])
        status_frame.pack(fill="x", pady=(0, 5))

        self.status_label = tk.Label(status_frame, text="❌ DISCONNECTED", fg=self.theme_colors[self.current_theme]["status_disconnected"], bg=self.theme_colors[self.current_theme]["header_bg"], font=("Segoe UI", 9, "bold"))
        self.status_label.pack(side=tk.LEFT, padx=10)

        self.connect_btn = tk.Button(status_frame, text="🔗 Connect", font=("Segoe UI", 9, "bold"), bg=self.theme_colors[self.current_theme]["button_bg"], fg=self.theme_colors[self.current_theme]["button_fg"], command=self.connect_serial, relief="raised", padx=5, pady=2)
        self.connect_btn.pack(side=tk.LEFT, padx=5)

        self.disconnect_btn = tk.Button(status_frame, text="🔌 Disconnect", font=("Segoe UI", 9, "bold"), bg=self.theme_colors[self.current_theme]["button_bg"], fg=self.theme_colors[self.current_theme]["button_fg"], command=self.disconnect_serial, state="disabled", relief="raised", padx=5, pady=2)
        self.disconnect_btn.pack(side=tk.LEFT, padx=5)

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
        self.baud_combo = ttk.Combobox(
            config_frame,
            textvariable=self.baud_var,
            values=["9600", "19200", "38400", "57600", "115200", "230400", "460800", "921600"],
            state="readonly",
            width=10
        )
        self.baud_combo.pack(side=tk.LEFT, padx=(0, 10))

        ttk.Label(config_frame, text="Data Size:").pack(side=tk.LEFT, padx=(10, 0))
        self.bytesize_var = tk.StringVar(value="8")
        bytesize_combo = ttk.Combobox(
            config_frame,
            textvariable=self.bytesize_var,
            values=["5", "6", "7", "8"],
            state="readonly",
            width=5
        )
        bytesize_combo.pack(side=tk.LEFT, padx=(0, 10))

        ttk.Label(config_frame, text="Parity:").pack(side=tk.LEFT, padx=(10, 0))
        self.parity_var = tk.StringVar(value="None")
        parity_combo = ttk.Combobox(
            config_frame,
            textvariable=self.parity_var,
            values=["None", "Even", "Odd", "Mark", "Space"],
            state="readonly",
            width=8
        )
        parity_combo.pack(side=tk.LEFT, padx=(0, 10))

        ttk.Label(config_frame, text="Stop Bits:").pack(side=tk.LEFT, padx=(10, 0))
        self.stopbits_var = tk.StringVar(value="1")
        stopbits_combo = ttk.Combobox(
            config_frame,
            textvariable=self.stopbits_var,
            values=["1", "1.5", "2"],
            state="readonly",
            width=5
        )
        stopbits_combo.pack(side=tk.LEFT, padx=(0, 10))

        self.refresh_btn = ttk.Button(config_frame, text="🔄 Refresh Ports", command=self.refresh_ports)
        self.refresh_btn.pack(side=tk.LEFT, padx=(10, 0))

        # Notebook
        notebook = ttk.Notebook(main_container)
        notebook.pack(fill="both", expand=True)

        # Terminal Tab
        terminal_frame = ttk.Frame(notebook)
        notebook.add(terminal_frame, text="Terminal")
        self.build_terminal_tab(terminal_frame)

        # Parser Tab
        parser_config_frame = ttk.Frame(notebook)
        notebook.add(parser_config_frame, text="Parser")
        self.build_parser_config_tab(parser_config_frame)

        # Graph Tab
        graph_frame = ttk.Frame(notebook)
        notebook.add(graph_frame, text="📈 Live Graph")
        self.build_graph_tab(graph_frame)

        # Settings Tab
        settings_frame = ttk.Frame(notebook)
        notebook.add(settings_frame, text="⚙️ Settings")
        self.build_settings_tab(settings_frame)

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
        self.send_btn = ttk.Button(send_frame, text="Send", command=self.send_data)
        self.send_btn.pack(fill="x")

        # Quick Commands
        cmd_frame = ttk.LabelFrame(left_frame, text="🚀 Quick Commands")
        cmd_frame.pack(fill="x", pady=(0, 10))
        for cmd in ["TEST", "AT", "PING", "HELLO"]:
            btn = tk.Button(cmd_frame, text=cmd, font=("Segoe UI", 10, "bold"), bg=self.theme_colors[self.current_theme]["accent"], fg="white", relief="raised", padx=10, pady=5, command=lambda c=cmd: self.send_data(c))
            btn.pack(fill="x", pady=2)

        # Clear & Export
        tools_frame = ttk.Frame(left_frame)
        tools_frame.pack(fill="x", pady=(10, 0))
        ttk.Button(tools_frame, text="Clear", command=self.clear_display).pack(side=tk.LEFT, fill="x", expand=True, padx=(0, 5))
        ttk.Button(tools_frame, text="Export CSV", command=self.export_data).pack(side=tk.RIGHT, fill="x", expand=True, padx=(5, 0))

        # Display
        display_frame = ttk.LabelFrame(right_frame, text="📡 Received Data")
        display_frame.pack(fill="both", expand=True)
        self.output_text = scrolledtext.ScrolledText(display_frame, wrap=tk.WORD, font=self.current_font)
        self.output_text.pack(fill="both", expand=True, padx=5, pady=5)
        self.output_text.config(state="disabled")

    def build_parser_config_tab(self, parent):
        info_label = ttk.Label(parent, text="Add parser patterns (e.g., 'ADC Volt:', 'Temp:')")
        info_label.pack(anchor="w", padx=10, pady=10)

        self.parser_container = ttk.Frame(parent)
        self.parser_container.pack(fill="both", expand=True, padx=10, pady=10)

        # Add default parsers
        for pattern in self.default_parsers:
            self.add_parser_row(pattern)

        # Add new parser button
        btn_frame = ttk.Frame(parent)
        btn_frame.pack(fill="x", padx=10, pady=5)
        ttk.Button(btn_frame, text="➕ Add Parser", command=self.add_parser_row).pack(side=tk.LEFT)
        ttk.Button(btn_frame, text="🗑️ Clear All", command=self.clear_parsers).pack(side=tk.RIGHT)

        # Parsed output display
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

        # Reset history
        self.parser_history = {p: deque(maxlen=self.max_history) for p in patterns}
        self.parser_values = {p: None for p in patterns}

    def build_graph_tab(self, parent):
        bg = self.theme_colors[self.current_theme]["graph_bg"]
        fg = self.theme_colors[self.current_theme]["graph_fg"]

        self.figure = Figure(figsize=(10, 6), dpi=100, facecolor=bg)
        self.ax = self.figure.add_subplot(111, facecolor=bg)
        
        self.ax.set_title("Live Parsed Values", color=fg, fontsize=12)
        self.ax.set_xlabel("Time (s)", color=fg, fontsize=10)
        self.ax.set_ylabel("Value", color=fg, fontsize=10)
        self.ax.grid(True, alpha=0.4, color=fg, linestyle='--')
        self.ax.tick_params(colors=fg, labelsize=9)

        self.canvas = FigureCanvasTkAgg(self.figure, parent)
        self.canvas.get_tk_widget().config(bg=bg)
        self.canvas.get_tk_widget().pack(fill="both", expand=True, padx=10, pady=10)

        toolbar = NavigationToolbar2Tk(self.canvas, parent)
        toolbar.config(bg=bg)
        toolbar._message_label.config(bg=bg, fg=fg)
        for child in toolbar.winfo_children():
            if isinstance(child, tk.Button):
                child.config(bg=bg, fg=fg, activebackground=bg, activeforeground=fg)
            elif isinstance(child, tk.Label):
                child.config(bg=bg, fg=fg)
        toolbar.update()

        control_frame = ttk.Frame(parent)
        control_frame.pack(fill="x", padx=10, pady=5)
        self.graph_auto_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(control_frame, text="Auto-update graph", variable=self.graph_auto_var).pack(side=tk.LEFT)
        ttk.Button(control_frame, text="Refresh Now", command=self.update_graph).pack(side=tk.RIGHT)

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

        ttk.Button(settings_frame, text="✅ Apply All Settings", command=self.apply_settings).pack(fill="x", pady=(20, 0))

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
        if accent_color:
            self.theme_colors["light"]["accent"] = accent_color
            self.theme_colors["dark"]["accent"] = accent_color
            self.apply_theme()

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

        if hasattr(self, 'ax'):
            self.ax.set_facecolor(colors["graph_bg"])
            self.ax.tick_params(colors=colors["graph_fg"])
            self.ax.spines['bottom'].set_color(colors["graph_fg"])
            self.ax.spines['top'].set_color(colors["graph_fg"])
            self.ax.spines['left'].set_color(colors["graph_fg"])
            self.ax.spines['right'].set_color(colors["graph_fg"])
            self.ax.set_title("Live Parsed Values", color=colors["graph_fg"])
            self.ax.set_xlabel("Time (s)", color=colors["graph_fg"])
            self.ax.set_ylabel("Value", color=colors["graph_fg"])
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
        while self.running:
            if self.serial_conn and self.serial_conn.is_open and self.serial_conn.in_waiting > 0:
                try:
                    data = self.serial_conn.read(self.serial_conn.in_waiting)
                    decoded = data.decode('iso-8859-1', errors='replace')
                    self.buffer += decoded

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
                time.sleep(0.01)

    def send_data(self, text=None):
        if not self.running or not self.serial_conn or not self.serial_conn.is_open:
            messagebox.showwarning("Warning", "Not connected")
            return

        if text is None:
            text = self.send_entry.get().strip()
            if not text:
                return

        try:
            data_to_send = text + '\r\n'
            encoded = data_to_send.encode('iso-8859-1')
            self.serial_conn.write(encoded)
            self.serial_conn.flush()
            self.send_entry.delete(0, tk.END)
        except Exception as e:
            messagebox.showerror("Error", f"Send failed:\n{e}")

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

    def update_display(self):
        self.output_text.config(state="normal")
        if self.auto_scroll:
            for entry in self.display_data[-200:]:
                self.output_text.insert(tk.END, f"{entry['timestamp']} → {entry['data']}\n")
            self.output_text.see(tk.END)
        else:
            if self.display_data:
                latest = self.display_data[-1]
                self.output_text.delete(1.0, tk.END)
                self.output_text.insert(tk.END, f"{latest['timestamp']} → {latest['data']}\n")
        self.output_text.config(state="disabled")

    def update_graph(self):
        bg = self.theme_colors[self.current_theme]["graph_bg"]
        fg = self.theme_colors[self.current_theme]["graph_fg"]

        self.ax.clear()
        self.ax.set_facecolor(bg)
        self.ax.set_title("Live Parsed Values", color=fg, fontsize=12)
        self.ax.set_xlabel("Time (s)", color=fg, fontsize=10)
        self.ax.set_ylabel("Value", color=fg, fontsize=10)
        self.ax.grid(True, alpha=0.4, color=fg, linestyle='--')
        self.ax.tick_params(colors=fg, labelsize=9)

        if not self.parser_history:
            self.ax.text(0.5, 0.5, "No data to plot", transform=self.ax.transAxes, ha="center", color=fg, fontsize=12)
            self.figure.set_facecolor(bg)
            self.canvas.draw()
            return

        all_times = []
        for hist in self.parser_history.values():
            if hist:
                all_times.extend([t for t, v in hist])
        if not all_times:
            self.ax.text(0.5, 0.5, "No numeric data", transform=self.ax.transAxes, ha="center", color=fg, fontsize=12)
            self.figure.set_facecolor(bg)
            self.canvas.draw()
            return

        t0 = min(all_times)
        patterns = list(self.parser_history.keys())
        for i, (pattern, hist) in enumerate(self.parser_history.items()):
            if hist:
                times = [(t - t0) for t, v in hist]
                values = [v for t, v in hist]
                color = self.parser_colors[i % len(self.parser_colors)]
                self.ax.plot(times, values, marker='o', linestyle='-', label=pattern, linewidth=2, color=color)

        self.ax.legend(facecolor=bg, edgecolor=fg, fontsize=10, loc='upper right')
        self.figure.set_facecolor(bg)
        self.canvas.draw()

    def toggle_scroll(self):
        self.auto_scroll = self.scroll_var.get()
        self.update_display()

    def clear_display(self):
        self.display_data = []
        self.buffer = ""
        self.output_text.config(state="normal")
        self.output_text.delete(1.0, tk.END)
        self.output_text.config(state="disabled")
        self.parsed_output.config(state="normal")
        self.parsed_output.delete(1.0, tk.END)
        self.parsed_output.config(state="disabled")

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

    def switch_theme(self, theme):
        self.current_theme = theme
        self.apply_theme()

    def on_closing(self):
        self.running = False
        if self.serial_conn and self.serial_conn.is_open:
            self.serial_conn.close()
        self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = SerialTerminalApp(root)
    root.mainloop()