import tkinter as tk
from tkinter import filedialog, messagebox, ttk  # ← ADD messagebox HERE
from tkcalendar import DateEntry
from PIL import Image, ImageTk
import sys
import os
import logging
from core.scheduler import generate_duty_chart

def resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

# ==================== GUI ====================
class DutyChartApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Duty Chart Generator")
        self.root.geometry("1000x700")
        self.root.minsize(800, 600)
        self.input_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.ratio_choice = tk.StringVar(value="1:3:6")
        self.theme_mode = tk.StringVar(value="dark")
        self.setup_styles(theme="dark")
        self.setup_widgets()
        self.root.bind("<Configure>", self.on_resize)

    def setup_styles(self, theme="dark"):
        self.style = ttk.Style()
        if theme == "dark":
            bg_color = "#181825"
            fg_color = "#f6f6f6"
            btn_bg = "#222222"
            btn_fg = "#222222"
            hover_bg = "#2d2d2d"
            entry_bg = "#232336"
            entry_fg = "#222222"
            summary_bg = "#232336"
            sum_fg = "#f6f6f6"
            header_fg = "#FFD700"  # Golden color for header in dark theme
        else:
            bg_color = "#f8fafc"
            fg_color = "#222222"
            btn_bg = "#e3e8f0"
            btn_fg = "#222222"
            hover_bg = "#ddebf9"
            entry_bg = "#fff"
            entry_fg = "#222222"
            summary_bg = "#f2f6fa"
            sum_fg = "#222222"
            header_fg = "#06a0c0"  # Blue color for header in light theme

        self.bg_color = bg_color
        self.fg_color = fg_color
        self.btn_bg = btn_bg
        self.btn_fg = btn_fg
        self.hover_bg = hover_bg
        self.entry_bg = entry_bg
        self.entry_fg = entry_fg
        self.summary_bg = summary_bg
        self.sum_fg = sum_fg
        self.header_fg = header_fg

        self.style.configure("TFrame", background=bg_color)
        self.style.configure("TLabel", font=("Segoe UI", 10), background=bg_color, foreground=fg_color)
        self.style.configure("Title.TLabel", font=("Segoe UI", 20, "bold"), foreground="#06a0c0", background=bg_color)
        self.style.configure("Header.TLabel", font=("Segoe UI", 14, "bold"), foreground=header_fg, background=bg_color)
        self.style.configure("TEntry", fieldbackground=entry_bg, background=entry_bg, foreground=entry_fg)
        self.style.configure("TButton", background=btn_bg, foreground=btn_fg, font=("Segoe UI", 11, "bold"))
        self.style.map("TButton",
            background=[("active", hover_bg), ("!active", btn_bg)],
            foreground=[("active", btn_fg), ("!active", btn_fg)]
        )
        self.style.configure("TRadiobutton", font=("Segoe UI", 10), background=bg_color, foreground=fg_color)
        self.style.configure("TProgressbar", thickness=20, background="#06a0c0", troughcolor=btn_bg, bordercolor=bg_color)
        self.style.configure("TLabelframe", font=("Segoe UI", 12, "bold"), foreground="#06a0c0", background=bg_color)
        self.style.configure("TLabelframe.Label", font=("Segoe UI", 12, "bold"), foreground="#06a0c0", background=bg_color)
        self.style.configure("Vertical.TScrollbar", background=bg_color, troughcolor=btn_bg, arrowcolor=fg_color)

    def toggle_theme(self):
        new_theme = "light" if self.theme_mode.get() == "dark" else "dark"
        self.theme_mode.set(new_theme)
        self.setup_styles(theme=new_theme)
        self.setup_widgets(reset=True)

    def setup_widgets(self, reset=False):
        if reset:
            for child in self.root.winfo_children():
                child.destroy()
        main_frame = ttk.Frame(self.root, padding=20, style="TFrame")
        main_frame.pack(fill="both", expand=True)
        header_frame = ttk.Frame(main_frame, style="TFrame")
        header_frame.pack(fill="x")



        def resource_path(relative_path):
            """ Get absolute path to resource, works for dev and PyInstaller """
            if hasattr(sys, '_MEIPASS'):
                return os.path.join(sys._MEIPASS, relative_path)
            return os.path.join(os.path.dirname(os.path.abspath(__file__)), relative_path)

        try:
            logo_path = resource_path("../assets/logo.png")
            print(f"Looking for logo at: {logo_path}")
            print(f"File exists? {os.path.exists(logo_path)}")
            logo_image = Image.open(logo_path)
            logo_image = logo_image.resize((50, 50), Image.Resampling.LANCZOS)
            self.logo = ImageTk.PhotoImage(logo_image)
            logo_label = ttk.Label(header_frame, image=self.logo, background=self.bg_color)
            logo_label.pack(side="left", padx=5, pady=(0, 10))
        except Exception as e:
            logging.error(f"Failed to load logo: {str(e)}")
            logo_label = ttk.Label(header_frame, text="[Logo]", style="Header.TLabel")
            logo_label.pack(side="left", padx=5, pady=(0, 10))


        # Header text
        ttk.Label(header_frame, text="GOVERNMENT COLLEGE OF ENGINEERING, SALEM - 636011", style="Header.TLabel", anchor="center").pack(side="left", fill="x", expand=True, pady=(0, 10))
        # Toggle theme button
        toggle_btn = ttk.Button(header_frame, text="Toggle Theme", command=self.toggle_theme, style="TButton")
        toggle_btn.pack(side="right", pady=(0, 10))
        # Duty Chart Generator heading
        ttk.Label(main_frame, text="Duty Chart Generator", style="Title.TLabel").pack(pady=10)
        input_frame = ttk.LabelFrame(main_frame, text="Input File", padding=10, style="TLabelframe")
        input_frame.pack(fill="x", pady=5)
        input_entry = ttk.Entry(input_frame, textvariable=self.input_path, width=50, style="TEntry")
        input_entry.pack(side="left", padx=5)
        input_entry.configure(foreground=self.entry_fg, background=self.entry_bg)
        input_button = ttk.Button(input_frame, text="Browse", command=self.browse_input, style="TButton")
        input_button.pack(side="left", padx=5)
        output_frame = ttk.LabelFrame(main_frame, text="Output File", padding=10, style="TLabelframe")
        output_frame.pack(fill="x", pady=5)
        output_entry = ttk.Entry(output_frame, textvariable=self.output_path, width=50, style="TEntry")
        output_entry.pack(side="left", padx=5)
        output_entry.configure(foreground=self.entry_fg, background=self.entry_bg)
        output_button = ttk.Button(output_frame, text="Save As", command=self.browse_output, style="TButton")
        output_button.pack(side="left", padx=5)
        date_frame = ttk.LabelFrame(main_frame, text="Date Ranges", padding=10, style="TLabelframe")
        date_frame.pack(fill="x", pady=5)
        slot1_frame = ttk.Frame(date_frame, style="TFrame")
        slot1_frame.pack(fill="x", pady=5)
        ttk.Label(slot1_frame, text="Slot 1 Start:", style="TLabel").pack(side="left")
        self.slot1_start = DateEntry(slot1_frame, date_pattern="dd/mm/yyyy")
        self.slot1_start.pack(side="left", padx=5)
        ttk.Label(slot1_frame, text="End:", style="TLabel").pack(side="left")
        self.slot1_end = DateEntry(slot1_frame, date_pattern="dd/mm/yyyy")
        self.slot1_end.pack(side="left", padx=5)
        slot2_frame = ttk.Frame(date_frame, style="TFrame")
        slot2_frame.pack(fill="x", pady=5)
        ttk.Label(slot2_frame, text="Slot 2 Start:", style="TLabel").pack(side="left")
        self.slot2_start = DateEntry(slot2_frame, date_pattern="dd/mm/yyyy")
        self.slot2_start.pack(side="left", padx=5)
        ttk.Label(slot2_frame, text="End:", style="TLabel").pack(side="left")
        self.slot2_end = DateEntry(slot2_frame, date_pattern="dd/mm/yyyy")
        self.slot2_end.pack(side="left", padx=5)
        ratio_frame = ttk.LabelFrame(main_frame, text="Duty Ratio (Prof:ASP:AP)", padding=10, style="TLabelframe")
        ratio_frame.pack(fill="x", pady=5)
        for val in ["1:3:6", "1:3:7", "1:4:8"]:
            radio = ttk.Radiobutton(ratio_frame, text=val, value=val, variable=self.ratio_choice, style="TRadiobutton")
            radio.pack(side="left", padx=10)
        self.generate_button = ttk.Button(main_frame, text="Generate Duty Chart", command=self.run, style="TButton")
        self.generate_button.pack(pady=20)
        self.progress = ttk.Progressbar(main_frame, mode="determinate", maximum=100, style="TProgressbar")
        self.progress.pack(fill="x", pady=5)
        summary_frame = ttk.LabelFrame(main_frame, text="Summary", padding=10, style="TLabelframe")
        summary_frame.pack(fill="both", expand=True, pady=5)
        self.summary_box = tk.Text(
            summary_frame, height=15, wrap="word",
            font=("Segoe UI", 10), background=self.summary_bg, fg=self.sum_fg,
            relief="flat", borderwidth=0
        )
        scrollbar = ttk.Scrollbar(summary_frame, orient="vertical", command=self.summary_box.yview, style="Vertical.TScrollbar")
        self.summary_box.config(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        self.summary_box.pack(fill="both", expand=True)
        self.summary_box.tag_configure("header", font=("Segoe UI", 12, "bold"), foreground="#06a0c0")
        self.tooltip = None
        self.generate_button.bind("<Enter>", self.on_button_hover)
        self.generate_button.bind("<Leave>", self.on_button_leave)

    def on_button_hover(self, event):
        self.style.configure("TButton", background=self.hover_bg)
    def on_button_leave(self, event):
        self.style.configure("TButton", background=self.btn_bg)
    def show_tooltip(self, widget, text):
        if self.tooltip:
            self.tooltip.destroy()
        try:
            x, y, _, _ = widget.bbox("insert")
        except:
            x = y = 0
        x += widget.winfo_rootx() + 25
        y += widget.winfo_rooty() + 25
        self.tooltip = tk.Toplevel(widget)
        self.tooltip.wm_overrideredirect(True)
        self.tooltip.wm_geometry(f"+{x}+{y}")
        label = tk.Label(self.tooltip, text=text, background=self.hover_bg, foreground=self.fg_color, relief="solid", borderwidth=1,
                         font=("Segoe UI", 9))
        label.pack()
    def hide_tooltip(self):
        if self.tooltip:
            self.tooltip.destroy()
            self.tooltip = None
    def on_resize(self, event):
        width = self.root.winfo_width()
        entry_width = max(30, int(width / 20))
        try:
            for entry in [self.root.winfo_children()[0].winfo_children()[1].winfo_children()[0],
                          self.root.winfo_children()[0].winfo_children()[2].winfo_children()[0]]:
                entry.configure(width=entry_width)
        except Exception:
            pass
    def browse_input(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            self.input_path.set(file_path)
    def browse_output(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                 filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            self.output_path.set(file_path)
    def run(self):
        try:
            slot1 = (self.slot1_start.get_date(), self.slot1_end.get_date())
            slot2 = (self.slot2_start.get_date(), self.slot2_end.get_date())
            if slot1[0] > slot1[1] or slot2[0] > slot2[1]:
                messagebox.showerror("Error", "Start date must be before end date for both slots.")
                return
            if not self.input_path.get().strip() or not self.output_path.get().strip():
                messagebox.showerror("Error", "Please select both input and output files.")
                return

            self.generate_button.config(state='disabled')
            self.progress['value'] = 0
            self.root.update()

            # Run scheduler
            result = generate_duty_chart(
                self.input_path.get(),
                self.output_path.get(),
                slot1,
                slot2,
                self.ratio_choice.get()
            )

            # Check for error
            if result[0] is None:
                messagebox.showerror("Error", "Failed to generate chart. See log for details.")
                self.generate_button.config(state='normal')
                return

            # Unpack safely
            assignment_summary, ratio_violations, duty_quota_violations, slot_violations, _ = result

            # Update progress
            for i in range(0, 81, 20):
                self.progress['value'] = i
                self.root.update()
                self.root.after(100)
            self.progress['value'] = 100
            self.root.update()

            # Show summary
            self.summary_box.delete("1.0", tk.END)
            self.summary_box.insert(tk.END, "Assignment Summary:\n", "header")
            self.summary_box.insert(tk.END, f"{assignment_summary}\n\n")
            self.summary_box.insert(tk.END, "Duty Quota Violations:\n", "header")
            self.summary_box.insert(tk.END, "\n".join(duty_quota_violations) + "\n" if duty_quota_violations else "No duty quota violations.\n")
            self.summary_box.insert(tk.END, "\nSlot Preference Violations:\n", "header")
            filtered = [v for v in slot_violations if "(A.P(Contract))" not in v]
            self.summary_box.insert(tk.END, "\n".join(filtered) + "\n" if filtered else "No slot preference violations (permanent staff).\n")

            messagebox.showinfo("Success", "Duty chart generated successfully!")

        except Exception as e:
            logging.error(f"GUI run error: {str(e)}")
            messagebox.showerror("Error", f"Unexpected error: {str(e)}")
        finally:
            self.progress['value'] = 0
            self.generate_button.config(state='normal')

