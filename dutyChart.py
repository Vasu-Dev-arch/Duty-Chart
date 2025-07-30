import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkcalendar import DateEntry
import math
import pandas as pd
import re
import os
from datetime import datetime, timedelta
from difflib import SequenceMatcher
import logging
from tkinter import messagebox

logging.basicConfig(filename='duty_chart_app.log', level=logging.DEBUG,
                    format='%(asctime)s - %(levelname)s - %(message)s')

def normalize_name(name):
    if pd.isna(name): return ""
    cleaned = re.sub(r"^(Dr\.?|Prof\.?|Mr\.?|Mrs\.?|Ms\.?)\s*", "", str(name).strip(), flags=re.IGNORECASE)
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    cleaned = re.sub(r"\.(?=\w)", " ", cleaned)
    parts = [part.lower() for part in cleaned.split() if part]
    return ' '.join(parts)

def normalize_designation(desig):
    if pd.isna(desig): return ""
    desig = str(desig).strip().lower()
    designation_map = {
        'professor': 'Professor', 'prof': 'Professor',
        'assoc. professor': 'Assoc. Professor', 'asp': 'Assoc. Professor', 'associate prof': 'Assoc. Professor',
        'asst. professor': 'Asst. Professor', 'ap': 'Asst. Professor', 'asst prof': 'Asst. Professor',
        'a.p(contract)': 'A.P(Contract)', 'gl': 'A.P(Contract)', 'guest lecturer': 'A.P(Contract)'
    }
    return designation_map.get(desig, desig.title())

def fuzzy_match_name(staff_name, pref_name, threshold=0.65):
    try:
        staff_norm = normalize_name(staff_name)
        pref_norm = normalize_name(pref_name)
        staff_parts = staff_norm.split()
        pref_parts = pref_norm.split()
        if set(staff_parts) & set(pref_parts):
            return True
        score = SequenceMatcher(None, staff_norm, pref_norm).ratio()
        staff_raw = re.sub(r"[.\s]+", "", str(staff_name).lower())
        pref_raw = re.sub(r"[.\s]+", "", str(pref_name).lower())
        raw_score = SequenceMatcher(None, staff_raw, pref_raw).ratio()
        return score >= threshold or raw_score >= 0.9
    except Exception as e:
        logging.error(f"Fuzzy match failed for {staff_name} {pref_name}: {e}")
        return False

def find_column(df, keywords):
    for col in df.columns:
        col_clean = col.strip().lower().replace('\n', '').replace('\r', '')
        if any(keyword.lower() in col_clean for keyword in keywords):
            return col
    return None

def safe_parse_date(val):
    try:
        if pd.isna(val): return None
        if isinstance(val, (int, float)):
            base_date = datetime(1899, 12, 30)
            return (base_date + timedelta(days=int(val))).date()
        if isinstance(val, pd.Timestamp):
            return val.date()
        return pd.to_datetime(val).date()
    except Exception as e:
        logging.error(f"Failed to parse date {val}: {e}")
        return None

def parse_timestamp(ts):
    try:
        parsed = pd.to_datetime(ts, errors='coerce')
        if pd.isna(parsed): return pd.NaT
        return parsed.tz_localize(None)
    except Exception as e:
        logging.error(f"Failed to parse timestamp {ts}: {e}")
        return pd.NaT

def can_assign(name, date, used_on_day):
    if date not in used_on_day: return True
    return name not in used_on_day[date]

def generate_duty_chart(input_path, output_path, slot1_range, slot2_range, ratio_choice):
    try:
        input_path = input_path.strip('"').strip()
        output_path = output_path.strip('"').strip()
        if not os.path.exists(input_path):
            messagebox.showerror("Error", f"Input file not found: {input_path}")
            return None, [], [], [], {}

        # Load sheets
        xls = pd.ExcelFile(input_path)
        sheets = {s.strip().lower().replace('\n','').replace('\r',''):s for s in xls.sheet_names}
        def get_sheet(variants):
            for v in variants:
                vclean = v.strip().lower().replace('\n','').replace('\r','')
                if vclean in sheets:
                    return sheets[vclean]
            return None
        session_sheet = get_sheet(['session strength', 'sessionwise strength'])
        staff_sheet = get_sheet(['staff list', 'staff details'])
        pref_sheet = get_sheet(['slot preference'])
        if not session_sheet or not staff_sheet or not pref_sheet:
            messagebox.showerror("Error", "Required sheets missing in input Excel.")
            return None, [], [], [], {}

        session_df = pd.read_excel(xls, session_sheet)
        staff_df = pd.read_excel(xls, staff_sheet)
        pref_df = pd.read_excel(xls, pref_sheet)

        session_df.columns = [c.strip().lower().replace('\n','').replace('\r','') for c in session_df.columns]
        staff_df.columns = [c.strip().lower().replace('\n','').replace('\r','') for c in staff_df.columns]
        pref_df.columns = [c.strip().lower().replace('\n','').replace('\r','') for c in pref_df.columns]

        session_cols = {'date': find_column(session_df, ['date']), 'fn': find_column(session_df, ['fn', 'forenoon', 'morning']), 'an': find_column(session_df, ['an', 'afternoon'])}
        staff_cols = {'name': find_column(staff_df, ['name']), 'designation': find_column(staff_df, ['designation'])}
        pref_cols = {'timestamp': find_column(pref_df, ['timestamp']), 'name': find_column(pref_df, ['name']), 'preferred slot': find_column(pref_df, ['preferred slot', 'slot'])}

        missing_cols = []
        for df_name, cols in [('Session Strength', session_cols), ('Staff List', staff_cols), ('Slot Preference', pref_cols)]:
            for col_name, col in cols.items():
                if col is None:
                    missing_cols.append(f"{col_name} in {df_name}")
        if missing_cols:
            messagebox.showerror("Error", f"Missing columns: {', '.join(missing_cols)}")
            return None, [], [], [], {}

        session_df = session_df.rename(columns={session_cols['date']:'date', session_cols['fn']:'fn', session_cols['an']:'an'})
        staff_df = staff_df.rename(columns={staff_cols['name']:'name', staff_cols['designation']:'designation'})
        pref_df = pref_df.rename(columns={pref_cols['timestamp']:'timestamp', pref_cols['name']:'name', pref_cols['preferred slot']:'preferred slot'})

        session_df['date'] = session_df['date'].apply(safe_parse_date)
        session_df = session_df.dropna(subset=['date'])
        session_df['fn'] = pd.to_numeric(session_df['fn'], errors='coerce').fillna(0)
        session_df['an'] = pd.to_numeric(session_df['an'], errors='coerce').fillna(0)

        staff_df['original_name'] = staff_df['name']
        staff_df['name'] = staff_df['name'].apply(normalize_name)
        staff_df['designation'] = staff_df['designation'].apply(normalize_designation)
        staff_df = staff_df.drop_duplicates(subset=['name'])

        pref_df['original_name'] = pref_df['name']
        pref_df['name'] = pref_df['name'].apply(normalize_name)
        pref_df['preferred slot'] = pref_df['preferred slot'].str.strip().str.title().replace('', 'Any')
        pref_df['timestamp'] = pref_df['timestamp'].apply(parse_timestamp)
        pref_df = pref_df.sort_values('timestamp').drop_duplicates(subset=['name'], keep='last')

        # fuzzy matching pref names to staff names if needed
        staff_names = set(staff_df['name'])
        pref_names = set(pref_df['name'])
        unmatched_pref = pref_names - staff_names
        unmatched_staff = staff_names - pref_names
        fuzzy_matches = {}
        for p_name in unmatched_pref:
            for s_name in unmatched_staff:
                if fuzzy_match_name(s_name, p_name):
                    fuzzy_matches[p_name] = s_name
                    break
        pref_df['name'] = pref_df['name'].replace(fuzzy_matches)
        pref_df = pref_df[pref_df['name'].isin(staff_names)]

        merged_df = pd.merge(
            staff_df[['name', 'original_name', 'designation']],
            pref_df[['name', 'original_name', 'timestamp', 'preferred slot']],
            on='name', how='left'
        )
        merged_df['preferred slot'] = merged_df['preferred slot'].fillna('Any')
        merged_df['original_name_x'] = merged_df['original_name_x'].fillna(merged_df['name'])
        merged_df = merged_df.rename(columns={'original_name_x':'original_name'}).drop(columns=['original_name_y'], errors='ignore')
        merged_df = merged_df.drop_duplicates(subset=['name'])

        all_dates = sorted(session_df['date'].unique())
        slot_dates = {'Slot 1': set(), 'Slot 2': set()}
        for d in all_dates:
            if slot1_range[0] <= d <= slot1_range[1]:
                slot_dates['Slot 1'].add(d)
            elif slot2_range[0] <= d <= slot2_range[1]:
                slot_dates['Slot 2'].add(d)

        sessions = []
        for _, row in session_df.iterrows():
            for label in ['fn','an']:
                count = row[label]
                if count > 0:
                    date = row["date"]
                    sessions.append((date, label.upper(), math.ceil(count / 30), count))
        sessions.sort(key=lambda x: (x[0], x[1]))

        # User-selected duty ratio caps
        ratio_map = {
            '1:3:6': {'Professor': 1, 'Assoc. Professor': 3, 'Asst. Professor': 6},
            '1:3:7': {'Professor': 1, 'Assoc. Professor': 3, 'Asst. Professor': 7},
            '1:4:8': {'Professor': 1, 'Assoc. Professor': 4, 'Asst. Professor': 8},
        }
        if ratio_choice not in ratio_map:
            messagebox.showerror("Error", f"Invalid duty ratio selected: {ratio_choice}")
            return None, [], [], [], {}
        designation_caps = ratio_map[ratio_choice]
        designation_caps['A.P(Contract)'] = float('inf')

        assigned_counts = {name: 0 for name in merged_df['name']}
        used_on_day = {d: set() for d in all_dates}
        duty_data = {name: {} for name in merged_df['name']}
        staff_slot_assignment = {}

        ratio_violations = []
        duty_quota_violations = []
        slot_preference_violations = []

        for date, session, needed, student_count in sessions:
            used_on_day.setdefault(date, set())
            remaining = needed
            slot_of_date = 'Slot 1' if date in slot_dates['Slot 1'] else 'Slot 2'

            def assign_perms(desg):
                nonlocal remaining
                perms = merged_df[merged_df['designation'] == desg]
                if desg in ['Professor', 'Assoc. Professor']:
                    perms = perms[perms['preferred slot'] == slot_of_date]
                    perm_list = perms.to_dict('records')
                else:
                    pref = perms[perms['preferred slot'] == slot_of_date].sort_values('timestamp', na_position='last')
                    anyp = perms[perms['preferred slot'] == 'Any'].sort_values('timestamp', na_position='last')
                    perm_list = pd.concat([pref, anyp]).to_dict('records')
                for p in perm_list:
                    n = p['name']
                    if assigned_counts[n] >= designation_caps[desg]:
                        continue
                    if n in used_on_day[date]:
                        continue
                    if n in staff_slot_assignment and staff_slot_assignment[n] != slot_of_date:
                        continue
                    if n not in staff_slot_assignment:
                        staff_slot_assignment[n] = slot_of_date
                    if can_assign(n, date, used_on_day):
                        duty_data.setdefault(n, {})
                        duty_data[n].setdefault(date, []).append(session)
                        used_on_day[date].add(n)
                        assigned_counts[n] += 1
                        remaining -= 1
                        if p['preferred slot'] != 'Any' and p['preferred slot'] != slot_of_date:
                            slot_preference_violations.append(f"{p['original_name']} ({desg}) assigned to {slot_of_date} but preferred {p['preferred slot']}")
                        if remaining == 0:
                            break
                return

            for d in ['Professor', 'Assoc. Professor', 'Asst. Professor']:
                assign_perms(d)
                if remaining == 0:
                    break

            if remaining > 0:
                gls = merged_df[merged_df['designation'] == 'A.P(Contract)']['name']
                available_gls = [n for n in gls if n not in used_on_day[date]
                                 and (n not in staff_slot_assignment or staff_slot_assignment[n] == slot_of_date)]
                available_gls = sorted(available_gls, key=lambda x: assigned_counts[x])
                for n in available_gls:
                    if remaining == 0:
                        break
                    if n not in staff_slot_assignment:
                        staff_slot_assignment[n] = slot_of_date
                    if can_assign(n, date, used_on_day):
                        duty_data.setdefault(n, {})
                        duty_data[n].setdefault(date, []).append(session)
                        used_on_day[date].add(n)
                        assigned_counts[n] += 1
                        remaining -= 1

            actually_assigned = sum(
                1 for name in assigned_counts if name in duty_data and date in duty_data[name] and session in duty_data[name][date]
            )
            if actually_assigned < needed:
                ratio_violations.append(f"{date} {session}: {needed} invigilators needed for {student_count} students, but only {actually_assigned} assigned.")

        # Report perm staff under-caps
        for _, person in merged_df[merged_df["designation"].isin(["Professor", "Assoc. Professor", "Asst. Professor"])].iterrows():
            name = person['name']
            cap = designation_caps.get(person['designation'], 0)
            assigned = assigned_counts.get(name, 0)
            if assigned < cap:
                duty_quota_violations.append(f"{person['original_name']} ({person['designation']}) assigned {assigned}/{cap} duties")

        output_rows = []
        for _, row in merged_df.iterrows():
            name = row['name']
            desig = row['designation']
            orig_name = row['original_name']
            dept = row['department'] if 'department' in merged_df.columns else ''
            assigned_slot = staff_slot_assignment.get(name)
            assigned_slot_str = assigned_slot if assigned_slot in ['Slot 1', 'Slot 2'] else 'None'
            total_duties = assigned_counts.get(name, 0)
            user_row = {'Name': orig_name, 'Designation': desig, 'Department': dept,
                        'Total Duties': total_duties, 'Assigned Slot': assigned_slot_str}
            for d in all_dates:
                sessions_assigned = duty_data.get(name, {}).get(d, [])
                user_row[d] = ' '.join(sessions_assigned) if sessions_assigned else ''
            output_rows.append(user_row)

        output_df = pd.DataFrame(output_rows)
        col_order = ['Name', 'Designation', 'Department', 'Total Duties', 'Assigned Slot'] + all_dates
        output_df = output_df[col_order]
        output_df.to_excel(output_path, index=False)

        total = sum(assigned_counts.values())
        prof_count = sum(assigned_counts[n] for n in merged_df[merged_df['designation'] == 'Professor']['name'])
        asp_count = sum(assigned_counts[n] for n in merged_df[merged_df['designation'] == 'Assoc. Professor']['name'])
        ap_count = sum(assigned_counts[n] for n in merged_df[merged_df['designation'] == 'Asst. Professor']['name'])
        gl_count = sum(assigned_counts[n] for n in merged_df[merged_df['designation'] == 'A.P(Contract)']['name'])
        summary = f"Final chart (Ratio: {ratio_choice}): {prof_count} Professor, {asp_count} Assoc. Professor, {ap_count} Asst. Professor, {gl_count} A.P(Contract), Total duties assigned: {total}"
        return summary, ratio_violations, duty_quota_violations, slot_preference_violations, merged_df[['name', 'original_name']].set_index('name')['original_name'].to_dict()

    except Exception as e:
        logging.error(f"Failed to generate chart: {str(e)}")
        messagebox.showerror("Error", f"Failed to generate chart: {str(e)}\nCheck duty_chart_app.log for details.")
        return None, [], [], [], {}

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
            btn_fg = "#222222"  # Button text dark
            hover_bg = "#2d2d2d"
            entry_bg = "#232336"
            entry_fg = "#222222"
            summary_bg = "#232336"
            sum_fg = "#f6f6f6"
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

        self.bg_color = bg_color
        self.fg_color = fg_color
        self.btn_bg = btn_bg
        self.btn_fg = btn_fg
        self.hover_bg = hover_bg
        self.entry_bg = entry_bg
        self.entry_fg = entry_fg
        self.summary_bg = summary_bg
        self.sum_fg = sum_fg

        self.style.configure("TFrame", background=bg_color)
        self.style.configure("TLabel", font=("Segoe UI", 10), background=bg_color, foreground=fg_color)
        self.style.configure("Title.TLabel", font=("Segoe UI", 20, "bold"), foreground="#06a0c0", background=bg_color)
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
        ttk.Label(main_frame, text="Duty Chart Generator", style="Title.TLabel").pack(pady=10)
        toggle_btn = ttk.Button(main_frame, text="Toggle Theme", command=self.toggle_theme, style="TButton")
        toggle_btn.pack(pady=(0, 10))
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
            for entry in [self.root.winfo_children()[0].winfo_children()[2].winfo_children()[0],
                          self.root.winfo_children()[0].winfo_children()[3].winfo_children()[0]]:
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
            for i in range(0, 81, 20):
                self.progress['value'] = i
                self.root.update()
                self.root.after(100)
            assignment_summary, ratio_violations, duty_quota_violations, slot_preference_violations, _ = generate_duty_chart(
                self.input_path.get(), self.output_path.get(), slot1, slot2, self.ratio_choice.get())
            self.progress['value'] = 100
            self.root.update()
            self.summary_box.delete("1.0", tk.END)
            self.summary_box.insert(tk.END, "Assignment Summary:\n", "header")
            self.summary_box.insert(tk.END, f"{assignment_summary if assignment_summary else 'All sessions assigned'}\n\n")
            self.summary_box.insert(tk.END, "70:30 Rule Violations:\n", "header")
            self.summary_box.insert(
                tk.END, "\n".join(ratio_violations) + "\n" if ratio_violations else "No 70:30 rule violations.\n")
            self.summary_box.insert(tk.END, "\nDuty Quota Violations:\n", "header")
            self.summary_box.insert(
                tk.END, "\n".join(duty_quota_violations) + "\n" if duty_quota_violations else "No duty quota violations.\n")
            self.summary_box.insert(tk.END, "\nSlot Preference Violations:\n", "header")
            self.summary_box.insert(
                tk.END, "\n".join(slot_preference_violations) + "\n" if slot_preference_violations else "No slot preference violations.\n")
            messagebox.showinfo("Success", "Duty chart generated successfully! Check the output file and log for details.")
        except Exception as e:
            logging.error(f"GUI run error: {str(e)}")
            messagebox.showerror("Error", f"Error: {str(e)}\nCheck duty_chart_app.log for details.")
        finally:
            self.progress['value'] = 0
            self.generate_button.config(state='normal')
if __name__ == "__main__":
    root = tk.Tk()
    app = DutyChartApp(root)
    root.mainloop()

