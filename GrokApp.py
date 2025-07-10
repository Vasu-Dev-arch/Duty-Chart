try:
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk
    from tkcalendar import DateEntry
    from datetime import timedelta
except ImportError as e:
    print(f"Required modules (tkinter, tkcalendar) missing. Install them (e.g., `pip install tkcalendar`).")
    exit(1)

import pandas as pd
import numpy as np
from datetime import datetime
import re
import os
import math
import logging
from difflib import SequenceMatcher

# Setup logging
logging.basicConfig(filename='duty_chart_app.log', level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

# ----------------------------- Helper Functions -----------------------------
def normalize_name(name):
    if pd.isna(name):
        return ""
    cleaned = re.sub(r"^(Dr\.?|Prof\.?|Mr\.?|Mrs\.?|Ms\.?)\s*", "", str(name).strip(), flags=re.IGNORECASE)
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    return cleaned.lower()

def fuzzy_match_name(name1, name2, threshold=0.9):
    try:
        score = SequenceMatcher(None, name1.lower(), name2.lower()).ratio()
        return score >= threshold
    except:
        logging.error(f"Fuzzy match failed for {name1} vs {name2}")
        return False

def find_column(df, keywords):
    for col in df.columns:
        col_clean = col.strip().lower().replace('\n', '').replace('\r', '')
        if any(keyword.lower() in col_clean for keyword in keywords):
            logging.info(f"Found column '{col}' matching keywords {keywords}")
            return col
    logging.warning(f"No column found matching keywords {keywords}")
    return None

def safe_parse_date(val):
    try:
        if pd.isna(val):
            return None
        if isinstance(val, (int, float)):
            base_date = datetime(1899, 12, 30)
            return (base_date + timedelta(days=int(val))).date()
        if isinstance(val, pd.Timestamp):
            return val.date()
        return pd.to_datetime(val).date()
    except:
        logging.error(f"Failed to parse date {val}")
        return None

def parse_timestamp(ts):
    try:
        parsed = pd.to_datetime(ts, errors='coerce')
        if pd.isna(parsed):
            logging.error(f"Invalid timestamp: {ts}")
            return pd.NaT
        return parsed.tz_localize(None)
    except:
        logging.error(f"Failed to parse timestamp {ts}")
        return pd.NaT

# ----------------------------- Duty Chart Generator -----------------------------
def generate_duty_chart(input_path, output_path, slot1_range, slot2_range):
    try:
        input_path = input_path.strip('"').strip()
        output_path = output_path.strip('"').strip()
        logging.info(f"Input path: {input_path}, Output path: {output_path}")

        # Load Excel file
        if not os.path.exists(input_path):
            logging.error(f"Input file not found: {input_path}")
            messagebox.showerror("Error", f"Input file not found: {input_path}")
            return None, None, None, None, {}

        xls = pd.ExcelFile(input_path)
        sheets = {s.strip().lower().replace('\n', '').replace('\r', ''): s for s in xls.sheet_names}

        # Find sheets
        sheet_map = {
            'session strength': ['Session Strength', 'session strength', 'student strength', 'sessionstrength'],
            'staff list': ['Staff List', 'staff list', 'faculty list', 'faculty'],
            'slot preference': ['Slot Preference', 'slot preference', 'slot preferences', 'preference', 'preferences']
        }
        found_sheets = {}
        for key, variations in sheet_map.items():
            for variation in variations:
                if variation.strip().lower().replace('\n', '').replace('\r', '') in sheets:
                    found_sheets[key] = sheets[variation.strip().lower().replace('\n', '').replace('\r', '')]
                    break
            if key not in found_sheets:
                logging.error(f"Missing sheet: {key}. Found sheets: {', '.join(xls.sheet_names)}")
                messagebox.showerror("Error", f"Missing sheet: {key}. Found sheets: {', '.join(xls.sheet_names)}")
                return None, None, None, None, {}

        # Load data
        session_df = pd.read_excel(xls, found_sheets['session strength'])
        staff_df = pd.read_excel(xls, found_sheets['staff list'])
        pref_df = pd.read_excel(xls, found_sheets['slot preference'])

        # Normalize column names
        session_df.columns = [c.strip().lower().replace('\n', '').replace('\r', '') for c in session_df.columns]
        staff_df.columns = [c.strip().lower().replace('\n', '').replace('\r', '') for c in staff_df.columns]
        pref_df.columns = [c.strip().lower().replace('\n', '').replace('\r', '') for c in pref_df.columns]

        # Find required columns
        session_cols = {
            'date': find_column(session_df, ['date']),
            'fn': find_column(session_df, ['fn', 'forenoon', 'morning']),
            'an': find_column(session_df, ['an', 'afternoon'])
        }
        staff_cols = {
            'name': find_column(staff_df, ['name of the faculty', 'name', 'faculty']),
            'designation': find_column(staff_df, ['designation', 'design', 'desig'])
        }
        pref_cols = {
            'timestamp': find_column(pref_df, ['timestamp']),
            'name': find_column(pref_df, ['name of the faculty', 'name', 'faculty']),
            'preferred slot': find_column(pref_df, ['preferred slot', 'slot', 'preferredslot'])
        }

        # Validate columns
        missing_cols = []
        for df_name, cols in [('Session Strength', session_cols), ('Staff List', staff_cols), ('Slot Preference', pref_cols)]:
            for col_name, col in cols.items():
                if col is None:
                    missing_cols.append(f"{col_name} in {df_name}")
        if missing_cols:
            logging.error(f"Missing columns: {', '.join(missing_cols)}")
            messagebox.showerror("Error", f"Missing columns: {', '.join(missing_cols)}")
            return None, None, None, None, {}

        # Rename columns
        session_df = session_df.rename(columns={session_cols['date']: 'date', session_cols['fn']: 'fn', session_cols['an']: 'an'})
        staff_df = staff_df.rename(columns={staff_cols['name']: 'name', staff_cols['designation']: 'designation'})
        pref_df = pref_df.rename(columns={pref_cols['timestamp']: 'timestamp', pref_cols['name']: 'name', 
                                         pref_cols['preferred slot']: 'preferred slot'})

        # Drop extra columns
        for col in pref_df.columns:
            if col != 'name' and find_column(pref_df, ['name of the faculty', 'name', 'faculty']) == col:
                pref_df = pref_df.drop(columns=[col])
            if col != 'designation' and find_column(pref_df, ['designation', 'design', 'desig']) == col:
                pref_df = pref_df.drop(columns=[col])
        if 'designation' in pref_df.columns:
            pref_df = pref_df.drop(columns=['designation'])

        # Process data
        session_df['date'] = session_df['date'].apply(safe_parse_date)
        session_df = session_df.dropna(subset=['date'])
        session_df['fn'] = pd.to_numeric(session_df['fn'], errors='coerce').fillna(0)
        session_df['an'] = pd.to_numeric(session_df['an'], errors='coerce').fillna(0)

        staff_df['original_name'] = staff_df['name']
        staff_df['name'] = staff_df['name'].apply(normalize_name)
        staff_df['designation'] = staff_df['designation'].str.strip().str.upper()
        staff_df = staff_df.drop_duplicates(subset=['name'])

        pref_df['original_name'] = pref_df['name']
        pref_df['name'] = pref_df['name'].apply(normalize_name)
        pref_df['preferred slot'] = pref_df['preferred slot'].str.strip().str.title().replace('', 'Any')
        pref_df['timestamp'] = pref_df['timestamp'].apply(parse_timestamp)
        pref_df = pref_df.sort_values('timestamp').drop_duplicates(subset=['name'], keep='last')

        # Fuzzy match names
        staff_names = set(staff_df['name'])
        pref_names = set(pref_df['name'])
        unmatched_staff = staff_names - pref_names
        unmatched_pref = pref_names - staff_names
        fuzzy_matches = {}
        for staff_name in unmatched_staff:
            for pref_name in unmatched_pref:
                if fuzzy_match_name(staff_name, pref_name):
                    fuzzy_matches[pref_name] = staff_name
                    logging.info(f"Fuzzy matched {pref_name} (Preference) to {staff_name} (Staff)")
        if fuzzy_matches:
            pref_df['name'] = pref_df['name'].replace(fuzzy_matches)
            unmatched_pref = pref_names - staff_names - set(fuzzy_matches.keys())
            unmatched_staff = staff_names - pref_names - set(fuzzy_matches.values())

        if unmatched_staff:
            logging.info(f"Staff names not in Slot Preference (defaulting to Any): {unmatched_staff}")
        if unmatched_pref:
            logging.warning(f"Preference names not in Staff List (ignored): {unmatched_pref}")
            pref_df = pref_df[pref_df['name'].isin(staff_names)]

        # Merge data
        merged_df = pd.merge(staff_df[['name', 'original_name', 'designation']], 
                            pref_df[['name', 'original_name', 'timestamp', 'preferred slot']], 
                            on='name', how='left')
        merged_df['preferred slot'] = merged_df['preferred slot'].fillna('Any')
        merged_df['original_name_x'] = merged_df['original_name_x'].fillna(merged_df['name'])
        merged_df = merged_df.rename(columns={'original_name_x': 'original_name'})
        merged_df = merged_df.drop(columns=['original_name_y'], errors='ignore')
        merged_df = merged_df.drop_duplicates(subset=['name'])

        # Validate slot dates
        all_dates = sorted(session_df['date'].unique())
        slot_dates = {'Slot 1': set(), 'Slot 2': set()}
        for d in all_dates:
            if slot1_range[0] <= d <= slot1_range[1]:
                slot_dates['Slot 1'].add(d)
            elif slot2_range[0] <= d <= slot2_range[1]:
                slot_dates['Slot 2'].add(d)
        logging.info(f"Slot 1 dates: {sorted(slot_dates['Slot 1'])}, Slot 2 dates: {sorted(slot_dates['Slot 2'])}")

        # Calculate duties
        slot1_duties = sum(math.ceil(row[s] / 30) for _, row in session_df.iterrows() for s in ['fn', 'an'] if row['date'] in slot_dates['Slot 1'])
        slot2_duties = sum(math.ceil(row[s] / 30) for _, row in session_df.iterrows() for s in ['fn', 'an'] if row['date'] in slot_dates['Slot 2'])
        logging.info(f"Slot 1 needs {slot1_duties} duties, Slot 2 needs {slot2_duties} duties")

        # Define sessions
        sessions = [(row['date'], s, math.ceil(row[s] / 30)) for _, row in session_df.iterrows() for s in ['fn', 'an'] if math.ceil(row[s] / 30) > 0]
        sessions.sort(key=lambda x: (x[0], -x[2]))  # Sort by date, then by required duties (descending)
        logging.info(f"Sessions to assign: {len(sessions)}")

        # Initialize tracking
        assigned_counts = {name: 0 for name in merged_df['name']}
        used_on_day = {d: set() for d in all_dates}
        duty_data = {name: {} for name in merged_df['name']}
        assigned_slots = {name: None for name in merged_df['name']}

        # Duty caps configurations
        duty_configs = [
            {'PROF': 1, 'ASP': 3, 'AP': 6, 'GL': float('inf'), 'perm_ratio': 0.7, 'gl_ratio': 0.3, 'name': '1:3:6 (70:30)'},
            {'PROF': 1, 'ASP': 4, 'AP': 8, 'GL': float('inf'), 'perm_ratio': 0.7, 'gl_ratio': 0.3, 'name': '1:4:8 (70:30)'}
        ]

        ratio_violations = []
        duty_quota_violations = []

        # Assignment with fallback logic
        for config_idx, designation_caps in enumerate(duty_configs):
            assigned_counts = {name: 0 for name in merged_df['name']}
            used_on_day = {d: set() for d in all_dates}
            duty_data = {name: {} for name in merged_df['name']}
            assigned_slots = {name: None for name in merged_df['name']}
            ratio_violations = []
            duty_quota_violations = []
            success = True

            # Assign permanent staff (PROF, ASP, AP)
            for desig in ['PROF', 'ASP', 'AP']:
                candidates = merged_df[merged_df['designation'] == desig][['name', 'original_name', 'preferred slot', 'timestamp']]
                if desig == 'PROF':
                    candidates = sorted(candidates.to_dict('records'), key=lambda x: x['name'] != normalize_name('Dr. K. Venkatesan'))
                elif desig == 'AP':
                    candidates = sorted(candidates.to_dict('records'), key=lambda x: x['timestamp'] if not pd.isna(x['timestamp']) else pd.Timestamp.max)
                else:
                    candidates = candidates.to_dict('records')

                for candidate in candidates:
                    name = candidate['name']
                    orig_name = candidate['original_name']
                    pref_slot = candidate['preferred slot'] if candidate['preferred slot'] in ['Slot 1', 'Slot 2'] else 'Any'
                    valid_slots = [pref_slot] if pref_slot in ['Slot 1', 'Slot 2'] and desig in ['PROF', 'ASP'] else ['Slot 1', 'Slot 2'] if pref_slot == 'Any' else [pref_slot]
                    duties_needed = designation_caps[desig]
                    assigned = 0

                    for slot in valid_slots:
                        valid_dates = sorted(slot_dates[slot])
                        for date, session, required in [(d, s, r) for d, s, r in sessions if d in valid_dates]:
                            if name not in used_on_day[date] and assigned_counts[name] < duties_needed:
                                current_slot = 'Slot 1' if date in slot_dates['Slot 1'] else 'Slot 2'
                                if assigned_slots[name] is not None and assigned_slots[name] != current_slot and desig in ['ASP', 'AP']:
                                    continue  # Prevent slot splitting
                                current_assigned = len([n for n in used_on_day[date] if date in duty_data[n] and session.upper() in duty_data[n][date]])
                                perm_needed = math.ceil(required * designation_caps['perm_ratio'])
                                perm_assigned = len([n for n in used_on_day[date] if date in duty_data[n] and session.upper() in duty_data[n][date] and merged_df[merged_df['name'] == n]['designation'].iloc[0] in ['PROF', 'ASP', 'AP']])
                                if perm_assigned < perm_needed and current_assigned < required:
                                    duty_data[name][date] = duty_data[name].get(date, []) + [session.upper()]
                                    used_on_day[date].add(name)
                                    assigned_counts[name] += 1
                                    assigned += 1
                                    assigned_slots[name] = current_slot
                                    logging.info(f"Assigned {orig_name} ({desig}) to {date} {session} (Slot {current_slot})")
                                    if assigned == duties_needed:
                                        break
                        if assigned == duties_needed:
                            break
                    if assigned < duties_needed:
                        duty_quota_violations.append(f"{orig_name} ({desig}) assigned {assigned}/{duties_needed} duties (preferred slot: {pref_slot})")

            # Ensure APs get full duty quota
            for name in merged_df[merged_df['designation'] == 'AP']['name']:
                if assigned_counts[name] < designation_caps['AP']:
                    orig_name = merged_df[merged_df['name'] == name]['original_name'].iloc[0]
                    pref_slot = merged_df[merged_df['name'] == name]['preferred slot'].iloc[0]
                    duties_needed = designation_caps['AP']
                    assigned = assigned_counts[name]
                    valid_slots = [pref_slot] if pref_slot in ['Slot 1', 'Slot 2'] else ['Slot 1', 'Slot 2']
                    for slot in valid_slots:
                        if assigned_slots[name] is not None and assigned_slots[name] != slot:
                            continue  # Respect no-split rule
                        valid_dates = sorted(slot_dates[slot])
                        # Try reassigning GL duties
                        for date, session, required in [(d, s, r) for d, s, r in sessions if d in valid_dates]:
                            if name not in used_on_day[date] and assigned_counts[name] < duties_needed:
                                gls_assigned = [n for n in used_on_day[date] if date in duty_data[n] and session.upper() in duty_data[n][date] and merged_df[merged_df['name'] == n]['designation'].iloc[0] == 'GL']
                                if gls_assigned:
                                    gl_name = gls_assigned[0]
                                    gl_orig_name = merged_df[merged_df['name'] == gl_name]['original_name'].iloc[0]
                                    duty_data[gl_name][date].remove(session.upper())
                                    if not duty_data[gl_name][date]:
                                        del duty_data[gl_name][date]
                                        used_on_day[date].remove(gl_name)
                                    assigned_counts[gl_name] -= 1
                                    duty_data[name][date] = duty_data[name].get(date, []) + [session.upper()]
                                    used_on_day[date].add(name)
                                    assigned_counts[name] += 1
                                    assigned += 1
                                    assigned_slots[name] = slot
                                    logging.info(f"Reassigned {gl_orig_name} (GL) duty to {orig_name} (AP) for {date} {session} (Slot {slot})")
                                    if assigned == duties_needed:
                                        break
                        if assigned == duties_needed:
                            break
                        # Relax 70:30 if necessary
                        for date, session, required in [(d, s, r) for d, s, r in sessions if d in valid_dates]:
                            if name not in used_on_day[date] and assigned_counts[name] < duties_needed:
                                current_assigned = len([n for n in used_on_day[date] if date in duty_data[n] and session.upper() in duty_data[n][date]])
                                if current_assigned < required:
                                    duty_data[name][date] = duty_data[name].get(date, []) + [session.upper()]
                                    used_on_day[date].add(name)
                                    assigned_counts[name] += 1
                                    assigned += 1
                                    assigned_slots[name] = slot
                                    logging.info(f"Assigned {orig_name} (AP) to {date} {session} (Slot {slot}, quota fulfillment, relaxed 70:30)")
                                    if assigned == duties_needed:
                                        break
                        if assigned == duties_needed:
                            break
                    if assigned < duties_needed:
                        duty_quota_violations.append(f"{orig_name} (AP) assigned {assigned}/{duties_needed} duties (preferred slot: {pref_slot})")

            # Fill with GLs
            for date, session, required in sessions:
                current_assigned = len([n for n in used_on_day[date] if date in duty_data[n] and session.upper() in duty_data[n][date]])
                perm_assigned = len([n for n in used_on_day[date] if date in duty_data[n] and session.upper() in duty_data[n][date] and merged_df[merged_df['name'] == n]['designation'].iloc[0] in ['PROF', 'ASP', 'AP']])
                perm_target = math.ceil(required * designation_caps['perm_ratio'])
                if perm_assigned < perm_target:
                    ratio_violations.append(f"{date} {session}: Permanent staff ratio violation ({perm_assigned}/{perm_target} needed)")
                remaining_needed = required - current_assigned
                if remaining_needed > 0:
                    available_gls = sorted(
                        [n for n in merged_df[merged_df['designation'] == 'GL']['name'] if n not in used_on_day[date]],
                        key=lambda x: assigned_counts[x]
                    )
                    for name in available_gls[:remaining_needed]:
                        orig_name = merged_df[merged_df['name'] == name]['original_name'].iloc[0]
                        duty_data[name][date] = duty_data[name].get(date, []) + [session.upper()]
                        used_on_day[date].add(name)
                        assigned_counts[name] += 1
                        logging.info(f"Assigned {orig_name} (GL) to {date} {session} (Slot {'Slot 1' if date in slot_dates['Slot 1'] else 'Slot 2'})")

            # Validate assignments
            for date, session, required in sessions:
                current_assigned = len([n for n in used_on_day[date] if date in duty_data[n] and session.upper() in duty_data[n][date]])
                if current_assigned != required:
                    success = False
                    logging.warning(f"Failed to assign all duties for {date} {session}: {current_assigned}/{required}")
                    duty_quota_violations.append(f"Failed to assign all duties for {date} {session}: {current_assigned}/{required}")
                    break

            if success:
                # Validate no split assignments for ASPs and APs
                split_violations = []
                for name in merged_df[merged_df['designation'].isin(['ASP', 'AP'])]['name']:
                    assigned_dates = list(duty_data[name].keys())
                    slots_assigned = set('Slot 1' if d in slot_dates['Slot 1'] else 'Slot 2' for d in assigned_dates if d in slot_dates['Slot 1'] or d in slot_dates['Slot 2'])
                    if len(slots_assigned) > 1:
                        orig_name = merged_df[merged_df['name'] == name]['original_name'].iloc[0]
                        split_violations.append(f"{orig_name} assigned to multiple slots: {slots_assigned}")
                        logging.error(f"{orig_name} assigned to multiple slots: {slots_assigned}")
                if split_violations:
                    success = False
                    logging.warning(f"Split slot violations with config {designation_caps['name']}: {split_violations}")
                    duty_quota_violations.extend(split_violations)
                else:
                    # Check if all APs have their full quota
                    ap_violations = [f"{merged_df[merged_df['name'] == n]['original_name'].iloc[0]} (AP) assigned {assigned_counts[n]}/{designation_caps['AP']} duties" 
                                     for n in merged_df[merged_df['designation'] == 'AP']['name'] if assigned_counts[n] < designation_caps['AP']]
                    if not ap_violations:
                        break
                    else:
                        duty_quota_violations.extend(ap_violations)
                        success = False
            else:
                logging.info(f"Config {designation_caps['name']} failed, trying next fallback")

        if not success:
            logging.warning("All configurations failed, using best effort with last config")
            duty_quota_violations.append("All configurations failed, using best effort with last config")

        # Generate output Excel
        output_rows = []
        for name in merged_df['name']:
            desig = merged_df[merged_df['name'] == name]['designation'].iloc[0]
            orig_name = merged_df[merged_df['name'] == name]['original_name'].iloc[0]
            row = {'Name': orig_name, 'Designation': desig}
            for d in all_dates:
                sessions = duty_data.get(name, {}).get(d, [])
                row[d] = ' '.join(sessions) if sessions else ''
            output_rows.append(row)

        output_df = pd.DataFrame(output_rows)
        output_df.insert(0, 'Name', output_df.pop('Name'))
        output_df.insert(1, 'Designation', output_df.pop('Designation'))
        output_df.to_excel(output_path, index=False)

        # Summarize
        total_duties = sum(assigned_counts.values())
        prof_duties = sum(assigned_counts[name] for name in merged_df[merged_df['designation'] == 'PROF']['name'])
        asp_duties = sum(assigned_counts[name] for name in merged_df[merged_df['designation'] == 'ASP']['name'])
        ap_duties = sum(assigned_counts[name] for name in merged_df[merged_df['designation'] == 'AP']['name'])
        gl_duties = sum(assigned_counts[name] for name in merged_df[merged_df['designation'] == 'GL']['name'])
        assignment_summary = f"Final chart (Config: {designation_caps['name']}): {prof_duties} PROF, {asp_duties} ASP, {ap_duties} AP, {gl_duties} GL, Total: {total_duties}\n"
        return assignment_summary, ratio_violations, duty_quota_violations, None, merged_df[['name', 'original_name']].set_index('name')['original_name'].to_dict()

    except Exception as e:
        logging.error(f"Failed to generate chart: {str(e)}")
        messagebox.showerror("Error", f"Failed to generate chart: {str(e)}\nCheck duty_chart_app.log for details.")
        return None, None, None, None, {}

# ----------------------------- GUI Class -----------------------------
class DutyChartApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Duty Chart Generator")
        self.input_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.setup_widgets()

    def setup_widgets(self):
        frm = ttk.Frame(self.root, padding=10)
        frm.grid()
        ttk.Label(frm, text="Input Excel File:").grid(row=0, column=0, sticky="w")
        ttk.Entry(frm, textvariable=self.input_path, width=50).grid(row=0, column=1)
        ttk.Button(frm, text="Browse", command=self.browse_input).grid(row=0, column=2)
        ttk.Label(frm, text="Output Excel File:").grid(row=1, column=0, sticky="w")
        ttk.Entry(frm, textvariable=self.output_path, width=50).grid(row=1, column=1)
        ttk.Button(frm, text="Save As", command=self.browse_output).grid(row=1, column=2)
        self.slot1_start = DateEntry(frm)
        self.slot1_end = DateEntry(frm)
        self.slot2_start = DateEntry(frm)
        self.slot2_end = DateEntry(frm)
        ttk.Label(frm, text="Slot 1 Start:").grid(row=2, column=0, sticky="w")
        self.slot1_start.grid(row=2, column=1, sticky="w")
        ttk.Label(frm, text="End:").grid(row=2, column=2, sticky="w")
        self.slot1_end.grid(row=2, column=3, sticky="w")
        ttk.Label(frm, text="Slot 2 Start:").grid(row=3, column=0, sticky="w")
        self.slot2_start.grid(row=3, column=1, sticky="w")
        ttk.Label(frm, text="End:").grid(row=3, column=2, sticky="w")
        self.slot2_end.grid(row=3, column=3, sticky="w")
        ttk.Button(frm, text="Generate Duty Chart", command=self.run).grid(row=4, column=1, pady=10)
        self.summary_box = tk.Text(frm, width=80, height=20)
        self.summary_box.grid(row=5, column=0, columnspan=4)

    def browse_input(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            self.input_path.set(file_path)

    def browse_output(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            self.output_path.set(file_path)

    def run(self):
        try:
            slot1 = (self.slot1_start.get_date(), self.slot1_end.get_date())
            slot2 = (self.slot2_start.get_date(), self.slot2_end.get_date())
            assignment_summary, ratio_violations, duty_quota_violations, _, _ = generate_duty_chart(self.input_path.get(), self.output_path.get(), slot1, slot2)
            self.summary_box.delete("1.0", tk.END)
            self.summary_box.insert(tk.END, "Assignment Summary:\n")
            self.summary_box.insert(tk.END, f"{assignment_summary if assignment_summary else 'All sessions assigned'}\n")
            self.summary_box.insert(tk.END, "\n70:30 Rule Violations:\n")
            if ratio_violations:
                for violation in ratio_violations:
                    self.summary_box.insert(tk.END, f"{violation}\n")
            else:
                self.summary_box.insert(tk.END, "No 70:30 rule violations.\n")
            self.summary_box.insert(tk.END, "\nDuty Quota Violations:\n")
            if duty_quota_violations:
                for violation in duty_quota_violations:
                    self.summary_box.insert(tk.END, f"{violation}\n")
            else:
                self.summary_box.insert(tk.END, "No duty quota violations.\n")
            messagebox.showinfo("Success", "Duty chart generated successfully! Check the output file and log for details.")
        except Exception as e:
            logging.error(f"GUI run error: {str(e)}")
            messagebox.showerror("Unexpected Error", f"Error: {str(e)}\nCheck duty_chart_app.log for details.")

if __name__ == "__main__":
    try:
        root = tk.Tk()
        app = DutyChartApp(root)
        root.mainloop()
    except Exception as e:
        logging.error(f"Main execution error: {str(e)}")
        messagebox.showerror("Error", f"Error starting application: {str(e)}")