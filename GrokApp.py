import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkcalendar import DateEntry
from datetime import timedelta
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
    # Remove titles and extra spaces, preserve order
    cleaned = re.sub(r"^(Dr\.?|Prof\.?|Mr\.?|Mrs\.?|Ms\.?)\s*", "", str(name).strip(), flags=re.IGNORECASE)
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    # Replace periods in initials with spaces for consistency
    cleaned = re.sub(r"\.(?=\w)", " ", cleaned)
    # Split into components and lowercase
    parts = cleaned.split()
    parts = [part.lower() for part in parts if part]
    return ' '.join(parts)

def normalize_designation(desig):
    if pd.isna(desig):
        return ""
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
        # Normalize both names
        staff_norm = normalize_name(staff_name)
        pref_norm = normalize_name(pref_name)
        staff_parts = staff_norm.split()
        pref_parts = pref_norm.split()
        
        # Check for common significant parts (length > 3)
        common_parts = set(staff_parts) & set(pref_parts)
        if any(len(part) > 3 for part in common_parts):
            logging.info(f"Fuzzy matched {pref_name} to {staff_name} based on significant part: {common_parts}")
            return True
        
        # Calculate similarity score on normalized names
        score = SequenceMatcher(None, staff_norm, pref_norm).ratio()
        logging.debug(f"Fuzzy match attempt: {staff_name} vs {pref_name}, normalized: {staff_norm} vs {pref_norm}, score: {score:.3f}")
        
        # Additional check with raw names (no spaces or punctuation)
        staff_raw = re.sub(r"[.\s]+", "", staff_name.lower())
        pref_raw = re.sub(r"[.\s]+", "", pref_name.lower())
        raw_score = SequenceMatcher(None, staff_raw, pref_raw).ratio()
        logging.debug(f"Raw match attempt: {staff_raw} vs {pref_raw}, score: {raw_score:.3f}")
        
        if score >= threshold or raw_score >= 0.9:
            logging.info(f"Fuzzy matched {pref_name} to {staff_name} with score {max(score, raw_score):.3f}")
            return True
        
        logging.debug(f"No match for {staff_name} vs {pref_name}, scores: normalized={score:.3f}, raw={raw_score:.3f}, threshold={threshold}")
        return False
    except Exception as e:
        logging.error(f"Fuzzy match failed for {staff_name} vs {pref_name}: {e}")
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
    except Exception as e:
        logging.error(f"Failed to parse date {val}: {e}")
        return None

def parse_timestamp(ts):
    try:
        parsed = pd.to_datetime(ts, errors='coerce')
        if pd.isna(parsed):
            logging.error(f"Invalid timestamp: {ts}")
            return pd.NaT
        return parsed.tz_localize(None)
    except Exception as e:
        logging.error(f"Failed to parse timestamp {ts}: {e}")
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

        # Find sheets with exact or near-exact names
        sheet_map = {
            'session strength': ['Session Strength', 'Sessionwise Strength'],
            'staff list': ['Staff List', 'Staff Details'],
            'slot preference': ['Slot Preference']
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

        # Process data
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

        # Fuzzy match names
        staff_names = set(staff_df['name'])
        pref_names = set(pref_df['name'])
        unmatched_staff = staff_names - pref_names
        unmatched_pref = pref_names - staff_names
        fuzzy_matches = {}
        for staff_name in unmatched_staff:
            staff_original = staff_df[staff_df['name'] == staff_name]['original_name'].iloc[0]
            best_score = 0
            best_match = None
            for pref_name in unmatched_pref:
                pref_original = pref_df[pref_df['name'] == pref_name]['original_name'].iloc[0]
                if fuzzy_match_name(staff_original, pref_original):
                    score = SequenceMatcher(None, normalize_name(staff_original), normalize_name(pref_original)).ratio()
                    raw_score = SequenceMatcher(None, re.sub(r"[.\s]+", "", staff_original.lower()), re.sub(r"[.\s]+", "", pref_original.lower())).ratio()
                    final_score = max(score, raw_score)
                    if final_score > best_score:
                        best_score = final_score
                        best_match = pref_name
            if best_match:
                fuzzy_matches[best_match] = staff_name
                logging.info(f"Fuzzy matched {pref_df[pref_df['name'] == best_match]['original_name'].iloc[0]} (Preference) to {staff_original} (Staff) with score {best_score:.3f}")
        
        if fuzzy_matches:
            pref_df['name'] = pref_df['name'].replace(fuzzy_matches)
            unmatched_pref = pref_names - staff_names - set(fuzzy_matches.keys())
            unmatched_staff = staff_names - pref_names - set(fuzzy_matches.values())

        if unmatched_staff:
            unmatched_original = [staff_df[staff_df['name'] == n]['original_name'].iloc[0] for n in unmatched_staff]
            logging.info(f"Staff names not in Slot Preference (defaulting to Any): {unmatched_original}")
        if unmatched_pref:
            unmatched_original = [pref_df[pref_df['name'] == n]['original_name'].iloc[0] for n in unmatched_pref]
            logging.warning(f"Preference names not in Staff List (ignored): {unmatched_original}")
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

        # Validate slot dates with overlap handling
        all_dates = sorted(session_df['date'].unique())
        slot_dates = {'Slot 1': set(), 'Slot 2': set()}
        for d in all_dates:
            if slot1_range[0] <= d <= slot1_range[1]:
                slot_dates['Slot 1'].add(d)
            elif slot2_range[0] <= d <= slot2_range[1] and d not in slot_dates['Slot 1']:
                slot_dates['Slot 2'].add(d)
        logging.info(f"Slot 1 dates: {sorted(slot_dates['Slot 1'])}, Slot 2 dates: {sorted(slot_dates['Slot 2'])}")

        # Calculate duties
        slot1_duties = sum(math.ceil(row[s] / 30) for _, row in session_df.iterrows() for s in ['fn', 'an'] if row['date'] in slot_dates['Slot 1'])
        slot2_duties = sum(math.ceil(row[s] / 30) for _, row in session_df.iterrows() for s in ['fn', 'an'] if row['date'] in slot_dates['Slot 2'])
        logging.info(f"Slot 1 needs {slot1_duties} duties, Slot 2 needs {slot2_duties} duties")

        # Define sessions
        sessions = [(row['date'], s, math.ceil(row[s] / 30)) for _, row in session_df.iterrows() for s in ['fn', 'an'] if math.ceil(row[s] / 30) > 0]
        sessions.sort(key=lambda x: (x[0], -x[2]))  # Sort by date, then by required duties (descending)

        # Initialize tracking
        assigned_counts = {name: 0 for name in merged_df['name']}
        used_on_day = {d: set() for d in all_dates}
        duty_data = {name: {} for name in merged_df['name']}
        assigned_slots = {name: None for name in merged_df['name']}

        # Duty caps configurations with fallback
        duty_configs = [
            {'Professor': 1, 'Assoc. Professor': 3, 'Asst. Professor': 6, 'A.P(Contract)': float('inf'), 'perm_ratio': 0.7, 'gl_ratio': 0.3, 'name': '1:3:6 (70:30)'},
            {'Professor': 1, 'Assoc. Professor': 4, 'Asst. Professor': 8, 'A.P(Contract)': float('inf'), 'perm_ratio': 0.7, 'gl_ratio': 0.3, 'name': '1:4:8 (70:30)'},
            {'Professor': 1, 'Assoc. Professor': 3, 'Asst. Professor': 6, 'A.P(Contract)': float('inf'), 'perm_ratio': 0.6, 'gl_ratio': 0.4, 'name': '1:3:6 (60:40)'},
            {'Professor': 1, 'Assoc. Professor': 4, 'Asst. Professor': 8, 'A.P(Contract)': float('inf'), 'perm_ratio': 0.8, 'gl_ratio': 0.2, 'name': '1:4:8 (80:20)'}
        ]

        ratio_violations = []
        duty_quota_violations = []
        slot_preference_violations = []

        # Assignment with fallback logic
        for config_idx, designation_caps in enumerate(duty_configs):
            assigned_counts = {name: 0 for name in merged_df['name']}
            used_on_day = {d: set() for d in all_dates}
            duty_data = {name: {} for name in merged_df['name']}
            assigned_slots = {name: None for name in merged_df['name']}
            success = True

            # Assign permanent staff (Professor, Assoc. Professor, Asst. Professor)
            for desig in ['Professor', 'Assoc. Professor', 'Asst. Professor']:
                candidates = merged_df[merged_df['designation'] == desig][['name', 'original_name', 'preferred slot', 'timestamp']]
                if desig == 'Asst. Professor':
                    candidates = sorted(candidates.to_dict('records'), key=lambda x: x['timestamp'] if not pd.isna(x['timestamp']) else pd.Timestamp.max)
                else:
                    candidates = candidates.to_dict('records')

                for candidate in candidates:
                    name = candidate['name']
                    orig_name = candidate['original_name']
                    pref_slot = candidate['preferred slot'] if candidate['preferred slot'] in ['Slot 1', 'Slot 2'] else 'Any'
                    valid_slots = [pref_slot] if pref_slot in ['Slot 1', 'Slot 2'] else ['Slot 1', 'Slot 2']
                    duties_needed = designation_caps[desig]
                    assigned = 0

                    for slot in valid_slots:
                        valid_dates = sorted(slot_dates[slot])
                        for date, session, required in [(d, s, r) for d, s, r in sessions if d in valid_dates]:
                            if name not in used_on_day[date] and assigned_counts[name] < duties_needed:
                                current_slot = 'Slot 1' if date in slot_dates['Slot 1'] else 'Slot 2'
                                if assigned_slots[name] is not None and assigned_slots[name] != current_slot:
                                    continue  # Prevent slot splitting
                                current_assigned = len([n for n in used_on_day[date] if date in duty_data[n] and session.upper() in duty_data[n][date]])
                                perm_needed = math.ceil(required * designation_caps['perm_ratio'])
                                perm_assigned = len([n for n in used_on_day[date] if date in duty_data[n] and session.upper() in duty_data[n][date] and merged_df[merged_df['name'] == n]['designation'].iloc[0] in ['Professor', 'Assoc. Professor', 'Asst. Professor']])
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

            # Fill with A.P(Contract)
            for date, session, required in sessions:
                current_assigned = len([n for n in used_on_day[date] if date in duty_data[n] and session.upper() in duty_data[n][date]])
                perm_assigned = len([n for n in used_on_day[date] if date in duty_data[n] and session.upper() in duty_data[n][date] and merged_df[merged_df['name'] == n]['designation'].iloc[0] in ['Professor', 'Assoc. Professor', 'Asst. Professor']])
                perm_target = math.ceil(required * designation_caps['perm_ratio'])
                if perm_assigned < perm_target:
                    ratio_violations.append(f"{date} {session}: Permanent staff ratio violation ({perm_assigned}/{perm_target} needed)")
                remaining_needed = required - current_assigned
                if remaining_needed > 0:
                    available_gls = sorted(
                        [n for n in merged_df[merged_df['designation'] == 'A.P(Contract)']['name'] if n not in used_on_day[date]],
                        key=lambda x: assigned_counts[x]
                    )
                    for name in available_gls[:remaining_needed]:
                        orig_name = merged_df[merged_df['name'] == name]['original_name'].iloc[0]
                        duty_data[name][date] = duty_data[name].get(date, []) + [session.upper()]
                        used_on_day[date].add(name)
                        assigned_counts[name] += 1
                        logging.info(f"Assigned {orig_name} (A.P(Contract)) to {date} {session} (Slot {'Slot 1' if date in slot_dates['Slot 1'] else 'Slot 2'})")

            # Validate assignments
            for date, session, required in sessions:
                current_assigned = len([n for n in used_on_day[date] if date in duty_data[n] and session.upper() in duty_data[n][date]])
                if current_assigned != required:
                    success = False
                    logging.warning(f"Failed to assign all duties for {date} {session}: {current_assigned}/{required}")
                    duty_quota_violations.append(f"Failed to assign all duties for {date} {session}: {current_assigned}/{required}")
                    break

            if success:
                break

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
            total_duties = sum(len(duty_data[name].get(d, [])) for d in all_dates)
            assigned_slot = assigned_slots[name] if assigned_slots[name] is not None else "Mixed" if desig == 'A.P(Contract)' else "None"
            row['Total Duties'] = total_duties
            row['Assigned Slot'] = assigned_slot
            output_rows.append(row)

        output_df = pd.DataFrame(output_rows)
        output_df.insert(0, 'Name', output_df.pop('Name'))
        output_df.insert(1, 'Designation', output_df.pop('Designation'))
        output_df.to_excel(output_path, index=False)

        # Summarize
        total_duties = sum(assigned_counts.values())
        prof_duties = sum(assigned_counts[name] for name in merged_df[merged_df['designation'] == 'Professor']['name'])
        asp_duties = sum(assigned_counts[name] for name in merged_df[merged_df['designation'] == 'Assoc. Professor']['name'])
        ap_duties = sum(assigned_counts[name] for name in merged_df[merged_df['designation'] == 'Asst. Professor']['name'])
        gl_duties = sum(assigned_counts[name] for name in merged_df[merged_df['designation'] == 'A.P(Contract)']['name'])
        assignment_summary = f"Final chart (Config: {designation_caps['name']}): {prof_duties} Professor, {asp_duties} Assoc. Professor, {ap_duties} Asst. Professor, {gl_duties} A.P(Contract), Total: {total_duties}\n"
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