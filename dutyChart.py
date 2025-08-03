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
    cleaned = re.sub(r"^(Dr\.?|Prof\.?|Mr\.?|Mrs\.?|Ms\.?)\s*", "", str(name).strip(), flags=re.IGNORECASE)
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    cleaned = re.sub(r"\.(?=\w)", " ", cleaned)
    parts = [part.lower() for part in cleaned.split() if part]
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
        staff_norm = normalize_name(staff_name)
        pref_norm = normalize_name(pref_name)
        staff_parts = staff_norm.split()
        pref_parts = pref_norm.split()
        
        common_parts = set(staff_parts) & set(pref_parts)
        if any(len(part) > 3 for part in common_parts):
            logging.info(f"Fuzzy matched {pref_name} to {staff_name} based on significant part: {common_parts}")
            return True
        
        score = SequenceMatcher(None, staff_norm, pref_norm).ratio()
        logging.debug(f"Fuzzy match attempt: {staff_name} vs {pref_name}, normalized: {staff_norm} vs {pref_norm}, score: {score:.3f}")
        
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
def generate_duty_chart(input_path, output_path, slot1_range, slot2_range, ratio_choice):
    try:
        input_path = input_path.strip('"').strip()
        output_path = output_path.strip('"').strip()
        logging.info(f"Input path: {input_path}, Output path: {output_path}, Ratio: {ratio_choice}")

        if not os.path.exists(input_path):
            logging.error(f"Input file not found: {input_path}")
            messagebox.showerror("Error", f"Input file not found: {input_path}")
            return None, None, None, None, {}

        xls = pd.ExcelFile(input_path)
        sheets = {s.strip().lower().replace('\n', '').replace('\r', ''): s for s in xls.sheet_names}

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

        session_df = pd.read_excel(xls, found_sheets['session strength'])
        staff_df = pd.read_excel(xls, found_sheets['staff list'])
        pref_df = pd.read_excel(xls, found_sheets['slot preference'])

        session_df.columns = [c.strip().lower().replace('\n', '').replace('\r', '') for c in session_df.columns]
        staff_df.columns = [c.strip().lower().replace('\n', '').replace('\r', '') for c in staff_df.columns]
        pref_df.columns = [c.strip().lower().replace('\n', '').replace('\r', '') for c in pref_df.columns]

        session_cols = {
            'date': find_column(session_df, ['date']),
            'fn': find_column(session_df, ['fn', 'forenoon', 'morning']),
            'an': find_column(session_df, ['an', 'afternoon'])
        }
        staff_cols = {
            'name': find_column(staff_df, ['name of the faculty', 'name', 'faculty']),
            'designation': find_column(staff_df, ['designation', 'design', 'desig']),
            'department': find_column(staff_df, ['department', 'dept'])
        }
        pref_cols = {
            'timestamp': find_column(pref_df, ['timestamp']),
            'name': find_column(pref_df, ['name of the faculty', 'name', 'faculty']),
            'preferred slot': find_column(pref_df, ['preferred slot', 'slot', 'preferredslot'])
        }

        missing_cols = []
        for df_name, cols in [('Session Strength', session_cols), ('Staff List', staff_cols), ('Slot Preference', pref_cols)]:
            for col_name, col in cols.items():
                if col is None:
                    missing_cols.append(f"{col_name} in {df_name}")
        if missing_cols:
            logging.error(f"Missing columns: {', '.join(missing_cols)}")
            messagebox.showerror("Error", f"Missing columns: {', '.join(missing_cols)}")
            return None, None, None, None, {}

        session_df = session_df.rename(columns={session_cols['date']: 'date', session_cols['fn']: 'fn', session_cols['an']: 'an'})
        staff_df = staff_df.rename(columns={staff_cols['name']: 'name', staff_cols['designation']: 'designation', staff_cols['department']: 'department'})
        pref_df = pref_df.rename(columns={pref_cols['timestamp']: 'timestamp', pref_cols['name']: 'name', 
                                         pref_cols['preferred slot']: 'preferred slot'})

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

        if unmatched_staff or unmatched_pref:
            unmatched_staff_orig = [staff_df[staff_df['name'] == n]['original_name'].iloc[0] for n in unmatched_staff]
            unmatched_pref_orig = [pref_df[pref_df['name'] == n]['original_name'].iloc[0] for n in unmatched_pref]
            logging.warning(f"Unmatched staff (defaulting to Any): {unmatched_staff_orig}, Unmatched preferences (ignored): {unmatched_pref_orig}")
            pref_df = pref_df[pref_df['name'].isin(staff_names)]

        merged_df = pd.merge(staff_df[['name', 'original_name', 'designation', 'department']], 
                            pref_df[['name', 'original_name', 'timestamp', 'preferred slot']], 
                            on='name', how='left')
        merged_df['preferred slot'] = merged_df['preferred slot'].fillna('Any')
        merged_df['original_name_x'] = merged_df['original_name_x'].fillna(merged_df['name'])
        merged_df = merged_df.rename(columns={'original_name_x': 'original_name'}).drop(columns=['original_name_y'], errors='ignore')
        merged_df = merged_df.drop_duplicates(subset=['name'])

        all_dates = sorted(session_df['date'].unique())
        slot_dates = {'Slot 1': set(), 'Slot 2': set()}
        for d in all_dates:
            if slot1_range[0] <= d <= slot1_range[1]:
                slot_dates['Slot 1'].add(d)
            elif slot2_range[0] <= d <= slot2_range[1]:
                slot_dates['Slot 2'].add(d)
        logging.info(f"Slot 1 dates: {sorted(slot_dates['Slot 1'])}, Slot 2 dates: {sorted(slot_dates['Slot 2'])}")

        slot1_duties = sum(math.ceil(row[s] / 30) for _, row in session_df.iterrows() for s in ['fn', 'an'] if row['date'] in slot_dates['Slot 1'])
        slot2_duties = sum(math.ceil(row[s] / 30) for _, row in session_df.iterrows() for s in ['fn', 'an'] if row['date'] in slot_dates['Slot 2'])
        logging.info(f"Slot 1 needs {slot1_duties} duties, Slot 2 needs {slot2_duties} duties")

        sessions = [(row['date'], s, math.ceil(row[s] / 30)) for _, row in session_df.iterrows() for s in ['fn', 'an'] if math.ceil(row[s] / 30) > 0]
        sessions.sort(key=lambda x: (x[0], -x[2]))

        assigned_counts = {name: 0 for name in merged_df['name']}
        used_on_day = {d: set() for d in all_dates}
        duty_data = {name: {} for name in merged_df['name']}
        assigned_slots = {name: None for name in merged_df['name']}

        # Parse ratio_choice and set duty caps
        ratio_map = {
            '1:3:6': {'Professor': 1, 'Assoc. Professor': 3, 'Asst. Professor': 6, 'A.P(Contract)': float('inf'), 'perm_ratio': 0.7, 'gl_ratio': 0.3},
            '1:3:7': {'Professor': 1, 'Assoc. Professor': 3, 'Asst. Professor': 7, 'A.P(Contract)': float('inf'), 'perm_ratio': 0.7, 'gl_ratio': 0.3},
            '1:4:8': {'Professor': 1, 'Assoc. Professor': 4, 'Asst. Professor': 8, 'A.P(Contract)': float('inf'), 'perm_ratio': 0.7, 'gl_ratio': 0.3}
        }
        if ratio_choice not in ratio_map:
            logging.error(f"Invalid ratio choice: {ratio_choice}")
            messagebox.showerror("Error", f"Invalid ratio choice: {ratio_choice}")
            return None, None, None, None, {}
        designation_caps = ratio_map[ratio_choice]

        ratio_violations = []
        duty_quota_violations = []
        slot_preference_violations = []

        # Assign duties for permanent staff (Professor, Assoc. Professor, Asst. Professor)
        for desig in ['Professor', 'Assoc. Professor', 'Asst. Professor']:
            candidates = merged_df[merged_df['designation'] == desig][['name', 'original_name', 'preferred slot', 'timestamp']]
            candidates = sorted(candidates.to_dict('records'), key=lambda x: x['timestamp'] if not pd.isna(x['timestamp']) else pd.Timestamp.max) if desig == 'Asst. Professor' else candidates.to_dict('records')

            for candidate in candidates:
                name = candidate['name']
                orig_name = candidate['original_name']
                pref_slot = candidate['preferred slot'] if candidate['preferred slot'] in ['Slot 1', 'Slot 2'] else 'Any'
                valid_slots = [pref_slot] if pref_slot in ['Slot 1', 'Slot 2'] else ['Slot 1', 'Slot 2']
                duties_needed = designation_caps[desig]
                assigned = 0

                # Prioritize slot with higher demand for APs with 'Any' preference
                if pref_slot == 'Any' and desig == 'Asst. Professor':
                    valid_slots = ['Slot 1', 'Slot 2'] if slot1_duties >= slot2_duties else ['Slot 2', 'Slot 1']

                for slot in valid_slots:
                    valid_dates = sorted(slot_dates[slot])
                    for date, session, required in [(d, s, r) for d, s, r in sessions if d in valid_dates]:
                        if name not in used_on_day[date] and assigned_counts[name] < duties_needed:
                            current_slot = 'Slot 1' if date in slot_dates['Slot 1'] else 'Slot 2'
                            if assigned_slots[name] is not None and assigned_slots[name] != current_slot:
                                continue
                            current_assigned = len([n for n in used_on_day[date] if date in duty_data[n] and session.upper() in duty_data[n][date]])
                            perm_needed = math.ceil(required * designation_caps['perm_ratio'])
                            perm_assigned = len([n for n in used_on_day[date] if date in duty_data[n] and session.upper() in duty_data[n][date] and merged_df[merged_df['name'] == n]['designation'].iloc[0] in ['Professor', 'Assoc. Professor', 'Asst. Professor']])
                            if perm_assigned < perm_needed and current_assigned < required:
                                duty_data[name][date] = duty_data[name].get(date, []) + [session.upper()]
                                used_on_day[date].add(name)
                                assigned_counts[name] += 1
                                assigned += 1
                                assigned_slots[name] = current_slot
                                if pref_slot != 'Any' and pref_slot != current_slot:
                                    slot_preference_violations.append(f"{orig_name} ({desig}) assigned to {current_slot} but preferred {pref_slot}")
                                logging.info(f"Assigned {orig_name} ({desig}) to {date} {session} (Slot {current_slot})")
                                if assigned == duties_needed:
                                    break
                    if assigned == duties_needed:
                        break
                if assigned < duties_needed:
                    duty_quota_violations.append(f"{orig_name} ({desig}) assigned {assigned}/{duties_needed} duties (preferred slot: {pref_slot})")

        # Assign remaining duties to Guest Lecturers (A.P(Contract))
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
                    pref_slot = merged_df[merged_df['name'] == name]['preferred slot'].iloc[0]
                    current_slot = 'Slot 1' if date in slot_dates['Slot 1'] else 'Slot 2'
                    if assigned_slots[name] is not None and assigned_slots[name] != current_slot:
                        continue
                    duty_data[name][date] = duty_data[name].get(date, []) + [session.upper()]
                    used_on_day[date].add(name)
                    assigned_counts[name] += 1
                    assigned_slots[name] = current_slot
                    if pref_slot != 'Any' and pref_slot != current_slot:
                        slot_preference_violations.append(f"{orig_name} (A.P(Contract)) assigned to {current_slot} but preferred {pref_slot}")
                    logging.info(f"Assigned {orig_name} (A.P(Contract)) to {date} {session} (Slot {current_slot})")

        # Check for unassigned duties
        for date, session, required in sessions:
            current_assigned = len([n for n in used_on_day[date] if date in duty_data[n] and session.upper() in duty_data[n][date]])
            if current_assigned != required:
                logging.warning(f"Failed to assign all duties for {date} {session}: {current_assigned}/{required}")
                duty_quota_violations.append(f"Failed to assign all duties for {date} {session}: {current_assigned}/{required}")

        # Generate output Excel with additional analysis rows
        output_rows = []
        for name in merged_df['name']:
            desig = merged_df[merged_df['name'] == name]['designation'].iloc[0]
            orig_name = merged_df[merged_df['name'] == name]['original_name'].iloc[0]
            dept = merged_df[merged_df['name'] == name]['department'].iloc[0]
            row = {'Name': orig_name, 'Designation': desig, 'Department': dept}
            for d in all_dates:
                sessions = duty_data.get(name, {}).get(d, [])
                row[d] = ' '.join(sessions) if sessions else ''
            total_duties = sum(len(duty_data[name].get(d, [])) for d in all_dates)
            assigned_slot = assigned_slots[name] if assigned_slots[name] is not None else "None"
            row['Total Duties'] = total_duties
            row['Assigned Slot'] = assigned_slot
            output_rows.append(row)

        output_df = pd.DataFrame(output_rows)
        output_df.insert(0, 'Name', output_df.pop('Name'))
        output_df.insert(1, 'Designation', output_df.pop('Designation'))
        output_df.insert(2, 'Department', output_df.pop('Department'))
        output_df.insert(3, 'Total Duties', output_df.pop('Total Duties'))
        output_df.insert(4, 'Assigned Slot', output_df.pop('Assigned Slot'))

        # Add analysis rows
        fn_duties_row = {'Name': 'Total FN Duties', 'Designation': '', 'Department': '', 'Total Duties': '', 'Assigned Slot': ''}
        an_duties_row = {'Name': 'Total AN Duties', 'Designation': '', 'Department': '', 'Total Duties': '', 'Assigned Slot': ''}
        fn_perm_pct_row = {'Name': 'FN Permanent %', 'Designation': '', 'Department': '', 'Total Duties': '', 'Assigned Slot': ''}
        an_perm_pct_row = {'Name': 'AN Permanent %', 'Designation': '', 'Department': '', 'Total Duties': '', 'Assigned Slot': ''}

        for date in all_dates:
            fn_count = sum(1 for name in merged_df['name'] if date in duty_data.get(name, {}) and 'FN' in duty_data[name][date])
            an_count = sum(1 for name in merged_df['name'] if date in duty_data.get(name, {}) and 'AN' in duty_data[name][date])
            fn_perm_count = sum(1 for name in merged_df[merged_df['designation'].isin(['Professor', 'Assoc. Professor', 'Asst. Professor'])]['name'] if date in duty_data.get(name, {}) and 'FN' in duty_data[name][date])
            an_perm_count = sum(1 for name in merged_df[merged_df['designation'].isin(['Professor', 'Assoc. Professor', 'Asst. Professor'])]['name'] if date in duty_data.get(name, {}) and 'AN' in duty_data[name][date])
            fn_duties_row[date] = fn_count
            an_duties_row[date] = an_count
            fn_perm_pct_row[date] = f"{(fn_perm_count / fn_count * 100):.1f}%" if fn_count > 0 else '0.0%'
            an_perm_pct_row[date] = f"{(an_perm_count / an_count * 100):.1f}%" if an_count > 0 else '0.0%'

        output_df = pd.concat([output_df, pd.DataFrame([fn_duties_row, an_duties_row, fn_perm_pct_row, an_perm_pct_row])], ignore_index=True)
        output_df.to_excel(output_path, index=False)

        total_duties = sum(assigned_counts.values())
        prof_duties = sum(assigned_counts[name] for name in merged_df[merged_df['designation'] == 'Professor']['name'])
        asp_duties = sum(assigned_counts[name] for name in merged_df[merged_df['designation'] == 'Assoc. Professor']['name'])
        ap_duties = sum(assigned_counts[name] for name in merged_df[merged_df['designation'] == 'Asst. Professor']['name'])
        gl_duties = sum(assigned_counts[name] for name in merged_df[merged_df['designation'] == 'A.P(Contract)']['name'])
        assignment_summary = f"Final chart (Ratio: {ratio_choice}): {prof_duties} Professor, {asp_duties} Assoc. Professor, {ap_duties} Asst. Professor, {gl_duties} A.P(Contract), Total: {total_duties}\n"
        return assignment_summary, ratio_violations, duty_quota_violations, slot_preference_violations, merged_df[['name', 'original_name']].set_index('name')['original_name'].to_dict()

    except Exception as e:
        logging.error(f"Failed to generate chart: {str(e)}")
        messagebox.showerror("Error", f"Failed to generate chart: {str(e)}\nCheck duty_chart_app.log for details.")
        return None, None, None, None, {}
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

