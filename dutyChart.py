import tkinter as tk
from tkinter import filedialog, messagebox, ttk, font
import pandas as pd
import numpy as np
import re
import os
import math
import logging
from difflib import SequenceMatcher
from datetime import timedelta, datetime

try:
    from tkcalendar import DateEntry
except ImportError:
    logging.error("tkcalendar module not found. Please install it using 'pip install tkcalendar'.")
    messagebox.showerror("Module Error", "The 'tkcalendar' module is required. Please install it by running 'pip install tkcalendar' in your terminal and restart the application.")
    raise ImportError("tkcalendar module is required. Install it with 'pip install tkcalendar'.")

logging.basicConfig(filename='duty_chart_app.log', level=logging.DEBUG,
                    format='%(asctime)s - %(levelname)s - %(message)s')

# ----------------------------- Helper Functions -----------------------------

def normalize_name(name):
    if pd.isna(name):
        return ""
    cleaned = re.sub(r"^(Dr\.?|Prof\.?|Mr\.?|Mrs\.?|Ms\.?)\s*", "",
                     str(name).strip(), flags=re.IGNORECASE)
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    cleaned = re.sub(r"\.(?=\w)", " ", cleaned)
    parts = [part.lower() for part in cleaned.split() if part]
    return ' '.join(parts)

def fuzzy_match_name(staff_name, pref_name, threshold=0.65):
    try:
        staff_norm = normalize_name(staff_name)
        pref_norm = normalize_name(pref_name)
        staff_parts = staff_norm.split()
        pref_parts = pref_norm.split()
        common_parts = set(staff_parts) & set(pref_parts)
        if any(len(part) > 3 for part in common_parts):
            return True
        score = SequenceMatcher(None, staff_norm, pref_norm).ratio()
        staff_raw = re.sub(r"[.\s]+", "", staff_name.lower())
        pref_raw = re.sub(r"[.\s]+", "", pref_name.lower())
        raw_score = SequenceMatcher(None, staff_raw, pref_raw).ratio()
        return score >= threshold or raw_score >= 0.9
    except Exception as e:
        logging.error(f"Fuzzy match failed: {e}")
        return False

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

def find_column(df, keywords):
    for col in df.columns:
        col_clean = col.strip().lower().replace('\n', '').replace('\r', '')
        if any(keyword.lower() in col_clean for keyword in keywords):
            return col
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
        val_str = str(val).strip()
        if ',' in val_str or ' ' in val_str:
            val_str = val_str.split(',')[0].split()[0].strip()
        for fmt in ['%d/%m/%Y', '%Y-%m-%d', '%m/%d/%Y', '%d-%m-%Y']:
            try:
                return datetime.strptime(val_str, fmt).date()
            except ValueError:
                continue
        return pd.to_datetime(val_str, errors='coerce').date() if val_str else None
    except Exception as e:
        logging.error(f"Failed to parse date {val}: {e}")
        return None

def parse_timestamp(ts):
    try:
        parsed = pd.to_datetime(ts, errors='coerce')
        if pd.isna(parsed):
            return pd.NaT
        return parsed.tz_localize(None)
    except Exception as e:
        logging.error(f"Failed to parse timestamp {ts}: {e}")
        return pd.NaT

def can_assign(name, date, used_on_day):
    if date not in used_on_day:
        return True
    return name not in used_on_day[date]

def generate_duty_chart(app, input_path, output_path, slot1_range, slot2_range, ratio_choice):
    try:
        app.update_progress(10)
        input_path = input_path.strip('"').strip()
        output_path = output_path.strip('"').strip()

        if not os.path.exists(input_path):
            raise FileNotFoundError(f"Input file not found: {input_path}")
        if not os.access(input_path, os.R_OK):
            raise PermissionError(f"No read permission for input file: {input_path}")
        if os.path.exists(output_path) and not os.access(os.path.dirname(output_path) or '.', os.W_OK):
            raise PermissionError(f"No write permission for output directory: {os.path.dirname(output_path) or '.'}")

        logging.info(f"Input path: {input_path}, Output path: {output_path}")

        xls = pd.ExcelFile(input_path)
        app.update_progress(20)
        sheets = {s.strip().lower().replace('\n', ' ').replace('\r', ''): s.strip() for s in xls.sheet_names}
        sheet_map = {
            'session strength': ['session strength', 'sessionwise strength'],
            'staff list': ['staff list', 'staff details'],
            'slot preference': ['slot preference']
        }
        found_sheets = {}
        for key, variations in sheet_map.items():
            for variation in variations:
                v = variation.strip().lower()
                if v in sheets:
                    found_sheets[key] = sheets[v]
                    break
            if key not in found_sheets:
                logging.error(f"Missing sheet: {key}. Found sheets: {', '.join(xls.sheet_names)}")
                messagebox.showerror("Error", f"Missing sheet '{key}'. Found sheets: {', '.join(xls.sheet_names)}")
                return None, None, None, None, {}

        session_df = pd.read_excel(xls, found_sheets['session strength'])
        staff_df = pd.read_excel(xls, found_sheets['staff list'])
        pref_df = pd.read_excel(xls, found_sheets['slot preference'])

        app.update_progress(30)

        if session_df.empty or staff_df.empty or pref_df.empty:
            logging.error("One or more input sheets are empty.")
            messagebox.showerror("Error", "One or more input sheets are empty.")
            return None, None, None, None, {}

        # Normalize columns to lower-case and clean whitespace/newlines
        session_df.columns = [c.strip().lower().replace('\n', '').replace('\r', '') for c in session_df.columns]
        staff_df.columns = [c.strip().lower().replace('\n', '').replace('\r', '') for c in staff_df.columns]
        pref_df.columns = [c.strip().lower().replace('\r', '') for c in pref_df.columns]

        session_cols = {'date': find_column(session_df, ['date']),
                        'fn': find_column(session_df, ['fn', 'forenoon']),
                        'an': find_column(session_df, ['an', 'afternoon'])}
        staff_cols = {'name': find_column(staff_df, ['name']),
                      'designation': find_column(staff_df, ['designation']),
                      'department': find_column(staff_df, ['department', 'dept'])}
        pref_cols = {'timestamp': find_column(pref_df, ['timestamp']),
                     'name': find_column(pref_df, ['name']),
                     'preferred slot': find_column(pref_df, ['preferred slot', 'slot'])}

        missing_cols = []
        for df_name, cols in [('Session Strength', session_cols), ('Staff List', staff_cols), ('Slot Preference', pref_cols)]:
            for col_name, col in cols.items():
                if col is None:
                    missing_cols.append(f"{col_name} in {df_name}")

        if missing_cols:
            logging.error(f"Missing columns: {', '.join(missing_cols)}")
            messagebox.showerror("Error", f"Missing columns: {', '.join(missing_cols)}")
            return None, None, None, None, {}

        session_df = session_df.rename(columns={session_cols['date']: 'date',
                                                session_cols['fn']: 'fn',
                                                session_cols['an']: 'an'})
        staff_df = staff_df.rename(columns={staff_cols['name']: 'name',
                                            staff_cols['designation']: 'designation',
                                            staff_cols['department']: 'department'})
        pref_df = pref_df.rename(columns={pref_cols['timestamp']: 'timestamp',
                                          pref_cols['name']: 'name',
                                          pref_cols['preferred slot']: 'preferred slot'})

        session_df['date'] = session_df['date'].apply(safe_parse_date)
        session_df = session_df.dropna(subset=['date'])
        session_df['fn'] = pd.to_numeric(session_df['fn'], errors='coerce').fillna(0)
        session_df['an'] = pd.to_numeric(session_df['an'], errors='coerce').fillna(0)

        app.update_progress(40)

        if session_df['fn'].isna().all() or session_df['an'].isna().all():
            logging.error("Forenoon or Afternoon columns contain no valid data.")
            messagebox.showerror("Error", "Forenoon or Afternoon columns contain no valid data.")
            return None, None, None, None, {}

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

        merged_df = pd.merge(staff_df[['name', 'original_name', 'designation', 'department']],
                             pref_df[['name', 'original_name', 'timestamp', 'preferred slot']],
                             on='name', how='left')

        merged_df['preferred slot'] = merged_df['preferred slot'].fillna('Any')
        merged_df['original_name_x'] = merged_df['original_name_x'].fillna(merged_df['name'])

        merged_df = merged_df.rename(columns={'original_name_x': 'original_name'}).drop(
            columns=['original_name_y'], errors='ignore')

        merged_df = merged_df.drop_duplicates(subset=['name'])

        app.update_progress(50)

        all_dates = sorted(session_df['date'].unique())

        slot_dates = {'Slot 1': set(), 'Slot 2': set()}

        for d in all_dates:
            if slot1_range[0] <= d <= slot1_range[1]:
                slot_dates['Slot 1'].add(d)
            elif slot2_range[0] <= d <= slot2_range[1]:
                slot_dates['Slot 2'].add(d)

        sessions = []
        total_duties = 0
        for _, row in session_df.iterrows():
            date = row['date']
            if date in slot_dates['Slot 1'].union(slot_dates['Slot 2']):
                fn_duties = math.ceil(row['fn'] / 30)
                an_duties = math.ceil(row['an'] / 30)
                total_duties += fn_duties + an_duties
                if fn_duties > 0:
                    sessions.append((date, 'FN', fn_duties))
                if an_duties > 0:
                    sessions.append((date, 'AN', an_duties))

        sessions.sort(key=lambda x: (x[0], -x[2]))

        app.update_progress(60)

        num_professors = len(merged_df[merged_df['designation'] == 'Professor'])

        if total_duties < num_professors:
            logging.warning(f"Insufficient duties ({total_duties}) for {num_professors} Professors. Some Professors may not be assigned.")
            messagebox.showwarning("Warning", f"Insufficient duties ({total_duties}) to assign 1 duty to each of {num_professors} Professors.")

        ratio_map = {
            '1:3:6': {'Professor': 1, 'Assoc. Professor': 3, 'Asst. Professor': 6},
            '1:3:7': {'Professor': 1, 'Assoc. Professor': 3, 'Asst. Professor': 7},
            '1:4:8': {'Professor': 1, 'Assoc. Professor': 4, 'Asst. Professor': 8},
        }

        if ratio_choice not in ratio_map:
            messagebox.showerror("Error", f"Invalid duty ratio selected: {ratio_choice}")
            return None, None, None, None, {}

        designation_caps = ratio_map[ratio_choice]
        designation_caps['A.P(Contract)'] = float('inf')

        assigned_counts = {name: 0 for name in merged_df['name']}
        duty_data = {}
        used_on_day = {}
        assigned_slots = {}
        ratio_violations = []
        duty_quota_violations = []
        slot_preference_violations = []
        seventy_thirty_violations = []

        perm_staff = merged_df[merged_df['designation'].isin(['Professor', 'Assoc. Professor', 'Asst. Professor'])]
        gl_staff = merged_df[merged_df['designation'] == 'A.P(Contract)']

        # First pass: Assign 1 duty to each Professor
        professors = perm_staff[perm_staff['designation'] == 'Professor'].sample(frac=1).reset_index(drop=True)

        for _, person in professors.iterrows():
            name = person['name']
            pref_slot = person['preferred slot'] if person['preferred slot'] in ['Slot 1', 'Slot 2'] else 'Any'
            assigned = False
            for i, (date, session, remaining) in enumerate(sessions):
                if remaining <= 0:
                    continue
                slot = 'Slot 1' if date in slot_dates['Slot 1'] else 'Slot 2'
                if pref_slot != 'Any' and pref_slot != slot:
                    continue
                if not can_assign(name, date, used_on_day):
                    continue
                duty_data.setdefault(name, {})
                duty_data[name].setdefault(date, []).append(session)
                used_on_day.setdefault(date, set()).add(name)
                assigned_counts[name] = assigned_counts.get(name, 0) + 1
                assigned_slots[name] = slot
                sessions[i] = (date, session, remaining - 1)
                assigned = True
                break

            if not assigned and pref_slot != 'Any':
                for i, (date, session, remaining) in enumerate(sessions):
                    if remaining <= 0:
                        continue
                    if not can_assign(name, date, used_on_day):
                        continue
                    duty_data.setdefault(name, {})
                    duty_data[name].setdefault(date, []).append(session)
                    used_on_day.setdefault(date, set()).add(name)
                    assigned_counts[name] = assigned_counts.get(name, 0) + 1
                    slot = 'Slot 1' if date in slot_dates['Slot 1'] else 'Slot 2'

                    assigned_slots[name] = slot

                    sessions[i] = (date, session, remaining - 1)
                    slot_preference_violations.append(
                        f"{person['original_name']} (Professor) assigned to {slot} but preferred {pref_slot}")
                    assigned = True
                    break

            if not assigned:
                duty_quota_violations.append(f"{person['original_name']} (Professor) assigned 0/1 duties")

        app.update_progress(70)

        # Second pass: Assign remaining duties with 70:30 enforcement
        for date, session, total_needed in sessions:
            if total_needed <= 0:
                continue
            used_on_day.setdefault(date, set())
            permanents_needed = math.ceil(total_needed * 0.7)  # 70% permanent staff
            gl_needed = total_needed - permanents_needed  # 30% guest lecturers
            original_permanents_needed = permanents_needed
            slot_of_date = 'Slot 1' if date in slot_dates['Slot 1'] else 'Slot 2'
            assigned_this_session = set()

            def assign_perm_group(designation):
                nonlocal permanents_needed
                candidates = perm_staff[(perm_staff['designation'] == designation) & (
                    (perm_staff['preferred slot'] == slot_of_date) | (perm_staff['preferred slot'] == 'Any'))]
                if designation == 'Asst. Professor':
                    candidates = candidates.sort_values('timestamp', na_position='last')
                else:
                    candidates = candidates.sample(frac=1).reset_index(drop=True)
                for _, person in candidates.iterrows():
                    if permanents_needed <= 0:
                        break
                    name = person['name']
                    cap = designation_caps[designation]
                    already_assigned = assigned_counts.get(name, 0)
                    if already_assigned >= cap or name in used_on_day[date]:
                        continue
                    if can_assign(name, date, used_on_day):
                        duty_data.setdefault(name, {})
                        duty_data[name].setdefault(date, []).append(session)
                        used_on_day[date].add(name)
                        assigned_counts[name] = already_assigned + 1
                        assigned_slots[name] = slot_of_date
                        permanents_needed -= 1
                        assigned_this_session.add(name)

            total_perm_assigned = 0
            for desig, cap in [('Assoc. Professor', designation_caps['Assoc. Professor']),
                               ('Asst. Professor', designation_caps['Asst. Professor'])]:
                candidates = perm_staff[perm_staff['designation'] == desig]
                if desig == 'Asst. Professor':
                    candidates = candidates.sort_values('timestamp', na_position='last')
                else:
                    candidates = candidates.sample(frac=1).reset_index(drop=True)
                target = math.ceil(permanents_needed * (cap / (designation_caps['Professor'] + designation_caps['Assoc. Professor'] + designation_caps['Asst. Professor'])))
                assigned = 0
                for _, person in candidates.iterrows():
                    if assigned >= target or permanents_needed <= 0:
                        break
                    name = person['name']
                    already_assigned = assigned_counts.get(name, 0)
                    if already_assigned >= cap or name in used_on_day[date]:
                        continue
                    if can_assign(name, date, used_on_day):
                        duty_data.setdefault(name, {})
                        duty_data[name].setdefault(date, []).append(session)
                        used_on_day[date].add(name)
                        assigned_counts[name] = already_assigned + 1
                        assigned_slots[name] = slot_of_date
                        permanents_needed -= 1
                        assigned_this_session.add(name)
                        assigned += 1
                        total_perm_assigned += 1

            # Relax slot preferences if needed
            if permanents_needed > 0:
                for desig in ['Assoc. Professor', 'Asst. Professor']:
                    candidates = perm_staff[perm_staff['designation'] == desig]
                    if desig == 'Asst. Professor':
                        candidates = candidates.sort_values('timestamp', na_position='last')
                    else:
                        candidates = candidates.sample(frac=1).reset_index(drop=True)
                    for _, person in candidates.iterrows():
                        if permanents_needed <= 0:
                            break
                        name = person['name']
                        cap = designation_caps[desig]
                        already_assigned = assigned_counts.get(name, 0)
                        if already_assigned >= cap or name in used_on_day[date]:
                            continue
                        if can_assign(name, date, used_on_day):
                            duty_data.setdefault(name, {})
                            duty_data[name].setdefault(date, []).append(session)
                            used_on_day[date].add(name)
                            assigned_counts[name] = already_assigned + 1
                            assigned_slots[name] = slot_of_date
                            permanents_needed -= 1
                            assigned_this_session.add(name)
                            if person['preferred slot'] not in ['Any', slot_of_date]:
                                slot_preference_violations.append(
                                    f"{person['original_name']} ({desig}) assigned to {slot_of_date} but preferred {person['preferred slot']}")
                            total_perm_assigned += 1

            if permanents_needed > 0:
                ratio_violations.append(
                    f"{date} {session}: Could not assign required permanents ({original_permanents_needed - permanents_needed}/{original_permanents_needed}) due to caps or availability")

            # Assign guest lecturers (30%)
            assigned_gl_this_session = 0
            available_g_ls = gl_staff[~gl_staff['name'].isin(used_on_day[date])].copy()
            available_g_ls['assigned_count'] = available_g_ls['name'].map(lambda n: assigned_counts.get(n, 0))
            available_g_ls = available_g_ls.sort_values('assigned_count')
            for _, person in available_g_ls.iterrows():
                if gl_needed <= 0:
                    break
                name = person['name']
                if can_assign(name, date, used_on_day):
                    duty_data.setdefault(name, {})
                    duty_data[name].setdefault(date, []).append(session)
                    used_on_day[date].add(name)
                    assigned_counts[name] = assigned_counts.get(name, 0) + 1
                    # FIXED HERE: get() only takes 2 args, use conditional assignment:
                    if name in assigned_slots:
                        current_slot = assigned_slots[name]
                    else:
                        current_slot = slot_of_date
                    # If assigned to different slot before, mark as 'Mixed'
                    if current_slot != slot_of_date:
                        assigned_slots[name] = 'Mixed'
                    else:
                        assigned_slots[name] = slot_of_date
                    gl_needed -= 1
                    assigned_gl_this_session += 1

            # Validate 70:30 ratio
            total_assigned = total_perm_assigned + assigned_gl_this_session
            if total_assigned > 0 and (total_perm_assigned / total_assigned < 0.7 - 0.05 or total_perm_assigned / total_assigned > 0.7 + 0.05):
                seventy_thirty_violations.append(
                    f"{date} {session}: 70:30 ratio violated. Assigned {total_perm_assigned}/{total_assigned} permanents ({(total_perm_assigned / total_assigned) * 100:.1f}%)")

            if gl_needed > 0:
                ratio_violations.append(
                    f"{date} {session}: Could not assign all GL duties. Remaining: {gl_needed}")

        app.update_progress(80)

        for _, person in perm_staff.iterrows():
            name = person['name']
            desig = person['designation']
            cap = designation_caps.get(desig, 0)
            assigned = assigned_counts.get(name, 0)
            if assigned < cap:
                duty_quota_violations.append(
                    f"{person['original_name']} ({desig}) assigned {assigned}/{cap} duties")

        output_rows = []
        for _, row in merged_df.iterrows():
            name = row['name']
            desig = row['designation']
            dept = row['department'] if 'department' in merged_df.columns else ''
            orig_name = row['original_name']
            assigned_slot = assigned_slots.get(name)
            assigned_slot_str = assigned_slot if assigned_slot in ['Slot 1', 'Slot 2'] else ('Mixed' if assigned_slot == 'Mixed' else 'None')
            total_duties = assigned_counts.get(name, 0)
            row_data = {'Name': orig_name, 'Designation': desig,
                        'Department': dept, 'Total Duties': total_duties,
                        'Assigned Slot': assigned_slot_str}
            for d in all_dates:
                sessions_assigned = duty_data.get(name, {}).get(d, [])
                row_data[d] = ' '.join(sorted(sessions_assigned)) if sessions_assigned else ''
            output_rows.append(row_data)

        app.update_progress(90)

        output_df = pd.DataFrame(output_rows)
        col_order = ['Name', 'Designation', 'Department',
                     'Total Duties', 'Assigned Slot'] + all_dates
        output_df = output_df[col_order]
        output_df.to_excel(output_path, index=False)

        logging.info(f"Duty chart saved to {output_path}")

        app.update_progress(100)

        total = sum(assigned_counts.values())
        prof_count = sum(assigned_counts[name]
                         for name in merged_df[merged_df['designation'] == 'Professor']['name'])
        asp_count = sum(assigned_counts[name]
                        for name in merged_df[merged_df['designation'] == 'Assoc. Professor']['name'])
        ap_count = sum(assigned_counts[name]
                       for name in merged_df[merged_df['designation'] == 'Asst. Professor']['name'])
        gl_count = sum(assigned_counts[name]
                       for name in merged_df[merged_df['designation'] == 'A.P(Contract)']['name'])

        summary = f"Final chart (Ratio: {ratio_choice}): {prof_count} Professor, {asp_count} Assoc. Professor, {ap_count} Asst. Professor, {gl_count} A.P(Contract), Total duties assigned: {total}"

        return summary, ratio_violations, duty_quota_violations, slot_preference_violations, seventy_thirty_violations, merged_df[['name', 'original_name']].set_index('name')['original_name'].to_dict()

    except (FileNotFoundError, PermissionError) as e:
        logging.error(f"File access error: {str(e)}")
        messagebox.showerror("Error", f"File access error: {str(e)}\nCheck duty_chart_app.log for details.")
        return None, None, None, None, None, {}
    except Exception as e:
        logging.error(f"Failed to generate chart: {str(e)}")
        messagebox.showerror("Error", f"Failed to generate chart: {str(e)}\nCheck duty_chart_app.log for details.")
        return None, None, None, None, None, {}

# ----------------------------- GUI Class -----------------------------

class DutyChartApp:

    def __init__(self, root):
        self.root = root
        self.root.title("Duty Chart Generator")
        self.root.geometry("1000x700")
        self.root.minsize(800, 600)

        self.input_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.ratio_choice = tk.StringVar(value="1:3:6")
        self.theme = "dark"

        # Initialize font for tooltip and widgets
        try:
            self.tooltip_font = ("Helvetica", 8) if "Helvetica" in font.families() else ("Arial", 8)
        except Exception as e:
            logging.warning(f"Font detection failed: {e}. Falling back to sans-serif.")
            self.tooltip_font = ("sans-serif", 8)

        self.setup_styles()
        self.setup_widgets()
        self.root.bind("<Configure>", self.on_resize)

    def setup_styles(self):
        self.style = ttk.Style()
        try:
            self.style.theme_use("clam")
        except tk.TclError as e:
            logging.warning(f"Failed to set 'clam' theme: {e}. Falling back to default theme.")
            self.style.theme_use("default")
        self.apply_dark_theme()

    def apply_dark_theme(self):
        self.theme = "dark"
        font_ = self.tooltip_font
        font_bold = (font_[0], font_[1], "bold")
        font_title = (font_[0], 20, "bold")
        font_label = (font_[0], 12, "bold")

        self.style.configure("TFrame", background="#1e2937")
        self.style.configure("TLabel", font=font_, background="#1e2937", foreground="#ffffff")
        self.style.configure("Title.TLabel", font=font_title, background="#1e2937", foreground="#60a5fa")
        self.style.configure("TEntry", font=font_, fieldbackground="#374151", foreground="#000000")
        self.style.configure("TButton", font=font_bold, background="#4b5563", foreground="#000000")
        self.style.map("TButton",
                       background=[("active", "#374151"), ("disabled", "#6b7280")],
                       foreground=[("active", "#000000"), ("disabled", "#000000")])
        self.style.configure("TRadiobutton", font=font_, background="#1e2937", foreground="#ffffff")
        self.style.map("TRadiobutton",
                       foreground=[("selected", "#ffffff"), ("active", "#ffffff")])
        self.style.configure("TProgressbar", thickness=40, troughcolor="#374151", background="#14b8a6")
        self.style.configure("Vertical.TScrollbar", background="#374151", troughcolor="#1e2937")
        self.style.map("Vertical.TScrollbar", background=[("active", "#4b5563")])

        # Update widgets if exist
        if hasattr(self, 'summary_box'):
            self.summary_box.configure(bg="#374151", fg="#ffffff")
            self.summary_box.tag_configure("header", font=font_label, foreground="#60a5fa")
            self.summary_box.tag_configure("loading", font=font_ + ("italic",), foreground="#ffffff")

        # Update input/output date range background etc.
        for w in [getattr(self, n, None) for n in ['slot1_start', 'slot1_end', 'slot2_start', 'slot2_end']]:
            if w is not None:
                w.configure(background="#374151", foreground="#000000")

        self.update_button_styles()

    def apply_light_theme(self):
        self.theme = "light"
        font_ = self.tooltip_font
        font_bold = (font_[0], font_[1], "bold")
        font_title = (font_[0], 20, "bold")
        font_label = (font_[0], 12, "bold")

        self.style.configure("TFrame", background="#f3f4f6")
        self.style.configure("TLabel", font=font_, background="#f3f4f6", foreground="#000000")
        self.style.configure("Title.TLabel", font=font_title, background="#f3f4f6", foreground="#60a5fa")
        self.style.configure("TEntry", font=font_, fieldbackground="#e5e7eb", foreground="#000000")
        self.style.configure("TButton", font=font_bold, background="#d1d5db", foreground="#000000")
        self.style.map("TButton",
                       background=[("active", "#6b7280"), ("disabled", "#e5e7eb")],
                       foreground=[("active", "#000000"), ("disabled", "#000000")])
        self.style.configure("TRadiobutton", font=font_, background="#f3f4f6", foreground="#000000")
        self.style.map("TRadiobutton",
                       foreground=[("selected", "#000000"), ("active", "#000000")])
        self.style.configure("TProgressbar", thickness=40, troughcolor="#d1d5db", background="#14b8a6")
        self.style.configure("Vertical.TScrollbar", background="#d1d5db", troughcolor="#f3f4f6")
        self.style.map("Vertical.TScrollbar", background=[("active", "#9ca3af")])

        # Update widgets if exist
        if hasattr(self, 'summary_box'):
            self.summary_box.configure(bg="#e5e7eb", fg="#000000")
            self.summary_box.tag_configure("header", font=font_label, foreground="#60a5fa")
            self.summary_box.tag_configure("loading", font=font_ + ("italic",), foreground="#000000")

        # Update input/output date range background etc.
        for w in [getattr(self, n, None) for n in ['slot1_start', 'slot1_end', 'slot2_start', 'slot2_end']]:
            if w is not None:
                w.configure(background="#e5e7eb", foreground="#000000")

        self.update_button_styles()

    def update_button_styles(self):
        if not hasattr(self, 'generate_button') or not hasattr(self, 'clear_button'):
            return
        font_bold = (self.tooltip_font[0], self.tooltip_font[1], "bold")
        try:
            if self.theme == "dark":
                self.generate_button.configure(
                    bg="#4b5563", fg="#ffffff", activebackground="#374151",
                    borderwidth=2, relief="groove", font=font_bold, padx=10, pady=5)
                self.clear_button.configure(
                    bg="#4b5563", fg="#ffffff", activebackground="#374151",
                    borderwidth=2, relief="groove", font=font_bold, padx=10, pady=5)
            else:
                self.generate_button.configure(
                    bg="#d1d5db", fg="#000000", activebackground="#6b7280",
                    borderwidth=2, relief="groove", font=font_bold, padx=10, pady=5)
                self.clear_button.configure(
                    bg="#d1d5db", fg="#000000", activebackground="#6b7280",
                    borderwidth=2, relief="groove", font=font_bold, padx=10, pady=5)
        except tk.TclError as e:
            logging.error(f"Button style configuration error: {str(e)}")
            # fallback
            if self.theme == "dark":
                self.generate_button.configure(bg="#4b5563", fg="#ffffff", activebackground="#374151")
                self.clear_button.configure(bg="#4b5563", fg="#ffffff", activebackground="#374151")
            else:
                self.generate_button.configure(bg="#d1d5db", fg="#000000", activebackground="#6b7280")
                self.clear_button.configure(bg="#d1d5db", fg="#000000", activebackground="#6b7280")

    def button_hover_enter(self, event):
        try:
            if self.theme == "dark":
                event.widget.configure(bg="#6b7280")
            else:
                event.widget.configure(bg="#9ca3af")
        except tk.TclError as e:
            logging.error(f"Button hover enter error: {str(e)}")

    def button_hover_leave(self, event):
        try:
            if self.theme == "dark":
                event.widget.configure(bg="#4b5563")
            else:
                event.widget.configure(bg="#d1d5db")
        except tk.TclError as e:
            logging.error(f"Button hover leave error: {str(e)}")

    def toggle_theme(self):
        if self.theme == "dark":
            self.apply_light_theme()
        else:
            self.apply_dark_theme()
        for frame in [self.input_frame, self.output_frame, self.date_frame, self.ratio_frame, self.summary_frame]:
            frame.configure(bg="#f3f4f6" if self.theme == "light" else "#1e2937")
        for inner_frame in [self.input_inner, self.output_inner, self.date_inner, self.ratio_inner, self.summary_inner]:
            inner_frame.configure(bg="#f3f4f6" if self.theme == "light" else "#1e2937")
        for label in [self.input_label, self.output_label, self.date_label, self.ratio_label, self.summary_label]:
            label.configure(bg="#f3f4f6" if self.theme == "light" else "#1e2937")
        self.root.update_idletasks()

    def clear_fields(self):
        self.input_path.set("")
        self.output_path.set("")
        self.slot1_start.set_date(datetime.today())
        self.slot1_end.set_date(datetime.today())
        self.slot2_start.set_date(datetime.today())
        self.slot2_end.set_date(datetime.today())
        self.ratio_choice.set("1:3:6")
        self.summary_box.delete("1.0", tk.END)
        self.progress['value'] = 0

    def setup_widgets(self):
        try:
            font_label = (self.tooltip_font[0], 12, "bold")
            main_frame = ttk.Frame(self.root, padding=20, style="TFrame")
            main_frame.pack(fill="both", expand=True)

            header_frame = ttk.Frame(main_frame, style="TFrame")
            header_frame.pack(fill="x")
            ttk.Label(header_frame, text="Duty Chart Generator", style="Title.TLabel").pack(side="left", pady=10)

            theme_button = ttk.Button(header_frame, text="Toggle Theme", command=self.toggle_theme, style="TButton")
            theme_button.pack(side="right", padx=10, pady=5)
            theme_button.bind("<Enter>", lambda e: self.show_tooltip(theme_button, "Switch between dark and light themes"))
            theme_button.bind("<Leave>", lambda e: self.hide_tooltip())

            # Input section
            self.input_frame = tk.Frame(main_frame, bg="#1e2937" if self.theme == "dark" else "#f3f4f6", borderwidth=0,
                                        relief="flat")
            self.input_frame.pack(fill="x", pady=5)

            self.input_label = tk.Label(self.input_frame, text="Input File", font=font_label,
                                        bg="#1e2937" if self.theme == "dark" else "#f3f4f6", fg="#60a5fa")
            self.input_label.pack(anchor="w", padx=10, pady=(0, 5))

            self.input_inner = tk.Frame(self.input_frame, bg="#1e2937" if self.theme == "dark" else "#f3f4f6")
            self.input_inner.pack(fill="x", padx=10)

            input_entry = ttk.Entry(self.input_inner, textvariable=self.input_path, style="TEntry")
            input_entry.pack(side="left", fill="x", expand=True, padx=5)
            input_entry.bind("<Enter>", lambda e: self.show_tooltip(input_entry, "Select the Excel file with session, staff, and preference data"))
            input_entry.bind("<Leave>", lambda e: self.hide_tooltip())

            input_button = ttk.Button(self.input_inner, text="Browse", command=self.browse_input, style="TButton")
            input_button.pack(side="left", padx=5)
            input_button.bind("<Enter>", lambda e: self.show_tooltip(input_button, "Open file explorer to choose input Excel file"))
            input_button.bind("<Leave>", lambda e: self.hide_tooltip())

            # Output section
            self.output_frame = tk.Frame(main_frame, bg="#1e2937" if self.theme == "dark" else "#f3f4f6", borderwidth=0,
                                         relief="flat")
            self.output_frame.pack(fill="x", pady=5)

            self.output_label = tk.Label(self.output_frame, text="Output File", font=font_label,
                                         bg="#1e2937" if self.theme == "dark" else "#f3f4f6", fg="#60a5fa")
            self.output_label.pack(anchor="w", padx=10, pady=(0, 5))

            self.output_inner = tk.Frame(self.output_frame, bg="#1e2937" if self.theme == "dark" else "#f3f4f6")
            self.output_inner.pack(fill="x", padx=10)

            output_entry = ttk.Entry(self.output_inner, textvariable=self.output_path, style="TEntry")
            output_entry.pack(side="left", fill="x", expand=True, padx=5)
            output_entry.bind("<Enter>", lambda e: self.show_tooltip(output_entry, "Specify the path for the output Excel file"))
            output_entry.bind("<Leave>", lambda e: self.hide_tooltip())

            output_button = ttk.Button(self.output_inner, text="Save As", command=self.browse_output, style="TButton")
            output_button.pack(side="left", padx=5)
            output_button.bind("<Enter>", lambda e: self.show_tooltip(output_button, "Open file explorer to set output file location"))
            output_button.bind("<Leave>", lambda e: self.hide_tooltip())

            # Date ranges
            self.date_frame = tk.Frame(main_frame, bg="#1e2937" if self.theme == "dark" else "#f3f4f6", borderwidth=0,
                                       relief="flat")
            self.date_frame.pack(fill="x", pady=5)

            self.date_label = tk.Label(self.date_frame, text="Date Ranges", font=font_label,
                                       bg="#1e2937" if self.theme == "dark" else "#f3f4f6", fg="#60a5fa")
            self.date_label.pack(anchor="w", padx=10, pady=(0, 5))

            self.date_inner = tk.Frame(self.date_frame, bg="#1e2937" if self.theme == "dark" else "#f3f4f6")
            self.date_inner.pack(fill="x", padx=10)

            # Slot 1 range
            slot1_frame = ttk.Frame(self.date_inner, style="TFrame")
            slot1_frame.pack(fill="x", pady=5)
            ttk.Label(slot1_frame, text="Slot 1 Start:", style="TLabel").pack(side="left")

            self.slot1_start = DateEntry(slot1_frame, date_pattern="dd/mm/yyyy", background="#374151", foreground="#000000")
            self.slot1_start.pack(side="left", padx=5)
            self.slot1_start.bind("<Enter>", lambda e: self.show_tooltip(self.slot1_start, "Select start date for Slot 1"))
            self.slot1_start.bind("<Leave>", lambda e: self.hide_tooltip())

            ttk.Label(slot1_frame, text="End:", style="TLabel").pack(side="left")

            self.slot1_end = DateEntry(slot1_frame, date_pattern="dd/mm/yyyy", background="#374151", foreground="#000000")
            self.slot1_end.pack(side="left", padx=5)
            self.slot1_end.bind("<Enter>", lambda e: self.show_tooltip(self.slot1_end, "Select end date for Slot 1"))
            self.slot1_end.bind("<Leave>", lambda e: self.hide_tooltip())

            # Slot 2 range
            slot2_frame = ttk.Frame(self.date_inner, style="TFrame")
            slot2_frame.pack(fill="x", pady=5)
            ttk.Label(slot2_frame, text="Slot 2 Start:", style="TLabel").pack(side="left")

            self.slot2_start = DateEntry(slot2_frame, date_pattern="dd/mm/yyyy", background="#374151", foreground="#000000")
            self.slot2_start.pack(side="left", padx=5)
            self.slot2_start.bind("<Enter>", lambda e: self.show_tooltip(self.slot2_start, "Select start date for Slot 2"))
            self.slot2_start.bind("<Leave>", lambda e: self.hide_tooltip())

            ttk.Label(slot2_frame, text="End:", style="TLabel").pack(side="left")

            self.slot2_end = DateEntry(slot2_frame, date_pattern="dd/mm/yyyy", background="#374151", foreground="#000000")
            self.slot2_end.pack(side="left", padx=5)
            self.slot2_end.bind("<Enter>", lambda e: self.show_tooltip(self.slot2_end, "Select end date for Slot 2"))
            self.slot2_end.bind("<Leave>", lambda e: self.hide_tooltip())

            # Duty Ratio options
            self.ratio_frame = tk.Frame(main_frame, bg="#1e2937" if self.theme == "dark" else "#f3f4f6", borderwidth=0,
                                       relief="flat")
            self.ratio_frame.pack(fill="x", pady=5)

            self.ratio_label = tk.Label(self.ratio_frame, text="Duty Ratio (Prof:ASP:AP)", font=font_label,
                                        bg="#1e2937" if self.theme == "dark" else "#f3f4f6", fg="#60a5fa")
            self.ratio_label.pack(anchor="w", padx=10, pady=(0, 5))

            self.ratio_inner = tk.Frame(self.ratio_frame, bg="#1e2937" if self.theme == "dark" else "#f3f4f6")
            self.ratio_inner.pack(fill="x", padx=10)

            for i, val in enumerate(["1:3:6", "1:3:7", "1:4:8"]):
                radio = ttk.Radiobutton(self.ratio_inner, text=val, value=val, variable=self.ratio_choice, style="TRadiobutton")
                radio.pack(side="left", padx=10)
                radio.bind("<Enter>", lambda e, v=val: self.show_tooltip(radio, f"Select duty ratio {v} for Professors, Associate Professors, and Assistant Professors"))
                radio.bind("<Leave>", lambda e: self.hide_tooltip())

            button_frame = ttk.Frame(main_frame, style="TFrame")
            button_frame.pack(pady=20)

            self.generate_button = tk.Button(button_frame, text="Generate Duty Chart", command=self.run)
            self.generate_button.pack(side="left", padx=5)
            self.generate_button.bind("<Enter>", self.button_hover_enter)
            self.generate_button.bind("<Leave>", self.button_hover_leave)
            self.generate_button.bind("<Enter>", lambda e: self.show_tooltip(self.generate_button, "Generate the duty chart based on input data and settings"), add="+")
            self.generate_button.bind("<Leave>", lambda e: self.hide_tooltip(), add="+")

            self.clear_button = tk.Button(button_frame, text="Clear", command=self.clear_fields)
            self.clear_button.pack(side="left", padx=5)
            self.clear_button.bind("<Enter>", self.button_hover_enter)
            self.clear_button.bind("<Leave>", self.button_hover_leave)
            self.clear_button.bind("<Enter>", lambda e: self.show_tooltip(self.clear_button, "Clear all input fields and summary"), add="+")
            self.clear_button.bind("<Leave>", lambda e: self.hide_tooltip(), add="+")

            self.update_button_styles()

            self.progress = ttk.Progressbar(main_frame, mode="determinate", maximum=100, style="TProgressbar")
            self.progress.pack(fill="x", pady=5, padx=50)
            self.progress.bind("<Enter>", lambda e: self.show_tooltip(self.progress, "Shows progress while generating the duty chart"))
            self.progress.bind("<Leave>", lambda e: self.hide_tooltip())

            self.summary_frame = tk.Frame(main_frame, bg="#1e2937" if self.theme == "dark" else "#f3f4f6", borderwidth=0,
                                          relief="flat")
            self.summary_frame.pack(fill="both", expand=True, pady=5)

            self.summary_label = tk.Label(self.summary_frame, text="Summary", font=font_label,
                                          bg="#1e2937" if self.theme == "dark" else "#f3f4f6", fg="#60a5fa")
            self.summary_label.pack(anchor="w", padx=10, pady=(0, 5))

            self.summary_inner = tk.Frame(self.summary_frame, bg="#1e2937" if self.theme == "dark" else "#f3f4f6")
            self.summary_inner.pack(fill="both", expand=True, padx=10)

            self.summary_box = tk.Text(self.summary_inner, height=15, wrap="word", font=self.tooltip_font,
                                       bg="#374151", fg="#ffffff", relief="flat", borderwidth=0)
            scrollbar = ttk.Scrollbar(self.summary_inner, orient="vertical", command=self.summary_box.yview, style="Vertical.TScrollbar")
            self.summary_box.config(yscrollcommand=scrollbar.set)
            scrollbar.pack(side="right", fill="y")
            self.summary_box.pack(fill="both", expand=True)
            self.summary_box.bind("<Enter>", lambda e: self.show_tooltip(self.summary_box, "Displays the assignment summary and any violations"))
            self.summary_box.bind("<Leave>", lambda e: self.hide_tooltip())
            self.summary_box.tag_configure("header", font=font_label, foreground="#60a5fa")
            self.summary_box.tag_configure("loading", font=self.tooltip_font + ("italic",), foreground="#ffffff")

            self.tooltip = None

        except Exception as e:
            logging.error(f"Widget setup error: {str(e)}")
            messagebox.showerror("Error", f"Failed to initialize GUI: {str(e)}\nCheck duty_chart_app.log for details.")
            raise

    def show_tooltip(self, widget, text):
        try:
            if self.tooltip:
                self.tooltip.destroy()
            x, y = widget.winfo_rootx() + 25, widget.winfo_rooty() + 25
            self.tooltip = tk.Toplevel(widget)
            self.tooltip.wm_overrideredirect(True)
            self.tooltip.wm_geometry(f"+{x}+{y}")
            tooltip_bg = "#374151" if self.theme == "dark" else "#e5e7eb"
            label = tk.Label(self.tooltip, text=text, background=tooltip_bg, foreground="#000000", relief="solid", borderwidth=1, font=self.tooltip_font)
            label.pack()
        except Exception as e:
            logging.error(f"Tooltip creation error: {str(e)}")
            self.tooltip = None

    def hide_tooltip(self):
        if self.tooltip:
            self.tooltip.destroy()
            self.tooltip = None

    def on_resize(self, event):
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

    def update_progress(self, value):
        self.progress['value'] = value
        self.root.update_idletasks()

    def run(self):
        try:
            slot1_start = self.slot1_start.get_date()
            slot1_end = self.slot1_end.get_date()
            slot2_start = self.slot2_start.get_date()
            slot2_end = self.slot2_end.get_date()

            if any(d is None for d in [slot1_start, slot1_end, slot2_start, slot2_end]):
                messagebox.showerror("Error", "Invalid date selection. Please ensure all date fields are set.")
                return

            slot1 = (slot1_start, slot1_end)
            slot2 = (slot2_start, slot2_end)

            if slot1[0] > slot1[1] or slot2[0] > slot2[1]:
                messagebox.showerror("Error", "Start date must be before end date for both slots.")
                return

            if not self.input_path.get().strip() or not self.output_path.get().strip():
                messagebox.showerror("Error", "Please select both input and output files.")
                return

            self.generate_button.configure(state="disabled")
            self.clear_button.configure(state="disabled")
            self.progress['value'] = 0
            self.summary_box.delete("1.0", tk.END)
            self.summary_box.insert(tk.END, "Generating duty chart, please wait...\n", "loading")
            self.root.update_idletasks()

            logging.info(f"Calling generate_duty_chart with args: input_path={self.input_path.get()}, output_path={self.output_path.get()}, slot1={slot1}, slot2={slot2}, ratio_choice={self.ratio_choice.get()}")

            assignment_summary, ratio_violations, duty_quota_violations, slot_preference_violations, seventy_thirty_violations, _ = generate_duty_chart(
                self, self.input_path.get(), self.output_path.get(), slot1, slot2, self.ratio_choice.get())

            self.progress['value'] = 100
            self.summary_box.delete("1.0", tk.END)

            self.summary_box.insert(tk.END, "Assignment Summary:\n", "header")
            self.summary_box.insert(tk.END, f"{assignment_summary if assignment_summary else 'All sessions assigned'}\n\n")

            self.summary_box.insert(tk.END, "70:30 Rule Violations:\n", "header")
            self.summary_box.insert(tk.END, "\n".join(seventy_thirty_violations) + "\n" if seventy_thirty_violations else "No 70:30 rule violations.\n")

            self.summary_box.insert(tk.END, "Other Ratio Violations:\n", "header")
            self.summary_box.insert(tk.END, "\n".join(ratio_violations) + "\n" if ratio_violations else "No other ratio violations.\n")

            self.summary_box.insert(tk.END, "Duty Quota Violations:\n", "header")
            self.summary_box.insert(tk.END, "\n".join(duty_quota_violations) + "\n" if duty_quota_violations else "No duty quota violations.\n")

            self.summary_box.insert(tk.END, "Slot Preference Violations:\n", "header")
            self.summary_box.insert(tk.END, "\n".join(slot_preference_violations) + "\n" if slot_preference_violations else "No slot preference violations.\n")

            self.progress['value'] = 0

            self.generate_button.configure(state="normal")
            self.clear_button.configure(state="normal")

            messagebox.showinfo("Success", "Duty chart generated successfully! Check the output file and log for details.")

        except Exception as e:
            self.progress['value'] = 0
            self.generate_button.configure(state="normal")
            self.clear_button.configure(state="normal")
            self.summary_box.delete("1.0", tk.END)
            self.summary_box.insert(tk.END, f"Error: {str(e)}\n", "header")
            logging.error(f"GUI run error: {str(e)}")
            messagebox.showerror("Error", f"Error: {str(e)}\nCheck duty_chart_app.log for details.")
            self.root.update_idletasks()


if __name__ == "__main__":
    try:
        root = tk.Tk()
        app = DutyChartApp(root)
        root.mainloop()
    except Exception as e:
        logging.error(f"Main execution error: {str(e)}")
        messagebox.showerror("Error", f"Error starting application: {str(e)}\nCheck duty_chart_app.log for details.")
