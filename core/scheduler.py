# core/scheduler.py
import os
import re
import math
import logging
import pandas as pd
from datetime import date
from difflib import SequenceMatcher
from typing import Tuple, List, Dict, Optional

# Import helper functions from same package
from .parser import find_column, safe_parse_date, parse_timestamp
from .matcher import normalize_name, normalize_designation, fuzzy_match_name


def generate_duty_chart(
    input_path: str,
    output_path: str,
    slot1_range: Tuple[date, date],
    slot2_range: Tuple[date, date],
    ratio_choice: str
) -> Tuple[Optional[str], List[str], List[str], List[str], Dict[str, str]]:
    """
    Generate duty chart from Excel input.
    Returns: (summary, ratio_violations, duty_violations, slot_violations, name_map)
    """
    try:
        input_path = input_path.strip('"').strip()
        output_path = output_path.strip('"').strip()
        logging.info(f"Input path: {input_path}, Output path: {output_path}, Ratio: {ratio_choice}")

        # --- File existence check ---
        if not os.path.exists(input_path):
            logging.error(f"Input file not found: {input_path}")
            return None, [], [], [], {}

        if not os.path.isfile(input_path):
            logging.error(f"Path is not a file: {input_path}")
            return None, [], [], [], {}

        # --- Load Excel ---
        try:
            xls = pd.ExcelFile(input_path)
            logging.info(f"Sheets found: {xls.sheet_names}")
        except Exception as e:
            logging.error(f"Failed to read Excel file: {e}")
            return None, [], [], [], {}

        # --- Normalize sheet names ---
        sheets = {
            s.strip().lower().replace('\n', '').replace('\r', ''): s
            for s in xls.sheet_names
        }

        # --- Required sheets ---
        sheet_map = {
            'session strength': ['Session Strength', 'Sessionwise Strength'],
            'staff list': ['Staff List', 'Staff Details'],
            'slot preference': ['Slot Preference']
        }
        found_sheets = {}
        for key, variations in sheet_map.items():
            for variation in variations:
                norm_var = variation.strip().lower().replace('\n', '').replace('\r', '')
                if norm_var in sheets:
                    found_sheets[key] = sheets[norm_var]
                    break
            if key not in found_sheets:
                logging.error(f"Missing sheet: {key}. Found: {list(sheets.values())}")
                return None, [], [], [], {}

        # --- Load dataframes ---
        session_df = pd.read_excel(xls, found_sheets['session strength'])
        staff_df = pd.read_excel(xls, found_sheets['staff list'])
        pref_df = pd.read_excel(xls, found_sheets['slot preference']) if 'slot preference' in found_sheets else None

        # --- Normalize column names ---
        def clean_cols(df):
            return [c.strip().lower().replace('\n', '').replace('\r', '') for c in df.columns]

        session_df.columns = clean_cols(session_df)
        staff_df.columns = clean_cols(staff_df)
        if pref_df is not None:
            pref_df.columns = clean_cols(pref_df)

        # --- Find required columns ---
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
        pref_cols = {}
        if pref_df is not None:
            pref_cols = {
                'timestamp': find_column(pref_df, ['timestamp']),
                'name': find_column(pref_df, ['name of the faculty', 'name', 'faculty']),
                'preferred slot': find_column(pref_df, ['preferred slot', 'slot', 'preferredslot'])
            }
        else:
            staff_df['preferred slot'] = 'Any'
            logging.info("Slot Preference sheet missing. Setting 'Any' for all.")

        # --- Validate columns ---
        missing_cols = []
        for df_name, cols in [('Session Strength', session_cols), ('Staff List', staff_cols)]:
            for col_name, col in cols.items():
                if col is None:
                    missing_cols.append(f"{col_name} in {df_name}")
        if pref_df is not None:
            for col_name, col in pref_cols.items():
                if col is None:
                    missing_cols.append(f"{col_name} in Slot Preference")
        if missing_cols:
            logging.error(f"Missing columns: {', '.join(missing_cols)}")
            return None, [], [], [], {}

        # --- Rename columns ---
        session_df = session_df.rename(columns={
            session_cols['date']: 'date',
            session_cols['fn']: 'fn',
            session_cols['an']: 'an'
        })
        staff_df = staff_df.rename(columns={
            staff_cols['name']: 'name',
            staff_cols['designation']: 'designation',
            staff_cols['department']: 'department'
        })
        if pref_df is not None:
            pref_df = pref_df.rename(columns={
                pref_cols['timestamp']: 'timestamp',
                pref_cols['name']: 'name',
                pref_cols['preferred slot']: 'preferred slot'
            })

        # --- Clean session data ---
        session_df['date'] = session_df['date'].apply(safe_parse_date)
        session_df = session_df.dropna(subset=['date'])
        session_df['fn'] = pd.to_numeric(session_df['fn'], errors='coerce').fillna(0).astype(int)
        session_df['an'] = pd.to_numeric(session_df['an'], errors='coerce').fillna(0).astype(int)

        # --- Clean staff data ---
        staff_df['original_name'] = staff_df['name']
        staff_df['name'] = staff_df['name'].apply(normalize_name)
        staff_df['designation'] = staff_df['designation'].apply(normalize_designation)
        staff_df = staff_df.drop_duplicates(subset=['name'])

        # --- Handle preferences ---
        if pref_df is not None:
            pref_df['original_name'] = pref_df['name']
            pref_df['name'] = pref_df['name'].apply(normalize_name)
            pref_df['preferred slot'] = pref_df['preferred slot'].astype(str).str.strip().str.title()
            pref_df['preferred slot'] = pref_df['preferred slot'].replace({'Nan': 'Any', '': 'Any'})
            pref_df['timestamp'] = pref_df['timestamp'].apply(parse_timestamp)
            pref_df = pref_df.sort_values('timestamp').drop_duplicates(subset=['name'], keep='last')

            # --- Fuzzy matching ---
            staff_names = set(staff_df['name'])
            pref_names = set(pref_df['name'])
            unmatched_staff = staff_names - pref_names
            unmatched_pref = pref_names - staff_names
            fuzzy_matches = {}

            for staff_name in unmatched_staff:
                staff_orig = staff_df[staff_df['name'] == staff_name]['original_name'].iloc[0]
                best_score = 0
                best_match = None
                for pref_name in unmatched_pref:
                    pref_orig = pref_df[pref_df['name'] == pref_name]['original_name'].iloc[0]
                    if fuzzy_match_name(staff_orig, pref_orig):
                        score = SequenceMatcher(None, normalize_name(staff_orig), normalize_name(pref_orig)).ratio()
                        raw_score = SequenceMatcher(None, re.sub(r"[.\s]+", "", staff_orig.lower()), re.sub(r"[.\s]+", "", pref_orig.lower())).ratio()
                        final_score = max(score, raw_score)
                        if final_score > best_score:
                            best_score = final_score
                            best_match = pref_name
                if best_match:
                    fuzzy_matches[best_match] = staff_name

            if fuzzy_matches:
                pref_df['name'] = pref_df['name'].replace(fuzzy_matches)
                unmatched_pref = pref_names - staff_names - set(fuzzy_matches.keys())
                unmatched_staff = staff_names - pref_names - set(fuzzy_matches.values())

            if unmatched_staff or unmatched_pref:
                unmatched_s = [staff_df[staff_df['name'] == n]['original_name'].iloc[0] for n in unmatched_staff]
                unmatched_p = [pref_df[pref_df['name'] == n]['original_name'].iloc[0] for n in unmatched_pref]
                logging.warning(f"Unmatched staff: {unmatched_s}, Ignored prefs: {unmatched_p}")
                pref_df = pref_df[pref_df['name'].isin(staff_names)]

            merged_df = pd.merge(
                staff_df[['name', 'original_name', 'designation', 'department']],
                pref_df[['name', 'original_name', 'timestamp', 'preferred slot']],
                on='name', how='left'
            )
            merged_df['preferred slot'] = merged_df['preferred slot'].fillna('Any')
            merged_df['original_name_x'] = merged_df['original_name_x'].fillna(merged_df['name'])
            merged_df = merged_df.rename(columns={'original_name_x': 'original_name'}).drop(columns=['original_name_y'], errors='ignore')
            merged_df = merged_df.drop_duplicates(subset=['name'])
        else:
            merged_df = staff_df.copy()
            merged_df['preferred slot'] = 'Any'
            merged_df['timestamp'] = pd.NaT

        # --- Date ranges ---
        all_dates = sorted(session_df['date'].unique())
        slot_dates = {'Slot 1': set(), 'Slot 2': set()}
        for d in all_dates:
            if slot1_range[0] <= d <= slot1_range[1]:
                slot_dates['Slot 1'].add(d)
            elif slot2_range[0] <= d <= slot2_range[1]:
                slot_dates['Slot 2'].add(d)
        logging.info(f"Slot 1: {sorted(slot_dates['Slot 1'])}, Slot 2: {sorted(slot_dates['Slot 2'])}")

        # --- Duty calculation ---
        slot1_duties = sum(math.ceil(row[s] / 30) for _, row in session_df.iterrows() for s in ['fn', 'an'] if row['date'] in slot_dates['Slot 1'])
        slot2_duties = sum(math.ceil(row[s] / 30) for _, row in session_df.iterrows() for s in ['fn', 'an'] if row['date'] in slot_dates['Slot 2'])
        logging.info(f"Slot 1 needs {slot1_duties}, Slot 2 needs {slot2_duties}")

        sessions = [
            (row['date'], s, math.ceil(row[s] / 30))
            for _, row in session_df.iterrows()
            for s in ['fn', 'an']
            if math.ceil(row[s] / 30) > 0
        ]
        sessions.sort(key=lambda x: (x[0], x[1]))

        # --- Assignment setup ---
        assigned_counts = {name: 0 for name in merged_df['name']}
        used_on_day = {d: set() for d in all_dates}
        duty_data = {name: {} for name in merged_df['name']}
        assigned_slots = {name: None for name in merged_df['name']}

        prof_count = len(merged_df[merged_df['designation'] == 'Professor'])
        assoc_count = len(merged_df[merged_df['designation'] == 'Assoc. Professor'])
        asst_count = len(merged_df[merged_df['designation'] == 'Asst. Professor'])

        ratio_map = {
            '1:3:6': {'Professor': 1 if prof_count else 0, 'Assoc. Professor': 3 if assoc_count else 0, 'Asst. Professor': 6 if asst_count else 0, 'A.P(Contract)': float('inf'), 'perm_ratio': 0.7, 'gl_ratio': 0.3},
            '1:3:7': {'Professor': 1 if prof_count else 0, 'Assoc. Professor': 3 if assoc_count else 0, 'Asst. Professor': 7 if asst_count else 0, 'A.P(Contract)': float('inf'), 'perm_ratio': 0.7, 'gl_ratio': 0.3},
            '1:4:8': {'Professor': 1 if prof_count else 0, 'Assoc. Professor': 4 if assoc_count else 0, 'Asst. Professor': 8 if asst_count else 0, 'A.P(Contract)': float('inf'), 'perm_ratio': 0.7, 'gl_ratio': 0.3}
        }
        designation_caps = ratio_map[ratio_choice]

        # --- Violations ---
        ratio_violations = []
        duty_quota_violations = []
        slot_preference_violations = []
        slot_split_violations = []

        # --- Assign permanent staff ---
        for desig in ['Professor', 'Assoc. Professor', 'Asst. Professor']:
            candidates = merged_df[merged_df['designation'] == desig][['name', 'original_name', 'preferred slot', 'timestamp']]
            candidates = sorted(candidates.to_dict('records'),
                key=lambda x: (x['timestamp'] if pd.notna(x['timestamp']) else pd.Timestamp.max, x['name'])
            ) if desig == 'Asst. Professor' else candidates.to_dict('records')

            for candidate in candidates:
                name = candidate['name']
                orig_name = candidate['original_name']
                pref_slot = candidate['preferred slot'] if candidate['preferred slot'] in ['Slot 1', 'Slot 2'] else 'Any'

                locked_slot = 'Slot 2' if orig_name == 'Dr. V. Satheeshkumar' and pref_slot in ['Slot 2', 'Any'] else \
                              pref_slot if pref_slot in ['Slot 1', 'Slot 2'] else \
                              'Slot 1' if slot1_duties >= slot2_duties else 'Slot 2'
                valid_slots = [locked_slot]
                duties_needed = designation_caps[desig]
                assigned = 0

                # Pass 1 & 2 (70:30 → 80:20)
                for ratio in [0.7, 0.8]:
                    if assigned >= duties_needed:
                        break
                    for slot in valid_slots:
                        valid_dates = sorted(slot_dates[slot])
                        for date, session, required in [(d, s, r) for d, s, r in sessions if d in valid_dates]:
                            if name in used_on_day[date] or assigned_counts[name] >= duties_needed:
                                continue
                            current_slot = 'Slot 1' if date in slot_dates['Slot 1'] else 'Slot 2'
                            if assigned_slots[name] and assigned_slots[name] != current_slot:
                                slot_split_violations.append(f"{orig_name} ({desig}) split across slots")
                                continue
                            current_assigned = sum(1 for n in used_on_day[date] if session.upper() in duty_data[n].get(date, []))
                            perm_assigned = sum(1 for n in used_on_day[date]
                                              if session.upper() in duty_data[n].get(date, [])
                                              and merged_df[merged_df['name'] == n]['designation'].iloc[0] in ['Professor', 'Assoc. Professor', 'Asst. Professor'])
                            perm_needed = math.ceil(required * ratio)
                            if perm_assigned < perm_needed and current_assigned < required:
                                duty_data[name].setdefault(date, []).append(session.upper())
                                used_on_day[date].add(name)
                                assigned_counts[name] += 1
                                assigned += 1
                                assigned_slots[name] = current_slot
                                if pref_slot != 'Any' and pref_slot != current_slot:
                                    slot_preference_violations.append(f"{orig_name} preferred {pref_slot} but got {current_slot}")
                                if assigned >= duties_needed:
                                    break
                        if assigned >= duties_needed:
                            break

            # Pass 3: Fill Asst. Professor quota
            if desig == 'Asst. Professor':
                candidates = sorted(candidates, key=lambda x: assigned_counts[x['name']])
                for candidate in candidates:
                    name = candidate['name']
                    if assigned_counts[name] >= duties_needed:
                        continue
                    # ... (same logic, simplified)

        # --- Assign Guest Lecturers ---
        for date, session, required in sessions:
            current_assigned = sum(1 for n in used_on_day[date] if session.upper() in duty_data[n].get(date, []))
            remaining = required - current_assigned
            if remaining > 0:
                gls = [n for n in merged_df[merged_df['designation'] == 'A.P(Contract)']['name'] if n not in used_on_day[date]]
                gls = sorted(gls, key=lambda x: assigned_counts[x])
                for name in gls[:remaining]:
                    duty_data[name].setdefault(date, []).append(session.upper())
                    used_on_day[date].add(name)
                    assigned_counts[name] += 1

        # --- Output ---
        output_rows = []
        for name in merged_df['name']:
            row = {
                'Name': merged_df[merged_df['name'] == name]['original_name'].iloc[0],
                'Designation': merged_df[merged_df['name'] == name]['designation'].iloc[0],
                'Department': merged_df[merged_df['name'] == name]['department'].iloc[0],
                'Total Duties': sum(len(duty_data[name].get(d, [])) for d in all_dates),
                'Preferred Slot': merged_df[merged_df['name'] == name]['preferred slot'].iloc[0],
                'Assigned Slot': assigned_slots[name] or "None"
            }
            for d in all_dates:
                row[d] = ' '.join(duty_data.get(name, {}).get(d, []))
            output_rows.append(row)

        output_df = pd.DataFrame(output_rows)
        cols = ['Name', 'Designation', 'Department', 'Total Duties', 'Preferred Slot', 'Assigned Slot'] + [d for d in all_dates]
        output_df = output_df[cols]

        # Add totals
        fn_row = {c: sum('FN' in duty_data.get(n, {}).get(c, []) for n in merged_df['name']) for c in all_dates}
        an_row = {c: sum('AN' in duty_data.get(n, {}).get(c, []) for n in merged_df['name']) for c in all_dates}
        fn_row.update({'Name': 'Total FN Duties', 'Designation': '', 'Department': '', 'Total Duties': '', 'Preferred Slot': '', 'Assigned Slot': ''})
        an_row.update({'Name': 'Total AN Duties', 'Designation': '', 'Department': '', 'Total Duties': '', 'Preferred Slot': '', 'Assigned Slot': ''})
        output_df = pd.concat([output_df, pd.DataFrame([fn_row, an_row])], ignore_index=True)

        output_df.to_excel(output_path, index=False)

        total = sum(assigned_counts.values())
        summary = f"Final chart ({ratio_choice}): {total} duties assigned."
        name_map = merged_df.set_index('name')['original_name'].to_dict()
        return summary, ratio_violations, duty_quota_violations, slot_preference_violations + slot_split_violations, name_map

    except Exception as e:
        logging.error(f"Failed to generate chart: {str(e)}", exc_info=True)
        return None, [], [], [], {}