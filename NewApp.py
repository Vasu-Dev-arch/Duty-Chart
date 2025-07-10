# Updated 

import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from datetime import datetime, timedelta
import math
import re
import os

# --- Utility functions ---
def normalize_name(name):
    name = re.sub(r'^(Prof\.|Dr\.)\s*', '', name, flags=re.IGNORECASE).strip()
    name = re.sub(r'\s+', ' ', name)
    parts = name.split()
    if len(parts) > 1 and len(parts[-1]) == 1:
        parts = parts[:-1]
    return ' '.join(parts).lower()

# --- Main Application ---
class DutyChartApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Duty Chart Generator")
        self.build_gui()

    def build_gui(self):
        tk.Label(self.root, text="Input Excel File:").grid(row=0, column=0)
        self.input_path = tk.StringVar()
        tk.Entry(self.root, textvariable=self.input_path, width=50).grid(row=0, column=1)
        tk.Button(self.root, text="Browse", command=self.browse_input).grid(row=0, column=2)

        tk.Label(self.root, text="Slot 1 (YYYY-MM-DD to YYYY-MM-DD):").grid(row=1, column=0)
        self.slot1 = tk.StringVar()
        tk.Entry(self.root, textvariable=self.slot1).grid(row=1, column=1)

        tk.Label(self.root, text="Slot 2 (YYYY-MM-DD to YYYY-MM-DD):").grid(row=2, column=0)
        self.slot2 = tk.StringVar()
        tk.Entry(self.root, textvariable=self.slot2).grid(row=2, column=1)

        tk.Label(self.root, text="Output Excel File:").grid(row=3, column=0)
        self.output_path = tk.StringVar()
        tk.Entry(self.root, textvariable=self.output_path, width=50).grid(row=3, column=1)
        tk.Button(self.root, text="Browse", command=self.browse_output).grid(row=3, column=2)

        tk.Button(self.root, text="Generate Duty Chart", command=self.run).grid(row=4, column=1, pady=10)

        self.summary = tk.Text(self.root, height=10, width=80)
        self.summary.grid(row=5, column=0, columnspan=3)

    def browse_input(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if path:
            self.input_path.set(path)

    def browse_output(self):
        path = filedialog.asksaveasfilename(defaultextension=".xlsx")
        if path:
            self.output_path.set(path)

    def get_slot_dates(self, slot_str):
        try:
            start_str, end_str = slot_str.strip().split('to')
            start = datetime.strptime(start_str.strip(), "%Y-%m-%d")
            end = datetime.strptime(end_str.strip(), "%Y-%m-%d")
            return [(start + timedelta(days=i)).strftime("%Y-%m-%d") for i in range((end-start).days+1)]
        except:
            raise ValueError("Invalid slot date format. Use YYYY-MM-DD to YYYY-MM-DD")
    def run(self):
        try:
            # Load Excel
            xl = pd.ExcelFile(self.input_path.get())
            strength = pd.read_excel(xl, 'Session Strength', parse_dates=['Date'])
            staff = pd.read_excel(xl, 'Staff List')
            pref = pd.read_excel(xl, 'Slot Preference')

            # Normalize names
            staff.rename(columns={"Name of the Faculty": "Name"}, inplace=True)
            pref.columns = [col.strip() for col in pref.columns]
            staff['NormName'] = staff['Name'].apply(normalize_name)
            pref['NormName'] = pref['Name'].apply(normalize_name)

            # Remove duplicate NormNames before setting index
            pref = pref.drop_duplicates(subset='NormName', keep='first')
            staff = staff.drop_duplicates(subset='NormName', keep='first')

            # Convert timestamp
            pref['Timestamp'] = pd.to_datetime(pref['Timestamp'], utc=True, errors='coerce').dt.tz_convert(None)

            # Build mappings
            preferences = pref.set_index('NormName').to_dict('index')
            designations = staff.set_index('NormName')['Designation'].to_dict()

            # Fix strength dates
            strength['Date'] = pd.to_datetime(strength['Date'], errors='coerce')

            # Build sessions
            sessions = {}
            for _, row in strength.iterrows():
                date = row['Date'].strftime("%Y-%m-%d")
                sessions[(date, 'FN')] = math.ceil(row['FN'] / 30)
                sessions[(date, 'AN')] = math.ceil(row['AN'] / 30)

            # Get slot ranges from user
            slot1_dates = self.get_slot_dates(self.slot1.get())
            slot2_dates = self.get_slot_dates(self.slot2.get())
            slot_map = {d: 'Slot 1' for d in slot1_dates}
            slot_map.update({d: 'Slot 2' for d in slot2_dates})

            # Initialize assignment data
            caps = {"Prof": 1, "ASP": 3, "AP": 6, "GL": 999}
            assigned = {n: [] for n in staff['NormName']}
            ap_conflicts = []

            # Assign duties
            for (date, session), needed in sessions.items():
                slot = slot_map.get(date)
                if not slot:
                    continue

                assigned_today = []
                perm_needed = math.ceil(needed * 0.7)
                gl_needed = needed - perm_needed

                # Assign permanent staff
                for des in ['Prof', 'ASP', 'AP']:
                    candidates = [
                        n for n in assigned
                        if designations[n] == des
                        and len(assigned[n]) < caps[des]
                        and (date, session) not in assigned[n]
                    ]

                    if des == 'AP':
                        candidates.sort(key=lambda x: preferences.get(x, {}).get('Timestamp', pd.Timestamp.max))
                    for name in candidates:
                        pref_slot = preferences.get(name, {}).get('Preferred Slot')
                        if des in ['Prof', 'ASP'] and pref_slot != slot:
                            continue
                        if des == 'AP' and pref_slot != slot:
                            ap_conflicts.append((
                                staff.loc[staff['NormName'] == name, 'Name'].values[0],
                                des, pref_slot, slot
                            ))
                            continue
                        assigned[name].append((date, session))
                        assigned_today.append(name)
                        if len(assigned_today) >= perm_needed:
                            break
                    if len(assigned_today) >= perm_needed:
                        break

                # Assign GLs if needed
                if len(assigned_today) < needed:
                    gls = [
                        n for n in assigned
                        if designations[n] == 'GL'
                        and (date, session) not in assigned[n]
                    ]
                    for name in gls:
                        assigned[name].append((date, session))
                        assigned_today.append(name)
                        if len(assigned_today) >= needed:
                            break

            # Prepare output
            names_dict = staff.set_index('NormName')['Name'].to_dict()
            dates_sorted = sorted({d for (d, _) in sessions})
            output_df = pd.DataFrame('', index=[names_dict[n] for n in assigned], columns=dates_sorted)

            for name, duties in assigned.items():
                real_name = names_dict.get(name, name)
                for d, s in duties:
                    val = output_df.at[real_name, d]
                    output_df.at[real_name, d] = (val + ' ' if val else '') + s

            output_df.to_excel(self.output_path.get())

            # Summary
            self.summary.delete(1.0, tk.END)
            self.summary.insert(tk.END, f"Duty Ratio Used: 1:3:6 (fallback logic)\n\n")
            self.summary.insert(tk.END, "APs Not Given Preferred Slot:\n")
            for name, des, pref, assigned_slot in ap_conflicts:
                self.summary.insert(tk.END, f"{name} ({des}) — Preferred: {pref}, Assigned: {assigned_slot}\n")

            messagebox.showinfo("Success", "Duty chart generated successfully!")

        except Exception as e:
            messagebox.showerror("Error", str(e))


if __name__ == '__main__':
    root = tk.Tk()
    app = DutyChartApp(root)
    root.mainloop()
