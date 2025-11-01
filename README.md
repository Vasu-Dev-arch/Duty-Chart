# Duty Chart Generator – GCE Salem

**Automates exam duty assignment** for **50+ staff** across **30+ dates** using **70:30 ratio**, **slot preferences**, and **designation hierarchy**.

**Currently Used by Government College of Engineering, Salem** exam cell.  
**Saves 12+ hours** per exam schedule.

![Python](https://img.shields.io/badge/python-3.9%2B-blue)
![Tkinter](https://img.shields.io/badge/GUI-Tkinter-green)
![Pandas](https://img.shields.io/badge/Data-Pandas-orange)

---

## Features

- **Modern GUI** with dark/light theme toggle
- **Excel input/output** (`.xlsx`)
- **Fuzzy name matching** between staff list and preferences
- **Designation hierarchy**: Professor to Assoc. Professor to Asst. Professor to A.P(Contract)
- **Flexible duty ratios**: `1:3:6`, `1:3:7`, `1:4:8`
- **70:30 permanent:guest** staff ratio enforcement
- **Slot preference** enforcement with violation reporting
- **Real-time progress bar** and summary
- **Modular architecture**: `core/`, `gui/`, `config/`

---

## Screenshots

| GUI (Dark Mode) | Input Selection | Sample Input | Sample Output |
|------------------|------------------|--------------|---------------|
| ![GUI Dark Mode](./screenshots/GUI-1.png) | ![Browse Input](./screenshots/GUI-2.png) | ![Input Sample](./screenshots/input.png) | ![Output Sample](./screenshots/output.png) |

---

## Sample Files

| File | Description |
|------|-----------|
| [SampleInput.xlsx](./samples/SampleInput.xlsx) | Test input with 3 sheets: `Session Strength`, `Staff List`, `Slot Preference` |
| [output.xlsx](./samples/output.xlsx) | Generated duty chart with totals and violations |

---

## How to Run

```bash
# 1. Clone the repo
git clone https://github.com/YOUR_USERNAME/dutychart.git
cd dutychart

# 2. Create virtual environment
python -m venv .venv
.venv\Scripts\activate

# 3. Install dependencies
pip install -r requirements.txt

# 4. Run the app
python main.py

```
---

## Project Structure

```text
DutyChart/
├── core/          → Business logic (parser, matcher, scheduler)
├── gui/           → Tkinter interface
├── config/        → Logging setup
├── assets/        → Logo
├── samples/       → Test input/output
├── screenshots/   → Demo images
├── main.py        → Entry point
└── requirements.txt
```
---

## Tech Stack

- **Python 3.9+**
- **Pandas** – Data processing
- **OpenPyXL** – Excel I/O
- **Tkinter** – Native GUI
- **Pillow** – Image handling
- **tkcalendar** – Date picker

---

## Author

**Vasudevan**  
CSE Student, Government College of Engineering, Salem  

<a href="https://github.com/Vasu-Dev-arch" target="_blank">Github</a> | <a href="https://linkedin.com/in/vasudevan-j" target="_blank">LinkedIn</a> | <a href="https://vasu-dev-portfolio.netlify.app" target="_blank">Portfolio</a>