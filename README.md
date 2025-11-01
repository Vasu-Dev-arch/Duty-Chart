# Duty Chart Generator – GCE Salem

**Automates exam duty assignment** for 50+ staff across 30+ dates using **70:30 ratio**, **slot preferences**, and **designation hierarchy**.

Used by **Government College of Engineering, Salem** exam cell.  
Saves **12+ hours** per exam schedule.

---

## Features
- GUI with dark/light theme
- Excel input/output
- Fuzzy name matching
- Designation hierarchy
- 1:3:6, 1:3:7 and 1:3:8 ratio
- 70:30 permanent:guest ratio
- Slot preference enforcement
- Violation reporting
- Built with Python + Tkinter + Pandas   
![Python](https://img.shields.io/badge/python-3.9%2B-blue)
![Tkinter](https://img.shields.io/badge/GUI-Tkinter-green)
![Pandas](https://img.shields.io/badge/Data-Pandas-orange)

---

## Screenshots
[GUI Dark Mode](./screenshots/GUI-1.png) 
![GUI Dark Mode](./screenshots/GUI-1.png)

[Input Process](./screenshots/GUI-2.png) 
![Input Process](./screenshots/GUI-2.png)

[Input Sample](./screenshots/input.png) 
![Input Sample](./screenshots/input.png)

[Output Sample](./screenshots/output.png) 
![Output Sample](./screenshots/output.png)

---

## How to Run

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python main.py