# 🌱 Supplier Data Automation

This project automates the full process of importing, validating, and consolidating supplier data for sustainability reports.  
Originally developed for internal operations, all files and paths here have been fully anonymized.

---

## ⚙️ Overview

**Automated tasks include:**
1. Checking for new supplier Excel files.
2. Refreshing all Power Query connections.
3. Comparing imported files vs. database table.
4. Appending new data to the main report.
5. Archiving processed files.
6. Final verification and data refresh.

**Frequency:** Monthly  
**Time saved:** ~1.5h per run (≈18h/year)

---

## 🧩 Technologies Used
- Python (`pandas`, `xlwings`, `win32com.client`, `shutil`, `os`, `time`)
- Excel Power Query
- Process automation principles

---

## 💡 Key Features
- Fully automated data import and refresh cycle  
- Cross-check between file names and database entries  
- Automatic archiving of processed files  
- Error logging and recovery-safe execution  

---

## 📂 Repository Structure

