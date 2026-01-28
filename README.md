# ARMS Approver Automation

Automates adding approvers in the **ARMS (ARMS2 Unit Owner Packages)** web portal using Selenium.  
Reads account data from **Excel**, resumes from where it left off, and submits multiple approvers for each record.

---

## ✅ Features
- Reads OU ID & Account Name from **Excel**
- Adds **3 approvers** for every "New approver" link found
- Connects to **existing Chrome/Edge session** via remote debugging
- **Resume-on-crash** using `progress.json`
- Handles dynamic tables, autocomplete suggestions, and retries
- Stable automation tab management

---

## 📦 Requirements

Install everything using:

```bash
pip install -r requirements.txt

```
---

## Open Chrome/Edge Browser in Debug mode

```bash
& "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" `
>>   --remote-debugging-port=9222 `
>>   --user-data-dir="C:\EdgeSession"

```


