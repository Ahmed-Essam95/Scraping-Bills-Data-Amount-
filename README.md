# Billing Amount Automation Bot

This Python automation script uses **Selenium** and **openpyxl** to log into a web application, extract monthly billing amounts for accounts, and write the data to Excel. It mimics user behavior to navigate through multiple pages, interact with complex UI elements, and summarize invoice amounts per month.

---

## Project Purpose

Before automation, manually checking and extracting invoice data for hundreds of cst accounts (each with multiple dials) was a time-consuming process. The script now:
- **Logs in automatically**
- **Searches for each account**
- **Navigates through the billing and financial pages**
- **Extracts and summarizes invoice amounts**
- **Writes the results into structured Excel files**

This has saved over **95% of the manual task time**.

---

## Features

🔐 Automated login and secure navigation — logs into the billing portal and navigates safely through multiple sections.

🧭 Full web interaction — handles dropdown menus, buttons, radio fields, and tables with dynamic content.

⚙️ Smart scrolling & wait management — manages page loading delays and ensures all data is captured accurately.

📄 Invoice data extraction — scrapes billing and invoice details across 110+ pages with structured logic.

📊 Excel integration (OpenPyXL) — writes monthly breakdowns, invoice totals, and account summaries directly into Excel.

🔁 Real-time progress tracking — updates each account status as “Done / Not Done” in the source sheet.

🧠 Robust error handling — takes screenshots and refreshes sessions automatically if errors occur.

---

## Tools Used

- **Python**
- **Selenium** – browser automation
- **openpyxl** – Excel file reading and writing


## Author
Ahmed Essam

