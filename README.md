# Lux Class Command Center

A desktop operations application built for **Lux Beauty Lash Academy** to manage classes, students, enrollments, payments, inventory, readiness checks, reports, and backups from one place.

This project began as a simple class profit calculator and grew into a multi-tab business operations system while I was learning Python, SQL, relational database design, and desktop application development.

## What the app does

- Plans class revenue, expenses, profit, margin, break-even, and target profit
- Stores student records
- Schedules class sessions
- Enrolls students into classes
- Tracks tuition payments and balances
- Maintains live class rosters
- Tracks certification status
- Manages supply inventory
- Calculates per-student and fixed-per-class supply requirements
- Creates class readiness checks
- Provides search and filters across student and class data
- Generates CSV reports for rosters, payments, students, and readiness
- Creates and restores SQLite backups
- Includes plain-English help for a non-technical end user

## Tech stack

- **Python**
- **Tkinter / ttk**
- **SQLite**
- **SQL**
- **CSV**
- Python standard-library modules for backup, file handling, date/time, ZIP archives, and platform support

No external Python packages are required for the core application.

## Data model

The application uses related tables for:

- `students`
- `class_sessions`
- `enrollments`
- `payments`
- `supplies`
- `class_supply_requirements`

A student can enroll in multiple class sessions, and each enrollment can contain multiple payment records.

Payment status is derived from the payment ledger instead of being manually duplicated.

## Selected business logic

### Payment status

- `$0 paid` → **Not Paid**
- `Paid > $0 but below tuition` → **Deposit Paid**
- `Paid >= tuition` → **Paid in Full**

### Supply readiness

```text
Needed = (Per Student × Enrolled Students) + Fixed Per Class
```

### Class readiness

The readiness screen combines enrollment, outstanding balances, certification tracking, and supplies into one live operational view.

## Screenshots

Portfolio screenshots will be added using **fake/demo data only**. Recommended views:

1. Dashboard
2. Class Readiness
3. Students + Search/Filters
4. Enrollments + Payment Ledger
5. Reports or Supplies

## Running the application

### Windows

Install Python 3 with Tkinter support, then run:

```powershell
python .\lux_command_center.py
```

### macOS

Install a compatible Python 3 build with Tkinter support, then run:

```bash
python3 lux_command_center.py
```

## Privacy and repository safety

This public portfolio repository intentionally does **not** include the live business database, student/customer information, payment records, backup ZIPs, or local CSV history.

Those files are excluded with `.gitignore`.

## What I learned

This project gave me hands-on experience with:

- translating a real business workflow into software requirements
- relational database design
- writing and debugging SQL queries
- CRUD operations
- validation and duplicate prevention
- financial calculations
- stateful desktop interfaces
- report generation
- backup/restore design
- cross-platform considerations
- designing software for a non-technical end user
- iterative testing and feature development

## Development approach

This was an **AI-assisted learning project**. I used ChatGPT as a coding tutor, debugging partner, and development assistant while I learned concepts step by step, tested each feature, made product decisions, and iterated around a real business workflow.

I am documenting that openly because the goal is to demonstrate both the finished system **and** my ability to learn, reason through problems, test software, work with data, and turn business needs into functioning tools.

## Current status

The application has reached a working pilot stage for Lux Beauty Lash Academy.

Possible future improvements include packaged executables, automated tests, richer analytics, and migration to a networked or web-based architecture if multi-user access becomes necessary.

## Author

**Guinevere Negrete-Perez**

Portfolio project focused on Python, SQL, data workflows, desktop application development, and real-world business process automation.
