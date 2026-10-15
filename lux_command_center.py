import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
from tkinter import filedialog
import math
import csv
import os
import sqlite3
import zipfile
import tempfile
import shutil
from datetime import datetime


# =====================================================
# APPLICATION LOCATION
# =====================================================

# When Lux is launched by double-clicking a shortcut instead of
# from PowerShell, Windows may choose a different working folder.
#
# We force the working folder to be the folder containing this
# Python file so the database, CSV history, and backup folder
# are always found beside the application.

PROGRAM_FOLDER = os.path.dirname(
    os.path.abspath(__file__)
)

os.chdir(
    PROGRAM_FOLDER
)


APP_VERSION = "4.4"


# =====================================================
# HELPERS
# =====================================================

def parse_saved_number(value):
    # Older CSV records may contain formatted text:
    #
    # "$2,394.00"
    # "77.4%"
    #
    # Python cannot convert those directly into numbers,
    # so remove the formatting first.

    cleaned_value = (
        str(value)
        .replace("$", "")
        .replace("%", "")
        .replace(",", "")
        .strip()
    )

    return float(cleaned_value)


def normalize_class_date(value):
    # Humans like entering dates in different ways.
    #
    # We will accept:
    #
    # 10/24/2026
    #
    # OR
    #
    # 2026-10-24
    #
    # But SQLite will always STORE:
    #
    # 2026-10-24
    #
    # YYYY-MM-DD sorts properly in a database.

    cleaned_value = value.strip()


    possible_formats = (
        "%m/%d/%Y",
        "%Y-%m-%d"
    )


    for date_format in possible_formats:

        try:

            parsed_date = datetime.strptime(
                cleaned_value,
                date_format
            )

            return parsed_date.strftime(
                "%Y-%m-%d"
            )

        except ValueError:

            # "pass" means:
            #
            # Don't do anything here.
            # Just try the next format.

            pass


    # If neither format worked,
    # create a ValueError ourselves.

    raise ValueError


def normalize_class_time(value):
    # We will accept times such as:
    #
    # 9:00 AM
    # 09:00
    # 9 AM
    #
    # SQLite will store time using:
    #
    # HH:MM
    #
    # Example:
    #
    # 09:00
    # 14:30
    #
    # This makes chronological sorting easy.

    cleaned_value = (
        value
        .strip()
        .upper()
    )


    possible_formats = (
        "%I:%M %p",
        "%I %p",
        "%H:%M"
    )


    for time_format in possible_formats:

        try:

            parsed_time = datetime.strptime(
                cleaned_value,
                time_format
            )

            return parsed_time.strftime(
                "%H:%M"
            )

        except ValueError:

            pass


    raise ValueError


def display_class_date(value):
    # Convert the database format:
    #
    # 2026-10-24
    #
    # into the more familiar:
    #
    # 10/24/2026

    try:

        parsed_date = datetime.strptime(
            value,
            "%Y-%m-%d"
        )

        return parsed_date.strftime(
            "%m/%d/%Y"
        )

    except ValueError:

        return value


def display_class_time(value):
    # Convert:
    #
    # 09:00
    #
    # into:
    #
    # 9:00 AM

    try:

        parsed_time = datetime.strptime(
            value,
            "%H:%M"
        )


        formatted_time = parsed_time.strftime(
            "%I:%M %p"
        )


        # lstrip("0") removes a leading zero.
        #
        # 09:00 AM
        #
        # becomes:
        #
        # 9:00 AM

        return formatted_time.lstrip(
            "0"
        )

    except ValueError:

        return value


# =====================================================
# CLASS PRESETS
# =====================================================

# A dictionary stores:
#
# key : value
#
# Example:
#
# "Intro Class" : 399

class_prices = {
    "Intro Class": 399,
    "Masterclass": 1200,
    "Custom Class": 0
}


# =====================================================
# DATABASE SETUP
# =====================================================

def connect_database():
    # Every database operation should come through
    # this helper from now on.
    #
    # Why?
    #
    # SQLite foreign-key rules are OFF by default
    # for each new connection. Turning them on here
    # means our enrollment relationships are protected
    # every time Python opens the database.
    #
    # Example:
    #
    # If Jane is enrolled in Class Session #3,
    # SQLite can stop us from accidentally deleting
    # Jane or Session #3 while that enrollment exists.

    connection = sqlite3.connect(
        "lux_command_center.db"
    )

    connection.execute(
        "PRAGMA foreign_keys = ON"
    )

    return connection


def initialize_database():

    # sqlite3.connect() opens our database.
    #
    # If the file does not exist,
    # SQLite creates it automatically.

    connection = connect_database()


    # cursor lets Python send
    # SQL instructions to SQLite.

    cursor = connection.cursor()


    # -------------------------------------------------
    # STUDENTS TABLE
    # -------------------------------------------------

    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS students (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            first_name TEXT NOT NULL,
            last_name TEXT NOT NULL,
            phone TEXT,
            email TEXT,
            class_type TEXT NOT NULL,
            payment_status TEXT NOT NULL,
            certification_status TEXT NOT NULL,
            notes TEXT,
            created_at TEXT NOT NULL
        )
        """
    )


    # -------------------------------------------------
    # CLASS SESSIONS TABLE
    # -------------------------------------------------

    # Each scheduled class gets its OWN permanent ID.
    #
    # Example:
    #
    # Class Session #1
    # Intro Class
    # 2026-10-24
    # 09:00
    #
    # Even another Intro Class will get
    # a different session ID.

    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS class_sessions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            class_type TEXT NOT NULL,
            class_date TEXT NOT NULL,
            class_time TEXT NOT NULL,
            planned_students INTEGER NOT NULL,
            capacity INTEGER NOT NULL,
            tuition REAL NOT NULL,
            kit_cost REAL NOT NULL,
            other_expenses REAL NOT NULL,
            desired_profit REAL NOT NULL,
            status TEXT NOT NULL,
            created_at TEXT NOT NULL
        )
        """
    )


    # -------------------------------------------------
    # ENROLLMENTS TABLE
    # -------------------------------------------------

    # This is our BRIDGE table.
    #
    # students and class_sessions can now be connected
    # without stuffing one class name directly into
    # the student's identity record.
    #
    # student_id points to students.id
    # session_id points to class_sessions.id
    #
    # UNIQUE(student_id, session_id) means the same
    # student cannot accidentally be enrolled twice
    # in the same class session.

    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS enrollments (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            student_id INTEGER NOT NULL,
            session_id INTEGER NOT NULL,
            payment_status TEXT NOT NULL,
            certification_status TEXT NOT NULL,
            enrolled_at TEXT NOT NULL,

            UNIQUE(student_id, session_id),

            FOREIGN KEY(student_id)
                REFERENCES students(id)
                ON DELETE RESTRICT,

            FOREIGN KEY(session_id)
                REFERENCES class_sessions(id)
                ON DELETE RESTRICT
        )
        """
    )


    # -------------------------------------------------
    # PAYMENTS TABLE
    # -------------------------------------------------

    # One enrollment can now have MANY payments.
    #
    # Example:
    #
    # Enrollment #4
    #     Payment #1 = $100 deposit
    #     Payment #2 = $150 second payment
    #     Payment #3 = $149 final payment
    #
    # This is called a ONE-TO-MANY relationship:
    #
    # one enrollment -> many payment records
    #
    # We store each transaction separately instead of
    # overwriting one "amount paid" number. That gives
    # Lux an actual payment HISTORY.

    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS payments (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            enrollment_id INTEGER NOT NULL,
            amount REAL NOT NULL CHECK(amount > 0),
            payment_date TEXT NOT NULL,
            payment_method TEXT NOT NULL,
            notes TEXT,
            created_at TEXT NOT NULL,

            FOREIGN KEY(enrollment_id)
                REFERENCES enrollments(id)
                ON DELETE CASCADE
        )
        """
    )



    # -------------------------------------------------
    # SUPPLIES TABLE
    # -------------------------------------------------

    # This table stores Lux's current inventory.
    #
    # quantity_on_hand = what Lux physically has now
    # minimum_stock = the level where Lux wants a warning
    # cost_per_unit = approximate replacement cost
    #
    # REAL is used instead of INTEGER so Lux can later
    # track fractional units if a supply ever needs them.

    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS supplies (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            item_name TEXT NOT NULL COLLATE NOCASE UNIQUE,
            category TEXT NOT NULL,
            quantity_on_hand REAL NOT NULL DEFAULT 0
                CHECK(quantity_on_hand >= 0),
            minimum_stock REAL NOT NULL DEFAULT 0
                CHECK(minimum_stock >= 0),
            cost_per_unit REAL NOT NULL DEFAULT 0
                CHECK(cost_per_unit >= 0),
            supplier TEXT,
            notes TEXT,
            created_at TEXT NOT NULL
        )
        """
    )


    # -------------------------------------------------
    # CLASS SUPPLY REQUIREMENTS TABLE
    # -------------------------------------------------

    # This table connects a CLASS TYPE to a SUPPLY.
    #
    # Example:
    #
    # Intro Class -> Glue Rings -> 2 per student
    #
    # quantity_per_student grows with enrollment.
    #
    # fixed_quantity is used for something needed once
    # per class, regardless of the number of students.
    #
    # Example:
    #
    # 1 sanitation tray per class
    #
    # UNIQUE(class_type, supply_id) prevents the same
    # supply from being accidentally entered twice for
    # the same class type.

    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS class_supply_requirements (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            class_type TEXT NOT NULL,
            supply_id INTEGER NOT NULL,
            quantity_per_student REAL NOT NULL DEFAULT 0
                CHECK(quantity_per_student >= 0),
            fixed_quantity REAL NOT NULL DEFAULT 0
                CHECK(fixed_quantity >= 0),
            notes TEXT,
            created_at TEXT NOT NULL,

            UNIQUE(class_type, supply_id),

            FOREIGN KEY(supply_id)
                REFERENCES supplies(id)
                ON DELETE RESTRICT
        )
        """
    )


    cursor.execute(
        """
        CREATE INDEX IF NOT EXISTS
        idx_class_supply_requirements_class_type
        ON class_supply_requirements(class_type)
        """
    )


    # commit() permanently saves
    # database changes.

    connection.commit()


    connection.close()


# =====================================================
# PROGRAM MEMORY
# =====================================================

# Keeps track of which CSV class PLAN
# is loaded for editing.

loaded_plan_index = None


# Keeps track of which SQLite STUDENT
# is loaded for editing.

loaded_student_id = None


# Keeps track of which SQLite CLASS SESSION
# is loaded for editing.
#
# None means:
#
# "We are currently scheduling
# a brand-new class."

loaded_session_id = None


# Keeps track of which ENROLLMENT
# is loaded for editing.
#
# Unlike student_id or session_id,
# this is the ID of the relationship itself.

loaded_enrollment_id = None


# Connects visible History rows
# with complete CSV information.

history_rows = {}


# These dictionaries translate the friendly text
# shown in Enrollment dropdowns back into permanent
# SQLite IDs.
#
# Example:
#
# "#1 | Jane Doe"  ->  1

student_enrollment_lookup = {}
session_enrollment_lookup = {}



# Keeps track of the inventory item currently loaded
# into the Supplies edit form.

loaded_supply_id = None


# Keeps track of the class supply requirement currently
# loaded for editing.

loaded_requirement_id = None


# Friendly dropdown text -> permanent SQLite supply ID.

requirement_supply_lookup = {}


# Friendly scheduled-class text -> permanent session ID.

readiness_session_lookup = {}


# The dedicated v4.0 Class Readiness tab has its own
# session dropdown. This separate lookup keeps its
# friendly display text connected to permanent IDs.

class_readiness_session_lookup = {}


# =====================================================
# CLASS DROPDOWN
# =====================================================

def class_selected(event):

    selected_class = (
        class_combo.get()
    )


    tuition_entry.delete(
        0,
        tk.END
    )


    if selected_class == "Custom Class":

        tuition_entry.insert(
            0,
            ""
        )

    else:

        tuition_entry.insert(
            0,
            class_prices[selected_class]
        )


# =====================================================
# CALCULATE
# =====================================================

def calculate():

    try:

        students = int(
            students_entry.get()
        )

        tuition = float(
            tuition_entry.get()
        )

        kit_cost = float(
            kit_entry.get()
        )

        other_expenses = float(
            other_entry.get()
        )

        target_profit = float(
            target_profit_entry.get()
        )

        capacity = int(
            capacity_entry.get()
        )


        if students < 0 or capacity <= 0:
            raise ValueError


        # ---------------------------------------------
        # BASIC CLASS MATH
        # ---------------------------------------------

        revenue = (
            tuition * students
        )


        kit_expense = (
            kit_cost * students
        )


        total_expenses = (
            kit_expense + other_expenses
        )


        profit = (
            revenue - total_expenses
        )


        contribution_per_student = (
            tuition - kit_cost
        )


        # ---------------------------------------------
        # BREAK-EVEN
        # ---------------------------------------------

        if contribution_per_student > 0:

            if other_expenses > 0:

                break_even_students = math.ceil(
                    other_expenses
                    / contribution_per_student
                )

            else:

                break_even_students = 1


            if break_even_students == 1:

                break_even_text = (
                    "1 student"
                )

            else:

                break_even_text = (
                    f"{break_even_students} students"
                )


        else:

            break_even_text = (
                "Not possible"
            )


        # ---------------------------------------------
        # PROFIT GOAL
        # ---------------------------------------------

        if contribution_per_student > 0:

            students_for_goal = math.ceil(
                (other_expenses + target_profit)
                / contribution_per_student
            )


            if students_for_goal < 1:

                students_for_goal = 1


            if students_for_goal == 1:

                goal_text = (
                    "1 student"
                )

            else:

                goal_text = (
                    f"{students_for_goal} students"
                )


        else:

            goal_text = (
                "Not possible"
            )


        # ---------------------------------------------
        # PROFIT GOAL STATUS
        # ---------------------------------------------

        if contribution_per_student > 0:

            if students >= students_for_goal:

                goal_status_text = (
                    "✓ Profit goal reached"
                )

            else:

                students_remaining = (
                    students_for_goal - students
                )


                if students_remaining == 1:

                    goal_status_text = (
                        "1 more student needed"
                    )

                else:

                    goal_status_text = (
                        f"{students_remaining} "
                        f"more students needed"
                    )


        else:

            goal_status_text = (
                "Goal not possible at current pricing"
            )


        # ---------------------------------------------
        # CAPACITY
        # ---------------------------------------------

        if students < capacity:

            seats_remaining = (
                capacity - students
            )


            if seats_remaining == 1:

                capacity_text = (
                    "1 seat available"
                )

            else:

                capacity_text = (
                    f"{seats_remaining} seats available"
                )


        elif students == capacity:

            capacity_text = (
                "Class is full"
            )


        else:

            over_capacity = (
                students - capacity
            )


            if over_capacity == 1:

                capacity_text = (
                    "⚠ 1 student over capacity"
                )

            else:

                capacity_text = (
                    f"⚠ {over_capacity} students "
                    f"over capacity"
                )


        # ---------------------------------------------
        # GOAL VS CAPACITY
        # ---------------------------------------------

        if contribution_per_student > 0:

            if students_for_goal <= capacity:

                seats_after_goal = (
                    capacity - students_for_goal
                )


                if seats_after_goal == 0:

                    goal_capacity_text = (
                        "✓ Goal fits exactly at "
                        "full capacity"
                    )

                elif seats_after_goal == 1:

                    goal_capacity_text = (
                        "✓ Goal fits • "
                        "1 seat would remain"
                    )

                else:

                    goal_capacity_text = (
                        f"✓ Goal fits • "
                        f"{seats_after_goal} "
                        f"seats would remain"
                    )


            else:

                goal_over_capacity = (
                    students_for_goal - capacity
                )


                if goal_over_capacity == 1:

                    goal_capacity_text = (
                        "⚠ Goal exceeds capacity "
                        "by 1 student"
                    )

                else:

                    goal_capacity_text = (
                        f"⚠ Goal exceeds capacity by "
                        f"{goal_over_capacity} students"
                    )


        else:

            goal_capacity_text = (
                "Cannot evaluate at current pricing"
            )


        # ---------------------------------------------
        # PROFIT PER STUDENT
        # ---------------------------------------------

        if students > 0:

            profit_per_student = (
                profit / students
            )

        else:

            profit_per_student = 0


        # ---------------------------------------------
        # PROFIT MARGIN
        # ---------------------------------------------

        if revenue > 0:

            profit_margin = (
                profit / revenue
            ) * 100

        else:

            profit_margin = 0


        # ---------------------------------------------
        # UPDATE RESULTS
        # ---------------------------------------------

        revenue_result.config(
            text=f"${revenue:,.2f}"
        )


        expenses_result.config(
            text=f"${total_expenses:,.2f}"
        )


        profit_result.config(
            text=f"${profit:,.2f}"
        )

        profit_per_student_result.config(
            text=f"${profit_per_student:,.2f}"
        )


        profit_margin_result.config(
            text=f"{profit_margin:.1f}%"
        )


        break_even_result.config(
            text=break_even_text
        )


        goal_result.config(
            text=goal_text
        )


        goal_status_result.config(
            text=goal_status_text
        )


        capacity_result.config(
            text=capacity_text
        )


        goal_capacity_result.config(
            text=goal_capacity_text
        )


    except ValueError:

        messagebox.showerror(
            "Oops!",
            "Please enter valid numbers in all fields."
        )


# =====================================================
# RESET CALCULATOR
# =====================================================

def reset():

    global loaded_plan_index
    global loaded_session_id


    class_combo.set(
        "Intro Class"
    )


    students_entry.delete(
        0,
        tk.END
    )


    tuition_entry.delete(
        0,
        tk.END
    )

    tuition_entry.insert(
        0,
        "399"
    )


    kit_entry.delete(
        0,
        tk.END
    )

    kit_entry.insert(
        0,
        "75"
    )


    other_entry.delete(
        0,
        tk.END
    )

    other_entry.insert(
        0,
        "0"
    )


    target_profit_entry.delete(
        0,
        tk.END
    )

    target_profit_entry.insert(
        0,
        "0"
    )


    capacity_entry.delete(
        0,
        tk.END
    )

    capacity_entry.insert(
        0,
        "6"
    )


    # Clear scheduling information too.

    class_date_entry.delete(
        0,
        tk.END
    )


    class_time_entry.delete(
        0,
        tk.END
    )


    revenue_result.config(
        text="$0.00"
    )

    expenses_result.config(
        text="$0.00"
    )

    profit_result.config(
        text="$0.00"
    )

    profit_per_student_result.config(
        text="$0.00"
    )

    profit_margin_result.config(
        text="0.0%"
    )

    break_even_result.config(
        text="0 students"
    )

    goal_result.config(
        text="0 students"
    )

    goal_status_result.config(
        text=""
    )

    capacity_result.config(
        text=""
    )

    goal_capacity_result.config(
        text=""
    )


    loaded_plan_index = None

    loaded_session_id = None


    session_edit_status_label.config(
        text="Scheduling a new class session"
    )


# =====================================================
# SAVE CLASS PLAN TO CSV
# =====================================================

def save_plan():

    try:

        selected_class = (
            class_combo.get()
        )

        students = int(
            students_entry.get()
        )

        tuition = float(
            tuition_entry.get()
        )

        kit_cost = float(
            kit_entry.get()
        )

        other_expenses = float(
            other_entry.get()
        )

        target_profit = float(
            target_profit_entry.get()
        )

        capacity = int(
            capacity_entry.get()
        )


        if students < 0 or capacity <= 0:
            raise ValueError


        revenue = (
            tuition * students
        )

        kit_expense = (
            kit_cost * students
        )

        total_expenses = (
            kit_expense + other_expenses
        )

        profit = (
            revenue - total_expenses
        )


        contribution_per_student = (
            tuition - kit_cost
        )


        if students > 0:

            profit_per_student = (
                profit / students
            )

        else:

            profit_per_student = 0


        if revenue > 0:

            profit_margin = (
                profit / revenue
            ) * 100

        else:

            profit_margin = 0


        if contribution_per_student > 0:

            if other_expenses > 0:

                break_even_students = math.ceil(
                    other_expenses
                    / contribution_per_student
                )

            else:

                break_even_students = 1


            students_for_goal = math.ceil(
                (other_expenses + target_profit)
                / contribution_per_student
            )


            if students_for_goal < 1:
                students_for_goal = 1


        else:

            break_even_students = (
                "Not possible"
            )

            students_for_goal = (
                "Not possible"
            )


        file_name = (
            "lux_class_history.csv"
        )


        file_exists = (
            os.path.exists(file_name)
        )


        # "a" means APPEND.

        with open(
            file_name,
            "a",
            newline="",
            encoding="utf-8"
        ) as file:


            writer = csv.writer(
                file
            )


            if not file_exists:

                writer.writerow([
                    "Saved Date",
                    "Class Type",
                    "Students",
                    "Tuition Per Student",
                    "Kit Cost Per Student",
                    "Other Expenses",
                    "Desired Profit",
                    "Class Capacity",
                    "Revenue",
                    "Total Expenses",
                    "Estimated Profit",
                    "Profit Per Student",
                    "Profit Margin",
                    "Break-Even Students",
                    "Students Needed for Goal"
                ])


            writer.writerow([
                datetime.now().strftime(
                    "%Y-%m-%d %H:%M"
                ),
                selected_class,
                students,
                tuition,
                kit_cost,
                other_expenses,
                target_profit,
                capacity,
                revenue,
                total_expenses,
                profit,
                profit_per_student,
                profit_margin,
                break_even_students,
                students_for_goal
            ])


        messagebox.showinfo(
            "Saved!",
            "Class plan saved to lux_class_history.csv"
        )


        refresh_history_table()


    except ValueError:

        messagebox.showerror(
            "Oops!",
            "Please enter valid numbers before saving."
        )


# =====================================================
# UPDATE CLASS PLAN IN CSV
# =====================================================

def update_loaded_plan():

    global loaded_plan_index


    if loaded_plan_index is None:

        messagebox.showwarning(
            "No Loaded Plan",
            "Load a saved plan from History first."
        )

        return


    try:

        selected_class = (
            class_combo.get()
        )

        students = int(
            students_entry.get()
        )

        tuition = float(
            tuition_entry.get()
        )

        kit_cost = float(
            kit_entry.get()
        )

        other_expenses = float(
            other_entry.get()
        )

        target_profit = float(
            target_profit_entry.get()
        )

        capacity = int(
            capacity_entry.get()
        )


        if students < 0 or capacity <= 0:
            raise ValueError


        revenue = (
            tuition * students
        )


        kit_expense = (
            kit_cost * students
        )


        total_expenses = (
            kit_expense + other_expenses
        )


        profit = (
            revenue - total_expenses
        )


        contribution_per_student = (
            tuition - kit_cost
        )


        if students > 0:

            profit_per_student = (
                profit / students
            )

        else:

            profit_per_student = 0


        if revenue > 0:

            profit_margin = (
                profit / revenue
            ) * 100

        else:

            profit_margin = 0


        if contribution_per_student > 0:

            if other_expenses > 0:

                break_even_students = math.ceil(
                    other_expenses
                    / contribution_per_student
                )

            else:

                break_even_students = 1


            students_for_goal = math.ceil(
                (other_expenses + target_profit)
                / contribution_per_student
            )


            if students_for_goal < 1:
                students_for_goal = 1


        else:

            break_even_students = (
                "Not possible"
            )

            students_for_goal = (
                "Not possible"
            )


        file_name = (
            "lux_class_history.csv"
        )


        with open(
            file_name,
            "r",
            newline="",
            encoding="utf-8"
        ) as file:

            reader = csv.DictReader(
                file
            )

            fieldnames = (
                reader.fieldnames
            )

            all_rows = list(
                reader
            )


        if loaded_plan_index >= len(all_rows):

            messagebox.showerror(
                "Update Error",
                "That saved plan could not be found."
            )

            loaded_plan_index = None

            return


        all_rows[loaded_plan_index] = {

            "Saved Date": datetime.now().strftime(
                "%Y-%m-%d %H:%M"
            ),

            "Class Type": selected_class,
            "Students": students,
            "Tuition Per Student": tuition,
            "Kit Cost Per Student": kit_cost,
            "Other Expenses": other_expenses,
            "Desired Profit": target_profit,
            "Class Capacity": capacity,
            "Revenue": revenue,
            "Total Expenses": total_expenses,
            "Estimated Profit": profit,
            "Profit Per Student": profit_per_student,
            "Profit Margin": profit_margin,
            "Break-Even Students": break_even_students,
            "Students Needed for Goal": students_for_goal
        }


        with open(
            file_name,
            "w",
            newline="",
            encoding="utf-8"
        ) as file:

            writer = csv.DictWriter(
                file,
                fieldnames=fieldnames
            )

            writer.writeheader()

            writer.writerows(
                all_rows
            )


        messagebox.showinfo(
            "Updated!",
            "The loaded class plan was updated."
        )


        loaded_plan_index = None

        refresh_history_table()


    except ValueError:

        messagebox.showerror(
            "Oops!",
            "Please enter valid numbers before updating."
        )


# =====================================================
# REFRESH HISTORY
# =====================================================

def refresh_history_table():

    history_rows.clear()


    # get_children() gives us all visible rows.

    for item in history_table.get_children():

        # This clears only the DISPLAY.
        #
        # It does not delete CSV records.

        history_table.delete(
            item
        )


    file_name = (
        "lux_class_history.csv"
    )


    if not os.path.exists(file_name):

        history_status_label.config(
            text="No saved class plans yet."
        )

        return


    row_count = 0


    with open(
        file_name,
        "r",
        newline="",
        encoding="utf-8"
    ) as file:

        reader = csv.DictReader(
            file
        )


        # enumerate() gives us both:
        #
        # row_index = row number
        # row = actual saved information

        for row_index, row in enumerate(reader):

            revenue = parse_saved_number(
                row["Revenue"]
            )

            profit = parse_saved_number(
                row["Estimated Profit"]
            )

            margin = parse_saved_number(
                row["Profit Margin"]
            )


            item_id = history_table.insert(
                "",
                tk.END,
                values=(
                    row["Saved Date"],
                    row["Class Type"],
                    row["Students"],
                    f"${revenue:,.2f}",
                    f"${profit:,.2f}",
                    f"{margin:.1f}%"
                )
            )


            # table item ABC123
            #        ↓
            # index = 2
            # row = all saved class information

            history_rows[item_id] = {
                "index": row_index,
                "row": row
            }


            row_count += 1


    if row_count == 1:

        history_status_label.config(
            text="1 saved class plan"
        )

    else:

        history_status_label.config(
            text=f"{row_count} saved class plans"
        )


# =====================================================
# GO TO HISTORY
# =====================================================

def go_to_history():

    refresh_history_table()


    notebook.select(
        history_tab
    )


# =====================================================
# LOAD HISTORY PLAN
# =====================================================

def load_selected_plan():

    global loaded_plan_index
    global loaded_session_id


    selected_items = (
        history_table.selection()
    )


    if not selected_items:

        messagebox.showwarning(
            "Select a Plan",
            "Please select a saved class plan first."
        )

        return


    selected_item = (
        selected_items[0]
    )


    loaded_plan_index = history_rows[
        selected_item
    ]["index"]


    row = history_rows[
        selected_item
    ]["row"]


    # A CSV history plan is NOT automatically
    # tied to an existing scheduled session.

    loaded_session_id = None


    session_edit_status_label.config(
        text="Scheduling a new class session"
    )


    class_combo.set(
        row["Class Type"]
    )


    students_entry.delete(
        0,
        tk.END
    )

    students_entry.insert(
        0,
        row["Students"]
    )


    tuition_entry.delete(
        0,
        tk.END
    )

    tuition_entry.insert(
        0,
        str(
            parse_saved_number(
                row["Tuition Per Student"]
            )
        )
    )


    kit_entry.delete(
        0,
        tk.END
    )

    kit_entry.insert(
        0,
        str(
            parse_saved_number(
                row["Kit Cost Per Student"]
            )
        )
    )


    other_entry.delete(
        0,
        tk.END
    )

    other_entry.insert(
        0,
        str(
            parse_saved_number(
                row["Other Expenses"]
            )
        )
    )


    target_profit_entry.delete(
        0,
        tk.END
    )

    target_profit_entry.insert(
        0,
        str(
            parse_saved_number(
                row["Desired Profit"]
            )
        )
    )


    capacity_entry.delete(
        0,
        tk.END
    )

    capacity_entry.insert(
        0,
        str(
            int(
                parse_saved_number(
                    row["Class Capacity"]
                )
            )
        )
    )


    # History plans did not previously contain
    # schedule dates or times.
    #
    # Clear these so the user intentionally
    # chooses a date/time before scheduling.

    class_date_entry.delete(
        0,
        tk.END
    )


    class_time_entry.delete(
        0,
        tk.END
    )


    calculate()


    notebook.select(
        calculator_tab
    )


# =====================================================
# DELETE HISTORY PLAN
# =====================================================

def delete_selected_plan():

    global loaded_plan_index


    selected_items = (
        history_table.selection()
    )


    if not selected_items:

        messagebox.showwarning(
            "Select a Plan",
            "Please select a saved class plan first."
        )

        return


    selected_item = (
        selected_items[0]
    )


    selected_data = history_rows[
        selected_item
    ]


    target_index = selected_data[
        "index"
    ]


    row = selected_data[
        "row"
    ]


    confirmed = messagebox.askyesno(
        "Delete Plan",
        (
            f"Delete the saved "
            f"{row['Class Type']} plan "
            f"from {row['Saved Date']}?"
        )
    )


    # If they did NOT confirm:
    #
    # stop immediately.
    # Go home. Function over. 😂

    if not confirmed:
        return


    file_name = (
        "lux_class_history.csv"
    )


    with open(
        file_name,
        "r",
        newline="",
        encoding="utf-8"
    ) as file:

        reader = csv.DictReader(
            file
        )

        fieldnames = (
            reader.fieldnames
        )

        # reader moves through the CSV
        # one record at a time.
        #
        # list(reader) gathers all records
        # into one Python list.

        all_rows = list(
            reader
        )


    # Delete the item at this position.
    #
    # target_index = 2
    #
    # removes the THIRD item because
    # Python starts counting at zero.

    del all_rows[
        target_index
    ]


    with open(
        file_name,
        "w",
        newline="",
        encoding="utf-8"
    ) as file:

        writer = csv.DictWriter(
            file,
            fieldnames=fieldnames
        )

        writer.writeheader()

        writer.writerows(
            all_rows
        )


    loaded_plan_index = None


    messagebox.showinfo(
        "Deleted",
        "Saved class plan deleted."
    )


    refresh_history_table()


# =====================================================
# CREATE / SCHEDULE CLASS SESSION
# =====================================================

def schedule_class_session():

    try:

        selected_class = (
            class_combo.get()
        )


        planned_students = int(
            students_entry.get()
        )


        tuition = float(
            tuition_entry.get()
        )


        kit_cost = float(
            kit_entry.get()
        )


        other_expenses = float(
            other_entry.get()
        )


        target_profit = float(
            target_profit_entry.get()
        )


        capacity = int(
            capacity_entry.get()
        )


        # Turn human-friendly date/time input
        # into standardized database values.

        class_date = normalize_class_date(
            class_date_entry.get()
        )


        class_time = normalize_class_time(
            class_time_entry.get()
        )


        if planned_students < 0 or capacity <= 0:
            raise ValueError


        connection = connect_database()


        cursor = connection.cursor()


        # INSERT INTO creates a brand-new
        # scheduled class session.

        cursor.execute(
            """
            INSERT INTO class_sessions (
                class_type,
                class_date,
                class_time,
                planned_students,
                capacity,
                tuition,
                kit_cost,
                other_expenses,
                desired_profit,
                status,
                created_at
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                selected_class,
                class_date,
                class_time,
                planned_students,
                capacity,
                tuition,
                kit_cost,
                other_expenses,
                target_profit,
                "Scheduled",
                datetime.now().strftime(
                    "%Y-%m-%d %H:%M"
                )
            )
        )


        connection.commit()

        connection.close()


        refresh_class_sessions_table()


        messagebox.showinfo(
            "Class Scheduled",
            (
                f"{selected_class} scheduled for "
                f"{display_class_date(class_date)} "
                f"at {display_class_time(class_time)}."
            )
        )


        # Automatically take the user
        # to the Classes tab.

        notebook.select(
            classes_tab
        )


    except ValueError:

        messagebox.showerror(
            "Schedule Error",
            (
                "Please enter valid numbers, a date like "
                "10/24/2026, and a time like 9:00 AM."
            )
        )


# =====================================================
# READ / REFRESH CLASS SESSIONS
# =====================================================

def refresh_class_sessions_table(*args):

    # v4.1 adds SEARCH + FILTERS to the Classes tab.
    #
    # The Treeview is still showing the same class-session records,
    # but now SQL decides which rows belong in the current view.
    #
    # The *args lets Tkinter call this function from events such as
    # <<ComboboxSelected>> without us caring about the event object.

    selected_session_id = None


    current_selection = (
        class_sessions_table.selection()
    )


    if current_selection:

        current_values = class_sessions_table.item(
            current_selection[0],
            "values"
        )


        if current_values:

            selected_session_id = int(
                current_values[0]
            )


    # Clear only the visible table.

    for item in class_sessions_table.get_children():

        class_sessions_table.delete(
            item
        )


    connection = connect_database()

    cursor = connection.cursor()


    # Read the current filter controls.
    # These StringVars exist by the time this function is first called.

    search_text = (
        class_search_var.get()
        .strip()
        .lower()
    )

    class_type_filter = (
        class_type_filter_var.get()
    )

    date_filter = (
        class_date_filter_var.get()
    )

    seats_filter = (
        class_seats_filter_var.get()
    )

    status_filter = (
        class_status_filter_var.get()
    )


    # Build the WHERE and HAVING pieces one condition at a time.
    #
    # WHERE filters ordinary class-session columns.
    # HAVING filters the grouped COUNT(enrollments.id) result.

    where_conditions = []
    where_parameters = []

    having_conditions = []
    having_parameters = []


    if search_text:

        search_pattern = (
            f"%{search_text}%"
        )

        where_conditions.append(
            """
            (
                LOWER(class_sessions.class_type) LIKE ?
                OR LOWER(class_sessions.class_date) LIKE ?
                OR LOWER(class_sessions.class_time) LIKE ?
                OR LOWER(class_sessions.status) LIKE ?
            )
            """
        )

        where_parameters.extend([
            search_pattern,
            search_pattern,
            search_pattern,
            search_pattern
        ])


    if class_type_filter != "All Classes":

        where_conditions.append(
            "class_sessions.class_type = ?"
        )

        where_parameters.append(
            class_type_filter
        )


    today_text = datetime.now().strftime(
        "%Y-%m-%d"
    )


    if date_filter == "Upcoming":

        where_conditions.append(
            "class_sessions.class_date >= ?"
        )

        where_parameters.append(
            today_text
        )

    elif date_filter == "Today":

        where_conditions.append(
            "class_sessions.class_date = ?"
        )

        where_parameters.append(
            today_text
        )

    elif date_filter == "Past":

        where_conditions.append(
            "class_sessions.class_date < ?"
        )

        where_parameters.append(
            today_text
        )


    if status_filter != "All Statuses":

        where_conditions.append(
            "class_sessions.status = ?"
        )

        where_parameters.append(
            status_filter
        )


    if seats_filter == "Open Seats":

        having_conditions.append(
            "COUNT(enrollments.id) < class_sessions.capacity"
        )

    elif seats_filter == "Full":

        having_conditions.append(
            "COUNT(enrollments.id) = class_sessions.capacity"
        )

    elif seats_filter == "Over Capacity":

        having_conditions.append(
            "COUNT(enrollments.id) > class_sessions.capacity"
        )


    where_sql = ""

    if where_conditions:

        where_sql = (
            " WHERE "
            + " AND ".join(
                where_conditions
            )
        )


    having_sql = ""

    if having_conditions:

        having_sql = (
            " HAVING "
            + " AND ".join(
                having_conditions
            )
        )


    query = (
        """
        SELECT
            class_sessions.id,
            class_sessions.class_date,
            class_sessions.class_time,
            class_sessions.class_type,
            class_sessions.planned_students,
            COUNT(enrollments.id),
            class_sessions.capacity,
            class_sessions.tuition,
            class_sessions.status
        FROM class_sessions

        LEFT JOIN enrollments
            ON enrollments.session_id = class_sessions.id
        """
        + where_sql
        + """
        GROUP BY
            class_sessions.id
        """
        + having_sql
        + """
        ORDER BY
            class_sessions.class_date ASC,
            class_sessions.class_time ASC,
            class_sessions.id ASC
        """
    )


    cursor.execute(
        query,
        where_parameters + having_parameters
    )


    sessions = cursor.fetchall()


    # Total count is useful when filters are active.
    # Example:
    #
    # 3 of 18 classes

    cursor.execute(
        "SELECT COUNT(*) FROM class_sessions"
    )

    total_classes = cursor.fetchone()[0]


    connection.close()


    selected_item_after_refresh = None


    for session in sessions:

        session_id = session[0]
        class_date = session[1]
        class_time = session[2]
        class_type = session[3]
        planned_students = session[4]
        enrolled_students = session[5]
        capacity = session[6]
        tuition = session[7]
        status = session[8]


        item_id = class_sessions_table.insert(
            "",
            tk.END,
            values=(
                session_id,
                display_class_date(
                    class_date
                ),
                display_class_time(
                    class_time
                ),
                class_type,
                planned_students,
                enrolled_students,
                capacity,
                f"${tuition:,.2f}",
                status
            )
        )


        if session_id == selected_session_id:

            selected_item_after_refresh = (
                item_id
            )


    visible_count = len(
        sessions
    )


    filters_active = any([
        bool(search_text),
        class_type_filter != "All Classes",
        date_filter != "All Dates",
        seats_filter != "All Seats",
        status_filter != "All Statuses"
    ])


    if filters_active:

        class_session_count_label.config(
            text=(
                f"{visible_count} of "
                f"{total_classes} classes"
            )
        )

    elif total_classes == 1:

        class_session_count_label.config(
            text="1 scheduled class"
        )

    else:

        class_session_count_label.config(
            text=f"{total_classes} scheduled classes"
        )


    # Restore the selected class only if it is still visible.

    if selected_item_after_refresh is not None:

        class_sessions_table.selection_set(
            selected_item_after_refresh
        )

        class_sessions_table.focus(
            selected_item_after_refresh
        )

        class_sessions_table.see(
            selected_item_after_refresh
        )

        refresh_selected_class_roster()

    else:

        clear_class_roster()


# =====================================================
# CLEAR CLASS FILTERS
# =====================================================

def clear_class_filters():

    class_search_var.set(
        ""
    )

    class_type_filter_var.set(
        "All Classes"
    )

    class_date_filter_var.set(
        "All Dates"
    )

    class_seats_filter_var.set(
        "All Seats"
    )

    class_status_filter_var.set(
        "All Statuses"
    )

    refresh_class_sessions_table()


# =====================================================
# CLASS ROSTER
# =====================================================

def clear_class_roster():

    # Some functions can run before the Classes-tab
    # roster widgets have been created.
    #
    # This guard makes the helper safe during startup.

    if "class_roster_table" not in globals():
        return


    for item in class_roster_table.get_children():

        class_roster_table.delete(
            item
        )


    class_roster_title_label.config(
        text="Select a class above to view its roster"
    )


    class_roster_capacity_label.config(
        text=""
    )


    class_roster_money_label.config(
        text=""
    )


    class_roster_payment_label.config(
        text=""
    )


def refresh_selected_class_roster(event=None, show_warning=False):

    # Treeview selection gives us the visible row.
    # The first value in that row is the permanent
    # SQLite session ID.

    selected_items = (
        class_sessions_table.selection()
    )


    if not selected_items:

        clear_class_roster()


        if show_warning:

            messagebox.showwarning(
                "Select a Class",
                "Please select a scheduled class first."
            )

        return


    selected_item = selected_items[0]


    selected_values = class_sessions_table.item(
        selected_item,
        "values"
    )


    if not selected_values:

        clear_class_roster()
        return


    session_id = int(
        selected_values[0]
    )


    connection = connect_database()

    cursor = connection.cursor()


    # First fetch information ABOUT the class itself.

    cursor.execute(
        """
        SELECT
            class_type,
            class_date,
            class_time,
            planned_students,
            capacity,
            tuition,
            status
        FROM class_sessions
        WHERE id = ?
        """,
        (
            session_id,
        )
    )


    session = cursor.fetchone()


    if session is None:

        connection.close()

        clear_class_roster()

        return


    # -------------------------------------------------
    # ROSTER JOIN + PAYMENT TOTALS
    # -------------------------------------------------
    #
    # This query now reaches through THREE related
    # tables:
    #
    # students -> enrollments -> payments
    #
    # LEFT JOIN payments is important because a student
    # must still appear on the roster even when they
    # have paid $0.00.
    #
    # SUM(payments.amount) adds every payment belonging
    # to that enrollment.
    #
    # COALESCE(..., 0) changes a database NULL into 0
    # when there are no payment records yet.

    cursor.execute(
        """
        SELECT
            students.id,
            students.first_name,
            students.last_name,
            enrollments.payment_status,
            enrollments.certification_status,
            students.phone,
            students.email,
            enrollments.id,
            COALESCE(SUM(payments.amount), 0)
        FROM enrollments

        JOIN students
            ON students.id = enrollments.student_id

        LEFT JOIN payments
            ON payments.enrollment_id = enrollments.id

        WHERE enrollments.session_id = ?

        GROUP BY
            enrollments.id

        ORDER BY
            students.last_name ASC,
            students.first_name ASC,
            students.id ASC
        """,
        (
            session_id,
        )
    )


    roster = cursor.fetchall()

    connection.close()


    # Clear the old visible roster before inserting
    # the newly queried one.

    for item in class_roster_table.get_children():

        class_roster_table.delete(
            item
        )


    tuition = session[5]


    for student in roster:

        student_id = student[0]

        first_name = student[1]

        last_name = student[2]

        payment_status = student[3]

        certification_status = student[4]

        phone = student[5] or ""

        email = student[6] or ""

        amount_paid = float(
            student[8] or 0
        )

        balance_due = max(
            tuition - amount_paid,
            0
        )


        class_roster_table.insert(
            "",
            tk.END,
            values=(
                student_id,
                f"{first_name} {last_name}",
                payment_status,
                f"${amount_paid:,.2f}",
                f"${balance_due:,.2f}",
                certification_status,
                phone,
                email
            )
        )


    class_type = session[0]

    class_date = session[1]

    class_time = session[2]

    planned_students = session[3]

    capacity = session[4]

    status = session[6]


    enrolled_count = len(
        roster
    )


    seats_remaining = max(
        capacity - enrolled_count,
        0
    )


    expected_tuition = (
        tuition * enrolled_count
    )


    collected_revenue = sum(
        float(student[8] or 0)
        for student in roster
    )


    outstanding_balance = max(
        expected_tuition - collected_revenue,
        0
    )


    paid_in_full_count = sum(
        1
        for student in roster
        if student[3] == "Paid in Full"
    )


    deposit_paid_count = sum(
        1
        for student in roster
        if student[3] == "Deposit Paid"
    )


    not_paid_count = sum(
        1
        for student in roster
        if student[3] == "Not Paid"
    )


    # -------------------------------------------------
    # ROSTER SUMMARY
    # -------------------------------------------------

    class_roster_title_label.config(
        text=(
            f"Session #{session_id} • "
            f"{class_type} • "
            f"{display_class_date(class_date)} at "
            f"{display_class_time(class_time)} • "
            f"{status}"
        )
    )


    if seats_remaining == 1:

        seat_text = "1 seat remaining"

    else:

        seat_text = (
            f"{seats_remaining} seats remaining"
        )


    class_roster_capacity_label.config(
        text=(
            f"Planned: {planned_students} • "
            f"Actually enrolled: "
            f"{enrolled_count}/{capacity} • "
            f"{seat_text}"
        )
    )


    class_roster_money_label.config(
        text=(
            f"Expected: ${expected_tuition:,.2f} • "
            f"Collected: ${collected_revenue:,.2f} • "
            f"Outstanding: ${outstanding_balance:,.2f}"
        )
    )


    class_roster_payment_label.config(
        text=(
            f"Paid in Full: {paid_in_full_count} • "
            f"Deposit/Partial: {deposit_paid_count} • "
            f"Not Paid: {not_paid_count}"
        )
    )


# =====================================================
# LOAD SELECTED CLASS SESSION
# =====================================================

def load_selected_session():

    global loaded_session_id
    global loaded_plan_index


    selected_items = (
        class_sessions_table.selection()
    )


    if not selected_items:

        messagebox.showwarning(
            "Select a Class",
            "Please select a scheduled class first."
        )

        return


    selected_item = (
        selected_items[0]
    )


    selected_values = class_sessions_table.item(
        selected_item,
        "values"
    )


    # The first displayed value is
    # the permanent database ID.

    session_id = int(
        selected_values[0]
    )

    connection = connect_database()


    cursor = connection.cursor()


    # WHERE id = ?
    #
    # finds ONE exact class session.

    cursor.execute(
        """
        SELECT
            class_type,
            class_date,
            class_time,
            planned_students,
            capacity,
            tuition,
            kit_cost,
            other_expenses,
            desired_profit,
            status
        FROM class_sessions
        WHERE id = ?
        """,
        (
            session_id,
        )
    )


    session = cursor.fetchone()


    connection.close()


    if session is None:

        messagebox.showerror(
            "Class Not Found",
            "That class session could not be found."
        )

        refresh_class_sessions_table()

        return


    # Remember the permanent SQLite ID.
    #
    # UPDATE SESSION will use this.

    loaded_session_id = (
        session_id
    )


    # This is NOT a CSV history editing session.

    loaded_plan_index = None


    # ---------------------------------------------
    # LOAD SESSION INTO CALCULATOR
    # ---------------------------------------------

    class_combo.set(
        session[0]
    )


    class_date_entry.delete(
        0,
        tk.END
    )

    class_date_entry.insert(
        0,
        display_class_date(
            session[1]
        )
    )


    class_time_entry.delete(
        0,
        tk.END
    )

    class_time_entry.insert(
        0,
        display_class_time(
            session[2]
        )
    )


    students_entry.delete(
        0,
        tk.END
    )

    students_entry.insert(
        0,
        session[3]
    )


    capacity_entry.delete(
        0,
        tk.END
    )

    capacity_entry.insert(
        0,
        session[4]
    )


    tuition_entry.delete(
        0,
        tk.END
    )

    tuition_entry.insert(
        0,
        session[5]
    )


    kit_entry.delete(
        0,
        tk.END
    )

    kit_entry.insert(
        0,
        session[6]
    )


    other_entry.delete(
        0,
        tk.END
    )

    other_entry.insert(
        0,
        session[7]
    )


    target_profit_entry.delete(
        0,
        tk.END
    )

    target_profit_entry.insert(
        0,
        session[8]
    )


    session_edit_status_label.config(
        text=(
            f"Editing Class Session "
            f"#{session_id}"
        )
    )


    calculate()


    notebook.select(
        calculator_tab
    )


# =====================================================
# UPDATE LOADED CLASS SESSION
# =====================================================

def update_loaded_session():

    global loaded_session_id


    if loaded_session_id is None:

        messagebox.showwarning(
            "No Loaded Session",
            (
                "Load a scheduled class from "
                "the Classes tab first."
            )
        )

        return


    try:

        selected_class = (
            class_combo.get()
        )


        planned_students = int(
            students_entry.get()
        )


        tuition = float(
            tuition_entry.get()
        )


        kit_cost = float(
            kit_entry.get()
        )


        other_expenses = float(
            other_entry.get()
        )


        target_profit = float(
            target_profit_entry.get()
        )


        capacity = int(
            capacity_entry.get()
        )


        class_date = normalize_class_date(
            class_date_entry.get()
        )


        class_time = normalize_class_time(
            class_time_entry.get()
        )


        if planned_students < 0 or capacity <= 0:
            raise ValueError


        connection = connect_database()


        cursor = connection.cursor()


        # UPDATE changes an existing record.
        #
        # WHERE id = ? is what prevents us
        # from changing every session
        # in the database at once.
        #
        # Very important WHERE.
        # Tiny word. Massive responsibility. 😂

        cursor.execute(
            """
            UPDATE class_sessions
            SET
                class_type = ?,
                class_date = ?,
                class_time = ?,
                planned_students = ?,
                capacity = ?,
                tuition = ?,
                kit_cost = ?,
                other_expenses = ?,
                desired_profit = ?
            WHERE id = ?
            """,
            (
                selected_class,
                class_date,
                class_time,
                planned_students,
                capacity,
                tuition,
                kit_cost,
                other_expenses,
                target_profit,
                loaded_session_id
            )
        )


        connection.commit()

        connection.close()


        messagebox.showinfo(
            "Class Updated",
            (
                f"Class Session "
                f"#{loaded_session_id} "
                f"was updated."
            )
        )


        loaded_session_id = None


        session_edit_status_label.config(
            text="Scheduling a new class session"
        )


        refresh_class_sessions_table()


    except ValueError:

        messagebox.showerror(
            "Update Error",
            (
                "Please enter valid numbers, a date like "
                "10/24/2026, and a time like 9:00 AM."
            )
        )


# =====================================================
# DELETE SELECTED CLASS SESSION
# =====================================================

def delete_selected_session():

    global loaded_session_id


    selected_items = (
        class_sessions_table.selection()
    )


    if not selected_items:

        messagebox.showwarning(
            "Select a Class",
            "Please select a scheduled class first."
        )

        return


    selected_item = (
        selected_items[0]
    )


    selected_values = class_sessions_table.item(
        selected_item,
        "values"
    )


    session_id = int(
        selected_values[0]
    )


    class_date = (
        selected_values[1]
    )


    class_time = (
        selected_values[2]
    )


    class_type = (
        selected_values[3]
    )


    confirmed = messagebox.askyesno(
        "Delete Class Session",
        (
            f"Delete {class_type} on "
            f"{class_date} at {class_time}?"
        )
    )


    if not confirmed:
        return


    try:

        connection = connect_database()

        cursor = connection.cursor()


        # Foreign keys now protect our relationships.
        #
        # If students are enrolled in this session,
        # SQLite will refuse to delete it until those
        # enrollment records are removed first.

        cursor.execute(
            """
            DELETE FROM class_sessions
            WHERE id = ?
            """,
            (
                session_id,
            )
        )


        connection.commit()

        connection.close()


    except sqlite3.IntegrityError:

        try:
            connection.close()
        except Exception:
            pass


        messagebox.showwarning(
            "Class Has Enrollments",
            (
                "This class session still has enrolled students. "
                "Remove those enrollments first so the roster "
                "history is not accidentally orphaned."
            )
        )

        return


    if loaded_session_id == session_id:

        loaded_session_id = None


        session_edit_status_label.config(
            text="Scheduling a new class session"
        )


    refresh_class_sessions_table()

    clear_class_roster()

    refresh_enrollment_choices()

    refresh_enrollments_table()


    messagebox.showinfo(
        "Class Deleted",
        "The scheduled class was deleted."
    )


# =====================================================
# CLEAR STUDENT FORM
# =====================================================

def clear_student_form():

    global loaded_student_id


    student_first_name_entry.delete(
        0,
        tk.END
    )


    student_last_name_entry.delete(
        0,
        tk.END
    )


    student_phone_entry.delete(
        0,
        tk.END
    )


    student_email_entry.delete(
        0,
        tk.END
    )


    student_class_combo.set(
        "Intro Class"
    )


    payment_status_combo.set(
        "Not Paid"
    )


    certification_status_combo.set(
        "Not Started"
    )


    # Text widgets use a different coordinate system
    # than Entry widgets.
    #
    # "1.0" means:
    #
    # line 1
    # character 0

    student_notes_text.delete(
        "1.0",
        tk.END
    )


    loaded_student_id = None


    student_edit_status_label.config(
        text="Adding a new student"
    )


# =====================================================
# CREATE STUDENT
# =====================================================

def add_student():

    # .strip() removes accidental spaces
    # from the beginning and end.

    first_name = (
        student_first_name_entry
        .get()
        .strip()
    )


    last_name = (
        student_last_name_entry
        .get()
        .strip()
    )


    phone = (
        student_phone_entry
        .get()
        .strip()
    )


    email = (
        student_email_entry
        .get()
        .strip()
    )


    class_type = (
        student_class_combo.get()
    )


    payment_status = (
        payment_status_combo.get()
    )


    certification_status = (
        certification_status_combo.get()
    )


    notes = (
        student_notes_text
        .get(
            "1.0",
            tk.END
        )
        .strip()
    )


    if not first_name or not last_name:

        messagebox.showwarning(
            "Missing Name",
            "Please enter the student's first and last name."
        )

        return


    connection = connect_database()


    cursor = connection.cursor()


    # INSERT INTO means:
    #
    # CREATE a new database record.
    #
    # ? symbols are placeholders.
    #
    # Python safely supplies the values
    # separately underneath.

    cursor.execute(
        """
        INSERT INTO students (
            first_name,
            last_name,
            phone,
            email,
            class_type,
            payment_status,
            certification_status,
            notes,
            created_at
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            first_name,
            last_name,
            phone,
            email,
            class_type,
            payment_status,
            certification_status,
            notes,
            datetime.now().strftime(
                "%Y-%m-%d %H:%M"
            )
        )
    )


    connection.commit()

    connection.close()


    clear_student_form()

    refresh_students_table()


    messagebox.showinfo(
        "Student Saved",
        f"{first_name} {last_name} was added."
    )


# =====================================================
# READ / REFRESH STUDENTS
# =====================================================

def refresh_students_table(*args):

    # v4.1 gives the Students tab a real search-and-filter layer.
    #
    # We still return ONE row per student. Filters that depend on
    # enrollments use EXISTS subqueries so a student does not appear
    # multiple times just because they took multiple classes.

    for item in students_table.get_children():

        students_table.delete(
            item
        )


    connection = connect_database()

    cursor = connection.cursor()


    search_text = (
        student_search_var.get()
        .strip()
        .lower()
    )

    class_filter = (
        student_class_filter_var.get()
    )

    payment_filter = (
        student_payment_filter_var.get()
    )

    certification_filter = (
        student_cert_filter_var.get()
    )

    balance_filter = (
        student_balance_filter_var.get()
    )


    conditions = []
    parameters = []


    # -------------------------------------------------
    # SEARCH
    # -------------------------------------------------
    # Search name, phone, or email with one box.

    if search_text:

        search_pattern = (
            f"%{search_text}%"
        )

        conditions.append(
            """
            (
                LOWER(students.first_name) LIKE ?
                OR LOWER(students.last_name) LIKE ?
                OR LOWER(students.first_name || ' ' || students.last_name) LIKE ?
                OR LOWER(COALESCE(students.phone, '')) LIKE ?
                OR LOWER(COALESCE(students.email, '')) LIKE ?
            )
            """
        )

        parameters.extend([
            search_pattern,
            search_pattern,
            search_pattern,
            search_pattern,
            search_pattern
        ])


    # -------------------------------------------------
    # CLASS FILTER
    # -------------------------------------------------
    # EXISTS means:
    #
    # "Keep this student if at least one matching enrollment exists."

    if class_filter != "All Classes":

        conditions.append(
            """
            EXISTS (
                SELECT 1
                FROM enrollments

                JOIN class_sessions
                    ON class_sessions.id = enrollments.session_id

                WHERE enrollments.student_id = students.id
                  AND class_sessions.class_type = ?
            )
            """
        )

        parameters.append(
            class_filter
        )


    # -------------------------------------------------
    # CERTIFICATION FILTER
    # -------------------------------------------------

    if certification_filter != "All Certifications":

        conditions.append(
            """
            EXISTS (
                SELECT 1
                FROM enrollments

                WHERE enrollments.student_id = students.id
                  AND enrollments.certification_status = ?
            )
            """
        )

        parameters.append(
            certification_filter
        )


    # -------------------------------------------------
    # PAYMENT FILTER
    # -------------------------------------------------
    # Payment status is calculated from the REAL ledger now,
    # not the older profile-level payment label.

    if payment_filter == "Not Paid":

        conditions.append(
            """
            EXISTS (
                SELECT 1
                FROM enrollments

                JOIN class_sessions
                    ON class_sessions.id = enrollments.session_id

                WHERE enrollments.student_id = students.id
                  AND COALESCE(
                        (
                            SELECT SUM(payments.amount)
                            FROM payments
                            WHERE payments.enrollment_id = enrollments.id
                        ),
                        0
                      ) <= 0.005
            )
            """
        )

    elif payment_filter == "Deposit Paid":

        conditions.append(
            """
            EXISTS (
                SELECT 1
                FROM enrollments

                JOIN class_sessions
                    ON class_sessions.id = enrollments.session_id

                WHERE enrollments.student_id = students.id
                  AND COALESCE(
                        (
                            SELECT SUM(payments.amount)
                            FROM payments
                            WHERE payments.enrollment_id = enrollments.id
                        ),
                        0
                      ) > 0.005
                  AND COALESCE(
                        (
                            SELECT SUM(payments.amount)
                            FROM payments
                            WHERE payments.enrollment_id = enrollments.id
                        ),
                        0
                      ) < class_sessions.tuition - 0.005
            )
            """
        )

    elif payment_filter == "Paid in Full":

        conditions.append(
            """
            EXISTS (
                SELECT 1
                FROM enrollments

                JOIN class_sessions
                    ON class_sessions.id = enrollments.session_id

                WHERE enrollments.student_id = students.id
                  AND COALESCE(
                        (
                            SELECT SUM(payments.amount)
                            FROM payments
                            WHERE payments.enrollment_id = enrollments.id
                        ),
                        0
                      ) >= class_sessions.tuition - 0.005
            )
            """
        )


    # -------------------------------------------------
    # BALANCE FILTER
    # -------------------------------------------------

    if balance_filter == "Balance Due":

        conditions.append(
            """
            EXISTS (
                SELECT 1
                FROM enrollments

                JOIN class_sessions
                    ON class_sessions.id = enrollments.session_id

                WHERE enrollments.student_id = students.id
                  AND COALESCE(
                        (
                            SELECT SUM(payments.amount)
                            FROM payments
                            WHERE payments.enrollment_id = enrollments.id
                        ),
                        0
                      ) < class_sessions.tuition - 0.005
            )
            """
        )

    elif balance_filter == "No Balance":

        conditions.append(
            """
            EXISTS (
                SELECT 1
                FROM enrollments

                JOIN class_sessions
                    ON class_sessions.id = enrollments.session_id

                WHERE enrollments.student_id = students.id
            )
            AND NOT EXISTS (
                SELECT 1
                FROM enrollments

                JOIN class_sessions
                    ON class_sessions.id = enrollments.session_id

                WHERE enrollments.student_id = students.id
                  AND COALESCE(
                        (
                            SELECT SUM(payments.amount)
                            FROM payments
                            WHERE payments.enrollment_id = enrollments.id
                        ),
                        0
                      ) < class_sessions.tuition - 0.005
            )
            """
        )


    where_sql = ""

    if conditions:

        where_sql = (
            " WHERE "
            + " AND ".join(
                conditions
            )
        )


    query = (
        """
        SELECT
            students.id,
            students.first_name,
            students.last_name,
            (
                SELECT COUNT(*)
                FROM enrollments
                WHERE enrollments.student_id = students.id
            ) AS enrollment_count,
            COALESCE(
                (
                    SELECT SUM(
                        CASE
                            WHEN class_sessions.tuition > COALESCE(
                                (
                                    SELECT SUM(payments.amount)
                                    FROM payments
                                    WHERE payments.enrollment_id = enrollments.id
                                ),                                0
                            )
                            THEN class_sessions.tuition - COALESCE(
                                (
                                    SELECT SUM(payments.amount)
                                    FROM payments
                                    WHERE payments.enrollment_id = enrollments.id
                                ),
                                0
                            )
                            ELSE 0
                        END
                    )
                    FROM enrollments

                    JOIN class_sessions
                        ON class_sessions.id = enrollments.session_id

                    WHERE enrollments.student_id = students.id
                ),
                0
            ) AS outstanding_balance,
            students.phone,
            students.email
        FROM students
        """
        + where_sql
        + """
        ORDER BY
            students.last_name COLLATE NOCASE ASC,
            students.first_name COLLATE NOCASE ASC,
            students.id ASC
        """
    )


    cursor.execute(
        query,
        parameters
    )


    students = cursor.fetchall()


    cursor.execute(
        "SELECT COUNT(*) FROM students"
    )

    total_students = cursor.fetchone()[0]


    connection.close()


    for student in students:

        student_id = student[0]
        first_name = student[1]
        last_name = student[2]
        enrollment_count = student[3]
        outstanding_balance = student[4]
        phone = student[5]
        email = student[6]


        students_table.insert(
            "",
            tk.END,
            values=(
                student_id,
                f"{first_name} {last_name}",
                enrollment_count,
                f"${outstanding_balance:,.2f}",
                phone,
                email
            )
        )


    visible_count = len(
        students
    )


    filters_active = any([
        bool(search_text),
        class_filter != "All Classes",
        payment_filter != "All Payments",
        certification_filter != "All Certifications",
        balance_filter != "All Balances"
    ])


    if filters_active:

        student_count_label.config(
            text=(
                f"{visible_count} of "
                f"{total_students} students"
            )
        )

    elif total_students == 1:

        student_count_label.config(
            text="1 student"
        )

    else:

        student_count_label.config(
            text=f"{total_students} students"
        )


# =====================================================
# CLEAR STUDENT FILTERS
# =====================================================

def clear_student_filters():

    student_search_var.set(
        ""
    )

    student_class_filter_var.set(
        "All Classes"
    )

    student_payment_filter_var.set(
        "All Payments"
    )

    student_cert_filter_var.set(
        "All Certifications"
    )

    student_balance_filter_var.set(
        "All Balances"
    )

    refresh_students_table()


# =====================================================
# LOAD SELECTED STUDENT
# =====================================================

def load_selected_student():

    global loaded_student_id


    selected_items = (
        students_table.selection()
    )


    if not selected_items:

        messagebox.showwarning(
            "Select a Student",
            "Please select a student first."
        )

        return


    selected_item = (
        selected_items[0]
    )


    selected_values = students_table.item(
        selected_item,
        "values"
    )


    student_id = int(
        selected_values[0]
    )


    connection = connect_database()


    cursor = connection.cursor()


    cursor.execute(
        """
        SELECT
            first_name,
            last_name,
            phone,
            email,
            class_type,
            payment_status,
            certification_status,
            notes
        FROM students
        WHERE id = ?
        """,
        (
            student_id,
        )
    )


    student = cursor.fetchone()


    connection.close()


    if student is None:

        messagebox.showerror(
            "Student Not Found",
            "That student could not be found in the database."
        )

        refresh_students_table()

        return


    loaded_student_id = (
        student_id
    )


    student_first_name_entry.delete(
        0,
        tk.END
    )

    student_first_name_entry.insert(
        0,
        student[0]
    )


    student_last_name_entry.delete(
        0,
        tk.END
    )

    student_last_name_entry.insert(
        0,
        student[1]
    )


    student_phone_entry.delete(
        0,
        tk.END
    )

    student_phone_entry.insert(
        0,
        student[2] or ""
    )


    student_email_entry.delete(
        0,
        tk.END
    )

    student_email_entry.insert(
        0,
        student[3] or ""
    )


    student_class_combo.set(
        student[4]
    )


    payment_status_combo.set(
        student[5]
    )


    certification_status_combo.set(
        student[6]
    )


    student_notes_text.delete(
        "1.0",
        tk.END
    )


    student_notes_text.insert(
        "1.0",
        student[7] or ""
    )


    student_edit_status_label.config(
        text=(
            f"Editing Student #{student_id}: "
            f"{student[0]} {student[1]}"
        )
    )


# =====================================================
# UPDATE LOADED STUDENT
# =====================================================

def update_loaded_student():

    global loaded_student_id


    if loaded_student_id is None:

        messagebox.showwarning(
            "No Loaded Student",
            "Load a student from the table first."
        )

        return


    first_name = (
        student_first_name_entry
        .get()
        .strip()
    )


    last_name = (
        student_last_name_entry
        .get()
        .strip()
    )


    phone = (
        student_phone_entry
        .get()
        .strip()
    )


    email = (
        student_email_entry
        .get()
        .strip()
    )


    class_type = (
        student_class_combo.get()
    )


    payment_status = (
        payment_status_combo.get()
    )


    certification_status = (
        certification_status_combo.get()
    )


    notes = (
        student_notes_text
        .get(
            "1.0",
            tk.END
        )
        .strip()
    )


    if not first_name or not last_name:

        messagebox.showwarning(
            "Missing Name",
            "Please enter the student's first and last name."
        )

        return


    connection = connect_database()


    cursor = connection.cursor()


    # UPDATE changes an existing record.
    #
    # WHERE id = ? is critical.
    #
    # Without WHERE, every student could
    # receive the same information.
    #
    # Database nightmare fuel. 😂

    cursor.execute(
        """
        UPDATE students
        SET
            first_name = ?,
            last_name = ?,
            phone = ?,
            email = ?,
            class_type = ?,
            payment_status = ?,
            certification_status = ?,
            notes = ?
        WHERE id = ?
        """,
        (
            first_name,
            last_name,
            phone,
            email,
            class_type,
            payment_status,
            certification_status,
            notes,
            loaded_student_id
        )
    )


    connection.commit()

    connection.close()


    messagebox.showinfo(
        "Student Updated",
        f"{first_name} {last_name}'s record was updated."
    )


    clear_student_form()

    refresh_students_table()


# =====================================================
# DELETE SELECTED STUDENT
# =====================================================

def delete_selected_student():

    global loaded_student_id


    selected_items = (
        students_table.selection()
    )


    if not selected_items:

        messagebox.showwarning(
            "Select a Student",
            "Please select a student first."
        )

        return


    selected_item = (
        selected_items[0]
    )


    selected_values = students_table.item(
        selected_item,
        "values"
    )


    student_id = int(
        selected_values[0]
    )


    student_name = (
        selected_values[1]
    )


    confirmed = messagebox.askyesno(
        "Delete Student",
        f"Delete {student_name} from the student database?"
    )


    if not confirmed:
        return


    try:

        connection = connect_database()

        cursor = connection.cursor()


        # Because enrollments now point to students,
        # SQLite protects any student who still has
        # enrollment history.
        #
        # This prevents a classic database disaster:
        # an enrollment pointing to a student
        # who no longer exists.

        cursor.execute(
            """
            DELETE FROM students
            WHERE id = ?
            """,
            (
                student_id,
            )
        )


        connection.commit()

        connection.close()


    except sqlite3.IntegrityError:

        try:
            connection.close()
        except Exception:
            pass


        messagebox.showwarning(
            "Student Has Enrollments",
            (
                f"{student_name} is still enrolled in one or more "
                "class sessions. Remove those enrollments first "
                "before deleting the student record."
            )
        )

        return


    if loaded_student_id == student_id:

        clear_student_form()


    refresh_students_table()

    refresh_enrollment_choices()

    refresh_enrollments_table()


    messagebox.showinfo(
        "Student Deleted",
        f"{student_name} was deleted."
    )



# =====================================================
# ENROLLMENT RELATIONSHIPS
# =====================================================

def refresh_enrollment_choices():

    global student_enrollment_lookup
    global session_enrollment_lookup


    # The Enrollment tab does not exist yet while
    # Python is defining the functions above.
    #
    # This small guard lets other parts of the program
    # safely call this function only after the GUI
    # widgets have been created.

    if "student_enrollment_combo" not in globals():
        return


    current_student_choice = (
        student_enrollment_combo.get()
    )


    current_session_choice = (
        session_enrollment_combo.get()
    )


    connection = connect_database()

    cursor = connection.cursor()


    # -------------------------------------------------
    # STUDENT CHOICES
    # -------------------------------------------------

    cursor.execute(
        """
        SELECT
            id,
            first_name,
            last_name
        FROM students
        ORDER BY
            last_name ASC,
            first_name ASC,
            id ASC
        """
    )


    students = cursor.fetchall()


    # -------------------------------------------------
    # CLASS SESSION CHOICES
    # -------------------------------------------------

    # LEFT JOIN lets us count actual enrollments
    # even when a class currently has ZERO students.
    #
    # COUNT(enrollments.id) is our first taste of
    # asking multiple related tables one question.

    cursor.execute(
        """
        SELECT
            class_sessions.id,
            class_sessions.class_date,
            class_sessions.class_time,
            class_sessions.class_type,
            COUNT(enrollments.id),
            class_sessions.capacity
        FROM class_sessions
        LEFT JOIN enrollments
            ON enrollments.session_id = class_sessions.id
        GROUP BY
            class_sessions.id
        ORDER BY
            class_sessions.class_date ASC,
            class_sessions.class_time ASC,
            class_sessions.id ASC
        """
    )


    sessions = cursor.fetchall()


    connection.close()


    # -------------------------------------------------
    # BUILD FRIENDLY DROPDOWN TEXT
    # -------------------------------------------------

    student_enrollment_lookup = {}

    student_choices = []


    for student in students:

        student_id = student[0]

        first_name = student[1]

        last_name = student[2]


        display_text = (
            f"#{student_id} | "
            f"{first_name} {last_name}"
        )


        student_enrollment_lookup[
            display_text
        ] = student_id


        student_choices.append(
            display_text
        )


    session_enrollment_lookup = {}

    session_choices = []


    for session in sessions:

        session_id = session[0]

        class_date = session[1]

        class_time = session[2]

        class_type = session[3]

        enrolled_count = session[4]

        capacity = session[5]


        seats_left = (
            capacity - enrolled_count
        )


        if seats_left <= 0:

            seat_text = "FULL"

        elif seats_left == 1:

            seat_text = "1 seat left"

        else:

            seat_text = (
                f"{seats_left} seats left"
            )


        display_text = (
            f"#{session_id} | "
            f"{display_class_date(class_date)} "
            f"{display_class_time(class_time)} | "
            f"{class_type} | "
            f"{enrolled_count}/{capacity} enrolled | "
            f"{seat_text}"
        )


        session_enrollment_lookup[
            display_text
        ] = session_id


        session_choices.append(
            display_text
        )


    student_enrollment_combo[
        "values"
    ] = student_choices


    session_enrollment_combo[
        "values"
    ] = session_choices


    # Keep the current selection if it still exists.
    #
    # If someone was deleted or a class disappeared,
    # clear the now-invalid choice.

    if current_student_choice in student_enrollment_lookup:

        student_enrollment_combo.set(
            current_student_choice
        )

    elif loaded_enrollment_id is None:

        student_enrollment_combo.set(
            ""
        )


    if current_session_choice in session_enrollment_lookup:

        session_enrollment_combo.set(
            current_session_choice
        )

    elif loaded_enrollment_id is None:

        session_enrollment_combo.set(
            ""
        )


def clear_enrollment_form():

    global loaded_enrollment_id


    loaded_enrollment_id = None


    # Return the two identity dropdowns to normal
    # read-only selection mode.

    student_enrollment_combo.config(
        state="readonly"
    )


    session_enrollment_combo.config(
        state="readonly"
    )


    student_enrollment_combo.set(
        ""
    )


    session_enrollment_combo.set(
        ""
    )


    # Payment status is now AUTOMATIC.
    #
    # Actual payment transactions determine whether
    # an enrollment is Not Paid, Deposit Paid, or
    # Paid in Full.

    enrollment_payment_combo.set(
        "Not Paid"
    )


    enrollment_certification_combo.set(
        "Not Started"
    )


    enrollment_edit_status_label.config(
        text="Creating a new enrollment"
    )


    clear_payment_section()


def get_enrollment_money(enrollment_id, cursor):

    # Return tuition and ACTUAL recorded payments for
    # one enrollment.
    #
    # SUM() is an aggregate function. It takes many
    # payment rows and turns them into one total.

    cursor.execute(
        """
        SELECT
            class_sessions.tuition,
            COALESCE(SUM(payments.amount), 0)
        FROM enrollments

        JOIN class_sessions
            ON class_sessions.id = enrollments.session_id

        LEFT JOIN payments
            ON payments.enrollment_id = enrollments.id

        WHERE enrollments.id = ?

        GROUP BY enrollments.id
        """,
        (
            enrollment_id,
        )
    )


    result = cursor.fetchone()


    if result is None:
        return None


    tuition = float(
        result[0]
    )


    total_paid = float(
        result[1] or 0
    )


    balance_due = max(
        tuition - total_paid,
        0
    )


    return (
        tuition,
        total_paid,
        balance_due
    )


def sync_enrollment_payment_status(enrollment_id, connection):

    # The old version let us manually type a payment
    # status. v3.7 has better evidence: actual dollars.
    #
    # $0 paid             -> Not Paid
    # some money paid     -> Deposit Paid
    # tuition fully paid  -> Paid in Full

    cursor = connection.cursor()


    money = get_enrollment_money(
        enrollment_id,
        cursor
    )


    if money is None:
        return


    tuition = money[0]

    total_paid = money[1]


    if total_paid <= 0:

        new_status = "Not Paid"

    elif total_paid >= tuition:

        new_status = "Paid in Full"

    else:

        new_status = "Deposit Paid"


    cursor.execute(
        """
        UPDATE enrollments
        SET payment_status = ?
        WHERE id = ?
        """,
        (
            new_status,
            enrollment_id
        )
    )


def clear_payment_section():

    # This function is safe during startup before
    # payment widgets exist.

    if "payment_history_table" not in globals():
        return


    payment_enrollment_label.config(
        text="Load an enrollment above to manage payments"
    )


    payment_summary_label.config(
        text=""
    )


    payment_legacy_note_label.config(
        text=""    )


    payment_amount_entry.delete(
        0,
        tk.END
    )


    payment_date_entry.delete(
        0,
        tk.END
    )


    payment_date_entry.insert(
        0,
        datetime.now().strftime(
            "%m/%d/%Y"
        )
    )


    payment_method_combo.set(
        "Cash"
    )


    payment_notes_entry.delete(
        0,
        tk.END
    )


    for item in payment_history_table.get_children():

        payment_history_table.delete(
            item
        )


def refresh_payment_section():

    if "payment_history_table" not in globals():
        return


    if loaded_enrollment_id is None:

        clear_payment_section()
        return


    connection = connect_database()

    cursor = connection.cursor()


    cursor.execute(
        """
        SELECT
            students.first_name,
            students.last_name,
            class_sessions.class_type,
            class_sessions.class_date,
            class_sessions.class_time,
            enrollments.payment_status
        FROM enrollments

        JOIN students
            ON students.id = enrollments.student_id

        JOIN class_sessions
            ON class_sessions.id = enrollments.session_id

        WHERE enrollments.id = ?
        """,
        (
            loaded_enrollment_id,
        )
    )


    enrollment = cursor.fetchone()


    if enrollment is None:

        connection.close()

        clear_payment_section()
        return


    money = get_enrollment_money(
        loaded_enrollment_id,
        cursor
    )


    cursor.execute(
        """
        SELECT
            id,
            payment_date,
            amount,
            payment_method,
            notes
        FROM payments
        WHERE enrollment_id = ?
        ORDER BY
            payment_date ASC,
            id ASC
        """,
        (
            loaded_enrollment_id,
        )
    )


    payments = cursor.fetchall()

    connection.close()


    first_name = enrollment[0]

    last_name = enrollment[1]

    class_type = enrollment[2]

    class_date = enrollment[3]

    class_time = enrollment[4]

    payment_status = enrollment[5]


    payment_enrollment_label.config(
        text=(
            f"Enrollment #{loaded_enrollment_id} • "
            f"{first_name} {last_name} • "
            f"{class_type} • "
            f"{display_class_date(class_date)} at "
            f"{display_class_time(class_time)}"
        )
    )


    if money is not None:

        tuition = money[0]

        total_paid = money[1]

        balance_due = money[2]


        payment_summary_label.config(
            text=(
                f"Tuition: ${tuition:,.2f} • "
                f"Recorded Paid: ${total_paid:,.2f} • "
                f"Balance Due: ${balance_due:,.2f}"
            )
        )


        # Existing v3.5/v3.6 records may have a text
        # payment status but no dollar transactions.
        # We refuse to invent money that was never
        # recorded.

        if total_paid == 0 and payment_status != "Not Paid":

            payment_legacy_note_label.config(
                text=(
                    f"Legacy status says '{payment_status}', "
                    "but no dollar payment history has been recorded yet."
                )
            )

        else:

            payment_legacy_note_label.config(
                text=""
            )


    for item in payment_history_table.get_children():

        payment_history_table.delete(
            item
        )


    for payment in payments:

        payment_id = payment[0]

        payment_date = payment[1]

        amount = float(
            payment[2]
        )

        payment_method = payment[3]

        notes = payment[4] or ""


        payment_history_table.insert(
            "",
            tk.END,
            values=(
                payment_id,
                display_class_date(
                    payment_date
                ),
                f"${amount:,.2f}",
                payment_method,
                notes
            )
        )


def refresh_enrollments_table():

    # Clear the visible table.

    for item in enrollment_table.get_children():

        enrollment_table.delete(
            item
        )


    connection = connect_database()

    cursor = connection.cursor()


    # -------------------------------------------------
    # JOIN + SUM
    # -------------------------------------------------
    #
    # We now combine:
    #
    # enrollments
    # students
    # class_sessions
    # payments
    #
    # GROUP BY keeps one visible row per enrollment,
    # while SUM(payments.amount) totals all of that
    # enrollment's payment transactions.

    cursor.execute(
        """
        SELECT
            enrollments.id,
            students.first_name,
            students.last_name,
            class_sessions.id,
            class_sessions.class_date,
            class_sessions.class_time,
            class_sessions.class_type,
            enrollments.payment_status,
            enrollments.certification_status,
            class_sessions.tuition,
            COALESCE(SUM(payments.amount), 0)
        FROM enrollments

        JOIN students
            ON students.id = enrollments.student_id

        JOIN class_sessions
            ON class_sessions.id = enrollments.session_id

        LEFT JOIN payments
            ON payments.enrollment_id = enrollments.id

        GROUP BY
            enrollments.id

        ORDER BY
            class_sessions.class_date ASC,
            class_sessions.class_time ASC,
            students.last_name ASC,
            students.first_name ASC
        """
    )


    enrollments = cursor.fetchall()

    connection.close()


    for enrollment in enrollments:

        enrollment_id = enrollment[0]

        first_name = enrollment[1]

        last_name = enrollment[2]

        session_id = enrollment[3]

        class_date = enrollment[4]

        class_time = enrollment[5]

        class_type = enrollment[6]

        payment_status = enrollment[7]

        certification_status = enrollment[8]

        tuition = float(
            enrollment[9]
        )

        total_paid = float(
            enrollment[10] or 0
        )

        balance_due = max(
            tuition - total_paid,
            0
        )


        enrollment_table.insert(
            "",
            tk.END,
            values=(
                enrollment_id,
                f"{first_name} {last_name}",
                session_id,
                display_class_date(
                    class_date
                ),
                display_class_time(
                    class_time
                ),
                class_type,
                f"${tuition:,.2f}",
                f"${total_paid:,.2f}",
                f"${balance_due:,.2f}",
                payment_status,
                certification_status
            )
        )


    if len(enrollments) == 1:

        enrollment_count_label.config(
            text="1 enrollment"
        )

    else:

        enrollment_count_label.config(
            text=f"{len(enrollments)} enrollments"
        )


    refresh_enrollment_choices()


def enroll_student():

    student_id = student_enrollment_lookup.get(
        student_enrollment_combo.get()
    )


    session_id = session_enrollment_lookup.get(
        session_enrollment_combo.get()
    )


    if student_id is None or session_id is None:

        messagebox.showwarning(
            "Choose Student and Class",
            (
                "Please choose both a student "
                "and a scheduled class session."
            )
        )

        return


    # New enrollments always begin with $0 recorded,
    # so payment status begins as Not Paid.

    payment_status = "Not Paid"


    certification_status = (
        enrollment_certification_combo.get()
    )


    connection = connect_database()

    cursor = connection.cursor()


    cursor.execute(
        """
        SELECT id
        FROM enrollments
        WHERE
            student_id = ?
            AND session_id = ?
        """,
        (
            student_id,
            session_id
        )
    )


    existing_enrollment = (
        cursor.fetchone()
    )


    if existing_enrollment is not None:

        connection.close()


        messagebox.showwarning(
            "Already Enrolled",
            (
                "That student is already enrolled "
                "in this class session."
            )
        )

        return


    cursor.execute(
        """
        SELECT capacity
        FROM class_sessions
        WHERE id = ?
        """,
        (
            session_id,
        )
    )


    class_record = cursor.fetchone()


    if class_record is None:

        connection.close()


        messagebox.showerror(
            "Class Not Found",
            "That scheduled class could not be found."
        )

        refresh_enrollment_choices()

        return


    capacity = class_record[0]


    cursor.execute(
        """
        SELECT COUNT(*)
        FROM enrollments
        WHERE session_id = ?
        """,
        (
            session_id,
        )
    )


    enrolled_count = (
        cursor.fetchone()[0]
    )


    if enrolled_count >= capacity:

        connection.close()


        messagebox.showwarning(
            "Class Full",
            (
                "This class has reached its capacity. "
                "No additional students can be enrolled."
            )
        )

        refresh_enrollment_choices()

        return


    try:

        cursor.execute(
            """
            INSERT INTO enrollments (
                student_id,
                session_id,
                payment_status,
                certification_status,
                enrolled_at
            )
            VALUES (?, ?, ?, ?, ?)
            """,
            (
                student_id,
                session_id,
                payment_status,
                certification_status,
                datetime.now().strftime(
                    "%Y-%m-%d %H:%M"
                )
            )
        )


        connection.commit()

        connection.close()


    except sqlite3.IntegrityError:

        connection.close()


        messagebox.showwarning(
            "Enrollment Not Saved",
            (
                "SQLite blocked this enrollment because "
                "the student/class relationship already exists "
                "or no longer points to a valid record."
            )
        )

        return


    clear_enrollment_form()

    refresh_enrollments_table()

    refresh_class_sessions_table()


    messagebox.showinfo(
        "Student Enrolled",
        "The student was enrolled successfully."
    )


def load_selected_enrollment():

    global loaded_enrollment_id


    selected_items = (
        enrollment_table.selection()
    )


    if not selected_items:

        messagebox.showwarning(
            "Select an Enrollment",
            "Please select an enrollment first."
        )

        return


    selected_item = (
        selected_items[0]
    )


    selected_values = enrollment_table.item(
        selected_item,
        "values"
    )


    enrollment_id = int(
        selected_values[0]
    )


    connection = connect_database()

    cursor = connection.cursor()


    cursor.execute(
        """
        SELECT
            student_id,
            session_id,
            payment_status,
            certification_status
        FROM enrollments
        WHERE id = ?
        """,
        (
            enrollment_id,
        )
    )


    enrollment = cursor.fetchone()

    connection.close()


    if enrollment is None:

        messagebox.showerror(
            "Enrollment Not Found",
            "That enrollment could not be found."
        )

        refresh_enrollments_table()

        return


    loaded_enrollment_id = (
        enrollment_id
    )


    student_id = enrollment[0]

    session_id = enrollment[1]


    refresh_enrollment_choices()


    for display_text, lookup_id in student_enrollment_lookup.items():

        if lookup_id == student_id:

            student_enrollment_combo.set(
                display_text
            )

            break


    for display_text, lookup_id in session_enrollment_lookup.items():

        if lookup_id == session_id:

            session_enrollment_combo.set(
                display_text
            )

            break


    enrollment_payment_combo.set(
        enrollment[2]
    )


    enrollment_certification_combo.set(
        enrollment[3]
    )


    student_enrollment_combo.config(
        state="disabled"
    )


    session_enrollment_combo.config(
        state="disabled"
    )


    enrollment_edit_status_label.config(
        text=(
            f"Editing Enrollment "
            f"#{enrollment_id}"
        )
    )


    refresh_payment_section()


def update_loaded_enrollment():

    global loaded_enrollment_id


    if loaded_enrollment_id is None:

        messagebox.showwarning(
            "No Loaded Enrollment",
            (
                "Load an enrollment from the table first."
            )
        )

        return


    certification_status = (
        enrollment_certification_combo.get()
    )


    connection = connect_database()

    cursor = connection.cursor()


    # Payment status is no longer manually edited.
    # Actual payment transactions are the source of
    # truth. This button now updates certification.

    cursor.execute(
        """
        UPDATE enrollments
        SET certification_status = ?
        WHERE id = ?
        """,
        (
            certification_status,
            loaded_enrollment_id
        )
    )


    connection.commit()

    connection.close()


    updated_id = loaded_enrollment_id


    refresh_enrollments_table()

    refresh_selected_class_roster()

    refresh_payment_section()


    messagebox.showinfo(
        "Enrollment Updated",
        f"Enrollment #{updated_id} certification was updated."
    )


def add_payment():

    if loaded_enrollment_id is None:

        messagebox.showwarning(
            "Load an Enrollment",
            (
                "Load an enrollment first so Lux knows "
                "which student/class relationship should receive the payment."
            )
        )

        return


    try:

        amount = float(
            payment_amount_entry.get()
        )


        if amount <= 0:
            raise ValueError


        payment_date = normalize_class_date(
            payment_date_entry.get()
        )


    except ValueError:

        messagebox.showerror(
            "Invalid Payment",
            (
                "Enter a payment amount greater than zero "
                "and a date like 10/12/2026."
            )
        )

        return


    payment_method = (
        payment_method_combo.get().strip()
    )


    notes = (
        payment_notes_entry.get().strip()
    )


    connection = connect_database()

    cursor = connection.cursor()


    money = get_enrollment_money(
        loaded_enrollment_id,
        cursor
    )


    if money is None:

        connection.close()

        messagebox.showerror(
            "Enrollment Not Found",
            "That enrollment no longer exists."
        )

        clear_enrollment_form()

        return


    balance_due = money[2]


    # For this version we block accidental overpayments.
    # Refunds/credits deserve their own future workflow
    # instead of silently creating negative balances.

    if amount > balance_due + 0.005:

        connection.close()

        messagebox.showwarning(
            "Payment Too Large",
            (
                f"The remaining balance is ${balance_due:,.2f}. "
                "This payment would exceed the tuition balance."
            )
        )

        return


    try:

        cursor.execute(
            """
            INSERT INTO payments (
                enrollment_id,
                amount,
                payment_date,
                payment_method,
                notes,
                created_at
            )
            VALUES (?, ?, ?, ?, ?, ?)
            """,
            (
                loaded_enrollment_id,
                amount,
                payment_date,
                payment_method,
                notes,
                datetime.now().strftime(
                    "%Y-%m-%d %H:%M"
                )
            )
        )


        sync_enrollment_payment_status(
            loaded_enrollment_id,
            connection
        )


        connection.commit()

        connection.close()


    except sqlite3.IntegrityError:

        connection.close()

        messagebox.showerror(
            "Payment Not Saved",
            "SQLite rejected that payment record."
        )

        return


    payment_amount_entry.delete(
        0,
        tk.END
    )


    payment_notes_entry.delete(
        0,
        tk.END
    )


    refresh_payment_section()

    refresh_enrollments_table()

    refresh_selected_class_roster()


    # Reload the current enrollment so the automatic
    # payment-status dropdown reflects the new total.

    connection = connect_database()

    cursor = connection.cursor()

    cursor.execute(
        """
        SELECT payment_status
        FROM enrollments
        WHERE id = ?
        """,
        (
            loaded_enrollment_id,
        )
    )

    status_row = cursor.fetchone()

    connection.close()


    if status_row is not None:

        enrollment_payment_combo.set(
            status_row[0]
        )


    messagebox.showinfo(
        "Payment Recorded",
        f"${amount:,.2f} was added to the payment history."
    )


def delete_selected_payment():

    if loaded_enrollment_id is None:

        messagebox.showwarning(
            "Load an Enrollment",
            "Load an enrollment first."
        )

        return


    selected_items = (        payment_history_table.selection()
    )


    if not selected_items:

        messagebox.showwarning(
            "Select a Payment",
            "Please select a payment transaction first."
        )

        return


    selected_item = selected_items[0]


    selected_values = payment_history_table.item(
        selected_item,
        "values"
    )


    payment_id = int(
        selected_values[0]
    )


    amount_text = selected_values[2]


    confirmed = messagebox.askyesno(
        "Delete Payment",
        (
            f"Delete payment #{payment_id} "
            f"for {amount_text}?"
        )
    )


    if not confirmed:
        return


    connection = connect_database()

    cursor = connection.cursor()


    cursor.execute(
        """
        DELETE FROM payments
        WHERE
            id = ?
            AND enrollment_id = ?
        """,
        (
            payment_id,
            loaded_enrollment_id
        )
    )


    sync_enrollment_payment_status(
        loaded_enrollment_id,
        connection
    )


    connection.commit()

    connection.close()


    refresh_payment_section()

    refresh_enrollments_table()

    refresh_selected_class_roster()


    connection = connect_database()

    cursor = connection.cursor()

    cursor.execute(
        """
        SELECT payment_status
        FROM enrollments
        WHERE id = ?
        """,
        (
            loaded_enrollment_id,
        )
    )

    status_row = cursor.fetchone()

    connection.close()


    if status_row is not None:

        enrollment_payment_combo.set(
            status_row[0]
        )


    messagebox.showinfo(
        "Payment Deleted",
        "The selected payment transaction was deleted."
    )


def remove_selected_enrollment():

    global loaded_enrollment_id


    selected_items = (
        enrollment_table.selection()
    )


    if not selected_items:

        messagebox.showwarning(
            "Select an Enrollment",
            "Please select an enrollment first."
        )

        return


    selected_item = (
        selected_items[0]
    )


    selected_values = enrollment_table.item(
        selected_item,
        "values"
    )


    enrollment_id = int(
        selected_values[0]
    )


    student_name = (
        selected_values[1]
    )


    class_date = (
        selected_values[3]
    )


    class_type = (
        selected_values[5]
    )


    confirmed = messagebox.askyesno(
        "Remove Enrollment",
        (
            f"Remove {student_name} from "
            f"{class_type} on {class_date}?"
        )
    )


    if not confirmed:
        return


    connection = connect_database()

    cursor = connection.cursor()


    cursor.execute(
        """
        SELECT COUNT(*)
        FROM payments
        WHERE enrollment_id = ?
        """,
        (
            enrollment_id,
        )
    )


    payment_count = cursor.fetchone()[0]


    if payment_count > 0:

        connection.close()

        messagebox.showwarning(
            "Payment History Exists",
            (
                "This enrollment has recorded payment history. "
                "Lux will not erase the enrollment while financial "
                "transactions are attached to it."
            )
        )

        return


    cursor.execute(
        """
        DELETE FROM enrollments
        WHERE id = ?
        """,
        (
            enrollment_id,
        )
    )


    connection.commit()

    connection.close()


    if loaded_enrollment_id == enrollment_id:

        clear_enrollment_form()


    refresh_enrollments_table()

    refresh_class_sessions_table()


    messagebox.showinfo(
        "Enrollment Removed",
        (
            f"{student_name} was removed "
            f"from that class session."
        )
    )



# =====================================================
# DASHBOARD / COMMAND CENTER SUMMARY
# =====================================================

def refresh_dashboard():

    # The Dashboard is a READ-ONLY summary.
    #
    # It does not create duplicate totals or store a second
    # copy of business data. Every number is calculated fresh
    # from the real SQLite tables each time the Dashboard opens.
    #
    # This is a very important database idea:
    #
    # STORE the facts once.
    # QUERY the facts many different ways.

    if "dashboard_upcoming_table" not in globals():
        return


    today = datetime.now().strftime(
        "%Y-%m-%d"
    )


    connection = connect_database()

    cursor = connection.cursor()


    # -------------------------------------------------
    # SUMMARY CARDS
    # -------------------------------------------------
    #
    # payment_totals first groups every transaction by
    # enrollment_id.
    #
    # The second part joins those totals to upcoming
    # enrollment records.
    #
    # Then SQLite gives us:
    #
    # COUNT() = how many active enrollments
    # SUM()   = expected / collected / outstanding money
    #
    # CASE lets SQL make a decision inside a query.

    cursor.execute(
        """
        WITH payment_totals AS (
            SELECT
                enrollment_id,
                COALESCE(SUM(amount), 0) AS total_paid
            FROM payments
            GROUP BY enrollment_id
        ),
        active_enrollments AS (
            SELECT
                enrollments.id,
                enrollments.student_id,
                class_sessions.tuition,
                COALESCE(payment_totals.total_paid, 0) AS total_paid
            FROM enrollments

            JOIN class_sessions
                ON class_sessions.id = enrollments.session_id

            LEFT JOIN payment_totals
                ON payment_totals.enrollment_id = enrollments.id

            WHERE class_sessions.class_date >= ?
              AND class_sessions.status != 'Cancelled'
        )

        SELECT
            COUNT(*),
            COALESCE(SUM(tuition), 0),
            COALESCE(SUM(total_paid), 0),
            COALESCE(
                SUM(
                    CASE
                        WHEN tuition > total_paid
                        THEN tuition - total_paid
                        ELSE 0
                    END
                ),
                0
            ),
            COUNT(
                DISTINCT CASE
                    WHEN tuition > total_paid
                    THEN student_id
                END
            )
        FROM active_enrollments
        """,
        (
            today,
        )
    )


    summary = cursor.fetchone()


    active_enrollments = int(
        summary[0] or 0
    )

    expected_tuition = float(
        summary[1] or 0
    )

    collected_revenue = float(
        summary[2] or 0
    )

    outstanding_balance = float(
        summary[3] or 0
    )

    students_with_balance = int(
        summary[4] or 0
    )


    # -------------------------------------------------
    # UPCOMING CLASSES
    # -------------------------------------------------
    #
    # GROUP BY gives us ONE dashboard row per class.
    # COUNT(enrollments.id) tells us the real roster size.
    # SUM(payment_totals.total_paid) tells us how much
    # money has actually been recorded for that class.

    cursor.execute(
        """
        WITH payment_totals AS (
            SELECT
                enrollment_id,
                COALESCE(SUM(amount), 0) AS total_paid
            FROM payments
            GROUP BY enrollment_id
        )

        SELECT
            class_sessions.id,
            class_sessions.class_date,
            class_sessions.class_time,
            class_sessions.class_type,
            class_sessions.capacity,
            class_sessions.tuition,
            class_sessions.status,
            COUNT(enrollments.id) AS enrolled_count,
            COALESCE(SUM(payment_totals.total_paid), 0) AS collected
        FROM class_sessions

        LEFT JOIN enrollments
            ON enrollments.session_id = class_sessions.id

        LEFT JOIN payment_totals
            ON payment_totals.enrollment_id = enrollments.id

        WHERE class_sessions.class_date >= ?
          AND class_sessions.status != 'Cancelled'

        GROUP BY
            class_sessions.id

        ORDER BY
            class_sessions.class_date ASC,
            class_sessions.class_time ASC,
            class_sessions.id ASC

        LIMIT 10
        """,
        (
            today,
        )
    )


    upcoming_classes = cursor.fetchall()


    # -------------------------------------------------
    # NEEDS ATTENTION
    # -------------------------------------------------
    #
    # These are upcoming enrollments whose recorded
    # payments are still smaller than tuition.
    #
    # We calculate the balance from the ledger rather
    # than trusting a manually typed status label.

    cursor.execute(
        """
        WITH payment_totals AS (
            SELECT
                enrollment_id,
                COALESCE(SUM(amount), 0) AS total_paid
            FROM payments
            GROUP BY enrollment_id
        )

        SELECT
            enrollments.id,
            students.first_name,
            students.last_name,
            class_sessions.class_type,
            class_sessions.class_date,
            class_sessions.class_time,
            class_sessions.tuition,
            COALESCE(payment_totals.total_paid, 0) AS total_paid
        FROM enrollments

        JOIN students
            ON students.id = enrollments.student_id

        JOIN class_sessions
            ON class_sessions.id = enrollments.session_id

        LEFT JOIN payment_totals
            ON payment_totals.enrollment_id = enrollments.id

        WHERE class_sessions.class_date >= ?
          AND class_sessions.status != 'Cancelled'
          AND class_sessions.tuition > COALESCE(payment_totals.total_paid, 0)

        ORDER BY
            class_sessions.class_date ASC,
            class_sessions.class_time ASC,
            students.last_name ASC,
            students.first_name ASC

        LIMIT 12
        """,
        (
            today,
        )
    )


    attention_rows = cursor.fetchall()


    connection.close()


    # -------------------------------------------------
    # UPDATE SUMMARY CARDS
    # -------------------------------------------------

    dashboard_enrollment_value.config(
        text=f"{active_enrollments:,}"
    )


    dashboard_expected_value.config(
        text=f"${expected_tuition:,.2f}"
    )


    dashboard_collected_value.config(
        text=f"${collected_revenue:,.2f}"
    )


    dashboard_outstanding_value.config(
        text=f"${outstanding_balance:,.2f}"
    )


    dashboard_balance_count_value.config(
        text=f"{students_with_balance:,}"
    )


    # -------------------------------------------------
    # NEXT CLASS CARD
    # -------------------------------------------------

    if upcoming_classes:

        next_class = upcoming_classes[0]

        next_session_id = next_class[0]
        next_date = next_class[1]
        next_time = next_class[2]
        next_type = next_class[3]
        next_capacity = int(next_class[4])
        next_tuition = float(next_class[5])
        next_status = next_class[6]
        next_enrolled = int(next_class[7])
        next_collected = float(next_class[8] or 0)

        next_seats = max(
            next_capacity - next_enrolled,
            0
        )

        next_expected = (
            next_tuition * next_enrolled
        )

        next_outstanding = max(
            next_expected - next_collected,
            0
        )


        dashboard_next_class_title.config(
            text=(
                f"Session #{next_session_id} • {next_type} • "
                f"{next_status}"
            )
        )


        dashboard_next_class_schedule.config(
            text=(
                f"{display_class_date(next_date)} at "
                f"{display_class_time(next_time)}"
            )
        )


        dashboard_next_class_details.config(
            text=(
                f"{next_enrolled}/{next_capacity} enrolled • "
                f"{next_seats} seats left • "
                f"Expected ${next_expected:,.2f} • "
                f"Collected ${next_collected:,.2f} • "
                f"Outstanding ${next_outstanding:,.2f}"
            )
        )

    else:

        dashboard_next_class_title.config(
            text="No upcoming class sessions"
        )

        dashboard_next_class_schedule.config(
            text=""
        )

        dashboard_next_class_details.config(
            text="Schedule a class in the Calculator when Lux is ready."
        )


    # -------------------------------------------------
    # UPCOMING CLASS TABLE
    # -------------------------------------------------

    for item in dashboard_upcoming_table.get_children():

        dashboard_upcoming_table.delete(
            item
        )


    for class_row in upcoming_classes:

        session_id = class_row[0]
        class_date = class_row[1]
        class_time = class_row[2]
        class_type = class_row[3]
        capacity = int(class_row[4])
        tuition = float(class_row[5])
        status = class_row[6]
        enrolled = int(class_row[7])
        collected = float(class_row[8] or 0)

        seats_left = max(
            capacity - enrolled,
            0
        )

        expected = (
            tuition * enrolled
        )

        outstanding = max(
            expected - collected,
            0
        )


        dashboard_upcoming_table.insert(
            "",
            tk.END,
            values=(
                session_id,
                display_class_date(
                    class_date
                ),
                display_class_time(
                    class_time
                ),
                class_type,
                f"{enrolled}/{capacity}",
                seats_left,
                f"${collected:,.2f}",
                f"${outstanding:,.2f}",
                status
            )
        )


    # -------------------------------------------------
    # NEEDS-ATTENTION TABLE
    # -------------------------------------------------

    for item in dashboard_attention_table.get_children():

        dashboard_attention_table.delete(
            item
        )


    for attention in attention_rows:

        enrollment_id = attention[0]
        first_name = attention[1]
        last_name = attention[2]
        class_type = attention[3]
        class_date = attention[4]
        class_time = attention[5]
        tuition = float(attention[6])
        total_paid = float(attention[7] or 0)

        balance = max(
            tuition - total_paid,
            0
        )


        dashboard_attention_table.insert(
            "",
            tk.END,
            values=(
                enrollment_id,
                f"{first_name} {last_name}",
                class_type,
                display_class_date(
                    class_date
                ),
                display_class_time(
                    class_time
                ),
                f"${total_paid:,.2f}",
                f"${balance:,.2f}"
            )
        )


    if attention_rows:

        dashboard_attention_status.config(
            text=(
                f"{len(attention_rows)} upcoming enrollment"
                f"{'s' if len(attention_rows) != 1 else ''} "
                f"with a balance due"
            )
        )

    else:

        dashboard_attention_status.config(
            text="No upcoming balances need attention 🎉"
        )


    dashboard_last_updated_label.config(
        text=(
            "Dashboard refreshed "
            + datetime.now().strftime(
                "%m/%d/%Y %I:%M %p"
            ).lstrip("0")
        )
    )


def dashboard_view_selected_class():

    # A dashboard is useful only if it can take the user
    # directly to the underlying record.
    #
    # This function reads the selected Dashboard session ID,
    # switches to Classes, finds that SAME permanent SQLite ID,
    # and opens its live roster.

    selected_items = (
        dashboard_upcoming_table.selection()
    )


    if selected_items:

        selected_values = dashboard_upcoming_table.item(
            selected_items[0],
            "values"
        )

        session_id = int(
            selected_values[0]
        )

    else:

        children = dashboard_upcoming_table.get_children()

        if not children:

            messagebox.showwarning(
                "No Upcoming Class",
                "There is no upcoming class to open."
            )

            return


        first_values = dashboard_upcoming_table.item(
            children[0],
            "values"
        )

        session_id = int(
            first_values[0]
        )


    notebook.select(
        classes_tab
    )


    refresh_class_sessions_table()


    for item in class_sessions_table.get_children():

        values = class_sessions_table.item(
            item,
            "values"
        )

        if values and int(values[0]) == session_id:

            class_sessions_table.selection_set(
                item
            )

            class_sessions_table.focus(
                item
            )

            class_sessions_table.see(
                item
            )

            refresh_selected_class_roster()

            return


    messagebox.showwarning(
        "Class Not Found",
        "That class session could not be found."
    )


def dashboard_open_students():

    notebook.select(
        students_tab
    )

    clear_student_form()


def dashboard_open_enrollments():

    notebook.select(
        enrollments_tab
    )

    refresh_enrollment_choices()

    refresh_enrollments_table()


def dashboard_open_calculator():

    notebook.select(
        calculator_tab
    )


def enrollment_tab_selected(event):

    selected_tab = notebook.select()


    # Refresh the Dashboard whenever the user opens it.
    # This keeps the summary tied to the live database instead
    # of showing yesterday's numbers.

    if selected_tab == str(dashboard_tab):

        refresh_dashboard()


    if selected_tab == str(enrollments_tab):

        refresh_enrollment_choices()

        refresh_enrollments_table()

        refresh_payment_section()


    if selected_tab == str(class_readiness_tab):

        refresh_class_readiness_all()


    if selected_tab == str(supplies_tab):

        refresh_supplies_all()


    if selected_tab == str(reports_tab):

        refresh_reports_all()




# =====================================================
# v4.0 CLASS READINESS COMMAND CENTER
# =====================================================

# v4.0 does NOT create a new readiness database table.
#
# That is deliberate.
#
# "Readiness" is a CURRENT business answer calculated
# from facts already stored in:
#
# class_sessions
# enrollments
# payments
# class_supply_requirements
# supplies
#
# Store facts once. Ask the database many questions.
# This avoids duplicate totals drifting out of sync.


def refresh_class_readiness_session_choices():

    class_readiness_session_lookup.clear()


    connection = connect_database()

    cursor = connection.cursor()


    cursor.execute(
        """
        SELECT
            class_sessions.id,
            class_sessions.class_date,
            class_sessions.class_time,
            class_sessions.class_type,
            class_sessions.capacity,
            COUNT(enrollments.id)
        FROM class_sessions
        LEFT JOIN enrollments
            ON enrollments.session_id =
               class_sessions.id
        GROUP BY
            class_sessions.id,
            class_sessions.class_date,
            class_sessions.class_time,
            class_sessions.class_type,
            class_sessions.capacity
        ORDER BY
            class_sessions.class_date ASC,
            class_sessions.class_time ASC,
            class_sessions.id ASC
        """
    )


    sessions = cursor.fetchall()

    connection.close()


    display_values = []


    for session in sessions:

        display_text = (
            f"#{session[0]} | "
            f"{display_class_date(session[1])} "
            f"{display_class_time(session[2])} | "
            f"{session[3]} | "
            f"{session[5]}/{session[4]} enrolled"
        )


        class_readiness_session_lookup[
            display_text
        ] = session[0]


        display_values.append(
            display_text
        )


    current_value = (
        class_readiness_session_combo.get()
        if "class_readiness_session_combo" in globals()
        else ""
    )


    if "class_readiness_session_combo" in globals():

        class_readiness_session_combo[
            "values"
        ] = display_values


        if current_value not in display_values:

            if display_values:

                class_readiness_session_combo.set(
                    display_values[0]
                )
            else:

                class_readiness_session_combo.set(
                    ""
                )


def get_class_readiness_snapshot(
    session_id
):

    # This function gathers everything needed to answer:
    #
    # "Is this scheduled class ready to run?"
    #
    # It returns ONE Python dictionary containing
    # schedule, enrollment, money, certification,
    # and supply information.

    connection = connect_database()

    cursor = connection.cursor()


    cursor.execute(
        """
        SELECT
            id,
            class_type,
            class_date,
            class_time,
            capacity,
            tuition,
            status
        FROM class_sessions
        WHERE id = ?
        """,
        (
            session_id,
        )
    )


    session = cursor.fetchone()


    if session is None:

        connection.close()

        return None


    cursor.execute(
        """
        SELECT
            enrollments.id,
            students.id,
            students.first_name,
            students.last_name,
            enrollments.certification_status,
            COALESCE(SUM(payments.amount), 0)
        FROM enrollments
        JOIN students
            ON students.id =
               enrollments.student_id
        LEFT JOIN payments
            ON payments.enrollment_id =
               enrollments.id
        WHERE enrollments.session_id = ?
        GROUP BY
            enrollments.id,
            students.id,
            students.first_name,
            students.last_name,
            enrollments.certification_status
        ORDER BY
            students.last_name COLLATE NOCASE,
            students.first_name COLLATE NOCASE
        """,
        (
            session_id,
        )
    )


    student_rows = cursor.fetchall()


    cursor.execute(
        """
        SELECT
            supplies.id,
            supplies.item_name,
            supplies.category,
            supplies.quantity_on_hand,
            supplies.minimum_stock,
            supplies.cost_per_unit,
            class_supply_requirements.quantity_per_student,
            class_supply_requirements.fixed_quantity
        FROM class_supply_requirements
        JOIN supplies
            ON supplies.id =
               class_supply_requirements.supply_id
        WHERE class_supply_requirements.class_type = ?
        ORDER BY
            supplies.category,
            supplies.item_name
        """,
        (
            session[1],
        )
    )


    supply_rows = cursor.fetchall()


    connection.close()


    class_type = session[1]

    class_date = session[2]

    class_time = session[3]

    capacity = int(
        session[4]
    )

    tuition = float(
        session[5]
    )

    class_status = session[6]


    enrolled_count = len(
        student_rows
    )


    expected_tuition = (
        tuition * enrolled_count
    )


    students = []

    total_paid = 0

    balance_due_count = 0


    certification_counts = {
        "Not Started": 0,
        "In Progress": 0,
        "Certified": 0
    }


    for row in student_rows:

        amount_paid = float(
            row[5]
        )


        balance_due = max(
            0,
            tuition - amount_paid
        )


        total_paid += amount_paid


        if balance_due > 0.005:

            balance_due_count += 1


        if amount_paid <= 0.005:

            payment_status = (
                "Not Paid"
            )

        elif balance_due <= 0.005:

            payment_status = (
                "Paid in Full"
            )

        else:

            payment_status = (
                "Deposit Paid"
            )


        certification_status = (
            row[4]
        )


        if (
            certification_status
            not in certification_counts
        ):

            certification_counts[
                certification_status
            ] = 0


        certification_counts[
            certification_status
        ] += 1


        students.append(
            {
                "enrollment_id": row[0],
                "student_id": row[1],
                "name": (
                    f"{row[2]} {row[3]}"
                ),
                "certification": certification_status,
                "paid": amount_paid,
                "balance": balance_due,
                "payment_status": payment_status
            }
        )


    outstanding = max(
        0,
        expected_tuition - total_paid
    )


    supplies = []

    shortage_items = 0

    low_after_items = 0

    shortage_cost = 0


    for row in supply_rows:

        on_hand = float(
            row[3]
        )

        minimum_stock = float(
            row[4]
        )

        cost_per_unit = float(
            row[5]
        )

        per_student = float(
            row[6]
        )

        fixed_quantity = float(
            row[7]
        )


        required_quantity = (
            per_student * enrolled_count
        ) + fixed_quantity


        shortage = max(
            0,
            required_quantity - on_hand
        )


        remaining_after = (
            on_hand - required_quantity
        )


        if shortage > 0:

            supply_status = "SHORTAGE"

            shortage_items += 1

            shortage_cost += (
                shortage * cost_per_unit
            )


        elif remaining_after <= minimum_stock:

            supply_status = (
                "LOW AFTER CLASS"
            )

            low_after_items += 1


        else:

            supply_status = "READY"


        supplies.append(
            {
                "id": row[0],
                "name": row[1],
                "category": row[2],
                "need": required_quantity,
                "have": on_hand,
                "short": shortage,
                "status": supply_status
            }
        )


    # -------------------------------------------------
    # BUILD THE CHECKLIST
    # -------------------------------------------------

    checklist = []

    critical_items = []

    warning_items = []


    today_text = datetime.now().strftime(
        "%Y-%m-%d"
    )


    if class_date < today_text:

        schedule_status = "PAST DATE"

        schedule_detail = (
            "This class date has already passed."
        )

        warning_items.append(
            "Class date is in the past"
        )

    else:

        schedule_status = "READY"

        schedule_detail = (
            f"{display_class_date(class_date)} "
            f"at {display_class_time(class_time)} • "
            f"Status: {class_status}"
        )


    checklist.append(
        (
            "Schedule",
            schedule_status,
            schedule_detail
        )
    )


    if enrolled_count == 0:

        enrollment_status = "NOT READY"

        enrollment_detail = (
            f"0/{capacity} enrolled • "
            f"No students are enrolled yet."
        )

        critical_items.append(
            "No students enrolled"
        )


    elif enrolled_count > capacity:

        over_capacity = (
            enrolled_count - capacity
        )

        enrollment_status = "NOT READY"

        enrollment_detail = (
            f"{enrolled_count}/{capacity} enrolled • "
            f"{over_capacity} over capacity."
        )

        critical_items.append(
            "Enrollment exceeds capacity"
        )


    elif enrolled_count == capacity:

        enrollment_status = "READY"

        enrollment_detail = (
            f"{enrolled_count}/{capacity} enrolled • "
            f"Class is full."
        )


    else:

        seats_left = (
            capacity - enrolled_count
        )

        enrollment_status = "READY"

        enrollment_detail = (
            f"{enrolled_count}/{capacity} enrolled • "
            f"{seats_left} seat(s) still available."
        )


    checklist.append(
        (
            "Enrollment",
            enrollment_status,
            enrollment_detail
        )
    )


    if balance_due_count == 0:

        payment_status = "READY"

        payment_detail = (
            f"${total_paid:,.2f} collected • "
            f"No outstanding tuition."
        )

    else:

        payment_status = "FOLLOW UP"

        payment_detail = (
            f"${total_paid:,.2f} of "
            f"${expected_tuition:,.2f} collected • "
            f"{balance_due_count} student(s) "
            f"have a balance • "
            f"${outstanding:,.2f} outstanding."
        )

        warning_items.append(
            "Outstanding student balances"
        )


    checklist.append(
        (
            "Payments",
            payment_status,
            payment_detail
        )
    )


    if not supply_rows:

        supplies_status = "NOT CONFIGURED"

        supplies_detail = (
            f"No supply requirements are configured "
            f"for {class_type}."
        )

        warning_items.append(
            "Supply requirements are not configured"
        )


    elif shortage_items > 0:

        supplies_status = "NOT READY"

        supplies_detail = (
            f"{shortage_items} shortage item(s) • "
            f"Approx. ${shortage_cost:,.2f} "
            f"needed to cover shortages."
        )

        critical_items.append(
            "Supply shortage"
        )


    elif low_after_items > 0:

        supplies_status = "READY • WATCH STOCK"

        supplies_detail = (
            f"All required supplies are available • "
            f"{low_after_items} item(s) will be "
            f"at/below minimum stock afterward."
        )

        warning_items.append(
            "Inventory will run low after class"
        )


    else:

        supplies_status = "READY"

        supplies_detail = (
            "All configured class supplies are available."
        )


    checklist.append(
        (
            "Supplies",
            supplies_status,
            supplies_detail
        )
    )


    not_started = certification_counts.get(
        "Not Started",
        0
    )

    in_progress = certification_counts.get(
        "In Progress",
        0
    )

    certified = certification_counts.get(
        "Certified",
        0
    )


    certification_detail = (
        f"{not_started} not started • "
        f"{in_progress} in progress • "
        f"{certified} certified"
    )


    # Certification is intentionally INFORMATIONAL for
    # an upcoming class. "Not Started" is normal before
    # training begins, so it should not falsely mark an
    # otherwise-ready class as broken.

    certification_status = "TRACKING"


    if (
        class_date < today_text
        and (not_started + in_progress) > 0
    ):

        certification_status = (
            "FOLLOW UP"
        )

        warning_items.append(
            "Past class has incomplete certification records"
        )


    checklist.append(
        (
            "Certification",
            certification_status,
            certification_detail
        )
    )


    if critical_items:

        overall_status = "NOT READY"

        overall_detail = (
            " • ".join(
                critical_items
            )
        )


    elif warning_items:

        overall_status = (
            "READY • FOLLOW-UP NEEDED"
        )

        overall_detail = (
            " • ".join(
                warning_items
            )
        )


    else:

        overall_status = "READY TO RUN"

        overall_detail = (
            "Enrollment, payments, and supplies "
            "have no current blockers."
        )


    return {
        "session_id": session_id,
        "class_type": class_type,
        "class_date": class_date,
        "class_time": class_time,
        "capacity": capacity,
        "tuition": tuition,
        "class_status": class_status,
        "enrolled": enrolled_count,
        "seats_left": max(
            0,
            capacity - enrolled_count
        ),
        "expected": expected_tuition,
        "paid": total_paid,
        "outstanding": outstanding,
        "balance_due_count": balance_due_count,
        "cert_counts": certification_counts,
        "students": students,
        "supplies": supplies,
        "shortage_items": shortage_items,
        "low_after_items": low_after_items,
        "shortage_cost": shortage_cost,
        "checklist": checklist,
        "overall_status": overall_status,
        "overall_detail": overall_detail
    }


def refresh_class_readiness(
    event=None
):

    if (
        "class_readiness_checklist_table"
        not in globals()
    ):

        return


    for table in (
        class_readiness_checklist_table,
        class_readiness_students_table,
        class_readiness_supplies_table
    ):

        for item in table.get_children():

            table.delete(
                item
            )


    session_id = (
        class_readiness_session_lookup.get(
            class_readiness_session_combo.get()
        )
    )


    if session_id is None:

        class_readiness_title_label.config(
            text="No scheduled class selected"
        )

        class_readiness_overall_label.config(
            text=""
        )

        class_readiness_overall_detail.config(
            text=""
        )

        return


    snapshot = get_class_readiness_snapshot(
        session_id
    )


    if snapshot is None:

        class_readiness_title_label.config(
            text="That class could not be found."
        )

        return


    class_readiness_title_label.config(
        text=(
            f"Session #{snapshot['session_id']} • "
            f"{snapshot['class_type']} • "
            f"{display_class_date(snapshot['class_date'])} "
            f"at {display_class_time(snapshot['class_time'])}"
        )
    )


    class_readiness_overall_label.config(
        text=snapshot[
            "overall_status"
        ]
    )


    class_readiness_overall_detail.config(
        text=snapshot[
            "overall_detail"
        ]
    )


    # -------------------------------------------------
    # SUMMARY CARDS
    # -------------------------------------------------

    class_readiness_enrollment_value.config(
        text=(
            f"{snapshot['enrolled']}/"
            f"{snapshot['capacity']}"
        )
    )


    class_readiness_enrollment_detail.config(
        text=(
            f"{snapshot['seats_left']} seat(s) left"
        )
    )


    class_readiness_payment_value.config(
        text=(
            f"${snapshot['paid']:,.2f}"
        )
    )


    class_readiness_payment_detail.config(
        text=(
            f"${snapshot['outstanding']:,.2f} outstanding • "
            f"{snapshot['balance_due_count']} balance(s)"
        )
    )


    if snapshot["shortage_items"] > 0:

        supply_value_text = (
            f"{snapshot['shortage_items']} SHORT"
        )

    elif not snapshot["supplies"]:

        supply_value_text = (
            "NOT SET"
        )

    else:

        supply_value_text = "READY"


    class_readiness_supply_value.config(
        text=supply_value_text
    )


    class_readiness_supply_detail.config(
        text=(
            f"{snapshot['low_after_items']} low after class • "
            f"${snapshot['shortage_cost']:,.2f} shortage cost"
        )
    )


    cert_counts = snapshot[
        "cert_counts"
    ]


    class_readiness_cert_value.config(
        text=(
            f"{cert_counts.get('Certified', 0)} "
            f"Certified"
        )
    )


    class_readiness_cert_detail.config(
        text=(
            f"{cert_counts.get('Not Started', 0)} not started • "
            f"{cert_counts.get('In Progress', 0)} in progress"
        )
    )


    # -------------------------------------------------
    # CHECKLIST
    # -------------------------------------------------

    for area, status, detail in snapshot[
        "checklist"
    ]:

        if status in (
            "NOT READY",
        ):

            tag_name = "critical"

        elif status in (
            "FOLLOW UP",
            "NOT CONFIGURED",
            "READY • WATCH STOCK",
            "PAST DATE"
        ):

            tag_name = "warning"

        elif status == "TRACKING":

            tag_name = "info"

        else:

            tag_name = "ready"


        class_readiness_checklist_table.insert(
            "",
            tk.END,
            values=(
                area,
                status,
                detail
            ),
            tags=(
                tag_name,
            )
        )


    # -------------------------------------------------
    # STUDENT READINESS
    # -------------------------------------------------

    for student in snapshot[
        "students"
    ]:

        if student["balance"] <= 0.005:

            student_status = "CLEAR"

            tag_name = "ready"

        else:

            student_status = "BALANCE DUE"

            tag_name = "warning"


        class_readiness_students_table.insert(
            "",
            tk.END,
            values=(
                student["student_id"],
                student["name"],
                f"${snapshot['tuition']:,.2f}",
                f"${student['paid']:,.2f}",
                f"${student['balance']:,.2f}",
                student["payment_status"],
                student["certification"],
                student_status
            ),
            tags=(
                tag_name,
            )
        )


    # -------------------------------------------------
    # SUPPLY READINESS
    # -------------------------------------------------

    for supply in snapshot[
        "supplies"
    ]:

        if supply["status"] == "SHORTAGE":

            tag_name = "critical"

        elif supply["status"] == "LOW AFTER CLASS":

            tag_name = "warning"

        else:

            tag_name = "ready"


        class_readiness_supplies_table.insert(
            "",
            tk.END,
            values=(
                supply["name"],
                supply["category"],
                format_supply_quantity(
                    supply["need"]
                ),
                format_supply_quantity(
                    supply["have"]
                ),
                format_supply_quantity(
                    supply["short"]
                ),
                supply["status"]
            ),
            tags=(
                tag_name,
            )
        )


    class_readiness_updated_label.config(
        text=(
            "Readiness recalculated "
            + datetime.now().strftime(
                "%m/%d/%Y %I:%M %p"
            ).lstrip("0")
        )
    )


def refresh_class_readiness_all():

    refresh_class_readiness_session_choices()

    refresh_class_readiness()


def set_class_readiness_session(
    session_id
):

    refresh_class_readiness_session_choices()


    for display_text, stored_id in (
        class_readiness_session_lookup.items()
    ):

        if stored_id == session_id:

            class_readiness_session_combo.set(
                display_text
            )

            refresh_class_readiness()

            return True


    return False

def open_selected_class_readiness():

    selected_items = (
        class_sessions_table.selection()
    )


    if not selected_items:

        messagebox.showwarning(
            "Select a Class",
            (
                "Please select a scheduled class "
                "before opening readiness."
            )
        )

        return


    values = class_sessions_table.item(
        selected_items[0],
        "values"
    )


    session_id = int(
        values[0]
    )


    notebook.select(
        class_readiness_tab
    )


    set_class_readiness_session(
        session_id
    )


def dashboard_open_readiness():

    selected_items = (
        dashboard_upcoming_table.selection()
    )


    if selected_items:

        values = dashboard_upcoming_table.item(
            selected_items[0],
            "values"
        )

        session_id = int(
            values[0]
        )


    else:

        children = (
            dashboard_upcoming_table.get_children()
        )


        if not children:

            notebook.select(
                class_readiness_tab
            )

            refresh_class_readiness_all()

            return


        values = dashboard_upcoming_table.item(
            children[0],
            "values"
        )

        session_id = int(
            values[0]
        )


    notebook.select(
        class_readiness_tab
    )


    set_class_readiness_session(
        session_id
    )


def class_readiness_open_roster():

    session_id = (
        class_readiness_session_lookup.get(
            class_readiness_session_combo.get()
        )
    )


    if session_id is None:

        messagebox.showwarning(
            "Select a Class",
            "Please select a class first."
        )

        return


    notebook.select(
        classes_tab
    )


    refresh_class_sessions_table()


    for item in class_sessions_table.get_children():

        values = class_sessions_table.item(
            item,
            "values"
        )


        if (
            values
            and int(values[0]) == session_id
        ):

            class_sessions_table.selection_set(
                item
            )

            class_sessions_table.focus(
                item
            )

            class_sessions_table.see(
                item
            )

            refresh_selected_class_roster()

            return


def class_readiness_open_enrollments():

    notebook.select(
        enrollments_tab
    )

    refresh_enrollment_choices()

    refresh_enrollments_table()

    refresh_payment_section()


def class_readiness_open_supplies():

    session_id = (
        class_readiness_session_lookup.get(
            class_readiness_session_combo.get()
        )
    )


    notebook.select(
        supplies_tab
    )


    refresh_supplies_all()


    if session_id is None:

        return


    for display_text, stored_id in (
        readiness_session_lookup.items()
    ):

        if stored_id == session_id:

            readiness_session_combo.set(
                display_text
            )

            refresh_supply_readiness()

            break


# =====================================================
# SUPPLIES + CLASS READINESS
# =====================================================

def format_supply_quantity(value):

    # Inventory is stored as REAL so it can support
    # whole numbers OR fractional quantities.
    #
    # We display clean whole numbers without ".0".

    number = float(value)


    if number.is_integer():

        return str(
            int(number)
        )


    return f"{number:.2f}".rstrip("0").rstrip(".")


def get_supply_stock_status(
    quantity_on_hand,
    minimum_stock
):

    quantity_on_hand = float(
        quantity_on_hand
    )

    minimum_stock = float(
        minimum_stock
    )


    if quantity_on_hand <= 0:

        return "OUT OF STOCK"


    if quantity_on_hand <= minimum_stock:

        return "LOW STOCK"


    return "OK"


def clear_supply_form():

    global loaded_supply_id


    loaded_supply_id = None


    inventory_item_entry.delete(
        0,
        tk.END
    )


    inventory_category_combo.set(
        "Disposables"
    )


    inventory_quantity_entry.delete(
        0,
        tk.END
    )

    inventory_quantity_entry.insert(
        0,
        "0"
    )


    inventory_minimum_entry.delete(
        0,
        tk.END
    )

    inventory_minimum_entry.insert(
        0,
        "0"
    )


    inventory_cost_entry.delete(
        0,
        tk.END
    )

    inventory_cost_entry.insert(
        0,
        "0"
    )


    inventory_supplier_entry.delete(
        0,
        tk.END
    )


    inventory_notes_entry.delete(
        0,
        tk.END
    )


    inventory_edit_status_label.config(
        text="Adding a new inventory item"
    )


def add_supply():

    item_name = (
        inventory_item_entry
        .get()
        .strip()
    )


    category = (
        inventory_category_combo
        .get()
        .strip()
    )


    supplier = (
        inventory_supplier_entry
        .get()
        .strip()
    )


    notes = (
        inventory_notes_entry
        .get()
        .strip()
    )


    if not item_name:

        messagebox.showwarning(
            "Missing Item Name",
            "Please enter a supply item name."
        )

        return


    try:

        quantity_on_hand = float(
            inventory_quantity_entry.get()
        )

        minimum_stock = float(
            inventory_minimum_entry.get()
        )

        cost_per_unit = float(
            inventory_cost_entry.get()
        )


        if (
            quantity_on_hand < 0
            or minimum_stock < 0
            or cost_per_unit < 0
        ):

            raise ValueError


    except ValueError:

        messagebox.showerror(
            "Inventory Error",
            (
                "Quantity, minimum stock, and cost "
                "must be valid non-negative numbers."
            )
        )

        return


    connection = connect_database()

    cursor = connection.cursor()


    try:

        cursor.execute(
            """
            INSERT INTO supplies (
                item_name,
                category,
                quantity_on_hand,
                minimum_stock,
                cost_per_unit,
                supplier,
                notes,
                created_at
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                item_name,
                category,
                quantity_on_hand,
                minimum_stock,
                cost_per_unit,
                supplier,
                notes,
                datetime.now().strftime(
                    "%Y-%m-%d %H:%M"
                )
            )
        )


        connection.commit()


    except sqlite3.IntegrityError:

        connection.close()


        messagebox.showwarning(
            "Duplicate Supply",
            (
                "That inventory item already exists. "
                "Load the existing item if you want to update it."
            )
        )

        return


    connection.close()


    clear_supply_form()

    refresh_supplies_all()


    messagebox.showinfo(
        "Supply Added",
        f"{item_name} was added to inventory."
    )


def refresh_supplies_table():

    if "supplies_table" not in globals():

        return


    for item in supplies_table.get_children():

        supplies_table.delete(
            item
        )


    connection = connect_database()

    cursor = connection.cursor()


    cursor.execute(
        """
        SELECT
            id,
            item_name,
            category,
            quantity_on_hand,
            minimum_stock,
            cost_per_unit,
            supplier
        FROM supplies
        ORDER BY
            category,
            item_name
        """
    )


    supplies = cursor.fetchall()


    connection.close()


    low_count = 0


    for supply in supplies:

        status = get_supply_stock_status(
            supply[3],
            supply[4]
        )


        if status != "OK":

            low_count += 1


        supplies_table.insert(
            "",
            tk.END,
            values=(
                supply[0],
                supply[1],
                supply[2],
                format_supply_quantity(
                    supply[3]
                ),
                format_supply_quantity(
                    supply[4]
                ),
                status,
                f"${supply[5]:,.2f}",
                supply[6] or ""
            )
        )


    if len(supplies) == 1:

        item_text = "1 inventory item"

    else:

        item_text = (
            f"{len(supplies)} inventory items"
        )


    if low_count == 1:

        low_text = "1 needs attention"

    else:

        low_text = (
            f"{low_count} need attention"
        )


    supplies_count_label.config(
        text=f"{item_text} • {low_text}"
    )


def load_selected_supply():

    global loaded_supply_id


    selected_items = (
        supplies_table.selection()
    )


    if not selected_items:

        messagebox.showwarning(
            "Select an Item",
            "Please select an inventory item first."
        )

        return


    values = supplies_table.item(
        selected_items[0],
        "values"
    )


    supply_id = int(
        values[0]
    )


    connection = connect_database()

    cursor = connection.cursor()


    cursor.execute(
        """
        SELECT
            item_name,
            category,
            quantity_on_hand,
            minimum_stock,
            cost_per_unit,
            supplier,
            notes
        FROM supplies
        WHERE id = ?
        """,
        (
            supply_id,
        )
    )


    supply = cursor.fetchone()


    connection.close()


    if supply is None:

        messagebox.showerror(
            "Supply Not Found",
            "That inventory item could not be found."
        )

        refresh_supplies_all()

        return


    loaded_supply_id = (
        supply_id
    )


    inventory_item_entry.delete(
        0,
        tk.END
    )

    inventory_item_entry.insert(
        0,
        supply[0]
    )


    inventory_category_combo.set(
        supply[1]
    )


    inventory_quantity_entry.delete(
        0,
        tk.END
    )

    inventory_quantity_entry.insert(
        0,
        format_supply_quantity(
            supply[2]
        )
    )


    inventory_minimum_entry.delete(
        0,
        tk.END
    )

    inventory_minimum_entry.insert(
        0,
        format_supply_quantity(
            supply[3]
        )
    )


    inventory_cost_entry.delete(
        0,
        tk.END
    )

    inventory_cost_entry.insert(
        0,
        f"{supply[4]:.2f}"
    )


    inventory_supplier_entry.delete(
        0,
        tk.END
    )

    inventory_supplier_entry.insert(
        0,
        supply[5] or ""
    )


    inventory_notes_entry.delete(
        0,
        tk.END
    )

    inventory_notes_entry.insert(
        0,
        supply[6] or ""
    )


    inventory_edit_status_label.config(
        text=(
            f"Editing Inventory Item "
            f"#{supply_id}: {supply[0]}"
        )
    )


def update_loaded_supply():

    global loaded_supply_id


    if loaded_supply_id is None:

        messagebox.showwarning(
            "No Loaded Item",
            (
                "Load an inventory item from "
                "the table first."
            )
        )

        return


    item_name = (
        inventory_item_entry
        .get()
        .strip()
    )


    category = (
        inventory_category_combo
        .get()
        .strip()
    )


    supplier = (
        inventory_supplier_entry
        .get()
        .strip()
    )


    notes = (
        inventory_notes_entry
        .get()
        .strip()
    )


    if not item_name:

        messagebox.showwarning(
            "Missing Item Name",
            "Please enter a supply item name."
        )

        return


    try:

        quantity_on_hand = float(
            inventory_quantity_entry.get()
        )

        minimum_stock = float(
            inventory_minimum_entry.get()
        )

        cost_per_unit = float(
            inventory_cost_entry.get()
        )


        if (
            quantity_on_hand < 0
            or minimum_stock < 0
            or cost_per_unit < 0
        ):

            raise ValueError


    except ValueError:

        messagebox.showerror(
            "Inventory Error",
            (
                "Quantity, minimum stock, and cost "
                "must be valid non-negative numbers."
            )
        )

        return


    connection = connect_database()

    cursor = connection.cursor()


    try:

        cursor.execute(
            """
            UPDATE supplies
            SET
                item_name = ?,
                category = ?,
                quantity_on_hand = ?,
                minimum_stock = ?,
                cost_per_unit = ?,
                supplier = ?,
                notes = ?
            WHERE id = ?
            """,
            (
                item_name,
                category,
                quantity_on_hand,
                minimum_stock,
                cost_per_unit,
                supplier,
                notes,
                loaded_supply_id
            )
        )


        connection.commit()


    except sqlite3.IntegrityError:

        connection.close()


        messagebox.showwarning(
            "Duplicate Supply",
            (
                "Another inventory item already uses "
                "that name."
            )
        )

        return


    connection.close()


    messagebox.showinfo(
        "Supply Updated",
        f"{item_name} was updated."
    )


    clear_supply_form()

    refresh_supplies_all()


def delete_selected_supply():

    global loaded_supply_id


    selected_items = (
        supplies_table.selection()
    )


    if not selected_items:

        messagebox.showwarning(
            "Select an Item",
            "Please select an inventory item first."
        )

        return


    values = supplies_table.item(
        selected_items[0],
        "values"
    )


    supply_id = int(
        values[0]
    )

    item_name = values[1]


    connection = connect_database()

    cursor = connection.cursor()


    cursor.execute(
        """
        SELECT COUNT(*)
        FROM class_supply_requirements
        WHERE supply_id = ?
        """,
        (
            supply_id,
        )
    )


    requirement_count = int(
        cursor.fetchone()[0]
    )


    connection.close()


    if requirement_count > 0:

        messagebox.showwarning(
            "Supply Is In Use",
            (
                f"{item_name} is used in "
                f"{requirement_count} class supply "
                f"requirement(s).\n\n"
                "Delete those requirement links first "
                "so Lux does not quietly erase class "
                "planning rules."
            )
        )

        return


    confirmed = messagebox.askyesno(
        "Delete Supply",
        (
            f"Delete {item_name} from inventory?\n\n"
            "This removes the inventory record."
        )
    )


    if not confirmed:

        return


    connection = connect_database()

    cursor = connection.cursor()


    cursor.execute(
        """
        DELETE FROM supplies
        WHERE id = ?
        """,
        (
            supply_id,
        )    )


    connection.commit()

    connection.close()


    if loaded_supply_id == supply_id:

        clear_supply_form()


    refresh_supplies_all()


    messagebox.showinfo(
        "Supply Deleted",
        f"{item_name} was deleted."
    )


def refresh_requirement_supply_choices():

    requirement_supply_lookup.clear()


    connection = connect_database()

    cursor = connection.cursor()


    cursor.execute(
        """
        SELECT
            id,
            item_name,
            category
        FROM supplies
        ORDER BY
            category,
            item_name
        """
    )


    supplies = cursor.fetchall()


    connection.close()


    display_values = []


    for supply in supplies:

        display_text = (
            f"#{supply[0]} | "
            f"{supply[1]} | "
            f"{supply[2]}"
        )


        requirement_supply_lookup[
            display_text
        ] = supply[0]


        display_values.append(
            display_text
        )


    requirement_supply_combo[
        "values"
    ] = display_values


    current_value = (
        requirement_supply_combo.get()
    )


    if current_value not in display_values:

        if display_values:

            requirement_supply_combo.set(
                display_values[0]
            )

        else:

            requirement_supply_combo.set(
                ""
            )


def find_supply_display_text(
    supply_id
):

    for display_text, mapped_id in (
        requirement_supply_lookup.items()
    ):

        if int(mapped_id) == int(supply_id):

            return display_text


    return ""


def clear_requirement_form():

    global loaded_requirement_id


    loaded_requirement_id = None


    requirement_class_combo.set(
        "Intro Class"
    )


    if requirement_supply_combo[
        "values"
    ]:

        requirement_supply_combo.set(
            requirement_supply_combo[
                "values"
            ][0]
        )

    else:

        requirement_supply_combo.set(
            ""
        )


    requirement_per_student_entry.delete(
        0,
        tk.END
    )

    requirement_per_student_entry.insert(
        0,
        "0"
    )


    requirement_fixed_entry.delete(
        0,
        tk.END
    )

    requirement_fixed_entry.insert(
        0,
        "0"
    )


    requirement_notes_entry.delete(
        0,
        tk.END
    )


    requirement_edit_status_label.config(
        text="Adding a class supply requirement"
    )


def add_supply_requirement():

    class_type = (
        requirement_class_combo.get()
        .strip()
    )


    supply_display = (
        requirement_supply_combo.get()
    )


    supply_id = (
        requirement_supply_lookup.get(
            supply_display
        )
    )


    if supply_id is None:

        messagebox.showwarning(
            "No Supply Selected",
            (
                "Add at least one inventory item, "
                "then choose it here."
            )
        )

        return


    try:

        quantity_per_student = float(
            requirement_per_student_entry.get()
        )

        fixed_quantity = float(
            requirement_fixed_entry.get()
        )


        if (
            quantity_per_student < 0
            or fixed_quantity < 0
        ):

            raise ValueError


        if (
            quantity_per_student == 0
            and fixed_quantity == 0
        ):

            messagebox.showwarning(
                "Nothing Required",
                (
                    "Enter a per-student amount, "
                    "a fixed per-class amount, or both."
                )
            )

            return


    except ValueError:

        messagebox.showerror(
            "Requirement Error",
            (
                "Requirement quantities must be "
                "valid non-negative numbers."
            )
        )

        return


    notes = (
        requirement_notes_entry
        .get()
        .strip()
    )


    connection = connect_database()

    cursor = connection.cursor()


    try:

        cursor.execute(
            """
            INSERT INTO class_supply_requirements (
                class_type,
                supply_id,
                quantity_per_student,
                fixed_quantity,
                notes,
                created_at
            )
            VALUES (?, ?, ?, ?, ?, ?)
            """,
            (
                class_type,
                supply_id,
                quantity_per_student,
                fixed_quantity,
                notes,
                datetime.now().strftime(
                    "%Y-%m-%d %H:%M"
                )
            )
        )


        connection.commit()


    except sqlite3.IntegrityError:

        connection.close()


        messagebox.showwarning(
            "Requirement Already Exists",
            (
                "That supply is already linked to "
                f"{class_type}.\n\n"
                "Load the existing requirement if "
                "you want to change its quantities."
            )
        )

        return


    connection.close()


    clear_requirement_form()

    refresh_requirements_table()

    refresh_supply_readiness()


    messagebox.showinfo(
        "Requirement Added",
        (
            "Class supply requirement saved."
        )
    )


def refresh_requirements_table():

    if "requirements_table" not in globals():

        return


    for item in requirements_table.get_children():

        requirements_table.delete(
            item
        )


    connection = connect_database()

    cursor = connection.cursor()


    cursor.execute(
        """
        SELECT
            class_supply_requirements.id,
            class_supply_requirements.class_type,
            supplies.item_name,
            class_supply_requirements.quantity_per_student,
            class_supply_requirements.fixed_quantity,
            class_supply_requirements.notes
        FROM class_supply_requirements
        JOIN supplies
            ON supplies.id =
               class_supply_requirements.supply_id
        ORDER BY
            class_supply_requirements.class_type,
            supplies.item_name
        """
    )


    requirements = cursor.fetchall()


    connection.close()


    for requirement in requirements:

        requirements_table.insert(
            "",
            tk.END,
            values=(
                requirement[0],
                requirement[1],
                requirement[2],
                format_supply_quantity(
                    requirement[3]
                ),
                format_supply_quantity(
                    requirement[4]
                ),
                requirement[5] or ""
            )
        )


    if len(requirements) == 1:

        requirement_count_label.config(
            text="1 class supply rule"
        )

    else:

        requirement_count_label.config(
            text=(
                f"{len(requirements)} "
                f"class supply rules"
            )
        )


def load_selected_requirement():

    global loaded_requirement_id


    selected_items = (
        requirements_table.selection()
    )


    if not selected_items:

        messagebox.showwarning(
            "Select a Requirement",
            (
                "Please select a class supply "
                "requirement first."
            )
        )

        return


    values = requirements_table.item(
        selected_items[0],
        "values"
    )


    requirement_id = int(
        values[0]
    )


    connection = connect_database()

    cursor = connection.cursor()


    cursor.execute(
        """
        SELECT
            class_type,
            supply_id,
            quantity_per_student,
            fixed_quantity,
            notes
        FROM class_supply_requirements
        WHERE id = ?
        """,
        (
            requirement_id,
        )
    )


    requirement = cursor.fetchone()


    connection.close()


    if requirement is None:

        messagebox.showerror(
            "Requirement Not Found",
            (
                "That class supply requirement "
                "could not be found."
            )
        )

        refresh_requirements_table()

        return


    loaded_requirement_id = (
        requirement_id
    )


    requirement_class_combo.set(
        requirement[0]
    )


    supply_display = find_supply_display_text(
        requirement[1]
    )


    if supply_display:

        requirement_supply_combo.set(
            supply_display
        )


    requirement_per_student_entry.delete(
        0,
        tk.END
    )

    requirement_per_student_entry.insert(
        0,
        format_supply_quantity(
            requirement[2]
        )
    )


    requirement_fixed_entry.delete(
        0,
        tk.END
    )

    requirement_fixed_entry.insert(
        0,
        format_supply_quantity(
            requirement[3]
        )
    )


    requirement_notes_entry.delete(
        0,
        tk.END
    )

    requirement_notes_entry.insert(
        0,
        requirement[4] or ""
    )


    requirement_edit_status_label.config(
        text=(
            f"Editing Requirement "
            f"#{requirement_id}"
        )
    )


def update_loaded_requirement():

    global loaded_requirement_id


    if loaded_requirement_id is None:

        messagebox.showwarning(
            "No Loaded Requirement",
            (
                "Load a class supply requirement "
                "from the table first."
            )
        )

        return


    class_type = (
        requirement_class_combo.get()
        .strip()
    )


    supply_id = (
        requirement_supply_lookup.get(
            requirement_supply_combo.get()
        )
    )


    if supply_id is None:

        messagebox.showwarning(
            "No Supply Selected",
            "Choose an inventory item first."
        )

        return


    try:

        quantity_per_student = float(
            requirement_per_student_entry.get()
        )

        fixed_quantity = float(
            requirement_fixed_entry.get()
        )


        if (
            quantity_per_student < 0
            or fixed_quantity < 0
        ):

            raise ValueError


        if (
            quantity_per_student == 0
            and fixed_quantity == 0
        ):

            messagebox.showwarning(
                "Nothing Required",
                (
                    "Enter a per-student amount, "
                    "a fixed per-class amount, or both."
                )
            )

            return


    except ValueError:

        messagebox.showerror(
            "Requirement Error",
            (
                "Requirement quantities must be "
                "valid non-negative numbers."
            )
        )

        return


    notes = (
        requirement_notes_entry
        .get()
        .strip()
    )


    connection = connect_database()

    cursor = connection.cursor()


    try:

        cursor.execute(
            """
            UPDATE class_supply_requirements
            SET
                class_type = ?,
                supply_id = ?,
                quantity_per_student = ?,
                fixed_quantity = ?,
                notes = ?
            WHERE id = ?
            """,
            (
                class_type,
                supply_id,
                quantity_per_student,
                fixed_quantity,
                notes,
                loaded_requirement_id
            )
        )


        connection.commit()


    except sqlite3.IntegrityError:

        connection.close()


        messagebox.showwarning(
            "Duplicate Requirement",
            (
                "That supply is already linked to "
                "that class type."
            )
        )

        return


    connection.close()


    messagebox.showinfo(
        "Requirement Updated",
        "Class supply requirement updated."
    )


    clear_requirement_form()

    refresh_requirements_table()

    refresh_supply_readiness()


def delete_selected_requirement():

    global loaded_requirement_id


    selected_items = (
        requirements_table.selection()
    )


    if not selected_items:

        messagebox.showwarning(
            "Select a Requirement",
            (
                "Please select a class supply "
                "requirement first."
            )
        )

        return


    values = requirements_table.item(
        selected_items[0],
        "values"
    )


    requirement_id = int(
        values[0]
    )


    confirmed = messagebox.askyesno(
        "Delete Requirement",
        (
            f"Remove {values[2]} from the "
            f"{values[1]} supply requirements?"
        )
    )


    if not confirmed:

        return


    connection = connect_database()

    cursor = connection.cursor()


    cursor.execute(
        """
        DELETE FROM class_supply_requirements
        WHERE id = ?
        """,
        (
            requirement_id,
        )
    )


    connection.commit()

    connection.close()


    if loaded_requirement_id == requirement_id:

        clear_requirement_form()


    refresh_requirements_table()

    refresh_supply_readiness()


def refresh_readiness_session_choices():

    readiness_session_lookup.clear()


    connection = connect_database()

    cursor = connection.cursor()


    cursor.execute(
        """
        SELECT
            class_sessions.id,
            class_sessions.class_date,
            class_sessions.class_time,
            class_sessions.class_type,
            class_sessions.capacity,
            COUNT(enrollments.id)
        FROM class_sessions
        LEFT JOIN enrollments
            ON enrollments.session_id =
               class_sessions.id
        GROUP BY
            class_sessions.id,
            class_sessions.class_date,
            class_sessions.class_time,
            class_sessions.class_type,
            class_sessions.capacity
        ORDER BY
            class_sessions.class_date ASC,
            class_sessions.class_time ASC,
            class_sessions.id ASC
        """
    )


    sessions = cursor.fetchall()


    connection.close()


    display_values = []


    for session in sessions:

        display_text = (
            f"#{session[0]} | "
            f"{display_class_date(session[1])} "
            f"{display_class_time(session[2])} | "
            f"{session[3]} | "
            f"{session[5]}/{session[4]} enrolled"
        )


        readiness_session_lookup[
            display_text
        ] = session[0]


        display_values.append(
            display_text
        )


    current_value = (
        readiness_session_combo.get()
    )


    readiness_session_combo[
        "values"
    ] = display_values


    if current_value not in display_values:

        if display_values:

            readiness_session_combo.set(
                display_values[0]
            )

        else:

            readiness_session_combo.set(
                ""
            )


def refresh_supply_readiness(
    event=None
):

    if "readiness_table" not in globals():

        return


    for item in readiness_table.get_children():

        readiness_table.delete(
            item
        )


    session_id = (
        readiness_session_lookup.get(
            readiness_session_combo.get()
        )
    )


    if session_id is None:

        readiness_session_summary_label.config(
            text="No scheduled class selected."
        )

        readiness_overall_status_label.config(
            text=""
        )

        return


    connection = connect_database()

    cursor = connection.cursor()


    cursor.execute(
        """
        SELECT
            class_type,
            class_date,
            class_time,
            capacity,
            status
        FROM class_sessions
        WHERE id = ?
        """,
        (
            session_id,
        )
    )


    session = cursor.fetchone()


    if session is None:

        connection.close()

        readiness_session_summary_label.config(
            text="That class session could not be found."
        )

        return


    cursor.execute(
        """
        SELECT COUNT(*)
        FROM enrollments
        WHERE session_id = ?
        """,
        (
            session_id,
        )
    )


    enrolled_count = int(
        cursor.fetchone()[0]
    )


    cursor.execute(
        """
        SELECT
            supplies.item_name,
            supplies.category,
            supplies.quantity_on_hand,
            supplies.minimum_stock,
            supplies.cost_per_unit,
            class_supply_requirements.quantity_per_student,
            class_supply_requirements.fixed_quantity
        FROM class_supply_requirements
        JOIN supplies
            ON supplies.id =
               class_supply_requirements.supply_id
        WHERE class_supply_requirements.class_type = ?
        ORDER BY
            supplies.category,
            supplies.item_name        """,
        (
            session[0],
        )
    )


    requirements = cursor.fetchall()


    connection.close()


    readiness_session_summary_label.config(
        text=(
            f"Session #{session_id} • "
            f"{session[0]} • "
            f"{display_class_date(session[1])} "
            f"at {display_class_time(session[2])} • "
            f"{enrolled_count}/{session[3]} enrolled"
        )
    )


    shortage_items = 0

    low_after_items = 0

    ready_items = 0

    estimated_shortage_cost = 0


    for requirement in requirements:

        item_name = requirement[0]

        category = requirement[1]

        on_hand = float(
            requirement[2]
        )

        minimum_stock = float(
            requirement[3]
        )

        cost_per_unit = float(
            requirement[4]
        )

        per_student = float(
            requirement[5]
        )

        fixed_quantity = float(
            requirement[6]
        )


        required_quantity = (
            per_student * enrolled_count
        ) + fixed_quantity


        shortage = max(
            0,
            required_quantity - on_hand
        )


        remaining_after_class = (
            on_hand - required_quantity
        )


        if shortage > 0:

            status = (
                f"NEED "
                f"{format_supply_quantity(shortage)}"
            )

            shortage_items += 1

            estimated_shortage_cost += (
                shortage * cost_per_unit
            )


        elif remaining_after_class <= minimum_stock:

            status = (
                "READY • LOW AFTER CLASS"
            )

            low_after_items += 1


        else:

            status = "READY"

            ready_items += 1


        readiness_table.insert(
            "",
            tk.END,
            values=(
                item_name,
                category,
                format_supply_quantity(
                    per_student
                ),
                format_supply_quantity(
                    fixed_quantity
                ),
                enrolled_count,
                format_supply_quantity(
                    required_quantity
                ),
                format_supply_quantity(
                    on_hand
                ),
                format_supply_quantity(
                    shortage
                ),
                status
            )
        )


    if not requirements:

        readiness_overall_status_label.config(
            text=(
                f"No supply requirements are configured "
                f"for {session[0]} yet."
            )
        )


    elif shortage_items > 0:

        readiness_overall_status_label.config(
            text=(
                f"⚠ {shortage_items} item(s) need ordering • "
                f"Estimated shortage cost: "
                f"${estimated_shortage_cost:,.2f}"
            )
        )


    else:

        readiness_overall_status_label.config(
            text=(
                f"✓ Supplies cover this class • "
                f"{low_after_items} item(s) will be "
                f"at or below minimum stock afterward"
            )
        )


def refresh_supplies_all():

    refresh_supplies_table()

    refresh_requirement_supply_choices()

    refresh_requirements_table()

    refresh_readiness_session_choices()

    refresh_supply_readiness()



# =====================================================
# v4.2 REPORTS + CSV EXPORTS
# =====================================================

# Reports do NOT create another source-of-truth table.
#
# They read the same live database used by the Dashboard,
# Classes, Students, Enrollments, Payments, and Readiness.
#
# Store the facts once. Format them many ways.

report_session_lookup = {}
current_report_export = None


def refresh_report_session_choices():

    # Build a human-readable list of scheduled classes.
    # The dictionary remembers which permanent session ID
    # belongs to each visible description.

    report_session_lookup.clear()

    connection = connect_database()
    cursor = connection.cursor()

    cursor.execute(
        """
        SELECT
            class_sessions.id,
            class_sessions.class_date,
            class_sessions.class_time,
            class_sessions.class_type,
            class_sessions.capacity,
            COUNT(enrollments.id)
        FROM class_sessions
        LEFT JOIN enrollments
            ON enrollments.session_id = class_sessions.id
        GROUP BY
            class_sessions.id,
            class_sessions.class_date,
            class_sessions.class_time,
            class_sessions.class_type,
            class_sessions.capacity
        ORDER BY
            class_sessions.class_date ASC,
            class_sessions.class_time ASC,
            class_sessions.id ASC
        """
    )

    sessions = cursor.fetchall()
    connection.close()

    values = ["All Classes"]
    report_session_lookup["All Classes"] = None

    for session in sessions:

        display_text = (
            f"#{session[0]} | "
            f"{display_class_date(session[1])} "
            f"{display_class_time(session[2])} | "
            f"{session[3]} | "
            f"{session[5]}/{session[4]} enrolled"
        )

        values.append(display_text)
        report_session_lookup[display_text] = session[0]

    if "report_session_combo" not in globals():
        return

    current_value = report_session_combo.get()
    report_session_combo["values"] = values

    if current_value not in values:
        report_session_combo.set("All Classes")


def set_report_session_to_first_real_class():

    # Class Roster and Readiness reports need ONE real class,
    # so if "All Classes" is selected, gently choose the first
    # actual session instead of making the user wrestle the UI.

    values = list(report_session_combo["values"])

    for value in values:
        if value != "All Classes":
            report_session_combo.set(value)
            return True

    return False


def report_type_changed(event=None):

    if "report_type_combo" not in globals():
        return

    report_type = report_type_combo.get()

    refresh_report_session_choices()

    if report_type in (
        "Class Roster",
        "Class Readiness"
    ):

        report_session_combo.configure(state="readonly")
        report_balance_filter_combo.configure(state="disabled")

        if report_session_combo.get() == "All Classes":
            set_report_session_to_first_real_class()

    elif report_type == "Payment & Balance":

        report_session_combo.configure(state="readonly")
        report_balance_filter_combo.configure(state="readonly")

    else:

        # Student Directory is already a business-wide report.
        report_session_combo.set("All Classes")
        report_session_combo.configure(state="disabled")
        report_balance_filter_combo.configure(state="disabled")

    preview_current_report(silent=True)


def configure_report_preview(columns, headings, widths=None, anchors=None):

    # A Treeview can change its columns at runtime.
    # That lets one preview area display several report types.

    for item in report_preview_table.get_children():
        report_preview_table.delete(item)

    report_preview_table["columns"] = columns

    widths = widths or {}
    anchors = anchors or {}

    for column in columns:
        report_preview_table.heading(
            column,
            text=headings.get(column, column)
        )

        report_preview_table.column(
            column,
            width=widths.get(column, 120),
            anchor=anchors.get(column, "w"),
            stretch=True
        )


def payment_status_from_amounts(tuition, paid):

    balance = max(0.0, float(tuition) - float(paid))

    if float(paid) <= 0.005:
        return "Not Paid"

    if balance <= 0.005:
        return "Paid in Full"

    return "Deposit Paid"


def build_class_roster_report(session_id):

    connection = connect_database()
    cursor = connection.cursor()

    cursor.execute(
        """
        SELECT
            id,
            class_type,
            class_date,
            class_time,
            capacity,
            tuition,
            status
        FROM class_sessions
        WHERE id = ?
        """,
        (session_id,)
    )

    session = cursor.fetchone()

    if session is None:
        connection.close()
        return None

    cursor.execute(
        """
        SELECT
            enrollments.id,
            students.id,
            students.first_name,
            students.last_name,
            students.phone,
            students.email,
            enrollments.certification_status,
            COALESCE(SUM(payments.amount), 0)
        FROM enrollments
        JOIN students
            ON students.id = enrollments.student_id
        LEFT JOIN payments
            ON payments.enrollment_id = enrollments.id
        WHERE enrollments.session_id = ?
        GROUP BY
            enrollments.id,
            students.id,
            students.first_name,
            students.last_name,
            students.phone,
            students.email,
            enrollments.certification_status
        ORDER BY
            students.last_name COLLATE NOCASE,
            students.first_name COLLATE NOCASE
        """,
        (session_id,)
    )

    records = cursor.fetchall()
    connection.close()

    tuition = float(session[5])
    expected = tuition * len(records)
    collected = sum(float(row[7]) for row in records)
    outstanding = max(0.0, expected - collected)

    preview_rows = []
    export_rows = []

    for row in records:

        paid = float(row[7])
        balance = max(0.0, tuition - paid)
        payment_status = payment_status_from_amounts(tuition, paid)

        student_name = f"{row[2]} {row[3]}"

        preview_rows.append((
            row[1],
            student_name,
            payment_status,
            f"${paid:,.2f}",
            f"${balance:,.2f}",
            row[6],
            row[4] or "",
            row[5] or ""
        ))

        export_rows.append([
            row[0],
            row[1],
            student_name,
            row[4] or "",
            row[5] or "",
            f"{tuition:.2f}",
            f"{paid:.2f}",
            f"{balance:.2f}",
            payment_status,
            row[6]
        ])

    date_text = display_class_date(session[2])
    time_text = display_class_time(session[3])

    csv_rows = [
        ["LUX BEAUTY LASH ACADEMY"],
        ["CLASS ROSTER REPORT"],
        [],
        ["Session ID", session[0]],
        ["Class", session[1]],
        ["Date", date_text],
        ["Time", time_text],
        ["Status", session[6]],
        ["Capacity", session[4]],
        ["Enrolled", len(records)],
        ["Expected Tuition", f"{expected:.2f}"],
        ["Collected", f"{collected:.2f}"],
        ["Outstanding", f"{outstanding:.2f}"],
        [],
        [
            "Enrollment ID",
            "Student ID",
            "Student",
            "Phone",
            "Email",
            "Tuition",
            "Paid",
            "Balance",
            "Payment Status",
            "Certification"
        ],
        *export_rows
    ]

    return {
        "title": f"Class Roster • {session[1]} • {date_text} at {time_text}",
        "summary": (
            f"{len(records)}/{session[4]} enrolled • "
            f"Expected ${expected:,.2f} • "
            f"Collected ${collected:,.2f} • "
            f"Outstanding ${outstanding:,.2f}"
        ),
        "columns": (
            "student_id",
            "student",
            "payment",
            "paid",
            "balance",
            "certification",
            "phone",
            "email"
        ),
        "headings": {
            "student_id": "Student ID",
            "student": "Student",
            "payment": "Payment Status",
            "paid": "Paid",
            "balance": "Balance",
            "certification": "Certification",
            "phone": "Phone",
            "email": "Email"
        },
        "widths": {
            "student_id": 70,
            "student": 150,
            "payment": 115,
            "paid": 85,
            "balance": 85,
            "certification": 110,
            "phone": 115,
            "email": 180
        },
        "anchors": {
            "student_id": "center",
            "paid": "e",
            "balance": "e"
        },
        "preview_rows": preview_rows,
        "csv_rows": csv_rows,
        "filename": (
            f"Lux_Class_Roster_{session[1].replace(' ', '_')}_"
            f"{session[2]}.csv"
        )
    }


def build_payment_balance_report(session_id=None, balance_filter="All Balances"):

    connection = connect_database()
    cursor = connection.cursor()

    sql = """
        SELECT
            enrollments.id,
            students.first_name,
            students.last_name,
            class_sessions.id,
            class_sessions.class_type,
            class_sessions.class_date,
            class_sessions.class_time,
            class_sessions.tuition,
            enrollments.certification_status,
            COALESCE(SUM(payments.amount), 0)
        FROM enrollments
        JOIN students
            ON students.id = enrollments.student_id
        JOIN class_sessions
            ON class_sessions.id = enrollments.session_id
        LEFT JOIN payments
            ON payments.enrollment_id = enrollments.id
    """

    params = []

    if session_id is not None:
        sql += " WHERE class_sessions.id = ? "
        params.append(session_id)

    sql += """
        GROUP BY
            enrollments.id,
            students.first_name,
            students.last_name,
            class_sessions.id,
            class_sessions.class_type,
            class_sessions.class_date,
            class_sessions.class_time,
            class_sessions.tuition,
            enrollments.certification_status
        ORDER BY
            class_sessions.class_date ASC,
            class_sessions.class_time ASC,
            students.last_name COLLATE NOCASE,
            students.first_name COLLATE NOCASE
    """

    cursor.execute(sql, params)
    records = cursor.fetchall()
    connection.close()

    preview_rows = []
    export_rows = []

    expected_total = 0.0
    collected_total = 0.0
    outstanding_total = 0.0
    kept_count = 0

    for row in records:

        tuition = float(row[7])
        paid = float(row[9])
        balance = max(0.0, tuition - paid)
        payment_status = payment_status_from_amounts(tuition, paid)

        if balance_filter == "Balance Due" and balance <= 0.005:
            continue

        if balance_filter == "Paid in Full" and balance > 0.005:
            continue

        kept_count += 1
        expected_total += tuition
        collected_total += paid
        outstanding_total += balance

        student_name = f"{row[1]} {row[2]}"
        class_text = f"#{row[3]} {row[4]}"
        date_text = display_class_date(row[5])
        time_text = display_class_time(row[6])

        preview_rows.append((
            row[0],
            student_name,
            class_text,
            date_text,
            f"${tuition:,.2f}",
            f"${paid:,.2f}",
            f"${balance:,.2f}",
            payment_status
        ))

        export_rows.append([
            row[0],
            student_name,
            row[3],
            row[4],
            date_text,
            time_text,
            f"{tuition:.2f}",
            f"{paid:.2f}",
            f"{balance:.2f}",
            payment_status,
            row[8]
        ])

    scope_text = "All Classes"

    if session_id is not None:
        scope_text = report_session_combo.get()

    csv_rows = [
        ["LUX BEAUTY LASH ACADEMY"],
        ["PAYMENT & BALANCE REPORT"],
        [],
        ["Scope", scope_text],
        ["Balance Filter", balance_filter],
        ["Enrollments", kept_count],
        ["Expected Tuition", f"{expected_total:.2f}"],
        ["Collected", f"{collected_total:.2f}"],
        ["Outstanding", f"{outstanding_total:.2f}"],
        [],
        [
            "Enrollment ID",
            "Student",
            "Session ID",
            "Class",
            "Date",
            "Time",
            "Tuition",
            "Paid",
            "Balance",
            "Payment Status",
            "Certification"
        ],
        *export_rows
    ]

    return {
        "title": "Payment & Balance Report",
        "summary": (
            f"{kept_count} enrollment(s) • "
            f"Expected ${expected_total:,.2f} • "
            f"Collected ${collected_total:,.2f} • "
            f"Outstanding ${outstanding_total:,.2f}"
        ),
        "columns": (
            "enrollment_id",
            "student",
            "class",
            "date",
            "tuition",
            "paid",
            "balance",
            "payment"
        ),
        "headings": {
            "enrollment_id": "Enrollment ID",
            "student": "Student",
            "class": "Class",
            "date": "Date",
            "tuition": "Tuition",
            "paid": "Paid",
            "balance": "Balance",
            "payment": "Payment Status"
        },
        "widths": {
            "enrollment_id": 85,
            "student": 150,
            "class": 150,
            "date": 95,
            "tuition": 85,
            "paid": 85,
            "balance": 85,
            "payment": 115
        },
        "anchors": {
            "enrollment_id": "center",
            "date": "center",
            "tuition": "e",
            "paid": "e",
            "balance": "e"
        },
        "preview_rows": preview_rows,
        "csv_rows": csv_rows,
        "filename": (
            "Lux_Payment_Balance_Report_"
            + datetime.now().strftime("%Y-%m-%d")
            + ".csv"
        )
    }


def build_student_directory_report():

    connection = connect_database()
    cursor = connection.cursor()

    cursor.execute(
        """
        SELECT
            students.id,
            students.first_name,
            students.last_name,
            students.phone,
            students.email,
            COUNT(enrollments.id),
            COALESCE(
                SUM(
                    CASE
                        WHEN class_sessions.tuition IS NULL
                            THEN 0
                        WHEN (
                            class_sessions.tuition -
                            COALESCE(payment_totals.total_paid, 0)
                        ) > 0
                            THEN (
                                class_sessions.tuition -
                                COALESCE(payment_totals.total_paid, 0)
                            )
                        ELSE 0
                    END
                ),
                0
            )
        FROM students
        LEFT JOIN enrollments
            ON enrollments.student_id = students.id
        LEFT JOIN class_sessions
            ON class_sessions.id = enrollments.session_id
        LEFT JOIN (
            SELECT
                enrollment_id,
                SUM(amount) AS total_paid
            FROM payments
            GROUP BY enrollment_id
        ) AS payment_totals
            ON payment_totals.enrollment_id = enrollments.id
        GROUP BY
            students.id,
            students.first_name,
            students.last_name,
            students.phone,
            students.email
        ORDER BY
            students.last_name COLLATE NOCASE,
            students.first_name COLLATE NOCASE
        """
    )

    records = cursor.fetchall()
    connection.close()

    preview_rows = []
    export_rows = []
    total_outstanding = 0.0

    for row in records:

        outstanding = float(row[6])
        total_outstanding += outstanding
        student_name = f"{row[1]} {row[2]}"

        preview_rows.append((
            row[0],
            student_name,
            row[3] or "",
            row[4] or "",
            row[5],
            f"${outstanding:,.2f}"
        ))

        export_rows.append([
            row[0],
            student_name,
            row[3] or "",
            row[4] or "",
            row[5],
            f"{outstanding:.2f}"
        ])

    csv_rows = [
        ["LUX BEAUTY LASH ACADEMY"],
        ["STUDENT DIRECTORY REPORT"],
        [],
        ["Students", len(records)],
        ["Combined Outstanding Balance", f"{total_outstanding:.2f}"],
        [],
        [
            "Student ID",
            "Student",
            "Phone",
            "Email",
            "Enrollment Count",
            "Outstanding Balance"
        ],
        *export_rows
    ]

    return {
        "title": "Student Directory Report",
        "summary": (
            f"{len(records)} student(s) • "
            f"Combined outstanding balance ${total_outstanding:,.2f}"
        ),
        "columns": (
            "student_id",
            "student",
            "phone",
            "email",
            "enrollments",
            "outstanding"
        ),
        "headings": {
            "student_id": "Student ID",
            "student": "Student",
            "phone": "Phone",
            "email": "Email",
            "enrollments": "Enrollments",
            "outstanding": "Outstanding"
        },
        "widths": {
            "student_id": 70,
            "student": 170,
            "phone": 120,
            "email": 220,
            "enrollments": 90,
            "outstanding": 110
        },
        "anchors": {
            "student_id": "center",
            "enrollments": "center",
            "outstanding": "e"
        },
        "preview_rows": preview_rows,
        "csv_rows": csv_rows,
        "filename": (
            "Lux_Student_Directory_"
            + datetime.now().strftime("%Y-%m-%d")
            + ".csv"
        )
    }


def build_class_readiness_report(session_id):

    snapshot = get_class_readiness_snapshot(session_id)

    if snapshot is None:
        return None

    date_text = display_class_date(snapshot["class_date"])
    time_text = display_class_time(snapshot["class_time"])

    preview_rows = list(snapshot["checklist"])

    csv_rows = [
        ["LUX BEAUTY LASH ACADEMY"],
        ["CLASS READINESS REPORT"],
        [],
        ["Session ID", snapshot["session_id"]],
        ["Class", snapshot["class_type"]],
        ["Date", date_text],
        ["Time", time_text],
        ["Overall Status", snapshot["overall_status"]],
        ["Overall Detail", snapshot["overall_detail"]],
        ["Enrollment", f"{snapshot['enrolled']}/{snapshot['capacity']}"],
        ["Expected Tuition", f"{snapshot['expected']:.2f}"],
        ["Collected", f"{snapshot['paid']:.2f}"],
        ["Outstanding", f"{snapshot['outstanding']:.2f}"],
        ["Supply Shortages", snapshot["shortage_items"]],
        ["Estimated Shortage Cost", f"{snapshot['shortage_cost']:.2f}"],
        [],
        ["MASTER CHECKLIST"],
        ["Area", "Status", "What Lux Found"]
    ]

    for row in snapshot["checklist"]:
        csv_rows.append(list(row))

    csv_rows.extend([
        [],
        ["STUDENT READINESS"],
        [
            "Student ID",
            "Student",
            "Tuition",
            "Paid",
            "Balance",
            "Payment Status",
            "Certification",
            "Status"
        ]
    ])

    for student in snapshot["students"]:

        student_status = (
            "CLEAR"
            if student["balance"] <= 0.005
            else "BALANCE DUE"
        )

        csv_rows.append([
            student["student_id"],
            student["name"],
            f"{snapshot['tuition']:.2f}",
            f"{student['paid']:.2f}",
            f"{student['balance']:.2f}",
            student["payment_status"],
            student["certification"],
            student_status
        ])

    csv_rows.extend([
        [],
        ["SUPPLY READINESS"],
        [
            "Supply",
            "Category",
            "Need",
            "Have",
            "Short",
            "Status"
        ]
    ])

    for supply in snapshot["supplies"]:
        csv_rows.append([
            supply["name"],
            supply["category"],
            format_supply_quantity(supply["need"]),
            format_supply_quantity(supply["have"]),
            format_supply_quantity(supply["short"]),
            supply["status"]
        ])

    return {
        "title": (
            f"Class Readiness • {snapshot['class_type']} • "
            f"{date_text} at {time_text}"
        ),
        "summary": (
            f"{snapshot['overall_status']} • "
            f"{snapshot['overall_detail']}"
        ),
        "columns": (
            "area",
            "status",
            "detail"
        ),
        "headings": {
            "area": "Area",
            "status": "Status",
            "detail": "What Lux Found"
        },
        "widths": {
            "area": 120,
            "status": 150,            "detail": 520
        },
        "anchors": {
            "area": "w",
            "status": "center",
            "detail": "w"
        },
        "preview_rows": preview_rows,
        "csv_rows": csv_rows,
        "filename": (
            f"Lux_Readiness_{snapshot['class_type'].replace(' ', '_')}_"
            f"{snapshot['class_date']}.csv"
        )
    }


def preview_current_report(event=None, silent=False):

    global current_report_export

    if "report_type_combo" not in globals():
        return

    report_type = report_type_combo.get()
    balance_filter = report_balance_filter_combo.get()
    session_id = report_session_lookup.get(report_session_combo.get())

    report = None

    if report_type == "Class Roster":

        if session_id is None:
            if not set_report_session_to_first_real_class():
                if not silent:
                    messagebox.showwarning(
                        "No Classes",
                        "Create a scheduled class before building a roster report."
                    )
                return

            session_id = report_session_lookup.get(report_session_combo.get())

        report = build_class_roster_report(session_id)

    elif report_type == "Payment & Balance":

        report = build_payment_balance_report(
            session_id=session_id,
            balance_filter=balance_filter
        )

    elif report_type == "Student Directory":

        report = build_student_directory_report()

    elif report_type == "Class Readiness":

        if session_id is None:
            if not set_report_session_to_first_real_class():
                if not silent:
                    messagebox.showwarning(
                        "No Classes",
                        "Create a scheduled class before building a readiness report."
                    )
                return

            session_id = report_session_lookup.get(report_session_combo.get())

        report = build_class_readiness_report(session_id)

    if report is None:

        current_report_export = None

        if not silent:
            messagebox.showerror(
                "Report Error",
                "Lux could not build that report from the current records."
            )

        return

    current_report_export = report

    report_preview_title_label.config(text=report["title"])
    report_preview_summary_label.config(text=report["summary"])

    configure_report_preview(
        report["columns"],
        report["headings"],
        report.get("widths"),
        report.get("anchors")
    )

    for row in report["preview_rows"]:
        report_preview_table.insert(
            "",
            tk.END,
            values=row
        )

    report_preview_count_label.config(
        text=f"{len(report['preview_rows'])} preview row(s)"
    )


def export_current_report_csv():

    global current_report_export

    # Always rebuild immediately before export so the CSV
    # reflects the live database, not an older preview.

    preview_current_report(silent=False)

    if current_report_export is None:
        return

    save_path = filedialog.asksaveasfilename(
        title="Save Lux Report",
        defaultextension=".csv",
        initialfile=current_report_export["filename"],
        filetypes=(
            ("CSV files", "*.csv"),
            ("All files", "*.*")
        )
    )

    if not save_path:
        return

    try:

        # utf-8-sig includes a tiny Unicode marker that helps
        # Excel recognize UTF-8 correctly when names contain
        # accents or other non-ASCII characters.

        with open(
            save_path,
            "w",
            newline="",
            encoding="utf-8-sig"
        ) as file:

            writer = csv.writer(file)
            writer.writerows(current_report_export["csv_rows"])

        messagebox.showinfo(
            "Report Exported",
            f"Report saved to:\n\n{save_path}"
        )

    except OSError as error:

        messagebox.showerror(
            "Export Error",
            f"Lux could not save the report.\n\n{error}"
        )


def refresh_reports_all():

    if "report_type_combo" not in globals():
        return

    refresh_report_session_choices()
    report_type_changed()



# =====================================================
# BACKUP + RESTORE
# =====================================================

# v4.3 gives Lux a proper data-safety system.
#
# A backup ZIP can contain:
#
#   lux_command_center.db
#       The SQLite database containing students,
#       classes, enrollments, payments, supplies, etc.
#
#   lux_class_history.csv
#       The older saved calculator-plan history,
#       when that file exists.
#
#   backup_info.txt
#       A small human-readable note describing when
#       and why the backup was created.
#
#   code_snapshot/lux_command_center.py
#       A copy of the Python program that created
#       the backup. Restore does NOT overwrite the
#       running program with this file. It is only
#       there as an emergency reference copy.
#
# IMPORTANT:
# SQLite gets copied using SQLite's own backup()
# method instead of simply copying an open database.
# That gives us a consistent snapshot.


DATABASE_FILE = "lux_command_center.db"
HISTORY_FILE = "lux_class_history.csv"


def get_program_folder():

    # __file__ is the path of the running Python file.
    #
    # abspath() turns it into a complete path.
    # dirname() removes the filename and leaves
    # only the folder.

    try:

        return os.path.dirname(
            os.path.abspath(__file__)
        )

    except NameError:

        # This fallback is mostly useful while
        # testing code in unusual environments.

        return os.getcwd()


def get_local_backup_folder():

    # Keep automatic backups in their own folder
    # beside the Lux application.

    backup_folder = os.path.join(
        get_program_folder(),
        "LuxBackups"
    )


    os.makedirs(
        backup_folder,
        exist_ok=True
    )


    return backup_folder


def update_backup_status(message):

    # The function can be called by several backup
    # operations. The try keeps it harmless if it is
    # ever called before the GUI label exists.

    try:

        backup_status_label.config(
            text=message
        )

    except NameError:

        pass


def refresh_backup_file_status():

    # Show the user which live data files currently
    # exist and roughly how large the database is.

    database_path = os.path.join(
        get_program_folder(),
        DATABASE_FILE
    )


    history_path = os.path.join(
        get_program_folder(),
        HISTORY_FILE
    )


    if os.path.exists(database_path):

        database_size = (
            os.path.getsize(database_path)
            / 1024
        )


        database_status_label.config(
            text=(
                f"SQLite database: READY "
                f"({database_size:,.1f} KB)"
            )
        )

    else:

        database_status_label.config(
            text="SQLite database: NOT FOUND"
        )


    if os.path.exists(history_path):

        history_backup_status_label.config(
            text="CSV class-plan history: FOUND"
        )

    else:

        history_backup_status_label.config(
            text=(
                "CSV class-plan history: "
                "No file yet"
            )
        )


def create_backup_archive(
    destination_path,
    reason="Manual backup",
    announce=True
):

    # This is the core backup engine.
    #
    # Other buttons can call this function and choose
    # WHERE the resulting ZIP should be saved.

    temp_folder = tempfile.mkdtemp(
        prefix="lux_backup_"
    )


    try:

        live_database_path = os.path.join(
            get_program_folder(),
            DATABASE_FILE
        )


        temp_database_path = os.path.join(
            temp_folder,
            DATABASE_FILE
        )


        # ---------------------------------------------
        # CREATE A CONSISTENT SQLITE SNAPSHOT
        # ---------------------------------------------

        source_connection = sqlite3.connect(
            live_database_path
        )


        backup_connection = sqlite3.connect(
            temp_database_path
        )


        try:

            # SQLite's backup() copies a safe snapshot
            # of the database into the temporary file.

            source_connection.backup(
                backup_connection
            )

        finally:

            backup_connection.close()

            source_connection.close()


        # ---------------------------------------------
        # BACKUP INFORMATION FILE
        # ---------------------------------------------

        backup_time = datetime.now().strftime(
            "%Y-%m-%d %H:%M:%S"
        )


        info_path = os.path.join(
            temp_folder,
            "backup_info.txt"
        )


        history_path = os.path.join(
            get_program_folder(),
            HISTORY_FILE
        )


        history_included = os.path.exists(
            history_path
        )


        with open(
            info_path,
            "w",
            encoding="utf-8"
        ) as info_file:

            info_file.write(
                "Lux Class Command Center Backup\n"
            )

            info_file.write(
                "Application Version: 4.3\n"
            )

            info_file.write(
                f"Created: {backup_time}\n"
            )

            info_file.write(
                f"Reason: {reason}\n"
            )

            info_file.write(
                (
                    "CSV History Included: "
                    f"{'Yes' if history_included else 'No'}\n"
                )
            )


        # ---------------------------------------------
        # BUILD THE ZIP ARCHIVE
        # ---------------------------------------------

        with zipfile.ZipFile(
            destination_path,
            "w",
            zipfile.ZIP_DEFLATED
        ) as backup_zip:

            backup_zip.write(
                temp_database_path,
                arcname=DATABASE_FILE
            )


            backup_zip.write(
                info_path,
                arcname="backup_info.txt"
            )


            if history_included:

                backup_zip.write(
                    history_path,
                    arcname=HISTORY_FILE
                )


            # Keep a copy of the program that CREATED
            # the backup. This is a reference snapshot
            # only. Restore never replaces the running
            # application code with it.

            try:

                current_program = os.path.abspath(
                    __file__
                )


                if os.path.exists(
                    current_program
                ):

                    backup_zip.write(
                        current_program,
                        arcname=(
                            "code_snapshot/"
                            "lux_command_center.py"
                        )
                    )

            except NameError:

                pass


        update_backup_status(
            (
                "Last backup created: "
                f"{os.path.basename(destination_path)}"
            )
        )


        refresh_backup_file_status()


        if announce:

            messagebox.showinfo(
                "Backup Created",
                (
                    "Lux data was backed up successfully.\n\n"
                    f"{destination_path}"
                )
            )


        return destination_path


    except Exception as error:

        if announce:

            messagebox.showerror(
                "Backup Failed",
                (
                    "Lux could not create the backup.\n\n"
                    f"{error}"
                )
            )


        return None


    finally:

        shutil.rmtree(
            temp_folder,
            ignore_errors=True
        )


def save_backup_as():

    # Let the user choose where the backup goes.
    #
    # This is especially useful for:
    #
    # OneDrive
    # Google Drive folders
    # USB drives
    # another computer location

    timestamp = datetime.now().strftime(
        "%Y-%m-%d_%H%M%S"
    )


    destination_path = filedialog.asksaveasfilename(
        title="Save Lux Backup",
        defaultextension=".zip",
        filetypes=[
            (
                "Lux Backup ZIP",
                "*.zip"
            )
        ],
        initialfile=(
            f"Lux_Backup_{timestamp}.zip"
        )
    )


    if not destination_path:

        return


    create_backup_archive(
        destination_path,
        reason="Manual Save Backup As",
        announce=True
    )


def quick_backup():

    # One-click local backup.
    #
    # These go into:
    #
    # LuxClassCalculator\LuxBackups
    #
    # This is convenient for frequent backups.
    # A second copy somewhere outside the computer
    # is still wise for important real business data.

    timestamp = datetime.now().strftime(
        "%Y-%m-%d_%H%M%S"
    )


    backup_folder = get_local_backup_folder()


    destination_path = os.path.join(
        backup_folder,
        f"Lux_Backup_{timestamp}.zip"
    )


    create_backup_archive(
        destination_path,
        reason="Quick local backup",
        announce=True
    )


def create_pre_restore_safety_backup():

    # Restoring replaces current data.
    #
    # Before doing that, Lux automatically saves
    # the CURRENT state so a mistaken restore can
    # be undone.

    timestamp = datetime.now().strftime(
        "%Y-%m-%d_%H%M%S"
    )


    backup_folder = get_local_backup_folder()


    destination_path = os.path.join(
        backup_folder,
        (
            "Before_Restore_"
            f"{timestamp}.zip"
        )
    )


    return create_backup_archive(
        destination_path,
        reason=(
            "Automatic safety backup "
            "before restore"
        ),
        announce=False
    )


def validate_backup_database(database_path):

    # Before replacing anything, make sure the
    # database inside the ZIP:
    #
    # 1. Opens successfully
    # 2. Passes SQLite's integrity check
    # 3. Looks like a Lux database

    connection = sqlite3.connect(
        database_path
    )


    try:

        integrity_result = connection.execute(
            "PRAGMA integrity_check"
        ).fetchone()


        if (
            integrity_result is None
            or integrity_result[0].lower() != "ok"
        ):

            raise ValueError(
                "SQLite integrity check failed."
            )


        table_rows = connection.execute(
            """
            SELECT name
            FROM sqlite_master
            WHERE type = 'table'
            """
        ).fetchall()


        table_names = {
            row[0]
            for row in table_rows
        }


        # These two tables have existed since our
        # early relational versions and are enough
        # to identify this as a Lux database.
        #
        # initialize_database() will safely add any
        # newer tables if an older valid backup is
        # ever restored.

        required_tables = {
            "students",
            "class_sessions"
        }


        if not required_tables.issubset(
            table_names
        ):

            raise ValueError(
                (
                    "This ZIP does not contain a "
                    "recognized Lux database."
                )
            )


        return True


    finally:

        connection.close()


def refresh_everything_after_restore():

    # A restore changes data underneath the GUI.
    #
    # Refresh every page so the screen immediately
    # matches the newly restored database.

    refresh_class_sessions_table()

    refresh_history_table()

    refresh_students_table()

    refresh_enrollment_choices()

    refresh_enrollments_table()

    refresh_payment_section()

    refresh_supplies_all()

    refresh_dashboard()

    refresh_class_readiness_all()

    refresh_reports_all()

    refresh_backup_file_status()


def restore_backup():

    # Let the user choose a Lux backup ZIP.

    backup_path = filedialog.askopenfilename(
        title="Restore Lux Backup",
        filetypes=[
            (
                "Lux Backup ZIP",
                "*.zip"
            )
        ]
    )


    if not backup_path:

        return


    # ---------------------------------------------
    # CHECK THE ZIP BEFORE TOUCHING LIVE DATA
    # ---------------------------------------------

    try:

        with zipfile.ZipFile(
            backup_path,
            "r"
        ) as backup_zip:

            zip_names = set(
                backup_zip.namelist()
            )


            if DATABASE_FILE not in zip_names:

                raise ValueError(
                    (
                        "This backup does not contain "
                        "lux_command_center.db."
                    )
                )


            backup_description = (
                os.path.basename(
                    backup_path
                )
            )


            if "backup_info.txt" in zip_names:

                info_text = backup_zip.read(
                    "backup_info.txt"
                ).decode(
                    "utf-8",
                    errors="replace"
                )


                backup_description = (
                    f"{backup_description}\n\n"
                    f"{info_text.strip()}"
                )


    except (
        zipfile.BadZipFile,
        ValueError
    ) as error:

        messagebox.showerror(
            "Invalid Backup",
            (
                "Lux cannot restore this file.\n\n"
                f"{error}"
            )
        )

        return


    confirmed = messagebox.askyesno(
        "Restore Backup?",
        (
            "Restoring will replace the CURRENT "
            "Lux database with the selected backup.\n\n"
            "Lux will first create an automatic "
            "safety backup of your current data.\n\n"
            f"{backup_description}\n\n"
            "Continue?"
        )
    )


    if not confirmed:

        return


    # ---------------------------------------------
    # SAFETY BACKUP CURRENT DATA FIRST
    # ---------------------------------------------

    safety_backup_path = (
        create_pre_restore_safety_backup()
    )


    if not safety_backup_path:

        messagebox.showerror(
            "Restore Stopped",
            (
                "Lux could not create the automatic "
                "pre-restore safety backup, so the "
                "restore was cancelled."
            )
        )

        return


    temp_folder = tempfile.mkdtemp(
        prefix="lux_restore_"
    )


    try:

        temp_database_path = os.path.join(
            temp_folder,
            DATABASE_FILE
        )


        temp_history_path = os.path.join(
            temp_folder,
            HISTORY_FILE
        )


        with zipfile.ZipFile(
            backup_path,
            "r"
        ) as backup_zip:

            # Extract ONLY the specific files we expect.
            #
            # We intentionally do not use extractall().
            # That prevents strange ZIP paths from
            # writing files in unexpected places.

            with backup_zip.open(
                DATABASE_FILE
            ) as source_file:

                with open(
                    temp_database_path,
                    "wb"
                ) as destination_file:

                    shutil.copyfileobj(
                        source_file,
                        destination_file
                    )


            history_is_in_backup = (
                HISTORY_FILE
                in backup_zip.namelist()
            )


            if history_is_in_backup:

                with backup_zip.open(
                    HISTORY_FILE
                ) as source_file:

                    with open(
                        temp_history_path,
                        "wb"
                    ) as destination_file:

                        shutil.copyfileobj(
                            source_file,
                            destination_file
                        )


        # Validate BEFORE replacing live data.

        validate_backup_database(
            temp_database_path
        )


        live_database_path = os.path.join(
            get_program_folder(),
            DATABASE_FILE
        )


        live_history_path = os.path.join(
            get_program_folder(),
            HISTORY_FILE
        )


        # Replace the database only after validation.

        shutil.copy2(
            temp_database_path,
            live_database_path
        )


        # Restore CSV history if this backup contains it.
        #
        # If an older backup does not include a CSV,
        # Lux leaves the current CSV alone rather than
        # deleting history unexpectedly.

        if history_is_in_backup:

            shutil.copy2(
                temp_history_path,
                live_history_path
            )

        # If this is an older valid database,
        # safely add any schema pieces introduced
        # in newer Lux versions.

        initialize_database()


        refresh_everything_after_restore()


        update_backup_status(
            (
                "Restored: "
                f"{os.path.basename(backup_path)}"
            )
        )


        messagebox.showinfo(
            "Restore Complete",
            (
                "Lux data was restored successfully.\n\n"
                "A safety copy of the data you had "
                "before restoring was saved here:\n\n"
                f"{safety_backup_path}"
            )
        )


    except Exception as error:

        messagebox.showerror(
            "Restore Failed",
            (
                "Lux could not restore this backup.\n\n"
                f"{error}\n\n"
                "Your pre-restore safety backup "
                "was still created here:\n\n"
                f"{safety_backup_path}"
            )
        )


    finally:

        shutil.rmtree(
            temp_folder,
            ignore_errors=True
        )


def open_local_backup_folder():

    # Windows can open a folder directly with
    # os.startfile().
    #
    # If Lux is ever run on another operating system,
    # the error message keeps the failure polite.

    backup_folder = get_local_backup_folder()


    try:

        os.startfile(
            backup_folder
        )

    except Exception:

        messagebox.showinfo(
            "Backup Folder",
            backup_folder
        )




# =====================================================
# v4.4 USER-FRIENDLY POLISH
# =====================================================

# v4.4 is less about adding new business data and more about
# making the existing system easy for a non-technical person
# to use every day.
#
# We add:
#
#   • Tooltips
#   • Keyboard shortcuts
#   • A bottom status bar
#   • A Help / Quick Guide tab
#   • A menu bar
#   • Friendlier navigation
#   • A maximized startup window
#
# None of these features change the underlying business data.


class ToolTip:

    # A tooltip is the little help bubble that appears when the
    # mouse rests over a button.
    #
    # It gives Becca plain-English help without permanently
    # filling the screen with instructions.

    def __init__(
        self,
        widget,
        text
    ):

        self.widget = widget

        self.text = text

        self.tip_window = None

        self.after_id = None


        self.widget.bind(
            "<Enter>",
            self.schedule,
            add="+"
        )


        self.widget.bind(
            "<Leave>",
            self.hide,
            add="+"
        )


        self.widget.bind(
            "<ButtonPress>",
            self.hide,
            add="+"
        )


    def schedule(
        self,
        event=None
    ):

        self.cancel_schedule()


        self.after_id = self.widget.after(
            500,
            self.show
        )


    def cancel_schedule(
        self
    ):

        if self.after_id is not None:

            self.widget.after_cancel(
                self.after_id
            )

            self.after_id = None


    def show(
        self
    ):

        if self.tip_window is not None:

            return


        try:

            x = (
                self.widget.winfo_rootx()
                + 20
            )

            y = (
                self.widget.winfo_rooty()
                + self.widget.winfo_height()
                + 5
            )

        except tk.TclError:

            return


        self.tip_window = tk.Toplevel(
            self.widget
        )


        self.tip_window.wm_overrideredirect(
            True
        )


        self.tip_window.wm_geometry(
            f"+{x}+{y}"
        )


        label = tk.Label(
            self.tip_window,
            text=self.text,
            justify="left",
            bg="#FFFBEA",
            fg="#333333",
            relief="solid",
            bd=1,
            font=("Arial", 9),
            padx=7,
            pady=5,
            wraplength=320
        )


        label.pack()


    def hide(
        self,
        event=None
    ):

        self.cancel_schedule()


        if self.tip_window is not None:

            self.tip_window.destroy()

            self.tip_window = None


def install_becca_tooltips():

    # Match button text to a friendly explanation.
    #
    # We walk through the whole application after every widget
    # has been created, so we do not need to manually rewrite
    # every existing button definition.

    tooltip_text = {

        "CALCULATE":
            "Run a what-if class profit calculation. This does not change student or payment records.",

        "SAVE PLAN":
            "Save this calculator scenario to the CSV planning history.",

        "SCHEDULE NEW CLASS":
            "Turn the calculator setup into a real scheduled class session.",

        "UPDATE SESSION":
            "Save changes to the scheduled class currently loaded into the Calculator.",

        "VIEW ROSTER":
            "Show the actual students enrolled in the selected class.",

        "READINESS":
            "Check whether the selected class is ready based on enrollment, money, certification tracking, and supplies.",

        "ADD STUDENT":
            "Create a new permanent student record.",

        "UPDATE LOADED":
            "Save changes to the student currently loaded in the form.",

        "ENROLL STUDENT":
            "Connect the selected student to the selected scheduled class.",

        "ADD PAYMENT":
            "Record a real payment against the loaded enrollment.",

        "PREVIEW / REFRESH":
            "Rebuild this report from the latest live database information.",

        "EXPORT CSV":
            "Save the current report as a CSV file for Excel or Google Sheets.",

        "QUICK BACKUP":
            "Create a timestamped backup ZIP inside the local LuxBackups folder.",

        "SAVE BACKUP AS...":
            "Create a backup ZIP somewhere you choose, such as OneDrive or a USB drive.",

        "RESTORE BACKUP...":
            "Restore an earlier Lux data snapshot. Lux creates a safety backup first.",

        "CLEAR FILTERS":
            "Remove all search filters and show every matching record again.",

        "REFRESH":
            "Reload this screen from the latest database information.",

        "REFRESH RESULTS":
            "Reload this screen from the latest database information.",

        "CLASS READINESS":
            "Open the all-in-one readiness checklist for a scheduled class.",

        "OPEN SUPPLIES":
            "Jump directly to inventory and class supply requirements."
    }


    def walk_widgets(
        parent
    ):

        for child in parent.winfo_children():

            if isinstance(
                child,
                tk.Button
            ):

                try:

                    button_text = child.cget(
                        "text"
                    )

                except tk.TclError:

                    button_text = ""


                if button_text in tooltip_text:

                    ToolTip(
                        child,
                        tooltip_text[button_text]
                    )


            walk_widgets(
                child
            )


    walk_widgets(
        window
    )


def refresh_app_status(
    event=None
):

    # The bottom status bar gives Becca a tiny always-visible
    # snapshot without needing to leave the current page.

    try:

        connection = sqlite3.connect(
            "lux_command_center.db"
        )


        cursor = connection.cursor()


        student_count = cursor.execute(
            """
            SELECT COUNT(*)
            FROM students
            """
        ).fetchone()[0]


        today_database = datetime.now().strftime(
            "%Y-%m-%d"
        )


        upcoming_count = cursor.execute(
            """
            SELECT COUNT(*)
            FROM class_sessions
            WHERE class_date >= ?
            AND status = 'Scheduled'
            """,
            (
                today_database,
            )
        ).fetchone()[0]


        outstanding = cursor.execute(
            """
            SELECT
                COALESCE(
                    SUM(
                        CASE
                            WHEN cs.tuition
                                 - COALESCE(p.total_paid, 0) > 0
                            THEN cs.tuition
                                 - COALESCE(p.total_paid, 0)
                            ELSE 0
                        END
                    ),
                    0
                )
            FROM enrollments e

            JOIN class_sessions cs
                ON cs.id = e.session_id

            LEFT JOIN (
                SELECT
                    enrollment_id,
                    SUM(amount) AS total_paid
                FROM payments
                GROUP BY enrollment_id
            ) p
                ON p.enrollment_id = e.id
            """
        ).fetchone()[0]


        connection.close()


        app_status_label.config(
            text=(
                f"Students: {student_count}"
                f"   |   Upcoming Classes: {upcoming_count}"
                f"   |   Outstanding: ${outstanding:,.2f}"
            )
        )


        app_status_right_label.config(
            text=(
                "Data saves automatically"
                f"   •   Lux v{APP_VERSION}"
            )
        )


    except Exception:

        # The status bar is convenience information.
        # A problem here should never stop the application.

        app_status_label.config(
            text="Lux data status unavailable"
        )


def refresh_current_view(
    event=None
):

    # F5 refreshes whichever main page Becca is currently viewing.

    selected_tab = notebook.select()


    if selected_tab == str(
        dashboard_tab
    ):

        refresh_dashboard()


    elif selected_tab == str(
        classes_tab
    ):

        refresh_class_sessions_table()

        refresh_selected_class_roster()


    elif selected_tab == str(
        class_readiness_tab
    ):

        refresh_class_readiness_all()


    elif selected_tab == str(
        students_tab
    ):

        refresh_students_table()


    elif selected_tab == str(
        enrollments_tab
    ):

        refresh_enrollment_choices()

        refresh_enrollments_table()

        refresh_payment_section()


    elif selected_tab == str(
        reports_tab
    ):

        refresh_reports_all()


    elif selected_tab == str(
        backup_tab
    ):

        refresh_backup_file_status()


    elif selected_tab == str(
        history_tab
    ):

        refresh_history_table()


    elif selected_tab == str(
        supplies_tab
    ):

        refresh_supplies_all()


    refresh_app_status()


def focus_current_search(
    event=None
):

    # Ctrl+F jumps directly into Search on the two pages where
    # searching is most useful.

    selected_tab = notebook.select()


    if selected_tab == str(
        students_tab
    ):

        student_search_entry.focus_set()

        student_search_entry.select_range(
            0,
            tk.END
        )


    elif selected_tab == str(
        classes_tab
    ):

        class_search_entry.focus_set()

        class_search_entry.select_range(
            0,
            tk.END
        )


    else:

        notebook.select(
            students_tab
        )

        student_search_entry.focus_set()


    return "break"


def show_about_lux():

    messagebox.showinfo(
        "About Lux Class Command Center",
        (
            "Lux Class Command Center\n"
            f"Version {APP_VERSION}\n\n"
            "Built for Lux Beauty Lash Academy.\n\n"
            "Students • Classes • Enrollments • Payments\n"
            "Readiness • Inventory • Reports • Backups\n\n"
            "Business data is stored locally in SQLite."
        )
    )


def open_help_tab():

    notebook.select(
        help_tab
    )


def on_notebook_changed(
    event=None
):

    # Update the status bar whenever Becca changes pages.

    refresh_app_status()



# =====================================================
# COLORS
# =====================================================

lux_pink = "#F4C6D7"
soft_pink = "#FFF5F8"
dark_text = "#333333"
gold = "#B58A3A"


# =====================================================
# MAIN WINDOW
# =====================================================

window = tk.Tk()


window.title(
    "Lux Class Command Center v4.4"
)


screen_height = (
    window.winfo_screenheight()
)


window_height = min(
    680,
    screen_height - 80
)


window.geometry(
    f"950x{window_height}"
)


window.configure(
    bg=soft_pink
)


# Start maximized on Windows when possible.
#
# The app still keeps its normal geometry as a fallback.

try:

    window.state(
        "zoomed"
    )

except tk.TclError:

    pass


# Give ttk widgets a cleaner, more consistent appearance.

lux_style = ttk.Style(
    window
)


lux_style.configure(
    "TNotebook.Tab",
    font=("Arial", 9, "bold"),
    padding=(9, 6)
)


lux_style.configure(
    "Treeview",
    font=("Arial", 9),
    rowheight=25
)


lux_style.configure(
    "Treeview.Heading",
    font=("Arial", 9, "bold")
)


# -----------------------------------------------------
# ALWAYS-VISIBLE STATUS BAR
# -----------------------------------------------------

app_status_bar = tk.Frame(
    window,
    bg=dark_text
)


app_status_bar.pack(
    side="bottom",
    fill="x"
)


app_status_label = tk.Label(
    app_status_bar,
    text="Loading Lux data...",
    font=("Arial", 9),
    bg=dark_text,
    fg="white",
    anchor="w"
)


app_status_label.pack(
    side="left",
    padx=10,
    pady=4
)


app_status_right_label = tk.Label(
    app_status_bar,
    text=f"Lux v{APP_VERSION}",
    font=("Arial", 9),
    bg=dark_text,
    fg="white",
    anchor="e"
)


app_status_right_label.pack(
    side="right",
    padx=10,
    pady=4
)


# =====================================================
# COMMAND CENTER TABS
# =====================================================

notebook = ttk.Notebook(
    window
)


notebook.pack(
    fill="both",
    expand=True
)


dashboard_tab = tk.Frame(
    notebook,
    bg=soft_pink
)


calculator_tab = tk.Frame(
    notebook,
    bg=soft_pink
)


classes_tab = tk.Frame(
    notebook,
    bg=soft_pink
)


class_readiness_tab = tk.Frame(
    notebook,
    bg=soft_pink
)


history_tab = tk.Frame(
    notebook,
    bg=soft_pink
)


students_tab = tk.Frame(
    notebook,
    bg=soft_pink
)


enrollments_tab = tk.Frame(
    notebook,
    bg=soft_pink
)


reports_tab = tk.Frame(
    notebook,
    bg=soft_pink
)


backup_tab = tk.Frame(
    notebook,
    bg=soft_pink
)


help_tab = tk.Frame(
    notebook,
    bg=soft_pink
)


supplies_tab = tk.Frame(
    notebook,
    bg=soft_pink
)


notebook.add(
    dashboard_tab,
    text="Dashboard"
)


notebook.add(
    calculator_tab,
    text="Calculator"
)


notebook.add(
    classes_tab,
    text="Classes"
)


notebook.add(
    class_readiness_tab,
    text="Readiness"
)


notebook.add(
    students_tab,
    text="Students"
)


notebook.add(
    enrollments_tab,
    text="Enrollments"
)


notebook.add(
    reports_tab,
    text="Reports"
)


notebook.add(
    backup_tab,
    text="Backup"
)


notebook.add(
    history_tab,
    text="History"
)


notebook.add(
    supplies_tab,
    text="Supplies"
)


notebook.add(
    help_tab,
    text="Help"
)


# =====================================================
# MENU BAR + KEYBOARD SHORTCUTS
# =====================================================

menu_bar = tk.Menu(
    window
)


file_menu = tk.Menu(
    menu_bar,
    tearoff=0
)


file_menu.add_command(
    label="Quick Backup",
    command=quick_backup
)


file_menu.add_command(
    label="Save Backup As...",
    command=save_backup_as
)


file_menu.add_separator()


file_menu.add_command(
    label="Exit",
    command=window.destroy
)


menu_bar.add_cascade(
    label="File",
    menu=file_menu
)


go_menu = tk.Menu(
    menu_bar,
    tearoff=0
)


go_menu.add_command(
    label="Dashboard",
    command=lambda: notebook.select(
        dashboard_tab
    )
)


go_menu.add_command(
    label="Students",
    command=lambda: notebook.select(
        students_tab
    )
)


go_menu.add_command(
    label="Classes",
    command=lambda: notebook.select(
        classes_tab
    )
)


go_menu.add_command(
    label="Enrollments / Payments",
    command=lambda: notebook.select(
        enrollments_tab
    )
)


go_menu.add_command(
    label="Readiness",
    command=lambda: notebook.select(
        class_readiness_tab
    )
)


go_menu.add_command(
    label="Reports",
    command=lambda: notebook.select(
        reports_tab
    )
)


go_menu.add_command(    label="Supplies",
    command=lambda: notebook.select(
        supplies_tab
    )
)


go_menu.add_command(
    label="Backup",
    command=lambda: notebook.select(
        backup_tab
    )
)


menu_bar.add_cascade(
    label="Go To",
    menu=go_menu
)


help_menu = tk.Menu(
    menu_bar,
    tearoff=0
)


help_menu.add_command(
    label="Becca's Quick Guide",
    command=open_help_tab
)


help_menu.add_command(
    label="About Lux",
    command=show_about_lux
)


menu_bar.add_cascade(
    label="Help",
    menu=help_menu
)


window.config(
    menu=menu_bar
)


# Keyboard shortcuts keep frequent jobs fast.
#
# Ctrl+D  Dashboard
# Ctrl+S  Students
# Ctrl+E  Enrollments / Payments
# Ctrl+R  Readiness
# Ctrl+B  Backup
# Ctrl+H  Help
# Ctrl+F  Search
# F5      Refresh current page

window.bind_all(
    "<Control-d>",
    lambda event: notebook.select(
        dashboard_tab
    )
)


window.bind_all(
    "<Control-s>",
    lambda event: notebook.select(
        students_tab
    )
)


window.bind_all(
    "<Control-e>",
    lambda event: notebook.select(
        enrollments_tab
    )
)


window.bind_all(
    "<Control-r>",
    lambda event: notebook.select(
        class_readiness_tab
    )
)


window.bind_all(
    "<Control-b>",
    lambda event: notebook.select(
        backup_tab
    )
)


window.bind_all(
    "<Control-h>",
    lambda event: notebook.select(
        help_tab
    )
)


window.bind_all(
    "<Control-f>",
    focus_current_search
)


window.bind_all(
    "<F5>",
    refresh_current_view
)


notebook.bind(
    "<<NotebookTabChanged>>",
    on_notebook_changed
)


# =====================================================
# DASHBOARD TAB
# =====================================================

# The Dashboard is intentionally a SUMMARY screen.
# It reads the database and helps Becca answer:
#
# What is coming up?
# How many students are enrolled?
# How much tuition is expected?
# How much has actually been collected?
# Who still owes money?
#
# The underlying records still live in Students,
# Classes, Enrollments, and Payments.


# -----------------------------------------------------
# SCROLLABLE DASHBOARD PAGE
# -----------------------------------------------------

# The Dashboard contains several business summaries.
# Giving the whole page a scrollbar keeps it usable on
# shorter laptop screens instead of playing GUI Tetris.

dashboard_main_container = tk.Frame(
    dashboard_tab,
    bg=soft_pink
)


dashboard_main_container.pack(
    fill="both",
    expand=True
)


dashboard_canvas = tk.Canvas(
    dashboard_main_container,
    bg=soft_pink,
    highlightthickness=0
)


dashboard_scrollbar = ttk.Scrollbar(
    dashboard_main_container,
    orient="vertical",
    command=dashboard_canvas.yview
)


dashboard_scrollable_frame = tk.Frame(
    dashboard_canvas,
    bg=soft_pink
)


dashboard_scrollable_frame.bind(
    "<Configure>",
    lambda event: dashboard_canvas.configure(
        scrollregion=dashboard_canvas.bbox("all")
    )
)


dashboard_canvas_window = dashboard_canvas.create_window(
    (0, 0),
    window=dashboard_scrollable_frame,
    anchor="nw"
)


def resize_dashboard_scrollable_frame(event):

    dashboard_canvas.itemconfigure(
        dashboard_canvas_window,
        width=event.width
    )


dashboard_canvas.bind(
    "<Configure>",
    resize_dashboard_scrollable_frame
)


dashboard_canvas.configure(
    yscrollcommand=dashboard_scrollbar.set
)


dashboard_canvas.pack(
    side="left",
    fill="both",
    expand=True
)


dashboard_scrollbar.pack(
    side="right",
    fill="y"
)


def dashboard_mousewheel_scroll(event):

    if notebook.select() == str(dashboard_tab):

        dashboard_canvas.yview_scroll(
            int(
                -1 * (event.delta / 120)
            ),
            "units"
        )


window.bind_all(
    "<MouseWheel>",
    dashboard_mousewheel_scroll,
    add="+"
)


# -----------------------------------------------------
# DASHBOARD HEADER
# -----------------------------------------------------

dashboard_header = tk.Frame(
    dashboard_scrollable_frame,
    bg=lux_pink
)


dashboard_header.pack(
    fill="x"
)


tk.Label(
    dashboard_header,
    text="LUX BEAUTY LASH ACADEMY",
    font=("Arial", 20, "bold"),
    bg=lux_pink,
    fg=dark_text
).pack(
    pady=(10, 1)
)


tk.Label(
    dashboard_header,
    text="COMMAND CENTER DASHBOARD",
    font=("Arial", 12, "bold"),
    bg=lux_pink,
    fg=gold
).pack(
    pady=(0, 3)
)


tk.Label(
    dashboard_header,
    text="Becca's at-a-glance home base • Start here, then jump straight to the job you need.",
    font=("Arial", 9, "italic"),
    bg=lux_pink,
    fg=dark_text
).pack(
    pady=(0, 7)
)


dashboard_last_updated_label = tk.Label(
    dashboard_header,
    text="",
    font=("Arial", 8),
    bg=lux_pink,
    fg=dark_text
)


dashboard_last_updated_label.pack(
    pady=(0, 8)
)


# -----------------------------------------------------
# NEXT CLASS
# -----------------------------------------------------

dashboard_next_class_frame = tk.Frame(
    dashboard_scrollable_frame,
    bg="white",
    bd=2,
    relief="groove"
)


dashboard_next_class_frame.pack(
    fill="x",
    padx=18,
    pady=(10, 6)
)


tk.Label(
    dashboard_next_class_frame,
    text="NEXT CLASS",
    font=("Arial", 10, "bold"),
    bg="white",
    fg=gold
).pack(
    pady=(7, 1)
)


dashboard_next_class_title = tk.Label(
    dashboard_next_class_frame,
    text="No upcoming class sessions",
    font=("Arial", 14, "bold"),
    bg="white",
    fg=dark_text
)


dashboard_next_class_title.pack()


dashboard_next_class_schedule = tk.Label(
    dashboard_next_class_frame,
    text="",
    font=("Arial", 11, "bold"),
    bg="white",
    fg=dark_text
)


dashboard_next_class_schedule.pack(
    pady=(1, 1)
)


dashboard_next_class_details = tk.Label(
    dashboard_next_class_frame,
    text="",
    font=("Arial", 9),
    bg="white",
    fg=dark_text
)


dashboard_next_class_details.pack(
    pady=(0, 7)
)


# -----------------------------------------------------
# SUMMARY CARDS
# -----------------------------------------------------

dashboard_cards_frame = tk.Frame(
    dashboard_scrollable_frame,
    bg=soft_pink
)


dashboard_cards_frame.pack(
    fill="x",
    padx=14,
    pady=5
)


for column_number in range(5):

    dashboard_cards_frame.columnconfigure(
        column_number,
        weight=1
    )


def make_dashboard_card(column, title):

    card = tk.Frame(
        dashboard_cards_frame,
        bg="white",
        bd=1,
        relief="groove"
    )

    card.grid(
        row=0,
        column=column,
        sticky="nsew",
        padx=3
    )


    tk.Label(
        card,
        text=title,
        font=("Arial", 8, "bold"),
        bg="white",
        fg=dark_text,
        wraplength=150
    ).pack(
        pady=(7, 2)
    )


    value_label = tk.Label(
        card,
        text="0",
        font=("Arial", 14, "bold"),
        bg="white",
        fg=gold
    )

    value_label.pack(
        pady=(0, 7)
    )


    return value_label


dashboard_enrollment_value = make_dashboard_card(
    0,
    "ACTIVE ENROLLMENTS"
)


dashboard_expected_value = make_dashboard_card(
    1,
    "EXPECTED TUITION"
)


dashboard_collected_value = make_dashboard_card(
    2,
    "COLLECTED"
)


dashboard_outstanding_value = make_dashboard_card(
    3,
    "OUTSTANDING"
)


dashboard_balance_count_value = make_dashboard_card(
    4,
    "STUDENTS WITH BALANCE"
)


# -----------------------------------------------------
# QUICK ACTIONS
# -----------------------------------------------------

dashboard_quick_frame = tk.Frame(
    dashboard_scrollable_frame,
    bg=soft_pink
)


dashboard_quick_frame.pack(
    pady=6
)


tk.Button(
    dashboard_quick_frame,
    text="VIEW NEXT / SELECTED CLASS",
    font=("Arial", 9, "bold"),
    bg=gold,
    fg="white",
    padx=10,
    pady=6,
    command=dashboard_view_selected_class
).grid(
    row=0,
    column=0,
    padx=3
)


tk.Button(
    dashboard_quick_frame,
    text="ENROLLMENTS / PAYMENTS",
    font=("Arial", 9, "bold"),
    bg=lux_pink,
    fg=dark_text,
    padx=10,
    pady=6,
    command=dashboard_open_enrollments
).grid(
    row=0,
    column=1,
    padx=3
)


tk.Button(
    dashboard_quick_frame,
    text="ADD / VIEW STUDENTS",
    font=("Arial", 9, "bold"),
    padx=10,
    pady=6,
    command=dashboard_open_students
).grid(
    row=0,
    column=2,
    padx=3
)


tk.Button(
    dashboard_quick_frame,
    text="CALCULATOR",
    font=("Arial", 9, "bold"),
    padx=10,
    pady=6,
    command=dashboard_open_calculator
).grid(
    row=0,
    column=3,
    padx=3
)


tk.Button(
    dashboard_quick_frame,
    text="REFRESH",
    font=("Arial", 9, "bold"),
    padx=10,
    pady=6,
    command=refresh_dashboard
).grid(
    row=0,
    column=4,
    padx=3
)


tk.Button(
    dashboard_quick_frame,
    text="CLASS READINESS",
    font=("Arial", 9, "bold"),
    bg=lux_pink,
    fg=dark_text,
    padx=10,
    pady=6,
    command=dashboard_open_readiness
).grid(
    row=1,
    column=0,
    columnspan=5,
    pady=(5, 0)
)


# -----------------------------------------------------
# UPCOMING CLASSES TABLE
# -----------------------------------------------------

tk.Label(
    dashboard_scrollable_frame,
    text="UPCOMING CLASSES",
    font=("Arial", 12, "bold"),
    bg=lux_pink,
    fg=dark_text
).pack(
    fill="x",
    pady=(5, 0)
)


dashboard_upcoming_frame = tk.Frame(
    dashboard_scrollable_frame,
    bg=soft_pink
)


dashboard_upcoming_frame.pack(
    fill="x",
    padx=12,
    pady=6
)


dashboard_upcoming_columns = (
    "id",
    "date",
    "time",
    "class",
    "enrolled",
    "seats",
    "collected",
    "outstanding",
    "status"
)


dashboard_upcoming_table = ttk.Treeview(
    dashboard_upcoming_frame,
    columns=dashboard_upcoming_columns,
    show="headings",
    height=5
)


for column_name, heading_text in (
    ("id", "ID"),
    ("date", "Date"),
    ("time", "Time"),
    ("class", "Class"),
    ("enrolled", "Enrolled"),
    ("seats", "Seats Left"),
    ("collected", "Collected"),
    ("outstanding", "Outstanding"),
    ("status", "Status")
):

    dashboard_upcoming_table.heading(
        column_name,
        text=heading_text
    )


dashboard_upcoming_table.column(
    "id",
    width=40,
    anchor="center"
)


dashboard_upcoming_table.column(
    "date",
    width=90,
    anchor="center"
)


dashboard_upcoming_table.column(
    "time",
    width=80,
    anchor="center"
)


dashboard_upcoming_table.column(
    "class",
    width=120
)


dashboard_upcoming_table.column(
    "enrolled",
    width=75,
    anchor="center"
)


dashboard_upcoming_table.column(
    "seats",
    width=70,
    anchor="center"
)


dashboard_upcoming_table.column(
    "collected",
    width=95,
    anchor="e"
)


dashboard_upcoming_table.column(
    "outstanding",
    width=100,
    anchor="e"
)


dashboard_upcoming_table.column(
    "status",
    width=85,
    anchor="center"
)


dashboard_upcoming_vertical = ttk.Scrollbar(
    dashboard_upcoming_frame,
    orient="vertical",
    command=dashboard_upcoming_table.yview
)


dashboard_upcoming_horizontal = ttk.Scrollbar(
    dashboard_upcoming_frame,
    orient="horizontal",
    command=dashboard_upcoming_table.xview
)


dashboard_upcoming_table.configure(
    yscrollcommand=dashboard_upcoming_vertical.set,
    xscrollcommand=dashboard_upcoming_horizontal.set
)


dashboard_upcoming_table.grid(
    row=0,
    column=0,
    sticky="nsew"
)


dashboard_upcoming_vertical.grid(
    row=0,
    column=1,
    sticky="ns"
)


dashboard_upcoming_horizontal.grid(
    row=1,
    column=0,
    sticky="ew"
)


dashboard_upcoming_frame.columnconfigure(
    0,
    weight=1
)


dashboard_upcoming_table.bind(
    "<Double-1>",
    lambda event: dashboard_view_selected_class()
)


# -----------------------------------------------------
# NEEDS ATTENTION
# -----------------------------------------------------

tk.Label(
    dashboard_scrollable_frame,
    text="NEEDS ATTENTION",
    font=("Arial", 12, "bold"),
    bg=lux_pink,
    fg=dark_text
).pack(
    fill="x",
    pady=(7, 0)
)


dashboard_attention_status = tk.Label(
    dashboard_scrollable_frame,
    text="",
    font=("Arial", 9, "italic"),
    bg=soft_pink,
    fg=dark_text
)


dashboard_attention_status.pack(
    pady=(4, 1)
)


dashboard_attention_frame = tk.Frame(
    dashboard_scrollable_frame,
    bg=soft_pink
)


dashboard_attention_frame.pack(
    fill="x",
    padx=12,
    pady=(3, 14)
)


dashboard_attention_columns = (
    "enrollment_id",
    "student",
    "class",
    "date",
    "time",
    "paid",
    "balance"
)


dashboard_attention_table = ttk.Treeview(
    dashboard_attention_frame,
    columns=dashboard_attention_columns,
    show="headings",
    height=5
)


for column_name, heading_text in (
    ("enrollment_id", "Enrollment ID"),
    ("student", "Student"),
    ("class", "Class"),
    ("date", "Date"),
    ("time", "Time"),
    ("paid", "Paid"),
    ("balance", "Balance Due")
):

    dashboard_attention_table.heading(
        column_name,
        text=heading_text
    )


dashboard_attention_table.column(
    "enrollment_id",
    width=80,
    anchor="center"
)


dashboard_attention_table.column(
    "student",
    width=150
)


dashboard_attention_table.column(
    "class",
    width=120
)


dashboard_attention_table.column(
    "date",
    width=90,
    anchor="center"
)


dashboard_attention_table.column(
    "time",
    width=80,
    anchor="center"
)


dashboard_attention_table.column(
    "paid",
    width=90,
    anchor="e"
)


dashboard_attention_table.column(
    "balance",
    width=100,
    anchor="e"
)


dashboard_attention_vertical = ttk.Scrollbar(
    dashboard_attention_frame,
    orient="vertical",
    command=dashboard_attention_table.yview
)


dashboard_attention_horizontal = ttk.Scrollbar(
    dashboard_attention_frame,
    orient="horizontal",
    command=dashboard_attention_table.xview
)


dashboard_attention_table.configure(
    yscrollcommand=dashboard_attention_vertical.set,
    xscrollcommand=dashboard_attention_horizontal.set
)


dashboard_attention_table.grid(
    row=0,
    column=0,
    sticky="nsew"
)


dashboard_attention_vertical.grid(
    row=0,
    column=1,
    sticky="ns"
)


dashboard_attention_horizontal.grid(
    row=1,
    column=0,
    sticky="ew"
)


dashboard_attention_frame.columnconfigure(
    0,
    weight=1
)


# =====================================================
# CALCULATOR TAB
# =====================================================

main_container = tk.Frame(
    calculator_tab,
    bg=soft_pink
)


main_container.pack(
    fill="both",
    expand=True
)


canvas = tk.Canvas(
    main_container,
    bg=soft_pink,
    highlightthickness=0
)


scrollbar = ttk.Scrollbar(
    main_container,
    orient="vertical",
    command=canvas.yview
)


scrollable_frame = tk.Frame(
    canvas,
    bg=soft_pink
)


scrollable_frame.bind(
    "<Configure>",
    lambda event: canvas.configure(
        scrollregion=canvas.bbox("all")
    )
)


canvas_window = canvas.create_window(
    (0, 0),
    window=scrollable_frame,
    anchor="nw"
)


def resize_scrollable_frame(event):

    canvas.itemconfigure(
        canvas_window,
        width=event.width
    )

canvas.bind(
    "<Configure>",
    resize_scrollable_frame
)


canvas.configure(
    yscrollcommand=scrollbar.set
)


canvas.pack(
    side="left",
    fill="both",
    expand=True
)


scrollbar.pack(
    side="right",
    fill="y"
)


def mousewheel_scroll(event):

    if notebook.select() == str(calculator_tab):

        canvas.yview_scroll(
            int(
                -1 * (event.delta / 120)
            ),
            "units"
        )


window.bind_all(
    "<MouseWheel>",
    mousewheel_scroll
)


# =====================================================
# CALCULATOR HEADER
# =====================================================

header = tk.Frame(
    scrollable_frame,
    bg=lux_pink
)


header.pack(
    fill="x"
)


title = tk.Label(
    header,
    text="LUX BEAUTY LASH ACADEMY",
    font=("Arial", 20, "bold"),
    bg=lux_pink,
    fg=dark_text
)


title.pack(
    pady=(12, 3)
)


subtitle = tk.Label(
    header,
    text="Class Command Center",
    font=("Arial", 13),
    bg=lux_pink,
    fg=dark_text
)


subtitle.pack(
    pady=(0, 12)
)


# =====================================================
# CALCULATOR INPUTS
# =====================================================

input_frame = tk.Frame(
    scrollable_frame,
    bg=soft_pink
)


input_frame.pack(
    pady=8
)


# -----------------------------------------------------
# CLASS TYPE
# -----------------------------------------------------

tk.Label(
    input_frame,
    text="Class Type",
    font=("Arial", 11, "bold"),
    bg=soft_pink
).grid(
    row=0,
    column=0,
    sticky="w",
    pady=5
)


class_combo = ttk.Combobox(
    input_frame,
    values=list(
        class_prices.keys()
    ),
    state="readonly",
    width=18
)


class_combo.set(
    "Intro Class"
)


class_combo.grid(
    row=0,
    column=1,
    padx=15
)


class_combo.bind(
    "<<ComboboxSelected>>",
    class_selected
)


# -----------------------------------------------------
# PLANNED STUDENTS
# -----------------------------------------------------

tk.Label(
    input_frame,
    text="Number of Students",
    font=("Arial", 11, "bold"),
    bg=soft_pink
).grid(
    row=1,
    column=0,
    sticky="w",
    pady=5
)


students_entry = tk.Entry(
    input_frame,
    font=("Arial", 12),
    width=15,
    justify="center"
)


students_entry.grid(
    row=1,
    column=1,
    padx=15
)


# -----------------------------------------------------
# TUITION
# -----------------------------------------------------

tk.Label(
    input_frame,
    text="Tuition Per Student",
    font=("Arial", 11, "bold"),
    bg=soft_pink
).grid(
    row=2,
    column=0,
    sticky="w",
    pady=5
)


tuition_entry = tk.Entry(
    input_frame,
    font=("Arial", 12),
    width=15,
    justify="center"
)


tuition_entry.insert(
    0,
    "399"
)


tuition_entry.grid(
    row=2,
    column=1,
    padx=15
)


# -----------------------------------------------------
# KIT COST
# -----------------------------------------------------

tk.Label(
    input_frame,
    text="Kit Cost Per Student",
    font=("Arial", 11, "bold"),
    bg=soft_pink
).grid(
    row=3,
    column=0,
    sticky="w",
    pady=5
)


kit_entry = tk.Entry(
    input_frame,
    font=("Arial", 12),
    width=15,
    justify="center"
)


kit_entry.insert(
    0,
    "75"
)


kit_entry.grid(
    row=3,
    column=1,
    padx=15
)


# -----------------------------------------------------
# OTHER EXPENSES
# -----------------------------------------------------

tk.Label(
    input_frame,
    text="Other Class Expenses",
    font=("Arial", 11, "bold"),
    bg=soft_pink
).grid(
    row=4,
    column=0,
    sticky="w",
    pady=5
)


other_entry = tk.Entry(
    input_frame,
    font=("Arial", 12),
    width=15,
    justify="center"
)


other_entry.insert(
    0,
    "0"
)


other_entry.grid(
    row=4,
    column=1,
    padx=15
)


# -----------------------------------------------------
# DESIRED PROFIT
# -----------------------------------------------------

tk.Label(
    input_frame,
    text="Desired Profit",
    font=("Arial", 11, "bold"),
    bg=soft_pink
).grid(
    row=5,
    column=0,
    sticky="w",
    pady=5
)


target_profit_entry = tk.Entry(
    input_frame,
    font=("Arial", 12),
    width=15,
    justify="center"
)


target_profit_entry.insert(
    0,
    "0"
)


target_profit_entry.grid(
    row=5,
    column=1,
    padx=15
)


# -----------------------------------------------------
# CAPACITY
# -----------------------------------------------------

tk.Label(
    input_frame,
    text="Class Capacity",
    font=("Arial", 11, "bold"),
    bg=soft_pink
).grid(
    row=6,
    column=0,
    sticky="w",
    pady=5
)


capacity_entry = tk.Entry(
    input_frame,
    font=("Arial", 12),
    width=15,
    justify="center"
)


capacity_entry.insert(
    0,
    "6"
)


capacity_entry.grid(
    row=6,
    column=1,
    padx=15
)


# -----------------------------------------------------
# CLASS DATE
# -----------------------------------------------------

tk.Label(
    input_frame,
    text="Class Date",
    font=("Arial", 11, "bold"),
    bg=soft_pink
).grid(
    row=7,
    column=0,
    sticky="w",
    pady=5
)


class_date_entry = tk.Entry(
    input_frame,
    font=("Arial", 12),
    width=15,
    justify="center"
)


class_date_entry.grid(
    row=7,
    column=1,
    padx=15
)


tk.Label(
    input_frame,
    text="Example: 10/24/2026",
    font=("Arial", 8),
    bg=soft_pink,
    fg=dark_text
).grid(
    row=8,
    column=1,
    pady=(0, 3)
)


# -----------------------------------------------------
# CLASS TIME
# -----------------------------------------------------

tk.Label(
    input_frame,
    text="Class Time",
    font=("Arial", 11, "bold"),
    bg=soft_pink
).grid(
    row=9,
    column=0,
    sticky="w",
    pady=5
)


class_time_entry = tk.Entry(
    input_frame,
    font=("Arial", 12),
    width=15,
    justify="center"
)


class_time_entry.grid(
    row=9,
    column=1,
    padx=15
)


tk.Label(
    input_frame,
    text="Example: 9:00 AM",
    font=("Arial", 8),
    bg=soft_pink,
    fg=dark_text
).grid(
    row=10,
    column=1,
    pady=(0, 3)
)


# Shows whether we are creating
# or editing a scheduled session.

session_edit_status_label = tk.Label(
    scrollable_frame,
    text="Scheduling a new class session",
    font=("Arial", 10, "italic"),
    bg=soft_pink,
    fg=gold
)


session_edit_status_label.pack(
    pady=(0, 4)
)


# =====================================================
# CALCULATOR BUTTONS
# =====================================================

button_frame = tk.Frame(
    scrollable_frame,
    bg=soft_pink
)


button_frame.pack(
    pady=6
)


calculate_button = tk.Button(
    button_frame,
    text="CALCULATE",
    font=("Arial", 11, "bold"),
    bg=gold,
    fg="white",
    padx=25,
    pady=8,
    command=calculate
)


calculate_button.grid(
    row=0,
    column=0,
    padx=10
)


reset_button = tk.Button(
    button_frame,
    text="RESET",
    font=("Arial", 11, "bold"),
    padx=25,
    pady=8,
    command=reset
)


reset_button.grid(
    row=0,
    column=1,
    padx=10
)


save_button = tk.Button(
    button_frame,
    text="SAVE PLAN",
    font=("Arial", 11, "bold"),
    bg=lux_pink,
    fg=dark_text,
    padx=25,
    pady=8,
    command=save_plan
)


save_button.grid(
    row=1,
    column=0,
    padx=5,
    pady=(8, 0)
)


update_button = tk.Button(
    button_frame,
    text="UPDATE PLAN",
    font=("Arial", 11, "bold"),
    padx=25,
    pady=8,
    command=update_loaded_plan
)


update_button.grid(
    row=1,
    column=1,
    padx=5,
    pady=(8, 0)
)


# CREATE SQLite class session.

schedule_button = tk.Button(
    button_frame,
    text="SCHEDULE NEW CLASS",
    font=("Arial", 11, "bold"),
    bg=gold,
    fg="white",
    padx=20,
    pady=8,
    command=schedule_class_session
)


schedule_button.grid(
    row=2,
    column=0,
    padx=5,
    pady=(8, 0)
)


# UPDATE SQLite class session.

update_session_button = tk.Button(
    button_frame,
    text="UPDATE SESSION",
    font=("Arial", 11, "bold"),
    bg=lux_pink,
    fg=dark_text,
    padx=20,
    pady=8,
    command=update_loaded_session
)


update_session_button.grid(
    row=2,
    column=1,
    padx=5,
    pady=(8, 0)
)


history_button = tk.Button(
    button_frame,
    text="GO TO HISTORY",
    font=("Arial", 11, "bold"),
    padx=25,
    pady=8,
    command=go_to_history
)


history_button.grid(
    row=3,
    column=0,
    columnspan=2,
    pady=(8, 0)
)


# =====================================================
# CALCULATOR RESULTS
# =====================================================

results_frame = tk.Frame(
    scrollable_frame,
    bg="white",
    bd=2,
    relief="groove"
)


results_frame.pack(
    padx=40,
    pady=8,
    fill="x"
)


tk.Label(
    results_frame,
    text="CLASS SUMMARY",
    font=("Arial", 13, "bold"),
    bg="white"
).pack(
    pady=(8, 4)
)


tk.Label(
    results_frame,
    text="Revenue",
    font=("Arial", 10),
    bg="white"
).pack()


revenue_result = tk.Label(
    results_frame,
    text="$0.00",
    font=("Arial", 14, "bold"),
    bg="white"
)


revenue_result.pack(
    pady=(0, 4)
)


tk.Label(
    results_frame,
    text="Total Expenses",
    font=("Arial", 10),
    bg="white"
).pack()


expenses_result = tk.Label(
    results_frame,
    text="$0.00",
    font=("Arial", 14, "bold"),
    bg="white"
)


expenses_result.pack(
    pady=(0, 4)
)


tk.Label(
    results_frame,
    text="Estimated Profit",
    font=("Arial", 10, "bold"),
    bg="white"
).pack()


profit_result = tk.Label(
    results_frame,
    text="$0.00",
    font=("Arial", 18, "bold"),
    bg="white",
    fg=gold
)


profit_result.pack(
    pady=(0, 4)
)


tk.Label(
    results_frame,
    text="Profit Per Student",
    font=("Arial", 10),
    bg="white"
).pack()


profit_per_student_result = tk.Label(
    results_frame,
    text="$0.00",
    font=("Arial", 13, "bold"),
    bg="white"
)


profit_per_student_result.pack(
    pady=(0, 4)
)


tk.Label(
    results_frame,
    text="Profit Margin",
    font=("Arial", 10),
    bg="white"
).pack()


profit_margin_result = tk.Label(
    results_frame,
    text="0.0%",
    font=("Arial", 13, "bold"),
    bg="white"
)


profit_margin_result.pack(
    pady=(0, 4)
)


tk.Label(
    results_frame,
    text="Break-Even Enrollment",
    font=("Arial", 10),
    bg="white"
).pack()


break_even_result = tk.Label(
    results_frame,
    text="0 students",
    font=("Arial", 13, "bold"),
    bg="white",
    fg=gold
)


break_even_result.pack(
    pady=(0, 4)
)


tk.Label(
    results_frame,
    text="Students Needed for Profit Goal",
    font=("Arial", 10),
    bg="white"
).pack()


goal_result = tk.Label(
    results_frame,
    text="0 students",
    font=("Arial", 13, "bold"),
    bg="white",
    fg=gold
)


goal_result.pack(
    pady=(0, 2)
)


goal_status_result = tk.Label(
    results_frame,
    text="",
    font=("Arial", 11, "bold"),
    bg="white",
    fg=dark_text
)


goal_status_result.pack(
    pady=(0, 6)
)


tk.Label(
    results_frame,
    text="Capacity Status",
    font=("Arial", 10),
    bg="white"
).pack()


capacity_result = tk.Label(
    results_frame,
    text="",
    font=("Arial", 12, "bold"),
    bg="white",
    fg=dark_text
)


capacity_result.pack(
    pady=(0, 5)
)


tk.Label(
    results_frame,
    text="Profit Goal Capacity Check",
    font=("Arial", 10),
    bg="white"
).pack()


goal_capacity_result = tk.Label(
    results_frame,
    text="",
    font=("Arial", 11, "bold"),
    bg="white",
    fg=gold
)


goal_capacity_result.pack(
    pady=(0, 10)
)



# =====================================================
# v4.0 CLASS READINESS TAB
# =====================================================

# This page combines several database areas into one
# operational answer:
#
# "Can Lux actually run this class?"
#
# Readiness is calculated live. Nothing here asks Becca
# to maintain a second copy of the same totals.


class_readiness_main = tk.Frame(
    class_readiness_tab,
    bg=soft_pink
)


class_readiness_main.pack(
    fill="both",
    expand=True
)


class_readiness_canvas = tk.Canvas(
    class_readiness_main,
    bg=soft_pink,
    highlightthickness=0
)


class_readiness_scrollbar = ttk.Scrollbar(
    class_readiness_main,
    orient="vertical",
    command=class_readiness_canvas.yview
)


class_readiness_scrollable_frame = tk.Frame(
    class_readiness_canvas,
    bg=soft_pink
)


class_readiness_scrollable_frame.bind(
    "<Configure>",
    lambda event: class_readiness_canvas.configure(
        scrollregion=class_readiness_canvas.bbox(
            "all"
        )
    )
)


class_readiness_canvas_window = (
    class_readiness_canvas.create_window(
        (0, 0),
        window=class_readiness_scrollable_frame,
        anchor="nw"
    )
)


def resize_class_readiness_frame(
    event
):

    class_readiness_canvas.itemconfigure(
        class_readiness_canvas_window,
        width=event.width
    )


class_readiness_canvas.bind(
    "<Configure>",
    resize_class_readiness_frame
)


class_readiness_canvas.configure(
    yscrollcommand=class_readiness_scrollbar.set
)


class_readiness_canvas.pack(
    side="left",
    fill="both",
    expand=True
)


class_readiness_scrollbar.pack(
    side="right",
    fill="y"
)


def class_readiness_mousewheel(
    event
):

    if notebook.select() == str(
        class_readiness_tab
    ):

        class_readiness_canvas.yview_scroll(
            int(
                -1 * (event.delta / 120)
            ),
            "units"
        )


window.bind_all(
    "<MouseWheel>",
    class_readiness_mousewheel,
    add="+"
)


# -----------------------------------------------------
# HEADER
# -----------------------------------------------------

tk.Label(
    class_readiness_scrollable_frame,
    text="CLASS READINESS CHECKLIST",
    font=("Arial", 20, "bold"),
    bg=lux_pink,
    fg=dark_text
).pack(
    fill="x"
)

tk.Label(
    class_readiness_scrollable_frame,
    text=(
        "One live view of enrollment, money, "
        "certification tracking, and supplies."
    ),
    font=("Arial", 10),
    bg=soft_pink,
    fg=dark_text
).pack(
    pady=(6, 3)
)


class_readiness_session_combo = ttk.Combobox(
    class_readiness_scrollable_frame,
    state="readonly",
    width=90
)


class_readiness_session_combo.pack(
    padx=20,
    pady=(3, 4),
    fill="x"
)


class_readiness_session_combo.bind(
    "<<ComboboxSelected>>",
    refresh_class_readiness
)


class_readiness_title_label = tk.Label(
    class_readiness_scrollable_frame,
    text="Select a scheduled class",
    font=("Arial", 11, "bold"),
    bg=soft_pink,
    fg=dark_text
)


class_readiness_title_label.pack(
    pady=(3, 1)
)


class_readiness_overall_label = tk.Label(
    class_readiness_scrollable_frame,
    text="",
    font=("Arial", 18, "bold"),
    bg=soft_pink,
    fg=gold
)


class_readiness_overall_label.pack(
    pady=(2, 1)
)


class_readiness_overall_detail = tk.Label(
    class_readiness_scrollable_frame,
    text="",
    font=("Arial", 9),
    bg=soft_pink,
    fg=dark_text,
    wraplength=850,
    justify="center"
)


class_readiness_overall_detail.pack(
    pady=(0, 4)
)


class_readiness_updated_label = tk.Label(
    class_readiness_scrollable_frame,
    text="",
    font=("Arial", 8, "italic"),
    bg=soft_pink,
    fg=dark_text
)


class_readiness_updated_label.pack(
    pady=(0, 5)
)


# -----------------------------------------------------
# SUMMARY CARDS
# -----------------------------------------------------

class_readiness_cards = tk.Frame(
    class_readiness_scrollable_frame,
    bg=soft_pink
)


class_readiness_cards.pack(
    fill="x",
    padx=12,
    pady=5
)


for column in range(4):

    class_readiness_cards.columnconfigure(
        column,
        weight=1
    )


def make_class_readiness_card(
    column,
    title
):

    frame = tk.LabelFrame(
        class_readiness_cards,
        text=title,
        font=("Arial", 9, "bold"),
        bg="white",
        fg=dark_text,
        padx=8,
        pady=5
    )


    frame.grid(
        row=0,
        column=column,
        padx=4,
        sticky="nsew"
    )


    value_label = tk.Label(
        frame,
        text="",
        font=("Arial", 15, "bold"),
        bg="white",
        fg=gold
    )


    value_label.pack()


    detail_label = tk.Label(
        frame,
        text="",
        font=("Arial", 8),
        bg="white",
        fg=dark_text,
        wraplength=190,
        justify="center"
    )


    detail_label.pack(
        pady=(2, 0)
    )


    return (
        value_label,
        detail_label
    )


(
    class_readiness_enrollment_value,
    class_readiness_enrollment_detail
) = make_class_readiness_card(
    0,
    "ENROLLMENT"
)


(
    class_readiness_payment_value,
    class_readiness_payment_detail
) = make_class_readiness_card(
    1,
    "PAYMENTS"
)


(
    class_readiness_supply_value,
    class_readiness_supply_detail
) = make_class_readiness_card(
    2,
    "SUPPLIES"
)


(
    class_readiness_cert_value,
    class_readiness_cert_detail
) = make_class_readiness_card(
    3,
    "CERTIFICATION"
)


# -----------------------------------------------------
# QUICK ACTIONS
# -----------------------------------------------------

class_readiness_actions = tk.Frame(
    class_readiness_scrollable_frame,
    bg=soft_pink
)


class_readiness_actions.pack(
    pady=5
)


for column in range(4):

    class_readiness_actions.columnconfigure(
        column,
        weight=1
    )


tk.Button(
    class_readiness_actions,
    text="VIEW CLASS ROSTER",
    font=("Arial", 9, "bold"),
    bg=gold,
    fg="white",
    padx=10,
    pady=6,
    command=class_readiness_open_roster
).grid(
    row=0,
    column=0,
    padx=3
)


tk.Button(
    class_readiness_actions,
    text="ENROLLMENTS / PAYMENTS",
    font=("Arial", 9, "bold"),
    bg=lux_pink,
    fg=dark_text,
    padx=10,
    pady=6,
    command=class_readiness_open_enrollments
).grid(
    row=0,
    column=1,
    padx=3
)


tk.Button(
    class_readiness_actions,
    text="OPEN SUPPLIES",
    font=("Arial", 9, "bold"),
    padx=10,
    pady=6,
    command=class_readiness_open_supplies
).grid(
    row=0,
    column=2,
    padx=3
)


tk.Button(
    class_readiness_actions,
    text="REFRESH",
    font=("Arial", 9, "bold"),
    padx=10,
    pady=6,
    command=refresh_class_readiness_all
).grid(
    row=0,
    column=3,
    padx=3
)


# -----------------------------------------------------
# MASTER CHECKLIST
# -----------------------------------------------------

tk.Label(
    class_readiness_scrollable_frame,
    text="MASTER CHECKLIST",
    font=("Arial", 12, "bold"),
    bg=lux_pink,
    fg=dark_text
).pack(
    fill="x",
    pady=(7, 0)
)


class_readiness_checklist_frame = tk.Frame(
    class_readiness_scrollable_frame,
    bg=soft_pink
)


class_readiness_checklist_frame.pack(
    fill="x",
    padx=12,
    pady=(4, 8)
)


class_readiness_checklist_columns = (
    "area",
    "status",
    "detail"
)


class_readiness_checklist_table = ttk.Treeview(
    class_readiness_checklist_frame,
    columns=class_readiness_checklist_columns,
    show="headings",
    height=5
)


class_readiness_checklist_table.heading(
    "area",
    text="Area"
)


class_readiness_checklist_table.heading(
    "status",
    text="Status"
)


class_readiness_checklist_table.heading(
    "detail",
    text="What Lux Found"
)


class_readiness_checklist_table.column(
    "area",
    width=120
)


class_readiness_checklist_table.column(
    "status",
    width=150,
    anchor="center"
)


class_readiness_checklist_table.column(
    "detail",
    width=620
)


class_readiness_checklist_table.tag_configure(
    "ready",
    background="#ECF7EE"
)


class_readiness_checklist_table.tag_configure(
    "warning",
    background="#FFF6D9"
)


class_readiness_checklist_table.tag_configure(
    "critical",
    background="#FCE8E8"
)


class_readiness_checklist_table.tag_configure(
    "info",
    background="#EEF3FA"
)


class_readiness_checklist_table.pack(
    fill="x"
)


# -----------------------------------------------------
# STUDENT MONEY / STATUS TABLE
# -----------------------------------------------------

tk.Label(
    class_readiness_scrollable_frame,
    text="STUDENT READINESS",
    font=("Arial", 12, "bold"),
    bg=lux_pink,
    fg=dark_text
).pack(
    fill="x",
    pady=(4, 0)
)


class_readiness_students_frame = tk.Frame(
    class_readiness_scrollable_frame,
    bg=soft_pink
)


class_readiness_students_frame.pack(
    fill="x",
    padx=12,
    pady=(4, 8)
)


class_readiness_student_columns = (
    "id",
    "student",
    "tuition",
    "paid",
    "balance",
    "payment",
    "certification",
    "status"
)


class_readiness_students_table = ttk.Treeview(
    class_readiness_students_frame,
    columns=class_readiness_student_columns,
    show="headings",
    height=7
)


for column_name, heading_text in (
    ("id", "ID"),
    ("student", "Student"),
    ("tuition", "Tuition"),
    ("paid", "Paid"),
    ("balance", "Balance"),
    ("payment", "Payment"),
    ("certification", "Certification"),
    ("status", "Status")
):

    class_readiness_students_table.heading(
        column_name,
        text=heading_text
    )


class_readiness_students_table.column(
    "id",
    width=45,
    anchor="center"
)


class_readiness_students_table.column(
    "student",
    width=150
)


for column_name in (
    "tuition",
    "paid",
    "balance"
):

    class_readiness_students_table.column(
        column_name,
        width=85,
        anchor="e"
    )


class_readiness_students_table.column(
    "payment",
    width=105,
    anchor="center"
)


class_readiness_students_table.column(
    "certification",
    width=105,
    anchor="center"
)


class_readiness_students_table.column(
    "status",
    width=100,
    anchor="center"
)


class_readiness_students_table.tag_configure(
    "ready",
    background="#ECF7EE"
)


class_readiness_students_table.tag_configure(
    "warning",
    background="#FFF6D9"
)


class_readiness_students_vertical = ttk.Scrollbar(
    class_readiness_students_frame,
    orient="vertical",
    command=class_readiness_students_table.yview
)


class_readiness_students_horizontal = ttk.Scrollbar(
    class_readiness_students_frame,
    orient="horizontal",
    command=class_readiness_students_table.xview
)


class_readiness_students_table.configure(
    yscrollcommand=class_readiness_students_vertical.set,
    xscrollcommand=class_readiness_students_horizontal.set
)


class_readiness_students_table.grid(
    row=0,
    column=0,
    sticky="nsew"
)


class_readiness_students_vertical.grid(
    row=0,
    column=1,
    sticky="ns"
)


class_readiness_students_horizontal.grid(
    row=1,
    column=0,
    sticky="ew"
)


class_readiness_students_frame.columnconfigure(
    0,
    weight=1
)


# -----------------------------------------------------
# SUPPLY READINESS TABLE
# -----------------------------------------------------

tk.Label(
    class_readiness_scrollable_frame,
    text="SUPPLY READINESS",
    font=("Arial", 12, "bold"),
    bg=lux_pink,
    fg=dark_text
).pack(
    fill="x",
    pady=(4, 0)
)


class_readiness_supplies_frame = tk.Frame(
    class_readiness_scrollable_frame,
    bg=soft_pink
)


class_readiness_supplies_frame.pack(
    fill="x",
    padx=12,
    pady=(4, 18)
)


class_readiness_supply_columns = (
    "supply",
    "category",
    "need",
    "have",
    "short",
    "status"
)


class_readiness_supplies_table = ttk.Treeview(
    class_readiness_supplies_frame,
    columns=class_readiness_supply_columns,
    show="headings",
    height=7
)


for column_name, heading_text in (
    ("supply", "Supply"),
    ("category", "Category"),
    ("need", "Need"),
    ("have", "Have"),
    ("short", "Short"),
    ("status", "Status")
):

    class_readiness_supplies_table.heading(
        column_name,
        text=heading_text
    )


class_readiness_supplies_table.column(
    "supply",
    width=190
)


class_readiness_supplies_table.column(
    "category",
    width=120
)


for column_name in (
    "need",
    "have",
    "short"
):

    class_readiness_supplies_table.column(
        column_name,
        width=80,
        anchor="center"
    )


class_readiness_supplies_table.column(
    "status",
    width=150,
    anchor="center"
)


class_readiness_supplies_table.tag_configure(
    "ready",
    background="#ECF7EE"
)


class_readiness_supplies_table.tag_configure(
    "warning",
    background="#FFF6D9"
)


class_readiness_supplies_table.tag_configure(
    "critical",
    background="#FCE8E8"
)


class_readiness_supplies_vertical = ttk.Scrollbar(
    class_readiness_supplies_frame,
    orient="vertical",
    command=class_readiness_supplies_table.yview
)


class_readiness_supplies_horizontal = ttk.Scrollbar(
    class_readiness_supplies_frame,
    orient="horizontal",
    command=class_readiness_supplies_table.xview
)


class_readiness_supplies_table.configure(
    yscrollcommand=class_readiness_supplies_vertical.set,
    xscrollcommand=class_readiness_supplies_horizontal.set
)


class_readiness_supplies_table.grid(
    row=0,
    column=0,
    sticky="nsew"
)


class_readiness_supplies_vertical.grid(
    row=0,
    column=1,
    sticky="ns"
)


class_readiness_supplies_horizontal.grid(
    row=1,
    column=0,
    sticky="ew"
)


class_readiness_supplies_frame.columnconfigure(
    0,
    weight=1
)


# =====================================================
# CLASSES TAB
# =====================================================

classes_header = tk.Label(
    classes_tab,
    text="CLASS SESSIONS",
    font=("Arial", 20, "bold"),
    bg=lux_pink,
    fg=dark_text
)


classes_header.pack(
    fill="x",
    pady=(0, 5)
)


class_session_count_label = tk.Label(
    classes_tab,
    text="0 scheduled classes",
    font=("Arial", 11),
    bg=soft_pink,
    fg=dark_text
)


class_session_count_label.pack(
    pady=(5, 3)
)


tk.Label(
    classes_tab,
    text=(
        "Select a class to see its live roster below. "
        "Planned = calculator estimate • "
        "Enrolled = actual student records."
    ),
    font=("Arial", 10),
    bg=soft_pink,
    fg=dark_text
).pack(
    pady=(0, 5)
)


# -----------------------------------------------------
# CLASS SEARCH + FILTERS
# -----------------------------------------------------

class_filter_frame = tk.Frame(
    classes_tab,
    bg=soft_pink
)


class_filter_frame.pack(
    fill="x",
    padx=15,
    pady=(3, 4)
)


# Free-text search.
# Typing refreshes the table immediately.

class_search_var = tk.StringVar()


tk.Label(
    class_filter_frame,
    text="Search",
    font=("Arial", 9, "bold"),
    bg=soft_pink
).grid(
    row=0,
    column=0,
    sticky="w",
    padx=(0, 4)
)


class_search_entry = tk.Entry(
    class_filter_frame,
    textvariable=class_search_var,
    width=18
)


class_search_entry.grid(
    row=0,
    column=1,
    padx=(0, 8)
)


class_search_entry.bind(
    "<KeyRelease>",
    refresh_class_sessions_table
)


# Class type filter.

class_type_filter_var = tk.StringVar(
    value="All Classes"
)


tk.Label(
    class_filter_frame,
    text="Class",
    font=("Arial", 9, "bold"),
    bg=soft_pink
).grid(
    row=0,
    column=2,
    sticky="w",
    padx=(0, 4)
)


class_type_filter_combo = ttk.Combobox(
    class_filter_frame,
    textvariable=class_type_filter_var,
    values=(
        "All Classes",
        *class_prices.keys()
    ),
    state="readonly",
    width=14
)


class_type_filter_combo.grid(
    row=0,
    column=3,
    padx=(0, 8)
)


class_type_filter_combo.bind(
    "<<ComboboxSelected>>",
    refresh_class_sessions_table
)


# Date-window filter.

class_date_filter_var = tk.StringVar(
    value="All Dates"
)


tk.Label(
    class_filter_frame,
    text="When",
    font=("Arial", 9, "bold"),
    bg=soft_pink
).grid(
    row=0,
    column=4,
    sticky="w",
    padx=(0, 4)
)


class_date_filter_combo = ttk.Combobox(
    class_filter_frame,
    textvariable=class_date_filter_var,
    values=(
        "All Dates",
        "Upcoming",
        "Today",
        "Past"
    ),
    state="readonly",
    width=11
)


class_date_filter_combo.grid(
    row=0,
    column=5,
    padx=(0, 8)
)


class_date_filter_combo.bind(
    "<<ComboboxSelected>>",
    refresh_class_sessions_table
)


# Seat availability filter.

class_seats_filter_var = tk.StringVar(
    value="All Seats"
)


tk.Label(
    class_filter_frame,
    text="Seats",
    font=("Arial", 9, "bold"),
    bg=soft_pink
).grid(
    row=0,
    column=6,
    sticky="w",
    padx=(0, 4)
)


class_seats_filter_combo = ttk.Combobox(
    class_filter_frame,
    textvariable=class_seats_filter_var,
    values=(
        "All Seats",
        "Open Seats",
        "Full",
        "Over Capacity"
    ),
    state="readonly",
    width=13
)


class_seats_filter_combo.grid(
    row=0,
    column=7,
    padx=(0, 8)
)


class_seats_filter_combo.bind(
    "<<ComboboxSelected>>",
    refresh_class_sessions_table
)


# Status filter.

class_status_filter_var = tk.StringVar(
    value="All Statuses"
)


tk.Label(
    class_filter_frame,
    text="Status",
    font=("Arial", 9, "bold"),
    bg=soft_pink
).grid(
    row=1,
    column=0,
    sticky="w",
    padx=(0, 4),
    pady=(4, 0)
)

class_status_filter_combo = ttk.Combobox(
    class_filter_frame,
    textvariable=class_status_filter_var,
    values=(
        "All Statuses",
        "Scheduled",
        "Completed",
        "Cancelled"
    ),
    state="readonly",
    width=14
)


class_status_filter_combo.grid(
    row=1,
    column=1,
    padx=(0, 8),
    pady=(4, 0)
)


clear_class_filters_button = tk.Button(
    class_filter_frame,
    text="CLEAR FILTERS",
    font=("Arial", 9, "bold"),
    padx=10,
    pady=3,
    command=clear_class_filters
)


clear_class_filters_button.grid(
    row=1,
    column=2,
    columnspan=2,
    sticky="w",
    pady=(4, 0)
)


class_filter_hint = tk.Label(
    class_filter_frame,
    text=(
        "Search class name, database date/time, or status. "
        "Filters can be combined."
    ),
    font=("Arial", 8, "italic"),
    bg=soft_pink,
    fg=dark_text
)


class_filter_hint.grid(
    row=1,
    column=4,
    columnspan=4,
    sticky="w",
    pady=(4, 0)
)


# -----------------------------------------------------
# CLASS SESSION TABLE
# -----------------------------------------------------

class_sessions_frame = tk.Frame(
    classes_tab,
    bg=soft_pink
)


class_sessions_frame.pack(
    fill="both",
    expand=True,
    padx=15,
    pady=(5, 4)
)


class_session_columns = (
    "id",
    "date",
    "time",
    "class",
    "planned",
    "enrolled",
    "capacity",
    "tuition",
    "status"
)


class_sessions_table = ttk.Treeview(
    class_sessions_frame,
    columns=class_session_columns,
    show="headings",
    height=6
)


class_sessions_table.heading(
    "id",
    text="ID"
)


class_sessions_table.heading(
    "date",
    text="Date"
)


class_sessions_table.heading(
    "time",
    text="Time"
)


class_sessions_table.heading(
    "class",
    text="Class"
)


class_sessions_table.heading(
    "planned",
    text="Planned"
)


class_sessions_table.heading(
    "enrolled",
    text="Enrolled"
)


class_sessions_table.heading(
    "capacity",
    text="Capacity"
)


class_sessions_table.heading(
    "tuition",
    text="Tuition"
)


class_sessions_table.heading(
    "status",
    text="Status"
)


class_sessions_table.column(
    "id",
    width=40,
    anchor="center"
)


class_sessions_table.column(
    "date",
    width=95,
    anchor="center"
)


class_sessions_table.column(
    "time",
    width=85,
    anchor="center"
)


class_sessions_table.column(
    "class",
    width=120
)


class_sessions_table.column(
    "planned",
    width=65,
    anchor="center"
)


class_sessions_table.column(
    "enrolled",
    width=65,
    anchor="center"
)


class_sessions_table.column(
    "capacity",
    width=65,
    anchor="center"
)


class_sessions_table.column(
    "tuition",
    width=85,
    anchor="e"
)


class_sessions_table.column(
    "status",
    width=85,
    anchor="center"
)


class_session_vertical_scrollbar = ttk.Scrollbar(
    class_sessions_frame,
    orient="vertical",
    command=class_sessions_table.yview
)


class_session_horizontal_scrollbar = ttk.Scrollbar(
    class_sessions_frame,
    orient="horizontal",
    command=class_sessions_table.xview
)


class_sessions_table.configure(
    yscrollcommand=class_session_vertical_scrollbar.set,
    xscrollcommand=class_session_horizontal_scrollbar.set
)


class_sessions_table.grid(
    row=0,
    column=0,
    sticky="nsew"
)


class_session_vertical_scrollbar.grid(
    row=0,
    column=1,
    sticky="ns"
)


class_session_horizontal_scrollbar.grid(
    row=1,
    column=0,
    sticky="ew"
)


class_sessions_frame.rowconfigure(
    0,
    weight=1
)


class_sessions_frame.columnconfigure(
    0,
    weight=1
)


# Clicking a class automatically refreshes the roster.
#
# <<TreeviewSelect>> is a Tkinter "virtual event".
# It fires whenever the selected row changes.

class_sessions_table.bind(
    "<<TreeviewSelect>>",
    refresh_selected_class_roster
)


# -----------------------------------------------------
# CLASS SESSION BUTTONS
# -----------------------------------------------------

class_session_buttons = tk.Frame(
    classes_tab,
    bg=soft_pink
)


class_session_buttons.pack(
    pady=(2, 6)
)


view_roster_button = tk.Button(
    class_session_buttons,
    text="VIEW ROSTER",
    font=("Arial", 10, "bold"),
    bg=gold,
    fg="white",
    padx=12,
    pady=6,
    command=lambda: refresh_selected_class_roster(
        show_warning=True
    )
)


view_roster_button.grid(
    row=0,
    column=0,
    padx=4
)


load_session_button = tk.Button(
    class_session_buttons,
    text="LOAD TO CALCULATOR",
    font=("Arial", 10, "bold"),
    bg=lux_pink,
    fg=dark_text,
    padx=12,
    pady=6,
    command=load_selected_session
)


load_session_button.grid(
    row=0,
    column=1,
    padx=4
)


delete_session_button = tk.Button(
    class_session_buttons,
    text="DELETE SELECTED",
    font=("Arial", 10, "bold"),
    padx=12,
    pady=6,
    command=delete_selected_session
)


delete_session_button.grid(
    row=0,
    column=2,
    padx=4
)


refresh_sessions_button = tk.Button(
    class_session_buttons,
    text="REFRESH RESULTS",
    font=("Arial", 10, "bold"),
    bg=lux_pink,
    fg=dark_text,
    padx=12,
    pady=6,
    command=refresh_class_sessions_table
)


refresh_sessions_button.grid(
    row=0,
    column=3,
    padx=4
)


class_readiness_button = tk.Button(
    class_session_buttons,
    text="READINESS",
    font=("Arial", 10, "bold"),
    bg=gold,
    fg="white",
    padx=12,
    pady=6,
    command=open_selected_class_readiness
)


class_readiness_button.grid(
    row=0,
    column=4,
    padx=4
)


# =====================================================
# LIVE CLASS ROSTER
# =====================================================

roster_header = tk.Label(
    classes_tab,
    text="LIVE CLASS ROSTER",
    font=("Arial", 13, "bold"),
    bg=lux_pink,
    fg=dark_text
)


roster_header.pack(
    fill="x",
    pady=(2, 3)
)


class_roster_title_label = tk.Label(
    classes_tab,
    text="Select a class above to view its roster",
    font=("Arial", 10, "bold"),
    bg=soft_pink,
    fg=dark_text
)


class_roster_title_label.pack(
    pady=(2, 1)
)


class_roster_capacity_label = tk.Label(
    classes_tab,
    text="",
    font=("Arial", 10),
    bg=soft_pink,
    fg=dark_text
)


class_roster_capacity_label.pack(
    pady=1
)


class_roster_money_label = tk.Label(
    classes_tab,
    text="",
    font=("Arial", 10, "bold"),
    bg=soft_pink,
    fg=gold
)


class_roster_money_label.pack(
    pady=1
)


class_roster_payment_label = tk.Label(
    classes_tab,
    text="",
    font=("Arial", 9),
    bg=soft_pink,
    fg=dark_text
)


class_roster_payment_label.pack(
    pady=(1, 2)
)


# -----------------------------------------------------
# ROSTER TABLE
# -----------------------------------------------------

class_roster_frame = tk.Frame(
    classes_tab,
    bg=soft_pink
)


class_roster_frame.pack(
    fill="both",
    expand=True,
    padx=15,
    pady=(2, 10)
)


class_roster_columns = (
    "student_id",
    "name",
    "payment",
    "paid",
    "balance",
    "certification",
    "phone",
    "email"
)


class_roster_table = ttk.Treeview(
    class_roster_frame,
    columns=class_roster_columns,
    show="headings",
    height=6
)


class_roster_table.heading(
    "student_id",
    text="Student ID"
)


class_roster_table.heading(
    "name",
    text="Student"
)


class_roster_table.heading(
    "payment",
    text="Payment Status"
)


class_roster_table.heading(
    "paid",
    text="Paid"
)


class_roster_table.heading(
    "balance",
    text="Balance"
)


class_roster_table.heading(
    "certification",
    text="Certification"
)


class_roster_table.heading(
    "phone",
    text="Phone"
)


class_roster_table.heading(
    "email",
    text="Email"
)


class_roster_table.column(
    "student_id",
    width=70,
    anchor="center"
)


class_roster_table.column(
    "name",
    width=145
)


class_roster_table.column(
    "payment",
    width=105,
    anchor="center"
)


class_roster_table.column(
    "paid",
    width=85,
    anchor="e"
)


class_roster_table.column(
    "balance",
    width=85,
    anchor="e"
)


class_roster_table.column(
    "certification",
    width=110,
    anchor="center"
)


class_roster_table.column(
    "phone",
    width=115
)


class_roster_table.column(
    "email",
    width=180
)


class_roster_vertical_scrollbar = ttk.Scrollbar(
    class_roster_frame,
    orient="vertical",
    command=class_roster_table.yview
)


class_roster_horizontal_scrollbar = ttk.Scrollbar(
    class_roster_frame,
    orient="horizontal",
    command=class_roster_table.xview
)


class_roster_table.configure(
    yscrollcommand=class_roster_vertical_scrollbar.set,
    xscrollcommand=class_roster_horizontal_scrollbar.set
)


class_roster_table.grid(
    row=0,
    column=0,
    sticky="nsew"
)


class_roster_vertical_scrollbar.grid(
    row=0,
    column=1,
    sticky="ns"
)


class_roster_horizontal_scrollbar.grid(
    row=1,
    column=0,
    sticky="ew"
)


class_roster_frame.rowconfigure(
    0,
    weight=1
)


class_roster_frame.columnconfigure(
    0,
    weight=1
)


# =====================================================
# HISTORY TAB
# =====================================================

history_header = tk.Label(
    history_tab,
    text="CLASS HISTORY",
    font=("Arial", 20, "bold"),
    bg=lux_pink,
    fg=dark_text
)


history_header.pack(
    fill="x",
    pady=(0, 8)
)


history_status_label = tk.Label(
    history_tab,
    text="",
    font=("Arial", 11),
    bg=soft_pink,
    fg=dark_text
)


history_status_label.pack(
    pady=(5, 5)
)


history_table_frame = tk.Frame(
    history_tab,
    bg=soft_pink
)


history_table_frame.pack(
    fill="both",
    expand=True,
    padx=15,
    pady=10
)


history_columns = (
    "date",
    "class",
    "students",
    "revenue",
    "profit",
    "margin"
)


history_table = ttk.Treeview(
    history_table_frame,
    columns=history_columns,
    show="headings"
)


history_table.heading(
    "date",
    text="Saved Date"
)


history_table.heading(
    "class",
    text="Class Type"
)


history_table.heading(
    "students",
    text="Students"
)


history_table.heading(
    "revenue",
    text="Revenue"
)


history_table.heading(
    "profit",
    text="Estimated Profit"
)


history_table.heading(
    "margin",
    text="Profit Margin"
)


history_table.column(
    "date",
    width=140
)


history_table.column(
    "class",
    width=130
)


history_table.column(
    "students",
    width=75,
    anchor="center"
)


history_table.column(
    "revenue",
    width=110,
    anchor="e"
)


history_table.column(
    "profit",
    width=120,
    anchor="e"
)


history_table.column(
    "margin",
    width=90,
    anchor="center"
)


history_vertical_scrollbar = ttk.Scrollbar(
    history_table_frame,
    orient="vertical",
    command=history_table.yview
)


history_horizontal_scrollbar = ttk.Scrollbar(
    history_table_frame,
    orient="horizontal",
    command=history_table.xview
)


history_table.configure(
    yscrollcommand=history_vertical_scrollbar.set,
    xscrollcommand=history_horizontal_scrollbar.set
)


history_table.grid(
    row=0,
    column=0,
    sticky="nsew"
)


history_vertical_scrollbar.grid(
    row=0,
    column=1,
    sticky="ns"
)


history_horizontal_scrollbar.grid(
    row=1,
    column=0,
    sticky="ew"
)


history_table_frame.rowconfigure(
    0,
    weight=1
)


history_table_frame.columnconfigure(
    0,
    weight=1
)


history_buttons = tk.Frame(
    history_tab,
    bg=soft_pink
)


history_buttons.pack(
    pady=(0, 15)
)


load_history_button = tk.Button(
    history_buttons,
    text="LOAD SELECTED",
    font=("Arial", 11, "bold"),
    bg=lux_pink,
    fg=dark_text,
    padx=15,
    pady=8,
    command=load_selected_plan
)


load_history_button.grid(
    row=0,
    column=0,
    padx=5
)


delete_history_button = tk.Button(
    history_buttons,
    text="DELETE SELECTED",
    font=("Arial", 11, "bold"),
    padx=15,
    pady=8,
    command=delete_selected_plan
)


delete_history_button.grid(
    row=0,
    column=1,
    padx=5
)


refresh_history_button = tk.Button(
    history_buttons,
    text="REFRESH HISTORY",
    font=("Arial", 11, "bold"),
    bg=gold,
    fg="white",
    padx=15,
    pady=8,
    command=refresh_history_table
)


refresh_history_button.grid(
    row=0,
    column=2,
    padx=5
)


# =====================================================
# STUDENTS TAB
# =====================================================

student_header = tk.Label(
    students_tab,
    text="STUDENT MANAGEMENT",
    font=("Arial", 20, "bold"),
    bg=lux_pink,
    fg=dark_text
)


student_header.pack(
    fill="x",
    pady=(0, 5)
)


student_count_label = tk.Label(
    students_tab,
    text="0 students",
    font=("Arial", 11),
    bg=soft_pink,
    fg=dark_text
)


student_count_label.pack(
    pady=(4, 1)
)


student_edit_status_label = tk.Label(
    students_tab,
    text="Adding a new student",
    font=("Arial", 10, "italic"),
    bg=soft_pink,
    fg=gold
)


student_edit_status_label.pack(
    pady=(0, 5)
)


# -----------------------------------------------------
# STUDENT SEARCH + FILTERS
# -----------------------------------------------------

student_filter_frame = tk.Frame(
    students_tab,
    bg=soft_pink
)


student_filter_frame.pack(
    fill="x",
    padx=15,
    pady=(3, 5)
)


# One search box covers name, phone, and email.

student_search_var = tk.StringVar()


tk.Label(
    student_filter_frame,
    text="Search",
    font=("Arial", 9, "bold"),
    bg=soft_pink
).grid(
    row=0,
    column=0,
    sticky="w",
    padx=(0, 4)
)


student_search_entry = tk.Entry(    student_filter_frame,
    textvariable=student_search_var,
    width=22
)


student_search_entry.grid(
    row=0,
    column=1,
    padx=(0, 10)
)


student_search_entry.bind(
    "<KeyRelease>",
    refresh_students_table
)


# These filters use the relational enrollment/payment data.
# A student can still appear only once even if several enrollments match.

student_class_filter_var = tk.StringVar(
    value="All Classes"
)


student_class_filter_combo = ttk.Combobox(
    student_filter_frame,
    textvariable=student_class_filter_var,
    values=(
        "All Classes",
        *class_prices.keys()
    ),
    state="readonly",
    width=15
)


student_class_filter_combo.grid(
    row=0,
    column=2,
    padx=(0, 8)
)


student_class_filter_combo.bind(
    "<<ComboboxSelected>>",
    refresh_students_table
)


student_payment_filter_var = tk.StringVar(
    value="All Payments"
)


student_payment_filter_combo = ttk.Combobox(
    student_filter_frame,
    textvariable=student_payment_filter_var,
    values=(
        "All Payments",
        "Not Paid",
        "Deposit Paid",
        "Paid in Full"
    ),
    state="readonly",
    width=14
)


student_payment_filter_combo.grid(
    row=0,
    column=3,
    padx=(0, 8)
)


student_payment_filter_combo.bind(
    "<<ComboboxSelected>>",
    refresh_students_table
)


student_cert_filter_var = tk.StringVar(
    value="All Certifications"
)


student_cert_filter_combo = ttk.Combobox(
    student_filter_frame,
    textvariable=student_cert_filter_var,
    values=(
        "All Certifications",
        "Not Started",
        "In Progress",
        "Certified"
    ),
    state="readonly",
    width=16
)


student_cert_filter_combo.grid(
    row=0,
    column=4,
    padx=(0, 8)
)


student_cert_filter_combo.bind(
    "<<ComboboxSelected>>",
    refresh_students_table
)


student_balance_filter_var = tk.StringVar(
    value="All Balances"
)


student_balance_filter_combo = ttk.Combobox(
    student_filter_frame,
    textvariable=student_balance_filter_var,
    values=(
        "All Balances",
        "Balance Due",
        "No Balance"
    ),
    state="readonly",
    width=13
)


student_balance_filter_combo.grid(
    row=0,
    column=5,
    padx=(0, 8)
)


student_balance_filter_combo.bind(
    "<<ComboboxSelected>>",
    refresh_students_table
)


clear_student_filters_button = tk.Button(
    student_filter_frame,
    text="CLEAR FILTERS",
    font=("Arial", 9, "bold"),
    padx=10,
    pady=3,
    command=clear_student_filters
)


clear_student_filters_button.grid(
    row=0,
    column=6,
    padx=(2, 0)
)


student_filter_hint = tk.Label(
    student_filter_frame,
    text=(
        "Class, payment, certification, and balance filters use "
        "the student's actual enrollment records."
    ),
    font=("Arial", 8, "italic"),
    bg=soft_pink,
    fg=dark_text
)


student_filter_hint.grid(
    row=1,
    column=0,
    columnspan=7,
    sticky="w",
    pady=(4, 0)
)


# -----------------------------------------------------
# STUDENT FORM
# -----------------------------------------------------

student_form_frame = tk.Frame(
    students_tab,
    bg=soft_pink
)


student_form_frame.pack(
    pady=5
)


# FIRST NAME

tk.Label(
    student_form_frame,
    text="First Name",
    font=("Arial", 10, "bold"),
    bg=soft_pink
).grid(
    row=0,
    column=0,
    sticky="w",
    padx=5,
    pady=4
)


student_first_name_entry = tk.Entry(
    student_form_frame,
    width=20
)


student_first_name_entry.grid(
    row=0,
    column=1,
    padx=5,
    pady=4
)


# LAST NAME

tk.Label(
    student_form_frame,
    text="Last Name",
    font=("Arial", 10, "bold"),
    bg=soft_pink
).grid(
    row=0,
    column=2,
    sticky="w",
    padx=5,
    pady=4
)


student_last_name_entry = tk.Entry(
    student_form_frame,
    width=20
)


student_last_name_entry.grid(
    row=0,
    column=3,
    padx=5,
    pady=4
)


# PHONE

tk.Label(
    student_form_frame,
    text="Phone",
    font=("Arial", 10, "bold"),
    bg=soft_pink
).grid(
    row=1,
    column=0,
    sticky="w",
    padx=5,
    pady=4
)


student_phone_entry = tk.Entry(
    student_form_frame,
    width=20
)


student_phone_entry.grid(
    row=1,
    column=1,
    padx=5,
    pady=4
)


# EMAIL

tk.Label(
    student_form_frame,
    text="Email",
    font=("Arial", 10, "bold"),
    bg=soft_pink
).grid(
    row=1,
    column=2,
    sticky="w",
    padx=5,
    pady=4
)


student_email_entry = tk.Entry(
    student_form_frame,
    width=20
)


student_email_entry.grid(
    row=1,
    column=3,
    padx=5,
    pady=4
)


# CLASS

tk.Label(
    student_form_frame,
    text="Class",
    font=("Arial", 10, "bold"),
    bg=soft_pink
).grid(
    row=2,
    column=0,
    sticky="w",
    padx=5,
    pady=4
)


student_class_combo = ttk.Combobox(
    student_form_frame,
    values=list(
        class_prices.keys()
    ),
    state="readonly",
    width=18
)


student_class_combo.set(
    "Intro Class"
)


student_class_combo.grid(
    row=2,
    column=1,
    padx=5,
    pady=4
)


# PAYMENT

tk.Label(
    student_form_frame,
    text="Payment",
    font=("Arial", 10, "bold"),
    bg=soft_pink
).grid(
    row=2,
    column=2,
    sticky="w",
    padx=5,
    pady=4
)


payment_status_combo = ttk.Combobox(
    student_form_frame,
    values=(
        "Not Paid",
        "Deposit Paid",
        "Paid in Full"
    ),
    state="readonly",
    width=18
)


payment_status_combo.set(
    "Not Paid"
)


payment_status_combo.grid(
    row=2,
    column=3,
    padx=5,
    pady=4
)


# CERTIFICATION

tk.Label(
    student_form_frame,
    text="Certification",
    font=("Arial", 10, "bold"),
    bg=soft_pink
).grid(
    row=3,
    column=0,
    sticky="w",
    padx=5,
    pady=4
)


certification_status_combo = ttk.Combobox(
    student_form_frame,
    values=(
        "Not Started",
        "In Progress",
        "Certified"
    ),
    state="readonly",
    width=18
)


certification_status_combo.set(
    "Not Started"
)


certification_status_combo.grid(
    row=3,
    column=1,
    padx=5,
    pady=4
)


# NOTES

tk.Label(
    student_form_frame,
    text="Notes",
    font=("Arial", 10, "bold"),
    bg=soft_pink
).grid(
    row=3,
    column=2,
    sticky="nw",
    padx=5,
    pady=4
)


# Entry = one line
#
# Text = multiple lines

student_notes_text = tk.Text(
    student_form_frame,
    width=22,
    height=3
)


student_notes_text.grid(
    row=3,
    column=3,
    padx=5,
    pady=4
)


# -----------------------------------------------------
# STUDENT FORM BUTTONS
# -----------------------------------------------------

student_form_buttons = tk.Frame(
    students_tab,
    bg=soft_pink
)


student_form_buttons.pack(
    pady=6
)


add_student_button = tk.Button(
    student_form_buttons,
    text="ADD STUDENT",
    font=("Arial", 11, "bold"),
    bg=gold,
    fg="white",
    padx=18,
    pady=8,
    command=add_student
)


add_student_button.grid(
    row=0,
    column=0,
    padx=5
)


update_student_button = tk.Button(
    student_form_buttons,
    text="UPDATE LOADED",
    font=("Arial", 11, "bold"),
    bg=lux_pink,
    fg=dark_text,
    padx=18,
    pady=8,
    command=update_loaded_student
)


update_student_button.grid(
    row=0,
    column=1,
    padx=5
)


clear_student_button = tk.Button(
    student_form_buttons,
    text="CLEAR FORM",
    font=("Arial", 11, "bold"),
    padx=18,
    pady=8,
    command=clear_student_form
)


clear_student_button.grid(
    row=0,
    column=2,
    padx=5
)


# -----------------------------------------------------
# STUDENT TABLE
# -----------------------------------------------------

student_table_frame = tk.Frame(
    students_tab,
    bg=soft_pink
)


student_table_frame.pack(
    fill="both",
    expand=True,
    padx=15,
    pady=(5, 5)
)


student_columns = (
    "id",
    "name",
    "enrollments",
    "outstanding",
    "phone",
    "email"
)


students_table = ttk.Treeview(
    student_table_frame,
    columns=student_columns,
    show="headings"
)


students_table.heading(
    "id",
    text="ID"
)


students_table.heading(
    "name",
    text="Student"
)


students_table.heading(
    "enrollments",
    text="Enrollments"
)


students_table.heading(
    "outstanding",
    text="Outstanding"
)


students_table.heading(
    "phone",
    text="Phone"
)


students_table.heading(
    "email",
    text="Email"
)


students_table.column(
    "id",
    width=45,
    anchor="center"
)


students_table.column(
    "name",
    width=180
)


students_table.column(
    "enrollments",
    width=90,
    anchor="center"
)


students_table.column(
    "outstanding",
    width=110,
    anchor="e"
)


students_table.column(
    "phone",
    width=125
)


students_table.column(
    "email",
    width=200
)


student_vertical_scrollbar = ttk.Scrollbar(
    student_table_frame,
    orient="vertical",
    command=students_table.yview
)


student_horizontal_scrollbar = ttk.Scrollbar(
    student_table_frame,
    orient="horizontal",
    command=students_table.xview
)


students_table.configure(
    yscrollcommand=student_vertical_scrollbar.set,
    xscrollcommand=student_horizontal_scrollbar.set
)


students_table.grid(
    row=0,
    column=0,
    sticky="nsew"
)


student_vertical_scrollbar.grid(
    row=0,
    column=1,
    sticky="ns"
)


student_horizontal_scrollbar.grid(
    row=1,
    column=0,
    sticky="ew"
)


student_table_frame.rowconfigure(
    0,
    weight=1
)


student_table_frame.columnconfigure(
    0,
    weight=1
)


# -----------------------------------------------------
# STUDENT TABLE CRUD BUTTONS
# -----------------------------------------------------

student_table_buttons = tk.Frame(
    students_tab,
    bg=soft_pink
)


student_table_buttons.pack(
    pady=(0, 15)
)


load_student_button = tk.Button(
    student_table_buttons,
    text="LOAD SELECTED",
    font=("Arial", 11, "bold"),
    bg=lux_pink,
    fg=dark_text,
    padx=15,
    pady=8,
    command=load_selected_student
)


load_student_button.grid(
    row=0,
    column=0,
    padx=5
)


delete_student_button = tk.Button(
    student_table_buttons,
    text="DELETE SELECTED",
    font=("Arial", 11, "bold"),
    padx=15,
    pady=8,
    command=delete_selected_student
)


delete_student_button.grid(
    row=0,
    column=1,
    padx=5
)


refresh_student_button = tk.Button(
    student_table_buttons,
    text="REFRESH RESULTS",
    font=("Arial", 11, "bold"),
    bg=gold,
    fg="white",
    padx=15,
    pady=8,
    command=refresh_students_table
)


refresh_student_button.grid(
    row=0,
    column=2,
    padx=5
)



# =====================================================
# ENROLLMENTS TAB
# =====================================================

# -----------------------------------------------------
# SCROLLABLE ENROLLMENTS PAGE
# -----------------------------------------------------

# The Enrollments tab now contains several sections:
# controls, the enrollment table, and the payment ledger.
#
# On shorter screens, the bottom buttons can fall below the
# visible window. A Canvas lets the whole page scroll vertically.

enrollment_main_container = tk.Frame(
    enrollments_tab,
    bg=soft_pink
)


enrollment_main_container.pack(
    fill="both",
    expand=True
)


enrollment_canvas = tk.Canvas(
    enrollment_main_container,
    bg=soft_pink,
    highlightthickness=0
)


enrollment_page_scrollbar = ttk.Scrollbar(
    enrollment_main_container,
    orient="vertical",
    command=enrollment_canvas.yview
)


enrollment_scrollable_frame = tk.Frame(
    enrollment_canvas,
    bg=soft_pink
)


enrollment_scrollable_frame.bind(
    "<Configure>",
    lambda event: enrollment_canvas.configure(
        scrollregion=enrollment_canvas.bbox("all")
    )
)


enrollment_canvas_window = enrollment_canvas.create_window(
    (0, 0),
    window=enrollment_scrollable_frame,
    anchor="nw"
)


def resize_enrollment_scrollable_frame(event):

    # Keep the page the same width as the visible canvas.
    # That way we only need vertical page scrolling.

    enrollment_canvas.itemconfigure(
        enrollment_canvas_window,
        width=event.width
    )


enrollment_canvas.bind(
    "<Configure>",
    resize_enrollment_scrollable_frame
)


enrollment_canvas.configure(
    yscrollcommand=enrollment_page_scrollbar.set
)


enrollment_canvas.pack(
    side="left",
    fill="both",
    expand=True
)


enrollment_page_scrollbar.pack(
    side="right",
    fill="y"
)


def enrollment_mousewheel_scroll(event):

    # Scroll this page only when the Enrollments tab is active.

    if notebook.select() == str(enrollments_tab):

        enrollment_canvas.yview_scroll(
            int(
                -1 * (event.delta / 120)
            ),
            "units"
        )


# add="+" preserves the Calculator tab's existing wheel binding.
window.bind_all(
    "<MouseWheel>",
    enrollment_mousewheel_scroll,
    add="+"
)




enrollment_header = tk.Label(
    enrollment_scrollable_frame,
    text="CLASS ENROLLMENTS",
    font=("Arial", 20, "bold"),
    bg=lux_pink,
    fg=dark_text
)


enrollment_header.pack(
    fill="x",
    pady=(0, 5)
)


enrollment_count_label = tk.Label(
    enrollment_scrollable_frame,
    text="0 enrollments",
    font=("Arial", 11),
    bg=soft_pink,
    fg=dark_text
)


enrollment_count_label.pack(
    pady=(4, 1)
)


enrollment_edit_status_label = tk.Label(
    enrollment_scrollable_frame,
    text="Creating a new enrollment",
    font=("Arial", 10, "italic"),
    bg=soft_pink,
    fg=gold
)


enrollment_edit_status_label.pack(
    pady=(0, 3)
)


tk.Label(
    enrollment_scrollable_frame,
    text=(
        "Choose an existing student and an existing class session. "
        "Payment and certification status now belong to THIS enrollment."
    ),
    font=("Arial", 10),
    bg=soft_pink,
    fg=dark_text
).pack(
    pady=(0, 8)
)


# -----------------------------------------------------
# STUDENT + CLASS RELATIONSHIP CHOICES
# -----------------------------------------------------

enrollment_choice_frame = tk.Frame(
    enrollment_scrollable_frame,
    bg=soft_pink
)


enrollment_choice_frame.pack(
    pady=4,
    padx=10,
    fill="x"
)


tk.Label(
    enrollment_choice_frame,
    text="Student",
    font=("Arial", 10, "bold"),
    bg=soft_pink
).grid(
    row=0,
    column=0,
    sticky="w",
    padx=5,
    pady=(0, 3)
)


tk.Label(
    enrollment_choice_frame,
    text="Scheduled Class Session",
    font=("Arial", 10, "bold"),
    bg=soft_pink
).grid(
    row=0,
    column=1,
    sticky="w",
    padx=5,
    pady=(0, 3)
)

student_enrollment_combo = ttk.Combobox(
    enrollment_choice_frame,
    state="readonly",
    width=32
)


student_enrollment_combo.grid(
    row=1,
    column=0,
    sticky="ew",
    padx=5
)


session_enrollment_combo = ttk.Combobox(
    enrollment_choice_frame,
    state="readonly",
    width=68
)


session_enrollment_combo.grid(
    row=1,
    column=1,
    sticky="ew",
    padx=5
)


enrollment_choice_frame.columnconfigure(
    0,
    weight=1
)


enrollment_choice_frame.columnconfigure(
    1,
    weight=2
)


# -----------------------------------------------------
# ENROLLMENT-SPECIFIC STATUS
# -----------------------------------------------------

enrollment_status_frame = tk.Frame(
    enrollment_scrollable_frame,
    bg=soft_pink
)


enrollment_status_frame.pack(
    pady=6
)


tk.Label(
    enrollment_status_frame,
    text="Payment Status (automatic)",
    font=("Arial", 10, "bold"),
    bg=soft_pink
).grid(
    row=0,
    column=0,
    padx=5,
    pady=3
)


enrollment_payment_combo = ttk.Combobox(
    enrollment_status_frame,
    values=(
        "Not Paid",
        "Deposit Paid",
        "Paid in Full"
    ),
    state="disabled",
    width=18
)


enrollment_payment_combo.set(
    "Not Paid"
)


enrollment_payment_combo.grid(
    row=0,
    column=1,
    padx=5,
    pady=3
)


tk.Label(
    enrollment_status_frame,
    text="Certification",
    font=("Arial", 10, "bold"),
    bg=soft_pink
).grid(
    row=0,
    column=2,
    padx=5,
    pady=3
)


enrollment_certification_combo = ttk.Combobox(
    enrollment_status_frame,
    values=(
        "Not Started",
        "In Progress",
        "Certified"
    ),
    state="readonly",
    width=18
)


enrollment_certification_combo.set(
    "Not Started"
)


enrollment_certification_combo.grid(
    row=0,
    column=3,
    padx=5,
    pady=3
)


# -----------------------------------------------------
# ENROLLMENT BUTTONS
# -----------------------------------------------------

enrollment_buttons = tk.Frame(
    enrollment_scrollable_frame,
    bg=soft_pink
)


enrollment_buttons.pack(
    pady=(2, 6)
)


enroll_student_button = tk.Button(
    enrollment_buttons,
    text="ENROLL STUDENT",
    font=("Arial", 10, "bold"),
    bg=gold,
    fg="white",
    padx=12,
    pady=6,
    command=enroll_student
)


enroll_student_button.grid(
    row=0,
    column=0,
    padx=4,
    pady=3
)


load_enrollment_button = tk.Button(
    enrollment_buttons,
    text="LOAD SELECTED",
    font=("Arial", 10, "bold"),
    bg=lux_pink,
    fg=dark_text,
    padx=12,
    pady=6,
    command=load_selected_enrollment
)


load_enrollment_button.grid(
    row=0,
    column=1,
    padx=4,
    pady=3
)


update_enrollment_button = tk.Button(
    enrollment_buttons,
    text="UPDATE CERTIFICATION",
    font=("Arial", 10, "bold"),
    bg=lux_pink,
    fg=dark_text,
    padx=12,
    pady=6,
    command=update_loaded_enrollment
)


update_enrollment_button.grid(
    row=0,
    column=2,
    padx=4,
    pady=3
)


remove_enrollment_button = tk.Button(
    enrollment_buttons,
    text="REMOVE ENROLLMENT",
    font=("Arial", 10, "bold"),
    padx=12,
    pady=6,
    command=remove_selected_enrollment
)


remove_enrollment_button.grid(
    row=0,
    column=3,
    padx=4,
    pady=3
)


clear_enrollment_button = tk.Button(
    enrollment_buttons,
    text="CLEAR FORM",
    font=("Arial", 10, "bold"),
    padx=12,
    pady=6,
    command=clear_enrollment_form
)


clear_enrollment_button.grid(
    row=1,
    column=0,
    padx=4,
    pady=3
)


refresh_enrollment_button = tk.Button(
    enrollment_buttons,
    text="REFRESH",
    font=("Arial", 10, "bold"),
    bg=gold,
    fg="white",
    padx=12,
    pady=6,
    command=refresh_enrollments_table
)


refresh_enrollment_button.grid(
    row=1,
    column=1,
    padx=4,
    pady=3
)


go_students_button = tk.Button(
    enrollment_buttons,
    text="GO TO STUDENTS",
    font=("Arial", 10, "bold"),
    padx=12,
    pady=6,
    command=lambda: notebook.select(
        students_tab
    )
)


go_students_button.grid(
    row=1,
    column=2,
    padx=4,
    pady=3
)


go_classes_button = tk.Button(
    enrollment_buttons,
    text="GO TO CLASSES",
    font=("Arial", 10, "bold"),
    padx=12,
    pady=6,
    command=lambda: notebook.select(
        classes_tab
    )
)


go_classes_button.grid(
    row=1,
    column=3,
    padx=4,
    pady=3
)


# -----------------------------------------------------
# ENROLLMENT TABLE
# -----------------------------------------------------

enrollment_table_frame = tk.Frame(
    enrollment_scrollable_frame,
    bg=soft_pink
)


enrollment_table_frame.pack(
    fill="both",
    expand=True,
    padx=15,
    pady=(3, 10)
)


enrollment_columns = (
    "id",
    "student",
    "session",
    "date",
    "time",
    "class",
    "tuition",
    "paid",
    "balance",
    "payment",
    "certification"
)


enrollment_table = ttk.Treeview(
    enrollment_table_frame,
    columns=enrollment_columns,
    show="headings",
    height=7
)


enrollment_table.heading(
    "id",
    text="Enrollment ID"
)


enrollment_table.heading(
    "student",
    text="Student"
)


enrollment_table.heading(
    "session",
    text="Session ID"
)


enrollment_table.heading(
    "date",
    text="Date"
)


enrollment_table.heading(
    "time",
    text="Time"
)


enrollment_table.heading(
    "class",
    text="Class"
)


enrollment_table.heading(
    "tuition",
    text="Tuition"
)


enrollment_table.heading(
    "paid",
    text="Paid"
)


enrollment_table.heading(
    "balance",
    text="Balance"
)


enrollment_table.heading(
    "payment",
    text="Payment Status"
)


enrollment_table.heading(
    "certification",
    text="Certification"
)


enrollment_table.column(
    "id",
    width=85,
    anchor="center"
)


enrollment_table.column(
    "student",
    width=145
)


enrollment_table.column(
    "session",
    width=70,
    anchor="center"
)


enrollment_table.column(
    "date",
    width=90,
    anchor="center"
)


enrollment_table.column(
    "time",
    width=80,
    anchor="center"
)


enrollment_table.column(
    "class",
    width=115
)


enrollment_table.column(
    "tuition",
    width=85,
    anchor="e"
)


enrollment_table.column(
    "paid",
    width=85,
    anchor="e"
)


enrollment_table.column(
    "balance",
    width=85,
    anchor="e"
)


enrollment_table.column(
    "payment",
    width=105,
    anchor="center"
)


enrollment_table.column(
    "certification",
    width=105,
    anchor="center"
)


enrollment_vertical_scrollbar = ttk.Scrollbar(
    enrollment_table_frame,
    orient="vertical",
    command=enrollment_table.yview
)


enrollment_horizontal_scrollbar = ttk.Scrollbar(
    enrollment_table_frame,
    orient="horizontal",
    command=enrollment_table.xview
)


enrollment_table.configure(
    yscrollcommand=enrollment_vertical_scrollbar.set,
    xscrollcommand=enrollment_horizontal_scrollbar.set
)


enrollment_table.grid(
    row=0,
    column=0,
    sticky="nsew"
)


enrollment_vertical_scrollbar.grid(
    row=0,
    column=1,
    sticky="ns"
)


enrollment_horizontal_scrollbar.grid(
    row=1,
    column=0,
    sticky="ew"
)


enrollment_table_frame.rowconfigure(
    0,
    weight=1
)


enrollment_table_frame.columnconfigure(
    0,
    weight=1
)


# =====================================================
# PAYMENT LEDGER
# =====================================================

# A ledger is simply a transaction history.
#
# Instead of saying only "Deposit Paid", Lux can now
# remember the actual dollars and each separate payment.

payment_header = tk.Label(
    enrollment_scrollable_frame,
    text="PAYMENT LEDGER",
    font=("Arial", 13, "bold"),
    bg=lux_pink,
    fg=dark_text
)


payment_header.pack(
    fill="x",
    pady=(2, 2)
)


payment_enrollment_label = tk.Label(
    enrollment_scrollable_frame,
    text="Load an enrollment above to manage payments",
    font=("Arial", 9, "bold"),
    bg=soft_pink,
    fg=dark_text
)


payment_enrollment_label.pack(
    pady=(2, 1)
)


payment_summary_label = tk.Label(
    enrollment_scrollable_frame,
    text="",
    font=("Arial", 10, "bold"),
    bg=soft_pink,
    fg=gold
)


payment_summary_label.pack(
    pady=1
)


payment_legacy_note_label = tk.Label(
    enrollment_scrollable_frame,
    text="",
    font=("Arial", 8, "italic"),
    bg=soft_pink,
    fg=dark_text
)


payment_legacy_note_label.pack(
    pady=(0, 2)
)


payment_entry_frame = tk.Frame(
    enrollment_scrollable_frame,
    bg=soft_pink
)


payment_entry_frame.pack(
    pady=3
)


tk.Label(
    payment_entry_frame,
    text="Amount",
    font=("Arial", 9, "bold"),
    bg=soft_pink
).grid(
    row=0,
    column=0,
    padx=3
)


payment_amount_entry = tk.Entry(
    payment_entry_frame,
    width=10,
    justify="center"
)


payment_amount_entry.grid(
    row=0,
    column=1,
    padx=3
)


tk.Label(
    payment_entry_frame,
    text="Date",
    font=("Arial", 9, "bold"),
    bg=soft_pink
).grid(
    row=0,
    column=2,
    padx=3
)


payment_date_entry = tk.Entry(
    payment_entry_frame,
    width=12,
    justify="center"
)


payment_date_entry.insert(
    0,
    datetime.now().strftime(
        "%m/%d/%Y"
    )
)


payment_date_entry.grid(
    row=0,
    column=3,
    padx=3
)


tk.Label(
    payment_entry_frame,
    text="Method",
    font=("Arial", 9, "bold"),
    bg=soft_pink
).grid(
    row=0,
    column=4,
    padx=3
)


payment_method_combo = ttk.Combobox(
    payment_entry_frame,
    values=(
        "Cash",
        "Zelle",
        "Venmo",
        "Cash App",
        "Credit/Debit",
        "Other"
    ),
    state="readonly",
    width=13
)


payment_method_combo.set(
    "Cash"
)


payment_method_combo.grid(
    row=0,
    column=5,
    padx=3
)


tk.Label(
    payment_entry_frame,
    text="Notes",
    font=("Arial", 9, "bold"),
    bg=soft_pink
).grid(
    row=0,
    column=6,
    padx=3
)


payment_notes_entry = tk.Entry(
    payment_entry_frame,
    width=22
)


payment_notes_entry.grid(
    row=0,
    column=7,
    padx=3
)


payment_button_frame = tk.Frame(
    enrollment_scrollable_frame,
    bg=soft_pink
)


payment_button_frame.pack(
    pady=(2, 4)
)


add_payment_button = tk.Button(
    payment_button_frame,
    text="ADD PAYMENT",
    font=("Arial", 9, "bold"),
    bg=gold,
    fg="white",
    padx=12,
    pady=5,
    command=add_payment
)


add_payment_button.grid(
    row=0,
    column=0,
    padx=4
)


delete_payment_button = tk.Button(
    payment_button_frame,
    text="DELETE SELECTED PAYMENT",
    font=("Arial", 9, "bold"),
    padx=12,
    pady=5,
    command=delete_selected_payment
)


delete_payment_button.grid(
    row=0,
    column=1,
    padx=4
)


payment_history_frame = tk.Frame(
    enrollment_scrollable_frame,
    bg=soft_pink
)


payment_history_frame.pack(
    fill="both",
    expand=True,
    padx=15,
    pady=(2, 8)
)


payment_history_columns = (
    "id",
    "date",
    "amount",
    "method",
    "notes"
)


payment_history_table = ttk.Treeview(
    payment_history_frame,
    columns=payment_history_columns,
    show="headings",
    height=5
)


payment_history_table.heading(
    "id",
    text="Payment ID"
)


payment_history_table.heading(
    "date",
    text="Date"
)


payment_history_table.heading(
    "amount",
    text="Amount"
)


payment_history_table.heading(
    "method",
    text="Method"
)


payment_history_table.heading(
    "notes",
    text="Notes"
)


payment_history_table.column(
    "id",
    width=70,
    anchor="center"
)


payment_history_table.column(
    "date",
    width=95,
    anchor="center"
)


payment_history_table.column(
    "amount",
    width=90,
    anchor="e"
)


payment_history_table.column(
    "method",
    width=100,
    anchor="center"
)


payment_history_table.column(
    "notes",
    width=280
)


payment_history_scrollbar = ttk.Scrollbar(
    payment_history_frame,
    orient="vertical",
    command=payment_history_table.yview
)


payment_history_table.configure(
    yscrollcommand=payment_history_scrollbar.set
)


payment_history_table.grid(
    row=0,
    column=0,
    sticky="nsew"
)


payment_history_scrollbar.grid(
    row=0,
    column=1,
    sticky="ns"
)


payment_history_frame.rowconfigure(
    0,
    weight=1
)


payment_history_frame.columnconfigure(
    0,
    weight=1
)




# =====================================================
# BACKUP TAB
# =====================================================

backup_header = tk.Label(
    backup_tab,
    text="BACKUP & RESTORE",
    font=("Arial", 20, "bold"),
    bg=lux_pink,
    fg=dark_text
)


backup_header.pack(
    fill="x",
    pady=(0, 8)
)


tk.Label(
    backup_tab,
    text=(
        "Protect Lux business data and restore "
        "a previous snapshot when needed."
    ),
    font=("Arial", 11),
    bg=soft_pink,
    fg=dark_text
).pack(
    pady=(4, 12)
)


# -----------------------------------------------------
# CURRENT DATA STATUS
# -----------------------------------------------------

backup_current_frame = tk.LabelFrame(
    backup_tab,
    text="CURRENT DATA",
    font=("Arial", 11, "bold"),
    bg=soft_pink,
    fg=dark_text,
    padx=20,
    pady=12
)


backup_current_frame.pack(
    fill="x",
    padx=35,
    pady=(0, 12)
)


database_status_label = tk.Label(
    backup_current_frame,
    text="SQLite database: Checking...",
    font=("Arial", 11),
    bg=soft_pink,
    fg=dark_text,
    anchor="w"
)


database_status_label.pack(
    fill="x",
    pady=3
)


history_backup_status_label = tk.Label(
    backup_current_frame,
    text="CSV class-plan history: Checking...",
    font=("Arial", 11),
    bg=soft_pink,