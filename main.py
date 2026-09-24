from flask import (
    Flask,
    render_template,
    request,
    jsonify,
    send_file,
    redirect,
    url_for,
    session,
)
from functools import wraps
import pandas as pd
from datetime import datetime, date, timedelta, time
import calendar
import io
import re
import math
import json
import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from werkzeug.security import generate_password_hash, check_password_hash
from sqlalchemy import create_engine, text
import holidays
from zoneinfo import ZoneInfo

# --- Database Setup (Neon PostgreSQL) ---
DATABASE_URL = os.environ.get("DATABASE_URL")
if DATABASE_URL and DATABASE_URL.startswith("postgres://"):
    DATABASE_URL = DATABASE_URL.replace("postgres://", "postgresql://", 1)[cite: 1]

engine = create_engine(DATABASE_URL) if DATABASE_URL else None[cite: 1]

if engine:
    try:
        with engine.begin() as conn:
            conn.execute(text("""
                CREATE TABLE IF NOT EXISTS save_slots (
                    slot_id VARCHAR(100) PRIMARY KEY,
                    username VARCHAR(50) NOT NULL,
                    title VARCHAR(255) NOT NULL,
                    updated_at VARCHAR(50) NOT NULL,
                    start_date VARCHAR(20),
                    end_date VARCHAR(20),
                    total_employees INT,
                    payload JSONB NOT NULL
                );
                CREATE TABLE IF NOT EXISTS company_holidays (
                    id SERIAL PRIMARY KEY,
                    holiday_date VARCHAR(10) UNIQUE NOT NULL,
                    holiday_name VARCHAR(255) NOT NULL,
                    created_at VARCHAR(50) NOT NULL
                );
            """))
            print("Database tables initialized successfully.")
    except Exception as e:
        print("Database initialization error:", e)[cite: 1]

app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "prod-session-key-attendance-2026")[cite: 1]

DRAFTS_DIR = os.path.join(os.path.dirname(__file__), "user_drafts")[cite: 1]
os.makedirs(DRAFTS_DIR, exist_ok=True)[cite: 1]

# Admin Credentials Store
USERS = {"admin": generate_password_hash("jiaen123")}[cite: 1]

USER_DATAFRAMES = {}[cite: 1]

# Business Rules
WORK_START_TIME = time(9, 0, 0)[cite: 1]
WEEKDAY_END_TIME = time(18, 0, 0)[cite: 1]
SATURDAY_END_TIME = time(13, 30, 0)[cite: 1]
LUNCH_LIMIT_MINS = 70[cite: 1]

# Malaysia Public Holidays Engine (English-first)
MY_HOLIDAYS = holidays.country_holidays("MY", language="en")[cite: 1]

HOLIDAY_EN_MAP = {
    "Hari Kebangsaan": "National Day",
    "Hari Malaysia": "Malaysia Day",
    "Hari Pekerja": "Labour Day",
    "Tahun Baru Cina": "Chinese New Year",
    "Hari Raya Aidilfitri": "Hari Raya Aidilfitri",
    "Hari Raya Aidiladha": "Hari Raya Haji",
    "Hari Deepavali": "Deepavali",
    "Hari Keputeraan Yang di-Pertuan Agong": "King's Birthday",
    "Hari Keputeraan Nabi Muhammad S.A.W. (Maulidur Rasul)": "Prophet Muhammad's Birthday",
    "Hari Krismas": "Christmas Day",
    "Tahun Baru": "New Year's Day",
    "Hari Wesak": "Wesak Day",
    "Awal Muharram": "Awal Muharram",
    "Thaipusam": "Thaipusam",
}[cite: 1]


def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if "user" not in session:[cite: 1]
            if request.path.startswith("/api/"):[cite: 1]
                return jsonify({"error": "Authentication required"}), 401[cite: 1]
            return redirect(url_for("login"))[cite: 1]
        return f(*args, **kwargs)[cite: 1]

    return decorated_function[cite: 1]


def load_user_slots(username):
    if not engine:[cite: 1]
        return [][cite: 1]
    try:
        with engine.connect() as conn:[cite: 1]
            result = conn.execute(
                text(
                    "SELECT payload FROM save_slots WHERE username = :u ORDER BY updated_at DESC"
                ),
                {"u": username},
            )[cite: 1]
            return [row[0] for row in result.fetchall()][cite: 1]
    except Exception as e:
        print("Error loading slots:", e)[cite: 1]
        return [][cite: 1]


def save_user_slots(username, slot_entry):
    if not engine:[cite: 1]
        return[cite: 1]
    try:
        with engine.begin() as conn:[cite: 1]
            conn.execute(
                text("""
                INSERT INTO save_slots (slot_id, username, title, updated_at, start_date, end_date, total_employees, payload)
                VALUES (:id, :u, :t, :time, :s_date, :e_date, :tot, :p)
                ON CONFLICT (slot_id) DO UPDATE 
                SET title = EXCLUDED.title,
                    updated_at = EXCLUDED.updated_at,
                    start_date = EXCLUDED.start_date,
                    end_date = EXCLUDED.end_date,
                    total_employees = EXCLUDED.total_employees,
                    payload = EXCLUDED.payload;
            """),
                {
                    "id": slot_entry["id"],
                    "u": username,
                    "t": slot_entry["title"],
                    "time": slot_entry["updatedAt"],
                    "s_date": slot_entry.get("startDate", ""),
                    "e_date": slot_entry.get("endDate", ""),
                    "tot": slot_entry.get("totalEmployees", 0),
                    "p": json.dumps(slot_entry),
                },
            )[cite: 1]
    except Exception as e:
        print("Error saving slot to DB:", e)[cite: 1]


def delete_user_slot(username, slot_id):
    if not engine:[cite: 1]
        return[cite: 1]
    try:
        with engine.begin() as conn:[cite: 1]
            conn.execute(
                text("DELETE FROM save_slots WHERE slot_id = :id AND username = :u"),
                {"id": slot_id, "u": username},
            )[cite: 1]
    except Exception as e:
        print("Error deleting slot from DB:", e)[cite: 1]


def load_db_holidays():
    if not engine:
        return {}
    try:
        with engine.connect() as conn:
            result = conn.execute(text("SELECT holiday_date, holiday_name FROM company_holidays ORDER BY holiday_date ASC"))
            return {row[0]: row[1] for row in result.fetchall()}
    except Exception as e:
        print("Error loading company holidays:", e)
        return {}


def save_db_holiday(h_date, h_name):
    if not engine:
        return False
    try:
        myt_now = datetime.now(ZoneInfo("Asia/Kuala_Lumpur")).strftime("%Y-%m-%d %H:%M:%S")
        with engine.begin() as conn:
            conn.execute(text("""
                INSERT INTO company_holidays (holiday_date, holiday_name, created_at)
                VALUES (:d, :n, :c)
                ON CONFLICT (holiday_date) DO UPDATE
                SET holiday_name = EXCLUDED.holiday_name,
                    created_at = EXCLUDED.created_at;
            """), {"d": h_date, "n": h_name, "c": myt_now})
        return True
    except Exception as e:
        print("Error saving company holiday:", e)
        return False


def delete_db_holiday(h_date):
    if not engine:
        return False
    try:
        with engine.begin() as conn:
            conn.execute(text("DELETE FROM company_holidays WHERE holiday_date = :d"), {"d": h_date})
        return True
    except Exception as e:
        print("Error deleting company holiday:", e)
        return False


def get_malaysia_holiday_name(d_obj, db_holidays_map=None):
    d_str = d_obj.strftime("%Y-%m-%d")
    if db_holidays_map and d_str in db_holidays_map:
        return db_holidays_map[d_str]

    raw_name = MY_HOLIDAYS.get(d_obj)[cite: 1]
    if not raw_name:[cite: 1]
        return None[cite: 1]
    for my_name, en_name in HOLIDAY_EN_MAP.items():[cite: 1]
        if my_name.lower() in raw_name.lower():[cite: 1]
            return en_name[cite: 1]
    return raw_name.replace(" (Observed)", "").replace(" (observed)", "").strip()[cite: 1]


def clean_time_str(t_str):
    """Truncates :SS from HH:MM:SS without rounding up, yielding clean HH:MM."""
    if not t_str or pd.isna(t_str):
        return ""
    t_str = str(t_str).strip()
    parts = t_str.split(":")
    if len(parts) >= 2:
        return f"{parts[0].zfill(2)}:{parts[1].zfill(2)}"
    return t_str


def parse_time_str(t_str):
    """Parses HH:MM or HH:MM:SS time string strictly at minute-level, ignoring seconds."""
    if not t_str or str(t_str).strip() in ["--", "nan", ""]:
        return None
    s = str(t_str).strip()
    try:
        return datetime.strptime(s[:5], "%H:%M")
    except Exception:
        return None


def format_mins_to_time(minutes):
    if pd.isna(minutes) or minutes <= 0:[cite: 1]
        return "0:00"[cite: 1]
    hrs = int(minutes // 60)[cite: 1]
    mins = int(round(minutes % 60))[cite: 1]
    if mins == 60:[cite: 1]
        hrs += 1[cite: 1]
        mins = 0[cite: 1]
    return f"{hrs}:{mins:02d}"[cite: 1]


def deduplicate_close_punches(punch_list, threshold_minutes=3):
    valid_punches = [][cite: 1]
    close_duplicates = [][cite: 1]

    for p in punch_list:[cite: 1]
        dt = parse_time_str(p)[cite: 1]
        if not dt:[cite: 1]
            continue[cite: 1]
        if not valid_punches:[cite: 1]
            valid_punches.append((dt, p))[cite: 1]
            continue[cite: 1]

        prev_dt, _ = valid_punches[-1][cite: 1]
        diff_sec = (dt - prev_dt).total_seconds()[cite: 1]
        if diff_sec <= (threshold_minutes * 60):[cite: 1]
            close_duplicates.append(p)[cite: 1]
        else:
            valid_punches.append((dt, p))[cite: 1]

    return [p for _, p in valid_punches], close_duplicates[cite: 1]


def categorize_punches(raw_punch_list, is_saturday):
    punch_list, close_duplicates = deduplicate_close_punches(
        raw_punch_list, threshold_minutes=3
    )[cite: 1]
    p_tuples = [][cite: 1]
    for p in punch_list:[cite: 1]
        dt = parse_time_str(p)[cite: 1]
        if dt:[cite: 1]
            p_tuples.append((dt.time(), p))[cite: 1]

    c_in, b_out, b_in, c_out = "--", "--", "--", "--"[cite: 1]
    assigned_punches = [][cite: 1]

    if is_saturday:[cite: 1]
        morning = [
            p
            for p in p_tuples
            if p[0].hour < 11 or (p[0].hour == 11 and p[0].minute < 30)
        ][cite: 1]
        afternoon = [
            p
            for p in p_tuples
            if p[0].hour > 11 or (p[0].hour == 11 and p[0].minute >= 30)
        ][cite: 1]
        if morning:[cite: 1]
            c_in = morning[0][1][cite: 1]
            assigned_punches.append(c_in)[cite: 1]
        if afternoon:[cite: 1]
            c_out = afternoon[-1][1][cite: 1]
            assigned_punches.append(c_out)[cite: 1]
        elif len(morning) > 1:[cite: 1]
            c_out = morning[-1][1][cite: 1]
            assigned_punches.append(c_out)[cite: 1]
    else:
        morning = [
            p
            for p in p_tuples
            if p[0].hour < 11 or (p[0].hour == 11 and p[0].minute < 15)
        ][cite: 1]
        evening = [
            p
            for p in p_tuples
            if p[0].hour > 15 or (p[0].hour == 15 and p[0].minute > 30)
        ][cite: 1]

        if morning:[cite: 1]
            c_in = morning[0][1][cite: 1]
            assigned_punches.append(c_in)[cite: 1]

        lunch_out_candidates = [
            p
            for p in p_tuples
            if (p[0].hour == 11 and p[0].minute >= 15)
            or (p[0].hour == 12)
            or (p[0].hour == 13 and p[0].minute < 15)
        ][cite: 1]
        lunch_in_candidates = [
            p
            for p in p_tuples
            if (p[0].hour == 13 and p[0].minute >= 15)
            or (p[0].hour == 14)
            or (p[0].hour == 15 and p[0].minute <= 30)
        ][cite: 1]

        if lunch_out_candidates:[cite: 1]
            b_out = lunch_out_candidates[-1][1][cite: 1]
            assigned_punches.append(b_out)[cite: 1]
        if lunch_in_candidates:[cite: 1]
            b_in = lunch_in_candidates[0][1][cite: 1]
            assigned_punches.append(b_in)[cite: 1]

        if b_out == "--" and b_in == "--":[cite: 1]
            all_lunch = [
                p
                for p in p_tuples
                if (p[0].hour > 11 or (p[0].hour == 11 and p[0].minute >= 15))
                and (p[0].hour < 15 or (p[0].hour == 15 and p[0].minute <= 30))
            ][cite: 1]
            if len(all_lunch) >= 2:[cite: 1]
                b_out = all_lunch[0][1][cite: 1]
                b_in = all_lunch[-1][1][cite: 1]
                assigned_punches.extend([b_out, b_in])[cite: 1]
            elif len(all_lunch) == 1:[cite: 1]
                if all_lunch[0][0].hour < 13:[cite: 1]
                    b_out = all_lunch[0][1][cite: 1]
                else:
                    b_in = all_lunch[0][1][cite: 1]
                assigned_punches.append(all_lunch[0][1])[cite: 1]

        if evening:[cite: 1]
            c_out = evening[-1][1][cite: 1]
            assigned_punches.append(c_out)[cite: 1]

    extra_punches = [
        p for p in punch_list if p not in assigned_punches
    ] + close_duplicates[cite: 1]
    multi_earlier = extra_punches[0] if len(extra_punches) >= 1 else "--"[cite: 1]
    multi_later = extra_punches[-1] if len(extra_punches) >= 2 else "--"[cite: 1]

    issues = [][cite: 1]
    if is_saturday:[cite: 1]
        if c_in == "--":[cite: 1]
            issues.append("Missing Clock In")[cite: 1]
        if c_out == "--":[cite: 1]
            issues.append("Missing Clock Out")[cite: 1]
        if len(punch_list) > 2:[cite: 1]
            issues.append(f"Multiple Punches ({len(punch_list)})")[cite: 1]
        status = (
            "Saturday (Half Day)" if not issues else f"Saturday ({', '.join(issues)})"
        )[cite: 1]
    else:
        if c_in == "--":[cite: 1]
            issues.append("Missing Clock In")[cite: 1]
        if b_out == "--" and b_in != "--":[cite: 1]
            issues.append("Missing Break Out")[cite: 1]
        if b_out != "--" and b_in == "--":[cite: 1]
            issues.append("Missing Break In")[cite: 1]
        if b_out == "--" and b_in == "--":[cite: 1]
            issues.append("No Lunch Punched")[cite: 1]
        if c_out == "--":[cite: 1]
            issues.append("Missing Clock Out")[cite: 1]
        if len(punch_list) > 4:[cite: 1]
            issues.append(f"Multiple Punches ({len(punch_list)})")[cite: 1]
        status = "Normal" if not issues else ", ".join(issues)[cite: 1]

    return c_in, b_out, b_in, c_out, status, multi_earlier, multi_later[cite: 1]


def process_time_card(
    df_raw, start_date_str=None, end_date_str=None, special_entries=None
):
    if special_entries is None:[cite: 1]
        special_entries = [][cite: 1]

    db_holidays = load_db_holidays()

    special_lookup = {}[cite: 1]
    for entry in special_entries:[cite: 1]
        special_lookup[(entry["date"], entry["target"])] = (
            entry["type"],
            entry.get("remark", "").strip(),
        )[cite: 1]

    header_idx = None[cite: 1]
    for idx, row in df_raw.iloc[:5].iterrows():[cite: 1]
        if "Employee ID" in row.values:[cite: 1]
            header_idx = idx[cite: 1]
            break[cite: 1]

    if header_idx is not None:[cite: 1]
        df = df_raw.iloc[header_idx + 1 :].copy()[cite: 1]
        df.columns = df_raw.iloc[header_idx].values[cite: 1]
    else:
        df = df_raw.copy()[cite: 1]

    df.columns = [str(c).strip() for c in df.columns][cite: 1]

    raw_punches = {}[cite: 1]
    emp_meta = {}[cite: 1]
    dates_in_file = [][cite: 1]

    for _, row in df.iterrows():[cite: 1]
        emp_id = str(row.get("Employee ID", "")).strip()[cite: 1]
        name = str(row.get("First Name", "")).strip()[cite: 1]
        dept = str(row.get("Department", "")).strip()[cite: 1]
        date_str = str(row.get("Date", "")).split()[0][cite: 1]
        raw_times_str = str(row.get("Time", ""))[cite: 1]

        if (
            pd.isna(raw_times_str)
            or not raw_times_str.strip()
            or raw_times_str == "nan"
        ):[cite: 1]
            continue[cite: 1]

        emp_meta[emp_id] = {"name": name, "dept": dept}[cite: 1]
        
        # Clean every punch to HH:MM immediately upon reading file (ignore seconds without rounding)
        raw_times_list = [clean_time_str(t) for t in raw_times_str.split(",") if t.strip()]
        raw_punches[(emp_id, date_str)] = ",".join(raw_times_list)

        try:
            dates_in_file.append(datetime.strptime(date_str, "%Y-%m-%d").date())[cite: 1]
        except Exception:
            pass[cite: 1]

    if dates_in_file:[cite: 1]
        detected_date = dates_in_file[0][cite: 1]
        year = detected_date.year[cite: 1]
        month = detected_date.month[cite: 1]
        last_day = calendar.monthrange(year, month)[1][cite: 1]
        min_d = date(year, month, 1)[cite: 1]
        max_d = date(year, month, last_day)[cite: 1]
    else:
        today = date.today()[cite: 1]
        last_day = calendar.monthrange(today.year, today.month)[1][cite: 1]
        min_d = date(today.year, today.month, 1)[cite: 1]
        max_d = date(today.year, today.month, last_day)[cite: 1]

    if start_date_str:[cite: 1]
        try:
            min_d = datetime.strptime(start_date_str, "%Y-%m-%d").date()[cite: 1]
        except Exception:
            pass[cite: 1]
    if end_date_str:[cite: 1]
        try:
            max_d = datetime.strptime(end_date_str, "%Y-%m-%d").date()[cite: 1]
        except Exception:
            pass[cite: 1]

    calendar_dates = [][cite: 1]
    curr = min_d[cite: 1]
    while curr <= max_d:[cite: 1]
        calendar_dates.append(curr)[cite: 1]
        curr += timedelta(days=1)[cite: 1]

    records = [][cite: 1]
    for emp_id, meta in sorted(emp_meta.items()):[cite: 1]
        name = meta["name"][cite: 1]
        dept = meta["dept"][cite: 1]

        for w_date in calendar_dates:[cite: 1]
            date_str = w_date.strftime("%Y-%m-%d")[cite: 1]
            is_sunday = w_date.weekday() == 6[cite: 1]
            is_saturday = w_date.weekday() == 5[cite: 1]
            day_name = w_date.strftime("%a")[cite: 1]

            # Sunday Protection: Always strictly Sunday
            if is_sunday:[cite: 1]
                records.append(
                    {
                        "Employee ID": emp_id,
                        "Name": name,
                        "Department": dept,
                        "Date": f"{date_str} ({day_name})",
                        "Clock In": "--",
                        "Break Out (Lunch)": "--",
                        "Break In (Back)": "--",
                        "Clock Out": "--",
                        "Lunch Duration": "--",
                        "Late to Work": "--",
                        "Late Time (Lunch)": "--",
                        "Total Late Time": "--",
                        "Early Leave": "--",
                        "Work Hours": "--",
                        "Status / Alert": "Sunday",
                        "Multiple Punch (Earlier)": "--",
                        "Multiple Punch (Later)": "--",
                        "_work_mins": 0,
                        "_lunch_mins": 0,
                        "_late_work_mins": 0,
                        "_late_lunch_mins": 0,
                        "_total_late_mins": 0,
                        "_early_leave_mins": 0,
                        "_is_absent": 0,
                        "_is_sunday": 1,
                        "_is_offday": 1,
                    }
                )[cite: 1]
                continue[cite: 1]

            special_info = special_lookup.get((date_str, name)) or special_lookup.get(
                (date_str, "ALL")
            )[cite: 1]
            special_type = None[cite: 1]

            if special_info:[cite: 1]
                st_type, st_remark = special_info[cite: 1]
                if st_remark:[cite: 1]
                    special_type = (
                        f"{st_type} ({st_remark})"
                        if st_type not in st_remark
                        else st_remark
                    )[cite: 1]
                else:
                    special_type = st_type

            # Check database-configured custom holidays
            if not special_type and date_str in db_holidays:
                special_type = f"Public Holiday ({db_holidays[date_str]})"

            raw_times_str = raw_punches.get((emp_id, date_str))[cite: 1]
            is_half_day = bool(special_type and "(0.5 Day)" in special_type)[cite: 1]

            # Days without biometric punches
            if not raw_times_str:[cite: 1]
                if special_type and not is_half_day:[cite: 1]
                    status_text = special_type[cite: 1]
                    is_off = 1[cite: 1]
                else:
                    # Absence concept replaced with punch irregularity
                    if is_half_day:[cite: 1]
                        status_text = f"{special_type} (Missing Clock In, Missing Clock Out)"[cite: 1]
                    elif is_saturday:[cite: 1]
                        status_text = "Saturday (Missing Clock In, Missing Clock Out)"[cite: 1]
                    else:
                        status_text = "Missing Clock In, No Lunch Punched, Missing Clock Out"[cite: 1]
                    is_off = 0[cite: 1]

                records.append(
                    {
                        "Employee ID": emp_id,
                        "Name": name,
                        "Department": dept,
                        "Date": f"{date_str} ({day_name})",
                        "Clock In": "--",
                        "Break Out (Lunch)": "--",
                        "Break In (Back)": "--",
                        "Clock Out": "--",
                        "Lunch Duration": "--",
                        "Late to Work": "--",
                        "Late Time (Lunch)": "--",
                        "Total Late Time": "--",
                        "Early Leave": "--",
                        "Work Hours": "0:00" if not is_off else "--",
                        "Status / Alert": status_text,
                        "Multiple Punch (Earlier)": "--",
                        "Multiple Punch (Later)": "--",
                        "_work_mins": 0,
                        "_lunch_mins": 0,
                        "_late_work_mins": 0,
                        "_late_lunch_mins": 0,
                        "_total_late_mins": 0,
                        "_early_leave_mins": 0,
                        "_is_absent": 0,
                        "_is_sunday": 0,
                        "_is_offday": is_off,
                    }
                )[cite: 1]
                continue[cite: 1]

            punch_list = [t.strip() for t in raw_times_str.split(",") if t.strip()][cite: 1]

            # Handle 0.5 Day leave: Earlier = Clock In, Later = Clock Out, Lunch = None
            if is_half_day and not is_saturday:[cite: 1]
                clean_punches, extra_dups = deduplicate_close_punches(
                    punch_list, threshold_minutes=3
                )[cite: 1]
                sorted_punches = sorted(
                    clean_punches, key=lambda p: parse_time_str(p) or datetime.min
                )[cite: 1]

                if len(sorted_punches) >= 2:[cite: 1]
                    c_in = sorted_punches[0][cite: 1]
                    c_out = sorted_punches[-1][cite: 1]
                    extra = sorted_punches[1:-1] + extra_dups[cite: 1]
                elif len(sorted_punches) == 1:[cite: 1]
                    dt = parse_time_str(sorted_punches[0])[cite: 1]
                    if dt and dt.hour < 12:[cite: 1]
                        c_in = sorted_punches[0][cite: 1]
                        c_out = "--"[cite: 1]
                    else:
                        c_in = "--"[cite: 1]
                        c_out = sorted_punches[0][cite: 1]
                    extra = extra_dups[cite: 1]
                else:
                    c_in, c_out, extra = "--", "--", [][cite: 1]

                b_out, b_in = "--", "--"[cite: 1]
                multi_earlier = extra[0] if len(extra) >= 1 else "--"[cite: 1]
                multi_later = extra[-1] if len(extra) >= 2 else "--"[cite: 1]

                issues = [][cite: 1]
                if c_in == "--":[cite: 1]
                    issues.append("Missing Clock In")[cite: 1]
                if c_out == "--":[cite: 1]
                    issues.append("Missing Clock Out")[cite: 1]
                status = (
                    f"{special_type} ({', '.join(issues)})"
                    if issues
                    else special_type
                )[cite: 1]
            else:
                c_in, b_out, b_in, c_out, status, multi_earlier, multi_later = (
                    categorize_punches(punch_list, is_saturday)
                )[cite: 1]
                if special_type:[cite: 1]
                    status = (
                        f"{special_type} (Worked)"
                        if status == "Normal"
                        else f"{special_type} ({status})"
                    )[cite: 1]

            lunch_mins, work_mins, late_work_mins, late_lunch_mins, early_leave_mins = (
                0,
                0,
                0,
                0,
                0,
            )[cite: 1]

            if c_in != "--":[cite: 1]
                dt_cin = parse_time_str(c_in)[cite: 1]
                if dt_cin:[cite: 1]
                    start_dt = dt_cin.replace(
                        hour=WORK_START_TIME.hour,
                        minute=WORK_START_TIME.minute,
                    )[cite: 1]
                    if dt_cin > start_dt:[cite: 1]
                        late_work_mins = round((dt_cin - start_dt).total_seconds() / 60)[cite: 1]

            if not is_saturday and not is_half_day and b_out != "--" and b_in != "--":[cite: 1]
                dt_bout = parse_time_str(b_out)[cite: 1]
                dt_bin = parse_time_str(b_in)[cite: 1]
                if dt_bout and dt_bin and dt_bin > dt_bout:[cite: 1]
                    lunch_mins = round((dt_bin - dt_bout).total_seconds() / 60)[cite: 1]
                    if lunch_mins > LUNCH_LIMIT_MINS:[cite: 1]
                        late_lunch_mins = lunch_mins - LUNCH_LIMIT_MINS[cite: 1]

            early_remark = ""[cite: 1]
            if c_out != "--":[cite: 1]
                dt_cout = parse_time_str(c_out)[cite: 1]
                if dt_cout:[cite: 1]
                    target_end = SATURDAY_END_TIME if is_saturday else WEEKDAY_END_TIME[cite: 1]
                    end_dt = dt_cout.replace(
                        hour=target_end.hour,
                        minute=target_end.minute,
                    )[cite: 1]
                    if not is_half_day and dt_cout < end_dt:[cite: 1]
                        early_leave_mins = round(
                            (end_dt - dt_cout).total_seconds() / 60
                        )[cite: 1]
                        if early_leave_mins > 0:[cite: 1]
                            half_hour_blocks = math.ceil(early_leave_mins / 30)[cite: 1]
                            if half_hour_blocks == 1:[cite: 1]
                                early_remark = "Early up 30 mins"[cite: 1]
                            else:
                                total_hours = half_hour_blocks * 0.5[cite: 1]
                                hr_str = (
                                    f"{int(total_hours)}"
                                    if total_hours.is_integer()
                                    else f"{total_hours}"
                                )[cite: 1]
                                hr_label = "hour" if total_hours == 1.0 else "hours"[cite: 1]
                                early_remark = f"Early up {hr_str} {hr_label}"[cite: 1]

            if early_remark:[cite: 1]
                if status in ["Normal", "Saturday (Half Day)"] or any(
                    k in status for k in ["off day", "Leave", "Holiday", "MC"]
                ):[cite: 1]
                    status = f"{status} ({early_remark})"[cite: 1]
                else:
                    status = f"{status}, {early_remark}"[cite: 1]

            total_late_mins = late_work_mins + late_lunch_mins[cite: 1]

            if c_in != "--" and c_out != "--":[cite: 1]
                dt_cin = parse_time_str(c_in)[cite: 1]
                dt_cout = parse_time_str(c_out)[cite: 1]
                if dt_cin and dt_cout and dt_cout > dt_cin:[cite: 1]
                    gross_mins = round((dt_cout - dt_cin).total_seconds() / 60)[cite: 1]
                    work_mins = max(0, gross_mins - lunch_mins)[cite: 1]

            records.append(
                {
                    "Employee ID": emp_id,
                    "Name": name,
                    "Department": dept,
                    "Date": f"{date_str} ({day_name})",
                    "Clock In": c_in,
                    "Break Out (Lunch)": b_out,
                    "Break In (Back)": b_in,
                    "Clock Out": c_out,
                    "Lunch Duration": (
                        format_mins_to_time(lunch_mins) if lunch_mins > 0 else "--"
                    ),
                    "Late to Work": (
                        format_mins_to_time(late_work_mins)
                        if late_work_mins > 0
                        else "--"
                    ),
                    "Late Time (Lunch)": (
                        format_mins_to_time(late_lunch_mins)
                        if late_lunch_mins > 0
                        else "--"
                    ),
                    "Total Late Time": (
                        format_mins_to_time(total_late_mins)
                        if total_late_mins > 0
                        else "--"
                    ),
                    "Early Leave": (
                        format_mins_to_time(early_leave_mins)
                        if early_leave_mins > 0
                        else "--"
                    ),
                    "Work Hours": (
                        format_mins_to_time(work_mins) if work_mins > 0 else "--"
                    ),
                    "Status / Alert": status,
                    "Multiple Punch (Earlier)": multi_earlier,
                    "Multiple Punch (Later)": multi_later,
                    "_work_mins": work_mins,
                    "_lunch_mins": lunch_mins,
                    "_late_work_mins": late_work_mins,
                    "_late_lunch_mins": late_lunch_mins,
                    "_total_late_mins": total_late_mins,
                    "_early_leave_mins": early_leave_mins,
                    "_is_absent": 0,
                    "_is_sunday": 0,
                    "_is_offday": 1 if (special_type and not is_half_day) else 0,
                }
            )[cite: 1]

    return pd.DataFrame(records), min_d.strftime("%Y-%m-%d"), max_d.strftime("%Y-%m-%d")[cite: 1]


def build_excel_workbook(df):
    wb = openpyxl.Workbook()[cite: 1]
    wb.remove(wb.active)[cite: 1]

    header_fill = PatternFill(
        start_color="0F2D59", end_color="0F2D59", fill_type="solid"
    )[cite: 1]
    total_fill = PatternFill(
        start_color="D9E1F2", end_color="D9E1F2", fill_type="solid"
    )[cite: 1]
    offday_fill = PatternFill(
        start_color="FFF5F5", end_color="FFF5F5", fill_type="solid"
    )[cite: 1]
    late_fill = PatternFill(start_color="FEF08A", end_color="FEF08A", fill_type="solid")[cite: 1]
    alert_fill = PatternFill(
        start_color="FED7AA", end_color="FED7AA", fill_type="solid"
    )[cite: 1]
    multi_fill = PatternFill(
        start_color="F1F5F9", end_color="F1F5F9", fill_type="solid"
    )[cite: 1]

    font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")[cite: 1]
    font_bold = Font(name="Calibri", size=11, bold=True)[cite: 1]
    font_regular = Font(name="Calibri", size=11)[cite: 1]
    font_merged_banner = Font(name="Calibri", size=11, bold=True, color="DC2626")[cite: 1]

    thin_border = Border(
        left=Side(style="thin", color="CBD5E1"),
        right=Side(style="thin", color="CBD5E1"),
        top=Side(style="thin", color="CBD5E1"),
        bottom=Side(style="thin", color="CBD5E1"),
    )[cite: 1]
    align_center = Alignment(horizontal="center", vertical="center")[cite: 1]
    align_left = Alignment(horizontal="left", vertical="center")[cite: 1]

    # 1. Summary Sheet
    ws_summary = wb.create_sheet(title="Overview Summary")[cite: 1]
    ws_summary.views.sheetView[0].showGridLines = True[cite: 1]

    summary_headers = [
        "Employee ID",
        "Employee Name",
        "Department",
        "Total Days Worked",
        "Total Work Hours",
        "Total Lunch Time",
        "Total Late (Work)",
        "Total Late (Lunch)",
        "Total Late (Sum)",
        "Total Early Leave",
        "Punch Irregularities Count",
    ][cite: 1]
    ws_summary.append(summary_headers)[cite: 1]

    for col_idx in range(1, len(summary_headers) + 1):[cite: 1]
        cell = ws_summary.cell(row=1, column=col_idx)[cite: 1]
        cell.fill = header_fill[cite: 1]
        cell.font = font_header[cite: 1]
        cell.alignment = align_center[cite: 1]

    non_work_categories = [
        "Holiday",
        "Public Holiday",
        "Annual Leave",
        "Replacement Leave",
        "MC",
        "Unpaid Leave",
        "Team A off day",
        "Team B off day",
        "Sunday",
    ][cite: 1]

    overview_data = [][cite: 1]
    for (emp_id, emp_name), emp_group in df.groupby(
        ["Employee ID", "Name"], sort=False
    ):[cite: 1]
        dept = emp_group["Department"].iloc[0][cite: 1]
        days_present = len(
            emp_group[
                (emp_group["Clock In"] != "--") | (emp_group["Clock Out"] != "--")
            ]
        )[cite: 1]
        total_work_m = emp_group["_work_mins"].sum()[cite: 1]
        total_lunch_m = emp_group["_lunch_mins"].sum()[cite: 1]
        total_late_w_m = emp_group["_late_work_mins"].sum()[cite: 1]
        total_late_l_m = emp_group["_late_lunch_mins"].sum()[cite: 1]
        grand_total_late_m = emp_group["_total_late_mins"].sum()[cite: 1]
        total_early_m = emp_group["_early_leave_mins"].sum()[cite: 1]
        anomalies = len(
            emp_group[
                ~emp_group["Status / Alert"].str.startswith("Normal")
                & ~emp_group["Status / Alert"].str.startswith("Saturday (Half Day)")
                & (emp_group["Status / Alert"] != "Sunday")
                & (
                    ~emp_group["Status / Alert"].str.contains(
                        "|".join(non_work_categories)
                    )
                    | emp_group["Status / Alert"].str.contains(r"\(0\.5 Day\)")
                )
            ]
        )[cite: 1]

        overview_data.append(
            [
                emp_id,
                emp_name,
                dept,
                days_present,
                format_mins_to_time(total_work_m),
                format_mins_to_time(total_lunch_m),
                format_mins_to_time(total_late_w_m),
                format_mins_to_time(total_late_l_m),
                format_mins_to_time(grand_total_late_m),
                format_mins_to_time(total_early_m),
                anomalies,
            ]
        )[cite: 1]

    for row_idx, row_vals in enumerate(overview_data, start=2):[cite: 1]
        for col_idx, val in enumerate(row_vals, start=1):[cite: 1]
            cell = ws_summary.cell(row=row_idx, column=col_idx, value=val)[cite: 1]
            cell.font = font_regular[cite: 1]
            cell.border = thin_border[cite: 1]
            cell.alignment = align_center if col_idx not in [2, 3] else align_left[cite: 1]

            if col_idx in [7, 8, 9, 10] and str(val) != "0:00":[cite: 1]
                cell.fill = late_fill[cite: 1]
                cell.font = font_bold[cite: 1]
            elif col_idx == 11 and val > 0:[cite: 1]
                cell.fill = alert_fill[cite: 1]
                cell.font = font_bold[cite: 1]

    for col in ws_summary.columns:[cite: 1]
        max_len = max(len(str(cell.value or "")) for cell in col)[cite: 1]
        col_letter = get_column_letter(col[0].column)[cite: 1]
        ws_summary.column_dimensions[col_letter].width = max(max_len + 4, 14)[cite: 1]

    # 2. Individual Sheets
    employee_cols = [
        "Date",
        "Clock In",
        "Break Out (Lunch)",
        "Break In (Back)",
        "Clock Out",
        "Lunch Duration",
        "Late to Work",
        "Late Time (Lunch)",
        "Total Late Time",
        "Early Leave",
        "Work Hours",
        "Status / Alert",
        "Multiple Punch (Earlier)",
        "Multiple Punch (Later)",
    ][cite: 1]

    for (emp_id, emp_name), emp_group in df.groupby(
        ["Employee ID", "Name"], sort=False
    ):[cite: 1]
        safe_name = re.sub(r"[\/\\\?\*\:\[\]]", "", emp_name).strip()[cite: 1]
        sheet_title = safe_name[:31] if safe_name else f"Emp_{emp_id}"[cite: 1]

        counter = 1[cite: 1]
        orig_title = sheet_title[cite: 1]
        while sheet_title in wb.sheetnames:[cite: 1]
            sheet_title = f"{orig_title[:28]}_{counter}"[cite: 1]
            counter += 1[cite: 1]

        ws_emp = wb.create_sheet(title=sheet_title)[cite: 1]
        ws_emp.views.sheetView[0].showGridLines = True[cite: 1]

        ws_emp["A1"] = f"Employee ID: {emp_id}"[cite: 1]
        ws_emp["A1"].font = font_bold[cite: 1]
        ws_emp["C1"] = f"Name: {emp_name}"[cite: 1]
        ws_emp["C1"].font = font_bold[cite: 1]
        ws_emp["G1"] = f"Department: {emp_group['Department'].iloc[0]}"[cite: 1]
        ws_emp["G1"].font = font_bold[cite: 1]

        for col_idx, h_name in enumerate(employee_cols, 1):[cite: 1]
            cell = ws_emp.cell(row=3, column=col_idx, value=h_name)[cite: 1]
            cell.fill = header_fill[cite: 1]
            cell.font = font_header[cite: 1]
            cell.alignment = align_center[cite: 1]

        start_row = 4[cite: 1]
        for r_offset, (_, row_data) in enumerate(emp_group.iterrows()):[cite: 1]
            curr_row = start_row + r_offset[cite: 1]
            status_val = str(row_data["Status / Alert"])[cite: 1]
            is_half_day = "(0.5 Day)" in status_val[cite: 1]
            is_sunday = status_val == "Sunday"[cite: 1]
            is_offday = (
                not is_half_day
                and any(
                    cat in status_val
                    for cat in [
                        "Holiday",
                        "Public Holiday",
                        "Annual Leave",
                        "Replacement Leave",
                        "MC",
                        "Unpaid Leave",
                        "Team A off day",
                        "Team B off day",
                    ]
                )
                and (row_data["Clock In"] == "--" and row_data["Clock Out"] == "--")
            )[cite: 1]

            if is_sunday or is_offday:[cite: 1]
                banner_text = "Sunday" if is_sunday else status_val.upper()[cite: 1]
                date_cell = ws_emp.cell(row=curr_row, column=1, value=row_data["Date"])[cite: 1]
                date_cell.font = font_bold[cite: 1]
                date_cell.border = thin_border[cite: 1]
                date_cell.alignment = align_center[cite: 1]

                for c_idx in range(2, len(employee_cols) + 1):[cite: 1]
                    c = ws_emp.cell(row=curr_row, column=c_idx)[cite: 1]
                    c.border = thin_border[cite: 1]
                    c.fill = offday_fill[cite: 1]

                ws_emp.merge_cells(
                    start_row=curr_row,
                    start_column=2,
                    end_row=curr_row,
                    end_column=len(employee_cols),
                )[cite: 1]
                merged_cell = ws_emp.cell(row=curr_row, column=2, value=banner_text)[cite: 1]
                merged_cell.alignment = align_center[cite: 1]
                merged_cell.font = font_merged_banner[cite: 1]
                continue[cite: 1]

            row_vals = [
                row_data["Date"],
                row_data["Clock In"],
                row_data["Break Out (Lunch)"],
                row_data["Break In (Back)"],
                row_data["Clock Out"],
                row_data["Lunch Duration"],
                row_data["Late to Work"],
                row_data["Late Time (Lunch)"],
                row_data["Total Late Time"],
                row_data["Early Leave"],
                row_data["Work Hours"],
                row_data["Status / Alert"],
                row_data["Multiple Punch (Earlier)"],
                row_data["Multiple Punch (Later)"],
            ][cite: 1]

            for col_idx, val in enumerate(row_vals, 1):[cite: 1]
                cell = ws_emp.cell(row=curr_row, column=col_idx, value=val)[cite: 1]
                cell.font = font_regular[cite: 1]
                cell.border = thin_border[cite: 1]
                cell.alignment = align_center if col_idx != 12 else align_left[cite: 1]

                # Highlight Lunch Duration in plain RED text only (no background fill) if > 1:10 (70m)
                if col_idx == 6 and row_data["_lunch_mins"] > 70:
                    cell.font = Font(name="Calibri", size=11, bold=True, color="B91C1C")

                if col_idx in [7, 8, 9, 10] and val != "--":[cite: 1]
                    cell.fill = late_fill[cite: 1]
                    cell.font = font_bold[cite: 1]
                if col_idx == 12 and (
                    "Missing" in val or "No Lunch" in val or "Early" in val
                ):[cite: 1]
                    cell.fill = alert_fill[cite: 1]
                    cell.font = font_bold[cite: 1]
                if col_idx in [13, 14] and val != "--":[cite: 1]
                    cell.fill = multi_fill[cite: 1]

            # Merge columns 3 & 4 (Break Out and Break In) into the soft amber block for half-day leaves
            if is_half_day:[cite: 1]
                ws_emp.merge_cells(
                    start_row=curr_row,
                    start_column=3,
                    end_row=curr_row,
                    end_column=4,
                )[cite: 1]
                half_day_label = status_val.split(" (Missing")[0].strip()[cite: 1]
                hd_cell = ws_emp.cell(row=curr_row, column=3, value=f"🌤 {half_day_label}")[cite: 1]
                hd_cell.alignment = align_center[cite: 1]
                hd_cell.font = Font(name="Calibri", size=10, bold=True, color="92400E")[cite: 1]
                hd_cell.fill = PatternFill(
                    start_color="FEF3C7", end_color="FEF3C7", fill_type="solid"
                )[cite: 1]
                ws_emp.cell(row=curr_row, column=4).border = thin_border[cite: 1]

        tot_row = start_row + len(emp_group)[cite: 1]
        tot_lunch = format_mins_to_time(emp_group["_lunch_mins"].sum())[cite: 1]
        tot_late_w = format_mins_to_time(emp_group["_late_work_mins"].sum())[cite: 1]
        tot_late_l = format_mins_to_time(emp_group["_late_lunch_mins"].sum())[cite: 1]
        tot_late_all = format_mins_to_time(emp_group["_total_late_mins"].sum())[cite: 1]
        tot_early = format_mins_to_time(emp_group["_early_leave_mins"].sum())[cite: 1]
        tot_work = format_mins_to_time(emp_group["_work_mins"].sum())[cite: 1]

        ws_emp.cell(row=tot_row, column=1, value="MONTHLY TOTAL").font = font_bold[cite: 1]
        ws_emp.cell(row=tot_row, column=1).alignment = align_center[cite: 1]
        ws_emp.cell(row=tot_row, column=6, value=tot_lunch).alignment = align_center[cite: 1]
        ws_emp.cell(row=tot_row, column=7, value=tot_late_w).alignment = align_center[cite: 1]
        ws_emp.cell(row=tot_row, column=8, value=tot_late_l).alignment = align_center[cite: 1]
        ws_emp.cell(row=tot_row, column=9, value=tot_late_all).alignment = align_center[cite: 1]
        ws_emp.cell(row=tot_row, column=10, value=tot_early).alignment = align_center[cite: 1]
        ws_emp.cell(row=tot_row, column=11, value=tot_work).alignment = align_center[cite: 1]

        for col_idx in range(1, len(employee_cols) + 1):[cite: 1]
            c = ws_emp.cell(row=tot_row, column=col_idx)[cite: 1]
            c.fill = total_fill[cite: 1]
            c.font = font_bold[cite: 1]
            c.border = thin_border[cite: 1]

        for col in ws_emp.columns:[cite: 1]
            max_len = max(len(str(cell.value or "")) for cell in col)[cite: 1]
            col_letter = get_column_letter(col[0].column)[cite: 1]
            ws_emp.column_dimensions[col_letter].width = max(max_len + 3, 14)[cite: 1]

    output = io.BytesIO()[cite: 1]
    wb.save(output)[cite: 1]
    output.seek(0)[cite: 1]
    return output[cite: 1]


# --- Authentication Endpoints ---
@app.route("/login", methods=["GET", "POST"])
def login():
    error = None[cite: 1]
    if request.method == "POST":[cite: 1]
        username = request.form.get("username", "").strip()[cite: 1]
        password = request.form.get("password", "")[cite: 1]

        if username in USERS and check_password_hash(USERS[username], password):[cite: 1]
            session["user"] = username[cite: 1]
            return redirect(url_for("index"))[cite: 1]
        else:
            error = "Invalid username or password"[cite: 1]

    return render_template("login.html", error=error)[cite: 1]


@app.route("/logout")
def logout():
    session.pop("user", None)[cite: 1]
    return redirect(url_for("login"))[cite: 1]


# --- Application Endpoints ---
@app.route("/")
@login_required
def index():
    return render_template("index.html", current_user=session.get("user"))[cite: 1]


@app.route("/api/process", methods=["POST"])
@login_required
def api_process():
    file = request.files.get("file")[cite: 1]
    if not file:[cite: 1]
        return jsonify({"error": "No file uploaded"}), 400[cite: 1]

    start_date = request.form.get("start_date")[cite: 1]
    end_date = request.form.get("end_date")[cite: 1]
    special_entries_raw = request.form.get("special_entries", "[]")[cite: 1]

    try:
        special_entries = json.loads(special_entries_raw)[cite: 1]
    except Exception:
        special_entries = [][cite: 1]

    try:
        raw_df = pd.read_excel(file, header=None)[cite: 1]
        df_processed, detected_start, detected_end = process_time_card(
            raw_df, start_date, end_date, special_entries
        )[cite: 1]

        current_user = session["user"][cite: 1]
        USER_DATAFRAMES[current_user] = df_processed[cite: 1]

        return jsonify(
            {
                "records": df_processed.to_dict(orient="records"),
                "detected_start": detected_start,
                "detected_end": detected_end,
            }
        )[cite: 1]
    except Exception as e:
        return jsonify({"error": str(e)}), 500[cite: 1]


@app.route("/api/update_records", methods=["POST"])
@login_required
def api_update_records():
    data = request.get_json()[cite: 1]
    if data and "records" in data:[cite: 1]
        current_user = session["user"][cite: 1]
        USER_DATAFRAMES[current_user] = pd.DataFrame(data["records"])[cite: 1]
        return jsonify({"status": "success"})[cite: 1]
    return jsonify({"error": "No data received"}), 400[cite: 1]


# --- Save Slot Endpoints (Postgres Linked) ---
@app.route("/api/slots", methods=["GET"])
@login_required
def api_get_slots():
    slots = load_user_slots(session["user"])[cite: 1]
    slot_headers = [
        {
            "id": s["id"],
            "title": s["title"],
            "updatedAt": s["updatedAt"],
            "startDate": s.get("startDate", ""),
            "endDate": s.get("endDate", ""),
            "totalEmployees": s.get("totalEmployees", 0),
        }
        for s in slots
    ][cite: 1]
    return jsonify({"slots": slot_headers})[cite: 1]


@app.route("/api/slots/save", methods=["POST"])
@login_required
def api_save_slot():
    data = request.get_json()[cite: 1]
    if not data or "records" not in data:[cite: 1]
        return jsonify({"error": "No data provided"}), 400[cite: 1]

    username = session["user"][cite: 1]
    slot_id = data.get("slotId")[cite: 1]
    myt_now = datetime.now(ZoneInfo("Asia/Kuala_Lumpur"))[cite: 1]
    title = (
        data.get("title", "").strip()
        or f"Save {myt_now.strftime('%d/%m/%Y %H:%M')}"
    )[cite: 1]
    unique_emps = len(set(r.get("Name") for r in data["records"]))[cite: 1]

    new_slot_entry = {
        "id": slot_id if slot_id else f"slot_{int(myt_now.timestamp() * 1000)}",
        "title": title,
        "updatedAt": myt_now.strftime("%Y-%m-%d %H:%M:%S"),
        "startDate": data.get("startDate", ""),
        "endDate": data.get("endDate", ""),
        "specialEntries": data.get("specialEntries", []),
        "totalEmployees": unique_emps,
        "records": data["records"],
    }[cite: 1]

    save_user_slots(username, new_slot_entry)[cite: 1]
    USER_DATAFRAMES[username] = pd.DataFrame(data["records"])[cite: 1]
    return jsonify({"status": "success", "slotId": new_slot_entry["id"]})[cite: 1]


@app.route("/api/slots/load/<slot_id>", methods=["GET"])
@login_required
def api_load_slot(slot_id):
    username = session["user"][cite: 1]
    slots = load_user_slots(username)[cite: 1]
    for s in slots:[cite: 1]
        if s["id"] == slot_id:[cite: 1]
            USER_DATAFRAMES[username] = pd.DataFrame(s["records"])[cite: 1]
            return jsonify({"status": "success", "slot": s})[cite: 1]
    return jsonify({"error": "Slot not found"}), 404[cite: 1]


@app.route("/api/slots/delete/<slot_id>", methods=["POST"])
@login_required
def api_delete_slot(slot_id):
    delete_user_slot(session["user"], slot_id)[cite: 1]
    return jsonify({"status": "success"})[cite: 1]


# --- Database Holiday Endpoints ---
@app.route("/api/holidays", methods=["GET"])
@login_required
def api_get_holidays():
    holidays_dict = load_db_holidays()
    holiday_list = [{"date": d, "name": n} for d, n in holidays_dict.items()]
    return jsonify({"holidays": holiday_list})


@app.route("/api/holidays/save", methods=["POST"])
@login_required
def api_save_holiday():
    data = request.get_json() or {}
    h_date = data.get("date", "").strip()
    h_name = data.get("name", "").strip()
    if not h_date or not h_name:
        return jsonify({"error": "Date and Holiday Name are required"}), 400
    
    try:
        dt = datetime.strptime(h_date, "%Y-%m-%d").date()
        if dt.weekday() == 6:
            return jsonify({"error": "Cannot assign a holiday on Sunday"}), 400
    except Exception:
        return jsonify({"error": "Invalid date format"}), 400

    if save_db_holiday(h_date, h_name):
        return jsonify({"status": "success"})
    return jsonify({"error": "Failed to save holiday to database"}), 500


@app.route("/api/holidays/delete", methods=["POST"])
@login_required
def api_delete_holiday():
    data = request.get_json() or {}
    h_date = data.get("date", "").strip()
    if not h_date:
        return jsonify({"error": "Date is required"}), 400
    if delete_db_holiday(h_date):
        return jsonify({"status": "success"})
    return jsonify({"error": "Failed to delete holiday"}), 500


@app.route("/api/download", methods=["GET"])
@login_required
def api_download():
    current_user = session["user"][cite: 1]
    user_df = USER_DATAFRAMES.get(current_user)[cite: 1]
    if user_df is None:[cite: 1]
        return "No processed data available for this user session", 400[cite: 1]

    excel_file = build_excel_workbook(user_df)[cite: 1]
    return send_file(
        excel_file,
        as_attachment=True,
        download_name=f"Monthly_Attendance_Summary_{current_user}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )[cite: 1]


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))[cite: 1]
    app.run(host="0.0.0.0", port=port, debug=False)[cite: 1]