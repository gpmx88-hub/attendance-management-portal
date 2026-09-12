from flask import Flask, render_template, request, jsonify, send_file
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
import holidays

app = Flask(__name__)
CURRENT_DF = None

# Business Rules
WORK_START_TIME = time(9, 0, 0)         # Mon - Sat start: 09:00:00
WEEKDAY_END_TIME = time(18, 0, 0)       # Mon - Fri finish: 18:00:00
SATURDAY_END_TIME = time(13, 30, 0)     # Saturday finish: 13:30:00
LUNCH_LIMIT_MINS = 70                   # Lunch limit: 1 hour 10 mins

# Malaysia Public Holidays Engine (English-first)
MY_HOLIDAYS = holidays.country_holidays('MY', language='en')

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
}

def get_malaysia_holiday_name(d_obj):
    raw_name = MY_HOLIDAYS.get(d_obj)
    if not raw_name:
        return None
    for my_name, en_name in HOLIDAY_EN_MAP.items():
        if my_name.lower() in raw_name.lower():
            return en_name
    return raw_name.replace(" (Observed)", "").replace(" (observed)", "").strip()

def parse_time_str(t_str):
    try:
        return datetime.strptime(t_str.strip(), "%H:%M:%S")
    except Exception:
        return None

def format_mins_to_time(minutes):
    if pd.isna(minutes) or minutes <= 0:
        return "0:00"
    hrs = int(minutes // 60)
    mins = int(round(minutes % 60))
    if mins == 60:
        hrs += 1
        mins = 0
    return f"{hrs}:{mins:02d}"

def deduplicate_close_punches(punch_list, threshold_minutes=3):
    valid_punches = []
    close_duplicates = []

    for p in punch_list:
        dt = parse_time_str(p)
        if not dt:
            continue
        if not valid_punches:
            valid_punches.append((dt, p))
            continue

        prev_dt, _ = valid_punches[-1]
        diff_sec = (dt - prev_dt).total_seconds()
        if diff_sec <= (threshold_minutes * 60):
            close_duplicates.append(p)
        else:
            valid_punches.append((dt, p))

    return [p for _, p in valid_punches], close_duplicates

def categorize_punches(raw_punch_list, is_saturday):
    punch_list, close_duplicates = deduplicate_close_punches(raw_punch_list, threshold_minutes=3)
    p_tuples = []
    for p in punch_list:
        dt = parse_time_str(p)
        if dt:
            p_tuples.append((dt.time(), p))

    c_in, b_out, b_in, c_out = "--", "--", "--", "--"
    assigned_punches = []

    if is_saturday:
        morning = [p for p in p_tuples if p[0].hour < 11 or (p[0].hour == 11 and p[0].minute < 30)]
        afternoon = [p for p in p_tuples if p[0].hour > 11 or (p[0].hour == 11 and p[0].minute >= 30)]
        if morning:
            c_in = morning[0][1]
            assigned_punches.append(c_in)
        if afternoon:
            c_out = afternoon[-1][1]
            assigned_punches.append(c_out)
        elif len(morning) > 1:
            c_out = morning[-1][1]
            assigned_punches.append(c_out)
    else:
        morning = [p for p in p_tuples if p[0].hour < 11 or (p[0].hour == 11 and p[0].minute < 15)]
        evening = [p for p in p_tuples if p[0].hour > 15 or (p[0].hour == 15 and p[0].minute > 30)]

        if morning:
            c_in = morning[0][1]
            assigned_punches.append(c_in)

        lunch_out_candidates = [p for p in p_tuples if (p[0].hour == 11 and p[0].minute >= 15) or (p[0].hour == 12) or (p[0].hour == 13 and p[0].minute < 15)]
        lunch_in_candidates = [p for p in p_tuples if (p[0].hour == 13 and p[0].minute >= 15) or (p[0].hour == 14) or (p[0].hour == 15 and p[0].minute <= 30)]

        if lunch_out_candidates:
            b_out = lunch_out_candidates[-1][1]
            assigned_punches.append(b_out)
        if lunch_in_candidates:
            b_in = lunch_in_candidates[0][1]
            assigned_punches.append(b_in)

        if b_out == "--" and b_in == "--":
            all_lunch = [p for p in p_tuples if (p[0].hour > 11 or (p[0].hour == 11 and p[0].minute >= 15)) and (p[0].hour < 15 or (p[0].hour == 15 and p[0].minute <= 30))]
            if len(all_lunch) >= 2:
                b_out = all_lunch[0][1]
                b_in = all_lunch[-1][1]
                assigned_punches.extend([b_out, b_in])
            elif len(all_lunch) == 1:
                if all_lunch[0][0].hour < 13:
                    b_out = all_lunch[0][1]
                else:
                    b_in = all_lunch[0][1]
                assigned_punches.append(all_lunch[0][1])

        if evening:
            c_out = evening[-1][1]
            assigned_punches.append(c_out)

    extra_punches = [p for p in punch_list if p not in assigned_punches] + close_duplicates
    multi_earlier = "--"
    multi_later = "--"
    if len(extra_punches) == 1:
        multi_earlier = extra_punches[0]
    elif len(extra_punches) >= 2:
        multi_earlier = extra_punches[0]
        multi_later = extra_punches[-1]

    issues = []
    if is_saturday:
        if c_in == "--": issues.append("Missing Clock In")
        if c_out == "--": issues.append("Missing Clock Out")
        if len(punch_list) > 2: issues.append(f"Multiple Punches ({len(punch_list)})")
        status = "Saturday (Half Day)" if not issues else f"Saturday ({', '.join(issues)})"
    else:
        if c_in == "--": issues.append("Missing Clock In")
        if b_out == "--" and b_in != "--": issues.append("Missing Break Out")
        if b_out != "--" and b_in == "--": issues.append("Missing Break In")
        if b_out == "--" and b_in == "--": issues.append("No Lunch Punched")
        if c_out == "--": issues.append("Missing Clock Out")
        if len(punch_list) > 4: issues.append(f"Multiple Punches ({len(punch_list)})")
        status = "Normal" if not issues else ", ".join(issues)

    return c_in, b_out, b_in, c_out, status, multi_earlier, multi_later

def process_time_card(df_raw, start_date_str=None, end_date_str=None, special_entries=None):
    if special_entries is None:
        special_entries = []

    special_lookup = {}
    for entry in special_entries:
        special_lookup[(entry['date'], entry['target'])] = entry['type']

    header_idx = None
    for idx, row in df_raw.iloc[:5].iterrows():
        if "Employee ID" in row.values:
            header_idx = idx
            break

    if header_idx is not None:
        df = df_raw.iloc[header_idx + 1:].copy()
        df.columns = df_raw.iloc[header_idx].values
    else:
        df = df_raw.copy()

    df.columns = [str(c).strip() for c in df.columns]

    raw_punches = {}
    emp_meta = {}
    dates_in_file = []

    for _, row in df.iterrows():
        emp_id = str(row.get("Employee ID", "")).strip()
        name = str(row.get("First Name", "")).strip()
        dept = str(row.get("Department", "")).strip()
        date_str = str(row.get("Date", "")).split()[0]
        raw_times_str = str(row.get("Time", ""))

        if pd.isna(raw_times_str) or not raw_times_str.strip() or raw_times_str == "nan":
            continue

        emp_meta[emp_id] = {"name": name, "dept": dept}
        raw_punches[(emp_id, date_str)] = raw_times_str
        try:
            dates_in_file.append(datetime.strptime(date_str, "%Y-%m-%d").date())
        except Exception:
            pass

    # Full month auto-detection
    if dates_in_file:
        detected_date = dates_in_file[0]
        year = detected_date.year
        month = detected_date.month
        last_day = calendar.monthrange(year, month)[1]
        min_d = date(year, month, 1)
        max_d = date(year, month, last_day)
    else:
        today = date.today()
        last_day = calendar.monthrange(today.year, today.month)[1]
        min_d = date(today.year, today.month, 1)
        max_d = date(today.year, today.month, last_day)

    if start_date_str:
        try: min_d = datetime.strptime(start_date_str, "%Y-%m-%d").date()
        except Exception: pass
    if end_date_str:
        try: max_d = datetime.strptime(end_date_str, "%Y-%m-%d").date()
        except Exception: pass

    calendar_dates = []
    curr = min_d
    while curr <= max_d:
        calendar_dates.append(curr)
        curr += timedelta(days=1)

    records = []
    for emp_id, meta in sorted(emp_meta.items()):
        name = meta["name"]
        dept = meta["dept"]

        for w_date in calendar_dates:
            date_str = w_date.strftime("%Y-%m-%d")
            is_sunday = (w_date.weekday() == 6)
            is_saturday = (w_date.weekday() == 5)
            day_name = w_date.strftime("%a")

            special_type = special_lookup.get((date_str, name)) or special_lookup.get((date_str, "ALL"))
            if special_type == "Holiday":
                h_name = get_malaysia_holiday_name(w_date)
                special_type = f"Public Holiday ({h_name})" if h_name else "Public Holiday"

            raw_times_str = raw_punches.get((emp_id, date_str))

            if is_sunday:
                records.append({
                    "Employee ID": emp_id,
                    "Name": name,
                    "Department": dept,
                    "Date": f"{date_str} ({day_name})",
                    "Clock In": "--", "Break Out (Lunch)": "--", "Break In (Back)": "--", "Clock Out": "--",
                    "Lunch Duration": "--", "Late to Work": "--", "Late Time (Lunch)": "--",
                    "Total Late Time": "--", "Early Leave": "--", "Work Hours": "--",
                    "Status / Alert": "Sunday",
                    "Multiple Punch (Earlier)": "--", "Multiple Punch (Later)": "--",
                    "_work_mins": 0, "_lunch_mins": 0, "_late_work_mins": 0, "_late_lunch_mins": 0,
                    "_total_late_mins": 0, "_early_leave_mins": 0, "_is_absent": 0, "_is_sunday": 1, "_is_offday": 1
                })
                continue

            if not raw_times_str:
                if special_type:
                    status_text = special_type
                    is_absent_flag = 0
                    is_off = 1
                else:
                    status_text = "Absent"
                    is_absent_flag = 1
                    is_off = 0

                records.append({
                    "Employee ID": emp_id,
                    "Name": name,
                    "Department": dept,
                    "Date": f"{date_str} ({day_name})",
                    "Clock In": "--", "Break Out (Lunch)": "--", "Break In (Back)": "--", "Clock Out": "--",
                    "Lunch Duration": "--", "Late to Work": "--", "Late Time (Lunch)": "--",
                    "Total Late Time": "--", "Early Leave": "--",
                    "Work Hours": "0:00" if is_absent_flag else "--",
                    "Status / Alert": status_text,
                    "Multiple Punch (Earlier)": "--", "Multiple Punch (Later)": "--",
                    "_work_mins": 0, "_lunch_mins": 0, "_late_work_mins": 0, "_late_lunch_mins": 0,
                    "_total_late_mins": 0, "_early_leave_mins": 0, "_is_absent": is_absent_flag,
                    "_is_sunday": 0, "_is_offday": is_off
                })
                continue

            punch_list = [t.strip() for t in raw_times_str.split(",") if t.strip()]
            c_in, b_out, b_in, c_out, status, multi_earlier, multi_later = categorize_punches(punch_list, is_saturday)

            if special_type:
                status = f"{special_type} (Worked)" if status == "Normal" else f"{special_type} ({status})"

            lunch_mins, work_mins, late_work_mins, late_lunch_mins, early_leave_mins = 0, 0, 0, 0, 0

            if c_in != "--":
                dt_cin = parse_time_str(c_in)
                if dt_cin:
                    start_dt = dt_cin.replace(hour=WORK_START_TIME.hour, minute=WORK_START_TIME.minute, second=WORK_START_TIME.second)
                    if dt_cin > start_dt:
                        late_work_mins = round((dt_cin - start_dt).total_seconds() / 60)

            if not is_saturday and b_out != "--" and b_in != "--":
                dt_bout = parse_time_str(b_out)
                dt_bin = parse_time_str(b_in)
                if dt_bout and dt_bin and dt_bin > dt_bout:
                    lunch_mins = round((dt_bin - dt_bout).total_seconds() / 60)
                    if lunch_mins > LUNCH_LIMIT_MINS:
                        late_lunch_mins = lunch_mins - LUNCH_LIMIT_MINS

            early_remark = ""
            if c_out != "--":
                dt_cout = parse_time_str(c_out)
                if dt_cout:
                    target_end = SATURDAY_END_TIME if is_saturday else WEEKDAY_END_TIME
                    end_dt = dt_cout.replace(hour=target_end.hour, minute=target_end.minute, second=target_end.second)
                    if dt_cout < end_dt:
                        early_leave_mins = round((end_dt - dt_cout).total_seconds() / 60)
                        if early_leave_mins > 0:
                            half_hour_blocks = math.ceil(early_leave_mins / 30)
                            if half_hour_blocks == 1:
                                early_remark = "Early up 30 mins"
                            else:
                                total_hours = half_hour_blocks * 0.5
                                hr_str = f"{int(total_hours)}" if total_hours.is_integer() else f"{total_hours}"
                                hr_label = "hour" if total_hours == 1.0 else "hours"
                                early_remark = f"Early up {hr_str} {hr_label}"

            if early_remark:
                if status in ["Normal", "Saturday (Half Day)"] or any(k in status for k in ["off day", "Leave", "Holiday", "MC"]):
                    status = f"{status} ({early_remark})"
                else:
                    status = f"{status}, {early_remark}"

            total_late_mins = late_work_mins + late_lunch_mins

            if c_in != "--" and c_out != "--":
                dt_cin = parse_time_str(c_in)
                dt_cout = parse_time_str(c_out)
                if dt_cin and dt_cout and dt_cout > dt_cin:
                    gross_mins = round((dt_cout - dt_cin).total_seconds() / 60)
                    work_mins = max(0, gross_mins - lunch_mins)

            records.append({
                "Employee ID": emp_id,
                "Name": name,
                "Department": dept,
                "Date": f"{date_str} ({day_name})",
                "Clock In": c_in,
                "Break Out (Lunch)": b_out,
                "Break In (Back)": b_in,
                "Clock Out": c_out,
                "Lunch Duration": format_mins_to_time(lunch_mins) if lunch_mins > 0 else "--",
                "Late to Work": format_mins_to_time(late_work_mins) if late_work_mins > 0 else "--",
                "Late Time (Lunch)": format_mins_to_time(late_lunch_mins) if late_lunch_mins > 0 else "--",
                "Total Late Time": format_mins_to_time(total_late_mins) if total_late_mins > 0 else "--",
                "Early Leave": format_mins_to_time(early_leave_mins) if early_leave_mins > 0 else "--",
                "Work Hours": format_mins_to_time(work_mins) if work_mins > 0 else "--",
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
                "_is_offday": 1 if special_type else 0
            })

    return pd.DataFrame(records), min_d.strftime("%Y-%m-%d"), max_d.strftime("%Y-%m-%d")

def build_excel_workbook(df):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    header_fill = PatternFill(start_color="0F2D59", end_color="0F2D59", fill_type="solid")
    total_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    offday_fill = PatternFill(start_color="FFF5F5", end_color="FFF5F5", fill_type="solid")
    absent_fill = PatternFill(start_color="FCA5A5", end_color="FCA5A5", fill_type="solid")
    late_fill = PatternFill(start_color="FEF08A", end_color="FEF08A", fill_type="solid")
    alert_fill = PatternFill(start_color="FED7AA", end_color="FED7AA", fill_type="solid")
    multi_fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")

    font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    font_bold = Font(name="Calibri", size=11, bold=True)
    font_regular = Font(name="Calibri", size=11)
    font_merged_banner = Font(name="Calibri", size=11, bold=True, color="DC2626")

    thin_border = Border(
        left=Side(style="thin", color="CBD5E1"),
        right=Side(style="thin", color="CBD5E1"),
        top=Side(style="thin", color="CBD5E1"),
        bottom=Side(style="thin", color="CBD5E1"),
    )
    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")

    # Overview Summary
    ws_summary = wb.create_sheet(title="Overview Summary")
    ws_summary.views.sheetView[0].showGridLines = True

    summary_headers = [
        "Employee ID", "Employee Name", "Department",
        "Total Days Worked", "Total Absent Days", "Total Work Hours", "Total Lunch Time", 
        "Total Late (Work)", "Total Late (Lunch)", "Total Late (Sum)", "Total Early Leave", "Issues Count"
    ]
    ws_summary.append(summary_headers)

    for col_idx in range(1, len(summary_headers) + 1):
        cell = ws_summary.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = font_header
        cell.alignment = align_center

    non_work_categories = ["Holiday", "Public Holiday", "Annual Leave", "MC", "Team A off day", "Team B off day", "Sunday"]

    overview_data = []
    for (emp_id, emp_name), emp_group in df.groupby(["Employee ID", "Name"], sort=False):
        dept = emp_group["Department"].iloc[0]
        days_present = len(emp_group[(emp_group["_is_absent"] == 0) & (~emp_group["Status / Alert"].str.contains('|'.join(non_work_categories)))])
        days_absent = len(emp_group[emp_group["_is_absent"] == 1])
        total_work_m = emp_group["_work_mins"].sum()
        total_lunch_m = emp_group["_lunch_mins"].sum()
        total_late_w_m = emp_group["_late_work_mins"].sum()
        total_late_l_m = emp_group["_late_lunch_mins"].sum()
        grand_total_late_m = emp_group["_total_late_mins"].sum()
        total_early_m = emp_group["_early_leave_mins"].sum()
        anomalies = len(emp_group[
            ~emp_group["Status / Alert"].str.startswith("Normal") & 
            ~emp_group["Status / Alert"].str.startswith("Saturday (Half Day)") & 
            (emp_group["Status / Alert"] != "Absent") &
            (~emp_group["Status / Alert"].str.contains('|'.join(non_work_categories)))
        ])
        
        overview_data.append([
            emp_id, emp_name, dept, days_present, days_absent,
            format_mins_to_time(total_work_m),
            format_mins_to_time(total_lunch_m),
            format_mins_to_time(total_late_w_m),
            format_mins_to_time(total_late_l_m),
            format_mins_to_time(grand_total_late_m),
            format_mins_to_time(total_early_m),
            anomalies
        ])

    for row_idx, row_vals in enumerate(overview_data, start=2):
        for col_idx, val in enumerate(row_vals, start=1):
            cell = ws_summary.cell(row=row_idx, column=col_idx, value=val)
            cell.font = font_regular
            cell.border = thin_border
            cell.alignment = align_center if col_idx not in [2, 3] else align_left

            if col_idx == 5 and val > 0:
                cell.fill = absent_fill
                cell.font = font_bold
            elif col_idx in [8, 9, 10, 11] and str(val) != "0:00":
                cell.fill = late_fill
                cell.font = font_bold
            elif col_idx == 12 and val > 0:
                cell.fill = alert_fill
                cell.font = font_bold

    for col in ws_summary.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws_summary.column_dimensions[col_letter].width = max(max_len + 4, 14)

    # Individual Employee Sheets
    employee_cols = [
        "Date", "Clock In", "Break Out (Lunch)", "Break In (Back)",
        "Clock Out", "Lunch Duration", "Late to Work", "Late Time (Lunch)", 
        "Total Late Time", "Early Leave", "Work Hours", "Status / Alert",
        "Multiple Punch (Earlier)", "Multiple Punch (Later)"
    ]

    for (emp_id, emp_name), emp_group in df.groupby(["Employee ID", "Name"], sort=False):
        safe_name = re.sub(r"[\/\\\?\*\:\[\]]", "", emp_name).strip()
        sheet_title = safe_name[:31] if safe_name else f"Emp_{emp_id}"

        counter = 1
        orig_title = sheet_title
        while sheet_title in wb.sheetnames:
            sheet_title = f"{orig_title[:28]}_{counter}"
            counter += 1

        ws_emp = wb.create_sheet(title=sheet_title)
        ws_emp.views.sheetView[0].showGridLines = True

        ws_emp["A1"] = f"Employee ID: {emp_id}"
        ws_emp["A1"].font = font_bold
        ws_emp["C1"] = f"Name: {emp_name}"
        ws_emp["C1"].font = font_bold
        ws_emp["G1"] = f"Department: {emp_group['Department'].iloc[0]}"
        ws_emp["G1"].font = font_bold

        for col_idx, h_name in enumerate(employee_cols, 1):
            cell = ws_emp.cell(row=3, column=col_idx, value=h_name)
            cell.fill = header_fill
            cell.font = font_header
            cell.alignment = align_center

        start_row = 4
        for r_offset, (_, row_data) in enumerate(emp_group.iterrows()):
            curr_row = start_row + r_offset
            status_val = str(row_data["Status / Alert"])
            is_sunday = (status_val == "Sunday")
            is_offday = any(cat in status_val for cat in ["Holiday", "Public Holiday", "Annual Leave", "MC", "Team A off day", "Team B off day"]) and (row_data["Clock In"] == "--" and row_data["Clock Out"] == "--")

            if is_sunday or is_offday:
                banner_text = "Sunday" if is_sunday else status_val.upper()
                date_cell = ws_emp.cell(row=curr_row, column=1, value=row_data["Date"])
                date_cell.font = font_bold
                date_cell.border = thin_border
                date_cell.alignment = align_center

                for c_idx in range(2, len(employee_cols) + 1):
                    c = ws_emp.cell(row=curr_row, column=c_idx)
                    c.border = thin_border
                    c.fill = offday_fill

                ws_emp.merge_cells(start_row=curr_row, start_column=2, end_row=curr_row, end_column=len(employee_cols))
                merged_cell = ws_emp.cell(row=curr_row, column=2, value=banner_text)
                merged_cell.alignment = align_center
                merged_cell.font = font_merged_banner
                continue

            row_vals = [
                row_data["Date"], row_data["Clock In"], row_data["Break Out (Lunch)"], row_data["Break In (Back)"],
                row_data["Clock Out"], row_data["Lunch Duration"], row_data["Late to Work"], row_data["Late Time (Lunch)"],
                row_data["Total Late Time"], row_data["Early Leave"], row_data["Work Hours"], row_data["Status / Alert"],
                row_data["Multiple Punch (Earlier)"], row_data["Multiple Punch (Later)"]
            ]
            is_absent = (status_val == "Absent")

            for col_idx, val in enumerate(row_vals, 1):
                cell = ws_emp.cell(row=curr_row, column=col_idx, value=val)
                cell.font = font_regular
                cell.border = thin_border
                cell.alignment = align_center if col_idx != 12 else align_left

                if is_absent:
                    cell.fill = absent_fill
                    if col_idx == 12: cell.font = font_bold
                else:
                    if col_idx in [7, 8, 9, 10] and val != "--":
                        cell.fill = late_fill
                        cell.font = font_bold
                    if col_idx == 12 and ("Missing" in val or "No Lunch" in val or "Early" in val):
                        cell.fill = alert_fill
                        cell.font = font_bold
                    if col_idx in [13, 14] and val != "--":
                        cell.fill = multi_fill

        tot_row = start_row + len(emp_group)
        tot_lunch = format_mins_to_time(emp_group["_lunch_mins"].sum())
        tot_late_w = format_mins_to_time(emp_group["_late_work_mins"].sum())
        tot_late_l = format_mins_to_time(emp_group["_late_lunch_mins"].sum())
        tot_late_all = format_mins_to_time(emp_group["_total_late_mins"].sum())
        tot_early = format_mins_to_time(emp_group["_early_leave_mins"].sum())
        tot_work = format_mins_to_time(emp_group["_work_mins"].sum())

        ws_emp.cell(row=tot_row, column=1, value="MONTHLY TOTAL").font = font_bold
        ws_emp.cell(row=tot_row, column=1).alignment = align_center
        ws_emp.cell(row=tot_row, column=6, value=tot_lunch).alignment = align_center
        ws_emp.cell(row=tot_row, column=7, value=tot_late_w).alignment = align_center
        ws_emp.cell(row=tot_row, column=8, value=tot_late_l).alignment = align_center
        ws_emp.cell(row=tot_row, column=9, value=tot_late_all).alignment = align_center
        ws_emp.cell(row=tot_row, column=10, value=tot_early).alignment = align_center
        ws_emp.cell(row=tot_row, column=11, value=tot_work).alignment = align_center

        for col_idx in range(1, len(employee_cols) + 1):
            c = ws_emp.cell(row=tot_row, column=col_idx)
            c.fill = total_fill
            c.font = font_bold
            c.border = thin_border

        for col in ws_emp.columns:
            max_len = max(len(str(cell.value or "")) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws_emp.column_dimensions[col_letter].width = max(max_len + 3, 14)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/api/process", methods=["POST"])
def api_process():
    global CURRENT_DF
    file = request.files.get("file")
    if not file:
        return jsonify({"error": "No file uploaded"}), 400

    start_date = request.form.get("start_date")
    end_date = request.form.get("end_date")
    special_entries_raw = request.form.get("special_entries", "[]")

    try:
        special_entries = json.loads(special_entries_raw)
    except Exception:
        special_entries = []

    try:
        raw_df = pd.read_excel(file, header=None)
        CURRENT_DF, detected_start, detected_end = process_time_card(raw_df, start_date, end_date, special_entries)
        return jsonify({
            "records": CURRENT_DF.to_dict(orient="records"),
            "detected_start": detected_start,
            "detected_end": detected_end
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/api/update_records", methods=["POST"])
def api_update_records():
    global CURRENT_DF
    data = request.get_json()
    if data and "records" in data:
        CURRENT_DF = pd.DataFrame(data["records"])
        return jsonify({"status": "success"})
    return jsonify({"error": "No data received"}), 400

@app.route("/api/download", methods=["GET"])
def api_download():
    global CURRENT_DF
    if CURRENT_DF is None:
        return "No processed data available", 400

    excel_file = build_excel_workbook(CURRENT_DF)
    return send_file(
        excel_file,
        as_attachment=True,
        download_name="Monthly_Attendance_Summary.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)