
import io
import re
from dataclasses import dataclass
from datetime import datetime, date, time, timedelta
from typing import Dict, List, Tuple, Optional, Set

import math
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import FormulaRule, DataBarRule

# =========================================================
# Rota Generator Engine — Clean Rebuild v24.1 (template-locked)
# - Uses the v14+++ scheduling core that enforces Phones/Bookings/EMIS/Docman targets
# - Outputs SITE timelines with COLOUR (conditional formatting) + Coverage by slot by site
# - Totals sheet uses formulas linked to SITE timelines so edits update totals
# =========================================================

DAY_START = time(8, 0)
DAY_END   = time(18, 30)
SLOT_MIN  = 30

SITES = ["SLGP", "JEN", "BGS"]

FD_BANDS = [
    (time(8, 0),  time(11, 0)),
    (time(11, 0), time(13, 30)),
    (time(13, 30), time(16, 0)),
    (time(16, 0), time(18, 30)),
]
TRIAGE_BANDS = [
    (time(8, 0),  time(10, 30)),
    (time(10, 30), time(13, 0)),
    (time(13, 30), time(16, 0)),
]

BREAK_WINDOW = (time(12, 0), time(14, 0))
BREAK_CANDIDATES = [time(12, 0), time(12, 30), time(13, 0), time(13, 30)]
BREAK_THRESHOLD_HOURS = 6.0

# Block mins (slots)
MIN_PHONES = 3          # 1.5h
MAX_PHONES = 8          # 4h
MIN_DEFAULT = 5         # 2.5h
MAX_DEFAULT = 9         # 4.5h
MIN_DOCMAN = 4          # 2h (hard rule)
MIN_EMIS   = 4          # 2h (hard rule)
MAX_EMIS_DOCMAN = 8     # 4h (hard rule)

THICK = Side(style="thick")
THIN  = Side(style="thin")
CELL_BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
DAY_TOP = Border(top=THICK, left=THIN, right=THIN, bottom=THIN)

ROLE_COLORS = {
    "FrontDesk_SLGP": "FFF2CC",
    "FrontDesk_JEN":  "FFF2CC",
    "FrontDesk_BGS":  "FFF2CC",
    "Triage_Admin_SLGP": "D9EAD3",
    "Triage_Admin_JEN":  "D9EAD3",
    "Email_Box": "CFE2F3",
    "Phones": "C9DAF8",
    "Bookings": "FCE5CD",
    "EMIS": "EAD1DC",
    "Docman": "D0E0E3",
    "Awaiting_PSA_Admin": "D0E0E3",
    "Misc_Tasks": "FFFFFF",
    "Unassigned": "FFFFFF",
    "Break": "DDDDDD",
    "Holiday": "FFF2CC",
    "Bank Holiday": "FFE599",
    "Sick": "F4CCCC",
    "": "DDDDDD",
}

def _fill_for(value: str) -> PatternFill:
    return PatternFill('solid', fgColor=ROLE_COLORS.get(value, 'FFFFFF'))



def timeslots() -> List[time]:
    cur = datetime(2000, 1, 1, DAY_START.hour, DAY_START.minute)
    end = datetime(2000, 1, 1, DAY_END.hour, DAY_END.minute)
    out: List[time] = []
    while cur < end:
        out.append(cur.time())
        cur += timedelta(minutes=SLOT_MIN)
    return out

def ensure_monday(d: date) -> date:
    return d - timedelta(days=d.weekday())

def day_name(d: date) -> str:
    return ["Mon","Tue","Wed","Thu","Fri","Sat","Sun"][d.weekday()]

def t_in_range(t: time, a: time, b: time) -> bool:
    return (t >= a) and (t < b)

def normalize(s: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(s).strip().lower())

def find_sheet(xls: pd.ExcelFile, candidates: List[str]) -> Optional[str]:
    names = {normalize(n): n for n in xls.sheet_names}
    for c in candidates:
        k = normalize(c)
        if k in names:
            return names[k]
    for n in xls.sheet_names:
        nn = normalize(n)
        for c in candidates:
            if normalize(c) in nn:
                return n
    return None

def pick_col(df: pd.DataFrame, candidates: List[str], required: bool=True) -> Optional[str]:
    cols = {normalize(c): c for c in df.columns}
    for cand in candidates:
        k = normalize(cand)
        if k in cols:
            return cols[k]
    for c in df.columns:
        nc = normalize(c)
        for cand in candidates:
            if normalize(cand) in nc:
                return c
    if required:
        raise KeyError(f"Missing required column among {candidates}. Available: {list(df.columns)}")
    return None

def to_time(x):
    if pd.isna(x) or x == "":
        return None
    if isinstance(x, time):
        return x
    if isinstance(x, datetime):
        return x.time()
    if isinstance(x, (float, int)):
        seconds = int(round(float(x) * 86400))
        return (datetime(2000,1,1) + timedelta(seconds=seconds)).time()
    return pd.to_datetime(str(x)).time()

def to_date(x):
    if pd.isna(x) or x == "":
        return None
    if isinstance(x, (datetime, pd.Timestamp)):
        return x.date()
    if isinstance(x, date):
        return x
    # IMPORTANT: your templates are UK format
    return pd.to_datetime(x, dayfirst=True).date()

def dt_of(d: date, t: time) -> datetime:
    return datetime(d.year, d.month, d.day, t.hour, t.minute)

def add_minutes(t: time, mins: int) -> time:
    return (datetime(2000,1,1,t.hour,t.minute) + timedelta(minutes=mins)).time()

def yn(v) -> bool:
    if pd.isna(v):
        return False
    s = str(v).strip().lower()
    return s in {"y","yes","true","1","t"}

@dataclass
class Staff:
    name: str
    home: str
    can_frontdesk: bool
    can_triage: bool
    can_email: bool
    can_phones: bool
    can_bookings: bool
    can_emis: bool
    can_docman: bool
    weights: Dict[str, int]
    frontdesk_only: bool
    break_required: bool

@dataclass
class TemplateData:
    staff: List[Staff]
    hours_map: Dict[str, Dict[str, Optional[time]]]
    hols: List[Tuple[str, date, date, str]]
    call_handlers: pd.DataFrame
    handler_leave: pd.DataFrame
    phones_targets: Dict[Tuple[str, time], int]
    bookings_targets: Dict[Tuple[str, time], int]
    weekly_targets: Dict[str, float]
    swaps: List[Tuple[date, str, str, Optional[time], Optional[time]]]

def read_template(uploaded_bytes: bytes) -> TemplateData:
    xls = pd.ExcelFile(io.BytesIO(uploaded_bytes))

    staff_sheet = find_sheet(xls, ["Staff"])
    hours_sheet = find_sheet(xls, ["WorkingHours", "Hours"])
    hols_sheet  = find_sheet(xls, ["Holidays", "Leave", "Absence"])
    callh_sheet = find_sheet(xls, ["CallHandlers", "Call Handlers"])
    hleave_sheet = find_sheet(xls, ["Handler_Leave", "Handler Leave", "CallHandler_Leave"])
    tph_sheet   = find_sheet(xls, ["Targets_Phones_Hourly", "PhonesTargets"])
    tbk_sheet   = find_sheet(xls, ["Targets_Bookings_Hourly", "BookingsTargets"])
    tweek_sheet = find_sheet(xls, ["Targets_Weekly"])
    swaps_sheet = find_sheet(xls, ["Swaps"])

    if not staff_sheet or not hours_sheet:
        raise ValueError(f"Missing Staff/WorkingHours sheets. Found: {xls.sheet_names}")

    staff_df = pd.read_excel(xls, sheet_name=staff_sheet)
    hours_df = pd.read_excel(xls, sheet_name=hours_sheet)
    hols_df  = pd.read_excel(xls, sheet_name=hols_sheet) if hols_sheet else pd.DataFrame()
    callh_df = pd.read_excel(xls, sheet_name=callh_sheet) if callh_sheet else pd.DataFrame()
    hleave_df = pd.read_excel(xls, sheet_name=hleave_sheet) if hleave_sheet else pd.DataFrame()
    tph_df   = pd.read_excel(xls, sheet_name=tph_sheet) if tph_sheet else pd.DataFrame()
    tbk_df   = pd.read_excel(xls, sheet_name=tbk_sheet) if tbk_sheet else pd.DataFrame()
    tweek_df = pd.read_excel(xls, sheet_name=tweek_sheet) if tweek_sheet else pd.DataFrame()
    swaps_df = pd.read_excel(xls, sheet_name=swaps_sheet) if swaps_sheet else pd.DataFrame()

    # --- Staff
    name_c = pick_col(staff_df, ["Name","StaffName"])
    home_c = pick_col(staff_df, ["HomeSite","Site","BaseSite"], required=False)

    staff_df = staff_df.copy()
    staff_df["Name"] = staff_df[name_c].astype(str).str.strip()
    staff_df["HomeSite"] = staff_df[home_c].astype(str).str.strip().str.upper() if home_c else ""

    def bcol(cands, default=False):
        c = pick_col(staff_df, cands, required=False)
        if not c:
            return pd.Series([default]*len(staff_df))
        return staff_df[c].apply(yn)

    staff_df["CanFrontDesk"] = bcol(["CanFrontDesk"])
    staff_df["CanTriage"]    = bcol(["CanTriage"])
    staff_df["CanEmail"]     = bcol(["CanEmail"])
    staff_df["CanPhones"]    = bcol(["CanPhones"])
    staff_df["CanBookings"]  = bcol(["CanBookings"])
    staff_df["CanEMIS"]      = bcol(["CanEMIS"])
    staff_df["CanDocman"]    = bcol(["CanDocman_PSA"]) | bcol(["CanDocman_AWAIT"]) | bcol(["CanDocman"])

    staff_df["FrontDeskOnly"] = bcol(["FrontDeskOnly"], default=False)
    staff_df["BreakRequired"] = bcol(["BreakRequired"], default=True)

    weight_cols = {
        "FrontDesk":"FrontDeskWeight",
        "Triage":"TriageWeight",
        "Email":"EmailWeight",
        "Phones":"PhonesWeight",
        "Bookings":"BookingsWeight",
        "EMIS":"EmisWeight",
        "Docman":"DocmanWeight",
        "Awaiting":"AwaitingWeight",
        "Misc":"MiscWeight",
    }

    staff_list: List[Staff] = []
    for _, r in staff_df.iterrows():
        weights: Dict[str,int] = {}
        for k, col in weight_cols.items():
            v = r.get(col, 3)
            try:
                if pd.isna(v) or v == "":
                    v = 3
                v = int(float(v))
            except Exception:
                v = 3
            weights[k] = max(0, min(5, v))
        staff_list.append(
            Staff(
                name=str(r["Name"]).strip(),
                home=str(r.get("HomeSite","")).strip().upper(),
                can_frontdesk=bool(r.get("CanFrontDesk", False)),
                can_triage=bool(r.get("CanTriage", False)),
                can_email=bool(r.get("CanEmail", False)),
                can_phones=bool(r.get("CanPhones", False)),
                can_bookings=bool(r.get("CanBookings", False)),
                can_emis=bool(r.get("CanEMIS", False)),
                can_docman=bool(r.get("CanDocman", False)),
                weights=weights,
                frontdesk_only=bool(r.get("FrontDeskOnly", False)),
                break_required=bool(r.get("BreakRequired", True)),
            )
        )

    # --- Working hours
    hours_df = hours_df.copy()
    hn = pick_col(hours_df, ["Name","StaffName"])
    hours_df["Name"] = hours_df[hn].astype(str).str.strip()

    for dn in ["Mon","Tue","Wed","Thu","Fri"]:
        sc = pick_col(hours_df, [f"{dn}Start", f"{dn} Start", f"{dn}_Start"], required=False)
        ec = pick_col(hours_df, [f"{dn}End", f"{dn} End", f"{dn}_End"], required=False)
        hours_df[f"{dn}Start"] = hours_df[sc].apply(to_time) if sc else None
        hours_df[f"{dn}End"]   = hours_df[ec].apply(to_time) if ec else None

    hours_map = {}
    for _, r in hours_df.iterrows():
        hours_map[r["Name"]] = {k: r.get(k) for k in hours_df.columns}

    # --- Holidays (ranges)
    hols: List[Tuple[str,date,date,str]] = []
    if hols_df is not None and not hols_df.empty:
        ncol = pick_col(hols_df, ["Name","StaffName"], required=False) or hols_df.columns[0]
        sdcol = pick_col(hols_df, ["StartDate","Start"], required=False) or hols_df.columns[1]
        edcol = pick_col(hols_df, ["EndDate","End"], required=False) or hols_df.columns[2]
        notes_c = pick_col(hols_df, ["Notes","Note","Reason"], required=False)
        for _, r in hols_df.iterrows():
            nm = str(r.get(ncol,"")).strip()
            sd = to_date(r.get(sdcol))
            ed = to_date(r.get(edcol))
            note = "" if (not notes_c or pd.isna(r.get(notes_c))) else str(r.get(notes_c)).strip().lower()
            kind = "Holiday"
            if "sick" in note or "sickness" in note:
                kind = "Sick"
            elif "bank" in note:
                kind = "Bank Holiday"
            if nm and sd and ed:
                hols.append((nm, sd, ed, kind))

    # --- Targets (hourly)
    def parse_hourly(df: pd.DataFrame) -> Dict[Tuple[str, time], int]:
        out: Dict[Tuple[str,time], int] = {}
        if df is None or df.empty:
            return out
        time_col = pick_col(df, ["Time"], required=False) or df.columns[0]
        ddf = df.copy()
        ddf["Time"] = ddf[time_col].apply(to_time)
        for dn in ["Mon","Tue","Wed","Thu","Fri"]:
            if dn not in ddf.columns:
                continue
            for _, r in ddf.iterrows():
                hh = r.get("Time")
                if not hh:
                    continue
                val = r.get(dn)
                if pd.isna(val) or val == "":
                    continue
                out[(dn, time(hh.hour, 0))] = int(float(val))
        return out

    phones_targets   = parse_hourly(tph_df)
    bookings_targets = parse_hourly(tbk_df)

    weekly_targets = {"Bookings": 0.0, "EMIS": 0.0, "Docman": 0.0}
    if tweek_df is not None and not tweek_df.empty:
        task_c = pick_col(tweek_df, ["Task"], required=False) or tweek_df.columns[0]
        val_c  = pick_col(tweek_df, ["WeekHoursTarget","Target","Hours"], required=False) or tweek_df.columns[1]
        for _, r in tweek_df.iterrows():
            tsk = str(r.get(task_c,"")).strip()
            val = r.get(val_c)
            if pd.isna(val) or val == "":
                continue
            if tsk in weekly_targets:
                weekly_targets[tsk] = float(val)

    # Swaps
    swaps: List[Tuple[date, str, str, Optional[time], Optional[time]]] = []
    if swaps_df is not None and not swaps_df.empty:
        dcol = pick_col(swaps_df, ["Date"], required=False) or swaps_df.columns[0]
        ncol = pick_col(swaps_df, ["Name"], required=False) or swaps_df.columns[1]
        swcol = pick_col(swaps_df, ["SwapWith"], required=False)
        nscol = pick_col(swaps_df, ["NewStart"], required=False)
        necol = pick_col(swaps_df, ["NewEnd"], required=False)
        for _, r in swaps_df.iterrows():
            dd = to_date(r.get(dcol))
            if not dd:
                continue
            nm = str(r.get(ncol,"")).strip()
            sw = str(r.get(swcol,"")).strip() if swcol else ""
            ns = to_time(r.get(nscol)) if nscol else None
            ne = to_time(r.get(necol)) if necol else None
            if nm:
                swaps.append((dd, nm, sw, ns, ne))

    return TemplateData(
        staff=staff_list,
        hours_map=hours_map,
        hols=hols,
        call_handlers=callh_df if callh_df is not None else pd.DataFrame(),
        handler_leave=hleave_df if hleave_df is not None else pd.DataFrame(),
        phones_targets=phones_targets,
        bookings_targets=bookings_targets,
        weekly_targets=weekly_targets,
        swaps=swaps,
    )

# ---------- helpers ----------
def holiday_kind(name: str, d: date, hols: List[Tuple[str,date,date,str]]) -> Optional[str]:
    for n, s, e, k in hols:
        if n.strip().lower() == name.strip().lower() and s and e and s <= d <= e:
            return k
    return None

def shift_window(hours_map: Dict[str, Dict[str, Optional[time]]], d: date, name: str) -> Tuple[Optional[time], Optional[time]]:
    dn = day_name(d)
    hr = hours_map.get(name)
    if hr is None:
        return None, None
    return hr.get(f"{dn}Start"), hr.get(f"{dn}End")

def is_working(hours_map: Dict[str, Dict[str, Optional[time]]], d: date, t: time, name: str) -> bool:
    stt, end = shift_window(hours_map, d, name)
    return bool(stt and end and (t >= stt) and (t < end))

def apply_swaps(hours_map: Dict[str, Dict[str, Optional[time]]], swaps, week_dates: List[date]) -> Dict[str, Dict[str, Optional[time]]]:
    out = {k: dict(v) for k, v in hours_map.items()}
    in_week = set(week_dates)
    for dd, nm, sw, ns, ne in swaps:
        if dd not in in_week:
            continue
        dn = day_name(dd)
        if nm not in out:
            continue
        if sw and sw in out:
            a1, a2 = out[nm].get(f"{dn}Start"), out[nm].get(f"{dn}End")
            b1, b2 = out[sw].get(f"{dn}Start"), out[sw].get(f"{dn}End")
            out[nm][f"{dn}Start"], out[nm][f"{dn}End"] = b1, b2
            out[sw][f"{dn}Start"], out[sw][f"{dn}End"] = a1, a2
        elif ns and ne:
            out[nm][f"{dn}Start"], out[nm][f"{dn}End"] = ns, ne
    return out

def parse_handler_leave(df: pd.DataFrame) -> List[Tuple[str, date, date]]:
    if df is None or df.empty:
        return []
    ncol = pick_col(df, ["Name","HandlerName","CallHandler"], required=False) or df.columns[0]
    sdcol = pick_col(df, ["LeaveStartDate","LeaveStart","StartDate"], required=False) or df.columns[1]
    edcol = pick_col(df, ["LeaveEndDate","LeaveEnd","EndDate"], required=False) or df.columns[2]
    out = []
    for _, r in df.iterrows():
        nm = str(r.get(ncol,"")).strip()
        sd = to_date(r.get(sdcol))
        ed = to_date(r.get(edcol))
        if nm and sd and ed:
            out.append((nm, sd, ed))
    return out

def handler_working(callh_row: pd.Series, d: date, t: time) -> bool:
    dn = day_name(d)
    stt = to_time(callh_row.get(f"{dn}Start"))
    end = to_time(callh_row.get(f"{dn}End"))
    return bool(stt and end and (t >= stt) and (t < end))

def phones_required(tpl: TemplateData, d: date, t: time) -> int:
    dn = day_name(d)
    hour_key = time(t.hour, 0)
    base = int(tpl.phones_targets.get((dn, hour_key), 0))

    # Add 1 for each call handler who would otherwise be working but is on leave that day
    leave_ranges = parse_handler_leave(tpl.handler_leave)
    off = 0
    if tpl.call_handlers is not None and not tpl.call_handlers.empty:
        for _, r in tpl.call_handlers.iterrows():
            nm = str(r.get("Name","") or r.get("HandlerName","")).strip()
            if not nm:
                continue
            if not handler_working(r, d, t):
                continue
            for ln, sd, ed in leave_ranges:
                if ln.strip().lower() == nm.strip().lower() and sd <= d <= ed:
                    off += 1
                    break
    return base + off

def awaiting_site_for_day(d: date) -> str:
    wd = d.weekday()
    if wd in (0,4):  # Mon/Fri
        return "SLGP"
    if wd in (1,3):  # Tue/Thu
        return "JEN"
    return "BGS"

def email_site_for_day(d: date) -> str:
    return awaiting_site_for_day(d)

def pick_breaks_site_balanced(staff_list: List[Staff], hours_map, hols, week_dates: List[date], fixed_assignments: Set[Tuple[date,time,str]]) -> Dict[Tuple[date,time], Set[str]]:
    breaks: Dict[Tuple[date,time], Set[str]] = {}
    break_load: Dict[Tuple[date,str,time], int] = {}
    for d in week_dates:
        for st in staff_list:
            if not st.break_required:
                continue
            if holiday_kind(st.name, d, hols):
                continue
            stt, end = shift_window(hours_map, d, st.name)
            if not stt or not end:
                continue
            dur = (dt_of(d, end) - dt_of(d, stt)).total_seconds() / 3600.0
            if dur <= BREAK_THRESHOLD_HOURS:
                continue
            midpoint = dt_of(d, stt) + (dt_of(d, end) - dt_of(d, stt)) / 2
            best = None
            for bt in BREAK_CANDIDATES:
                if bt < stt or add_minutes(bt, 30) > end:
                    continue
                if not t_in_range(bt, BREAK_WINDOW[0], BREAK_WINDOW[1]):
                    continue
                if (d, bt, st.name) in fixed_assignments:
                    continue
                before = (dt_of(d, bt) - dt_of(d, stt)).total_seconds() / 3600.0
                after = (dt_of(d, end) - dt_of(d, add_minutes(bt, 30))).total_seconds() / 3600.0
                frag_penalty = 0
                if before < 1.0:
                    frag_penalty += 10_000
                if after < 1.0:
                    frag_penalty += 10_000
                load = break_load.get((d, st.home, bt), 0)
                dist = abs((dt_of(d, bt) - midpoint).total_seconds())
                score = frag_penalty + (load * 3600) + dist
                if best is None or score < best[0]:
                    best = (score, bt)
            if best:
                bt = best[1]
                breaks.setdefault((d, bt), set()).add(st.name)
                break_load[(d, st.home, bt)] = break_load.get((d, st.home, bt), 0) + 1
    return breaks

# ---------- Scheduling ----------
def task_weight(st: Staff, task_key: str) -> int:
    return int(st.weights.get(task_key, 3) if st.weights is not None else 3)

def block_limits(task: str) -> Tuple[int,int]:
    if task == "Phones":
        return MIN_PHONES, MAX_PHONES
    if task == "Docman":
        return MIN_DOCMAN, MAX_DEFAULT
    return MIN_DEFAULT, MAX_DEFAULT

def schedule_week(tpl: TemplateData, week_start: date):
    slots = timeslots()
    dates = [week_start + timedelta(days=i) for i in range(5)]
    hours_map = apply_swaps(tpl.hours_map, tpl.swaps, dates)

    staff_by_name = {s.name: s for s in tpl.staff}
    staff_names = [s.name for s in tpl.staff]

    a: Dict[Tuple[date,time,str], str] = {}
    gaps: List[Tuple[date,time,str,str]] = []

    mins_task: Dict[Tuple[str,str], int] = {}

    def add_mins(nm: str, task_key: str, mins: int):
        mins_task[(nm, task_key)] = mins_task.get((nm, task_key), 0) + mins

    def is_free(nm: str, d: date, t: time) -> bool:
        return (d,t,nm) not in a

    fixed_slots: Set[Tuple[date,time,str]] = set()

    def can_cover_full_band(nm: str, d: date, bs: time, be: time) -> bool:
        stt, endt = shift_window(hours_map, d, nm)
        if not stt or not endt:
            return False
        if not ((stt <= bs) and (endt >= be)):
            return False
        for tt in slots:
            if tt < bs or tt >= be:
                continue
            if (d, tt, nm) in fixed_slots:
                return False
            if (d, tt, nm) in a:
                return False
        return True

    def pick_for_band(candidates: List[str], d: date, task_key: str, bs: time, be: time) -> Optional[str]:
        ok = []
        for nm in candidates:
            if holiday_kind(nm, d, tpl.hols):
                continue
            if not can_cover_full_band(nm, d, bs, be):
                continue
            ok.append(nm)
        if not ok:
            return None

        def score(nm: str):
            st = staff_by_name[nm]
            fd_only = 1 if (task_key == "FrontDesk" and st.frontdesk_only) else 0
            w = task_weight(st, task_key)
            used = mins_task.get((nm, task_key), 0)
            # Highest weight first, then least used
            return (-fd_only, -w, used, nm.lower())

        ok.sort(key=score)
        return ok[0]

    # Front Desk fixed bands
    for d in dates:
        for site in SITES:
            role = f"FrontDesk_{site}"
            cands = [s.name for s in tpl.staff if s.can_frontdesk and s.home == site]
            for bs, be in FD_BANDS:
                chosen = pick_for_band(cands, d, "FrontDesk", bs, be)
                if not chosen:
                    gaps.append((d, bs, role, "No suitable staff for full band"))
                    continue
                for tt in slots:
                    if tt < bs or tt >= be:
                        continue
                    a[(d, tt, chosen)] = role
                    fixed_slots.add((d, tt, chosen))
                    add_mins(chosen, "FrontDesk", SLOT_MIN)

    # Triage fixed bands (SLGP/JEN)
    for d in dates:
        for site in ("SLGP","JEN"):
            role = f"Triage_Admin_{site}"
            cands = [s.name for s in tpl.staff if s.can_triage and s.home == site]
            for bs, be in TRIAGE_BANDS:
                chosen = pick_for_band(cands, d, "Triage", bs, be)
                if not chosen:
                    gaps.append((d, bs, role, "No suitable staff for full band"))
                    continue
                for tt in slots:
                    if tt < bs or tt >= be:
                        continue
                    a[(d, tt, chosen)] = role
                    fixed_slots.add((d, tt, chosen))
                    add_mins(chosen, "Triage", SLOT_MIN)

    breaks = pick_breaks_site_balanced(tpl.staff, hours_map, tpl.hols, dates, fixed_slots)

    def on_break(nm: str, d: date, t: time) -> bool:
        return nm in breaks.get((d,t), set())

    active: Dict[Tuple[date,str], Tuple[str,int]] = {}

    def task_key_for_task(task: str) -> str:
        if task.startswith("FrontDesk_"):
            return "FrontDesk"
        if task.startswith("Triage_Admin_"):
            return "Triage"
        if task == "Email_Box":
            return "Email"
        if task == "Awaiting_PSA_Admin":
            return "Awaiting"
        if task == "Phones":
            return "Phones"
        if task == "Bookings":
            return "Bookings"
        if task == "EMIS":
            return "EMIS"
        if task == "Docman":
            return "Docman"
        if task == "Misc_Tasks":
            return "Misc"
        return "Misc"

    def eligible(nm: str, task: str, d: date, t: time, allow_cross_site: bool=False) -> bool:
        st = staff_by_name[nm]
        if holiday_kind(nm, d, tpl.hols):
            return False
        if not is_working(hours_map, d, t, nm):
            return False
        if on_break(nm, d, t):
            return False
        if task == "Email_Box":
            if not st.can_email:
                return False
            if allow_cross_site:
                return True
            return st.home == email_site_for_day(d)
        if task == "Awaiting_PSA_Admin":
            if not st.can_docman:
                return False
            if allow_cross_site:
                return True
            return st.home == awaiting_site_for_day(d)
        if task == "Phones":
            return st.can_phones
        if task == "Bookings":
            if not st.can_bookings:
                return False
            if allow_cross_site:
                return True
            return st.home == "SLGP"
        if task == "EMIS":
            return st.can_emis
        if task == "Docman":
            return st.can_docman
        if task == "Misc_Tasks":
            return True
        return False

    def start_block(nm: str, task: str, d: date, start_idx: int) -> bool:
        mn, mx = block_limits(task)
        stt, end = shift_window(hours_map, d, nm)
        if not stt or not end:
            return False
        end_idx = start_idx
        while end_idx < len(slots) and slots[end_idx] < end:
            if (d, slots[end_idx], nm) in fixed_slots:
                break
            if nm in breaks.get((d, slots[end_idx]), set()):
                break
            end_idx += 1

        remaining = end_idx - start_idx
        if remaining <= 0:
            return False

        # No 30-min floaters: allow < min only if it's end-of-shift remainder
        if remaining < mn:
            active[(d,nm)] = (task, start_idx + remaining)
            return True

        L = min(mx, remaining)
        L = max(mn, L)
        active[(d,nm)] = (task, start_idx + L)
        return True

    def apply_active(nm: str, d: date, idx: int) -> bool:
        b = active.get((d,nm))
        if not b:
            return False
        task, end_idx = b
        if idx >= end_idx:
            del active[(d,nm)]
            return False
        t = slots[idx]
        if not is_free(nm,d,t):
            del active[(d,nm)]
            return False
        a[(d,t,nm)] = task
        add_mins(nm, task_key_for_task(task), SLOT_MIN)
        return True

    def stop_block(nm: str, d: date):
        if (d,nm) in active:
            del active[(d,nm)]

    def pick_candidates(task: str, d: date, t: time, allow_cross_site: bool=False, prefer_sites: Optional[List[str]]=None) -> List[str]:
        cands = []
        for nm in staff_names:
            if not is_free(nm,d,t):
                continue
            if not eligible(nm, task, d, t, allow_cross_site=allow_cross_site):
                continue
            cands.append(nm)

        def score(nm: str):
            st = staff_by_name[nm]
            key = task_key_for_task(task)
            w = task_weight(st, key)
            used = mins_task.get((nm, key), 0)
            site_pen = 0
            if prefer_sites and st.home not in prefer_sites:
                site_pen = 1
            return (-w, site_pen, used, nm.lower())

        cands.sort(key=score)
        return cands

    def assign_block(nm: str, task: str, d: date, idx: int):
        b = active.get((d,nm))
        if b and b[0] == task:
            apply_active(nm, d, idx)
            return
        if b and b[0] != task:
            stop_block(nm, d)
        ok = start_block(nm, task, d, idx)
        if not ok:
            a[(d, slots[idx], nm)] = "Misc_Tasks"
            add_mins(nm, "Misc", SLOT_MIN)
            return
        apply_active(nm, d, idx)

    # weekly target minutes
    target_book = int(round((tpl.weekly_targets.get("Bookings", 0.0) or 0.0) * 60))
    target_emis = int(round(20.0 * 60))  # HARD RULE: 20h/week
    target_doc  = int(round(14.0 * 60))  # HARD RULE: 14h/week

    def total_mins(task_key: str) -> int:
        return sum(v for (nm, tk), v in mins_task.items() if tk == task_key)

    def bookings_needed_this_slot(d: date, idx: int) -> int:
        # Use HOURLY matrix after 10:30; apply to both half-hours
        t = slots[idx]
        dn = day_name(d)
        hour_key = time(t.hour, 0)
        base = int(tpl.bookings_targets.get((dn, hour_key), 0))
        if t < time(10,30):
            base = 0

        # also apply weekly pressure if behind
        if target_book <= 0:
            return base
        done = total_mins("Bookings")
        remaining = max(0, target_book - done)
        if remaining <= 0:
            return base
        rem_slots = 0
        for dd in dates:
            for tt in slots:
                if dd < d:
                    continue
                if dd == d and tt < t:
                    continue
                if tt >= time(10,30):
                    rem_slots += 1
        if rem_slots <= 0:
            return base
        ppl_pressure = math.ceil(remaining / (rem_slots * SLOT_MIN))
        return max(base, ppl_pressure)

    def _set_assignment(d: date, t: time, nm: str, new_task: str):
        """Overwrite (d,t,nm) assignment safely, keeping mins_task approximately consistent.
        Only used for controlled "reclaim" moves (e.g., from Misc/EMIS/Docman -> higher-pressure task).
        """
        old = a.get((d,t,nm))
        if old == new_task:
            return
        if old:
            old_key = task_key_for_task(old)
            mins_task[(nm, old_key)] = max(0, mins_task.get((nm, old_key), 0) - SLOT_MIN)
        a[(d,t,nm)] = new_task
        add_mins(nm, task_key_for_task(new_task), SLOT_MIN)

    def enforce(task: str, need: int, d: date, idx: int, allow_cross_site: bool=False, prefer_sites: Optional[List[str]]=None, reclaim_from: Optional[Set[str]] = None):
        """Ensure at least `need` staff are on `task` at this slot.
        If reclaim_from is provided, may reassign staff currently on one of those tasks.
        """
        t = slots[idx]
        reclaim_from = reclaim_from or set()
        while True:
            current = len([nm for nm in staff_names if a.get((d,t,nm)) == task])
            if current >= need:
                return

            # Candidates: free first; then reclaimable assignments
            free_cands = pick_candidates(task, d, t, allow_cross_site=allow_cross_site, prefer_sites=prefer_sites)
            if free_cands:
                nm = free_cands[0]
                assign_block(nm, task, d, idx)
                continue

            if reclaim_from:
                reclaimable = []
                for nm in staff_names:
                    cur = a.get((d,t,nm))
                    if cur not in reclaim_from:
                        continue
                    if not eligible(nm, task, d, t, allow_cross_site=allow_cross_site):
                        continue
                    reclaimable.append(nm)
                if reclaimable:
                    key = task_key_for_task(task)
                    reclaimable = sorted(reclaimable, key=lambda nm: (-task_weight(staff_by_name[nm], key), mins_task.get((nm, key), 0), nm.lower()))
                    nm = reclaimable[0]
                    stop_block(nm, d)
                    _set_assignment(d, t, nm, task)
                    continue

            gaps.append((d, t, task, f"Short by {need-current}"))
            return

    def pressure_needed(task_key: str, target_mins: int, done_mins: int, d: date, idx: int, window_start: time = DAY_START) -> int:
        """How many people should we schedule *this slot* for a weekly-hours target.
        This is a soft pressure (ceil of remaining / remaining slots).
        """
        if target_mins <= 0:
            return 0
        remaining = max(0, target_mins - done_mins)
        if remaining <= 0:
            return 0
        t = slots[idx]
        rem_slots = 0
        for dd in dates:
            for tt in slots:
                if dd < d:
                    continue
                if dd == d and tt < t:
                    continue
                if tt >= window_start:
                    rem_slots += 1
        if rem_slots <= 0:
            return 0
        return max(0, math.ceil(remaining / (rem_slots * SLOT_MIN)))

    # main loop
    for d in dates:
        for idx, t in enumerate(slots):
            # If weekly target for EMIS/Docman is already met, do NOT let blocks keep running.
            # This prevents overshooting targets just because someone was mid-block.
            if target_emis > 0 or target_doc > 0:
                emis_done = total_mins("EMIS")
                doc_done = total_mins("Docman")
                for nm in staff_names:
                    b = active.get((d, nm))
                    if not b:
                        continue
                    task_running = b[0]
                    if task_running == "EMIS" and target_emis > 0 and emis_done >= target_emis:
                        stop_block(nm, d)
                    elif task_running == "Docman" and target_doc > 0 and doc_done >= target_doc:
                        stop_block(nm, d)

            # apply active blocks for stability
            for nm in staff_names:
                if (d,t,nm) in a:
                    continue
                if on_break(nm,d,t):
                    continue
                apply_active(nm, d, idx)

            # Email 10:30–16:00 (site-of-day, then cross-site if needed)
            if t_in_range(t, time(10,30), time(16,0)):
                enforce("Email_Box", 1, d, idx, allow_cross_site=False, reclaim_from={"Misc_Tasks","EMIS","Docman"})
                if len([nm for nm in staff_names if a.get((d,t,nm)) == "Email_Box"]) < 1:
                    enforce("Email_Box", 1, d, idx, allow_cross_site=True, reclaim_from={"Misc_Tasks","EMIS","Docman"})

            # Awaiting 10:00–16:00
            if t_in_range(t, time(10,0), time(16,0)):
                enforce("Awaiting_PSA_Admin", 1, d, idx, allow_cross_site=False, reclaim_from={"Misc_Tasks","EMIS","Docman"})
                if len([nm for nm in staff_names if a.get((d,t,nm)) == "Awaiting_PSA_Admin"]) < 1:
                    enforce("Awaiting_PSA_Admin", 1, d, idx, allow_cross_site=True, reclaim_from={"Misc_Tasks","EMIS","Docman"})

            # Phones per hourly matrix (hard)
            req_p = phones_required(tpl, d, t)
            if req_p > 0:
                enforce("Phones", req_p, d, idx, allow_cross_site=True)

            # Bookings per hourly matrix (and weekly pressure), SLGP first then cross-site
            req_b = bookings_needed_this_slot(d, idx)
            if req_b > 0:
                enforce("Bookings", req_b, d, idx, allow_cross_site=False)
                if len([nm for nm in staff_names if a.get((d,t,nm)) == "Bookings"]) < req_b:
                    enforce("Bookings", req_b, d, idx, allow_cross_site=True)

        # Docman/EMIS weekly targets (soft): allocate as close as possible to targets.
        # Strategy:
        #  - Calculate "pressure" (people needed this slot) for each task based on remaining minutes / remaining slots
        #  - Enforce the higher-pressure task first, allowing reclaim from Misc/other admin tasks
        #  - Fill any remaining free staff to the more-behind task (site preference JEN/BGS first; SLGP last)

        emis_done = total_mins("EMIS")
        doc_done  = total_mins("Docman")

        doc_need  = pressure_needed("Docman", target_doc, doc_done, d, idx)
        emis_need = pressure_needed("EMIS",   target_emis, emis_done, d, idx)

        # Decide which is more behind (ratio remaining/target)
        doc_ratio  = (max(0, target_doc - doc_done) / target_doc) if target_doc else 0.0
        emis_ratio = (max(0, target_emis - emis_done) / target_emis) if target_emis else 0.0

        # Enforce in descending ratio order; allow reclaim from Misc and the other task (but never from Phones/Bookings/etc.)
        reclaim_low = {"Misc_Tasks"}
        if doc_ratio >= emis_ratio:
            if doc_need > 0:
                enforce("Docman", doc_need, d, idx, allow_cross_site=True, prefer_sites=["JEN","BGS"], reclaim_from=reclaim_low | {"EMIS"})
            if emis_need > 0:
                enforce("EMIS", emis_need, d, idx, allow_cross_site=True, prefer_sites=["JEN","BGS"], reclaim_from=reclaim_low | {"Docman"})
        else:
            if emis_need > 0:
                enforce("EMIS", emis_need, d, idx, allow_cross_site=True, prefer_sites=["JEN","BGS"], reclaim_from=reclaim_low | {"Docman"})
            if doc_need > 0:
                enforce("Docman", doc_need, d, idx, allow_cross_site=True, prefer_sites=["JEN","BGS"], reclaim_from=reclaim_low | {"EMIS"})

        #         # Fill remaining free staff:
        # - While EMIS/Docman weekly targets are still outstanding, assign those tasks (equal priority by deficit).
        # - Prefer JEN/BGS staff for EMIS/Docman first (ordering), then SLGP if still behind.
        # - Misc_Tasks is TRUE overflow only (used only when both targets met, or staff not eligible for either).
        free_staff = []
        for nm in staff_names:
            if not is_working(hours_map, d, t, nm):
                continue
            if holiday_kind(nm, d, tpl.hols):
                continue
            if on_break(nm, d, t):
                continue
            if not is_free(nm, d, t):
                continue
            free_staff.append(nm)

        # Prefer JEN/BGS staff first for EMIS/Docman, then SLGP
        free_staff.sort(key=lambda nm: (0 if staff_by_name[nm].home in ("JEN","BGS") else 1, nm.lower()))

        for nm in free_staff:
            doc_done  = total_mins("Docman")
            emis_done = total_mins("EMIS")
            doc_rem   = max(0, target_doc - doc_done)
            emis_rem  = max(0, target_emis - emis_done)

            # HARD RULE: Misc is strictly last. If either target still needs time,
            # allocate up to 1 Docman + 1 EMIS at a time (per slot), otherwise Misc.
            doc_now = len([x for x in staff_names if a.get((d,t,x)) == "Docman"])
            emis_now = len([x for x in staff_names if a.get((d,t,x)) == "EMIS"])

            if doc_rem <= 0 and emis_rem <= 0:
                assign_block(nm, "Misc_Tasks", d, idx)
                continue

            # If both still needed, pick the one that is proportionally further behind
            choice = None
            if doc_rem > 0 and emis_rem > 0:
                doc_ratio = doc_rem / max(1, target_doc)
                emis_ratio = emis_rem / max(1, target_emis)
                choice = "Docman" if doc_ratio >= emis_ratio else "EMIS"
            elif doc_rem > 0:
                choice = "Docman"
            elif emis_rem > 0:
                choice = "EMIS"

            if choice == "Docman" and doc_now < 1:
                assign_block(nm, "Docman", d, idx)
                continue
            if choice == "EMIS" and emis_now < 1:
                assign_block(nm, "EMIS", d, idx)
                continue

            # If preferred choice cap already filled, try the other (if available)
            if choice == "Docman" and emis_rem > 0 and emis_now < 1:
                assign_block(nm, "EMIS", d, idx)
                continue
            if choice == "EMIS" and doc_rem > 0 and doc_now < 1:
                assign_block(nm, "Docman", d, idx)
                continue

            # Otherwise, everyone else goes to Misc
            assign_block(nm, "Misc_Tasks", d, idx)
# Smooth any single-slot fragments (non-fixed)
    def smooth_single_slot_blocks():
        FIXED_PREFIXES = ("FrontDesk_", "Triage_Admin_")
        SPECIAL = ("Break", "Holiday", "Bank Holiday", "Sick")
        for d in dates:
            for nm in staff_names:
                seq = [a.get((d, t, nm)) for t in slots]
                i = 0
                while i < len(slots):
                    task = seq[i]
                    if not task:
                        i += 1
                        continue
                    j = i + 1
                    while j < len(slots) and seq[j] == task:
                        j += 1
                    if (j - i) == 1:
                        tsk = str(task)
                        if tsk.startswith(FIXED_PREFIXES) or tsk in SPECIAL:
                            i = j
                            continue
                        prev = seq[i-1] if i-1 >= 0 else None
                        nxt = seq[j] if j < len(slots) else None
                        chosen = None
                        if prev and (not str(prev).startswith(FIXED_PREFIXES)) and str(prev) not in SPECIAL:
                            chosen = prev
                        elif nxt and (not str(nxt).startswith(FIXED_PREFIXES)) and str(nxt) not in SPECIAL:
                            chosen = nxt
                        if chosen:
                            # Don't violate EMIS/Docman concurrency caps when smoothing
                            if chosen in ("EMIS","Docman"):
                                cur = sum(1 for n2 in staff_names if a.get((d, slots[i], n2)) == chosen)
                                if cur >= 1:
                                    chosen = None
                            if chosen:
                                a[(d, slots[i], nm)] = chosen
                                seq[i] = chosen
                    i = j
    smooth_single_slot_blocks()

    # Enforce hard concurrency caps after smoothing (at most 1 EMIS + 1 Docman per slot)
    for d in dates:
        for t in slots:
            for task in ("EMIS","Docman"):
                assignees = [nm for nm in staff_names if a.get((d,t,nm)) == task]
                if len(assignees) <= 1:
                    continue
                # Keep highest weight, then least used
                def _score(nm):
                    st = staff_by_name[nm]
                    w = task_weight(st, task)
                    used = mins_task.get((nm, task), 0)
                    return (-w, used, nm.lower())
                assignees.sort(key=_score)
                keep = assignees[0]
                for extra in assignees[1:]:
                    a[(d,t,extra)] = "Misc_Tasks"

    return a, breaks, gaps, dates, slots, hours_map

# ---------- Excel output ----------
def _apply_day_borders(ws):
    # Every day starts at 08:00 (row groups of 21 slots per day)
    for rr in range(2, ws.max_row + 1):
        if ws.cell(rr, 2).value == "08:00":
            for cc in range(1, ws.max_column + 1):
                ws.cell(rr, cc).border = DAY_TOP
        for cc in range(1, ws.max_column + 1):
            cell = ws.cell(rr, cc)
            if cell.border is None or cell.border == Border():
                cell.border = CELL_BORDER

def _add_colour_rules(ws, data_range: str, tl_cell: str):
    # Equality based conditional formatting (reliable for edits)
    for task, color in ROLE_COLORS.items():
        if task == "":
            continue
        fill = PatternFill("solid", fgColor=color)
        # exact match
        ws.conditional_formatting.add(
            data_range,
            FormulaRule(formula=[f'{tl_cell}="{task}"'], fill=fill, stopIfTrue=False)
        )

def build_workbook(tpl: TemplateData, start_monday: date, weeks: int) -> Workbook:
    wb = Workbook()
    wb.remove(wb.active)

    all_staff = [s.name for s in tpl.staff]
    staff_by_name = {s.name: s for s in tpl.staff}
    slots = timeslots()

    for w in range(weeks):
        wk_start = start_monday + timedelta(days=7*w)
        a, breaks, gaps, dates, slots, hours_map = schedule_week(tpl, wk_start)

        # Site timelines (ONLY site staff columns)
        site_sheets = {}
        for site in ("SLGP","JEN","BGS"):
            site_staff = [nm for nm in all_staff if staff_by_name[nm].home == site]
            ws = wb.create_sheet(f"Week{w+1}_{site}_Timeline")
            ws.append(["Date","Time"] + site_staff)
            for c in ws[1]:
                c.font = Font(bold=True)
                c.alignment = Alignment(horizontal="center", vertical="center")
            ws.freeze_panes = "C2"
            ws.column_dimensions["A"].width = 14
            ws.column_dimensions["B"].width = 8
            for i in range(len(site_staff)):
                ws.column_dimensions[get_column_letter(3+i)].width = 18
            site_sheets[site] = (ws, site_staff)

        # Fill values
        for d in dates:
            for t in slots:
                date_label = d.strftime("%a %d-%b")
                time_label = t.strftime("%H:%M")
                for site, (ws, site_staff) in site_sheets.items():
                    row = [date_label, time_label]
                    for nm in site_staff:
                        hk = holiday_kind(nm, d, tpl.hols)
                        if hk:
                            val = hk
                        elif not is_working(hours_map, d, t, nm):
                            val = ""
                        elif nm in breaks.get((d,t), set()):
                            val = "Break"
                        else:
                            val = a.get((d,t,nm), "Misc_Tasks")
                        row.append(val)
                    ws.append(row)

        # Conditional formatting for colours (works even after editing)
        for site, (ws, site_staff) in site_sheets.items():
            if not site_staff:
                continue
            start = get_column_letter(3) + "2"
            end   = get_column_letter(2 + len(site_staff)) + str(ws.max_row)
            rng = f"{get_column_letter(3)}2:{get_column_letter(2+len(site_staff))}{ws.max_row}"
            _add_colour_rules(ws, rng, start)
            for rr in range(2, ws.max_row+1):
                for cc in range(3, ws.max_column+1):
                    cell = ws.cell(rr, cc)
                    v = str(cell.value or "")
                    cell.fill = _fill_for(v)
                    cell.alignment = Alignment(wrap_text=True, vertical="top")
            _apply_day_borders(ws)

        # Coverage by slot by site (names)
        ws_cov = wb.create_sheet(f"Week{w+1}_Coverage_By_Slot_By_Site")
        cov_cols = ["FD_SLGP","FD_JEN","FD_BGS","Triage_SLGP","Triage_JEN","Phones","Bookings","EMIS","Docman","Awaiting","Email","Misc"]
        ws_cov.append(["Date","Time"] + cov_cols)
        for c in ws_cov[1]:
            c.font = Font(bold=True)
            c.alignment = Alignment(horizontal="center", vertical="center")
        ws_cov.freeze_panes = "C2"
        ws_cov.column_dimensions["A"].width = 14
        ws_cov.column_dimensions["B"].width = 8
        for i in range(len(cov_cols)):
            ws_cov.column_dimensions[get_column_letter(3+i)].width = 22

        def names_for(prefix: str, d: date, t: time) -> str:
            return ", ".join([nm for nm in all_staff if str(a.get((d,t,nm),"")).startswith(prefix)])

        for d in dates:
            for t in slots:
                row = [d.strftime("%a %d-%b"), t.strftime("%H:%M")]
                row.append(names_for("FrontDesk_SLGP", d, t))
                row.append(names_for("FrontDesk_JEN", d, t))
                row.append(names_for("FrontDesk_BGS", d, t))
                row.append(names_for("Triage_Admin_SLGP", d, t))
                row.append(names_for("Triage_Admin_JEN", d, t))
                row.append(", ".join([nm for nm in all_staff if a.get((d,t,nm))=="Phones"]))
                row.append(", ".join([nm for nm in all_staff if a.get((d,t,nm))=="Bookings"]))
                row.append(", ".join([nm for nm in all_staff if a.get((d,t,nm))=="EMIS"]))
                row.append(", ".join([nm for nm in all_staff if a.get((d,t,nm))=="Docman"]))
                row.append(", ".join([nm for nm in all_staff if a.get((d,t,nm))=="Awaiting_PSA_Admin"]))
                row.append(", ".join([nm for nm in all_staff if a.get((d,t,nm))=="Email_Box"]))
                row.append(", ".join([nm for nm in all_staff if a.get((d,t,nm), "") in ("Misc_Tasks","Unassigned")]))
                ws_cov.append(row)


        # Colour coverage columns by meaning (static fill)
        col_map = {
            "FD_": ROLE_COLORS.get("FrontDesk_SLGP", "FFF2CC"),
            "Triage_": ROLE_COLORS.get("Triage_Admin_SLGP", "D9EAD3"),
            "Phones": ROLE_COLORS.get("Phones", "C9DAF8"),
            "Bookings": ROLE_COLORS.get("Bookings", "FCE5CD"),
            "EMIS": ROLE_COLORS.get("EMIS", "EAD1DC"),
            "Docman": ROLE_COLORS.get("Docman", "D0E0E3"),
            "Awaiting": ROLE_COLORS.get("Awaiting_PSA_Admin", "D0E0E3"),
            "Email": ROLE_COLORS.get("Email_Box", "CFE2F3"),
            "Misc": ROLE_COLORS.get("Misc_Tasks", "FFFFFF"),
        }
        for cc in range(3, ws_cov.max_column+1):
            hdr = str(ws_cov.cell(1, cc).value or "")
            hexcol = None
            for k, v in col_map.items():
                if hdr.startswith(k) or hdr == k:
                    hexcol = v
                    break
            if hexcol:
                f = PatternFill("solid", fgColor=hexcol)
                for rr in range(2, ws_cov.max_row+1):
                    ws_cov.cell(rr, cc).fill = f

        _apply_day_borders(ws_cov)
        for rr in range(2, ws_cov.max_row+1):
            for cc in range(1, ws_cov.max_column+1):
                ws_cov.cell(rr, cc).alignment = Alignment(wrap_text=True, vertical="top")

        # Totals (FORMULA from site timelines so edits update totals)
        ws_tot = wb.create_sheet(f"Week{w+1}_Totals")
        task_keys = ["FrontDesk","Triage","Email","Awaiting","Phones","Bookings","EMIS","Docman","Misc","Break"]
        ws_tot.append(["Name"] + task_keys + ["WeeklyTotal"])
        for c in ws_tot[1]:
            c.font = Font(bold=True)
            c.alignment = Alignment(horizontal="center", vertical="center")
        ws_tot.freeze_panes = "B2"
        ws_tot.column_dimensions["A"].width = 22
        for i in range(2, 2+len(task_keys)+1):
            ws_tot.column_dimensions[get_column_letter(i)].width = 12

        # Build COUNTIF formulas across each site timeline for each staff column
        for r, nm in enumerate(all_staff, start=2):
            row = [nm]
            # Find staff column index per site (if staff not in site sheet, countif range will be empty)
            for task in task_keys:
                if task == "FrontDesk":
                    crit = "FrontDesk_*"
                elif task == "Triage":
                    crit = "Triage_Admin_*"
                elif task == "Email":
                    crit = "Email_Box"
                elif task == "Awaiting":
                    crit = "Awaiting_PSA_Admin"
                elif task == "Phones":
                    crit = "Phones"
                elif task == "Bookings":
                    crit = "Bookings"
                elif task == "EMIS":
                    crit = "EMIS"
                elif task == "Docman":
                    crit = "Docman"
                elif task == "Misc":
                    crit = "Misc_Tasks"
                elif task == "Break":
                    crit = "Break"
                else:
                    crit = task

                parts = []
                for site in ("SLGP","JEN","BGS"):
                    ws_name = f"Week{w+1}_{site}_Timeline"
                    ws_site = wb[ws_name]
                    # find column for staff in that sheet
                    col_idx = None
                    for cc in range(3, ws_site.max_column+1):
                        if str(ws_site.cell(1,cc).value) == nm:
                            col_idx = cc
                            break
                    if col_idx is None:
                        continue
                    col_letter = get_column_letter(col_idx)
                    rng = f"{ws_name}!{col_letter}$2:{col_letter}${ws_site.max_row}"
                    parts.append(f'COUNTIF({rng},"{crit}")')
                if not parts:
                    row.append("=0")
                else:
                    row.append(f"=0.5*({'+'.join(parts)})")
            start_letter = get_column_letter(2)
            end_letter   = get_column_letter(1 + len(task_keys))
            row.append(f"=SUM({start_letter}{r}:{end_letter}{r})")
            ws_tot.append(row)

        # Dashboard (quick sense check)
        ws_dash = wb.create_sheet(f"Week{w+1}_Dashboard")
        ws_dash.column_dimensions["A"].width = 46
        ws_dash.column_dimensions["B"].width = 14
        ws_dash.column_dimensions["C"].width = 14
        ws_dash.column_dimensions["D"].width = 10
        ws_dash["A1"] = "Coverage Dashboard"
        ws_dash["A1"].font = Font(bold=True, size=14)

        total_slots = len(dates) * len(slots)

        def achieved(task: str) -> float:
            return round(sum(0.5 for d in dates for t in slots for nm in all_staff if a.get((d,t,nm)) == task), 2)

        phones_ok = 0
        for d in dates:
            for t in slots:
                req = phones_required(tpl, d, t)
                actual = sum(1 for nm in all_staff if a.get((d,t,nm)) == "Phones")
                if actual >= req:
                    phones_ok += 1

        ws_dash.append(["Metric","Achieved","Target","%"])
        for c in ws_dash[2]:
            c.font = Font(bold=True)
            c.alignment = Alignment(horizontal="center", vertical="center")

        book_h = achieved("Bookings")
        emis_h = achieved("EMIS")
        doc_h  = achieved("Docman")

        book_t = float(tpl.weekly_targets.get("Bookings", 0.0) or 0.0)
        emis_t = float(tpl.weekly_targets.get("EMIS", 0.0) or 0.0)
        doc_t  = float(tpl.weekly_targets.get("Docman", 0.0) or 0.0)

        rows = [
            ("Phones coverage (slots meeting requirement)", phones_ok, total_slots, (phones_ok/total_slots if total_slots else 1.0)),
            ("Bookings hours", book_h, book_t, (book_h/book_t if book_t else 1.0)),
            ("EMIS hours", emis_h, emis_t, (emis_h/emis_t if emis_t else 1.0)),
            ("Docman hours", doc_h, doc_t, (doc_h/doc_t if doc_t else 1.0)),
        ]
        for r in rows:
            ws_dash.append([r[0], r[1], r[2], round(float(r[3]), 3)])

        ws_dash.conditional_formatting.add(f"D3:D{ws_dash.max_row}", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color="63C384"))
        for rr in range(2, ws_dash.max_row+1):
            for cc in range(1, 5):
                ws_dash.cell(rr, cc).border = CELL_BORDER
                ws_dash.cell(rr, cc).alignment = Alignment(vertical="center")

        # Notes & gaps
        ws_g = wb.create_sheet(f"Week{w+1}_NotesAndGaps")
        ws_g.append(["Date","Time","Task","Note"])
        for c in ws_g[1]:
            c.font = Font(bold=True)
        for d, t, task, note in gaps:
            ws_g.append([d.isoformat(), t.strftime("%H:%M") if t else "", task, note])

    return wb