
import io
import re
from dataclasses import dataclass
from datetime import datetime, date, time, timedelta
from typing import Dict, List, Tuple, Optional, Set

import math

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows

# =========================================================
# Rota Generator Engine — v15 (Weighted, Block-Stable, No Floaters)
# =========================================================

DAY_START = time(8, 0)
DAY_END = time(18, 30)
SLOT_MIN = 30

SITES = ["SLGP", "JEN", "BGS"]

FD_BANDS = [
    (time(8, 0), time(11, 0)),
    (time(11, 0), time(13, 30)),
    (time(13, 30), time(16, 0)),
    (time(16, 0), time(18, 30)),
]
TRIAGE_BANDS = [
    (time(8, 0), time(10, 30)),
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
MIN_DOCMAN = 6          # 3h

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
    return pd.to_datetime(x).date()

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
    buddies: Dict[str, str]

def read_template(uploaded_bytes: bytes) -> TemplateData:
    xls = pd.ExcelFile(io.BytesIO(uploaded_bytes))

    staff_sheet = find_sheet(xls, ["Staff"])
    hours_sheet = find_sheet(xls, ["WorkingHours", "Hours"])
    hols_sheet = find_sheet(xls, ["Holidays", "Leave", "Absence"])
    callh_sheet = find_sheet(xls, ["CallHandlers", "Call Handlers"])
    hleave_sheet = find_sheet(xls, ["Handler_Leave", "Handler Leave", "CallHandler_Leave"])
    tph_sheet = find_sheet(xls, ["Targets_Phones_Hourly", "PhonesTargets", "Targets Phones Hourly"])
    tbk_sheet = find_sheet(xls, ["Targets_Bookings_Hourly", "BookingsTargets", "Targets Bookings Hourly"])
    tweek_sheet = find_sheet(xls, ["Targets_Weekly", "Targets Weekly"])
    swaps_sheet = find_sheet(xls, ["Swaps"])
    new_sheet = find_sheet(xls, ["NewStarters", "New Starters"])

    if not staff_sheet or not hours_sheet:
        raise ValueError(f"Missing Staff/WorkingHours sheets. Found: {xls.sheet_names}")

    staff_df = pd.read_excel(xls, sheet_name=staff_sheet)
    hours_df = pd.read_excel(xls, sheet_name=hours_sheet)
    hols_df = pd.read_excel(xls, sheet_name=hols_sheet) if hols_sheet else pd.DataFrame()
    callh_df = pd.read_excel(xls, sheet_name=callh_sheet) if callh_sheet else pd.DataFrame()
    hleave_df = pd.read_excel(xls, sheet_name=hleave_sheet) if hleave_sheet else pd.DataFrame()
    tph_df = pd.read_excel(xls, sheet_name=tph_sheet) if tph_sheet else pd.DataFrame()
    tbk_df = pd.read_excel(xls, sheet_name=tbk_sheet) if tbk_sheet else pd.DataFrame()
    tweek_df = pd.read_excel(xls, sheet_name=tweek_sheet) if tweek_sheet else pd.DataFrame()
    swaps_df = pd.read_excel(xls, sheet_name=swaps_sheet) if swaps_sheet else pd.DataFrame()
    new_df = pd.read_excel(xls, sheet_name=new_sheet) if new_sheet else pd.DataFrame()

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
        if staff_df[c].dtype == bool:
            return staff_df[c].fillna(default)
        return staff_df[c].apply(yn)

    staff_df["CanFrontDesk"] = bcol(["CanFrontDesk"])
    staff_df["CanTriage"] = bcol(["CanTriage"])
    staff_df["CanEmail"] = bcol(["CanEmail"])
    staff_df["CanPhones"] = bcol(["CanPhones"])
    staff_df["CanBookings"] = bcol(["CanBookings"])
    staff_df["CanEMIS"] = bcol(["CanEMIS"])
    staff_df["CanDocman"] = bcol(["CanDocman_PSA"]) | bcol(["CanDocman_AWAIT"]) | bcol(["CanDocman"])

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
    hs = pick_col(hours_df, ["HomeSite","Site","BaseSite"], required=False)
    hours_df["Name"] = hours_df[hn].astype(str).str.strip()
    hours_df["HomeSite"] = hours_df[hs].astype(str).str.strip().str.upper() if hs else ""

    for dn in ["Mon","Tue","Wed","Thu","Fri"]:
        sc = pick_col(hours_df, [f"{dn}Start", f"{dn} Start", f"{dn}_Start"], required=False)
        ec = pick_col(hours_df, [f"{dn}End", f"{dn} End", f"{dn}_End"], required=False)
        hours_df[f"{dn}Start"] = hours_df[sc].apply(to_time) if sc else None
        hours_df[f"{dn}End"] = hours_df[ec].apply(to_time) if ec else None

    hours_map = {}
    for _, r in hours_df.iterrows():
        hours_map[r["Name"]] = {k: r.get(k) for k in hours_df.columns}

    # --- Holidays ranges
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

    # --- Targets
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

    phones_targets = parse_hourly(tph_df)
    bookings_targets = parse_hourly(tbk_df)

    weekly_targets = {"Bookings": 0.0, "EMIS": 0.0, "Docman": 0.0}
    if tweek_df is not None and not tweek_df.empty:
        task_c = pick_col(tweek_df, ["Task"], required=False) or tweek_df.columns[0]
        val_c = pick_col(tweek_df, ["WeekHoursTarget","Target","Hours"], required=False) or tweek_df.columns[1]
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

    # Buddies
    buddies: Dict[str,str] = {}
    if new_df is not None and not new_df.empty:
        nc = pick_col(new_df, ["NewStarterName","NewStarter","Starter"], required=False) or new_df.columns[0]
        bc = pick_col(new_df, ["BuddyName","Buddy"], required=False) or new_df.columns[1]
        for _, r in new_df.iterrows():
            n = str(r.get(nc,"")).strip()
            b = str(r.get(bc,"")).strip()
            if n and b:
                buddies[n] = b

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
        buddies=buddies,
    )

# ---------- Availability helpers ----------
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

# ---------- Swaps ----------
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

# ---------- Call handler leave impact on phones ----------
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
    base = int(tpl.phones_targets.get((dn, hour_key), 0) or 0)
    if base <= 0:
        return 0

    leave_ranges = parse_handler_leave(tpl.handler_leave)
    off = 0
    if tpl.call_handlers is not None and not tpl.call_handlers.empty:
        for _, r in tpl.call_handlers.iterrows():
            nm = str(r.get("Name","")).strip()
            if not nm:
                continue
            if not handler_working(r, d, t):
                continue
            for ln, sd, ed in leave_ranges:
                if ln.strip().lower() == nm.strip().lower() and sd <= d <= ed:
                    off += 1
                    break
    return base + off

# ---------- Site-of-day ----------
def awaiting_site_for_day(d: date) -> str:
    wd = d.weekday()
    if wd in (0,4):  # Mon/Fri
        return "SLGP"
    if wd in (1,3):  # Tue/Thu
        return "JEN"
    return "BGS"

def email_site_for_day(d: date) -> str:
    # As agreed: same pattern as awaiting (can fall back cross-site only if needed)
    return awaiting_site_for_day(d)

# ---------- Break placement ----------
def pick_breaks_site_balanced(staff_list: List[Staff], hours_map, hols, week_dates: List[date], fixed_assignments: Set[Tuple[date,time,str]]) -> Dict[Tuple[date,time], Set[str]]:
    """
    Breaks only for staff with break_required AND shift > 6h.
    Spread within site by balancing counts per time slot.
    Avoid placing break where staff is locked to fixed assignment (FD/Triage) in that slot.
    """
    breaks: Dict[Tuple[date,time], Set[str]] = {}
    break_load: Dict[Tuple[date,str,time], int] = {}

    slots = timeslots()

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
                # avoid if fixed assigned at that slot
                if (d, bt, st.name) in fixed_assignments:
                    continue

                # avoid tiny fragments
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
    # Default weight = 3 if not present
    return int(st.weights.get(task_key, 3) if st.weights is not None else 3)

def block_limits(task: str) -> Tuple[int,int]:
    if task == "Phones":
        return MIN_PHONES, MAX_PHONES
    if task == "Docman":
        return MIN_DOCMAN, MAX_DEFAULT
    # fixed tasks not here
    return MIN_DEFAULT, MAX_DEFAULT

def schedule_week(tpl: TemplateData, week_start: date):
    slots = timeslots()
    dates = [week_start + timedelta(days=i) for i in range(5)]

    # apply swaps
    hours_map = apply_swaps(tpl.hours_map, tpl.swaps, dates)

    staff_by_name = {s.name: s for s in tpl.staff}
    staff_names = [s.name for s in tpl.staff]

    # Assignments (d,t,name)->task
    a: Dict[Tuple[date,time,str], str] = {}
    gaps: List[Tuple[date,time,str,str]] = []

    # minutes tracking
    mins_task: Dict[Tuple[str,str], int] = {}           # (name, taskKey)->minutes
    mins_task_day: Dict[Tuple[date,str,str], int] = {}  # (d,name,taskKey)->minutes

    def add_mins(d: date, nm: str, task_key: str, mins: int):
        mins_task[(nm, task_key)] = mins_task.get((nm, task_key), 0) + mins
        mins_task_day[(d, nm, task_key)] = mins_task_day.get((d, nm, task_key), 0) + mins

    def assigned(nm: str, d: date, t: time) -> Optional[str]:
        return a.get((d,t,nm))

    def is_free(nm: str, d: date, t: time) -> bool:
        return (d,t,nm) not in a

    # --- Fixed assignments: Front Desk + Triage bands
    fixed_slots: Set[Tuple[date,time,str]] = set()

    def can_cover_full_band(nm: str, d: date, bs: time, be: time) -> bool:
        stt, end = shift_window(hours_map, d, nm)
        if not stt or not end:
            return False
        return (stt <= bs) and (end >= be)

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
            # FrontDeskOnly gets absolute precedence for FD bands
            fd_only = 1 if (task_key == "FrontDesk" and st.frontdesk_only) else 0
            # weight-first, then least used
            w = task_weight(st, task_key)
            used = mins_task.get((nm, task_key), 0)
            return (-fd_only, -w, used, nm.lower())

        ok.sort(key=score)
        return ok[0]

    # Front Desk
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
                    add_mins(d, chosen, "FrontDesk", SLOT_MIN)

    # Triage (SLGP/JEN)
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
                    add_mins(d, chosen, "Triage", SLOT_MIN)

    # --- Breaks (only >6h and BreakRequired)
    breaks = pick_breaks_site_balanced(tpl.staff, hours_map, tpl.hols, dates, fixed_slots)

    def on_break(nm: str, d: date, t: time) -> bool:
        return nm in breaks.get((d,t), set())

    # --- Active blocks for variable tasks
    active: Dict[Tuple[date,str], Tuple[str,int]] = {}  # (d,name)->(task, end_idx_excl)

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
            # someone already assigned (fixed or enforced), stop block
            del active[(d,nm)]
            return False
        a[(d,t,nm)] = task
        add_mins(d, nm, task_key_for_task(task), SLOT_MIN)
        return True

    def stop_block(nm: str, d: date):
        if (d,nm) in active:
            del active[(d,nm)]

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
        if task.startswith("FrontDesk_") or task.startswith("Triage_Admin_"):
            return False  # already handled fixed
        if task == "Email_Box":
            # by-day site preference; if allow_cross_site True, any site allowed
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
        return True

    def start_block(nm: str, task: str, d: date, start_idx: int, allow_short_end: bool=True) -> bool:
        mn, mx = block_limits(task)
        stt, end = shift_window(hours_map, d, nm)
        if not stt or not end:
            return False

        # compute end_idx by shift end or break or fixed assignment boundary
        end_idx = start_idx
        while end_idx < len(slots) and slots[end_idx] < end:
            # stop if a fixed assignment exists at this slot for nm
            if (d, slots[end_idx], nm) in fixed_slots:
                break
            # stop at break start
            if nm in breaks.get((d, slots[end_idx]), set()):
                break
            end_idx += 1

        remaining = end_idx - start_idx
        if remaining <= 0:
            return False

        # no floaters: must be >= mn unless it's end-of-shift remainder (allow_short_end)
        if remaining < mn and not allow_short_end:
            return False

        L = remaining if remaining < mn else min(mx, remaining)
        # if we have enough remaining to meet min, enforce min
        if remaining >= mn:
            L = max(mn, L)

        active[(d,nm)] = (task, start_idx + L)
        return True

    def pick_candidates(task: str, d: date, t: time, allow_cross_site: bool=False, prefer_sites: Optional[List[str]]=None) -> List[str]:
        cands = []
        for nm in staff_names:
            if not is_free(nm,d,t):
                continue
            if not eligible(nm, task, d, t, allow_cross_site=allow_cross_site):
                continue
            cands.append(nm)

        # site preference for EMIS/Docman: JEN/BGS first
        if prefer_sites:
            cands.sort(key=lambda nm: (0 if staff_by_name[nm].home in prefer_sites else 1, nm.lower()))

        def score(nm: str):
            st = staff_by_name[nm]
            key = task_key_for_task(task)
            w = task_weight(st, key)
            used = mins_task.get((nm, key), 0)
            return (-w, used, nm.lower())

        cands.sort(key=score)
        return cands

    def assign_block(nm: str, task: str, d: date, idx: int):
        # if active same task apply
        b = active.get((d,nm))
        if b and b[0] == task:
            apply_active(nm, d, idx)
            return
        if b and b[0] != task:
            stop_block(nm, d)

        # start block
        ok = start_block(nm, task, d, idx, allow_short_end=True)
        if not ok:
            # fallback: if only 1 slot left, DO NOT assign floater; instead mark misc and log
            # We'll avoid assigning anything < min except true end remainder already handled in start_block allow_short_end.
            a[(d, slots[idx], nm)] = "Misc_Tasks"
            add_mins(d, nm, "Misc", SLOT_MIN)
            return
        apply_active(nm, d, idx)

    # --- Weekly targets (minutes)
    target_book = int(round((tpl.weekly_targets.get("Bookings", 0.0) or 0.0) * 60))
    target_emis = int(round((tpl.weekly_targets.get("EMIS", 0.0) or 0.0) * 60))
    target_doc  = int(round((tpl.weekly_targets.get("Docman", 0.0) or 0.0) * 60))

    def total_mins(task_key: str) -> int:
        return sum(v for (nm, tk), v in mins_task.items() if tk == task_key)

    # helper: compute dynamic bookings people-per-slot requirement to approach weekly target
    def bookings_needed_this_slot(d: date, idx: int) -> int:
        if target_book <= 0:
            return 0
        t = slots[idx]
        if t < time(10,30):
            return 0
        done = total_mins("Bookings")
        remaining = max(0, target_book - done)
        if remaining <= 0:
            return 0

        # remaining available slots across week from (d,idx) onward in booking window
        rem_slots = 0
        for dd in dates:
            for tt in slots:
                if dd < d:
                    continue
                if dd == d and tt < t:
                    continue
                if tt >= time(10,30):
                    rem_slots += 1
        # each person contributes 30 mins per slot
        if rem_slots <= 0:
            return 0
        # required people this slot to hit target smoothly
        ppl = math.ceil(remaining / (rem_slots * SLOT_MIN))
        return max(0, ppl)

    def enforce(task: str, need: int, d: date, idx: int, allow_cross_site: bool=False, prefer_sites: Optional[List[str]]=None, note_task_key: str=""):
        t = slots[idx]
        while True:
            current = len([nm for nm in staff_names if a.get((d,t,nm)) == task])
            if current >= need:
                return
            cands = pick_candidates(task, d, t, allow_cross_site=allow_cross_site, prefer_sites=prefer_sites)
            if not cands:
                gaps.append((d, t, task, f"Short by {need-current}"))
                return
            nm = cands[0]
            assign_block(nm, task, d, idx)

    # --- Main loop by slot
    for d in dates:
        for idx, t in enumerate(slots):
            # Skip non-working: handled per eligible check
            # Apply active blocks first for stability
            for nm in staff_names:
                if (d,t,nm) in a:
                    continue
                if on_break(nm,d,t):
                    continue
                apply_active(nm, d, idx)

            # Mandatory: Email (10:30-16) on site-of-day; cross-site only if needed
            if t_in_range(t, time(10,30), time(16,0)):
                enforce("Email_Box", 1, d, idx, allow_cross_site=False)
                # if still not met, allow cross-site
                cur = len([nm for nm in staff_names if a.get((d,t,nm)) == "Email_Box"])
                if cur < 1:
                    enforce("Email_Box", 1, d, idx, allow_cross_site=True)

            # Mandatory: Awaiting (10-16) site-of-day; cross-site only if needed
            if t_in_range(t, time(10,0), time(16,0)):
                enforce("Awaiting_PSA_Admin", 1, d, idx, allow_cross_site=False)
                cur = len([nm for nm in staff_names if a.get((d,t,nm)) == "Awaiting_PSA_Admin"])
                if cur < 1:
                    enforce("Awaiting_PSA_Admin", 1, d, idx, allow_cross_site=True)

            # Phones (hard) all day per matrix
            req_p = phones_required(tpl, d, t)
            if req_p > 0:
                enforce("Phones", req_p, d, idx, allow_cross_site=True)

            # Bookings: dynamic weekly pressure, SLGP first; can use other sites if still short
            need_b = bookings_needed_this_slot(d, idx)
            if need_b > 0:
                # enforce on SLGP first
                enforce("Bookings", need_b, d, idx, allow_cross_site=False)
                cur = len([nm for nm in staff_names if a.get((d,t,nm)) == "Bookings"])
                if cur < need_b:
                    # top-up cross-site (but NEVER steal from phones below req - enforce uses only free staff)
                    enforce("Bookings", need_b, d, idx, allow_cross_site=True)

            # Fill remaining staff with EMIS/Docman until targets met, pref JEN/BGS
            emis_done = total_mins("EMIS")
            doc_done  = total_mins("Docman")

            # decide filler order
            filler_tasks = []
            if target_doc > 0 and doc_done < target_doc:
                filler_tasks.append("Docman")
            if target_emis > 0 and emis_done < target_emis:
                filler_tasks.append("EMIS")
            # if both met, misc
            if not filler_tasks:
                filler_tasks = ["Misc_Tasks"]

            # assign fillers to any remaining free staff, pref site logic
            for nm in staff_names:
                if not is_working(hours_map, d, t, nm):
                    continue
                if holiday_kind(nm, d, tpl.hols):
                    continue
                if on_break(nm, d, t):
                    continue
                if not is_free(nm, d, t):
                    continue

                # choose filler respecting site preference
                chosen = None
                for ft in filler_tasks:
                    if ft in ("EMIS","Docman"):
                        # JEN/BGS first; SLGP only if still needed (handled by prefer_sites ranking)
                        pass
                    if eligible(nm, ft, d, t, allow_cross_site=True):
                        chosen = ft
                        break
                if not chosen:
                    chosen = "Misc_Tasks"

                # for EMIS/Docman, prefer JEN/BGS by candidate picker rather than per-person; but we're here per-person.
                # We'll implement: if task is EMIS/Docman and nm is SLGP and there exists any free JEN/BGS who can do it, give SLGP misc instead.
                if chosen in ("EMIS","Docman") and staff_by_name[nm].home == "SLGP":
                    # look for any other free JEN/BGS candidate for same task
                    any_other = False
                    for other in staff_names:
                        if other == nm:
                            continue
                        if staff_by_name[other].home not in ("JEN","BGS"):
                            continue
                        if not is_free(other, d, t):
                            continue
                        if eligible(other, chosen, d, t, allow_cross_site=True):
                            any_other = True
                            break
                    if any_other:
                        chosen = "Misc_Tasks"

                assign_block(nm, chosen, d, idx)

    # Post-pass: enforce FD exactly one per site per slot (safety)
    for d in dates:
        for t in slots:
            for site in SITES:
                role = f"FrontDesk_{site}"
                assigned_staff = [nm for nm in staff_names if a.get((d,t,nm)) == role]
                if len(assigned_staff) > 1:
                    # keep the highest weight (FrontDesk), then least used
                    def fd_score(nm):
                        st = staff_by_name[nm]
                        w = task_weight(st, "FrontDesk")
                        used = mins_task.get((nm, "FrontDesk"), 0)
                        return (-w, used, nm.lower())
                    assigned_staff.sort(key=fd_score)
                    keep = assigned_staff[0]
                    for extra in assigned_staff[1:]:
                        a[(d,t,extra)] = "Misc_Tasks"
                        gaps.append((d,t,role,f"Removed duplicate FD assignment for {extra}; kept {keep}"))

    return a, breaks, gaps, dates, slots, hours_map

# ---------- Excel output ----------
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
    "Break": "DDDDDD",
    "Holiday": "FFF2CC",
    "Bank Holiday": "FFE599",
    "Sick": "F4CCCC",
    "": "DDDDDD",
}

def fill_for(value: str) -> PatternFill:
    return PatternFill("solid", fgColor=ROLE_COLORS.get(value, "FFFFFF"))

THICK = Side(style="thick")
THIN = Side(style="thin")
DAY_BORDER = Border(top=THICK)
CELL_BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

def build_workbook(tpl: TemplateData, start_monday: date, weeks: int) -> Workbook:
    wb = Workbook()
    wb.remove(wb.active)

    staff_names = [s.name for s in tpl.staff]
    slots = timeslots()

    for w in range(weeks):
        wk_start = start_monday + timedelta(days=7*w)
        a, breaks, gaps, dates, slots, hours_map = schedule_week(tpl, wk_start)

        # Master
        ws = wb.create_sheet(f"Week{w+1}_MasterTimeline")
        ws.append(["Date","Time"] + staff_names)
        for c in ws[1]:
            c.font = Font(bold=True)
            c.alignment = Alignment(horizontal="center", vertical="center")
        ws.freeze_panes = "C2"

        for d in dates:
            for t in slots:
                row = [d.strftime("%a %d-%b"), t.strftime("%H:%M")]
                for nm in staff_names:
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

        # Style
        for rr in range(2, ws.max_row+1):
            if ws.cell(rr,2).value == "08:00":
                for cc in range(1, ws.max_column+1):
                    ws.cell(rr,cc).border = DAY_BORDER
            for cc in range(1, ws.max_column+1):
                cell = ws.cell(rr,cc)
                cell.border = CELL_BORDER
                if cc >= 3:
                    val = str(cell.value or "")
                    cell.fill = fill_for(val)
                    cell.alignment = Alignment(vertical="top", wrap_text=True)

        # Notes/gaps
        ws_g = wb.create_sheet(f"Week{w+1}_NotesAndGaps")
        ws_g.append(["Date","Time","Task","Note"])
        for c in ws_g[1]:
            c.font = Font(bold=True)
        for d, t, task, note in gaps:
            ws_g.append([d.isoformat(), t.strftime("%H:%M") if t else "", task, note])

        # Totals
        ws_tot = wb.create_sheet(f"Week{w+1}_Totals")
        tasks_tot = ["FrontDesk_SLGP","FrontDesk_JEN","FrontDesk_BGS","Triage_Admin_SLGP","Triage_Admin_JEN","Email_Box","Phones","Awaiting_PSA_Admin","Bookings","EMIS","Docman","Misc_Tasks","Break"]
        ws_tot.append(["Name"] + tasks_tot + ["WeeklyTotalHours"])
        for c in ws_tot[1]:
            c.font = Font(bold=True)
            c.alignment = Alignment(horizontal="center", vertical="center")

        header_map = {ws.cell(1,c).value: c for c in range(1, ws.max_column+1)}

        for r_i, nm in enumerate(staff_names, start=2):
            ws_tot.cell(r_i,1).value = nm
            mc = header_map.get(nm)
            col_letter = ws.cell(1, mc).column_letter
            rng = f"'{ws.title}'!{col_letter}2:{col_letter}{ws.max_row}"
            for j, task in enumerate(tasks_tot, start=2):
                ws_tot.cell(r_i,j).value = f'=COUNTIF({rng},"{task}")*0.5'
            start = ws_tot.cell(r_i,2).coordinate
            end = ws_tot.cell(r_i,1+len(tasks_tot)).coordinate
            ws_tot.cell(r_i,2+len(tasks_tot)).value = f"=SUM({start}:{end})"

        ws_tot.append([])
        ws_tot.append(["Weekly Targets (hours)",
                       "Bookings", tpl.weekly_targets.get("Bookings",0.0),
                       "EMIS", tpl.weekly_targets.get("EMIS",0.0),
                       "Docman", tpl.weekly_targets.get("Docman",0.0)])
        ws_tot["A"+str(ws_tot.max_row)].font = Font(bold=True)

    return wb
