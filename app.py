import io
import re
from dataclasses import dataclass
from datetime import datetime, date, time, timedelta
from typing import Dict, List, Tuple, Optional

import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows

# =========================================================
# Rota Generator — v11 Structured Priority (Core Build)
# =========================================================

# ---------------- Password (optional) ----------------
def require_password():
    pw = st.secrets.get("APP_PASSWORD")
    if not pw:
        return True
    if st.session_state.get("authed"):
        return True
    with st.form("login"):
        entered = st.text_input("Password", type="password")
        ok = st.form_submit_button("Log in")
        if ok and entered == pw:
            st.session_state.authed = True
            st.success("Logged in.")
            return True
        if ok:
            st.error("Incorrect password.")
    st.stop()

# ---------------- Helpers ----------------
def normalize(s: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(s).strip().lower())

def find_sheet(xls: pd.ExcelFile, candidates):
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

def pick_col(df: pd.DataFrame, candidates, required=True):
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
        return (datetime(2000, 1, 1) + timedelta(seconds=seconds)).time()
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
    return (datetime(2000, 1, 1, t.hour, t.minute) + timedelta(minutes=mins)).time()

def t_in_range(t: time, a: time, b: time) -> bool:
    return (t >= a) and (t < b)

def ensure_monday(d: date) -> date:
    return d - timedelta(days=d.weekday())

DAY_START = time(8, 0)
DAY_END = time(18, 30)
SLOT_MIN = 30

def timeslots():
    cur = datetime(2000, 1, 1, DAY_START.hour, DAY_START.minute)
    end = datetime(2000, 1, 1, DAY_END.hour, DAY_END.minute)
    out = []
    while cur < end:
        out.append(cur.time())
        cur += timedelta(minutes=SLOT_MIN)
    return out

# Bands (fixed)
FD_BANDS = [(time(8,0), time(11,0)), (time(11,0), time(13,30)), (time(13,30), time(16,0)), (time(16,0), time(18,30))]
TRIAGE_BANDS = [(time(8,0), time(10,30)), (time(10,30), time(13,0)), (time(13,30), time(16,0))]

BREAK_WINDOW = (time(12,0), time(14,0))
BREAK_THRESHOLD_HOURS = 6.0

SITES = ["SLGP", "JEN", "BGS"]

def awaiting_site_for_day(d: date) -> str:
    wd = d.weekday()
    if wd in (0,4):
        return "SLGP"
    if wd in (1,3):
        return "JEN"
    return "BGS"

def day_name(d: date) -> str:
    return ["Mon","Tue","Wed","Thu","Fri","Sat","Sun"][d.weekday()]

def holiday_kind(name: str, d: date, hols):
    for n, s, e, k in hols:
        if n.strip().lower() == name.strip().lower() and s and e and s <= d <= e:
            return k
    return None

def yn(v):
    if pd.isna(v):
        return False
    s = str(v).strip().lower()
    return s in ["y","yes","true","1","t"]

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

def read_template(uploaded_bytes: bytes):
    xls = pd.ExcelFile(io.BytesIO(uploaded_bytes))

    staff_sheet = find_sheet(xls, ["Staff"])
    hours_sheet = find_sheet(xls, ["WorkingHours", "Hours"])
    hols_sheet = find_sheet(xls, ["Holidays", "Leave", "Absence"])
    callh_sheet = find_sheet(xls, ["CallHandlers", "Call Handlers"])
    tphones_sheet = find_sheet(xls, ["Targets_Phones_Hourly"])
    tweekly_sheet = find_sheet(xls, ["Targets_Weekly"])
    swaps_sheet = find_sheet(xls, ["Swaps"])
    new_sheet = find_sheet(xls, ["NewStarters"])

    if not staff_sheet or not hours_sheet:
        raise ValueError(f"Missing Staff/WorkingHours sheets. Found: {xls.sheet_names}")

    staff_df = pd.read_excel(xls, sheet_name=staff_sheet)
    hours_df = pd.read_excel(xls, sheet_name=hours_sheet)
    hols_df = pd.read_excel(xls, sheet_name=hols_sheet) if hols_sheet else pd.DataFrame()
    callh_df = pd.read_excel(xls, sheet_name=callh_sheet) if callh_sheet else pd.DataFrame()
    phones_df = pd.read_excel(xls, sheet_name=tphones_sheet) if tphones_sheet else pd.DataFrame()
    weekly_df = pd.read_excel(xls, sheet_name=tweekly_sheet) if tweekly_sheet else pd.DataFrame()
    swaps_df = pd.read_excel(xls, sheet_name=swaps_sheet) if swaps_sheet else pd.DataFrame()
    new_df = pd.read_excel(xls, sheet_name=new_sheet) if new_sheet else pd.DataFrame()

    # Staff parsing
    name_c = pick_col(staff_df, ["Name", "StaffName"])
    home_c = pick_col(staff_df, ["HomeSite", "Site", "BaseSite"], required=False)
    staff_df = staff_df.copy()
    staff_df["Name"] = staff_df[name_c].astype(str).str.strip()
    staff_df["HomeSite"] = staff_df[home_c].astype(str).str.strip().str.upper() if home_c else ""

    def col_bool(df, candidates):
        c = pick_col(df, candidates, required=False)
        if not c:
            return pd.Series([False]*len(df))
        return df[c].apply(yn)

    staff_df["CanFrontDesk"] = col_bool(staff_df, ["CanFrontDesk"])
    staff_df["CanTriage"] = col_bool(staff_df, ["CanTriage"])
    staff_df["CanEmail"] = col_bool(staff_df, ["CanEmail"])
    staff_df["CanPhones"] = col_bool(staff_df, ["CanPhones"])
    staff_df["CanBookings"] = col_bool(staff_df, ["CanBookings"])
    staff_df["CanEMIS"] = col_bool(staff_df, ["CanEMIS"])
    staff_df["CanDocman"] = col_bool(staff_df, ["CanDocman_PSA"]) | col_bool(staff_df, ["CanDocman_AWAIT"])

    staff_list: List[Staff] = []
    for _, r in staff_df.iterrows():
        staff_list.append(Staff(
            name=str(r["Name"]).strip(),
            home=str(r.get("HomeSite","")).strip().upper(),
            can_frontdesk=bool(r.get("CanFrontDesk",False)),
            can_triage=bool(r.get("CanTriage",False)),
            can_email=bool(r.get("CanEmail",False)),
            can_phones=bool(r.get("CanPhones",False)),
            can_bookings=bool(r.get("CanBookings",False)),
            can_emis=bool(r.get("CanEMIS",False)),
            can_docman=bool(r.get("CanDocman",False)),
        ))

    # Working hours parsing
    hours_df = hours_df.copy()
    hours_name_c = pick_col(hours_df, ["Name","StaffName"])
    hours_df["Name"] = hours_df[hours_name_c].astype(str).str.strip()
    for dn in ["Mon","Tue","Wed","Thu","Fri"]:
        sc = pick_col(hours_df, [f"{dn}Start", f"{dn} Start", f"{dn}_Start"], required=False)
        ec = pick_col(hours_df, [f"{dn}End", f"{dn} End", f"{dn}_End"], required=False)
        hours_df[f"{dn}Start"] = hours_df[sc].apply(to_time) if sc else None
        hours_df[f"{dn}End"] = hours_df[ec].apply(to_time) if ec else None
    hours_map = {r["Name"]: r for _, r in hours_df.iterrows()}

    # Holidays as ranges
    hols = []
    if not hols_df.empty:
        hn = pick_col(hols_df, ["Name","StaffName"], required=False) or hols_df.columns[0]
        hs = pick_col(hols_df, ["StartDate","Start"], required=False) or hols_df.columns[1]
        he = pick_col(hols_df, ["EndDate","End"], required=False) or hols_df.columns[2]
        notes_c = pick_col(hols_df, ["Notes","Note","Reason"], required=False)
        for _, r in hols_df.iterrows():
            nm = str(r[hn]).strip()
            sd = to_date(r[hs])
            ed = to_date(r[he])
            note = "" if (not notes_c or pd.isna(r[notes_c])) else str(r[notes_c]).strip().lower()
            kind = "Holiday"
            if "sick" in note or "sickness" in note:
                kind = "Sick"
            elif "bank" in note:
                kind = "Bank Holiday"
            hols.append((nm, sd, ed, kind))

    # Call handlers (hours + leave)
    call_handlers = []
    if not callh_df.empty:
        ncol = pick_col(callh_df, ["Name","HandlerName","CallHandler","Call Handler"], required=False) or callh_df.columns[0]
        callh_df = callh_df.copy()
        callh_df["Name"] = callh_df[ncol].astype(str).str.strip()
        for dn in ["Mon","Tue","Wed","Thu","Fri"]:
            sc = pick_col(callh_df, [f"{dn}Start"], required=False)
            ec = pick_col(callh_df, [f"{dn}End"], required=False)
            callh_df[f"{dn}Start"] = callh_df[sc].apply(to_time) if sc else None
            callh_df[f"{dn}End"] = callh_df[ec].apply(to_time) if ec else None
        ls = pick_col(callh_df, ["LeaveStartDate(optional)","LeaveStart","Leave Start"], required=False)
        le = pick_col(callh_df, ["LeaveEndDate(optional)","LeaveEnd","Leave End"], required=False)
        callh_df["LeaveStart"] = callh_df[ls].apply(to_date) if ls else None
        callh_df["LeaveEnd"] = callh_df[le].apply(to_date) if le else None
        for _, r in callh_df.iterrows():
            call_handlers.append(r)

    # Phones hourly targets
    phones_targets = {}  # (weekday, hour_time)->int
    if not phones_df.empty:
        time_col = pick_col(phones_df, ["Time"], required=False) or phones_df.columns[0]
        phones_df = phones_df.copy()
        phones_df["Time"] = phones_df[time_col].astype(str).str.strip()
        for dn in ["Mon","Tue","Wed","Thu","Fri"]:
            if dn in phones_df.columns:
                for _, r in phones_df.iterrows():
                    hh = to_time(r["Time"])
                    if not hh:
                        continue
                    val = r.get(dn, 0)
                    if pd.isna(val) or val == "":
                        continue
                    phones_targets[(dn, hh)] = int(float(val))

    # Weekly targets
    weekly_targets = {"Bookings": None, "EMIS": None, "Docman": None}
    if not weekly_df.empty:
        task_c = pick_col(weekly_df, ["Task"], required=False) or weekly_df.columns[0]
        val_c = pick_col(weekly_df, ["WeekHoursTarget","Target","Hours"], required=False) or weekly_df.columns[1]
        for _, r in weekly_df.iterrows():
            tsk = str(r[task_c]).strip()
            val = r[val_c]
            if pd.isna(val) or val == "":
                continue
            weekly_targets[tsk] = float(val)

    # Swaps
    swaps = []
    if not swaps_df.empty:
        dcol = pick_col(swaps_df, ["Date"], required=False) or swaps_df.columns[0]
        ncol = pick_col(swaps_df, ["Name"], required=False) or swaps_df.columns[1]
        swcol = pick_col(swaps_df, ["SwapWith (OPTION A)","SwapWith","Swap With"], required=False)
        nscol = pick_col(swaps_df, ["NewStart (OPTION B)","NewStart","New Start"], required=False)
        necol = pick_col(swaps_df, ["NewEnd (OPTION B)","NewEnd","New End"], required=False)
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

    # Buddy map
    buddies = {}
    if not new_df.empty:
        nc = pick_col(new_df, ["NewStarterName","NewStarter","Starter"], required=False) or new_df.columns[0]
        bc = pick_col(new_df, ["BuddyName","Buddy"], required=False) or new_df.columns[1]
        for _, r in new_df.iterrows():
            n = str(r.get(nc,"")).strip()
            b = str(r.get(bc,"")).strip()
            if n and b:
                buddies[n] = b

    return staff_list, hours_map, hols, call_handlers, phones_targets, weekly_targets, swaps, buddies

def shift_window(hours_map, week_start: date, d: date, name: str):
    dn = day_name(d)
    hr = hours_map.get(name)
    if hr is None:
        return None, None
    return hr.get(f"{dn}Start"), hr.get(f"{dn}End")

def is_working(hours_map, week_start: date, d: date, t: time, name: str):
    stt, end = shift_window(hours_map, week_start, d, name)
    return bool(stt and end and (t >= stt) and (t < end))

def handler_working(callh_row, d: date, t: time):
    ls = callh_row.get("LeaveStart")
    le = callh_row.get("LeaveEnd")
    if ls and le and ls <= d <= le:
        return False
    dn = day_name(d)
    stt = callh_row.get(f"{dn}Start")
    end = callh_row.get(f"{dn}End")
    return bool(stt and end and (t >= stt) and (t < end))

def phones_required(phones_targets, d: date, t: time, call_handlers):
    dn = day_name(d)
    hour_key = time(t.hour, 0)  # applies to both half-hours
    base = phones_targets.get((dn, hour_key), 0)
    off = 0
    for r in call_handlers:
        if not handler_working(r, d, t):
            off += 1
    return int(base + off)

def pick_breaks_site_balanced(staff_list, hours_map, hols, week_start: date, dates: List[date]):
    breaks = {}
    break_load = {}  # (d, site, t)->count
    candidate_times = [time(12,0), time(12,30), time(13,0), time(13,30)]
    for d in dates:
        for st in staff_list:
            if holiday_kind(st.name, d, hols):
                continue
            stt, end = shift_window(hours_map, week_start, d, st.name)
            if not stt or not end:
                continue
            dur = (dt_of(d, end) - dt_of(d, stt)).total_seconds()/3600
            if dur <= BREAK_THRESHOLD_HOURS:
                continue
            midpoint = dt_of(d, stt) + (dt_of(d, end) - dt_of(d, stt))/2
            best = None
            for bt in candidate_times:
                if bt < stt or add_minutes(bt, 30) > end:
                    continue
                if not t_in_range(bt, BREAK_WINDOW[0], BREAK_WINDOW[1]):
                    continue
                load = break_load.get((d, st.home, bt), 0)
                dist = abs((dt_of(d, bt) - midpoint).total_seconds())
                score = (load * 3600) + dist
                if best is None or score < best[0]:
                    best = (score, bt)
            if best:
                bt = best[1]
                breaks.setdefault((d, bt), set()).add(st.name)
                break_load[(d, st.home, bt)] = break_load.get((d, st.home, bt), 0) + 1
    return breaks

def schedule_week(staff_list, hours_map, hols, call_handlers, phones_targets, weekly_targets, swaps, buddies, week_start: date):
    slots = timeslots()
    dates = [week_start + timedelta(days=i) for i in range(5)]
    hours_map2 = {k: dict(v) for k, v in hours_map.items()}

    # Apply swaps
    for dd, nm, sw, ns, ne in swaps:
        if dd < week_start or dd > week_start + timedelta(days=4):
            continue
        dn = day_name(dd)
        if nm not in hours_map2:
            continue
        if sw and sw in hours_map2:
            a = hours_map2[nm].get(f"{dn}Start"), hours_map2[nm].get(f"{dn}End")
            b = hours_map2[sw].get(f"{dn}Start"), hours_map2[sw].get(f"{dn}End")
            hours_map2[nm][f"{dn}Start"], hours_map2[nm][f"{dn}End"] = b
            hours_map2[sw][f"{dn}Start"], hours_map2[sw][f"{dn}End"] = a
        elif ns and ne:
            hours_map2[nm][f"{dn}Start"] = ns
            hours_map2[nm][f"{dn}End"] = ne

    breaks = pick_breaks_site_balanced(staff_list, hours_map2, hols, week_start, dates)

    a: Dict[Tuple[date,time,str], str] = {}
    gaps: List[Tuple[date,time,str,str]] = []
    active: Dict[Tuple[date,str], Tuple[str,int]] = {}
    fd_bands_done: Dict[Tuple[date,str], int] = {}
    task_minutes: Dict[Tuple[str,str], int] = {}

    staff_by_name = {s.name: s for s in staff_list}

    def on_break(name, d, t):
        return name in breaks.get((d,t), set())

    def is_free(name, d, t):
        return (d,t,name) not in a

    def staff_on_task(task, d, t):
        return [nm for (dd,tt,nm), rr in a.items() if dd==d and tt==t and rr==task]

    def start_block(name, task, d, start_idx, min_slots, max_slots):
        stt, end = shift_window(hours_map2, week_start, d, name)
        if not stt or not end:
            return False
        end_idx = start_idx
        while end_idx < len(slots) and slots[end_idx] < end:
            end_idx += 1
        for k in range(start_idx, min(end_idx, len(slots))):
            if name in breaks.get((d, slots[k]), set()):
                end_idx = k
                break
        remaining = end_idx - start_idx
        if remaining <= 0:
            return False
        if remaining < min_slots:
            L = remaining
        else:
            L = min(max_slots, remaining)
            if L < min_slots and remaining >= min_slots:
                L = min_slots
        active[(d,name)] = (task, start_idx + L)
        return True

    def apply_active(name, d, idx):
        b = active.get((d,name))
        if not b:
            return False
        task, end_idx = b
        if idx >= end_idx:
            del active[(d,name)]
            return False
        t = slots[idx]
        a[(d,t,name)] = task
        task_minutes[(name, task)] = task_minutes.get((name, task), 0) + SLOT_MIN
        return True

    def can_do(name, task, d, t):
        st = staff_by_name[name]
        if holiday_kind(name, d, hols):
            return False
        if not is_working(hours_map2, week_start, d, t, name):
            return False
        if on_break(name, d, t):
            return False
        if task.startswith("FrontDesk_"):
            site = task.split("_",1)[1]
            return st.can_frontdesk and st.home == site
        if task.startswith("Triage_Admin_"):
            site = task.split("_")[-1]
            return st.can_triage and st.home == site
        if task == "Email_Box":
            return st.can_email and st.home in ("JEN","BGS")
        if task == "Phones":
            return st.can_phones
        if task == "Awaiting_PSA_Admin":
            return st.can_docman and st.home == awaiting_site_for_day(d)
        if task == "Bookings":
            return st.can_bookings and st.home == "SLGP"
        if task == "EMIS":
            return st.can_emis
        if task == "Docman":
            return st.can_docman
        return True

    def pick_candidate(task, d, idx, t):
        free = []
        for name in staff_by_name.keys():
            if not is_free(name, d, t):
                continue
            if not can_do(name, task, d, t):
                continue
            free.append(name)
        free.sort(key=lambda nm: task_minutes.get((nm, task), 0))
        return free[0] if free else None

    def enforce_frontdesk_band(d, band_start, band_end, site):
        task = f"FrontDesk_{site}"
        candidates = []
        for st in staff_list:
            if st.home != site or not st.can_frontdesk:
                continue
            ok = True
            for tt in slots:
                if tt < band_start or tt >= band_end:
                    continue
                if not is_working(hours_map2, week_start, d, tt, st.name):
                    ok = False; break
                if on_break(st.name, d, tt):
                    ok = False; break
            if not ok:
                continue
            bands = fd_bands_done.get((d, st.name), 0)
            candidates.append((bands, task_minutes.get((st.name, task), 0), st.name))
        candidates.sort()
        if not candidates:
            gaps.append((d, band_start, task, "No suitable staff for front desk band"))
            return
        chosen = candidates[0][2]
        if fd_bands_done.get((d, chosen), 0) >= 1:
            for bands, _, nm in candidates:
                if bands == 0:
                    chosen = nm
                    break
        fd_bands_done[(d, chosen)] = fd_bands_done.get((d, chosen), 0) + 1
        for tt in slots:
            if tt < band_start or tt >= band_end:
                continue
            a[(d,tt,chosen)] = task
            task_minutes[(chosen, task)] = task_minutes.get((chosen, task), 0) + SLOT_MIN

    def enforce_triage_band(d, band_start, band_end, site):
        task = f"Triage_Admin_{site}"
        candidates = []
        for st in staff_list:
            if st.home != site or not st.can_triage:
                continue
            ok = True
            for tt in slots:
                if tt < band_start or tt >= band_end:
                    continue
                if not is_working(hours_map2, week_start, d, tt, st.name):
                    ok = False; break
                if on_break(st.name, d, tt):
                    ok = False; break
            if not ok:
                continue
            candidates.append((task_minutes.get((st.name, task), 0), st.name))
        candidates.sort()
        if not candidates:
            gaps.append((d, band_start, task, "No suitable staff for triage band"))
            return
        chosen = candidates[0][1]
        for tt in slots:
            if tt < band_start or tt >= band_end:
                continue
            a[(d,tt,chosen)] = task
            task_minutes[(chosen, task)] = task_minutes.get((chosen, task), 0) + SLOT_MIN

    # Phase 1 lock FD + triage
    for d in dates:
        for site in SITES:
            for bs, be in FD_BANDS:
                enforce_frontdesk_band(d, bs, be, site)
        for site in ("SLGP","JEN"):
            for bs, be in TRIAGE_BANDS:
                enforce_triage_band(d, bs, be, site)

    # Phase 2 slot fill
    for d in dates:
        for idx, t in enumerate(slots):
            # Email 10:30-16:00
            if t_in_range(t, time(10,30), time(16,0)):
                if len(staff_on_task("Email_Box", d, t)) < 1:
                    cand = pick_candidate("Email_Box", d, idx, t)
                    if cand:
                        if (d,cand) not in active:
                            start_block(cand, "Email_Box", d, idx, min_slots=5, max_slots=8)
                        apply_active(cand, d, idx)

            # Awaiting 10-16
            if t_in_range(t, time(10,0), time(16,0)):
                if len(staff_on_task("Awaiting_PSA_Admin", d, t)) < 1:
                    cand = pick_candidate("Awaiting_PSA_Admin", d, idx, t)
                    if cand:
                        if (d,cand) not in active:
                            start_block(cand, "Awaiting_PSA_Admin", d, idx, min_slots=4, max_slots=6)
                        apply_active(cand, d, idx)
                    else:
                        gaps.append((d,t,"Awaiting_PSA_Admin","No awaiting cover candidate"))

            # Phones hard
            req = phones_required(phones_targets, d, t, call_handlers)
            cur = len(staff_on_task("Phones", d, t))
            while cur < req:
                cand = pick_candidate("Phones", d, idx, t)
                if not cand:
                    gaps.append((d,t,"Phones",f"Short by {req-cur}"))
                    break
                if (d,cand) not in active:
                    start_block(cand, "Phones", d, idx, min_slots=5, max_slots=8)
                apply_active(cand, d, idx)
                cur = len(staff_on_task("Phones", d, t))

            # Bookings (from 10:30)
            if t_in_range(t, time(10,30), DAY_END):
                if len(staff_on_task("Bookings", d, t)) < 1:
                    cand = pick_candidate("Bookings", d, idx, t)
                    if cand:
                        if (d,cand) not in active:
                            start_block(cand, "Bookings", d, idx, min_slots=5, max_slots=8)
                        apply_active(cand, d, idx)

            # Filler EMIS / Docman
            emis_target = weekly_targets.get("EMIS") or 0
            doc_target = weekly_targets.get("Docman") or 0
            emis_done = sum(v for (nm, task), v in task_minutes.items() if task == "EMIS")/60.0
            doc_done = sum(v for (nm, task), v in task_minutes.items() if task == "Docman")/60.0
            prefer = "EMIS" if (emis_target and emis_done < emis_target and (emis_target-emis_done) >= (doc_target-doc_done if doc_target else 0)) else "Docman"

            for st in staff_list:
                if not is_working(hours_map2, week_start, d, t, st.name):
                    continue
                if holiday_kind(st.name, d, hols):
                    continue
                if on_break(st.name, d, t):
                    continue
                if not is_free(st.name, d, t):
                    continue

                chosen = None
                if prefer == "EMIS" and st.can_emis:
                    chosen = "EMIS"
                elif prefer == "Docman" and st.can_docman:
                    chosen = "Docman"
                else:
                    if st.can_emis:
                        chosen = "EMIS"
                    elif st.can_docman:
                        chosen = "Docman"
                    elif st.can_bookings and t_in_range(t, time(10,30), DAY_END):
                        chosen = "Bookings"
                    elif st.can_phones:
                        chosen = "Phones"
                    else:
                        chosen = "Unassigned"

                if chosen == "Docman":
                    min_slots, max_slots = 6, 9
                elif chosen == "EMIS":
                    min_slots, max_slots = 4, 9
                elif chosen in ("Phones","Bookings"):
                    min_slots, max_slots = 5, 8
                else:
                    min_slots, max_slots = 2, 4

                if chosen == "Unassigned":
                    a[(d,t,st.name)] = "Unassigned"
                else:
                    if (d, st.name) not in active:
                        start_block(st.name, chosen, d, idx, min_slots=min_slots, max_slots=max_slots)
                    apply_active(st.name, d, idx)

    return a, breaks, gaps, dates, slots, hours_map2

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
    "Unassigned": "FFFFFF",
    "Break": "CCCCCC",
    "Holiday": "FFF2CC",
    "Bank Holiday": "FFE599",
    "Sick": "F4CCCC",
    "": "DDDDDD",
}

def fill_for(value: str):
    return PatternFill("solid", fgColor=ROLE_COLORS.get(value, "FFFFFF"))

THICK = Side(style="thick")
THIN = Side(style="thin")
DAY_BORDER = Border(top=THICK)
CELL_BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

def build_workbook(staff_list, a, breaks, gaps, dates, slots, hols, hours_map2, week_start, week_num: int):
    wb = Workbook()
    wb.remove(wb.active)
    staff_names = [s.name for s in staff_list]

    ws = wb.create_sheet(f"Week{week_num}_MasterTimeline")
    ws.append(["Date","Time"] + staff_names)
    for c in ws[1]:
        c.font = Font(bold=True)
        c.alignment = Alignment(horizontal="center", vertical="center")
    ws.freeze_panes = "C2"

    for d in dates:
        for t in slots:
            row = [d.strftime("%a %d-%b"), t.strftime("%H:%M")]
            for nm in staff_names:
                hk = holiday_kind(nm, d, hols)
                if hk:
                    val = hk
                elif not is_working(hours_map2, week_start, d, t, nm):
                    val = ""
                elif nm in breaks.get((d,t), set()):
                    val = "Break"
                else:
                    val = a.get((d,t,nm), "Unassigned")
                row.append(val)
            ws.append(row)

    for rr in range(2, ws.max_row+1):
        if ws.cell(rr,2).value == "08:00":
            for cc in range(1, ws.max_column+1):
                ws.cell(rr,cc).border = DAY_BORDER
        for cc in range(1, ws.max_column+1):
            ws.cell(rr,cc).border = CELL_BORDER
            if cc >= 3:
                val = str(ws.cell(rr,cc).value or "")
                ws.cell(rr,cc).fill = fill_for(val)
                ws.cell(rr,cc).alignment = Alignment(vertical="top", wrap_text=True)

    ws2 = wb.create_sheet(f"Week{week_num}_CoverageAtAGlance")
    tasks = [
        "FrontDesk_SLGP","FrontDesk_JEN","FrontDesk_BGS",
        "Triage_Admin_SLGP","Triage_Admin_JEN",
        "Email_Box","Phones","Awaiting_PSA_Admin","Bookings","EMIS","Docman"
    ]
    ws2.append(["Date","Time"] + tasks)
    for c in ws2[1]:
        c.font = Font(bold=True)
        c.alignment = Alignment(horizontal="center", vertical="center")
    ws2.freeze_panes = "C2"
    for d in dates:
        for t in slots:
            row = [d.strftime("%a %d-%b"), t.strftime("%H:%M")]
            for task in tasks:
                ppl = [nm for nm in staff_names if a.get((d,t,nm)) == task]
                row.append(", ".join(ppl))
            ws2.append(row)
    for rr in range(2, ws2.max_row+1):
        if ws2.cell(rr,2).value == "08:00":
            for cc in range(1, ws2.max_column+1):
                ws2.cell(rr,cc).border = DAY_BORDER
        for cc in range(1, ws2.max_column+1):
            ws2.cell(rr,cc).border = CELL_BORDER
            ws2.cell(rr,cc).alignment = Alignment(vertical="top", wrap_text=True)

    ws3 = wb.create_sheet(f"Week{week_num}_Gaps")
    ws3.append(["Date","Time","Task","Issue"])
    for c in ws3[1]:
        c.font = Font(bold=True)
        c.alignment = Alignment(horizontal="center", vertical="center")
    for d,t,task,issue in gaps:
        ws3.append([d.isoformat(), "" if t is None else t.strftime("%H:%M"), task, issue])

    return wb

# ---------------- Streamlit UI ----------------
st.set_page_config(page_title="Rota Generator v11 Structured Priority", layout="wide")
require_password()

st.title("Rota Generator v11 — Structured Priority (Core)")

uploaded = st.file_uploader("Upload your v11 template (.xlsx)", type=["xlsx"])
c1, c2 = st.columns(2)
with c1:
    start_date = st.date_input("Week commencing (Monday)", value=date.today())
with c2:
    weeks = int(st.number_input("Number of weeks", min_value=1, max_value=12, value=1, step=1))

start_monday = ensure_monday(start_date)

if uploaded:
    try:
        staff_list, hours_map, hols, call_handlers, phones_targets, weekly_targets, swaps, buddies = read_template(uploaded.getvalue())
        st.success(f"Loaded. Staff={len(staff_list)} | CallHandlers={len(call_handlers)} | PhonesMatrixRows={len(phones_targets)}")
        st.write("Weekly targets:", weekly_targets)

        if st.button("Generate rota and download Excel", type="primary"):
            out_wb = Workbook()
            out_wb.remove(out_wb.active)

            for w in range(weeks):
                wk_start = start_monday + timedelta(days=7*w)
                a, breaks, gaps, dates, slots, hours_map2 = schedule_week(
                    staff_list, hours_map, hols, call_handlers, phones_targets,
                    weekly_targets, swaps, buddies, wk_start
                )
                wb_w = build_workbook(staff_list, a, breaks, gaps, dates, slots, hols, hours_map2, wk_start, w+1)
                for sh in wb_w.worksheets:
                    out_wb._add_sheet(sh)

            bio = io.BytesIO()
            out_wb.save(bio)
            bio.seek(0)
            out_name = f"rota_v11_{start_monday.isoformat()}_{weeks}w.xlsx"
            st.download_button(
                "📊 Download Excel rota",
                data=bio.getvalue(),
                file_name=out_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
    except Exception as e:
        st.error("Could not process the template.")
        st.exception(e)
else:
    st.info("Upload your completed template to generate the rota.")
