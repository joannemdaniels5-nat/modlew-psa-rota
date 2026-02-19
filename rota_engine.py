# -*- coding: utf-8 -*-
"""
ModLew PSA Rota Engine — v31 (Strict Priority, Editable Site Timelines, Dynamic Coverage+Totals)

Key behaviours (as requested):
1) Front Desk coverage – absolute first priority – 1 per site, every slot, no gaps (band-locked, shiftable ±30m).
2) Triage Admin – weight-based, band-locked, no misc leakage.
3) Phones – REQUIRED hourly matrix enforced (applies to both 30-min slots within the hour).
4) Email / Awaiting – site-of-day enforced where possible (fallback cross-site if no eligible staff).
5) EMIS / Docman – hard weekly targets (EMIS=20h, Docman=14h by default), site preference JEN+BGS above SLGP,
   block size 2–4h, and max 1 person on each at any slot.
6) Bookings – SLGP first, then BGS/JEN if needed; used to push towards weekly target if provided.
7) Misc – only when everything else is satisfied.

Excel:
- Site timelines contain ONLY that site’s staff columns (editable).
- Totals + Coverage sheets are formula-driven off Site timelines (so edits propagate).
- Colours are applied via conditional formatting (so edits recolour on open).
- Day sections have repeated name headers for print readability.
"""
from __future__ import annotations

import io
import re
import math
from dataclasses import dataclass
from datetime import datetime, date, time, timedelta
from typing import Dict, List, Tuple, Optional, Set

import pandas as pd
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.formatting.rule import FormulaRule, DataBarRule

# ---------------- Core time config ----------------
DAY_START = time(8, 0)
DAY_END = time(18, 30)
SLOT_MIN = 30

SITES = ["SLGP", "JEN", "BGS"]

# Fixed bands (can shift ±30m to handle breaks / short overlaps)
FD_BANDS = [(time(8,0), time(11,0)), (time(11,0), time(13,30)), (time(13,30), time(16,0)), (time(16,0), time(18,30))]
TRIAGE_BANDS = [(time(8,0), time(10,30)), (time(10,30), time(13,0)), (time(13,30), time(16,0))]

# Breaks
BREAK_WINDOW = (time(12,0), time(14,0))
BREAK_CANDIDATES = [time(12,0), time(12,30), time(13,0), time(13,30)]
BREAK_THRESHOLD_HOURS = 6.0

# Block sizes (slots)
MIN_DEFAULT = 5   # 2.5h
MAX_DEFAULT = 9   # 4.5h
MIN_PHONES  = 3   # 1.5h
MAX_PHONES  = 8   # 4h
MIN_ADMIN   = 4   # 2h (EMIS/Docman)
MAX_ADMIN   = 8   # 4h (EMIS/Docman)

# Hard weekly targets (hours) for EMIS/Docman unless changed in code
HARD_EMIS_HOURS = 20.0
HARD_DOC_HOURS  = 14.0

ROLE_COLORS = {
    "FrontDesk_SLGP": "FFF2CC",
    "FrontDesk_JEN":  "FFF2CC",
    "FrontDesk_BGS":  "FFF2CC",
    "Triage_Admin_SLGP": "D9EAD3",
    "Triage_Admin_JEN":  "D9EAD3",
    "Email_Box": "CFE2F3",
    "Awaiting_PSA_Admin": "D0E0E3",
    "Phones": "C9DAF8",
    "Bookings": "FCE5CD",
    "EMIS": "EAD1DC",
    "Docman": "D0E0E3",
    "Misc_Tasks": "FFFFFF",
    "Break": "DDDDDD",
    "Holiday": "FFF2CC",
    "Bank Holiday": "FFE599",
    "Sick": "F4CCCC",
    "": "DDDDDD",
}

THICK = Side(style="thick")
THIN  = Side(style="thin")
CELL_BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

def timeslots() -> List[time]:
    cur = datetime(2000,1,1,DAY_START.hour,DAY_START.minute)
    end = datetime(2000,1,1,DAY_END.hour,DAY_END.minute)
    out=[]
    while cur < end:
        out.append(cur.time())
        cur += timedelta(minutes=SLOT_MIN)
    return out

def ensure_monday(d: date) -> date:
    return d - timedelta(days=d.weekday())

def day_name(d: date) -> str:
    return ["Mon","Tue","Wed","Thu","Fri","Sat","Sun"][d.weekday()]

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
    # UK templates often dd/mm/yyyy
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

# ---------------- Data models ----------------
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
    phones_targets: Dict[Tuple[str, time], int]   # (Mon, 08:00 hour) -> required call handlers
    bookings_targets: Dict[Tuple[str, time], int] # optional hourly matrix (if present)
    weekly_targets: Dict[str, float]              # optional (Bookings can be used)
    swaps: List[Tuple[date, str, str, Optional[time], Optional[time]]]

# ---------------- Site-of-day rules ----------------
def awaiting_site_for_day(d: date) -> str:
    wd = d.weekday()
    if wd in (0,4):  # Mon/Fri
        return "SLGP"
    if wd in (1,3):  # Tue/Thu
        return "JEN"
    return "BGS"

def email_site_for_day(d: date) -> str:
    return awaiting_site_for_day(d)

# ---------------- Template reader ----------------
def read_template(uploaded_bytes: bytes) -> TemplateData:
    xls = pd.ExcelFile(io.BytesIO(uploaded_bytes))

    staff_sheet = find_sheet(xls, ["Staff"])
    hours_sheet = find_sheet(xls, ["WorkingHours","Hours"])
    hols_sheet  = find_sheet(xls, ["Holidays","Leave","Absence"])
    tph_sheet   = find_sheet(xls, ["Targets_Phones_Hourly","PhonesTargets","Targets Phones Hourly"])

    # Optional
    tbk_sheet   = find_sheet(xls, ["Targets_Bookings_Hourly","BookingsTargets","Targets Bookings Hourly"])
    tweek_sheet = find_sheet(xls, ["Targets_Weekly","Targets Weekly"])
    swaps_sheet = find_sheet(xls, ["Swaps"])

    if not staff_sheet or not hours_sheet:
        raise ValueError(f"Missing Staff/WorkingHours sheets. Found: {xls.sheet_names}")

    if not tph_sheet:
        raise ValueError("Phone hourly target sheet is REQUIRED (Targets_Phones_Hourly).")

    staff_df = pd.read_excel(xls, sheet_name=staff_sheet)
    hours_df = pd.read_excel(xls, sheet_name=hours_sheet)
    hols_df  = pd.read_excel(xls, sheet_name=hols_sheet) if hols_sheet else pd.DataFrame()
    tph_df   = pd.read_excel(xls, sheet_name=tph_sheet)
    tbk_df   = pd.read_excel(xls, sheet_name=tbk_sheet) if tbk_sheet else pd.DataFrame()
    tweek_df = pd.read_excel(xls, sheet_name=tweek_sheet) if tweek_sheet else pd.DataFrame()
    swaps_df = pd.read_excel(xls, sheet_name=swaps_sheet) if swaps_sheet else pd.DataFrame()

    # Staff
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
        weights={}
        for k,col in weight_cols.items():
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

    # Working hours
    hours_df = hours_df.copy()
    hn = pick_col(hours_df, ["Name","StaffName"])
    hours_df["Name"] = hours_df[hn].astype(str).str.strip()
    for dn in ["Mon","Tue","Wed","Thu","Fri"]:
        sc = pick_col(hours_df, [f"{dn}Start", f"{dn} Start", f"{dn}_Start"], required=False)
        ec = pick_col(hours_df, [f"{dn}End", f"{dn} End", f"{dn}_End"], required=False)
        hours_df[f"{dn}Start"] = hours_df[sc].apply(to_time) if sc else None
        hours_df[f"{dn}End"]   = hours_df[ec].apply(to_time) if ec else None
    hours_map = {r["Name"]: {k: r.get(k) for k in hours_df.columns} for _, r in hours_df.iterrows()}

    # Holidays ranges (optional)
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

    def parse_hourly(df: pd.DataFrame) -> Dict[Tuple[str,time], int]:
        out={}
        if df is None or df.empty:
            return out
        time_col = pick_col(df, ["Time"], required=False) or df.columns[0]
        ddf=df.copy()
        ddf["Time"]=ddf[time_col].apply(to_time)
        for dn in ["Mon","Tue","Wed","Thu","Fri"]:
            if dn not in ddf.columns:
                continue
            for _, r in ddf.iterrows():
                hh=r.get("Time")
                if not hh:
                    continue
                val=r.get(dn)
                if pd.isna(val) or val=="":
                    continue
                out[(dn, time(hh.hour,0))]=int(float(val))
        return out

    phones_targets = parse_hourly(tph_df)
    if not phones_targets:
        raise ValueError("Phone hourly targets sheet is present but no targets were parsed.")

    bookings_targets = parse_hourly(tbk_df) if tbk_df is not None else {}

    weekly_targets={"Bookings":0.0}
    if tweek_df is not None and not tweek_df.empty:
        task_c = pick_col(tweek_df, ["Task"], required=False) or tweek_df.columns[0]
        val_c  = pick_col(tweek_df, ["WeekHoursTarget","Target","Hours"], required=False) or tweek_df.columns[1]
        for _, r in tweek_df.iterrows():
            tsk=str(r.get(task_c,"")).strip()
            val=r.get(val_c)
            if pd.isna(val) or val=="":
                continue
            if tsk.lower()=="bookings":
                weekly_targets["Bookings"]=float(val)

    # Swaps (optional). Supports swapping shifts AND swapping days off.
    swaps=[]
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
        phones_targets=phones_targets,
        bookings_targets=bookings_targets,
        weekly_targets=weekly_targets,
        swaps=swaps,
    )

# ---------------- Availability + swaps ----------------
def holiday_kind(name: str, d: date, hols: List[Tuple[str,date,date,str]]) -> Optional[str]:
    for n,s,e,k in hols:
        if n.strip().lower()==name.strip().lower() and s and e and s<=d<=e:
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
    """Swap shifts or swap days off (i.e., swap whole day's start/end even if None)."""
    out = {k: dict(v) for k, v in hours_map.items()}
    in_week=set(week_dates)
    for dd, nm, sw, ns, ne in swaps:
        if dd not in in_week:
            continue
        dn=day_name(dd)
        if nm not in out:
            continue
        if sw and sw in out:
            a1,a2 = out[nm].get(f"{dn}Start"), out[nm].get(f"{dn}End")
            b1,b2 = out[sw].get(f"{dn}Start"), out[sw].get(f"{dn}End")
            out[nm][f"{dn}Start"], out[nm][f"{dn}End"] = b1,b2
            out[sw][f"{dn}Start"], out[sw][f"{dn}End"] = a1,a2
        elif (ns is not None) and (ne is not None):
            out[nm][f"{dn}Start"], out[nm][f"{dn}End"] = ns, ne
    return out

# ---------------- Phones requirement ----------------
def phones_required(tpl: TemplateData, d: date, t: time) -> int:
    dn = day_name(d)
    hour_key = time(t.hour, 0)
    return int(tpl.phones_targets.get((dn, hour_key), 0) or 0)

# ---------------- Break placement ----------------
def pick_breaks_site_balanced(
    staff_list: List[Staff],
    hours_map: Dict[str, Dict[str, Optional[time]]],
    hols: List[Tuple[str,date,date,str]],
    week_dates: List[date],
    fixed_assignments: Set[Tuple[date,time,str]],
) -> Dict[Tuple[date,time], Set[str]]:
    """
    Breaks only for staff with break_required AND shift > 6h.
    Spread within site by balancing counts per time slot.
    Avoid placing break where staff is locked to fixed assignment (FD/Triage) in that slot.
    """
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
            dur = (dt_of(d, end) - dt_of(d, stt)).total_seconds()/3600.0
            if dur <= BREAK_THRESHOLD_HOURS:
                continue

            midpoint = dt_of(d, stt) + (dt_of(d, end)-dt_of(d, stt))/2

            best=None
            for bt in BREAK_CANDIDATES:
                if bt < stt or add_minutes(bt, 30) > end:
                    continue
                if not (BREAK_WINDOW[0] <= bt < BREAK_WINDOW[1]):
                    continue
                if (d, bt, st.name) in fixed_assignments:
                    continue

                # Avoid fragmenting into <1h before/after
                before = (dt_of(d, bt) - dt_of(d, stt)).total_seconds()/3600.0
                after  = (dt_of(d, end) - dt_of(d, add_minutes(bt,30))).total_seconds()/3600.0
                frag_penalty = 0
                if before < 1.0:
                    frag_penalty += 10_000
                if after < 1.0:
                    frag_penalty += 10_000

                load = break_load.get((d, st.home, bt), 0)
                dist = abs((dt_of(d, bt) - midpoint).total_seconds())
                score = frag_penalty + (load*3600) + dist
                if best is None or score < best[0]:
                    best=(score, bt)

            if best:
                bt=best[1]
                breaks.setdefault((d, bt), set()).add(st.name)
                break_load[(d, st.home, bt)] = break_load.get((d, st.home, bt), 0) + 1

    return breaks

# ---------------- Scheduling ----------------
def task_weight(st: Staff, key: str) -> int:
    return int(st.weights.get(key, 3) if st.weights else 3)

def block_limits(task: str) -> Tuple[int,int]:
    if task == "Phones":
        return MIN_PHONES, MAX_PHONES
    if task in ("EMIS","Docman"):
        return MIN_ADMIN, MAX_ADMIN
    return MIN_DEFAULT, MAX_DEFAULT

def schedule_week(tpl: TemplateData, week_start: date):
    slots = timeslots()
    dates = [week_start + timedelta(days=i) for i in range(5)]
    hours_map = apply_swaps(tpl.hours_map, tpl.swaps, dates)

    staff_by_name = {s.name: s for s in tpl.staff}
    staff_names = [s.name for s in tpl.staff]

    # Assignments
    a: Dict[Tuple[date,time,str], str] = {}
    gaps: List[Tuple[date,time,str,str]] = []
    fixed_slots: Set[Tuple[date,time,str]] = set()

    mins_task: Dict[Tuple[str,str], int] = {}  # (name, taskKey) -> minutes

    def add_mins(nm: str, key: str, mins: int):
        mins_task[(nm,key)] = mins_task.get((nm,key),0) + mins

    def used(nm: str, key: str) -> int:
        return mins_task.get((nm,key),0)

    def is_free(nm: str, d: date, t: time) -> bool:
        return (d,t,nm) not in a

    # ----- Phase 1: lock fixed bands (FD then Triage) -----
    def band_variants(bs: time, be: time) -> List[Tuple[time,time,int]]:
        """Return candidate (start,end,shift_abs_minutes) allowing ±30 min shift for whole band."""
        variants=[]
        for shift in (-30,0,30):
            s2 = add_minutes(bs, shift)
            e2 = add_minutes(be, shift)
            # keep within day bounds
            if s2 < DAY_START or e2 > DAY_END:
                continue
            variants.append((s2,e2,abs(shift)))
        # prefer zero shift
        variants.sort(key=lambda x: x[2])
        return variants

    def can_cover_band(nm: str, d: date, bs: time, be: time) -> bool:
        stt, end = shift_window(hours_map, d, nm)
        if not stt or not end:
            return False
        if not (stt <= bs and end >= be):
            return False
        # no overlap with existing fixed
        for tt in slots:
            if tt < bs or tt >= be:
                continue
            if (d, tt, nm) in fixed_slots:
                return False
        return True

    def lock_band(d: date, bs: time, be: time, role: str, candidates: List[str], weight_key: str):
        best=None
        best_window=None
        for s2,e2,sh in band_variants(bs,be):
            ok=[nm for nm in candidates if (not holiday_kind(nm,d,tpl.hols)) and can_cover_band(nm,d,s2,e2)]
            if not ok:
                continue
            # weight-first then least-used
            ok.sort(key=lambda nm: (-task_weight(staff_by_name[nm], weight_key),
                                    - (1 if (weight_key=="FrontDesk" and staff_by_name[nm].frontdesk_only) else 0),
                                    used(nm, weight_key),
                                    nm.lower()))
            pick=ok[0]
            score=(sh, -task_weight(staff_by_name[pick], weight_key), used(pick, weight_key), pick.lower())
            if best is None or score < best[0]:
                best=(score,pick)
                best_window=(s2,e2)
        if best is None or best_window is None:
            gaps.append((d, bs, role, "No suitable staff for band (even with ±30m shift)"))
            return
        pick=best[1]
        s2,e2=best_window
        for tt in slots:
            if tt < s2 or tt >= e2:
                continue
            a[(d,tt,pick)] = role
            fixed_slots.add((d,tt,pick))
            add_mins(pick, weight_key, SLOT_MIN)

    for d in dates:
        for site in SITES:
            role=f"FrontDesk_{site}"
            cands=[s.name for s in tpl.staff if s.home==site and s.can_frontdesk]
            for bs,be in FD_BANDS:
                lock_band(d, bs, be, role, cands, "FrontDesk")

    for d in dates:
        for site in ("SLGP","JEN"):
            role=f"Triage_Admin_{site}"
            cands=[s.name for s in tpl.staff if s.home==site and s.can_triage]
            for bs,be in TRIAGE_BANDS:
                lock_band(d, bs, be, role, cands, "Triage")

    # ----- Phase 2: breaks (after fixed placement) -----
    breaks = pick_breaks_site_balanced(tpl.staff, hours_map, tpl.hols, dates, fixed_slots)
    def on_break(nm: str, d: date, t: time) -> bool:
        return nm in breaks.get((d,t), set())

    # ----- Phase 3: block-based fill by strict priority -----
    active: Dict[Tuple[date,str], Tuple[str,int]] = {}  # (d,name)->(task,end_idx_excl)

    def task_key_for(task: str) -> str:
        if task.startswith("FrontDesk_"): return "FrontDesk"
        if task.startswith("Triage_Admin_"): return "Triage"
        if task == "Email_Box": return "Email"
        if task == "Awaiting_PSA_Admin": return "Awaiting"
        if task == "Phones": return "Phones"
        if task == "Bookings": return "Bookings"
        if task == "EMIS": return "EMIS"
        if task == "Docman": return "Docman"
        return "Misc"

    def eligible(nm: str, task: str, d: date, t: time, *, allow_cross_site: bool=False) -> bool:
        st=staff_by_name[nm]
        if holiday_kind(nm,d,tpl.hols): return False
        if not is_working(hours_map, d, t, nm): return False
        if on_break(nm,d,t): return False
        # fixed tasks already allocated
        if task.startswith("FrontDesk_") or task.startswith("Triage_Admin_"):
            return False

        if task == "Phones":
            return st.can_phones

        if task == "Email_Box":
            if not st.can_email: return False
            return True if allow_cross_site else (st.home == email_site_for_day(d))

        if task == "Awaiting_PSA_Admin":
            if not st.can_docman: return False
            return True if allow_cross_site else (st.home == awaiting_site_for_day(d))

        if task == "Bookings":
            if not st.can_bookings: return False
            if allow_cross_site: return True
            return st.home == "SLGP"

        if task == "EMIS":
            return st.can_emis

        if task == "Docman":
            return st.can_docman

        if task == "Misc_Tasks":
            return True
        return True

    def pick_candidates(task: str, d: date, t: time, *, allow_cross_site=False, prefer_sites: Optional[List[str]]=None) -> List[str]:
        c=[]
        for nm in staff_names:
            if not is_free(nm,d,t): 
                continue
            if not eligible(nm, task, d, t, allow_cross_site=allow_cross_site):
                continue
            c.append(nm)

        if prefer_sites:
            c.sort(key=lambda nm: (0 if staff_by_name[nm].home in prefer_sites else 1, nm.lower()))

        key=task_key_for(task)
        c.sort(key=lambda nm: (-task_weight(staff_by_name[nm], key), used(nm, key), nm.lower()))
        return c

    def stop_block(nm: str, d: date):
        active.pop((d,nm), None)

    def apply_active(nm: str, d: date, idx: int) -> bool:
        b=active.get((d,nm))
        if not b:
            return False
        task, end_idx=b
        if idx >= end_idx:
            stop_block(nm,d)
            return False
        t=slots[idx]
        if not is_free(nm,d,t):
            stop_block(nm,d)
            return False
        a[(d,t,nm)] = task
        add_mins(nm, task_key_for(task), SLOT_MIN)
        return True

    def start_block(nm: str, task: str, d: date, start_idx: int) -> bool:
        mn,mx=block_limits(task)
        stt,end=shift_window(hours_map,d,nm)
        if not stt or not end:
            return False
        end_idx=start_idx
        while end_idx < len(slots) and slots[end_idx] < end:
            tt=slots[end_idx]
            if (d,tt,nm) in fixed_slots:
                break
            if nm in breaks.get((d,tt), set()):
                break
            end_idx += 1
        remaining=end_idx-start_idx
        if remaining <= 0:
            return False
        # no floaters: if we can't meet min and it's not genuine end remainder, don't start
        if remaining < mn and remaining > 1:
            return False
        L = min(mx, remaining)
        if remaining >= mn:
            L = max(mn, L)
        active[(d,nm)] = (task, start_idx+L)
        return True

    def assign_block(nm: str, task: str, d: date, idx: int):
        b=active.get((d,nm))
        if b and b[0]==task:
            apply_active(nm,d,idx)
            return
        if b and b[0]!=task:
            stop_block(nm,d)
        if not start_block(nm, task, d, idx):
            # last-resort: extend previous task if possible else misc
            a[(d,slots[idx],nm)]="Misc_Tasks"
            add_mins(nm,"Misc",SLOT_MIN)
            return
        apply_active(nm,d,idx)

    # helper counts per slot
    def on_task(task: str, d: date, t: time) -> List[str]:
        return [nm for nm in staff_names if a.get((d,t,nm))==task]

    # Bookings weekly pressure (optional)
    target_book = int(round((tpl.weekly_targets.get("Bookings",0.0) or 0.0)*60))
    def bookings_needed(d: date, idx: int) -> int:
        if target_book <= 0:
            return 0
        t=slots[idx]
        if t < time(10,30):
            return 0
        done = sum(v for (nm,k), v in mins_task.items() if k=="Bookings")
        remaining=max(0, target_book-done)
        if remaining<=0:
            return 0
        rem_slots=0
        for dd in dates:
            for tt in slots:
                if dd < d: 
                    continue
                if dd==d and tt < t:
                    continue
                if tt >= time(10,30):
                    rem_slots += 1
        if rem_slots<=0:
            return 0
        ppl = math.ceil(remaining/(rem_slots*SLOT_MIN))
        return max(0,ppl)

    # EMIS/Docman hard targets (minutes)
    target_emis = int(HARD_EMIS_HOURS*60)
    target_doc  = int(HARD_DOC_HOURS*60)

    def total_mins(key: str) -> int:
        return sum(v for (nm,k),v in mins_task.items() if k==key)

    # Main slot loop
    for d in dates:
        for idx,t in enumerate(slots):
            # extend active blocks first
            for nm in staff_names:
                if (d,t,nm) in a:
                    continue
                if on_break(nm,d,t):
                    continue
                apply_active(nm,d,idx)

            # --- 3) Phones (hard) ---
            req_p = phones_required(tpl, d, t)
            if req_p > 0:
                # ensure we never exceed available staff; fill highest weight first
                while len(on_task("Phones", d, t)) < req_p:
                    cands = pick_candidates("Phones", d, t, allow_cross_site=True)
                    if not cands:
                        gaps.append((d,t,"Phones", f"Short by {req_p-len(on_task('Phones',d,t))}"))
                        break
                    assign_block(cands[0], "Phones", d, idx)

            # --- 4) Email / Awaiting (site-of-day preferred) ---
            if t >= time(10,30) and t < time(18,30):
                # Email becomes optional after 16 but still attempt if feasible
                need_email = 1 if t < time(16,0) else 1
                if len(on_task("Email_Box", d, t)) < need_email:
                    cands = pick_candidates("Email_Box", d, t, allow_cross_site=False)
                    if not cands:
                        cands = pick_candidates("Email_Box", d, t, allow_cross_site=True)
                    if cands:
                        assign_block(cands[0], "Email_Box", d, idx)
                    else:
                        if t < time(16,0):
                            gaps.append((d,t,"Email_Box","No eligible staff"))

            if t >= time(10,0) and t < time(16,0):
                if len(on_task("Awaiting_PSA_Admin", d, t)) < 1:
                    cands = pick_candidates("Awaiting_PSA_Admin", d, t, allow_cross_site=False)
                    if not cands:
                        cands = pick_candidates("Awaiting_PSA_Admin", d, t, allow_cross_site=True)
                    if cands:
                        assign_block(cands[0], "Awaiting_PSA_Admin", d, idx)
                    else:
                        gaps.append((d,t,"Awaiting_PSA_Admin","No eligible staff"))

            # --- 5) EMIS / Docman (hard weekly targets, max 1 each at a time) ---
            doc_done = total_mins("Docman")
            emis_done = total_mins("EMIS")

            # Only assign if we're behind target
            want_doc = (doc_done < target_doc) and (len(on_task("Docman", d, t)) < 1)
            want_emis = (emis_done < target_emis) and (len(on_task("EMIS", d, t)) < 1)

            # prefer Docman first when both behind (you can tweak)
            if want_doc:
                cands = pick_candidates("Docman", d, t, allow_cross_site=True, prefer_sites=["JEN","BGS"])
                if cands:
                    assign_block(cands[0], "Docman", d, idx)
            if want_emis:
                cands = pick_candidates("EMIS", d, t, allow_cross_site=True, prefer_sites=["JEN","BGS"])
                if cands:
                    assign_block(cands[0], "EMIS", d, idx)

            # --- 6) Bookings (soft weekly pressure) ---
            need_b = bookings_needed(d, idx)
            if need_b > 0:
                while len(on_task("Bookings", d, t)) < need_b:
                    cands = pick_candidates("Bookings", d, t, allow_cross_site=False)
                    if not cands:
                        cands = pick_candidates("Bookings", d, t, allow_cross_site=True)
                    if not cands:
                        break
                    assign_block(cands[0], "Bookings", d, idx)

            # --- 7) Misc only when everything else satisfied ---
            for nm in staff_names:
                if not is_working(hours_map, d, t, nm):
                    continue
                if holiday_kind(nm,d,tpl.hols):
                    continue
                if on_break(nm,d,t):
                    continue
                if not is_free(nm,d,t):
                    continue
                assign_block(nm, "Misc_Tasks", d, idx)

    # Post-pass: smooth isolated single-slot fragments (non-fixed, non-break)
    def smooth_single_slot():
        fixed_prefix=("FrontDesk_","Triage_Admin_")
        special={"Break","Holiday","Bank Holiday","Sick"}
        for d in dates:
            for nm in staff_names:
                seq=[a.get((d,tt,nm)) for tt in slots]
                i=0
                while i < len(slots):
                    task=seq[i]
                    if not task:
                        i+=1; continue
                    j=i+1
                    while j < len(slots) and seq[j]==task:
                        j+=1
                    if (j-i)==1:
                        if str(task).startswith(fixed_prefix) or str(task) in special:
                            i=j; continue
                        prev=seq[i-1] if i-1>=0 else None
                        nxt=seq[j] if j < len(seq) else None
                        chosen=None
                        if prev and (not str(prev).startswith(fixed_prefix)) and str(prev) not in special:
                            chosen=prev
                        elif nxt and (not str(nxt).startswith(fixed_prefix)) and str(nxt) not in special:
                            chosen=nxt
                        if chosen:
                            a[(d,slots[i],nm)] = chosen
                            seq[i]=chosen
                    i=j
    smooth_single_slot()

    return a, breaks, gaps, dates, slots, hours_map

# ---------------- Excel builder ----------------
def _add_conditional_colours(ws, data_range: str, tl_cell: str):
    """Conditional formatting so colours update when users edit timeline cells."""
    for task, color in ROLE_COLORS.items():
        if task == "":
            continue
        fill = PatternFill("solid", fgColor=color)
        ws.conditional_formatting.add(
            data_range,
            FormulaRule(formula=[f'{tl_cell}="{task}"'], fill=fill, stopIfTrue=False)
        )

def _apply_day_borders(ws, start_row: int, rows_per_day: int):
    # Thick line above each day's header row
    for r in range(start_row, ws.max_row+1):
        if (r - start_row) % rows_per_day == 0:
            for c in range(1, ws.max_column+1):
                cell=ws.cell(r,c)
                cell.border = Border(left=cell.border.left, right=cell.border.right, top=THICK, bottom=cell.border.bottom)

def display_site_for_assignment(role: str, d: date, staff_home: str) -> str:
    if role.endswith("_SLGP"): return "SLGP"
    if role.endswith("_JEN"):  return "JEN"
    if role.endswith("_BGS"):  return "BGS"
    if role == "Bookings": return "SLGP"
    if role == "Awaiting_PSA_Admin": return awaiting_site_for_day(d)
    if role == "Email_Box": return email_site_for_day(d)
    # Phones/EMIS/Docman/Misc -> home site for display
    return staff_home

def build_workbook(tpl: TemplateData, start_monday: date, weeks: int) -> Workbook:
    wb=Workbook()
    wb.remove(wb.active)

    slots=timeslots()
    rows_per_day = len(slots) + 1  # +1 repeated header per day

    # For each week create:
    # - WeekX_SLGP_Timeline / JEN / BGS (editable values + conditional colours)
    # - WeekX_Totals (formulas)
    # - WeekX_Coverage (formulas; names per task)
    # - WeekX_NotesAndGaps
    for w in range(weeks):
        wk_start = start_monday + timedelta(days=7*w)
        a, breaks, gaps, dates, slots, hours_map = schedule_week(tpl, wk_start)

        staff_by_name = {s.name: s for s in tpl.staff}

        # ---- Site timelines ----
        site_staff = {
            site: [s.name for s in tpl.staff if s.home == site]
            for site in SITES
        }

        site_sheets={}
        for site in SITES:
            ws = wb.create_sheet(f"Week{w+1}_{site}_Timeline")
            site_sheets[site]=ws
            names=site_staff[site]
            ws.append(["Date","Time"] + names)
            for c in ws[1]:
                c.font=Font(bold=True)
                c.alignment=Alignment(horizontal="center", vertical="center")
                c.border=CELL_BORDER
            ws.freeze_panes="C2"
            ws.column_dimensions["A"].width=14
            ws.column_dimensions["B"].width=8
            for i,_ in enumerate(names):
                ws.column_dimensions[get_column_letter(3+i)].width=18

            # body with repeated headers per day
            for d in dates:
                # repeated header row
                ws.append([d.strftime("%a %d-%b"), ""] + names)
                rr=ws.max_row
                for cc in range(1, ws.max_column+1):
                    cell=ws.cell(rr,cc)
                    cell.font=Font(bold=True)
                    cell.alignment=Alignment(horizontal="center")
                    cell.fill=PatternFill("solid", fgColor="F2F2F2")
                    cell.border=CELL_BORDER

                for t in slots:
                    row=[d.strftime("%a %d-%b"), t.strftime("%H:%M")]
                    for nm in names:
                        hk = holiday_kind(nm, d, tpl.hols)
                        if hk:
                            val=hk
                        elif not is_working(hours_map, d, t, nm):
                            val=""
                        elif nm in breaks.get((d,t), set()):
                            val="Break"
                        else:
                            role=a.get((d,t,nm), "Misc_Tasks")
                            ds = display_site_for_assignment(role, d, staff_by_name[nm].home)
                            val = role if ds==site else ""
                        row.append(val)
                    ws.append(row)

            # Conditional colours for editable cells (exclude headers + repeated headers)
            if names:
                first_data_row = 2
                data_range = f"{get_column_letter(3)}{first_data_row}:{get_column_letter(2+len(names))}{ws.max_row}"
                tl_cell = f"{get_column_letter(3)}{first_data_row}"
                _add_conditional_colours(ws, data_range, tl_cell)

            # Borders
            for rr in range(2, ws.max_row+1):
                for cc in range(1, ws.max_column+1):
                    ws.cell(rr,cc).border=CELL_BORDER
                    ws.cell(rr,cc).alignment=Alignment(vertical="top", wrap_text=True)

            _apply_day_borders(ws, start_row=2, rows_per_day=rows_per_day)

        # ---- Totals (dynamic formulas from site timelines) ----
        ws_tot = wb.create_sheet(f"Week{w+1}_Totals")
        ws_tot.append(["Name","FrontDesk","Triage","Phones","Email","Awaiting","Bookings","EMIS","Docman","Misc","Break","WeeklyTotal"])
        for c in ws_tot[1]:
            c.font=Font(bold=True)
            c.alignment=Alignment(horizontal="center", vertical="center")
            c.border=CELL_BORDER
        ws_tot.freeze_panes="A2"
        ws_tot.column_dimensions["A"].width=22
        for col in range(2, 12+1):
            ws_tot.column_dimensions[get_column_letter(col)].width=12

        def countif_expr(sheet, col_letter, crit):
            # covers up to row 1000 comfortably
            return f'COUNTIF({sheet}!{col_letter}$2:{col_letter}$1000,"{crit}")'

        # For each site, staff columns differ; totals should use the site's timeline for that staff (only if present).
        for nm in [s.name for s in tpl.staff]:
            row=[nm]
            # Identify which site sheet contains the staff
            home = staff_by_name[nm].home
            sh = f"Week{w+1}_{home}_Timeline"
            # Find staff column index in that sheet
            names = site_staff.get(home, [])
            if nm not in names:
                # not in site staff list for some reason
                col_letter = None
            else:
                col_letter = get_column_letter(3 + names.index(nm))

            def halfhours(crit):
                if not col_letter:
                    return "0"
                return f"=0.5*{countif_expr(sh, col_letter, crit)}"

            row += [
                halfhours("FrontDesk*"),
                halfhours("Triage*"),
                halfhours("Phones"),
                halfhours("Email_Box"),
                halfhours("Awaiting_PSA_Admin"),
                halfhours("Bookings"),
                halfhours("EMIS"),
                halfhours("Docman"),
                halfhours("Misc_Tasks"),
                halfhours("Break"),
            ]
            # Weekly total = sum columns B-K
            r_idx = ws_tot.max_row + 1
            row.append(f"=SUM(B{r_idx}:K{r_idx})")
            ws_tot.append(row)

        for rr in range(2, ws_tot.max_row+1):
            for cc in range(1, ws_tot.max_column+1):
                ws_tot.cell(rr,cc).border=CELL_BORDER
                ws_tot.cell(rr,cc).alignment=Alignment(vertical="center")

        # ---- Coverage (dynamic formulas; names per task per slot) ----
        ws_cov = wb.create_sheet(f"Week{w+1}_Coverage")
        tasks = ["FrontDesk_SLGP","FrontDesk_JEN","FrontDesk_BGS","Triage_Admin_SLGP","Triage_Admin_JEN","Phones","Email_Box","Awaiting_PSA_Admin","Bookings","EMIS","Docman","Misc_Tasks"]
        ws_cov.append(["Date","Time"] + tasks)
        for c in ws_cov[1]:
            c.font=Font(bold=True); c.alignment=Alignment(horizontal="center"); c.border=CELL_BORDER
        ws_cov.freeze_panes="C2"
        ws_cov.column_dimensions["A"].width=14
        ws_cov.column_dimensions["B"].width=8
        for i,_ in enumerate(tasks):
            ws_cov.column_dimensions[get_column_letter(3+i)].width=26

        # helper formulas for each task from each site timeline at a specific row
        def names_formula(site, row_idx_in_site, task_value):
            ws_name = f"Week{w+1}_{site}_Timeline"
            names = site_staff[site]
            if not names:
                return '""'
            start_col = get_column_letter(3)
            end_col = get_column_letter(2+len(names))
            header_rng = f"{ws_name}!${start_col}$1:${end_col}$1"
            row_rng = f"{ws_name}!${start_col}${row_idx_in_site}:${end_col}${row_idx_in_site}"
            # Use FILTER(headers, row=task)
            return f'IFERROR(TEXTJOIN(", ",TRUE,FILTER({header_rng},{row_rng}="{task_value}")),"")'

        # Build coverage rows aligned with site sheets row indices.
        # Site sheets include header row 1 then alternating (day header row + slots rows). We'll mirror same ordering.
        # We'll generate per day: one spacer/dayheader row in coverage too for print, then slot rows.
        for d in dates:
            ws_cov.append([d.strftime("%a %d-%b"), ""] + [""]*len(tasks))
            hdr_row = ws_cov.max_row
            for cc in range(1, ws_cov.max_column+1):
                cell=ws_cov.cell(hdr_row,cc)
                cell.font=Font(bold=True)
                cell.fill=PatternFill("solid", fgColor="F2F2F2")
                cell.border=CELL_BORDER
            for t in slots:
                ws_cov.append([d.strftime("%a %d-%b"), t.strftime("%H:%M")] + [""]*len(tasks))
                cov_rr = ws_cov.max_row
                # Determine corresponding row index in each site sheet:
                # Each site sheet layout per day: 1 repeated header + len(slots) rows.
                # Row 2 starts first day's repeated header, then slot rows.
                day_index = (d - dates[0]).days
                # for each day: offset = 2 + day_index*(rows_per_day) ; slot row within day = 1 + slot_idx (since row0=header)
                base = 2 + day_index*rows_per_day
                slot_idx = slots.index(t)
                site_row = base + 1 + slot_idx  # +1 for day header row

                # Fill each coverage cell with combined names across sites
                for i, task in enumerate(tasks, start=3):
                    parts = []
                    for site in SITES:
                        parts.append(names_formula(site, site_row, task))
                    ws_cov.cell(cov_rr, i).value = f"=TEXTJOIN(\", \",TRUE,{','.join(parts)})"
                    ws_cov.cell(cov_rr, i).alignment=Alignment(wrap_text=True, vertical="top")
                    ws_cov.cell(cov_rr, i).border=CELL_BORDER

        # Colour entire coverage columns lightly by task (simple fill on header + col)
        for col_i, task in enumerate(tasks, start=3):
            fill = PatternFill("solid", fgColor=ROLE_COLORS.get(task, "FFFFFF"))
            # only shade header row 1
            ws_cov.cell(1,col_i).fill = fill

        _apply_day_borders(ws_cov, start_row=2, rows_per_day=rows_per_day)

        # ---- Notes/Gaps ----
        ws_g = wb.create_sheet(f"Week{w+1}_NotesAndGaps")
        ws_g.append(["Date","Time","Task","Note"])
        for c in ws_g[1]:
            c.font=Font(bold=True); c.border=CELL_BORDER
        for d,t,task,note in gaps:
            ws_g.append([d.isoformat(), t.strftime("%H:%M") if t else "", task, note])

    return wb
