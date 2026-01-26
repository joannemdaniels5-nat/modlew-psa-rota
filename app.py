import io
import re
from datetime import datetime, date, time, timedelta

import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows

# =========================================================
# Rota Generator — v11 (Lewisham PSA rules)
# =========================================================
# Key hard rules implemented here:
# - Open Mon–Fri 08:00–18:30 in 30-min slots
# - Front Desk (SLGP/JEN/BGS):
#     * exactly 1 per site per slot
#     * home-site only
#     * each block 1.5–3h (3–6 slots)
#     * blocks end at break (no 30- or 60-min pre-break stints)
# - Carol Church:
#     * SLGP Front Desk only
#     * 08:00–11:00 Mon–Fri
# - Triage Admin (SLGP & JEN):
#     * required 08:00–16:00 Mon–Fri
#     * home-site only
#     * blocks 2–4h (4–8 slots)
# - Breaks:
#     * if shift > 6h, one 30-min break
#     * break chosen between 11:30–14:00
#     * spread across staff to avoid everyone off at the same time
# - All other tasks:
#     * minimum 2h blocks (4+ slots), shorter allowed only at END of shift
# - Awaiting Response / PSA Admin:
#     * 10:00–16:00 mandatory, 16:00–18:30 optional
#     * site-of-day pattern: Mon/Fri SLGP, Tue/Thu JEN, Wed BGS
#     * home-site only
# - Phones / Bookings / Email / EMIS / Docman:
#     * Phones & Bookings have fixed minimum coverage
#     * Preferences (soft, not hard):
#         - Bookings: prefer SLGP but allow others
#         - Email: prefer site-of-day (Mon BGS, Tue–Thu JEN, Fri SLGP) but allow others
#         - Phones: prefer JEN/BGS but allow SLGP
#         - EMIS/Docman: prefer JEN/BGS but allow SLGP
#     * Per-staff weights (0–5) are tie-breakers only, never hard blocks

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

# ---------------- Parsing helpers ----------------
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

# ---------------- Read template (NO Parameters required) ----------------
def read_template(uploaded_bytes: bytes):
    xls = pd.ExcelFile(io.BytesIO(uploaded_bytes))
    staff_sheet = find_sheet(xls, ["Staff", "Skills", "People"])
    hours_sheet = find_sheet(xls, ["WorkingHours", "Hours", "Availability"])
    hols_sheet = find_sheet(xls, ["Holidays", "Leave", "Absence"])

    if not staff_sheet:
        raise ValueError(f"Could not find Staff sheet. Found: {xls.sheet_names}")
    if not hours_sheet:
        raise ValueError(f"Could not find WorkingHours sheet. Found: {xls.sheet_names}")

    staff_df = pd.read_excel(xls, sheet_name=staff_sheet)
    hours_df = pd.read_excel(xls, sheet_name=hours_sheet)
    hols_df = pd.read_excel(xls, sheet_name=hols_sheet) if hols_sheet else pd.DataFrame()

    name_c = pick_col(staff_df, ["Name", "StaffName"])
    home_c = pick_col(staff_df, ["HomeSite", "Site", "BaseSite"], required=False)

    staff_df = staff_df.copy()
    staff_df["Name"] = staff_df[name_c].astype(str).str.strip()
    staff_df["HomeSite"] = staff_df[home_c].astype(str).str.strip().str.upper() if home_c else ""

    def yn(v):
        if pd.isna(v):
            return False
        s = str(v).strip().lower()
        return s in ["y", "yes", "true", "1"]

    for c in staff_df.columns:
        if normalize(c).startswith("can") or normalize(c) in ["iscarolchurchyn", "iscarolchurch"]:
            staff_df[c] = staff_df[c].apply(yn)

    carol_c = pick_col(staff_df, ["IsCarolChurch(Y/N)", "IsCarolChurch", "Carol"], required=False)
    if carol_c:
        staff_df["IsCarolChurch"] = staff_df[carol_c].apply(bool)
    else:
        staff_df["IsCarolChurch"] = staff_df["Name"].str.lower().eq("carol church")

    hours_df = hours_df.copy()
    hours_name_c = pick_col(hours_df, ["Name", "StaffName"])
    hours_df["Name"] = hours_df[hours_name_c].astype(str).str.strip()

    for d in ["Mon", "Tue", "Wed", "Thu", "Fri"]:
        sc = pick_col(hours_df, [f"{d}Start", f"{d} Start", f"{d}_Start"], required=False)
        ec = pick_col(hours_df, [f"{d}End", f"{d} End", f"{d}_End"], required=False)
        hours_df[f"{d}Start"] = hours_df[sc].apply(to_time) if sc else None
        hours_df[f"{d}End"] = hours_df[ec].apply(to_time) if ec else None

    hols = []
    if not hols_df.empty:
        hn = pick_col(hols_df, ["Name", "StaffName"], required=False) or hols_df.columns[0]
        hs = pick_col(hols_df, ["StartDate", "Start"], required=False) or hols_df.columns[1]
        he = pick_col(hols_df, ["EndDate", "End"], required=False) or hols_df.columns[2]
        notes_c = pick_col(hols_df, ["Notes", "Note", "Reason"], required=False)
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

    return staff_df, hours_df, hols, {"staff": staff_sheet, "hours": hours_sheet, "holidays": hols_sheet or ""}

# ---------------- Business constants ----------------
DAY_START = time(8, 0)
DAY_END = time(18, 30)
SLOT_MIN = 30

# Block lengths
MIN_BLOCK_SLOTS_DEFAULT = 4   # 2h
FD_MIN_BLOCK_SLOTS = 3        # 1.5h for Front Desk
TRIAGE_MIN_BLOCK_SLOTS = 4    # 2h

MAX_BLOCK_SLOTS_DEFAULT = 6   # 3h
PHONES_MAX_BLOCK_SLOTS = 7      # 3.5h max for Phones
FD_MAX_BLOCK_SLOTS = 6        # 3h
TRIAGE_MAX_BLOCK_SLOTS = 100   # effectively no max (up to full day)
EMAIL_MAX_BLOCK_SLOTS = 8     # up to 4h

BREAK_THRESHOLD_HOURS = 6.0

# Phones targets by weekday (0=Mon..4=Fri) and hour, as per workforce model
PHONES_TARGETS = {
    0: {  # Monday
        8: 5, 9: 4, 10: 4, 11: 3, 12: 3, 13: 3, 14: 3, 15: 3, 16: 2, 17: 1, 18: 1,
    },
    1: {  # Tuesday
        8: 4, 9: 3, 10: 3, 11: 3, 12: 3, 13: 2, 14: 2, 15: 1, 16: 1, 17: 0, 18: 0,
    },
    2: {  # Wednesday
        8: 3, 9: 3, 10: 3, 11: 3, 12: 3, 13: 3, 14: 3, 15: 2, 16: 2, 17: 1, 18: 1,
    },
    3: {  # Thursday
        8: 3, 9: 3, 10: 3, 11: 3, 12: 2, 13: 1, 14: 1, 15: 1, 16: 0, 17: 1, 18: 0,
    },
    4: {  # Friday
        8: 4, 9: 4, 10: 4, 11: 4, 12: 3, 13: 2, 14: 2, 15: 2, 16: 2, 17: 1, 18: 0,
    },
}

BREAK_CANDIDATE_TIMES = [
    time(11, 30), time(12, 0), time(12, 30),
    time(13, 0), time(13, 30), time(14, 0),
]

def timeslots():
    cur = datetime(2000, 1, 1, DAY_START.hour, DAY_START.minute)
    end = datetime(2000, 1, 1, DAY_END.hour, DAY_END.minute)
    out = []
    while cur < end:
        out.append(cur.time())
        cur += timedelta(minutes=SLOT_MIN)
    return out

# Mandatory coverage baseline (per slot)
MANDATORY_RULES = [
    ("FrontDesk_SLGP", DAY_START, DAY_END, 1),
    ("FrontDesk_JEN",  DAY_START, DAY_END, 1),
    ("FrontDesk_BGS",  DAY_START, DAY_END, 1),
    ("Triage_Admin_SLGP", DAY_START, time(16, 0), 1),
    ("Triage_Admin_JEN",  DAY_START, time(16, 0), 1),
    ("Email_Box", DAY_START, DAY_END, 1),
    ("Phones", DAY_START, DAY_END, 2),
    ("Bookings", DAY_START, DAY_END, 3),
]

def awaiting_site_for_day(d: date) -> str:
    wd = d.weekday()  # Mon=0
    if wd in (0, 4):  # Mon/Fri
        return "SLGP"
    if wd in (1, 3):  # Tue/Thu
        return "JEN"
    return "BGS"      # Wed

def awaiting_required(d: date, t: time) -> bool:
    return t_in_range(t, time(10, 0), time(16, 0))

def awaiting_optional(d: date, t: time) -> bool:
    return t_in_range(t, time(16, 0), DAY_END)

def filler_order_for_day(d: date):
    # Mon/Fri prioritise phones over EMIS/Docman (busy phones days)
    if d.weekday() in (0, 4):
        return ["Phones", "Bookings", "Awaiting_PSA_Admin", "Emis_Tasks", "Docman_Tasks"]
    return ["Phones", "Bookings", "Emis_Tasks", "Docman_Tasks", "Awaiting_PSA_Admin"]

def email_site_for_day(d: date) -> str:
    # Option C: prefer BGS Mon, JEN Tue–Thu, SLGP Fri
    wd = d.weekday()
    if wd == 0:
        return "BGS"
    if wd in (1, 2, 3):
        return "JEN"
    return "SLGP"

# ---------------- Skills + weights ----------------
def can(sr, col: str) -> bool:
    return bool(sr.get(col, False))

def skill_allowed(sr, role: str) -> bool:
    if role.startswith("FrontDesk"):
        return can(sr, "CanFrontDesk")
    if role.startswith("Triage_Admin"):
        return can(sr, "CanTriage")
    if role == "Email_Box":
        return can(sr, "CanEmail")
    if role == "Phones":
        return can(sr, "CanPhones")
    if role == "Bookings":
        return can(sr, "CanBookings")
    if role == "Emis_Tasks":
        return can(sr, "CanEMIS")
    if role == "Docman_Tasks":
        return can(sr, "CanDocman_PSA") or can(sr, "CanDocman_AWAIT")
    if role == "Awaiting_PSA_Admin":
        return can(sr, "CanDocman_PSA") or can(sr, "CanDocman_AWAIT")
    if role == "Unassigned":
        return True
    return False

def skill_weight(sr, role: str) -> int:
    # Optional per-person tie-break preference (0–5). Missing/blank -> 3.
    mapping = {
        "Phones": "PhonesWeight",
        "Bookings": "BookingsWeight",
        "Emis_Tasks": "EmisWeight",
        "Docman_Tasks": "DocmanWeight",
        "Awaiting_PSA_Admin": "AwaitingWeight",
        "Email_Box": "EmailWeight",
        "Triage_Admin_SLGP": "TriageWeight",
        "Triage_Admin_JEN": "TriageWeight",
        "FrontDesk_SLGP": "FrontDeskWeight",
        "FrontDesk_JEN": "FrontDeskWeight",
        "FrontDesk_BGS": "FrontDeskWeight",
    }
    col = mapping.get(role)
    if not col:
        return 3
    try:
        v = sr.get(col, 3)
        if pd.isna(v):
            return 3
        v = int(float(v))
        return max(0, min(5, v))
    except Exception:
        return 3

def role_site(role: str, d: date) -> str:
    if role.endswith("_SLGP"):
        return "SLGP"
    if role.endswith("_JEN"):
        return "JEN"
    if role.endswith("_BGS"):
        return "BGS"
    if role == "Bookings":
        # bookings can be any site but strongly preferred at SLGP (handled via bias)
        return ""
    if role == "Awaiting_PSA_Admin":
        return awaiting_site_for_day(d)
    return ""

def is_frontdesk(role: str) -> bool:
    return role.startswith("FrontDesk_")

def is_triage(role: str) -> bool:
    return role.startswith("Triage_Admin_")

def holiday_kind(name: str, d: date, hols):
    for n, s, e, k in hols:
        if n.strip().lower() == name.strip().lower() and s and e and s <= d <= e:
            return k
    return None

# =========================================================
# Break placement (closest to midpoint, spread by load)
# =========================================================
def build_maps(staff_df, hours_df):
    staff_by_name = {r["Name"]: r for _, r in staff_df.iterrows()}
    hmap = {r["Name"]: r for _, r in hours_df.iterrows()}
    return staff_by_name, hmap

def shift_window(hmap, week_start, d, name):
    days = ["Mon", "Tue", "Wed", "Thu", "Fri"]
    dname = days[(d - week_start).days]
    hr = hmap.get(name)
    if hr is None:
        return None, None
    return hr.get(f"{dname}Start"), hr.get(f"{dname}End")

def is_working(hmap, week_start, d, t, name):
    stt, end = shift_window(hmap, week_start, d, name)
    return bool(stt and end and (t >= stt) and (t < end))

def pick_break_slot(d: date, stt: time, end: time, load_map):
    if not stt or not end:
        return None
    shift_hours = (dt_of(d, end) - dt_of(d, stt)).total_seconds() / 3600.0
    if shift_hours <= BREAK_THRESHOLD_HOURS:
        return None

    midpoint = dt_of(d, stt) + (dt_of(d, end) - dt_of(d, stt)) / 2
    candidates = []
    for t in BREAK_CANDIDATE_TIMES:
        if t < stt or add_minutes(t, 30) > end:
            continue
        dist = abs((dt_of(d, t) - midpoint).total_seconds())
        load = load_map.get((d, t), 0)
        # sort by load first (spread breaks), then closeness to midpoint
        candidates.append(((load, dist), t))
    if not candidates:
        return None
    candidates.sort(key=lambda x: x[0])
    return candidates[0][1]

# =========================================================
# Block-based scheduler
# =========================================================
def min_slots_for_role(role: str) -> int:
    if is_frontdesk(role):
        return FD_MIN_BLOCK_SLOTS
    if is_triage(role):
        return TRIAGE_MIN_BLOCK_SLOTS
    if role in ("Phones", "Bookings", "Email_Box", "Emis_Tasks", "Docman_Tasks", "Awaiting_PSA_Admin"):
        return MIN_BLOCK_SLOTS_DEFAULT
    return MIN_BLOCK_SLOTS_DEFAULT

def max_slots_for_role(role: str) -> int:
    if is_frontdesk(role):
        return FD_MAX_BLOCK_SLOTS
    if is_triage(role):
        return TRIAGE_MAX_BLOCK_SLOTS
    if role == "Email_Box":
        return EMAIL_MAX_BLOCK_SLOTS
    return MAX_BLOCK_SLOTS_DEFAULT

def site_bias(sr, role: str, d: date) -> float:
    home = str(sr.get("HomeSite", "")).strip().upper()
    if role == "Bookings":
        return 0.0 if home == "SLGP" else 2.0
    if role == "Email_Box":
        pref = email_site_for_day(d)
        if home == pref:
            return 0.0
        if home in ("JEN", "BGS"):
            return 1.0
        return 2.0
    if role in ("Phones", "Emis_Tasks", "Docman_Tasks"):
        # Prefer JEN/BGS for phones + EMIS/Docman
        return 0.0 if home in ("JEN", "BGS") else 2.0
    return 0.0

def rota_generate_week(staff_df, hours_df, hols, week_start: date):
    slots = timeslots()
    dates = [week_start + timedelta(days=i) for i in range(5)]
    staff_by_name, hmap = build_maps(staff_df, hours_df)

    # Breaks with load-based spreading
    breaks = {}
    break_load = {}  # (d,t)->count
    for d in dates:
        for name in staff_by_name.keys():
            if holiday_kind(name, d, hols):
                continue
            stt, end = shift_window(hmap, week_start, d, name)
            b = pick_break_slot(d, stt, end, break_load)
            if b:
                breaks.setdefault((d, b), set()).add(name)
                break_load[(d, b)] = break_load.get((d, b), 0) + 1

    # Assignments: (d,t,name)->role
    a = {}
    # Active blocks: (d,name)->(role, end_idx_excl)
    active = {}
    # Minutes per role: (d,name,role)->minutes
    role_minutes = {}
    gaps = []

    def on_break(name, d, t):
        return name in breaks.get((d, t), set())

    def is_free(name, d, t):
        return (d, t, name) not in a

    def staff_assigned_to(role, d, t):
        return [nm for (dd, tt, nm), rr in a.items() if dd == d and tt == t and rr == role]

    def needed_for_slot(d, t):
        wanted = []
        for role, ta, tb, need in MANDATORY_RULES:
            if t_in_range(t, ta, tb):
                wanted.append((role, need))
        if awaiting_required(d, t) or awaiting_optional(d, t):
            wanted.append(("Awaiting_PSA_Admin", 1))
        return wanted

    def candidate_ok(name, role, d, t):
        sr = staff_by_name[name]

        if holiday_kind(name, d, hols):
            return False
        if not is_working(hmap, week_start, d, t, name):
            return False
        if on_break(name, d, t):
            return False
        if not is_free(name, d, t):
            return False
        if not skill_allowed(sr, role):
            return False

        home = str(sr.get("HomeSite", "")).strip().upper()
        target_site = role_site(role, d)

        # Carol Church special case: only SLGP Front Desk, 08:00–11:00
        if sr.get("IsCarolChurch", False):
            if role != "FrontDesk_SLGP":
                return False
            # 08:00–11:00 window
            if not t_in_range(t, time(8, 0), time(11, 0)):
                return False
            return True

        # Front desk: strict home-site only
        if is_frontdesk(role):
            return home == target_site

        # Triage: strict home-site only
        if is_triage(role):
            return home == target_site

        # Awaiting: strict to site-of-day (no cross-site)
        if role == "Awaiting_PSA_Admin":
            return home == awaiting_site_for_day(d)

        # Bookings: any site with skill (bias handled separately)
        if role == "Bookings":
            return True

        # Email: any staff with CanEmail (bias by site-of-day handled separately)
        if role == "Email_Box":
            return True

        # Phones/Emis/Docman: any staff with skill (bias handled separately)
        if role in ("Phones", "Emis_Tasks", "Docman_Tasks"):
            return True

        return True

    def pick_block_len(name, role, d, start_idx):
        stt, end = shift_window(hmap, week_start, d, name)
        if not stt or not end:
            return 0

        # shift end index
        end_idx = start_idx
        while end_idx < len(slots) and slots[end_idx] < end:
            end_idx += 1

        # break cut
        break_idx = None
        for k in range(start_idx, min(len(slots), end_idx)):
            if name in breaks.get((d, slots[k]), set()):
                break_idx = k
                break

        hard_end = break_idx if break_idx is not None else end_idx
        remaining = hard_end - start_idx
        if remaining <= 0:
            return 0

        min_slots = min_slots_for_role(role)
        max_slots = max_slots_for_role(role)

        # If remaining before a break is less than min, do NOT start a block that would end just before break
        if remaining < min_slots and break_idx is not None:
            return 0

        # If it's the final tail of the shift (no break ahead), allow short remainder
        if remaining < min_slots and break_idx is None:
            return remaining

        return min(max_slots, remaining)

    def start_block(name, role, d, start_idx):
        L = pick_block_len(name, role, d, start_idx)
        if L <= 0:
            return False
        active[(d, name)] = (role, start_idx + L)
        return True

    def apply_block_if_active(name, d, idx):
        b = active.get((d, name))
        if not b:
            return False
        role, end_idx = b
        if idx >= end_idx:
            del active[(d, name)]
            return False
        t = slots[idx]
        a[(d, t, name)] = role
        role_minutes[(d, name, role)] = role_minutes.get((d, name, role), 0) + SLOT_MIN
        return True

    def enforce_frontdesk_exactly_one(d, idx):
        t = slots[idx]
        for site in ("SLGP", "JEN", "BGS"):
            role = f"FrontDesk_{site}"
            current = staff_assigned_to(role, d, t)

            if len(current) > 1:
                gaps.append((d, t, role, f"More than 1 assigned ({', '.join(current)})"))
                for extra in current[1:]:
                    a[(d, t, extra)] = "Unassigned"
                current = current[:1]

            if len(current) == 1:
                continue

            cands = []
            for name, sr in staff_by_name.items():
                if not candidate_ok(name, role, d, t):
                    continue
                used_mins = role_minutes.get((d, name, role), 0)
                primary = used_mins / 60.0

                # continuity / swap friction for this role
                prev_idx = idx - 1
                if prev_idx >= 0:
                    prev_t = slots[prev_idx]
                    prev_role = a.get((d, prev_t, name))
                    if prev_role == role:
                        primary -= 2.0  # strong bonus for staying on same role
                    elif prev_role and prev_role != role:
                        primary += 2.0  # friction for swapping roles unnecessarily

                tie = -(skill_weight(sr, role) - 3)
                cands.append(((primary, tie), name))

            if not cands:
                gaps.append((d, t, role, "No suitable staff available"))
                continue

            cands.sort(key=lambda x: x[0])
            pick = cands[0][1]
            if (d, pick) not in active:
                start_block(pick, role, d, idx)
            apply_block_if_active(pick, d, idx)

    def enforce_role_need(role, need, d, idx):
        t = slots[idx]
        current = staff_assigned_to(role, d, t)
        while len(current) < need:
            cands = []
            for name, sr in staff_by_name.items():
                b = active.get((d, name))
                if b is not None and b[0] != role:
                    continue
                if not candidate_ok(name, role, d, t):
                    continue

                used_mins = role_minutes.get((d, name, role), 0)
                primary = used_mins / 60.0 + site_bias(sr, role, d)

                # continuity / swap friction for this role
                prev_idx = idx - 1
                if prev_idx >= 0:
                    prev_t = slots[prev_idx]
                    prev_role = a.get((d, prev_t, name))
                    if prev_role == role:
                        primary -= 2.0
                    elif prev_role and prev_role != role:
                        primary += 2.0

                tie = -(skill_weight(sr, role) - 3)
                cands.append(((primary, tie), name))

            if not cands:
                # Only log gaps for mandatory window; awaiting optional after 16:00 can be skipped
                if role != "Awaiting_PSA_Admin" or awaiting_required(d, t):
                    gaps.append((d, t, role, "No suitable staff available"))
                break

            cands.sort(key=lambda x: x[0])
            pick = cands[0][1]
            if (d, pick) not in active:
                start_block(pick, role, d, idx)
            apply_block_if_active(pick, d, idx)
            current = staff_assigned_to(role, d, t)

    # Main loop
    for d in dates:
        for idx, t in enumerate(slots):
            # extend existing blocks first
            for name in staff_by_name.keys():
                if (d, t, name) in a:
                    continue
                apply_block_if_active(name, d, idx)

            # enforce front desk per site
            enforce_frontdesk_exactly_one(d, idx)

            # enforce all other mandatory roles
            for role, need in needed_for_slot(d, t):
                if role.startswith("FrontDesk_"):
                    continue
                enforce_role_need(role, need, d, idx)

            # Fill remaining free staff with filler tasks (do not leave working people unassigned)
            filler_roles = filler_order_for_day(d)

            for name, sr in staff_by_name.items():
                if holiday_kind(name, d, hols):
                    continue
                if not is_working(hmap, week_start, d, t, name):
                    continue
                if on_break(name, d, t):
                    continue
                if not is_free(name, d, t):
                    continue

                home = str(sr.get("HomeSite", "")).strip().upper()
                chosen = None

                # Enforce "majority of SLGP on bookings" by strongly biasing SLGP fillers to Bookings
                if home == "SLGP" and skill_allowed(sr, "Bookings"):
                    chosen = "Bookings"

                if not chosen:
                    best = None
                    for role in filler_roles:
                        # Awaiting optional only after 16:00; mandatory window already enforced above
                        if role == "Awaiting_PSA_Admin" and not awaiting_optional(d, t):
                            continue
                        if not skill_allowed(sr, role):
                            continue
                        if role == "Awaiting_PSA_Admin" and home != awaiting_site_for_day(d):
                            continue

                        used = role_minutes.get((d, name, role), 0)
                        # Don't keep someone on one filler all day: cap 6h per filler (except Email which is allowed longer)
                        if role != "Email_Box" and used >= 360:
                            continue
                        primary = used / 60.0 + site_bias(sr, role, d)

                        # continuity / swap friction for fillers
                        prev_idx = idx - 1
                        if prev_idx >= 0:
                            prev_t = slots[prev_idx]
                            prev_role = a.get((d, prev_t, name))
                            if prev_role == role:
                                primary -= 2.0
                            elif prev_role and prev_role != role:
                                primary += 2.0

                        tie = -(skill_weight(sr, role) - 3)
                        score = (primary, tie)
                        if best is None or score < best[0]:
                            best = (score, role)
                    if best:
                        chosen = best[1]

                if not chosen:
                    chosen = "Additional Capacity"

                if (d, name) not in active:
                    start_block(name, chosen, d, idx)
                apply_block_if_active(name, d, idx)

    # Break sanity check
    for d in dates:
        for name in staff_by_name.keys():
            if holiday_kind(name, d, hols):
                continue
            stt, end = shift_window(hmap, week_start, d, name)
            if not stt or not end:
                continue
            shift_hours = (dt_of(d, end) - dt_of(d, stt)).total_seconds() / 3600.0
            if shift_hours <= BREAK_THRESHOLD_HOURS:
                continue
            had_break = False
            for bt in BREAK_CANDIDATE_TIMES:
                if name in breaks.get((d, bt), set()):
                    had_break = True
                    break
            if not had_break:
                gaps.append((d, None, "Break", f"{name}: shift > 6h but no break could be placed in window"))

    return a, breaks, gaps, hmap

# =========================================================
# Excel output (colour-coded)
# =========================================================
ROLE_COLORS = {
    "FrontDesk_SLGP": "FFF2CC",
    "FrontDesk_JEN":  "FFF2CC",
    "FrontDesk_BGS":  "FFF2CC",
    "Triage_Admin_SLGP": "D9EAD3",
    "Triage_Admin_JEN":  "D9EAD3",
    "Email_Box": "CFE2F3",
    "Phones": "C9DAF8",
    "Bookings": "FCE5CD",
    "Emis_Tasks": "EAD1DC",
    "Docman_Tasks": "D0E0E3",
    "Awaiting_PSA_Admin": "D0E0E3",
    "Additional Capacity": "E2EFDA",
    "Unassigned": "FFFFFF",
    "Break": "CFE2F3",
    "Holiday": "FFF2CC",
    "Bank Holiday": "FFE599",
    "Sick": "F4CCCC",
    "": "DDDDDD",  # non-working blank grey
}

def fill_for(value: str):
    return PatternFill("solid", fgColor=ROLE_COLORS.get(value, "FFFFFF"))

def build_workbook(staff_df, hours_df, hols, start_monday: date, weeks: int):
    wb = Workbook()
    wb.remove(wb.active)
    slots = timeslots()

    for w in range(weeks):
        week_start = start_monday + timedelta(days=7 * w)
        dates = [week_start + timedelta(days=i) for i in range(5)]
        staff_names = list(staff_df["Name"].astype(str))

        a, breaks, gaps, hmap = rota_generate_week(staff_df, hours_df, hols, week_start)

        # Grid
        ws_grid = wb.create_sheet(f"Week{w+1}_Grid")
        ws_grid.append(["Time"] + [d.strftime("%a %d-%b") for d in dates])
        for c in ws_grid[1]:
            c.font = Font(bold=True)
            c.alignment = Alignment(horizontal="center", vertical="center")

        display_roles = [
            "FrontDesk_SLGP","FrontDesk_JEN","FrontDesk_BGS",
            "Triage_Admin_SLGP","Triage_Admin_JEN",
            "Email_Box","Phones","Bookings",
            "Awaiting_PSA_Admin","Emis_Tasks","Docman_Tasks",
        ]

        for t in slots:
            row = [t.strftime("%H:%M")]
            for d in dates:
                parts = []
                for role in display_roles:
                    ppl = [nm for nm in staff_names if a.get((d, t, nm)) == role]
                    if ppl:
                        parts.append(f"{role}: " + ", ".join(ppl))
                row.append("\n".join(parts))
            ws_grid.append(row)

        ws_grid.column_dimensions["A"].width = 8
        for col in range(2, 7):
            ws_grid.column_dimensions[chr(64 + col)].width = 42
        for r in range(2, 2 + len(slots)):
            ws_grid.row_dimensions[r].height = 95
            for c in range(2, 7):
                ws_grid.cell(r, c).alignment = Alignment(wrap_text=True, vertical="top")

        # Coverage gaps
        ws_gaps = wb.create_sheet(f"Week{w+1}_CoverageGaps")
        ws_gaps.append(["Date", "Time", "Role", "Issue"])
        for c in ws_gaps[1]:
            c.font = Font(bold=True)
        for d, t, role, issue in gaps:
            ws_gaps.append([d.isoformat(), "" if t is None else t.strftime("%H:%M"), role, issue])

        # Totals
        ws_tot = wb.create_sheet(f"Week{w+1}_Totals")
        rows = []
        for d in dates:
            for t in slots:
                for nm in staff_names:
                    role = a.get((d, t, nm))
                    if role:
                        rows.append([d, nm, role, 0.5])
                for nm in breaks.get((d, t), set()):
                    rows.append([d, nm, "Break", 0.5])

        df = pd.DataFrame(rows, columns=["Date", "Name", "Task", "Hours"])
        if df.empty:
            df = pd.DataFrame(columns=["Date", "Name", "Task", "Hours"])

        # Merge Emis_Tasks + Docman_Tasks into a single public label "EMIS/Docman Tasks"
        df["Task"] = df["Task"].replace({
            "Emis_Tasks": "EMIS/Docman Tasks",
            "Docman_Tasks": "EMIS/Docman Tasks",
        })

        pivot_w = (
            df.groupby(["Name", "Task"])["Hours"].sum()
            .reset_index()
            .pivot(index="Name", columns="Task", values="Hours")
            .fillna(0.0)
        )
        pivot_w["WeeklyTotal"] = pivot_w.sum(axis=1)
        pivot_w = pivot_w.reset_index()

        pivot_d = df.groupby(["Date", "Name", "Task"])["Hours"].sum().reset_index()

        ws_tot.append(["Weekly totals per staff by task (hours)"])
        ws_tot["A1"].font = Font(bold=True, size=12)
        ws_tot.append([])
        for r in dataframe_to_rows(pivot_w, index=False, header=True):
            ws_tot.append(r)

        header_row = 3
        for cell in ws_tot[header_row]:
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center")

        ws_tot.append([])
        ws_tot.append(["Daily totals per staff by task (hours)"])
        ws_tot[f"A{ws_tot.max_row}"].font = Font(bold=True, size=12)
        ws_tot.append([])
        for r in dataframe_to_rows(pivot_d, index=False, header=True):
            ws_tot.append(r)

        # Timelines (All staff)
        ws_tl = wb.create_sheet(f"Week{w+1}_Timelines_All")
        ws_tl.append(["Date", "Time"] + staff_names)
        for c in ws_tl[1]:
            c.font = Font(bold=True)
            c.alignment = Alignment(horizontal="center", vertical="center")
        ws_tl.freeze_panes = "C2"

        for d in dates:
            for t in slots:
                row = [d.strftime("%a %d-%b"), t.strftime("%H:%M")]
                for nm in staff_names:
                    hk = holiday_kind(nm, d, hols)
                    if hk:
                        row.append(hk)
                    elif not is_working(hmap, week_start, d, t, nm):
                        row.append("")  # blank grey
                    elif nm in breaks.get((d, t), set()):
                        row.append("Break")
                    else:
                        row.append(a.get((d, t, nm), "Unassigned"))
                ws_tl.append(row)

        for rr in range(2, ws_tl.max_row + 1):
            for cc in range(3, ws_tl.max_column + 1):
                val = str(ws_tl.cell(rr, cc).value or "")
                ws_tl.cell(rr, cc).fill = fill_for(val)
                ws_tl.cell(rr, cc).alignment = Alignment(wrap_text=True, vertical="top")

        # Timelines by site
        for site in ("SLGP", "JEN", "BGS"):
            site_staff = list(staff_df.loc[staff_df["HomeSite"].astype(str).str.upper() == site, "Name"].astype(str))
            if not site_staff:
                continue
            ws_site = wb.create_sheet(f"Week{w+1}_{site}_Timelines")
            ws_site.append(["Date", "Time"] + site_staff)
            for c in ws_site[1]:
                c.font = Font(bold=True)
                c.alignment = Alignment(horizontal="center", vertical="center")
            ws_site.freeze_panes = "C2"

            for d in dates:
                for t in slots:
                    row = [d.strftime("%a %d-%b"), t.strftime("%H:%M")]
                    for nm in site_staff:
                        hk = holiday_kind(nm, d, hols)
                        if hk:
                            row.append(hk)
                        elif not is_working(hmap, week_start, d, t, nm):
                            row.append("")  # blank grey
                        elif nm in breaks.get((d, t), set()):
                            row.append("Break")
                        else:
                            row.append(a.get((d, t, nm), "Unassigned"))
                    ws_site.append(row)

            for rr in range(2, ws_site.max_row + 1):
                for cc in range(3, ws_site.max_column + 1):
                    val = str(ws_site.cell(rr, cc).value or "")
                    ws_site.cell(rr, cc).fill = fill_for(val)
                    ws_site.cell(rr, cc).alignment = Alignment(wrap_text=True, vertical="top")



        # Per-site task hours by staff (for KPIs)
        # We use HomeSite from staff_df to allocate people to SLGP/JEN/BGS sheets.
        site_map = {}
        for _, row in staff_df.iterrows():
            nm = str(row["Name"]).strip()
            hs = str(row.get("HomeSite", "")).upper()
            if nm and hs in ("SLGP", "JEN", "BGS"):
                site_map[nm] = hs

        # Build a tidy per-staff task hours table from pivot_w
        # pivot_w index = Name, columns = tasks
        if not pivot_w.empty:
            # Normalise columns we care about
            wanted_tasks = [
                "Phones",
                "Bookings",
                "EMIS/Docman Tasks",
                "FrontDesk_SLGP",
                "FrontDesk_JEN",
                "FrontDesk_BGS",
                "Triage_Admin_SLGP",
                "Triage_Admin_JEN",
                "Email_Box",
                "Awaiting_PSA_Admin",
                "Additional Capacity",
            ]
            for col in wanted_tasks:
                if col not in pivot_w.columns:
                    pivot_w[col] = 0.0

            # Collapse front desk + triage into single columns
            pivot_w["Front Desk"] = (
                pivot_w["FrontDesk_SLGP"]
                + pivot_w["FrontDesk_JEN"]
                + pivot_w["FrontDesk_BGS"]
            )
            pivot_w["Triage Admin"] = (
                pivot_w["Triage_Admin_SLGP"]
                + pivot_w["Triage_Admin_JEN"]
            )

            # PSA Admin / Awaiting Response stays as Awaiting_PSA_Admin
            pivot_w["PSA/AR"] = pivot_w["Awaiting_PSA_Admin"]

            # Final per-staff task table
            task_cols = [
                "Phones",
                "Bookings",
                "EMIS/Docman Tasks",
                "Front Desk",
                "Triage Admin",
                "Email_Box",
                "PSA/AR",
                "Additional Capacity",
            ]

            per_staff = pivot_w[task_cols].copy()
            per_staff = per_staff.reset_index()  # bring Name out of index
            per_staff["Total"] = per_staff[task_cols].sum(axis=1)

            # Create one sheet per site
            for site in ("SLGP", "JEN", "BGS"):
                ws_site_tasks = wb.create_sheet(f"Week{w+1}_{site}_TaskByStaff")
                ws_site_tasks.append(
                    ["Staff", "Phones", "Bookings", "EMIS/Docman Tasks",
                     "Front Desk", "Triage Admin", "Email", "PSA/AR",
                     "Additional Capacity", "Total"]
                )
                for c in ws_site_tasks[1]:
                    c.font = Font(bold=True)
                    c.alignment = Alignment(horizontal="center", vertical="center")

                for _, row in per_staff.iterrows():
                    nm = row["Name"]
                    hs = site_map.get(nm)
                    if hs != site:
                        continue
                    ws_site_tasks.append([
                        nm,
                        float(row["Phones"]),
                        float(row["Bookings"]),
                        float(row["EMIS/Docman Tasks"]),
                        float(row["Front Desk"]),
                        float(row["Triage Admin"]),
                        float(row["Email_Box"]),
                        float(row["PSA/AR"]),
                        float(row["Additional Capacity"]),
                        float(row["Total"]),
                    ])

        # Weekly KPI summary for Bookings and EMIS/Docman
        ws_kpi = wb.create_sheet(f"Week{w+1}_KPI_Summary")
        ws_kpi.append(["Task", "WeekHoursTarget", "WeekHoursAchieved", "WeekGap"])
        for c in ws_kpi[1]:
            c.font = Font(bold=True)
            c.alignment = Alignment(horizontal="center", vertical="center")

        # Fixed weekly targets as requested
        weekly_targets = {
            "Bookings": 128.0,
            "EMIS/Docman Tasks": 35.0,
        }

        # Compute achieved hours from df (post-merged EMIS/Docman)
        week_hours = {"Bookings": 0.0, "EMIS/Docman Tasks": 0.0}
        if not df.empty:
            for _, row in df.iterrows():
                task = str(row["Task"])
                h = float(row["Hours"])
                if task == "Bookings":
                    week_hours["Bookings"] += h
                elif task == "EMIS/Docman Tasks":
                    week_hours["EMIS/Docman Tasks"] += h

        for task in ["Bookings", "EMIS/Docman Tasks"]:
            tgt = weekly_targets.get(task, 0.0)
            ach = week_hours.get(task, 0.0)
            ws_kpi.append([task, tgt, ach, ach - tgt])

        # Overall task hours summary for the week (all tasks)
        ws_th = wb.create_sheet(f"Week{w+1}_TaskHours")
        ws_th.append(["Task", "WeekHoursAchieved"])
        for c in ws_th[1]:
            c.font = Font(bold=True)
            c.alignment = Alignment(horizontal="center", vertical="center")

        task_hours = {}
        if not df.empty:
            for _, row in df.iterrows():
                tname = str(row["Task"])
                h = float(row["Hours"])
                task_hours[tname] = task_hours.get(tname, 0.0) + h
        for tname, h in sorted(task_hours.items()):
            ws_th.append([tname, h])


        # Safety alerts (after 16:00, require at least one Front Desk + one other staff on site)
        ws_safe = wb.create_sheet(f"Week{w+1}_SafetyAlerts")
        ws_safe.append(["Date", "Time", "Site", "FrontDeskCount", "NonFrontDeskCount", "Status"])
        for c in ws_safe[1]:
            c.font = Font(bold=True)
            c.alignment = Alignment(horizontal="center", vertical="center")

        for d in dates:
            # Only Mon–Fri (0=Mon..4=Fri)
            if d.weekday() > 4:
                continue
            for t in slots:
                # Only from 16:00 onwards
                if t < time(16, 0):
                    continue
                # Count staff per site
                site_fd = {"SLGP": 0, "JEN": 0, "BGS": 0}
                site_non = {"SLGP": 0, "JEN": 0, "BGS": 0}
                for nm in staff_names:
                    hk = holiday_kind(nm, d, hols)
                    if hk:
                        continue
                    if not is_working(hmap, week_start, d, t, nm):
                        continue
                    role = a.get((d, t, nm), "")
                    sr = staff_by_name[nm]
                    home = str(sr.get("HomeSite", "")).upper()
                    if home not in site_fd:
                        continue
                    if role.startswith("FrontDesk_"):
                        site_fd[home] += 1
                    elif role:
                        site_non[home] += 1

                for site in ["SLGP", "JEN", "BGS"]:
                    fd_count = site_fd[site]
                    non_count = site_non[site]
                    if fd_count >= 1 and non_count >= 1:
                        status = "SAFE"
                    elif fd_count == 0 and non_count == 0:
                        status = "No staff on site"
                    elif fd_count == 0:
                        status = "No Front Desk cover"
                    else:
                        status = "Front Desk lone worker"
                    if status != "SAFE":
                        ws_safe.append([
                            d.isoformat(),
                            t.strftime("%H:%M"),
                            site,
                            fd_count,
                            non_count,
                            status,
                        ])

        # Simple per-site headcount heatmap (numbers only, user can apply conditional formatting)
        ws_hm = wb.create_sheet(f"Week{w+1}_SiteHeadcount")
        ws_hm.append(["Date", "Time", "SLGP_Count", "JEN_Count", "BGS_Count"])
        for c in ws_hm[1]:
            c.font = Font(bold=True)
            c.alignment = Alignment(horizontal="center", vertical="center")

        for d in dates:
            for t in slots:
                counts = {"SLGP": 0, "JEN": 0, "BGS": 0}
                for nm in staff_names:
                    hk = holiday_kind(nm, d, hols)
                    if hk:
                        continue
                    if not is_working(hmap, week_start, d, t, nm):
                        continue
                    sr = staff_by_name[nm]
                    home = str(sr.get("HomeSite", "")).upper()
                    if home in counts:
                        counts[home] += 1
                ws_hm.append([
                    d.isoformat(),
                    t.strftime("%H:%M"),
                    counts["SLGP"],
                    counts["JEN"],
                    counts["BGS"],
                ])


        # Phones targets vs achieved for PSA phones only (gap vs target)
        ws_ph = wb.create_sheet(f"Week{w+1}_Phones_Gap")
        ws_ph.append(["Date", "Time", "Target", "PSA_Achieved", "Gap"])
        for c in ws_ph[1]:
            c.font = Font(bold=True)
            c.alignment = Alignment(horizontal="center", vertical="center")

        for d in dates:
            dow = d.weekday()  # 0=Mon
            for tslot in slots:
                # Map each 30-min slot to the hour bucket in PHONES_TARGETS
                hour_key = tslot.hour
                base_target = PHONES_TARGETS.get(dow, {}).get(hour_key, 0)

                # Count PSA phones achieved from rota
                psa_achieved = 0
                for nm in staff_names:
                    hk = holiday_kind(nm, d, hols)
                    if hk:
                        continue
                    if not is_working(hmap, week_start, d, tslot, nm):
                        continue
                    role = a.get((d, tslot, nm), "")
                    if role == "Phones":
                        psa_achieved += 1
                gap = psa_achieved - base_target
                ws_ph.append([
                    d.isoformat(),
                    tslot.strftime("%H:%M"),
                    base_target,
                    psa_achieved,
                    gap,
                ])
        # Weekly targets vs achieved for Bookings and EMIS/Docman (hours)
        ws_wtg = wb.create_sheet(f"Week{w+1}_TaskTargets")
        ws_wtg.append(["Task", "WeekHoursTarget", "WeekHoursAchieved", "WeekGap"])
        for c in ws_wtg[1]:
            c.font = Font(bold=True)
            c.alignment = Alignment(horizontal="center", vertical="center")

        # Calculate achieved hours from df (already built above)
        week_hours = {"Bookings": 0.0, "EMIS/Docman Tasks": 0.0}
        if not df.empty:
            for _, row in df.iterrows():
                task = str(row["Task"])
                h = float(row["Hours"])
                if task == "Bookings":
                    week_hours["Bookings"] += h
                elif task == "EMIS/Docman Tasks":
                    week_hours["EMIS/Docman Tasks"] += h

        for task in ["Bookings", "EMIS/Docman Tasks"]:
            ach = week_hours.get(task, 0.0)
            ws_wtg.append([task, 0.0, ach, ach - 0.0])
    return wb

# =========================================================
# Streamlit UI
# =========================================================
st.set_page_config(page_title="Rota Generator v14", layout="wide")
require_password()

st.title("Rota Generator v14 — PSA Workforce-aligned rota")

st.markdown(
    """
### How to use
1. Download the starter template below.
2. Fill in **Staff**, **WorkingHours**, and (optional) **Holidays**.
3. Optional: add task **weights (0–5)** per staff member to prefer stronger/faster people for specific tasks.
4. Upload the completed template and click **Generate** to download the rota as Excel.

Weights are **tie-breakers only** – if left blank, behaviour defaults to neutral (3).
"""
)

# Template download (must exist in repo root)
tpl_bytes = None
for p in ("rota_generator_template.xlsx", "rota_generator_template_v2.xlsx"):
    try:
        with open(p, "rb") as f:
            tpl_bytes = f.read()
        break
    except Exception:
        continue

if tpl_bytes:
    st.download_button(
        "📥 Download starter template",
        data=tpl_bytes,
        file_name="rota_generator_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
    st.info("Add 'rota_generator_template.xlsx' to your repo root to enable template download.")

uploaded = st.file_uploader("Upload your completed template (.xlsx)", type=["xlsx"])

c1, c2 = st.columns(2)
with c1:
    start_date = st.date_input("Week commencing (Monday)", value=date.today())
with c2:
    mode = st.selectbox("Run period", ["1 week", "4 weeks (month)", "Custom # weeks"])

if mode == "Custom # weeks":
    weeks = int(st.number_input("Number of weeks", min_value=1, max_value=12, value=2, step=1))
elif mode == "4 weeks (month)":
    weeks = 4
else:
    weeks = 1

start_monday = ensure_monday(start_date)

if uploaded:
    try:
        staff_df, hours_df, hols, found = read_template(uploaded.getvalue())
        st.success(f"Template loaded. Sheets: Staff={found['staff']} | Hours={found['hours']} | Holidays={found['holidays'] or 'None'}")
        if st.button("Generate rota and download Excel", type="primary"):
            wb = build_workbook(staff_df, hours_df, hols, start_monday, weeks)
            bio = io.BytesIO()
            wb.save(bio)
            bio.seek(0)
            out_name = f"rota_{start_monday.isoformat()}_{weeks}w.xlsx"
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