import streamlit as st
import pandas as pd
import openpyxl
from openpyxl import load_workbook
import io
import json
import re
import requests
from datetime import datetime, date
import traceback

# ─── Page config ────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Rent Roll AI Agent",
    page_icon="🏢",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
    .main { padding-top: 0.5rem; }
    .log-box {
        background: #0f1117;
        border: 1px solid #2d3748;
        border-radius: 8px;
        padding: 1rem 1.2rem;
        font-family: 'Courier New', monospace;
        font-size: 0.78rem;
        max-height: 400px;
        overflow-y: auto;
        color: #e2e8f0;
        line-height: 1.6;
    }
    .log-step  { color: #63b3ed; }
    .log-think { color: #f6ad55; }
    .log-action{ color: #68d391; }
    .log-warn  { color: #fc8181; }
    .log-ok    { color: #9ae6b4; }
</style>
""", unsafe_allow_html=True)

TEMPLATE_PATH = "Rent_Roll_template.xlsx"

# ─── Rent charge-code patterns ───────────────────────────────────────────────
RENT_INCLUDE = [
    r"^rent$", r"^rnt$", r"^base$", r"^contract$", r"^net$",
    r"^hap$", r"^sec8$", r"^s8$", r"^hud$", r"^rnta$",
    r"^subsidy$", r"^sub$", r"^ttp$", r"^tenant$", r"^tr$",
    r"^lihtc$", r"^credit$", r"^bmr$", r"^aff$", r"^usda$", r"^rd$",
    r"^rentmkt$", r"^stl$",
    r"\brent\b", r"\blease rent\b", r"\bbase rent\b",
    r"\bcontract rent\b", r"\bhud\b", r"\bhap\b",
    r"^rent-", r"^hudr",
]
RENT_EXCLUDE = [
    r"pet", r"reno", r"park", r"garag", r"storag", r"trash", r"water", r"sewer",
    r"electric", r"gas", r"internet", r"cable", r"utility", r"admin",
    r"pest", r"valet", r"amenity", r"concession", r"late\s*fee", r"nsf",
    r"real\s*estate\s*tax", r"pack", r"locker", r"waste",
    r"building\s*facilit", r"reimb", r"comfee", r"feeamty",
    r"ltor", r"ltol", r"loss\s*to\s*old", r"^real$",
]

def is_rent_code(code):
    if not code: return False
    cl = str(code).strip().lower()
    for ex in RENT_EXCLUDE:
        if re.search(ex, cl): return False
    for pat in RENT_INCLUDE:
        if re.search(pat, cl, re.IGNORECASE): return True
    return False

SKIP_KW  = ["pending renewal","future resident","applicant","future","prospect",
             "notice to vacate","previous","former"]
VACAT_STATUS_KW = ["vacant","model","down unit","offline","-- vacant --","excluded"]

def should_skip(status):
    sl = str(status or "").lower()
    return any(kw in sl for kw in SKIP_KW)

def is_vacant(status, name):
    sl = str(status or "").lower()
    nl = str(name or "").lower()
    for kw in VACAT_STATUS_KW:
        if kw in sl: return True
    # Name-based: "-- vacant --", or name starts with "Model" or "Vacant" (system placeholders)
    if nl in ("-- vacant --", "vacant"):
        return True
    if re.match(r"^(model|vacant)\s", nl):   # e.g. "Model R00000008" but NOT "Model, Model"
        return True
    return False

def fmt_date(val):
    if val is None: return ""
    if isinstance(val, (datetime, date)):
        return val.strftime("%m-%d-%Y")
    s = str(val).strip()
    for fmt in ["%m/%d/%Y","%Y-%m-%d","%m-%d-%Y","%d/%m/%Y","%m/%d/%y"]:
        try: return datetime.strptime(s, fmt).strftime("%m-%d-%Y")
        except: pass
    try:
        d = pd.to_datetime(s, errors="coerce")
        if not pd.isna(d): return d.strftime("%m-%d-%Y")
    except: pass
    return ""

def to_num(val):
    if val is None: return None
    if isinstance(val, (int, float)): return float(val)
    try: return float(str(val).replace(",","").replace("$","").strip())
    except: return None

def clean_type(val):
    if not val: return ""
    return re.sub(r"^unit\s*type\s*[:\-]\s*","",str(val),flags=re.IGNORECASE).strip()

def gcol(row, idx, default=None):
    if idx is None or idx >= len(row): return default
    return row[idx]

# ─── Claude API ───────────────────────────────────────────────────────────────
def call_claude(messages, system, max_tokens=2000):
    api_key = st.secrets.get("ANTHROPIC_API_KEY", "")
    resp = requests.post(
        "https://api.anthropic.com/v1/messages",
        headers={"Content-Type":"application/json","x-api-key": api_key,
                 "anthropic-version":"2023-06-01"},
        json={"model":"claude-sonnet-4-20250514","max_tokens":max_tokens,
              "system":system,"messages":messages},
    )
    data = resp.json()
    if "content" not in data: raise ValueError(f"API error: {data}")
    return "".join(b.get("text","") for b in data["content"] if b.get("type")=="text")

# ─── Parsers ─────────────────────────────────────────────────────────────────

def _finalize(current, units):
    if not current: return
    if current.get("_vacant"):
        current["effective_rent"] = None
    elif current.get("_charge_sum",0) > 0 and not current.get("effective_rent"):
        current["effective_rent"] = current["_charge_sum"]
    current.pop("_charge_sum", None)
    current.pop("_vacant", None)
    units.append(current)

def find_header(rows, keywords, max_search=20):
    for i, row in enumerate(rows[:max_search]):
        rs = " ".join(str(v or "").lower() for v in row)
        if all(kw in rs for kw in keywords):
            return i
    return None

def hdr_map(headers, *kws):
    for kw in kws:
        for i, h in enumerate(headers):
            if kw in str(h or "").lower().replace("\n"," "):
                return i
    return None


class YardiParser:
    """Yardi Voyager / Breeze: single header row, one data row per unit."""
    name = "Yardi"

    def can_handle(self, rows, fname):
        for row in rows[:12]:
            rs = " ".join(str(v or "").lower() for v in row)
            if ("resh id" in rs or "unit/lease status" in rs or
                ("unit" in rs and "floorplan" in rs and "sqft" in rs and "lease rent" in rs)):
                return True
        return False

    def extract(self, rows):
        hi = find_header(rows, ["unit","status"], max_search=15)
        if hi is None: return []
        h = [str(v or "").lower().replace("\n"," ").strip() for v in rows[hi]]
        c = dict(
            unit   = hdr_map(h,"unit","bldg-unit"),
            fp     = hdr_map(h,"floorplan","unit type","floor plan"),
            sqft   = hdr_map(h,"sqft","sq ft","size"),
            status = hdr_map(h,"status","designation"),
            name   = hdr_map(h,"name"),
            movin  = hdr_map(h,"move-in","move in"),
            lstart = hdr_map(h,"lease start","lease\nstart"),
            lend   = hdr_map(h,"lease end","lease\nend"),
            mkt    = hdr_map(h,"market rent","market\nrent"),
            eff    = hdr_map(h,"lease rent","effective rent","actual rent"),
        )
        units = []
        for row in rows[hi+1:]:
            # Detect mid-file section dividers / repeated header rows in OneSite
            r1_val = row[1] if len(row) > 1 else None
            r1_str = str(r1_val or "").strip()
            if r1_str == "details" or r1_str == "\nUnit" or (
                r1_val is None and len(row) > 36 and row[36] and 
                "trans" in str(row[36]).lower()
            ):
                # Mid-file section divider - skip but keep accumulating for current unit
                continue

            uv = gcol(row, c["unit"])
            if not uv: continue
            uvs = str(uv).strip()
            if not uvs or uvs.lower() in ["unit","none","nan"]: continue
            if any(kw in uvs.lower() for kw in ["total","summary","subtotal","grand","future residents"]): break
            status = str(gcol(row, c["status"]) or "")
            name   = str(gcol(row, c["name"])   or "")
            if should_skip(status): continue
            vacat  = is_vacant(status, name)
            eff    = to_num(gcol(row, c["eff"]))
            if vacat: eff = None
            units.append({
                "unit_no":       uvs,
                "unit_type":     clean_type(gcol(row, c["fp"])),
                "sqft":          to_num(gcol(row, c["sqft"])),
                "status":        status,
                "resident_name": name if not vacat else "",
                "move_in":       fmt_date(gcol(row, c["movin"])),
                "lease_start":   fmt_date(gcol(row, c["lstart"])),
                "lease_end":     fmt_date(gcol(row, c["lend"])),
                "market_rent":   to_num(gcol(row, c["mkt"])),
                "effective_rent": eff,
                "building": None,
            })
        return units


class OneSiteParser:
    """RealPage OneSite: wide columns (58), unit row + charge rows, \n in headers."""
    name = "OneSite/RealPage"

    def can_handle(self, rows, fname):
        for row in rows[:5]:
            rs = " ".join(str(v or "").lower() for v in row)
            if "onesite" in rs or "onsite" in rs: return True
        return False

    def _find_header_row(self, rows):
        for i, row in enumerate(rows[:15]):
            for v in row:
                if v and "\nUnit" in str(v) or (v and "unit" in str(v).lower() and "\n" in str(v)):
                    return i
            # Also check row with Trans Code
            rs = " ".join(str(v or "") for v in row)
            if "Trans" in rs and "Code" in rs and "Floorplan" in rs:
                return i
        return None

    def extract(self, rows):
        hi = self._find_header_row(rows)
        if hi is None: return []
        h = rows[hi]
        c = dict(
            unit   = hdr_map(h,"\nunit","unit\n"),
            fp     = hdr_map(h,"floorplan","floor plan"),
            sqft   = hdr_map(h,"sqft","sq ft"),
            status = hdr_map(h,"unit/lease\nstatus","status"),
            name   = hdr_map(h,"name\n","\nname"),
            movin  = hdr_map(h,"move-in\nmove-out","move-in"),
            lstart = hdr_map(h,"lease\nstart"),
            lend   = hdr_map(h,"lease\nend"),
            mkt    = hdr_map(h,"market\n+ addl.","market rent","market\nrent"),
            tcode  = hdr_map(h,"trans\ncode","code"),
            lrent  = hdr_map(h,"lease\nrent","lease rent"),
        )
        # Fallback column positions for known OneSite layout
        if c["unit"] is None:   c["unit"]   = 1
        if c["fp"] is None:     c["fp"]     = 3
        if c["sqft"] is None:   c["sqft"]   = 14
        if c["status"] is None: c["status"] = 18
        if c["name"] is None:   c["name"]   = 20
        if c["movin"] is None:  c["movin"]  = 24
        if c["lstart"] is None: c["lstart"] = 28
        if c["lend"] is None:   c["lend"]   = 30
        if c["mkt"] is None:    c["mkt"]    = 33
        if c["tcode"] is None:  c["tcode"]  = 36
        if c["lrent"] is None:  c["lrent"]  = 43

        # OneSite has a known column-offset artifact after unmerging:
        # 'Lease Rent' header may land at idx N but actual values sit at idx N+1.
        # Probe first charge row to auto-detect the real data column.
        for probe_row in rows[hi+1:hi+50]:
            tcode_val = gcol(probe_row, c["tcode"])
            if tcode_val and is_rent_code(str(tcode_val)):
                for offset in range(c["tcode"]+1, min(c["tcode"]+12, len(probe_row))):
                    v = probe_row[offset]
                    if v is not None and isinstance(v, (int, float)) and float(v) > 0:
                        c["lrent"] = offset
                        break
                break

        units = []
        current = None
        for row in rows[hi+1:]:
            # Skip mid-file section-divider rows (OneSite re-prints "details" + headers mid-file)
            r1_raw = row[1] if len(row) > 1 else None
            r1_str = str(r1_raw or "").strip()
            if r1_str in ("details",) or r1_str == "\nUnit":
                continue   # keep current unit alive, charges continue after header

            uv = gcol(row, c["unit"])
            if uv:
                uvs = str(uv).strip()
                if not uvs or uvs.lower() in ["unit","none","nan"]: continue
                if any(kw in uvs.lower() for kw in ["total","summary"]): break
                _finalize(current, units)
                name   = str(gcol(row, c["name"]) or "")
                status = str(gcol(row, c["status"]) or "")
                if should_skip(status): current=None; continue
                vacat  = is_vacant(status, name)
                movin_raw = re.split(r"_x000D_|\\r", str(gcol(row, c["movin"]) or ""))[0].strip()
                sqft_raw  = str(gcol(row, c["sqft"]) or "").replace(",","")
                cc = str(gcol(row, c["tcode"]) or "")
                ca = to_num(gcol(row, c["lrent"]))
                current = {
                    "unit_no":       uvs,
                    "unit_type":     clean_type(gcol(row, c["fp"])),
                    "sqft":          to_num(sqft_raw) if sqft_raw else None,
                    "status":        status,
                    "resident_name": name if not vacat else "",
                    "move_in":       fmt_date(movin_raw),
                    "lease_start":   fmt_date(gcol(row, c["lstart"])),
                    "lease_end":     fmt_date(gcol(row, c["lend"])),
                    "market_rent":   to_num(gcol(row, c["mkt"])),
                    "effective_rent": None,
                    "_charge_sum":   0,
                    "_vacant":       vacat,
                    "building": None,
                }
                if cc and ca and is_rent_code(cc): current["_charge_sum"] += ca or 0
            else:
                if current is None: continue
                # Skip future-tenant sub-rows embedded in current unit block
                row_str = " ".join(str(v or "").lower() for v in row if v)
                if any(kw in row_str for kw in ["applicant","pending renewal","pending renew"]):
                    current["_skip_future"] = True
                    continue
                if current.get("_skip_future"):
                    continue
                cc = str(gcol(row, c["tcode"]) or "")
                ca = to_num(gcol(row, c["lrent"]))
                if cc and ca and is_rent_code(cc): current["_charge_sum"] += ca or 0
        _finalize(current, units)
        return units


class MRIParser:
    """MRI Living: 13-14 col, 2-row headers.
    Variant A (14 col): multi-row charges, col6=ChargeCode, col7=Amount (Stone_Loch, Canal)
    Variant B (13 col): single effective rent, col6=ActualRent (Retreat_at_the_Park)
    """
    name = "MRI Living"

    def can_handle(self, rows, fname):
        # MRI has a 2-row split header: row N has "Unit|Unit Type|Unit|Resident|Name|Market|Charge|Amount"
        # and row N+1 has "Sq Ft" in position 2 completing the split.
        # Also: title row contains "Rent Roll with Lease Charges" or "As Of ="
        has_title = any("rent roll with lease charges" in " ".join(str(v or "").lower() for v in r) or
                        "as of =" in " ".join(str(v or "").lower() for v in r)
                        for r in rows[:6])
        if not has_title:
            return False
        for i, row in enumerate(rows[3:10], 3):
            rs = " ".join(str(v or "").lower() for v in row)
            if ("unit" in rs and ("charge" in rs or "actual" in rs) and "market" in rs
                    and "bldg-unit" not in rs and "bldg" not in rs):
                # Check next row for "sq ft"
                if i+1 < len(rows):
                    r2 = " ".join(str(v or "").lower() for v in rows[i+1])
                    if "sq ft" in r2:
                        return True
        return False

    def _detect_variant(self, rows):
        """Returns 'charges' or 'actual' based on header row."""
        for row in rows[4:8]:
            rs = " ".join(str(v or "").lower() for v in row)
            if "charge" in rs and "amount" in rs:
                return "charges"
            if "actual" in rs:
                return "actual"
        return "charges"

    def _find_data_start(self, rows):
        for i, row in enumerate(rows[:20]):
            r0 = str(row[0] or "").strip()
            if r0 and row[4] and not any(kw in r0.lower() for kw in
               ["rent roll","as of","month","unit","current","notice","vacant","total","summary"]):
                if re.match(r"[\w\d\-]+", r0):
                    return i
        return 7

    def extract(self, rows):
        variant    = self._detect_variant(rows)
        data_start = self._find_data_start(rows)

        units = []
        current = None

        for row in rows[data_start:]:
            r0 = str(row[0] or "").strip()
            r4 = str(row[4] or "").strip() if len(row) > 4 else ""
            if any(kw in r0.lower() for kw in ["total","summary","grand","future residents","future resident"]): break
            # Section headers
            if r0 and not r4 and all(row[j] is None for j in range(1, min(5, len(row)))):
                continue

            is_new = bool(r0 and r4)
            if is_new:
                _finalize(current, units)
                name  = r4
                vacat = is_vacant("", name)
                mkt   = to_num(row[5]) if len(row) > 5 else None

                if variant == "actual":
                    # col 6 = actual rent amount directly
                    eff = to_num(row[6]) if len(row) > 6 else None
                    if vacat: eff = None
                    units.append({
                        "unit_no":       r0,
                        "unit_type":     clean_type(row[1] if len(row) > 1 else None),
                        "sqft":          to_num(row[2]) if len(row) > 2 else None,
                        "status":        "",
                        "resident_name": name if not vacat else "",
                        "move_in":       fmt_date(row[9] if len(row) > 9 else None),
                        "lease_start":   "",
                        "lease_end":     fmt_date(row[10] if len(row) > 10 else None),
                        "market_rent":   mkt,
                        "effective_rent": eff,
                        "building": None,
                    })
                    current = None
                else:
                    # Variant A: multi-row charges
                    cc = str(row[6] or "").strip() if len(row) > 6 else ""
                    ca = to_num(row[7]) if len(row) > 7 else None
                    current = {
                        "unit_no":       r0,
                        "unit_type":     clean_type(row[1] if len(row) > 1 else None),
                        "sqft":          to_num(row[2]) if len(row) > 2 else None,
                        "status":        "",
                        "resident_name": name if not vacat else "",
                        "move_in":       fmt_date(row[10] if len(row) > 10 else None),
                        "lease_start":   "",
                        "lease_end":     fmt_date(row[11] if len(row) > 11 else None),
                        "market_rent":   mkt,
                        "effective_rent": None,
                        "_charge_sum":   0,
                        "_vacant":       vacat,
                        "building": None,
                    }
                    if cc and ca and is_rent_code(cc): current["_charge_sum"] += ca or 0
            else:
                if current is None or variant == "actual": continue
                cc = str(row[6] or "").strip() if len(row) > 6 else ""
                ca = to_num(row[7]) if len(row) > 7 else None
                if cc and ca and "total" not in cc.lower() and is_rent_code(cc):
                    current["_charge_sum"] += ca or 0

        if variant == "charges":
            _finalize(current, units)
        return units


class AppFolioParser:
    """AppFolio / ReNew style: unit type in section headers, multi-row charges."""
    name = "AppFolio"

    def can_handle(self, rows, fname):
        for row in rows[:12]:
            rs = " ".join(str(v or "").lower() for v in row)
            if ("bldg-unit" in rs or "unit details" in rs) and ("charge code" in rs or "scheduled" in rs):
                return True
        return False

    def extract(self, rows):
        hi = find_header(rows, ["bldg-unit","resident","lease"], max_search=15)
        if hi is None:
            hi = find_header(rows, ["bldg","resident","lease start"], max_search=15)
        if hi is None: return []

        h = [str(v or "").lower().replace("\n"," ").strip() for v in rows[hi]]
        c = dict(
            unit   = hdr_map(h,"bldg-unit","unit"),
            sqft   = hdr_map(h,"sqft","sq ft"),
            status = hdr_map(h,"unit status","status"),
            res    = hdr_map(h,"resident"),
            movin  = hdr_map(h,"move-in","move in"),
            lstart = hdr_map(h,"lease start"),
            lend   = hdr_map(h,"lease end"),
            mkt    = hdr_map(h,"budgeted rent","market rent","market"),
            cc     = hdr_map(h,"charge code"),
            sched  = hdr_map(h,"scheduled charges","scheduled"),
            ledger = hdr_map(h,"ledger"),
        )

        units = []
        current_type = None
        current = None

        for row in rows[hi+1:]:
            r0 = str(row[0] or "").strip()
            # Section header: single populated cell that is not a unit number
            if r0 and all(row[j] is None for j in range(1, min(6, len(row)))):
                if not re.match(r"^\d", r0) and "total" not in r0.lower():
                    current_type = clean_type(r0)
                    continue

            # "Charge Total:" skip
            ledger_val = str(gcol(row, c["ledger"]) or "").lower()
            if "charge total" in ledger_val: continue

            uv = gcol(row, c["unit"])
            if not uv:
                # Charge-only row (no unit number) - accumulate into current unit
                if current is None: continue
                cc_v = str(gcol(row, c["cc"]) or "")
                ca_v = to_num(gcol(row, c["sched"]))
                if cc_v and ca_v and is_rent_code(cc_v): current["_charge_sum"] += ca_v or 0
                continue

            uvs = str(uv).strip()
            if not uvs or uvs.lower() in ["unit","bldg-unit","none","nan"]: continue
            if any(kw in uvs.lower() for kw in ["future residents","future resident"]): break
            # Skip section subtotals like "Auburn Total:" - only stop on grand/property totals
            if re.match(r"^(grand total|property total|total units|total$)", uvs.lower()): break
            if uvs.lower() in ["total","summary","subtotal"]: break

            res_val    = gcol(row, c["res"])
            status_val = str(gcol(row, c["status"]) or "")
            is_unit_row = bool(status_val or res_val)

            if is_unit_row and re.match(r"[\w\d\-]", uvs):
                _finalize(current, units)
                if should_skip(status_val): current=None; continue
                name  = str(res_val or "")
                vacat = is_vacant(status_val, name)
                cc_v  = str(gcol(row, c["cc"])    or "")
                ca_v  = to_num(gcol(row, c["sched"]))
                current = {
                    "unit_no":       uvs,
                    "unit_type":     current_type or "",
                    "sqft":          to_num(gcol(row, c["sqft"])),
                    "status":        status_val,
                    "resident_name": name if not vacat else "",
                    "move_in":       fmt_date(gcol(row, c["movin"])),
                    "lease_start":   fmt_date(gcol(row, c["lstart"])),
                    "lease_end":     fmt_date(gcol(row, c["lend"])),
                    "market_rent":   to_num(gcol(row, c["mkt"])),
                    "effective_rent": None,
                    "_charge_sum":   0,
                    "_vacant":       vacat,
                    "building": None,
                }
                if cc_v and ca_v and is_rent_code(cc_v): current["_charge_sum"] += ca_v or 0
            else:
                # Row has a unit number but no status/res -> inline charge row
                if current is None: continue
                cc_v = str(gcol(row, c["cc"]) or "")
                ca_v = to_num(gcol(row, c["sched"]))
                if cc_v and ca_v and is_rent_code(cc_v): current["_charge_sum"] += ca_v or 0
        _finalize(current, units)
        return units


class RentManagerParser:
    """Rent Manager / WAT: cols Unit, Type, Market, Name+ID+Date, LeaseDates, ..., Desc, ChargeAmt."""
    name = "Rent Manager"

    def can_handle(self, rows, fname):
        for row in rows[:15]:
            rs = " ".join(str(v or "").lower() for v in row)
            if "lease dates" in rs and ("description" in rs or "charge amount" in rs):
                return True
        return False

    def extract(self, rows):
        hi = find_header(rows, ["lease dates","description"], max_search=20)
        if hi is None: return []

        units = []
        current = None

        for row in rows[hi+1:]:
            r0 = str(row[0] or "").strip()
            if any(kw in r0.lower() for kw in ["total","summary","grand","future residents","future resident"]): break
            is_unit = bool(r0 and re.match(r"^\d+", r0) and len(row) > 1 and row[1])
            if is_unit:
                _finalize(current, units)
                unit_no   = r0
                unit_type = str(row[1] or "").split(",")[0].strip()
                mkt       = to_num(row[2]) if len(row) > 2 else None
                name_raw  = str(row[3] or "")
                name_m    = re.match(r"^([^T\d]+?)(?:\s+T\d+|$)", name_raw)
                name      = name_m.group(1).strip() if name_m else name_raw.split("T00")[0].strip()
                dates_str = str(row[4] or "") if len(row) > 4 else ""
                dates     = re.findall(r"\d{2}/\d{2}/\d{4}", dates_str)
                vacat     = is_vacant("", name)
                desc      = str(row[8] or "") if len(row) > 8 else ""
                ca        = to_num(row[9]) if len(row) > 9 else None
                current = {
                    "unit_no":       unit_no,
                    "unit_type":     unit_type,
                    "sqft":          None,
                    "status":        "",
                    "resident_name": name if not vacat else "",
                    "move_in":       "",
                    "lease_start":   fmt_date(dates[0]) if len(dates) > 0 else "",
                    "lease_end":     fmt_date(dates[1]) if len(dates) > 1 else "",
                    "market_rent":   mkt,
                    "effective_rent": None,
                    "_charge_sum":   0,
                    "_vacant":       vacat,
                    "building": None,
                }
                if desc and ca and is_rent_code(desc): current["_charge_sum"] += ca or 0
            else:
                if current is None: continue
                desc = str(row[8] or "") if len(row) > 8 else ""
                ca   = to_num(row[9]) if len(row) > 9 else None
                if desc and ca and is_rent_code(desc): current["_charge_sum"] += ca or 0
        _finalize(current, units)
        return units


class AIFallbackParser:
    """Uses Claude to map any unknown format, then generic extraction."""
    name = "AI (Claude)"

    def detect_schema(self, rows, filename):
        preview = []
        seen = 0
        for i, r in enumerate(rows):
            if any(v is not None for v in r):
                preview.append(f"Row {i}: {[str(v)[:35] if v is not None else None for v in r[:28]]}")
                seen += 1
            if seen >= 45: break

        system = "You are an expert multi-family rent roll data analyst. Return ONLY valid JSON. No markdown."
        prompt = f"""Analyze raw rent roll from '{filename}'.

{chr(10).join(preview)}

Return JSON:
{{
  "format": "software name",
  "header_row": <0-based int>,
  "data_start_row": <0-based int>,
  "structure": "one_row_per_unit" | "multi_row_charges",
  "cols": {{
    "unit_no": <int|null>, "unit_type": <int|null>, "sqft": <int|null>,
    "status": <int|null>, "resident_name": <int|null>,
    "move_in": <int|null>, "lease_start": <int|null>, "lease_end": <int|null>,
    "market_rent": <int|null>, "effective_rent": <int|null>,
    "charge_code": <int|null>, "charge_amount": <int|null>
  }},
  "building_col": <int|null>,
  "unit_type_in_section_header": true|false,
  "notes": "brief"
}}"""
        try:
            raw = call_claude([{"role":"user","content":prompt}], system, 1500)
            raw = re.sub(r"```[a-z]*\n?","",raw).strip("` \n")
            return json.loads(raw)
        except Exception as e:
            return {
                "format":"unknown","header_row":0,"data_start_row":1,
                "structure":"multi_row_charges",
                "cols":{"unit_no":0,"unit_type":1,"sqft":2,"status":3,
                        "resident_name":4,"move_in":5,"lease_start":6,
                        "lease_end":7,"market_rent":8,"effective_rent":9,
                        "charge_code":None,"charge_amount":None},
                "building_col":None,"unit_type_in_section_header":False,"notes":str(e),
            }

    def extract(self, rows, schema):
        cols       = schema.get("cols",{})
        structure  = schema.get("structure","multi_row_charges")
        data_start = schema.get("data_start_row",1)
        cur_bldg   = None
        units      = []
        current    = None

        def g(row, key, default=None):
            idx = cols.get(key)
            if idx is None or idx >= len(row): return default
            return row[idx]

        for i, row in enumerate(rows[data_start:], data_start):
            rstr = " ".join(str(v or "").lower() for v in row)
            if re.search(r"\b(grand total|total units|summary)\b", rstr): break

            bc = schema.get("building_col")
            if bc is not None and bc < len(row) and row[bc]:
                if all(row[j] is None for j in range(len(row)) if j != bc):
                    cur_bldg = str(row[bc]).strip()
                    continue

            uv = g(row,"unit_no")
            if not uv: continue
            uvs = str(uv).strip()
            if not uvs or uvs.lower() in ["unit","none","nan"]: continue
            if any(kw in uvs.lower() for kw in ["total","summary","subtotal"]): continue

            status = str(g(row,"status") or "")
            name   = str(g(row,"resident_name") or "")

            if structure == "one_row_per_unit":
                if should_skip(status): continue
                vacat = is_vacant(status, name)
                eff   = to_num(g(row,"effective_rent"))
                if vacat: eff = None
                units.append({
                    "unit_no":       uvs,
                    "unit_type":     clean_type(g(row,"unit_type") or cur_bldg),
                    "sqft":          to_num(g(row,"sqft")),
                    "status":        status,
                    "resident_name": name if not vacat else "",
                    "move_in":       fmt_date(g(row,"move_in")),
                    "lease_start":   fmt_date(g(row,"lease_start")),
                    "lease_end":     fmt_date(g(row,"lease_end")),
                    "market_rent":   to_num(g(row,"market_rent")),
                    "effective_rent": eff,
                    "building": cur_bldg,
                })
            else:  # multi_row_charges
                is_new = bool(uv and (g(row,"resident_name") or g(row,"market_rent")))
                if is_new:
                    _finalize(current, units)
                    if should_skip(status): current=None; continue
                    vacat = is_vacant(status, name)
                    cc = str(g(row,"charge_code") or "")
                    ca = to_num(g(row,"charge_amount"))
                    current = {
                        "unit_no":       uvs,
                        "unit_type":     clean_type(g(row,"unit_type") or cur_bldg),
                        "sqft":          to_num(g(row,"sqft")),
                        "status":        status,
                        "resident_name": name if not vacat else "",
                        "move_in":       fmt_date(g(row,"move_in")),
                        "lease_start":   fmt_date(g(row,"lease_start")),
                        "lease_end":     fmt_date(g(row,"lease_end")),
                        "market_rent":   to_num(g(row,"market_rent")),
                        "effective_rent": to_num(g(row,"effective_rent")),
                        "_charge_sum":   0,
                        "_vacant":       vacat,
                        "building": cur_bldg,
                    }
                    if cc and ca and is_rent_code(cc): current["_charge_sum"] += ca or 0
                else:
                    if current is None: continue
                    cc = str(g(row,"charge_code") or "")
                    ca = to_num(g(row,"charge_amount"))
                    if cc and ca and is_rent_code(cc): current["_charge_sum"] += ca or 0

        if structure == "multi_row_charges":
            _finalize(current, units)
        return units


# ─── Agent ───────────────────────────────────────────────────────────────────
class RentRollAgent:

    PARSERS = [
        OneSiteParser(),
        RentManagerParser(),
        MRIParser(),
        AppFolioParser(),
        YardiParser(),
    ]

    def __init__(self, log_fn=None):
        self.log = log_fn or (lambda l, m: None)

    def _find_raw_sheet(self, wb):
        for s in wb.sheetnames:
            if "raw" in s.lower(): return s
        for s in wb.sheetnames:
            if "rent roll" not in s.lower(): return s
        return wb.sheetnames[0]

    def _to_rows(self, ws, max_rows=3000):
        rows = []
        for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, max_rows), values_only=True):
            rows.append(list(row))
        return rows

    def _validate(self, units):
        if not units: return {"ok":False,"msg":"No units extracted"}
        eff = [u["effective_rent"] for u in units if u.get("effective_rent")]
        if not eff: return {"ok":False,"msg":"All units have zero effective rent"}
        avg = sum(eff)/len(eff)
        if avg < 50:  return {"ok":False,"msg":f"Avg eff rent ${avg:.0f} too low"}
        if avg > 100000: return {"ok":False,"msg":f"Avg eff rent ${avg:.0f} too high"}
        return {"ok":True,"msg":f"{len(units)} units · {len(eff)} occupied · avg ${avg:,.0f}"}

    def _write_template(self, units, rent_roll_date):
        # keep_links=False strips the external link cache that causes Excel repair errors
        wb = load_workbook(TEMPLATE_PATH, keep_links=False)
        ws = wb["Rent Roll"]

        # ONLY write to the 9 permitted columns:
        # C=3  Unit No        D=4  Unit Type       H=8  Size
        # I=9  Move-in date   J=10 Lease start     K=11 Lease end
        # N=14 Resident Name  O=15 Market Rent      Q=17 Effective Rent
        # All other columns (formulas, summaries, etc.) are left completely untouched.

        # Set Rent Roll date in G5 (col 7) — the one non-data cell we must fill
        try: ws.cell(5, 7).value = datetime.strptime(rent_roll_date, "%Y-%m-%d")
        except: pass

        ALLOWED_COLS = {3, 4, 8, 9, 10, 11, 14, 15, 17}  # C D H I J K N O Q

        for idx, u in enumerate(units):
            r = 9 + idx
            if r > 620: break

            # C — Unit No
            ws.cell(r, 3).value = u.get("unit_no")

            # D — Unit Type
            ws.cell(r, 4).value = u.get("unit_type") or ""

            # H — Size (sqft)
            sqft = u.get("sqft")
            if sqft:
                ws.cell(r, 8).value = int(sqft)

            # I J K — Move-in, Lease Start, Lease End (stored as dates so Excel formats them)
            for col, field in [(9,"move_in"), (10,"lease_start"), (11,"lease_end")]:
                v = u.get(field)
                if v:
                    try: ws.cell(r, col).value = datetime.strptime(v, "%m-%d-%Y")
                    except: pass

            # N — Resident Name
            # For vacant/model/down units: show the status label, not blank
            name = u.get("resident_name") or ""
            status = u.get("status") or ""
            if not name:
                # Derive a label from status (e.g. "Vacant", "Model", "Down")
                sl = status.lower()
                if "model" in sl:
                    name = "Model"
                elif "down" in sl or "offline" in sl:
                    name = "Down"
                elif not u.get("effective_rent"):
                    name = "Vacant"
            ws.cell(r, 14).value = name

            # O — Market Rent (no decimals)
            mkt = u.get("market_rent")
            if mkt:
                ws.cell(r, 15).value = round(mkt)

            # Q — Effective Rent (blank if zero/none, no decimals)
            eff = u.get("effective_rent")
            if eff and eff > 0:
                ws.cell(r, 17).value = round(eff)
            # else leave blank — formula cells below row 9 remain untouched

        out = io.BytesIO()
        wb.save(out)
        out.seek(0)
        return out.read()

    def run(self, raw_bytes, filename, rent_roll_date):
        self.log("step", f"🚀 Agent started — {filename}")
        try:
            raw_wb = load_workbook(io.BytesIO(raw_bytes), data_only=True)
        except Exception as e:
            return {"ok":False,"error":f"Cannot open file: {e}"}

        sheet = self._find_raw_sheet(raw_wb)
        ws    = raw_wb[sheet]
        self.log("action", f"📄 Sheet: '{sheet}'  ({ws.max_row}r × {ws.max_column}c)")
        rows  = self._to_rows(ws)

        units       = []
        parser_name = None

        # Try known parsers
        for p in self.PARSERS:
            if p.can_handle(rows, filename):
                self.log("think", f"🔍 Matched parser: {p.name}")
                try:
                    units = p.extract(rows)
                    parser_name = p.name
                    self.log("action", f"   → {len(units)} units extracted")
                    break
                except Exception as e:
                    self.log("warn", f"   Parser failed: {e} — trying next")

        # AI fallback
        ai = AIFallbackParser()
        if not units:
            self.log("think", "🧠 No known format matched — invoking Claude for schema detection")
            schema = ai.detect_schema(rows, filename)
            self.log("action", f"   Format: {schema.get('format')} | {schema.get('structure')}")
            units       = ai.extract(rows, schema)
            parser_name = f"AI ({schema.get('format','?')})"
            self.log("action", f"   → {len(units)} units extracted")

        # Validate
        v = self._validate(units)
        self.log("ok" if v["ok"] else "warn",
                 ("✅ " if v["ok"] else "⚠️ ") + f"Validation: {v['msg']}")

        # Self-correction: if validation failed, try AI override
        if not v["ok"]:
            self.log("think", "🔄 Attempting AI self-correction...")
            try:
                schema   = ai.detect_schema(rows, filename)
                units2   = ai.extract(rows, schema)
                v2       = self._validate(units2)
                if v2["ok"] or len(units2) > len(units):
                    units, v = units2, v2
                    parser_name += " [corrected]"
                    self.log("ok", f"✅ Correction: {v['msg']}")
                else:
                    self.log("warn", "Correction did not improve — keeping original")
            except Exception as e:
                self.log("warn", f"Correction error: {e}")

        # Write template
        try:
            out_bytes = self._write_template(units, rent_roll_date)
            self.log("ok", "📥 Template populated successfully")
        except Exception as e:
            return {"ok":False,"error":f"Template write failed: {e}"}

        occ   = sum(1 for u in units if u.get("effective_rent"))
        t_eff = sum(u.get("effective_rent",0) or 0 for u in units)
        t_mkt = sum(u.get("market_rent",0) or 0 for u in units)

        return {
            "ok": True,
            "output_bytes": out_bytes,
            "metrics": {
                "total_units": len(units),
                "occupied": occ,
                "vacant": len(units) - occ,
                "total_effective_rent": t_eff,
                "total_market_rent": t_mkt,
                "parser": parser_name,
            },
            "validation": v,
            "preview": units[:5],
        }


# ─── Streamlit UI ─────────────────────────────────────────────────────────────
def main():
    st.title("🏢 Rent Roll AI Agent")
    st.caption("Auto-detects any rent roll format · extracts & transforms data · populates your standard template")
    st.divider()

    with st.sidebar:
        st.header("⚙️ Settings")
        rr_date = st.date_input("Rent Roll As-Of Date", value=datetime.today()).strftime("%Y-%m-%d")
        st.divider()
        st.markdown("**Built-in parsers**")
        for n in ["Yardi Voyager/Breeze","RealPage OneSite","MRI Living","AppFolio","Rent Manager"]:
            st.markdown(f"• {n}")
        st.markdown("• *Any other via Claude AI*")
        st.divider()
        st.markdown("**Agent logic**")
        st.markdown("""
- 🔍 Format auto-detection
- 💸 Rent code identification
- 🏗️ Building/section grouping
- 🚫 Skips applicants & future
- 🏚️ Zeros vacant/model/down
- ✅ Self-validates results
- 🔄 Claude correction loop
        """)

    uploaded = st.file_uploader("📂 Upload Raw Rent Roll (.xlsx)", type=["xlsx"])
    if not uploaded:
        st.info("Upload a raw rent roll file to begin.")
        return

    st.success(f"**{uploaded.name}** — {uploaded.size:,} bytes")
    _, mid, _ = st.columns([1,2,1])
    with mid:
        run_btn = st.button("🚀 Run Agent", type="primary", use_container_width=True)
    if not run_btn: return

    log_box = st.empty()
    log_lines = []

    def add_log(level, msg):
        cls  = {"step":"log-step","think":"log-think","action":"log-action",
                "warn":"log-warn","ok":"log-ok"}.get(level,"")
        icon = {"step":"▶","think":"💭","action":"⚡","warn":"⚠","ok":"✓"}.get(level,"·")
        ts   = datetime.now().strftime("%H:%M:%S")
        log_lines.append(f'<span class="{cls}">[{ts}] {icon} {msg}</span>')
        log_box.markdown(
            f'<div class="log-box">{"<br>".join(log_lines)}</div>',
            unsafe_allow_html=True,
        )

    raw_bytes = uploaded.read()
    with st.spinner(""):
        result = RentRollAgent(log_fn=add_log).run(raw_bytes, uploaded.name, rr_date)

    st.divider()
    if not result["ok"]:
        st.error(f"❌ {result.get('error','Unknown error')}")
        return

    m = result["metrics"]
    v = result["validation"]

    st.subheader("📊 Results")
    c1,c2,c3,c4,c5 = st.columns(5)
    c1.metric("Total Units",   m["total_units"])
    c2.metric("Occupied",      m["occupied"])
    c3.metric("Vacant/Other",  m["vacant"])
    c4.metric("Total Eff Rent",f"${m['total_effective_rent']:,.0f}")
    c5.metric("Total Mkt Rent",f"${m['total_market_rent']:,.0f}")

    vi = "✅" if v["ok"] else "⚠️"
    st.info(f"{vi} **Validation:** {v['msg']}")
    st.caption(f"Parser: **{m['parser']}**")

    if result.get("preview"):
        with st.expander("🔎 Preview first 5 units"):
            df = pd.DataFrame(result["preview"])
            show = [c for c in ["unit_no","unit_type","sqft","resident_name",
                                 "move_in","lease_start","lease_end",
                                 "market_rent","effective_rent"] if c in df.columns]
            st.dataframe(df[show], use_container_width=True)

    st.divider()
    out_name = uploaded.name.replace(".xlsx","") + "_Rent_Roll.xlsx"
    st.download_button(
        "📥 Download Filled Rent Roll Template",
        data=result["output_bytes"],
        file_name=out_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )


if __name__ == "__main__":
    main()
