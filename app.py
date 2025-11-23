import streamlit as st
import pandas as pd
from io import BytesIO
from dateutil import parser as dtparser
from google.oauth2.service_account import Credentials
import gspread
from gspread.utils import rowcol_to_a1
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A3
from reportlab.lib.units import mm

# ===== Config =====
st.set_page_config(page_title="Permission Cell — North-West", layout="wide")

SHEET_ID = st.secrets["sheet"]["id"]
SHEET_NAME = st.secrets["sheet"]["name"]

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# Expected headers (case/space/underscore-insensitive)
NEED = [
    "refno","appno","dated","acname","acno","district","organizername","organizermobile",
    "party","designation","typeprog","venueprog","psvenue","date","time","route","gathering",
    "localpolice","traffic","landown","fire","permission","reason","orderno","wardno","orderdate"
]

def _norm(s: str) -> str:
    return str(s or "").strip().lower().replace(" ", "").replace("_", "")

@st.cache_resource(show_spinner=False)
def _open_ws():
    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"], scopes=SCOPES
    )
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(SHEET_ID)
    return sh.worksheet(SHEET_NAME)

def _fetch_table():
    """Return (values, head_map, col_count). values includes header row 0."""
    ws = _open_ws()
    values = ws.get_all_values()
    if not values:
        raise RuntimeError("Empty sheet.")
    heads = values[0]
    head_map = { _norm(h): i for i, h in enumerate(heads) }
    # Validate headers
    missing = [k for k in NEED if k not in head_map]
    if missing:
        raise RuntimeError(f"Missing headers: {', '.join(missing)}")
    return values, head_map, len(heads)

@st.cache_data(show_spinner=False, ttl=20)
def load_dataframe():
    values, H, col_count = _fetch_table()
    rows = []
    for r in range(1, len(values)):
        row = values[r]
        if not any(row):
            continue
        d = {k: row[H[k]] if H[k] < len(row) else "" for k in H}
        d["_row"] = r + 1
        rows.append(d)
    df = pd.DataFrame(rows)
    return df

def check_unique(df: pd.DataFrame, refno: str, appno: str, exclude_row: int|None):
    ref_unique = True
    app_unique = True
    if refno:
        ref_unique = not any((df["refno"].astype(str) == str(refno)) & (df["_row"] != exclude_row))
    if appno:
        app_unique = not any((df["appno"].astype(str) == str(appno)) & (df["_row"] != exclude_row))
    return ref_unique, app_unique

def _max_numeric(text_series: pd.Series) -> int:
    best = 0
    for v in text_series.dropna().astype(str):
        n = "".join(ch for ch in v if ch.isdigit())
        if n.isdigit():
            best = max(best, int(n))
    return best

def generate_unique_ids(df: pd.DataFrame, acno_raw: str|None):
    app_next = _max_numeric(df.get("appno", pd.Series([], dtype=str))) + 1
    appno = str(app_next)

    ac = "".join(ch for ch in str(acno_raw or "00") if ch.isdigit())
    prefix = (ac.zfill(2) if ac else "00") + "AC"
    # like GAS: start suffix at 39999 and grow
    suffix_max = 39999
    for ref in df.get("refno", pd.Series([], dtype=str)).astype(str):
        if ref.startswith(prefix):
            suf = ref[len(prefix):]
            if suf.isdigit():
                suffix_max = max(suffix_max, int(suf))
    # find next unused
    ws = _open_ws()
    values = ws.get_all_values()
    H = { _norm(h): i for i,h in enumerate(values[0]) }
    taken = { str(values[r][H["refno"]]) for r in range(1, len(values)) if H["refno"] < len(values[r]) }
    tries = 0
    while tries < 50:
        suffix_max += 1
        refno = f"{prefix}{str(suffix_max).zfill(5)}"
        if refno not in taken:
            return refno, appno
        tries += 1
    raise RuntimeError("Could not generate unique IDs after many tries.")

def to_row_payload(head_map: dict, col_count: int, payload: dict):
    out = [""] * col_count
    for k, v in payload.items():
        nk = _norm(k)
        if nk in head_map:
            out[head_map[nk]] = v
    return out

def update_row(row_index: int, payload: dict):
    ws = _open_ws()
    values, H, col_count = _fetch_table()
    out = to_row_payload(H, col_count, payload)
    rng = f"{rowcol_to_a1(row_index,1)}:{rowcol_to_a1(row_index,col_count)}"
    ws.update(rng, [out], value_input_option="USER_ENTERED")

def add_row(payload: dict) -> int:
    ws = _open_ws()
    values, H, col_count = _fetch_table()
    out = to_row_payload(H, col_count, payload)
    ws.append_row(out, value_input_option="USER_ENTERED")
    # Re-fetch to get the new row index (safe & simple)
    values2 = ws.get_all_values()
    return len(values2)  # last row index

def search_by_ref(ref: str) -> dict|None:
    df = load_dataframe()
    needle = _norm(ref)
    for _, row in df.iterrows():
        if _norm(row["refno"]) == needle:
            return row.to_dict()
    return None

def format_date_fallback(s: str, placeholder="______/_______/2025"):
    s = (s or "").strip()
    if not s:
        return placeholder
    try:
        d = dtparser.parse(s, dayfirst=True, fuzzy=True)
        return d.strftime("%d/%m/%Y")
    except Exception:
        return s

def pack_view(row: dict) -> dict:
    """Return dict for the order view."""
    return {
        "refno": row.get("refno",""),
        "appno": row.get("appno",""),
        "dated": format_date_fallback(row.get("dated","")),
        "acname": row.get("acname",""),
        "acno": row.get("acno",""),
        "wardno": row.get("wardno",""),
        "district": row.get("district",""),
        "organizername": row.get("organizername",""),
        "organizermobile": row.get("organizermobile",""),
        "party": row.get("party",""),
        "designation": row.get("designation",""),
        "typeprog": row.get("typeprog",""),
        "venueprog": row.get("venueprog",""),
        "psvenue": row.get("psvenue",""),
        "date": format_date_fallback(row.get("date","")),
        "time": row.get("time",""),
        "route": row.get("route",""),
        "gathering": row.get("gathering",""),
        "localpolice": row.get("localpolice",""),
        "traffic": row.get("traffic",""),
        "landown": row.get("landown",""),
        "fire": row.get("fire",""),
        "permission": row.get("permission",""),
        "reason": row.get("reason",""),
        "orderno": row.get("orderno",""),
        "orderdate": format_date_fallback(row.get("orderdate","")),
    }

def pdf_bytes_from_view(view: dict) -> bytes:
    """Very compact A3 PDF (portrait) with key fields."""
    buf = BytesIO()
    c = canvas.Canvas(buf, pagesize=A3)  # (841.89 x 1190.55) points portrait
    W, H = A3
    x = 20*mm
    y = H - 20*mm
    line = 7*mm

    c.setFont("Helvetica-Bold", 16)
    c.drawCentredString(W/2, y, "PERMISSION CELL / SINGLE WINDOW — NORTH-WEST (KANJHAWALA)")
    y -= 10*mm

    c.setFont("Helvetica-Bold", 12)
    c.drawString(x, y, f"Ref No.: {view['refno']}"); y -= line
    c.drawString(x, y, f"App No.: {view['appno']}"); y -= line
    c.drawString(x, y, f"Dated  : {view['dated']}"); y -= (line+2*mm)

    c.setFont("Helvetica-Bold", 13)
    c.drawCentredString(W/2, y, "ORDER")
    y -= (line + 4*mm)

    c.setFont("Helvetica", 11)
    def rowlbl(num, label, value):
        nonlocal y
        c.setFont("Helvetica-Bold", 11)
        c.drawString(x, y, f"{num}. {label}")
        c.setFont("Helvetica", 11)
        c.drawString(x+190, y, f": {value}")
        y -= line

    rowlbl(1, "Ward & No. (AC/Ward)", f"{view['acname']}  (AC-{view['acno']}) (Ward-{view['wardno']})")
    rowlbl(2, "Election District", view["district"])
    rowlbl(3, "Organizer & Contact", f"{view['organizername']} ({view['organizermobile']})")
    rowlbl(4, "Party & Designation", f"{view['party']}, {view['designation']}")
    rowlbl(5, "Type of Programme", view["typeprog"])
    rowlbl(6, "Venue (PS)", f"{view['venueprog']} ({view['psvenue']})")
    rowlbl(7, "Date", view["date"])
    rowlbl(8, "Time", view["time"])
    rowlbl(9, "Route/Distance", view["route"])
    rowlbl(10, "Permitted gathering", view["gathering"])

    c.setFont("Helvetica-Bold", 11)
    c.drawString(x, y, "11. NOC obtained from"); y -= line
    c.setFont("Helvetica", 11)
    c.drawString(x+18, y, f"Local Police : {view['localpolice']}"); y -= line
    c.drawString(x+18, y, f"Traffic      : {view['traffic']}"); y -= line
    c.drawString(x+18, y, f"Land owning  : {view['landown']}"); y -= line
    c.drawString(x+18, y, f"Fire Deptt   : {view['fire']}"); y -= (line + 2*mm)

    c.setFont("Helvetica-Bold", 11)
    c.drawString(x, y, "12. Permission / Reason"); y -= line
    c.setFont("Helvetica", 11)
    c.drawString(x+18, y, f"Permission : {view['permission']}"); y -= line
    c.drawString(x+18, y, f"Reason     : {view['reason']}"); y -= (line + 2*mm)

    c.setFont("Helvetica", 11)
    c.drawString(x, 25*mm, f"No. {view['appno']} /ACP(P)RO/PC-(NORTH-WEST)")
    c.drawRightString(W - 20*mm, 25*mm, f"Dated : {view['dated']}")

    c.showPage()
    c.save()
    buf.seek(0)
    return buf.read()

# ===== UI =====
if "offset" not in st.session_state: st.session_state.offset = 0
if "selected" not in st.session_state: st.session_state.selected = None
if "filter" not in st.session_state: st.session_state.filter = ""
PAGE = 60

st.title("Permission Cell / Single Window — North-West")

# Top search & actions
top_left, top_mid, top_right = st.columns([0.42, 0.28, 0.30])
with top_left:
    ref_query = st.text_input("Search by Reference No.", placeholder="e.g. 28AC44838")
with top_mid:
    if st.button("Search"):
        with st.spinner("Searching…"):
            hit = search_by_ref(ref_query)
        if hit:
            st.session_state.selected = hit
            st.success("Loaded.")
        else:
            st.error("No record found.")
with top_right:
    new_click = st.button("New Entry", type="primary")

df = load_dataframe()

# ===== Left: List (descending appno) with Filter + Load more
left, right = st.columns([0.36, 0.64], gap="small")

with left:
    st.subheader("Applications")
    st.session_state.filter = st.text_input("Filter (app/ref/organizer/party/type)", value=st.session_state.filter)

    tmp = df.copy()
    q = st.session_state.filter.strip().lower()
    if q:
        mask = (
            tmp["appno"].astype(str).str.lower().str.contains(q) |
            tmp["refno"].astype(str).str.lower().str.contains(q) |
            tmp["organizername"].astype(str).str.lower().str.contains(q) |
            tmp["party"].astype(str).str.lower().str.contains(q) |
            tmp["typeprog"].astype(str).str.lower().str.contains(q)
        )
        tmp = tmp[mask]

    def _num(v):
        s = "".join(ch for ch in str(v) if ch.isdigit())
        return int(s) if s.isdigit() else -10**9

    tmp = tmp.sort_values(by="appno", key=lambda s: s.map(_num), ascending=False).reset_index(drop=True)

    page_df = tmp.iloc[0: st.session_state.offset + PAGE]
    for _, r in page_df.iterrows():
        label = f"**{r['appno']}**  ·  {r.get('organizername','')[:24]}{'…' if len(str(r.get('organizername',''))) > 24 else ''}"
        sub = f"{r.get('party','')}  ·  {r.get('typeprog','')}  ·  {r.get('refno','')}"
        if st.button(label, key=f"pick_{r['appno']}"):
            st.session_state.selected = r.to_dict()
            st.toast(f"Loaded {r['appno']}")
        st.caption(sub)

    if (st.session_state.offset + PAGE) < len(tmp):
        if st.button("Load more"):
            st.session_state.offset += PAGE

    if st.button("Refresh list"):
        st.cache_data.clear()
        st.session_state.offset = 0
        st.rerun()

# ===== Right: View + Form
with right:
    if new_click:
        st.session_state.selected = None  # clear -> new entry mode

    selected = st.session_state.selected

    st.subheader("Order Preview")
    if selected:
        view = pack_view(selected)
        # Compact visual in Markdown (styled rows)
        st.markdown(f"""
**Ref No.**: {view['refno']} &nbsp;&nbsp; **App No.**: {view['appno']} &nbsp;&nbsp; **Dated**: {view['dated']}

**1. Ward & No. (AC/Ward):** {view['acname']}  (AC-{view['acno']}) (Ward-{view['wardno']})  
**2. District:** {view['district']}  
**3. Organizer & Contact:** {view['organizername']} ({view['organizermobile']})  
**4. Party & Designation:** {view['party']}, {view['designation']}  
**5. Type of Programme:** {view['typeprog']}  
**6. Venue (PS):** {view['venueprog']} ({view['psvenue']})  
**7. Date:** {view['date']}  
**8. Time:** {view['time']}  
**9. Route:** {view['route']}  
**10. Permitted gathering:** {view['gathering']}

**11. NOC obtained from**  
• Local Police: {view['localpolice']}  
• Traffic: {view['traffic']}  
• Land owning: {view['landown']}  
• Fire: {view['fire']}

**12. Permission / Reason**  
• Permission: {view['permission']}  
• Reason: {view['reason']}

_No. {view['appno']} /ACP(P)RO/PC-(NORTH-WEST) &nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Dated: {view['dated']}_
        """)
        pdf_data = pdf_bytes_from_view(view)
        st.download_button("Download PDF", data=pdf_data, file_name=f"Order_{view['appno']}.pdf", mime="application/pdf")

    # ===== Form (Edit / New)
    st.divider()
    st.subheader("Edit / Add")
    with st.form(key="edit_form", clear_on_submit=False):
        # When selected is None -> New Entry mode
        col1, col2, col3 = st.columns(3)
        with col1:
            refno = st.text_input("Ref No.", value=(selected or {}).get("refno",""))
            acname = st.text_input("Ward / Area Name", value=(selected or {}).get("acname",""))
            organizername = st.text_input("Organizer", value=(selected or {}).get("organizername",""))
            party = st.text_input("Party", value=(selected or {}).get("party",""))
            typeprog = st.text_input("Type of Programme", value=(selected or {}).get("typeprog",""))
            venueprog = st.text_input("Venue", value=(selected or {}).get("venueprog",""))
            localpolice = st.text_input("Local Police", value=(selected or {}).get("localpolice",""))
            permission = st.text_input("Permission", value=(selected or {}).get("permission",""))

        with col2:
            appno = st.text_input("Application No.", value=(selected or {}).get("appno",""))
            acno = st.text_input("AC No.", value=(selected or {}).get("acno",""))
            organizermobile = st.text_input("Organizer Mobile", value=(selected or {}).get("organizermobile",""))
            designation = st.text_input("Designation", value=(selected or {}).get("designation",""))
            psvenue = st.text_input("Police Station", value=(selected or {}).get("psvenue",""))
            date_str = st.text_input("Date (DD-MM-YYYY)", value=(selected or {}).get("date",""))
            traffic = st.text_input("Traffic", value=(selected or {}).get("traffic",""))
            reason = st.text_area("Reason (if any)", value=(selected or {}).get("reason",""))

        with col3:
            dated = st.text_input("Dated (DD-MM-YYYY)", value=(selected or {}).get("dated",""))
            wardno = st.text_input("Ward No.", value=(selected or {}).get("wardno",""))
            district = st.text_input("District", value=(selected or {}).get("district",""))
            time_str = st.text_input("Time (e.g., 02:00 PM TO 05:00 PM)", value=(selected or {}).get("time",""))
            route = st.text_input("Route / Distance", value=(selected or {}).get("route",""))
            gathering = st.text_input("Permitted Gathering", value=(selected or {}).get("gathering",""))
            landown = st.text_input("Land Owning", value=(selected or {}).get("landown",""))
            fire = st.text_input("Fire", value=(selected or {}).get("fire",""))
            orderno = st.text_input("Order No. (optional)", value=(selected or {}).get("orderno",""))
            orderdate = st.text_input("Order Date (optional, DD-MM-YYYY)", value=(selected or {}).get("orderdate",""))

        submitted_update = st.form_submit_button("Update existing", use_container_width=True)
        submitted_add = st.form_submit_button("Add as new", use_container_width=True, type="primary")

    if submitted_update or submitted_add:
        row_idx = (selected or {}).get("_row") if submitted_update else None
        df = load_dataframe()
        ref_unique, app_unique = check_unique(df, refno.strip(), appno.strip(), row_idx)

        # handle auto-generate for "Add new"
        gen_ref, gen_app = None, None
        if submitted_add and (not refno.strip() or not appno.strip()):
            try:
                gen_ref, gen_app = generate_unique_ids(df, acno)
            except Exception as e:
                st.error(f"Auto-generate failed: {e}")
                st.stop()

        if not submitted_add:
            if not refno.strip() or not appno.strip():
                st.error("Ref No. and Application No. are required for update.")
                st.stop()

        if not ref_unique:
            st.error("Duplicate Reference No. — must be unique.")
            st.stop()
        if not app_unique:
            st.error("Duplicate Application No. — must be unique.")
            st.stop()

        payload = {
            "refno": gen_ref or refno.strip(),
            "appno": gen_app or appno.strip(),
            "dated": dated.strip(),
            "acname": acname.strip(),
            "acno": acno.strip(),
            "district": district.strip(),
            "organizername": organizername.strip(),
            "organizermobile": organizermobile.strip(),
            "party": party.strip(),
            "designation": designation.strip(),
            "typeprog": typeprog.strip(),
            "venueprog": venueprog.strip(),
            "psvenue": psvenue.strip(),
            "date": date_str.strip(),
            "time": time_str.strip(),
            "route": route.strip(),
            "gathering": gathering.strip(),
            "localpolice": localpolice.strip(),
            "traffic": traffic.strip(),
            "landown": landown.strip(),
            "fire": fire.strip(),
            "permission": permission.strip(),
            "reason": reason.strip(),
            "orderno": orderno.strip(),
            "wardno": wardno.strip(),
            "orderdate": orderdate.strip(),
        }

        try:
            if submitted_update and row_idx:
                with st.spinner("Updating record…"):
                    update_row(int(row_idx), payload)
                st.success("Updated.")
                st.cache_data.clear()
                df2 = load_dataframe()
                match = df2.loc[df2["refno"] == payload["refno"]]
                st.session_state.selected = match.iloc[0].to_dict() if not match.empty else None
            elif submitted_add:
                with st.spinner("Adding new entry…"):
                    new_row = add_row(payload)
                st.success(f"Added as new (row {new_row}).")
                st.cache_data.clear()
                df2 = load_dataframe()
                match = df2.loc[df2["refno"] == payload["refno"]]
                st.session_state.selected = match.iloc[0].to_dict() if not match.empty else None
                st.session_state.offset = 0  # so newest shows on top
        except Exception as e:
            st.error(f"Operation failed: {e}")
