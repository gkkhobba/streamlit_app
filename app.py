import streamlit as st
import pandas as pd
from dateutil import parser as dtparser
from google.oauth2.service_account import Credentials
import gspread
from jinja2 import Template
from io import BytesIO

# --- Optional import: WeasyPrint (HTML->PDF). If missing, we fall back to ReportLab.
HAS_WEASYPRINT = False
try:
    from weasyprint import HTML  # type: ignore
    HAS_WEASYPRINT = True
except Exception:
    HAS_WEASYPRINT = False

# Fallback PDF generator (compact A3) when WeasyPrint is unavailable
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A3
from reportlab.lib.units import mm

# ========= Streamlit setup =========
st.set_page_config(page_title="Permission Cell — North-West", layout="wide")

# ========= Config =========
SHEET_ID = st.secrets["sheet"]["id"]
SHEET_NAME = st.secrets["sheet"]["name"]
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

NEED = [
    "refno","appno","dated","acname","acno","district","organizername","organizermobile",
    "party","designation","typeprog","venueprog","psvenue","date","time","route","gathering",
    "localpolice","traffic","landown","fire","permission","reason","orderno","wardno","orderdate"
]

def _norm(s: str) -> str:
    return str(s or "").strip().lower().replace(" ", "").replace("_", "")

@st.cache_resource(show_spinner=False)
def _ws():
    creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=SCOPES)
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(SHEET_ID)
    return sh.worksheet(SHEET_NAME)

def _fetch_table():
    ws = _ws()
    values = ws.get_all_values()
    if not values:
        raise RuntimeError("Empty sheet.")
    heads = values[0]
    H = { _norm(h): i for i,h in enumerate(heads) }
    missing = [k for k in NEED if k not in H]
    if missing:
        raise RuntimeError(f"Missing headers: {', '.join(missing)}")
    return ws, values, H, len(heads)

@st.cache_data(show_spinner=False, ttl=20)
def load_df():
    ws, values, H, _ = _fetch_table()
    rows = []
    for r in range(1, len(values)):
        row = values[r]
        if not any(row):  # skip entirely blank rows
            continue
        d = {k: (row[H[k]] if H[k] < len(row) else "") for k in H}
        d["_row"] = r+1
        rows.append(d)
    return pd.DataFrame(rows)

def check_unique(df: pd.DataFrame, refno: str, appno: str, exclude_row: int|None):
    ref_unique = True
    app_unique = True
    if refno:
        ref_unique = not any((df["refno"].astype(str)==str(refno)) & (df["_row"] != exclude_row))
    if appno:
        app_unique = not any((df["appno"].astype(str)==str(appno)) & (df["_row"] != exclude_row))
    return ref_unique, app_unique

def _max_numeric(series: pd.Series) -> int:
    best = 0
    for v in series.dropna().astype(str):
        digits = "".join(ch for ch in v if ch.isdigit())
        if digits.isdigit():
            best = max(best, int(digits))
    return best

def generate_ids(df: pd.DataFrame, acno_raw: str|None):
    # Application number is just next highest numeric
    app_next = _max_numeric(df.get("appno", pd.Series([], dtype=str))) + 1
    appno = str(app_next)

    # Reference number  : <2-digit AC>AC<5-digit suffix starting after 39999>
    ac = "".join(ch for ch in str(acno_raw or "00") if ch.isdigit())
    prefix = (ac.zfill(2) if ac else "00") + "AC"
    suffix = 39999
    for ref in df.get("refno", pd.Series([], dtype=str)).astype(str):
        if ref.startswith(prefix):
            tail = ref[len(prefix):]
            if tail.isdigit(): suffix = max(suffix, int(tail))
    # Re-scan live sheet to avoid rare races
    ws, values, H, _ = _fetch_table()
    taken = {values[r][H["refno"]] for r in range(1, len(values)) if H["refno"] < len(values[r])}
    tries = 0
    while tries < 50:
        suffix += 1
        refno = f"{prefix}{str(suffix).zfill(5)}"
        if refno not in taken:
            return refno, appno
        tries += 1
    raise RuntimeError("ID generation failed after many tries.")

def to_row(H: dict, width: int, payload: dict):
    out = [""] * width
    for k, v in payload.items():
        nk = _norm(k)
        if nk in H:
            out[H[nk]] = v
    return out

def update_row(row_index: int, payload: dict):
    ws, _, H, width = _fetch_table()
    rng = gspread.utils.rowcol_to_a1(row_index,1) + ":" + gspread.utils.rowcol_to_a1(row_index, width)
    ws.update(rng, [to_row(H, width, payload)], value_input_option="USER_ENTERED")

def add_row(payload: dict) -> int:
    ws, _, H, width = _fetch_table()
    ws.append_row(to_row(H, width, payload), value_input_option="USER_ENTERED")
    return len(ws.get_all_values())

def search_by_ref(ref: str):
    df = load_df()
    needle = _norm(ref)
    for _, row in df.iterrows():
        if _norm(row["refno"]) == needle:
            return row.to_dict()
    return None

def fmt_date(s: str, placeholder="______/_______/2025"):
    s = (s or "").strip()
    if not s: return placeholder
    try:
        d = dtparser.parse(s, dayfirst=True, fuzzy=True)
        return d.strftime("%d/%m/%Y")
    except Exception:
        return s

def pack_view(row: dict) -> dict:
    return {
        "refno": row.get("refno",""),
        "appno": row.get("appno",""),
        "dated": fmt_date(row.get("dated","")),
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
        "date": fmt_date(row.get("date","")),
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
        "orderdate": fmt_date(row.get("orderdate","")),
    }

# ======= A3 HTML (for on-screen preview & WeasyPrint when available) =======
HTML_TMPL = Template(r"""
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
@page { size: A3; margin: 10mm 12mm; }
*{ box-sizing:border-box; }
body{ font: 14pt/1.28 "Inter", system-ui, -apple-system, "Segoe UI", Roboto, Arial, "Noto Sans", sans-serif; color:#0f172a; }
.sheet{ position:relative; border:1px solid #d1d5db; border-radius:8px; padding:10mm 12mm; }
.wm{ position:absolute; inset:0; margin:auto; width:42%; opacity:.07; filter:grayscale(100%); z-index:0; }
.topband{ display:grid; grid-template-columns:90px 1fr 90px; gap:8px; align-items:center;
  border:2px solid #111; border-radius:8px; padding:8px 10px; background:#fff; position:relative; z-index:1; }
.logo{ width:90px; height:90px; object-fit:contain; }
.t1{ font-weight:900; font-size:22pt; text-transform:uppercase; text-align:center; }
.t2{ font-weight:800; font-size:16pt; text-transform:uppercase; text-align:center; }
.t3{ font-weight:800; font-size:14pt; text-transform:uppercase; text-align:center; }
.infostrip{ display:grid; grid-template-columns:1fr 1fr 1fr; gap:6px; margin-top:8px; z-index:1; position:relative; }
.infostrip > div{ border:1.4px solid #111; border-radius:6px; padding:6px 8px; font-weight:800; background:#fff; }
.order-title{ text-align:center; font-weight:900; font-size:16pt; margin:8mm 0 5mm; text-transform:uppercase; }
table{ width:100%; border-collapse:collapse; }
th,td{ border:1px solid #111; padding:6px 8px; vertical-align:middle; }
.idx{ width:6%; text-align:center; font-weight:900; }
.lab{ width:47%; font-weight:800; }
.val{ width:47%; font-weight:600; word-break:break-word; white-space:pre-wrap; }
.grid2{ display:grid; grid-template-columns:1fr 1fr; gap:0px 18px; }
.muted{ color:#6b7280; font-weight:600; }
.signs{ display:grid; row-gap:18px; margin-top:12mm; }
.sigL{ justify-self:start; }
.sigR{ justify-self:end; }
.sigbox{ display:inline-block; border-top:1px solid #000; padding-top:4px; text-transform:uppercase; font-weight:700; }
.meta{ display:flex; justify-content:space-between; margin-top:6mm; font-weight:800; }
.tc{ margin-top:7mm; }
.tc .ttl{ font-weight:900; margin-bottom:4px; text-transform:uppercase; }
.tc ol{ margin:0; padding-left:18px; }
.tc li{ margin:2px 0; line-height:1.22; }
.small{ font-size:12pt; }
</style>
</head>
<body>
  <div class="sheet">
    <img class="wm" src="https://upload.wikimedia.org/wikipedia/commons/5/55/Emblem_of_India.svg" alt="">
    <div class="topband">
      <img class="logo" src="https://upload.wikimedia.org/wikipedia/commons/5/55/Emblem_of_India.svg" alt="">
      <div>
        <div class="t1">OFFICE OF THE INCHARGE</div>
        <div class="t2">PERMISSION CELL / SINGLE WINDOW</div>
        <div class="t2">DISTRICT ELECTION OFFICER : NORTH-WEST</div>
        <div class="t3">KANJHAWALA DELHI - 110081</div>
      </div>
      <img class="logo" src="https://upload.wikimedia.org/wikipedia/commons/3/32/Swachh_Bharat_Mission_Logo.svg" alt="">
    </div>

    <div class="infostrip">
      <div>Ref No. <b>{{ view.refno or "________" }}</b></div>
      <div>Application No. <b>{{ view.appno or "——" }}</b></div>
      <div>Dated : <b>{{ view.dated or "______/_______/2025" }}</b></div>
    </div>

    <div class="order-title">ORDER</div>

    <table>
      <tr><th class="idx">1.</th><th class="lab">Name of Municipal Corporation Ward &amp; No.</th>
        <td class="val"><span>{{ view.acname }}</span> <span class="muted">(AC-{{ view.acno }})</span><span class="muted"> (Ward-{{ view.wardno }})</span></td></tr>
      <tr><th class="idx">2.</th><th class="lab">Name of the Election District</th>
        <td class="val">{{ view.district }}</td></tr>
      <tr><th class="idx">3.</th><th class="lab">Name of the organizer &amp; Contact No</th>
        <td class="val"><span>{{ view.organizername }}</span> ( <span>{{ view.organizermobile }}</span> )</td></tr>
      <tr><th class="idx">4.</th><th class="lab">Party affiliation and his designation</th>
        <td class="val"><span>{{ view.party }}</span>, <span>{{ view.designation }}</span></td></tr>
      <tr><th class="idx">5.</th><th class="lab">Type of programme (meeting procession, rally, nukkad natak, pad yatra etc. with loudspeaker or without it)</th>
        <td class="val">{{ view.typeprog }}</td></tr>
      <tr><th class="idx">6.</th><th class="lab">Name of venue with police Station</th>
        <td class="val"><span>{{ view.venueprog }}</span> ( <span>{{ view.psvenue }}</span> )</td></tr>
      <tr><th class="idx">7.</th><th class="lab">Date</th>
        <td class="val">{{ view.date or "______/_______/2025" }}</td></tr>
      <tr><th class="idx">8.</th><th class="lab">Timing of Programme (Start and ending)</th>
        <td class="val">{{ view.time }}</td></tr>
      <tr><th class="idx">9.</th><th class="lab">Route and approximate distance to be covered (in case of pad yatra, procession etc.)</th>
        <td class="val">{{ view.route }}</td></tr>
      <tr><th class="idx">10.</th><th class="lab">Permitted gathering</th>
        <td class="val">{{ view.gathering }}</td></tr>

      <tr><th class="idx">11.</th><th class="lab">NOC obtained from</th>
        <td class="val">
          <div class="grid2 small">
            <div>Local Police :- <b>{{ view.localpolice }}</b></div>
            <div>Traffic Police:- <b>{{ view.traffic }}</b></div>
            <div>Land owning agency:- <b>{{ view.landown }}</b></div>
            <div>Fire Deptt:- <b>{{ view.fire }}</b></div>
          </div>
        </td>
      </tr>

      <tr><th class="idx">12.</th><th class="lab">Permission granted or not, if not the reason for not granting the permission</th>
        <td class="val"><b>{{ view.permission }}</b><div class="muted">{{ view.reason }}</div></td></tr>
    </table>

    <div class="signs">
      <div class="sigL"><div class="sigbox">INSPECTOR (PERMISSION CELL)</div></div>
      <div class="sigR"><div class="sigbox">INCHARGE (PERMISSION CELL), NORTH-WEST (KANJHAWALA), DELHI</div></div>
    </div>

    <div class="meta">
      <div>No. <b>{{ view.appno or "——" }}</b> /ACP(P)RO/PC-(NORTH-WEST)</div>
      <div>Dated : <b>{{ view.dated or "______/_______/2025" }}</b></div>
    </div>

    <section class="tc">
      <div class="ttl">TERMS &amp; CONDITIONS</div>
      <ol>
        <li>Instructions/guidelines issued by the Election Commission of India/State Election Commission in connection with Bye-Elections of MCD-2025 shall be complied with.</li>
        <li>The date, time and place of the programme shall not be changed after issuing this permission.</li>
        <li>Direction and advice of Police Officers on duty should be complied with to maintain law and order.</li>
        <li>No effigies of opponents are allowed to be carried for burning.</li>
        <li>Only 1/3 of the carriage way shall be used and the flow of traffic should remain smooth.</li>
        <li>The organizer shall exercise control over carrying of such articles which may be misused by undesirable elements.</li>
        <li>Pitch of the loudspeaker shall be controlled so that it is not audible beyond the audience.</li>
        <li>The permission is not transferable.</li>
        <li>The Model Code of Conduct regarding Bye-Elections of MCD-2025 will be complied with while organizing rallies, pad yatra etc.</li>
        <li>As per ECI Guidelines, the temporary party office must be 200 meters away from any existing polling station.</li>
        <li>The Permission is subject to Guidelines of Hon’ble Supreme Court / National Green Tribunal.</li>
      </ol>
    </section>
  </div>
</body>
</html>
""")

def html_from_view(view: dict) -> str:
    return HTML_TMPL.render(view=view)

# Fallback ReportLab PDF (compact layout)
def pdf_reportlab(view: dict) -> bytes:
    buf = BytesIO()
    c = canvas.Canvas(buf, pagesize=A3)  # portrait
    W, H = A3
    x = 20*mm
    y = H - 20*mm
    line = 7*mm

    c.setFont("Helvetica-Bold", 16)
    c.drawCentredString(W/2, y, "PERMISSION CELL / SINGLE WINDOW — NORTH-WEST (KANJHAWALA)")
    y -= 10*mm

    c.setFont("Helvetica-Bold", 12)
    c.drawString(x, y, f"Ref No.: {view.get('refno','')}"); y -= line
    c.drawString(x, y, f"App No.: {view.get('appno','')}"); y -= line
    c.drawString(x, y, f"Dated  : {view.get('dated','')}"); y -= (line+2*mm)

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

    rowlbl(1, "Ward & No. (AC/Ward)", f"{view.get('acname','')}  (AC-{view.get('acno','')}) (Ward-{view.get('wardno','')})")
    rowlbl(2, "Election District", view.get("district",""))
    rowlbl(3, "Organizer & Contact", f"{view.get('organizername','')} ({view.get('organizermobile','')})")
    rowlbl(4, "Party & Designation", f"{view.get('party','')}, {view.get('designation','')}")
    rowlbl(5, "Type of Programme", view.get("typeprog",""))
    rowlbl(6, "Venue (PS)", f"{view.get('venueprog','')} ({view.get('psvenue','')})")
    rowlbl(7, "Date", view.get("date",""))
    rowlbl(8, "Time", view.get("time",""))
    rowlbl(9, "Route/Distance", view.get("route",""))
    rowlbl(10, "Permitted gathering", view.get("gathering",""))

    c.setFont("Helvetica-Bold", 11)
    c.drawString(x, y, "11. NOC obtained from"); y -= line
    c.setFont("Helvetica", 11)
    c.drawString(x+18, y, f"Local Police : {view.get('localpolice','')}"); y -= line
    c.drawString(x+18, y, f"Traffic      : {view.get('traffic','')}"); y -= line
    c.drawString(x+18, y, f"Land owning  : {view.get('landown','')}"); y -= line
    c.drawString(x+18, y, f"Fire Deptt   : {view.get('fire','')}"); y -= (line + 2*mm)

    c.setFont("Helvetica-Bold", 11)
    c.drawString(x, y, "12. Permission / Reason"); y -= line
    c.setFont("Helvetica", 11)
    c.drawString(x+18, y, f"Permission : {view.get('permission','')}"); y -= line
    c.drawString(x+18, y, f"Reason     : {view.get('reason','')}"); y -= (line + 2*mm)

    c.setFont("Helvetica", 11)
    c.drawString(x, 25*mm, f"No. {view.get('appno','')} /ACP(P)RO/PC-(NORTH-WEST)")
    c.drawRightString(W - 20*mm, 25*mm, f"Dated : {view.get('dated','')}")

    c.showPage()
    c.save()
    buf.seek(0)
    return buf.read()

def pdf_from_view(view: dict) -> bytes:
    if HAS_WEASYPRINT:
        html = html_from_view(view)
        buf = BytesIO()
        HTML(string=html, base_url=".").write_pdf(buf)  # type: ignore
        buf.seek(0)
        return buf.read()
    # fallback
    return pdf_reportlab(view)

# ========= UI State =========
if "offset" not in st.session_state: st.session_state.offset = 0
if "filter" not in st.session_state: st.session_state.filter = ""
if "selected" not in st.session_state: st.session_state.selected = None
PAGE = 60

st.title("Permission Cell / Single Window — North-West")

# Top bar
c1, c2, c3 = st.columns([0.45, 0.25, 0.30])
with c1:
    ref_query = st.text_input("Search by Reference No.", placeholder="e.g. 28AC44838")
with c2:
    if st.button("Search"):
        with st.spinner("Searching…"):
            hit = search_by_ref(ref_query)
        if hit: st.session_state.selected = hit; st.success("Loaded.")
        else: st.error("No record found.")
with c3:
    new_click = st.button("New Entry", type="primary")

df = load_df()

# Left list + Right details
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
    show_df = tmp.iloc[0: st.session_state.offset + PAGE]

    for _, r in show_df.iterrows():
        lbl = f"**{r['appno']}**  ·  {r.get('organizername','')[:24]}{'…' if len(str(r.get('organizername',''))) > 24 else ''}"
        sub = f"{r.get('party','')}  ·  {r.get('typeprog','')}  ·  {r.get('refno','')}"
        if st.button(lbl, key=f"pick_{r['appno']}"):
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

with right:
    if new_click:
        st.session_state.selected = None

    selected = st.session_state.selected
    st.subheader("A3 Order Preview")

    if selected:
        view = pack_view(selected)
        html = html_from_view(view)
        # Preview HTML in an iframe-like container
        st.components.v1.html(html, height=1150, scrolling=True)

        pdf_data = pdf_from_view(view)
        label = "Download A3 PDF (WeasyPrint)" if HAS_WEASYPRINT else "Download A3 PDF (fallback)"
        st.download_button(label, data=pdf_data, file_name=f"Order_{view['appno'] or 'NA'}.pdf", mime="application/pdf")

        # Also let users download the raw HTML (can print to PDF in browser)
        st.download_button("Download A3 HTML", data=html.encode("utf-8"), file_name=f"Order_{view['appno'] or 'NA'}.html", mime="text/html")

    st.divider()
    st.subheader("Edit / Add")

    with st.form("edit_form", clear_on_submit=False):
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
        submitted_add = st.form_submit_button("Add as new (auto-generate allowed)", type="primary", use_container_width=True)

    if submitted_update or submitted_add:
        row_idx = (selected or {}).get("_row") if submitted_update else None
        df = load_df()

        # Auto-generate if adding and blank
        gen_ref, gen_app = None, None
        if submitted_add and (not refno.strip() or not appno.strip()):
            try:
                gen_ref, gen_app = generate_ids(df, acno)
            except Exception as e:
                st.error(f"Auto-generate failed: {e}")
                st.stop()

        if submitted_update and (not refno.strip() or not appno.strip()):
            st.error("Ref No. and Application No. are required for update.")
            st.stop()

        ref_check = gen_ref or refno.strip()
        app_check = gen_app or appno.strip()
        ref_unique, app_unique = check_unique(df, ref_check, app_check, row_idx)
        if not ref_unique: st.error("Duplicate Reference No. — must be unique."); st.stop()
        if not app_unique: st.error("Duplicate Application No. — must be unique."); st.stop()

        payload = {
            "refno": ref_check,
            "appno": app_check,
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
            else:
                with st.spinner("Adding new entry…"):
                    new_row = add_row(payload)
                st.success(f"Added as new (row {new_row}).")

            # Refresh & re-select
            st.cache_data.clear()
            df2 = load_df()
            match = df2.loc[df2["refno"] == payload["refno"]]
            st.session_state.selected = match.iloc[0].to_dict() if not match.empty else None
            st.session_state.offset = 0
            st.rerun()
        except Exception as e:
            st.error(f"Operation failed: {e}")
