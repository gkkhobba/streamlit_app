# streamlit_app.py
# Permission Cell / Single Window — Streamlit + Google Sheets (gspread)

import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
import re

st.set_page_config(page_title="Permission Cell / Single Window — North-West",
                   layout="wide", initial_sidebar_state="expanded")

# ====== Config via secrets ======
SPREADSHEET_ID = st.secrets["google_sheets"]["spreadsheet_id"]
SHEET_NAME     = st.secrets["google_sheets"]["sheet_name"]

# ====== Connect Google Sheets ======
@st.cache_resource(show_spinner=False)
def _connect_ws():
    # Either service-account dict or GCP default (we use secrets -> dict)
    sa_info = dict(st.secrets["gcp_service_account"])
    scopes = ["https://www.googleapis.com/auth/spreadsheets",
              "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_info(sa_info, scopes=scopes)
    gc = gspread.authorize(creds)
    ss = gc.open_by_key(SPREADSHEET_ID)
    ws = ss.worksheet(SHEET_NAME)
    return ws

ws = _connect_ws()

# ====== Column map / helpers ======
NEED = [
    'refno','appno','dated','acname','acno','district','organizername','organizermobile',
    'party','designation','typeprog','venueprog','psvenue','date','time','route','gathering',
    'localpolice','traffic','landown','fire','permission','reason','orderno','wardno','orderdate'
]

def _norm(s: str) -> str:
    return re.sub(r'[\s_]+', '', str(s or '').strip().lower())

@st.cache_data(ttl=30, show_spinner=False)
def get_values():
    # all display values as strings
    return ws.get_all_values()

def get_map(values):
    if not values or len(values) < 1:
        st.stop()
    heads = [_norm(h) for h in values[0]]
    m = {}
    for k in NEED:
        try:
            m[k] = heads.index(_norm(k))
        except ValueError:
            st.error(f'Missing header "{k}" in sheet.')
            st.stop()
    return m

def pack(row, m, sheet_row):
    # Return list similar to your Apps Script pack
    return [
        row[m['refno']], row[m['appno']], row[m['dated']],
        row[m['acname']], row[m['acno']], row[m['district']],
        row[m['organizername']], row[m['organizermobile']],
        row[m['party']], row[m['designation']], row[m['typeprog']],
        row[m['venueprog']], row[m['psvenue']], row[m['date']],
        row[m['time']], row[m['route']], row[m['gathering']],
        row[m['localpolice']], row[m['traffic']], row[m['landown']],
        row[m['fire']], row[m['permission']], row[m['reason']],
        row[m['orderno']], row[m['wardno']], row[m['orderdate']], sheet_row
    ]

def check_unique(refno: str, appno: str, exclude_row: int | None):
    values = get_values(); m = get_map(values)
    ref_unique = True; app_unique = True
    for r in range(1, len(values)):
        row_idx = r + 1
        if exclude_row and row_idx == exclude_row:
            continue
        vref = str(values[r][m['refno']])
        vapp = str(values[r][m['appno']])
        if refno and vref == str(refno):
            ref_unique = False
        if appno and vapp == str(appno):
            app_unique = False
        if not ref_unique and not app_unique:
            break
    return ref_unique, app_unique

def generate_ids(acno_raw: str | None):
    values = get_values(); m = get_map(values)
    # appno = max numeric + 1
    max_app = 0
    for r in range(1, len(values)):
        s = re.sub(r'\D', '', str(values[r][m['appno']] or ''))
        if s.isdigit():
            max_app = max(max_app, int(s))
    appno = str(max_app + 1)

    # refno = {AC}AC{suffix} ; suffix grows
    acno = re.sub(r'\D', '', str(acno_raw or '00'))
    prefix = (acno.zfill(2) if acno else '00') + 'AC'
    max_suffix = 39999
    for r in range(1, len(values)):
        ref = str(values[r][m['refno']] or '').strip()
        if ref.startswith(prefix):
            suf = ref[len(prefix):]
            if suf.isdigit():
                max_suffix = max(max_suffix, int(suf))
    refno = prefix + str(max_suffix + 1).zfill(5)
    return refno, appno

def update_row(sheet_row: int, updates: dict):
    values = get_values(); m = get_map(values)
    if not sheet_row or sheet_row < 2:
        raise ValueError("Invalid row index.")
    row = ws.row_values(sheet_row)
    if not row:
        row = [""] * ws.col_count

    # enforce uniqueness if provided
    ref_new = updates.get('refno')
    app_new = updates.get('appno')
    if ref_new or app_new:
        ref_ok, app_ok = check_unique(ref_new, app_new, sheet_row)
        if ref_new and not ref_ok:
            raise ValueError("Duplicate Reference No. — must be unique.")
        if app_new and not app_ok:
            raise ValueError("Duplicate Application No. — must be unique.")

    # expand row to num columns (for safe write)
    cols = ws.col_count
    row += [""] * (cols - len(row))

    for k, v in updates.items():
        if k in m:
            row[m[k]] = v

    # write back
    rng = f"A{sheet_row}:{gspread.utils.rowcol_to_a1(1, cols)[0]}{sheet_row}"
    ws.update(rng, [row[:cols]])

def add_new_entry(entry: dict):
    values = get_values(); m = get_map(values)
    ref = (entry.get('refno') or '').strip()
    app = (entry.get('appno') or '').strip()

    if not ref or not app:
        gen_ref, gen_app = generate_ids(entry.get('acno'))
        ref = ref or gen_ref
        app = app or gen_app
    else:
        ref_ok, app_ok = check_unique(ref, app, None)
        if not ref_ok:
            raise ValueError("Duplicate Reference No. — must be unique.")
        if not app_ok:
            raise ValueError("Duplicate Application No. — must be unique.")

    payload = {**entry, 'refno': ref, 'appno': app}
    # Build an output row aligned to headers
    out = [""] * ws.col_count
    for k in NEED:
        if k in m and k in payload:
            out[m[k]] = payload.get(k, "")
    ws.append_row(out, value_input_option="USER_ENTERED")

    # return the appended row
    last = ws.get_all_values()[-1]
    sheet_row = ws.row_count  # not reliable if empty below; safer compute:
    sheet_row = len(get_values())  # actual used rows
    return pack(last + [""] * (ws.col_count - len(last)), m, sheet_row)

def search_by_ref(refno: str):
    values = get_values(); m = get_map(values)
    needle = _norm(refno)
    for r in range(1, len(values)):
        if _norm(values[r][m['refno']]) == needle:
            return pack(values[r] + [""] * (ws.col_count - len(values[r])), m, r + 1)
    return None

def get_by_app(appno: str):
    values = get_values(); m = get_map(values)
    needle = str(appno).strip()
    for r in range(1, len(values)):
        if str(values[r][m['appno']]).strip() == needle:
            return pack(values[r] + [""] * (ws.col_count - len(values[r])), m, r + 1)
    return None

def list_applications(limit=60, offset=0, query=""):
    values = get_values(); m = get_map(values)
    items = []
    for r in range(1, len(values)):
        row = values[r]
        it = dict(
            appno=(row[m['appno']] or "").strip(),
            refno=(row[m['refno']] or "").strip(),
            dated=(row[m['dated']] or "").strip(),
            organizername=(row[m['organizername']] or "").strip(),
            party=(row[m['party']] or "").strip(),
            typeprog=(row[m['typeprog']] or "").strip(),
            rowIndex=r+1
        )
        if it["appno"]:
            items.append(it)

    q = (query or "").lower()
    if q:
        items = [it for it in items if any(q in (str(it[k]) or "").lower()
                                           for k in ["appno","refno","organizername","party","typeprog"])]

    def num(app):  # sort desc by numeric appno if possible
        s = re.sub(r'\D', '', app or '')
        return int(s) if s.isdigit() else -1
    items.sort(key=lambda it: (num(it["appno"]), it["appno"]), reverse=True)

    total = len(items)
    page = items[offset: offset+limit]
    has_more = offset + len(page) < total
    next_offset = offset + len(page)
    return page, total, has_more, next_offset

def dash_value(x, default="—"):
    return x if (x and str(x).strip()) else default

def render_order_html(packed):
    # packed indexes = same as Apps Script
    # 0 refno, 1 appno, 2 dated, 3 acname, 4 acno, 5 district, 6 organizer, 7 mobile,
    # 8 party, 9 desg, 10 type, 11 venue, 12 psvenue, 13 date, 14 time, 15 route, 16 gathering,
    # 17 localpolice, 18 traffic, 19 landown, 20 fire, 21 permission, 22 reason, 23 orderno, 24 wardno, 25 orderdate, 26 row
    def ph(s): return dash_value(s)
    html = f"""
    <style>
      body {{ font-family: system-ui, -apple-system, Segoe UI, Roboto, Arial, "Noto Sans", sans-serif; color:#0f172a; }}
      .sheet {{
        width: min(1120px, 100%); margin: 8px auto; background:#fff; border:1px solid #ddd; border-radius:8px; padding:18px 20px;
        box-shadow: 0 6px 20px rgba(0,0,0,.06);
      }}
      .head {{ display:grid; grid-template-columns: 80px 1fr 80px; gap:8px; align-items:center; border:2px solid #111; border-radius:8px; padding:8px 10px; }}
      .head img {{ width:72px; height:72px; object-fit:contain; }}
      .title {{ text-align:center; line-height:1.1; }}
      .title .t1{{ font-weight:900; font-size:20px; text-transform:uppercase; }}
      .title .t2{{ font-weight:800; font-size:16px; text-transform:uppercase; }}
      .title .t3{{ font-weight:700; font-size:14px; text-transform:uppercase; }}
      .strip {{ display:grid; grid-template-columns:1fr 1fr 1fr; gap:8px; margin:8px 0; }}
      .box {{ border:1.4px solid #111; border-radius:6px; padding:6px 8px; font-weight:800; background:#fff; }}
      table {{ width:100%; border-collapse:collapse; }}
      td, th {{ border:1px solid #111; padding:6px 8px; vertical-align:top; }}
      th.idx {{ width:50px; font-weight:900; text-align:center; }}
      th.lab {{ width:45%; text-align:left; font-weight:800; }}
      td.val {{ width:55%; font-weight:600; }}
      .grid2 {{ display:grid; grid-template-columns:1fr 1fr; gap:4px 16px; }}
      .muted {{ color:#64748b; font-weight:600; }}
      .order {{ text-align:center; font-weight:900; margin:10px 0 8px; }}
      @media (max-width: 640px) {{
        .strip {{ grid-template-columns:1fr; }}
      }}
    </style>
    <div class="sheet">
      <div class="head">
        <img src="https://upload.wikimedia.org/wikipedia/commons/5/55/Emblem_of_India.svg">
        <div class="title">
          <div class="t1">OFFICE OF THE INCHARGE</div>
          <div class="t2">PERMISSION CELL / SINGLE WINDOW</div>
          <div class="t2">DISTRICT ELECTION OFFICER : NORTH-WEST</div>
          <div class="t3">KANJHAWALA DELHI - 110081</div>
        </div>
        <img src="https://upload.wikimedia.org/wikipedia/commons/3/32/Swachh_Bharat_Mission_Logo.svg">
      </div>

      <div class="strip">
        <div class="box">Ref No. <b>{ph(packed[0])}</b></div>
        <div class="box">Application No. <b>{ph(packed[1])}</b></div>
        <div class="box">Dated : <b>{ph(packed[2] or "______/_______/2025")}</b></div>
      </div>

      <div class="order">ORDER</div>

      <table>
        <tr><th class="idx">1.</th><th class="lab">Name of Municipal Corporation Ward &amp; No.</th>
            <td class="val">{ph(packed[3])} <span class="muted">(AC-{ph(packed[4])})</span> <span class="muted">(Ward-{ph(packed[24])})</span></td></tr>
        <tr><th class="idx">2.</th><th class="lab">Name of the Election District</th><td class="val">{ph(packed[5])}</td></tr>
        <tr><th class="idx">3.</th><th class="lab">Name of the organizer &amp; Contact No</th>
            <td class="val">{ph(packed[6])} ( {ph(packed[7])} )</td></tr>
        <tr><th class="idx">4.</th><th class="lab">Party affiliation and his designation</th>
            <td class="val">{ph(packed[8])}, {ph(packed[9])}</td></tr>
        <tr><th class="idx">5.</th><th class="lab">Type of programme (meeting procession, rally...)</th><td class="val">{ph(packed[10])}</td></tr>
        <tr><th class="idx">6.</th><th class="lab">Name of venue with police Station</th>
            <td class="val">{ph(packed[11])} ( {ph(packed[12])} )</td></tr>
        <tr><th class="idx">7.</th><th class="lab">Date</th><td class="val">{ph(packed[13] or "______/_______/2025")}</td></tr>
        <tr><th class="idx">8.</th><th class="lab">Timing of Programme</th><td class="val">{ph(packed[14])}</td></tr>
        <tr><th class="idx">9.</th><th class="lab">Route / Distance</th><td class="val">{ph(packed[15])}</td></tr>
        <tr><th class="idx">10.</th><th class="lab">Permitted gathering</th><td class="val">{ph(packed[16])}</td></tr>
        <tr><th class="idx">11.</th><th class="lab">NOC obtained from</th>
          <td class="val">
            <div class="grid2">
              <div>Local Police :- <b>{ph(packed[17])}</b></div>
              <div>Traffic Police:- <b>{ph(packed[18])}</b></div>
              <div>Land owning agency:- <b>{ph(packed[19])}</b></div>
              <div>Fire Deptt:- <b>{ph(packed[20])}</b></div>
            </div>
          </td>
        </tr>
        <tr><th class="idx">12.</th><th class="lab">Permission / reason if not granted</th>
            <td class="val"><b>{ph(packed[21])}</b><div class="muted">{ph(packed[22])}</div></td></tr>
      </table>

      <div style="display:flex;justify-content:space-between;margin-top:12px;font-weight:800">
        <div>No. <b>{ph(packed[1])}</b> /ACP(P)RO/PC-(NORTH-WEST)</div>
        <div>Dated : <b>{ph(packed[2] or "______/_______/2025")}</b></div>
      </div>

      <div style="margin-top:14px;font-weight:900">TERMS &amp; CONDITIONS</div>
      <ol style="margin-top:4px;padding-left:18px;line-height:1.25">
        <li>Instructions/guidelines issued by the ECI/SEC for Bye-Elections of MCD-2025 shall be complied with.</li>
        <li>No change in date/time/place after issue of permission.</li>
        <li>Follow directions of Police Officers on duty.</li>
        <li>No burning of effigies.</li>
        <li>Only 1/3 carriage way to be used; traffic to remain smooth.</li>
        <li>Organizer shall control articles that may be misused by undesirable elements.</li>
        <li>Loudspeaker volume so that it is not audible beyond audience.</li>
        <li>Permission is non-transferable.</li>
        <li>Model Code of Conduct to be complied with during events.</li>
        <li>Temporary party office must be 200m away from any polling station.</li>
        <li>Subject to guidelines of Hon’ble Supreme Court / NGT.</li>
      </ol>
    </div>
    """
    return html

# ====== Session state ======
if "offset" not in st.session_state: st.session_state.offset = 0
if "query"  not in st.session_state: st.session_state.query = ""
if "selected_app" not in st.session_state: st.session_state.selected_app = None
if "selected_pack" not in st.session_state: st.session_state.selected_pack = None
if "selected_row" not in st.session_state: st.session_state.selected_row = None

# ====== Sidebar: list + paging ======
with st.sidebar:
    st.markdown("### Applications")
    st.text_input("Filter (app/ref/organizer)", key="query",
                  placeholder="Search text…")
    colA, colB = st.columns([1,1])
    with colA:
        if st.button("Refresh", use_container_width=True):
            st.session_state.offset = 0
            st.rerun()
    with colB:
        if st.button("New Entry", type="primary", use_container_width=True):
            st.session_state.selected_pack = None
            st.session_state.selected_row = None
            st.session_state.selected_app = None
            st.session_state.offset = st.session_state.offset  # keep
            st.rerun()

    page, total, has_more, next_offset = list_applications(
        limit=60, offset=st.session_state.offset, query=st.session_state.query)

    for it in page:
        lbl = f"**{it['appno']}**  ·  {it['party'] or ''}"
        sub = f"{(it['organizername'] or '')[:30]}{'…' if (it['organizername'] and len(it['organizername'])>30) else ''}"
        if st.button(lbl + "\n" + sub, key=f"app_{it['appno']}", use_container_width=True):
            packrow = get_by_app(it['appno'])
            if packrow:
                st.session_state.selected_app = it['appno']
                st.session_state.selected_pack = packrow
                st.session_state.selected_row = packrow[-1]
                st.toast(f"Loaded {it['appno']}")
                st.rerun()

    if has_more:
        if st.button("Load more", use_container_width=True):
            st.session_state.offset = next_offset
            st.rerun()
    else:
        st.caption(f"{total} items")

# ====== Main area ======

st.title("Permission Cell / Single Window — North-West")
with st.container(border=True):
    # Search by Reference No.
    ref = st.text_input("Search by Reference No.", placeholder="e.g. 28AC44838")
    col1, col2, col3 = st.columns([1,1,1])
    with col1:
        if st.button("Find by Ref No."):
            if not ref.strip():
                st.error("Enter a Reference No.")
            else:
                with st.spinner("Searching…"):
                    row = search_by_ref(ref.strip())
                if row:
                    st.session_state.selected_pack = row
                    st.session_state.selected_row = row[-1]
                    st.session_state.selected_app = row[1]
                    st.toast("Loaded")
                    st.rerun()
                else:
                    st.error("No record found")

    with col2:
        if st.session_state.selected_pack:
            # Print: offer HTML download (user can print to PDF)
            html = render_order_html(st.session_state.selected_pack)
            st.download_button("Download print-friendly HTML", data=html,
                               file_name=f"Order_{st.session_state.selected_pack[1]}.html",
                               mime="text/html", use_container_width=True)

    with col3:
        if st.session_state.selected_pack:
            st.success(f"Selected: App {st.session_state.selected_pack[1]} | Ref {st.session_state.selected_pack[0]}")

# Show the order (printable view)
if st.session_state.selected_pack:
    st.components.v1.html(render_order_html(st.session_state.selected_pack),
                          height=1280, scrolling=True)

# ====== Form: Add / Edit ======
st.markdown("### Add / Edit Entry")

is_editing = st.session_state.selected_pack is not None
initial = {}
if is_editing:
    p = st.session_state.selected_pack
    initial = dict(
        refno=p[0], appno=p[1], dated=p[2],
        acname=p[3], acno=p[4], district=p[5],
        organizername=p[6], organizermobile=p[7],
        party=p[8], designation=p[9], typeprog=p[10],
        venueprog=p[11], psvenue=p[12], date=p[13], time=p[14], route=p[15],
        gathering=p[16], localpolice=p[17], traffic=p[18], landown=p[19], fire=p[20],
        permission=p[21], reason=p[22], orderno=p[23], wardno=p[24], orderdate=p[25]
    )

with st.form("edit_add_form", clear_on_submit=False, border=True):
    cols = st.columns(3)
    refno = cols[0].text_input("Ref No.", value=initial.get("refno",""))
    appno = cols[1].text_input("Application No.", value=initial.get("appno",""))
    dated = cols[2].text_input("Dated (DD-MM-YYYY)", value=initial.get("dated",""))

    cols2 = st.columns(3)
    acname = cols2[0].text_input("Ward / Area Name", value=initial.get("acname",""))
    acno   = cols2[1].text_input("AC No.", value=initial.get("acno",""))
    wardno = cols2[2].text_input("Ward No.", value=initial.get("wardno",""))

    cols3 = st.columns(3)
    district = cols3[0].text_input("District", value=initial.get("district",""))
    organizername = cols3[1].text_input("Organizer", value=initial.get("organizername",""))
    organizermobile = cols3[2].text_input("Organizer Mobile", value=initial.get("organizermobile",""))

    cols4 = st.columns(3)
    party = cols4[0].text_input("Party", value=initial.get("party",""))
    designation = cols4[1].text_input("Designation", value=initial.get("designation",""))
    typeprog = cols4[2].text_input("Type of Programme", value=initial.get("typeprog",""))

    venueprog = st.text_input("Venue", value=initial.get("venueprog",""))
    psvenue   = st.text_input("Police Station (PS)", value=initial.get("psvenue",""))

    cols5 = st.columns(3)
    date = cols5[0].text_input("Date (DD-MM-YYYY)", value=initial.get("date",""))
    time = cols5[1].text_input("Time (e.g., 02:00 PM TO 05:00 PM)", value=initial.get("time",""))
    route = cols5[2].text_input("Route / Distance", value=initial.get("route",""))

    cols6 = st.columns(3)
    gathering = cols6[0].text_input("Permitted Gathering", value=initial.get("gathering",""))
    localpolice = cols6[1].text_input("Local Police", value=initial.get("localpolice",""))
    traffic = cols6[2].text_input("Traffic", value=initial.get("traffic",""))

    cols7 = st.columns(3)
    landown = cols7[0].text_input("Land Owning", value=initial.get("landown",""))
    fire = cols7[1].text_input("Fire", value=initial.get("fire",""))
    permission = cols7[2].text_input("Permission", value=initial.get("permission",""))

    reason = st.text_area("Reason (if any)", value=initial.get("reason",""))

    cols8 = st.columns(3)
    orderno   = cols8[0].text_input("Order No. (optional)", value=initial.get("orderno",""))
    orderdate = cols8[1].text_input("Order Date (DD-MM-YYYY)", value=initial.get("orderdate",""))
    # cols8[2] left blank

    left, mid, right = st.columns([1,1,2])
    with left:
        check = st.form_submit_button("Check duplicates")
    with mid:
        add_btn = st.form_submit_button("Add as New", type="primary")
    with right:
        upd_btn = st.form_submit_button("Update Selected")

    payload = dict(
        refno=refno, appno=appno, dated=dated, acname=acname, acno=acno, district=district,
        organizername=organizername, organizermobile=organizermobile, party=party,
        designation=designation, typeprog=typeprog, venueprog=venueprog, psvenue=psvenue,
        date=date, time=time, route=route, gathering=gathering, localpolice=localpolice,
        traffic=traffic, landown=landown, fire=fire, permission=permission, reason=reason,
        orderno=orderno, orderdate=orderdate, wardno=wardno
    )

    if check:
        ref_ok, app_ok = check_unique(refno, appno, st.session_state.selected_row if is_editing else None)
        st.info(f"Ref Unique: **{ref_ok}**, App Unique: **{app_ok}**")

    if add_btn:
        try:
            with st.spinner("Adding new entry…"):
                row = add_new_entry(payload)
            st.success("Added as new.")
            st.session_state.selected_pack = row
            st.session_state.selected_row  = row[-1]
            st.session_state.selected_app  = row[1]
            st.session_state.offset = 0   # refresh list from start
            get_values.clear()            # invalidate cache
            st.rerun()
        except Exception as e:
            st.error(str(e))

    if upd_btn:
        if not is_editing or not st.session_state.selected_row:
            st.error("Select a record first (from sidebar or via Ref No. search).")
        else:
            try:
                with st.spinner("Updating record…"):
                    update_row(st.session_state.selected_row, payload)
                st.success("Updated.")
                get_values.clear()
                # reload the updated row
                row = search_by_ref(payload.get("refno") or st.session_state.selected_pack[0])
                if row:
                    st.session_state.selected_pack = row
                    st.session_state.selected_row  = row[-1]
                    st.session_state.selected_app  = row[1]
                st.rerun()
            except Exception as e:
                st.error(str(e))
