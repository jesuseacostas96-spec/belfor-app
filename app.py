import streamlit as st
import zipfile, io, re, json, os
from collections import OrderedDict, defaultdict
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

st.set_page_config(page_title="BELFOR Equipment Log", page_icon="🏗️", layout="centered")
st.markdown("""
<style>
.block-container{padding-top:2rem;max-width:760px}
.stButton>button{background:#0D2B55;color:white;border-radius:6px;border:none;padding:0.6rem 2rem;font-weight:700;font-size:15px;width:100%}
</style>""", unsafe_allow_html=True)
st.markdown("""
<div style='background:#0D2B55;padding:1.5rem 2rem;border-radius:10px;margin-bottom:1.5rem'>
<div style='color:#D4A017;font-size:0.75rem;letter-spacing:3px;font-weight:700'>BELFOR PROPERTY RESTORATION</div>
<div style='color:white;font-size:1.6rem;font-weight:800;margin-top:4px'>Equipment Log Generator</div>
<div style='color:#90b4d4;font-size:0.85rem;margin-top:4px'>WhatsApp chat → BELFOR Weekly EQ Template (1 tab per floor)</div>
</div>""", unsafe_allow_html=True)

try:
    api_key = st.secrets["ANTHROPIC_API_KEY"]
except Exception:
    st.error("API key not found. Add ANTHROPIC_API_KEY to Streamlit Secrets.")
    st.stop()

with st.expander("📱 How to export WhatsApp chat"):
    st.markdown("1. Open WhatsApp group\n2. Tap group name\n3. Export Chat → Without Media\n4. Upload the .zip here")

uploaded   = st.file_uploader("Upload WhatsApp .zip or .txt", type=["zip","txt"], label_visibility="collapsed")
job_name   = st.text_input("Job name / address", placeholder="e.g. Emera Port Royale - 3333 Port Royale Dr")
job_number = st.text_input("Job # / JDE", placeholder="e.g. JDE 100623171")

def extract_txt(f):
    if f.name.endswith(".zip"):
        with zipfile.ZipFile(io.BytesIO(f.read())) as z:
            for name in z.namelist():
                if name.endswith(".txt"):
                    return z.read(name).decode("utf-8", errors="ignore")
    return f.read().decode("utf-8", errors="ignore")

def parse_messages(raw):
    lines = raw.replace("\r","").split("\n")
    messages, current = [], None
    for line in lines:
        m = re.match(r'^\[(\d{1,2}/\d{1,2}/\d{2,4}),\s*(\d{1,2}:\d{2}:\d{2}\s*[AP]M)\]\s*([^:]+):\s*(.+)$', line)
        if m:
            if current: messages.append(current)
            current = {"date":m.group(1),"time":m.group(2),"sender":m.group(3).strip(),"text":m.group(4).strip()}
        elif current and line.strip() and not line.startswith("‎"):
            current["text"] += "\n" + line.strip()
    if current: messages.append(current)
    return messages

def extract_equipment(messages):
    import anthropic
    client = anthropic.Anthropic(api_key=api_key)
    SYSTEM = """Extract equipment placement from WhatsApp restoration messages.
Only confirmed placements (not requests/assessments). Return JSON array only:
[{"date":"3/15/26","unit":"Unit 702","ams":["CODE1"],"dhs":["CODE2"]}]
No explanation, no markdown."""
    all_results = []
    for i in range(0, len(messages), 50):
        chunk = messages[i:i+50]
        text = "\n".join(f"[{m['date']} {m['time']}] {m['sender']}: {m['text']}" for m in chunk)
        try:
            resp = client.messages.create(
                model="claude-sonnet-4-5",
                max_tokens=4000,
                system=SYSTEM,
                messages=[{"role":"user","content":f"Extract:\n\n{text}"}]
            )
            raw = resp.content[0].text.strip().replace("```json","").replace("```","")
            all_results.extend(json.loads(raw))
        except Exception as e:
            st.warning(f"Chunk error: {e}")
    consolidated = OrderedDict()
    for r in all_results:
        key = (r.get("date",""), r.get("unit",""))
        if key not in consolidated:
            consolidated[key] = {"date":r["date"],"unit":r["unit"],"ams":[],"dhs":[]}
        consolidated[key]["ams"].extend(r.get("ams",[]))
        consolidated[key]["dhs"].extend(r.get("dhs",[]))
    return list(consolidated.values())

def get_floor(unit):
    u = unit.lower()
    for k,v in [("7th",7),("6th",6),("5th",5),("4th",4),("3rd",3),("2nd",2),("1st",1)]:
        if k in u: return v
    try: return int(''.join(filter(str.isdigit,unit))) // 100
    except: return 0

FLOOR_LABELS = {7:"7th Floor",6:"6th Floor",5:"5th Floor",4:"4th Floor",
                3:"3rd Floor",2:"2nd Floor",1:"1st Floor"}

def sort_unit(unit):
    if "hallway" in unit.lower(): return (0, unit)
    try: return (1, -int(''.join(filter(str.isdigit,unit))))
    except: return (1, unit)

def build_excel(equipment, job_name, job_number):
    AM_BG = "EBF4FB"; DH_BG = "FEF9E7"
    thin  = Side(style="thin",   color="CCCCCC")
    thick = Side(style="medium", color="0D2B55")

    template_path = os.path.join(os.path.dirname(__file__), "Weekly_EQ.xlsx")
    wb = load_workbook(template_path)

    floors = defaultdict(list)
    for loc in equipment:
        floors[get_floor(loc["unit"])].append(loc)

    first_floor_num = sorted(floors.keys(), reverse=True)[0]
    wb.active.title = FLOOR_LABELS.get(first_floor_num, f"Floor {first_floor_num}")

    for i, floor_num in enumerate(sorted(floors.keys(), reverse=True)):
        label = FLOOR_LABELS.get(floor_num, f"Floor {floor_num}")
        if i == 0:
            ws = wb.active
        else:
            src_title = FLOOR_LABELS.get(first_floor_num, f"Floor {first_floor_num}")
            ws = wb.copy_worksheet(wb[src_title])
            ws.title = label

        ws["A2"] = job_number or ""
        ws["E2"] = job_name or ""

        dates = sorted(set(loc["date"] for loc in floors[floor_num]))
        ws["I10"] = dates[0] if dates else ""

        sorted_locs = sorted(floors[floor_num], key=lambda x: sort_unit(x["unit"]))
        all_rows = []
        for loc in sorted_locs:
            first = True
            for code in loc["ams"]:
                all_rows.append({"code":code,"desc":"AM","location":loc["unit"] if first else "","first":first})
                first = False
            for code in loc["dhs"]:
                all_rows.append({"code":code,"desc":"DH","location":loc["unit"] if first else "","first":first})
                first = False

        # Clear previous data
        for r in range(13, 63):
            for c in [3,5,7,9]:
                ws.cell(row=r, column=c).value = None

        for idx, eq in enumerate(all_rows[:50]):
            row = 13 + idx
            is_am = eq["desc"] == "AM"
            bg = AM_BG if is_am else DH_BG
            for col, val in [(3,eq["code"]),(5,eq["desc"]),(7,eq["location"]),(9,2)]:
                c = ws.cell(row=row, column=col, value=val)
                c.fill = PatternFill("solid", start_color=bg)
                c.font = Font(name="Arial", size=9, bold=(col==5),
                              color=("1A4F7A" if is_am else "7B5200"))
                c.alignment = Alignment(
                    horizontal="center" if col in (5,9) else "left",
                    vertical="center")
                c.border = Border(left=thin,right=thin,
                                  top=(thick if eq["first"] else thin),
                                  bottom=thin)

        am_c = sum(1 for r in all_rows if r["desc"]=="AM")
        dh_c = sum(1 for r in all_rows if r["desc"]=="DH")
        ws.cell(row=63, column=9).value = am_c
        ws.cell(row=65, column=9).value = dh_c

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)

    total_am = sum(len(l["ams"]) for l in equipment)
    total_dh = sum(len(l["dhs"]) for l in equipment)
    apt_c    = sum(1 for l in equipment if "unit" in l["unit"].lower())
    hall_c   = sum(1 for l in equipment if "hallway" in l["unit"].lower())
    return buf, total_am, total_dh, apt_c, hall_c, len(floors)

if uploaded:
    st.markdown("### Process Chat")
    if st.button("⚙️ Generate BELFOR Weekly EQ Template"):
        with st.spinner("Reading chat..."):
            raw = extract_txt(uploaded)
            messages = parse_messages(raw)
        st.info(f"Found {len(messages)} messages")

        with st.spinner("Analyzing with AI..."):
            equipment = extract_equipment(messages)

        if not equipment:
            st.error("No equipment records found.")
        else:
            with st.spinner("Building Excel by floor..."):
                try:
                    buf, total_am, total_dh, apt_c, hall_c, num_floors = build_excel(
                        equipment, job_name, job_number)
                    st.success(f"✅ Done — {num_floors} floor tabs generated!")
                    c1,c2,c3,c4 = st.columns(4)
                    c1.metric("Floors", num_floors)
                    c2.metric("Locations", apt_c + hall_c)
                    c3.metric("Air Movers", total_am)
                    c4.metric("Dehumidifiers", total_dh)
                    fname = (job_number or "Equipment").replace(" ","_") + "_Weekly_EQ.xlsx"
                    st.download_button(
                        "⬇️ Download BELFOR Weekly EQ",
                        buf, fname,
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True)
                except Exception as e:
                    st.error(f"Error: {e}")

st.markdown("---")
st.markdown("<div style='text-align:center;color:#999;font-size:0.75rem'>BELFOR Property Restoration • Powered by Claude AI</div>", unsafe_allow_html=True)
