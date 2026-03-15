import streamlit as st
import zipfile, io, re, json
from collections import OrderedDict
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

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
<div style='color:#90b4d4;font-size:0.85rem;margin-top:4px'>WhatsApp chat to Excel</div>
</div>""", unsafe_allow_html=True)

try:
    api_key = st.secrets["ANTHROPIC_API_KEY"]
except Exception:
    st.error("API key not found. Add ANTHROPIC_API_KEY to Streamlit Secrets.")
    st.stop()

with st.expander("How to export WhatsApp chat"):
    st.markdown("1. Open WhatsApp group\n2. Tap group name\n3. Export Chat > Without Media\n4. Upload the .zip here")

uploaded = st.file_uploader("Upload WhatsApp .zip or .txt", type=["zip","txt"], label_visibility="collapsed")
project_name = st.text_input("Project name (optional)", placeholder="e.g. Emera Port Royale - JDE 100623171")

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

def extract_equipment(messages, api_key):
    import anthropic
    client = anthropic.Anthropic(api_key=api_key)
    SYSTEM = """Extract equipment placement from WhatsApp restoration messages.
Only confirmed placements (not requests). Return JSON array only:
[{"date":"3/15/26","unit":"Unit 702","ams":["CODE1"],"dhs":["CODE2"]}]"""
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

def sort_key(r):
    floor_map = {"7th":7,"6th":6,"5th":5,"4th":4,"3rd":3,"2nd":2,"1st":1}
    u = r["unit"]
    if "hallway" in u.lower():
        for k,v in floor_map.items():
            if k.lower() in u.lower(): return (1,-v,u)
        return (1,0,u)
    try: return (0,-int(''.join(filter(str.isdigit,u))),u)
    except: return (0,0,u)

def build_excel(rows, project):
    C_NAVY,C_BLUE,C_ACCENT = "0D2B55","1560A4","2980B9"
    C_AM_HDR,C_DH_HDR = "1A4F7A","7B5200"
    C_AM_L,C_AM_D = "EBF4FB","D6EAF8"
    C_DH_L,C_DH_D = "FEF9E7","FDEBD0"
    C_GOLD,C_WHITE = "D4A017","FFFFFF"
    med = Side(style="medium",color="0D2B55")
    thin = Side(style="thin",color="B8CCE4")
    def tb(c): c.border=Border(left=med,right=med,top=med,bottom=med)
    def nb(c): c.border=Border(left=thin,right=thin,top=thin,bottom=thin)
    def hset(c,val,bg,fg=C_WHITE,sz=10,bold=True,center=True,wrap=False):
        c.value=val
        c.font=Font(name="Arial",bold=bold,color=fg,size=sz)
        c.fill=PatternFill("solid",start_color=bg)
        c.alignment=Alignment(horizontal="center" if center else "left",vertical="center",wrap_text=wrap,indent=0 if center else 1)
    wb=Workbook(); ws=wb.active; ws.title="Equipment Log"
    ws.sheet_view.showGridLines=False; ws.freeze_panes="A5"
    COLS=[("No.",5),("Date",10),("Unit / Location",24),("Air Mover Codes",54),("AMs",7),("Dehumidifier Codes",40),("DHs",7),("Total",8)]
    for i,(_,w) in enumerate(COLS,1): ws.column_dimensions[get_column_letter(i)].width=w
    ws.merge_cells("A1:H1"); hset(ws["A1"],"BELFOR PROPERTY RESTORATION",C_NAVY,C_GOLD,sz=10); tb(ws["A1"]); ws.row_dimensions[1].height=18
    ws.merge_cells("A2:H2"); hset(ws["A2"],(project or "EQUIPMENT PLACEMENT LOG").upper(),C_NAVY,C_WHITE,sz=15); tb(ws["A2"]); ws.row_dimensions[2].height=32
    ws.merge_cells("A3:D3"); hset(ws["A3"],"Water Damage Mitigation — Equipment Tracking",C_BLUE,C_WHITE,sz=10,bold=False,center=False); tb(ws["A3"])
    ws.merge_cells("E3:H3")
    dates=sorted(set(r["date"] for r in rows))
    hset(ws["E3"],f"Date: {dates[0]}" if len(dates)==1 else f"{dates[0]} - {dates[-1]}",C_BLUE,C_WHITE,sz=10,bold=False); tb(ws["E3"]); ws.row_dimensions[3].height=18
    hdr_styles={"Air Mover Codes":(C_AM_HDR,C_WHITE),"AMs":(C_AM_HDR,C_WHITE),"Dehumidifier Codes":(C_DH_HDR,C_WHITE),"DHs":(C_DH_HDR,C_WHITE),"Total":(C_NAVY,C_GOLD)}
    for i,(h,_) in enumerate(COLS,1):
        bg,fg=hdr_styles.get(h,(C_ACCENT,C_WHITE))
        c=ws.cell(row=4,column=i); hset(c,h,bg,fg,sz=10); tb(c)
    ws.row_dimensions[4].height=22
    total_am=total_dh=0
    for i,r in enumerate(rows):
        row_n=i+5; is_hall="hallway" in r["unit"].lower(); even=i%2==0
        n_am,n_dh=len(r["ams"]),len(r["dhs"]); total_am+=n_am; total_dh+=n_dh
        am_str="   |   ".join(r["ams"]) if r["ams"] else "—"
        dh_str="   |   ".join(r["dhs"]) if r["dhs"] else "—"
        bg_base="E8EEF4" if is_hall else (C_AM_L if even else C_WHITE)
        bg_am=C_AM_D if even else C_AM_L; bg_dh=C_DH_D if even else C_DH_L
        bg_tot="1A3F6F" if is_hall else C_NAVY
        cells=[(i+1,bg_base,Font(name="Arial",color="8FA8C8",size=9),True,False),
               (r["date"],bg_base,Font(name="Arial",color="1F3864",size=10,bold=not is_hall),True,False),
               (r["unit"],bg_base,Font(name="Arial",bold=not is_hall,italic=is_hall,color=C_NAVY,size=10),False,False),
               (am_str,bg_am,Font(name="Arial",color="1A3E5C",size=9),False,True),
               (n_am,bg_am,Font(name="Arial",bold=True,color=C_AM_HDR,size=11),True,False),
               (dh_str,bg_dh,Font(name="Arial",color="5D3A00",size=9),False,True),
               (n_dh,bg_dh,Font(name="Arial",bold=True,color=C_DH_HDR,size=11),True,False),
               (n_am+n_dh,bg_tot,Font(name="Arial",bold=True,color=C_GOLD,size=11),True,False)]
        for col,(val,bg,fnt,center,wrap) in enumerate(cells,1):
            c=ws.cell(row=row_n,column=col,value=val)
            c.font=fnt; c.fill=PatternFill("solid",start_color=bg)
            c.alignment=Alignment(horizontal="center" if center else "left",vertical="center",wrap_text=wrap,indent=0 if center else 1)
            nb(c)
        ws.row_dimensions[row_n].height=max(18,14*max(1,-(-max(n_am,1)//6),-(-max(n_dh,1)//3)))
    tr=len(rows)+5
    apt_c=sum(1 for r in rows if "unit" in r["unit"].lower())
    hall_c=sum(1 for r in rows if "hallway" in r["unit"].lower())
    ws.merge_cells(f"A{tr}:C{tr}"); hset(ws.cell(tr,1),f"TOTALS — {len(rows)} LOCATIONS","0A1F3D",sz=11,center=False); tb(ws.cell(tr,1))
    hset(ws.cell(tr,4),f"{total_am} placed",C_AM_HDR,sz=11); tb(ws.cell(tr,4))
    hset(ws.cell(tr,5),total_am,C_AM_HDR,C_GOLD,sz=13); tb(ws.cell(tr,5))
    hset(ws.cell(tr,6),f"{total_dh} placed",C_DH_HDR,sz=11); tb(ws.cell(tr,6))
    hset(ws.cell(tr,7),total_dh,C_DH_HDR,C_GOLD,sz=13); tb(ws.cell(tr,7))
    hset(ws.cell(tr,8),total_am+total_dh,"0A1F3D",C_GOLD,sz=14); tb(ws.cell(tr,8)); ws.row_dimensions[tr].height=26
    sr=tr+1; ws.merge_cells(f"A{sr}:H{sr}")
    c=ws.cell(sr,1,f"  {apt_c} Apartments  |  {hall_c} Hallways  |  {len(rows)} Locations  |  {total_am} Air Movers  |  {total_dh} Dehumidifiers  |  {total_am+total_dh} Total")
    c.font=Font(name="Arial",bold=True,color=C_NAVY,size=10); c.fill=PatternFill("solid",start_color="D6E8F7")
    c.alignment=Alignment(horizontal="center",vertical="center"); tb(c); ws.row_dimensions[sr].height=20
    buf=io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf,total_am,total_dh,apt_c,hall_c

if uploaded:
    st.markdown("### Process Chat")
    if st.button("Generate Equipment Log"):
        with st.spinner("Reading chat..."):
            raw=extract_txt(uploaded)
            messages=parse_messages(raw)
        st.info(f"Found {len(messages)} messages")
        with st.spinner("Analyzing with AI..."):
            rows=extract_equipment(messages, api_key)
            rows=sorted(rows,key=sort_key)
        if not rows:
            st.error("No equipment records found.")
        else:
            with st.spinner("Building Excel..."):
                buf,total_am,total_dh,apt_c,hall_c=build_excel(rows,project_name)
            st.success("Done!")
            c1,c2,c3,c4=st.columns(4)
            c1.metric("Apts",apt_c); c2.metric("Hallways",hall_c); c3.metric("Air Movers",total_am); c4.metric("Dehumidifiers",total_dh)
            fname=(project_name.replace(" ","_")[:30] if project_name else "Equipment_Log")+".xlsx"
            st.download_button("Download Excel",buf,fname,"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)

st.markdown("---")
st.markdown("<div style='text-align:center;color:#999;font-size:0.75rem'>BELFOR Property Restoration • Powered by Claude AI</div>",unsafe_allow_html=True)
