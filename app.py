"""Streamlit web app: OSA raw TXT -> 50-400mA / 80-400mA chart Excel files."""
import io
import os
import tempfile
import zipfile

import streamlit as st

from fill_chart_templates import fill_template

HERE = os.path.dirname(os.path.abspath(__file__))
def find_template(name):
    """Look in templates/ first, then next to app.py."""
    for p in (os.path.join(HERE, "templates", name), os.path.join(HERE, name)):
        if os.path.isfile(p):
            return p
    return os.path.join(HERE, name)


TEMPLATES = {
    "50-400mA": find_template("Template-50-400mA.xlsx"),
    "80-400mA": find_template("Template-80-400mA.xlsx"),
}

st.set_page_config(page_title="OSA Chart Auto-Fill", page_icon="📈", layout="centered")
st.title("OSA Chart Auto-Fill")
st.caption("Upload raw TXT files → get the 50-400mA and 80-400mA chart workbooks.")

with st.sidebar:
    st.header("Templates")
    which = st.multiselect("Generate", list(TEMPLATES), default=list(TEMPLATES))
    st.write("Optional: use your own template instead of the built-in one")
    custom = {lab: st.file_uploader(f"{lab} template (.xlsx)", type="xlsx", key=lab) for lab in TEMPLATES}

files = st.file_uploader("Raw data TXT file(s)", type="txt", accept_multiple_files=True)

if files and which and st.button("Generate Excel files", type="primary"):
    results, log = [], []
    with tempfile.TemporaryDirectory() as tmp:
        tpl = {}
        for lab in which:
            if custom[lab] is not None:
                p = os.path.join(tmp, f"custom-{lab}.xlsx")
                with open(p, "wb") as fh:
                    fh.write(custom[lab].getvalue())
                tpl[lab] = p
            else:
                tpl[lab] = TEMPLATES[lab]

        bar = st.progress(0.0)
        total = len(files) * len(which)
        done = 0
        for f in files:
            stem = os.path.splitext(f.name)[0]
            txt = os.path.join(tmp, f.name)
            with open(txt, "wb") as fh:
                fh.write(f.getvalue())
            for lab in which:
                out = os.path.join(tmp, f"{stem}-Chart-{lab}.xlsx")
                try:
                    nraw, nmap, blanks, sn = fill_template(tpl[lab], txt, out)
                    with open(out, "rb") as fh:
                        results.append((os.path.basename(out), fh.read()))
                    status = "OK" if blanks == 0 else f"⚠️ {blanks} Mapping rows without RawData match"
                    log.append({"File": f.name, "Template": lab, "SN": sn,
                                "RawData rows": nraw, "Mapping rows": nmap, "Status": status})
                except Exception as e:
                    log.append({"File": f.name, "Template": lab, "SN": "", "RawData rows": 0,
                                "Mapping rows": 0, "Status": f"❌ {e}"})
                done += 1
                bar.progress(done / total)
    st.session_state["results"], st.session_state["log"] = results, log

if st.session_state.get("log"):
    st.dataframe(st.session_state["log"], use_container_width=True, hide_index=True)
    results = st.session_state["results"]
    if len(results) == 1:
        name, data = results[0]
        st.download_button(f"⬇ {name}", data, file_name=name,
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    elif results:
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
            for name, data in results:
                z.writestr(name, data)
        st.download_button(f"⬇ Download all ({len(results)} files, zip)", buf.getvalue(),
                           file_name="OSA_Charts.zip", mime="application/zip", type="primary")
        with st.expander("Download individually"):
            for name, data in results:
                st.download_button(name, data, file_name=name, key=name,
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
