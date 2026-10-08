#!/usr/bin/env python3
"""
fill_chart_templates.py
-----------------------
Fill the 50-400mA and 80-400mA "Chart" Excel templates from an OSA raw-data TXT
(tab-separated, e.g. 48226H40040_A2026100620533560886.txt).

For every TXT it creates:
    <txt name>-Chart-50-400mA.xlsx
    <txt name>-Chart-80-400mA.xlsx

What it does to each template copy
  * RawData sheet : replaced with the TXT data, laid out like the Power Query output
                    (SN, Time, Program, Station="Temp", Column1=<temp number>, Channel, ...)
  * Mapping sheet : formulas are kept untouched; their cached values are recalculated
                    so the file is correct even before Excel recalculates
  * Charts        : the old SN in chart titles is replaced by the new SN, chart caches refreshed
  * Workbook      : set to fully recalculate when opened in Excel
Everything else (charts, styles, column widths, filters) is preserved byte-for-byte.

Usage
    python fill_chart_templates.py RAW.txt [RAW2.txt ...]
    python fill_chart_templates.py "D:\\OSA\\RawTxtFolder"            (all *.txt in folder)
    python fill_chart_templates.py RAW.txt --t50 T50.xlsx --t80 T80.xlsx --out D:\\Output

Default template paths: Template-50-400mA.xlsx / Template-80-400mA.xlsx next to this script.
Requires only the Python standard library (3.8+).
"""
import argparse
import csv
import glob
import os
import re
import sys
import zipfile
from xml.sax.saxutils import escape

HERE = os.path.dirname(os.path.abspath(__file__))
DEFAULT_T50 = os.path.join(HERE, "Template-50-400mA.xlsx")
DEFAULT_T80 = os.path.join(HERE, "Template-80-400mA.xlsx")

NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


# ----------------------------------------------------------------- helpers
def col_letter(n):  # 1 -> A
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def col_index(letters):  # A -> 1
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch) - 64
    return n


def to_num(s):
    """'1260.00' -> 1260 ; '-3.83' -> -3.83 ; non-number -> None"""
    try:
        f = float(s)
    except (TypeError, ValueError):
        return None
    if f != f or f in (float("inf"), float("-inf")):  # NaN / +-Inf -> keep as text
        return None
    return int(f) if f.is_integer() else f


def fmt_num(v):
    if isinstance(v, int):
        return str(v)
    return repr(round(v, 10))


# ----------------------------------------------------------------- read TXT
def read_txt(path):
    with open(path, "r", encoding="utf-8-sig", errors="replace", newline="") as fh:
        rows = list(csv.reader(fh, delimiter="\t"))
    rows = [r for r in rows if any(c.strip() for c in r)]
    header = [h.strip() for h in rows[0]]
    data = [dict(zip(header, [c.strip() for c in r])) for r in rows[1:]]
    for need in ("SN", "Time", "Station", "Channel", "WL(nm)", "SMSR(dB)", "Set_Bias(mA)"):
        if need not in header:
            raise ValueError(f"{os.path.basename(path)}: column '{need}' not found in TXT header")
    return header, data


def split_station(st):
    """'Temp_-5' -> ('Temp', -5)"""
    if "_" in st:
        a, b = st.rsplit("_", 1)
        n = to_num(b)
        if n is not None:
            return a, n
    return st, None


# ----------------------------------------------------------------- xlsx parts
def sheet_paths(z):
    wb = z.read("xl/workbook.xml").decode("utf-8")
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    rid2target = dict(re.findall(r'Id="([^"]+)"[^>]*Target="([^"]+)"', rels))
    rid2target.update({k: v for v, k in re.findall(r'Target="([^"]+)"[^>]*Id="([^"]+)"', rels)})
    out = {}
    for name, rid in re.findall(r'<sheet [^>]*name="([^"]+)"[^>]*r:id="([^"]+)"', wb):
        t = rid2target[rid].lstrip("/")
        out[name] = t if t.startswith("xl/") else "xl/" + t
    return out


def shared_strings(z):
    if "xl/sharedStrings.xml" not in z.namelist():
        return []
    x = z.read("xl/sharedStrings.xml").decode("utf-8")
    out = []
    for si in re.findall(r"<si>(.*?)</si>", x, flags=re.S):
        txt = "".join(re.findall(r"<t[^>]*>(.*?)</t>", si, flags=re.S))
        out.append(txt.replace("&amp;", "&").replace("&lt;", "<").replace("&gt;", ">"))
    return out


def unescape(s):
    return s.replace("&lt;", "<").replace("&gt;", ">").replace("&quot;", '"').replace("&amp;", "&")


# ----------------------------------------------------------------- core
def build_rawdata_rows(template_headers, txt_data):
    """Return list of row-value lists laid out in template column order."""
    rows = []
    for d in txt_data:
        prefix, tnum = split_station(d.get("Station", ""))
        if tnum is None:
            tnum = to_num(d.get("Set_Temp"))
        vals = []
        for h in template_headers:
            if h == "Station":
                vals.append(prefix)
            elif h == "Column1":
                vals.append(tnum)
            else:
                raw = d.get(h, "")
                if raw.lower() in ("inf", "-inf", "+inf", "nan"):  # Power Query turns these into blanks
                    raw = ""
                n = to_num(raw)
                vals.append(n if n is not None else raw)
        rows.append(vals)
    return rows


def rawdata_xml(sheet_xml, header_row_xml, rows):
    ncol = len(rows[0]) if rows else 1
    parts = [header_row_xml]
    for i, vals in enumerate(rows, start=2):
        cells = []
        for j, v in enumerate(vals, start=1):
            ref = f"{col_letter(j)}{i}"
            if v is None or v == "":
                continue
            if isinstance(v, (int, float)):
                cells.append(f'<c r="{ref}"><v>{fmt_num(v)}</v></c>')
            else:
                cells.append(f'<c r="{ref}" t="inlineStr"><is><t>{escape(v)}</t></is></c>')
        parts.append(f'<row r="{i}" spans="1:{ncol}">{"".join(cells)}</row>')
    new = re.sub(r"<sheetData>.*</sheetData>|<sheetData/>",
                 lambda m: "<sheetData>" + "".join(parts) + "</sheetData>", sheet_xml, flags=re.S)
    last = len(rows) + 1
    new = re.sub(r'<dimension ref="[^"]*"/>', f'<dimension ref="A1:{col_letter(ncol)}{last}"/>', new)
    return new


class Grid:
    """Values of RawData by column letter for quick lookups."""

    def __init__(self, headers, rows):
        self.rows = rows
        self.col = {col_letter(i + 1): i for i in range(len(headers))}

    def cell(self, letter, rownum):  # rownum is the Excel row (2 = first data row)
        i = rownum - 2
        if 0 <= i < len(self.rows):
            return self.rows[i][self.col[letter]]
        return None

    @staticmethod
    def _key(v):
        if isinstance(v, (int, float)):
            return float(v)
        n = to_num(v) if isinstance(v, str) else None
        return float(n) if n is not None else str(v).strip().lower()

    def lookup(self, out_col, crit):
        """SUMIFS with COUNTIFS guard (blank if no match). crit: list of (col_letter, value)"""
        cols = tuple(self.col[c] for c, _ in crit)
        o = self.col[out_col]
        cache = self.__dict__.setdefault("_idx", {})
        if (cols, o) not in cache:
            idx = {}
            for r in self.rows:
                k = tuple(self._key(r[i]) for i in cols)
                x = r[o]
                hit, tot = idx.get(k, (0, 0.0))
                idx[k] = (hit + 1, tot + (x if isinstance(x, (int, float)) else 0))
            cache[(cols, o)] = idx
        hit, tot = cache[(cols, o)].get(tuple(self._key(v) for _, v in crit), (0, 0.0))
        return round(tot, 10) if hit else ""


def update_mapping(sheet_xml, grid):
    """Recalculate cached values of the Mapping formulas (formulas themselves untouched)."""
    # read the generator formulas from row 2 (shared or plain)
    f2 = {c: unescape(f) for c, f in re.findall(r'<c r="([A-Z]+)2"[^>]*><f[^>]*>([^<]*)</f>', sheet_xml)}
    m = re.search(r"(-?\d+)\+(\d+)\*MOD\(ROW\(\)-2,(\d+)\)", f2.get("F", ""))
    if not m:
        raise ValueError("Mapping!F2 current formula not recognised")
    i0, istep, per = int(m.group(1)), int(m.group(2)), int(m.group(3))
    mt = re.search(r"(-?\d+)\+(\d+)\*INT\(\(ROW\(\)-2\)/(\d+)\)", f2.get("E", ""))
    t0, tstep, tper = int(mt.group(1)), int(mt.group(2)), int(mt.group(3))
    mc = re.search(r"MOD\(INT\(\(ROW\(\)-2\)/(\d+)\),(\d+)\)\+1", f2.get("D", ""))
    cper, cmod = int(mc.group(1)), int(mc.group(2))

    def lookup_cols(formula):
        # SUMIFS(RawData!$P:$P,RawData!$E:$E,$E2,RawData!$F:$F,"CH"&$D2,RawData!$AF:$AF,$F2)
        s = re.search(r"SUMIFS\(RawData!\$([A-Z]+):\$[A-Z]+,(.*)\)\)\s*$", formula)
        out = s.group(1)
        pairs = re.findall(r'RawData!\$([A-Z]+):\$[A-Z]+,("CH"&)?\$([A-Z]+)\d+', s.group(2))
        return out, pairs

    lk = {c: lookup_cols(f2[c]) for c in ("G", "H") if c in f2 and "SUMIFS" in f2[c]}

    vals = {}  # (col,row) -> value

    def compute(col, r):
        k = r - 2
        if col == "D":
            return (k // cper) % cmod + 1
        if col == "E":
            return t0 + tstep * (k // tper)
        if col == "F":
            return i0 + istep * (k % per)
        if col == "C":
            v = grid.cell("B", 2)
            return "" if v in (None, "") else v
        if col == "B":
            v = grid.cell("A", r)
            return 0 if v in (None, "") else v
        if col in lk:
            out, pairs = lk[col]
            crit = []
            for raw_col, ch_prefix, map_col in pairs:
                v = vals[(map_col, r)]
                crit.append((raw_col, f"CH{v}" if ch_prefix else v))
            return grid.lookup(out, crit)
        if col == "I":
            a, b = vals.get(("G", r)), vals.get(("G", r - 1))
            if isinstance(a, (int, float)) and isinstance(b, (int, float)):
                return round(a - b, 10)
            if a is None or b is None:
                return None
            return "#VALUE!"
        return None

    order = ["B", "C", "D", "E", "F", "G", "H", "I"]
    rownums = sorted({int(n) for n in re.findall(r'<c r="[A-Z]+(\d+)"', sheet_xml)})
    for r in rownums:
        if r < 2:
            continue
        for c in order:
            vals[(c, r)] = compute(c, r)

    pat = re.compile(r'<c r="([A-Z]+)(\d+)"([^>/]*)>(<f[^>]*/>|<f[^>]*>[^<]*</f>)(?:<v>[^<]*</v>)?</c>')

    def rep(mm):
        col, r, attrs, f = mm.group(1), int(mm.group(2)), mm.group(3), mm.group(4)
        v = vals.get((col, r))
        attrs = re.sub(r'\s*t="[^"]*"', "", attrs)
        if v is None:
            return f'<c r="{col}{r}"{attrs}>{f}</c>'
        if isinstance(v, str) and v.startswith("#"):
            return f'<c r="{col}{r}"{attrs} t="e">{f}<v>{v}</v></c>'
        if isinstance(v, str):
            return f'<c r="{col}{r}"{attrs} t="str">{f}<v>{escape(v)}</v></c>'
        return f'<c r="{col}{r}"{attrs}>{f}<v>{fmt_num(v)}</v></c>'

    return pat.sub(rep, sheet_xml), vals


def update_chart(xml, old_sn, new_sn, vals):
    if old_sn and new_sn and old_sn != new_sn:
        xml = xml.replace(old_sn, new_sn)

    def rep(m):
        col, r1, r2 = m.group(2), int(m.group(3)), int(m.group(4))
        body = m.group(5)
        fc = re.search(r"<c:formatCode>.*?</c:formatCode>", body)
        pts = []
        for i, r in enumerate(range(r1, r2 + 1)):
            v = vals.get((col, r))
            if isinstance(v, (int, float)):
                pts.append(f'<c:pt idx="{i}"><c:v>{fmt_num(v)}</c:v></c:pt>')
        cache = (fc.group(0) if fc else "") + f'<c:ptCount val="{r2 - r1 + 1}"/>' + "".join(pts)
        return f"{m.group(1)}<c:numCache>{cache}</c:numCache>"

    xml = re.sub(r"(<c:f>Mapping!\$([A-Z]+)\$(\d+):\$[A-Z]+\$(\d+)</c:f>)<c:numCache>(.*?)</c:numCache>",
                 rep, xml, flags=re.S)
    return xml


def fill_template(template, txt_path, out_path):
    header, data = read_txt(txt_path)
    z = zipfile.ZipFile(template)
    sp = sheet_paths(z)
    ss = shared_strings(z)
    raw_path, map_path = sp["RawData"], sp["Mapping"]

    raw_xml = z.read(raw_path).decode("utf-8")
    hdr_row = re.search(r'<row r="1"[^>]*>.*?</row>', raw_xml, flags=re.S).group(0)
    tmpl_headers = []
    for c in re.findall(r"<c [^>]*?(?:/>|>.*?</c>)", hdr_row, flags=re.S):
        v = re.search(r"<v>([^<]*)</v>", c)
        t = re.search(r't="(\w+)"', c)
        if t and t.group(1) == "s":
            tmpl_headers.append(ss[int(v.group(1))])
        else:
            it = re.search(r"<t[^>]*>([^<]*)</t>", c)
            tmpl_headers.append(unescape(it.group(1)) if it else (v.group(1) if v else ""))
    missing = [h for h in tmpl_headers if h not in header and h not in ("Station", "Column1")]
    if missing:
        print(f"   ! TXT has no column(s) {missing} - left blank")

    # old SN (for chart titles) = template RawData!A2
    old_sn = None
    a2 = re.search(r'<c r="A2"[^>]*>.*?</c>', raw_xml, flags=re.S)
    if a2:
        v = re.search(r"<v>([^<]*)</v>", a2.group(0))
        if v and 't="s"' in a2.group(0):
            old_sn = ss[int(v.group(1))]
        else:
            it = re.search(r"<t[^>]*>([^<]*)</t>", a2.group(0))
            old_sn = it.group(1) if it else (v.group(1) if v else None)

    rows = build_rawdata_rows(tmpl_headers, data)
    new_sn = rows[0][tmpl_headers.index("SN")] if "SN" in tmpl_headers and rows else None
    new_raw = rawdata_xml(raw_xml, hdr_row, rows)

    grid = Grid(tmpl_headers, rows)
    new_map, vals = update_mapping(z.read(map_path).decode("utf-8"), grid)

    wb = z.read("xl/workbook.xml").decode("utf-8")
    if "fullCalcOnLoad" not in wb:
        wb = re.sub(r"<calcPr([^>]*?)/>", r'<calcPr\1 fullCalcOnLoad="1"/>', wb)

    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    tmp = out_path + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zo:
        for it in z.infolist():
            name = it.filename
            if name == raw_path:
                data_b = new_raw.encode("utf-8")
            elif name == map_path:
                data_b = new_map.encode("utf-8")
            elif name == "xl/workbook.xml":
                data_b = wb.encode("utf-8")
            elif name.startswith("xl/charts/chart") and name.endswith(".xml"):
                data_b = update_chart(z.read(name).decode("utf-8"), old_sn, new_sn, vals).encode("utf-8")
            else:
                data_b = z.read(name)
            zo.writestr(it, data_b, compress_type=zipfile.ZIP_DEFLATED)
    z.close()
    os.replace(tmp, out_path)

    g = [v for (c, r), v in vals.items() if c == "G"]
    blanks = sum(1 for v in g if v == "")
    return len(rows), len(g), blanks, new_sn


def pick_txt_files():
    """Open a Windows file dialog to choose one or more raw TXT files."""
    try:
        import tkinter as tk
        from tkinter import filedialog
    except Exception:
        print("No TXT given and tkinter is unavailable. Usage: fill_chart_templates.py RAW.txt")
        return []
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    files = filedialog.askopenfilenames(title="Select OSA raw TXT file(s)",
                                        filetypes=[("Raw data TXT", "*.txt"), ("All files", "*.*")])
    root.destroy()
    return list(files)


def main():
    ap = argparse.ArgumentParser(description="Fill 50-400mA / 80-400mA chart templates from OSA raw TXT")
    ap.add_argument("inputs", nargs="*", help="TXT file(s) or folder(s); if omitted a file picker opens")
    ap.add_argument("--t50", default=DEFAULT_T50, help="50-400mA template xlsx")
    ap.add_argument("--t80", default=DEFAULT_T80, help="80-400mA template xlsx")
    ap.add_argument("--out", default=None, help="output folder (default: same folder as each TXT)")
    a = ap.parse_args()

    inputs = list(a.inputs)
    if not inputs:  # double-clicked: let the user pick TXT files
        inputs = pick_txt_files()
        if not inputs:
            sys.exit("No TXT file selected.")
    txts = []
    for p in inputs:
        txts += sorted(glob.glob(os.path.join(p, "*.txt"))) if os.path.isdir(p) else [p]
    if not txts:
        sys.exit("No TXT files found.")
    templates = [("50-400mA", a.t50), ("80-400mA", a.t80)]
    for label, t in templates:
        if not os.path.isfile(t):
            sys.exit(f"Template not found: {t}")

    for txt in txts:
        stem = os.path.splitext(os.path.basename(txt))[0]
        outdir = a.out or os.path.dirname(os.path.abspath(txt))
        print(f"\n{os.path.basename(txt)}")
        for label, t in templates:
            out = os.path.join(outdir, f"{stem}-Chart-{label}.xlsx")
            try:
                nraw, nmap, blanks, sn = fill_template(t, txt, out)
                flag = "OK" if blanks == 0 else f"WARNING: {blanks} Mapping rows found no RawData match"
                print(f"  [{label}] SN={sn}  RawData rows={nraw}  Mapping rows={nmap}  -> {flag}")
                print(f"           saved: {out}")
            except Exception as e:  # keep going with the next file
                print(f"  [{label}] FAILED: {e}")


if __name__ == "__main__":
    main()
