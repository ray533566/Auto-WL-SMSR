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

Every Template-*.xlsx next to this script (or in a templates/ folder) is used, e.g.
Template-50-400mA.xlsx, Template-80-400mA.xlsx, Template-100-400mA.xlsx. To add a new current
range just drop in a new Template-<label>.xlsx - no code change needed.

Usage
    python fill_chart_templates.py RAW.txt [RAW2.txt ...]
    python fill_chart_templates.py "D:\\OSA\\RawTxtFolder"            (all *.txt in folder)
    python fill_chart_templates.py RAW.txt --templates D:\\MyTemplates --out D:\\Output

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


def find_templates(*folders):
    """All Template-<label>.xlsx files (e.g. Template-100-400mA.xlsx) -> [(label, path)], sorted by current."""
    found = {}
    for d in folders:
        for p in glob.glob(os.path.join(d, "Template-*.xlsx")):
            name = os.path.basename(p)
            if name.startswith("~$"):  # Excel lock file
                continue
            label = name[len("Template-"):-len(".xlsx")]
            found.setdefault(label, p)
    def key(lab):
        m = re.match(r"(\d+)", lab)
        return (int(m.group(1)) if m else 10**9, lab)
    return [(lab, found[lab]) for lab in sorted(found, key=key)]

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
    new = re.sub(r'<autoFilter ref="A1:([A-Z]+)\d+"', lambda m: f'<autoFilter ref="A1:{m.group(1)}{last}"', new)
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
    """Recalculate cached values of the Mapping formulas (formulas themselves untouched).

    Each column is classified by its formula, so columns can be added or moved freely:
      MOD(INT((ROW()-2)/n),m)+1        -> channel number
      a+b*INT((ROW()-2)/n)             -> temperature
      a+b*MOD(ROW()-2,n)               -> current
      RawData!A2 (relative)            -> same-row RawData cell
      IF(RawData!$B$2="","",...)       -> fixed RawData cell
      IF(COUNTIFS(..)=0,"",SUMIFS(..)) -> lookup in RawData
      G3-G2                            -> difference to previous row
      RawData!A2&"-text" (header row)  -> text
    """
    cell_pat = re.compile(r'<c r="([A-Z]+)(\d+)"[^>]*?>(?:<f[^>]*>([^<]*)</f>)?')
    first_f = {}  # col -> (row, formula) : first explicit formula below the header
    hdr_f = {}    # col -> formula in row 1
    for col, r, f in cell_pat.findall(sheet_xml):
        if not f:
            continue
        r = int(r)
        if r == 1:
            hdr_f[col] = unescape(f)
        elif col not in first_f:
            first_f[col] = (r, unescape(f))

    kinds = {}
    for col, (r0, f) in first_f.items():
        m = re.fullmatch(r"MOD\(INT\(\(ROW\(\)-2\)/(\d+)\),(\d+)\)\+1", f)
        if m:
            kinds[col] = ("ch", int(m.group(1)), int(m.group(2)))
            continue
        m = re.fullmatch(r"(-?\d+)\+(\d+)\*INT\(\(ROW\(\)-2\)/(\d+)\)", f)
        if m:
            kinds[col] = ("block", int(m.group(1)), int(m.group(2)), int(m.group(3)))
            continue
        m = re.fullmatch(r"(-?\d+)\+(\d+)\*MOD\(ROW\(\)-2,(\d+)\)", f)
        if m:
            kinds[col] = ("cycle", int(m.group(1)), int(m.group(2)), int(m.group(3)))
            continue
        m = re.fullmatch(r"RawData!([A-Z]+)(\d+)", f)
        if m:
            kinds[col] = ("rel", m.group(1), int(m.group(2)) - r0)
            continue
        m = re.fullmatch(r'IF\(RawData!\$([A-Z]+)\$(\d+)="","",RawData!\$[A-Z]+\$\d+\)', f)
        if m:
            kinds[col] = ("fixed", m.group(1), int(m.group(2)))
            continue
        m = re.search(r"SUMIFS\(RawData!\$([A-Z]+):\$[A-Z]+,(.*)\)\)\s*$", f)
        if m:
            pairs = re.findall(r'RawData!\$([A-Z]+):\$[A-Z]+,("CH"&)?\$([A-Z]+)\d+', m.group(2))
            kinds[col] = ("lookup", m.group(1), pairs)
            continue
        m = re.fullmatch(rf"([A-Z]+){r0}-([A-Z]+){r0 - 1}", f)
        if m and m.group(1) == m.group(2):
            kinds[col] = ("diff", m.group(1))
            continue
        print(f"   ! Mapping column {col}: formula not recognised, cached value left as is ({f[:60]})")

    def compute(kind, r):
        k = r - 2
        t = kind[0]
        if t == "ch":
            return (k // kind[1]) % kind[2] + 1
        if t == "block":
            return kind[1] + kind[2] * (k // kind[3])
        if t == "cycle":
            return kind[1] + kind[2] * (k % kind[3])
        if t == "rel":
            v = grid.cell(kind[1], r + kind[2])
            return 0 if v in (None, "") else v
        if t == "fixed":
            v = grid.cell(kind[1], kind[2])
            return "" if v in (None, "") else v
        if t == "lookup":
            crit = []
            for raw_col, ch_prefix, map_col in kind[2]:
                v = vals.get((map_col, r))
                crit.append((raw_col, f"CH{v}" if ch_prefix else v))
            return grid.lookup(kind[1], crit)
        if t == "diff":
            a, b = vals.get((kind[1], r)), vals.get((kind[1], r - 1))
            if isinstance(a, (int, float)) and isinstance(b, (int, float)):
                return round(a - b, 10)
            if a is None or b is None:
                return None
            return "#VALUE!"
        return None

    vals = {}  # (col,row) -> value
    rownums = sorted({int(n) for n in re.findall(r'<c r="[A-Z]+(\d+)"', sheet_xml)})
    stage = {"ch": 0, "block": 0, "cycle": 0, "rel": 0, "fixed": 0, "lookup": 1, "diff": 2}
    order = sorted(kinds, key=lambda c: (stage[kinds[c][0]], col_index(c)))
    for r in rownums:
        if r < 2:
            continue
        for c in order:
            if r < first_f[c][0]:   # e.g. no ΔWL formula in the first data row
                continue
            vals[(c, r)] = compute(kinds[c], r)

    # header-row formulas such as RawData!A2&"-Case_Temp"
    for col, f in hdr_f.items():
        parts = re.findall(r'RawData!\$?([A-Z]+)\$?(\d+)|"([^"]*)"', f)
        if parts and re.fullmatch(r'(?:RawData!\$?[A-Z]+\$?\d+|"[^"]*")(?:&(?:RawData!\$?[A-Z]+\$?\d+|"[^"]*"))*', f):
            txt = ""
            for c_, r_, lit in parts:
                if c_:
                    v = grid.cell(c_, int(r_))
                    txt += "" if v is None else (fmt_num(v) if isinstance(v, (int, float)) else str(v))
                else:
                    txt += lit
            vals[(col, 1)] = txt

    pat = re.compile(r'<c r="([A-Z]+)(\d+)"([^>/]*)>(<f[^>]*/>|<f[^>]*>[^<]*</f>)(?:<v>[^<]*</v>)?</c>')

    def rep(mm):
        col, r, attrs, f = mm.group(1), int(mm.group(2)), mm.group(3), mm.group(4)
        if (col, r) not in vals:
            return mm.group(0)
        v = vals[(col, r)]
        attrs = re.sub(r'\s*t="[^"]*"', "", attrs)
        if v is None:
            return f'<c r="{col}{r}"{attrs}>{f}</c>'
        if isinstance(v, str) and v.startswith("#"):
            return f'<c r="{col}{r}"{attrs} t="e">{f}<v>{v}</v></c>'
        if isinstance(v, str):
            return f'<c r="{col}{r}"{attrs} t="str">{f}<v>{escape(v)}</v></c>'
        return f'<c r="{col}{r}"{attrs}>{f}<v>{fmt_num(v)}</v></c>'

    lookup_cols = [c for c in order if kinds[c][0] == "lookup"]
    return pat.sub(rep, sheet_xml), vals, (lookup_cols[0] if lookup_cols else None)


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
    new_map, vals, wl_col = update_mapping(z.read(map_path).decode("utf-8"), grid)

    wb = z.read("xl/workbook.xml").decode("utf-8")
    wb = re.sub(r"(RawData!\$A\$1:\$[A-Z]+\$)\d+", lambda m: f"{m.group(1)}{len(rows) + 1}", wb)
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

    g = [v for (c, r), v in vals.items() if c == wl_col and r >= 2]
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
    ap = argparse.ArgumentParser(description="Fill Template-*.xlsx chart templates from OSA raw TXT")
    ap.add_argument("inputs", nargs="*", help="TXT file(s) or folder(s); if omitted a file picker opens")
    ap.add_argument("--templates", default=None,
                    help="folder with Template-*.xlsx (default: this script's folder and its templates/ folder)")
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
    folders = [a.templates] if a.templates else [HERE, os.path.join(HERE, "templates")]
    templates = find_templates(*folders)
    if not templates:
        sys.exit(f"No Template-*.xlsx found in {folders}")
    print("Templates:", ", ".join(lab for lab, _ in templates))

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
