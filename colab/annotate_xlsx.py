"""Write an annotated copy of Tito's workbook from the CSVs in colab/.

Tito's xlsx stays untouched. This edits the zip directly (cells and sheets are only added), so formulas,
data validations and charts are preserved byte for byte. Notes go in a free column beside each input;
Survey / Summary / Sources become extra tabs. Standard library only.
"""
import csv, io, re, zipfile
from xml.sax.saxutils import escape

L = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
def _cidx(s):
    n = 0
    for ch in s: n = n * 26 + ord(ch) - 64
    return n

def _bump(st, tag, add, n_add):
    m = re.search(r'<%s count="(\d+)"' % tag, st); n = int(m.group(1))
    st = st.replace(m.group(0), '<%s count="%d"' % (tag, n + n_add), 1).replace("</%s>" % tag, add + "</%s>" % tag, 1)
    return st, n

def _read(path_or_text):
    return list(csv.DictReader(io.StringIO(path_or_text))) if isinstance(path_or_text, str) else list(path_or_text)

def fmt(v):
    try: x = float(v)
    except (TypeError, ValueError): return str(v)
    return ("%.4g" % x)

def note_text(r):
    rng = "range %s / %s / %s %s" % (fmt(r["low"]), fmt(r["central"]), fmt(r["high"]), r["unit"])
    dois = " ".join("https://doi.org/" + d for d in r["source_dois"].split()) if r["source_dois"] else ""
    t = "%s (low / Tito / high). Origin: %s. %s" % (rng, r["origin"], r["note"])
    if dois: t += " " + dois
    if r["conflict"]: t = r["conflict"] + " | " + t
    return t

def annotate(src, dst, ranges, survey=None, summary=None, sources=None):
    zin = zipfile.ZipFile(src)
    files = {n: zin.read(n) for n in zin.namelist()}; order = list(zin.namelist())
    st = files["xl/styles.xml"].decode("utf-8")
    st, f0 = _bump(st, "fonts", '<font><b/><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>'
                   '<font><sz val="11"/><name val="Calibri"/></font><font><i/><sz val="11"/><color rgb="FF375623"/><name val="Calibri"/></font>'
                   '<font><b/><i/><sz val="11"/><color rgb="FFC00000"/><name val="Calibri"/></font>', 5)
    st, l0 = _bump(st, "fills", '<fill><patternFill patternType="solid"><fgColor rgb="FF1F3864"/><bgColor indexed="64"/></patternFill></fill>'
                   '<fill><patternFill patternType="solid"><fgColor rgb="FFFFD966"/><bgColor indexed="64"/></patternFill></fill>'
                   '<fill><patternFill patternType="solid"><fgColor rgb="FFD9D9D9"/><bgColor indexed="64"/></patternFill></fill>', 3)
    wrap = '<alignment vertical="top" wrapText="1"/>'; top = '<alignment vertical="top"/>'
    def xf(font, fill, a=wrap): return '<xf numFmtId="0" fontId="%d" fillId="%d" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1">%s</xf>' % (font, fill, a)
    st, x0 = _bump(st, "cellXfs", xf(f0 + 2, 0) + xf(f0, 0) + xf(f0 + 1, l0) + xf(f0, l0 + 1) + xf(f0, l0 + 2) + xf(f0 + 3, 0, top) + xf(f0 + 4, 0, top), 7)
    S_TXT, S_BOLD, S_HDR, S_BAN, S_SEC, S_NOTE, S_CONF = range(x0, x0 + 7)
    files["xl/styles.xml"] = st.encode("utf-8")

    wbx = files["xl/workbook.xml"].decode("utf-8"); rel = files["xl/_rels/workbook.xml.rels"].decode("utf-8")
    rid_of = dict(re.findall(r'<sheet name="([^"]+)" sheetId="\d+" r:id="(rId\d+)"', wbx))
    tgt = dict(re.findall(r'<Relationship Id="(rId\d+)"[^>]*?Target="(worksheets/[^"]+)"', rel))
    tgt.update({a: b for b, a in re.findall(r'Target="(worksheets/[^"]+)"[^>]*?Id="(rId\d+)"', rel)})
    path_of = {n.replace("&amp;", "&"): "xl/" + tgt[r] for n, r in rid_of.items()}

    def split_cols(x, ci, width):
        m = re.search(r"<cols>(.*?)</cols>", x, re.S)
        if not m: return x
        out, done = [], False
        for c in re.findall(r"<col [^>]*/>", m.group(1)):
            a = dict(re.findall(r'(\w+)="([^"]*)"', c)); lo, hi = int(a["min"]), int(a["max"])
            if lo <= ci <= hi and not done:
                done = True
                def mk(l, h, w=None):
                    b = dict(a); b["min"], b["max"] = str(l), str(h)
                    if w: b["width"], b["customWidth"] = str(w), "1"
                    return "<col " + " ".join('%s="%s"' % kv for kv in b.items()) + "/>"
                if lo < ci: out.append(mk(lo, ci - 1))
                out.append(mk(ci, ci, width))
                if ci < hi: out.append(mk(ci + 1, hi))
            else: out.append(c)
        if not done: out.append('<col min="%d" max="%d" width="%s" customWidth="1"/>' % (ci, ci, width))
        return x[:m.start(1)] + "".join(out) + x[m.end(1):]

    by_sheet = {}
    for r in _read(ranges): by_sheet.setdefault(r["sheet"], []).append(r)
    for sheet, rs in by_sheet.items():
        p = path_of[sheet]; x = files[p].decode("utf-8")
        dm = re.search(r'<dimension ref="[A-Z]+\d+:([A-Z]+)(\d+)"/>', x)
        colL = L[_cidx(dm.group(1))] if dm else "J"
        ci = _cidx(colL)
        for r in sorted(rs, key=lambda r: int(re.sub(r"\D", "", r["cell"]))):
            row = int(re.sub(r"\D", "", r["cell"])); text = note_text(r)
            newc = '<c r="%s%d" s="%d" t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>' % (colL, row, S_CONF if r["conflict"] else S_NOTE, escape(text))
            m = re.search(r'<row r="%d"(?: [^>]*?)?(/>|>(.*?)</row>)' % row, x, re.S)
            if not m: continue
            full = m.group(0)
            if m.group(1) == "/>": rowx = full[:-2] + ">" + newc + "</row>"
            else:
                inner = m.group(2); cells = list(re.finditer(r'<c r="([A-Z]+)%d"' % row, inner))
                if any(_cidx(c.group(1)) == ci for c in cells): continue
                pos = next((c.start() for c in cells if _cidx(c.group(1)) > ci), len(inner))
                rowx = full[:m.start(2) - m.start(0)] + inner[:pos] + newc + inner[pos:] + "</row>"
            rowx = re.sub(r'^(<row r="%d"[^>]*?) spans="[^"]*"' % row, r"\1", rowx)
            x = x[:m.start()] + rowx + x[m.end():]
        x = split_cols(x, ci, 120)
        if dm: x = x.replace(dm.group(0), dm.group(0).replace(dm.group(1) + dm.group(2), colL + dm.group(2)) if _cidx(dm.group(1)) < ci else dm.group(0))
        files[p] = x.encode("utf-8")

    def cell(r, c, v, s):
        ref = "%s%d" % (L[c], r)
        if v is None or v == "": return '<c r="%s" s="%d"/>' % (ref, s)
        return '<c r="%s" s="%d" t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>' % (ref, s, escape(str(v)))
    def table(banner, rows, widths):
        hdr = list(rows[0].keys()); n = len(hdr); out = ['<row r="1" ht="48" customHeight="1">%s</row>' % cell(1, 0, banner, S_BAN)]
        out.append('<row r="3">%s</row>' % "".join(cell(3, i, h, S_HDR) for i, h in enumerate(hdr)))
        for k, r in enumerate(rows):
            out.append('<row r="%d">%s</row>' % (k + 4, "".join(cell(k + 4, i, r[h], S_TXT) for i, h in enumerate(hdr))))
        cols = "".join('<col min="%d" max="%d" width="%s" customWidth="1"/>' % (i + 1, i + 1, widths[i] if i < len(widths) else 30) for i in range(n))
        return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetPr><tabColor rgb="FFFFC000"/></sheetPr>'
                '<sheetViews><sheetView workbookViewId="0"><pane ySplit="3" topLeftCell="A4" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight="15"/>'
                '<cols>%s</cols><sheetData>%s</sheetData><mergeCells count="1"><mergeCell ref="A1:%s1"/></mergeCells></worksheet>' % (cols, "".join(out), L[n - 1]))
    BAN = "ADDED by precis/Reto - NOT part of Tito's draft. Generated from the CSV files in the colab/ folder; Tito's values and formulas are unchanged."
    new = [(n, rows, w) for n, rows, w in (("Survey (added)", survey, [34, 38, 16, 22, 30, 30, 44, 70, 16]), ("Summary (added)", summary, [34, 40, 24, 34, 52, 30, 52, 16]),
                                       ("Sources (added)", sources, [70, 60, 70])) if rows]
    base_sid = 100; app = files["docProps/app.xml"].decode("utf-8"); ct = files["[Content_Types].xml"].decode("utf-8")
    for k, (nm, rows, w) in enumerate(new):
        rows = _read(rows); part = "xl/worksheets/sheet%d.xml" % (100 + k); rid = "rId%d" % (900 + k)
        wbx = wbx.replace("</sheets>", '<sheet name="%s" sheetId="%d" r:id="%s"/></sheets>' % (nm, base_sid + k, rid))
        rel = rel.replace("</Relationships>", '<Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet%d.xml"/></Relationships>' % (rid, 100 + k))
        ct = ct.replace("</Types>", '<Override PartName="/%s" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>' % part)
        app = app.replace("</vt:vector></TitlesOfParts>", "<vt:lpstr>%s</vt:lpstr></vt:vector></TitlesOfParts>" % nm)
        files[part] = table(BAN, rows, w).encode("utf-8"); order.append(part)
    if new:
        m = re.search(r'<vt:i4>(\d+)</vt:i4>', app); n0 = int(m.group(1))
        app = app.replace(m.group(0), "<vt:i4>%d</vt:i4>" % (n0 + len(new)), 1).replace('<vt:vector size="%d" baseType="lpstr">' % n0, '<vt:vector size="%d" baseType="lpstr">' % (n0 + len(new)))
    files["xl/workbook.xml"] = wbx.encode("utf-8"); files["xl/_rels/workbook.xml.rels"] = rel.encode("utf-8")
    files["[Content_Types].xml"] = ct.encode("utf-8"); files["docProps/app.xml"] = app.encode("utf-8")
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as z:
        for n in order: z.writestr(n, files[n])
    return dst

def set_values(src, dst, edits):
    """Write numeric input values into a copy of the workbook (cells keep their style; Excel recalculates on open)."""
    zin = zipfile.ZipFile(src); files = {n: zin.read(n) for n in zin.namelist()}; order = list(zin.namelist())
    wbx = files["xl/workbook.xml"].decode("utf-8"); rel = files["xl/_rels/workbook.xml.rels"].decode("utf-8")
    rid_of = dict(re.findall(r'<sheet name="([^"]+)" sheetId="\d+" r:id="(rId\d+)"', wbx))
    tgt = dict(re.findall(r'<Relationship Id="(rId\d+)"[^>]*?Target="(worksheets/[^"]+)"', rel))
    tgt.update({a: b for b, a in re.findall(r'Target="(worksheets/[^"]+)"[^>]*?Id="(rId\d+)"', rel)})
    for (sheet, cell), v in edits.items():
        p = "xl/" + tgt[rid_of[sheet]]; x = files[p].decode("utf-8")
        pat = re.compile(r'<c r="%s"([^>]*?)>\s*<v>[^<]*</v>\s*</c>' % cell)
        m = pat.search(x)
        if not m: raise KeyError((sheet, cell))
        attrs = re.sub(r' t="[^"]*"', "", m.group(1))
        x = x[:m.start()] + '<c r="%s"%s><v>%r</v></c>' % (cell, attrs, float(v)) + x[m.end():]
        files[p] = x.encode("utf-8")
    wbx = re.sub(r"<calcPr([^>]*?)/>", lambda m: "<calcPr%s fullCalcOnLoad=\"1\"/>" % m.group(1) if "fullCalcOnLoad" not in m.group(1) else m.group(0), wbx)
    files["xl/workbook.xml"] = wbx.encode("utf-8")
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as z:
        for n in order: z.writestr(n, files[n])
    return dst

if __name__ == "__main__":
    import sys
    d = sys.argv[3]
    rd = lambda f: open(d + "/" + f, encoding="utf-8").read()
    annotate(sys.argv[1], sys.argv[2], rd("ranges.csv"), rd("survey.csv"), rd("summary.csv"), rd("sources.csv"))
