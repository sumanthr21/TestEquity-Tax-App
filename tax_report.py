"""Append the two Tax Specialists' EOD workbooks to the data sheets of 'Test Equity - OS - Tax Compliance Specialist'.
Each agent's workbook is a running tracker (one sheet per task, all dates). The app picks out ONE date, maps each
tracker sheet to its report sheet, and appends at XML level (charts/formulas survive). EOD sheets are STREAMED,
because one tracker sheet is ~87 MB."""
import re, zipfile, io, datetime as dt
from html import unescape
from xml.sax.saxutils import escape

# report sheet -> (names the EOD tracker sheet may have, EOD column -> report column, report columns to style)
SHIFT4 = lambda cols: {c: chr(ord(c) + 4) for c in cols}   # EOD B..  ->  report F..
SHEETS = {
    "Invalid Report": (["invalid report"], {"B": "F", "C": "L", "D": "S", "E": "N", "F": "V"}),
    "Tax Certificate Upload - T-hold": (["t-hold", "thold"], SHIFT4("BCDEFGH")),
    "Tax Certificate Upload - Email": (["tax certificate - email", "e-mail", "email"], SHIFT4("BCDEFGHI")),
    "Credit Memo": (["credit memo"], SHIFT4("BCD")),
    "Web UI - Tax cert": (["web ui - tax cert", "webui tax cert"], SHIFT4("BCDEFG")),
}
EPOCH = dt.datetime(1899, 12, 30)
CH = 8_000_000

def week_of(d):  # Mon-Sun weeks; Week 1 = week holding the month's first weekday (matches the report)
    first = d.replace(day=1)
    while first.weekday() > 4: first += dt.timedelta(days=1)
    mon = lambda x: x - dt.timedelta(days=x.weekday())
    return f"Week {(mon(d) - mon(first)).days // 7 + 1}"

def month_label(d, sheets):
    full = {4: "April", 6: "June", 7: "July"}.get(d.month, d.strftime("%b")); yy = "'" + d.strftime("%y")
    for c in (full + yy, d.strftime("%b") + yy, d.strftime("%B") + yy):
        if c in sheets: return c
    return full + yy

def serial(d): return (dt.datetime(d.year, d.month, d.day) - EPOCH).days

def _load(z):
    wb = z.read("xl/workbook.xml").decode(); rels = z.read("xl/_rels/workbook.xml.rels").decode()
    rid = {re.search(r'Id="([^"]+)"', t).group(1): re.search(r'Target="([^"]+)"', t).group(1) for t in re.findall(r"<Relationship [^>]*>", rels)}
    paths = {unescape(n): "xl/" + rid[r].lstrip("/").replace("xl/", "") for n, r in re.findall(r'<sheet [^>]*?name="([^"]+)"[^>]*?r:id="([^"]+)"', wb)}
    sst = [unescape(re.sub(r"<[^>]+>", "", s)) for s in re.findall(r"<si>(.*?)</si>", z.read("xl/sharedStrings.xml").decode(), flags=re.S)] if "xl/sharedStrings.xml" in z.namelist() else []
    return paths, sst

def _cells(body, sst):
    out = {}
    for col, attrs, inner in re.findall(r'<c r="([A-Z]+)\d+"([^>]*?)(?:/>|>(.*?)</c>)', body, flags=re.S):
        if not inner: continue
        v = re.search(r"<v>(.*?)</v>", inner, flags=re.S)
        if 't="s"' in attrs and v: out[col] = ("s", sst[int(v.group(1))])
        elif "inlineStr" in attrs:
            t = re.findall(r"<t[^>]*>(.*?)</t>", inner, flags=re.S); out[col] = ("s", unescape("".join(t)))
        elif 't="str"' in attrs and v: out[col] = ("s", unescape(v.group(1)))
        elif v and not re.search(r't="(b|e)"', attrs): out[col] = ("n", v.group(1))
    return out

def eod_rows(z, path, sst, date):
    """Stream one tracker sheet; yield {col: (kind, value)} for rows whose column A equals the date."""
    want = str(serial(date)).encode(); pat = re.compile(rb'<c r="A\d+"[^>]*><v>(\d+)(?:\.0+)?</v>'); buf = b""
    with z.open(path) as f:
        while True:
            chunk = f.read(CH)
            parts = (buf + chunk).split(b"</row>")
            buf = parts.pop() if chunk else b""
            for p in parts:
                m = pat.search(p)
                if m and m.group(1) == want: yield _cells(p.decode("utf-8", "ignore"), sst)
            if not chunk: break

def find_sheet(paths, names):
    for n, p in paths.items():
        if n.strip().lower() in names: return n, p
    return None, None

def guess_agent(z_paths, agents):
    """Mahalakshmi's tracker has an 'E-Mail' sheet; Abhinaya's has 'Tax Certificate - Email' / 'OT Tracker'."""
    low = {n.strip().lower() for n in z_paths}
    if "e-mail" in low: return next((a for a in agents if "mahalakshmi" in a.lower()), None)
    if "tax certificate - email" in low or "ot tracker" in low: return next((a for a in agents if "abhinaya" in a.lower()), None)
    return None

def learn(report_bytes):
    z = zipfile.ZipFile(io.BytesIO(report_bytes)); paths, sst = _load(z); done = set(); names = {}; styles = {}; lastrow = {}
    for sheet in SHEETS:
        x = z.read(paths[sheet]).decode()
        rows = re.findall(r'<row r="(\d+)"[^>]*>(.*?)</row>', x, flags=re.S)
        for rn, body in rows:
            c = _cells(body, sst)
            if "B" in c and c["B"][0] == "n" and "E" in c: done.add((sheet, int(float(c["B"][1])), c["E"][1]))
        for rn, body in rows[-300:]:
            c = _cells(body, sst)
            if "E" in c and c["E"][0] == "s": names[c["E"][1]] = names.get(c["E"][1], 0) + 1
        lastrow[sheet] = int(rows[-1][0]); styles[sheet] = dict(re.findall(r'<c r="([A-Z]+)\d+" s="(\d+)"', re.search(rf'<row r="{rows[-1][0]}".*?</row>', x, flags=re.S).group(0)))
    agents = [n for n, k in sorted(names.items(), key=lambda kv: -kv[1]) if n != "Name"][:2]   # exact spelling, incl. the leading non-breaking space in ' Abhinaya B'
    return {"agents": agents, "done": done, "styles": styles, "last": lastrow, "sheets": set(paths)}

def _cell(ref, v, style):
    s = f' s="{style}"' if style else ""
    if v is None: return ""
    kind, val = v
    if kind == "n": return f'<c r="{ref}"{s}><v>{val}</v></c>'
    return f'<c r="{ref}"{s} t="inlineStr"><is><t xml:space="preserve">{escape(val)}</t></is></c>'

def update_report(report_bytes, jobs, write_as=None):
    """jobs: list of (eod_file, agent, date). Returns (new_xlsx_bytes, log).  write_as: test hook (write rows under another date)."""
    ctx = learn(report_bytes); log = []; plan = {}   # sheet -> list of (agent, date, rows)
    zin = zipfile.ZipFile(io.BytesIO(report_bytes)); rpaths, _ = _load(zin)
    for f, agent, date in jobs:
        ez = zipfile.ZipFile(f); ep, esst = _load(ez); out_date = write_as or date
        for sheet, (names, cmap) in SHEETS.items():
            n, p = find_sheet(ep, names)
            if not p: log.append(f"WARNING {agent}: no tracker sheet for '{sheet}' in {getattr(f, 'name', f)}"); continue
            rows = list(eod_rows(ez, p, esst, date))
            if not rows: log.append(f"{agent}: {sheet}: no rows for {date:%d-%b-%Y}"); continue
            if (sheet, serial(out_date), agent) in ctx["done"]:
                raise ValueError(f"Already loaded: {sheet} / {agent} / {out_date:%d-%b-%Y}. Nothing was added.")
            plan.setdefault(sheet, []).append((agent, out_date, rows, cmap))
    if not plan: raise ValueError("No rows found for that date in the files. Check the date.")
    new = {}; latest = max(d for v in plan.values() for _, d, _, _ in v); total = 0
    for sheet, items in plan.items():
        p = rpaths[sheet]; x = zin.read(p).decode(); r = ctx["last"][sheet]; st = ctx["styles"][sheet]; xr = []
        for agent, d, rows, cmap in items:
            for row in rows:
                r += 1; cells = {"A": ("s", month_label(d, ctx["sheets"])), "B": ("n", str(serial(d))), "C": ("s", week_of(d)), "E": ("s", agent)}
                for ec, rc in cmap.items():
                    if ec in row: cells[rc] = row[ec]
                xr.append(f'<row r="{r}">' + "".join(_cell(f"{c}{r}", v, st.get(c)) for c, v in sorted(cells.items(), key=lambda kv: (len(kv[0]), kv[0]))) + "</row>")
            log.append(f"{agent}: {sheet}: +{len(rows)} rows for {d:%d-%b-%Y}"); total += len(rows)
        x = x.replace("</sheetData>", "".join(xr) + "</sheetData>", 1)
        x = re.sub(r'<dimension ref="A1:([A-Z]+)\d+"', lambda m: f'<dimension ref="A1:{m.group(1)}{r}"', x, count=1); new[p] = x
    m = month_label(latest, ctx["sheets"])
    if m in rpaths:
        x = zin.read(rpaths[m]).decode(); x, k = re.subn(r'(<c r="B6"[^>]*>)<v>\d+</v>', lambda q: f'{q.group(1)}<v>{serial(latest)}</v>', x, count=1)
        new[rpaths[m]] = x; log.append(f"{m}!B6 set to {latest:%d-%b-%Y}" if k else f"WARNING could not set {m}!B6")
    else: log.append(f"WARNING there is no '{m}' sheet yet - copy last month's sheet, rename it, then upload again")
    wbx = zin.read("xl/workbook.xml").decode()
    new["xl/workbook.xml"] = re.sub(r"<calcPr[^>]*/>", '<calcPr calcId="191029" fullCalcOnLoad="1"/>', wbx) if "<calcPr" in wbx else wbx.replace("</workbook>", '<calcPr fullCalcOnLoad="1"/></workbook>')
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zo:
        for it in zin.infolist():
            if it.filename != "xl/calcChain.xml": zo.writestr(it, new.get(it.filename, zin.read(it.filename)))
    log.insert(0, f"Total: +{total} rows")
    return out.getvalue(), log

def latest_date(f):
    """Most recent date in the tracker's small sheets (used as the default in the app)."""
    z = zipfile.ZipFile(f); paths, sst = _load(z); best = 0
    for sheet in ("Credit Memo", "Invalid Report", "T-Hold", "E-Mail", "Tax Certificate - Email"):
        n, p = find_sheet(paths, [sheet.lower()])
        if not p or z.getinfo(p).file_size > 20_000_000: continue
        for v in re.findall(rb'<c r="A\d+"[^>]*><v>(\d{5})(?:\.0+)?</v>', z.read(p)): best = max(best, int(v))
    return (EPOCH + dt.timedelta(days=best)).date() if best else None
