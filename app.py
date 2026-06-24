# app.py  -  JRC Truck Plan Builder (single Flask app for Vercel)
import io, re, json, base64, tempfile, os
from collections import Counter
from http.server import BaseHTTPRequestHandler
import openpyxl
from openpyxl.styles import PatternFill, Alignment, Font
from openpyxl.utils import get_column_letter

# ===================== formatter =====================
NEW_COLUMNS = ["JRC EQUIPMENT", "JRC TRUCK", "JRC RATE", "L FT", "W FT", "H FT", "POUNDS"]
ID_HEADER = "ID"

# Header fills: the three JRC-decision columns are green; the four converted-
# measurement columns are yellow. This makes the human-entered/decided section and
# the machine-converted section visually distinct at a glance.
GREEN_HEADERS = {"JRC EQUIPMENT", "JRC TRUCK", "JRC RATE"}
YELLOW_HEADERS = {"L FT", "W FT", "H FT", "POUNDS"}

GREEN = PatternFill(start_color="C1F0C8", end_color="C1F0C8", fill_type="solid")
YELLOW = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
# Over-dimension data cells are BLUE (a load over 12 ft tall, over 8'6" wide, or
# over 47,500 lb); an extreme over-width load (> 16 ft) is RED instead.
BLUE = PatternFill(start_color="83CCEB", end_color="83CCEB", fill_type="solid")
RED = PatternFill(start_color="FF2F2F", end_color="FF2F2F", fill_type="solid")
# Extreme over-width (> 16 ft) cells: YELLOW fill with RED text (replaces the old
# solid-red fill). Header text colours match the header fill (Excel's Good/Neutral
# styles): green headers -> dark-green text, yellow headers -> brown text.
YELLOW_HL = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
RED_FONT = Font(color="FF0000")
GREEN_FONT = Font(color="006100")
YELLOW_FONT = Font(color="9C5700")

# Practical per-truck cargo payload ceiling (lb). A load at or under this can ride
# a truck; the consolidation grouping must keep each truck's combined cargo weight
# within it, and any single item over it is highlighted blue. This is the federal
# 80,000 lb gross less a typical tractor + open-deck trailer tare.
PAYLOAD_CEILING_LB = 47500

LEN_TO_FT = {
    "mm": 0.00328084, "cm": 0.0328084, "m": 3.2808399, "meter": 3.2808399,
    "metre": 3.2808399, "in": 1/12.0, "inch": 1/12.0, "inches": 1/12.0,
    "\"": 1/12.0, "ft": 1.0, "foot": 1.0, "feet": 1.0, "'": 1.0,
    "yd": 3.0, "yard": 3.0,
}
WT_TO_LB = {
    "g": 0.00220462, "gram": 0.00220462, "kg": 2.20462262, "kilogram": 2.20462262,
    "kgs": 2.20462262, "t": 2204.62262, "tonne": 2204.62262, "tonnes": 2204.62262,
    "metric ton": 2204.62262, "lb": 1.0, "lbs": 1.0, "pound": 1.0, "pounds": 1.0,
    "ton": 2000.0, "us ton": 2000.0, "short ton": 2000.0,
}


def to_ft(meas):
    u = str(meas["unit"]).strip().lower()
    if u not in LEN_TO_FT:
        raise ValueError(f"Unknown length unit {meas['unit']!r}. Flag this; do not guess.")
    return round(float(meas["value"]) * LEN_TO_FT[u], 1)   # one decimal place


def to_lb(meas):
    u = str(meas["unit"]).strip().lower()
    if u not in WT_TO_LB:
        raise ValueError(f"Unknown weight unit {meas['unit']!r}. Flag this; do not guess.")
    return int(round(float(meas["value"]) * WT_TO_LB[u]))   # nearest whole pound


def renumber_by_rule(rows):
    """Number trucks by equipment class: double drops (incl. extended) first,
    then stepdecks, then flatbeds. Within double drops / stepdecks sort by tallest
    load first; within flatbeds sort by longest length first. Ties -> first row."""
    def rank(e):
        e=(e or "").upper()
        if "DOUBLE DROP" in e or "MINIDECK" in e or "MINI DECK" in e: return 0
        if "STEPDECK" in e: return 1
        return 2                            # FLATBED
    groups={}
    for r in rows:
        gid=r["truck"]
        g=groups.setdefault(gid,{"max_h":float("-inf"),"max_l":float("-inf"),
                                 "first_row":r["excel_row"],"equip":r.get("equipment")})
        g["max_h"]=max(g["max_h"],r["h_ft"])
        g["max_l"]=max(g["max_l"],r["l_ft"])
        g["first_row"]=min(g["first_row"],r["excel_row"])
    def sortkey(kv):
        g=kv[1]; rk=rank(g["equip"])
        within = -g["max_l"] if rk==2 else -g["max_h"]   # flatbed by length, else height
        return (rk, within, g["first_row"])
    ordered=sorted(groups.items(), key=sortkey)
    return {gid:i for i,(gid,_) in enumerate(ordered, start=1)}


def renumber_by_height(rows_with_h):
    """Map each model 'truck' id to a final number so the truck carrying the
    tallest load is 1, the next tallest group is 2, etc. Items keep their
    grouping; only the labels change. Ties broken by first appearance (row)."""
    groups = {}
    for r in rows_with_h:
        gid = r["truck"]
        g = groups.setdefault(gid, {"max_h": float("-inf"), "first_row": r["excel_row"]})
        g["max_h"] = max(g["max_h"], r["h_ft"])
        g["first_row"] = min(g["first_row"], r["excel_row"])
    ordered = sorted(groups.items(), key=lambda kv: (-kv[1]["max_h"], kv[1]["first_row"]))
    return {gid: i for i, (gid, _) in enumerate(ordered, start=1)}


def autofit_columns(ws, header_row):
    """Widen every VISIBLE column so the longest cell value is not cut off, and
    wrap the header row. Hidden columns are skipped (left hidden)."""
    from openpyxl.utils import get_column_letter
    for col in range(1, ws.max_column + 1):
        letter = get_column_letter(col)
        if ws.column_dimensions[letter].hidden:
            continue
        longest = 0
        for cell in ws[letter]:
            if cell.value is not None:
                for line in str(cell.value).split("\n"):
                    longest = max(longest, len(line))
        ws.column_dimensions[letter].width = min(max(longest + 2, 9), 90)
    # wrap + bold-ish header so long titles stay readable
    for cell in ws[header_row]:
        if cell.value is not None:
            cell.alignment = Alignment(wrap_text=False, vertical="center")
    ws.row_dimensions[header_row].height = 15


def build(spec):
    wb = openpyxl.load_workbook(spec["input_file"])   # keep formatting/structure
    ws = wb[spec["sheet"]]
    header_row = spec.get("header_row", 1)

    original_max_col = ws.max_column

    # 1) Hide every column NOT in keep_visible; make sure kept columns are visible.
    #    We NEVER delete a column. Match by exact header text (model already
    #    resolved fuzzy variants like "case #" -> "Case Number" into actual headers).
    keep = {str(h).strip().lower() for h in spec.get("keep_visible", [])}
    for c in range(1, original_max_col + 1):
        letter = get_column_letter(c)
        header = ws.cell(row=header_row, column=c).value
        header_norm = str(header).strip().lower() if header is not None else ""
        ws.column_dimensions[letter].hidden = header_norm not in keep

    # 2) Append the JRC columns to the RIGHT of all existing columns, in order,
    #    then the identifier column on the far right. Colour the new headers.
    first_new = original_max_col + 1
    headers = NEW_COLUMNS + [ID_HEADER]
    for offset, name in enumerate(headers):
        cell = ws.cell(row=header_row, column=first_new + offset, value=name)
        if name in GREEN_HEADERS:
            cell.fill = GREEN
            cell.font = GREEN_FONT
        elif name in YELLOW_HEADERS:
            cell.fill = YELLOW
            cell.font = YELLOW_FONT
    col_idx = {name: first_new + i for i, name in enumerate(NEW_COLUMNS)}
    id_col = first_new + len(NEW_COLUMNS)

    # 3) Convert + round every row first (we need heights before numbering trucks).
    computed = []
    for row in spec["rows"]:
        computed.append({
            "excel_row": row["excel_row"],
            "l_ft": to_ft(row["length"]),
            "w_ft": to_ft(row["width"]),
            "h_ft": to_ft(row["height"]),
            "lbs": to_lb(row["weight"]),
            "equipment": row.get("equipment"),
            "truck": row.get("truck"),
            "rate": row.get("rate"),
        })

    # 4) Truck numbering: tallest loads get the lowest numbers.
    truck_map = renumber_by_rule(computed)

    # 5) Fill each data row and apply the highlight rules to the converted cells.
    for i, row in enumerate(computed, start=1):
        r = row["excel_row"]
        ws.cell(row=r, column=col_idx["L FT"], value=row["l_ft"])
        ws.cell(row=r, column=col_idx["W FT"], value=row["w_ft"])
        ws.cell(row=r, column=col_idx["H FT"], value=row["h_ft"])
        ws.cell(row=r, column=col_idx["POUNDS"], value=row["lbs"])
        ws.cell(row=r, column=col_idx["JRC EQUIPMENT"], value=row["equipment"])
        ws.cell(row=r, column=col_idx["JRC TRUCK"], value=truck_map[row["truck"]])
        ws.cell(row=r, column=col_idx["JRC RATE"], value=row["rate"])   # blank if None
        ws.cell(row=r, column=id_col, value=i)

        # Height  > 12 ft     -> BLUE the H FT cell.
        # Weight  > 47,500 lb -> BLUE the POUNDS cell (the payload ceiling).
        # Width   > 8.5 ft    -> BLUE the W FT cell; if width > 16 ft, RED instead
        # (red supersedes blue for an extreme over-width load).
        if row["h_ft"] > 12:
            ws.cell(row=r, column=col_idx["H FT"]).fill = BLUE
        if row["lbs"] > PAYLOAD_CEILING_LB:
            ws.cell(row=r, column=col_idx["POUNDS"]).fill = BLUE
        if row["w_ft"] > 16:
            wfc = ws.cell(row=r, column=col_idx["W FT"])
            wfc.fill = YELLOW_HL
            wfc.font = RED_FONT
        elif row["w_ft"] > 8.5:
            ws.cell(row=r, column=col_idx["W FT"]).fill = BLUE

    autofit_columns(ws, header_row)
    wb.save(spec["output_file"])
    return spec["output_file"]



# ===================== 2-D packer =====================
DECKW=8.5
NORMAL={3:32.0,2:29.0,1:48.0,0:53.0}      # normal trailer LENGTH limit per class
def cap_len(cls,ext): return 53.0 if (ext or cls==0) else (48.0 if cls==1 else NORMAL[cls])
def cap_wt(cls): return 35000 if cls in (2,3) else 47500   # double-drop family 35k; flat/step 47.5k

def prep(items):
    # add packing geometry: depth=along-trailer (max horiz dim), lane=across-deck (min horiz dim)
    for r,it in items.items():
        L,W=it['L'],it['W']
        across=min(L,W); depth=max(L,W)
        it['depth']=depth
        it['ow']= across>DECKW+1e-9
        it['lane']= DECKW if it['ow'] else across
        it['owwidth']= across if it['ow'] else 0.0
        it['oversize']= depth>NORMAL[it['rk']]+1e-9   # too long for its band's normal well
    return items

def pack(items, order):
    # EXTENDED trailers (ext double drop / ext minideck) are stretched to fit ONE
    # oversize piece. You may add more pieces ONLY if they do NOT further lengthen the
    # trailer -- i.e. they fit BESIDE the oversize piece within its existing length
    # (an existing shelf), never in a NEW length position. Adding a piece that needs
    # more length is a divisible load and is not permittable.
    # Oversize pieces never join another truck (and two oversize can't share).
    trucks=[]
    for r in order:
        it=items[r]; placed=False
        if not it['oversize']:
            for t in trucks:
                newcls=max(t['rk'],it['rk'])
                if t['wt']+it['wt']>cap_wt(newcls)+1e-6: continue
                if it['ow'] and t['ow_w']>DECKW and it['owwidth']>t['ow_w']+1e-9: continue
                done=False
                if not it['ow']:
                    for sh in t['shelves']:                 # side-by-side within an existing shelf
                        if sh['ow']: continue
                        if sh['depth']+1e-9>=it['depth'] and sh['across']+it['lane']<=DECKW+1e-9:
                            sh['across']+=it['lane']; done=True; break
                if not done and not t['ext']:               # NEW length position only on NON-extended trucks
                    if t['len']+it['depth']<=cap_len(newcls,False)+1e-9:
                        t['shelves'].append({'depth':it['depth'],'across':it['lane'],'ow':it['ow']})
                        t['len']+=it['depth']; done=True
                if done:
                    t['wt']+=it['wt']; t['rk']=newcls
                    if it['ow']: t['ow_w']=max(t['ow_w'],it['owwidth'])
                    t['rows'].append(r); placed=True; break
        if not placed:
            trucks.append({'shelves':[{'depth':it['depth'],'across':it['lane'],'ow':it['ow']}],
                           'len':it['depth'],'wt':it['wt'],'rk':it['rk'],'ext':it['oversize'],
                           'ow_w':it['owwidth'] if it['ow'] else 0.0,'rows':[r]})
    return trucks

def best_pack(items):
    items=prep(items)
    keys={
     'rk,depth':lambda r:(-items[r]['rk'],-items[r]['depth'],-items[r]['wt']),
     'rk,wt':lambda r:(-items[r]['rk'],-items[r]['wt'],-items[r]['depth']),
     'depth':lambda r:(-items[r]['depth'],-items[r]['rk']),
     'rk,ow,depth':lambda r:(-items[r]['rk'],-int(items[r]['ow']),-items[r]['depth']),
    }
    best=None
    for nm,k in keys.items():
        # fresh copy of truck state each run
        for r in items: pass
        t=pack(items,sorted(items,key=k))
        if best is None or len(t)<len(best[1]): best=(nm,t)
    return best

def label(rows,items):
    cls=max(items[x]['rk'] for x in rows); ext=any(items[x]['oversize'] for x in rows)
    if cls==3: return 'EXT MINIDECK' if ext else 'MINIDECK'
    if cls==2: return 'EXTENDED DOUBLE DROP' if ext else 'DOUBLE DROP'
    return 'STEPDECK' if cls==1 else 'FLATBED'

def validate(trucks,items):
    bad=[]
    for t in trucks:
        cls=t['rk']
        if t['wt']>cap_wt(cls)+1 and len(t['rows'])>1: bad.append(('WT',t['rows'],t['wt'],cap_wt(cls)))
        for sh in t['shelves']:
            if sh['across']>DECKW+1e-6: bad.append(('WIDTH',t['rows'],sh['across']))
        if not t['ext'] and t['len']>cap_len(cls,False)+1e-6 and len(t['rows'])>1: bad.append(('LEN',t['rows'],t['len']))
    return bad

if __name__=='__main__':
    import sys
    items={int(k):v for k,v in json.load(open(sys.argv[1])).items()}
    nm,trucks=best_pack(items)
    print('best heuristic',nm,'->',len(trucks),'trucks')
    bad=validate(trucks,items); print('violations',len(bad), bad[:5])
    out=[{'rows':sorted(t['rows']),'eq':label(t['rows'],items)} for t in trucks]
    json.dump(out,open(sys.argv[2],'w'))


# ===================== engine =====================
def load_any(file_bytes, filename):
    """Return (openpyxl_workbook, xlsx_bytes). Handles .xlsx natively and old .xls
    via xlrd (no LibreOffice needed -> works on Vercel serverless)."""
    if filename.lower().endswith(".xls"):
        import xlrd                      # add 'xlrd' to requirements.txt for prod
        book = xlrd.open_workbook(file_contents=file_bytes)
        wb = openpyxl.Workbook(); wb.remove(wb.active)
        for sh in book.sheets():
            ws = wb.create_sheet(sh.name)
            for r in range(sh.nrows):
                for c in range(sh.ncols):
                    ws.cell(r+1, c+1, sh.cell_value(r, c))
        buf = io.BytesIO(); wb.save(buf)
        return openpyxl.load_workbook(io.BytesIO(buf.getvalue()), data_only=True), buf.getvalue()
    return openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True), file_bytes

# ---------------------------------------------------------------- detection ---
DIM_UNIT_RE = re.compile(r"(cm|mm|m\b|in\b|inch|inches|ft|foot|feet)", re.I)

def _norm(s):
    return re.sub(r"[^a-z0-9]", "", str(s).lower()) if s is not None else ""

def find_header_row(ws, max_scan=30):
    """Header row = the first row (within the first few) that looks like column
    titles for length/width/height + a description/case column."""
    best, best_score = 1, -1
    for r in range(1, min(max_scan, ws.max_row) + 1):
        vals = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
        norm = [_norm(v) for v in vals]
        score = 0
        for key in ("ln", "length", "lnin", "lncm", "wd", "width", "ht", "height",
                    "description", "casenumber", "grswtlbs", "grosssweight", "weight"):
            if any(key in n for n in norm):
                score += 1
        if score > best_score:
            best, best_score = r, score
    return best

def _match(headers, *cands):
    norm = {i: _norm(h) for i, h in headers.items()}
    # exact-ish contains match, in candidate priority order
    for cand in cands:
        c = _norm(cand)
        for i, n in norm.items():
            if n == c:
                return i
    for cand in cands:
        c = _norm(cand)
        for i, n in norm.items():
            if c and c in n:
                return i
    return None

def detect_columns(ws, header_row):
    headers = {c: ws.cell(header_row, c).value for c in range(1, ws.max_column + 1)
               if ws.cell(header_row, c).value is not None}
    col = {}
    col["desc"] = _match(headers, "description", "box description", "project name", "mli")
    col["case"] = _match(headers, "case number", "box/case number", "order release id",
                          "case no", "case #")
    # prefer imperial dimension columns, fall back to metric
    col["len_in"] = _match(headers, "ln_in", "length (in", "lenght (in")
    col["wd_in"]  = _match(headers, "wd_in", "width (in")
    col["ht_in"]  = _match(headers, "ht_in", "height (in")
    col["wt_lb"]  = _match(headers, "grs_wt_lbs", "gross_weight (lb", "total weight",
                           "gross weight")
    col["len_cm"] = _match(headers, "ln_cm", "lenght (cm", "length (cm")
    col["wd_cm"]  = _match(headers, "wd_cm", "width (cm")
    col["ht_cm"]  = _match(headers, "ht_cm", "height (cm")
    col["wt_kg"]  = _match(headers, "grs_wt_kg", "gross_weight (kg")
    # generic fallbacks
    if col["len_in"] is None and col["len_cm"] is None:
        col["len_in"] = _match(headers, "length", "ln");
    if col["wd_in"] is None and col["wd_cm"] is None:
        col["wd_in"] = _match(headers, "width", "wd")
    if col["ht_in"] is None and col["ht_cm"] is None:
        col["ht_in"] = _match(headers, "height", "ht")
    if col["wt_lb"] is None and col["wt_kg"] is None:
        col["wt_lb"] = _match(headers, "weight")
    return headers, col

def pick_dims(col):
    """Return (len_col, wd_col, ht_col, wt_col, unit_len, unit_wt) preferring imperial."""
    if col["len_in"] and col["wd_in"] and col["ht_in"]:
        ul = "in"
    elif col["len_cm"] and col["wd_cm"] and col["ht_cm"]:
        return col["len_cm"], col["wd_cm"], col["ht_cm"], (col["wt_kg"] or col["wt_lb"]), "cm", ("kg" if col["wt_kg"] else "lb")
    else:
        ul = "in"
    return col["len_in"], col["wd_in"], col["ht_in"], (col["wt_lb"] or col["wt_kg"]), ul, ("lb" if col["wt_lb"] else "kg")

# ----------------------------------------------------------------- numbers ---
def num(v):
    try:
        return float(str(v).replace(",", "").strip())
    except Exception:
        return None

def feet(v, unit):
    f = {"in": 1/12.0, "cm": 0.0328084, "mm": 0.00328084, "m": 3.2808399, "ft": 1.0}[unit]
    return v * f

# --------------------------------------------------------------- equipment ---
def bclass(h_ft):
    return 3 if h_ft > 14.5 else 2 if h_ft > 10 else 1 if h_ft > 8.5 else 0

# -------------------------------------------------------------- main entry ---
def _scan(ws):
    hr = find_header_row(ws)
    headers, col = detect_columns(ws, hr)
    lc, wc, hc, wtc, ul, uw = pick_dims(col)
    return hr, headers, col, lc, wc, hc, wtc, ul, uw

def _rows_for(ws, hr, col, lc, wc, hc, wtc, ul, uw):
    items, skip = {}, {"consolidated": 0, "zero": 0}
    for r in range(hr + 1, ws.max_row + 1):
        lv = ws.cell(r, lc).value
        case = ws.cell(r, col["case"]).value if col["case"] else None
        desc = ws.cell(r, col["desc"]).value if col["desc"] else None
        if lv is None and case is None and desc is None:
            continue
        if case and "onsolidat" in str(case):
            skip["consolidated"] += 1; continue
        L, W, H, WT = num(lv), num(ws.cell(r, wc).value), num(ws.cell(r, hc).value), num(ws.cell(r, wtc).value)
        if (L in (0, None)) and (WT in (0, None)):
            skip["zero"] += 1; continue
        Lf, Wf, Hf = feet(L, ul), feet(W, ul), feet(H, ul)
        wt_lb = WT if uw == "lb" else WT * 2.20462262
        items[r] = dict(L=Lf, W=Wf, H=Hf, wt=wt_lb, rk=bclass(Hf), ow=Wf > 8.5)
    return items, skip

def process(file_bytes, filename="manifest.xlsx"):
    """Run the full pipeline over EVERY cargo tab. Returns (xlsx_bytes, summary)."""
    import tempfile, os
    from collections import Counter
    wb, xlsx_bytes = load_any(file_bytes, filename)

    cargo = []
    for sh in wb.worksheets:
        if sh.max_row < 2:
            continue
        hr, headers, col, lc, wc, hc, wtc, ul, uw = _scan(sh)
        if all([lc, wc, hc, wtc]):
            cargo.append((sh, hr, headers, col, lc, wc, hc, wtc, ul, uw))
    if not cargo:
        raise ValueError("Couldn't find a tab with Length/Width/Height/Weight columns. "
                         "If the cargo list is on a specific tab, make sure that tab has those headers.")

    cur = xlsx_bytes
    tot_pieces = tot_trucks = tot_ow = 0; tot_wt = 0.0
    mix = Counter(); superloads = []; skipped = {"consolidated": 0, "zero": 0}; tabs = []
    for (sh, hr, headers, col, lc, wc, hc, wtc, ul, uw) in cargo:
        items, skip = _rows_for(sh, hr, col, lc, wc, hc, wtc, ul, uw)
        if not items:
            continue
        _, trucks = best_pack(items)
        groups = [{"rows": sorted(t["rows"]), "eq": label(t["rows"], items)} for t in trucks]
        keep = []
        for _k, idx in [("desc", col["desc"]), ("case", col["case"]),
                        (None, lc), (None, wc), (None, hc), (None, wtc)]:
            if idx:
                keep.append(headers[idx])
        rows_spec, gid = [], 0
        for g in groups:
            gid += 1
            for r in g["rows"]:
                rows_spec.append({"excel_row": r,
                    "length": {"value": num(sh.cell(r, lc).value), "unit": ul},
                    "width":  {"value": num(sh.cell(r, wc).value), "unit": ul},
                    "height": {"value": num(sh.cell(r, hc).value), "unit": ul},
                    "weight": {"value": num(sh.cell(r, wtc).value), "unit": uw},
                    "equipment": g["eq"], "truck": "g%03d" % gid, "rate": None})
        rows_spec.sort(key=lambda x: x["excel_row"])
        tin = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False); tin.write(cur); tin.close()
        tout = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False); tout.close()
        build({"input_file": tin.name, "output_file": tout.name, "sheet": sh.title,
               "header_row": hr, "keep_visible": keep, "rows": rows_spec})
        cur = open(tout.name, "rb").read(); os.unlink(tin.name); os.unlink(tout.name)

        s = make_summary(items, groups, skip, len(items))
        tot_pieces += s["pieces"]; tot_trucks += s["trucks"]; tot_ow += s["over_width_pieces"]; tot_wt += s["total_weight_lb"]
        for k, v in s["equipment_mix"].items():
            mix[k] += v
        for row, reason in s["superloads"]:
            superloads.append(("%s (%s)" % (row, sh.title), reason))
        skipped["consolidated"] += skip["consolidated"]; skipped["zero"] += skip["zero"]
        tabs.append({"tab": sh.title, "trucks": s["trucks"], "pieces": s["pieces"]})

    if tot_pieces == 0:
        raise ValueError("No cargo rows found.")
    summary = {"pieces": tot_pieces, "trucks": tot_trucks, "equipment_mix": dict(mix),
               "over_width_pieces": tot_ow, "superloads": superloads, "skipped": skipped,
               "total_weight_lb": round(tot_wt), "tabs": tabs}
    return cur, summary

# ----------------------------------------------------------------- summary ---
def make_summary(items, groups, skipped, n):
    from collections import Counter
    mix = Counter(g["eq"] for g in groups)
    ow = sum(1 for i in items.values() if i["W"] > 8.5)
    superload = []
    for r, i in items.items():
        if i["H"] > 13.5: superload.append((r, "over-height %.1f ft" % i["H"]))
        if i["W"] > 16: superload.append((r, "over-width %.1f ft" % i["W"]))
        if i["wt"] > 47500: superload.append((r, "heavy-haul %.0f lb" % i["wt"]))
        if max(i["L"], i["W"]) > 55: superload.append((r, "over-length %.1f ft" % max(i["L"], i["W"])))
    return {
        "pieces": n,
        "trucks": len(groups),
        "equipment_mix": dict(mix),
        "over_width_pieces": ow,
        "superloads": superload,
        "skipped": skipped,
        "total_weight_lb": round(sum(i["wt"] for i in items.values())),
    }

# ===================== Flask app (single entrypoint: app.py) =====================
import os, base64 as _b64
from flask import Flask, request, jsonify, send_file, Response
app = Flask(__name__)

_PAGE = _b64.b64decode("PCFkb2N0eXBlIGh0bWw+CjxodG1sIGxhbmc9ImVuIj4KPGhlYWQ+CjxtZXRhIGNoYXJzZXQ9InV0Zi04Ij4KPG1ldGEgbmFtZT0idmlld3BvcnQiIGNvbnRlbnQ9IndpZHRoPWRldmljZS13aWR0aCxpbml0aWFsLXNjYWxlPTEiPgo8dGl0bGU+VHJ1Y2sgUGxhbiBCdWlsZGVyICBKUkMgVHJhbnNwb3J0YXRpb248L3RpdGxlPgo8bGluayByZWw9InByZWNvbm5lY3QiIGhyZWY9Imh0dHBzOi8vZm9udHMuZ29vZ2xlYXBpcy5jb20iPgo8bGluayByZWw9InByZWNvbm5lY3QiIGhyZWY9Imh0dHBzOi8vZm9udHMuZ3N0YXRpYy5jb20iIGNyb3Nzb3JpZ2luPgo8bGluayBocmVmPSJodHRwczovL2ZvbnRzLmdvb2dsZWFwaXMuY29tL2NzczI/ZmFtaWx5PUludGVyOndnaHRANDAwOzUwMDs2MDA7NzAwJmRpc3BsYXk9c3dhcCIgcmVsPSJzdHlsZXNoZWV0Ij4KPHN0eWxlPgogOnJvb3R7CiAgIC0tYmc6IzBiMTIyMDsgLS1iZzI6IzBmMWEzMDsgLS1wYW5lbDojMGYxODMwOyAtLXBhbmVsMjojMTMyMDNjOwogICAtLWxpbmU6cmdiYSgyNTUsMjU1LDI1NSwuMDkpOyAtLWxpbmUyOnJnYmEoMjU1LDI1NSwyNTUsLjE2KTsKICAgLS1pbms6I2VlZjJmOTsgLS1zdWI6IzlhYTdjMjsgLS1oaW50OiM2YzdhOTk7CiAgIC0tYnJhbmQ6IzNiODJmNjsgLS1icmFuZDI6IzYwYTVmYTsgLS1icmFuZGluazojYmNkNGZmOwogICAtLW9rOiMzNGQzOTk7IC0tb2tiZzpyZ2JhKDUyLDIxMSwxNTMsLjEyKTsgLS13YXJuOiNmODcxNzE7IC0td2FybmJnOnJnYmEoMjQ4LDExMywxMTMsLjEyKTsKICAgLS1yOjE0cHg7IC0tcjI6MjBweDsKIH0KICp7Ym94LXNpemluZzpib3JkZXItYm94fQogaHRtbCxib2R5e21hcmdpbjowO2hlaWdodDoxMDAlfQogYm9keXsKICAgZm9udC1mYW1pbHk6SW50ZXIsLWFwcGxlLXN5c3RlbSxCbGlua01hY1N5c3RlbUZvbnQsIlNlZ29lIFVJIixSb2JvdG8sc2Fucy1zZXJpZjsKICAgY29sb3I6dmFyKC0taW5rKTtiYWNrZ3JvdW5kOnZhcigtLWJnKTsKICAgYmFja2dyb3VuZC1pbWFnZTpyYWRpYWwtZ3JhZGllbnQoOTAwcHggNTAwcHggYXQgODAlIC0xMCUscmdiYSg1OSwxMzAsMjQ2LC4xOCksdHJhbnNwYXJlbnQgNjAlKSwKICAgICAgICAgICAgICAgICAgICByYWRpYWwtZ3JhZGllbnQoNzAwcHggNTAwcHggYXQgMCUgMCUscmdiYSg5OSwxMDIsMjQxLC4xMiksdHJhbnNwYXJlbnQgNTUlKTsKICAgLXdlYmtpdC1mb250LXNtb290aGluZzphbnRpYWxpYXNlZDtsaW5lLWhlaWdodDoxLjU7CiB9CiAud3JhcHttYXgtd2lkdGg6NzYwcHg7bWFyZ2luOjAgYXV0bztwYWRkaW5nOjI4cHggMjBweCA2MHB4fQogbmF2e2Rpc3BsYXk6ZmxleDthbGlnbi1pdGVtczpjZW50ZXI7Z2FwOjExcHg7Y29sb3I6I2ZmZn0KIC5sb2dve3dpZHRoOjM4cHg7aGVpZ2h0OjM4cHg7Ym9yZGVyLXJhZGl1czoxMXB4O2JhY2tncm91bmQ6bGluZWFyLWdyYWRpZW50KDE0MGRlZywjMjU2M2ViLCM3YzNhZWQpOwogICBkaXNwbGF5OmdyaWQ7cGxhY2UtaXRlbXM6Y2VudGVyO2ZvbnQtc2l6ZToyMHB4O2JveC1zaGFkb3c6MCA4cHggMjRweCAtOHB4ICMyNTYzZWI4OH0KIG5hdiAubmFtZXtmb250LXdlaWdodDo3MDA7Zm9udC1zaXplOjE2cHg7bGV0dGVyLXNwYWNpbmc6LjJweH0KIG5hdiAudGFne2ZvbnQtc2l6ZToxMnB4O2NvbG9yOnZhcigtLXN1Yil9CiBuYXYgLnBpbGx7bWFyZ2luLWxlZnQ6YXV0bztmb250LXNpemU6MTJweDtjb2xvcjp2YXIoLS1icmFuZGluayk7YmFja2dyb3VuZDpyZ2JhKDU5LDEzMCwyNDYsLjE0KTsKICAgYm9yZGVyOjFweCBzb2xpZCByZ2JhKDU5LDEzMCwyNDYsLjMpO2JvcmRlci1yYWRpdXM6MzBweDtwYWRkaW5nOjVweCAxMnB4fQogLmhlcm97cGFkZGluZzoxNHB4IDAgMjBweDt0ZXh0LWFsaWduOmNlbnRlcn0KIC5oZXJvIGgxe2ZvbnQtc2l6ZToyNnB4O2xpbmUtaGVpZ2h0OjEuMTI7bWFyZ2luOjAgMCA4cHg7bGV0dGVyLXNwYWNpbmc6LS41cHg7Zm9udC13ZWlnaHQ6NzAwO2JhY2tncm91bmQ6bGluZWFyLWdyYWRpZW50KDEyMGRlZywjZmZmLCNiY2Q0ZmYpOy13ZWJraXQtYmFja2dyb3VuZC1jbGlwOnRleHQ7YmFja2dyb3VuZC1jbGlwOnRleHQ7Y29sb3I6dHJhbnNwYXJlbnR9CiAuaGVybyBwe21hcmdpbjowIGF1dG87bWF4LXdpZHRoOjU0MHB4O2NvbG9yOnZhcigtLXN1Yik7Zm9udC1zaXplOjE2cHh9CiAucGFuZWx7YmFja2dyb3VuZDpsaW5lYXItZ3JhZGllbnQoMTgwZGVnLHZhcigtLXBhbmVsKSx2YXIoLS1wYW5lbDIpKTtib3JkZXI6MXB4IHNvbGlkIHZhcigtLWxpbmUpOwogICBib3JkZXItcmFkaXVzOnZhcigtLXIyKTtwYWRkaW5nOjEwcHg7Ym94LXNoYWRvdzowIDMwcHggODBweCAtNDBweCAjMDAwYSwgaW5zZXQgMCAxcHggMCByZ2JhKDI1NSwyNTUsMjU1LC4wNCl9CiAuZHJvcHtib3JkZXI6MS41cHggZGFzaGVkIHZhcigtLWxpbmUyKTtib3JkZXItcmFkaXVzOnZhcigtLXIpO3BhZGRpbmc6NDhweCAyMnB4O3RleHQtYWxpZ246Y2VudGVyOwogICBjdXJzb3I6cG9pbnRlcjt0cmFuc2l0aW9uOi4xOHM7YmFja2dyb3VuZDpyZ2JhKDI1NSwyNTUsMjU1LC4wMTUpfQogLmRyb3A6aG92ZXJ7Ym9yZGVyLWNvbG9yOnZhcigtLWJyYW5kMik7YmFja2dyb3VuZDpyZ2JhKDU5LDEzMCwyNDYsLjA2KX0KIC5kcm9wLm92ZXJ7Ym9yZGVyLWNvbG9yOnZhcigtLWJyYW5kMik7YmFja2dyb3VuZDpyZ2JhKDU5LDEzMCwyNDYsLjEpO3RyYW5zZm9ybTpzY2FsZSgxLjAwNCl9CiAuZHJvcCAuaWN7d2lkdGg6NThweDtoZWlnaHQ6NThweDtib3JkZXItcmFkaXVzOjE2cHg7bWFyZ2luOjAgYXV0byAxNHB4O2Rpc3BsYXk6Z3JpZDtwbGFjZS1pdGVtczpjZW50ZXI7CiAgIGZvbnQtc2l6ZToyN3B4O2JhY2tncm91bmQ6cmdiYSg1OSwxMzAsMjQ2LC4xNCk7Ym9yZGVyOjFweCBzb2xpZCByZ2JhKDU5LDEzMCwyNDYsLjI4KTtjb2xvcjp2YXIoLS1icmFuZDIpfQogLmRyb3AgLnR7Zm9udC1zaXplOjE3cHg7Zm9udC13ZWlnaHQ6NjAwfSAuZHJvcCAuc3tjb2xvcjp2YXIoLS1zdWIpO2ZvbnQtc2l6ZToxM3B4O21hcmdpbi10b3A6NXB4fQogLmFjdGlvbnN7ZGlzcGxheTpmbGV4O2FsaWduLWl0ZW1zOmNlbnRlcjtnYXA6MTJweDtwYWRkaW5nOjE0cHggMTJweCA2cHg7ZmxleC13cmFwOndyYXB9CiBidXR0b257Zm9udC1mYW1pbHk6aW5oZXJpdDtmb250LXdlaWdodDo2MDA7Zm9udC1zaXplOjE1cHg7Ym9yZGVyOjA7Ym9yZGVyLXJhZGl1czoxMXB4O3BhZGRpbmc6MTNweCAyMnB4O2N1cnNvcjpwb2ludGVyO3RyYW5zaXRpb246LjE2c30KIC5wcmltYXJ5e2JhY2tncm91bmQ6bGluZWFyLWdyYWRpZW50KDE4MGRlZywjM2I4MmY2LCMyNTYzZWIpO2NvbG9yOiNmZmY7Ym94LXNoYWRvdzowIDEycHggMjZweCAtMTJweCAjMjU2M2ViY2N9CiAucHJpbWFyeTpob3Zlcjpub3QoOmRpc2FibGVkKXtmaWx0ZXI6YnJpZ2h0bmVzcygxLjA3KX0gYnV0dG9uOmRpc2FibGVke29wYWNpdHk6LjQ7Y3Vyc29yOmRlZmF1bHQ7Ym94LXNoYWRvdzpub25lfQogLmdob3N0e2JhY2tncm91bmQ6cmdiYSgyNTUsMjU1LDI1NSwuMDUpO2NvbG9yOnZhcigtLWluayk7Ym9yZGVyOjFweCBzb2xpZCB2YXIoLS1saW5lKX0KIC5zdGF0dXN7ZGlzcGxheTpmbGV4O2FsaWduLWl0ZW1zOmNlbnRlcjtnYXA6OXB4O2NvbG9yOnZhcigtLXN1Yik7Zm9udC1zaXplOjE0cHh9CiAuc3Bpbnt3aWR0aDoxNnB4O2hlaWdodDoxNnB4O2JvcmRlcjoycHggc29saWQgcmdiYSgyNTUsMjU1LDI1NSwuMTgpO2JvcmRlci10b3AtY29sb3I6dmFyKC0tYnJhbmQyKTtib3JkZXItcmFkaXVzOjUwJTthbmltYXRpb246c3AgLjdzIGxpbmVhciBpbmZpbml0ZX0KIEBrZXlmcmFtZXMgc3B7dG97dHJhbnNmb3JtOnJvdGF0ZSgzNjBkZWcpfX0KIC5mZWF0c3tkaXNwbGF5OmdyaWQ7Z3JpZC10ZW1wbGF0ZS1jb2x1bW5zOnJlcGVhdCgzLDFmcik7Z2FwOjEycHg7bWFyZ2luLXRvcDoxOHB4fQogLmZlYXR7YmFja2dyb3VuZDpyZ2JhKDI1NSwyNTUsMjU1LC4wMjUpO2JvcmRlcjoxcHggc29saWQgdmFyKC0tbGluZSk7Ym9yZGVyLXJhZGl1czp2YXIoLS1yKTtwYWRkaW5nOjE1cHggMTRweH0KIC5mZWF0IC5maXtjb2xvcjp2YXIoLS1icmFuZDIpO2ZvbnQtc2l6ZToxOHB4fSAuZmVhdCBoNHttYXJnaW46OHB4IDAgM3B4O2ZvbnQtc2l6ZToxNHB4O2ZvbnQtd2VpZ2h0OjYwMH0KIC5mZWF0IHB7bWFyZ2luOjA7Y29sb3I6dmFyKC0tc3ViKTtmb250LXNpemU6MTIuNXB4O2xpbmUtaGVpZ2h0OjEuNX0KIC5yZXN1bHR7bWFyZ2luLXRvcDoxOHB4fQogLnJjYXJke2JhY2tncm91bmQ6bGluZWFyLWdyYWRpZW50KDE4MGRlZyx2YXIoLS1wYW5lbCksdmFyKC0tcGFuZWwyKSk7Ym9yZGVyOjFweCBzb2xpZCB2YXIoLS1saW5lKTtib3JkZXItcmFkaXVzOnZhcigtLXIyKTtwYWRkaW5nOjIycHh9CiAucmhlYWR7ZGlzcGxheTpmbGV4O2FsaWduLWl0ZW1zOmNlbnRlcjtnYXA6MTBweDtmb250LXdlaWdodDo2MDA7Zm9udC1zaXplOjE2cHg7bWFyZ2luLWJvdHRvbToxOHB4fQogLmNoZWNre3dpZHRoOjI2cHg7aGVpZ2h0OjI2cHg7Ym9yZGVyLXJhZGl1czo1MCU7YmFja2dyb3VuZDp2YXIoLS1va2JnKTtjb2xvcjp2YXIoLS1vayk7ZGlzcGxheTpncmlkO3BsYWNlLWl0ZW1zOmNlbnRlcjtmb250LXNpemU6MTVweH0KIC5zdGF0c3tkaXNwbGF5OmdyaWQ7Z3JpZC10ZW1wbGF0ZS1jb2x1bW5zOnJlcGVhdCg0LDFmcik7Z2FwOjEycHh9CiAuc3RhdHtiYWNrZ3JvdW5kOnJnYmEoMjU1LDI1NSwyNTUsLjAzKTtib3JkZXI6MXB4IHNvbGlkIHZhcigtLWxpbmUpO2JvcmRlci1yYWRpdXM6dmFyKC0tcik7cGFkZGluZzoxNnB4IDE0cHh9CiAuc3RhdCAubntmb250LXNpemU6MzBweDtmb250LXdlaWdodDo3MDA7bGV0dGVyLXNwYWNpbmc6LTFweH0gLnN0YXQgLmx7Y29sb3I6dmFyKC0tc3ViKTtmb250LXNpemU6MTEuNXB4O3RleHQtdHJhbnNmb3JtOnVwcGVyY2FzZTtsZXR0ZXItc3BhY2luZzouNnB4O21hcmdpbi10b3A6M3B4fQogLm1peHtkaXNwbGF5OmZsZXg7ZmxleC13cmFwOndyYXA7Z2FwOjhweDttYXJnaW46MThweCAwfQogLnBpbGwye2Rpc3BsYXk6aW5saW5lLWZsZXg7YWxpZ24taXRlbXM6Y2VudGVyO2dhcDo3cHg7YmFja2dyb3VuZDpyZ2JhKDI1NSwyNTUsMjU1LC4wNCk7Ym9yZGVyOjFweCBzb2xpZCB2YXIoLS1saW5lKTtib3JkZXItcmFkaXVzOjMwcHg7cGFkZGluZzo3cHggMTNweDtmb250LXNpemU6MTNweH0KIC5waWxsMiBie2JhY2tncm91bmQ6dmFyKC0tYnJhbmQpO2NvbG9yOiNmZmY7Ym9yZGVyLXJhZGl1czoyMHB4O21pbi13aWR0aDoyMHB4O3RleHQtYWxpZ246Y2VudGVyO3BhZGRpbmc6MXB4IDdweDtmb250LXNpemU6MTJweH0KIC5iYW5uZXJ7Ym9yZGVyLXJhZGl1czp2YXIoLS1yKTtwYWRkaW5nOjE0cHggMTZweDttYXJnaW46NHB4IDAgMThweDtmb250LXNpemU6MTRweDtkaXNwbGF5OmZsZXg7Z2FwOjExcHh9CiAuYmFubmVyLndhcm57YmFja2dyb3VuZDp2YXIoLS13YXJuYmcpO2JvcmRlcjoxcHggc29saWQgcmdiYSgyNDgsMTEzLDExMywuMyk7Y29sb3I6I2ZlY2FjYX0KIC5iYW5uZXIub2t7YmFja2dyb3VuZDp2YXIoLS1va2JnKTtib3JkZXI6MXB4IHNvbGlkIHJnYmEoNTIsMjExLDE1MywuMyk7Y29sb3I6I2JiZjdkMH0KIC5iYW5uZXIgaXtmb250LXNpemU6MThweDttYXJnaW4tdG9wOjFweH0KIC5iYW5uZXIgLnR0bHtmb250LXdlaWdodDo2MDB9CiAuZGx7ZGlzcGxheTppbmxpbmUtZmxleDthbGlnbi1pdGVtczpjZW50ZXI7Z2FwOjlweDtiYWNrZ3JvdW5kOmxpbmVhci1ncmFkaWVudCgxODBkZWcsIzEwYjk4MSwjMDU5NjY5KTtjb2xvcjojZmZmOwogICBib3JkZXItcmFkaXVzOjExcHg7cGFkZGluZzoxM3B4IDIycHg7Zm9udC13ZWlnaHQ6NjAwO3RleHQtZGVjb3JhdGlvbjpub25lO2ZvbnQtc2l6ZToxNXB4O2JveC1zaGFkb3c6MCAxMnB4IDI2cHggLTEycHggIzA1OTY2OWNjfQogLmRsOmhvdmVye2ZpbHRlcjpicmlnaHRuZXNzKDEuMDcpfQogLmVycntiYWNrZ3JvdW5kOnZhcigtLXdhcm5iZyk7Ym9yZGVyOjFweCBzb2xpZCByZ2JhKDI0OCwxMTMsMTEzLC4zKTtjb2xvcjojZmVjYWNhO2JvcmRlci1yYWRpdXM6dmFyKC0tcik7cGFkZGluZzoxNXB4fQogLmZvb3R7dGV4dC1hbGlnbjpjZW50ZXI7Y29sb3I6dmFyKC0taGludCk7Zm9udC1zaXplOjEycHg7bWFyZ2luLXRvcDozNHB4fQogLmhpZGV7ZGlzcGxheTpub25lfQogQG1lZGlhKG1heC13aWR0aDo1NjBweCl7LnN0YXRze2dyaWQtdGVtcGxhdGUtY29sdW1uczpyZXBlYXQoMiwxZnIpfS5mZWF0c3tncmlkLXRlbXBsYXRlLWNvbHVtbnM6MWZyfS5oZXJvIGgxe2ZvbnQtc2l6ZTozMnB4fX0KIC53ZWxjb21le2Rpc3BsYXk6ZmxleDtmbGV4LWRpcmVjdGlvbjpjb2x1bW47YWxpZ24taXRlbXM6Y2VudGVyO2dhcDo4cHg7cGFkZGluZzo2cHggMCAxMHB4fQogLndlbGNvbWUgaW1ne2Rpc3BsYXk6YmxvY2s7aGVpZ2h0OmF1dG87d2lkdGg6NDQwcHg7bWF4LXdpZHRoOjg2dnc7aW1hZ2UtcmVuZGVyaW5nOmF1dG99CiAud2VsY29tZSAud3R4dHtmb250LXNpemU6NDZweDtmb250LXdlaWdodDo3MDA7bGV0dGVyLXNwYWNpbmc6LTFweDtsaW5lLWhlaWdodDoxO2JhY2tncm91bmQ6bGluZWFyLWdyYWRpZW50KDEyMGRlZywjZmZmLCNiY2Q0ZmYpOy13ZWJraXQtYmFja2dyb3VuZC1jbGlwOnRleHQ7YmFja2dyb3VuZC1jbGlwOnRleHQ7Y29sb3I6dHJhbnNwYXJlbnR9Cjwvc3R5bGU+CjwvaGVhZD4KPGJvZHk+CjxkaXYgY2xhc3M9IndyYXAiPgogCiAKIDxuYXY+CiAgIDxkaXYgY2xhc3M9ImxvZ28iPiYjMTI4NjY2OzwvZGl2PgogICA8ZGl2PjxkaXYgY2xhc3M9Im5hbWUiPkpSQyZuYnNwO1RyYW5zcG9ydGF0aW9uPC9kaXY+PGRpdiBjbGFzcz0idGFnIj5UcnVjayBQbGFuIEJ1aWxkZXI8L2Rpdj48L2Rpdj4KICAgPGRpdiBjbGFzcz0icGlsbCI+QmV0YTwvZGl2PgogPC9uYXY+CiA8ZGl2IGNsYXNzPSJ3ZWxjb21lIj48c3BhbiBjbGFzcz0id3R4dCI+V2VsY29tZSBEUlIhPC9zcGFuPjxpbWcgc3JjPSJ3ZWxjb21lLnBuZyIgYWx0PSJEUlIgd2VsY29tZSIgb25lcnJvcj0idGhpcy5zdHlsZS5kaXNwbGF5PSdub25lJyI+PC9kaXY+CgogPGRpdiBjbGFzcz0iaGVybyI+CiAgIDxoMT5Gcm9tIG1hbmlmZXN0IHRvIHRydWNrIHBsYW4gaW4gc2Vjb25kczwvaDE+CiAgIDxwPlVwbG9hZCBhIHZlc3NlbCBtYW5pZmVzdCBvciBwYWNraW5nIGxpc3QuIFdlIGFzc2lnbiBldmVyeSBwaWVjZSB0byB0aGUgcmlnaHQgdHJhaWxlciwgY29uc29saWRhdGUgb250byB0aGUgZmV3ZXN0IGxlZ2FsIHRydWNrcywgYW5kIGZsYWcgcGVybWl0IGxvYWRzIGJlZm9yZSBoYW5kaW5nIGJhY2sgYSBmaW5pc2hlZCB3b3JrYm9vay48L3A+CiA8L2Rpdj4KCiA8ZGl2IGNsYXNzPSJwYW5lbCI+CiAgIDxkaXYgaWQ9ImRyb3AiIGNsYXNzPSJkcm9wIj4KICAgICA8ZGl2IGNsYXNzPSJpYyI+JiMxMTAxNDs8L2Rpdj4KICAgICA8ZGl2IGNsYXNzPSJ0IiBpZD0iZHJvcHQiPkRyb3AgeW91ciBzcHJlYWRzaGVldCBoZXJlPC9kaXY+CiAgICAgPGRpdiBjbGFzcz0icyIgaWQ9ImRyb3BzIj5vciBjbGljayB0byBicm93c2UgJm5ic3A7Jm1pZGRvdDsmbmJzcDsgLnhsc3ggb3IgLnhsczwvZGl2PgogICA8L2Rpdj4KICAgPGlucHV0IGlkPSJmaWxlIiB0eXBlPSJmaWxlIiBhY2NlcHQ9Ii54bHN4LC54bHMiIGNsYXNzPSJoaWRlIj4KICAgPGRpdiBjbGFzcz0iYWN0aW9ucyI+CiAgICAgPGJ1dHRvbiBjbGFzcz0icHJpbWFyeSIgaWQ9ImdvIiBkaXNhYmxlZD5CdWlsZCB0cnVjayBwbGFuPC9idXR0b24+CiAgICAgPGJ1dHRvbiBjbGFzcz0iZ2hvc3QgaGlkZSIgaWQ9InJlc2V0Ij5TdGFydCBvdmVyPC9idXR0b24+CiAgICAgPHNwYW4gY2xhc3M9InN0YXR1cyIgaWQ9InN0YXR1cyI+PC9zcGFuPgogICA8L2Rpdj4KIDwvZGl2PgoKIDxkaXYgY2xhc3M9ImZlYXRzIiBpZD0iZmVhdHMiPgogICA8ZGl2IGNsYXNzPSJmZWF0Ij48ZGl2IGNsYXNzPSJmaSI+JiM5ODgxOzwvZGl2PjxoND5BdXRvIHRyYWlsZXIgc2VsZWN0aW9uPC9oND48cD5GbGF0YmVkLCBzdGVwZGVjaywgZG91YmxlIGRyb3AsIG1pbmlkZWNrICBjaG9zZW4gYnkgaGVpZ2h0LCBsZW5ndGggYW5kIHdpZHRoLjwvcD48L2Rpdj4KICAgPGRpdiBjbGFzcz0iZmVhdCI+PGRpdiBjbGFzcz0iZmkiPiYjMTI4MjMwOzwvZGl2PjxoND4yLUQgY29uc29saWRhdGlvbjwvaDQ+PHA+U2lkZS1ieS1zaWRlIGFuZCBlbmQtdG8tZW5kIHBhY2tpbmcgd2l0aGluIGV2ZXJ5IGxlZ2FsIGxlbmd0aCwgd2lkdGggYW5kIHdlaWdodCBsaW1pdC48L3A+PC9kaXY+CiAgIDxkaXYgY2xhc3M9ImZlYXQiPjxkaXYgY2xhc3M9ImZpIj4mIzk4ODg7PC9kaXY+PGg0PlBlcm1pdCBmbGFnczwvaDQ+PHA+T3Zlci13aWR0aCwgb3Zlci1oZWlnaHQgYW5kIGhlYXZ5LWhhdWwgc3VwZXJsb2FkcyBjYWxsZWQgb3V0IGF1dG9tYXRpY2FsbHkuPC9wPjwvZGl2PgogPC9kaXY+CgogPGRpdiBpZD0icmVzdWx0IiBjbGFzcz0icmVzdWx0Ij48L2Rpdj4KIDxkaXYgY2xhc3M9ImZvb3QiPkpSQyBUcmFuc3BvcnRhdGlvbiAmbWlkZG90OyB0cmFpbGVyIHNlbGVjdGlvbiAmYW1wOyBjb25zb2xpZGF0aW9uIGVuZ2luZTwvZGl2Pgo8L2Rpdj4KCjxzY3JpcHQ+CmNvbnN0ICQ9cz0+ZG9jdW1lbnQucXVlcnlTZWxlY3RvcihzKTsKY29uc3QgZHJvcD0kKCIjZHJvcCIpLGZpbGU9JCgiI2ZpbGUiKSxnbz0kKCIjZ28iKSxyZXNldD0kKCIjcmVzZXQiKSxzdGF0dXM9JCgiI3N0YXR1cyIpLHJlc3VsdD0kKCIjcmVzdWx0IiksZmVhdHM9JCgiI2ZlYXRzIik7CmxldCBjaG9zZW49bnVsbDsKZHJvcC5vbmNsaWNrPSgpPT5maWxlLmNsaWNrKCk7ClsiZHJhZ292ZXIiLCJkcmFnbGVhdmUiLCJkcm9wIl0uZm9yRWFjaChlPT5kcm9wLmFkZEV2ZW50TGlzdGVuZXIoZSxldj0+e2V2LnByZXZlbnREZWZhdWx0KCk7ZHJvcC5jbGFzc0xpc3QudG9nZ2xlKCJvdmVyIixlPT09ImRyYWdvdmVyIik7fSkpOwpkcm9wLmFkZEV2ZW50TGlzdGVuZXIoImRyb3AiLGV2PT57aWYoZXYuZGF0YVRyYW5zZmVyLmZpbGVzWzBdKXtjaG9zZW49ZXYuZGF0YVRyYW5zZmVyLmZpbGVzWzBdO3BpY2soKTt9fSk7CmZpbGUub25jaGFuZ2U9KCk9PntpZihmaWxlLmZpbGVzWzBdKXtjaG9zZW49ZmlsZS5maWxlc1swXTtwaWNrKCk7fX07CmZ1bmN0aW9uIHBpY2soKXskKCIjZHJvcHQiKS50ZXh0Q29udGVudD1jaG9zZW4ubmFtZTskKCIjZHJvcHMiKS50ZXh0Q29udGVudD0iUmVhZHkgdG8gYnVpbGQiO2dvLmRpc2FibGVkPWZhbHNlO30KcmVzZXQub25jbGljaz0oKT0+e2Nob3Nlbj1udWxsO2ZpbGUudmFsdWU9IiI7Z28uZGlzYWJsZWQ9dHJ1ZTtyZXNldC5jbGFzc0xpc3QuYWRkKCJoaWRlIik7c3RhdHVzLnRleHRDb250ZW50PSIiO3Jlc3VsdC5pbm5lckhUTUw9IiI7ZmVhdHMuY2xhc3NMaXN0LnJlbW92ZSgiaGlkZSIpOyQoIiNkcm9wdCIpLnRleHRDb250ZW50PSJEcm9wIHlvdXIgc3ByZWFkc2hlZXQgaGVyZSI7JCgiI2Ryb3BzIikuaW5uZXJIVE1MPSJvciBjbGljayB0byBicm93c2UgJm5ic3A7Jm1pZGRvdDsmbmJzcDsgLnhsc3ggb3IgLnhscyI7fTsKY29uc3QgZm10PW49Pm4udG9Mb2NhbGVTdHJpbmcoKTsKZ28ub25jbGljaz1hc3luYygpPT57CiBnby5kaXNhYmxlZD10cnVlO3Jlc3VsdC5pbm5lckhUTUw9IiI7c3RhdHVzLmlubmVySFRNTD0nPHNwYW4gY2xhc3M9InNwaW4iPjwvc3Bhbj4gQnVpbGRpbmcgcGxhbic7CiBjb25zdCBmZD1uZXcgRm9ybURhdGEoKTtmZC5hcHBlbmQoImZpbGUiLGNob3Nlbik7CiB0cnl7CiAgIGNvbnN0IHI9YXdhaXQgZmV0Y2goIi9hcGkvcHJvY2VzcyIse21ldGhvZDoiUE9TVCIsYm9keTpmZH0pOwogICBjb25zdCBqPWF3YWl0IHIuanNvbigpOwogICBpZihqLmVycm9yKXtzdGF0dXMudGV4dENvbnRlbnQ9IiI7cmVzdWx0LmlubmVySFRNTD0nPGRpdiBjbGFzcz0icmNhcmQiPjxkaXYgY2xhc3M9ImVyciI+PGI+Q291bGRudCBwcm9jZXNzIHRoaXMgZmlsZS48L2I+PGJyPicrai5lcnJvcisnPC9kaXY+PC9kaXY+Jztnby5kaXNhYmxlZD1mYWxzZTtyZXR1cm47fQogICBzdGF0dXMudGV4dENvbnRlbnQ9IiI7cmVzZXQuY2xhc3NMaXN0LnJlbW92ZSgiaGlkZSIpO2ZlYXRzLmNsYXNzTGlzdC5hZGQoImhpZGUiKTsKICAgY29uc3Qgcz1qLnN1bW1hcnk7CiAgIGNvbnN0IG1peD1PYmplY3QuZW50cmllcyhzLmVxdWlwbWVudF9taXgpLnNvcnQoKGEsYik9PmJbMV0tYVsxXSkubWFwKChbayx2XSk9Pic8c3BhbiBjbGFzcz0icGlsbDIiPjxiPicrdisnPC9iPicraysnPC9zcGFuPicpLmpvaW4oIiIpOwogICBsZXQgYmFubmVyOwogICBpZihzLnN1cGVybG9hZHMubGVuZ3RoKXtjb25zdCB1PVsuLi5uZXcgU2V0KHMuc3VwZXJsb2Fkcy5tYXAoeD0+IlJvdyAiK3hbMF0rIiAgIit4WzFdKSldOwogICAgIGJhbm5lcj0nPGRpdiBjbGFzcz0iYmFubmVyIHdhcm4iPjxpPiYjOTg4ODs8L2k+PGRpdj48c3BhbiBjbGFzcz0idHRsIj4nK3UubGVuZ3RoKycgcGVybWl0IC8gc3VwZXJsb2FkIGZsYWcnKyh1Lmxlbmd0aD4xPyJzIjoiIikrJzwvc3Bhbj48YnI+Jyt1LmpvaW4oIjxicj4iKSsnPC9kaXY+PC9kaXY+Jzt9CiAgIGVsc2UgYmFubmVyPSc8ZGl2IGNsYXNzPSJiYW5uZXIgb2siPjxpPiYjMTAwMDM7PC9pPjxkaXY+PHNwYW4gY2xhc3M9InR0bCI+Tm8gc3VwZXJsb2Fkcy48L3NwYW4+IE5vdGhpbmcgZXhjZWVkcyAxNiBmdCB3aWRlLCAxMyYjMzk7NiYjMzQ7IHRhbGwsIG9yIHRoZSB3ZWlnaHQgY2VpbGluZy48L2Rpdj48L2Rpdj4nOwogICBjb25zdCBjZWxsPShuLGwpPT4nPGRpdiBjbGFzcz0ic3RhdCI+PGRpdiBjbGFzcz0ibiI+JytuKyc8L2Rpdj48ZGl2IGNsYXNzPSJsIj4nK2wrJzwvZGl2PjwvZGl2Pic7CiAgIHJlc3VsdC5pbm5lckhUTUw9JzxkaXYgY2xhc3M9InJjYXJkIj48ZGl2IGNsYXNzPSJyaGVhZCI+PHNwYW4gY2xhc3M9ImNoZWNrIj4mIzEwMDAzOzwvc3Bhbj4gUGxhbiByZWFkeSAgJytjaG9zZW4ubmFtZSsnPC9kaXY+JysKICAgICAnPGRpdiBjbGFzcz0ic3RhdHMiPicrY2VsbChzLnRydWNrcywiVHJ1Y2tzIikrY2VsbChzLnBpZWNlcywiUGllY2VzIikrY2VsbChzLm92ZXJfd2lkdGhfcGllY2VzLCJPdmVyLXdpZHRoIikrY2VsbChNYXRoLnJvdW5kKHMudG90YWxfd2VpZ2h0X2xiLzEwMDApKyJrIiwiVG90YWwgbGIiKSsnPC9kaXY+JysKICAgICAnPGRpdiBjbGFzcz0ibWl4Ij4nK21peCsnPC9kaXY+JytiYW5uZXIrCiAgICAgKHMuc2tpcHBlZC5jb25zb2xpZGF0ZWQ/JzxwIHN0eWxlPSJjb2xvcjp2YXIoLS1zdWIpO2ZvbnQtc2l6ZToxM3B4O21hcmdpbjowIDAgMTZweCI+JytzLnNraXBwZWQuY29uc29saWRhdGVkKycgY29uc29saWRhdGVkIGxpbmUocykgbGVmdCBibGFuay48L3A+JzoiIikrCiAgICAgJzxhIGNsYXNzPSJkbCIgZG93bmxvYWQ9Iicrai5maWxlbmFtZSsnIiBocmVmPSJkYXRhOmFwcGxpY2F0aW9uL3ZuZC5vcGVueG1sZm9ybWF0cy1vZmZpY2Vkb2N1bWVudC5zcHJlYWRzaGVldG1sLnNoZWV0O2Jhc2U2NCwnK2ouZmlsZSsnIj4mIzExMDE1OyBEb3dubG9hZCB0cnVjayBwbGFuICgueGxzeCk8L2E+PC9kaXY+JzsKIH1jYXRjaChlKXtzdGF0dXMudGV4dENvbnRlbnQ9IiI7cmVzdWx0LmlubmVySFRNTD0nPGRpdiBjbGFzcz0icmNhcmQiPjxkaXYgY2xhc3M9ImVyciI+PGI+RXJyb3IuPC9iPiAnK2UrJzwvZGl2PjwvZGl2Pic7Z28uZGlzYWJsZWQ9ZmFsc2U7fQp9Owo8L3NjcmlwdD4KPC9ib2R5Pgo8L2h0bWw+Cg==")

@app.route('/')
def home():
    return Response(_PAGE, mimetype='text/html')

@app.route('/welcome.png')
def welcome_png():
    p = os.path.join(os.path.dirname(__file__), 'welcome.png')
    return send_file(p, mimetype='image/png') if os.path.exists(p) else ('', 404)

@app.route('/api/process', methods=['POST'])
def api_process():
    f = request.files.get('file')
    if not f:
        return jsonify({'error': 'No file received.'}), 200
    try:
        out, summary = process(f.read(), f.filename or 'manifest.xlsx')
        base = (f.filename or 'manifest.xlsx').rsplit('.', 1)[0]
        return jsonify({'summary': summary,
                        'filename': base + ' JRCTEST.xlsx',
                        'file': _b64.b64encode(out).decode()})
    except Exception as e:
        return jsonify({'error': str(e)}), 200


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8000, debug=True)
