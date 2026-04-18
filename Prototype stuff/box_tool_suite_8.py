"""
Box Tool Suite — All tools in one window
  Tool 1 · Box XML Generator
  Tool 2 · Rate / Count Adjuster
  Tool 3 · NCash Updater (simple CSV)
  Tool 4 · NCash Updater (parent-box CSV + sub-box)
  Tool 5 · NCash ↔ Ticket Calculator

Run: python box_tool_suite.py
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import csv, io, re, os, copy

# ══════════════════════════════════════════════════════════════════════════════
# SHARED — character / options data  (used by Tool 1)
# ══════════════════════════════════════════════════════════════════════════════
CHR_NAMES = ["Bunny","Buffalo","Sheep","Dragon","Fox","Lion","Cat","Raccoon","Paula"]
CHR_JOBS  = ["1st","2nd","3rd"]
CHR_FLAG_MAP = {
    "Bunny 1st":1,"Buffalo 1st":2,"Sheep 1st":4,"Dragon 1st":8,"Fox 1st":16,
    "Lion 1st":32,"Cat 1st":64,"Raccoon 1st":124,"Paula 1st":256,
    "Bunny 2nd":512,"Buffalo 2nd":1024,"Sheep 2nd":2048,"Dragon 2nd":4096,
    "Fox 2nd":8192,"Lion 2nd":16384,"Cat 2nd":32768,"Raccoon 2nd":65536,
    "Paula 2nd":131072,"Bunny 3rd":262144,"Buffalo 3rd":524288,"Sheep 3rd":1048576,
    "Dragon 3rd":2097152,"Fox 3rd":4194304,"Lion 3rd":8388608,"Cat 3rd":16777216,
    "Raccoon 3rd":33554432,"Paula 3rd":67108864,
}
CHR_FLAG_REVERSE = {v:k for k,v in CHR_FLAG_MAP.items()}

# ── Job-name to (race, job_tier) lookup for CSV auto-populate ───────────────
# Each entry: normalised_keyword -> (CHR_NAMES entry, "1st"/"2nd"/"3rd")
_JOB_NAME_MAP = {}
def _reg(*keywords, race, tier):
    for kw in keywords:
        _JOB_NAME_MAP[re.sub(r"[^a-z0-9]","", kw.lower())] = (race, tier)
# Bunny
_reg("schoolgirl","bunnyschoolgirl", race="Bunny", tier="1st")
_reg("boxer","bunnyboxer", race="Bunny", tier="2nd")
_reg("champion","duelist","bunnychampion","bunnyduelist", race="Bunny", tier="3rd")
# Buffalo
_reg("fighter","buffalofighter", race="Buffalo", tier="1st")
_reg("warrior","buffalowarrior", race="Buffalo", tier="2nd")
_reg("mercenary","gladiator","buffalomercenary","buffalogladiator", race="Buffalo", tier="3rd")
# Sheep
_reg("librarian","sheeplibrarian", race="Sheep", tier="1st")
_reg("bard","sheepbard", race="Sheep", tier="2nd")
_reg("witch","soulmaster","sheepwitch","sheepsoulmaster", race="Sheep", tier="3rd")
# Dragon
_reg("shaman","dragonshaman", race="Dragon", tier="1st")
_reg("magician","dragonmagician", race="Dragon", tier="2nd")
_reg("priest","darklord","wizard","dragonpriest","dragondarklord","dragonwizard", race="Dragon", tier="3rd")
# Fox
_reg("foxfirst","fox1st", race="Fox", tier="1st")
_reg("foxsecond","fox2nd", race="Fox", tier="2nd")
_reg("foxthird","fox3rd", race="Fox", tier="3rd")
# Lion
_reg("lionfirst","lion1st", race="Lion", tier="1st")
_reg("lionsecond","lion2nd", race="Lion", tier="2nd")
_reg("lionthird","lion3rd", race="Lion", tier="3rd")
# Cat
_reg("model","catmodel", race="Cat", tier="1st")
_reg("entertainer","catentertainer", race="Cat", tier="2nd")
_reg("primadonna","diva","catprimadonna","catdiva", race="Cat", tier="3rd")
# Raccoon
_reg("teacher","raccoonteacher", race="Raccoon", tier="1st")
_reg("cardmaster","raccooncardmaster", race="Raccoon", tier="2nd")
_reg("gambler","duke","raccoongambler","raccounduke", race="Raccoon", tier="3rd")
# Paula / Polar Bear / Bear
_reg("animallover","paulaanimallover","polarbearanimallover", race="Paula", tier="1st")
_reg("trainer","paulatrainer","polarbeartrainer", race="Paula", tier="2nd")
_reg("zoologist","paulazoologist","polarbearzoologist", race="Paula", tier="3rd")

# Race aliases: bear / polarbear / paula all -> "Paula"
_RACE_ALIAS = {}
for aliases, canon in [
    (["bunny"], "Bunny"), (["buffalo"], "Buffalo"),
    (["sheep"], "Sheep"), (["dragon"], "Dragon"),
    (["fox"], "Fox"), (["lion"], "Lion"), (["cat"], "Cat"),
    (["raccoon"], "Raccoon"),
    (["paula","bear","polarbear","polar bear","polar"], "Paula"),
]:
    for a in aliases: _RACE_ALIAS[re.sub(r"[^a-z]","", a.lower())] = canon

# Tier aliases
_TIER_ALIAS = {}
for aliases, canon in [
    (["1st","first","1","job1","firstjob","1stjob","job1st"], "1st"),
    (["2nd","second","2","job2","secondjob","2ndjob","job2nd"], "2nd"),
    (["3rd","third","3","job3","thirdjob","3rdjob","job3rd"], "3rd"),
]:
    for a in aliases: _TIER_ALIAS[re.sub(r"[^a-z0-9]","", a.lower())] = canon

def resolve_chr_flag(text):
    """Try to resolve free-text like '2nd job Sheep' or 'Witch' to a CHR_FLAG_MAP value.
    Returns int flag or None."""
    raw = text.strip()
    norm = re.sub(r"[^a-z0-9]", "", raw.lower())
    # Direct job name lookup
    if norm in _JOB_NAME_MAP:
        race, tier = _JOB_NAME_MAP[norm]
        return CHR_FLAG_MAP.get(f"{race} {tier}")
    # Try to parse "tier race" or "race tier" patterns
    # Strip common joiners
    parts = re.split(r"[-–—/\s]+", raw.lower())
    parts = [re.sub(r"[^a-z0-9]","", p) for p in parts if p.strip()]
    found_race = found_tier = None
    for p in parts:
        if p in _RACE_ALIAS: found_race = _RACE_ALIAS[p]
        if p in _TIER_ALIAS: found_tier = _TIER_ALIAS[p]
    # Also try each part as a job name (e.g. "witch")
    if not found_race or not found_tier:
        for p in parts:
            if p in _JOB_NAME_MAP:
                r, t = _JOB_NAME_MAP[p]
                if not found_race: found_race = r
                if not found_tier: found_tier = t
    if found_race and found_tier:
        return CHR_FLAG_MAP.get(f"{found_race} {found_tier}")
    return None
OPTIONS_CHECKS = [
    ("Not Buyable",256),("Not Sellable",512),("Not Exchangeable",1024),
    ("Not Pickable",2048),("Not Droppable",4096),("Not Vanishable",8192),
    ("No Angelina Bank",65536),("No Lisa Bank",131072),
]

# ══════════════════════════════════════════════════════════════════════════════
# SHARED — XML regex  (used by Tools 1/2/3/4)
# ══════════════════════════════════════════════════════════════════════════════
ROW_RE   = re.compile(r'<ROW>.*?</ROW>', re.DOTALL)
CDATA_RE = re.compile(r'<!\[CDATA\[(.*?)\]\]>', re.DOTALL)

def _get_tag(block, tag):
    m = re.search(rf'<{re.escape(tag)}>(.*?)</{re.escape(tag)}>', block, re.DOTALL)
    if not m: return ""
    cd = CDATA_RE.search(m.group(1))
    return cd.group(1).strip() if cd else m.group(1).strip()

def _set_tag(block, tag, val):
    return re.sub(rf'<{re.escape(tag)}>.*?</{re.escape(tag)}>',
                  f'<{tag}>{val}</{tag}>', block, flags=re.DOTALL)

# ══════════════════════════════════════════════════════════════════════════════
# SHARED — ItemParam / NCash helpers  (Tools 1/3/4)
# ══════════════════════════════════════════════════════════════════════════════
TARGET_FILES = {"itemparam2.xml","itemparamcm2.xml","itemparamex2.xml","itemparamex.xml"}
PRESENT_FILE = "presentitemparam2.xml"

def build_item_lib(files):
    lib = {}
    for _, text in files:
        for row in ROW_RE.findall(text):
            rid  = _get_tag(row, "ID")
            name = _get_tag(row, "Name")
            if rid.isdigit() and name:
                lib[rid] = name
    return lib

def bulk_update_ncash(xml_text, updates):
    found = {k: False for k in updates}
    def replace_row(m):
        block = m.group(0)
        rid   = _get_tag(block, "ID")
        if rid not in updates: return block
        found[rid] = True
        return re.sub(r'<Ncash>\d+</Ncash>', f'<Ncash>{updates[rid]}</Ncash>', block)
    return ROW_RE.sub(replace_row, xml_text), found

# ══════════════════════════════════════════════════════════════════════════════
# SHARED — Tool 1 CSV + XML builders
# ══════════════════════════════════════════════════════════════════════════════
# ── Column headers that are NEVER box-name columns ──────────────────────────
# ── Column-header sets for CSV field recognition ─────────────────────────────
_SKIP_HEADERS = {
    "id","level","rate","lv","luck","lvl","chance","prob","itemcnt","count",
    "droprate","dropcnt","dropid","drop","qty","quantity","amount","weight",
}

def _norm_hdr(h):
    return re.sub(r"[^a-z0-9]", "", h.strip().lower())

# Headers that identify box-name columns (the column header IS the box name)
# A column is a box-name column when its header is non-empty, not a known skip
# word, not purely numeric, and NOT a recognised field/metadata header.
_FIELD_HDRS = {
    "name","comment","use","filename","fn","bundlenum","bn",
    "cmtfilename","cmtfn","cmtbundlenum","cmtbn",
    "weight","value","minlevel","money","ncash","tickets","ticket",
    "options","chrtypeflags","chrtypeflag","recycle","recyclable",
    "boxid","contents","content","id",
}

def _is_box_name_header(h):
    hn = _norm_hdr(h)
    if not hn: return False
    if hn.isdigit(): return False
    if hn in _SKIP_HEADERS: return False
    if hn in _FIELD_HDRS: return False
    return True

def _parse_options_cell(cell):
    """Parse e.g. 'Not Buyable, Not Sellable' → (opt_checks list, recycle_int)."""
    opt_checks = [False] * 8
    recycle = 0
    tokens = [re.sub(r"[^a-z0-9]", "", t.lower()) for t in re.split(r"[,;|]+", cell) if t.strip()]
    for tok in tokens:
        for i, (lbl, _) in enumerate(OPTIONS_CHECKS):
            if re.sub(r"[^a-z0-9]", "", lbl.lower()) in tok or tok in re.sub(r"[^a-z0-9]", "", lbl.lower()):
                opt_checks[i] = True
        if tok in ("recyclable", "recycle", "recyc"):                         recycle = 262144
        if tok in ("nonrecyclable", "norecycle", "nonrecycle", "notrecyclable"): recycle = 8388608
    return opt_checks, recycle

def _parse_chrtypeflags_cell(cell):
    """Parse comma/semicolon separated ChrTypeFlags text → list of flag ints."""
    flags = []
    for part in re.split(r"[,;|/]+", cell):
        part = part.strip()
        if not part: continue
        if part.isdigit(): flags.append(int(part)); continue
        f = resolve_chr_flag(part)
        if f is not None: flags.append(f)
    return flags

# Map normalised header → cfg_override key (special keys start with _)
_HDR_TO_CFGKEY = {
    "name":         "name",
    "comment":      "comment",
    "use":          "use",
    "filename":     "file_name",   "fn": "file_name",
    "bundlenum":    "bundle_num",  "bn": "bundle_num",
    "cmtfilename":  "cmt_file_name", "cmtfn": "cmt_file_name",
    "cmtbundlenum": "cmt_bundle_num", "cmtbn": "cmt_bundle_num",
    "weight":       "weight",
    "value":        "value",
    "minlevel":     "min_level",
    "money":        "money",
    "ncash":        "_ncash",      # handled specially
    "tickets":      "_tickets",    # handled specially
    "ticket":       "_tickets",
    "options":      "_options",
    "chrtypeflags": "_chr",
    "chrtypeflag":  "_chr",
    "recycle":      "_recycle",
    "recyclable":   "_recycle",
    "boxid":        "_boxid",      # box's own ID
    "id":           "_id",         # ambiguous – could be box ID or item ID
    "contents":     "_contents",
    "content":      "_contents",
}

def _apply_field_col(cfg_override, key, val):
    """Apply a parsed field value to cfg_override dict."""
    if key == "_options":
        oc, rv = _parse_options_cell(val)
        cfg_override["opt_checks"] = oc
        if rv: cfg_override["opt_recycle"] = rv
    elif key == "_chr":
        flags = _parse_chrtypeflags_cell(val)
        if flags: cfg_override["chr_type_flags"] = flags
    elif key == "_recycle":
        nv = _norm_hdr(val)
        if nv in ("recyclable","recycle","recyc"):                       cfg_override["opt_recycle"] = 262144
        elif nv in ("nonrecyclable","norecycle","nonrecycle","notrecyclable"): cfg_override["opt_recycle"] = 8388608
    elif key == "_ncash":
        try: cfg_override["ncash_direct"] = int(round(float(val)))
        except: pass
    elif key == "_tickets":
        try: cfg_override["ticket"] = val
        except: pass
    elif not key.startswith("_"):
        cfg_override[key] = val

def parse_grouped_csv(text):
    """Parse a box-definition CSV.

    Layout rules (same as original, with additions):
    - Header row: columns whose headers are not recognised field/skip words are
      "box-name columns".  Each defines a GROUP.
    - A group's span is: columns from (previous group's box-name col + 1) up to
      and including its own box-name col.
    - Within each span:
        * The box-name col: each data row's cell is an item NAME (or box name if
          a "Box ID" / "Box …" column is present).
        * Any recognised field header (Name, Comment, Options, ChrTypeFlags,
          BoxID, Contents, …) is parsed and stored in cfg_override.
        * Numeric middle columns are item DropRate.
        * "ItemCnt" / "Count" column sets item_cnt.
        * "Contents" column overrides item name source.
        * "Box ID" or "BoxID" column sets the box's own ID in cfg_override.
    - Returns list of {box_name, items, cfg_override} dicts.
    """
    reader = csv.reader(io.StringIO(text))
    rows   = list(reader)
    if not rows: return []
    headers = [h.strip() for h in rows[0]]
    hnorms  = [_norm_hdr(h) for h in headers]
    data_rows = rows[1:]

    # Identify box-name columns
    box_col_indices = [i for i, h in enumerate(headers) if _is_box_name_header(h)]
    if not box_col_indices: return []

    results = []
    prev_end = -1
    # Compute next box-col for each group (to know where trailing fields end)
    next_box_cols = box_col_indices[1:] + [len(headers)]

    for bi, bc in enumerate(box_col_indices):
        # Core span: prev_end+1 .. bc  (the original item columns)
        # Trailing span: bc+1 .. next_box_col-1 BUT only recognised field headers
        next_bc = next_box_cols[bi]
        _TRAILING_OK = {k for k,v in _HDR_TO_CFGKEY.items()
                        if v not in ("_id",) and k not in _SKIP_HEADERS}
        trailing = [ci for ci in range(bc + 1, next_bc)
                    if _norm_hdr(headers[ci]) in _TRAILING_OK]
        span = list(range(prev_end + 1, bc + 1)) + trailing
        prev_end = bc

        box_name_hdr = headers[bc]   # the column header = box name for the group

        # Classify every column in the span
        id_col       = None   # item ID column index (per-item)
        box_id_col   = None   # box's own ID column index
        contents_col = None   # alternate item-name column
        rate_cols    = []     # numeric rate columns
        itemcnt_col  = None
        field_cols   = {}     # col_index -> cfg_key (field overrides)

        for ci in span:
            if ci == bc: continue   # box-name col itself
            hn = hnorms[ci]
            if hn in _HDR_TO_CFGKEY:
                cfgkey = _HDR_TO_CFGKEY[hn]
                if cfgkey == "_boxid":
                    box_id_col = ci
                elif cfgkey == "_id":
                    # bare "ID" column — treat as item ID
                    id_col = ci
                elif cfgkey == "_contents":
                    contents_col = ci
                else:
                    field_cols[ci] = cfgkey
            elif hn in ("itemcnt", "count"):
                itemcnt_col = ci
            elif hn in _SKIP_HEADERS:
                rate_cols.append(ci)   # treat rate/level/etc as rate source
            else:
                rate_cols.append(ci)   # unknown middle col → try as rate

        # Build cfg_override from field columns (take first non-empty value per field)
        group_cfg = {}
        for ci, cfgkey in field_cols.items():
            for row in data_rows:
                val = row[ci].strip() if ci < len(row) else ""
                if val:
                    _apply_field_col(group_cfg, cfgkey, val)
                    break

        # Box's own ID from box_id_col (first non-empty)
        if box_id_col is not None:
            for row in data_rows:
                val = row[box_id_col].strip() if box_id_col < len(row) else ""
                if val.isdigit():
                    group_cfg["id"] = val
                    break

        # Build items list
        items = []
        for row in data_rows:
            # Item name comes from box-name col (or contents_col if present)
            raw_name = row[bc].strip() if bc < len(row) else ""
            if not raw_name: continue

            item_name = raw_name
            if contents_col is not None:
                item_name = row[contents_col].strip() if contents_col < len(row) else raw_name

            # Item ID
            item_id = ""
            if id_col is not None:
                v = row[id_col].strip() if id_col < len(row) else ""
                if v.isdigit(): item_id = v

            # Rate: first parseable integer from rate_cols
            rate = None
            for ci in rate_cols:
                v = row[ci].strip() if ci < len(row) else ""
                try: rate = int(v.replace("%","").strip()); break
                except: pass

            # ItemCnt
            item_cnt = 1
            if itemcnt_col is not None:
                v = row[itemcnt_col].strip() if itemcnt_col < len(row) else ""
                try: item_cnt = int(v)
                except: pass

            items.append({"id": item_id, "name": item_name,
                          "extra": [], "rate": rate, "item_cnt": item_cnt})

        if items:
            results.append({"box_name": box_name_hdr,
                            "items": items,
                            "cfg_override": group_cfg})

    return results

def substitute_box_name(template, old, new):
    if not old or not template: return template
    return re.sub(re.escape(old), new, template, flags=re.IGNORECASE)

def apply_name_template(template, prev, new):
    if not template or not prev: return new
    return substitute_box_name(template, prev, new)

def deduplicate_name(proposed, existing):
    if proposed not in existing: return proposed
    i = 2
    while f"{proposed} ({i})" in existing: i += 1
    return f"{proposed} ({i})"

def build_options_str(check_states, recycle_val):
    base    = [2,32]
    checked = [v for (_,v),on in zip(OPTIONS_CHECKS,check_states) if on]
    rec     = [recycle_val] if recycle_val > 0 else []
    return "/".join(str(x) for x in sorted(set(base+checked+rec)))

def build_itemparam_row(cfg):
    opts      = build_options_str(cfg["opt_checks"], cfg["opt_recycle"])
    chr_flags = sum(cfg["chr_type_flags"])
    fn        = cfg["file_name"]
    bn        = cfg["bundle_num"]
    return f"""<ROW>
<ID>{cfg['id']}</ID>
<Class>1</Class>
<Type>15</Type>
<SubType>0</SubType>
<ItemFType>0</ItemFType>
<n><![CDATA[{cfg['name']}]]></n>
<Comment><![CDATA[{cfg['comment']}]]></Comment>
<Use><![CDATA[{cfg['use']}]]></Use>
<Name_Eng><![CDATA[ ]]></Name_Eng>
<Comment_Eng><![CDATA[ ]]></Comment_Eng>
<FileName><![CDATA[{fn}]]></FileName>
<BundleNum>{bn}</BundleNum>
<InvFileName><![CDATA[{fn}]]></InvFileName>
<InvBundleNum>{bn}</InvBundleNum>
<CmtFileName><![CDATA[{cfg['cmt_file_name']}]]></CmtFileName>
<CmtBundleNum>{cfg['cmt_bundle_num']}</CmtBundleNum>
<EquipFileName><![CDATA[ ]]></EquipFileName>
<PivotID>0</PivotID>
<PaletteId>0</PaletteId>
<Options>{opts}</Options>
<HideHat>0</HideHat>
<ChrTypeFlags>{chr_flags}</ChrTypeFlags>
<GroundFlags>0</GroundFlags>
<SystemFlags>0</SystemFlags>
<OptionsEx>0</OptionsEx>
<Weight>{cfg['weight']}</Weight>
<Value>{cfg['value']}</Value>
<MinLevel>{cfg['min_level']}</MinLevel>
<Effect>22</Effect>
<EffectFlags2>0</EffectFlags2>
<SelRange>0</SelRange>
<Life>0</Life>
<Depth>0</Depth>
<Delay>0.000000</Delay>
<AP>0</AP>
<HP>0</HP>
<HPCon>0</HPCon>
<MP>0</MP>
<MPCon>0</MPCon>
<Money>{cfg['money']}</Money>
<APPlus>0</APPlus>
<ACPlus>0</ACPlus>
<DXPlus>0</DXPlus>
<MaxMPPlus>0</MaxMPPlus>
<MAPlus>0</MAPlus>
<MDPlus>0</MDPlus>
<MaxWTPlus>0</MaxWTPlus>
<DAPlus>0</DAPlus>
<LKPlus>0</LKPlus>
<MaxHPPlus>0</MaxHPPlus>
<DPPlus>0</DPPlus>
<HVPlus>0</HVPlus>
<HPRecoveryRate>0.000000</HPRecoveryRate>
<MPRecoveryRate>0.000000</MPRecoveryRate>
<CardNum>0</CardNum>
<CardGenGrade>0</CardGenGrade>
<CardGenParam>0.000000</CardGenParam>
<DailyGenCnt>0</DailyGenCnt>
<PartFileName><![CDATA[ ]]></PartFileName>
<ChrFTypeFlag>0</ChrFTypeFlag>
<ChrGender>0</ChrGender>
<ExistType>0</ExistType>
<Ncash>{cfg['ncash']}</Ncash>
<NewCM>0</NewCM>
<FamCM>0</FamCM>
<Summary><![CDATA[ ]]></Summary>
<ShopFileName><![CDATA[ ]]></ShopFileName>
<ShopBundleNum>0</ShopBundleNum>
<MinStatType>0</MinStatType>
<MinStatLv>0</MinStatLv>
<RefineIndex>0</RefineIndex>
<RefineType>0</RefineType>
<CompoundSlot>0</CompoundSlot>
<SetItemID>0</SetItemID>
<ReformCount>0</ReformCount>
<GroupId>0</GroupId>
</ROW>"""

def build_presentparam_row(box_id, items, ptype, drop_cnt, default_rate,
                           item_cnts=None, box_name=None):
    """box_name appended as <!-- name --> after <Id>.
    items[i]["name"] appended as <!-- name --> after each <DropId_#>."""
    is_distrib      = (ptype == 2)
    actual_drop_cnt = len(items) if is_distrib else drop_cnt
    id_line = f"<Id>{box_id}</Id>"
    if box_name:
        id_line += f" <!-- {box_name} -->"
    lines = ["<ROW>", id_line,
             f"<Type>{ptype}</Type>", f"<DropCnt>{actual_drop_cnt}</DropCnt>"]
    for i in range(20):
        if i < len(items):
            did   = items[i]["id"]
            drate = 100 if is_distrib else (items[i].get("rate") or default_rate)
            icnt  = (item_cnts[i] if item_cnts and i < len(item_cnts) else 1) or 1
            item_name = items[i].get("name", "").strip()
            drop_id_line = f"<DropId_{i}>{did}</DropId_{i}>"
            if item_name:
                drop_id_line += f" <!-- {item_name} -->"
        else:
            did, drate, icnt = 0, 0, 0
            drop_id_line = f"<DropId_{i}>{did}</DropId_{i}>"
        lines += [drop_id_line,
                  f"<DropRate_{i}>{drate}</DropRate_{i}>",
                  f"<ItemCnt_{i}>{icnt}</ItemCnt_{i}>"]
    lines.append("</ROW>")
    return "\n".join(lines)

def build_recycle_except_row(box_id, name):
    return f"<ROW>\n<ItemID>{box_id}</ItemID>\n<Comment><![CDATA[{name}]]></Comment>\n</ROW>"

# ══════════════════════════════════════════════════════════════════════════════
# SHARED — Tool 2 XML helpers
# ══════════════════════════════════════════════════════════════════════════════
def real_drop_slots(block):
    pairs = re.findall(r'<DropId_(\d+)>(\d+)</DropId_\d+>', block)
    return [(int(i),v) for i,v in sorted(pairs,key=lambda x:int(x[0])) if v!="0"]

def apply_cfg_to_row(block, cfg):
    block = _set_tag(block,"Type",str(cfg["type"]))
    block = _set_tag(block,"DropCnt",str(cfg["drop_cnt"]))
    for pos,(idx,_) in enumerate(real_drop_slots(block)):
        sc = cfg["slots"][pos] if pos<len(cfg["slots"]) else {"rate":100,"count":1}
        block = _set_tag(block,f"DropRate_{idx}",str(sc["rate"]))
        block = _set_tag(block,f"ItemCnt_{idx}",str(sc["count"]))
    return block

def load_itemparam_folder(folder):
    lib = {}
    for fname in os.listdir(folder):
        if not fname.lower().endswith(".xml"): continue
        try:
            with open(os.path.join(folder,fname),encoding="utf-8-sig",errors="replace") as f:
                text = f.read()
            for row in ROW_RE.findall(text):
                rid  = _get_tag(row,"ID")
                name = _get_tag(row,"n")
                if rid.isdigit() and name: lib[rid] = name
        except: pass
    return lib

# ══════════════════════════════════════════════════════════════════════════════
# SHARED — Tool 2 CSV parser
# ══════════════════════════════════════════════════════════════════════════════
_T2_SKIP = {"id","level","rate","lv","luck","lvl","chance","prob"}

def parse_box_id_csv(text):
    reader  = csv.reader(io.StringIO(text))
    rows    = list(reader)
    if not rows: return {}
    headers = [h.strip() for h in rows[0]]
    id_positions = [i for i,h in enumerate(headers) if h.strip().lower()=="id"]
    if not id_positions: id_positions = [0]
    box_map = {}
    for g,id_pos in enumerate(id_positions):
        next_id = id_positions[g+1] if g+1<len(id_positions) else len(headers)
        gcols   = list(range(id_pos,next_id))
        ghdrs   = [headers[c] for c in gcols]
        name_local = next((i for i,h in enumerate(ghdrs)
                           if bool(h) and h.strip().lower() not in _T2_SKIP and not h.strip().isdigit()),None)
        for row in rows[1:]:
            id_val = row[id_pos].strip() if id_pos<len(row) else ""
            if not id_val or not id_val.isdigit(): continue
            if name_local is not None:
                nc = gcols[name_local]
                name_val = row[nc].strip() if nc<len(row) else ""
            else:
                name_val = ""
            box_map[id_val] = name_val
    return box_map

# ══════════════════════════════════════════════════════════════════════════════
# SHARED — Tool 3 CSV parser
# ══════════════════════════════════════════════════════════════════════════════
def parse_csv_text_t3(text):
    stripped = text.strip()
    if not stripped: return []
    all_rows = list(csv.reader(io.StringIO(stripped)))
    if not all_rows: return []
    raw_headers = [h.strip() for h in all_rows[0]]
    data_rows   = all_rows[1:]
    items, seen = [], set()
    def add(id_str, cost=None):
        id_str = id_str.strip()
        if id_str and id_str.isdigit() and id_str not in seen:
            seen.add(id_str); items.append({"id":id_str,"ticket_cost":cost})
    item_col_positions = [i for i,h in enumerate(raw_headers) if re.match(r'Item\d+_ID',h,re.I)]
    if item_col_positions:
        for row in data_rows:
            for pos in item_col_positions:
                if pos<len(row): add(row[pos])
        return items
    id_positions = [i for i,h in enumerate(raw_headers) if h.lower()=="id"]
    if id_positions:
        for row in data_rows:
            for pos in id_positions:
                if pos<len(row): add(row[pos])
        return items
    if len(raw_headers)>=2:
        for row in data_rows:
            if not row: continue
            raw_cost = row[1].strip() if len(row)>1 else ""
            try:    cost = float(raw_cost)
            except: cost = None
            add(row[0], cost)
        return items
    for row in data_rows:
        if row: add(row[0])
    return items

# ══════════════════════════════════════════════════════════════════════════════
# SHARED — Tool 4 CSV parser
# ══════════════════════════════════════════════════════════════════════════════
_NON_ID_HEADERS = {
    "acplus","ap","applus","bundlenum","cardgengrade","cardgenparam","cardnum",
    "chrftypeflag","chrgender","chrtypeflags","class","cmtbundlenum","cmtfilename",
    "comment","comment_eng","compoundslot","daplus","dpplus","dxplus","dailygencnt",
    "delay","depth","effect","effectflags2","equipfilename","existtype","famcm",
    "filename","groundflags","groupid","hp","hpcon","hprecoveryrate","hvplus",
    "hidehat","invbundlenum","invfilename","itemftype","lkplus","life","maplus",
    "mdplus","mp","mpcon","mprecoveryrate","maxhpplus","maxmpplus","maxwtplus",
    "minlevel","minstatLv","minstattype","money","name","name_eng","newcm",
    "options","optionsex","paletteid","partfilename","pivotid","refineindex",
    "refinetype","reformcount","selrange","setitemid","shopbundlenum","shopfilename",
    "subtype","summary","systemflags","type","use","value","weight",
    "level","rate","lv","luck","lvl","chance","prob","drop","qty",
    "quantity","count","amount","row",
    "droprate","dropcnt","dropid","itemcnt","dropcnt_0","dropcnt_1",
    "droprate_0","droprate_1","droprate_2","droprate_3","droprate_4",
    "dropid_0","dropid_1","dropid_2","dropid_3","dropid_4",
    "itemcnt_0","itemcnt_1","itemcnt_2","itemcnt_3","itemcnt_4",
}
_TICKET_NAMES     = {"tickets","ticket"}
_NCASH_NAMES      = {"ncash","ncash_val","ncashval"}
_BOX_TICKET_NAMES = {
    "tickets of box contents","box contents tickets",
    "box content tickets","tickets of box content",
    "sub-box tickets","subbox tickets",
}

def _find_value_col(raw_headers, id_pos):
    next_id = next((i for i in range(id_pos+1,len(raw_headers))
                    if raw_headers[i].lower()=="id"), len(raw_headers))
    for i in range(id_pos+1, next_id):
        h = raw_headers[i].lower()
        if h in _TICKET_NAMES: return i,"tickets"
        if h in _NCASH_NAMES:  return i,"ncash"
    return None, None

def _find_box_ticket_col(raw_headers, id_pos):
    next_id = next((i for i in range(id_pos+1,len(raw_headers))
                    if raw_headers[i].lower()=="id"), len(raw_headers))
    for i in range(id_pos+1, next_id):
        if raw_headers[i].lower().strip() in _BOX_TICKET_NAMES: return i
    return None

def parse_parentbox_csv(text):
    stripped = text.strip()
    if not stripped: return []
    all_rows = list(csv.reader(io.StringIO(stripped)))
    if not all_rows: return []
    raw_headers = [h.strip() for h in all_rows[0]]
    data_rows   = all_rows[1:]
    id_positions = [i for i,h in enumerate(raw_headers) if h.lower()=="id"]
    val_map, box_tick_map = {}, {}
    for id_pos in id_positions:
        vcol,vtype = _find_value_col(raw_headers, id_pos)
        if vcol is not None: val_map[id_pos] = (vcol,vtype)
        btcol = _find_box_ticket_col(raw_headers, id_pos)
        if btcol is not None: box_tick_map[id_pos] = btcol
    items, seen = [], set()
    def add(id_str, ticket_cost, ncash_direct, box_ticket_cost, group_idx=0):
        id_str = id_str.strip()
        if id_str and id_str.isdigit() and id_str not in seen:
            seen.add(id_str)
            items.append({"id":id_str,"ticket_cost":ticket_cost,"ncash_direct":ncash_direct,
                          "box_ticket_cost":box_ticket_cost,"group_idx":group_idx,"name":""})
    def _parse_num(row, col):
        if col is None or col>=len(row): return None
        try:    return float(row[col].strip())
        except: return None
    if id_positions:
        for row in data_rows:
            for gi,id_pos in enumerate(id_positions):
                if id_pos>=len(row): continue
                id_val = row[id_pos].strip()
                if not (id_val and id_val.isdigit()): continue
                ticket_cost = ncash_direct = None
                if id_pos in val_map:
                    vcol,vtype = val_map[id_pos]
                    num = _parse_num(row, vcol)
                    if num is not None:
                        if vtype=="tickets": ticket_cost  = num
                        else:               ncash_direct = int(round(num))
                btcol = box_tick_map.get(id_pos)
                box_ticket_cost = _parse_num(row,btcol) if btcol is not None else None
                add(id_val, ticket_cost, ncash_direct, box_ticket_cost, group_idx=gi)
        return items
    for row in data_rows:
        for i,cell in enumerate(row):
            hdr = raw_headers[i].lower() if i<len(raw_headers) else ""
            if hdr not in _NON_ID_HEADERS:
                add(cell, None, None, None)
    return items

def extract_drop_ids_from_present(present_text, box_ids):
    result = {}
    for row in ROW_RE.findall(present_text):
        bid = _get_tag(row,"Id")
        if bid not in box_ids: continue
        drops = []
        for i in range(20):
            did = _get_tag(row,f"DropId_{i}")
            if did and did.isdigit() and did!="0": drops.append(did)
        result[bid] = drops
    return result


# ══════════════════════════════════════════════════════════════════════════════
# SESSION STORE  — shared state between tools
# ══════════════════════════════════════════════════════════════════════════════
class AppSession:
    """Central store. Tool 1 writes here; Tools 2-4 can import from it."""
    def __init__(self):
        self.box_id_list_csv  = None   # Tool1 output: box IDs CSV text
        self.box_id_map       = {}     # Tool1 output: {box_id: box_name}
        self.box_contents_csv = None   # Tool2 output (future)
        self.present_xml_path = None

# ══════════════════════════════════════════════════════════════════════════════
# UI HELPERS — shared across all tool frames
# ══════════════════════════════════════════════════════════════════════════════
BG      = "#1e1e2e"
BG2     = "#181825"
BG3     = "#313244"
BG4     = "#45475a"
FG      = "#cdd6f4"
FG_DIM  = "#a6adc8"
FG_GREY = "#6c7086"
ACC1    = "#cba6f7"   # purple  — tool 1
ACC2    = "#89dceb"   # cyan    — tool 2
ACC3    = "#f38ba8"   # red     — tool 3
ACC4    = "#fab387"   # peach   — tool 4
ACC5    = "#f9e2af"   # yellow  — calculator
GREEN   = "#a6e3a1"
BLUE    = "#89b4fa"

def mk_section(parent, title):
    f = tk.LabelFrame(parent, text=title, bg=BG, fg=BLUE,
                      font=("Consolas",10,"bold"), bd=1, relief="groove")
    f.pack(fill="x", padx=12, pady=5)
    return f

def mk_btn(parent, text, command, color=BG3, fg=FG, **kw):
    return tk.Button(parent, text=text, command=command, bg=color, fg=fg,
                     font=("Consolas",10), relief="flat", padx=12, pady=6, **kw)

def mk_scroll_canvas(parent, init_width=900):
    """Returns (canvas, container_frame). Container scrolls inside canvas.
    init_width: initial width hint so content renders before the first Configure event."""
    sb     = ttk.Scrollbar(parent, orient="vertical")
    canvas = tk.Canvas(parent, bg=BG, highlightthickness=0,
                       yscrollcommand=sb.set, width=init_width)
    sb.configure(command=canvas.yview)
    sb.pack(side="right", fill="y")
    canvas.pack(side="left", fill="both", expand=True)
    cont = tk.Frame(canvas, bg=BG)
    wid  = canvas.create_window((0,0), window=cont, anchor="nw", width=init_width)

    def _on_inner_resize(e):
        canvas.configure(scrollregion=canvas.bbox("all"))

    def _on_canvas_resize(e):
        canvas.itemconfig(wid, width=e.width)

    cont.bind("<Configure>", _on_inner_resize)
    canvas.bind("<Configure>", _on_canvas_resize)

    def _on_enter(e): canvas.bind_all("<MouseWheel>", lambda ev: canvas.yview_scroll(-1*(ev.delta//120),"units"))
    def _on_leave(e): canvas.unbind_all("<MouseWheel>")
    canvas.bind("<Enter>", _on_enter)
    canvas.bind("<Leave>", _on_leave)

    return canvas, cont

class ScrolledFrame(tk.Frame):
    """A frame with a vertical scrollbar. Place children in .inner.
    Uses a canvas with an explicit initial width so content renders immediately."""
    def __init__(self, parent, **kw):
        super().__init__(parent, **kw)
        self.configure(bg=BG)
        self._canvas = tk.Canvas(self, bg=BG, highlightthickness=0,
                                 borderwidth=0, width=900)
        self._sb = ttk.Scrollbar(self, orient="vertical", command=self._canvas.yview)
        self._canvas.configure(yscrollcommand=self._sb.set)
        self._sb.pack(side="right", fill="y")
        self._canvas.pack(side="left", fill="both", expand=True)
        self.inner = tk.Frame(self._canvas, bg=BG)
        self._win = self._canvas.create_window((0, 0), window=self.inner,
                                               anchor="nw", width=900)
        self.inner.bind("<Configure>", self._on_inner)
        self._canvas.bind("<Configure>", self._on_canvas)
        self._canvas.bind("<Enter>",
            lambda e: self._canvas.bind_all("<MouseWheel>", self._scroll))
        self._canvas.bind("<Leave>",
            lambda e: self._canvas.unbind_all("<MouseWheel>"))

    def _on_inner(self, e):
        self._canvas.configure(scrollregion=self._canvas.bbox("all"))

    def _on_canvas(self, e):
        self._canvas.itemconfig(self._win, width=e.width)

    def _scroll(self, e):
        self._canvas.yview_scroll(-1 * (e.delta // 120), "units")

def make_output_tab(nb, title, content, fname, root):
    frm = tk.Frame(nb, bg=BG); nb.add(frm, text=title)
    br  = tk.Frame(frm, bg=BG); br.pack(side="bottom", fill="x")
    mk_btn(br, "📋 Copy All",
           lambda c=content: (root.clipboard_clear(), root.clipboard_append(c),
                               messagebox.showinfo("Copied","Copied to clipboard.")),
           padx=10, pady=4).pack(side="left", padx=6, pady=4)
    def _save(c=content, f=fname):
        p = filedialog.asksaveasfilename(initialfile=f, defaultextension=".xml",
                filetypes=[("XML","*.xml"),("CSV","*.csv"),("Text","*.txt"),("All","*.*")])
        if p:
            with open(p,"w",encoding="utf-8") as fh: fh.write(c)
            messagebox.showinfo("Saved",f"Saved to {p}")
    mk_btn(br,"💾 Save As…",_save,color=GREEN,fg=BG2,padx=10,pady=4).pack(side="left",padx=6,pady=4)
    txt = scrolledtext.ScrolledText(frm, font=("Consolas",9), bg=BG2, fg=FG)
    txt.pack(fill="both", expand=True, padx=4, pady=4)
    txt.insert("1.0", content); txt.config(state="disabled")

# ══════════════════════════════════════════════════════════════════════════════
# TOOL 1 — Box XML Generator
# ══════════════════════════════════════════════════════════════════════════════
class Tool1(tk.Frame):
    def __init__(self, parent, root, session):
        super().__init__(parent, bg=BG)
        self.root = root; self.session = session
        self.saved_settings = None
        self.continue_mode  = None
        self.groups         = []
        self.current_group_idx = 0
        self.box_configs    = []
        self.no_csv_mode    = False
        self._build_load_screen()

    def _clear(self):
        for w in self.winfo_children(): w.destroy()

    # ── Load screen ──────────────────────────────────────────────────────────
    def _build_load_screen(self):
        self._clear()
        frm = tk.Frame(self, bg=BG); frm.pack(expand=True)
        tk.Label(frm, text="BOX XML GENERATOR", font=("Consolas",20,"bold"),
                 bg=BG, fg=ACC1).pack(pady=(30,5))
        tk.Label(frm, text="Load a CSV with groups:  ID  |  Level/Rate  |  Parent Box Name",
                 bg=BG, fg=FG_DIM, font=("Consolas",10)).pack(pady=5)
        bf = tk.Frame(frm, bg=BG); bf.pack(pady=15)
        mk_btn(bf,"📂  Load CSV File",   self._load_csv_file).pack(side="left",padx=8)
        mk_btn(bf,"📋  Paste CSV Text",  self._paste_csv    ).pack(side="left",padx=8)
        mk_btn(bf,"✏️  No CSV — Manual Entry", self._start_no_csv,
               color=BG4).pack(side="left",padx=8)

    def _load_csv_file(self):
        p = filedialog.askopenfilename(filetypes=[("CSV","*.csv"),("All","*.*")])
        if p:
            with open(p, encoding="utf-8-sig") as f: self._process_csv(f.read())

    def _paste_csv(self):
        win = tk.Toplevel(self.root); win.title("Paste CSV")
        win.geometry("600x400"); win.configure(bg=BG)
        tk.Label(win, text="Paste CSV below:", bg=BG, fg=FG,
                 font=("Consolas",10)).pack(anchor="w",padx=10,pady=5)
        txt = scrolledtext.ScrolledText(win, font=("Consolas",9))
        txt.pack(fill="both",expand=True,padx=10,pady=5)
        def confirm():
            self._process_csv(txt.get("1.0","end")); win.destroy()
        mk_btn(win,"Confirm",confirm,color=GREEN,fg=BG2).pack(pady=8)

    def _process_csv(self, text):
        groups = parse_grouped_csv(text)
        if not groups:
            messagebox.showerror("Error","No valid box groups found in CSV."); return
        self.groups = groups; self.current_group_idx=0; self.box_configs=[]
        self.continue_mode=None; self.saved_settings=None; self.no_csv_mode=False
        self._build_config_screen()

    def _start_no_csv(self):
        self.no_csv_mode=True; self.current_group_idx=0; self.box_configs=[]
        self.continue_mode=None; self.saved_settings=None
        self.groups=[{"box_name":"Manual Entry","items":[
            {"id":"","name":"","extra":[],"rate":None,"item_cnt":1}]}]
        self._build_config_screen()

    # ── Config screen ─────────────────────────────────────────────────────────
    def _build_config_screen(self):
        self._clear()
        idx      = self.current_group_idx
        grp      = self.groups[idx]
        box_name = grp["box_name"]
        total    = len(self.groups)
        s        = self.saved_settings or {}
        prev_box_name = s.get("box_name","")

        # Wrapper uses grid so rows have explicit weights
        wrap = tk.Frame(self, bg=BG)
        wrap.pack(fill="both", expand=True)
        wrap.grid_rowconfigure(0, weight=0)
        wrap.grid_rowconfigure(1, weight=1)
        wrap.grid_rowconfigure(2, weight=0)
        wrap.grid_columnconfigure(0, weight=1)

        # Header — row 0, fixed height
        hdr = tk.Frame(wrap, bg=BG2)
        hdr.grid(row=0, column=0, sticky="ew")
        tk.Label(hdr, text=f"Box {idx+1} / {total}:  {box_name}",
                 font=("Consolas",14,"bold"), bg=BG2, fg=ACC1, pady=8
                 ).pack(side="left", padx=15)
        if self.continue_mode:
            tk.Label(hdr, text="🤖 AUTO" if self.continue_mode=="automate" else "👁 MONITOR",
                     font=("Consolas",10), bg=BG2, fg=ACC4).pack(side="right", padx=15)

        # Nav bar — row 2, always visible
        nav = tk.Frame(wrap, bg=BG2)
        nav.grid(row=2, column=0, sticky="ew")

        # Scrollable content — row 1, expands to fill
        scroll_host = tk.Frame(wrap, bg=BG)
        scroll_host.grid(row=1, column=0, sticky="nsew")
        canvas, container = mk_scroll_canvas(scroll_host)

        def section(title): return mk_section(container, title)
        def row_entry(p,lbl,var,w=38):
            r=tk.Frame(p,bg=BG); r.pack(fill="x",padx=8,pady=2)
            tk.Label(r,text=lbl,width=26,anchor="w",bg=BG,fg=FG,font=("Consolas",9)).pack(side="left")
            tk.Entry(r,textvariable=var,width=w,bg=BG3,fg=FG,insertbackground=FG,
                     font=("Consolas",9),relief="flat").pack(side="left",padx=4)
        def row_num(p,lbl,var,w=10):
            r=tk.Frame(p,bg=BG); r.pack(fill="x",padx=8,pady=2)
            tk.Label(r,text=lbl,width=26,anchor="w",bg=BG,fg=FG,font=("Consolas",9)).pack(side="left")
            tk.Entry(r,textvariable=var,width=w,bg=BG3,fg=FG,insertbackground=FG,
                     font=("Consolas",9),relief="flat").pack(side="left",padx=4)
        def note(p,txt):
            tk.Label(p,text=txt,bg=BG,fg=FG_GREY,font=("Consolas",8)).pack(anchor="w",padx=10,pady=(3,0))

        # cfg_override from CSV — must be resolved FIRST before any StringVar uses it
        cfg_ov = grp.get("cfg_override", {})

        try:    next_id = str(int(s.get("id",""))+1)
        except: next_id = s.get("id","")
        if cfg_ov.get("id"): next_id = cfg_ov["id"]
        saved_name_tmpl = s.get("name_template","")
        if saved_name_tmpl and prev_box_name:
            proposed_name = apply_name_template(saved_name_tmpl, prev_box_name, box_name)
        else:
            proposed_name = box_name
        used_names = [c["name"] for c in self.box_configs]
        proposed_name = deduplicate_name(proposed_name, used_names)
        if cfg_ov.get("name"): proposed_name = deduplicate_name(cfg_ov["name"], used_names)
        saved_cmt_tmpl = s.get("comment_template","")
        initial_comment = apply_name_template(saved_cmt_tmpl,prev_box_name,box_name) if saved_cmt_tmpl and prev_box_name else s.get("comment","Special dice that contains amazing items.")
        if cfg_ov.get("comment"): initial_comment = cfg_ov["comment"]
        saved_use_tmpl = s.get("use_template","")
        initial_use = apply_name_template(saved_use_tmpl,prev_box_name,box_name) if saved_use_tmpl and prev_box_name else s.get("use","Event Box.")
        if cfg_ov.get("use"): initial_use = cfg_ov["use"]

        v_id=tk.StringVar(value=next_id); v_name=tk.StringVar(value=proposed_name)
        v_comment=tk.StringVar(value=initial_comment); v_use=tk.StringVar(value=initial_use)
        v_file_name=tk.StringVar(value=cfg_ov.get("file_name") or s.get("file_name",r"data\item\itm_pre_107.nri"))
        v_bundle_num=tk.StringVar(value=cfg_ov.get("bundle_num") or s.get("bundle_num","0"))
        v_cmt_file_name=tk.StringVar(value=cfg_ov.get("cmt_file_name") or s.get("cmt_file_name",r"data\item\itm_pre_illu_107.nri"))
        v_cmt_bundle=tk.StringVar(value=cfg_ov.get("cmt_bundle_num") or s.get("cmt_bundle_num","0"))
        ov_checks = cfg_ov.get("opt_checks", s.get("opt_checks",[False]*8))
        ov_recycle = cfg_ov.get("opt_recycle", s.get("opt_recycle",0))
        ov_chr     = cfg_ov.get("chr_type_flags", s.get("chr_type_flags",[]))
        opt_check_vars=[tk.BooleanVar(value=ov_checks[i] if i<len(ov_checks) else False) for i in range(8)]
        opt_recycle_var=tk.IntVar(value=ov_recycle)
        chr_selected=list(ov_chr)
        present_type_var=tk.IntVar(value=s.get("present_type",0))
        drop_cnt_var=tk.StringVar(value=s.get("drop_cnt","1"))
        default_rate_var=tk.StringVar(value=s.get("default_rate","50"))
        remember_present=tk.BooleanVar(value=s.get("remember_present",False))
        v_weight=tk.StringVar(value=cfg_ov.get("weight") or s.get("weight","1"))
        v_value=tk.StringVar(value=cfg_ov.get("value") or s.get("value","0"))
        v_min_level=tk.StringVar(value=cfg_ov.get("min_level") or s.get("min_level","1"))
        v_money=tk.StringVar(value=cfg_ov.get("money") or s.get("money","0"))
        v_ticket=tk.StringVar(value=cfg_ov.get("ticket") or s.get("ticket","0"))
        dc_mode_var=tk.StringVar(value=s.get("dc_mode","flexible"))
        rate_mode_var=tk.StringVar(value=s.get("rate_mode","manual"))
        item_cnt_mode_var=tk.StringVar(value=s.get("item_cnt_mode","flexible"))
        item_cnt_univ_var=tk.StringVar(value=s.get("item_cnt_univ","1"))

        live_items=list(grp["items"])
        item_rate_vars=[]; item_cnt_vars=[]
        for it in live_items:
            if present_type_var.get()==2: initial_r="100"
            elif it.get("rate") is not None: initial_r=str(it["rate"])
            else:
                try: initial_r=str(int(s.get("default_rate","50")))
                except: initial_r="50"
            item_rate_vars.append(tk.StringVar(value=initial_r))
            item_cnt_vars.append(tk.StringVar(value=str(it.get("item_cnt",1))))

        # Basic Info
        sb = section("  ItemParam – Basic Info  ")
        row_entry(sb,"Box ID (itemparam):",v_id,20)
        note(sb,"  Name auto-filled from CSV header — editable.")
        row_entry(sb,"Name (CDATA):",v_name)
        row_entry(sb,"Comment (CDATA):",v_comment)
        row_entry(sb,"Use (CDATA):",v_use)

        # Filepaths
        sf = section("  Filepaths & Bundle Numbers  ")
        note(sf,"  FileName — also written to InvFileName.")
        row_entry(sf,"FileName (CDATA):",v_file_name)
        row_num(sf,"BundleNum:",v_bundle_num,8)
        row_entry(sf,"CmtFileName (CDATA):",v_cmt_file_name)
        row_num(sf,"CmtBundleNum:",v_cmt_bundle,8)

        # Options
        so = section("  Options  ")
        chk_frm=tk.Frame(so,bg=BG); chk_frm.pack(anchor="w",padx=8,pady=4)
        for i,(lbl,_) in enumerate(OPTIONS_CHECKS):
            tk.Checkbutton(chk_frm,text=lbl,variable=opt_check_vars[i],
                           bg=BG,fg=FG,selectcolor=BG3,activebackground=BG,
                           font=("Consolas",9)).grid(row=i//4,column=i%4,sticky="w",padx=6,pady=2)
        rec_frm=tk.Frame(so,bg=BG); rec_frm.pack(anchor="w",padx=8,pady=4)
        tk.Label(rec_frm,text="Recycle:",bg=BG,fg=FG,font=("Consolas",9)).pack(side="left",padx=(0,8))
        for lbl,val in [("None",0),("Recyclable",262144),("Non-Recyclable",8388608)]:
            tk.Radiobutton(rec_frm,text=lbl,variable=opt_recycle_var,value=val,
                           bg=BG,fg=FG,selectcolor=BG3,activebackground=BG,
                           font=("Consolas",9)).pack(side="left",padx=6)
        nc_container=tk.Frame(so,bg=BG); nc_container.pack(anchor="w",padx=8,pady=2)
        nc_inner=tk.Frame(nc_container,bg=BG)
        tk.Label(nc_inner,text="Ticket value (NCash = tickets × 133, rounded):",
                 bg=BG,fg=FG,font=("Consolas",9)).pack(side="left")
        tk.Entry(nc_inner,textvariable=v_ticket,width=10,bg=BG3,fg=FG,
                 insertbackground=FG,font=("Consolas",9),relief="flat").pack(side="left",padx=8)
        ncash_lbl=tk.Label(nc_inner,text="→ NCash: 0",bg=BG,fg=GREEN,font=("Consolas",9))
        ncash_lbl.pack(side="left")
        def update_ncash(*_):
            try: ncash_lbl.config(text=f"→ NCash: {round(float(v_ticket.get())*133)}")
            except: ncash_lbl.config(text="→ NCash: ?")
        v_ticket.trace_add("write",update_ncash); update_ncash()
        def toggle_ncash(*_):
            if opt_recycle_var.get()==262144: nc_inner.pack(anchor="w")
            else: nc_inner.pack_forget()
        opt_recycle_var.trace_add("write",toggle_ncash); toggle_ncash()

        # ChrTypeFlags
        sc_chr=section("  ChrTypeFlags  ")
        pf=tk.Frame(sc_chr,bg=BG); pf.pack(fill="x",padx=8,pady=4)
        tk.Label(pf,text="Character Type:",bg=BG,fg=FG,font=("Consolas",9)).pack(side="left")
        chr_name_combo=ttk.Combobox(pf,values=CHR_NAMES,state="readonly",width=14,font=("Consolas",9))
        chr_name_combo.pack(side="left",padx=(6,12))
        tk.Label(pf,text="Job:",bg=BG,fg=FG,font=("Consolas",9)).pack(side="left")
        chr_job_combo=ttk.Combobox(pf,values=CHR_JOBS,state="readonly",width=6,font=("Consolas",9))
        chr_job_combo.pack(side="left",padx=(6,12))
        lb_frm=tk.Frame(sc_chr,bg=BG); lb_frm.pack(fill="x",padx=8,pady=(0,6))
        tk.Label(lb_frm,text="Added:",bg=BG,fg=FG_GREY,font=("Consolas",8)).pack(anchor="w")
        chr_lb=tk.Listbox(lb_frm,height=4,width=36,bg=BG3,fg=FG,
                          font=("Consolas",9),selectbackground=BG4,activestyle="none")
        chr_lb.pack(anchor="w")
        def refresh_chr_lb():
            chr_lb.delete(0,"end")
            for val in chr_selected: chr_lb.insert("end",CHR_FLAG_REVERSE.get(val,str(val)))
        def add_chr_flag():
            name=chr_name_combo.get(); job=chr_job_combo.get()
            if not name or not job: return
            key=f"{name} {job}"; val=CHR_FLAG_MAP.get(key)
            if val and val not in chr_selected and len(chr_selected)<24:
                chr_selected.append(val); refresh_chr_lb()
        def rem_chr_flag():
            sel=chr_lb.curselection()
            if sel: chr_selected.pop(sel[0]); refresh_chr_lb()
        tk.Button(pf,text="+",command=add_chr_flag,bg=GREEN,fg=BG2,
                  font=("Consolas",11,"bold"),relief="flat",width=3).pack(side="left",padx=2)
        tk.Button(pf,text="−",command=rem_chr_flag,bg=ACC3,fg=BG2,
                  font=("Consolas",11,"bold"),relief="flat",width=3).pack(side="left",padx=2)
        refresh_chr_lb()

        # Numeric
        sn=section("  Numeric Fields  ")
        rf=tk.Frame(sn,bg=BG); rf.pack(fill="x",padx=8,pady=4)
        for ci,(lbl,var) in enumerate([("Weight:",v_weight),("Value:",v_value),
                                        ("MinLevel:",v_min_level),("Money:",v_money)]):
            tk.Label(rf,text=lbl,bg=BG,fg=FG,font=("Consolas",9),width=10,anchor="w").grid(row=0,column=ci*2,padx=4)
            tk.Entry(rf,textvariable=var,width=10,bg=BG3,fg=FG,insertbackground=FG,
                     font=("Consolas",9),relief="flat").grid(row=0,column=ci*2+1,padx=4)

        # PresentItemParam2 settings
        sp=section("  PresentItemParam2 Settings  ")
        id_disp=tk.Label(sp,text=f"Box ID: {next_id}  |  {box_name}",bg=BG,fg=FG_GREY,font=("Consolas",9))
        id_disp.pack(anchor="w",padx=8,pady=(4,0))
        v_id.trace_add("write",lambda *_: id_disp.config(text=f"Box ID: {v_id.get() or '—'}  |  {box_name}"))
        type_frm=tk.Frame(sp,bg=BG); type_frm.pack(anchor="w",padx=8,pady=4)
        tk.Label(type_frm,text="Drop Type:",bg=BG,fg=FG,font=("Consolas",9)).pack(side="left")
        for lbl,val in [("Random",0),("Distributive",2)]:
            tk.Radiobutton(type_frm,text=lbl,variable=present_type_var,value=val,
                           bg=BG,fg=FG,selectcolor=BG3,activebackground=BG,
                           font=("Consolas",9)).pack(side="left",padx=8)
        tk.Checkbutton(type_frm,text="Remember this setting",variable=remember_present,
                       bg=BG,fg=ACC4,selectcolor=BG3,activebackground=BG,
                       font=("Consolas",9)).pack(side="left",padx=14)

        dc_mode_frm=tk.Frame(sp,bg=BG); dc_mode_frm.pack(anchor="w",padx=8,pady=(2,0))
        tk.Label(dc_mode_frm,text="DropCnt:",bg=BG,fg=FG,font=("Consolas",9)).pack(side="left")
        tk.Radiobutton(dc_mode_frm,text="Flexible  (= item count)",variable=dc_mode_var,value="flexible",
                       bg=BG,fg=FG,selectcolor=BG3,activebackground=BG,font=("Consolas",9),
                       command=lambda:_toggle_dc_mode()).pack(side="left",padx=6)
        tk.Radiobutton(dc_mode_frm,text="Manual",variable=dc_mode_var,value="manual",
                       bg=BG,fg=FG,selectcolor=BG3,activebackground=BG,font=("Consolas",9),
                       command=lambda:_toggle_dc_mode()).pack(side="left",padx=4)
        dc_frm=tk.Frame(sp,bg=BG); dc_frm.pack(anchor="w",padx=8,pady=2)
        dc_lbl=tk.Label(dc_frm,text="DropCnt value:",bg=BG,fg=FG,font=("Consolas",9))
        dc_ent=tk.Entry(dc_frm,textvariable=drop_cnt_var,width=6,bg=BG3,fg=FG,
                        insertbackground=FG,font=("Consolas",9),relief="flat")
        rate_note=tk.Label(sp,text="",bg=BG,fg=FG_GREY,font=("Consolas",8))
        rate_note.pack(anchor="w",padx=8)
        def _toggle_dc_mode():
            if dc_mode_var.get()=="manual": dc_lbl.pack(side="left"); dc_ent.pack(side="left",padx=6)
            else: dc_lbl.pack_forget(); dc_ent.pack_forget()

        dr_mode_frm=tk.Frame(sp,bg=BG); dr_mode_frm.pack(anchor="w",padx=8,pady=(4,0))
        tk.Label(dr_mode_frm,text="DropRate:",bg=BG,fg=FG,font=("Consolas",9)).pack(side="left")
        tk.Radiobutton(dr_mode_frm,text="Manual",variable=rate_mode_var,value="manual",
                       bg=BG,fg=FG,selectcolor=BG3,activebackground=BG,font=("Consolas",9),
                       command=lambda:_toggle_rate_mode()).pack(side="left",padx=6)
        tk.Radiobutton(dr_mode_frm,text="Distributive  (set all entries equally)",
                       variable=rate_mode_var,value="distributive",
                       bg=BG,fg=FG,selectcolor=BG3,activebackground=BG,font=("Consolas",9),
                       command=lambda:_toggle_rate_mode()).pack(side="left",padx=4)
        dr_dist_frm=tk.Frame(sp,bg=BG)
        tk.Label(dr_dist_frm,text="Default DropRate:",bg=BG,fg=FG,font=("Consolas",9)).pack(side="left")
        dr_dist_ent=tk.Entry(dr_dist_frm,textvariable=default_rate_var,width=6,bg=BG3,fg=FG,
                             insertbackground=FG,font=("Consolas",9),relief="flat")
        dr_dist_ent.pack(side="left",padx=6)
        dr_dist_preview=tk.Label(dr_dist_frm,text="",bg=BG,fg=FG_GREY,font=("Consolas",8))
        dr_dist_preview.pack(side="left")
        def _apply_dist_rate(*_):
            if rate_mode_var.get()=="distributive":
                try:   val=str(int(default_rate_var.get()))
                except: val="50"
                for var in item_rate_vars: var.set(val)
                dr_dist_preview.config(text=f"→ all items set to {val}")
        default_rate_var.trace_add("write",_apply_dist_rate)

        # Box Contents table — defined BEFORE the ItemCnt radios so the command lambda works
        sec_items=tk.LabelFrame(container,text="  Box Contents  ",
                                bg=BG,fg=BLUE,font=("Consolas",10,"bold"),bd=1,relief="groove")
        sec_items.pack(fill="x",padx=12,pady=5)
        items_outer=tk.Frame(sec_items,bg=BG); items_outer.pack(fill="x",padx=8,pady=4)
        rate_entry_widgets=[]; icnt_entry_widgets=[]
        icnt_univ_frm=tk.Frame(sp,bg=BG)
        tk.Label(icnt_univ_frm,text="ItemCnt (all):",bg=BG,fg=FG,font=("Consolas",9)).pack(side="left")
        tk.Entry(icnt_univ_frm,textvariable=item_cnt_univ_var,width=6,bg=BG3,fg=FG,
                 insertbackground=FG,font=("Consolas",9),relief="flat").pack(side="left",padx=6)

        def _rebuild_items_table():
            nonlocal rate_entry_widgets, icnt_entry_widgets
            while len(item_rate_vars)<len(live_items): item_rate_vars.append(tk.StringVar(value=default_rate_var.get() or "50"))
            while len(item_rate_vars)>len(live_items): item_rate_vars.pop()
            while len(item_cnt_vars)<len(live_items): item_cnt_vars.append(tk.StringVar(value="1"))
            while len(item_cnt_vars)>len(live_items): item_cnt_vars.pop()
            for w in items_outer.winfo_children(): w.destroy()
            rate_entry_widgets=[]; icnt_entry_widgets=[]
            sec_items.config(text=f"  Box Contents ({len(live_items)} items)  ")
            icnt_mode=item_cnt_mode_var.get(); is_no_csv=self.no_csv_mode
            if icnt_mode=="universal": icnt_univ_frm.pack(anchor="w",padx=8,pady=(0,4))
            else: icnt_univ_frm.pack_forget()
            cols=[("#",3),("ID",9 if is_no_csv else 10),("Name",40 if is_no_csv else 44),("DropRate",8)]
            if icnt_mode=="manual": cols.append(("ItemCnt",7))
            if is_no_csv: cols.append(("",5))
            for ci,(txt,w) in enumerate(cols):
                tk.Label(items_outer,text=txt,width=w,bg=BG2,fg=BLUE,
                         font=("Consolas",9,"bold"),anchor="w").grid(row=0,column=ci,padx=2,pady=2,sticky="w")
            for i,item in enumerate(live_items):
                bg = BG if i%2==0 else BG2; col=0
                tk.Label(items_outer,text=str(i),width=3,bg=bg,fg=FG_GREY,
                         font=("Consolas",9)).grid(row=i+1,column=col,padx=2,pady=1); col+=1
                if is_no_csv:
                    id_var=tk.StringVar(value=item.get("id",""))
                    tk.Entry(items_outer,textvariable=id_var,width=9,bg=BG3,fg=FG,
                             insertbackground=FG,font=("Consolas",9),relief="flat").grid(row=i+1,column=col,padx=2)
                    id_var.trace_add("write",lambda *_,v=id_var,it=item: it.update({"id":v.get()}))
                    name_var=tk.StringVar(value=item.get("name",""))
                    tk.Entry(items_outer,textvariable=name_var,width=40,bg=BG3,fg=FG,
                             insertbackground=FG,font=("Consolas",9),relief="flat").grid(row=i+1,column=col+1,padx=2,sticky="w")
                    name_var.trace_add("write",lambda *_,v=name_var,it=item: it.update({"name":v.get()}))
                else:
                    tk.Label(items_outer,text=item.get("id",""),width=9,bg=bg,fg=FG,
                             font=("Consolas",9)).grid(row=i+1,column=col,padx=2)
                    tk.Label(items_outer,text=item.get("name","")[:52],width=44,bg=bg,fg=FG_DIM,
                             font=("Consolas",9),anchor="w").grid(row=i+1,column=col+1,padx=2,sticky="w")
                col+=2
                rate_ent=tk.Entry(items_outer,textvariable=item_rate_vars[i],width=8,bg=BG3,fg=FG,
                                  insertbackground=FG,font=("Consolas",9),relief="flat")
                rate_ent.grid(row=i+1,column=col,padx=2); rate_entry_widgets.append(rate_ent); col+=1
                if icnt_mode=="manual":
                    icnt_ent=tk.Entry(items_outer,textvariable=item_cnt_vars[i],width=7,bg=BG3,fg=FG,
                                      insertbackground=FG,font=("Consolas",9),relief="flat")
                    icnt_ent.grid(row=i+1,column=col,padx=2); icnt_entry_widgets.append(icnt_ent); col+=1
                if is_no_csv:
                    def make_remove(idx):
                        def _rem():
                            if len(live_items)<=1: return
                            live_items.pop(idx)
                            if idx<len(item_rate_vars): item_rate_vars.pop(idx)
                            if idx<len(item_cnt_vars): item_cnt_vars.pop(idx)
                            _rebuild_items_table()
                        return _rem
                    tk.Button(items_outer,text="−",width=2,command=make_remove(i),
                              bg=ACC3,fg=BG2,font=("Consolas",9,"bold"),relief="flat"
                              ).grid(row=i+1,column=col,padx=2)
            if is_no_csv and len(live_items)<20:
                def _add_item():
                    live_items.append({"id":"","name":"","extra":[],"rate":None,"item_cnt":1})
                    _rebuild_items_table()
                tk.Button(items_outer,text="＋ Add row",command=_add_item,
                          bg=GREEN,fg=BG2,font=("Consolas",9),relief="flat",padx=6,pady=2
                          ).grid(row=len(live_items)+1,column=0,columnspan=6,sticky="w",padx=4,pady=4)
            _toggle_rate_mode_inner()

        # ItemCnt mode radios — placed here so _rebuild_items_table is already defined
        icnt_mode_frm=tk.Frame(sp,bg=BG); icnt_mode_frm.pack(anchor="w",padx=8,pady=(4,2))
        tk.Label(icnt_mode_frm,text="ItemCnt:",bg=BG,fg=FG,font=("Consolas",9)).pack(side="left")
        for _lbl_ic,_val_ic in [("Flexible (=1)","flexible"),("Universal","universal"),("Manual (per row)","manual")]:
            tk.Radiobutton(icnt_mode_frm,text=_lbl_ic,variable=item_cnt_mode_var,value=_val_ic,
                           bg=BG,fg=FG,selectcolor=BG3,activebackground=BG,font=("Consolas",9),
                           command=_rebuild_items_table).pack(side="left",padx=4)

        def _toggle_rate_mode_inner():
            is_dist=rate_mode_var.get()=="distributive"
            for ent in rate_entry_widgets:
                ent.config(state="disabled" if is_dist else "normal",disabledforeground=FG_GREY)
            if is_dist: _apply_dist_rate()

        def _toggle_rate_mode():
            if rate_mode_var.get()=="distributive":
                dr_dist_frm.pack(anchor="w",padx=8,pady=2); _toggle_rate_mode_inner()
            else:
                dr_dist_frm.pack_forget()
                for ent in rate_entry_widgets: ent.config(state="normal")
                dr_dist_preview.config(text="")

        def toggle_drop_fields(*_):
            _toggle_dc_mode(); rate_note.config(text="")

        _rebuild_items_table()
        present_type_var.trace_add("write",toggle_drop_fields)
        _toggle_dc_mode(); toggle_drop_fields(); _toggle_rate_mode()

        # Gather / nav
        def gather_config():
            try: ncash=round(float(v_ticket.get() or "0")*133)
            except: ncash=0
            item_rates=[]
            for iv in item_rate_vars:
                try: item_rates.append(int(iv.get()))
                except: item_rates.append(100 if present_type_var.get()==2 else 50)
            icnt_mode=item_cnt_mode_var.get()
            if icnt_mode=="flexible": item_cnts=[1]*len(live_items)
            elif icnt_mode=="universal":
                try: uval=int(item_cnt_univ_var.get()) or 1
                except: uval=1
                item_cnts=[uval]*len(live_items)
            else:
                item_cnts=[]
                for iv in item_cnt_vars:
                    try: item_cnts.append(int(iv.get()) or 1)
                    except: item_cnts.append(1)
            return {
                "id":v_id.get().strip(),"name":v_name.get(),
                "name_template":v_name.get(),"comment":v_comment.get(),
                "comment_template":v_comment.get(),"use":v_use.get(),
                "use_template":v_use.get(),"box_name":box_name,
                "items":list(live_items),"item_rates":item_rates,"item_cnts":item_cnts,
                "item_cnt_mode":icnt_mode,"item_cnt_univ":item_cnt_univ_var.get() or "1",
                "file_name":v_file_name.get(),"bundle_num":v_bundle_num.get() or "0",
                "cmt_file_name":v_cmt_file_name.get(),"cmt_bundle_num":v_cmt_bundle.get() or "0",
                "weight":v_weight.get() or "1","value":v_value.get() or "0",
                "min_level":v_min_level.get() or "1","money":v_money.get() or "0",
                "ncash":ncash,"ticket":v_ticket.get() or "0",
                "opt_checks":[bv.get() for bv in opt_check_vars],
                "opt_recycle":opt_recycle_var.get(),"chr_type_flags":list(chr_selected),
                "present_type":present_type_var.get(),
                "drop_cnt":drop_cnt_var.get() or "1" if dc_mode_var.get()=="manual" else str(len(live_items)),
                "default_rate":default_rate_var.get() or "50",
                "remember_present":remember_present.get(),
                "dc_mode":dc_mode_var.get(),"rate_mode":rate_mode_var.get(),
            }

        def save_settings(cfg):
            self.saved_settings={k:v for k,v in cfg.items() if k!="items"}

        # nav is already created above, pinned to bottom of outer

        def go_next():
            cfg=gather_config()
            if not cfg["id"]: messagebox.showwarning("Missing ID","Please enter a Box ID."); return
            blanks=[]
            if not cfg["name"].strip(): blanks.append("Name")
            if not cfg["file_name"].strip(): blanks.append("FileName")
            if not cfg["cmt_file_name"].strip(): blanks.append("CmtFileName")
            if blanks:
                if not messagebox.askyesno("Missed a spot","These fields are empty:\n\n  "+"\n  ".join(blanks)+"\n\nContinue anyway?"): return
            self.box_configs.append(cfg); save_settings(cfg); self.current_group_idx+=1
            if self.no_csv_mode or self.current_group_idx>=len(self.groups): self._build_output_screen()
            elif self.continue_mode=="automate": self._automate_remaining(cfg)
            elif self.continue_mode=="monitor": self._build_config_screen()
            else: self._ask_automate_or_monitor(cfg)

        mk_btn(nav,"◀  Back to CSV",self._build_load_screen).pack(side="left",padx=10,pady=8)
        if idx>0:
            def go_prev():
                self.current_group_idx-=1
                if self.box_configs: self.box_configs.pop()
                self._build_config_screen()
            mk_btn(nav,"◀  Previous Box",go_prev).pack(side="left",padx=4,pady=8)
        if self.continue_mode:
            mk_btn(nav,"⚙ Change Mode",lambda: (setattr(self,"continue_mode",None),self._build_config_screen()),
                   color=BG4).pack(side="left",padx=4,pady=8)
        next_lbl="Generate XML ✓" if idx==total-1 else "Next Box ▶"
        mk_btn(nav,next_lbl,go_next,color=GREEN,fg=BG2,font=("Consolas",10,"bold")).pack(side="right",padx=10,pady=8)

    def _ask_automate_or_monitor(self, last_cfg):
        remaining=len(self.groups)-self.current_group_idx
        win=tk.Toplevel(self.root); win.title("Continue?"); win.geometry("530x260")
        win.configure(bg=BG); win.grab_set()
        tk.Label(win,text=f"{remaining} box(es) remaining.",bg=BG,fg=FG,font=("Consolas",13,"bold")).pack(pady=14)
        tk.Label(win,text="How would you like to continue?",bg=BG,fg=FG_DIM,font=("Consolas",10)).pack()
        remember_var=tk.BooleanVar(value=False)
        bf=tk.Frame(win,bg=BG); bf.pack(pady=10)
        def choose(mode):
            if remember_var.get(): self.continue_mode=mode
            win.destroy()
            if mode=="automate": self._automate_remaining(last_cfg)
            else: self._build_config_screen()
        mk_btn(bf,"🤖  Automate  —  use saved settings for all remaining boxes",
               lambda:choose("automate"),color=ACC1,fg=BG2).pack(pady=5)
        mk_btn(bf,"👁  Monitor  —  review each box individually",
               lambda:choose("monitor"),color=BLUE,fg=BG2).pack(pady=5)
        tk.Checkbutton(win,text="Remember my choice for the rest of this session",
                       variable=remember_var,bg=BG,fg=ACC4,selectcolor=BG3,
                       activebackground=BG,font=("Consolas",9)).pack(pady=4)

    def _automate_remaining(self, last_cfg):
        try: base_id=int(last_cfg["id"])
        except:
            messagebox.showerror("Error","Cannot automate: last Box ID was not a plain integer.")
            self.continue_mode="monitor"; self._build_config_screen(); return
        try: default_rate=int(last_cfg.get("default_rate",50))
        except: default_rate=50
        id_counter=base_id+1; prev_name=last_cfg["box_name"]
        used_names=[c["name"] for c in self.box_configs]
        for i in range(self.current_group_idx, len(self.groups)):
            grp=self.groups[i]; cfg=copy.deepcopy(last_cfg)
            cfg["id"]=str(id_counter)
            proposed=apply_name_template(cfg.get("name_template",""),prev_name,grp["box_name"])
            proposed=deduplicate_name(proposed,used_names)
            cfg["name"]=proposed; cfg["name_template"]=proposed
            cfg["comment"]=apply_name_template(cfg.get("comment_template",""),prev_name,grp["box_name"])
            cfg["use"]=apply_name_template(cfg.get("use_template",""),prev_name,grp["box_name"])
            cfg["comment_template"]=cfg["comment"]; cfg["use_template"]=cfg["use"]
            cfg["box_name"]=grp["box_name"]; cfg["items"]=grp["items"]
            is_distrib=cfg["present_type"]==2
            cfg["item_rates"]=[100 if is_distrib else (it.get("rate") if it.get("rate") is not None else default_rate) for it in grp["items"]]
            icnt_mode=cfg.get("item_cnt_mode","flexible")
            if icnt_mode=="flexible": cfg["item_cnts"]=[1]*len(grp["items"])
            elif icnt_mode=="universal":
                try: uval=int(cfg.get("item_cnt_univ","1")) or 1
                except: uval=1
                cfg["item_cnts"]=[uval]*len(grp["items"])
            else: cfg["item_cnts"]=[it.get("item_cnt",1) for it in grp["items"]]
            used_names.append(proposed); prev_name=grp["box_name"]
            self.box_configs.append(cfg); id_counter+=1
        self.current_group_idx=len(self.groups); self._build_output_screen()

    def _build_output_screen(self):
        self._clear()
        tk.Label(self,text="Generated XML Output",font=("Consolas",14,"bold"),
                 bg=BG,fg=ACC1).pack(pady=10)
        nb=ttk.Notebook(self); nb.pack(fill="both",expand=True,padx=12,pady=4)
        itemparam_rows=[]; presentparam_rows=[]; recycle_except_rows=[]
        for cfg in self.box_configs:
            try: default_rate=int(cfg.get("default_rate",50))
            except: default_rate=50
            try: drop_cnt=int(cfg.get("drop_cnt",1))
            except: drop_cnt=1
            is_distrib=(cfg["present_type"]==2)
            items_with_rates=[]
            for j,it in enumerate(cfg["items"]):
                rate=100 if is_distrib else (cfg["item_rates"][j] if j<len(cfg["item_rates"]) else default_rate)
                items_with_rates.append({**it,"rate":rate})
            item_cnts=cfg.get("item_cnts") or [1]*len(cfg["items"])
            itemparam_rows.append(build_itemparam_row(cfg))
            presentparam_rows.append(build_presentparam_row(cfg["id"],items_with_rates,cfg["present_type"],drop_cnt,default_rate,item_cnts=item_cnts,box_name=cfg.get("name","")))
            if cfg.get("opt_recycle",0) in (0,8388608):
                recycle_except_rows.append(build_recycle_except_row(cfg["id"],cfg["name"]))
        csv_lines=["ID,BoxName"]+[f"{c['id']},{c['box_name']}" for c in self.box_configs]
        csv_text="\n".join(csv_lines)
        # Write to shared session so other tools can import
        self.session.box_id_list_csv = csv_text
        self.session.box_id_map = {c["id"]: c["box_name"] for c in self.box_configs}
        _exports=[("itemparam_rows.xml","\n".join(itemparam_rows)),
                  ("presentparam_rows.xml","\n".join(presentparam_rows)),
                  ("box_id_list.csv",csv_text)]
        if recycle_except_rows: _exports.append(("RecycleExceptItem_rows.xml","\n".join(recycle_except_rows)))
        make_output_tab(nb,"itemparam.xml rows","\n".join(itemparam_rows),"itemparam_rows.xml",self.root)
        make_output_tab(nb,"PresentItemParam2.xml rows","\n".join(presentparam_rows),"presentparam_rows.xml",self.root)
        make_output_tab(nb,"Box ID List (→ Tool 2)","\n".join(csv_lines),"box_id_list.csv",self.root)
        if recycle_except_rows: make_output_tab(nb,"RecycleExceptItem.xml rows","\n".join(recycle_except_rows),"RecycleExceptItem_rows.xml",self.root)
        nb.select(1)
        bot=tk.Frame(self,bg=BG); bot.pack(fill="x",pady=6)
        def export_all():
            folder=filedialog.askdirectory(title="Choose export folder")
            if not folder: return
            saved=[]
            for fname,content in _exports:
                with open(os.path.join(folder,fname),"w",encoding="utf-8") as f: f.write(content)
                saved.append(fname)
            messagebox.showinfo("Export Complete",f"Saved to:\n{folder}\n\n"+"\n".join(saved))
        mk_btn(bot,"💾  Export All Files",export_all,color=ACC1,fg=BG2,font=("Consolas",11,"bold")).pack(side="left",padx=14)
        mk_btn(bot,"◀  Start Over",self._build_load_screen).pack(side="left",padx=4)

# ══════════════════════════════════════════════════════════════════════════════
# TOOL 2 — PresentItemParam2 Rate Adjuster
# ══════════════════════════════════════════════════════════════════════════════
class Tool2(tk.Frame):
    def __init__(self, parent, root, session):
        super().__init__(parent, bg=BG)
        self.root=root; self.session=session; self.csv_text=""; self.xml_text=""
        self.item_lib={}; self.mode_var=tk.StringVar(value="automatic")
        self._rate_var=tk.StringVar(value="100"); self._count_var=tk.StringVar(value="1")
        self._lib_status=tk.StringVar(value="No library loaded  (item names won't appear)")
        self._mode_panel_frame=None; self._build_load_screen()

    def _clear(self):
        for w in self.winfo_children(): w.destroy()

    def _build_load_screen(self):
        self._clear()
        tk.Label(self,text="PRESENTITEMPARAM2 RATE ADJUSTER",font=("Consolas",16,"bold"),
                 bg=BG,fg=ACC2).pack(pady=(18,2))
        tk.Label(self,text="CSV must contain the BOX IDs  (<Id> in PresentItemParam2).\nUse the 'Box ID List' exported by Tool 1, or any 2-col ID / Name CSV.",
                 bg=BG,fg=FG_GREY,font=("Consolas",8),justify="center").pack(pady=(0,6))

        csv_frm=mk_section(self,"  Step 1 — Box ID CSV  ")
        csv_status=tk.StringVar(value="No file loaded")
        tk.Label(csv_frm,textvariable=csv_status,bg=BG,fg=FG_GREY,font=("Consolas",9)).pack(side="left",padx=10)
        def load_csv():
            p=filedialog.askopenfilename(filetypes=[("CSV","*.csv"),("All","*.*")])
            if not p: return
            with open(p,encoding="utf-8-sig") as f: self.csv_text=f.read()
            bm=parse_box_id_csv(self.csv_text)
            csv_status.set(f"✓  {os.path.basename(p)}  ({len(bm)} box IDs found)")
        mk_btn(csv_frm,"📂 Load CSV",load_csv,padx=10,pady=4).pack(side="right",padx=8,pady=6)
        def import_session_t2():
            if not self.session.box_id_list_csv:
                messagebox.showinfo("No Session Data","Run Tool 1 first to generate box IDs.")
                return
            self.csv_text = self.session.box_id_list_csv
            bm = parse_box_id_csv(self.csv_text)
            csv_status.set(f"✓  Imported from Tool 1  ({len(bm)} box IDs)")
        mk_btn(csv_frm,"⬇  Import from Tool 1",import_session_t2,color=ACC2,fg=BG2,padx=8,pady=4).pack(side="right",padx=4,pady=6)

        xml_frm=mk_section(self,"  Step 2 — PresentItemParam2.xml  ")
        xml_status=tk.StringVar(value="No file loaded")
        tk.Label(xml_frm,textvariable=xml_status,bg=BG,fg=FG_GREY,font=("Consolas",9)).pack(side="left",padx=10)
        def load_xml():
            p=filedialog.askopenfilename(filetypes=[("XML","*.xml"),("All","*.*")])
            if not p: return
            with open(p,encoding="utf-8-sig") as f: self.xml_text=f.read()
            xml_status.set(f"✓  {os.path.basename(p)}")
        mk_btn(xml_frm,"📂 Load XML",load_xml,padx=10,pady=4).pack(side="right",padx=8,pady=6)

        mode_frm=mk_section(self,"  Step 3 — Mode  ")
        mf=tk.Frame(mode_frm,bg=BG); mf.pack(anchor="w",padx=10,pady=6)
        for lbl,val in [("Manual     — review and configure each box individually","manual"),
                        ("Automatic  — apply the same values to every matched box","automatic")]:
            tk.Radiobutton(mf,text=lbl,variable=self.mode_var,value=val,
                           bg=BG,fg=FG,selectcolor=BG3,activebackground=BG,font=("Consolas",10),
                           command=self._refresh_mode_panel).pack(anchor="w",pady=2)

        self._mode_panel_frame=tk.Frame(self,bg=BG)
        self._mode_panel_frame.pack(fill="x",padx=30,pady=2)
        self._refresh_mode_panel()
        mk_btn(self,"▶  Continue →",self._on_continue,color=GREEN,fg=BG2,
               font=("Consolas",12,"bold")).pack(pady=14)

    def _refresh_mode_panel(self):
        if not self._mode_panel_frame: return
        for w in self._mode_panel_frame.winfo_children(): w.destroy()
        if self.mode_var.get()=="automatic":
            frm=tk.LabelFrame(self._mode_panel_frame,text="  Adjustment Values  ",
                              bg=BG,fg=BLUE,font=("Consolas",10,"bold"),bd=1,relief="groove")
            frm.pack(fill="x")
            def num_row(lbl,var,note=""):
                r=tk.Frame(frm,bg=BG); r.pack(fill="x",padx=10,pady=4)
                tk.Label(r,text=lbl,width=18,anchor="w",bg=BG,fg=FG,font=("Consolas",9)).pack(side="left")
                tk.Entry(r,textvariable=var,width=8,bg=BG3,fg=FG,insertbackground=FG,
                         font=("Consolas",9),relief="flat").pack(side="left",padx=6)
                tk.Label(r,text=note,bg=BG,fg=FG_GREY,font=("Consolas",8)).pack(side="left")
            num_row("Adjust Rate:",self._rate_var,"(1–32766)  applied to every used DropRate_# slot")
            num_row("Adjust Count:",self._count_var,"(1–32766)  applied to every used ItemCnt_# slot")
            tk.Label(frm,text="  Type will be set to 2.  DropCnt will be set to the number of real items.",
                     bg=BG,fg=FG_GREY,font=("Consolas",8)).pack(anchor="w",padx=10,pady=(0,6))
        else:
            frm=tk.LabelFrame(self._mode_panel_frame,text="  ItemParam Library (optional)  ",
                              bg=BG,fg=BLUE,font=("Consolas",10,"bold"),bd=1,relief="groove")
            frm.pack(fill="x")
            tk.Label(frm,textvariable=self._lib_status,bg=BG,fg=FG_GREY,font=("Consolas",9)).pack(side="left",padx=10,pady=6)
            def load_lib():
                folder=filedialog.askdirectory(title="Select folder containing ItemParam XML files")
                if not folder: return
                self.item_lib=load_itemparam_folder(folder)
                self._lib_status.set(f"✓  {len(self.item_lib)} items from {os.path.basename(folder)}")
            mk_btn(frm,"📂 Load Folder",load_lib,padx=10,pady=4).pack(side="right",padx=8,pady=6)

    def _on_continue(self):
        if not self.csv_text: messagebox.showwarning("Missing","Please load a CSV first."); return
        if not self.xml_text: messagebox.showwarning("Missing","Please load PresentItemParam2.xml first."); return
        box_map=parse_box_id_csv(self.csv_text)
        if not box_map: messagebox.showerror("Error","No box IDs found in CSV."); return
        matched=[]; seen=set()
        for row in ROW_RE.findall(self.xml_text):
            rid=_get_tag(row,"Id")
            if rid in box_map and rid not in seen:
                matched.append((rid,box_map[rid],row)); seen.add(rid)
        if not matched:
            messagebox.showwarning("No Matches","None of the CSV box IDs matched any <Id> in the XML.\n\nMake sure you're using the Box ID CSV from Tool 1."); return
        if self.mode_var.get()=="automatic":
            try:
                rate=int(self._rate_var.get()); count=int(self._count_var.get())
                if not (1<=rate<=32766) or not (1<=count<=32766): raise ValueError
            except: messagebox.showerror("Invalid","Rate and Count must be integers 1–32766."); return
            self._run_automatic(matched,rate,count)
        else: self._run_manual(matched)

    def _run_automatic(self, matched, rate, count):
        matched_ids={rid:row for rid,_,row in matched}; csv_rows=[]
        def replace_row(m):
            row=m.group(0); rid=_get_tag(row,"Id")
            if rid not in matched_ids: return row
            slots=real_drop_slots(row)
            cfg={"type":2,"drop_cnt":len(slots),"slots":[{"rate":rate,"count":count} for _ in slots]}
            new_row=apply_cfg_to_row(row,cfg)
            drop_ids=[v for _,v in real_drop_slots(new_row)]
            name=next((n for r,n,_ in matched if r==rid),"")
            csv_rows.append([rid,name,*drop_ids]); return new_row
        full_out=ROW_RE.sub(replace_row,self.xml_text)
        self._build_output_screen(full_out,csv_rows,len(matched))

    def _run_manual(self, matched):
        self.manual_matched=matched; self.manual_idx=0
        self.manual_configs={}; self.manual_saved=None; self.manual_continue_mode=None
        self._build_manual_screen()

    def _build_manual_screen(self):
        self._clear()
        idx=self.manual_idx; total=len(self.manual_matched)
        rid,csv_name,row_block=self.manual_matched[idx]
        slots=real_drop_slots(row_block)
        s=self.manual_saved or {}
        last_type=s.get("type",2); last_dc=s.get("drop_cnt",len(slots)); last_slots=s.get("slots",[])

        wrap=tk.Frame(self,bg=BG); wrap.pack(fill="both",expand=True)
        wrap.grid_rowconfigure(0,weight=0); wrap.grid_rowconfigure(1,weight=1); wrap.grid_rowconfigure(2,weight=0)
        wrap.grid_columnconfigure(0,weight=1)
        hdr=tk.Frame(wrap,bg=BG2); hdr.grid(row=0,column=0,sticky="ew")
        hdr_txt=f"  Box {idx+1} / {total}   ID: {rid}"
        if csv_name: hdr_txt+=f"   —   {csv_name}"
        tk.Label(hdr,text=hdr_txt,font=("Consolas",12,"bold"),bg=BG2,fg=ACC2,pady=8).pack(side="left",padx=10)
        if self.manual_continue_mode:
            tk.Label(hdr,text="🤖 AUTO" if self.manual_continue_mode=="automate" else "👁 MONITOR",
                     font=("Consolas",10),bg=BG2,fg=ACC4).pack(side="right",padx=15)
        nav2=tk.Frame(wrap,bg=BG2); nav2.grid(row=2,column=0,sticky="ew")
        scroll_host2=tk.Frame(wrap,bg=BG); scroll_host2.grid(row=1,column=0,sticky="nsew")
        canvas,cont=mk_scroll_canvas(scroll_host2)

        sec_type=mk_section(cont,"  Drop Type  ")
        type_var=tk.IntVar(value=last_type)
        tf=tk.Frame(sec_type,bg=BG); tf.pack(anchor="w",padx=8,pady=6)
        for lbl,val in [("Random (Type=0)",0),("Egalitarian (Type=2)",2)]:
            tk.Radiobutton(tf,text=lbl,variable=type_var,value=val,
                           bg=BG,fg=FG,selectcolor=BG3,activebackground=BG,font=("Consolas",10)).pack(side="left",padx=8)
        dc_frame=tk.Frame(sec_type,bg=BG); dc_frame.pack(anchor="w",padx=8,pady=(0,6))
        dc_var=tk.StringVar(value=str(last_dc))
        dc_lbl=tk.Label(dc_frame,text="Types of Items (DropCnt):",bg=BG,fg=FG,font=("Consolas",9))
        dc_ent=tk.Entry(dc_frame,textvariable=dc_var,width=6,bg=BG3,fg=FG,insertbackground=FG,font=("Consolas",9),relief="flat")
        dc_auto=tk.Label(dc_frame,text=f"DropCnt auto-set to {len(slots)}  (all items in this box)",bg=BG,fg=FG_GREY,font=("Consolas",9))
        def toggle_dc(*_):
            if type_var.get()==0: dc_auto.pack_forget(); dc_lbl.pack(side="left"); dc_ent.pack(side="left",padx=6)
            else: dc_lbl.pack_forget(); dc_ent.pack_forget(); dc_auto.pack(anchor="w")
        type_var.trace_add("write",toggle_dc); toggle_dc()

        sec_slots=mk_section(cont,f"  Drop Slots  ({len(slots)} items)  ")
        hrow=tk.Frame(sec_slots,bg=BG2); hrow.pack(fill="x",padx=8,pady=2)
        for txt,w in [("Slot #",6),("Drop ID",12),("Item Name",34),("Rate %",9),("Item Count",10)]:
            tk.Label(hrow,text=txt,width=w,bg=BG2,fg=BLUE,font=("Consolas",9,"bold"),anchor="w").pack(side="left",padx=3)
        slot_rate_vars=[]; slot_count_vars=[]
        for pos,(sidx,drop_id) in enumerate(slots):
            bg=BG if pos%2==0 else BG2
            srow=tk.Frame(sec_slots,bg=bg); srow.pack(fill="x",padx=8,pady=1)
            prev_r=last_slots[pos]["rate"] if pos<len(last_slots) else 100
            prev_c=last_slots[pos]["count"] if pos<len(last_slots) else 1
            tk.Label(srow,text=str(sidx),width=6,bg=bg,fg=FG_GREY,font=("Consolas",9)).pack(side="left",padx=3)
            tk.Label(srow,text=drop_id,width=12,bg=bg,fg=BG4,font=("Consolas",9)).pack(side="left",padx=3)
            name=self.item_lib.get(drop_id,"—")
            tk.Label(srow,text=name[:36],width=34,bg=bg,fg=FG_DIM,font=("Consolas",9),anchor="w").pack(side="left",padx=3)
            rv=tk.StringVar(value=str(prev_r))
            tk.Entry(srow,textvariable=rv,width=7,bg=BG3,fg=FG,insertbackground=FG,font=("Consolas",9),relief="flat").pack(side="left",padx=3)
            slot_rate_vars.append(rv)
            cv=tk.StringVar(value=str(prev_c))
            tk.Entry(srow,textvariable=cv,width=7,bg=BG3,fg=FG,insertbackground=FG,font=("Consolas",9),relief="flat").pack(side="left",padx=3)
            slot_count_vars.append(cv)

        def gather():
            t=type_var.get()
            dc=len(slots) if t==2 else max(1,int(dc_var.get() or 1))
            sl=[]
            for rv,cv in zip(slot_rate_vars,slot_count_vars):
                try: r=max(1,min(32766,int(rv.get())))
                except: r=100
                try: c=max(1,min(32766,int(cv.get())))
                except: c=1
                sl.append({"rate":r,"count":c})
            return {"type":t,"drop_cnt":dc,"slots":sl}

        def save_and_advance():
            if type_var.get()==0:
                dc_raw=dc_var.get().strip()
                if not dc_raw or not dc_raw.isdigit() or int(dc_raw)<1:
                    if not messagebox.askyesno("Missed a spot","DropCnt is blank/invalid. Will default to 1. Continue?"): return
            cfg=gather(); self.manual_configs[rid]=cfg; self.manual_saved=cfg
            self.manual_idx+=1
            if self.manual_idx>=total: self._finish_manual()
            elif self.manual_continue_mode=="automate": self._automate_manual_remaining(cfg)
            elif self.manual_continue_mode=="monitor": self._build_manual_screen()
            else: self._ask_manual_mode(cfg)

        # Use nav2 which is pinned outside the scroll canvas
        mk_btn(nav2,"◀  Start Over",self._build_load_screen).pack(side="left",padx=8,pady=6)
        if idx>0:
            mk_btn(nav2,"◀  Prev",lambda:(setattr(self,"manual_idx",self.manual_idx-1),self._build_manual_screen())).pack(side="left",padx=4,pady=6)
        if self.manual_continue_mode:
            mk_btn(nav2,"⚙ Change Mode",lambda:(setattr(self,"manual_continue_mode",None),self._build_manual_screen()),color=BG4).pack(side="left",padx=4,pady=6)
        mk_btn(nav2,"Finish ✓" if idx==total-1 else "Next ▶",save_and_advance,color=GREEN,fg=BG2,font=("Consolas",10,"bold")).pack(side="right",padx=8,pady=6)

    def _ask_manual_mode(self, last_cfg):
        remaining=len(self.manual_matched)-self.manual_idx
        win=tk.Toplevel(self.root); win.title("Continue?"); win.geometry("520x240")
        win.configure(bg=BG); win.grab_set()
        tk.Label(win,text=f"{remaining} box(es) remaining.",bg=BG,fg=FG,font=("Consolas",13,"bold")).pack(pady=12)
        remember=tk.BooleanVar(value=False)
        bf=tk.Frame(win,bg=BG); bf.pack(pady=8)
        def choose(mode):
            if remember.get(): self.manual_continue_mode=mode
            win.destroy()
            if mode=="automate": self._automate_manual_remaining(last_cfg)
            else: self._build_manual_screen()
        mk_btn(bf,"🤖  Automate  —  copy settings to all remaining boxes",lambda:choose("automate"),color=ACC1,fg=BG2).pack(pady=4)
        mk_btn(bf,"👁  Monitor  —  review each box",lambda:choose("monitor"),color=BLUE,fg=BG2).pack(pady=4)
        tk.Checkbutton(win,text="Remember for rest of session",variable=remember,
                       bg=BG,fg=ACC4,selectcolor=BG3,activebackground=BG,font=("Consolas",9)).pack(pady=4)

    def _automate_manual_remaining(self, last_cfg):
        while self.manual_idx<len(self.manual_matched):
            rid,_,row_block=self.manual_matched[self.manual_idx]
            slots=real_drop_slots(row_block); cfg=copy.deepcopy(last_cfg)
            while len(cfg["slots"])<len(slots): cfg["slots"].append(cfg["slots"][-1] if cfg["slots"] else {"rate":100,"count":1})
            cfg["slots"]=cfg["slots"][:len(slots)]
            if cfg["type"]==2: cfg["drop_cnt"]=len(slots)
            self.manual_configs[rid]=cfg; self.manual_idx+=1
        self._finish_manual()

    def _finish_manual(self):
        csv_rows=[]
        def replace_row(m):
            row=m.group(0); rid=_get_tag(row,"Id")
            if rid not in self.manual_configs: return row
            new_row=apply_cfg_to_row(row,self.manual_configs[rid])
            drop_ids=[v for _,v in real_drop_slots(new_row)]
            name=next((n for r,n,_ in self.manual_matched if r==rid),"")
            csv_rows.append([rid,name,*drop_ids]); return new_row
        full_out=ROW_RE.sub(replace_row,self.xml_text)
        self._build_output_screen(full_out,csv_rows,len(self.manual_configs))

    def _build_output_screen(self, full_xml, csv_rows, count):
        self._clear()
        tk.Label(self,text=f"Done — {count} box(es) modified",font=("Consolas",13,"bold"),
                 bg=BG,fg=GREEN).pack(pady=12)
        nb=ttk.Notebook(self); nb.pack(fill="both",expand=True,padx=12,pady=4)
        max_items=max((len(r)-2 for r in csv_rows),default=0)
        header=["BoxID","BoxName"]+[f"Item{i+1}_ID" for i in range(max_items)]
        csv_cont="\n".join([",".join(header)]+[",".join(str(x) for x in r) for r in csv_rows])
        make_output_tab(nb,"Full PresentItemParam2.xml (modified)",full_xml,"PresentItemParam2_modified.xml",self.root)
        make_output_tab(nb,"Box Contents CSV (→ Tool 3)",csv_cont,"box_contents_for_tool3.csv",self.root)
        nb.select(0)
        bot=tk.Frame(self,bg=BG); bot.pack(fill="x",pady=6)
        def export_all():
            folder=filedialog.askdirectory(title="Choose export folder")
            if not folder: return
            for fname,content in [("PresentItemParam2_modified.xml",full_xml),("box_contents_for_tool3.csv",csv_cont)]:
                with open(os.path.join(folder,fname),"w",encoding="utf-8") as f: f.write(content)
            messagebox.showinfo("Export Complete",f"Saved to:\n{folder}")
        mk_btn(bot,"💾  Export All Files",export_all,color=ACC2,fg=BG2,font=("Consolas",11,"bold")).pack(side="left",padx=14)
        mk_btn(bot,"◀  Start Over",self._build_load_screen).pack(side="left",padx=4)

# ══════════════════════════════════════════════════════════════════════════════
# TOOL 3 — NCash / Ticket Updater  (simple CSV)
# ══════════════════════════════════════════════════════════════════════════════
class Tool3(tk.Frame):
    def __init__(self, parent, root, session):
        super().__init__(parent, bg=BG)
        self.root=root; self.session=session; self.csv_items=[]; self.xml_files=[]; self.item_lib={}
        self.mode_var=tk.StringVar(value="uniform"); self._build_load_screen()

    def _clear(self):
        for w in self.winfo_children(): w.destroy()

    def _build_load_screen(self):
        self._clear()
        tk.Label(self,text="NCASH / TICKET UPDATER",font=("Consolas",18,"bold"),bg=BG,fg=ACC3).pack(pady=(24,4))
        tk.Label(self,text="Formula: NCash = round(tickets × 133)",bg=BG,fg=FG_DIM,font=("Consolas",10)).pack(pady=(0,8))
        csv_status=tk.StringVar(value="No file loaded")
        xml_status=tk.StringVar(value="No file loaded")

        csv_frm=mk_section(self,"Step 1 — Box Contents CSV (from Tool 2, or ID list)")
        tk.Label(csv_frm,textvariable=csv_status,bg=BG,fg=FG_GREY,font=("Consolas",9)).pack(side="left",padx=10)
        def load_csv():
            p=filedialog.askopenfilename(filetypes=[("CSV","*.csv"),("All","*.*")])
            if not p: return
            with open(p,encoding="utf-8-sig") as f: text=f.read()
            items=parse_csv_text_t3(text)
            if not items: messagebox.showerror("Error","No item IDs found in CSV."); return
            self.csv_items=items
            if self.item_lib:
                for it in self.csv_items: it["name"]=self.item_lib.get(it["id"],"")
            csv_status.set(f"✓  {os.path.basename(p)}  —  {len(items)} items")
        mk_btn(csv_frm,"📂 Load",load_csv,padx=10,pady=4).pack(side="right",padx=8,pady=6)
        def import_session_t3():
            csv = self.session.box_id_list_csv or self.session.box_contents_csv
            if not csv:
                messagebox.showinfo("No Session Data","Run Tool 1 or 2 first to generate data.")
                return
            items = parse_csv_text_t3(csv)
            if not items:
                messagebox.showerror("Error","No item IDs found in session data."); return
            self.csv_items = items
            if self.item_lib:
                for it in self.csv_items: it["name"] = self.item_lib.get(it["id"],"")
            src_label = "Tool 1" if self.session.box_id_list_csv else "Tool 2"
            csv_status.set(f"✓  Imported from {src_label}  —  {len(items)} items")
        mk_btn(csv_frm,"⬇  Import from Tool 1/2",import_session_t3,color=ACC3,fg=BG2,padx=8,pady=4).pack(side="right",padx=4,pady=6)

        xml_frm=mk_section(self,"Step 2 — ItemParam XML (pick any one of the 4 files)")
        tk.Label(xml_frm,textvariable=xml_status,bg=BG,fg=FG_GREY,font=("Consolas",9)).pack(side="left",padx=10)
        def load_xml():
            p=filedialog.askopenfilename(title="Select any one of the 4 ItemParam XML files",
                                         filetypes=[("XML","*.xml"),("All","*.*")])
            if not p: return
            folder=os.path.dirname(p); loaded=[]
            for fname in os.listdir(folder):
                if fname.lower() in TARGET_FILES:
                    try:
                        with open(os.path.join(folder,fname),encoding="utf-8-sig",errors="replace") as f:
                            loaded.append((fname,f.read()))
                    except: pass
            if not loaded: messagebox.showerror("Error","None of the 4 ItemParam files found."); return
            self.xml_files=loaded; self.item_lib=build_item_lib(loaded)
            for it in self.csv_items: it["name"]=self.item_lib.get(it["id"],"")
            xml_status.set(f"✓  {len(loaded)}/4 files  |  {len(self.item_lib)} items indexed")
        mk_btn(xml_frm,"📂 Load",load_xml,padx=10,pady=4).pack(side="right",padx=8,pady=6)

        mode_frm=mk_section(self,"Step 3 — Mode")
        mf=tk.Frame(mode_frm,bg=BG); mf.pack(anchor="w",padx=10,pady=6)
        for lbl,val in [("Uniform  —  one ticket cost applied to every item","uniform"),
                        ("Manual   —  set ticket cost per item individually","manual")]:
            tk.Radiobutton(mf,text=lbl,variable=self.mode_var,value=val,
                           bg=BG,fg=FG,selectcolor=BG3,activebackground=BG,font=("Consolas",10)).pack(anchor="w",pady=2)

        def proceed():
            if not self.csv_items: messagebox.showwarning("Missing","Load a CSV first."); return
            if not self.xml_files: messagebox.showwarning("Missing","Load ItemParam XML first."); return
            if self.mode_var.get()=="uniform": self._build_uniform_screen()
            else: self._build_manual_screen()
        mk_btn(self,"▶  Continue →",proceed,color=GREEN,fg=BG2,font=("Consolas",12,"bold")).pack(pady=18)

    def _build_uniform_screen(self):
        self._clear()
        tk.Label(self,text="Uniform Ticket Cost",font=("Consolas",14,"bold"),bg=BG,fg=ACC3).pack(pady=(20,4))
        tk.Label(self,text=f"Applied to all {len(self.csv_items)} items.",bg=BG,fg=FG_DIM,font=("Consolas",10)).pack(pady=(0,12))
        frm=tk.Frame(self,bg=BG); frm.pack()
        tk.Label(frm,text="Ticket Cost:",bg=BG,fg=FG,font=("Consolas",12)).pack(side="left",padx=8)
        tv=tk.StringVar()
        ent=tk.Entry(frm,textvariable=tv,width=12,bg=BG3,fg=FG,insertbackground=FG,font=("Consolas",12),relief="flat")
        ent.pack(side="left",padx=8); ent.focus()
        ncash_var=tk.StringVar(value="NCash: —")
        tk.Label(self,textvariable=ncash_var,bg=BG,fg=GREEN,font=("Consolas",12,"bold")).pack(pady=8)
        def on_change(*_):
            try: ncash_var.set(f"NCash: {round(float(tv.get())*133)}")
            except: ncash_var.set("NCash: —")
        tv.trace_add("write",on_change)
        def apply_uniform():
            try: cost=float(tv.get())
            except: messagebox.showwarning("Invalid","Enter a valid ticket cost."); return
            for it in self.csv_items: it["ticket_cost"]=cost
            self._process_and_show()
        bot=tk.Frame(self,bg=BG); bot.pack(pady=16)
        mk_btn(bot,"◀  Back",self._build_load_screen).pack(side="left",padx=8)
        mk_btn(bot,"✓  Apply to All & Update XML",apply_uniform,color=GREEN,fg=BG2,font=("Consolas",11,"bold")).pack(side="left",padx=8)

    def _build_manual_screen(self):
        self._clear()
        tk.Label(self,text="Manual Ticket Costs",font=("Consolas",14,"bold"),bg=BG,fg=ACC3).pack(pady=(12,2))
        tk.Label(self,text="Leave blank to skip an item.",bg=BG,fg=FG_DIM,font=("Consolas",9)).pack(pady=(0,4))
        outer=tk.Frame(self,bg=BG); outer.pack(fill="both",expand=True,padx=20,pady=4)
        canvas,cont=mk_scroll_canvas(outer)
        hdr=tk.Frame(cont,bg=BG2); hdr.pack(fill="x",pady=2)
        for txt,w in [("Item ID",12),("Item Name",36),("Ticket Cost",14),("NCash (calc)",14)]:
            tk.Label(hdr,text=txt,width=w,bg=BG2,fg=BLUE,font=("Consolas",9,"bold"),anchor="w").pack(side="left",padx=6,pady=4)
        ticket_vars=[]
        for i,item in enumerate(self.csv_items):
            bg=BG if i%2==0 else BG2
            row=tk.Frame(cont,bg=bg); row.pack(fill="x")
            tk.Label(row,text=item["id"],width=12,bg=bg,fg=BG4,font=("Consolas",9),anchor="w").pack(side="left",padx=6,pady=2)
            name=item.get("name") or self.item_lib.get(item["id"],"—")
            tk.Label(row,text=name[:38],width=36,bg=bg,fg=FG_DIM,font=("Consolas",9),anchor="w").pack(side="left",padx=6,pady=2)
            tv=tk.StringVar()
            if item.get("ticket_cost") is not None: tv.set(str(item["ticket_cost"]))
            ticket_vars.append(tv)
            tk.Entry(row,textvariable=tv,width=12,bg=BG3,fg=FG,insertbackground=FG,font=("Consolas",9),relief="flat").pack(side="left",padx=6,pady=2)
            ncash_lbl=tk.Label(row,text="—",width=14,bg=bg,fg=GREEN,font=("Consolas",9),anchor="w")
            ncash_lbl.pack(side="left",padx=6)
            def make_trace(var,lbl):
                def cb(*_):
                    try: lbl.config(text=str(round(float(var.get())*133)))
                    except: lbl.config(text="—")
                var.trace_add("write",cb); cb()
            make_trace(tv,ncash_lbl)
        def confirm():
            blanks=[]
            for i,item in enumerate(self.csv_items):
                raw=ticket_vars[i].get().strip()
                try: item["ticket_cost"]=float(raw)
                except: item["ticket_cost"]=None; blanks.append(item["id"])
            if blanks:
                if not messagebox.askyesno("Missed a spot",f"{len(blanks)} item(s) will be SKIPPED:\n\n"+", ".join(blanks[:20])+("\n\nContinue anyway?")): return
            self._process_and_show()
        bot=tk.Frame(self,bg=BG); bot.pack(fill="x",pady=6)
        mk_btn(bot,"◀  Back",self._build_load_screen).pack(side="left",padx=14)
        mk_btn(bot,"✓  Apply & Update XML",confirm,color=GREEN,fg=BG2,font=("Consolas",11,"bold")).pack(side="right",padx=14)

    def _process_and_show(self):
        updates={it["id"]:round(it["ticket_cost"]*133) for it in self.csv_items if it["ticket_cost"] is not None}
        name_map={it["id"]:it.get("name","") for it in self.csv_items}
        file_results=[]
        for fname,text in self.xml_files:
            modified,found_map=bulk_update_ncash(text,updates)
            file_results.append((fname,modified,found_map))
        found_in={}
        for fname,_,found_map in file_results:
            for iid,hit in found_map.items():
                if hit and iid not in found_in: found_in[iid]=fname
        results=[]
        for item in self.csv_items:
            iid=item["id"]; name=name_map.get(iid,"")
            if item["ticket_cost"] is None: results.append((iid,name,None,None))
            else: results.append((iid,name,updates[iid],found_in.get(iid)))
        self._build_output_screen(file_results,results,updates)

    def _build_output_screen(self, file_results, results, updates):
        self._clear()
        updated_count=sum(1 for _,_,n,f in results if n is not None and f is not None)
        skipped_count=sum(1 for _,_,n,_ in results if n is None)
        missing_count=sum(1 for _,_,n,f in results if n is not None and f is None)
        tk.Label(self,text=f"✓ Updated: {updated_count}    ⚠ Not found: {missing_count}    — Skipped: {skipped_count}",
                 font=("Consolas",10,"bold"),bg=BG,fg=GREEN).pack(pady=8)
        nb=ttk.Notebook(self); nb.pack(fill="both",expand=True,padx=12,pady=4)
        exports=[]
        for fname,modified_text,found_map in file_results:
            if not any(hit for hit in found_map.values()): continue
            exports.append((fname,modified_text))
            make_output_tab(nb,os.path.splitext(fname)[0],modified_text,fname,self.root)
        col_hdr=f"{'ID':<15}{'Name':<34}{'NCash':<13}Status"; col_sep="─"*74
        log_parts=[]
        for fname,_,found_map in file_results:
            file_rows=[(iid,name,ncash,ff) for iid,name,ncash,ff in results if ff==fname]
            if not file_rows: log_parts.append(f"{fname}  →  No matching IDs — Skipped!\n"); continue
            log_parts.append(f"{fname}  →  {len(file_rows)} match(es)\n  {col_hdr}\n  {col_sep}")
            for iid,name,ncash,_ in file_rows: log_parts.append(f"  {iid:<15}{(name or '—')[:32]:<34}{ncash:<13}✓ Updated")
            log_parts.append("")
        unassigned=[(iid,name,ncash,ff) for iid,name,ncash,ff in results if ff is None]
        missing_rows=[(iid,name,ncash) for iid,name,ncash,_ in unassigned if ncash is not None]
        skipped_rows=[(iid,name) for iid,name,ncash,_ in unassigned if ncash is None]
        log_parts.append("── Unassigned / Skipped ──────────────────────────────────")
        for iid,name,ncash in missing_rows: log_parts.append(f"  {iid:<15}{(name or '—')[:32]:<34}{ncash:<13}⚠ Not found")
        for iid,name in skipped_rows: log_parts.append(f"  {iid:<15}{(name or '—')[:32]:<34}{'—':<13}SKIPPED")
        if not missing_rows and not skipped_rows: log_parts.append("  (none)")
        log_content="\n".join(log_parts)
        exports.append(("ncash_update_log.txt",log_content))
        make_output_tab(nb,"Update Log",log_content,"ncash_update_log.txt",self.root)
        nb.select(0)
        bot=tk.Frame(self,bg=BG); bot.pack(fill="x",pady=6)
        def export_all():
            folder=filedialog.askdirectory(title="Choose export folder")
            if not folder: return
            saved=[]
            for efname,content in exports:
                with open(os.path.join(folder,efname),"w",encoding="utf-8") as f: f.write(content)
                saved.append(efname)
            messagebox.showinfo("Export Complete",f"Saved to:\n{folder}\n\n"+"\n".join(saved))
        mk_btn(bot,"💾  Export All Files",export_all,color=ACC1,fg=BG2,font=("Consolas",11,"bold")).pack(side="left",padx=14)
        mk_btn(bot,"◀  Start Over",self._build_load_screen).pack(side="left",padx=4)

# ══════════════════════════════════════════════════════════════════════════════
# TOOL 4 — NCash Updater  (parent-box CSV + sub-box via PresentItemParam2)
# ══════════════════════════════════════════════════════════════════════════════
class Tool4(tk.Frame):
    def __init__(self, parent, root, session):
        super().__init__(parent, bg=BG)
        self.root=root; self.session=session
        self.parent_items=[]; self.xml_files=[]; self.item_lib={}
        self.present_text=None; self.present_enabled=tk.BooleanVar(value=False)
        self.parent_mode_var=tk.StringVar(value="uniform")
        self.sub_mode_var=tk.StringVar(value="uniform")
        self.sub_items=[]; self._build_load_screen()

    def _clear(self):
        for w in self.winfo_children(): w.destroy()

    def _build_load_screen(self):
        self._clear()
        tk.Label(self,text="NCASH UPDATER — PARENT BOX",font=("Consolas",16,"bold"),bg=BG,fg=ACC4).pack(pady=(18,2))
        tk.Label(self,text="Formula: NCash = round(tickets × 133)",bg=BG,fg=FG_DIM,font=("Consolas",9)).pack(pady=(0,6))
        outer=tk.Frame(self,bg=BG); outer.pack(fill="both",expand=True)
        scroll_host4=tk.Frame(outer,bg=BG); scroll_host4.pack(fill="both",expand=True)
        canvas,cont=mk_scroll_canvas(scroll_host4)
        # Session import panel
        if self.session.box_id_list_csv:
            sess_frm=mk_section(cont,"  ⬇  Session Import  ")
            def import_t4():
                items=parse_parentbox_csv(self.session.box_id_list_csv)
                if not items:
                    items=[{"id":i,"tickets":None,"ncash_direct":None,"box_ticket_cost":None,"group_idx":0,"name":n}
                           for i,n in self.session.box_id_map.items()]
                if not items: messagebox.showinfo("Empty","Session has no usable box IDs."); return
                self.parent_items=items
                messagebox.showinfo("Imported",f"Imported {len(items)} box IDs from Tool 1 session.")
            mk_btn(sess_frm,"Import box IDs from Tool 1",import_t4,color=ACC4,fg=BG2).pack(side="left",padx=10,pady=6)
            tk.Label(sess_frm,text=f"{len(parse_box_id_csv(self.session.box_id_list_csv))} IDs available",
                     bg=BG,fg=FG_GREY,font=("Consolas",9)).pack(side="left",padx=8)

        csv_status=tk.StringVar(value="No file loaded")
        xml_status=tk.StringVar(value="No file loaded")
        pres_status=tk.StringVar(value="Not loaded")

        s1=mk_section(cont,"  Step 1 — Parent-Box CSV  (ID, Tickets/NCash, Tickets of Box Contents)  ")
        tk.Label(s1,textvariable=csv_status,bg=BG,fg=FG_GREY,font=("Consolas",9)).pack(side="left",padx=10)
        def load_csv():
            p=filedialog.askopenfilename(filetypes=[("CSV","*.csv"),("All","*.*")])
            if not p: return
            with open(p,encoding="utf-8-sig") as f: text=f.read()
            items=parse_parentbox_csv(text)
            if not items: messagebox.showerror("Error","No valid item IDs found in CSV."); return
            self.parent_items=items
            csv_status.set(f"✓  {os.path.basename(p)}  —  {len(items)} IDs")
        mk_btn(s1,"📂 Load CSV",load_csv,padx=10,pady=4).pack(side="right",padx=8,pady=6)

        s2=mk_section(cont,"  Step 2 — ItemParam XML  (pick any one of the 4 files)  ")
        tk.Label(s2,textvariable=xml_status,bg=BG,fg=FG_GREY,font=("Consolas",9)).pack(side="left",padx=10)
        def load_xml():
            p=filedialog.askopenfilename(filetypes=[("XML","*.xml"),("All","*.*")])
            if not p: return
            folder=os.path.dirname(p); loaded=[]
            for fname in os.listdir(folder):
                if fname.lower() in TARGET_FILES:
                    try:
                        with open(os.path.join(folder,fname),encoding="utf-8-sig",errors="replace") as f:
                            loaded.append((fname,f.read()))
                    except: pass
            if not loaded: messagebox.showerror("Error","None of the 4 ItemParam files found."); return
            self.xml_files=loaded; self.item_lib=build_item_lib(loaded)
            for it in self.parent_items: it["name"]=self.item_lib.get(it["id"],"")
            # Also silently scan for PresentItemParam2.xml
            ppath=os.path.join(folder,PRESENT_FILE)
            if os.path.exists(ppath):
                try:
                    with open(ppath,encoding="utf-8-sig",errors="replace") as f:
                        self.present_text=f.read()
                    pres_status.set(f"✓  {PRESENT_FILE} auto-loaded")
                except: pass
            xml_status.set(f"✓  {len(loaded)}/4 files  |  {len(self.item_lib)} items indexed")
        mk_btn(s2,"📂 Load",load_xml,padx=10,pady=4).pack(side="right",padx=8,pady=6)

        s3=mk_section(cont,"  Step 3 — Mode  (parent-box IDs)  ")
        mf=tk.Frame(s3,bg=BG); mf.pack(anchor="w",padx=10,pady=6)
        for lbl,val in [("Uniform  —  one value per group","uniform"),("Manual  —  set per item","manual")]:
            tk.Radiobutton(mf,text=lbl,variable=self.parent_mode_var,value=val,
                           bg=BG,fg=FG,selectcolor=BG3,activebackground=BG,font=("Consolas",10)).pack(anchor="w",pady=2)

        s4=mk_section(cont,"  Step 4 (Optional) — Sub-box NCash via PresentItemParam2  ")
        tk.Label(s4,textvariable=pres_status,bg=BG,fg=FG_GREY,font=("Consolas",9)).pack(anchor="w",padx=10,pady=(4,0))
        tk.Checkbutton(s4,text="Enable sub-box NCash update via PresentItemParam2",
                       variable=self.present_enabled,bg=BG,fg=FG,selectcolor=BG3,
                       activebackground=BG,font=("Consolas",10)).pack(anchor="w",padx=10,pady=4)
        def load_present_manual():
            p=filedialog.askopenfilename(filetypes=[("XML","*.xml"),("All","*.*")])
            if not p: return
            with open(p,encoding="utf-8-sig",errors="replace") as f: self.present_text=f.read()
            pres_status.set(f"✓  {os.path.basename(p)}")
        mk_btn(s4,"📂 Load PresentItemParam2.xml manually",load_present_manual,padx=10,pady=4).pack(anchor="w",padx=10,pady=(0,6))

        s5=mk_section(cont,"  Step 5 (Optional) — Sub-box Mode  ")
        sf=tk.Frame(s5,bg=BG); sf.pack(anchor="w",padx=10,pady=6)
        for lbl,val in [("Uniform  —  one value for all sub items","uniform"),("Manual  —  set per sub item","manual")]:
            tk.Radiobutton(sf,text=lbl,variable=self.sub_mode_var,value=val,
                           bg=BG,fg=FG,selectcolor=BG3,activebackground=BG,font=("Consolas",10)).pack(anchor="w",pady=2)

        bot_frm=tk.Frame(cont,bg=BG); bot_frm.pack(fill="x",pady=10)
        def proceed():
            if not self.parent_items: messagebox.showwarning("Missing","Load a CSV first."); return
            if not self.xml_files: messagebox.showwarning("Missing","Load ItemParam XML first."); return
            if self.parent_mode_var.get()=="uniform": self._build_uniform_screen()
            else: self._build_manual_screen(self.parent_items,"parent",self._after_parent_configured)
        mk_btn(bot_frm,"▶  Continue →",proceed,color=GREEN,fg=BG2,font=("Consolas",12,"bold")).pack(pady=10)

    # ── Uniform screen (group-aware) ──────────────────────────────────────────
    def _build_uniform_screen(self, _saved_group_vals=None):
        self._clear()
        groups={}
        for it in self.parent_items:
            gi=it.get("group_idx",0)
            groups.setdefault(gi,[]).append(it)
        group_keys=sorted(groups.keys())
        saved_vals=_saved_group_vals or {}
        confirmed_vals={}  # gi -> {ticket_cost or ncash_direct}
        current_group_pos=[0]

        def show_group(pos):
            self._clear()
            gi=group_keys[pos]
            group_items=groups[gi]
            sample_names=[self.item_lib.get(it["id"],"") for it in group_items if self.item_lib.get(it["id"],"")][:3]
            tk.Label(self,text=f"Uniform Settings — Group {pos+1} of {len(group_keys)}",
                     font=("Consolas",14,"bold"),bg=BG,fg=ACC4).pack(pady=(18,4))
            if sample_names: tk.Label(self,text="e.g. "+", ".join(sample_names),bg=BG,fg=FG_DIM,font=("Consolas",9)).pack()

            if confirmed_vals:
                prev_frm=tk.LabelFrame(self,text="  Previously confirmed  ",bg=BG,fg=FG_GREY,
                                       font=("Consolas",9),bd=1,relief="groove")
                prev_frm.pack(fill="x",padx=24,pady=4)
                for prev_gi,pval in confirmed_vals.items():
                    pg_items=groups[prev_gi]
                    pg_sample=[self.item_lib.get(it["id"],"") for it in pg_items if self.item_lib.get(it["id"],"")][:2]
                    desc=", ".join(pg_sample) or f"Group {prev_gi+1}"
                    if pval.get("ticket_cost") is not None: disp=f"Tickets={pval['ticket_cost']} → NCash={round(pval['ticket_cost']*133)}"
                    else: disp=f"NCash={pval.get('ncash_direct','?')}"
                    tk.Label(prev_frm,text=f"  Group {list(group_keys).index(prev_gi)+1}: {desc[:40]}  →  {disp}",
                             bg=BG,fg=FG_DIM,font=("Consolas",9)).pack(anchor="w",padx=6,pady=1)

            type_var=tk.StringVar(value="tickets")
            sample_it=group_items[0]
            if sample_it.get("ticket_cost") is not None: type_var.set("tickets"); init_val=str(sample_it["ticket_cost"])
            elif sample_it.get("ncash_direct") is not None: type_var.set("ncash"); init_val=str(sample_it["ncash_direct"])
            else: init_val=saved_vals.get(gi,{}).get("init_val","")
            val_var=tk.StringVar(value=init_val)

            inp_frm=tk.Frame(self,bg=BG); inp_frm.pack(pady=10)
            tk.Radiobutton(inp_frm,text="Tickets ×133",variable=type_var,value="tickets",bg=BG,fg=FG,selectcolor=BG3,activebackground=BG,font=("Consolas",10)).pack(side="left",padx=8)
            tk.Radiobutton(inp_frm,text="NCash exact",variable=type_var,value="ncash",bg=BG,fg=FG,selectcolor=BG3,activebackground=BG,font=("Consolas",10)).pack(side="left",padx=8)
            ent_frm=tk.Frame(self,bg=BG); ent_frm.pack(pady=6)
            tk.Label(ent_frm,text="Value:",bg=BG,fg=FG,font=("Consolas",11)).pack(side="left",padx=8)
            ent=tk.Entry(ent_frm,textvariable=val_var,width=14,bg=BG3,fg=FG,insertbackground=FG,font=("Consolas",12),relief="flat")
            ent.pack(side="left",padx=8); ent.focus()
            result_lbl=tk.Label(self,text="",bg=BG,fg=GREEN,font=("Consolas",11,"bold"))
            result_lbl.pack(pady=4)
            def update_result(*_):
                try:
                    if type_var.get()=="tickets": result_lbl.config(text=f"→ NCash: {round(float(val_var.get())*133)}")
                    else: result_lbl.config(text=f"→ Approx tickets: {round(float(val_var.get())/133,4)}")
                except: result_lbl.config(text="")
            val_var.trace_add("write",update_result); type_var.trace_add("write",update_result); update_result()

            def confirm_group():
                try: v=float(val_var.get())
                except: messagebox.showwarning("Invalid","Enter a valid number."); return
                if type_var.get()=="tickets": confirmed_vals[gi]={"ticket_cost":v,"ncash_direct":None}
                else: confirmed_vals[gi]={"ticket_cost":None,"ncash_direct":int(round(v))}
                if pos+1<len(group_keys): show_group(pos+1)
                else:
                    for it in self.parent_items:
                        gval=confirmed_vals.get(it.get("group_idx",0),{})
                        it["ticket_cost"]=gval.get("ticket_cost"); it["ncash_direct"]=gval.get("ncash_direct")
                    self._after_parent_configured()

            def go_back():
                if pos>0:
                    prev_gi=group_keys[pos-1]
                    confirmed_vals.pop(prev_gi,None)
                    show_group(pos-1)
                else: self._build_load_screen()

            nav=tk.Frame(self,bg=BG); nav.pack(pady=14)
            mk_btn(nav,"◀  Back",go_back).pack(side="left",padx=10)
            mk_btn(nav,"Confirm ▶" if pos<len(group_keys)-1 else "✓  Apply All",confirm_group,color=GREEN,fg=BG2,font=("Consolas",11,"bold")).pack(side="left",padx=10)

        show_group(0)

    # ── Manual screen (parent or sub) ─────────────────────────────────────────
    def _build_manual_screen(self, items, label, on_done):
        self._clear()
        tk.Label(self,text=f"Manual — {label.title()} Item Costs",font=("Consolas",14,"bold"),bg=BG,fg=ACC4).pack(pady=(12,2))
        tk.Label(self,text="Leave blank to skip.",bg=BG,fg=FG_DIM,font=("Consolas",9)).pack(pady=(0,4))
        outer=tk.Frame(self,bg=BG); outer.pack(fill="both",expand=True,padx=20,pady=4)
        canvas,cont=mk_scroll_canvas(outer)
        hdr=tk.Frame(cont,bg=BG2); hdr.pack(fill="x",pady=2)
        for txt,w in [("Item ID",12),("Item Name",34),("Type",12),("Value",14),("NCash →",14)]:
            tk.Label(hdr,text=txt,width=w,bg=BG2,fg=BLUE,font=("Consolas",9,"bold"),anchor="w").pack(side="left",padx=4,pady=4)
        type_vars=[]; val_vars=[]
        for i,item in enumerate(items):
            bg=BG if i%2==0 else BG2
            row=tk.Frame(cont,bg=bg); row.pack(fill="x")
            tk.Label(row,text=item["id"],width=12,bg=bg,fg=BG4,font=("Consolas",9),anchor="w").pack(side="left",padx=4,pady=2)
            name=item.get("name") or self.item_lib.get(item["id"],"—")
            tk.Label(row,text=name[:36],width=34,bg=bg,fg=FG_DIM,font=("Consolas",9),anchor="w").pack(side="left",padx=4,pady=2)
            if item.get("ticket_cost") is not None: init_type,init_val="tickets",str(item["ticket_cost"])
            elif item.get("ncash_direct") is not None: init_type,init_val="ncash",str(item["ncash_direct"])
            else: init_type,init_val="tickets",""
            tv=tk.StringVar(value=init_type); type_vars.append(tv)
            type_combo=ttk.Combobox(row,textvariable=tv,values=["tickets","ncash"],state="readonly",width=8,font=("Consolas",9))
            type_combo.pack(side="left",padx=4,pady=2)
            vv=tk.StringVar(value=init_val); val_vars.append(vv)
            tk.Entry(row,textvariable=vv,width=12,bg=BG3,fg=FG,insertbackground=FG,font=("Consolas",9),relief="flat").pack(side="left",padx=4,pady=2)
            ncash_lbl=tk.Label(row,text="—",width=14,bg=bg,fg=GREEN,font=("Consolas",9),anchor="w")
            ncash_lbl.pack(side="left",padx=4)
            def make_trace(tv2,vv2,lbl):
                def cb(*_):
                    try:
                        v=float(vv2.get())
                        if tv2.get()=="tickets": lbl.config(text=str(round(v*133)))
                        else: lbl.config(text=str(int(round(v))))
                    except: lbl.config(text="—")
                tv2.trace_add("write",cb); vv2.trace_add("write",cb); cb()
            make_trace(tv,vv,ncash_lbl)

        def confirm():
            blanks=[]
            for i,item in enumerate(items):
                raw=val_vars[i].get().strip()
                try:
                    v=float(raw)
                    if type_vars[i].get()=="tickets": item["ticket_cost"]=v; item["ncash_direct"]=None
                    else: item["ncash_direct"]=int(round(v)); item["ticket_cost"]=None
                except: item["ticket_cost"]=None; item["ncash_direct"]=None; blanks.append(item["id"])
            if blanks:
                if not messagebox.askyesno("Missed a spot",f"{len(blanks)} item(s) will be SKIPPED. Continue anyway?"): return
            on_done()

        bot=tk.Frame(self,bg=BG); bot.pack(fill="x",pady=6)
        mk_btn(bot,"◀  Back",self._build_load_screen).pack(side="left",padx=14)
        mk_btn(bot,"✓  Apply",confirm,color=GREEN,fg=BG2,font=("Consolas",11,"bold")).pack(side="right",padx=14)

    def _after_parent_configured(self):
        if self.present_enabled.get() and self.present_text:
            self._build_sub_configure_screen()
        else:
            self._process_and_show()

    def _build_sub_configure_screen(self):
        box_ids={it["id"] for it in self.parent_items}
        drop_map=extract_drop_ids_from_present(self.present_text, box_ids)
        self.sub_items=[]
        seen=set()
        for it in self.parent_items:
            for drop_id in drop_map.get(it["id"],[]):
                if drop_id not in seen:
                    seen.add(drop_id)
                    tc=it.get("box_ticket_cost")
                    self.sub_items.append({"id":drop_id,"name":self.item_lib.get(drop_id,""),
                                           "ticket_cost":tc,"ncash_direct":None,"box_ticket_cost":tc,"group_idx":0})
        if not self.sub_items:
            messagebox.showinfo("No sub-items","No DropId entries found for the parent box IDs in PresentItemParam2.")
            self._process_and_show(); return
        all_prefilled=all(it.get("ticket_cost") is not None or it.get("ncash_direct") is not None for it in self.sub_items)
        if all_prefilled:
            self._process_and_show(); return
        if self.sub_mode_var.get()=="uniform":
            self._build_sub_uniform_screen()
        else:
            self._build_manual_screen(self.sub_items,"sub-box",self._process_and_show)

    def _build_sub_uniform_screen(self):
        self._clear()
        tk.Label(self,text="Uniform Sub-Box Tickets",font=("Consolas",14,"bold"),bg=BG,fg=ACC4).pack(pady=(20,4))
        tk.Label(self,text=f"Applied to all {len(self.sub_items)} sub-box drop IDs.",bg=BG,fg=FG_DIM,font=("Consolas",10)).pack(pady=(0,12))
        sample_it=next((it for it in self.sub_items if it.get("ticket_cost") is not None),None)
        init_val=str(sample_it["ticket_cost"]) if sample_it else ""
        frm=tk.Frame(self,bg=BG); frm.pack()
        tv=tk.StringVar(value=init_val)
        tk.Label(frm,text="Ticket Cost:",bg=BG,fg=FG,font=("Consolas",12)).pack(side="left",padx=8)
        ent=tk.Entry(frm,textvariable=tv,width=14,bg=BG3,fg=FG,insertbackground=FG,font=("Consolas",12),relief="flat")
        ent.pack(side="left",padx=8); ent.focus()
        ncash_lbl=tk.Label(self,text="",bg=BG,fg=GREEN,font=("Consolas",12,"bold")); ncash_lbl.pack(pady=6)
        def on_change(*_):
            try: ncash_lbl.config(text=f"→ NCash: {round(float(tv.get())*133)}")
            except: ncash_lbl.config(text="")
        tv.trace_add("write",on_change)
        def apply():
            try: cost=float(tv.get())
            except: messagebox.showwarning("Invalid","Enter a valid ticket cost."); return
            for it in self.sub_items: it["ticket_cost"]=cost; it["ncash_direct"]=None
            self._process_and_show()
        bot=tk.Frame(self,bg=BG); bot.pack(pady=16)
        mk_btn(bot,"◀  Back",self._build_load_screen).pack(side="left",padx=8)
        mk_btn(bot,"✓  Apply & Update XML",apply,color=GREEN,fg=BG2,font=("Consolas",11,"bold")).pack(side="left",padx=8)

    def _resolve_ncash(self, item):
        if item.get("ticket_cost") is not None: return round(item["ticket_cost"]*133)
        if item.get("ncash_direct") is not None: return item["ncash_direct"]
        return None

    def _process_and_show(self):
        all_items=list(self.parent_items)+list(self.sub_items)
        updates={}
        for it in all_items:
            n=self._resolve_ncash(it)
            if n is not None: updates[it["id"]]=n
        file_results=[]
        for fname,text in self.xml_files:
            modified,found_map=bulk_update_ncash(text,updates)
            file_results.append((fname,modified,found_map))
        found_in={}
        for fname,_,found_map in file_results:
            for iid,hit in found_map.items():
                if hit and iid not in found_in: found_in[iid]=fname
        parent_ids={it["id"] for it in self.parent_items}
        sub_ids={it["id"] for it in self.sub_items}
        updated_p=sum(1 for iid in parent_ids if found_in.get(iid))
        updated_s=sum(1 for iid in sub_ids if found_in.get(iid))
        missing=sum(1 for iid,ncash in updates.items() if not found_in.get(iid))
        skipped=sum(1 for it in all_items if self._resolve_ncash(it) is None)
        self._build_output_screen(file_results,updates,found_in,parent_ids,sub_ids,updated_p,updated_s,missing,skipped)

    def _build_output_screen(self, file_results, updates, found_in, parent_ids, sub_ids,
                              updated_p, updated_s, missing, skipped):
        self._clear()
        total_updated=updated_p+updated_s
        tk.Label(self,text=f"✓ Updated: {total_updated}  (parent: {updated_p}, sub-box: {updated_s})   ⚠ Not found: {missing}   — Skipped: {skipped}",
                 font=("Consolas",10,"bold"),bg=BG,fg=GREEN).pack(pady=8)
        nb=ttk.Notebook(self); nb.pack(fill="both",expand=True,padx=12,pady=4)
        exports=[]
        for fname,modified_text,found_map in file_results:
            if not any(hit for hit in found_map.values()): continue
            exports.append((fname,modified_text))
            make_output_tab(nb,os.path.splitext(fname)[0],modified_text,fname,self.root)
        log_parts=[]
        for fname,_,found_map in file_results:
            hits=[(iid,updates[iid],found_in.get(iid)=="parent" or iid in parent_ids) for iid,hit in found_map.items() if hit]
            if not hits: log_parts.append(f"{fname}  →  No matches"); continue
            log_parts.append(f"{fname}  →  {len(hits)} update(s)")
            for iid,ncash,is_parent in hits:
                label="parent" if iid in parent_ids else "sub-drop"
                name=self.item_lib.get(iid,"—")
                log_parts.append(f"  [{label}]  {iid:<12}  {name[:30]:<30}  NCash={ncash}")
            log_parts.append("")
        log_content="\n".join(log_parts)
        exports.append(("ncash_update_log.txt",log_content))
        make_output_tab(nb,"Update Log",log_content,"ncash_update_log.txt",self.root)
        nb.select(0)
        bot=tk.Frame(self,bg=BG); bot.pack(fill="x",pady=6)
        def export_all():
            folder=filedialog.askdirectory(title="Choose export folder")
            if not folder: return
            for efname,content in exports:
                with open(os.path.join(folder,efname),"w",encoding="utf-8") as f: f.write(content)
            messagebox.showinfo("Export Complete",f"Saved to:\n{folder}")
        mk_btn(bot,"💾  Export All Files",export_all,color=ACC4,fg=BG2,font=("Consolas",11,"bold")).pack(side="left",padx=14)
        mk_btn(bot,"◀  Start Over",self._build_load_screen).pack(side="left",padx=4)

# ══════════════════════════════════════════════════════════════════════════════
# TOOL 5 — NCash ↔ Ticket Calculator
# ══════════════════════════════════════════════════════════════════════════════
class Tool5(tk.Frame):
    def __init__(self, parent, root, session):
        super().__init__(parent, bg=BG)
        self.root=root; self.session=session; self._build()

    def _build(self):
        for w in self.winfo_children(): w.destroy()
        tk.Label(self,text="NCash  ↔  Ticket  Calculator",font=("Consolas",16,"bold"),
                 bg=BG,fg=ACC5).pack(pady=(30,4))
        tk.Label(self,text="Formula:  NCash = round( Tickets × 133 )",
                 font=("Consolas",9),bg=BG,fg=FG_GREY).pack(pady=(0,18))

        box_a=tk.LabelFrame(self,text="  Tickets  →  NCash  ",bg=BG,fg=BLUE,
                             font=("Consolas",10,"bold"),bd=1,relief="groove")
        box_a.pack(fill="x",padx=40,pady=8)
        ra=tk.Frame(box_a,bg=BG); ra.pack(padx=14,pady=12)
        tk.Label(ra,text="Tickets:",width=10,anchor="w",font=("Consolas",11),bg=BG,fg=FG).pack(side="left")
        v_tickets=tk.StringVar()
        tk.Entry(ra,textvariable=v_tickets,width=14,bg=BG3,fg=FG,insertbackground=FG,
                 font=("Consolas",13),relief="flat").pack(side="left",padx=8)
        tk.Label(ra,text="=",font=("Consolas",13),bg=BG,fg=FG_GREY).pack(side="left",padx=4)
        lbl_ncash=tk.Label(ra,text="—",width=14,anchor="w",font=("Consolas",13,"bold"),bg=BG,fg=GREEN)
        lbl_ncash.pack(side="left",padx=4)
        tk.Label(ra,text="NCash",font=("Consolas",10),bg=BG,fg=FG_GREY).pack(side="left")

        box_b=tk.LabelFrame(self,text="  NCash  →  Tickets  ",bg=BG,fg=BLUE,
                             font=("Consolas",10,"bold"),bd=1,relief="groove")
        box_b.pack(fill="x",padx=40,pady=8)
        rb=tk.Frame(box_b,bg=BG); rb.pack(padx=14,pady=12)
        tk.Label(rb,text="NCash:",width=10,anchor="w",font=("Consolas",11),bg=BG,fg=FG).pack(side="left")
        v_ncash=tk.StringVar()
        tk.Entry(rb,textvariable=v_ncash,width=14,bg=BG3,fg=FG,insertbackground=FG,
                 font=("Consolas",13),relief="flat").pack(side="left",padx=8)
        tk.Label(rb,text="=",font=("Consolas",13),bg=BG,fg=FG_GREY).pack(side="left",padx=4)
        lbl_tickets=tk.Label(rb,text="—",width=14,anchor="w",font=("Consolas",13,"bold"),bg=BG,fg=ACC5)
        lbl_tickets.pack(side="left",padx=4)
        tk.Label(rb,text="Tickets",font=("Consolas",10),bg=BG,fg=FG_GREY).pack(side="left")

        def calc_ncash(*_):
            try: lbl_ncash.config(text=f"{round(float(v_tickets.get())*133):,}")
            except: lbl_ncash.config(text="—")
        def calc_tickets(*_):
            try:
                raw=float(v_ncash.get())/133
                lbl_tickets.config(text=f"{int(raw):,}" if raw==int(raw) else f"{round(raw,4):,}")
            except: lbl_tickets.config(text="—")
        v_tickets.trace_add("write",calc_ncash)
        v_ncash.trace_add("write",calc_tickets)

# ══════════════════════════════════════════════════════════════════════════════
# COMBINED SHELL
# ══════════════════════════════════════════════════════════════════════════════
TOOLS = [
    ("1", "Box XML\nGenerator",   ACC1, Tool1),
    ("2", "Rate / Count\nAdjuster", ACC2, Tool2),
    ("3", "NCash Updater\n(Simple)", ACC3, Tool3),
    ("4", "NCash Updater\n(Parent Box)", ACC4, Tool4),
    ("5", "NCash ↔ Ticket\nCalculator", ACC5, Tool5),
]

class CombinedApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Box Tool Suite")
        self.geometry("1100x820")
        self.configure(bg=BG2)
        self._current_tool = None
        self._tool_instances = {}
        self._nav_buttons = {}
        self.session = AppSession()
        self._build_layout()
        self._switch_tool(0)

    def _build_layout(self):
        # ── Left sidebar ──────────────────────────────────────────────────
        sidebar = tk.Frame(self, bg=BG2, width=148)
        sidebar.pack(side="left", fill="y")
        sidebar.pack_propagate(False)

        tk.Label(sidebar, text="BOX\nTOOL\nSUITE", font=("Consolas",13,"bold"),
                 bg=BG2, fg=FG, justify="center").pack(pady=(20,16))

        sep = tk.Frame(sidebar, bg=BG4, height=1); sep.pack(fill="x", padx=10, pady=4)

        for i,(num,label,color,_) in enumerate(TOOLS):
            frm = tk.Frame(sidebar, bg=BG2, cursor="hand2")
            frm.pack(fill="x", padx=8, pady=3)
            dot = tk.Label(frm, text="●", font=("Consolas",9), bg=BG2, fg=color, width=2)
            dot.pack(side="left")
            btn = tk.Button(frm, text=f"  {label}", font=("Consolas",9),
                            bg=BG2, fg=FG_DIM, relief="flat", anchor="w",
                            justify="left", padx=4, pady=6,
                            activebackground=BG3, activeforeground=FG,
                            command=lambda idx=i: self._switch_tool(idx))
            btn.pack(side="left", fill="x", expand=True)
            self._nav_buttons[i] = (frm, btn, dot, color)
            frm.bind("<Button-1>", lambda e, idx=i: self._switch_tool(idx))

        sep2 = tk.Frame(sidebar, bg=BG4, height=1); sep2.pack(fill="x", padx=10, pady=(14,4))
        tk.Label(sidebar, text="Session data\nflows between\ntools. Use\n'Import Session'\nin each tool.",
                 font=("Consolas",7), bg=BG2, fg=FG_GREY, justify="left").pack(padx=12, pady=6, anchor="w")

        # ── Content area ──────────────────────────────────────────────────
        self._content = tk.Frame(self, bg=BG)
        self._content.pack(side="left", fill="both", expand=True)

    def _switch_tool(self, idx):
        if self._current_tool == idx: return

        # Update nav highlight
        for i,(frm,btn,dot,color) in self._nav_buttons.items():
            if i == idx:
                frm.config(bg=BG3); btn.config(bg=BG3, fg=FG); dot.config(bg=BG3, fg=color)
            else:
                frm.config(bg=BG2); btn.config(bg=BG2, fg=FG_DIM); dot.config(bg=BG2, fg=self._nav_buttons[i][3])

        # Hide current
        if self._current_tool is not None:
            if self._current_tool in self._tool_instances:
                self._tool_instances[self._current_tool].pack_forget()

        # Instantiate or show
        if idx not in self._tool_instances:
            _,_,_,ToolClass = TOOLS[idx]
            instance = ToolClass(self._content, self, self.session)
            self._tool_instances[idx] = instance

        self._tool_instances[idx].pack(fill="both", expand=True)
        self._current_tool = idx
        num,label,_,_ = TOOLS[idx]
        self.title(f"Box Tool Suite  —  Tool {num}: {label.replace(chr(10),' ')}")


if __name__ == "__main__":
    CombinedApp().mainloop()
