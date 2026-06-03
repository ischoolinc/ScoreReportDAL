# -*- coding: utf-8 -*-
"""Inspect row order and subject numbers in expanded docx tables."""
import re
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

sys.stdout.reconfigure(encoding="utf-8")

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def wtag(name):
    return f"{{{W}}}{name}"


def row_subject_num(tr):
    parts = []
    in_f = False
    texts = []
    for elem in tr.iter():
        tag = elem.tag.split("}")[-1]
        if tag == "fldChar":
            ft = elem.get(wtag("fldCharType"))
            if ft == "begin":
                texts = []
                in_f = True
            elif ft == "end" and in_f:
                s = "".join(texts)
                m = re.search(r"MERGEFIELD\s+(.+?)(?:\s+\\|\s+\*|$)", s)
                if m:
                    parts.append(m.group(1).strip())
                in_f = False
        elif tag == "instrText" and in_f and elem.text:
            texts.append(elem.text)
    for name in parts:
        m = re.match(r"科目(\d+)_", name)
        if m:
            return int(m.group(1))
    # split pattern: 科目 + N_
    full = "".join(t.text or "" for t in tr.iter(wtag("instrText")))
    m = re.search(r"科目(\d+)_", full)
    if m:
        return int(m.group(1))
    m2 = re.search(r"(\d+)_", full)
    if m2 and "科目" in full:
        return int(m2.group(1))
    return None


def first_cell_text(tr):
    cells = tr.findall(wtag("tc"))
    if not cells:
        return ""
    return "".join(t.text or "" for t in cells[0].iter(wtag("t"))).strip()


label = sys.argv[1] if len(sys.argv) > 1 else "NEW"
fname = (
    "畢業預警通知單合併欄位總表.docx.bak"
    if label == "BAK"
    else "畢業預警通知單合併欄位總表.docx"
)
p = Path(__file__).resolve().parent.parent / "Template" / fname
print(f"File: {fname}")
with zipfile.ZipFile(p) as z:
    root = ET.fromstring(z.read("word/document.xml"))

tables = root.findall(f".//{wtag('tbl')}")
for ti in [6, 7, 10, 11]:
    tbl = tables[ti]
    rows = tbl.findall(wtag("tr"))
    print(f"\n=== Table {ti}: {len(rows)} rows ===")
    data_rows = []
    for ri, tr in enumerate(rows):
        sn = row_subject_num(tr)
        fc = first_cell_text(tr)
        if sn is not None:
            data_rows.append((ri, fc, sn))

    print(f"Data rows with subject fields: {len(data_rows)}")
    # show around 45-55 and end
    for ri, fc, sn in data_rows:
        if sn <= 5 or 45 <= sn <= 55 or sn >= 125:
            print(f"  xml_row={ri:4d}  first_cell={fc!r:6s}  merge_subject={sn}")

    nums = [sn for _, _, sn in data_rows]
    # detect gaps/duplicates
    issues = []
    for i in range(1, len(nums)):
        if nums[i] != nums[i - 1] + 1 and not (nums[i - 1] is None):
            issues.append((i, nums[i - 1], nums[i]))
    if issues:
        print("  ORDER ISSUES (prev -> next):")
        for i, a, b in issues[:20]:
            print(f"    index {i}: {a} -> {b}")
    dup = [n for n in set(nums) if nums.count(n) > 1]
    if dup:
        print(f"  DUPLICATES: {sorted(dup)[:20]}...")
    missing = [n for n in range(1, 131) if n not in nums]
    if missing:
        print(f"  MISSING nums: {missing[:30]}{'...' if len(missing)>30 else ''}")
