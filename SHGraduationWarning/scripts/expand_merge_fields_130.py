# -*- coding: utf-8 -*-
"""Expand subject merge fields in 畢業預警通知單合併欄位總表.docx from 50 to 130."""
import re
import shutil
import zipfile
from copy import deepcopy
from pathlib import Path

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def wtag(name):
    return f"{{{W}}}{name}"


def table_rows(tbl):
    """Return only direct child w:tr elements (not nested table rows)."""
    return [ch for ch in tbl if ch.tag == wtag("tr")]


def parse_merge_fields_in_element(elem):
    """Return list of merge field names in document order within elem."""
    names = []
    for fld in elem.iter(wtag("fldSimple")):
        instr = fld.get(wtag("instr")) or ""
        m = re.search(r"MERGEFIELD\s+(.+?)(?:\s+\\|\s+\*|$)", instr)
        if m:
            names.append(m.group(1).strip())

    for p_elem in elem.iter(wtag("p")):
        texts = []
        in_field = False
        for child in p_elem.iter():
            tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
            if tag == "fldChar":
                ftype = child.get(wtag("fldCharType"))
                if ftype == "begin":
                    texts = []
                    in_field = True
                elif ftype == "end" and in_field:
                    full = "".join(texts)
                    m = re.search(r"MERGEFIELD\s+(.+?)(?:\s+\\|\s+\*|$)", full)
                    if m:
                        names.append(m.group(1).strip())
                    in_field = False
            elif tag == "instrText" and in_field and child.text:
                texts.append(child.text)
    return names


def replace_subject_number_in_element(elem, old_num, new_num):
    """Replace subject index in merge fields (Word may split 科目 / 50_ / suffix)."""
    old_prefix = f"科目{old_num}_"
    new_prefix = f"科目{new_num}_"
    old_part = f"{old_num}_"
    new_part = f"{new_num}_"

    for fld in elem.iter(wtag("fldSimple")):
        instr = fld.get(wtag("instr"))
        if not instr:
            continue
        if old_prefix in instr:
            fld.set(wtag("instr"), instr.replace(old_prefix, new_prefix))
        elif old_part in instr:
            fld.set(wtag("instr"), instr.replace(old_part, new_part))

    for it in elem.iter(wtag("instrText")):
        if not it.text:
            continue
        if old_prefix in it.text:
            it.text = it.text.replace(old_prefix, new_prefix)
        elif it.text == old_part:
            it.text = new_part

    for t in elem.iter(wtag("t")):
        if not t.text:
            continue
        if old_prefix in t.text:
            t.text = t.text.replace(old_prefix, new_prefix)
        elif t.text == old_part:
            t.text = new_part


def replace_row_number_in_cells(tr, old_num, new_num):
    """Replace visible row number in first column (non-merge text)."""
    old_s = str(old_num)
    new_s = str(new_num)
    cells = tr.findall(wtag("tc"))
    if not cells:
        return
    first_tc = cells[0]
    for t in first_tc.iter(wtag("t")):
        if t.text and t.text.strip() == old_s:
            t.text = new_s
            return


def row_max_subject_num(tr):
    names = parse_merge_fields_in_element(tr)
    max_n = 0
    for n in names:
        m = re.match(r"科目(\d+)_", n)
        if m:
            max_n = max(max_n, int(m.group(1)))
    return max_n, names


def first_cell_number(tr):
    cells = tr.findall(wtag("tc"))
    if not cells:
        return None
    text = "".join(t.text or "" for t in cells[0].iter(wtag("t"))).strip()
    return int(text) if text.isdigit() else None


def expand_table(tbl, source_num=50, start_num=51, end_num=130):
    rows = table_rows(tbl)
    template_row = None
    for tr in rows:
        max_n, _ = row_max_subject_num(tr)
        if max_n == source_num and first_cell_number(tr) == source_num:
            template_row = tr

    if template_row is None:
        return 0

    # Must use index in tbl direct children (tblPr/tblGrid shift row indices).
    insert_pos = list(tbl).index(template_row)
    added = 0

    for n in range(start_num, end_num + 1):
        new_row = deepcopy(template_row)
        replace_subject_number_in_element(new_row, source_num, n)
        replace_row_number_in_cells(new_row, source_num, n)
        insert_pos += 1
        tbl.insert(insert_pos, new_row)
        added += 1

    return added


def verify_row_order(tbl, end_num=130):
    """Ensure data rows are sequential 1..end_num."""
    rows = table_rows(tbl)
    nums = []
    for tr in rows:
        max_n, _ = row_max_subject_num(tr)
        if max_n > 0:
            nums.append((first_cell_number(tr), max_n))
    issues = []
    prev = 0
    for cell, max_n in nums:
        if max_n != prev + 1:
            issues.append((prev, max_n, cell))
        prev = max_n
    return issues


def verify_fields(root, checks):
    all_names = []
    for elem in root:
        all_names.extend(parse_merge_fields_in_element(elem))
    results = {}
    for name in checks:
        results[name] = name in all_names
    nums = set()
    for name in all_names:
        m = re.match(r"科目(\d+)_", name)
        if m:
            nums.add(int(m.group(1)))
    results["_max_subject_num"] = max(nums) if nums else 0
    results["_count_狀態"] = sum(
        1 for n in all_names if re.match(r"科目\d+_狀態$", n)
    )
    return results


def main():
    docx_path = (
        Path(__file__).resolve().parent.parent
        / "Template"
        / "畢業預警通知單合併欄位總表.docx"
    )
    backup_path = docx_path.with_suffix(".docx.bak")

    if not backup_path.exists():
        shutil.copy2(docx_path, backup_path)
        print(f"Backup created: {backup_path}")

    # Always rebuild from original backup to avoid compounding errors.
    source_path = backup_path
    with zipfile.ZipFile(source_path, "r") as zin:
        document_xml = zin.read("word/document.xml")
        other_files = {
            name: zin.read(name)
            for name in zin.namelist()
            if name != "word/document.xml"
        }

    root = etree.fromstring(document_xml)
    tables = root.findall(f".//{wtag('tbl')}")

    expanded_tables = []
    for ti, tbl in enumerate(tables):
        rows = table_rows(tbl)
        max_n = 0
        for tr in rows:
            n, _ = row_max_subject_num(tr)
            max_n = max(max_n, n)
        if max_n == 50:
            added = expand_table(tbl, source_num=50, start_num=51, end_num=130)
            if added:
                issues = verify_row_order(tbl)
                if issues:
                    print(f"Table {ti}: ROW ORDER ERROR {issues[:5]}")
                    return 1
                expanded_tables.append((ti, added))
                print(f"Table {ti}: added {added} rows (51-130)")

    # Must use lxml (not xml.etree) so Word namespace prefixes (w:, mc:, etc.) stay intact.
    new_xml = etree.tostring(
        root, xml_declaration=True, encoding="UTF-8", standalone=True
    )
    if b"ns0:" in new_xml[:4096] or b"<ns0:document" in new_xml[:4096]:
        print("\nXML namespace corruption detected (ns0:). Aborting.")
        return 1
    if b"<w:document" not in new_xml[:4096]:
        print("\nMissing w:document root. Aborting.")
        return 1

    checks = [
        "科目51_狀態",
        "科目51_科目名稱",
        "科目51_學分數",
        "科目100_狀態",
        "科目100_科目名稱",
        "科目100_學分數",
        "科目130_狀態",
        "科目130_修課學年度",
        "科目130_修課學期",
        "科目130_科目名稱",
        "科目130_科目級別",
        "科目130_學分數",
        "科目130_應修總學分數_可補修重修_打勾",
        "科目130_應修所有必修課程_可補修重修_打勾",
        "科目130_應修所有部定必修課程_可補修重修_打勾",
        "科目130_應修專業及實習總學分數_可補修重修_打勾",
        "科目130_總學分數_可補修重修_打勾",
        "科目130_必修學分數_可補修重修_打勾",
        "科目130_部訂必修學分數_可補修重修_打勾",
        "科目130_校訂必修學分數_可補修重修_打勾",
        "科目130_選修學分數_可補修重修_打勾",
        "科目130_專業及實習總學分數_可補修重修_打勾",
        "科目130_實習學分數_可補修重修_打勾",
        "科目130_修課學分數統計_核心科目表序號1_規則_可補修重修_打勾",
        "科目130_修課學分數統計_核心科目表序號5_規則_可補修重修_打勾",
        "科目130_取得學分數統計_核心科目表序號1_規則_可補修重修_打勾",
        "科目130_取得學分數統計_核心科目表序號5_規則_可補修重修_打勾",
    ]
    verify_root = etree.fromstring(new_xml)
    results = verify_fields(verify_root, checks)

    print("\nVerification:")
    all_ok = True
    for k, v in results.items():
        if k.startswith("_"):
            print(f"  {k}: {v}")
            continue
        status = "OK" if v else "MISSING"
        if not v:
            all_ok = False
        print(f"  {k}: {status}")

    if not all_ok or results["_max_subject_num"] < 130:
        print("\nVerification FAILED - not writing docx")
        return 1

    temp_path = docx_path.with_suffix(".docx.tmp")
    with zipfile.ZipFile(source_path, "r") as zin:
        compress_map = {info.filename: info.compress_type for info in zin.infolist()}
    with zipfile.ZipFile(temp_path, "w", zipfile.ZIP_DEFLATED) as zout:
        zout.writestr(
            "word/document.xml",
            new_xml,
            compress_type=compress_map.get("word/document.xml", zipfile.ZIP_DEFLATED),
        )
        for name, data in other_files.items():
            zout.writestr(
                name, data, compress_type=compress_map.get(name, zipfile.ZIP_DEFLATED)
            )

    temp_path.replace(docx_path)
    print(f"\nSaved: {docx_path}")
    print(f"Expanded tables: {expanded_tables}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
