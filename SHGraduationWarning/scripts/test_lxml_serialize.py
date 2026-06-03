# -*- coding: utf-8 -*-
import sys
import zipfile
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")

try:
    from lxml import etree
except ImportError:
    print("no lxml")
    raise SystemExit(1)

bak = Path(__file__).resolve().parent.parent / "Template" / "畢業預警通知單合併欄位總表.docx.bak"
with zipfile.ZipFile(bak) as z:
    xml_bytes = z.read("word/document.xml")

root = etree.fromstring(xml_bytes)
out = etree.tostring(
    root, xml_declaration=True, encoding="UTF-8", standalone=True
).decode("utf-8")
print("head:", out[:200])
print("has ns0:", "ns0:" in out[:500])
print("has w:document:", "<w:document" in out[:500])
