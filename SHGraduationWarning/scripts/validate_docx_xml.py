# -*- coding: utf-8 -*-
import re
import sys
import zipfile
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")

base = Path(__file__).resolve().parent.parent / "Template"


def analyze(path):
    print(f"\n=== {path.name} ===")
    with zipfile.ZipFile(path) as z:
        bad = z.testzip()
        print("zip test:", "OK" if bad is None else bad)
        for name in sorted(z.namelist()):
            if name.endswith(".xml") or name.endswith(".rels"):
                data = z.read(name)
                try:
                    data.decode("utf-8")
                except UnicodeDecodeError as e:
                    print(f"  BAD encoding: {name}: {e}")

        doc = z.read("word/document.xml").decode("utf-8")
        print("document.xml bytes:", len(doc.encode("utf-8")))
        print("declaration:", doc[:80].replace("\n", " "))

        issues = []
        if "ns0:" in doc or "ns1:" in doc:
            issues.append("ElementTree renamed namespaces (ns0:, ns1:)")
        if 'xmlns:w=""' in doc:
            issues.append("empty xmlns:w")
        if "<w:document" not in doc:
            issues.append("missing w:document root")
        if doc.count("<w:document") != 1:
            issues.append(f"w:document count={doc.count('<w:document')}")
        # ET often drops mc:Ignorable etc on body
        if "mc:Ignorable" not in doc and "mc:Ignorable" in z.read(
            "word/document.xml" if path.suffix == ".bak" else "word/document.xml"
        ):
            pass

        # Unescaped invalid chars
        if re.search(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", doc):
            issues.append("control characters in XML")

        # Mismatched tags rough check via lxml if available
        try:
            from lxml import etree

            etree.fromstring(doc.encode("utf-8"))
            print("lxml parse: OK")
        except ImportError:
            print("lxml: not installed")
        except Exception as e:
            issues.append(f"lxml parse error: {e}")

        if issues:
            print("ISSUES:")
            for i in issues:
                print(" ", i)
        else:
            print("No obvious XML issues")

        # sample root attrs
        m = re.search(r"<w:document[^>]{0,800}>", doc)
        if m:
            print("root:", m.group()[:400])


for fname in ["畢業預警通知單合併欄位總表.docx.bak", "畢業預警通知單合併欄位總表.docx"]:
    p = base / fname
    if p.exists():
        analyze(p)
