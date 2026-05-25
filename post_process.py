"""Post-process luanvan.docx:
- Track chapter number from each Heading 1 (parses "CHƯƠNG X")
- Prepend auto-numbering field to every Image Caption / Table Caption:
    "Hình <chapter>.<SEQ Hình \\s 1>: <original caption>"
    "Bảng <chapter>.<SEQ Bảng \\s 1>: <original caption>"
- SEQ field resets per chapter (\\s 1). Press F9 in Word to update.
"""
import re
import copy
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

CHAPTER_RE = re.compile(r"CH[ƯƯ]Ơ?NG\s+(\d+)", re.IGNORECASE)


def make_run(text, *, bold=False):
    r = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")
    rFonts = OxmlElement("w:rFonts")
    for a in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rFonts.set(qn(a), "Times New Roman")
    rPr.append(rFonts)
    sz = OxmlElement("w:sz"); sz.set(qn("w:val"), "26"); rPr.append(sz)
    szCs = OxmlElement("w:szCs"); szCs.set(qn("w:val"), "26"); rPr.append(szCs)
    color = OxmlElement("w:color"); color.set(qn("w:val"), "000000"); rPr.append(color)
    if bold:
        rPr.append(OxmlElement("w:b"))
        rPr.append(OxmlElement("w:bCs"))
    r.append(rPr)
    t = OxmlElement("w:t")
    t.set(qn("xml:space"), "preserve")
    t.text = text
    r.append(t)
    return r


def make_seq_field(seq_name):
    """Build runs implementing { SEQ <name> \\* ARABIC \\s 1 } field."""
    runs = []

    r1 = OxmlElement("w:r")
    fc1 = OxmlElement("w:fldChar"); fc1.set(qn("w:fldCharType"), "begin")
    r1.append(fc1)
    runs.append(r1)

    r2 = OxmlElement("w:r")
    it = OxmlElement("w:instrText")
    it.set(qn("xml:space"), "preserve")
    it.text = f" SEQ {seq_name} \\* ARABIC \\s 1 "
    r2.append(it)
    runs.append(r2)

    r3 = OxmlElement("w:r")
    fc2 = OxmlElement("w:fldChar"); fc2.set(qn("w:fldCharType"), "separate")
    r3.append(fc2)
    runs.append(r3)

    r4 = OxmlElement("w:r")
    t = OxmlElement("w:t"); t.text = "1"
    r4.append(t)
    runs.append(r4)

    r5 = OxmlElement("w:r")
    fc3 = OxmlElement("w:fldChar"); fc3.set(qn("w:fldCharType"), "end")
    r5.append(fc3)
    runs.append(r5)

    return runs


def style_of(p):
    pPr = p.find(qn("w:pPr"))
    if pPr is None:
        return None
    pStyle = pPr.find(qn("w:pStyle"))
    if pStyle is None:
        return None
    return pStyle.get(qn("w:val"))


def prepend_runs(p, runs):
    pPr = p.find(qn("w:pPr"))
    insert_after = pPr if pPr is not None else None
    if insert_after is None:
        for i, r in enumerate(runs):
            p.insert(i, r)
    else:
        idx = list(p).index(pPr) + 1
        for i, r in enumerate(runs):
            p.insert(idx + i, r)


def remove_existing_prefix(p, prefix_re):
    """If the caption text already starts with 'Hình X' or 'Bảng X', strip it."""
    texts = p.findall(f".//{{{W}}}t")
    if not texts:
        return
    full = "".join(t.text or "" for t in texts)
    m = prefix_re.match(full)
    if not m:
        return
    remaining = full[m.end():]
    # Clear all existing <w:t> by emptying; then put remainder in the first
    for t in texts:
        t.text = ""
        t.set(qn("xml:space"), "preserve")
    texts[0].text = remaining


PREFIX_RE = re.compile(r"^\s*(?:Hình|Bảng)\s+\d+(?:\.\d+)?\s*[:：.]?\s*", re.IGNORECASE)


def process(path):
    doc = Document(path)
    body = doc.element.body
    chapter = 0

    fig_count_per_chapter = {}
    tbl_count_per_chapter = {}

    # Walk all paragraphs in document order (including those inside tables shouldn't matter)
    for p in body.iter(qn("w:p")):
        st = style_of(p)
        if st == "Heading1":
            # extract chapter num from text
            text = "".join(t.text or "" for t in p.iter(qn("w:t")))
            m = CHAPTER_RE.search(text)
            if m:
                chapter = int(m.group(1))
            else:
                chapter += 1
        elif st in ("ImageCaption", "TableCaption"):
            label = "Hình" if st == "ImageCaption" else "Bảng"
            seq_name = label
            ch = chapter if chapter > 0 else 1

            # Strip any existing "Hình X.Y:" / "Bảng X.Y:" prefix that pandoc may have written
            remove_existing_prefix(p, PREFIX_RE)

            new_runs = [
                make_run(f"{label} {ch}.", bold=True),
                *[r for r in make_seq_field(seq_name)],
                make_run(": ", bold=True),
            ]
            # Force bold on SEQ result runs
            for r in new_runs:
                rPr = r.find(qn("w:rPr"))
                if rPr is not None and rPr.find(qn("w:b")) is None:
                    # check if this run has no font props (the field char runs)
                    pass

            prepend_runs(p, new_runs)

            # also center align caption
            pPr = p.find(qn("w:pPr"))
            if pPr is None:
                pPr = OxmlElement("w:pPr")
                p.insert(0, pPr)
            jc = pPr.find(qn("w:jc"))
            if jc is None:
                jc = OxmlElement("w:jc")
                pPr.append(jc)
            jc.set(qn("w:val"), "center")

    doc.save(path)
    print("post-processed")


if __name__ == "__main__":
    import sys
    path = sys.argv[1] if len(sys.argv) > 1 else "luanvan.docx"
    process(path)
