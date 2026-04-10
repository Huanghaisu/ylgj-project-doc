from pathlib import Path
import re
import zipfile


ROOT = Path(r"E:\省平台\运力管家-黄骅港项目")
EXTRACTED = ROOT / "99-temp" / "财务分析提取"
TARGET = ROOT / "04-财务模型" / "20-工作稿 黄骅港项目 - 财务分析.xlsx"
NEW_SHEET_NAME = "7-双层熔断测算区"


def read_text(path: Path) -> str:
    return path.read_text(encoding="utf-8")


def build_sheet10() -> bytes:
    # Reuse the already-generated sheet10 from the current workbook if present; otherwise fall back to extracted temp if any.
    with zipfile.ZipFile(TARGET, "r") as zf:
        try:
            return zf.read("xl/worksheets/sheet10.xml")
        except KeyError:
            raise RuntimeError("sheet10.xml not found in current workbook")


def patch_workbook_xml(raw: str) -> str:
    insert = f'<sheet name="{NEW_SHEET_NAME}" sheetId="10" r:id="rId14"/>'
    if NEW_SHEET_NAME in raw:
        return raw
    raw = raw.replace("</sheets>", insert + "</sheets>")
    if "<calcPr " in raw:
        raw = raw.replace("<calcPr calcId=\"191029\"/>", "<calcPr calcId=\"191029\" fullCalcOnLoad=\"1\" forceFullCalc=\"1\"/>")
    return raw


def patch_workbook_rels(raw: str) -> str:
    raw = re.sub(r'<Relationship Id="rId13" Type="http://schemas\.openxmlformats\.org/officeDocument/2006/relationships/calcChain" Target="calcChain\.xml"/>', "", raw)
    if 'Target="worksheets/sheet10.xml"' not in raw:
        raw = raw.replace("</Relationships>", '<Relationship Id="rId14" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet10.xml"/></Relationships>')
    return raw


def patch_content_types(raw: str) -> str:
    raw = re.sub(r'<Override PartName="/xl/calcChain\.xml" ContentType="application/vnd\.openxmlformats-officedocument\.spreadsheetml\.calcChain\+xml"/>', "", raw)
    if '/xl/worksheets/sheet10.xml' not in raw:
        raw = raw.replace("</Types>", '<Override PartName="/xl/worksheets/sheet10.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>')
    return raw


def patch_app_xml(raw: str) -> str:
    raw = raw.replace("<vt:i4>9</vt:i4>", "<vt:i4>10</vt:i4>", 1)
    raw = raw.replace('size="30"', 'size="31"', 1)
    if f"<vt:lpstr>{NEW_SHEET_NAME}</vt:lpstr>" not in raw:
        raw = raw.replace("</vt:vector></TitlesOfParts>", f"<vt:lpstr>{NEW_SHEET_NAME}</vt:lpstr></vt:vector></TitlesOfParts>")
    return raw


def rebuild():
    workbook_xml = patch_workbook_xml(read_text(EXTRACTED / "xl" / "workbook.xml"))
    workbook_rels = patch_workbook_rels(read_text(EXTRACTED / "xl" / "_rels" / "workbook.xml.rels"))
    content_types = patch_content_types(read_text(EXTRACTED / "[Content_Types].xml"))
    app_xml = patch_app_xml(read_text(EXTRACTED / "docProps" / "app.xml"))
    sheet10 = build_sheet10()

    tmp = TARGET.with_suffix(".xlsx.tmp")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zf:
        for file in EXTRACTED.rglob("*"):
            if file.is_dir():
                continue
            rel = file.relative_to(EXTRACTED).as_posix()
            if rel == "xl/calcChain.xml":
                continue
            if rel == "xl/workbook.xml":
                zf.writestr(rel, workbook_xml.encode("utf-8"))
            elif rel == "xl/_rels/workbook.xml.rels":
                zf.writestr(rel, workbook_rels.encode("utf-8"))
            elif rel == "[Content_Types].xml":
                zf.writestr(rel, content_types.encode("utf-8"))
            elif rel == "docProps/app.xml":
                zf.writestr(rel, app_xml.encode("utf-8"))
            else:
                zf.write(file, rel)
        zf.writestr("xl/worksheets/sheet10.xml", sheet10)

    tmp.replace(TARGET)
    print("repaired workbook with preserved namespaces")


if __name__ == "__main__":
    rebuild()
