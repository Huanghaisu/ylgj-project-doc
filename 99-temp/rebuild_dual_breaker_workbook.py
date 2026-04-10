from pathlib import Path
import zipfile
import xml.etree.ElementTree as ET


ROOT = Path(r"E:\省平台\运力管家-黄骅港项目")
EXTRACTED = ROOT / "99-temp" / "财务分析提取"
TARGET = ROOT / "04-财务模型" / "20-工作稿 黄骅港项目 - 财务分析.xlsx"
NEW_SHEET_NAME = "7-双层熔断测算区"

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL_DOC = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_REL_PKG = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"
NS_APP = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
NS_VT = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"

ET.register_namespace("", NS_MAIN)
ET.register_namespace("r", NS_REL_DOC)
ET.register_namespace("", NS_REL_PKG)
ET.register_namespace("", NS_CT)
ET.register_namespace("", NS_APP)
ET.register_namespace("vt", NS_VT)


def qn(ns, tag):
    return f"{{{ns}}}{tag}"


def load_xml(path):
    return ET.parse(path).getroot()


def read_sheet_cell(sheet_file, ref):
    root = load_xml(EXTRACTED / "xl" / "worksheets" / sheet_file)
    for c in root.findall(f".//{qn(NS_MAIN,'c')}"):
        if c.attrib.get("r") == ref:
            v = c.find(qn(NS_MAIN, "v"))
            if v is None or v.text is None:
                return 0.0
            return float(v.text)
    return 0.0


def make_inline(cell_ref, text, style):
    c = ET.Element(qn(NS_MAIN, "c"), {"r": cell_ref, "s": str(style), "t": "inlineStr"})
    is_el = ET.SubElement(c, qn(NS_MAIN, "is"))
    t = ET.SubElement(is_el, qn(NS_MAIN, "t"))
    t.text = text
    return c


def make_formula(cell_ref, formula, value, style=318, text=False):
    attrs = {"r": cell_ref, "s": str(style)}
    if text:
        attrs["t"] = "str"
    c = ET.Element(qn(NS_MAIN, "c"), attrs)
    f = ET.SubElement(c, qn(NS_MAIN, "f"))
    f.text = formula
    v = ET.SubElement(c, qn(NS_MAIN, "v"))
    v.text = str(value)
    return c


def blank(cell_ref, style=318):
    return ET.Element(qn(NS_MAIN, "c"), {"r": cell_ref, "s": str(style)})


def add_row(sheet_data, row_num, cells, ht=None):
    attrs = {"r": str(row_num)}
    if ht is not None:
        attrs["ht"] = str(ht)
        attrs["customHeight"] = "1"
    row = ET.SubElement(sheet_data, qn(NS_MAIN, "row"), attrs)
    for c in cells:
        row.append(c)


def compute_metrics():
    metrics = {}
    metrics["price"] = read_sheet_cell("sheet2.xml", "C6")
    metrics["fleet"] = read_sheet_cell("sheet2.xml", "C13")
    metrics["load"] = read_sheet_cell("sheet2.xml", "C9")
    metrics["monthly_trips"] = read_sheet_cell("sheet2.xml", "C12")
    metrics["gross_income"] = read_sheet_cell("sheet4.xml", "F6")
    metrics["vehicle_cost"] = (
        read_sheet_cell("sheet4.xml", "F9")
        + read_sheet_cell("sheet4.xml", "F10")
        + read_sheet_cell("sheet4.xml", "F11")
    )
    metrics["energy_cost"] = read_sheet_cell("sheet4.xml", "F12")
    metrics["driver_cost"] = read_sheet_cell("sheet4.xml", "F13")
    metrics["maint_cost"] = read_sheet_cell("sheet4.xml", "F14")
    metrics["manage_cost"] = read_sheet_cell("sheet4.xml", "F15")
    metrics["other_cost"] = read_sheet_cell("sheet4.xml", "F16")
    metrics["tax_cost"] = read_sheet_cell("sheet4.xml", "F17")
    metrics["daily_rate"] = read_sheet_cell("sheet2.xml", "C25")
    metrics["freight_days"] = read_sheet_cell("sheet2.xml", "C22")
    metrics["carrier_per_vehicle"] = read_sheet_cell("sheet5.xml", "D28")

    metrics["tons"] = metrics["fleet"] * metrics["load"] * metrics["monthly_trips"]
    metrics["net_income"] = metrics["gross_income"]
    metrics["actual_price"] = metrics["net_income"] / metrics["tons"] if metrics["tons"] else 0
    metrics["freight_funding_cost"] = metrics["net_income"] * metrics["daily_rate"] * metrics["freight_days"]
    metrics["rent_funding_cost"] = 0
    metrics["non_carrier_total"] = (
        metrics["vehicle_cost"]
        + metrics["energy_cost"]
        + metrics["driver_cost"]
        + metrics["maint_cost"]
        + metrics["manage_cost"]
        + metrics["other_cost"]
        + metrics["tax_cost"]
        + metrics["freight_funding_cost"]
        + metrics["rent_funding_cost"]
    )
    metrics["carrier_total"] = metrics["carrier_per_vehicle"] * metrics["fleet"]
    metrics["carrier_line"] = metrics["carrier_total"] / metrics["tons"] if metrics["tons"] else 0
    metrics["project_total"] = metrics["non_carrier_total"] + metrics["carrier_total"]
    metrics["project_line"] = metrics["project_total"] / metrics["tons"] if metrics["tons"] else 0
    metrics["line_check"] = "通过" if metrics["carrier_line"] > metrics["project_line"] else "口径待复核"
    if metrics["line_check"] != "通过":
        metrics["status"] = "口径待复核"
        metrics["action"] = "先复核承运分包总价与非承运成本口径"
    elif metrics["actual_price"] > metrics["carrier_line"]:
        metrics["status"] = "正常执行区"
        metrics["action"] = "正常执行"
    elif metrics["actual_price"] > metrics["project_line"]:
        metrics["status"] = "承运层自担波动区"
        metrics["action"] = "承运层自担波动，项目继续执行"
    else:
        metrics["status"] = "正式熔断区"
        metrics["action"] = "停新增、保在途；启动停运重谈"
    metrics["resume"] = "暂不恢复新增投放"
    return metrics


def build_sheet10(metrics):
    root = ET.Element(qn(NS_MAIN, "worksheet"))
    sheet_pr = ET.SubElement(root, qn(NS_MAIN, "sheetPr"))
    ET.SubElement(sheet_pr, qn(NS_MAIN, "pageSetUpPr"), {"fitToPage": "1"})
    ET.SubElement(root, qn(NS_MAIN, "dimension"), {"ref": "B2:G40"})
    views = ET.SubElement(root, qn(NS_MAIN, "sheetViews"))
    view = ET.SubElement(views, qn(NS_MAIN, "sheetView"), {"showGridLines": "0", "workbookViewId": "0", "zoomScale": "90", "zoomScaleNormal": "90"})
    ET.SubElement(view, qn(NS_MAIN, "pane"), {"ySplit": "5", "topLeftCell": "B6", "activePane": "bottomLeft", "state": "frozen"})
    ET.SubElement(view, qn(NS_MAIN, "selection"), {"pane": "bottomLeft", "activeCell": "E16", "sqref": "E16"})
    ET.SubElement(root, qn(NS_MAIN, "sheetFormatPr"), {"defaultColWidth": "10", "defaultRowHeight": "15"})
    cols = ET.SubElement(root, qn(NS_MAIN, "cols"))
    for mi, ma, width in [(2, 2, 24), (3, 5, 18), (6, 6, 12), (7, 7, 34)]:
        ET.SubElement(cols, qn(NS_MAIN, "col"), {"min": str(mi), "max": str(ma), "width": str(width), "customWidth": "1"})
    sheet_data = ET.SubElement(root, qn(NS_MAIN, "sheetData"))

    add_row(sheet_data, 2, [make_inline("B2", NEW_SHEET_NAME, 360)] + [blank(f"{c}2", 360) for c in "CDEFG"], 20.25)
    add_row(sheet_data, 3, [make_inline("B3", "用于按月录入实际收入、分包结算和成本口径，自动形成双层熔断判断。", 361)] + [blank(f"{c}3", 361) for c in "CDEFG"])
    add_row(sheet_data, 5, [make_inline("B5", "项目", 313), make_inline("C5", "默认引用值", 313), make_inline("D5", "手工调整值", 313), make_inline("E5", "采用值", 313), make_inline("F5", "单位", 313), make_inline("G5", "说明", 313)])

    base_rows = [
        (6, "名义运价", "'1-基础参数'!C6", metrics["price"], "元/吨", "引用基础参数表。"),
        (7, "车队规模", "'1-基础参数'!C13", metrics["fleet"], "辆", "引用基础参数表。"),
        (8, "载重", "'1-基础参数'!C9", metrics["load"], "吨/车", "引用基础参数表。"),
        (9, "月趟数", "'1-基础参数'!C12", metrics["monthly_trips"], "趟/车/月", "引用基础参数表。"),
        (10, "理论月运量", "E7*E8*E9", metrics["tons"], "吨", "按规模、载重、趟数测算。"),
    ]
    for row_num, label, formula, value, unit, note in base_rows:
        add_row(sheet_data, row_num, [make_inline(f"B{row_num}", label, 314), make_formula(f"C{row_num}", formula, value), blank(f"D{row_num}"), make_formula(f"E{row_num}", f'IF(D{row_num}="",C{row_num},D{row_num})', value), make_inline(f"F{row_num}", unit, 316), make_inline(f"G{row_num}", note, 317)])

    add_row(sheet_data, 12, [make_inline("B12", "一、实际到手结算吨价", 358)] + [blank(f"{c}12", 358) for c in "CDEFG"], 15.75)
    items = [
        (13, "货主侧当期最终确认结算收入", "'3-收益测算'!F6", metrics["gross_income"], "元", "默认引用收益测算表收入。"),
        (14, "折让/扣罚/异常扣款", "0", 0, "元", "按月录入实际减项。"),
        (15, "当期净结算收入", "E13-E14", metrics["net_income"], "元", "形成正式比较口径分子。"),
        (16, "实际完成运输吨数", "E10", metrics["tons"], "吨", "可按实际完成吨数覆盖。"),
        (17, "实际到手结算吨价", "IFERROR(E15/E16,0)", metrics["actual_price"], "元/吨", "净结算收入÷实际完成运输吨数。"),
    ]
    for row_num, label, formula, value, unit, note in items:
        add_row(sheet_data, row_num, [make_inline(f"B{row_num}", label, 314), make_formula(f"C{row_num}", formula, value), blank(f"D{row_num}"), make_formula(f"E{row_num}", f'IF(D{row_num}="",C{row_num},D{row_num})', value), make_inline(f"F{row_num}", unit, 316), make_inline(f"G{row_num}", note, 317)])

    add_row(sheet_data, 19, [make_inline("B19", "二、非承运完全成本总额", 358)] + [blank(f"{c}19", 358) for c in "CDEFG"], 15.75)
    cost_rows = [
        (20, "车辆相关成本", "'3-收益测算'!F9+'3-收益测算'!F10+'3-收益测算'!F11", metrics["vehicle_cost"], "元", "车辆购置/租赁/保险合并口径。"),
        (21, "能源成本", "'3-收益测算'!F12", metrics["energy_cost"], "元", "引用收益测算表。"),
        (22, "司机人工成本", "'3-收益测算'!F13", metrics["driver_cost"], "元", "引用收益测算表。"),
        (23, "维保成本", "'3-收益测算'!F14", metrics["maint_cost"], "元", "引用收益测算表。"),
        (24, "管理成本", "'3-收益测算'!F15", metrics["manage_cost"], "元", "引用收益测算表。"),
        (25, "其他费用", "'3-收益测算'!F16", metrics["other_cost"], "元", "如有额外费用可覆盖。"),
        (26, "税费", "'3-收益测算'!F17", metrics["tax_cost"], "元", "引用收益测算表。"),
        (27, "运费账期资金成本", "E15*'1-基础参数'!C25*'1-基础参数'!C22", metrics["freight_funding_cost"], "元", "按净结算收入和账期测算。"),
        (28, "车租账期资金成本", "0", 0, "元", "资产车口径默认填0。"),
        (29, "非承运完全成本总额", "SUM(E20:E28)", metrics["non_carrier_total"], "元", "项目总熔断线第一部分。"),
    ]
    for row_num, label, formula, value, unit, note in cost_rows:
        add_row(sheet_data, row_num, [make_inline(f"B{row_num}", label, 314), make_formula(f"C{row_num}", formula, value), blank(f"D{row_num}"), make_formula(f"E{row_num}", f'IF(D{row_num}="",C{row_num},D{row_num})', value, 329 if row_num == 29 else 318), make_inline(f"F{row_num}", unit, 316), make_inline(f"G{row_num}", note, 317)])

    add_row(sheet_data, 31, [make_inline("B31", "三、双层熔断判断", 358)] + [blank(f"{c}31", 358) for c in "CDEFG"], 15.75)
    breaker_rows = [
        (32, "承运分包结算总价", "'3.5-资金清分'!D28*E7", metrics["carrier_total"], "元", "默认按资金清分表单车分包额×车队规模。"),
        (33, "承运调节线", "IFERROR(E32/E16,0)", metrics["carrier_line"], "元/吨", "承运分包结算总价÷实际完成运输吨数。"),
        (34, "项目总成本", "E29+E32", metrics["project_total"], "元", "非承运完全成本总额+承运分包结算总价。"),
        (35, "项目总熔断线", "IFERROR(E34/E16,0)", metrics["project_line"], "元/吨", "项目总成本÷实际完成运输吨数。"),
        (36, "线序校验", 'IF(E33>E35,"通过","口径待复核")', metrics["line_check"], "文本", "承运调节线应高于项目总熔断线。"),
        (37, "价格区间判定", 'IF(E36<>"通过","口径待复核",IF(E17>E33,"正常执行区",IF(E17>E35,"承运层自担波动区","正式熔断区")))', metrics["status"], "文本", "先判承运调节线，再判项目总熔断线。"),
        (38, "动作结论", 'IF(E37="口径待复核","先复核承运分包总价与非承运成本口径",IF(E37="正常执行区","正常执行",IF(E37="承运层自担波动区","承运层自担波动，项目继续执行","停新增、保在途；启动停运重谈")))', metrics["action"], "文本", "项目是否停做，仅以项目总熔断线为准。"),
        (39, "恢复执行建议", 'IF(E37="口径待复核","先复核后再判断","暂不恢复新增投放")', metrics["resume"], "文本", "满足价格、周期、回款三项条件后恢复。"),
    ]
    for row_num, label, formula, value, unit, note in breaker_rows:
        style = 315 if unit == "文本" else (329 if row_num in (33, 35) else 318)
        add_row(sheet_data, row_num, [make_inline(f"B{row_num}", label, 314), make_formula(f"C{row_num}", formula, value, style, text=(unit == "文本")), blank(f"D{row_num}", 317 if unit == "文本" else 318), make_formula(f"E{row_num}", f'IF(D{row_num}="",C{row_num},D{row_num})', value, style, text=(unit == "文本")), make_inline(f"F{row_num}", unit, 316), make_inline(f"G{row_num}", note, 317)])

    merge_cells = ET.SubElement(root, qn(NS_MAIN, "mergeCells"), {"count": "4"})
    for ref in ["B2:G2", "B3:G3", "B12:G12", "B19:G19", "B31:G31"]:
        ET.SubElement(merge_cells, qn(NS_MAIN, "mergeCell"), {"ref": ref})
    merge_cells.set("count", "5")

    ET.SubElement(root, qn(NS_MAIN, "pageMargins"), {"left": "0.75", "right": "0.75", "top": "1", "bottom": "1", "header": "0.5", "footer": "0.5"})
    ET.SubElement(root, qn(NS_MAIN, "pageSetup"), {"orientation": "landscape", "horizontalDpi": "300", "verticalDpi": "300"})
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def updated_workbook_xml():
    root = load_xml(EXTRACTED / "xl" / "workbook.xml")
    sheets = root.find(qn(NS_MAIN, "sheets"))
    ET.SubElement(sheets, qn(NS_MAIN, "sheet"), {"name": NEW_SHEET_NAME, "sheetId": "10", qn(NS_REL_DOC, "id"): "rId14"})
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def updated_workbook_rels():
    root = load_xml(EXTRACTED / "xl" / "_rels" / "workbook.xml.rels")
    for rel in list(root):
        if rel.attrib.get("Type", "").endswith("/calcChain"):
            root.remove(rel)
    ET.SubElement(root, qn(NS_REL_PKG, "Relationship"), {"Id": "rId14", "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet", "Target": "worksheets/sheet10.xml"})
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def updated_content_types():
    root = load_xml(EXTRACTED / "[Content_Types].xml")
    for node in list(root):
        if node.attrib.get("PartName") == "/xl/calcChain.xml":
            root.remove(node)
    ET.SubElement(root, qn(NS_CT, "Override"), {"PartName": "/xl/worksheets/sheet10.xml", "ContentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"})
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def updated_app_xml():
    root = load_xml(EXTRACTED / "docProps" / "app.xml")
    vector = root.find(f"{qn(NS_APP,'HeadingPairs')}/{qn(NS_VT,'vector')}")
    variants = vector.findall(qn(NS_VT, "variant"))
    variants[1].find(qn(NS_VT, "i4")).text = "10"
    titles_vector = root.find(f"{qn(NS_APP,'TitlesOfParts')}/{qn(NS_VT,'vector')}")
    titles_vector.set("size", str(int(titles_vector.attrib["size"]) + 1))
    ET.SubElement(titles_vector, qn(NS_VT, "lpstr")).text = NEW_SHEET_NAME
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def rebuild():
    metrics = compute_metrics()
    tmp = TARGET.with_suffix(".xlsx.tmp")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        for file in EXTRACTED.rglob("*"):
            if file.is_dir():
                continue
            rel = file.relative_to(EXTRACTED).as_posix()
            if rel == "xl/calcChain.xml":
                continue
            if rel == "xl/workbook.xml":
                z.writestr(rel, updated_workbook_xml())
            elif rel == "xl/_rels/workbook.xml.rels":
                z.writestr(rel, updated_workbook_rels())
            elif rel == "[Content_Types].xml":
                z.writestr(rel, updated_content_types())
            elif rel == "docProps/app.xml":
                z.writestr(rel, updated_app_xml())
            else:
                z.write(file, rel)
        z.writestr("xl/worksheets/sheet10.xml", build_sheet10(metrics))
    tmp.replace(TARGET)
    print("rebuilt workbook")


if __name__ == "__main__":
    rebuild()
