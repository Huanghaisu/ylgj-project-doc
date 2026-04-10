import copy
import os
import zipfile
import xml.etree.ElementTree as ET


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


WORKBOOK = r"E:\省平台\运力管家-黄骅港项目\04-财务模型\20-工作稿 黄骅港项目 - 财务分析.xlsx"
NEW_SHEET_NAME = "7-双层熔断测算区"
NEW_SHEET_FILE = "xl/worksheets/sheet10.xml"


def qn(ns, tag):
    return f"{{{ns}}}{tag}"


def read_cell_value(zf, sheet_path, cell_ref):
    root = ET.fromstring(zf.read(sheet_path))
    for cell in root.findall(f".//{qn(NS_MAIN,'c')}"):
        if cell.attrib.get("r") != cell_ref:
            continue
        value = cell.find(qn(NS_MAIN, "v"))
        if value is None or value.text is None or value.text == "":
            return 0.0
        try:
            return float(value.text)
        except ValueError:
            return value.text
    raise KeyError(f"{sheet_path}:{cell_ref} not found")


def make_formula_cell(ref, formula, value, style=318, cell_type=None):
    cell = ET.Element(qn(NS_MAIN, "c"), {"r": ref, "s": str(style)})
    if cell_type:
        cell.set("t", cell_type)
    f = ET.SubElement(cell, qn(NS_MAIN, "f"))
    f.text = formula
    v = ET.SubElement(cell, qn(NS_MAIN, "v"))
    v.text = str(value)
    return cell


def make_inline_cell(ref, text, style=314):
    cell = ET.Element(qn(NS_MAIN, "c"), {"r": ref, "s": str(style), "t": "inlineStr"})
    is_el = ET.SubElement(cell, qn(NS_MAIN, "is"))
    t = ET.SubElement(is_el, qn(NS_MAIN, "t"))
    t.text = text
    return cell


def make_blank_cell(ref, style=318):
    return ET.Element(qn(NS_MAIN, "c"), {"r": ref, "s": str(style)})


def add_row(sheet_data, row_num, cells, height=None):
    attrs = {"r": str(row_num)}
    if height is not None:
        attrs["ht"] = str(height)
        attrs["customHeight"] = "1"
    row = ET.SubElement(sheet_data, qn(NS_MAIN, "row"), attrs)
    for cell in cells:
        row.append(cell)


def build_sheet_xml(metrics):
    root = ET.Element(
        qn(NS_MAIN, "worksheet"),
        {
            qn(NS_REL_DOC, "uid"): "{00000000-0001-0000-0900-000000000000}",
        },
    )

    sheet_pr = ET.SubElement(root, qn(NS_MAIN, "sheetPr"))
    ET.SubElement(sheet_pr, qn(NS_MAIN, "pageSetUpPr"), {"fitToPage": "1"})
    ET.SubElement(root, qn(NS_MAIN, "dimension"), {"ref": "B2:G48"})

    sheet_views = ET.SubElement(root, qn(NS_MAIN, "sheetViews"))
    sheet_view = ET.SubElement(
        sheet_views,
        qn(NS_MAIN, "sheetView"),
        {"showGridLines": "0", "workbookViewId": "0", "zoomScale": "90", "zoomScaleNormal": "90"},
    )
    ET.SubElement(
        sheet_view,
        qn(NS_MAIN, "pane"),
        {"ySplit": "5", "topLeftCell": "B6", "activePane": "bottomLeft", "state": "frozen"},
    )
    ET.SubElement(sheet_view, qn(NS_MAIN, "selection"), {"pane": "bottomLeft", "activeCell": "E15", "sqref": "E15"})

    ET.SubElement(root, qn(NS_MAIN, "sheetFormatPr"), {"defaultColWidth": "10", "defaultRowHeight": "15"})
    cols = ET.SubElement(root, qn(NS_MAIN, "cols"))
    for min_idx, max_idx, width in [(2, 2, 24), (3, 5, 18), (6, 6, 12), (7, 7, 34)]:
        ET.SubElement(cols, qn(NS_MAIN, "col"), {"min": str(min_idx), "max": str(max_idx), "width": str(width), "customWidth": "1"})

    sheet_data = ET.SubElement(root, qn(NS_MAIN, "sheetData"))

    add_row(sheet_data, 2, [make_inline_cell("B2", NEW_SHEET_NAME, 360)] + [make_blank_cell(f"{col}2", 360) for col in "CDEFG"], 20.25)
    add_row(sheet_data, 3, [make_inline_cell("B3", "用于将“承运调节线+项目总熔断线”直接落到财务模型，默认引用现有算表，可按月人工修正。", 361)] + [make_blank_cell(f"{col}3", 361) for col in "CDEFG"])

    add_row(
        sheet_data,
        5,
        [
            make_inline_cell("B5", "项目", 313),
            make_inline_cell("C5", "默认引用值", 313),
            make_inline_cell("D5", "手工调整值", 313),
            make_inline_cell("E5", "采用值", 313),
            make_inline_cell("F5", "单位", 313),
            make_inline_cell("G5", "说明", 313),
        ],
    )

    add_row(sheet_data, 6, [make_inline_cell("B6", "名义运价", 314), make_formula_cell("C6", "'1-基础参数'!C6", metrics["price"], 318), make_blank_cell("D6", 318), make_formula_cell("E6", 'IF(D6="",C6,D6)', metrics["price"], 318), make_inline_cell("F6", "元/吨", 316), make_inline_cell("G6", "引用基础参数表当前运价。", 317)])
    add_row(sheet_data, 7, [make_inline_cell("B7", "车队规模", 314), make_formula_cell("C7", "'1-基础参数'!C13", metrics["fleet"], 318), make_blank_cell("D7", 318), make_formula_cell("E7", 'IF(D7="",C7,D7)', metrics["fleet"], 318), make_inline_cell("F7", "辆", 316), make_inline_cell("G7", "引用基础参数表当前车队规模。", 317)])
    add_row(sheet_data, 8, [make_inline_cell("B8", "载重", 314), make_formula_cell("C8", "'1-基础参数'!C9", metrics["load"], 318), make_blank_cell("D8", 318), make_formula_cell("E8", 'IF(D8="",C8,D8)', metrics["load"], 318), make_inline_cell("F8", "吨/车", 316), make_inline_cell("G8", "引用基础参数表单车额定载重。", 317)])
    add_row(sheet_data, 9, [make_inline_cell("B9", "月趟数", 314), make_formula_cell("C9", "'1-基础参数'!C12", metrics["monthly_trips"], 318), make_blank_cell("D9", 318), make_formula_cell("E9", 'IF(D9="",C9,D9)', metrics["monthly_trips"], 318), make_inline_cell("F9", "趟/车/月", 316), make_inline_cell("G9", "引用基础参数表自动计算值。", 317)])
    add_row(sheet_data, 10, [make_inline_cell("B10", "理论月运量", 314), make_formula_cell("C10", "E7*E8*E9", metrics["tons"], 318), make_blank_cell("D10", 318), make_formula_cell("E10", 'IF(D10="",C10,D10)', metrics["tons"], 318), make_inline_cell("F10", "吨", 316), make_inline_cell("G10", "按车队规模、载重、月趟数测算。", 317)])

    add_row(sheet_data, 11, [make_inline_cell("B11", "一、实际到手结算吨价测算", 358)] + [make_blank_cell(f"{col}11", 358) for col in "CDEFG"], 15.75)
    add_row(sheet_data, 12, [make_inline_cell("B12", "货主侧当期最终确认结算收入", 314), make_formula_cell("C12", "'3-收益测算'!F6", metrics["gross_income"], 318), make_blank_cell("D12", 318), make_formula_cell("E12", 'IF(D12="",C12,D12)', metrics["gross_income"], 318), make_inline_cell("F12", "元", 316), make_inline_cell("G12", "默认引用购置车队月度运输收入，可按实际对账结果改写。", 317)])
    add_row(sheet_data, 13, [make_inline_cell("B13", "折让/扣罚/异常扣款", 314), make_formula_cell("C13", "0", 0, 318), make_blank_cell("D13", 318), make_formula_cell("E13", 'IF(D13="",C13,D13)', 0, 318), make_inline_cell("F13", "元", 316), make_inline_cell("G13", "录入货主侧已确认减项。", 317)])
    add_row(sheet_data, 14, [make_inline_cell("B14", "当期净结算收入", 314), make_formula_cell("C14", "E12-E13", metrics["net_income"], 318), make_blank_cell("D14", 318), make_formula_cell("E14", 'IF(D14="",C14,D14)', metrics["net_income"], 318), make_inline_cell("F14", "元", 316), make_inline_cell("G14", "作为正式比较口径的分子。", 317)])
    add_row(sheet_data, 15, [make_inline_cell("B15", "实际完成运输吨数", 314), make_formula_cell("C15", "E10", metrics["tons"], 318), make_blank_cell("D15", 318), make_formula_cell("E15", 'IF(D15="",C15,D15)', metrics["tons"], 318), make_inline_cell("F15", "吨", 316), make_inline_cell("G15", "可按当月实际完成吨数覆盖默认值。", 317)])
    add_row(sheet_data, 16, [make_inline_cell("B16", "实际到手结算吨价", 314), make_formula_cell("C16", "IFERROR(E14/E15,0)", metrics["actual_price"], 318), make_blank_cell("D16", 318), make_formula_cell("E16", 'IF(D16="",C16,D16)', metrics["actual_price"], 318), make_inline_cell("F16", "元/吨", 316), make_inline_cell("G16", "净结算收入÷实际完成运输吨数。", 317)])

    add_row(sheet_data, 17, [make_inline_cell("B17", "二、非承运完全成本总额", 358)] + [make_blank_cell(f"{col}17", 358) for col in "CDEFG"], 15.75)
    add_row(sheet_data, 18, [make_inline_cell("B18", "车辆相关成本", 314), make_formula_cell("C18", "'3-收益测算'!F9+'3-收益测算'!F10+'3-收益测算'!F11", metrics["vehicle_cost"], 318), make_blank_cell("D18", 318), make_formula_cell("E18", 'IF(D18="",C18,D18)', metrics["vehicle_cost"], 318), make_inline_cell("F18", "元", 316), make_inline_cell("G18", "购置/租赁/保险合并口径，默认引用购置车队。", 317)])
    add_row(sheet_data, 19, [make_inline_cell("B19", "能源成本", 314), make_formula_cell("C19", "'3-收益测算'!F12", metrics["energy_cost"], 318), make_blank_cell("D19", 318), make_formula_cell("E19", 'IF(D19="",C19,D19)', metrics["energy_cost"], 318), make_inline_cell("F19", "元", 316), make_inline_cell("G19", "引用收益测算表月度能源费用。", 317)])
    add_row(sheet_data, 20, [make_inline_cell("B20", "司机人工成本", 314), make_formula_cell("C20", "'3-收益测算'!F13", metrics["driver_cost"], 318), make_blank_cell("D20", 318), make_formula_cell("E20", 'IF(D20="",C20,D20)', metrics["driver_cost"], 318), make_inline_cell("F20", "元", 316), make_inline_cell("G20", "引用收益测算表月度司机成本。", 317)])
    add_row(sheet_data, 21, [make_inline_cell("B21", "维保成本", 314), make_formula_cell("C21", "'3-收益测算'!F14", metrics["maint_cost"], 318), make_blank_cell("D21", 318), make_formula_cell("E21", 'IF(D21="",C21,D21)', metrics["maint_cost"], 318), make_inline_cell("F21", "元", 316), make_inline_cell("G21", "引用收益测算表月度维保费用。", 317)])
    add_row(sheet_data, 22, [make_inline_cell("B22", "管理成本", 314), make_formula_cell("C22", "'3-收益测算'!F15", metrics["manage_cost"], 318), make_blank_cell("D22", 318), make_formula_cell("E22", 'IF(D22="",C22,D22)', metrics["manage_cost"], 318), make_inline_cell("F22", "元", 316), make_inline_cell("G22", "引用收益测算表月度管理费用。", 317)])
    add_row(sheet_data, 23, [make_inline_cell("B23", "其他费用", 314), make_formula_cell("C23", "'3-收益测算'!F16", metrics["other_cost"], 318), make_blank_cell("D23", 318), make_formula_cell("E23", 'IF(D23="",C23,D23)', metrics["other_cost"], 318), make_inline_cell("F23", "元", 316), make_inline_cell("G23", "如有进场费等，可在手工调整列覆盖。", 317)])
    add_row(sheet_data, 24, [make_inline_cell("B24", "税费", 314), make_formula_cell("C24", "'3-收益测算'!F17", metrics["tax_cost"], 318), make_blank_cell("D24", 318), make_formula_cell("E24", 'IF(D24="",C24,D24)', metrics["tax_cost"], 318), make_inline_cell("F24", "元", 316), make_inline_cell("G24", "引用收益测算表税费。", 317)])
    add_row(sheet_data, 25, [make_inline_cell("B25", "运费账期资金成本", 314), make_formula_cell("C25", "E14*'1-基础参数'!C25*'1-基础参数'!C22", metrics["freight_funding_cost"], 318), make_blank_cell("D25", 318), make_formula_cell("E25", 'IF(D25="",C25,D25)', metrics["freight_funding_cost"], 318), make_inline_cell("F25", "元", 316), make_inline_cell("G25", "按当期净结算收入、日利率和运费账期测算。", 317)])
    add_row(sheet_data, 26, [make_inline_cell("B26", "车租账期资金成本", 314), make_formula_cell("C26", "0", metrics["rent_funding_cost"], 318), make_blank_cell("D26", 318), make_formula_cell("E26", 'IF(D26="",C26,D26)', metrics["rent_funding_cost"], 318), make_inline_cell("F26", "元", 316), make_inline_cell("G26", "资产车口径默认填0；如存在车租链条，可手工补录。", 317)])
    add_row(sheet_data, 27, [make_inline_cell("B27", "非承运完全成本总额", 314), make_blank_cell("C27", 318), make_blank_cell("D27", 318), make_formula_cell("E27", "SUM(E18:E26)", metrics["non_carrier_total"], 329), make_inline_cell("F27", "元", 316), make_inline_cell("G27", "作为项目总熔断线的第一部分。", 317)])

    add_row(sheet_data, 28, [make_inline_cell("B28", "三、承运调节线与项目总熔断线", 358)] + [make_blank_cell(f"{col}28", 358) for col in "CDEFG"], 15.75)
    add_row(sheet_data, 29, [make_inline_cell("B29", "承运分包结算总价", 314), make_formula_cell("C29", "'3.5-资金清分'!D28*E7", metrics["carrier_total"], 318), make_blank_cell("D29", 318), make_formula_cell("E29", 'IF(D29="",C29,D29)', metrics["carrier_total"], 318), make_inline_cell("F29", "元", 316), make_inline_cell("G29", "默认按资金清分表承运分包款单车值×车队规模测算。", 317)])
    add_row(sheet_data, 30, [make_inline_cell("B30", "承运调节线", 314), make_blank_cell("C30", 318), make_blank_cell("D30", 318), make_formula_cell("E30", "IFERROR(E29/E15,0)", metrics["carrier_line"], 329), make_inline_cell("F30", "元/吨", 316), make_inline_cell("G30", "承运分包结算总价÷实际完成运输吨数。", 317)])
    add_row(sheet_data, 31, [make_inline_cell("B31", "项目总成本", 314), make_blank_cell("C31", 318), make_blank_cell("D31", 318), make_formula_cell("E31", "E27+E29", metrics["project_total"], 329), make_inline_cell("F31", "元", 316), make_inline_cell("G31", "非承运完全成本总额+承运分包结算总价。", 317)])
    add_row(sheet_data, 32, [make_inline_cell("B32", "项目总熔断线", 314), make_blank_cell("C32", 318), make_blank_cell("D32", 318), make_formula_cell("E32", "IFERROR(E31/E15,0)", metrics["project_line"], 329), make_inline_cell("F32", "元/吨", 316), make_inline_cell("G32", "项目总成本÷实际完成运输吨数。", 317)])
    add_row(sheet_data, 33, [make_inline_cell("B33", "线序校验", 314), make_blank_cell("C33", 317), make_blank_cell("D33", 317), make_formula_cell("E33", 'IF(E30>E32,"通过","口径待复核")', metrics["line_check"], 315, "str"), make_inline_cell("F33", "文本", 316), make_inline_cell("G33", "承运调节线应高于项目总熔断线；倒挂时先复核口径。", 317)])
    add_row(sheet_data, 34, [make_inline_cell("B34", "价格区间判定", 314), make_blank_cell("C34", 317), make_blank_cell("D34", 317), make_formula_cell("E34", 'IF(E33<>"通过","口径待复核",IF(E16>E30,"正常执行区",IF(E16>E32,"承运层自担波动区","正式熔断区")))', metrics["status"], 315, "str"), make_inline_cell("F34", "文本", 316), make_inline_cell("G34", "先判承运调节线，再判项目总熔断线。", 317)])
    add_row(sheet_data, 35, [make_inline_cell("B35", "动作结论", 314), make_blank_cell("C35", 317), make_blank_cell("D35", 317), make_formula_cell("E35", 'IF(E34="口径待复核","先复核承运分包总价与非承运成本口径",IF(E34="正常执行区","正常执行",IF(E34="承运层自担波动区","承运层自担波动，项目继续执行","停新增、保在途；启动停运重谈")))', metrics["action"], 315, "str"), make_inline_cell("F35", "文本", 316), make_inline_cell("G35", "项目是否停做，仅以项目总熔断线为准。", 317)])

    add_row(sheet_data, 37, [make_inline_cell("B37", "四、恢复执行复核", 358)] + [make_blank_cell(f"{col}37", 358) for col in "CDEFG"], 15.75)
    add_row(sheet_data, 38, [make_inline_cell("B38", "连续一个确认周期稳定", 314), make_inline_cell("C38", "否", 317), make_blank_cell("D38", 317), make_formula_cell("E38", 'IF(D38="",C38,D38)', "否", 315, "str"), make_inline_cell("F38", "是/否", 316), make_inline_cell("G38", "满足恢复执行的周期条件时填“是”。", 317)])
    add_row(sheet_data, 39, [make_inline_cell("B39", "回款周期未继续恶化", 314), make_inline_cell("C39", "否", 317), make_blank_cell("D39", 317), make_formula_cell("E39", 'IF(D39="",C39,D39)', "否", 315, "str"), make_inline_cell("F39", "是/否", 316), make_inline_cell("G39", "结合账期实际变化复核。", 317)])
    add_row(sheet_data, 40, [make_inline_cell("B40", "恢复执行建议", 314), make_blank_cell("C40", 317), make_blank_cell("D40", 317), make_formula_cell("E40", 'IF(AND(E34<>"口径待复核",E16>E32,E38="是",E39="是"),"可恢复新增投放","暂不恢复新增投放")', metrics["resume"], 315, "str"), make_inline_cell("F40", "文本", 316), make_inline_cell("G40", "同时满足价格、周期、回款三项条件后恢复。", 317)])

    add_row(sheet_data, 42, [make_inline_cell("B42", "五、测试场景校验", 358)] + [make_blank_cell(f"{col}42", 358) for col in "CDEFG"], 15.75)
    add_row(sheet_data, 43, [make_inline_cell("B43", "场景1：高于承运调节线", 314), make_formula_cell("C43", "E30+1", metrics["scenario_1_price"], 318), make_blank_cell("D43", 317), make_formula_cell("E43", 'IF(E33<>"通过","口径待复核",IF(C43>E30,"正常执行区",IF(C43>E32,"承运层自担波动区","正式熔断区")))', metrics["scenario_1_status"], 315, "str"), make_inline_cell("F43", "元/吨", 316), make_inline_cell("G43", "验证正常执行区。", 317)])
    add_row(sheet_data, 44, [make_inline_cell("B44", "场景2：跌破承运线但高于总线", 314), make_formula_cell("C44", "ROUND((E30+E32)/2,2)", metrics["scenario_2_price"], 318), make_blank_cell("D44", 317), make_formula_cell("E44", 'IF(E33<>"通过","口径待复核",IF(C44>E30,"正常执行区",IF(C44>E32,"承运层自担波动区","正式熔断区")))', metrics["scenario_2_status"], 315, "str"), make_inline_cell("F44", "元/吨", 316), make_inline_cell("G44", "验证承运方先行吸收波动。", 317)])
    add_row(sheet_data, 45, [make_inline_cell("B45", "场景3：击穿项目总线", 314), make_formula_cell("C45", "E32-1", metrics["scenario_3_price"], 318), make_blank_cell("D45", 317), make_formula_cell("E45", 'IF(E33<>"通过","口径待复核",IF(C45>E30,"正常执行区",IF(C45>E32,"承运层自担波动区","正式熔断区")))', metrics["scenario_3_status"], 315, "str"), make_inline_cell("F45", "元/吨", 316), make_inline_cell("G45", "验证项目级正式熔断。", 317)])
    add_row(sheet_data, 46, [make_inline_cell("B46", "场景4：当前测算价", 314), make_formula_cell("C46", "E16", metrics["actual_price"], 318), make_blank_cell("D46", 317), make_formula_cell("E46", 'IF(E33<>"通过","口径待复核",IF(C46>E30,"正常执行区",IF(C46>E32,"承运层自担波动区","正式熔断区")))', metrics["status"], 315, "str"), make_inline_cell("F46", "元/吨", 316), make_inline_cell("G46", "回看当前口径下的项目状态。", 317)])
    add_row(sheet_data, 47, [make_inline_cell("B47", "场景5：当前动作", 314), make_blank_cell("C47", 317), make_blank_cell("D47", 317), make_formula_cell("E47", 'IF(E46="口径待复核","先复核承运分包总价与非承运成本口径",IF(E46="正常执行区","正常执行",IF(E46="承运层自担波动区","承运层自担波动，项目继续执行","停新增、保在途；启动停运重谈")))', metrics["action"], 315, "str"), make_inline_cell("F47", "文本", 316), make_inline_cell("G47", "与场景4判定保持一致。", 317)])
    add_row(sheet_data, 48, [make_inline_cell("B48", "使用提示", 314), make_blank_cell("C48", 317), make_blank_cell("D48", 317), make_inline_cell("E48", "优先改动D列", 315), make_inline_cell("F48", "文本", 316), make_inline_cell("G48", "C列保留原模型引用，D列录入月度实际值，E列自动形成正式判断。", 317)])

    merge_cells = ET.SubElement(root, qn(NS_MAIN, "mergeCells"), {"count": "7"})
    for ref in ["B2:G2", "B3:G3", "B11:G11", "B17:G17", "B28:G28", "B37:G37", "B42:G42"]:
        ET.SubElement(merge_cells, qn(NS_MAIN, "mergeCell"), {"ref": ref})

    ET.SubElement(root, qn(NS_MAIN, "phoneticPr"), {"fontId": "43", "type": "noConversion"})
    ET.SubElement(root, qn(NS_MAIN, "pageMargins"), {"left": "0.75", "right": "0.75", "top": "1", "bottom": "1", "header": "0.5", "footer": "0.5"})
    ET.SubElement(root, qn(NS_MAIN, "pageSetup"), {"orientation": "landscape", "horizontalDpi": "300", "verticalDpi": "300", "fitToWidth": "1", "fitToHeight": "0"})

    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def update_workbook_xml(data):
    root = ET.fromstring(data)
    sheets = root.find(qn(NS_MAIN, "sheets"))
    max_sheet_id = max(int(sheet.attrib["sheetId"]) for sheet in sheets.findall(qn(NS_MAIN, "sheet")))
    new_sheet = ET.SubElement(
        sheets,
        qn(NS_MAIN, "sheet"),
        {"name": NEW_SHEET_NAME, "sheetId": str(max_sheet_id + 1), qn(NS_REL_DOC, "id"): "rId14"},
    )
    calc_pr = root.find(qn(NS_MAIN, "calcPr"))
    if calc_pr is None:
        calc_pr = ET.SubElement(root, qn(NS_MAIN, "calcPr"))
    calc_pr.set("fullCalcOnLoad", "1")
    calc_pr.set("forceFullCalc", "1")
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def update_workbook_rels(data):
    root = ET.fromstring(data)
    ET.SubElement(
        root,
        qn(NS_REL_PKG, "Relationship"),
        {
            "Id": "rId14",
            "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
            "Target": "worksheets/sheet10.xml",
        },
    )
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def update_content_types(data):
    root = ET.fromstring(data)
    ET.SubElement(
        root,
        qn(NS_CT, "Override"),
        {
            "PartName": "/xl/worksheets/sheet10.xml",
            "ContentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
        },
    )
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def update_app_xml(data):
    root = ET.fromstring(data)
    heading_pairs = root.find(qn(NS_APP, "HeadingPairs"))
    vector = heading_pairs.find(f"{qn(NS_VT,'vector')}")
    variants = vector.findall(qn(NS_VT, "variant"))
    sheet_count = variants[1].find(qn(NS_VT, "i4"))
    sheet_count.text = str(int(sheet_count.text) + 1)

    titles = root.find(qn(NS_APP, "TitlesOfParts"))
    titles_vector = titles.find(qn(NS_VT, "vector"))
    titles_vector.set("size", str(int(titles_vector.attrib["size"]) + 1))
    new_title = ET.SubElement(titles_vector, qn(NS_VT, "lpstr"))
    new_title.text = NEW_SHEET_NAME
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def main():
    tmp_path = WORKBOOK + ".tmp"
    with zipfile.ZipFile(WORKBOOK, "r") as src:
        metrics = {
            "price": read_cell_value(src, "xl/worksheets/sheet2.xml", "C6"),
            "fleet": read_cell_value(src, "xl/worksheets/sheet2.xml", "C13"),
            "load": read_cell_value(src, "xl/worksheets/sheet2.xml", "C9"),
            "monthly_trips": read_cell_value(src, "xl/worksheets/sheet2.xml", "C12"),
            "gross_income": read_cell_value(src, "xl/worksheets/sheet4.xml", "F6"),
            "vehicle_cost": read_cell_value(src, "xl/worksheets/sheet4.xml", "F9") + read_cell_value(src, "xl/worksheets/sheet4.xml", "F10") + read_cell_value(src, "xl/worksheets/sheet4.xml", "F11"),
            "energy_cost": read_cell_value(src, "xl/worksheets/sheet4.xml", "F12"),
            "driver_cost": read_cell_value(src, "xl/worksheets/sheet4.xml", "F13"),
            "maint_cost": read_cell_value(src, "xl/worksheets/sheet4.xml", "F14"),
            "manage_cost": read_cell_value(src, "xl/worksheets/sheet4.xml", "F15"),
            "other_cost": read_cell_value(src, "xl/worksheets/sheet4.xml", "F16"),
            "tax_cost": read_cell_value(src, "xl/worksheets/sheet4.xml", "F17"),
            "daily_rate": read_cell_value(src, "xl/worksheets/sheet2.xml", "C25"),
            "freight_days": read_cell_value(src, "xl/worksheets/sheet2.xml", "C22"),
            "carrier_per_vehicle": read_cell_value(src, "xl/worksheets/sheet5.xml", "D28"),
        }

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
        metrics["scenario_1_price"] = metrics["carrier_line"] + 1
        metrics["scenario_2_price"] = round((metrics["carrier_line"] + metrics["project_line"]) / 2, 2)
        metrics["scenario_3_price"] = metrics["project_line"] - 1
        metrics["scenario_1_status"] = "正常执行区" if metrics["line_check"] == "通过" else "口径待复核"
        metrics["scenario_2_status"] = "承运层自担波动区" if metrics["line_check"] == "通过" else "口径待复核"
        metrics["scenario_3_status"] = "正式熔断区" if metrics["line_check"] == "通过" else "口径待复核"

        with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as dst:
            for info in src.infolist():
                if info.filename == NEW_SHEET_FILE:
                    continue
                data = src.read(info.filename)
                if info.filename == "xl/workbook.xml":
                    data = update_workbook_xml(data)
                elif info.filename == "xl/_rels/workbook.xml.rels":
                    data = update_workbook_rels(data)
                elif info.filename == "[Content_Types].xml":
                    data = update_content_types(data)
                elif info.filename == "docProps/app.xml":
                    data = update_app_xml(data)
                dst.writestr(info, data)

            dst.writestr(NEW_SHEET_FILE, build_sheet_xml(metrics))

    os.replace(tmp_path, WORKBOOK)
    print("added", NEW_SHEET_NAME)


if __name__ == "__main__":
    main()
