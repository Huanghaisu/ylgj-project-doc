from pathlib import Path
import zipfile
import xml.etree.ElementTree as ET


ROOT = Path.cwd()
SOURCE_BOOK = next(ROOT.glob("04-*/*.xlsx"))
BASE_COPY = ROOT / "99-temp" / "contract-logistics-breaker-fixed-dedup.xlsx"
OUTPUT = ROOT / "99-temp" / "contract-logistics-breaker-rental-v5.xlsx"
NS = {"a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def read_numeric(book: Path, sheet: str, cell: str) -> float:
    with zipfile.ZipFile(book) as zf:
        root = ET.fromstring(zf.read(f"xl/worksheets/{sheet}"))
    node = root.find(f".//a:c[@r='{cell}']/a:v", NS)
    return float(node.text) if node is not None and node.text else 0.0


def inline_str(text: str) -> str:
    return f'<is><t>{text}</t></is>'


def build_sheet10() -> bytes:
    tons = read_numeric(SOURCE_BOOK, "sheet2.xml", "C9") * read_numeric(SOURCE_BOOK, "sheet2.xml", "C12")
    default_revenue = read_numeric(SOURCE_BOOK, "sheet4.xml", "E6")
    default_price = default_revenue / tons if tons else 0.0

    driver = read_numeric(SOURCE_BOOK, "sheet4.xml", "E13")
    manage = read_numeric(SOURCE_BOOK, "sheet4.xml", "E15")
    freight_interest = read_numeric(SOURCE_BOOK, "sheet5.xml", "E12")
    rent_interest = read_numeric(SOURCE_BOOK, "sheet5.xml", "E13")
    vehicle_rent = read_numeric(SOURCE_BOOK, "sheet5.xml", "E15")
    insurance = read_numeric(SOURCE_BOOK, "sheet4.xml", "E11")
    energy = read_numeric(SOURCE_BOOK, "sheet4.xml", "E12")
    maint = read_numeric(SOURCE_BOOK, "sheet4.xml", "E14")
    tax = default_revenue * 0.09

    rigid_cover = freight_interest + rent_interest + vehicle_rent + insurance + energy + maint + tax
    carrier_margin = default_revenue - rigid_cover - driver - manage
    carrier_bear_line = rigid_cover + driver + manage
    rigid_line_price = rigid_cover / tons if tons else 0.0
    carrier_bear_line_price = carrier_bear_line / tons if tons else 0.0

    default_zone = (
        "正常执行区"
        if carrier_margin > 0
        else ("承运执行层自担区" if default_revenue >= rigid_cover else "合同物流企业执行边界区")
    )
    default_check = "通过" if default_revenue >= rigid_cover else "越界预警"
    default_action = (
        "继续执行"
        if default_zone == "正常执行区"
        else ("承运侧承担执行层缺口" if default_zone == "承运执行层自担区" else "暂停新增并启动价格重谈")
    )
    default_line_check = "通过" if rigid_line_price < default_price else "口径待复核"

    rows = [
        (2, [
            ('B2', 360, 'inlineStr', inline_str('7-运费调节测算区')),
            ('C2', 360, None, None), ('D2', 360, None, None), ('E2', 360, None, None), ('F2', 360, None, None), ('G2', 360, None, None)
        ]),
        (3, [
            ('B3', 361, 'inlineStr', inline_str('采用租赁模式单车口径。承运毛利改为随运价自动变化的结果变量，仅保留人工覆盖口。')),
            ('C3', 361, None, None), ('D3', 361, None, None), ('E3', 361, None, None), ('F3', 361, None, None), ('G3', 361, None, None)
        ]),
        (5, [
            ('B5', 313, 'inlineStr', inline_str('项目')),
            ('C5', 313, 'inlineStr', inline_str('默认引用值')),
            ('D5', 313, 'inlineStr', inline_str('手工调整值')),
            ('E5', 313, 'inlineStr', inline_str('采用值')),
            ('F5', 313, 'inlineStr', inline_str('单位')),
            ('G5', 313, 'inlineStr', inline_str('说明')),
        ]),
        (6, [
            ('B6', 314, 'inlineStr', inline_str('单车月运输吨数')),
            ('C6', 318, None, f"<f>'1-基础参数'!C9*'1-基础参数'!C12</f><v>{tons}</v>"),
            ('D6', 318, None, None),
            ('E6', 318, None, f'<f>IF(D6=&quot;&quot;,C6,D6)</f><v>{tons}</v>'),
            ('F6', 316, 'inlineStr', inline_str('吨/月')),
            ('G6', 317, 'inlineStr', inline_str('载重乘月趟数。')),
        ]),
        (7, [
            ('B7', 314, 'inlineStr', inline_str('货主支付物流费用')),
            ('C7', 318, None, f"<f>'3-收益测算'!E6</f><v>{default_revenue}</v>"),
            ('D7', 318, None, None),
            ('E7', 318, None, f'<f>IF(D7=&quot;&quot;,IF(D8=&quot;&quot;,C-,D8*E6),D-)</f><v>{default_revenue}</v>'),
            ('F7', 316, 'inlineStr', inline_str('元/月')),
            ('G7', 317, 'inlineStr', inline_str('默认引用收入；调整实际到手运价后自动反推。')),
        ]),
        (8, [
            ('B8', 314, 'inlineStr', inline_str('实际到手运价')),
            ('C8', 318, None, f'<f>IFERROR(C7/C6,0)</f><v>{default_price}</v>'),
            ('D8', 318, None, None),
            ('E8', 318, None, f'<f>IF(D8=&quot;&quot;,IFERROR(E7/E6,0),D8)</f><v>{default_price}</v>'),
            ('F8', 316, 'inlineStr', inline_str('元/吨趟')),
            ('G8', 317, 'inlineStr', inline_str('主调参入口。')),
        ]),
        (10, [
            ('B10', 358, 'inlineStr', inline_str('一、承运侧清分项目')),
            ('C10', 358, None, None), ('D10', 358, None, None), ('E10', 358, None, None), ('F10', 358, None, None), ('G10', 358, None, None)
        ]),
        (11, [
            ('B11', 314, 'inlineStr', inline_str('承运司机费用')),
            ('C11', 318, None, f"<f>'3-收益测算'!E13</f><v>{driver}</v>"),
            ('D11', 318, None, None),
            ('E11', 318, None, f'<f>IF(D11=&quot;&quot;,C11,D11)</f><v>{driver}</v>'),
            ('F11', 316, 'inlineStr', inline_str('元/月')),
            ('G11', 317, 'inlineStr', inline_str('承运执行层司机成本。')),
        ]),
        (12, [
            ('B12', 314, 'inlineStr', inline_str('承运管理费用')),
            ('C12', 318, None, f"<f>'3-收益测算'!E15</f><v>{manage}</v>"),
            ('D12', 318, None, None),
            ('E12', 318, None, f'<f>IF(D12=&quot;&quot;,C12,D12)</f><v>{manage}</v>'),
            ('F12', 316, 'inlineStr', inline_str('元/月')),
            ('G12', 317, 'inlineStr', inline_str('承运执行层管理成本。')),
        ]),
        (13, [
            ('B13', 314, 'inlineStr', inline_str('承运毛利')),
            ('C13', 329, None, f'<f>E7-E24-E11-E12</f><v>{carrier_margin}</v>'),
            ('D13', 318, None, None),
            ('E13', 329, None, f'<f>IF(D13=&quot;&quot;,E7-E24-E11-E12,D13)</f><v>{carrier_margin}</v>'),
            ('F13', 316, 'inlineStr', inline_str('元/月')),
            ('G13', 317, 'inlineStr', inline_str('默认随运价联动计算；如需谈判测算，可在D13覆盖。')),
        ]),
        (14, [
            ('B14', 314, 'inlineStr', inline_str('扣减承运侧后可用金额')),
            ('C14', 329, None, f'<f>E7-E11-E12-E13</f><v>{rigid_cover}</v>'),
            ('D14', 318, None, None),
            ('E14', 329, None, f'<f>IF(D14=&quot;&quot;,E7-E11-E12-E13,D14)</f><v>{rigid_cover}</v>'),
            ('F14', 316, 'inlineStr', inline_str('元/月')),
            ('G14', 317, 'inlineStr', inline_str('货主支付物流费用扣减承运司机、管理、毛利后的余额。')),
        ]),
        (16, [
            ('B16', 358, 'inlineStr', inline_str('二、合同物流企业刚性支出覆盖要求')),
            ('C16', 358, None, None), ('D16', 358, None, None), ('E16', 358, None, None), ('F16', 358, None, None), ('G16', 358, None, None)
        ]),
        (17, [
            ('B17', 314, 'inlineStr', inline_str('运费垫资利息')),
            ('C17', 318, None, f"<f>'3.5-资金清分'!E12</f><v>{freight_interest}</v>"),
            ('D17', 318, None, None),
            ('E17', 318, None, f'<f>IF(D17=&quot;&quot;,C17,D17)</f><v>{freight_interest}</v>'),
            ('F17', 316, 'inlineStr', inline_str('元/月')),
            ('G17', 317, 'inlineStr', inline_str('合同物流企业资金收益覆盖要求。')),
        ]),
        (18, [
            ('B18', 314, 'inlineStr', inline_str('车租垫资利息')),
            ('C18', 318, None, f"<f>'3.5-资金清分'!E13</f><v>{rent_interest}</v>"),
            ('D18', 318, None, None),
            ('E18', 318, None, f'<f>IF(D18=&quot;&quot;,C18,D18)</f><v>{rent_interest}</v>'),
            ('F18', 316, 'inlineStr', inline_str('元/月')),
            ('G18', 317, 'inlineStr', inline_str('租赁模式车租账期覆盖要求。')),
        ]),
        (19, [
            ('B19', 314, 'inlineStr', inline_str('车辆租金（牵引车+挂车）')),
            ('C19', 318, None, f"<f>'3.5-资金清分'!E15</f><v>{vehicle_rent}</v>"),
            ('D19', 318, None, None),
            ('E19', 318, None, f'<f>IF(D19=&quot;&quot;,C19,D19)</f><v>{vehicle_rent}</v>'),
            ('F19', 316, 'inlineStr', inline_str('元/月')),
            ('G19', 317, 'inlineStr', inline_str('租赁模式牵引车及挂车租金。')),
        ]),
        (20, [
            ('B20', 314, 'inlineStr', inline_str('保险费用')),
            ('C20', 318, None, f"<f>'3-收益测算'!E11</f><v>{insurance}</v>"),
            ('D20', 318, None, None),
            ('E20', 318, None, f'<f>IF(D20=&quot;&quot;,C20,D20)</f><v>{insurance}</v>'),
            ('F20', 316, 'inlineStr', inline_str('元/月')),
            ('G20', 317, 'inlineStr', inline_str('租赁模式默认值为0，保留覆盖口。')),
        ]),
        (21, [
            ('B21', 314, 'inlineStr', inline_str('能源费用')),
            ('C21', 318, None, f"<f>'3-收益测算'!E12</f><v>{energy}</v>"),
            ('D21', 318, None, None),
            ('E21', 318, None, f'<f>IF(D21=&quot;&quot;,C21,D21)</f><v>{energy}</v>'),
            ('F21', 316, 'inlineStr', inline_str('元/月')),
            ('G21', 317, 'inlineStr', inline_str('合同物流企业实际承担的能源支出。')),
        ]),
        (22, [
            ('B22', 314, 'inlineStr', inline_str('维保费用')),
            ('C22', 318, None, f"<f>'3-收益测算'!E14</f><v>{maint}</v>"),
            ('D22', 318, None, None),
            ('E22', 318, None, f'<f>IF(D22=&quot;&quot;,C22,D22)</f><v>{maint}</v>'),
            ('F22', 316, 'inlineStr', inline_str('元/月')),
            ('G22', 317, 'inlineStr', inline_str('合同物流企业实际承担的维保支出。')),
        ]),
        (23, [
            ('B23', 314, 'inlineStr', inline_str('货主支付物流费用税费（9%）')),
            ('C23', 329, None, f'<f>E7*9%</f><v>{tax}</v>'),
            ('D23', 318, None, None),
            ('E23', 329, None, f'<f>IF(D23=&quot;&quot;,E7*9%,D23)</f><v>{tax}</v>'),
            ('F23', 316, 'inlineStr', inline_str('元/月')),
            ('G23', 317, 'inlineStr', inline_str('按货主支付物流费用的9%测算，随收入自动联动。')),
        ]),
        (24, [
            ('B24', 314, 'inlineStr', inline_str('刚性支出覆盖要求合计')),
            ('C24', 329, None, f'<f>SUM(E17:E23)</f><v>{rigid_cover}</v>'),
            ('D24', 318, None, None),
            ('E24', 329, None, f'<f>IF(D24=&quot;&quot;,SUM(E17:E23),D24)</f><v>{rigid_cover}</v>'),
            ('F24', 316, 'inlineStr', inline_str('元/月')),
            ('G24', 317, 'inlineStr', inline_str('合同物流企业要求被覆盖的刚性支出与资金收益。')),
        ]),
        (25, [
            ('B25', 314, 'inlineStr', inline_str('合同物流企业执行边界线')),
            ('C25', 329, None, f'<f>IFERROR(E24/E6,0)</f><v>{rigid_line_price}</v>'),
            ('D25', 318, None, None),
            ('E25', 329, None, f'<f>IF(D25=&quot;&quot;,IFERROR(E24/E6,0),D25)</f><v>{rigid_line_price}</v>'),
            ('F25', 316, 'inlineStr', inline_str('元/吨趟')),
            ('G25', 317, 'inlineStr', inline_str('刚性支出覆盖要求合计除以单车月运输吨数。')),
        ]),
        (26, [
            ('B26', 314, 'inlineStr', inline_str('承运侧承压边界')),
            ('C26', 329, None, f'<f>E11+E12+E24</f><v>{carrier_bear_line}</v>'),
            ('D26', 318, None, None),
            ('E26', 329, None, f'<f>IF(D26=&quot;&quot;,E11+E12+E24,D26)</f><v>{carrier_bear_line}</v>'),
            ('F26', 316, 'inlineStr', inline_str('元/月')),
            ('G26', 317, 'inlineStr', inline_str('低于该边界后，承运司机费、管理费及毛利缺口由承运侧承担。')),
        ]),
        (27, [
            ('B27', 314, 'inlineStr', inline_str('承运侧承压边界折吨趟价')),
            ('C27', 329, None, f'<f>IFERROR(E26/E6,0)</f><v>{carrier_bear_line_price}</v>'),
            ('D27', 318, None, None),
            ('E27', 329, None, f'<f>IF(D27=&quot;&quot;,IFERROR(E26/E6,0),D27)</f><v>{carrier_bear_line_price}</v>'),
            ('F27', 316, 'inlineStr', inline_str('元/吨趟')),
            ('G27', 317, 'inlineStr', inline_str('用于识别承运执行层承压区间。')),
        ]),
        (28, [
            ('B28', 358, 'inlineStr', inline_str('三、风控判断')),
            ('C28', 358, None, None), ('D28', 358, None, None), ('E28', 358, None, None), ('F28', 358, None, None), ('G28', 358, None, None)
        ]),
        (29, [
            ('B29', 314, 'inlineStr', inline_str('差额/安全边际')),
            ('C29', 329, None, f'<f>E7-(E11+E12+E13+E24)</f><v>0</v>'),
            ('D29', 318, None, None),
            ('E29', 329, None, f'<f>IF(D29=&quot;&quot;,E7-(E11+E12+E13+E24),D29)</f><v>0</v>'),
            ('F29', 316, 'inlineStr', inline_str('元/月')),
            ('G29', 317, 'inlineStr', inline_str('承运毛利随运价自动变化时，该项应恒等于0；如手工覆盖承运毛利，则反映人工设定差额。')),
        ]),
        (30, [
            ('B30', 314, 'inlineStr', inline_str('风险区间判断')),
            ('C30', 315, None, f'<f>IF(E7&lt;E24,&quot;合同物流企业执行边界区&quot;,IF(E13&lt;=0,&quot;承运执行层自担区&quot;,&quot;正常执行区&quot;))</f><v>{default_zone}</v>'),
            ('D30', 317, None, None),
            ('E30', 315, None, f'<f>IF(D30=&quot;&quot;,C30,D30)</f><v>{default_zone}</v>'),
            ('F30', 316, 'inlineStr', inline_str('文本')),
            ('G30', 317, 'inlineStr', inline_str('先保合同物流企业刚性支出，再由承运侧承担执行层波动。')),
        ]),
        (31, [
            ('B31', 314, 'inlineStr', inline_str('覆盖校验')),
            ('C31', 315, None, f'<f>IF(E7&gt;=E24,&quot;通过&quot;,&quot;越界预警&quot;)</f><v>{default_check}</v>'),
            ('D31', 317, None, None),
            ('E31', 315, None, f'<f>IF(D31=&quot;&quot;,C31,D31)</f><v>{default_check}</v>'),
            ('F31', 316, 'inlineStr', inline_str('文本')),
            ('G31', 317, 'inlineStr', inline_str('仅当货主支付物流费用低于合同物流企业刚性支出覆盖要求时报警。')),
        ]),
        (32, [
            ('B32', 314, 'inlineStr', inline_str('动作结论')),
            ('C32', 315, None, f'<f>IF(E30=&quot;正常执行区&quot;,&quot;继续执行&quot;,IF(E30=&quot;承运执行层自担区&quot;,&quot;承运侧承担执行层缺口&quot;,&quot;暂停新增并启动价格重谈&quot;))</f><v>{default_action}</v>'),
            ('D32', 317, None, None),
            ('E32', 315, None, f'<f>IF(D32=&quot;&quot;,C32,D32)</f><v>{default_action}</v>'),
            ('F32', 316, 'inlineStr', inline_str('文本')),
            ('G32', 317, 'inlineStr', inline_str('承运毛利及承运执行层缺口均由承运物流企业承担。')),
        ]),
        (33, [
            ('B33', 314, 'inlineStr', inline_str('线序校验')),
            ('C33', 315, None, f'<f>IF(E25&lt;E8,&quot;通过&quot;,&quot;口径待复核&quot;)</f><v>{default_line_check}</v>'),
            ('D33', 317, None, None),
            ('E33', 315, None, f'<f>IF(D33=&quot;&quot;,C33,D33)</f><v>{default_line_check}</v>'),
            ('F33', 316, 'inlineStr', inline_str('文本')),
            ('G33', 317, 'inlineStr', inline_str('合同物流企业执行边界线应低于当前实际到手运价。')),
        ]),
        (35, [
            ('B35', 358, 'inlineStr', inline_str('四、平衡校验')),
            ('C35', 358, None, None), ('D35', 358, None, None), ('E35', 358, None, None), ('F35', 358, None, None), ('G35', 358, None, None)
        ]),
        (36, [
            ('B36', 314, 'inlineStr', inline_str('平衡关系')),
            ('C36', 315, 'inlineStr', inline_str('货主支付物流费用 = 承运司机费用 + 承运管理费用 + 承运毛利 + 合同物流企业刚性支出覆盖要求 + 差额/安全边际')),
            ('D36', 317, None, None),
            ('E36', 315, 'inlineStr', inline_str('承运毛利随运价自动变化，不再作为固定前提值。')),
            ('F36', 316, 'inlineStr', inline_str('文本')),
            ('G36', 317, 'inlineStr', inline_str('用于核对清分逻辑与风控判断的一致性。')),
        ]),
        (37, [
            ('B37', 314, 'inlineStr', inline_str('勾稽校验差额')),
            ('C37', 329, None, '<f>E7-(E11+E12+E13+E24+E29)</f><v>0</v>'),
            ('D37', 318, None, None),
            ('E37', 329, None, '<f>IF(D37=&quot;&quot;,E7-(E11+E12+E13+E24+E29),D37)</f><v>0</v>'),
            ('F37', 316, 'inlineStr', inline_str('元/月')),
            ('G3-', 317, 'inlineStr', inline_str('应恒等于0。')),
        ]),
        (38, [
            ('B38', 314, 'inlineStr', inline_str('使用方式')),
            ('C38', 315, 'inlineStr', inline_str('优先改D8')),
            ('D38', 317, None, None),
            ('E38', 315, 'inlineStr', inline_str('调整实际到手运价后，收入、承运毛利、风险区间、覆盖校验、动作结论、线序校验将自动联动；如需强制覆盖承运毛利，可改D13。')),
            ('F38', 316, 'inlineStr', inline_str('文本')),
            ('G38', 317, 'inlineStr', inline_str('如需直接录入收入，可在D7覆盖。')),
        ]),
    ]

    xml_rows = []
    for row_no, cells in rows:
        row_xml = [f'    <row r="{row_no}">']
        for ref, style, cell_type, payload in cells:
            if payload is None:
                row_xml.append(f'      <c r="{ref}" s="{style}"/>')
            elif cell_type == 'inlineStr':
                row_xml.append(f'      <c r="{ref}" s="{style}" t="inlineStr">{payload}</c>')
            else:
                row_xml.append(f'      <c r="{ref}" s="{style}">{payload}</c>')
        row_xml.append('    </row>')
        xml_rows.append("\n".join(row_xml))

    xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>
  <dimension ref="B2:G38"/>
  <sheetViews>
    <sheetView showGridLines="0" workbookViewId="0" zoomScale="90" zoomScaleNormal="90">
      <pane ySplit="5" topLeftCell="B6" activePane="bottomLeft" state="frozen"/>
      <selection pane="bottomLeft" activeCell="D8" sqref="D8"/>
    </sheetView>
  </sheetViews>
  <sheetFormatPr defaultColWidth="10" defaultRowHeight="15"/>
  <cols>
    <col min="2" max="2" width="24" customWidth="1"/>
    <col min="3" max="5" width="18" customWidth="1"/>
    <col min="6" max="6" width="12" customWidth="1"/>
    <col min="-" max="-" width="42" customWidth="1"/>
  </cols>
  <sheetData>
{chr(10).join(xml_rows)}
  </sheetData>
  <mergeCells count="6">
    <mergeCell ref="B2:G2"/>
    <mergeCell ref="B3:G3"/>
    <mergeCell ref="B10:G10"/>
    <mergeCell ref="B16:G16"/>
    <mergeCell ref="B28:G28"/>
    <mergeCell ref="B35:G35"/>
  </mergeCells>
  <pageMargins left="0.-5" right="0.-5" top="1" bottom="1" header="0.5" footer="0.5"/>
  <pageSetup orientation="landscape" horizontalDpi="300" verticalDpi="300"/>
</worksheet>
"""
    return xml.encode("utf-8")


def build():
    with zipfile.ZipFile(BASE_COPY) as zin, zipfile.ZipFile(OUTPUT, "w", zipfile.ZIP_DEFLATED) as zout:
        seen = set()
        for info in zin.infolist():
            if info.filename in seen:
                continue
            seen.add(info.filename)
            data = build_sheet10() if info.filename == "xl/worksheets/sheet10.xml" else zin.read(info.filename)
            zout.writestr(info.filename, data)
    print(OUTPUT.resolve())


if __name__ == "__main__":
    build()




