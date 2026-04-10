from pathlib import Path
import re
import zipfile
import xml.etree.ElementTree as ET


ROOT = Path(r"E:\省平台\运力管家-黄骅港项目")
SRC_DIR = ROOT / "99-temp" / "reextract"
OUT_FILE = ROOT / "99-temp" / "contract-logistics-breaker-fixed.xlsx"
SHEET_NAME = "7-双层熔断测算区"


NS = {"a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def read_text(path: Path) -> str:
    return path.read_text(encoding="utf-8")


def read_numeric(sheet: str, cell: str) -> float:
    root = ET.parse(SRC_DIR / "xl" / "worksheets" / sheet).getroot()
    node = root.find(f".//a:c[@r='{cell}']/a:v", NS)
    return float(node.text) if node is not None and node.text else 0.0


def build_sheet() -> str:
    revenue = read_numeric("sheet4.xml", "D6")
    tons = read_numeric("sheet2.xml", "C9") * read_numeric("sheet2.xml", "C12")
    driver = read_numeric("sheet4.xml", "D13")
    manage = read_numeric("sheet4.xml", "D15")
    carrier_margin = read_numeric("sheet5.xml", "D46")
    freight_interest = read_numeric("sheet5.xml", "D12")
    vehicle_rent = read_numeric("sheet3.xml", "H10")
    insurance = read_numeric("sheet4.xml", "D11")
    energy = read_numeric("sheet4.xml", "D12")
    maint = read_numeric("sheet4.xml", "D14")
    actual_price = revenue / tons if tons else 0.0
    available = revenue - driver - manage - carrier_margin
    rigid_cost = freight_interest + vehicle_rent + insurance + energy + maint
    breaker_line = rigid_cost / tons if tons else 0.0
    status = "熔断报警" if available < rigid_cost else "正常执行"
    action = "触发熔断预警，暂停新增并重谈价格" if available < rigid_cost else "继续执行"

    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>
  <dimension ref="B2:G33"/>
  <sheetViews>
    <sheetView showGridLines="0" workbookViewId="0" zoomScale="90" zoomScaleNormal="90">
      <pane ySplit="5" topLeftCell="B6" activePane="bottomLeft" state="frozen"/>
      <selection pane="bottomLeft" activeCell="E20" sqref="E20"/>
    </sheetView>
  </sheetViews>
  <sheetFormatPr defaultColWidth="10" defaultRowHeight="15"/>
  <cols>
    <col min="2" max="2" width="24" customWidth="1"/>
    <col min="3" max="5" width="18" customWidth="1"/>
    <col min="6" max="6" width="12" customWidth="1"/>
    <col min="7" max="7" width="38" customWidth="1"/>
  </cols>
  <sheetData>
    <row r="2" ht="20.25" customHeight="1">
      <c r="B2" s="360" t="inlineStr"><is><t>{SHEET_NAME}</t></is></c><c r="C2" s="360"/><c r="D2" s="360"/><c r="E2" s="360"/><c r="F2" s="360"/><c r="G2" s="360"/>
    </row>
    <row r="3">
      <c r="B3" s="361" t="inlineStr"><is><t>按单车口径模拟合同物流企业熔断。单位统一为元/吨趟，判断逻辑为“扣减承运侧后是否还能覆盖合同物流企业刚性支出”。</t></is></c><c r="C3" s="361"/><c r="D3" s="361"/><c r="E3" s="361"/><c r="F3" s="361"/><c r="G3" s="361"/>
    </row>
    <row r="5">
      <c r="B5" s="313" t="inlineStr"><is><t>项目</t></is></c><c r="C5" s="313" t="inlineStr"><is><t>默认引用值</t></is></c><c r="D5" s="313" t="inlineStr"><is><t>手工调整值</t></is></c><c r="E5" s="313" t="inlineStr"><is><t>采用值</t></is></c><c r="F5" s="313" t="inlineStr"><is><t>单位</t></is></c><c r="G5" s="313" t="inlineStr"><is><t>说明</t></is></c>
    </row>
    <row r="6">
      <c r="B6" s="314" t="inlineStr"><is><t>货主支付物流费用</t></is></c><c r="C6" s="318"><f>'3-收益测算'!D6</f><v>{revenue}</v></c><c r="D6" s="318"/><c r="E6" s="318"><f>IF(D6=&quot;&quot;,C6,D6)</f><v>{revenue}</v></c><c r="F6" s="316" t="inlineStr"><is><t>元/月</t></is></c><c r="G6" s="317" t="inlineStr"><is><t>单车月收入。</t></is></c>
    </row>
    <row r="7">
      <c r="B7" s="314" t="inlineStr"><is><t>单车月运输吨数</t></is></c><c r="C7" s="318"><f>'1-基础参数'!C9*'1-基础参数'!C12</f><v>{tons}</v></c><c r="D7" s="318"/><c r="E7" s="318"><f>IF(D7=&quot;&quot;,C7,D7)</f><v>{tons}</v></c><c r="F7" s="316" t="inlineStr"><is><t>吨/月</t></is></c><c r="G7" s="317" t="inlineStr"><is><t>载重×月趟数。</t></is></c>
    </row>
    <row r="8">
      <c r="B8" s="314" t="inlineStr"><is><t>实际到手运价</t></is></c><c r="C8" s="318"><f>IFERROR(E6/E7,0)</f><v>{actual_price}</v></c><c r="D8" s="318"/><c r="E8" s="318"><f>IF(D8=&quot;&quot;,C8,D8)</f><v>{actual_price}</v></c><c r="F8" s="316" t="inlineStr"><is><t>元/吨趟</t></is></c><c r="G8" s="317" t="inlineStr"><is><t>用于与熔断线比较。</t></is></c>
    </row>
    <row r="10" ht="15.75" customHeight="1">
      <c r="B10" s="358" t="inlineStr"><is><t>一、承运侧先扣项目</t></is></c><c r="C10" s="358"/><c r="D10" s="358"/><c r="E10" s="358"/><c r="F10" s="358"/><c r="G10" s="358"/>
    </row>
    <row r="11">
      <c r="B11" s="314" t="inlineStr"><is><t>承运司机费用</t></is></c><c r="C11" s="318"><f>'3-收益测算'!D13</f><v>{driver}</v></c><c r="D11" s="318"/><c r="E11" s="318"><f>IF(D11=&quot;&quot;,C11,D11)</f><v>{driver}</v></c><c r="F11" s="316" t="inlineStr"><is><t>元/月</t></is></c><c r="G11" s="317" t="inlineStr"><is><t>承运执行层人工。</t></is></c>
    </row>
    <row r="12">
      <c r="B12" s="314" t="inlineStr"><is><t>承运管理费用</t></is></c><c r="C12" s="318"><f>'3-收益测算'!D15</f><v>{manage}</v></c><c r="D12" s="318"/><c r="E12" s="318"><f>IF(D12=&quot;&quot;,C12,D12)</f><v>{manage}</v></c><c r="F12" s="316" t="inlineStr"><is><t>元/月</t></is></c><c r="G12" s="317" t="inlineStr"><is><t>承运执行层管理。</t></is></c>
    </row>
    <row r="13">
      <c r="B13" s="314" t="inlineStr"><is><t>承运毛利</t></is></c><c r="C13" s="318"><f>'3.5-资金清分'!D46</f><v>{carrier_margin}</v></c><c r="D13" s="318"/><c r="E13" s="318"><f>IF(D13=&quot;&quot;,C13,D13)</f><v>{carrier_margin}</v></c><c r="F13" s="316" t="inlineStr"><is><t>元/月</t></is></c><c r="G13" s="317" t="inlineStr"><is><t>承运物流企业执行层收益。</t></is></c>
    </row>
    <row r="14">
      <c r="B14" s="314" t="inlineStr"><is><t>扣减承运侧后可用金额</t></is></c><c r="C14" s="329"><f>E6-E11-E12-E13</f><v>{available}</v></c><c r="D14" s="318"/><c r="E14" s="329"><f>IF(D14=&quot;&quot;,C14,D14)</f><v>{available}</v></c><c r="F14" s="316" t="inlineStr"><is><t>元/月</t></is></c><c r="G14" s="317" t="inlineStr"><is><t>用于覆盖合同物流企业刚性支出。</t></is></c>
    </row>
    <row r="16" ht="15.75" customHeight="1">
      <c r="B16" s="358" t="inlineStr"><is><t>二、合同物流企业刚性支出</t></is></c><c r="C16" s="358"/><c r="D16" s="358"/><c r="E16" s="358"/><c r="F16" s="358"/><c r="G16" s="358"/>
    </row>
    <row r="17">
      <c r="B17" s="314" t="inlineStr"><is><t>垫资利息</t></is></c><c r="C17" s="318"><f>'3.5-资金清分'!D12</f><v>{freight_interest}</v></c><c r="D17" s="318"/><c r="E17" s="318"><f>IF(D17=&quot;&quot;,C17,D17)</f><v>{freight_interest}</v></c><c r="F17" s="316" t="inlineStr"><is><t>元/月</t></is></c><c r="G17" s="317" t="inlineStr"><is><t>合同物流企业实际垫资收益/成本口径。</t></is></c>
    </row>
    <row r="18">
      <c r="B18" s="314" t="inlineStr"><is><t>车辆租金（牵引车+挂车）</t></is></c><c r="C18" s="318"><f>'2-单车成本测算'!H10</f><v>{vehicle_rent}</v></c><c r="D18" s="318"/><c r="E18" s="318"><f>IF(D18=&quot;&quot;,C18,D18)</f><v>{vehicle_rent}</v></c><c r="F18" s="316" t="inlineStr"><is><t>元/月</t></is></c><c r="G18" s="317" t="inlineStr"><is><t>默认引用租赁口径，可按实际车租调整。</t></is></c>
    </row>
    <row r="19">
      <c r="B19" s="314" t="inlineStr"><is><t>保险费用</t></is></c><c r="C19" s="318"><f>'3-收益测算'!D11</f><v>{insurance}</v></c><c r="D19" s="318"/><c r="E19" s="318"><f>IF(D19=&quot;&quot;,C19,D19)</f><v>{insurance}</v></c><c r="F19" s="316" t="inlineStr"><is><t>元/月</t></is></c><c r="G19" s="317" t="inlineStr"><is><t>按单车月度保险口径。</t></is></c>
    </row>
    <row r="20">
      <c r="B20" s="314" t="inlineStr"><is><t>能源费用</t></is></c><c r="C20" s="318"><f>'3-收益测算'!D12</f><v>{energy}</v></c><c r="D20" s="318"/><c r="E20" s="318"><f>IF(D20=&quot;&quot;,C20,D20)</f><v>{energy}</v></c><c r="F20" s="316" t="inlineStr"><is><t>元/月</t></is></c><c r="G20" s="317" t="inlineStr"><is><t>合同物流企业实际承担的能源支出。</t></is></c>
    </row>
    <row r="21">
      <c r="B21" s="314" t="inlineStr"><is><t>维保费用</t></is></c><c r="C21" s="318"><f>'3-收益测算'!D14</f><v>{maint}</v></c><c r="D21" s="318"/><c r="E21" s="318"><f>IF(D21=&quot;&quot;,C21,D21)</f><v>{maint}</v></c><c r="F21" s="316" t="inlineStr"><is><t>元/月</t></is></c><c r="G21" s="317" t="inlineStr"><is><t>合同物流企业实际承担的维保支出。</t></is></c>
    </row>
    <row r="22">
      <c r="B22" s="314" t="inlineStr"><is><t>合同物流企业刚性支出合计</t></is></c><c r="C22" s="329"><f>SUM(E17:E21)</f><v>{rigid_cost}</v></c><c r="D22" s="318"/><c r="E22" s="329"><f>IF(D22=&quot;&quot;,C22,D22)</f><v>{rigid_cost}</v></c><c r="F22" s="316" t="inlineStr"><is><t>元/月</t></is></c><c r="G22" s="317" t="inlineStr"><is><t>熔断判断的核心覆盖成本。</t></is></c>
    </row>
    <row r="23">
      <c r="B23" s="314" t="inlineStr"><is><t>合同物流企业熔断线</t></is></c><c r="C23" s="329"><f>IFERROR(E22/E7,0)</f><v>{breaker_line}</v></c><c r="D23" s="318"/><c r="E23" s="329"><f>IF(D23=&quot;&quot;,C23,D23)</f><v>{breaker_line}</v></c><c r="F23" s="316" t="inlineStr"><is><t>元/吨趟</t></is></c><c r="G23" s="317" t="inlineStr"><is><t>刚性支出合计÷单车月运输吨数。</t></is></c>
    </row>
    <row r="25" ht="15.75" customHeight="1">
      <c r="B25" s="358" t="inlineStr"><is><t>三、熔断判断</t></is></c><c r="C25" s="358"/><c r="D25" s="358"/><c r="E25" s="358"/><c r="F25" s="358"/><c r="G25" s="358"/>
    </row>
    <row r="26">
      <c r="B26" s="314" t="inlineStr"><is><t>覆盖校验</t></is></c><c r="C26" s="315"><f>IF(E14&gt;=E22,&quot;通过&quot;,&quot;熔断报警&quot;)</f><v>{status}</v></c><c r="D26" s="317"/><c r="E26" s="315"><f>IF(D26=&quot;&quot;,C26,D26)</f><v>{status}</v></c><c r="F26" s="316" t="inlineStr"><is><t>文本</t></is></c><c r="G26" s="317" t="inlineStr"><is><t>扣减承运侧后可用金额是否足以覆盖刚性支出。</t></is></c>
    </row>
    <row r="27">
      <c r="B27" s="314" t="inlineStr"><is><t>动作结论</t></is></c><c r="C27" s="315"><f>IF(E26=&quot;通过&quot;,&quot;继续执行&quot;,&quot;触发熔断预警，暂停新增并重谈价格&quot;)</f><v>{action}</v></c><c r="D27" s="317"/><c r="E27" s="315"><f>IF(D27=&quot;&quot;,C27,D27)</f><v>{action}</v></c><c r="F27" s="316" t="inlineStr"><is><t>文本</t></is></c><c r="G27" s="317" t="inlineStr"><is><t>合同物流企业不承担运价下行风险。</t></is></c>
    </row>
    <row r="28">
      <c r="B28" s="314" t="inlineStr"><is><t>线序校验</t></is></c><c r="C28" s="315"><f>IF(E23&lt;E8,&quot;通过&quot;,&quot;口径待复核&quot;)</f><v>{"通过" if breaker_line < actual_price else "口径待复核"}</v></c><c r="D28" s="317"/><c r="E28" s="315"><f>IF(D28=&quot;&quot;,C28,D28)</f><v>{"通过" if breaker_line < actual_price else "口径待复核"}</v></c><c r="F28" s="316" t="inlineStr"><is><t>文本</t></is></c><c r="G28" s="317" t="inlineStr"><is><t>熔断线应低于当前实际到手运价。</t></is></c>
    </row>
    <row r="30">
      <c r="B30" s="314" t="inlineStr"><is><t>说明</t></is></c><c r="C30" s="315" t="inlineStr"><is><t>当前副本先锁合同物流企业熔断逻辑。</t></is></c><c r="D30" s="317"/><c r="E30" s="315" t="inlineStr"><is><t>承运调节线后续再单独设计。</t></is></c><c r="F30" s="316" t="inlineStr"><is><t>文本</t></is></c><c r="G30" s="317" t="inlineStr"><is><t>避免把承运风险和合同物流企业刚性支出混算。</t></is></c>
    </row>
    <row r="31">
      <c r="B31" s="314" t="inlineStr"><is><t>使用方式</t></is></c><c r="C31" s="315" t="inlineStr"><is><t>优先改D列</t></is></c><c r="D31" s="317"/><c r="E31" s="315" t="inlineStr"><is><t>按月输入实际承运毛利、车租、能源、维保、保险及结算金额。</t></is></c><c r="F31" s="316" t="inlineStr"><is><t>文本</t></is></c><c r="G31" s="317" t="inlineStr"><is><t>适合先做单车熔断模拟，再扩展到全项目。</t></is></c>
    </row>
  </sheetData>
  <mergeCells count="5">
    <mergeCell ref="B2:G2"/>
    <mergeCell ref="B3:G3"/>
    <mergeCell ref="B10:G10"/>
    <mergeCell ref="B16:G16"/>
    <mergeCell ref="B25:G25"/>
  </mergeCells>
  <pageMargins left="0.75" right="0.75" top="1" bottom="1" header="0.5" footer="0.5"/>
  <pageSetup orientation="landscape" horizontalDpi="300" verticalDpi="300"/>
</worksheet>
"""


def patch_workbook_xml(raw: str) -> str:
    if SHEET_NAME not in raw:
        raw = raw.replace("</sheets>", f'<sheet name="{SHEET_NAME}" sheetId="10" r:id="rId14"/></sheets>')
    raw = raw.replace('<calcPr calcId="191029"/>', '<calcPr calcId="191029" fullCalcOnLoad="1" forceFullCalc="1"/>')
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
    if f"<vt:lpstr>{SHEET_NAME}</vt:lpstr>" not in raw:
        raw = raw.replace("</vt:vector></TitlesOfParts>", f"<vt:lpstr>{SHEET_NAME}</vt:lpstr></vt:vector></TitlesOfParts>")
    return raw


def build():
    sheet10 = build_sheet().encode("utf-8")
    with zipfile.ZipFile(OUT_FILE, "w", zipfile.ZIP_DEFLATED) as zf:
        for file in SRC_DIR.rglob("*"):
            if file.is_dir():
                continue
            rel = file.relative_to(SRC_DIR).as_posix()
            if rel in ("xl/calcChain.xml", "xl/worksheets/sheet10.xml"):
                continue
            if rel == "xl/workbook.xml":
                zf.writestr(rel, patch_workbook_xml(read_text(file)).encode("utf-8"))
            elif rel == "xl/_rels/workbook.xml.rels":
                zf.writestr(rel, patch_workbook_rels(read_text(file)).encode("utf-8"))
            elif rel == "[Content_Types].xml":
                zf.writestr(rel, patch_content_types(read_text(file)).encode("utf-8"))
            elif rel == "docProps/app.xml":
                zf.writestr(rel, patch_app_xml(read_text(file)).encode("utf-8"))
            else:
                zf.write(file, rel)
        zf.writestr("xl/worksheets/sheet10.xml", sheet10)
    print(OUT_FILE)


if __name__ == "__main__":
    build()
