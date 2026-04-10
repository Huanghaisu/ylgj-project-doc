from pathlib import Path
import zipfile
import re


ROOT = Path(r"E:\省平台\运力管家-黄骅港项目")
EXTRACTED = ROOT / "99-temp" / "财务分析提取"
TARGET = ROOT / "04-财务模型" / "20-工作稿 黄骅港项目 - 财务分析.xlsx"
NEW_SHEET_NAME = "7-双层熔断测算区"


def read_text(path: Path) -> str:
    return path.read_text(encoding="utf-8")


def patch_workbook_xml(raw: str) -> str:
    if NEW_SHEET_NAME not in raw:
        raw = raw.replace("</sheets>", f'<sheet name="{NEW_SHEET_NAME}" sheetId="10" r:id="rId14"/></sheets>')
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


def build_sheet10() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>
  <dimension ref="B2:G31"/>
  <sheetViews>
    <sheetView showGridLines="0" workbookViewId="0" zoomScale="90" zoomScaleNormal="90">
      <pane ySplit="5" topLeftCell="B6" activePane="bottomLeft" state="frozen"/>
      <selection pane="bottomLeft" activeCell="E18" sqref="E18"/>
    </sheetView>
  </sheetViews>
  <sheetFormatPr defaultColWidth="10" defaultRowHeight="15"/>
  <cols>
    <col min="2" max="2" width="24" customWidth="1"/>
    <col min="3" max="5" width="18" customWidth="1"/>
    <col min="6" max="6" width="12" customWidth="1"/>
    <col min="7" max="7" width="34" customWidth="1"/>
  </cols>
  <sheetData>
    <row r="2" ht="20.25" customHeight="1">
      <c r="B2" s="360" t="inlineStr"><is><t>7-双层熔断测算区</t></is></c><c r="C2" s="360"/><c r="D2" s="360"/><c r="E2" s="360"/><c r="F2" s="360"/><c r="G2" s="360"/>
    </row>
    <row r="3">
      <c r="B3" s="361" t="inlineStr"><is><t>采用单车口径、单位统一为元/吨趟，先锁定项目总熔断线，再单独复核承运调节线。</t></is></c><c r="C3" s="361"/><c r="D3" s="361"/><c r="E3" s="361"/><c r="F3" s="361"/><c r="G3" s="361"/>
    </row>
    <row r="5">
      <c r="B5" s="313" t="inlineStr"><is><t>项目</t></is></c><c r="C5" s="313" t="inlineStr"><is><t>默认引用值</t></is></c><c r="D5" s="313" t="inlineStr"><is><t>手工调整值</t></is></c><c r="E5" s="313" t="inlineStr"><is><t>采用值</t></is></c><c r="F5" s="313" t="inlineStr"><is><t>单位</t></is></c><c r="G5" s="313" t="inlineStr"><is><t>说明</t></is></c>
    </row>
    <row r="6">
      <c r="B6" s="314" t="inlineStr"><is><t>名义运价</t></is></c><c r="C6" s="318"><f>'1-基础参数'!C6</f><v>210</v></c><c r="D6" s="318"/><c r="E6" s="318"><f>IF(D6="",C6,D6)</f><v>210</v></c><c r="F6" s="316" t="inlineStr"><is><t>元/吨趟</t></is></c><c r="G6" s="317" t="inlineStr"><is><t>项目对外名义运价。</t></is></c>
    </row>
    <row r="7">
      <c r="B7" s="314" t="inlineStr"><is><t>单车载重</t></is></c><c r="C7" s="318"><f>'1-基础参数'!C9</f><v>32</v></c><c r="D7" s="318"/><c r="E7" s="318"><f>IF(D7="",C7,D7)</f><v>32</v></c><c r="F7" s="316" t="inlineStr"><is><t>吨</t></is></c><c r="G7" s="317" t="inlineStr"><is><t>单车额定载重。</t></is></c>
    </row>
    <row r="8">
      <c r="B8" s="314" t="inlineStr"><is><t>单车月趟数</t></is></c><c r="C8" s="318"><f>'1-基础参数'!C12</f><v>11.2</v></c><c r="D8" s="318"/><c r="E8" s="318"><f>IF(D8="",C8,D8)</f><v>11.2</v></c><c r="F8" s="316" t="inlineStr"><is><t>趟/月</t></is></c><c r="G8" s="317" t="inlineStr"><is><t>单车月度执行趟数。</t></is></c>
    </row>
    <row r="9">
      <c r="B9" s="314" t="inlineStr"><is><t>单车月运输吨数</t></is></c><c r="C9" s="318"><f>E7*E8</f><v>358.4</v></c><c r="D9" s="318"/><c r="E9" s="318"><f>IF(D9="",C9,D9)</f><v>358.4</v></c><c r="F9" s="316" t="inlineStr"><is><t>吨/月</t></is></c><c r="G9" s="317" t="inlineStr"><is><t>载重×月趟数。</t></is></c>
    </row>
    <row r="11" ht="15.75" customHeight="1">
      <c r="B11" s="358" t="inlineStr"><is><t>一、单车损益模拟</t></is></c><c r="C11" s="358"/><c r="D11" s="358"/><c r="E11" s="358"/><c r="F11" s="358"/><c r="G11" s="358"/>
    </row>
    <row r="12">
      <c r="B12" s="314" t="inlineStr"><is><t>单车月收入</t></is></c><c r="C12" s="318"><f>'3-收益测算'!D6</f><v>75264</v></c><c r="D12" s="318"/><c r="E12" s="318"><f>IF(D12="",C12,D12)</f><v>75264</v></c><c r="F12" s="316" t="inlineStr"><is><t>元/月</t></is></c><c r="G12" s="317" t="inlineStr"><is><t>单车收入=运价×载重×月趟数。</t></is></c>
    </row>
    <row r="13">
      <c r="B13" s="314" t="inlineStr"><is><t>单车月总成本</t></is></c><c r="C13" s="318"><f>'3-收益测算'!D18</f><v>71505.9978166513</v></c><c r="D13" s="318"/><c r="E13" s="318"><f>IF(D13="",C13,D13)</f><v>71505.9978166513</v></c><c r="F13" s="316" t="inlineStr"><is><t>元/月</t></is></c><c r="G13" s="317" t="inlineStr"><is><t>引用资产车单车月度总成本。</t></is></c>
    </row>
    <row r="14">
      <c r="B14" s="314" t="inlineStr"><is><t>单车月毛利</t></is></c><c r="C14" s="318"><f>E12-E13</f><v>3758.00218334865</v></c><c r="D14" s="318"/><c r="E14" s="318"><f>IF(D14="",C14,D14)</f><v>3758.00218334865</v></c><c r="F14" s="316" t="inlineStr"><is><t>元/月</t></is></c><c r="G14" s="317" t="inlineStr"><is><t>用于判断单车是否仍有执行空间。</t></is></c>
    </row>
    <row r="15">
      <c r="B15" s="314" t="inlineStr"><is><t>单车实际到手运价</t></is></c><c r="C15" s="318"><f>IFERROR(E12/E9,0)</f><v>210</v></c><c r="D15" s="318"/><c r="E15" s="318"><f>IF(D15="",C15,D15)</f><v>210</v></c><c r="F15" s="316" t="inlineStr"><is><t>元/吨趟</t></is></c><c r="G15" s="317" t="inlineStr"><is><t>单车月收入÷单车月运输吨数。</t></is></c>
    </row>
    <row r="17" ht="15.75" customHeight="1">
      <c r="B17" s="358" t="inlineStr"><is><t>二、单车熔断判断</t></is></c><c r="C17" s="358"/><c r="D17" s="358"/><c r="E17" s="358"/><c r="F17" s="358"/><c r="G17" s="358"/>
    </row>
    <row r="18">
      <c r="B18" s="314" t="inlineStr"><is><t>项目总熔断线（单车）</t></is></c><c r="C18" s="329"><f>'3-收益测算'!D27</f><v>199.514502836639</v></c><c r="D18" s="318"/><c r="E18" s="329"><f>IF(D18="",C18,D18)</f><v>199.514502836639</v></c><c r="F18" s="316" t="inlineStr"><is><t>元/吨趟</t></is></c><c r="G18" s="317" t="inlineStr"><is><t>单车总成本÷单车月运输吨数。该值应低于名义运价。</t></is></c>
    </row>
    <row r="19">
      <c r="B19" s="314" t="inlineStr"><is><t>当前分包结算折吨趟价</t></is></c><c r="C19" s="318"><f>'3.5-资金清分'!D28/('1-基础参数'!C9*'1-基础参数'!C12)</f><v>34.1373623287671</v></c><c r="D19" s="318"/><c r="E19" s="318"><f>IF(D19="",C19,D19)</f><v>34.1373623287671</v></c><c r="F19" s="316" t="inlineStr"><is><t>元/吨趟</t></is></c><c r="G19" s="317" t="inlineStr"><is><t>仅反映现行清分口径，不作为正式项目熔断线。</t></is></c>
    </row>
    <row r="20">
      <c r="B20" s="314" t="inlineStr"><is><t>线序校验</t></is></c><c r="C20" s="315"><f>IF(E18&lt;E6,"通过","口径待复核")</f><v>通过</v></c><c r="D20" s="317"/><c r="E20" s="315"><f>IF(D20="",C20,D20)</f><v>通过</v></c><c r="F20" s="316" t="inlineStr"><is><t>文本</t></is></c><c r="G20" s="317" t="inlineStr"><is><t>总熔断线应低于名义运价。</t></is></c>
    </row>
    <row r="21">
      <c r="B21" s="314" t="inlineStr"><is><t>项目状态</t></is></c><c r="C21" s="315"><f>IF(E20&lt;&gt;"通过","口径待复核",IF(E15&gt;E18,"可执行","正式熔断"))</f><v>可执行</v></c><c r="D21" s="317"/><c r="E21" s="315"><f>IF(D21="",C21,D21)</f><v>可执行</v></c><c r="F21" s="316" t="inlineStr"><is><t>文本</t></is></c><c r="G21" s="317" t="inlineStr"><is><t>当前阶段先按单车总熔断线做正式判断。</t></is></c>
    </row>
    <row r="22">
      <c r="B22" s="314" t="inlineStr"><is><t>动作结论</t></is></c><c r="C22" s="315"><f>IF(E21="可执行","继续执行",IF(E21="正式熔断","停新增、保在途","先复核口径"))</f><v>继续执行</v></c><c r="D22" s="317"/><c r="E22" s="315"><f>IF(D22="",C22,D22)</f><v>继续执行</v></c><c r="F22" s="316" t="inlineStr"><is><t>文本</t></is></c><c r="G22" s="317" t="inlineStr"><is><t>单车实际到手运价高于单车总熔断线时继续执行。</t></is></c>
    </row>
    <row r="24" ht="15.75" customHeight="1">
      <c r="B24" s="358" t="inlineStr"><is><t>三、场景校验</t></is></c><c r="C24" s="358"/><c r="D24" s="358"/><c r="E24" s="358"/><c r="F24" s="358"/><c r="G24" s="358"/>
    </row>
    <row r="25">
      <c r="B25" s="314" t="inlineStr"><is><t>场景1：名义运价</t></is></c><c r="C25" s="318"><f>E6</f><v>210</v></c><c r="D25" s="317"/><c r="E25" s="315"><f>IF(C25&gt;E18,"可执行","正式熔断")</f><v>可执行</v></c><c r="F25" s="316" t="inlineStr"><is><t>元/吨趟</t></is></c><c r="G25" s="317" t="inlineStr"><is><t>验证名义运价高于总熔断线。</t></is></c>
    </row>
    <row r="26">
      <c r="B26" s="314" t="inlineStr"><is><t>场景2：总熔断线</t></is></c><c r="C26" s="318"><f>E18</f><v>199.514502836639</v></c><c r="D26" s="317"/><c r="E26" s="315"><f>IF(C26&gt;E18,"可执行","正式熔断")</f><v>正式熔断</v></c><c r="F26" s="316" t="inlineStr"><is><t>元/吨趟</t></is></c><c r="G26" s="317" t="inlineStr"><is><t>验证触线即熔断。</t></is></c>
    </row>
    <row r="27">
      <c r="B27" s="314" t="inlineStr"><is><t>场景3：低于总熔断线1元</t></is></c><c r="C27" s="318"><f>E18-1</f><v>198.514502836639</v></c><c r="D27" s="317"/><c r="E27" s="315"><f>IF(C27&gt;E18,"可执行","正式熔断")</f><v>正式熔断</v></c><c r="F27" s="316" t="inlineStr"><is><t>元/吨趟</t></is></c><c r="G27" s="317" t="inlineStr"><is><t>验证低于总熔断线时停止新增。</t></is></c>
    </row>
    <row r="29">
      <c r="B29" s="314" t="inlineStr"><is><t>说明</t></is></c><c r="C29" s="315" t="inlineStr"><is><t>当前页先锁单车项目总熔断线=199.51元/吨趟。</t></is></c><c r="D29" s="317"/><c r="E29" s="315" t="inlineStr"><is><t>承运调节线待按新的分责逻辑重构。</t></is></c><c r="F29" s="316" t="inlineStr"><is><t>文本</t></is></c><c r="G29" s="317" t="inlineStr"><is><t>避免把不同层级口径直接叠加，导致熔断线高于名义运价。</t></is></c>
    </row>
    <row r="31">
      <c r="B31" s="314" t="inlineStr"><is><t>使用方式</t></is></c><c r="C31" s="315" t="inlineStr"><is><t>优先改D列</t></is></c><c r="D31" s="317"/><c r="E31" s="315" t="inlineStr"><is><t>单车月收入、单车月成本、实际运价均可按月覆盖。</t></is></c><c r="F31" s="316" t="inlineStr"><is><t>文本</t></is></c><c r="G31" s="317" t="inlineStr"><is><t>先把单车熔断跑通，再决定是否扩展到车队总量或双层分责。</t></is></c>
    </row>
  </sheetData>
  <mergeCells count="5">
    <mergeCell ref="B2:G2"/>
    <mergeCell ref="B3:G3"/>
    <mergeCell ref="B11:G11"/>
    <mergeCell ref="B17:G17"/>
    <mergeCell ref="B24:G24"/>
  </mergeCells>
  <pageMargins left="0.75" right="0.75" top="1" bottom="1" header="0.5" footer="0.5"/>
  <pageSetup orientation="landscape" horizontalDpi="300" verticalDpi="300"/>
</worksheet>
"""


def rebuild():
    workbook_xml = patch_workbook_xml(read_text(EXTRACTED / "xl" / "workbook.xml"))
    workbook_rels = patch_workbook_rels(read_text(EXTRACTED / "xl" / "_rels" / "workbook.xml.rels"))
    content_types = patch_content_types(read_text(EXTRACTED / "[Content_Types].xml"))
    app_xml = patch_app_xml(read_text(EXTRACTED / "docProps" / "app.xml"))
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
        zf.writestr("xl/worksheets/sheet10.xml", build_sheet10().encode("utf-8"))
    tmp.replace(TARGET)
    print("rebuilt workbook with single-vehicle breaker sheet")


if __name__ == "__main__":
    rebuild()
