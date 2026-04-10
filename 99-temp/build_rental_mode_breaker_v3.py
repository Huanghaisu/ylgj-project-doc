from pathlib import Path
import zipfile
import xml.etree.ElementTree as ET


ROOT = Path.cwd()
SOURCE_BOOK = next(ROOT.glob("04-*/*.xlsx"))
BASE_COPY = ROOT / "99-temp" / "contract-logistics-breaker-fixed-dedup.xlsx"
OUTPUT = ROOT / "99-temp" / "contract-logistics-breaker-rental-v4.xlsx"
NS = {"a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def read_numeric(book: Path, sheet: str, cell: str) -> float:
    with zipfile.ZipFile(book) as zf:
        root = ET.fromstring(zf.read(f"xl/worksheets/{sheet}"))
    node = root.find(f".//a:c[@r='{cell}']/a:v", NS)
    return float(node.text) if node is not None and node.text else 0.0


def build_sheet10() -> bytes:
    tons = read_numeric(SOURCE_BOOK, "sheet2.xml", "C9") * read_numeric(SOURCE_BOOK, "sheet2.xml", "C12")

    default_revenue = read_numeric(SOURCE_BOOK, "sheet4.xml", "E6")
    driver = read_numeric(SOURCE_BOOK, "sheet4.xml", "E13")
    manage = read_numeric(SOURCE_BOOK, "sheet4.xml", "E15")
    carrier_gross_default = read_numeric(SOURCE_BOOK, "sheet5.xml", "E46")

    freight_interest = read_numeric(SOURCE_BOOK, "sheet5.xml", "E12")
    rent_interest = read_numeric(SOURCE_BOOK, "sheet5.xml", "E13")
    vehicle_rent = read_numeric(SOURCE_BOOK, "sheet5.xml", "E15")
    insurance = read_numeric(SOURCE_BOOK, "sheet4.xml", "E11")
    energy = read_numeric(SOURCE_BOOK, "sheet4.xml", "E12")
    maint = read_numeric(SOURCE_BOOK, "sheet4.xml", "E14")

    rigid_cover = freight_interest + rent_interest + vehicle_rent + insurance + energy + maint
    trigger_revenue = driver + manage + carrier_gross_default + rigid_cover
    trigger_price = trigger_revenue / tons if tons else 0.0
    actual_price = default_revenue / tons if tons else 0.0
    available = default_revenue - driver - manage - carrier_gross_default
    margin = available - rigid_cover
    trigger_revenue = driver + manage + carrier_gross_default + rigid_cover
    trigger_price = trigger_revenue / tons if tons else 0.0

    xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>
  <dimension ref="B2:G40"/>
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
    <col min="7" max="7" width="40" customWidth="1"/>
  </cols>
  <sheetData>
    <row r="2" ht="20.25" customHeight="1">
      <c r="B2" s="360" t="inlineStr"><is><t>7-双层熔断测算区</t></is></c><c r="C2" s="360"/><c r="D2" s="360"/><c r="E2" s="360"/><c r="F2" s="360"/><c r="G2" s="360"/>
    </row>
    <row r="3">
      <c r="B3" s="361" t="inlineStr"><is><t>采用租赁模式单车口径。以“实际到手运价”为主调参入口，自动反推货主支付物流费用，并联动风控结论。</t></is></c><c r="C3" s="361"/><c r="D3" s="361"/><c r="E3" s="361"/><c r="F3" s="361"/><c r="G3" s="361"/>
    </row>
    <row r="5">
      <c r="B5" s="313" t="inlineStr"><is><t>项目</t></is></c><c r="C5" s="313" t="inlineStr"><is><t>默认引用值</t></is></c><c r="D5" s="313" t="inlineStr"><is><t>手工调整值</t></is></c><c r="E5" s="313" t="inlineStr"><is><t>采用值</t></is></c><c r="F5" s="313" t="inlineStr"><is><t>单位</t></is></c><c r="G5" s="313" t="inlineStr"><is><t>说明</t></is></c>
    </row>

    <row r="6">
      <c r="B6" s="314" t="inlineStr"><is><t>单车月运输吨数</t></is></c><c r="C6" s="318"><f>'1-基础参数'!C9*'1-基础参数'!C12</f><v>{tons}</v></c><c r="D6" s="318"/><c r="E6" s="318"><f>IF(D6=&quot;&quot;,C6,D6)</f><v>{tons}</v></c><c r="F6" s="316" t="inlineStr"><is><t>吨/月</t></is></c><c r="G6" s="317" t="inlineStr"><is><t>载重×月趟数。</t></is></c>
    </row>
    <row r="7">
      <c r="B7" s="314" t="inlineStr"><is><t>货主支付物流费用</t></is></c><c r="C7" s="318"><f>'3-收益测算'!E6</f><v>{default_revenue}</v></c><c r="D7" s="318"/><c r="E7" s="318"><f>IF(D7=&quot;&quot;,IF(D8=&quot;&quot;,C7,D8*E6),D7)</f><v>{default_revenue}</v></c><c r="F7" s="316" t="inlineStr"><is><t>元/月</t></is></c><c r="G7" s="317" t="inlineStr"><is><t>默认引用收入；若调整实际到手运价，则自动反推。</t></is></c>
    </row>
    <row r="8">
      <c r="B8" s="314" t="inlineStr"><is><t>实际到手运价</t></is></c><c r="C8" s="318"><f>IFERROR(C7/C6,0)</f><v>{actual_price}</v></c><c r="D8" s="318"/><c r="E8" s="318"><f>IF(D8=&quot;&quot;,IFERROR(E7/E6,0),D8)</f><v>{actual_price}</v></c><c r="F8" s="316" t="inlineStr"><is><t>元/吨趟</t></is></c><c r="G8" s="317" t="inlineStr"><is><t>主调参入口；调整后反推货主支付物流费用。</t></is></c>
    </row>

    <row r="10" ht="15.75" customHeight="1">
      <c r="B10" s="358" t="inlineStr"><is><t>一、承运侧先扣项目</t></is></c><c r="C10" s="358"/><c r="D10" s="358"/><c r="E10" s="358"/><c r="F10" s="358"/><c r="G10" s="358"/>
    </row>
    <row r="11">
      <c r="B11" s="314" t="inlineStr"><is><t>承运司机费用</t></is></c><c r="C11" s="318"><f>'3-收益测算'!E13</f><v>{driver}</v></c><c r="D11" s="318"/><c r="E11" s="318"><f>IF(D11=&quot;&quot;,C11,D11)</f><v>{driver}</v></c><c r="F11" s="316" t="inlineStr"><is><t>元/月</t></is></c><c r="G11" s="317" t="inlineStr"><is><t>承运执行层司机成本。</t></is></c>
    </row>
    <row r="12">
      <c r="B12" s="314" t="inlineStr"><is><t>承运管理费用</t></is></c><c r="C12" s="318"><f>'3-收益测算'!E15</f><v>{manage}</v></c><c r="D12" s="318"/><c r="E12" s="318"><f>IF(D12=&quot;&quot;,C12,D12)</f><v>{manage}</v></c><c r="F12" s="316" t="inlineStr"><is><t>元/月</t></is></c><c r="G12" s="317" t="inlineStr"><is><t>承运执行层管理成本。</t></is></c>
    </row>
    <row r="13">
      <c r="B13" s="314" t="inlineStr"><is><t>承运毛利</t></is></c><c r="C13" s="318"><f>'3.5-资金清分'!E46</f><v>{carrier_gross_default}</v></c><c r="D13" s="318"/><c r="E13" s="318"><f>IF(D13=&quot;&quot;,C13,D13)</f><v>{carrier_gross_default}</v></c><c r="F13" s="316" t="inlineStr"><is><t>元/月</t></is></c><c r="G13" s="317" t="inlineStr"><is><t>变量项，可按清分逻辑或谈判结果调整。</t></is></c>
    </row>
    <row r="14">
      <c r="B14" s="314" t="inlineStr"><is><t>扣减承运侧后可用金额</t></is></c><c r="C14" s="329"><f>E7-E11-E12-E13</f><v>{available}</v></c><c r="D14" s="318"/><c r="E14" s="329"><f>IF(D14=&quot;&quot;,C14,D14)</f><v>{available}</v></c><c r="F14" s="316" t="inlineStr"><is><t>元/月</t></is></c><c r="G14" s="317" t="inlineStr"><is><t>货主支付物流费用扣减承运司机、管理、毛利后的剩余金额。</t></is></c>
    </row>

    <row r="16" ht="15.75" customHeight="1">
      <c r="B16" s="358" t="inlineStr"><is><t>二、合同物流企业刚性支出覆盖要求</t></is></c><c r="C16" s="358"/><c r="D16" s="358"/><c r="E16" s="358"/><c r="F16" s="358"/><c r="G16" s="358"/>
    </row>
    <row r="17">
      <c r="B17" s="314" t="inlineStr"><is><t>运费垫资利息</t></is></c><c r="C17" s="318"><f>'3.5-资金清分'!E12</f><v>{freight_interest}</v></c><c r="D17" s="318"/><c r="E17" s="318"><f>IF(D17=&quot;&quot;,C17,D17)</f><v>{freight_interest}</v></c><c r="F17" s="316" t="inlineStr"><is><t>元/月</t></is></c><c r="G17" s="317" t="inlineStr"><is><t>合同物流企业资金收益/覆盖要求。</t></is></c>
    </row>
    <row r="18">
      <c r="B18" s="314" t="inlineStr"><is><t>车租垫资利息</t></is></c><c r="C18" s="318"><f>'3.5-资金清分'!E13</f><v>{rent_interest}</v></c><c r="D18" s="318"/><c r="E18" s="318"><f>IF(D18=&quot;&quot;,C18,D18)</f><v>{rent_interest}</v></c><c r="F18" s="316" t="inlineStr"><is><t>元/月</t></is></c><c r="G18" s="317" t="inlineStr"><is><t>租赁模式车租账期收益/覆盖要求。</t></is></c>
    </row>
    <row r="19">
      <c r="B19" s="314" t="inlineStr"><is><t>车辆租金（牵引车+挂车）</t></is></c><c r="C19" s="318"><f>'3.5-资金清分'!E15</f><v>{vehicle_rent}</v></c><c r="D19" s="318"/><c r="E19" s="318"><f>IF(D19=&quot;&quot;,C19,D19)</f><v>{vehicle_rent}</v></c><c r="F19" s="316" t="inlineStr"><is><t>元/月</t></is></c><c r="G19" s="317" t="inlineStr"><is><t>租赁模式下牵引车+挂车租金。</t></is></c>
    </row>
    <row r="20">
      <c r="B20" s="314" t="inlineStr"><is><t>保险费用</t></is></c><c r="C20" s="318"><f>'3-收益测算'!E11</f><v>{insurance}</v></c><c r="D20" s="318"/><c r="E20" s="318"><f>IF(D20=&quot;&quot;,C20,D20)</f><v>{insurance}</v></c><c r="F20" s="316" t="inlineStr"><is><t>元/月</t></is></c><c r="G20" s="317" t="inlineStr"><is><t>租赁模式默认值为0，保留手工覆盖口。</t></is></c>
    </row>
    <row r="21">
      <c r="B21" s="314" t="inlineStr"><is><t>能源费用</t></is></c><c r="C21" s="318"><f>'3-收益测算'!E12</f><v>{energy}</v></c><c r="D21" s="318"/><c r="E21" s="318"><f>IF(D21=&quot;&quot;,C21,D21)</f><v>{energy}</v></c><c r="F21" s="316" t="inlineStr"><is><t>元/月</t></is></c><c r="G21" s="317" t="inlineStr"><is><t>合同物流企业实际承担的能源支出。</t></is></c>
    </row>
    <row r="22">
      <c r="B22" s="314" t="inlineStr"><is><t>维保费用</t></is></c><c r="C22" s="318"><f>'3-收益测算'!E14</f><v>{maint}</v></c><c r="D22" s="318"/><c r="E22" s="318"><f>IF(D22=&quot;&quot;,C22,D22)</f><v>{maint}</v></c><c r="F22" s="316" t="inlineStr"><is><t>元/月</t></is></c><c r="G22" s="317" t="inlineStr"><is><t>合同物流企业实际承担的维保支出。</t></is></c>
    </row>
    <row r="23">
      <c r="B23" s="314" t="inlineStr"><is><t>刚性支出覆盖要求合计</t></is></c><c r="C23" s="329"><f>SUM(E17:E22)</f><v>{rigid_cover}</v></c><c r="D23" s="318"/><c r="E23" s="329"><f>IF(D23=&quot;&quot;,C23,D23)</f><v>{rigid_cover}</v></c><c r="F23" s="316" t="inlineStr"><is><t>元/月</t></is></c><c r="G23" s="317" t="inlineStr"><is><t>合同物流企业要求被覆盖的刚性支出与资金收益。</t></is></c>
    </row>
    <row r="24">
      <c r="B24" s="314" t="inlineStr"><is><t>合同物流企业熔断线</t></is></c><c r="C24" s="329"><f>IFERROR(E23/E6,0)</f><v>{rigid_cover / tons if tons else 0.0}</v></c><c r="D24" s="318"/><c r="E24" s="329"><f>IF(D24=&quot;&quot;,C24,D24)</f><v>{rigid_cover / tons if tons else 0.0}</v></c><c r="F24" s="316" t="inlineStr"><is><t>元/吨趟</t></is></c><c r="G24" s="317" t="inlineStr"><is><t>刚性支出覆盖要求合计÷单车月运输吨数。</t></is></c>
    </row>
    <row r="25">
      <c r="B25" s="314" t="inlineStr"><is><t>货主支付物流费用熔断点</t></is></c><c r="C25" s="329"><f>E11+E12+E13+E23</f><v>{trigger_revenue}</v></c><c r="D25" s="318"/><c r="E25" s="329"><f>IF(D25=&quot;&quot;,C25,D25)</f><v>{trigger_revenue}</v></c><c r="F25" s="316" t="inlineStr"><is><t>元/月</t></is></c><c r="G25" s="317" t="inlineStr"><is><t>承运毛利为变量时，该熔断点同步变化。</t></is></c>
    </row>
    <row r="26">
      <c r="B26" s="314" t="inlineStr"><is><t>货主支付物流费用熔断点折吨趟价</t></is></c><c r="C26" s="329"><f>IFERROR(E25/E6,0)</f><v>{trigger_price}</v></c><c r="D26" s="318"/><c r="E26" s="329"><f>IF(D26=&quot;&quot;,C26,D26)</f><v>{trigger_price}</v></c><c r="F26" s="316" t="inlineStr"><is><t>元/吨趟</t></is></c><c r="G26" s="317" t="inlineStr"><is><t>用于与实际到手运价直接比较。</t></is></c>
    </row>

    <row r="28" ht="15.75" customHeight="1">
      <c r="B28" s="358" t="inlineStr"><is><t>三、风控判断</t></is></c><c r="C28" s="358"/><c r="D28" s="358"/><c r="E28" s="358"/><c r="F28" s="358"/><c r="G28" s="358"/>
    </row>
    <row r="29">
      <c r="B29" s="314" t="inlineStr"><is><t>差额/安全边际</t></is></c><c r="C29" s="329"><f>E14-E23</f><v>{margin}</v></c><c r="D29" s="318"/><c r="E29" s="329"><f>IF(D29=&quot;&quot;,C29,D29)</f><v>{margin}</v></c><c r="F29" s="316" t="inlineStr"><is><t>元/月</t></is></c><c r="G29" s="317" t="inlineStr"><is><t>扣减承运侧后用于覆盖合同物流企业刚性支出的余额。</t></is></c>
    </row>
    <row r="30">
      <c r="B30" s="314" t="inlineStr"><is><t>风险区间判断</t></is></c><c r="C30" s="315"><f>IF(E7&gt;=E25,&quot;正常执行区&quot;,IF(E7&gt;=E25-E13,&quot;承运毛利受压区&quot;,IF(E7&gt;=E23,&quot;承运执行层自担区&quot;,&quot;合同物流企业熔断区&quot;)))</f><v>{"正常执行区" if default_revenue >= trigger_revenue else ("承运毛利受压区" if default_revenue >= trigger_revenue - carrier_gross_default else ("承运执行层自担区" if default_revenue >= rigid_cover else "合同物流企业熔断区"))}</v></c><c r="D30" s="317"/><c r="E30" s="315"><f>IF(D30=&quot;&quot;,C30,D30)</f><v>{"正常执行区" if default_revenue >= trigger_revenue else ("承运毛利受压区" if default_revenue >= trigger_revenue - carrier_gross_default else ("承运执行层自担区" if default_revenue >= rigid_cover else "合同物流企业熔断区"))}</v></c><c r="F30" s="316" t="inlineStr"><is><t>文本</t></is></c><c r="G30" s="317" t="inlineStr"><is><t>承运毛利变化时，区间边界同步变化。</t></is></c>
    </row>
    <row r="31">
      <c r="B31" s="314" t="inlineStr"><is><t>覆盖校验</t></is></c><c r="C31" s="315"><f>IF(E7&gt;=E25,&quot;通过&quot;,&quot;熔断报警&quot;)</f><v>{"通过" if default_revenue >= trigger_revenue else "熔断报警"}</v></c><c r="D31" s="317"/><c r="E31" s="315"><f>IF(D31=&quot;&quot;,C31,D31)</f><v>{"通过" if default_revenue >= trigger_revenue else "熔断报警"}</v></c><c r="F31" s="316" t="inlineStr"><is><t>文本</t></is></c><c r="G31" s="317" t="inlineStr"><is><t>只有不足覆盖合同物流企业熔断点时才触发熔断报警。</t></is></c>
    </row>
    <row r="32">
      <c r="B32" s="314" t="inlineStr"><is><t>动作结论</t></is></c><c r="C32" s="315"><f>IF(E30=&quot;正常执行区&quot;,&quot;继续执行&quot;,IF(E30=&quot;承运毛利受压区&quot;,&quot;承运毛利压缩/归零&quot;,IF(E30=&quot;承运执行层自担区&quot;,&quot;承运侧承担执行层缺口&quot;,&quot;暂停新增并启动价格重谈&quot;)))</f><v>{"继续执行" if default_revenue >= trigger_revenue else ("承运毛利压缩/归零" if default_revenue >= trigger_revenue - carrier_gross_default else ("承运侧承担执行层缺口" if default_revenue >= rigid_cover else "暂停新增并启动价格重谈"))}</v></c><c r="D32" s="317"/><c r="E32" s="315"><f>IF(D32=&quot;&quot;,C32,D32)</f><v>{"继续执行" if default_revenue >= trigger_revenue else ("承运毛利压缩/归零" if default_revenue >= trigger_revenue - carrier_gross_default else ("承运侧承担执行层缺口" if default_revenue >= rigid_cover else "暂停新增并启动价格重谈"))}</v></c><c r="F32" s="316" t="inlineStr"><is><t>文本</t></is></c><c r="G32" s="317" t="inlineStr"><is><t>承运毛利为变量时，动作结论同步变化。</t></is></c>
    </row>
    <row r="33">
      <c r="B33" s="314" t="inlineStr"><is><t>线序校验</t></is></c><c r="C33" s="315"><f>IF(E26&lt;E8,&quot;通过&quot;,&quot;口径待复核&quot;)</f><v>{"通过" if trigger_price < actual_price else "口径待复核"}</v></c><c r="D33" s="317"/><c r="E33" s="315"><f>IF(D33=&quot;&quot;,C33,D33)</f><v>{"通过" if trigger_price < actual_price else "口径待复核"}</v></c><c r="F33" s="316" t="inlineStr"><is><t>文本</t></is></c><c r="G33" s="317" t="inlineStr"><is><t>熔断变量折吨趟价应低于当前实际到手运价。</t></is></c>
    </row>

    <row r="35" ht="15.75" customHeight="1">
      <c r="B35" s="358" t="inlineStr"><is><t>四、平衡校验</t></is></c><c r="C35" s="358"/><c r="D35" s="358"/><c r="E35" s="358"/><c r="F35" s="358"/><c r="G35" s="358"/>
    </row>
    <row r="36">
      <c r="B36" s="314" t="inlineStr"><is><t>平衡关系</t></is></c><c r="C36" s="315" t="inlineStr"><is><t>货主支付物流费用 = 承运司机费用 + 承运管理费用 + 承运毛利 + 合同物流企业刚性支出覆盖要求 + 差额/安全边际</t></is></c><c r="D36" s="317"/><c r="E36" s="315" t="inlineStr"><is><t>货主支付物流费用熔断点已包含承运毛利，不得再与承运毛利重复相加。</t></is></c><c r="F36" s="316" t="inlineStr"><is><t>文本</t></is></c><c r="G36" s="317" t="inlineStr"><is><t>用于解释71,682与4,012不能再次相加的原因。</t></is></c>
    </row>
    <row r="37">
      <c r="B37" s="314" t="inlineStr"><is><t>勾稽校验差额</t></is></c><c r="C37" s="329"><f>E7-(E11+E12+E13+E23+E29)</f><v>0</v></c><c r="D37" s="318"/><c r="E37" s="329"><f>IF(D37=&quot;&quot;,C37,D37)</f><v>0</v></c><c r="F37" s="316" t="inlineStr"><is><t>元/月</t></is></c><c r="G37" s="317" t="inlineStr"><is><t>应恒等于0。</t></is></c>
    </row>
    <row r="38">
      <c r="B38" s="314" t="inlineStr"><is><t>使用方式</t></is></c><c r="C38" s="315" t="inlineStr"><is><t>优先改D8和D13</t></is></c><c r="D38" s="317"/><c r="E38" s="315" t="inlineStr"><is><t>调整实际到手运价后自动反推收入；调整承运毛利后同步刷新熔断点、风险区间和动作结论。</t></is></c><c r="F38" s="316" t="inlineStr"><is><t>文本</t></is></c><c r="G38" s="317" t="inlineStr"><is><t>如需直接录入收入，可在D7覆盖。</t></is></c>
    </row>
  </sheetData>
  <mergeCells count="6">
    <mergeCell ref="B2:G2"/>
    <mergeCell ref="B3:G3"/>
    <mergeCell ref="B10:G10"/>
    <mergeCell ref="B16:G16"/>
    <mergeCell ref="B28:G28"/>
    <mergeCell ref="B35:G35"/>
  </mergeCells>
  <pageMargins left="0.75" right="0.75" top="1" bottom="1" header="0.5" footer="0.5"/>
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
