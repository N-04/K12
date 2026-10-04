"""生成包含内嵌数据工作簿的可编辑 PPT 图表。"""

import html
from decimal import Decimal, InvalidOperation
from io import BytesIO
import zipfile
from typing import Any

from .excel import excel_column_letters


REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
OFFICE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def chart_has_values(series: dict[str, Any]) -> bool:
    """判断系列是否包含实际数值，区分零与空白。"""
    return any(value not in (None, "") for value in series["values"].values())


def build_chart(chart: dict[str, Any]) -> tuple[str, bytes]:
    """将系列缓存写成原生图表和对应的内嵌 Excel 数据。"""
    series = chart.get("series", [])
    indexes = sorted({index for item in series for index in set(item["categories"]) | set(item["values"])})
    if not series or not indexes or any(not chart_has_values(item) for item in series):
        raise ValueError("图表没有可用数据")
    if indexes[-1] >= 200000:
        raise ValueError("图表点数过多，请分批转换")
    indexes = list(range(indexes[-1] + 1))
    categories = []
    for index in indexes:
        labels = {item["categories"][index] for item in series if item["categories"].get(index)}
        if len(labels) > 1:
            raise ValueError("图表系列类别不同，暂以数据页保留")
        categories.append(next(iter(labels), ""))
    rows: list[list[Any]] = [["类别", *[item["name"] for item in series]]]
    for offset, index in enumerate(indexes):
        row: list[Any] = [categories[offset]]
        for item in series:
            value = item["values"].get(index, "")
            if value != "":
                try:
                    number = Decimal(str(value))
                except (InvalidOperation, ValueError) as exc:
                    raise ValueError("图表数值不是有效数字，暂以数据页保留") from exc
                if not number.is_finite():
                    raise ValueError("图表数值不是有限数字")
                value = number
            row.append(value)
        rows.append(row)
    category_cache = _cache(categories, False)
    parts = []
    for index, item in enumerate(series):
        column = excel_column_letters(index + 2)
        values = [row[index + 1] for row in rows[1:]]
        marker_xml, smooth_xml = "", ""
        if chart.get("chart_type") == "lineChart":
            symbol = item.get("marker_symbol") or ("circle" if chart.get("line_marker") else "none")
            if symbol not in {"auto", "circle", "dash", "diamond", "dot", "none", "picture", "plus", "square", "star", "triangle", "x"}:
                raise ValueError("折线标记类型无效")
            if symbol == "picture":
                raise ValueError("图片折线标记暂以数据页保留")
            size_xml = ""
            if item.get("marker_size") is not None:
                try:
                    size = int(item["marker_size"])
                except (TypeError, ValueError) as exc:
                    raise ValueError("折线标记大小无效") from exc
                if not 2 <= size <= 72:
                    raise ValueError("折线标记大小超出范围")
                size_xml = f'<c:size val="{size}"/>'
            marker_xml = f'<c:marker><c:symbol val="{symbol}"/>{size_xml}</c:marker>'
            smooth_xml = f'<c:smooth val="{int(bool(item.get("smooth")))}"/>'
        parts.append(f'<c:ser><c:idx val="{index}"/><c:order val="{index}"/><c:tx><c:strRef><c:f>Data!${column}$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>{html.escape(item["name"])}</c:v></c:pt></c:strCache></c:strRef></c:tx>{marker_xml}<c:cat><c:strRef><c:f>Data!$A$2:$A${len(rows)}</c:f>{category_cache}</c:strRef></c:cat><c:val><c:numRef><c:f>Data!${column}$2:${column}${len(rows)}</c:f>{_cache(values, True)}</c:numRef></c:val>{smooth_xml}</c:ser>')
    plot = _chart_plot(chart, "".join(parts))
    xml = f'<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="{OFFICE_REL}"><c:chart><c:autoTitleDeleted val="1"/><c:plotArea>{plot}</c:plotArea><c:legend><c:legendPos val="r"/></c:legend><c:dispBlanksAs val="gap"/></c:chart><c:externalData r:id="workbook"><c:autoUpdate val="0"/></c:externalData></c:chartSpace>'
    return xml, _chart_workbook(rows)


def _chart_plot(chart: dict[str, Any], series_xml: str) -> str:
    """按源类型生成柱状、折线或单系列饼图的绘图区。"""
    chart_type = chart.get("chart_type", "barChart")
    if chart_type == "pieChart":
        if len(chart["series"]) != 1:
            raise ValueError("饼图需要单个系列，暂以数据页保留")
        return f'<c:pieChart><c:varyColors val="1"/>{series_xml}<c:firstSliceAng val="0"/></c:pieChart>'
    if chart_type not in {"barChart", "lineChart"}:
        raise ValueError("暂不支持此图表类型，请查看图表数据页")
    direction = "bar" if chart.get("bar_direction") == "bar" and chart_type == "barChart" else "col"
    grouping = chart.get("grouping", "clustered")
    allowed = {"standard", "stacked", "percentStacked"} if chart_type == "lineChart" else {"clustered", "stacked", "percentStacked", "standard"}
    if grouping not in allowed:
        grouping = "standard" if chart_type == "lineChart" else "clustered"
    prefix = f'<c:barDir val="{direction}"/>' if chart_type == "barChart" else ""
    overlap = '<c:overlap val="100"/>' if chart_type == "barChart" and grouping in {"stacked", "percentStacked"} else ""
    marker = "1" if chart.get("line_marker") else "0"
    options = f'<c:marker val="{marker}"/><c:smooth val="0"/>' if chart_type == "lineChart" else overlap
    category_position, value_position = ("l", "b") if direction == "bar" else ("b", "l")
    axes = f'<c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="{category_position}"/><c:tickLblPos val="nextTo"/><c:crossAx val="2"/><c:crosses val="autoZero"/></c:catAx><c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="{value_position}"/><c:majorGridlines/><c:numFmt formatCode="General" sourceLinked="1"/><c:tickLblPos val="nextTo"/><c:crossAx val="1"/><c:crosses val="autoZero"/></c:valAx>'
    return f'<c:{chart_type}>{prefix}<c:grouping val="{grouping}"/>{series_xml}{options}<c:axId val="1"/><c:axId val="2"/></c:{chart_type}>{axes}'


def _cache(values: list[Any], numeric: bool) -> str:
    """写入完整点数与稀疏缓存，空数值不伪造为零。"""
    tag = "numCache" if numeric else "strCache"
    points = "".join(f'<c:pt idx="{index}"><c:v>{html.escape(str(value))}</c:v></c:pt>' for index, value in enumerate(values) if not numeric or value != "")
    format_code = "<c:formatCode>General</c:formatCode>" if numeric else ""
    return f'<c:{tag}>{format_code}<c:ptCount val="{len(values)}"/>{points}</c:{tag}>'


def _chart_workbook(rows: list[list[Any]]) -> bytes:
    """打包可由 PowerPoint 编辑图表数据的最小 XLSX。"""
    sheet_rows = []
    for row_index, row in enumerate(rows, 1):
        cells = []
        for column, value in enumerate(row, 1):
            ref = f"{excel_column_letters(column)}{row_index}"
            if isinstance(value, (int, float, Decimal)):
                cells.append(f'<c r="{ref}"><v>{value}</v></c>')
            else:
                cells.append(f'<c r="{ref}" t="inlineStr"><is><t xml:space="preserve">{html.escape(str(value))}</t></is></c>')
        sheet_rows.append(f'<row r="{row_index}">{"".join(cells)}</row>')
    buffer = BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>')
        archive.writestr("_rels/.rels", f'<Relationships xmlns="{REL_NS}"><Relationship Id="book" Type="{OFFICE_REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>')
        archive.writestr("xl/workbook.xml", f'<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="{OFFICE_REL}"><sheets><sheet name="Data" sheetId="1" r:id="sheet"/></sheets></workbook>')
        archive.writestr("xl/_rels/workbook.xml.rels", f'<Relationships xmlns="{REL_NS}"><Relationship Id="sheet" Type="{OFFICE_REL}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>')
        archive.writestr("xl/worksheets/sheet1.xml", '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>' + "".join(sheet_rows) + '</sheetData></worksheet>')
    return buffer.getvalue()


def chart_frame() -> str:
    """生成指向图表部件且尺寸明确的幻灯片图形框。"""
    return f'<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="4" name="Chart"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="457200" y="1463040"/><a:ext cx="8229600" cy="4937760"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="{OFFICE_REL}" r:id="chart"/></a:graphicData></a:graphic></p:graphicFrame>'
