"""实际启动桌面包，验证资源、接口、客户端与独立数据目录。"""

import argparse
import base64
import hashlib
import json
import zipfile
import socket
import subprocess
import sys
import tempfile
import time
from pathlib import Path
from xml.etree import ElementTree as ET
from urllib.error import URLError
from urllib.request import Request, urlopen

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from k12.local_client import _macos_office_output_validation_error
from k12.media import IMAGE_TYPES


def check_macro_sources(origin: str) -> None:
    """验证实际桌面程序的宏源导入、列表和删除。"""
    source = 'Attribute VB_Name = "PackageModule"\nPublic Sub PackageMacro()\nEnd Sub'
    request = Request(origin + "/api/macros/import", data=json.dumps({"file_name": "package.bas", "source": source}).encode(), headers={"Content-Type": "application/json"})
    with urlopen(request, timeout=15) as response:
        assert response.status == 201, "包内宏源导入失败"
        imported = json.load(response)["source"]
    assert "source" not in imported, "导入响应泄露宏源正文"
    assert imported["sha256"] == hashlib.sha256(source.encode()).hexdigest(), "宏源哈希不匹配"
    with urlopen(origin + "/api/macros", timeout=15) as response:
        macros = json.load(response)["macros"]
    found = [item for item in macros if item.get("source_id") == imported["id"]]
    assert len(found) == 1 and found[0]["macro_name"] == "PackageModule.PackageMacro", "包内真实宏声明未进入列表"
    request = Request(origin + "/api/macros/sources/" + imported["id"], method="DELETE")
    with urlopen(request, timeout=15) as response:
        assert json.load(response)["deleted"], "包内导入模块删除失败"
    with urlopen(origin + "/api/macros", timeout=15) as response:
        assert not any(item.get("source_id") == imported["id"] for item in json.load(response)["macros"]), "删除后宏列表残留"


def check_image_format_conversion(origin: str, root: Path, image_format: str = "svg", top_down: bool = False) -> None:
    """通过包内 API 验证图片格式的实际双向保留。"""
    source = root / "vector.docx"
    vector = b'<svg xmlns="http://www.w3.org/2000/svg" width="8" height="4"><rect width="8" height="4" fill="red"/></svg>'
    raster = base64.b64decode('iVBORw0KGgoAAAANSUhEUgAAAAIAAAABCAIAAAB7QOjdAAAAD0lEQVR4nGP4z8DA8J8BAAf/Af8Bf4mnAAAAAElFTkSuQmCC')
    with zipfile.ZipFile(source, "w") as archive:
        archive.writestr("[Content_Types].xml", '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="svg" ContentType="image/svg+xml"/></Types>')
        archive.writestr("word/document.xml", '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t>SVG 验证</w:t></w:r></w:p><w:p><w:r><w:drawing><wp:inline><wp:extent cx="2000000" cy="1000000"/><a:blip r:embed="image"><a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}"><asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="vector"/></a:ext></a:extLst></a:blip></wp:inline></w:drawing></w:r></w:p></w:body></w:document>')
        archive.writestr("word/_rels/document.xml.rels", '<Relationships><Relationship Id="image" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/source.png"/><Relationship Id="vector" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/source.svg"/></Relationships>')
        archive.writestr("word/media/source.png", raster)
        archive.writestr("word/media/source.svg", vector)
    expected_media = [vector, raster]
    if image_format != "svg":
        fixture = Path(__file__).resolve().parents[1] / "tests/fixtures/images" / ("sample." + image_format)
        image = fixture.read_bytes()
        if top_down:
            # 夹具为单色 BMP，反转存储方向后像素内容仍相同。
            height = int.from_bytes(image[22:26], "little", signed=True)
            image = image[:22] + (-height).to_bytes(4, "little", signed=True) + image[26:]
        with zipfile.ZipFile(source) as archive:
            parts = {name: archive.read(name) for name in archive.namelist() if "/media/" not in name}
        document = ET.fromstring(parts["word/document.xml"])
        blip = document.find(".//{*}blip")
        for child in list(blip):
            blip.remove(child)
        parts["word/document.xml"] = ET.tostring(document)
        relations = ET.fromstring(parts["word/_rels/document.xml.rels"])
        relations.remove(next(node for node in relations if node.get("Id") == "vector"))
        relations[0].set("Target", "media/source." + image_format)
        parts["word/_rels/document.xml.rels"] = ET.tostring(relations)
        content_types = ET.fromstring(parts["[Content_Types].xml"])
        if not any(node.get("Extension") == image_format for node in content_types):
            ET.SubElement(content_types, "{http://schemas.openxmlformats.org/package/2006/content-types}Default", {"Extension": image_format, "ContentType": IMAGE_TYPES[image_format]})
        parts["[Content_Types].xml"] = ET.tostring(content_types)
        parts["word/media/source." + image_format] = image
        with zipfile.ZipFile(source, "w") as archive:
            for name, data in parts.items():
                archive.writestr(name, data)
        expected_media = [image]
    def post(route: str, payload: dict) -> dict:
        """提交临时验证请求。"""
        request = Request(origin + route, data=json.dumps(payload).encode(), headers={"Content-Type": "application/json"})
        with urlopen(request, timeout=15) as response:
            return json.load(response)
    for kind, extension in (("word_to_ppt", ".pptx"), ("ppt_to_word", ".docx")):
        file = post("/api/files", {"file_name": source.name, "file_path": str(source), "file_size": source.stat().st_size})["file"]
        if top_down and kind == "word_to_ppt":
            scan = post("/api/tasks", {"task_type": "small_image_scan", "file_ids": [file["id"]]})["task"]
            assert scan["status"] == "成功", "自上而下 BMP 检索失败"
            with urlopen(origin + "/api/reports?task_id=" + scan["id"], timeout=15) as response:
                images = json.load(response)["reports"][0]["analysis"]["smallImages"]
            assert len(images) == 1 and images[0]["width"] == 8 and images[0]["height"] == 4, "自上而下 BMP 尺寸被误判"
        task = post("/api/tasks", {"task_type": kind, "file_ids": [file["id"]]})["task"]
        assert task["status"] == "成功", f"包内 {image_format} 转换失败：{kind} {task.get('error_message', '')}"
        with urlopen(origin + "/api/reports?task_id=" + task["id"], timeout=15) as response:
            reports = json.load(response)["reports"]
        artifact = next(report for report in reports if report["task_id"] == task["id"])["analysis"]["artifacts"][0]
        with urlopen(origin + artifact["url"], timeout=3) as response:
            content = response.read()
        assert hashlib.sha256(content).hexdigest() == artifact["sha256"], "SVG 产物下载哈希不匹配"
        source = root / ("vector-result" + extension)
        source.write_bytes(content)
        assert not _macos_office_output_validation_error(source, extension), "SVG 产物关系无效"
        with zipfile.ZipFile(source) as archive:
            media = [archive.read(name) for name in archive.namelist() if "/media/" in name]
            assert len(media) == len(expected_media) and all(image in media for image in expected_media), f"包内 {image_format} 图片内容丢失"


def check_conversion(origin: str, root: Path) -> None:
    """通过真实 API 转换临时文档并复验下载内容。"""
    source = root / "sample.docx"
    image = base64.b64decode('iVBORw0KGgoAAAANSUhEUgAAAAIAAAABCAIAAAB7QOjdAAAAD0lEQVR4nGP4z8DA8J8BAAf/Af8Bf4mnAAAAAElFTkSuQmCC')
    with zipfile.ZipFile(source, "w") as archive:
        archive.writestr("[Content_Types].xml", '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>')
        archive.writestr("word/document.xml", '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>打包验证标题</w:t></w:r></w:p><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:i/><w:u w:val="double" w:color="CC3300"/><w:sz w:val="32"/><w:rFonts w:ascii="Times New Roman" w:eastAsia="宋体"/></w:rPr><w:t>PACKAGE_CONVERSION_BODY</w:t></w:r></w:p><w:tbl><w:tblGrid><w:gridCol w:w="1500"/><w:gridCol w:w="3000"/></w:tblGrid><w:tr><w:trPr><w:trHeight w:val="360"/></w:trPr><w:tc><w:tcPr><w:gridSpan w:val="2"/><w:vMerge w:val="restart"/></w:tcPr><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>PACKAGE_MERGED_TABLE</w:t></w:r></w:p><w:p/><w:p/></w:tc></w:tr><w:tr><w:trPr><w:trHeight w:val="720"/></w:trPr><w:tc><w:tcPr><w:gridSpan w:val="2"/><w:vMerge/></w:tcPr><w:p/></w:tc></w:tr></w:tbl><w:p><w:r><w:t>PACKAGE_AFTER_TABLE</w:t></w:r></w:p><w:p><w:r><w:pict xmlns:v="urn:schemas-microsoft-com:vml" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><v:shape style="width:144pt;height:72pt;rotation:-90;flip:x y"><v:imagedata r:id="image" cropleft="8192f" croptop=".25" cropright="10%"/></v:shape></w:pict></w:r></w:p></w:body></w:document>')
        archive.writestr("word/media/source.png", image)
        archive.writestr("word/_rels/document.xml.rels", '<Relationships><Relationship Id="image" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/source.png"/></Relationships>')
    def post(route: str, payload: dict) -> dict:
        """提交临时测试请求并读取 JSON。"""
        request = Request(origin + route, data=json.dumps(payload).encode(), headers={"Content-Type": "application/json"})
        with urlopen(request, timeout=15) as response:
            return json.load(response)
    file = post("/api/files", {"file_name": source.name, "file_path": str(source), "file_size": source.stat().st_size})["file"]
    scan = post("/api/tasks", {"task_type": "small_image_scan", "file_ids": [file["id"]]})["task"]
    assert scan["status"] == "成功", "包内图片检索失败"
    with urlopen(origin + "/api/reports?task_id=" + scan["id"], timeout=15) as response:
        scan_report = next(item for item in json.load(response)["reports"] if item["task_id"] == scan["id"])
    scanned = scan_report["analysis"]["smallImages"]
    assert len(scanned) == 1 and "正文" in scanned[0]["location"], "包内图片真实位置缺失"
    assert "第 8 段" in scanned[0]["location"], "包内图片段落位置缺失"
    assert scanned[0]["page_index"] == 0, "包内图片媒体编号被误当作页码"
    with urlopen(origin + f"/api/files/{file['id']}/preview", timeout=3) as response:
        preview = json.load(response)["preview"]
    image_sources = [page for page in preview["pages"] if page["kind"] == "图片来源记录"]
    assert len(image_sources) == 1 and scanned[0]["location"] in image_sources[0]["text"], "包内来源预览与图片报告位置不一致"
    task = post("/api/tasks", {"task_type": "word_to_ppt", "file_ids": [file["id"]]})["task"]
    assert task["status"] == "成功", "包内 Word 转 PPT 失败"
    with urlopen(origin + "/api/reports?task_id=" + task["id"], timeout=15) as response:
        reports = json.load(response)["reports"]
    report = next(item for item in reports if item["task_id"] == task["id"])
    artifact = report["analysis"]["artifacts"][0]
    assert "图片按独立页保留" in artifact["message"], "Word 转 PPT 布局提示丢失"
    with urlopen(origin + artifact["url"], timeout=3) as response:
        content = response.read()
    assert hashlib.sha256(content).hexdigest() == artifact["sha256"], "下载哈希不匹配"
    output = root / "result.pptx"
    output.write_bytes(content)
    assert not _macos_office_output_validation_error(output, ".pptx"), "包内 PPT 转换产物关系不完整"
    with zipfile.ZipFile(output) as archive:
        text = b"".join(archive.read(name) for name in archive.namelist() if name.startswith("ppt/slides/") and name.endswith(".xml"))
        assert b"PACKAGE_CONVERSION_BODY" in text, "包内转换遗漏正文"
        styled_runs = [run for name in archive.namelist() if name.startswith("ppt/slides/slide") and name.endswith(".xml")
                       for run in ET.fromstring(archive.read(name)).findall(".//{*}r")
                       if run.findtext("{*}t") == "PACKAGE_CONVERSION_BODY"]
        assert len(styled_runs) == 1, "包内格式正文缺失或重复"
        paragraphs = [paragraph for name in archive.namelist() if name.startswith("ppt/slides/slide") and name.endswith(".xml")
                      for paragraph in ET.fromstring(archive.read(name)).findall(".//{*}p")
                      if paragraph.findtext("{*}r/{*}t") == "PACKAGE_CONVERSION_BODY"]
        assert paragraphs[0].find("{*}pPr").get("algn") == "ctr", "包内正文居中丢失"
        properties = styled_runs[0].find("{*}rPr")
        assert properties.get("b") == "1" and properties.get("i") == "1" and properties.get("sz") == "1600", "包内加粗、斜体或字号丢失"
        assert properties.find("{*}latin").get("typeface") == "Times New Roman", "包内英文字体丢失"
        assert properties.find("{*}ea").get("typeface") == "宋体", "包内中文字体丢失"
        assert properties.get("u") == "dbl", "包内双下划线丢失"
        assert properties.find("{*}uFill/{*}solidFill/{*}srgbClr").get("val") == "CC3300", "包内下划线颜色丢失"
        media = [name for name in archive.namelist() if name.startswith("ppt/media/")]
        assert len(media) == 1 and archive.read(media[0]) == image, "PPT 图片字节丢失或改变"
        assert b"<p:pic>" in text and b'r:embed="image"' in text, "PPT 图片形状或引用丢失"
        crops = [node.attrib for name in archive.namelist() if name.startswith("ppt/slides/") and name.endswith(".xml")
                 for node in ET.fromstring(archive.read(name)).findall(".//{*}srcRect")]
        assert crops == [{"l": "12500", "t": "25000", "r": "10000"}], "包内旧式图片裁剪丢失"
        transforms = [node.attrib for name in archive.namelist() if name.startswith("ppt/slides/") and name.endswith(".xml")
                      for node in ET.fromstring(archive.read(name)).findall(".//{*}pic/{*}spPr/{*}xfrm")]
        assert transforms == [{"rot": "16200000", "flipH": "1", "flipV": "1"}], "包内旧式图片旋转或翻转丢失"
        assert b"PACKAGE_AFTER_TABLE" in text, "包内转换遗漏表格后正文"
        cells = [cell for name in archive.namelist() if name.startswith("ppt/slides/") and name.endswith(".xml")
                 for cell in ET.fromstring(archive.read(name)).findall(".//{*}tc")]
        origin_cell = next((cell for cell in cells if cell.get("rowSpan") == "2" and cell.get("gridSpan") == "2"), None)
        assert origin_cell is not None, "包内转换遗漏合并表格"
        assert "PACKAGE_MERGED_TABLE" in "".join(origin_cell.itertext()), "合并表格正文缺失"
        assert origin_cell.find("{*}txBody/{*}p/{*}pPr").get("algn") == "ctr", "PPT 表格居中丢失"
        assert len(origin_cell.findall("{*}txBody/{*}p")) == 3, "PPT 表格空段落丢失"
        assert any(cell.get("hMerge") == "1" and cell.get("vMerge") == "1" for cell in cells), "合并覆盖单元格缺失"
        tables = [table for name in archive.namelist() if name.startswith("ppt/slides/") and name.endswith(".xml")
                  for table in ET.fromstring(archive.read(name)).findall(".//{*}tbl")]
        widths = [int(column.get("w")) for column in tables[0].findall("{*}tblGrid/{*}gridCol")]
        heights = [int(row.get("h")) for row in tables[0].findall("{*}tr")]
        assert widths[1] == 2 * widths[0], "包内转换未保留列宽比例"
        assert heights[1] == 2 * heights[0], "包内转换未保留行高比例"
    # 在实际包内 PPT 上加入外部关系，确认往返转换不误读网络目标。
    with zipfile.ZipFile(output) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    # 将格式移到母版、主题及本地默认层，验证实际包内的继承链。
    master_name = "ppt/slideMasters/slideMaster1.xml"
    master = ET.fromstring(parts[master_name])
    master_properties = master.find("{*}txStyles/{*}otherStyle/{*}lvl1pPr/{*}defRPr")
    master_properties.set("b", "1")
    master_properties.set("i", "1")
    master_properties.set("lang", "zh-CN")
    parts[master_name] = ET.tostring(master)
    theme_name = "ppt/theme/theme1.xml"
    theme = ET.fromstring(parts[theme_name])
    theme.find("{*}themeElements/{*}fontScheme/{*}minorFont/{*}latin").set("typeface", "Times New Roman")
    theme.find("{*}themeElements/{*}fontScheme/{*}minorFont/{*}ea").set("typeface", "宋体")
    parts[theme_name] = ET.tostring(theme)
    for name in list(parts):
        if not name.startswith("ppt/slides/slide") or not name.endswith(".xml"):
            continue
        slide = ET.fromstring(parts[name])
        for shape in slide.findall(".//{*}sp"):
            for paragraph in shape.findall("{*}txBody/{*}p"):
                run = next((run for run in paragraph.findall("{*}r") if run.findtext("{*}t") == "PACKAGE_CONVERSION_BODY"), None)
                if run is None:
                    continue
                namespace = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
                properties = run.find("{*}rPr")
                run.remove(properties)
                properties.attrib.pop("b", None)
                properties.attrib.pop("i", None)
                for font in list(properties):
                    if font.tag.rsplit("}", 1)[-1] in {"latin", "ea"}:
                        properties.remove(font)
                list_style = shape.find("{*}txBody/{*}lstStyle")
                level = ET.SubElement(list_style, namespace + "lvl1pPr")
                ET.SubElement(level, namespace + "defRPr", {"sz": properties.attrib.pop("sz")})
                paragraph_properties = paragraph.find("{*}pPr")
                if paragraph_properties is None:
                    paragraph_properties = ET.Element(namespace + "pPr")
                    paragraph.insert(0, paragraph_properties)
                properties.tag = namespace + "defRPr"
                paragraph_properties.append(properties)
        parts[name] = ET.tostring(slide)
    relation_name = "ppt/slides/_rels/slide1.xml.rels"
    relations = ET.fromstring(parts[relation_name])
    ET.SubElement(relations, "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship",
                  {"Id": "link", "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
                   "Target": "https://example.com/course", "TargetMode": "External"})
    parts[relation_name] = ET.tostring(relations)
    with zipfile.ZipFile(output, "w") as archive:
        for name, data in parts.items():
            archive.writestr(name, data)
    file = post("/api/files", {"file_name": output.name, "file_path": str(output), "file_size": output.stat().st_size})["file"]
    task = post("/api/tasks", {"task_type": "ppt_to_word", "file_ids": [file["id"]]})["task"]
    assert task["status"] == "成功", "包内 PPT 转 Word 失败"
    with urlopen(origin + "/api/reports?task_id=" + task["id"], timeout=15) as response:
        reports = json.load(response)["reports"]
    artifact = next(item for item in reports if item["task_id"] == task["id"])["analysis"]["artifacts"][0]
    assert "图片按行内绘图保留" in artifact["message"], "PPT 转 Word 布局提示丢失"
    with urlopen(origin + artifact["url"], timeout=3) as response:
        content = response.read()
    assert hashlib.sha256(content).hexdigest() == artifact["sha256"], "反向转换下载哈希不匹配"
    output = root / "roundtrip.docx"
    output.write_bytes(content)
    assert not _macos_office_output_validation_error(output, ".docx"), "包内 Word 转换产物关系不完整"
    with zipfile.ZipFile(output) as archive:
        document = ET.fromstring(archive.read("word/document.xml"))
        media = [name for name in archive.namelist() if name.startswith("word/media/")]
        assert len(media) == 1 and archive.read(media[0]) == image, "Word 图片字节丢失或改变"
        assert len(document.findall(".//{*}drawing")) == 1, "Word 图片绘图丢失"
        assert document.find(".//{*}srcRect").attrib == {"l": "12500", "t": "25000", "r": "10000", "b": "0"}, "Word 往返裁剪丢失"
        assert document.find(".//{*}xfrm").attrib == {"rot": "16200000", "flipH": "1", "flipV": "1"}, "Word 往返旋转或翻转丢失"
        extent = document.find(".//{*}inline/{*}extent")
        effects = document.find(".//{*}effectExtent")
        assert effects is not None and int(effects.get("t")) > 0, "Word 旋转图片未预留行高"
        assert int(extent.get("cx")) + int(effects.get("l")) + int(effects.get("r")) <= 5715001, "Word 旋转图片超出正文宽度"
    cells = document.findall(".//{*}tbl/{*}tr/{*}tc")
    value_key = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val"
    styled = [run for run in document.findall(".//{*}r") if run.findtext("{*}t") == "PACKAGE_CONVERSION_BODY"]
    assert len(styled) == 1, "包内反向格式正文缺失或重复"
    paragraph = next(node for node in document.findall(".//{*}p") if node.findtext("{*}r/{*}t") == "PACKAGE_CONVERSION_BODY")
    assert paragraph.find("{*}pPr/{*}jc").get(value_key) == "center", "包内反向正文居中丢失"
    properties = styled[0].find("{*}rPr")
    assert properties.find("{*}b").get(value_key) == "1" and properties.find("{*}i").get(value_key) == "1", "包内反向默认加粗或斜体丢失"
    assert properties.find("{*}sz").get(value_key) == "32", "包内反向列表字号丢失"
    fonts = properties.find("{*}rFonts")
    font_key = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    assert fonts.get(font_key + "ascii") == "Times New Roman" and fonts.get(font_key + "eastAsia") == "宋体", "包内反向字体丢失"
    underline = properties.find("{*}u")
    assert underline.get(value_key) == "double" and underline.get(font_key + "color") == "CC3300", "包内反向双下划线或颜色丢失"
    assert len(cells) == 2, "Word 合并网格不正确"
    assert all(cell.find("{*}tcPr/{*}gridSpan").get(value_key) == "2" for cell in cells), "Word 横向合并丢失"
    assert [cell.find("{*}tcPr/{*}vMerge").get(value_key) for cell in cells] == ["restart", "continue"], "Word 纵向合并丢失"
    assert "".join(document.itertext()).count("PACKAGE_MERGED_TABLE") == 1, "合并正文重复或丢失"
    assert cells[0].find("{*}p/{*}pPr/{*}jc").get(value_key) == "center", "Word 表格居中丢失"
    assert len(cells[0].findall("{*}p")) == 3, "Word 表格空段落丢失"


def check_excel_chart(origin: str, root: Path) -> None:
    """验证包内 Excel 转 PPT 输出原生图表和内嵌工作簿。"""
    source = root / "chart.xlsx"
    with zipfile.ZipFile(source, "w") as archive:
        archive.writestr("xl/workbook.xml", '<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="统计" r:id="sheet"/></sheets></workbook>')
        archive.writestr("xl/_rels/workbook.xml.rels", '<Relationships><Relationship Id="sheet" Target="worksheets/sheet1.xml"/></Relationships>')
        archive.writestr("xl/worksheets/sheet1.xml", '<worksheet><sheetData/></worksheet>')
        archive.writestr("xl/worksheets/_rels/sheet1.xml.rels", '<Relationships><Relationship Target="../drawings/drawing1.xml"/><Relationship Target="../comments1.xml"/></Relationships>')
        archive.writestr("xl/comments1.xml", '<comments><authors><author>打包验证</author></authors><commentList><comment ref="A1" authorId="0"><text><t>PACKAGE_EXCEL_COMMENT</t></text></comment></commentList></comments>')
        archive.writestr("xl/drawings/drawing1.xml", '<drawing/>')
        archive.writestr("xl/drawings/_rels/drawing1.xml.rels", '<Relationships><Relationship Target="../charts/chart1.xml"/></Relationships>')
        archive.writestr("xl/charts/chart1.xml", '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:lineChart><c:ser><c:tx><c:v>收入</c:v></c:tx><c:cat><c:strLit><c:pt idx="0"><c:v>一月</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>98</c:v></c:pt></c:numLit></c:val></c:ser></c:lineChart></c:plotArea></c:chart></c:chartSpace>')
    def post(route: str, payload: dict) -> dict:
        """提交临时图表转换请求。"""
        request = Request(origin + route, data=json.dumps(payload).encode(), headers={"Content-Type": "application/json"})
        with urlopen(request, timeout=15) as response:
            return json.load(response)
    file = post("/api/files", {"file_name": source.name, "file_path": str(source), "file_size": source.stat().st_size})["file"]
    task = post("/api/tasks", {"task_type": "excel_to_ppt", "file_ids": [file["id"]], "options": {"excel": {"retainComments": True}}})["task"]
    assert task["status"] == "成功", "包内 Excel 图表转换失败"
    with urlopen(origin + "/api/reports?task_id=" + task["id"], timeout=15) as response:
        reports = json.load(response)["reports"]
    artifact = next(item for item in reports if item["task_id"] == task["id"])["analysis"]["artifacts"][0]
    with urlopen(origin + artifact["url"], timeout=3) as response:
        content = response.read()
    assert hashlib.sha256(content).hexdigest() == artifact["sha256"], "图表产物下载哈希不匹配"
    output = root / "chart.pptx"
    output.write_bytes(content)
    with zipfile.ZipFile(output) as archive:
        charts = [name for name in archive.namelist() if name.startswith("ppt/charts/chart")]
        assert charts and b"lineChart" in archive.read(charts[0]), "原生折线图缺失"
        slides = b"".join(archive.read(name) for name in archive.namelist() if name.startswith("ppt/slides/") and name.endswith(".xml"))
        assert b"PACKAGE_EXCEL_COMMENT" in slides, "包内 Excel 批注正文缺失"
        assert any(name.startswith("ppt/embeddings/") and name.endswith(".xlsx") for name in archive.namelist()), "图表内嵌工作簿缺失"


def main() -> None:
    """使用临时数据与空闲端口，结束后关闭本次启动的进程。"""
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('executable', type=Path)
    args = parser.parse_args()
    executable = args.executable.resolve()
    subprocess.run([str(executable), '--local-client', '--help'], check=True, timeout=15,
                   stdout=subprocess.DEVNULL)
    with socket.socket() as probe:
        probe.bind(('127.0.0.1', 0))
        port = probe.getsockname()[1]
    with tempfile.TemporaryDirectory(prefix='k12-desktop-check-') as tmp:
        root = Path(tmp)
        data = root / 'data'
        with (root / 'process.log').open('wb') as log:
            process = subprocess.Popen([str(executable), '--no-open-browser', '--port', str(port),
                                        '--data-dir', str(data)], cwd=root, stdout=log, stderr=log)
            origin = f'http://127.0.0.1:{port}'
            try:
                deadline = time.monotonic() + 20
                while time.monotonic() < deadline:
                    if process.poll() is not None:
                        raise RuntimeError('桌面包提前退出：' + (root / 'process.log').read_text(errors='replace'))
                    try:
                        with urlopen(origin + '/', timeout=1) as response:
                            assert b'<html' in response.read().lower(), '首页内容不完整'
                        break
                    except URLError:
                        time.sleep(.1)
                else:
                    raise RuntimeError('桌面包启动超时')
                for route in ['/app.js', '/styles.css']:
                    with urlopen(origin + route, timeout=3) as response:
                        assert response.status == 200 and len(response.read()) > 100, f'资源缺失：{route}'
                with urlopen(origin + '/api/settings', timeout=3) as response:
                    assert isinstance(json.load(response), dict), '设置接口返回无效'
                assert (data / 'k12.sqlite3').is_file(), '数据库未写入指定目录'
                check_macro_sources(origin)
                check_conversion(origin, root)
                for image_format in ("svg", "png", "jpg", "jpeg", "gif", "bmp", "tif", "tiff"):
                    check_image_format_conversion(origin, root, image_format)
                check_image_format_conversion(origin, root, "bmp", top_down=True)
                check_excel_chart(origin, root)
                with urlopen(origin + '/api/acceptance-matrix', timeout=60) as response:
                    matrix = json.load(response)['acceptanceMatrix']
                requirements = {item['key']: item for group in matrix['groups'] for item in group['items']}
                assert requirements['17.2.4']['status'] == '部分实测', '包内图片样本被误标为完整验收通过'
                assert '图片引用和原始字节已验证' in requirements['17.2.4']['evidence'], '包内图片验收证据未更新'
                assert '17.2.4' in {item['key'] for item in matrix['uncovered_risks']}, '包内完整图片验收缺口被隐藏'
                print('桌面包客户端、资源、API、转换下载及数据目录验证通过')
            finally:
                if process.poll() is None:
                    process.terminate()
                    try:
                        process.wait(timeout=10)
                    except subprocess.TimeoutExpired:
                        process.kill()
                        process.wait(timeout=5)


if __name__ == '__main__':
    main()
