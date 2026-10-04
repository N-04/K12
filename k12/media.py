"""读取 Word 内嵌图片并生成 PPT 原生图片形状。"""

import base64
import math
import posixpath
import re
import struct
from decimal import Decimal, ROUND_HALF_UP
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

from .ooxml import parse_compatible_xml
from .excel import excel_column_letters

IMAGE_TYPES = {"png": "image/png", "jpg": "image/jpeg", "jpeg": "image/jpeg", "gif": "image/gif",
               "bmp": "image/bmp", "tif": "image/tiff", "tiff": "image/tiff", "emf": "image/x-emf",
               "wmf": "image/x-wmf", "svg": "image/svg+xml"}
RELATION = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"


def image_locations(archive: zipfile.ZipFile) -> dict[str, list[str]]:
    """按实际引用标记 Word 正文、页眉页脚和 PPT 图片来源。"""
    locations = excel_image_locations(archive)
    names = set(archive.namelist())
    slide_pages = {}
    if {"ppt/presentation.xml", "ppt/_rels/presentation.xml.rels"} <= names:
        relations = {item.get("Id"): item for item in ET.fromstring(archive.read("ppt/_rels/presentation.xml.rels"))}
        for index, node in enumerate(ET.fromstring(archive.read("ppt/presentation.xml")).findall("{*}sldIdLst/{*}sldId"), 1):
            relation = relations.get(node.get(RELATION + "id"))
            if relation is None or relation.get("TargetMode") == "External" or not relation.get("Type", "").endswith("/slide"):
                raise ValueError("PPT 页面关系无效，无法定位图片")
            target = relation.get("Target", "")
            part = posixpath.normpath(posixpath.join("ppt", target)) if not target.startswith("/") else target.lstrip("/")
            slide_pages[part] = index
    stories = {}
    if "word/document.xml" in archive.namelist() and "word/_rels/document.xml.rels" in archive.namelist():
        relations = {item.get("Id"): item for item in ET.fromstring(archive.read("word/_rels/document.xml.rels"))}
        for node in parse_compatible_xml(archive.read("word/document.xml")).iter():
            kind = node.tag.rsplit("}", 1)[-1].removesuffix("Reference")
            relation = relations.get(node.get(RELATION + "id"))
            if kind in {"header", "footer"} and relation is not None and relation.get("TargetMode") != "External" and relation.get("Type", "").endswith("/" + kind):
                target = relation.get("Target", "")
                part = posixpath.normpath(posixpath.join("word", target)) if not target.startswith("/") else target.lstrip("/")
                if part in stories and stories[part] != kind:
                    raise ValueError("Word 同一部件同时被当作页眉和页脚引用")
                story = parse_compatible_xml(archive.read(part))
                if story.tag.rsplit("}", 1)[-1] != ("hdr" if kind == "header" else "ftr"):
                    raise ValueError("Word 页眉页脚部件类型不匹配")
                stories[part] = kind
    for part in archive.namelist():
        if not (part == "word/document.xml" or part in stories or part in slide_pages):
            continue
        if part.startswith(("word/header", "word/footer")) and part not in stories:
            continue
        folder, name = posixpath.split(part)
        relations_name = f"{folder}/_rels/{name}.rels"
        if relations_name not in archive.namelist():
            continue
        relations = {item.get("Id"): item for item in ET.fromstring(archive.read(relations_name))}
        label = "正文（页码待排版定位）" if part == "word/document.xml" else (
            "页眉" if stories.get(part) == "header" else "页脚" if stories.get(part) == "footer" else f"幻灯片第 {slide_pages[part]} 页")
        root = parse_compatible_xml(archive.read(part))
        parents = {child: parent for parent in root.iter() for child in parent}
        paragraphs = {paragraph: index for index, paragraph in enumerate(root.findall(".//{*}p"), 1)} if part.startswith("word/") else {}
        for node in root.iter():
            if node.tag.rsplit("}", 1)[-1] not in {"blip", "imagedata"}:
                continue
            reference = node.get(RELATION + "embed") or node.get(RELATION + "id")
            relation = relations.get(reference)
            if relation is None or relation.get("TargetMode") == "External" or not relation.get("Type", "").endswith("/image"):
                continue
            target = relation.get("Target", "")
            media = posixpath.normpath(posixpath.join(folder, target)) if not target.startswith("/") else target.lstrip("/")
            location = f"{label} · {part}"
            paragraph = node
            while paragraph in parents and paragraph not in paragraphs:
                paragraph = parents[paragraph]
            if paragraph in paragraphs:
                location += f" · 第 {paragraphs[paragraph]} 段"
            if location not in locations.setdefault(media, []):
                locations[media].append(location)
    return locations


def excel_image_locations(archive: zipfile.ZipFile) -> dict[str, list[str]]:
    """沿工作簿、工作表和绘图关系定位图片的单元格锚点。"""
    names = set(archive.namelist())
    if "xl/workbook.xml" not in names:
        return {}

    def targets(part: str, kind: str) -> dict[str, str]:
        """只解析同包内部的指定关系类型。"""
        folder, name = posixpath.split(part)
        relations = f"{folder}/_rels/{name}.rels"
        if relations not in names:
            return {}
        result = {}
        for relation in ET.fromstring(archive.read(relations)):
            if relation.get("TargetMode") == "External" or not relation.get("Type", "").endswith("/" + kind):
                continue
            target = relation.get("Target", "")
            result[relation.get("Id")] = posixpath.normpath(posixpath.join(folder, target)) if not target.startswith("/") else target.lstrip("/")
        return result

    locations = {}
    sheets = targets("xl/workbook.xml", "worksheet")
    for sheet in ET.fromstring(archive.read("xl/workbook.xml")).findall(".//{*}sheet"):
        part = sheets.get(sheet.get(RELATION + "id"))
        if not part:
            continue
        drawings = targets(part, "drawing")
        for drawing in ET.fromstring(archive.read(part)).findall(".//{*}drawing"):
            drawing_part = drawings.get(drawing.get(RELATION + "id"))
            if not drawing_part:
                continue
            images = targets(drawing_part, "image")
            for anchor in parse_compatible_xml(archive.read(drawing_part)):
                position = anchor.find("{*}from")
                cell = "绝对位置（无单元格锚点）"
                if position is not None:
                    row = int(position.findtext("{*}row", "-1"))
                    column = int(position.findtext("{*}col", "-1"))
                    if not 0 <= row < 1048576 or not 0 <= column < 16384:
                        raise ValueError("Excel 图片单元格锚点超出范围")
                    cell = f"{excel_column_letters(column + 1)}{row + 1}"
                for picture in anchor.findall(".//{*}pic"):
                    blip = picture.find(".//{*}blip")
                    media = images.get(blip.get(RELATION + "embed")) if blip is not None else None
                    if media:
                        location = f"工作表 {sheet.get('name', '')} · {cell}"
                        if location not in locations.setdefault(media, []):
                            locations[media].append(location)
    return locations


def word_story_images(path: Path, document: ET.Element) -> list[dict]:
    """读取实际引用的页眉页脚图片，同一部件只提取一次。"""
    references = [node for node in document.iter() if node.tag.rsplit("}", 1)[-1] in {"headerReference", "footerReference"}]
    if not references:
        return []
    images, seen = [], set()
    with zipfile.ZipFile(path) as archive:
        relations = ET.fromstring(archive.read("word/_rels/document.xml.rels"))
        by_id = {item.get("Id"): item for item in relations}
        for reference in references:
            kind = "header" if reference.tag.endswith("headerReference") else "footer"
            relation = by_id.get(reference.get(RELATION + "id"))
            if relation is None or relation.get("TargetMode") == "External" or not relation.get("Type", "").endswith("/" + kind):
                raise ValueError("Word 页眉页脚引用无效")
            target = relation.get("Target", "")
            part = posixpath.normpath(posixpath.join("word", target)) if not target.startswith("/") else target.lstrip("/")
            if not part.startswith("word/") or "\\" in part or not part.endswith(".xml"):
                raise ValueError("Word 页眉页脚部件路径无效")
            story = parse_compatible_xml(archive.read(part))
            if story.tag.rsplit("}", 1)[-1] != ("hdr" if kind == "header" else "ftr"):
                raise ValueError("Word 页眉页脚部件类型不匹配")
            if part in seen:
                continue
            seen.add(part)
            for image in word_images(path, story, part):
                images.append({"text": "", "level": 0, "image": image,
                               "image_label": "页眉图片" if kind == "header" else "页脚图片"})
    return images


def _vml_length(value: str) -> int:
    """把旧式图片的物理长度换算为 EMU，不猜测百分比尺寸。"""
    match = re.fullmatch(r"\s*([0-9]+(?:\.[0-9]+)?)\s*(in|cm|mm|pt|pc|px)\s*", value.lower())
    if not match:
        return 0
    factors = {"in": 914400, "cm": 360000, "mm": 36000, "pt": 12700, "pc": 152400, "px": 9525}
    return int(Decimal(match[1]) * factors[match[2]])


def _raster_dimensions(data: bytes) -> tuple[int, int]:
    """仅从常见光栅图片头读取尺寸，不解码或重采样图片。"""
    if data.startswith(b"\x89PNG\r\n\x1a\n") and len(data) >= 24:
        return struct.unpack(">II", data[16:24])
    if data[:6] in {b"GIF87a", b"GIF89a"} and len(data) >= 10:
        return struct.unpack("<HH", data[6:10])
    if data.startswith(b"BM") and len(data) >= 26:
        size = struct.unpack("<I", data[14:18])[0]
        if size == 12:
            return struct.unpack("<HH", data[18:22])
        if size >= 40:
            width, height = struct.unpack("<ii", data[18:26])
            return width, abs(height)
    if data.startswith(b"\xff\xd8"):
        position = 2
        while position + 4 <= len(data):
            if data[position] != 255:
                break
            while position < len(data) and data[position] == 255:
                position += 1
            if position >= len(data):
                break
            marker = data[position]
            position += 1
            if marker in {0xD9, 0xDA}:
                break
            length = int.from_bytes(data[position:position + 2], "big")
            if length < 2 or position + length > len(data):
                break
            if marker in {0xC0, 0xC1, 0xC2, 0xC3, 0xC5, 0xC6, 0xC7, 0xC9, 0xCA, 0xCB, 0xCD, 0xCE, 0xCF} and length >= 7:
                height, width = struct.unpack(">HH", data[position + 3:position + 7])
                return width, height
            position += length
    return 0, 0


def _vml_crop(value: str) -> int:
    """把小数、百分比和 65536 分制裁剪值换成 DrawingML 比例。"""
    value = value.strip()
    if not re.fullmatch(r"[+-]?(?:\d+(?:\.\d*)?|\.\d+)%?|[+-]?\d+f", value):
        raise ValueError("旧式图片裁剪值无效")
    suffix = value[-1:]
    fraction = Decimal(value[:-1]) / (65536 if suffix == "f" else 100) if suffix in {"f", "%"} else Decimal(value)
    if not -1 <= fraction <= 1:
        raise ValueError("旧式图片裁剪值超出范围")
    return int((fraction * 100000).to_integral_value(rounding=ROUND_HALF_UP))


def _vml_transform(style: dict[str, str]) -> dict:
    """读取旧式图片的角度和翻转方向，统一到 DrawingML 变换。"""
    transform = {}
    if "rotation" in style:
        angle = style["rotation"].strip().lower()
        if not re.fullmatch(r"[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:deg)?", angle):
            raise ValueError("旧式图片旋转角度无效")
        degrees = Decimal(angle.removesuffix("deg")) % 360
        transform["rot"] = int((degrees * 60000).to_integral_value(rounding=ROUND_HALF_UP)) % 21600000
    if "flip" in style:
        direction = "".join(style["flip"].lower().split())
        if direction not in {"x", "y", "xy", "yx"}:
            raise ValueError("旧式图片翻转方向无效")
        transform.update({key: "1" for axis, key in (("x", "flipH"), ("y", "flipV")) if axis in direction})
    return transform


def word_images(path: Path, element: ET.Element, source_part: str = "word/document.xml") -> list[dict]:
    """读取显式内嵌图片引用；拒绝外部引用、缺失部件与路径逃逸。"""
    references = [node for node in element.iter() if node.tag.rsplit("}", 1)[-1] in {"blip", "imagedata"}]
    if not references:
        return []
    parents = {child: parent for parent in element.iter() for child in parent}
    images = []
    with zipfile.ZipFile(path) as archive:
        folder, name = posixpath.split(source_part)
        relations = ET.fromstring(archive.read(f"{folder}/_rels/{name}.rels"))
        by_id = {item.get("Id"): item for item in relations}
        for node in references:
            reference = node.get(RELATION + "embed") or node.get(RELATION + "link") or node.get(RELATION + "id")
            shape = node
            while shape in parents and shape.tag.rsplit("}", 1)[-1] not in {"pic", "shape", "inline", "anchor"}:
                shape = parents[shape]
            extent = shape.find(".//{*}xfrm/{*}ext")
            if extent is None:
                extent = shape.find(".//{*}extent")
            width, height = (int(extent.get("cx", "0")), int(extent.get("cy", "0"))) if extent is not None else (0, 0)
            style = {key.strip().lower(): value.strip() for part in shape.get("style", "").split(";")
                     if ":" in part for key, value in [part.split(":", 1)]}
            if extent is None:
                width, height = _vml_length(style.get("width", "")), _vml_length(style.get("height", ""))
            crop_node = shape.find(".//{*}srcRect")
            crop = {side: int(crop_node.get(side, "0")) for side in ("l", "t", "r", "b")} if crop_node is not None else {}
            if node.tag.rsplit("}", 1)[-1] == "imagedata":
                crop = {side: _vml_crop(node.get(attribute)) for side, attribute in
                        (("l", "cropleft"), ("t", "croptop"), ("r", "cropright"), ("b", "cropbottom"))
                        if node.get(attribute) is not None}
            transform_node = shape.find(".//{*}xfrm")
            transform = {}
            if transform_node is not None:
                transform["rot"] = int(transform_node.get("rot", "0")) % 21600000
                transform.update({key: "1" for key in ("flipH", "flipV") if transform_node.get(key) in {"1", "true"}})
            elif node.tag.rsplit("}", 1)[-1] == "imagedata":
                transform = _vml_transform(style)
            relation = by_id.get(reference)
            if relation is None or relation.get("TargetMode") == "External":
                raise ValueError("文档图片引用缺失或为外部链接，无法保留")
            if not relation.get("Type", "").endswith("/image"):
                raise ValueError("文档图片关系类型无效")
            target = relation.get("Target", "")
            part = posixpath.normpath(posixpath.join(folder, target)) if not target.startswith("/") else target.lstrip("/")
            if not part.startswith(source_part.split("/")[0] + "/media/") or "\\" in part:
                raise ValueError("文档图片部件路径无效")
            extension = part.rsplit(".", 1)[-1].lower()
            if extension not in IMAGE_TYPES:
                raise ValueError("文档图片格式暂不支持")
            if archive.getinfo(part).file_size > 20 * 1024 * 1024:
                raise ValueError("文档单张图片超过 20 MB，无法保留")
            data = archive.read(part)
            if not data:
                raise ValueError("文档图片部件为空")
            if width <= 0 or height <= 0:
                intrinsic_width, intrinsic_height = _raster_dimensions(data)
                if intrinsic_width <= 0 or intrinsic_height <= 0:
                    raise ValueError("文档图片缺少可用显示尺寸")
                if width > 0:
                    height = max(1, width * intrinsic_height // intrinsic_width)
                elif height > 0:
                    width = max(1, height * intrinsic_width // intrinsic_height)
                else:
                    width, height = intrinsic_width * 9525, intrinsic_height * 9525
            images.append({"data": base64.b64encode(data).decode("ascii"), "extension": extension,
                           "width": width, "height": height, "crop": crop, "transform": transform})
    return images


def picture_xml(image: dict) -> str:
    """按源显示比例缩放到图片页，生成带关系引用的原生形状。"""
    width, height = int(image.get("width") or 8229600), int(image.get("height") or 4937760)
    if width <= 0 or height <= 0:
        raise ValueError("文档图片显示尺寸无效")
    transform = image.get("transform") or {}
    angle = math.radians(int(transform.get("rot", 0)) / 60000)
    bounding_width = abs(width * math.cos(angle)) + abs(height * math.sin(angle))
    bounding_height = abs(width * math.sin(angle)) + abs(height * math.cos(angle))
    scale = min(8229600 / bounding_width, 4937760 / bounding_height)
    crop = {key: int(value) for key, value in (image.get("crop") or {}).items() if key in {"l", "t", "r", "b"}}
    if crop.get("l", 0) + crop.get("r", 0) >= 100000 or crop.get("t", 0) + crop.get("b", 0) >= 100000:
        raise ValueError("文档图片裁剪范围没有可见区域")
    crop_xml = '<a:srcRect ' + " ".join(f'{key}="{value}"' for key, value in crop.items()) + '/>' if crop else ""
    transform_xml = " ".join(f'{key}="{int(transform[key])}"' for key in ("rot", "flipH", "flipV") if key in transform)
    width, height = max(1, int(width * scale)), max(1, int(height * scale))
    x, y = 457200 + (8229600 - width) // 2, 1463040 + (4937760 - height) // 2
    return f'<p:pic><p:nvPicPr><p:cNvPr id="5" name="Picture"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr><p:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="image"/>{crop_xml}<a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm {transform_xml}><a:off x="{x}" y="{y}"/><a:ext cx="{width}" cy="{height}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>'


def word_picture_xml(image: dict, index: int) -> str:
    """把内嵌图片写为 Word 行内绘图，复用裁剪与变换属性。"""
    picture = picture_xml(image).replace("<p:", "<pic:").replace("</p:", "</pic:").replace('r:embed="image"', f'r:embed="image{index}"')
    # 使用源显示尺寸，正文宽度和高度不足时才缩小。
    root = ET.fromstring(f'<root xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">{picture}</root>')
    extent = root.find(".//{*}xfrm/{*}ext")
    width, height = int(image["width"]), int(image["height"])
    angle = math.radians(int((image.get("transform") or {}).get("rot", 0)) / 60000)
    cosine, sine = abs(math.cos(angle)), abs(math.sin(angle))
    bounding_width = width * cosine + height * sine
    bounding_height = width * sine + height * cosine
    scale = min(1, 5715000 / max(width, bounding_width), 8863330 / max(height, bounding_height))
    width, height = max(1, int(width * scale)), max(1, int(height * scale))
    # 为旋转后的外边界预留行高和左右空间，避免遮挡相邻正文。
    horizontal = max(0, math.ceil((width * cosine + height * sine - width) / 2 - 1e-6))
    vertical = max(0, math.ceil((width * sine + height * cosine - height) / 2 - 1e-6))
    extent.set("cx", str(width))
    extent.set("cy", str(height))
    offset = root.find(".//{*}xfrm/{*}off")
    offset.set("x", "0")
    offset.set("y", "0")
    properties = root.find(".//{*}nvPicPr")
    extra = properties.find("{*}nvPr")
    if extra is not None:
        properties.remove(extra)
    picture = ET.tostring(root[0], encoding="unicode")
    return f'<w:p><w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><wp:extent cx="{width}" cy="{height}"/><wp:effectExtent l="{horizontal}" r="{horizontal}" t="{vertical}" b="{vertical}"/><wp:docPr id="{index}" name="图片 {index}"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">{picture}</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>'
