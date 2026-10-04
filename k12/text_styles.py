"""解析 Word 正文的显式格式与样式继承。"""

import re
import posixpath
from urllib.parse import urlsplit
from xml.etree import ElementTree as ET

from .ooxml import parse_compatible_xml

WORD = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
UNDERLINES = {
    "none": "none", "single": "sng", "words": "words",
    "double": "dbl", "thick": "heavy", "dotted": "dotted",
    "dottedHeavy": "dottedHeavy", "dash": "dash", "dashedHeavy": "dashHeavy",
    "dashLong": "dashLong", "dashLongHeavy": "dashLongHeavy", "dotDash": "dotDash",
    "dashDotHeavy": "dotDashHeavy", "dotDotDash": "dotDotDash", "dashDotDotHeavy": "dotDotDashHeavy",
    "wave": "wavy", "wavyHeavy": "wavyHeavy", "wavyDouble": "wavyDbl",
}


def theme_language_script(language: str) -> str | None:
    """映射常见东亚语言与显式 BCP 47 文字系统，未知语言不猜测。"""
    parts = language.lower().split("-")
    for part in parts[1:]:
        if len(part) == 4 and part.isalpha():
            return part.title()
    if parts[0] == "zh":
        if any(part in {"tw", "hk", "mo"} for part in parts):
            return "Hant"
        if any(part in {"cn", "sg"} for part in parts):
            return "Hans"
        return None
    return {"ja": "Jpan", "ko": "Hang", "en": "Latn"}.get(parts[0])


def apply_run_properties(style: dict, properties: ET.Element | None, toggle: bool = False) -> None:
    """样式中的加粗和斜体采用切换语义，直接格式覆盖最终值。"""
    if properties is not None:
        for key in ("b", "i"):
            flag = properties.find("{*}" + key)
            if flag is not None:
                enabled = flag.get(WORD + "val", "1") not in {"0", "false", "off"}
                if not toggle or enabled:
                    style[key] = not style.get(key, False) if toggle else enabled
        value_key = WORD
        underline = properties.find("{*}u")
        if underline is not None:
            value = underline.get(WORD + "val", "single")
            if value not in UNDERLINES:
                raise ValueError("Word 下划线类型无效")
            style["underline"] = UNDERLINES[value]
            style.pop("underline_color", None)
            style.pop("underline_color_unresolved", None)
            underline_color = underline.get(WORD + "color", "auto")
            if underline_color != "auto":
                if not re.fullmatch(r"[0-9a-fA-F]{6}", underline_color):
                    raise ValueError("Word 下划线颜色无效")
                style["underline_color"] = underline_color.upper()
            if any(underline.get(WORD + attribute) for attribute in ("themeColor", "themeTint", "themeShade")):
                style["underline_color_unresolved"] = True
                style["underline_color_theme"] = {attribute: underline.get(WORD + attribute) for attribute in ("themeColor", "themeTint", "themeShade")}
            else:
                style.pop("underline_color_theme", None)
        color = properties.find("{*}color")
        if color is not None:
            style.pop("color", None)
            style.pop("theme_color_unresolved", None)
            if any(color.get(value_key + attribute) for attribute in ("themeColor", "themeTint", "themeShade")):
                style["theme_color_unresolved"] = True
                style["color_theme"] = {attribute: color.get(WORD + attribute) for attribute in ("themeColor", "themeTint", "themeShade")}
            else:
                style.pop("color_theme", None)
            color_value = color.get(value_key + "val", "auto")
            if color_value != "auto":
                if not re.fullmatch(r"[0-9a-fA-F]{6}", color_value):
                    raise ValueError("Word 文字颜色无效")
                style["color"] = color_value.upper()
        size = properties.find("{*}sz")
        if size is not None:
            size_value = int(size.get(value_key + "val", "0"))
            if not 1 <= size_value <= 3276:
                raise ValueError("Word 文字字号超出可转换范围")
            style["size"] = size_value * 50
        fonts = properties.find("{*}rFonts")
        if fonts is not None:
            for key, attribute in (("font", "ascii"), ("east_asia", "eastAsia")):
                theme = fonts.get(value_key + attribute + "Theme")
                if theme:
                    style[key + "_theme"] = theme
                    style.pop(key, None)
                if fonts.get(value_key + attribute):
                    style[key] = fonts.get(value_key + attribute)
                    if not theme:
                        style.pop(key + "_theme", None)


class WordTextStyles:
    """依次应用文档默认值、段落样式、字符样式和直接格式。"""

    def __init__(self, root: ET.Element | None, theme: ET.Element | None = None, settings: ET.Element | None = None):
        """建立样式索引，并读取文档默认文字格式。"""
        self.styles = {} if root is None else {node.get(WORD + "styleId"): node for node in root.findall("{*}style")}
        self.theme = theme
        self.color_mapping = settings.find("{*}clrSchemeMapping") if settings is not None else None
        self.theme_languages = settings.find("{*}themeFontLang") if settings is not None else None
        self.defaults = {}
        if root is not None:
            apply_run_properties(self.defaults, root.find("{*}docDefaults/{*}rPrDefault/{*}rPr"))
        self.default_ids = {node.get(WORD + "type"): key for key, node in self.styles.items()
                            if node.get(WORD + "default") in {"1", "true", "on"}}

    @classmethod
    def from_archive(cls, archive):
        """从文档关系读取主题，支持自定义主题文件名。"""
        styles = parse_compatible_xml(archive.read("word/styles.xml")) if "word/styles.xml" in archive.namelist() else None
        parts = {}
        if "word/_rels/document.xml.rels" in archive.namelist():
            relations = ET.fromstring(archive.read("word/_rels/document.xml.rels"))
            for relation in relations:
                kind = relation.get("Type", "").rsplit("/", 1)[-1]
                if kind not in {"theme", "settings"}:
                    continue
                uri = urlsplit(relation.get("Target", ""))
                target = posixpath.normpath(uri.path.lstrip("/") if uri.path.startswith("/") else "word/" + uri.path)
                if relation.get("TargetMode") == "External" or uri.scheme or uri.netloc or uri.query or target.startswith("../"):
                    raise ValueError("Word 主题或设置关系无效")
                parts[kind] = parse_compatible_xml(archive.read(target))
        return cls(styles, parts.get("theme"), parts.get("settings"))

    def theme_font(self, reference: str) -> str | None:
        """按主题语言选择补充字体，缺失时读取区域字体定义。"""
        match = re.fullmatch(r"(major|minor)(Ascii|HAnsi|EastAsia)", reference)
        if self.theme is None or not match:
            return None
        group, category = match.groups()
        collection = self.theme.find("{*}themeElements/{*}fontScheme/{*}" + group + "Font")
        if collection is None:
            return None
        language = self.theme_languages.get(WORD + ("eastAsia" if category == "EastAsia" else "val"), "") if self.theme_languages is not None else ""
        script = theme_language_script(language)
        if script:
            for font in collection.findall("{*}font"):
                if font.get("script") == script and font.get("typeface"):
                    return font.get("typeface")
        node = collection.find("{*}" + ("ea" if category == "EastAsia" else "latin"))
        return (node.get("typeface") or None) if node is not None else None

    def theme_color(self, reference: dict) -> str | None:
        """解析 RGB 主题色及系统色回退，未验证的明暗调整保持限制提示。"""
        if self.theme is None:
            return None
        name = reference.get("themeColor")
        roles = {"text1": ("t1", "dark1"), "text2": ("t2", "dark2"), "background1": ("bg1", "light1"), "background2": ("bg2", "light2")}
        if name in roles:
            role, default = roles[name]
            name = self.color_mapping.get(WORD + role, default) if self.color_mapping is not None else default
        name = {"dark1": "dk1", "dark2": "dk2", "light1": "lt1", "light2": "lt2", "hyperlink": "hlink", "followedHyperlink": "folHlink"}.get(name, name)
        scheme = self.theme.find("{*}themeElements/{*}clrScheme")
        node = next((child for child in scheme if child.tag.rsplit("}", 1)[-1] == name), None) if scheme is not None else None
        if node is None or len(node) != 1 or len(node[0]):
            return None
        color = node[0]
        kind = color.tag.rsplit("}", 1)[-1]
        value = color.get("val") if kind == "srgbClr" else color.get("lastClr") if kind == "sysClr" else None
        if not value or not re.fullmatch(r"[0-9a-fA-F]{6}", value):
            return None
        if reference.get("themeTint") or reference.get("themeShade"):
            return None
        return value.upper()

    def chain(self, style_id: str | None, kind: str) -> list[ET.Element]:
        """沿 basedOn 读取同类样式；拒绝循环，忽略缺失或异类父样式。"""
        chain, visited = [], set()
        while style_id in self.styles:
            if style_id in visited:
                raise ValueError("Word 样式继承存在循环")
            visited.add(style_id)
            node = self.styles[style_id]
            if node.get(WORD + "type") != kind:
                break
            chain.append(node)
            parent = node.find("{*}basedOn")
            style_id = parent.get(WORD + "val") if parent is not None else None
        return list(reversed(chain))

    def run(self, paragraph: ET.Element, run: ET.Element) -> dict:
        """未指定样式时使用默认样式，显式字体属性按字段继承。"""
        result = self.defaults.copy()
        for kind, reference in (("paragraph", paragraph.find("{*}pPr/{*}pStyle")),
                                ("character", run.find("{*}rPr/{*}rStyle"))):
            style_id = reference.get(WORD + "val") if reference is not None else self.default_ids.get(kind)
            for node in self.chain(style_id, kind):
                apply_run_properties(result, node.find("{*}rPr"), toggle=True)
        apply_run_properties(result, run.find("{*}rPr"))
        for key, warning in (("color", "theme_color_unresolved"), ("underline_color", "underline_color_unresolved")):
            reference = result.pop(key + "_theme", None)
            if reference:
                color = self.theme_color(reference)
                if color:
                    result[key] = color
                    result.pop(warning, None)
        for key in ("font", "east_asia"):
            reference = result.pop(key + "_theme", None)
            if reference:
                font = self.theme_font(reference)
                if font:
                    result[key] = font
                else:
                    result["theme_font_unresolved"] = True
        return result
