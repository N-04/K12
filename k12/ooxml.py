"""解析 OOXML 兼容分支，避免重复处理现代与旧式内容。"""

from io import StringIO
from xml.etree import ElementTree as ET

COMPATIBILITY = "{http://schemas.openxmlformats.org/markup-compatibility/2006}"
SUPPORTED_NAMESPACES = {
    "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "http://schemas.openxmlformats.org/drawingml/2006/main",
    "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "http://schemas.openxmlformats.org/presentationml/2006/main",
    "urn:schemas-microsoft-com:vml",
}


def parse_compatible_xml(xml: str | bytes) -> ET.Element:
    """按局部命名空间选择首个可理解分支，否则使用回退内容。"""
    text = xml.decode("utf-8") if isinstance(xml, bytes) else xml
    scopes, pending, choices = [], {}, {}
    parser = ET.iterparse(StringIO(text), events=("start-ns", "start", "end"))
    for event, node in parser:
        if event == "start-ns":
            pending[node[0]] = node[1]
        elif event == "start":
            scope = dict(scopes[-1]) if scopes else {}
            scope.update(pending)
            pending.clear()
            scopes.append(scope)
            if node.tag == COMPATIBILITY + "Choice":
                required = node.get("Requires", "").split()
                choices[node] = bool(required) and all(scope.get(prefix) in SUPPORTED_NAMESPACES for prefix in required)
        else:
            if node.tag == COMPATIBILITY + "AlternateContent":
                selected = next((child for child in node if child.tag == COMPATIBILITY + "Choice" and choices.get(child)), None)
                if selected is None:
                    selected = node.find(COMPATIBILITY + "Fallback")
                if selected is None:
                    raise ValueError("OOXML 兼容内容没有可用分支")
                children = list(selected)
                for child in list(node):
                    node.remove(child)
                node.extend(children)
            scopes.pop()
    return parser.root
