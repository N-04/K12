"""解析导出的 VBA 模块源文件，不执行宏或修改 Office 文档。"""

import re


def vba_macros(source: str) -> list[dict]:
    """列出标准模块中可直接调用的公开无参数 Sub，保留来源行号。"""
    if len(source) > 2_000_000 or "\0" in source:
        raise ValueError("VBA 源文件过大或不是文本")
    module = ""
    statements = []
    pending, first_line = "", 0
    for number, line in enumerate(source.lstrip("\ufeff").splitlines(), 1):
        # 清除注释与字符串内容，避免把文字中的 Sub 误认为声明。
        code, quoted, index = [], False, 0
        while index < len(line):
            character = line[index]
            if character == '"':
                if quoted and index + 1 < len(line) and line[index + 1] == '"':
                    index += 2
                    continue
                quoted = not quoted
                code.append(" ")
            elif character == "'" and not quoted:
                break
            elif not quoted:
                code.append(character)
            index += 1
        attribute = re.fullmatch(r'\s*Attribute\s+VB_Name\s*=\s*"([^"\r\n]+)"\s*', line, re.I)
        if attribute:
            if module or len(attribute[1]) > 255 or not re.fullmatch(r"[^\W\d_]\w*", attribute[1]):
                raise ValueError("VBA 模块名称无效或重复")
            module = attribute[1]
        if not pending:
            first_line = number
        text = "".join(code).strip()
        if text.endswith(" _"):
            pending += text[:-1]
            continue
        for statement in (pending + text).split(":"):
            if re.match(r"\s*Rem(?:\s|$)", statement, re.I):
                break
            statements.append((first_line, statement.strip()))
        pending = ""
    if pending:
        raise ValueError("VBA 源文件续行不完整")
    if any(re.fullmatch(r"Option\s+Private\s+Module", text, re.I) for _, text in statements):
        return []
    macros, names = [], set()
    for number, statement in statements:
        match = re.fullmatch(r"(?:(Public|Private|Friend)\s+)?(?:Static\s+)?Sub\s+([^\W\d_]\w*)\s*(?:\(\s*\))?", statement, re.I)
        if not match or (match[1] and match[1].lower() != "public"):
            continue
        name = match[2]
        if len(name) > 255:
            raise ValueError("VBA 宏名称超过 255 个字符")
        if name.casefold() in names:
            raise ValueError("VBA 模块中存在重复宏名称")
        names.add(name.casefold())
        macros.append({"macro_name": name, "module_name": module, "source_line": number,
                       "qualified_name": f"{module}.{name}" if module else name})
    return macros
