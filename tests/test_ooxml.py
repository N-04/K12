"""验证 OOXML 兼容分支只处理一种可理解的内容。"""

import unittest

from k12.ooxml import parse_compatible_xml


class CompatibilityTests(unittest.TestCase):
    """覆盖别名前缀、局部重定义、回退和无可用分支。"""

    def test_supported_choice_uses_namespace_uri(self) -> None:
        """前缀可以改名，判断依据是实际命名空间。"""
        xml = '<root xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:alias="http://schemas.openxmlformats.org/drawingml/2006/main"><mc:AlternateContent><mc:Choice Requires="alias"><text>现代</text></mc:Choice><mc:Fallback><text>旧式</text></mc:Fallback></mc:AlternateContent></root>'
        self.assertEqual([node.text for node in parse_compatible_xml(xml).iter("text")], ["现代"])

    def test_local_namespace_redefinition_uses_fallback(self) -> None:
        """局部前缀重定义不被外层已知 URI 覆盖。"""
        xml = '<root xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><mc:AlternateContent><mc:Choice xmlns:a="urn:unknown" Requires="a"><text>不可理解</text></mc:Choice><mc:Fallback><text>旧式</text></mc:Fallback></mc:AlternateContent></root>'
        self.assertEqual([node.text for node in parse_compatible_xml(xml).iter("text")], ["旧式"])

    def test_missing_supported_branch_fails(self) -> None:
        """没有可理解选择也没有回退时，明确失败而不是丢失内容。"""
        xml = '<root xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:AlternateContent><mc:Choice Requires="unknown"><text>未知</text></mc:Choice></mc:AlternateContent></root>'
        with self.assertRaisesRegex(ValueError, "没有可用分支"):
            parse_compatible_xml(xml)
