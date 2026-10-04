"""验证真实 VBA 源文本中的可调用宏识别边界。"""

import unittest
import tempfile

from k12.vba_sources import vba_macros


class VbaSourceTests(unittest.TestCase):
    def test_word_bridge_checks_ownership_before_saving(self) -> None:
        """Word 桥接只在核对本次源文件后取得可关闭的文档引用。"""
        from k12.local_client import _macos_office_normalize_script
        script = _macos_office_normalize_script("word")
        self.assertLess(script.index("Source document is already open"), script.index("open sourceFile"))
        self.assertLess(script.index("Opened document ownership mismatch"), script.index("set documentRef to active document"))
        self.assertIn("if documentRef is not missing value then", script)

    def test_macro_queue_rejects_removed_or_changed_source(self) -> None:
        """真实宏队列不能把失效来源降级为自由名称宏。"""
        from k12.store import AppStore
        from k12.processor import TaskProcessor
        with tempfile.TemporaryDirectory() as temporary:
            store = AppStore(temporary)
            processor = TaskProcessor(store)
            imported = processor.import_macro_source({"file_name": "test.bas", "source": "Public Sub Prepare()\nEnd Sub"})
            task = {"id": "queue", "options": {"selectedMacros": [{"id": imported["id"] + "_0"}]}}
            definition = store.get_macro_source(imported["id"])
            definition["source"] += "\n' changed"
            store.save_macro_source(definition)
            with self.assertRaisesRegex(ValueError, "不一致"):
                processor._macro_items({"id": "word"}, task)
            store.delete_macro_source(imported["id"])
            with self.assertRaisesRegex(ValueError, "已删除"):
                processor._macro_items({"id": "word"}, task)

    def test_source_delete_preserves_queued_tasks_and_templates(self) -> None:
        """有任务或模板引用时拒绝删除，解除引用后删除真实来源。"""
        from k12.store import AppStore
        from k12.processor import TaskProcessor
        with tempfile.TemporaryDirectory() as temporary:
            store = AppStore(temporary)
            processor = TaskProcessor(store)
            source = processor.import_macro_source({"file_name": "test.bas", "source": "Public Sub Prepare()\nEnd Sub"})
            macro_id = source["id"] + "_0"
            task = {"id": "pending", "status": "待确认", "options": {"selectedMacros": [{"id": macro_id}]}}
            store.save_task(task)
            with self.assertRaisesRegex(ValueError, "待执行任务"):
                processor.delete_macro_source(source["id"])
            task["status"] = "成功"
            store.save_task(task)
            template = processor.save_macro_template({"name": "测试", "macro_sequence": [{"id": macro_id}]})
            with self.assertRaisesRegex(ValueError, "模板引用"):
                processor.delete_macro_source(source["id"])
            processor.delete_macro_template(template["id"])
            self.assertTrue(processor.delete_macro_source(source["id"]))
            self.assertEqual(store.list_macro_sources(), [])
            self.assertEqual(len(processor.macro_library()), 5)

    def test_import_http_and_source_metadata_handoff(self) -> None:
        """实际导入接口返回声明，客户端宏动作保留来源哈希与行号。"""
        import json
        import threading
        from pathlib import Path
        from urllib.request import Request, urlopen
        from k12.server import K12Server, K12RequestHandler
        with tempfile.TemporaryDirectory() as temporary:
            server = K12Server(("127.0.0.1", 0), K12RequestHandler, Path(temporary))
            worker = threading.Thread(target=server.serve_forever, daemon=True)
            worker.start()
            try:
                origin = f"http://127.0.0.1:{server.server_address[1]}"
                payload = {"file_name": "lesson.bas", "source": 'Attribute VB_Name = "Lesson"\nPublic Sub Prepare()\nEnd Sub'}
                request = Request(origin + "/api/macros/import", data=json.dumps(payload).encode(), headers={"Content-Type": "application/json"})
                with urlopen(request, timeout=5) as response:
                    self.assertEqual(response.status, 201)
                    source = json.load(response)["source"]
                self.assertNotIn("source", source)
                with urlopen(origin + "/api/macros/sources/" + source["id"], timeout=5) as response:
                    definition = json.load(response)["source"]
                self.assertEqual(definition["source"], payload["source"])
                self.assertEqual(definition["sha256"], source["sha256"])
                with urlopen(origin + "/api/macros", timeout=5) as response:
                    macro = next(item for item in json.load(response)["macros"] if item.get("source_id") == source["id"])
                macro["execute_status"] = "待本地客户端执行"
                actions = server.processor._local_payload_actions({"task_type": "macro_sequence"}, [], {"analysis": {"macros": [macro]}})
                queued = next(item for item in actions if item["type"] == "macro_sequence")["macros"][0]
                self.assertEqual(queued["source_id"], source["id"])
                self.assertEqual(queued["source_sha256"], source["sha256"])
                self.assertEqual(queued["source_line"], 2)
                self.assertEqual(queued["module_name"], "Lesson")
                self.assertNotIn("source", queued)
            finally:
                server.shutdown()
                server.server_close()
                worker.join(timeout=5)

    def test_import_persists_and_enters_macro_library(self) -> None:
        """导入声明进入真实宏库，重启后保存，源正文不进入公开返回值。"""
        from k12.store import AppStore
        from k12.processor import TaskProcessor
        with tempfile.TemporaryDirectory() as temporary:
            processor = TaskProcessor(AppStore(temporary))
            payload = {"file_name": "lesson.bas", "source": 'Attribute VB_Name = "Lesson"\nPublic Sub Prepare()\nEnd Sub'}
            imported = processor.import_macro_source(payload)
            self.assertNotIn("source", imported)
            self.assertEqual(processor.import_macro_source(payload)["id"], imported["id"])
            restored = TaskProcessor(AppStore(temporary))
            macros = [item for item in restored.macro_library() if item.get("source_id") == imported["id"]]
            self.assertEqual(len(macros), 1)
            self.assertEqual(macros[0]["macro_name"], "Lesson.Prepare")
            self.assertTrue(macros[0]["definition_available"])
            self.assertNotIn("source", macros[0])
            self.assertEqual(restored.store.list_macro_sources()[0]["source"], payload["source"])
            definition = restored.store.get_macro_source(imported["id"])
            definition["source"] += "\n' changed"
            restored.store.save_macro_source(definition)
            with self.assertRaisesRegex(ValueError, "不一致"):
                restored.macro_source_definition(imported["id"])
            with self.assertRaisesRegex(ValueError, "2 MB"):
                restored.import_macro_source({"file_name": "large.bas", "source": "中" * 700_000})

    def test_public_subs_and_source_lines(self) -> None:
        """公开无参过程可调用，带参数、私有过程和函数不作为宏。"""
        source = 'Attribute VB_Name = "教案模块"\nPublic Sub 整理标题()\nEnd Sub\nPrivate Sub Hidden()\nEnd Sub\nSub NeedsArgument(value As String)\nEnd Sub\nPublic Function Value() As String\nEnd Function\nSub Refresh\nEnd Sub'
        self.assertEqual(vba_macros(source), [
            {"macro_name": "整理标题", "module_name": "教案模块", "source_line": 2, "qualified_name": "教案模块.整理标题"},
            {"macro_name": "Refresh", "module_name": "教案模块", "source_line": 10, "qualified_name": "教案模块.Refresh"},
        ])

    def test_comments_strings_and_continuation(self) -> None:
        """注释和字符串中的伪声明被排除，续行和同一行声明可识别。"""
        source = "' Sub Fake()\nRem comment: Sub FakeTwo()\nvalue = \"Sub FakeThree(): Public Sub FakeFour()\"\nPublic _\nSub Real(): End Sub"
        self.assertEqual([item["macro_name"] for item in vba_macros(source)], ["Real"])
        self.assertEqual(vba_macros(source)[0]["source_line"], 4)

    def test_private_module_and_invalid_source(self) -> None:
        """私有模块不公开宏，重复名称及二进制源明确失败。"""
        self.assertEqual(vba_macros("Option Private Module\nPublic Sub Hidden()"), [])
        for source in ("Sub Same()\nSub SAME()", "Public _", "\0"):
            with self.assertRaises(ValueError):
                vba_macros(source)

    def test_invalid_or_duplicate_module_names(self) -> None:
        """限定调用名只由合法模块标识组成，拒绝重复模块声明。"""
        for module in ("bad.name", "bad name", "123Module", "Module[1]", "_Module", "M" * 256):
            with self.assertRaisesRegex(ValueError, "模块名称"):
                vba_macros(f'Attribute VB_Name = "{module}"\nPublic Sub Prepare()')
        with self.assertRaisesRegex(ValueError, "重复"):
            vba_macros('Attribute VB_Name = "First"\nAttribute VB_Name = "Second"\nSub Prepare()')
        self.assertEqual(vba_macros('\ufeffAttribute VB_Name = "Module_1"\nSub Prepare()')[0]["qualified_name"], "Module_1.Prepare")
        self.assertEqual(vba_macros("Sub _Hidden()"), [])
        with self.assertRaisesRegex(ValueError, "255"):
            vba_macros("Sub " + "M" * 256 + "()")
