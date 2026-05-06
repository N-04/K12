import sys
import os
import platform
from PySide6.QtWidgets import (QApplication, QMainWindow, QLabel, QVBoxLayout, 
                             QWidget, QMessageBox, QCheckBox, QFrame)
from PySide6.QtCore import Qt, QTimer

# 核心逻辑模块
# import core_logic 

class WordMasterApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Word & PDF 深度处理系统")
        self.resize(500, 480)
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        # 标题
        title = QLabel("K12文档自动化处理工具")
        title.setStyleSheet("font-size: 18px; font-weight: bold; margin-bottom: 10px;")
        layout.addWidget(title)

        # 功能复选框 (根据最新需求调整)
        self.cb_punc = QCheckBox("中英文标点纠错 (，？：等一键统一)")
        self.cb_punc.setChecked(True)
        
        self.cb_ocr = QCheckBox("识别并还原“图片字符” (解决文中隐藏图片字)")
        self.cb_ocr.setChecked(True)
        
        self.cb_mathpix = QCheckBox("PDF 识别")
        self.cb_mathpix.setChecked(False) # 默认关闭，按需开启

        self.cb_macros = QCheckBox("执行格式宏 (仅限 Windows & Word)") #设置制表位、文档格式以及公式居中
        self.cb_macros.setChecked(True)

        for cb in [self.cb_punc, self.cb_ocr, self.cb_mathpix, self.cb_macros]:
            layout.addWidget(cb)

        # 分割线
        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        layout.addWidget(line)

        # 拖拽区域
        self.drop_label = QLabel("\n\n将文件拖入此区域运行\n支持 .pdf / .doc / .docx\n\n")
        self.drop_label.setAlignment(Qt.AlignCenter)
        self.drop_label.setStyleSheet("""
            QLabel {
                border: 3px dashed #2196F3;
                border-radius: 15px;
                background-color: #f8fbff;
                font-size: 14px;
                color: #1976D2;
            }
        """)
        layout.addWidget(self.drop_label)

        # 状态显示
        self.status_bar = QLabel("系统准备就绪")
        self.status_bar.setStyleSheet("color: #666; font-size: 12px;")
        layout.addWidget(self.status_bar)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

        # 开启拖拽支持
        self.setAcceptDrops(True)

    def show_auto_msg(self, title, text, duration=1500):
        """自动消失的提示框"""
        msg = QMessageBox(self)
        msg.setWindowTitle(title)
        msg.setText(text)
        msg.setStandardButtons(QMessageBox.NoButton)
        QTimer.singleShot(duration, msg.close)
        msg.exec()

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.accept()
            self.drop_label.setStyleSheet("border: 3px dashed #4CAF50; background-color: #f1f8e9;")
        else:
            event.ignore()

    def dragLeaveEvent(self, event):
        self.reset_drop_style()

    def dropEvent(self, event):
        self.reset_drop_style()
        urls = event.mimeData().urls()
        files = [u.toLocalFile() for u in urls]

        for file_path in files:
            name = os.path.basename(file_path)
            self.status_bar.setText(f"处理中: {name}")
            
            # 1. 弹出启动提示
            self.show_auto_msg("任务启动", f"正在深度处理：{name}")

            # 2. 准备配置参数
            options = {
                'punc': self.cb_punc.isChecked(),
                'ocr_replace': self.cb_ocr.isChecked(),
                'mathpix': self.cb_mathpix.isChecked(),
                'run_macros': self.cb_macros.isChecked()
            }

            try:
                # --- 核心逻辑调用 ---
                # 获取处理统计结果
                replaced_chars, fixed_puncs = core_logic.process_word_step(file_path, options)
                
                # 如果在 Windows 上且开启了宏
                if platform.system() == "Windows" and options['run_macros']:
                    core_logic.run_word_vba_logic(file_path)

                # 3. 结果弹窗 (用户要求：显示替换了几处)
                result_text = f"文件 {name} 处理完成！"
                if replaced_chars > 0:
                    result_text += f"\n\n成功识别并还原图片文字：{replaced_chars} 处"
                if fixed_puncs > 0:
                    result_text += f"\n已自动纠正标点符号段落：{fixed_puncs} 处"
                
                QMessageBox.information(self, "任务报告", result_text)

            except Exception as e:
                QMessageBox.critical(self, "错误", f"处理失败：\n{str(e)}")

        self.status_bar.setText("就绪")

    def reset_drop_style(self):
        self.drop_label.setStyleSheet("""
            QLabel {
                border: 3px dashed #2196F3;
                border-radius: 15px;
                background-color: #f8fbff;
                font-size: 14px;
                color: #1976D2;
            }
        """)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = WordMasterApp()
    window.show()
    sys.exit(app.exec())