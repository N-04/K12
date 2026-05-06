import os
import requests
import json
import pypandoc
from docx import Document
import win32com.client as win32
import easyocr

# --- Mathpix 配置 (去官网获取) ---
MATHPIX_APP_ID = "你的APP_ID"
MATHPIX_APP_KEY = "你的APP_KEY"

def process_pdf_step(file_path):
    """PDF -> Mathpix -> MD -> Docx"""
    # 这一步需要联网调用 API，建议先在 Windows/Mac 环境下测试 API 通信
    print(f"正在上传 Mathpix: {file_path}")
    # (此处填入之前给你的 requests 上传逻辑)
    # 假设最后得到了 md_content
    # md_to_docx(md_content, file_path.replace(".pdf", ".docx"))
    
    options = {
    "conversion_formats": {"docx": True}, # 直接要 docx
    "math_inline_delimiters": ["$", "$"],
    "rm_spaces": True
}

# 初始化 OCR 引擎（全局初始化，避免重复加载）
# 'ch_sim' 处理简体中文，'en' 处理英文
reader = easyocr.Reader(['ch_sim', 'en'])

def process_word_step(file_path, options):
    """
    Word 文档深度处理：标点纠错 + 图片字还原
    returns: (replaced_char_count, punc_count)
    """
    doc = Document(file_path)
    char_replace_count = 0
    punc_fix_count = 0
    
    # 尺寸阈值：250,000 EMU 约为 12pt (五号字) 大小
    CHAR_THRESHOLD = 250000 

    for para in doc.paragraphs:
        # --- 1. 标点符号纠错 ---
        if options.get('punc'):
            original_text = para.text
            # 常见的全角/半角转换逻辑
            para.text = para.text.replace(',', '，').replace('?', '？').replace(':', '：')
            if para.text != original_text:
                punc_fix_count += 1

        # --- 2. 图片字识别与还原 ---
        if options.get('ocr_replace'):
            for run in para.runs:
                # 检查 Run 中是否包含图片 XML
                if 'pic:pic' in run._element.xml:
                    for inline in run._element.xpath('.//wp:inline'):
                        width = inline.extent.cx
                        height = inline.extent.cy
                        
                        # 仅处理符合单字尺寸特征的图片
                        if width < CHAR_THRESHOLD and height < CHAR_THRESHOLD:
                            try:
                                # 提取图片二进制流
                                img_id = inline.xpath('.//a:blip/@r:embed')[0]
                                image_bytes = doc.part.related_parts[img_id].blob
                                
                                # OCR 识别
                                result = reader.readtext(image_bytes)
                                
                                if result:
                                    detected_char = result[0][1]
                                    # 彻底移除图片，替换为识别出的纯文本
                                    run.clear()
                                    run.text = detected_char
                                    char_replace_count += 1
                            except Exception as e:
                                print(f"单字还原跳过: {e}")

    # 只有在确实发生改变时才保存，减少 IO 开销
    if char_replace_count > 0 or punc_fix_count > 0:
        doc.save(file_path)
    
    return char_replace_count, punc_fix_count

def run_word_vba_logic(file_path):
    try:
        # 1. 启动 Word
        word = win32.gencache.EnsureDispatch('Word.Application')
        word.Visible = False
        doc = word.Documents.Open(file_path)
        
        # 获取全文范围，代替 VBA 中的 Selection
        full_range = doc.Content

        # --- 步骤 A: 设置制表位 ---
        print("正在清除并重新设置制表位...")
        full_range.ParagraphFormat.TabStops.ClearAll()
        # 转换为 Points (1 厘米约等于 28.35 磅)
        doc.DefaultTabStop = 0.74 * 28.35 
        full_range.ParagraphFormat.TabStops.Add(Position=3.66 * 28.35, Alignment=0)
        full_range.ParagraphFormat.TabStops.Add(Position=7.33 * 28.35, Alignment=0)
        full_range.ParagraphFormat.TabStops.Add(Position=10.99 * 28.35, Alignment=0)

        # --- 步骤 B: 公式垂直居中 ---
        print("正在设置公式垂直居中...")
        # 对应 VBA 中的 .BaseLineAlignment = wdBaselineAlignCenter (3)
        full_range.ParagraphFormat.BaseLineAlignment = 3

        
        # --- 步骤 C: 设置 Word 文档格式 ---
        print("正在统一字体和行间距...")
        full_range.Font.NameFarEast = "宋体"
        full_range.Font.Name = "Times New Roman"
        full_range.Font.Size = 12
        
        full_range.ParagraphFormat.LineSpacingRule = 1 # wdLineSpace1pt5 对应的是 1.5倍行距
        full_range.ParagraphFormat.Alignment = 0        # wdAlignParagraphLeft
        full_range.ParagraphFormat.FirstLineIndent = 0
        
        # 将所有浮动图片转为嵌入式
        for shape in doc.Shapes:
            shape.ConvertToInlineShape()      

        
        # 最终保存
        doc.Save()
        doc.Close()
        print("VBA 相关逻辑处理完毕")
        
    except Exception as e:
        print(f"Windows 自动化出错: {e}")
    except ImportError:
        print("非 Windows 环境，跳过宏执行")