import docx
from docx.shared import Inches
import json
import os
from reportlab.lib.pagesizes import A4, landscape
from reportlab.pdfgen import canvas

def extract_word_content(doc_path):
    """提取Word文档内容（文本和图片）"""
    doc = docx.Document(doc_path)
    content = []
    
    # 创建临时目录保存图片
    if not os.path.exists('temp_images'):
        os.makedirs('temp_images')
    
    # 提取段落文本
    for para in doc.paragraphs:
        if para.text.strip():
            content.append({
                'type': 'text',
                'content': para.text,
                'style': {
                    'font_size': para.style.font.size.pt if para.style.font.size else None,
                    'bold': para.style.font.bold,
                    'italic': para.style.font.italic
                }
            })
    
    # 提取图片
    for rel in doc.part.rels.values():
        if "image" in rel.target_ref:
            img_data = rel.target_part.blob
            img_ext = rel.target_ref.split('.')[-1]
            img_filename = f'temp_images/image_{len(content)}.{img_ext}'
            with open(img_filename, 'wb') as f:
                f.write(img_data)
            content.append({
                'type': 'image',
                'path': img_filename,
                'width': Inches(4)  # 默认图片宽度
            })
    
    return content

def generate_booklet_pdf(content, output_path, is_landscape=False):
    """生成对折小册子PDF"""
    # 小册子尺寸：210×140mm (A4的1/4)
    page_width, page_height = 210, 140  # mm
    if is_landscape:
        page_width, page_height = page_height, page_width  # 横版交换宽高
    
    # 转换为点 (1mm = 2.83465点)
    page_width_pt = page_width * 2.83465
    page_height_pt = page_height * 2.83465
    
    # 创建A4尺寸的canvas（对折打印需要A4纸）
    c = canvas.Canvas(output_path, pagesize=landscape(A4) if is_landscape else A4)
    a4_width, a4_height = landscape(A4) if is_landscape else A4
    
    # 计算每张A4纸可容纳的小册子页数（4页）
    # 这里简化处理，实际需要根据内容分页
    # 目前仅演示前4页内容排列
    for i in range(0, len(content), 4):
        # 第1面：第4页（左上）和第1页（右上）
        c.drawString(50, a4_height - 50, f"Page {i+4}: {content[i]['content'][:20]}..." if i+3 < len(content) else "Empty")
        c.drawString(a4_width/2 + 50, a4_height - 50, f"Page {i+1}: {content[i]['content'][:20]}..." if i < len(content) else "Empty")
        
        # 第2面：第2页（左上）和第3页（右上）
        c.showPage()
        c.drawString(50, a4_height - 50, f"Page {i+2}: {content[i+1]['content'][:20]}..." if i+1 < len(content) else "Empty")
        c.drawString(a4_width/2 + 50, a4_height - 50, f"Page {i+3}: {content[i+2]['content'][:20]}..." if i+2 < len(content) else "Empty")
        c.showPage()
    
    c.save()
    print(f"拼版PDF已生成: {output_path}")

if __name__ == "__main__":
    # 示例调用（实际应从web应用接收文件路径）
    doc_path = "example.docx"  # 实际应用中替换为上传文件路径
    output_pdf = "booklet_output.pdf"
    
    try:
        content = extract_word_content(doc_path)
        generate_booklet_pdf(content, output_pdf, is_landscape=False)  # 默认为竖版
        print("小册子拼版处理完成")
    except Exception as e:
        print(f"处理失败: {str(e)}")