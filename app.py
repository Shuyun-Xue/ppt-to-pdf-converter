import streamlit as st
import os
import tempfile
from pptx import Presentation
from fpdf import FPDF
from PIL import Image, ImageDraw
import io

def render_shape(shape, draw, offset_x=0, offset_y=0):
    """渲染PPT中的形状"""
    if hasattr(shape, 'text'):
        # 渲染文本
        text_frame = shape.text_frame
        if text_frame.text:
            draw.text((offset_x + shape.left, offset_y + shape.top), 
                     text_frame.text, 
                     fill='black')
    
    if hasattr(shape, 'fill'):
        # 渲染形状
        if shape.shape_type == 1:  # 矩形
            draw.rectangle(
                [offset_x + shape.left, offset_y + shape.top,
                 offset_x + shape.left + shape.width,
                 offset_y + shape.top + shape.height],
                outline='black'
            )

def convert_slide_to_image(slide):
    """将PPT幻灯片转换为图像"""
    # 获取幻灯片尺寸（转换EMU到像素）
    width = int(slide.shapes.width * 0.75)
    height = int(slide.shapes.height * 0.75)
    
    # 创建一个新的图像
    img = Image.new('RGB', (width, height), 'white')
    draw = ImageDraw.Draw(img)
    
    # 渲染所有形状
    for shape in slide.shapes:
        render_shape(shape, draw)
    
    return img

def convert_ppt_to_pdf(input_file_path):
    """
    将PPT转换为PDF（纯Python实现）
    """
    try:
        # 创建临时目录存放图片
        with tempfile.TemporaryDirectory() as temp_dir:
            # 加载PPT文件
            prs = Presentation(input_file_path)
            
            # 创建PDF文档
            pdf = FPDF()
            
            # 设置PDF页面大小为PPT大小
            first_slide = prs.slides[0] if prs.slides else None
            if first_slide:
                width = first_slide.shapes.width * 0.75 / 96 * 25.4  # 转换为毫米
                height = first_slide.shapes.height * 0.75 / 96 * 25.4
                pdf.set_page_size((width, height))
            
            # 遍历所有幻灯片
            for i, slide in enumerate(prs.slides):
                # 将幻灯片转换为图像
                img = convert_slide_to_image(slide)
                
                # 保存图像
                img_path = os.path.join(temp_dir, f'slide_{i}.png')
                img.save(img_path, 'PNG')
                
                # 添加到PDF
                pdf.add_page()
                pdf.image(img_path, x=0, y=0, w=pdf.w, h=pdf.h)
            
            # 保存PDF
            output_path = os.path.splitext(input_file_path)[0] + '.pdf'
            pdf.output(output_path)
            
            return output_path
            
    except Exception as e:
        st.error(f"转换出错: {str(e)}")
        return None

def main():
    st.set_page_config(
        page_title="PPT转PDF工具",
        page_icon="📄",
        layout="centered"
    )
    
    st.title("PPT转PDF工具")
    st.write("上传PPT文件，自动转换为PDF格式")
    
    # 添加使用说明
    st.info("""
    使用说明：
    1. 点击"选择PPT文件"上传您的PPT文件（支持.ppt和.pptx格式）
    2. 点击"转换为PDF"按钮开始转换
    3. 转换完成后，点击"下载PDF文件"保存结果
    
    注意：当前版本支持基本的文本和形状转换，复杂的动画效果和某些特殊格式可能无法完全保留。
    """)
    
    uploaded_file = st.file_uploader("选择PPT文件", type=['ppt', 'pptx'])
    
    if uploaded_file is not None:
        # 创建临时文件来保存上传的PPT
        with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(uploaded_file.name)[1]) as tmp_file:
            tmp_file.write(uploaded_file.getvalue())
            input_path = tmp_file.name
        
        if st.button("转换为PDF"):
            with st.spinner('正在转换中...'):
                # 转换文件
                pdf_path = convert_ppt_to_pdf(input_path)
                
                if pdf_path and os.path.exists(pdf_path):
                    # 读取生成的PDF文件
                    with open(pdf_path, 'rb') as pdf_file:
                        pdf_data = pdf_file.read()
                    
                    # 提供下载链接
                    st.success("转换成功！")
                    st.download_button(
                        label="下载PDF文件",
                        data=pdf_data,
                        file_name=os.path.splitext(uploaded_file.name)[0] + '.pdf',
                        mime='application/pdf'
                    )
                    
                    # 清理临时文件
                    try:
                        os.remove(input_path)
                        os.remove(pdf_path)
                    except:
                        pass
                else:
                    st.error("转换失败，请重试")

if __name__ == '__main__':
    main() 