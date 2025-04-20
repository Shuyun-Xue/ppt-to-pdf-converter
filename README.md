# PPT to PDF Converter

一个简单的在线 PPT 转 PDF 工具，使用纯 Python 实现，可以在线将 PowerPoint 文件转换为 PDF 格式。

## 功能特点

- 支持 .ppt 和 .pptx 格式文件
- 在线转换，无需安装软件
- 简单易用的网页界面
- 快速下载转换后的 PDF 文件

## 技术栈

- Python 3.6+
- Streamlit
- python-pptx
- FPDF2
- Pillow

## 本地运行

1. 克隆仓库：
```bash
git clone https://github.com/Shuyun-Xue/ppt-to-pdf-converter.git
cd ppt-to-pdf-converter
```

2. 安装依赖：
```bash
pip install -r requirements.txt
```

3. 运行应用：
```bash
streamlit run app.py
```

## 在线使用

访问：[https://ppt-to-pdf-converter.streamlit.app](https://ppt-to-pdf-converter.streamlit.app)

## 注意事项

- 当前版本支持基本的文本和形状转换
- 复杂的动画效果和某些特殊格式可能无法完全保留
- 建议在转换前备份重要文件

## License

MIT License 