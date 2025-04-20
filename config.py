import os

# 服务器配置
SERVER_HOST = "0.0.0.0"  # 允许外部访问
SERVER_PORT = 8501

# 文件配置
UPLOAD_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'uploads')
TEMP_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'temp')
MAX_FILE_SIZE = 50 * 1024 * 1024  # 50MB

# PowerPoint配置
POWERPOINT_VISIBLE = False  # 服务器上运行时不显示PowerPoint窗口

# 安全配置
ALLOWED_EXTENSIONS = {'ppt', 'pptx'}
MAX_CONCURRENT_CONVERSIONS = 5

# 创建必要的目录
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(TEMP_DIR, exist_ok=True) 