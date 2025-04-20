# 安装必要的Windows功能
Write-Host "Installing IIS..."
Install-WindowsFeature -Name Web-Server -IncludeManagementTools

# 创建应用目录
$appPath = "C:\ppt-to-pdf"
Write-Host "Creating application directory at $appPath"
New-Item -ItemType Directory -Force -Path $appPath
New-Item -ItemType Directory -Force -Path "$appPath\uploads"
New-Item -ItemType Directory -Force -Path "$appPath\temp"
New-Item -ItemType Directory -Force -Path "$appPath\logs"

# 复制应用文件
Write-Host "Copying application files..."
Copy-Item ".\*" -Destination $appPath -Recurse -Force

# 创建虚拟环境并安装依赖
Write-Host "Setting up Python virtual environment..."
python -m venv "$appPath\venv"
& "$appPath\venv\Scripts\Activate.ps1"
pip install -r "$appPath\requirements.txt"

# 创建Windows服务
$serviceName = "PPTtoPDFConverter"
$serviceDisplayName = "PPT to PDF Converter Service"
$serviceDescription = "Web service for converting PowerPoint files to PDF"

Write-Host "Creating Windows service..."
New-Service -Name $serviceName `
    -DisplayName $serviceDisplayName `
    -Description $serviceDescription `
    -StartupType Automatic `
    -BinaryPathName "C:\ppt-to-pdf\venv\Scripts\python.exe C:\ppt-to-pdf\app.py"

# 启动服务
Write-Host "Starting service..."
Start-Service -Name $serviceName

Write-Host "Deployment completed!" 