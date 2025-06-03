import os
import sys
from PyInstaller.__main__ import run

# 获取当前脚本所在目录
current_dir = os.path.dirname(os.path.abspath(__file__))

# 设置图标路径
icon_path = os.path.join(current_dir, 'pdf.ico')

# 设置主程序路径
main_py = os.path.join(current_dir, 'PDF 目录提取.py')

if not os.path.exists(icon_path):
    print(f"错误：找不到图标文件 {icon_path}")
    sys.exit(1)

if not os.path.exists(main_py):
    print(f"错误：找不到主程序文件 {main_py}")
    sys.exit(1)

# PyInstaller打包命令
options = [
    main_py,
    '--name=PDF目录提取器',
    '--onefile',  # 打包成单个文件
    '--windowed',  # 不显示控制台窗口
    f'--icon={icon_path}',  # 设置应用图标
    '--noconfirm',  # 覆盖输出目录
    '--add-data', f'{icon_path};.',  # 将图标文件添加到打包文件中
    '--clean',  # 清理临时文件
    '--noupx',  # 不使用UPX压缩，提高启动速度
    '--version-file=version.txt',  # 版本信息文件
]

# 创建版本信息文件
version_info = '''
VSVersionInfo(
  ffi=FixedFileInfo(
    filevers=(1, 0, 0, 0),
    prodvers=(1, 0, 0, 0),
    mask=0x3f,
    flags=0x0,
    OS=0x40004,
    fileType=0x1,
    subtype=0x0,
    date=(0, 0)
  ),
  kids=[
    StringFileInfo(
      [
      StringTable(
        u'080404b0',
        [StringStruct(u'CompanyName', u'EzXxY'),
        StringStruct(u'FileDescription', u'PDF目录提取器'),
        StringStruct(u'FileVersion', u'1.0.0'),
        StringStruct(u'InternalName', u'PDF目录提取器'),
        StringStruct(u'LegalCopyright', u'Copyright (C) 2024'),
        StringStruct(u'OriginalFilename', u'PDF目录提取器.exe'),
        StringStruct(u'ProductName', u'PDF目录提取器'),
        StringStruct(u'ProductVersion', u'1.0.0')])
      ]
    ),
    VarFileInfo([VarStruct(u'Translation', [2052, 1200])])
  ]
)
'''

# 保存版本信息文件
with open('version.txt', 'w', encoding='utf-8') as f:
    f.write(version_info)

try:
    # 运行打包命令
    run(options)
finally:
    # 清理版本信息文件
    if os.path.exists('version.txt'):
        os.remove('version.txt') 