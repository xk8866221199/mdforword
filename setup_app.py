"""
py2app 打包配置
将 MD → Word 应用打包为 macOS .app 文件

使用方法:
    python3 setup_app.py py2app
"""
from setuptools import setup

APP = ['run_app.py']
DATA_FILES = [
    ('templates', ['templates/index.html']),
    ('static', ['static/script.js', 'static/style.css']),
]
OPTIONS = {
    'argv_emulation': False,
    'iconfile': None,  # 可替换为自定义 .icns 图标文件
    'plist': {
        'CFBundleName': 'MD to Word',
        'CFBundleDisplayName': 'MD → Word',
        'CFBundleIdentifier': 'com.mdforword.app',
        'CFBundleVersion': '1.0.0',
        'CFBundleShortVersionString': '1.0.0',
        'NSHighResolutionCapable': True,
        'LSMinimumSystemVersion': '10.15',
    },
    'packages': [
        'flask', 'jinja2', 'werkzeug', 'markupsafe', 'click',
        'docx', 'markdown_it', 'mdit_py_plugins',
        'webview',
        'converter',
    ],
    'includes': [
        'lxml', 'lxml.etree', 'lxml._elementpath',
    ],
}

setup(
    app=APP,
    name='MD to Word',
    data_files=DATA_FILES,
    options={'py2app': OPTIONS},
    setup_requires=['py2app'],
)
