# MD → Word 转换器 (Markdown to Docx Converter)

![MD to Word](https://img.shields.io/badge/status-active-success.svg) ![Python](https://img.shields.io/badge/python-3.8+-blue.svg) ![Flask](https://img.shields.io/badge/flask-3.0-orange.svg)

一个轻量级、无需后台数据库的 Markdown 转 Word（`.docx`）文档应用。支持通过 **Web 浏览器** 或 **macOS 原生桌面客户端** 两种方式运行。

## ✨ 核心特性

- **多模式运行**: 支持作为本地 Web 服务运行，或直接启动为轻量级 Mac 原生桌面应用（基于 `pywebview`）。
- **即时预览**: 输入 Markdown 文本，右侧实时显示美观的 HTML 预览界面。
- **卓越的转换质量**:
  - 精准支持多级无序列表（自动转换为实心圆 ●、空心圆 ○、方块 ■ 等符号）及有序列表
  - 支持复杂表格解析（自动调整列宽，保证排版整齐）
  - 支持内联公式和块级公式（LaTeX 语法支持，转换为 Word 原生公式对象）
  - 支持删除线、加粗、斜体等多种文本格式
- **下载安全稳定**: 原生系统下载机制，无惧前端跨域或安全沙箱对大尺寸文件的限制。
- **零配置即用**: 无需配置数据库，克隆即运行。

## 🎯 效果演示

*提供一组 Web 和桌面应用的截图*

---

## 🚀 快速上手 (Web 版本)

如果你需要在局域网内给多人使用，或者部署在自己的服务器上，推荐使用 Web 版。

### 1. 环境准备

确保你已经安装了 Python (建议 3.8 或以上版本)。

### 2. 克隆仓库与安装依赖

```bash
git clone https://github.com/your-username/mdforword.git
cd mdforword

# 建议使用虚拟环境（可选）
# python3 -m venv venv
# source venv/bin/activate

# 安装所有依赖
pip install -r requirements.txt
```

### 3. 运行 Web 服务

```bash
python3 app.py
```

服务启动后，在浏览器中打开: [http://127.0.0.1:5001](http://127.0.0.1:5001)

---

## 🖥 桌面版使用 (macOS 专属)

如果你只是个人在本地使用，或者想拥有更「原生软件」的体验，本应用内置了桌面版入口。

桌面版无需单独开启浏览器，直接弹出独立的 App 窗口，且下载生成的文件时会直接调用系统的「保存文件」对话框，体验极佳。

### 运行开发版桌面 App

```bash
# 在项目根目录下，直接运行桌面启动脚本：
python3 run_app.py
```

### 将应用打包为双击运行的 `.app` 文件 (仅限 macOS)

你可以将代码固化成一个真正的 macOS 应用程序，存放在启动台中随时使用：

```bash
# 安装打包工具
pip install py2app

# 清理历史构建
rm -rf build dist

# 开始执行打包
python3 setup_app.py py2app

# 复制到「应用程序」文件夹
cp -R "dist/MD to Word.app" ~/Applications/
```
之后你就可以在 Launchpad 或「应用程序」文件夹里找到 **MD to Word** 并双击使用了！

---

## 📁 目录结构

```text
mdforword/
├── app.py                # Web 服务后端主入口 (Flask)
├── run_app.py            # Mac 桌面应用启动器 (pywebview)
├── setup_app.py          # py2app 桌面应用打包配置
├── requirements.txt      # Python 依赖清单
├── converter/            # 核心转换引擎模块
│   ├── docx_builder.py       # 将解析后的结构创建为 Word 文档
│   ├── md_parser.py          # 解析 Markdown 为自定义块对象
│   ├── latex_converter.py    # OmML (Word 公式) 转换器
│   └── ...
├── static/               # 前端静态资源
│   ├── script.js             # 实时预览与下载交互逻辑
│   └── style.css             # 深色/浅色极简美学 UI 样式
└── templates/            # 前端 HTML 模板
    └── index.html            # 唯一的主体页面
```

## 🛠️ 技术栈

**前端:**
- 原生 HTML5 / CSS3 (CSS Variables 主题系统)
- Vanilla JavaScript (ES6+，无框架依赖)
- Markdown-it (虽然主要使用自研正则实现特定 UI，但可随时扩展)

**后端:**
- Python 3
- [Flask](https://flask.palletsprojects.com/) (极简 Web 框架)
- [python-docx](https://python-docx.readthedocs.io/) (强大的 Word 文档生成库)
- [pywebview](https://pywebview.flowrl.com/) (负责封装 Web 界面为系统原生级桌面窗口)

## 🤝 贡献指南

1. Fork 本仓库
2. 创建您的 Feature 分支 (`git checkout -b feature/AmazingFeature`)
3. 提交您的修改 (`git commit -m 'Add some AmazingFeature'`)
4. 推送至分支 (`git push origin feature/AmazingFeature`)
5. 发起一个 Pull Request

## 📄 许可证

本项目采用 [MIT 许可证](LICENSE) - 详情请参阅 LICENSE 文件。
