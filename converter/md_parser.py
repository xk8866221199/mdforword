"""
Markdown 解析模块
使用 markdown-it-py 将 Markdown 文本解析为 Token 流
"""
from markdown_it import MarkdownIt
from mdit_py_plugins.front_matter import front_matter_plugin


def create_parser():
    """创建并返回配置好的 Markdown 解析器"""
    md = MarkdownIt("commonmark", {"breaks": True, "html": False})
    # 启用表格扩展
    md.enable("table")
    # 启用删除线
    md.enable("strikethrough")
    # 前置元数据插件
    front_matter_plugin(md)
    return md


def parse_markdown(text: str) -> list:
    """
    解析 Markdown 文本，返回 Token 列表

    Args:
        text: Markdown 格式的文本

    Returns:
        Token 列表
    """
    md = create_parser()
    tokens = md.parse(text)
    return tokens
