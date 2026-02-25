"""
Word 文档样式定义模块
定义字体、字号、颜色等样式常量
"""
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH


# ============================================================
# 字体配置
# ============================================================
class Fonts:
    """字体名称常量"""
    # 中文字体
    CN_BODY = "微软雅黑"
    CN_HEADING = "黑体"
    CN_CODE = "等线"

    # 英文字体
    EN_BODY = "Calibri"
    EN_HEADING = "Calibri"
    EN_CODE = "Consolas"


# ============================================================
# 字号配置（单位：磅）
# ============================================================
class FontSizes:
    """各级标题和正文字号"""
    H1 = Pt(22)
    H2 = Pt(18)
    H3 = Pt(15)
    H4 = Pt(13)
    H5 = Pt(11)
    H6 = Pt(10.5)
    BODY = Pt(11)
    CODE = Pt(10)
    CODE_BLOCK = Pt(9.5)
    SMALL = Pt(9)

    HEADING_MAP = {
        1: H1, 2: H2, 3: H3,
        4: H4, 5: H5, 6: H6,
    }


# ============================================================
# 颜色配置
# ============================================================
class Colors:
    """颜色常量"""
    # 标题颜色
    HEADING = RGBColor(0x1A, 0x1A, 0x2E)

    # 正文
    BODY = RGBColor(0x33, 0x33, 0x33)

    # 代码
    CODE_TEXT = RGBColor(0xD4, 0x3F, 0x3F)       # 行内代码文字
    CODE_BG = RGBColor(0xF5, 0xF5, 0xF5)         # 代码块背景 (浅灰)

    # 代码块
    CODE_BLOCK_TEXT = RGBColor(0x2B, 0x2B, 0x2B)
    CODE_BLOCK_BG = RGBColor(0xF8, 0xF8, 0xF8)
    CODE_BLOCK_BORDER = RGBColor(0xE0, 0xE0, 0xE0)

    # 引用
    QUOTE_TEXT = RGBColor(0x55, 0x55, 0x55)
    QUOTE_BORDER = RGBColor(0xDD, 0xDD, 0xDD)

    # 链接
    LINK = RGBColor(0x06, 0x6E, 0xE0)

    # 表格
    TABLE_HEADER_BG = RGBColor(0x44, 0x72, 0xC4)
    TABLE_HEADER_TEXT = RGBColor(0xFF, 0xFF, 0xFF)
    TABLE_BORDER = RGBColor(0xBF, 0xBF, 0xBF)
    TABLE_ALT_BG = RGBColor(0xF2, 0xF2, 0xF2)


# ============================================================
# 段落间距配置
# ============================================================
class Spacing:
    """段落间距（单位：磅）"""
    HEADING_BEFORE = Pt(12)
    HEADING_AFTER = Pt(6)
    BODY_BEFORE = Pt(0)
    BODY_AFTER = Pt(6)
    CODE_BLOCK_BEFORE = Pt(6)
    CODE_BLOCK_AFTER = Pt(6)
    LIST_BEFORE = Pt(0)
    LIST_AFTER = Pt(2)

    LINE_SPACING = 1.15  # 行距倍数


# ============================================================
# 页面配置
# ============================================================
class PageLayout:
    """页面布局"""
    TOP_MARGIN = Cm(2.54)
    BOTTOM_MARGIN = Cm(2.54)
    LEFT_MARGIN = Cm(3.18)
    RIGHT_MARGIN = Cm(3.18)
