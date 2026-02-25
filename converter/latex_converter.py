"""
LaTeX 数学表达式转 Unicode 模块
将 Gemini 复制文本中的 LaTeX 数学表达式转换为 Unicode 字符
"""
import re


# ============================================================
# 希腊字母映射
# ============================================================
GREEK_LETTERS = {
    # 小写
    r'\alpha': 'α', r'\beta': 'β', r'\gamma': 'γ', r'\delta': 'δ',
    r'\epsilon': 'ε', r'\varepsilon': 'ε', r'\zeta': 'ζ', r'\eta': 'η',
    r'\theta': 'θ', r'\vartheta': 'ϑ', r'\iota': 'ι', r'\kappa': 'κ',
    r'\lambda': 'λ', r'\mu': 'μ', r'\nu': 'ν', r'\xi': 'ξ',
    r'\pi': 'π', r'\varpi': 'ϖ', r'\rho': 'ρ', r'\varrho': 'ϱ',
    r'\sigma': 'σ', r'\varsigma': 'ς', r'\tau': 'τ', r'\upsilon': 'υ',
    r'\phi': 'φ', r'\varphi': 'ϕ', r'\chi': 'χ', r'\psi': 'ψ',
    r'\omega': 'ω',
    # 大写
    r'\Alpha': 'Α', r'\Beta': 'Β', r'\Gamma': 'Γ', r'\Delta': 'Δ',
    r'\Epsilon': 'Ε', r'\Zeta': 'Ζ', r'\Eta': 'Η', r'\Theta': 'Θ',
    r'\Iota': 'Ι', r'\Kappa': 'Κ', r'\Lambda': 'Λ', r'\Mu': 'Μ',
    r'\Nu': 'Ν', r'\Xi': 'Ξ', r'\Pi': 'Π', r'\Rho': 'Ρ',
    r'\Sigma': 'Σ', r'\Tau': 'Τ', r'\Upsilon': 'Υ', r'\Phi': 'Φ',
    r'\Chi': 'Χ', r'\Psi': 'Ψ', r'\Omega': 'Ω',
}

# ============================================================
# 数学运算符和符号映射
# ============================================================
MATH_SYMBOLS = {
    r'\times': '×', r'\div': '÷', r'\pm': '±', r'\mp': '∓',
    r'\cdot': '·', r'\ast': '∗', r'\star': '⋆',
    r'\leq': '≤', r'\le': '≤', r'\geq': '≥', r'\ge': '≥',
    r'\neq': '≠', r'\ne': '≠', r'\approx': '≈', r'\equiv': '≡',
    r'\sim': '∼', r'\simeq': '≃', r'\cong': '≅', r'\propto': '∝',
    r'\infty': '∞', r'\partial': '∂', r'\nabla': '∇',
    r'\sum': '∑', r'\prod': '∏', r'\int': '∫',
    r'\iint': '∬', r'\iiint': '∭', r'\oint': '∮',
    r'\forall': '∀', r'\exists': '∃', r'\nexists': '∄',
    r'\in': '∈', r'\notin': '∉', r'\ni': '∋',
    r'\subset': '⊂', r'\supset': '⊃', r'\subseteq': '⊆', r'\supseteq': '⊇',
    r'\cup': '∪', r'\cap': '∩', r'\emptyset': '∅', r'\varnothing': '∅',
    r'\land': '∧', r'\lor': '∨', r'\lnot': '¬', r'\neg': '¬',
    r'\Rightarrow': '⇒', r'\Leftarrow': '⇐', r'\Leftrightarrow': '⇔',
    r'\rightarrow': '→', r'\leftarrow': '←', r'\leftrightarrow': '↔',
    r'\uparrow': '↑', r'\downarrow': '↓',
    r'\mapsto': '↦', r'\to': '→',
    r'\ldots': '…', r'\cdots': '⋯', r'\vdots': '⋮', r'\ddots': '⋱',
    r'\angle': '∠', r'\triangle': '△', r'\square': '□',
    r'\circ': '∘', r'\bullet': '•', r'\diamond': '◇',
    r'\sqrt': '√', r'\cbrt': '∛',
    r'\prime': '′', r'\degree': '°',
    r'\%': '%', r'\$': '$', r'\&': '&', r'\#': '#',
    r'\{': '{', r'\}': '}', r'\_': '_',
    r'\quad': ' ', r'\qquad': '  ',
    r'\,': ' ', r'\;': ' ', r'\!': '',
    r'\text': '',  # \text{...} 单独处理
    r'\mathrm': '', r'\mathbf': '', r'\mathit': '',
    r'\left': '', r'\right': '',
    r'\Big': '', r'\big': '', r'\Bigg': '', r'\bigg': '',
}

# ============================================================
# Unicode 上标/下标字符
# ============================================================
SUPERSCRIPT_MAP = {
    '0': '⁰', '1': '¹', '2': '²', '3': '³', '4': '⁴',
    '5': '⁵', '6': '⁶', '7': '⁷', '8': '⁸', '9': '⁹',
    '+': '⁺', '-': '⁻', '=': '⁼', '(': '⁽', ')': '⁾',
    'n': 'ⁿ', 'i': 'ⁱ', 'x': 'ˣ',
}

SUBSCRIPT_MAP = {
    '0': '₀', '1': '₁', '2': '₂', '3': '₃', '4': '₄',
    '5': '₅', '6': '₆', '7': '₇', '8': '₈', '9': '₉',
    '+': '₊', '-': '₋', '=': '₌', '(': '₍', ')': '₎',
    'a': 'ₐ', 'e': 'ₑ', 'i': 'ᵢ', 'j': 'ⱼ',
    'k': 'ₖ', 'n': 'ₙ', 'o': 'ₒ', 'p': 'ₚ',
    'r': 'ᵣ', 's': 'ₛ', 't': 'ₜ', 'u': 'ᵤ', 'x': 'ₓ',
}

# ============================================================
# 分数占位符格式（用于 Word OMML 数学公式渲染）
# ============================================================
FRAC_PLACEHOLDER = '⟦FRAC:{num}:{den}⟧'


def _convert_frac(match_str):
    """将 \\frac{a}{b} 转换为分数占位符，供 docx_builder 渲染为数学公式"""
    pattern = r'\\frac\s*\{([^{}]+)\}\s*\{([^{}]+)\}'
    m = re.match(pattern, match_str)
    if not m:
        return match_str

    numerator = m.group(1).strip()
    denominator = m.group(2).strip()

    # 递归处理分子分母中的 LaTeX（但不处理分数本身）
    numerator = _convert_latex_content(numerator)
    denominator = _convert_latex_content(denominator)

    # 使用占位符格式，让 docx_builder 渲染为 Word 原生数学分数
    return FRAC_PLACEHOLDER.format(num=numerator, den=denominator)



def _convert_sqrt(match_str):
    """将 \\sqrt{x} 转换为 √x 或 \\sqrt[n]{x} 转换为 ⁿ√x"""
    # \sqrt[n]{x}
    pattern_n = r'\\sqrt\s*\[([^\]]+)\]\s*\{([^{}]+)\}'
    m = re.match(pattern_n, match_str)
    if m:
        n = m.group(1).strip()
        content = _convert_latex_content(m.group(2).strip())
        if n == '3':
            return f"∛{content}"
        n_super = ''.join(SUPERSCRIPT_MAP.get(c, c) for c in n)
        return f"{n_super}√{content}"

    # \sqrt{x}
    pattern = r'\\sqrt\s*\{([^{}]+)\}'
    m = re.match(pattern, match_str)
    if m:
        content = _convert_latex_content(m.group(1).strip())
        return f"√{content}"

    return match_str


def _convert_superscript(content):
    """将内容转换为上标 Unicode"""
    content = _convert_latex_content(content)
    return ''.join(SUPERSCRIPT_MAP.get(c, c) for c in content)


def _convert_subscript(content):
    """将内容转换为下标 Unicode"""
    content = _convert_latex_content(content)
    return ''.join(SUBSCRIPT_MAP.get(c, c) for c in content)


def _convert_latex_content(text):
    """转换 LaTeX 内容中的符号（不包括 $ 定界符）"""
    result = text

    # 处理 \text{...}、\mathrm{...} 等文本命令
    result = re.sub(r'\\(?:text|mathrm|mathbf|mathit|textbf|textit)\s*\{([^{}]+)\}',
                    r'\1', result)

    # 处理 \frac{a}{b}
    while r'\frac' in result:
        new_result = re.sub(r'\\frac\s*\{([^{}]+)\}\s*\{([^{}]+)\}',
                            lambda m: _convert_frac(m.group(0)), result)
        if new_result == result:
            break
        result = new_result

    # 处理 \sqrt[n]{x} 和 \sqrt{x}
    while r'\sqrt' in result:
        new_result = re.sub(r'\\sqrt\s*(?:\[[^\]]+\])?\s*\{[^{}]+\}',
                            lambda m: _convert_sqrt(m.group(0)), result)
        if new_result == result:
            break
        result = new_result

    # 处理上标 ^{...} 和 ^x
    result = re.sub(r'\^\{([^{}]+)\}',
                    lambda m: _convert_superscript(m.group(1)), result)
    result = re.sub(r'\^([0-9a-zA-Z])',
                    lambda m: _convert_superscript(m.group(1)), result)

    # 处理下标 _{...} 和 _x
    result = re.sub(r'_\{([^{}]+)\}',
                    lambda m: _convert_subscript(m.group(1)), result)
    result = re.sub(r'_([0-9a-zA-Z])',
                    lambda m: _convert_subscript(m.group(1)), result)

    # 替换希腊字母（先替换长的，避免部分匹配）
    for latex, unicode_char in sorted(GREEK_LETTERS.items(), key=lambda x: -len(x[0])):
        result = result.replace(latex, unicode_char)

    # 替换数学符号（先替换长的）
    for latex, unicode_char in sorted(MATH_SYMBOLS.items(), key=lambda x: -len(x[0])):
        result = result.replace(latex, unicode_char)

    # 清理多余空格
    result = re.sub(r'\s+', ' ', result).strip()

    return result


def convert_latex_in_text(text):
    """
    将文本中所有 LaTeX 数学表达式转换为 Unicode

    处理两种形式：
    - 行内公式：$...$
    - 展示公式：$$...$$

    Args:
        text: 包含 LaTeX 表达式的文本

    Returns:
        LaTeX 表达式被替换为 Unicode 的文本
    """
    if '$' not in text:
        return text

    # 先处理 $$...$$ 展示公式
    result = re.sub(
        r'\$\$(.+?)\$\$',
        lambda m: _convert_latex_content(m.group(1)),
        text,
        flags=re.DOTALL
    )

    # 再处理 $...$ 行内公式（避免匹配已处理的 $$）
    result = re.sub(
        r'(?<!\$)\$(?!\$)(.+?)(?<!\$)\$(?!\$)',
        lambda m: _convert_latex_content(m.group(1)),
        result
    )

    return result
