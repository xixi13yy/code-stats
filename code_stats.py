#!/usr/bin/env python3
"""代码量统计工具（表格版 + 图形界面）。

在 cloc.py 的行数统计内核之上补齐三件事：

1. 按「语言 / 源文件数 / 代码行数 / 空行数 / 注释行数 / 函数个数 / 最大值 / 最小值 / 均值 / 中位数」
   的表格格式输出（与 testcase3_测试结果 的列一致）；
2. 统计函数个数，并给出单个函数代码行数的最大 / 最小 / 均值 / 中位数；
3. 图形界面：输入路径、选择统计方式，点一下就出表，可导出 CSV 或直接复制进 Excel。

用法：
    python code_stats.py                      # 打开图形界面
    python code_stats.py <路径>               # 命令行直接出表
    python code_stats.py <路径> -m hand       # 按 .clocignore 排除生成内容
    python code_stats.py <路径> --csv out.csv # 同时导出 CSV

命令行参数：
    path                 目标目录或文件；不给则打开图形界面
    -m, --mode           raw=原始统计 / hand=手写口径(读 .clocignore) / ext=仅指定扩展名
    -e, --ext            仅统计这些扩展名，逗号分隔（mode=ext 时使用）
    -d, --distribution   func=按函数代码行 / file=按文件代码行（默认 func）
    --no-others          只输出 C / C++ / JAVA / Python / C# 五行
    --csv               把表格另存为 CSV（utf-8-sig，Excel 可直接打开）
"""
from __future__ import annotations

import argparse
import ast
import csv
import os
import queue
import re
import statistics
import sys
import threading
import time
import unicodedata
from dataclasses import dataclass, field
from pathlib import Path

# Windows 控制台默认 GBK，直接输出中文会乱码
if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:  # 某些环境 stdout 不可重配置
        pass

# 允许 code_stats.py 被拷到别处时依然能找到 cloc.py
sys.path.insert(0, str(Path(__file__).resolve().parent))
import cloc  # noqa: E402  同目录的行数统计内核（语言表 / 读文件 / 数行 / 忽略规则）

# ---------------------------------------------------------------- 表格定义

#: 表格列，顺序与 testcase3_测试结果 完全一致
COLUMNS = ("语言", "源文件数", "代码行数", "空行数", "注释行数",
           "函数个数", "最大值", "最小值", "均值", "中位数")

#: 固定的五个语言行（顺序同图片）
TEMPLATE_LANGS = ("C", "C++", "JAVA", "Python", "C#")

#: cloc 的语言名 -> 表格里显示的名字（.h/.hpp 归到 C++）
LANG_DISPLAY = {
    "C": "C",
    "C++": "C++",
    "C/C++ Header": "C++",
    "C++ Header": "C++",
    "Java": "JAVA",
    "Python": "Python",
    "C#": "C#",
    "JavaScript": "JavaScript",
    "TypeScript": "TypeScript",
    "SQL": "SQL",
    "HTML": "HTML",
    "CSS": "CSS",
    "Markdown": "Markdown",
    "JSON": "JSON",
    "Shell": "Shell",
    "Go": "Go",
    "Rust": "Rust",
}

#: 花括号语言（用大括号界定函数体）
BRACE_LANGS = {
    "C", "C++", "C/C++ Header", "C++ Header", "Java", "JavaScript",
    "TypeScript", "C#", "Go", "Rust", "PHP", "Swift", "Kotlin", "Scala", "Dart",
}

#: 出现在“名字(”之前时说明不是函数定义，而是调用/表达式
NO_FUNC_PREV = {
    "new", "return", "throw", "case", "delete", "typeof", "sizeof", "await",
    "yield", "else", "do", "defined", "assert", "static_assert", "decltype",
    "alignof", "noexcept", "not", "and", "or", "in", "is", "lambda", "co_await",
    "co_return", "co_yield",
}

#: 不是函数名的关键字（`if (`、`for (` 这类控制结构）
KEYWORDS = {
    "if", "for", "while", "switch", "catch", "do", "else", "try", "finally",
    "return", "sizeof", "new", "delete", "throw", "case", "default", "assert",
    "defined", "typeof", "await", "yield", "with", "foreach", "lock", "using",
    "synchronized", "defer", "go", "select", "match", "when", "unless", "until",
    "operator", "function", "static_assert", "decltype", "alignof", "noexcept",
    "co_await", "co_return", "co_yield", "super", "this", "print", "println",
}


# ---------------------------------------------------------------- 函数识别

def extract_functions(text: str, lang: str) -> list[tuple[str, int, int]]:
    """识别源码里的函数，返回 [(函数名, 起始行, 结束行), ...]（行号从 1 开始）。

    - Python 用 ast 精确解析；语法错误时退回缩进启发式
    - 花括号语言用“名字 + 参数表 + 函数体”的启发式扫描
    - SQL 只认 CREATE FUNCTION / PROCEDURE
    - 其余语言（HTML/CSS/Markdown/JSON…）不统计函数
    """
    if lang == "Python":
        return _python_functions(text)
    if lang in BRACE_LANGS:
        return _brace_functions(text)
    if lang == "SQL":
        return _sql_functions(text)
    return []


def _python_functions(text: str) -> list[tuple[str, int, int]]:
    """用 ast 精确统计 def / async def（含类方法、嵌套函数）。"""
    try:
        tree = ast.parse(text)
    except (SyntaxError, ValueError):
        return _python_functions_regex(text)
    out: list[tuple[str, int, int]] = []
    for node in ast.walk(tree):
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
            start = node.lineno
            end = getattr(node, "end_lineno", None) or start
            out.append((node.name, start, end))
    out.sort(key=lambda item: item[1])
    return out


def _python_functions_regex(text: str) -> list[tuple[str, int, int]]:
    """语法解析失败时的兜底：按缩进找 def 块。"""
    lines = text.splitlines()
    pat = re.compile(r"^(\s*)(?:async\s+)?def\s+([A-Za-z_]\w*)\s*\(")
    out: list[tuple[str, int, int]] = []
    for i, line in enumerate(lines):
        m = pat.match(line)
        if not m:
            continue
        indent = len(m.group(1))
        start = i + 1
        end = start
        for j in range(i + 1, len(lines)):
            cur = lines[j]
            if not cur.strip():
                continue
            cur_indent = len(cur) - len(cur.lstrip())
            if cur_indent <= indent:
                break
            end = j + 1
        out.append((m.group(2), start, end))
    return out


def _sql_functions(text: str) -> list[tuple[str, int, int]]:
    """统计 CREATE FUNCTION / CREATE PROCEDURE。"""
    pattern = re.compile(
        r"\bCREATE\s+(?:OR\s+REPLACE\s+)?(?:DEFINER\s*=\s*\S+\s+)?"
        r"(FUNCTION|PROCEDURE)\s+([\w.`]+)",
        re.I,
    )
    out = []
    for m in pattern.finditer(text):
        line = text.count("\n", 0, m.start()) + 1
        out.append((m.group(2), line, line))
    return out


def _mask(text: str) -> str:
    """把注释与字符串内容替换成空格（保留换行），便于后续做括号匹配。"""
    out = list(text)
    n = len(text)
    i = 0
    while i < n:
        ch = text[i]
        if ch == "/" and i + 1 < n and text[i + 1] == "/":
            j = text.find("\n", i)
            j = n if j == -1 else j
            for k in range(i, j):
                out[k] = " "
            i = j
            continue
        if ch == "/" and i + 1 < n and text[i + 1] == "*":
            j = text.find("*/", i + 2)
            j = n if j == -1 else j + 2
            for k in range(i, j):
                if out[k] != "\n":
                    out[k] = " "
            i = j
            continue
        if ch in "\"'`":
            quote = ch
            j = i + 1
            while j < n:
                cj = text[j]
                if cj == "\\":
                    j += 2
                    continue
                if cj == quote:
                    j += 1
                    break
                if cj == "\n" and quote != "`":  # 普通字符串不跨行
                    break
                j += 1
            for k in range(i, min(j, n)):
                if out[k] != "\n":
                    out[k] = " "
            i = j
            continue
        i += 1
    return "".join(out)


def _match_pair(masked: str, pos: int, open_ch: str, close_ch: str) -> int | None:
    """从 pos（指向 open_ch）出发，返回匹配的 close_ch 下标。"""
    depth = 0
    for i in range(pos, len(masked)):
        c = masked[i]
        if c == open_ch:
            depth += 1
        elif c == close_ch:
            depth -= 1
            if depth == 0:
                return i
    return None


#: 函数名 + 可选模板参数 + 参数表，例如 `foo(`、`Get<T>(`
_NAME_CALL = re.compile(r"([A-Za-z_$][\w$]*)\s*(?:<[^<>(){};]{1,120}>\s*)?\(")
#: 箭头函数 `const f = (a) => ...`
_ARROW = re.compile(
    r"([A-Za-z_$][\w$]*)\s*=\s*(?:async\s+)?"
    r"(?:\([^()]*\)|[A-Za-z_$][\w$]*)\s*=>"
)


#: 参数表之后允许出现的修饰词（其余词一律说明这不是函数定义）
_FUNC_QUALIFIERS = {
    "const", "volatile", "mutable", "noexcept", "override", "final", "requires",
}


def _brace_functions(text: str) -> list[tuple[str, int, int]]:
    """花括号语言的函数识别：名字 + 参数表 + 函数体，靠大括号配对定位函数体。"""
    masked = _mask(text)
    n = len(masked)
    found: list[tuple[str, int, int]] = []
    seen: set[int] = set()
    #: 构造器初始化列表等“不是函数”的区域，稍后遇到的匹配要跳过
    skip_ranges: list[tuple[int, int]] = []

    def add(name: str, start_pos: int, end_pos: int) -> None:
        if start_pos in seen or end_pos < start_pos:
            return
        seen.add(start_pos)
        start_line = text.count("\n", 0, start_pos) + 1
        end_line = text.count("\n", 0, end_pos) + 1
        found.append((name, start_line, end_line))

    for m in _NAME_CALL.finditer(masked):
        name = m.group(1)
        if name in KEYWORDS:
            continue
        name_pos = m.start(1)
        if any(a <= name_pos < b for a, b in skip_ranges):
            continue

        before = masked[max(0, name_pos - 60):name_pos]
        prev_word_match = re.search(r"([\w$]+)\s*$", before)
        prev_word = prev_word_match.group(1) if prev_word_match else ""
        if prev_word in NO_FUNC_PREV:
            continue
        stripped = before.rstrip()
        if stripped.endswith((".", "->")):  # obj.foo(...) 这类调用
            continue

        p_open = m.end() - 1
        p_close = _match_pair(masked, p_open, "(", ")")
        if p_close is None:
            continue

        # 参数表之后只接受：修饰词 / 引用符 / 初始化列表 / throws 子句 / { / =>
        q = p_close + 1
        limit = min(n, p_close + 600)
        body = None
        arrow = False
        while q < limit:
            while q < limit and masked[q] in " \t\r\n":
                q += 1
            if q >= limit:
                break
            c = masked[q]
            if c == "{":
                body = q
                break
            if masked.startswith("=>", q):
                arrow = True
                q += 2
                break
            if c == ":":
                # C++ 构造器初始化列表 / C# 基类构造调用：登记为跳过区
                r = q + 1
                while r < limit and masked[r] not in "{;":
                    r += 1
                if r < limit and masked[r] == "{":
                    skip_ranges.append((q, r))
                    body = r
                break
            if c in "&*[":
                q += 1
                continue
            if c.isalnum() or c == "_":
                word_match = re.match(r"[\w]+", masked[q:])
                word = word_match.group(0) if word_match else ""
                if word in _FUNC_QUALIFIERS:
                    q += len(word)
                    if q < limit and masked[q] == "(":  # noexcept(...)
                        end = _match_pair(masked, q, "(", ")")
                        if end is None:
                            break
                        q = end + 1
                    continue
                if word == "throws":  # Java 异常声明
                    q += len(word)
                    while q < limit and (masked[q].isalnum() or masked[q] in "_.,<> \t\r\n"):
                        q += 1
                    continue
                if word == "where":  # C# 泛型约束
                    while q < limit and masked[q] not in "{;":
                        q += 1
                    continue
                break  # 出现别的词（如 `> 0`），说明这是表达式不是定义
            if c == "-" and masked.startswith("->", q):  # 尾置返回类型
                q += 2
                while q < limit and (masked[q].isalnum() or masked[q] in "_<>:,*& \t\r\n"):
                    q += 1
                continue
            break

        if body is not None:
            end = _match_pair(masked, body, "{", "}")
            if end is None:
                continue
            add(name, name_pos, end)
        elif arrow:
            semi = masked.find(";", q)
            if semi == -1:
                semi = n - 1
            add(name, name_pos, semi)

    # 箭头函数：`const f = (a, b) => { ... }` / `const f = a => a + 1`
    for m in _ARROW.finditer(masked):
        name = m.group(1)
        name_pos = m.start(1)
        q = m.end()
        while q < len(masked) and masked[q] in " \t\r\n":
            q += 1
        if q < len(masked) and masked[q] == "{":
            end = _match_pair(masked, q, "{", "}")
            if end is None:
                continue
            add(name, name_pos, end)
        else:
            semi = masked.find(";", q)
            if semi == -1:
                continue
            add(name, name_pos, semi)

    found.sort(key=lambda item: item[1])
    return found


# ---------------------------------------------------------------- 统计

@dataclass
class LangStat:
    """单个语言的累计量。"""
    files: int = 0
    blank: int = 0
    comment: int = 0
    code: int = 0
    funcs: int = 0
    file_codes: list[int] = field(default_factory=list)
    func_codes: list[int] = field(default_factory=list)


@dataclass
class FileRow:
    """明细里的一行。"""
    path: str
    lang: str
    blank: int
    comment: int
    code: int
    funcs: int


@dataclass
class Result:
    """一次统计的完整结果。"""
    rows: list[tuple] = field(default_factory=list)
    files: list[FileRow] = field(default_factory=list)
    languages: dict[str, LangStat] = field(default_factory=dict)
    scanned: int = 0
    skipped_ignored: int = 0
    skipped_dup: int = 0
    skipped_binary: int = 0
    elapsed: float = 0.0
    distribution: str = "func"
    ignore_file: str = ""
    root: str = ""


def _fmt_num(value) -> str:
    """数字格式化：整数不带小数点，其余保留 1 位。"""
    if value is None:
        return "-"
    if abs(value - round(value)) < 1e-9:
        return str(int(round(value)))
    return f"{value:.1f}"


def _describe(values: list[int]) -> tuple[str, str, str, str]:
    """返回 (最大值, 最小值, 均值, 中位数)；空集全部给 '-'。"""
    if not values:
        return "-", "-", "-", "-"
    return (
        _fmt_num(max(values)),
        _fmt_num(min(values)),
        _fmt_num(statistics.fmean(values)),
        _fmt_num(statistics.median(values)),
    )


def _to_row(lang: str, stat: LangStat, distribution: str) -> tuple:
    values = stat.func_codes if distribution == "func" else stat.file_codes
    maximum, minimum, mean, median = _describe(values)
    return (
        lang, stat.files, stat.code, stat.blank, stat.comment,
        stat.funcs, maximum, minimum, mean, median,
    )


def analyze(
    root: str | Path,
    mode: str = "raw",
    ext: str | None = None,
    distribution: str = "func",
    append_others: bool = True,
    ignore_name: str = ".clocignore",
    extra_exclude: str | None = None,
) -> Result:
    """统计 root 下的代码量。

    mode:
        raw  —— 原始统计，不排除任何东西（除了缓存目录）
        hand —— 手写口径，读取根目录的 .clocignore 排除生成内容
        ext  —— 只统计 ext 指定的扩展名
    distribution:
        func —— 最大值/最小值/均值/中位数 按“单个函数的代码行数”算
        file —— 按“单个文件的代码行数”算
    """
    start = time.perf_counter()
    target = Path(root)
    if not target.exists():
        raise FileNotFoundError(f"路径不存在：{target}")

    only_ext = None
    if mode == "ext":
        if not ext:
            raise ValueError("统计方式为“仅指定扩展名”时必须填写扩展名，例如 .py,.js")
        only_ext = {
            e if e.startswith(".") else "." + e
            for e in (x.strip().lower() for x in ext.split(","))
            if e
        }

    # 忽略清单
    patterns: list[str] = []
    ignore_path = ""
    if mode == "hand":
        base = target if target.is_dir() else target.parent
        candidate = base / ignore_name
        if candidate.is_file():
            patterns = cloc.load_ignore(candidate)
            ignore_path = str(candidate)

    excludes = set(cloc.DEFAULT_EXCLUDE_DIRS)
    if extra_exclude:
        excludes |= {d.strip() for d in extra_exclude.split(",") if d.strip()}

    root_for_rel = target if target.is_dir() else target.parent
    stats: dict[str, LangStat] = {}
    file_rows: list[FileRow] = []
    seen: dict[str, str] = {}
    skipped_ignored = skipped_dup = skipped_binary = 0
    scanned = 0

    def handle(path: Path) -> None:
        nonlocal skipped_ignored, skipped_dup, skipped_binary, scanned
        info = cloc.lang_of(path)
        if info is None:
            return
        lang_name, line_comments, block_comments = info
        if only_ext is not None and path.suffix.lower() not in only_ext:
            return
        try:
            rel = path.relative_to(root_for_rel).as_posix()
        except ValueError:
            rel = path.name
        if cloc.is_ignored(rel, path.name, patterns):
            skipped_ignored += 1
            return
        text, digest = cloc.read_source(path)
        if text is None:
            skipped_binary += 1
            return
        if digest in seen:  # 内容完全相同的文件只算一份
            skipped_dup += 1
            return
        seen[digest] = str(path)
        scanned += 1

        blank, comment, code = cloc.count_text(text, line_comments, block_comments)
        funcs = extract_functions(text, lang_name)

        st = stats.setdefault(lang_name, LangStat())
        st.files += 1
        st.blank += blank
        st.comment += comment
        st.code += code
        st.funcs += len(funcs)
        st.file_codes.append(code)
        st.func_codes.extend(end - start_line + 1 for _, start_line, end in funcs)
        file_rows.append(FileRow(str(path), lang_name, blank, comment, code, len(funcs)))

    if target.is_file():
        handle(target)
    else:
        for dirpath, dirnames, filenames in os.walk(target):
            dirnames[:] = sorted(
                d for d in dirnames if d not in excludes and not d.startswith(".")
            )
            for fn in sorted(filenames):
                handle(Path(dirpath) / fn)

    # cloc 语言名归并成表格里的语言名（.h 与 .cpp 都算 C++）
    merged: dict[str, LangStat] = {}
    for lang_name, st in stats.items():
        display = LANG_DISPLAY.get(lang_name, lang_name)
        m = merged.setdefault(display, LangStat())
        m.files += st.files
        m.blank += st.blank
        m.comment += st.comment
        m.code += st.code
        m.funcs += st.funcs
        m.file_codes.extend(st.file_codes)
        m.func_codes.extend(st.func_codes)

    rows: list[tuple] = []
    for lang in TEMPLATE_LANGS:
        rows.append(_to_row(lang, merged.get(lang, LangStat()), distribution))
    if append_others:
        for lang, st in sorted(merged.items(), key=lambda kv: -kv[1].code):
            if lang in TEMPLATE_LANGS:
                continue
            rows.append(_to_row(lang, st, distribution))

    return Result(
        rows=rows,
        files=file_rows,
        languages=merged,
        scanned=scanned,
        skipped_ignored=skipped_ignored,
        skipped_dup=skipped_dup,
        skipped_binary=skipped_binary,
        elapsed=time.perf_counter() - start,
        distribution=distribution,
        ignore_file=ignore_path,
        root=str(target),
    )


# ---------------------------------------------------------------- 输出

def _width(text: str) -> int:
    """字符串显示宽度（中文算 2）。"""
    return sum(2 if unicodedata.east_asian_width(c) in "WF" else 1 for c in str(text))


def _pad(text, width: int, align: str = "left") -> str:
    text = str(text)
    gap = max(0, width - _width(text))
    return text + " " * gap if align == "left" else " " * gap + text


def format_table(result: Result) -> str:
    """渲染成等宽字符表格。"""
    cells = [[str(c) for c in row] for row in result.rows]
    widths = [max(_width(COLUMNS[i]), *(_width(r[i]) for r in cells)) + 2 for i in range(len(COLUMNS))]
    head = "".join(_pad(COLUMNS[i], widths[i], "left" if i == 0 else "right") for i in range(len(COLUMNS)))
    lines = [head, "-" * _width(head)]
    for row in cells:
        lines.append("".join(
            _pad(row[i], widths[i], "left" if i == 0 else "right")
            for i in range(len(COLUMNS))
        ))
    unit = "单个函数代码行数" if result.distribution == "func" else "单个文件代码行数"
    lines.append("")
    lines.append(f"统计目录：{result.root}")
    if result.ignore_file:
        lines.append(f"忽略清单：{result.ignore_file}")
    lines.append(f"已统计 {result.scanned} 个文件；最大值/最小值/均值/中位数 = {unit}的分布")
    if result.skipped_ignored:
        lines.append(f"按清单跳过 {result.skipped_ignored} 个文件")
    if result.skipped_dup:
        lines.append(f"内容重复跳过 {result.skipped_dup} 个文件")
    if result.skipped_binary:
        lines.append(f"二进制/无法解码跳过 {result.skipped_binary} 个文件")
    lines.append(f"耗时 {result.elapsed:.3f} 秒")
    return "\n".join(lines)


def to_tsv(result: Result) -> str:
    """制表符分隔，可直接粘进 Excel / Word 表格。"""
    lines = ["\t".join(COLUMNS)]
    for row in result.rows:
        lines.append("\t".join(str(c) for c in row))
    return "\n".join(lines)


def save_csv(result: Result, path: str | Path) -> Path:
    """导出 CSV。utf-8-sig 保证 Excel 打开中文不乱码。"""
    out = Path(path)
    with out.open("w", newline="", encoding="utf-8-sig") as fh:
        writer = csv.writer(fh)
        writer.writerow(COLUMNS)
        writer.writerows(result.rows)
    return out


# ---------------------------------------------------------------- 图形界面

MODE_LABELS = {
    "原始统计（全部文件）": "raw",
    "手写口径（按 .clocignore 排除生成内容）": "hand",
    "仅指定扩展名": "ext",
}
DIST_LABELS = {
    "按函数代码行数": "func",
    "按文件代码行数": "file",
}


def _enable_dpi_awareness() -> None:
    """让窗口在高 DPI 屏幕上清晰显示（否则 Windows 会把整个窗口模糊拉伸）。"""
    if sys.platform != "win32":
        return
    try:
        import ctypes
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)  # 系统级 DPI 感知
        except Exception:
            ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass


def launch_gui(smoke_path: str | None = None) -> int:
    """打开图形界面。

    smoke_path 只用于自动化冒烟测试：打开窗口、自动跑一次统计、打印结果后自动关闭。
    """
    _enable_dpi_awareness()
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    class App(tk.Tk):
        def __init__(self) -> None:
            super().__init__()
            self.title("代码量统计工具")
            self.geometry("1160x640")
            self.minsize(900, 500)
            self.result: Result | None = None
            self._queue: queue.Queue = queue.Queue()
            self._smoke_deadline = 0.0

            self.path_var = tk.StringVar(value=str(Path.cwd()))
            self.mode_var = tk.StringVar(value=list(MODE_LABELS)[0])
            self.dist_var = tk.StringVar(value=list(DIST_LABELS)[0])
            self.ext_var = tk.StringVar(value=".py,.js,.java,.c,.cpp,.cs")
            self.others_var = tk.BooleanVar(value=False)  # 默认只出五行，贴合测试用例表
            self.status_var = tk.StringVar(value="选择目录后点「开始统计」")

            self._build()

            if smoke_path:
                self.after(300, lambda: self._smoke_start(smoke_path))

        # ---------------- 布局
        def _build(self) -> None:
            pad = {"padx": 6, "pady": 4}
            top = ttk.Frame(self)
            top.pack(fill="x", **pad)

            ttk.Label(top, text="目标路径").grid(row=0, column=0, sticky="w")
            ttk.Entry(top, textvariable=self.path_var).grid(
                row=0, column=1, sticky="ew", padx=4)
            ttk.Button(top, text="浏览…", command=self._browse).grid(row=0, column=2)
            top.columnconfigure(1, weight=1)

            ttk.Label(top, text="统计方式").grid(row=1, column=0, sticky="w", pady=(6, 0))
            self.mode_box = ttk.Combobox(
                top, textvariable=self.mode_var, values=list(MODE_LABELS),
                state="readonly", width=34)
            self.mode_box.grid(row=1, column=1, sticky="w", padx=4, pady=(6, 0))
            self.mode_box.bind("<<ComboboxSelected>>", lambda _e: self._sync_mode())
            ext_box = ttk.Frame(top)
            ext_box.grid(row=1, column=2, sticky="w", pady=(6, 0))
            ttk.Label(ext_box, text="扩展名").pack(side="left")
            self.ext_entry = ttk.Entry(ext_box, textvariable=self.ext_var, width=22)
            self.ext_entry.pack(side="left", padx=4)

            ttk.Label(top, text="分布口径").grid(row=2, column=0, sticky="w", pady=(6, 0))
            ttk.Combobox(
                top, textvariable=self.dist_var, values=list(DIST_LABELS),
                state="readonly", width=22
            ).grid(row=2, column=1, sticky="w", padx=4, pady=(6, 0))
            ttk.Checkbutton(
                top, text="追加其他语言行（JavaScript / SQL / HTML …）",
                variable=self.others_var).grid(row=2, column=2, sticky="w", pady=(6, 0))

            bar = ttk.Frame(self)
            bar.pack(fill="x", **pad)
            self.run_btn = ttk.Button(bar, text="开始统计", command=self.start)
            self.run_btn.pack(side="left")
            ttk.Button(bar, text="导出 CSV", command=self.export_csv).pack(side="left", padx=6)
            ttk.Button(bar, text="复制表格", command=self.copy_table).pack(side="left")
            ttk.Button(bar, text="查看明细", command=self.show_detail).pack(side="left", padx=6)

            table = ttk.Frame(self)
            table.pack(fill="both", expand=True, **pad)
            self.tree = ttk.Treeview(
                table, columns=COLUMNS, show="headings", height=12, selectmode="browse")
            widths = (110, 80, 80, 70, 80, 80, 60, 60, 60, 70)
            for col, w in zip(COLUMNS, widths):
                self.tree.heading(col, text=col)
                self.tree.column(
                    col, width=w, minwidth=54, stretch=True,
                    anchor="w" if col == "语言" else "e")
            yscroll = ttk.Scrollbar(table, orient="vertical", command=self.tree.yview)
            xscroll = ttk.Scrollbar(table, orient="horizontal", command=self.tree.xview)
            self.tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
            self.tree.grid(row=0, column=0, sticky="nsew")
            yscroll.grid(row=0, column=1, sticky="ns")
            xscroll.grid(row=1, column=0, sticky="ew")
            table.rowconfigure(0, weight=1)
            table.columnconfigure(0, weight=1)
            self.tree.bind("<Double-1>", lambda _e: self.show_detail())

            ttk.Label(self, textvariable=self.status_var, anchor="w",
                      relief="sunken", padding=(6, 3)).pack(fill="x", side="bottom")
            self._sync_mode()

        def _sync_mode(self) -> None:
            is_ext = MODE_LABELS[self.mode_var.get()] == "ext"
            self.ext_entry.configure(state="normal" if is_ext else "disabled")

        def _browse(self) -> None:
            chosen = filedialog.askdirectory(initialdir=self.path_var.get() or ".")
            if chosen:
                self.path_var.set(chosen)

        # ---------------- 自动化冒烟测试用
        def _smoke_start(self, path: str) -> None:
            self.path_var.set(path)
            self.start()
            self._smoke_deadline = time.time() + 30
            self.after(200, self._smoke_check)

        def _smoke_check(self) -> None:
            if self.result is not None:
                items = self.tree.get_children()
                first = self.tree.item(items[0], "values") if items else ()
                print(f"SMOKE OK rows={len(self.result.rows)} tree_rows={len(items)} "
                      f"files={self.result.scanned} elapsed={self.result.elapsed:.2f}s")
                print("SMOKE first_row=" + " | ".join(str(c) for c in first))
                self.destroy()
            elif time.time() > self._smoke_deadline:
                print("SMOKE TIMEOUT")
                self.destroy()
            else:
                self.after(200, self._smoke_check)

        # ---------------- 统计（子线程，避免界面卡死）
        def start(self) -> None:
            root = self.path_var.get().strip().strip('"')
            if not root:
                messagebox.showwarning("提示", "请先填写目标路径")
                return
            if not Path(root).exists():
                messagebox.showerror("路径不存在", f"找不到：{root}")
                return
            mode = MODE_LABELS[self.mode_var.get()]
            if mode == "ext" and not self.ext_var.get().strip():
                messagebox.showwarning("提示", "请填写要统计的扩展名，例如 .py,.js")
                return
            distribution = DIST_LABELS[self.dist_var.get()]
            self.run_btn.configure(state="disabled")
            self.status_var.set("统计中，请稍候…")
            args = dict(
                root=root, mode=mode, ext=self.ext_var.get().strip(),
                distribution=distribution, append_others=self.others_var.get(),
            )
            threading.Thread(target=self._work, args=(args,), daemon=True).start()
            self.after(100, self._poll)

        def _work(self, args: dict) -> None:
            try:
                self._queue.put(("ok", analyze(**args)))
            except Exception as exc:  # 把异常带回主线程显示
                self._queue.put(("err", exc))

        def _poll(self) -> None:
            try:
                kind, payload = self._queue.get_nowait()
            except queue.Empty:
                self.after(100, self._poll)
                return
            self.run_btn.configure(state="normal")
            if kind == "err":
                self.status_var.set("统计失败")
                messagebox.showerror("统计失败", str(payload))
                return
            self.result = payload
            self._render(payload)
            if MODE_LABELS[self.mode_var.get()] == "hand" and not payload.ignore_file:
                self.status_var.set(
                    self.status_var.get() + "；未找到 .clocignore，本次按原始统计")
                messagebox.showwarning(
                    "未找到排除清单",
                    "目标目录下没有 .clocignore，本次按「原始统计」处理。\n\n"
                    "解决办法：把工具目录里的 clocignore.example 复制到目标目录，"
                    "改名为 .clocignore 并按需增删规则。")

        def _render(self, result: Result) -> None:
            self.tree.delete(*self.tree.get_children())
            for row in result.rows:
                self.tree.insert("", "end", values=row)
            unit = "函数" if result.distribution == "func" else "文件"
            note = f"已统计 {result.scanned} 个文件，耗时 {result.elapsed:.2f} 秒"
            if result.skipped_ignored:
                note += f"；按清单跳过 {result.skipped_ignored}"
            if result.skipped_dup:
                note += f"；重复跳过 {result.skipped_dup}"
            if result.skipped_binary:
                note += f"；二进制跳过 {result.skipped_binary}"
            self.status_var.set(f"{note}；分布按{unit}代码行数")

        # ---------------- 导出
        def export_csv(self) -> None:
            if not self.result:
                messagebox.showinfo("提示", "请先统计一次")
                return
            default = f"代码量统计_{time.strftime('%Y%m%d_%H%M%S')}.csv"
            path = filedialog.asksaveasfilename(
                defaultextension=".csv", initialfile=default,
                filetypes=[("CSV 文件", "*.csv"), ("所有文件", "*.*")])
            if not path:
                return
            save_csv(self.result, path)
            self.status_var.set(f"已导出：{path}")

        def copy_table(self) -> None:
            if not self.result:
                messagebox.showinfo("提示", "请先统计一次")
                return
            self.clipboard_clear()
            self.clipboard_append(to_tsv(self.result))
            self.status_var.set("表格已复制，可直接粘贴到 Excel / Word")

        def show_detail(self) -> None:
            if not self.result or not self.result.files:
                messagebox.showinfo("提示", "请先统计一次")
                return
            selected = self.tree.selection()
            only_lang = None
            if selected:
                only_lang = str(self.tree.item(selected[0], "values")[0])
                if self.result.languages.get(
                        only_lang, LangStat()).files == 0:
                    only_lang = None

            win = tk.Toplevel(self)
            title = f"文件明细 - {only_lang}" if only_lang else "文件明细"
            win.title(title)
            win.geometry("900x480")
            cols = ("文件", "语言", "代码行", "空行", "注释行", "函数个数")
            tree = ttk.Treeview(win, columns=cols, show="headings")
            for col, w in zip(cols, (560, 100, 80, 70, 80, 80)):
                tree.heading(col, text=col)
                tree.column(col, width=w, anchor="w" if col in ("文件", "语言") else "e")
            rows = [f for f in self.result.files
                    if only_lang is None or f.lang == only_lang
                    or LANG_DISPLAY.get(f.lang, f.lang) == only_lang]
            rows.sort(key=lambda f: (-f.code, f.path))
            for f in rows:
                tree.insert("", "end", values=(
                    f.path, LANG_DISPLAY.get(f.lang, f.lang),
                    f.code, f.blank, f.comment, f.funcs))
            yscroll = ttk.Scrollbar(win, orient="vertical", command=tree.yview)
            tree.configure(yscrollcommand=yscroll.set)
            tree.pack(side="left", fill="both", expand=True)
            yscroll.pack(side="right", fill="y")

    App().mainloop()
    return 0


# ---------------------------------------------------------------- 入口

def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="代码量统计（表格格式 + 图形界面）",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "示例：\n"
            "  python code_stats.py                          # 打开图形界面\n"
            "  python code_stats.py 四则运算项目              # 命令行直接出表\n"
            "  python code_stats.py . -m hand                 # 按 .clocignore 只算手写\n"
            "  python code_stats.py . -m hand --csv out.csv   # 顺便导出 CSV\n"
        ),
    )
    parser.add_argument("path", nargs="?", help="目标目录或文件；不给则打开图形界面")
    parser.add_argument("-m", "--mode", choices=["raw", "hand", "ext"], default="raw",
                        help="统计方式：raw 原始 / hand 手写口径(.clocignore) / ext 仅指定扩展名")
    parser.add_argument("-e", "--ext", help="扩展名，逗号分隔（mode=ext 时使用）")
    parser.add_argument("-d", "--distribution", choices=["func", "file"], default="func",
                        help="分布口径：func 按函数代码行（默认）/ file 按文件代码行")
    parser.add_argument("--no-others", action="store_true", help="只输出五行模板语言")
    parser.add_argument("--csv", help="把表格导出为 CSV")
    parser.add_argument("--gui", action="store_true", help="强制打开图形界面")
    parser.add_argument("--gui-smoke", metavar="PATH", help=argparse.SUPPRESS)
    args = parser.parse_args(argv)

    if args.gui_smoke:
        return launch_gui(smoke_path=args.gui_smoke)
    if args.gui or not args.path:
        return launch_gui()

    # 给了 -e 但没写 -m 时，自动按“仅指定扩展名”处理
    if args.ext and args.mode == "raw":
        args.mode = "ext"

    try:
        result = analyze(
            args.path, mode=args.mode, ext=args.ext,
            distribution=args.distribution, append_others=not args.no_others,
        )
    except (FileNotFoundError, ValueError) as exc:
        print(exc, file=sys.stderr)
        return 2

    if args.mode == "hand" and not result.ignore_file:
        print("提示：目标目录下没有 .clocignore，本次按原始统计处理。\n"
              "      可复制 代码量统计工具/clocignore.example 到目标目录并改名 .clocignore。\n",
              file=sys.stderr)

    print(format_table(result))
    if args.csv:
        out = save_csv(result, args.csv)
        print(f"\n已导出 CSV：{out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
