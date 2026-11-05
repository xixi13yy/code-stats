#!/usr/bin/env python3
"""统计代码量：文件数、空行、注释行、代码行。

用法:
    python cloc.py [路径] [-e .py,.js] [-E .html] [-x node_modules,dist] [--by-file] [--time]

默认会自动去重（内容完全相同的文件只算一份）。可在统计根目录放一个
.clocignore 文件列出要跳过的文件，每行一条规则：
    # 注释
    .html                        扩展名
    generate_report.py           文件名
    plane/generate_report.py     相对路径
    plane/**                     通配符
"""
from __future__ import annotations

import argparse
import fnmatch
import hashlib
import os
import sys
import time
from dataclasses import dataclass, field
from pathlib import Path

# Windows 控制台默认 cp936，直接输出中文会乱码
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

# 扩展名 -> (语言名, 行注释列表, 块注释对列表)
# 块注释以 (开始, 结束) 表示；开始/结束为空串表示该语言无此行注释形式
LANGS: dict[str, tuple[str, tuple[str, ...], tuple[tuple[str, str], ...]]] = {
    ".py": ("Python", ("#",), (('"""', '"""'), ("'''", "'''"))),
    ".pyw": ("Python", ("#",), (('"""', '"""'), ("'''", "'''"))),
    ".js": ("JavaScript", ("//",), (("/*", "*/"),)),
    ".jsx": ("JavaScript", ("//",), (("/*", "*/"),)),
    ".mjs": ("JavaScript", ("//",), (("/*", "*/"),)),
    ".cjs": ("JavaScript", ("//",), (("/*", "*/"),)),
    ".ts": ("TypeScript", ("//",), (("/*", "*/"),)),
    ".tsx": ("TypeScript", ("//",), (("/*", "*/"),)),
    ".java": ("Java", ("//",), (("/*", "*/"),)),
    ".c": ("C", ("//",), (("/*", "*/"),)),
    ".h": ("C/C++ Header", ("//",), (("/*", "*/"),)),
    ".cc": ("C++", ("//",), (("/*", "*/"),)),
    ".cpp": ("C++", ("//",), (("/*", "*/"),)),
    ".cxx": ("C++", ("//",), (("/*", "*/"),)),
    ".hpp": ("C++ Header", ("//",), (("/*", "*/"),)),
    ".hh": ("C++ Header", ("//",), (("/*", "*/"),)),
    ".cs": ("C#", ("//",), (("/*", "*/"),)),
    ".go": ("Go", ("//",), (("/*", "*/"),)),
    ".rs": ("Rust", ("//",), (("/*", "*/"),)),
    ".swift": ("Swift", ("//",), (("/*", "*/"),)),
    ".kt": ("Kotlin", ("//",), (("/*", "*/"),)),
    ".kts": ("Kotlin", ("//",), (("/*", "*/"),)),
    ".scala": ("Scala", ("//",), (("/*", "*/"),)),
    ".php": ("PHP", ("//", "#"), (("/*", "*/"),)),
    ".rb": ("Ruby", ("#",), (("=begin", "=end"),)),
    ".pl": ("Perl", ("#",), ()),
    ".pm": ("Perl", ("#",), ()),
    ".sh": ("Shell", ("#",), ()),
    ".bash": ("Shell", ("#",), ()),
    ".zsh": ("Shell", ("#",), ()),
    ".ps1": ("PowerShell", ("#",), (("<#", "#>"),)),
    ".bat": ("Batch", ("rem", "::"), ()),
    ".sql": ("SQL", ("--",), (("/*", "*/"),)),
    ".lua": ("Lua", ("--",), (("--[[", "]]"),)),
    ".r": ("R", ("#",), ()),
    ".m": ("MATLAB/ObjC", ("%", "//"), (("/*", "*/"), ("%{", "%}"))),
    ".html": ("HTML", (), (("<!--", "-->"),)),
    ".htm": ("HTML", (), (("<!--", "-->"),)),
    ".xml": ("XML", (), (("<!--", "-->"),)),
    ".css": ("CSS", (), (("/*", "*/"),)),
    ".scss": ("SCSS", ("//",), (("/*", "*/"),)),
    ".less": ("LESS", ("//",), (("/*", "*/"),)),
    ".vue": ("Vue", ("//",), (("/*", "*/"), ("<!--", "-->"))),
    ".svelte": ("Svelte", ("//",), (("/*", "*/"), ("<!--", "-->"))),
    ".json": ("JSON", (), ()),
    ".yaml": ("YAML", ("#",), ()),
    ".yml": ("YAML", ("#",), ()),
    ".toml": ("TOML", ("#",), ()),
    ".ini": ("INI", (";", "#"), ()),
    ".cfg": ("INI", (";", "#"), ()),
    ".md": ("Markdown", (), (("<!--", "-->"),)),
    ".tex": ("TeX", ("%",), ()),
    ".asm": ("Assembly", (";",), ()),
    ".s": ("Assembly", (";",), ()),
    ".v": ("Verilog", ("//",), (("/*", "*/"),)),
    ".sv": ("SystemVerilog", ("//",), (("/*", "*/"),)),
    ".vhd": ("VHDL", ("--",), ()),
    ".vhdl": ("VHDL", ("--",), ()),
    ".dart": ("Dart", ("//",), (("/*", "*/"),)),
    ".ex": ("Elixir", ("#",), ()),
    ".exs": ("Elixir", ("#",), ()),
    ".erl": ("Erlang", ("%",), ()),
    ".hs": ("Haskell", ("--",), (("{-", "-}"),)),
    ".ml": ("OCaml", (), (("(*", "*)"),)),
    ".jl": ("Julia", ("#",), (("#=", "=#"),)),
    ".sqlite": ("SQL", ("--",), ()),
    ".gradle": ("Gradle", ("//",), (("/*", "*/"),)),
    ".cmake": ("CMake", ("#",), ()),
    ".dockerfile": ("Dockerfile", ("#",), ()),
}

SPECIAL_NAMES = {
    "Dockerfile": "Dockerfile",
    "Makefile": "Makefile",
    "makefile": "Makefile",
    "GNUmakefile": "Makefile",
    "CMakeLists.txt": "CMake",
    "Rakefile": "Ruby",
    "Gemfile": "Ruby",
    ".gitignore": "GitIgnore",
    ".dockerignore": "DockerIgnore",
    ".env": "DotEnv",
}

DEFAULT_EXCLUDE_DIRS = {
    ".git", ".svn", ".hg", "node_modules", "__pycache__", ".venv", "venv",
    "env", ".env", "dist", "build", "target", "out", "bin", "obj",
    ".idea", ".vscode", ".vs", "vendor", "third_party", ".mypy_cache",
    ".pytest_cache", ".gradle", ".next", ".nuxt", "coverage", ".tox",
}

MAX_BYTES = 5 * 1024 * 1024  # 跳过大于 5MB 的文件，多半是数据/二进制


@dataclass
class Stat:
    files: int = 0
    blank: int = 0
    comment: int = 0
    code: int = 0
    per_file: list[tuple[str, int, int, int]] = field(default_factory=list)

    @property
    def total(self) -> int:
        return self.blank + self.comment + self.code


def lang_of(path: Path) -> tuple[str, tuple[str, ...], tuple[tuple[str, str], ...]] | None:
    if path.name in SPECIAL_NAMES:
        name = SPECIAL_NAMES[path.name]
        return LANGS.get(path.suffix.lower(), (name, ("#",), ())) if path.suffix else (name, ("#",), ())
    return LANGS.get(path.suffix.lower())


def read_source(path: Path):
    """读出源码文本和内容哈希。二进制或无法解码时文本为 None。"""
    try:
        data = path.read_bytes()
    except OSError:
        return None, None
    if len(data) > MAX_BYTES:
        return None, None
    digest = hashlib.md5(data).hexdigest()
    # 前 8KB 出现 NUL 字节即视为二进制（图片/编译产物等改了扩展名的文件）
    if b"\x00" in data[:8192]:
        return None, None
    for enc in ("utf-8", "gbk"):
        try:
            return data.decode(enc), digest
        except UnicodeDecodeError:
            continue
    return None, None


def count_text(text: str, line_comments, block_comments):
    """返回 (空行, 注释行, 代码行)。"""
    blank = comment = code = 0
    block_end = None  # 当前所处块注释的结束标记

    for raw in text.splitlines():
        line = raw.strip()
        if not line:
            blank += 1
            continue

        remainder = line
        has_code = False

        while remainder:
            if block_end is not None:
                idx = remainder.find(block_end)
                if idx == -1:
                    remainder = ""
                    break
                remainder = remainder[idx + len(block_end):]
                block_end = None
                continue

            # 找到最早的块注释起点和行注释起点
            bc_pos, bc_pair = None, None
            for start, end in block_comments:
                pos = remainder.find(start)
                if pos != -1 and (bc_pos is None or pos < bc_pos):
                    bc_pos, bc_pair = pos, (start, end)
            lc_pos = None
            for mark in line_comments:
                pos = remainder.find(mark)
                if pos != -1 and (lc_pos is None or pos < lc_pos):
                    lc_pos = pos

            # 行注释要排除 "://" 这类 URL，"rem" 要作为独立单词
            if lc_pos is not None and lc_pos > 0 and mark_is_embedded(remainder, lc_pos):
                lc_pos = None

            if lc_pos is not None and (bc_pos is None or lc_pos < bc_pos):
                if remainder[:lc_pos].strip():
                    has_code = True
                remainder = ""
                break

            if bc_pos is not None:
                if remainder[:bc_pos].strip():
                    has_code = True
                start, end = bc_pair
                remainder = remainder[bc_pos + len(start):]
                block_end = end
                if end == start:  # Python 三引号：本行内找配对
                    block_end = end
                continue

            if remainder.strip():
                has_code = True
            break

        if has_code:
            code += 1
        else:
            comment += 1

    return blank, comment, code


def load_ignore(path: Path) -> list[str]:
    try:
        # utf-8-sig：记事本保存的 .clocignore 带 BOM，不剥掉会让第一行规则失效
        lines = path.read_text(encoding="utf-8-sig", errors="replace").splitlines()
    except OSError:
        return []
    return [p for p in (ln.split("#", 1)[0].strip().lstrip("\ufeff") for ln in lines) if p]


def is_ignored(rel: str, name: str, patterns: list[str]) -> str | None:
    """返回命中的规则，未命中返回 None。rel 用 / 分隔的相对路径。"""
    for pat in patterns:
        # 裸扩展名写法：.html / *.html 都按后缀匹配
        if pat.startswith(".") and "/" not in pat and "*" not in pat:
            if name.lower().endswith(pat.lower()):
                return pat
        if pat.startswith("*."):
            if name.lower().endswith(pat[1:].lower()):
                return pat
        if fnmatch.fnmatch(name, pat) or fnmatch.fnmatch(rel, pat):
            return pat
        # 目录前缀：plane/ 命中 plane/ 下的所有文件
        if pat.endswith("/") and rel.startswith(pat):
            return pat
    return None


def mark_is_embedded(text: str, pos: int) -> bool:
    """判断 pos 处的注释标记是否只是字符串/URL 的一部分。"""
    prev = text[pos - 1]
    mark = text[pos]
    if mark == "/" and prev == ":":
        return True  # http://
    if mark.isalnum() and (prev.isalnum() or prev == "_"):
        return True  # rem 出现在单词中间
    return False


def main() -> int:
    parser = argparse.ArgumentParser(description="统计代码量")
    parser.add_argument("path", nargs="?", default=".", help="要统计的目录或文件，默认当前目录")
    parser.add_argument("-e", "--ext", help="只统计这些扩展名，逗号分隔，如 .py,.js")
    parser.add_argument("-E", "--exclude-ext",
                        help="排除这些扩展名，逗号分隔，如 .html,.json（与 -e 同时用时优先生效）")
    parser.add_argument("-x", "--exclude", help="额外排除的目录名，逗号分隔")
    parser.add_argument("--by-file", action="store_true", help="列出每个文件的明细")
    parser.add_argument("--ignore-file", help="忽略清单文件，默认用根目录下的 .clocignore")
    parser.add_argument("--no-dedup", action="store_true",
                        help="关闭内容去重（默认相同内容的文件只算一份）")
    parser.add_argument("--no-default-exclude", action="store_true", help="不排除 node_modules 等默认目录")
    parser.add_argument("-s", "--sort", choices=["code", "name", "files"], default="code", help="排序方式")
    parser.add_argument("--time", action="store_true", help="结束时显示统计耗时")
    args = parser.parse_args()

    start = time.perf_counter()

    root = Path(args.path)
    if not root.exists():
        print(f"路径不存在: {root}", file=sys.stderr)
        return 1

    only_ext = None
    if args.ext:
        only_ext = {e if e.startswith(".") else "." + e for e in args.ext.lower().split(",")}

    drop_ext = set()
    if args.exclude_ext:
        drop_ext = {e if e.startswith(".") else "." + e for e in args.exclude_ext.lower().split(",")}
        drop_ext = {e.strip() for e in drop_ext if e.strip()}

    excludes = set() if args.no_default_exclude else set(DEFAULT_EXCLUDE_DIRS)
    if args.exclude:
        excludes |= {d.strip() for d in args.exclude.split(",") if d.strip()}

    ignore_path = Path(args.ignore_file) if args.ignore_file else (root / ".clocignore")
    ignore_patterns = load_ignore(ignore_path)

    stats: dict[str, Stat] = {}
    total = Stat()
    seen: dict[str, str] = {}          # 内容哈希 -> 首次出现的文件
    dup_files = 0
    dup_lines = 0
    ignored_files = 0
    binary_files = 0

    def handle(f: Path) -> None:
        nonlocal dup_files, dup_lines, ignored_files, binary_files
        info = lang_of(f)
        if info is None:
            return
        name, line_comments, block_comments = info
        if only_ext and f.suffix.lower() not in only_ext:
            return
        if f.suffix.lower() in drop_ext:
            return
        try:
            rel = f.relative_to(root if root.is_dir() else root.parent).as_posix()
        except ValueError:
            rel = f.name
        if is_ignored(rel, f.name, ignore_patterns):
            ignored_files += 1
            return
        text, digest = read_source(f)
        if text is None:
            binary_files += 1
            return
        if not args.no_dedup:
            if digest in seen:
                dup_files += 1
                dup_lines += count_text(text, line_comments, block_comments)[2]
                return
            seen[digest] = str(f)
        blank, comment, code = count_text(text, line_comments, block_comments)
        st = stats.setdefault(name, Stat())
        st.files += 1
        st.blank += blank
        st.comment += comment
        st.code += code
        st.per_file.append((str(f), blank, comment, code))
        total.files += 1
        total.blank += blank
        total.comment += comment
        total.code += code

    if root.is_file():
        handle(root)
    else:
        for dirpath, dirnames, filenames in os.walk(root):
            dirnames[:] = sorted(d for d in dirnames if d not in excludes and not d.startswith("."))
            for fn in sorted(filenames):
                handle(Path(dirpath) / fn)

    if not stats:
        print("没有找到可统计的代码文件。可用 -e 指定扩展名，或检查路径。")
        return 0

    order = {
        "code": lambda kv: -kv[1].code,
        "name": lambda kv: kv[0].lower(),
        "files": lambda kv: -kv[1].files,
    }[args.sort]
    rows = sorted(stats.items(), key=order)

    width = max(len("Language"), max(len(n) for n in stats))
    header = f"{'Language':<{width}} {'Files':>7} {'Blank':>8} {'Comment':>9} {'Code':>9}"
    print(header)
    print("-" * len(header))
    for name, st in rows:
        print(f"{name:<{width}} {st.files:>7} {st.blank:>8} {st.comment:>9} {st.code:>9}")
    print("-" * len(header))
    print(f"{'SUM':<{width}} {total.files:>7} {total.blank:>8} {total.comment:>9} {total.code:>9}")

    ratio = total.comment / total.code * 100 if total.code else 0
    print(f"\n代码行 {total.code}，注释行 {total.comment}，空行 {total.blank}，总行数 {total.total}")
    print(f"注释率 {ratio:.1f}%  |  平均每文件 {total.code / total.files:.1f} 行代码")

    if dup_files:
        print(f"已去重 {dup_files} 个重复文件（少算 {dup_lines} 行代码）")
    if ignored_files:
        print(f"已按 {ignore_path.name} 跳过 {ignored_files} 个文件")
    if binary_files:
        print(f"已跳过 {binary_files} 个二进制/无法解码的文件")

    elapsed = time.perf_counter() - start
    if args.time:
        print(f"耗时 {elapsed:.3f} 秒（{total.files} 个文件，{elapsed / total.files * 1000:.2f} ms/文件）")

    if args.by_file:
        print("\n明细:")
        files = sorted(
            ((n, pf) for n, st in stats.items() for pf in st.per_file),
            key=lambda x: -x[1][3],
        )
        fw = max(len(f) for _, (f, _, _, _) in files)
        print(f"{'File':<{fw}} {'Blank':>8} {'Comment':>9} {'Code':>9}")
        for name, (f, blank, comment, code) in files:
            print(f"{f:<{fw}} {blank:>8} {comment:>9} {code:>9}")

    return 0


if __name__ == "__main__":
    sys.exit(main())
