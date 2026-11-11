# -*- coding: utf-8 -*-
"""排版 profile：加载 TOML、base 派生合并、字段校验、字体可用性解析。

profile 是唯一事实源：md2docx.py 按它渲染，check_docx.py 按它校验，
避免"脚本常量"与"规范文档"各写一份数值而互相漂移。

用法:
    from profile import load_profile
    cfg, notices = load_profile("tech-doc")

    python scripts/profile.py --list          # 列出可用 profile
    python scripts/profile.py --show tech-doc # 打印合并后的完整参数

约定:
- profile 查找顺序：`--profile` 给的路径 → 环境变量 WORD_DOC_PROFILES 列出的目录
  → 仓库根 profiles/ → 当前目录 profiles/。
- `base = "父 profile"` 做深度合并；[[headings]] 按 level 合并，子层只写差异项。
- 未实现的功能字段（如 mirror = true）**直接报错**，不允许静默忽略——
  静默忽略会产出"看着合规、其实不对"的文档。
- 标了 verified = false 的块会在加载时提示，校验器对这些项只告警不报错。
"""
from __future__ import annotations

import copy
import os
import sys

try:  # Python 3.11+
    import tomllib
except ModuleNotFoundError:  # pragma: no cover
    tomllib = None

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
REPO_ROOT = os.path.dirname(SCRIPT_DIR)
DEFAULT_PROFILE = "ug-thesis-bit"

PAGE_SIZES = {
    "A4": (21.0, 29.7),
    "Letter": (21.59, 27.94),
    "B5": (17.6, 25.0),
    "16K": (18.4, 26.0),
}

DEFAULTS = {
    "name": "",
    "label": "",
    "description": "",
    "base": "",
    "standards": [],
    "verified": True,
    "references_standard": "2025",
    "page": {
        "size": "A4",
        "margins": {"top": 2.54, "bottom": 2.54, "left": 3.18, "right": 3.18},
        "gutter": 0.0,
        "mirror": False,
        "header_distance": 1.5,
        "footer_distance": 1.75,
        "different_first_page": False,
    },
    "body": {
        "font_cn": "宋体",
        "font_en": "Times New Roman",
        "size": 12,
        "align": "justify",
        "first_line_indent": {"chars": 2, "mode": "first_line"},
        "space_before": {"pt": 0.0},
        "space_after": {"pt": 0.0},
        "char_spacing": 0.0,
        "image_line_spacing": 1.5,
        "max_image_width": 14.5,
        "line_spacing": {"mode": "fixed", "value": 22},
    },
    "headings": [
        {"level": 1, "font": "黑体", "size": 16, "bold": False, "align": "left",
         "space_before": {"lines": 0.5}, "space_after": {"lines": 0.5},
         "page_break_before": False},
    ],
    "numbering": {
        "scheme": "decimal-three",
        "max_depth": 3,
        "figure": {"scheme": "by_chapter", "sep": "-", "caption": "below"},
        "table": {"scheme": "by_chapter", "sep": "-", "caption": "above"},
    },
    "table": {
        "style": "three_line",
        "cell_font": "宋体",
        "cell_size": 10,
        "header_repeat": True,
        "border": {"outer": 12, "header": 6},
        "caption": {"font": "宋体", "size": 9,
                    "space_before": {"pt": 0.0}, "space_after": {"pt": 0.0}},
    },
    "figure": {
        "caption_font": "宋体",
        "caption_size": 9,
        "caption_space_before": {"pt": 0.0},
        "caption_space_after": {"pt": 0.0},
    },
    "header_footer": {
        "header_text": "",
        "header_font": "宋体",
        "header_size": 9,
        "header_rule": "none",
        "page_number": {
            "position": "footer-center",
            "format": "dash",
            "font": "宋体",
            "size": 9,
            "start_at": 1,
        },
    },
    "structure": {
        "toc": False,
        "toc_title": "目录",
        "front_numbering": "none",
        "body_start_at": 1,
    },
    "special": {
        "signature_page": False,
        "seal_space": False,
        "amount_in_words": False,
        "doc_control_page": False,
        "revision_history": False,
        "classification": "",
    },
    "fonts": {"fallback": {}},
}

ENUMS = {
    "page.size": set(PAGE_SIZES),
    "body.align": {"left", "center", "right", "justify"},
    "body.line_spacing.mode": {"fixed", "multiple", "grid"},
    "body.first_line_indent.mode": {"first_line", "hanging", "none"},
    "headings.align": {"left", "center", "right", "justify"},
    "numbering.scheme": {"decimal-three", "cn-gov", "article", "chapter"},
    "numbering.figure.scheme": {"by_chapter", "continuous"},
    "numbering.table.scheme": {"by_chapter", "continuous"},
    "numbering.figure.caption": {"below", "above"},
    "numbering.table.caption": {"below", "above"},
    "table.style": {"three_line", "grid"},
    "header_footer.page_number.position": {"footer-center", "header-right", "odd-even-outside"},
    "header_footer.page_number.format": {"plain", "dash", "cn", "roman"},
    "structure.front_numbering": {"none", "plain", "roman"},
}

# 中文字体族在 Windows 字体注册表里常以英文名登记
FONT_ALIASES = {
    "宋体": ["simsun"],
    "新宋体": ["nsimsun"],
    "黑体": ["simhei"],
    "楷体": ["kaiti", "simkai"],
    "仿宋": ["fangsong", "simfang"],
    "微软雅黑": ["microsoft yahei", "msyh"],
    "等线": ["dengxian", "deng"],
    "华文中宋": ["stzhongsong"],
    "华文楷体": ["stkaiti"],
    "华文仿宋": ["stfangsong"],
    "方正小标宋简体": ["fzxbs"],
}


class ProfileError(Exception):
    """profile 缺失、字段非法或使用了未实现的能力。"""


def _read_toml(path):
    if tomllib is None:  # pragma: no cover
        raise ProfileError("需要 Python 3.11+ 的 tomllib 才能读取 profile")
    # utf-8-sig 兼容 Windows 记事本/PowerShell 写入的 BOM
    with open(path, "r", encoding="utf-8-sig") as f:
        return tomllib.loads(f.read())


def find_profile(spec):
    """把 profile 名或路径解析成 .toml 文件路径。"""
    if not spec:
        spec = DEFAULT_PROFILE
    candidates = []
    if spec.endswith(".toml") or os.sep in spec or "/" in spec:
        candidates.append(spec)
    else:
        dirs = [d for d in os.environ.get("WORD_DOC_PROFILES", "").split(os.pathsep) if d]
        dirs += [os.path.join(REPO_ROOT, "profiles"),
                 os.path.join(os.getcwd(), "profiles"),
                 os.getcwd()]
        for d in dirs:
            candidates.append(os.path.join(d, spec + ".toml"))
    for path in candidates:
        if os.path.isfile(path):
            return os.path.abspath(path)
    raise ProfileError(
        f"找不到 profile「{spec}」。已查找: {', '.join(os.path.dirname(p) for p in candidates)}")


def _is_number(value):
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def _validate(node, default, path, errors):
    """按 DEFAULTS 的形状校验键名与类型；额外允许 verified / note。"""
    if path == "fonts.fallback":  # 自由映射：字体名 → 候选字体数组
        for key, value in node.items():
            if not isinstance(value, list) or not all(isinstance(x, str) for x in value):
                errors.append(f"fonts.fallback.{key}：应为字符串数组")
        return
    for key, value in node.items():
        if key.startswith("_") or key in ("verified", "note"):
            continue
        full = f"{path}.{key}" if path else key
        if key not in default:
            errors.append(f"{full}：未知字段（拼写错误或本版本不支持）")
            continue
        ref = default[key]
        if isinstance(ref, dict):
            if not isinstance(value, dict):
                errors.append(f"{full}：应为表，实际是 {type(value).__name__}")
            else:
                _validate(value, ref, full, errors)
        elif isinstance(ref, list):
            if not isinstance(value, list):
                errors.append(f"{full}：应为数组，实际是 {type(value).__name__}")
            elif key == "headings":
                for item in value:
                    if not isinstance(item, dict):
                        errors.append(f"{full}：数组元素应为表")
                    else:
                        _validate(item, DEFAULTS["headings"][0], full, errors)
            elif not all(isinstance(x, str) for x in value):
                errors.append(f"{full}：数组元素应为字符串")
        elif isinstance(ref, bool):
            if not isinstance(value, bool):
                errors.append(f"{full}：应为 true/false")
        elif _is_number(ref):
            if not _is_number(value):
                errors.append(f"{full}：应为数字")
        elif isinstance(ref, str):
            if not isinstance(value, str):
                errors.append(f"{full}：应为字符串")
    # space_before / space_after 允许 {pt: x} 或 {lines: y}
    if path.endswith("space_before") or path.endswith("space_after"):
        extra = set(node) - {"pt", "lines"} - {"verified", "note"}
        if extra:
            errors.append(f"{path}：只接受 pt 或 lines，多了 {sorted(extra)}")


def _check_enums(cfg, errors):
    for dotted in ENUMS:
        allowed = ENUMS[dotted]
        if dotted.startswith("headings."):
            field = dotted.split(".", 1)[1]
            for heading in cfg.get("headings", []):
                if heading.get(field) not in allowed:
                    errors.append(f"headings[{heading.get('level')}].{field}：取值 "
                                  f"{heading.get(field)!r} 不合法，可选 {sorted(allowed)}")
            continue
        value = get_path(cfg, dotted)
        if value not in allowed:
            errors.append(f"{dotted}：取值 {value!r} 不合法，可选 {sorted(allowed)}")


def get_path(cfg, dotted, default=None):
    node = cfg
    for part in dotted.split("."):
        if not isinstance(node, dict) or part not in node:
            return default
        node = node[part]
    return node


def set_path(cfg, dotted, value):
    parts = dotted.split(".")
    node = cfg
    for part in parts[:-1]:
        node = node[part]
    node[parts[-1]] = value


def _deep_merge(base, over):
    """深度合并：dict 递归，[[headings]] 按 level 合并，其余列表整体替换。"""
    result = copy.deepcopy(base)
    for key, value in over.items():
        if key == "headings" and isinstance(value, list) and isinstance(result.get(key), list):
            result[key] = _merge_headings(result[key], value)
        elif isinstance(value, dict) and isinstance(result.get(key), dict):
            result[key] = _deep_merge(result[key], value)
        else:
            result[key] = copy.deepcopy(value)
    return result


def _merge_headings(base_list, over_list):
    by_level = {h.get("level"): copy.deepcopy(h) for h in base_list}
    order = [h.get("level") for h in base_list]
    for item in over_list:
        level = item.get("level")
        if level in by_level:
            by_level[level] = _deep_merge(by_level[level], item)
        else:
            by_level[level] = copy.deepcopy(item)
            order.append(level)
    return [by_level[lv] for lv in order if lv in by_level]


# 未实现的能力：(字段路径, 判断"用户是否真的用了它")
UNSUPPORTED = {
    "page.gutter": lambda v: bool(v),
    "page.mirror": lambda v: bool(v),
    "body.char_spacing": lambda v: bool(v),
    "body.line_spacing.mode": lambda v: v == "grid",
    "body.first_line_indent.mode": lambda v: v != "first_line",
    "numbering.scheme": lambda v: v != "decimal-three",
    "table.style": lambda v: v != "three_line",
    "header_footer.header_rule": lambda v: v not in (None, "", "none"),
    "header_footer.page_number.position": lambda v: v == "odd-even-outside",
}

UNSUPPORTED_SPECIAL = ("signature_page", "seal_space", "amount_in_words",
                       "doc_control_page", "revision_history")


def _unsupported(cfg):
    problems = []
    for dotted, used in UNSUPPORTED.items():
        value = get_path(cfg, dotted)
        if used(value):
            problems.append(f"{dotted} = {value!r}")
    for key in UNSUPPORTED_SPECIAL:
        if cfg.get("special", {}).get(key):
            problems.append(f"special.{key} = true")
    if cfg.get("special", {}).get("classification"):
        problems.append("special.classification（密级标识）")
    return problems


def _collect_unverified(cfg, path=""):
    """收集标了 verified = false 的块。"""
    found = []
    if isinstance(cfg, dict):
        if cfg.get("verified") is False and path:
            found.append(path)
        for key, value in cfg.items():
            if key.startswith("_") or key in ("verified", "note"):
                continue
            if isinstance(value, dict):
                found.extend(_collect_unverified(value, f"{path}.{key}" if path else key))
            elif isinstance(value, list) and key == "headings":
                prefix = f"{path}.{key}" if path else key
                for item in value:
                    if isinstance(item, dict) and item.get("verified") is False:
                        found.append(f"{prefix}[{item.get('level')}]")
    return found


# ---------------------------------------------------------------- 字体
_FONT_DB = None


def _normalize_font_name(name):
    name = name.strip().lower()
    for suffix in (" (truetype)", " (opentype)", " (all res)", " & "):
        if name.endswith(suffix):
            name = name[: -len(suffix)]
    return name


def installed_fonts():
    """返回本机已安装字体名集合（小写）；无法判定时返回 None。"""
    global _FONT_DB
    if _FONT_DB is not None:
        return _FONT_DB
    if os.name != "nt":
        return None
    names = set()
    try:
        import winreg  # noqa: PLC0415  仅 Windows
        key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,
                             r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts")
        try:
            index = 0
            while True:
                try:
                    value_name, value_data, _ = winreg.EnumValue(key, index)
                except OSError:
                    break
                names.add(_normalize_font_name(value_name))
                names.add(_normalize_font_name(str(value_data)))
                index += 1
        finally:
            key.Close()
    except Exception:
        return None
    font_dir = os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts")
    if os.path.isdir(font_dir):
        for entry in os.listdir(font_dir):
            stem, _ = os.path.splitext(entry)
            names.add(stem.lower())
    _FONT_DB = names
    return names


def font_available(name):
    """字体是否可用；无法判定时返回 True（不阻断转换）。"""
    db = installed_fonts()
    if db is None:
        return True
    lowered = name.strip().lower()
    if lowered in db:
        return True
    for alias in FONT_ALIASES.get(name.strip(), []):
        if alias in db:
            return True
        if any(alias in entry for entry in db):
            return True
    return False


FONT_PATHS = [
    "body.font_cn", "body.font_en", "header_footer.header_font",
    "figure.caption_font", "table.cell_font", "table.caption.font",
]


def _font_paths(cfg):
    paths = list(FONT_PATHS)
    for heading in cfg.get("headings", []):
        paths.append(f"headings[{heading.get('level')}].font")
    return paths


def _font_value(cfg, dotted):
    if dotted.startswith("headings["):
        level = int(dotted.split("[")[1].split("]")[0])
        for heading in cfg.get("headings", []):
            if heading.get("level") == level:
                return heading.get("font")
        return None
    return get_path(cfg, dotted)


def _set_font_value(cfg, dotted, value):
    if dotted.startswith("headings["):
        level = int(dotted.split("[")[1].split("]")[0])
        for heading in cfg.get("headings", []):
            if heading.get("level") == level:
                heading["font"] = value
        return
    set_path(cfg, dotted, value)


def resolve_fonts(cfg):
    """按 fonts.fallback 链替换本机缺失字体，返回告警文本列表。"""
    notices = []
    fallback_map = cfg.get("fonts", {}).get("fallback", {}) or {}
    for dotted in _font_paths(cfg):
        name = _font_value(cfg, dotted)
        if not name or font_available(name):
            continue
        chain = [name] + list(fallback_map.get(name, []))
        replacement = next((c for c in chain[1:] if font_available(c)), None)
        if replacement:
            _set_font_value(cfg, dotted, replacement)
            notices.append(f"字体「{name}」本机未安装，已按 fallback 链替换为「{replacement}」（{dotted}）")
        else:
            notices.append(f"字体「{name}」本机未安装，且 fallback 链无可用替代，Word 将回退默认字体（{dotted}）")
    return notices


# ---------------------------------------------------------------- 加载
def _resolve_chain(spec):
    """沿 base 链自底向上收集配置，返回 (merged_without_defaults, chain_names)。"""
    layers = []
    seen = set()
    current = spec
    while current:
        path = find_profile(current)
        if path in seen:
            raise ProfileError(f"profile base 循环引用: {path}")
        seen.add(path)
        raw = _read_toml(path)
        raw["_source"] = path
        layers.append(raw)
        base = raw.get("base") or ""
        if not base:
            break
        current = base
    merged = {}
    for layer in reversed(layers):
        merged = _deep_merge(merged, layer)
    chain = [os.path.basename(layer["_source"]) for layer in layers]
    return merged, chain


def load_profile(spec=None, check_fonts=True):
    """加载并解析 profile。

    返回 (cfg, notices)：cfg 是填好默认值、字体已解析的字典；
    notices 是字体替换、未核实项等提示文本（调用方负责打印）。
    """
    raw, chain = _resolve_chain(spec or DEFAULT_PROFILE)
    errors = []
    _validate(raw, DEFAULTS, "", errors)
    if errors:
        raise ProfileError("profile 字段错误:\n  " + "\n  ".join(errors))

    cfg = _deep_merge(DEFAULTS, raw)
    _check_enums(cfg, errors)
    unsupported = _unsupported(cfg)
    if unsupported:
        errors.append("以下能力本版本尚未实现，请改用已支持的取值:\n    "
                      + "\n    ".join(unsupported))
    if errors:
        raise ProfileError("profile 不可用:\n  " + "\n  ".join(errors))

    cfg["_meta"] = {"chain": chain, "source": raw.get("_source", "")}

    notices = []
    unverified = _collect_unverified(cfg)
    if unverified:
        notices.append(f"profile「{cfg.get('name') or spec}」含 {len(unverified)} 处未核实取值: "
                       + ", ".join(unverified[:6]) + (" …" if len(unverified) > 6 else ""))
    if check_fonts:
        notices.extend(resolve_fonts(cfg))
    return cfg, notices


# ---------------------------------------------------------------- CLI
def _list_profiles():
    dirs = [os.path.join(REPO_ROOT, "profiles")]
    for extra in os.environ.get("WORD_DOC_PROFILES", "").split(os.pathsep):
        if extra:
            dirs.append(extra)
    rows = []
    for directory in dirs:
        if not os.path.isdir(directory):
            continue
        for name in sorted(os.listdir(directory)):
            if name.endswith(".toml"):
                path = os.path.join(directory, name)
                try:
                    raw = _read_toml(path)
                    label = raw.get("label", "")
                except Exception as exc:  # 坏文件也要列出来
                    label = f"<解析失败: {exc}>"
                rows.append((os.path.splitext(name)[0], label, path))
    return rows


def main(argv=None):
    import argparse
    parser = argparse.ArgumentParser(description="查看或校验排版 profile")
    parser.add_argument("--list", action="store_true", help="列出可用 profile")
    parser.add_argument("--show", metavar="NAME", help="打印合并后的完整参数")
    parser.add_argument("--check", metavar="NAME", help="只校验字段合法性")
    args = parser.parse_args(argv)

    if args.list or not (args.show or args.check):
        rows = _list_profiles()
        if not rows:
            print("未找到任何 profile")
            return 1
        width = max(len(name) for name, _, _ in rows)
        for name, label, path in rows:
            print(f"{name:<{width}}  {label}")
        return 0

    target = args.show or args.check
    try:
        cfg, notices = load_profile(target)
    except ProfileError as exc:
        print(f"错误: {exc}", file=sys.stderr)
        return 1
    for notice in notices:
        print(f"提示: {notice}", file=sys.stderr)
    if args.check:
        print(f"profile「{target}」字段校验通过")
        return 0
    _dump(cfg)
    return 0


def _dump(cfg):
    def walk(node, indent=0):
        pad = "  " * indent
        for key, value in node.items():
            if key.startswith("_"):
                continue
            if isinstance(value, dict):
                print(f"{pad}{key}:")
                walk(value, indent + 1)
            elif isinstance(value, list):
                if all(not isinstance(item, (dict, list)) for item in value):
                    print(f"{pad}{key} = {value!r}")
                    continue
                for item in value:
                    if isinstance(item, dict):
                        tag = f"[{item['level']}]" if "level" in item else ""
                        print(f"{pad}{key}{tag}:")
                        walk(item, indent + 1)
                    else:
                        print(f"{pad}{key} = {item!r}")
            else:
                print(f"{pad}{key} = {value!r}")
    walk(cfg)


if __name__ == "__main__":
    sys.exit(main())
