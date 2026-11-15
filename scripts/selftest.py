#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""管线自检：不需要 Word、不联网，几十秒内验证「转换 → 校验」这条链路没坏。

覆盖四类约定：
- 编码：UTF-8 BOM 的 md 转换时忽略、check_md 必须报错；变量文件带 BOM 能用；
- 表格：列数与表头不一致时补齐并告警、check_md 报错；缺 |---| 分隔行两边都告警；
- 版式：header-right 页码不覆盖页眉项目名、图注表注上下位置、TOC 深度、
  --page-size 覆盖、--no-header/--no-footer 豁免、坏图/缺图不中断；
- profile：废弃键 start_at 报错并给替代、max_depth 越界报错；
- 模板接管：{{正文}} 只认正文里的锚点、变量替换、BOM 变量文件。

用法：
    python scripts/selftest.py          # 跑全部，失败时退出码 1
    python scripts/selftest.py --keep   # 保留中间产物，便于排查

改脚本后先跑它，再跑真实文档；它失败说明是脚本坏了，不是文档坏了。
"""

import argparse
import os
import shutil
import struct
import subprocess
import sys
import sysconfig
import tempfile
import zlib

from docx import Document

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PY = sys.executable
# pip 安装后没有仓库结构：示例文档从 share 目录取，临时目录放系统 temp
INSTALLED = not os.path.isdir(os.path.join(REPO_ROOT, "scripts"))
ROOT = REPO_ROOT if not INSTALLED else tempfile.mkdtemp(prefix="wdp-selftest-")
EXAMPLE = os.path.join(REPO_ROOT, "examples", "示例文档.md")
if INSTALLED:
    _share = os.path.join(sysconfig.get_path("data"), "share", "word-doc-pipeline")
    EXAMPLE = os.path.join(_share, "examples", "示例文档.md")

# 安装后脚本没有 .py 路径，改用 console script 入口（找不到则退回调用 cli 模块）
ENTRIES = {
    "md2docx.py": ("wdp-md2docx", "md2docx"),
    "check_md.py": ("wdp-check-md", "check_md"),
    "check_docx.py": ("wdp-check-docx", "check_docx"),
    "profile.py": ("wdp-profile", "profile"),
    "build_docs.py": ("wdp-build-docs", "build_docs"),
    "fill_template.py": ("wdp-fill-template", "fill_template"),
}


def _installed_prefix(script):
    exe, func = ENTRIES[os.path.basename(script)]
    path = shutil.which(exe)
    if not path:
        local = os.path.join(os.path.dirname(sys.executable), exe + (".exe" if os.name == "nt" else ""))
        path = local if os.path.isfile(local) else None
    if path:
        return [path]
    code = f"import sys; from word_doc_pipeline.cli import {func}; sys.exit({func}())"
    return [PY, "-c", code]

RESULTS = []
TMP = None


# ---------------------------------------------------------------- 基础设施

def check(label, ok, detail=""):
    RESULTS.append((label, bool(ok), detail))
    print(("  ok   " if ok else "  FAIL ") + label + (f"\n         {detail}" if not ok else ""))


def cli(*args):
    """跑一个脚本，返回 (退出码, 合并后的输出)。"""
    args = list(args)
    if INSTALLED:
        args = [*_installed_prefix(args[0]), *args[1:]]
    else:
        args = [PY, *args]
    env = dict(os.environ, PYTHONUTF8="1")
    proc = subprocess.run(args, cwd=ROOT, capture_output=True, text=True,
                          encoding="utf-8", errors="replace", env=env)
    return proc.returncode, (proc.stdout or "") + (proc.stderr or "")


def write(path, text, bom=False):
    data = text.encode("utf-8")
    if bom:
        data = b"\xef\xbb\xbf" + data      # 记事本的默认写法
    with open(path, "wb") as f:
        f.write(data)


def make_png(path, width=60, height=30):
    """手写一张最小 PNG，免得为自检引入 Pillow 依赖。"""
    def chunk(tag, data):
        return (struct.pack(">I", len(data)) + tag + data
                + struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF))

    raw = b"".join(b"\x00" + b"\xcc\x33\x33" * width for _ in range(height))
    with open(path, "wb") as f:
        f.write(b"\x89PNG\r\n\x1a\n"
                + chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0))
                + chunk(b"IDAT", zlib.compress(raw))
                + chunk(b"IEND", b""))


def profile_file(name, extra=""):
    path = os.path.join(TMP, name + ".toml")
    write(path, f'base = "ug-thesis-bit"\nname = "{name}"\n{extra}')
    return path


# ---------------------------------------------------------------- 编码

def test_encoding():
    print("\n[编码]")
    md = os.path.join(TMP, "bom.md")
    write(md, "# 第一章 标题\n\n正文一段。\n", bom=True)
    rc, out = cli("scripts/md2docx.py", md, os.path.join(TMP, "bom.docx"))
    first = Document(os.path.join(TMP, "bom.docx")).paragraphs[0]
    check("带 BOM 的 md 能转换", rc == 0, out[-400:])
    check("BOM 不影响首个标题", first.style.name.startswith("Heading") and first.text == "第一章 标题",
          f"{first.style.name!r} {first.text!r}")
    check("转换时提示已忽略 BOM", "BOM" in out, out[-300:])

    rc, out = cli("scripts/check_md.py", md)
    check("check_md 对 BOM 报错", rc == 1 and "BOM" in out, out[-400:])


# ---------------------------------------------------------------- 表格

def test_tables():
    print("\n[表格]")
    md = os.path.join(TMP, "cols.md")
    write(md, "# 第一章 标题\n\n**表 1-1 列数不齐**\n\n"
              "| a | b | c |\n| --- | --- | --- |\n| 1 | 2 |\n| 1 | 2 | 3 |\n")
    rc, out = cli("scripts/check_md.py", md)
    check("check_md 报列数不一致", rc == 1 and "表头 3 列" in out, out[-400:])
    rc, out = cli("scripts/md2docx.py", md, os.path.join(TMP, "cols.docx"))
    table = Document(os.path.join(TMP, "cols.docx")).tables[0]
    check("列数不齐不会让转换崩溃", rc == 0, out[-400:])
    check("缺的那列被补齐", len(table.columns) == 3 and len(table.rows) == 3,
          f"{len(table.columns)}x{len(table.rows)}")
    check("补齐时给出告警", "列" in out and "补齐" in out, out[-400:])

    md = os.path.join(TMP, "nosep.md")
    write(md, "# 第一章 标题\n\n| a | b |\n| 1 | 2 |\n")
    rc, out = cli("scripts/check_md.py", md)
    check("check_md 对缺分隔行告警", "分隔行" in out, out[-400:])
    rc, out = cli("scripts/md2docx.py", md, os.path.join(TMP, "nosep.docx"))
    check("缺分隔行不生成表格", len(Document(os.path.join(TMP, "nosep.docx")).tables) == 0)
    check("转换时提示缺分隔行", "分隔行" in out, out[-400:])


# ---------------------------------------------------------------- 版式

def test_layout():
    print("\n[版式]")
    # header-right：页眉项目名与页码必须共存
    prof = profile_file("hdr-right", '[header_footer.page_number]\nposition = "header-right"\n')
    rc, out = cli("scripts/md2docx.py", EXAMPLE, os.path.join(TMP, "hdr.docx"),
                  "--profile", prof, "--project", "项目名X")
    para = Document(os.path.join(TMP, "hdr.docx")).sections[0].header.paragraphs[0]
    check("header-right 能转换", rc == 0, out[-400:])
    check("页眉保留项目名", "项目名X" in para.text, para.text)
    check("页眉同一段里有 PAGE 域", "PAGE" in para._p.xml)
    check("页眉有右对齐制表位", "tab" in para._p.xml)
    rc, out = cli("scripts/check_docx.py", os.path.join(TMP, "hdr.docx"),
                  "--profile", prof, "--project", "项目名X")
    check("header-right 产物通过校验", rc == 0, out[-500:])

    # 表注在表格下方（作者仍写在表格上方）
    prof = profile_file("cap", '[numbering.table]\ncaption = "below"\n')
    md = os.path.join(TMP, "cap.md")
    write(md, "# 第一章 标题\n\n**表 1-1 表注在下**\n\n| a | b |\n| --- | --- |\n| 1 | 2 |\n")
    rc, out = cli("scripts/md2docx.py", md, os.path.join(TMP, "cap.docx"), "--profile", prof)
    order = []
    for element in Document(os.path.join(TMP, "cap.docx")).element.body.iterchildren():
        tag = element.tag.split("}")[-1]
        if tag == "tbl":
            order.append("table")
        elif tag == "p":
            text = "".join(n.text or "" for n in element.iter() if n.tag.endswith("}t")).strip()
            if text:
                order.append(text)
    check("表注渲染在表格下方", rc == 0 and "table" in order
          and order.index("表 1-1 表注在下") > order.index("table"), str(order))

    # 图注在上 / 在下
    make_png(os.path.join(TMP, "pic.png"))
    prof = profile_file("fig-up", '[numbering.figure]\ncaption = "above"\n')
    md = os.path.join(TMP, "fig.md")
    write(md, "# 第一章 标题\n\n![图 1-1 图注在上](pic.png)\n")
    rc, out = cli("scripts/md2docx.py", md, os.path.join(TMP, "fig-up.docx"), "--profile", prof)
    texts = [p.text.strip() for p in Document(os.path.join(TMP, "fig-up.docx")).paragraphs if p.text.strip()]
    check("图注在上", rc == 0 and len(texts) >= 2 and texts[1] == "图 1-1 图注在上", str(texts))
    rc, out = cli("scripts/md2docx.py", md, os.path.join(TMP, "fig-down.docx"))
    texts = [p.text.strip() for p in Document(os.path.join(TMP, "fig-down.docx")).paragraphs if p.text.strip()]
    check("默认图注在下", rc == 0 and texts and texts[-1] == "图 1-1 图注在上", str(texts))

    # 损坏的图片只告警，不中断整篇
    write(os.path.join(TMP, "broken.png"), "not a png")
    md = os.path.join(TMP, "badimg.md")
    write(md, "# 第一章 标题\n\n![图 1-1 坏图](broken.png)\n\n正文仍在。\n")
    rc, out = cli("scripts/md2docx.py", md, os.path.join(TMP, "badimg.docx"))
    body = "\n".join(p.text for p in Document(os.path.join(TMP, "badimg.docx")).paragraphs)
    check("坏图不中断转换", rc == 0 and "正文仍在" in body, out[-400:])
    check("坏图给出告警", "无法解析" in out, out[-400:])

    md = os.path.join(TMP, "missimg.md")
    write(md, "# 第一章 标题\n\n![图 1-1 缺图](nope.png)\n")
    rc, out = cli("scripts/md2docx.py", md, os.path.join(TMP, "missimg.docx"))
    check("缺图只告警", rc == 0 and "不存在" in out, out[-400:])

    # TOC 深度跟随 numbering.max_depth
    rc, out = cli("scripts/md2docx.py", EXAMPLE, os.path.join(TMP, "toc4.docx"),
                  "--profile", "cn-gov", "--toc")
    xml = Document(os.path.join(TMP, "toc4.docx")).element.xml
    check("TOC 深度取 max_depth（cn-gov = 4）", rc == 0 and 'TOC \\o "1-4"' in xml, out[-300:])
    rc, out = cli("scripts/md2docx.py", EXAMPLE, os.path.join(TMP, "toc3.docx"), "--toc")
    xml = Document(os.path.join(TMP, "toc3.docx")).element.xml
    check("TOC 深度默认 1-3", rc == 0 and 'TOC \\o "1-3"' in xml, out[-300:])

    # --page-size 覆盖后，校验按生成标记里的尺寸判断
    rc, out = cli("scripts/md2docx.py", EXAMPLE, os.path.join(TMP, "b5.docx"), "--page-size", "B5")
    rc2, out2 = cli("scripts/check_docx.py", os.path.join(TMP, "b5.docx"))
    check("--page-size 覆盖不会被误判", rc == 0 and rc2 == 0, out2[-500:])

    # --no-header / --no-footer 的豁免
    rc, out = cli("scripts/md2docx.py", EXAMPLE, os.path.join(TMP, "nohf.docx"),
                  "--no-header", "--no-footer")
    rc2, out2 = cli("scripts/check_docx.py", os.path.join(TMP, "nohf.docx"))
    check("--no-header/--no-footer 豁免生效", rc == 0 and rc2 == 0, out2[-500:])


# ---------------------------------------------------------------- profile

def test_profile():
    print("\n[profile]")
    prof = profile_file("stale", "[header_footer.page_number]\nstart_at = 1\n")
    rc, out = cli("scripts/profile.py", "--check", prof)
    check("废弃键 start_at 报错", rc != 0, out[-400:])
    check("报错里给出替代键", "body_start_at" in out, out[-400:])

    prof = profile_file("deep", "[numbering]\nmax_depth = 7\n")
    rc, out = cli("scripts/profile.py", "--check", prof)
    check("max_depth 越界报错", rc != 0 and "max_depth" in out, out[-400:])

    for name in ("ug-thesis-bit", "thesis-bit", "tech-doc", "cn-gov"):
        rc, out = cli("scripts/profile.py", "--check", name)
        check(f"profile {name} 通过校验", rc == 0, out[-400:])


# ---------------------------------------------------------------- 模板接管

def test_template():
    print("\n[模板接管]")
    template = os.path.join(TMP, "tpl.docx")
    doc = Document()
    doc.add_paragraph("{{项目名}}")
    doc.add_paragraph("{{正文}}")
    doc.add_paragraph("落款：{{项目名}}")
    doc.sections[0].header.paragraphs[0].text = "{{正文}}"   # 陷阱：页眉里的锚点不算
    doc.save(template)

    var_file = os.path.join(TMP, "vars.toml")
    write(var_file, '"项目名" = "张三"\n', bom=True)
    out_docx = os.path.join(TMP, "filled.docx")
    rc, out = cli("scripts/fill_template.py", template, EXAMPLE, out_docx,
                  "--vars", var_file)
    doc = Document(out_docx)
    texts = [p.text.strip() for p in doc.paragraphs]
    check("模板填充能跑通", rc == 0, out[-400:])
    check("变量已替换", "张三" in texts, str(texts))
    check("带 BOM 的变量文件可用", "张三" in texts, str(texts))
    check("正文注入正文锚点处", any("示例文档" in t for t in texts), str(texts[:6]))
    check("页眉里的 {{正文}} 不当锚点",
          "示例文档" not in doc.sections[0].header.paragraphs[0].text,
          doc.sections[0].header.paragraphs[0].text)


# ---------------------------------------------------------------- 批量

def test_batch():
    print("\n[批量]")
    batch = os.path.join(TMP, "batch")
    os.makedirs(batch, exist_ok=True)
    write(os.path.join(batch, "doc1.md"), "# 第一章 标题\n\n正文。\n")
    write(os.path.join(batch, "doc2.md"), "# 第一章 标题\n\n正文二。\n")
    rc, out = cli("scripts/build_docs.py", batch, "--profile", "tech-doc", "--no-check")
    check("build_docs --no-check 批量转换", rc == 0
          and os.path.isfile(os.path.join(batch, "doc1.docx")), out[-400:])
    rc, out = cli("scripts/build_docs.py", batch, "--profile", "tech-doc")
    check("build_docs 转完自动校验", rc == 0, out[-400:])


# ---------------------------------------------------------------- 入口

def main():
    global TMP
    parser = argparse.ArgumentParser(description="管线自检")
    parser.add_argument("--keep", action="store_true", help="保留中间产物")
    args = parser.parse_args()

    TMP = os.path.join(ROOT, "tmp", f"selftest-{os.getpid()}")
    os.makedirs(TMP, exist_ok=True)
    try:
        test_encoding()
        test_tables()
        test_layout()
        test_profile()
        test_template()
        test_batch()
    finally:
        if args.keep:
            print(f"\n中间产物保留在 {TMP}")
        else:
            shutil.rmtree(TMP, ignore_errors=True)
    if INSTALLED and not args.keep:
        shutil.rmtree(ROOT, ignore_errors=True)

    failed = [r for r in RESULTS if not r[1]]
    print(f"\n{len(RESULTS) - len(failed)}/{len(RESULTS)} 项通过")
    for label, _, detail in failed:
        print(f"  FAIL {label}")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
