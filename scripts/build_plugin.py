# -*- coding: utf-8 -*-
"""把仓库源码组装成可安装的 DSH 插件 bundle。

    python scripts/build_plugin.py
    python scripts/build_plugin.py --out "C:/path/to/pptx-studio-bundle"

``dsh-plugin/`` 是 bundle 的源码（清单、patch、Host 插件、locale、图标）；
本脚本把 ``src/pptx_native/`` 与 ``scripts/export_slide.ps1`` 复制进去，
保证 Python 包只有一份源码，不在仓库里重复提交。
"""
from __future__ import annotations

import argparse
import json
import shutil
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
PLUGIN_DIR = ROOT / "dsh-plugin"
PACKAGE_DIR = ROOT / "src" / "pptx_native"
EXPORT_SCRIPT = ROOT / "scripts" / "export_slide.ps1"
IGNORE = shutil.ignore_patterns("__pycache__", "*.pyc", ".pytest_cache")


def build(out) -> Path:
    """组装 bundle 到 ``out``，返回该目录。"""
    out = Path(out).resolve()
    if out.exists():
        shutil.rmtree(out)
    shutil.copytree(PLUGIN_DIR, out)

    shutil.copytree(PACKAGE_DIR, out / "python" / "pptx_native", ignore=IGNORE)
    (out / "scripts").mkdir(parents=True, exist_ok=True)
    shutil.copy2(EXPORT_SCRIPT, out / "scripts" / "export_slide.ps1")

    manifest = json.loads((out / "package.json").read_text(encoding="utf-8"))
    assert manifest["name"] == "@local/pptx-studio", manifest
    assert manifest["dsh"]["bundle"]["patch"] == "./cordis.patch.yml"
    assert (out / "python" / "pptx_native" / "__init__.py").is_file()
    assert (out / "scripts" / "export_slide.ps1").is_file()
    assert not list((out / "python").rglob("__pycache__"))
    return out


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="组装 DSH 插件 bundle")
    parser.add_argument("--out", default=str(ROOT / "build" / "pptx-studio-bundle"),
                        help="输出目录（默认 build/pptx-studio-bundle）")
    args = parser.parse_args(argv)
    out = build(args.out)
    files = sum(1 for p in out.rglob("*") if p.is_file())
    print("bundle:", out)
    print("files:", files)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
