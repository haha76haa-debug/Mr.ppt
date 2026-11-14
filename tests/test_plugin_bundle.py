# -*- coding: utf-8 -*-
"""DSH 插件 bundle 测试：清单完整、构建产出包含 Python 包与导出脚本。"""
from __future__ import annotations

import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
PLUGIN = ROOT / "dsh-plugin"


def test_manifest_and_bundle_files():
    manifest = json.loads((PLUGIN / "package.json").read_text(encoding="utf-8"))
    assert manifest["name"] == "@local/pptx-studio"
    assert manifest["type"] == "module"
    assert manifest["dsh"]["bundle"]["patch"] == "./cordis.patch.yml"

    for rel in ("index.js", "cordis.patch.yml", "README.md", "icon.svg",
                "locale/en.json", "locale/zh.json"):
        assert (PLUGIN / rel).is_file(), rel

    patch = (PLUGIN / "cordis.patch.yml").read_text(encoding="utf-8")
    assert "@local/pptx-studio" in patch and "pptx-studio" in patch

    for locale in ("en", "zh"):
        data = json.loads((PLUGIN / "locale" / f"{locale}.json").read_text(encoding="utf-8"))
        assert data["meta"]["title"] == "PPT Studio"
        assert data["meta"]["description"]


def test_plugin_entry_registers_pptx_tool():
    source = (PLUGIN / "index.js").read_text(encoding="utf-8")
    for token in ("name: 'pptx'", "sample", "probe", "lint", "render", "themes",
                  "export const inject = ['tools']"):
        assert token in source, token


def test_build_plugin_assembles_bundle(tmp_path):
    sys.path.insert(0, str(ROOT / "scripts"))
    try:
        import build_plugin
    finally:
        sys.path.pop(0)

    out = build_plugin.build(tmp_path / "bundle")
    assert (out / "index.js").is_file()
    assert (out / "python" / "pptx_native" / "__init__.py").is_file()
    assert (out / "python" / "pptx_native" / "layouts.py").is_file()
    assert (out / "scripts" / "export_slide.ps1").is_file()
    assert (out / "locale" / "zh.json").is_file()
    # 构建产物里不应带缓存目录
    assert not list((out / "python").rglob("__pycache__"))
