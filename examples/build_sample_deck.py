# -*- coding: utf-8 -*-
"""示例：用版式库搭一份 12 页 deck。

    python examples/build_sample_deck.py build/sample.pptx

等价于 ``pptx-native sample build/sample.pptx``。
不依赖任何外部素材：图片墙用「图片占位框」，换成自己的图即可。
"""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from pptx_native.sample import build_sample  # noqa: E402


def main(argv=None):
    argv = argv or sys.argv[1:]
    out = argv[0] if argv else "build/sample.pptx"
    print("sample deck:", build_sample(out))
    return 0


if __name__ == "__main__":
    sys.exit(main())
