# -*- coding: utf-8 -*-
"""pptx-native：把 PPT 做成「可编辑原生形状」，并用 PowerPoint 真机校验。

公开 API 都从子模块导出，方便脚本里直接 import。
"""
__version__ = "0.2.0"

from . import (checks, colors, geometry, groups, images, layouts, probe,  # noqa: F401
               render, sample, styling, theme)
from .geometry import abs_rect, iter_shapes, to_child_space  # noqa: F401
from .styling import (fill_text, grad, layered_text, set_alpha_fill,  # noqa: F401
                      set_font, shape, textbox, vertical_label)
from .theme import SLIDE_H, SLIDE_W, TYPE_SCALE, Theme  # noqa: F401

__all__ = [
    "checks", "colors", "geometry", "groups", "images", "layouts", "probe",
    "render", "sample", "styling", "theme",
    "abs_rect", "iter_shapes", "to_child_space",
    "fill_text", "grad", "layered_text", "set_alpha_fill", "set_font", "shape",
    "textbox", "vertical_label",
    "Theme", "TYPE_SCALE", "SLIDE_W", "SLIDE_H",
    "__version__",
]
