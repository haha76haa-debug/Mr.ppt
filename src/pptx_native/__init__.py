# -*- coding: utf-8 -*-
"""pptx-native锛氭妸 PPT 鍋氭垚銆屽彲缂栬緫鍘熺敓褰㈢姸銆嶏紝骞剁敤 PowerPoint 鐪熸満鏍￠獙銆?
鍏紑 API 閮戒粠瀛愭ā鍧楀鍑猴紝鏂逛究鑴氭湰閲岀洿鎺?import銆?"""
__version__ = "0.2.1"

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

