# -*- coding: utf-8 -*-
"""设计系统：主题色、字体、字号阶梯、12 栏网格。

版式代码（``layouts.py``）只依赖 ``Theme``。换主题 = 换一组常量，
版式一个字符都不用改；这是「同一套模板做一百份 PPT」的前提。

坐标一律用英寸（1 in = 914400 EMU），画布默认 13.333 x 7.5（16:9）。
"""
from __future__ import annotations

from dataclasses import asdict, dataclass, field, replace

from pptx.dml.color import RGBColor

#: 16:9 画布尺寸（英寸）
SLIDE_W = 13.333
SLIDE_H = 7.5

#: 字号阶梯（pt）。一页最多用 3~4 级，层级才立得住。
TYPE_SCALE = {
    "display": 88,      # 章节序号
    "hero": 44,         # 封面主标题
    "h1": 34,           # 页面标题
    "h2": 26,           # 引用 / 大段强调
    "h3": 19,           # 小标题
    "lead": 20,         # 导语
    "body": 16,         # 正文
    "small": 13,        # 辅助说明
    "caption": 11,      # 注释 / 页码
    "stat": 46,         # 关键数字
    "stat_unit": 16,    # 数字单位
}

_HEX = set("0123456789ABCDEFabcdef")
_ALIAS = {"dark": "primary_dark", "light": "primary_light", "white": "on_dark"}


def _is_hex(name: str) -> bool:
    return len(name) == 6 and all(c in _HEX for c in name)


@dataclass(frozen=True)
class Theme:
    """一套视觉规范。所有字段都是显式值，不依赖 PPT 主题继承。"""

    name: str = "经典蓝"
    # 颜色角色：只允许这几种，避免页面上出现第 5 种蓝
    primary: str = "0B52C4"
    primary_dark: str = "12325F"
    primary_light: str = "2E7BE8"
    accent: str = "E0A03C"
    text: str = "1F2329"
    muted: str = "6B7280"
    bg: str = "FFFFFF"
    surface: str = "F2F6FC"
    line: str = "DCE2EC"
    on_dark: str = "FFFFFF"
    # 字体：标题 / 正文；中文环境默认微软雅黑（Windows 自带）
    font_title: str = "Microsoft YaHei"
    font_body: str = "Microsoft YaHei"
    # 网格
    margin: float = 0.72
    gutter: float = 0.28
    cols: int = 12
    scale: dict = field(default_factory=lambda: dict(TYPE_SCALE))

    # ------------------------------------------------------------- 颜色
    def rgb(self, role):
        """角色名或 6 位十六进制 → ``RGBColor``。"""
        if isinstance(role, RGBColor):
            return role
        name = _ALIAS.get(role, role)
        value = name if _is_hex(name) else getattr(self, name)
        return RGBColor.from_string(value.upper())

    def tuple(self, role):
        """角色名 → ``(r, g, b)``，给 ``styling.grad`` 的 stops 用。"""
        value = str(self.rgb(role))
        return tuple(int(value[i:i + 2], 16) for i in (0, 2, 4))

    # ------------------------------------------------------------- 文字
    def size(self, key: str) -> int:
        return self.scale[key]

    def font(self, kind: str = "body") -> str:
        return self.font_title if kind in ("title", "heading") else self.font_body

    # ------------------------------------------------------------- 网格
    @property
    def content_width(self) -> float:
        return SLIDE_W - 2 * self.margin

    @property
    def content_height(self) -> float:
        return SLIDE_H - 2 * self.margin

    def col(self, start: int = 0, span: int = 1):
        """12 栏网格 → ``(x, w)``，单位英寸。``col(0, 12)`` 就是整幅内容宽度。"""
        if start < 0 or span < 1 or start + span > self.cols:
            raise ValueError("超出 %d 栏网格: start=%r span=%r" % (self.cols, start, span))
        unit = (self.content_width - (self.cols - 1) * self.gutter) / self.cols
        x = self.margin + start * (unit + self.gutter)
        w = span * unit + (span - 1) * self.gutter
        return x, w

    # ------------------------------------------------------------- 派生
    def with_fonts(self, title: str, body: str | None = None) -> "Theme":
        return replace(self, font_title=title, font_body=body or title)

    def as_dict(self) -> dict:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: dict) -> "Theme":
        known = {k: v for k, v in data.items() if k in cls.__dataclass_fields__}
        return cls(**known)

    # ------------------------------------------------------------- 预设
    @classmethod
    def blue(cls) -> "Theme":
        """经典蓝（默认）：稳重、通用，适合工作汇报与答辩。"""
        return cls()

    @classmethod
    def dark(cls) -> "Theme":
        """深蓝：发布会 / 路演，深底浅字。"""
        return cls(name="深海蓝", primary="1B3E78", primary_dark="0B1F3A",
                   primary_light="3D7FD6", accent="E8B04B", text="16202E",
                   muted="5F6B7A", surface="EEF3FA")

    @classmethod
    def minimal(cls) -> "Theme":
        """石墨 + 单一强调色：极简、克制，适合作品集与学术汇报。"""
        return cls(name="石墨", primary="2B2F33", primary_dark="151719",
                   primary_light="6E7378", accent="C8613A", text="1C1F22",
                   muted="70767C", surface="F4F4F2")

    @classmethod
    def warm(cls) -> "Theme":
        """暖赤陶：品牌 / 文化类内容，比蓝色更有温度。"""
        return cls(name="暖赤陶", primary="B5452A", primary_dark="6E2A1A",
                   primary_light="D9794F", accent="2F6E6A", text="2A2320",
                   muted="7A6B63", surface="FAF3EE")


__all__ = ["Theme", "TYPE_SCALE", "SLIDE_W", "SLIDE_H"]
