# -*- coding: utf-8 -*-
"""生成一份中性示例 deck：演示版式库怎么用，内容全部是可替换的占位。

    pptx-native sample demo.pptx --theme blue

不依赖任何外部素材：图片墙用「图片占位框」，换成自己的图即可。
"""
from __future__ import annotations

from pathlib import Path

from . import layouts
from .theme import Theme

#: 命令行里可选的预设主题
THEMES = {
    "blue": Theme.blue,
    "dark": Theme.dark,
    "minimal": Theme.minimal,
    "warm": Theme.warm,
}

SLIDES = [
    "cover", "agenda", "section", "bullets", "two_col", "stats",
    "timeline", "process", "image_grid", "table", "chart", "end",
]


def build_sample(out="sample.pptx", theme: Theme | None = None) -> str:
    """按版式库搭一份 12 页示例 deck，返回输出路径。"""
    theme = theme or Theme.blue()
    out_path = Path(out)
    if str(out_path.parent) not in ("", ".") and not out_path.parent.exists():
        out_path.parent.mkdir(parents=True, exist_ok=True)

    prs = layouts.new_deck(theme)

    layouts.cover(
        prs, theme,
        kicker="PPTX-NATIVE · 通用示例",
        title="示例演示文稿",
        subtitle="用 12 栏网格和原生形状搭一套可编辑的页面",
        meta="所有内容均为占位，可直接替换成你的材料",
    )

    layouts.agenda(
        prs, theme, title="本次内容",
        items=["背景与目标", "方法与路径", "关键数据", "里程碑", "下一步"],
        page=2,
    )

    layouts.section(prs, theme, number=1, title="背景与目标",
                    subtitle="先把问题说清楚，再谈方案", page=3)

    layouts.bullets(
        prs, theme, title="我们要解决的问题", kicker="问题定义", page=4,
        items=[
            ("信息分散", "数据散落在表格、文档和聊天记录里，找一次要十分钟。"),
            ("口径不一", "同一个指标两种算法，开会先花半小时对齐定义。"),
            ("复盘困难", "结果没有留痕，出了问题定位不到具体环节。"),
        ],
    )

    layouts.two_col(
        prs, theme, title="现状与目标", kicker="对比", page=5,
        left=("现状", ["3 套并存的口径", "周报靠手工汇总", "复盘平均滞后 2 周"]),
        right=("目标", ["单一指标定义", "周报自动生成", "次日完成复盘"]),
    )

    layouts.stats(
        prs, theme, title="关键数据", kicker="仪表盘", page=6,
        items=[("128", "个", "季度交付需求"), ("96", "%", "按时交付率"),
               ("4", "个", "并行迭代"), ("12", "个", "里程碑达成")],
        note="以上数字均为示例占位。",
    )

    layouts.timeline(
        prs, theme, title="里程碑", kicker="计划", page=7,
        items=[("第 1 周", "立项", "明确目标与验收标准"),
               ("第 4 周", "原型", "关键路径跑通"),
               ("第 8 周", "上线", "灰度 10% 用户"),
               ("第 12 周", "复盘", "输出可执行改进项")],
    )

    layouts.process(
        prs, theme, title="四步落地", kicker="方法", page=8,
        steps=[("对齐", "统一指标定义与验收口径"),
               ("落地", "按里程碑拆解到人"),
               ("监控", "每周看板自动汇总"),
               ("复盘", "输出可执行的改进项")],
    )

    layouts.image_grid(
        prs, theme, title="图片墙", kicker="素材", page=9, cols=3,
        items=[(None, "图片 01"), (None, "图片 02"), (None, "图片 03"),
               (None, "图片 04"), (None, "图片 05"), (None, "图片 06")],
    )

    layouts.table(
        prs, theme, title="指标对照", kicker="表格", page=10,
        header=["指标", "目标", "实际"],
        rows=[["交付需求", "≥ 120 个", "128 个"],
              ["按时交付率", "≥ 95%", "96%"],
              ["平均周期", "≤ 10 天", "9 天"],
              ["复盘滞后", "≤ 3 天", "2 天"]],
        weight=[2, 1.2, 1.2],
        note="数据为示例占位。",
    )

    layouts.chart(
        prs, theme, title="季度趋势", kicker="图表", page=11, kind="column",
        categories=["Q1", "Q2", "Q3", "Q4"],
        series=[("计划", [24, 30, 35, 39]), ("实际", [22, 29, 34, 38])],
        note="右键图表 → 编辑数据，即可替换成真实数据。",
    )

    layouts.end(
        prs, theme,
        subtitle="版式当函数用：换主题、换内容，不改结构。",
        contact="pptx-native · MIT License",
    )

    prs.save(str(out_path))
    return str(out_path)


if __name__ == "__main__":       # pragma: no cover
    print(build_sample())
