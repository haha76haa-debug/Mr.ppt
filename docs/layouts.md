# 版式库：13 种页面原型

所有函数在 `src/pptx_native/layouts.py`，统一签名：

```python
slide = layouts.xxx(prs, theme, **内容参数)
```

* 只产出**原生形状 + 真文字**（可改字、可改色、可换图）。
* 颜色/字体全部来自 `Theme`，换主题不改代码。
* 坐标单位英寸；内容区固定 `BODY_TOP=1.82 ~ BODY_BOTTOM=6.82`。

---

## 封面 `cover`

```python
layouts.cover(prs, theme, kicker="QUARTERLY REVIEW", title="季度回顾",
              subtitle="2026 Q1 · 交付与质量", meta="产品团队 · 2026-04")
```

深色渐变底 + 右下装饰圆（20% 透明）。标题 `hero 44`、导语 `lead 20`。
`kicker` 是标题上方的小标签（全大写、字距 3pt）。

**适用**：任何正式 deck。**忌**：封面堆联系方式和目录（目录单独一页）。

---

## 目录 `agenda`

```python
layouts.agenda(prs, theme, title="本次内容",
               items=["背景", "方法", "数据", "计划"])
```

编号 `01…`，≤4 项单列，≥5 项自动两列（列优先）。`kicker` 默认「议程」。

---

## 章节过渡 `section`

```python
layouts.section(prs, theme, number=2, title="方法与路径", subtitle="先对齐口径")
```

左侧 4.25in 浅色竖条 + 大序号（88pt），右侧标题与一句话导语。
每 3~5 页内容插一个过渡页，给观众喘息。

---

## 要点页 `bullets`

```python
layouts.bullets(prs, theme, title="我们要解决的问题", kicker="问题定义",
                items=[("信息分散", "数据散落在多个工具里"),
                       ("口径不一", "同一指标两种算法")])
```

每行一个圆角方块 + 小标题 + 一句说明，行高均分，行间细分隔线。
`items` 也可以是纯字符串（只有小标题）。

**原则**：一行 ≤ 2 行说明；超过就拆页。要点 3~5 个最佳。

---

## 对比页 `two_col`

```python
layouts.two_col(prs, theme, title="现状与目标",
                left=("现状", ["3 套口径", "手工周报"]),
                right=("目标", "单一口径 + 自动汇总"))
```

两张浅色卡片，左侧主色条、右侧次色条；正文可为字符串或字符串列表。

---

## 数据页 `stats`

```python
layouts.stats(prs, theme, title="关键数据", kicker="仪表盘",
              items=[("128", "个", "季度交付"), ("96", "%", "按时率")],
              note="数据来源：内部看板")
```

3~5 个数字并排：顶部细色条 + 大数字（46pt）+ 小单位（16pt）+ 灰色标签。
数字与单位拆两个 run，基线对齐。

**要点**：只放大真正重要的 3~5 个；其余数字进备注或表格。

---

## 时间线 `timeline`

```python
layouts.timeline(prs, theme, title="里程碑",
                 items=[("第 1 周", "立项", "对齐目标"),
                        ("第 4 周", "原型", "跑通关键路径")])
```

横向主线 + 均分节点（白心蓝环），节点上方时间、下方标题与说明。
节点等分，不按真实时间比例（时间跨度悬殊时另做甘特图）。

---

## 流程页 `process`

```python
layouts.process(prs, theme, title="四步落地",
                steps=[("对齐", "统一口径"), ("落地", "拆到人")])
```

N 张卡片 + 序号圆形徽标 + 卡片间箭头。3~5 步最佳。

---

## 图片墙 `image_grid`

```python
layouts.image_grid(prs, theme, title="现场", cols=3,
                   items=[("a.jpg", "开幕"), (None, "待补图"), ("b.jpg", "合影")])
```

等分网格，图片按 cover 裁切填满格子，底部半透明注释条。
`source=None` 生成虚线占位框——**先排版、后插图**，图片可替换。

跨渲染器注意：如果照片放在带 `scene3d` 的组合里，注释条必须同组（见 native-editing.md）。

---

## 引用页 `quote`

```python
layouts.quote(prs, theme, text="把复杂留给自己，把简单留给用户。", author="— 产品原则")
```

浅底 + 大引号 + 26pt 引言 + 短横线 + 出处。一页只放一句。

---

## 表格 `table`

```python
layouts.table(prs, theme, title="指标对照",
              header=["指标", "目标", "实际"],
              rows=[["交付", "≥120", "128"]],
              weight=[2, 1, 1], note="数据为示例")
```

原生表格：清掉模板样式，首行主色、隔行浅底、只留细下划线，列宽按 `weight` 分配。

**规则**：表格 ≤ 6 行 × 5 列；数字右对齐（本版左对齐，需要时自己加）；
表格里不要塞长句。

---

## 图表 `chart`

```python
layouts.chart(prs, theme, title="季度趋势", kind="column",
              categories=["Q1", "Q2", "Q3", "Q4"],
              series=[("计划", [24, 30, 35, 39]), ("实际", [22, 29, 34, 38])])
```

原生 Chart（右键「编辑数据」可改），`kind` 支持 `column / bar / line`。
单系列自动加数据标签；多系列显示底部图例；网格线用 `line` 色。

**规则**：柱状图从 0 起、别截断；折线 ≤ 3 条；不要 3D 饼图；
图例能删就删（系列名写进标题或标签）。

---

## 结尾 `end`

```python
layouts.end(prs, theme, title="谢谢", subtitle="欢迎提问", contact="team@example.com")
```

与封面同色系，收得干净。

---

## 页面节奏建议

| 顺序 | 版式 |
|---|---|
| 1 | `cover` |
| 2 | `agenda` |
| 3 | `section` 01 |
| 4~6 | `bullets` / `two_col` / `stats` |
| 7 | `section` 02 |
| 8~10 | `timeline` / `process` / `chart` / `table` / `image_grid` |
| 11 | `quote`（可选） |
| 12 | `end` |

同一份 deck 里，**同一种信息用同一种版式**：所有「数据」都用 `stats`，
所有「对比」都用 `two_col`。观众会因此更快读懂。
