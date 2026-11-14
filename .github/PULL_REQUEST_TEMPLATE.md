## 这个 PR 做了什么

<!-- 一句话说明，可关联 Issue: closes #12 -->

## 改动类型

- [ ] 🐞 Bug 修复
- [ ] ✨ 新功能（新版式 / 新命令）
- [ ] 📝 文档
- [ ] ✅ 测试
- [ ] 🔧 重构 / 杂项

## 怎么验证的

```
# 贴命令与结果
pytest -q
python scripts/selfcheck.py
pptx-native sample build/sample.pptx
```

- [ ] 测试全绿（pytest 或 selfcheck）
- [ ] 涉及几何/颜色的改动跑了 `pptx-native diff`，确认没有误伤
- [ ] 如涉及真机渲染，附 PowerPoint 导出的截图
- [ ] 更新了 `CHANGELOG.md`
- [ ] 更新了 README / docs（如新增或修改 CLI）

## 备注

<!-- 有没有已知限制、后续要做的 -->
