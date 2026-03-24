# 发布隔离流程

## 规则

1. **`main` 分支 = Streamlit Cloud 自动部署入口**，任何推送到 `main` 的代码会立即上线。
2. **禁止直接 `git push --force` 到 `main`**，除非已在本对话中明确核实了所有改动。
3. 日常开发在 **feature / staging 分支** 上进行，验证通过后通过 PR 合并。

## 验证最新版的方法

- 打开页面后，标题下方会显示 **build tag**（格式：`build <sha> · <UTC时间>`）。
- 与最近一次 `git log --oneline -1` 的 SHA 对比即可确认是否为最新版本。
- 使用无痕窗口（`Cmd+Shift+N`）避免缓存干扰。

## 被覆盖后的恢复

```bash
git fetch origin
git log --oneline -10          # 确认被覆盖的提交是否还在
git reset --hard <目标SHA>     # 回到正确版本
git push --force origin main   # 恢复线上
```

## 长期建议

- 为"正式版"和"测试版"各自创建独立的 Streamlit Cloud App（指向不同分支）。
- 或在同一仓库使用 `main`（正式）和 `staging`（测试）两个分支，各自绑定各自的 Streamlit App。
