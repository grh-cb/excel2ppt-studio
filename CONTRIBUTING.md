# 贡献指南

## 开发环境

```powershell
python -m pip install -e ".[dev,web]"
python -m pytest
ruff check src tests
```

要求 Python 3.10+。Windows 上预览能力需要 WPS Office（可选）。

## 分支与提交

- 分支：`feat/<短描述>`、`fix/<短描述>`、`docs/<短描述>`；不直推 `main`。
- 提交信息使用约定式：`feat(renderer): 支持表格动态增删行`、`fix(scanner): 组合形状坐标换算`。
- 一次提交只做一个逻辑变更。

## 测试要求

- 新增能力必须带 pytest；测试打在公共 seam 上（`template_scanner` / `renderer` / `mapping`）。
- 测试 fixtures **程序化生成**（python-pptx / openpyxl），不提交二进制样例，保证可读与可维护。
- 修复 bug 先写失败测试。

## 代码规范

- `ruff check` 必须通过；类型标注覆盖公共函数。
- 不允许静默吞异常：捕获后必须转成可读错误或告警。
- 失败必须可见：CLI 返回非零退出码 + 中文单行错误；降级路径打 WARNING。

## 文档同步

改动涉及以下内容时同步更新：

| 改动 | 文档 |
|---|---|
| 新术语 / 绑定类型 | `docs/ARCHITECTURE.md`、`CONTEXT.md`（如存在） |
| 架构决策 | `docs/adr/NNNN-*.md` |
| 用户可见能力 | `README.md`、`CHANGELOG.md` |

## 许可

提交即表示同意以 MIT 许可发布你的贡献。
