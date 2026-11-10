# Codex 审查报告：数据层（第四轮）

> 工具：`codex exec --json`，输入 21.4 万 token，输出 1.9 万 token
> 范围：`excel_reader.py`、`excel_images.py`、`titles.py`、`mapping.py`
> 结论：高优先级 8 条、中优先级 13 条

## 高优先级

| # | 问题 | 位置 | 修复要点 |
|---|---|---|---|
| 1 | **非首行表头在 Profile 中丢失**（已核实） | `mapping.py:132-138` | `build_profile()` 未传 `header_row`，Profile 恒为 1；调用方（CLI/Web）需传入并写进 Profile |
| 2 | 合并单元格静默丢数据 | `excel_reader.py:106,121` | `data_only+read_only` 下合并区只有左上角有值；应检测 `mergeCells`：多级表头直接拒绝，纵向合并按策略填充并告警 |
| 3 | 公式无缓存值被当空单元格 | `excel_reader.py:106` | 程序生成、未重算的文件会返回 `None`；应加公式模式扫描，对"有公式无缓存"告警而非静默 |
| 4 | LLM 返回缺项/重复项被当作完全成功 | `titles.py:97,70` | 每批检查键集合与请求一致；重复序号先拒绝；`null` 不应转成 `"None"` |
| 5 | LLM 调用无重试、无总时限 | `titles.py:33,46,96` | 增加 `timeout/max_attempts/total_deadline`，只对 408/429/5xx/网络错误退避重试 |
| 6 | ZIP 中 ID 与关系按位置 `zip` 配对 | `excel_images.py:39,47` | 改为 XML 解析：以图片对象为单元读 name + r:embed；校验数量与重复 |
| 7 | ZIP 异常未包装、无资源限制 | `excel_images.py:27,38,58` | `BadZipFile/RuntimeError/UnicodeDecodeError` 统一转 `ExcelImageError`；限制成员数/单成员大小/总解压量 |
| 8 | Profile 保存非原子 | `mapping.py:142-146` | 内存序列化 → 同目录唯一临时文件 → `fsync` → `os.replace` |

## 中优先级（摘要）

- #9 `max_column` 可为 `None` / 被 dimension 放大 / 小于真实数据时静默截断。
- #10 重复表头用 `names.count()` 是 O(n²)，改 `Counter`。
- #11 `read_dispimg_formulas` 完全未包装错误，空工作簿会裸 `KeyError`。
- #12 全量读入 `raw_rows`，无行数/规模上限。
- #13 `\{.*\}` 贪婪正则，多个 JSON 块或标题含花括号会解析失败。
- #14 成功响应非法 JSON / 编码错误未包装；响应体无大小上限。
- #15 `batch_size=0` 抛原始 `ValueError`；空任务被 CLI 当成功。
- #16 `load_titles` 未包装 I/O/JSON 错误，`null`/数组被 `str()` 成怪值。
- #17 `load_profile` 只捕 `JSONDecodeError`，`Profile.from_dict` 的 `KeyError/TypeError` 会泄漏；未校验 version/header_row/重复 field_id。
- #18 `assert field.table is not None` 用于生产控制流，`-O` 下失效。
- #19 所有图片元素自动绑第一列（模板两个图片位时静默重复）。
- #20 请求了但 Excel 里不存在的 ID 不计入 skipped，静默当"无图"；`__skipped__` 键污染接口。
- #21 `list_dispimg_ids()` 注释说返回 ID→成员名，实际值恒为空串。

## 与前三轮的交叉

| 结论 | 首轮 | 本轮 |
|---|---|---|
| `ws.max_column` 依赖 | P2-9 | #9 ✅ |
| `excel_images` 解析脆弱 | P2-10 | #6/#7/#20/#21 ✅加深 |
| `titles` 无重试/JSON 脆弱/base_url 风险 | P2-11 | #4/#5/#13/#14/#15/#16 ✅加深 |
| `save_profile` 非原子 | P2-13 | #8 ✅一致 |
| 图片自动绑第一列 | P2-13 | #19 ✅一致 |
| **`header_row` 丢失** | 未发现 | **#1 新增（已核实）** |
| **合并单元格** | 未发现 | **#2 新增** |
| **公式无缓存值** | 未发现 | **#3 新增** |
