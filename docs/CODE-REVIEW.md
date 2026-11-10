# 代码审查报告（v0.6.0）

> **第二轮补充**：Codex（`codex exec`，DeepSeek flash）另做了一轮审查，见
> [`CODE-REVIEW-CODEX.md`](CODE-REVIEW-CODEX.md)。它额外发现并**实证**了 2 条本报告未覆盖的问题：
> 1. `a:srcRect` 子元素顺序错误（真实输出每页 `blipFill` 都是 `blip → stretch → srcRect`，标准应为 `blip → srcRect → stretch`）——列为 P0；
> 2. SmartArt 的 `dgm:relIds`（`r:dm/r:lo/r:qs/r:cs`）未重映射；以及 `value_map` 中 `0`/`False` 被误判未命中。
> 两轮报告在 partname 重名、并发 `.tmp`、路径绕过、槽位静默、图表静默上**结论一致**。

> 审查方式：独立审查 agent（只读），逐行阅读核心模块，并用已安装的 python-pptx 1.0.2 / openpyxl 源码核实行为、对仓库内产物做只读量化。
> 审查范围：`renderer.py`、`template_scanner.py`、`verifier.py`、`titles.py`、`web/app.py`、`excel_reader.py`、`mapping.py` 及其依赖模块。
> 结论：**1 个 P0（必现型输出损坏）、10 个 P1、15 个 P2**。

---

## P0 阻塞

### P0-1 模板 slideN 编号有空洞时，克隆页与保留页 partname 重名

- 位置：`renderer.py:131`（`add_slide`）、`renderer.py:619`（`_drop_slide`）
- 原因：python-pptx 的 `_next_slide_partname` 只按 `len(sldIdLst)+1` 命名，不检查重名；模板本身编号有空洞（如删过中间页）时会撞上已存在的 part。
- 已复现：3 页模板删中间页 → 输出 zip 出现重复条目 `ppt/slides/slide3.xml`，后写覆盖前写，被保留的模板页内容错乱，图片/图表 rels 也被顶掉。
- `render()` 仍返回成功，`verifier` 只看 shape/几何/字号，发现不了。
- 修复：
  ```python
  dest = prs.slides.add_slide(src_slide.slide_layout)
  dest.part.partname = prs.part.package.next_partname("/ppt/slides/slide%d.xml")
  ```
  保存前断言 partname 唯一；补"带空洞模板"的回归测试。

---

## P1 应修

| 编号 | 问题 | 位置 | 修复要点 |
|---|---|---|---|
| P1-1 | 图片临时目录在多数异常路径泄漏 | `renderer.py:455,530,632` | 清理放进覆盖整个渲染的 `try/finally` / `ExitStack` |
| P1-2 | 输出路径与输入文件的"同一文件"检查可被大小写/尾空格绕过 | `renderer.py:502` | `normcase(realpath())` 比较 + `samefile` |
| P1-3 | 固定 `.tmp` 名 + 秒级默认输出名，并发渲染互相踩踏 | `renderer.py:624`、`app.py:234` | `mkstemp` 唯一临时文件；默认名加毫秒/UUID |
| P1-4 | 绑定的 slot 名找不到时静默 0 替换 | `renderer.py:229-238` | 返回 missing 并 `warnings.append` |
| P1-5 | 换图/隐藏图不删旧 rel，每页保留死媒体 | `renderer.py:359,601` | `drop_rel` 旧 rId；跨页共享 part 缓存（实测 100 页多 28MB） |
| P1-6 | `/api/upload` 先全量读进内存再判断 1GB | `app.py:155` | 按 `Content-Length` 预拒 + 分块流式写盘 |
| P1-7 | 上传目录永不清理 | `app.py:38,153` | 启动清理 / 按 mtime TTL / 会话级临时目录 |
| P1-8 | 非法 `pagination` 直接 500 HTML | `app.py:314`、`renderer.py:415` | Pydantic 校验 + `RenderError` 包装 |
| P1-9 | 预览超时只杀 worker，WPS 变孤儿进程 | `preview.py:82`、`preview_worker.py:54` | `taskkill /T` 或 Job Object；`encoding="utf-8"` |
| P1-10 | Excel 单元格里的本地路径被直接接受 | `renderer.py:476` | 显式开关 + 限定根目录/扩展名/大小 |

---

## P2 建议（摘要）

- P2-1 表格少绑列不报错、追加行残留模板样例文字；`_set_cell_text` 不清第二段。
- P2-2 未指定 `slot` 时把同一值写进该 shape 所有槽位。
- P2-3 图表数值转换失败静默变 `None`。
- P2-4 verifier 盲区：不递归组合内部、字体签名不含 `a:latin`、残留不认 `%x%`、负索引/负页数、无期望页数校验。
- P2-5 预览的 `page/total` 变量恒为 1。
- P2-6 预览缓存键不含模板 mtime。
- P2-7 `_PREVIEW_LOCK` 覆盖整段渲染+COM，吞吐低。
- P2-8 `/api/preview-page` 丢弃 `_guard_file` 的归一化路径。
- P2-9 `excel_reader` 依赖 `ws.max_column`（可能 None / 被尾部格式撑大）。
- P2-10 `excel_images` 解析脆弱（编码、zip 顺序、Target 变体、`__skipped__` 脏接口）。
- P2-11 `titles` 无重试、JSON 解析脆弱、`base_url` 可被请求方指定（key 外泄面）。
- P2-12 死代码、`SheetData.warnings` 恒空、`cross_run` 未接。
- P2-13 `_unique_partname` O(P²)、`replaced` 量纲混用、字符串匹配去重、`save_profile` 非原子。
- P2-14 文档与实现不一致（URL 图片、指纹描述、预览一致性）。
- P2-15 测试缺口（模板空洞、并发、失败清理、titles/preview/excel_images 全套）。

---

## 最值得先修的三条

1. **P0-1**：唯一"必现型输出损坏"，修法小、收益大。
2. **P1-1 + P1-5 + P1-6 + P1-7**：临时目录泄漏、每页死媒体（28MB/100 页）、上传内存峰值、上传目录垃圾。
3. **P1-4**：槽位名不匹配时静默不替换——"成功"的报告掩盖了没替换的事实。

## 已核实无问题

- rId 重映射覆盖 `r:id/r:embed/r:link`，`PartFactory` 还原 `ChartPart`，跳过 `NOTES_SLIDE` 正确。
- 常规连续编号模板下 `_drop_slide` 生效；原子写路径正确。
- key 未落盘/未打日志；上传文件名 UUID + 后缀白名单，无路径穿越。
