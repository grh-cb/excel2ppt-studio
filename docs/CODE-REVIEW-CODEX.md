# Codex 审查报告（第二轮，交叉验证）

> 工具：`codex exec --json`（codex-cli 0.163.0，DeepSeek `deepseek-flash`，`model_reasoning_effort=low`）
> 输入：65.3 万 token，输出 1.9 万 token，用时约 8 分钟
> 特点：它会主动用 python-pptx 内存态做实验验证，因此结论带实证

## P0-1 图片替换写出顺序错误的 `a:srcRect`（已实证）

- 位置：`renderer.py:342`（`etree.SubElement(blip_fill, qn("a:srcRect"))`）
- 问题：`SubElement` 追加到末尾，实际结构变成 `blip → stretch → srcRect`；而 `CT_BlipFillProperties` 要求 `blip → srcRect → stretch/tile → extLst`。
- **实测**（本仓库真实输出 `out/real/full.pptx`）：
  ```
  pic 5: blipFill children = ['blip', 'stretch', 'srcRect']
  pic 8: blipFill children = ['blip', 'stretch', 'srcRect']
  ```
- 后果：只要模板图片原本没有 `srcRect`，`contain / cover / stretch` 都会写出架构无效 XML；PowerPoint 可能提示修复并丢失裁剪/填充设置。WPS 宽容，所以之前用 WPS 预览没暴露。
- 修复：用顺序感知 API
  ```python
  src_rect = blip_fill.get_or_add_srcRect()
  ```
  并补一个"子元素顺序"回归测试。

## P0-2 模板 slide 编号空洞 → partname 重名

- 位置：`renderer.py:125`
- 与上一轮 P0-1 一致（详见 `docs/CODE-REVIEW.md`）。
- 修复：
  ```python
  dest.part.partname = prs.part.package.next_partname("/ppt/slides/slide%d.xml")
  ```
  保存前断言唯一。

## P1 列表

| 编号 | 问题 | 位置 | 备注 |
|---|---|---|---|
| C-1 | SmartArt 的 `dgm:relIds`（`r:dm/r:lo/r:qs/r:cs`）未重映射 | `renderer.py:49,114` | Codex 已实证 `_remap_rids` 不改这些属性；应改为"属性在关系命名空间且值命中 rid_map 就替换" |
| C-2 | 并发渲染共用固定 `<out>.tmp` | `renderer.py:608` | 与上一轮 P1-3 一致；改 `tempfile.mkstemp` |
| C-3 | 输入输出同文件检查可被路径别名绕过 | `renderer.py:486` | 与上一轮 P1-2 一致；`samefile` + `normcase(realpath)` |
| C-4 | `value_map` 中 `0`/`False`/空串被真值判断误判为未命中 | `renderer.py:553` | **上一轮未发现**；用哨兵 `object()` 区分"键不存在" |
| C-5 | 图表数值转换失败静默写 `None` | `renderer.py:325` | 与上一轮 P2-3 一致；应带行号/列名警告 |
| C-6 | 指定槽位找不到时静默返回 0 | `renderer.py:221` | 与上一轮 P1-4 一致 |
| C-7 | 演讲者备注被无条件丢弃且无警告 | `renderer.py:133` | 与上一轮判断不同：上一轮认为跳过是正确的（避免反向引用拖入模板页），本轮认为属于数据丢失；合理处理是"跳过 + 一次性警告" |

## 两轮报告的一致性

| 结论 | 第一轮（独立 agent） | 第二轮（Codex） |
|---|---|---|
| partname 重名 | P0-1 | P0-2 ✅一致 |
| 并发 `.tmp` 踩踏 | P1-3 | C-2 ✅一致 |
| 输出路径绕过同文件检查 | P1-2 | C-3 ✅一致 |
| 槽位名找不到静默 0 替换 | P1-4 | C-6 ✅一致 |
| 图表数值静默 `None` | P2-3 | C-5 ✅一致 |
| 备注页跳过 | 判为正确 | 判为数据丢失（分歧，建议加警告） |
| `srcRect` 子元素顺序 | 未发现 | **P0-1（已实证）** |
| SmartArt relIds 未重映射 | 未发现 | **C-1** |
| `value_map` 假值误判 | 未发现 | **C-4** |
| 临时目录泄漏 / 死媒体 / 上传内存 | P1-1/5/6/7 | 未覆盖（未审这些文件） |

## 结论

两轮报告互补：Codex 强在"OOXML 规范细节 + 内存态实证"，第一轮强在"资源与 Web 层"。合并后 P0 应为 **2 条**（partname 重名 + srcRect 顺序），其中 srcRect 是**当前真实输出每页都会触发**的问题，优先级最高。
