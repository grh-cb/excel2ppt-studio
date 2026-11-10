# 更新日志

本项目遵循 [语义化版本](https://semver.org/lang/zh-CN/)，格式参考 [Keep a Changelog](https://keepachangelog.com/zh-CN/1.1.0/)。

## [0.7.1] - 2026-11-10

### 新增
- 多模板编排简化为「正文映射 + 封面/结尾文字配置」，不需要在画布上映射封面
- 模板体检面板：Step1 显示分数与问题清单（`e2p doctor` 同源）
- 自动识别正文页，并给出封面/正文/结尾的页型判断
- 映射下拉新增「固定文本…」，用于封面标题、结尾致辞

### 修复
- 生成时删除全部母版页：封面/结尾不再跑到正文前面
- 封面/结尾只写入「有文字的文本框」，不再写进装饰空框
- Step3 预览支持整册（封面 → 正文 → 结尾）并缓存
- 校验不再把业务结论 `ok:false` 当成请求失败
- `constant` 绑定不再被 layout 过滤（封面/结尾因此丢失）
- 尾部空表头列自动裁剪；重复列名自动重命名为「列名(2)」
- 整段替换写入字号最大的 run，正文不再出现 14pt 首行
- 图片 `a:srcRect` 顺序修正；克隆页 partname 唯一化

## [0.7.0] - 2026-11-10

### 新增
- **多模板编排**：`layout.cover / section / end`（同一 pptx 的多页），配合 `pagination.group_by`
  生成「封面 → 章节 → 正文 → 结尾」；缺省 `layout` 时行为与单页模式完全一致
- **模板体检**：`e2p doctor`、`POST /api/doctor`、Step1 体检卡片（分数 + 问题清单）
  - 结构：元素分布、字号/字体种类、自动调整
  - 可映射性：段落槽位提示、跨 run 占位符、图片位与 Excel 图片列交叉检查
- 页码变量在多模板下按真实页序号计算（含封面与章节页）

### 修复
- `read_dispimg_formulas` 自动查找含 DISPIMG 的工作表（台账常把「使用说明」放第一页）

## [0.6.1] - 2026-11-09

### 修复（代码审查第一批）
- 图片 `a:srcRect` 改用顺序感知 API，修正 `blip → stretch → srcRect` 的 OOXML 顺序错误
- `clone_slide` 为克隆页分配唯一 partname，模板编号有空洞时不再产生重复 zip 条目
- 图片临时目录的清理覆盖整个渲染过程，失败路径不再泄漏 `e2p_img_*`
- 换图与隐藏空图时删除旧关系，消除每页残留的模板图（100 页实测减少 26.9MB）
- `build_profile` 保留 `header_row`，表头不在第 1 行时不再错位
- 图片占位符（`PP_PLACEHOLDER.PICTURE`）识别为 `image`，不再落成文本 `body`

## [0.6.0] - 2026-11-09

### 新增
- `e2p titles`：用 DeepSeek / OpenAI 兼容接口批量生成 ≤16 字语义短标题，输出 `{序号: 标题}`
- binding 支持 `value_map`：按 `key_column` 查表覆盖列值（标题列无数据时的解法）
- binding 支持 `var`：`page` / `total` / `date` 页码类变量
- Web 映射下拉新增「变量」分组（页码、总页数、生成日期）
- `docs/SHARE.md`：DSH 社区分发的三条路径与索引格式
- 便携版：`packaging/build_portable.ps1` 生成双击即用的 exe（自动选端口、无需 Python）
- `docs/PORTABLE.md`：便携版分发与使用说明（路径 / 上传两种选文件方式都保留）

## [0.5.1] - 2026-11-09

真实台账（100 行 / 296MB / 240 张内嵌图）实测后修复。

### 修复
- **克隆跳过 notesSlide**：备注页持有反向 slide 引用，之前每克隆一页都会把模板页整页复制进包，导致 `slideN.xml` 重名、模板页残留、WPS 渲染错页
- **DISPIMG 公式值**：`=DISPIMG("ID_...",1)` 现在能正确解析出图片 ID（之前只认裸 `ID_xxx`）
- **按需解压图片**：只提取本页用到的图片，300MB 台账不会整体解压
- **空图片位隐藏**：图片列为空时删除占位形状，不再保留模板原图；校验对图片缺失降级为警告

### 新增
- **段落级槽位**（`第N段`）：无占位符模板（一页已填好的样例）可按段落逐段映射，只替换该段、其余段落不动
- 图片绑定支持 `options.on_empty = hide`

## [0.5.0] - 2026-11-07

### 变更
- 「认元素」页即可就地映射：点画布元素弹出浮层，直接选 Excel 列；右侧详情同步
- 已映射元素：实线高亮 + 标签显示 `→ 列名`；未映射保持虚线
- 「连映射」页改为「预览」页：按页渲染 WPS 真实缩略图，支持翻页、切换分页、映射摘要
- 新增 `POST /api/preview-page`：单页渲染 + LRU 缓存 + 无 WPS 时降级几何预览

### 新增
- `render(rows_override=...)`：只渲染指定页的行，供预览复用
- 自动匹配 / 导入导出 Profile 移到第 2 步工具条

## [0.4.0] - 2026-11-07

里程碑 M0–M4 的合并发布。

### 新增（M4 · 图表与 DSH 插件）
- 图表数据绑定：`chart.replace_data` 保留图表类型/配色/坐标轴，同步内嵌工作簿
- `PartFactory` 克隆修复：图表 part 保留 `ChartPart` 类型，`shape.chart` 可访问
- 分页：`one_per_row` / `n_per_page`（Web 映射页可选）
- DSH 插件 `dsh-plugin/`：Host 工具 `excel2ppt_build` / `excel2ppt_inspect` / `excel2ppt_verify`，Client 输入框入口按钮
- CLI `--json` 输出（build / verify），供插件与脚本解析

### 新增（M3 · 校验与预览）
- `core.verifier`：格式指纹（位置/尺寸/字体签名/残留槽位），`e2p verify` 命令
- `core.preview`：WPS / PowerPoint 单页导出，COM 调用隔离到子进程，失败降级
- `core.text_fit`：文本溢出估算（中文 1em / ASCII 0.55em 近似），只告警不改字号
- Web：生成页增加「校验零漂移」「预览第 1 页」，展示校验明细与真实渲染图
- `/api/verify`、`/api/preview` 两个端点

### 新增（M2 · Web GUI）
- FastAPI 本地服务：`/api/scan`、`/api/excel`、`/api/suggest`、`/api/build`、`/api/upload`
- 四步向导：选文件 → 认元素（画布 + 列表 + 详情联动）→ 连映射（逐槽位选列）→ 生成
- 画布按模板真实 EMU 坐标绘制元素框，按类型着色
- 数据预览（前 20 行 + 类型标记）
- Profile 导出/导入 JSON，并自动存浏览器 localStorage
- 分页模式：每行一页 / 每 N 行一页
- `e2p web` 命令，自动打开浏览器

### 新增（M0/M1）
- 模板扫描器：文本框 / 表格 / 图片 / 图表 / 组合形状提取，含绝对坐标换算与槽位探测
- Excel 读取器：表头校验、类型推断、WPS DISPIMG 图片提取
- 克隆渲染引擎：部件树深拷贝 + rId 重映射，文本/表格/图片零漂移替换
- 映射模型与自动建议（列名 ↔ 槽位名）
- CLI：`inspect` / `suggest` / `build`

### 说明
- 首版超长文本只告警，不改字号

## [0.1.0] - 2026-11-07

### 新增
- 项目初始化：MIT 许可、pyproject、CI、架构文档与 ADR
