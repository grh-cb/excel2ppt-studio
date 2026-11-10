# 架构说明

## 一句话

**读模板 → 认元素 → 连映射 → 克隆页 → 换文本 → 校验 → 出片。**

## 分层

```
┌─────────────────────────────────────────────────────┐
│ 入口层  cli.py（e2p）        web/app.py（M2）        │
├─────────────────────────────────────────────────────┤
│ 应用层  mapping.py（建议/Profile）  renderer.py      │
├─────────────────────────────────────────────────────┤
│ 领域层  template_scanner.py  excel_reader.py        │
│         slot_detect.py       excel_images.py        │
├─────────────────────────────────────────────────────┤
│ 基础层  model.py（dataclass）  verifier.py（M3）     │
└─────────────────────────────────────────────────────┘
```

## 模块职责

| 模块 | 职责 | 不做 |
|---|---|---|
| `model.py` | 纯数据模型与 JSON 序列化 | 不碰 OOXML |
| `slot_detect.py` | 从文本/形状元数据识别槽位 | 不解析 PPT |
| `template_scanner.py` | 打开模板、递归遍历、产出 `Field[]` | 不改文件 |
| `excel_reader.py` | 表头校验、类型推断、行数据 | 不读 PPT |
| `excel_images.py` | 从 xlsx 包内提取 DISPIMG 图片 | 不决定排版 |
| `mapping.py` | 自动建议 + Profile 读写 + 校验 | 不渲染 |
| `renderer.py` | 克隆、替换、分页、原子写 | 不猜映射 |
| `verifier.py` | 格式指纹、页数、图片校验 | 不修文件 |

## 关键不变量（INV）

| 编号 | 不变量 | 违反后果 |
|---|---|---|
| INV-1 | 输出页必须是模板页 XML 的深拷贝，禁止新建 shape | 版式漂移 |
| INV-2 | 替换只写 `<a:t>`（及显式开启的 `a:rPr/@sz`） | 字体/颜色丢失 |
| INV-3 | 克隆时必须重映射所有 `r:id` / `r:embed` / `r:link` | 图片/图表断链、文件损坏 |
| INV-4 | 克隆页复用原 `slideLayout`，不复制版式树 | 版式不同步、文件膨胀 |
| INV-5 | 一律原子写（`.tmp` → `os.replace`） | 中途失败顶掉成品 |
| INV-6 | 任何降级/跳过都必须有 WARNING 或错误码 | 静默数据丢失 |

## 数据流

```
template.pptx ──scanner──► Field[] ─┐
                                    ├──mapping──► Profile ──renderer──► out.pptx
excel.xlsx ───reader────► Column[] ─┘                                    │
                                                                         ▼
                                                              verifier（指纹/页数）
```

## 术语表

| 术语 | 定义 |
|---|---|
| Field | 模板中的一个可映射元素（文本/表格/图片/图表/组合） |
| Slot | Field 内可被数据填充的位置（`{{x}}`、命名形状、替代文字等） |
| Binding | 一个 Field 与一个数据来源的绑定关系 |
| Profile | 一份可保存/复用的完整映射方案（JSON） |
| 克隆 | 对模板页部件树做深拷贝并重映射关系 |
| 格式指纹 | 归一化后用于比对 XML 结构的哈希/树 |
| 主槽位 | 一个 Field 上优先匹配的槽位（多槽位时取第一个显式槽位） |

## 为什么不用「重建页面」

见 [ADR-0001](adr/0001-clone-preserving-render.md)。核心：重建必然重写格式属性，克隆则天然继承。

## 扩展点

| 需求 | 扩展位置 |
|---|---|
| 新元素类型（SmartArt） | `template_scanner` 增加分支 + `renderer` 适配器 |
| 新数据源（CSV/数据库） | 新增 reader，实现 `(columns, rows)` 契约 |
| 新绑定类型 | `model.Binding.kind` + `renderer` 分派 |
| 新预览引擎（LibreOffice） | `preview.py` 增加后端，保持接口 |
