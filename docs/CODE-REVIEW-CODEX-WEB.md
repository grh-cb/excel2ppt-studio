# Codex 审查报告：Web 层与预览（第三轮）

> 工具：`codex exec --json`，输入 32.3 万 token，输出 1.8 万 token
> 范围：`web/app.py`、`core/preview.py`、`core/preview_worker.py`，并追踪 `renderer.py` 调用链
> 结论：高优先级 7 条、中优先级 3 条

## 高优先级

| # | 问题 | 位置 | 修复要点 |
|---|---|---|---|
| 1 | 上传在大小检查前整体读入内存；异步端点里同步写盘 | `app.py:153-156` | 先看 `Content-Length` 返 413 → 分块读 → `.part` + 原子替换 → `run_in_threadpool` |
| 2 | 上传只校验扩展名，不验证 OOXML 内容（ZIP bomb 风险） | `app.py:147-150` | 落盘后检查 ZIP 头与必需成员（`[Content_Types].xml` 等），限制成员数与解压总量 |
| 3 | 路径接口信任任意本机绝对路径 | `app.py:124,230,271,298` | 上传返回不可猜 upload ID，服务端映射；`resolve(strict=True)` + `is_relative_to` + 后缀校验；拒绝 UNC/设备路径/重解析点；输出路径由服务端生成 |
| 4 | COM 超时只杀 worker，Office/WPS 子进程树未清理 | `preview.py:76`、`preview_worker.py:52` | Windows Job Object（`KILL_ON_JOB_CLOSE`）+ `TerminateJobObject`；`taskkill /T` 可靠性弱于 Job |
| 5 | 预览缓存键不含模板版本；降级结果也被缓存 | `app.py:102-109,374` | 键加入模板/Excel 的 `(mtime_ns,size,file_id)` 或 SHA-256 + 外部图片依赖 + 后端；降级结果短 TTL 或不缓存 |
| 6 | COM 并发控制不完整：锁覆盖整段渲染，`/api/preview` 又绕过锁 | `app.py:302,275` | 专用 COM 信号量（全局最多一个导出）；缓存锁只保护字典；同 key 用 in-flight event 去重 |
| 7 | 默认输出名秒级精度 + 固定 `.tmp` 并发冲突 | `app.py:231-233`、`renderer.py:608` | UUID/随机后缀输出名；独占临时文件；同目标并发返回 409 或排队 |

## 中优先级

| # | 问题 | 位置 | 修复要点 |
|---|---|---|---|
| 8 | 上传目录无 TTL/会话归属，永久累积 | `app.py:36,151-157` | 每次启动独立目录 + lifespan 清理 + TTL / 会话关闭接口 |
| 9 | 渲染内部 `e2p_img_*` 在异常路径泄漏 | `renderer.py:442,616` | 创建后立即进入覆盖全程的 `try/finally`；启动时扫描超龄目录 |
| 10 | 错误响应契约不统一 | `app.py:120,126,262,358` | 统一 `{ok:false,error:{code,message,details}}`；413/400/422/503/504 语义化；降级用 `ok:true,degraded:true` |

## 与第一轮的交叉

| 结论 | 第一轮 | Codex |
|---|---|---|
| 上传先生读后判大小 | P1-6 | #1 ✅ |
| 上传目录不清理 | P1-7 | #8 ✅ |
| 预览超时孤儿进程 | P1-9 | #4 ✅（方案更完整） |
| 缓存键不含模板 mtime | P2-6 | #5 ✅（并补充"降级结果被缓存"新问题） |
| `_PREVIEW_LOCK` 覆盖过宽 | P2-7 | #6 ✅（并补充 `/api/preview` 绕过锁） |
| 默认输出名/tmp 并发 | P1-3 | #7 ✅ |
| 图片临时目录泄漏 | P1-1 | #9 ✅ |
| 错误契约 | P2-13 | #10 ✅ |
| **上传不验 OOXML 内容** | 未提 | **#2 新增** |
| **路径校验不完整（UNC/设备/重解析点）** | 部分 | **#3 更系统** |

## 判断

Web 层这 10 条里，真正影响"给同事用"的是 **#1、#3、#4、#8**：
- #1 决定 300MB 台账会不会把内存打满；
- #3 决定"同事发来的台账"能不能读到本机任意文件；
- #4 决定预览失败后 WPS 残留进程会不会越积越多；
- #8 决定临时盘会不会被上传文件堆满。
