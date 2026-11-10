# Codex 审查报告：校验与扫描层（第五轮）

> 范围：`verifier.py`、`template_scanner.py`
> 结论：P1 共 10 条、P2 共 7 条；关键结论均用内存 `python-pptx` 复现

## verifier.py

| # | 问题 | 位置 | 实证/修复 |
|---|---|---|---|
| V1 | 组合形状内部完全不参与比对 | `verifier.py:161,169` | 子元素增删/换 ID/换类型/调位置都不被发现；需递归带路径指纹 |
| V2 | 坐标指纹只比 `off/ext`，旋转翻转被测不出 | `verifier.py:88,218` | **实测：旋转 45° 后 rect 不变、校验通过**；需纳入 `rot/flipH/flipV/chOff/chExt` |
| V3 | 字体签名没有字体名，且丢 run 顺序 | `verifier.py:126` | **实测：Arial→Times New Roman 仍通过**；需加 `a:latin/a:ea/a:cs` 并按序保留重数 |
| V4 | 残留占位符检测不完整，且只算 warning | `verifier.py:24,142,239` | 不认 `%x%`；**跨 run 拆开的 `{`+`{slot}}` 实测 residual 为空**；`ok` 只由 error 决定 |
| V5 | 不校验实际页数；`max_pages` 部分校验仍返回 `ok=True` | `verifier.py:288,296` | 需 `expected_pages`；部分校验应标 `complete=False` |
| V6 | 页码/总页数文本无语义校验 | `verifier.py:9` | 第 N 页显示 1、总页数不等于实际页数都能通过 |
| V7 | shape ID 缺失/重复被静默跳过或覆盖 | `verifier.py:151,172` | 应报 `missing_shape_id` / `duplicate_shape_id` |
| V8 | 负 `slide_index` 静默取最后一页；异常未包装 | `verifier.py:278,287` | 校验 `0 <= slide_index < len` |

## template_scanner.py

| # | 问题 | 位置 | 实证/修复 |
|---|---|---|---|
| S1 | 组合坐标换算忽略旋转/翻转 | `scanner.py:100,229-248` | 只做轴向缩放位移，旋转组合坐标错；应矩阵组合后求 AABB |
| S2 | 嵌套组合每层 `int()` 截断，误差累积 | `scanner.py:233,242` | 递归用 float，末端统一取整 |
| S3 | `chOff/chExt` 缺失/非法时回退到错误坐标系 | `scanner.py:49,111,237` | 应标记不支持并告警，而非 1:1 假设 |
| S4 | 图片占位符被识别成文本 + `body` | `scanner.py:290,310` | **实测：默认布局 Picture Placeholder → kind=text, role=body**；应先判 `PP_PLACEHOLDER.PICTURE` |
| S5 | 空文本框/空自选图形都变成 `body` | `scanner.py:84,310` | `if paragraphs:` 对空 text frame 也为真；应区分"有内容/有意义空位/装饰" |
| S6 | 标题判定只看字号，遗漏 DATE/FOOTER/SLIDE_NUMBER | `scanner.py:179,197` | 大号正文误判标题，主题字号标题反而成 body |
| S7 | 跨 run 占位符丢失显式语义 | `scanner.py:314` | `find_cross_run_slots` 闲置未调用 |
| S8 | 业务解析异常未包装成 `TemplateError` | `scanner.py:356,369` | 表格/图表/字体元数据读取异常直接泄漏 |
| S9 | 表头判定只认"首行全非空" | `scanner.py:130` | 数据行可能被误判表头；缺一个表头名又判无表头 |
