# Excel2PPT Studio · DSH 插件

把 Excel2PPT Studio 接进 DeepSeek Harness：Excel 台账按单页 PPT 模板逐行生成多页幻灯片，**克隆母版页、只换内容**，排版/颜色/字体/位置零漂移。

## 三步安装

```powershell
# 1. 装后端（PyPI）
pip install excel2ppt-studio

# 2. 装插件（在 Harness 里）
plugin_manager install_bundle target=dsh-excel2ppt-studio

# 3. 需要图形界面时
e2p web
```

> 不想装 Python？直接下载[便携版](https://github.com/grh-cb/excel2ppt-studio/releases/latest)，解压双击即可，无需第 1 步。

## 能力

| 能力 | 形态 |
|---|---|
| `excel2ppt_build` | Host 工具：按 Profile 把 Excel 生成多页 PPT |
| `excel2ppt_inspect` | Host 工具：查看模板元素与槽位 |
| `excel2ppt_verify` | Host 工具：零漂移校验 |
| 输入框下方 `Excel2PPT` 按钮 | Client 入口：打开本地四步向导 |

## 配置

```yaml
- id: excel2ppt
  name: 'dsh-excel2ppt-studio'
  config:
    pythonPath: 'python'
    studioUrl: 'http://127.0.0.1:8765'
    timeoutMs: 180000
```

| 字段 | 说明 |
|---|---|
| `pythonPath` | 运行 `python -m excel2ppt_studio` 的解释器（默认 `python`） |
| `studioUrl` | Client 按钮打开的地址 |
| `timeoutMs` | 单次生成/校验超时 |

## 其他链接

- 包：https://pypi.org/project/excel2ppt-studio/
- 源码与 Release：https://github.com/grh-cb/excel2ppt-studio
- 市场索引：https://grh-cb.github.io/excel2ppt-studio/market-index.json

## 说明

- 插件不依赖 npm 包，只用 Node 标准库，安装时不需要额外解析依赖树。
- 生成/校验失败时，Python 的中文错误信息会原样作为工具错误返回。
