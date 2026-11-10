# Excel2PPT Studio · DSH 插件

把 Excel2PPT Studio 接进 DeepSeek Harness：

| 能力 | 形态 |
|---|---|
| `excel2ppt_build` | Host 工具：按 Profile 把 Excel 生成多页 PPT |
| `excel2ppt_inspect` | Host 工具：查看模板元素与槽位 |
| `excel2ppt_verify` | Host 工具：零漂移校验 |
| 输入框下方 `Excel2PPT` 按钮 | Client 入口：打开本地四步向导 |

## 前置

1. 本仓库已安装到 Python：
   ```powershell
   python -m pip install -e ".[dev,web]"
   ```
2. `cordis.patch.yml` 里的 `pythonPath` 指向该 Python（默认 `C:\Program Files\Python313\python.exe`）。
3. 需要界面时，先启动 Studio：
   ```powershell
   e2p web
   ```

## 安装

在 Harness 里用插件管理器安装本目录（绝对路径）：

```text
plugin_manager install_bundle target=<仓库>\dsh-plugin
```

安装后可在插件列表看到「Excel2PPT Studio」；Host 工具立即对 Agent 可见，Client 按钮出现在输入框下方。

## 配置

```yaml
- id: excel2ppt
  name: '@local/excel2ppt-studio'
  config:
    pythonPath: 'C:\Program Files\Python313\python.exe'
    studioUrl: 'http://127.0.0.1:8765'
    timeoutMs: 180000
```

| 字段 | 说明 |
|---|---|
| `pythonPath` | 运行 `python -m excel2ppt_studio` 的解释器 |
| `studioUrl` | Client 按钮打开的地址 |
| `timeoutMs` | 单次生成/校验超时 |

## 说明

- 插件不依赖 npm 包，只用 Node 标准库，安装时不需要额外解析依赖树。
- Client 端不在 Harness 页面里嵌入 iframe，而是新开标签页，避免主题与布局互相污染。
- 生成/校验失败时，Python 的中文错误信息会原样作为工具错误返回。
