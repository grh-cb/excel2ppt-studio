# 发布指南

## 1. 版本与变更

- 遵循 [语义化版本](https://semver.org/lang/zh-CN/)：`0.x` 阶段小版本可有破坏性变更。
- 改版本号要同步三处：
  1. `pyproject.toml` → `project.version`
  2. `src/excel2ppt_studio/__init__.py` → `__version__`
  3. `CHANGELOG.md` → 新增对应条目（把 `Unreleased` 改成发布日期）

## 2. 发布前检查清单

```powershell
cd E:\dsh\work\excel2ppt-studio
$py = 'C:\Program Files\Python313\python.exe'

# 静态检查 + 全量测试
& $py -m ruff check src tests
& $py -m pytest -q

# 示例可生成
& $py examples\make_sample.py

# CLI 可用
& $py -m excel2ppt_studio inspect examples\sample\template.pptx
& $py -m excel2ppt_studio build --profile <profile.json> --json

# Web 可启动
& $py -m excel2ppt_studio web --port 8765 --no-browser
```

手工走查（必须）：

1. 打开 `http://127.0.0.1:8765`，选 `examples/sample` 下两个文件。
2. 画布元素与清单联动；数据预览正确。
3. 自动映射后能手工改；导出/导入 Profile。
4. 生成 → 校验零漂移 → 预览第 1 页。

## 3. GitHub 发布

本机若未安装 Git：先装 [Git for Windows](https://git-scm.com/download/win)。

```powershell
cd E:\dsh\work\excel2ppt-studio

git init -b main
git add .
git commit -m "feat: excel2ppt-studio v0.4.0（克隆模板零漂移生成）"

# 在 GitHub 建空仓库后
git remote add origin https://github.com/<user>/excel2ppt-studio.git
git push -u origin main

# 打标签
git tag v0.4.0
git push origin v0.4.0
```

CI（`.github/workflows/ci.yml`）会在 push / PR 时跑 Windows + Ubuntu × Python 3.10/3.13 的 `ruff` 与 `pytest`。

> 首次 push 前把 `pyproject.toml` 和 `README.md` 里的 `your-name` 占位换成真实用户名。

## 4. 发布到 PyPI（可选）

```powershell
& $py -m pip install build twine
& $py -m build
& $py -m twine upload dist/*
```

包名 `excel2ppt-studio`，命令行入口 `e2p`。发布前确认 `LICENSE` 与 `README.md` 已随包打包。

## 5. Windows 可执行程序

```powershell
powershell -ExecutionPolicy Bypass -File packaging\build_exe.ps1
```

产物：`dist\e2p\e2p.exe`（onedir，目录整体分发）。验证：

```powershell
& .\dist\e2p\e2p.exe --help
& .\dist\e2p\e2p.exe web --port 8766 --no-browser
```

注意：exe 内已打包 Web 静态资源；预览功能仍依赖外部 WPS/PowerPoint。

## 6. DSH 插件

目录安装（推荐，跟随仓库版本）：

```text
plugin_manager install_bundle target=<仓库>\dsh-plugin
```

发布到 npm（可选）：

1. 把 `dsh-plugin/package.json` 的 `name` 改为自己的 scope（如 `@your-name/excel2ppt-studio`），`private` 改 `false`。
2. `npm publish --access public`。
3. 用户在 Harness 里 `plugin_manager install_bundle target=@your-name/excel2ppt-studio`。

插件依赖本机 Python：用户需先 `pip install -e .`，并在 `cordis.patch.yml` 里改 `pythonPath`。

## 7. 发布后验证

| 项 | 做法 |
|---|---|
| 干净安装 | 新建 venv：`pip install excel2ppt-studio`，跑 `e2p --help` 与一次生成 |
| CI | 确认 GitHub Actions 两个平台绿 |
| exe | 在一台未装 Python 的机器上运行 `e2p.exe web` |
| 插件 | 在 Harness 里确认插件行与三个工具可见，Client 按钮可打开页面 |

## 8. 回滚

- PyPI / npm：撤回对应版本（`twine yank` / `npm unpublish`）。
- GitHub：`git revert` 或回退 tag。
- DSH 插件：`plugin_manager remove_bundle`。
