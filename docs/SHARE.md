# 分享到 DSH 社区

DSH 目前**没有官方插件市场**：Web 的 Plugins 页只有「已安装」列表，Discover 标签只读用户自己添加的**索引源**。所以分享有三条路径，按投入从低到高排列。

---

## 方式 A：发布到 npm（推荐，用户一条命令安装）

```powershell
cd dsh-plugin
npm publish --access public
```

用户安装：

```text
plugin_manager install_bundle target=dsh-excel2ppt-studio
```

### 发布前改 `dsh-plugin/package.json`

| 字段 | 改为 |
|---|---|
| `name` | 自己的 scope，如 `@your-name/dsh-excel2ppt-studio` |
| `version` | 语义化版本，如 `0.2.0` |
| `private` | `false` |
| `publishConfig.access` | `public` |
| `repository.url` | GitHub 仓库地址 |
| `engines.dsh` | 兼容的 DSH 版本，如 `>=0.2.1-rc.5` |
| `files` | 确认包含 `index.js`、`client.js`、`cordis.patch.yml`、`locale/*.json`、`icon.svg` |

### 注意：插件只是壳

`dsh-plugin` 不包含 Python 后端。用户还要：

```powershell
pip install excel2ppt-studio
```

并在 `cordis.patch.yml` 里确认 `pythonPath` 指向装了该包的解释器。README 必须写清楚，否则装完插件调工具会报「No module named excel2ppt_studio」。

---

## 方式 B：Git 源（不发布 npm）

把仓库推到 GitHub，然后让用户按 commit 安装：

```text
plugin_manager install_bundle target=git+https://github.com/<user>/excel2ppt-studio.git#<40位commit>
```

pnpm 支持 `git+https` 规格；包在仓库子目录时用 `path:` 参数（市场索引里对应 `source.path`）。

适合内部团队，不需要 npm 账号。

---

## 方式 C：自建插件市场索引（社区目录）

做一个 JSON 索引，托管在 GitHub Pages / raw.githubusercontent，用户把它加进 Discover，就能浏览、搜索、一键安装。

### 索引格式（format 1）

```json
{
  "format": 1,
  "title": { "zh": "Excel2PPT 插件源", "en": "Excel2PPT Market" },
  "description": { "zh": "办公文档自动化插件集合", "en": "Office document automation plugins" },
  "metadata": {
    "categories": { "office": { "title": { "zh": "办公文档", "en": "Office" } } },
    "labels": {},
    "sortKeys": {}
  },
  "plugins": [
    {
      "name": "@your-name/dsh-excel2ppt-studio",
      "title": { "zh": "Excel2PPT Studio", "en": "Excel2PPT Studio" },
      "description": {
        "zh": "用单页 PPT 模板把 Excel 台账逐行生成多页幻灯片，克隆母版页零漂移。",
        "en": "Generate multi-page slides from Excel rows with a zero-drift template clone."
      },
      "author": "一片风干叶",
      "homepage": "https://github.com/<user>/excel2ppt-studio",
      "repository": {
        "url": "https://github.com/<user>/excel2ppt-studio",
        "directory": "dsh-plugin"
      },
      "metadata": {
        "category": "office",
        "keywords": ["excel", "ppt", "powerpoint", "report"],
        "labels": [],
        "sortKeys": {}
      },
      "versions": [
        {
          "version": "0.2.0",
          "publishTimestamp": 1762560000,
          "source": { "type": "npm" },
          "engines": { "dsh": ">=0.2.1-rc.5" }
        }
      ]
    }
  ]
}
```

- `versions` 按 SemVer 降序；`publishTimestamp` 是 Unix 秒。
- `source` 可以是 `{"type":"npm"}` 或 `{"type":"git","url":"https://...","commit":"<40位sha>","path":"dsh-plugin"}`。
- `metadata.category` 必须在 `metadata.categories` 里声明，否则校验失败。
- 图标可用 `data:` URL 或绝对 URL（相对引用会按索引 URL 解析）。

### 用户怎么加

1. Web → **Plugins** → **More**（Experimental plugins）→ 打开 **Custom plugin markets**。
2. 进入 **Discover** → 添加索引 URL（`https://.../index.json`）。
3. 搜索并安装。

> 该市场能力是实验性 bundle（`@deepseek-ai/dsh-experimental-plugin-market-bundle`），默认关闭，需要用户手动开。

---

## 发布检查清单

- [ ] `dsh-plugin/package.json`：`name` 唯一、`version`、`private:false`、`files` 完整
- [ ] `cordis.patch.yml` 的 `pythonPath` 写成**相对通用**的默认值，README 说明怎么改
- [ ] `locale/zh.json`、`locale/en.json` 有 title/description
- [ ] `icon.svg` 合法且 < 256 KiB
- [ ] `client.js` 只依赖 `react`（browser 模块表），不引其他 Harness 包
- [ ] README 写清 Python 依赖与 Studio 启动方式
- [ ] 在干净 profile 里装一次验证（`plugin_manager install_bundle` → `application: applied`）
- [ ] 版本号与主仓库 `pyproject.toml` 保持同步

---

## 建议的发布节奏

| 阶段 | 动作 |
|---|---|
| 内部试用 | 方式 B（Git 源），迭代快 |
| 小范围分享 | 方式 A（npm），README + 安装命令 |
| 社区目录 | 方式 C，建一个索引仓库收录自己的 2–3 个插件 |
