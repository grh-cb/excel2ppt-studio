# 便携版（Windows，无需安装 Python）

给同事一个压缩包，解压双击就能用。目标机器只需要 **Windows 10/11 x64 + 浏览器**。

## 包里有什么

```
Excel2PPT-Studio\
  Excel2PPT-Studio.exe      ← 双击这个
  _internal\                ← 运行库与前端资源（不要删）
```

## 使用步骤

1. 解压到本地磁盘（建议路径短、不要放 OneDrive/网络盘，例如 `D:\Excel2PPT`）。
2. 双击 `Excel2PPT-Studio.exe`。
3. 黑色控制台窗口会显示地址（默认 `http://127.0.0.1:8765`，被占用会自动换端口），浏览器自动打开。
4. 第 1 步有两种选文件方式，**都保留**：
   - 点 **「上传」** 从浏览器选文件（同事机器推荐，不需要知道路径）；
   - 或直接在输入框**粘贴本机绝对路径**（同一台机器上更快，310MB 不用复制）。
5. 后面就是四步向导：认元素 → 连映射 → 预览 → 生成。
6. **关闭黑色窗口即停止服务**，PPT 文件已经写到你在第 4 步看到的路径。

## 系统要求

| 项 | 说明 |
|---|---|
| 系统 | Windows 10 / 11 x64 |
| 浏览器 | Edge / Chrome（任意现代浏览器） |
| Office | **生成不需要**；预览需要 WPS 或 PowerPoint，没有则自动降级为几何预览 |
| 网络 | 完全离线可用（AI 短标题除外，它要调 DeepSeek） |

## 首次运行可能遇到

| 现象 | 处理 |
|---|---|
| SmartScreen「Windows 已保护你的电脑」 | 点「更多信息」→「仍要运行」 |
| 杀毒软件拦截 | PyInstaller 打包的程序常被误报，添加信任或白名单 |
| 端口被占用 | 启动器会自动换端口，看控制台里的地址 |
| 上传大文件慢 | 310MB 台账约 10–30 秒，属正常 |

## 数据安全

- 服务只监听 `127.0.0.1`，不对外暴露。
- 文件只在本机读写，不上传云端。
- AI 短标题会把问题描述发给 DeepSeek；不用该功能时完全不联网。

## 与安装版的差异

| 项 | 便携版 | 安装版（pip） |
|---|---|---|
| 需要 Python | 否 | 是 |
| 升级 | 换整个文件夹 | `pip install -U` |
| 体积 | 约 90MB（压缩后 40MB 左右） | 几 MB |
| 扩展 | 不能加 Python 插件 | 可以 |

## 自己构建

```powershell
powershell -ExecutionPolicy Bypass -File packaging\build_portable.ps1
Compress-Archive -Path dist\Excel2PPT-Studio -DestinationPath dist\Excel2PPT-Studio-win64.zip -Force
```

产物：`dist\Excel2PPT-Studio\Excel2PPT-Studio.exe`。
