# 示例

运行一次即可生成示例模板与数据：

```powershell
python examples/make_sample.py
```

产物：

| 文件 | 说明 |
|---|---|
| `sample/template.pptx` | 单页模板：标题 `{{项目名称}}`、正文两槽位、3 列表格、图片位、图表 |
| `sample/sample_data.xlsx` | 3 行数据：项目名称/单位/问题描述/备注/问题图片/数量 |
| `sample/example_photo.png` | 示例图片（供图片列使用） |

然后启动界面：

```powershell
python -m excel2ppt_studio web
```

依次选择两个文件 → 查看元素 → 自动映射 → 生成，即可得到 3 页 PPT。

> 这些文件由脚本生成，不随仓库提交二进制；`examples/sample/` 已在 `.gitignore` 中。
