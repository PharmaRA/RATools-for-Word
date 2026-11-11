<div align="center">
  <img src="_image/logo.png" alt="RATools for Word" width="320" />
  <h1>RATools for Word</h1>
  <p>面向药品注册资料编写的 Word 效率插件</p>
</div>

把样式排版、题注引用、审阅清理和批量文件处理集中到 Word 的 RATools 选项卡，减少重复操作。插件使用 `.dotm`，中文和英文样式使用 `.dotx` 模板。

**[下载插件](https://github.com/PharmaRA/RATools-for-Word/releases) · [完整使用手册](docs/manual/index.md) · [安装指南](docs/manual/installation.md) · [常见问题](docs/manual/faq.md)**

## 可以用来做什么

| 使用需求 | 主要功能 | 详细教程 |
| --- | --- | --- |
| 统一文档格式 | 中英文模板、编号/无编号标题 1–9、正文及列表样式、字体段落快捷操作 | [模板与样式](docs/manual/templates.md) |
| 整理图表和页面 | 表格文字样式、表格自动调整、图片段落、页边距、页眉页脚链接 | [表格图片](docs/manual/tables-images.md)、[页面布局](docs/manual/page-layout.md) |
| 管理题注与引用 | 生成题注及引用、粘贴并重连引用、恢复题注标签、目录、链接标蓝与域格式保护 | [题注与引用](docs/manual/references.md) |
| 审阅和制作清洁稿 | 按审阅者筛选、仅接受格式修订、批量接受全部修订并删除批注、高亮检测报告 | [审阅指南](docs/manual/review.md) |
| 批量处理文件 | Word 转 PDF、规范文件名、批量设置页边距、重命名当前文件 | [批量模式与保存规则](docs/manual/batch-basics.md) |
| 整理文本与样式 | 提取缩略语、科学术语下标、标题编号转文本、未使用样式清理 | [全部宏功能](docs/manual/interface.md) |

还提供可搜索的宏管理器和六行四列快捷悬浮窗。每个功能的入口、操作步骤、结果和限制见使用手册。

## 开始使用

1. 从 [Releases](https://github.com/PharmaRA/RATools-for-Word/releases) 下载 `.dotm` 插件和所需 `.dotx` 模板，放到同一个专用目录，例如 `D:\RATools`。
2. 按[图文安装指南](docs/manual/installation.md)设置 Word 启动位置和受信任位置，然后重新启动 Word。
3. 打开普通文档，在 RATools 的“加载”下拉菜单选择中文或英文模板，再应用需要的样式。
4. 按[第一次使用](docs/manual/quick-start.md)完成一份带标题、题注和引用的练习文档。

日常写作请打开普通文档，无需双击 `.dotm`。批量工具是否写回原文件各不相同，运行前先阅读对应的保存规则。

## 环境与版本

- 项目已有使用说明所列环境为 **Windows 10/11 + Microsoft 365 桌面版 Word**。
- 其他 Office 版本的完整兼容性需另行验证；网页版 Word、Mac Word 和 WPS 不属于已确认的运行环境。
- 当前源码版本标识为 **v0.8.8**。下载时以 Releases 附件与发布说明为准，已加载插件的版本可在“关于”中查看。
- 版本变化见 [CHANGELOG](CHANGELOG.md)。

## 使用手册与开发文档

- [完整使用手册](docs/manual/index.md)：安装、每项功能、常见工作流程和排错。仓库内 Markdown 可直接阅读。
- [自定义模板与扩展宏](docs/manual/customization.md)：调整 `.dotx` 样式或添加自己的宏。
- [开发文档索引](docs/README.md)：源码编码、构建测试、发布与维护说明。
- [文档站构建与发布](docs/docs-site.md)：Astro Starlight 本地预览及 GitHub Pages 配置。站点部署后地址为 `https://PharmaRA.github.io/RATools-for-Word/`；当前分支尚未发布。

项目采用源码优先流程：VBA 标准模块、类和窗体分别位于 `modules/`、`class_modules/`、`userforms/`；`dotm/` 为插件解包内容，`template/` 为样式模板。参与开发前请阅读[编码契约](docs/vba-source-encoding.md)与[构建测试指南](docs/local-dotm-build-testing.md)。

## 反馈

问题或建议可提交到 [Issues](https://github.com/PharmaRA/RATools-for-Word/issues)。请注明插件版本、Word 版本、功能名称及复现步骤；示例文档请先移除敏感内容。
