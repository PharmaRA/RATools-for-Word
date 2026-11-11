# RATools for Word 使用手册

RATools for Word 是面向药品注册资料编写的 Windows Word 插件。它把常用排版操作集中在一个功能区，并提供样式、题注引用、审阅清理和批量文件处理工具。

**第一次使用：** 从[安装插件](installation.md)开始，然后完成[第一次使用](quick-start.md)中的练习。已经安装的用户，可以通过左侧目录或搜索框查找功能名称。

## 按你要完成的工作查找

| 我想做什么 | 从这里开始 |
| --- | --- |
| 把整份资料的标题、正文和列表格式统一起来 | [加载模板](templates.md) → [应用样式](styles.md) |
| 调整表格、图片、页眉页脚和页面尺寸 | [表格与图片](tables-images.md)、[页面与页边距](page-layout.md) |
| 给图表编号，并在正文中引用 | [题注与引用基础](references.md)、[生成题注与引用](caption-pairs.md) |
| 复制已有图表和正文，让新引用跟随新图表 | [粘贴并重连引用](paste-relink.md) |
| 保留文字修订，先接受格式修改 | [接受格式修订](accept-format-revisions.md) |
| 只查看某位同事的批注和修订 | [按审阅者筛选](review.md#reviewer-filter) |
| 把审阅稿整理为清洁稿 | [接受修订并删除批注](clean-documents.md)、[检测高亮](highlights.md) |
| 一次导出多个 Word 文件为 PDF | [Word 批量转 PDF](pdf-export.md) |
| 规范文件名或修改当前文件名称 | [批量修改文件名](batch-rename.md)、[当前文件工具](current-file.md) |
| 提取英文缩略语、整理科学术语的下标 | [提取缩略语](abbreviations.md)、[科学术语下标](scientific-terms.md) |

## 如何阅读本手册

- **入口**说明按钮或宏在哪；很多功能区按钮只显示图标，把鼠标停上去可以查看名称。
- **处理范围**说明作用于选区、当前文档还是多个文件。没有选中文字时，某些功能会处理正文整体。
- **保存与结果**说明内容是否已经写回原文件、是否生成新文件。各工具的保存方式并不相同。
- **示例和限制**帮助你判断结果是否符合预期。批处理前可以先用一份副本熟悉流程。

文中的“宏列表”指 RATools 的宏管理器；“Word 原生”表示按钮打开的是 Word 自带功能，其窗口文字可能随 Office 版本变化。无需编写 VBA 即可使用手册介绍的功能。

## 版本与适用环境

本手册按仓库 **v0.8.8 标识的当前源码**编写。下载的发布包可能早于源码中的修复；遇到行为不同，先在“关于”中查看版本，并查阅[更新日志](https://github.com/PharmaRA/RATools-for-Word/blob/main/CHANGELOG.md)。

项目已有使用说明所列环境为 Windows 10/11 的 Microsoft 365 桌面版 Word。其他 Office 版本的完整兼容性尚不能据此保证；网页版 Word、Mac Word 和 WPS 不属于本手册已确认的运行环境。

安装与模板截图来自仓库既有教程，可能显示旧版按钮布局。请以本文中的按钮名称和你的实际界面为准。

需要开发、构建或维护插件，请进入[开发文档索引](https://github.com/PharmaRA/RATools-for-Word/blob/main/docs/README.md)。
