# 宏管理器、悬浮窗与关于

<a id="macro-manager"></a>
## 宏管理器

**入口：RATools → 宏列表中的宏列表按钮。** 打开的窗口标题为“宏管理器”。

1. 在列表中单击功能名称，阅读下方的详细说明。
2. 点击运行按钮，或直接双击该功能。
3. 若功能较多，在搜索框中输入关键词，例如“PDF”“页边距”“修订”。搜索同时匹配中文名称、宏过程名和功能说明，不区分英文大小写。
4. 清空搜索框可以重新显示全部功能；刷新按钮重新载入列表并清空搜索。
5. 窗口可调整大小，也允许回到文档继续编辑。运行前先确认当前活动文档和选区。

宏管理器只是启动入口。是否出现模式选择、是否保存原文件、是否生成报告，都由具体工具决定。

### 宏列表中的全部功能

| 功能 | 详细说明 |
| --- | --- |
| 超链接和域批量设置为蓝色 | [统一链接颜色](field-format.md#blue-links) |
| Word批量转PDF | [选择导出方式、输出路径与报告](pdf-export.md) |
| 批量修改文件名 | [清洗名称、处理同名与占用](batch-rename.md) |
| 标题自动编号转文本 | [固定标题编号](heading-numbers.md) |
| 重命名当前文件 | [保存新名称并处理原文件](current-file.md#rename-current) |
| 批量设置页边距 | [所有节四边设为 2.54 厘米](page-layout.md#batch-margins) |
| 一键表格自动调整 | [调整当前文档的表格](tables-images.md#autofit-all) |
| 批量接受修订并删除批注 | [制作清洁稿及保存规则](clean-documents.md) |
| 页眉和页脚设置为“链接到前一节” | [统一节间链接](page-layout.md#link-sections) |
| 清理未使用的模板样式 | [清理未使用的 -F 样式](unused-styles.md) |
| 批量检测高亮内容 | [检测、报告与跳转](highlights.md) |
| 提取缩略语 | [提取候选列表并排序](abbreviations.md) |
| 标准化科学术语下标 | [支持的术语与格式规则](scientific-terms.md) |
| 接受格式修订 | [保留内容修订，接受格式变化](accept-format-revisions.md) |
| 按审阅者筛选 | [按作者控制标记显示](review.md#reviewer-filter) |
| 粘贴并重连引用 | [让粘贴出的引用对应新题注](paste-relink.md) |
| 生成题注与引用 | [一次建立题注和引用](caption-pairs.md) |

<a id="quick-toolbar"></a>
## 快捷悬浮窗

**入口：RATools → 宏列表 → 快捷悬浮窗。** 再点击可以隐藏；窗口自身的关闭按钮也可以关闭窗口。

悬浮窗提供六行四列共 24 个按钮。它不会阻塞文档编辑，可拖动到顺手的位置，并保存上次位置。先在文档中选好目标，再点击悬浮窗；鼠标停在图标上可以查看名称和说明。

| 行 | 从左到右的四个功能 | 使用说明 |
| --- | --- | --- |
| 第 1 行 | 编号标题 1、2、3、4 | [编号标题](styles.md#numbered-headings) |
| 第 2 行 | 无编号标题 1、2、3、4 | [无编号标题](styles.md#unnumbered-headings) |
| 第 3 行 | 正文、表标题、图标题、段前分页 | [基础样式](styles.md#basic-styles)、[表格图片](tables-images.md)、[段前分页](formatting.md#page-break-before) |
| 第 4 行 | 格式刷、段落设置、清除格式、设为蓝色 | [字体与段落](formatting.md) |
| 第 5 行 | 根据窗口调整表格、插入题注、插入交叉引用、更新域 | [当前表格调整](tables-images.md#table-tools)、[题注与引用](references.md) |
| 第 6 行 | 粘贴并重连引用、超链接和域批量设置为蓝色、接受修订并删除批注、检测高亮内容 | [重连](paste-relink.md)、[标蓝](field-format.md#blue-links)、[清稿](clean-documents.md)、[高亮](highlights.md) |

标题 5–9 和无编号标题 5–9 在功能区中使用。悬浮窗中的批量宏仍会弹出相应模式选择；点击悬浮窗的清稿按钮不意味着只处理选中的文字。

切换 Word 文档后，按钮作用于当前活动文档。文档都已关闭时，应先打开文档再使用编辑类按钮。

<a id="about-update"></a>
## 关于与检查更新

**入口：RATools → 关于。**

| 操作 | 结果 |
| --- | --- |
| 查看当前版本 | 显示已加载插件的版本标识；反馈问题时可记录此版本 |
| GitHub | 用浏览器打开项目 GitHub 仓库 |
| Gitee | 用浏览器打开项目 Gitee 仓库 |
| 检查更新 | 查询发布版本；发现新版本后，可按弹窗提示打开 GitHub 或 Gitee 的发布详情页 |
| 关闭 | 退出关于窗口 |

检查更新需要联网。网络失败不代表当前版本已经最新，可手动查看[发布页面](https://github.com/PharmaRA/RATools-for-Word/releases)。插件不会因此自动替换当前文件，后续按照[更新步骤](installation.md)安装。
