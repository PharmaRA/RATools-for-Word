# 生成题注与引用：验收记录

日期：2026-11-10。验证环境：本机 Windows / Word COM。

## 交付

- 版本：`v0.8.8`。成品：`dist/RATools_v0.8.8.dotm`，340853 字节。
- SHA256：`482EA1F73251CB22E7AAAB27A2ED815BAA58170F7FFFA49BC836AA6E822B0076`。
- 入口：RATools「引用」组及宏列表中的「生成题注与引用」。
- 使用说明：项目根目录 `README.md` 的「生成题注与引用」小节。
- 发布构建使用 `Build-RAToolsDotm.ps1 -Version v0.8.8`。成品包含当前源码和既有悬浮窗扩充，已同步回 `dotm/`；VBA 二进制、VBA 数据和 Ribbon XML 与成品逐字节一致。没有替换已安装插件。
- 首次实现前备份：`backups/caption-reference-20261110-164036/`；其中 `before.patch` 保留开始时的工作区差异。发布前备份：`backups/release-v0.8.8-20261110-224116/`。本记录随 v0.8.8 本地提交保留；未推送或发布远程 Release。

## 已通过

`tests/CaptionReferencePair.Tests.ps1` 的以下 10 组真实 Word 用例，在源码导入模式和 `-PackagePath dist/RATools_v0.8.8.dotm` 成品模式均通过：

1. 相邻两段且交叉引用在前、题注在后，真实 SEQ/REF、可跳转引用、原生交叉引用列表、原文保留、一次撤销。
2. 连续编号、同标题独立目标、中文和含空格的自定义标签、空标题、题注与引用分别剪切移动。
3. 段落中间插入、拒绝覆盖文字选区、非法输入及文档保护。
4. 窗体原生标签列表、暂存新标签、取消无插入或标签注册。
5. 窗体生成按钮事件、自定义标签和重新打开后的标签记忆。
6. 表格、页眉和现有域内部插入点保护。
7. 新生成的配对通过原有「粘贴并重连引用」复制，原引用保持原目标、新引用指向新题注。
8. 两位数编号和 Word 原生罗马数字编号设置。
9. 移除旧题注但保留引用后，创建无关题注不会复用旧目标或错误连接旧引用。
10. 引用默认为标准蓝色；编号从 9 更新为 10、100 后，保留蓝色、粗体、斜体和字号；域仍可更新，题注和原文不被染蓝。

初次实现时另外通过 `PasteRelinkCrossReference.Tests.ps1`、`PasteRelinkCrossReferenceCases.Tests.ps1`、`PasteRelinkSafety.Tests.ps1` 三个回归套件，以及 `SourceCompile.Tests.ps1`、`RibbonCallbacks.Tests.ps1`、`ScriptEncoding.Tests.ps1`。未运行其他无关功能的全量测试。

## 本日后续调整：引用在前及工具栏布局

- 生成结果改为首段交叉引用、次段题注；窗体、宏列表提示和 README 示例同步修改。
- 「生成题注与引用」图标改为 `CitationInsert`，在 Word VBA 内实际加载成功；当前 Ribbon 与悬浮工具栏均无其他使用者。
- `HyperlinkInsert` 从 `boxR2` 移到 `boxR1` 末尾；前两行各 5 个按钮，与下一行第五列的 `btnPasteRelinkCrossReference` 对齐。最终模板包内的 Ribbon XML 也已核对。
- 先将段落顺序测试改为引用在前，在旧实现上观察到预期失败；实现调整后，源码模式与成品模式的上述 9 组测试均通过。
- 本轮重新通过全源码编译与 Ribbon 回调检查；独立子代理只读 review 通过，未发现新增阻塞问题。布局位置根据 Ribbon XML 验证，未进行实际屏幕像素检查。
- 本轮修改前备份（含上一版成品）：`backups/caption-reference-order-20261110-220549/`。未新增磁盘临时脚本；构建临时目录由构建器移入回收站。

## Review 与限制

- v0.8.8 发布检查：版本常量、CHANGELOG 和 README 一致；成品内 `GetAppVersion` 返回 `v0.8.8`。`QuickToolbar.Tests.ps1 -DotmPath dist/RATools_v0.8.8.dotm` 的源码及成品检查、`SourceCompile`、`RibbonCallbacks`、`QuickToolbarWindow`、`ScriptEncoding` 均通过。悬浮窗本次待提交改动经独立只读 review，无阻塞问题。
- 本轮成品移动测试最初出现两次间歇失败：一次为“Moving first caption must preserve its reference target”，一次为 Word“此命令无效”。测试已补充执行步骤和实际结果诊断，随后同一成品连续三次完整通过；根因尚未确认，保留为剪切／粘贴自动化验证的已知限制，不能据此宣称该间歇问题已修复。

- 本日格式调整：新生成的 REF 域自动加入 `\* MERGEFORMAT`，结果设置为 `wdColorBlue`（RGB 0、0、255），不锁定域的更新。新增格式用例先在旧实现上失败，修复后源码与成品各 10 组测试通过；全源码编译及本轮独立只读 review 通过。
- 格式调整前备份（含上一版成品）：`backups/caption-reference-format-20261110-223033/`。本轮未新增磁盘临时脚本，构建临时目录由构建器移入回收站。

- 第一轮独立只读 review 发现递增书签名可能被孤立引用复用。新增用例先在真实 Word 中失败，再改用完整 GUID 的 38 字符书签名；修复后通过。
- 第二轮独立只读 review 复核上述修复及回归用例，未发现新增阻塞问题。
- 初版限可编辑文档的正文、表格外普通插入点。移动时应剪切整段，调整顺序后更新域；引用位于题注前方时可能需要再次更新。
- 窗体控件与事件经过运行时验证，但截图预览尝试失败，未完成界面外观人工验收。没有宣称 Office 各版本或 32/64 位完整兼容性验收。
- 本次临时探测、导出及截图脚本已移入回收站，失败预览创建的独立 Word 测试进程已按进程 ID 和启动时间确认后关闭。
