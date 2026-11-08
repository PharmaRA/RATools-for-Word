# 2026-11-10 代码审阅与逐项修复计划

基线：`12e6c4f`，以审阅时工作区为准，含尚未提交的快捷工具栏修改。
用户要求：先记录发现，再按优先级逐项修复或优化；每项验证并 review 后单独提交。

当前：原 13 项均已修复并分别提交；最终验收中确认 PDF 文档所有权观察为 R14，另行修复并独立提交。整体验证和成品位置见文末。

## 执行约束

- 不混入原有未提交修改；重要文件先备份到 `backups/code-review-2026-11-10/`。
- 不物理删除文件；临时产物清理使用回收站，失败时记录保留位置。
- VBA 使用 GBK/936、CRLF；含中文的 PowerShell 使用 UTF-8 BOM。
- 每项先建立失败复现，再修复、运行针对性回归、进行子代理只读 review、更新本表并提交。
- 发现不成立时记录撤回依据；不能用静态分析或测试片段断言代替真实行为验收。
- 提交使用 Fiveonine 身份。本轮暂不运行含物理删除的原构建钩子；以明确的测试证据替代。产物同步在构建安全修复后处理，不覆盖原有未提交模板修改。

## 发现清单和验收

编号沿用审阅报告，按表中顺序处理。P1 为优先修复，P2 为后续修复。

| 编号 | 优先级 | 问题与证据入口 | 必须验证的行为 | 状态 |
| --- | --- | --- | --- | --- |
| R01 | P1 | `modules/Macro_AcceptFormatRevisions.bas`：异步 ZIP 仅凭目录出现/300ms 大小稳定判断完成，随后覆盖原文档；恢复提示无对应恢复 | 完整保留包部件和二进制数据；仅接受字符/段落格式修订；损坏输入、打包失败、替换失败不得破坏原文件；完成后可重新打开 | 已修复并通过针对性回归、review；启动限制见下 |
| R02 | P1 | 同模块 `FindLocalSyncCopy`：精确映射失败后按同名文件寻找本地副本 | 同名不同目录/账户不误处理；无法确定映射时对当前 Document 回退 | 已修复，回归和 review 通过 |
| R03 | P1 | `Macro_PasteRelinkCrossReference.bas`：缓存绝对坐标后逐条改写域代码 | 多个题注前有多条引用，新旧书签名长短不同时仍指向正确题注；原文不变、一次撤销 | 已修复，Word 回归和 review 通过 |
| R04 | P1 | 同模块 `MatchPastedCaption`：重复标签/描述取首个未使用项 | 反序引用重复描述、空描述题注不串号；单一明确匹配仍正常；优化候选索引但不牺牲唯一性 | 已修复并建立候选索引，Word 回归和 review 通过 |
| R05 | P1 | `Core_Session.bas`：Save 错误吞掉后仍不保存关闭；清理/边距调用方报成功 | 保存失败返回失败并保留修改；批处理统计准确；正常保存及用户已有文档行为不退化 | 已修复，真实 Word 回归、全工程编译和 review 通过 |
| R06 | P1 | `scripts/RATools.Build.psm1`：先删除旧 dotm 再解压校验，允许目标等于 SafeRoot | 坏包/中断保留旧目录；拒绝根目录、输入包位于替换目标、链接目标；成功替换可恢复且不用物理删除 | 已修复，构建回归和 review 通过 |
| R07 | P1 | `Macro_BatchConvertWordToPDF.bas`：同主文件名不同扩展名输出到同一 PDF | 批次内冲突和已有输出不无提示覆盖；每个成功结果对应独立有效 PDF | 已修复，真实导出及竞争发布回归、review 通过 |
| R08 | P2 | 同模块：导出结束统一解锁全部域 | 混合 Locked 状态在成功和异常路径均恢复，覆盖正文/页眉/文本框 | 已修复，真实 Word 回归和 review 通过 |
| R09 | P2 | `Macro_RemoveUnusedStyles.bas`：只在正文检查样式是否使用 | 页眉/页脚/脚注/文本框使用中的样式保留；真正未用样式正常清理 | 已修复，真实 Word 回归和 review 通过 |
| R10 | P2 | `Macro_ReviewerFilter.bas`：正文修订计数为零提前返回 | 只有页眉或页脚修订且零批注时仍列出审阅者；无标记文档仍返回空 | 已修复，真实 Word 三例和 review 通过 |
| R11 | P2 | `Macro_BatchRenameFiles.bas`：大小写重命名将源文件自身判为冲突 | `Report.DOCX` 改为 `report.docx` 不加后缀；真正冲突仍不覆盖 | 已修复，真实文件回归和 review 通过 |
| R12 | P2 | `Core_Window.bas`：仅凭尺寸和矩形相交判定 tooltip，桌面窗口可能误置顶 | 外部进程/非提示窗口被排除；合法 tooltip 能置顶且不抢焦点 | 已修复，实际 Win32 窗口回归和 review 通过；视觉限制见下 |
| R13 | P2 | `Install-RAToolsBuildHook.ps1`：构建版本来自未暂存 CHANGELOG | 未暂存/部分暂存 CHANGELOG 不产生版本不一致的提交；完整暂存正常 | 生成器已修复，真实钩子回归和 review 通过 |
| R14 | P1 | 最终验收确认：PDF 批处理关闭用户已打开的未保存文档 | 已有文档仍成功导出，成功/失败均保留正文、待保存状态和原域锁；只关闭本次打开的文档 | 已修复，真实 Word 回归和 review 通过 |

## 审阅证据与限制

三名子代理分别只读审阅批处理、引用/内容、界面；构建发现另经独立复核。
初次审阅未执行 Word 集成测试、未改动项目文件。
路径函数的只读实测确认：接受目标等于根目录，拒绝相邻目录。
Microsoft 文档确认 [CopyHere 不通知复制完成](https://learn.microsoft.com/en-us/windows/win32/shell/folder-copyhere)及 [Document.Revisions.Count 只计正文](https://learn.microsoft.com/en-us/office/vba/api/word.revisions)。

原待确认观察“PDF 批处理可能关闭用户已有文档”已在最终验收真实复现，升级为 R14 并修复，见文末。

## 验证与提交记录

- 文档阶段：记录 13 项发现、验收条件及原有工作区修改边界；尚无修复完成。
- 本文只记录已执行的验证。源码修复、自动测试、真实 Word、打包产物分别记录，不互相替代。

### R01：同步处理和原件恢复

- 失败复现：旧生成脚本对损坏的 `word/document.xml` 返回成功，`MalformedXmlPreservesOriginal` 断言失败。
- 实现：独立 PowerShell/.NET 进程同步修改 ZIP；关闭后重开并核对每个部件的 SHA256；同卷 `File.Replace` 保留原件副本；修改前校验源文件哈希；Word 重开失败时尝试恢复原件。备份文件不覆盖，临时材料仅送回收站。
- review 补充并修复：保存事件取消后检查 `doc.Saved`；事务路径碰撞时只清理本次拥有的文件；回滚状态未确认时保留所有恢复材料。
- 命令：`powershell.exe -NoProfile -ExecutionPolicy Bypass -File tests\AcceptFormatRevisions.Tests.ps1`，退出码 **0**。
- 通过：损坏 XML 保留原件、全包/正文/内容修订/页眉格式处理/二进制字节保留、锁定目标替换失败保护、已有事务材料保持、恢复副本哈希一致、Word 编译、真实 DOCX 处理后重新打开、Word API 回退、保存事件取消时保留未保存文档。
- 环境限制：本机 Word 内启动 PowerShell 返回“拒绝的权限”。真实文档的快速处理验收使用 **VBA 生成脚本 → 外部 PowerShell 执行 → Word 重新打开**；未声称 Word 内快速启动成功。进程不可用时的 Word API 回退已通过故障注入验收。正常使用本机可能走较慢的兼容路径。
- 未实测：提交后结果写入失败且回滚同时失败、Word 拒绝重开后的恢复分支；这些分支已接受只读代码 review，不能等同于故障注入通过。
- review：`review_reference` 独立读取实现并复查三个阻塞的修复后通过；最终新增测试执行证据由主代理核实。
- 编码和 PowerShell 解析检查通过；已注册到 `Run-AllTests.ps1`。未运行含物理删除清理的全套旧测试，未同步现有未提交的二进制模板。
- 独立提交：`2340c76`。

### R02：禁止猜测云端文件身份

- 旧版失败复现：模拟云端 URL 无精确映射时，入口选择同步根下另一份 `Report.docx`，而不是当前文档回退，测试明确失败。
- 移除同名文件搜索及其私有辅助函数。精确映射和目标存在检查保留；无法确定文件时使用入口捕获的当前 `Document`。
- `powershell.exe -NoProfile -ExecutionPolicy Bypass -File tests\FormatRevisionPath.Tests.ps1`：退出码 0，映射缺失、精确映射、精确文件缺失三例通过。
- 同模块 `AcceptFormatRevisions.Tests.ps1` 全部针对性回归再次通过；R01 的本机启动限制保持记录。
- `review_reference` 只读 review 通过。R02 测试仅替换外部发现、URL、UI 和写入边界，证明路径选择与当前文档回退决策；不把夹具未变当作真实修改验收。
- 独立提交：`4b4b74d`。

### R03：先建立书签再改写引用

- 真实 Word 失败复现：两条前置引用采用 `A` 和 `LongCaptionBetaSourceName`，旧实现第二条粘贴引用返回 `ha captionTable 4`，混入上一题注文字。
- 实现分为两步：先建立所有新书签并收集待改写域，再统一替换书签名。编辑引起的位置变化由 Word 的书签自动跟随。
- `PasteRelinkSafety.Tests.ps1` 通过：原引用 Table 1/2 保持，新引用准确为 Table 3/4；单次 Undo 恢复原文文本（未据此宣称所有格式和书签状态完全一致）。
- `PasteRelinkCrossReference.Tests.ps1` 和 `PasteRelinkCrossReferenceCases.Tests.ps1` 均退出 0，覆盖单题注、不同引用形状、无题注、文档中间粘贴。
- `review_reference` 独立只读 review 通过；新用例注册到总测试入口。
- 独立提交：`28f3bfe`。

### R04：唯一对应检查及题注索引

- 真实 Word 失败复现：重复描述反序引用会改写错误目标；review 又指出“仅一个粘贴候选、多个源身份”的误认领，补充用例也在修复前失败。
- 按标签和规范化描述建立粘贴候选索引，同时对粘贴范围外的原文 SEQ 题注按段落身份去重索引。源或目标不唯一时保留原引用，并在操作结束后提示手动确认。
- 原来逐引用遍历全部题注改为一次建立索引；算法开销为原文域、粘贴题注和引用数量的线性遍历。未测量实际 Word 耗时，不报告具体提速。
- `PasteRelinkSafety.Tests.ps1`：长短书签名、重复描述、空描述、一个粘贴候选与多个源身份均通过。
- `PasteRelinkCrossReference.Tests.ps1`、`PasteRelinkCrossReferenceCases.Tests.ps1` 再次退出 0，明确唯一的匹配仍能重连。
- `review_reference` 复查通过。仅有一条外部 REF 的变体未单独实测；实现索引不依赖引用数量。
- 独立提交：`94e81d5`。

### R05：保存失败保留文档并返回失败

- 真实 Word RED：DocumentBeforeSave 取消后旧函数仍不保存关闭。新增 TrySaveDocument 检查保存异常及 doc.Saved；关闭函数返回结果且检查关闭取消；两个批处理根据实际结果统计。
- 保存/关闭失败保留文档与待保存内容；交互 Word 尝试显示该文档，隐藏自动化保持应用可见性。页边距部分设置失败也使用同一原则，显示失败不会中断批处理。
- `BatchSaveSafety.Tests.ps1` 退出 0：清理与边距的取消保存/正常落盘、用户已打开文档的取消保存、未保存文档取消、关闭取消、部分边距失败、持久化抛错后的失败结果/错误信息/内容保留。
- Save 抛错与部分边距失败在外部操作边界注入，真实错误处理和文档生命周期仍执行；交互式 Visible=True 的窗口显示分支仅经源码 review，不称为 UI 实测。
- 先前许可报错发生在隐藏自动化被强制显示后的调用中；独立创建/保存通过，修正恢复可见性及测试文件锁处理后整套通过，不再认定为全局许可阻塞。未改 Office 许可设置。
- `SourceCompile.Tests.ps1` 全部标准模块、类和窗体编译探针退出 0；`review_batch` 最终复查通过。
- 本项因早先 Word 验收障碍在 R06–R13 后完成，保持独立提交。
- 独立提交：`09f1811`。

### R06：可恢复构建发布

- 旧版失败复现：坏包在解压报错前先移走原目录。基线测试只把物理删除替换成回收站操作，保留原控制流；未执行物理删除。
- 新实现先在唯一临时目录解压并验证必要部件，然后将旧目录移入安全根下的 `backups/`，再发布；最终重命名失败时恢复旧目录。成品文件也先复制校验、备份旧输出再发布，失败时恢复。
- 拒绝安全根/文件系统根、位于待替换目录内的输入包、输出位于源目录、链接及链接祖先；检查完成前不创建输出父目录。临时材料只送回收站，回收失败报告保留位置。
- `BuildSafety.Tests.ps1` 退出 0：坏包、缺部件、危险目标、目录发布故障回滚、文件发布故障回滚、原件备份、打包读取失败保留成品、枚举失败在 Continue 下仍停止、受控 junction 拒绝及拒绝前无目录创建。
- 用合成仓库执行 `Build-RAToolsDotm.ps1 -SkipVbaImport` 的完整打包、目录同步和输出发布通过；未对真实工作区 dotm 执行覆盖。
- `BuildRATools.Tests.ps1` 全部通过，测试自身清理也已改为回收站。枚举故障注入首先命中安全检查阶段；打包枚举的 Stop 策略由源码 review 确认，没有声称分别注入了两个阶段。
- `review_batch` 复查通过；补充入口链接输出用例 RED→GREEN 后再确认。R05 测试入口变更与本项分开暂存。
- 独立提交：`64c1c14`。

### R07：PDF 不覆盖发布

- 真实 Word 失败复现：同目录 Report.doc 与 Report.docx 导出时覆盖既有 Report.pdf。
- 先独占创建临时输出文件，再真实导出 PDF；发布使用不覆盖目标的 MoveFile，已存在的名称自动加 `_1`、`_2` 等后缀。选择后被其他写入者抢占则重新选名；成功计数发生在发布之后，报告记录实际路径。
- `PdfOutputSafety.Tests.ps1` 退出 0：原 PDF 哈希不变、两份源文档生成独立有效 PDF、报告路径准确；同名目录保持、竞争者抢占文件保持、真实发布重试生成有效 PDF。
- `BatchConvertWordToPDF.Tests.ps1` 原有回归退出 0；`review_batch` 独立只读 review 通过。
- 验证通过 ConvertActiveDocument 覆盖共用导出函数；未声称动态覆盖非冲突发布失败和临时回收失败，这两条异常路径仅经源码 review。
- 独立提交：`0464d38`。

### R08：逐域恢复原锁定状态

- 真实 Word 失败复现：默认导出后原先锁定的正文域被解锁。
- 先完整捕获各 Field 对象及 Locked 布尔值，再锁定；成功及异常都逐项恢复。恢复前保存原始错误；恢复不完整计失败并提示。
- `PdfFieldState.Tests.ps1` 退出 0：真实 PDF 成功导出和注入导出失败后，正文混合状态、页眉锁定、文本框未锁定均恢复；错误信息保持。
- review 发现并修复可映射错误吞掉恢复警告：先转换导出错误再附加恢复警告。复合失败测试注入恢复结果 False，旧消息顺序 RED、修复后 GREEN；不宣称验证了 Word Locked 赋值实际抛错。
- `PdfOutputSafety.Tests.ps1` 与原 `BatchConvertWordToPDF.Tests.ps1` 回归通过。去掉旧测试中强制要求全量解锁的源码断言，保留辅助函数行为测试。
- `review_batch` 独立复查通过；未改批量导出中原有文档所有权行为。
- 独立提交：`3ac1420`。

### R09：保护非正文及依赖中的样式

- 真实 Word RED→GREEN：仅用于页眉的样式、页脚表格样式、被使用中派生样式引用的基样式、页眉文本框样式，均先复现误删后修复。
- 遍历 StoryRanges/NextStoryRange，并显式检查正文、各节页眉页脚 Shapes 及组合/画布。段落/表格直接比较样式，字符样式使用格式查找；单用 Find 在实测页眉中会返回 False，不能据此删除。
- 基样式只要被其他样式引用就保留（包括未直接应用的派生样式）；任何读取异常保留样式，避免误删。
- `UnusedStylesSafety.Tests.ps1` 退出 0：正文、两节页眉、页脚、脚注、正文文本框、页眉文本框、字符和表格样式、继承基样式均保留；孤立 Unused-F 被清理。
- `review_reference` 复查通过。组合/画布递归及异常保留分支未单独运行；保守处理不支持的图形可能留下部分未用样式。
- 独立提交：`0f5916f`。

### R10：收集非正文修订作者

- 真实 Word RED：仅页眉修订且正文修订/批注均为零，返回 0 位作者。
- 移除只看正文计数的提前返回，保留原 XML 收集和对象模型回退流程。
- `ReviewerStorySafety.Tests.ps1` 退出 0：仅页眉、仅页脚、无标记分别返回 1/1/0，不输出实际作者信息。
- `review_ui` 独立只读 review 通过。
- 独立提交：`22ac0ac`。

### R11：大小写重命名排除源文件自身

- 真实 Windows 文件系统 RED→GREEN：Report.DOCX 原来成为 report_1.docx，修后成为准确小写的 report.docx。
- 冲突循环将同一路径的大小写差异排除，其他文件/目录仍生成后缀；复制回退明确禁止覆盖。
- `RenamePathSafety.Tests.ps1` 退出 0：调用完整宏，验证准确文件名、数量、三个文件内容，以及 A B.DOCX 与既有 a-b.docx 冲突时生成 a-b_1.docx。
- 原 `BatchRenameFiles.Tests.ps1` 的 10 个字符串清洗用例通过；`review_ui` 只读 review 通过。
- 运行范围为普通 Windows 不区分大小写目录。目录名称冲突、复制回退失败和启用大小写敏感目录未单独实测。
- 独立提交：`af0f8c3`。

### R12：校验提示窗口身份

- 实际 Win32 窗口 RED：Word 内普通 STATIC 窗口因几何条件符合被判为提示。
- 修复先确认候选及窗体均属于当前 Word 进程，再检查 F3 Tooltip 类名前缀；最后才检查尺寸和矩形相交，API 失败拒绝候选。
- `TooltipIdentity.Tests.ps1` 退出 0：当前进程普通窗口、外部进程同 tooltip 类名前缀窗口被排除；合法合成提示经实际枚举/SetWindowPos 置顶，其余测试窗口不置顶且前台焦点不变。仅替换窗体发现入口，测试窗口均在屏幕外且不激活。
- `QuickToolbarWindow.Tests.ps1` 原有 74 条断言通过；`review_ui` 只读 review 通过，声明按 VBA7 LongPtr/Long 类型审阅。
- 未验证真实 MSForms 提示出现时机和视觉效果，也未分别运行 32/64 位 Office；合成窗口身份测试不替代这些人工兼容性检查。
- 独立提交：`5f6ce77`。

### R13：提交前检查版本输入的暂存一致性

- 临时 Git 仓库安装实际生成钩子，旧版在 CHANGELOG index=v1.1.0、working=v1.2.0 时仍构建，测试 RED。
- 把 CHANGELOG.md 加入 VBA 输入的未暂存及未跟踪检查，检查在构建之前执行。
- `BuildHookSafety.Tests.ps1` 退出 0：部分暂存拒绝且不构建；完整暂存成功且构建读取 v1.2.0；VBA 未暂存仍拒绝。构建外部边界替身仅记录调用/输入，Git 检查运行真实钩子。
- `review_ui` 只读 review 通过。未跟踪 CHANGELOG 分支仅经源码检查，未独立运行。
- 当前项目实际已安装钩子尚未重装，不能称其保护已更新；本项交付为钩子生成器与回归测试。

- R13 独立提交：`5d72eca`。

### R14：PDF 批处理保留用户文档

- 真实 Word RED：用户已有 Report.docx 并含未保存编辑，ConvertOneFile 返回后该文档被关闭。
- 打开前识别并复用已有文档，保留其可见性；仅关闭批处理自行打开的对象。已有文档的默认 WYSIWYG 导出捕获并恢复原域锁状态，异常路径同样恢复并保留原错误信息。
- `PdfOutputSafety.Tests.ps1` 退出 0：真实导出成功计数 1/0，原文档保持打开、待保存文本一致、混合域锁恢复；注入导出失败计数 0/1，原错误、文档、文本和域锁均保持。
- `PdfFieldState.Tests.ps1` 回归退出 0；`review_batch` 只读 review 通过。独立成品需在本项提交后重新构建。
