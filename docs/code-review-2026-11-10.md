# 2026-11-10 代码审阅与逐项修复计划

基线：`12e6c4f`，以审阅时工作区为准，含尚未提交的快捷工具栏修改。
用户要求：先记录发现，再按优先级逐项修复或优化；每项验证并 review 后单独提交。

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
| R05 | P1 | `Core_Session.bas`：Save 错误吞掉后仍不保存关闭；清理/边距调用方报成功 | 保存失败返回失败并保留修改；批处理统计准确；正常保存及用户已有文档行为不退化 | 修复中，未提交；Word 许可错误阻断回归，review 待办见下 |
| R06 | P1 | `scripts/RATools.Build.psm1`：先删除旧 dotm 再解压校验，允许目标等于 SafeRoot | 坏包/中断保留旧目录；拒绝根目录、输入包位于替换目标、链接目标；成功替换可恢复且不用物理删除 | 已修复，构建回归和 review 通过 |
| R07 | P1 | `Macro_BatchConvertWordToPDF.bas`：同主文件名不同扩展名输出到同一 PDF | 批次内冲突和已有输出不无提示覆盖；每个成功结果对应独立有效 PDF | 已修复，真实导出及竞争发布回归、review 通过 |
| R08 | P2 | 同模块：导出结束统一解锁全部域 | 混合 Locked 状态在成功和异常路径均恢复，覆盖正文/页眉/文本框 | 已修复，真实 Word 回归和 review 通过 |
| R09 | P2 | `Macro_RemoveUnusedStyles.bas`：只在正文检查样式是否使用 | 页眉/页脚/脚注/文本框使用中的样式保留；真正未用样式正常清理 | 待处理 |
| R10 | P2 | `Macro_ReviewerFilter.bas`：正文修订计数为零提前返回 | 只有页眉或页脚修订且零批注时仍列出审阅者；无标记文档仍返回空 | 待处理 |
| R11 | P2 | `Macro_BatchRenameFiles.bas`：大小写重命名将源文件自身判为冲突 | `Report.DOCX` 改为 `report.docx` 不加后缀；真正冲突仍不覆盖 | 待处理 |
| R12 | P2 | `Core_Window.bas`：仅凭尺寸和矩形相交判定 tooltip，桌面窗口可能误置顶 | 外部进程/非提示窗口被排除；真实 Word tooltip 能置顶且不抢焦点 | 待处理 |
| R13 | P2 | `Install-RAToolsBuildHook.ps1`：构建版本来自未暂存 CHANGELOG | 未暂存/部分暂存 CHANGELOG 不产生版本不一致的提交；完整暂存正常 | 待处理 |

## 审阅证据与限制

三名子代理分别只读审阅批处理、引用/内容、界面；构建发现另经独立复核。
初次审阅未执行 Word 集成测试、未改动项目文件。
路径函数的只读实测确认：接受目标等于根目录，拒绝相邻目录。
Microsoft 文档确认 [CopyHere 不通知复制完成](https://learn.microsoft.com/en-us/windows/win32/shell/folder-copyhere)及 [Document.Revisions.Count 只计正文](https://learn.microsoft.com/en-us/office/vba/api/word.revisions)。

待确认观察：PDF 批处理直接打开并关闭文件，未记录是否为用户已打开的文档。须在涉及 PDF 生命周期的修复中验证；尚未按确定缺陷另行计数。

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

### R05：保存失败不能丢弃文档（进行中）

- 真实 Word 失败复现：DocumentBeforeSave 取消保存后，旧 CloseDocumentQuietly 仍丢弃关闭文档。
- 当前未提交实现：新增 TrySaveDocument，检查保存结果和 Saved；关闭函数返回 Boolean；清理/边距调用方使用保存关闭结果统计失败；保存失败显示保留文档。
- 已执行通过的局部用例：取消保存后保留文档和待保存正文。
- 随后 Word 在创建/编辑下一份测试文档时报告“此方法或属性无效，因为 使用此应用程序的许可协议已经过期。”Quit 也返回相同错误。结束专属测试实例后重试仍复现。正常落盘、两个批处理结果尚未验证，不能计为完成。
- review_batch 待办：关闭被事件取消的分支也要显示隐藏文档；已补入该测试但尚未执行。另需验证 Save 抛异常和用户已打开文档的清理分支。全部通过后再复审、提交。
- 接续命令：`powershell.exe -NoProfile -ExecutionPolicy Bypass -File tests\BatchSaveSafety.Tests.ps1`。未执行任何许可配置修改。
- 后续定位：独立隐藏 Word 的创建、页边距设置和保存通过，不能把报错直接认定为全局许可失效。改进测试诊断后，先发现已打开文件哈希检查受锁影响，已调整为关闭夹具后检查；随后实际失败位于保留文档的可见性断言，Word 开始拒绝调用且退出仍报许可过期。此验收仍未通过，需要继续定位显示窗口相关环境行为。
- 已将恢复窗口访问改为文档自身 `Windows(1)`，关闭取消也进入相同恢复分支；这些更改仍未通过完整验收，不提交。R05 阻塞期间先推进同为 P1、无 Word 依赖的 R06。
- 二进制 dotm 仍保留用户原有未提交内容，不包含本轮源码修复。

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
