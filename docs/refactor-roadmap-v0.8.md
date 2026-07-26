# RATools-for-Word 重构路线图（v0.8 系列）

> 分支：`worktree-refactor-v0.8`（worktree：`.claude/worktrees/refactor-v0.8`）
> 基线：`75c8b2f`（v0.7.1，main，工作区干净）
> 制定日期：2026-07-26

## 一、重构原则

沿用 RATools-for-PDF 重构的成功经验：

1. **行为不变**：除「阶段 1 已知 bug 修复」外，任何阶段不改变用户可感知的功能行为。
2. **安全网先行**：先建统一测试入口与版本一致性校验，再动代码。
3. **每阶段独立可合并**：每阶段（甚至阶段内每个模块迁移）都是一个可独立构建、测试、合并回 main 的提交，不存在"半成品必须一次性合完"的状态。
4. **验收三件套**：源码测试全绿 → 本地构建成功 → 成品 `dist/*.dotm` 手测关键功能（Ribbon 出现、宏列表、悬浮窗、至少 1 个批处理宏）。
5. **Ribbon XML 零改动约束**：`dotm/customUI/customUI14.xml` 的 onAction/tag 是外部契约。Ribbon 回调按名字在全工程范围解析，因此**把回调 Sub 移动到其他模块不需要改 XML**，只要保持 Public 且全工程唯一。这是拆分 mRibbon 的安全基础。

## 二、现状诊断（7 个核心问题）

### P1. mRibbon.bas 是上帝模块（684 行，全仓变更最频繁：26 次提交）
混杂 6 类职责：Ribbon 生命周期（`mRibbon.bas:16-49`）、样式导入引擎（`60-132`）、约 60 项中英样式名映射 Select Case（`152-212`）、零散文档命令（`415-549`）、宏注册表数据（`562-659`）、Wrapper 适配层（`666-684`）。
另有隐患：公共变量与模块同名 `Public mRibbon As IRibbonUI`（`mRibbon.bas:5`）。

### P2. 5 个批处理宏各自复制同一套样板，且细节互相矛盾
- 模式选择 InputBox 4 份近似拷贝（`Mod_BatchAcceptAndClean.bas:19-25`、`Mod_BatchConvertWordToPDF.bas:51-57`、`Mod_BatchSetMargins.bas:15-20`、`Mod_BatchDetectHighlights.bas:28-34`），`Mod_BatchRenameFiles.bas:32-37` 又是另一种交互模型。
- FileDialog 样板 11 处；过滤串 4 种写法。
- FSO 递归遍历 4+1 份实现，`Mod_BatchSetMargins` 的 `ProcessFolder`（`139-162`）不递归且 `Dir("*.doc*")` 会误匹配 `.dot/.dotx`，与其余宏行为不一致。
- `ScreenUpdating` 开关 29 处，多处无错误保护；已确认泄漏 bug 见阶段 1。
- 公共名冲突：`ProcessFile` 同签名出现在 `Mod_BatchAcceptAndClean.bas:95` 与 `Mod_BatchSetMargins.bas:83`。

### P3. 样式名"-F"体系存在三种独立事实源
- mRibbon 的 60 项映射 Select Case（`mRibbon.bas:152-212`）
- Mod_QuickToolbar 的 GBK 中文字面量（`Mod_QuickToolbar.bas:70-76`）
- Mod_QuickToolbarActions 的 Unicode 码点构造（`Mod_QuickToolbarActions.bas:61-84`）
`TARGET_SUFFIX = "-F"` 常量之外 `Mod_RemoveUnusedStyles.bas:34` 再写一次字面量；`StyleExists` 私有在 mRibbon 里，其他模块用不了。

### P4. 悬浮窗动作分发被历史增量劈成两半
前 8 个动作在 `Mod_QuickToolbar.bas:68-89`，后 13 个在 `Mod_QuickToolbarActions.bas:7-42`，未知 key 二次分发后才报错。测试桩（MsoTestMode）混入生产代码。19 按钮定义表在 3 处维护：`Sync-QuickToolbarForm.ps1:19-39`、生成的 `.frm`、`QuickToolbar.Tests.ps1:72-114`；布局算法有 3 份实现。

### P5. 测试体系名不副实，覆盖极不均衡
4 个 `*.Tests.ps1` 均非 Pester（自制 assert + throw，本机 Pester 仅 3.4.0），无统一 runner，CI 完全不跑测试。18 个 .bas 中 14 个零覆盖——包括最大的 mRibbon 和纯逻辑可离线测的 `Mod_UpdateChecker`（版本比较/JSON 解析）。

### P6. 版本与发布链路有断点
- CI（`release-dotm.yml`）只压包已提交的 `dotm/`，从不校验 git tag 与 `vbaProject.bin` 内 APP_VERSION 一致。
- `README.md:9` 版本号纯手工维护。
- **README 仓库链接指向错误 owner**：`README.md:46` 指向 `Fon509/RATools/releases`，而实际仓库是 `PharmaRA/RATools-for-Word`（`Mod_UpdateChecker.bas:5-9` 是对的）——用户点下载链接会 404。
- `.gitignore:3` 的 `/docs` 会静默忽略一切新文档（`docs/local-dotm-build-testing.md` 是先跟踪后加规则才幸存；`docs/release-workflow.md` 至今未被跟踪）。**本路线图文档提交时需 `git add -f` 或先修 .gitignore。**

### P7. 编码是最脆弱的横切面
源码 GBK 存盘 + `Sync-QuickToolbarForm.ps1:388` 硬编码 CP936 + `Set-RAToolsAppVersion` 用 UTF-8 无 BOM 写回（恰好 `Mod_UpdateChecker.bas` 纯 ASCII 才没出事）+ `.gitattributes` 仅 `* text=auto`。项目内已出现 `FromCodePoints` 规避手段且复制了 3 份（`Mod_QuickToolbarActions.bas:78-84`、`Mod_UpdateChecker.bas:210-215`、`frmAbout.frm:72-77`）。

## 三、分阶段路线图

每阶段末尾的【验收】= 上文"验收三件套"，特殊要求另注。

### 阶段 0：安全网与基线（无行为变更）
1. 新增 `tests/Run-AllTests.ps1`：顺序执行 4 个测试脚本 + 汇总 PASS/FAIL；参数 `-SkipCom` 只跑无 Word 依赖的测试。
2. 新增版本一致性校验脚本（CHANGELOG 顶部 ↔ `APP_VERSION` ↔ `README.md:9`），并入 Run-AllTests。
3. CI 增加轻量测试 job：在 windows-latest 跑 `BuildRATools.Tests.ps1` + 版本一致性（均无 COM 依赖）；发布 job 增加 tag ↔ CHANGELOG ↔ APP_VERSION 源码级校验。
4. 决策并修复 `.gitignore` 的 `/docs` 陷阱（建议：改为仅忽略 `/docs/superpowers/` 与明确的私有文档，让 docs 默认可跟踪）。
5. 记录基线：本地构建 + 全部测试绿 + 成品手测清单留档。

### 阶段 1：已知 bug 修复（小步行为修正，先于重构以免被重构掩盖）
每项独立提交：
1. `Mod_BatchAutoFitTablesToWindow.bas:23` 无表格分支 `Exit Sub` 前未恢复 `ScreenUpdating`。
2. `mRibbon.ImportStyles`：`Documents.Open` 失败时卡在 `ScreenUpdating=False` + 等待光标（`mRibbon.bas:80-84` 无错误保护）；`Mod_ConvertHeadingNumbers`、`Mod_HyperlinksToBlue` 同类问题。
3. `Mod_BatchAcceptAndClean.bas:50-68`：打开失败仍计入成功数，完成提示虚报。
4. `Mod_BatchSetMargins.ProcessFolder`：不递归 + `*.doc*` 误匹配模板文件 → 统一为"递归 + doc/docx/docm 白名单 + 排除 ~$"（与其余批处理一致）。
5. `README.md:46` 仓库链接 `Fon509/RATools` → `PharmaRA/RATools-for-Word`（`README.md:195` 一并核对）。

### 阶段 2：卫生清理（行为不变）
1. 7 个缺失模块补 `Option Explicit`（BatchRenameFiles、BatchAutoFitTablesToWindow、ConvertHeadingNumbers、ExtractAbbreviations、HyperlinksToBlue、LinkToThePreviousSection、RenameCurrentDocument），随手修正由此暴露的未声明变量。
2. 删除死代码 `Mod_BatchDetectHighlights.bas:295-300`（`CheckForHighlight`）。
3. `ProcessFile` 冲突：两处均 Private 化并改名（`ProcessMarginsFile` / `ProcessCleanFile`）。
4. 局部变量 `count` 遮蔽成员名的 4 处改名；批处理计数器 `Integer` → `Long`。
5. 统一错误标签命名（定 `CleanFail`/`ErrH` 其一）与 MsgBox 标题规范（每模块常量标题），仅做机械统一，不改逻辑。

### 阶段 3：公共基础层提取（行为不变，逐模块迁移）
新增 4 个公共模块（命名 `Mod_Core_*`，与业务 `Mod_*` 区分）：
- `Mod_Core_Session`：`BeginBatchUI`/`EndBatchUI`（ScreenUpdating/DisplayAlerts/StatusBar 守卫，错误路径保证恢复）、`SafeOpenDocument(path, readOnly, visible)`/`CloseDocumentQuietly(doc, save)`（吸收 `Mod_BatchSetMargins.bas:88-98` 的"文件已打开检测"这一最佳实现）。
- `Mod_Core_Files`：FSO 获取、`CollectWordFiles(folder, recursive)`（~$ 排除 + doc/docx/docm 白名单），替换 4+1 份递归实现。
- `Mod_Core_UI`：三模式选择对话框（含 `StrPtr` 取消检测这一最佳实现）、Word 文件/文件夹 FileDialog 封装（统一过滤串）、进度 StatusBar 格式、MsgBox 包装。
- `Mod_Core_Text`：`FromCodePoints` 唯一实现（frmAbout 因窗体隔离可保留私有拷贝，注释指向单源）。

迁移顺序从简到繁、每模块一个提交：`BatchSetMargins` → `BatchAutoFitTablesToWindow` → `BatchAcceptAndClean` → `BatchDetectHighlights` → `BatchRenameFiles` → `BatchConvertWordToPDF`（顺带把 `Mod_BatchConvertWordToPDF.bas:11-17` 的 7 个模块级状态收敛为过程内上下文传参）。

### 阶段 4：样式名单一事实源
1. 新增 `Mod_StyleNames`：`STYLE_SUFFIX` 常量、中英映射表（数据数组形式，替代 60 项 Select Case）、`NumberedHeadingStyle(n)` 等构造函数、公共 `StyleExists`。
2. mRibbon（`GetTargetStyleName`）、`Mod_QuickToolbar`、`Mod_QuickToolbarActions`、`Mod_RemoveUnusedStyles` 全部改调 `Mod_StyleNames`。
3. 补离线测试：映射表完整性（中文名 ↔ 英文名与 `template/master-template-en.dotx` 实际样式名一致——用 COM 解包 dotx 校验一次并固化为静态断言）。

### 阶段 5：拆分 mRibbon（依赖阶段 4）
目标：mRibbon 只留 Ribbon 生命周期与回调转发（预计 <150 行）。
1. `Mod_StyleEngine`：`ImportStyles`、`GetStyleFilePath`、`ApplyRAToolsStyle`、`HandleStyleErr`。
2. `Mod_DocCommands`：`SetTextBlue`、`TogglePageBreakBefore`、`AutoFitTableWindow`、`OpenDocumentFolder`、`ProtectFieldFormat`、`Align*_Click`、`ShowStylePane`、`btnCap_Click`。
3. `Mod_MacroRegistry`：`GetMyMacroRegistry` 数据表（frmMacroList 的 `Application.Run` 字符串调用改为直接调用）。
4. 公共变量 `mRibbon` 更名为 `gRibbonUI`（消除与模块同名），`clsAppEvents` 的反向调用同步收窄。
5. **约束**：所有 onAction 回调 Sub 名与签名不变，`customUI14.xml` 零改动（git diff 验证）。
【验收】额外要求：成品 dotm 中逐 group 点击全部 Ribbon 按钮手测一遍。

### 阶段 6：悬浮窗统一分发 + 按钮定义单源
1. 分发合并：`RunQuickToolbarAction` 归一为单个分发表（吸收 `TryRunExpandedQuickToolbarAction`），样式名走 `Mod_StyleNames`；测试桩标注隔离（`' === TEST SUPPORT ===` 区块）。
2. 按钮定义单源：19 按钮表提为 `scripts/QuickToolbarButtons.psd1`，`Sync-QuickToolbarForm.ps1` 与 `QuickToolbar.Tests.ps1` 共同消费；布局算法收敛为 Sync 脚本单份（测试从 psd1 推导期望值而非重算三遍）。
3. 需要重跑 `Sync-QuickToolbarForm.ps1` 再构建（.frx 会重生成，属预期 churn）。

### 阶段 7：构建/测试基础设施收尾
1. COM 样板提取：`RATools.Build.psm1` 新增 `Start-RAToolsWordSession`/`Stop-RAToolsWordSession`（含 AutomationSecurity、释放与双 GC），Build/Sync/3 个 COM 测试全部改用（消除 5 份拷贝）。
2. 测试补盲区（优先纯逻辑）：`Mod_UpdateChecker` 的 `CompareVersions` 与 JSON 提取（可完全离线）、`Mod_BatchRenameFiles` 的文件名清洗规则（逻辑提纯后离线测）、批处理宏各 1 个 COM smoke（借鉴 `BatchDetectHighlights.Tests.ps1` 的 Document.Variables 回传模式）。
3. pre-commit hook 与 `Get-RAToolsProjectLayout` 目录清单联动（消除硬编码重复）。

### 阶段 8（可选，需用户决策）：编码策略统一
现状 GBK 是 Word/VBE 在中文系统的现实约束，不建议激进改动。三个选项：
- **A（推荐，保守）**：文档化"GBK 契约"，`.gitattributes` 为 `*.bas/*.cls/*.frm` 显式声明 `text eol=crlf`，防 `text=auto` 归一化风险；`Set-RAToolsAppVersion` 改为按原编码写回。
- B：git `working-tree-encoding=GBK`（仓库内转 UTF-8 存储，diff 可读，但改写全部 blob，风险中）。
- C：全面码点化中文字面量（可读性差，仅作字符串资源模块的兜底手段）。

## 四、风险与约束备忘

| 风险 | 缓解 |
|---|---|
| Ribbon 回调改名/丢失 → 按钮静默失效 | 阶段 5 约束 XML 零改动；新增静态测试：解析 customUI14.xml 的全部 onAction，与源码 Public Sub 清单比对 |
| GBK 字面量在编辑中损坏 → 样式匹配失败 | 编辑器统一 GBK 打开；阶段 4 后样式名收敛到少数文件；阶段 8A 的 .gitattributes 防线 |
| `MACROBUTTON` 域按名引用 `JumpToHighlightFromReport`（`Mod_BatchDetectHighlights.bas:183`） | 该 Sub 名列入"禁改名清单"（roadmap 附录维护） |
| `Application.Run` 字符串调用（frmMacroList、Sync 脚本） | 阶段 5.3 消除 frmMacroList 处；Sync 脚本处保留并注释 |
| 每次构建 `vbaProject.bin` 必然 churn | 已知成本，提交粒度按阶段控制；不在中间提交里反复重建 |
| Word COM 仅本地可用 | CI 只跑无 COM 测试；COM 测试与成品手测留在本地验收清单 |

## 五、执行顺序与里程碑

```
阶段0 ──► 阶段1 ──► 阶段2 ──► 阶段3 ──► 阶段4 ──► 阶段5 ──► 阶段6 ──► 阶段7 ──► (阶段8)
安全网    bug修复   卫生清理   公共层     样式单源   拆mRibbon  悬浮窗     基建收尾    编码决策
 [完成]    [完成]    [完成]     [完成]     [完成]     [完成]     [完成]     [完成]      待决策
```

- 阶段 0-2 合计工作量小，可合并为一个 PR/版本（如 v0.7.2 维护版）。
- 阶段 3-6 是重构主体，建议每阶段一个版本号递进（v0.8.x）。
- 阶段 5 完成时是最有感知的里程碑：mRibbon 从 684 行瘦身到 <150 行。

### 执行记录（2026-07-26，分支 worktree-refactor-v0.8）

阶段 0-7 已全部完成，共 17 个提交。执行中的重要发现与偏差：

- **新增安全网超出计划**：SourceCompile.Tests（全工程编译探针）、
  RibbonCallbacks.Tests（XML 回调 ↔ Public Sub 静态比对）、
  StyleNameMapping.Tests（映射表 ↔ 英文模板实名校验）三个计划外测试，
  合计使测试套件从 4 个扩到 10 个。
- **StyleNameMapping 测试当场抓出两个存量 bug**：`THeading Left-F` 拼写错误
  （"标题左对齐"的英文回退从未生效）与 `指导-F -> Instruction-F` 死映射
  （英文模板无此样式），已修复/移除。
- **VBA 续行数上限**：单条语句约 24 个续行，46 项映射的单条 Array 字面量
  会被 VBE 拒绝导入（错误 0x800A9D00"输入超出文件结尾"），已改为
  Collection 逐条追加。同样报错也会由残留 WINWORD 进程引起，
  Stop-RAToolsWordSession 已统一释放逻辑。
- **GBK 编辑约定**：通用编辑工具会把 GBK 源码重写为 UTF-8 导致中文损坏，
  新增 scripts/dev/Edit-GbkSource.ps1 作为安全替换工具；所有源码改动
  统一经 936 代码页读写。
- **阶段 7 补充**：批量重命名的清洗规则提纯为 CleanFileBaseName 纯函数
  （10 用例）；UpdateChecker 版本比较/JSON 提取 18 用例。
- 阶段 8（编码策略）保持待决策状态，倾向方案 A（GBK 契约 + .gitattributes）。

## 附录：禁改名清单（外部契约）

- `customUI14.xml` 引用的全部 onAction 回调（`Onload`、`btnStyle_Click`、`btnChar_Click`、`AttachTemplate`、`SetTextBlue`、`TogglePageBreakBefore`、`ShowStylePane`、`AutoFitTableWindow`、`ShowMacroListWindow`、`ToggleQuickToolbar`、`OpenDocumentFolder`、`ProtectFieldFormat`、`AlignLeft_Click`、`AlignCenter_Click`、`AlignRight_Click`、`AlignJustify_Click`、`btnCap_Click`、`Wrapper_*`）
- `GetMyMacroRegistry` 注册的宏名（frmMacroList 经 `Application.Run` 按名调用）
- `JumpToHighlightFromReport`（生成文档内 MACROBUTTON 域引用）
- `GetAppVersion` / `Private Const APP_VERSION`（构建脚本正则 `RATools.Build.psm1:315` 依赖其精确形态）
- 测试桩接口：`GetQuickToolbarButtonCount`、`ReleaseQuickToolbarForTest`、`SetQuickToolbarMsoTestMode`、`GetLastQuickToolbarMsoForTest`
