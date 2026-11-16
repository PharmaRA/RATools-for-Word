# 审阅悬浮窗实施与验收

依据会话 219 已确认方案，从 `57db4c9` 建立工具窗开发工作区。当前分支为 `feat/workflow-toolbars`，工作目录为 `C:\Repositories\RATools-for-Word\.worktrees\workflow-toolbars`。本次不提交、不合并。

该工作区用于审阅、撰写、定稿等阶段工具窗的持续开发；本文记录已完成的审阅悬浮窗，撰写与定稿工具窗仍属后续开发范围。

2026-11-17 按管理方式调整，已通过 `git worktree move` 迁入主仓库的 `.worktrees/` 并重命名分支，保留全部未提交修改、测试证据和插件产物。主仓库的本地 `.git/info/exclude` 已加入 `/.worktrees/`，无需修改项目 `.gitignore`。旧测试日志中的绝对路径是迁移前的历史执行位置，保留原始记录；下文相对路径均以当前 worktree 为根。

范围：独立非模态窗口，七项现有功能；日常审阅、格式处理、定稿前检查、清稿处理四区。功能区保留新建批注、修订开关，用审阅窗入口替代格式处理及审阅者筛选按钮，原宏列表入口保留。

实现：`UI_ReviewToolbar` 管理已加载实例、动作与忙碌状态；`frmReviewToolbar` 显示当前文档和修订状态、保存独立位置。应用事件刷新而不隐式加载窗体，执行前重新获取活动文档；重开文档后重新刷新。批注和筛选保留交互焦点，其余动作交还文档焦点。`Sync-ReviewToolbarForm.ps1` 生成七个文字按钮的 FRX；使用现有回收站清理函数。

验证范围：Word COM 导入与编译、七项实际动作分发、双窗共存与 owner、文档切换和关闭、无文档、忙碌/异常恢复、位置恢复；独立 dotm 打包及源码匹配。破坏性业务验证仅使用测试文档。交付前对子模块及直接依赖进行只读 review。具体结果在执行后补入。

## 2026-11-17 实际验证

构建和下列检查均在 Windows 本机运行，退出码均为 0。未执行完整 `Run-AllTests`，以下结果仅代表选定范围。

| 检查 | 结果与范围 |
| --- | --- |
| `Sync-ReviewToolbarForm.ps1` | Word COM 生成真实 `.frm/.frx`，七个文字按钮、四个分区；临时目录移至回收站 |
| `ReviewToolbar.Tests.ps1` | 52 条集成断言、6 条无文档断言；验证分发、忙碌/重入、异常恢复、双窗、owner、置顶、切文档、关闭重开、位置、隐藏/卸载；六项带交互的业务入口在内存替换为记录器 |
| `ReviewToolbarActions.Tests.ps1` | 13 条真实动作断言：原生批注新增及原选区锚点、筛选面板打开/复用、清稿接受修订/删除批注/关闭跟踪/保留待保存状态；仅模式选择、完成提示与错误呈现于内存适配 |
| `BatchDetectHighlights.Tests.ps1` | 当前实际高亮检测逻辑、脚注及错误路径通过 |
| `BatchDetectCleanVersions.Tests.ps1` | 25 个场景、607 条断言，含只读文件检测、已打开未保存文档、DOC/DOCM、宏禁用与分类报告 |
| `AcceptFormatRevisions.Tests.ps1` | 处理包完整性、内容修订保留、处理后真实 Word 重开、API 回退、取消保存通过 |
| `ReviewerFilter.Tests.ps1` | 更新后的入口契约、多作者收集和可见性控制通过 |
| `QuickToolbarWindow.Tests.ps1` | 74 条既有窗口层断言通过 |
| `SourceCompile.Tests.ps1` | 所有标准模块、类模块及窗体导入编译通过 |
| `BuiltPackage.Tests.ps1` | 成品可由 Word 打开；33 个标准模块源码匹配，41 个源码组件齐全 |
| `ReviewToolbarPackage.Tests.ps1` | 包内 UserForm 类型、非模态、七个控件、焦点属性、窗体/事件类源码和 Ribbon 入口迁移通过 |

附加定点检查：新 VBA 与改动的事件类保持 GBK/CRLF；新脚本通过语法/编码检查；七项审阅动作的手册路径/锚点与 Ribbon 映射一致，`git diff --check` 通过。旧筛选测试已同步入口迁移，两个新 COM suite 已注册到 `Run-AllTests.ps1`。

执行命令（在本 worktree 根目录）：

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File scripts/Sync-ReviewToolbarForm.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File tests/ReviewToolbar.Tests.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File tests/ReviewToolbarActions.Tests.ps1
# 上表其余源码检查使用相同入口：tests/<名称>.Tests.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File scripts/Build-RAToolsDotm.ps1 -AppVersion v0.8.8 -NoSyncDotmDirectory -OutputPath dist/RATools_v0.8.8_review-toolbar_20261117.dotm
powershell.exe -NoProfile -ExecutionPolicy Bypass -File tests/BuiltPackage.Tests.ps1 -PackagePath dist/RATools_v0.8.8_review-toolbar_20261117.dotm
powershell.exe -NoProfile -ExecutionPolicy Bypass -File tests/ReviewToolbarPackage.Tests.ps1 -PackagePath dist/RATools_v0.8.8_review-toolbar_20261117.dotm
```

产物：`dist/RATools_v0.8.8_review-toolbar_20261117.dotm`，367,532 字节；SHA256：`B96086D4D0E2ADCE23B834C2F152E13634B636F639C2BCDA9B01AC7F4F34B9D6`。

证据日志：`backups/review-toolbar-20261117/verification/`；改动前备份同在 `backups/review-toolbar-20261117/`。中断测试留下的四个夹具目录已移入回收站，可恢复；其余测试夹具按脚本自动回收。原 main 工作区保持干净，未提交、合并或安装插件到 Word 启动目录。

只读 review 发现的提示气泡接入缺口已修复，复用 `EnsureTooltipOnTop` 与鼠标移动事件，无定时器；测试错误呈现的循环风险也已修复。自动化覆盖不能替代实际屏幕下的视觉与焦点验收。

### 仍需人工确认的交互边界

在 Word 内加载成品后，检查窗口位置/文字排版和提示气泡、使用键盘输入批注、筛选窗已有实例再次打开时的交互焦点，以及 Word 界面修订开关切换后返回审阅窗的状态显示。本次没有完成 32/64 位 Office、多显示器/DPI 组合矩阵。

格式修订测试记录到本机 Office 限制：`Word could not launch processor: UNAVAILABLE=拒绝的权限`。外部执行生成的处理程序及 Word API 回退均通过，不能据此声称 Word 内启动外部处理程序已通过。
