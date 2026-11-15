# Clean 检测最终验收记录

日期：2026-11-17。接续 clean检测02 的两项补测和最终只读复审。

## 实际执行与结果

在仓库根目录、本机 Windows Word COM 环境执行：

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File tests/BatchDetectCleanVersions.Tests.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File tests/Run-AllTests.ps1
```

- 新功能：25 个场景、607 项断言通过，退出码 0。证据：`final-new-feature.log`。
- 完整套件：31 PASS、0 FAIL、0 SKIP，退出码 0；包含最终源码编译。证据：`final-run-all-tests.log`。
- 本次补测通过真实二进制 `.doc` 的复合文件签名校验，并验证 Clean 与带修订的非 Clean 文件、不保存原文件、哈希不变。
- 真实 `.docm` 包含 VBA 工程及 `Document_Open`。检测时独立标记文件不存在；允许宏的阳性对照写入恰好一行标记并产生修订。检测前后文件哈希相同。
- 补测复现了同一 Word 实例强制禁宏打开含 VBA 文件时中断调用的问题。最小复现确认专用隐藏实例可以完整返回；修复限定为尚未打开的 `.doc/.docm`，已打开源继续复用当前对象和未保存内容。
- 专用实例打开后实测只读状态、宏安全级别；源文档窗口在收尾后关闭。窗口关闭断言不能单独证明进程退出，实例 Quit 的成功/失败收尾路径已静态审查。

## 最终复审

子代理对五个交付文件及直接依赖 `Core_Session.bas`、`Core_Files.bas` 完成只读复审，最终未发现阻断问题。范围包含三条件判定、已打开源保留、专用实例打开/关闭核验、错误降级、宏隔离和报告。

组合图形修订未单独造样实测；已有正文、脚注、尾注、页眉页脚、文本框和画布修订场景。完整套件仍报告既有格式修订处理器从 Word 内启动的权限限制 `UNAVAILABLE=拒绝的权限`，其外部处理器和 Word 回退用例通过；该限制不计作 clean 宏隔离通过证据。

原对话日志 `new-feature-18.log`、`run-all-tests.log`、`source-compile.log` 仅保留为历史记录；最终结果以 `final-*` 通过日志及下列源码哈希为准。本次诊断脚本、失败日志和测试 fixture 已移入回收站，模块和测试改动前备份保留于 `backups/clean-detection-final-20261117-073832/`。

## 最终交付文件 SHA256

| 文件 | SHA256 |
| --- | --- |
| modules/Macro_BatchDetectCleanVersions.bas | 10B71FC9262F7D8F5BEEFB16AF56FAF82A6CBE3F90736F2E3B70416584143774 |
| modules/UI_MacroRegistry.bas | F8215F8DFBCDB68BD4981DBA396163AEEE7F4B77B953958AFDE7F88AD8767D0B |
| tests/BatchDetectCleanVersions.Tests.ps1 | BBB127FF4E318A7AF0A140CD2D82BC43D611CDA21310ACEA524D762698F8974A |
| tests/Run-AllTests.ps1 | 691D037A4722E365FC8F5D90076E8F2A38425F9E434ACC24C71A5B3F73A914ED |
| docs/manual/clean-documents.md | 1DFCC934D0D2D925E440D013E48E9F22F99AA0B2E859201229ED14DF5DC67586 |
