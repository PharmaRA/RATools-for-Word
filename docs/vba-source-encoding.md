# VBA 源码编码契约（GBK/936）

本项目 `modules/*.bas`、`class_modules/*.cls`、`userforms/*.frm` 统一以
**GBK（代码页 936）+ CRLF** 存盘。这是刻意选择而非历史遗留，修改这些文件前请先读完本文。

## 为什么是 GBK

Word/VBE 在中文 Windows 上按系统 ANSI 代码页（936）导入导出 VBA 源码：

- 导入 UTF-8 源文件时 VBE 按 936 解码，中文注释与字符串字面量全部乱码
  （v0.7.0 悬浮工具栏中文乱码事故的根因，见 CHANGELOG）。
- VBE 导出的 .frm/.bas 也是 936 编码，`Sync-QuickToolbarForm.ps1` 依赖该约定回读。

因此源码在磁盘上必须保持 GBK，直到 Microsoft 让 VBE 支持 UTF-8（不要指望）。

## 契约内容

1. **磁盘编码**：`.bas/.cls/.frm` 一律 GBK(936)，CRLF 行尾。
2. **.gitattributes**：三种扩展名显式声明 `text eol=crlf`，防止 `* text=auto`
   把某些 GBK 字节序列误判为二进制或做出意外归一化；`.frx/.dotm/.dotx/.bin/.png`
   显式 `binary`。
3. **脚本读写**：任何 PowerShell/构建脚本读写 VBA 源码必须显式使用
   `[System.Text.Encoding]::GetEncoding(936)`。已有实现：
   - `scripts/RATools.Build.psm1` 的 `Set-RAToolsAppVersion`（按 GBK 读写回）
   - `scripts/Sync-QuickToolbarForm.ps1`（按 936 回读 VBE 导出）
   - `tests/*.Tests.ps1` 的 `Read-VbaSource` 等辅助函数
4. **通用编辑工具禁区**：常见编辑器/AI 编辑工具默认按 UTF-8 写文件，
   会把 GBK 中文整体损坏。对含中文的源码做机械替换时使用
   `scripts/dev/Edit-GbkSource.ps1`（按 936 读→精确替换→按 936 写，
   带出现次数断言防误替换）。
5. **纯 ASCII 修改不受限**：如果改动只涉及 ASCII 字符且工具保持原字节
   不动，无需特殊处理；但"整文件重写"式的工具仍会破坏文件中已有的中文。

## 验证手段

- `tests/BuildRATools.Tests.ps1` 的 "Set-RAToolsAppVersion preserves GBK
  Chinese content"：断言版本改写后中文完好且磁盘仍为 GBK 字节。
- `tests/SourceCompile.Tests.ps1`：全部源码经 Word COM 导入编译，
  编码损坏通常会在此暴露（字面量样式名失配、语法错误）。
- 快速人工检查：`file modules/*.bas` 应显示 `ISO-8859`（GBK 被误标的常见值）
  或含中文时非 UTF-8；`git diff` 中中文显示为乱码是正常现象
  （终端按 UTF-8 解码 GBK），只要 Word 内显示正确即可。

## 例外清单

- `Mod_NormalizeScientificTerms.bas`、部分纯 ASCII 模块：无中文内容，
  GBK 与 ASCII 兼容，无特殊处理。
- `frmAbout.frm` 保留一份私有 `FromCodePoints`（码点构造中文），
  这是窗体导入隔离下的历史规避手段；新代码一律直接写 GBK 中文字面量，
  公共版本在 `Mod_Core_Text`。
