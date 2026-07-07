# 本地 DOTM 构建测试流程

本文档用于测试 `scripts/Build-RAToolsDotm.ps1` 本地构建流程是否正常。建议按顺序执行，先验证不改动 `dotm/` 的路径，再执行正式同步。

## 1. 关闭 Word

先关闭所有 Word 窗口，尤其不要打开 `RATools*.dotm`，避免模板文件被占用。

## 2. 启用 Word VBA 项目访问权限

打开 Word，进入：

```text
文件 -> 选项 -> 信任中心 -> 信任中心设置 -> 宏设置
```

勾选：

```text
信任对 VBA 项目对象模型的访问
```

本地构建脚本需要通过 Word COM 自动导入 `.bas`、`.cls`、`.frm` 文件，这个设置是必需的。

## 3. 运行纯脚本测试

在仓库根目录运行：

```powershell
powershell -ExecutionPolicy Bypass -File tests\BuildRATools.Tests.ps1
```

预期最后输出：

```text
PASS BuildRATools.Tests
```

## 4. 测试旧式打包路径

这一步只测试“基于当前 `dotm/` 文件夹直接打包”的旧方案，不导入 VBA 源码，也不同步 `dotm/`。

```powershell
powershell -ExecutionPolicy Bypass -File scripts\Build-RAToolsDotm.ps1 -Version test -SkipVbaImport -NoSyncDotmDirectory
```

预期输出包含：

```text
Built dotm: ...\dist\RATools_test.dotm
```

## 5. 测试新式源码导入路径，但暂不同步 dotm/

这一步会通过 Word COM 从 `modules/`、`class_modules/`、`userforms/` 导入源码并生成测试用 `.dotm`，但不会重写 `dotm/`。

```powershell
powershell -ExecutionPolicy Bypass -File scripts\Build-RAToolsDotm.ps1 -Version comtest -NoSyncDotmDirectory
```

预期输出包含：

```text
Starting Word automation
Removing VBA component: ...
Importing modules\...
Importing class_modules\...
Importing userforms\...
Saving updated dotm
Built dotm: ...\dist\RATools_comtest.dotm
```

如果这里失败，并提示无法访问 VBA project，请重新检查第 2 步的 Word 信任中心设置。

## 6. 手动检查生成的 dotm

打开：

```text
dist\RATools_comtest.dotm
```

检查以下内容：

- Word 没有报错。
- RATools 选项卡能正常出现。
- 宏列表窗口能正常打开。
- 至少运行一个熟悉的功能，确认宏可以执行。
- 在 Word 的“文件 -> 信息”中检查基本信息，确认作者、最后修改者、公司等字段没有个人信息。

## 7. 确认工作区没有被测试产物污染

运行：

```powershell
git status --short
```

`dist/` 是本地构建输出目录，已被 `.gitignore` 忽略，不应作为待提交文件出现。

## 8. 正式同步 dotm/

前面步骤都正常后，再运行正式构建：

```powershell
powershell -ExecutionPolicy Bypass -File scripts\Build-RAToolsDotm.ps1
```

这一步会：

- 从 `modules/`、`class_modules/`、`userforms/` 导入 VBA 源码。
- 生成 `dist\RATools_local.dotm`。
- 将生成后的 `.dotm` 解包并同步回 `dotm/`。

运行后检查变化：

```powershell
git status --short
```

如果看到 `dotm/word/vbaProject.bin` 等文件变化，这是预期结果。

## 9. 可选：安装提交前自动构建 hook

确认手动构建流程稳定后，可以安装本地 `pre-commit` hook：

```powershell
powershell -ExecutionPolicy Bypass -File scripts\Install-RAToolsBuildHook.ps1
```

安装后，提交前会自动运行本地构建并暂存 `dotm/`。

如果需要卸载：

```powershell
powershell -ExecutionPolicy Bypass -File scripts\Install-RAToolsBuildHook.ps1 -Uninstall
```

hook 会要求相关 VBA 源码先暂存，避免提交中的源码和 `dotm/` 不一致。
