# Release 发布流程

本文档用于发布 RATools for Word 新版本。当前推荐流程是：本地从源码构建并同步 `dotm/`，提交后由 GitHub Actions 从 tag 对应的 `dotm/` 目录打包 Release 资产。

## 发布来源原则

正式 Release 中的 `.dotm` 文件应由 GitHub Actions 从仓库中的 `dotm/` 目录生成，不直接上传本地 `dist/RATools_local.dotm`。

推荐链路：

```text
modules/ class_modules/ userforms/
  -> scripts\Build-RAToolsDotm.ps1
  -> dotm/
  -> git tag vX.Y.Z
  -> GitHub Actions
  -> Release 资产
```

本地 `dist/` 只用于自测和临时检查；`dotm/` 才是提交进仓库、供远程发布工作流打包的内容。

## 1. 确认版本号

确定本次要发布的版本号，例如：

```text
v0.6.4
```

版本号需要同时用于：

- `CHANGELOG.md` 顶部标题。
- `modules/Engine_UpdateChecker.bas` 中的 `APP_VERSION`，构建脚本会自动同步。
- Git tag。
- GitHub Release 名称和资产名。

`CHANGELOG.md` 的标题必须写成：

```markdown
# v0.6.4
```

发布 workflow 会按这个标题提取 Release notes。如果 tag 是 `v0.6.4`，但 `CHANGELOG.md` 中没有 `# v0.6.4`，远程打包会失败。

## 2. 更新 CHANGELOG

在 `CHANGELOG.md` 顶部新增版本段落，格式参考历史版本：

```markdown
# v0.6.4

- 新增……`✨ 新增` (Added)
- 优化……`🛠 优化` (Changed)
- 修复……`🐛 修复` (Fixed)
```

建议只写用户能理解的变化，不写过多内部实现细节。

## 3. 关闭 Word 并确认信任设置

关闭所有 Word 窗口，尤其不要打开 `RATools*.dotm`。

如果本机还没有启用 VBA 项目访问权限，打开 Word，进入：

```text
文件 -> 选项 -> 信任中心 -> 信任中心设置 -> 宏设置
```

勾选：

```text
信任对 VBA 项目对象模型的访问
```

## 4. 运行本地测试

在仓库根目录运行：

```powershell
powershell -ExecutionPolicy Bypass -File tests\BuildRATools.Tests.ps1
powershell -ExecutionPolicy Bypass -File tests\NormalizeScientificTerms.Tests.ps1
```

预期结果：

```text
PASS BuildRATools.Tests
PASS NormalizeScientificTerms
```

如果新增了其他测试，也应在发布前一起运行。

## 5. 正式构建并同步 dotm/

运行：

```powershell
powershell -ExecutionPolicy Bypass -File scripts\Build-RAToolsDotm.ps1
```

默认情况下，脚本会读取 `CHANGELOG.md` 顶部版本作为 `APP_VERSION`。如果需要临时覆盖，可以显式传入：

```powershell
powershell -ExecutionPolicy Bypass -File scripts\Build-RAToolsDotm.ps1 -AppVersion v0.6.4
```

这一步会：

- 按 `CHANGELOG.md` 顶部版本同步 `modules\Engine_UpdateChecker.bas` 中的 `APP_VERSION`。
- 从 `modules/`、`class_modules/`、`userforms/` 导入 VBA 源码。
- 清理作者、最后修改者、公司等文档元数据。
- 生成 `dist\RATools_local.dotm`。
- 将生成后的 `.dotm` 解包并同步回 `dotm/`。

运行后检查工作区：

```powershell
git status --short
```

如果 `dotm/word/vbaProject.bin` 等文件变化，这是预期结果，需要随本次发布提交。

如果 `modules/Engine_UpdateChecker.bas` 变化，通常是 `APP_VERSION` 被同步到了当前发布版本，也需要随本次发布提交。

## 6. 手动检查本地 dotm

打开：

```text
dist\RATools_local.dotm
```

检查：

- Word 没有报错。
- RATools 选项卡能正常出现。
- 宏列表窗口能正常打开。
- 本次新增或修改的功能可以正常执行。
- “文件 -> 信息”中的作者、最后修改者、公司等基本信息不含个人信息。
- 关于对话框或宏内显示的当前版本号与本次发布版本一致。

检查完成后关闭 Word。

## 7. 提交发布相关变更

确认差异：

```powershell
git status --short
git diff --stat
```

暂存需要发布的文件，例如：

```powershell
git add CHANGELOG.md modules class_modules userforms dotm tests README.md docs scripts .gitignore
```

创建中文提交：

```powershell
git commit -m "v0.6.4：发布说明"
```

提交信息可以按实际内容调整，例如：

```powershell
git commit -m "v0.6.5：优化批量处理功能"
```

## 8. 创建 tag

确认当前提交就是要发布的提交：

```powershell
git log --oneline -3
```

创建 tag：

```powershell
git tag v0.6.4
```

确认 tag 指向正确提交：

```powershell
git show --stat v0.6.4
```

## 9. 推送 main 和 tag

先推送主分支：

```powershell
git push origin main
```

再推送 tag，触发 GitHub Actions 发布 workflow：

```powershell
git push origin v0.6.4
```

也可以一次性推送：

```powershell
git push origin main --tags
```

如果只是正常新增提交和新增 tag，不需要 force push。

## 10. 等待 GitHub Actions 完成

推送 tag 后，GitHub Actions 会自动运行 `Release DOTM` workflow。

远程 workflow 会：

- 检出 tag 对应的仓库内容。
- 从 `CHANGELOG.md` 中提取对应版本的 Release notes。
- 将 `dotm/` 压缩为 `RATools_vX.Y.Z.dotm`。
- 上传 `template/master-template-cn.dotx`。
- 上传 `template/master-template-en.dotx`。

如果 workflow 失败，优先检查：

- tag 名是否与 `CHANGELOG.md` 标题完全一致。
- `CHANGELOG.md` 对应版本段落是否为空。
- `dotm/` 是否存在且已提交。
- 两个 `template/*.dotx` 是否存在。

## 11. 发布后检查 Release

在 GitHub Release 页面检查资产是否包含：

```text
RATools_v0.6.4.dotm
master-template-cn.dotx
master-template-en.dotx
```

下载 Release 中的 `.dotm`，再做一次快速检查：

- Word 能正常加载。
- RATools 选项卡能出现。
- 本次更新的关键功能可用。
- “文件 -> 信息”中没有个人信息。

## 12. 重新触发已有 tag 的发布

如果 tag 已经存在，但需要重新跑发布 workflow，可以在 GitHub Actions 页面手动运行 `Release DOTM`，输入已有 tag，例如：

```text
v0.6.4
```

手动触发时仍然会从该 tag 对应的 `dotm/` 目录打包，不会使用本地 `dist/` 文件。

## 发布前最小清单

- `CHANGELOG.md` 有对应 `# vX.Y.Z` 段落。
- 本地测试通过。
- `scripts\Build-RAToolsDotm.ps1` 已正式运行。
- `dotm/` 的变化已提交。
- 本地 `dist\RATools_local.dotm` 手动检查通过。
- tag 指向正确提交。
- GitHub Actions 发布成功。
- Release 资产下载后检查通过。
