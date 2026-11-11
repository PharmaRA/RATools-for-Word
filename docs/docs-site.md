# 文档站构建、检查与发布

用户正文位于 `docs/manual/`，站点是独立的 Astro Starlight 工程，源目录为 `website/`。同一份 Markdown 可在仓库阅读，也可构建为带中文搜索的静态网站。开发文档保留在 `docs/`，不作为用户手册正文发布。

## 本地安装

站点需要 Node.js 22.23.4 或更高版本（当前锁定的 Astro 7.4.10 要求；CI 使用 Node 24）。从仓库根目录运行：

```powershell
cd website
npm ci
```

若当前网络需要本机 HTTP 代理，可为 npm 设置 `http://127.0.0.1:7890`。代理配置以本机情况为准，不把账户信息写入仓库。

## 同步、检查、构建与预览

```powershell
.\.venv-docs\Scripts\python.exe scripts/docs/check_docs.py
cd website
npm run build
cd ..
.\.venv-docs\Scripts\python.exe scripts/docs/check_docs.py --site-dir website/dist --base-path /RATools-for-Word
```

`npm run build` 会先自动同步 `docs/manual/`，再生成静态站点。预览构建产物：

```powershell
cd website
npm run preview -- --host 127.0.0.1 --port 4321
```

浏览器访问 `http://127.0.0.1:4321/RATools-for-Word/`。搜索使用 Pagefind，必须在 HTTP 预览或部署后的网站上使用，直接双击 HTML 的 `file://` 方式无法加载搜索索引。

`website/src/content/docs/manual/` 和 `website/public/manual/assets/` 是从 `docs/manual/` 生成的镜像，不应手工修改，也已在版本库中忽略。侧栏定义位于 `website/src/data/sidebar.json`，同步脚本会核对它是否覆盖全部手册页。同步脚本负责：

- 补上 Starlight 需要的标题、摘要和编辑链接；
- 把 `.md` 相对链接改写为站内路由；
- 转义代码块语言；
- 保留旧版标题锚点；
- 复制截图资源。

编辑正文时只改 `docs/manual/`，然后运行 `npm run build` 或站点开发命令；两者都会自动同步。

### 检查内容

- 从实际 Ribbon XML 提取所有有效动作和 Word 原生入口，排除 XML 注释中的按钮。
- 从 GBK 编码的宏注册表提取有效注册项，从悬浮窗事件提取动作键。
- 与 `docs/feature-map.json` 逐项比对，验证目标页面、锚点及源入口名称；新增、删除或改接回调时要求同步维护映射。
- 检查手册、README 和 docs 索引中的本地链接和图片。
- 对生成的 HTML 检查站内链接、锚点和静态资源，并确认 Pagefind 搜索索引存在。

这些检查证明入口覆盖和链接有效，不能代替对功能说明准确性的人工审阅。行为变化时应对照对应 VBA 模块更新说明，尤其是处理范围、保存、备份及输出规则。

## GitHub Actions 与 Pages

`.github/workflows/docs.yml` 对相关 PR 和 main 分支变更执行同步、检查、构建和链接校验，并上传 `documentation-site` 工件供审阅。**这些自动构建不会直接发布网站。**

正式发布步骤：

1. 合并经审阅的文档到 main。
2. 在仓库 **Settings → Pages → Build and deployment → Source** 选择 **GitHub Actions**。
3. 在 **Actions → Documentation → Run workflow** 中选择 **main** 并运行。
4. 等待 build 和 deploy 两个任务完成，通过 deploy 的环境链接访问网站。
5. 检查安装页、一个复杂功能页和中文搜索，确认后再对外分发站点链接。

本配置的目标地址是 `https://PharmaRA.github.io/RATools-for-Word/`。这只是部署目标；创建配置文件并不证明网站已经上线。以后每次需要更新线上手册，都在 main 上手动运行同一工作流。

发布使用 GitHub Pages 工件，不向 `gh-pages` 分支强推，不修改插件发布工作流。相关机制见 [GitHub Pages 官方工作流说明](https://docs.github.com/en/pages/getting-started-with-github-pages/using-custom-workflows-with-github-pages)。

## 如何维护内容

### 设计说明

页面只使用本地字体，避免访问外部字体服务。站点从仓库现有 Logo 提取黑金配色，首页按任务入口组织，内页保持三栏阅读结构。深色与浅色主题共用同一套可读性变量；窄屏下侧栏收为菜单，表格允许横向滚动。

### 内容更新规则

- 简单按钮按用途组合在一个页面，每个按钮保留具体操作说明；复杂宏独立成页。
- 一个功能至少说明入口、适用场景、操作步骤、处理范围、结果、保存和限制。不要从函数名或宏列表短描述推断行为。
- 截图复用自根目录 `_image/` 的既有公开教程素材，站点所需副本放在 `docs/manual/assets/`；源图更新时同步相关副本，旧版界面示意需标注。
- 根 README 保持简短，详细教程只在手册维护；开发构建细节放开发文档。
- 当前手册按 v0.8.8 标识的源码编写，不承诺早期发布包已包含后来修复；版本变化时同步首页说明。