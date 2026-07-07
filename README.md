<div align="center"> <img src="_image/logo.png" alt="Logo" width="400"/> <h3>RATools for Word - 专为药品注册（RA）打造的 Word 效率插件</h3> <p>基于实战经验开发，提升申报资料编写效率，助力注册申报工作更上一层楼。</p> </div>



## 📖 项目简介

本插件基于 `.dotm`（启用宏的模板）定义功能核心，并通过 `.dotx` 文件管理样式模板。旨在解决药品RA人在文档编写中频繁切换选项卡、格式调整繁琐等痛点。

**主要特性：**

- **实战导向**：功能源于开发者在注册申报过程中的长期实践，精准解决痛点。
- **兼容性**：已在 Windows 10 / 11系统下的 Microsoft 365 Word环境中测试通过，理论上支持Office 2010之后的至今所有Word版本（请自行测试）。
- **持续迭代**：后续将根据实际需求，持续添加更多便捷功能。

<div align="center"> <img src="_image/0.png" width=100%/> </div>

## 🧱 项目结构与开发说明

本项目结构如下：

- `dotm/`：存放 `.dotm` 插件解包后的主要内容，包括 `vbaProject.bin`、Ribbon XML、自定义图片资源等，是插件的核心载体。
- `modules/`：存放从 `vbaProject.bin` 导出的标准模块代码（`.bas`），便于阅读、搜索和版本对比。
- `class_modules/`：存放从 `vbaProject.bin` 导出的类模块代码（`.cls`），主要用于事件和对象封装。
- `userforms/`：存放从 `vbaProject.bin` 导出的窗体文件（`.frm/.frx`），用于宏管理器等界面功能。
- `template/`：存放可供插件加载的 Word 样式模板（`.dotx`），用于统一文档格式。

为避免源码与实际插件内容不一致，推荐使用“源码优先”的本地构建流程：

- 日常开发优先直接修改 `modules/`、`class_modules/`、`userforms/` 中的源码文件。
- 修改完成后运行 `powershell -ExecutionPolicy Bypass -File scripts\Build-RAToolsDotm.ps1`。
- 构建脚本会以当前 `dotm/` 为模板基底，自动把源码导入 VBA 工程，清理作者、最后修改者、公司等文档元数据，保存新的 `.dotm`，同步回 `dotm/`，并生成 `dist/RATools_local.dotm`。
- 如需在提交前自动同步，可运行 `powershell -ExecutionPolicy Bypass -File scripts\Install-RAToolsBuildHook.ps1` 安装本地 `pre-commit` hook。hook 会要求相关 VBA 源码先暂存，避免提交中的源码和 `dotm/` 不一致；卸载可运行同一脚本并加上 `-Uninstall`。

原有流程仍然保留：如果你仍然在 Word 的 VBA 编辑器中修改主程序，可以保存 `.dotm` 后手动解包回 `dotm/`；远程发布工作流仍会基于 `dotm/` 打包发布。

> 注意：本地源码导入依赖 Word COM 自动化，需要在 Word 信任中心启用“信任对 VBA 项目对象模型的访问”。

## ⚙️ 安装与配置

为了确保插件正常运行，请严格按照以下步骤进行配置。

### 1. 下载与文件准备

- 前往 [Releases](https://github.com/Fon509/RATools/releases) 下载最新的 `.dotm` 和 `.dotx` 文件。
- **创建目录**：建议在 D 盘根目录创建`RATools`文件夹，将 `.dotx` 和 `.dotm` 放入其中

> **注意**：您也可以自定义文件路径，但后续步骤需对应修改路径配置。本文档以推荐路径 `D:\RATools` 为例。

<div align="center"> <img src="_image/folder.png" width=75%/> </div>

### 2. 配置 Word 启动路径

1. 打开 Word，点击左上角 **「文件」** -> **「选项」**。

   <div align="center"> <img src="_image/1.png" width=25%/> <img src="_image/2.png" width=25%/> </div> 

2. 在弹出的对话框中点击 **「高级」**，向下滑动至“常规”栏目，点击 **「文件位置」**。

   <div align="center"> <img src="_image/3.png" width=100%/> </div> 

3. 选中 **「启动」** 项，点击“修改”，选择存放 `RAtools.dotm` 的文件夹路径（例如：`D:\RATools`）。

<div align="center"> <img src="_image/4.png" width=75%/> </div>

### 3. 添加受信任位置

为防止宏被系统安全策略拦截，需将插件目录设为受信任位置：

1. 在 Word 选项对话框左侧，选择 **「信任中心」** -> **「信任中心设置」**。

   <div align="center"> <img src="_image/5.png" width=75%/> </div> 

2. 选择 **「受信任位置」**，点击 **「添加新位置」**。

   <div align="center"> <img src="_image/6.png" width=100%/> </div> 

3. 点击 **「浏览」**，选择插件文件夹路径（`D:\RATools`），确认无误后点击 **「确定」** 保存所有设置。

<div align="center"> <img src="_image/7.png" width=50%/> </div>
<div align="center">  <img src="_image/8.png" width=100%/> </div>

## 🚀 使用指南

### 加载模板

配置成功后，打开或新建Word文件，顶部会出现 **「RATools」** 选项卡（请不要双击 `.dotm` 新建文件）。

1. 点击 **「点击加载」** 按钮即可挂载模板样式（每个文件仅首次需要进行加载）。

   <div align="center"> <img src="_image/9.png" width=100%/> </div> 

2. 在v0.5.0的更新中，加入了中英文模版功能，目前挂载模版逻辑如下：

   - 如果 `RAtools.dotm` 同目录下仅存在 `master-template-cn.dotx` 或 `master-template-en.dotx` 中的一个，默认挂载该模版中的样式。

     <div align="center"> <img src="_image/10_1.png" width=40%/> </div>

   - 如果 `RAtools.dotm` 同目录下同时存在 `master-template-cn.dotx` 和 `master-template-en.dotx` ，会弹出对话框选择需要挂载的模板或自行选择其他模板文件。

     <div align="center"> <img src="_image/10_2.png" width=40%/> </div>

   - 如果 `RAtools.dotm` 同目录下不存在任何模板文件，会弹出对话框进行模版文件选择。

     <div align="center"> <img src="_image/10_3.png" width=40%/> </div>

3. 成功加载后将提示 **「操作完成」**。

### 功能模块详解

当前版本集成了三大核心模块：

#### 1. 样式快速应用

基于 `.dotx` 定义的标准样式模板，提供一键应用预设样式功能，统一文档格式标准。

#### 2. 常用选项

将分散在 Word 不同选项卡中的高频功能（如字体、段落等）聚合至同一面板，减少鼠标点击与页面切换，显著提升操作流。

在功能区单独的「关于」项下，可打开关于对话框，查看当前版本、访问 GitHub/Gitee 仓库，并手动检查更新；发现新版本后，可选择前往 GitHub 或 Gitee 的 Release 详情页进行更新。

#### 3. 增强型宏工具

内置宏列表对话框，比 Word 原生界面更清晰直观，同时在v0.5.1版本中添加了搜索功能。目前已内置以下宏：

| **功能名称**                       | **说明**                                                     |
| ---------------------------------- | ------------------------------------------------------------ |
| **超链接和域批量设置为蓝色**       | 智能遍历文档，将所有超链接和域（REF/PAGEREF等）的颜色设置为蓝色，但在处理过程中会自动排除图表题注和页码。（如有 Bug 欢迎反馈） |
| ~~**域格式保护**~~                 | ~~扫描全文或选区内的引用域，自动添加 \* MERGEFORMAT 开关，防止更新域后格式丢失。~~<br />调整为通过功能区按钮实现 |
| **Word批量转PDF**                  | 批量将单个或多个Word转为PDF，并通过Word标题创建PDF书签。     |
| **批量修改文件名**                 | 批量修改文件名<br/>1. 仅保留汉字、字母、数字、中划线和下划线<br/>2. 字母/数字间的空格改为中划线，汉字与字符间的空格（以及其他剩余空格）直接删除，其他非法字符替换为中划线 "-"<br/>3. 支持“文件夹模式”和“多文件选择模式”<br/>4. 如果文件被占用无法重命名，自动创建改名后的副本 |
| **标题自动编号转文本**             | 将文档中所有标题（大纲 1-9 级）的自动编号转换为固定的静态文本。 |
| **重命名当前文件**                 | 无需关闭文件，直接重命名当前文件。                           |
| **一键设置页边距**                 | 一键将单个或多个文件页面上、下、左、右的页边距设置为 2.54厘米（即标准的 1 英寸）。 |
| **一键表格自动调整**               | 将文档中所有表格批量设置为“根据窗口自动调整”                 |
| **批量接受修订并删除批注**         | 批量将单个或多个文档的tracking版转换为clean版，接受所有修订并停止修订同时删除文档中的所有批注。 |
| **页眉和页脚设置为“链接到前一节”** | 遍历文档中除第一节以外的所有节，将所有页眉和页脚设置为“链接到前一节”。 |
| **清理未使用的模板样式**           | 一键清理文档中所有未被使用的自定义样式（仅针对以 -F 结尾的样式），保持文档整洁。 |
| **批量检测高亮内容**               | 批量检测文档是否有突出显示颜色的内容，以便在最终clean前进行调整。 |
| **提取缩略语**                     | 利用Word内置的通配符功能提取全大写英文缩略语。               |

## ⬆️ 进阶用法

### 1. 修改.dotx文件实现样式自定义

RATools 的样式应用功能依赖于底层的 `.dotx` 模板文件。如果你有独特的样式偏好（如特定的字体、字号或段落间距），可以通过修改模板来实现，从而满足自己独特的样式偏好：

1. 根据你的使用环境，在Word中打开 `master-template-cn.dotx`（中文）或 `master-template-en.dotx`（英文）。

   <div align="center"> <img src="_image/11_1.png" width=40%/> </div>
   <div align="center"> <img src="_image/11_2.png" width=100%/> </div>

2. 打开 `.dotx` 模板文件后，在Word的「样式」面板中，找到并修改你想要调整的带-F后缀的样式（例如 「正文-F」等）。

   <div align="center"> <img src="_image/11_3.png" width=100%/> </div>

3. 保存并关闭模板文件。

之后使用RATools的样式功能时，插件将自动应用你自定义后的格式标准。

### 2. 创建属于自己的宏并添加至宏列表中

RATools 支持扩展。如果你具备 VBA 开发能力，可以将自己的常用脚本集成到工具中：

1. 打开 RATools 的主程序文件（`.dotm`），但是不要直接打开 `D:\RATools`中的`.dotm`文件，可以复制到其他路径打开。

   <div align="center"> <img src="_image/12_1.png" width=100%/> </div>

2. 按 `Alt + F11` 进入 VBA 编辑器。

3. 在工程资源管理器中右键点击「RATools」，选择“插入” -> “模块”，创建一个新的模块。

   <div align="center"> <img src="_image/12_2.png" width=40%/> </div>

4. 编写你的 `Public Sub` 过程（宏代码）。

5. 在「mRibbon」模块中末尾处添加你所创建的宏代码的信息。

   <div align="center"> <img src="_image/12_3.png" width=75%/> </div>

6. 测试无误后左上角保存代码并将保存后的 `.dotm` 文件移动至 `D:\RATools`，删除之前的版本（注意备份）。

你的自定义宏现在可以通过 Word 的宏列表或 RATools 的宏管理功能进行调用，实现功能的个性化扩展。

## 📝 交流与反馈

如果您在使用过程中遇到问题或有新的功能建议，欢迎提交 Issue 或联系开发者。

## 📅 更新日志

查看版本更新历史，请参阅 [CHANGELOG](https://github.com/Fon509/RATools-for-Word/blob/main/CHANGELOG.md)。
