Attribute VB_Name = "mRibbon"
Option Explicit

'=====================  模 块 级 变 量  =====================
Public mRibbon        As IRibbonUI     ' 缓存 Ribbon (改为 Public 以便调试)
Private mAppEvents    As clsAppEvents  ' 事件监听对象

'=====================  配 置 区 域  =====================
Private Const FILE_PREFIX As String = "RAtools"
Private Const FILE_NAME_CN As String = "master-template-cn.dotx"
Private Const FILE_NAME_EN As String = "master-template-en.dotx"
Private Const TARGET_SUFFIX As String = "-F"

'=====================  Ribbon 必 要 回 调  =====================
' Ribbon 加载完成时触发
Public Sub Onload(ribbon As IRibbonUI)
    Set mRibbon = ribbon
    InitEvents
    ActivateRAToolsTab
End Sub

'=====================  事 件 初 始 化  =====================
Public Sub InitEvents()
    If mAppEvents Is Nothing Then
        Set mAppEvents = New clsAppEvents
    End If
End Sub

' 供外部调用的激活方法 (使用极速异步模式)
Public Sub ActivateRAToolsTab()
    ' 使用 Now 可以在当前所有操作完成后立即执行
    Application.OnTime Now, "DoActivateTab"
    
    ' 确保监听器在线
    If mAppEvents Is Nothing Then InitEvents
End Sub

' 真正的执行过程
Public Sub DoActivateTab()
    On Error Resume Next
    If Not mRibbon Is Nothing Then
        mRibbon.ActivateTab "RATool"
    End If
End Sub

' AutoExec 宏
Public Sub AutoExec()
    InitEvents
End Sub

'=====================  导 入 样 式  =====================
' 说明：双重导入策略，避免后续段落样式丢失，支持 -F 后缀及目录样式

' Ribbon 按钮调用的入口（带成功提示）
Public Sub AttachTemplate(ByVal control As IRibbonControl)
    ' 显式调用，显示成功消息
    ImportStyles isSilent:=False
End Sub

Public Function ImportStyles(Optional isSilent As Boolean = False) As Boolean
    Dim tmplPath As String
    Dim sourceDoc As Document
    Dim currentDoc As Document
    Dim sty As Style
    Dim stylesList As New Collection
    Dim vStyleName As Variant
    Dim pass As Integer
    Dim sName As String
    
    ' 1. 获取路径
    tmplPath = GetStyleFilePath
    If tmplPath = "" Then
        ImportStyles = False
        Exit Function
    End If
    
    Set currentDoc = ActiveDocument
    
    ' 性能优化：关闭屏幕更新，鼠标设为等待状态
    Application.ScreenUpdating = False
    System.Cursor = wdCursorWait
    
    ' 2. 后台打开模版 (只读/不可见)
    Set sourceDoc = Documents.Open(fileName:=tmplPath, ReadOnly:=True, Visible:=False)
    
    On Error Resume Next
    ' 3. 快速筛选：建立待导入名单
    For Each sty In sourceDoc.Styles
        sName = sty.NameLocal
        ' 仅匹配 -F 结尾 或 TOC/图表目录 相关
        If (UCase(Right(sName, Len(TARGET_SUFFIX))) = UCase(TARGET_SUFFIX)) Or _
           (UCase(Left(sName, 3)) = "TOC") Or _
           (InStr(sName, "图表目录") > 0) Or _
           (InStr(sName, "Table of Figures") > 0) Then
            stylesList.Add sName
        End If
    Next sty
    
    ' 获取完名单立即关闭模版，释放内存
    sourceDoc.Close SaveChanges:=wdDoNotSaveChanges
    Set sourceDoc = Nothing
    
    If stylesList.count = 0 Then
        System.Cursor = wdCursorNormal ' 恢复鼠标
        Application.ScreenUpdating = True
        If Not isSilent Then MsgBox "模版中没有找到符合条件（-F 或 TOC）的样式。", vbExclamation
        ImportStyles = False
        Exit Function
    End If
       
    ' 4. 执行导入 (保留双重导入以修复 BasedOn 链接)
    ' 虽然双重导入会多花一点时间，但为了样式层级关系的正确性，这步不能省。
    For pass = 1 To 2
        For Each vStyleName In stylesList
            On Error Resume Next
            Application.OrganizerCopy Source:=tmplPath, Destination:=currentDoc.FullName, _
                Name:=vStyleName, Object:=wdOrganizerObjectStyles
            On Error GoTo 0
        Next vStyleName
    Next pass
    
    ' 恢复状态
    System.Cursor = wdCursorNormal
    Application.ScreenUpdating = True
    
    ' 仅在非静默模式下弹窗
    If Not isSilent Then
        MsgBox "操作完成！已成功导入 " & stylesList.count & " 个样式。", vbInformation, "导入成功"
    End If
    
    ImportStyles = True
End Function

'=====================  智 能 样 式 映 射  =====================
' 作用：将 UI 传来的中文标签（如“标题1-F”）转换为文档中实际存在的样式名
Private Function GetTargetStyleName(ByVal uiTagName As String) As String
    Dim doc As Document
    Set doc = ActiveDocument
    
    ' 1. 优先检查：如果文档中直接存在该中文样式，直接返回
    ' 这样保证了中文模板加载时，速度最快且完全兼容
    If StyleExists(doc, uiTagName) Then
        GetTargetStyleName = uiTagName
        Exit Function
    End If
    
    ' 2. 映射检查：如果中文找不到，尝试查找对应的英文名称
    ' 【注意】这里列出了常见的英文样式名，请确保与 master-template-en.dotx 中的实际名称一致
    Dim mapName As String
    mapName = ""
    
    Select Case uiTagName
        ' --- 基础样式 ---
        Case "正文-F":       mapName = "Body Text with Indentation-F"
        Case "正文无缩进-F": mapName = "Body Text-F"
        Case "正文无间距-F": mapName = "Body Text no Space-F"
        Case "标题居中-F":   mapName = "Heading Center-F"
        Case "标题左对齐-F": mapName = "THeading Left-F"
        Case "目录标题-F":   mapName = "TOC Heading-F"
        
        ' --- 标题类 ---
        Case "标题1-F": mapName = "Heading 1-F"
        Case "标题2-F": mapName = "Heading 2-F"
        Case "标题3-F": mapName = "Heading 3-F"
        Case "标题4-F": mapName = "Heading 4-F"
        Case "标题5-F": mapName = "Heading 5-F"
        Case "标题6-F": mapName = "Heading 6-F"
        Case "标题7-F": mapName = "Heading 7-F"
        Case "标题8-F": mapName = "Heading 8-F"
        Case "标题9-F": mapName = "Heading 9-F"
        Case "无编号标题1-F": mapName = "UN Heading 1-F"
        Case "无编号标题2-F": mapName = "UN Heading 2-F"
        Case "无编号标题3-F": mapName = "UN Heading 3-F"
        Case "无编号标题4-F": mapName = "UN Heading 4-F"
        Case "无编号标题5-F": mapName = "UN Heading 5-F"
        Case "无编号标题6-F": mapName = "UN Heading 6-F"
        Case "无编号标题7-F": mapName = "UN Heading 7-F"
        Case "无编号标题8-F": mapName = "UN Heading 8-F"
        Case "无编号标题9-F": mapName = "UN Heading 9-F"
        Case "附录标题-F": mapName = "Appendix Title-F"
        
        ' --- 表格类 ---
        Case "表头左对齐-F": mapName = "Table Heading Left-F"
        Case "表头居中-F":   mapName = "Table Heading Center-F"
        Case "表头右对齐-F": mapName = "Table Heading Right-F"
        Case "表格文本左对齐-F": mapName = "Table Cell Left-F"
        Case "表格文本居中-F":   mapName = "Table Cell Center-F"
        Case "表格文本右对齐-F": mapName = "Table Cell Right-F"
        Case "表格文本无间距-F": mapName = "Table Cell no Space-F"
        Case "表格编号列表-F": mapName = "Table List Number-F"
        Case "表格项目符号列表-F": mapName = "Table List Bullet-F"
        Case "表格注释-F":   mapName = "Table Note-F"
        Case "表标题-F":     mapName = "Table Title-F"
        
        ' --- 图片类 ---
        Case "图片-F":       mapName = "Figure-F"
        Case "图标题-F":     mapName = "Figure Title-F"
        
        ' --- 列表类 ---
        Case "编号列表-F":     mapName = "List Number-F"
        Case "项目符号列表-F": mapName = "List Bullet-F"
        Case "参考文献列表-F": mapName = "List Reference-F"
        
        ' --- 其他 ---
        Case "页眉-F": mapName = "Header-F"
        Case "页脚-F": mapName = "Footer-F"
        Case "脚注-F": mapName = "Footnote-F"
        Case "超链接-F": mapName = "Hyperlink-F"
        Case "指导-F": mapName = "Instruction-F"
        
            
    End Select
    
    ' 3. 如果找到了映射名，检查文档里是否存在这个英文样式
    If mapName <> "" Then
        If StyleExists(doc, mapName) Then
            GetTargetStyleName = mapName
            Exit Function
        End If
    End If
    
    ' 4. 如果都没找到，还是返回原始标签（让它在 ApplyStyle 里正常报错）
    GetTargetStyleName = uiTagName
End Function

' 辅助函数：检查样式是否存在
Private Function StyleExists(doc As Document, sName As String) As Boolean
    On Error Resume Next
    Dim s As Style
    Set s = doc.Styles(sName)
    StyleExists = (Err.Number = 0)
    On Error GoTo 0
End Function

'=====================  应 用 样 式  =====================
Private Sub ApplyStyle(ByVal uiTagName As String)
    Dim realStyleName As String
    
    ' 获取实际样式名（自动处理中英文映射）
    realStyleName = GetTargetStyleName(uiTagName)
    
    ' 尝试应用样式
    On Error GoTo ErrH
    Selection.Style = ActiveDocument.Styles(realStyleName)
    Exit Sub
    
ErrH:
    HandleStyleErr
End Sub

'=====================  段 落 样 式  =====================
Public Sub btnStyle_Click(ByVal control As IRibbonControl)
    On Error GoTo ErrH
    ApplyStyle control.Tag
    Exit Sub
ErrH:
    HandleStyleErr
End Sub

'=====================  字 符 样 式  =====================
Public Sub btnChar_Click(ByVal control As IRibbonControl)
    On Error GoTo ErrH
    
    Dim uiTagName As String
    Dim realStyleName As String
    Dim defaultStyleName As String
    
    uiTagName = control.Tag
    
    ' 1. 获取目标样式的真实名称 (比如 "Heading 1-F")
    realStyleName = GetTargetStyleName(uiTagName)
    
    ' 2. 检查当前是否已经是该样式 (用于重复点击取消)
    If Selection.Style = realStyleName Then
        ' 也要动态获取“正文-F”的实际名称 (可能是 "Body Text-F")
        defaultStyleName = GetTargetStyleName("正文-F")
        realStyleName = defaultStyleName
    End If
    
    ' 3. 应用样式
    Selection.Style = ActiveDocument.Styles(realStyleName)
    Exit Sub
    
ErrH:
    HandleStyleErr
End Sub

' 确保加载逻辑支持版本号匹配
Private Function EnsureMainTemplate() As Boolean
    Dim mMainTemplate As Template
    If mMainTemplate Is Nothing Then
        Dim t As Template
        For Each t In Templates
            ' 模糊匹配：只要文件名以 RAtools 开头，且是 .dotm 结尾即可
            If UCase(t.Name) Like UCase(FILE_PREFIX) & "*.DOTM" Then
                Set mMainTemplate = t
                Exit For
            End If
        Next
    End If
    EnsureMainTemplate = Not mMainTemplate Is Nothing
    If Not EnsureMainTemplate Then _
        MsgBox "未检测到以 " & FILE_PREFIX & " 开头的加载项！", vbCritical
End Function

'样式错误统一提示
Private Sub HandleStyleErr()
    ' 5941: 集合成员不存在(样式不存在); 91: 对象变量未设置
    If Err.Number = 5941 Or Err.Number = 91 Then
        Dim ans As VbMsgBoxResult
        ans = MsgBox("当前文档未包含目标样式，请先加载主模板 dotx！" & vbCrLf & vbCrLf & _
                     "点击【确定】立即加载模板。", vbOKCancel + vbExclamation, "提示")
        
        If ans = vbOK Then
            Err.Clear
            ' 传入 Nothing，直接触发加载模板的主逻辑
            ' 此时会进入 GetStyleFilePath，如果找到模板则询问，找不到则进入 PickFile
            AttachTemplate Nothing
        End If
    Else
        MsgBox "样式应用失败：" & Err.Description, vbCritical
    End If
End Sub

' 获取路径函数
' 逻辑：
' 1. 获取 RAtools.dotm 所在目录
' 2. 检测该目录下是否存在 -cn.dotx 和 -en.dotx
' 3. 根据存在情况决定直接加载或弹窗选择
Private Function GetStyleFilePath() As String
    Dim basePath As String
    Dim pathCN As String, pathEN As String
    Dim existCN As Boolean, existEN As Boolean
    Dim t As Template
    
    ' 1. 获取 RAtools.dotm (本工具) 的所在路径
    ' 尝试通过 ThisDocument 获取（如果代码就在该模板里）
    On Error Resume Next
    basePath = ThisDocument.Path
    If Err.Number <> 0 Or basePath = "" Then
        ' 如果失败（比如代码被导入其他地方），则遍历模板集合寻找 RAtools*.dotm
        Err.Clear
        For Each t In Templates
            ' 模糊匹配：支持 RAtools_v1.0.dotm 等变体
            If UCase(t.Name) Like UCase(FILE_PREFIX) & "*.DOTM" Then
                basePath = t.Path
                Exit For
            End If
        Next
    End If
    On Error GoTo 0
    
    ' 如果仍找不到路径（极少见情况），跳到手动选择
    If basePath = "" Then GoTo ManualSelect
    
    ' 2. 构建目标文件路径
    pathCN = basePath & Application.PathSeparator & FILE_NAME_CN
    pathEN = basePath & Application.PathSeparator & FILE_NAME_EN
    
    existCN = (Dir(pathCN) <> "")
    existEN = (Dir(pathEN) <> "")
    
    ' 3. 判断逻辑
    If existCN And existEN Then
        ' 都有：弹窗选择
        Dim ans As VbMsgBoxResult
        ans = MsgBox("在工具目录下检测到中英文两种模板：" & vbCrLf & vbCrLf & _
                     "【是 (Yes)】 加载中文模板 (-cn)" & vbCrLf & _
                     "【否 (No)】  加载英文模板 (-en)" & vbCrLf & vbCrLf & _
                     "点击【取消】手动选择其他文件。", _
                     vbYesNoCancel + vbQuestion, "选择样式模板")
                     
        If ans = vbYes Then
            GetStyleFilePath = pathCN
            Exit Function
        ElseIf ans = vbNo Then
            GetStyleFilePath = pathEN
            Exit Function
        Else
            ' 用户点击取消，直接跳转到 PickFile 标签，
            ' 跳过 ManualSelect 的 "未找到" 提示
            GoTo PickFile
        End If
        
    ElseIf existCN Then
        ' 只有中文
        GetStyleFilePath = pathCN
        Exit Function
        
    ElseIf existEN Then
        ' 只有英文
        GetStyleFilePath = pathEN
        Exit Function
    End If
    
    ' 如果都不存在，进入下面的手动选择逻辑

ManualSelect:
    ' 只有当文件真的不存在，或者路径为空时，才应该显示这个提示
    If MsgBox("在工具同级目录下未找到默认样式模板：" & vbCrLf & _
              FILE_NAME_CN & vbCrLf & _
              FILE_NAME_EN & vbCrLf & vbCrLf & _
              "是否手动选择文件？", vbYesNo + vbQuestion) = vbNo Then Exit Function

PickFile: ' 直接进入文件选择
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Word 模板", "*.dot;*.dotx;*.dotm"
        If .Show = -1 Then GetStyleFilePath = .SelectedItems(1)
    End With
End Function

'=====================  一 键 大 写  =====================
Public Sub btnCap_Click(ByVal control As IRibbonControl)
    On Error Resume Next
    Selection.Range.Case = wdUpperCase
End Sub

'================  设置文字为蓝色  ================
Public Sub SetTextBlue(control As IRibbonControl)
    On Error Resume Next
    Selection.Font.Color = wdColorBlue
End Sub

'================  域格式保护  ================
'为选区内 REF/PAGEREF 加 \* MERGEFORMAT,保护域格式
'包含智能判断(全文/选区) + 结果弹窗
Public Sub ProtectFieldFormat(control As IRibbonControl)
    Dim fld As field, rng As Range
    Dim targetFields As Fields ' 目标域集合
    Dim msgTip As String       ' 提示信息

    ' 判断：如果是光标插入点(wdSelectionIP)则处理全文，否则处理选区
    If Selection.Type = wdSelectionIP Then
        Set targetFields = ActiveDocument.Fields
        msgTip = "未选中文字，已对【全文】域代码进行格式保护。"
    Else
        Set targetFields = Selection.Fields
        msgTip = "已对【选中区域】域代码进行格式保护。"
    End If

    ' 遍历处理
    For Each fld In targetFields
        If fld.Type = wdFieldRef Or fld.Type = wdFieldPageRef Then
            Set rng = fld.Code
            If InStr(1, rng.Text, "mergeformat", vbTextCompare) = 0 Then
                rng.Text = rng.Text & " \* MERGEFORMAT "
                fld.Update
            End If
        End If
    Next fld

    ' 操作完成后弹出提示
    MsgBox msgTip, vbInformation, "操作完成"
End Sub

'================  打开文件所在文件夹  ================
Public Sub OpenDocumentFolder(control As IRibbonControl)
        
    ' 1. 声明一个变量，用来临时存储文档的路径
    Dim docPath As String
    
    ' 2. 获取当前活动文档的路径
    docPath = ActiveDocument.Path
    
    ' 3. 检查文档是否已经保存过（如果没有保存过，路径会是空的）
    If docPath = "" Then
        ' 弹出一个消息框提醒用户
        MsgBox "这个文档还没有保存过哦！请先保存文档，然后再尝试打开所在文件夹。", vbExclamation, "温馨提示"
        ' 停止运行后面的代码
        Exit Sub
    End If
    
    ' 4. 如果路径存在，使用 Windows 资源管理器打开该路径
    ' explorer.exe 是 Windows 自带的文件夹管理程序
    Shell "explorer.exe """ & docPath & """", vbNormalFocus
    
End Sub

'================  段前分页切换  ================
' 说明：切换选中段落的“段前分页”属性 (PageBreakBefore)
' 逻辑：如果是混合状态或关闭状态 -> 设为开启；如果是纯开启状态 -> 设为关闭
Public Sub TogglePageBreakBefore(control As IRibbonControl)
    On Error Resume Next
    Dim currentStatus As Long
    
    ' 获取当前选中段落的段前分页状态
    ' 0 = False (关), -1 = True (开), 9999999 = wdUndefined (混合)
    currentStatus = Selection.ParagraphFormat.PageBreakBefore
    
    ' 如果全关(0)，则开启(-1)
    ' 如果混合(wdUndefined)，也统一开启(-1)
    ' 如果全开(-1)，则关闭(0)
    If currentStatus = -1 Then
        Selection.ParagraphFormat.PageBreakBefore = 0 ' 关闭
    Else
        Selection.ParagraphFormat.PageBreakBefore = -1 ' 开启
    End If
End Sub

'================  表格功能：根据窗口自动调整  ================
Public Sub AutoFitTableWindow(control As IRibbonControl)
    On Error Resume Next
    
    ' 检查光标是否在表格内
    If Selection.Information(wdWithInTable) Then
        ' 将当前所在的表格设置为：根据窗口自动调整
        Selection.Tables(1).AutoFitBehavior wdAutoFitWindow
    Else
        MsgBox "请先将光标定位在表格内部。", vbExclamation, "提示"
    End If
End Sub

'================  下拉选择对齐方式  ================
'================  下拉菜单：左对齐  ================
Public Sub AlignLeft_Click(control As IRibbonControl)
    Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
End Sub

'================  顶部大按钮：直接设为居中  ================
Public Sub AlignCenter_Click(control As IRibbonControl)
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
End Sub

'================  下拉菜单：右对齐  ================
Public Sub AlignRight_Click(control As IRibbonControl)
    Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
End Sub

'================  下拉菜单：两端对齐  ================
Public Sub AlignJustify_Click(control As IRibbonControl)
    Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
End Sub


' 显示/隐藏样式管理窗格
Public Sub ShowStylePane(control As IRibbonControl)
    On Error GoTo ErrorHandler
    
    ' 尝试使用内置命令打开样式窗格
    Application.CommandBars.ExecuteMso "StylesPane"
    
    Exit Sub
    
ErrorHandler:
    ' 如果内置命令失败，使用快捷键
    SendKeys "%^{+}s", True
End Sub

'=====================  宏 列 表 管 理  =====================

' 1. Ribbon 回调：点击按钮弹出窗体
' 在 Ribbon XML 中，将按钮的 onAction 指向这个 Sub
Public Sub ShowMacroListWindow(control As IRibbonControl)
    frmMacroList.Show
End Sub


' 2. 供窗体调用的数据源函数
' 返回值：Variant 数组
Public Function GetMyMacroRegistry() As Variant
    Dim items As New Collection
    Dim vArr() As Variant
    Dim i As Long
    
    ' ================= 配置区域 =================
    
    ' 格式：items.Add Array("宏代码名", "列表显示的名称", "下方显示的详细介绍")
    
    ' 第1个
    items.Add Array("SetHyperlinksAndFieldsToBlue", _
                    "超链接和域批量设置为蓝色", _
                    "智能遍历文档，将所有超链接和域（REF/PAGEREF等）的颜色设置为蓝色，但在处理过程中会自动排除图表题注和页码。")
      
    ' 第2个：功能调整为通过按钮实现
    ' items.Add Array("Wrapper_ProtectFieldFormat", _
    '                 "域格式保护", _
    '                 "扫描全文或选区内的引用域，自动添加 \* MERGEFORMAT 开关，防止更新域后格式丢失。")
                    
    ' 第3个
    items.Add Array("BatchConvertWordToPDF", _
                    "Word批量转PDF", _
                    "批量将单个或多个Word转为PDF，并通过Word标题创建PDF书签。")
    
    ' 第4个
    items.Add Array("BatchRenameFiles", _
                    "批量修改文件名", _
                    "批量修改文件名" & vbCrLf & _
                    "1. 仅保留汉字、小写字母、数字、中划线和下划线" & vbCrLf & _
                    "2. 汉字与字符间的空格（以及其他剩余空格）直接删除，字母数字间的空格改为中划线 ""-""，其他非法字符替换为中划线 ""-""" & vbCrLf & _
                    "3. 支持“文件夹模式”和“多文件选择模式”" & vbCrLf & _
                    "4. 如果文件被占用无法重命名，自动创建改名后的副本")
    
    ' 第5个
    items.Add Array("ConvertHeadingNumbers", _
                    "标题自动编号转文本", _
                    "将文档中所有标题（大纲 1-9 级）的自动编号转换为固定的静态文本。")
    
    ' 第6个
    items.Add Array("RenameCurrentDocument", _
                    "重命名当前文件", _
                    "无需关闭文件，直接重命名当前文件。")
    
    ' 第7个
    items.Add Array("BatchSetMargins", _
                    "批量设置页边距", _
                    "批量将单个或多个文档页面上、下、左、右的页边距设置为 2.54厘米（即标准的 1 英寸）。")
    
    ' 第8个
    items.Add Array("BatchAutoFitTablesToWindow", _
                    "一键表格自动调整", _
                    "将文档中所有表格批量设置为“根据窗口自动调整”。")
                    
    ' 第9个
    items.Add Array("BatchAcceptAndClean", _
                    "批量接受修订并删除批注", _
                    "批量将单个或多个文档的tracking版转换为clean版，接受所有修订并停止修订同时删除文档中的所有批注。")
                    
    ' 第10个
    items.Add Array("LinkToThePreviousSection", _
                    "页眉和页脚设置为“链接到前一节”", _
                    "遍历文档中除第一节以外的所有节，将所有页眉和页脚设置为“链接到前一节”。")
                    
    ' 第11个
    items.Add Array("RemoveUnusedStyles", _
                    "清理未使用的模板样式", _
                    "一键清理文档中所有未被使用的自定义样式（仅针对以 -F 结尾的样式），保持文档整洁。")
                    
    ' 第12个
    items.Add Array("BatchDetectHighlights", _
                    "批量检测高亮内容", _
                    "检测文档是否有突出显示颜色的内容，在最终clean前进行调整。")
                    
    ' 第13个
    items.Add Array("ExtractAbbreviations", _
                    "提取缩略语", _
                    "利用了Word内置的通配符功能提取全大写英文缩略语。")
                    
    ' 如果以后要加新宏，直接复制粘贴即可，无需修改其他地方
    ' 如果需要control参数的宏，需要下面做一个Wrapper，见下面Wrapper包装器下的内容，同时需要在上面添加
    
    ' ================= 配置结束 =================
    
    If items.count > 0 Then
        ReDim vArr(0 To items.count - 1)
        For i = 1 To items.count
            vArr(i - 1) = items(i)
        Next i
        GetMyMacroRegistry = vArr
    Else
        GetMyMacroRegistry = Empty
    End If
End Function

'=====================  Wrapper 包装器  =====================
' 解释：因为很多宏是 Ribbon 回调 (带 control 参数)，通常是通过按钮直接调用功能
' Application.Run 无法自动提供 control 参数，直接运行会报错。
' 所以我们需要一些不带参数的“外壳”过程。

Public Sub Wrapper_ProtectFieldFormat()
    ' 调用原有的逻辑
    ' 注意：因为原 Sub 需要 control 参数，我们传 Nothing 进去
    ' 只要原 Sub 内部没用到 control.ID 或 control.Tag，这样写就是安全的
    ProtectFieldFormat Nothing
End Sub

Public Sub Wrapper_RenameCurrentDocument(control As IRibbonControl)
    RenameCurrentDocument
End Sub

Public Sub Wrapper_ShowAbout(control As IRibbonControl)
    frmAbout.Show
End Sub

