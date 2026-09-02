Attribute VB_Name = "Mod_StyleEngine"
Option Explicit

' =============================================
' 样式引擎：模板样式导入与样式应用
' 从 mRibbon 拆出（v0.8 重构阶段5）。
' AttachTemplate/AttachTemplateCN/AttachTemplateEN/AttachTemplateCustom
' 以及 btnStyle_Click/btnChar_Click 是 Ribbon onAction 回调，
' 名称与签名不可改（customUI14.xml 按名引用）。
' =============================================

Private Const FILE_PREFIX As String = "RAtools"
Private Const FILE_NAME_CN As String = "master-template-cn.dotx"
Private Const FILE_NAME_EN As String = "master-template-en.dotx"


'=====================  导 入 样 式  =====================
' 说明：双重导入策略，避免后续段落样式丢失，支持 -F 后缀及目录样式

' Ribbon 按钮调用入口：默认加载（大图标）
Public Sub AttachTemplate(ByVal control As IRibbonControl)
    ImportStyles isSilent:=False
End Sub

' Ribbon 按钮调用入口：加载中文模板
Public Sub AttachTemplateCN(ByVal control As IRibbonControl)
    Dim tmplPath As String
    tmplPath = GetCNTemplatePath()
    If tmplPath <> "" Then
        ImportStyles isSilent:=False, specificPath:=tmplPath
    End If
End Sub

' Ribbon 按钮调用入口：加载英文模板
Public Sub AttachTemplateEN(ByVal control As IRibbonControl)
    Dim tmplPath As String
    tmplPath = GetENTemplatePath()
    If tmplPath <> "" Then
        ImportStyles isSilent:=False, specificPath:=tmplPath
    End If
End Sub

' Ribbon 按钮调用入口：浏览其他模板
Public Sub AttachTemplateCustom(ByVal control As IRibbonControl)
    Dim tmplPath As String
    tmplPath = PickTemplateFile()
    If tmplPath <> "" Then
        ImportStyles isSilent:=False, specificPath:=tmplPath
    End If
End Sub


Public Function ImportStyles(Optional isSilent As Boolean = False, Optional ByVal specificPath As String = "") As Boolean
    Dim tmplPath As String
    Dim sourceDoc As Document
    Dim currentDoc As Document
    Dim sty As Style
    Dim stylesList As New Collection
    Dim vStyleName As Variant
    Dim pass As Integer
    Dim sName As String

    ' 1. 获取路径（若指定具体路径则直接使用，否则执行常规探测）
    If specificPath <> "" Then
        tmplPath = specificPath
    Else
        tmplPath = GetStyleFilePath()
    End If

    If tmplPath = "" Then
        ImportStyles = False
        Exit Function
    End If

    Set currentDoc = ActiveDocument

    ' 性能优化：关闭屏幕更新，鼠标设为等待状态
    Application.ScreenUpdating = False
    System.Cursor = wdCursorWait

    ' 2. 后台打开模版 (只读/不可见)
    On Error GoTo OpenFail
    Set sourceDoc = Documents.Open(fileName:=tmplPath, ReadOnly:=True, Visible:=False)

    On Error Resume Next
    ' 3. 快速筛选：建立待导入名单
    For Each sty In sourceDoc.Styles
        sName = sty.NameLocal
        ' 仅匹配 -F 结尾 或 TOC/图表目录 相关
        If (UCase(Right(sName, Len(STYLE_SUFFIX))) = UCase(STYLE_SUFFIX)) Or _
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
    Exit Function

OpenFail:
    ' 打开模板失败（被占用/损坏/路径失效）时必须恢复屏幕刷新与光标
    System.Cursor = wdCursorNormal
    Application.ScreenUpdating = True
    If Not isSilent Then MsgBox "无法打开样式模板：" & Err.Description, vbCritical
    ImportStyles = False
End Function


'=====================  智 能 样 式 映 射  =====================
' 作用：将 UI 传来的中文标签（如“标题1-F”）转换为文档中实际存在的样式名
Private Function GetTargetStyleName(ByVal uiTagName As String) As String
    ' 中英映射与存在性探测统一由 Mod_StyleNames 提供
    GetTargetStyleName = ResolveStyleName(ActiveDocument, uiTagName)
End Function



'=====================  应 用 样 式  =====================
Public Sub ApplyRAToolsStyle(ByVal uiTagName As String)
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
    ApplyRAToolsStyle control.Tag
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


' 样式错误统一提示
Private Sub HandleStyleErr()
    ' 5941: 集合成员不存在(样式不存在); 91: 对象变量未设置
    If Err.Number = 5941 Or Err.Number = 91 Then
        Dim ans As VbMsgBoxResult
        ans = MsgBox("当前文档未包含目标样式，请先加载主模板 dotx！" & vbCrLf & vbCrLf & _
                     "点击【确定】立即加载模板。", vbOKCancel + vbExclamation, "提示")

        If ans = vbOK Then
            Err.Clear
            ' 传入 Nothing，直接触发加载模板的主逻辑
            ' 此时会进入 GetStyleFilePath，如果找到模板则询问，找不到则进入 PickTemplateFile
            AttachTemplate Nothing
        End If
    Else
        MsgBox "样式应用失败：" & Err.Description, vbCritical
    End If
End Sub


' 获取 RAtools.dotm 所在目录
Private Function GetPluginBasePath() As String
    Dim basePath As String
    Dim t As Template

    On Error Resume Next
    basePath = ThisDocument.Path
    If Err.Number <> 0 Or basePath = "" Then
        Err.Clear
        For Each t In Templates
            If UCase(t.Name) Like UCase(FILE_PREFIX) & "*.DOTM" Then
                basePath = t.Path
                Exit For
            End If
        Next
    End If
    On Error GoTo 0
    GetPluginBasePath = basePath
End Function


' 交互式选取模板文件
Private Function PickTemplateFile() As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Word 模板", "*.dot;*.dotx;*.dotm"
        If .Show = -1 Then PickTemplateFile = .SelectedItems(1)
    End With
End Function


' 获取中文样式模板路径
Private Function GetCNTemplatePath() As String
    Dim basePath As String
    Dim p As String

    basePath = GetPluginBasePath()
    If basePath <> "" Then
        p = basePath & Application.PathSeparator & FILE_NAME_CN
        If Dir(p) <> "" Then
            GetCNTemplatePath = p
            Exit Function
        End If
    End If

    If MsgBox("在工具同级目录下未找到中文样式模板：" & vbCrLf & _
              FILE_NAME_CN & vbCrLf & vbCrLf & _
              "是否手动选择文件？", vbYesNo + vbQuestion, "选择中文样式模板") = vbYes Then
        GetCNTemplatePath = PickTemplateFile()
    End If
End Function


' 获取英文样式模板路径
Private Function GetENTemplatePath() As String
    Dim basePath As String
    Dim p As String

    basePath = GetPluginBasePath()
    If basePath <> "" Then
        p = basePath & Application.PathSeparator & FILE_NAME_EN
        If Dir(p) <> "" Then
            GetENTemplatePath = p
            Exit Function
        End If
    End If

    If MsgBox("在工具同级目录下未找到英文样式模板：" & vbCrLf & _
              FILE_NAME_EN & vbCrLf & vbCrLf & _
              "是否手动选择文件？", vbYesNo + vbQuestion, "选择英文样式模板") = vbYes Then
        GetENTemplatePath = PickTemplateFile()
    End If
End Function


' 路径获取主函数（大图标或默认调用）
Private Function GetStyleFilePath() As String
    Dim basePath As String
    Dim pathCN As String, pathEN As String
    Dim existCN As Boolean, existEN As Boolean

    basePath = GetPluginBasePath()
    If basePath = "" Then GoTo ManualSelect

    pathCN = basePath & Application.PathSeparator & FILE_NAME_CN
    pathEN = basePath & Application.PathSeparator & FILE_NAME_EN

    existCN = (Dir(pathCN) <> "")
    existEN = (Dir(pathEN) <> "")

    If existCN And existEN Then
        Dim ans As VbMsgBoxResult
        ans = MsgBox("在工具目录下检测到中英文两种模板：" & vbCrLf & vbCrLf & _
                     "【是 (Yes)】 加载中文模板 (-cn)" & vbCrLf & _
                     "【否 (No)】  加载英文模板 (-en)" & vbCrLf & vbCrLf & _
                     "点击【取消】手动选择其他文件。" & vbCrLf & _
                     "（提示：也可点击功能区“加载模板”下方的下拉箭头直接选择）", _
                     vbYesNoCancel + vbQuestion, "选择样式模板")

        If ans = vbYes Then
            GetStyleFilePath = pathCN
            Exit Function
        ElseIf ans = vbNo Then
            GetStyleFilePath = pathEN
            Exit Function
        Else
            GetStyleFilePath = PickTemplateFile()
            Exit Function
        End If

    ElseIf existCN Then
        GetStyleFilePath = pathCN
        Exit Function

    ElseIf existEN Then
        GetStyleFilePath = pathEN
        Exit Function
    End If

ManualSelect:
    If MsgBox("在工具同级目录下未找到默认样式模板：" & vbCrLf & _
              FILE_NAME_CN & vbCrLf & _
              FILE_NAME_EN & vbCrLf & vbCrLf & _
              "是否手动选择文件？", vbYesNo + vbQuestion, "未找到样式模板") = vbYes Then
        GetStyleFilePath = PickTemplateFile()
    End If
End Function
