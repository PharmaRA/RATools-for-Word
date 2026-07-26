Attribute VB_Name = "Mod_StyleNames"
Option Explicit

' =============================================
' 样式名单一事实源
' 项目 "-F" 样式体系的全部知识集中于此：
' - 后缀常量与判断
' - 中文 UI 名 -> 英文模板名映射（master-template-en.dotx）
' - 常用样式名构造函数
' - 样式存在性探测
' 映射数据与 template/master-template-en.dotx 的实际样式名一致，
' 修改模板样式名时必须同步本表（有 StyleNameMapping 测试守护）。
' =============================================

Public Const STYLE_SUFFIX As String = "-F"

Private mMapKeys As Variant
Private mMapValues As Variant
Private mMapReady As Boolean

' 中英映射数据表。返回 Collection，元素为 Array(中文名, 英文名)。
' 注意：VBA 单条语句续行上限约 24 个，故逐条追加而非单条 Array 字面量。
Public Function GetStyleNameMappingData() As Collection
    Dim data As New Collection

    data.Add Array("正文-F", "Body Text with Indentation-F")
    data.Add Array("正文无缩进-F", "Body Text-F")
    data.Add Array("正文无间距-F", "Body Text no Space-F")
    data.Add Array("标题居中-F", "Heading Center-F")
    data.Add Array("标题左对齐-F", "Heading Left-F")
    data.Add Array("目录标题-F", "TOC Heading-F")
    data.Add Array("标题1-F", "Heading 1-F")
    data.Add Array("标题2-F", "Heading 2-F")
    data.Add Array("标题3-F", "Heading 3-F")
    data.Add Array("标题4-F", "Heading 4-F")
    data.Add Array("标题5-F", "Heading 5-F")
    data.Add Array("标题6-F", "Heading 6-F")
    data.Add Array("标题7-F", "Heading 7-F")
    data.Add Array("标题8-F", "Heading 8-F")
    data.Add Array("标题9-F", "Heading 9-F")
    data.Add Array("无编号标题1-F", "UN Heading 1-F")
    data.Add Array("无编号标题2-F", "UN Heading 2-F")
    data.Add Array("无编号标题3-F", "UN Heading 3-F")
    data.Add Array("无编号标题4-F", "UN Heading 4-F")
    data.Add Array("无编号标题5-F", "UN Heading 5-F")
    data.Add Array("无编号标题6-F", "UN Heading 6-F")
    data.Add Array("无编号标题7-F", "UN Heading 7-F")
    data.Add Array("无编号标题8-F", "UN Heading 8-F")
    data.Add Array("无编号标题9-F", "UN Heading 9-F")
    data.Add Array("附录标题-F", "Appendix Title-F")
    data.Add Array("表头左对齐-F", "Table Heading Left-F")
    data.Add Array("表头居中-F", "Table Heading Center-F")
    data.Add Array("表头右对齐-F", "Table Heading Right-F")
    data.Add Array("表格文本左对齐-F", "Table Cell Left-F")
    data.Add Array("表格文本居中-F", "Table Cell Center-F")
    data.Add Array("表格文本右对齐-F", "Table Cell Right-F")
    data.Add Array("表格文本无间距-F", "Table Cell no Space-F")
    data.Add Array("表格编号列表-F", "Table List Number-F")
    data.Add Array("表格项目符号列表-F", "Table List Bullet-F")
    data.Add Array("表格注释-F", "Table Note-F")
    data.Add Array("表标题-F", "Table Title-F")
    data.Add Array("图片-F", "Figure-F")
    data.Add Array("图标题-F", "Figure Title-F")
    data.Add Array("编号列表-F", "List Number-F")
    data.Add Array("项目符号列表-F", "List Bullet-F")
    data.Add Array("参考文献列表-F", "List Reference-F")
    data.Add Array("页眉-F", "Header-F")
    data.Add Array("页脚-F", "Footer-F")
    data.Add Array("脚注-F", "Footnote-F")
    data.Add Array("超链接-F", "Hyperlink-F")

    Set GetStyleNameMappingData = data
End Function

Private Sub EnsureMapLoaded()
    Dim data As Collection
    Dim i As Long

    If mMapReady Then Exit Sub

    Set data = GetStyleNameMappingData()
    ReDim keys(1 To data.count) As String
    ReDim vals(1 To data.count) As String

    For i = 1 To data.count
        keys(i) = data(i)(0)
        vals(i) = data(i)(1)
    Next i

    mMapKeys = keys
    mMapValues = vals
    mMapReady = True
End Sub

' 中文 UI 名 -> 英文模板名；无映射时返回空串
Public Function MapStyleNameToEnglish(ByVal uiTagName As String) As String
    Dim i As Long

    EnsureMapLoaded
    For i = LBound(mMapKeys) To UBound(mMapKeys)
        If mMapKeys(i) = uiTagName Then
            MapStyleNameToEnglish = mMapValues(i)
            Exit Function
        End If
    Next i
    MapStyleNameToEnglish = ""
End Function

' 判断样式名是否属于 "-F" 体系（不区分大小写）
Public Function HasStyleSuffix(ByVal styleName As String) As Boolean
    HasStyleSuffix = (UCase$(Right$(styleName, Len(STYLE_SUFFIX))) = UCase$(STYLE_SUFFIX))
End Function

' 常用样式名构造函数（悬浮窗与 Ribbon 共用）
Public Function NumberedHeadingStyle(ByVal level As Long) As String
    NumberedHeadingStyle = "标题" & CStr(level) & STYLE_SUFFIX
End Function

Public Function UnnumberedHeadingStyle(ByVal level As Long) As String
    UnnumberedHeadingStyle = "无编号标题" & CStr(level) & STYLE_SUFFIX
End Function

Public Function BodyTextStyle() As String
    BodyTextStyle = "正文" & STYLE_SUFFIX
End Function

Public Function TableTitleStyle() As String
    TableTitleStyle = "表标题" & STYLE_SUFFIX
End Function

Public Function FigureTitleStyle() As String
    FigureTitleStyle = "图标题" & STYLE_SUFFIX
End Function

' 探测文档中样式是否存在
Public Function StyleExists(ByVal doc As Document, ByVal styleName As String) As Boolean
    Dim s As Style

    On Error Resume Next
    Set s = doc.Styles(styleName)
    StyleExists = (Err.Number = 0)
    On Error GoTo 0
End Function

' 解析 UI 标签对应的实际样式名：
' 1) 文档中直接存在 -> 原名
' 2) 有英文映射且文档中存在 -> 英文名
' 3) 否则返回原名（由调用方处理未找到错误）
Public Function ResolveStyleName(ByVal doc As Document, ByVal uiTagName As String) As String
    Dim mappedName As String

    If StyleExists(doc, uiTagName) Then
        ResolveStyleName = uiTagName
        Exit Function
    End If

    mappedName = MapStyleNameToEnglish(uiTagName)
    If mappedName <> "" Then
        If StyleExists(doc, mappedName) Then
            ResolveStyleName = mappedName
            Exit Function
        End If
    End If

    ResolveStyleName = uiTagName
End Function