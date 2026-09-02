Attribute VB_Name = "Macro_ExtractAbbreviations"
Option Explicit
Sub ExtractAbbreviations()
    Dim docSource As Document
    Dim docTarget As Document
    Dim rng As Range
    Dim dict As Object
    Dim wordText As String
    Dim key As Variant
    Dim sortRange As Range

    ' 检查是否在Windows环境下运行 (Scripting.Dictionary 依赖Windows)
    #If Mac Then
        MsgBox "抱歉，此宏使用了 Scripting.Dictionary，目前仅支持 Windows 版本的 Word。", vbCritical
        Exit Sub
    #End If

    ' 初始化
    Set docSource = ActiveDocument
    ' 创建字典对象用于去重
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbBinaryCompare ' 区分大小写

    ' 设置搜索范围为全文档正文
    Set rng = docSource.Content
    rng.Find.ClearFormatting

    ' 配置通配符查找参数
    With rng.Find
        ' 查找以大写字母开头，且包含2个及以上大写字母的单词
        .Text = "<[A-Z][A-Z0-9]{1,}>"
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = True
        .MatchWildcards = True
    End With

    ' 循环查找并添加到字典
    Do While rng.Find.Execute
        wordText = Trim(rng.Text)
        
        ' 字典去重逻辑：如果字典中没有这个词，则添加进去
        If Not dict.Exists(wordText) Then
            dict.Add wordText, 1
        End If
        rng.Collapse wdCollapseEnd
    Loop

    ' 检查是否找到了缩略语，并输出结果
    If dict.count > 0 Then
        ' 新建一个Word文档
        Set docTarget = Documents.Add
        
        ' 写入标题
        docTarget.Range.Text = "自动提取的缩略语列表（共去重提取 " & dict.count & " 个）：" & vbCrLf & vbCrLf
        
        ' 遍历字典，将去重后的缩略语写入新文档
        For Each key In dict.Keys
            docTarget.Range.InsertAfter key & vbCrLf
        Next key
        
        ' 对提取的缩略语进行字母A-Z排序（跳过前两行的标题）
        Set sortRange = docTarget.Range
        sortRange.Start = docTarget.Paragraphs(3).Range.Start
        sortRange.Sort ExcludeHeader:=False, FieldNumber:=1, _
            SortFieldType:=wdSortFieldAlphanumeric, _
            SortOrder:=wdSortOrderAscending

        MsgBox "提取完成！共找到" & dict.count & "个唯一的纯大写缩略语。" & vbCrLf & "已为您新建文档并按字母A-Z完成排序。", vbInformation, "提取成功"
    Else
        MsgBox "扫描完毕，当前文档中未找到连续的大写英文字母缩略语。", vbExclamation, "无发现"
    End If
    
    ' 释放对象内存
    Set dict = Nothing
    Set docSource = Nothing
    Set docTarget = Nothing
End Sub

