Attribute VB_Name = "UI_MacroRegistry"
Option Explicit

' =============================================
' 宏注册表：宏列表窗口数据源与窗口入口
' 从 mRibbon 拆出（v0.8 重构阶段5）。
' GetMyMacroRegistry 中的宏名按名被调用（Application.Run），不可改名。
' =============================================

'=====================  宏 列 表 管 理  =====================

' 1. Ribbon 回调：点击按钮弹出窗体
' 在 Ribbon XML 中，将按钮的 onAction 指向这个 Sub
Public Sub ShowMacroListWindow(control As IRibbonControl)
    frmMacroList.Show vbModeless
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
                    "批量将单个或多个Word转为PDF文档，支持锁定域所见即所得、仅刷新页码或刷新整个目录，并通过Word标题创建PDF书签。")
    
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
                    
    ' 第14个
    items.Add Array("NormalizeScientificTerms", _
                    "标准化科学术语下标", _
                    "将STD10、Cmax、AUC0-t等常见科学术语中的数字或后缀设置为Word下标格式。")
                    
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


Public Sub Wrapper_ShowAbout(control As IRibbonControl)
    frmAbout.Show
End Sub
